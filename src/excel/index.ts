import type ExcelJS from 'exceljs';
import type { AgendaItem, FlatAgendaWithDepth, MappingJsonV3, ColumnRegion, GridDetection } from '../types';
import { flattenAgendasWithDepth } from '../agenda';

// ── タグ定義 ──
const TAG_TITLE = '{{title}}';
const TAG_SPEAKER = '{{speaker}}';
const TAG_CONTENT = '{{content}}';
const ALL_TAGS = [TAG_TITLE, TAG_SPEAKER, TAG_CONTENT];

const DEFAULT_ROWS_PER_PAGE = 40;

/**
 * column_regions を行スロットにグループ分けする。
 * row_offset があればそれを使い、なければ列範囲の重複を検出して自動グループ化。
 */
const groupRegionsByRow = (regions: ColumnRegion[]): ColumnRegion[][] => {
    // row_offset が設定されているか確認
    const hasRowOffset = regions.some(r => r.row_offset !== undefined && r.row_offset > 0);

    if (hasRowOffset) {
        // row_offset ベースでグループ化
        const maxOffset = Math.max(...regions.map(r => r.row_offset ?? 0));
        const groups: ColumnRegion[][] = Array.from({ length: maxOffset + 1 }, () => []);
        for (const region of regions) {
            groups[region.row_offset ?? 0].push(region);
        }
        return groups.filter(g => g.length > 0);
    }

    // row_offset なし: 列範囲の重複を検出して自動グループ化
    const groups: ColumnRegion[][] = [];
    for (const region of regions) {
        const startCol = colLetterToNumber(region.col_start);
        const endCol = colLetterToNumber(region.col_end);
        let placed = false;
        for (const group of groups) {
            const overlaps = group.some(existing => {
                const eStart = colLetterToNumber(existing.col_start);
                const eEnd = colLetterToNumber(existing.col_end);
                return startCol <= eEnd && endCol >= eStart;
            });
            if (!overlaps) {
                group.push(region);
                placed = true;
                break;
            }
        }
        if (!placed) {
            groups.push([region]);
        }
    }
    return groups;
};

/**
 * セル値からプレーンテキストを抽出する。
 * RichText / Formula / その他の型に対応。
 */
const extractCellText = (value: unknown): string => {
    if (typeof value === 'string') return value;
    if (value == null) return '';
    if (typeof value === 'object') {
        // RichText: { richText: [{ text: '...' }, ...] }
        if ('richText' in value && Array.isArray((value as Record<string, unknown>).richText)) {
            return ((value as Record<string, unknown>).richText as { text: string }[])
                .map(r => r.text)
                .join('');
        }
        // Formula: { formula: '=...', result: '...' }
        if ('result' in value) {
            return String((value as Record<string, unknown>).result || '');
        }
    }
    return String(value);
};

const cellContainsTag = (value: unknown): boolean => {
    const text = extractCellText(value);
    if (!text) return false;
    return ALL_TAGS.some(tag => text.includes(tag));
};

const resolveTagValue = (template: string, item: FlatAgendaWithDepth): string => {
    const indent = '\u3000'.repeat(item.depth);
    let resolved = template;
    resolved = resolved.replace(TAG_TITLE, `${indent}${item.title}`);
    resolved = resolved.replace(TAG_SPEAKER, item.speaker || '');
    resolved = resolved.replace(TAG_CONTENT, item.refinedTranscript || item.rawTranscript || '');
    return resolved;
};

const findTemplateRows = (worksheet: ExcelJS.Worksheet): number[] => {
    const templateRows: number[] = [];
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: false }, (cell) => {
            if (cellContainsTag(cell.value)) {
                if (!templateRows.includes(rowNumber)) {
                    templateRows.push(rowNumber);
                }
            }
        });
    });
    return templateRows.sort((a, b) => a - b);
};

/**
 * テンプレートExcel内の全セルをスキャンし、検出されたタグを返す
 */
export const detectTags = async (
    ExcelJSModule: typeof ExcelJS,
    buffer: Buffer
): Promise<string[]> => {
    const workbook = new ExcelJSModule.Workbook();
    await workbook.xlsx.load(buffer as unknown as ArrayBuffer);

    const tags = new Set<string>();
    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) return [];

    worksheet.eachRow({ includeEmpty: false }, (row) => {
        row.eachCell({ includeEmpty: false }, (cell) => {
            const val = extractCellText(cell.value);
            for (const tag of ALL_TAGS) {
                if (val.includes(tag)) {
                    tags.add(tag);
                }
            }
        });
    });

    return Array.from(tags);
};

/**
 * 列番号を列名（A, B, ..., Z, AA, AB, ...）に変換する
 */
const colNumberToLetter = (col: number): string => {
    let result = '';
    let n = col;
    while (n > 0) {
        n--;
        result = String.fromCharCode(65 + (n % 26)) + result;
        n = Math.floor(n / 26);
    }
    return result;
};

/**
 * 方眼Excel判定: 全列幅の変動係数が閾値以下かつ列数が多い場合に方眼と判定
 */
const detectHouganGrid = (worksheet: ExcelJS.Worksheet): GridDetection => {
    const colCount = worksheet.columnCount || 0;
    if (colCount < 20) return { is_hougan: false, base_cell_size: null };

    const widths: number[] = [];
    for (let c = 1; c <= colCount; c++) {
        const col = worksheet.getColumn(c);
        widths.push(col.width || 8.43);
    }

    const mean = widths.reduce((a, b) => a + b, 0) / widths.length;
    const variance = widths.reduce((sum, w) => sum + (w - mean) ** 2, 0) / widths.length;
    const stddev = Math.sqrt(variance);
    const cv = mean > 0 ? stddev / mean : 1;

    return {
        is_hougan: cv < 0.1,
        base_cell_size: cv < 0.1 ? Math.round(mean * 100) / 100 : null,
    };
};

/**
 * タグを含むセルの結合セル範囲から列範囲を検出する
 */
const detectColumnRegions = (
    worksheet: ExcelJS.Worksheet,
    templateRowNumbers: number[],
    colCount: number,
): ColumnRegion[] => {
    const regions: ColumnRegion[] = [];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const merges: string[] = (worksheet as any).model?.merges || [];
    const firstRow = templateRowNumbers.length > 0 ? templateRowNumbers[0] : 1;

    for (const rowNum of templateRowNumbers) {
        const row = worksheet.getRow(rowNum);
        for (let c = 1; c <= colCount; c++) {
            const cell = row.getCell(c);
            const text = extractCellText(cell.value);
            const matchedTag = ALL_TAGS.find(tag => text.includes(tag));
            if (!matchedTag) continue;

            // 既に同じタグのリージョンがあればスキップ
            if (regions.some(r => r.tag === matchedTag)) continue;

            // 結合セル範囲を探す
            let colStart = c;
            let colEnd = c;
            for (const merge of merges) {
                const match = merge.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
                if (!match) continue;
                const mStartRow = parseInt(match[2], 10);
                const mEndRow = parseInt(match[4], 10);
                if (rowNum >= mStartRow && rowNum <= mEndRow) {
                    // 列文字を列番号に変換してチェック
                    const mStartCol = colLetterToNumber(match[1]);
                    const mEndCol = colLetterToNumber(match[3]);
                    if (c >= mStartCol && c <= mEndCol) {
                        colStart = mStartCol;
                        colEnd = mEndCol;
                        break;
                    }
                }
            }

            // ラベル: テンプレート行の直上行のセル値を取得
            let label = matchedTag;
            if (rowNum > 1) {
                const headerRow = worksheet.getRow(rowNum - 1);
                const headerText = extractCellText(headerRow.getCell(colStart).value);
                if (headerText) label = headerText;
            }

            regions.push({
                tag: matchedTag,
                col_start: colNumberToLetter(colStart),
                col_end: colNumberToLetter(colEnd),
                label,
                row_offset: rowNum - firstRow,
            });
        }
    }

    return regions;
};

/**
 * 列名（A, AA, AC等）を列番号に変換
 */
const colLetterToNumber = (letters: string): number => {
    let num = 0;
    for (let i = 0; i < letters.length; i++) {
        num = num * 26 + (letters.charCodeAt(i) - 64);
    }
    return num;
};

/**
 * テンプレートExcelを解析し、v3 mapping_json を生成する
 */
export const analyzeTemplate = async (
    ExcelJSModule: typeof ExcelJS,
    buffer: Buffer
): Promise<MappingJsonV3> => {
    const workbook = new ExcelJSModule.Workbook();
    await workbook.xlsx.load(buffer as unknown as ArrayBuffer);

    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) {
        return {
            version: 3,
            detected_tags: [],
            print_area: { data_start_row: 1, data_end_row: 40, repeat_header: false, footer_rows: 0 },
            column_regions: [],
            grid_detection: { is_hougan: false, base_cell_size: null },
        };
    }

    // タグ検出
    const tags = new Set<string>();
    worksheet.eachRow({ includeEmpty: false }, (row) => {
        row.eachCell({ includeEmpty: false }, (cell) => {
            const val = extractCellText(cell.value);
            for (const tag of ALL_TAGS) {
                if (val.includes(tag)) tags.add(tag);
            }
        });
    });

    const templateRowNumbers = findTemplateRows(worksheet);
    const colCount = worksheet.columnCount || 20;

    // print_area 推定（タグ位置から初期値を生成。ユーザーがUIで修正する前提）
    const firstTemplateRow = templateRowNumbers.length > 0 ? templateRowNumbers[0] : 1;
    const lastTemplateRow = templateRowNumbers.length > 0 ? templateRowNumbers[templateRowNumbers.length - 1] : 1;
    const totalRows = worksheet.rowCount;

    const print_area = {
        data_start_row: firstTemplateRow,
        data_end_row: firstTemplateRow + 20,
        repeat_header: firstTemplateRow > 1,
        footer_rows: 0,
    };

    // 列範囲推定（初期値。ユーザーがUIで修正する前提）
    const column_regions = detectColumnRegions(worksheet, templateRowNumbers, colCount);

    // 方眼判定
    const grid_detection = detectHouganGrid(worksheet);

    return {
        version: 3,
        detected_tags: Array.from(tags),
        print_area,
        column_regions,
        grid_detection,
    };
};

/**
 * スタイルオブジェクトを深いコピーで複製する
 */
const deepCopyStyle = (style: {
    font?: Partial<ExcelJS.Font>;
    fill?: ExcelJS.Fill;
    border?: Partial<ExcelJS.Borders>;
    alignment?: Partial<ExcelJS.Alignment>;
    numFmt?: string;
}) => ({
    font: style.font ? JSON.parse(JSON.stringify(style.font)) : undefined,
    fill: style.fill ? JSON.parse(JSON.stringify(style.fill)) : undefined,
    border: style.border ? JSON.parse(JSON.stringify(style.border)) : undefined,
    alignment: style.alignment ? JSON.parse(JSON.stringify(style.alignment)) : undefined,
    numFmt: style.numFmt,
});

interface CellStyle {
    col: number;
    font?: Partial<ExcelJS.Font>;
    fill?: ExcelJS.Fill;
    border?: Partial<ExcelJS.Borders>;
    alignment?: Partial<ExcelJS.Alignment>;
    numFmt?: string;
}

const captureCellStyle = (cell: ExcelJS.Cell, col: number): CellStyle => ({
    col,
    font: cell.font ? JSON.parse(JSON.stringify(cell.font)) : undefined,
    fill: cell.fill ? JSON.parse(JSON.stringify(cell.fill)) as ExcelJS.Fill : undefined,
    border: cell.border ? JSON.parse(JSON.stringify(cell.border)) : undefined,
    alignment: cell.alignment ? JSON.parse(JSON.stringify(cell.alignment)) : undefined,
    numFmt: cell.numFmt || undefined,
});

interface TemplateRowData {
    cells: { col: number; template: string; isTag: boolean }[];
    height: number | undefined;
    styles: CellStyle[];
}

/**
 * セルタグ方式でExcelテンプレートに議題データを流し込む。
 */
export const generateExcelFromTagTemplate = async (
    ExcelJSModule: typeof ExcelJS,
    agendas: AgendaItem[],
    templateBuffer: Buffer,
    options?: { rowsPerPage?: number }
): Promise<Buffer> => {
    const rowsPerPage = options?.rowsPerPage || DEFAULT_ROWS_PER_PAGE;
    const flatAgendas = flattenAgendasWithDepth(agendas);

    const workbook = new ExcelJSModule.Workbook();
    await workbook.xlsx.load(templateBuffer as unknown as ArrayBuffer);

    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) {
        throw new Error('テンプレートにシートが見つかりません');
    }

    const templateRowNumbers = findTemplateRows(worksheet);

    if (templateRowNumbers.length === 0) {
        const buffer = await workbook.xlsx.writeBuffer();
        return Buffer.from(buffer as ArrayBuffer);
    }

    const firstTemplateRow = templateRowNumbers[0];
    const lastTemplateRow = templateRowNumbers[templateRowNumbers.length - 1];
    const colCount = worksheet.columnCount || 20;

    // Step 1: 列幅を保存
    const columnWidths: number[] = [];
    for (let c = 1; c <= colCount; c++) {
        const col = worksheet.getColumn(c);
        columnWidths.push(col.width || 8.43);
    }

    // Step 2: 結合セル情報を保存
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const merges: string[] = (worksheet as any).model?.merges
        ? [...(worksheet as any).model.merges]
        : [];

    // Step 3: テンプレート行のデータとスタイルを記憶（タグを含む行のみ）
    const templateRowsData: TemplateRowData[] = [];
    for (const r of templateRowNumbers) {
        const row = worksheet.getRow(r);
        const cells: TemplateRowData['cells'] = [];
        const styles: CellStyle[] = [];

        for (let c = 1; c <= colCount; c++) {
            const cell = row.getCell(c);
            const val = extractCellText(cell.value);
            cells.push({ col: c, template: val, isTag: cellContainsTag(cell.value) });
            styles.push(captureCellStyle(cell, c));
        }

        templateRowsData.push({ cells, height: row.height || undefined, styles });
    }

    // Step 4: ヘッダー行のデータを記憶
    interface HeaderRowData {
        cells: { col: number; value: unknown }[];
        height: number | undefined;
        styles: CellStyle[];
    }

    const headerRowsData: HeaderRowData[] = [];
    for (let r = 1; r < firstTemplateRow; r++) {
        const row = worksheet.getRow(r);
        const cells: HeaderRowData['cells'] = [];
        const styles: CellStyle[] = [];

        for (let c = 1; c <= colCount; c++) {
            const cell = row.getCell(c);
            cells.push({ col: c, value: cell.value });
            styles.push(captureCellStyle(cell, c));
        }
        headerRowsData.push({ cells, height: row.height || undefined, styles });
    }

    // Step 5: フッター行のデータを記憶
    const totalOriginalRows = worksheet.rowCount;
    const footerRowsData: HeaderRowData[] = [];
    for (let r = lastTemplateRow + 1; r <= totalOriginalRows; r++) {
        const row = worksheet.getRow(r);
        const cells: HeaderRowData['cells'] = [];
        const styles: CellStyle[] = [];

        for (let c = 1; c <= colCount; c++) {
            const cell = row.getCell(c);
            cells.push({ col: c, value: cell.value });
            styles.push(captureCellStyle(cell, c));
        }
        footerRowsData.push({ cells, height: row.height || undefined, styles });
    }

    // Step 6: クリア（結合セルも解除）
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wsModel = worksheet as any;
    if (wsModel.model?.merges) {
        wsModel.model.merges = [];
    }
    for (let r = firstTemplateRow; r <= totalOriginalRows; r++) {
        const row = worksheet.getRow(r);
        for (let c = 1; c <= colCount; c++) {
            row.getCell(c).value = null;
        }
    }

    // Step 7: ヘッダー部分の結合セルを再適用
    for (const merge of merges) {
        const match = merge.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
        if (!match) continue;
        const startRow = parseInt(match[2], 10);
        if (startRow < firstTemplateRow) {
            worksheet.mergeCells(merge);
        }
    }

    // テンプレート行の結合セル情報を抽出（行オフセットで保存）
    const templateMerges: { merge: string; rowOffset: number; startRow: number; endRow: number }[] = [];
    for (const merge of merges) {
        const match = merge.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
        if (!match) continue;
        const startRow = parseInt(match[2], 10);
        const endRow = parseInt(match[4], 10);
        if (startRow >= firstTemplateRow && endRow <= lastTemplateRow) {
            templateMerges.push({
                merge,
                rowOffset: startRow - firstTemplateRow,
                startRow,
                endRow,
            });
        }
    }

    // Step 8: 議題データを展開
    const headerRowCount = headerRowsData.length;
    const footerRowCount = footerRowsData.length;
    const dataRowsPerPage = Math.max(rowsPerPage - headerRowCount - footerRowCount, 1);

    let currentRow = firstTemplateRow;
    let rowsOnCurrentPage = 0;

    const applyStyles = (row: ExcelJS.Row, styles: CellStyle[]) => {
        for (const s of styles) {
            const cell = row.getCell(s.col);
            const copied = deepCopyStyle(s);
            if (copied.font) cell.font = copied.font;
            if (copied.fill) cell.fill = copied.fill;
            if (copied.border) cell.border = copied.border;
            if (copied.alignment) cell.alignment = copied.alignment;
            if (copied.numFmt) cell.numFmt = copied.numFmt;
        }
    };

    const writeHeaderRows = (startRow: number): number => {
        for (let i = 0; i < headerRowsData.length; i++) {
            const hd = headerRowsData[i];
            const row = worksheet.getRow(startRow + i);
            if (hd.height !== undefined) row.height = hd.height;
            for (const c of hd.cells) {
                row.getCell(c.col).value = c.value as ExcelJS.CellValue;
            }
            applyStyles(row, hd.styles);
            row.commit();
        }

        // ヘッダー部分の結合セルを再適用（行オフセット調整）
        for (const merge of merges) {
            const match = merge.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
            if (!match) continue;
            const origStartRow = parseInt(match[2], 10);
            const origEndRow = parseInt(match[4], 10);
            if (origStartRow < firstTemplateRow) {
                const offset = startRow - 1;
                const newMerge = `${match[1]}${origStartRow + offset}:${match[3]}${origEndRow + offset}`;
                try { worksheet.mergeCells(newMerge); } catch { /* already merged */ }
            }
        }

        return startRow + headerRowsData.length;
    };

    const writeFooterRows = (startRow: number): number => {
        for (let i = 0; i < footerRowsData.length; i++) {
            const fd = footerRowsData[i];
            const row = worksheet.getRow(startRow + i);
            if (fd.height !== undefined) row.height = fd.height;
            for (const c of fd.cells) {
                row.getCell(c.col).value = c.value as ExcelJS.CellValue;
            }
            applyStyles(row, fd.styles);
            row.commit();
        }
        return startRow + footerRowsData.length;
    };

    const templateRowCount = templateRowsData.length;

    for (let agendaIdx = 0; agendaIdx < flatAgendas.length; agendaIdx++) {
        const item = flatAgendas[agendaIdx];

        if (dataRowsPerPage > 0 && rowsOnCurrentPage >= dataRowsPerPage) {
            currentRow = writeFooterRows(currentRow);

            // ページブレークを挿入（フッター直後の行に設定）
            worksheet.getRow(currentRow).addPageBreak();

            currentRow = writeHeaderRows(currentRow);
            rowsOnCurrentPage = 0;
        }

        const dataStartRow = currentRow;

        for (let trdIdx = 0; trdIdx < templateRowsData.length; trdIdx++) {
            const trd = templateRowsData[trdIdx];
            const row = worksheet.getRow(currentRow);
            if (trd.height !== undefined) row.height = trd.height;

            for (const cellData of trd.cells) {
                const cell = row.getCell(cellData.col);
                cell.value = cellData.isTag
                    ? resolveTagValue(cellData.template, item)
                    : (cellData.template || null);
            }

            applyStyles(row, trd.styles);
            for (const cellData of trd.cells) {
                if (cellData.isTag) {
                    const cell = row.getCell(cellData.col);
                    cell.alignment = { ...cell.alignment, wrapText: true, vertical: 'top' };
                }
            }
            row.commit();
            currentRow++;
            rowsOnCurrentPage++;
        }

        // テンプレート行の結合セルをデータ行に再適用
        for (const tm of templateMerges) {
            const match = tm.merge.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
            if (!match) continue;
            const rowShift = dataStartRow - firstTemplateRow;
            const newMerge = `${match[1]}${parseInt(match[2], 10) + rowShift}:${match[3]}${parseInt(match[4], 10) + rowShift}`;
            try { worksheet.mergeCells(newMerge); } catch { /* skip if overlap */ }
        }
    }

    writeFooterRows(currentRow);

    // 列幅を復元
    for (let c = 1; c <= columnWidths.length; c++) {
        worksheet.getColumn(c).width = columnWidths[c - 1];
    }

    const buffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(buffer as ArrayBuffer);
};

// ── テキスト幅計算ユーティリティ ──

/**
 * 文字の表示幅を返す（全角=2, 半角=1）
 */
const charWidth = (code: number): number => {
    // CJK統合漢字、全角カナ、全角記号、ハングル等
    if (
        (code >= 0x3000 && code <= 0x9FFF) ||
        (code >= 0xF900 && code <= 0xFAFF) ||
        (code >= 0xFE30 && code <= 0xFE4F) ||
        (code >= 0xFF01 && code <= 0xFF60) ||
        (code >= 0xFFE0 && code <= 0xFFE6) ||
        (code >= 0xAC00 && code <= 0xD7AF) ||
        (code >= 0x20000 && code <= 0x2FFFF)
    ) {
        return 2;
    }
    return 1;
};

/**
 * テキストの表示幅（半角単位）を計算する
 */
const measureTextWidth = (text: string): number => {
    let width = 0;
    for (const ch of text) {
        width += charWidth(ch.codePointAt(0) || 0);
    }
    return width;
};

/**
 * セル幅（ExcelJS文字数単位）に対する折り返し行数を計算する
 */
const calcWrapLineCount = (text: string, cellWidthChars: number): number => {
    if (!text || cellWidthChars <= 0) return 1;

    // 改行で分割し、各行の折り返しを計算
    const lines = text.split('\n');
    let totalLines = 0;
    for (const line of lines) {
        const width = measureTextWidth(line);
        // ExcelJS幅は半角文字数基準
        totalLines += Math.max(Math.ceil(width / cellWidthChars), 1);
    }
    return totalLines;
};

/**
 * 折り返し行数から行高さを計算する（デフォルト行高さ15pt, フォントサイズ11pt基準）
 */
const calcRowHeight = (wrapLines: number, defaultRowHeight = 15): number => {
    return Math.max(wrapLines * defaultRowHeight, defaultRowHeight);
};

/**
 * v3 mapping_json を使用してExcelテンプレートに議題データを流し込む。
 * 列範囲ベース・高さベース改ページ対応。
 */
export const generateExcelFromV3Template = async (
    ExcelJSModule: typeof ExcelJS,
    agendas: AgendaItem[],
    templateBuffer: Buffer,
    mappingJson: MappingJsonV3
): Promise<Buffer> => {
    const flatAgendas = flattenAgendasWithDepth(agendas);

    const workbook = new ExcelJSModule.Workbook();
    await workbook.xlsx.load(templateBuffer as unknown as ArrayBuffer);

    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) {
        throw new Error('テンプレートにシートが見つかりません');
    }

    const { print_area, column_regions } = mappingJson;
    const { data_start_row, data_end_row, repeat_header, footer_rows } = print_area;

    const colCount = worksheet.columnCount || 20;

    // 列幅を保存
    const columnWidths: number[] = [];
    for (let c = 1; c <= colCount; c++) {
        columnWidths.push(worksheet.getColumn(c).width || 8.43);
    }

    // 結合セル情報を保存
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wsModel = worksheet as any;
    const merges: string[] = wsModel.model?.merges ? [...wsModel.model.merges] : [];

    // ヘッダー行を記憶（data_start_row より前）
    const headerRowCount = data_start_row - 1;
    const headerRowsData: { cells: { col: number; value: unknown }[]; height: number | undefined; styles: CellStyle[] }[] = [];
    for (let r = 1; r < data_start_row; r++) {
        const row = worksheet.getRow(r);
        const cells: { col: number; value: unknown }[] = [];
        const styles: CellStyle[] = [];
        for (let c = 1; c <= colCount; c++) {
            const cell = row.getCell(c);
            cells.push({ col: c, value: cell.value });
            styles.push(captureCellStyle(cell, c));
        }
        headerRowsData.push({ cells, height: row.height || undefined, styles });
    }

    // column_regions を行グループに分割
    const rowGroups = groupRegionsByRow(column_regions);
    const rowsPerItem = rowGroups.length;

    // 各行グループのテンプレート行スタイルを記憶
    const templateStylesPerRow: CellStyle[][] = [];
    const templateRowHeights: number[] = [];
    for (let r = 0; r < rowsPerItem; r++) {
        const row = worksheet.getRow(data_start_row + r);
        const styles: CellStyle[] = [];
        for (let c = 1; c <= colCount; c++) {
            styles.push(captureCellStyle(row.getCell(c), c));
        }
        templateStylesPerRow.push(styles);
        templateRowHeights.push(row.height || 15);
    }
    const templateRowHeight = templateRowHeights[0];

    // column_region がカバーする列範囲を構築（テンプレートマージの競合排除用）
    const regionCoveredRanges = column_regions.map(r => ({
        start: colLetterToNumber(r.col_start),
        end: colLetterToNumber(r.col_end),
        rowOffset: r.row_offset ?? 0,
    }));

    // テンプレート行の結合セル情報を抽出（パターン行のみ）
    const templateMerges: string[] = [];
    for (const merge of merges) {
        const match = merge.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
        if (!match) continue;
        const startRow = parseInt(match[2], 10);
        const endRow = parseInt(match[4], 10);
        // パターン行の範囲内かつ単一行の結合のみ
        if (startRow >= data_start_row && endRow < data_start_row + rowsPerItem && startRow === endRow) {
            // column_region とカバー範囲が重複する結合はスキップ
            const mStartCol = colLetterToNumber(match[1]);
            const mEndCol = colLetterToNumber(match[3]);
            const rowOffset = startRow - data_start_row;
            const overlapsWithRegion = regionCoveredRanges.some(r =>
                r.rowOffset === rowOffset && mStartCol <= r.end && mEndCol >= r.start
            );
            if (!overlapsWithRegion) {
                templateMerges.push(merge);
            }
        }
    }

    // フッター行を記憶
    const totalOriginalRows = worksheet.rowCount;
    const footerStartRow = data_end_row + 1;
    const footerRowsData: { cells: { col: number; value: unknown }[]; height: number | undefined; styles: CellStyle[] }[] = [];
    for (let r = footerStartRow; r <= Math.min(footerStartRow + footer_rows - 1, totalOriginalRows); r++) {
        const row = worksheet.getRow(r);
        const cells: { col: number; value: unknown }[] = [];
        const styles: CellStyle[] = [];
        for (let c = 1; c <= colCount; c++) {
            const cell = row.getCell(c);
            cells.push({ col: c, value: cell.value });
            styles.push(captureCellStyle(cell, c));
        }
        footerRowsData.push({ cells, height: row.height || undefined, styles });
    }

    // データ領域をクリア
    if (wsModel.model?.merges) {
        // テンプレート行以降の結合セルを解除
        wsModel.model.merges = merges.filter((merge: string) => {
            const match = merge.match(/^[A-Z]+(\d+):/);
            return match ? parseInt(match[1], 10) < data_start_row : true;
        });
    }
    for (let r = data_start_row; r <= totalOriginalRows; r++) {
        const row = worksheet.getRow(r);
        for (let c = 1; c <= colCount; c++) {
            row.getCell(c).value = null;
        }
    }

    // 列範囲ごとの実効セル幅（半角文字数）を算出
    const regionWidths = new Map<string, number>();
    for (const region of column_regions) {
        const startCol = colLetterToNumber(region.col_start);
        const endCol = colLetterToNumber(region.col_end);
        let totalWidth = 0;
        for (let c = startCol; c <= endCol; c++) {
            totalWidth += columnWidths[c - 1] || 8.43;
        }
        regionWidths.set(region.tag, totalWidth);
    }

    // 1ページのデータ領域の物理的な高さ
    const pageDataHeight = (data_end_row - data_start_row + 1) * templateRowHeight;

    let currentRow = data_start_row;
    let currentPageHeight = 0;

    const applyStylesLocal = (row: ExcelJS.Row, styles: CellStyle[]) => {
        for (const s of styles) {
            const cell = row.getCell(s.col);
            const copied = deepCopyStyle(s);
            if (copied.font) cell.font = copied.font;
            if (copied.fill) cell.fill = copied.fill;
            if (copied.border) cell.border = copied.border;
            if (copied.alignment) cell.alignment = copied.alignment;
            if (copied.numFmt) cell.numFmt = copied.numFmt;
        }
    };

    const writeHeaders = (startRow: number): number => {
        for (let i = 0; i < headerRowsData.length; i++) {
            const hd = headerRowsData[i];
            const row = worksheet.getRow(startRow + i);
            if (hd.height !== undefined) row.height = hd.height;
            for (const c of hd.cells) {
                row.getCell(c.col).value = c.value as ExcelJS.CellValue;
            }
            applyStylesLocal(row, hd.styles);
            row.commit();
        }
        // ヘッダーの結合セルを再適用
        for (const merge of merges) {
            const match = merge.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
            if (!match) continue;
            const origStartRow = parseInt(match[2], 10);
            const origEndRow = parseInt(match[4], 10);
            if (origStartRow < data_start_row) {
                const offset = startRow - 1;
                const newMerge = `${match[1]}${origStartRow + offset}:${match[3]}${origEndRow + offset}`;
                try { worksheet.mergeCells(newMerge); } catch { /* already merged */ }
            }
        }
        return startRow + headerRowsData.length;
    };

    const writeFooters = (startRow: number): number => {
        for (let i = 0; i < footerRowsData.length; i++) {
            const fd = footerRowsData[i];
            const row = worksheet.getRow(startRow + i);
            if (fd.height !== undefined) row.height = fd.height;
            for (const c of fd.cells) {
                row.getCell(c.col).value = c.value as ExcelJS.CellValue;
            }
            applyStylesLocal(row, fd.styles);
            row.commit();
        }
        // フッターの結合セルを再適用
        for (const merge of merges) {
            const match = merge.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
            if (!match) continue;
            const origStartRow = parseInt(match[2], 10);
            const origEndRow = parseInt(match[4], 10);
            if (origStartRow >= footerStartRow && origEndRow <= footerStartRow + footer_rows - 1) {
                const offset = startRow - footerStartRow;
                const newMerge = `${match[1]}${origStartRow + offset}:${match[3]}${origEndRow + offset}`;
                try { worksheet.mergeCells(newMerge); } catch { /* already merged */ }
            }
        }
        return startRow + footerRowsData.length;
    };

    // 議題データを展開
    for (const item of flatAgendas) {
        const indent = '\u3000'.repeat(item.depth);

        // 各リージョンのテキストを解決
        const resolvedValues = new Map<string, string>();
        for (const region of column_regions) {
            let value = '';
            if (region.tag === TAG_TITLE) value = `${indent}${item.title}`;
            else if (region.tag === TAG_SPEAKER) value = item.speaker || '';
            else if (region.tag === TAG_CONTENT) value = item.refinedTranscript || item.rawTranscript || '';
            resolvedValues.set(region.tag, value);
        }

        // 行グループごとの行高さを計算
        const rowHeights: number[] = [];
        let totalItemHeight = 0;
        for (let g = 0; g < rowGroups.length; g++) {
            let maxWrapLines = 1;
            for (const region of rowGroups[g]) {
                const text = resolvedValues.get(region.tag) || '';
                const cellWidth = regionWidths.get(region.tag) || 8.43;
                const wrapLines = calcWrapLineCount(text, cellWidth);
                maxWrapLines = Math.max(maxWrapLines, wrapLines);
            }
            const h = calcRowHeight(maxWrapLines, templateRowHeights[g] || templateRowHeight);
            rowHeights.push(h);
            totalItemHeight += h;
        }

        // 改ページ判定
        if (currentPageHeight + totalItemHeight > pageDataHeight && currentPageHeight > 0) {
            currentRow = writeFooters(currentRow);

            // ページブレークを挿入（フッター直後の行に設定）
            worksheet.getRow(currentRow).addPageBreak();

            if (repeat_header) {
                currentRow = writeHeaders(currentRow);
            }
            currentPageHeight = 0;
        }

        // 各行グループのデータを書き込み
        for (let g = 0; g < rowGroups.length; g++) {
            const row = worksheet.getRow(currentRow + g);
            row.height = rowHeights[g];

            // テンプレート行のスタイルを適用
            applyStylesLocal(row, templateStylesPerRow[g] || templateStylesPerRow[0]);

            // 列範囲ベースでデータを書き込む（セル結合 + 書き込み）
            for (const region of rowGroups[g]) {
                const startCol = colLetterToNumber(region.col_start);
                const endCol = colLetterToNumber(region.col_end);
                const value = resolvedValues.get(region.tag) || '';

                if (startCol !== endCol) {
                    try {
                        worksheet.mergeCells(currentRow + g, startCol, currentRow + g, endCol);
                    } catch { /* already merged */ }
                }

                const cell = row.getCell(startCol);
                cell.value = value;
                cell.alignment = { ...cell.alignment, wrapText: true, vertical: 'top' };
            }

            // テンプレート行の結合セルパターンを再適用（column_region 外の装飾的結合のみ）
            for (const merge of templateMerges) {
                const match = merge.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
                if (!match) continue;
                const origRow = parseInt(match[2], 10);
                const mergeRowOffset = origRow - data_start_row;
                // この行グループに属する結合のみ適用
                if (mergeRowOffset !== g) continue;
                const rowShift = currentRow - data_start_row;
                const newMerge = `${match[1]}${origRow + rowShift}:${match[3]}${parseInt(match[4], 10) + rowShift}`;
                try { worksheet.mergeCells(newMerge); } catch { /* skip */ }
            }

            row.commit();
        }

        currentRow += rowsPerItem;
        currentPageHeight += totalItemHeight;
    }

    // 最後のフッターを書き込み
    writeFooters(currentRow);

    // 列幅を復元
    for (let c = 1; c <= columnWidths.length; c++) {
        worksheet.getColumn(c).width = columnWidths[c - 1];
    }

    const buffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(buffer as ArrayBuffer);
};

/**
 * テンプレートなしで既定のフォーマットでExcelを生成する
 */
export const generateExcelBuffer = async (
    ExcelJSModule: typeof ExcelJS,
    agendas: AgendaItem[]
): Promise<Buffer> => {
    const flatAgendas = flattenAgendasWithDepth(agendas);
    const workbook = new ExcelJSModule.Workbook();

    // --- SHEET 1: 議事録(AI清書) ---
    const sheet1 = workbook.addWorksheet('議事録(AI清書)');

    sheet1.columns = [
        { header: '議題', key: 'title', width: 30 },
        { header: '内容', key: 'content', width: 60 },
        { header: '発言者', key: 'speaker', width: 15 },
    ];

    flatAgendas.forEach(item => {
        const indent = '\u3000'.repeat(item.depth);
        sheet1.addRow({
            title: `${indent}${item.title}`,
            content: item.refinedTranscript || item.rawTranscript || '',
            speaker: item.speaker || '',
        });
    });

    const headerRow1 = sheet1.getRow(1);
    headerRow1.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
    headerRow1.alignment = { vertical: 'middle', horizontal: 'center' };
    headerRow1.commit();

    sheet1.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        row.alignment = { vertical: 'top', wrapText: true };
        row.commit();
    });

    // --- SHEET 2: 原文ログ(検索用) ---
    const sheet2 = workbook.addWorksheet('原文ログ(検索用)');

    sheet2.columns = [
        { header: '対象議題', key: 'title', width: 25 },
        { header: '発言者', key: 'speaker', width: 15 },
        { header: '原文テキスト', key: 'content', width: 80 },
    ];

    flatAgendas.forEach(item => {
        sheet2.addRow({
            title: item.title,
            speaker: item.speaker || '',
            content: item.rawTranscript || ''
        });
    });

    const headerRow2 = sheet2.getRow(1);
    headerRow2.font = { bold: true, color: { argb: 'FF000000' } };
    headerRow2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
    headerRow2.alignment = { vertical: 'middle', horizontal: 'center' };
    headerRow2.commit();

    sheet2.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        row.alignment = { vertical: 'top', wrapText: true };
        row.commit();
    });

    const buffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(buffer);
};

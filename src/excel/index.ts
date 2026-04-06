import type ExcelJS from 'exceljs';
import type { AgendaItem, FlatAgendaWithDepth } from '../types';
import { flattenAgendasWithDepth } from '../agenda';

// ── タグ定義 ──
const TAG_TITLE = '{{title}}';
const TAG_SPEAKER = '{{speaker}}';
const TAG_CONTENT = '{{content}}';
const ALL_TAGS = [TAG_TITLE, TAG_SPEAKER, TAG_CONTENT];

const DEFAULT_ROWS_PER_PAGE = 40;

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

            if (!wsModel.pageSetup) wsModel.pageSetup = {};
            if (!wsModel.rowBreaks) wsModel.rowBreaks = [];
            wsModel.rowBreaks.push(currentRow - 1);

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

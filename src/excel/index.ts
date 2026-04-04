import type ExcelJS from 'exceljs';
import type { AgendaItem, FlatAgendaWithDepth } from '../types';
import { flattenAgendasWithDepth } from '../agenda';
import { fetchImageAsBuffer } from '../image';

// ── タグ定義 ──
const TAG_TITLE = '{{title}}';
const TAG_SPEAKER = '{{speaker}}';
const TAG_CONTENT = '{{content}}';
const ALL_TAGS = [TAG_TITLE, TAG_SPEAKER, TAG_CONTENT];

const DEFAULT_ROWS_PER_PAGE = 40;

const cellContainsTag = (value: unknown): boolean => {
    if (typeof value !== 'string') return false;
    return ALL_TAGS.some(tag => value.includes(tag));
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
            const val = String(cell.value || '');
            for (const tag of ALL_TAGS) {
                if (val.includes(tag)) {
                    tags.add(tag);
                }
            }
        });
    });

    return Array.from(tags);
};

interface TemplateRowData {
    cells: { col: number; template: string; isTag: boolean }[];
    height: number | undefined;
    styles: {
        col: number;
        font?: Partial<ExcelJS.Font>;
        fill?: ExcelJS.Fill;
        border?: Partial<ExcelJS.Borders>;
        alignment?: Partial<ExcelJS.Alignment>;
        numFmt?: string;
    }[];
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

    // Step 2: テンプレート行のデータとスタイルを記憶
    const templateRowsData: TemplateRowData[] = [];
    for (let r = firstTemplateRow; r <= lastTemplateRow; r++) {
        const row = worksheet.getRow(r);
        const cells: TemplateRowData['cells'] = [];
        const styles: TemplateRowData['styles'] = [];

        for (let c = 1; c <= colCount; c++) {
            const cell = row.getCell(c);
            const val = String(cell.value || '');
            cells.push({ col: c, template: val, isTag: cellContainsTag(val) });
            styles.push({
                col: c,
                font: cell.font ? { ...cell.font } : undefined,
                fill: cell.fill ? { ...cell.fill } as ExcelJS.Fill : undefined,
                border: cell.border ? { ...cell.border } : undefined,
                alignment: cell.alignment ? { ...cell.alignment } : undefined,
                numFmt: cell.numFmt || undefined,
            });
        }

        templateRowsData.push({ cells, height: row.height || undefined, styles });
    }

    // Step 3: ヘッダー行のデータを記憶
    interface HeaderRowData {
        cells: { col: number; value: unknown }[];
        height: number | undefined;
        styles: TemplateRowData['styles'];
    }

    const headerRowsData: HeaderRowData[] = [];
    for (let r = 1; r < firstTemplateRow; r++) {
        const row = worksheet.getRow(r);
        const cells: HeaderRowData['cells'] = [];
        const styles: TemplateRowData['styles'] = [];

        for (let c = 1; c <= colCount; c++) {
            const cell = row.getCell(c);
            cells.push({ col: c, value: cell.value });
            styles.push({
                col: c,
                font: cell.font ? { ...cell.font } : undefined,
                fill: cell.fill ? { ...cell.fill } as ExcelJS.Fill : undefined,
                border: cell.border ? { ...cell.border } : undefined,
                alignment: cell.alignment ? { ...cell.alignment } : undefined,
                numFmt: cell.numFmt || undefined,
            });
        }
        headerRowsData.push({ cells, height: row.height || undefined, styles });
    }

    // Step 4: フッター行のデータを記憶
    const totalOriginalRows = worksheet.rowCount;
    const footerRowsData: HeaderRowData[] = [];
    for (let r = lastTemplateRow + 1; r <= totalOriginalRows; r++) {
        const row = worksheet.getRow(r);
        const cells: HeaderRowData['cells'] = [];
        const styles: TemplateRowData['styles'] = [];

        for (let c = 1; c <= colCount; c++) {
            const cell = row.getCell(c);
            cells.push({ col: c, value: cell.value });
            styles.push({
                col: c,
                font: cell.font ? { ...cell.font } : undefined,
                fill: cell.fill ? { ...cell.fill } as ExcelJS.Fill : undefined,
                border: cell.border ? { ...cell.border } : undefined,
                alignment: cell.alignment ? { ...cell.alignment } : undefined,
                numFmt: cell.numFmt || undefined,
            });
        }
        footerRowsData.push({ cells, height: row.height || undefined, styles });
    }

    // Step 5: クリア
    for (let r = firstTemplateRow; r <= totalOriginalRows; r++) {
        const row = worksheet.getRow(r);
        for (let c = 1; c <= colCount; c++) {
            row.getCell(c).value = null;
        }
    }

    // Step 6: 議題データを展開
    const headerRowCount = headerRowsData.length;
    const footerRowCount = footerRowsData.length;
    const dataRowsPerPage = rowsPerPage - headerRowCount - footerRowCount;

    let currentRow = firstTemplateRow;
    let rowsOnCurrentPage = 0;

    const applyStyles = (row: ExcelJS.Row, styles: TemplateRowData['styles']) => {
        for (const s of styles) {
            const cell = row.getCell(s.col);
            if (s.font) cell.font = s.font;
            if (s.fill) cell.fill = s.fill;
            if (s.border) cell.border = s.border;
            if (s.alignment) cell.alignment = s.alignment;
            if (s.numFmt) cell.numFmt = s.numFmt;
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

    for (let agendaIdx = 0; agendaIdx < flatAgendas.length; agendaIdx++) {
        const item = flatAgendas[agendaIdx];

        if (dataRowsPerPage > 0 && rowsOnCurrentPage >= dataRowsPerPage) {
            currentRow = writeFooterRows(currentRow);

            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const wsModel = worksheet as any;
            if (!wsModel.pageSetup) wsModel.pageSetup = {};
            if (!wsModel.rowBreaks) wsModel.rowBreaks = [];
            wsModel.rowBreaks.push(currentRow - 1);

            currentRow = writeHeaderRows(currentRow);
            rowsOnCurrentPage = 0;
        }

        for (const trd of templateRowsData) {
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
    }

    writeFooterRows(currentRow);

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

    // --- SHEET 3: 図表（enrichedContent がある場合のみ） ---
    const hasRichContent = flatAgendas.some(a => a.enrichedContent);
    if (hasRichContent) {
        const sheet3 = workbook.addWorksheet('図表');
        let sheetRow = 1;

        for (const item of flatAgendas) {
            const ec = item.enrichedContent;
            if (!ec) continue;

            const titleRow = sheet3.getRow(sheetRow);
            titleRow.getCell(1).value = item.title;
            titleRow.getCell(1).font = { bold: true, size: 14 };
            titleRow.commit();
            sheetRow++;

            for (const img of ec.images) {
                try {
                    const imageId = workbook.addImage({ base64: img.base64, extension: 'png' });
                    const maxW = 500;
                    let w = img.width;
                    let h = img.height;
                    if (w > maxW) { h = h * (maxW / w); w = maxW; }
                    sheet3.addImage(imageId, {
                        tl: { col: 0, row: sheetRow - 1 } as ExcelJS.Anchor,
                        ext: { width: w, height: h },
                    });
                    sheetRow += Math.max(Math.ceil(h / 20), 1);
                } catch (e) {
                    console.warn(`[Excel] Failed to embed image for "${item.title}":`, e);
                }
            }

            for (const url of ec.imageUrls) {
                const fetched = await fetchImageAsBuffer(url);
                if (!fetched) continue;
                try {
                    const imageId = workbook.addImage({
                        buffer: fetched.buffer as unknown as ExcelJS.Buffer,
                        extension: fetched.ext,
                    });
                    sheet3.addImage(imageId, {
                        tl: { col: 0, row: sheetRow - 1 } as ExcelJS.Anchor,
                        ext: { width: 400, height: 300 },
                    });
                    sheetRow += 16;
                } catch (e) {
                    console.warn(`[Excel] Failed to embed URL image:`, e);
                }
            }

            for (const tbl of ec.tables) {
                const hRow = sheet3.getRow(sheetRow);
                tbl.headers.forEach((h, i) => {
                    const cell = hRow.getCell(i + 1);
                    cell.value = h;
                    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
                    cell.border = {
                        top: { style: 'thin' }, bottom: { style: 'thin' },
                        left: { style: 'thin' }, right: { style: 'thin' },
                    };
                });
                hRow.commit();
                sheetRow++;

                for (const row of tbl.rows) {
                    const dataRow = sheet3.getRow(sheetRow);
                    row.forEach((val, i) => {
                        const cell = dataRow.getCell(i + 1);
                        cell.value = val;
                        cell.border = {
                            top: { style: 'thin' }, bottom: { style: 'thin' },
                            left: { style: 'thin' }, right: { style: 'thin' },
                        };
                    });
                    dataRow.commit();
                    sheetRow++;
                }
                sheetRow++;
            }
            sheetRow++;
        }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(buffer);
};

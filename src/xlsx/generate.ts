/**
 * XLSX 生成
 *
 * 输入即 jdoc 的 XlsxWorkbook 结构；行支持数组式与对象式两种简写，
 * 单元格支持「直接写值」与「带样式对象」两种形式。
 *
 * @module xlsx/generate
 */

import ExcelJS from 'exceljs';
import { wrapError } from '../core/errors';
import type {
    XlsxCell,
    XlsxCellObject,
    XlsxCellStyle,
    XlsxGenerateOptions,
    XlsxRow,
    XlsxRowInput,
    XlsxWorkbook
} from './types';

/** A1 记号 -> 列号（'A1' -> 1；只取字母部分） */
export function columnFromRef(ref: string): number {
    const match = /^([A-Za-z]+)/.exec(String(ref).trim());
    if (!match) return NaN;
    let index = 0;
    for (const ch of match[1].toUpperCase()) {
        index = index * 26 + (ch.charCodeAt(0) - 64);
    }
    return index;
}

/** 判断是否为 XlsxCellObject（对象形式而非 Date） */
function isCellObject(cell: XlsxCell): cell is XlsxCellObject {
    return (
        !!cell &&
        typeof cell === 'object' &&
        !(cell instanceof Date) &&
        !Array.isArray(cell)
    );
}

/** #RRGGBB -> ARGB（exceljs 需要 FF 前缀） */
function toArgb(color: string | undefined): string | undefined {
    if (!color) return undefined;
    const hex = color.replace(/^#/, '').toUpperCase();
    if (hex.length === 6) return `FF${hex}`;
    if (hex.length === 8) return hex;
    return undefined;
}

/** 统一边框线型 -> exceljs 线型 */
const BORDER_STYLE: Record<string, string> = {
    thin: 'thin',
    medium: 'medium',
    thick: 'thick',
    dashed: 'dashed',
    dotted: 'dotted',
    double: 'double',
    none: 'none'
};

/** 应用单元格样式 */
function applyStyle(cell: ExcelJS.Cell, style: XlsxCellStyle | undefined): void {
    if (!style) return;

    if (style.font) {
        cell.font = {
            name: style.font.name,
            size: style.font.size,
            bold: style.font.bold,
            italic: style.font.italic,
            underline: style.font.underline,
            strike: style.font.strike,
            ...(toArgb(style.font.color) ? { color: { argb: toArgb(style.font.color) } } : {})
        };
    }

    const fill = toArgb(style.fill);
    if (fill) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
    }

    if (style.align || style.valign || style.wrapText) {
        cell.alignment = {
            horizontal: style.align,
            vertical: style.valign,
            wrapText: style.wrapText
        };
    }

    if (style.numFmt) cell.numFmt = style.numFmt;

    if (style.border) {
        const border: Record<string, { style: string; color?: { argb: string } }> = {};
        for (const side of ['top', 'right', 'bottom', 'left'] as const) {
            const line = style.border[side];
            if (!line) continue;
            const lineStyle = BORDER_STYLE[line.style || 'thin'] || 'thin';
            const color = toArgb(line.color);
            border[side] = color ? { style: lineStyle, color: { argb: color } } : { style: lineStyle };
        }
        if (Object.keys(border).length) cell.border = border as never;
    }
}

/** 写入单元格值 */
function applyValue(cell: ExcelJS.Cell, source: XlsxCell): void {
    if (source === null || source === undefined) return;

    if (Array.isArray(source)) {
        // 富文本片段数组
        cell.value = {
            richText: source.map((part) => ({
                text: part.text,
                ...(part.font ? { font: applyFont(part.font) } : {})
            }))
        } as never;
        return;
    }

    if (!isCellObject(source)) {
        cell.value = source as ExcelJS.CellValue;
        return;
    }

    const obj = source;

    if (obj.richText) {
        cell.value = {
            richText: obj.richText.map((part) => ({
                text: part.text,
                ...(part.font ? { font: applyFont(part.font) } : {})
            }))
        } as never;
    } else if (obj.formula !== undefined) {
        cell.value = { formula: obj.formula, result: obj.result ?? undefined } as never;
    } else if (obj.hyperlink) {
        cell.value = { text: String(obj.value ?? obj.text ?? ''), hyperlink: obj.hyperlink } as never;
    } else if (obj.value !== undefined) {
        cell.value = obj.value as ExcelJS.CellValue;
    } else if (obj.text !== undefined) {
        cell.value = String(obj.text);
    }

    applyStyle(cell, obj.style);
}

/** XlsxFontStyle -> exceljs font */
function applyFont(font: NonNullable<XlsxCellStyle['font']>): ExcelJS.Font {
    const color = toArgb(font.color);
    return {
        name: font.name,
        size: font.size,
        bold: font.bold,
        italic: font.italic,
        underline: font.underline,
        strike: font.strike,
        ...(color ? { color: { argb: color } } : {})
    } as ExcelJS.Font;
}

/** 归一化后的行 */
interface NormalizedRow {
    /** 列号 -> 单元格 */
    cells: Map<number, XlsxCell>;
    /** 行高（磅） */
    height?: number;
    /** 是否隐藏 */
    hidden?: boolean;
}

/** 把一行输入归一化为列映射（同时兼容数组式 / 对象式 / 解析后的 XlsxRow） */
function normalizeRow(row: XlsxRowInput | XlsxRow): NormalizedRow {
    const cells = new Map<number, XlsxCell>();
    const result: NormalizedRow = { cells };

    if (Array.isArray(row)) {
        row.forEach((cell, index) => {
            if (cell !== undefined && cell !== null) cells.set(index + 1, cell);
        });
        return result;
    }

    if (row && typeof row === 'object') {
        const record = row as Record<string, unknown>;

        // 解析结果直传：{ index, cells, height, hidden }
        if (Array.isArray(record.cells)) {
            (record.cells as (XlsxCell | null)[]).forEach((cell, index) => {
                if (cell !== undefined && cell !== null) cells.set(index + 1, cell);
            });
            if (record.height !== undefined) result.height = Number(record.height);
            if (record.hidden) result.hidden = true;
            return result;
        }

        for (const [key, cell] of Object.entries(record)) {
            // 键支持三种写法：纯列号 '3'、列字母 'C'、A1 记号 'C3'
            const keyText = String(key).trim();
            let col = Number(keyText);
            if (!Number.isInteger(col) || col <= 0) col = columnFromRef(keyText);
            if (!Number.isInteger(col) || col <= 0) continue;
            cells.set(col, cell);
        }
    }
    return result;
}

/**
 * 生成 XLSX 工作簿
 * @param input - XlsxWorkbook，或直接传 sheets 数组（简写）
 * @param options - 文档属性
 * @returns XLSX 字节数据
 *
 * @example
 * ```js
 * const bytes = await jsonToXlsx([
 *   { name: '订单', rows: [
 *       ['姓名', '金额'],
 *       ['张三', { value: 1200, style: { numFmt: '#,##0.00' } }]
 *   ]}
 * ]);
 * ```
 */
export async function generateXlsx(
    input: XlsxWorkbook | XlsxWorkbook['sheets'],
    options: XlsxGenerateOptions = {}
): Promise<Uint8Array> {
    try {
        const spec: XlsxWorkbook = Array.isArray(input) ? { sheets: input } : input;
        if (!spec || !Array.isArray(spec.sheets) || !spec.sheets.length) {
            throw new TypeError('缺少 sheets 内容');
        }

        const workbook = new ExcelJS.Workbook();
        if (options.creator || spec.metadata?.creator) workbook.creator = options.creator || spec.metadata!.creator!;
        if (spec.metadata?.lastModifiedBy) workbook.lastModifiedBy = spec.metadata.lastModifiedBy;
        const created = options.created || spec.metadata?.created;
        if (created) {
            const date = new Date(created);
            if (!Number.isNaN(date.getTime())) workbook.created = date;
        }
        if (spec.metadata?.modified) {
            const modified = new Date(spec.metadata.modified);
            if (!Number.isNaN(modified.getTime())) workbook.modified = modified;
        }

        for (const sheetSpec of spec.sheets) {
            const worksheet = workbook.addWorksheet(sheetSpec.name || `Sheet${workbook.worksheets.length + 1}`, {
                state: sheetSpec.hidden ? 'hidden' : 'visible'
            });

            // 列宽/隐藏列
            if (sheetSpec.columns?.length) {
                worksheet.columns = sheetSpec.columns.map((col) => ({
                    width: col.width,
                    hidden: col.hidden
                }));
            }

            let maxColumn = sheetSpec.columns?.length || 0;

            sheetSpec.rows.forEach((rowInput, rowIndex) => {
                const row = normalizeRow(rowInput);
                for (const [col, source] of row.cells) {
                    if (col > maxColumn) maxColumn = col;
                    applyValue(worksheet.getCell(rowIndex + 1, col), source);
                }
                if (row.height !== undefined || row.hidden) {
                    const target = worksheet.getRow(rowIndex + 1);
                    if (row.height !== undefined) target.height = row.height;
                    if (row.hidden) target.hidden = true;
                }
            });

            if (sheetSpec.merges?.length) {
                for (const range of sheetSpec.merges) worksheet.mergeCells(range);
            }
            if (sheetSpec.frozen) {
                worksheet.views = [
                    {
                        state: 'frozen',
                        xSplit: sheetSpec.frozen.columns || 0,
                        ySplit: sheetSpec.frozen.rows || 0
                    }
                ];
            }
            if (sheetSpec.orientation) {
                worksheet.pageSetup = { ...(worksheet.pageSetup as object), orientation: sheetSpec.orientation };
            }
        }

        // exceljs 的 writeBuffer 只接受 JSZip 生成选项，压缩与否由 zip 选项透传
        const buffer = (await workbook.xlsx.writeBuffer(options.zip === false ? {} : {})) as ArrayBuffer;
        return new Uint8Array(buffer);
    } catch (error) {
        throw wrapError('xlsx', '生成', error);
    }
}

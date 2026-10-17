/**
 * XLSX 解析与生成公开入口
 *
 * @module xlsx
 */

import { wrapError } from '../core/errors';
import type { DocInput } from '../core/input';
import { parseXlsx } from './parse';
import { generateXlsx } from './generate';
import { xlsxToHtmlString } from './to-html';
import type { XlsxGenerateOptions, XlsxMetadata, XlsxParseOptions, XlsxSheet, XlsxWorkbook } from './types';

export * from './types';

/** 解析结果 */
export interface XlsxResult {
    /** 工作簿元数据 */
    metadata?: XlsxMetadata;
    /** 所有工作表 */
    sheets: XlsxSheet[];
    /** 解析提示 */
    messages: string[];
}

/** HTML 解析结果 */
export interface XlsxHtmlResult extends XlsxResult {
    /** 渲染好的 HTML（含默认样式表） */
    html: string;
}

/**
 * 解析 XLSX 为工作簿结构
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 解析选项（可指定只解析某些表）
 *
 * @example
 * ```js
 * const { sheets } = await xlsxToJson(file);
 * console.log(sheets[0].name, sheets[0].rows.length);
 * ```
 */
export async function xlsxToJson(
    input: DocInput,
    options: XlsxParseOptions = {}
): Promise<XlsxResult> {
    try {
        const { workbook, messages } = await parseXlsx(input, options);
        return { metadata: workbook.metadata, sheets: workbook.sheets, messages };
    } catch (error) {
        throw wrapError('xlsx', '解析', error);
    }
}

/**
 * 解析 XLSX 为 HTML 表格
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 解析选项
 *
 * @example
 * ```js
 * const { html } = await xlsxToHtml(file);
 * container.innerHTML = html;
 * ```
 */
export async function xlsxToHtml(
    input: DocInput,
    options: XlsxParseOptions & { includeStyle?: boolean } = {}
): Promise<XlsxHtmlResult> {
    const result = await xlsxToJson(input, options);
    const html = xlsxToHtmlString(
        { metadata: result.metadata, sheets: result.sheets },
        { includeStyle: options.includeStyle !== false }
    );
    return { ...result, html };
}

/**
 * 提取指定工作表为二维数组（适合直接喂给图表库）
 * @param input - 文件二进制
 * @param sheetName - 表名；缺省取第一张表
 * @returns { headers, rows }：headers 为首行，rows 为其余行的原始值
 *
 * @example
 * ```js
 * const { headers, rows } = await xlsxToMatrix(file, 'Sheet1');
 * ```
 */
export async function xlsxToMatrix(
    input: DocInput,
    sheetName?: string
): Promise<{ headers: unknown[]; rows: unknown[][] }> {
    const { sheets } = await xlsxToJson(input, { skipStyles: true });
    if (!sheets.length) return { headers: [], rows: [] };

    const sheet = sheetName ? sheets.find((item) => item.name === sheetName) : sheets[0];
    if (!sheet) throw wrapError('xlsx', '解析', new Error(`未找到工作表：${sheetName}`));

    const matrix = sheet.rows.map((row) => {
        const cells = Array.isArray(row) ? row : ((row as { cells?: unknown[] }).cells ?? []);
        return cells.map((cell) => {
            if (cell && typeof cell === 'object' && !(cell instanceof Date) && !Array.isArray(cell)) {
                const obj = cell as { value?: unknown; result?: unknown };
                return obj.value !== undefined ? obj.value : obj.result;
            }
            return cell as unknown;
        });
    });

    const [headers = [], ...rows] = matrix;
    return { headers, rows };
}

/**
 * 生成 XLSX 文件
 * @param input - XlsxWorkbook，或直接传 sheets 数组（简写）
 * @param options - 文档属性
 *
 * @example
 * ```js
 * const bytes = await jsonToXlsx([
 *   { name: '订单', rows: [['姓名', '金额'], ['张三', 1200]] }
 * ]);
 * ```
 */
export async function jsonToXlsx(
    input: XlsxWorkbook | XlsxWorkbook['sheets'],
    options: XlsxGenerateOptions = {}
): Promise<Uint8Array> {
    try {
        return await generateXlsx(input, options);
    } catch (error) {
        throw wrapError('xlsx', '生成', error);
    }
}

/**
 * 把工作簿结构渲染为 HTML（不经过文件解析，便于从 JSON 直接预览）
 * @param input - XlsxWorkbook 或 sheets 数组
 */
export function sheetsToHtml(
    input: XlsxWorkbook | XlsxWorkbook['sheets'],
    options: { includeStyle?: boolean } = {}
): string {
    return xlsxToHtmlString(input, options);
}

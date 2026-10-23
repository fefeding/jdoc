/**
 * XLSX -> 文档块模型（DocxBlock）
 *
 * XLSX 的内容模型是「工作表 + 单元格网格」，与 docx/md/html 的块序列不同，
 * 因此这里把每张工作表映射成「标题 + 表格」：
 *
 *   工作表名（heading）-> 表格（table）
 *
 * 单元格的富文本、超链接、底纹与合并区域都会尽量保留，
 * 这样 xlsx -> md / html / docx / pdf 四条目标路径都能拿到接近原表的内容。
 *
 * @module convert/xlsx
 */

import type { DocInput } from '../core/input';
import type { DocAsset, DocAssetMode } from '../core/asset';
import { xlsxToJson } from '../xlsx/index';
import { cellText } from '../xlsx/to-html';
import { resolveRefValues, xlsxChartToSpec, type XlsxChartSeriesData } from '../xlsx/chart';
import type { XlsxCellObject, XlsxMetadata, XlsxRow, XlsxSheet } from '../xlsx/types';
import type { DocxBlock, DocxCell, DocxCellDetail, DocxMetadata, DocxRun, DocxTable } from '../docx/types';

/** xlsxToBlocks 选项 */
export interface XlsxToBlocksOptions {
    /** 只转换指定工作表（缺省全部） */
    sheets?: string[];
    /** 是否把工作表名作为标题输出，默认 true */
    sheetHeading?: boolean;
    /** 首行是否作为表头（加粗、跨页重复），默认 true */
    header?: boolean;
    /** 单元格内图片的处理方式，默认 'inline'（图片作为独立图片块跟在表格后） */
    assets?: DocAssetMode;
    /** 是否跳过全空行，默认 true */
    skipEmptyRows?: boolean;
    /** 单表最多取多少行（防止超大表拖垮转换），默认 5000 */
    maxRows?: number;
}

/** xlsxToBlocks 结果 */
export interface XlsxToBlocksResult {
    /** 块序列 */
    blocks: DocxBlock[];
    /** 文档元数据（标题/作者等） */
    metadata?: DocxMetadata;
    /** 二进制资源（单元格图片） */
    assets?: DocAsset[];
}

/** 合并区域 */
interface MergeSpan {
    colspan: number;
    rowspan: number;
}

/** 'A1:C2' -> { startRow, startCol, endRow, endCol }（均为 1 起） */
function parseRange(range: string): { startRow: number; startCol: number; endRow: number; endCol: number } | undefined {
    const match = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(String(range).trim().toUpperCase());
    if (!match) return undefined;
    const toCol = (letters: string) => {
        let index = 0;
        for (const ch of letters) index = index * 26 + (ch.charCodeAt(0) - 64);
        return index;
    };
    const startRow = Number(match[2]);
    const startCol = toCol(match[1]);
    if (!match[3]) return { startRow, startCol, endRow: startRow, endCol: startCol };
    return { startRow, startCol, endRow: Number(match[4]), endCol: toCol(match[3]) };
}

/**
 * 合并区索引
 *
 * 主单元格记录 colspan / rowspan，被覆盖的坐标进 covered 集合（输出时跳过）。
 * 与 docx 的表格模型一致：被合并覆盖的位置不产单元格。
 */
function buildMergeIndex(sheet: XlsxSheet): { spans: Map<string, MergeSpan>; covered: Set<string> } {
    const spans = new Map<string, MergeSpan>();
    const covered = new Set<string>();
    for (const range of sheet.merges || []) {
        const parsed = parseRange(range);
        if (!parsed) continue;
        spans.set(`${parsed.startRow}:${parsed.startCol}`, {
            colspan: Math.max(1, parsed.endCol - parsed.startCol + 1),
            rowspan: Math.max(1, parsed.endRow - parsed.startRow + 1)
        });
        for (let r = parsed.startRow; r <= parsed.endRow; r++) {
            for (let c = parsed.startCol; c <= parsed.endCol; c++) {
                if (r === parsed.startRow && c === parsed.startCol) continue;
                covered.add(`${r}:${c}`);
            }
        }
    }
    return { spans, covered };
}

/** 单元格 -> run 序列（富文本保留样式，超链接保留链接） */
function cellToRuns(cell: XlsxCellObject | null | undefined): DocxRun[] {
    if (!cell) return [];
    if (cell.richText?.length) {
        return cell.richText.map((part) => {
            const run: DocxRun = { text: part.text ?? '' };
            const font = part.font;
            if (font?.bold) run.bold = true;
            if (font?.italic) run.italic = true;
            if (font?.underline) run.underline = true;
            if (font?.strike) run.strike = true;
            if (font?.color) run.color = font.color;
            if (font?.name) run.font = font.name;
            if (font?.size) run.size = font.size;
            return run;
        });
    }
    const text = cellText(cell);
    if (!text) return [];
    if (cell.hyperlink) return [{ text, link: cell.hyperlink }];
    return [{ text }];
}

/** 单元格 -> 表格单元格（带合并跨度与底纹） */
function cellToBlockCell(cell: XlsxCellObject | null | undefined, span?: MergeSpan): DocxCell {
    const runs = cellToRuns(cell);
    const detail: DocxCellDetail = { type: 'cell', text: runs.map((r) => r.text).join('') };
    if (runs.length) detail.runs = runs;
    if (span && span.colspan > 1) detail.colspan = span.colspan;
    if (span && span.rowspan > 1) detail.rowspan = span.rowspan;
    const fill = cell?.style?.fill;
    if (fill) detail.fill = fill;
    return detail;
}

/** 行是否全空（没有任何文本与图片） */
function isEmptyRow(row: XlsxRow): boolean {
    return !(row.cells || []).some((cell) => {
        if (!cell) return false;
        if (cell.image) return true;
        return cellToRuns(cell).some((r) => r.text.trim().length > 0);
    });
}

/** 工作表行归一（兼容数组式输入与对象式输入） */
function normalizeRows(sheet: XlsxSheet): XlsxRow[] {
    return sheet.rows.map((row, index) => {
        if (Array.isArray(row)) {
            return { index: index + 1, cells: row as (XlsxCellObject | null)[] };
        }
        const record = row as unknown as { index?: number; cells?: (XlsxCellObject | null)[] };
        if (Array.isArray(record.cells)) {
            return { index: record.index ?? index + 1, cells: record.cells };
        }
        // 对象式输入：{ 1: 'a', C: 'b' }
        const cells: (XlsxCellObject | null)[] = [];
        for (const [key, value] of Object.entries(row as Record<string, unknown>)) {
            const text = String(key).trim();
            let col = Number(text);
            if (!Number.isInteger(col) || col <= 0) {
                const letters = /^([A-Za-z]+)/.exec(text);
                if (!letters) continue;
                col = 0;
                for (const ch of letters[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
            }
            cells[col - 1] = value as XlsxCellObject;
        }
        return { index: index + 1, cells };
    });
}

/** 工作表 -> 表格块 */
function sheetToTable(
    sheet: XlsxSheet,
    rows: XlsxRow[],
    options: { header: boolean; skipEmptyRows: boolean }
): DocxBlock | undefined {
    const { spans, covered } = buildMergeIndex(sheet);
    const tableRows: DocxCell[][] = [];

    for (const row of rows) {
        if (options.skipEmptyRows && isEmptyRow(row)) continue;
        const cells: DocxCell[] = [];
        (row.cells || []).forEach((cell, colIndex) => {
            const col = colIndex + 1;
            if (covered.has(`${row.index}:${col}`)) return;
            cells.push(cellToBlockCell(cell, spans.get(`${row.index}:${col}`)));
        });
        tableRows.push(cells);
    }

    if (!tableRows.length) return undefined;
    return {
        type: 'table',
        // 只有一行时不存在「表头 + 数据行」的结构，标成表头反而会误导
        header: options.header && tableRows.length > 1,
        rows: tableRows
    };
}

/** 列号 -> 列字母（1 -> A，27 -> AA） */
function columnLetter(index: number): string {
    let n = Math.max(1, Math.floor(index));
    let letters = '';
    while (n > 0) {
        const rem = (n - 1) % 26;
        letters = String.fromCharCode(65 + rem) + letters;
        n = Math.floor((n - 1) / 26);
    }
    return letters;
}

/** 工作表内的单元格图片 -> 独立图片块（跟在表格之后） */
function sheetImages(sheet: XlsxSheet, rows: XlsxRow[]): DocxBlock[] {
    const images: DocxBlock[] = [];
    for (const row of rows) {
        (row.cells || []).forEach((cell, colIndex) => {
            if (!cell?.image) return;
            images.push({
                type: 'image',
                src: cell.image,
                assetId: cell.assetId,
                width: cell.imageWidth,
                height: cell.imageHeight,
                caption: `${sheet.name}!${columnLetter(colIndex + 1)}${row.index}`
            });
        });
    }
    return images;
}

/** 纯文本单元格（避免字面量 type 被拓宽成 string） */
function chartTextCell(text: string): DocxCell {
    return { type: 'cell', text };
}

/** EMU -> px（OOXML 尺寸单位，1 px = 9525 EMU） */
function emuToPx(value: number | undefined): number | undefined {
    if (!value || !Number.isFinite(value)) return undefined;
    return Math.round(value / 9525);
}

/** 引用取值 -> 数值（空值 / 非数字 -> null，缺值在图表里显示为断口） */
function toNumbers(values: (string | number | null)[]): (number | null)[] {
    return values.map((v) => {
        if (typeof v === 'number') return Number.isFinite(v) ? v : null;
        if (v === null || v === undefined || v === '') return null;
        // 单元格里可能是带货币符号 / 千分位的显示文本
        const n = Number(String(v).replace(/[,¥$€%\s]/g, ''));
        return Number.isFinite(n) ? n : null;
    });
}

/**
 * 工作表内的原生图表 -> 图表块
 *
 * Excel 的图表不在单元格网格里（存在 xl/charts/chartN.xml + 绘图锚点），
 * 因此表格转换天然看不到它——这里按引用回查当前单元格取值后转成 ChartSpec，
 * HTML 目标渲染成矢量 SVG，其余目标退化成兜底数据表。
 */
function sheetCharts(sheet: XlsxSheet, allSheets: XlsxSheet[]): DocxBlock[] {
    return (sheet.charts || []).map((chart) => {
        const data: XlsxChartSeriesData[] = (chart.series || []).map((s) => {
            const cats = s.catRef
                ? resolveRefValues(s.catRef, allSheets, sheet.name)
                : (s.cats ?? []);
            const valsRef = s.valRef ? resolveRefValues(s.valRef, allSheets, sheet.name) : undefined;
            const nameRef = s.nameRef ? resolveRefValues(s.nameRef, allSheets, sheet.name) : undefined;
            return {
                name: (nameRef?.[0] ?? s.name) ? String(nameRef?.[0] ?? s.name ?? '') : undefined,
                cats,
                vals: toNumbers(valsRef ?? s.vals ?? []),
                xVals: s.xRef ? toNumbers(resolveRefValues(s.xRef, allSheets, sheet.name)) : s.xVals,
                color: s.color
            };
        });
        const spec = xlsxChartToSpec(chart, data);
        const categories = data[0]?.cats ?? [];
        const fallback: DocxTable = {
            type: 'table',
            header: true,
            rows: [
                [chartTextCell(''), ...data.map((d) => chartTextCell(d.name ?? ''))],
                ...categories.map((cat, i) => [
                    chartTextCell(cat === null || cat === undefined ? '' : String(cat)),
                    ...data.map((d) => chartTextCell(String(d.vals?.[i] ?? '')))
                ])
            ]
        };
        return {
            type: 'chart',
            spec,
            width: emuToPx(chart.anchor?.ext?.cx),
            height: emuToPx(chart.anchor?.ext?.cy),
            caption: chart.title,
            fallback
        } as DocxBlock;
    });
}

/** XLSX 元数据 -> 文档元数据 */
function toDocxMetadata(metadata?: XlsxMetadata): DocxMetadata | undefined {
    if (!metadata) return undefined;
    const result: DocxMetadata = {};
    const put = (key: keyof DocxMetadata, value: unknown) => {
        if (value === undefined || value === null || value === '') return;
        result[key] = String(value);
    };
    put('title', metadata.title);
    put('author', metadata.author ?? metadata.creator);
    put('subject', metadata.subject);
    put('keywords', metadata.keywords);
    put('description', metadata.description);
    put('creator', metadata.creator);
    put('lastModifiedBy', metadata.lastModifiedBy);
    put('company', metadata.company);
    return Object.keys(result).length ? result : undefined;
}

/**
 * 把 XLSX 解析为文档块模型
 *
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 工作表筛选、表头、图片等选项
 * @returns 块序列与元数据
 *
 * @example
 * ```js
 * const { blocks } = await xlsxToBlocks(file);
 * console.log(blocks[0]); // { type: 'heading', level: 1, text: 'Sheet1' }
 * ```
 */
export async function xlsxToBlocks(
    input: DocInput,
    options: XlsxToBlocksOptions = {}
): Promise<XlsxToBlocksResult> {
    const {
        sheets,
        assets: assetsMode = 'inline',
        header = true,
        sheetHeading = true,
        skipEmptyRows = true,
        maxRows = 5000
    } = options;

    const parsed = await xlsxToJson(input, {
        sheets,
        assets: assetsMode,
        includeRichText: true,
        includeHyperlinks: true
    });

    const blocks: DocxBlock[] = [];
    for (const sheet of parsed.sheets) {
        if (sheetHeading && sheet.name) {
            blocks.push({ type: 'heading', level: 1, text: sheet.name });
        }
        // 行数上限对表格与图片同时生效，避免超大表拖垮转换
        const rows = normalizeRows(sheet).slice(0, Math.max(1, maxRows));
        const table = sheetToTable(sheet, rows, { header, skipEmptyRows });
        if (table) blocks.push(table);
        blocks.push(...sheetImages(sheet, rows));
        blocks.push(...sheetCharts(sheet, parsed.sheets));
    }

    return {
        blocks,
        metadata: toDocxMetadata(parsed.metadata),
        assets: parsed.assets?.length ? parsed.assets : undefined
    };
}

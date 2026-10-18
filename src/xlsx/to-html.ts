/**
 * XLSX 结构 -> HTML
 *
 * 用于表格预览：合并单元格按 rowspan/colspan 展开，数值按 numFmt 尽量格式化。
 *
 * @module xlsx/to-html
 */

import {
    assetToDataUrl,
    getAsset,
    isAssetRef,
    isDocAsset,
    sniffMime,
    toAssetTable,
    type AssetTable,
    type DocAsset
} from '../core/asset';
import type { XlsxCell, XlsxCellObject, XlsxRow, XlsxSheet, XlsxWorkbook } from './types';

/** HTML 转义 */
export function escapeHtml(value: unknown): string {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

/** 'A1:C2' -> { startRow, startCol, endRow, endCol }（均为 1 起） */
function parseRange(range: string) {
    const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(String(range).toUpperCase());
    if (!match) return undefined;
    const toCol = (letters: string) => {
        let index = 0;
        for (const ch of letters) index = index * 26 + (ch.charCodeAt(0) - 64);
        return index;
    };
    return {
        startCol: toCol(match[1]),
        startRow: Number(match[2]),
        endCol: toCol(match[3]),
        endRow: Number(match[4])
    };
}

/** 单元格 -> 显示文本 */
export function cellText(cell: XlsxCell | null | undefined): string {
    if (cell === null || cell === undefined) return '';
    if (Array.isArray(cell)) return cell.map((part) => part.text).join('');
    if (cell instanceof Date) return cell.toISOString().slice(0, 10);
    if (typeof cell !== 'object') return String(cell);

    const obj = cell as XlsxCellObject;
    if (obj.richText) return obj.richText.map((part) => part.text).join('');
    if (obj.formula !== undefined) {
        return obj.result === null || obj.result === undefined ? '' : String(obj.result);
    }
    if (obj.text !== undefined) return obj.text;
    if (obj.value === null || obj.value === undefined) return '';
    if (obj.value instanceof Date) return obj.value.toISOString().slice(0, 10);
    return String(obj.value);
}

/**
 * 单元格内容 -> HTML：图片走 <img>，其余按显示文本
 *
 * 图片可能是内联 data URL，也可能是资源引用（回查工作簿资源表）。
 */
export function cellContentToHtml(cell: XlsxCell | null | undefined, assets?: AssetTable): string {
    if (cell && typeof cell === 'object' && !Array.isArray(cell) && !(cell instanceof Date)) {
        const obj = cell as XlsxCellObject;
        if (obj.image) {
            const src = resolveImageSrc(obj.image, assets);
            const dims = [
                obj.imageWidth ? `width="${Math.round(obj.imageWidth * (96 / 72))}"` : '',
                obj.imageHeight ? `height="${Math.round(obj.imageHeight * (96 / 72))}"` : ''
            ]
                .filter(Boolean)
                .join(' ');
            return src ? `<img src="${escapeHtml(src)}" ${dims} alt="" style="max-width:100%"/>` : '';
        }
    }
    return escapeHtml(cellText(cell));
}

/** 单元格图片 src -> 可放进 <img> 的地址 */
function resolveImageSrc(image: NonNullable<XlsxCellObject['image']>, assets?: AssetTable): string | undefined {
    if (typeof image === 'string') return image;
    const resolved = resolveAssetBytesSync(image, assets);
    return resolved;
}

/** 同步解析图片字节（HTML 渲染路径要求同步） */
function resolveAssetBytesSync(
    value: unknown,
    assets?: AssetTable
): string | undefined {
    if (isAssetRef(value)) {
        const ref = value as { assetId?: string; asset?: DocAsset };
        const asset = ref.asset ?? getAsset(assets, ref.assetId);
        return asset ? assetToDataUrl(asset) : undefined;
    }
    if (isDocAsset(value)) return assetToDataUrl(value);
    if (value instanceof Uint8Array || value instanceof ArrayBuffer) {
        const bytes = value instanceof Uint8Array ? value : new Uint8Array(value);
        return assetToDataUrl(bytes, sniffMime(bytes));
    }
    return undefined;
}

/** 单元格 -> inline style */
function cellStyle(cell: XlsxCell | null | undefined): string {
    if (!cell || typeof cell !== 'object' || cell instanceof Date || Array.isArray(cell)) return '';
    const style = (cell as XlsxCellObject).style;
    if (!style) return '';

    const parts: string[] = [];
    if (style.font?.bold) parts.push('font-weight:600');
    if (style.font?.italic) parts.push('font-style:italic');
    if (style.font?.color) parts.push(`color:${style.font.color}`);
    if (style.font?.size) parts.push(`font-size:${style.font.size}pt`);
    if (style.fill) parts.push(`background-color:${style.fill}`);
    if (style.align) parts.push(`text-align:${style.align}`);
    if (style.valign) parts.push(`vertical-align:${style.valign}`);
    if (style.wrapText) parts.push('white-space:pre-wrap');
    return parts.length ? ` style="${escapeHtml(parts.join(';'))}"` : '';
}

/** 工作表 -> HTML table */
export function sheetToHtml(sheet: XlsxSheet, assets?: AssetTable): string {
    // 解析结果形式：rows 是 XlsxRow；输入形式：XlsxRowInput
    const normalized: XlsxRow[] = sheet.rows.map((row, index) => {
        if (Array.isArray(row)) {
            return { index: index + 1, cells: row as (XlsxCellObject | null)[] };
        }
        const record = row as unknown as { index?: number; cells?: (XlsxCellObject | null)[] };
        if (Array.isArray(record.cells)) {
            return { index: record.index ?? index + 1, cells: record.cells };
        }
        // 对象式输入：{ 1: 'a', C: 'b' }
        const cells: (XlsxCellObject | null)[] = [];
        for (const [key, value] of Object.entries(row as Record<string, XlsxCell>)) {
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

    // 合并区：主单元格所在坐标 -> rowspan/colspan，被覆盖的坐标跳过
    const spans = new Map<string, { rowspan: number; colspan: number }>();
    const covered = new Set<string>();
    for (const range of sheet.merges || []) {
        const parsed = parseRange(range);
        if (!parsed) continue;
        spans.set(`${parsed.startRow}:${parsed.startCol}`, {
            rowspan: parsed.endRow - parsed.startRow + 1,
            colspan: parsed.endCol - parsed.startCol + 1
        });
        for (let r = parsed.startRow; r <= parsed.endRow; r++) {
            for (let c = parsed.startCol; c <= parsed.endCol; c++) {
                if (r === parsed.startRow && c === parsed.startCol) continue;
                covered.add(`${r}:${c}`);
            }
        }
    }

    const body = normalized
        .map((row) => {
            const cells: string[] = [];
            row.cells.forEach((cell, colIndex) => {
                const col = colIndex + 1;
                if (covered.has(`${row.index}:${col}`)) return;
                const span = spans.get(`${row.index}:${col}`);
                const attrs = [
                    span ? ` rowspan="${span.rowspan}"` : '',
                    span && span.colspan > 1 ? ` colspan="${span.colspan}"` : ''
                ].join('');
                cells.push(`<td${attrs}${cellStyle(cell)}>${cellContentToHtml(cell, assets)}</td>`);
            });
            return `<tr>${cells.join('')}</tr>`;
        })
        .join('');

    return `<table class="xlsx-table">${body}</table>`;
}

/** 默认样式表 */
export const XLSX_HTML_STYLE = `.xlsx-wrap{overflow:auto;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,"PingFang SC","Microsoft YaHei",sans-serif}
.xlsx-title{font-weight:600;margin:10px 0 6px}
.xlsx-table{border-collapse:collapse;font-size:12px;min-width:100%}
.xlsx-table td{border:1px solid #d0d7de;padding:4px 8px;white-space:pre-wrap}
`;

/**
 * 把工作簿渲染为 HTML
 * @param input - XlsxWorkbook 或 sheets 数组
 * @param options.includeStyle - 是否附带默认样式表
 */
export function xlsxToHtmlString(
    input: XlsxWorkbook | XlsxWorkbook['sheets'],
    options: { includeStyle?: boolean } = {}
): string {
    const sheets = Array.isArray(input) ? input : input?.sheets || [];
    const assets = Array.isArray(input) || !input?.assets?.length ? undefined : toAssetTable(input.assets);
    const body = sheets
        .map(
            (sheet) =>
                `<div class="xlsx-sheet"><div class="xlsx-title">${escapeHtml(sheet.name)}</div>${sheetToHtml(sheet, assets)}</div>`
        )
        .join('');
    const style = options.includeStyle === false ? '' : `<style>${XLSX_HTML_STYLE}</style>`;
    return `${style}<div class="xlsx-wrap">${body}</div>`;
}

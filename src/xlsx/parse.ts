/**
 * XLSX 解析
 *
 * 基于 exceljs 读取工作簿，把 exceljs 的单元格模型收敛到 jdoc 统一的
 * XlsxCellObject 结构，使解析结果可直接回传给 jsonToXlsx 实现往返。
 *
 * @module xlsx/parse
 */

import ExcelJS from 'exceljs';
import { wrapError } from '../core/errors';
import { toUint8Array, type DocInput } from '../core/input';
import {
    assetToDataUrl,
    createAsset,
    createAssetId,
    mimeOf,
    sniffMime,
    type DocAsset,
    type DocAssetMode
} from '../core/asset';
import type {
    XlsxBorderStyle,
    XlsxCellObject,
    XlsxCellStyle,
    XlsxMetadata,
    XlsxParseOptions,
    XlsxRichText,
    XlsxRow,
    XlsxSheet,
    XlsxWorkbook
} from './types';

/** 列号 -> 列字母：1 -> A，27 -> AA */
export function columnLetter(index: number): string {
    let n = Math.max(1, Math.floor(index));
    let letters = '';
    while (n > 0) {
        const rem = (n - 1) % 26;
        letters = String.fromCharCode(65 + rem) + letters;
        n = Math.floor((n - 1) / 26);
    }
    return letters;
}

/** exceljs 边框线型 -> 统一枚举 */
function normalizeBorder(style: unknown): XlsxBorderStyle['style'] {
    const value = String(style || '').toLowerCase();
    if (!value || value === 'none') return undefined;
    if (value.includes('double')) return 'double';
    if (value.includes('thick')) return 'thick';
    if (value.includes('medium')) return 'medium';
    if (value.includes('dashed') || value.includes('dash')) return 'dashed';
    if (value.includes('dotted') || value.includes('dot')) return 'dotted';
    return 'thin';
}

/**
 * 颜色归一化为 #RRGGBB。
 * exceljs 的颜色有三种形态：裸字符串、{ argb }、{ theme, tint }。
 * theme 配色依赖主题表的完整上下文，无法稳定还原，这里直接忽略。
 */
function normalizeArgb(color: unknown): string | undefined {
    if (!color) return undefined;

    // { argb: 'FFFF0000' } 形态
    if (typeof color === 'object') {
        const argb = (color as { argb?: unknown }).argb;
        if (argb === undefined || argb === null) return undefined;
        return normalizeArgb(argb);
    }

    const raw = String(color).replace(/^#/, '');
    if (!/^[0-9a-fA-F]{6}$/.test(raw) && !/^[0-9a-fA-F]{8}$/.test(raw)) return undefined;
    return `#${raw.slice(-6).toUpperCase()}`;
}

/** exceljs cell.style -> XlsxCellStyle */
function readStyle(cell: ExcelJS.Cell, skip: boolean): XlsxCellStyle | undefined {
    if (skip) return undefined;
    const style = cell.style;
    if (!style) return undefined;

    const result: XlsxCellStyle = {};
    const font = style.font as unknown as Record<string, unknown> | undefined;
    if (font) {
        const mapped: Record<string, unknown> = {};
        if (font.name) mapped.name = font.name;
        if (font.size) mapped.size = Number(font.size);
        if (font.bold) mapped.bold = true;
        if (font.italic) mapped.italic = true;
        if (font.underline) mapped.underline = true;
        if (font.strike) mapped.strike = true;
        const color = normalizeArgb(font.color);
        if (color) mapped.color = color;
        if (Object.keys(mapped).length) result.font = mapped as XlsxCellStyle['font'];
    }

    const fill = style.fill as unknown as Record<string, unknown> | undefined;
    if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
        const fg = normalizeArgb(fill.fgColor);
        if (fg) result.fill = fg;
    }

    const alignment = style.alignment as unknown as Record<string, unknown> | undefined;
    if (alignment) {
        const h = String(alignment.horizontal || '').toLowerCase();
        if (h === 'left' || h === 'center' || h === 'right') result.align = h;
        const v = String(alignment.vertical || '').toLowerCase();
        if (v === 'top' || v === 'middle' || v === 'bottom') result.valign = v;
        if (alignment.wrapText) result.wrapText = true;
    }

    if (style.numFmt) result.numFmt = String(style.numFmt);

    const borders = style.border as unknown as Record<string, Record<string, unknown>> | undefined;
    if (borders) {
        const map: NonNullable<XlsxCellStyle['border']> = {};
        for (const side of ['top', 'right', 'bottom', 'left'] as const) {
            const line = borders[side];
            if (!line || !line.style) continue;
            map[side] = { style: normalizeBorder(line.style), color: normalizeArgb(line.color) };
        }
        if (Object.keys(map).length) result.border = map;
    }

    return Object.keys(result).length ? result : undefined;
}

/** exceljs 单元格值 -> XlsxCellObject */
function readCell(cell: ExcelJS.Cell, options: XlsxParseOptions): XlsxCellObject | null {
    const value = cell.value as unknown;
    const style = readStyle(cell, !!options.skipStyles);
    const result: XlsxCellObject = {};

    const isEmpty =
        value === null ||
        value === undefined ||
        (typeof value === 'object' &&
            !Array.isArray(value) &&
            Object.keys(value as object).length === 0);

    if (!isEmpty) {
        if (typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date)) {
            const record = value as Record<string, unknown>;

            if (Array.isArray(record.richText)) {
                if (options.includeRichText !== false) {
                    result.richText = record.richText as XlsxRichText[];
                } else {
                    result.value = record.richText
                        .map((part) => (part as XlsxRichText).text)
                        .join('');
                }
            } else if (record.formula !== undefined || record.sharedFormula !== undefined) {
                // sharedFormula 的公式体在主单元格上，这里保留引用关系
                result.formula = String(record.formula ?? record.sharedFormula ?? '');
                result.result = (record.result as XlsxCellObject['value']) ?? null;
            } else if (record.hyperlink !== undefined) {
                result.value = (record.text as XlsxCellObject['value']) ?? '';
                if (options.includeHyperlinks !== false) result.hyperlink = String(record.hyperlink);
            } else if (record.error !== undefined) {
                result.value = String(record.error);
            } else if (record.text !== undefined) {
                result.value = record.text as XlsxCellObject['value'];
            } else {
                const first = Object.values(record)[0];
                result.value = (first as XlsxCellObject['value']) ?? null;
            }
        } else if (Array.isArray(value)) {
            // 多行文本被 exceljs 表达为字符串数组
            result.value = (value as unknown[])
                .map((part) => String((part as { text?: string })?.text ?? part))
                .join('\n');
        } else if (value instanceof Date) {
            result.value = value;
        } else {
            result.value = value as XlsxCellObject['value'];
        }
    }

    if (style) result.style = style;
    if (options.includeNotes && cell.note) {
        const note = cell.note as { texts?: { text?: string }[]; text?: string };
        const noteText = note.texts ? note.texts.map((t) => t.text || '').join('') : note.text;
        if (noteText) result.note = noteText;
    }

    return Object.keys(result).length ? result : null;
}

/** 读取工作簿元数据 */
function readMetadata(workbook: ExcelJS.Workbook): XlsxMetadata {
    const metadata: XlsxMetadata = {};
    const put = (key: keyof XlsxMetadata, value: unknown) => {
        if (value === undefined || value === null || value === '') return;
        const text = value instanceof Date ? value.toISOString() : String(value);
        (metadata as Record<string, string>)[key] = text;
    };

    put('creator', workbook.creator);
    put('lastModifiedBy', workbook.lastModifiedBy);
    put('created', workbook.created);
    put('modified', workbook.modified);

    // title/subject 等写在 coreProperties 里（部分生成器会填）
    const props = (workbook as unknown as { model?: { properties?: Record<string, unknown> } }).model?.properties;
    if (props) {
        put('title', props.title);
        put('subject', props.subject);
        put('description', props.description);
        put('keywords', props.keywords);
        put('category', props.category);
        put('company', props.company);
        if (!metadata.author && props.lastModifiedBy) put('author', props.lastModifiedBy);
    }
    if (!metadata.author && workbook.creator) put('author', workbook.creator);

    return Object.keys(metadata).length ? metadata : undefined;
}

/** 读取单个工作表 */
/**
 * 读取工作表内嵌图片，挂到对应单元格并收集为资源
 *
 * exceljs 的 `worksheet.getImages()` 给出每张图的左上角锚点，
 * 资源表按媒体 id 去重，锚点记录「哪张表的哪个单元格」。
 */
function readSheetImages(
    worksheet: ExcelJS.Worksheet,
    sheet: XlsxSheet,
    ctx: { assets: DocAsset[]; byId: Map<string, string>; mode: DocAssetMode }
): void {
    if (ctx.mode === 'ignore') return;
    const images = worksheet.getImages();
    if (!images?.length) return;

    // exceljs 的 workbook.media 是数组，下标即 imageId（旧版本可能是 Map，这里兼容两种）
    const rawMedia = (worksheet.workbook as unknown as {
        media?: Array<{ type?: string; extension?: string; buffer?: Uint8Array }> | Map<number, { type?: string; extension?: string; buffer?: Uint8Array }>;
    }).media;
    const lookupMedia = (id: number) =>
        Array.isArray(rawMedia)
            ? rawMedia[id]
            : (rawMedia as Map<number, { type?: string; extension?: string; buffer?: Uint8Array }>)?.get(id);

    for (const image of images) {
        // exceljs 的类型把 imageId 声明为 string，而 media 表以 number 为键
        const mediaId = Number(image.imageId);
        const entry = lookupMedia(mediaId);
        const bytes = entry?.buffer ? new Uint8Array(entry.buffer) : undefined;
        if (!bytes || bytes.byteLength === 0) continue;

        const tl = image.range?.tl;
        // exceljs 的 tl.col / tl.row 为 0 基，这里转成 1 基的行列号
        const rowNo = (tl?.row ?? 0) + 1;
        const colNo = (tl?.col ?? 0) + 1;
        // exceljs 的 type 只是 'image' 这类粗粒度标记，需补全为完整 MIME
        const rawType = entry?.type && entry.type.includes('/') ? entry.type : '';
        const mime =
            rawType ||
            mimeOf(entry?.extension ? `x.${entry.extension}` : '', '') ||
            sniffMime(bytes);
        const id = createAssetId('xlsx', ctx.assets.length + 1);
        const asset = createAsset({
            id,
            data: bytes,
            mime,
            part: `xl/media/${mediaId}.${entry?.extension || 'png'}`,
            anchor: {
                scope: 'xlsx',
                kind: 'cell',
                index: rowNo,
                path: `${sheet.name}!${columnLetter(colNo)}${rowNo}`,
                label: `${sheet.name}!${columnLetter(colNo)}${rowNo}`
            }
        });
        ctx.assets.push(asset);
        ctx.byId.set(String(mediaId), id);

        // 挂到锚定单元格：单元格内容为空时才用图片占位，已有值则保留文字
        const target = sheet.rows[rowNo - 1] as XlsxRow | undefined;
        if (!target) continue;
        const cells = (target.cells ||= []);
        const cell = (cells[colNo - 1] ||= {});
        const dataUrl = ctx.mode === 'extract' ? { assetId: id } : assetToDataUrl(asset);
        cell.image = dataUrl as never;
        cell.assetId = id;
        // exceljs 的 ImageRange 类型只声明 tl/br，ext（像素尺寸）为可选扩展信息
        const ext = (image.range as { ext?: { width?: number; height?: number } } | undefined)?.ext;
        // 像素 -> 磅（96dpi）
        if (ext?.width) cell.imageWidth = Math.round(ext.width * 0.75);
        if (ext?.height) cell.imageHeight = Math.round(ext.height * 0.75);
    }
}

function readSheet(worksheet: ExcelJS.Worksheet, options: XlsxParseOptions, ctx?: {
    assets: DocAsset[];
    byId: Map<string, string>;
    mode: DocAssetMode;
}): XlsxSheet {
    const rows: XlsxRow[] = [];

    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const cells: (XlsxCellObject | null)[] = [];
        const columnCount = Math.max(row.cellCount, row.actualCellCount);
        for (let col = 1; col <= columnCount; col++) {
            cells[col - 1] = readCell(row.getCell(col), options);
        }
        // 压缩尾部空单元格，减少 JSON 体积
        while (cells.length && !cells[cells.length - 1]) cells.pop();

        const item: XlsxRow = { index: rowNumber, cells };
        if (row.height !== undefined && row.height !== 15) item.height = Number(row.height);
        if (row.hidden) item.hidden = true;
        rows.push(item);
    });

    // 压缩尾部空行
    while (rows.length && !rows[rows.length - 1].cells.length) rows.pop();

    const sheet: XlsxSheet = { name: worksheet.name, rows: rows as XlsxSheet['rows'] };

    const columns = (worksheet.columns || [])
        .map((col) => {
            const item: { width?: number; hidden?: boolean; header?: string } = {};
            if (col.width !== undefined) item.width = Number(col.width);
            if (col.hidden) item.hidden = true;
            return Object.keys(item).length ? item : undefined;
        })
        .filter(Boolean) as XlsxSheet['columns'];
    if (columns && columns.length) sheet.columns = columns;

    // exceljs：worksheet.model.merges 是完整范围 ref 数组（如 'A1:C1'）；
    // 而 _merges 的 key 只是主单元格地址（如 'A1'），不能直接当合并范围使用。
    const model = worksheet.model as unknown as { merges?: string[] } | undefined;
    let merges: string[] = Array.isArray(model?.merges) ? model!.merges!.slice() : [];
    if (!merges.length) {
        const raw = (worksheet as unknown as { _merges?: Record<string, { model?: { top: number; left: number; bottom: number; right: number } }> })._merges || {};
        const colLetter = (n: number): string => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
        merges = Object.keys(raw)
            .map((k) => {
                const m = raw[k]?.model;
                if (!m) return k;
                const a = colLetter(m.left) + m.top;
                const b = colLetter(m.right) + m.bottom;
                return a === b ? a : a + ':' + b;
            })
            .filter(Boolean);
    }
    if (merges.length) sheet.merges = merges;

    // views 的 xSplit/ySplit 属于「冻结」专属字段，联合类型里被抹平，这里取原始记录
    const view = worksheet.views?.[0] as unknown as { state?: string; xSplit?: number; ySplit?: number } | undefined;
    if (view && typeof view.ySplit === 'number' && view.ySplit > 0) {
        sheet.frozen = { rows: view.ySplit, columns: typeof view.xSplit === 'number' ? view.xSplit : 0 };
    }
    if (worksheet.state === 'hidden' || worksheet.state === 'veryHidden') sheet.hidden = true;
    if (worksheet.pageSetup?.orientation === 'landscape') sheet.orientation = 'landscape';

    if (ctx) readSheetImages(worksheet, sheet, ctx);

    return sheet;
}

/**
 * 解析 XLSX 为工作簿结构
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 解析选项
 */
export async function parseXlsx(
    input: DocInput,
    options: XlsxParseOptions = {}
): Promise<{ workbook: XlsxWorkbook; messages: string[] }> {
    const messages: string[] = [];
    try {
        const bytes = await toUint8Array(input);
        if (bytes.byteLength < 4) throw new TypeError('文件过小');

        const workbook = new ExcelJS.Workbook();
        // exceljs 需要 ArrayBuffer 或其视图，不能直接传 Blob
        await workbook.xlsx.load(bytes as unknown as ArrayBuffer);

        const assetCtx = {
            assets: [] as DocAsset[],
            byId: new Map<string, string>(),
            mode: (options.assets ?? 'ignore') as DocAssetMode
        };
        const wanted = options.sheets?.length ? new Set(options.sheets) : undefined;
        const sheets: XlsxSheet[] = [];
        for (const worksheet of workbook.worksheets) {
            if (wanted && !wanted.has(worksheet.name)) continue;
            sheets.push(readSheet(worksheet, options, assetCtx));
        }
        if (wanted) {
            for (const name of wanted) {
                if (!sheets.some((sheet) => sheet.name === name)) messages.push(`未找到工作表：${name}`);
            }
        }

        const metadata = readMetadata(workbook);
        const result: XlsxWorkbook = { metadata, sheets };
        if (assetCtx.assets.length) result.assets = assetCtx.assets;
        return { workbook: result, messages };
    } catch (error) {
        throw wrapError('xlsx', '解析', error);
    }
}

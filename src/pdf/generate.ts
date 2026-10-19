/**
 * PDF 生成
 *
 * 基于 pdf-lib 输出 PDF。坐标系对调用方保持「顶部原点」（与 HTML 心智模型一致），
 * 内部换算为 pdf-lib 的左下原点。
 *
 * 中文字体：pdf-lib 的 14 种标准字体不含 CJK 字形，需要通过 fonts 传入 TTF/OTF 字节；
 * 未传入时使用标准字体，此时中文会抛出明确错误而不是静默输出空白。
 *
 * @module pdf/generate
 */

import { PDFDocument, StandardFonts, rgb, type PDFFont, type PDFImage, type RGB } from 'pdf-lib';
import { JdocError, wrapError } from '../core/errors';
import {
    getAsset,
    isAssetRef,
    toAssetTable,
    type AssetTable,
    type DocAsset
} from '../core/asset';
import { toUint8Array } from '../core/input';
import type {
    PdfDocumentSpec,
    PdfElement,
    PdfFontSet,
    PdfGenerateOptions,
    PdfImageElement,
    PdfPageSize,
    PdfRun,
    PdfRunsElement,
    PdfTableCell,
    PdfTableElement,
    PdfTabStop,
    PdfTextElement
} from './types';
import { registerPdfFontkit } from './fonts';

/** 纸张尺寸（pt，1pt = 1/72 英寸） */
const PAGE_SIZES: Record<string, [number, number]> = {
    A3: [841.89, 1190.55],
    A4: [595.28, 841.89],
    A5: [419.53, 595.28],
    Letter: [612, 792],
    Legal: [612, 1008]
};

/** 命名颜色 */
const NAMED_COLORS: Record<string, string> = {
    black: '#000000',
    white: '#FFFFFF',
    red: '#FF0000',
    green: '#008000',
    blue: '#0000FF',
    gray: '#808080',
    grey: '#808080',
    silver: '#C0C0C0',
    yellow: '#FFFF00',
    orange: '#FFA500',
    purple: '#800080',
    navy: '#000080',
    teal: '#008080',
    maroon: '#800000',
    olive: '#808000',
    lime: '#00FF00',
    aqua: '#00FFFF',
    cyan: '#00FFFF',
    fuchsia: '#FF00FF',
    magenta: '#FF00FF'
};

/** 颜色字符串 -> RGB；无法解析时返回黑色 */
export function parseColor(value: string | undefined, fallback: RGB = rgb(0, 0, 0)): RGB {
    if (!value) return fallback;
    let hex = value.trim();

    const named = NAMED_COLORS[hex.toLowerCase()];
    if (named) hex = named;
    if (hex.startsWith('#')) hex = hex.slice(1);
    if (hex.length === 3) hex = hex.split('').map((ch) => ch + ch).join('');
    if (!/^[0-9a-fA-F]{6}$/.test(hex)) return fallback;

    return rgb(
        parseInt(hex.slice(0, 2), 16) / 255,
        parseInt(hex.slice(2, 4), 16) / 255,
        parseInt(hex.slice(4, 6), 16) / 255
    );
}

/** 纸张尺寸 -> [宽, 高]（已考虑横竖屏） */
export function resolvePageSize(
    size: PdfPageSize | undefined,
    landscape: boolean
): [number, number] {
    let width: number;
    let height: number;

    if (Array.isArray(size)) {
        [width, height] = size;
    } else if (size && typeof size === 'object') {
        width = size.width;
        height = size.height;
    } else {
        const key = typeof size === 'string' ? size : 'A4';
        [width, height] = PAGE_SIZES[key] || PAGE_SIZES.A4;
    }

    return landscape ? [height, width] : [width, height];
}

/** 页边距归一化 */
function resolveMargin(margin: PdfPageSpecLike['margin']): { top: number; right: number; bottom: number; left: number } {
    const DEFAULT = 56; // 约 2cm
    if (typeof margin === 'number') {
        return { top: margin, right: margin, bottom: margin, left: margin };
    }
    if (margin && typeof margin === 'object') {
        return {
            top: margin.top ?? DEFAULT,
            right: margin.right ?? DEFAULT,
            bottom: margin.bottom ?? DEFAULT,
            left: margin.left ?? DEFAULT
        };
    }
    return { top: DEFAULT, right: DEFAULT, bottom: DEFAULT, left: DEFAULT };
}

/** 页边距类型别名 */
type PdfPageSpecLike = { margin?: number | { top?: number; right?: number; bottom?: number; left?: number } };

/** 字体集合：按 normal/bold/italic/boldItalic 分类 */
export interface FontBundle {
    normal: PDFFont;
    bold: PDFFont;
    italic: PDFFont;
    boldItalic: PDFFont;
    /** 是否为自定义内嵌字体（决定中文是否可用） */
    embedded: boolean;
}

/** 标准字体族名 -> pdf-lib 常量 */
const STANDARD_FAMILIES = [
    { key: 'normal', constant: StandardFonts.Helvetica },
    { key: 'bold', constant: StandardFonts.HelveticaBold },
    { key: 'italic', constant: StandardFonts.HelveticaOblique },
    { key: 'boldItalic', constant: StandardFonts.HelveticaBoldOblique }
] as const;

/** 内嵌字体字节归一化 */
async function toFontBytes(source: string | ArrayBuffer | Uint8Array): Promise<Uint8Array> {
    if (typeof source !== 'string') return toUint8Array(source);
    // base64 或 data URL
    const match = /^data:[^;,]*(;base64)?,(.*)$/s.exec(source.trim());
    const payload = match ? match[2] : source;
    const clean = payload.replace(/\s+/g, '');
    const atobFn = (globalThis as { atob?: (s: string) => string }).atob;
    const binary = atobFn ? atobFn(clean) : Buffer.from(clean, 'base64').toString('binary');
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i) & 0xff;
    return bytes;
}

/** 构建字体集合 */
export async function buildFonts(
    doc: PDFDocument,
    fonts: PdfFontSet | undefined
): Promise<FontBundle> {
    if (fonts && (fonts.regular || fonts.bold)) {
        // 内嵌自定义字体必须注册 fontkit，否则 embedFont(TTF) 会抛 "Unknown font format"
        doc.registerFontkit((await registerPdfFontkit()) as never);

        const embed = async (source: string | ArrayBuffer | Uint8Array | undefined, fallback?: string) => {
            const bytes = source ? await toFontBytes(source) : fallback ? await toFontBytes(fallback) : undefined;
            return bytes ? doc.embedFont(bytes) : undefined;
        };

        const regular = await embed(fonts.regular);
        const normal = regular || (await doc.embedFont(StandardFonts.Helvetica));
        const bold = (await embed(fonts.bold, typeof fonts.regular === 'string' ? fonts.regular : undefined)) || normal;
        const italic = (await embed(fonts.italic, typeof fonts.regular === 'string' ? fonts.regular : undefined)) || normal;
        const boldItalic =
            (await embed(fonts.boldItalic, typeof fonts.bold === 'string' ? fonts.bold : undefined)) || bold;
        return { normal, bold, italic, boldItalic, embedded: true };
    }

    // 标准 14 字体：Helvetica + Times + Courier 共 4 组，取 Helvetica 族做默认
    const normal = await doc.embedFont(StandardFonts.Helvetica);
    const bold = await doc.embedFont(StandardFonts.HelveticaBold);
    const italic = await doc.embedFont(StandardFonts.HelveticaOblique);
    const boldItalic = await doc.embedFont(StandardFonts.HelveticaBoldOblique);
    return { normal, bold, italic, boldItalic, embedded: false };
}

/** 检测文本中是否含有 CJK 字符（标准字体无法渲染） */
function hasCjk(text: string): boolean {
    return /[\u2E80-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF]/.test(text);
}

/** 文本按宽度换行（中文按字断行，西文优先按空格断行） */
export function wrapText(text: string, font: PDFFont, size: number, maxWidth: number): string[] {
    if (!text) return [];
    if (!maxWidth || maxWidth <= 0) return text.split('\n');

    const lines: string[] = [];
    for (const paragraph of String(text).split('\n')) {
        if (!paragraph) {
            lines.push('');
            continue;
        }

        let line = '';
        for (const char of Array.from(paragraph)) {
            const candidate = line + char;
            let width: number;
            try {
                width = font.widthOfTextAtSize(candidate, size);
            } catch {
                // 标准字体遇到缺字形会抛错，降级为「放得下就继续」
                width = candidate.length * size * 0.5;
            }

            if (width <= maxWidth || !line) {
                line = candidate;
                continue;
            }

            // 超宽：在最后一个空格处断行，避免西文单词被劈开
            const lastSpace = line.lastIndexOf(' ');
            if (lastSpace > 0 && /[\x20-\x7E]/.test(char)) {
                lines.push(line.slice(0, lastSpace));
                line = line.slice(lastSpace + 1) + char;
            } else {
                lines.push(line);
                line = char;
            }
        }
        if (line) lines.push(line);
    }
    return lines;
}

/** 图片字节类型识别 */
function detectImageType(bytes: Uint8Array): 'png' | 'jpg' | undefined {
    if (bytes.length > 8 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return 'png';
    if (bytes.length > 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'jpg';
    return undefined;
}

/**
 * 图片源归一化为字节
 *
 * 支持内联 data URL / base64 / 二进制，以及资源引用 `{ assetId }`
 * （引用需在 `jsonToPdf` 的 `assets` 里给出）。
 */
async function readImageBytes(
    src: PdfImageElement['src'],
    assets?: AssetTable
): Promise<Uint8Array> {
    if (isAssetRef(src)) {
        const ref = src as { assetId?: string; asset?: DocAsset };
        const asset = ref.asset ?? getAsset(assets, ref.assetId);
        if (!asset) throw new TypeError('图片引用无法解析：未提供对应的 assets');
        return asset.data;
    }
    return toUint8Array(src as string);
}

/**
 * 自动分页的流式状态：绘制函数逐行询问剩余空间，不足即换页。
 *
 * 由 generatePdf 的自动分页循环创建并注入 DrawContext；显式 pages 模式无 flow，
 * 绘制函数保持原有「一次画完、返回总高」的行为。
 */
interface FlowState {
    /** 当前页（换页时被循环替换） */
    page: ReturnType<PDFDocument['addPage']>;
    /** 页顶边距（顶部原点的基准） */
    marginTop: number;
    /** 当前页内容光标（自内容区顶部起算，pt） */
    cursorY: number;
    /** 单页可用高度（pageHeight - 上边距 - 下边距） */
    usableHeight: number;
    /** 确保剩余空间足够容纳 need；不足则换页。返回是否发生了换页 */
    ensure: (need: number) => boolean;
}

/** 绘制上下文 */
interface DrawContext {
    doc: PDFDocument;
    fonts: FontBundle;
    /** 可用资源表：用于解析图片元素的 { assetId } 引用 */
    assets: AssetTable;
    page: ReturnType<PDFDocument['addPage']>;
    pageHeight: number;
    contentWidth: number;
    originX: number;
    originY: number;
    defaultSize: number;
    defaultColor: RGB;
    /** 自动分页流状态：存在时绘制函数逐行换页而非整块挪页 */
    flow?: FlowState;
}

/** 当前绘制目标页（flow 模式下随换页更新） */
function currentPage(ctx: DrawContext): DrawContext['page'] {
    return ctx.flow ? ctx.flow.page : ctx.page;
}

/** 顶部原点光标（不含元素 y 偏移；flow 模式取流光标） */
function topCursor(ctx: DrawContext, elementY = 0): number {
    return ctx.flow ? ctx.flow.marginTop + ctx.flow.cursorY : ctx.originY + elementY;
}

/** 在页面上绘制文本（含换行与对齐），返回占用高度 */
function drawTextElement(
    ctx: DrawContext,
    element: PdfTextElement | (Omit<PdfTextElement, 'type'> & { type: 'heading'; level: number })
): number {
    const isHeading = element.type === 'heading';
    const size = element.size ?? (isHeading ? Math.max(10, 24 - ((element.level - 1) as number) * 2.6) : ctx.defaultSize);

    // 标题与加粗走 bold 字体，斜体走 italic，二者同时命中用 boldItalic
    const bold = isHeading || (('bold' in element ? element.bold : false) ?? false);
    const italic = !isHeading && (('italic' in element ? element.italic : false) ?? false);
    const font = bold && italic
        ? ctx.fonts.boldItalic
        : bold
          ? ctx.fonts.bold
          : italic
            ? ctx.fonts.italic
            : ctx.fonts.normal;

    const lineHeight = (element.lineHeight ?? 1.35) * size;
    const width = element.width ?? ctx.contentWidth;
    const lines = wrapText(element.text ?? '', font, size, width);

    const startTop = topCursor(ctx, element.y ?? 0);
    let y = startTop;

    // 标题保持整块不拆页（Word 标题样式不跨页拆行）
    if (ctx.flow && isHeading) ctx.flow.ensure(lines.length * lineHeight);

    let localY = y;
    for (const line of lines) {
        let x = ctx.originX + (element.x ?? 0);
        if (element.align === 'center' || element.align === 'right') {
            let lineWidth: number;
            try {
                lineWidth = font.widthOfTextAtSize(line, size);
            } catch {
                lineWidth = line.length * size * 0.5;
            }
            const offset = width - lineWidth;
            x += element.align === 'center' ? offset / 2 : Math.max(0, offset);
        }

        if (ctx.flow) {
            ctx.flow.ensure(lineHeight);
            localY = ctx.flow.marginTop + ctx.flow.cursorY;
        }

        currentPage(ctx).drawText(line, {
            x,
            // 顶部原点 -> pdf-lib 左下原点
            y: ctx.pageHeight - localY - size,
            size,
            font,
            color: parseColor(element.color, ctx.defaultColor)
        });
        if (ctx.flow) ctx.flow.cursorY += lineHeight;
        localY += lineHeight;
    }

    return (ctx.flow ? ctx.flow.marginTop + ctx.flow.cursorY : localY) - startTop;
}

/** 取某 run 对应的 pdf-lib 字体（bold / italic 组合） */
function fontForRun(ctx: DrawContext, run: PdfRun, size: number): PDFFont {
    const bold = run.bold ?? false;
    const italic = run.italic ?? false;
    void size;
    return bold && italic ? ctx.fonts.boldItalic : bold ? ctx.fonts.bold : italic ? ctx.fonts.italic : ctx.fonts.normal;
}

/** 测一段 run 文本在给定字号下的像素宽度（含字体回退容错） */
function measureRun(ctx: DrawContext, run: PdfRun, text: string, size: number): number {
    try {
        return fontForRun(ctx, run, size).widthOfTextAtSize(text, size);
    } catch {
        return text.length * size * 0.5;
    }
}

/**
 * 绘制富文本段落（内联样式的多 run 混排）
 *
 * 先把 runs 切成语义「片段」：CJK 字符逐字成段（按字断行），西文按词成段
 * （词内不拆），样式切换/空白也作为片段边界。随后贪心地把片段塞进当前行，
 * 放不下就在该片段前换行，因此加粗/彩色/行内代码可落在不连续的 x 偏移上。
 */
function drawRunsElement(ctx: DrawContext, element: PdfRunsElement): number {
    const size = element.size ?? ctx.defaultSize;
    const lineHeight = (element.lineHeight ?? 1.35) * size;
    const indentLeft = element.indentLeft ?? 0;
    const indentRight = element.indentRight ?? 0;
    const firstIndent = element.indentFirstLine ?? 0;
    const width = element.width ?? ctx.contentWidth - indentLeft - indentRight;
    const baseX = ctx.originX + (element.x ?? 0) + indentLeft;

    /** 渲染片段：一段同一样式的文本（可能是单字、一个西文词、一段空白） */
    interface Piece {
        text: string;
        run: PdfRun;
    }
    const pieces: Piece[] = [];
    const isCjk = (ch: string) => /[　-〿㐀-鿿豈-﫿＀-￯]/.test(ch);
    const isLatin = (ch: string) => /[A-Za-z0-9]/.test(ch);

    for (const run of element.runs) {
        if (!run.text) continue;
        let word = '';
        let wordRun: PdfRun | undefined;
        const flushWord = () => {
            if (word) {
                pieces.push({ text: word, run: wordRun! });
                word = '';
            }
        };
        for (const ch of Array.from(run.text)) {
            if (ch === '\n') {
                flushWord();
                pieces.push({ text: '\n', run });
                continue;
            }
            if (isCjk(ch) || !/[\x00-\x7F]/.test(ch)) {
                // CJK 与任何非 ASCII：逐字成段
                flushWord();
                pieces.push({ text: ch, run });
            } else if (ch === '\t') {
                // 制表符单独成段：绘制时按 tabStops 跳位（无 tabStops 时退化为空格）
                flushWord();
                pieces.push({ text: '\t', run });
            } else if (ch === ' ') {
                flushWord();
                pieces.push({ text: ' ', run });
            } else if (isLatin(ch)) {
                // 西文词：连同标点尽量聚合，但样式切换时先收尾
                if (wordRun && wordRun !== run) flushWord();
                word += ch;
                wordRun = run;
            } else {
                // 其它 ASCII 标点：并入西文词缓冲（避免 "hello," 被拆开）
                if (wordRun && wordRun !== run) flushWord();
                word += ch;
                wordRun = run;
            }
        }
        flushWord();
    }

    /** 一行：片段序列 + 是否首行（首行缩进收窄可用宽） */
    interface Line {
        pieces: Piece[];
        first: boolean;
    }
    const lines: Line[] = [];
    let line: Piece[] = [];
    let lineWidth = 0;
    let lineFirst = true;
    const pushLine = () => {
        lines.push({ pieces: line, first: lineFirst });
        line = [];
        lineWidth = 0;
        lineFirst = false;
    };
    const tabStops = element.tabStops?.length
        ? [...element.tabStops].sort((a, b) => (a.pos ?? Infinity) - (b.pos ?? Infinity))
        : undefined;
    for (const piece of pieces) {
        if (piece.text === '\n') {
            pushLine();
            continue;
        }
        const lineAvail = width - (lineFirst ? firstIndent : 0);
        // tab 的落点在绘制时按制表位确定，宽度不预占（有 tabStops 时按 0 计）
        const w =
            piece.text === '\t' ? (tabStops ? 0 : size * 0.5) : measureRun(ctx, piece.run, piece.text, piece.run.size ?? size);
        if (lineWidth + w > lineAvail && line.length && piece.text !== '\t') {
            pushLine();
        }
        line.push(piece);
        lineWidth += w;
    }
    pushLine();

    const drawLine = (l: Line, localY: number) => {
        const lineTextWidth = l.pieces.reduce(
            (sum, p) => (p.text === '\t' ? sum : sum + measureRun(ctx, p.run, p.text, p.run.size ?? size)),
            0
        );
        let x = baseX + (l.first ? firstIndent : 0);
        if (element.align === 'center' || element.align === 'right') {
            const offset = width - lineTextWidth;
            x += element.align === 'center' ? offset / 2 : Math.max(0, offset);
        }

        // 制表位：预扫每个 tab 到「下一个 tab / 行末」之间的片段总宽（右对齐需要）
        const segWidths = new Map<number, number>();
        if (tabStops) {
            for (let i = 0; i < l.pieces.length; i++) {
                if (l.pieces[i].text !== '\t') continue;
                let sum = 0;
                for (let j = i + 1; j < l.pieces.length && l.pieces[j].text !== '\t'; j++) {
                    sum += measureRun(ctx, l.pieces[j].run, l.pieces[j].text, l.pieces[j].run.size ?? size);
                }
                segWidths.set(i, sum);
            }
        }

        for (let i = 0; i < l.pieces.length; i++) {
            const p = l.pieces[i];
            if (p.text === '\t') {
                if (!tabStops) {
                    x += size * 0.5; // 无制表位定义：退化为半角空格
                    continue;
                }
                // 跳到第一个位于当前位置右侧的制表位（pos 缺省 = 段落右端）
                const stop = tabStops.find((s) => (s.pos ?? width) > x - baseX + 0.5);
                if (!stop) continue;
                const stopPos = stop.pos ?? width;
                const segW = segWidths.get(i) ?? 0;
                const segSize = p.run.size ?? size;
                const segColor = p.run.color ? parseColor(p.run.color, ctx.defaultColor) : ctx.defaultColor;
                let contentX = baseX + stopPos;
                if (stop.align === 'right') contentX -= segW;
                else if (stop.align === 'center') contentX -= segW / 2;
                if (stop.leader) {
                    // 右/居中对齐时前导符只填「内容与制表位之间」的空白
                    drawLeader(ctx, x, contentX, localY, segSize, stop.leader, segColor);
                }
                x = Math.max(x, contentX); // 内容超宽时不倒退
                continue;
            }
            const rs = p.run.size ?? size;
            const w = measureRun(ctx, p.run, p.text, rs);
            const color = p.run.color ? parseColor(p.run.color, ctx.defaultColor) : ctx.defaultColor;
            if (p.run.code || p.run.link) {
                const bg = p.run.code ? parseColor('#F2F2F2') : undefined;
                if (bg) {
                    currentPage(ctx).drawRectangle({
                        x: x - 1,
                        y: ctx.pageHeight - localY - rs,
                        width: w + 2,
                        height: rs + 2,
                        color: bg
                    });
                }
                currentPage(ctx).drawText(p.text, {
                    x,
                    y: ctx.pageHeight - localY - rs,
                    size: rs,
                    font: fontForRun(ctx, p.run, rs),
                    color: p.run.link ? parseColor('#1565C0', ctx.defaultColor) : color
                });
                if (p.run.link) {
                    currentPage(ctx).drawLine({
                        start: { x, y: ctx.pageHeight - localY - 2 },
                        end: { x: x + w, y: ctx.pageHeight - localY - 2 },
                        thickness: 0.5,
                        color: parseColor('#1565C0', ctx.defaultColor)
                    });
                }
            } else {
                currentPage(ctx).drawText(p.text, {
                    x,
                    y: ctx.pageHeight - localY - rs,
                    size: rs,
                    font: fontForRun(ctx, p.run, rs),
                    color
                });
            }
            x += w;
        }
    };

    const startTop = topCursor(ctx, element.y ?? 0);
    let localY = startTop;

    for (let i = 0; i < lines.length; i++) {
        const l = lines[i];
        if (ctx.flow) {
            // widow/orphan：首行不独留页尾（orphan）、末行不独留页首（widow）
            if (i === 0 && lines.length > 1) ctx.flow.ensure(Math.min(2, lines.length) * lineHeight);
            if (i === lines.length - 2 && lines.length > 2) ctx.flow.ensure(2 * lineHeight);
            ctx.flow.ensure(lineHeight);
            localY = ctx.flow.marginTop + ctx.flow.cursorY;
        }
        drawLine(l, localY);
        if (ctx.flow) ctx.flow.cursorY += lineHeight;
        localY += lineHeight;
    }

    const used = (ctx.flow ? ctx.flow.marginTop + ctx.flow.cursorY : localY) - startTop;
    // 段后间距在 flow 模式下由自动分页循环统一进给，不占元素内容高
    return used + (ctx.flow ? 0 : (element.spacingAfter ?? 4));
}

/** 绘制制表位前导符（目录的点线导引等） */
function drawLeader(
    ctx: DrawContext,
    from: number,
    to: number,
    localY: number,
    size: number,
    leader: NonNullable<PdfTabStop['leader']>,
    color: RGB
): void {
    if (to - from < 1) return;
    const page = currentPage(ctx);
    // 前导符位于基线略下方
    const y = ctx.pageHeight - localY - size * 0.28;
    switch (leader) {
        case 'dot':
            for (let px = from + 0.5; px < to; px += 3) {
                page.drawLine({ start: { x: px, y }, end: { x: px + 0.9, y }, thickness: 0.8, color });
            }
            break;
        case 'middleDot':
            for (let px = from + 0.5; px < to; px += 3) {
                page.drawLine({ start: { x: px, y: y + 0.8 }, end: { x: px + 0.9, y: y + 0.8 }, thickness: 1.6, color });
            }
            break;
        case 'hyphen':
            for (let px = from + 0.5; px < to; px += 4) {
                page.drawLine({
                    start: { x: px, y },
                    end: { x: Math.min(to - 0.5, px + 2), y },
                    thickness: 0.8,
                    color
                });
            }
            break;
        case 'underscore':
            page.drawLine({ start: { x: from, y }, end: { x: to, y }, thickness: 0.6, color });
            break;
    }
}

/**
 * 绘制表格，返回占用高度
 *
 * 单元格文本按列宽换行、行高按内容自适应；flow 模式下按行跨页，
 * 跨页后重复表头行（与 Word「标题行重复」行为一致）。
 */
function drawTableElement(ctx: DrawContext, element: PdfTableElement): number {
    const fontSize = element.fontSize ?? 10;
    const padding = element.padding ?? 4;
    const lineH = fontSize * 1.35;
    const hasHeader = element.header !== false && element.rows.length > 0;
    const columnCount = element.rows.reduce((max, row) => Math.max(max, row.length), 0);
    if (!columnCount) return 0;

    const widths = element.columnWidths?.length
        ? element.columnWidths
        : new Array(columnCount).fill(ctx.contentWidth / columnCount);

    const borderColor = element.borderColor === 'none' ? undefined : parseColor(element.borderColor || '#CCCCCC');
    const headerFill = element.headerFill ? parseColor(element.headerFill) : undefined;

    const startTop = topCursor(ctx, element.y ?? 0);
    const startX = ctx.originX + (element.x ?? 0);

    /**
     * 画一行表格。
     * @param y 该行顶部光标（顶部原点）
     * @param rowHeight 行高
     * @param row 单元格数据
     * @param isHeader 是否表头（加粗 + 底纹）
     * @param rowLines 每个单元格已换行的文本
     */
    const drawRow = (y: number, rowHeight: number, row: PdfTableCell[], isHeader: boolean, rowLines: string[][]) => {
        const page = currentPage(ctx);
        if (isHeader && headerFill) {
            page.drawRectangle({
                x: startX,
                y: ctx.pageHeight - y - rowHeight,
                width: widths.reduce((sum, w) => sum + w, 0),
                height: rowHeight,
                color: headerFill
            });
        }

        let x = startX;
        row.forEach((cell, colIndex) => {
            const cellWidth = widths[colIndex] ?? widths[widths.length - 1] ?? ctx.contentWidth / columnCount;

            if (borderColor) {
                page.drawRectangle({
                    x,
                    y: ctx.pageHeight - y - rowHeight,
                    width: cellWidth,
                    height: rowHeight,
                    borderColor,
                    borderWidth: 0.5,
                    color: isHeader ? headerFill : undefined
                });
            }

            const font = isHeader ? ctx.fonts.bold : ctx.fonts.normal;
            const lines = rowLines[colIndex] ?? [];
            // 多行文本在单元格内垂直居中
            const textBlock = lines.length * lineH;
            const offsetY = y + (rowHeight - textBlock) / 2;
            lines.forEach((text, li) => {
                let textWidth: number;
                try {
                    textWidth = font.widthOfTextAtSize(text, fontSize);
                } catch {
                    textWidth = text.length * fontSize * 0.5;
                }
                const drawX = x + Math.max(padding, (cellWidth - textWidth) / 2);
                page.drawText(text, {
                    x: drawX,
                    y: ctx.pageHeight - offsetY - li * lineH - fontSize - (lineH - fontSize) / 2,
                    size: fontSize,
                    font,
                    color: ctx.defaultColor
                });
            });

            x += cellWidth;
        });
    };

    /** 预排版一行：每单元格按各自列宽换行，并求该行高 */
    const layoutRow = (row: PdfTableCell[], isHeader: boolean): { rowHeight: number; rowLines: string[][] } => {
        const font = isHeader ? ctx.fonts.bold : ctx.fonts.normal;
        const rowLines = row.map((cell, colIndex) => {
            const cellWidth = widths[colIndex] ?? widths[widths.length - 1] ?? ctx.contentWidth / columnCount;
            const text = typeof cell === 'string' ? cell : cellToPlainText(cell);
            return wrapText(text ?? '', font, fontSize, Math.max(8, cellWidth - padding * 2));
        });
        const maxLines = rowLines.reduce((m, ls) => Math.max(m, ls.length), 0);
        const rowHeight = element.rowHeight ?? maxLines * lineH + padding * 2;
        return { rowHeight, rowLines };
    };

    let localY = startTop;

    element.rows.forEach((row, rowIndex) => {
        const isHeader = hasHeader && rowIndex === 0;
        const { rowHeight, rowLines } = layoutRow(row, isHeader);

        if (ctx.flow) {
            // 表头与首条数据行保持同页（避免表头孤悬页尾）
            const bind = isHeader && element.rows.length > 1 ? rowHeight + (layoutRow(element.rows[1], false).rowHeight) : rowHeight;
            const moved = ctx.flow.ensure(bind);
            if (moved && hasHeader && rowIndex > 0) {
                // 跨页：重复表头行
                const head = layoutRow(element.rows[0], true);
                drawRow(ctx.flow.marginTop + ctx.flow.cursorY, head.rowHeight, element.rows[0], true, head.rowLines);
                ctx.flow.cursorY += head.rowHeight;
            }
            localY = ctx.flow.marginTop + ctx.flow.cursorY;
        }

        drawRow(localY, rowHeight, row, isHeader, rowLines);

        if (ctx.flow) ctx.flow.cursorY += rowHeight;
        localY += rowHeight;
    });

    return (ctx.flow ? ctx.flow.marginTop + ctx.flow.cursorY : localY) - startTop;
}

/** 表格单元格 -> 纯文本（嵌套元素取其文本拼接） */
function cellToPlainText(cell: PdfElement[]): string {
    return cell
        .map((element) => {
            if (element.type === 'text' || element.type === 'heading') return element.text ?? '';
            return '';
        })
        .join(' ')
        .trim();
}

/** 元素段前间距（pt，自动分页时由循环层进给） */
function spacingBeforeOf(element: PdfElement): number {
    const el = element as { spacingBefore?: number };
    return el.spacingBefore ?? 0;
}

/** 元素段后间距（pt；runs 元素默认 4 与原行为一致） */
function spacingAfterOf(element: PdfElement): number {
    if (element.type === 'runs') return element.spacingAfter ?? 4;
    const el = element as { spacingAfter?: number };
    return el.spacingAfter ?? 0;
}

/** 绘制单个元素，返回占用高度 */
async function drawElement(ctx: DrawContext, element: PdfElement): Promise<number> {
    switch (element.type) {
        case 'text':
        case 'heading':
            return drawTextElement(ctx, element as never);

        case 'line': {
            ctx.page.drawLine({
                start: { x: ctx.originX + element.x1, y: ctx.pageHeight - element.y1 },
                end: { x: ctx.originX + element.x2, y: ctx.pageHeight - element.y2 },
                thickness: element.lineWidth ?? 1,
                color: parseColor(element.color, rgb(0.6, 0.6, 0.6)),
                dashArray: element.dash
            });
            return element.lineWidth ?? 1;
        }

        case 'hr': {
            // 流式分隔线：在当前内容位置画一条横贯内容区的细线
            const thickness = element.lineWidth ?? 0.5;
            const before = element.spacingBefore ?? 4;
            const after = element.spacingAfter ?? 4;
            const total = before + thickness + after;
            if (ctx.flow) ctx.flow.ensure(total);
            const y = ctx.pageHeight - topCursor(ctx) - before;
            currentPage(ctx).drawLine({
                start: { x: ctx.originX, y },
                end: { x: ctx.originX + ctx.contentWidth, y },
                thickness,
                color: parseColor(element.color, rgb(0.6, 0.6, 0.6))
            });
            if (ctx.flow) ctx.flow.cursorY += total;
            return total;
        }

        case 'rect': {
            const y = ctx.originY + element.y;
            const options = {
                x: ctx.originX + element.x,
                y: ctx.pageHeight - y - element.height,
                width: element.width,
                height: element.height,
                color: element.fill ? parseColor(element.fill) : undefined,
                borderColor: element.stroke ? parseColor(element.stroke) : undefined,
                borderWidth: element.stroke ? element.lineWidth ?? 1 : 0
            };
            if (element.radius && element.radius > 0) {
                // pdf-lib 没有圆角矩形 API，用 SVG path 绘制
                const { page, pageHeight, originX, originY } = ctx;
                const px = originX + element.x;
                const py = pageHeight - y - element.height;
                const r = Math.min(element.radius, element.width / 2, element.height / 2);
                const path = [
                    `M ${px + r} ${py}`,
                    `L ${px + element.width - r} ${py}`,
                    `A ${r} ${r} 0 0 1 ${px + element.width} ${py + r}`,
                    `L ${px + element.width} ${py + element.height - r}`,
                    `A ${r} ${r} 0 0 1 ${px + element.width - r} ${py + element.height}`,
                    `L ${px + r} ${py + element.height}`,
                    `A ${r} ${r} 0 0 1 ${px} ${py + element.height - r}`,
                    `L ${px} ${py + r}`,
                    `A ${r} ${r} 0 0 1 ${px + r} ${py}`
                ].join(' ');
                page.drawSvgPath(path, {
                    borderColor: options.borderColor,
                    borderWidth: options.borderWidth,
                    color: options.color
                });
            } else {
                ctx.page.drawRectangle(options);
            }
            return element.height;
        }

        case 'image': {
            try {
                const bytes = await readImageBytes(element.src, ctx.assets);
                const kind = detectImageType(bytes);
                if (!kind) throw new TypeError('仅支持 PNG / JPEG 图片');
                const image: PDFImage = kind === 'png' ? await ctx.doc.embedPng(bytes) : await ctx.doc.embedJpg(bytes);
                const width = element.width ?? image.width;
                const height = element.height ?? image.height;
                if (ctx.flow) ctx.flow.ensure(height);
                currentPage(ctx).drawImage(image, {
                    x: ctx.originX + (element.x ?? 0),
                    y: ctx.pageHeight - topCursor(ctx, element.y ?? 0) - height,
                    width,
                    height
                });
                if (ctx.flow) ctx.flow.cursorY += height;
                return height;
            } catch {
                return 0;
            }
        }

        case 'runs':
            return drawRunsElement(ctx, element);

        case 'table':
            return drawTableElement(ctx, element);

        case 'pageBreak':
            return Number.POSITIVE_INFINITY;

        default:
            return 0;
    }
}

/** 生成 PDF
 * @param spec - 文档规格（页面或元素序列）
 * @param options - 生成选项
 * @returns PDF 字节数据
 */
export async function generatePdf(
    spec: PdfDocumentSpec,
    options: PdfGenerateOptions = {}
): Promise<Uint8Array> {
    try {
        if (!spec || typeof spec !== 'object') throw new TypeError('缺少文档规格');

        const doc = await PDFDocument.create();
        const fonts = await buildFonts(doc, { ...(spec.fonts || {}), ...(options.fonts || {}) });

        // 资源表：显式传入优先，否则沿用规格自带的 assets（往返零搬运）
        const rawAssets = options.assets ?? spec.assets;
        const assets: AssetTable = Array.isArray(rawAssets) ? toAssetTable(rawAssets) : (rawAssets || {});

        // 未内嵌字体时提前拦下中文，避免输出空白方块
        if (!fonts.embedded) {
            for (const element of collectElements(spec)) {
                const cjkText =
                    element.type === 'text' || element.type === 'heading'
                        ? element.text ?? ''
                        : element.type === 'runs'
                          ? element.runs.map((r) => r.text).join('')
                          : element.type === 'table'
                            ? element.rows.flat().map((c) => (typeof c === 'string' ? c : cellToPlainText(c))).join('')
                            : '';
                if (hasCjk(cjkText)) {
                    throw new JdocError(
                        '生成中文 PDF 需要内嵌字体：请通过 fonts.regular 传入中文 TTF/OTF 字节',
                        'pdf'
                    );
                }
            }
        }

        const defaultSize = spec.defaultSize ?? 12;
        const defaultColor = parseColor(spec.defaultColor, rgb(0.12, 0.12, 0.12));

        // 元数据
        if (spec.metadata) {
            doc.setTitle(spec.metadata.title || '');
            doc.setAuthor(spec.metadata.author || '');
            doc.setSubject(spec.metadata.subject || '');
            doc.setKeywords(spec.metadata.keywords ? spec.metadata.keywords.split(/[,;\s]+/).filter(Boolean) : []);
            doc.setProducer(spec.metadata.producer || '@fefeding/jdoc');
            doc.setCreator(spec.metadata.creator || '@fefeding/jdoc');
            if (spec.metadata.created) {
                const date = new Date(spec.metadata.created);
                if (!Number.isNaN(date.getTime())) doc.setCreationDate(date);
            }
            if (spec.metadata.modified) {
                const date = new Date(spec.metadata.modified);
                if (!Number.isNaN(date.getTime())) doc.setModificationDate(date);
            }
        }

        /** 显式分页：每页独立坐标系，元素超出页高不再自动拆页 */
        const renderPage = async (elements: PdfElement[], pageSpec: PdfPageSpecLike & { size?: PdfPageSize; orientation?: 'portrait' | 'landscape' }) => {
            const [pageWidth, pageHeight] = resolvePageSize(pageSpec.size, pageSpec.orientation === 'landscape');
            const margin = resolveMargin(pageSpec.margin);
            const page = doc.addPage([pageWidth, pageHeight]);
            const ctx: DrawContext = {
                doc,
                fonts,
                assets,
                page,
                pageHeight,
                contentWidth: pageWidth - margin.left - margin.right,
                originX: margin.left,
                originY: margin.top,
                defaultSize,
                defaultColor
            };

            let cursorY = 0;
            for (const element of elements) {
                if (element.type === 'pageBreak') {
                    cursorY = Number.POSITIVE_INFINITY;
                    break;
                }
                const used = await drawElement({ ...ctx, originY: margin.top + cursorY }, element);
                cursorY += used;
            }
        };

        if (spec.pages?.length) {
            for (const pageSpec of spec.pages) {
                await renderPage(pageSpec.elements || spec.elements || [], pageSpec);
            }
        } else {
            // 自动分页：行级流式排版，内容逐行溢出即另起一页
            const elements = spec.elements || [];
            const [pageWidth, pageHeight] = resolvePageSize(spec.pageSize, false);
            const margin = resolveMargin(spec.margin);
            const usableHeight = pageHeight - margin.top - margin.bottom;

            let page = doc.addPage([pageWidth, pageHeight]);
            let pageNumber = 1; // 从 1 开始，供 onElementLayout 回调使用
            let ctx: DrawContext = {
                doc,
                fonts,
                assets,
                page,
                pageHeight,
                contentWidth: pageWidth - margin.left - margin.right,
                originX: margin.left,
                originY: margin.top,
                defaultSize,
                defaultColor
            };

            const newPage = () => {
                page = doc.addPage([pageWidth, pageHeight]);
                ctx = { ...ctx, page };
                flow.page = page;
                flow.cursorY = 0;
                pageNumber += 1;
            };
            const flow: FlowState = {
                page,
                marginTop: margin.top,
                cursorY: 0,
                usableHeight,
                ensure(need: number) {
                    // 页首不换页（元素高于整页时避免死循环）
                    if (this.cursorY + need <= this.usableHeight || this.cursorY === 0) return false;
                    newPage();
                    return true;
                }
            };

            // 段间距：Word 规则为「前段 spacingAfter + 本段 spacingBefore」相加
            let prevAfter = 0;
            for (const element of elements) {
                if (element.type === 'pageBreak') {
                    newPage();
                    prevAfter = 0;
                    continue;
                }

                const before = spacingBeforeOf(element);
                const after = spacingAfterOf(element);
                const gap = prevAfter + before;
                if (gap > 0) {
                    flow.cursorY += Math.min(gap, Math.max(0, usableHeight - flow.cursorY));
                }

                spec.onElementLayout?.(element, pageNumber);
                await drawElement({ ...ctx, flow }, element);
                prevAfter = after;
            }
        }

        const bytes = await doc.save({ useObjectStreams: options.compress !== false });
        return bytes;
    } catch (error) {
        throw wrapError('pdf', '生成', error);
    }
}

/** 收集文档中所有元素（供前置校验使用） */
function collectElements(spec: PdfDocumentSpec): PdfElement[] {
    if (spec.pages?.length) return spec.pages.flatMap((page) => page.elements || []);
    return spec.elements || [];
}

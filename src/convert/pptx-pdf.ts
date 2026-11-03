/**
 * PPTX -> PDF（幻灯片版式直排）
 *
 * 幻灯片是「固定画布 + 绝对定位元素」，与 docx/md 的流式文档不同。经 blocks
 * 中转的流式排版会把 16:9 幻灯片重排成 A4 竖版：绝对定位的文本框互相叠压、
 * 大图溢出成多页、页尾大面积留白。这里按幻灯片原稿直排：
 *
 *   - 每张幻灯片一页，页面尺寸 = 幻灯片尺寸（解析器坐标为 96dpi px，转 pt ×0.75）
 *   - 文本框 / 图片 / 表格 / 形状按原始位置与尺寸绝对定位（依赖 pages 模式
 *     对显式 y 元素的绝对定位语义，见 generate.ts renderPage）
 *   - 渐变填充取首停色近似，旋转 / 翻转 / 自定义几何不还原（PDF 生成端无对应能力）
 *
 * @module convert/pptx-pdf
 */

import type { DocInput } from '../core/input';
import { pptxToStandard } from '../pptx/index';
import { jsonToPdf } from '../pdf/index';
import { detectSystemFonts } from '../pdf/fonts';
import type {
    PdfElement,
    PdfFontSet,
    PdfPageSpec,
    PdfRun,
    PdfRunsElement,
    PdfTableCell,
    PdfTableCellObject
} from '../pdf/types';
import type {
    PptxDocument,
    PptxElement,
    PptxFill,
    PptxGroupElement,
    PptxImageElement,
    PptxShapeElement,
    PptxSlide,
    PptxTableElement,
    PptxTableCell,
    PptxTextElement,
    PptxTextRun
} from '@fefeding/ppt-parser';

/** 解析器导出的坐标单位是 px（96dpi，16:9 幻灯片为 1280×720），PDF 用 pt */
const PX_TO_PT = 0.75;

/** 幻灯片上未显式给字号的文本默认 18pt（与 PowerPoint / WPS 默认一致） */
const DEFAULT_FONT_SIZE = 18;

/** pptxToPdf 选项 */
export interface PptxToPdfOptions {
    /** 内嵌字体（中文必须）；未提供且内容含中文时自动探测系统字体 */
    fonts?: PdfFontSet;
    /** 是否自动探测系统字体（默认 true，仅 Node 端有效） */
    autoFont?: boolean;
}

/** 颜色归一：解析器的颜色可能不带 # 前缀 */
function normalizeColor(value: string | undefined | null): string | undefined {
    if (!value || value === 'none') return undefined;
    return value.startsWith('#') ? value : `#${value}`;
}

/** #RRGGBB -> [r, g, b]（非法输入返回 null） */
function rgbOf(color: string | undefined): [number, number, number] | null {
    const hex = normalizeColor(color)?.slice(1);
    if (!hex || !/^[0-9a-fA-F]{6}$/.test(hex)) return null;
    return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
}

/** [r, g, b] -> #RRGGBB */
function toHex(rgb: [number, number, number]): string {
    return `#${rgb.map((v) => Math.round(v).toString(16).padStart(2, '0')).join('')}`;
}

/** 前景按 alpha 混合到背景（模拟透明度：PDF 生成端只支持不透明填充） */
function blendOver(fg: string, alpha: number, bg: string): string {
    const f = rgbOf(fg);
    const b = rgbOf(bg);
    if (!f || !b) return fg;
    if (alpha >= 1) return fg;
    return toHex([0, 1, 2].map((i) => f[i] * alpha + b[i] * (1 - alpha)) as [number, number, number]);
}

/**
 * 填充 -> 不透明单色（相对幻灯片背景混合透明度）
 *
 * 渐变取「最不透明」的 stop 近似：装饰形状常用「全透明白 + 半透明主题色」
 * 的渐变，取首停色会画出突兀的实心白块。整体透明度 >= 97% 的形状肉眼不可见，
 * 直接跳过不画。
 */
function fillToColor(fill: PptxFill | undefined, slideBg: string): string | undefined {
    if (!fill || fill === 'none') return undefined;
    if (typeof fill === 'string') return blendOver(normalizeColor(fill)!, 1, slideBg);
    const f = fill as {
        type?: string;
        color?: string;
        transparency?: number;
        stops?: { color: string; transparency?: number }[];
        fg?: string;
    };
    if (f.type === 'gradient' && f.stops?.length) {
        let best = -1;
        let bestAlpha = 0;
        for (const stop of f.stops) {
            const alpha = 1 - (stop.transparency ?? 0) / 100;
            if (alpha > bestAlpha) {
                bestAlpha = alpha;
                best = f.stops.indexOf(stop);
            }
        }
        if (best < 0 || bestAlpha < 0.03) return undefined;
        return blendOver(normalizeColor(f.stops[best].color)!, bestAlpha, slideBg);
    }
    if (f.type === 'pattern') return blendOver(normalizeColor(f.fg)!, 1, slideBg);
    if (f.type === 'image') return undefined;
    const alpha = 1 - (f.transparency ?? 0) / 100;
    const color = normalizeColor(f.color);
    if (!color || alpha < 0.03) return undefined;
    return blendOver(color, alpha, slideBg);
}

/** 字符宽度估算系数（CJK 全宽 1.0，西文数字 0.6，其余标点 0.5） */
function charWidthFactor(ch: string): number {
    if (/[　-〿㐀-鿿豈-﫿＀-￯]/.test(ch)) return 1;
    if (/[A-Za-z0-9@%]/.test(ch)) return 0.6;
    return 0.5;
}

/** 估算一段文本的渲染宽度（pt） */
function estimateWidth(text: string, size: number): number {
    let w = 0;
    for (const ch of Array.from(text)) w += charWidthFactor(ch) * size;
    return w;
}

/**
 * 单行文本框的溢出缩字。
 *
 * WPS/PowerPoint 对「高度只容一行」的标题框会在排不下时自动缩小字号
 * （normAutofit 很多时候不写进文件，由排版器实时处理）；PDF 端直接换行的
 * 表现是孤字换行或截断，观感很差。这里在框高只容一行且估算行宽超框时，
 * 按比例整体缩小字号（下限 0.4）。多行文本框不动——换行本就是原稿排版。
 */
function shrinkSingleLineBox(el: PptxTextElement, lines: PdfRun[][]): void {
    const boxW = el.width * PX_TO_PT;
    const boxH = el.height * PX_TO_PT;
    const sizes = lines.flat().map((r) => r.size ?? 0);
    const maxSize = Math.max(...sizes, 0);
    if (!maxSize || boxW <= 0) return;
    // 行高 1.2：框高容不下两行才视为单行框
    if (boxH >= maxSize * 1.2 * 1.8) return;
    const maxLineW = Math.max(
        ...lines.map((line) => line.reduce((sum, r) => sum + estimateWidth(r.text, r.size ?? maxSize), 0))
    );
    if (maxLineW <= boxW) return;
    // 0.96 安全边距：估算字宽与真实字体度量有偏差，压线仍可能换行出孤字
    const k = Math.max(0.4, (boxW * 0.96) / maxLineW);
    for (const line of lines) {
        for (const r of line) {
            if (r.size) r.size = Math.round(r.size * k * 10) / 10;
        }
    }
}

/** PPTX 文本片段 -> PDF run（解析器字号已是 pt，不需要坐标缩放） */
function runToPdfRun(run: PptxTextRun, fallbackSize?: number, fontScale = 1): PdfRun {
    const r: PdfRun = { text: run.text ?? '' };
    if (run.bold) r.bold = true;
    if (run.italic) r.italic = true;
    if (run.strike) r.strike = true;
    if (run.color) r.color = normalizeColor(run.color);
    const size = (run.fontSize ?? fallbackSize) * fontScale;
    if (size) r.size = size;
    if (run.href) r.link = run.href;
    return r;
}

/** 文本框 -> runs 元素：多段落用 \n 合并为一个绝对定位段落 */
function textToElement(el: PptxTextElement): PdfElement | null {
    // 「溢出则缩字」（normAutofit fontScale）：标题框排不下整行时 WPS/PPT 会
    // 自动缩小字号，不套用这个比例会在 PDF 里换行溢出文本框
    const fontScale = (el.fontScale ?? 100) / 100;
    const paragraphs = el.paragraphs?.length
        ? el.paragraphs
        : el.runs?.length || el.text
          ? [{ runs: el.runs, text: el.text }]
          : [];
    const lines: PdfRun[][] = [];
    let align: PdfRunsElement['align'];
    for (const p of paragraphs) {
        const runs = p.runs?.length
            ? p.runs.map((r) => runToPdfRun(r, el.fontSize, fontScale))
            : p.text
              ? [{ text: p.text, size: (el.fontSize ?? DEFAULT_FONT_SIZE) * fontScale }]
              : [];
        if (!runs.some((r) => r.text?.trim())) {
            // 空段落保留为空行（与幻灯片上的留行一致）
            lines.push([{ text: '' }]);
            continue;
        }
        lines.push(runs);
        align = align ?? (p.align as PdfRunsElement['align']);
    }
    if (!lines.length) return null;
    shrinkSingleLineBox(el, lines);
    const runs: PdfRun[] = [];
    lines.forEach((line, i) => {
        if (i > 0) runs.push({ text: '\n' });
        runs.push(...line);
    });
    return {
        type: 'runs',
        x: el.x * PX_TO_PT,
        y: el.y * PX_TO_PT,
        width: el.width * PX_TO_PT,
        runs,
        align,
        // 幻灯片排版行距比文档紧，取 1.2 近似
        lineHeight: 1.2,
        spacingAfter: 0
    };
}

/** 形状 -> 矩形（旋转 / 翻转 / 自定义几何不还原；透明填充已折算到背景色） */
function shapeToElement(el: PptxShapeElement, slideBg: string): PdfElement | null {
    if (!el.width && !el.height) return null;
    const fill = fillToColor(el.fill, slideBg);
    const line = el.line && el.line !== 'none' ? el.line : undefined;
    const stroke = normalizeColor(line?.color);
    if (!fill && !stroke) return null;
    return {
        type: 'rect',
        x: el.x * PX_TO_PT,
        y: el.y * PX_TO_PT,
        width: el.width * PX_TO_PT,
        height: el.height * PX_TO_PT,
        fill,
        stroke,
        lineWidth: line?.width ? Math.max(0.5, line.width) : undefined
    };
}

/** 图片 -> 绝对定位图片元素 */
function imageToElement(el: PptxImageElement): PdfElement | null {
    const src = el.data || el.src;
    if (!src || typeof src !== 'string') return null;
    return {
        type: 'image',
        src,
        x: el.x * PX_TO_PT,
        y: el.y * PX_TO_PT,
        width: el.width * PX_TO_PT,
        height: el.height * PX_TO_PT
    };
}

/** 单元格 -> PDF 单元格（保留合并信息，否则合并表头会错位） */
function cellToPdfCell(cell: PptxTableCell): PdfTableCell {
    const text =
        cell.text ??
        cell.paragraphs
            ?.map((p) => p.text ?? p.runs?.map((r) => r.text ?? '').join('') ?? '')
            .join('\n') ??
        '';
    if (!cell.colSpan && !cell.rowSpan && !cell.fill) return text;
    const obj: PdfTableCellObject = { content: text };
    if (cell.colSpan && cell.colSpan > 1) obj.colspan = cell.colSpan;
    if (cell.rowSpan && cell.rowSpan > 1) obj.rowspan = cell.rowSpan;
    const fill = normalizeColor(cell.fill);
    if (fill) obj.fill = fill;
    return obj;
}

/** 表格字号：取首个显式给出的 run 字号，缺省 14pt */
function tableFontSize(el: PptxTableElement): number {
    for (const row of el.rows) {
        for (const cell of row.cells) {
            for (const p of cell.paragraphs || []) {
                for (const r of p.runs || []) {
                    if (r.fontSize) return r.fontSize;
                }
            }
        }
    }
    return 14;
}

/** 表格 -> 绝对定位表格元素 */
function tableToElement(el: PptxTableElement): PdfElement {
    return {
        type: 'table',
        x: el.x * PX_TO_PT,
        y: el.y * PX_TO_PT,
        rows: el.rows.map((row) =>
            row.cells.flatMap((cell) => (cell.hMerge || cell.vMerge ? [] : [cellToPdfCell(cell)]))
        ),
        header: false,
        fontSize: tableFontSize(el),
        columnWidths: el.colWidths?.map((w) => w * PX_TO_PT)
    };
}

/**
 * 递归收集元素（展平组合）
 *
 * 组合子元素坐标语义与编辑器渲染端一致：`relative` 表示相对组合原点（需加上
 * 组合位置），其余（page / 绝对坐标）已是页面绝对坐标。
 */
function collectElements(
    list: PptxElement[],
    out: PdfElement[],
    slideBg: string,
    offset?: { x: number; y: number }
): void {
    const dx = offset?.x ?? 0;
    const dy = offset?.y ?? 0;
    for (const el of list) {
        switch (el.type) {
            case 'text': {
                const t = el as PptxTextElement;
                const e = textToElement({ ...t, x: t.x + dx, y: t.y + dy });
                if (e) out.push(e);
                break;
            }
            case 'shape': {
                const s = el as PptxShapeElement;
                const e = shapeToElement({ ...s, x: s.x + dx, y: s.y + dy }, slideBg);
                if (e) out.push(e);
                break;
            }
            case 'image': {
                const img = el as PptxImageElement;
                const e = imageToElement({ ...img, x: img.x + dx, y: img.y + dy });
                if (e) out.push(e);
                break;
            }
            case 'table': {
                const t = el as PptxTableElement;
                out.push(tableToElement({ ...t, x: t.x + dx, y: t.y + dy }));
                break;
            }
            case 'group': {
                const g = el as PptxGroupElement;
                if (g.childrenCoordinates === 'relative') {
                    collectElements(g.children || [], out, slideBg, { x: (g.x || 0) + dx, y: (g.y || 0) + dy });
                } else {
                    collectElements(g.children || [], out, slideBg);
                }
                break;
            }
            default:
                // chart / diagram / video 等：PDF 直排暂不绘制（图表数据在
                // pptx -> md/html/docx 路径仍有呈现）
                break;
        }
    }
}

/** 幻灯片背景色（纯色 / 渐变取首停色），无背景返回白色兜底 */
function slideBackgroundColor(slide: PptxSlide): string {
    const bg = slide.background;
    if (typeof bg === 'string') return normalizeColor(bg) ?? '#ffffff';
    if (bg) {
        const b = bg as { type?: string; color?: string; stops?: { color: string }[] };
        const color =
            b.type === 'gradient' ? normalizeColor(b.stops?.[0]?.color) : normalizeColor(b.color);
        if (color) return color;
    }
    return '#ffffff';
}

/** 幻灯片是否含中文（决定是否探测系统字体） */
function slidesHaveCjk(slides: PptxSlide[]): boolean {
    let found = false;
    const walk = (els?: PptxElement[]) => {
        for (const el of els || []) {
            if (found) return;
            if (el.type === 'text') {
                const t = el as PptxTextElement;
                const text =
                    t.text ||
                    (t.paragraphs || []).map((p) => p.text || p.runs?.map((r) => r.text ?? '').join('') || '').join('');
                if (/[一-鿿]/.test(text)) found = true;
            } else if (el.type === 'group') {
                walk((el as PptxGroupElement).children);
            }
        }
    };
    for (const s of slides) walk(s.elements);
    return found;
}

/**
 * 把 PPTX 按幻灯片版式直排为 PDF（每页 = 一张幻灯片）
 *
 * @param input - PPTX 字节 / base64 / data URL
 * @param options - 字体选项
 * @returns PDF 字节数据
 */
export async function pptxToPdf(input: DocInput, options: PptxToPdfOptions = {}): Promise<Uint8Array> {
    const doc = (await pptxToStandard(input, { assets: 'both' })) as PptxDocument;
    const pageW = doc.slideSize.width * PX_TO_PT;
    const pageH = doc.slideSize.height * PX_TO_PT;

    // 自动字体：含中文且未提供字体时探测系统字体（仅 Node 端）
    let fonts = options.fonts;
    if (!fonts && options.autoFont !== false && slidesHaveCjk(doc.slides)) {
        try {
            fonts = await detectSystemFonts();
        } catch {
            // 浏览器端无系统字体：交给 jsonToPdf 抛友好的中文提示
        }
    }

    const pages: PdfPageSpec[] = [];
    for (const slide of doc.slides) {
        if (slide.hidden) continue;
        const bg = slideBackgroundColor(slide);
        const elements: PdfElement[] = [{ type: 'rect', x: 0, y: 0, width: pageW, height: pageH, fill: bg }];
        collectElements(slide.elements || [], elements, bg);
        pages.push({ size: [pageW, pageH], margin: 0, elements });
    }
    if (!pages.length) pages.push({ size: [pageW, pageH], margin: 0, elements: [] });

    return jsonToPdf({ pages, fonts, defaultSize: DEFAULT_FONT_SIZE }, {});
}

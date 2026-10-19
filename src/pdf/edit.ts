/**
 * PDF 叠加编辑
 *
 * 在**既有 PDF** 上追加内容并保存，与 `generate.ts` 的「从零生成」是两件事。
 * 实现路径是 pdf-lib 的 `PDFDocument.load()`：原文件的对象结构、页面内容流、
 * 矢量图形、嵌入字体全部原样保留，只往现有页的内容流末尾追加绘制指令。
 *
 * 这条路不能走「解析成模型再重新生成」——PDF 里大量信息（矢量路径、精确字距、
 * 字体子集、表单域、注释）无法从文本层还原，重排等于毁掉原文档。
 *
 * 坐标系与 `generate.ts` 保持一致：**顶部原点、单位 pt**，内部换算为 pdf-lib 的左下原点。
 *
 * @module pdf/edit
 */

import { PDFDocument, degrees, rgb, type Color, type PDFImage, type PDFPage, type RGB } from 'pdf-lib';
import { JdocError, wrapError } from '../core/errors';
import { toUint8Array, type DocInput } from '../core/input';
import type { PdfFontSet, PdfMetadata } from './types';
import { buildFonts, wrapText, type FontBundle } from './generate';

/** 叠加对象的公共几何：x/y 为左上角，单位 pt */
interface OverlayBase {
    /** 左边界（pt，相对页面左边） */
    x: number;
    /** 上边界（pt，相对页面顶边） */
    y: number;
    /** 不透明度 0~1，缺省 1 */
    opacity?: number;
}

/** 图片 / 签名 */
export interface PdfOverlayImage extends OverlayBase {
    type: 'image';
    /** 图片字节（PNG / JPEG） */
    data: Uint8Array | ArrayBuffer;
    /** 缺省按 PNG 处理 */
    format?: 'png' | 'jpg';
    width: number;
    height: number;
    /** 顺时针旋转角度，缺省 0 */
    rotate?: 0 | 90 | 180 | 270;
}

/** 文本块 */
export interface PdfOverlayText extends OverlayBase {
    type: 'text';
    /** 文本内容；给了 width 就自动换行，也可用 \n 强制断行 */
    text: string;
    /** 文本框宽（pt），用于换行；缺省不换行 */
    width?: number;
    /** 字号（pt），缺省 12 */
    size?: number;
    /** 字体族，缺省 normal */
    fontFamily?: 'normal' | 'bold' | 'italic' | 'boldItalic';
    /** 文字颜色，缺省黑色 */
    color?: string;
    /** 行距倍数，缺省 1.35 */
    lineHeight?: number;
    /** 对齐，缺省 left */
    align?: 'left' | 'center' | 'right';
}

/** 手绘 / SVG 路径（签名、圈画、箭头） */
export interface PdfOverlayPath extends OverlayBase {
    type: 'path';
    /** SVG path 的 d 属性，绝对坐标、原点在左上 */
    d: string;
    /** 缩放，缺省 1 */
    scale?: number;
    /** 描边色，缺省红色（签名场景惯例） */
    stroke?: string;
    /** 描边宽度，缺省 2 */
    strokeWidth?: number;
}

/** 矩形（高亮块、遮挡、边框） */
export interface PdfOverlayRect extends OverlayBase {
    type: 'rect';
    width: number;
    height: number;
    /** 填充色，缺省透明 */
    fill?: string;
    /** 描边色，缺省透明 */
    stroke?: string;
    /** 描边宽度（pt），缺省 1 */
    strokeWidth?: number;
}

/** 叠加对象 */
export type PdfOverlay = PdfOverlayImage | PdfOverlayText | PdfOverlayPath | PdfOverlayRect;

/** 页面旋转 */
export interface PdfRotateOp {
    /** 页序（0 起） */
    page: number;
    /** 顺时针角度 */
    degrees: 90 | 180 | 270;
}

/** 编辑规格 */
export interface PdfEditSpec {
    /**
     * 按页分组的叠加对象，数组下标即页序（0 起）
     *
     * 同一页内按数组顺序叠放，靠后的盖在靠前的上面。
     */
    overlays?: PdfOverlay[][];
    /** 删除的页序（0 起） */
    removePages?: number[];
    /** 旋转页面（改的是页面本身的 /Rotate，不影响叠加物坐标） */
    rotate?: PdfRotateOp[];
    /**
     * 重排页序：`[2, 0, 1]` 表示新第 1 页用原第 3 页
     *
     * 与 removePages 都作用于最终页集合：先删后重排。
     */
    pageOrder?: number[];
    /** 内嵌字体；中文或自定义字形时必需 */
    fonts?: PdfFontSet;
    /** 文档元数据覆盖 */
    metadata?: Partial<PdfMetadata>;
}

/** 颜色字符串 -> pdf-lib Color */
function toColor(value: string | undefined, fallback?: RGB): Color | undefined {
    if (!value) return fallback;
    const hex = value.trim().replace(/^#/, '');
    if (hex.length === 3) {
        const [r, g, b] = hex.split('').map((c) => parseInt(c + c, 16) / 255);
        return rgb(r, g, b);
    }
    if (hex.length >= 6) {
        return rgb(parseInt(hex.slice(0, 2), 16) / 255, parseInt(hex.slice(2, 4), 16) / 255, parseInt(hex.slice(4, 6), 16) / 255);
    }
    return fallback;
}

/**
 * 搬运文档元数据到另一份文档
 *
 * pdf-lib 的 setter 不接受 undefined（元数据缺失时 getXxx() 返回 undefined，
 * 直接 set 会抛 `title must be of type string`），因此逐项跳过空值。
 */
function carryMetadata(from: PDFDocument, to: PDFDocument): void {
    const title = from.getTitle();
    const author = from.getAuthor();
    const subject = from.getSubject();
    const keywords = from.getKeywords();
    const creator = from.getCreator();
    const producer = from.getProducer();
    const created = from.getCreationDate();
    const modified = from.getModificationDate();
    if (title !== undefined) to.setTitle(title);
    if (author !== undefined) to.setAuthor(author);
    if (subject !== undefined) to.setSubject(subject);
    if (keywords !== undefined) to.setKeywords([keywords]);
    if (creator !== undefined) to.setCreator(creator);
    if (producer !== undefined) to.setProducer(producer);
    if (created !== undefined) to.setCreationDate(created);
    if (modified !== undefined) to.setModificationDate(modified);
}

/** 顶部原点 -> pdf-lib 左下原点 */
function flipY(pageHeight: number, y: number, boxHeight: number): number {
    return pageHeight - y - boxHeight;
}

/**
 * SVG path 的粗略包围盒
 *
 * pdf-lib 的 drawSvgPath 以左下为原点且不接受负坐标，必须知道路径自身的偏移量
 * 才能摆到目标位置。这里假定传入的是绝对坐标路径（手绘笔迹天然如此），扫描所有
 * 坐标对取极值；相对命令（r* / 相对 c*）的数值是增量，直接当绝对会算偏，
 * 因此调用方应先把相对路径转成绝对。
 */
function pathBBox(d: string): { minX: number; minY: number; width: number; height: number } | null {
    let minX = Infinity;
    let minY = Infinity;
    let maxX = -Infinity;
    let maxY = -Infinity;
    const re = /([MmLlHhVvCcSsQqTtAaZz])([^MmLlHhVvCcSsQqTtAaZz]*)/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(d))) {
        const args = (m[2].match(/-?\d*\.?\d+(?:e[-+]?\d+)?/gi) || []).map(Number);
        for (let i = 0; i + 1 < args.length; i += 2) {
            const x = args[i];
            const y = args[i + 1];
            if (!Number.isFinite(x) || !Number.isFinite(y)) continue;
            if (x < minX) minX = x;
            if (x > maxX) maxX = x;
            if (y < minY) minY = y;
            if (y > maxY) maxY = y;
        }
    }
    if (!Number.isFinite(minX) || !Number.isFinite(minY)) return null;
    return { minX, minY, width: maxX - minX, height: maxY - minY };
}

/** 绘制叠加图片 */
async function drawImage(doc: PDFDocument, page: PDFPage, item: PdfOverlayImage, pageHeight: number): Promise<void> {
    const bytes = await toUint8Array(item.data);
    const image: PDFImage = item.format === 'jpg' ? await doc.embedJpg(bytes) : await doc.embedPng(bytes);
    const opacity = item.opacity ?? 1;
    const rotate = item.rotate ?? 0;
    const x = item.x;
    const y = flipY(pageHeight, item.y, item.height);

    // 90/270 时视觉包围盒宽高对调：pdf-lib 绕图片中心旋转，
    // 因此要把中心还原成「未旋转尺寸下的中心」再平移回去
    if (rotate === 90 || rotate === 270) {
        const visualW = item.height;
        const visualH = item.width;
        const ccx = x + visualW / 2;
        const ccy = y + visualH / 2;
        page.drawImage(image, {
            x: ccx - item.width / 2,
            y: ccy - item.height / 2,
            width: item.width,
            height: item.height,
            rotate: degrees(rotate),
            opacity
        });
        return;
    }
    page.drawImage(image, { x, y, width: item.width, height: item.height, rotate: degrees(rotate), opacity });
}

/** 绘制叠加文本块 */
function drawText(fonts: FontBundle, page: PDFPage, item: PdfOverlayText, pageHeight: number): void {
    const size = item.size ?? 12;
    const font = fonts[item.fontFamily ?? 'normal'] ?? fonts.normal;
    const raw = String(item.text ?? '');
    const lines = item.width ? wrapText(raw, font, size, item.width) : raw.split('\n');
    const lineHeight = (item.lineHeight ?? 1.35) * size;
    const boxHeight = lineHeight * Math.max(1, lines.length);
    const top = flipY(pageHeight, item.y, boxHeight);
    const color = toColor(item.color, rgb(0, 0, 0));
    const align = item.align ?? 'left';

    lines.forEach((line, i) => {
        if (!line) return;
        let x = item.x;
        if (item.width) {
            let w = 0;
            try {
                w = font.widthOfTextAtSize(line, size);
            } catch {
                w = line.length * size * 0.5;
            }
            if (align === 'center') x = item.x + (item.width - w) / 2;
            else if (align === 'right') x = item.x + item.width - w;
        }
        // pdf-lib 的 drawText 以基线起算，逐行下移一个行高
        page.drawText(line, { x, y: top + boxHeight - lineHeight * (i + 1) + size * 0.22, size, font, color });
    });
}

/** 绘制手绘路径 */
function drawPath(page: PDFPage, item: PdfOverlayPath, pageHeight: number): void {
    const bbox = pathBBox(item.d);
    if (!bbox) return;
    const scale = item.scale ?? 1;
    const y = flipY(pageHeight, item.y, bbox.height * scale);
    page.drawSvgPath(item.d, {
        x: item.x - bbox.minX * scale,
        y: y + bbox.minY * scale,
        scale,
        borderWidth: item.strokeWidth ?? 2,
        borderColor: toColor(item.stroke, rgb(0.8, 0.1, 0.2)),
        borderLineCap: 1
    });
}

/** 绘制矩形 */
function drawRect(page: PDFPage, item: PdfOverlayRect, pageHeight: number): void {
    page.drawRectangle({
        x: item.x,
        y: flipY(pageHeight, item.y, item.height),
        width: item.width,
        height: item.height,
        color: toColor(item.fill),
        borderColor: toColor(item.stroke),
        borderWidth: item.stroke ? (item.strokeWidth ?? 1) : undefined,
        opacity: item.opacity,
        borderOpacity: item.opacity
    });
}

/**
 * 在既有 PDF 上叠加内容并保存
 *
 * @param input - 原 PDF 字节
 * @param spec - 编辑规格（叠加对象与页面操作）
 * @returns 编辑后的 PDF 字节
 * @remarks
 * 中文/自定义字形需通过 `spec.fonts` 传入 TTF/OTF 字节（依赖 `@pdf-lib/fontkit`），
 * 与 `jsonToPdf` 同一约束：pdf-lib 的标准 14 字体不含 CJK 字形。
 *
 * @example
 * ```js
 * const bytes = await editPdf(file, {
 *   overlays: [[{ type: 'rect', x: 40, y: 60, width: 200, height: 18, fill: '#ffe08a', opacity: 0.5 }]],
 *   rotate: [{ page: 0, degrees: 90 }]
 * });
 * ```
 */
export async function editPdf(input: DocInput, spec: PdfEditSpec): Promise<Uint8Array> {
    try {
        const source = await PDFDocument.load(await toUint8Array(input), {
            ignoreEncryption: true,
            // 保留原文件元数据；spec.metadata 显式覆盖时在后面单独设置
            updateMetadata: false
        });

        // ---- 1. 页集合调整（删页 / 重排）----
        // 放到叠加之前：spec.overlays 的下标语义是「最终页序」，用户看到的页序。
        // 涉及页集合变化时用中转文档承载新顺序——pdf-lib 无法把页「搬回」同一文档
        // （copyPages 跨文档只注册 ref，而 ref 早已在原 context 登记），
        // 逐个 removePage/addPage 又会让页树与页数组失配，save() 时断言失败。
        const original = [...source.getPages()];
        let removed = new Set<number>(spec.removePages ?? []);
        let order: number[] | null = null;
        if (spec.pageOrder?.length) {
            const valid = spec.pageOrder.filter((i) => i >= 0 && i < original.length);
            const identity = valid.length === original.length && valid.every((v, i) => v === i);
            if (valid.length && !identity) order = valid;
        }
        const needsRestage = removed.size > 0 || !!order;
        let doc = source;
        if (needsRestage) {
            const kept = original.map((_, i) => i).filter((i) => !removed.has(i));
            // order 里的值本身就是原文档页序；没有 order 时就是「删掉指定页后的自然序」
            const pick = (order ?? kept).filter((i) => i >= 0 && i < original.length);
            if (!pick.length) throw new JdocError('删除后文档已无页面', 'pdf');
            const staging = await PDFDocument.create();
            const copied = await staging.copyPages(source, pick);
            copied.forEach((p) => staging.addPage(p));
            // 元数据随文档走，否则重排/删页后标题作者全丢
            carryMetadata(source, staging);
            doc = staging;
            removed = new Set<number>();
        }

        // ---- 2. 页面旋转（改的是页面 /Rotate，与叠加物坐标无关）----
        for (const op of spec.rotate ?? []) {
            const page = doc.getPages()[op.page];
            if (!page) continue;
            page.setRotation(degrees((page.getRotation().angle + op.degrees) % 360));
        }

        // ---- 3. 叠加 ----
        const fonts = await buildFonts(doc, spec.fonts);
        const pages = doc.getPages();
        const overlays = spec.overlays ?? [];
        for (let i = 0; i < pages.length; i++) {
            const list = overlays[i];
            if (!list?.length) continue;
            const page = pages[i];
            const pageHeight = page.getHeight();
            for (const item of list) {
                if (item.type === 'image') await drawImage(doc, page, item, pageHeight);
                else if (item.type === 'text') drawText(fonts, page, item, pageHeight);
                else if (item.type === 'path') drawPath(page, item, pageHeight);
                else drawRect(page, item, pageHeight);
            }
        }

        // ---- 4. 元数据覆盖 ----
        const meta = spec.metadata;
        if (meta) {
            if (meta.title !== undefined) doc.setTitle(meta.title);
            if (meta.author !== undefined) doc.setAuthor(meta.author);
            if (meta.subject !== undefined) doc.setSubject(meta.subject);
            if (meta.keywords !== undefined) doc.setKeywords(Array.isArray(meta.keywords) ? meta.keywords : [meta.keywords]);
            if (meta.creator !== undefined) doc.setCreator(meta.creator);
            if (meta.producer !== undefined) doc.setProducer(meta.producer);
        }

        if (!doc.getPageCount()) throw new JdocError('删除后文档已无页面', 'pdf');
        return await doc.save();
    } catch (err) {
        throw wrapError('pdf', '编辑', err);
    }
}

/** 合并多个 PDF（按给定顺序拼接） */
export async function mergePdfs(inputs: DocInput[], spec: { fonts?: PdfFontSet } = {}): Promise<Uint8Array> {
    try {
        if (!inputs.length) throw new JdocError('没有可合并的 PDF', 'pdf');
        const out = await PDFDocument.create();
        // 注册 fontkit / 标准字体，让合并后的文档仍可继续叠加文本
        await buildFonts(out, spec.fonts);
        for (const input of inputs) {
            const src = await PDFDocument.load(await toUint8Array(input), { ignoreEncryption: true, updateMetadata: false });
            const pages = await out.copyPages(src, src.getPageIndices());
            pages.forEach((p) => out.addPage(p));
        }
        return await out.save();
    } catch (err) {
        throw wrapError('pdf', '合并', err);
    }
}

/** 拆分 PDF：按页码区间导出为多个新文件（区间为 1 起的闭区间） */
export async function splitPdf(input: DocInput, ranges: [number, number][]): Promise<Uint8Array[]> {
    try {
        const doc = await PDFDocument.load(await toUint8Array(input), { ignoreEncryption: true, updateMetadata: false });
        const total = doc.getPageCount();
        const parts: Uint8Array[] = [];
        for (const [from, to] of ranges) {
            const start = Math.max(0, Math.min(total - 1, Math.floor(from) - 1));
            const end = Math.max(start, Math.min(total - 1, Math.floor(to) - 1));
            const indices: number[] = [];
            for (let i = start; i <= end; i++) indices.push(i);
            if (!indices.length) continue;
            const part = await PDFDocument.create();
            const pages = await part.copyPages(doc, indices);
            pages.forEach((p) => part.addPage(p));
            parts.push(await part.save());
        }
        return parts;
    } catch (err) {
        throw wrapError('pdf', '拆分', err);
    }
}
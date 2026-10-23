/**
 * PDF 页面图像提取（含版面位置）
 *
 * pdfjs 的 `page.objs` 里放的是**解码后的原始像素**，且默认不在 operator list 阶段
 * 就绪（`objs.get()` 会抛「isn't resolved yet」，必须走回调异步取）。这里：
 *
 * 1. 走一遍渲染指令，维护 CTM，在 paintImage* 指令处用「单位方块经 CTM 变换」
 *    算出图像在页面上的实际位置与显示尺寸（PDF 用户坐标，pt，原点左下）
 * 2. 用回调方式把图像对象取出来（带超时，避免个别对象永远不 resolve 而挂住整条链路）
 * 3. 原始像素编码成 PNG（JPEG/PNG 直通时不重新编码）
 *
 * 之所以要位置：PDF 没有「块」的概念，图文是绝对定位的，只有拿到坐标才能把
 * 图片按阅读顺序插回文本流（否则 pdf -> md / html 会整篇丢图）。
 *
 * @module pdf/images
 */

import { sniffMime } from '../core/asset';
import { encodePng } from '../core/png';
import type { PdfImage } from './types';

/**
 * pdfjs 的 operator list
 *
 * 声明成宽松联合是刻意的：parse.ts 里把 fnArray 描述成 `number | string[]`
 * （pdfjs 真实的形态是数字码数组），严格声明会让调用点类型不兼容。
 */
interface PdfjsOperatorList {
    fnArray: ArrayLike<number | string> | number | string[];
    argsArray?: unknown[][];
}

/** pdfjs 的对象容器（同步 get 未就绪时会抛，带回调则异步返回） */
interface PdfjsObjectStore {
    get(name: string, callback?: (data: unknown) => void): unknown;
    has?(name: string): boolean;
}

/** pdfjs 页面（仅本模块用到的部分） */
export interface PdfjsImagePage {
    getOperatorList?(): Promise<PdfjsOperatorList>;
    objs?: PdfjsObjectStore;
    commonObjs?: PdfjsObjectStore;
}

/**
 * pdfjs 模块（需要 OPS 表把数字码还原成操作名）
 *
 * 索引签名是刻意的：调用方手里是完整的 pdfjs 模块对象（含 getDocument 等），
 * 这里只关心 OPS，没有索引签名会被 TS 判成「无共同属性」而不兼容。
 */
export interface PdfjsImageModule {
    OPS?: Record<string, number>;
    [key: string]: unknown;
}

/** pdfjs 解出的图像对象 */
interface PdfjsImageObject {
    width?: number;
    height?: number;
    kind?: number;
    data?: Uint8Array | ArrayBuffer;
}

/** 单位矩阵 [a,b,c,d,e,f] */
const IDENTITY = [1, 0, 0, 1, 0, 0];

/** 提取选项 */
export interface PdfImageOptions {
    /** 小于该显示尺寸（pt）的图不产出（背景填充、1x1 占位图等），默认 8 */
    minSizePt?: number;
    /** 小于该像素尺寸的图不产出，默认 8 */
    minPixels?: number;
    /**
     * 单张图的像素上限，超出后盒式降采样，默认 2e6（约 1600x1250）
     *
     * 扫描件整页图动辄 3000x4000（14MP），原样编码成 PNG 会让 md / html 膨胀到
     * 几百 MB，而它在页面上的显示尺寸通常只有几百像素。设为 0 表示不降采样。
     */
    maxPixels?: number;
    /** 单个图像对象的解析超时（ms），默认 5000 */
    timeoutMs?: number;
}

/** 参与 CTM 维护的操作码 */
interface OpCodes {
    save?: number;
    restore?: number;
    transform?: number;
    formBegin?: number;
    formEnd?: number;
    /** 绘制图像的操作码集合 */
    images: Set<number>;
    /** 数字码 -> 操作名（fnArray 为名字时的兜底） */
    names: Record<number, string>;
}

/** 从 pdfjs 的 OPS 表解析出需要的操作码（版本无关：pdfjs 6 的码值与旧版不同） */
function readOpCodes(pdfjs: PdfjsImageModule | undefined): OpCodes {
    const ops = pdfjs?.OPS || {};
    const names: Record<number, string> = {};
    const images = new Set<number>();
    const wanted: Record<string, number | undefined> = {};
    for (const [name, code] of Object.entries(ops)) {
        if (typeof code !== 'number') continue;
        names[code] = name;
        if (/^paint(Image|Jpeg|Inline)/.test(name) || /^paintImageMaskXObject$/.test(name)) images.add(code);
        if (name === 'save' || name === 'restore' || name === 'transform') wanted[name] = code;
        if (name === 'paintFormXObjectBegin' || name === 'paintFormXObjectEnd') wanted[name] = code;
    }
    return {
        save: wanted.save,
        restore: wanted.restore,
        transform: wanted.transform,
        formBegin: wanted.paintFormXObjectBegin,
        formEnd: wanted.paintFormXObjectEnd,
        images,
        names
    };
}

/**
 * 提取一页里的所有图像（含位置与字节）
 * @param pdfjs - pdfjs 模块（提供 OPS 表）
 * @param page - pdfjs 页面
 * @param options - 过滤与超时选项
 */
export async function extractPageImages(
    pdfjs: PdfjsImageModule | undefined,
    page: PdfjsImagePage,
    options: PdfImageOptions = {}
): Promise<PdfImage[]> {
    const opList = await page.getOperatorList?.();
    if (!opList) return [];

    const codes = readOpCodes(pdfjs);
    const minSizePt = options.minSizePt ?? 8;
    const minPixels = options.minPixels ?? 8;
    const maxPixels = options.maxPixels ?? 2e6;
    const timeoutMs = options.timeoutMs ?? 5000;

    const fnArray = (opList.fnArray as ArrayLike<number | string>) || [];
    const argsArray = opList.argsArray || [];
    const total = typeof fnArray.length === 'number' ? fnArray.length : 0;

    // objId -> 已收集的图像（同一张图被多次绘制时保留面积最大的那次）
    const collected = new Map<string, PdfImage>();
    const stack: number[][] = [];
    let ctm = IDENTITY.slice();

    for (let i = 0; i < total; i++) {
        const raw = fnArray[i];
        const code = typeof raw === 'number' ? raw : -1;
        const name = typeof raw === 'string' ? raw : codes.names[code];
        const args = (argsArray[i] || []) as unknown[];

        // 注意拼接方向：pdfjs 的 op list 用「后写的矩阵先作用于坐标」的行向量约定，
        // 即 CTM_new = CTM_old × M（若按列向量习惯写成 M × CTM_old，多段 cm 会把
        // 平移量再乘一次缩放，算出来的位置会跑到页面外）。
        if (code === codes.save) stack.push(ctm.slice());
        else if (code === codes.restore) ctm = stack.pop() || IDENTITY.slice();
        else if (code === codes.transform) ctm = mulMatrix(ctm, (args as number[]) || IDENTITY);
        else if (code === codes.formBegin) {
            stack.push(ctm.slice());
            ctm = mulMatrix(ctm, (args[0] as number[]) || IDENTITY);
        } else if (code === codes.formEnd) ctm = stack.pop() || ctm;
        else if (code >= 0 && codes.images.has(code)) {
            const arg0 = args[0];
            const objId = typeof arg0 === 'string' ? arg0 : `inline:${i}`;
            const inline = arg0 && typeof arg0 === 'object' ? (arg0 as PdfjsImageObject) : undefined;
            const rect = rectOfUnitSquare(ctm);
            if (!isFinite(rect.width) || !isFinite(rect.height)) continue;
            const prev = collected.get(objId);
            if (prev && prev.width * prev.height >= rect.width * rect.height) continue;

            const obj = inline || (await resolveImage(page, objId, timeoutMs));
            const image = await toPdfImage(obj, rect, objId, maxPixels);
            if (!image) continue;
            if (image.width < minSizePt || image.height < minSizePt) continue;
            if (image.pixelWidth < minPixels || image.pixelHeight < minPixels) continue;
            collected.set(objId, image);
        } else if (typeof raw === 'string' && name && /^paint(Image|Jpeg)/.test(name)) {
            // fnArray 直接给名字的版本（pdfjs 6+ 的部分分支）
            const rect = rectOfUnitSquare(ctm);
            const objId = typeof args[0] === 'string' ? args[0] : `inline:${i}`;
            const obj = (args[0] && typeof args[0] === 'object' ? args[0] : await resolveImage(page, objId, timeoutMs)) as
                | PdfjsImageObject
                | undefined;
            const image = await toPdfImage(obj, rect, objId, maxPixels);
            if (image && image.width >= minSizePt && image.height >= minSizePt) collected.set(objId, image);
        }
    }

    return [...collected.values()];
}

/** 单位方块 (0,0)-(1,1) 经 CTM 变换后的外接矩形（PDF 用户坐标，pt） */
function rectOfUnitSquare(ctm: number[]): { x: number; y: number; width: number; height: number } {
    const [a, b, c, d, e, f] = ctm;
    const xs = [e, a + e, c + e, a + c + e];
    const ys = [f, b + f, d + f, b + d + f];
    const x = Math.min(...xs);
    const y = Math.min(...ys);
    return { x, y, width: Math.max(...xs) - x, height: Math.max(...ys) - y };
}

/** 2D 仿射矩阵相乘（a·b），输入/输出均为 [a,b,c,d,e,f] */
function mulMatrix(a: number[], b: number[]): number[] {
    return [
        a[0] * b[0] + a[2] * b[1],
        a[1] * b[0] + a[3] * b[1],
        a[0] * b[2] + a[2] * b[3],
        a[1] * b[2] + a[3] * b[3],
        a[0] * b[4] + a[2] * b[5] + a[4],
        a[1] * b[4] + a[3] * b[5] + a[5]
    ];
}

/** 取图像对象：objs / commonObjs 都试，未就绪时用回调异步等（带超时） */
async function resolveImage(page: PdfjsImagePage, objId: string, timeoutMs: number): Promise<PdfjsImageObject | undefined> {
    const fromPage = await getObject(page.objs, objId, timeoutMs);
    if (fromPage) return fromPage;
    // 内联图像挂在 commonObjs 上
    return getObject(page.commonObjs, objId, timeoutMs);
}

function getObject(
    store: PdfjsObjectStore | undefined,
    id: string,
    timeoutMs: number
): Promise<PdfjsImageObject | undefined> {
    if (!store) return Promise.resolve(undefined);
    return new Promise((resolve) => {
        let settled = false;
        const done = (value: PdfjsImageObject | undefined) => {
            if (settled) return;
            settled = true;
            resolve(value);
        };
        try {
            const sync = store.get(id);
            if (sync) return done(sync as PdfjsImageObject);
        } catch {
            // 未就绪，继续走回调
        }
        let timer: ReturnType<typeof setTimeout> | undefined;
        try {
            store.get(id, (data) => {
                if (timer) clearTimeout(timer);
                done(data as PdfjsImageObject);
            });
        } catch {
            if (timer) clearTimeout(timer);
            return done(undefined);
        }
        timer = setTimeout(() => done(undefined), timeoutMs);
    });
}

/** 图像对象 -> 统一结构（原始像素编码成 PNG，已是 JPEG/PNG 的直通） */
async function toPdfImage(
    obj: PdfjsImageObject | undefined,
    rect: { x: number; y: number; width: number; height: number },
    objId: string,
    maxPixels: number
): Promise<PdfImage | undefined> {
    const raw = toBytes(obj?.data);
    if (!raw?.length) return undefined;
    const pixelWidth = Math.max(0, Math.round(obj?.width || 0));
    const pixelHeight = Math.max(0, Math.round(obj?.height || 0));
    if (!pixelWidth || !pixelHeight) return undefined;

    const sniffed = sniffMime(raw);
    if (sniffed.startsWith('image/')) {
        // 本身就是 JPEG/PNG 之类的编码流，直通不重新编码
        return buildImage(objId, rect, pixelWidth, pixelHeight, raw, sniffed);
    }

    // pdfjs 给的是解码后的原始像素（灰度/RGB/RGBA），需要自己编码
    const channels = channelsOf(obj?.kind, raw.length, pixelWidth, pixelHeight);
    let pixels = normalizePixels(raw, pixelWidth, pixelHeight, obj?.kind, channels);
    let width = pixelWidth;
    let height = pixelHeight;
    if (maxPixels > 0 && width * height > maxPixels) {
        // 扫描件整页图动辄 3000x4000，原样编码会让 md / html 膨胀到几百 MB；
        // 显示尺寸通常只有几百像素，这里按面积上限盒式降采样。
        const scaled = resample(pixels, width, height, channels, Math.sqrt(maxPixels / (width * height)));
        pixels = scaled.data;
        width = scaled.width;
        height = scaled.height;
    }
    const data = await encodePng({ data: pixels, width, height });
    return buildImage(objId, rect, width, height, data, 'image/png');
}

function buildImage(
    objId: string,
    rect: { x: number; y: number; width: number; height: number },
    pixelWidth: number,
    pixelHeight: number,
    data: Uint8Array,
    mime: string
): PdfImage {
    return {
        id: objId,
        x: Math.round(rect.x * 100) / 100,
        y: Math.round(rect.y * 100) / 100,
        width: Math.round(rect.width * 100) / 100,
        height: Math.round(rect.height * 100) / 100,
        pixelWidth,
        pixelHeight,
        data,
        mime
    };
}

/**
 * 取像素字节（pdfjs 各版本给的形态不一致：Uint8Array / Uint8ClampedArray / ArrayBuffer）
 *
 * 只看 `instanceof Uint8Array` 会漏掉 Uint8ClampedArray（pdfjs 对 RGB 图就给这个），
 * 表现为「大图一张都提取不到」。
 */
function toBytes(data: unknown): Uint8Array | undefined {
    if (!data) return undefined;
    if (data instanceof Uint8Array) return data;
    if (data instanceof Uint8ClampedArray) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
    if (data instanceof ArrayBuffer) return new Uint8Array(data);
    if (ArrayBuffer.isView(data)) {
        return new Uint8Array((data as ArrayBufferView).buffer, (data as ArrayBufferView).byteOffset, (data as ArrayBufferView).byteLength);
    }
    return undefined;
}

/** 通道数：1bpp 灰度 -> 1，RGB -> 3，RGBA -> 4，缺省按数据长度推断 */
function channelsOf(kind: number | undefined, dataLength: number, width: number, height: number): number {
    if (kind === 1) return 1;
    if (kind === 2) return 3;
    if (kind === 3) return 4;
    const perPixel = dataLength / Math.max(1, width * height);
    if (perPixel >= 4) return 4;
    if (perPixel >= 3) return 3;
    return 1;
}

/** 归一化成每通道 8 位（1bpp 灰度按位展开，其余原样拷贝） */
function normalizePixels(
    raw: Uint8Array,
    width: number,
    height: number,
    kind: number | undefined,
    channels: number
): Uint8Array {
    if (kind === 1) {
        // 1bpp：每字节 8 个像素，MSB 在先；位为 1 表示着色（黑），0 表示不着色（白）
        const srcStride = Math.ceil(width / 8);
        const out = new Uint8Array(width * height);
        for (let y = 0; y < height; y++) {
            for (let x = 0; x < width; x++) {
                const byte = raw[y * srcStride + (x >> 3)] ?? 0;
                out[y * width + x] = (byte >> (7 - (x & 7))) & 1 ? 0 : 255;
            }
        }
        return out;
    }
    const stride = width * channels;
    const out = new Uint8Array(stride * height);
    out.set(raw.subarray(0, Math.min(out.length, raw.length)), 0);
    return out;
}

/** 盒式降采样：输出像素取源区域平均值，避免最近邻采样的锯齿 */
function resample(
    src: Uint8Array,
    width: number,
    height: number,
    channels: number,
    scale: number
): { data: Uint8Array; width: number; height: number } {
    const outWidth = Math.max(1, Math.round(width * scale));
    const outHeight = Math.max(1, Math.round(height * scale));
    if (outWidth === width && outHeight === height) return { data: src, width, height };
    const out = new Uint8Array(outWidth * outHeight * channels);
    const xRatio = width / outWidth;
    const yRatio = height / outHeight;
    for (let y = 0; y < outHeight; y++) {
        const y0 = Math.floor(y * yRatio);
        const y1 = Math.min(height, Math.max(y0 + 1, Math.ceil((y + 1) * yRatio)));
        for (let x = 0; x < outWidth; x++) {
            const x0 = Math.floor(x * xRatio);
            const x1 = Math.min(width, Math.max(x0 + 1, Math.ceil((x + 1) * xRatio)));
            const dst = (y * outWidth + x) * channels;
            for (let c = 0; c < channels; c++) {
                let sum = 0;
                let count = 0;
                for (let sy = y0; sy < y1; sy++) {
                    const rowBase = sy * width * channels;
                    for (let sx = x0; sx < x1; sx++) {
                        sum += src[rowBase + sx * channels + c];
                        count++;
                    }
                }
                out[dst + c] = count ? Math.round(sum / count) : 0;
            }
        }
    }
    return { data: out, width: outWidth, height: outHeight };
}

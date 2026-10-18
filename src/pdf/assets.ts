/**
 * PDF 内嵌资源提取
 *
 * 两条通道：
 * 1. **附件**（`doc.getAttachments()`）——最可靠，附件本来就是独立文件
 * 2. **页面图像**（`page.getOperatorList()` + `page.objs`）——扫描件里的插图
 *
 * 页面图像默认不提取：它需要走一遍渲染指令，代价高（182 页课本会明显变慢），
 * 调用方按需开启 `assets: 'extract'` 即可。
 *
 * @module pdf/assets
 */

import {
    createAsset,
    createAssetId,
    mimeOf,
    sniffMime,
    type DocAsset,
    type DocAssetMode
} from '../core/asset';

/** pdfjs 操作符中与内嵌图像相关的几个码 */
const PAINT_IMAGE_OPS = new Set([
    'paintImageXObject',
    'paintJpegXObject',
    'paintInlineImageXObject',
    'paintImageMaskXObject'
]);

/** pdfjs 的附件描述（不同版本字段略有差异） */
export interface PdfAttachment {
    filename?: string;
    rawFilename?: string;
    description?: string;
    content?: Uint8Array;
}

/** pdfjs 文档对象（仅声明本模块用到的部分） */
export interface PdfAssetDocument {
    numPages: number;
    getPage(pageNumber: number): Promise<PdfAssetPage>;
    /** 附件清单：可能是 Map、数组或以文件名为键的对象 */
    getAttachments?(): Promise<Map<string, PdfAttachment> | PdfAttachment[] | Record<string, PdfAttachment> | null>;
    /** 按文件名取附件字节（pdfjs 惰性提供内容，描述里通常只有文件名） */
    getAttachmentContent?(filename: string): Promise<Uint8Array>;
}

/** pdfjs 页面对象 */
export interface PdfAssetPage {
    getOperatorList?(): Promise<{ fnArray: number | string[]; argsArray?: unknown[][] }>;
    objs?: { get(name: string): unknown; has?(name: string): boolean };
    commonObjs?: { get(name: string): unknown };
}

/** pdfjs 操作符编号（与 pdfjs-dist 的 OPS 表一致，用到的几项） */
const OPS = {
    paintImageXObject: 85,
    paintInlineImageXObject: 86,
    paintImageMaskXObject: 88,
    paintJpegXObject: 93
} as const;

const OP_NAMES: Record<number, string> = Object.fromEntries(
    Object.entries(OPS).map(([name, code]) => [code, name])
);

/**
 * 提取 PDF 的内嵌资源
 *
 * @param pdf - 已加载的 pdfjs 文档
 * @param options - 提取选项
 * @returns 资源数组（附件在前，页面图像在后）
 */
export async function extractPdfAssets(
    pdf: PdfAssetDocument,
    options: {
        mode?: DocAssetMode;
        /** 附件文件名在清单中的键，作为资源名 */
        maxPages?: number;
        pages?: number[];
        /** 操作符名映射（测试可注入） */
        opNames?: Record<number, string>;
    } = {}
): Promise<DocAsset[]> {
    const mode = options.mode ?? 'extract';
    if (mode === 'ignore') return [];

    const assets: DocAsset[] = [];
    const names = options.opNames ?? OP_NAMES;

    // 1) 附件：作为独立资源，锚点标为 other
    if (typeof pdf.getAttachments === 'function') {
        for (const [key, entry] of normalizeAttachments(await pdf.getAttachments())) {
            const filename = entry.filename || entry.rawFilename || key;
            // pdfjs 的附件描述里往往只有文件名，字节需再取一次
            let bytes = entry.content ? new Uint8Array(entry.content) : undefined;
            if (!bytes?.byteLength && typeof pdf.getAttachmentContent === 'function') {
                try {
                    const payload = await pdf.getAttachmentContent(filename);
                    if (payload?.byteLength) bytes = new Uint8Array(payload);
                } catch {
                    // 单个附件取不到不影响其它资源
                }
            }
            if (!bytes?.byteLength) continue;

            const id = createAssetId('pdf', assets.length + 1);
            assets.push(
                createAsset({
                    id,
                    data: bytes,
                    name: filename,
                    // 附件文件名通常带扩展名，比魔数更可信（如 note.txt）
                    mime: mimeOf(filename, sniffMime(bytes)),
                    part: `EmbeddedFile:${key}`,
                    anchor: {
                        scope: 'pdf',
                        kind: 'other',
                        index: assets.length,
                        path: `attachments/${key}`,
                        label: `附件 ${filename}`
                    },
                    meta: { attachmentName: key }
                })
            );
        }
    }

    // 2) 页面图像：走渲染指令找出被绘制的图像对象
    const wanted = new Set(options.pages || []);
    const limit = options.maxPages || pdf.numPages;
    const seen = new Set<string>();

    for (let pageNo = 1; pageNo <= pdf.numPages && assets.length - pageNo < limit; pageNo++) {
        if (wanted.size && !wanted.has(pageNo)) continue;
        let page: PdfAssetPage;
        try {
            page = await pdf.getPage(pageNo);
        } catch {
            continue;
        }
        if (typeof page.getOperatorList !== 'function') continue;

        let operatorList: { fnArray: number | string[]; argsArray?: unknown[][] };
        try {
            operatorList = await page.getOperatorList();
        } catch {
            continue;
        }

        const fnArray = operatorList.fnArray as ArrayLike<number | string>;
        const argsArray = operatorList.argsArray || [];
        for (let i = 0; i < fnArray.length; i++) {
            const fn = fnArray[i];
            const name = typeof fn === 'string' ? fn : names[fn];
            if (!name || !PAINT_IMAGE_OPS.has(name)) continue;

            const args = argsArray[i] || [];
            // 按页隔离：同一张图常在多页复用，键必须带页码才不会互相吞掉
            const key = `p${pageNo}:${name === 'paintInlineImageXObject' ? `inline${i}` : String(args[0])}`;
            if (!key || key === 'undefined' || seen.has(key)) continue;
            seen.add(key);

            const image = lookupImage(page, key.split(':').slice(1).join(':') || key);
            const bytes = image?.data ? new Uint8Array(image.data as ArrayBufferLike) : undefined;
            if (!bytes?.byteLength) continue;

            const id = createAssetId('pdf', assets.length + 1);
            assets.push(
                createAsset({
                    id,
                    data: bytes,
                    mime: sniffMime(bytes),
                    part: `Page${pageNo}:${name}`,
                    anchor: {
                        scope: 'pdf',
                        kind: 'page',
                        index: pageNo,
                        path: `pages[${pageNo - 1}]`,
                        label: `第 ${pageNo} 页`
                    },
                    width: typeof image?.width === 'number' ? image.width : undefined,
                    height: typeof image?.height === 'number' ? image.height : undefined
                })
            );
        }
    }

    return assets;
}

/** 附件清单归一化为 [键, 描述] 列表（pdfjs 可能是 Map / 数组 / 对象） */
function normalizeAttachments(
    raw: Map<string, PdfAttachment> | PdfAttachment[] | Record<string, PdfAttachment> | null | undefined
): [string, PdfAttachment][] {
    if (!raw) return [];
    if (raw instanceof Map) return Array.from(raw.entries());
    if (Array.isArray(raw)) {
        return raw.map((entry, index) => [entry?.filename || `attachment-${index + 1}`, entry] as [string, PdfAttachment]);
    }
    return Object.entries(raw);
}

/** 从页面对象表里取出图像数据 */
function lookupImage(
    page: PdfAssetPage,
    key: string
): { data?: unknown; width?: number; height?: number } | undefined {
    const fromPage = safeGet(page.objs, key);
    if (fromPage) return fromPage as { data?: unknown; width?: number; height?: number };
    // 内联图像挂在 commonObjs 上，且键名带前缀
    const suffix = key.split(':').pop() || key;
    return safeGet(page.commonObjs, suffix) as { data?: unknown; width?: number; height?: number } | undefined;
}

function safeGet(store: { get(name: string): unknown } | undefined, key: string): unknown {
    if (!store) return undefined;
    try {
        return store.get(key);
    } catch {
        return undefined;
    }
}
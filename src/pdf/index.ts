/**
 * PDF 解析与生成公开入口
 *
 * @module pdf
 */

import { toArrayBuffer, type DocInput } from '../core/input';
import { wrapError } from '../core/errors';
import type { DocAsset } from '../core/asset';
import { parsePdf } from './parse';
import { generatePdf } from './generate';
import { editPdf, mergePdfs, splitPdf } from './edit';
import type {
    PdfDocument,
    PdfDocumentSpec,
    PdfGenerateOptions,
    PdfMetadata,
    PdfPage,
    PdfParseOptions
} from './types';

export * from './types';
export * from './edit';
export { editPdf, mergePdfs, splitPdf };

/** 解析结果 */
export interface PdfResult {
    /** 文档元数据 */
    metadata?: PdfMetadata;
    /** 总页数 */
    pageCount: number;
    /** 页面详情（取决于解析选项） */
    pages?: PdfPage[];
    /**
     * 内嵌资源（附件 / 页面图像），仅当 `assets` 选项非 `ignore` 时返回
     *
     * 附件锚点为 `kind: 'other'`，页面图像为 `kind: 'page'` 并记录页码。
     */
    assets?: DocAsset[];
    /** 解析提示 */
    messages: string[];
}

/**
 * 解析 PDF：返回元数据、页数与逐页文本
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 解析选项（items=true 时额外返回带坐标的文本片段）
 *
 * @example
 * ```js
 * const { pageCount, pages } = await pdfToJson(file);
 * console.log(pageCount, pages[0].text);
 * ```
 */
export async function pdfToJson(input: DocInput, options: PdfParseOptions = {}): Promise<PdfResult> {
    try {
        const { document, messages } = await parsePdf(input, options);
        return {
            metadata: document.metadata,
            pageCount: document.pageCount,
            pages: document.pages,
            assets: document.assets,
            messages
        };
    } catch (error) {
        throw wrapError('pdf', '解析', error);
    }
}

/**
 * 提取 PDF 全文（按页用空行分隔）
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 解析选项
 *
 * @example
 * ```js
 * const text = await pdfToText(file);
 * ```
 */
export async function pdfToText(input: DocInput, options: PdfParseOptions = {}): Promise<string> {
    try {
        const { document } = await parsePdf(input, { ...options, items: false });
        return (document.pages || [])
            .map((page) => page.text || '')
            .filter((text) => text.length > 0)
            .join('\n\n');
    } catch (error) {
        throw wrapError('pdf', '解析', error);
    }
}

/**
 * 读取 PDF 基本信息（页数、尺寸、元数据），不提取正文，速度最快
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 */
export async function pdfInfo(input: DocInput): Promise<PdfDocument & { messages: string[] }> {
    try {
        const { document, messages } = await parsePdf(input, { text: false });
        return { ...document, messages };
    } catch (error) {
        throw wrapError('pdf', '解析', error);
    }
}

/** 页面渲染所需的画布实现（Node 端需额外安装 @napi-rs/canvas） */
const CANVAS_PACKAGE = '@napi-rs/canvas';

/**
 * 把指定页渲染为 PNG 图片数据
 * @param input - 文件二进制
 * @param pageNumber - 页号（1 起），默认第 1 页
 * @param scale - 缩放倍数，默认 2（约 144dpi）
 * @returns PNG 字节数据
 * @remarks Node 端需先安装 `@napi-rs/canvas` 才能渲染；浏览器端无需额外依赖。
 */
export async function pdfPageToPng(
    input: DocInput,
    pageNumber = 1,
    scale = 2
): Promise<Uint8Array> {
    try {
        const data = await toArrayBuffer(input);
        const canvasFactory = (await import(/* @vite-ignore */ CANVAS_PACKAGE).catch(() => undefined)) as
            | {
                  createCanvas: (w: number, h: number) => {
                      getContext: (type: '2d') => unknown;
                      toBuffer: (mime: string) => Buffer;
                  };
              }
            | undefined;

        const entries = ['pdfjs-dist/legacy/build/pdf.mjs', 'pdfjs-dist/build/pdf.mjs'];
        let pdfjs: PdfjsForRender | undefined;
        for (const entry of entries) {
            try {
                pdfjs = (await import(/* @vite-ignore */ entry)) as PdfjsForRender;
                break;
            } catch {
                // 继续尝试下一个入口
            }
        }
        if (!pdfjs) throw new Error('pdfjs-dist 未安装，无法渲染 PDF 页面');
        if (!canvasFactory) {
            throw new Error('缺少画布实现，Node 端请先安装：npm i @napi-rs/canvas');
        }

        const pdf = await pdfjs.getDocument({ data, isEvalSupported: false }).promise;
        const page = await pdf.getPage(pageNumber);
        const viewport = page.getViewport({ scale });
        const canvas = canvasFactory.createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
        await page.render({
            canvasContext: canvas.getContext('2d') as never,
            viewport,
            canvas: canvas as never
        }).promise;
        return new Uint8Array(canvas.toBuffer('image/png'));
    } catch (error) {
        throw wrapError('pdf', '页面渲染', error);
    }
}

/** 页面渲染所需的 pdfjs 最小接口 */
interface PdfjsForRender {
    getDocument(params: Record<string, unknown>): {
        promise: Promise<{
            getPage: (pageNumber: number) => Promise<{
                getViewport: (options: { scale: number }) => { width: number; height: number };
                render: (options: Record<string, unknown>) => { promise: Promise<void> };
            }>;
        }>;
    };
}

/**
 * 生成 PDF 文件
 * @param spec - 文档规格（页面或元素序列）
 * @param options - 生成选项
 *
 * @example
 * ```js
 * const bytes = await jsonToPdf({
 *   metadata: { title: '报告' },
 *   elements: [
 *     { type: 'heading', level: 1, text: '季度报告' },
 *     { type: 'text', text: '正文内容……' },
 *     { type: 'table', header: true, rows: [['项', '值'], ['营收', '1200']] }
 *   ]
 * });
 * ```
 */
export async function jsonToPdf(
    spec: PdfDocumentSpec,
    options: PdfGenerateOptions = {}
): Promise<Uint8Array> {
    return generatePdf(spec, options);
}

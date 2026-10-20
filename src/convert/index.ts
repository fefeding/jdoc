/**
 * 文档互转统一门面
 *
 * 以「文档块模型（DocxBlock）」为中间表示，提供任意格式间的转换：
 *
 *   md / docx / html / pdf  ->  blocks  ->  md / docx / html / pdf
 *
 * 各格式分别提供细粒度函数（如 `markdownToBlocks` / `blocksToPdf`），
 * `convert()` 则是「自动识别源格式 -> 转成目标格式」的一站式入口。
 *
 * @module convert
 */

import { toUint8Array, detectFormat, type DocInput } from '../core/input';
import type { DocFormat } from '../core/errors';
import { docxToJson, jsonToDocx, blocksToHtml } from '../docx/index';
import { markdownToBlocks, blocksToMarkdown } from './markdown';
import { htmlToBlocks } from './html';
import { pdfToBlocks } from './pdf';
import { blocksToPdf } from './blocks-to-pdf';
import type { DocxBlock, DocxDocument } from '../docx/types';
import type { DocAsset } from '../core/asset';
import type { PdfFontSet } from '../pdf/types';

export { markdownToBlocks, blocksToMarkdown } from './markdown';
export { htmlToBlocks } from './html';
export { pdfToBlocks } from './pdf';
export { blocksToPdf } from './blocks-to-pdf';
export type { MarkdownToBlocksOptions, BlocksToMarkdownOptions } from './markdown';
export type { PdfToBlocksOptions, PdfToBlocksResult } from './pdf';
export type { BlocksToPdfOptions } from './blocks-to-pdf';

/** 可转换的目标格式 */
export type ConvertTarget = 'docx' | 'pdf' | 'md' | 'html';

/** convert 选项 */
export interface ConvertOptions {
    /** 源格式（缺省自动识别；文本类 Html 需显式指定 'html'） */
    from?: DocFormat | 'html';
    /** 内嵌字体（生成 PDF 且含中文时必须） */
    fonts?: PdfFontSet;
    /** 是否自动探测系统字体（仅 Node 端），默认 true */
    autoFont?: boolean;
    /** PDF 解析质量：'fast' | 'layout'，默认 'layout' */
    quality?: 'fast' | 'layout';
    /** 资源表，用于还原图片 assetId */
    assets?: DocAsset[] | Record<string, DocAsset> | { assets?: DocAsset[] };
    /**
     * pdfjs worker 地址（Node 端）。当 pdfjs 被打包器内联进 bundle（VS Code 扩展、
     * Electron 应用等）时，它内置的相对 worker 路径会失效，PDF 解析会直接失败，
     * 需显式给出 worker 文件的地址。
     */
    workerSrc?: string;
}

/** 把字节解码为 UTF-8 文本 */
function decodeText(bytes: Uint8Array): string {
    if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) bytes = bytes.subarray(3);
    return new TextDecoder('utf-8', { fatal: false }).decode(bytes);
}

/** 粗略判断文本是否像 HTML */
function looksLikeHtml(text: string): boolean {
    return /^\s*<(!doctype\s+html|html|head|body|div|p|table|section|article)\b/i.test(text);
}

/**
 * 格式互转：源格式 -> 文档块模型 -> 目标格式
 *
 * @param input - 源文件字节 / data URL / base64 / Blob
 * @param target - 目标格式
 * @param options - 源格式指定、字体、资源表等
 * @returns 二进制目标（docx / pdf）返回 Uint8Array；文本目标（md / html）返回字符串
 *
 * @example
 * ```js
 * const pdf = await convert(docxBytes, 'pdf');            // docx -> pdf
 * const md  = await convert(pdfBytes, 'md');              // pdf -> markdown
 * const html = await convert(markdownBytes, 'html');       // markdown -> html
 * ```
 */
export async function convert(
    input: DocInput,
    target: ConvertTarget,
    options: ConvertOptions = {}
): Promise<Uint8Array | string> {
    let from = options.from as string | undefined;
    // 字符串输入：文本类格式（md/html/txt）按原始文本处理；二进制格式（docx/pdf）
    // 按 base64 / data URL 处理（与 DocInput 约定一致）。data: 前缀一律走 base64。
    let bytes: Uint8Array;
    const textish = from === 'md' || from === 'html' || from === 'txt' || from === undefined;
    if (typeof input === 'string' && !input.startsWith('data:') && textish) {
        bytes = new TextEncoder().encode(input);
    } else {
        bytes = await toUint8Array(input);
    }

    if (!from) {
        const detected = await detectFormat(bytes);
        if (detected) from = detected;
        else {
            // 文本类：按内容嗅探 html，否则当作 markdown
            const text = decodeText(bytes);
            from = looksLikeHtml(text) ? 'html' : 'md';
        }
    }

    // ---- 源 -> blocks ----
    let blocks: DocxBlock[] = [];
    let metadata: DocxDocument['metadata'];
    let assets: DocAsset[] | undefined;

    if (from === 'docx') {
        const result = await docxToJson(bytes, { assets: 'extract' });
        blocks = result.blocks;
        metadata = result.metadata;
        assets = result.assets;
    } else if (from === 'md') {
        blocks = markdownToBlocks(bytes);
    } else if (from === 'html') {
        blocks = htmlToBlocks(decodeText(bytes));
    } else if (from === 'pdf') {
        const result = await pdfToBlocks(bytes, { quality: options.quality, workerSrc: options.workerSrc });
        blocks = result.blocks;
        metadata = result.metadata as DocxDocument['metadata'];
    } else {
        // txt 或其它：当作纯文本 markdown
        blocks = markdownToBlocks(decodeText(bytes));
    }

    // ---- blocks -> 目标 ----
    const pool = mergeAssets(options.assets, assets);
    if (target === 'docx') {
        return jsonToDocx({ metadata, blocks, assets: Array.isArray(pool) ? pool : (pool ? Object.values(pool) : undefined) });
    }
    if (target === 'pdf') {
        return blocksToPdf({ metadata, blocks }, { assets: pool, fonts: options.fonts, autoFont: options.autoFont });
    }
    if (target === 'md') {
        return blocksToMarkdown({ blocks }, { assets: pool });
    }
    if (target === 'html') {
        return blocksToHtml({ metadata, blocks }, { includeStyle: true });
    }
    throw new TypeError(`不支持的目标格式：${target}`);
}

/** 合并调用方传入的资源表与解析得到的资源表（统一为数组形态） */
function mergeAssets(
    fromOptions: ConvertOptions['assets'],
    fromParse: DocAsset[] | undefined
): DocAsset[] | undefined {
    const list: DocAsset[] = [];
    if (fromOptions) {
        if (Array.isArray(fromOptions)) list.push(...fromOptions);
        else if ('assets' in fromOptions && fromOptions.assets) list.push(...(fromOptions.assets as DocAsset[] | undefined || []));
        else list.push(...Object.values(fromOptions));
    }
    if (fromParse) list.push(...fromParse);
    return list.length ? list : undefined;
}

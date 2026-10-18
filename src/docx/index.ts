/**
 * DOCX 解析与生成公开入口
 *
 * 命名与 pptxToHtml / jsonToPptx 保持一致的风格：
 * 解析函数 xxxToYyy(input, options)，生成函数 jsonToXxx(doc, options)。
 *
 * @module docx
 */

import { wrapError } from '../core/errors';
import type { DocInput } from '../core/input';
import type { DocAsset, DocAssetMode } from '../core/asset';
import { blocksToText, parseDocxInternal } from './parse';
import { generateDocx } from './generate';
import { docxToHtmlString } from './to-html';
import type {
    DocxBlock,
    DocxDocument,
    DocxGenerateOptions,
    DocxMetadata,
    DocxParseOptions,
    DocxSection
} from './types';

export * from './types';

/** 解析结果公共字段 */
export interface DocxParseResult {
    /** 文档元数据 */
    metadata?: DocxMetadata;
    /** 页面设置（纸张与页边距） */
    section?: DocxSection;
    /** 内容块序列（可直接回传给 jsonToDocx 实现往返） */
    blocks: DocxBlock[];
    /** 解析提示（缺部件、图片丢失等），不影响主流程 */
    messages: string[];
    /** 全文纯文本 */
    text: string;
    /**
     * 文档内二进制资源（图片等），仅在 `assets` 选项为 `extract` / `both` 时返回
     *
     * 每项带 `anchor` 指明它在文档中的位置，块上通过 `assetId` 引用。
     */
    assets?: DocAsset[];
}

/** HTML 解析结果 */
export interface DocxHtmlResult extends DocxParseResult {
    /** 渲染好的 HTML（含默认样式表） */
    html: string;
}

/** docxToJson 的可选项 */
export interface DocxToJsonOptions extends DocxParseOptions {
    /**
     * 是否把图片内联为 data URL。
     * 默认 true；显式关闭可显著减小返回体积。
     */
    includeImages?: boolean;
    /**
     * 二进制资源处理方式，默认 `inline`（块里是 data URL）
     *
     * - `inline`：内联 data URL，浏览器直接显示
     * - `extract`：字节进 `assets`，块上只留 `assetId`（JSON 体积小，适合后台存储）
     * - `both`：两者都给
     * - `ignore`：丢弃
     *
     * @example
     * ```js
     * const { blocks, assets } = await docxToJson(file, { assets: 'extract' });
     * // 后台落盘
     * for (const a of assets) await writeFile(a.name, assetToNodeStream(a));
     * ```
     */
    assets?: DocAssetMode;
}

/** docxToHtml 的可选项 */
export interface DocxToHtmlOptions extends DocxToJsonOptions {
    /** 是否附带默认样式表，默认 true */
    includeStyle?: boolean;
}

/**
 * 解析 DOCX 为结构化 JSON
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 解析选项
 * @returns { metadata, blocks, messages, text }
 *
 * @example
 * ```js
 * const { blocks, text } = await docxToJson(file);
 * ```
 */
export async function docxToJson(
    input: DocInput,
    options: DocxToJsonOptions = {}
): Promise<DocxParseResult> {
    try {
        const output = await parseDocxInternal(
            input,
            options.includeImages !== false,
            options.assets ?? 'inline'
        );
        const { metadata, blocks, section, assets } = output.document;
        return { metadata, section, blocks, assets, messages: output.messages, text: output.text };
    } catch (error) {
        throw wrapError('docx', '解析', error);
    }
}

/**
 * 解析 DOCX 为 HTML
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 解析选项
 * @returns { html, metadata, section, blocks, messages, text }
 *
 * @example
 * ```js
 * const { html } = await docxToHtml(file);
 * container.innerHTML = html;
 * ```
 */
export async function docxToHtml(
    input: DocInput,
    options: DocxToHtmlOptions = {}
): Promise<DocxHtmlResult> {
    try {
        const output = await parseDocxInternal(
            input,
            options.includeImages !== false,
            options.assets ?? 'inline'
        );
        const { metadata, blocks, section, assets } = output.document;
        // HTML 预览一律用内联 data URL，否则 <img> 没有可加载的地址
        const html = docxToHtmlString(
            assets?.length ? { metadata, section, blocks, assets } : { metadata, section, blocks },
            { includeStyle: options.includeStyle !== false }
        );
        return { html, metadata, section, blocks, assets, messages: output.messages, text: output.text };
    } catch (error) {
        throw wrapError('docx', '解析', error);
    }
}

/**
 * 提取 DOCX 全文纯文本
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 */
export async function docxToText(input: DocInput): Promise<string> {
    try {
        const output = await parseDocxInternal(input, false);
        return output.text;
    } catch (error) {
        throw wrapError('docx', '解析', error);
    }
}

/**
 * 把块结构渲染为 HTML（不经过文件解析，便于从 JSON 直接预览）
 * @param input - DocxDocument 或 blocks 数组
 */
export function blocksToHtml(
    input: DocxDocument | DocxBlock[],
    options: { includeStyle?: boolean } = {}
): string {
    return docxToHtmlString(input, options);
}

/**
 * 把块序列提取为纯文本
 * @param blocks - 内容块序列
 */
export function jsonToText(blocks: DocxBlock[]): string {
    return blocksToText(blocks);
}

/**
 * 生成 DOCX 文件
 * @param input - DocxDocument，或直接传 blocks 数组（简写）
 * @param options - 页面与默认样式
 * @returns DOCX 字节数据
 *
 * @example
 * ```js
 * const bytes = await jsonToDocx({
 *   metadata: { title: '周报' },
 *   blocks: [
 *     { type: 'heading', level: 1, text: '本周进展' },
 *     { type: 'paragraph', text: '完成 A、B 模块' },
 *     { type: 'table', header: true, rows: [['模块','进度'],['A','100%']] }
 *   ]
 * });
 * ```
 */
export async function jsonToDocx(
    input: DocxDocument | DocxBlock[],
    options: DocxGenerateOptions = {}
): Promise<Uint8Array> {
    try {
        return await generateDocx(input, options);
    } catch (error) {
        throw wrapError('docx', '生成', error);
    }
}

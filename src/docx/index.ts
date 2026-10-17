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
import { blocksToText, parseDocxInternal } from './parse';
import { generateDocx } from './generate';
import { docxToHtmlString } from './to-html';
import type { DocxBlock, DocxDocument, DocxGenerateOptions, DocxMetadata, DocxParseOptions } from './types';

export * from './types';

/** 解析结果公共字段 */
export interface DocxParseResult {
    /** 文档元数据 */
    metadata?: DocxMetadata;
    /** 内容块序列（可直接回传给 jsonToDocx 实现往返） */
    blocks: DocxBlock[];
    /** 解析提示（缺部件、图片丢失等），不影响主流程 */
    messages: string[];
    /** 全文纯文本 */
    text: string;
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
        const output = await parseDocxInternal(input, options.includeImages !== false);
        const { metadata, blocks } = output.document;
        return { metadata, blocks, messages: output.messages, text: output.text };
    } catch (error) {
        throw wrapError('docx', '解析', error);
    }
}

/**
 * 解析 DOCX 为 HTML
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 解析选项
 * @returns { html, metadata, blocks, messages, text }
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
        const output = await parseDocxInternal(input, options.includeImages !== false);
        const { metadata, blocks } = output.document;
        const html = docxToHtmlString(
            { metadata, blocks },
            { includeStyle: options.includeStyle !== false }
        );
        return { html, metadata, blocks, messages: output.messages, text: output.text };
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

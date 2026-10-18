/**
 * Markdown 解析
 *
 * 基于 marked（GFM）把 Markdown 渲染为可直接展示的 HTML，并对内嵌原始 HTML 做白名单清洗。
 * 与其余格式保持一致的命名：`markdownToHtml` / `markdownToText`。
 *
 * @module markdown
 */

import { marked } from 'marked';
import { wrapError } from '../core/errors';
import { toUint8Array, type DocInput } from '../core/input';
import { sanitizeHtml, htmlToText, escapeHtml } from './sanitize';
import { MARKDOWN_CSS } from './to-html';
import type { MarkdownHeading, MarkdownHtmlResult, MarkdownMetadata, MarkdownOptions } from './types';

export * from './types';

/** 字节 -> UTF-8 文本（Markdown/TXT 均为文本格式） */
function decodeText(bytes: Uint8Array): string {
    // 去掉 UTF-8 BOM
    if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
        bytes = bytes.subarray(3);
    }
    return new TextDecoder('utf-8', { fatal: false }).decode(bytes);
}

/** 从 token 流推断元信息（标题大纲 / 字数） */
function readMetadata(markdown: string, text: string): MarkdownMetadata | undefined {
    const headings: MarkdownHeading[] = [];
    try {
        for (const token of marked.lexer(markdown)) {
            if (token.type === 'heading') {
                headings.push({ level: (token as { depth?: number }).depth || 1, text: (token as { text?: string }).text || '' });
            }
        }
    } catch {
        // 大纲提取失败不影响正文渲染
    }
    const metadata: MarkdownMetadata = {};
    const first = headings.find((h) => h.level === 1) || headings[0];
    if (first?.text) metadata.title = first.text;
    if (headings.length) metadata.headings = headings;
    metadata.wordCount = text.replace(/\s+/g, '').length;
    return Object.keys(metadata).length ? metadata : undefined;
}

/**
 * 渲染 Markdown 为 HTML
 * @param input - Uint8Array / ArrayBuffer / Blob / base64 字符串（与其它格式一致，字符串按 base64 解析）
 * @param options - 渲染选项
 * @returns { html, text, metadata, messages }
 *
 * @example
 * ```js
 * const { html } = await markdownToHtml(bytes);
 * container.innerHTML = html;
 * ```
 */
export async function markdownToHtml(
    input: DocInput,
    options: MarkdownOptions = {}
): Promise<MarkdownHtmlResult> {
    const messages: string[] = [];
    try {
        const markdown = decodeText(await toUint8Array(input));
        const allowHtml = options.allowHtml !== false;

        const renderer = new marked.Renderer();
        if (!allowHtml) {
            // 关闭原始 HTML：让 marked 把内嵌 HTML 当作纯文本输出（自动转义）
            renderer.html = (token: { text?: string }) => escapeHtml(String(token?.text ?? ''));
        }
        const raw = marked.parse(markdown, {
            gfm: options.gfm !== false,
            breaks: options.breaks === true,
            async: false,
            renderer
        }) as unknown;
        const rendered = typeof raw === 'string' ? raw : String(raw ?? '');

        const html = allowHtml ? sanitizeHtml(rendered) : rendered;
        if (allowHtml && /<(script|iframe|style)\b/i.test(rendered)) {
            messages.push('Markdown 中的原始 HTML 已按白名单清洗（脚本/嵌入内容被移除）');
        }

        const text = htmlToText(html);
        const metadata = readMetadata(markdown, text);
        const style = options.includeStyle === false ? '' : `<style>${MARKDOWN_CSS}</style>`;

        return {
            html: `${style}<div class="jdoc-md">${html}</div>`,
            text,
            metadata,
            messages
        };
    } catch (error) {
        throw wrapError('md', '解析', error);
    }
}

/**
 * 提取 Markdown 的纯文本（去标记）
 * @param input - 同 markdownToHtml
 */
export async function markdownToText(input: DocInput, options: MarkdownOptions = {}): Promise<string> {
    try {
        const { text } = await markdownToHtml(input, { ...options, includeStyle: false });
        return text;
    } catch (error) {
        throw wrapError('md', '解析', error);
    }
}

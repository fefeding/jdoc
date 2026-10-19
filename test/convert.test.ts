import { describe, it, expect } from 'vitest';
import {
    markdownToBlocks,
    blocksToMarkdown,
    htmlToBlocks,
    pdfToBlocks,
    blocksToPdf,
    convert
} from '../src/convert/index';
import { jsonToDocx } from '../src/docx/index';
import { blocksToHtml } from '../src/docx/index';
import { docxToJson } from '../src/docx/index';
import type { DocxBlock } from '../src/docx/types';

describe('convert: markdown <-> blocks', () => {
    it('解析标题/段落/列表/表格/代码块', () => {
        const md = [
            '# 标题一',
            '',
            '这是一段**加粗**与*斜体*混排的文字。',
            '',
            '- 项目一',
            '- 项目二',
            '  - 嵌套项',
            '',
            '| 列A | 列B |',
            '| --- | --- |',
            '| 1 | 2 |',
            '',
            '```js',
            'const a = 1;',
            '```'
        ].join('\n');

        const blocks = markdownToBlocks(md);
        expect(blocks[0]).toMatchObject({ type: 'heading', level: 1 });
        const para = blocks.find((b) => b.type === 'paragraph') as Extract<DocxBlock, { type: 'paragraph' }>;
        expect(para.runs?.some((r) => r.bold)).toBe(true);
        const list = blocks.find((b) => b.type === 'list') as Extract<DocxBlock, { type: 'list' }>;
        expect(list.items.length).toBe(2);
        const nested = (list.items[1] as unknown as { children?: unknown[] }).children;
        expect(Array.isArray(nested)).toBe(true);
        const table = blocks.find((b) => b.type === 'table') as Extract<DocxBlock, { type: 'table' }>;
        expect(table.rows.length).toBe(2);
        const code = blocks.find((b) => b.type === 'code') as Extract<DocxBlock, { type: 'code' }>;
        expect(code.language).toBe('js');
    });

    it('blocks -> markdown 可往返（md -> blocks -> md 关键字保留）', () => {
        const md = '# 标题\n\n正文 **粗** 与 `代码`。\n\n- a\n- b\n';
        const back = blocksToMarkdown(markdownToBlocks(md));
        expect(back).toContain('# 标题');
        expect(back).toContain('**粗**');
        expect(back).toContain('`代码`');
        expect(back).toContain('- a');
    });
});

describe('convert: html -> blocks', () => {
    it('解析标题/段落/列表/表格/图片', () => {
        const html = `
            <h1>标题</h1>
            <p>段落 <strong>粗</strong> 与 <em>斜</em>。</p>
            <ul><li>一</li><li>二</li></ul>
            <table><tr><th>A</th><th>B</th></tr><tr><td>1</td><td>2</td></tr></table>
            <img src="data:image/png;base64,xxx" alt="图" />
        `;
        const blocks = htmlToBlocks(html);
        expect(blocks[0]).toMatchObject({ type: 'heading', level: 1 });
        const img = blocks.find((b) => b.type === 'image') as Extract<DocxBlock, { type: 'image' }>;
        expect(img.src).toBe('data:image/png;base64,xxx');
        const table = blocks.find((b) => b.type === 'table') as Extract<DocxBlock, { type: 'table' }>;
        expect(table.rows.length).toBe(2);
        expect(table.header).toBe(true);
    });
});

describe('convert: 统一门面', () => {
    it('markdown -> html 返回带标签的字符串', async () => {
        const html = (await convert('# 你好\n\n正文', 'html')) as string;
        expect(html).toContain('<h1');
        expect(html).toContain('你好');
    });

    it('markdown -> docx 返回 DOCX 字节', async () => {
        const bytes = (await convert('# 标题\n\n段落', 'docx')) as Uint8Array;
        expect(bytes[0]).toBe(0x50); // P
        expect(bytes[1]).toBe(0x4b); // K
    });

    it('docx -> markdown 保留标题', async () => {
        const blocks: DocxBlock[] = [
            { type: 'heading', level: 1, text: '从 DOCX 来' },
            { type: 'paragraph', text: '一段内容' }
        ];
        const docx = await jsonToDocx({ blocks });
        const md = (await convert(docx, 'md', { from: 'docx' })) as string;
        expect(md).toContain('# 从 DOCX 来');
        expect(md).toContain('一段内容');
    });

    it('docx -> docx 往返（blocks 数量一致）', async () => {
        const blocks: DocxBlock[] = [
            { type: 'heading', level: 2, text: '章节' },
            { type: 'paragraph', text: '内容 A' },
            { type: 'paragraph', text: '内容 B' }
        ];
        const docx = await jsonToDocx({ blocks });
        const parsed = await docxToJson(docx);
        expect(parsed.blocks.length).toBe(3);
    });
});

describe('convert: PDF 生成与解析', () => {
    it('blocks -> pdf 生成可解析的 PDF（无中文无需字体）', async () => {
        const blocks: DocxBlock[] = [
            { type: 'heading', level: 1, text: 'Hello Title' },
            { type: 'paragraph', text: 'This is a paragraph of English text for layout test.' },
            { type: 'list', ordered: false, items: ['first', 'second'] }
        ];
        const pdf = await blocksToPdf(blocks);
        expect(pdf[0]).toBe(0x25); // %
        expect(pdf[1]).toBe(0x50); // P
        expect(pdf[2]).toBe(0x44); // D
        expect(pdf[3]).toBe(0x46); // F
    });

    it('pdf -> blocks 还原出标题', async () => {
        const blocks: DocxBlock[] = [
            { type: 'heading', level: 1, text: 'Chapter One' },
            { type: 'paragraph', text: 'Body text of the chapter.' }
        ];
        const pdf = await blocksToPdf(blocks);
        const result = await pdfToBlocks(pdf, { quality: 'fast' });
        expect(result.blocks.some((b) => b.type === 'heading' && /Chapter One/.test((b as { text?: string }).text || ''))).toBe(true);
    });

    it('blocks -> pdf 中文自动探测系统字体（Node 端）', async () => {
        const blocks: DocxBlock[] = [{ type: 'paragraph', text: '中文内容测试生成 PDF' }];
        const pdf = await blocksToPdf(blocks, { autoFont: true });
        expect(pdf[0]).toBe(0x25);
        const result = await pdfToBlocks(pdf, { quality: 'fast' });
        expect(result.blocks.some((b) => /中文内容测试生成 PDF/.test((b as { text?: string }).text || ''))).toBe(true);
    });
});

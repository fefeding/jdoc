import { describe, expect, it } from 'vitest';
import { docxToHtml, docxToJson, docxToText, jsonToDocx, blocksToHtml } from '../src/docx/index';
import type { DocxBlock, DocxDocument } from '../src/docx/types';

/** 收集块里的所有纯文本，便于断言 */
function collectText(blocks: DocxBlock[]): string {
    return JSON.stringify(blocks);
}

const sample: DocxDocument = {
    metadata: { title: '季度报告', author: 'jdoc', subject: '测试' },
    blocks: [
        { type: 'heading', level: 1, text: '季度报告' },
        { type: 'paragraph', text: '本季度营收增长显著。' },
        {
            type: 'paragraph',
            runs: [
                { text: '重点 ' },
                { text: '加粗', bold: true },
                { text: '，' },
                { text: '彩色', color: '#FF0000', size: 14 },
                { text: '，' },
                { text: '链接', link: 'https://example.com' }
            ]
        },
        { type: 'list', ordered: false, items: ['要点一', '要点二', '要点三'] },
        { type: 'list', ordered: true, items: ['第一步', '第二步'] },
        {
            type: 'table',
            header: true,
            rows: [
                ['模块', '进度', '备注'],
                ['核心', '100%', '已完成'],
                ['外围', '60%', '进行中']
            ]
        },
        { type: 'quote', text: '引用一段话' },
        { type: 'hr' },
        { type: 'pageBreak' },
        { type: 'heading', level: 2, text: '第二部分' },
        { type: 'paragraph', text: '结尾段落', align: 'center' }
    ]
};

describe('docx：JSON -> DOCX -> JSON 往返', () => {
    it('生成的文件是合法 ZIP 容器', async () => {
        const bytes = await jsonToDocx(sample);
        expect(bytes.byteLength).toBeGreaterThan(1000);
        // ZIP 魔数
        expect(bytes[0]).toBe(0x50);
        expect(bytes[1]).toBe(0x4b);
    });

    it('保留标题层级与文本', async () => {
        const { blocks } = await docxToJson(await jsonToDocx(sample));
        const headings = blocks.filter((b) => b.type === 'heading');
        expect(headings.length).toBe(2);
        expect(headings[0]).toMatchObject({ type: 'heading', level: 1 });
        expect(collectText(headings[0])).toContain('季度报告');
    });

    it('保留 run 级样式（加粗 / 颜色 / 字号 / 超链接）', async () => {
        const { blocks } = await docxToJson(await jsonToDocx(sample));
        const styled = blocks.find((b) => b.type === 'paragraph' && 'runs' in b && b.runs && b.runs.length > 1);
        expect(styled).toBeDefined();

        const runs = (styled as { runs: Record<string, unknown>[] }).runs;
        const bold = runs.find((r) => r.bold === true);
        expect(bold?.text).toBe('加粗');

        const colored = runs.find((r) => typeof r.color === 'string' && r.color !== '');
        expect(String(colored?.color).toUpperCase()).toContain('FF0000');
        expect(colored?.size).toBe(14);

        const linked = runs.find((r) => r.link);
        expect(linked?.link).toBe('https://example.com');
    });

    it('保留有序/无序列表类型', async () => {
        const { blocks } = await docxToJson(await jsonToDocx(sample));
        const lists = blocks.filter((b) => b.type === 'list') as { ordered?: boolean; items: unknown[] }[];
        expect(lists.length).toBe(2);
        expect(lists.some((l) => l.ordered === false)).toBe(true);
        expect(lists.some((l) => l.ordered === true)).toBe(true);
        expect(lists[0].items.length).toBe(3);
    });

    it('保留表格内容与行列', async () => {
        const { blocks } = await docxToJson(await jsonToDocx(sample));
        const table = blocks.find((b) => b.type === 'table') as { rows: unknown[][]; header?: boolean } | undefined;
        expect(table).toBeDefined();
        expect(table!.rows.length).toBe(3);
        expect(table!.rows[0].length).toBe(3);
        expect(collectText(table!.rows[0])).toContain('模块');
    });

    it('保留元数据', async () => {
        const { metadata } = await docxToJson(await jsonToDocx(sample));
        expect(metadata?.title).toBe('季度报告');
    });

    it('提取全文文本', async () => {
        const text = await docxToText(await jsonToDocx(sample));
        expect(text).toContain('季度报告');
        expect(text).toContain('要点一');
        expect(text).toContain('模块');
    });
});

describe('docx：HTML 输出', () => {
    it('解析时返回可直接插入的 HTML', async () => {
        const bytes = await jsonToDocx(sample);
        const { html } = await docxToHtml(bytes);
        expect(html).toContain('<h1');
        expect(html).toContain('<table');
        expect(html).toContain('<ul>');
        expect(html).toContain('<ol>');
        expect(html).toContain('<blockquote');
        expect(html).toContain('docx-body');
    });

    it('HTML 正确转义特殊字符', async () => {
        const bytes = await jsonToDocx([{ type: 'paragraph', text: '<script>alert(1)</script>' }]);
        const { html } = await docxToHtml(bytes);
        expect(html).not.toContain('<script>alert(1)</script>');
        expect(html).toContain('&lt;script&gt;');
    });

    it('支持从 JSON 直接渲染 HTML（无需生成文件）', () => {
        const html = blocksToHtml(sample.blocks);
        expect(html).toContain('季度报告');
    });

    it('includeStyle=false 时不带样式表', async () => {
        const bytes = await jsonToDocx(sample);
        const { html } = await docxToHtml(bytes, { includeStyle: false });
        expect(html).not.toContain('<style>');
    });
});

describe('docx：异常处理', () => {
    it('空内容也能生成合法文档', async () => {
        const bytes = await jsonToDocx({ blocks: [] });
        expect(bytes.byteLength).toBeGreaterThan(500);
    });

    it('解析非法数据抛带格式标识的错误', async () => {
        await expect(docxToJson(new Uint8Array([1, 2, 3]))).rejects.toThrow(/docx/);
    });

    it('缺少 blocks 时给出明确错误', async () => {
        await expect(jsonToDocx({} as never)).rejects.toThrow(/blocks/);
    });
});

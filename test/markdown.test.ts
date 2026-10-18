import { describe, expect, it } from 'vitest';
import { markdownToHtml, markdownToText, textToHtml, detectFormat } from '../src/index';

const enc = new TextEncoder();
const bytes = (s: string) => enc.encode(s);

describe('markdown：渲染为 HTML', () => {
    it('渲染标题、强调与代码', async () => {
        const { html } = await markdownToHtml(bytes('# 标题\n\n这是 **粗体** 与 `code`。\n'));
        expect(html).toContain('<h1>标题</h1>');
        expect(html).toContain('<strong>粗体</strong>');
        expect(html).toContain('<code>code</code>');
        expect(html).toContain('class="jdoc-md"');
    });

    it('支持 GFM 表格与任务列表', async () => {
        const { html } = await markdownToHtml(bytes('| a | b |\n|---|---|\n| 1 | 2 |\n'));
        expect(html).toContain('<table>');
        expect(html).toContain('<th>a</th>');
        expect(html).toContain('<td>1</td>');
    });

    it('推断元信息：标题大纲与字数', async () => {
        const { metadata } = await markdownToHtml(bytes('# 一级\n\n## 二级\n\n正文'));
        expect(metadata?.title).toBe('一级');
        expect(metadata?.headings?.length).toBe(2);
        expect(metadata?.headings?.[1]).toEqual({ level: 2, text: '二级' });
        expect(metadata?.wordCount).toBeGreaterThan(0);
    });

    it('清洗内嵌的危险 HTML（script / 事件属性 / javascript: 协议）', async () => {
        const { html, messages } = await markdownToHtml(
            bytes('<script>alert(1)</script>\n\n<img src="javascript:alert(2)" onerror="evil()">\n')
        );
        expect(html).not.toMatch(/<script/i);
        expect(html).not.toMatch(/onerror/i);
        expect(html).not.toMatch(/javascript:/i);
        expect(messages.some((m) => m.includes('清洗'))).toBe(true);
    });

    it('allowHtml 关闭时原始 HTML 被转义', async () => {
        const { html } = await markdownToHtml(bytes('<b>x</b>\n'), { allowHtml: false });
        expect(html).toContain('&lt;b&gt;x&lt;/b&gt;');
    });

    it('includeStyle 关闭时不带样式表', async () => {
        const { html } = await markdownToHtml(bytes('# t\n'), { includeStyle: false });
        expect(html).not.toContain('<style>');
    });

    it('markdownToText 去标记', async () => {
        const text = await markdownToText(bytes('# 标题\n\n**粗体** 段落\n'));
        expect(text).toContain('标题');
        expect(text).toContain('粗体');
        expect(text).not.toContain('#');
    });
});

describe('text：纯文本渲染', () => {
    it('转义 HTML 特殊字符', async () => {
        const { html, text } = await textToHtml(bytes('a < b & "c"'));
        expect(html).toContain('&lt;');
        expect(html).toContain('&amp;');
        expect(text).toBe('a < b & "c"');
    });

    it('保留换行与缩进（pre-wrap）', async () => {
        const { html } = await textToHtml(bytes('line1\n  line2'));
        expect(html).toContain('class="jdoc-txt"');
        expect(html).toContain('line1\n  line2');
    });
});

describe('detectFormat：文本类格式嗅探', () => {
    it('Markdown 识别为 md', async () => {
        expect(await detectFormat(bytes('# 标题\n\n- 列表\n'))).toBe('md');
    });

    it('普通文本识别为 txt', async () => {
        expect(await detectFormat(bytes('just a note\nsecond line\n'))).toBe('txt');
    });

    it('二进制格式不受影响', async () => {
        expect(await detectFormat(bytes('%PDF-1.7\n'))).toBe('pdf');
    });
});

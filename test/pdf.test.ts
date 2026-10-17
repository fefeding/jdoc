import { describe, expect, it } from 'vitest';
import { jsonToPdf, pdfInfo, pdfToJson, pdfToText } from '../src/pdf/index';
import type { PdfDocumentSpec } from '../src/pdf/types';

const sample: PdfDocumentSpec = {
    metadata: { title: 'Quarterly Report', author: 'jdoc', subject: 'demo' },
    defaultSize: 12,
    elements: [
        { type: 'heading', level: 1, text: 'Quarterly Report' },
        { type: 'text', text: 'Revenue grew steadily across all regions this quarter.' },
        {
            type: 'paragraphless' as never,
            text: ''
        },
        { type: 'heading', level: 2, text: 'Details' },
        {
            type: 'text',
            text: 'Body copy with a link to the source.',
            color: '#3366CC',
            size: 11
        },
        {
            type: 'table',
            header: true,
            columnWidths: [200, 120, 160],
            rows: [
                ['Item', 'Amount', 'Note'],
                ['Service', '$1,200', 'Paid'],
                ['Support', '$300', 'Pending']
            ]
        },
        { type: 'line', x1: 0, y1: 0, x2: 400, y2: 0, lineWidth: 2, color: '#CCCCCC' },
        { type: 'rect', x: 0, y: 0, width: 120, height: 40, fill: '#E8F0FE', stroke: '#3366CC', lineWidth: 1 },
        { type: 'rect', x: 140, y: 0, width: 120, height: 40, fill: '#FFF4E5', radius: 8 },
        { type: 'pageBreak' },
        { type: 'heading', level: 1, text: 'Second Page' },
        { type: 'text', text: 'Content after the explicit page break.' }
    ]
};

describe('pdf：生成', () => {
    it('生成的文件带 %PDF- 头且非空', async () => {
        const bytes = await jsonToPdf(sample);
        expect(bytes.byteLength).toBeGreaterThan(1000);
        expect(new TextDecoder().decode(bytes.slice(0, 5))).toBe('%PDF-');
    });

    it('多元素文档可生成', async () => {
        const bytes = await jsonToPdf({
            elements: [
                { type: 'heading', level: 1, text: 'A' },
                { type: 'text', text: 'B' },
                { type: 'table', rows: [['x', 'y'], ['1', '2']] }
            ]
        });
        expect(bytes.byteLength).toBeGreaterThan(500);
    });

    it('支持显式分页配置', async () => {
        const bytes = await jsonToPdf({
            pages: [
                { size: 'A5', orientation: 'landscape', elements: [{ type: 'text', text: 'page one' }] },
                { size: [400, 500], elements: [{ type: 'text', text: 'page two' }] }
            ]
        });
        expect(bytes.byteLength).toBeGreaterThan(500);
    });

    it('未内嵌字体时拒绝中文并给出明确指引', async () => {
        await expect(jsonToPdf({ elements: [{ type: 'text', text: '中文内容' }] })).rejects.toThrow(/字体/);
    });

    it('非法输入给出明确错误', async () => {
        await expect(jsonToPdf(null as never)).rejects.toThrow(/规格/);
    });
});

describe('pdf：解析（PDF -> JSON）', () => {
    it('读回页数与页面尺寸', async () => {
        const bytes = await jsonToPdf({ elements: [{ type: 'text', text: 'Hello PDF' }] });
        const { pageCount, pages } = await pdfToJson(bytes);
        expect(pageCount).toBe(1);
        expect(pages?.[0].width).toBeGreaterThan(500);
        expect(pages?.[0].height).toBeGreaterThan(700);
    });

    it('读回元数据', async () => {
        const bytes = await jsonToPdf(sample);
        const { metadata } = await pdfToJson(bytes);
        expect(metadata?.title).toBe('Quarterly Report');
        expect(metadata?.author).toBe('jdoc');
    });

    it('提取文本内容', async () => {
        const bytes = await jsonToPdf(sample);
        const text = await pdfToText(bytes);
        expect(text).toContain('Quarterly Report');
        expect(text).toContain('Revenue grew steadily');
        expect(text).toContain('Second Page');
        expect(text).toContain('$1,200');
    });

    it('分页后文本用空行分隔', async () => {
        const bytes = await jsonToPdf(sample);
        const text = await pdfToText(bytes);
        expect(text).toContain('\n\n');
    });

    it('items=true 时返回带坐标的文本片段', async () => {
        const bytes = await jsonToPdf({ elements: [{ type: 'text', text: 'Positioned' }] });
        const { pages } = await pdfToJson(bytes, { items: true });
        const items = pages?.[0].items;
        expect(items && items.length).toBeGreaterThan(0);
        expect(typeof items?.[0]?.x).toBe('number');
        expect(typeof items?.[0]?.y).toBe('number');
        expect(items?.[0]?.fontSize).toBeGreaterThan(0);
    });

    it('pdfInfo 只取元信息（不提取正文）', async () => {
        const bytes = await jsonToPdf(sample);
        const info = await pdfInfo(bytes);
        expect(info.pageCount).toBe(2);
        expect(info.pages === undefined || info.pages.every((p) => p.text === undefined)).toBe(true);
    });

    it('maxPages 限制返回页数', async () => {
        const bytes = await jsonToPdf(sample);
        const { pages } = await pdfToJson(bytes, { maxPages: 1 });
        expect(pages?.length).toBe(1);
    });

    it('非 PDF 数据抛出明确错误', async () => {
        await expect(pdfToJson(new TextEncoder().encode('NOT A PDF FILE'))).rejects.toThrow(/PDF/);
    });
});

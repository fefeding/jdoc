import { describe, it, expect } from 'vitest';
import { jsonToPptx } from '../src/pptx/index';
import { pptxDocumentToBlocks, pptxToBlocks } from '../src/convert/pptx';
import { convert } from '../src/convert/index';
import type { PptxDocument } from '../src/pptx/index';
import type { DocxBlock } from '../src/docx/types';

/** 构造一个最小的演示文稿：标题 + 项目符号 + 图片 + 表格 + 图表 */
function buildDeck(): PptxDocument {
    return {
        version: '1.0',
        slideSize: { width: 960, height: 540 },
        metadata: { title: '示例', author: '测试' },
        slides: [
            {
                elements: [
                    {
                        type: 'text',
                        x: 0,
                        y: 0,
                        width: 960,
                        height: 80,
                        fontSize: 32,
                        paragraphs: [{ text: '季度汇报', runs: [{ text: '季度汇报', bold: true, fontSize: 32 }] }]
                    },
                    {
                        type: 'text',
                        x: 0,
                        y: 100,
                        width: 960,
                        height: 300,
                        paragraphs: [
                            { text: '完成目标', bullet: true, runs: [{ text: '完成目标' }] },
                            { text: '存在问题', bullet: true, runs: [{ text: '存在问题' }] }
                        ]
                    },
                    {
                        type: 'image',
                        x: 0,
                        y: 420,
                        width: 120,
                        height: 90,
                        data: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
                    }
                ]
            },
            {
                elements: [
                    {
                        type: 'text',
                        x: 0,
                        y: 0,
                        width: 960,
                        height: 80,
                        fontSize: 28,
                        paragraphs: [{ text: '数据', runs: [{ text: '数据', bold: true, fontSize: 28 }] }]
                    },
                    {
                        type: 'table',
                        x: 0,
                        y: 120,
                        width: 400,
                        height: 120,
                        rows: [
                            { cells: [{ text: '合并标题', colSpan: 2 }] },
                            { cells: [{ text: 'A' }, { text: 'B' }] },
                            { cells: [{ text: 'C' }, { text: 'D' }] }
                        ]
                    },
                    {
                        type: 'chart',
                        x: 0,
                        y: 260,
                        width: 400,
                        height: 240,
                        title: '销量',
                        chartType: 'barChart',
                        categories: ['Q1', 'Q2'],
                        series: [{ name: '销量', values: [10, 20] }]
                    }
                ]
            }
        ]
    };
}

describe('convert: pptx -> blocks', () => {
    const deck = buildDeck();

    it('每张幻灯片产出标题，标题取字号最大的文本框', () => {
        const { blocks } = pptxDocumentToBlocks(deck);
        const headings = blocks.filter((b) => b.type === 'heading') as Extract<DocxBlock, { type: 'heading' }>[];
        expect(headings.length).toBe(2);
        expect(headings[0].text).toBe('季度汇报');
        expect(headings[1].text).toBe('数据');
    });

    it('项目符号分组为列表', () => {
        const { blocks } = pptxDocumentToBlocks(deck);
        const list = blocks.find((b) => b.type === 'list') as Extract<DocxBlock, { type: 'list' }>;
        expect(list).toBeTruthy();
        expect(list.items.length).toBe(2);
        expect(list.items[0]).toEqual([{ text: '完成目标' }]);
    });

    it('图片元素 -> 图片块', () => {
        const { blocks } = pptxDocumentToBlocks(deck);
        const image = blocks.find((b) => b.type === 'image') as Extract<DocxBlock, { type: 'image' }>;
        expect(image).toBeTruthy();
        expect(String(image.src).startsWith('data:image/png')).toBe(true);
    });

    it('表格合并单元格还原为 colspan，延续单元格被跳过', () => {
        const { blocks } = pptxDocumentToBlocks(deck);
        const tables = blocks.filter((b) => b.type === 'table') as Extract<DocxBlock, { type: 'table' }>[];
        // 第 1 张是图表数据表，第 2 张是演示表格
        const table = tables.find((t) => t.rows.length === 3);
        expect(table).toBeTruthy();
        expect(table!.rows[0].length).toBe(1); // 合并标题占满两列
        expect((table!.rows[0][0] as { colspan?: number }).colspan).toBe(2);
    });

    it('图表 -> 图表块（矢量 spec + 兜底数据表）', () => {
        const { blocks } = pptxDocumentToBlocks(deck);
        const chart = blocks.find((b) => b.type === 'chart') as Extract<DocxBlock, { type: 'chart' }>;
        expect(chart).toBeTruthy();
        // OOXML 的 barChart 默认 barDir=col，即纵向柱形
        expect(chart.spec.type).toBe('column');
        expect(chart.spec.title).toBe('销量');
        expect(chart.spec.series?.[0]?.values).toEqual([10, 20]);
        // 画不了矢量图的目标（md / docx / pdf）退化成这张数据表
        expect(chart.fallback?.rows.length).toBe(3);
        expect((chart.fallback!.rows[2][1] as { text?: string }).text).toBe('20');
    });
});

describe('convert: pptx 统一门面', () => {
    it('pptx -> md / html / docx / pdf 全部成功', async () => {
        const bytes = await jsonToPptx(buildDeck());
        expect(bytes[0]).toBe(0x50);

        const md = (await convert(bytes, 'md', { from: 'pptx' })) as string;
        expect(md).toContain('# 季度汇报');
        expect(md).toContain('- 完成目标');
        // 含合并单元格的表输出 HTML 表格（管道表无法表达 colspan）
        expect(md).toContain('<table>');
        expect(md).toContain('colspan="2"');
        expect(md).toContain('合并标题');

        const html = (await convert(bytes, 'html', { from: 'pptx' })) as string;
        expect(html).toContain('<table');

        const docx = (await convert(bytes, 'docx', { from: 'pptx' })) as Uint8Array;
        expect(docx[0]).toBe(0x50);

        const pdf = (await convert(bytes, 'pdf', { from: 'pptx' })) as Uint8Array;
        expect(pdf[0]).toBe(0x25);
    }, 60000);

    it('pptxToBlocks 透传元数据', async () => {
        const bytes = await jsonToPptx(buildDeck());
        const { metadata } = await pptxToBlocks(bytes);
        expect(metadata?.title).toBe('示例');
        expect(metadata?.author).toBe('测试');
    });
});

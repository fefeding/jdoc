/**
 * 目录（TOC）还原验证：docx 目录块 -> PDF 的排版与页码回填。
 *
 * 覆盖：条目文本 + 点线导引 + 右对齐页码、层级缩进、实测页码两遍回填、
 * 无匹配标题时回退 TOC 域缓存页码。
 */
import { describe, expect, it } from 'vitest';
import { blocksToPdf } from '../src/convert/blocks-to-pdf';
import { detectSystemFonts } from '../src/pdf/fonts';
import type { DocxBlock } from '../src/docx/types';

async function loadPdfjs(): Promise<any> {
    const mod: any = await import('pdfjs-dist/legacy/build/pdf.mjs').catch(() => import('pdfjs-dist'));
    return mod;
}

/** 读 PDF：每页文本项（含坐标，用于验证右对齐与缩进） */
async function pdfItems(bytes: Uint8Array): Promise<{ str: string; x: number; y: number }[][]> {
    const pdfjs = await loadPdfjs();
    const doc = await pdfjs.getDocument({ data: bytes.slice() }).promise;
    const pages: { str: string; x: number; y: number }[][] = [];
    for (let p = 1; p <= doc.numPages; p++) {
        const page = await doc.getPage(p);
        const content = await page.getTextContent();
        pages.push(
            (content.items as Array<{ str: string; transform: number[] }>)
                .filter((it) => it.str.trim())
                .map((it) => ({ str: it.str, x: it.transform[4], y: it.transform[5] }))
        );
    }
    return pages;
}

const TOC_BLOCK: DocxBlock = {
    type: 'toc',
    entries: [
        { level: 1, text: '第一章 概述' },
        { level: 2, text: '1.1 编写目的' },
        { level: 1, text: '第二章 安装' }
    ]
};

describe('目录（TOC）还原', () => {
    it('渲染条目文本、实测页码，页码右对齐到内容区右端', async () => {
        const fonts = await detectSystemFonts();
        const blocks: DocxBlock[] = [
            TOC_BLOCK,
            { type: 'heading', level: 1, text: '第一章 概述' },
            { type: 'paragraph', runs: [{ text: '内容'.repeat(800) }] },
            { type: 'heading', level: 2, text: '1.1 编写目的' },
            { type: 'paragraph', runs: [{ text: '内容'.repeat(800) }] },
            { type: 'heading', level: 1, text: '第二章 安装' },
            { type: 'paragraph', runs: [{ text: '内容'.repeat(800) }] }
        ];
        const pages = await pdfItems(await blocksToPdf(blocks, { fonts }));

        // 第 1 页应含目录条目文本 + 页码（实测：各标题所在页 >= 2）
        const tocPage = pages[0];
        const texts = tocPage.map((it) => it.str);
        expect(texts).toContain('第一章 概述');
        expect(texts).toContain('1.1 编写目的');
        expect(texts).toContain('第二章 安装');
        // 页码：应出现 3 个（实测页码，非域缓存）
        const numbers = tocPage.map((it) => it.str).filter((s) => /^\d+$/.test(s));
        expect(numbers.length).toBe(3);
        // 目录顺序与文档顺序一致：页码非递减
        const pageNums = numbers.map(Number);
        expect(pageNums.every((n, i) => i === 0 || n >= pageNums[i - 1])).toBe(true);
        // 正文有 3 段长内容，末尾标题页码应 >= 2
        expect(Math.max(...pageNums)).toBeGreaterThanOrEqual(2);

        // 页码右对齐：A4 宽 595.28 - 右边距 90 = 505.28（右端），数字宽约 0.5em = 5.25pt
        for (const n of numbers) {
            const item = tocPage.find((it) => it.str === n);
            expect(item).toBeDefined();
            const right = item!.x + n.length * 5.25;
            expect(right).toBeGreaterThan(503);
            expect(right).toBeLessThan(507);
        }
    });

    it('层级缩进：一级条目左对齐边距，二级条目右移', async () => {
        const fonts = await detectSystemFonts();
        const pages = await pdfItems(
            await blocksToPdf(
                [
                    TOC_BLOCK,
                    { type: 'heading', level: 1, text: '第一章 概述' },
                    { type: 'heading', level: 2, text: '1.1 编写目的' },
                    { type: 'heading', level: 1, text: '第二章 安装' }
                ],
                { fonts }
            )
        );
        const find = (text: string) => pages[0].find((it) => it.str === text)!;
        const l1 = find('第一章 概述');
        const l2 = find('1.1 编写目的');
        expect(l2).toBeDefined();
        // 二级缩进 18pt
        expect(l2.x - l1.x).toBeGreaterThan(14);
        expect(l2.x - l1.x).toBeLessThan(22);
    });

    it('匹配不到标题时回退 TOC 域缓存页码', async () => {
        const fonts = await detectSystemFonts();
        const blocks: DocxBlock[] = [
            { type: 'toc', entries: [{ level: 1, text: '不存在的标题', page: 7 }] },
            { type: 'paragraph', runs: [{ text: '正文' }] }
        ];
        const pages = await pdfItems(await blocksToPdf(blocks, { fonts }));
        expect(pages[0].map((it) => it.str)).toContain('7');
    });

    it('页码回填不影响分页：两遍排版页数一致（目录页数稳定）', async () => {
        const fonts = await detectSystemFonts();
        const entries = Array.from({ length: 30 }, (_, i) => ({ level: 1, text: `第${i}章` }));
        const blocks: DocxBlock[] = [
            { type: 'toc', entries },
            ...entries.map((e) => ({ type: 'heading', level: 1, text: e.text }) as DocxBlock),
            { type: 'paragraph', runs: [{ text: '正文'.repeat(300) }] }
        ];
        const withToc = await pdfItems(await blocksToPdf(blocks, { fonts }));
        // 同样的内容但无目录块：目录页数不影响后续分页的关键是目录自身占的行数固定
        const noTocBlocks = blocks.filter((b) => b.type !== 'toc');
        const withoutToc = await pdfItems(await blocksToPdf(noTocBlocks, { fonts }));
        // 无目录时 31 个元素；有目录时多出目录占的页（1 页）
        expect(withToc.length).toBe(withoutToc.length + 1);
    });
});
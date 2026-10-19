/**
 * 分页行为验证：blocksToPdf 输出的 PDF 用 pdfjs 读回，检查页数与关键排版指标。
 *
 * 覆盖：行级跨页（段落应在行边界断开而非整段挪页）、Word 中文默认参数
 * （A4 + 上下 72pt / 左右 90pt 边距 + 10.5pt 正文）、行距/段距透传。
 */
import { describe, expect, it } from 'vitest';
import { PDFDocument } from 'pdf-lib';
import { blocksToPdf } from '../src/convert/blocks-to-pdf';
import { detectSystemFonts } from '../src/pdf/fonts';
import type { DocxBlock } from '../src/docx/types';

/** 与 src/pdf/parse.ts 一致的 pdfjs 加载方式（Node 下用 legacy 构建） */
async function loadPdfjs(): Promise<any> {
    const mod: any = await import('pdfjs-dist/legacy/build/pdf.mjs').catch(() => import('pdfjs-dist'));
    return mod;
}

/** 读 PDF 每页文本与首个文本项坐标 */
async function pdfPages(
    bytes: Uint8Array
): Promise<{ text: string; firstX: number; firstY: number }[]> {
    const pdfjs = await loadPdfjs();
    const doc = await pdfjs.getDocument({ data: bytes.slice() }).promise;
    const out: { text: string; firstX: number; firstY: number }[] = [];
    for (let p = 1; p <= doc.numPages; p++) {
        const page = await doc.getPage(p);
        const content = await page.getTextContent();
        const items = content.items as Array<{ str: string; transform: number[] }>;
        const text = items.map((it) => it.str).join('');
        const first = items.find((it) => it.str.trim());
        out.push({
            text,
            firstX: first ? first.transform[4] : 0,
            firstY: first ? first.transform[5] : 0
        });
    }
    return out;
}

describe('blocksToPdf 分页与 Word 排版基线', () => {
    it('长段落按行跨页而非整段挪页（页尾不留整段空白）', async () => {
        const fonts = await detectSystemFonts();
        // A4 可用高 841.89-144≈697.9pt；每行 10.5×1.3≈13.65pt → 每页约 51 行
        // 造约 70 行的段落 + 次段：若整段挪页，第一页只会有空段落
        const para1: DocxBlock = {
            type: 'paragraph',
            runs: [{ text: '甲'.repeat(70 * 36) }]
        };
        const para2: DocxBlock = { type: 'paragraph', runs: [{ text: '第二段落开始' }] };
        const bytes = await blocksToPdf([para1, para2], { fonts });
        const pages = await pdfPages(bytes);
        expect(pages.length).toBe(2);
        // 第一页必须有第一段内容（行级跨页），且第二段开头在第二页
        expect(pages[0].text).toContain('甲');
        expect(pages[1].text).toContain('第二段落开始');
        // 第二页以第一段的延续开头（段落被拆开，而非整段挪页）
        expect(pages[1].text.startsWith('甲')).toBe(true);
    });

    it('页边距与纸张为 Word 中文默认（A4，上下 72pt，左右 90pt）', async () => {
        const fonts = await detectSystemFonts();
        const bytes = await blocksToPdf([{ type: 'paragraph', runs: [{ text: '边距测试' }] }], { fonts });
        const doc = await PDFDocument.load(bytes);
        const { width, height } = doc.getPage(0).getSize();
        expect(Math.round(width)).toBe(595);
        expect(Math.round(height)).toBe(842);
        const pages = await pdfPages(bytes);
        // 顶边距 72：首行基线 ≈ 842-72-10.5 ≈ 759；左边距 90：x ≈ 90
        expect(pages[0].firstY).toBeGreaterThan(750);
        expect(pages[0].firstY).toBeLessThan(765);
        expect(pages[0].firstX).toBeGreaterThan(88);
        expect(pages[0].firstX).toBeLessThan(92);
    });

    it('行距透传参与分页（2 倍行距页数多于单倍行距）', async () => {
        const fonts = await detectSystemFonts();
        const lines: DocxBlock[] = [];
        for (let i = 0; i < 60; i++) lines.push({ type: 'paragraph', runs: [{ text: `第${i}段落内容测试` }] });

        const single = (await pdfPages(await blocksToPdf(lines, { fonts }))).length;
        const spaced = lines.map((b) => {
            const p = b as Extract<DocxBlock, { type: 'paragraph' }>;
            return { ...p, lineSpacing: 2 };
        });
        const double = (await pdfPages(await blocksToPdf(spaced, { fonts }))).length;
        expect(double).toBeGreaterThan(single);
    });

    it('表格跨页重复表头', async () => {
        const fonts = await detectSystemFonts();
        const rows: string[][] = [['表头一', '表头二']];
        for (let i = 0; i < 48; i++) rows.push([`很长很长很长很长很长很长的单元格内容${i}`, `B${i}`]);
        const bytes = await blocksToPdf([{ type: 'table', header: true, rows } as DocxBlock], { fonts });
        const pages = await pdfPages(bytes);
        expect(pages.length).toBeGreaterThan(1);
        // 跨页后第二页应再次出现表头文本
        expect(pages[1].text).toContain('表头一');
        expect(pages[0].text).toContain('B0');
    });

    it('spacingBefore 参与分页（大段距把后续内容推到下一页）', async () => {
        const fonts = await detectSystemFonts();
        const blocks: DocxBlock[] = [];
        for (let i = 0; i < 45; i++) blocks.push({ type: 'paragraph', runs: [{ text: `段${i}` }] });
        blocks.push({
            type: 'paragraph',
            runs: [{ text: '被推走的段落' }],
            spacingBefore: 600
        } as DocxBlock);
        const bytes = await blocksToPdf(blocks, { fonts });
        const pages = await pdfPages(bytes);
        expect(pages.length).toBe(2);
        expect(pages[1].text).toContain('被推走的段落');
    });
});

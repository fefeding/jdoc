import { describe, it, expect } from 'vitest';
import { jsonToXlsx } from '../src/xlsx/index';
import { xlsxToBlocks } from '../src/convert/xlsx';
import { convert } from '../src/convert/index';
import type { DocxBlock, DocxCell } from '../src/docx/types';

/** 表格单元格 -> 纯文本（兼容 string / run 数组 / detail 三种形态） */
const cellText = (cell: DocxCell | undefined): string => {
    if (cell == null) return '';
    if (typeof cell === 'string') return cell;
    if (Array.isArray(cell)) return cell.map((r) => r.text ?? '').join('');
    const detail = cell as { runs?: { text?: string }[]; text?: string };
    if (detail.runs) return detail.runs.map((r) => r.text ?? '').join('');
    return detail.text ?? '';
};

describe('convert: xlsx -> blocks', () => {
    it('工作表映射为「标题 + 表格」', async () => {
        const bytes = await jsonToXlsx([
            { name: '销售', rows: [['姓名', '金额'], ['张三', 1200]] }
        ]);
        const { blocks } = await xlsxToBlocks(bytes);
        expect(blocks[0]).toMatchObject({ type: 'heading', level: 1, text: '销售' });
        const table = blocks[1] as Extract<DocxBlock, { type: 'table' }>;
        expect(table.type).toBe('table');
        expect(table.rows.length).toBe(2);
        expect(table.header).toBe(true);
        expect(cellText(table.rows[0][0])).toBe('姓名');
        expect(cellText(table.rows[1][0])).toBe('张三');
        expect(cellText(table.rows[1][1])).toBe('1200');
    });

    it('多张工作表依次输出', async () => {
        const bytes = await jsonToXlsx([
            { name: '表一', rows: [['A']] },
            { name: '表二', rows: [['B']] }
        ]);
        const { blocks } = await xlsxToBlocks(bytes);
        const headings = blocks.filter((b) => b.type === 'heading');
        expect(headings.length).toBe(2);
        expect(headings.map((h) => (h as { text?: string }).text)).toEqual(['表一', '表二']);
    });

    it('合并单元格还原为 colspan，被覆盖的位置不产单元格', async () => {
        const bytes = await jsonToXlsx([
            { name: 'S', rows: [['标题', '', ''], ['a', 'b', 'c']], merges: ['A1:C1'] }
        ]);
        const { blocks } = await xlsxToBlocks(bytes);
        const table = blocks.find((b) => b.type === 'table') as Extract<DocxBlock, { type: 'table' }>;
        expect(table.rows[0].length).toBe(1);
        expect((table.rows[0][0] as { colspan?: number }).colspan).toBe(3);
        expect(table.rows[1].length).toBe(3);
    });

    it('空行默认跳过', async () => {
        const bytes = await jsonToXlsx([
            { name: 'S', rows: [['A'], [''], ['B']] }
        ]);
        const { blocks } = await xlsxToBlocks(bytes);
        const table = blocks.find((b) => b.type === 'table') as Extract<DocxBlock, { type: 'table' }>;
        expect(table.rows.length).toBe(2);
    });
});

describe('convert: xlsx 统一门面', () => {
    const build = () =>
        jsonToXlsx([
            { name: 'Sheet1', rows: [['列A', '列B'], ['值1', '值2']] }
        ]);

    it('xlsx -> markdown 输出标题与表格', async () => {
        const md = (await convert(await build(), 'md', { from: 'xlsx' })) as string;
        expect(md).toContain('# Sheet1');
        expect(md).toContain('| 列A | 列B |');
        expect(md).toContain('值1');
    });

    it('xlsx -> html 输出 table', async () => {
        const html = (await convert(await build(), 'html', { from: 'xlsx' })) as string;
        expect(html).toContain('<table');
        expect(html).toContain('列A');
    });

    it('xlsx -> docx 返回 DOCX 字节', async () => {
        const bytes = (await convert(await build(), 'docx', { from: 'xlsx' })) as Uint8Array;
        expect(bytes[0]).toBe(0x50); // P
        expect(bytes[1]).toBe(0x4b); // K
    });

    it('xlsx -> pdf 返回 PDF 字节', async () => {
        const bytes = (await convert(await build(), 'pdf', { from: 'xlsx' })) as Uint8Array;
        expect(bytes[0]).toBe(0x25); // %
    });
});

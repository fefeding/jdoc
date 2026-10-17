import { describe, expect, it } from 'vitest';
import { jsonToXlsx, sheetsToHtml, xlsxToHtml, xlsxToJson, xlsxToMatrix } from '../src/xlsx/index';
import type { XlsxCellObject, XlsxWorkbook } from '../src/xlsx/types';

/** 取出第一个有值的单元格对象 */
function firstCellObject(cell: unknown): XlsxCellObject {
    return cell as XlsxCellObject;
}

const sample: XlsxWorkbook = {
    metadata: { creator: 'jdoc', title: '订单表' },
    sheets: [
        {
            name: '订单',
            columns: [{ width: 20 }, { width: 12 }, { width: 30 }],
            frozen: { rows: 1 },
            rows: [
                ['姓名', '金额', '日期'],
                [
                    '张三',
                    { value: 1200.5, style: { numFmt: '#,##0.00', align: 'right' } },
                    { value: '备注：{ text: "含特殊字符 <>&" }' }
                ],
                [
                    { value: '李四', style: { font: { bold: true, color: '#FF0000' }, fill: '#FFFF00' } },
                    'xlsxToText',
                    { formula: 'B2*2', result: 2401 }
                ],
                [{ value: '链接', hyperlink: 'https://example.com', style: { font: { color: '#0563C1' } } }, '', '']
            ],
            merges: ['A1:C1']
        },
        { name: '空表', rows: [['x']] }
    ]
};

describe('xlsx：JSON -> XLSX -> JSON 往返', () => {
    it('生成的文件是合法 ZIP 容器', async () => {
        const bytes = await jsonToXlsx(sample);
        expect(bytes.byteLength).toBeGreaterThan(1000);
        expect(bytes[0]).toBe(0x50);
        expect(bytes[1]).toBe(0x4b);
    });

    it('保留多个工作表及其顺序', async () => {
        const { sheets } = await xlsxToJson(await jsonToXlsx(sample));
        expect(sheets.length).toBe(2);
        expect(sheets[0].name).toBe('订单');
        expect(sheets[1].name).toBe('空表');
    });

    it('保留单元格值与特殊字符', async () => {
        const { sheets } = await xlsxToJson(await jsonToXlsx(sample));
        const rows = sheets[0].rows;
        const flat = JSON.stringify(rows);
        expect(flat).toContain('张三');
        expect(flat).toContain('李四');
        expect(flat).toContain('含特殊字符 <>&');
    });

    it('保留数值与数字格式', async () => {
        const { sheets } = await xlsxToJson(await jsonToXlsx(sample));
        const cell = firstCellObject(sheets[0].rows[1].cells[1]);
        expect(typeof cell.value).toBe('number');
        expect(cell.style?.numFmt).toBe('#,##0.00');
        expect(cell.style?.align).toBe('right');
    });

    it('保留字体样式与填充色', async () => {
        const { sheets } = await xlsxToJson(await jsonToXlsx(sample));
        const cell = firstCellObject(sheets[0].rows[2].cells[0]);
        expect(cell.style?.font?.bold).toBe(true);
        expect(String(cell.style?.font?.color).toUpperCase()).toContain('FF0000');
        expect(String(cell.style?.fill).toUpperCase()).toContain('FFFF00');
    });

    it('保留公式', async () => {
        const { sheets } = await xlsxToJson(await jsonToXlsx(sample));
        const cell = firstCellObject(sheets[0].rows[2].cells[2]);
        expect(cell.formula).toContain('B2');
    });

    it('保留超链接', async () => {
        const { sheets } = await xlsxToJson(await jsonToXlsx(sample));
        const cell = firstCellObject(sheets[0].rows[3].cells[0]);
        expect(cell.hyperlink).toBe('https://example.com');
    });

    it('保留合并区域与列宽', async () => {
        const { sheets } = await xlsxToJson(await jsonToXlsx(sample));
        expect(sheets[0].merges?.length).toBeGreaterThan(0);
        expect(sheets[0].columns?.[0]?.width).toBe(20);
    });

    it('保留元数据', async () => {
        const { metadata } = await xlsxToJson(await jsonToXlsx(sample));
        expect(metadata?.creator).toBe('jdoc');
    });

    it('skipStyles 时仍能读到值但体积更小', async () => {
        const withStyle = await xlsxToJson(await jsonToXlsx(sample));
        const without = await xlsxToJson(await jsonToXlsx(sample), { skipStyles: true });
        expect(JSON.stringify(without.sheets).length).toBeLessThan(JSON.stringify(withStyle.sheets).length);
        expect(JSON.stringify(without.sheets)).toContain('张三');
    });

    it('按表名过滤', async () => {
        const { sheets } = await xlsxToJson(await jsonToXlsx(sample), { sheets: ['空表'] });
        expect(sheets.length).toBe(1);
        expect(sheets[0].name).toBe('空表');
    });
});

describe('xlsx：矩阵与 HTML 输出', () => {
    it('xlsxToMatrix 取首行作表头', async () => {
        // 故意用不含合并区的表：合并区会让该行各列都返回主单元格的值
        const plain = {
            sheets: [{ name: '数据', rows: [['姓名', '金额'], ['张三', 1200], ['李四', 900]] }]
        };
        const { headers, rows } = await xlsxToMatrix(await jsonToXlsx(plain), '数据');
        expect(headers).toEqual(['姓名', '金额']);
        expect(rows.length).toBe(2);
        expect(rows[0][0]).toBe('张三');
        expect(rows[0][1]).toBe(1200);
    });

    it('合并区内各列返回主单元格的值（Excel 原生语义）', async () => {
        const { headers } = await xlsxToMatrix(await jsonToXlsx(sample), '订单');
        // sample 的 A1:C1 是合并区
        expect(headers).toEqual(['姓名', '姓名', '姓名']);
    });

    it('输出 HTML 表格并正确转义', async () => {
        const { html } = await xlsxToHtml(await jsonToXlsx(sample));
        expect(html).toContain('<table');
        expect(html).toContain('张三');
        expect(html).toContain('&lt;&gt;&amp;');
        expect(html).toContain('订单');
    });

    it('从 JSON 直接渲染 HTML', () => {
        const html = sheetsToHtml(sample);
        expect(html).toContain('张三');
    });
});

describe('xlsx：稀疏行写法', () => {
    it('对象式行支持列号与 A1 记号', async () => {
        const bytes = await jsonToXlsx([
            { name: 'S', rows: [{ 1: 'A列', 3: 'C列' }, { B2: 'B列定位' }] as never }
        ]);
        const { sheets } = await xlsxToJson(bytes);
        const flat = JSON.stringify(sheets[0].rows);
        expect(flat).toContain('A列');
        expect(flat).toContain('C列');
    });
});

describe('xlsx：异常处理', () => {
    it('缺少 sheets 时给出明确错误', async () => {
        await expect(jsonToXlsx({} as never)).rejects.toThrow(/sheets/);
    });

    it('解析非法数据抛带格式标识的错误', async () => {
        await expect(xlsxToJson(new Uint8Array([1, 2, 3]))).rejects.toThrow(/xlsx/);
    });
});

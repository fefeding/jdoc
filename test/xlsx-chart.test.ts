import { describe, expect, it } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { jsonToXlsx, xlsxToJson, xlsxChartToSpec } from '../src/xlsx/index';
import { buildChartXml } from '../src/xlsx/chart-write';
import { attr, findAll, parseXml } from '../src/core/xml';
import { evaluateFormula, extractFormulaRefs, type FormulaContext } from '../src/xlsx/formula';
import type { XlsxCellObject, XlsxRow, XlsxSheet } from '../src/xlsx/types';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const EXAMPLE = path.resolve(__dirname, '..', 'examples', '销售分析表 副本.xlsx');

function readExample(): Uint8Array {
    return new Uint8Array(fs.readFileSync(EXAMPLE));
}

/** 用工作簿搭一个公式求值上下文 */
function contextOf(sheets: XlsxSheet[], current: string): FormulaContext {
    return {
        sheetName: current,
        hasSheet: (name) => sheets.some((s) => s.name === name),
        read(sheetName, row, col) {
            const sheet = sheets.find((s) => s.name === sheetName);
            const r = (sheet?.rows[row] as XlsxRow | undefined)?.cells;
            const cell = r ? (r[col] as XlsxCellObject | null) : null;
            if (!cell) return { value: null, text: '' };
            if (cell.formula !== undefined) {
                return { value: cell.result ?? null, text: cell.result != null ? String(cell.result) : '' };
            }
            if (cell.value !== undefined) return { value: cell.value as never, text: String(cell.value ?? '') };
            if (cell.richText?.length) {
                const text = cell.richText.map((p) => p.text || '').join('');
                return { value: text, text };
            }
            if (cell.text !== undefined) return { value: cell.text, text: cell.text };
            return { value: null, text: '' };
        }
    };
}

describe('xlsx 公式引擎', () => {
    it('中文表名的跨表引用可求值', () => {
        const ctx: FormulaContext = {
            sheetName: '汇总看板',
            hasSheet: (n) => n === '销售数据' || n === '汇总看板',
            read(sheetName, row, col) {
                if (sheetName === '销售数据') {
                    if (col === 6 && row === 2) return { value: 100 };
                    if (col === 2 && row === 2) return { value: '华东' };
                    return { value: null, text: '' };
                }
                return { value: null, text: '' };
            }
        };
        expect(evaluateFormula('SUM(销售数据!G3:G32)', ctx)).toBe(100);
        expect(evaluateFormula('SUMIF(销售数据!C3:C32,"华东",销售数据!G3:G32)', ctx)).toBe(100);
    });

    it('区域求和与跨表求和一致', () => {
        const ctx: FormulaContext = {
            sheetName: 'S',
            read: (_s, row, col) => ({ value: row * 10 + col })
        };
        expect(evaluateFormula('SUM(A1:B2)', ctx)).toBe(0 + 1 + 10 + 11);
    });

    it('TEXT 按格式输出日期', () => {
        const ctx: FormulaContext = { sheetName: 'S', read: () => ({ value: 46023 }) };
        expect(evaluateFormula('TEXT(46023,"yyyy-mm")', ctx)).toBe('2026-01');
    });

    it('提取引用（含跨表）', () => {
        const refs = extractFormulaRefs('SUMIF(销售数据!C3:C32,"华东",销售数据!G3:G32)');
        expect(refs.length).toBe(2);
        expect(refs[0].sheet).toBe('销售数据');
        expect(refs[0].ref).toBe('C3:C32');
    });

    it('未知函数抛错（调用方可保留原值）', () => {
        const ctx: FormulaContext = { sheetName: 'S', read: () => ({ value: 1 }) };
        expect(() => evaluateFormula('NOSUCHFN(1)', ctx)).toThrow();
    });
});

describe('xlsx 图表解析', () => {
    it('示例文件解析出 3 个图表', async () => {
        const { sheets } = await xlsxToJson(readExample());
        const board = sheets.find((s) => s.name === '汇总看板');
        expect(board?.charts?.length).toBe(3);
        const types = (board?.charts || []).map((c) => c.type).sort();
        expect(types).toEqual(['column', 'line', 'pie']);
    });

    it('图表带有标题、系列与锚点', async () => {
        const { sheets } = await xlsxToJson(readExample());
        const board = sheets.find((s) => s.name === '汇总看板');
        const chart = board?.charts?.[0];
        expect(chart?.title).toBe('各区域销售额对比');
        expect(chart?.anchor.from.row).toBe(23);
        expect(chart?.series[0].valRef).toBe("'汇总看板'!$B$8:$B$10");
        // 缓存值由工作簿回查得到
        expect(chart?.series[0].vals?.length).toBe(3);
    });

    it('往返后图表仍在（保存不丢图表）', async () => {
        const parsed = await xlsxToJson(readExample());
        const bytes = await jsonToXlsx({ sheets: parsed.sheets });
        const again = await xlsxToJson(bytes);
        const board = again.sheets.find((s) => s.name === '汇总看板');
        expect(board?.charts?.length).toBe(3);
        expect(board?.charts?.[0].title).toBe('各区域销售额对比');
    });

    it('生成的包结构自洽（XML 合法、关系可解析、内容类型齐全）', async () => {
        const parsed = await xlsxToJson(readExample());
        const bytes = await jsonToXlsx({ sheets: parsed.sheets });
        const zip = await JSZip.loadAsync(bytes);
        const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);

        // 1. 所有 XML 部件都能被解析
        for (const name of names.filter((n) => n.endsWith('.xml') || n.endsWith('.rels'))) {
            const text = await zip.file(name)!.async('string');
            expect(() => parseXml(text), name).not.toThrow();
        }

        // 2. 关系目标都存在
        for (const name of names.filter((n) => n.endsWith('.rels'))) {
            const text = await zip.file(name)!.async('string');
            const base = name.replace(/_rels\/[^/]+$/, '');
            for (const rel of findAll(parseXml(text), 'Relationship')) {
                const target = attr(rel, 'Target');
                if (!target || /^[a-z]+:/i.test(target)) continue; // 外链
                const resolved = resolve(base, target);
                expect(names, `${name} -> ${target}`).toContain(resolved);
            }
        }

        // 3. 内容类型覆盖图表与绘图部件
        const ct = await zip.file('[Content_Types].xml')!.async('string');
        expect(ct).toContain('/xl/charts/chart1.xml');
        expect(ct).toContain('/xl/drawings/drawing1.xml');
        // 4. 工作表里有 drawing 引用
        const sheet = await zip.file('xl/worksheets/sheet2.xml')!.async('string');
        expect(sheet).toMatch(/<drawing r:id="rId\d+"\/>/);
    });
});

/** 相对路径解析（rels 目标 -> 包内路径） */
function resolve(base: string, target: string): string {
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/').filter(Boolean);
    for (const seg of target.split('/')) {
        if (seg === '.' || seg === '') continue;
        if (seg === '..') parts.pop();
        else parts.push(seg);
    }
    return parts.join('/');
}

describe('xlsx 图表能力对齐（与 PPTX 共用渲染内核）', () => {
    const anchor = { from: { col: 0, row: 0 }, to: { col: 6, row: 12 } };

    it('写回：雷达图生成 radarChart 部件', () => {
        const xml = buildChartXml({ type: 'radar', anchor, series: [{ name: 's', vals: [1, 2] }] }, 1);
        expect(xml).toContain('<c:radarChart>');
        expect(xml).toContain('<c:catAx>');
    });

    it('写回：散点图用 xVal / yVal（不是 cat / val）', () => {
        const xml = buildChartXml(
            { type: 'scatter', anchor, series: [{ name: 's', xVals: [1, 2], vals: [3, 4] }] },
            2
        );
        expect(xml).toContain('<c:xVal>');
        expect(xml).toContain('<c:yVal>');
        expect(xml).not.toContain('<c:cat>');
    });

    it('写回：数据标签 / 环形内径 / 平滑', () => {
        const pie = buildChartXml(
            { type: 'pie', anchor, series: [{ name: 's', vals: [1] }], labels: { percent: true, value: true } },
            3
        );
        expect(pie).toContain('<c:showPercent val="1"/>');
        expect(pie).toContain('<c:showVal val="1"/>');
        const ring = buildChartXml({ type: 'doughnut', anchor, holeSize: 40, series: [{ name: 's', vals: [1] }] }, 4);
        expect(ring).toContain('<c:holeSize val="40"/>');
        const line = buildChartXml({ type: 'line', anchor, smooth: true, series: [{ name: 's', vals: [1] }] }, 5);
        expect(line).toContain('<c:smooth val="1"/>');
    });

    it('往返：雷达图类型不再被降级成折线', async () => {
        const bytes = await jsonToXlsx({
            sheets: [
                {
                    name: 'S',
                    rows: [
                        { index: 1, cells: [{ value: 'A' }, { value: 1 }] },
                        { index: 2, cells: [{ value: 'B' }, { value: 2 }] }
                    ],
                    charts: [
                        {
                            type: 'radar',
                            anchor,
                            title: '雷达',
                            series: [{ name: 's', catRef: 'S!$A$1:$A$2', valRef: 'S!$B$1:$B$2' }]
                        }
                    ]
                }
            ]
        });
        const parsed = await xlsxToJson(bytes);
        const chart = parsed.sheets[0].charts?.[0];
        expect(chart?.type).toBe('radar');
        expect(chart?.title).toBe('雷达');
    });

    it('图表模型可映射为共享渲染模型（ChartSpec）', () => {
        const spec = xlsxChartToSpec(
            {
                type: 'column',
                anchor,
                title: 'T',
                legend: 'r',
                valNumFmt: '¥#,##0',
                series: [{ name: 's', catRef: 'S!$A$1:$A$2', valRef: 'S!$B$1:$B$2' }]
            },
            [{ name: 's', cats: ['A', 'B'], vals: [1, 2] }]
        );
        expect(spec.type).toBe('column');
        expect(spec.title).toBe('T');
        expect(spec.legend).toBe('r');
        expect(spec.numFmt).toBe('¥#,##0');
        expect(spec.categories).toEqual(['A', 'B']);
        expect(spec.series?.[0].values).toEqual([1, 2]);
    });

    it('数据标签缺省按 OOXML 规范视为开启（CT_Boolean 默认 true）', async () => {
        const parsed = await xlsxToJson(readExample());
        const board = parsed.sheets.find((s) => s.name === '汇总看板');
        const pie = board?.charts?.find((c) => c.type === 'pie');
        // 文件里只写了 <c:showPercent val="1"/>，其余开关缺省即开启
        expect(pie?.labels?.percent).toBe(true);
        expect(pie?.labels?.value).toBe(true);
        expect(pie?.labels?.category).toBe(true);
    });
});

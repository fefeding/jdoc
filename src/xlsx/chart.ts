/**
 * XLSX 图表解析
 *
 * ExcelJS 不支持图表，图表部件（`xl/charts/chartN.xml`）与绘图锚点
 * （`xl/drawings/drawingN.xml`）需要直接读 OOXML：
 *
 *   sheet.xml ──rels──> drawing1.xml ──rels──> chart1.xml
 *                        （锚点：位置与尺寸）
 *
 * 解析结果落到 `XlsxSheet.charts`：图表类型、标题、数据系列（引用 + 缓存值）、
 * 锚点（起止单元格 / 尺寸）。缓存值由引用回查工作簿单元格得到，
 * 因此没有 `<c:numCache>` 的文件（很多生成器都不写缓存）也能画出图。
 *
 * @module xlsx/chart
 */

import JSZip from 'jszip';
import { attr, findAll, findFirst, parseXml, textOf, type XmlElement } from '../core/xml';
import { dirOf, parseRels, relsPath, resolveTarget } from '../core/ooxml';
import { parseRangeRef } from './formula';
import type { ChartSpec } from '../chart/svg';
import type { XlsxCellObject, XlsxChart, XlsxChartSeries, XlsxRow, XlsxSheet } from './types';

const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** 读取 ZIP 内文本部件 */
async function readText(zip: JSZip, path: string): Promise<string> {
    const file = zip.file(path);
    if (!file) return '';
    return file.async('string');
}

/** 子元素文本 */
function childText(node: XmlElement | undefined, tag: string): string {
    if (!node) return '';
    return textOf(findFirst(node, tag)).trim();
}

/** 单元格取值（公式取计算结果） */
function cellValueOf(cell: XlsxCellObject | null | undefined): string | number | null {
    if (!cell) return null;
    if (cell.formula !== undefined) {
        if (cell.result === undefined || cell.result === null) return null;
        return cell.result as string | number;
    }
    if (cell.value !== undefined && cell.value !== null) {
        const v = cell.value as unknown;
        // 日期按序列号参与图表数值
        if (v instanceof Date) return Math.round((v.getTime() - Date.UTC(1899, 11, 30)) / 86400000);
        return v as string | number;
    }
    if (cell.richText?.length) return cell.richText.map((p) => p.text || '').join('');
    if (cell.text !== undefined && cell.text !== null) return cell.text;
    return null;
}

/** 单元格显示文本 */
function cellTextOf(cell: XlsxCellObject | null | undefined): string {
    const v = cellValueOf(cell);
    if (v === null) return '';
    return String(v);
}

/**
 * 按引用回查工作簿取值
 *
 * @param ref - 引用，如 `'汇总看板'!$B$8:$B$10`
 * @param sheets - 工作表列表
 * @param fallbackSheet - 引用未带表名时的默认表
 */
export function resolveRefValues(
    ref: string,
    sheets: XlsxSheet[],
    fallbackSheet?: string
): (string | number | null)[] {
    const parsed = parseRangeRef(ref);
    if (!parsed) return [];
    const sheet = parsed.sheet
        ? sheets.find((s) => s.name === parsed.sheet)
        : sheets.find((s) => s.name === fallbackSheet);
    if (!sheet) return [];
    const out: (string | number | null)[] = [];
    for (let r = parsed.r1; r <= parsed.r2; r++) {
        const row = (sheet.rows[r] as XlsxRow | undefined) ?? undefined;
        for (let c = parsed.c1; c <= parsed.c2; c++) {
            const cell = row?.cells ? (row.cells[c] as XlsxCellObject | null) : null;
            out.push(cellValueOf(cell));
        }
    }
    return out;
}

/** 解析图表 XML */
function parseChartXml(xml: string): Omit<XlsxChart, 'anchor'> | null {
    if (!xml) return null;
    const root = findFirst(parseXml(xml), 'chartSpace');
    if (!root) return null;

    const plotArea = findFirst(root, 'plotArea');
    const chartNode = findFirst(root, 'chart');

    // 图表类型
    let type: XlsxChart['type'] | null = null;
    let barDir: string | undefined;
    let grouping: string | undefined;
    if (plotArea) {
        const kinds: [string, XlsxChart['type']][] = [
            ['barChart', 'bar'],
            ['pieChart', 'pie'],
            ['lineChart', 'line'],
            ['areaChart', 'area'],
            ['doughnutChart', 'doughnut'],
            ['scatterChart', 'scatter'],
            ['radarChart', 'radar']
        ];
        for (const [tag, kind] of kinds) {
            const node = findFirst(plotArea, tag);
            if (!node) continue;
            type = kind;
            if (tag === 'barChart') {
                const dir = childText(node, 'barDir');
                barDir = dir || 'col';
            }
            const g = childText(node, 'grouping');
            if (g) grouping = g;
            break;
        }
    }
    if (!type) return null;
    if (type === 'bar' && barDir === 'col') type = 'column';

    // 标题
    let title: string | undefined;
    const titleNode = findFirst(root, 'title');
    if (titleNode) {
        const rich = findAll(titleNode, 't');
        if (rich.length) {
            title = rich.map((t) => textOf(t).trim()).filter(Boolean).join('');
        } else {
            const f = childText(titleNode, 'f');
            if (f) title = f;
        }
    }

    // 数据系列
    const series: XlsxChartSeries[] = [];
    for (const ser of plotArea ? findAll(plotArea, 'ser') : []) {
        const item: XlsxChartSeries = {};
        const tx = findFirst(ser, 'tx');
        if (tx) {
            const nameRef = childText(tx, 'f');
            if (nameRef) item.nameRef = nameRef;
            const v = childText(tx, 'v');
            if (v) item.name = v;
        }
        const cat = findFirst(ser, 'cat');
        if (cat) {
            const ref = childText(cat, 'f');
            if (ref) item.catRef = ref;
        }
        const val = findFirst(ser, 'val');
        if (val) {
            const ref = childText(val, 'f');
            if (ref) item.valRef = ref;
        }
        // 散点图：x 值走 c:xVal（不是 c:cat）
        const xVal = findFirst(ser, 'xVal');
        if (xVal) {
            const ref = childText(xVal, 'f');
            if (ref) item.xRef = ref;
        }
        // 逐点颜色：c:dPt/c:spPr/a:solidFill/a:srgbClr@val
        const pointColors: string[] = [];
        for (const dPt of findAll(ser, 'dPt')) {
            const idx = parseInt(childText(dPt, 'idx') || '-1', 10);
            if (idx < 0) continue;
            const clr = findFirst(dPt, 'srgbClr');
            const val = clr ? attr(clr, 'val') : '';
            if (val) pointColors[idx] = `#${val.replace('#', '').toUpperCase()}`;
        }
        if (pointColors.length) item.pointColors = pointColors;
        series.push(item);
    }

    // 图例
    let legend: XlsxChart['legend'] = 'none';
    const legendNode = findFirst(root, 'legend');
    if (legendNode) {
        const pos = childText(legendNode, 'legendPos');
        legend = (['l', 'r', 't', 'b'] as const).includes(pos as 'l')
            ? (pos as 'l' | 'r' | 't' | 'b')
            : 'r';
    }

    // 数据标签：showVal / showPercent / showCatName
    const labels: { value: boolean; percent: boolean; category: boolean } = {
        value: false,
        percent: false,
        category: false
    };
    for (const dLbls of findAll(root, 'dLbls')) {
        // CT_Boolean 的 val 属性默认是 true：dLbls 里没写的开关 = 开启。
        // 只写 <c:showPercent val="1"/> 的文件（不少生成器如此），在 Excel 里
        // 实际是「分类名 + 数值 + 百分比」全开，按缺省 true 才对得上。
        const flag = (tag: string): boolean => {
            const node = findFirst(dLbls, tag);
            if (!node) return true;
            const v = attr(node, 'val');
            if (v !== undefined) return v === '1' || v === 'true';
            const t = textOf(node).trim();
            return t === '' || t === '1' || t === 'true';
        };
        if (flag('showVal')) labels.value = true;
        if (flag('showPercent')) labels.percent = true;
        if (flag('showCatName')) labels.category = true;
    }
    const showPercent = labels.percent;

    // 环形图内径
    let holeSize: number | undefined;
    const holeNode = findFirst(root, 'holeSize');
    if (holeNode) {
        const n = parseInt(textOf(holeNode).trim() || attr(holeNode, 'val') || '', 10);
        if (!Number.isNaN(n) && n > 0) holeSize = n;
    }

    // 折线标记 / 平滑 / 饼图逐点配色
    let marker: boolean | undefined;
    let smooth: boolean | undefined;
    let varyColors: boolean | undefined;
    for (const node of plotArea ? findAll(plotArea, 'marker') : []) {
        const sym = childText(node, 'symbol');
        marker = sym ? sym !== 'none' : attr(node, 'val') !== '0';
        break;
    }
    for (const node of plotArea ? findAll(plotArea, 'smooth') : []) {
        smooth = attr(node, 'val') !== '0' && textOf(node).trim() !== '0';
        break;
    }
    for (const node of plotArea ? findAll(plotArea, 'varyColors') : []) {
        varyColors = attr(node, 'val') === '1' || textOf(node).trim() === '1';
        break;
    }

    // 数值轴数字格式
    let valNumFmt: string | undefined;
    if (plotArea) {
        const valAx = findFirst(plotArea, 'valAx');
        const numFmt = valAx ? findFirst(valAx, 'numFmt') : undefined;
        const code = numFmt ? attr(numFmt, 'formatCode') : undefined;
        if (code) valNumFmt = code;
    }

    const result: Omit<XlsxChart, 'anchor'> = { type, series };
    if (title) result.title = title;
    if (legend !== 'none') result.legend = legend;
    if (labels.value || labels.percent || labels.category) result.labels = labels;
    if (showPercent) result.showPercent = true;
    if (valNumFmt) result.valNumFmt = valNumFmt;
    if (grouping) {
        result.grouping = grouping as XlsxChart['grouping'];
    }
    if (holeSize !== undefined) result.holeSize = holeSize;
    if (marker !== undefined) result.marker = marker;
    if (smooth !== undefined) result.smooth = smooth;
    if (varyColors !== undefined) result.varyColors = varyColors;
    return result;
}

/** 解析绘图锚点 */
function parseDrawingAnchors(xml: string): {
    from: { col: number; row: number; colOff: number; rowOff: number };
    to?: { col: number; row: number; colOff: number; rowOff: number };
    ext?: { cx: number; cy: number };
    rid: string;
}[] {
    const out: ReturnType<typeof parseDrawingAnchors> = [];
    const nodes = parseXml(xml);
    for (const tag of ['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor']) {
        for (const anchor of findAll(nodes, tag)) {
            const fromNode = findFirst(anchor, 'from');
            if (!fromNode) continue;
            const read = (node: XmlElement | undefined): { col: number; row: number; colOff: number; rowOff: number } => ({
                col: parseInt(childText(node, 'col') || '0', 10) || 0,
                row: parseInt(childText(node, 'row') || '0', 10) || 0,
                colOff: parseInt(childText(node, 'colOff') || '0', 10) || 0,
                rowOff: parseInt(childText(node, 'rowOff') || '0', 10) || 0
            });
            const chartNode = findAll(anchor, 'chart')[0];
            const rid = chartNode ? attr(chartNode, 'id') || attr(chartNode, `${NS_REL}id`) : undefined;
            if (!rid) continue;
            const item: ReturnType<typeof parseDrawingAnchors>[number] = {
                from: read(fromNode),
                rid
            };
            const toNode = findFirst(anchor, 'to');
            if (toNode) item.to = read(toNode);
            const extNode = findFirst(anchor, 'ext');
            if (extNode) {
                // <xdr:ext cx="5040000" cy="2880000"/>：尺寸在属性上
                item.ext = {
                    cx: parseInt(attr(extNode, 'cx') || '0', 10) || 0,
                    cy: parseInt(attr(extNode, 'cy') || '0', 10) || 0
                };
            }
            out.push(item);
        }
    }
    return out;
}

/**
 * 读取工作簿中的图表
 *
 * @param zip - XLSX 的 ZIP 容器
 * @param sheets - 已解析的工作表（用于回查图表引用的数据）
 * @returns 表名 -> 图表列表
 */
export async function readWorkbookCharts(
    zip: JSZip,
    sheets: XlsxSheet[]
): Promise<Map<string, XlsxChart[]>> {
    const result = new Map<string, XlsxChart[]>();
    if (!sheets.length) return result;

    const wbXml = await readText(zip, 'xl/workbook.xml');
    const wbRelsXml = await readText(zip, 'xl/_rels/workbook.xml.rels');
    if (!wbXml) return result;
    const wbRels = parseRels(wbRelsXml);

    // 表名 -> 工作表部件路径
    const sheetParts: { name: string; part: string }[] = [];
    for (const sheetNode of findAll(parseXml(wbXml), 'sheet')) {
        const name = attr(sheetNode, 'name');
        const rid = attr(sheetNode, 'id');
        if (!name || !rid) continue;
        const rel = wbRels.get(rid);
        if (!rel) continue;
        sheetParts.push({ name, part: resolveTarget('xl', rel.target) });
    }

    for (const { name, part } of sheetParts) {
        const sheetRelsXml = await readText(zip, relsPath(part));
        if (!sheetRelsXml) continue;
        const sheetRels = parseRels(sheetRelsXml);
        const drawings: string[] = [];
        for (const [, rel] of sheetRels) {
            if (rel.type.endsWith('/drawing')) {
                drawings.push(resolveTarget(dirOf(part), rel.target));
            }
        }
        if (!drawings.length) continue;

        const charts: XlsxChart[] = [];
        for (const drawing of drawings) {
            const drawingXml = await readText(zip, drawing);
            if (!drawingXml) continue;
            const drawingRels = parseRels(await readText(zip, relsPath(drawing)));
            for (const anchor of parseDrawingAnchors(drawingXml)) {
                const rel = drawingRels.get(anchor.rid);
                if (!rel) continue;
                const chartPart = resolveTarget(dirOf(drawing), rel.target);
                const chartXml = await readText(zip, chartPart);
                const parsed = parseChartXml(chartXml);
                if (!parsed) continue;

                const chart: XlsxChart = {
                    ...parsed,
                    anchor: {
                        from: anchor.from,
                        ...(anchor.to ? { to: anchor.to } : {}),
                        ...(anchor.ext ? { ext: anchor.ext } : {})
                    }
                };

                // 回查系列数据（引用 -> 缓存值），并补系列名
                for (const series of chart.series) {
                    if (series.nameRef && !series.name) {
                        const values = resolveRefValues(series.nameRef, sheets, name);
                        if (values.length) series.name = String(values[0] ?? '');
                    }
                    if (series.catRef) {
                        const cats = resolveRefValues(series.catRef, sheets, name);
                        if (cats.length) series.cats = cats;
                    }
                    if (series.valRef) {
                        const vals = resolveRefValues(series.valRef, sheets, name);
                        if (vals.length) series.vals = vals.map((v) => (typeof v === 'number' ? v : v === null ? null : Number(v)));
                    }
                    if (series.xRef) {
                        const xs = resolveRefValues(series.xRef, sheets, name);
                        if (xs.length) series.xVals = xs.map((v) => (typeof v === 'number' ? v : v === null ? null : Number(v)));
                    }
                }

                charts.push(chart);
            }
        }
        if (charts.length) result.set(name, charts);
    }

    return result;
}

/**
 * 图表系列数据（调用方按引用回查后的实时值）
 *
 * 与 `XlsxChart.series` 分开：图表定义里的 `vals` 是文件里的缓存值，
 * 编辑器要的是「当前单元格里的值」，两者可能已经不一致。
 */
export interface XlsxChartSeriesData {
    /** 系列名 */
    name?: string;
    /** 分类轴取值 */
    cats: (string | number | null)[];
    /** 数值轴取值 */
    vals: (number | null)[];
    /** 散点图 x 值 */
    xVals?: (number | null)[];
    /** 系列色 */
    color?: string;
}

/**
 * 把 XLSX 图表转成共享渲染模块的 `ChartSpec`
 *
 * 绘图交给 `chart/svg`（与 PPTX 共用），这里只做格式相关的字段映射。
 *
 * @param chart - 图表定义
 * @param data - 各系列的实时取值
 */
export function xlsxChartToSpec(chart: XlsxChart, data: XlsxChartSeriesData[]): ChartSpec {
    const series = (chart.series || []).map((s, i) => {
        const d = data[i];
        return {
            name: d?.name ?? s.name ?? '',
            values: d?.vals ?? s.vals ?? [],
            xValues: d?.xVals ?? s.xVals,
            color: d?.color ?? s.color,
            pointColors: s.pointColors
        };
    });
    const spec: ChartSpec = {
        type: chart.type,
        categories: data[0]?.cats ?? chart.series?.[0]?.cats ?? [],
        series
    };
    if (chart.title) spec.title = chart.title;
    if (chart.legend) spec.legend = chart.legend;
    const labels = chart.labels ?? (chart.showPercent ? { percent: true } : undefined);
    if (labels) spec.labels = labels;
    if (chart.grouping && chart.grouping !== 'standard') spec.grouping = chart.grouping;
    if (chart.holeSize !== undefined) spec.holeSize = chart.holeSize;
    if (chart.marker !== undefined) spec.marker = chart.marker;
    if (chart.smooth !== undefined) spec.smooth = chart.smooth;
    if (chart.varyColors !== undefined) spec.varyColors = chart.varyColors;
    if (chart.valNumFmt) spec.numFmt = chart.valNumFmt;
    return spec;
}

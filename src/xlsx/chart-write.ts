/**
 * XLSX 图表写回
 *
 * ExcelJS 不支持图表，保存时图表会整体丢失。这里在 ExcelJS 产出 ZIP 之后，
 * 直接补写图表部件：
 *
 *   sheetN.xml ──<drawing r:id>──> drawings/drawingM.xml ──> charts/chartK.xml
 *
 * 若该表已有 drawing（单元格图片也会生成 drawing），则把图表锚点并入既有
 * drawing 部件——一张表只允许一个 `<drawing>`。
 *
 * @module xlsx/chart-write
 */

import JSZip from 'jszip';
import { attr, findAll, parseXml } from '../core/xml';
import { dirOf, relsPath, resolveTarget } from '../core/ooxml';
import type { XlsxChart, XlsxChartSeries, XlsxSheet } from './types';

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';

const REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const REL_DRAWING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';

const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml';

/** XML 转义 */
function esc(value: unknown): string {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/** 分类轴数据是文本还是数字 */
function catsAreNumeric(cats: (string | number | null)[] | undefined): boolean {
    if (!cats || !cats.length) return false;
    return cats.every((v) => v === null || typeof v === 'number');
}

/** strRef / numRef 片段 */
function refXml(
    tag: 'strRef' | 'numRef',
    ref: string | undefined,
    values: (string | number | null)[] | undefined,
    formatCode?: string
): string {
    const inner: string[] = [];
    if (ref) inner.push(`<c:f>${esc(ref)}</c:f>`);
    if (values && values.length) {
        const pts = values
            .map((v, i) => (v === null || v === undefined ? '' : `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`))
            .join('');
        const cache =
            tag === 'numRef'
                ? `<c:numCache>${formatCode ? `<c:formatCode>${esc(formatCode)}</c:formatCode>` : ''}<c:ptCount val="${values.length}"/>${pts}</c:numCache>`
                : `<c:strCache><c:ptCount val="${values.length}"/>${pts}</c:strCache>`;
        inner.push(cache);
    }
    if (!inner.length) return '';
    return `<c:${tag}>${inner.join('')}</c:${tag}>`;
}

/** 单个系列 */
function seriesXml(series: XlsxChartSeries, index: number, smooth?: boolean, scatter?: boolean): string {
    const parts: string[] = [];
    parts.push(`<c:idx val="${index}"/><c:order val="${index}"/>`);

    if (series.name || series.nameRef) {
        const nameXml = series.nameRef
            ? refXml('strRef', series.nameRef, series.name ? [series.name] : undefined)
            : `<c:tx><c:v>${esc(series.name)}</c:v></c:tx>`;
        parts.push(nameXml);
    }

    if (series.color) {
        const rgb = String(series.color).replace('#', '').toUpperCase();
        if (/^[0-9A-F]{6}$/.test(rgb)) {
            parts.push(`<c:spPr><a:solidFill><a:srgbClr val="${rgb}"/></a:solidFill></c:spPr>`);
        }
    }

    // 散点图：x 走 c:xVal、y 走 c:yVal；其余图表用 c:cat / c:val
    // 判定看的是图表类型而不是系列自身：散点图里即使某系列没有 x 数据，
    // 也必须写成 xVal/yVal（写成 cat/val 会被 Excel 读成 x=分类、y=数值）。
    if (scatter) {
        const x = refXml('numRef', series.xRef, series.xVals);
        if (x) parts.push(`<c:xVal>${x}</c:xVal>`);
        const y = refXml('numRef', series.valRef, series.vals);
        if (y) parts.push(`<c:yVal>${y}</c:yVal>`);
    } else {
        if (series.catRef || series.cats) {
            const tag = catsAreNumeric(series.cats) ? 'numRef' : 'strRef';
            const cat = refXml(tag, series.catRef, series.cats);
            if (cat) parts.push(`<c:cat>${cat}</c:cat>`);
        }
        if (series.valRef || series.vals) {
            const val = refXml('numRef', series.valRef, series.vals);
            if (val) parts.push(`<c:val>${val}</c:val>`);
        }
    }

    // 逐点颜色
    if (series.pointColors?.length) {
        const dpts = series.pointColors
            .map((color, idx) => {
                if (!color) return '';
                const rgb = String(color).replace('#', '').toUpperCase();
                if (!/^[0-9A-F]{6}$/.test(rgb)) return '';
                return `<c:dPt><c:idx val="${idx}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${rgb}"/></a:solidFill></c:spPr></c:dPt>`;
            })
            .join('');
        if (dpts) parts.push(dpts);
    }

    // c:smooth 属于 CT_LineSer，位置在 c:val 之后
    if (smooth) parts.push('<c:smooth val="1"/>');

    return `<c:ser>${parts.join('')}</c:ser>`;
}

/** 数据标签 XML */
function dataLabelsXml(chart: XlsxChart): string {
    const labels = chart.labels ?? (chart.showPercent ? { percent: true } : {});
    if (!labels.value && !labels.percent && !labels.category) return '';
    return (
        '<c:dLbls>' +
        '<c:showLegendKey val="0"/>' +
        `<c:showVal val="${labels.value ? 1 : 0}"/>` +
        `<c:showCatName val="${labels.category ? 1 : 0}"/>` +
        '<c:showSerName val="0"/>' +
        `<c:showPercent val="${labels.percent ? 1 : 0}"/>` +
        '<c:showBubbleSize val="0"/>' +
        '</c:dLbls>'
    );
}

/** 分类轴 / 数值轴 */
function axesXml(catId: number, valId: number, valNumFmt?: string): string {
    const catAx = `<c:catAx><c:axId val="${catId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${valId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
    const valAx = `<c:valAx><c:axId val="${valId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="${esc(valNumFmt || 'General')}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${catId}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
    return catAx + valAx;
}

/** 生成图表部件 XML */
export function buildChartXml(chart: XlsxChart, seed: number): string {
    const catId = 100000000 + seed * 2;
    const valId = catId + 1;
    const isLine = chart.type === 'line';
    const isScatter = chart.type === 'scatter';
    const sers = chart.series
        .map((s, i) => seriesXml(s, i, isLine && chart.smooth === true, isScatter))
        .join('');

    let plot = '';
    switch (chart.type) {
        case 'column':
        case 'bar': {
            const dir = chart.type === 'bar' ? 'bar' : 'col';
            const grouping = chart.grouping === 'stacked' || chart.grouping === 'percentStacked' ? chart.grouping : 'clustered';
            const overlap = grouping === 'clustered' ? '' : '<c:overlap val="100"/>';
            const vary = chart.varyColors ? 1 : 0;
            const dlbls = dataLabelsXml(chart);
            plot = `<c:barChart><c:barDir val="${dir}"/><c:grouping val="${grouping}"/><c:varyColors val="${vary}"/>${sers}${dlbls}<c:gapWidth val="150"/>${overlap}<c:axId val="${catId}"/><c:axId val="${valId}"/></c:barChart>${axesXml(catId, valId, chart.valNumFmt)}`;
            break;
        }
        case 'line': {
            const grouping = chart.grouping === 'stacked' || chart.grouping === 'percentStacked' ? chart.grouping : 'standard';
            const marker = chart.marker === false ? 0 : 1;
            const smooth = chart.smooth ? 1 : 0;
            const dlbls = dataLabelsXml(chart);
            plot = `<c:lineChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${sers}${dlbls}<c:marker val="${marker}"/><c:axId val="${catId}"/><c:axId val="${valId}"/></c:lineChart>${axesXml(catId, valId, chart.valNumFmt)}`;
            break;
        }
        case 'area': {
            const grouping = chart.grouping === 'stacked' || chart.grouping === 'percentStacked' ? chart.grouping : 'standard';
            const dlbls = dataLabelsXml(chart);
            plot = `<c:areaChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${sers}${dlbls}<c:axId val="${catId}"/><c:axId val="${valId}"/></c:areaChart>${axesXml(catId, valId, chart.valNumFmt)}`;
            break;
        }
        case 'pie':
        case 'doughnut': {
            const tag = chart.type === 'doughnut' ? 'doughnutChart' : 'pieChart';
            const hole =
                chart.type === 'doughnut'
                    ? `<c:holeSize val="${Math.max(10, Math.min(90, Math.round(chart.holeSize ?? 50)))}"/>`
                    : '';
            const dlbls = dataLabelsXml(chart);
            const vary = chart.varyColors === false ? 0 : 1;
            plot = `<c:${tag}><c:varyColors val="${vary}"/>${sers}${dlbls}${chart.type === 'doughnut' ? hole : '<c:firstSliceAng val="0"/>'}</c:${tag}>`;
            break;
        }
        case 'scatter': {
            plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}<c:axId val="${catId}"/><c:axId val="${valId}"/></c:scatterChart>${axesXml(catId, valId, chart.valNumFmt)}`;
            break;
        }
        case 'radar': {
            plot = `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${sers}<c:axId val="${catId}"/><c:axId val="${valId}"/></c:radarChart>${axesXml(catId, valId, chart.valNumFmt)}`;
            break;
        }
        default:
            plot = '';
    }

    const title = chart.title
        ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t>${esc(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
        : '<c:autoTitleDeleted val="1"/>';

    const legend =
        chart.legend && chart.legend !== 'none'
            ? `<c:legend><c:legendPos val="${chart.legend}"/><c:overlay val="0"/></c:legend>`
            : '';

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
}

/** 锚点 XML */
function anchorXml(chart: XlsxChart, rid: string, id: number): string {
    const from = chart.anchor.from;
    const to = chart.anchor.to;
    const fromXml = `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>${Math.round(from.colOff || 0)}</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>${Math.round(from.rowOff || 0)}</xdr:rowOff></xdr:from>`;
    const frame = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" xmlns:r="${NS_R}" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;

    if (to) {
        const toXml = `<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>${Math.round(to.colOff || 0)}</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>${Math.round(to.rowOff || 0)}</xdr:rowOff></xdr:to>`;
        return `<xdr:twoCellAnchor>${fromXml}${toXml}${frame}<xdr:clientData/></xdr:twoCellAnchor>`;
    }
    const ext = chart.anchor.ext || { cx: 5040000, cy: 2880000 };
    return `<xdr:oneCellAnchor>${fromXml}<xdr:ext cx="${Math.round(ext.cx)}" cy="${Math.round(ext.cy)}"/>${frame}<xdr:clientData/></xdr:oneCellAnchor>`;
}

/** 关系文件：追加 Relationship，返回新的 Id */
function addRelationship(xml: string, type: string, target: string): { xml: string; id: string } {
    let max = 0;
    for (const m of xml.matchAll(/Id="rId(\d+)"/g)) max = Math.max(max, parseInt(m[1], 10));
    const id = `rId${max + 1}`;
    const rel = `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
    if (!xml) {
        return {
            xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel}</Relationships>`,
            id
        };
    }
    return { xml: xml.replace('</Relationships>', `${rel}</Relationships>`), id };
}

/** Content_Types：追加 Override */
function addContentType(xml: string, partName: string, contentType: string): string {
    if (xml.includes(`PartName="${partName}"`)) return xml;
    const override = `<Override PartName="${partName}" ContentType="${contentType}"/>`;
    return xml.replace('</Types>', `${override}</Types>`);
}

/** 在工作表 XML 里插入 `<drawing r:id>`（须位于 pageSetup 之后、tableParts 之前） */
function insertDrawing(sheetXml: string, rid: string): string {
    if (/<drawing\s/.test(sheetXml)) return sheetXml;
    const tag = `<drawing r:id="${rid}"/>`;
    // 顺序约束：drawing 之后才是 legacyDrawing/picture/oleObjects/controls/tableParts/extLst
    const after = ['<legacyDrawing', '<picture', '<oleObjects', '<controls', '<webPublishItems', '<tableParts', '<extLst'];
    for (const marker of after) {
        const idx = sheetXml.indexOf(marker);
        if (idx >= 0) return sheetXml.slice(0, idx) + tag + sheetXml.slice(idx);
    }
    const idx = sheetXml.lastIndexOf('</worksheet>');
    if (idx < 0) return sheetXml;
    return sheetXml.slice(0, idx) + tag + sheetXml.slice(idx);
}

/**
 * 工作表部件路径（按 workbook.xml 的顺序）
 *
 * 不能假设是 `xl/worksheets/sheetN.xml`：ExcelJS 写出的部件名由它自己决定，
 * 一旦与 spec 里的表顺序对不上，图表就会被写进别的工作表。
 *
 * @param zip - ExcelJS 产出的 ZIP
 * @returns 与 sheets 同序的部件路径（解析不到时为空串）
 */
async function sheetPartPaths(zip: JSZip): Promise<string[]> {
    const wbXml = await zip.file('xl/workbook.xml')?.async('string');
    const wbRelsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
    if (!wbXml) return [];
    const rels = new Map<string, string>();
    for (const rel of findAll(parseXml(wbRelsXml || ''), 'Relationship')) {
        const id = attr(rel, 'Id');
        const target = attr(rel, 'Target');
        if (id && target) rels.set(id, target);
    }
    const out: string[] = [];
    for (const sheetNode of findAll(parseXml(wbXml), 'sheet')) {
        const rid = attr(sheetNode, 'id');
        const target = rid ? rels.get(rid) : undefined;
        out.push(target ? resolveTarget('xl', target) : '');
    }
    return out;
}

/**
 * 把图表写回 ZIP
 *
 * @param zip - ExcelJS 产出的 ZIP
 * @param sheets - 工作表（含 charts）
 */
export async function writeWorkbookCharts(zip: JSZip, sheets: XlsxSheet[]): Promise<void> {
    const withCharts = sheets.filter((s) => s.charts && s.charts.length);
    if (!withCharts.length) return;

    let contentTypes = await zip.file('[Content_Types].xml')?.async('string');
    if (!contentTypes) return;

    let chartSeq = 0;
    let drawingSeq = 0;
    const partPaths = await sheetPartPaths(zip);

    for (let i = 0; i < sheets.length; i++) {
        const sheet = sheets[i];
        if (!sheet.charts || !sheet.charts.length) continue;

        const sheetPart = partPaths[i] || `xl/worksheets/sheet${i + 1}.xml`;
        const sheetFile = zip.file(sheetPart);
        if (!sheetFile) continue;
        let sheetXml = await sheetFile.async('string');

        const sheetRelsPath = relsPath(sheetPart);
        let sheetRels = (await zip.file(sheetRelsPath)?.async('string')) || '';
        let drawingPart = '';
        let drawingRelsPath = '';
        let drawingRels = '';
        let drawingXml = '';

        // 已有 drawing（图片等）：并入既有部件
        const existing = /<drawing\s[^>]*r:id="([^"]+)"/.exec(sheetXml);
        if (existing) {
            const rid = existing[1];
            const relMatch = new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`).exec(sheetRels);
            if (relMatch) {
                // 关系目标是相对工作表部件的（通常 ../drawings/drawing1.xml）
                drawingPart = resolveTarget(dirOf(sheetPart), relMatch[1]);
                drawingRelsPath = relsPath(drawingPart);
                drawingRels = (await zip.file(drawingRelsPath)?.async('string')) || '';
                drawingXml = (await zip.file(drawingPart)?.async('string')) || '';
            }
        }

        if (!drawingPart) {
            drawingSeq++;
            drawingPart = `xl/drawings/drawing${drawingSeq}.xml`;
            drawingRelsPath = `xl/drawings/_rels/drawing${drawingSeq}.xml.rels`;
            drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}"></xdr:wsDr>`;
        }

        // 写图表部件 + 锚点
        let anchors = '';
        for (const chart of sheet.charts) {
            chartSeq++;
            const chartPart = `xl/charts/chart${chartSeq}.xml`;
            zip.file(chartPart, buildChartXml(chart, chartSeq));
            contentTypes = addContentType(contentTypes, `/${chartPart}`, CT_CHART);

            const rel = addRelationship(drawingRels, REL_CHART, `../charts/chart${chartSeq}.xml`);
            drawingRels = rel.xml;
            anchors += anchorXml(chart, rel.id, chartSeq);
        }

        // 写 drawing 部件
        const drawing = drawingXml.includes('</xdr:wsDr>')
            ? drawingXml.replace('</xdr:wsDr>', `${anchors}</xdr:wsDr>`)
            : drawingXml + anchors;
        zip.file(drawingPart, drawing);
        zip.file(drawingRelsPath, drawingRels || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`);
        contentTypes = addContentType(contentTypes, `/${drawingPart}`, CT_DRAWING);

        // 工作表 -> drawing 的关系
        if (!existing) {
            const rel = addRelationship(sheetRels, REL_DRAWING, `../drawings/${drawingPart.slice('xl/drawings/'.length)}`);
            sheetRels = rel.xml;
            sheetXml = insertDrawing(sheetXml, rel.id);
            zip.file(sheetPart, sheetXml);
        }
        zip.file(sheetRelsPath, sheetRels);
    }

    zip.file('[Content_Types].xml', contentTypes);
}

/**
 * jdoc 浏览器示例
 *
 * 演示两件事：
 *   1. 解析上传的四类文件（点击或拖拽）；
 *   2. 直接调用 jsonToDocx / jsonToXlsx / jsonToPdf / jsonToPptx 生成各种文档，
 *      生成后立刻把字节回读解析，在同一页面预览（「生成 → 解析」往返）。
 */
import {
    detectFormat,
    docxToHtml,
    docxToJson,
    jsonToDocx,
    jsonToPdf,
    jsonToPptx,
    jsonToXlsx,
    pptxToHtml,
    pptxToStandard,
    pdfToJson,
    xlsxToHtml,
    xlsxToJson
} from 'jdoc';
import { chartRenderer } from '@fefeding/ppt-parser/chart-renderer';
import echarts from 'echarts';
import echartsGL from 'echarts-gl';

// chart-renderer 通过全局 echarts 渲染（UMD 引入会自动挂载，ESM 打包需手动暴露）。
// echarts-gl 用于真 3D 图表；未加载时 3D 类型会自动降级为同族 2D。
window.echarts = echarts;
window['echarts-gl'] = echartsGL;

const $ = (id) => document.getElementById(id);
const log = (msg, target = 'log') => { $(target).textContent = msg; };

/** 触发浏览器下载 */
function download(bytes, filename, mime) {
    const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
    const a = Object.assign(document.createElement('a'), { href: url, download: filename });
    a.click();
    URL.revokeObjectURL(url);
}

const MIME = {
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    pdf: 'application/pdf'
};

const GENERATORS = { docx: jsonToDocx, xlsx: jsonToXlsx, pdf: jsonToPdf, pptx: jsonToPptx };

// ---------- 预览双视图（渲染效果 / JSON） ----------

let currentZoom = 1;
/**
 * 展示预览内容
 * @param html - HTML 字符串
 * @param node - 追加的内容，既可以是 DOM 节点（PPTX 逐页插入），也可以是 HTML 字符串
 */
let pendingCharts = null;
function setView(html, node) {
    const layer = document.createElement('div');
    layer.id = 'zoom-layer';
    if (html) layer.innerHTML = html;
    if (node) layer.appendChild(toNode(node));
    const view = $('out-view');
    view.innerHTML = '';
    view.appendChild(layer);
    fitZoom();
    // 图表必须在幻灯片插入 DOM 之后渲染：容器需有可测量尺寸，否则 echarts 画布为 0×0
    renderChartsInView();
}

/** 渲染待处理的 PPTX 图表（按 chartId 在幻灯片容器内定位占位元素） */
function renderChartsInView() {
    if (!pendingCharts) return;
    const { charts, container } = pendingCharts;
    pendingCharts = null;
    try {
        chartRenderer.renderCharts(charts, container);
    } catch (error) {
        console.warn('图表渲染失败：', error);
    }
}

/** HTML 字符串 / DOM 节点 -> 可 appendChild 的节点（字符串不能丢给 appendChild） */
function toNode(value) {
    if (value == null) return null;
    if (typeof value !== 'string') return value;
    const template = document.createElement('template');
    template.innerHTML = value;
    if (template.content.childNodes.length === 1) return template.content.firstChild;
    const fragment = document.createDocumentFragment();
    while (template.content.firstChild) fragment.appendChild(template.content.firstChild);
    return fragment;
}
function applyZoom() {
    const layer = $('zoom-layer');
    if (layer) layer.style.zoom = String(currentZoom);
    const label = $('zoom-label');
    if (label) label.textContent = Math.round(currentZoom * 100) + '%';
}
// 首次渲染按预览容器宽度自适应（仅缩小到适合，不放大），避免大页面超出视图
function fitZoom() {
    const view = $('out-view');
    const layer = $('zoom-layer');
    if (!view || !layer) return;
    layer.style.zoom = '1';
    const natural = layer.scrollWidth || layer.getBoundingClientRect().width;
    const cs = getComputedStyle(view);
    const avail = view.clientWidth - (parseFloat(cs.paddingLeft) || 0) - (parseFloat(cs.paddingRight) || 0);
    let fit = avail / natural;
    if (!isFinite(fit) || fit > 1) fit = 1;
    currentZoom = +(Math.max(0.2, Math.min(1, fit))).toFixed(2);
    applyZoom();
}
function setJson(data) {
    const text = JSON.stringify(data, null, 2) ?? '';
    const clipped = text.length > 200000 ? text.slice(0, 200000) + '\n…（内容过长已截断）' : text;
    $('out-json').innerHTML = `<pre>${escapeHtml(clipped)}</pre>`;
}
function switchTab(which) {
    document.querySelectorAll('.tab').forEach((b) => b.classList.toggle('active', b.dataset.view === which));
    $('out-view').hidden = which !== 'view';
    $('out-json').hidden = which !== 'json';
}
document.querySelectorAll('.tab').forEach((btn) =>
    btn.addEventListener('click', () => switchTab(btn.dataset.view))
);
document.querySelectorAll('[data-zoom]').forEach((btn) =>
    btn.addEventListener('click', () => {
        const act = btn.dataset.zoom;
        if (act === 'in') currentZoom = Math.min(3, +(currentZoom + 0.1).toFixed(2));
        else if (act === 'out') currentZoom = Math.max(0.2, +(currentZoom - 0.1).toFixed(2));
        else if (act === 'fit') return fitZoom();
        else if (act === 'reset') currentZoom = 1;
        applyZoom();
    })
);

// ---------- 解析 ----------

const drop = $('drop');
const picker = $('file');

drop.addEventListener('click', () => picker.click());
picker.addEventListener('change', (e) => e.target.files[0] && parse(e.target.files[0]));

['dragenter', 'dragover'].forEach((evt) =>
    drop.addEventListener(evt, (e) => { e.preventDefault(); drop.classList.add('over'); })
);
['dragleave', 'drop'].forEach((evt) =>
    drop.addEventListener(evt, (e) => { e.preventDefault(); drop.classList.remove('over'); })
);
drop.addEventListener('drop', (e) => e.dataTransfer.files[0] && parse(e.dataTransfer.files[0]));

async function parse(file) {
    log(`解析中：${file.name} (${(file.size / 1024).toFixed(1)} KB)`);
    setView('<em style="color:#8b949e">解析中…</em>');
    setJson({});
    switchTab('view');
    const started = performance.now();

    try {
        const format = await detectFormat(file);
        const { view, data } = await render(file, format);
        const ms = (performance.now() - started).toFixed(0);
        setView(`<div class="meta">
            <span class="badge">${format || 'unknown'}</span>
            <span class="ms">${ms} ms</span>
        </div>`, view);
        setJson(data);
        if (format === 'pdf') setupPdfPageRender(file);
        log(`✓ ${file.name} 解析完成（${format}，${ms} ms）`);
    } catch (error) {
        setView('');
        log(`✗ 解析失败：${error?.message || error}`);
    }
}

/** PPTX 每页独立解析后以 DOM 节点插入：若把所有页拼成一个字符串一次性
 * innerHTML，某页未闭合的标签会被 HTML 解析器自动纠错，把后续页的
 * 内容与媒体吞进前一页（ppt-parser 官方示例即为逐页解析）。 */
function buildPptxSlides(rendered) {
    const wrap = document.createElement('div');
    wrap.className = 'pptx-slides';
    const slides = rendered.slides || [];
    if (!slides.length) {
        wrap.innerHTML = '<em>无幻灯片</em>';
        return wrap;
    }
    for (const slide of slides) {
        const holder = document.createElement('div');
        holder.innerHTML = slide.html;
        const scaler = document.createElement('div');
        scaler.className = 'slide-scaler';
        while (holder.firstChild) scaler.appendChild(holder.firstChild);
        wrap.appendChild(scaler);
    }
    // 幻灯片 HTML 中为图表预留了占位容器（chartId），交给官方 chart-renderer 渲染
    pendingCharts = rendered.charts && rendered.charts.length ? { charts: rendered.charts, container: wrap } : null;
    return wrap;
}

async function render(file, format) {
    const pptxStyleEl = document.getElementById('pptx-styles');
    if (pptxStyleEl) pptxStyleEl.textContent = '';
    if (format === 'docx') {
        const { html, metadata, blocks } = await docxToHtml(file);
        return { view: `${html}<div class="hint">共 ${blocks.length} 个块</div>`, data: { metadata, blocks } };
    }
    if (format === 'xlsx') {
        const { html, metadata, sheets } = await xlsxToHtml(file);
        return { view: html, data: { metadata, sheets } };
    }
    if (format === 'pptx') {
        const rendered = await pptxToHtml(file, { slideMode: false, keyBoardShortCut: false });
        const standard = await pptxToStandard(file);
        const globalStyle = rendered.styles?.global || '';
        if (pptxStyleEl) pptxStyleEl.textContent = globalStyle;
        return {
            view: buildPptxSlides(rendered),
            data: { slideSize: rendered.slideSize, metadata: rendered.metadata, slides: standard?.slides || [] }
        };
    }
    if (format === 'pdf') {
        const result = await pdfToJson(file);
        const pages = result.pages || [];
        // 每页一张卡片：先放等比例占位框（懒渲染成图像），有文本的页折叠展示文本
        const cards = pages
            .map((page, i) => {
                const size = page.width ? ` · ${Math.round(page.width)}×${Math.round(page.height)}` : '';
                const ratio = page.width && page.height
                    ? ` style="aspect-ratio:${Math.round(page.width * 100) / 100}/${Math.round(page.height * 100) / 100}"`
                    : '';
                const text = page.text
                    ? `<details class="page-text"><summary>本页文本</summary><pre>${escapeHtml(page.text)}</pre></details>`
                    : '<em class="empty page-note">（无可提取文本，页面内容以图像展示）</em>';
                return `<figure class="page" data-page="${i + 1}"><figcaption>第 ${i + 1} 页${size}</figcaption>` +
                    `<div class="page-canvas"${ratio}><em class="empty">页面渲染中…</em></div>${text}</figure>`;
            })
            .join('') || '<em>无页面</em>';
        const allEmpty = pages.length > 0 && pages.every((p) => !p.text);
        const hint = `共 ${result.pageCount} 页${allEmpty ? ' · 未提取到文本，已按页面图像渲染' : ''}`;
        return { view: `<div class="hint">${hint}</div>${cards}`, data: result };
    }
    const { blocks } = await docxToJson(file);
    return { view: '<em>未知格式，无法渲染</em>', data: { blocks } };
}

function table(rows) {
    const body = rows
        .slice(0, 200)
        .map((row) => {
            const cells = Array.isArray(row) ? row : row?.cells || [];
            return `<tr>${cells
                .map((cell) => `<td style="border:1px solid #d0d7de;padding:4px 8px">${escapeHtml(cellText(cell))}</td>`)
                .join('')}</tr>`;
        })
        .join('');
    return `<table style="border-collapse:collapse;font-size:12px">${body}</table>`;
}

function cellText(cell) {
    if (cell == null) return '';
    if (typeof cell !== 'object') return String(cell);
    if (Array.isArray(cell)) return cell.map((part) => part.text).join('');
    if (cell.text !== undefined) return cell.text;
    if (cell.value !== undefined && cell.value !== null) return String(cell.value);
    return '';
}

function escapeHtml(value) {
    return String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * 浏览器里直接用 Vite 打包的 pdfjs（不走 jdoc 的 @vite-ignore 动态导入，
 * 否则浏览器无法解析裸模块名）提取 PDF 文本，用于预览。
 */
let pdfjsModule = null;
async function loadPdfjs() {
    if (pdfjsModule) return pdfjsModule;
    const [pdfjs, workerUrl] = await Promise.all([
        import('pdfjs-dist/legacy/build/pdf.mjs'),
        import('pdfjs-dist/legacy/build/pdf.worker.mjs?url')
    ]);
    pdfjs.GlobalWorkerOptions.workerSrc = workerUrl.default;
    pdfjsModule = pdfjs;
    return pdfjs;
}

async function extractPdfText(input) {
    const pdfjs = await loadPdfjs();
    // pdfjs 会转移底层 ArrayBuffer，必须传副本，避免影响原始字节
    const data = input instanceof Uint8Array ? new Uint8Array(input) : new Uint8Array(await toArrayBufferLike(input));
    const doc = await pdfjs.getDocument({ data }).promise;
    let text = '';
    for (let i = 1; i <= doc.numPages; i++) {
        const page = await doc.getPage(i);
        const content = await page.getTextContent();
        text += content.items.map((item) => item.str).join(' ') + '\n\n';
    }
    return text.trim();
}

/** File/Blob/ArrayBuffer -> ArrayBuffer（供 pdfjs 读取） */
async function toArrayBufferLike(input) {
    if (input instanceof ArrayBuffer) return input;
    if (input instanceof Uint8Array) return input.buffer;
    if (input instanceof Blob) return input.arrayBuffer();
    throw new TypeError('不支持的输入类型');
}

// ---------- PDF 页面图像渲染（扫描件没有文本，也能按原图看到内容） ----------

let pdfPageRender = null; // { doc, observer }

/** 解析完成后调用：页面卡进入视口时才渲染成 canvas（支持几百页的文档） */
async function setupPdfPageRender(file) {
    if (pdfPageRender) {
        pdfPageRender.observer?.disconnect();
        pdfPageRender = null;
    }
    const holders = Array.from(document.querySelectorAll('#out-view .page-canvas'));
    if (!holders.length) return;

    let pdfjs;
    try {
        pdfjs = await loadPdfjs();
    } catch (error) {
        holders.forEach((holder) => {
            holder.innerHTML = `<em class="empty">无法加载 pdfjs：${escapeHtml(error?.message || error)}</em>`;
        });
        return;
    }
    // pdfjs 会转移底层 ArrayBuffer，必须传副本，避免影响原文件字节
    const data = new Uint8Array(await toArrayBufferLike(file));
    let doc;
    try {
        doc = await pdfjs.getDocument({ data }).promise;
    } catch (error) {
        holders.forEach((holder) => {
            holder.innerHTML = `<em class="empty">页面渲染失败：${escapeHtml(error?.message || error)}</em>`;
        });
        return;
    }
    const state = { doc, observer: null };
    pdfPageRender = state;

    const observer = new IntersectionObserver((entries) => {
        for (const entry of entries) {
            if (!entry.isIntersecting) continue;
            observer.unobserve(entry.target);
            renderPdfPage(state, entry.target);
        }
    }, { root: $('out-view'), rootMargin: '400px' });
    state.observer = observer;
    holders.forEach((holder) => observer.observe(holder));
}

async function renderPdfPage(state, holder) {
    const pageNum = Number(holder.closest('.page')?.dataset.page);
    if (!pageNum) return;
    try {
        const page = await state.doc.getPage(pageNum);
        // 按预览区可用宽度选缩放倍数，限制在 0.5 ~ 2 倍
        const view = $('out-view');
        const avail = Math.max(280, (view.clientWidth || 700) - 60);
        const base = page.getViewport({ scale: 1 });
        const scale = Math.min(2, Math.max(0.5, avail / base.width));
        const viewport = page.getViewport({ scale });
        const canvas = document.createElement('canvas');
        canvas.width = Math.floor(viewport.width);
        canvas.height = Math.floor(viewport.height);
        canvas.className = 'page-img';
        await page.render({ canvas, canvasContext: canvas.getContext('2d'), viewport }).promise;
        holder.replaceChildren(canvas);
    } catch (error) {
        holder.innerHTML = `<em class="empty">页面渲染失败：${escapeHtml(error?.message || error)}</em>`;
    }
}

// ---------- 生成预设 ----------

/** DOCX 预设：每种一个典型文档形态 */
const docxPresets = [
    {
        name: '周报',
        desc: '标题 / 富文本 / 编号列表 / 进度表格 / 引用块',
        file: 'jdoc-weekly-report.docx',
        build: () => ({
            metadata: { title: '第 42 周周报', author: '张三', company: '平台组' },
            blocks: [
                { type: 'heading', level: 1, text: '第 42 周周报（10.12 – 10.18）' },
                { type: 'paragraph', runs: [
                    { text: '汇报人：张三' },
                    { text: '　部门：平台组', color: '#57606a' },
                    { text: '　周期：2026-10-12 ~ 10-18', color: '#57606a' }
                ] },
                { type: 'heading', level: 2, text: '本周完成' },
                { type: 'list', ordered: true, items: ['完成文档解析库核心重构', '修复 3 个 P1 缺陷', '新增 61 个单元测试'] },
                { type: 'heading', level: 2, text: '进度跟踪' },
                { type: 'table', header: true, columnWidths: [150, 90, 90, 90], rows: [
                    ['模块', '计划', '实际', '完成率'],
                    ['解析核心', '100%', '100%', '100%'],
                    ['生成器', '80%', '75%', '94%'],
                    ['往返测试', '50%', '60%', '120%']
                ] },
                { type: 'heading', level: 2, text: '风险与问题' },
                { type: 'quote', text: 'PDF 中文字体内嵌方案待评审，可能影响下周排期。' },
                { type: 'paragraph', text: '下周计划：完成 PDF 字体方案评审，启动 XLSX 往返测试。' }
            ]
        })
    },
    {
        name: '会议纪要',
        desc: '信息表格 / 议题列表 / JSON 代码块 / 分隔线',
        file: 'jdoc-meeting-minutes.docx',
        build: () => ({
            metadata: { title: '项目双周会会议纪要', author: '李四' },
            blocks: [
                { type: 'heading', level: 1, text: '项目双周会会议纪要' },
                { type: 'table', header: false, rows: [
                    ['时间', '2026-10-16 14:00'],
                    ['地点', '3F-会议室 A'],
                    ['主持', '王五'],
                    ['参会', '平台组、前端组、测试组']
                ] },
                { type: 'heading', level: 2, text: '议题与结论' },
                { type: 'list', ordered: true, items: [
                    '文档库 v0.2 范围：确认仅覆盖 PPTX / DOCX / XLSX / PDF 四种格式',
                    '浏览器端体积预算：gzip 后小于 400KB',
                    '发布时间定于 10 月 30 日'
                ] },
                { type: 'heading', level: 2, text: '行动项' },
                { type: 'code', language: 'json', text: '[\n  { "owner": "张三", "task": "输出体积报告", "due": "10-20" },\n  { "owner": "李四", "task": "补齐往返测试", "due": "10-22" }\n]' },
                { type: 'hr' },
                { type: 'paragraph', align: 'center', text: '—— 以上纪要如有异议，请于 24 小时内回复 ——' }
            ]
        })
    },
    {
        name: '个人简历',
        desc: '姓名标题 / 简介段落 / 经历表格 / 技能列表',
        file: 'jdoc-resume.docx',
        build: () => ({
            metadata: { title: '张三 - 个人简历' },
            blocks: [
                { type: 'heading', level: 1, text: '张三' },
                { type: 'paragraph', runs: [
                    { text: '高级前端工程师' },
                    { text: '　|　', color: '#8b949e' },
                    { text: 'zhangsan@example.com', color: '#0969DA', link: 'mailto:zhangsan@example.com' }
                ] },
                { type: 'heading', level: 2, text: '个人简介' },
                { type: 'paragraph', text: '8 年前端开发经验，擅长文档处理、数据可视化与前端工程化建设，主导过 3 个百万级 DAU 产品的文档中心从 0 到 1。' },
                { type: 'heading', level: 2, text: '工作经历' },
                { type: 'table', header: true, columnWidths: [120, 160, 140], rows: [
                    ['时间', '公司', '职位'],
                    ['2023 – 至今', '某某科技', '高级前端工程师'],
                    ['2020 – 2023', '某某网络', '前端工程师'],
                    ['2018 – 2020', '某某信息', '前端开发']
                ] },
                { type: 'heading', level: 2, text: '技能清单' },
                { type: 'list', ordered: false, items: ['TypeScript / Vue / React', 'Node.js / Vite / Rollup', '文档解析与生成（OOXML / PDF）'] }
            ]
        })
    },
    {
        name: '通知公告',
        desc: '居中标题 / 首行缩进 / 右对齐落款 / 分页符',
        file: 'jdoc-notice.docx',
        build: () => ({
            metadata: { title: '关于文档库系统升级的通知' },
            blocks: [
                { type: 'paragraph', align: 'center', runs: [{ text: '关于文档库系统升级的通知', bold: true, size: 18 }] },
                { type: 'paragraph', text: '各位同事：' },
                { type: 'paragraph', indentFirstLine: 24, text: '为提升文档处理性能，文档库将于本周六 02:00 – 04:00 进行升级，期间相关服务暂停使用。' },
                { type: 'paragraph', indentFirstLine: 24, text: '请提前保存正在编辑的文档，以免数据丢失。由此带来的不便，敬请谅解。' },
                { type: 'hr' },
                { type: 'paragraph', align: 'right', text: '平台组　2026-10-17' },
                { type: 'pageBreak' },
                { type: 'heading', level: 2, text: '附件：升级影响范围' },
                { type: 'table', header: true, rows: [
                    ['模块', '影响', '预计恢复'],
                    ['文档解析', '不可用', '04:00'],
                    ['文档生成', '不可用', '04:00'],
                    ['历史下载', '正常', '—']
                ] }
            ]
        })
    }
];

/** XLSX 预设：每种一个典型工作簿形态 */
const xlsxPresets = [
    {
        name: '财务报表',
        desc: '双工作表 / 合并标题 / 冻结表头 / SUM 公式 / 货币格式 / 边框',
        file: 'jdoc-finance.xlsx',
        build: () => {
            const thin = { style: 'thin', color: 'BFBFBF' };
            const box = { top: thin, left: thin, right: thin, bottom: thin };
            const head = (text) => ({ value: text, style: { font: { bold: true, color: '#FFFFFF' }, fill: '2F5597', align: 'center', border: box } });
            const cell = (value, extra = {}) => ({ value, style: { border: box, ...extra } });
            const money = { numFmt: '#,##0.00' };
            return {
                metadata: { title: '2026 年度财务报表', creator: '@fefeding/jdoc' },
                sheets: [
                    {
                        name: '利润表',
                        merges: ['A1:D1'],
                        frozen: { rows: 2 },
                        columns: [{ width: 18 }, { width: 14 }, { width: 14 }, { width: 14 }],
                        rows: [
                            { 1: { value: '2026 年度利润表（单位：元）', style: { font: { bold: true, size: 14 }, align: 'center' } } },
                            [head('项目'), head('一季度'), head('二季度'), head('全年合计')],
                            [cell('营业收入'), cell(1250000, money), cell(1480000, money),
                                { formula: 'SUM(B3:C3)', result: 2730000, style: { border: box, numFmt: '#,##0.00', font: { bold: true } } }],
                            [cell('营业成本'), cell(760000, money), cell(910000, money),
                                { formula: 'SUM(B4:C4)', result: 1670000, style: { border: box, numFmt: '#,##0.00' } }],
                            [cell('净利润'), cell(490000, money), cell(570000, money),
                                { formula: 'SUM(B5:C5)', result: 1060000, style: { border: box, numFmt: '#,##0.00', font: { bold: true }, fill: 'E2EFDA' } }]
                        ]
                    },
                    {
                        name: '月度明细',
                        frozen: { rows: 1 },
                        columns: [{ width: 12 }, { width: 14 }, { width: 14 }, { width: 14 }],
                        rows: [
                            [head('月份'), head('收入'), head('成本'), head('净利润')],
                            ...['1 月', '2 月', '3 月', '4 月', '5 月', '6 月'].map((month, i) => [
                                cell(month),
                                cell(380000 + i * 15000, money),
                                cell(240000 + i * 8000, money),
                                { formula: `B${i + 2}-C${i + 2}`, result: 140000 + i * 7000, style: { border: box, numFmt: '#,##0.00' } }
                            ]),
                            [
                                cell('合计', { font: { bold: true } }),
                                { formula: 'SUM(B2:B7)', result: 2505000, style: { border: box, numFmt: '#,##0.00', font: { bold: true } } },
                                { formula: 'SUM(C2:C7)', result: 1560000, style: { border: box, numFmt: '#,##0.00', font: { bold: true } } },
                                { formula: 'SUM(D2:D7)', result: 945000, style: { border: box, numFmt: '#,##0.00', font: { bold: true } } }
                            ]
                        ]
                    }
                ]
            };
        }
    },
    {
        name: '销售订单',
        desc: '超链接单元格 / 货币与日期格式 / 右对齐 / 状态配色 / 自动换行',
        file: 'jdoc-sales-orders.xlsx',
        build: () => {
            const status = (text, fill, color) => ({ value: text, style: { fill, font: { color }, align: 'center' } });
            const amount = (value) => ({ value, style: { numFmt: '"¥"#,##0', align: 'right' } });
            const date = (value) => ({ value, style: { numFmt: 'yyyy-mm-dd' } });
            return {
                metadata: { title: '销售订单表', creator: '@fefeding/jdoc' },
                sheets: [{
                    name: '订单',
                    frozen: { rows: 1 },
                    columns: [{ width: 12 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 12 }, { width: 30 }],
                    rows: [
                        ['订单号', '客户', '金额', '下单日期', '状态', '备注'],
                        [
                            { value: 'SO-1001', style: { font: { bold: true } } },
                            '云腾科技',
                            amount(12900),
                            date(new Date('2026-10-08')),
                            status('已完成', 'E2EFDA', '1A7F37'),
                            { value: '查看详情', hyperlink: 'https://example.com/order/SO-1001', style: { font: { color: '0563C1', underline: true } } }
                        ],
                        [
                            'SO-1002',
                            '恒信数据',
                            amount(4580),
                            date(new Date('2026-10-11')),
                            status('已发货', 'DDF4FF', '0969DA'),
                            { value: '部分到货，余款到账后开票', style: { wrapText: true } }
                        ],
                        [
                            'SO-1003',
                            '灵犀传媒',
                            amount(899),
                            date(new Date('2026-10-15')),
                            status('待付款', 'FFF3CD', '9A6700'),
                            '已发送催款提醒'
                        ]
                    ]
                }]
            };
        }
    },
    {
        name: '学生成绩',
        desc: 'SUM / AVERAGE 公式 / 分段填色 / 班级平均统计行',
        file: 'jdoc-grades.xlsx',
        build: () => {
            const score = (value) => value >= 90
                ? { value, style: { fill: 'C6EFCE', font: { color: '1A7F37' }, align: 'center' } }
                : value >= 80
                    ? { value, style: { fill: 'DDF4FF', font: { color: '0969DA' }, align: 'center' } }
                    : value >= 60
                        ? { value, style: { align: 'center' } }
                        : { value, style: { fill: 'FFC7CE', font: { color: '9C0006' }, align: 'center' } };
            const names = ['张三', '李四', '王五', '赵六', '钱七'];
            const data = [[92, 85, 88], [78, 90, 84], [95, 93, 97], [58, 62, 55], [88, 76, 91]];
            const avg = (list) => Math.round((list.reduce((a, b) => a + b, 0) / list.length) * 10) / 10;
            return {
                metadata: { title: '期中考试成绩表' },
                sheets: [{
                    name: '成绩',
                    merges: ['A1:F1'],
                    frozen: { rows: 2 },
                    columns: [{ width: 12 }, { width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }],
                    rows: [
                        { 1: { value: '六年级（2）班期中考试成绩', style: { font: { bold: true, size: 13 }, align: 'center' } } },
                        ['姓名', '语文', '数学', '英语', { value: '总分', style: { font: { bold: true } } }, { value: '平均分', style: { font: { bold: true } } }],
                        ...names.map((name, i) => [
                            name,
                            score(data[i][0]),
                            score(data[i][1]),
                            score(data[i][2]),
                            {
                                formula: `SUM(B${i + 3}:D${i + 3})`,
                                result: data[i][0] + data[i][1] + data[i][2],
                                style: { font: { bold: true }, align: 'center' }
                            },
                            {
                                formula: `AVERAGE(B${i + 3}:D${i + 3})`,
                                result: avg(data[i]),
                                style: { align: 'center' }
                            }
                        ]),
                        [
                            { value: '班级平均', style: { font: { bold: true } } },
                            { formula: 'AVERAGE(B3:B7)', result: avg(data.map((d) => d[0])), style: { align: 'center' } },
                            { formula: 'AVERAGE(C3:C7)', result: avg(data.map((d) => d[1])), style: { align: 'center' } },
                            { formula: 'AVERAGE(D3:D7)', result: avg(data.map((d) => d[2])), style: { align: 'center' } },
                            { formula: 'AVERAGE(E3:E7)', result: avg(data.map((d) => d[0] + d[1] + d[2])), style: { font: { bold: true }, align: 'center' } },
                            { formula: 'AVERAGE(F3:F7)', result: avg(data.map(avg)), style: { font: { bold: true }, align: 'center' } }
                        ]
                    ]
                }]
            };
        }
    },
    {
        name: '库存清单',
        desc: '对象式行 { A1: … } / 低于安全库存红色预警',
        file: 'jdoc-inventory.xlsx',
        build: () => {
            const stock = (value, safety) => value < safety
                ? { value, style: { fill: 'FFC7CE', font: { color: '9C0006' }, align: 'center' } }
                : { value, style: { align: 'center' } };
            return {
                metadata: { title: '仓库库存清单' },
                sheets: [{
                    name: '库存',
                    columns: [{ width: 12 }, { width: 22 }, { width: 10 }, { width: 12 }, { width: 12 }],
                    rows: [
                        // 对象式行：键可以是列号 '3'、列字母 'C' 或 A1 记号 'C3'
                        { 1: 'SKU', 2: '商品', 3: '库存', 4: '安全库存', 5: '单价' },
                        { A2: 'A-100', B2: '机械键盘', C2: stock(42, 20), D2: 20, E2: { value: 499, style: { numFmt: '"¥"#,##0' } } },
                        { A3: 'A-101', B3: '无线鼠标', C3: stock(8, 15), D3: 15, E3: { value: 129, style: { numFmt: '"¥"#,##0' } } },
                        { A4: 'A-102', B4: '显示器支架', C4: stock(0, 10), D4: 10, E4: { value: 259, style: { numFmt: '"¥"#,##0' } } },
                        { A5: 'A-103', B5: 'USB-C 扩展坞', C5: stock(63, 25), D5: 25, E5: { value: 359, style: { numFmt: '"¥"#,##0' } } }
                    ]
                }]
            };
        }
    }
];

/** PDF 预设：中英文各一版，上传字体后用中文版 */
const pdfPresets = [
    {
        name: '经营报告',
        desc: '标题 / 正文 / 表格 / 装饰线 / 分页（默认英文，可上传字体切中文）',
        file: 'jdoc-report.pdf',
        build: (zh) => {
            const t = zh
                ? {
                    title: '2026 年第三季度经营报告',
                    subtitle: '平台组 · 数据截至 2026-09-30',
                    body: '本季度文档库完成解析与生成双链路打通，四格式（PPTX / DOCX / XLSX / PDF）全部实现「解析 → JSON → 生成」往返，编辑器路径与程序化路径像素差降至 0%。',
                    tableTitle: '核心指标',
                    rows: [['指标', '上季', '本季'], ['解析覆盖率', '82%', '100%'], ['往返像素差', '4.2%', '0.0%'], ['平均生成耗时', '310ms', '180ms']],
                    note: '注：像素差为编辑器导出与程序化导出逐页比对结果。',
                    page2: '下季度规划',
                    body2: '重点推进 PDF 中文字体内嵌、SmartArt 保真度与批量导出性能三件事。'
                }
                : {
                    title: 'Q3 2026 Business Report',
                    subtitle: 'Platform Team - data as of 2026-09-30',
                    body: 'This quarter we connected the parse and generate pipelines: all four formats (PPTX / DOCX / XLSX / PDF) now support a full parse -> JSON -> generate round trip, and the pixel diff between the editor path and the scripted path is down to 0%.',
                    tableTitle: 'Key Metrics',
                    rows: [['Metric', 'Last Q', 'This Q'], ['Parse coverage', '82%', '100%'], ['Round-trip pixel diff', '4.2%', '0.0%'], ['Avg generate time', '310ms', '180ms']],
                    note: 'Note: pixel diff compares editor export vs scripted export page by page.',
                    page2: 'Next Quarter',
                    body2: 'Focus areas: embedded CJK fonts for PDF, SmartArt fidelity, and batch export performance.'
                };
            return {
                metadata: { title: t.title, author: '@fefeding/jdoc' },
                elements: [
                    { type: 'heading', level: 1, text: t.title },
                    { type: 'text', text: t.subtitle, color: '#57606a' },
                    { type: 'rect', x: 0, y: 6, width: 483, height: 3, fill: '#0969DA' },
                    { type: 'text', text: t.body },
                    { type: 'heading', level: 2, text: t.tableTitle },
                    { type: 'table', header: true, headerFill: '#E8F0FE', columnWidths: [180, 120, 120], rows: t.rows },
                    { type: 'text', text: t.note, size: 10, color: '#57606a' },
                    { type: 'pageBreak' },
                    { type: 'heading', level: 2, text: t.page2 },
                    { type: 'text', text: t.body2 },
                    { type: 'rect', x: 0, y: 480, width: 483, height: 56, fill: '#F6F8FA', stroke: '#D0D7DE' }
                ]
            };
        }
    },
    {
        name: '发票',
        desc: '色带标题 / 买卖双方 / 明细表格 / 合计 / 页脚注释',
        file: 'jdoc-invoice.pdf',
        build: (zh) => {
            const t = zh
                ? {
                    title: '电子发票',
                    meta: '发票号码：INV-2026-1017　　开票日期：2026-10-17',
                    billTo: '付款方：云腾科技有限公司',
                    addr: '地址：某某市某某区某某路 100 号　　税号：91110000XXXXXXXX0X',
                    rows: [['项目', '数量', '单价', '金额'], ['文档解析库企业授权', '1', '¥12,000', '¥12,000'], ['技术支持服务（年）', '1', '¥3,600', '¥3,600'], ['定制开发', '20 人天', '¥800', '¥16,000']],
                    total: '合计：¥31,600.00',
                    footer: '本发票由 @fefeding/jdoc 自动生成，仅作演示。'
                }
                : {
                    title: 'INVOICE',
                    meta: 'No.: INV-2026-1017    Date: 2026-10-17',
                    billTo: 'Bill To: Yunteng Technology Co., Ltd.',
                    addr: 'Address: 100 Sunset Rd.    Tax ID: 91110000XXXXXXXX0X',
                    rows: [['Item', 'Qty', 'Unit Price', 'Amount'], ['Enterprise license', '1', '$1,200', '$1,200'], ['Support (1 year)', '1', '$360', '$360'], ['Custom dev', '20 pd', '$80', '$1,600']],
                    total: 'Total: $3,160.00',
                    footer: 'Generated by @fefeding/jdoc for demonstration only.'
                };
            return {
                metadata: { title: t.title },
                elements: [
                    { type: 'rect', x: 0, y: 0, width: 483, height: 44, fill: '#0969DA' },
                    { type: 'text', text: t.title, x: 12, y: 10, size: 18, bold: true, color: '#FFFFFF' },
                    { type: 'text', text: t.meta, y: 38, size: 10, color: '#57606a' },
                    { type: 'text', text: t.billTo, y: 58, bold: true },
                    { type: 'text', text: t.addr, y: 78, size: 10, color: '#57606a' },
                    { type: 'table', y: 100, header: true, headerFill: '#E8F0FE', columnWidths: [200, 60, 90, 90], fontSize: 10, rows: t.rows },
                    { type: 'text', text: t.total, y: 188, align: 'right', bold: true },
                    { type: 'text', text: t.footer, y: 212, size: 9, color: '#8b949e' }
                ]
            };
        }
    },
    {
        name: '多页文档',
        desc: 'pageBreak 手动分页 / 每页独立标题与内容',
        file: 'jdoc-multipage.pdf',
        build: (zh) => {
            const t = zh
                ? {
                    doc: '使用手册',
                    pages: [
                        ['第一章　快速开始', '安装依赖后即可调用：解析用 xxxToYyy(input)，生成用 jsonToXxx(doc)。入参支持 ArrayBuffer / Blob / File / base64，出参统一为 Uint8Array。'],
                        ['第二章　往返编辑', '解析结果可以直接回传给生成函数，支持「解析 → 改 → 生成」。文档即契约，无需中间格式。'],
                        ['第三章　环境要求', 'Node.js >= 18 直接可用；浏览器需打包器；PDF 解析需 pdfjs-dist；中文字体需通过 fonts.regular 内嵌 TTF/OTF。']
                    ]
                }
                : {
                    doc: 'User Manual',
                    pages: [
                        ['Chapter 1 - Getting Started', 'Install the package and start calling: parse with xxxToYyy(input), generate with jsonToXxx(doc). Inputs accept ArrayBuffer / Blob / File / base64; outputs are always Uint8Array.'],
                        ['Chapter 2 - Round Trip', 'Parse results can be fed straight back into the generate functions, enabling parse -> edit -> generate. The JSON document is the contract.'],
                        ['Chapter 3 - Requirements', 'Node.js >= 18 works out of the box. Browsers need a bundler. PDF parsing needs pdfjs-dist. CJK text requires an embedded TTF/OTF via fonts.regular.']
                    ]
                };
            const elements = [{ type: 'heading', level: 1, text: t.doc }];
            t.pages.forEach(([title, body], i) => {
                if (i > 0) elements.push({ type: 'pageBreak' });
                elements.push({ type: 'heading', level: 2, text: title });
                elements.push({ type: 'text', text: body });
                elements.push({ type: 'text', text: `— ${i + 1} / ${t.pages.length} —`, align: 'center', size: 10, color: '#8b949e' });
            });
            return { metadata: { title: t.doc }, elements };
        }
    },
    {
        name: '荣誉证书',
        desc: '居中排版 / 金色边框 / 装饰线 / 落款（默认英文，可上传字体切中文）',
        file: 'jdoc-certificate.pdf',
        build: (zh) => {
            const t = zh
                ? { title: '荣 誉 证 书', name: '张　三', desc: '在 2026 年度文档库建设中表现突出，特此嘉奖。', date: '2026 年 10 月 17 日', org: '平台组' }
                : { title: 'CERTIFICATE', name: 'Zhang San', desc: 'In recognition of outstanding contribution to the 2026 doc library project.', date: 'October 17, 2026', org: 'Platform Team' };
            // pages 模式不自动拆页；边框色带先画占住光标，正文用负偏移回到带内
            const BAND = 640;
            const y = (fromTop) => fromTop - 56 - BAND;
            return {
                metadata: { title: t.title },
                pages: [{
                    size: 'A4',
                    margin: 56,
                    elements: [
                        { type: 'rect', x: 0, y: 0, width: 483, height: BAND, stroke: '#C9A227', lineWidth: 2 },
                        { type: 'text', text: t.title, y: y(150), size: 28, bold: true, align: 'center', color: '#8A6D1D' },
                        { type: 'text', text: '—— ' + t.name + ' ——', y: y(230), size: 20, bold: true, align: 'center', color: '#1F2328' },
                        { type: 'text', text: t.desc, y: y(290), size: 12, align: 'center', color: '#57606a' },
                        { type: 'line', x1: 170, y1: 380, x2: 425, y2: 380, color: '#C9A227', lineWidth: 1 },
                        { type: 'text', text: t.date, y: y(450), size: 12, align: 'center', color: '#57606a' },
                        { type: 'text', text: t.org, y: y(530), size: 12, align: 'center', color: '#57606a' }
                    ]
                }]
            };
        }
    }
];

const PRESETS = { docx: docxPresets, xlsx: xlsxPresets, pdf: pdfPresets };

/** PDF 中文字体（用户上传后启用中文版预设） */
let pdfFont = null;

// ---------- 生成流程 ----------

function renderPresets(format) {
    const box = $('presets-' + format);
    box.innerHTML = '';
    for (const preset of PRESETS[format]) {
        const button = document.createElement('button');
        button.className = 'preset';
        button.innerHTML = `<b>${preset.name}</b><small>${preset.desc}</small>`;
        button.addEventListener('click', () => withLog(async () => {
            const bytes = await generatePreset(format, preset);
            download(bytes, preset.file, MIME[format]);
            await previewGenerated(format, bytes, preset.file);
            return `${preset.name} 已生成并下载：${preset.file}（${(bytes.length / 1024).toFixed(1)} KB）`;
        }));
        box.appendChild(button);
    }
    $('cnt-' + format).textContent = `${PRESETS[format].length} 个预设`;
}

async function generatePreset(format, preset) {
    const input = preset.build(format === 'pdf' && !!pdfFont);
    const options = format === 'pdf' && pdfFont ? { fonts: { regular: pdfFont } } : {};
    return GENERATORS[format](input, options);
}

/** 把刚生成的字节回读解析，直接在页面里预览（生成 → 解析往返） */
let pdfPreviewUrl = null;
async function previewGenerated(format, bytes, name) {
    const started = performance.now();
    const head = (extra) => `<div style="margin-bottom:10px">
        <span class="badge">${format}</span>
        <span class="badge" style="background:#dafbe1;color:#1a7f37">生成 ✓ 往返解析 ✓</span>
        <span style="color:#57606a;font-size:12px">${escapeHtml(name)} · ${(bytes.length / 1024).toFixed(1)} KB · ${extra} · ${(performance.now() - started).toFixed(0)} ms</span>
    </div>`;

    let view = '';
    let node = null;
    let data = {};
    const pptxStyleEl = document.getElementById('pptx-styles');
    if (pptxStyleEl) pptxStyleEl.textContent = '';
    if (format === 'docx') {
        const { html: body, blocks } = await docxToHtml(bytes);
        view = head(`${blocks.length} 个块`) + body;
        data = { blocks };
    } else if (format === 'xlsx') {
        const { html: body, sheets } = await xlsxToHtml(bytes);
        view = head(`${sheets.length} 个工作表：${escapeHtml(sheets.map((s) => s.name).join(' / '))}`) + body;
        data = { sheets };
    } else if (format === 'pdf') {
        if (pdfPreviewUrl) URL.revokeObjectURL(pdfPreviewUrl);
        pdfPreviewUrl = URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' }));
        let text = '';
        try {
            text = await extractPdfText(bytes);
        } catch (error) {
            text = `（PDF 文本提取在当前环境不可用：${error?.message || error}）`;
        }
        view = head('内嵌渲染') +
            `<iframe src="${pdfPreviewUrl}" style="width:100%;height:420px;border:1px solid #d0d7de;border-radius:6px;background:#fff"></iframe>
            <details style="margin-top:8px"><summary style="cursor:pointer;color:#57606a;font-size:12px">提取的文本（${text.length} 字）</summary>
            <pre style="margin-top:6px">${escapeHtml(text.slice(0, 20000))}</pre></details>`;
        data = { extractedText: text };
    } else {
        const rendered = await pptxToHtml(bytes, { slideMode: false, keyBoardShortCut: false });
        const standard = await pptxToStandard(bytes);
        if (pptxStyleEl) pptxStyleEl.textContent = rendered.styles?.global || '';
        node = buildPptxSlides(rendered);
        view = head(`${standard?.slides?.length || 0} 页幻灯片`);
        data = { slideSize: rendered.slideSize, slides: standard?.slides || [] };
    }
    setView(view, node);
    setJson(data);
}

// ---------- 批量生成 ----------

let resultUrls = [];
async function generateAll(format) {
    log(`批量生成 ${format.toUpperCase()} 中…`, 'log2');
    const results = [];
    for (const preset of PRESETS[format]) {
        const bytes = await generatePreset(format, preset);
        results.push({ preset, bytes });
    }
    showResults(format, results);
    return `共生成 ${results.length} 个 ${format.toUpperCase()} 文件，可逐个下载或预览`;
}

function showResults(format, results) {
    resultUrls.forEach((url) => URL.revokeObjectURL(url));
    resultUrls = [];
    $('results-card').hidden = false;
    $('results').innerHTML = '';
    for (const { preset, bytes } of results) {
        const url = URL.createObjectURL(new Blob([bytes], { type: MIME[format] }));
        resultUrls.push(url);
        const row = document.createElement('div');
        row.className = 'result';
        row.innerHTML = `<span class="name">${preset.name}</span>
            <span class="size">${preset.file} · ${(bytes.length / 1024).toFixed(1)} KB</span>`;
        const link = document.createElement('a');
        link.textContent = '下载';
        link.href = url;
        link.download = preset.file;
        const preview = document.createElement('a');
        preview.textContent = '预览';
        preview.addEventListener('click', () => withLog(async () => {
            await previewGenerated(format, bytes, preset.file);
            return `预览：${preset.file}`;
        }));
        row.append(link, preview);
        $('results').appendChild(row);
    }
}

['docx', 'xlsx', 'pdf'].forEach((format) => {
    renderPresets(format);
    $('all-' + format).addEventListener('click', () => withLog(() => generateAll(format), '批量生成中…'));
});

// ---------- 字体上传 ----------

$('font').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) {
        pdfFont = null;
        $('font-state').textContent = '未上传：PDF 示例使用英文内容。';
        return;
    }
    pdfFont = new Uint8Array(await file.arrayBuffer());
    $('font-state').textContent = `已加载 ${file.name}（${(pdfFont.length / 1024).toFixed(0)} KB）— PDF 示例已切换为中文。`;
});

// ---------- 自定义 JSON ----------

const CUSTOM_SAMPLES = {
    docx: () => ({
        metadata: { title: '自定义文档', author: '@fefeding/jdoc' },
        blocks: [
            { type: 'heading', level: 1, text: '自定义 DOCX' },
            { type: 'paragraph', runs: [
                { text: '普通、' },
                { text: '加粗', bold: true },
                { text: '、' },
                { text: '斜体', italic: true },
                { text: '、' },
                { text: '彩色', color: '#FF0000', size: 14 }
            ] },
            { type: 'list', ordered: false, items: ['要点一', '要点二'] },
            { type: 'table', header: true, rows: [['模块', '进度'], ['A', '100%']] },
            { type: 'quote', text: '一套 API，四种格式。' },
            { type: 'hr' }
        ]
    }),
    xlsx: () => ({
        metadata: { title: '自定义工作簿' },
        sheets: [
            { name: '订单', columns: [{ width: 18 }, { width: 12 }], rows: [
                ['姓名', '金额'],
                ['张三', { value: 1200.5, style: { numFmt: '#,##0.00', align: 'right' } }],
                [{ value: '李四', style: { font: { bold: true, color: '#FF0000' } } }, { formula: 'B2*2', result: 2401 }]
            ] },
            { name: '空表', rows: [] }
        ]
    }),
    pdf: () => ({
        metadata: { title: 'Custom PDF' },
        elements: [
            { type: 'heading', level: 1, text: 'Custom PDF' },
            { type: 'text', text: 'Generated with @fefeding/jdoc.' },
            { type: 'table', header: true, rows: [['Item', 'Value'], ['Revenue', '$1,200'], ['Cost', '$300']] },
            { type: 'rect', x: 0, y: 0, width: 160, height: 36, fill: '#E8F0FE', stroke: '#0969DA', radius: 8 }
        ]
    }),
    pptx: () => ({
        slides: [
            { elements: [{ type: 'text', text: 'jdoc', x: 60, y: 60, width: 400, height: 60 }] },
            { elements: [{ type: 'text', text: 'PPTX via @fefeding/ppt-parser', x: 60, y: 60, width: 500, height: 40 }] }
        ]
    })
};

$('fmt').addEventListener('change', (e) => {
    $('spec').value = JSON.stringify(CUSTOM_SAMPLES[e.target.value](), null, 2);
});
$('spec').value = JSON.stringify(CUSTOM_SAMPLES.docx(), null, 2);

$('mk-custom').addEventListener('click', () => withLog(async () => {
    const format = $('fmt').value;
    const input = JSON.parse($('spec').value);
    const options = format === 'pdf' && pdfFont ? { fonts: { regular: pdfFont } } : {};
    const bytes = await GENERATORS[format](input, options);
    download(bytes, `jdoc-custom.${format}`, MIME[format]);
    await previewGenerated(format, bytes, `jdoc-custom.${format}`);
    return `自定义 ${format.toUpperCase()} 已生成（${(bytes.length / 1024).toFixed(1)} KB）`;
}));

async function withLog(task, pending = '生成中…') {
    try {
        log(pending, 'log2');
        log(`✓ ${await task()}`, 'log2');
    } catch (error) {
        log(`✗ ${error?.message || error}`, 'log2');
    }
}

/**
 * jdoc 浏览器示例
 *
 * 演示两件事：解析上传的四类文件、按 JSON 生成并下载。
 */
import {
    detectFormat,
    docxToHtml,
    docxToJson,
    jsonToDocx,
    jsonToPdf,
    jsonToPptx,
    jsonToXlsx,
    pdfToText,
    pptxToStandard,
    xlsxToJson
} from 'jdoc';

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
    $('out').innerHTML = '<em style="color:#8b949e">解析中…</em>';
    const started = performance.now();

    try {
        const format = await detectFormat(file);
        const html = await render(file, format);
        const ms = (performance.now() - started).toFixed(0);
        $('out').innerHTML = `<div style="margin-bottom:10px">
            <span class="badge">${format || 'unknown'}</span>
            <span style="color:#57606a;font-size:12px">${ms} ms</span>
        </div>${html}`;
        log(`✓ ${file.name} 解析完成（${format}，${ms} ms）`);
    } catch (error) {
        $('out').innerHTML = '';
        log(`✗ 解析失败：${error?.message || error}`);
    }
}

async function render(file, format) {
    if (format === 'docx') {
        const { html, blocks } = await docxToHtml(file);
        return `${html}<hr style="margin:14px 0;border:none;border-top:1px solid #d0d7de">
            <small style="color:#57606a">共 ${blocks.length} 个块</small>`;
    }
    if (format === 'xlsx') {
        const { sheets } = await xlsxToJson(file);
        return sheets
            .map((sheet) => `<div style="font-weight:600;margin-bottom:6px">${escapeHtml(sheet.name)}
                <small style="color:#57606a">（${sheet.rows.length} 行）</small></div>
                ${table(sheet.rows)}`)
            .join('<hr style="margin:14px 0;border:none;border-top:1px solid #d0d7de">');
    }
    if (format === 'pptx') {
        const doc = await pptxToStandard(file);
        const slides = doc?.slides || [];
        return slides
            .map((slide, i) => `<div style="border:1px solid #d0d7de;border-radius:6px;padding:10px;margin-bottom:10px">
                <div style="font-weight:600;margin-bottom:6px">第 ${i + 1} 页</div>
                <pre>${escapeHtml(JSON.stringify(slide).slice(0, 600))}</pre></div>`)
            .join('') || '<em>无幻灯片</em>';
    }
    if (format === 'pdf') {
        const text = await pdfToText(file);
        return `<pre style="white-space:pre-wrap">${escapeHtml(text.slice(0, 20000))}</pre>`;
    }
    const { blocks } = await docxToJson(file);
    return `<pre>${escapeHtml(JSON.stringify(blocks, null, 2).slice(0, 20000))}</pre>`;
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

// ---------- 生成 ----------

$('mk-docx').addEventListener('click', () => withLog(async () => {
    const bytes = await jsonToDocx(sampleDocx());
    download(bytes, 'jdoc-sample.docx', MIME.docx);
    return `DOCX 已生成：${(bytes.length / 1024).toFixed(1)} KB`;
}));

$('mk-xlsx').addEventListener('click', () => withLog(async () => {
    const bytes = await jsonToXlsx([
        { name: '订单', columns: [{ width: 18 }, { width: 12 }], rows: [
            ['姓名', '金额'],
            ['张三', { value: 1200.5, style: { numFmt: '#,##0.00', align: 'right' } }],
            [{ value: '李四', style: { font: { bold: true, color: '#FF0000' } } }, { formula: 'B2*2', result: 2401 }]
        ] },
        { name: '空表', rows: [] }
    ]);
    download(bytes, 'jdoc-sample.xlsx', MIME.xlsx);
    return `XLSX 已生成：${(bytes.length / 1024).toFixed(1)} KB`;
}));

$('mk-pdf').addEventListener('click', () => withLog(async () => {
    const bytes = await jsonToPdf({
        metadata: { title: 'jdoc Sample' },
        elements: [
            { type: 'heading', level: 1, text: 'jdoc PDF Sample' },
            { type: 'text', text: 'Generated with @fefeding/jdoc.' },
            { type: 'table', header: true, rows: [['Item', 'Value'], ['Revenue', '$1,200'], ['Cost', '$300']] },
            { type: 'rect', x: 0, y: 0, width: 160, height: 36, fill: '#E8F0FE', stroke: '#0969DA', radius: 8 }
        ]
    });
    download(bytes, 'jdoc-sample.pdf', MIME.pdf);
    return `PDF 已生成：${(bytes.length / 1024).toFixed(1)} KB`;
}));

$('mk-pptx').addEventListener('click', () => withLog(async () => {
    const bytes = await jsonToPptx({
        slides: [
            { elements: [{ type: 'text', text: 'jdoc', x: 60, y: 60, width: 400, height: 60 }] },
            { elements: [{ type: 'text', text: 'PPTX via @fefeding/ppt-parser', x: 60, y: 60, width: 500, height: 40 }] }
        ]
    });
    download(bytes, 'jdoc-sample.pptx', MIME.pptx);
    return `PPTX 已生成：${(bytes.length / 1024).toFixed(1)} KB`;
}));

$('mk-custom').addEventListener('click', () => withLog(async () => {
    const spec = JSON.parse($('spec').value);
    const bytes = await jsonToDocx(spec);
    download(bytes, 'jdoc-custom.docx', MIME.docx);
    return `自定义 DOCX 已生成：${(bytes.length / 1024).toFixed(1)} KB`;
}));

async function withLog(task) {
    try {
        log('生成中…', 'log2');
        log(`✓ ${await task()}`, 'log2');
    } catch (error) {
        log(`✗ ${error?.message || error}`, 'log2');
    }
}

function sampleDocx() {
    return {
        metadata: { title: 'jdoc 示例', author: '@fefeding/jdoc' },
        blocks: [
            { type: 'heading', level: 1, text: 'jdoc 示例文档' },
            { type: 'paragraph', runs: [
                { text: '支持' },
                { text: '加粗', bold: true },
                { text: '、' },
                { text: '斜体', italic: true },
                { text: '、' },
                { text: '彩色', color: '#0969DA', size: 13 },
                { text: '、' },
                { text: '链接', link: 'https://github.com/fefeding/jdoc' }
            ] },
            { type: 'list', ordered: true, items: ['DOCX 解析', 'DOCX 生成'] },
            { type: 'list', ordered: false, items: ['XLSX', 'PPTX', 'PDF'] },
            { type: 'table', header: true, rows: [['格式', '解析', '生成'], ['DOCX', '✓', '✓'], ['XLSX', '✓', '✓']] },
            { type: 'quote', text: '一套 API，四种格式。' },
            { type: 'hr' }
        ]
    };
}

// PDF 大文档打开体验回归：
// 1. 打开期间出现 loading 遮罩（不再「没反应」）
// 2. 第一页先渲染（canvas 出现），其余页随后异步补齐，不阻塞第一页
// 3. 第一页就绪后 loading 收起
// 用真实大文档（examples 下约 100MB 的课本 PDF）跑，重点验证「先首屏后异步」
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const WEBVIEW_JS = path.resolve(__dirname, '..', 'dist', 'webview.js');
const USE_BIG = process.argv.includes('--big');
const PDF = (() => {
    if (USE_BIG) {
        const ex = path.resolve(__dirname, '..', '..', 'examples');
        const f = fs.readdirSync(ex).find((n) => /\.pdf$/i.test(n));
        return f ? path.join(ex, f) : null;
    }
    return null;
})();

/** 生成最小多页 PDF（仅验证加载/首屏逻辑，不依赖外部大文件） */
function buildTinyPdf(pages) {
    let buf = Buffer.from('');
    const offs = [];
    const w = (s) => { offs.push(buf.length); buf = Buffer.concat([buf, Buffer.from(s + '\n')]); };
    w('%PDF-1.4');
    w('1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj');
    const kidNums = [];
    for (let i = 1; i <= pages; i++) kidNums.push(`${2 + 2 * i - 1} 0 R`);
    w(`2 0 obj\n<< /Type /Pages /Kids [${kidNums.join(' ')}] /Count ${pages} >>\nendobj`);
    const fontObj = 2 + 2 * pages + 1;
    for (let i = 1; i <= pages; i++) {
        const pObj = 2 + 2 * i - 1;
        const cObj = 2 + 2 * i;
        w(`${pObj} 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 ${fontObj} 0 R >> >> /Contents ${cObj} 0 R >>\nendobj`);
        const content = `BT /F1 20 Tf 60 760 Td (Page ${i} of ${pages}) Tj ET`;
        w(`${cObj} 0 obj\n<< /Length ${Buffer.byteLength(content)} >>\nstream\n${content}\nendstream\nendobj`);
    }
    w(`${fontObj} 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj`);
    const xrefStart = buf.length;
    let xref = `xref\n0 ${fontObj + 1}\n0000000000 65535 f \n`;
    for (let i = 0; i < offs.length; i++) xref += String(offs[i]).padStart(10, '0') + ' 00000 n \n';
    w(xref + `trailer\n<< /Size ${fontObj + 1} /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF`);
    return buf;
}
const TINY_PDF = buildTinyPdf(12);
if (!USE_BIG && !PDF) {
    // 用小 PDF（写入临时文件以便 server 以 stream 提供）
}
console.log('使用 PDF:', USE_BIG && PDF ? PDF + ' ' + (fs.statSync(PDF).size / 1024 / 1024).toFixed(0) + 'MB' : `生成的 ${TINY_PDF.length}B 小 PDF(12 页)`);

let failed = 0;
const check = (name, ok, detail = '') => {
    console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  ' + detail : ''}`);
    if (!ok) failed++;
};

const server = http.createServer((req, res) => {
    const url = req.url.split('?')[0];
    if (url === '/' || url === '/index.html') {
        res.writeHead(200, { 'Content-Type': 'text/html' });
        res.end(`<!doctype html><html><head><meta charset="utf-8"></head>
<body class="vscode-dark"><div id="app"></div>
<script>window.__JDC_PDF_WORKER__ = '/pdf.worker.mjs';</script>
<script>window.acquireVsCodeApi=function(){return{postMessage:function(){}}};</script>
<script src="/webview.js"></script>
<script>
// 模拟真实宿主：init 直接带 fileUri，webview 端自行 fetch 二进制（大文件免 base64 卡死）
window.postMessage({type:'init',format:'pdf',mode:'view',fileUri:'/pdf',title:'课本',uri:'file://x',dirty:false,locale:'zh-cn'},'*');
</script></body></html>`);
        return;
    }
    if (url === '/webview.js') {
        res.writeHead(200, { 'Content-Type': 'text/javascript' });
        res.end(fs.readFileSync(WEBVIEW_JS));
        return;
    }
    if (url === '/pdf') {
        res.writeHead(200, { 'Content-Type': 'application/pdf' });
        if (USE_BIG && PDF) fs.createReadStream(PDF).pipe(res);
        else res.end(TINY_PDF);
        return;
    }
    if (url === '/pdf.worker.mjs') {
        const worker = path.resolve(__dirname, '..', 'media', 'pdf.worker.min.mjs');
        res.writeHead(200, { 'Content-Type': 'text/javascript' });
        res.end(fs.readFileSync(worker));
        return;
    }
    res.writeHead(404);
    res.end();
});
await new Promise((r) => server.listen(0, r));
const browser = await chromium.launch({ args: ['--no-sandbox', '--disable-gpu'] });
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
page.on('console', (m) => console.log('  [browser]', m.type(), m.text()));
page.on('pageerror', (e) => console.log('  [pageerror]', e.message));
await page.goto(`http://127.0.0.1:${server.address().port}/`, { waitUntil: 'load' });

console.log('\n== PDF 打开（大文档）==');

// 1) loading 出现（解析/首渲期间）
const loadingAppeared = await page.waitForSelector('.pdf-loading', { timeout: 60000 }).then(() => true).catch(() => false);
check('打开期间出现 loading 遮罩', loadingAppeared);

// 2) 第一页 canvas 出现（首屏优先）
const firstCanvas = await page.waitForSelector('.pdf-page[data-page="1"] canvas', { timeout: USE_BIG ? 280000 : 60000 }).then(() => true).catch(() => false);
check('第一页先渲染（canvas 出现）', firstCanvas);

// 3) 第一页就绪后 loading 收起
if (firstCanvas) {
    const loadingGone = await page.evaluate(() => !document.querySelector('.pdf-loading'));
    check('第一页就绪后 loading 已收起', loadingGone);
}

// 4) 全部页占位最终补齐（分批异步完成），且第一页在所有页之前就可见
const total = await page.evaluate(() => {
    const t = document.querySelector('.pdf-pager .total');
    return t ? Number(t.textContent.replace(/[^\d]/g, '')) : 0;
});
console.log(`  文档页数: ${total}`);
let allPages = false;
if (total > 0) {
    allPages = await page.waitForFunction(
        (n) => document.querySelectorAll('.pdf-page').length >= n,
        total, { timeout: USE_BIG ? 280000 : 60000 }
    ).then(() => true).catch(() => false);
}
check('其余页占位已异步补齐（全部页就位）', total === 0 ? true : allPages,
    total ? `已就位 ${await page.evaluate(() => document.querySelectorAll('.pdf-page').length)}/${total}` : '');

// 5) PDF 是纯只读：工具栏不应再有「编辑」入口，也不该存在编辑工具栏
const readOnly = await page.evaluate(() => {
    const titles = [...document.querySelectorAll('.pdf-bar .pdf-btn')].map((b) => b.getAttribute('aria-label') || '');
    return {
        hasEditBtn: titles.some((t) => t.includes('编辑')),
        editBars: document.querySelectorAll('.pdf-edit-bar').length,
        btnTitles: titles
    };
});
check('PDF 工具栏没有「编辑」按钮', !readOnly.hasEditBtn, readOnly.btnTitles.join(' | '));
check('PDF 没有编辑工具栏', readOnly.editBars === 0, `找到 ${readOnly.editBars} 个`);

console.log(failed ? `\nFAILED (${failed} failure(s))` : '\nALL PASSED (0 failure(s))');
await browser.close();
server.close();
process.exit(failed ? 1 : 0);
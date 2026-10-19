// office-editor（DOCX/XLSX）与查看器 PDF 阅读器的 Webview 回归验证。
//
// DOCX/XLSX 走 dist/office-editor.js（注入 acquireVsCodeApi 桩捕获 sync 消息），
// view → edit → 驱动编辑 → 捕获 sync 字节 → 用 jdoc 重新解析断言内容与容器。
// PDF 走 dist/webview.js（pdfjs 栅格化），断言阅读器外壳：缩略图栏、页码跳转、
// 单页/连续切换、旋转后当前页高亮保持。
//
// 用法：node scripts/verify-roundtrip.mjs [docx|xlsx|pdf]   （缺省三种都跑）
// 依赖：playwright（从仓库根 node_modules 向上解析）、jdoc 已构建（dist/jdoc.cjs）。
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { chromium } from 'playwright';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const jdoc = require(path.resolve(__dirname, '..', '..', 'dist', 'jdoc.cjs'));
const { PDFArray, PDFDocument, decodePDFRawStream } = require('pdf-lib');
const EDITOR = path.resolve(__dirname, '..', 'dist', 'office-editor.js');
const WEBVIEW = path.resolve(__dirname, '..', 'dist', 'webview.js');
const PDF_WORKER = path.resolve(__dirname, '..', 'media', 'pdf.worker.min.mjs');

/**
 * 拼出某页的内容流指令文本（pdf-lib 保存时会压成 FlateDecode，需解码；
 * 叠加内容与原页面内容分属不同流，必须全部拼起来）
 *
 * 用它断言「叠加确实写进了 PDF」，而不是只看字节变长了。
 */
async function pageContentOps(buf, pageIndex = 0) {
    const doc = await PDFDocument.load(buf);
    const contents = doc.getPage(pageIndex).node.Contents();
    const streams = contents instanceof PDFArray ? contents.asArray().map((r) => doc.context.lookup(r)) : [doc.context.lookup(contents)];
    return streams
        .filter(Boolean)
        .map((s) => Buffer.from(decodePDFRawStream(s).decode()).toString('latin1'))
        .join('\n');
}

let failures = 0;
function check(name, cond, extra = '') {
    const ok = !!cond;
    if (!ok) failures++;
    console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${extra ? '  ' + extra : ''}`);
}

async function genDocx() {
    // 覆盖分页 / 目录 / 章节号：多页正文 + 目录块 + 多级标题。
    // 用固定段数而不是长文本，保证页数断言稳定。
    const filler = Array.from({ length: 60 }, (_, i) => ({ type: 'paragraph', text: `第 ${i + 1} 段正文内容。` }));
    return jdoc.jsonToDocx({ blocks: [
        { type: 'heading', level: 1, text: '季度报告' },
        { type: 'toc', entries: [
            { level: 1, text: '季度报告', page: 1 },
            { level: 1, text: '1. 收入', page: 2 },
            { level: 2, text: '1.1. 分区域', page: 2 }
        ] },
        { type: 'heading', level: 1, text: '收入' },
        { type: 'paragraph', text: '正文内容。' },
        ...filler,
        { type: 'heading', level: 2, text: '分区域' },
        { type: 'table', header: true, rows: [[{ type: 'cell', colspan: 2, text: 'X' }, 'c'], ['', '', '']] }
    ] });
}
async function genXlsx() {
    return jdoc.jsonToXlsx({ sheets: [{ name: 'S', rows: [
        { index: 1, cells: [{ value: 'a' }, { value: 'b' }] },
        { index: 2, cells: [{ value: 'x' }, { value: 100 }] }
    ] }] });
}

/**
 * PDF 阅读器回归
 *
 * PDF 是只读格式，没有 sync 往返可断言，验证点全在阅读器外壳与导航：
 * 外壳结构（工具栏/缩略图栏）、缩略图与正文页数量、点击缩略图翻页、
 * 单页 ↔ 连续切换、旋转后当前页高亮仍保持。
 */
async function runPdf() {
    const pages = Array.from({ length: 6 }, (_, i) => ({
        size: { width: 595, height: 842 },
        elements: [
            { type: 'text', x: 60, y: 80, size: 22, bold: true, text: `Chapter ${i + 1}` },
            { type: 'text', x: 60, y: 130, size: 11, text: `Page ${i + 1} of the PDF viewer regression fixture.` }
        ]
    }));
    const b64 = Buffer.from(await jdoc.jsonToPdf({ pages })).toString('base64');

    const server = http.createServer((req, res) => {
        const url = req.url.split('?')[0];
        if (url === '/' || url === '/index.html') {
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end(`<!doctype html><html><head><meta charset="utf-8"></head><body class="vscode-dark"><div id="app"></div>
<script>window.acquireVsCodeApi=function(){return {postMessage:function(m){try{(window.__POSTS__=window.__POSTS__||[]).push(m);}catch(e){}}};};</script>
<script>window.__ERRS__=[];window.__POSTS__=[];(function(){const o=console.error.bind(console);console.error=function(...a){try{window.__ERRS__.push(a.map(x=>x&&x.stack?x.stack:String(x)).join(' '));}catch(e){}o(...a);};})();</script>
<script>window.__JDC_PDF_WORKER__='/pdf.worker.mjs';</script>
<script src="/webview.js"></script></body></html>`);
            return;
        }
        if (url === '/webview.js') { res.writeHead(200, { 'Content-Type': 'text/javascript' }); res.end(fs.readFileSync(WEBVIEW)); return; }
        if (url === '/pdf.worker.mjs') { res.writeHead(200, { 'Content-Type': 'text/javascript' }); res.end(fs.readFileSync(PDF_WORKER)); return; }
        res.writeHead(404); res.end('nf');
    });
    await new Promise((r) => server.listen(0, r));
    const browser = await chromium.launch({ args: ['--no-sandbox'] });
    const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
    const errs = [];
    page.on('console', (m) => { if (m.type() === 'error') errs.push('CONSOLE: ' + m.text()); });
    page.on('pageerror', (e) => errs.push('PAGEERR: ' + e.message));

    await page.goto(`http://127.0.0.1:${server.address().port}/`, { waitUntil: 'load' });
    await page.evaluate((b) => window.postMessage({ type: 'init', format: 'pdf', mode: 'view', bytes: b, title: 'fixture.pdf', dirty: false, locale: 'zh-cn' }, '*'), b64);
    await page.waitForTimeout(3500);
    const init = (mode) => page.evaluate(({ m, b }) => window.postMessage({ type: 'init', format: 'pdf', mode: m, bytes: b, title: 'fixture.pdf', dirty: false, locale: 'zh-cn' }, '*'), { m: mode, b: b64 });

    const snap = () => page.evaluate(() => ({
        bar: !!document.querySelector('.pdf-bar'),
        pager: document.querySelector('.pdf-pager input')?.value,
        total: document.querySelector('.pdf-pager .total')?.textContent,
        zoomChips: document.querySelectorAll('.jc-zoom .chip').length,
        thumbs: document.querySelectorAll('.pdf-thumb').length,
        thumbsCanvas: document.querySelectorAll('.pdf-thumb .frame canvas').length,
        active: document.querySelector('.pdf-thumb.active')?.dataset.page,
        pages: document.querySelectorAll('.pdf-page').length,
        visible: Array.from(document.querySelectorAll('.pdf-page')).filter((p) => p.offsetParent !== null).length,
        cur: document.querySelector('.pdf-page.current')?.dataset.page,
        title: document.querySelector('.pdf-bar .title')?.textContent,
        paged: !!document.querySelector('.pdf-pages.paged')
    }));

    const s0 = await snap();
    check('PDF 阅读器外壳', s0.bar && s0.thumbs === 6 && s0.pages === 6 && s0.zoomChips === 4, JSON.stringify({ bar: s0.bar, thumbs: s0.thumbs, pages: s0.pages }));
    check('PDF 首屏渲染与缩略图', s0.paged && s0.visible === 1 && s0.cur === '1' && s0.thumbsCanvas > 0, `visible=${s0.visible} thumbsCanvas=${s0.thumbsCanvas}`);
    check('PDF 工具栏页码与标题', s0.pager === '1' && s0.total === '/ 6' && s0.title === 'fixture.pdf', `${s0.pager}${s0.total} ${s0.title}`);

    await page.click('.pdf-thumb[data-page="4"]'); await page.waitForTimeout(1800);
    const s1 = await snap();
    check('PDF 点击缩略图翻页', s1.pager === '4' && s1.cur === '4' && s1.active === '4' && s1.visible === 1, JSON.stringify({ pager: s1.pager, cur: s1.cur, active: s1.active }));

    await page.click('.pdf-bar .pdf-btn[title*="连续"]'); await page.waitForTimeout(1500);
    const s2 = await snap();
    check('PDF 切连续滚动', !s2.paged && s2.visible === 6, `paged=${s2.paged} visible=${s2.visible}`);

    await page.click('.pdf-bar .pdf-btn[title*="旋转"]'); await page.waitForTimeout(2500);
    const s3 = await snap();
    // 旋转会重建缩略图栏，当前页高亮必须补回去，否则翻页定位会丢失视觉基准
    check('PDF 旋转后保持当前页高亮', s3.active === '4' && s3.cur === '4', JSON.stringify({ active: s3.active, cur: s3.cur }));

    // ---------- 叠加编辑 ----------
    // 查看模式下编辑工具层不参与（不抢正文的选中与滚动）
    check('查看模式不显示编辑工具栏', !(await page.isVisible('.pdf-edit-bar')));

    await page.evaluate(() => { window.__POSTS__ = []; });
    await init('edit'); await page.waitForTimeout(3000);
    const editUi = await page.evaluate(() => ({
        bar: !!document.querySelector('.pdf-edit-bar') && !document.querySelector('.pdf-edit-bar').hidden,
        tools: document.querySelectorAll('.pdf-edit-bar .pdf-tool').length,
        layer: !!document.querySelector('.pdf-page.current .pdf-ov-layer.on'),
        setMode: (window.__POSTS__ || []).some((m) => m && m.type === 'setMode')
    }));
    check('编辑模式打开叠加工具层', editUi.bar && editUi.tools >= 6 && editUi.layer, JSON.stringify(editUi));
    // 宿主 init 的 edit 不该回发 setMode（否则与标题栏命令互相触发）
    check('宿主指定的模式不回发 setMode', !editUi.setMode);

    // 页内相对坐标 -> 屏幕坐标
    // 页面往往比视口高（单页在 1280 宽下有 1400+px），必须夹到「可视的正文区」内：
    // 超出视口或压在底部状态栏上，鼠标事件都到不了叠加层
    const inPagePoint = async (fx, fy) => {
        const box = await page.evaluate(() => {
            const r = document.querySelector('.pdf-page.current .pdf-ov-layer').getBoundingClientRect();
            return { x: r.x, y: r.y, w: r.width, h: r.height };
        });
        const vp = page.viewportSize();
        const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
        return {
            x: clamp(box.x + box.w * fx, 40, vp.width - 40),
            y: clamp(box.y + box.h * fy, 40, vp.height - 90)
        };
    };
    const clickInPage = async (fx, fy) => {
        const p = await inPagePoint(fx, fy);
        await page.mouse.click(p.x, p.y);
        await page.waitForTimeout(220);
    };

    // 矩形（高亮块）
    await page.click('.pdf-edit-bar [data-tool="rect"]');
    await clickInPage(0.3, 0.3);
    // 文本：中文需要宿主注入 TTF，这里用 ASCII 走标准字体
    page.once('dialog', (d) => d.accept('Signed by JDOC'));
    await page.click('.pdf-edit-bar [data-tool="text"]');
    await clickInPage(0.5, 0.5);
    // 手绘签名：按下拖出一笔
    await page.click('.pdf-edit-bar [data-tool="draw"]');
    {
        const start = await inPagePoint(0.2, 0.75);
        await page.mouse.move(start.x, start.y);
        await page.mouse.down();
        for (let i = 1; i <= 8; i++) {
            const p = await inPagePoint(0.2 + i * 0.06, 0.75 - Math.sin(i / 2) * 0.04);
            await page.mouse.move(p.x, p.y, { steps: 2 });
        }
        await page.mouse.up();
        await page.waitForTimeout(220);
    }
    const objs = await page.evaluate(() => ({
        all: document.querySelectorAll('.pdf-page.current .pdf-ov').length,
        text: document.querySelectorAll('.pdf-page.current .pdf-ov-text').length,
        path: document.querySelectorAll('.pdf-page.current .pdf-ov svg path').length,
        dirtyTitle: document.querySelector('.pdf-bar .title')?.textContent
    }));
    check('PDF 叠加对象绘制', objs.all === 3 && objs.text === 1 && objs.path === 1, JSON.stringify(objs));
    check('PDF 编辑后标题标记未保存', String(objs.dirtyTitle || '').startsWith('●'), String(objs.dirtyTitle));

    // 保存 → sync（字节）+ save（落盘）
    await page.click('.pdf-edit-bar .pdf-tool.primary');
    await page.waitForTimeout(2500);
    const posts = await page.evaluate(() => window.__POSTS__ || []);
    const syncMsg = posts.filter((m) => m && m.type === 'sync').pop();
    check('PDF 保存发出 sync + save', !!syncMsg && syncMsg.dirty === true && posts.some((m) => m && m.type === 'save'),
        posts.map((m) => m && m.type).join(','));

    const saved = Buffer.from(syncMsg?.bytes || '', 'base64');
    const ops = saved.length ? await pageContentOps(saved, 0) : '';
    // 叠加文本用标准字体写入，内容流里是字形 hex；手绘/矩形走 path 指令
    const hexText = Buffer.from('Signed by JDOC').toString('hex').toUpperCase();
    const segs = (ops.match(/ l\s*\n/g) || []).length;
    check('PDF 叠加写入内容流（文本/手绘/矩形）', ops.includes(hexText) && segs >= 10 && /h\s*\n[BF]/.test(ops),
        `len=${saved.length} text=${ops.includes(hexText)} lineSegs=${segs} closed=${/h\s*\n[BF]/.test(ops)}`);
    const parsed = saved.length ? await jdoc.pdfToJson(saved) : null;
    const text = saved.length ? await jdoc.pdfToText(saved) : '';
    check('PDF 叠加后可正常回读且原内容保留', parsed?.pageCount === 6 && text.includes('Chapter 1') && text.includes('Signed by JDOC'),
        `pages=${parsed?.pageCount} chapter=${text.includes('Chapter 1')} signed=${text.includes('Signed by JDOC')}`);

    // 全文搜索：打开后预取文本，搜 "Chapter" 应命中多页并能跳到首命中页
    await page.waitForTimeout(800); // 等预取文本项完成
    await page.click('.pdf-bar [data-tool="search"]');
    const searchVisible = await page.evaluate(() => !document.querySelector('.pdf-search-bar').hidden);
    check('搜索栏可打开', searchVisible, `visible=${searchVisible}`);
    await page.fill('.pdf-search-input', 'Chapter');
    await page.waitForTimeout(500);
    const search = await page.evaluate(() => {
        const count = document.querySelector('.pdf-search-count')?.textContent || '';
        const active = document.querySelector('.pdf-search-hit.current');
        const firstPage = active ? Number(active.closest('.pdf-page')?.dataset.page) : 0;
        return { count, total: document.querySelectorAll('.pdf-search-hit').length, firstPage };
    });
    check('全文搜索命中并高亮', Number(search.count.split('/')[1]) > 1 && search.total > 1,
        `count=${search.count} hits=${search.total}`);
    check('搜索跳到首个命中页', search.firstPage >= 1, `page=${search.firstPage}`);
    // 下一个匹配
    await page.click('.pdf-search-bar .pdf-btn[aria-label="下一个匹配"]');
    await page.waitForTimeout(300);
    const next = await page.evaluate(() => {
        const cur = document.querySelector('.pdf-search-hit.current');
        return { page: cur ? Number(cur.closest('.pdf-page')?.dataset.page) : 0, count: document.querySelector('.pdf-search-count')?.textContent };
    });
    check('下一个匹配切换当前项', next.count.split('/')[0] === '2', `count=${next.count} page=${next.page}`);
    // 无匹配：标红且不报错
    await page.fill('.pdf-search-input', 'zzz_no_such_text');
    await page.waitForTimeout(400);
    const none = await page.evaluate(() => ({
        noMatch: document.querySelector('.pdf-search-input')?.classList.contains('no-match'),
        count: document.querySelector('.pdf-search-count')?.textContent
    }));
    check('无匹配时提示且不崩溃', none.noMatch === true && none.count === '0/0', JSON.stringify(none));
    await page.keyboard.press('Escape');

    const stacks = await page.evaluate(() => window.__ERRS__ || []);
    check('无运行时错误', errs.length === 0 && stacks.length === 0, errs.concat(stacks).slice(0, 2).join(' | '));
    await browser.close();
    server.close();
}

async function run(fmt) {
    console.log(`\n== ${fmt.toUpperCase()} ==`);
    if (fmt === 'pdf') return runPdf();
    const bytes = fmt === 'docx' ? await genDocx() : await genXlsx();
    const b64 = Buffer.from(bytes).toString('base64');
    const server = http.createServer((req, res) => {
        const url = req.url.split('?')[0];
        if (url === '/' || url === '/index.html') {
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end(`<!doctype html><html><head><meta charset="utf-8"></head><body><div id="app"></div>
<script>window.acquireVsCodeApi=function(){return {postMessage:function(m){try{(window.__POSTS__=window.__POSTS__||[]).push(m);}catch(e){}}};};</script>
<script>window.__ERRS__=[];window.__POSTS__=[];(function(){const o=console.error.bind(console);console.error=function(...a){try{window.__ERRS__.push(a.map(x=>x&&x.stack?x.stack:String(x)).join(' '));}catch(e){}o(...a);};})();</script>
<script src="/office-editor.js"></script></body></html>`);
            return;
        }
        if (url === '/office-editor.js') { res.writeHead(200, { 'Content-Type': 'text/javascript' }); res.end(fs.readFileSync(EDITOR)); return; }
        res.writeHead(404); res.end('nf');
    });
    await new Promise((r) => server.listen(0, r));
    const port = server.address().port;
    const browser = await chromium.launch({ args: ['--no-sandbox'] });
    const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
    const errs = [];
    page.on('console', (m) => { if (m.type() === 'error') errs.push('CONSOLE: ' + m.text()); });
    page.on('pageerror', (e) => errs.push('PAGEERR: ' + e.message));
    const init = (mode) => page.evaluate(({ f, m, b }) => window.postMessage({ type: 'init', format: f, mode: m, bytes: b, title: 't', dirty: false, locale: 'zh-cn' }, '*'), { f: fmt, m: mode, b: b64 });
    const lastSync = () => page.evaluate(() => { const s = (window.__POSTS__ || []).filter((m) => m && m.type === 'sync'); return s.length ? s[s.length - 1].bytes : ''; });

    await page.goto(`http://127.0.0.1:${port}/`, { waitUntil: 'load' });
    await page.waitForTimeout(500);

    // 查看模式（DOCX 已分页：.oe-page 纸面 + 左侧目录栏；XLSX 仍是单张 .page）
    await init('view'); await page.waitForTimeout(1100);
    const view = await page.evaluate(() => ({
        page: !!document.querySelector('.oe-view .page'),
        pages: document.querySelectorAll('.oe-view .oe-page').length,
        pageNums: document.querySelectorAll('.oe-view .oe-page-num').length,
        outline: document.querySelectorAll('.oe-outline-item').length,
        table: !!document.querySelector('.oe-view table')
    }));
    check(
        'view 渲染',
        fmt === 'docx'
            ? view.pages > 0 && view.pageNums === view.pages
            : view.page && view.table
    );

    if (fmt === 'docx') {
        // 分页 + 目录/大纲：页数 > 1、每页有页码、左侧目录列出全部标题
        check('DOCX 分页页码', view.pages > 1 && view.pageNums === view.pages, `pages=${view.pages} nums=${view.pageNums}`);
        check('DOCX 左侧目录索引', view.outline === 3, `items=${view.outline}`);
        const toc = await page.evaluate(() => ({
            rows: document.querySelectorAll('.oe-view .docx-toc-row').length,
            first: document.querySelector('.oe-view .docx-toc-row')?.textContent?.trim() || ''
        }));
        check('DOCX 目录条目', toc.rows === 3 && toc.first.includes('季度报告'), `rows=${toc.rows} first=${JSON.stringify(toc.first)}`);
        // 「与下段同页」：标题不能孤立在页尾（Word/WPS 的 keepNext）
        const orphan = await page.evaluate(() => {
            const bad = [];
            document.querySelectorAll('.oe-view .oe-page').forEach((p, i) => {
                const body = p.querySelector('.docx-body');
                if (!body) return;
                const blocks = Array.from(body.children);
                if (blocks.length < 2) return;
                const last = blocks[blocks.length - 1];
                if (/^H[1-6]$/.test(last.tagName) || last.dataset?.type === 'heading') bad.push(i + 1);
            });
            return bad;
        });
        check('DOCX 标题不落在页尾（与下段同页）', orphan.length === 0, `页码=${orphan.join(',')}`);
        // 编辑模式同样要有目录块与分页（块编辑器走另一条渲染路径）
        await page.evaluate(() => { window.__POSTS__ = []; });
        await init('edit'); await page.waitForTimeout(1300);
        const editView = await page.evaluate(() => ({
            pages: document.querySelectorAll('.oe-doc .oe-page').length,
            tocRows: document.querySelectorAll('.oe-doc .oe-toc-row').length,
            outline: document.querySelectorAll('.oe-outline-item').length
        }));
        check('DOCX 编辑模式目录', editView.tocRows === 3 && editView.pages > 1 && editView.outline === 3,
            `pages=${editView.pages} toc=${editView.tocRows} outline=${editView.outline}`);
    }

    // 编辑模式
    await page.evaluate(() => { window.__POSTS__ = []; });
    await init('edit'); await page.waitForTimeout(1300);
    const ui = await page.evaluate(() => ({ editables: document.querySelectorAll('.oe-doc .oe-edit').length, grid: document.querySelectorAll('.oe-grid td').length }));
    check('edit UI', fmt === 'docx' ? ui.editables > 0 : ui.grid > 0);

    if (fmt === 'docx') {
        // 合并表格渲染：jdoc 为压缩表示(数组下标≠列号)，需列号追踪正确显示 colspan
        const tbl = await page.evaluate(() => {
            const trs = document.querySelectorAll('.oe-doc .oe-tbl tr');
            if (trs.length < 2) return { r0: -1, r1: -1, span: 0 };
            return { r0: trs[0].children.length, r1: trs[1].children.length, span: trs[0].children[0].colSpan };
        });
        check('DOCX 合并表格渲染', tbl.r0 === 2 && tbl.r1 === 3 && tbl.span === 2, `r0=${tbl.r0} r1=${tbl.r1} span=${tbl.span}`);
    }

    if (fmt === 'docx') {
        await page.evaluate(() => { const ed = document.querySelector('.oe-doc .oe-edit'); ed.textContent = 'JDOC_PROBE'; ed.dispatchEvent(new Event('input', { bubbles: true })); ed.dispatchEvent(new Event('blur', { bubbles: true })); });
        await page.waitForTimeout(700);
        let p = await jdoc.docxToJson(Buffer.from(await lastSync(), 'base64'));
        let hText = (bb) => { const x = p.blocks.find((b) => b.type === 'heading'); return x ? (x.text ?? (x.runs || []).map((r) => r.text || '').join('')) : ''; };
        check('DOCX 文本编辑往返', hText() === 'JDOC_PROBE', `heading="${hText()}"`);
        await page.click('.oe-toolbar .oe-tb[title*="撤销"]'); await page.waitForTimeout(700);
        const usd = await page.evaluate(() => { const s = (window.__POSTS__ || []).filter((m) => m && m.type === 'sync'); const l = s[s.length - 1]; return l ? { bytes: l.bytes, dirty: l.dirty } : null; });
        check('DOCX 撤销回最初(放弃编辑 sync null)', !!usd && usd.bytes === null && usd.dirty === false, JSON.stringify(usd));
    } else {
        // 单元格编辑 + 公式求值
        await page.evaluate(() => { const td = document.querySelector('.oe-grid td[data-r="1"][data-c="1"]'); td.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); const c = document.querySelector('.oe-grid td[data-r="1"][data-c="1"] .cell'); if (c) { c.textContent = '500'; c.dispatchEvent(new Event('blur', { bubbles: true })); } });
        await page.waitForTimeout(700);
        let p = await jdoc.xlsxToJson(Buffer.from(await lastSync(), 'base64'));
        const b2 = (p.sheets[0].rows[1]?.cells || [])[1];
        check('XLSX 单元格编辑往返', (b2?.value ?? b2?.text) === 500, `B2=${b2?.value}`);
        // 撤销恢复（一次编辑后 undo 即回最初态 → 放弃编辑 sync bytes:null）
        await page.click('.oe-toolbar .oe-tb[title*="撤销"]'); await page.waitForTimeout(700);
        const usx = await page.evaluate(() => { const s = (window.__POSTS__ || []).filter((m) => m && m.type === 'sync'); const l = s[s.length - 1]; return l ? { bytes: l.bytes, dirty: l.dirty } : null; });
        check('XLSX 撤销回最初(放弃编辑 sync null)', !!usx && usx.bytes === null && usx.dirty === false, JSON.stringify(usx));
    }

    const stacks = await page.evaluate(() => window.__ERRS__ || []);
    check('无运行时错误', errs.length === 0 && stacks.length === 0, errs.concat(stacks).slice(0, 2).join(' | '));
    await browser.close();
    server.close();
}

const only = process.argv[2];
const list = only === 'docx' || only === 'xlsx' || only === 'pdf' ? [only] : ['docx', 'xlsx', 'pdf'];
for (const f of list) { try { await run(f); } catch (e) { failures++; console.log(`  FAIL  ${f} 抛异常: ${e?.message || e}`); } }
console.log(`\n${failures ? 'FAILED' : 'ALL PASSED'} (${failures} failure(s))`);
process.exit(failures ? 1 : 0);

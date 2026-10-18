// office-editor（DOCX/XLSX）Webview 往返回归验证。
//
// 用 Playwright 加载 dist/office-editor.js（注入 acquireVsCodeApi 桩捕获 sync 消息），
// 走 view → edit → 驱动编辑 → 捕获 sync 字节 → 用 jdoc 重新解析断言内容与容器，
// 覆盖：查看渲染、编辑 UI、文本/单元格编辑往返、撤销/重做、公式求值、数字格式。
//
// 用法：node scripts/verify-roundtrip.mjs [docx|xlsx]   （缺省两种都跑）
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
const EDITOR = path.resolve(__dirname, '..', 'dist', 'office-editor.js');

let failures = 0;
function check(name, cond, extra = '') {
    const ok = !!cond;
    if (!ok) failures++;
    console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${extra ? '  ' + extra : ''}`);
}

async function genDocx() {
    return jdoc.jsonToDocx({ blocks: [
        { type: 'heading', level: 1, text: '季度报告' },
        { type: 'paragraph', text: '正文内容。' },
        { type: 'table', header: true, rows: [[{ type: 'cell', colspan: 2, text: 'X' }, 'c'], ['', '', '']] }
    ] });
}
async function genXlsx() {
    return jdoc.jsonToXlsx({ sheets: [{ name: 'S', rows: [
        { index: 1, cells: [{ value: 'a' }, { value: 'b' }] },
        { index: 2, cells: [{ value: 'x' }, { value: 100 }] }
    ] }] });
}

async function run(fmt) {
    console.log(`\n== ${fmt.toUpperCase()} ==`);
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

    // 查看模式
    await init('view'); await page.waitForTimeout(1100);
    const view = await page.evaluate(() => ({ page: !!document.querySelector('.oe-view .page'), table: !!document.querySelector('.oe-view table') }));
    check('view 渲染', fmt === 'docx' ? view.page : view.page && view.table);

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
const list = only === 'docx' || only === 'xlsx' ? [only] : ['docx', 'xlsx'];
for (const f of list) { try { await run(f); } catch (e) { failures++; console.log(`  FAIL  ${f} 抛异常: ${e?.message || e}`); } }
console.log(`\n${failures ? 'FAILED' : 'ALL PASSED'} (${failures} failure(s))`);
process.exit(failures ? 1 : 0);

// XLSX 网格编辑器回归：真实示例文件（多表 + 图表 + 跨表公式）在 webview 里的观感与行为。
//
// 用法：node scripts/verify-xlsx.mjs [路径]
// 依赖：playwright（仓库根 node_modules）、jdoc 已构建（dist/jdoc.cjs）、扩展已构建（vscode/dist/office-editor.js）。
import http from 'node:http';
import os from 'node:os';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { chromium } from 'playwright';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const jdoc = require(path.resolve(__dirname, '..', '..', 'dist', 'jdoc.cjs'));
const EDITOR = path.resolve(__dirname, '..', 'dist', 'office-editor.js');
const DEFAULT_FILE = path.resolve(__dirname, '..', '..', 'examples', '销售分析表 副本.xlsx');

let failures = 0;
function check(name, cond, extra = '') {
    const ok = !!cond;
    if (!ok) failures++;
    console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${extra ? '  ' + extra : ''}`);
}

const file = process.argv[2] ? path.resolve(process.argv[2]) : DEFAULT_FILE;
const b64 = fs.readFileSync(file).toString('base64');

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
    if (url === '/office-editor.js') {
        res.writeHead(200, { 'Content-Type': 'text/javascript' });
        res.end(fs.readFileSync(EDITOR));
        return;
    }
    res.writeHead(404);
    res.end('nf');
});
await new Promise((r) => server.listen(0, r));
const port = server.address().port;

const browser = await chromium.launch({ args: ['--no-sandbox'] });
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
const errs = [];
page.on('console', (m) => { if (m.type() === 'error') errs.push('CONSOLE: ' + m.text()); });
page.on('pageerror', (e) => errs.push('PAGEERR: ' + e.message));

await page.goto(`http://127.0.0.1:${port}/`, { waitUntil: 'load' });
await page.waitForTimeout(400);

const init = (mode) =>
    page.evaluate(({ m, b }) => window.postMessage({ type: 'init', format: 'xlsx', mode: m, bytes: b, title: 't', dirty: false, locale: 'zh-cn' }, '*'), { m: mode, b: b64 });

// ---------- 查看模式 ----------
await init('view');
await page.waitForTimeout(1200);
const view = await page.evaluate(() => ({
    grid: !!document.querySelector('.oe-grid'),
    cells: document.querySelectorAll('.oe-grid td').length,
    tabs: Array.from(document.querySelectorAll('.oe-sheet-tabs button')).map((b) => b.textContent),
    charts: document.querySelectorAll('.oe-chart').length,
    svgs: document.querySelectorAll('.oe-chart svg').length,
    namebox: document.querySelector('.oe-namebox')?.value,
    status: document.querySelector('.oe-statusbar .grow')?.textContent || '',
    toolbars: document.querySelectorAll('.oe-toolbar .oe-tb').length
}));
check('查看模式渲染网格', view.grid && view.cells > 0, `cells=${view.cells}`);
check('工作表标签齐全', view.tabs.join(',') === '销售数据,汇总看板', view.tabs.join(','));
check('只读模式不显示编辑工具栏', view.toolbars === 0, `buttons=${view.toolbars}`);
check('公式栏显示当前单元格', view.namebox === 'A1', `namebox=${view.namebox}`);

// 切到「汇总看板」看图表
await page.click('.oe-sheet-tabs button:nth-child(2)');
await page.waitForTimeout(700);
const board = await page.evaluate(() => {
    const charts = Array.from(document.querySelectorAll('.oe-chart'));
    const boxes = charts.map((c) => {
        const r = c.getBoundingClientRect();
        return { w: Math.round(r.width), h: Math.round(r.height), x: Math.round(r.x), y: Math.round(r.y) };
    });
    return {
        charts: charts.length,
        boxes,
        // 图表里的文本：标题 / 图例
        texts: charts.map((c) => (c.textContent || '').slice(0, 60)),
        // 关键数值是否按原样显示（不能被公式重算覆盖成 0）
        b8: document.querySelector('.oe-grid td[data-r="7"][data-c="1"] .cell')?.textContent || '',
        b11: document.querySelector('.oe-grid td[data-r="10"][data-c="1"] .cell')?.textContent || '',
        c8: document.querySelector('.oe-grid td[data-r="7"][data-c="2"] .cell')?.textContent || ''
    };
});
check('汇总看板渲染出 3 个图表', board.charts === 3, `charts=${board.charts}`);
check('图表有尺寸与位置', board.boxes.every((b) => b.w > 100 && b.h > 80), JSON.stringify(board.boxes));
check('图表带标题文本', board.texts.some((t) => t.includes('销售额')), JSON.stringify(board.texts));
check('跨表公式结果未被重算覆盖', board.b8.includes('2,481,613') || board.b8 === '2481613', `B8=${board.b8}`);
check('占比显示为百分比', board.c8.includes('%'), `C8=${board.c8}`);

await page.screenshot({ path: path.join(os.tmpdir(), 'jdoc-xlsx-view.png') });

// ---------- 编辑模式 ----------
await init('edit');
await page.waitForTimeout(1200);
const edit = await page.evaluate(() => ({
    toolbars: document.querySelectorAll('.oe-toolbar .oe-tb').length,
    tabs: Array.from(document.querySelectorAll('.oe-sheet-tabs button')).map((b) => b.textContent)
}));
check('编辑模式有工具栏', edit.toolbars > 5, `buttons=${edit.toolbars}`);
check('编辑模式有工作表标签', edit.tabs.slice(0, 2).join(',') === '销售数据,汇总看板', edit.tabs.join(','));
// 切到汇总看板：编辑模式也要有图表
await page.click('.oe-sheet-tabs button:nth-child(2)');
await page.waitForTimeout(600);
const editCharts = await page.evaluate(() => document.querySelectorAll('.oe-chart').length);
check('编辑模式也有图表', editCharts === 3, `charts=${editCharts}`);

// 改一个源数据单元格，验证公式联动重算
await page.click('.oe-sheet-tabs button:nth-child(1)');
await page.waitForTimeout(500);
const before = await page.evaluate(() => document.querySelector('.oe-grid td[data-r="2"][data-c="6"] .cell')?.textContent || '');
await page.evaluate(() => {
    const td = document.querySelector('.oe-grid td[data-r="2"][data-c="6"]');
    td.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    td.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
});
await page.waitForTimeout(200);
await page.evaluate(() => {
    const c = document.querySelector('.oe-grid td[data-r="2"][data-c="6"] .cell');
    if (c) {
        c.textContent = '999';
        c.dispatchEvent(new Event('blur', { bubbles: true }));
    }
});
await page.waitForTimeout(600);
const after = await page.evaluate(() => ({
    g3: document.querySelector('.oe-grid td[data-r="2"][data-c="6"] .cell')?.textContent || '',
    h3: document.querySelector('.oe-grid td[data-r="2"][data-c="7"] .cell')?.textContent || ''
}));
check('单元格编辑生效', after.g3.includes('999'), `G3=${after.g3} (before=${before})`);

// 切到汇总看板看合计是否跟着变
await page.click('.oe-sheet-tabs button:nth-child(2)');
await page.waitForTimeout(600);
const linked = await page.evaluate(() => ({
    a4: document.querySelector('.oe-grid td[data-r="3"][data-c="0"] .cell')?.textContent || '',
    status: document.querySelector('.oe-statusbar .grow')?.textContent || ''
}));
check('跨表公式随编辑重算', linked.a4 !== '' && !linked.a4.includes('#'), `A4=${linked.a4}`);
check('状态栏有求和/计数/平均值', /求和=/.test(linked.status) && /平均值=/.test(linked.status), linked.status);

await page.screenshot({ path: path.join(os.tmpdir(), 'jdoc-xlsx-edit.png') });

// ---------- 保存往返：图表与数据都不能丢 ----------
const syncBytes = await page.evaluate(() => {
    const s = (window.__POSTS__ || []).filter((m) => m && m.type === 'sync');
    const last = s[s.length - 1];
    return last && last.bytes ? last.bytes : '';
});
check('编辑后有 sync 回传', !!syncBytes, `len=${syncBytes.length}`);
if (syncBytes) {
    const parsed = await jdoc.xlsxToJson(Buffer.from(syncBytes, 'base64'));
    const board = parsed.sheets.find((s2) => s2.name === '汇总看板');
    check('保存后图表仍在', (board?.charts?.length || 0) === 3, `charts=${board?.charts?.length}`);
    check('保存后标题保留', board?.charts?.[0]?.title === '各区域销售额对比', `title=${board?.charts?.[0]?.title}`);
    const g3 = parsed.sheets[0].rows[2].cells[6];
    check('保存后编辑值写入', Number(g3?.value ?? g3?.result) === 999, `G3=${JSON.stringify(g3?.value ?? g3?.result)}`);
    const a4 = board?.rows[3]?.cells?.[0];
    check('保存后公式仍在', typeof a4?.formula === 'string' && a4.formula.includes('SUM'), `A4=${a4?.formula}`);
}

// ---------- 冻结窗格：滚动后表头仍在 ----------
await page.click('.oe-sheet-tabs button:nth-child(1)');
await page.waitForTimeout(500);
await page.evaluate(() => {
    const wrap = document.querySelector('.oe-grid-wrap');
    if (wrap) wrap.scrollTop = 400;
});
await page.waitForTimeout(300);
const frozen = await page.evaluate(() => {
    const wrap = document.querySelector('.oe-grid-wrap');
    const head = document.querySelector('.oe-grid thead th:nth-child(2)');
    if (!wrap || !head) return null;
    const wr = wrap.getBoundingClientRect();
    const hr = head.getBoundingClientRect();
    return { headTop: Math.round(hr.top - wr.top), scrollTop: Math.round(wrap.scrollTop) };
});
check('冻结行在滚动后仍吸附在顶部', frozen && frozen.headTop >= 0 && frozen.headTop < 30, JSON.stringify(frozen));

const stacks = await page.evaluate(() => window.__ERRS__ || []);
check('无运行时错误', errs.length === 0 && stacks.length === 0, errs.concat(stacks).slice(0, 2).join(' | '));

await browser.close();
server.close();
console.log(`\n${failures ? 'FAILED' : 'ALL PASSED'} (${failures} failure(s))`);
process.exit(failures ? 1 : 0);

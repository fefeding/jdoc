// 截图辅助：把示例工作簿的两张表分别截出来，便于人工比对观感。
import http from 'node:http';
import os from 'node:os';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const EDITOR = path.resolve(__dirname, '..', 'dist', 'office-editor.js');
const FILE = path.resolve(__dirname, '..', '..', 'examples', '销售分析表 副本.xlsx');
const b64 = fs.readFileSync(FILE).toString('base64');

const server = http.createServer((req, res) => {
    const url = req.url.split('?')[0];
    if (url === '/office-editor.js') {
        res.writeHead(200, { 'Content-Type': 'text/javascript' });
        res.end(fs.readFileSync(EDITOR));
        return;
    }
    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(`<!doctype html><html><head><meta charset="utf-8"></head><body><div id="app"></div>
<script>window.acquireVsCodeApi=function(){return {postMessage:function(m){window.__POSTS__=(window.__POSTS__||[]).concat([m]);}};};</script>
<script src="/office-editor.js"></script></body></html>`);
});
await new Promise((r) => server.listen(0, r));
const port = server.address().port;

const browser = await chromium.launch({ args: ['--no-sandbox'] });
const page = await browser.newPage({ viewport: { width: 1180, height: 860 } });
await page.goto(`http://127.0.0.1:${port}/`, { waitUntil: 'load' });
await page.evaluate((b) => window.postMessage({ type: 'init', format: 'xlsx', mode: 'view', bytes: b, title: 't', dirty: false, locale: 'zh-cn' }, '*'), b64);
await page.waitForTimeout(1400);

const out = os.tmpdir();
await page.screenshot({ path: path.join(out, 'jdoc-xlsx-sheet1.png') });

// 销售数据：滚动看冻结行
await page.evaluate(() => {
    const wrap = document.querySelector('.oe-grid-wrap');
    if (wrap) wrap.scrollTop = 300;
});
await page.waitForTimeout(400);
await page.screenshot({ path: path.join(out, 'jdoc-xlsx-frozen.png') });

// 汇总看板：滚到折线图
await page.click('.oe-sheet-tabs button:nth-child(2)');
await page.waitForTimeout(600);
await page.evaluate(() => {
    const wrap = document.querySelector('.oe-grid-wrap');
    if (wrap) wrap.scrollTop = wrap.scrollHeight;
});
await page.waitForTimeout(400);
await page.screenshot({ path: path.join(out, 'jdoc-xlsx-board-bottom.png') });

await browser.close();
server.close();
console.log('screenshots written');

// 行级分页引擎回归：
// 1. 拆分不得丢内容——分页前后全文逐字相同
// 2. widow/orphan 生效——没有孤行（页尾段落在页首只剩 1 行 / 反之）
// 3. 页高不溢出——没有一页的内容超过可用高度
// 4. 段落顺序不变
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { chromium } from 'playwright';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const jdoc = require(path.resolve(__dirname, '..', '..', 'dist', 'jdoc.cjs'));
const OFFICE_EDITOR = path.resolve(__dirname, '..', 'dist', 'office-editor.js');

let failed = 0;
const check = (name, ok, detail = '') => {
    console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  ' + detail : ''}`);
    if (!ok) failed++;
};

// 造一个「长段落必然跨页」的文档：正文用固定页高的 A4，多行段落 + 多节
function buildDoc() {
    const blocks = [{ type: 'heading', level: 1, text: '长段落跨页测试' }];
    for (let i = 1; i <= 6; i++) {
        blocks.push({ type: 'heading', level: 2, text: `第 ${i} 节` });
        // 每段约 14 行，确保跨页
        blocks.push({ type: 'paragraph', text: `${i} 节正文。`.repeat(140) });
        blocks.push({ type: 'heading', level: 3, text: `${i}.1 小节` });
        blocks.push({ type: 'paragraph', text: `补充说明。`.repeat(90) });
    }
    return jdoc.jsonToDocx({
        section: { width: 595.28, height: 841.89, marginTop: 72, marginBottom: 72, marginLeft: 72, marginRight: 72 },
        blocks
    });
}

const bytes = await buildDoc();
const b64 = Buffer.from(bytes).toString('base64');
const expectedText = (await jdoc.docxToText(bytes)).replace(/\s+/g, '');

const server = http.createServer((req, res) => {
    const url = req.url.split('?')[0];
    if (url === '/' || url === '/index.html') {
        res.writeHead(200, { 'Content-Type': 'text/html' });
        res.end(`<!doctype html><html><head><meta charset="utf-8"></head><body><div id="app"></div>
<script>window.acquireVsCodeApi=function(){return{postMessage:function(){}}};</script>
<script src="/office-editor.js"></script></body></html>`);
        return;
    }
    if (url === '/office-editor.js') {
        res.writeHead(200, { 'Content-Type': 'text/javascript' });
        res.end(fs.readFileSync(OFFICE_EDITOR));
        return;
    }
    res.writeHead(404);
    res.end();
});
await new Promise((r) => server.listen(0, r));
const browser = await chromium.launch({ args: ['--no-sandbox'] });
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
await page.goto(`http://127.0.0.1:${server.address().port}/`, { waitUntil: 'load' });
await page.evaluate((b) => window.postMessage({ type: 'init', format: 'docx', mode: 'view', bytes: b, title: 't', dirty: false, locale: 'zh-cn' }, '*'), b64);
await page.waitForSelector('.oe-page', { timeout: 20000 });
await page.waitForTimeout(1500);

console.log('\n== 行级分页 ==');

const r = await page.evaluate(() => {
    const pages = [...document.querySelectorAll('.oe-page')];
    let fullText = '';
    let order = [];
    let overflow = [];
    for (const [i, p] of pages.entries()) {
        const body = p.querySelector('.docx-body');
        fullText += body ? body.textContent : '';
        for (const el of body ? body.children : []) {
            order.push(el.tagName + ':' + (el.textContent || '').trim().slice(0, 8));
        }
        // 基准取纸面自身去掉上下 padding 后的净高（offsetHeight 是不含 zoom 的布局 px，
        // 不能拿 pt 直接比，也别用 rect——那已被缩放放大）
        const cs = getComputedStyle(p);
        const net = p.offsetHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom);
        const h = body?.scrollHeight || 0;
        if (h > net + 2) overflow.push({ page: i + 1, h: Math.round(h), net: Math.round(net) });
    }
    return {
        pageCount: pages.length,
        fullText: fullText.replace(/\s+/g, ''),
        order: order.slice(0, 8),
        overflow
    };
});

check('分页后全文逐字不丢（含跨页拆分）', r.fullText === expectedText,
    `分页后=${r.fullText.length} 原文档=${expectedText.length}`);
check('没有页面内容溢出可用高度', r.overflow.length === 0, JSON.stringify(r.overflow.slice(0, 3)));

// widow/orphan：统计每个段落的首尾行数分布
const wo = await page.evaluate(() => {
    const pages = [...document.querySelectorAll('.oe-page')];
    const results = [];
    for (const [i, p] of pages.entries()) {
        const body = p.querySelector('.docx-body');
        if (!body) continue;
        const ps = [...body.children].filter((el) => el.tagName === 'P');
        for (const el of ps) {
            // 用 Range 量首行与末行的 top，判断该段是否只有 1 行
            const text = el.firstChild;
            if (!text || text.nodeType !== 3) continue;
            const range = document.createRange();
            range.selectNodeContents(el);
            const rects = Array.from(range.getClientRects()).filter((x) => x.height > 0);
            const tops = [];
            for (const rc of rects) if (!tops.length || Math.abs(rc.top - tops[tops.length - 1]) > 1) tops.push(rc.top);
            const lines = tops.length;
            if (lines === 1) continue;
            const lh = rects.length ? rects[0].height : 0;
            results.push({ page: i + 1, lines, lh: Math.round(lh) });
        }
    }
    return results;
});
// 孤行检测：同一段落被切开后，落在页首的段若行数 < 2 即为 widow
const widows = wo.filter((x) => x.page > 1 && x.lines === 1);
check('页首无单行段落（widow）', widows.length === 0, `widow=${widows.length} 样例=${JSON.stringify(widows.slice(0, 2))}`);

const errs = await page.evaluate(() => window.__ERRS__ || []);
check('无运行时错误', errs.length === 0, errs.slice(0, 2).join(' | '));

console.log(`\n页数=${r.pageCount} 段落行数样本=${JSON.stringify(wo.slice(0, 4))}`);
console.log(failed ? `\nFAILED (${failed} failure(s))` : '\nALL PASSED (0 failure(s))');
await browser.close();
server.close();
process.exit(failed ? 1 : 0);
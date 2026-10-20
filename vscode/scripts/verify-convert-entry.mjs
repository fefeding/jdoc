// 验证：内容区右键菜单（转换格式）——三条 webview 链路
//   1. office-editor（DOCX / XLSX）
//   2. 顶层 webview（PDF / Markdown / TXT）
//   3. pptx-editor（PPTX，画布右键菜单）
// 断言：右键能弹出自定义菜单、菜单里有「转换格式」、点击后回传 convert 且带上文档 uri。
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
const WEBVIEW = path.resolve(__dirname, '..', 'dist', 'webview.js');
const PPTX_EDITOR = path.resolve(__dirname, '..', 'dist', 'pptx-editor.js');

let failed = 0;
function check(name, ok, detail = '') {
    console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  ' + detail : ''}`);
    if (!ok) failed++;
}

function makeServer(script) {
    return http.createServer((req, res) => {
        const url = req.url.split('?')[0];
        if (url === '/' || url === '/index.html') {
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end(`<!doctype html><html><head><meta charset="utf-8"></head><body><div id="app"></div>
<script>window.__POSTS__=[];window.acquireVsCodeApi=function(){return {postMessage:function(m){window.__POSTS__.push(m);}};};</script>
<script src="/app.js"></script></body></html>`);
            return;
        }
        if (url === '/app.js') {
            res.writeHead(200, { 'Content-Type': 'text/javascript' });
            res.end(fs.readFileSync(script));
            return;
        }
        res.writeHead(404);
        res.end();
    });
}

const docxB64 = Buffer.from(
    await jdoc.jsonToDocx({
        section: { width: 595.28, height: 841.89, marginTop: 72, marginBottom: 72, marginLeft: 72, marginRight: 72 },
        blocks: [
            { type: 'heading', level: 1, text: '测试文档' },
            ...Array.from({ length: 30 }, (_, i) => ({ type: 'paragraph', text: `第 ${i + 1} 段正文内容。`.repeat(10) }))
        ]
    })
).toString('base64');
const mdB64 = Buffer.from('# 标题\n\n正文内容。\n', 'utf8').toString('base64');
const pptxB64 = Buffer.from(
    await jdoc.jsonToPptx({
        slides: [{ title: '第一页', bullets: ['要点一', '要点二'] }]
    })
).toString('base64');

const browser = await chromium.launch({ args: ['--no-sandbox'] });

/** 在页面里右键并返回菜单信息 */
async function rightClickMenu(page, selector) {
    return page.evaluate(async (sel) => {
        const target = sel ? document.querySelector(sel) : document.body;
        if (!target) return { err: 'no target' };
        const r = target.getBoundingClientRect();
        const x = Math.round(r.left + Math.min(80, r.width / 2));
        const y = Math.round(r.top + Math.min(40, r.height / 2));
        target.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: x, clientY: y }));
        await new Promise((r2) => setTimeout(r2, 200));
        const menu = document.querySelector('.jc-ctxmenu');
        if (!menu) return { err: 'no menu', hasConvert: false };
        const items = [...menu.querySelectorAll('.jc-ctxmenu-item')].map((i) => i.textContent.replace(/⌘C$/, '').trim());
        return { items, hasConvert: items.some((t) => t.includes('转换格式')) };
    }, selector);
}

async function clickConvert(page) {
    return page.evaluate(async () => {
        window.__POSTS__.length = 0;
        const item = [...document.querySelectorAll('.jc-ctxmenu-item')].find((i) => i.textContent.includes('转换格式'));
        if (!item) return null;
        item.click();
        await new Promise((r) => setTimeout(r, 250));
        const msgs = window.__POSTS__.filter((m) => m && m.type === 'convert');
        return { count: msgs.length, uri: msgs[0]?.uri ?? null };
    });
}

// ---- 1. office-editor（DOCX）----
{
    console.log('\n== 内容区右键：DOCX（office-editor）==');
    const server = makeServer(OFFICE_EDITOR);
    const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
    await new Promise((r) => server.listen(0, r));
    await page.goto(`http://127.0.0.1:${server.address().port}/`, { waitUntil: 'load' });
    await page.evaluate(
        (b) => window.postMessage({ type: 'init', format: 'docx', mode: 'view', bytes: b, title: 't.docx', uri: 'file:///tmp/t.docx', dirty: false, locale: 'zh-cn' }, '*'),
        docxB64
    );
    await page.waitForSelector('.oe-page', { timeout: 20000 });
    await page.waitForTimeout(800);

    const menu = await rightClickMenu(page, '.oe-page');
    check('DOCX 内容区右键弹出自定义菜单', !menu.err, menu.err || `items=${JSON.stringify(menu.items)}`);
    check('DOCX 菜单含「转换格式」', menu.hasConvert === true, `items=${JSON.stringify(menu.items)}`);

    const click = await clickConvert(page);
    check('DOCX 点击「转换格式」回传 convert 消息', click?.count === 1, `count=${click?.count}`);
    check('DOCX convert 消息带上文档 uri', click?.uri === 'file:///tmp/t.docx', `uri=${click?.uri}`);

    // 菜单应随再次右键/点击外部而更新，且不残留
    const closed = await page.evaluate(() => document.querySelectorAll('.jc-ctxmenu').length);
    check('DOCX 点击菜单项后菜单关闭', closed === 0, `残留=${closed}`);

    // 有选区时应同时提供「复制」，无选区时只给「转换格式」
    const withSel = await page.evaluate(async () => {
        const p = document.querySelector('.oe-page p, .oe-page h1');
        if (!p) return { err: 'no text' };
        const range = document.createRange();
        range.selectNodeContents(p);
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
        p.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 100, clientY: 100 }));
        await new Promise((r) => setTimeout(r, 200));
        const m = document.querySelector('.jc-ctxmenu');
        const items = m ? [...m.querySelectorAll('.jc-ctxmenu-item')].map((i) => i.textContent.trim()) : [];
        sel.removeAllRanges();
        return { items };
    });
    check(
        'DOCX 有选区时菜单含「复制」与「转换格式」',
        withSel.items?.some((t) => t.includes('复制')) && withSel.items?.some((t) => t.includes('转换格式')),
        `items=${JSON.stringify(withSel.items)}`
    );
    await page.close();
    server.close();
}

// ---- 2. 顶层 webview（Markdown）----
{
    console.log('\n== 内容区右键：Markdown（webview）==');
    const server = makeServer(WEBVIEW);
    const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
    await new Promise((r) => server.listen(0, r));
    await page.goto(`http://127.0.0.1:${server.address().port}/`, { waitUntil: 'load' });
    await page.evaluate(
        (b) => window.postMessage({ type: 'init', format: 'md', mode: 'view', bytes: b, title: 't.md', uri: 'file:///tmp/t.md', dirty: false, locale: 'zh-cn' }, '*'),
        mdB64
    );
    await page.waitForSelector('.jc-content', { timeout: 15000 });
    await page.waitForTimeout(600);

    const menu = await rightClickMenu(page, '.jc-content');
    check('Markdown 内容区右键弹出自定义菜单', !menu.err, menu.err || `items=${JSON.stringify(menu.items)}`);
    check('Markdown 菜单含「转换格式」', menu.hasConvert === true, `items=${JSON.stringify(menu.items)}`);

    const click = await clickConvert(page);
    check('Markdown 点击「转换格式」回传 convert 消息', click?.count === 1, `count=${click?.count}`);
    check('Markdown convert 消息带上文档 uri', click?.uri === 'file:///tmp/t.md', `uri=${click?.uri}`);
    await page.close();
    server.close();
}

// ---- 3. pptx-editor（PPTX）----
{
    console.log('\n== 内容区右键：PPTX（pptx-editor）==');
    const server = makeServer(PPTX_EDITOR);
    const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
    await new Promise((r) => server.listen(0, r));
    await page.goto(`http://127.0.0.1:${server.address().port}/`, { waitUntil: 'load' });
    await page.evaluate(
        (b) => window.postMessage({ type: 'init', format: 'pptx', mode: 'preview', bytes: b, title: 't.pptx', uri: 'file:///tmp/t.pptx', dirty: false, locale: 'zh-cn' }, '*'),
        pptxB64
    );
    await page.waitForSelector('.menu, .slide, .canvas', { timeout: 20000 });
    await page.waitForTimeout(1200);

    // PPTX 用自带的 menu.ts（.menu 类），不是 .jc-ctxmenu
    const menu = await page.evaluate(async () => {
        const target =
            document.querySelector('.canvas-area, .canvas, .slide-area, .stage') ||
            document.querySelector('.slide') ||
            document.body;
        const r = target.getBoundingClientRect();
        const x = Math.round(r.left + Math.min(60, r.width / 2));
        const y = Math.round(r.top + Math.min(30, r.height / 2));
        target.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: x, clientY: y }));
        await new Promise((res) => setTimeout(res, 250));
        const m = document.querySelector('.menu');
        if (!m) return { err: 'no menu' };
        const items = [...m.querySelectorAll('.menu-item, [role="menuitem"], div')].map((i) => (i.textContent || '').trim()).filter(Boolean);
        return { items, text: m.textContent };
    });
    check('PPTX 画布右键弹出菜单', !menu.err, menu.err || `text=${JSON.stringify((menu.text || '').slice(0, 80))}`);
    check('PPTX 菜单含「转换格式」', (menu.text || '').includes('转换格式'), `text=${JSON.stringify((menu.text || '').slice(0, 120))}`);

    const click = await page.evaluate(async () => {
        window.__POSTS__.length = 0;
        const m = document.querySelector('.menu');
        if (!m) return null;
        // PPTX 菜单项是带 .menu-item 或可点击的 div
        const item = [...m.children].find((i) => (i.textContent || '').includes('转换格式'));
        if (!item) return null;
        item.click();
        await new Promise((r) => setTimeout(r, 250));
        const msgs = window.__POSTS__.filter((x) => x && x.type === 'convert');
        return { count: msgs.length, uri: msgs[0]?.uri ?? null };
    });
    check('PPTX 点击「转换格式」回传 convert 消息', click?.count === 1, `count=${click?.count}`);
    check('PPTX convert 消息带上文档 uri', click?.uri === 'file:///tmp/t.pptx', `uri=${click?.uri}`);
    await page.close();
    server.close();
}

console.log(failed ? `\nFAILED (${failed} failure(s))` : '\nALL PASSED (0 failure(s))');
await browser.close();
process.exit(failed ? 1 : 0);
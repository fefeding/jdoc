// 验证：webview 状态栏的「转换格式」入口存在且点击会回传 convert 消息
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

let failed = 0;
function check(name, ok, detail = '') {
    console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  ' + detail : ''}`);
    if (!ok) failed++;
}

const docxBytes = jdoc.jsonToDocx({
    section: { width: 595.28, height: 841.89, marginTop: 72, marginBottom: 72, marginLeft: 72, marginRight: 72 },
    blocks: [
        { type: 'heading', level: 1, text: '测试文档' },
        ...Array.from({ length: 30 }, (_, i) => ({ type: 'paragraph', text: `第 ${i + 1} 段正文内容。`.repeat(10) }))
    ]
});

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

const b64 = Buffer.from(await docxBytes).toString('base64');
const browser = await chromium.launch({ args: ['--no-sandbox'] });

// ---- office-editor（docx / xlsx）----
{
    const server = makeServer(OFFICE_EDITOR);
    const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
    await new Promise((r) => server.listen(0, r));
    await page.goto(`http://127.0.0.1:${server.address().port}/`, { waitUntil: 'load' });
    await page.evaluate(
        (b) => window.postMessage({ type: 'init', format: 'docx', mode: 'view', bytes: b, title: 't.docx', dirty: false, locale: 'zh-cn' }, '*'),
        b64
    );
    await page.waitForSelector('.oe-page', { timeout: 20000 });
    await page.waitForTimeout(800);

    console.log('\n== 转换入口（office-editor：DOCX/XLSX）==');
    const chip = await page.evaluate(() => {
        const b = [...document.querySelectorAll('.oe-statusbar .chip')].find((x) => x.textContent.includes('转换'));
        return b ? { text: b.textContent, title: b.title } : null;
    });
    check('DOCX 状态栏有「转换格式」按钮', !!chip, chip ? `title="${chip.title}"` : '未找到');

    const posted = await page.evaluate(async () => {
        const b = [...document.querySelectorAll('.oe-statusbar .chip')].find((x) => x.textContent.includes('转换'));
        if (!b) return null;
        window.__POSTS__.length = 0;
        b.click();
        await new Promise((r) => setTimeout(r, 200));
        return window.__POSTS__.filter((m) => m && m.type === 'convert').length;
    });
    check('点击后回传 convert 消息给宿主', posted === 1, `convert 消息数=${posted}`);
    await page.close();
    server.close();
}

// ---- 顶层 webview（pdf / md / txt）----
{
    const server = makeServer(WEBVIEW);
    const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
    await new Promise((r) => server.listen(0, r));
    await page.goto(`http://127.0.0.1:${server.address().port}/`, { waitUntil: 'load' });
    const mdB64 = Buffer.from('# 标题\n\n正文内容。\n', 'utf8').toString('base64');
    await page.evaluate(
        (b) => window.postMessage({ type: 'init', format: 'md', mode: 'view', bytes: b, title: 't.md', uri: 'file:///t.md', dirty: false, locale: 'zh-cn' }, '*'),
        mdB64
    );
    await page.waitForSelector('.jc-statusbar', { timeout: 15000 });
    await page.waitForTimeout(500);

    console.log('\n== 转换入口（webview：MD/TXT/PDF）==');
    const chip = await page.evaluate(() => {
        const b = [...document.querySelectorAll('.jc-statusbar .chip')].find((x) => x.textContent.includes('转换'));
        return b ? { text: b.textContent, title: b.title } : null;
    });
    check('Markdown 状态栏有「转换格式」按钮', !!chip, chip ? `title="${chip.title}"` : '未找到');

    const posted = await page.evaluate(async () => {
        const b = [...document.querySelectorAll('.jc-statusbar .chip')].find((x) => x.textContent.includes('转换'));
        if (!b) return null;
        window.__POSTS__.length = 0;
        b.click();
        await new Promise((r) => setTimeout(r, 200));
        return window.__POSTS__.filter((m) => m && m.type === 'convert').length;
    });
    check('点击后回传 convert 消息给宿主', posted === 1, `convert 消息数=${posted}`);
    await page.close();
    server.close();
}

console.log(failed ? `\nFAILED (${failed} failure(s))` : '\nALL PASSED (0 failure(s))');
await browser.close();
process.exit(failed ? 1 : 0);
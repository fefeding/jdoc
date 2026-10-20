// 栅格化回归验证：在真实浏览器里驱动 webview.js 的 rasterize 消息，
// 确认「源格式 → PDF → 每页 PNG」这条链路在webview 侧真的能出图。
//
// 为什么必须在浏览器里验：宿主端没有 canvas（pdfjs 的 Node 渲染依赖 @napi-rs/canvas，
// 约 27MB 平台相关原生二进制，无法随扩展分发），所以图片导出完全依赖 webview。
// 只验证编译通过毫无意义——这里断言的是真实产出的 PNG 字节与页数。
//
// 断言：
// 1. webview 脚本就绪后主动发出 ready（宿主的 init/rasterize 都靠它握手）
// 2. rasterizeDone 的页数 == PDF 实际页数
// 3. rasterizePage 每页都是合法 PNG，且尺寸 == 页面尺寸 × scale（允许 ±2px 舍入）
// 4. PNG 不是空白（页面铺了白底，抽样像素应为白色，证明确实渲染了内容）
//
// 用法：node scripts/verify-rasterize.mjs
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { chromium } from 'playwright';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const ROOT = path.resolve(__dirname, '..', '..');
const jdoc = require(path.join(ROOT, 'dist', 'jdoc.cjs'));
const WEBVIEW = path.join(__dirname, '..', 'dist', 'webview.js');
const MEDIA = path.join(__dirname, '..', 'media');

let failed = 0;
function check(name, ok, detail = '') {
    console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  ' + detail : ''}`);
    if (!ok) failed++;
}

/** 造一份多页 PDF（长中文正文会自动跨页） */
async function makePdf() {
    const md = ['# 栅格化验证'].concat(
        Array.from({ length: 60 }, (_, i) => `## 第 ${i + 1} 节\n\n` + `这是第 ${i + 1} 节的正文内容，用于把文档撑到多页。`.repeat(12))
    ).join('\n\n');
    return jdoc.convert(md, 'pdf', { from: 'md' });
}

const MIME = { '.js': 'text/javascript', '.mjs': 'text/javascript' };
const server = http.createServer((req, res) => {
    const url = req.url.split('?')[0];
    if (url === '/' || url === '/index.html') {
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(`<!doctype html><html><head><meta charset="utf-8"></head><body><div id="app"></div>
<script>
  window.__POSTS__ = [];
  window.__ERRS__ = [];
  window.__JDC_PDF_WORKER__ = '/pdf.worker.min.mjs';
  window.acquireVsCodeApi = function () {
    return { postMessage: function (m) { try { window.__POSTS__.push(m); } catch (e) {} } };
  };
  (function () {
    var o = console.error.bind(console);
    console.error = function () {
      try { window.__ERRS__.push(Array.prototype.map.call(arguments, function (x) { return x && x.stack ? x.stack : String(x); }).join(' ')); } catch (e) {}
      o.apply(null, arguments);
    };
  })();
</script>
<script src="/webview.js"></script></body></html>`);
        return;
    }
    const file = url === '/webview.js' ? WEBVIEW : path.join(MEDIA, url.replace(/^\//, ''));
    if (fs.existsSync(file) && fs.statSync(file).isFile()) {
        res.writeHead(200, { 'Content-Type': MIME[path.extname(file)] || 'application/octet-stream' });
        res.end(fs.readFileSync(file));
        return;
    }
    res.writeHead(404);
    res.end();
});

const browser = await chromium.launch({ args: ['--no-sandbox'] });
const page = await browser.newPage({ viewport: { width: 1200, height: 900 } });
await new Promise((r) => server.listen(0, r));
await page.goto(`http://127.0.0.1:${server.address().port}/`, { waitUntil: 'load' });

console.log('\n== RASTERIZE ==');

const pdf = await makePdf();
const expected = (await jdoc.pdfToJson(pdf, { text: false })).pageCount;
console.log(`  （样本PDF ${(pdf.length / 1024).toFixed(0)}KB，共 ${expected} 页）`);

// 1. 脚本就绪即握手
await page.waitForFunction(() => window.__POSTS__?.some((m) => m.type === 'ready'), null, { timeout: 15000 });
check('webview 就绪后主动发出 ready', true);

// 2. 发栅格化请求
await page.evaluate((bytes) => window.postMessage({ type: 'rasterize', bytes, scale: 2 }, '*'), Buffer.from(pdf).toString('base64'));
await page.waitForFunction(
    () => window.__POSTS__?.some((m) => m.type === 'rasterizeDone' || m.type === 'rasterizeError'),
    null,
    { timeout: 60000 }
);

const posts = await page.evaluate(() => window.__POSTS__);
const errs = await page.evaluate(() => window.__ERRS__);

const fail = posts.find((m) => m.type === 'rasterizeError');
check('栅格化未报错', !fail, fail ? fail.message : '');
if (fail) {
    console.log('\n  webview 错误：', (errs || []).slice(0, 3).join('\n  '));
}

const done = posts.find((m) => m.type === 'rasterizeDone');
check('收到 rasterizeDone', !!done, done ? `count=${done.count}` : '');
check('页数与 PDF 实际页数一致', done?.count === expected, `期望 ${expected}，实际 ${done?.count}`);

const pages = posts.filter((m) => m.type === 'rasterizePage').sort((a, b) => a.index - b.index);
check('逐页回传数量正确', pages.length === expected, `收到 ${pages.length} 页`);

// 3. 每页都是合法 PNG 且尺寸正确（A4 595.28×841.89pt，scale=2 → 约 1191×1684）
const inspected = pages.map((p) => {
    const b = Buffer.from(p.dataUrl.slice(p.dataUrl.indexOf(',') + 1), 'base64');
    return {
        index: p.index,
        size: b.length,
        // PNG 完整签名 8 字节：89 50 4E 47 0D 0A 1A 0A
        sig: b.subarray(0, 8).toString('hex'),
        w: b.readUInt32BE(16),
        h: b.readUInt32BE(20),
        b64: p.dataUrl
    };
});
const PNG_SIG = '89504e470d0a1a0a';
check('全部是合法 PNG（签名 89504e47…）', inspected.every((p) => p.sig === PNG_SIG), `${inspected.length} 页`);
const sized = inspected.every((p) => Math.abs(p.w - 1191) <= 2 && Math.abs(p.h - 1684) <= 2);
check(
    '尺寸 == A4 × scale=2',
    sized,
    inspected.map((p) => `${p.index}:${p.w}x${p.h}`).join(' ')
);
check('PNG 体积合理（非空白）', inspected.every((p) => p.size > 8000), `最小 ${Math.min(...inspected.map((p) => p.size))}B`);

// 4. 抽样第1 页中心像素：应为白色（证明铺了白底且渲染出内容，不是黑图）
const px = await page.evaluate(async (dataUrl) => {
    const img = new Image();
    img.src = dataUrl;
    await img.decode();
    const c = document.createElement('canvas');
    c.width = img.width; c.height = img.height;
    const ctx = c.getContext('2d');
    ctx.drawImage(img, 0, 0);
    const d = ctx.getImageData(Math.floor(img.width / 2), Math.floor(img.height / 2), 1, 1).data;
    return [d[0], d[1], d[2], d[3]];
}, inspected[0].b64);
check('首页中心像素为不透明白色', px[0] > 200 && px[1] > 200 && px[2] > 200 && px[3] === 255, `rgba(${px.join(',')})`);

await browser.close();
server.close();

console.log(`\n${failed === 0 ? '全部通过' : failed + ' 项失败'}\n`);
process.exit(failed === 0 ? 0 : 1);
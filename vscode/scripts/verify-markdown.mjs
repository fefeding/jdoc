// Markdown 预览三处问题的回归：
// 1. 内容外框不应再画 1px 描边（与 office-editor 纸面一致：仅阴影 + 圆角）
// 2. 围栏代码块应被高亮（token 有颜色），且源码不被破坏
// 3. 无 class 的裸代码块也应高亮；mermaid 不应被着色
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const WEBVIEW_JS = path.resolve(__dirname, '..', 'dist', 'webview.js');

let failed = 0;
const check = (name, ok, detail = '') => {
    console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  ' + detail : ''}`);
    if (!ok) failed++;
};

const MD = [
    '# 标题',
    '',
    '> 引用',
    '',
    '```python',
    'def greet(name: str) -> str:',
    '    # 打招呼',
    '    return f"hello {name}"',
    '```',
    '',
    '```ts',
    'export const answer: number = 42;',
    '```',
    '',
    '```',
    'plain text block',
    '```',
    '',
    '```mermaid',
    'graph TD; A-->B;',
    '```',
    ''
].join('\n');
const b64 = Buffer.from(MD, 'utf8').toString('base64');

const server = http.createServer((req, res) => {
    const url = req.url.split('?')[0];
    if (url === '/' || url === '/index.html') {
        res.writeHead(200, { 'Content-Type': 'text/html' });
        res.end(`<!doctype html><html><head><meta charset="utf-8"></head>
<body class="vscode-dark"><div id="app"></div>
<script>window.acquireVsCodeApi=function(){return{postMessage:function(){}}};</script>
<script src="/webview.js"></script></body></html>`);
        return;
    }
    if (url === '/webview.js') {
        res.writeHead(200, { 'Content-Type': 'text/javascript' });
        res.end(fs.readFileSync(WEBVIEW_JS));
        return;
    }
    res.writeHead(404);
    res.end();
});
await new Promise((r) => server.listen(0, r));
const browser = await chromium.launch({ args: ['--no-sandbox'] });
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
await page.goto(`http://127.0.0.1:${server.address().port}/`, { waitUntil: 'load' });
await page.evaluate((b) => window.postMessage({ type: 'init', format: 'md', mode: 'view', bytes: b, title: 't', dirty: false, locale: 'zh-cn' }, '*'), b64);
await page.waitForSelector('.md-body', { timeout: 20000 });
await page.waitForTimeout(1200);

console.log('\n== Markdown 预览 ==');

const r = await page.evaluate(() => {
    const body = document.querySelector('.md-body');
    const bs = getComputedStyle(body);
    const codes = [...body.querySelectorAll('pre > code')];
    const colorOf = (el, sel) => {
        const n = el.querySelector(sel);
        return n ? getComputedStyle(n).color : null;
    };
    const py = codes.find((c) => c.className.includes('python')) || codes[0];
    const ts = codes.find((c) => c.className.includes('ts')) || codes[0];
    return {
        borderTop: bs.borderTopWidth,
        borderStyle: bs.borderTopStyle,
        borderRadius: bs.borderRadius,
        boxShadow: bs.boxShadow,
        codeCount: codes.length,
        pyTokens: py ? py.querySelectorAll('span[class^="hljs-"]').length : 0,
        pyKeywordColor: colorOf(py, '.hljs-keyword'),
        pyStringColor: colorOf(py, '.hljs-string'),
        pyCommentColor: colorOf(py, '.hljs-comment'),
        tsTypeColor: colorOf(ts, '.hljs-title, .hljs-type, .hljs-built_in'),
        // 源码完整性：文本内容必须与原始一致（高亮只加 span，不改字符）
        pyText: py ? py.textContent.replace(/\s+/g, ' ').trim() : '',
        mermaidHighlighted: codes.some((c) => c.className.includes('mermaid') && c.querySelector('.hljs-keyword, .hljs-string'))
    };
});

check('内容外框无 1px 描边', r.borderStyle === 'none' || parseFloat(r.borderTop) === 0,
    `border=${r.borderTop} ${r.borderStyle}`);
check('圆角与阴影仍保留（纸面观感）', parseFloat(r.borderRadius) > 0, `radius=${r.borderRadius}`);

check('python 代码块已高亮', r.pyTokens > 0, `token 数=${r.pyTokens}`);
check('关键字有颜色', !!r.pyKeywordColor && r.pyKeywordColor !== 'rgb(0, 0, 0)', `keyword=${r.pyKeywordColor}`);
check('字符串有颜色', !!r.pyStringColor, `string=${r.pyStringColor}`);
check('注释有颜色', !!r.pyCommentColor, `comment=${r.pyCommentColor}`);
check('类型/标题有颜色', !!r.tsTypeColor, `ts=${r.tsTypeColor}`);
check('高亮不破坏源码内容', r.pyText.includes('def greet(name: str) -> str:') && r.pyText.includes('hello {name}'),
    r.pyText.slice(0, 60));
check('mermaid 块未被着色（留给图表渲染）', !r.mermaidHighlighted);

console.log(failed ? `\nFAILED (${failed} failure(s))` : '\nALL PASSED (0 failure(s))');
await browser.close();
server.close();
process.exit(failed ? 1 : 0);
// 缩放回归验证：适应是否真的生效 + 缩放后左侧目录跳转是否落到正确位置。
//
// 直接驱动 office-editor（DOCX），断言：
// 1. 打开后「适应」自动算出的倍率不是 1（窄页面应放大铺满，而不是锁死在 100%）
// 2. 点「适应宽度」后倍率与 auto 标记正确
// 3. 非 100% 缩放下点击左侧目录项，目标标题实际落在内容区顶部附近（换算正确）
// 4. 缩放变化后侧栏高亮会跟着更新
//
// 用法：node scripts/verify-zoom.mjs
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
function check(name, ok, detail = '') {
    console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  ' + detail : ''}`);
    if (!ok) failed++;
}

function docxBytes() {
    const blocks = [];
    for (let i = 1; i <= 12; i++) {
        blocks.push({ type: 'heading', level: 2, text: `第 ${i} 章` });
        for (let j = 0; j < 6; j++) {
            blocks.push({ type: 'paragraph', text: `第 ${i} 章正文段落 ${j + 1}。` .repeat(14) });
        }
    }
    return jdoc.jsonToDocx({
        section: { width: 595.28, height: 841.89, marginTop: 72, marginBottom: 72, marginLeft: 72, marginRight: 72 },
        blocks
    });
}

const server = http.createServer((req, res) => {
    const url = req.url.split('?')[0];
    if (url === '/' || url === '/index.html') {
        res.writeHead(200, { 'Content-Type': 'text/html' });
        res.end(`<!doctype html><html><head><meta charset="utf-8"></head><body><div id="app"></div>
<script>window.acquireVsCodeApi=function(){return {postMessage:function(m){try{(window.__POSTS__=window.__POSTS__||[]).push(m);}catch(e){}}};};</script>
<script>window.__ERRS__=[];(function(){const o=console.error.bind(console);console.error=function(...a){try{window.__ERRS__.push(a.map(x=>x&&x.stack?x.stack:String(x)).join(' '));}catch(e){}o(...a);};})();</script>
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

const browser = await chromium.launch({ args: ['--no-sandbox'] });
const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
await new Promise((r) => server.listen(0, r));
await page.goto(`http://127.0.0.1:${server.address().port}/`, { waitUntil: 'load' });

const b64 = Buffer.from(await docxBytes()).toString('base64');
await page.evaluate(
    (bytes) => window.postMessage({ type: 'init', format: 'docx', mode: 'view', bytes, title: 'zoom.docx', dirty: false, locale: 'zh-cn' }, '*'),
    b64
);
await page.waitForSelector('.oe-page', { timeout: 15000 });
await page.waitForTimeout(800);

console.log('\n== ZOOM ==');

// 1. 打开后自动适应：A4(≈794px) 在 1440 宽视口里应放大到 >1，而不是锁死 100%
const autoZoom = await page.evaluate(() => {
    const label = [...document.querySelectorAll('.jc-zoom .chip')].find((b) => /%$/.test(b.textContent));
    const fit = [...document.querySelectorAll('.jc-zoom .chip')].find((b) => /适应/.test(b.textContent));
    return { label: label?.textContent, fitOn: fit?.classList.contains('on') };
});
const autoPct = Number(String(autoZoom.label).replace('%', ''));
check('打开后自动适应宽度（窄页面放大而非锁死 100%）', autoZoom.fitOn && autoPct > 100, `zoom=${autoZoom.label} fitOn=${autoZoom.fitOn}`);

// 2. 缩放真的作用到了 DOM 上（target 的 zoom 属性非空）
const applied = await page.evaluate(() => {
    const root = document.querySelector('.oe-content > *');
    return { zoom: root?.style.zoom || '', w: Math.round(root?.getBoundingClientRect().width || 0) };
});
check('缩放已应用到内容根', applied.zoom !== '' && applied.zoom !== '1', `style.zoom=${applied.zoom} visualWidth=${applied.w}`);

// 3. 点「适应宽度」后仍是 auto 且倍率稳定
await page.click('.jc-zoom .chip[title="适应宽度"]');
await page.waitForTimeout(300);
const afterFit = await page.evaluate(() => {
    const label = [...document.querySelectorAll('.jc-zoom .chip')].find((b) => /%$/.test(b.textContent));
    return label?.textContent;
});
check('点「适应」回到自动宽度', Number(String(afterFit).replace('%', '')) > 100, `zoom=${afterFit}`);

// 4. 非 100% 缩放下点击左侧目录项，目标标题应落在内容区顶部附近
await page.click('.jc-zoom .chip[title="实际大小（100%）"]'); // 先设成手动 100%
await page.waitForTimeout(200);
const jump = await page.evaluate(async (idx) => {
    const settle = async () => {
        // smooth 滚动时长随距离变化，轮询到 scrollTop 连续两次不变才算停稳
        const content = document.querySelector('.oe-content');
        let prev = -1;
        for (let i = 0; i < 40; i++) {
            await new Promise((r) => setTimeout(r, 100));
            if (content.scrollTop === prev) break;
            prev = content.scrollTop;
        }
    };
    const items = [...document.querySelectorAll('.oe-outline-item')];
    const target = items[idx];
    const id = target.dataset.id;
    target.click();
    await settle();
    const heading = document.querySelector(`[data-oe-anchor="${id}"]`);
    const content = document.querySelector('.oe-content');
    if (!heading || !content) return { ok: false, reason: 'missing' };
    const delta = heading.getBoundingClientRect().top - content.getBoundingClientRect().top;
    const zoom = Number(getComputedStyle(document.querySelector('.oe-content > *')).zoom || 1);
    return { ok: true, delta: Math.round(delta), zoom, scrollTop: Math.round(content.scrollTop) };
}, 3);
check('缩放后目录跳转定位正确（目标标题落在内容区顶部 ±30px）', jump.ok && Math.abs(jump.delta) <= 30,
    `delta=${jump.delta}px zoom=${jump.zoom} scrollTop=${jump.scrollTop}`);

// 5. 换到缩小的倍率再跳一次，确认换算不是只在放大时碰巧成立
await page.click('.jc-zoom .chip[title="缩小"]');
await page.click('.jc-zoom .chip[title="缩小"]');
await page.waitForTimeout(300);
const jump2 = await page.evaluate(async (idx) => {
    const settle = async () => {
        const content = document.querySelector('.oe-content');
        let prev = -1;
        for (let i = 0; i < 40; i++) {
            await new Promise((r) => setTimeout(r, 100));
            if (content.scrollTop === prev) break;
            prev = content.scrollTop;
        }
    };
    const items = [...document.querySelectorAll('.oe-outline-item')];
    const target = items[idx];
    const id = target.dataset.id;
    target.click();
    await settle();
    const heading = document.querySelector(`[data-oe-anchor="${id}"]`);
    const content = document.querySelector('.oe-content');
    if (!heading || !content) return { ok: false };
    const delta = heading.getBoundingClientRect().top - content.getBoundingClientRect().top;
    const zoom = Number(getComputedStyle(document.querySelector('.oe-content > *')).zoom || 1);
    return { ok: true, delta: Math.round(delta), zoom };
}, 5);
check('缩小倍率下目录跳转同样定位正确', jump2.ok && Math.abs(jump2.delta) <= 30, `delta=${jump2.delta}px zoom=${jump2.zoom}`);

// 6. 放大到 >100% 再跳一次：视觉坐标放大后最容易暴露「视觉坐标与布局坐标混用」
await page.click('.jc-zoom .chip[title="适应宽度"]');
await page.waitForTimeout(300);
const jump3 = await page.evaluate(async (idx) => {
    const settle = async () => {
        const content = document.querySelector('.oe-content');
        let prev = -1;
        for (let i = 0; i < 40; i++) {
            await new Promise((r) => setTimeout(r, 100));
            if (content.scrollTop === prev) break;
            prev = content.scrollTop;
        }
    };
    const items = [...document.querySelectorAll('.oe-outline-item')];
    const target = items[idx];
    const id = target.dataset.id;
    target.click();
    await settle();
    const heading = document.querySelector(`[data-oe-anchor="${id}"]`);
    const content = document.querySelector('.oe-content');
    if (!heading || !content) return { ok: false };
    const delta = heading.getBoundingClientRect().top - content.getBoundingClientRect().top;
    const zoom = Number(getComputedStyle(document.querySelector('.oe-content > *')).zoom || 1);
    return { ok: true, delta: Math.round(delta), zoom };
}, 6);
check('放大（适应宽度）后目录跳转定位正确', jump3.ok && Math.abs(jump3.delta) <= 30, `delta=${jump3.delta}px zoom=${jump3.zoom}`);

// 7. 缩放会重排内容但不一定触发 scroll：侧栏高亮必须跟着重算，否则会停在旧标题上
const sync = await page.evaluate(async () => {
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    const content = document.querySelector('.oe-content');
    // 与 trackActiveHeading 同样的 24px 阈值，独立算一遍「视口顶部标题」
    const pickTop = () => {
        const limit = content.getBoundingClientRect().top + 24;
        let id = null;
        for (const el of content.querySelectorAll('[data-oe-anchor]')) {
            if (el.getBoundingClientRect().top <= limit) id = el.dataset.oeAnchor;
            else break;
        }
        return id;
    };
    content.scrollTop = Math.round(content.scrollHeight * 0.45);
    await sleep(400);
    const zoomChip = [...document.querySelectorAll('.jc-zoom .chip')].find((b) => b.title === '适应宽度');
    zoomChip.click();
    await sleep(700);
    return { expected: pickTop(), active: document.querySelector('.oe-outline-item.active')?.dataset.id };
});
check('缩放后侧栏高亮与视口顶部标题一致', sync.expected !== null && sync.expected === sync.active,
    `expected=${sync.expected} active=${sync.active}`);

// 8. XLSX 编辑模式：.oe-grid 每次重绘都被重建，缩放必须重新绑到新网格上，否则会失效
const xb64 = Buffer.from(
    await jdoc.jsonToXlsx({
        sheets: [
            {
                name: 'S1',
                rows: Array.from({ length: 40 }, (_, r) => Array.from({ length: 8 }, (_, c) => `r${r + 1}c${c + 1}`))
            }
        ]
    })
).toString('base64');
await page.evaluate(
    (b) => window.postMessage({ type: 'init', format: 'xlsx', mode: 'edit', bytes: b, title: 'z.xlsx', dirty: false, locale: 'zh-cn' }, '*'),
    xb64
);
await page.waitForSelector('.oe-grid', { timeout: 15000 });
await page.waitForTimeout(600);
const gridBefore = await page.evaluate(() => document.querySelector('.oe-grid')?.style.zoom || '');
await page.click('.jc-zoom .chip[title="放大"]');
await page.waitForTimeout(400);
const gridAfter = await page.evaluate(() => document.querySelector('.oe-grid')?.style.zoom || '');
check('XLSX 缩放作用在网格上（重绘后不丢）', gridBefore !== gridAfter && gridAfter !== '',
    `before=${gridBefore || '(none)'} after=${gridAfter || '(none)'}`);

// 9. 幂等性：反复触发自动重算（折叠/展开大纲）不得让倍率一路漂移。
//    旧实现用 rect/zoom 自引用，paper 退化成撑满容器时每次乘一个 <1 的系数，会收敛到 MIN_ZOOM。
const drift = await page.evaluate(async () => {
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    const readZoom = () => {
        const label = [...document.querySelectorAll('.jc-zoom .chip')].find((b) => /%$/.test(b.textContent));
        return Number(String(label?.textContent).replace('%', ''));
    };
    document.querySelector('.jc-zoom .chip[title="适应宽度"]').click();
    await sleep(250);
    const toggle = document.querySelector('.oe-outline-toggle');
    const samples = [readZoom()];
    for (let i = 0; i < 6; i++) {
        toggle.click(); // 折叠/展开来回切，每次都会触发 bindZoom 重算
        await sleep(120);
        toggle.click();
        await sleep(120);
        samples.push(readZoom());
    }
    return samples;
});
check('反复重算不漂移（幂等，不滑向 MIN_ZOOM）', Math.min(...drift) > 60 && Math.max(...drift) - Math.min(...drift) <= 2,
    `samples=${drift.join(',')}`);

// 10. 真实 docx 的末尾条目：标题推不到内容区顶部时（后面内容不足一屏），
//     必须退化为「滚到所在页的页顶」，而不是卡在文档底部让目标偏出预期位置。
//     该缺陷只在缩放后的真实文档上出现：适应宽度下最后一条曾偏 168px 且页顶不可见。
const EXAMPLES = path.resolve(__dirname, '..', '..', 'examples');
const realDocx = fs
    .readdirSync(EXAMPLES)
    .filter((f) => f.endsWith('.docx'))
    .map((f) => ({ name: f, b64: fs.readFileSync(path.join(EXAMPLES, f)).toString('base64') }));

for (const doc of realDocx) {
    await page.evaluate(
        (bytes) =>
            window.postMessage(
                { type: 'init', format: 'docx', mode: 'view', bytes, title: 'real.docx', dirty: false, locale: 'zh-cn' },
                '*'
            ),
        doc.b64
    );
    await page.waitForSelector('.oe-page', { timeout: 20000 });
    await page.waitForTimeout(1200);

    const last = await page.evaluate(async () => {
        const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
        const c = document.querySelector('.oe-content');
        const items = [...document.querySelectorAll('.oe-outline-item')];
        if (!items.length) return { err: 'no outline item' };
        const t = items[items.length - 1];
        const id = t.dataset.id;
        const zoom = Number(getComputedStyle(document.querySelector('.oe-content > *')).zoom || 1);
        t.click();
        let prev = -1;
        for (let k = 0; k < 50; k++) {
            await sleep(100);
            if (c.scrollTop === prev) break;
            prev = c.scrollTop;
        }
        const h = document.querySelector(`[data-oe-anchor="${id}"]`);
        if (!h) return { err: 'no anchor' };
        const hr = h.getBoundingClientRect();
        const cr = c.getBoundingClientRect();
        const pg = h.closest('.oe-page');
        return {
            id,
            zoom,
            label: t.textContent.slice(0, 18),
            delta: Math.round(hr.top - cr.top),
            visible: hr.top < cr.bottom && hr.bottom > cr.top,
            pageTopAligned: pg
                ? Math.round(pg.getBoundingClientRect().top - cr.top + (parseFloat(getComputedStyle(pg).paddingTop) || 0) - 12)
                : null,
            atBottom: Math.round(c.scrollTop) >= Math.round(c.scrollHeight - c.clientHeight) - 1
        };
    });

    // 判据：标题顶部对齐，或已退化为页顶对齐（含页面 padding 补偿）；且目标必须可见
    const aligned = Math.abs(last.delta ?? 999) <= 30 || Math.abs(last.pageTopAligned ?? 999) <= 30;
    check(
        `真实文档末尾条目可定位（${doc.name}）`,
        !last.err && last.visible && aligned,
        `label=${last.label} zoom=${last.zoom} delta=${last.delta} pageTopAligned=${last.pageTopAligned} atBottom=${last.atBottom}`
    );
}

const errs = await page.evaluate(() => window.__ERRS__ || []);
check('无运行时错误', errs.length === 0, errs.slice(0, 2).join(' | '));

console.log(failed ? `\nFAILED (${failed} failure(s))` : '\nALL PASSED (0 failure(s))');
await browser.close();
server.close();
process.exit(failed ? 1 : 0);
// 缩放回归验证：适应是否真的生效 + 缩放后左侧目录跳转是否落到正确位置。
//
// 直接驱动 office-editor（DOCX），断言：
// 1. 打开后「适应」符合「只缩不放」语义：宽屏保持 100% 居中，窄屏才按屏宽缩小
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

// 1. 打开后自动适应：A4(≈794px) 在 1440 宽视口里并不比可视区宽，
//    应保持 100% 实际大小（不再放大铺满），且仍处于自动模式。
const autoZoom = await page.evaluate(() => {
    const label = [...document.querySelectorAll('.jc-zoom .chip')].find((b) => /%$/.test(b.textContent));
    const fit = [...document.querySelectorAll('.jc-zoom .chip')].find((b) => /适应/.test(b.textContent));
    return { label: label?.textContent, fitOn: fit?.classList.contains('on') };
});
const autoPct = Number(String(autoZoom.label).replace('%', ''));
check('打开后自动适应：宽屏保持 100%（不放大铺满）', autoZoom.fitOn && autoPct === 100,
    `zoom=${autoZoom.label} fitOn=${autoZoom.fitOn}`);

// 2. 只有内容比可视区宽时才按屏宽缩小，且缩放确实落到 DOM 上（zoom=1 时 style.zoom 会被清空，
//    所以必须在窄视口下验证才有非空 style.zoom）
await page.setViewportSize({ width: 700, height: 900 });
await page.waitForTimeout(600);
const narrow = await page.evaluate(() => {
    const label = [...document.querySelectorAll('.jc-zoom .chip')].find((b) => /%$/.test(b.textContent));
    const root = document.querySelector('.oe-content > *');
    return { label: label?.textContent, zoom: root?.style.zoom || '' };
});
const narrowPct = Number(String(narrow.label).replace('%', ''));
check('窄视口按屏宽缩小（<100% 且 zoom 已应用到内容根）',
    narrowPct < 100 && narrowPct >= 20 && narrow.zoom !== '' && narrow.zoom !== '1',
    `zoom=${narrow.label} style.zoom=${narrow.zoom || '(none)'}`);

// 3. 视口放宽后自动恢复；点「适应宽度」同样稳定在 100%
await page.setViewportSize({ width: 1440, height: 900 });
await page.waitForTimeout(600);
await page.click('.jc-zoom .chip[title="适应宽度"]');
await page.waitForTimeout(300);
const afterFit = await page.evaluate(() => {
    const label = [...document.querySelectorAll('.jc-zoom .chip')].find((b) => /%$/.test(b.textContent));
    return label?.textContent;
});
check('点「适应」回到自动宽度（宽屏 100%）', Number(String(afterFit).replace('%', '')) === 100, `zoom=${afterFit}`);

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
//    （「适应宽度」已改为只缩不放、不再放大，故这里手动置 100% 再点「＋」制造放大态）
await page.click('.jc-zoom .chip[title="实际大小（100%）"]');
await page.click('.jc-zoom .chip[title="放大"]');
await page.click('.jc-zoom .chip[title="放大"]');
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
check('放大（>100%）后目录跳转定位正确', jump3.ok && Math.abs(jump3.delta) <= 30, `delta=${jump3.delta}px zoom=${jump3.zoom}`);

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
const gridBefore = await page.evaluate(() => document.querySelector('.oe-grid-canvas')?.style.zoom || '');
await page.click('.jc-zoom .chip[title="放大"]');
await page.waitForTimeout(400);
const gridAfter = await page.evaluate(() => document.querySelector('.oe-grid-canvas')?.style.zoom || '');
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

// 10. 真实 docx × 多种倍率 × 全部条目：目录跳转必须准确。
//     这组组合此前从未被覆盖——合成文档只测了 3 个条目、真实文档只测了最后一个，
//     「真实文档 + 中间条目 + 缩放」正好是线上失准的场景。
//
//     判据分两种：
//     - 能贴顶时（绝大多数）：标题落在内容区顶部 ±30px 内；
//     - 物理上贴不了顶时（文档末尾，标题后面不足一屏）：退化为滚到底/页顶，
//       此时 delta 可以大于 30，但标题必须**完整可见**。
//     旧口径用 getComputedStyle(page).paddingTop 判页顶对齐，那是未缩放的布局值
//     （实测恒为 96px，与 zoom 无关），与视觉坐标系混用后无法作为判据。
const EXAMPLES = path.resolve(__dirname, '..', '..', 'examples');
const realDocx = fs
    .readdirSync(EXAMPLES)
    .filter((f) => f.endsWith('.docx'))
    .map((f) => ({ name: f, b64: fs.readFileSync(path.join(EXAMPLES, f)).toString('base64') }));

const ZOOM_MODES = [
    ['适应', async () => clickChip(page, '适应宽度')],
    ['100%', async () => clickChip(page, '实际大小（100%）')],
    ['放大', async () => {
        await clickChip(page, '实际大小（100%）');
        await clickChip(page, '放大');
    }],
    ['缩小', async () => {
        await clickChip(page, '实际大小（100%）');
        await clickChip(page, '缩小');
        await clickChip(page, '缩小');
        await clickChip(page, '缩小');
    }]
];

async function clickChip(pg, title) {
    await pg.click(`.jc-zoom .chip[title="${title}"]`);
    await pg.waitForTimeout(450);
}

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

    for (const [modeName, setup] of ZOOM_MODES) {
        await setup();
        const res = await page.evaluate(async () => {
            const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
            const c = document.querySelector('.oe-content');
            const items = [...document.querySelectorAll('.oe-outline-item')];
            const zoom = Number(getComputedStyle(document.querySelector('.oe-content > *')).zoom || 1);
            const bad = [];
            let checked = 0;
            for (const t of items) {
                const id = t.dataset.id;
                c.scrollTop = 0;
                await sleep(120);
                t.click();
                let prev = -1;
                for (let k = 0; k < 40; k++) {
                    await sleep(80);
                    if (c.scrollTop === prev) break;
                    prev = c.scrollTop;
                }
                await sleep(200);
                const h = document.querySelector(`[data-oe-anchor="${id}"]`);
                if (!h) {
                    bad.push({ label: t.textContent.slice(0, 12), why: 'no-anchor' });
                    continue;
                }
                checked++;
                const cr = c.getBoundingClientRect();
                const hr = h.getBoundingClientRect();
                const delta = Math.round(hr.top - cr.top);
                // 能贴顶 -> 必须贴顶；贴不了顶 -> 必须完整可见
                const ok = Math.abs(delta) <= 30 || (hr.top >= cr.top - 1 && hr.bottom <= cr.bottom + 1);
                if (!ok) {
                    bad.push({
                        label: t.textContent.slice(0, 12),
                        delta,
                        viewH: Math.round(cr.height),
                        headTop: Math.round(hr.top - cr.top),
                        headBottom: Math.round(hr.bottom - cr.top)
                    });
                }
            }
            return { zoom, checked, bad };
        });
        check(
            `真实文档全条目跳转准确（${doc.name} @ ${modeName}）`,
            res.bad.length === 0 && res.checked > 0,
            `zoom=${res.zoom} 检查${res.checked}项` + (res.bad.length ? ` 失准: ${JSON.stringify(res.bad.slice(0, 3))}` : '')
        );
    }
}

// 11. 极小倍率（32% / 20%）下全条目跳转。线上事故场景：文档缩到 32% 后点「4.2
//     准入级别计算」，视图却滚到了「7.总结」。上面的「缩小」只到 75%，覆盖不到。
async function setZoomPct(pct) {
    await page.evaluate(async (target) => {
        const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
        const read = () => {
            const b = [...document.querySelectorAll('.jc-zoom .chip')].find((x) => /%$/.test(x.textContent));
            return Number(String(b?.textContent).replace('%', ''));
        };
        const out = document.querySelector('.jc-zoom .chip[title="缩小"]');
        const inn = document.querySelector('.jc-zoom .chip[title="放大"]');
        document.querySelector('.jc-zoom .chip[title="实际大小（100%）"]').click();
        await sleep(150);
        let guard = 0;
        while (read() > target + 1 && guard++ < 60) {
            out.click();
            await sleep(90);
        }
        guard = 0;
        while (read() < target - 1 && guard++ < 60) {
            inn.click();
            await sleep(90);
        }
    }, pct);
    await page.waitForTimeout(450);
}

/** 在浏览器里遍历全部目录条目并判定位是否准确 */
async function sweepOutline() {
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    const c = document.querySelector('.oe-content');
    const items = [...document.querySelectorAll('.oe-outline-item')];
    const bad = [];
    let checked = 0;
    for (const t of items) {
        const id = t.dataset.id;
        t.click();
        let prev = -1;
        for (let k = 0; k < 40; k++) {
            await sleep(80);
            if (c.scrollTop === prev) break;
            prev = c.scrollTop;
        }
        await sleep(200);
        const h = document.querySelector(`[data-oe-anchor="${id}"]`);
        if (!h) {
            bad.push({ label: t.textContent.slice(0, 12), why: 'no-anchor' });
            continue;
        }
        checked++;
        const cr = c.getBoundingClientRect();
        const hr = h.getBoundingClientRect();
        const delta = Math.round(hr.top - cr.top);
        // 能贴顶 -> 必须贴顶；贴不了顶（文档末尾） -> 必须完整可见
        const ok = Math.abs(delta) <= 30 || (hr.top >= cr.top - 1 && hr.bottom <= cr.bottom + 1);
        if (!ok) {
            bad.push({
                label: t.textContent.slice(0, 12),
                delta,
                scrollTop: Math.round(c.scrollTop),
                max: c.scrollHeight - c.clientHeight
            });
        }
    }
    return { checked, bad, zoom: document.querySelector('.oe-content > *')?.style.zoom };
}

for (const pct of [32, 20]) {
    await setZoomPct(pct);
    const res = await page.evaluate(sweepOutline);
    check(
        `极小倍率全条目跳转准确（目标 ${pct}%，实测 ${res.zoom}）`,
        res.bad.length === 0 && res.checked > 0,
        `检查${res.checked}项` + (res.bad.length ? ` 失准: ${JSON.stringify(res.bad.slice(0, 3))}` : '')
    );
}

// 12. 祖先层还有一层 CSS zoom 时的跳转。VS Code 的窗口缩放以 CSS zoom 落在 webview
//     的 body/html 上，此时 getBoundingClientRect 是「视觉坐标」、scrollTop 是「内容
//     坐标」，两者比值 = 那一层倍率。旧实现直接把 rect 差加 scrollTop，落点被整体
//     放大，小倍率下更早撞上可滚动上限而被 clamp 到文末（正是上面第 11 项的事故现象）。
for (const bodyZoom of ['1.4', '0.8']) {
    await setZoomPct(32);
    await page.evaluate((z) => {
        document.body.style.zoom = z;
    }, bodyZoom);
    await page.waitForTimeout(350);
    const res = await page.evaluate(sweepOutline);
    check(
        `祖先 CSS zoom=${bodyZoom} 时全条目跳转仍准确`,
        res.bad.length === 0 && res.checked > 0,
        `zoom=${res.zoom} 检查${res.checked}项` + (res.bad.length ? ` 失准: ${JSON.stringify(res.bad.slice(0, 3))}` : '')
    );
}
await page.evaluate(() => {
    document.body.style.zoom = '';
});
await setZoomPct(100);
await page.waitForTimeout(300);

// 13. 跳转期间侧栏高亮必须锁定在用户点的那条：smooth 滚动一路触发滚动联动，
//     它按「视口顶部实际是哪个标题」改写高亮，落稳那一刻往往不是刚点的那条。
const lockCase = await page.evaluate(async () => {
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    const items = [...document.querySelectorAll('.oe-outline-item')];
    // 挑一条不在视口顶部的（最容易出现高亮错位）
    const t = items[Math.min(6, items.length - 1)];
    const id = t.dataset.id;
    t.click();
    await sleep(120);
    const midLock = document.querySelector('.oe-outline-item.active')?.dataset.id;
    let prev = -1;
    const c = document.querySelector('.oe-content');
    for (let k = 0; k < 40; k++) {
        await sleep(100);
        if (c.scrollTop === prev) break;
        prev = c.scrollTop;
    }
    await sleep(400);
    return { id, midLock, settled: document.querySelector('.oe-outline-item.active')?.dataset.id };
});
check(
    '跳转后侧栏高亮停在用户点的条目',
    lockCase.id === lockCase.settled && lockCase.id === lockCase.midLock,
    `target=${lockCase.id} duringScroll=${lockCase.midLock} settled=${lockCase.settled}`
);

// 14. 跳转锁必须释放：落稳后手动滚动，高亮要重新跟随
const unlockCase = await page.evaluate(async () => {
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    const c = document.querySelector('.oe-content');
    c.scrollTop = Math.round(c.scrollHeight * 0.6);
    await sleep(600);
    return { active: document.querySelector('.oe-outline-item.active')?.dataset.id || null };
});
check('跳转锁会释放（落稳后手动滚动仍联动高亮）', unlockCase.active !== null, `active=${unlockCase.active}`);

// 15. 动画让位：跳转动画进行中用户滚动（wheel），动画必须立刻停、不得把视图拽回目标。
//     旧实现（原生 smooth + 落稳补跳）做不到这点；也让「滚动锚定改写 scrollTop」
//     这类非用户输入不再误伤动画。
const yieldCase = await page.evaluate(async () => {
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    const c = document.querySelector('.oe-content');
    c.scrollTop = 0;
    await sleep(200);
    const items = [...document.querySelectorAll('.oe-outline-item')];
    const t = items[items.length - 1]; // 最远条目，动画路径最长
    const id = t.dataset.id;
    t.click();
    await sleep(100); // 动画进行中
    c.dispatchEvent(new WheelEvent('wheel', { deltaY: 240, bubbles: true }));
    await sleep(700);
    const h = document.querySelector(`[data-oe-anchor="${id}"]`);
    const delta = Math.round(h.getBoundingClientRect().top - c.getBoundingClientRect().top);
    return { id, delta };
});
check(
    '动画中用户滚动会立刻让位（不拽回目标）',
    Math.abs(yieldCase.delta - 12) > 40, // 若没让位，动画仍会把标题拽到 +12px 贴顶
    `delta=${yieldCase.delta}px`
);

const errs = await page.evaluate(() => window.__ERRS__ || []);
check('无运行时错误', errs.length === 0, errs.slice(0, 2).join(' | '));

console.log(failed ? `\nFAILED (${failed} failure(s))` : '\nALL PASSED (0 failure(s))');
await browser.close();
server.close();
process.exit(failed ? 1 : 0);
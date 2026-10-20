// office-editor 入口：DOCX / XLSX 的查看（view）与编辑（edit）。
// 独立自包含入口，避免与查看版 webview.js 的双 message 监听 / 双 acquireVsCodeApi 冲突。
//
// 数据流：host 发 init(bytes) → view 用 docxToHtml/xlsxToHtml 只读渲染；edit 用块式/网格
// 编辑器载入 blocks/sheets 模型 → 编辑 → jsonToDocx/jsonToXlsx 生成字节 → sync 回 host 保存。
//
// 外壳：左侧目录/编号索引栏（大纲）+ 右侧内容区。DOCX 两个模式都按纸张高度分页，
// 每页独立成盒并带页码，观感对齐 Word/WPS；XLSX 是网格工作表，不分页、无大纲。
import { docxToHtml, type DocxSection } from '@fefeding/jdoc';
import { officeCss } from './style';
import { CHROME_CSS } from '../chrome';
import { convertMenuEntry, copySelection, hasTextSelection, installContextMenu, notifyConvertAck, postConvert, setConvertStatusSink, setCurrentDocUri, setHostPostMessage } from '../chrome';
import { createZoom, createZoomCluster } from '../zoom';
import { setLocale, t } from '../i18n';
import { toast } from './framework';
import { createDocxEditor, type DocxEditor } from './docx';
import { createXlsxEditor, type XlsxEditor } from './xlsx';
import { b64ToBytes, bytesToB64, h, escapeHtml } from './util';
import { createOutline, buildOutline, trackActiveHeading, type OutlineView, type ActiveHeadingTracker } from './outline';
import { paginate, waitImages, type PageMetrics } from './paging';
import type { DocFormat, EditorMode, HostToWebview, WebviewToHost } from '../../protocol';

declare function acquireVsCodeApi(): { postMessage(msg: unknown): void };
const vscode = acquireVsCodeApi();

let format: DocFormat = 'pptx';
let mode: EditorMode = 'view';
let latest: Uint8Array = new Uint8Array(0);
let original: Uint8Array = new Uint8Array(0);
let gen = 0; // init 代数：丢弃过期的异步 applyMode 结果，防快速切换竞态

let editor: DocxEditor | XlsxEditor | null = null;
let root: HTMLElement;
let toolbarEl: HTMLElement;
let contentEl: HTMLElement;
let statusEl: HTMLElement;
let statusBarEl: HTMLElement;
let convertBtnEl: HTMLElement;

const zoom = createZoom();
/** 当前已绑定缩放的内容根，用于跳过重复绑定 */
let zoomBoundTo: HTMLElement | null = null;

/** 左侧目录栏（XLSX 无大纲，整条隐藏） */
let outline: OutlineView | null = null;
/** 滚动联动句柄（refresh 用于缩放后重算高亮） */
let untrack: ActiveHeadingTracker | null = null;
/** 目录跳转进行中：期间忽略滚动联动对侧栏高亮的改写，落稳后自动释放 */
let jumpLock = false;
/** 查看模式渲染代数：丢弃过期的异步分页结果 */
let viewGen = 0;

function post(msg: WebviewToHost): void {
    vscode.postMessage(msg);
}

// chrome.ts 的转换入口需要向宿主发消息，但它不能自己调 acquireVsCodeApi()
// （每个 webview 只能调一次，第二次会抛错），由这里注入本入口的 post 函数
setHostPostMessage((m) => post(m as WebviewToHost));
// 转换过程的状态提示接到状态栏：点击后立刻有反馈，失败时不再是一片死寂
setConvertStatusSink((text) => {
    if (statusEl) statusEl.textContent = text;
});

const hooks = {
    onSync: (bytes: Uint8Array | null): void => {
        if (bytes === null) {
            // 放弃编辑回到导入态
            latest = original;
            post({ type: 'sync', bytes: null, dirty: false });
        } else {
            latest = bytes;
            post({ type: 'sync', bytes: bytesToB64(bytes), dirty: true });
        }
    },
    onToast: (message: string, kind?: 'info' | 'warn' | 'error'): void => toast(message, kind),
    /** 编辑器重建内容后请外壳重建大纲 / 重新分页 */
    onContentChanged: () => refreshOutline()
};

// ---------- 外壳 ----------
function buildShell(): void {
    root = document.getElementById('app')!;
    root.innerHTML = '';
    const app = h('div', { class: 'oe' });

    // 不画顶栏：文档标题 / 未保存点由 VS Code 编辑器标题栏呈现；查看/编辑切换放在
    // VS Code 标题栏（editor/title 菜单），由 host 通过 setMode 消息驱动。
    toolbarEl = h('div', { class: 'oe-toolbar' });
    contentEl = h('div', { class: 'oe-content' });
    statusEl = h('span', { class: 'grow' });

    outline = createOutline({
        onJump: (_item, el) => scrollToAnchor(el),
        onToggle: () => bindZoom()
    });

    statusBarEl = h('div', { class: 'oe-statusbar' });
    convertBtnEl = convertBtn();
    statusBarEl.append(statusEl, convertBtnEl, createZoomCluster(zoom));
    app.append(toolbarEl, h('div', { class: 'oe-body' }, [outline.el, contentEl]), statusBarEl);
    root.append(app);

    // 缩放会改变内容高度与各标题的视口位置，但缩放本身不触发 scroll 事件，
    // 必须主动重算一次，否则侧栏高亮会停在上一个标题上。
    zoom.setOnChange(() => untrack?.refresh());

    // 编辑器在增删行列 / 切工作表 / undo 时会整体重建内容根，需重新绑定缩放目标。
    // 必须监听 subtree：XLSX 重建的是 .oe-xl 内部的 .oe-grid，contentEl 的直接子节点
    // 并没变，只看 childList 会漏掉，缩放就再也绑不回新网格上。
    new MutationObserver(() => bindZoom()).observe(contentEl, { childList: true, subtree: true });
}

/**
 * 状态栏「转换格式」按钮
 *
 * DOCX 打开后是 custom editor，正文在 webview 里：内容区的右键事件被webview 吞掉，
 * VS Code 的 editor/context 菜单**根本不会出现**，explorer/context 又只在资源管理器
 * 文件树里生效——只靠右键菜单，转换功能在编辑器里等于没有入口（这正是之前
 * 「移除 editor/title、全部收敛到右键」那次改动留下的死角）。这里在状态栏常驻一个
 * 按钮作为可靠入口，点击后交回宿主执行 jdoc.convert。
 */
function convertBtn(): HTMLElement {
    const btn = h('button', { class: 'chip', type: 'button' });
    btn.textContent = t('chrome.convert');
    btn.title = t('chrome.convertTitle');
    // 走 chrome.ts 的 postConvert：它会带上当前文档 uri，宿主才能精确定位目标文件。
    // 不能直接用 vscode.postMessage——那条路拿不到 uri。
    btn.addEventListener('click', () => postConvert());
    return btn;
}

/**
 * 解析真正能滚动的容器
 *
 * 正常就是 .oe-content。若它当前不可滚（外层布局塌缩、或缩放后内容比视口短），
 * 上溯找第一个真正能滚的祖先；都没有时返回 null —— 此时**必须保持当前视图**，
 * 任何 scrollTo 都会被浏览器归零，效果就是「点任何条目都被甩回顶部」。
 */
function resolveScroller(): HTMLElement | null {
    if (contentEl.scrollHeight > contentEl.clientHeight + 1) return contentEl;
    for (let el = contentEl.parentElement; el; el = el.parentElement) {
        const oy = getComputedStyle(el).overflowY;
        if ((oy === 'auto' || oy === 'scroll' || oy === 'overlay') && el.scrollHeight > el.clientHeight + 1) {
            return el;
        }
    }
    return null;
}

/** 把带缩放的文档滚动到视口内（大纲点击用） */
function scrollToAnchor(el: HTMLElement): void {
    el.classList.add('oe-flash');
    window.setTimeout(() => el.classList.remove('oe-flash'), 900);

    const scroller = resolveScroller();
    // 无处可滚：保持当前视图（否则 scrollTo 会被归零，反而把用户甩回顶部）
    if (!scroller) return;
    // 目标已不在文档里（重排后失效）：不猜位置，保持当前视图
    if (!el.isConnected || !scroller.contains(el)) return;

    /** 本次落点是否退化为「页顶对齐」；退化后 delta 天然偏大，不参与二次校正 */
    let degraded = false;
    // 跳转期间锁定侧栏高亮：smooth 滚动一路触发的 trackActiveHeading 会按「视口顶部
    // 实际是哪个标题」反复改写高亮，落稳那一刻算出的往往不是刚点的那条（末条退化为
    // 页顶对齐时尤其明显）。锁定到落稳后再交还给滚动联动。
    jumpLock = true;

    // rect 与 scrollTop 同属「视觉」坐标系：CSS zoom 生效后，滚动容器的 scrollTop
    // 也按缩放后的距离计，两者直接相加即为目标在滚动内容中的位置，不需要再除 zoom。
    const goto = (behavior: ScrollBehavior): void => {
        const sRect = scroller.getBoundingClientRect();
        const docTopOf = (node: HTMLElement): number =>
            node.getBoundingClientRect().top - sRect.top + scroller.scrollTop;
        const maxScroll = Math.max(0, scroller.scrollHeight - scroller.clientHeight);

        let desired = docTopOf(el) - 12;
        // 是否退化为「滚到所在页的页顶」。退化后标题本就推不到内容区顶部（delta 会
        // 明显大于阈值），二次校正不能再拿 delta 当判据，否则会白跳一次。
        degraded = false;
        if (desired > maxScroll) {
            // 末尾条目（末章、附录）后面不足一屏，标题物理上推不到内容区顶部。
            // 先按页顶对齐试算，全程只用 getBoundingClientRect 的**视觉坐标**：
            // 不能读 getComputedStyle(page).paddingTop——那是未缩放的布局值（实测恒为
            // 96px，与 zoom 无关），与视觉值相加就是坐标系混用，偏差 = padTop*(1-zoom)。
            const page = el.closest<HTMLElement>('.oe-page');
            if (page) {
                const pRect = page.getBoundingClientRect();
                // 页内首个正文块；没有就退回用标题自身（此时视觉 padding 恰为 0）
                const firstKid = page.querySelector<HTMLElement>('.docx-body') ?? page.firstElementChild;
                const visualPadTop = firstKid instanceof HTMLElement ? firstKid.getBoundingClientRect().top - pRect.top : 0;
                const pageAligned = Math.max(0, Math.min(docTopOf(page) + visualPadTop - 12, maxScroll));

                // 页顶对齐仍不足以把标题带进视口时（页面底部本身就顶到文档底部、
                // maxScroll 连页顶都到不了），退化为「滚到底」——此时标题必然可见，
                // 与 Word 导航窗格在文档末尾的落点一致。判断用「目标底边 vs 可视底边」
                // 这一实际可见性，而不是继续在坐标系上纠缠。
                const targetBottom = docTopOf(el) + el.getBoundingClientRect().height;
                const reachable = pageAligned + (el.getBoundingClientRect().top - (pRect.top + visualPadTop));
                if (targetBottom - maxScroll <= scroller.clientHeight && reachable <= maxScroll) {
                    degraded = true;
                    desired = pageAligned;
                } else {
                    degraded = true;
                    desired = maxScroll;
                }
            } else {
                degraded = true;
                desired = maxScroll;
            }
        }

        scroller.scrollTo({ top: Math.max(0, Math.min(desired, maxScroll)), behavior });
    };

    goto('smooth');

    // 落稳后二次校正：smooth 期间若发生重排（图片解码后重分页）或缩放变化，
    // 首次算出的位置会偏（表现为滚过头 / 没滚到位）。等滚动停止再量一次，
    // 偏差仍超过阈值就用 instant 拉回——只校正一次，避免来回抖动。
    // 退化（页顶对齐）时标题本就推不到顶部，delta 天然大于阈值，不在此列。
    let stableSince = 0;
    let lastTop = scroller.scrollTop;
    let done = false;
    const watch = (): void => {
        if (done) return;
        const top = scroller.scrollTop;
        if (top !== lastTop) {
            lastTop = top;
            stableSince = performance.now();
            window.setTimeout(watch, 100);
            return;
        }
        if (stableSince === 0) stableSince = performance.now();
        if (performance.now() - stableSince < 180) {
            window.setTimeout(watch, 100);
            return;
        }
        done = true;
        scroller.removeEventListener('scroll', watch);
        if (degraded || !el.isConnected) {
            jumpLock = false;
            return;
        }
        const delta = el.getBoundingClientRect().top - scroller.getBoundingClientRect().top;
        if (Math.abs(delta) > 24) {
            goto('instant');
            return; // goto 会重新上锁，等下一轮 watch 落稳再释放
        }
        jumpLock = false;
    };
    scroller.addEventListener('scroll', watch, { passive: true });
    window.setTimeout(watch, 100);
    // 兜底：watch 因容器被移除 / 事件不触发而没跑到时，必须释放跳转锁，
    // 否则侧栏高亮会永久停在点击项，后续手动滚动也不再联动。
    window.setTimeout(() => {
        jumpLock = false;
    }, 3000);
}

/** 依据当前内容重建左侧目录并接上滚动联动 */
function refreshOutline(): void {
    untrack?.dispose();
    untrack = null;
    if (!outline) return;
    if (format !== 'docx') {
        outline.update([]);
        return;
    }
    const view = contentEl.firstElementChild;
    if (!view) {
        outline.update([]);
        return;
    }
    const items = buildOutline(view as HTMLElement);
    outline.update(items);
    untrack = trackActiveHeading(
        contentEl,
        items.map((item) => item.id),
        (id) => outline?.setActive(id),
        () => jumpLock
    );
}

/** DOCX section -> 纸张设置；缺省 A4 纵向 1 英寸页边距 */
function metricsOf(section: DocxSection | undefined): PageMetrics {
    if (!section || !(section.width > 0) || !(section.height > 0)) {
        return { width: 595.28, height: 841.89, marginTop: 72, marginBottom: 72, marginLeft: 72, marginRight: 72 };
    }
    return {
        width: section.width,
        height: section.height,
        marginTop: section.marginTop,
        marginBottom: section.marginBottom,
        marginLeft: section.marginLeft,
        marginRight: section.marginRight
    };
}

/** 把缩放绑定到当前内容根（查看 / 编辑 / 内容重建后都要重新绑定） */
function bindZoom(): void {
    const el = contentEl.firstElementChild as HTMLElement | null;
    if (!el) {
        zoomBoundTo = null;
        zoom.bind({ target: null });
        return;
    }
    // XLSX 缩放作用在网格画布上：.oe-xl 是 height:100% 的 flex 容器，缩放它会把整个
    // 面板（含工作表标签与冻结区）按倍率压缩，和滚动区对不上。只缩画布本身，
    // 滚动容器 .oe-grid-wrap 保持满高，冻结行列仍按 sticky 正常吸附。
    let target: HTMLElement;
    let container: HTMLElement;
    let paper: HTMLElement;
    if (el.classList.contains('oe-xl')) {
        // 目标取 .oe-grid-canvas：表格与浮在上面的图表层都在它内部，一起缩放、
        // 不会错位；而 .oe-grid 每次重绘（增删行列 / 切工作表 / undo）都会被整体
        // 重建，canvas 本身不变，缩放倍率因此能在重绘后保留。
        const canvas = el.querySelector<HTMLElement>('.oe-grid-canvas');
        const grid = el.querySelector<HTMLElement>('.oe-grid');
        if (!canvas || !grid) return; // 网格还没建好，等下一次内容变化再绑
        target = canvas;
        container = el.querySelector<HTMLElement>('.oe-grid-wrap') ?? el;
        paper = grid;
    } else {
        // DOCX 必须以真实纸面为基准：若退化成 el（内容根），它的宽度就是容器宽，
        // fit 会算出「缩一点点」的无意义倍率。纸面未就绪时先不绑，等内容变化
        // 再来一次（观察器监听 subtree，纸面插入必然触发）。
        const page = el.querySelector<HTMLElement>('.oe-page');
        if (!page) return;
        target = el;
        container = contentEl;
        paper = page;
    }
    // 只有「同一个、且仍在文档里的」target 才算已绑定：旧 target 被替换后必须重绑
    if (zoomBoundTo === target && target.isConnected) return;
    zoomBoundTo = target;
    zoom.bind({ target, container, paper });
}

/**
 * 语言落地：状态栏里的文案（缩放簇、转换按钮）是创建时烘焙的，
 * 而外壳在 init 消息到达前就建好了——init 里的 locale 才是权威值，
 * 所以这里在拿到语言后重建一次带文案的控件。
 */
function applyLocale(): void {
    if (!statusBarEl) return;
    const oldCluster = statusBarEl.querySelector('.jc-zoom');
    if (oldCluster) statusBarEl.replaceChild(createZoomCluster(zoom), oldCluster);
    if (convertBtnEl) {
        convertBtnEl.textContent = t('chrome.convert');
        convertBtnEl.title = t('chrome.convertTitle');
    }
}

// ---------- 模式 ----------
async function applyMode(): Promise<void> {
    updateModeHint();
    if (mode === 'edit') await enterEdit();
    else await enterView();
}

async function enterView(): Promise<void> {
    // 离开编辑：先把编辑产出同步为字节，再忠实渲染
    if (editor) {
        await editor.syncNow();
        editor.destroy();
        editor = null;
    }
    const my = ++viewGen;
    toolbarEl.innerHTML = '';
    statusEl.textContent = t('oe.readOnly');
    contentEl.innerHTML = '';
    // 解析/分页期间先给出反馈：大文档（教材、报表）在 await 期间是纯白屏，
    // 用户会以为卡死。这里先挂一个提示，解析完成后再替换。
    contentEl.append(h('div', { class: 'oe-hint', text: t('oe.parsing', format.toUpperCase()) }));
    try {
        if (format === 'xlsx') {
            // XLSX 查看态复用编辑器的网格渲染（只读）：列宽行高、边框、图表、
            // 工作表标签都与 Excel 一致，不再退化成一张普通 HTML 表格
            const x = createXlsxEditor(hooks, { readonly: true });
            x.setStatusHandler((s) => (statusEl.textContent = s));
            editor = x;
            editor.mount(toolbarEl, contentEl);
            await editor.load(latest);
        } else if (format === 'docx') {
            const { html, section } = await docxToHtml(latest);
            const view = await renderDocxPages(html, section);
            if (my !== viewGen) return;
            contentEl.replaceChildren(view);
        } else {
            contentEl.append(h('div', { class: 'oe-hint', html: `office-editor 仅支持 DOCX / XLSX（当前：${escapeHtml(format)}）` }));
        }
        bindZoom();
        refreshOutline();
    } catch (err: unknown) {
        const msg = (err as { message?: string })?.message || String(err);
        contentEl.append(h('pre', { class: 'oe-error', text: '渲染失败：' + msg }));
        post({ type: 'error', message: msg });
    }
}

/**
 * 把 docxToHtml 产出的单页 HTML 拆成多页
 *
 * jdoc 的 HTML 形如 `<style>…<div class="docx-page"><div class="docx-body">块…</div></div>`，
 * 这里剥掉最外层纸面，取出块序列按纸张高度重新装箱，每页各自套一层纸面，
 * 于是原有的 `.docx-body` 样式（标题层级、目录点线、高亮等）原样生效。
 */
async function renderDocxPages(html: string, section: DocxSection | undefined): Promise<HTMLElement> {
    const metrics = metricsOf(section);
    // 离屏解出块序列：jdoc 自带 .docx-page 外壳，分页要自己控制纸面
    const probe = document.createElement('div');
    probe.innerHTML = html;
    const body = probe.querySelector<HTMLElement>('.docx-body') ?? probe;

    // jdoc 样式表随 HTML 一起回来，去重后挂到外壳头部（分页后仍需全局生效）
    const styleText = Array.from(probe.querySelectorAll('style'))
        .map((el) => el.textContent || '')
        .join('\n');
    if (styleText && !document.getElementById('docx-html-style')) {
        const style = document.createElement('style');
        style.id = 'docx-html-style';
        style.textContent = styleText;
        document.head.appendChild(style);
    }

    const nodes = Array.from(body.childNodes).filter((n): n is HTMLElement => n instanceof HTMLElement);
    body.textContent = '';

    const view = h('div', { class: 'oe-view oe-view-paged' });
    // 查看模式是只读渲染，允许按行拆块 + widow/orphan；编辑模式不开——拆块会破坏
    // 「块 ↔ model.blocks[i]」的对应关系，contenteditable 的光标 / 选区 / undo 会失准
    const paging = (lineBreak: boolean) =>
        paginate(nodes, {
            metrics,
            pageClass: 'oe-page docx-page',
            bodyClass: 'docx-body',
            pageNumber: true,
            lineBreak,
            mount: contentEl
        });
    // 首轮绝不能拆：拆分后的后半段是新元素、不在 nodes 里，而分页要跑两轮
    //（等图片解码后按最终高度重切），第二轮就找不到那段内容、等于凭空丢失。
    // 首轮只做块级装箱，等图片就绪再启用行级拆分。
    view.appendChild(paging(false));

    // 图片高度未就绪时分页会算错，等解码完成再按最终高度切一次
    await waitImages(view);
    if (!view.isConnected) return view;
    const pages = view.firstElementChild as HTMLElement;
    view.replaceChild(paging(true), pages);
    return view;
}

async function enterEdit(): Promise<void> {
    contentEl.innerHTML = '';
    toolbarEl.innerHTML = '';
    if (editor) {
        editor.destroy();
        editor = null;
    }
    try {
        if (format === 'docx') {
            editor = createDocxEditor(hooks);
            statusEl.textContent = '';
        } else if (format === 'xlsx') {
            const x = createXlsxEditor(hooks);
            x.setStatusHandler((s) => (statusEl.textContent = s));
            editor = x;
        } else {
            contentEl.append(h('div', { class: 'oe-hint', text: `编辑仅支持 DOCX / XLSX（当前：${format}）` }));
            return;
        }
        editor.mount(toolbarEl, contentEl);
        await editor.load(latest);
        bindZoom();
        refreshOutline();
    } catch (err: unknown) {
        const msg = (err as { message?: string })?.message || String(err);
        contentEl.append(h('pre', { class: 'oe-error', text: '加载失败：' + msg }));
        post({ type: 'error', message: msg });
    }
}

// ---------- 消息 ----------
window.addEventListener('message', async (event: MessageEvent<HostToWebview>) => {
    const msg = event.data;
    if (!msg || typeof msg !== 'object') return;
    switch (msg.type) {
        case 'init': {
            const my = ++gen; // 丢弃此后被更新的 init 所取代的异步结果
            format = msg.format;
            mode = msg.mode === 'edit' ? 'edit' : 'view';
            // 语言：init 里的 locale 是权威值（外壳建得比 init 早，文案要补一次）
            if (setLocale(msg.locale)) applyLocale();
            // 记住当前文档 uri：转换请求必须带上，否则宿主只能靠 activeTextEditor 猜
            setCurrentDocUri(msg.uri ?? null);
            latest = b64ToBytes(msg.bytes);
            original = latest;
            if (editor) {
                editor.destroy();
                editor = null;
            }
            zoomBoundTo = null; // 新文档：解除上一份文档的绑定，并把缩放复位到适应宽度
            zoom.fit();
            await applyMode();
            if (my !== gen) return;
            break;
        }
        case 'setMode':
            // 来自 VS Code 标题栏的查看/编辑按钮
            if (msg.mode !== mode) {
                mode = msg.mode;
                post({ type: 'setMode', mode });
                await applyMode();
            }
            break;
        case 'bytes': {
            latest = b64ToBytes(msg.bytes);
            original = latest;
            if (mode === 'view') await enterView();
            break;
        }
        case 'saved':
            toast('已保存');
            break;
        case 'reverted':
            break;
        case 'info':
            toast(msg.message, msg.kind || 'info');
            break;
        case 'convertAck':
            // 宿主已收到转换请求、选择框即将弹出
            notifyConvertAck();
            break;
    }
});

function updateModeHint(): void {
    if (statusEl && mode === 'view') statusEl.textContent = t('oe.readOnly');
}

// ---------- 启动 ----------
function main(): void {
    const style = document.createElement('style');
    style.textContent = CHROME_CSS + officeCss;
    document.head.appendChild(style);
    buildShell();

    // 文档内容区右键菜单：custom editor 的正文在 webview 里，VS Code 的
    // editor/context 菜单不会出现，这里自绘一份，把「转换格式」补进内容区右键。
    // 可编辑区域（编辑模式下的 contenteditable / 表格单元格 / 输入框）保留浏览器
    // 原生菜单——那里的复制、粘贴、撤销等原生行为更完整。
    installContextMenu(
        () => {
            const entries = [];
            if (hasTextSelection()) {
                entries.push({ label: t('main.copy'), action: () => copySelection(), hint: '⌘C' });
                entries.push({});
            }
            entries.push(convertMenuEntry());
            return entries;
        },
        (target) =>
            target.isContentEditable ||
            !!target.closest('[contenteditable="true"], input, textarea, .oe-tbl td, .oe-tbl th')
    );

    post({ type: 'ready' });
}

main();

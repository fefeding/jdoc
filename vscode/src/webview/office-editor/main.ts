// office-editor 入口：DOCX / XLSX 的查看（view）与编辑（edit）。
// 独立自包含入口，避免与查看版 webview.js 的双 message 监听 / 双 acquireVsCodeApi 冲突。
//
// 数据流：host 发 init(bytes) → view 用 docxToHtml/xlsxToHtml 只读渲染；edit 用块式/网格
// 编辑器载入 blocks/sheets 模型 → 编辑 → jsonToDocx/jsonToXlsx 生成字节 → sync 回 host 保存。
//
// 外壳：左侧目录/编号索引栏（大纲）+ 右侧内容区。DOCX 两个模式都按纸张高度分页，
// 每页独立成盒并带页码，观感对齐 Word/WPS；XLSX 是网格工作表，不分页、无大纲。
import { docxToHtml, xlsxToHtml, type DocxSection } from '@fefeding/jdoc';
import { officeCss } from './style';
import { CHROME_CSS } from '../chrome';
import { createZoom, createZoomCluster } from '../zoom';
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

const zoom = createZoom();
/** 当前已绑定缩放的内容根，用于跳过重复绑定 */
let zoomBoundTo: HTMLElement | null = null;

/** 左侧目录栏（XLSX 无大纲，整条隐藏） */
let outline: OutlineView | null = null;
/** 滚动联动句柄（refresh 用于缩放后重算高亮） */
let untrack: ActiveHeadingTracker | null = null;
/** 查看模式渲染代数：丢弃过期的异步分页结果 */
let viewGen = 0;

function post(msg: WebviewToHost): void {
    vscode.postMessage(msg);
}

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

    app.append(toolbarEl, h('div', { class: 'oe-body' }, [outline.el, contentEl]), h('div', { class: 'oe-statusbar' }, [statusEl, createZoomCluster(zoom)]));
    root.append(app);

    // 缩放会改变内容高度与各标题的视口位置，但缩放本身不触发 scroll 事件，
    // 必须主动重算一次，否则侧栏高亮会停在上一个标题上。
    zoom.setOnChange(() => untrack?.refresh());

    // 编辑器在增删行列 / 切工作表 / undo 时会整体重建内容根，需重新绑定缩放目标。
    // 必须监听 subtree：XLSX 重建的是 .oe-xl 内部的 .oe-grid，contentEl 的直接子节点
    // 并没变，只看 childList 会漏掉，缩放就再也绑不回新网格上。
    new MutationObserver(() => bindZoom()).observe(contentEl, { childList: true, subtree: true });
}

/** 把带缩放的文档滚动到视口内（大纲点击用） */
function scrollToAnchor(el: HTMLElement): void {
    // rect 与 scrollTop 同属「视觉」坐标系：CSS zoom 生效后，滚动容器的 scrollTop
    // 也按缩放后的距离计，两者直接相加即为目标在滚动内容中的位置，不需要再除 zoom。
    const top = el.getBoundingClientRect().top - contentEl.getBoundingClientRect().top + contentEl.scrollTop;
    contentEl.scrollTo({ top: Math.max(0, top - 12), behavior: 'smooth' });
    el.classList.add('oe-flash');
    window.setTimeout(() => el.classList.remove('oe-flash'), 900);
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
        (id) => outline?.setActive(id)
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
    // XLSX 缩放作用在网格表上：.oe-xl 是 height:100% 的 flex 容器，缩放它会把整个
    // 面板（含工作表标签与冻结区）按倍率压缩，和滚动区对不上。只缩表格本身，
    // 滚动容器 .oe-grid-wrap 保持满高，冻结行列仍按 sticky 正常吸附。
    let target: HTMLElement;
    let container: HTMLElement;
    let paper: HTMLElement;
    if (el.classList.contains('oe-xl')) {
        // .oe-grid 每次重绘（增删行列 / 切工作表 / undo）都会被整体重建，
        // 而 .oe-xl 这个根节点不变。若拿 .oe-xl 当「已绑定」判据，缩放会一直
        // 挂在已被替换掉的旧 grid 上，表现为缩放对 XLSX 完全无效。
        const grid = el.querySelector<HTMLElement>('.oe-grid');
        if (!grid) return; // 网格还没建好，等下一次内容变化再绑
        target = grid;
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
    statusEl.textContent = '查看模式（只读）';
    contentEl.innerHTML = '';
    // 解析/分页期间先给出反馈：大文档（教材、报表）在 await 期间是纯白屏，
    // 用户会以为卡死。这里先挂一个提示，解析完成后再替换。
    contentEl.append(h('div', { class: 'oe-hint', text: `正在解析 ${format.toUpperCase()}…` }));
    try {
        if (format === 'xlsx') {
            const html = (await xlsxToHtml(latest)).html;
            const view = h('div', { class: 'oe-view' });
            const page = h('div', { class: 'page' });
            page.innerHTML = html;
            view.append(page);
            contentEl.replaceChildren(view);
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
    }
});

function updateModeHint(): void {
    if (statusEl && mode === 'view') statusEl.textContent = '查看模式（只读）';
}

// ---------- 启动 ----------
function main(): void {
    const style = document.createElement('style');
    style.textContent = CHROME_CSS + officeCss;
    document.head.appendChild(style);
    buildShell();
    post({ type: 'ready' });
}

main();

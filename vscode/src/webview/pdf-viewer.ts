// PDF 阅读器：左侧缩略图栏 + 顶部页码/缩放工具栏 + 单页/连续双视图。
//
// 与其他格式的区别：PDF 已自带分页语义，不需要自己切页（那是 office-editor 的分页器
// 在做的事）。这里做的是"把已有的页"组织成可导航的阅读界面：
// - 缩略图栏：pdfjs 把每页栅格化到小 canvas（不是 CSS 缩放正文，否则几百页会拖垮布局），
//   懒渲染 + 串行队列，避免一次开文档就把主线程占满。
// - 页码：工具栏可输入跳转；单页模式下用上一页/下一页，翻页即换页。
// - 视图：默认单页（一次一页，对齐常见 PDF 阅读器），可切连续滚动。
// - 旋转：viewport.rotation 变化后页面尺寸会换（width/height 对调），必须重算占位与
//   缩略图，所以旋转不是 CSS transform，而是把渲染重走一遍。
import { createZoomCluster, type ZoomController } from './zoom';
import { createOverlayEditor, type OverlayEditor } from './pdf-editor';

/** 缩略图渲染宽度（CSS px，与 .pdf-thumb .frame 宽度一致） */
const THUMB_WIDTH = 132;
/** 缩略图栏宽度 */
const SIDEBAR_WIDTH = 196;
/** 正文页最小 / 最大宽度 */
const MIN_PAGE_WIDTH = 280;
const MAX_PAGE_WIDTH = 1200;
/** 缩略图渲染队列并发度：一次一页，避免大文档开页时卡死主线程 */
const THUMB_CONCURRENCY = 2;

/** PDF 阅读器外观（注入 webview 头部，与 baseCss 同级） */
export const PDF_CSS = `
/* ---------- 外壳：工具栏 + 缩略图栏 + 正文 ---------- */
.pdf-bar { display: flex; align-items: center; gap: 6px; flex: 0 0 auto; padding: 5px 10px; background: var(--jc-surface-2); border-bottom: 1px solid var(--jc-border); }
.pdf-bar .spring { flex: 1 1 auto; }
.pdf-bar .title { max-width: 32vw; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--jc-muted); font-size: 12px; }
.pdf-btn { appearance: none; border: 0; background: transparent; color: var(--jc-muted); cursor: pointer; width: 24px; height: 24px; border-radius: 5px; display: inline-flex; align-items: center; justify-content: center; padding: 0; flex: 0 0 auto; font-size: 13px; line-height: 1; }
.pdf-btn:hover:not(:disabled) { background: var(--jc-hover); color: var(--jc-text); }
.pdf-btn:disabled { opacity: .35; cursor: default; }
.pdf-btn.on { background: var(--jc-active); color: var(--jc-active-fg); }
.pdf-btn svg { width: 15px; height: 15px; fill: none; stroke: currentColor; stroke-width: 1.7; stroke-linecap: round; stroke-linejoin: round; }
.pdf-sep { width: 1px; height: 16px; background: var(--jc-border); flex: 0 0 auto; }
/* 页码：可输入跳转 */
.pdf-pager { display: inline-flex; align-items: center; gap: 3px; color: var(--jc-muted); font-size: 12px; }
.pdf-pager input { width: 46px; height: 22px; text-align: center; border: 1px solid var(--jc-border); border-radius: 4px; background: var(--jc-surface); color: var(--jc-text); font: inherit; padding: 0 2px; }
.pdf-pager input:focus { outline: 1px solid var(--jc-focus); outline-offset: -1px; }
.pdf-pager .total { min-width: 26px; font-variant-numeric: tabular-nums; }

.pdf-body { flex: 1 1 auto; display: flex; min-height: 0; }
/* 左侧缩略图栏 */
.pdf-thumbs { width: ${SIDEBAR_WIDTH}px; flex: 0 0 ${SIDEBAR_WIDTH}px; overflow-y: auto; overflow-x: hidden; background: var(--jc-surface-2); border-right: 1px solid var(--jc-border); padding: 10px 0 24px; }
.pdf-thumbs[hidden] { display: none; }
.pdf-thumbs::-webkit-scrollbar { width: 10px; }
.pdf-thumbs::-webkit-scrollbar-thumb { background: var(--vscode-scrollbarSlider-background, rgba(121,121,121,.4)); border-radius: 6px; border: 2px solid transparent; background-clip: content-box; }
.pdf-thumb { appearance: none; border: 0; background: transparent; cursor: pointer; display: flex; flex-direction: column; align-items: center; gap: 4px; padding: 6px 8px; width: 100%; color: var(--jc-muted); font: inherit; }
.pdf-thumb:hover { background: var(--jc-hover); }
/* 当前页：靠缩略图边框与底色区分，不再整行高亮（整行色块在深色主题下过于抢眼） */
.pdf-thumb.active { background: var(--jc-hover); color: var(--jc-text); }
.pdf-thumb .frame { width: ${THUMB_WIDTH}px; background: #fff; border: 1px solid var(--jc-border); box-shadow: 0 1px 2px rgba(0,0,0,.25); overflow: hidden; display: block; line-height: 0; }
.pdf-thumb.active .frame { border-color: var(--jc-accent); box-shadow: 0 0 0 2px var(--jc-accent); }
.pdf-thumb .frame canvas { display: block; width: ${THUMB_WIDTH}px; }
.pdf-thumb .frame .ph { display: block; height: 60px; }
.pdf-thumb .num { font-size: 11px; font-variant-numeric: tabular-nums; }

/* ---------- 正文页 ---------- */
.pdf-content { padding: 16px 24px 32px; }
.pdf-pages { display: flex; flex-direction: column; align-items: center; }
.pdf-page { background: #fff; border: 1px solid var(--jc-border, #d0d7de); margin: 0 auto 16px; box-shadow: 0 1px 3px rgba(0,0,0,.1); }
.pdf-page .page-text { white-space: pre-wrap; word-break: break-word; }
.pdf-canvas-page { padding: 0; overflow: hidden; position: relative; line-height: 0; }
.pdf-canvas-page canvas { display: block; width: 100%; height: 100%; }
.pdf-canvas-page.pending { display: flex; align-items: center; justify-content: center; color: var(--jc-muted, #656d76); font-size: 12px; line-height: 1.6; }
.pdf-canvas-page .pdf-text-layer { position: absolute; inset: 0; color: transparent; }
.pdf-canvas-page .pdf-text-layer span { position: absolute; white-space: pre; line-height: 1; font-family: sans-serif; transform-origin: 0 0; }
/* 单页视图：只留当前页。旋转后每页宽高会变，不能靠 CSS 统一裁切 */
.pdf-pages.paged .pdf-page { display: none; }
.pdf-pages.paged .pdf-page.current { display: block; margin-bottom: 0; }

/* ---------- 编辑模式（叠加层） ---------- */
/* 工具栏独立一行：顶栏已经放了导航/缩放/视图，叠加工具再挤进去会挤爆 */
.pdf-edit-bar { display: flex; align-items: center; gap: 6px; flex: 0 0 auto; padding: 4px 10px; background: var(--jc-surface); border-bottom: 1px solid var(--jc-border); overflow-x: auto; }
.pdf-edit-bar[hidden] { display: none; }
.pdf-edit-bar .spring { flex: 1 1 auto; }
.pdf-tool { appearance: none; border: 1px solid transparent; background: transparent; color: var(--jc-muted); cursor: pointer; height: 22px; padding: 0 9px; border-radius: 4px; font: inherit; font-size: 12px; white-space: nowrap; flex: 0 0 auto; }
.pdf-tool:hover { background: var(--jc-hover); color: var(--jc-text); }
.pdf-tool.on { background: var(--jc-active); color: var(--jc-active-fg); }
.pdf-tool.primary { border-color: var(--jc-border); background: var(--jc-surface-2); color: var(--jc-text); }
.pdf-tool.primary:hover { border-color: var(--jc-accent); }
.pdf-field { color: var(--jc-muted); font-size: 12px; display: inline-flex; align-items: center; flex: 0 0 auto; }
.pdf-num { width: 46px; height: 22px; border: 1px solid var(--jc-border); border-radius: 4px; background: var(--jc-surface); color: var(--jc-text); font: inherit; font-size: 12px; text-align: center; padding: 0 2px; }
.pdf-color { width: 26px; height: 20px; padding: 0; border: 1px solid var(--jc-border); border-radius: 4px; background: transparent; cursor: pointer; flex: 0 0 auto; }

/* 叠加层：未进入编辑模式时不吃指针事件，避免挡住正文的选中与滚动 */
.pdf-ov-layer { position: absolute; inset: 0; pointer-events: none; }
.pdf-ov-layer.on { pointer-events: auto; cursor: crosshair; }
.pdf-ov { position: absolute; }
.pdf-ov-text .t { white-space: pre-wrap; word-break: break-word; font-family: inherit; }
.pdf-ov img { width: 100%; height: 100%; display: block; }
.pdf-ov svg { display: block; overflow: visible; }
.pdf-ov.hov { outline: 1px dashed var(--jc-accent); }
.pdf-ov.sel { outline: 1px solid var(--jc-accent); }
/* 8 向缩放手柄 */
.pdf-ov-h { position: absolute; width: 8px; height: 8px; background: var(--jc-surface); border: 1px solid var(--jc-accent); border-radius: 2px; }
.pdf-ov-h.h-nw { left: -5px; top: -5px; cursor: nwse-resize; }
.pdf-ov-h.h-n  { left: 50%; top: -5px; margin-left: -4px; cursor: ns-resize; }
.pdf-ov-h.h-ne { right: -5px; top: -5px; cursor: nesw-resize; }
.pdf-ov-h.h-e  { right: -5px; top: 50%; margin-top: -4px; cursor: ew-resize; }
.pdf-ov-h.h-se { right: -5px; bottom: -5px; cursor: nwse-resize; }
.pdf-ov-h.h-s  { left: 50%; bottom: -5px; margin-left: -4px; cursor: ns-resize; }
.pdf-ov-h.h-sw { left: -5px; bottom: -5px; cursor: nesw-resize; }
.pdf-ov-h.h-w  { left: -5px; top: 50%; margin-top: -4px; cursor: ew-resize; }
.pdf-ov-del { position: absolute; right: -9px; top: -9px; width: 18px; height: 18px; border-radius: 50%; background: var(--jc-danger, #d13438); color: #fff; font-size: 13px; line-height: 16px; text-align: center; cursor: pointer; user-select: none; }

/* ---------- 全文搜索 ---------- */
/* 单独一行（与编辑栏对称），不挤占顶部导航栏 */
.pdf-search-bar { display: flex; align-items: center; gap: 6px; flex: 0 0 auto; padding: 4px 10px; background: var(--jc-surface); border-bottom: 1px solid var(--jc-border); }
.pdf-search-bar[hidden] { display: none; }
.pdf-search-input { width: 220px; height: 24px; border: 1px solid var(--jc-border); border-radius: 4px; background: var(--jc-surface); color: var(--jc-text); font: inherit; font-size: 12px; padding: 0 8px; }
.pdf-search-input:focus { outline: 1px solid var(--jc-focus); outline-offset: -1px; }
.pdf-search-input.no-match { border-color: var(--jc-danger, #d13438); }
.pdf-search-count { color: var(--jc-muted); font-size: 12px; font-variant-numeric: tabular-nums; min-width: 38px; text-align: center; }
.pdf-search-bar .pdf-btn { width: 22px; height: 22px; }
.pdf-search-bar .pdf-btn:disabled { opacity: .3; }
/* 命中点：透明文本层之上，不抢事件 */
.pdf-search-layer { position: absolute; inset: 0; pointer-events: none; z-index: 2; }
.pdf-search-hit { position: absolute; background: rgba(255, 212, 0, .45); border-radius: 2px; }
.pdf-search-hit.current { background: rgba(255, 140, 0, .65); box-shadow: 0 0 0 1px rgba(255, 140, 0, .9); }
`;

function el<K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string): HTMLElementTagNameMap[K] {
    const node = document.createElement(tag);
    if (cls) node.className = cls;
    if (text !== undefined) node.textContent = text;
    return node;
}

/** 内联 SVG 图标：几个按钮不值得引图标字体或额外请求 */
function svgIcon(path: string): SVGSVGElement {
    const ns = 'http://www.w3.org/2000/svg';
    const svg = document.createElementNS(ns, 'svg');
    svg.setAttribute('viewBox', '0 0 24 24');
    const p = document.createElementNS(ns, 'path');
    p.setAttribute('d', path);
    svg.appendChild(p);
    return svg;
}

const ICON = {
    menu: 'M4 6h16M4 12h16M4 18h16',
    prev: 'M15 6l-6 6 6 6',
    next: 'M9 6l6 6-6 6',
    page: 'M7 3h7l5 5v13H7zM14 3v5h5',
    rotate: 'M4.5 12a7.5 7.5 0 1 1 2.2 5.3M4.5 12V6.5M4.5 12H10',
    edit: 'M4 20h4l10-10-4-4L4 16zM13.5 6.5l4 4',
    search: 'M11 4a7 7 0 1 0 0 14 7 7 0 0 0 0-14zM20 20l-4.3-4.3'
};

/** PDF 阅读器实例 */
export interface PdfViewer {
    /** 主内容滚动容器 */
    readonly content: HTMLElement;
    /** 打开一份 PDF */
    open(bytes: Uint8Array): Promise<void>;
    /** 顶部工具栏显示的文件名 */
    setTitle(name: string): void;
    /**
     * 切换查看 / 编辑模式（编辑模式叠加 overlay 工具层）
     *
     * @param notify 是否回调 onEditingChange（宿主据此同步标题栏的查看/编辑命令）
     */
    setEditing(on: boolean, notify?: boolean): void;
    readonly editing: boolean;
    /** 编辑产出新字节，交回 host 写盘 */
    onSave(handler: (bytes: Uint8Array) => void): void;
    /** 宿主已落盘：清掉标题上的未保存标记 */
    markSaved(): void;
    /** 卸载：销毁 pdfjs 文档、解绑监听与观察器 */
    close(): void;
}

/**
 * 构建 PDF 阅读器并挂到 app
 *
 * @param app 外壳根（调用方负责 height:100% 的布局）
 * @param zoom 共用缩放控制器；本模块自行绑定 paper 为当前页
 * @param onStatus 状态栏文案回调
 * @param onEditingChange 用户在阅读器里开关编辑模式时回调，供宿主同步「查看/编辑」上下文
 */
export function createPdfViewer(
    app: HTMLElement,
    zoom: ZoomController,
    onStatus: (s: string) => void,
    onEditingChange?: (editing: boolean) => void
): PdfViewer {
    // ---------- 结构与工具栏 ----------
    const bar = el('div', 'pdf-bar');
    const body = el('div', 'pdf-body');
    const thumbs = el('div', 'pdf-thumbs');
    const content = el('div', 'jc-content pdf-content');

    const btn = (icon: string, title: string, fn: () => void): HTMLButtonElement => {
        const b = el('button', 'pdf-btn');
        b.type = 'button';
        b.title = title;
        b.setAttribute('aria-label', title);
        b.appendChild(svgIcon(icon));
        b.addEventListener('click', fn);
        return b;
    };

    // ---------- 叠加编辑层 ----------
    // 必须在 append 之前建好：它的工具栏要插在 bar 之后单独一行
    let editing = false;
    /** 有未保存的叠加内容 */
    let dirty = false;
    let saveHandler: ((bytes: Uint8Array) => void) | null = null;
    const editor = createOverlayEditor({
        // 真正的原始字节在 open() 时通过 reset 注入
        original: new Uint8Array(0),
        zoom: () => zoom.getZoom(),
        onSave: (bytes) => saveHandler?.(bytes),
        onDirty: (d) => {
            dirty = d;
            paintTitle();
        },
        onToast: (message, kind) => onStatus(`${kind === 'error' ? '✕ ' : ''}${message}`)
    });
    const editBar = editor.bar;
    editBar.hidden = true;

    // ---------- 全文搜索栏（独立一行，与编辑栏对称） ----------
    const searchBar = el('div', 'pdf-search-bar');
    searchBar.hidden = true;
    const searchInput = el('input', 'pdf-search-input') as HTMLInputElement;
    searchInput.type = 'text';
    searchInput.placeholder = '搜索全文…';
    searchInput.setAttribute('aria-label', '搜索');
    const searchCount = el('span', 'pdf-search-count', '');
    const searchPrev = btn(ICON.prev, '上一个匹配', () => gotoHit(-1));
    const searchNext = btn(ICON.next, '下一个匹配', () => gotoHit(1));
    const searchClose = btn('M6 6l12 12M18 6L6 18', '关闭搜索', () => toggleSearch(false));
    searchBar.append(searchInput, searchCount, searchPrev, searchNext, searchClose);

    body.append(thumbs, content);
    app.append(bar, editBar, searchBar, body);

    const toggleBtn = btn(ICON.menu, '折叠 / 展开缩略图栏', () => {
        thumbs.hidden = !thumbs.hidden;
        toggleBtn.classList.toggle('on', !thumbs.hidden);
        applyZoom();
    });
    const prevBtn = btn(ICON.prev, '上一页', () => step(-1));
    const nextBtn = btn(ICON.next, '下一页', () => step(1));

    const pageInput = el('input');
    pageInput.type = 'text';
    pageInput.inputMode = 'numeric';
    pageInput.title = '输入页码回车跳转';
    pageInput.setAttribute('aria-label', '页码');
    const pageTotal = el('span', 'total', '/ 0');
    const pager = el('div', 'pdf-pager');
    pager.append(prevBtn, pageInput, pageTotal, nextBtn);

    const viewBtn = btn(ICON.page, '单页 / 连续滚动', () => setPaged(!state.paged));
    const rotateBtn = btn(ICON.rotate, '顺时针旋转 90°', () => rotate());
    const editBtn = btn(ICON.edit, '编辑：叠加签名 / 图片 / 文本 / 手绘', () => setEditing(!editing));
    const searchBtn = btn(ICON.search, '搜索全文', () => toggleSearch());
    searchBtn.dataset.tool = 'search';
    const title = el('div', 'title');

    /** 标题：文件名 + 未保存标记（叠加编辑的未保存态由 webview 自行维护） */
    function paintTitle(): void {
        const name = title.dataset.name ?? '';
        title.textContent = dirty ? `● ${name}` : name;
        title.title = name;
    }

    /**
     * 切换编辑模式
     *
     * 叠加层随每页渲染时创建（未渲染的页等渲染完再建），因此这里只切
     * pointer-events 与工具栏显隐，不需要重建页面。
     */
    function setEditing(on: boolean, notify = true): void {
        if (editing === on) return;
        editing = on;
        if (on) editor.enable();
        else editor.disable();
        editBar.hidden = !on;
        editBtn.classList.toggle('on', on);
        if (notify) onEditingChange?.(on);
    }

    bar.append(
        toggleBtn,
        el('div', 'pdf-sep'),
        pager,
        createZoomCluster(zoom),
        viewBtn,
        rotateBtn,
        editBtn,
        searchBtn,
        el('div', 'spring'),
        title
    );

    // ---------- 状态 ----------
    const state = {
        doc: null as any,
        total: 0,
        current: 1,
        /** 顺时针累计旋转（0/90/180/270） */
        rotation: 0,
        /** 单页视图 */
        paged: true,
        pages: [] as HTMLElement[],
        rendered: new Set<number>(),
        thumbsRendered: new Set<number>(),
        /** 每页文本项（PDF pt transform），预取后用于搜索；下标 0 起 */
        textItems: [] as Array<Array<{ str: string; transform: number[]; width: number; height: number }> | null>
    };

    let wrap: HTMLElement | null = null;
    /** 正文页宽（px），按内容区可用宽度算 */
    let pageWidth = 800;
    let disposed = true;
    let io: IntersectionObserver | null = null;
    let thumbIo: IntersectionObserver | null = null;
    let onScroll: (() => void) | null = null;
    let onResize: (() => void) | null = null;

    /** 缩放绑定：paper 必须挂当前页，单页视图下非当前页量到的宽度是 0 */
    function applyZoom(): void {
        if (!wrap) return;
        zoom.bind({ target: wrap, container: content, paper: state.pages[state.current - 1] ?? null });
    }

    function measure(): void {
        pageWidth = Math.max(MIN_PAGE_WIDTH, Math.min(MAX_PAGE_WIDTH, (content.clientWidth || 900) - 48));
    }

    function isVisible(index: number): boolean {
        return state.paged ? index === state.current : true;
    }

    /**
     * 按当前可用宽度与旋转角修正某页占位尺寸
     *
     * 每页只在「自己即将被渲染」或「尺寸相关的状态变了」时修正一次。relayout
     * 若逐页 getPage，几百页的文档会一次性把主线程占满；未轮到的页等真正可见时
     * 由 renderPage 兜底修正，视觉上只是切到该页时宽度就位。
     */
    async function ensureSize(index: number): Promise<{ page: any; base: any; cssScale: number }> {
        const host = state.pages[index - 1];
        const page = await state.doc.getPage(index);
        const base = page.getViewport({ scale: 1, rotation: state.rotation });
        const cssScale = pageWidth / base.width;
        if (host && !disposed) {
            host.style.width = `${Math.round(base.width * cssScale)}px`;
            host.style.height = `${Math.round(base.height * cssScale)}px`;
        }
        return { page, base, cssScale };
    }

    /** 尺寸变了 canvas 位图不再匹配，需重渲染；只处理可见页与已渲染页 */
    function relayout(): void {
        if (!wrap || !state.doc) return;
        measure();
        for (let i = 1; i <= state.total; i++) {
            if (!isVisible(i) && !state.rendered.has(i)) continue;
            void ensureSize(i).then(({ base, cssScale, page }) => {
                if (disposed) return;
                state.rendered.delete(i);
                void renderPage(i, page, base, cssScale);
            });
        }
        applyZoom();
    }

    /** 渲染正文页（canvas + 透明文本层）；page/base/cssScale 可由调用方预先算好复用 */
    async function renderPage(index: number, page?: any, base?: any, cssScale?: number): Promise<void> {
        if (disposed || !state.doc || state.rendered.has(index)) return;
        const host = state.pages[index - 1];
        if (!host) return;
        state.rendered.add(index);
        // 下面会 replaceChildren 重建页内容，旧叠加层节点随之失效。
        // 编辑层内部一律用 0 起页序（与库层 overlays[] 的下标一致）
        editor.invalidate(index - 1);
        try {
            let pg = page;
            let baseVp = base;
            let scale = cssScale;
            if (!pg || baseVp === undefined || scale === undefined) {
                const sized = await ensureSize(index);
                pg = sized.page;
                baseVp = sized.base;
                scale = sized.cssScale;
            }
            if (disposed) return;
            const dpr = Math.min(2, window.devicePixelRatio || 1);
            const viewport = pg.getViewport({ scale: scale * dpr, rotation: state.rotation });

            const canvas = el('canvas');
            canvas.width = Math.max(1, Math.floor(viewport.width));
            canvas.height = Math.max(1, Math.floor(viewport.height));
            canvas.style.width = `${Math.round(baseVp.width * scale)}px`;
            canvas.style.height = `${Math.round(baseVp.height * scale)}px`;

            host.classList.remove('pending');
            host.replaceChildren(canvas);
            const ctx = canvas.getContext('2d');
            if (!ctx) throw new Error('无法创建 canvas 上下文');
            await pg.render({ canvasContext: ctx, viewport }).promise;
            if (disposed) return;

            // 透明文本层：视觉不可见但可选中/复制。
            // 用单位视口换算而非手算 PDF 左下坐标，这样旋转后依然对齐。
            try {
                const tc = await pg.getTextContent();
                const items = ((tc.items || []) as any[]).filter((it) => typeof it?.str === 'string' && it.str);
                if (items.length) {
                    const layer = el('div', 'pdf-text-layer');
                    for (const it of items) {
                        const t = (it.transform || []) as number[];
                        const [vx, vy] = baseVp.convertToViewportPoint(t[4] || 0, t[5] || 0);
                        const fs = Math.abs(t[3] || it.height || 0) || 10;
                        const span = el('span');
                        span.textContent = it.str;
                        span.style.left = `${vx * scale}px`;
                        span.style.top = `${Math.max(0, vy - fs) * scale}px`;
                        span.style.fontSize = `${fs * scale}px`;
                        layer.appendChild(span);
                    }
                    host.appendChild(layer);
                }
            } catch {
                // 文本层失败不影响图形渲染
            }
            // 叠加编辑层挂在页内（canvas / 文本层之上），坐标换算由 editor 自己做。
            // 尺寸单位：cssWidth 为布局 px，宽高为 PDF 点（baseVp 是 scale=1 的视口）。
            editor.paint(index - 1, host, Math.round(baseVp.width * scale), baseVp.width, baseVp.height);
            onStatus(`共 ${state.total} 页`);
            // 搜索高亮依赖坐标换算，页（重）渲染后需要按当前命中重画
            if (query && hits.length) void paintSearch();
        } catch (err: any) {
            state.rendered.delete(index);
            if (disposed) return;
            host.classList.remove('pending');
            host.textContent = `第 ${index} 页渲染失败：${String(err?.message || err)}`;
        }
    }

    // ---------- 缩略图（串行队列） ----------
    const thumbQueue: number[] = [];
    let thumbBusy = 0;

    function pumpThumbs(): void {
        while (thumbBusy < THUMB_CONCURRENCY && thumbQueue.length) {
            const index = thumbQueue.shift()!;
            thumbBusy += 1;
            void renderThumb(index).finally(() => {
                thumbBusy -= 1;
                pumpThumbs();
            });
        }
    }

    async function renderThumb(index: number): Promise<void> {
        if (disposed || !state.doc || state.thumbsRendered.has(index)) return;
        state.thumbsRendered.add(index);
        const frame = thumbs.querySelector<HTMLElement>(`.pdf-thumb[data-page="${index}"] .frame`);
        if (!frame) return;
        try {
            const page = await state.doc.getPage(index);
            if (disposed) return;
            const base = page.getViewport({ scale: 1, rotation: state.rotation });
            const scale = THUMB_WIDTH / base.width;
            const dpr = Math.min(2, window.devicePixelRatio || 1);
            const viewport = page.getViewport({ scale: scale * dpr, rotation: state.rotation });

            const canvas = el('canvas');
            canvas.width = Math.max(1, Math.floor(viewport.width));
            canvas.height = Math.max(1, Math.floor(viewport.height));
            canvas.style.height = `${Math.round(base.height * scale)}px`;
            const ctx = canvas.getContext('2d');
            if (!ctx) return;
            await page.render({ canvasContext: ctx, viewport }).promise;
            if (disposed) return;
            frame.replaceChildren(canvas);
        } catch {
            // 缩略图失败保持占位灰块，不影响正文
        }
    }

    function buildThumbs(): void {
        thumbs.replaceChildren();
        thumbQueue.length = 0;
        for (let i = 1; i <= state.total; i++) {
            const item = el('button', 'pdf-thumb');
            item.type = 'button';
            item.dataset.page = String(i);
            item.title = `第 ${i} 页`;
            const frame = el('span', 'frame');
            frame.appendChild(el('span', 'ph'));
            item.append(frame, el('span', 'num', String(i)));
            item.addEventListener('click', () => goTo(i));
            thumbs.appendChild(item);
        }
        // 重建后把当前页高亮补回去（旋转会重建整栏）
        syncThumbs();
        thumbIo?.disconnect();
        thumbIo = new IntersectionObserver(
            (entries) => {
                for (const e of entries) {
                    if (!e.isIntersecting) continue;
                    thumbIo!.unobserve(e.target as Element);
                    thumbQueue.push(Number((e.target as HTMLElement).dataset.page));
                }
                pumpThumbs();
            },
            { root: thumbs, rootMargin: '300px 0px' }
        );
        thumbs.querySelectorAll('.pdf-thumb').forEach((n) => thumbIo!.observe(n));
    }

    /** 同步缩略图当前页高亮 */
    function syncThumbs(): void {
        for (const t of Array.from(thumbs.querySelectorAll('.pdf-thumb'))) {
            t.classList.toggle('active', Number((t as HTMLElement).dataset.page) === state.current);
        }
    }

    // ---------- 导航 ----------
    function setPaged(on: boolean): void {
        state.paged = on;
        wrap?.classList.toggle('paged', on);
        viewBtn.classList.toggle('on', on);
        if (!on) goTo(state.current, false);
        relayout();
        void renderPage(state.current);
        void renderThumb(state.current);
        if (query) void paintSearch();
    }

    function goTo(index: number, scroll = true): void {
        if (!state.total) return;
        const n = Math.min(state.total, Math.max(1, Math.round(index)));
        const changed = n !== state.current;
        state.current = n;
        pageInput.value = String(n);
        pageTotal.textContent = `/ ${state.total}`;
        prevBtn.disabled = n <= 1;
        nextBtn.disabled = n >= state.total;

        for (const p of state.pages) p.classList.remove('current');
        state.pages[n - 1]?.classList.add('current');

        syncThumbs();
        if (changed) {
            thumbs.querySelector<HTMLElement>(`.pdf-thumb[data-page="${n}"]`)?.scrollIntoView({ block: 'nearest' });
        }

        void renderPage(n);
        void renderThumb(n);
        if (!scroll) return;
        if (state.paged) {
            content.scrollTop = 0;
        } else {
            const top = (state.pages[n - 1]?.offsetTop ?? 0) - 16;
            content.scrollTo({ top: Math.max(0, top), behavior: 'smooth' });
        }
        applyZoom();
    }

    function step(delta: number): void {
        goTo(state.current + delta);
    }

    function rotate(): void {
        state.rotation = (state.rotation + 90) % 360;
        // 旋转换掉每页的尺寸与位图：正文与缩略图都得重来
        state.rendered.clear();
        state.thumbsRendered.clear();
        buildThumbs();
        relayout();
        if (query) void paintSearch();
    }

    /** 连续滚动模式下按视口位置反推当前页 */
    let syncing = false;
    function syncFromScroll(): void {
        if (syncing || state.paged || !state.total || !wrap) return;
        syncing = true;
        requestAnimationFrame(() => {
            syncing = false;
            if (disposed || state.paged) return;
            const top = content.getBoundingClientRect().top + 24;
            let found = state.current;
            for (let i = 1; i <= state.total; i++) {
                const p = state.pages[i - 1];
                if (!p) continue;
                if (p.getBoundingClientRect().top <= top) found = i;
                else break;
            }
            if (found !== state.current) goTo(found, false);
        });
    }

    function onKey(e: KeyboardEvent): void {
        if (!state.total) return;
        const t = e.target as HTMLElement | null;
        if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable)) return;
        if (e.metaKey || e.ctrlKey || e.altKey) return;
        if (e.key === 'PageDown') {
            e.preventDefault();
            state.paged ? step(1) : content.scrollBy({ top: content.clientHeight });
        } else if (e.key === 'PageUp') {
            e.preventDefault();
            state.paged ? step(-1) : content.scrollBy({ top: -content.clientHeight });
        } else if (e.key === 'Home') {
            e.preventDefault();
            goTo(1);
        } else if (e.key === 'End') {
            e.preventDefault();
            goTo(state.total);
        } else if (state.paged && e.key === 'ArrowDown' && e.shiftKey) {
            e.preventDefault();
            step(1);
        } else if (state.paged && e.key === 'ArrowUp' && e.shiftKey) {
            e.preventDefault();
            step(-1);
        }
    }

    // ---------- 全文搜索 ----------
    let query = '';
    /** 命中点：{ page: 0 起以外的 1 起页序, item: 文本项, subStart/subLen: 在 item.str 内的子串偏移 } */
    let hits: Array<{ page: number; item: any; subStart: number; subLen: number }> = [];
    let activeHit = -1;
    let searchTimer: number | null = null;

    function updateCount(): void {
        searchCount.textContent = hits.length ? `${activeHit + 1}/${hits.length}` : '0/0';
    }

    /** 打开文档后预取每页文本项（PDF pt transform），搜索时不必再 parse 一遍 */
    async function prefetchText(): Promise<void> {
        if (!state.doc) return;
        state.textItems = new Array(state.total).fill(null);
        await Promise.all(
            Array.from({ length: state.total }, async (_, i) => {
                try {
                    const pg = await state.doc.getPage(i + 1);
                    const tc = await pg.getTextContent();
                    state.textItems[i] = ((tc.items || []) as any[])
                        .filter((it: any) => typeof it?.str === 'string' && it.str)
                        .map((it: any) => ({ str: it.str, transform: it.transform || [], width: it.width || 0, height: it.height || 0 }));
                } catch {
                    state.textItems[i] = [];
                }
            })
        );
        if (disposed) return;
        if (query) void paintSearch();
    }

    /** 把命中画到对应页：复用文本层同一套坐标换算（含旋转），高亮矩形轴对齐近似 */
    async function paintSearch(): Promise<void> {
        if (disposed) return;
        for (const host of state.pages) host.querySelector('.pdf-search-layer')?.remove();
        if (!query || !hits.length) return;
        const byPage = new Map<number, typeof hits>();
        for (const h of hits) {
            if (!byPage.has(h.page)) byPage.set(h.page, []);
            byPage.get(h.page)!.push(h);
        }
        for (const [page, list] of byPage) {
            const host = state.pages[page - 1];
            if (!host) continue;
            const sized = await ensureSize(page);
            if (disposed) return;
            const layer = el('div', 'pdf-search-layer');
            for (const h of list) {
                const [vx, vy] = sized.base.convertToViewportPoint(h.item.transform[4] || 0, h.item.transform[5] || 0);
                const fs = Math.abs(h.item.transform[3] || h.item.height || 0) || 10;
                const itemLeft = vx * sized.cssScale;
                const itemTop = Math.max(0, vy - fs) * sized.cssScale;
                const denom = Math.max(1, h.item.str.length);
                const subLeft = itemLeft + (h.item.width || 0) * (h.subStart / denom) * sized.cssScale;
                const subW = (h.item.width || 0) * (h.subLen / denom) * sized.cssScale;
                const div = el('div', 'pdf-search-hit' + (hits[activeHit] === h ? ' current' : ''));
                div.style.left = `${subLeft}px`;
                div.style.top = `${itemTop}px`;
                div.style.width = `${Math.max(2, subW)}px`;
                div.style.height = `${fs * sized.cssScale}px`;
                layer.appendChild(div);
            }
            host.appendChild(layer);
        }
    }

    function runSearch(q: string): void {
        query = q.trim();
        hits = [];
        activeHit = -1;
        if (!query) {
            searchInput.classList.remove('no-match');
            searchCount.textContent = '';
            searchPrev.disabled = true;
            searchNext.disabled = true;
            void paintSearch();
            return;
        }
        const re = (() => {
            try {
                return new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
            } catch {
                return null;
            }
        })();
        if (!re) {
            void paintSearch();
            return;
        }
        for (let i = 0; i < state.textItems.length; i++) {
            const items = state.textItems[i];
            if (!items || !items.length) continue;
            // 拼接本页文本：相邻 item 间补一个空格（近似词边界），并记录字符 -> item 映射
            let text = '';
            const spans: Array<{ start: number; end: number; item: any }> = [];
            for (const it of items) {
                const start = text.length;
                if (start > 0) text += ' ';
                const realStart = text.length;
                text += it.str;
                spans.push({ start: realStart, end: text.length, item: it });
            }
            re.lastIndex = 0;
            let m: RegExpExecArray | null;
            while ((m = re.exec(text))) {
                const s = m.index;
                const e = m.index + m[0].length;
                for (const sp of spans) {
                    if (sp.end <= s || sp.start >= e) continue;
                    const subStart = Math.max(0, s - sp.start);
                    const subLen = Math.min(e, sp.end) - Math.max(s, sp.start);
                    if (subLen > 0) hits.push({ page: i + 1, item: sp.item, subStart, subLen });
                }
                if (m.index === re.lastIndex) re.lastIndex++;
            }
        }
        const has = hits.length > 0;
        searchInput.classList.toggle('no-match', !has);
        searchPrev.disabled = !has;
        searchNext.disabled = !has;
        if (has) {
            activeHit = 0;
            updateCount();
            void paintSearch().then(() => gotoHit(0, true));
        } else {
            searchCount.textContent = '0/0';
            void paintSearch();
        }
    }

    function gotoHit(delta: number, initial = false): void {
        if (!hits.length) return;
        if (!initial) activeHit = (activeHit + delta + hits.length) % hits.length;
        updateCount();
        const h = hits[activeHit];
        void paintSearch().then(() => {
            const host = state.pages[h.page - 1];
            if (!host) return;
            const div = host.querySelector('.pdf-search-layer .pdf-search-hit.current') as HTMLElement | null;
            if (state.current !== h.page) goTo(h.page);
            if (div) requestAnimationFrame(() => div.scrollIntoView({ block: 'center' }));
        });
    }

    function clearSearch(): void {
        query = '';
        hits = [];
        activeHit = -1;
        searchInput.value = '';
        searchInput.classList.remove('no-match');
        searchCount.textContent = '';
        searchPrev.disabled = true;
        searchNext.disabled = true;
        void paintSearch();
    }

    function toggleSearch(force?: boolean): void {
        const on = force === undefined ? searchBar.hidden : force;
        searchBar.hidden = !on;
        searchBtn.classList.toggle('on', on);
        if (on) {
            searchInput.focus();
            if (query) runSearch(query);
        } else {
            clearSearch();
        }
    }

    searchInput.addEventListener('input', () => {
        if (searchTimer) clearTimeout(searchTimer);
        searchTimer = window.setTimeout(() => runSearch(searchInput.value), 200);
    });
    searchInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            gotoHit(e.shiftKey ? -1 : 1);
        } else if (e.key === 'Escape') {
            e.preventDefault();
            toggleSearch(false);
        }
    });

    // ---------- 打开 / 关闭 ----------
    function close(): void {
        disposed = true;
        io?.disconnect();
        io = null;
        thumbIo?.disconnect();
        thumbIo = null;
        thumbQueue.length = 0;
        thumbBusy = 0;
        editor.disable();
        editor.reset(new Uint8Array(0));
        editBar.hidden = !editing;
        editBtn.classList.toggle('on', editing);
        searchBar.hidden = true;
        searchBtn.classList.remove('on');
        clearSearch();
        state.textItems = [];
        if (onScroll) content.removeEventListener('scroll', onScroll);
        onScroll = null;
        if (onResize) window.removeEventListener('resize', onResize);
        onResize = null;
        window.removeEventListener('keydown', onKey);
        if (state.doc) {
            void state.doc.destroy?.();
            state.doc = null;
        }
        state.total = 0;
        state.current = 1;
        state.pages = [];
        state.rendered.clear();
        state.thumbsRendered.clear();
        wrap = null;
        content.replaceChildren();
        thumbs.replaceChildren();
        pageInput.value = '';
        pageTotal.textContent = '/ 0';
        prevBtn.disabled = true;
        nextBtn.disabled = true;
    }

    async function open(bytes: Uint8Array): Promise<void> {
        // 记住编辑态：close() 会拆掉编辑层，重开时要恢复用户所处的模式
        const wasEditing = editing;
        close();
        disposed = false;
        onStatus('正在打开 PDF…');
        // 编辑层的幂等保存基准：始终是这份打开时的字节
        editor.reset(bytes);
        // 与 jdoc 内部使用同一说明符，命中同一模块实例
        const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
        const worker = (window as any).__JDC_PDF_WORKER__;
        if (worker && pdfjs.GlobalWorkerOptions && !pdfjs.GlobalWorkerOptions.workerSrc) {
            pdfjs.GlobalWorkerOptions.workerSrc = worker;
        }
        // pdfjs 会把 data 转移给 worker（原 ArrayBuffer 会被 detach），必须传副本
        const doc: any = await pdfjs.getDocument({ data: bytes.slice() }).promise;
        if (disposed) {
            void doc.destroy?.();
            return;
        }
        state.doc = doc;
        state.total = doc.numPages || 0;
        state.current = 1;
        state.rotation = 0;
        state.rendered.clear();
        state.thumbsRendered.clear();

        measure();
        wrap = el('div', 'pdf-pages');
        // 先按首页比例撑占位，避免渲染完成时布局跳动
        const first = await doc.getPage(1);
        if (disposed) return;
        const ratio = first.getViewport({ scale: 1 }).height / first.getViewport({ scale: 1 }).width;
        for (let i = 1; i <= state.total; i++) {
            const p = el('div', 'pdf-page pdf-canvas-page pending');
            p.dataset.page = String(i);
            p.textContent = `第 ${i} 页`;
            p.style.width = `${pageWidth}px`;
            p.style.height = `${Math.round(pageWidth * ratio)}px`;
            wrap.appendChild(p);
            state.pages.push(p);
        }
        wrap.classList.toggle('paged', state.paged);
        content.replaceChildren(wrap);
        buildThumbs();

        // 连续模式靠观察器懒渲染正文页；单页模式由翻页主动渲染
        io = new IntersectionObserver(
            (entries) => {
                for (const e of entries) {
                    if (!e.isIntersecting) continue;
                    io!.unobserve(e.target as Element);
                    void renderPage(Number((e.target as HTMLElement).dataset.page));
                }
            },
            { root: content, rootMargin: '800px 0px' }
        );
        state.pages.forEach((p) => io!.observe(p));

        onScroll = () => syncFromScroll();
        content.addEventListener('scroll', onScroll, { passive: true });
        onResize = () => relayout();
        window.addEventListener('resize', onResize);
        window.addEventListener('keydown', onKey);

        viewBtn.classList.toggle('on', state.paged);
        toggleBtn.classList.toggle('on', !thumbs.hidden);
        setEditing(wasEditing, false);
        goTo(1);
        await renderPage(1);
        onStatus(`共 ${state.total} 页`);
        // 后台预取文本项，搜索时无需再 parse
        void prefetchText();
    }

    return {
        content,
        open,
        close,
        setEditing,
        get editing(): boolean {
            return editing;
        },
        onSave(handler: (bytes: Uint8Array) => void): void {
            saveHandler = handler;
        },
        markSaved(): void {
            dirty = false;
            paintTitle();
        },
        setTitle(name: string): void {
            title.dataset.name = name;
            paintTitle();
        }
    };
}

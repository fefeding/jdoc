// 缩放：所有文档格式（DOCX / XLSX / PDF / Markdown / 纯文本）共用的缩放控制器与状态栏缩放簇。
//
// 几个关键取舍：
// 1. 用 CSS `zoom` 而不是 `transform: scale()`。transform 方案必须再套一层元素来撑开
//    滚动区（否则缩小后留白、放大后被裁），而且会让 contenteditable 的光标定位、
//    XLSX 的 sticky 冻结行列发生偏移。`zoom` 直接参与布局计算，缩放后滚动区、
//    冻结行列与光标都保持自洽——VS Code webview 就是 Chromium，原生支持 `zoom`。
// 2. 默认「适应宽度」：内容比可视区宽时按比例缩小，避免宽表格 / 宽文档出屏外。
//    但上限锁在 100%——文档类内容被拉伸会失真，不如交给用户手动放大。
// 3. 用户手动设定倍率后保持该倍率（切窗口大小不跟随），只有打开新文档时才复位，
//    与 pptx-editor 的行为一致。

const MIN_ZOOM = 0.2;
const MAX_ZOOM = 4;
/** fit 时在可视宽度里预留的留白，避免纸面紧贴滚动条 */
const FIT_GUTTER = 32;

export interface ZoomBinding {
    /** 被缩放的元素 */
    target: HTMLElement | null;
    /** 可视区域：提供 clientWidth 供 fit 计算，并限定 Ctrl/⌘+滚轮的作用范围 */
    container?: HTMLElement | null;
    /** 纸张元素：量其自然宽度做 fit；缺省用 target 自身 */
    paper?: HTMLElement | null;
}

export interface ZoomController {
    /** 绑定缩放目标（查看 / 编辑 / 内容重建后都要重新绑定） */
    bind(b: ZoomBinding): void;
    /** 适应宽度，并切回自动模式 */
    fit(): void;
    /** 100% 实际大小 */
    actual(): void;
    /** 按倍率步进（factor > 1 为放大） */
    step(factor: number): void;
    getZoom(): number;
    isAuto(): boolean;
    setOnChange(fn: (zoom: number, auto: boolean) => void): void;
    destroy(): void;
}

function clamp(v: number, lo: number, hi: number): number {
    return Math.max(lo, Math.min(hi, v));
}

export function createZoom(): ZoomController {
    let target: HTMLElement | null = null;
    let container: HTMLElement | null = null;
    let paper: HTMLElement | null = null;
    /** null 表示自动适应宽度 */
    let userZoom: number | null = null;
    let zoom = 1;
    let onChange: ((zoom: number, auto: boolean) => void) | null = null;

    /**
     * 纸张的自然（未缩放）宽度。
     * getBoundingClientRect 返回的是缩放后的视觉宽度，需除以当前倍率还原；
     * offsetWidth 在 zoom 下的语义依赖浏览器实现，用 rect/zoom 更稳妥。
     */
    function naturalWidth(): number {
        const el = paper && paper.isConnected ? paper : target;
        if (!el || !el.isConnected) return 0;
        const w = el.getBoundingClientRect().width;
        return w > 0 ? w / (zoom || 1) : 0;
    }

    function fitZoom(): number {
        const nat = naturalWidth();
        if (!container || !nat) return 1;
        const avail = container.clientWidth - FIT_GUTTER;
        if (avail <= 0) return 1;
        return clamp(avail / nat, MIN_ZOOM, 1);
    }

    function apply(): void {
        zoom = userZoom == null ? fitZoom() : userZoom;
        if (target) {
            // 取 4 位小数，避免浮点噪声写出 0.8500000000000001 这类值
            const z = Math.round(zoom * 1e4) / 1e4;
            target.style.zoom = z === 1 ? '' : String(z);
        }
        onChange?.(zoom, userZoom == null);
    }

    function fit(): void {
        userZoom = null;
        apply();
    }

    function actual(): void {
        userZoom = 1;
        apply();
    }

    function step(factor: number): void {
        userZoom = clamp((userZoom ?? zoom) * factor, MIN_ZOOM, MAX_ZOOM);
        apply();
    }

    function bind(b: ZoomBinding): void {
        target = b.target;
        container = b.container ?? null;
        paper = b.paper ?? null;
        apply();
    }

    // Ctrl/⌘ + 滚轮缩放（与 pptx-editor 一致）；普通滚轮仍是上下滚动
    function onWheel(e: WheelEvent): void {
        if (!e.ctrlKey && !e.metaKey) return;
        const c = container;
        if (!c || !(e.target instanceof Node) || !c.contains(e.target)) return;
        e.preventDefault();
        step(e.deltaY < 0 ? 1.1 : 1 / 1.1);
    }

    function onResize(): void {
        // 只在自动模式下跟随窗口重算，用户手动设定的倍率保持不变
        if (userZoom == null) apply();
    }

    // 每个实例各自在 window 上挂监听，destroy() 用同一闭包引用移除，故多个控制器可安全并存
    // （不再依赖「模块级单例」这一隐性约定；不调用 destroy 才会泄漏）。
    window.addEventListener('wheel', onWheel, { passive: false });
    window.addEventListener('resize', onResize);

    return {
        bind,
        fit,
        actual,
        step,
        getZoom: () => zoom,
        isAuto: () => userZoom == null,
        setOnChange: (fn) => {
            onChange = fn;
        },
        destroy: () => {
            window.removeEventListener('wheel', onWheel);
            window.removeEventListener('resize', onResize);
            onChange = null;
        }
    };
}

/**
 * 状态栏缩放簇：− / 适应 / 百分比 / ＋。点击百分比复位到 100%。
 * 内部接管 controller 的 onChange，各入口只需把它挂到状态栏尾部。
 */
export function createZoomCluster(ctl: ZoomController, labels?: { fit?: string }): HTMLElement {
    const el = document.createElement('span');
    el.className = 'jc-zoom';

    const chip = (label: string, title: string, fn: () => void): HTMLButtonElement => {
        const b = document.createElement('button');
        b.type = 'button';
        b.className = 'chip';
        b.textContent = label;
        b.title = title;
        b.addEventListener('click', fn);
        return b;
    };

    const outBtn = chip('−', '缩小', () => ctl.step(1 / 1.1));
    const fitBtn = chip(labels?.fit || '适应', '适应宽度', () => ctl.fit());
    const labelBtn = chip('100%', '实际大小（100%）', () => ctl.actual());
    const inBtn = chip('+', '放大', () => ctl.step(1.1));
    el.append(outBtn, fitBtn, labelBtn, inBtn);

    const sync = (): void => {
        labelBtn.textContent = Math.round(ctl.getZoom() * 100) + '%';
        fitBtn.classList.toggle('on', ctl.isAuto());
    };
    ctl.setOnChange(sync);
    sync();
    return el;
}
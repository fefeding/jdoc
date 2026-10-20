// 缩放：所有文档格式（DOCX / XLSX / PDF / Markdown / 纯文本）共用的缩放控制器与状态栏缩放簇。
//
// 几个关键取舍：
// 1. 用 CSS `zoom` 而不是 `transform: scale()`。transform 方案必须再套一层元素来撑开
//    滚动区（否则缩小后留白、放大后被裁），而且会让 contenteditable 的光标定位、
//    XLSX 的 sticky 冻结行列发生偏移。`zoom` 直接参与布局计算，缩放后滚动区、
//    冻结行列与光标都保持自洽——VS Code webview 就是 Chromium，原生支持 `zoom`。
// 2. 默认「适应宽度」：**只缩不放**。内容比可视区宽时按屏宽等比缩小；内容本身就
//    不宽（如 A4 约 794px、Markdown 列宽上限 960px）则保持 100% 实际大小、居中留白。
//    曾把上限设成 1.5 让窄页面放大铺满，结果宽屏下 A4 被拉到 1.48、Markdown/PDF
//    更是顶到 1.5，正文行长远超舒适区间（60~80 字符）、字也偏大。放大留给用户手动
//    Ctrl/⌘+滚轮或点「＋」（上限 MAX_ZOOM）。
// 3. 用户手动设定倍率后保持该倍率（切窗口大小不跟随），只有打开新文档时才复位，
//    与 pptx-editor 的行为一致。

import { t } from './i18n';

const MIN_ZOOM = 0.2;
const MAX_ZOOM = 4;
/** fit 时在可视宽度里预留的留白，避免纸面紧贴滚动条 */
const FIT_GUTTER = 32;
/**
 * fit 的倍率上限：100%（即 fit 永不放大）。
 *
 * 目标宽度取「可视宽（扣留白）」与「内容自然宽」的较小者，所以宽屏上文档保持实际
 * 大小居中显示，而不是被拉伸铺满一屏——铺满既让行宽过长难读，字也偏大。
 */
const FIT_MAX_ZOOM = 1;

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
    let onChange: Array<(zoom: number, auto: boolean) => void> = [];

    /**
     * 纸张的自然（未缩放）宽度。
     *
     * 必须用 offsetWidth，不能用 `getBoundingClientRect().width / zoom`：后者是
     * **自引用**写法——分母正是要算出来的 zoom。一旦 paper 解析失败、退化成撑满
     * 容器的元素（解析提示、被重建的中间态），fit 就退化成
     * `(container - gutter) * zoom / container` 的收缩反馈：每绑定一次倍率就乘
     * 一个略小于 1 的系数，多次触发后一路收敛到 MIN_ZOOM（实测出现过 20%，
     * 页面被缩成一条）。offsetWidth 量的是元素**自身坐标系**的布局宽度，不随
     * 祖先 zoom 浮动，是稳定的自然宽度，也让 fit 变成幂等。
     *
     * 例外：paper 就是 target 自身时（XLSX 网格，zoom 直接作用其上），
     * offsetWidth 会随 zoom 变化，此时除回 zoom。
     */
    function naturalWidth(): number {
        const el = paper && paper.isConnected ? paper : target;
        if (!el || !el.isConnected) return 0;
        const w = el.offsetWidth;
        if (w <= 0) return 0;
        return el === target ? w / (zoom || 1) : w;
    }

    function fitZoom(): number {
        const nat = naturalWidth();
        if (!container || !nat) return 1;
        const avail = container.clientWidth - FIT_GUTTER;
        if (avail <= 0) return 1;
        return clamp(avail / nat, MIN_ZOOM, FIT_MAX_ZOOM);
    }

    function apply(): void {
        zoom = userZoom == null ? fitZoom() : userZoom;
        if (target) {
            // 取 4 位小数，避免浮点噪声写出 0.8500000000000001 这类值
            const z = Math.round(zoom * 1e4) / 1e4;
            target.style.zoom = z === 1 ? '' : String(z);
        }
        onChange.forEach((fn) => fn(zoom, userZoom == null));
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
            onChange.push(fn);
        },
        destroy: () => {
            window.removeEventListener('wheel', onWheel);
            window.removeEventListener('resize', onResize);
            onChange = [];
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

    const outBtn = chip('−', t('zoom.out'), () => ctl.step(1 / 1.1));
    const fitBtn = chip(labels?.fit || t('zoom.fit'), t('zoom.fitWidth'), () => ctl.fit());
    const labelBtn = chip('100%', t('zoom.actual'), () => ctl.actual());
    const inBtn = chip('+', t('zoom.in'), () => ctl.step(1.1));
    el.append(outBtn, fitBtn, labelBtn, inBtn);

    const sync = (): void => {
        labelBtn.textContent = Math.round(ctl.getZoom() * 100) + '%';
        fitBtn.classList.toggle('on', ctl.isAuto());
    };
    ctl.setOnChange(sync);
    sync();
    return el;
}
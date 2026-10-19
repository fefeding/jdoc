// 分页：把一串顺序内容按纸张高度贪心装箱成一叠「页」，逼近 Word/WPS 的分页观感。
//
// 为什么不用 CSS `break-inside` / `@page`：webview 不是打印上下文，浏览器不会为滚动
// 视口做分页（只有打印时才生效），所以只能自己按纸张高度把内容切到多个页盒里。
//
// 两个关键取舍：
// 1. 先在离屏容器里量一次每个块的高度（而不是边塞边测）。量一次可避免块被移动到页里
//    之后又回头改尺寸导致的反复回流，也保证图片加载完成后只需要整体量一次。
// 2. 计入相邻块之间的外边距间距。块级高度（offsetHeight）不含 margin，直接累加会让
//    每页比 Word 多塞几行，分页位置整体上移。

/** 纸张与页边距（磅，与 OOXML w:pgSz / w:pgMar 同单位） */
export interface PageMetrics {
    /** 页宽（磅） */
    width: number;
    /** 页高（磅） */
    height: number;
    /** 上边距 */
    marginTop: number;
    /** 下边距 */
    marginBottom: number;
    /** 左边距 */
    marginLeft: number;
    /** 右边距 */
    marginRight: number;
}

/** A4 纵向、1 英寸页边距 */
export const A4_PORTRAIT: PageMetrics = {
    width: 595.28,
    height: 841.89,
    marginTop: 72,
    marginBottom: 72,
    marginLeft: 72,
    marginRight: 72
};

/** 分页选项 */
export interface PaginateOptions {
    /** 纸张设置，缺省 A4 */
    metrics?: PageMetrics;
    /** 页根元素类名 */
    pageClass?: string;
    /** 页内正文容器类名（承载切分后的内容，尺寸由分页决定） */
    bodyClass?: string;
    /** 是否显示页码 */
    pageNumber?: boolean;
    /** 每页额外的自定义内容（页眉/水印等） */
    decorate?: (page: HTMLElement, index: number) => void;
    /**
     * 量测容器的挂载点（通常是内容滚动容器）
     *
     * 必须在文档内量测：游离元素的 getBoundingClientRect() 全为 0，会把所有块
     * 算成零高度而塞进第一页。挂载点要在被缩放的元素之外，量出的才是未缩放值。
     */
    mount?: HTMLElement | null;
}

/** 强制分页标记：元素带此属性时在其之前断页 */
export const PAGE_BREAK_ATTR = 'data-oe-pagebreak';

/**
 * 「与下段同页」时，标题下至少要露出的下一段高度（磅）
 *
 * 中文正文单倍行距约 15.6pt（见 w:docGrid），取 16pt 作为「首行」的估算值。
 */
const KEEP_NEXT_LEAD = 16;

/** 页码文本（页根上的 data-oe-page 属性也用它） */
export function pageLabel(index: number): string {
    return String(index + 1);
}

/** 页正文区的可用高度（磅） */
export function contentHeight(metrics: PageMetrics): number {
    return Math.max(40, metrics.height - metrics.marginTop - metrics.marginBottom);
}

/** 页正文的可用宽度（磅） */
export function contentWidth(metrics: PageMetrics): number {
    return Math.max(40, metrics.width - metrics.marginLeft - metrics.marginRight);
}

/**
 * 等待容器内所有图片完成解码
 *
 * 图片未加载时高度为 0，直接分页会把图算成"零高度"块塞进当前页，导致溢出。
 * data URL 的图片一般同 tick 就绪，但仍留一次 decode 兜底。
 */
export async function waitImages(host: HTMLElement, timeoutMs = 1500): Promise<void> {
    const imgs = Array.from(host.querySelectorAll('img'));
    if (!imgs.length) return;
    await Promise.all(
        imgs.map(
            (img) =>
                new Promise<void>((resolve) => {
                    if (img.complete && img.naturalWidth > 0) return resolve();
                    const done = (): void => {
                        img.removeEventListener('load', done);
                        img.removeEventListener('error', done);
                        resolve();
                    };
                    img.addEventListener('load', done);
                    img.addEventListener('error', done);
                    const t = window.setTimeout(done, timeoutMs);
                    img.addEventListener('load', () => window.clearTimeout(t), { once: true });
                })
        )
    );
}

/** 单个块的分页占用：进入本页需要的高度增量，以及自身高度 */
interface Slot {
    node: HTMLElement;
    /** 块自身高度（不含 margin） */
    own: number;
    /** 从上一个块之后到本块顶部占用的额外高度（含上边距） */
    gap: number;
    /** 标题块：需要与其后第一个块保持同页（Word 的「与下段同页」） */
    isHeading: boolean;
}

/**
 * 是否标题块
 *
 * 查看模式是 jdoc 渲染的原生 h1~h6；编辑模式没有 h 标签，标题是
 * `.oe-block[data-type="heading"]`。两种形态都要认，否则编辑模式分页会退化。
 */
function isHeadingNode(node: HTMLElement): boolean {
    if (/^H[1-6]$/.test(node.tagName)) return true;
    return node.dataset?.type === 'heading';
}

/**
 * 把节点按纸张高度切成一叠页，返回承载所有页的容器
 *
 * @param nodes 顺序内容节点（函数会接管它们的归属，原容器被清空）
 */
export function paginate(nodes: HTMLElement[], options: PaginateOptions = {}): HTMLElement {
    const metrics = options.metrics ?? A4_PORTRAIT;
    const pageClass = options.pageClass ?? 'oe-page';
    const bodyClass = options.bodyClass ?? 'oe-page-body';
    const avail = contentHeight(metrics);

    const host = document.createElement('div');
    host.className = 'oe-paginate-measure';
    host.setAttribute('aria-hidden', 'true');
    // 与真实页正文同宽，量出的高度才准；visibility:hidden 仍参与布局，也能触发图片加载
    host.style.cssText = `position:absolute;left:-100000px;top:0;width:${contentWidth(metrics)}pt;visibility:hidden;pointer-events:none;`;
    const mount = options.mount ?? null;
    if (mount) mount.appendChild(host);
    for (const node of nodes) host.appendChild(node);
    if (!mount) {
        // 没有挂载点时无法量测（游离元素 rect 全 0），退化为单页直出
        host.remove();
        const single = document.createElement('div');
        single.className = pageClass;
        const inner = document.createElement('div');
        inner.className = bodyClass;
        for (const node of nodes) inner.appendChild(node);
        single.appendChild(inner);
        return single;
    }

    // ---- 量一次：记录每个块的高度与相邻间距 ----
    const hostRect = host.getBoundingClientRect();
    const slots: Slot[] = [];
    let prevBottom = 0;
    for (const node of nodes) {
        const rect = node.getBoundingClientRect();
        const top = rect.top - hostRect.top;
        const bottom = rect.bottom - hostRect.top;
        slots.push({ node, own: Math.max(0, bottom - top), gap: Math.max(0, top - prevBottom), isHeading: isHeadingNode(node) });
        prevBottom = bottom;
    }
    host.remove();

    // ---- 装箱 ----
    const stack = document.createElement('div');
    stack.className = 'oe-pages';
    let body: HTMLElement | null = null;
    let used = 0;
    let index = 0;

    const newPage = (): void => {
        const page = document.createElement('div');
        page.className = pageClass;
        page.dataset.oePage = String(index);
        page.style.width = metrics.width + 'pt';
        page.style.minHeight = metrics.height + 'pt';
        page.style.padding =
            metrics.marginTop + 'pt ' + metrics.marginRight + 'pt ' + metrics.marginBottom + 'pt ' + metrics.marginLeft + 'pt';

        const inner = document.createElement('div');
        inner.className = bodyClass;
        inner.style.minHeight = avail + 'pt';
        page.appendChild(inner);
        stack.appendChild(page);

        if (options.pageNumber) {
            const num = document.createElement('div');
            num.className = 'oe-page-num';
            num.dataset.oePageNum = String(index);
            num.textContent = pageLabel(index);
            page.appendChild(num);
        }
        options.decorate?.(page, index);

        body = inner;
        used = 0;
        index++;
    };

    newPage();

    for (let i = 0; i < slots.length; i++) {
        const slot = slots[i];
        // 显式分页符：先收尾当前页，标记元素本身不渲染
        if (slot.node.getAttribute(PAGE_BREAK_ATTR) !== null) {
            slot.node.remove();
            newPage();
            continue;
        }
        // 「与下段同页」（Word 的 keepNext）：标题不能孤立在页尾。
        // 只判断标题自己放不放得下是不够的——标题塞得下、紧随其后的正文塞不下时，
        // 标题会被留在页尾而正文跑到下一页，Word/WPS 则把标题推到下一页页首。
        // 注意只要求「标题 + 下一段首行」能放下：下一段本身可以跨页续排，若按整段
        // 高度判定，长段落（必然要跨页）会把标题也顶走，反而多出空白页。
        if (slot.isHeading && used > 0) {
            const next = slots[i + 1];
            const hasNext = !!next && next.node.getAttribute(PAGE_BREAK_ATTR) === null;
            const lead = hasNext ? Math.min(next!.own, KEEP_NEXT_LEAD) : 0;
            const need2 = used + slot.gap + slot.own + (hasNext ? next!.gap + lead : 0);
            if (need2 > avail) newPage();
        }
        const need = used ? used + slot.gap + slot.own : slot.own;
        // 当前页已有内容且塞不下：断页重来；空页就装不下（超高表格/代码块）则标记后放行
        if (need > avail && used > 0) newPage();
        else if (used === 0 && slot.own > avail) slot.node.classList.add('oe-tall');
        body!.appendChild(slot.node);
        used = used ? used + slot.gap + slot.own : slot.own;
    }

    return stack;
}

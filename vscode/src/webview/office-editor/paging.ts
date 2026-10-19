// 分页：把一串顺序内容按纸张高度贪心装箱成一叠「页」，逼近 Word/WPS 的分页观感。
//
// 为什么不用 CSS `break-inside` / `@page`：webview 不是打印上下文，浏览器不会为滚动
// 视口做分页（只有打印时才生效），所以只能自己按纸张高度把内容切到多个页盒里。
//
// 三个关键取舍：
// 1. 先在离屏容器里量一次每个块的高度（而不是边塞边测）。量一次可避免块被移动到页里
//    之后又回头改尺寸导致的反复回流，也保证图片加载完成后只需要整体量一次。
// 2. 计入相邻块之间的外边距间距。块级高度（offsetHeight）不含 margin，直接累加会让
//    每页比 Word 多塞几行，分页位置整体上移。
// 3. 「与下段同页」+ widow/orphan：标题不落页尾，段落首尾不孤行。开启 lineBreak 后
//    段落允许在行边界切开（Word 的做法），代价是要在 DOM 上按行拆块。

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

/**
 * widow/orphan 的最小行数
 *
 * Word 默认开启 w:widowControl：段末行不得单独留在页首（widow），段首行不得单独
 * 留在页尾（orphan）。段落被切分时两半都至少要有这么多行，否则整段挪页。
 */
export const WIDOW_ORPHAN_LINES = 2;

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
    /**
     * 允许文本块在行边界跨页拆分，并对拆出的两半套用 widow/orphan
     *
     * 默认关闭（块级不可拆）。**编辑模式不要开**：拆分会在 DOM 上把一个块切成两个，
     * 破坏「一个块对应 model.blocks[i]」的对应关系，contenteditable 的光标定位、
     * 选区与 undo 都会失准。查看模式是只读渲染，拆开没有副作用。
     */
    lineBreak?: boolean;
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
    /** 可按行拆分（仅 lineBreak 打开时为 true） */
    splittable: boolean;
    /** 每个行末的字符偏移（元素文本流内），仅 splittable 为 true 时有值 */
    lineEnds: number[];
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

/** 该元素能否按行拆分：只认纯文本流，标题/表格/代码块/图片一律整体搬运 */
function isSplittable(node: HTMLElement): boolean {
    if (isHeadingNode(node)) return false;
    const tag = node.tagName;
    if (tag !== 'P' && tag !== 'LI' && tag !== 'BLOCKQUOTE' && tag !== 'TD' && tag !== 'TH') return false;
    // 内嵌块级元素（表格/图片/代码块）时无法安全按字符切开
    if (node.querySelector('table, pre, img, svg, br')) return false;
    return (node.textContent || '').length > 0;
}

/**
 * 量出元素内每个「行末」的字符偏移（相对整个元素的文本流）
 *
 * 浏览器只会给出行的矩形位置，不会告诉我们在 DOM 里对应哪个字符，所以这里：
 * 1. 用 Range.getClientRects() 拿到每个文本节点内部的行矩形（按 top 归并出不同行）
 * 2. 对每个行边界在对应文本节点内二分，定位到第一个落到该行的字符偏移
 *
 * 返回的偏移是「切口位置」：在 offset 处切开，前半是 offset 之前的字符。
 */
function measureLineOffsets(el: HTMLElement): number[] {
    const bounds: number[] = [];
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    const range = document.createRange();
    let base = 0;
    let node: Text | null;
    while ((node = walker.nextNode() as Text | null)) {
        const len = node.data.length;
        if (!len) continue;
        range.selectNodeContents(node);
        // 同一行可能有多个矩形（被内联元素切碎），按 top 归并
        const tops: number[] = [];
        for (const r of Array.from(range.getClientRects())) {
            if (r.height <= 0) continue;
            if (!tops.length || Math.abs(r.top - tops[tops.length - 1]) > 1) tops.push(r.top);
        }
        for (let i = 1; i < tops.length; i++) bounds.push(base + firstCharAtOrBelow(node, tops[i]));
        base += len;
    }
    return bounds;
}

/** 二分：文本节点内第一个 top ≥ 目标行 top 的字符偏移 */
function firstCharAtOrBelow(node: Text, top: number): number {
    const len = node.data.length;
    const range = document.createRange();
    let lo = 0;
    let hi = len;
    while (lo < hi) {
        const mid = (lo + hi) >> 1;
        range.setStart(node, mid);
        range.setEnd(node, mid + 1);
        if (range.getBoundingClientRect().top >= top - 1) hi = mid;
        else lo = mid + 1;
    }
    return lo;
}

/**
 * 在字符偏移处把元素切成两半，返回后半部分（同一 tagName/类名/内联样式）
 *
 * 只搬移节点、不重建：命中的文本节点用 splitText 切开，其后的兄弟节点整体移动。
 * 这样前半仍是原元素（保留 id / 锚点 / 事件），后半是它的浅克隆。
 */
function splitAtOffset(el: HTMLElement, offset: number): HTMLElement | null {
    const tail = el.cloneNode(false) as HTMLElement;
    let acc = 0;
    let moved = false;

    const move = (node: Node): void => {
        if (node.nodeType === Node.TEXT_NODE) {
            const len = node.nodeValue?.length ?? 0;
            if (acc + len <= offset) {
                acc += len;
                return;
            }
            const cut = offset - acc;
            if (cut > 0) {
                tail.appendChild(node.splitText(cut));
            } else {
                // 切口正好落在节点开头：整段搬走，原节点退化为空文本
                tail.appendChild(node);
            }
            acc = offset;
            moved = true;
            return;
        }
        if (node.nodeType === Node.ELEMENT_NODE) {
            const kids = Array.from(node.childNodes);
            for (const kid of kids) {
                if (moved) tail.appendChild(kid);
                else move(kid);
            }
            if (!moved) {
                // 元素节点内没有任何文字被搬走（如 <span><br></span>），整个保留
                return;
            }
            // 若原节点已被清空，就不必在 tail 里留空壳
            if (tail.lastChild === node && node.childNodes.length === 0) tail.removeChild(node);
        }
    };

    for (const child of Array.from(el.childNodes)) {
        if (moved) tail.appendChild(child);
        else move(child);
    }
    return moved ? tail : null;
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
    const lineBreak = options.lineBreak === true;
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
        slots.push({
            node,
            own: Math.max(0, bottom - top),
            gap: Math.max(0, top - prevBottom),
            isHeading: isHeadingNode(node),
            splittable: lineBreak && isSplittable(node),
            lineEnds: []
        });
        prevBottom = bottom;
    }
    // 行边界必须在节点还挂在离屏容器里时量：getClientRects 依赖真实布局，
    // 脱离文档后全为 0，一个字也测不出来。
    for (const slot of slots) {
        if (slot.splittable) slot.lineEnds = measureLineOffsets(slot.node);
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
        if (need > avail && used > 0) {
            // 放不下时先试按行拆分：能同时满足「页尾 ≥2 行」与「页首 ≥2 行」才拆，
            // 否则整块挪页，避免切出孤行（这正是 Word 的 widow/orphan 语义）。
            const remain = avail - (used + slot.gap);
            const total = slot.lineEnds.length + 1;
            const fitLines = slot.lineEnds.filter((end, idx) => (idx + 1) * (slot.own / total) <= remain).length;
            const leftover = total - fitLines;
            const canSplit =
                slot.splittable &&
                fitLines >= WIDOW_ORPHAN_LINES &&
                leftover >= WIDOW_ORPHAN_LINES &&
                slot.lineEnds.length > 0;
            if (canSplit) {
                const cut = slot.lineEnds[fitLines - 1];
                const tail = splitAtOffset(slot.node, cut);
                if (tail) {
                    // 前半留在本页：高度必须实测（行高未必均等，按比例估会累积误差）。
                    // 用 offsetHeight 而非 getBoundingClientRect：此刻节点已被移入 .oe-page-body
                    //（缩放子树内），rect 是乘过 zoom 的视觉值，而 avail 是未缩放页高，
                    // 混用会让 used 虚高 z 倍、后续块被无谓地推到下一页。
                    body!.appendChild(slot.node);
                    used += slot.gap + slot.node.offsetHeight;
                    // 后半必须插在「下一个块之前」，否则内容顺序会乱
                    slots.splice(i + 1, 0, {
                        node: tail,
                        own: tail.offsetHeight || slot.own * (leftover / total),
                        gap: 0,
                        isHeading: false,
                        splittable: slot.splittable,
                        lineEnds: slot.lineEnds.slice(fitLines).map((v) => v - cut)
                    });
                    newPage();
                    continue;
                }
            }
            newPage();
        } else if (used === 0 && slot.own > avail) slot.node.classList.add('oe-tall');
        body!.appendChild(slot.node);
        used = used ? used + slot.gap + slot.own : slot.own;
    }

    return stack;
}

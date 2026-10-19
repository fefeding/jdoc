// PDF 叠加编辑层（webview 侧）
//
// 与 pdf-viewer 的关系：阅读器负责页的渲染与导航，本模块只负责「在已渲染的页上
// 维护可编辑的叠加对象」，保存时把对象列表交给 jdoc 的 editPdf 回写。
//
// 为什么不做成 PDF 注释（annotation）：
//   - 注释在多数阅读器里不可见 / 不可编辑，用户要的是「看得见、改得了」；
//   - pdf-lib 写注释要手工构造 /Annots 对象数组与 AP 外形流，成本远高于追加内容流；
//   - 追加内容流是 pdf-lib 的强项，兼容性最好。代价是叠加物无法「撤掉原件内容」，
//     这与主流「标注后扁平化导出」的预期一致。
//
// 坐标：内部统一「PDF 点（pt）+ 页面左上原点」，与 jdoc 的 editPdf 一致；
// 只有写 DOM 样式与读鼠标位置时才除以 cssScale 换算像素。
import { editPdf, type PdfOverlay } from '@fefeding/jdoc';

/**
 * 编辑对象（UI 侧模型）
 *
 * 与库层 PdfOverlay 的差异有两处，都是为了让 8 向缩放有统一的矩形：
 * 1. 几何统一成 `width` + `boxH`（库层是 width/height，且 path 只有 x/y，尺寸由 d 决定）；
 * 2. 多运行时 id 与所属页。保存时再按类型还原成 PdfOverlay。
 *
 * 用交叉类型而非 interface extends：PdfOverlay 是判别联合，接口无法继承联合类型；
 * Omit 也要分配式地作用到每个分支，否则联合会被塌成公共字段。
 */
type DistOmit<T, K extends PropertyKey> = T extends unknown ? Omit<T, K> : never;
type EditObject = DistOmit<PdfOverlay, 'width' | 'height'> & {
    /** 运行时标识，供 DOM 与选中态使用 */
    key: string;
    /** 所属页序（0 起） */
    page: number;
    /** 命中区域宽度（pt） */
    width: number;
    /** 命中区域高度（pt） */
    boxH: number;
};

/** UI 侧编辑对象 -> 库层叠加对象 */
function toOverlay(o: EditObject): PdfOverlay {
    const { key, page, boxH, ...rest } = o;
    void key;
    void page;
    // text 的高度由库层按行数与行距推算；path 的尺寸写在 d 里；image/rect 才需要显式 height
    if (o.type === 'text') return { ...rest, width: o.width } as PdfOverlay;
    if (o.type === 'path') return { ...rest } as PdfOverlay;
    return { ...rest, height: o.boxH } as PdfOverlay;
}

type Tool = 'select' | 'text' | 'image' | 'rect' | 'draw';

export interface OverlayEditorOptions {
    /** 打开文档时的原始字节：每次保存都从它重算，避免叠加物累积 */
    original: Uint8Array;
    onSave(bytes: Uint8Array): void;
    onDirty(dirty: boolean): void;
    onToast(message: string, kind?: 'info' | 'warn' | 'error'): void;
    /**
     * 当前显示倍率（阅读器 zoom 用的是 CSS `zoom`）。鼠标坐标换算必须除掉它：
     * CSS zoom 下 getBoundingClientRect 与 clientX 都是「视觉像素」，而叠加层
     * 的定位值是「布局 px」，两者差一个倍率。缺省按 1 处理。
     */
    zoom?: () => number;
}

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const;
type Handle = (typeof HANDLES)[number];

export interface OverlayEditor {
    /** 编辑工具栏元素（由阅读器插入顶栏） */
    readonly bar: HTMLElement;
    readonly enabled: boolean;
    readonly tool: Tool;
    enable(): void;
    disable(): void;
    /**
     * 页 DOM 就绪后绘制该页叠加物
     *
     * @param pageIndex 页序，**0 起**（与库层 `overlays[]` 下标一致）
     */
    paint(pageIndex: number, pageEl: HTMLElement, cssWidth: number, pageWidthPt: number, pageHeightPt: number): void;
    /** 页几何变化后丢弃缓存，下一次 paint 重建 */
    invalidate(pageIndex: number): void;
    /** 保存：把叠加物交给库层 editPdf 回写 */
    save(): Promise<void>;
    hasObjects(): boolean;
    removeSelected(): void;
    /** 重置为打开时的原始字节（撤销编辑时） */
    reset(original: Uint8Array): void;
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string): HTMLElementTagNameMap[K] {
    const n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text !== undefined) n.textContent = text;
    return n;
}

let seq = 0;
const nextKey = (): string => `ov${++seq}`;

/** 毫米 -> pt（UI 以毫米给尺寸更直观） */
export const mm = (v: number): number => (v * 72) / 25.4;

export function createOverlayEditor(opts: OverlayEditorOptions): OverlayEditor {
    let original = opts.original;
    const objects: EditObject[] = [];
    /** pageIndex -> 叠加层与该页几何 */
    const layers = new Map<number, { layer: HTMLElement; cssWidth: number; w: number; h: number; zoom: () => number }>();
    const viewZoom: () => number = opts.zoom ?? (() => 1);
    let enabled = false;
    let tool: Tool = 'select';
    let selected: string | null = null;
    let textSize = 14;
    let textColor = '#111827';
    let strokeColor = '#e11d48';
    let fillColor = '#fde68a';

    // ---------- 工具栏 ----------
    const bar = el('div', 'pdf-edit-bar');

    const mkTool = (label: string, title: string, t: Tool): HTMLButtonElement => {
        const b = el('button', 'pdf-tool', label);
        b.type = 'button';
        b.title = title;
        b.dataset.tool = t;
        b.addEventListener('click', () => setTool(t));
        return b;
    };
    function setTool(t: Tool): void {
        tool = t;
        for (const x of Array.from(bar.querySelectorAll('.pdf-tool'))) x.classList.remove('on');
        bar.querySelector(`[data-tool="${t}"]`)?.classList.add('on');
    }

    const sizeInput = el('input', 'pdf-num');
    sizeInput.type = 'number';
    sizeInput.min = '6';
    sizeInput.max = '96';
    sizeInput.value = String(textSize);
    sizeInput.title = '字号（pt）';
    sizeInput.addEventListener('input', () => {
        textSize = Number(sizeInput.value) || 14;
    });

    const colorInput = el('input', 'pdf-color');
    colorInput.type = 'color';
    colorInput.value = textColor;
    colorInput.title = '文字颜色';
    colorInput.addEventListener('input', () => {
        textColor = colorInput.value;
    });

    const strokeInput = el('input', 'pdf-color');
    strokeInput.type = 'color';
    strokeInput.value = strokeColor;
    strokeInput.title = '描边颜色（手绘 / 矩形边框）';
    strokeInput.addEventListener('input', () => {
        strokeColor = strokeInput.value;
    });

    const delBtn = el('button', 'pdf-tool', '删除');
    delBtn.type = 'button';
    delBtn.title = '删除选中对象（Delete）';
    delBtn.addEventListener('click', () => removeSelected());

    const saveBtn = el('button', 'pdf-tool primary', '保存');
    saveBtn.type = 'button';
    saveBtn.title = '把叠加内容写入 PDF 并回存（Ctrl/Cmd+S）';
    saveBtn.addEventListener('click', () => void save());

    const fileInput = el('input');
    fileInput.type = 'file';
    fileInput.accept = 'image/png,image/jpeg';
    fileInput.style.display = 'none';

    bar.append(
        mkTool('选择', '选择 / 移动 / 缩放（V）', 'select'),
        mkTool('文本', '插入文本（T）', 'text'),
        mkTool('图片', '插入图片 / 签名（I）', 'image'),
        mkTool('矩形', '插入矩形 / 高亮（R）', 'rect'),
        mkTool('手绘', '手写签名（D）', 'draw'),
        el('span', 'pdf-sep'),
        el('label', 'pdf-field', '字号'),
        sizeInput,
        el('label', 'pdf-field', '文字'),
        colorInput,
        el('label', 'pdf-field', '描边'),
        strokeInput,
        el('span', 'spring'),
        delBtn,
        saveBtn,
        fileInput
    );
    setTool('select');

    // ---------- 绘制 ----------
    const round = (v: number): number => Math.round(v * 100) / 100;

    function toDataUrl(data: Uint8Array | ArrayBuffer): string {
        const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
        let bin = '';
        // 分块，避免大图触发调用栈上限
        for (let i = 0; i < bytes.length; i += 0x8000) {
            bin += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + 0x8000)) as number[]);
        }
        return `data:image/png;base64,${btoa(bin)}`;
    }

    function paintOne(obj: EditObject, scale: number, active: boolean): HTMLElement {
        const node = el('div', 'pdf-ov');
        node.dataset.key = obj.key;
        node.style.left = `${obj.x * scale}px`;
        node.style.top = `${obj.y * scale}px`;
        node.style.width = `${Math.max(4, obj.width * scale)}px`;
        node.style.height = `${Math.max(4, obj.boxH * scale)}px`;

        if (obj.type === 'text') {
            node.classList.add('pdf-ov-text');
            const t = el('div', 't');
            t.textContent = obj.text ?? '';
            t.style.fontSize = `${(obj.size ?? 14) * scale}px`;
            t.style.color = obj.color ?? '#111827';
            t.style.lineHeight = String(obj.lineHeight ?? 1.35);
            node.appendChild(t);
        } else if (obj.type === 'image') {
            const img = el('img');
            img.src = toDataUrl(obj.data);
            img.draggable = false;
            node.appendChild(img);
        } else if (obj.type === 'rect') {
            node.style.background = obj.fill ?? 'transparent';
            node.style.border = `${Math.max(1, (obj.strokeWidth ?? 1) * scale)}px solid ${obj.stroke ?? 'transparent'}`;
            node.style.opacity = String(obj.opacity ?? 1);
        } else if (obj.type === 'path') {
            const ns = 'http://www.w3.org/2000/svg';
            const svg = document.createElementNS(ns, 'svg');
            svg.setAttribute('width', '100%');
            svg.setAttribute('height', '100%');
            svg.setAttribute('viewBox', `0 0 ${Math.max(0.01, obj.width)} ${Math.max(0.01, obj.boxH)}`);
            const p = document.createElementNS(ns, 'path');
            p.setAttribute('d', obj.d ?? '');
            p.setAttribute('fill', 'none');
            p.setAttribute('stroke', obj.stroke ?? strokeColor);
            p.setAttribute('stroke-width', String(obj.strokeWidth ?? 2));
            p.setAttribute('stroke-linecap', 'round');
            p.setAttribute('stroke-linejoin', 'round');
            svg.appendChild(p);
            node.appendChild(svg);
        }

        if (active) {
            for (const h of HANDLES) {
                const g = el('div', `pdf-ov-h h-${h}`);
                g.dataset.handle = h;
                node.appendChild(g);
            }
            const del = el('div', 'pdf-ov-del', '×');
            del.title = '删除该对象';
            del.addEventListener('pointerdown', (e) => {
                e.stopPropagation();
                removeSelected();
            });
            node.appendChild(del);
        }
        return node;
    }

    function paint(pageIndex: number): void {
        const entry = layers.get(pageIndex);
        if (!entry) return;
        const scale = entry.cssWidth / entry.w;
        entry.layer.replaceChildren();
        for (const obj of objects) {
            if (obj.page === pageIndex) entry.layer.appendChild(paintOne(obj, scale, obj.key === selected));
        }
    }

    function paintAll(): void {
        for (const idx of layers.keys()) paint(idx);
    }

    function markDirty(): void {
        opts.onDirty(true);
    }

    function addObject(obj: EditObject): void {
        objects.push(obj);
        selected = obj.key;
        markDirty();
    }

    function removeSelected(): void {
        if (!selected) return;
        const i = objects.findIndex((o) => o.key === selected);
        if (i >= 0) objects.splice(i, 1);
        selected = null;
        markDirty();
        paintAll();
    }

    /** 按当前工具在页面落一个对象 */
    function placeAt(pageIndex: number, px: number, py: number, scale: number): void {
        const x = px / scale;
        const y = py / scale;
        if (tool === 'text') {
            const text = window.prompt('输入文字（嵌入中文需要宿主提供字体）', '在此处签名 / 备注');
            if (!text) return;
            const lines = Math.max(1, text.split('\n').length);
            addObject({
                key: nextKey(),
                type: 'text',
                page: pageIndex,
                x,
                y,
                width: mm(120),
                boxH: textSize * 1.35 * lines + textSize * 0.3,
                text,
                size: textSize,
                color: textColor,
                align: 'left'
            });
        } else if (tool === 'rect') {
            addObject({
                key: nextKey(),
                type: 'rect',
                page: pageIndex,
                x,
                y,
                width: mm(50),
                boxH: mm(8),
                fill: fillColor,
                stroke: strokeColor,
                strokeWidth: 1,
                opacity: 0.45
            });
        }
        paintAll();
    }

    // ---------- 手绘 ----------
    const drawStacks = new Map<string, number[]>();

    /** 点列 -> SVG path，坐标取相对包围盒左上 */
    function toPathD(stack: number[], minX: number, minY: number): string {
        const pts: string[] = [];
        for (let i = 0; i + 1 < stack.length; i += 2) {
            pts.push(`${round(stack[i] - minX)},${round(stack[i + 1] - minY)}`);
        }
        if (pts.length <= 1) return `M${pts[0] ?? '0,0'} L${pts[0] ?? '0,0'}`;
        return `M${pts[0]} ` + pts.slice(1).map((p) => `L${p}`).join(' ');
    }

    function drawTo(key: string, x: number, y: number): void {
        let stack = drawStacks.get(key);
        if (!stack) {
            stack = [];
            drawStacks.set(key, stack);
        }
        stack.push(x, y);
        const obj = objects.find((o) => o.key === key);
        if (!obj || obj.type !== 'path' || stack.length < 2) return;
        const xs = stack.filter((_, i) => i % 2 === 0);
        const ys = stack.filter((_, i) => i % 2 === 1);
        const minX = Math.min(...xs);
        const minY = Math.min(...ys);
        obj.x = minX;
        obj.y = minY;
        obj.width = Math.max(0.5, Math.max(...xs) - minX);
        obj.boxH = Math.max(0.5, Math.max(...ys) - minY);
        obj.d = toPathD(stack, minX, minY);
        paint(obj.page);
    }

    // ---------- 拖拽 / 缩放 ----------
    let drag: null | {
        key: string;
        mode: 'move' | 'draw' | Handle;
        startX: number;
        startY: number;
        ox: number;
        oy: number;
        ow: number;
        oh: number;
        scale: number;
    } = null;

    function onPointerDown(e: PointerEvent): void {
        if (!enabled) return;
        const target = e.target as HTMLElement;
        if (!(target instanceof Element)) return;

        if (tool === 'image') {
            fileInput.click();
            return;
        }
        const layerEl = target.closest<HTMLElement>('.pdf-ov-layer');
        const node = target.closest<HTMLElement>('.pdf-ov');
        const handleEl = target.closest<HTMLElement>('.pdf-ov-h');

        if (layerEl) {
            const entry = layers.get(Number(layerEl.dataset.page));
            if (!entry) return;
            const rect = layerEl.getBoundingClientRect();
            const z = viewZoom() || 1;
            // rect 与 clientX 都是视觉像素，先除掉 CSS zoom 回到布局 px，再按 pt 换算
            const scale = entry.cssWidth / entry.w;
            const px = (e.clientX - rect.left) / z;
            const py = (e.clientY - rect.top) / z;

            if (tool === 'draw') {
                const obj: EditObject = {
                    key: nextKey(),
                    type: 'path',
                    page: Number(layerEl.dataset.page),
                    x: px / scale,
                    y: py / scale,
                    width: 0.5,
                    boxH: 0.5,
                    d: 'M0 0',
                    stroke: strokeColor,
                    strokeWidth: 2
                };
                objects.push(obj);
                selected = obj.key;
                markDirty();
                drawStacks.set(obj.key, []);
                drag = { key: obj.key, mode: 'draw', startX: e.clientX, startY: e.clientY, ox: obj.x, oy: obj.y, ow: obj.width, oh: obj.boxH, scale: scale * z };
                e.preventDefault();
                return;
            }
            if (!node) {
                if (tool === 'select') {
                    selected = null;
                    paintAll();
                } else {
                    placeAt(Number(layerEl.dataset.page), px, py, scale);
                }
                return;
            }
        }

        if (!node) return;
        const key = node.dataset.key;
        const obj = objects.find((o) => o.key === key);
        if (!key || !obj) return;
        if (selected !== key) {
            selected = key;
            paintAll();
        }
        if (tool !== 'select' && obj.type !== 'path') return;

        const entry = layers.get(obj.page);
        if (!entry) return;
        drag = {
            key,
            mode: (handleEl?.dataset.handle as Handle) ?? 'move',
            startX: e.clientX,
            startY: e.clientY,
            ox: obj.x,
            oy: obj.y,
            ow: obj.width,
            oh: obj.boxH,
            scale: (entry.cssWidth / entry.w) * (viewZoom() || 1)
        };
        e.preventDefault();
    }

    function onPointerMove(e: PointerEvent): void {
        if (!drag) return;
        const obj = objects.find((o) => o.key === drag!.key);
        if (!obj) return;
        const dx = (e.clientX - drag.startX) / drag.scale;
        const dy = (e.clientY - drag.startY) / drag.scale;

        if (drag.mode === 'draw') {
            drawTo(obj.key, obj.x + dx, obj.y + dy);
            return;
        }
        if (drag.mode === 'move') {
            obj.x = drag.ox + dx;
            obj.y = drag.oy + dy;
        } else {
            applyResize(obj, drag.mode, drag.ox, drag.oy, drag.ow, drag.oh, dx, dy);
        }
        paint(obj.page);
    }

    function applyResize(
        obj: EditObject,
        mode: Handle,
        ox: number,
        oy: number,
        ow: number,
        oh: number,
        dx: number,
        dy: number
    ): void {
        let x = ox;
        let y = oy;
        let width = ow;
        let boxH = oh;
        if (mode.includes('e')) width = Math.max(2, ow + dx);
        if (mode.includes('s')) boxH = Math.max(2, oh + dy);
        if (mode.includes('w')) {
            width = Math.max(2, ow - dx);
            x = ox + (ow - width);
        }
        if (mode.includes('n')) {
            boxH = Math.max(2, oh - dy);
            y = oy + (oh - boxH);
        }
        obj.x = x;
        obj.y = y;
        obj.width = width;
        obj.boxH = boxH;
    }

    function onPointerUp(): void {
        if (!drag) return;
        const obj = objects.find((o) => o.key === drag!.key);
        if (drag.mode === 'draw' && obj && obj.type === 'path') {
            drawStacks.delete(obj.key);
            // 单击（没有轨迹）视为误触，丢弃
            if ((obj.d ?? '') === 'M0 0') {
                const i = objects.indexOf(obj);
                if (i >= 0) objects.splice(i, 1);
                selected = null;
            }
        }
        drag = null;
        markDirty();
        paintAll();
    }

    // ---------- 图片插入 ----------
    fileInput.addEventListener('change', () => {
        const file = fileInput.files?.[0];
        fileInput.value = '';
        const first = [...layers.entries()][0];
        if (!file || !first) return;
        const [pageIndex] = first;
        void file.arrayBuffer().then((buf) => {
            const bytes = new Uint8Array(buf);
            const probe = new Image();
            probe.onload = () => {
                const w = mm(40);
                addObject({
                    key: nextKey(),
                    type: 'image',
                    page: pageIndex,
                    x: mm(20),
                    y: mm(20),
                    width: w,
                    boxH: (w * probe.naturalHeight) / Math.max(1, probe.naturalWidth),
                    data: bytes,
                    format: file.type === 'image/jpeg' ? 'jpg' : 'png'
                });
                paintAll();
            };
            probe.onerror = () => opts.onToast('图片解码失败', 'error');
            probe.src = toDataUrl(bytes);
        });
    });

    // ---------- 保存 ----------
    async function save(): Promise<void> {
        if (!objects.length) {
            opts.onToast('没有可保存的叠加内容', 'warn');
            return;
        }
        try {
            const maxPage = objects.reduce((m, o) => Math.max(m, o.page), 0);
            const byPage: EditObject[][] = [];
            for (const o of objects) (byPage[o.page] ??= []).push(o);
            const overlays: PdfOverlay[][] = [];
            for (let i = 0; i <= maxPage; i++) {
                overlays[i] = (byPage[i] ?? []).map(toOverlay);
            }
            // 始终从打开时的原始字节重算，叠加物不会被重复写入
            const bytes = await editPdf(original, { overlays });
            opts.onSave(bytes);
        } catch (err: unknown) {
            opts.onToast(`保存失败：${(err as { message?: string })?.message || String(err)}`, 'error');
        }
    }

    // ---------- 生命周期 ----------
    function onKey(e: KeyboardEvent): void {
        if (!enabled) return;
        const t = e.target as HTMLElement | null;
        if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable)) return;
        if ((e.key === 'Delete' || e.key === 'Backspace') && selected) {
            e.preventDefault();
            removeSelected();
        } else if (e.key === 'Escape') {
            selected = null;
            paintAll();
        } else if (e.key === 'v' || e.key === 'V') setTool('select');
        else if (e.key === 't' || e.key === 'T') setTool('text');
        else if (e.key === 'i' || e.key === 'I') setTool('image');
        else if (e.key === 'r' || e.key === 'R') setTool('rect');
        else if (e.key === 'd' || e.key === 'D') setTool('draw');
        else if (e.key === 's' && (e.metaKey || e.ctrlKey)) {
            e.preventDefault();
            void save();
        }
    }

    function attach(): void {
        document.addEventListener('pointerdown', onPointerDown, true);
        document.addEventListener('pointermove', onPointerMove);
        document.addEventListener('pointerup', onPointerUp);
        document.addEventListener('keydown', onKey);
    }
    function detach(): void {
        document.removeEventListener('pointerdown', onPointerDown, true);
        document.removeEventListener('pointermove', onPointerMove);
        document.removeEventListener('pointerup', onPointerUp);
        document.removeEventListener('keydown', onKey);
    }

    return {
        bar,
        get enabled() {
            return enabled;
        },
        get tool() {
            return tool;
        },
        enable(): void {
            if (enabled) return;
            enabled = true;
            for (const v of layers.values()) v.layer.classList.add('on');
            attach();
        },
        disable(): void {
            if (!enabled) return;
            enabled = false;
            detach();
            selected = null;
            for (const v of layers.values()) v.layer.classList.remove('on');
        },
        paint(pageIndex, pageEl, cssWidth, pageWidthPt, pageHeightPt): void {
            let layer = pageEl.querySelector<HTMLElement>('.pdf-ov-layer');
            if (!layer) {
                layer = el('div', 'pdf-ov-layer');
                layer.dataset.page = String(pageIndex);
                pageEl.appendChild(layer);
            }
            layer.style.width = `${cssWidth}px`;
            layer.style.height = `${(pageHeightPt * cssWidth) / pageWidthPt}px`;
            layer.classList.toggle('on', enabled);
            layers.set(pageIndex, { layer, cssWidth, w: pageWidthPt, h: pageHeightPt, zoom: viewZoom });
            paint(pageIndex);
        },
        invalidate(pageIndex: number): void {
            layers.delete(pageIndex);
        },
        save,
        hasObjects(): boolean {
            return objects.length > 0;
        },
        removeSelected,
        reset(bytes: Uint8Array): void {
            original = bytes;
            objects.length = 0;
            selected = null;
            // 页 DOM 会被整体重建，指向旧 layer 的引用必须一并丢弃，
            // 否则后续 paint 会把叠加物画到已脱离文档的节点上
            layers.clear();
            drawStacks.clear();
            opts.onDirty(false);
        }
    };
}

// office-editor 公共工具：字节/base64、DOM、防抖、列名。

/** base64 → Uint8Array */
export function b64ToBytes(b64: string): Uint8Array {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
}

/** Uint8Array → base64（分块避免大数组展开爆栈） */
export function bytesToB64(bytes: Uint8Array): string {
    let bin = '';
    const CHUNK = 0x8000;
    for (let i = 0; i < bytes.length; i += CHUNK) {
        bin += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + CHUNK)) as unknown as number[]);
    }
    return btoa(bin);
}

export function escapeHtml(s: string): string {
    return s.replace(/[&<>"']/g, (c) => {
        switch (c) {
            case '&': return '&amp;';
            case '<': return '&lt;';
            case '>': return '&gt;';
            case '"': return '&quot;';
            default: return '&#39;';
        }
    });
}

type Attrs = Record<string, unknown>;

/** 轻量 DOM 构造器：h('div', {class, onclick, dataset, html, text, style}, [children]) */
export function h<K extends keyof HTMLElementTagNameMap>(
    tag: K,
    attrs: Attrs = {},
    children: (Node | string | null | undefined)[] = []
): HTMLElementTagNameMap[K] {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
        if (v == null) continue;
        if (k === 'class' || k === 'className') el.className = String(v);
        else if (k === 'text') el.textContent = String(v);
        else if (k === 'html') el.innerHTML = String(v);
        else if (k === 'style') el.setAttribute('style', String(v));
        else if (k === 'dataset') Object.assign(el.dataset, v as Record<string, string>);
        else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v as EventListener);
        else el.setAttribute(k, String(v));
    }
    for (const c of children) {
        if (c == null) continue;
        el.append(typeof c === 'string' ? document.createTextNode(c) : c);
    }
    return el;
}

export function debounce<A extends unknown[]>(fn: (...a: A) => void, ms: number): (...a: A) => void {
    let t: number | undefined;
    return (...a: A) => {
        if (t) window.clearTimeout(t);
        t = window.setTimeout(() => fn(...a), ms);
    };
}

export function clamp(v: number, lo: number, hi: number): number {
    return Math.max(lo, Math.min(hi, v));
}

/** 0 -> 'A', 25 -> 'Z', 26 -> 'AA'（XLSX 列名，1 起的列号转字母） */
export function colName(n: number): string {
    let s = '';
    let i = n;
    while (i > 0) {
        const r = (i - 1) % 26;
        s = String.fromCharCode(65 + r) + s;
        i = Math.floor((i - 1) / 26);
    }
    return s;
}

import { pptxToHtml, docxToHtml, xlsxToHtml, markdownToHtml, textToHtml } from '@fefeding/jdoc';
import { chartRenderer } from './chart-renderer';
import { b64ToBytes, escapeHtml } from './util';
import { baseCss } from './style';
import type { DocFormat, EditorMode, HostToWebview } from '../protocol';

declare function acquireVsCodeApi(): { postMessage(msg: unknown): void };
const vscode = acquireVsCodeApi();

let format: DocFormat = 'pptx';
let mode: EditorMode = 'view';
let currentBytes: Uint8Array = new Uint8Array(0);

function applyBaseStyle(): void {
    const style = document.createElement('style');
    style.textContent = baseCss;
    document.head.appendChild(style);
}

function injectGlobalStyles(css: string): void {
    let el = document.getElementById('pptx-styles') as HTMLStyleElement | null;
    if (!el) {
        el = document.createElement('style');
        el.id = 'pptx-styles';
        document.head.appendChild(el);
    }
    el.textContent = css;
}

function buildToolbar(toolbar: HTMLElement): void {
    toolbar.innerHTML = '';
    const viewBtn = document.createElement('button');
    viewBtn.className = 'tbtn' + (mode === 'view' ? ' active' : '');
    viewBtn.textContent = '查看';
    viewBtn.onclick = () => setMode('view');
    toolbar.appendChild(viewBtn);

    const editBtn = document.createElement('button');
    editBtn.className = 'tbtn' + (mode === 'edit' ? ' active' : '');
    editBtn.textContent = '编辑';
    editBtn.onclick = () => setMode('edit');
    toolbar.appendChild(editBtn);
}

function renderShell(): { toolbar: HTMLElement; content: HTMLElement } {
    const app = document.getElementById('app')!;
    app.innerHTML = '';
    const toolbar = document.createElement('div');
    toolbar.className = 'toolbar';
    app.appendChild(toolbar);
    const content = document.createElement('div');
    content.id = 'content';
    app.appendChild(content);
    return { toolbar, content };
}

function setMode(next: EditorMode): void {
    mode = next;
    vscode.postMessage({ type: 'setMode', mode });
    const { toolbar } = renderShell();
    buildToolbar(toolbar);
    void renderContent();
}

async function renderContent(): Promise<void> {
    const content = document.getElementById('content')!;
    if (mode === 'edit') {
        // PPTX / DOCX / XLSX 已各自使用独立编辑器入口（pptx-editor.js / office-editor.js）；
        // 能走到这里的仅 PDF / MD / TXT，按只读处理。切走时作废进行中的 PDF 渲染任务。
        pdfToken += 1;
        if (pdfDoc) {
            void pdfDoc.destroy?.();
            pdfDoc = null;
        }
        const name = format === 'md' ? 'Markdown' : format === 'txt' ? '纯文本' : 'PDF';
        content.innerHTML = `<div class="hint">${name}为只读查看，暂不支持编辑。</div>`;
        return;
    }
    try {
        if (format === 'pptx') await renderPptx(content, currentBytes);
        else if (format === 'docx') await renderDocx(content, currentBytes);
        else if (format === 'xlsx') await renderXlsx(content, currentBytes);
        else if (format === 'pdf') await renderPdf(content, currentBytes);
        else if (format === 'md') await renderMarkdown(content, currentBytes);
        else if (format === 'txt') await renderText(content, currentBytes);
    } catch (err: any) {
        content.innerHTML = `<pre class="error">渲染失败：${escapeHtml(String(err?.message || err))}</pre>`;
        vscode.postMessage({ type: 'error', message: String(err?.message || err) });
    }
}

async function renderPptx(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    const r = await pptxToHtml(bytes, { slideMode: false, keyBoardShortCut: false });
    injectGlobalStyles(r.styles?.global || '');
    const wrap = document.createElement('div');
    wrap.className = 'pptx-slides';
    for (const slide of r.slides || []) {
        const holder = document.createElement('div');
        holder.innerHTML = slide.html;
        const scaler = document.createElement('div');
        scaler.className = 'slide-scaler';
        while (holder.firstChild) scaler.appendChild(holder.firstChild);
        wrap.appendChild(scaler);
    }
    content.appendChild(wrap);
    // 图表：ECharts 已由 media 脚本注入为全局变量，chart-renderer 直接使用
    if (r.charts && r.charts.length) {
        chartRenderer.renderCharts(r.charts, wrap);
    }
}

async function renderDocx(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    const { html } = await docxToHtml(bytes);
    const body = document.createElement('div');
    body.className = 'docx-body';
    body.innerHTML = html;
    content.appendChild(body);
}

async function renderXlsx(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    const { html } = await xlsxToHtml(bytes);
    const body = document.createElement('div');
    body.className = 'xlsx-body';
    body.innerHTML = html;
    content.appendChild(body);
}

async function renderMarkdown(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    const { html, messages } = await markdownToHtml(bytes);
    const body = document.createElement('div');
    body.className = 'md-body';
    body.innerHTML = html;
    content.appendChild(body);
    // 外部环境链接在新标签页打开（webview 内无法跳转）
    body.querySelectorAll<HTMLAnchorElement>('a[href^="http"]').forEach((a) => {
        a.setAttribute('target', '_blank');
        a.setAttribute('rel', 'noopener noreferrer');
    });
    if (messages?.length) {
        const note = document.createElement('div');
        note.className = 'hint';
        note.textContent = messages.join('；');
        content.appendChild(note);
    }
}

async function renderText(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    const { html } = await textToHtml(bytes);
    const body = document.createElement('div');
    body.className = 'txt-body';
    body.innerHTML = html;
    content.appendChild(body);
}

/** 当前 PDF 渲染代数：渲染过程中切模式/重渲染时用于丢弃过期任务 */
let pdfToken = 0;
/** 当前打开的 pdfjs 文档，重新渲染前销毁以释放 worker 侧资源 */
let pdfDoc: any = null;
/** 已渲染页数计数回调 */
let pdfProgress: ((done: number, total: number) => void) | null = null;

/** 加载 pdfjs 并设置有效 worker（extension 注入 window.__JDC_PDF_WORKER__） */
async function loadPdfjs(): Promise<any> {
    // 与 jdoc 内部使用同一说明符，命中同一模块实例
    const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
    const worker = (window as any).__JDC_PDF_WORKER__;
    if (worker && pdfjs.GlobalWorkerOptions && !pdfjs.GlobalWorkerOptions.workerSrc) {
        pdfjs.GlobalWorkerOptions.workerSrc = worker;
    }
    return pdfjs;
}

/** 单页栅格化到 canvas，并在其上叠一层透明文本层（供选中/复制） */
async function renderPdfPage(page: any, host: HTMLElement, cssWidth: number): Promise<void> {
    const base = page.getViewport({ scale: 1 });
    const cssScale = cssWidth / base.width;
    const dpr = Math.min(2, window.devicePixelRatio || 1);
    const viewport = page.getViewport({ scale: cssScale * dpr });

    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.floor(viewport.width));
    canvas.height = Math.max(1, Math.floor(viewport.height));
    canvas.style.width = `${Math.round(base.width * cssScale)}px`;
    canvas.style.height = `${Math.round(base.height * cssScale)}px`;

    host.classList.remove('pending');
    host.style.width = `${Math.round(base.width * cssScale)}px`;
    host.style.height = `${Math.round(base.height * cssScale)}px`;
    host.replaceChildren(canvas);

    const ctx = canvas.getContext('2d');
    if (!ctx) throw new Error('无法创建 canvas 上下文');
    await page.render({ canvasContext: ctx, viewport }).promise;

    // 透明文本层：视觉不可见，但可选中/复制（PDF 原点在左下，需换算到 CSS 坐标）
    try {
        const tc = await page.getTextContent();
        const items = ((tc.items || []) as any[]).filter((it) => typeof it?.str === 'string' && it.str);
        if (!items.length) return;
        const layer = document.createElement('div');
        layer.className = 'pdf-text-layer';
        for (const it of items) {
            const t = (it.transform || []) as number[];
            const span = document.createElement('span');
            span.textContent = it.str;
            span.style.left = `${(t[4] || 0) * cssScale}px`;
            const fs = Math.abs(t[3] || it.height || 0) || 10;
            span.style.top = `${(base.height - (t[5] || 0) - fs) * cssScale}px`;
            span.style.fontSize = `${fs * cssScale}px`;
            layer.appendChild(span);
        }
        host.appendChild(layer);
    } catch {
        // 文本层失败不影响图形渲染
    }
}

async function renderPdf(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    const token = ++pdfToken;
    const stale = () => token !== pdfToken;

    const hint = document.createElement('div');
    hint.className = 'hint';
    hint.textContent = '正在打开 PDF…';
    content.replaceChildren(hint);

    const pdfjs = await loadPdfjs();
    // pdfjs 会把 data 转移给 worker（原 ArrayBuffer 会被 detach），必须传副本
    const doc = await pdfjs.getDocument({ data: bytes.slice() }).promise;
    if (stale()) {
        void doc.destroy?.();
        return;
    }
    if (pdfDoc && pdfDoc !== doc) void pdfDoc.destroy?.();
    pdfDoc = doc;

    const total = doc.numPages || 0;
    const wrap = document.createElement('div');
    wrap.className = 'pdf-pages';
    let done = 0;
    hint.textContent = `共 ${total} 页，按需渲染中…（0/${total}）`;
    pdfProgress = (d, t) => {
        if (stale()) return;
        hint.textContent = `共 ${t} 页，按需渲染中…（${d}/${t}）`;
    };
    content.replaceChildren(hint, wrap);

    // 页面宽度按容器可用宽度自适应（上限 1000px）
    const cssWidth = Math.max(320, Math.min(1000, (content.clientWidth || 900) - 40));

    const renderOne = async (index: number) => {
        if (stale()) return;
        const host = wrap.querySelector<HTMLElement>(`.pdf-page[data-page="${index}"]`);
        if (!host || !host.classList.contains('pending')) return;
        try {
            const page = await doc.getPage(index);
            if (stale()) return;
            await renderPdfPage(page, host, cssWidth);
        } catch (err: any) {
            if (stale()) return;
            host.classList.remove('pending');
            host.textContent = `第 ${index} 页渲染失败：${String(err?.message || err)}`;
        } finally {
            if (!stale()) {
                done += 1;
                pdfProgress?.(done, total);
            }
        }
    };

    // 先按首页比例撑出占位（避免滚动时布局跳动），进入视口前后再真正渲染
    const first = await doc.getPage(1);
    if (stale()) return;
    const v1 = first.getViewport({ scale: 1 });
    const phHeight = Math.round((cssWidth * v1.height) / v1.width);
    for (let i = 1; i <= total; i++) {
        const p = document.createElement('div');
        p.className = 'pdf-page pdf-canvas-page pending';
        p.dataset.page = String(i);
        p.style.width = `${cssWidth}px`;
        p.style.height = `${phHeight}px`;
        p.textContent = `第 ${i} 页`;
        wrap.appendChild(p);
    }

    const io = new IntersectionObserver(
        (entries) => {
            for (const e of entries) {
                if (!e.isIntersecting) continue;
                const el = e.target as HTMLElement;
                io.unobserve(el);
                void renderOne(Number(el.dataset.page));
            }
        },
        { root: null, rootMargin: '800px 0px' }
    );
    wrap.querySelectorAll('.pdf-page').forEach((el) => io.observe(el));

    // 首屏立即渲染，避免等 observer 回调
    await renderOne(1);
}

window.addEventListener('message', (e) => {
    const msg = e.data as HostToWebview;
    switch (msg.type) {
        case 'init':
            format = msg.format;
            mode = msg.mode || 'view';
            currentBytes = b64ToBytes(msg.bytes);
            applyBaseStyle();
            {
                const { toolbar } = renderShell();
                buildToolbar(toolbar);
            }
            void renderContent();
            vscode.postMessage({ type: 'ready' });
            break;
        case 'reverted':
            void renderContent();
            break;
        case 'saved':
        case 'info':
            break;
    }
});

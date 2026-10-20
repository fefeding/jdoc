import { pptxToHtml, docxToHtml, xlsxToHtml, markdownToHtml, textToHtml, renderMarkdownExtras } from '@fefeding/jdoc';
import { chartRenderer } from './chart-renderer';
import { b64ToBytes, bytesToB64, escapeHtml } from './util';
import { baseCss } from './style';
import { CHROME_CSS, createStatusBar } from './chrome';
import { createZoom, createZoomCluster } from './zoom';
import { createConvertChip, convertMenuEntry, copySelection, hasTextSelection, installContextMenu, setCurrentDocUri } from './chrome';
import { createPdfViewer, PDF_CSS, type PdfViewer } from './pdf-viewer';
import { highlightCodeBlocks, markdownCodeCss } from './markdown-code';
import type { DocFormat, EditorMode, HostToWebview } from '../protocol';

declare function acquireVsCodeApi(): { postMessage(msg: unknown): void };
const vscode = acquireVsCodeApi();

let format: DocFormat = 'pptx';
let mode: EditorMode = 'view';
let currentBytes: Uint8Array = new Uint8Array(0);

/**
 * 取文档字节：优先 fetch 宿主给的可访问文件 URI（大文件避免 base64 经 postMessage
 * 序列化卡死主线程），失败再回退到 init.bytes 的 base64。
 */
async function loadBytes(msg: HostToWebview): Promise<Uint8Array> {
    if (msg.type === 'init' && msg.fileUri) {
        try {
            const res = await fetch(msg.fileUri);
            if (res.ok) return new Uint8Array(await res.arrayBuffer());
        } catch {
            // 回退 base64
        }
    }
    const init = msg.type === 'init' ? msg : (msg as { bytes?: string });
    return b64ToBytes(init.bytes || '');
}
let statusText: HTMLElement | null = null;

/** 缩放：PDF / Markdown / 纯文本等只读内容共用，默认适应宽度，宽内容不会出屏外 */
const zoom = createZoom();

/** 把缩放绑定到当前渲染出的内容根 */
function bindZoom(): void {
    const content = document.getElementById('content');
    if (!content) return;
    const el = content.firstElementChild as HTMLElement | null;
    // 提示 / 错误块无需缩放
    if (!el || el.classList.contains('hint') || el.classList.contains('error')) {
        zoom.bind({ target: null, container: content });
        return;
    }
    zoom.bind({ target: el, container: content });
}

function applyBaseStyle(): void {
    const style = document.createElement('style');
    style.textContent = CHROME_CSS + baseCss + PDF_CSS + markdownCodeCss;
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

/** 更新底部状态栏文案 */
function setStatus(text: string): void {
    if (statusText) statusText.textContent = text;
}

/** PDF 阅读器实例（仅 pdf 格式创建） */
let pdfViewer: PdfViewer | null = null;

function renderShell(): { content: HTMLElement } {
    const app = document.getElementById('app')!;
    // 重建外壳前先拆旧实例：PDF 编辑层在 document 上挂了 pointer 监听，
    // 不 close 会随每次 init 累积一份
    pdfViewer?.close();
    app.innerHTML = '';

    // PDF 有自己的阅读器外壳（顶部工具栏 + 左侧缩略图栏 + 正文），缩放也由它绑定
    if (format === 'pdf') {
        app.className = 'jc-app pdf-app';
        const status = createStatusBar();
        statusText = status.textEl;
        pdfViewer = createPdfViewer(app, zoom, setStatus, (editing) => {
            // 阅读器里手动开关编辑模式，同步给宿主，标题栏的「查看/编辑」命令才不会失配
            vscode.postMessage({ type: 'setMode', mode: editing ? 'edit' : 'view' });
        });
        // 叠加编辑产出新字节：先回传字节（标脏），再请求落盘，
        // 一步完成「生成 + 写回原文件」，不必让用户再按一次保存。
        pdfViewer.onSave((bytes) => {
            vscode.postMessage({ type: 'sync', bytes: bytesToB64(bytes), dirty: true });
            vscode.postMessage({ type: 'save' });
        });
        // PDF 阅读器自带工具栏且已挂缩放簇（pdf-viewer 内），这里只补转换入口，
        // 避免同一份缩放簇出现两次
        status.el.append(createConvertChip());
        app.append(status.el);
        return { content: pdfViewer.content };
    }

    // 不画顶栏：标题 / 面包屑 / 未保存点由 VS Code 编辑器标题栏呈现，webview 再画一份只会重复。
    const content = document.createElement('div');
    content.className = 'jc-content';
    content.id = 'content';
    const status = createStatusBar();
    statusText = status.textEl;
    app.className = 'jc-app';
    app.append(content, status.el);
    // 缩放簇挂在状态栏尾部（createStatusBar 里已有一个 grow 占位，簇会自然靠右）
    status.el.append(createConvertChip(), createZoomCluster(zoom));
    return { content };
}

async function renderContent(): Promise<void> {
    const content = document.getElementById('content')!;
    // PDF 无论查看还是编辑都走阅读器外壳，编辑模式只是叠加一层 overlay 工具层
    if (format === 'pdf' && pdfViewer) {
        try {
            await pdfViewer.open(currentBytes);
            pdfViewer.setEditing(mode === 'edit', false);
        } catch (err: any) {
            console.error('[pdf] renderContent open failed:', err);
            pdfViewer.content.innerHTML = `<pre class="error">渲染失败：${escapeHtml(String(err?.message || err))}</pre>`;
            setStatus('渲染失败');
            vscode.postMessage({ type: 'error', message: String(err?.message || err) });
        }
        return;
    }
    if (mode === 'edit') {
        // PPTX / DOCX / XLSX 已各自使用独立编辑器入口（pptx-editor.js / office-editor.js）；
        // 能走到这里的仅 MD / TXT，按只读处理。
        const name = format === 'md' ? 'Markdown' : '纯文本';
        const hint = document.createElement('div');
        hint.className = 'hint';
        hint.textContent = `${name}为只读查看，暂不支持编辑。`;
        content.replaceChildren(hint);
        setStatus('只读预览');
        bindZoom();
        return;
    }
    try {
        if (format === 'pdf') {
            // 上面已处理；走到这里说明外壳没建成（理论上不会发生）
            return;
        }
        if (format === 'pptx') await renderPptx(content, currentBytes);
        else if (format === 'docx') await renderDocx(content, currentBytes);
        else if (format === 'xlsx') await renderXlsx(content, currentBytes);
        else if (format === 'md') await renderMarkdown(content, currentBytes);
        else if (format === 'txt') await renderText(content, currentBytes);
        bindZoom();
    } catch (err: any) {
        content.innerHTML = `<pre class="error">渲染失败：${escapeHtml(String(err?.message || err))}</pre>`;
        setStatus('渲染失败');
        bindZoom();
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
    const slideCount = (r.slides || []).length;
    setStatus(`共 ${slideCount} 页幻灯片${r.charts?.length ? `，${r.charts.length} 个图表` : ''}`);
}

async function renderDocx(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    const { html } = await docxToHtml(bytes);
    const body = document.createElement('div');
    body.className = 'docx-body';
    body.innerHTML = html;
    content.appendChild(body);

    setStatus('DOCX 文档预览');
}

async function renderXlsx(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    const { html } = await xlsxToHtml(bytes);
    const body = document.createElement('div');
    body.className = 'xlsx-body';
    body.innerHTML = html;
    content.appendChild(body);

    setStatus('XLSX 表格预览');
}

async function renderMarkdown(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    const { html, messages: baseMessages } = await markdownToHtml(bytes);
    let messages = baseMessages || [];
    const body = document.createElement('div');
    body.className = 'md-body';
    body.innerHTML = html;
    content.appendChild(body);
    // 外部环境链接在新标签页打开（webview 内无法跳转）
    body.querySelectorAll<HTMLAnchorElement>('a[href^="http"]').forEach((a) => {
        a.setAttribute('target', '_blank');
        a.setAttribute('rel', 'noopener noreferrer');
    });
    // 围栏代码块上色：必须在 mermaid 渲染之前，且 mermaid 自身跳过
    highlightCodeBlocks(body);
    // 增强内容：Mermaid 图表 / KaTeX 公式（动态加载依赖，失败时保留源码降级显示）
    try {
        // 主题判定：优先用 VS Code 给 webview body 加的类（实时跟随颜色主题），
        // 其次退回 html 上的 data-oe-theme，再退回系统偏好。
        const cls = document.body.classList;
        const dark = cls.contains('vscode-dark') || cls.contains('vscode-high-contrast') ||
            (document.documentElement.getAttribute('data-oe-theme') === 'dark');
        await renderMarkdownExtras(body, { theme: dark ? 'dark' : 'default' });
    } catch (err: any) {
        messages.push(`图表/公式渲染失败：${String(err?.message || err)}`);
    }
    setStatus(messages.length ? 'Markdown 预览（有清洗提示）' : 'Markdown 预览');
    if (messages.length) {
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

    setStatus('纯文本预览');
}

/**
 * 把整份 PDF 逐页渲染成 PNG，每完成一页立刻回传一页。
 *
 * 宿主端没有 canvas 可用（pdfjs 的 Node 渲染依赖 @napi-rs/canvas，约 27MB 的平台相关
 * 原生二进制，无法随扩展分发），所以由 webview 用浏览器原生 canvas 渲染；中文 CJK
 * 字形在浏览器里走 Path2D 也不会有宿主侧遇到的 napi InvalidArg 问题。
 *
 * 逐页回传而不是攒齐再发：几十页文档的 PNG 全塞进一条 postMessage 体会非常大，
 * 宿主边收边落盘，内存峰值只剩单页。
 */
async function rasterize(bytes: Uint8Array, scale: number): Promise<void> {
    try {
        const pdfjs: any = await import('pdfjs-dist/legacy/build/pdf.mjs');
        const worker = (window as any).__JDC_PDF_WORKER__;
        if (worker && pdfjs.GlobalWorkerOptions && !pdfjs.GlobalWorkerOptions.workerSrc) {
            pdfjs.GlobalWorkerOptions.workerSrc = worker;
        }
        // pdfjs 会把 data 转移给 worker，必须传副本，否则原始字节被 detach
        const doc = await pdfjs.getDocument({ data: bytes.slice(), isEvalSupported: false }).promise;
        for (let i = 1; i <= doc.numPages; i++) {
            const page = await doc.getPage(i);
            const viewport = page.getViewport({ scale });
            const canvas = document.createElement('canvas');
            canvas.width = Math.ceil(viewport.width);
            canvas.height = Math.ceil(viewport.height);
            const ctx = canvas.getContext('2d');
            if (!ctx) throw new Error('无法创建 canvas 2D 上下文');
            // 铺白底：PDF 页面本身透明，直接导出会得到透明 PNG，在部分看图器里显示为黑块
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, canvas.width, canvas.height);
            await page.render({ canvasContext: ctx, viewport }).promise;
            vscode.postMessage({ type: 'rasterizePage', index: i, dataUrl: canvas.toDataURL('image/png') });
        }
        vscode.postMessage({ type: 'rasterizeDone', count: doc.numPages });
    } catch (err: any) {
        vscode.postMessage({ type: 'rasterizeError', message: String(err?.message || err) });
    }
}

window.addEventListener('message', (e) => {
    const msg = e.data as HostToWebview;
    switch (msg.type) {
        case 'init':
            void (async () => {
                format = msg.format;
                mode = msg.mode || 'view';
                // 记住当前文档 uri：转换请求必须带上，否则宿主只能靠 activeTextEditor 猜
                setCurrentDocUri(msg.uri ?? null);
                currentBytes = await loadBytes(msg);
                applyBaseStyle();
                renderShell();
                pdfViewer?.setTitle(msg.title || '');
                zoom.fit(); // 新文档：缩放复位到适应宽度
                void renderContent();
            })();
            break;
        case 'rasterize':
            void rasterize(b64ToBytes(msg.bytes), msg.scale || 2);
            break;
        case 'reverted':
            void renderContent();
            break;
        case 'setMode':
            // 标题栏的「查看 / 编辑」命令：PDF 切到编辑模式就是打开 overlay 工具层
            mode = msg.mode;
            void renderContent();
            break;
        case 'saved':
            pdfViewer?.markSaved();
            break;
        case 'info':
            break;
    }
});

// 脚本就绪即握手，而不是等处理完 init 再回：宿主必须先确认 webview 已注册监听，
// 后续消息（init / rasterize）才不会被丢弃。全程只发一次，宿主收到 ready 后的回发
// 不会再次触发它，因此不存在握手循环。
vscode.postMessage({ type: 'ready' });

// 文档内容区右键菜单：custom editor 的正文在 webview 里，VS Code 的 editor/context
// 菜单不会出现，这里自绘一份，把「转换格式」补进内容区右键（PDF / Markdown / 纯文本）。
installContextMenu(() => {
    const entries = [];
    if (hasTextSelection()) {
        entries.push({ label: '复制', action: () => copySelection(), hint: '⌘C' });
        entries.push({});
    }
    entries.push(convertMenuEntry());
    return entries;
});

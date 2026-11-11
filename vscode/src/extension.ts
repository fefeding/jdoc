import * as vscode from 'vscode';
import * as path from 'path';
import { pathToFileURL } from 'url';
import { PDFDocument } from 'pdf-lib';
import { JdocDocument } from './document';
import { convert, type ConvertTarget } from '@fefeding/jdoc';
import { extToFormat, targetFormats, type SourceFormat } from './convert-formats';
import type { DocFormat, DocTabInfo, EditorMode, HostToWebview, WebviewToHost } from './protocol';

/**
 * 宿主侧解析 PDF 用的 pdfjs worker 地址。
 *
 * pdfjs 被 esbuild 内联进 extension.js 后，它内置的相对 worker 路径（"./pdf.worker.mjs"）
 * 会指向 dist/ 目录，worker 必然加载不到，PDF 转换会直接失败。esbuild 会把真实的
 * worker 复制到 dist/pdf.worker.mjs（见 esbuild.js 的 copyPdfWorker），这里把它
 * 的绝对地址显式传给 jdoc。
 */
const PDF_WORKER_URL = pathToFileURL(path.join(__dirname, 'pdf.worker.mjs')).href;

/**
 * 内嵌了 cMap/标准字体/wasm 的 pdf worker（webview 与宿主侧转换共用同一份）。
 *
 * 宿主侧 PDF 解析（convert）需要磁盘上的 cMap 等资源才能渲染中文等 CID 字体；
 * 打包后的扩展没有 node_modules，jdoc 的自动探测必然失败。传入该文件后，jdoc
 * 会把资源解包到系统临时目录（按内容哈希缓存，只做一次），无需在包里再放一份
 * 资源目录。
 */
const PDF_EMBEDDED_WORKER = path.join(__dirname, '..', 'media', 'pdf.worker.min.mjs');

/**
 * 给可能挂起的异步操作加超时。
 *
 * pdfjs 在 worker 加载失败（文件缺失 / 路径失效 / 宿主 CSP）时不会 reject，
 * 而是永远 pending——表现就是转换卡在「正在请求转换…」/「Converting document…」
 * 且无任何报错。这里统一兜一个超时，把静默挂起变成明确错误。
 */
function withTimeout<T>(p: Promise<T>, ms: number, label: string): Promise<T> {
    return Promise.race([
        p,
        new Promise<T>((_, reject) =>
            setTimeout(
                () => reject(new Error(`${label} 超时（${Math.round(ms / 1000)}s），可能是 PDF worker 未能加载`)),
                ms
            )
        )
    ]);
}

/** 各目标格式对应的展示名 / 扩展名 / 保存过滤器 */
const TARGET_META: Record<ConvertTarget, { label: string; ext: string; filter: Record<string, string[]> }> = {
    docx: { label: loc('extension.target.docx', 'Word Document (.docx)'), ext: 'docx', filter: { Word: ['docx'] } },
    pdf: { label: loc('extension.target.pdf', 'PDF Document (.pdf)'), ext: 'pdf', filter: { PDF: ['pdf'] } },
    md: { label: loc('extension.target.md', 'Markdown (.md)'), ext: 'md', filter: { Markdown: ['md', 'markdown'] } },
    html: { label: loc('extension.target.html', 'HTML Web Page (.html)'), ext: 'html', filter: { HTML: ['html', 'htm'] } }
};

/** 「导出为图片」在 QuickPick 里的哨兵值（库的 ConvertTarget 里没有 png，故用独立标记） */
const IMAGE_TARGET = '__images__';

const VIEW_TYPE = 'jdoc.documentEditor';

/**
 * 本地化回退：无语言包时 vscode.l10n.t 会原样返回 key，这里给出英文回退。
 */
function loc(key: string, fallback: string, ...args: string[]): string {
    const value = vscode.l10n.t(key, ...args);
    if (value && value !== key) return value;
    // 回退文案：与 webview 端一致的 {n} 位置占位符插值（支持 {0}/{1}/…），
    // 避免多参数消息在回退分支漏替换而残留字面量占位符。
    if (!args.length) return fallback;
    return fallback.replace(/\{(\d+)\}/g, (m, i) => {
        const v = args[Number(i)];
        return v == null ? m : String(v);
    });
}

function getNonce(): string {
    let text = '';
    const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    for (let i = 0; i < 32; i++) text += possible.charAt(Math.floor(Math.random() * possible.length));
    return text;
}

/** webview 回传的 data URL（image/png;base64,…）还原成 PNG 字节。 */
function dataUrlToBytes(dataUrl: string): Uint8Array {
    const comma = dataUrl.indexOf(',');
    const b64 = comma >= 0 ? dataUrl.slice(comma + 1) : dataUrl;
    return new Uint8Array(Buffer.from(b64, 'base64'));
}

class JdocEditorProvider implements vscode.CustomEditorProvider<JdocDocument> {
    public static readonly viewType = VIEW_TYPE;

    private readonly _onDidChangeCustomDocument = new vscode.EventEmitter<
        vscode.CustomDocumentEditEvent<JdocDocument> | vscode.CustomDocumentContentChangeEvent<JdocDocument>
    >();
    public readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

    private _activeDoc: JdocDocument | undefined;

    /** 当前活动的 webview 面板（标题栏命令需要向它发消息）。 */
    private _activePanel: vscode.WebviewPanel | undefined;

    /** Resolver for a pending "serialize now, before saving" handshake. */
    private _pendingFlush: (() => void) | null = null;

    /**
     * 已打开的文档（key 为 uri.toString()）。用于编辑器标题栏的「切换文档」命令：
     * 所有自定义编辑器都经由本 provider 打开，因此这里能拿到全量列表。
     */
    private readonly _openDocs = new Map<string, DocTabInfo>();

    constructor(private readonly context: vscode.ExtensionContext) {}

    async openCustomDocument(
        uri: vscode.Uri,
        _openContext: vscode.CustomDocumentOpenContext,
        _token: vscode.CancellationToken
    ): Promise<JdocDocument> {
        const document = await JdocDocument.create(uri);
        document.onDidChange(() => {
            this._onDidChangeCustomDocument.fire({ document });
            this.syncDocTab(document);
        });
        document.onDidRevert(() => {
            this._onDidChangeCustomDocument.fire({ document });
            this.syncDocTab(document);
        });
        document.onDidSave(() => {
            this._onDidChangeCustomDocument.fire({ document });
            this.syncDocTab(document);
        });
        return document;
    }

    async resolveCustomEditor(
        document: JdocDocument,
        webviewPanel: vscode.WebviewPanel,
        _token: vscode.CancellationToken
    ): Promise<void> {
        webviewPanel.webview.options = {
            enableScripts: true,
            enableCommandUris: false,
            // 文档所在目录也加入资源根，使 webview 能直接 fetch 原始文件（大 PDF 免 base64 卡死）
            localResourceRoots: [this.context.extensionUri, vscode.Uri.file(path.dirname(document.uri.fsPath))]
        };
        webviewPanel.webview.html = this.getWebviewHtml(webviewPanel.webview, document.format);

        const key = document.uri.toString();
        const title = document.uri.path.split('/').pop() || 'document';
        this._openDocs.set(key, { uri: key, title, format: document.format, dirty: document.isDirty });
        webviewPanel.onDidDispose(() => {
            this._openDocs.delete(key);
        });

        webviewPanel.onDidChangeViewState((e) => {
            if (e.webviewPanel.active) {
                this._activeDoc = document;
                this._activePanel = webviewPanel;
                this.syncContext();
            }
        });
        // 立即登记为活动文档：部分宿主在打开面板时不一定触发 onDidChangeViewState，
        // 不登记的话标题栏「转换格式」按钮（jdoc.convert 命令，不经 webview）取不到
        // _activeDoc，convertUri(undefined) 会落到 undefined，只弹一条易被忽略的提示，
        // 表现成「点了没反应」。焦点变化后上面的事件仍会持续校正。
        this._activeDoc = document;
        this._activePanel = webviewPanel;

        const send = (msg: HostToWebview) => webviewPanel.webview.postMessage(msg);
        const sendInit = async () => {
            const s = document.snapshot();
            send({
                type: 'init',
                format: s.format,
                mode: s.mode,
                bytes: s.bytes,
                // 可访问的原始文件 URI：webview 优先 fetch 它取二进制，绕开大文件 base64 瓶颈
                fileUri: webviewPanel.webview.asWebviewUri(document.uri).toString(),
                title,
                uri: key,
                dirty: document.isDirty,
                locale: vscode.env.language
            });
        };

        document.onDidSave(() => send({ type: 'saved' }));
        document.onDidRevert(() => {
            void sendInit();
            send({ type: 'reverted' });
        });

        webviewPanel.webview.onDidReceiveMessage(async (message: WebviewToHost) => {
            try {
                switch (message.type) {
                    case 'ready':
                        await sendInit();
                        break;
                    case 'setMode':
                        document.mode = message.mode;
                        if (this._activeDoc === document) this.syncContext();
                        break;
                    case 'sync':
                        document.setBytes(message.bytes);
                        document.setDirty(message.dirty);
                        this._pendingFlush?.();
                        break;
                    case 'flushDone':
                        this._pendingFlush?.();
                        break;
                    case 'flush':
                        // 保存前立即把挂起的编辑序列化，避免 Ctrl+S 抢在 350ms 防抖 / 异步序列化之前。
                        // syncBytes 会发出 `sync`（上面的分支释放等待），预览模式则回 `flushDone`。
                        await this.flushPendingSync(document);
                        break;
                    case 'save':
                        // 走 VS Code 标准保存：它会调用 saveCustomDocument（含保存前 flush 握手），
                        // 并正确清除标题栏的未保存标记。直接 document.save() 会绕过 VS Code。
                        await vscode.commands.executeCommand('workbench.action.files.save');
                        break;
                    case 'saveAs':
                        await this.saveAs(document, message.bytes);
                        break;
                    case 'exportImage':
                        await this.saveImageFile(document, message);
                        break;
                    case 'exportImages':
                        await this.saveImageFiles(document, message);
                        break;
                    case 'exportPdf':
                        await this.savePdf(document, message);
                        break;
                    case 'revert':
                        await document.revert();
                        await sendInit();
                        send({ type: 'info', kind: 'info', message: loc('extension.reloaded', 'Reloaded') });
                        break;
                    case 'convert': {
                        // 状态栏按钮 / 内容区右键：优先用 webview 带来的 uri，
                        // 兜底用当前面板的 document.uri（本 handler 就绑在这个面板上）。
                        // 绝不能让 convertUri 去猜 activeTextEditor——同时开着别的文本
                        // 文件时会转错对象（Markdown 走内置编辑器时尤其明显）。
                        let target: vscode.Uri | undefined;
                        if (message.uri) {
                            try {
                                target = vscode.Uri.parse(message.uri, true);
                            } catch {
                                target = undefined;
                            }
                        }
                        const finalTarget = target ?? document.uri;
                        // 写日志：转换链路出问题时，用户在「输出 → Log (Extension Host)」
                        // 里能直接看到请求有没有到达宿主、目标文件解析成了什么。
                        console.log(`[jdoc] convert 请求到达，uri=${finalTarget?.toString() ?? '(none)'}`);
                        // 先回执：让 webview 知道消息已送达、选择框即将弹出。
                        send({ type: 'convertAck' });
                        // 走命令路径，与 TAB 右键 / 资源管理器右键完全一致。
                        await vscode.commands.executeCommand('jdoc.convert', finalTarget);
                        break;
                    }
                    case 'error':
                        send({ type: 'info', kind: 'error', message: message.message });
                        break;
                }
            } catch (err: any) {
                send({ type: 'info', kind: 'error', message: err?.message || String(err) });
            }
        });

        try {
            await sendInit();
        } catch (err: any) {
            send({
                type: 'info',
                kind: 'error',
                message: loc('extension.initFailed', 'Initialization failed: {0}', String(err?.message || err))
            });
        }
    }

    /** 记录当前活动文档，用于标题栏命令（查看/编辑、切换文档）。 */
    private syncContext(): void {
        void vscode.commands.executeCommand('setContext', 'jdoc:editing', this._activeDoc?.mode === 'edit');
    }

    /** 标题栏「查看 / 编辑」命令：把模式推给当前活动的 webview。 */
    public async setActiveMode(mode: EditorMode): Promise<void> {
        const doc = this._activeDoc;
        if (!doc) return;
        doc.mode = mode;
        this.syncContext();
        await this._activePanel?.webview.postMessage({ type: 'setMode', mode } as HostToWebview);
    }

    /** 标题栏「切换文档」命令：在已打开的 jdoc 文档之间跳转。 */
    public async switchDocument(): Promise<void> {
        const docs = Array.from(this._openDocs.values());
        if (docs.length === 0) {
            vscode.window.showInformationMessage(loc('extension.noOpenDoc', 'No jdoc document is open.'));
            return;
        }
        const items = docs.map((d) => ({
            label: d.title,
            description: vscode.workspace.asRelativePath(vscode.Uri.parse(d.uri)),
            detail: d.dirty ? loc('extension.unsavedDoc', 'Unsaved changes') : undefined,
            uri: d.uri
        }));
        const picked = await vscode.window.showQuickPick(items, {
            placeHolder: loc('extension.switchDoc', 'Switch to an open document')
        });
        if (!picked) return;
        // 用同一个 viewType 打开：已打开则聚焦到该标签，未打开则新建。
        await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.parse(picked.uri), VIEW_TYPE);
    }

    /** 文档 dirty 状态变化时刷新已打开文档列表。 */
    private syncDocTab(document: JdocDocument): void {
        const rec = this._openDocs.get(document.uri.toString());
        if (!rec || rec.dirty === document.isDirty) return;
        rec.dirty = document.isDirty;
    }

    /**
     * Edits are pushed to the host with a 350ms debounce plus async PPTX serialization. Without
     * this handshake, Ctrl+S right after an edit can save the *previous* bytes (e.g. an animation
     * added a moment ago is absent after reopening). Ask the webview to flush and wait for its
     * `sync` reply (bounded by a timeout so saving never hangs).
     */
    private async flushPendingSync(document: JdocDocument): Promise<void> {
        const panel = this._activeDoc === document ? this._activePanel : undefined;
        if (!panel) return;
        await new Promise<void>((resolve) => {
            const timer = setTimeout(() => {
                this._pendingFlush = null;
                resolve();
            }, 2000);
            this._pendingFlush = () => {
                clearTimeout(timer);
                this._pendingFlush = null;
                resolve();
            };
            void panel.webview.postMessage({ type: 'flush' });
        });
    }

    // ---- 导出 PNG / PDF：webview 离屏渲染好位图，宿主负责文件对话框与写盘 ----
    private async saveImageFile(
        document: JdocDocument,
        msg: { bytes: string; suggestedName: string }
    ): Promise<void> {
        const uri = await vscode.window.showSaveDialog({
            defaultUri: vscode.Uri.joinPath(document.uri, '..', msg.suggestedName || 'slide.png'),
            filters: { PNG: ['png'] }
        });
        if (!uri) return;
        await vscode.workspace.fs.writeFile(uri, Buffer.from(msg.bytes, 'base64'));
        vscode.window.showInformationMessage(loc('extension.imageSaved', 'Image saved: {0}', uri.fsPath));
    }

    private async saveImageFiles(
        document: JdocDocument,
        msg: { images: { name: string; bytes: string }[] }
    ): Promise<void> {
        const folders = await vscode.window.showOpenDialog({
            canSelectFiles: false,
            canSelectFolders: true,
            canSelectMany: false,
            defaultUri: vscode.Uri.joinPath(document.uri, '..'),
            openLabel: loc('extension.chooseFolder', 'Choose Folder')
        });
        if (!folders?.length) return;
        const dir = folders[0];
        for (const img of msg.images) {
            await vscode.workspace.fs.writeFile(vscode.Uri.joinPath(dir, img.name), Buffer.from(img.bytes, 'base64'));
        }
        vscode.window.showInformationMessage(
            loc('extension.imagesSaved', 'Saved {0} images to {1}', String(msg.images.length), dir.fsPath)
        );
    }

    private async savePdf(
        document: JdocDocument,
        msg: { images: { bytes: string }[]; widthPx: number; heightPx: number; suggestedName: string }
    ): Promise<void> {
        if (!msg.images?.length) return;
        const pdf = await PDFDocument.create();
        // px → pt (1px = 0.75pt, 与解析器其它地方的换算一致)
        const w = (msg.widthPx || 1280) * 0.75;
        const h = (msg.heightPx || 720) * 0.75;
        for (const img of msg.images) {
            const embedded = await pdf.embedPng(Buffer.from(img.bytes, 'base64'));
            const page = pdf.addPage([w, h]);
            page.drawImage(embedded, { x: 0, y: 0, width: w, height: h });
        }
        const uri = await vscode.window.showSaveDialog({
            defaultUri: vscode.Uri.joinPath(document.uri, '..', msg.suggestedName || 'slides.pdf'),
            filters: { PDF: ['pdf'] }
        });
        if (!uri) return;
        await vscode.workspace.fs.writeFile(uri, Buffer.from(await pdf.save()));
        vscode.window.showInformationMessage(loc('extension.pdfSaved', 'PDF saved: {0}', uri.fsPath));
    }

    // ---- CustomEditorProvider 必需方法 ----
    async saveCustomDocument(document: JdocDocument, _token: vscode.CancellationToken): Promise<void> {
        await this.flushPendingSync(document);
        await document.save();
    }
    async saveCustomDocumentAs(
        document: JdocDocument,
        target: vscode.Uri,
        _token: vscode.CancellationToken
    ): Promise<void> {
        await document.save(target);
    }
    async revertCustomDocument(document: JdocDocument, _token: vscode.CancellationToken): Promise<void> {
        await document.revert();
    }
    async backupCustomDocument(
        document: JdocDocument,
        context: vscode.CustomDocumentBackupContext,
        _token: vscode.CancellationToken
    ): Promise<vscode.CustomDocumentBackup> {
        await document.backup(context.destination);
        return {
            id: context.destination.toString(),
            delete: () => {
                vscode.workspace.fs.delete(context.destination).then(
                    () => undefined,
                    () => undefined
                );
            }
        };
    }

    private async saveAs(document: JdocDocument, bytes?: string | null): Promise<void> {
        const filters: Record<string, string[]> = {
            PowerPoint: ['pptx'],
            Word: ['docx'],
            Excel: ['xlsx'],
            PDF: ['pdf'],
            Markdown: ['md', 'markdown'],
            Text: ['txt', 'log']
        };
        const uri = await vscode.window.showSaveDialog({
            defaultUri: document.uri,
            saveLabel: loc('extension.exportCopy', 'Export Copy'),
            filters: { Document: filters[document.format] || ['*'] }
        });
        if (!uri) return;
        if (bytes) document.setBytes(bytes);
        await document.save(uri);
        vscode.window.showInformationMessage(loc('extension.exported', 'Exported copy: {0}', uri.fsPath));
    }

    public async revertActive(): Promise<void> {
        if (this._activeDoc) await this._activeDoc.revert();
    }
    public async exportActive(): Promise<void> {
        if (this._activeDoc) await this.saveAs(this._activeDoc);
    }

    /**
     * 文档格式转换入口：把选中的文件转换成另一种格式（docx / xlsx / pdf / md / html / txt 互转）。
     * 通过 QuickPick 选择目标格式，转换后在扩展宿主（Node）完成并另存为新文件。
     */
    public async convertUri(uri?: vscode.Uri): Promise<void> {
        // uri：菜单点击时 VS Code 传入；activeTextEditor：内置编辑器打开的 md/txt；
        // _activeDoc：jdoc 自定义编辑器（此时 activeTextEditor 为 undefined），
        // 否则从命令面板触发会误报「请选择文件」。
        const target = uri || vscode.window.activeTextEditor?.document.uri || this._activeDoc?.uri;
        if (!target) {
            vscode.window.showInformationMessage(loc('extension.pickConvert', 'Select a file to convert with jdoc.'));
            return;
        }
        const sourceName = target.path.split('/').pop() || 'document';
        const from = extToFormat(target.path);
        if (!from) {
            vscode.window.showWarningMessage(
                loc(
                    'extension.unsupported',
                    'Unsupported file format for conversion. Supported: .docx, .xlsx, .pptx, .pdf, .md, .html, .txt'
                )
            );
            return;
        }

        const items = [
            { label: loc('extension.toImages', 'Images (PNG, one per page)'), target: IMAGE_TARGET },
            ...targetFormats(from, Object.keys(TARGET_META) as ConvertTarget[]).map((k) => ({
                label: TARGET_META[k].label,
                target: k
            }))
        ];
        const picked = await vscode.window.showQuickPick(items, {
            placeHolder: loc('extension.convertTo', 'Convert "{0}" to which format?', sourceName),
            title: loc('extension.convertTitle', 'jdoc: Convert Document Format')
        });
        if (!picked) return;

        if (picked.target === IMAGE_TARGET) {
            await this.exportImages(target, from, sourceName);
            return;
        }

        const to = picked.target as ConvertTarget;
        const meta = TARGET_META[to];
        await vscode.window.withProgress(
            { location: vscode.ProgressLocation.Notification, title: loc('extension.converting', 'Converting document…') },
            async () => {
                try {
                    const data = await vscode.workspace.fs.readFile(target);
                    const result = await withTimeout(
                        convert(data, to, { from, workerSrc: PDF_WORKER_URL, embeddedWorkerPath: PDF_EMBEDDED_WORKER }),
                        30000,
                        loc('extension.converting', 'Converting document…')
                    );
                    const base = sourceName.replace(/\.[^.]+$/, '');
                    const dir = target.path.replace(/[^/]+$/, '');
                    const outUri = await vscode.window.showSaveDialog({
                        defaultUri: target.with({ path: dir + base + '.' + meta.ext }),
                        saveLabel: loc('extension.saveResult', 'Save converted file'),
                        filters: meta.filter
                    });
                    if (!outUri) return;
                    const bytes = typeof result === 'string' ? new TextEncoder().encode(result) : result;
                    await vscode.workspace.fs.writeFile(outUri, bytes);
                    vscode.window.showInformationMessage(
                        loc('extension.converted', 'Converted and saved: {0}', outUri.fsPath)
                    );
                } catch (err: any) {
                    vscode.window.showErrorMessage(loc('extension.convertFailed', 'Conversion failed: {0}', String(err?.message || err)));
                }
            }
        );
    }

    /**
     * 在一次性webview 里把 PDF 逐页栅格化为 PNG。
     *
     * 为什么必须借道 webview：pdfjs 在 Node 侧渲染依赖 @napi-rs/canvas，那是约 27MB
     * 的平台相关原生二进制（skia.*.node），既无法被 esbuild 打包、也无法随扩展分发；
     * 而且 esbuild 产物里 pdfjs 的 fake worker 会因定位不到 pdf.worker.mjs 直接抛
     *「No GlobalWorkerOptions.workerSrc specified」。webview 里pdfjs 与浏览器 canvas
     * 都是现成的，中文 CJK 字形渲染也正常。
     *
     * 每渲染完一页onPage 就被调用一次并立刻落盘，内存峰值只有单页。
     * 返回实际写出的页数。
     */
    private rasterizePdf(
        pdf: Uint8Array,
        scale: number,
        onPage: (index: number, png: Uint8Array) => Promise<void>
    ): Promise<number> {
        return new Promise<number>((resolve, reject) => {
            const panel = vscode.window.createWebviewPanel(
                'jdoc.rasterize',
                loc('extension.rasterizing', 'Exporting images…'),
                { viewColumn: vscode.ViewColumn.Active, preserveFocus: true },
                { enableScripts: true, enableCommandUris: false }
            );
            panel.webview.html = this.getWebviewHtml(panel.webview, 'pdf');

            let settled = false;
            let count = 0;
            const finish = (err?: Error) => {
                if (settled) return;
                settled = true;
                panel.dispose();
                if (err) reject(err);
                else resolve(count);
            };

            panel.webview.onDidReceiveMessage((msg: WebviewToHost) => {
                switch (msg.type) {
                    case 'ready':
                        // 必须等握手：webview 脚本就绪前发出的消息会被丢弃
                        panel.webview.postMessage({
                            type: 'rasterize',
                            bytes: Buffer.from(pdf).toString('base64'),
                            scale
                        });
                        break;
                    case 'rasterizePage':
                        count = Math.max(count, msg.index);
                        void onPage(msg.index, dataUrlToBytes(msg.dataUrl)).catch((e) =>
                            finish(e instanceof Error ? e : new Error(String(e)))
                        );
                        break;
                    case 'rasterizeDone':
                        finish();
                        break;
                    case 'rasterizeError':
                        finish(new Error(msg.message));
                        break;
                    default:
                        break;
                }
            });
            panel.onDidDispose(() => finish(new Error(loc('extension.rasterizeCanceled', 'Image export canceled'))));
        });
    }

    /**
     * 导出为图片：源格式 → PDF → 每页一张 PNG，写入用户选定目录。
     *
     * 库的 ConvertTarget 只有 docx/pdf/md/html，没有 png，所以显式走两步：先用 pdf-lib
     * 转 PDF（已修子集化，中文文档也只有几百 KB），再交给一次性 webview 逐页栅格化。
     * 全程不需要任何原生依赖，没打开任何文档时也能用。
     */
    private async exportImages(target: vscode.Uri, from: SourceFormat, sourceName: string): Promise<void> {
        // 先选目录再栅格化：页数要等 webview 渲染完才知道，这样还能边渲染边落盘，
        // 不必把所有 PNG 暂存在内存里。
        const picked = await vscode.window.showOpenDialog({
            canSelectFolders: true,
            canSelectFiles: false,
            canSelectMany: false,
            title: loc('extension.pickImageFolder', 'Select image output folder')
        });
        const folder = picked?.[0];
        if (!folder) return;

        await vscode.window.withProgress(
            {
                location: vscode.ProgressLocation.Notification,
                title: loc('extension.exportingImages', 'Exporting images…')
            },
            async () => {
                try {
                    const data = await vscode.workspace.fs.readFile(target);
                    // PDF 源直接栅格化原始文件：经「解析成 blocks 再重建 PDF」的链路是有损的
                    // （封面装饰网格会被误判为表格、版面重排后页数膨胀），而 webview 里的
                    // pdfjs 能完美渲染原始页面（页数、字体、布局与查看器所见一致）。
                    // 非 PDF 源（docx/xlsx/pptx/md/html）仍需先转成 PDF。
                    const pdf =
                        from === 'pdf'
                            ? data
                            : ((await withTimeout(
                                  convert(data, 'pdf', { from, workerSrc: PDF_WORKER_URL, embeddedWorkerPath: PDF_EMBEDDED_WORKER }),
                                  30000,
                                  loc('extension.converting', 'Converting document…')
                              )) as Uint8Array);
                    const base = sourceName.replace(/\.[^.]+$/, '') || 'page';
                    // scale=2 约 144dpi，A4 输出约 1191x1684，清晰度足够印刷预览
                    const written = await this.rasterizePdf(pdf, 2, async (index, png) => {
                        await vscode.workspace.fs.writeFile(
                            vscode.Uri.joinPath(folder, `${base}-${index}.png`),
                            png
                        );
                    });
                    vscode.window.showInformationMessage(
                        `${loc('extension.imagesExported', 'Exported {0} images', String(written))} → ${folder.fsPath}`
                    );
                } catch (err: any) {
                    vscode.window.showErrorMessage(
                        loc('extension.convertFailed', 'Conversion failed: {0}', String(err?.message || err))
                    );
                }
            }
        );
    }

    private getWebviewHtml(webview: vscode.Webview, format: DocFormat): string {
        // pptx -> 幻灯片编辑器；docx/xlsx -> office 编辑器；pdf/md/txt -> 只读查看器
        const scriptFile =
            format === 'pptx'
                ? 'pptx-editor.js'
                : format === 'docx' || format === 'xlsx'
                    ? 'office-editor.js'
                    : 'webview.js';
        const scriptUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.context.extensionUri, 'dist', scriptFile)
        );
        // 图表依赖：echarts / echarts-gl 作为全局脚本按序加载（必须先于 webview.js）
        const echartsUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.context.extensionUri, 'media', 'echarts.min.js')
        );
        const echartsGlUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.context.extensionUri, 'media', 'echarts-gl.min.js')
        );
        // PDF（pdfjs）worker：webview 是单一打包产物，import.meta.url 派生的 worker 地址无效，
        // 因此这里显式给出可访问的 worker 资源 URL，由 webview 端写入全局供 jdoc 复用。
        const pdfWorkerUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.context.extensionUri, 'media', 'pdf.worker.min.mjs')
        );
        // pdfjs 资源目录占位：cMap/标准字体/wasm 已由构建脚本内嵌进 worker，
        // webview 端识别到内嵌后跳过 URL 探测，这些地址只是让 pdfjs 能拼出
        // 文件名的占位符（media/pdfjs 目录已不再随包分发）。
        const pdfAssetUri = (sub: string): string =>
            webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'media', 'pdfjs', sub)).toString();
        // KaTeX 样式（Markdown 公式渲染），CSS 内部以相对路径引用 media/fonts/*.woff2
        const katexCssUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.context.extensionUri, 'media', 'katex.min.css')
        );
        const nonce = getNonce();
        const oeTheme =
            vscode.window.activeColorTheme.kind === vscode.ColorThemeKind.Dark ||
            vscode.window.activeColorTheme.kind === vscode.ColorThemeKind.HighContrast
                ? 'dark'
                : 'light';
        const csp = [
            `default-src 'none';`,
            `img-src ${webview.cspSource} data: blob:;`,
            `media-src ${webview.cspSource} data: blob:;`,
            // KaTeX 字体（Markdown 公式）
            `font-src ${webview.cspSource} data:;`,
            `style-src ${webview.cspSource} 'unsafe-inline';`,
            // echarts-gl 内部用 new Function() 解析 expr()，需 unsafe-eval 否则 3D 图表抛错；
            // pdfjs worker 以 Worker 形式加载，需 worker-src（含 blob: 兜底）
            `script-src 'nonce-${nonce}' 'unsafe-eval' ${webview.cspSource};`,
            `worker-src ${webview.cspSource} blob: data:;`,
            // 允许 webview 直接 fetch 宿主给的文件 URI（asWebviewUri 同源），大文件免 base64 卡死
            `connect-src ${webview.cspSource} blob: data:;`
        ].join(' ');
        return `<!DOCTYPE html>
<html lang="${vscode.env.language}" data-oe-theme="${oeTheme}">
<head>
<meta charset="utf-8" />
<meta http-equiv="Content-Security-Policy" content="${csp}" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<link rel="stylesheet" href="${katexCssUri}" />
<title>jdoc Document</title>
</head>
<body>
<div id="app"></div>
<script nonce="${nonce}">window.__JDC_PDF_WORKER__ = "${pdfWorkerUri}";
window.__JDC_PDF_CMAPS__ = "${pdfAssetUri('cmaps/')}";
window.__JDC_PDF_FONTS__ = "${pdfAssetUri('standard_fonts/')}";
window.__JDC_PDF_WASM__ = "${pdfAssetUri('wasm/')}";</script>
<script nonce="${nonce}" src="${echartsUri}"></script>
<script nonce="${nonce}" src="${echartsGlUri}"></script>
<script type="module" nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
    }
}

export function activate(context: vscode.ExtensionContext): void {
    const provider = new JdocEditorProvider(context);
    context.subscriptions.push(
        vscode.window.registerCustomEditorProvider(VIEW_TYPE, provider, {
            webviewOptions: { retainContextWhenHidden: true },
            supportsMultipleEditorsPerDocument: false
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand('jdoc.revert', () => provider.revertActive()),
        vscode.commands.registerCommand('jdoc.exportCopy', () => provider.exportActive()),
        vscode.commands.registerCommand('jdoc.convert', (uri?: vscode.Uri) => provider.convertUri(uri)),
        // 标题栏（editor/title）里的查看/编辑与文档切换
        vscode.commands.registerCommand('jdoc.viewMode', () => provider.setActiveMode('view')),
        vscode.commands.registerCommand('jdoc.editMode', () => provider.setActiveMode('edit')),
        vscode.commands.registerCommand('jdoc.switchDocument', () => provider.switchDocument())
    );
}

export function deactivate(): void {
    // no-op
}

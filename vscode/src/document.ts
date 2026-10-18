import * as vscode from 'vscode';
import type { DocFormat, EditorMode } from './protocol';

/** 根据文件扩展名推断文档格式。 */
export function formatOfUri(uri: vscode.Uri): DocFormat {
    const ext = uri.path.split('.').pop()?.toLowerCase();
    if (ext === 'docx') return 'docx';
    if (ext === 'xlsx') return 'xlsx';
    if (ext === 'pdf') return 'pdf';
    return 'pptx'; // 默认按 pptx 处理未知 OOXML
}

/**
 * 通用文档模型：一个轻量的字节持有者。
 *
 * 所有解析/模型/编辑都发生在 webview（浏览器环境），使用 @fefeding/jdoc，
 * 因此 Host 永远不需要 DOM、也不需要触碰文档模型。职责仅限于：读文件、
 * 保存 webview 产出的最新字节、驱动 save / revert / backup / dirty 记账。
 */
export class JdocDocument implements vscode.CustomDocument {
    public readonly uri: vscode.Uri;
    public readonly format: DocFormat;

    private _original!: Uint8Array; // 从磁盘加载的文件字节
    private _current!: Uint8Array; // 最新字节（webview 编辑产出；否则等同原始）
    private _dirty = false;
    private _mode: EditorMode = 'view';

    private readonly _onDidChange = new vscode.EventEmitter<{ contentChanged: boolean }>();
    public readonly onDidChange = this._onDidChange.event;
    private readonly _onDidRevert = new vscode.EventEmitter<void>();
    public readonly onDidRevert = this._onDidRevert.event;
    private readonly _onDidSave = new vscode.EventEmitter<void>();
    public readonly onDidSave = this._onDidSave.event;

    private constructor(uri: vscode.Uri, format: DocFormat) {
        this.uri = uri;
        this.format = format;
    }

    static async create(uri: vscode.Uri): Promise<JdocDocument> {
        const doc = new JdocDocument(uri, formatOfUri(uri));
        const bytes = await vscode.workspace.fs.readFile(uri);
        doc._original = bytes;
        doc._current = bytes;
        return doc;
    }

    get mode(): EditorMode {
        return this._mode;
    }
    set mode(m: EditorMode) {
        this._mode = m;
    }

    get isDirty(): boolean {
        return this._dirty;
    }

    /** 接受 webview 产出的字节（base64）。bytes 为 null 表示放弃编辑回到导入态。 */
    setBytes(bytes: string | null): void {
        this._current = bytes ? Buffer.from(bytes, 'base64') : this._original;
        this.setDirty(true);
    }

    setDirty(dirty: boolean): void {
        if (this._dirty === dirty) return;
        this._dirty = dirty;
        this._onDidChange.fire({ contentChanged: dirty });
    }

    /** 给 webview 的载荷：原始字节 + 最新字节（均为 base64）。 */
    snapshot(): { format: DocFormat; mode: EditorMode; originalBytes: string; bytes: string } {
        return {
            format: this.format,
            mode: this._mode,
            originalBytes: Buffer.from(this._original).toString('base64'),
            bytes: Buffer.from(this._current).toString('base64')
        };
    }

    // ---- Save / Revert / Backup ----
    async save(targetUri?: vscode.Uri): Promise<void> {
        const dest = targetUri || this.uri;
        await vscode.workspace.fs.writeFile(dest, this._current);
        if (!targetUri) {
            this._original = this._current;
            this.setDirty(false);
            this._onDidSave.fire();
        }
    }

    async revert(): Promise<void> {
        const bytes = await vscode.workspace.fs.readFile(this.uri);
        this._original = bytes;
        this._current = bytes;
        this.setDirty(false);
        this._onDidRevert.fire();
    }

    async backup(destination: vscode.Uri): Promise<void> {
        await vscode.workspace.fs.writeFile(destination, this._current);
    }

    dispose(): void {
        this._onDidChange.dispose();
        this._onDidRevert.dispose();
        this._onDidSave.dispose();
    }
}

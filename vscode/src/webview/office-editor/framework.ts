// office-editor 共享框架：编辑器基类（dirty 记账 + 防抖 sync + 序列化基架）与全局 toast。

import { debounce } from './util';

export interface EditorHooks {
    /** 编辑产出新字节（Uint8Array）；由 main 转 base64 后经 sync 回 host。 */
    onSync(bytes: Uint8Array): void;
    onToast(message: string, kind?: 'info' | 'warn' | 'error'): void;
}

/**
 * DOCX / XLSX 编辑器的公共基类。
 * 子类实现：load（字节→模型）、serialize（模型→字节）、render（把工具栏与内容渲染到宿主容器）。
 */
export abstract class EditorBase {
    protected dirty = false;
    private syncing = false;
    private syncQueued = false;
    protected toolbarEl: HTMLElement | null = null;
    protected contentEl: HTMLElement | null = null;

    private readonly syncDebounced: () => void;

    constructor(protected readonly hooks: EditorHooks) {
        this.syncDebounced = debounce(() => void this.syncNow(), 400);
    }

    mount(toolbar: HTMLElement, content: HTMLElement): void {
        this.toolbarEl = toolbar;
        this.contentEl = content;
    }

    /** 载入文件字节并解析为编辑模型，随后渲染。 */
    abstract load(bytes: Uint8Array): Promise<void>;

    /** 把当前编辑模型生成为文件字节。 */
    protected abstract serialize(): Promise<Uint8Array>;

    /** 清理 DOM / 监听。 */
    abstract destroy(): void;

    get isDirty(): boolean {
        return this.dirty;
    }

    resetDirty(): void {
        this.dirty = false;
    }

    protected markDirty(): void {
        this.dirty = true;
        this.syncDebounced();
    }

    /**
     * 立即序列化并回传字节（切模式 / 保存前调用）。
     * 带并发锁：序列化进行中到来的再次同步只排队一次；成功回传后清零 dirty
     * （dirty 表示「有未同步到 host 的编辑」，而非「文档未保存」——后者由 host 记账）。
     */
    async syncNow(): Promise<void> {
        if (this.syncing) {
            this.syncQueued = true;
            return;
        }
        if (!this.dirty) return;
        this.syncing = true;
        try {
            const bytes = await this.serialize();
            this.dirty = false;
            this.hooks.onSync(bytes);
        } catch (e: unknown) {
            this.hooks.onToast('保存失败：' + errMsg(e), 'error');
        } finally {
            this.syncing = false;
            if (this.syncQueued) {
                this.syncQueued = false;
                void this.syncNow();
            }
        }
    }

    protected toast(message: string, kind: 'info' | 'warn' | 'error' = 'info'): void {
        this.hooks.onToast(message, kind);
    }
}

export function errMsg(e: unknown): string {
    return (e as { message?: string })?.message || String(e);
}

// ---------- 全局 toast ----------
let toastEl: HTMLElement | null = null;
let toastTimer = 0;
export function toast(message: string, kind: 'info' | 'warn' | 'error' = 'info'): void {
    if (!toastEl) {
        toastEl = document.createElement('div');
        toastEl.className = 'oe-toast';
        document.body.appendChild(toastEl);
    }
    toastEl.textContent = message;
    toastEl.className = 'oe-toast show' + (kind === 'error' ? ' err' : '');
    if (toastTimer) window.clearTimeout(toastTimer);
    toastTimer = window.setTimeout(() => toastEl?.classList.remove('show'), 2600);
}

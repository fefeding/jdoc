// Message protocol between extension host (Node) and webview (browser).
//
// The webview owns the document model and all editing (via @fefeding/ppt-parser's
// editor core: createStore / createActions / docFromPptx / docToPptx), so the host only
// moves PPTX bytes and tracks the dirty state. Both payloads are base64: VS Code's webview
// typed-array transport ($$vscode_array_buffer_reference$$) does not restore Uint8Array in
// custom editors, passing binary directly makes parsing fail.

export type EditorMode = 'preview' | 'edit';

/** 顶栏「文档切换标签」的一项（host 广播的已打开文档列表）。 */
export interface DocTabInfo {
  /** 文档 uri.toString()，点击时原样回传给 host。 */
  uri: string;
  /** 文件名，用作标签文案。 */
  title: string;
  /** 是否有未保存的修改（标签上显示小圆点）。 */
  dirty: boolean;
}

/** Webview → Host */
export type WebviewToHost =
  | { type: 'ready' }
  | { type: 'setMode'; mode: EditorMode }
  /**
   * Editor produced new PPTX bytes (edit mode). `bytes` is null when the edit mode was
   * left without any edit (e.g. undo back to the imported state), meaning "keep original".
   */
  | { type: 'sync'; bytes: string | null; dirty: boolean }
  | { type: 'save' }
  | { type: 'saveAs'; bytes: string | null }
  | { type: 'revert' }
  /** 顶栏文档标签被点击：切换到（或打开）该文档。 */
  | { type: 'openDocument'; uri: string }
  /** Surface a webview-side problem (parse/serialize failure) in the host UI. */
  | { type: 'error'; message: string };

/** Host → Webview */
export type HostToWebview =
  | {
    type: 'init';
    mode: EditorMode;
    /** Original file bytes (base64) — preview mode renders these faithfully. */
    originalBytes: string;
    /** Latest bytes (base64); equals originalBytes until the webview reports an edit. */
    bytes: string;
    title: string;
    /** 当前文档的 uri.toString()，用于在文档标签里高亮自己。 */
    uri: string;
    /** True when the document has unsaved changes (drives the webview's dirty dot). */
    dirty: boolean;
    /**
     * VS Code display language (e.g. `zh-cn`, `en`). The webview also reads `<html lang>`,
     * which the host fills with the same value; this field is the authoritative fallback.
     */
    locale: string;
  }
  /** 已打开文档列表：驱动顶栏右上角的文档切换标签。 */
  | { type: 'docs'; docs: DocTabInfo[]; activeUri: string }
  | { type: 'bytes'; bytes: string; dirty: boolean }
  | { type: 'saved' }
  | { type: 'reverted' }
  | { type: 'info'; message: string; kind?: 'info' | 'warn' | 'error' };

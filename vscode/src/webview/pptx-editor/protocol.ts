// Message protocol between extension host (Node) and webview (browser).
//
// The webview owns the document model and all editing (via @fefeding/ppt-parser's
// editor core: createStore / createActions / docFromPptx / docToPptx), so the host only
// moves PPTX bytes and tracks the dirty state. Both payloads are base64: VS Code's webview
// typed-array transport ($$vscode_array_buffer_reference$$) does not restore Uint8Array in
// custom editors, passing binary directly makes parsing fail.

export type EditorMode = 'preview' | 'edit';

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
  /** 请求宿主对当前文档执行 jdoc.convert（内容区右键「转换格式」）。 */
  | { type: 'convert'; uri?: string }
  /** Surface a webview-side problem (parse/serialize failure) in the host UI. */
  | { type: 'error'; message: string };

/** Host → Webview */
export type HostToWebview =
  | {
    type: 'init';
    mode: EditorMode;
    /** Original file bytes (base64). Host sends only one copy, so this is usually absent. */
    originalBytes?: string;
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
  /** 切换预览/编辑模式（由 VS Code 标题栏的 jdoc.viewMode / jdoc.editMode 命令触发）。 */
  | { type: 'setMode'; mode: EditorMode }
  | { type: 'bytes'; bytes: string; dirty: boolean }
  | { type: 'saved' }
  | { type: 'reverted' }
  | { type: 'info'; message: string; kind?: 'info' | 'warn' | 'error' };

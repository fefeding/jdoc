// Host (Node) ↔ Webview (browser) 消息协议。
//
// Webview 持有文档模型并负责渲染/编辑（使用 @fefeding/jdoc 的解析与生成能力），
// Host 仅负责在磁盘读写字节与维护 dirty 状态。所有字节以 base64 传递：
// VS Code 自定义编辑器里 typed-array 传输无法在 webview 端还原为 Uint8Array，
// 直接传二进制会导致解析失败。

export type DocFormat = 'pptx' | 'docx' | 'xlsx' | 'pdf' | 'md' | 'txt';
export type EditorMode = 'view' | 'edit';

/** Webview → Host */
export type WebviewToHost =
  | { type: 'ready' }
  | { type: 'setMode'; mode: EditorMode }
  /** 编辑产生的新字节（edit 模式）。bytes 为 null 表示放弃编辑回到导入态。 */
  | { type: 'sync'; bytes: string | null; dirty: boolean }
  | { type: 'save' }
  | { type: 'saveAs'; bytes: string | null }
  | { type: 'revert' }
  /** 把 webview 端的解析/生成错误暴露到 Host UI。 */
  | { type: 'error'; message: string };

/** Host → Webview */
export type HostToWebview =
  | {
    type: 'init';
    format: DocFormat;
    mode: EditorMode;
    /** 原始文件字节（base64）—— 预览模式忠实渲染。 */
    bytes: string;
    title: string;
    dirty: boolean;
    /** VS Code 显示语言（如 zh-cn / en），作为权威回退。 */
    locale: string;
  }
  | { type: 'reverted' }
  | { type: 'saved' }
  | { type: 'info'; message: string; kind?: 'info' | 'warn' | 'error' };

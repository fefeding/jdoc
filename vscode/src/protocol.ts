// Host (Node) ↔ Webview (browser) 消息协议。
//
// Webview 持有文档模型并负责渲染/编辑（使用 @fefeding/jdoc 的解析与生成能力），
// Host 仅负责在磁盘读写字节与维护 dirty 状态。所有字节以 base64 传递：
// VS Code 自定义编辑器里 typed-array 传输无法在 webview 端还原为 Uint8Array，
// 直接传二进制会导致解析失败。

export type DocFormat = 'pptx' | 'docx' | 'xlsx' | 'pdf' | 'md' | 'txt';
export type EditorMode = 'view' | 'edit';

/** 顶栏「文档切换标签」的一项（由 host 维护，每次变更后广播给所有 webview）。 */
export interface DocTabInfo {
  /** 文档 uri.toString()，点击时原样回传给 host。 */
  uri: string;
  /** 文件名，用作标签文案。 */
  title: string;
  format: DocFormat;
  /** 是否有未保存的修改（标签上显示小圆点）。 */
  dirty: boolean;
}

/** Webview → Host */
export type WebviewToHost =
  | { type: 'ready' }
  | { type: 'setMode'; mode: EditorMode }
  /** 编辑产生的新字节（edit 模式）。bytes 为 null 表示放弃编辑回到导入态。 */
  | { type: 'sync'; bytes: string | null; dirty: boolean }
  | { type: 'save' }
  | { type: 'saveAs'; bytes: string | null }
  | { type: 'revert' }
  /** 顶栏文档标签被点击：切换到（或打开）该文档。 */
  | { type: 'openDocument'; uri: string }
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
    /** 当前文档的 uri.toString()，用于在文档标签里高亮自己。 */
    uri: string;
    dirty: boolean;
    /** VS Code 显示语言（如 zh-cn / en），作为权威回退。 */
    locale: string;
  }
  /** 已打开文档列表：驱动顶栏右上角的文档切换标签。 */
  | { type: 'docs'; docs: DocTabInfo[]; activeUri: string }
  /** 宿主侧文档字节被外部更新（目前未主动发送，保留给 webview 的刷新分支）。 */
  | { type: 'bytes'; bytes: string; dirty: boolean }
  | { type: 'reverted' }
  | { type: 'saved' }
  | { type: 'info'; message: string; kind?: 'info' | 'warn' | 'error' };

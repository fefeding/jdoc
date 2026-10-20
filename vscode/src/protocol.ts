// Host (Node) ↔ Webview (browser) 消息协议。
//
// Webview 持有文档模型并负责渲染/编辑（使用 @fefeding/jdoc 的解析与生成能力），
// Host 仅负责在磁盘读写字节与维护 dirty 状态。所有字节以 base64 传递：
// VS Code 自定义编辑器里 typed-array 传输无法在 webview 端还原为 Uint8Array，
// 直接传二进制会导致解析失败。

export type DocFormat = 'pptx' | 'docx' | 'xlsx' | 'pdf' | 'md' | 'txt';
export type EditorMode = 'view' | 'edit';

/** 已打开文档的一项（host 侧维护，供标题栏「切换文档」命令列出）。 */
export interface DocTabInfo {
  /** 文档 uri.toString()。 */
  uri: string;
  /** 文件名，用作列表文案。 */
  title: string;
  format: DocFormat;
  /** 是否有未保存的修改。 */
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
  /** 状态栏「转换格式」按钮：请求宿主对当前文档执行 jdoc.convert（无参数=活动文档）。 */
/** 请求宿主对当前文档执行 jdoc.convert。uri 为当前文档的 uri（string），
 *  必须带上——否则宿主只能靠 activeTextEditor 猜测，同时开着其它文本文件时会转错对象。 */
| { type: 'convert'; uri?: string }
  /** 把webview 端的解析/生成错误暴露到 Host UI。 */
  | { type: 'error'; message: string }
  /** 栅格化结果：每渲染完一页立刻回传一页，避免整本文档的 PNG 挤在一条消息里。 */
  | { type: 'rasterizePage'; index: number; dataUrl: string }
  | { type: 'rasterizeDone'; count: number }
  | { type: 'rasterizeError'; message: string };

/** Host → Webview */
export type HostToWebview =
  | {
    type: 'init';
    format: DocFormat;
    mode: EditorMode;
    /** 原始文件字节（base64）—— 预览模式忠实渲染。 */
    bytes: string;
    /**
     * 可访问的原始文件 URI（宿主侧 asWebviewUri 结果）。大文件优先走 fetch 取
     * 二进制，避免把整份 base64 经 postMessage 序列化卡死主线程；缺失时回退 bytes。
     */
    fileUri?: string;
    title: string;
    /** 当前文档的 uri.toString()。 */
    uri: string;
    dirty: boolean;
    /** VS Code 显示语言（如 zh-cn / en），作为权威回退。 */
    locale: string;
  }
  /** 切换查看/编辑（由标题栏的 jdoc.viewMode / jdoc.editMode 命令触发）。 */
  | { type: 'setMode'; mode: EditorMode }
  /** 宿主侧文档字节被外部更新（目前未主动发送，保留给 webview 的刷新分支）。 */
  | { type: 'bytes'; bytes: string; dirty: boolean }
  | { type: 'reverted' }
  | { type: 'saved' }
  | { type: 'info'; message: string; kind?: 'info' | 'warn' | 'error' }
  /**
   * 栅格化请求：把整份 PDF 逐页渲染成 PNG，逐页回传（见 WebviewToHost.rasterizePage）。
   *
   * 之所以由 webview 而非宿主栅格化：pdfjs 在 Node 侧渲染依赖 @napi-rs/canvas，
   * 那是约 27MB 的平台相关原生二进制，既无法被esbuild 打包也不能随扩展分发；
   * 而且打包后 pdfjs 的 fake worker 会因定位不到 pdf.worker.mjs 直接失败。
   * webview 侧pdfjs 与浏览器 canvas 都是现成的，中文 CJK 字形渲染也正常。
   */
  | { type: 'rasterize'; bytes: string; scale: number };

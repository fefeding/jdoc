// 顶层 webview 国际化（查看器 / office / pptx 共用的 chrome 与 main 框架文案）。
//
// 设计要点（与 pptx-editor/i18n.ts 同源，但独立打包，互不影响）：
//   1. 默认英文，支持简体中文；任意词条缺失一律回退英文（保证不会因漏译而显示空白）。
//   2. 语言来源：宿主把 `vscode.env.language` 写进 `<html lang="…">`，模块加载时即读取，
//      无闪烁；`init` 消息里的 `locale`（权威）在启动时再校正一次。
//   3. 新增语言：在 `BUNDLES` 再注册一份与 `en` 同构的字典即可。
type Dict = Record<string, string>;

/** 英文词条集（规范键集合，中文包类型受其约束，漏译即编译错误）。 */
const en = {
  // ---------- chrome：状态栏「转换格式」入口 / 右键菜单 ----------
  'chrome.convert': 'Convert Format',
  'chrome.convertTitle': 'Convert to another format (DOCX / PDF / Markdown / HTML)',
  'chrome.convertMenu': 'Convert Format…',
  'chrome.convertAck': 'Please choose a target format from the list',
  'chrome.convertRequesting': 'Requesting conversion…',

  // ---------- 缩放簇（状态栏） ----------
  'zoom.out': 'Zoom out',
  'zoom.fit': 'Fit',
  'zoom.fitWidth': 'Fit width',
  'zoom.actual': 'Actual size (100%)',
  'zoom.in': 'Zoom in',

  // ---------- main：内容区状态栏 / 右键菜单 / 渲染提示 ----------
  'main.renderFailed': 'Render failed: {0}',
  'main.failed': 'Render failed',
  'main.plainText': 'Plain Text',
  'main.readOnlyHint': '{0} is read-only; editing is not supported yet.',
  'main.readOnlyStatus': 'Read-only preview',
  'main.slidesStatus': '{0} slides',
  'main.chartsSuffix': ', {0} charts',
  'main.docxStatus': 'DOCX document preview',
  'main.xlsxStatus': 'XLSX spreadsheet preview',
  'main.mdStatusWithNotes': 'Markdown preview (with cleanup notes)',
  'main.mdStatus': 'Markdown preview',
  'main.txtStatus': 'Plain text preview',
  'main.extrasFailed': 'Chart/formula rendering failed: {0}',
  'main.copy': 'Copy',
  'main.canvasError': 'Cannot create canvas 2D context',

  // ---------- office-editor 框架（状态栏 / 解析提示 / 右键菜单） ----------
  'oe.readOnly': 'Read-only mode',
  'oe.parsing': 'Parsing {0}…'
} as const;

type Key = keyof typeof en;

/** 简体中文语言包（类型受 `Key` 约束：多写或漏写 key 均会编译报错）。 */
const zhCN: Record<Key, string> = {
  'chrome.convert': '转换格式',
  'chrome.convertTitle': '转换为其他格式（DOCX / PDF / Markdown / HTML）',
  'chrome.convertMenu': '转换格式…',
  'chrome.convertAck': '请在弹出的列表中选择目标格式',
  'chrome.convertRequesting': '正在请求转换…',

  'zoom.out': '缩小',
  'zoom.fit': '适应',
  'zoom.fitWidth': '适应宽度',
  'zoom.actual': '实际大小（100%）',
  'zoom.in': '放大',

  'main.renderFailed': '渲染失败：{0}',
  'main.failed': '渲染失败',
  'main.plainText': '纯文本',
  'main.readOnlyHint': '{0} 为只读查看，暂不支持编辑。',
  'main.readOnlyStatus': '只读预览',
  'main.slidesStatus': '{0} 页幻灯片',
  'main.chartsSuffix': '，{0} 个图表',
  'main.docxStatus': 'DOCX 文档预览',
  'main.xlsxStatus': 'XLSX 表格预览',
  'main.mdStatusWithNotes': 'Markdown 预览（有清洗提示）',
  'main.mdStatus': 'Markdown 预览',
  'main.txtStatus': '纯文本预览',
  'main.extrasFailed': '图表/公式渲染失败：{0}',
  'main.copy': '复制',
  'main.canvasError': '无法创建 canvas 2D 上下文',

  'oe.readOnly': '查看模式（只读）',
  'oe.parsing': '正在解析 {0}…'
};

const BUNDLES: Record<string, Dict> = { en, 'zh-cn': zhCN };

export const DEFAULT_LOCALE = 'en';

/** 把 VS Code 的 locale（zh-cn / zh_CN / zh / en-US …）归一到已注册的语言包。 */
function normalizeLocale(locale?: string | null): string {
  const raw = String(locale || '').trim().toLowerCase().replace(/_/g, '-');
  if (!raw) return DEFAULT_LOCALE;
  if (BUNDLES[raw]) return raw;
  const base = raw.split('-')[0];
  if (BUNDLES[base]) return base;
  // zh-Hans / zh-Hant / zh-TW 暂统一回退到简中；注册 'zh-tw' 后会自动命中
  if (base === 'zh') return 'zh-cn';
  return DEFAULT_LOCALE;
}

let current: string =
  (typeof document !== 'undefined' && normalizeLocale(document.documentElement.lang)) || DEFAULT_LOCALE;

/** 切换语言（init 消息到达时调用），返回是否发生变化。 */
export function setLocale(locale?: string | null): boolean {
  const next = normalizeLocale(locale);
  if (next === current) return false;
  current = next;
  document.documentElement.lang = next;
  return true;
}

function lookup(key: string): string | undefined {
  const dict = BUNDLES[current] || BUNDLES[DEFAULT_LOCALE];
  return dict[key] ?? BUNDLES[DEFAULT_LOCALE][key];
}

/** `{0}` / `{1}` 位置占位符替换。 */
function interpolate(s: string, args: any[]): string {
  if (!args.length) return s;
  return s.replace(/\{(\d+)\}/g, (m, i) => {
    const v = args[Number(i)];
    return v == null ? m : String(v);
  });
}

/** 取词条；缺 key 时返回 key 本身（便于定位漏译）。 */
export function t(key: Key, ...args: any[]): string {
  const s = lookup(key);
  return interpolate(s == null ? key : s, args);
}

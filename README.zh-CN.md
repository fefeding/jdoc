# @fefeding/jdoc

文档解析与生成统一库：**PPTX / DOCX / XLSX / PDF** 四种格式一套 API，浏览器与 Node.js 双端可用。

## 特性

- **一套命名，四种格式** — 解析统一为 `xxxToYyy(input, options)`，生成统一为 `jsonToXxx(doc, options)`
- **入参无门槛** — `ArrayBuffer` / `TypedArray` / `Buffer` / `Blob` / `File` / `base64` / `data URL` 直接传，不用自己转换
- **出参统一** — 所有生成函数一律返回 `Uint8Array`，可直接 `new Blob([bytes])` 或写入文件系统
- **JSON 即契约** — 解析结果可直接回传给生成函数，支持「解析 → 改 → 生成」往返
- **自带 HTML 渲染** — DOCX / XLSX 解析结果可直接渲染成带内联样式的 HTML
- **零框架依赖** — 纯 TypeScript，同时输出 ESM / CJS / 浏览器产物

## 安装

```bash
npm install @fefeding/jdoc
```

## 快速开始

### 解析

```javascript
import { docxToHtml, xlsxToJson, pptxToStandard, pdfToText, detectFormat } from '@fefeding/jdoc';

// 直接把 File / Blob 丢进去，不用 arrayBuffer()
const html = await docxToHtml(file);                 // DOCX -> HTML
const { sheets } = await xlsxToJson(file);           // XLSX -> 工作簿结构
const doc = await pptxToStandard(file);              // PPTX -> 标准文档结构
const text = await pdfToText(file);                  // PDF -> 全文

console.log(await detectFormat(file));               // 'docx' | 'xlsx' | 'pptx' | 'pdf'
```

### 生成

```javascript
import { jsonToDocx, jsonToXlsx, jsonToPptx, jsonToPdf } from '@fefeding/jdoc';

const docx = await jsonToDocx([
  { type: 'heading', level: 1, text: '周报' },
  { type: 'paragraph', text: '本周完成 A、B 两个模块。' },
  { type: 'table', header: true, rows: [['模块', '进度'], ['A', '100%']] }
]);

const xlsx = await jsonToXlsx([
  { name: '订单', rows: [['姓名', '金额'], ['张三', 1200]] }
]);

const pdf = await jsonToPdf({
  metadata: { title: '报告' },
  elements: [
    { type: 'heading', level: 1, text: '季度报告' },
    { type: 'text', text: '正文内容……' },
    { type: 'table', header: true, rows: [['项', '值'], ['营收', '1200']] }
  ]
});

const pptx = await jsonToPptx({
  slides: [{ elements: [{ type: 'text', text: 'Hello', x: 50, y: 50, width: 200, height: 40 }] }]
});

// 浏览器：触发下载
const url = URL.createObjectURL(new Blob([docx], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
const a = Object.assign(document.createElement('a'), { href: url, download: 'report.docx' });
a.click();
URL.revokeObjectURL(url);
```

## API 一览

### 通用

| 函数 | 说明 |
| --- | --- |
| `detectFormat(input)` | 按魔数与 ZIP 内部结构识别格式，返回 `'docx' \| 'xlsx' \| 'pptx' \| 'pdf' \| null` |
| `JdocError` | 统一错误类型，带 `format` / `code` / `cause` |
| `toUint8Array(input)` / `toArrayBuffer(input)` / `bytesToBase64(bytes)` | 入参与编码工具 |
| `parseXml(xml)` / `findAll` / `findFirst` / `attr` / `textOf` | 零依赖 XML 解析工具 |

### 二进制资源（图片等）

文档里的图片等二进制内容统一抽象为**资源（asset）**：一份字节 + 位置锚点。
解析时用 `assets` 选项决定给不给独立资源表，生成时用引用写回。

| 模式 | 文档结构里 | `assets` | 适用场景 |
| --- | --- | --- | --- |
| `inline`（DOCX/XLSX/PPTX 默认） | data URL | 无 | 浏览器直接显示，保持既有行为 |
| `extract` | `{ assetId }` 引用 | 有 | 后台存储 / 另存为，JSON 体积小 |
| `both`（PPTX 默认） | data URL + `assetId` | 有 | 前端显示与后端取字节并行 |
| `ignore`（PDF 默认） | — | 无 | 完全丢弃 |

各格式的默认与能力：

| 格式 | 默认 | 可提取 | 锚点 `kind` |
| --- | --- | --- | --- |
| DOCX | `inline` | 块内图片 | `block`（表格内为 `cells[r][c]` 路径） |
| XLSX | `ignore` | 单元格锚定图片 | `cell`（如 `Sheet1!B3`） |
| PPTX | `both` | 元素图片（含形状填充） | `slide`（第几张幻灯片） |
| PDF | `ignore` | 附件 + 页面图像 | `other`（附件）、`page`（页码） |

> PDF 的页面图像需走一遍渲染指令，代价较高（上百页文档明显变慢），因此默认不提取；
> 附件提取很廉价，显式开启 `assets: 'extract'` 即可一并拿到。

```js
// 解析：抽出资源，块上只留引用（DOCX / XLSX / PPTX 同理）
const parsed = await docxToJson(file, { assets: 'extract' });
// const sheet   = await xlsxToJson(file, { assets: 'extract' });  // 单元格图片
// const deck    = await pptxToStandard(file, { assets: 'extract' }); // 幻灯片图片
// const pdf     = await pdfToJson(file, { assets: 'extract' });    // 附件 + 页面图像
parsed.assets[0];
// {
//   id: 'docx-0001',
//   name: 'docx-0001.png',
//   mime: 'image/png',
//   data: Uint8Array,
//   part: 'word/media/image1.png',   // 容器内原始部件路径
//   anchor: { scope: 'docx', kind: 'block', index: 1, path: 'blocks[1]', label: '第 2 个块' }
// }

// 同一份字节，按端取最合适的形式
assetToDataUrl(asset)      // data:image/png;base64,…      浏览器直接 <img src>
assetToBlob(asset)         // Blob                         浏览器预览 / 下载
assetToBase64(asset)       // base64                       存库字段
assetToArrayBuffer(asset)  // ArrayBuffer                  底层库 / postMessage
assetToNodeStream(asset)   // Node Readable                后台直接 pipe 落盘或转存对象存储
assetToWebStream(asset)    // ReadableStream               fetch / Response

// 生成：按引用写回，往返不必搬运 base64
const bytes = await jsonToDocx(parsed, { assets: parsed.assets });
```

**四种格式的引用写法一致**，图片字段给 `{ assetId }` 或直接给资源对象即可：

```js
jsonToDocx({ blocks: [{ type: 'image', src: { assetId: 'docx-0001' } }] }, { assets });
jsonToXlsx({ sheets: [{ name: 'S', rows: [['x', { image: { assetId: 'xlsx-0001' } }]] }] }, { assets });
jsonToPptx({ slides: [{ elements: [{ type: 'image', data: { assetId: 'pptx-0001' } }] }] }, { assets });
jsonToPdf({ elements: [{ type: 'image', src: { assetId: 'pdf-0001' } }] }, { assets });
```

- 资源表也可直接挂在文档规格上（`{ assets, blocks/elements/slides }`），无需再传选项
- 资源缺失时 DOCX 降级为占位段落、PDF 跳过该图，均不中断整篇生成
- PPTX 的 `extract` 模式元素上只留 `assetId`，生成时自动还原成 data URL

资源是**纯数据**（不含方法），可直接 `JSON.stringify` 落库；字节单独取用。
同一张图在文档里出现多次只存一份字节，块共享同一个 `assetId`。
`anchor` 让「资源 ↔ 文档位置」双向可查——既能由块找图片，也能由图片反查它在第几个块 / 哪张表哪格。

### 文档一体保存

解析结果与资源可打成自包含的 bundle（ZIP），解压即可还原，任意机器可编辑：

```js
const zip = await packDocument({
    format: 'docx',
    name: '技术协议',
    document: parsed,
    assets: parsed.assets,
    original: file   // 可选：一并保留原始文件
});
// 技术协议.jdbundle
// ├── manifest.json      格式、名称、资源清单（含位置锚点）
// ├── 技术协议.docx        原始文档
// ├── document.json      解析出的结构，可回传给 jsonToDocx
// └── assets/docx-0001.png

const restored = await unpackDocument(zip);      // 解包还原结构与字节
await listBundleAssets(zip);                     // 只列资源与其在文档中的位置
await isDocumentBundle(bytes);                   // 判断是否为 bundle，避免误判普通文档

// Node 端也能直接落成一个文件夹（便于人工查看与二次加工）
await writeDocumentBundle(bundle, 'out/技术协议');
```

### DOCX

| 函数 | 返回 |
| --- | --- |
| `docxToJson(input, options?)` | `{ metadata, blocks, messages, text }`，`assets: 'extract'` 时另含 `assets` |
| `docxToHtml(input, options?)` | 上述字段 + `html`（含默认样式表） |
| `docxToText(input)` | `string` |
| `jsonToDocx(doc \| blocks, options?)` | `Uint8Array` |
| `blocksToHtml(doc \| blocks, options?)` | `string`（不生成文件直接渲染） |

支持的块类型：`heading` / `paragraph` / `list` / `table` / `image` / `quote` / `code` / `hr` / `pageBreak`。
文本支持 `bold` `italic` `underline` `strike` `subscript` `superscript` `color` `size` `font` `highlight` `link`。

### XLSX

| 函数 | 返回 |
| --- | --- |
| `xlsxToJson(input, options?)` | `{ metadata, sheets, messages }`，`assets: 'extract'` 时另含 `assets` |
| `xlsxToHtml(input, options?)` | 上述字段 + `html` |
| `xlsxToMatrix(input, sheetName?)` | `{ headers, rows }`，适合直接喂图表库 |
| `jsonToXlsx(workbook \| sheets, options?)` | `Uint8Array` |
| `sheetsToHtml(workbook \| sheets, options?)` | `string` |

单元格支持三种写法，从简到繁：

```javascript
rows: [
  ['简写：直接写值'],
  [{ 1: '对象式：列号或 A1 记号做键' }, { C2: 'C 列' }],
  [{ value: 1200, style: { numFmt: '#,##0.00', align: 'right' } }],
  [{ formula: 'B2*2', result: 2401 }],
  [{ value: '链接', hyperlink: 'https://example.com' }]
]
```

样式支持：`font`（含 `bold` / `color` / `size` / `name`）、`fill`、`align`、`valign`、`wrapText`、`numFmt`、`border`。

### PDF

| 函数 | 返回 |
| --- | --- |
| `pdfToJson(input, options?)` | `{ metadata, pageCount, pages, messages }`，`assets` 选项开启时另含 `assets` |
| `pdfToText(input, options?)` | `string`（按页用空行分隔） |
| `pdfInfo(input)` | `{ metadata, pageCount, pages, messages }`，不提取正文，最快 |
| `jsonToPdf(spec, options?)` | `Uint8Array` |
| `pdfPageToPng(input, page?, scale?)` | `Uint8Array`（PNG，需 `@napi-rs/canvas`，仅 Node） |

支持的元素类型：`text` / `heading` / `image` / `line` / `rect` / `table` / `pageBreak`。
坐标采用**顶部原点**（与 HTML 一致），只需给出距左边距 / 距顶边距的距离。

### PPTX

| 函数 | 返回 |
| --- | --- |
| `pptxToHtml(input, options?)` | 幻灯片 HTML |
| `pptxToJson(input, options?)` | 结构化解析结果 |
| `pptxToStandard(input, options?)` | 标准文档结构（与 `jsonToPptx` 同源，可往返） |
| `pptxToFiles(input)` | 内部文件清单与内容 |
| `jsonToPptx(doc, options?)` | `Uint8Array` |
| `PPTXComposer` | 链式构建，`save()` 返回 `Uint8Array` |
| `editPptx(input)` | 增删页 / 改元数据 |

## 环境要求

| 场景 | 说明 |
| --- | --- |
| Node.js >= 18 | 直接可用，PDF 解析无需额外配置 |
| 浏览器 | 需使用打包器（Vite / Webpack / Rollup）；纯 ESM CDN 直接引入时，`exceljs` 需改用其 UMD 产物 |
| PDF 解析 | `pdfjs-dist` 为可选依赖，未安装时仅 `pdfTo*` 报错，其余功能不受影响：`npm i pdfjs-dist` |
| PDF 页面渲染 | Node 端另需 `npm i @napi-rs/canvas` |
| 生成含中文的 PDF | 需内嵌中文 TTF/OTF（PDF 标准字体不含 CJK 字形） |

```javascript
// 生成中文 PDF：把字体字节交给 fonts.regular
import fs from 'node:fs';
const bytes = await jsonToPdf({
  fonts: { regular: fs.readFileSync('./SourceHanSans-Regular.ttf') },
  elements: [{ type: 'text', text: '中文内容' }]
});
```

## 已知限制

- **PDF 只做解析文本与页面信息**，不做版面分析、表格结构化或图片提取。
- **DOCX 不还原分栏、页眉页脚、脚注、批注与修订**；文本框、SmartArt 等浮动对象会被忽略。
- **XLSX 不保留数据透视表、图表、条件格式与批注**（批注需显式开启 `includeNotes`）。
- **PPTX 由 `@fefeding/ppt-parser` 提供**，其保真度以该库为准（形状、图表、SmartArt 渲染与往返）。

## 开发

```bash
npm install
npm run test:run      # 运行测试
npm run type-check    # 类型检查
npm run build         # 构建 ESM / CJS / browser + .d.ts
```

## License

MIT

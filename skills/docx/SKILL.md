---
name: jdoc-docx
description: 解析、生成与往返编辑 DOCX 文档（@fefeding/jdoc）。当用户需要把 docx 转 HTML/JSON、从块模型生成 docx、提取文本、或在解析结果上做结构化编辑后再生成时使用。
---

# DOCX 文档技能（@fefeding/jdoc）

统一库 `@fefeding/jdoc` 的 DOCX 能力。解析得到「块模型（DocxBlock）」或 HTML，生成接收块模型返回 `Uint8Array`。

## 何时使用
- 把 `.docx` 转成可预览的 HTML / 结构化 JSON
- 用程序生成 Word 文档（报告、合同、导出）
- 在解析出的块上做增删改，再生成新 docx（往返编辑）

## 核心 API
| 函数 | 返回 |
| --- | --- |
| `docxToJson(input, options?)` | `{ metadata, blocks, messages, text }`，`assets:'extract'` 时含 `assets` |
| `docxToHtml(input, options?)` | 上述字段 + `html`（含默认样式表） |
| `docxToText(input)` | `string` |
| `jsonToDocx(doc \| blocks, options?)` | `Uint8Array` |
| `blocksToHtml(doc \| blocks, options?)` | `string`（不落文件直接渲染） |

入参 `input` 支持 `ArrayBuffer` / `TypedArray` / `Buffer` / `Blob` / `File` / `base64` / `data URL`，无需预转换。

## 典型用法
```js
import { docxToJson, jsonToDocx } from '@fefeding/jdoc';

// 解析
const { blocks, text } = await docxToJson(file);

// 生成
const bytes = await jsonToDocx([
  { type: 'heading', level: 1, text: '周报' },
  { type: 'paragraph', text: '本周完成 A、B 两个模块。' },
  { type: 'table', header: true, rows: [['模块', '进度'], ['A', '100%']] }
]);
```

## 块与样式
- 块类型：`heading` / `paragraph` / `list` / `table` / `image` / `quote` / `code` / `hr` / `pageBreak`
- 文本样式（`runs`）：`bold` `italic` `underline` `strike` `subscript` `superscript` `color` `size` `font` `highlight` `link`

## 注意事项
- 图片默认内联为 data URL；想拿到独立字节用 `docxToJson(file, { assets: 'extract' })`，图片按 `blocks[i]` 路径锚定。
- 不还原分栏、页眉页脚、脚注、批注、修订；文本框 / SmartArt 等浮动对象被跳过。
- 生成 DOCX 本身不需要嵌字体（由 Word 渲染）；仅生成 PDF 时才需嵌 CJK 字体。

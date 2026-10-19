---
name: jdoc-convert
description: 任意文档格式互转（@fefeding/jdoc）。当用户需要在 docx / pdf / markdown / html 之间一站式转换、或想把某格式转成中间「块模型」再继续加工时使用。
---

# 文档互转 技能（@fefeding/jdoc）

以「文档块模型（DocxBlock）」为中间表示，提供任意格式间的转换：

```
md / docx / html / pdf  ->  blocks  ->  md / docx / html / pdf
```

## 何时使用
- 一条命令完成格式转换（如 docx→pdf、pdf→md、md→html）
- 把某格式转成 blocks 后做结构化增删改，再生成目标格式
- 需要自动识别源格式（无需显式指定 `from`）

## 核心 API
| 函数 | 说明 |
| --- | --- |
| `convert(input, target, options?)` | 一站式：`target` ∈ `docx`/`pdf`/`md`/`html` |
| `markdownToBlocks` / `blocksToMarkdown` | Markdown ↔ blocks |
| `htmlToBlocks` | HTML → blocks |
| `pdfToBlocks` | PDF → blocks |
| `blocksToPdf` | blocks → PDF |

## 典型用法
```js
import { convert } from '@fefeding/jdoc';

const pdf  = await convert(docxBytes, 'pdf');     // docx -> pdf
const md   = await convert(pdfBytes,  'md');       // pdf  -> markdown
const html = await convert(markdownBytes, 'html'); // md   -> html
```

## 选项
- `from`：源格式（`'docx'\|'pdf'\|'md'\|'html'`），缺省自动识别；HTML 需显式指定。
- `fonts`：内嵌字体（生成 PDF 且含中文必须）。
- `autoFont`：Node 端自动探测系统字体（默认 true）。
- `quality`：PDF 解析质量 `'fast'\|'layout'`（默认 `'layout'`）。
- `assets`：资源表，用于还原图片 assetId。

## 注意事项
- 文本目标（md/html）返回 `string`；二进制目标（docx/pdf）返回 `Uint8Array`。
- PDF→blocks 走解析路径，对上百页文档有一定开销；纯文本类（md/html/txt）按原始文本处理。
- 复杂版式（多栏、浮动对象）在互转中会被简化，见各格式已知限制。

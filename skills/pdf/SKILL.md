---
name: jdoc-pdf
description: 解析、生成与编辑 PDF（@fefeding/jdoc）。当用户需要提取 PDF 文本/信息、从块模型生成 PDF、把 PDF 栅格化为 PNG、或在既有 PDF 上叠加签名/图片/文本/图形时使用。
---

# PDF 技能（@fefeding/jdoc）

## 何时使用
- 解析 PDF 文本与页面信息（无需渲染）
- 从块模型生成 PDF 文档
- 把某一页渲染成 PNG（Node，需 `@napi-rs/canvas`）
- 在既有 PDF 上盖章/签名/叠加（不破坏原内容）

## 核心 API
| 函数 | 返回 |
| --- | --- |
| `pdfToJson(input, options?)` | `{ metadata, pageCount, pages, messages }` |
| `pdfToText(input, options?)` | `string`（按页空行分隔） |
| `pdfInfo(input)` | 同上但不提取正文，最快 |
| `jsonToPdf(spec, options?)` | `Uint8Array` |
| `pdfPageToPng(input, page?, scale?)` | `Uint8Array`（PNG，Node） |
| `editPdf(input, spec)` | 追加叠加并保存 |
| `mergePdfs(inputs)` | 合并 |
| `splitPdf(input, ranges)` | 按页区间拆分 |

## 编辑既有 PDF（扁平化叠加）
```js
const bytes = await editPdf(file, {
  overlays: [
    [{ type: 'rect', x: 40, y: 60, width: 200, height: 18, fill: '#ffe08a', opacity: 0.5 }],
    [{ type: 'text', x: 40, y: 120, text: 'Approved', size: 14, color: '#334155' }]
  ],
  rotate: [{ page: 0, degrees: 90 }],
  removePages: [4],
  pageOrder: [2, 0, 1]
});
```

## 注意事项
- 坐标取**顶部原点**（同 HTML），给距左边距/顶边距的偏移即可。
- **含中文的 PDF 必须内嵌 CJK TTF/OTF**：PDF 标准 14 字体无中文字形。
- `editPdf` 是扁平化叠加（画进页面内容流，之后不可单独移除），但保留原矢量/字体/表单域。
- PDF 解析只覆盖文本与页面信息，无版面分析/表格结构化/图片提取。

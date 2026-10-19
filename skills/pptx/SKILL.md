---
name: jdoc-pptx
description: 解析、生成与编辑 PPTX 演示文稿（@fefeding/jdoc，底层 @fefeding/ppt-parser）。当用户需要把 pptx 渲染为 HTML/结构化 JSON、生成幻灯片、或增删页/改元数据时。
---

# PPTX 演示文稿技能（@fefeding/jdoc）

PPTX 能力由 `@fefeding/ppt-parser` 提供，jdoc 负责统一入参/出参类型。

## 何时使用
- 把 `.pptx` 渲染成可预览的幻灯片 HTML
- 生成演示文稿（程序化排版）
- 链式构建或编辑（增删页、改元数据）

## 核心 API
| 函数 | 返回 |
| --- | --- |
| `pptxToHtml(input, options?)` | 幻灯片 HTML |
| `pptxToJson(input, options?)` | 结构化解析结果 |
| `pptxToStandard(input, options?)` | 标准文档结构（与 `jsonToPptx` 同源，可往返） |
| `pptxToFiles(input)` | 内部文件清单与内容 |
| `jsonToPptx(doc, options?)` | `Uint8Array` |
| `PPTXComposer` | 链式构建，`save()` 返回 `Uint8Array` |
| `editPptx(input)` | 增删页 / 改元数据 |

## 典型用法
```js
import { jsonToPptx, PPTXComposer } from '@fefeding/jdoc';

const bytes = await jsonToPptx({
  slides: [{ elements: [{ type: 'text', text: 'Hello', x: 50, y: 50, width: 200, height: 40 }] }]
});

// 链式
const pptx = new PPTXComposer();
pptx.addSlide().addText('Title', { x: 50, y: 50 });
const out = await pptx.save();
```

## 注意事项
- 保真度以 `@fefeding/ppt-parser` 为准：形状、图表、SmartArt 的渲染与往返取决于该库。
- 图片默认 `both`（data URL + assetId）；`extract` 模式元素上只留 `assetId`，生成时自动还原。

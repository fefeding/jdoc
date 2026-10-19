---
name: jdoc-markdown
description: Markdown / 纯文本 的解析与渲染（@fefeding/jdoc）。当用户需要把 Markdown 转 HTML（含 Mermaid/KaTeX 支持）、提取纯文本、或安全渲染用户 Markdown 时使用。
---

# Markdown / 纯文本 技能（@fefeding/jdoc）

## 何时使用
- 把 Markdown 渲染成带样式的 HTML（博客、文档站点、预览）
- 提取 Markdown 的纯文本或标题大纲（`metadata.headings`）
- 纯文本转义后包 `<pre>` 预览
- 支持 Mermaid 图表与 KaTeX 公式渲染

## 核心 API
| 函数 | 返回 |
| --- | --- |
| `markdownToHtml(input, options?)` | `{ html, text, metadata, messages }` |
| `markdownToText(input, options?)` | 去标记纯文本 |
| `textToHtml(input, options?)` | `{ html, text, messages }` |
| `textToText(input)` | 原文（UTF-8 解码，去 BOM） |
| `renderMarkdownExtras(container)` | 浏览器端渲染 Mermaid/KaTeX，返回 `{ math, diagrams }` |

```js
const { html, metadata } = await markdownToHtml(bytes);
container.innerHTML = html;             // 自带 .jdoc-md 作用域样式
await renderMarkdownExtras(container);  // 动态 import Mermaid/KaTeX，不渲染不加载
```

## 注意事项
- 默认允许 Markdown 内嵌原始 HTML，但移除 `script` / `iframe` / 事件属性与 `javascript:` 等危险协议；`{ allowHtml: false }` 让原始 HTML 全部按纯文本转义。
- 公式/流程图语法可整体关闭：`{ math: false, diagrams: false }`。
- 渲染失败会回退到显示源码，不会白屏。

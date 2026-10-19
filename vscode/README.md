# jdoc Document Viewer

VS Code 扩展，基于 [`@fefeding/jdoc`](https://github.com/fefeding/jdoc) 在编辑器内**查看并编辑**办公文档：

- **PPTX** — 幻灯片渲染（含图表），可编辑
- **DOCX** — 文档渲染，可编辑
- **XLSX** — 表格渲染，可编辑
- **PDF** — 页面栅格化渲染（只读，支持扫描件）
- **Markdown**（`.md` / `.markdown`）— 排版预览（只读，支持 Mermaid 图表与 KaTeX 公式）
- **纯文本**（`.txt`）— 等宽预览（只读）

打开 PPTX / DOCX / XLSX / PDF 会自动以本扩展打开（作为默认编辑器）。
Markdown / TXT 因 VS Code 内置编辑器优先级更高，需在文件上右键选择 **“jdoc: 用 jdoc 打开预览”**（命令 `jdoc.preview`），或经“打开方式”选择 `jdoc Document Editor`。

支持编辑后保存到原文件（`Ctrl/Cmd+S`）、`Revert`（还原）、`Export Copy`（另存副本）。

## Markdown 预览能力

- **Mermaid**：` ```mermaid ` 围栏渲染为 SVG，支持 flowchart / sequence / class / state / ER / gantt / pie 等图
- **KaTeX**：支持 `$行内$`、`$$块级$$`、`\[…\]`、`\(…\)`，以及 ` ```math ` / ` ```tex ` 围栏
- 语法错误或依赖未加载时保留源码显示，不会白屏

> 注意：普通 `$5 和 $6` 这类货币写法不会被当成公式（要求 `$` 内侧不紧跟空白）。

## 开发

```bash
# 仓库根目录：先构建 jdoc 库（产出 dist/jdoc.browser.js）
npm install && npm run build

# 扩展目录
cd vscode
npm install
npm run dev        # 持续构建 extension.js + webview.js

# 在 VS Code 中打开仓库根，按 F5 选择 "Run Extension (watch)"
```

## 发布

推送 `v*` 标签或手动触发 `Publish Extension` 工作流，自动构建并发布到
**VS Code Marketplace** 与 **Open VSX**。需要仓库 Secrets：

- `VSCE_PAT` — Marketplace 发布令牌
- `OVSX_TOKEN` — Open VSX 访问令牌

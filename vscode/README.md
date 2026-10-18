# jdoc Document Viewer

VS Code 扩展，基于 [`@fefeding/jdoc`](https://github.com/fefeding/jdoc) 在编辑器内**查看并编辑**办公文档：

- **PPTX** — 幻灯片渲染（含图表），可编辑
- **DOCX** — 文档渲染，可编辑
- **XLSX** — 表格渲染，可编辑
- **PDF** — 页面渲染（只读）

打开上述任意文件即自动以本扩展打开（作为默认编辑器）。支持编辑后保存到原文件（`Ctrl/Cmd+S`）、`Revert`（还原）、`Export Copy`（另存副本）。

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

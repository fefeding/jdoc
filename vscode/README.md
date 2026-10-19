# jdoc Document Viewer

VS Code 扩展，基于 [`@fefeding/jdoc`](https://github.com/fefeding/jdoc) 在编辑器内**查看并编辑**办公文档：

- **PPTX** — 幻灯片渲染（含图表），可编辑
- **DOCX** — 文档渲染，可编辑
- **XLSX** — 表格渲染，可编辑
- **PDF** — 页面栅格化渲染（支持扫描件），可叠加编辑：签名 / 图片 / 文本 / 矩形 / 手绘
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

## PDF 叠加编辑

工具栏铅笔按钮（或标题栏「切换到编辑」）进入编辑模式，多出一行叠加工具栏：

| 工具 | 快捷键 | 说明 |
| --- | --- | --- |
| 选择 | `V` | 选中后可拖拽移动，8 向手柄缩放；`Delete` / `×` 删除，`Esc` 取消 |
| 文本 | `T` | 点击落点后输入内容，支持字号、文字颜色；嵌入中文需宿主提供 TTF |
| 图片 | `I` | 选图片文件（PNG / JPEG）后插入左上角，拖拽缩放定位 |
| 矩形 | `R` | 点击落一个高亮块 / 遮挡块（填充 + 描边 + 半透明） |
| 手绘 | `D` | 按下拖出一笔，用于手写签名 |

保存（`Ctrl/Cmd+S` 或工具栏「保存」）把叠加内容写进 PDF 并回写原文件。

叠加是**扁平化**的：内容被追加到页面内容流，之后无法再单独撤回。它不会改动原有正文——矢量
路径、字体子集、表单域等一律原样保留。页面管理与合并拆分见库层 `editPdf` / `mergePdfs` /
`splitPdf`。

## PDF 全文搜索

工具栏放大镜按钮打开搜索栏（独立一行，不挤占导航栏）。打开文档时后台预取每页文本项，
因此搜索无需再解析一遍、也不依赖当前是否渲染到该页。

- 输入即搜（防抖 200ms），命中以黄色高亮框标出，当前命中为橙色
- `Enter` 跳下一个、`Shift+Enter` 跳上一个，或用栏内 ↑/↓ 按钮；命中自动翻页并滚动到视口中央
- `Esc` 关闭搜索栏并清除高亮
- 无匹配时输入框标红并显示 `0/0`，不会崩
- 坐标复用文本层同一套换算（含旋转），高亮矩形为轴对齐近似（与文本层一致）
- 词边界用相邻文本项的水平间隙近似补空格，跨词的短语（如 `by JDOC`）会精确高亮到对应的若干段文本项

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

# jdoc Document Viewer

A VS Code extension that lets you **view and edit** office documents right inside the editor, built on [`@fefeding/jdoc`](https://github.com/fefeding/jdoc):

- **PPTX** — slide rendering (including charts), editable
- **DOCX** — document rendering (paginated by paper size), editable
- **XLSX** — spreadsheet rendering in an Excel-like grid (column/row headers, real column widths and row heights, cell borders / fills / fonts / number formats, merged cells, frozen panes), with **sheet tabs**, an in-cell **formula bar**, a status bar showing Sum / Count / Average for the selection, and **charts** (column / bar / pie / doughnut / line / area / scatter / radar) anchored to cells and refreshed live from the referenced cells; formulas (including cross-sheet references and Chinese sheet names) are evaluated incrementally when a cell is edited. Charts survive saving, and cells are editable. Chart drawing is shared with the PPTX editor through the library's `chart/svg` module, so both formats render the same chart types with the same legend / data-label / number-format rules.
- **PDF** — rasterized page rendering (scanned documents supported), with overlay editing: signature / image / text / rectangle / freehand; large documents show a loading indicator first and prioritize the first page, while the rest of the pages and the thumbnails are filled in incrementally in the background
- **Markdown** (`.md` / `.markdown`) — typeset preview (read-only, fenced code blocks with syntax highlighting, Mermaid diagrams and KaTeX formulas supported)
- **Plain text** (`.txt`) — monospace preview (read-only)

Opening a PPTX / DOCX / XLSX / PDF automatically uses this extension (as the default editor). Because VS Code's built-in editor has higher priority for Markdown / TXT, you need to right-click the file and choose **"jdoc: Open Preview with jdoc"** (command `jdoc.preview`), or pick `jdoc Document Editor` via "Open With".

After editing you can save back to the original file (`Ctrl/Cmd+S`), `Revert` (restore), or `Export Copy` (save a copy).

## Markdown Preview Features

- **Mermaid**: ` ```mermaid ` fenced blocks render as SVG, supporting flowchart / sequence / class / state / ER / gantt / pie diagrams
- **KaTeX**: supports `$inline$`, `$$block$$`, `\[…\]`, `\(…\)`, as well as ` ```math ` / ` ```tex ` fences
- Syntax errors or unloaded dependencies keep the source visible instead of blanking out

> Note: ordinary currency writing like `$5 and $6` is not treated as a formula (the character inside `$` must not be followed immediately by whitespace).

## PDF Overlay Editing

Click the pencil button on the toolbar (or "Switch to Edit" in the title bar) to enter edit mode, which adds a row of overlay tools:

| Tool | Shortcut | Description |
| --- | --- | --- |
| Select | `V` | After selecting, drag to move; 8-direction handles to resize; `Delete` / `×` to delete, `Esc` to cancel |
| Text | `T` | Click to place, then type; font size and text color supported; embedding Chinese requires a TTF font from the host |
| Image | `I` | Choose an image file (PNG / JPEG), inserted at the top-left; drag and resize to position |
| Rectangle | `R` | Click to drop a highlight / cover block (fill + stroke + semi-transparent) |
| Freehand | `D` | Press and drag to draw a stroke, for handwritten signatures |

Save (`Ctrl/Cmd+S` or the toolbar "Save") writes the overlay content into the PDF and back to the original file.

The overlay is **flattened**: content is appended to the page content stream and cannot be individually undone afterwards. It does not modify the original body text — vector paths, font subsets, form fields, etc. are all preserved as-is. For page management and merge/split, see the library-level `editPdf` / `mergePdfs` / `splitPdf`.

## DOCX Pagination and Typesetting Fidelity

DOCX paginates by paper height in view mode, trying to match the look of Word/WPS:

- **Line grid**: almost all Chinese Word documents enable `w:docGrid` (line grid); paragraphs without explicit line spacing align to the grid — without restoring this, each page fits a few extra lines and pagination shifts forward overall
- **Line spacing multiplier**: Word's "multiple" line spacing is a multiple of single-line spacing (about 1.2× the font size), not a multiple of the font size itself; using it directly as CSS `line-height` makes it about 20% too small
- **Keep with next**: headings won't be isolated at the bottom of a page (Word's keepNext)
- **Widow/orphan and cross-page control**: paragraphs can split across page boundaries at line boundaries, and at least 2 lines are kept at both page bottom and top (widow/orphan)
- **List numbering and indentation**: restores the numbering format of `numbering` (`a. b. c.` / `i. ii. iii.`) and the `w:ind` indentation

Line-level splitting is enabled only in **view mode**: in edit mode, splitting blocks would break the "block ↔ data model" correspondence, making contenteditable cursor positioning, selection, and undo inaccurate, so edit mode keeps whole-block movement.

## PDF Full-Text Search

The magnifier button on the toolbar opens the search bar (a separate row, not crowding the navigation bar). When a document is opened, the text items of each page are prefetched in the background, so searching requires no re-parsing and does not depend on whether the page is currently rendered.

- Search as you type (200ms debounce); matches are highlighted with yellow boxes, the current match in orange
- `Enter` jumps to next, `Shift+Enter` to previous, or use the ↑/↓ buttons in the bar; matches auto-flip pages and scroll to the viewport center
- `Esc` closes the search bar and clears highlights
- No match shows the input box in red with `0/0`, no crash
- Coordinates reuse the same transform as the text layer (including rotation); highlight rectangles are axis-aligned approximations (consistent with the text layer)
- Word boundaries are approximated with the horizontal gaps between adjacent text items to fill spaces; phrases spanning words (e.g. `by JDOC`) precisely highlight the corresponding text items

## Skills

The [`skills/`](../../skills) directory at the repository root provides capability-split agent skill packages (one `SKILL.md` each) for `@fefeding/jdoc`, loaded by AI assistants on demand: covering docx / xlsx / pptx / pdf / markdown / convert / assets / bundle, eight categories. The extension itself shares the same jdoc API with these skills.

## Development

```bash
# Repository root: build the jdoc library first (produces dist/jdoc.browser.js)
npm install && npm run build

# Extension directory
cd vscode
npm install
npm run dev        # continuously build extension.js + webview.js

# Open the repository root in VS Code, press F5 and choose "Run Extension (watch)"
```

## Publishing

Pushing a `v*` tag or manually triggering the `Publish Extension` workflow builds and publishes to **VS Code Marketplace** and **Open VSX**. Required repository Secrets:

- `VSCE_PAT` — Marketplace publish token
- `OVSX_TOKEN` — Open VSX access token

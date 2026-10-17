/**
 * DOCX 结构 -> HTML
 *
 * 输出的 HTML 自带内联样式，可直接插入页面预览；调用方也可只取 blocks 自行渲染。
 *
 * @module docx/to-html
 */

import type {
    DocxBlock,
    DocxCell,
    DocxDocument,
    DocxListItem,
    DocxRun,
    DocxTable
} from './types';

/** HTML 转义 */
export function escapeHtml(value: string): string {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

/** 转义并保留换行（run 内的软换行与制表符） */
function escapeWithBreaks(value: string): string {
    return escapeHtml(value).replace(/\n/g, '<br/>').replace(/\t/g, '&emsp;');
}

/** run -> 内联 span / a */
export function runToHtml(run: DocxRun): string {
    const styles: string[] = [];
    if (run.color) styles.push(`color:${run.color}`);
    if (run.size !== undefined) styles.push(`font-size:${run.size}pt`);
    if (run.font) styles.push(`font-family:${escapeHtml(run.font)}`);
    if (run.highlight) styles.push(`background-color:${run.highlight}`);

    let text = escapeWithBreaks(run.text);
    if (run.bold) text = `<strong>${text}</strong>`;
    if (run.italic) text = `<em>${text}</em>`;
    if (run.underline) text = `<u>${text}</u>`;
    if (run.strike) text = `<s>${text}</s>`;
    if (run.subscript) text = `<sub>${text}</sub>`;
    if (run.superscript) text = `<sup>${text}</sup>`;
    if (run.link) {
        text = `<a href="${escapeHtml(run.link)}" target="_blank" rel="noreferrer noopener">${text}</a>`;
    }
    if (styles.length) return `<span style="${escapeHtml(styles.join(';'))}">${text}</span>`;
    return text;
}

/** runs 或 text -> HTML */
function contentToHtml(text: string | undefined, runs: DocxRun[] | undefined): string {
    if (runs && runs.length) return runs.map(runToHtml).join('');
    return escapeWithBreaks(text || '');
}

/** 列表项 -> HTML（含嵌套子列表） */
function listItemToHtml(item: DocxListItem, ordered: boolean): string {
    const children = typeof item === 'string' ? [] : ((item as { children?: DocxListItem[] }).children || []);
    const self = typeof item === 'string'
        ? escapeWithBreaks(item)
        : (item as DocxRun[]).map(runToHtml).join('');

    if (!children.length) return `<li>${self}</li>`;
    const tag = ordered ? 'ol' : 'ul';
    return `<li>${self}<${tag}>${children.map((child) => listItemToHtml(child, ordered)).join('')}</${tag}></li>`;
}

/** 单元格 -> td/th */
function cellToHtml(cell: DocxCell, isHeader: boolean): string {
    const tag = isHeader ? 'th' : 'td';
    const content = typeof cell === 'string'
        ? escapeWithBreaks(cell)
        : (cell as DocxRun[]).map(runToHtml).join('');
    return `<${tag} style="border:1px solid #ccc;padding:4px 6px">${content || '&nbsp;'}</${tag}>`;
}

/** 表格 -> HTML */
function tableToHtml(block: DocxTable): string {
    const hasHeader = block.header !== false && block.rows.length > 0;
    const columnCount = block.rows.reduce((max, row) => Math.max(max, row.length), 0);
    const cellStyle = block.borderSize === 0
        ? ' style="border:none;padding:4px 6px"'
        : ' style="border:1px solid #ccc;padding:4px 6px"';

    const colgroup = block.columnWidths
        ? `<colgroup>${block.columnWidths
              .slice(0, columnCount)
              .map((width) => `<col style="width:${width}pt"/>`)
              .join('')}</colgroup>`
        : '';

    const body = block.rows
        .map((row, rowIndex) => {
            const header = hasHeader && rowIndex === 0;
            const cells = Array.from({ length: columnCount }, (_, i) => cellToHtml(row[i] ?? '', header)).join('');
            return `<tr>${cells}</tr>`;
        })
        .join('');

    const tableStyle = block.borderSize === 0
        ? ' style="border-collapse:collapse;border:none"'
        : ' style="border-collapse:collapse"';

    return `<table${tableStyle}>${colgroup}${body}</table>`.replace(
        /<t[dh](?![^>]*style=)/g,
        `<t$1${cellStyle}`
    );
}

/** 段落级属性 -> inline style */
function paragraphStyle(block: DocxBlock): string {
    const para = block as {
        align?: string;
        indentLeft?: number;
        indentRight?: number;
        indentFirstLine?: number;
        spacingBefore?: number;
        spacingAfter?: number;
        lineSpacing?: number;
    };
    const styles: string[] = [];
    if (para.align) styles.push(`text-align:${para.align}`);
    if (para.indentLeft) styles.push(`margin-left:${para.indentLeft}pt`);
    if (para.indentRight) styles.push(`margin-right:${para.indentRight}pt`);
    if (para.indentFirstLine) {
        styles.push(para.indentFirstLine < 0
            ? `text-indent:-${Math.abs(para.indentFirstLine)}pt`
            : `text-indent:${para.indentFirstLine}pt`);
    }
    if (para.spacingBefore) styles.push(`margin-top:${para.spacingBefore}pt`);
    if (para.spacingAfter) styles.push(`margin-bottom:${para.spacingAfter}pt`);
    if (para.lineSpacing && para.lineSpacing !== 1) styles.push(`line-height:${para.lineSpacing}`);
    return styles.length ? ` style="${escapeHtml(styles.join(';'))}"` : '';
}

/** 单块 -> HTML */
export function blockToHtml(block: DocxBlock): string {
    switch (block.type) {
        case 'heading': {
            const level = block.level >= 1 && block.level <= 6 ? block.level : 3;
            return `<h${level}${paragraphStyle(block)}>${contentToHtml(block.text, block.runs)}</h${level}>`;
        }
        case 'paragraph':
            return `<p${paragraphStyle(block)}>${contentToHtml(block.text, block.runs) || '&nbsp;'}</p>`;
        case 'quote':
            return `<blockquote${paragraphStyle(block)}>${contentToHtml(block.text, block.runs)}</blockquote>`;
        case 'code':
            return `<pre style="background:#f5f5f5;padding:8px;border-radius:4px;overflow:auto"><code>${escapeHtml(block.text)}</code></pre>`;
        case 'list': {
            const tag = block.ordered ? 'ol' : 'ul';
            return `<${tag}>${block.items.map((item) => listItemToHtml(item, block.ordered)).join('')}</${tag}>`;
        }
        case 'table':
            return tableToHtml(block);
        case 'image': {
            if (typeof block.src !== 'string') return '';
            const attrs = [
                block.width ? `width="${Math.round(block.width * (96 / 72))}"` : '',
                block.height ? `height="${Math.round(block.height * (96 / 72))}"` : '',
                block.alt ? `alt="${escapeHtml(block.alt)}"` : 'alt=""'
            ]
                .filter(Boolean)
                .join(' ');
            const img = `<img src="${escapeHtml(block.src)}" ${attrs} style="max-width:100%"/>`;
            const caption = block.caption
                ? `<figcaption style="text-align:center;font-size:11px;color:#666">${escapeHtml(block.caption)}</figcaption>`
                : '';
            return `<figure style="margin:12px 0;text-align:center">${img}${caption}</figure>`;
        }
        case 'pageBreak':
            return '<hr style="page-break-after:always;border:none;border-top:1px dashed #ccc"/>';
        case 'hr':
            return '<hr/>';
        default:
            return '';
    }
}

/** 默认样式表（便于开箱即用，也可忽略自带样式自行渲染） */
export const DOCX_HTML_STYLE = `.docx-body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,"PingFang SC","Microsoft YaHei",sans-serif;font-size:11pt;line-height:1.6;color:#1f2328}
.docx-body h1{font-size:20pt;margin:20px 0 10px}
.docx-body h2{font-size:17pt;margin:18px 0 9px}
.docx-body h3{font-size:15pt;margin:16px 0 8px}
.docx-body h4{font-size:13pt;margin:14px 0 7px}
.docx-body h5,.docx-body h6{font-size:12pt;margin:12px 0 6px}
.docx-body p{margin:6px 0}
.docx-body blockquote{margin:8px 0;padding:4px 12px;border-left:3px solid #d0d7de;color:#57606a}
.docx-body table{margin:10px 0;font-size:10.5pt}
.docx-body th{background:#f6f8fa;font-weight:600}
.docx-body img{max-width:100%}
`;

/**
 * 把 DOCX 结构渲染为 HTML
 * @param input - DocxDocument 或 blocks 数组
 * @param options.includeStyle - 是否附带默认样式表（默认 true）
 */
export function docxToHtmlString(
    input: DocxDocument | DocxBlock[],
    options: { includeStyle?: boolean } = {}
): string {
    const blocks = Array.isArray(input) ? input : input?.blocks || [];
    const body = blocks.map(blockToHtml).filter(Boolean).join('\n');
    const style = options.includeStyle === false ? '' : `<style>${DOCX_HTML_STYLE}</style>`;
    return `${style}<div class="docx-body">${body}</div>`;
}

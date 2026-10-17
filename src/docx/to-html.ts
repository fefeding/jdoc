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
    DocxCellBorder,
    DocxCellDetail,
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

/** 单元格边框单边 -> CSS 声明 */
function borderSideToCss(side: string, border: DocxCellBorder | undefined, fallback: string): string {
    if (border && border.style === 'nil') return `${side}:none`;
    if (border) {
        const width = border.size !== undefined ? border.size : 1;
        const lineStyle = border.style === 'double' ? 'double' : border.style === 'dashed' ? 'dashed' : border.style === 'dotted' ? 'dotted' : 'solid';
        return `${side}:${width}pt ${lineStyle} ${border.color || '#000000'}`;
    }
    return `${side}:${fallback}`;
}

/** 单元格内容 -> HTML（字符串 / 富文本 / 多段落对象） */
function cellContentToHtml(cell: DocxCell): string {
    if (typeof cell === 'string') return escapeWithBreaks(cell);
    if (Array.isArray(cell)) return cell.map(runToHtml).join('');
    const detail = cell as DocxCellDetail;
    if (detail.blocks && detail.blocks.length) return detail.blocks.map(blockToHtml).join('');
    if (detail.runs && detail.runs.length) return detail.runs.map(runToHtml).join('');
    return escapeWithBreaks(detail.text || '');
}

/** 单元格 -> td/th（含合并、边框、底纹、对齐、列宽） */
function cellToHtml(cell: DocxCell, isHeader: boolean, defaultBorder: string): string {
    const tag = isHeader ? 'th' : 'td';
    const content = cellContentToHtml(cell);
    const styles: string[] = [];

    const detail = typeof cell === 'object' && cell !== null && !Array.isArray(cell)
        ? (cell as DocxCellDetail)
        : undefined;

    if (detail) {
        const b = detail.borders;
        styles.push(
            borderSideToCss('border-top', b?.top, defaultBorder),
            borderSideToCss('border-right', b?.right, defaultBorder),
            borderSideToCss('border-bottom', b?.bottom, defaultBorder),
            borderSideToCss('border-left', b?.left, defaultBorder)
        );
        if (detail.fill) styles.push(`background-color:${detail.fill}`);
        if (detail.verticalAlign) styles.push(`vertical-align:${detail.verticalAlign}`);
        if (detail.width) styles.push(`width:${detail.width}pt`);
    } else {
        styles.push(
            borderSideToCss('border-top', undefined, defaultBorder),
            borderSideToCss('border-right', undefined, defaultBorder),
            borderSideToCss('border-bottom', undefined, defaultBorder),
            borderSideToCss('border-left', undefined, defaultBorder)
        );
    }
    styles.push('padding:4px 6px');

    const attrs = [`style="${escapeHtml(styles.join(';'))}"`];
    if (detail?.colspan && detail.colspan > 1) attrs.push(`colspan="${detail.colspan}"`);
    if (detail?.rowspan && detail.rowspan > 1) attrs.push(`rowspan="${detail.rowspan}"`);

    return `<${tag} ${attrs.join(' ')}>${content || '&nbsp;'}</${tag}>`;
}

/** 表格 -> HTML */
function tableToHtml(block: DocxTable): string {
    const hasHeader = block.header !== false && block.rows.length > 0;
    const columnCount = block.rows.reduce((max, row) => Math.max(max, row.length), 0);
    const defaultBorder = block.borderSize === 0 ? 'none' : '1pt solid #cccccc';

    const colgroup = block.columnWidths
        ? `<colgroup>${block.columnWidths
              .slice(0, columnCount)
              .map((width) => `<col style="width:${width}pt"/>`)
              .join('')}</colgroup>`
        : '';

    const body = block.rows
        .map((row, rowIndex) => {
            const header = hasHeader && rowIndex === 0;
            const cells = row.map((cell) => cellToHtml(cell ?? '', header, defaultBorder)).join('');
            return `<tr>${cells}</tr>`;
        })
        .join('');

    const tableStyle = block.borderSize === 0
        ? ' style="border-collapse:collapse;border:none"'
        : ' style="border-collapse:collapse"';

    return `<table${tableStyle}>${colgroup}${body}</table>`;
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
        case 'textBox': {
            const styles: string[] = ['margin:8px 0', 'padding:8px 12px', 'overflow:auto'];
            if (block.fill) {
                styles.push(`background-color:${block.fill}`);
                // 深色底自动反白（run 未显式指定颜色时保证可读）
                const hex = block.fill.replace(/^#/, '');
                if (/^[0-9a-fA-F]{6}$/.test(hex)) {
                    const r = parseInt(hex.slice(0, 2), 16);
                    const g = parseInt(hex.slice(2, 4), 16);
                    const b = parseInt(hex.slice(4, 6), 16);
                    if (r * 0.299 + g * 0.587 + b * 0.114 < 128) styles.push('color:#ffffff');
                }
            }
            if (block.stroke) styles.push(`border:1px solid ${block.stroke}`);
            if (block.width) styles.push(`width:${block.width}pt`);
            const inner = block.paragraphs.map(blockToHtml).join('');
            return `<figure class="docx-textbox" style="${escapeHtml(styles.join(';'))}">${inner}</figure>`;
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
.docx-page{background:#fff;margin:24px auto;box-shadow:0 1px 6px rgba(0,0,0,.2);box-sizing:border-box}
.docx-body h1{font-size:20pt;margin:20px 0 10px}
.docx-body h2{font-size:17pt;margin:18px 0 9px}
.docx-body h3{font-size:15pt;margin:16px 0 8px}
.docx-body h4{font-size:13pt;margin:14px 0 7px}
.docx-body h5,.docx-body h6{font-size:12pt;margin:12px 0 6px}
.docx-body p{margin:0}
.docx-body blockquote{margin:8px 0;padding:4px 12px;border-left:3px solid #d0d7de;color:#57606a}
.docx-body table{margin:10px 0;font-size:10.5pt}
.docx-body th{background:#f6f8fa;font-weight:600}
.docx-body img{max-width:100%}
.docx-textbox{margin:8px 0;font-size:10pt;line-height:1.45}
.docx-textbox p{margin:0}
`;

/**
 * 把 DOCX 结构渲染为 HTML
 * @param input - DocxDocument 或 blocks 数组（传 DocxDocument 时按 section 还原页面）
 * @param options.includeStyle - 是否附带默认样式表（默认 true）
 */
export function docxToHtmlString(
    input: DocxDocument | DocxBlock[],
    options: { includeStyle?: boolean } = {}
): string {
    const doc = Array.isArray(input) ? undefined : input;
    const blocks = Array.isArray(input) ? input : input?.blocks || [];
    const body = blocks.map(blockToHtml).filter(Boolean).join('\n');

    // 有页面设置时还原纸张：固定页宽 + 页边距，接近 Word 的版面
    let pageOpen = '<div class="docx-body">';
    let pageClose = '</div>';
    const section = doc?.section;
    if (section) {
        pageOpen =
            `<div class="docx-page" style="width:${section.width}pt;min-height:${section.height}pt;` +
            `padding:${section.marginTop}pt ${section.marginRight}pt ${section.marginBottom}pt ${section.marginLeft}pt">` +
            '<div class="docx-body">';
        pageClose = '</div></div>';
    }

    const style = options.includeStyle === false ? '' : `<style>${DOCX_HTML_STYLE}</style>`;
    return `${style}${pageOpen}${body}${pageClose}`;
}

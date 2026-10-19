/**
 * Markdown 增强内容渲染（浏览器端）
 *
 * `markdownToHtml` 只产出语义容器（保留源码作降级显示）；本模块在浏览器里把
 * Mermaid 图表与 KaTeX 公式真正渲染出来。两个依赖都是动态 import：
 * 不渲染时不加载，也不会影响 Node 端使用。
 *
 * 用法：
 * ```js
 * container.innerHTML = html;
 * await renderMarkdownExtras(container);
 * ```
 *
 * @module markdown/extras
 */

/** 把 <pre class="jdoc-mermaid"> 渲染为 SVG 图表 */
async function renderDiagrams(root: ParentNode, theme?: MermaidTheme): Promise<number> {
    const nodes = Array.from(root.querySelectorAll<HTMLElement>('.jdoc-mermaid'));
    if (!nodes.length) return 0;

    const { default: mermaid } = await import('mermaid');
    mermaid.initialize({
        startOnLoad: false,
        securityLevel: 'strict',
        suppressErrorRendering: true,
        theme: theme || 'default'
    });

    let done = 0;
    let seq = 0;
    for (const el of nodes) {
        const code = el.querySelector('code')?.textContent ?? el.textContent ?? '';
        if (!code.trim()) continue;
        try {
            const { svg } = await mermaid.render(`jdoc-mermaid-${Date.now()}-${seq++}`, code);
            const holder = document.createElement('div');
            holder.className = 'jdoc-mermaid-svg';
            holder.innerHTML = svg;
            // SVG 内已自带宽高，显式让容器可横向滚动
            el.replaceWith(holder);
            done += 1;
        } catch {
            // 语法错误时保留源码（容器仍是 <pre>，用户能看见原始文本）
        }
    }
    return done;
}

/** 把 [data-tex] 容器渲染为 KaTeX 公式 */
async function renderMath(root: ParentNode): Promise<number> {
    const nodes = Array.from(root.querySelectorAll<HTMLElement>('[data-tex]'));
    if (!nodes.length) return 0;

    const { default: katex } = await import('katex');
    let done = 0;
    for (const el of nodes) {
        const tex = el.dataset.tex ?? el.textContent ?? '';
        const display = el.classList.contains('jdoc-math-block');
        try {
            el.innerHTML = katex.renderToString(tex, {
                displayMode: display,
                throwOnError: false,
                strict: false
            });
            delete el.dataset.tex;
            done += 1;
        } catch {
            // 公式语法错误：保留 TeX 源码
        }
    }
    return done;
}

/** Mermaid 内置主题 */
export type MermaidTheme = 'default' | 'base' | 'forest' | 'dark' | 'neutral' | 'neo';

/** renderMarkdownExtras 的可选项 */
export interface MarkdownExtrasOptions {
    /** 是否渲染 Mermaid 图表，默认 true */
    diagrams?: boolean;
    /** 是否渲染 KaTeX 公式，默认 true */
    math?: boolean;
    /** Mermaid 主题，默认 default */
    theme?: MermaidTheme;
}

/** renderMarkdownExtras 的返回结构 */
export interface MarkdownExtrasResult {
    /** 成功渲染的图表数 */
    diagrams: number;
    /** 成功渲染的公式数 */
    math: number;
}

/**
 * 把 Markdown 预览里的 Mermaid 图表与 KaTeX 公式渲染出来
 * @param root - 预览容器（含 `.jdoc-mermaid` / `[data-tex]` 节点）
 * @param options - 渲染选项
 *
 * @example
 * ```js
 * await renderMarkdownExtras(document.querySelector('.md-body'));
 * ```
 */
export async function renderMarkdownExtras(
    root: ParentNode,
    options: MarkdownExtrasOptions = {}
): Promise<MarkdownExtrasResult> {
    if (!root || typeof root.querySelector !== 'function') {
        return { diagrams: 0, math: 0 };
    }
    const mathCount =
        options.math !== false && root.querySelector('[data-tex]') ? renderMath(root) : Promise.resolve(0);
    const diagramCount =
        options.diagrams !== false && root.querySelector('.jdoc-mermaid')
            ? renderDiagrams(root, options.theme)
            : Promise.resolve(0);

    const [math, diagrams] = await Promise.all([mathCount, diagramCount]);
    return { math, diagrams };
}

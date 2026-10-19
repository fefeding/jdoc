/**
 * Markdown 类型定义
 *
 * @module markdown/types
 */

/** Markdown 文档元信息（从内容推断，Markdown 本身没有元数据区） */
export interface MarkdownMetadata {
    /** 首个一级标题（无 h1 时取首个标题） */
    title?: string;
    /** 标题大纲 */
    headings?: MarkdownHeading[];
    /** 中英文混排的粗略字数统计（去标记后的字符数） */
    wordCount?: number;
}

/** 标题项 */
export interface MarkdownHeading {
    /** 标题层级（1-6） */
    level: number;
    /** 标题文本 */
    text: string;
}

/** markdownToHtml / markdownToText 的可选项 */
export interface MarkdownOptions {
    /**
     * 是否附带默认排版样式（作用域限定在 `.jdoc-md` 内）。
     * 默认 true；宿主已有自己的 Markdown 样式时可关闭。
     */
    includeStyle?: boolean;
    /**
     * 是否保留 Markdown 中内嵌的原始 HTML。
     * 默认 true（会经安全清洗：去掉 script/iframe/事件属性与危险协议）；
     * 关闭后原始 HTML 会被当作纯文本转义显示，渲染结果 100% 安全但丢失富内容。
     */
    allowHtml?: boolean;
    /** 是否启用 GFM（表格、删除线、任务列表、自动链接），默认 true */
    gfm?: boolean;
    /** 是否把单个换行当作 <br>，默认 false（按 CommonMark 规范） */
    breaks?: boolean;
    /**
     * 是否启用公式语法（$…$ / $$…$$ / ```math），默认 true。
     * 仅影响标记：真正渲染需浏览器端调用 `renderMarkdownExtras()`。
     */
    math?: boolean;
    /**
     * 是否启用 ```mermaid 图表语法，默认 true。
     * 真正渲染同样由 `renderMarkdownExtras()` 完成。
     */
    diagrams?: boolean;
}

/** markdownToHtml 的返回结构 */
export interface MarkdownHtmlResult {
    /** 渲染后的 HTML（含可选样式表，根元素带 `jdoc-md` 类） */
    html: string;
    /** 去标记后的纯文本 */
    text: string;
    /** 推断出的元信息 */
    metadata?: MarkdownMetadata;
    /** 解析提示（如“原始 HTML 已被清洗”），不影响主流程 */
    messages: string[];
}

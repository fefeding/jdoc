// Markdown 围栏代码块的语法高亮。
//
// 为什么自己接而不是让库侧（@fefeding/jdoc）做：marked 只负责出
// <pre><code class="language-xxx">，高亮是要按宿主配色渲染的——VS Code 有
// 深/浅/高对比三档主题，写死一套 hljs 主题色会在浅色下突兀。这里保留
// highlight.js 的词法分析（纯离线、无 DOM 依赖），配色交给 markdownCodeCss()
// 随主题注入。
//
// 语言只注册常用的若干种：highlight.js 的 full build 有近 200 种语言、
// 体积多出好几 MB，而技术文档里绝大多数代码块就落在下面这些语言上；
// 未注册的语言会退回自动探测，仍能高亮出常见结构。

import hljs from 'highlight.js/lib/core';
import bash from 'highlight.js/lib/languages/bash';
import c from 'highlight.js/lib/languages/c';
import cpp from 'highlight.js/lib/languages/cpp';
import csharp from 'highlight.js/lib/languages/csharp';
import css from 'highlight.js/lib/languages/css';
import diff from 'highlight.js/lib/languages/diff';
import dockerfile from 'highlight.js/lib/languages/dockerfile';
import go from 'highlight.js/lib/languages/go';
import ini from 'highlight.js/lib/languages/ini';
import java from 'highlight.js/lib/languages/java';
import javascript from 'highlight.js/lib/languages/javascript';
import json from 'highlight.js/lib/languages/json';
import less from 'highlight.js/lib/languages/less';
import lua from 'highlight.js/lib/languages/lua';
import makefile from 'highlight.js/lib/languages/makefile';
import markdown from 'highlight.js/lib/languages/markdown';
import nginx from 'highlight.js/lib/languages/nginx';
import objectivec from 'highlight.js/lib/languages/objectivec';
import perl from 'highlight.js/lib/languages/perl';
import php from 'highlight.js/lib/languages/php';
import python from 'highlight.js/lib/languages/python';
import r from 'highlight.js/lib/languages/r';
import ruby from 'highlight.js/lib/languages/ruby';
import rust from 'highlight.js/lib/languages/rust';
import scss from 'highlight.js/lib/languages/scss';
import shell from 'highlight.js/lib/languages/shell';
import sql from 'highlight.js/lib/languages/sql';
import swift from 'highlight.js/lib/languages/swift';
import typescript from 'highlight.js/lib/languages/typescript';
import xml from 'highlight.js/lib/languages/xml';
import yaml from 'highlight.js/lib/languages/yaml';

hljs.registerLanguage('bash', bash);
hljs.registerLanguage('shell', shell);
hljs.registerLanguage('sh', shell);
hljs.registerLanguage('zsh', shell);
hljs.registerLanguage('c', c);
hljs.registerLanguage('cpp', cpp);
hljs.registerLanguage('csharp', csharp);
hljs.registerLanguage('cs', csharp);
hljs.registerLanguage('css', css);
hljs.registerLanguage('diff', diff);
hljs.registerLanguage('dockerfile', dockerfile);
hljs.registerLanguage('docker', dockerfile);
hljs.registerLanguage('go', go);
hljs.registerLanguage('ini', ini);
hljs.registerLanguage('toml', ini);
hljs.registerLanguage('java', java);
hljs.registerLanguage('javascript', javascript);
hljs.registerLanguage('js', javascript);
hljs.registerLanguage('jsx', javascript);
hljs.registerLanguage('json', json);
hljs.registerLanguage('less', less);
hljs.registerLanguage('lua', lua);
hljs.registerLanguage('makefile', makefile);
hljs.registerLanguage('markdown', markdown);
hljs.registerLanguage('md', markdown);
hljs.registerLanguage('nginx', nginx);
hljs.registerLanguage('objectivec', objectivec);
hljs.registerLanguage('objc', objectivec);
hljs.registerLanguage('perl', perl);
hljs.registerLanguage('php', php);
hljs.registerLanguage('python', python);
hljs.registerLanguage('py', python);
hljs.registerLanguage('r', r);
hljs.registerLanguage('ruby', ruby);
hljs.registerLanguage('rb', ruby);
hljs.registerLanguage('rust', rust);
hljs.registerLanguage('rs', rust);
hljs.registerLanguage('scss', scss);
hljs.registerLanguage('sql', sql);
hljs.registerLanguage('swift', swift);
hljs.registerLanguage('typescript', typescript);
hljs.registerLanguage('ts', typescript);
hljs.registerLanguage('tsx', typescript);
hljs.registerLanguage('xml', xml);
hljs.registerLanguage('html', xml);
hljs.registerLanguage('vue', xml);
hljs.registerLanguage('svg', xml);
hljs.registerLanguage('yaml', yaml);
hljs.registerLanguage('yml', yaml);

/**
 * 对容器内所有围栏代码块上色
 *
 * 跳过 mermaid：它随后会被 renderMarkdownExtras 渲染成图表，着色反而多余
 * （源码也确实不是给人读的）。已带 hljs 标记的也跳过，避免二次转义。
 */
export function highlightCodeBlocks(root: HTMLElement): number {
    const blocks = root.querySelectorAll<HTMLElement>(
        'pre > code[class*="language-"]:not([class*="language-mermaid"]), pre > code:not([class])'
    );
    let done = 0;
    for (const el of Array.from(blocks)) {
        if (el.dataset.hls !== 'done' && !el.querySelector('.hljs-keyword, .hljs-string, .hljs-title')) {
            const lang = (el.className.match(/language-([\w+#-]+)/)?.[1] || '').toLowerCase();
            const code = el.textContent ?? '';
            try {
                const result = lang && hljs.getLanguage(lang)
                    ? hljs.highlight(code, { language: lang, ignoreIllegals: true })
                    : hljs.highlightAuto(code);
                el.innerHTML = result.value;
                el.dataset.hls = 'done';
                done++;
            } catch {
                // 未注册/高亮失败：保留原始文本，不影响阅读
            }
        }
    }
    return done;
}

/**
 * highlight.js token 配色
 *
 * 不用 hljs 自带主题（那是固定配色，深浅主题要整份切换），改为映射到 VS Code
 * 语义色：能用令牌的就跟随主题（如 --vscode-textPreformat-foreground 是
 * pre 内的前景色），token 专属色则按主题给两套，挂在 vscode-dark /
 * vscode-high-contrast 类下——VS Code 切换颜色主题时 body 类名实时变化，
 * CSS 自然跟着切，不需要监听事件。
 */
export const markdownCodeCss = `
.hljs { background: transparent; color: inherit; }
.hljs-comment, .hljs-quote { color: var(--hl-comment); font-style: italic; }
.hljs-keyword, .hljs-selector-tag, .hljs-literal, .hljs-doctag, .hljs-formula { color: var(--hl-keyword); }
.hljs-string, .hljs-regexp, .hljs-addition, .hljs-attribute, .hljs-meta .hljs-string { color: var(--hl-string); }
.hljs-number, .hljs-symbol, .hljs-bullet, .hljs-link, .hljs-selector-attr, .hljs-selector-pseudo { color: var(--hl-number); }
.hljs-title, .hljs-title.class_, .hljs-title.function_, .hljs-section, .hljs-name { color: var(--hl-title); }
.hljs-type, .hljs-class .hljs-title, .hljs-built_in, .hljs-class .hljs-title.class_ { color: var(--hl-type); }
.hljs-attr, .hljs-variable, .hljs-template-variable, .hljs-params, .hljs-property, .hljs-constant { color: var(--hl-attr); }
.hljs-meta, .hljs-selector-id, .hljs-selector-class, .hljs-tag { color: var(--hl-meta); }
.hljs-deletion { color: var(--vscode-gitDecoration-deletedResourceForeground, #f14c4c); }
.hljs-emphasis { font-style: italic; }
.hljs-strong { font-weight: 700; }
body {
  --hl-keyword: #c586c0; --hl-string: #ce9178; --hl-number: #b5cea8;
  --hl-title: #dcdcaa; --hl-type: #4ec9b0; --hl-attr: #9cdcfe; --hl-meta: #c8c8c8;
  --hl-comment: #6a9955;
}
body.vscode-light, body.vscode-high-contrast-light {
  --hl-keyword: #af00db; --hl-string: #a31515; --hl-number: #098658;
  --hl-title: #795e26; --hl-type: #267f99; --hl-attr: #001080; --hl-meta: #800000;
  --hl-comment: #008000;
}
body.vscode-high-contrast, body.vscode-high-contrast-light {
  --hl-keyword: #f9f2f4; --hl-string: #f8b886; --hl-number: #b9f9ca;
  --hl-title: #faf067; --hl-type: #7ee787; --hl-attr: #9cdcfe; --hl-meta: #ffffff;
  --hl-comment: #7ca668;
}
`;
// 精简 SVG 图标集（stroke 风格，currentColor）。
const S = (inner: string, vb = '0 0 24 24'): string =>
    `<svg viewBox="${vb}" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${inner}</svg>`;

const ICONS: Record<string, string> = {
    bold: '<path d="M6 4h7a4 4 0 0 1 0 8H6z"/><path d="M6 12h8a4 4 0 0 1 0 8H6z"/>',
    italic: '<line x1="19" y1="4" x2="10" y2="4"/><line x1="14" y1="20" x2="5" y2="20"/><line x1="15" y1="4" x2="9" y2="20"/>',
    underline: '<path d="M6 3v7a6 6 0 0 0 12 0V3"/><line x1="4" y1="21" x2="20" y2="21"/>',
    strike: '<path d="M17 5H9a4 4 0 0 0 0 8h6a4 4 0 0 1 0 8H7"/><line x1="4" y1="12" x2="20" y2="12"/>',
    alignLeft: '<line x1="4" y1="6" x2="20" y2="6"/><line x1="4" y1="12" x2="14" y2="12"/><line x1="4" y1="18" x2="18" y2="18"/>',
    alignCenter: '<line x1="4" y1="6" x2="20" y2="6"/><line x1="7" y1="12" x2="17" y2="12"/><line x1="5" y1="18" x2="19" y2="18"/>',
    alignRight: '<line x1="4" y1="6" x2="20" y2="6"/><line x1="10" y1="12" x2="20" y2="12"/><line x1="6" y1="18" x2="20" y2="18"/>',
    alignJustify: '<line x1="4" y1="6" x2="20" y2="6"/><line x1="4" y1="12" x2="20" y2="12"/><line x1="4" y1="18" x2="20" y2="18"/>',
    listUl: '<line x1="9" y1="6" x2="20" y2="6"/><line x1="9" y1="12" x2="20" y2="12"/><line x1="9" y1="18" x2="20" y2="18"/><circle cx="4.5" cy="6" r="1.2" fill="currentColor"/><circle cx="4.5" cy="12" r="1.2" fill="currentColor"/><circle cx="4.5" cy="18" r="1.2" fill="currentColor"/>',
    listOl: '<line x1="10" y1="6" x2="20" y2="6"/><line x1="10" y1="12" x2="20" y2="12"/><line x1="10" y1="18" x2="20" y2="18"/><path d="M4 4h1.5v4"/><path d="M3.5 10.5h2v1.5l-2 1.5h2"/><path d="M3.5 16.5h2v1.5l-2 1.5h2"/>',
    image: '<rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><path d="M21 15l-5-5L5 21"/>',
    table: '<rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="3" x2="9" y2="21"/>',
    hr: '<line x1="4" y1="12" x2="20" y2="12"/>',
    quote: '<path d="M7 15h3l2-4V6H6v5h3z"/><path d="M16 15h3l2-4V6h-6v5h3z"/>',
    code: '<polyline points="16 6 21 12 16 18"/><polyline points="8 6 3 12 8 18"/>',
    plus: '<line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>',
    trash: '<polyline points="3 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4h6v2"/>',
    up: '<line x1="12" y1="19" x2="12" y2="5"/><polyline points="5 12 12 5 19 12"/>',
    down: '<line x1="12" y1="5" x2="12" y2="19"/><polyline points="19 12 12 19 5 12"/>',
    undo: '<polyline points="9 14 4 9 9 4"/><path d="M20 20a8 8 0 0 0-8-8H4"/>',
    redo: '<polyline points="15 14 20 9 15 4"/><path d="M4 20a8 8 0 0 1 8-8h8"/>',
    heading: '<path d="M6 4v16"/><path d="M18 4v16"/><path d="M6 12h12"/>',
    text: '<polyline points="4 7 4 4 20 4 20 7"/><line x1="9" y1="20" x2="15" y2="20"/><line x1="12" y1="4" x2="12" y2="20"/>',
    fill: '<path d="M19 11l-8-8-9 9 8 8z"/><path d="M5 14l4 4"/><path d="M21 15v6h-6"/>',
    color: '<path d="M12 3l7 7-7 7-7-7z"/><line x1="5" y1="21" x2="19" y2="21"/>',
    pageBreak: '<rect x="3" y="4" width="18" height="6" rx="1"/><line x1="2" y1="14" x2="22" y2="14" stroke-dasharray="3 3"/><rect x="3" y="16" width="18" height="4" rx="1"/>',
    save: '<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/>',
    eye: '<path d="M1 12s4-7 11-7 11 7 11 7-4 7-11 7-11-7-11-7z"/><circle cx="12" cy="12" r="3"/>',
    pencil: '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/>'
};

/** 取图标 innerHTML（svg 字符串）。未知名返回空串。 */
export function icon(name: string): string {
    return ICONS[name] ? S(ICONS[name]) : '';
}

/**
 * OOXML 包内路径工具
 *
 * 关系（rels）与部件路径的解析规则在「读图表」「写图表」两端都要用，
 * 放一处避免两边算出不同的路径（写错部件 = 保存后图表丢失）。
 *
 * @module core/ooxml
 */

import { attr, findAll, parseXml } from './xml';

/** 部件所在目录 */
export function dirOf(path: string): string {
    const idx = path.lastIndexOf('/');
    return idx < 0 ? '' : path.slice(0, idx);
}

/** 关系 Target（相对路径）-> ZIP 内绝对路径 */
export function resolveTarget(baseDir: string, target: string): string {
    if (target.startsWith('/')) return target.slice(1);
    const parts = baseDir ? baseDir.split('/').filter(Boolean) : [];
    for (const seg of target.split('/')) {
        if (seg === '.' || seg === '') continue;
        if (seg === '..') parts.pop();
        else parts.push(seg);
    }
    return parts.join('/');
}

/** 关系文件路径 */
export function relsPath(part: string): string {
    const dir = dirOf(part);
    const name = part.slice(dir.length + 1);
    return dir ? dir + '/_rels/' + name + '.rels' : '_rels/' + name + '.rels';
}

/** 关系条目 */
export interface RelEntry {
    /** 关系类型 */
    type: string;
    /** 目标 */
    target: string;
}

/** 解析 .rels：Id -> 关系条目 */
export function parseRels(xml: string): Map<string, RelEntry> {
    const out = new Map<string, RelEntry>();
    if (!xml) return out;
    for (const rel of findAll(parseXml(xml), 'Relationship')) {
        const id = attr(rel, 'Id');
        if (!id) continue;
        out.set(id, { type: attr(rel, 'Type') || '', target: attr(rel, 'Target') || '' });
    }
    return out;
}

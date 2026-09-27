import { OfficeContentNode } from '../types.js';

/**
 * A named anchor (`<a id="x"></a>`, a bookmark target) while a document is read: the ids of the node
 * it stands at, which resolveAnchorMarks adds to that node once the tree is built. Never left in an AST.
 */
const ANCHOR_MARK = '__anchor_mark__';

/** Nodes whose content is a line of text: an anchor mark inside one is one of its own ids. */
const ANCHOR_INLINE_HOLDERS = new Set<string>(['paragraph', 'heading', 'list', 'cell', 'definitionTerm', 'definitionDescription']);

/** An anchor mark naming `ids`. */
export const anchorMark = (ids: string[]): OfficeContentNode => ({ type: ANCHOR_MARK as any, metadata: { anchorIds: ids } as any });

export const isAnchorMark = (node: OfficeContentNode): boolean => (node.type as string) === ANCHOR_MARK;

/** `node`'s ids, then those of `ids` it does not have yet. */
function addIds(node: OfficeContentNode, ids: string[]): void {
    const meta: any = node.metadata ?? (node.metadata = {} as any);
    const own: string[] = meta.anchorIds ?? [];
    const seen = new Set(own);
    const added = ids.filter(id => !seen.has(id) && !!seen.add(id));
    if (added.length) meta.anchorIds = own.concat(added);
}

/**
 * `nodes` (and every node's children, notes and comments) with each anchor mark's ids added to the
 * node it stands at, after that node's own ids (HtmlGenerator writes the first id on the element and
 * the rest as anchors before it). In a line of text that is the picture the mark stands right before,
 * else the line's container; among blocks it is the next block (the one before, when none follows; the
 * container, when there is none; an empty paragraph, at the top of a document holding nothing else).
 * Each node's ids are added once, so many marks cost time in proportion to their number. The tree is
 * at most the parser's depth limit deep.
 */
export function resolveAnchorMarks(nodes: OfficeContentNode[], parent?: OfficeContentNode): OfficeContentNode[] {
    for (const node of nodes) {
        if (node.children?.length) node.children = resolveAnchorMarks(node.children, node);
        if (node.notes?.length) node.notes = resolveAnchorMarks(node.notes);
        if (node.comments?.length) node.comments = resolveAnchorMarks(node.comments);
    }
    if (!nodes.some(isAnchorMark)) return nodes;
    const inLine = !!parent && ANCHOR_INLINE_HOLDERS.has(parent.type);
    const out: OfficeContentNode[] = [];
    const parentIds: string[] = [];
    let carried: string[] = [];
    for (const node of nodes) {
        if (isAnchorMark(node)) {
            for (const id of ((node.metadata as any)?.anchorIds ?? []) as string[]) carried.push(id);
            continue;
        }
        if (carried.length) {
            if (inLine && node.type !== 'image') for (const id of carried) parentIds.push(id);
            else addIds(node, carried);
            carried = [];
        }
        out.push(node);
    }
    if (carried.length) {
        if (inLine) for (const id of carried) parentIds.push(id);
        else if (out.length) addIds(out[out.length - 1], carried);
        else if (parent) addIds(parent, carried);
        else out.push({ type: 'paragraph', metadata: { anchorIds: carried } as any, children: [] });
    }
    if (parentIds.length) addIds(parent!, parentIds);
    return out;
}

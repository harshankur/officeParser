import { OfficeContentNode, OfficeContentNodeType, OfficeParserAST } from '../types.js';

/** Every node type the AST defines (a record, so a type added to the union must be added here). */
const KNOWN_NODE_TYPES: Record<OfficeContentNodeType, true> = {
    paragraph: true, heading: true, table: true, list: true, text: true, image: true, chart: true, drawing: true,
    slide: true, note: true, sheet: true, row: true, cell: true, page: true, break: true, code: true, comment: true,
    header: true, footer: true, slideMaster: true, embed: true, admonition: true, definitionList: true,
    definitionTerm: true, definitionDescription: true,
};

/** Where a node sits: among blocks, in a line of text, among a table's rows, or among a row's cells. */
type Place = 'block' | 'inline' | 'table' | 'row';

/** Nodes whose children are blocks (paragraphs, tables, lists) rather than a line of text. */
const BLOCK_CONTAINERS = new Set<string>(['page', 'slide', 'note', 'admonition', 'cell', 'header', 'footer', 'slideMaster', 'chart', 'drawing', 'comment']);

/** Nodes that sit in a line of text; among blocks, a run of them is a paragraph. */
const INLINE_TYPES = new Set<string>(['text', 'image', 'break']);

const isKnown = (type: string): boolean => Object.prototype.hasOwnProperty.call(KNOWN_NODE_TYPES, type);

/** Where the children of a node of `type` sit. */
const placeOfChildren = (type: string): Place =>
    type === 'table' || type === 'sheet' ? 'table' : type === 'row' ? 'row' : BLOCK_CONTAINERS.has(type) ? 'block' : 'inline';

/**
 * `nodes` with each node of a type the AST does not define replaced by its content (its children, or
 * its text) in the form its place takes: among blocks a run of inline content is a paragraph, in a row
 * the content is a cell, and in a table a row. Returns `nodes` itself when nothing changes.
 */
function replaceUnknownNodes(nodes: OfficeContentNode[], place: Place): OfficeContentNode[] {
    let changed = false;
    const out: OfficeContentNode[] = [];
    for (const node of nodes) {
        if (isKnown(node.type)) {
            let next = node;
            for (const key of ['children', 'notes', 'comments'] as const) {
                const list = node[key];
                if (!list?.length) continue;
                const replaced = replaceUnknownNodes(list, key === 'children' ? placeOfChildren(node.type) : 'block');
                if (replaced !== list) next = { ...next, [key]: replaced };
            }
            if (next !== node) changed = true;
            out.push(next);
            continue;
        }
        changed = true;
        const notes = node.notes?.length ? replaceUnknownNodes(node.notes, 'block') : undefined;
        const comments = node.comments?.length ? replaceUnknownNodes(node.comments, 'block') : undefined;
        const content = (inner: Place): OfficeContentNode[] => {
            if (!node.children?.length) {
                // Its text, with its formatting, link, notes and comments.
                return node.text || notes || comments ? [{
                    type: 'text', text: node.text ?? '',
                    ...(node.formatting && { formatting: node.formatting }), ...(node.metadata && { metadata: node.metadata }),
                    ...(notes && { notes }), ...(comments && { comments }),
                } as OfficeContentNode] : [];
            }
            let replaced = replaceUnknownNodes(node.children, inner);
            // An inline wrapper's formatting and link reach the text it holds (the text's own first).
            if (inner === 'inline' && (node.formatting || node.metadata)) {
                replaced = replaced.map(child => (child.type === 'text'
                    ? { ...child, ...((node.formatting || child.formatting) && { formatting: { ...node.formatting, ...child.formatting } }), ...((child.metadata ?? node.metadata) && { metadata: child.metadata ?? node.metadata }) } as OfficeContentNode
                    : child));
            }
            // Its notes and comments are carried by the last of its content.
            if (notes || comments) {
                const last = replaced.length ? { ...replaced[replaced.length - 1] } : { type: 'text', text: '' } as OfficeContentNode;
                if (notes) last.notes = [...(last.notes ?? []), ...notes];
                if (comments) last.comments = [...(last.comments ?? []), ...comments];
                replaced = [...replaced.slice(0, -1), last];
            }
            return replaced;
        };
        if (place === 'inline') {
            for (const child of content('inline')) out.push(child);
        } else if (place === 'row') {
            out.push({ type: 'cell', children: content('block') });
        } else if (place === 'table') {
            for (const child of content('table')) out.push(child.type === 'row' ? child : { type: 'row', children: [{ type: 'cell', children: [child] }] });
        } else {
            let run: OfficeContentNode[] = [];
            const flush = () => {
                if (run.length) out.push({ type: 'paragraph', children: run });
                run = [];
            };
            for (const child of content('block')) {
                if (INLINE_TYPES.has(child.type)) run.push(child);
                else { flush(); out.push(child); }
            }
            flush();
        }
    }
    return changed ? out : nodes;
}

/**
 * The AST with every node of a type it does not define (from a hand-built AST, or one written by a
 * newer version) replaced by that node's content, so every generator writes the content instead of
 * failing on the node or leaving it out. Returns `ast` itself (same object, `.to()` intact) when all
 * types are known; otherwise a shallow copy with new `content`. The input is never mutated.
 */
export function withKnownNodeTypes<T extends OfficeParserAST>(ast: T): T {
    const content = replaceUnknownNodes(ast.content, 'block');
    return content === ast.content ? ast : { ...ast, content };
}

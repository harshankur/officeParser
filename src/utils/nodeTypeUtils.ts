import { mapNodeLists } from './nodeListUtils.js';
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

/** The formatting and link of the unknown wrappers around a node, for the text in them (its own first). */
interface Inherited {
    formatting?: OfficeContentNode['formatting'];
    metadata?: OfficeContentNode['metadata'];
}

/** `text` with the formatting and link of the wrappers around it under its own. */
const withInherited = (text: OfficeContentNode, inherited: Inherited | undefined): OfficeContentNode => {
    if (!inherited || (!inherited.formatting && !inherited.metadata)) return text;
    return {
        ...text,
        ...((inherited.formatting || text.formatting) && { formatting: { ...inherited.formatting, ...text.formatting } }),
        ...((text.metadata ?? inherited.metadata) && { metadata: text.metadata ?? inherited.metadata }),
    } as OfficeContentNode;
};

/** A node of a known type with its lists' unknown nodes replaced, once per node (see withKnownNodeTypes). */
function knownResult(node: OfficeContentNode, done: Map<OfficeContentNode, OfficeContentNode>): OfficeContentNode {
    let next = done.get(node);
    if (!next) {
        done.set(node, node);
        next = node;
        for (const key of ['children', 'notes', 'comments'] as const) {
            const list = node[key];
            if (!list?.length) continue;
            const replaced = replaceUnknownNodes(list, key === 'children' ? placeOfChildren(node.type) : 'block', done);
            if (replaced !== list) next = { ...next, [key]: replaced };
        }
        done.set(node, next);
    }
    return next;
}

/**
 * `nodes` with each node of a type the AST does not define replaced by its content (its children, or
 * its text) in the form its place takes: among blocks a run of inline content is a paragraph, in a row
 * the content is a cell, and in a table a row. Returns `nodes` itself when nothing changes. Written
 * into `into` when given (the list a wrapper's content joins).
 */
function replaceUnknownNodes(nodes: OfficeContentNode[], place: Place, done: Map<OfficeContentNode, OfficeContentNode>, inherited?: Inherited, into?: OfficeContentNode[]): OfficeContentNode[] {
    let changed = !!inherited || !!into;
    const out = into ?? [];
    for (const node of nodes) {
        if (isKnown(node.type)) {
            const next = knownResult(node, done);
            if (next !== node) changed = true;
            out.push(next.type === 'text' ? withInherited(next, inherited) : next);
            continue;
        }
        changed = true;
        writeUnknown(node, place, done, place === 'inline' ? inherited : undefined, out);
    }
    return changed ? out : nodes;
}

/**
 * Writes an unknown node's content into `out` in the form `place` takes, straight into it however many
 * unknown wrappers nest in a line of text (each writing its own content, none copying another's): the
 * formatting and link of the wrappers reach the text in them as it is made. Copied up and given to all
 * of a wrapper's text at every level, 1,000 nested wrappers around 100,000 runs took 80 times as long
 * as none. Among blocks and in a table a nested wrapper's content is paragraphs or rows of its own,
 * which the wrapper's formatting does not reach.
 */
function writeUnknown(node: OfficeContentNode, place: Place, done: Map<OfficeContentNode, OfficeContentNode>, inherited: Inherited | undefined, out: OfficeContentNode[]): void {
    const notes = node.notes?.length ? replaceUnknownNodes(node.notes, 'block', done) : undefined;
    const comments = node.comments?.length ? replaceUnknownNodes(node.comments, 'block', done) : undefined;
    // What text in it takes: its own formatting over its wrappers', its link before theirs.
    const inner: Inherited | undefined = node.formatting || node.metadata
        ? { formatting: node.formatting ? { ...inherited?.formatting, ...node.formatting } : inherited?.formatting, metadata: node.metadata ?? inherited?.metadata }
        : inherited;
    const leaf = (): OfficeContentNode | undefined => node.text || notes || comments
        ? withInherited({ type: 'text', text: node.text ?? '' } as OfficeContentNode, inner)
        : undefined;
    // Puts its notes and comments on `list[at]`, a copy of it.
    const carry = (list: OfficeContentNode[], at: number): void => {
        const last = { ...list[at] };
        if (notes) last.notes = [...(last.notes ?? []), ...notes];
        if (comments) last.comments = [...(last.comments ?? []), ...comments];
        list[at] = last;
    };
    if (place === 'row') {
        // In a row, its content is one cell, as it stands among blocks.
        let children: OfficeContentNode[];
        if (node.children?.length) {
            children = replaceUnknownNodes(node.children, 'block', done, inner);
            if (children === node.children) children = children.slice();
        } else {
            const text = leaf();
            children = text ? [text] : [];
        }
        if (notes || comments) {
            if (!children.length) children.push({ type: 'text', text: '' });
            carry(children, children.length - 1);
        }
        out.push({ type: 'cell', children });
        return;
    }
    // Among blocks, inline content waits in a run for its paragraph. Where the last of its content went,
    // for its notes and comments.
    let run: OfficeContentNode[] = [];
    let lastList: OfficeContentNode[] | undefined;
    let lastAt = -1;
    const flush = () => {
        if (run.length) out.push({ type: 'paragraph', children: run });
        run = [];
    };
    const write = (child: OfficeContentNode) => {
        if (place === 'inline') {
            out.push(child);
            lastList = out; lastAt = out.length - 1;
        } else if (place === 'table' && child.type !== 'row') {
            const cellChildren = [child];
            out.push({ type: 'row', children: [{ type: 'cell', children: cellChildren }] });
            lastList = cellChildren; lastAt = 0;
        } else if (place === 'block' && INLINE_TYPES.has(child.type)) {
            run.push(child);
            lastList = run; lastAt = run.length - 1;
        } else {
            if (place === 'block') flush();
            out.push(child);
            lastList = out; lastAt = out.length - 1;
        }
    };
    if (!node.children?.length) {
        const text = leaf();
        if (text) write(text);
    } else {
        for (const child of node.children) {
            if (isKnown(child.type)) {
                const next = knownResult(child, done);
                write(next.type === 'text' ? withInherited(next, inner) : next);
                continue;
            }
            // A wrapper among blocks makes paragraphs of its own: the run so far ends before it.
            if (place === 'block') flush();
            const before = out.length;
            writeUnknown(child, place, done, place === 'inline' ? inner : undefined, out);
            if (out.length > before) { lastList = out; lastAt = out.length - 1; }
        }
    }
    // Its notes and comments are carried by the last of its content (an empty text of their own, when
    // it has none).
    if (notes || comments) {
        if (!lastList) write({ type: 'text', text: '' });
        carry(lastList!, lastAt);
    }
    flush();
}

/**
 * The AST with every node of a type it does not define (from a hand-built AST, or one written by a
 * newer version) replaced by that node's content, so every generator writes the content instead of
 * failing on the node or leaving it out. Returns `ast` itself (same object, `.to()` intact) when all
 * types are known; otherwise a shallow copy with new `content`. The input is never mutated.
 */
export function withKnownNodeTypes<T extends OfficeParserAST>(ast: T): T {
    // A node the AST shares (one note every reference to it holds) is read once, not once per path to
    // it: notes referring to each other twice each took time doubling per level.
    const seen = new Map();
    return mapNodeLists(ast, nodes => replaceUnknownNodes(nodes, 'block', seen));
}

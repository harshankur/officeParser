import { mapNodeLists, NodeMemo, nodeMemo } from './nodeListUtils.js';
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

/** What one replacement of unknown nodes keeps across the AST. */
interface Pass {
    /** Each known node's result, once (a node the AST shares keeps one result). */
    known: NodeMemo<OfficeContentNode>;
    /**
     * Each unknown node's content among blocks, in a table or in a row, once: there it takes no formatting
     * from around it, so it is the same wherever the node stands. A node shared through notes lists (which
     * sharedNodeVisits counts once, as writers write a note once) was written again along every path to
     * it: 28 nodes, each holding the next twice in its notes, ended the process out of memory.
     */
    unknown: NodeMemo<Partial<Record<Place, OfficeContentNode[]>>>;
    /** Copies this pass made to carry notes and comments, not yet shared: more are added to them in place. */
    owned: WeakSet<OfficeContentNode>;
    /** What writing the content may still cost (nodes carried, copied or written again), and past it. */
    workLeft: number;
    onTooLarge: () => never;
}

const spend = (pass: Pass, work: number): void => {
    pass.workLeft -= work;
    if (pass.workLeft < 0) pass.onTooLarge();
};

/** A node of a known type with its lists' unknown nodes replaced, once per node (see withKnownNodeTypes). */
function knownResult(node: OfficeContentNode, pass: Pass): OfficeContentNode {
    let next = pass.known.get(node);
    if (!next) {
        pass.known.set(node, node);
        next = node;
        for (const key of ['children', 'notes', 'comments'] as const) {
            const list = node[key];
            if (!list?.length) continue;
            const replaced = replaceUnknownNodes(list, key === 'children' ? placeOfChildren(node.type) : 'block', pass);
            if (replaced !== list) next = { ...next, [key]: replaced };
        }
        pass.known.set(node, next);
    }
    return next;
}

/**
 * `nodes` with each node of a type the AST does not define replaced by its content (its children, or
 * its text) in the form its place takes: among blocks a run of inline content is a paragraph, in a row
 * the content is a cell, and in a table a row. Returns `nodes` itself when nothing changes.
 */
function replaceUnknownNodes(nodes: OfficeContentNode[], place: Place, pass: Pass, inherited?: Inherited): OfficeContentNode[] {
    // Copied at its first change only (text under a wrapper changes from the first): a list most passes
    // leave as it is.
    let out: OfficeContentNode[] | undefined = inherited ? [] : undefined;
    for (let i = 0; i < nodes.length; i++) {
        const node = nodes[i];
        if (isKnown(node.type)) {
            const next = knownResult(node, pass);
            if (next !== node) out ??= nodes.slice(0, i);
            out?.push(next.type === 'text' ? withInherited(next, inherited) : next);
            continue;
        }
        out ??= nodes.slice(0, i);
        writeUnknown(node, place, pass, place === 'inline' ? inherited : undefined, out);
    }
    return out ?? nodes;
}

/**
 * Writes an unknown node's content into `out` in the form `place` takes, straight into it however many
 * unknown wrappers nest in a line of text (each writing its own content, none copying another's): the
 * formatting and link of the wrappers reach the text in them as it is made. Copied up and given to all
 * of a wrapper's text at every level, 1,000 nested wrappers around 100,000 runs took 80 times as long
 * as none. Among blocks and in a table a nested wrapper's content is paragraphs or rows of its own,
 * which the wrapper's formatting does not reach; there it is written once and shared (see Pass.unknown).
 */
function writeUnknown(node: OfficeContentNode, place: Place, pass: Pass, inherited: Inherited | undefined, out: OfficeContentNode[]): void {
    if (place !== 'inline') {
        const written = pass.unknown.get(node)?.[place];
        if (written) {
            spend(pass, written.length);
            for (const item of written) out.push(item);
            return;
        }
    }
    const start = out.length;
    const notes = node.notes?.length ? replaceUnknownNodes(node.notes, 'block', pass) : undefined;
    const comments = node.comments?.length ? replaceUnknownNodes(node.comments, 'block', pass) : undefined;
    // What text in it takes: its own formatting over its wrappers', its link before theirs.
    const inner: Inherited | undefined = node.formatting || node.metadata
        ? { formatting: node.formatting ? { ...inherited?.formatting, ...node.formatting } : inherited?.formatting, metadata: node.metadata ?? inherited?.metadata }
        : inherited;
    const leaf = (): OfficeContentNode | undefined => node.text || notes || comments
        ? withInherited({ type: 'text', text: node.text ?? '' } as OfficeContentNode, inner)
        : undefined;
    // Puts its notes and comments on `list[at]`: on a copy of it the first time, then on that copy in
    // place (each wrapper around the same content adds its own), where re-spreading the notes gathered
    // so far at every level took time in the square of the depth.
    const carry = (list: OfficeContentNode[], at: number): void => {
        let last = list[at];
        if (!pass.owned.has(last)) {
            spend(pass, (last.notes?.length ?? 0) + (last.comments?.length ?? 0));
            last = { ...last, notes: last.notes ? last.notes.slice() : undefined, comments: last.comments ? last.comments.slice() : undefined };
            if (!last.notes) delete last.notes;
            if (!last.comments) delete last.comments;
            pass.owned.add(last);
            list[at] = last;
        }
        spend(pass, (notes?.length ?? 0) + (comments?.length ?? 0));
        if (notes) { const own = last.notes ??= []; for (const note of notes) own.push(note); }
        if (comments) { const own = last.comments ??= []; for (const comment of comments) own.push(comment); }
    };
    const remember = () => {
        if (place === 'inline') return;
        // Shared from now on: nothing adds to these in place any more.
        const written = out.slice(start);
        for (const item of written) pass.owned.delete(item);
        let entry = pass.unknown.get(node);
        if (!entry) pass.unknown.set(node, entry = {});
        entry[place] = written;
    };
    if (place === 'row') {
        // In a row, its content is one cell, as it stands among blocks.
        let children: OfficeContentNode[];
        if (node.children?.length) {
            children = replaceUnknownNodes(node.children, 'block', pass, inner);
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
        remember();
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
                const next = knownResult(child, pass);
                write(next.type === 'text' ? withInherited(next, inner) : next);
                continue;
            }
            // A wrapper among blocks makes paragraphs of its own: the run so far ends before it.
            if (place === 'block') flush();
            const before = out.length;
            writeUnknown(child, place, pass, place === 'inline' ? inner : undefined, out);
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
    remember();
}

/**
 * The AST with every node of a type it does not define (from a hand-built AST, or one written by a
 * newer version) replaced by that node's content, so every generator writes the content instead of
 * failing on the node or leaving it out. Returns `ast` itself (same object, `.to()` intact) when all
 * types are known; otherwise a shallow copy with new `content`. The input is never mutated. What
 * writing unknown nodes' content costs past `maxWork` (nodes carried, copied or written again for
 * shared content) calls `onTooLarge`.
 */
export function withKnownNodeTypes<T extends OfficeParserAST>(ast: T, limits: { maxWork?: number; onTooLarge?: () => never; tree?: boolean } = {}): T {
    // A node the AST shares (one note every reference to it holds) is read once, not once per path to
    // it: notes referring to each other twice each took time doubling per level. A `tree` (see
    // sharedNodeVisits) shares none.
    const pass: Pass = {
        known: nodeMemo(limits.tree),
        unknown: nodeMemo(limits.tree),
        owned: new WeakSet(),
        workLeft: limits.maxWork ?? Infinity,
        onTooLarge: limits.onTooLarge ?? (() => { throw new RangeError('Invalid array length'); }),
    };
    return mapNodeLists(ast, nodes => replaceUnknownNodes(nodes, 'block', pass));
}

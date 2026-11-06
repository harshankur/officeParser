/**
 * Disclosures (`disclosure` nodes, HTML `<details>`): a summary and the blocks a reader opens it to see.
 * HTML, EPUB and Markdown write one back as `<details>`; every other writer has no block that opens and
 * closes, and is given it as what those formats have: its summary as a bold line, then its blocks, all
 * shown, so nothing in it is lost.
 */
import type { DisclosureMetadata, OfficeContentNode, OfficeParserAST } from '../types.js';
import { mapNodeLists, NodeMemo, nodeMemo } from './nodeListUtils.js';

/** What a browser shows a disclosure with no summary as. */
export const DEFAULT_DISCLOSURE_SUMMARY = 'Details';

/** A disclosure's summary as text: its own, or what a browser shows where it has none. */
export function disclosureSummary(node: OfficeContentNode): string {
    const summary = (node.metadata as DisclosureMetadata | undefined)?.summary;
    const text = typeof summary === 'string' ? summary.replace(/\s+/g, ' ').trim() : '';
    return text || DEFAULT_DISCLOSURE_SUMMARY;
}

/** A disclosure as a bold summary line, then its blocks (the summary carries its ids). */
export function openedDisclosure(node: OfficeContentNode, children: OfficeContentNode[] = node.children ?? []): OfficeContentNode[] {
    const summary = disclosureSummary(node);
    const anchorIds = (node.metadata as DisclosureMetadata | undefined)?.anchorIds;
    const line = {
        type: 'paragraph', text: summary,
        ...(anchorIds?.length ? { metadata: { anchorIds: [...anchorIds] } } : {}),
        children: [{ type: 'text', text: summary, formatting: { bold: true } }],
    } as OfficeContentNode;
    return [line, ...children];
}

/** A node list with each disclosure, at any depth, opened (see openedDisclosure). Unchanged lists are returned as they are. */
function openAll(nodes: OfficeContentNode[], done: NodeMemo<OfficeContentNode[]>): OfficeContentNode[] {
    let out: OfficeContentNode[] | undefined;
    for (let i = 0; i < nodes.length; i++) {
        const node = nodes[i];
        let result = done.get(node);
        if (!result) {
            let next = node;
            if (node.children?.length) {
                const children = openAll(node.children, done);
                if (children !== node.children) next = { ...next, children };
            }
            if (node.notes?.length) {
                const notes = openAll(node.notes, done);
                if (notes !== node.notes) next = { ...next, notes };
            }
            if (node.comments?.length) {
                const comments = openAll(node.comments, done);
                if (comments !== node.comments) next = { ...next, comments };
            }
            result = next.type === 'disclosure' ? openedDisclosure(next) : [next];
            done.set(node, result);
        }
        if (result.length !== 1 || result[0] !== node) out ??= nodes.slice(0, i);
        if (out) for (const n of result) out.push(n);
    }
    return out ?? nodes;
}

/**
 * The AST with every disclosure opened into a bold summary line and its blocks, for a writer with no
 * block that opens and closes. Returns `ast` itself when it has none; the input is never mutated. A
 * `tree` (see sharedNodeVisits) is read without a memo.
 */
export function withDisclosuresOpened<T extends OfficeParserAST>(ast: T, options: { tree?: boolean } = {}): T {
    const seen = nodeMemo<OfficeContentNode[]>(options.tree);
    return mapNodeLists(ast, nodes => openAll(nodes, seen));
}

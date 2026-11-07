/**
 * Disclosures (`disclosure` nodes, HTML `<details>`): a summary (a `summary` node, its first child) and
 * the blocks a reader opens it to see. HTML, EPUB and Markdown write one back as `<details>`; every other
 * writer has no block that opens and closes, and is given it as what those formats have: its summary in
 * bold, then its blocks, all shown, so nothing in it is lost.
 */
import type { DisclosureMetadata, OfficeContentNode, OfficeParserAST } from '../types.js';
import { mapNodeLists, NodeMemo, nodeMemo } from './nodeListUtils.js';

/** What a browser shows a disclosure with no summary as. */
export const DEFAULT_DISCLOSURE_SUMMARY = 'Details';

/** A disclosure's summary: its first child when that is a `summary` node. */
export function summaryOf(node: OfficeContentNode): OfficeContentNode | undefined {
    const first = node.children?.[0];
    return first?.type === 'summary' ? first : undefined;
}

/** A disclosure's body: its children after its summary. */
export function bodyOf(node: OfficeContentNode): OfficeContentNode[] {
    const children = node.children ?? [];
    return children[0]?.type === 'summary' ? children.slice(1) : children;
}

/**
 * A summary's one paragraph, when that is all it holds and the paragraph has no metadata (alignment, ids,
 * a style): what HTML writes as the `<summary>`'s line itself, since a `<p>` in it puts the line under the
 * marker. Undefined for anything else.
 */
export function summaryLine(summary: OfficeContentNode): OfficeContentNode | undefined {
    const only = summary.children?.length === 1 ? summary.children[0] : undefined;
    if (only?.type !== 'paragraph') return undefined;
    return Object.values((only.metadata ?? {}) as object).some(value => value !== undefined) ? undefined : only;
}

/** The text of nodes, blocks set apart by a space, no deeper than `depth` levels. */
function textOf(nodes: OfficeContentNode[], depth = 0): string {
    return nodes.map(node => (node.children?.length && depth < 64 ? textOf(node.children, depth + 1) : node.text ?? '')).join(' ');
}

/**
 * A summary's content when it is one line of plain text (one paragraph of unformatted, unlinked text
 * runs, no break, picture or note), white space collapsed: what can stand in a `<summary>` written as
 * text. Undefined for anything else.
 */
export function plainSummaryText(summary: OfficeContentNode): string | undefined {
    const blocks = summary.children ?? [];
    if (blocks.length !== 1 || blocks[0].type !== 'paragraph' || blocks[0].notes?.length || blocks[0].comments?.length) return undefined;
    const paragraph = blocks[0];
    if (Object.values((paragraph.metadata ?? {}) as object).some(value => value !== undefined)) return undefined;
    const runs = paragraph.children;
    if (!runs?.length) return (paragraph.text ?? '').replace(/\s+/g, ' ').trim();
    for (const run of runs) {
        if (run.type !== 'text' || run.children?.length || run.notes?.length || run.comments?.length) return undefined;
        if (run.formatting && Object.values(run.formatting).some(value => value !== undefined && value !== false)) return undefined;
        if (run.metadata && Object.values(run.metadata as object).some(value => value !== undefined)) return undefined;
    }
    return runs.map(run => run.text ?? '').join('').replace(/\s+/g, ' ').trim();
}

/** A disclosure's summary as text: its own, or what a browser shows where it has none. */
export function disclosureSummary(node: OfficeContentNode): string {
    const summary = summaryOf(node);
    const text = summary ? textOf(summary.children ?? []).replace(/\s+/g, ' ').trim() : '';
    return text || DEFAULT_DISCLOSURE_SUMMARY;
}

/** Inline nodes with their text bold (links, notes and the rest kept). */
function bolded(nodes: OfficeContentNode[]): OfficeContentNode[] {
    return nodes.map(node => (node.type === 'text' ? { ...node, formatting: { ...node.formatting, bold: true } } : node) as OfficeContentNode);
}

/**
 * A disclosure as its summary, then its blocks: the summary's paragraphs with their text bold and its
 * other blocks as they are (a "Details" line in bold where it has none). The first carries its ids.
 * `summary` and `body` are given where they were mapped already (see withDisclosuresOpened).
 */
export function openedDisclosure(node: OfficeContentNode, summary: OfficeContentNode[] | undefined = summaryOf(node)?.children, body: OfficeContentNode[] = bodyOf(node)): OfficeContentNode[] {
    const anchorIds = (node.metadata as DisclosureMetadata | undefined)?.anchorIds;
    let lines: OfficeContentNode[] = (summary ?? []).map(block => (block.type === 'paragraph'
        ? { ...block, children: bolded(block.children?.length ? block.children : block.text ? [{ type: 'text', text: block.text }] : []) }
        : block) as OfficeContentNode);
    if (!lines.length) lines = [{ type: 'paragraph', text: DEFAULT_DISCLOSURE_SUMMARY, children: [{ type: 'text', text: DEFAULT_DISCLOSURE_SUMMARY, formatting: { bold: true } }] } as OfficeContentNode];
    if (anchorIds?.length) {
        const first = lines[0];
        const own = (first.metadata as { anchorIds?: string[] } | undefined)?.anchorIds ?? [];
        lines = [{ ...first, metadata: { ...(first.metadata as object), anchorIds: [...anchorIds, ...own] } } as OfficeContentNode, ...lines.slice(1)];
    }
    return [...lines, ...body];
}

/** A node list with each disclosure, at any depth, opened (see openedDisclosure). Unchanged lists are returned as they are. */
function openAll(nodes: OfficeContentNode[], done: NodeMemo<OfficeContentNode[]>): OfficeContentNode[] {
    let out: OfficeContentNode[] | undefined;
    for (let i = 0; i < nodes.length; i++) {
        const node = nodes[i];
        let result = done.get(node);
        if (!result) {
            let next = node;
            // A disclosure's summary and body are opened apart: its summary node is read by openedDisclosure.
            const summary = next.type === 'disclosure' ? summaryOf(next) : undefined;
            const summaryChildren = summary?.children?.length ? openAll(summary.children, done) : summary?.children;
            const own = next.type === 'disclosure' ? bodyOf(next) : next.children;
            const children = own?.length ? openAll(own, done) : own;
            if (next.type !== 'disclosure' && children !== next.children) next = { ...next, children };
            if (node.notes?.length) {
                const notes = openAll(node.notes, done);
                if (notes !== node.notes) next = { ...next, notes };
            }
            if (node.comments?.length) {
                const comments = openAll(node.comments, done);
                if (comments !== node.comments) next = { ...next, comments };
            }
            // A disclosure opens into its summary and body; a summary standing alone is its blocks.
            result = next.type === 'disclosure' ? openedDisclosure(next, summaryChildren, children ?? [])
                : next.type === 'summary' ? (next.children ?? [])
                : [next];
            done.set(node, result);
        }
        if (result.length !== 1 || result[0] !== node) out ??= nodes.slice(0, i);
        if (out) for (const n of result) out.push(n);
    }
    return out ?? nodes;
}

/**
 * The AST with every disclosure opened into its summary in bold and its blocks, for a writer with no
 * block that opens and closes. Returns `ast` itself when it has none; the input is never mutated. A
 * `tree` (see sharedNodeVisits) is read without a memo.
 */
export function withDisclosuresOpened<T extends OfficeParserAST>(ast: T, options: { tree?: boolean } = {}): T {
    const seen = nodeMemo<OfficeContentNode[]>(options.tree);
    return mapNodeLists(ast, nodes => openAll(nodes, seen));
}

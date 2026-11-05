/**
 * Pictures' captions (`ImageMetadata.caption`). HTML, EPUB and Markdown write a captioned picture as a
 * figure; every other writer has no figure that holds a caption, and is given the caption as what
 * those formats have: a caption paragraph (`style: 'Caption'`) after the picture, so it is never lost.
 */
import type { ImageMetadata, OfficeContentNode, OfficeParserAST } from '../types.js';
import { mapNodeLists, NodeMemo, nodeMemo } from './nodeListUtils.js';

/** The caption a picture carries, when it has one. */
export function captionOf(node: OfficeContentNode | undefined): string | undefined {
    if (node?.type !== 'image') return undefined;
    const caption = (node.metadata as ImageMetadata | undefined)?.caption;
    if (typeof caption !== 'string') return undefined;
    const trimmed = caption.trim();
    return trimmed || undefined;
}

/** The caption paragraph a picture's caption is written as where there is no figure. */
function captionParagraph(caption: string): OfficeContentNode {
    return { type: 'paragraph', text: caption, metadata: { style: 'Caption' }, children: [{ type: 'text', text: caption }] } as OfficeContentNode;
}

/** The blocks whose line a picture can stand in (its caption follows the block). */
const LINE_BLOCKS: ReadonlySet<string> = new Set(['paragraph', 'heading', 'list', 'definitionTerm', 'definitionDescription']);

/** A picture without its caption (the paragraph after it carries it). */
function withoutCaption(node: OfficeContentNode): OfficeContentNode {
    const { caption: _caption, ...rest } = node.metadata as ImageMetadata;
    return { ...node, metadata: rest as ImageMetadata } as OfficeContentNode;
}

type Placed = { node: OfficeContentNode; captions: string[] };

/**
 * A node without the captions of its pictures, and those captions, to follow it: its own (a picture's),
 * and those of the pictures in its line (a paragraph's, a heading's, a list item's), which cannot hold a
 * paragraph. A node that does not change is returned as it is.
 */
function placedNode(node: OfficeContentNode, done: NodeMemo<Placed>): Placed {
    const known = done.get(node);
    if (known) return known;
    let next = node;
    const captions: string[] = [];
    if (node.children?.length) {
        let children: OfficeContentNode[];
        if (LINE_BLOCKS.has(node.type)) {
            let changed = false;
            const kept = node.children.map(child => {
                const placed = placedNode(child, done);
                if (placed.node !== child) changed = true;
                captions.push(...placed.captions);
                return placed.node;
            });
            children = changed ? kept : node.children;
        } else {
            children = placeCaptions(node.children, done);
        }
        if (children !== node.children) next = { ...next, children };
    }
    if (node.notes?.length) {
        const notes = placeCaptions(node.notes, done);
        if (notes !== node.notes) next = { ...next, notes };
    }
    const own = captionOf(next);
    if (own) { next = withoutCaption(next); captions.unshift(own); }
    const placed = { node: next, captions };
    done.set(node, placed);
    return placed;
}

/** A node list with each captioned picture, or block holding one in its line, followed by the caption paragraphs. */
function placeCaptions(nodes: OfficeContentNode[], done: NodeMemo<Placed>): OfficeContentNode[] {
    let out: OfficeContentNode[] | undefined;
    for (let i = 0; i < nodes.length; i++) {
        const placed = placedNode(nodes[i], done);
        if (placed.node !== nodes[i] || placed.captions.length) out ??= nodes.slice(0, i);
        if (out) {
            out.push(placed.node);
            for (const caption of placed.captions) out.push(captionParagraph(caption));
        }
    }
    return out ?? nodes;
}

/**
 * The AST with each picture's caption as a caption paragraph after it, for a writer with no figure (see
 * the module comment). Returns `ast` itself when no picture has a caption; the input is never mutated.
 * A `tree` (see sharedNodeVisits) is read without a memo.
 */
export function withCaptionParagraphs<T extends OfficeParserAST>(ast: T, options: { tree?: boolean } = {}): T {
    const seen = nodeMemo<Placed>(options.tree);
    return mapNodeLists(ast, nodes => placeCaptions(nodes, seen));
}

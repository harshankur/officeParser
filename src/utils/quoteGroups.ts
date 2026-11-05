/**
 * The paragraphs of one quote. The AST has no quote container: a quote is a paragraph style, and the
 * paragraphs of a quote of several (Markdown `> a` `>` `> b`, HTML `<blockquote><p>a</p><p>b</p>
 * </blockquote>`, a LaTeX `quote` environment) share a `ParagraphMetadata.quoteId`. Writers with a
 * quote that holds paragraphs write adjacent paragraphs of one quote as one.
 */
import type { OfficeContentNode, ParagraphMetadata } from '../types.js';

/** The quote a paragraph belongs to, when its source grouped it with others. */
export function quoteIdOf(node: OfficeContentNode | undefined): string | undefined {
    if (node?.type !== 'paragraph') return undefined;
    const id = (node.metadata as ParagraphMetadata | undefined)?.quoteId;
    return typeof id === 'string' && id ? id : undefined;
}

/** Whether two paragraphs are paragraphs of one quote (adjacency is the caller's to check). */
export function sameQuote(a: OfficeContentNode | undefined, b: OfficeContentNode | undefined): boolean {
    const id = quoteIdOf(a);
    return id !== undefined && id === quoteIdOf(b);
}

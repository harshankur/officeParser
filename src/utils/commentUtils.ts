import { CommentMetadata, OfficeContentNode, OfficeParserAST } from '../types.js';

/**
 * Whether `node` is a SOURCE comment - `<!-- ... -->` from Markdown or HTML, the author's hidden note -
 * rather than a document review comment (see `CommentMetadata.sourceSyntax`).
 */
export function isSourceComment(node: OfficeContentNode): boolean {
    return node.type === 'comment' && (node.metadata as CommentMetadata | undefined)?.sourceSyntax === 'html';
}

/** Recursively drop source comments from a node list. Lists and nodes that don't change are returned as-is. */
function pruneSourceComments(nodes: OfficeContentNode[]): OfficeContentNode[] {
    let changed = false;
    const out: OfficeContentNode[] = [];
    for (const node of nodes) {
        if (isSourceComment(node)) {
            changed = true;
            continue;
        }
        let next = node;
        if (node.children?.length) {
            const children = pruneSourceComments(node.children);
            if (children !== node.children) next = { ...next, children };
        }
        if (node.notes?.length) {
            const notes = pruneSourceComments(node.notes);
            if (notes !== node.notes) next = { ...next, notes };
        }
        if (next !== node) changed = true;
        out.push(next);
    }
    return changed ? out : nodes;
}

/**
 * The AST without source comments, for an output format that has no hidden-comment construct: a source
 * comment is a note the author hid, so it is left out rather than rendered as visible content. Returns
 * `ast` itself (same object, `.to()` intact) when there are none; otherwise a shallow copy with a pruned
 * `content`. The input is never mutated.
 */
export function withoutSourceComments<T extends OfficeParserAST>(ast: T): T {
    const content = pruneSourceComments(ast.content);
    return content === ast.content ? ast : { ...ast, content };
}

import { CommentMetadata, OfficeContentNode, OfficeParserAST } from '../types.js';

/**
 * Whether `node` is a SOURCE comment - `<!-- ... -->` from Markdown or HTML, the author's hidden note -
 * rather than a document review comment (see `CommentMetadata.sourceSyntax`).
 */
export function isSourceComment(node: OfficeContentNode): boolean {
    return node.type === 'comment' && (node.metadata as CommentMetadata | undefined)?.sourceSyntax === 'html';
}

/** Recursively drop source comments from a node list. Lists and nodes that don't change are returned as-is. */
function pruneSourceComments(nodes: OfficeContentNode[], done: Map<OfficeContentNode, OfficeContentNode>): OfficeContentNode[] {
    let changed = false;
    const out: OfficeContentNode[] = [];
    for (const node of nodes) {
        if (isSourceComment(node)) {
            changed = true;
            continue;
        }
        // Each node once, its result shared: a note every reference holds stays one note (copied per
        // reference, it was written out at each), and notes nested in notes take linear time.
        let next = done.get(node);
        if (!next) {
            done.set(node, node);
            next = node;
            if (node.children?.length) {
                const children = pruneSourceComments(node.children, done);
                if (children !== node.children) next = { ...next, children };
            }
            if (node.notes?.length) {
                const notes = pruneSourceComments(node.notes, done);
                if (notes !== node.notes) next = { ...next, notes };
            }
            done.set(node, next);
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
    const content = pruneSourceComments(ast.content, new Map());
    return content === ast.content ? ast : { ...ast, content };
}

import { OfficeContentNode, OfficeParserAST } from '../types.js';

/** The node lists of `auxiliary` a writer renders (headers and footers on every page, slide masters, the outline). */
const AUXILIARY_LISTS = ['headers', 'footers', 'slideMasters', 'outline'] as const;

/**
 * `ast` with `transform` applied to its content and to every node list in its auxiliary, in that order
 * (a transform keeping state, a budget or a memo of shared nodes, keeps it across all of them); `ast`
 * itself, `.to()` intact, when none changed. A pass over `content` alone left headers, footers, slide
 * masters and the outline to reach the writers unchecked: a sheet in a footer was never held to the grid
 * budget, and 400 bytes of AST made 270 MB of DOCX. An auxiliary list that is not an array is left out.
 */
export function mapNodeLists<T extends OfficeParserAST>(ast: T, transform: (nodes: OfficeContentNode[]) => OfficeContentNode[]): T {
    const content = Array.isArray(ast.content) ? transform(ast.content) : [];
    const aux = ast.auxiliary;
    let auxiliary = aux;
    if (aux && typeof aux === 'object') {
        for (const key of AUXILIARY_LISTS) {
            const list = (aux as Record<string, unknown>)[key];
            if (list === undefined) continue;
            const next = Array.isArray(list) ? transform(list) : undefined;
            if (next !== list) auxiliary = { ...auxiliary, [key]: next };
        }
    } else if (aux !== undefined) {
        auxiliary = undefined;
    }
    return content === ast.content && auxiliary === aux ? ast : { ...ast, content, auxiliary };
}

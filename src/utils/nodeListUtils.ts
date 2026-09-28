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

/**
 * How many more nodes a writer meets, walking every path through `children` (content and auxiliary),
 * than the AST holds. A node two parents share is written under each, so an AST built in code whose
 * nodes each hold the next one twice doubled a writer's work at every level: 30 levels would take half
 * an hour. Parsers share nodes only within the repeated-content budget (repeated ODF cells), and a note
 * or comment is written once however many nodes hold it, so those count once.
 */
export function sharedNodeVisits(ast: OfficeParserAST): number {
    const along = new Map<OfficeContentNode, number>();
    const writtenOnce = new Set<OfficeContentNode>();
    const pending: OfficeContentNode[] = [];
    const count = (node: OfficeContentNode): number => {
        const known = along.get(node);
        if (known !== undefined) return known;
        along.set(node, 1);
        let total = 1;
        for (const child of Array.isArray(node.children) ? node.children : []) if (child && typeof child === 'object') total += count(child);
        for (const list of [node.notes, node.comments]) {
            if (Array.isArray(list)) for (const held of list) if (held && typeof held === 'object' && !writtenOnce.has(held)) { writtenOnce.add(held); pending.push(held); }
        }
        along.set(node, total);
        return total;
    };
    let visits = 0;
    const roots: unknown[] = [...(Array.isArray(ast.content) ? ast.content : [])];
    const aux = ast.auxiliary as Record<string, unknown> | undefined;
    if (aux && typeof aux === 'object') for (const key of AUXILIARY_LISTS) if (Array.isArray(aux[key])) roots.push(...(aux[key] as unknown[]));
    for (const root of roots) if (root && typeof root === 'object') visits += count(root as OfficeContentNode);
    while (pending.length) visits += count(pending.pop()!);
    return visits - along.size;
}

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
 * than the AST holds, plus every reference a node's notes and comments make. A node two parents share
 * is written under each, so an AST built in code whose nodes each hold the next twice doubled a
 * writer's work at every level: 30 levels would take half an hour. Parsers share nodes only within the
 * repeated-content budget (repeated ODF cells), and a note or comment is written once however many
 * nodes hold it, so its content counts once; each reference to it counts too, since passes and writers
 * handle each (1,000 wrappers holding one list of 16,000 notes made 16 million). Each list is read
 * once, however many nodes hold it: an AST built in code can give one array to a million nodes.
 */
export function sharedNodeVisits(ast: OfficeParserAST): number {
    const along = new Map<OfficeContentNode, number>();
    const listTotals = new Map<unknown[], number>();
    const writtenOnce = new Set<OfficeContentNode>();
    const readHeld = new Set<unknown[]>();
    const pending: OfficeContentNode[] = [];
    let references = 0;
    const listTotal = (list: unknown[]): number => {
        const known = listTotals.get(list);
        if (known !== undefined) return known;
        listTotals.set(list, 0);
        let total = 0;
        for (const child of list) if (child && typeof child === 'object') total += count(child as OfficeContentNode);
        listTotals.set(list, total);
        return total;
    };
    const count = (node: OfficeContentNode): number => {
        const known = along.get(node);
        if (known !== undefined) return known;
        along.set(node, 1);
        let total = 1;
        if (Array.isArray(node.children)) total += listTotal(node.children);
        for (const list of [node.notes, node.comments]) {
            if (!Array.isArray(list)) continue;
            references += list.length;
            if (readHeld.has(list)) continue;
            readHeld.add(list);
            for (const held of list) if (held && typeof held === 'object' && !writtenOnce.has(held)) { writtenOnce.add(held); pending.push(held); }
        }
        along.set(node, total);
        return total;
    };
    let visits = 0;
    if (Array.isArray(ast.content)) visits += listTotal(ast.content);
    const aux = ast.auxiliary as Record<string, unknown> | undefined;
    if (aux && typeof aux === 'object') for (const key of AUXILIARY_LISTS) if (Array.isArray(aux[key])) visits += listTotal(aux[key] as unknown[]);
    while (pending.length) visits += count(pending.pop()!);
    return visits - along.size + references;
}

/**
 * Appends `items` to `target` one by one. `appendAll(target, items)` passes every item as an argument,
 * which throws a RangeError once there are more than the engine allows (about 120,000): a slide group
 * of that many shapes, a chapter of that many paragraphs, or a paragraph of that many runs failed the
 * parse, and a writer reported it as nesting too deep.
 */
export function appendAll<T>(target: T[], items: readonly T[]): void {
    for (const item of items) target.push(item);
}

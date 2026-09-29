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

/** Four levels of a list item's indentation (text output writes four spaces a level, at most 64 levels) weigh one unit. */
const indentationWeight = (metadata: unknown): number => {
    const value = metadata && typeof metadata === 'object' ? (metadata as { indentation?: unknown }).indentation : undefined;
    const levels = typeof value === 'number' ? value : typeof value === 'string' ? Number(value) : 0;
    return levels > 0 ? Math.min(64, Math.floor(levels)) >> 2 : 0;
};

/**
 * How much more a writer handles, walking every path through `children` (content and auxiliary), than
 * the AST holds; in units of a node (a node counts 1, and 16 characters of its text count 1 more; so do
 * each entry of its formatting, metadata and attributes and every 16 characters of their keys and
 * values, each note or comment reference it holds, four levels of a list item's indentation, and what
 * `weigh` adds for it, such as the empty positions a table's grid fills). A node two parents share is
 * written under each, so an AST built in code whose nodes each hold the next twice doubled a writer's
 * work at every level: 30 levels would take half an hour. The same goes for a list, a record or a value many nodes share, which JSON cannot express but
 * an AST built in code can: one list of 16,000 note references, one array of 3,000 ids or one record of
 * 20,000 keys given to thousands of nodes. What the AST holds counts each node, list and record once;
 * what a writer handles counts them at every path and every node holding them. Parsers share only
 * within the repeated-content budget (repeated ODF cells, a style given to many runs). A note or comment
 * is written once however many nodes hold it, so its content counts once (a slide's notes are written
 * with the slide, so they count at every path to it). Each list and record is read once however many
 * nodes hold it, and a list longer than `limit` (a sparse array of four billion slots) is not read at
 * all: `extra` is then Infinity. Returns that `extra`, and what the AST `held` in the same units.
 */
export function sharedNodeVisits(ast: OfficeParserAST, limit = Infinity, weigh?: (node: OfficeContentNode) => number): { extra: number; held: number } {
    const along = new Map<OfficeContentNode, number>();
    const listTotals = new Map<unknown[], number>();
    const referenceLists = new Set<unknown[]>();
    const valueTotals = new Map<object, number>();
    const writtenOnce = new Set<OfficeContentNode>();
    const pending: OfficeContentNode[] = [];
    let held = 0;
    let tooLong = false;
    // A string past 64 characters (a value) or 0 (text) weighs a node for every 16 characters beyond.
    const stringWeight = (value: unknown, allowance: number): number =>
        typeof value === 'string' && value.length > allowance ? (value.length - allowance) >> 4 : 0;
    // What the AST holds of a string: its weight the first time a node holds it. A long string is one
    // value however many nodes hold it (an AST built in code gave one of 100 MB to a thousand nodes, and
    // each counted as content of its own), so later holders count only as writing it again.
    const seenStrings = new Set<string>();
    const heldWeight = (value: unknown, allowance: number): number => {
        const weight = stringWeight(value, allowance);
        if (weight === 0 || (value as string).length <= 64) return weight;
        if (seenStrings.has(value as string)) return 0;
        seenStrings.add(value as string);
        return weight;
    };
    // What a record (formatting, metadata, attributes) or array weighs per node holding it, as writers
    // write it at each: one per entry, and one per 16 characters of its keys and strings. Counted once
    // towards what the AST holds, a long string once by value. Short values were free, and one record of
    // 32 attributes of 64 characters shared by 200,000 paragraphs made 492 MB of HTML.
    const valueTotal = (value: object): number => {
        const known = valueTotals.get(value);
        if (known !== undefined) return known;
        valueTotals.set(value, 0);
        let entries = 0;
        let chars = 0;
        let heldChars = 0;
        let nested = 0;
        const text = (item: string) => {
            chars += item.length;
            if (item.length <= 64) heldChars += item.length;
            else if (!seenStrings.has(item)) { seenStrings.add(item); heldChars += item.length; }
        };
        if (Array.isArray(value)) {
            if (value.length > limit) { tooLong = true; return 0; }
            entries = value.length;
            for (const item of value) {
                if (item && typeof item === 'object') nested += valueTotal(item);
                else if (typeof item === 'string') text(item);
            }
        } else {
            for (const key in value) {
                entries++;
                text(key);
                const item = (value as Record<string, unknown>)[key];
                if (item && typeof item === 'object') nested += valueTotal(item);
                else if (typeof item === 'string') text(item);
            }
        }
        const own = entries + (chars >> 4);
        held += entries + (heldChars >> 4);
        valueTotals.set(value, own + nested);
        return own + nested;
    };
    const listTotal = (list: unknown[]): number => {
        const known = listTotals.get(list);
        if (known !== undefined) return known;
        listTotals.set(list, 0);
        if (list.length > limit) { tooLong = true; return 0; }
        let total = 0;
        for (const child of list) if (child && typeof child === 'object') total += count(child as OfficeContentNode);
        listTotals.set(list, total);
        return total;
    };
    // What a list of references (notes, comments) weighs per node holding it: one per reference, the
    // list counted once towards what the AST holds.
    const referencesOf = (list: unknown[]): number => {
        if (referenceLists.has(list)) return list.length;
        referenceLists.add(list);
        if (list.length > limit) { tooLong = true; return 0; }
        held += list.length;
        for (const item of list) if (item && typeof item === 'object' && !writtenOnce.has(item as OfficeContentNode)) { writtenOnce.add(item as OfficeContentNode); pending.push(item as OfficeContentNode); }
        return list.length;
    };
    // A chart's data as writers read it (see normalizeAttachments): its labels and texts, and each
    // series' name, values and point labels, one level deep (an entry that is a list is written empty).
    // Each list, series and chart is counted once towards what the AST holds.
    const chartTotals = new Map<object, number>();
    const entriesTotal = (list: unknown): number => {
        if (typeof list === 'string') return stringWeight(list, 64);
        if (!Array.isArray(list)) return 0;
        const known = chartTotals.get(list);
        if (known !== undefined) return known;
        chartTotals.set(list, 0);
        if (list.length > limit) { tooLong = true; return 0; }
        let total = list.length;
        held += list.length;
        for (const item of list) if (typeof item === 'string') { total += stringWeight(item, 64); held += heldWeight(item, 64); }
        chartTotals.set(list, total);
        return total;
    };
    const chartTotal = (chart: Record<string, unknown>): number => {
        const known = chartTotals.get(chart);
        if (known !== undefined) return known;
        chartTotals.set(chart, 0);
        let total = entriesTotal(chart.labels) + entriesTotal(chart.rawTexts) + entriesTotal(chart.title) + entriesTotal(chart.xAxisTitle) + entriesTotal(chart.yAxisTitle);
        const dataSets = chart.dataSets;
        if (Array.isArray(dataSets)) {
            if (dataSets.length > limit) { tooLong = true; return 0; }
            total += dataSets.length;
            held += dataSets.length;
            for (const series of dataSets) {
                if (!series || typeof series !== 'object') continue;
                let own = chartTotals.get(series);
                if (own === undefined) {
                    chartTotals.set(series, 0);
                    const { name, values, pointLabels } = series as Record<string, unknown>;
                    own = entriesTotal(name) + entriesTotal(values) + entriesTotal(pointLabels);
                    chartTotals.set(series, own);
                }
                total += own;
            }
        }
        chartTotals.set(chart, total);
        return total;
    };
    const count = (node: OfficeContentNode): number => {
        const known = along.get(node);
        if (known !== undefined) return known;
        along.set(node, 1);
        // The node itself: its text (and raw source), and what its records and reference lists weigh.
        const raw = (node as { rawContent?: unknown }).rawContent;
        // A list item's indentation, which text output writes as four spaces a level (at most 64), and
        // what the caller weighs a node by: both written at every path to it.
        const extra = (node.type === 'list' ? indentationWeight(node.metadata) : 0) + (weigh ? weigh(node) : 0);
        const text = stringWeight(node.text, 0);
        held += 1 + heldWeight(node.text, 0) + heldWeight(raw, 0) + extra;
        let total = 1 + stringWeight(raw, 0) + extra;
        for (const record of [node.formatting, node.metadata, (node as { htmlAttributes?: unknown }).htmlAttributes]) {
            if (record && typeof record === 'object') total += valueTotal(record);
        }
        // A node's text beside its children is the text they hold (a paragraph's, of its runs), which
        // writers write in place of them (chunks) or not at all: it weighs what it outweighs them by.
        // Counted with them, one paragraph of 100 MB counted its text twice and was refused.
        if (Array.isArray(node.children) && node.children.length) total += Math.max(text, listTotal(node.children));
        else total += text;
        if (node.type === 'slide' && Array.isArray(node.notes)) total += listTotal(node.notes);
        else if (Array.isArray(node.notes)) total += referencesOf(node.notes);
        if (Array.isArray(node.comments)) total += referencesOf(node.comments);
        along.set(node, total);
        return total;
    };
    let visits = 0;
    // The document's own metadata is written once, but a record in it (its custom properties) can share
    // a list or a long string across thousands of keys, each written out.
    if (ast.metadata && typeof ast.metadata === 'object') visits += valueTotal(ast.metadata);
    // So can a chart's data (series sharing one list of values), which writers write with the chart. Only
    // the chart data of an attachment is weighed: its file is written once, whatever its size.
    if (Array.isArray(ast.attachments)) {
        if (ast.attachments.length > limit) return { extra: Infinity, held };
        for (const attachment of ast.attachments as unknown[]) {
            const chartData = attachment && typeof attachment === 'object' ? (attachment as { chartData?: unknown }).chartData : undefined;
            if (chartData && typeof chartData === 'object' && !Array.isArray(chartData)) visits += chartTotal(chartData as Record<string, unknown>);
        }
    }
    if (Array.isArray(ast.content)) visits += listTotal(ast.content);
    const aux = ast.auxiliary as Record<string, unknown> | undefined;
    if (aux && typeof aux === 'object') for (const key of AUXILIARY_LISTS) if (Array.isArray(aux[key])) visits += listTotal(aux[key] as unknown[]);
    while (pending.length) visits += count(pending.pop()!);
    return { extra: tooLong ? Infinity : visits - held, held };
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

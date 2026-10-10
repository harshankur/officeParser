import { mapNodeLists, NodeMemo, nodeMemo } from './nodeListUtils.js';
import { documentBytesOf, GRID_POSITIONS_PER_BYTE } from './budgetUtils.js';
import { OfficeContentNode, OfficeParserAST } from '../types.js';
import { layoutTableRows } from './tableLayout.js';

/**
 * How many empty positions the grids of one document's tables and sheets may hold in all, beyond their
 * cells. Every writer that places cells by their `row` and `col` fills the grid between them (HTML
 * draws every position of a sheet, CSV and Markdown write an empty cell for each), so a few cells far
 * apart, one spanning billions of columns, or many sheets each just within a budget of its own, would
 * make output no memory holds.
 */
export const MAX_SHEET_GRID_GAPS = 1_000_000;

type Cell = { node: OfficeContentNode; row: OfficeContentNode; rowIndex: number; r: number; c: number; rowSpan: number; colSpan: number };

const coordinate = (value: unknown): number | undefined =>
    typeof value === 'number' && Number.isFinite(value) && value >= 0 ? Math.floor(value) : undefined;

const span = (value: unknown): number =>
    typeof value === 'number' && Number.isFinite(value) && value > 1 ? Math.floor(value) : 1;

/**
 * `node` (a table or sheet) with its cells' coordinates rewritten when the grid they make holds more
 * empty positions than `budget.left`; `node` itself when it does not (its empty positions are then
 * taken from the budget). Each cell is where its `row` and `col` put it, or, lacking them, next in
 * reading order; its spans count, and so does the layout writers that place cells by occupancy make
 * (see tableLayout). A grid too large is laid out closer: first without the rows and columns no cell
 * starts or ends in (a far cell comes next to the others, a span covers the lines it touches); if that
 * is still too large (cells scattered along a diagonal, rows each spanning many below), each row's
 * cells follow one another, without spans.
 */
function boundGrid(node: OfficeContentNode, budget: GridBudget, filled?: Map<OfficeContentNode, number>): OfficeContentNode {
    const cells: Cell[] = [];
    let rowIndex = 0;
    // The furthest row a row names itself (a sparse sheet's row with no cells), which a writer fills up to.
    let namedRow = -1;
    for (const row of node.children ?? []) {
        if (row.type !== 'row') continue;
        namedRow = Math.max(namedRow, coordinate((row.metadata as { row?: unknown } | undefined)?.row) ?? -1);
        let col = 0;
        for (const cell of row.children ?? []) {
            if (cell.type !== 'cell') continue;
            const meta = cell.metadata as { row?: unknown; col?: unknown; rowSpan?: unknown; colSpan?: unknown } | undefined;
            const r = coordinate(meta?.row);
            const c = coordinate(meta?.col);
            const colSpan = span(meta?.colSpan);
            cells.push({ node: cell, row, rowIndex, r: r ?? rowIndex, c: c ?? col, rowSpan: span(meta?.rowSpan), colSpan });
            col = (c ?? col) + colSpan;
        }
        rowIndex++;
    }
    if (!cells.length) return node;

    let maxRow = Math.max(rowIndex - 1, namedRow);
    let maxCol = 0;
    for (const cell of cells) {
        maxRow = Math.max(maxRow, cell.r + cell.rowSpan - 1);
        maxCol = Math.max(maxCol, cell.c + cell.colSpan - 1);
    }
    const gaps = (rows: number, cols: number) => rows * cols - cells.length;
    // A layout is taken when the positions its writers make fit the budget: the grid its cells' places
    // fill (HTML, CSV and Markdown write every position), and the cells, merges and gaps writers that
    // place cells by occupancy (DOCX, ODT) write, which a staircase of rows each spanning a thousand
    // rows made a thousand times wider than the cells' own places said (a 1.5 KB EPUB, 500 MB of DOCX).
    const rowsOf = (table: OfficeContentNode) => (table.children ?? []).filter(row => row.type === 'row');
    const fits = (table: OfficeContentNode, fillGaps: number): boolean => {
        if (fillGaps > budget.left) return false;
        const laid = layoutTableRows(rowsOf(table), cells.length + budget.left);
        const needed = Math.max(fillGaps, laid ? laid.positions - cells.length : Infinity);
        if (needed > budget.left) return false;
        budget.left -= Math.max(0, needed);
        filled?.set(table, Math.max(0, needed));
        return true;
    };
    if (fits(node, gaps(maxRow + 1, maxCol + 1))) return node;

    // The lines a cell starts or ends in, in order, each mapped to its place among them.
    const linesOf = (ends: number[]): Map<number, number> => {
        const sorted = [...new Set(ends)].sort((a, b) => a - b);
        return new Map(sorted.map((line, i) => [line, i]));
    };
    const rowLines = linesOf(cells.flatMap(cell => [cell.r, cell.r + cell.rowSpan - 1]));
    const colLines = linesOf(cells.flatMap(cell => [cell.c, cell.c + cell.colSpan - 1]));
    const compactGaps = gaps(Math.max(rowLines.size, rowIndex), Math.max(colLines.size, 1));
    if (compactGaps <= budget.left) {
        const compact = new Map<OfficeContentNode, Record<string, unknown>>();
        for (const cell of cells) {
            const meta: Record<string, unknown> = { row: rowLines.get(cell.r)!, col: colLines.get(cell.c)! };
            if (cell.rowSpan > 1) meta.rowSpan = rowLines.get(cell.r + cell.rowSpan - 1)! - rowLines.get(cell.r)! + 1;
            if (cell.colSpan > 1) meta.colSpan = colLines.get(cell.c + cell.colSpan - 1)! - colLines.get(cell.c)! + 1;
            compact.set(cell.node, meta);
        }
        const compacted = rewrite(node, compact);
        if (fits(compacted, compactGaps)) { budget.laidOut++; return compacted; }
    }

    // Each row's cells one after another, in the order of the rows, without spans.
    const sequential = new Map<OfficeContentNode, Record<string, unknown>>();
    let previousRow: OfficeContentNode | undefined;
    let col = 0;
    for (const cell of cells) {
        if (cell.row !== previousRow) { previousRow = cell.row; col = 0; }
        sequential.set(cell.node, { row: cell.rowIndex, col, rowSpan: undefined, colSpan: undefined });
        col++;
    }
    budget.laidOut++;
    return rewrite(node, sequential);
}

/** `node` with each cell `rewritten` names given those metadata values (and each row the row of its first). */
function rewrite(node: OfficeContentNode, rewritten: Map<OfficeContentNode, Record<string, unknown>>): OfficeContentNode {
    return {
        ...node,
        children: (node.children ?? []).map(row => {
            if (row.type !== 'row') return row;
            const rowMeta = row.metadata as { row?: unknown } | undefined;
            if (!row.children?.some(cell => rewritten.has(cell))) {
                // A row without cells has no place in the new layout: it keeps none of its own.
                return rowMeta && rowMeta.row !== undefined ? { ...row, metadata: { ...rowMeta, row: undefined } as any } : row;
            }
            const children = row.children.map(cell => {
                const meta = rewritten.get(cell);
                return meta ? { ...cell, metadata: { ...(cell.metadata as object), ...meta } } as OfficeContentNode : cell;
            });
            const first = children.find(cell => rewritten.has(cell));
            return {
                ...row,
                ...(rowMeta && typeof rowMeta.row === 'number' && { metadata: { ...rowMeta, row: (first?.metadata as { row?: number } | undefined)?.row ?? rowMeta.row } as any }),
                children,
            };
        }),
    };
}

/** The empty positions a document's grids may still fill, and how many of its tables were laid out closer to fit. */
interface GridBudget { left: number; laidOut: number }

/** `nodes` with every table and sheet among them, or in them, bounded (see boundGrid); `nodes` when none changed. */
function boundGrids(nodes: OfficeContentNode[], budget: GridBudget, done: NodeMemo<OfficeContentNode>, filled?: Map<OfficeContentNode, number>): OfficeContentNode[] {
    // Copied at its first change only: a list most passes leave as it is.
    let out: OfficeContentNode[] | undefined;
    for (let i = 0; i < nodes.length; i++) {
        const node = nodes[i];
        // Each node once, its result shared: a node the AST shares (a note every reference holds) was
        // walked once per path to it, doubling per level of notes nested in notes.
        let next: OfficeContentNode | undefined = done.get(node);
        if (!next) {
            done.set(node, node);
            next = node;
            for (const key of ['children', 'notes', 'comments'] as const) {
                const list: OfficeContentNode[] | undefined = next[key];
                if (!list?.length) continue;
                const bounded = boundGrids(list, budget, done, filled);
                if (bounded !== list) next = { ...next, [key]: bounded };
            }
            if (next.type === 'table' || next.type === 'sheet') next = boundGrid(next, budget, filled);
            done.set(node, next);
        }
        if (next !== node) out ??= nodes.slice(0, i);
        out?.push(next);
    }
    return out ?? nodes;
}

/**
 * The empty positions one document's grids may fill: MAX_SHEET_GRID_GAPS, plus GRID_POSITIONS_PER_BYTE
 * for each byte of the document the AST was parsed from (see budgetUtils), so a large sparse sheet is
 * written where it stands. Padding short rows to their table's width (see BaseGenerator) takes as many.
 */
export const gridPositionsFor = (ast: OfficeParserAST): number => MAX_SHEET_GRID_GAPS + GRID_POSITIONS_PER_BYTE * documentBytesOf(ast);

/**
 * The AST with every table and sheet whose cell positions and spans make a grid too large to write
 * laid out closer (see boundGrid), all of the document's grids sharing one budget of empty positions
 * (see gridPositionsFor), so no writer fills billions of positions from a few cells, or many sheets
 * each just within a budget of their own. Returns `ast` itself (same object, `.to()` intact) when every
 * grid is within bounds; the input is never mutated. `filled`, when given, takes the empty positions
 * each resulting table or sheet fills: the budget is charged once for a table the AST shares, which
 * writers write along every path to it (sharedNodeVisits weighs them there). `onLaidOut` is told how
 * many tables were laid out closer, when any were: their cells no longer stand where the document put
 * them, which is reported. A `tree` (see sharedNodeVisits) is read without a memo.
 */
export function withBoundedSheetGrids<T extends OfficeParserAST>(ast: T, filled?: Map<OfficeContentNode, number>, options: { tree?: boolean; onLaidOut?: (tables: number) => void } = {}): T {
    // One budget for the document's content and its headers, footers, slide masters and outline.
    const budget: GridBudget = { left: gridPositionsFor(ast), laidOut: 0 };
    const done = nodeMemo<OfficeContentNode>(options.tree);
    const bounded = mapNodeLists(ast, nodes => boundGrids(nodes, budget, done, filled));
    if (budget.laidOut) options.onLaidOut?.(budget.laidOut);
    return bounded;
}

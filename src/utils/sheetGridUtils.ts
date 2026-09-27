import { OfficeContentNode, OfficeParserAST } from '../types.js';

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
 * reading order; its spans count. A grid too large is laid out closer: first without the rows and
 * columns no cell starts or ends in (a far cell comes next to the others, a span covers the lines it
 * touches); if that is still too large (cells scattered along a diagonal), each row's cells follow one
 * another, without spans.
 */
function boundGrid(node: OfficeContentNode, budget: { left: number }): OfficeContentNode {
    const cells: Cell[] = [];
    let rowIndex = 0;
    for (const row of node.children ?? []) {
        if (row.type !== 'row') continue;
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

    let maxRow = rowIndex - 1;
    let maxCol = 0;
    for (const cell of cells) {
        maxRow = Math.max(maxRow, cell.r + cell.rowSpan - 1);
        maxCol = Math.max(maxCol, cell.c + cell.colSpan - 1);
    }
    const gaps = (rows: number, cols: number) => rows * cols - cells.length;
    const full = gaps(maxRow + 1, maxCol + 1);
    if (full <= budget.left) {
        budget.left -= Math.max(0, full);
        return node;
    }

    // The lines a cell starts or ends in, in order, each mapped to its place among them.
    const linesOf = (ends: number[]): Map<number, number> => {
        const sorted = [...new Set(ends)].sort((a, b) => a - b);
        return new Map(sorted.map((line, i) => [line, i]));
    };
    const rowLines = linesOf(cells.flatMap(cell => [cell.r, cell.r + cell.rowSpan - 1]));
    const colLines = linesOf(cells.flatMap(cell => [cell.c, cell.c + cell.colSpan - 1]));
    const compactGaps = gaps(Math.max(rowLines.size, rowIndex), Math.max(colLines.size, 1));
    const compact = compactGaps <= budget.left;

    const rewritten = new Map<OfficeContentNode, Record<string, unknown>>();
    if (compact) {
        budget.left -= Math.max(0, compactGaps);
        for (const cell of cells) {
            const meta: Record<string, unknown> = { row: rowLines.get(cell.r)!, col: colLines.get(cell.c)! };
            if (cell.rowSpan > 1) meta.rowSpan = rowLines.get(cell.r + cell.rowSpan - 1)! - rowLines.get(cell.r)! + 1;
            if (cell.colSpan > 1) meta.colSpan = colLines.get(cell.c + cell.colSpan - 1)! - colLines.get(cell.c)! + 1;
            rewritten.set(cell.node, meta);
        }
    } else {
        // Each row's cells one after another, in the order of the rows.
        let previousRow: OfficeContentNode | undefined;
        let col = 0;
        for (const cell of cells) {
            if (cell.row !== previousRow) { previousRow = cell.row; col = 0; }
            rewritten.set(cell.node, { row: cell.rowIndex, col, rowSpan: undefined, colSpan: undefined });
            col++;
        }
    }
    return {
        ...node,
        children: (node.children ?? []).map(row => {
            if (row.type !== 'row' || !row.children?.some(cell => rewritten.has(cell))) return row;
            const children = row.children.map(cell => {
                const meta = rewritten.get(cell);
                return meta ? { ...cell, metadata: { ...(cell.metadata as object), ...meta } } as OfficeContentNode : cell;
            });
            const first = children.find(cell => rewritten.has(cell));
            const rowMeta = row.metadata as { row?: unknown } | undefined;
            return {
                ...row,
                ...(rowMeta && typeof rowMeta.row === 'number' && { metadata: { ...rowMeta, row: (first?.metadata as { row?: number } | undefined)?.row ?? rowMeta.row } as any }),
                children,
            };
        }),
    };
}

/** `nodes` with every table and sheet among them, or in them, bounded (see boundGrid); `nodes` when none changed. */
function boundGrids(nodes: OfficeContentNode[], budget: { left: number }): OfficeContentNode[] {
    let changed = false;
    const out = nodes.map(node => {
        let next = node;
        for (const key of ['children', 'notes', 'comments'] as const) {
            const list = next[key];
            if (!list?.length) continue;
            const bounded = boundGrids(list, budget);
            if (bounded !== list) next = { ...next, [key]: bounded };
        }
        if (next.type === 'table' || next.type === 'sheet') next = boundGrid(next, budget);
        if (next !== node) changed = true;
        return next;
    });
    return changed ? out : nodes;
}

/**
 * The AST with every table and sheet whose cell positions and spans make a grid too large to write
 * laid out closer (see boundGrid), all of the document's grids sharing one budget of empty positions,
 * so no writer fills billions of positions from a few cells, or many sheets each just within a budget
 * of their own. Returns `ast` itself (same object, `.to()` intact) when every grid is within bounds;
 * the input is never mutated.
 */
export function withBoundedSheetGrids<T extends OfficeParserAST>(ast: T): T {
    const content = boundGrids(ast.content, { left: MAX_SHEET_GRID_GAPS });
    return content === ast.content ? ast : { ...ast, content };
}

import { OfficeContentNode, OfficeParserAST } from '../types.js';

/**
 * How many empty positions a table's or sheet's grid may hold beyond its cells. Every writer that
 * places cells by their `row` and `col` fills the grid between them (HTML draws every position of a
 * sheet, CSV and Markdown write an empty cell for each), so a few cells far apart, or one spanning
 * billions of columns, would make a grid no memory holds.
 */
export const MAX_SHEET_GRID_GAPS = 2_000_000;

type Cell = { node: OfficeContentNode; row: OfficeContentNode; rowIndex: number; r?: number; c?: number; rowSpan: number; colSpan: number };

const coordinate = (value: unknown): number | undefined =>
    typeof value === 'number' && Number.isFinite(value) && value >= 0 ? Math.floor(value) : undefined;

const span = (value: unknown): number =>
    typeof value === 'number' && Number.isFinite(value) && value > 1 ? Math.floor(value) : 1;

/**
 * `node` (a table or sheet) with its cells' coordinates rewritten when the grid they make is too
 * large (see MAX_SHEET_GRID_GAPS); `node` itself when it is not. First the rows and columns no cell
 * starts or ends in are dropped (a far cell comes next to the others, a span covers the lines it
 * touches); if the grid is still too large (cells scattered along a diagonal), each row's cells are
 * placed one after another, as in a table without coordinates.
 */
function boundGrid(node: OfficeContentNode): OfficeContentNode {
    const cells: Cell[] = [];
    let rowIndex = 0;
    for (const row of node.children ?? []) {
        if (row.type !== 'row') continue;
        for (const cell of row.children ?? []) {
            if (cell.type !== 'cell') continue;
            const meta = cell.metadata as { row?: unknown; col?: unknown; rowSpan?: unknown; colSpan?: unknown } | undefined;
            cells.push({ node: cell, row, rowIndex, r: coordinate(meta?.row), c: coordinate(meta?.col), rowSpan: span(meta?.rowSpan), colSpan: span(meta?.colSpan) });
        }
        rowIndex++;
    }
    const placed = cells.filter(cell => cell.r !== undefined || cell.c !== undefined);
    if (!placed.length) return node;

    const budget = cells.length + MAX_SHEET_GRID_GAPS;
    let maxRow = rowIndex - 1;
    let maxCol = 0;
    for (const cell of placed) {
        if (cell.r !== undefined) maxRow = Math.max(maxRow, cell.r + cell.rowSpan - 1);
        if (cell.c !== undefined) maxCol = Math.max(maxCol, cell.c + cell.colSpan - 1);
    }
    if ((maxRow + 1) * (maxCol + 1) <= budget) return node;

    // The lines a cell starts or ends in, in order, each mapped to its place among them.
    const linesOf = (ends: number[]): Map<number, number> => {
        const sorted = [...new Set(ends)].sort((a, b) => a - b);
        return new Map(sorted.map((line, i) => [line, i]));
    };
    const rowLines = linesOf(placed.flatMap(cell => (cell.r === undefined ? [] : [cell.r, cell.r + cell.rowSpan - 1])));
    const colLines = linesOf(placed.flatMap(cell => (cell.c === undefined ? [] : [cell.c, cell.c + cell.colSpan - 1])));
    const compact = Math.max(rowLines.size, rowIndex) * Math.max(colLines.size, 1) <= budget;

    const rewritten = new Map<OfficeContentNode, Record<string, unknown>>();
    if (compact) {
        for (const cell of placed) {
            const meta: Record<string, unknown> = {};
            if (cell.r !== undefined) {
                meta.row = rowLines.get(cell.r)!;
                if (cell.rowSpan > 1) meta.rowSpan = rowLines.get(cell.r + cell.rowSpan - 1)! - rowLines.get(cell.r)! + 1;
            }
            if (cell.c !== undefined) {
                meta.col = colLines.get(cell.c)!;
                if (cell.colSpan > 1) meta.colSpan = colLines.get(cell.c + cell.colSpan - 1)! - colLines.get(cell.c)! + 1;
            }
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
function boundGrids(nodes: OfficeContentNode[]): OfficeContentNode[] {
    let changed = false;
    const out = nodes.map(node => {
        let next = node;
        for (const key of ['children', 'notes', 'comments'] as const) {
            const list = next[key];
            if (!list?.length) continue;
            const bounded = boundGrids(list);
            if (bounded !== list) next = { ...next, [key]: bounded };
        }
        if (next.type === 'table' || next.type === 'sheet') next = boundGrid(next);
        if (next !== node) changed = true;
        return next;
    });
    return changed ? out : nodes;
}

/**
 * The AST with every table and sheet whose cell coordinates make a grid too large to write bounded
 * (see boundGrid), so no writer fills billions of empty positions from a few cells. Returns `ast`
 * itself (same object, `.to()` intact) when every grid is within bounds; the input is never mutated.
 */
export function withBoundedSheetGrids<T extends OfficeParserAST>(ast: T): T {
    const content = boundGrids(ast.content);
    return content === ast.content ? ast : { ...ast, content };
}

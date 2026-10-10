import { OfficeContentNode } from '../types.js';

/** The widest grid, and the longest vertical merge, a writer lays a table out with. */
export const MAX_LAYOUT_COLUMNS = 1000;
export const MAX_LAYOUT_ROW_SPAN = 1000;

/** One position of a laid-out row, left to right. */
export type LayoutSlot =
    /** A cell of the row, at `col`, spanning `colSpan` columns and `rowSpan` rows. */
    | { kind: 'cell'; cell: OfficeContentNode; col: number; colSpan: number; rowSpan: number }
    /** The part of an earlier row's vertical merge this row holds, `span` columns wide. */
    | { kind: 'covered'; col: number; span: number }
    /** An empty cell filling a column the row skips (a sparse grid, or before a pending merge). */
    | { kind: 'gap'; col: number };

export interface TableLayout {
    /** Columns the grid is wide (at most MAX_LAYOUT_COLUMNS). */
    cols: number;
    rows: LayoutSlot[][];
    /** Positions written in all (each covered slot counts once per column it spans). */
    positions: number;
}

/**
 * How writers that place cells by occupancy (DOCX, ODT) lay `rows` out: each cell in the next column
 * no earlier row's vertical merge still holds, or at its own `col` when that lies further right; a
 * merge's later rows hold its columns until it ends. Undefined once the positions pass `limit`: a
 * table of rows each holding one cell spanning a thousand rows lays out a staircase, every row a column
 * wider than the last, and a budget that placed each row's cells by their order alone saw one column.
 * One layout for the writers and for that budget, so they agree.
 */
export function layoutTableRows(rows: OfficeContentNode[], limit = Infinity): TableLayout | undefined {
    // Pending vertical merges by their left column, and the rightmost of them (-1 for none): asked of
    // every position, a scan of all of them made a row of a thousand merges cost a million steps.
    const active = new Map<number, { remaining: number; span: number }>();
    let rightmost = -1;
    const end = (col: number) => {
        active.delete(col);
        if (col === rightmost) { rightmost = -1; for (const left of active.keys()) if (left > rightmost) rightmost = left; }
    };
    const out: LayoutSlot[][] = [];
    let positions = 0;
    let width = 1;
    for (const row of rows) {
        const cells = (row.children || []).filter(c => c.type === 'cell');
        const slots: LayoutSlot[] = [];
        let col = 0, ci = 0;
        while (ci < cells.length || rightmost >= col) {
            if (positions > limit) return undefined;
            const merge = active.get(col);
            if (merge) {
                slots.push({ kind: 'covered', col, span: merge.span });
                positions += merge.span;
                if (--merge.remaining <= 0) end(col);
                col += merge.span;
                continue;
            }
            if (ci >= cells.length) {
                // A merge is still pending further right: fill this column so it lands in place.
                if (rightmost <= col) break;
                slots.push({ kind: 'gap', col });
                positions++;
                col++;
                continue;
            }
            // A sparse grid (an XLSX sheet names each cell's column): fill the columns it skips.
            const nextCol = (cells[ci].metadata as { col?: unknown } | undefined)?.col;
            if (typeof nextCol === 'number' && nextCol > col && col < MAX_LAYOUT_COLUMNS) {
                slots.push({ kind: 'gap', col });
                positions++;
                col++;
                continue;
            }
            const cell = cells[ci++];
            const meta = cell.metadata as { colSpan?: unknown; rowSpan?: unknown } | undefined;
            const colSpan = Math.max(1, Math.min(MAX_LAYOUT_COLUMNS, Math.floor(Number(meta?.colSpan)) || 1));
            const rowSpan = Math.max(1, Math.min(MAX_LAYOUT_ROW_SPAN, Math.floor(Number(meta?.rowSpan)) || 1));
            slots.push({ kind: 'cell', cell, col, colSpan, rowSpan });
            positions += colSpan;
            if (rowSpan > 1) {
                active.set(col, { remaining: rowSpan - 1, span: colSpan });
                if (col > rightmost) rightmost = col;
            }
            col += colSpan;
        }
        if (col > width) width = col;
        out.push(slots);
    }
    return { cols: Math.min(MAX_LAYOUT_COLUMNS, width), rows: out, positions };
}

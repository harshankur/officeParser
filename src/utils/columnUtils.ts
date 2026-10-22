/**
 * Columns as a writer can set them: held to what a page holds, as the readers hold the columns they read
 * (see WordParser's pageSetupOf), so a document of columns no page holds (one made by hand, or one a
 * reader of another kind gave) cannot make a file an office suite will not open, or a page that loses
 * its text.
 *
 * @module utils/columnUtils
 */

import { PageSetup } from '../types.js';

/** The narrowest a column is set: half an inch, in points. */
const MIN_COLUMN = 36;

/**
 * `columns` held to a page `pageWidth` wide (points) whose text is `textWidth` wide: from 1 to 16 whole
 * columns, as many as half an inch each leaves room for; a gap of none to a quarter of the page, and no
 * more than leaves each column that half inch (sixteen columns of wide gaps left the text none, and an
 * office suite drew none of it); each own width more than nothing and no wider than the text (else the
 * columns share it evenly); undefined for one column, or none.
 */
export function heldColumns(columns: PageSetup['columns'] | undefined, pageWidth: number, textWidth: number): PageSetup['columns'] | undefined {
    if (!columns) return undefined;
    const room = Number.isFinite(textWidth) && textWidth > 0 ? textWidth : Math.max(0, pageWidth);
    const count = Math.min(16, Math.floor(room / MIN_COLUMN), Math.max(1, Math.round(Number(columns.count)) || 1));
    if (!(count >= 2)) return undefined;
    const gapOf = Number(columns.gap);
    const gap = Number.isFinite(gapOf) && gapOf > 0 ? Math.min(gapOf, Math.max(0, pageWidth) / 4, (room - count * MIN_COLUMN) / (count - 1)) : 0;
    const widths = Array.isArray(columns.widths) && columns.widths.length === count && columns.widths.every(w => Number.isFinite(w) && w > 0)
        ? columns.widths.map(w => Math.min(w, room)) : undefined;
    return { count, gap, ...(widths ? { widths } : {}), ...(columns.separator ? { separator: true } : {}) };
}

/**
 * A document as a workbook, for the spreadsheet writers (XLSX, ODS): its sheets as they are (a
 * spreadsheet's), or each of its tables as a sheet (or all on one), and its text outside them on a sheet
 * of its own; each cell typed as a spreadsheet types it, a number as a number with the format it is
 * shown in (a percentage, a currency, grouped thousands, its decimals), a date or a time as one, other
 * text as text. A spreadsheet's cell keeps its value and is shown as it was (CellMetadata.shownText); a
 * table's cell is read from its text (`19%`, `€1,250.00`, `2026-10-01`), a number with leading zeros
 * or more digits than a spreadsheet keeps (`007`, a card number) staying text.
 *
 * @module generators/sheets/workbook
 */

import { CellMetadata, CommentMetadata, ImageMode, ListMetadata, OfficeContentNode, RowMetadata, SheetMetadata, SlideMetadata } from '../../types.js';
import { stripInvalidXmlChars } from '../../utils/sanitize.js';
import { layoutTableRows } from '../../utils/tableLayout.js';

/** How a cell's value is held and shown. */
export interface CellValue {
    kind: 'number' | 'percent' | 'currency' | 'date' | 'datetime' | 'time' | 'text';
    /** The number held: a number's, a percentage's fraction, a date's or time's day count (1900 system) or day fraction. */
    number?: number;
    /** A date's or time's ISO text, as ODF holds it. */
    iso?: string;
    /** Decimals shown, thousands grouped, and a currency's symbol (before or after the number). */
    decimals?: number;
    grouping?: boolean;
    symbol?: string;
    symbolAfter?: boolean;
}

/** A cell: its text (a spreadsheet's stored value), what it shows (its value in its format), its value typed. */
export interface WorkCell { col: number; text: string; shown: string; value: CellValue; bold: boolean; colSpan: number; rowSpan: number; hidden?: boolean }
export interface WorkRow { cells: WorkCell[]; hidden?: boolean }
export interface WorkSheet { name: string; rows: WorkRow[]; hidden?: boolean; widths: number[] }

/** Spreadsheets' limits: rows, columns, a sheet name's length. */
export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;

const NUMBER = /^-?(?:\d+(?:\.\d+)?|\.\d+)$/;
/** A number as a spreadsheet stores it: its exponent too (`1E-10`), and up to seventeen digits (`0.30000000000000004`). */
const STORED_NUMBER = /^-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;
const GROUPED = /^-?\d{1,3}(?:,\d{3})+(?:\.\d+)?$/;
const CURRENCY_SYMBOLS = '$€£¥₹₩₽₺₪฿₫';

/** Whether a number is a whole one too long for a General format to show in full (twelve to fifteen digits, an id). */
export const wholeInFull = (v: CellValue): boolean => v.kind === 'number' && typeof v.number === 'number' && Number.isInteger(v.number) && Math.abs(v.number) >= 1e11 && Math.abs(v.number) < 1e15;

/** Decimals written after a number's point. */
const decimalsOf = (digits: string) => { const dot = digits.indexOf('.'); return dot < 0 ? 0 : digits.length - dot - 1; };

/**
 * A number's text as a number, or undefined: no leading zero but before a point (`007`, a code, stays
 * text), at most fifteen significant digits (more, a card number, a spreadsheet would round). A leading
 * plus (`+4915112345678`, a phone number) is not a number's: such text stays text.
 */
function plainNumber(text: string): number | undefined {
    if (text.startsWith('+')) return undefined;
    const bare = text.replace(/^-/, '').replace(/,/g, '');
    if (/^0\d/.test(bare)) return undefined;
    if (bare.replace(/\D/g, '').replace(/^0+/, '').length > 15) return undefined;
    const n = Number(text.replace(/,/g, ''));
    return Number.isFinite(n) ? n : undefined;
}

/** Days from 1899-12-30 (the 1900 date system, as a spreadsheet counts them) to a UTC date. */
const dayCount = (y: number, m: number, d: number) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86_400_000;

/** A date or time in ISO form (`2026-10-01`, `2026-10-01T14:05`, `14:05:00`) as a cell's value. */
function isoValue(text: string): CellValue | undefined {
    const dt = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$/.exec(text);
    if (dt) {
        const [y, mo, d] = [Number(dt[1]), Number(dt[2]), Number(dt[3])];
        if (mo < 1 || mo > 12 || d < 1 || d > 31 || y < 1900 || y > 9999) return undefined;
        const date = new Date(Date.UTC(y, mo - 1, d));
        if (date.getUTCMonth() !== mo - 1) return undefined;
        // (Before 1900-03-01 the system counts a day that never was: such a date stays text.)
        if (dayCount(y, mo, d) < 61) return undefined;
        if (dt[4] === undefined) return { kind: 'date', number: dayCount(y, mo, d), iso: text };
        const [h, mi, s] = [Number(dt[4]), Number(dt[5]), Number(dt[6] ?? 0)];
        if (h > 23 || mi > 59 || s > 59) return undefined;
        return { kind: 'datetime', number: dayCount(y, mo, d) + (h * 3600 + mi * 60 + s) / 86400, iso: `${dt[1]}-${dt[2]}-${dt[3]}T${dt[4]}:${dt[5]}:${dt[6] ?? '00'}` };
    }
    const t = /^(\d{2}):(\d{2})(?::(\d{2}))?$/.exec(text);
    if (t) {
        const [h, mi, s] = [Number(t[1]), Number(t[2]), Number(t[3] ?? 0)];
        if (h > 23 || mi > 59 || s > 59) return undefined;
        return { kind: 'time', number: (h * 3600 + mi * 60 + s) / 86400, iso: `PT${h}H${mi}M${s}S` };
    }
    return undefined;
}

/**
 * A spreadsheet's stored number (`stored`: a cell of a sheet, its text the value held) as a number, or
 * undefined: any finite number as a spreadsheet writes one, exponents and all seventeen digits a double
 * holds, but not a code with leading zeros (`007`, which a sheet holds as text) or more digits than a
 * double keeps (a card number).
 */
function storedNumber(text: string): number | undefined {
    const bare = text.replace(/^-/, '');
    if (/^0\d/.test(bare)) return undefined;
    const mantissa = bare.replace(/[eE].*$/, '').replace(/\D/g, '').replace(/^0+/, '');
    if (mantissa.length > 17) return undefined;
    const n = Number(text);
    return Number.isFinite(n) ? n : undefined;
}

/**
 * A cell's value from its text, and from what a spreadsheet showed for it (`shown`, where its text is
 * its stored value; `stored`, a sheet's cell, whose text is a value as the sheet holds it): a number, a
 * percentage, a currency amount, a date or a time, else text.
 */
export function cellValue(text: string, shown?: string, stored = shown !== undefined): CellValue {
    const s = text.trim();
    if (!s || s.length > 64) return { kind: 'text' };
    const iso = isoValue(s);
    if (iso) return iso;
    // A spreadsheet's number, shown through its format: the value held, the format read from what it showed.
    if (stored && STORED_NUMBER.test(s)) {
        const n = storedNumber(s);
        if (n !== undefined) {
            if (shown === undefined) return { kind: 'number', number: n };
            const look = cellValue(shown);
            if (look.kind === 'percent' || look.kind === 'currency' || (look.kind === 'number' && (look.grouping || look.decimals))) return { ...look, number: n };
            return { kind: 'number', number: n };
        }
        if (shown !== undefined) return { kind: 'text' };
    }
    if (NUMBER.test(s) || GROUPED.test(s)) {
        const n = plainNumber(s);
        return n === undefined ? { kind: 'text' } : { kind: 'number', number: n, decimals: decimalsOf(s.replace(/^[+-]/, '')), grouping: s.includes(',') };
    }
    const percent = /^([+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)\s?%$/.exec(s);
    if (percent) {
        const n = plainNumber(percent[1]);
        if (n !== undefined) return { kind: 'percent', number: n / 100, decimals: decimalsOf(percent[1]), grouping: percent[1].includes(',') };
    }
    // A currency amount: its symbol before (`$1,250.00`, `-€5`) or after (`120.00 €`); in brackets, a debit.
    const negative = /^\((.*)\)$/.exec(s);
    const body = negative ? negative[1].trim() : s;
    const money = new RegExp(`^([+-]?)([${CURRENCY_SYMBOLS}])\\s?([+-]?(?:\\d{1,3}(?:,\\d{3})+|\\d+)(?:\\.\\d+)?)$|^([+-]?(?:\\d{1,3}(?:,\\d{3})+|\\d+)(?:\\.\\d+)?)\\s?([${CURRENCY_SYMBOLS}])$`, 'u').exec(body);
    if (money) {
        const digits = money[3] ?? money[4];
        const n = plainNumber(digits.replace(/^[+-]/, ''));
        if (n !== undefined) {
            const sign = (negative || money[1] === '-' || digits.startsWith('-')) ? -1 : 1;
            return { kind: 'currency', number: sign * n, decimals: decimalsOf(digits.replace(/^[+-]/, '')), grouping: digits.includes(','), symbol: money[2] ?? money[5], symbolAfter: !!money[5] };
        }
    }
    return { kind: 'text' };
}

/** `text` cut to at most `max` UTF-16 units, at a whole character (never half of a surrogate pair). */
function cutTo(text: string, max: number): string {
    let out = '';
    for (const ch of text) { if (out.length + ch.length > max) break; out += ch; }
    return out;
}

/**
 * A sheet name a spreadsheet takes: at most 31 characters (cut at a whole character), none of
 * `[]:*?/\` nor what XML cannot hold, not blank, unique as a spreadsheet compares names (case aside).
 */
const nextSuffix = new WeakMap<Set<string>, Map<string, number>>();

export function sheetName(wanted: string, taken: Set<string>, fallback: string): string {
    const clean = (name: string) => cutTo(stripInvalidXmlChars(name).replace(/[[\]:*?/\\\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').trim().replace(/^'+|'+$/g, '').trim(), 31).trim();
    let base = clean(wanted) || clean(fallback) || 'Sheet';
    if (base.toLowerCase() === 'history') base = `${cutTo(base, 29)} 1`;
    let name = base;
    // The next number to try for each name, so many sheets wanting one name are named in time in proportion.
    let next = nextSuffix.get(taken);
    if (!next) nextSuffix.set(taken, next = new Map());
    let n = next.get(base.toLowerCase()) ?? 2;
    while (taken.has(name.toLowerCase())) { name = `${cutTo(base, 31 - String(n).length - 1).trim()} ${n}`; n++; }
    next.set(base.toLowerCase(), n);
    taken.add(name.toLowerCase());
    return name;
}

/**
 * The columns a sheet hides: those whose every cell in a row it shows is hidden. A parser marks a cell
 * hidden for its row or its column alike, so a hidden row's cells say nothing of their columns.
 */
export function hiddenColumns(sheet: WorkSheet): Set<number> {
    const shown = new Set<number>(), hidden = new Set<number>();
    for (const row of sheet.rows) {
        if (row.hidden) continue;
        for (const cell of row.cells) (cell.hidden ? hidden : shown).add(cell.col);
    }
    for (const col of shown) hidden.delete(col);
    return hidden;
}

/** Each column's width in characters, from its longest line of text (at least 8, at most 80). */
function widthsOf(rows: WorkRow[]): number[] {
    const widths: number[] = [];
    for (const row of rows) for (const cell of row.cells) {
        if (cell.colSpan > 1) continue;
        let longest = 0;
        for (const line of cell.shown.split('\n')) if (line.length > longest) longest = line.length;
        widths[cell.col] = Math.max(widths[cell.col] ?? 8, Math.min(80, longest + 2));
    }
    return Array.from({ length: widths.length }, (_, i) => widths[i] ?? 8);
}

/** What the model needs of a writer: a node's text, and whether a node is written at all (onNode). */
export interface WorkbookSource {
    textOf(node: OfficeContentNode): Promise<string>;
    keep(node: OfficeContentNode): Promise<boolean | string>;
    /** A chart's data as a table (its labels down, its series across), or undefined where it has none. */
    chartTable?(node: OfficeContentNode): OfficeContentNode | undefined;
}

/** A cell of `text` alone at `col`, typed from it. */
const textCell = (col: number, text: string, bold = false): WorkCell => ({ col, text, shown: text, value: cellValue(text), bold, colSpan: 1, rowSpan: 1 });

/** Whether all of a node's text is bold. */
function allBold(node: OfficeContentNode): boolean {
    let any = false, all = true;
    const walk = (n: OfficeContentNode, depth: number) => {
        if (n.type === 'text' && (n.text || '').trim()) { any = true; if (!n.formatting?.bold) all = false; }
        if (depth < 32) for (const c of n.children || []) walk(c, depth + 1);
    };
    walk(node, 0);
    return any && all;
}

/**
 * A table's or sheet's rows as a sheet's: its cells placed, typed, spans kept. A spreadsheet's cells are
 * where their own rows and columns put them (gaps kept as empty rows and columns); a table's, and a
 * sheet's whose cells name no columns, as DOCX and ODT lay a table out (see layoutTableRows), so a cell
 * merged down holds its columns in the rows below and no two merges overlap. A row onNode gives text
 * for is one cell holding it; a cell onNode gives text for holds that text.
 */
async function gridOf(node: OfficeContentNode, source: WorkbookSource, spreadsheet: boolean): Promise<WorkRow[]> {
    // The rows and cells kept, and the text onNode gave for any.
    const rows: { row: OfficeContentNode; cells: OfficeContentNode[]; text?: string }[] = [];
    const given = new Map<OfficeContentNode, string>();
    for (const row of node.children || []) {
        if (row.type !== 'row') continue;
        const kept = await source.keep(row);
        if (kept === false) continue;
        const cells: OfficeContentNode[] = [];
        if (typeof kept !== 'string') {
            for (const cell of row.children || []) {
                if (cell.type !== 'cell') continue;
                const k = await source.keep(cell);
                if (k === false) continue;
                if (typeof k === 'string') given.set(cell, k);
                cells.push(cell);
            }
        }
        rows.push({ row, cells, ...(typeof kept === 'string' ? { text: kept } : {}) });
    }
    const hiddenRow = (row: OfficeContentNode) => (row.metadata as RowMetadata | undefined)?.hidden ? { hidden: true } : {};
    const workCell = async (cell: OfficeContentNode, col: number, rowIndex: number, colSpan: number, rowSpan: number): Promise<WorkCell> => {
        const cm = cell.metadata as CellMetadata | undefined;
        const text = given.get(cell) ?? await source.textOf(cell);
        const shown = spreadsheet && typeof cm?.shownText === 'string' ? cm.shownText : undefined;
        return {
            col, text, shown: shown ?? text, value: cellValue(text, shown, spreadsheet), bold: cm?.style === 'header' || allBold(cell),
            colSpan: Math.max(1, Math.min(MAX_COLS - col, colSpan)), rowSpan: Math.max(1, Math.min(MAX_ROWS - rowIndex, rowSpan)),
            ...(cm?.hidden ? { hidden: true } : {}),
        };
    };
    const out: WorkRow[] = [];
    const placed = spreadsheet && rows.every(r => r.cells.every(c => typeof (c.metadata as CellMetadata | undefined)?.col === 'number'));
    if (placed) {
        let rowIndex = -1;
        for (const { row, cells: kept, text } of rows) {
            // A spreadsheet's row number, where its cells give it (gaps kept as empty rows), else the next.
            const own = (row.children || []).map(c => (c.metadata as CellMetadata | undefined)?.row).find((r): r is number => typeof r === 'number');
            rowIndex = typeof own === 'number' && own > rowIndex ? own : rowIndex + 1;
            if (rowIndex >= MAX_ROWS) break;
            while (out.length < rowIndex) out.push({ cells: [] });
            const cells: WorkCell[] = [];
            if (text !== undefined) {
                const first = (row.children || []).map(c => (c.metadata as CellMetadata | undefined)?.col).find((c): c is number => typeof c === 'number');
                cells.push(textCell(Math.min(MAX_COLS - 1, Math.max(0, first ?? 0)), text));
            }
            let col = -1;
            for (const cell of kept) {
                const cm = cell.metadata as CellMetadata;
                col = cm.col > col ? cm.col : col + 1;
                if (col >= MAX_COLS) break;
                const colSpan = Math.max(1, Math.floor(cm.colSpan ?? 1));
                cells.push(await workCell(cell, col, rowIndex, colSpan, Math.max(1, Math.floor(cm.rowSpan ?? 1))));
                col += colSpan - 1;
            }
            out.push({ cells, ...hiddenRow(row) });
        }
        return withinRows(out);
    }
    // Laid out as a table is: a row onNode gave text for is one cell, in the row's first column.
    const stand = new Map<OfficeContentNode, string>();
    const laidRows = rows.map(r => {
        if (r.text === undefined) return { type: 'row', children: r.cells } as OfficeContentNode;
        const cell: OfficeContentNode = { type: 'cell', text: r.text, children: [{ type: 'text', text: r.text }] };
        stand.set(cell, r.text);
        return { type: 'row', children: [cell] } as OfficeContentNode;
    });
    const layout = layoutTableRows(laidRows);
    for (const [ri, slots] of (layout?.rows ?? []).entries()) {
        if (ri >= MAX_ROWS) break;
        const cells: WorkCell[] = [];
        for (const slot of slots) {
            if (slot.kind !== 'cell' || slot.col >= MAX_COLS) continue;
            const text = stand.get(slot.cell);
            cells.push(text !== undefined ? textCell(slot.col, text) : await workCell(slot.cell, slot.col, ri, slot.colSpan, slot.rowSpan));
        }
        out.push({ cells, ...hiddenRow(rows[ri].row) });
    }
    return withinRows(out);
}

/** A table's rows with each merge held to them: one reaching past its last row overlapped what came next on the sheet. */
function withinRows(rows: WorkRow[]): WorkRow[] {
    rows.forEach((row, ri) => { for (const cell of row.cells) if (cell.rowSpan > rows.length - ri) cell.rowSpan = Math.max(1, rows.length - ri); });
    return rows;
}

/**
 * The numbers of a document's footnotes and endnotes (`Footnote 1`, `Endnote 1`), each kind in its order
 * through the document, as its text cites them (see cellTextProcessor) and the Notes sheet lists them.
 */
export function noteLabels(content: OfficeContentNode[]): Map<OfficeContentNode, string> {
    const labels = new Map<OfficeContentNode, string>();
    const counts = { footnote: 0, endnote: 0 };
    const label = (note: OfficeContentNode) => {
        if (labels.has(note)) return;
        const end = (note.metadata as { noteType?: string } | undefined)?.noteType === 'endnote';
        labels.set(note, end ? `Endnote ${++counts.endnote}` : `Footnote ${++counts.footnote}`);
    };
    const stack: { nodes: OfficeContentNode[]; at: number; depth: number }[] = [{ nodes: content, at: 0, depth: 0 }];
    while (stack.length) {
        const top = stack[stack.length - 1];
        if (top.at >= top.nodes.length) { stack.pop(); continue; }
        const node = top.nodes[top.at++];
        if (node.type === 'note') label(node);
        if (node.type !== 'slide') for (const note of node.notes || []) {
            label(note);
            if (top.depth < 32 && note.children?.length) stack.push({ nodes: note.children, at: 0, depth: top.depth + 1 });
        }
        if (top.depth < 32 && node.children?.length) stack.push({ nodes: node.children, at: 0, depth: top.depth + 1 });
    }
    return labels;
}

/**
 * The notes a document holds, as rows of a sheet: each footnote and endnote (cited in its text, or a
 * block of the page) by its number, each slide's speaker notes by the slide, and each comment by its
 * author, with their text.
 */
async function notesOf(content: OfficeContentNode[], source: WorkbookSource, labels: Map<OfficeContentNode, string>): Promise<WorkRow[]> {
    const rows: WorkRow[] = [];
    const seen = new Set<OfficeContentNode>();
    const add = async (note: OfficeContentNode, label: string) => {
        if (seen.has(note) || rows.length >= MAX_ROWS) return;
        seen.add(note);
        if (await source.keep(note) === false) return;
        // (A comment's blocks one by one: a comment itself, in a run's text, writes nothing there.)
        const text = (note.children?.length
            ? note.type === 'comment' ? (await Promise.all(note.children.map(c => source.textOf(c)))).join('\n') : await source.textOf(note)
            : note.text ?? '').trim();
        if (text) rows.push({ cells: [textCell(0, label), { ...textCell(1, text), value: { kind: 'text' } }] });
    };
    const commentLabel = (comment: OfficeContentNode) => {
        const author = (comment.metadata as CommentMetadata | undefined)?.author;
        return typeof author === 'string' && author.trim() ? `Comment (${author.trim()})` : 'Comment';
    };
    const visit = async (nodes: OfficeContentNode[], depth: number) => {
        for (const node of nodes) {
            if (seen.has(node)) continue;
            if (node.type === 'note') await add(node, labels.get(node) ?? 'Footnote');
            else if (node.type === 'comment') { await add(node, commentLabel(node)); continue; }
            else if (node.type === 'slide') {
                const slideNumber = (node.metadata as SlideMetadata | undefined)?.slideNumber;
                for (const note of node.notes || []) await add(note, `Slide ${slideNumber ?? ''}`.trim());
            } else for (const note of node.notes || []) {
                await add(note, labels.get(note) ?? 'Footnote');
                // A note cited inside a note.
                if (depth < 32 && note.children?.length) await visit(note.children, depth + 1);
            }
            for (const comment of node.comments || []) await add(comment, commentLabel(comment));
            if (depth < 32 && node.children?.length) await visit(node.children, depth + 1);
        }
    };
    await visit(content, 0);
    return rows;
}

/**
 * The document's workbook: a spreadsheet's sheets; or each table a sheet (`sheetPerTable`, else all on
 * one, an empty row between them), named after the heading before it, and the text outside tables on a
 * sheet of its own (`includeText`), a block a row. Pictures and charts, which a cell cannot hold, are
 * counted for the writer's warning.
 */
export async function buildWorkbook(content: OfficeContentNode[], source: WorkbookSource, options: { sheetPerTable: boolean; includeText: boolean; includeCharts: boolean; imageMode: ImageMode; includeFormatting: boolean; noteLabels: Map<OfficeContentNode, string> }): Promise<{ sheets: WorkSheet[]; unrepresentable: Set<string> }> {
    const sheets: WorkSheet[] = [];
    const taken = new Set<string>();
    const unrepresentable = new Set<string>();
    // A chart or picture left out by the caller (includeCharts: false, includeImages 'none' or 'ocr-text-only')
    // is not one lost.
    const picturesWanted = options.imageMode === 'image-only' || options.imageMode === 'image+ocr-text';
    const lost = (node: OfficeContentNode) => { if (node.type === 'image') { if (picturesWanted) unrepresentable.add('pictures'); } else if (node.type === 'chart' && options.includeCharts) unrepresentable.add('charts'); };
    // The notes, on a sheet of their own after the others.
    const withNotes = async () => {
        const notes = await notesOf(content, source, options.noteLabels);
        if (notes.length) sheets.push({ name: sheetName('Notes', taken, 'Notes'), rows: notes, widths: [12, Math.min(120, Math.max(8, longestOf(notes, 1)))] });
        // No bold where the caller asked for no formatting (includeFormatting: false).
        if (!options.includeFormatting) for (const sheet of sheets) for (const row of sheet.rows) for (const cell of row.cells) cell.bold = false;
        return { sheets, unrepresentable };
    };
    // A chart as a sheet of its data (its labels down, its series across), where it has data; else lost.
    let charts = 0;
    const chartRows = async (node: OfficeContentNode): Promise<WorkRow[] | undefined> => {
        if (!options.includeCharts) return undefined;
        const table = source.chartTable?.(node);
        if (!table) { lost(node); return undefined; }
        charts++;
        return gridOf(table, source, false);
    };
    const sheetNodes = content.filter(n => n.type === 'sheet');
    if (sheetNodes.length) {
        for (const [i, node] of sheetNodes.entries()) {
            if (await source.keep(node) === false) continue;
            const meta = node.metadata as SheetMetadata | undefined;
            const rows = await gridOf(node, source, true);
            const name = sheetName(meta?.sheetName ?? '', taken, `Sheet${i + 1}`);
            sheets.push({ name, rows, ...(meta?.hidden ? { hidden: true } : {}), widths: widthsOf(rows) });
            // Its charts, each a sheet of its data after it; its pictures, which a cell cannot hold, reported.
            for (const c of node.children || []) {
                if (c.type !== 'chart') { lost(c); continue; }
                if (await source.keep(c) === false) continue;
                const chart = await chartRows(c);
                if (chart) sheets.push({ name: sheetName(`${name} Chart ${charts}`, taken, `Chart ${charts}`), rows: chart, widths: widthsOf(chart) });
            }
        }
        return withNotes();
    }
    // A document: its tables, and its text between them.
    const text: WorkRow[] = [];
    const together: WorkRow[] = [];
    let heading = '', tables = 0;
    const visit = async (nodes: OfficeContentNode[], depth: number) => {
        for (const node of nodes) {
            const kept = await source.keep(node);
            if (kept === false) continue;
            if (node.type === 'table') {
                const rows = await gridOf(node, source, false);
                tables++;
                if (options.sheetPerTable) sheets.push({ name: sheetName(heading, taken, `Table ${tables}`), rows, widths: widthsOf(rows) });
                else { if (together.length) together.push({ cells: [] }); for (const r of rows) together.push(r); }
                continue;
            }
            if (node.type === 'image') {
                // A picture a cell cannot hold (reported): its words in its place (see pictureWords).
                const words = pictureWords(node, options.imageMode);
                if (words && options.includeText && text.length < MAX_ROWS) text.push({ cells: [{ ...textCell(0, words), value: { kind: 'text' } }] });
                lost(node);
                continue;
            }
            if (node.type === 'chart') {
                const rows = await chartRows(node);
                if (!rows) continue;
                if (options.sheetPerTable) sheets.push({ name: sheetName(heading ? `${heading} Chart` : '', taken, `Chart ${charts}`), rows, widths: widthsOf(rows) });
                else { if (together.length) together.push({ cells: [] }); for (const r of rows) together.push(r); }
                continue;
            }
            if (node.type === 'drawing') continue;
            // An embedded video or page: its address.
            if (node.type === 'embed') {
                const url = (node.metadata as { url?: string } | undefined)?.url ?? node.text;
                if (typeof url === 'string' && url.trim() && options.includeText && text.length < MAX_ROWS) text.push({ cells: [{ ...textCell(0, url.trim()), value: { kind: 'text' } }] });
                continue;
            }
            // Comments are written on the Notes sheet (see notesOf).
            if (node.type === 'comment') continue;
            // Notes are written on their own sheet (see notesOf).
            if (node.type === 'note') continue;
            const blockText = typeof kept === 'string' ? kept : null;
            if (['paragraph', 'heading', 'list', 'code', 'definitionTerm', 'definitionDescription'].includes(node.type)) {
                let t = (blockText ?? await source.textOf(node)).trim();
                if (node.type === 'heading' && t) heading = t;
                // A list item with its marker as the document shows it.
                if (node.type === 'list' && t && blockText === null) { const marker = listMarker(node.metadata as ListMetadata | undefined); if (marker) t = `${marker} ${t}`; }
                if (t && options.includeText && text.length < MAX_ROWS) text.push({ cells: [{ ...textCell(0, t, node.type === 'heading'), value: { kind: 'text' } }] });
                // A picture in the text, which a cell cannot hold.
                const walk = (n: OfficeContentNode, d: number) => { if (n.type === 'image' && picturesWanted) unrepresentable.add('pictures'); if (d < 16) for (const c of n.children || []) walk(c, d + 1); };
                walk(node, 0);
                continue;
            }
            if (depth < 32 && node.children?.length) await visit(node.children, depth + 1);
        }
    };
    await visit(content, 0);
    if (!options.sheetPerTable && together.length) sheets.push({ name: sheetName('Tables', taken, 'Tables'), rows: together, widths: widthsOf(together) });
    if (text.length) sheets.push({ name: sheetName('Text', taken, 'Text'), rows: text, widths: [Math.min(120, Math.max(8, longestOf(text, 0)))] });
    return withNotes();
}

/** The longest text of the cells at `col` of `rows`. */
function longestOf(rows: WorkRow[], col: number): number {
    let longest = 0;
    for (const r of rows) for (const c of r.cells) if (c.col === col && c.text.length > longest) longest = c.text.length;
    return longest;
}

/**
 * A node's text as a cell holds it, for the spreadsheet writers' traversal: its runs, a line for each
 * block and break, and a picture's recognized (OCR) text where the caller asks for it (includeImages
 * 'ocr-text-only' or 'image+ocr-text'); a picture itself a cell cannot hold.
 */
export function cellTextProcessor(imageMode: ImageMode, labels?: Map<OfficeContentNode, string>): (n: OfficeContentNode, childrenOutput: string) => string {
    const ocr = imageMode === 'ocr-text-only' || imageMode === 'image+ocr-text';
    // The notes a run cites, by their numbers on the Notes sheet.
    const cites = (n: OfficeContentNode) => (n.notes || []).map(note => labels?.get(note)).filter(Boolean).map(l => ` [${l}]`).join('');
    return (n, co) => {
        if (n.type === 'text' || n.type === 'code') return (n.text || '') + cites(n);
        if (n.type === 'embed') return ((n.metadata as { url?: string } | undefined)?.url ?? n.text ?? '').trim();
        if (n.type === 'comment') return '';
        if (n.type === 'image') return pictureWords(n, imageMode);
        if (n.type === 'break') return (n.metadata as { breakType?: string } | undefined)?.breakType === 'section' ? '' : '\n';
        if ((n.type === 'paragraph' || n.type === 'heading' || n.type === 'list') && co) return co + '\n';
        return co;
    };
}

/** A list item's marker as the document shows it: its own, else its number or a bullet. */
function listMarker(meta: ListMetadata | undefined): string {
    if (typeof meta?.marker === 'string') return meta.marker.replace(/\s+/g, ' ').trim();
    return meta?.listType === 'ordered' ? `${(typeof meta.itemIndex === 'number' && Number.isFinite(meta.itemIndex) ? Math.max(0, Math.trunc(meta.itemIndex)) : 0) + 1}.` : '•';
}

/**
 * A picture's words, which a cell holds where it cannot hold the picture: its alternative text (and its
 * recognized text, under includeImages 'image+ocr-text'), its recognized text alone under 'ocr-text-only',
 * none under 'none'.
 */
function pictureWords(node: OfficeContentNode, imageMode: ImageMode): string {
    const m = node.metadata as { altText?: string; ocrText?: string } | undefined;
    const alt = typeof m?.altText === 'string' ? m.altText.trim() : '', ocr = typeof m?.ocrText === 'string' ? m.ocrText.trim() : '';
    if (imageMode === 'none') return '';
    if (imageMode === 'ocr-text-only') return ocr;
    if (imageMode === 'image+ocr-text') return [alt, ocr].filter(Boolean).join(' ');
    return alt;
}

/**
 * A document as a workbook, for the spreadsheet writers (XLSX, ODS): its sheets as they are (a
 * spreadsheet's), or each of its tables as a sheet (or all on one), and its text outside them on a sheet
 * of its own; each cell typed as a spreadsheet types it, a number as a number with the format it is
 * shown in (a percentage, a currency, grouped thousands, its decimals), a date or a time as one, other
 * text as text. A spreadsheet's cell keeps the value and the kind of value it holds (CellMetadata.valueType,
 * CellMetadata.value) and is shown as it was (CellMetadata.shownText); a cell whose source does not say (a
 * table's, a CSV's) is read from its text (`19%`, `€1,250.00`, `2026-10-01`), a number with leading zeros
 * or more digits than a spreadsheet keeps (`007`, a card number) staying text.
 *
 * @module generators/sheets/workbook
 */

import { CellMetadata, CommentMetadata, ImageMode, ListMetadata, NodeBounds, OfficeContentNode, RowMetadata, SheetMetadata, SlideMetadata } from '../../types.js';
import { sanitizeOfficePackageUrl, stripInvalidXmlChars } from '../../utils/sanitize.js';
import { encUrl } from '../../utils/officeGenUtils.js';
import { layoutTableRows } from '../../utils/tableLayout.js';

/** How a cell's value is held and shown. */
export interface CellValue {
    kind: 'number' | 'percent' | 'currency' | 'date' | 'datetime' | 'time' | 'boolean' | 'error' | 'text';
    /** The number held: a number's, a percentage's fraction, a date's or time's day count (1900 system) or day fraction (a duration's days), a truth value's 1 or 0. */
    number?: number;
    /** A date's or time's ISO text, as ODF holds it (a time's as a duration, `PT30H0M0S`). */
    iso?: string;
    /** Decimals shown, thousands grouped, and a currency's symbol (before or after the number). */
    decimals?: number;
    grouping?: boolean;
    symbol?: string;
    symbolAfter?: boolean;
    /** The spreadsheet's own number format code (CellMetadata.numberFormat), which XLSX output writes as it is. */
    code?: string;
}

/** A cell: its text (a spreadsheet's stored value), what it shows (its value in its format), its value typed. */
export interface WorkCell { col: number; text: string; shown: string; value: CellValue; bold: boolean; colSpan: number; rowSpan: number; hidden?: boolean; /** Its links in order (a few at most): a writer writes the first it can (see writableLink). */ links?: string[] }
export interface WorkRow { cells: WorkCell[]; hidden?: boolean }
export interface WorkSheet { name: string; rows: WorkRow[]; hidden?: boolean; widths: number[]; /** A spreadsheet's sheet's own name, where the name written is another (cut, made unique): links name it. */ source?: string }

/** Spreadsheets' limits: rows, columns, a sheet name's length. */
export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;
/** What is reported of content past them. */
const PAST_ROWS = 'rows past a sheet\'s 1,048,576', PAST_COLS = 'columns past a sheet\'s 16,384';

const NUMBER = /^-?(?:\d+(?:\.\d+)?|\.\d+)$/;
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

/** The errors a spreadsheet cell holds, as Excel writes them. */
const ERRORS = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA']);

/** A day count (1900 system) as ISO text: a date, or a date and time. */
function isoOfDays(days: number, withTime: boolean): string {
    const ms = Math.round(days * 86_400_000 / 1000) * 1000;
    const at = new Date(Date.UTC(1899, 11, 30) + ms).toISOString();
    return withTime ? at.slice(0, 19) : at.slice(0, 10);
}

/**
 * A spreadsheet's cell as the value its file says it holds (CellMetadata.valueType: `held`, the value
 * itself, CellMetadata.value or its text), shown as it showed (`shown`): a number in the format read from
 * what it showed, a date or a time as one, a truth value, an error, or text however it reads.
 */
export function typedValue(valueType: NonNullable<CellMetadata['valueType']>, held: string, shown: string, code?: string): CellValue {
    const v = typedKind(valueType, held, shown);
    return code && v.kind !== 'text' && v.kind !== 'boolean' && v.kind !== 'error' ? { ...v, code } : v;
}

function typedKind(valueType: NonNullable<CellMetadata['valueType']>, held: string, shown: string): CellValue {
    const n = Number(held.trim());
    switch (valueType) {
        case 'number': {
            if (!held.trim() || !Number.isFinite(n)) return { kind: 'text' };
            const look = cellValue(shown);
            if (look.kind === 'percent' || look.kind === 'currency' || (look.kind === 'number' && (look.grouping || look.decimals))) return { ...look, number: n };
            return { kind: 'number', number: n };
        }
        case 'date': {
            // (Before 1900-03-01 the 1900 system counts a day that never was: such a date stays its text.)
            if (!held.trim() || !Number.isFinite(n) || n < 61 || n >= 2_958_466) return { kind: 'text' };
            const whole = Math.abs(n - Math.round(n)) < 1e-9;
            return whole ? { kind: 'date', number: Math.round(n), iso: isoOfDays(n, false) } : { kind: 'datetime', number: n, iso: isoOfDays(n, true) };
        }
        case 'time': {
            if (!held.trim() || !Number.isFinite(n) || n < 0) return { kind: 'text' };
            const seconds = Math.round(n * 86400);
            return { kind: 'time', number: n, iso: `PT${Math.floor(seconds / 3600)}H${Math.floor(seconds / 60) % 60}M${seconds % 60}S` };
        }
        case 'boolean': {
            const truth = /^(true|1)$/i.test(held.trim()) ? 1 : /^(false|0)$/i.test(held.trim()) ? 0 : undefined;
            return truth === undefined ? { kind: 'text' } : { kind: 'boolean', number: truth };
        }
        case 'error': return ERRORS.has(held.trim().toUpperCase()) ? { kind: 'error' } : { kind: 'text' };
        default: return { kind: 'text' };
    }
}

/**
 * A cell's value from its text (a table's, a CSV's: its source says nothing of it), and from what a
 * spreadsheet showed for it where its text is a value held (`shown`): a number of at most fifteen
 * significant digits, a percentage, a currency amount, a date or a time, else text.
 */
export function cellValue(text: string, shown?: string): CellValue {
    const s = text.trim();
    if (!s || s.length > 64) return { kind: 'text' };
    const iso = isoValue(s);
    if (iso) return iso;
    // A number shown through a format: the value held, the format read from what it showed.
    if (shown !== undefined && shown !== text && (NUMBER.test(s) || GROUPED.test(s))) {
        const n = plainNumber(s);
        if (n !== undefined) {
            const look = cellValue(shown);
            if (look.kind === 'percent' || look.kind === 'currency' || (look.kind === 'number' && (look.grouping || look.decimals))) return { ...look, number: n };
            return { kind: 'number', number: n };
        }
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
 * A sheet name a spreadsheet takes: at most `max` characters (Excel's 31; cut at a whole character), none
 * of `[]:*?/\` nor what XML cannot hold, not blank, unique as a spreadsheet compares names (case aside).
 */
const nextSuffix = new WeakMap<Set<string>, Map<string, number>>();

export function sheetName(wanted: string, taken: Set<string>, fallback: string, max = 31): string {
    // (Neither end an apostrophe nor a blank, after every cut too: a name cut to 31 could end in one.)
    const ends = (text: string) => text.replace(/^[\s']+|[\s']+$/g, '');
    const clean = (name: string) => ends(cutTo(ends(stripInvalidXmlChars(name).replace(/[[\]:*?/\\\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ')), max));
    let base = clean(wanted) || clean(fallback) || 'Sheet';
    if (base.toLowerCase() === 'history') base = `${cutTo(base, max - 2)} 1`;
    let name = base;
    // The next number to try for each name, so many sheets wanting one name are named in time in proportion.
    let next = nextSuffix.get(taken);
    if (!next) nextSuffix.set(taken, next = new Map());
    let n = next.get(base.toLowerCase()) ?? 2;
    while (taken.has(name.toLowerCase())) { name = `${ends(cutTo(base, max - String(n).length - 1)) || 'Sheet'} ${n}`; n++; }
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
    /** A chart's title, where it has one. */
    chartTitle?(node: OfficeContentNode): string | undefined;
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
async function gridOf(node: OfficeContentNode, source: WorkbookSource, spreadsheet: boolean, lost: Set<string>, media?: CellMedia): Promise<WorkRow[]> {
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
        const own = given.get(cell);
        const text = own ?? await source.textOf(cell);
        const shown = spreadsheet && typeof cm?.shownText === 'string' ? cm.shownText : undefined;
        // The kind of value its file says it holds, and the value (its text, else what it shows; see CellMetadata.value).
        const type = spreadsheet && typeof cm?.valueType === 'string' ? cm.valueType : undefined;
        const held = typeof cm?.value === 'string' ? cm.value : text;
        // Its pictures and charts, which a cell cannot hold (see CellMedia), and its first link.
        if (own === undefined && media) await mediaIn(cell, source, media);
        const links = own === undefined ? await linksIn(cell, source.keep) : [];
        return {
            col, text, shown: shown ?? text, value: type ? typedValue(type, held, shown ?? text, typeof cm?.numberFormat === 'string' ? cm.numberFormat : undefined) : cellValue(text, shown), bold: cm?.style === 'header' || allBold(cell),
            colSpan: Math.max(1, Math.min(MAX_COLS - col, colSpan)), rowSpan: Math.max(1, Math.min(MAX_ROWS - rowIndex, rowSpan)),
            ...(cm?.hidden ? { hidden: true } : {}), ...(links.length ? { links } : {}),
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
            // (Rows and columns past a spreadsheet's are reported, not written.)
            if (rowIndex >= MAX_ROWS) { lost.add(PAST_ROWS); break; }
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
                if (col >= MAX_COLS) { lost.add(PAST_COLS); break; }
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
        if (ri >= MAX_ROWS) { lost.add(PAST_ROWS); break; }
        const cells: WorkCell[] = [];
        for (const slot of slots) {
            if (slot.kind !== 'cell') continue;
            if (slot.col >= MAX_COLS) { lost.add(PAST_COLS); continue; }
            const text = stand.get(slot.cell);
            cells.push(text !== undefined ? textCell(slot.col, text) : await workCell(slot.cell, slot.col, ri, slot.colSpan, slot.rowSpan));
        }
        out.push({ cells, ...hiddenRow(rows[ri].row) });
    }
    return withinRows(out);
}

/**
 * A table's rows with each merge held to them (one reaching past its last row overlapped what came next on
 * the sheet) and to the area no earlier merge covers: a sheet whose cells name their places can give two
 * merges that overlap, which a spreadsheet repairs by dropping them. A merge that would overlap is none.
 * Bounded: past MAX_MERGE_POSITIONS positions marked, later merges are none.
 */
function withinRows(rows: WorkRow[]): WorkRow[] {
    const taken = new Set<string>();
    let marked = 0;
    rows.forEach((row, ri) => {
        for (const cell of row.cells) {
            if (cell.rowSpan > rows.length - ri) cell.rowSpan = Math.max(1, rows.length - ri);
            if (cell.rowSpan === 1 && cell.colSpan === 1) continue;
            const area = cell.rowSpan * cell.colSpan;
            let free = marked + area <= MAX_MERGE_POSITIONS;
            for (let r = ri; free && r < ri + cell.rowSpan; r++) for (let c = cell.col; free && c < cell.col + cell.colSpan; c++) if (taken.has(`${r},${c}`)) free = false;
            if (!free) { cell.rowSpan = 1; cell.colSpan = 1; continue; }
            for (let r = ri; r < ri + cell.rowSpan; r++) for (let c = cell.col; c < cell.col + cell.colSpan; c++) taken.add(`${r},${c}`);
            marked += area;
        }
    });
    return rows;
}

/** The most positions merges may cover on one sheet's check for overlaps. */
const MAX_MERGE_POSITIONS = 2_000_000;

/** The pictures and charts met in cells (a cell holds neither), and whether a cell held more than one link. */
export interface CellMedia { pictures: number; charts: OfficeContentNode[] }

/** The pictures and charts under a node, each asked of onNode: one left out, or replaced by a string, is none. */
async function mediaIn(node: OfficeContentNode, source: WorkbookSource, media: CellMedia, depth = 0): Promise<void> {
    for (const c of node.children || []) {
        if (c.type !== 'image' && c.type !== 'chart' && !c.children?.length) continue;
        const kept = await source.keep(c);
        if (kept === false || typeof kept === 'string') continue;
        if (c.type === 'image') media.pictures++;
        else if (c.type === 'chart') media.charts.push(c);
        if (depth < 32) await mediaIn(c, source, media, depth + 1);
    }
}

/**
 * Whether a link is one a cell takes: a place in the document (`#Sheet!A1`, `#bookmark`), or an address a
 * package holds (a web, mail or phone address, or one relative to the file), as DOCX output takes them.
 */
const linkable = (url: string) => url.startsWith('#') ? url.length > 1 : sanitizeOfficePackageUrl(url) !== '';

/**
 * A workbook's sheets by their names as a spreadsheet compares them (case aside), for finding the places
 * links name: by the names written, and by a spreadsheet's own where it was written under another (cut,
 * made unique), whose links name it so.
 */
export function sheetIndex(sheets: readonly { name: string; source?: string }[]): Map<string, string> {
    const index = new Map<string, string>();
    for (const s of sheets) if (s.source) index.set(s.source.toLowerCase(), s.name);
    // (A name written is that sheet's, whatever sheet was once called so.)
    for (const s of sheets) index.set(s.name.toLowerCase(), s.name);
    return index;
}

/** A place on a sheet: its sheet, its first cell, and its last where it names a range. */
export interface SheetPlace { sheet: string; cell: string; to?: string }

/**
 * A place in the workbook a link names (`#Sheet2!B4`, `#'Q1 sales'!A1`, a range `#Sheet2!A1:C3`, ODF's
 * `#Sheet2.B4` and `#$Sheet2.$A$1`, a sheet alone `#Sheet2`) on a sheet written, as that sheet's name and
 * the cell (A1 where it names none); undefined for a link elsewhere, or to a sheet no longer there (a
 * document's bookmark, a defined name).
 */
export function placeOf(link: string, sheets: ReadonlyMap<string, string>): SheetPlace | undefined {
    if (!link.startsWith('#') || link.length > 512) return undefined;
    // Excel's form (`#Sheet!A1`) or ODF's (`#Sheet.A1`, its name's own dots kept: the last one is the cell's).
    const m = /^#\$?(?:'((?:[^']|'')+)'|([^!]+?))(?:[!.]\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::(?:\$?(?:'(?:[^']|'')+'|[^!.:]+)[!.])?\$?([A-Za-z]{1,3})\$?(\d{1,7}))?)?$/.exec(link);
    if (!m) return undefined;
    const sheet = sheets.get((m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]).toLowerCase());
    if (!sheet) return undefined;
    return { sheet, cell: m[3] ? `${m[3].toUpperCase()}${m[4]}` : 'A1', ...(m[5] ? { to: `${m[5].toUpperCase()}${m[6]}` } : {}) };
}

/**
 * The first of a cell's links a workbook can hold (an address of at most `maxLength` characters as it is
 * written, XLSX's MAX_LINK_LENGTH as Excel opens it; or a place on a sheet written), and why each other
 * was not written, every one reported: one past the cell's first, a place no sheet holds, an address too long.
 */
export function writableLink(links: readonly string[] | undefined, sheets: ReadonlyMap<string, string>, maxLength = Infinity): { link?: string; place?: SheetPlace; lost: string[] } {
    let first: { link: string; place?: SheetPlace } | undefined;
    const lost = new Set<string>();
    for (const link of links ?? []) {
        let place: SheetPlace | undefined;
        if (link.startsWith('#')) {
            place = placeOf(link, sheets);
            if (!place) { lost.add('links to places outside the workbook\'s sheets'); continue; }
        } else if (encUrl(link).length > maxLength) { lost.add('links of more than 2,079 characters'); continue; }
        if (first) lost.add('links past a cell\'s first');
        else first = { link, ...(place ? { place } : {}) };
    }
    return { ...(first ?? {}), lost: [...lost] };
}

/** The longest address Excel opens a link to, and the most links a sheet holds. */
export const MAX_LINK_LENGTH = 2079;
export const MAX_SHEET_LINKS = 65_530;

/** A sheet's name in a reference to a place on it: quoted where it is more than a plain word, or reads as a cell or a truth. */
export function sheetRef(name: string): string {
    const plain = /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^R\d*C\d*$/i.test(name) && !/^(true|false)$/i.test(name);
    return plain ? name : `'${name.replace(/'/g, "''")}'`;
}

/**
 * A node's distinct links (web, mail, phone, a place), in order, a few at most (the writer takes the first
 * it can hold, see writableLink); a part onNode leaves out or replaces has none (its link went with it).
 */
async function linksIn(node: OfficeContentNode, keep: (node: OfficeContentNode) => Promise<boolean | string>): Promise<string[]> {
    const out: string[] = [];
    const walk = async (n: OfficeContentNode, depth: number) => {
        for (const c of n.children || []) {
            if (out.length >= 8) return;
            const k = await keep(c);
            if (k === false || typeof k === 'string') continue;
            const url = (c.metadata as { link?: unknown } | undefined)?.link;
            if (typeof url === 'string' && linkable(url.trim()) && !out.includes(url.trim())) out.push(url.trim());
            if (depth < 32) await walk(c, depth + 1);
        }
    };
    const own = (node.metadata as { link?: unknown } | undefined)?.link;
    if (typeof own === 'string' && linkable(own.trim())) out.push(own.trim());
    await walk(node, 0);
    return out;
}

/**
 * The numbers of a document's footnotes and endnotes (`Footnote 1`, `Endnote 1`), each kind in its order
 * through the document, as its text cites them (see cellTextProcessor) and the Notes sheet lists them.
 */
export async function noteLabels(content: OfficeContentNode[], keep: (node: OfficeContentNode) => Promise<boolean | string>): Promise<Map<OfficeContentNode, string>> {
    const labels = new Map<OfficeContentNode, string>();
    const counts = { footnote: 0, endnote: 0 };
    // (A note onNode leaves out is neither numbered nor cited.)
    const found: OfficeContentNode[] = [];
    const label = (note: OfficeContentNode) => { if (!labels.has(note)) { labels.set(note, ''); found.push(note); } };
    const stack: { nodes: OfficeContentNode[]; at: number; depth: number }[] = [{ nodes: content, at: 0, depth: 0 }];
    while (stack.length) {
        const top = stack[stack.length - 1];
        if (top.at >= top.nodes.length) { stack.pop(); continue; }
        const node = top.nodes[top.at++];
        // (The notes of a part onNode leaves out or replaces went with it.)
        if (node.type !== 'note') { const k = await keep(node); if (k === false || typeof k === 'string') continue; }
        if (node.type === 'note') label(node);
        if (node.type !== 'slide') for (const note of node.notes || []) {
            label(note);
            if (top.depth < 32 && note.children?.length) stack.push({ nodes: note.children, at: 0, depth: top.depth + 1 });
        }
        if (top.depth < 32 && node.children?.length) stack.push({ nodes: node.children, at: 0, depth: top.depth + 1 });
    }
    for (const note of found) {
        if (await keep(note) === false) { labels.delete(note); continue; }
        const end = (note.metadata as { noteType?: string } | undefined)?.noteType === 'endnote';
        labels.set(note, end ? `Endnote ${++counts.endnote}` : `Footnote ${++counts.footnote}`);
    }
    return labels;
}

/**
 * The notes a document holds, as rows of a sheet: each footnote and endnote (cited in its text, or a
 * block of the page) by its number, each slide's speaker notes by the slide, and each comment by its
 * author, with their text.
 */
async function notesOf(content: OfficeContentNode[], source: WorkbookSource, labels: Map<OfficeContentNode, string>, lost: Set<string>): Promise<WorkRow[]> {
    const rows: WorkRow[] = [];
    const seen = new Set<OfficeContentNode>();
    const add = async (note: OfficeContentNode, label: string) => {
        if (seen.has(note)) return;
        if (rows.length >= MAX_ROWS) { lost.add(PAST_ROWS); return; }
        seen.add(note);
        const kept = await source.keep(note);
        if (kept === false) return;
        // (A comment's blocks one by one: a comment itself, in a run's text, writes nothing there.)
        const text = (typeof kept === 'string' ? kept : note.children?.length
            ? note.type === 'comment' ? (await Promise.all(note.children.map(c => source.textOf(c)))).join('\n') : await source.textOf(note)
            : note.text ?? '').trim();
        // (Its first link the cell's, as a block's is on the Text sheet; a note onNode replaced has none.)
        const links = typeof kept === 'string' ? [] : await linksIn(note, source.keep);
        if (text) rows.push({ cells: [textCell(0, label), { ...textCell(1, text), value: { kind: 'text' }, ...(links.length ? { links } : {}) }] });
    };
    // (A comment onNode replaced is its text alone, no author.)
    const commentLabel = async (comment: OfficeContentNode) => {
        if (typeof await source.keep(comment) === 'string') return 'Comment';
        const author = (comment.metadata as CommentMetadata | undefined)?.author;
        return typeof author === 'string' && author.trim() ? `Comment (${author.trim()})` : 'Comment';
    };
    const visit = async (nodes: OfficeContentNode[], depth: number) => {
        for (const node of nodes) {
            if (seen.has(node)) continue;
            // A part onNode leaves out or replaces takes its notes and comments with it.
            if (node.type !== 'note' && node.type !== 'comment') { const k = await source.keep(node); if (k === false || typeof k === 'string') continue; }
            if (node.type === 'note') await add(node, labels.get(node) ?? 'Footnote');
            else if (node.type === 'comment') { await add(node, await commentLabel(node)); continue; }
            else if (node.type === 'slide') {
                const slideNumber = (node.metadata as SlideMetadata | undefined)?.slideNumber;
                for (const note of node.notes || []) await add(note, `Slide ${slideNumber ?? ''}`.trim());
            } else for (const note of node.notes || []) {
                await add(note, labels.get(note) ?? 'Footnote');
                // A note cited inside a note.
                if (depth < 32 && note.children?.length) await visit(note.children, depth + 1);
            }
            for (const comment of node.comments || []) await add(comment, await commentLabel(comment));
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
export async function buildWorkbook(content: OfficeContentNode[], source: WorkbookSource, options: { sheetPerTable: boolean; includeText: boolean; includeCharts: boolean; imageMode: ImageMode; includeFormatting: boolean; noteLabels: Map<OfficeContentNode, string>; /** The longest sheet name the format takes (Excel's 31). */ sheetNameLength?: number }): Promise<{ sheets: WorkSheet[]; unrepresentable: Set<string> }> {
    const sheets: WorkSheet[] = [];
    const taken = new Set<string>();
    const nameOf = (wanted: string, fallback: string) => sheetName(wanted, taken, fallback, options.sheetNameLength);
    const unrepresentable = new Set<string>();
    // A chart or picture left out by the caller (includeCharts: false, includeImages 'none' or 'ocr-text-only')
    // is not one lost.
    const picturesWanted = options.imageMode === 'image-only' || options.imageMode === 'image+ocr-text';
    const lost = (node: OfficeContentNode) => { if (node.type === 'image') { if (picturesWanted) unrepresentable.add('pictures'); } else if (node.type === 'chart' && options.includeCharts) unrepresentable.add('charts'); };
    // What cells and text held that a cell cannot: their pictures reported, their charts sheets of their data.
    const media: CellMedia = { pictures: 0, charts: [] };
    const settle = async (after: string) => {
        if (media.pictures && picturesWanted) unrepresentable.add('pictures');
        const charts = media.charts.splice(0);
        media.pictures = 0;
        for (const c of charts) {
            const rows = await chartRows(c);
            if (rows) sheets.push({ name: nameOf(chartName(c, after), `Chart ${chartCount}`), rows, widths: widthsOf(rows) });
        }
    };
    // A chart's sheet is named for its title, else for what it follows.
    const chartName = (node: OfficeContentNode, after: string) => {
        const title = (source.chartTitle?.(node) ?? '').trim();
        return title || (after ? `${after} Chart` : '');
    };
    // The notes, on a sheet of their own after the others.
    const withNotes = async () => {
        const notes = await notesOf(content, source, options.noteLabels, unrepresentable);
        if (notes.length) sheets.push({ name: nameOf('Notes', 'Notes'), rows: notes, widths: [12, Math.min(120, Math.max(8, longestOf(notes, 1)))] });
        // No bold where the caller asked for no formatting (includeFormatting: false).
        if (!options.includeFormatting) for (const sheet of sheets) for (const row of sheet.rows) for (const cell of row.cells) cell.bold = false;
        return { sheets, unrepresentable };
    };
    // A chart as a sheet of its data (its labels down, its series across), where it has data; else lost.
    let chartCount = 0;
    const chartRows = async (node: OfficeContentNode): Promise<WorkRow[] | undefined> => {
        if (!options.includeCharts) return undefined;
        const table = source.chartTable?.(node);
        if (!table) { lost(node); return undefined; }
        chartCount++;
        return gridOf(table, source, false, unrepresentable);
    };
    // A node onNode gives a string for, where a table or sheet stood: one cell of it.
    const oneCell = (text: string): WorkRow[] => [{ cells: [{ ...textCell(0, text), value: { kind: 'text' } }] }];
    // Whether a sheet has a row left (one past its last is reported, not written).
    const room = (rows: WorkRow[]) => rows.length < MAX_ROWS || (unrepresentable.add(PAST_ROWS), false);
    const sheetNodes = content.filter(n => n.type === 'sheet');
    if (sheetNodes.length) {
        for (const [i, node] of sheetNodes.entries()) {
            const kept = await source.keep(node);
            if (kept === false) continue;
            const meta = node.metadata as SheetMetadata | undefined;
            const rows = typeof kept === 'string' ? oneCell(kept) : await gridOf(node, source, true, unrepresentable, media);
            // (A sheet onNode gives a string for is that string alone, under no name of its own.)
            const name = nameOf(typeof kept === 'string' ? '' : meta?.sheetName ?? '', `Sheet${i + 1}`);
            const own = typeof kept === 'string' ? undefined : meta?.sheetName;
            sheets.push({ name, rows, ...(meta?.hidden ? { hidden: true } : {}), widths: widthsOf(rows), ...(own && own !== name ? { source: own } : {}) });
            if (typeof kept === 'string') continue;
            // Its charts and pictures (on it, or in its cells): each chart a sheet of its data after it; its
            // pictures, which a cell cannot hold, reported.
            for (const c of node.children || []) {
                if (c.type !== 'chart' && c.type !== 'image') continue;
                const k = await source.keep(c);
                if (k === false) continue;
                // (One onNode gives a string for: a row of that text on the sheet, after its own.)
                if (typeof k === 'string') { if (k.trim() && room(rows)) rows.push({ cells: [{ ...textCell(0, k), value: { kind: 'text' } }] }); continue; }
                if (c.type === 'image') media.pictures++; else media.charts.push(c);
            }
            await settle(name);
        }
        // Blocks beside the sheets (a heading, a paragraph): read as a document's are, below.
        if (!content.some(n => n.type !== 'sheet')) return withNotes();
    }
    // A document: its tables, and its text between them.
    const text: WorkRow[] = [];
    const together: WorkRow[] = [];
    let heading = '', tables = 0;
    const visit = async (nodes: OfficeContentNode[], depth: number) => {
        // The box of a shape onNode left out or replaced, whose text (the blocks after it with its bounds) went with it.
        let goneShape: NodeBounds | undefined;
        for (const node of nodes) {
            if (goneShape && node.bounds && ['paragraph', 'heading', 'list'].includes(node.type) && sameBounds(node.bounds, goneShape)) continue;
            goneShape = undefined;
            const kept = await source.keep(node);
            if (node.type === 'drawing' && kept !== true && node.bounds) goneShape = node.bounds;
            if (kept === false) continue;
            // A table or chart onNode gives a string for: a table of that one cell; a picture, an embed or any
            // other block: a row of the text.
            if (typeof kept === 'string' && !['paragraph', 'heading', 'list', 'code', 'definitionTerm', 'definitionDescription'].includes(node.type)) {
                if (node.type === 'table' || node.type === 'chart') {
                    tables++;
                    if (options.sheetPerTable) sheets.push({ name: nameOf(heading, `Table ${tables}`), rows: oneCell(kept), widths: [Math.min(80, kept.length + 2)] });
                    else { if (together.length) together.push({ cells: [] }); together.push(...oneCell(kept)); }
                } else if (kept.trim() && options.includeText && room(text)) text.push(...oneCell(kept));
                continue;
            }
            if (node.type === 'table') {
                const rows = await gridOf(node, source, false, unrepresentable, media);
                tables++;
                if (options.sheetPerTable) sheets.push({ name: nameOf(heading, `Table ${tables}`), rows, widths: widthsOf(rows) });
                else { if (together.length) together.push({ cells: [] }); for (const r of rows) together.push(r); }
                await settle(heading);
                continue;
            }
            if (node.type === 'image') {
                // A picture a cell cannot hold (reported): its words in its place (see pictureWords).
                const words = pictureWords(node, options.imageMode);
                if (words && options.includeText && room(text)) text.push({ cells: [{ ...textCell(0, words), value: { kind: 'text' } }] });
                lost(node);
                continue;
            }
            if (node.type === 'chart') {
                const rows = await chartRows(node);
                if (!rows) continue;
                if (options.sheetPerTable) sheets.push({ name: nameOf(chartName(node, heading), `Chart ${chartCount}`), rows, widths: widthsOf(rows) });
                else { if (together.length) together.push({ cells: [] }); for (const r of rows) together.push(r); }
                continue;
            }
            if (node.type === 'drawing') continue;
            // An embedded video or page: its address.
            if (node.type === 'embed') {
                const url = (node.metadata as { url?: string } | undefined)?.url ?? node.text;
                if (typeof url === 'string' && url.trim() && options.includeText && room(text)) text.push({ cells: [{ ...textCell(0, url.trim()), value: { kind: 'text' } }] });
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
                // Its first link the row's; a picture or chart in the text, which a cell cannot hold, reported or a sheet of its data.
                const links = blockText === null ? await linksIn(node, source.keep) : [];
                if (t && options.includeText && room(text)) text.push({ cells: [{ ...textCell(0, t, node.type === 'heading'), value: { kind: 'text' }, ...(links.length ? { links } : {}) }] });
                if (blockText === null) await mediaIn(node, source, media);
                await settle(heading);
                continue;
            }
            if (depth < 32 && node.children?.length) await visit(node.children, depth + 1);
        }
    };
    await visit(sheetNodes.length ? content.filter(n => n.type !== 'sheet') : content, 0);
    // (The tables one after another past a sheet's rows: the rest reported, a merge held to the last row.)
    if (together.length > MAX_ROWS) {
        together.length = MAX_ROWS;
        unrepresentable.add(PAST_ROWS);
        together.forEach((row, ri) => { for (const cell of row.cells) cell.rowSpan = Math.min(cell.rowSpan, MAX_ROWS - ri); });
    }
    if (!options.sheetPerTable && together.length) sheets.push({ name: nameOf('Tables', 'Tables'), rows: together, widths: widthsOf(together) });
    if (text.length) sheets.push({ name: nameOf('Text', 'Text'), rows: text, widths: [Math.min(120, Math.max(8, longestOf(text, 0)))] });
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

/** Whether two boxes are one (a shape's and the text in it). */
const sameBounds = (a: NodeBounds, b: NodeBounds) => Math.abs(a.x - b.x) < 0.5 && Math.abs(a.y - b.y) < 0.5 && Math.abs(a.width - b.width) < 0.5 && Math.abs(a.height - b.height) < 0.5;

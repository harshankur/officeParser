/**
 * Tables drawn with ruling lines, read from the lines themselves: a page's horizontal and vertical
 * rules (strokes, and the thin filled rectangles a browser draws borders as) are snapped and joined,
 * each set of rules that cross one another is a table's frame, the frame's distinct rule positions are
 * its grid, a cell is a grid cell or several where no rule divides them (so a merged title row or a
 * cell spanning rows keeps its span), and each text run goes to the cell its middle is in. A table
 * inside another table's cell is that cell's. The alignment-based reader (detectTables) took tables
 * from their text alone: one of two columns, one whose header wraps, one inside another, or one with
 * merged cells was read as prose or in pieces, and an empty cell was lost.
 *
 * Bounded: at most MAX_RULINGS rules a page, MAX_CROSSINGS crossing tests, MAX_GRID cells a frame.
 *
 * @module parsers/pdf/rulings
 */

import { CellMetadata, NodeBounds, OfficeContentNode } from '../../types.js';
import type { PdfShape } from './pdfShapes.js';
import type { PdfLayoutConfig, RawRun } from './pdfTypes.js';
import { CELL_PAGE_BOX } from './textLayout.js';

/** At most this many rules a page, this many crossing tests, and this many grid cells a frame. */
const MAX_RULINGS = 4_000;
const MAX_CROSSINGS = 2_000_000;
const MAX_GRID = 10_000;
/** Rules this close (in points) are one; rule positions this close are one grid line. */
const SNAP = 1.5;
const MERGE_GRID = 3;
/** A rule at least this long, and a filled rectangle at most this thick, is a rule. */
const MIN_RULE = 4;
const MAX_THICKNESS = 2.5;

/** A rule: horizontal (at `at`, from `from` to `to` across x) or vertical (across y). */
interface Rule { horizontal: boolean; at: number; from: number; to: number }

/** A frame read: its grid's lines, its cells (each a merged region), and its box. */
export interface RuledFrame {
    xs: number[];
    ys: number[];
    /** Each cell: its top-left grid position and spans. */
    cells: { row: number; col: number; rowSpan: number; colSpan: number }[];
    box: NodeBounds;
}

/** The rules a page's shapes draw. */
function rulesOf(shapes: readonly PdfShape[], pageW: number, pageH: number): Rule[] {
    const rules: Rule[] = [];
    const add = (horizontal: boolean, at: number, a: number, b: number) => {
        if (rules.length >= MAX_RULINGS) return;
        const from = Math.min(a, b), to = Math.max(a, b);
        if (to - from >= MIN_RULE) rules.push({ horizontal, at, from, to });
    };
    for (const s of shapes) {
        if (rules.length >= MAX_RULINGS) break;
        // The page's own background (a rectangle covering it) frames nothing.
        if (s.width >= 0.9 * pageW && s.height >= 0.9 * pageH) continue;
        if (s.curved) continue;
        if (s.filled && !s.stroked) {
            const white = !s.fill || /^#f[a-f0-9]f[a-f0-9]f[a-f0-9]$/i.test(s.fill);
            if (s.height <= MAX_THICKNESS && s.width >= MIN_RULE) add(true, s.y + s.height / 2, s.x, s.x + s.width);
            else if (s.width <= MAX_THICKNESS && s.height >= MIN_RULE) add(false, s.x + s.width / 2, s.y, s.y + s.height);
            else if (!white && s.segments.length <= 5) {
                // A shaded cell: its edges.
                add(true, s.y, s.x, s.x + s.width); add(true, s.y + s.height, s.x, s.x + s.width);
                add(false, s.x, s.y, s.y + s.height); add(false, s.x + s.width, s.y, s.y + s.height);
            }
            continue;
        }
        for (const [x1, y1, x2, y2] of s.segments) {
            if (Math.abs(y1 - y2) <= 0.5) add(true, (y1 + y2) / 2, x1, x2);
            else if (Math.abs(x1 - x2) <= 0.5) add(false, (x1 + x2) / 2, y1, y2);
        }
    }
    return rules;
}

/** Rules of one direction joined: those on one line (within SNAP) whose spans meet or overlap. */
function joinRules(rules: Rule[]): Rule[] {
    const sorted = [...rules].sort((a, b) => a.at - b.at || a.from - b.from);
    const out: Rule[] = [];
    let i = 0;
    while (i < sorted.length) {
        // A line: the rules within SNAP of the first's position.
        let j = i;
        while (j < sorted.length && sorted[j].at - sorted[i].at <= SNAP) j++;
        const line = sorted.slice(i, j).sort((a, b) => a.from - b.from);
        const at = line.reduce((s, r) => s + r.at, 0) / line.length;
        let cur: Rule | undefined;
        for (const r of line) {
            if (cur && r.from <= cur.to + SNAP) cur.to = Math.max(cur.to, r.to);
            else { cur = { horizontal: r.horizontal, at, from: r.from, to: r.to }; out.push(cur); }
        }
        i = j;
    }
    return out;
}

/** Positions within MERGE_GRID of each other as one (their mean): a cell's double border is one line. */
function gridLines(values: number[]): number[] {
    const sorted = [...values].sort((a, b) => a - b);
    const out: number[] = [];
    let group: number[] = [];
    for (const v of sorted) {
        if (group.length && v - group[0] > MERGE_GRID) { out.push(group.reduce((s, x) => s + x, 0) / group.length); group = []; }
        group.push(v);
    }
    if (group.length) out.push(group.reduce((s, x) => s + x, 0) / group.length);
    return out;
}

/** The index of the grid line nearest `v`. */
function nearest(lines: number[], v: number): number {
    let lo = 0, hi = lines.length - 1;
    while (hi - lo > 1) { const mid = (lo + hi) >> 1; if (lines[mid] <= v) lo = mid; else hi = mid; }
    return Math.abs(lines[lo] - v) <= Math.abs(lines[hi] - v) ? lo : hi;
}

/** The frames a page's shapes rule: sets of crossing rules with two of each direction at least, as grids. */
export function ruledFrames(shapes: readonly PdfShape[] | undefined, pageW: number, pageH: number): RuledFrame[] {
    if (!shapes?.length) return [];
    const all = rulesOf(shapes, pageW, pageH);
    const hs = joinRules(all.filter(r => r.horizontal)), vs = joinRules(all.filter(r => !r.horizontal));
    if (hs.length < 2 || vs.length < 2) return [];
    // Rules that cross (or meet, within SNAP) are one frame: a union-find over horizontal and vertical.
    const parent = [...hs.map((_, i) => i), ...vs.map((_, i) => hs.length + i)];
    const find = (i: number): number => { while (parent[i] !== i) { parent[i] = parent[parent[i]]; i = parent[i]; } return i; };
    const vByX = vs.map((v, i) => ({ v, i })).sort((a, b) => a.v.at - b.v.at);
    const xsSorted = vByX.map(e => e.v.at);
    let tests = 0;
    for (let h = 0; h < hs.length && tests < MAX_CROSSINGS; h++) {
        const r = hs[h];
        // The vertical rules within the horizontal one's span.
        let lo = 0, hi = xsSorted.length;
        while (lo < hi) { const mid = (lo + hi) >> 1; if (xsSorted[mid] < r.from - SNAP) lo = mid + 1; else hi = mid; }
        for (let k = lo; k < vByX.length && vByX[k].v.at <= r.to + SNAP && tests < MAX_CROSSINGS; k++, tests++) {
            const v = vByX[k].v;
            if (r.at >= v.from - SNAP && r.at <= v.to + SNAP) {
                const a = find(h), b = find(hs.length + vByX[k].i);
                if (a !== b) parent[a] = b;
            }
        }
    }
    const groups = new Map<number, { h: Rule[]; v: Rule[] }>();
    hs.forEach((r, i) => { const g = find(i); const e = groups.get(g) ?? { h: [], v: [] }; e.h.push(r); groups.set(g, e); });
    vs.forEach((r, i) => { const g = find(hs.length + i); const e = groups.get(g) ?? { h: [], v: [] }; e.v.push(r); groups.set(g, e); });
    const frames: RuledFrame[] = [];
    for (const { h, v } of groups.values()) {
        if (h.length < 2 || v.length < 2) continue;
        const xs = gridLines(v.map(r => r.at)), ys = gridLines(h.map(r => r.at));
        const rows = ys.length - 1, cols = xs.length - 1;
        if (rows < 1 || cols < 1 || rows * cols > MAX_GRID || rows * cols < 2) continue;
        // Which edges between grid cells a rule draws: a vertical edge at xs[c] over row r, and so on.
        const vEdge = Array.from({ length: xs.length }, () => new Array<boolean>(rows).fill(false));
        const hEdge = Array.from({ length: ys.length }, () => new Array<boolean>(cols).fill(false));
        for (const r of v) {
            const c = nearest(xs, r.at);
            for (let row = 0; row < rows; row++) {
                const top = ys[row], bottom = ys[row + 1];
                if (r.from <= top + MERGE_GRID && r.to >= bottom - MERGE_GRID) vEdge[c][row] = true;
            }
        }
        for (const r of h) {
            const rr = nearest(ys, r.at);
            for (let col = 0; col < cols; col++) {
                if (r.from <= xs[col] + MERGE_GRID && r.to >= xs[col + 1] - MERGE_GRID) hEdge[rr][col] = true;
            }
        }
        // Cells: grid cells joined where no rule divides them, each region a rectangle (else no table).
        const owner = Array.from({ length: rows }, () => new Array<number>(cols).fill(-1));
        const cells: RuledFrame['cells'] = [];
        let ok = true;
        for (let row = 0; row < rows && ok; row++) {
            for (let col = 0; col < cols && ok; col++) {
                if (owner[row][col] >= 0) continue;
                let colSpan = 1, rowSpan = 1;
                while (col + colSpan < cols && !vEdge[col + colSpan][row]) colSpan++;
                while (row + rowSpan < rows && !hEdge[row + rowSpan][col]) rowSpan++;
                for (let r = row; r < row + rowSpan; r++) for (let c = col; c < col + colSpan; c++) {
                    if (owner[r][c] >= 0) ok = false;
                    owner[r][c] = cells.length;
                }
                cells.push({ row, col, rowSpan, colSpan });
            }
        }
        if (!ok || cells.length < 2) continue;
        frames.push({ xs, ys, cells, box: { x: xs[0], y: ys[0], width: xs[xs.length - 1] - xs[0], height: ys[ys.length - 1] - ys[0] } });
    }
    return frames;
}

/** What turns a cell's runs into its content, and a box into the AST's. */
export interface CellLayout {
    cfg: PdfLayoutConfig;
    content: (runs: RawRun[]) => OfficeContentNode[];
    bounds: (box: NodeBounds) => NodeBounds | undefined;
}

/** How far a comb's boxes' widths may be from their middle one's, as a share of it (see isComb). */
const COMB_WIDTH_SPREAD = 0.15;
/**
 * The fewest boxes a row of choices fills (see isScale): two characters in two boxes ("1 2", "A B") may as well
 * be a value's.
 */
const SCALE_BOXES = 3;
/**
 * How a row of boxes the page draws, a character to a box, is a comb (see isComb): its boxes about square,
 * taller than wide or a little wider (no wider than COMB_SHAPE times their height; a row of digits alone up to
 * COMB_DIGITS_SHAPE, a date's in boxes of 20 by 15 or 16 by 12) and tight to their characters (no wider than
 * COMB_EMS of them): a date's digits of 14 points in square boxes of 20 (1.4 ems), of 8 in boxes of 18 (2.25
 * ems). A row of choices is ruled apart: weekdays' letters in boxes of 26 by 18 or 30 by 18 (1.44 and 1.67
 * times as wide as tall), a scale's numbers of 11 points in boxes of 30 by 24 (2.7 ems).
 */
const COMB_SHAPE = 1.3;
const COMB_DIGITS_SHAPE = 1.5;
const COMB_EMS = 2.5;
/** The characters a scale's choices count through, one after another, and a school's grades (no E). */
const SCALE_RUNS = ['0123456789', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'];
const GRADES = ['ABCDF', 'abcdf'];

/**
 * Whether a row of boxes, one character to a box, prints a scale's choices rather than holds a value: every
 * box filled, with characters that count up or down one at a time ("1 2 3 4 5", "0 1 2 ... 9", "A B C D E",
 * "5 4 3 2 1"), digits counting by another step ("1 3 5 7 9") or a school's grades ("A B C D F"), three at
 * least. A value's characters (a date's "15041985", a name's) do not count so; one that does, a comb filled
 * with "12345" however tight its boxes, is read as the row of choices it looks like, its characters still in
 * order, each in its cell.
 */
function isScale(chars: readonly string[]): boolean {
    if (chars.length < SCALE_BOXES || chars.some(c => [...c].length !== 1)) return false;
    const text = chars.join(''), back = [...chars].reverse().join('');
    if (SCALE_RUNS.some(run => run.includes(text) || run.includes(back)) || GRADES.includes(text) || GRADES.includes(back)) return true;
    // (Digits counting by a step of their own, the odd or even choices: "1 3 5 7 9", "2 4 6", "10 ... 0" by twos.)
    if (!chars.every(c => /^[0-9]$/.test(c))) return false;
    const step = Number(chars[1]) - Number(chars[0]);
    return step !== 0 && chars.every((c, i) => i === 0 || Number(c) - Number(chars[i - 1]) === step);
}

/**
 * Whether a frame is a comb: a row of like boxes (two at least, of widths within COMB_WIDTH_SPREAD of
 * their middle one's, about square: no wider than two and a half times their height, no taller than four
 * times their width) a form draws to be filled one character to a box, a date's digits or a tax number's.
 * It is one where its text is form field values (see RawRun.formField) alone, or where each box holds a
 * character at most, in boxes of a comb's shape and tight to them (see COMB_SHAPE, COMB_EMS), and its
 * characters are not a scale's choices (see isScale). A row of words, a row of letters, a scale's numbers or
 * grades, or a grid of two rows, is a table: where it is unsure, a row stays a table, each character in its
 * cell.
 */
function isComb(f: RuledFrame, contents: readonly RawRun[][]): boolean {
    if (f.ys.length !== 2 || f.cells.length < 2 || f.cells.some(c => c.colSpan !== 1 || c.rowSpan !== 1)) return false;
    const height = f.ys[1] - f.ys[0];
    const widths = f.cells.map(c => f.xs[c.col + 1] - f.xs[c.col]);
    const middle = [...widths].sort((a, b) => a - b)[widths.length >> 1];
    if (!(middle > 0 && height > 0) || middle > 2.5 * height || height > 4 * middle) return false;
    if (widths.some(w => Math.abs(w - middle) > COMB_WIDTH_SPREAD * middle)) return false;
    const texts = contents.flat().filter(r => r.text.trim());
    if (texts.length && texts.every(r => r.formField)) return true;
    const size = texts.reduce((m, r) => Math.max(m, r.fontSize), 0);
    if (size > 0 && middle > COMB_EMS * size) return false;
    const chars: string[] = [];
    for (const runs of contents) {
        let box = '', count = 0;
        for (const r of runs) for (const ch of r.text) if (ch.trim()) { if (++count > 1) return false; box = ch; }
        chars.push(box);
    }
    // (Wider boxes hold a comb's characters only where they are all digits: "Y N" or "M T W T F S S" in boxes
    // a half again as wide as tall are choices.)
    const digits = chars.every(c => !c || /^\p{Nd}$/u.test(c));
    if (middle > (digits ? COMB_DIGITS_SHAPE : COMB_SHAPE) * height) return false;
    // (In the boxes' order across the row.)
    const order = f.cells.map((c, i) => i).sort((a, b) => f.cells[a].col - f.cells[b].col);
    return !isScale(order.map(i => chars[i]));
}

/**
 * A comb's characters (see isComb) as one run across its boxes: a box left empty between two filled ones
 * is a space, those before the first and after the last are none. A form field's value, read where its
 * boxes are, as its field's are (RawRun.formField).
 */
function combRun(f: RuledFrame, contents: readonly RawRun[][]): RawRun | undefined {
    const order = f.cells.map((c, i) => i).sort((a, b) => f.cells[a].col - f.cells[b].col);
    const chars = order.map(i => contents[i].filter(r => r.text.trim()).sort((a, b) => a.x - b.x).map(r => r.text.trim()).join(''));
    const first = chars.findIndex(Boolean);
    if (first < 0) return undefined;
    let last = chars.length - 1;
    while (!chars[last]) last--;
    const runs = order.flatMap(i => contents[i]).filter(r => r.text.trim());
    const like = runs[0];
    const fontSize = runs.reduce((m, r) => Math.max(m, r.fontSize), 0) || like.fontSize;
    return {
        ...like, text: chars.slice(first, last + 1).map(c => c || ' ').join(''),
        x: f.xs[0], width: f.xs[f.xs.length - 1] - f.xs[0], fontSize, link: undefined, formField: 'text',
    };
}

/**
 * The tables a page's frames hold, with the runs they took: each run goes to the cell its middle is in,
 * a frame inside another's cell first (it is that cell's, among its paragraphs by its place). A frame
 * whose cells hold no text is no table. Nor is a comb (see isComb): its characters are one run (see
 * combRun), its cell's where it is inside a table's, else one of `combs`, read with the page's text; or
 * the form field values over it are read as they are. A
 * field's value is not cut at a table's rules (it is one value), and goes to the cell its middle is in.
 */
export function ruledTables(frames: RuledFrame[], runs: readonly RawRun[], layout: CellLayout): { tables: OfficeContentNode[]; consumed: Set<RawRun>; combs: RawRun[] } {
    const consumed = new Set<RawRun>();
    const combs: RawRun[] = [];
    const tables: OfficeContentNode[] = [];
    const inside = (a: NodeBounds, b: NodeBounds) => a !== b && a.x >= b.x - SNAP && a.y >= b.y - SNAP && a.x + a.width <= b.x + b.width + SNAP && a.y + a.height <= b.y + b.height + SNAP;
    // Smaller frames first, so a frame inside another's cell is read before it, as that cell's.
    const order = frames.map((f, i) => i).sort((a, b) => frames[a].box.width * frames[a].box.height - frames[b].box.width * frames[b].box.height);
    const built = new Map<number, OfficeContentNode>();
    const nestedIn = new Set<number>();
    for (const fi of order) {
        const f = frames[fi];
        const contents: RawRun[][] = f.cells.map(() => []);
        const cellIndex = new Map<string, number>();
        f.cells.forEach((c, i) => { for (let r = c.row; r < c.row + c.rowSpan; r++) for (let k = c.col; k < c.col + c.colSpan; k++) cellIndex.set(`${r},${k}`, i); });
        // (A piece of a run cut at a column's rule, and the run it was cut from.)
        const pieceOf = new Map<RawRun, RawRun>();
        const place = (run: RawRun) => {
            const cx = run.x + run.width / 2, cy = run.yTop + run.height / 2;
            if (!(cx >= f.xs[0] && cx <= f.xs[f.xs.length - 1] && cy >= f.ys[0] && cy <= f.ys[f.ys.length - 1])) return;
            let col = 0, row = 0;
            while (col < f.xs.length - 2 && cx > f.xs[col + 1]) col++;
            while (row < f.ys.length - 2 && cy > f.ys[row + 1]) row++;
            const at = cellIndex.get(`${row},${col}`);
            if (at !== undefined) contents[at].push(run);
        };
        // The rules between distinct cells of each row (a merged cell's own grid lines are none), by row.
        const bordersOf = new Map<number, number[]>();
        const rowBorders = (row: number) => {
            let out = bordersOf.get(row);
            if (!out) {
                out = [f.xs[0]];
                for (let k = 1; k < f.xs.length - 1; k++) if (cellIndex.get(`${row},${k}`) !== cellIndex.get(`${row},${k - 1}`)) out.push(f.xs[k]);
                out.push(f.xs[f.xs.length - 1]);
                bordersOf.set(row, out);
            }
            return out;
        };
        // (A comb read before, inside this frame's cell, is that cell's: see combs below.)
        for (const run of combs.length ? runs.concat(combs) : runs) {
            if (consumed.has(run) || run.angle !== 0) continue;
            // Only a run within the frame's rows is cut, and only at the borders of its own row's cells.
            const cy = run.yTop + run.height / 2;
            if (!(cy >= f.ys[0] && cy <= f.ys[f.ys.length - 1]) || run.x + run.width < f.xs[0] || run.x > f.xs[f.xs.length - 1]) continue;
            let row = 0;
            while (row < f.ys.length - 2 && cy > f.ys[row + 1]) row++;
            const pieces = run.formField ? [run] : cutAtRules(run, rowBorders(row));
            if (pieces.length > 1) for (const p of pieces) { pieceOf.set(p, run); place(p); }
            else place(run);
        }
        // The frames inside this one, each in the cell holding its middle.
        const innerOf = new Map<number, OfficeContentNode[]>();
        for (const [gi, node] of built) {
            const g = frames[gi];
            if (nestedIn.has(gi) || !inside(g.box, f.box)) continue;
            const cx = g.box.x + g.box.width / 2, cy = g.box.y + g.box.height / 2;
            const at = f.cells.findIndex(c => cx >= f.xs[c.col] && cx <= f.xs[c.col + c.colSpan] && cy >= f.ys[c.row] && cy <= f.ys[c.row + c.rowSpan]);
            if (at < 0) continue;
            nestedIn.add(gi);
            (innerOf.get(at) ?? innerOf.set(at, []).get(at)!).push(node);
        }
        if (!contents.some(c => c.some(r => r.text.trim())) && !innerOf.size) continue;
        if (!innerOf.size && isComb(f, contents)) {
            // (Values over it are read as they are, the frame's characters as one.)
            const comb = contents.some(c => c.some(r => r.text.trim() && !r.formField)) ? combRun(f, contents) : undefined;
            if (comb) {
                for (const c of contents) for (const r of c) consumed.add(pieceOf.get(r) ?? r);
                combs.push(comb);
            }
            continue;
        }
        // A cell joined across a grid line its text does not cross, its text in one column (or row) and
        // another, is cells a producer drew no rule between (partial borders), not a merged one: the
        // frame is left to the alignment reader, which reads such tables from their text.
        const falseMerge = f.cells.some((c, i) => {
            const spans = (along: 'x' | 'y', lines: number[], from: number, span: number) => {
                if (span < 2) return false;
                const bands = new Set<number>();
                for (const run of contents[i]) {
                    if (!run.text.trim()) continue;
                    const a = along === 'x' ? run.x : run.yTop, b = a + (along === 'x' ? run.width : run.height);
                    const band = lines.slice(from, from + span).findIndex((line, k) => a >= line - SNAP && b <= lines[from + k + 1] + SNAP);
                    if (band < 0) return false;
                    bands.add(band);
                }
                return bands.size > 1;
            };
            return spans('x', f.xs, c.col, c.colSpan) || spans('y', f.ys, c.row, c.rowSpan);
        });
        // A cell whose own lines hold text set apart in columns (a gap of two ems and more between runs on one
        // baseline) is several cells a producer drew no rule between: a box with lines between its rows alone
        // read each row as one cell. So is a frame of one column (row lines in a box, of a list or a table).
        const columned = f.xs.length - 1 < 2 || contents.some(runs => {
            const lines = new Map<number, RawRun[]>();
            for (const run of runs) if (run.text.trim()) { const key = Math.round(run.yBaseline); (lines.get(key) ?? lines.set(key, []).get(key)!).push(run); }
            for (const line of lines.values()) {
                line.sort((a, b) => a.x - b.x);
                for (let k = 1; k < line.length; k++) if (line[k].x - (line[k - 1].x + line[k - 1].width) >= 2 * Math.max(line[k].fontSize, line[k - 1].fontSize, 4)) return true;
            }
            return false;
        });
        if (falseMerge || columned) continue;
        const rows: OfficeContentNode[] = [];
        f.cells.forEach((c, i) => {
            for (const r of contents[i]) consumed.add(pieceOf.get(r) ?? r);
            const children = layout.content(contents[i]);
            // A table inside the cell, among its paragraphs by its place.
            for (const inner of innerOf.get(i) ?? []) {
                const y = inner.bounds?.y ?? 0;
                const atY = children.findIndex(n => (n.bounds?.y ?? Infinity) > y);
                if (atY < 0) children.push(inner); else children.splice(atY, 0, inner);
            }
            const meta: CellMetadata = { row: c.row, col: c.col } as CellMetadata;
            if (c.colSpan > 1) meta.colSpan = c.colSpan;
            if (c.rowSpan > 1) meta.rowSpan = c.rowSpan;
            const cell: OfficeContentNode = { type: 'cell', text: children.map(n => n.text || '').join('\n').trim(), children, metadata: meta };
            const pageBox = { x: f.xs[c.col], y: f.ys[c.row], width: f.xs[c.col + c.colSpan] - f.xs[c.col], height: f.ys[c.row + c.rowSpan] - f.ys[c.row] };
            CELL_PAGE_BOX.set(cell, pageBox);
            const box = layout.bounds(pageBox);
            if (box) cell.bounds = box;
            while (rows.length <= c.row) rows.push({ type: 'row', children: [] });
            rows[c.row].children!.push(cell);
        });
        for (const row of rows) row.text = (row.children || []).map(c => c.text || '').join(' ').trim();
        const table: OfficeContentNode = { type: 'table', children: rows, text: rows.map(r => r.text || '').join('\n') };
        const box = layout.bounds(f.box);
        if (box) table.bounds = box;
        built.set(fi, table);
    }
    for (const [fi, node] of built) if (!nestedIn.has(fi)) tables.push(node);
    // A comb inside a table's cell is that cell's; the others are the page's text.
    return { tables, consumed, combs: combs.filter(c => !consumed.has(c)) };
}

/** A filled rectangle where it paints (see backgroundFills), and its colour. */
interface Fill { x: number; y: number; width: number; height: number; fill: string }

/**
 * The rectangles a page fills with a colour that something may stand on, in the order the page draws them
 * (one drawn later covers those before it), each where it paints: its box within the clip it was painted
 * through (see PdfShape.clip), as a browser fills a whole box clipped to its left border's strip. Filled
 * with a colour and not see-through (a shadow under a box is no background), straight, of at most five
 * edges, thicker than a rule, and not the page's own background (nearly all of it in both directions). White ones too: a white card on a grey panel is what the text on
 * it stands on (see colourOf).
 */
function backgroundFills(shapes: readonly PdfShape[] | undefined, pageW: number, pageH: number): Fill[] {
    const out: Fill[] = [];
    for (const s of shapes ?? []) {
        if (!s.filled || !s.fill || s.seeThrough || s.curved || s.segments.length > 5) continue;
        let { x, y, width, height } = s;
        if (s.clip) {
            const x1 = Math.min(x + width, s.clip.x + s.clip.width), y1 = Math.min(y + height, s.clip.y + s.clip.height);
            x = Math.max(x, s.clip.x); y = Math.max(y, s.clip.y);
            width = x1 - x; height = y1 - y;
        }
        if (width > MAX_THICKNESS && height > MAX_THICKNESS && !(width >= 0.9 * pageW && height >= 0.9 * pageH)) out.push({ x, y, width, height, fill: s.fill });
    }
    return out;
}

/** The colour a fill gives what stands on it: none for white, the paper's own colour. */
function colourOf(fill: Fill): string | undefined {
    return fill.fill.toLowerCase() !== '#ffffff' ? fill.fill : undefined;
}

/**
 * Tables' cells shaded as their page shades them: each cell (read from the page, with no colour yet)
 * takes the colour of the rectangle the page draws last under its middle (none where it is white), so a
 * header row's white text is read on the colour it was set on, not on none.
 */
export function shadeCells(tables: readonly OfficeContentNode[], shapes: readonly PdfShape[] | undefined, pageW: number, pageH: number): void {
    if (!shapes?.length || !tables.length) return;
    const fills = backgroundFills(shapes, pageW, pageH);
    if (!fills.length) return;
    let tests = 0;
    const visit = (node: OfficeContentNode) => {
        for (const child of node.children ?? []) {
            const box = child.type === 'cell' ? CELL_PAGE_BOX.get(child) : undefined;
            if (box && !(child.metadata as CellMetadata | undefined)?.backgroundColor && tests < MAX_CROSSINGS) {
                const cx = box.x + box.width / 2, cy = box.y + box.height / 2;
                // (From the last drawn: the first under it is the one on top.)
                let under: Fill | undefined;
                for (let i = fills.length - 1; i >= 0 && tests < MAX_CROSSINGS; i--) {
                    tests++;
                    const f = fills[i];
                    if (cx >= f.x && cx <= f.x + f.width && cy >= f.y && cy <= f.y + f.height) { under = f; break; }
                }
                const colour = under && colourOf(under);
                if (colour) child.metadata = { ...(child.metadata as CellMetadata), backgroundColor: colour };
            }
            // The cell shows its colour: its text is not also on that colour (see shadeRuns).
            const shown = child.type === 'cell' ? (child.metadata as CellMetadata | undefined)?.backgroundColor : undefined;
            if (shown) clearRunShading(child, shown);
            visit(child);
        }
    };
    for (const t of tables) visit(t);
}

/** The bands a page's height is cut into, to find the rectangles holding a place without testing them all. */
const BANDS = 256;
/** A rectangle taller than this many bands is kept in a list of its own, tested for every place. */
const MAX_BANDS_A_RECTANGLE = 16;
/** How far (in points) a run's box may reach past the rectangle it stands on and still be on it. */
const REACH = 1.5;

/**
 * Text that stands on a colour keeps it: each run whose whole box is inside a filled rectangle of the page
 * takes the colour of the one the page draws last (see backgroundFills) as the run's `backgroundColor`, as a
 * word processor's paragraph shading reaches its runs: none where that one is white, as text on a white card
 * inside a grey panel is on white. White text on a purple bar is then read on the bar, not as white text on
 * none. A run only partly over a rectangle, one with a background already (a highlight) and invisible text
 * are left as they are; a table's cells, shaded as the page shades them (see shadeCells), take their runs'
 * background of the same colour back, the cell showing it.
 * Bounded: the rectangles are looked up by the band of the page they touch, and the tests are no more than
 * MAX_CROSSINGS in all, past which the page's other runs are left unshaded.
 */
export function shadeRuns(runs: readonly RawRun[], shapes: readonly PdfShape[] | undefined, pageW: number, pageH: number): void {
    if (!shapes?.length || !runs.length) return;
    const fills = backgroundFills(shapes, pageW, pageH);
    if (!fills.some(colourOf)) return;
    // The fills (in the order drawn) by the bands they touch: a run is looked for in the band its middle is in.
    const size = Math.max(pageH, 1) / BANDS;
    const band = (y: number) => Math.min(BANDS - 1, Math.max(0, Math.floor(y / size)));
    const bands: number[][] = [];
    const tall: number[] = [];
    fills.forEach((f, i) => {
        const lo = band(f.y), hi = band(f.y + f.height);
        if (hi - lo + 1 > MAX_BANDS_A_RECTANGLE) { tall.push(i); return; }
        for (let b = lo; b <= hi; b++) (bands[b] ??= []).push(i);
    });
    let tests = 0;
    for (const run of runs) {
        if (run.formatting.backgroundColor || run.formatting.invisible || !run.text.trim()) continue;
        const inBand = bands[band(run.yTop + run.height / 2)] ?? [];
        // The two lists are each in the order drawn: the rectangle on top holding the run is the last of either.
        let a = inBand.length - 1, b = tall.length - 1;
        while ((a >= 0 || b >= 0) && tests < MAX_CROSSINGS) {
            tests++;
            const f = fills[b < 0 || (a >= 0 && inBand[a] > tall[b]) ? inBand[a--] : tall[b--]];
            if (run.x >= f.x - REACH && run.x + run.width <= f.x + f.width + REACH && run.yTop >= f.y - REACH && run.yTop + run.height <= f.y + f.height + REACH) {
                const colour = colourOf(f);
                if (colour) run.formatting = { ...run.formatting, backgroundColor: colour };
                break;
            }
        }
        if (tests >= MAX_CROSSINGS) return;
    }
}

/** Takes the background `color` off the text under `node` (a shaded cell's own, which the cell shows). */
function clearRunShading(node: OfficeContentNode, color: string): void {
    for (const child of node.children ?? []) {
        if (child.formatting?.backgroundColor?.toLowerCase() === color.toLowerCase()) {
            const { backgroundColor: _shown, ...rest } = child.formatting;
            child.formatting = rest;
        }
        clearRunShading(child, color);
    }
}

/**
 * A run drawn across a table's column rules (a producer setting a row's cells in one string, "Apples 10")
 * cut at the word boundary nearest each rule it crosses, where text stands at least half an em on both
 * sides; each piece over its share of the run's width (in proportion to its letters). A run that crosses
 * no rule, or has no word boundary to cut at, is itself.
 */
export function cutAtRules(run: RawRun, xs: readonly number[]): RawRun[] {
    const em = run.fontSize || 12;
    const inner = xs.slice(1, -1).filter(x => x > run.x + em / 2 && x < run.x + run.width - em / 2);
    if (!inner.length || inner.length > 64 || run.text.length > 4096) return [run];
    // The word boundaries (where a run of spaces starts after a letter), found once.
    const bounds: number[] = [];
    const lastLetter = run.text.trimEnd().length;
    for (let i = 1; i < lastLetter; i++) if (/\s/.test(run.text[i]) && !/\s/.test(run.text[i - 1])) bounds.push(i);
    if (!bounds.length) return [run];
    const perChar = run.width / Math.max(1, run.text.length);
    const cuts: number[] = [];
    for (const x of inner) {
        // The boundary nearest the rule's place among the letters (by halving).
        const want = (x - run.x) / perChar;
        let lo = 0, hi = bounds.length - 1;
        while (lo < hi) { const mid = (lo + hi) >> 1; if (bounds[mid] < want) lo = mid + 1; else hi = mid; }
        const best = lo > 0 && Math.abs(bounds[lo - 1] - want) <= Math.abs(bounds[lo] - want) ? bounds[lo - 1] : bounds[lo];
        if (!cuts.includes(best)) cuts.push(best);
    }
    if (!cuts.length) return [run];
    cuts.sort((a, b) => a - b);
    const pieces: RawRun[] = [];
    let from = 0;
    const edges = [...inner].sort((a, b) => a - b);
    for (const [k, cut] of [...cuts, run.text.length].entries()) {
        const text = run.text.slice(from, cut).trim();
        // Each piece from its rule (the run's start for the first) to the next (the run's end for the last).
        const left = k === 0 ? run.x : Math.max(run.x, edges[Math.min(k - 1, edges.length - 1)]);
        const right = k === cuts.length ? run.x + run.width : Math.min(run.x + run.width, edges[Math.min(k, edges.length - 1)]);
        if (text) pieces.push({ ...run, text, x: left, width: Math.max(0, right - left) });
        from = cut;
    }
    return pieces.length > 1 ? pieces : [run];
}

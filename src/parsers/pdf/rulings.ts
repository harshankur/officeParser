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

/**
 * The tables a page's frames hold, with the runs they took: each run goes to the cell its middle is in,
 * a frame inside another's cell first (it is that cell's, among its paragraphs by its place). A frame
 * whose cells hold no text is no table.
 */
export function ruledTables(frames: RuledFrame[], runs: readonly RawRun[], layout: CellLayout): { tables: OfficeContentNode[]; consumed: Set<RawRun> } {
    const consumed = new Set<RawRun>();
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
        for (const run of runs) {
            if (consumed.has(run) || run.angle !== 0) continue;
            const cx = run.x + run.width / 2, cy = run.yTop + run.height / 2;
            if (cx < f.xs[0] || cx > f.xs[f.xs.length - 1] || cy < f.ys[0] || cy > f.ys[f.ys.length - 1]) continue;
            let col = 0, row = 0;
            while (col < f.xs.length - 2 && cx > f.xs[col + 1]) col++;
            while (row < f.ys.length - 2 && cy > f.ys[row + 1]) row++;
            const at = cellIndex.get(`${row},${col}`);
            if (at !== undefined) contents[at].push(run);
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
        if (falseMerge) continue;
        const rows: OfficeContentNode[] = [];
        f.cells.forEach((c, i) => {
            for (const r of contents[i]) consumed.add(r);
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
            const box = layout.bounds({ x: f.xs[c.col], y: f.ys[c.row], width: f.xs[c.col + c.colSpan] - f.xs[c.col], height: f.ys[c.row + c.rowSpan] - f.ys[c.row] });
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
    return { tables, consumed };
}

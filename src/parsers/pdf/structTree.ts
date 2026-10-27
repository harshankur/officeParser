/**
 * The tagged-PDF path: turns a page's structure tree into semantic AST nodes.
 *
 * pdf.js hands back a per-page tree of `{ role, children }` internal nodes and
 * `{ type: 'content', id }` leaves, where `id` is a marked-content id. Every text run collected in
 * the main pass already carries its `mcid`, so joining a leaf to its runs is a map lookup. Block
 * text (P, Hn, list items, cells) is rebuilt with the same line/spacing/bounds logic as the
 * geometric path via {@link runsToParagraph}, so both paths produce identical run shapes.
 *
 * @module parsers/pdf/structTree
 */

import { CellMetadata, ListMetadata, NodeBounds, NoteMetadata, OfficeContentNode } from '../../types.js';
import { unionAll } from './geometry.js';
import type { RuledFrame } from './rulings.js';
import { RawRun } from './pdfTypes.js';
import { median } from '../../utils/numberUtils.js';
import { CELL_PAGE_BOX, DocContext, PageContext, rightToLeftLean, runsToParagraph } from './textLayout.js';
import { appendAll } from '../../utils/nodeListUtils.js';

interface StructNode {
    role?: string;
    type?: string;
    id?: string;
    alt?: string;
    children?: StructNode[];
}

/** Options that gate which tagged roles are emitted. */
export interface TaggedOptions {
    ignoreNotes: boolean;
    /**
     * Document-wide list counter, owned by the caller and shared across pages. Numbering is keyed on
     * `listId`, so a per-page counter would hand every page's first list the same `pdf-list-1` and
     * make a generator continue page 2's list from page 1's. The geometric path counts the same way.
     */
    listCounter: { n: number };
    /** The page's ruled frames (see rulings), read when a table needs them: they place a tagged table's cells. */
    frames?: () => readonly RuledFrame[];
}

/** Result of walking one page's structure tree. */
export interface TaggedResult {
    nodes: OfficeContentNode[];
    /** Marked-content ids that were consumed, for the coverage/trust check. */
    coveredMcids: Set<string>;
}

interface WalkCtx {
    runsByMcid: Map<string, RawRun[]>;
    page: PageContext;
    doc: DocContext;
    covered: Set<string>;
    opts: TaggedOptions;
    listCounter: { n: number };
}

const HEADING = /^H([1-6])$/;
const role = (n: StructNode): string => n.role || n.type || '';

/**
 * Resolves the heading level to force for a tagged heading node, honoring `headingDetection`:
 * - `'off'`:       0 - never a heading, always a paragraph.
 * - `'font-size'`: `undefined` - ignore the tag's own level and let the geometric font-size/weight
 *                  heuristic in {@link runsToParagraph} decide (which may also demote it to a
 *                  paragraph), i.e. "always use the heuristic, even on tagged PDFs".
 * - `'auto'`, `'tags'`: the tag's declared level (`tagLevel`), trusting the document's own structure.
 */
function taggedHeadingLevel(cfg: DocContext['cfg'], tagLevel: number): number | undefined {
    if (cfg.headingDetection === 'off') return 0;
    if (cfg.headingDetection === 'font-size') return undefined;
    return tagLevel;
}

/** Walks a page's struct tree into ordered AST nodes. */
/**
 * Maximum struct-tree depth the recursive walkers below will descend. The tree comes straight from an
 * untrusted PDF's `/StructTreeRoot`, so a pathologically deep (or cyclic) `/K` chain would overflow the
 * stack in `walkNode`/`collectRuns`/`buildList`/etc. A real document is only a handful of levels deep;
 * anything past this cap is treated as an unusable tag tree and the page falls back to geometry.
 */
const MAX_STRUCT_DEPTH = 256;

/**
 * Iterative depth probe: returns true if any path in the tree is deeper than {@link MAX_STRUCT_DEPTH}
 * (a cycle qualifies, since depth grows without bound along it). Recursing here would itself overflow,
 * so it walks with an explicit stack. A well-formed tree is visited fully (O(nodes)); a malformed one
 * short-circuits as soon as any path passes the cap, so a cycle costs O(cap), not an infinite walk.
 */
function structTooDeep(root: StructNode): boolean {
    const stack: Array<{ n: StructNode; d: number }> = [{ n: root, d: 0 }];
    while (stack.length) {
        const { n, d } = stack.pop()!;
        if (d > MAX_STRUCT_DEPTH) return true;
        if (n.children) for (const c of n.children) stack.push({ n: c, d: d + 1 });
    }
    return false;
}

export function buildTaggedNodes(
    structTree: any, runsByMcid: Map<string, RawRun[]>, page: PageContext, doc: DocContext, opts: TaggedOptions,
): TaggedResult {
    const covered = new Set<string>();
    const ctx: WalkCtx = { runsByMcid, page, doc, covered, opts, listCounter: opts.listCounter };
    // Guard the recursive walk against a hostile deep/cyclic tree: too deep -> no tagged nodes, which
    // drives the page onto the geometry fallback rather than overflowing the stack.
    const nodes = (structTree && !structTooDeep(structTree as StructNode)) ? walkChildren(structTree as StructNode, ctx, 0) : [];
    return { nodes, coveredMcids: covered };
}

function walkChildren(node: StructNode, ctx: WalkCtx, sectionDepth: number): OfficeContentNode[] {
    const out: OfficeContentNode[] = [];
    for (const child of node.children || []) appendAll(out, walkNode(child, ctx, sectionDepth));
    return out;
}

function walkNode(node: StructNode, ctx: WalkCtx, sectionDepth: number): OfficeContentNode[] {
    if (node.type === 'content') {
        const p = paragraphFrom(node, ctx, 0);
        return p ? [p] : [];
    }
    const r = role(node);
    const hm = HEADING.exec(r);
    if (hm) {
        const p = paragraphFrom(node, ctx, taggedHeadingLevel(ctx.doc.cfg, parseInt(hm[1], 10)));
        return p ? [p] : [];
    }

    switch (r) {
        case 'H': return blockWithNotes(node, ctx, taggedHeadingLevel(ctx.doc.cfg, Math.min(6, Math.max(1, sectionDepth))));
        case 'P': case 'Caption': case 'Title': case 'Lbl': case 'LBody': return blockWithNotes(node, ctx, 0);
        // A table-of-contents entry (TOCI) wraps a Link + Span + leader dots + page number; emit it as
        // one paragraph instead of letting the default recursion shatter each content leaf into its own.
        case 'TOCI': {
            const nodes = blockWithNotes(node, ctx, 0);
            for (const n of nodes) stripDotLeaders(n);
            return nodes;
        }
        case 'Table': { const t = buildTable(node, ctx); return t ? [t] : []; }
        case 'L': return buildList(node, ctx, 0);
        case 'Note': case 'FENote': {
            if (ctx.opts.ignoreNotes) { markCovered(node, ctx); return []; }
            const n = buildNote(node, ctx);
            return n ? [n] : [];
        }
        case 'Sect': case 'Part': case 'Art': case 'Section':
            return walkChildren(node, ctx, sectionDepth + 1);
        case 'Formula': {
            // Real math reconstruction from positioned glyphs is out of scope; at least mark the block
            // as a formula so consumers can tell it apart from ordinary prose.
            const nodes = blockWithNotes(node, ctx, 0);
            for (const n of nodes) if (n.type === 'paragraph') n.metadata = { ...(n.metadata || {}), style: 'formula' };
            return nodes;
        }
        case 'Private': return [];
        // Transparent containers: recurse and flatten.
        default:
            return walkChildren(node, ctx, sectionDepth);
    }
}

/** Role subtrees whose text must never be folded into a parent paragraph. */
const ALWAYS_SKIP = new Set(['Note', 'FENote']);

/** Gathers runs under a node from content leaves, skipping notes and any extra role subtrees. */
function collectRuns(node: StructNode, ctx: WalkCtx, skip?: Set<string>): RawRun[] {
    const out: RawRun[] = [];
    const recurse = (n: StructNode) => {
        if (n.type === 'content' && n.id) {
            const runs = ctx.runsByMcid.get(n.id);
            if (runs) { appendAll(out, runs); ctx.covered.add(n.id); }
            return;
        }
        for (const c of n.children || []) {
            if (ALWAYS_SKIP.has(role(c)) || (skip && skip.has(role(c)))) continue;
            recurse(c);
        }
    };
    recurse(node);
    return out;
}

/** Emits a block paragraph/heading, attaching any nested footnotes to its last text run. */
function blockWithNotes(node: StructNode, ctx: WalkCtx, level: number | undefined): OfficeContentNode[] {
    const out: OfficeContentNode[] = [];
    const p = paragraphFrom(node, ctx, level);
    if (p) out.push(p);
    const notes: OfficeContentNode[] = [];
    for (const noteNode of findNotes(node)) {
        if (ctx.opts.ignoreNotes) { markCovered(noteNode, ctx); continue; }
        const n = buildNote(noteNode, ctx);
        if (n) notes.push(n);
    }
    if (notes.length) {
        // Attach to the paragraph's last text child so generators anchor the reference inline (the
        // convention WordParser uses); fall back to trailing sibling nodes when there is no text run.
        const anchor = p ? lastTextChild(p) : undefined;
        if (anchor) anchor.notes = [...(anchor.notes || []), ...notes];
        else appendAll(out, notes);
    }
    return out;
}

/** The last descendant `text` node of a block, or undefined. */
function lastTextChild(node: OfficeContentNode): OfficeContentNode | undefined {
    let found: OfficeContentNode | undefined;
    const walk = (n: OfficeContentNode) => {
        if (n.type === 'text') found = n;
        for (const c of n.children || []) walk(c);
    };
    for (const c of node.children || []) walk(c);
    return found;
}

/** Finds Note/FENote subtrees anywhere under a node, without descending into `skip` roles. */
function findNotes(node: StructNode, skip?: Set<string>): StructNode[] {
    const out: StructNode[] = [];
    const recurse = (n: StructNode) => {
        for (const c of n.children || []) {
            if (ALWAYS_SKIP.has(role(c))) out.push(c);
            else if (!skip || !skip.has(role(c))) recurse(c);
        }
    };
    recurse(node);
    return out;
}

/** Marks every content leaf under a node as covered without emitting it (e.g. ignored notes). */
function markCovered(node: StructNode, ctx: WalkCtx): void {
    if (node.type === 'content' && node.id) ctx.covered.add(node.id);
    for (const c of node.children || []) markCovered(c, ctx);
}

function paragraphFrom(node: StructNode, ctx: WalkCtx, level: number | undefined): OfficeContentNode | null {
    const runs = collectRuns(node, ctx);
    if (!runs.length) return null;
    return runsToParagraph(runs, ctx.page, ctx.doc, level);
}

/**
 * Structure roles that group content without carrying table semantics. A producer may wrap a table's
 * row groups (or an individual row's cells) in one of these; descend through them so the row/cell is
 * still found, but never through a nested `Table`/`TR` (that would pull a sub-table's rows up into this
 * one).
 */
const TRANSPARENT_GROUP = new Set(['Div', 'NonStruct', 'Part', 'Sect', 'Art', 'Document', 'Group']);

/** Bounds the wrapper descent in {@link collectRows}/{@link collectCells} against a hostile tag tree. */
const MAX_TABLE_WRAPPER_DEPTH = 32;

function collectRows(node: StructNode, depth = 0): StructNode[] {
    const out: StructNode[] = [];
    if (depth > MAX_TABLE_WRAPPER_DEPTH) return out;
    for (const c of node.children || []) {
        const r = role(c);
        if (r === 'TR') out.push(c);
        else if (r === 'THead' || r === 'TBody' || r === 'TFoot' || TRANSPARENT_GROUP.has(r)) appendAll(out, collectRows(c, depth + 1));
    }
    return out;
}

/** Cells of a row: TH/TD directly under the TR, or under a transparent wrapper the producer inserted. */
function collectCells(tr: StructNode, depth = 0): StructNode[] {
    const out: StructNode[] = [];
    if (depth > MAX_TABLE_WRAPPER_DEPTH) return out;
    for (const c of tr.children || []) {
        const r = role(c);
        if (r === 'TH' || r === 'TD') out.push(c);
        else if (TRANSPARENT_GROUP.has(r)) appendAll(out, collectCells(c, depth + 1));
    }
    return out;
}

/** Median of a numeric list (robust to alignment jitter); 0 for an empty list. */
/**
 * True when a cell is an empty placeholder: no text and no geometry. A tagged PDF pads the columns
 * (and rows) a merged cell covers with exactly these, keeping the grid rectangular. They are the
 * only cells span inference is ever allowed to absorb.
 */
function isEmptyCell(cell: OfficeContentNode): boolean {
    return !cell.bounds && !(cell.text || '').trim();
}

/**
 * Recovers horizontal cell merges (colSpan) that the tag tree padded with empty placeholder cells.
 * pdf.js does not expose the `/ColSpan` attribute, so a merge appears as a wide non-empty cell
 * followed by empty placeholders in the columns it visually covers. When a non-empty cell's box
 * extends past the start of an adjacent empty column, that column is absorbed: the placeholder is
 * dropped and `colSpan` is set. Because only *empty* neighbours a cell's own geometry covers are
 * ever merged, a regular grid (whose cells sit inside their own column) is never given a spurious
 * span. Positional `col` indices are left untouched so the generator rebuilds the grid from
 * `col` + `colSpan`. No-op without geometry (cells then have no bounds to reason from).
 */
/**
 * Cell-count ceiling for merged-cell inference. Both passes are worst-case O(cells^2) (a hostile tall
 * single-column table of empty placeholders makes every row scan to the bottom), and the tagged-PDF
 * path has no `maxTableCells`-style budget of its own, so cap the input: past this the table still
 * renders, just without colSpan/rowSpan recovery. A genuine merged-cell table is far smaller.
 */
const MAX_SPAN_INFERENCE_CELLS = 5000;

function inferColSpans(rows: OfficeContentNode[]): void {
    // Representative left edge per grid column, from the non-empty cells that occupy it.
    const colLefts: number[][] = [];
    for (const row of rows) {
        const cells = row.children || [];
        for (let c = 0; c < cells.length; c++) {
            const b = cells[c].bounds;
            if (b) (colLefts[c] ||= []).push(b.x);
        }
    }
    const colLeft = (c: number): number => {
        const xs = colLefts[c];
        return xs && xs.length ? median(xs) : NaN;
    };
    for (const row of rows) {
        const cells = row.children || [];
        const kept: OfficeContentNode[] = [];
        let c = 0;
        while (c < cells.length) {
            const cell = cells[c];
            if (isEmptyCell(cell)) { kept.push(cell); c++; continue; }
            const b = cell.bounds;
            let span = 1;
            if (b) {
                const right = b.x + b.width;
                // Absorb consecutive empty columns to the right whose start this cell's box passes.
                while (c + span < cells.length && isEmptyCell(cells[c + span])) {
                    const nextLeft = colLeft(c + span);
                    if (!Number.isFinite(nextLeft) || right <= nextLeft + 1) break;
                    span++;
                }
            }
            if (span > 1 && cell.metadata) (cell.metadata as CellMetadata).colSpan = span;
            kept.push(cell);
            c += span;
        }
        row.children = kept;
    }
}

/**
 * Recovers vertical cell merges (rowSpan) the tag tree padded with empty placeholder cells. pdf.js
 * does not expose the `/RowSpan` attribute, so a merge appears as a tall non-empty cell with empty
 * placeholders in the rows below it, at the same grid column. When a non-empty cell's box extends
 * well past the top of the next row's band and the cell directly below (same grid column) is an
 * empty placeholder, that placeholder is absorbed: it is dropped and `rowSpan` grows. Only detects
 * merges whose spanning cell has content tall enough to reach into the rows it covers - a vertically
 * merged cell with a single centred line is invisible to geometry and left as a full grid (correct,
 * just not marked as merged). Runs on the grid keyed by the positional `col`, which the tag tree
 * keeps rectangular, so it stays aligned with {@link inferColSpans}.
 */
function inferRowSpans(rows: OfficeContentNode[]): void {
    if (rows.length < 2) return;
    const byCol: Map<number, OfficeContentNode>[] = [];
    const rowTop: number[] = [];
    for (const row of rows) {
        const m = new Map<number, OfficeContentNode>();
        let top = Infinity;
        for (const cell of row.children || []) {
            const col = (cell.metadata as CellMetadata)?.col;
            if (typeof col === 'number' && !m.has(col)) m.set(col, cell);
            if (cell.bounds) top = Math.min(top, cell.bounds.y);
        }
        byCol.push(m);
        rowTop.push(Number.isFinite(top) ? top : NaN);
    }
    // Typical row pitch drives the coverage margin, so a cell must clearly enter the next band
    // (not merely touch its top, which every ordinary cell does) to count as spanning.
    const pitches: number[] = [];
    for (let r = 1; r < rowTop.length; r++) {
        if (Number.isFinite(rowTop[r]) && Number.isFinite(rowTop[r - 1])) pitches.push(rowTop[r] - rowTop[r - 1]);
    }
    const pitch = median(pitches.filter(p => p > 0));
    if (!(pitch > 0)) return;
    const margin = 0.4 * pitch;

    const toRemove = new Set<OfficeContentNode>();
    for (let r = 0; r < rows.length; r++) {
        for (const [col, cell] of byCol[r]) {
            if (toRemove.has(cell) || isEmptyCell(cell) || !cell.bounds) continue;
            const cSpan = (cell.metadata as CellMetadata)?.colSpan || 1;
            const bottom = cell.bounds.y + cell.bounds.height;
            let span = 1;
            for (let k = 1; r + k < rows.length; k++) {
                const below = byCol[r + k].get(col);
                if (!below || !isEmptyCell(below)) break;
                if (!Number.isFinite(rowTop[r + k]) || bottom < rowTop[r + k] + margin) break;
                toRemove.add(below);
                // A 2-D merge (this cell also spans columns) leaves placeholders under every column
                // it covers in the rows below; absorb those too, or they render as a phantom column.
                for (let cc = col + 1; cc < col + cSpan; cc++) {
                    const extra = byCol[r + k].get(cc);
                    if (extra && isEmptyCell(extra)) toRemove.add(extra);
                }
                span++;
            }
            if (span > 1 && cell.metadata) (cell.metadata as CellMetadata).rowSpan = span;
        }
    }
    if (toRemove.size) {
        for (const row of rows) row.children = (row.children || []).filter(c => !toRemove.has(c));
    }
}

/**
 * A tagged table's cells in the columns they stand in. A producer leaves out the cells a merged one covers
 * (Chrome and LibreOffice), and a row's empty cells (Chrome), so a row's n-th cell is not its n-th column:
 * a calendar's first week (four days, from a Thursday) was read from Monday, and a merged header's next
 * cell a column early. Where the page rules the table (see rulings), its frame says each cell's column and
 * spans. Else the columns are the full rows' (those of the most cells), each where its cells start (end,
 * read right to left), and a shorter row's cells go, in their order, to the columns they start nearest; a
 * cell whose text runs past the start of a column no cell of its row takes spans it. (A merged cell whose
 * text does not reach past its first column says nothing of its span without rules: the columns after it
 * are left empty.) A table whose rows are all full is left as tagged, and so is one whose cells cannot be
 * placed in order. Returns whether the cells were placed.
 */
function placeTaggedCells(rows: OfficeContentNode[], frames: readonly RuledFrame[], rtl: boolean): boolean {
    const width = Math.max(...rows.map(r => r.children?.length ?? 0));
    if (rows.every(r => (r.children?.length ?? 0) === width)) return false;
    const boxOf = (cell: OfficeContentNode) => CELL_PAGE_BOX.get(cell);
    // Where a cell starts and ends, in the table's reading direction (right to left as negative x).
    const start = (b: NodeBounds) => rtl ? -(b.x + b.width) : b.x;
    const end = (b: NodeBounds) => rtl ? -b.x : b.x + b.width;
    const table = unionAll(rows.flatMap(r => (r.children ?? []).map(boxOf)));
    if (!table) return false;

    // The ruled frame around the table (the smallest, so a table inside another's cell is its own), with as
    // many columns as the full rows have cells.
    const SNAP = 2;
    const frame = frames
        .filter(f => f.xs.length - 1 === width && f.ys.length - 1 === rows.length && table.x >= f.box.x - SNAP && table.y >= f.box.y - SNAP
            && table.x + table.width <= f.box.x + f.box.width + SNAP && table.y + table.height <= f.box.y + f.box.height + SNAP)
        .sort((a, b) => a.box.width * a.box.height - b.box.width * b.box.height)[0];
    const places: { col: number; colSpan: number; rowSpan: number }[][] = [];
    let ok = !!frame;
    if (frame) {
        // Each grid position's frame cell, and a frame cell's column in reading order.
        const owner = Array.from({ length: rows.length }, () => new Array<number>(width).fill(-1));
        frame.cells.forEach((c, k) => { for (let r = c.row; r < c.row + c.rowSpan; r++) for (let q = c.col; q < c.col + c.colSpan; q++) owner[r][q] = k; });
        const colOf = (c: RuledFrame['cells'][number]) => rtl ? width - c.col - c.colSpan : c.col;
        const taken = new Set<number>();
        for (const [ri, row] of rows.entries()) {
            const out: { col: number; colSpan: number; rowSpan: number }[] = [];
            let last = -1;
            for (const cell of row.children ?? []) {
                const b = boxOf(cell);
                let at = -1;
                if (b) {
                    // The frame cell its middle is in, starting in its row.
                    const cx = b.x + b.width / 2, cy = b.y + b.height / 2;
                    let lo = 0, hi = width - 1;
                    while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (frame.xs[mid] <= cx) lo = mid; else hi = mid - 1; }
                    const k = owner[ri][lo], c = frame.cells[k];
                    if (c && c.row === ri && cy >= frame.ys[c.row] && cy <= frame.ys[c.row + c.rowSpan]) at = k;
                } else {
                    // An empty cell: the row's next frame cell, in reading order.
                    for (let q = last + 1; q < width && at < 0; q++) {
                        const k = owner[ri][rtl ? width - 1 - q : q];
                        if (k >= 0 && frame.cells[k].row === ri && !taken.has(k)) at = k;
                    }
                }
                const fc = frame.cells[at];
                if (!fc || taken.has(at) || colOf(fc) <= last) { ok = false; break; }
                taken.add(at);
                out.push({ col: colOf(fc), colSpan: fc.colSpan, rowSpan: fc.rowSpan });
                last = colOf(fc) + fc.colSpan - 1;
            }
            if (!ok) break;
            places.push(out);
        }
    }
    if (!ok) {
        places.length = 0;
        // The full rows' columns: where each one's cells start.
        const full = rows.filter(r => r.children!.length === width);
        const starts: number[] = [];
        for (let k = 0; k < width; k++) {
            const at = full.map(r => boxOf(r.children![k])).filter((b): b is NodeBounds => !!b).map(start);
            if (!at.length) return false;
            starts.push(median(at));
            if (k && !(starts[k] > starts[k - 1])) return false;
        }
        for (const row of rows) {
            const cells = row.children ?? [];
            const cols: number[] = [];
            let last = -1;
            for (const [i, cell] of cells.entries()) {
                const b = boxOf(cell);
                // (Room left for the row's cells after this one.)
                const room = width - (cells.length - i);
                let col = last + 1;
                if (col > room) return false;
                if (b && cells.length < width) {
                    // The column of those it may take (from the one after the last, leaving room) it starts nearest.
                    let lo = col, hi = room;
                    while (lo < hi) { const mid = (lo + hi) >> 1; if (starts[mid] < start(b)) lo = mid + 1; else hi = mid; }
                    col = lo > col && Math.abs(start(b) - starts[lo - 1]) <= Math.abs(start(b) - starts[lo]) ? lo - 1 : lo;
                }
                cols.push(col);
                last = col;
            }
            places.push(cols.map((col, i) => {
                // Spanning the columns its text runs into that no cell of the row takes.
                const b = boxOf(cells[i]), next = i + 1 < cols.length ? cols[i + 1] : width;
                let lo = col + 1, hi = next;
                while (b && lo < hi) { const mid = (lo + hi) >> 1; if (end(b) > starts[mid] + 1) lo = mid + 1; else hi = mid; }
                return { col, colSpan: b ? lo - col : 1, rowSpan: 1 };
            }));
        }
    }
    rows.forEach((row, ri) => (row.children ?? []).forEach((cell, i) => {
        const meta = cell.metadata as CellMetadata;
        const p = places[ri][i];
        meta.col = p.col;
        if (p.colSpan > 1) meta.colSpan = p.colSpan;
        if (p.rowSpan > 1) meta.rowSpan = p.rowSpan;
    }));
    return true;
}

function buildTable(node: StructNode, ctx: WalkCtx): OfficeContentNode | null {
    const rows: OfficeContentNode[] = [];
    const tableRuns: RawRun[] = [];
    let rowIdx = 0;
    for (const tr of collectRows(node)) {
        const cells: OfficeContentNode[] = [];
        let colIdx = 0;
        for (const cellNode of collectCells(tr)) {
            const cr = role(cellNode);
            const cellChildren = walkChildren(cellNode, ctx, 0);
            const meta: CellMetadata = { row: rowIdx, col: colIdx };
            if (cr === 'TH') meta.style = 'header';
            const cell: OfficeContentNode = {
                type: 'cell',
                text: cellChildren.map(n => n.text || '').join(' ').trim(),
                children: cellChildren,
                metadata: meta,
            };
            const cb = unionAll(cellChildren.map(c => c.bounds));
            if (cb) cell.bounds = cb;
            const cellRuns = collectRuns(cellNode, ctx);
            appendAll(tableRuns, cellRuns);
            // (Its text's box: a blank a producer set after the cell before it, at that cell's end, is not where it stands.)
            const pageBox = unionAll(cellRuns.filter(r => r.text.trim()).map(r => ({ x: r.x, y: r.yTop, width: r.width, height: r.height })));
            if (pageBox) CELL_PAGE_BOX.set(cell, pageBox);
            cells.push(cell);
            colIdx++;
        }
        if (!cells.length) continue;
        const row: OfficeContentNode = { type: 'row', children: cells, text: cells.map(c => c.text || '').join(' ').trim() };
        const rb = unionAll(cells.map(c => c.bounds));
        if (rb) row.bounds = rb;
        rows.push(row);
        rowIdx++;
    }
    if (!rows.length) return null;
    // Recover merged cells the tags padded with empty placeholders (needs geometry). Column spans
    // run first, on the still-rectangular grid (they use positional indices); row spans run after,
    // keyed by the geometry-stable `col`, so they tolerate the placeholders columns already dropped.
    const totalCells = rows.reduce((s, r) => s + (r.children?.length || 0), 0);
    // A row with fewer cells than the table's fullest is placed by where its cells stand (see placeTaggedCells);
    // the placeholder inference below is for a grid the tags kept whole.
    const placed = totalCells <= MAX_SPAN_INFERENCE_CELLS && placeTaggedCells(rows, ctx.opts.frames?.() ?? [], rightToLeftLean(tableRuns) > 0);
    if (!placed && ctx.doc.cfg.includeBounds && totalCells <= MAX_SPAN_INFERENCE_CELLS) { inferColSpans(rows); inferRowSpans(rows); }
    const table: OfficeContentNode = { type: 'table', children: rows, text: rows.map(r => r.text || '').join('\n') };
    const tb = unionAll(rows.map(r => r.bounds));
    if (tb) table.bounds = tb;
    return table;
}

/** Classifies a list marker label into an ordered/unordered list type. */
function classifyListType(label: string): 'ordered' | 'unordered' {
    const t = label.trim();
    // Decimal (incl. multilevel "1.1."), single alpha "a.", or roman "iii." markers are ordered.
    if (/^\(?\d+(\.\d+)*[.)]?$/.test(t) || /^[a-zA-Z][.)]$/.test(t) || /^[ivxlcdmIVXLCDM]+[.)]$/.test(t)) return 'ordered';
    return 'unordered';
}

/** Collapses TOC dot-leaders ("Title ...... 3" -> "Title 3") in a node's text and text-run children. */
function stripDotLeaders(node: OfficeContentNode): void {
    // One leader, however many runs of dots it is split into ("Scope .... .... 12"), is one space. A
    // match starts where whitespace does, not at each space of a run (each read to its end); once its
    // first dots are found nothing after them can fail, so it never goes back over what it read.
    const clean = (s: string | undefined) => (s || '').replace(/(?<!\s)\s*\.{4,}(?:\s*\.{4,})*\s*/g, ' ');
    if (node.text) node.text = clean(node.text).trim();
    for (const c of node.children || []) {
        if (c.type === 'text') c.text = clean(c.text);
        else stripDotLeaders(c);
    }
}

/** Converts a roman numeral to its integer value, or 0 if it is not a valid roman numeral. */
function romanToInt(s: string): number {
    const map: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
    const t = s.toLowerCase();
    let total = 0, prev = 0;
    for (let i = t.length - 1; i >= 0; i--) {
        const v = map[t[i]];
        if (!v) return 0;
        total += v < prev ? -v : v;
        prev = v;
    }
    return total;
}

/**
 * Parses the number an ordered-list label represents: decimal ("3."), multi-level decimal ("1.2.",
 * uses the last component), roman ("iii."), or single alpha ("c." -> 3). Returns null when the label
 * carries no recognizable number.
 */
function parseListNumber(label: string): number | null {
    const t = label.trim();
    if (!t) return null;
    // All-decimal (possibly multi-level like "1.2."): use the last numeric component.
    if (/^[\d.()\[\]\s]+$/.test(t)) {
        const parts = t.match(/\d+/g);
        if (parts && parts.length) return parseInt(parts[parts.length - 1], 10);
    }
    const core = t.replace(/^[(\[]+/, '').replace(/(?<![.)\]\s])[.)\]\s]+$/, '');
    if (/^[ivxlcdm]+$/i.test(core)) { const n = romanToInt(core); return n > 0 ? n : null; }
    if (/^[a-z]$/i.test(core)) return core.toLowerCase().charCodeAt(0) - 96; // a -> 1
    return null;
}

/** Roles inside a list body that build their own nodes, so the item must not absorb their content. */
const NESTED_IN_LBODY = new Set(['L', 'Table']);

function buildList(node: StructNode, ctx: WalkCtx, indent: number): OfficeContentNode[] {
    const items: OfficeContentNode[] = [];
    const listId = `pdf-list-${++ctx.listCounter.n}`;
    let idx = 0;
    for (const li of (node.children || [])) {
        if (role(li) !== 'LI') {
            // Some producers nest content directly; recurse transparently.
            if (role(li) === 'L') appendAll(items, buildList(li, ctx, indent + 1));
            continue;
        }
        let label = '';
        const bodyNodes: OfficeContentNode[] = [];
        const nested: OfficeContentNode[] = [];
        const itemNotes: OfficeContentNode[] = [];
        for (const part of li.children || []) {
            const pr = role(part);
            if (pr === 'Lbl') {
                label = collectRuns(part, ctx).map(r => r.text).join('').trim();
            } else if (pr === 'L') {
                appendAll(nested, buildList(part, ctx, indent + 1));
            } else {
                // The item body: an LBody wrapper, OR the body placed directly under LI (a P/Span with
                // no LBody, which some producers emit). Treat any non-Lbl/non-L child as body so its
                // runs are covered and the item is never dropped as empty (which left the marker looking
                // like stray text outside the tag tree and fired a spurious PDF_STRUCT_TREE_UNRELIABLE).
                const directRuns = collectRuns(part, ctx, NESTED_IN_LBODY);
                if (directRuns.length) { const p = runsToParagraph(directRuns, ctx.page, ctx.doc, 0); if (p) bodyNodes.push(p); }
                // Footnotes hanging off the item text, collected exactly as a paragraph's are. Without
                // this their runs stay uncovered: they come back as stray paragraphs spliced into the
                // page and, worse, make the page look like it has text outside the tag tree. Nested
                // lists and tables build their own notes, so this must not descend into them.
                for (const noteNode of findNotes(part, NESTED_IN_LBODY)) {
                    if (ctx.opts.ignoreNotes) { markCovered(noteNode, ctx); continue; }
                    const n = buildNote(noteNode, ctx);
                    if (n) itemNotes.push(n);
                }
                for (const b of part.children || []) {
                    const br = role(b);
                    if (br === 'L') appendAll(nested, buildList(b, ctx, indent + 1));
                    else if (br === 'Table') { const t = buildTable(b, ctx); if (t) bodyNodes.push(t); }
                }
            }
        }
        const listType = classifyListType(label);
        // Number ordered items from the label the PDF actually rendered (decimal "3.", roman "iii.",
        // or alpha "c.") rather than the positional index, so a list interrupted by a paragraph and
        // split into two <L> elements resumes its numbering (3, 4) instead of restarting at 1. Falls
        // back to the position when the label carries no parseable number.
        let itemIndex = idx;
        if (listType === 'ordered') {
            const n = parseListNumber(label);
            if (n !== null && n >= 1) itemIndex = n - 1;
        }
        const meta: ListMetadata = {
            listType,
            indentation: indent,
            alignment: 'left',
            listId,
            itemIndex,
        };
        // The label as printed, where it is other than the item's number and a dot or a round bullet.
        const printed = label.trim();
        if (printed && printed !== (listType === 'ordered' ? `${itemIndex + 1}.` : '•')) meta.marker = printed;
        const item: OfficeContentNode = {
            type: 'list',
            text: bodyNodes.map(n => n.text || '').join(' ').trim(),
            children: bodyNodes,
            metadata: meta,
        };
        const b = unionAll(bodyNodes.map(n => n.bounds));
        if (b) item.bounds = b;
        if (itemNotes.length) {
            // Anchor on the item's last text run so generators render the citation inline, as the
            // paragraph path does; with no text run to hang them on, keep them as trailing children.
            const anchor = lastTextChild(item);
            if (anchor) anchor.notes = [...(anchor.notes || []), ...itemNotes];
            else item.children = [...(item.children || []), ...itemNotes];
        }
        // Skip an LI that carried no body text and no notes (a producer that wrapped only a nested
        // list in an LI, or an empty LI): emitting an empty list node just adds a blank bullet and
        // leaves its (uncovered) marker looking like text outside the tag tree. Any nested sublist is
        // still emitted.
        if (bodyNodes.length || itemNotes.length) { items.push(item); idx++; }
        appendAll(items, nested);
    }
    return items;
}

function buildNote(node: StructNode, ctx: WalkCtx): OfficeContentNode | null {
    const children = walkChildren(node, ctx, 0);
    if (!children.length) {
        const p = paragraphFrom(node, ctx, 0);
        if (!p) return null;
        children.push(p);
    }
    const meta: NoteMetadata = { noteType: 'footnote' };
    // Strip the leading marker glyph the PDF renders inside the note body ("1 In paged media…") so it
    // is not duplicated next to the generated citation ("[^1]: 1 In paged…"); keep it as the note id.
    // A number or symbol marker may drop its trailing punctuation ("1 In paged...", a dagger),
    // but a roman-numeral or single-letter marker MUST carry `.`/`)` - otherwise a note body that
    // simply begins with a word like "Did", "I", "A" or "Civil" would have its first word eaten.
    const markerRe = /^\s*(?:([0-9]+|[*†‡§])[.)]?|([ivxlcdm]+|[a-z])[.)])\s+/i;
    const firstText = children.map(n => n.text || '').join(' ').trim();
    const mk = firstText.match(markerRe);
    if (mk) {
        meta.noteId = mk[1] || mk[2];
        const strip = (n: OfficeContentNode): boolean => {
            if (n.type === 'text' && n.text) { const m = n.text.match(markerRe); if (m) { n.text = n.text.slice(m[0].length); return true; } return false; }
            if (n.text) { const m = n.text.match(markerRe); if (m) n.text = n.text.slice(m[0].length); }
            for (const c of n.children || []) if (strip(c)) return true;
            return false;
        };
        for (const child of children) if (strip(child)) break;
    }
    const note: OfficeContentNode = { type: 'note', text: children.map(n => n.text || '').join(' ').trim(), children, metadata: meta };
    const b = unionAll(children.map(n => n.bounds));
    if (b) note.bounds = b;
    return note;
}

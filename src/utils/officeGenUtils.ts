/**
 * Shared helpers for the office-package generators (DOCX, ODT, EPUB).
 *
 * These are the pure, format-agnostic pieces every hand-built package writer needs: base64 image
 * decoding, image-size sniffing, length/colour/date normalisation, and the reproducible zip instant.
 * They live here so DocxGenerator, OdtGenerator and EpubGenerator share one implementation instead
 * of drifting copies.
 *
 * @module officeGenUtils
 */

import { CellMetadata, EmbedMetadata, HeaderFooterMetadata, ImageMode, OfficeContentNode, OfficeParserAST, PageSetup, ParagraphSpacing, RowMetadata, SheetMetadata } from '../types.js';
import { cssColorHex } from './colorUtils.js';
import { escapeXml, stripInvalidXmlChars } from './sanitize.js';
import { lookupTable } from './lookupUtils.js';
import { hasPageSize } from './pageLimits.js';

/**
 * Resolves `config.includeImages` - a boolean, a CLI-provided `'true'`/`'false'` string, or an
 * {@link ImageMode} - to a single mode. Shared by every generator (via `BaseGenerator.imageMode()`
 * and the standalone native PDF engine) so image handling never drifts between them, and so a truthy
 * mode string such as `'none'` is never mistaken for "include images". Unknown values default to
 * `'image-only'`, the historical behaviour of a plain `true`.
 */
export function resolveImageMode(includeImages: unknown): ImageMode {
    let v = includeImages;
    if (v === 'true') v = true;
    if (v === 'false') v = false;
    if (v === false) return 'none';
    if (v === true || v === undefined) return 'image-only';
    if (v === 'image-only' || v === 'image+ocr-text' || v === 'ocr-text-only' || v === 'none') return v;
    return 'image-only';
}

/**
 * Whether a table row should be treated as a header row, mirroring the HtmlGenerator heuristic so
 * every generator agrees. A field a parser actually sets (never the test-only `isHeader`): the row's
 * or a cell's `style` containing "header" (PDF `TH` cells set `style: 'header'`), an explicit
 * `isHeader` flag, or - for the first row only, where it is a reliable signal - every cell bold.
 */
export function isHeaderRow(row: OfficeContentNode, isFirstRow: boolean): boolean {
    const cells = (row.children || []).filter(c => c.type === 'cell');
    if (!cells.length) return false;
    const meta = row.metadata as any;
    if (meta?.isHeader) return true;
    if (typeof meta?.style === 'string' && meta.style.toLowerCase().includes('header')) return true;
    const cellIsHeader = (c: OfficeContentNode) => {
        const cm = c.metadata as any;
        if (cm?.isHeader) return true;
        return typeof cm?.style === 'string' && cm.style.toLowerCase().includes('header');
    };
    if (cells.every(cellIsHeader)) return true;
    if (isFirstRow && cells.every(c => !!c.children?.length && c.children.every(ch => ch.formatting?.bold === true))) return true;
    return false;
}

/**
 * Largest grid a sparse spreadsheet may be expanded back into by {@link fillSheetRowGaps}. A sheet
 * whose only populated rows are 1 and 50,000 is a real (if unusual) file, and materialising every
 * empty row between them as table XML would dwarf the document; past this many rows the gaps are
 * left unfilled, which is the historical behaviour.
 */
const MAX_SHEET_GRID_ROWS = 4096;

/**
 * Re-inserts the empty rows a sparse spreadsheet AST leaves out.
 *
 * `ExcelParser` (and the ODS reader) emit a `row` node only for rows with at least one non-empty
 * cell, and each cell carries its absolute `{ row, col }`, so a sheet with values in rows 1 and 3
 * arrives as two adjacent `row` nodes. Rendering those as a two-row table silently moves the data up
 * a row. `HtmlGenerator` rebuilds the full grid from the indices; this gives the DOCX and ODT table
 * writers the same view without each re-deriving it.
 *
 * The synthesized row holds a single empty cell placed at the grid's last column, which the sparse
 * placement in those writers expands into a full-width run of empty cells.
 *
 * Rows whose cells carry no `row` index (an ordinary document table) are returned untouched.
 */
export function fillSheetRowGaps(rows: OfficeContentNode[], take: (positions: number) => boolean = () => true): OfficeContentNode[] {
    const cellsOf = (row: OfficeContentNode) => (row.children || []).filter(c => c.type === 'cell');
    const rowIndex = (row: OfficeContentNode): number | null => {
        const own = (row.metadata as any)?.row;
        if (typeof own === 'number') return own;
        for (const c of cellsOf(row)) {
            const r = (c.metadata as any)?.row;
            if (typeof r === 'number') return r;
        }
        return null;
    };
    if (!rows.some(r => rowIndex(r) !== null)) return rows;

    let maxCol = 0;
    for (const row of rows) {
        for (const c of cellsOf(row)) {
            const meta = c.metadata as any;
            if (typeof meta?.col === 'number') maxCol = Math.max(maxCol, meta.col + Math.max(1, meta.colSpan || 1) - 1);
        }
    }

    const out: OfficeContentNode[] = [];
    let expected = 0;
    for (const row of rows) {
        const idx = rowIndex(row);
        // Each filled row is as wide as the grid, so the rows are taken from the document's grid budget
        // (`take`): an empty row naming row 4,094 in a sheet 1,000 columns wide made 270 MB of DOCX.
        if (idx !== null && idx > expected && idx - expected + out.length <= MAX_SHEET_GRID_ROWS && take((idx - expected) * (maxCol + 1))) {
            for (let r = expected; r < idx; r++) {
                out.push({ type: 'row', children: [{ type: 'cell', metadata: { row: r, col: maxCol } as any, children: [] }] });
            }
        }
        out.push(row);
        expected = (idx === null ? expected : idx) + 1;
    }
    return out;
}

/**
 * Percent-encodes URL-unsafe characters (space, `<>"{}|^[]` and the like) for an href/Target sink.
 * `encodeURI` leaves existing `%xx` escapes and structural characters (`/?:@&=#`) intact, so it is
 * idempotent for already-valid URLs; Word rejects some of those raw characters in a relationship
 * Target with an "unreadable content" dialog. Falls back to the raw string if encoding throws.
 */
export function encUrl(url: string): string {
    // Only what a URL cannot hold is encoded (spaces, quotes, angle brackets, control and non-ASCII
    // characters, a `%` that starts no escape): an address already encoded (`a%20b`) was encoded again
    // (`a%2520b`), and opened another address.
    let out = '';
    for (let i = 0; i < url.length; i++) {
        const ch = url[i], code = url.charCodeAt(i);
        if (ch === '%') { out += /^[0-9a-fA-F]{2}$/.test(url.slice(i + 1, i + 3)) ? '%' : '%25'; continue; }
        if (code > 0x20 && code < 0x7f && !'"<>\\^`{|}'.includes(ch)) { out += ch; continue; }
        // A character beyond ASCII as its UTF-8 bytes (a surrogate pair as one character, a lone half as U+FFFD).
        const pair = code >= 0xd800 && code <= 0xdbff && i + 1 < url.length && url.charCodeAt(i + 1) >= 0xdc00 && url.charCodeAt(i + 1) <= 0xdfff;
        const whole = pair ? url.slice(i, i + 2) : code >= 0xd800 && code <= 0xdfff ? '�' : ch;
        if (pair) i++;
        out += encodeURIComponent(whole);
    }
    return out;
}

/** Decodes a base64 string to raw bytes, cross-env (atob exists in Node 16+ and browsers). */
export function decodeBase64(b64: string): Uint8Array {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
}

/**
 * Decoded byte length of a base64 string, computed from its length and padding without decoding.
 * (Standard base64 is 4 chars per 3 bytes; each trailing `=` drops one byte.) Assumes clean base64
 * with no embedded whitespace, which is how attachment data is stored.
 */
export function base64ByteLength(b64: string | undefined): number {
    if (!b64) return 0;
    const len = b64.length;
    let padding = 0;
    if (len >= 1 && b64.charCodeAt(len - 1) === 0x3d) padding++; // '='
    if (len >= 2 && b64.charCodeAt(len - 2) === 0x3d) padding++;
    return Math.max(0, Math.floor(len * 3 / 4) - padding);
}

const PT_PER_IN = 72;
const PT_PER_MM = 72 / 25.4;

/**
 * Portrait page dimensions in PostScript points (1/72 inch), keyed by lowercased paper-format name.
 * The A-series is ISO 216 (mm-derived); Letter/Legal/Tabloid/Ledger are US/ANSI (inch-derived).
 * Single source of truth for {@link paperSizePt}, so every generator lays out the same size the
 * same way.
 */
const PAPER_SIZES_PT: Record<string, { w: number; h: number }> = lookupTable({
    letter: { w: 8.5 * PT_PER_IN, h: 11 * PT_PER_IN },
    legal: { w: 8.5 * PT_PER_IN, h: 14 * PT_PER_IN },
    tabloid: { w: 11 * PT_PER_IN, h: 17 * PT_PER_IN },
    // Ledger and Tabloid are the same ANSI B sheet in the two orientations: Tabloid is portrait
    // (11x17), Ledger is landscape (17x11). Keeping Ledger at 17x11 matches Chromium/Puppeteer, which
    // is the default HTML PDF engine and defines `ledger` as 17x11, so every engine agrees on the size.
    // (This is the one non-portrait entry; the generators that assume portrait swap via `landscape` -
    // see the native engine and the DOCX/ODT `landscape ? h : w` orientation, which both key off it.)
    ledger: { w: 17 * PT_PER_IN, h: 11 * PT_PER_IN },
    a0: { w: 841 * PT_PER_MM, h: 1189 * PT_PER_MM },
    a1: { w: 594 * PT_PER_MM, h: 841 * PT_PER_MM },
    a2: { w: 420 * PT_PER_MM, h: 594 * PT_PER_MM },
    a3: { w: 297 * PT_PER_MM, h: 420 * PT_PER_MM },
    a4: { w: 210 * PT_PER_MM, h: 297 * PT_PER_MM },
    a5: { w: 148 * PT_PER_MM, h: 210 * PT_PER_MM },
    a6: { w: 105 * PT_PER_MM, h: 148 * PT_PER_MM },
});

/**
 * The document's language as a BCP 47 tag for output: `metadata.language` (where parsers put it),
 * else `nativeProperties.language`, else `fallback`. Reduced to the characters a language tag uses,
 * since it is written into markup (`lang`, `dc:language`, `/Lang`).
 */
export function documentLanguage(metadata: { language?: string; nativeProperties?: Record<string, any> } | undefined, fallback = 'en'): string {
    const raw = metadata?.language || metadata?.nativeProperties?.language;
    const tag = typeof raw === 'string' ? raw.replace(/[^A-Za-z0-9-]/g, '') : '';
    return tag || fallback;
}

/** Every {@link PaperFormat} name, lowercased: the page sizes a `format` option accepts. */
export const PAPER_FORMAT_NAMES: readonly string[] = [...Object.keys(PAPER_SIZES_PT), 'auto'];

/**
 * The header or footer a writer with one for the whole document writes (DOCX, ODT, RTF, LaTeX): the
 * first section's default one, shown, once; else the first shown one. A reader that gives loose blocks
 * (no header or footer node) gives them all. Every part's and every page's copy were written into the
 * one header: a Word document's three parts, a PDF's header on each of its pages.
 */
export function runningFurniture(nodes: OfficeContentNode[] | undefined, kind: 'header' | 'footer'): OfficeContentNode[] {
    if (!nodes?.length) return [];
    // Found once a list (writers ask for it at every section of thousands).
    let byKind = runningByList.get(nodes);
    if (!byKind) runningByList.set(nodes, byKind = new Map());
    const known = byKind.get(kind);
    if (known) return known;
    const found = findRunningFurniture(nodes, kind);
    byKind.set(kind, found);
    return found;
}
const runningByList = new WeakMap<OfficeContentNode[], Map<string, OfficeContentNode[]>>();

/**
 * The furniture found for these lists forgotten (see runningFurniture, sectionFurnitureIndex): asked
 * at the start of each generation, so a list edited in place since the last is read afresh.
 */
export function forgetFurniture(...lists: (OfficeContentNode[] | undefined)[]): void {
    for (const nodes of lists) if (nodes) { runningByList.delete(nodes); furnitureIndexes.delete(nodes); }
}

function findRunningFurniture(nodes: OfficeContentNode[], kind: 'header' | 'footer'): OfficeContentNode[] {
    const loose = nodes.filter(n => n.type !== kind);
    if (loose.length) return loose;
    const meta = (n: OfficeContentNode) => n.metadata as { type?: string; section?: number; hidden?: boolean; page?: number } | undefined;
    // One with something in it: an empty default header did not stand for a first page's letterhead.
    const filled = (n: OfficeContentNode, depth = 0): boolean => !!(n.text || '').trim() || (depth < 8 && (n.children || []).some(c => c.type === 'image' || c.type === 'chart' || filled(c, depth + 1)));
    const shown = nodes.filter(n => !meta(n)?.hidden && filled(n));
    const first = shown.filter(n => (meta(n)?.section ?? 0) === 0);
    // (A document of sections whose first has none of its own has none on its first pages: a later
    // section's is that section's, not the first's.)
    const sectioned = nodes.some(n => typeof meta(n)?.section === 'number' && meta(n)!.section! > 0);
    const pick = first.find(n => (meta(n)?.type ?? 'default') === 'default') ?? first[0] ?? (sectioned ? undefined : shown[0]);
    if (!pick) return [];
    // A header read from a page (a PDF's) is that page's paragraphs, each a node: all of them.
    const page = meta(pick)?.page;
    return page === undefined ? [pick] : shown.filter(n => meta(n)?.page === page && (meta(n)?.type ?? 'default') === (meta(pick)?.type ?? 'default'));
}

/**
 * A later section's own header or footer (its every-page part, `section` counted from 0, the first
 * section's being {@link runningFurniture}'s): the part a Word section has of its own, an empty one
 * included (its pages have none); undefined where the section has none of its own, and takes on the one
 * before it, as a word processor's section does. A reader that gives loose blocks, or a page's (a
 * PDF's), gives none.
 */
export function sectionFurniture(nodes: OfficeContentNode[] | undefined, kind: 'header' | 'footer', section: number, type: 'default' | 'first' | 'even' = 'default'): OfficeContentNode[] | undefined {
    if (!nodes?.length || section < (type === 'default' ? 1 : 0)) return undefined;
    const own = sectionFurnitureIndex(nodes, kind, type).bySection.get(section);
    return own ? [own] : undefined;
}

/**
 * Sections' own parts of a kind and type, by section, and their sections in order: read once a list (see
 * sectionFurniture). Every-page parts are later sections' (the first section's is runningFurniture's);
 * a first page's or even pages' are the first section's too, but for the one runningFurniture picked.
 */
const furnitureIndexes = new WeakMap<OfficeContentNode[], Map<string, { bySection: Map<number, OfficeContentNode>; sections: number[] }>>();
export function sectionFurnitureIndex(nodes: OfficeContentNode[], kind: 'header' | 'footer', type: 'default' | 'first' | 'even' = 'default'): { bySection: Map<number, OfficeContentNode>; sections: number[] } {
    let byKind = furnitureIndexes.get(nodes);
    if (!byKind) furnitureIndexes.set(nodes, byKind = new Map());
    let index = byKind.get(`${kind}|${type}`);
    if (!index) {
        const bySection = new Map<number, OfficeContentNode>();
        const running = type === 'default' ? undefined : runningFurniture(nodes, kind)[0];
        for (const n of nodes) {
            const m = n.metadata as { type?: string; section?: number; hidden?: boolean; page?: number } | undefined;
            if (n.type !== kind || n === running || m?.page !== undefined || m?.hidden || (m?.type ?? 'default') !== type) continue;
            const section = typeof m?.section === 'number' ? m.section : type === 'default' ? -1 : 0;
            if (section < (type === 'default' ? 1 : 0)) continue;
            if (!bySection.has(section)) bySection.set(section, n);
        }
        index = { bySection, sections: [...bySection.keys()].sort((a, b) => a - b) };
        byKind.set(`${kind}|${type}`, index);
    }
    return index;
}

/**
 * A title page's first page's parts where its section has none of its own: the latest earlier section's,
 * shown or not (one a section without a title page has, kept hidden, is still the one a later title page
 * takes on, as Word gives it), as a part of the titled section's own where that one is hidden: the later
 * title page showed none, or an older section's. `titled` says whether each section has a title page;
 * a reader of sections calls it once its parts are read, for each kind.
 */
export function titlePagesTakeHiddenFirsts(parts: OfficeContentNode[], kind: 'header' | 'footer', titled: readonly boolean[]): void {
    const meta = (n: OfficeContentNode) => n.metadata as { type?: string; section?: number; hidden?: boolean } | undefined;
    const bySection = new Map<number, OfficeContentNode>();
    for (const n of parts) if (n.type === kind && meta(n)?.type === 'first' && typeof meta(n)?.section === 'number' && !bySection.has(meta(n)!.section!)) bySection.set(meta(n)!.section!, n);
    let latest: OfficeContentNode | undefined;
    for (let section = 0; section < titled.length; section++) {
        const own = bySection.get(section);
        if (own) { latest = own; continue; }
        if (titled[section] && latest && meta(latest)?.hidden) {
            // (Its blocks shared with the part it copies: copied, a large part made a copy for each such section.)
            const copy: OfficeContentNode = { type: kind, text: latest.text, children: latest.children ?? [], metadata: { type: 'first', section } as HeaderFooterMetadata };
            parts.push(latest = copy);
        }
    }
}

/**
 * Which pages the running header or footer {@link runningFurniture} picked is on: every page
 * (`default`), the first alone (`first`: a letterhead, where the default one is empty) or the even pages
 * alone (`even`). A writer writes it as that kind, never as every page's.
 */
export function furnitureKind(nodes: OfficeContentNode[], kind: 'header' | 'footer'): 'default' | 'first' | 'even' {
    const type = nodes[0]?.type === kind ? (nodes[0].metadata as { type?: string } | undefined)?.type : undefined;
    return type === 'first' || type === 'even' ? type : 'default';
}

/**
 * The paper a writer lays its pages on, in points: the format's, turned by `landscape`; under 'auto' the
 * document's own first page (OfficeMetadata.pageSetup, turned by `landscape` when it is upright), else
 * A4. `name` is the format whose paper it is, upright, where it is one (for writers that name paper).
 * `margin` is the configured margins, a side left unset (`''`) being the document's own under 'auto',
 * else an inch; a page too small for them keeps half an inch of text room each way, its margins made
 * smaller to fit.
 */
export function resolvePaper(format: string | undefined, landscape: boolean | undefined, setup: PageSetup | undefined,
    margin?: { top?: number | string; right?: number | string; bottom?: number | string; left?: number | string }): { w: number; h: number; landscape: boolean; name?: string; margin: { top: number; right: number; bottom: number; left: number }; headerDistance?: number; footerDistance?: number } {
    const key = String(format ?? 'auto').toLowerCase();
    let w: number, h: number;
    // (A document's page between an inch and two hundred inches each way (see isPageSize); another is A4.)
    const own = key === 'auto' && hasPageSize(setup);
    if (own) {
        [w, h] = landscape && setup.width < setup.height ? [setup.height, setup.width] : [setup.width, setup.height];
    } else {
        const paper = PAPER_SIZES_PT[key] ?? PAPER_SIZES_PT.a4;
        [w, h] = landscape ? [paper.h, paper.w] : [paper.w, paper.h];
    }
    const upright = [Math.min(w, h), Math.max(w, h)];
    const name = Object.keys(PAPER_SIZES_PT).find(n => Math.abs(PAPER_SIZES_PT[n].w - upright[0]) < 1.5 && Math.abs(PAPER_SIZES_PT[n].h - upright[1]) < 1.5);
    const side = (value: number | string | undefined, document: number | undefined) =>
        (value === undefined || value === null || value === '') && own && document !== undefined && Number.isFinite(document) ? Math.max(0, document) : Math.max(0, marginPt(value === '' ? undefined : value));
    const m = { top: side(margin?.top, setup?.margin?.top), right: side(margin?.right, setup?.margin?.right), bottom: side(margin?.bottom, setup?.margin?.bottom), left: side(margin?.left, setup?.margin?.left) };
    // Half an inch of text room each way at least: margins wider than the page leaves are made smaller.
    const fit = (a: 'top' | 'left', b: 'bottom' | 'right', room: number) => {
        const total = m[a] + m[b], most = Math.max(0, room - 36);
        if (total > most) { const scale = total ? most / total : 0; m[a] = Math.round(m[a] * scale * 100) / 100; m[b] = Math.round(m[b] * scale * 100) / 100; }
    };
    fit('left', 'right', w);
    fit('top', 'bottom', h);
    // The header's and footer's distances from the edges, the document's under 'auto'.
    const distance = (value: number | undefined) => own && value !== undefined && Number.isFinite(value) && value > 0 ? value : undefined;
    const headerDistance = distance(setup?.headerDistance), footerDistance = distance(setup?.footerDistance);
    return { w, h, landscape: w > h, ...(name ? { name } : {}), margin: m, ...(headerDistance !== undefined ? { headerDistance } : {}), ...(footerDistance !== undefined ? { footerDistance } : {}) };
}

/** Portrait `{ w, h }` page size in points for a {@link PaperFormat} name (case-insensitive), defaulting to A4. */
export function paperSizePt(format: string | undefined): { w: number; h: number } {
    return PAPER_SIZES_PT[(format || 'a4').toLowerCase()] || PAPER_SIZES_PT.a4;
}

/**
 * Resolves a margin that may be a number (points) or a unit-labeled string (`'1in'`, `'2cm'`,
 * `'36pt'`, `'48px'`) to points. A bare number - or a bare numeric string - is points; a unit string
 * is converted via {@link lengthToPt}. Unresolvable input falls back to `fallback` (default 72pt).
 */
export function marginPt(v: number | string | undefined, fallback = 72): number {
    if (v == null) return fallback;
    if (typeof v === 'number') return Number.isFinite(v) ? v : fallback;
    const t = v.trim();
    // A bare number is points (`\d*\.?\d+`, the same numbers, tried every split of a long run of digits).
    if (/^-?(?:\d+(?:\.\d+)?|\.\d+)$/.test(t)) return parseFloat(t);
    const pt = lengthToPt(t);
    return pt == null ? fallback : pt;
}

/** Image MIME to file extension for a packaged media part. */
export const MIME_EXT: Record<string, string> = lookupTable({
    'image/jpeg': 'jpg', 'image/jpg': 'jpg', 'image/png': 'png', 'image/gif': 'gif',
    'image/svg+xml': 'svg', 'image/webp': 'webp', 'image/bmp': 'bmp', 'image/tiff': 'tiff',
});

/** Admonition type to accent color (matches the HTML generator's palette family). */
export const ADMONITION_COLOR: Record<string, string> = lookupTable({
    note: '0969DA', tip: '1A7F37', important: '8250DF', warning: '9A6700', caution: 'CF222E',
});

/**
 * A paragraph's or heading's spacing (ParagraphSpacing) fit to write: each value a finite number within
 * what a word processor takes (space and a fixed line height up to 1584pt, Word's 22 inches; a line
 * multiple up to 132), others dropped. Undefined when nothing is left.
 */
export function writableSpacing(meta: unknown): ParagraphSpacing | undefined {
    const spacing = (meta as { spacing?: unknown } | undefined)?.spacing as Record<string, unknown> | undefined;
    if (!spacing || typeof spacing !== 'object') return undefined;
    const within = (v: unknown, max: number, least: number) => typeof v === 'number' && Number.isFinite(v) && v >= least ? Math.min(v, max) : undefined;
    const out: ParagraphSpacing = {};
    // No space is a setting (a word processor's own default is some); a line distance of none is not.
    const before = within(spacing.before, 1584, 0), after = within(spacing.after, 1584, 0);
    const lineMultiple = within(spacing.lineMultiple, 132, 0.01), lineHeight = within(spacing.lineHeight, 1584, 0.01);
    if (before !== undefined) out.before = before;
    if (after !== undefined) out.after = after;
    if (lineMultiple !== undefined) out.lineMultiple = lineMultiple;
    else if (lineHeight !== undefined) { out.lineHeight = lineHeight; if (spacing.lineAtLeast === true) out.lineAtLeast = true; }
    return Object.keys(out).length ? out : undefined;
}

/** Parses a CSS-ish length ('12pt', '1in', '2cm', '10mm', '96px', or a bare number = px) to points. */
export function lengthToPt(value: string | number | undefined): number | null {
    if (value == null) return null;
    if (typeof value === 'number') return Number.isFinite(value) ? value * 0.75 : null; // px -> pt
    // Case-insensitive: CSS units are, and callers pass raw source values like "12PT"/"16PX".
    // The unit's trailing space is part of it: `\s*(unit)?\s*$` tried every split of a long run of spaces.
    const m = /^\s*(-?[\d.]+)\s*(?:(pt|in|cm|mm|px|%)\s*)?$/i.exec(value);
    if (!m) return null;
    const n = parseFloat(m[1]);
    if (!Number.isFinite(n)) return null;
    switch (m[2]?.toLowerCase()) {
        case 'pt': return n;
        case 'in': return n * 72;
        case 'cm': return n * 28.3465;
        case 'mm': return n * 2.83465;
        case '%': return null; // handled by the caller (fraction of content width)
        default: return n * 0.75; // px at 96dpi
    }
}

/** Sniffs intrinsic pixel size from the header of PNG / JPEG / GIF bytes. */
export function sniffImageSize(bytes: Uint8Array): { w: number; h: number } | null {
    if (bytes.length < 24) return null;
    // PNG: signature then IHDR width/height (big-endian) at offset 16.
    if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4E && bytes[3] === 0x47) {
        // Read UNSIGNED (`>>> 0`): a dimension with the top bit set (>= 2^31) would otherwise overflow
        // the signed `<< 24` into a negative value, fail the `> 0` check, return null, and let the
        // caller's megapixel cap be BYPASSED for exactly the absurd-dimension PNG the cap exists to stop.
        const w = ((bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19]) >>> 0;
        const h = ((bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23]) >>> 0;
        if (w > 0 && h > 0) return { w, h };
    }
    // GIF: logical screen width/height (little-endian) at offset 6.
    if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) {
        const w = bytes[6] | (bytes[7] << 8);
        const h = bytes[8] | (bytes[9] << 8);
        if (w > 0 && h > 0) return { w, h };
    }
    // BMP: width and height (signed, a negative height for rows stored top down) at offsets 18 and 22.
    if (bytes[0] === 0x42 && bytes[1] === 0x4D && bytes.length >= 26) {
        const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
        const w = view.getInt32(18, true), h = Math.abs(view.getInt32(22, true));
        if (w > 0 && h > 0) return { w, h };
    }
    // JPEG: scan segments for a SOF marker carrying height/width.
    if (bytes[0] === 0xFF && bytes[1] === 0xD8) {
        let o = 2;
        while (o + 9 < bytes.length) {
            if (bytes[o] !== 0xFF) { o++; continue; }
            const marker = bytes[o + 1];
            if (marker >= 0xC0 && marker <= 0xCF && marker !== 0xC4 && marker !== 0xC8 && marker !== 0xCC) {
                const h = (bytes[o + 5] << 8) | bytes[o + 6];
                const w = (bytes[o + 7] << 8) | bytes[o + 8];
                if (w > 0 && h > 0) return { w, h };
                return null;
            }
            const len = (bytes[o + 2] << 8) | bytes[o + 3];
            if (len <= 0) return null;
            o += 2 + len;
        }
    }
    return null;
}

/**
 * A colour as bare uppercase `RRGGBB`, or null: any CSS colour the AST holds (`#RGB`, `#RRGGBB` with or
 * without `#`, named, `rgb()`, `hsl()`; see cssColorHex). Read as hex alone, a named colour was dropped.
 */
export function hexColor(v: string | undefined): string | null {
    return cssColorHex(v);
}

/** Sanitizes a bookmark/anchor name to `[A-Za-z0-9_]`, leading letter/underscore, <=40 chars. */
export function toBookmarkNameRaw(name: string): string {
    let s = (name || '').replace(/[^A-Za-z0-9_]/g, '_');
    if (!/^[A-Za-z_]/.test(s)) s = '_' + s;
    return s.slice(0, 40) || '_';
}

/** Normalizes a Date or date-like string to a W3CDTF instant (`YYYY-MM-DDThh:mm:ssZ`), or null. */
export function toW3CDTF(v: unknown): string | null {
    let d: Date | null = null;
    if (v instanceof Date && !isNaN(v.getTime())) d = v;
    else if (typeof v === 'string' && v.trim() !== '') { const p = new Date(v); if (!isNaN(p.getTime())) d = p; }
    return d ? d.toISOString().replace(/\.\d+Z$/, 'Z') : null;
}

/**
 * Resolves the modification instant used for both a package's visible date property and every zip
 * entry's mtime, from a raw `effectiveMetadata.modified` value.
 *
 * Prefers a real Date, then a date-like string, and only falls back to the current time when neither
 * exists (the sole non-reproducible path, hence last). `iso` is `YYYY-MM-DDThh:mm:ssZ` (UTC, whole
 * seconds). `mtime` is additionally clamped to zip's DOS-timestamp range (1980-2099): fflate throws
 * rather than clamping, and an unset/epoch-zero date is the common out-of-range case.
 *
 * Known limitation: the one UTC wall-clock hour per year that falls in the running machine's DST
 * spring-forward gap has no local representation, so on a DST-observing machine the mtime for a
 * `modified` value in that hour shifts by 1h and its DOS timestamp differs from a non-DST machine's.
 * This cannot be corrected without reimplementing DOS-timestamp encoding (fflate reads local getters);
 * it affects only that hour and only the zip mtime, never `iso` or the visible metadata date.
 */
export function resolveZipInstant(raw: unknown): { iso: string; mtime: Date } {
    let resolved: Date | null = null;
    if (raw instanceof Date && !isNaN(raw.getTime())) resolved = raw;
    else if (typeof raw === 'string' && raw !== '') { const p = new Date(raw); if (!isNaN(p.getTime())) resolved = p; }
    resolved ??= new Date();
    // fflate stamps a zip entry's DOS timestamp from the mtime's LOCAL-time getters, so the same
    // instant would otherwise yield different archive bytes in different timezones. Build the mtime from
    // the instant's UTC calendar fields treated as LOCAL fields: its local getters then read back the
    // UTC wall clock identically on every machine, so the package is byte-for-byte reproducible
    // cross-timezone (the reproducibility contract callers rely on). The year is clamped into zip's DOS
    // range (1980-2099), since fflate throws rather than clamps and an epoch-zero date is a common case.
    const year = Math.min(2099, Math.max(1980, resolved.getUTCFullYear()));
    const mtime = new Date(year, resolved.getUTCMonth(), resolved.getUTCDate(), resolved.getUTCHours(), resolved.getUTCMinutes(), resolved.getUTCSeconds());
    return { iso: resolved.toISOString().replace(/\.\d+Z$/, 'Z'), mtime };
}

/** The id of a YouTube video `url` links to (a watch, embed, shorts or youtu.be link); '' for any other URL. */
function youtubeVideoIdOf(url: string): string {
    let parsed: URL;
    try {
        parsed = new URL(url);
    } catch {
        return '';
    }
    const host = parsed.hostname.replace(/^(?:www\.|m\.)/, '');
    const id = host === 'youtu.be' ? parsed.pathname.split('/')[1]
        : /^youtube(?:-nocookie)?\.com$/.test(host) ? (parsed.searchParams.get('v') ?? /^\/(?:embed|shorts)\/([^/]+)/.exec(parsed.pathname)?.[1]) : '';
    return id && /^[\w-]+$/.test(id) ? id : '';
}

/**
 * What an embed shows: a YouTube video (by its id), or a page in a frame (by its URL). An embed built
 * by hand may name no type, or a YouTube video by its URL alone, so it is read from what it carries: a
 * video id, else a YouTube URL, else any URL. Null when it carries none of them.
 */
export function resolveEmbed(meta: EmbedMetadata | undefined): { kind: 'youtube'; videoId: string } | { kind: 'iframe'; url: string } | null {
    if (!meta) return null;
    if (meta.embedType !== 'iframe') {
        // A video id is letters, digits, `_` and `-`, however it was given.
        const videoId = (typeof meta.videoId === 'string' && /^[\w-]+$/.test(meta.videoId) ? meta.videoId : '') || youtubeVideoIdOf(meta.url || '');
        if (videoId) return { kind: 'youtube', videoId };
    }
    return meta.url ? { kind: 'iframe', url: meta.url } : null;
}

/** The URL an embed shows: its own, else its YouTube video's watch page; '' when it has neither. */
export function embedUrl(meta: EmbedMetadata | undefined): string {
    if (meta?.url) return meta.url;
    const embed = resolveEmbed(meta);
    return embed?.kind === 'youtube' ? `https://www.youtube.com/watch?v=${embed.videoId}` : '';
}

/**
 * A spreadsheet as its printout shows it: each cell whose number the sheet shows through its number
 * format (CellMetadata.shownText) written as that text, `19%` rather than `0.19`. The AST as given when no
 * cell has one; otherwise a copy, sharing every node it does not change.
 */
export function withShownCellText(ast: OfficeParserAST): OfficeParserAST {
    const shown = (c: OfficeContentNode) => c.type === 'cell' && typeof (c.metadata as CellMetadata | undefined)?.shownText === 'string';
    if (!ast.content.some(n => n.type === 'sheet' && (n.children || []).some(r => (r.children || []).some(shown)))) return ast;
    return {
        ...ast,
        content: ast.content.map(n => n.type !== 'sheet' ? n : {
            ...n,
            children: (n.children || []).map(r => r.type !== 'row' || !(r.children || []).some(shown) ? r : {
                ...r,
                children: (r.children || []).map(c => {
                    if (!shown(c)) return c;
                    const text = (c.metadata as CellMetadata).shownText!;
                    // The cell's text, in its first run's formatting.
                    const run = (c.children || []).find(k => k.type === 'text');
                    return { ...c, text, children: [{ type: 'text', text, ...(run?.formatting ? { formatting: run.formatting } : {}) } as OfficeContentNode] } as OfficeContentNode;
                }),
            }),
        }),
    };
}

/**
 * A presentation as its printout shows it: the slides without their speaker notes (a slide's own
 * `notes`). The AST as given when no slide has any; otherwise a copy, sharing every other node.
 */
export function withoutSpeakerNotes(ast: OfficeParserAST): OfficeParserAST {
    if (!ast.content.some(n => n.type === 'slide' && n.notes?.length)) return ast;
    return { ...ast, content: ast.content.map(n => n.type === 'slide' && n.notes?.length ? { ...n, notes: undefined } : n) };
}

/**
 * A spreadsheet's content as a printout shows it: hidden sheets, rows and cells (SheetMetadata.hidden,
 * RowMetadata.hidden, CellMetadata.hidden) left out, and the rows and columns that remain numbered
 * without the gaps, a span across hidden columns shortened by them. The AST as given when nothing in
 * it is hidden; otherwise a copy, sharing the nodes it does not change.
 */
export function withoutHiddenSheetContent(ast: OfficeParserAST): OfficeParserAST {
    let changed = false;
    const isHiddenCell = (c: OfficeContentNode) => c.type === 'cell' && !!(c.metadata as CellMetadata | undefined)?.hidden;
    const content = ast.content.flatMap((node): OfficeContentNode[] => {
        if (node.type !== 'sheet') return [node];
        if ((node.metadata as SheetMetadata | undefined)?.hidden) { changed = true; return []; }
        const children = node.children || [];
        const isHiddenRow = (r: OfficeContentNode) => r.type === 'row' && !!(r.metadata as RowMetadata | undefined)?.hidden;
        if (!children.some(r => isHiddenRow(r) || (r.type === 'row' && (r.children || []).some(isHiddenCell)))) return [node];
        changed = true;
        // A column is hidden where a shown row's cell in it is; a row by its own mark.
        const hiddenCols = new Set<number>();
        const hiddenRows = new Set<number>();
        for (const r of children) {
            if (r.type !== 'row') continue;
            for (const c of r.children || []) {
                const m = c.metadata as CellMetadata | undefined;
                if (c.type !== 'cell' || !m?.hidden) continue;
                if (isHiddenRow(r)) { if (typeof m.row === 'number') hiddenRows.add(m.row); }
                else if (typeof m.col === 'number') hiddenCols.add(m.col);
            }
        }
        const cols = [...hiddenCols].sort((a, b) => a - b);
        const rowsHidden = [...hiddenRows].sort((a, b) => a - b);
        // How many of `sorted` are below `at` (a binary search).
        const below = (sorted: number[], at: number) => { let lo = 0, hi = sorted.length; while (lo < hi) { const mid = (lo + hi) >> 1; if (sorted[mid] < at) lo = mid + 1; else hi = mid; } return lo; };
        const kept: OfficeContentNode[] = [];
        for (const child of children) {
            if (child.type !== 'row') { kept.push(child); continue; }
            if (isHiddenRow(child)) continue;
            const cells = (child.children || []).filter(c => !isHiddenCell(c)).map(c => {
                const m = c.metadata as CellMetadata | undefined;
                if (c.type !== 'cell' || !m || (typeof m.col !== 'number' && typeof m.row !== 'number')) return c;
                const next: CellMetadata = { ...m };
                if (typeof m.col === 'number') {
                    next.col = m.col - below(cols, m.col);
                    if (m.colSpan && m.colSpan > 1) {
                        const span = m.colSpan - (below(cols, m.col + m.colSpan) - below(cols, m.col));
                        if (span > 1) next.colSpan = span; else delete next.colSpan;
                    }
                }
                if (typeof m.row === 'number') next.row = m.row - below(rowsHidden, m.row);
                return { ...c, metadata: next } as OfficeContentNode;
            });
            kept.push({ ...child, children: cells });
        }
        return [{ ...node, children: kept }];
    });
    return changed ? { ...ast, content } : ast;
}

const propText = (s: string): string => escapeXml(stripInvalidXmlChars(s));

/**
 * A document's custom properties as a package holds them: each a name (not blank, at most 255 characters,
 * as Office takes) and a value of its kind, a flag, a number, a date (between the years 1601 and 9999, as
 * a file time holds one), else text; an object or list as its JSON, not `[object Object]`.
 */
function customEntries(props: Record<string, unknown> | undefined): { name: string; kind: 'bool' | 'number' | 'date' | 'text'; value: string }[] {
    const out: { name: string; kind: 'bool' | 'number' | 'date' | 'text'; value: string }[] = [];
    // One property a name as Office compares them (trimmed, case aside): the first.
    const names = new Set<string>();
    for (const [key, value] of Object.entries(props || {})) {
        const name = key.trim().slice(0, 255);
        if (!name || value === undefined || names.has(name.toLowerCase())) continue;
        names.add(name.toLowerCase());
        const time = value instanceof Date ? value.getTime() : NaN;
        const year = Number.isFinite(time) ? (value as Date).getUTCFullYear() : NaN;
        if (typeof value === 'boolean') out.push({ name, kind: 'bool', value: value ? 'true' : 'false' });
        else if (typeof value === 'number' && Number.isFinite(value)) out.push({ name, kind: 'number', value: String(value) });
        else if (value instanceof Date && year >= 1601 && year <= 9999) out.push({ name, kind: 'date', value: value.toISOString().replace(/\.\d+Z$/, 'Z') });
        // A date no file time holds: its ISO text (the same on every machine), or 'Invalid Date'.
        else if (value instanceof Date) out.push({ name, kind: 'text', value: Number.isFinite(time) ? value.toISOString() : 'Invalid Date' });
        else {
            let text: string;
            if (value !== null && typeof value === 'object' && !(value instanceof Date)) { try { text = JSON.stringify(value) ?? String(value); } catch { text = String(value); } }
            else text = String(value);
            out.push({ name, kind: 'text', value: text });
        }
    }
    return out;
}

/**
 * A document's custom properties as an OOXML package's `docProps/custom.xml` (DOCX, XLSX, PPTX), each typed
 * as it is held (see customEntries); null where there are none.
 */
export function ooxmlCustomPropertiesXml(props: Record<string, unknown> | undefined): string | null {
    const entries = customEntries(props);
    if (!entries.length) return null;
    const tag = { bool: 'vt:bool', number: 'vt:r8', date: 'vt:filetime', text: 'vt:lpwstr' } as const;
    const xml = entries.map((e, i) => `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${i + 2}" name="${propText(e.name)}"><${tag[e.kind]}>${propText(e.value)}</${tag[e.kind]}></property>`).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
        + `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" `
        + `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${xml}</Properties>`;
}

/** A document's custom properties as ODF `meta:user-defined` elements (ODT, ODS, ODP), each typed as it is held (see customEntries). */
export function odfUserDefinedXml(props: Record<string, unknown> | undefined): string {
    const type = { bool: 'boolean', number: 'float', date: 'date', text: 'string' } as const;
    return customEntries(props).map(e => `<meta:user-defined meta:name="${propText(e.name)}" meta:value-type="${type[e.kind]}">${propText(e.value)}</meta:user-defined>`).join('');
}

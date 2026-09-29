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

import { EmbedMetadata, ImageMode, OfficeContentNode } from '../types.js';
import { cssColorHex } from './colorUtils.js';
import { lookupTable } from './lookupUtils.js';

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
    try { return encodeURI(url); } catch { return url; }
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
export const PAPER_FORMAT_NAMES: readonly string[] = Object.keys(PAPER_SIZES_PT);

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

/**
 * PDF Parser
 *
 * Extracts text, structure, metadata, images, links and attachments from PDF files using PDF.js
 * (pdfjs-dist).
 *
 * **Features**
 * - High-fidelity text assembly: baseline line clustering, gap-based word spacing (no more glued or
 *   broken words), super/subscript detection, hyphenation repair, and paragraph reconstruction.
 * - Reading-order recovery for multi-column and float-beside-text pages via a recursive XY-cut.
 * - Semantic structure from tagged PDFs (headings, tables, lists, footnotes) via the structure tree,
 *   with a geometric fallback for untagged files. (Tagged path: {@link module:parsers/pdf/structTree}.)
 * - Per-node page geometry (`bounds`) and page dimensions, on by default, opt out with
 *   `ignorePageGeometry`.
 * - Password-protected documents via the top-level `password`, or the `onPassword` callback to
 *   supply one lazily/interactively (the same config every encryptable format uses).
 * - Comprehensive metadata (including the document outline/bookmarks, page labels, and permissions),
 *   hyperlink extraction, image extraction with optional OCR, and embedded file attachments.
 *
 * **Pipeline**
 * 1. Open the document (one worker, one pass).
 * 2. Extract metadata and embedded attachments.
 * 3. For each selected page, collect normalized text runs, images, annotations and (when tagged) the
 *    structure tree in a single pass.
 * 4. Post-process purely in memory: resolve fonts once per document, build lines, segment blocks,
 *    reconstruct paragraphs/headings (or walk the structure tree), and interleave images.
 *
 * @module PdfParser
 * @see https://mozilla.github.io/pdf.js/ PDF.js documentation
 */

import { zlibSync } from 'fflate';
import { DEFAULT_OFFICE_PARSER_CONFIG } from '../defaults.js';
import { FullOfficeParserConfig, ImageMetadata, OfficeAttachment, OfficeContentNode, OfficeErrorType, OfficeMetadata, OfficeParserAST, OfficeWarningType, TextFormatting, TextMetadata } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { parseOfficeDate } from '../utils/dateUtils.js';
import { assertNode, isBrowser } from '../utils/envUtils.js';
import { checkAbortSignal, getOfficeError, logWarning } from '../utils/errorUtils.js';
import { crc32, createAttachment } from '../utils/imageUtils.js';
import { loadPdfJs } from '../utils/moduleLoader.js';
import { ocrDuringParse } from '../utils/ocrUtils.js';
import { collectColorMarks, ColorLookup, makeColorLookup } from './pdf/pdfColor.js';
import { computeRunBox, identityMatrix, mulMatrix, roundBounds, rotateBoundsToRendered, toMatrix6, unionAll } from './pdf/geometry.js';
import { PageExtract, PdfImage, PdfLayoutConfig, RawRun, ResolvedFont } from './pdf/pdfTypes.js';
import { blockToNodes, buildLines, computeDocContext, detectTables, DocContext, PageContext, recoverTaggedGrids, runsToParagraph, segmentIntoBlocks } from './pdf/textLayout.js';
import { buildTaggedNodes } from './pdf/structTree.js';
import { setOwn } from '../utils/lookupUtils.js';

/** Type guard for a pdf.js TextItem (marked-content items lack `str`/`transform`). */
function isTextItem(item: any): item is { str: string; transform: number[]; width: number; height: number; fontName: string; dir?: string; hasEOL?: boolean } {
    return item && typeof item.str === 'string' && Array.isArray(item.transform) && item.transform.length >= 6;
}

/** A link annotation matched to text by geometric overlap. */
interface ResolvedLink {
    /** Rect in layout (rotation-0) viewport space: [minX, minY, maxX, maxY]. */
    rect: [number, number, number, number];
    meta: TextMetadata;
}

/** A highlight annotation quad matched to text by geometric overlap; `color` is a hex string. */
interface ResolvedHighlight {
    /** Rect in layout (rotation-0) viewport space: [minX, minY, maxX, maxY]. */
    rect: [number, number, number, number];
    color: string;
}

/** An internal destination: the 0-based target page and its top y in PDF user space (y-up), if any. */
interface SectionTarget {
    pageIndex: number;
    /** Target top in PDF user space (y grows up), or null when the destination is whole-page (Fit). */
    pdfY: number | null;
}

/**
 * Collects internal-link destinations during the single collection pass so a later pass, once every
 * page's headings and their positions are known, can point each link at the nearest heading (falling
 * back to the page). `register` hands back a transient href that the post-pass rewrites in place, so
 * the placeholder never escapes into the returned AST.
 */
class SectionLinks {
    readonly targets: SectionTarget[] = [];
    register(target: SectionTarget): string {
        const k = this.targets.length;
        this.targets.push(target);
        return `#__pdfsec_${k}`;
    }
}

/**
 * Encodes raw RGBA pixel data into a PNG buffer (8-bit RGB, alpha flattened against white).
 *
 * PNG replaced an earlier uncompressed BMP encoder: a scanned page as BMP is many megabytes, and
 * inlining that as a base64 `data:` URI produced multi-megabyte single lines that broke downstream
 * consumers (e.g. Markdown renderers). Deflate typically shrinks a scanned page by an order of
 * magnitude, and PNG is a real image format that Tesseract and browsers both accept, so OCR and
 * embedding still work. Uses fflate's `zlibSync` (the same browser-safe dependency the ZIP reader
 * uses) rather than Node's `zlib`, so the browser bundle needs no polyfill.
 */
function encodePng(width: number, height: number, data: Uint8Array | Uint8ClampedArray): Buffer {
    // Raw image data: one filter byte (0 = none) per scanline, then RGB triples, top-to-bottom.
    const raw = new Uint8Array(height * (1 + width * 3));
    let o = 0;
    for (let y = 0; y < height; y++) {
        raw[o++] = 0; // filter: none
        for (let x = 0; x < width; x++) {
            const i = (y * width + x) * 4;
            const alpha = data[i + 3] / 255;
            raw[o++] = Math.round(data[i + 0] * alpha + 255 * (1 - alpha));
            raw[o++] = Math.round(data[i + 1] * alpha + 255 * (1 - alpha));
            raw[o++] = Math.round(data[i + 2] * alpha + 255 * (1 - alpha));
        }
    }
    const idatData = zlibSync(raw, { level: 6 });

    // Assemble chunks: each is length(4 BE) + type(4) + data + CRC32(4 BE) over type+data.
    const chunk = (type: string, body: Uint8Array): Buffer => {
        const typeBytes = Buffer.from(type, 'ascii');
        const out = Buffer.alloc(12 + body.length);
        out.writeUInt32BE(body.length, 0);
        typeBytes.copy(out, 4);
        Buffer.from(body).copy(out, 8);
        out.writeUInt32BE(crc32(out.subarray(4, 8 + body.length)), 8 + body.length);
        return out;
    };

    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(width, 0);
    ihdr.writeUInt32BE(height, 4);
    ihdr[8] = 8;   // bit depth
    ihdr[9] = 2;   // color type: 2 = truecolor RGB
    ihdr[10] = 0;  // compression: deflate
    ihdr[11] = 0;  // filter: adaptive
    ihdr[12] = 0;  // interlace: none

    const signature = Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
    return Buffer.concat([
        signature,
        chunk('IHDR', ihdr),
        chunk('IDAT', idatData),
        chunk('IEND', new Uint8Array(0)),
    ]);
}

/**
 * Converts raw PDF image pixel data (grayscale/RGB/RGBA) to a normalized RGBA buffer. The result is
 * re-encoded to PNG by {@link encodePng}, which is what enables OCR on PDF images.
 */
function convertToRgbaBuffer(data: Uint8Array | Uint8ClampedArray, width: number, height: number, kind?: number): Buffer {
    let rgbaData: Uint8ClampedArray;
    if (kind === 1) {
        // GRAYSCALE_1BPP: one BIT per pixel, most-significant bit first, each row padded to a whole
        // byte (stride = ceil(width/8)). The old code read one byte per pixel, so a CCITT/JBIG2 fax
        // scan came out as noise for the first 1/8 and black after. pdf.js has already applied any
        // Decode array, so a set bit is white (255) and a clear bit is black (0). (The RGBA allocation
        // is bounded because getDocument caps decodable images via `maxImageSize`.)
        const rowBytes = (width + 7) >> 3;
        rgbaData = new Uint8ClampedArray(width * height * 4);
        for (let y = 0; y < height; y++) {
            const rowStart = y * rowBytes;
            for (let x = 0; x < width; x++) {
                const byte = data[rowStart + (x >> 3)] ?? 0;
                const gray = ((byte >> (7 - (x & 7))) & 1) ? 255 : 0;
                const o = (y * width + x) * 4;
                rgbaData[o] = gray; rgbaData[o + 1] = gray; rgbaData[o + 2] = gray; rgbaData[o + 3] = 255;
            }
        }
    } else if (kind === 2 || data.length === width * height * 3) {
        rgbaData = new Uint8ClampedArray(width * height * 4);
        for (let i = 0; i < width * height; i++) {
            rgbaData[i * 4] = data[i * 3];
            rgbaData[i * 4 + 1] = data[i * 3 + 1];
            rgbaData[i * 4 + 2] = data[i * 3 + 2];
            rgbaData[i * 4 + 3] = 255;
        }
    } else {
        rgbaData = data instanceof Uint8ClampedArray ? data : new Uint8ClampedArray(data);
    }
    return Buffer.from(rgbaData.buffer, rgbaData.byteOffset, rgbaData.byteLength);
}

/** Reads and coerces the PDF-specific knobs from a resolved parser config. */
function resolvePdfLayoutConfig(config: FullOfficeParserConfig): PdfLayoutConfig {
    const p = config.pdfParserConfig ?? {};
    const num = (v: unknown, fallback: number) => {
        const n = typeof v === 'number' ? v : Number(v);
        return Number.isFinite(n) ? n : fallback;
    };
    return {
        useTags: p.useTags !== false,
        detectColumns: p.detectColumns !== false,
        mergeHyphenatedWords: p.mergeHyphenatedWords !== false,
        lineToleranceFactor: num(p.lineToleranceFactor, 0.35),
        spaceToleranceFactor: num(p.spaceToleranceFactor, 0.25),
        headingDetection: p.headingDetection ?? 'auto',
        normalizeText: p.normalizeText !== false,
        extractTextColor: p.extractTextColor !== false, // default true, matching defaults.ts and its siblings
        includeBounds: !config.ignorePageGeometry,
    };
}

/** Parses a 1-based page-range spec like "1-3,7" against a page count. Empty/invalid means all. */
function parsePageRange(spec: string | undefined, numPages: number): number[] {
    const all = () => Array.from({ length: numPages }, (_, i) => i + 1);
    if (!spec || !spec.trim()) return all();
    const pages = new Set<number>();
    for (const part of spec.split(',')) {
        const token = part.trim();
        if (!token) continue;
        const range = token.match(/^(\d+)\s*-\s*(\d+)$/);
        if (range) {
            const lo = parseInt(range[1], 10), hi = parseInt(range[2], 10);
            if (lo < 1 || hi < lo) return all();
            // Clamp to the document before iterating, not inside the loop: "1-999999999999" would
            // otherwise spin through a trillion integers and only then discard all but the first few.
            const end = Math.min(hi, numPages);
            for (let p = lo; p <= end; p++) pages.add(p);
        } else if (/^\d+$/.test(token)) {
            const p = parseInt(token, 10);
            if (p >= 1 && p <= numPages) pages.add(p);
        } else {
            return all();
        }
    }
    return pages.size ? [...pages].sort((a, b) => a - b) : all();
}

/** Resolves a font id once per document into name/weight/style plus ascent/descent for geometry. */
async function resolveFont(fontKey: string, commonObjs: any, styles: Record<string, any>): Promise<ResolvedFont> {
    const style = styles[fontKey] || {};
    const resolved: ResolvedFont = {
        bold: false,
        italic: false,
        ascent: typeof style.ascent === 'number' ? style.ascent : 0.8,
        descent: typeof style.descent === 'number' ? style.descent : -0.2,
    };
    try {
        if (commonObjs?.has?.(fontKey)) {
            const fontData: any = await new Promise((resolve) => commonObjs.get(fontKey, (d: any) => resolve(d)));
            const rawName: string | undefined = typeof fontData?.name === 'string' ? fontData.name : undefined;
            if (rawName) {
                resolved.name = rawName.replace(/^[A-Z]{6}\+/, '');
                const lower = rawName.toLowerCase();
                resolved.bold = lower.includes('bold') || (typeof fontData.black === 'boolean' && fontData.black);
                resolved.italic = lower.includes('italic') || lower.includes('oblique');
            }
        }
    } catch {
        // Font lookup failed; ascent/descent from styles are enough to proceed.
    }
    if (typeof style.fontFamily === 'string' && !resolved.name) resolved.name = style.fontFamily;
    return resolved;
}

/** Applies a 6-element affine matrix to a point. */
function applyMatrix(m: number[], x: number, y: number): [number, number] {
    return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}

/** Converts a PDF-space rect [x1,y1,x2,y2] to a normalized layout-viewport rect. */
function toViewportRect(viewport: any, rect: number[]): [number, number, number, number] {
    const m = viewport.transform;
    const [ax, ay] = applyMatrix(m, rect[0], rect[1]);
    const [bx, by] = applyMatrix(m, rect[2], rect[3]);
    return [Math.min(ax, bx), Math.min(ay, by), Math.max(ax, bx), Math.max(ay, by)];
}

/** True for a color at or extremely close to black (the default text fill), so it is not reported. */
function isNearBlack(hex: string): boolean {
    const h = hex.replace(/^#/, '');
    // Colours arrive as full `#rrggbb`; anything else is not a near-black colour we recognise (the old
    // `hex === '#000000'` fallback here could never be true, since that input is 6 digits long).
    if (h.length !== 6) return false;
    return parseInt(h.slice(0, 2), 16) <= 8 && parseInt(h.slice(2, 4), 16) <= 8 && parseInt(h.slice(4, 6), 16) <= 8;
}

/** Converts an [r,g,b] byte triple (pdf.js annotation color) to a lowercase `#rrggbb` hex string. */
function rgbArrayToHex(c: ArrayLike<number> | null | undefined): string | undefined {
    if (!c || c.length < 3) return undefined;
    const h = (n: number) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0');
    return `#${h(c[0])}${h(c[1])}${h(c[2])}`;
}

/** Maps a highlight annotation's quadPoints (or its rect) to per-quad viewport rects. */
function highlightRects(viewport: any, annot: any): [number, number, number, number][] {
    const qp: ArrayLike<number> | undefined = annot.quadPoints;
    if (qp && qp.length >= 8) {
        const rects: [number, number, number, number][] = [];
        for (let i = 0; i + 8 <= qp.length; i += 8) {
            const xs = [qp[i], qp[i + 2], qp[i + 4], qp[i + 6]];
            const ys = [qp[i + 1], qp[i + 3], qp[i + 5], qp[i + 7]];
            rects.push(toViewportRect(viewport, [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]));
        }
        return rects;
    }
    return annot.rect ? [toViewportRect(viewport, annot.rect)] : [];
}

/** The most distinct links, and highlight areas, read from one page (see resolveAnnotations). */
const MAX_PAGE_ANNOTATIONS = 1000;

/**
 * Resolves a page's Link and Highlight annotations. Links become geometry + hyperlink metadata
 * (honoring config flags); highlights become per-quad viewport rects carrying their color, so a run
 * they cover gets a `backgroundColor`. Both come from the single `getAnnotations()` call.
 */
async function resolveAnnotations(
    page: any, viewport: any, pdfDocument: any, config: FullOfficeParserConfig,
    destCache: Map<string, SectionTarget | null>, sectionLinks: SectionLinks,
): Promise<{ links: ResolvedLink[]; highlights: ResolvedHighlight[] }> {
    const links: ResolvedLink[] = [];
    const highlights: ResolvedHighlight[] = [];
    let annots: any[];
    try {
        annots = await page.getAnnotations();
    } catch (e) {
        logWarning(OfficeWarningType.ANNOTATION_EXTRACTION_FAILED, config, page.pageNumber, e);
        return { links, highlights };
    }
    // Each distinct highlight quad and link once, and at most MAX_PAGE_ANNOTATIONS of each: every run on
    // the page is matched against them, so a page listing one annotation many times, or a highlight of
    // a huge QuadPoints array, took time in the square of the page.
    const seenHighlights = new Set<string>(), seenLinks = new Set<string>();
    let capped = false;
    for (const annot of annots) {
        if (annot.subtype === 'Highlight') {
            // A highlight with no /C renders yellow (pdf.js synthesizes that appearance), so mirror it.
            const color = rgbArrayToHex(annot.color) ?? '#ffff00';
            for (const rect of highlightRects(viewport, annot)) {
                const key = `${rect.join(',')}|${color}`;
                if (seenHighlights.has(key)) continue;
                if (highlights.length >= MAX_PAGE_ANNOTATIONS) { capped = true; break; }
                seenHighlights.add(key);
                highlights.push({ rect, color });
            }
            continue;
        }
        if (annot.subtype !== 'Link' || !annot.rect) continue;
        const url: string | undefined = annot.url || annot.unsafeUrl || annot.data?.url;
        let meta: TextMetadata | undefined;
        if (url) {
            const internal = url.startsWith('#');
            if (internal && config.ignoreInternalLinks) continue;
            meta = { link: url, linkType: internal ? 'internal' : 'external' };
        } else if (annot.dest) {
            if (config.ignoreInternalLinks) continue;
            const target = await resolveDestFull(annot.dest, pdfDocument, destCache);
            meta = { link: target ? sectionLinks.register(target) : '#internal', linkType: 'internal' };
        }
        if (meta) {
            const rect = toViewportRect(viewport, annot.rect);
            const key = `${rect.join(',')}|${meta.link}`;
            if (seenLinks.has(key)) continue;
            if (links.length >= MAX_PAGE_ANNOTATIONS) { capped = true; continue; }
            seenLinks.add(key);
            links.push({ rect, meta });
        }
    }
    if (capped) logWarning(OfficeWarningType.ANNOTATION_EXTRACTION_FAILED, config, page.pageNumber, `the page has more than ${MAX_PAGE_ANNOTATIONS} distinct links or highlight areas; the rest were not read`);
    return { links, highlights };
}

/** Finds the highlight color covering a run: its vertical midpoint inside a quad with horizontal overlap. */
function highlightForBox(x: number, yTop: number, w: number, h: number, highlights: ResolvedHighlight[]): string | undefined {
    const midY = yTop + h / 2, minX = x, maxX = x + w;
    for (const hl of highlights) {
        const [lx1, ly1, lx2, ly2] = hl.rect;
        if (midY >= ly1 && midY <= ly2 && minX < lx2 && maxX > lx1) return hl.color;
    }
    return undefined;
}

/** Extracts a destination's top y (PDF user space) from its explicit array, per its fit type; null if none. */
function destTopY(explicit: any[]): number | null {
    const fit = explicit[1];
    const name: string | undefined = typeof fit === 'string' ? fit : fit?.name;
    const num = (v: unknown) => (typeof v === 'number' && Number.isFinite(v) ? v : null);
    // [pageRef, /XYZ, left, top, zoom] | [.., /FitH|/FitBH, top] | [.., /FitR, left, bottom, right, top]
    if (name === 'XYZ') return num(explicit[3]);
    if (name === 'FitH' || name === 'FitBH') return num(explicit[2]);
    if (name === 'FitR') return num(explicit[5]);
    return null; // Fit / FitB / FitV / FitBV: whole-page, no meaningful top
}

/**
 * Upper bound on the named-destination cache. It exists only to bound memory on a pathological file;
 * a real document (a hyperref thesis cross-references thousands of labels) must stay well under it,
 * because a name past the cap resolves to a dead `#internal` link. Misses are cached too, so a
 * repeated bad name costs one lookup, and this single bound covers both hits and misses.
 */
const MAX_DEST_CACHE = 50000;

/**
 * Resolves an internal destination to its target page (0-based) and top y, or null when it cannot be
 * resolved cheaply. Named destinations are resolved via `getDestination` and cached.
 */
async function resolveDestFull(
    dest: string | unknown[], pdfDocument: any, cache: Map<string, SectionTarget | null>,
): Promise<SectionTarget | null> {
    try {
        let explicit = dest;
        if (typeof dest === 'string') {
            if (cache.has(dest)) return cache.get(dest) ?? null;
            if (cache.size >= MAX_DEST_CACHE) return null;
            explicit = await pdfDocument.getDestination(dest);
            if (!explicit) { cache.set(dest, null); return null; }
        }
        if (Array.isArray(explicit) && explicit[0]) {
            const pageIndex = await pdfDocument.getPageIndex(explicit[0]);
            const target: SectionTarget = { pageIndex, pdfY: destTopY(explicit) };
            if (typeof dest === 'string') cache.set(dest, target);
            return target;
        }
    } catch {
        // fall through
    }
    if (typeof dest === 'string') cache.set(dest, null);
    return null;
}

/**
 * Splits an item's character range into segments by the annotation rects that overlap it, assigning
 * each character the link whose rect contains its (proportionally-estimated) centre. Adjacent
 * same-link characters coalesce, so a link covers only the text it spans rather than the whole run.
 */
function segmentByLinks(x: number, width: number, yTop: number, height: number, len: number, links: ResolvedLink[]): { start: number; end: number; link?: TextMetadata }[] {
    const charLink: (TextMetadata | undefined)[] = new Array(len).fill(undefined);
    // The characters whose centre a link's rect holds are a range, computed rather than tested one by
    // one; each character is assigned once (the first link holding it), skipping over assigned ones.
    const nextFree = new Int32Array(len + 1).map((_, i) => i);
    const free = (i: number): number => { let r = i; while (nextFree[r] !== r) r = nextFree[r]; while (nextFree[i] !== r) { const n = nextFree[i]; nextFree[i] = r; i = n; } return r; };
    for (const l of links) {
        const [lx1, ly1, lx2, ly2] = l.rect;
        if (!(yTop < ly2 && yTop + height > ly1)) continue; // require vertical overlap
        if (width === 0) {
            // No width to place characters by: the centre of every character is `x`.
            if (x >= lx1 && x <= lx2) for (let i = free(0); i < len; i = free(i + 1)) { charLink[i] = l.meta; nextFree[i] = i + 1; }
            continue;
        }
        if (!Number.isFinite(width)) continue;
        // The i whose centre x + (i + 0.5) / len * width lies in [lx1, lx2] (width may be negative).
        const a = ((lx1 - x) / width) * len - 0.5, b = ((lx2 - x) / width) * len - 0.5;
        const from = Math.max(0, Math.ceil(Math.min(a, b)));
        const to = Math.min(len - 1, Math.floor(Math.max(a, b)));
        for (let i = free(from); i <= to; i = free(i + 1)) { charLink[i] = l.meta; nextFree[i] = i + 1; }
    }
    const segs: { start: number; end: number; link?: TextMetadata }[] = [];
    let s = 0;
    for (let i = 1; i <= len; i++) {
        if (i === len || charLink[i] !== charLink[s]) { segs.push({ start: s, end: i, link: charLink[s] }); s = i; }
    }
    return segs;
}

/** Finds the hyperlink metadata for a run by intersecting its box with the page's link rects. */
function linkForBox(x: number, yTop: number, w: number, h: number, links: ResolvedLink[]): TextMetadata | undefined {
    const minX = x, maxX = x + w, minY = yTop, maxY = yTop + h;
    for (const l of links) {
        const [lx1, ly1, lx2, ly2] = l.rect;
        if (minX < lx2 && maxX > lx1 && minY < ly2 && maxY > ly1) return l.meta;
    }
    return undefined;
}

/** Warns when a document's extracted text is dominated by unmappable glyphs (bad/missing ToUnicode). */
function warnIfEncodingSuspect(runs: RawRun[], config: FullOfficeParserConfig): void {
    let total = 0, bad = 0;
    for (const r of runs) {
        for (const ch of r.text) {
            const c = ch.codePointAt(0)!;
            if (c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d) continue; // skip whitespace
            total++;
            // Private Use Area, replacement char, or a C0/C1 control character.
            if ((c >= 0xE000 && c <= 0xF8FF) || c === 0xFFFD || c < 0x20 || (c >= 0x7F && c <= 0x9F)) bad++;
        }
    }
    if (total >= 50 && bad / total >= 0.2) {
        logWarning(OfficeWarningType.PDF_TEXT_ENCODING_SUSPECT, config, `${Math.round((bad / total) * 100)}% of characters were unmappable`);
    }
}

/** Maps the granted-permission flags from `getPermissions()` to readable action names. */
function permissionNames(pdfjs: any, perms: number[]): string[] {
    const F = pdfjs.PermissionFlag || {};
    const pairs: [number | undefined, string][] = [
        [F.PRINT, 'print'], [F.MODIFY_CONTENTS, 'modify'], [F.COPY, 'copy'],
        [F.MODIFY_ANNOTATIONS, 'annotate'], [F.FILL_INTERACTIVE_FORMS, 'fillForms'],
        [F.COPY_FOR_ACCESSIBILITY, 'copyForAccessibility'], [F.ASSEMBLE, 'assemble'],
        [F.PRINT_HIGH_QUALITY, 'printHighQuality'],
    ];
    const set = new Set(perms);
    return pairs.filter(([flag]) => flag !== undefined && set.has(flag)).map(([, name]) => name);
}

/** Lists optional-content (layer) names and default visibility, defensively across pdf.js shapes. */
function listOptionalContentLayers(oc: any): { name: string; visible: boolean }[] {
    const out: { name: string; visible: boolean }[] = [];
    if (!oc || typeof oc.getGroups !== 'function') return out;
    let groups: any;
    try { groups = oc.getGroups(); } catch { return out; }
    if (!groups) return out;
    const entries: [string, any][] = groups instanceof Map
        ? [...groups.entries()]
        : Object.keys(groups).map(k => [k, groups[k]] as [string, any]);
    for (const [id, g] of entries) {
        const name = (g && (g.name || (g.data && g.data.name))) || id;
        let visible = true;
        try { if (typeof oc.isVisible === 'function') visible = !!oc.isVisible(id); } catch { /* keep default */ }
        out.push({ name: String(name), visible });
    }
    return out;
}

/** Nesting depth the outline walk will follow; deeper bookmarks are dropped, not recursed into. */
const MAX_OUTLINE_DEPTH = 64;
/** Total bookmarks converted; the rest are dropped. Bounds the per-item destination round-trips. */
const MAX_OUTLINE_ITEMS = 10000;

/**
 * Builds the document outline (bookmarks) as a tree of `list` nodes carrying destination links.
 *
 * Walked iteratively and under both a depth and a total-item cap. pdf.js builds the outline
 * iteratively with a RefSet, so a hostile `/First` chain 100k levels deep is handed back intact: a
 * recursive conversion would blow the stack (`RangeError`) and take the whole parse down with it,
 * and a flat outline of a million items would cost a million worker round-trips for its
 * destinations. A truncated outline is reported as a warning; the rest of the document is unaffected.
 */
async function buildOutline(
    pdfDocument: any, destCache: Map<string, SectionTarget | null>, sectionLinks: SectionLinks, config: FullOfficeParserConfig,
): Promise<OfficeContentNode[] | undefined> {
    let outline: any[] | null;
    try { outline = await pdfDocument.getOutline(); } catch { return undefined; }
    if (!Array.isArray(outline) || !outline.length) return undefined;

    const roots: OfficeContentNode[] = [];
    // Explicit DFS stack: each frame is one sibling list plus the array its nodes are appended to.
    const stack: { items: any[]; i: number; n: number; depth: number; out: OfficeContentNode[] }[] = [{ items: outline, i: 0, n: 0, depth: 0, out: roots }];
    let count = 0;
    let truncated = '';
    while (stack.length) {
        const frame = stack[stack.length - 1];
        if (frame.i >= frame.items.length) { stack.pop(); continue; }
        const item = frame.items[frame.i++];
        if (count >= MAX_OUTLINE_ITEMS) { truncated = `more than ${MAX_OUTLINE_ITEMS} bookmarks`; break; }
        count++;

        const title = typeof item?.title === 'string' ? item.title : '';
        let link: string | undefined;
        let linkType: 'internal' | 'external' | undefined;
        if (item?.url) { link = item.url; linkType = 'external'; }
        else if (item?.dest != null) {
            const target = await resolveDestFull(item.dest, pdfDocument, destCache);
            link = target ? sectionLinks.register(target) : '#internal';
            linkType = 'internal';
        }
        const label: OfficeContentNode = { type: 'text', text: title };
        if (link) label.metadata = { link, linkType };
        const node: OfficeContentNode = {
            type: 'list',
            text: title,
            children: [label],
            metadata: { listType: 'unordered', indentation: frame.depth, alignment: 'left', listId: 'pdf-outline', itemIndex: frame.n++ },
        };
        frame.out.push(node);

        const kids: any[] = Array.isArray(item?.items) ? item.items : [];
        if (!kids.length) continue;
        // Children are appended straight onto this node, after its label, as the child frame runs.
        if (frame.depth + 1 >= MAX_OUTLINE_DEPTH) truncated = truncated || `nesting deeper than ${MAX_OUTLINE_DEPTH} levels`;
        else stack.push({ items: kids, i: 0, n: 0, depth: frame.depth + 1, out: node.children! });
    }
    if (truncated) {
        logWarning(OfficeWarningType.PDF_OUTLINE_TRUNCATED, config, truncated);
    }
    return roots.length ? roots : undefined;
}

/** Lowercases and hyphenates text into a URL-fragment-safe anchor slug (Unicode letters/digits kept). */
function slugify(text: string): string {
    return text.toLowerCase().trim().replace(/[^\p{L}\p{N}]+/gu, '-').replace(/^-+|-+$/g, '').slice(0, 64);
}

/**
 * Points each collected internal link at the nearest heading on its target page, so a PDF
 * cross-reference or bookmark jumps to the actual section rather than the top of the page. The
 * referenced heading is given a stable anchor id (via its `anchorIds`), and every transient
 * `#__pdfsec_k` placeholder emitted during collection is rewritten to that anchor - or to `#page=N`
 * when no heading matches (whole-page destinations, rotated pages, or a target far from any heading).
 * Only headings actually referenced are given ids, so unreferenced ones are left untouched.
 */
function resolveSectionLinks(
    content: OfficeContentNode[], extracts: PageExtract[], sectionLinks: SectionLinks, rewriteRoots: OfficeContentNode[][],
): void {
    if (!sectionLinks.targets.length) return;

    const pageInfo = new Map<number, { authoredH: number; authoredY1: number; rotation: number }>();
    for (const e of extracts) pageInfo.set(e.pageNumber, { authoredH: e.authoredH, authoredY1: e.authoredY1, rotation: e.rotation });

    // Headings per page (top-to-bottom, with their rendered top y when known), plus a document-wide
    // count of each heading's base slug so a referenced heading can be disambiguated from same-text
    // headings elsewhere in the document.
    const headingsByPage = new Map<number, { node: OfficeContentNode; y: number | undefined }[]>();
    const slugCount = new Map<string, number>();
    for (const pageNode of content) {
        if (pageNode.type !== 'page' || typeof pageNode.metadata?.pageNumber !== 'number') continue;
        const hs: { node: OfficeContentNode; y: number | undefined }[] = [];
        const walk = (n: OfficeContentNode) => {
            if (n.type === 'heading') {
                hs.push({ node: n, y: n.bounds?.y });
                const s = slugify(n.text || '') || 'section';
                slugCount.set(s, (slugCount.get(s) || 0) + 1);
            }
            for (const c of n.children || []) walk(c);
        };
        for (const c of pageNode.children || []) walk(c);
        hs.sort((a, b) => (a.y ?? 0) - (b.y ?? 0));
        headingsByPage.set(pageNode.metadata.pageNumber, hs);
    }

    // Assign a unique, stable anchor id to a heading the first time a link targets it.
    const usedIds = new Set<string>();
    const idFor = new Map<OfficeContentNode, string>();
    const anchorFor = (h: OfficeContentNode): string => {
        const existing = idFor.get(h);
        if (existing) return existing;
        const base = slugify(h.text || '') || 'section';
        let id = base;
        // The generator gives every same-text heading the bare slug, so when this base is shared (or
        // already taken) use a suffixed id the generator never emits, so the link lands on THIS one.
        if ((slugCount.get(base) || 0) > 1 || usedIds.has(base)) {
            let i = 2;
            id = `${base}-${i}`;
            while (usedIds.has(id) || (slugCount.get(id) || 0) > 0) id = `${base}-${++i}`;
        }
        usedIds.add(id);
        idFor.set(h, id);
        const meta = (h.metadata ?? (h.metadata = { level: 1 })) as { anchorIds?: string[] };
        meta.anchorIds = [id, ...(meta.anchorIds || [])];
        return id;
    };

    const resolved = sectionLinks.targets.map((t): string => {
        const pnum = t.pageIndex + 1;
        const info = pageInfo.get(pnum);
        const hs = headingsByPage.get(pnum);
        // Match to the nearest heading only when geometry is present. Under `ignorePageGeometry` every
        // heading y is undefined, so fall back to the page anchor rather than binding every link to
        // the first heading (which a y=0 tie would otherwise do).
        if (t.pdfY != null && info && info.rotation === 0 && hs) {
            const positioned = hs.filter((h): h is { node: OfficeContentNode; y: number } => h.y !== undefined);
            if (positioned.length) {
                const targetY = info.authoredY1 - t.pdfY; // PDF user space (y-up) -> viewport (y-down), honouring a non-zero CropBox origin
                let best: OfficeContentNode | null = null, bestD = Infinity;
                for (const h of positioned) { const d = Math.abs(h.y - targetY); if (d < bestD) { bestD = d; best = h.node; } }
                if (best && bestD <= info.authoredH * 0.5) return `#${anchorFor(best)}`;
            }
        }
        return `#page=${pnum}`;
    });

    // Iterative walk over children *and* the side branches (`notes`, `comments`): a cross-reference
    // inside a footnote body hangs off `notes`, never off `children`, so a recursive children-only
    // rewrite would ship the raw `#__pdfsec_k` placeholder. Iterative so a pathologically deep
    // outline (see `buildOutline`) cannot blow the stack here either.
    const rewrite = (root: OfficeContentNode) => {
        const stack: OfficeContentNode[] = [root];
        while (stack.length) {
            const n = stack.pop()!;
            const meta = n.metadata as { link?: string } | undefined;
            if (meta && typeof meta.link === 'string') {
                const m = /^#__pdfsec_(\d+)$/.exec(meta.link);
                if (m) meta.link = resolved[Number(m[1])] ?? '#internal';
            }
            for (const c of n.children || []) stack.push(c);
            for (const c of n.notes || []) stack.push(c);
            for (const c of n.comments || []) stack.push(c);
        }
    };
    for (const roots of rewriteRoots) for (const n of roots) rewrite(n);
}

/** Collects one page's text runs, images and annotations into a PageExtract. */
async function collectPage(
    pdfjs: any, pdfDocument: any, pageNumber: number, config: FullOfficeParserConfig,
    pdfCfg: PdfLayoutConfig, fontCache: Map<string, ResolvedFont>, destCache: Map<string, SectionTarget | null>,
    sectionLinks: SectionLinks,
): Promise<PageExtract> {
    const page = await pdfDocument.getPage(pageNumber);
    const rotation = ((page.rotate % 360) + 360) % 360;
    const layoutViewport = page.getViewport({ scale: 1, rotation: 0 });
    const authoredW = layoutViewport.width;
    const authoredH = layoutViewport.height;
    // CropBox top in user space (viewBox = [x0, y0, x1, y1]); equals authoredH only when the CropBox
    // origin y is 0. A section-link destination's user-space y maps to viewport y as authoredY1 - y.
    const authoredY1 = Array.isArray((layoutViewport as any).viewBox) ? (layoutViewport as any).viewBox[3] : authoredH;
    const width = rotation % 180 === 0 ? authoredW : authoredH;
    const height = rotation % 180 === 0 ? authoredH : authoredW;

    const textContent = await page.getTextContent({ includeMarkedContent: true, disableNormalization: !pdfCfg.normalizeText });
    const styles: Record<string, any> = textContent.styles || {};

    // Resolve every font on the page once (document-scoped cache). Real font objects (with their
    // names, from which bold/italic are derived) only populate `page.commonObjs` during
    // `getOperatorList()`, never during `getTextContent()`, so fetch the operator list first when a
    // page introduces a not-yet-resolved font. Fonts are document-scoped, so in practice only the
    // first page or two pay this; the list is reused for image extraction below.
    const seen = new Set<string>();
    for (const item of textContent.items) if (isTextItem(item) && item.fontName) seen.add(item.fontName);
    const needFonts = [...seen].some(k => !fontCache.has(k));
    let ops: any = null;
    // Not `config.ocr`: page-image OCR runs over collected images, which require extractAttachments, so
    // ocr alone needs no operator list (and must not fetch one just to discard it).
    if (needFonts || config.extractAttachments || pdfCfg.extractTextColor) {
        try { ops = await page.getOperatorList(); } catch { ops = null; }
    }
    for (const key of seen) if (!fontCache.has(key)) fontCache.set(key, await resolveFont(key, page.commonObjs, styles));

    // Per-run fill color (opt-in): recovered from the operator list, which is now fetched above.
    const colorFor: ColorLookup | null = (pdfCfg.extractTextColor && ops)
        ? makeColorLookup(collectColorMarks(ops, layoutViewport.transform, pdfjs.OPS))
        : null;

    const { links, highlights } = await resolveAnnotations(page, layoutViewport, pdfDocument, config, destCache, sectionLinks);

    // Walk items, tracking the marked-content stack for mcid / Artifact scope.
    // Each entry carries the nearest mcid and whether an Artifact encloses it, set when pushed, so a run
    // reads them from the top entry: looked up through the whole stack for each run, content nested
    // thousands deep took time in the square of the page.
    const stack: { mcid: string | null; artifact: boolean }[] = [];
    const runs: RawRun[] = [];
    for (const item of textContent.items) {
        if (!isTextItem(item)) {
            const type = (item as any).type as string | undefined;
            if (type === 'beginMarkedContent' || type === 'beginMarkedContentProps') {
                const parent = stack[stack.length - 1];
                const id: string | null = (item as any).id ?? null;
                stack.push({ mcid: id ?? parent?.mcid ?? null, artifact: (item as any).tag === 'Artifact' || !!parent?.artifact });
            } else if (type === 'endMarkedContent') {
                stack.pop();
            }
            continue;
        }
        if (!item.str) continue;
        const font = fontCache.get(item.fontName) || { bold: false, italic: false, ascent: 0.8, descent: -0.2 };
        const m = pdfjs.Util.transform(layoutViewport.transform, item.transform);
        const box = computeRunBox(m, item.width || 0, font.ascent, font.descent);

        const formatting: TextFormatting = {};
        if (font.name) formatting.font = font.name;
        if (font.bold) formatting.bold = true;
        if (font.italic) formatting.italic = true;
        // Font size carries its unit, per the AST contract (TextFormatting.size); every other parser
        // appends 'pt', and generators that read it as a length (lengthToPt, CSS font-size) otherwise
        // misread a bare number as pixels and shrink the text to 75%.
        formatting.size = `${Math.round(box.fontSize * 2) / 2}pt`;
        // Fill color (opt-in): skip near-black, the default, so only real colors are reported.
        if (colorFor) {
            const color = colorFor(box.x, box.yBaseline, box.fontSize, box.width);
            if (color && !isNearBlack(color)) formatting.color = color;
        }
        // Highlight background (always on; from the annotation list already read above).
        if (highlights.length) {
            const bg = highlightForBox(box.x, box.yTop, box.width, box.height, highlights);
            if (bg) formatting.backgroundColor = bg;
        }

        const top = stack[stack.length - 1];
        const mcid = top?.mcid ?? null, inArtifact = !!top?.artifact;

        const dir = (item.dir === 'rtl' || item.dir === 'ttb') ? item.dir : 'ltr';
        const angle = box.angle === -1 ? 0 : box.angle;

        // Split the item at annotation x-boundaries so a link only covers the characters it actually
        // spans, instead of the whole phrase run it merely grazes. Adjacent unlinked segments re-merge
        // in the line builder. Only horizontal runs are split; others are tagged whole.
        const len = item.str.length;
        const segments = (angle === 0 && len > 0)
            ? segmentByLinks(box.x, box.width, box.yTop, box.height, len, links)
            : [{ start: 0, end: len, link: linkForBox(box.x, box.yTop, box.width, box.height, links) }];
        for (const seg of segments) {
            const segText = item.str.slice(seg.start, seg.end);
            if (!segText) continue;
            const segX = box.x + (len ? (seg.start / len) * box.width : 0);
            const segW = len ? ((seg.end - seg.start) / len) * box.width : box.width;
            runs.push({
                text: segText,
                x: segX, yTop: box.yTop, yBaseline: box.yBaseline, width: segW, height: box.height,
                fontSize: box.fontSize,
                dir,
                angle,
                mcid, inArtifact,
                formatting,
                link: seg.link,
            });
        }
    }

    // OCR of page images is emitted through the attachment path (`emitImage` requires
    // extractAttachments), so collecting images for `ocr` alone would decode and then drop them.
    // Gate on extractAttachments only; `ocr` without it raised OCR_REQUIRES_ATTACHMENTS above.
    const images = config.extractAttachments
        ? await collectImages(pdfjs, page, layoutViewport, config, pageNumber, ops)
        : [];

    let structTree: unknown | null = null;
    if (pdfCfg.useTags) {
        try { structTree = await page.getStructTree(); } catch { structTree = null; }
    }

    if (typeof page.cleanup === 'function') { try { page.cleanup(); } catch { /* best effort */ } }

    return { pageNumber, width, height, authoredW, authoredH, authoredY1, rotation, runs, images, structTree };
}

/** Extracts images from a page's operator list, positioned in layout-viewport space. */
async function collectImages(pdfjs: any, page: any, viewport: any, config: FullOfficeParserConfig, pageNumber: number, prefetchedOps?: any): Promise<PdfImage[]> {
    const images: PdfImage[] = [];
    try {
        const ops = prefetchedOps || await page.getOperatorList();
        const fnArray = ops.fnArray;
        const argsArray = ops.argsArray;
        // Graphics-state CTM, tracked exactly as `collectColorMarks` tracks it: an image fills the
        // unit square under the CTM in force when it is painted, and that CTM is built from the page
        // `cm`s, the surrounding `q`/`Q` brackets and any form XObject's `/Matrix`. Reading back only
        // the nearest preceding `transform` gets the image's own placement matrix but none of its
        // context, so a Chrome/Skia page (`1 0 0 -1 0 H cm` then `q w 0 0 h x y cm /Im Do Q`) placed
        // every image mirrored, and an image inside a form got the form-local box.
        let ctm = identityMatrix();
        const ctmStack: number[][] = [];
        let inlineSeq = 0;

        /**
         * Encodes one decoded image object to PNG at the current CTM and records it. Shared by the
         * named-XObject, repeated-XObject and inline-image ops so all three collect identically. The
         * megapixel guard bounds an inline image's allocation (pdf.js `maxImageSize` gates decoded
         * XObjects, but an inline image's declared dimensions reach `convertToRgbaBuffer` directly).
         */
        const pushImage = async (imgObj: any, imgName: string, placement: number[] = ctm): Promise<void> => {
            if (!imgObj) return;
            if (isBrowser && !imgObj.data && imgObj.bitmap) {
                try {
                    const canvas = document.createElement('canvas');
                    canvas.width = imgObj.width; canvas.height = imgObj.height;
                    const ctx = canvas.getContext('2d');
                    if (ctx) { ctx.drawImage(imgObj.bitmap, 0, 0); imgObj.data = ctx.getImageData(0, 0, imgObj.width, imgObj.height).data; imgObj.kind = 3; }
                } catch (e) {
                    logWarning(OfficeWarningType.IMAGE_PROCESSING_FAILED, config, undefined, e);
                }
            }
            if (!(imgObj.data && imgObj.width > 0 && imgObj.height > 0)) return;
            if (imgObj.width * imgObj.height > 40_000_000) {
                logWarning(OfficeWarningType.IMAGE_PROCESSING_FAILED, config, `image on page ${pageNumber} is too large to extract (${imgObj.width}x${imgObj.height} pixels)`);
                return;
            }
            // `placement` maps the image's unit square onto the page ([a,b,c,d,e,f]); it defaults to the
            // tracked CTM but is the first tile's matrix for a repeated image. Encode to PNG now, while
            // this page's raw pixel buffer is in hand, so the large uncompressed RGBA is freed as the
            // page goes out of scope instead of being retained until the emit pass.
            const bounds = imageBounds(viewport, placement);
            try {
                const rgba = convertToRgbaBuffer(imgObj.data, imgObj.width, imgObj.height, imgObj.kind);
                const png = encodePng(imgObj.width, imgObj.height, new Uint8Array(rgba));
                images.push({ name: imgName, bounds, png, pixelWidth: imgObj.width, pixelHeight: imgObj.height });
            } catch (e) {
                logWarning(OfficeWarningType.IMAGE_EXTRACTION_FAILED, config, `on page ${pageNumber}`, e);
            }
        };

        for (let j = 0; j < fnArray.length; j++) {
            const fn = fnArray[j];
            if (fn === pdfjs.OPS.save) { ctmStack.push(ctm); continue; }
            if (fn === pdfjs.OPS.restore) { if (ctmStack.length) ctm = ctmStack.pop()!; continue; }
            if (fn === pdfjs.OPS.transform) { const m = toMatrix6(argsArray[j]); if (m) ctm = mulMatrix(ctm, m); continue; }
            if (fn === pdfjs.OPS.paintFormXObjectBegin) {
                ctmStack.push(ctm);
                const m = toMatrix6(argsArray[j]?.[0]);
                if (m) ctm = mulMatrix(ctm, m);
                continue;
            }
            if (fn === pdfjs.OPS.paintFormXObjectEnd) { if (ctmStack.length) ctm = ctmStack.pop()!; continue; }
            if (fn === pdfjs.OPS.dependency) {
                for (const dep of argsArray[j]) {
                    try {
                        // Route each dependency to the pool that actually holds it, as pdf.js does:
                        // global objects (fonts, graphics state) are `g_`-prefixed and live in
                        // `commonObjs`; everything else is page-local in `objs`. Waiting on the wrong
                        // pool means the callback never fires and the 500 ms timeout always elapses,
                        // which on a multi-font document adds up to many seconds of pure sleep per page.
                        const pool = typeof dep === 'string' && dep.startsWith('g_') ? page.commonObjs : page.objs;
                        if (pool.has(dep)) continue;
                        await new Promise<void>((resolve) => {
                            const timeout = setTimeout(resolve, 500);
                            pool.get(dep, () => { clearTimeout(timeout); resolve(); });
                        });
                    } catch (e) {
                        logWarning(OfficeWarningType.DEPENDENCY_LOAD_FAILED, config, dep, e);
                    }
                }
            }
            // A named XObject image, or a repeated (tiled) one: args[0] is the object id in either pool.
            // The repeat positions do not change the pixels, so collect the bitmap once at the current
            // placement. (`paintImageMaskXObject` is deliberately not collected: a stencil mask is painted
            // in the current fill colour, which this pass does not track, so it has no faithful bitmap.)
            if (fn === pdfjs.OPS.paintImageXObject || fn === pdfjs.OPS.paintImageXObjectRepeat) {
                const imgName = argsArray[j][0];
                // A repeated (tiled) image is placed at each `[scaleX,0,0,scaleY, x,y]` from the args,
                // not at the bare CTM, so record the FIRST tile's box; the bitmap is collected once.
                let placement = ctm;
                if (fn === pdfjs.OPS.paintImageXObjectRepeat) {
                    const a = argsArray[j];
                    const pos = a[3];
                    if (typeof a[1] === 'number' && typeof a[2] === 'number' && pos && pos.length >= 2) {
                        placement = mulMatrix(ctm, [a[1], 0, 0, a[2], pos[0], pos[1]]);
                    }
                }
                try {
                    let hasObj = page.objs.has(imgName);
                    let targetObjs = page.objs;
                    if (!hasObj && page.commonObjs.has(imgName)) { hasObj = true; targetObjs = page.commonObjs; }
                    if (!hasObj) continue;
                    const imgObj: any = await new Promise((resolve) => targetObjs.get(imgName, (d: any) => resolve(d)));
                    await pushImage(imgObj, imgName, placement);
                } catch {
                    // Image access failed, continue.
                }
                continue;
            }
            // An inline image (BI...EI): the decoded pixel object is passed directly in the args, with
            // no id, so it never goes through page.objs.
            if (fn === pdfjs.OPS.paintInlineImageXObject) {
                try { await pushImage(argsArray[j][0], `inline_p${pageNumber}_${inlineSeq++}`); } catch { /* continue */ }
                continue;
            }
        }
    } catch (e) {
        logWarning(OfficeWarningType.IMAGE_EXTRACTION_FAILED, config, `from page ${pageNumber}`, e);
    }
    return images;
}

/** Maps an image CTM to a layout-viewport box (the image fills the unit square under the CTM). */
function imageBounds(viewport: any, ctm: number[]): { x: number; y: number; width: number; height: number } {
    const [a, b, c, d, e, f] = ctm;
    const xs = [e, e + a, e + c, e + a + c];
    const ys = [f, f + b, f + d, f + b + d];
    const rect = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
    const [vx1, vy1, vx2, vy2] = toViewportRect(viewport, rect);
    return { x: vx1, y: vy1, width: vx2 - vx1, height: vy2 - vy1 };
}

/**
 * Parses a PDF file and extracts content.
 *
 * @param buffer - The PDF file buffer
 * @param config - Parser configuration
 * @returns Promise resolving to the parsed AST
 */
export const parsePdf = async (buffer: Buffer, config: FullOfficeParserConfig): Promise<OfficeParserAST> => {
    checkAbortSignal(config.abortSignal);
    const pdfjs = await loadPdfJs();
    const pdfCfg = resolvePdfLayoutConfig(config);

    // --- Worker configuration ---
    const workerSrc = config.pdfWorkerSrc;
    if (isBrowser) {
        pdfjs.GlobalWorkerOptions.workerSrc = workerSrc;
    } else {
        assertNode('pdf-worker-auto-resolution', config);
        let resolved = false;
        if (workerSrc !== DEFAULT_OFFICE_PARSER_CONFIG.pdfWorkerSrc && workerSrc !== '') {
            pdfjs.GlobalWorkerOptions.workerSrc = workerSrc;
            resolved = true;
        } else if ((globalThis as any).pdfjsWorker) {
            resolved = true;
        } else {
            try {
                // @ts-ignore - 'require' is available in Node.js/CommonJS environment
                const localWorkerPath = require.resolve('pdfjs-dist/legacy/build/pdf.worker.mjs');
                const { pathToFileURL } = await import('url');
                pdfjs.GlobalWorkerOptions.workerSrc = pathToFileURL(localWorkerPath).href;
                resolved = true;
            } catch (e) {
                logWarning(OfficeWarningType.PDF_WORKER_FALLBACK, config, undefined, e);
            }
        }
        if (!resolved) pdfjs.GlobalWorkerOptions.workerSrc = workerSrc;
    }

    // Password / onPassword are top-level (they apply to every encryptable format); PDF reads them
    // the same way the OOXML/ODF decryptors do.
    const onPassword = config.onPassword;
    // Cap callback-driven retries so an onPassword that keeps returning a wrong password can't loop.
    const MAX_PASSWORD_ATTEMPTS = 3;
    let password: string | undefined = config.password || undefined;
    let passwordAttempts = 0;

    // Bound how many pixels pdf.js will decode per image. When we never read image bytes, set 1 so
    // image XObjects are skipped entirely: colour marks come from the operator list, not the decoded
    // bitmap, and getOperatorList otherwise decodes every image on every page for nothing. We read
    // image bytes only when collecting attachments (page-image OCR runs over collected images too, so
    // `ocr` alone - without extractAttachments - collects nothing and must NOT raise this). When we do
    // need pixels, cap at a generous 40 megapixels so a decompression-bomb image (a few bytes of
    // headers declaring enormous dimensions) cannot drive a multi-GB, uncatchable allocation.
    // `ocr: true` without `extractAttachments: true` collects no images, so no OCR runs and no page
    // image is decoded for it. The OCR_REQUIRES_ATTACHMENTS warning is raised centrally in OfficeParser
    // (it applies to every format, not just PDF), so it is not repeated here.
    const maxImageSize = config.extractAttachments ? 40_000_000 : 1;

    // Open the document, retrying with an onPassword-supplied password when the PDF is encrypted.
    while (true) {
        checkAbortSignal(config.abortSignal);
        // A fresh copy per attempt: pdf.js transfers the data buffer to its worker, detaching it, so
        // reusing the same Uint8Array on a password retry would fail with a transfer error.
        const loadingTask = pdfjs.getDocument({
            data: new Uint8Array(buffer),
            verbosity: 0,
            isEvalSupported: false,
            maxImageSize,
            password,
        });

        let pdfDocument;
        try {
            pdfDocument = await loadingTask.promise;
        } catch (e: any) {
            try { await loadingTask.destroy(); } catch { /* best effort cleanup */ }
            if (e?.name === 'PasswordException') {
                const need = pdfjs.PasswordResponses?.NEED_PASSWORD;
                const reason: 'required' | 'incorrect' = e.code === need ? 'required' : 'incorrect';
                if (onPassword && passwordAttempts < MAX_PASSWORD_ATTEMPTS) {
                    passwordAttempts++;
                    const supplied = await onPassword(reason);
                    if (supplied) { password = supplied; continue; }
                }
                throw getOfficeError(reason === 'required' ? OfficeErrorType.PASSWORD_REQUIRED : OfficeErrorType.PASSWORD_INCORRECT, config);
            }
            const message = e instanceof Error ? e.message : String(e);
            if (message.includes('workerSrc') || message.includes('No "GlobalWorkerOptions.workerSrc" specified')) {
                throw getOfficeError(OfficeErrorType.PDF_WORKER_MISSING, config);
            }
            throw e;
        }

        try {
            return await buildAst(pdfjs, pdfDocument, config, pdfCfg);
        } finally {
            try { await loadingTask.destroy(); } catch { /* best effort cleanup */ }
        }
    }
};

/** Assembles the AST from an opened document. Separated so the caller can guarantee task cleanup. */
async function buildAst(pdfjs: any, pdfDocument: any, config: FullOfficeParserConfig, pdfCfg: PdfLayoutConfig): Promise<OfficeParserAST> {
    const content: OfficeContentNode[] = [];
    const attachments: OfficeAttachment[] = [];
    const numPages = pdfDocument.numPages;

    // --- Metadata ---
    const meta = await pdfDocument.getMetadata().catch(() => ({ info: {} }));
    const info = (meta.info || {}) as Record<string, unknown>;
    const metadata: OfficeMetadata = {
        pages: numPages,
        title: info?.Title as string | undefined,
        author: info?.Author as string | undefined,
        subject: info?.Subject as string | undefined,
        // PDF /Subject is the document description; /Keywords are the keywords (previously /Keywords
        // was wrongly used as the description and keywords were dropped).
        description: info?.Subject as string | undefined,
        keywords: info?.Keywords as string | undefined,
        created: parseOfficeDate(info?.CreationDate as string | undefined),
        modified: parseOfficeDate(info?.ModDate as string | undefined),
    };
    if (typeof info?.Language === 'string') metadata.language = info.Language as string;

    const standardPdfInfoKeys = new Set([
        'Title', 'Author', 'Subject', 'Keywords', 'Creator', 'Producer',
        'CreationDate', 'ModDate', 'Trapped', 'IsAcroFormPresent', 'IsXFAPresent',
        'IsCollectionPresent', 'IsSignaturesPresent', 'PDFFormatVersion'
    ]);
    if (info) {
        metadata.nativeProperties = {};
        for (const [key, val] of Object.entries(info)) {
            if (key === 'Custom' && typeof val === 'object' && !Array.isArray(val) && !(val instanceof Date) && val !== null) {
                for (const [ck, cv] of Object.entries(val)) setOwn(metadata.nativeProperties, ck, cv);
            } else {
                setOwn(metadata.nativeProperties, key, val);
            }
        }
        const customProperties: Record<string, string | number | boolean | Date> = {};
        for (const key of Object.keys(info)) {
            if (standardPdfInfoKeys.has(key)) continue;
            const val = info[key];
            if (val === null || val === undefined) continue;
            if (key === 'Custom' && typeof val === 'object' && !Array.isArray(val) && !(val instanceof Date)) {
                for (const [ck, cv] of Object.entries(val)) {
                    if (cv === null || cv === undefined) continue;
                    if (typeof cv === 'string' || typeof cv === 'number' || typeof cv === 'boolean' || cv instanceof Date) setOwn(customProperties, ck, cv);
                }
                continue;
            }
            if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean' || val instanceof Date) setOwn(customProperties, key, val);
        }
        if (Object.keys(customProperties).length > 0) metadata.customProperties = customProperties;
    }
    if (meta.metadata) {
        if (!metadata.nativeProperties) metadata.nativeProperties = {};
        const xmp: any = meta.metadata;
        metadata.nativeProperties['XMP'] = typeof xmp.getAll === 'function' ? xmp.getAll() : xmp;
    }

    // Tagged flag for consumers.
    let markInfo: any = null;
    try { markInfo = await pdfDocument.getMarkInfo(); } catch { markInfo = null; }
    if (!metadata.nativeProperties) metadata.nativeProperties = {};
    metadata.nativeProperties['tagged'] = !!(markInfo && markInfo.Marked);
    if (markInfo) metadata.nativeProperties['markInfo'] = markInfo;

    // Document permissions (which user actions the file allows). `null` means all actions allowed;
    // pdf.js cannot validate signatures, so we only report presence, never validity.
    try {
        const perms: number[] | null = await pdfDocument.getPermissions();
        metadata.nativeProperties['permissions'] = perms ? permissionNames(pdfjs, perms) : 'all';
    } catch { /* not available */ }

    // Optional-content group (layer) names and default visibility, for consumers that care which
    // layers exist. Text inside hidden layers is still extracted (getTextContent ignores visibility).
    try {
        const oc = await pdfDocument.getOptionalContentConfig();
        const layers = listOptionalContentLayers(oc);
        if (layers.length) metadata.nativeProperties['layers'] = layers;
    } catch { /* no optional content */ }

    // AcroForm field values (filled form data). Many real-world PDFs (applications, invoices) carry
    // their content here rather than as page text. Reported structurally; extraction, not validation.
    try {
        const fieldObjects = await pdfDocument.getFieldObjects();
        if (fieldObjects) {
            const fields: Record<string, unknown> = {};
            for (const name of Object.keys(fieldObjects)) {
                const arr = (fieldObjects as any)[name];
                const f = Array.isArray(arr) ? arr[0] : arr;
                if (!f || typeof f !== 'object') continue;
                const entry: Record<string, unknown> = { value: f.value, type: f.type };
                if (f.defaultValue !== undefined && f.defaultValue !== null) entry.defaultValue = f.defaultValue;
                setOwn(fields, name, entry);
            }
            if (Object.keys(fields).length) metadata.nativeProperties['formFields'] = fields;
        }
    } catch { /* no AcroForm */ }

    // Printed page labels (e.g. roman-numeral front matter), distinct from the physical page index.
    let pageLabels: (string | null)[] | null = null;
    try { pageLabels = await pdfDocument.getPageLabels(); } catch { pageLabels = null; }

    // --- Embedded file attachments ---
    try {
        const embeddedFiles = await pdfDocument.getAttachments();
        if (embeddedFiles && config.extractAttachments) {
            for (const name in embeddedFiles) {
                const file = embeddedFiles[name];
                attachments.push(createAttachment(file.filename, Buffer.from(file.content)));
            }
        }
    } catch (e) {
        logWarning(OfficeWarningType.ATTACHMENT_EXTRACTION_FAILED, config, undefined, e);
    }

    // --- Single collection pass ---
    const pageNumbers = parsePageRange(config.pdfParserConfig?.pageRange, numPages);
    const fontCache = new Map<string, ResolvedFont>();
    const destCache = new Map<string, SectionTarget | null>();
    const sectionLinks = new SectionLinks();
    const extracts: PageExtract[] = [];
    for (const pageNum of pageNumbers) {
        checkAbortSignal(config.abortSignal);
        try {
            extracts.push(await collectPage(pdfjs, pdfDocument, pageNum, config, pdfCfg, fontCache, destCache, sectionLinks));
        } catch (e: any) {
            logWarning(OfficeWarningType.PAGE_LOAD_FAILED, config, pageNum, e);
        }
    }

    const allRuns = extracts.flatMap(e => e.runs);
    // Layout always runs in the authored (rotation-0) frame with geometry ON, whatever the caller
    // asked for: reading-order splicing, table span inference and header/footer bands are geometric
    // decisions, and making them depend on `ignorePageGeometry` or on `/Rotate` is how nodes ended up
    // ordered differently (or arbitrarily) between the two. `finalizeBounds` then rotates the boxes
    // into rendered space, or strips them entirely, once the page is assembled.
    const layoutCfg: PdfLayoutConfig = { ...pdfCfg, includeBounds: true };
    const docCtx = computeDocContext(allRuns, layoutCfg, config.newlineDelimiter);

    // Warn when extracted text is mostly unmappable glyphs (broken/missing ToUnicode), so consumers
    // can tell "genuinely empty" from "font could not be decoded" and reach for OCR.
    warnIfEncodingSuspect(allRuns, config);

    // Warn when a document yields essentially no text and OCR is off: it is very likely scanned, and
    // silent empty output is otherwise indistinguishable from a genuine failure. Only when OCR is off,
    // since with OCR the caller is already handling image-only content.
    if (!config.ocr) {
        // Scale the "looks empty" threshold to the pages actually processed, not the whole document:
        // with a pageRange, `allRuns` covers only the selected pages, so comparing against the full
        // numPages would spuriously warn whenever a small slice of a large PDF is requested.
        const processedPages = pageNumbers.length;
        const textChars = allRuns.reduce((sum, r) => sum + r.text.replace(/\s/g, '').length, 0);
        if (processedPages > 0 && textChars < Math.max(10, processedPages)) {
            logWarning(OfficeWarningType.PDF_NO_TEXT_EXTRACTED, config, processedPages);
        }
    }

    // Tagged-structure trust: the document must declare it is tagged and not flag it as suspect.
    const docTagged = !!(markInfo && markInfo.Marked);
    const docTrusted = docTagged && !markInfo.Suspects;
    let structWarned = false;
    const warnStruct = (reason: string) => {
        if (structWarned) return;
        structWarned = true;
        logWarning(OfficeWarningType.PDF_STRUCT_TREE_UNRELIABLE, config, reason);
    };
    if (pdfCfg.useTags && docTagged && !docTrusted) warnStruct('the document flags its tags as suspect');

    const auxHeaders: OfficeContentNode[] = [];
    const auxFooters: OfficeContentNode[] = [];

    const taggedListCounter = { n: 0 };
    let imageCounter = 0;
    for (const extract of extracts) {
        checkAbortSignal(config.abortSignal);
        const pageCtx: PageContext = { pageNumber: extract.pageNumber, authoredW: extract.authoredW, authoredH: extract.authoredH, rotation: extract.rotation };
        // The frame every node is built in: authored space, so a `/Rotate` page's boxes stay
        // comparable to the runs they came from. `finalizeBounds` maps them at the end of the page.
        const layoutCtx: PageContext = { ...pageCtx, rotation: 0 };

        // Artifact runs split by where they render: only the top and bottom bands are running
        // headers/footers. Everything else an Artifact scope covers (watermarks, figure labels,
        // decorative headings, whatever an InDesign or Acrobat export marked as artifact) is real
        // page text and joins the body flow instead of being dropped on the floor.
        const { header: headerRuns, footer: footerRuns, body: midArtifacts } = splitArtifacts(extract);
        const midSet = new Set(midArtifacts);
        const bodyRuns = extract.runs.filter(r => !r.inArtifact || midSet.has(r));

        let pageContent: OfficeContentNode[] | null = null;
        let fromTags = false;

        // Tagged path: use the structure tree when trusted and it covers most of the page's text.
        if (pdfCfg.useTags && docTrusted && extract.structTree) {
            const runsByMcid = new Map<string, RawRun[]>();
            for (const r of bodyRuns) {
                if (!r.mcid) continue;
                const list = runsByMcid.get(r.mcid);
                if (list) list.push(r); else runsByMcid.set(r.mcid, [r]);
            }
            // buildTaggedNodes guards against a hostile deep/cyclic tree internally, but wrap the call
            // too: any unexpected failure here must degrade to the geometry path, never fail the parse.
            let nodes: OfficeContentNode[] = [], coveredMcids = new Set<string>();
            try {
                ({ nodes, coveredMcids } = buildTaggedNodes(extract.structTree, runsByMcid, layoutCtx, docCtx, { ignoreNotes: config.ignoreNotes, listCounter: taggedListCounter }));
            } catch { warnStruct('the tag tree could not be processed'); }
            // Coverage measures the tag tree against the page's *tagged* text, so artifact runs
            // (which by definition live outside the structure) never drag the trust signal down.
            const textMcids = new Set(bodyRuns.filter(r => !r.inArtifact && r.text.trim() && r.mcid).map(r => r.mcid as string));
            let coveredText = 0;
            for (const m of textMcids) if (coveredMcids.has(m)) coveredText++;
            const coverage = textMcids.size ? coveredText / textMcids.size : 1;
            if (coverage >= 0.7) {
                pageContent = nodes;
                fromTags = true;
                // Stitch in any runs the tags did not cover, via the geometric path, splicing each
                // recovered node into reading order by its y instead of dumping them at the page end.
                const leftover = bodyRuns.filter(r => r.text.trim() && (!r.mcid || !coveredMcids.has(r.mcid)));
                if (leftover.length) {
                    for (const ln of geometricNodes(leftover, layoutCtx, docCtx, layoutCfg)) {
                        const y = ln.bounds?.y ?? Infinity;
                        const at = pageContent.findIndex(n => (n.bounds?.y ?? Infinity) > y);
                        if (at < 0) pageContent.push(ln); else pageContent.splice(at, 0, ln);
                    }
                    // Mid-page artifact text is expected to be outside the tag tree, so it is not
                    // evidence of a broken one: only untagged *body* text warrants the warning.
                    if (leftover.some(r => !r.inArtifact)) warnStruct('some text on a page was outside the tag tree');
                }
            } else {
                warnStruct('the tag tree covered too little of the page text');
            }
        }

        // Geometric fallback (untagged, distrusted, or low coverage).
        if (!pageContent) pageContent = geometricNodes(bodyRuns, layoutCtx, docCtx, layoutCfg);

        // Hybrid recovery: on the tagged path, additionally fold a contiguous run of loose sibling
        // paragraphs whose geometry forms a clean grid (a calendar tagged as one paragraph per day,
        // rather than as a Table) into a standard table, without disturbing the tagged tables. The
        // geometric path already recovers grid tables via detectTables; this brings the tagged path to
        // parity. Runs on authored boxes here, before finalizeBounds, so it works under
        // ignorePageGeometry too.
        if (fromTags) pageContent = recoverTaggedGrids(pageContent);

        // Images: emit as attachments/OCR, then splice each into the flow before the first text node
        // that sits lower on the page, so reading order is preserved without reordering text.
        for (const img of extract.images) {
            const node = await emitImage(img, extract.pageNumber, ++imageCounter, config, attachments);
            if (!node) continue;
            const y = node.bounds?.y ?? Infinity;
            const at = pageContent.findIndex(n => n.type !== 'image' && (n.bounds?.y ?? Infinity) > y);
            if (at < 0) pageContent.push(node); else pageContent.splice(at, 0, node);
        }

        // Running headers/footers route to auxiliary unless the caller asked for them to be dropped.
        if (!config.ignoreHeadersAndFooters) {
            const pageHeaders: OfficeContentNode[] = [];
            const pageFooters: OfficeContentNode[] = [];
            for (const node of geometricNodes(headerRuns, layoutCtx, docCtx, layoutCfg, false)) pageHeaders.push(retypeAsHeaderFooter(node, 'header'));
            for (const node of geometricNodes(footerRuns, layoutCtx, docCtx, layoutCfg, false)) pageFooters.push(retypeAsHeaderFooter(node, 'footer'));
            finalizeBounds(pageHeaders, pageCtx, pdfCfg.includeBounds);
            finalizeBounds(pageFooters, pageCtx, pdfCfg.includeBounds);
            auxHeaders.push(...pageHeaders);
            auxFooters.push(...pageFooters);
        }

        finalizeBounds(pageContent, pageCtx, pdfCfg.includeBounds);

        const pageNode: OfficeContentNode = {
            type: 'page',
            children: pageContent,
            text: pageContent.map(n => n.text).join(config.newlineDelimiter),
            metadata: { pageNumber: extract.pageNumber },
        };
        if (pdfCfg.includeBounds && pageNode.type === 'page' && pageNode.metadata) {
            pageNode.metadata.pageWidth = Math.round(extract.width * 100) / 100;
            pageNode.metadata.pageHeight = Math.round(extract.height * 100) / 100;
            if (extract.rotation) pageNode.metadata.rotation = extract.rotation;
        }
        if (pageNode.type === 'page' && pageNode.metadata) {
            const label = pageLabels?.[extract.pageNumber - 1];
            if (label && label !== String(extract.pageNumber)) pageNode.metadata.pageLabel = label;
        }
        content.push(pageNode);
    }

    // Document outline (bookmarks / TOC), when present, into auxiliary. A malformed or hostile
    // outline must never cost the caller the whole document, so anything thrown here is a warning.
    let outline: OfficeContentNode[] | undefined;
    if (!config.ignoreInternalLinks) {
        try { outline = await buildOutline(pdfDocument, destCache, sectionLinks, config); }
        catch { logWarning(OfficeWarningType.PDF_OUTLINE_TRUNCATED, config, 'it could not be read'); }
    }

    // Now that every page's headings and their positions are known, point each internal link at the
    // nearest heading on its target page (falling back to the page itself), rather than a page jump.
    resolveSectionLinks(content, extracts, sectionLinks, [content, auxHeaders, auxFooters, outline ?? []]);

    const auxiliary = (auxHeaders.length || auxFooters.length || outline)
        ? { ...(auxHeaders.length ? { headers: auxHeaders } : {}), ...(auxFooters.length ? { footers: auxFooters } : {}), ...(outline ? { outline } : {}) }
        : undefined;

    return createAST('pdf', metadata, content, attachments, config, auxiliary);
}

/**
 * Runs the geometric (untagged) assembly path over a set of runs. `detectGridTables` recovers grid
 * tables from the lines before column segmentation (off for margin artifacts, which are never
 * tables).
 */
function geometricNodes(runs: RawRun[], pageCtx: PageContext, docCtx: DocContext, pdfCfg: PdfLayoutConfig, detectGridTables = true): OfficeContentNode[] {
    const lines = buildLines(runs, pdfCfg);
    const { tables, consumed } = detectGridTables
        ? detectTables(lines, pageCtx, docCtx)
        : { tables: [] as { node: OfficeContentNode; y: number }[], consumed: new Set<typeof lines[number]>() };
    const remaining = consumed.size ? lines.filter(l => !consumed.has(l)) : lines;

    const out: OfficeContentNode[] = [];
    for (const block of segmentIntoBlocks(remaining, pdfCfg)) out.push(...blockToNodes(block, pageCtx, docCtx));
    // Splice each detected table into the flow at its vertical position.
    for (const t of tables) {
        const ty = t.node.bounds?.y ?? Infinity;
        const at = out.findIndex(n => (n.bounds?.y ?? Infinity) > ty);
        if (at < 0) out.push(t.node); else out.splice(at, 0, t.node);
    }
    // Rescue rotated text (90/180/270), which the horizontal line builder skips, so it is not
    // silently dropped. It is appended after the main flow, in the source content order (which is
    // usually the correct reading order); precise visual ordering of rotated text is a limitation.
    out.push(...rotatedTextNodes(runs, pageCtx, pdfCfg));
    return out;
}

/** Recovers rotated (90/180/270) runs as trailing paragraphs, one per angle, in content order. */
function rotatedTextNodes(runs: RawRun[], pageCtx: PageContext, pdfCfg: PdfLayoutConfig): OfficeContentNode[] {
    const out: OfficeContentNode[] = [];
    for (const angle of [90, 270, 180] as const) {
        const group = runs.filter(r => r.angle === angle && r.text.trim().length > 0);
        if (!group.length) continue;
        // Joined once, a space between runs that do not already meet at one; testing the end of the
        // growing text for each run instead copied all of it every time.
        const parts: string[] = [];
        let endsInSpace = true;
        for (const r of group) {
            if (!endsInSpace && !/^\s/.test(r.text)) parts.push(' ');
            parts.push(r.text);
            if (r.text) endsInSpace = /\s/.test(r.text[r.text.length - 1]);
        }
        const text = parts.join('').replace(/\s+/g, ' ').trim();
        if (!text) continue;
        const node: OfficeContentNode = { type: 'paragraph', text, children: [{ type: 'text', text }] };
        const box = unionAll(group.map(r => ({ x: r.x, y: r.yTop, width: r.width, height: r.height })));
        if (pdfCfg.includeBounds && box) {
            const rendered = roundBounds(rotateBoundsToRendered(box, pageCtx.rotation, pageCtx.authoredW, pageCtx.authoredH));
            node.bounds = rendered;
            if (node.children && node.children[0]) node.children[0].bounds = rendered;
        }
        out.push(node);
    }
    return out;
}

/**
 * Splits a page's `/Artifact` text into the running-header band, the running-footer band, and
 * everything in between.
 *
 * The bands are measured in RENDERED space (top and bottom 15% of the page as the reader sees it),
 * so a `/Rotate 90` page - whose running footer sits at authored *right*, halfway down in authored y
 * - is still classified as a footer instead of being mistaken for body text. The middle is returned
 * as body: an Artifact scope marks content as non-structural, not as unwanted, and a watermark, a
 * figure label or a decorative heading is text the caller asked to extract.
 */
function splitArtifacts(extract: PageExtract): { header: RawRun[]; footer: RawRun[]; body: RawRun[] } {
    const header: RawRun[] = [], footer: RawRun[] = [], body: RawRun[] = [];
    const renderedH = extract.rotation % 180 === 0 ? extract.authoredH : extract.authoredW;
    for (const r of extract.runs) {
        if (!r.inArtifact || !r.text.trim()) continue;
        const box = { x: r.x, y: r.yTop, width: r.width, height: r.height };
        const y = rotateBoundsToRendered(box, extract.rotation, extract.authoredW, extract.authoredH).y;
        if (y < 0.15 * renderedH) header.push(r);
        else if (y > 0.85 * renderedH) footer.push(r);
        else body.push(r);
    }
    return { header, footer, body };
}

/**
 * Converts a page's nodes from the authored frame the layout runs in to what the AST reports:
 * rendered-space boxes on a `/Rotate` page, or no boxes at all under `ignorePageGeometry`. Reading
 * order, table span inference and the header/footer bands are all decided on the authored boxes
 * before this runs, so neither setting can change the structure or the order of the nodes - only
 * whether, and in which frame, their geometry is reported.
 */
function finalizeBounds(nodes: OfficeContentNode[], page: PageContext, includeBounds: boolean): void {
    if (includeBounds && page.rotation === 0) return;   // authored space already is rendered space
    const stack = [...nodes];
    while (stack.length) {
        const n = stack.pop()!;
        if (n.bounds) {
            if (!includeBounds) delete n.bounds;
            else n.bounds = roundBounds(rotateBoundsToRendered(n.bounds, page.rotation, page.authoredW, page.authoredH));
        }
        for (const c of n.children || []) stack.push(c);
        for (const c of n.notes || []) stack.push(c);
        for (const c of n.comments || []) stack.push(c);
    }
}

/** Wraps a paragraph produced from margin artifacts as a header/footer node. */
function retypeAsHeaderFooter(node: OfficeContentNode, type: 'header' | 'footer'): OfficeContentNode {
    return { type, text: node.text, children: node.children, bounds: node.bounds, metadata: { type: 'default' } };
}

/**
 * Encodes/OCRs one image and returns its positioned image node (or null when nothing was emitted).
 * Its box is authored-space, like every other node's at this stage; {@link finalizeBounds} maps it.
 */
async function emitImage(
    img: PdfImage, pageNumber: number, index: number, config: FullOfficeParserConfig,
    attachments: OfficeAttachment[],
): Promise<OfficeContentNode | null> {
    if (!config.extractAttachments) return null;
    const attachmentName = `pdf_image_p${pageNumber}_${index}.png`;
    try {
        // Pixels were PNG-encoded at collection time; here we only name, attach and optionally OCR them.
        const png = img.png;
        const attachment = createAttachment(attachmentName, png);
        attachment.mimeType = 'image/png';
        if (config.ocr && img.pixelWidth >= 10 && img.pixelHeight >= 10) {
            const ocrText = await ocrDuringParse(png, config, attachmentName, 'image/png');
            if (ocrText !== undefined) attachment.ocrText = ocrText;
        }
        attachments.push(attachment);
        const metadata: ImageMetadata = { attachmentName };
        const node: OfficeContentNode = { type: 'image', text: attachment.ocrText || '', metadata };
        node.bounds = roundBounds(img.bounds);
        return node;
    } catch (e: any) {
        // A cancelled parse rejects; it is not a failed image.
        if (e?.name === 'AbortError') throw e;
        logWarning(OfficeWarningType.IMAGE_EXTRACTION_FAILED, config, attachmentName, e);
        return null;
    }
}

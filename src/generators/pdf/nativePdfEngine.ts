/**
 * A first-party PDF layout engine built directly on `pdf-lib`, an alternative to the default
 * HTML-through-a-headless-browser route (`PdfGenerator`'s Puppeteer path).
 *
 * It walks the AST and paints it onto pages itself - text with inline bold/italic/size/color,
 * headings, lists (bulleted and numbered, nested), grid tables with borders, embedded images, page
 * breaks and horizontal rules - handling word wrapping and pagination as it goes. It needs no browser
 * and runs unchanged in Node and the browser, so it is the way to produce a real PDF client-side (the
 * browser Puppeteer path can only hand back HTML). Its trade-off is fidelity: it draws in the fonts it
 * is given (`pdfConfig.fonts`, `pdfConfig.fallbackFonts`) or, given none, in the Standard-14 Helvetica
 * and Courier faces, never in the document's own, and its layout is a clean reflow, not a pixel match
 * of the HTML renderer. Choose it with `pdfConfig.engine: 'native'`.
 *
 * `pdf-lib` is an optional peer dependency, imported lazily so it is only required when this engine
 * is selected; `@pdf-lib/fontkit` is another, required only when fonts are given.
 *
 * @module generators/pdf/nativePdfEngine
 */

import { EmbedMetadata, FullGeneratorConfig, ImageMode, OfficeContentNode, OfficeErrorType, OfficeMetadata, OfficeParserAST, OfficeWarningType, TextFormatting } from '../../types.js';
import { getAbortError, getOfficeError } from '../../utils/errorUtils.js';
import { documentLanguage, embedUrl, isHeaderRow, lengthToPt, paperSizePt, resolveImageMode, sniffImageSize } from '../../utils/officeGenUtils.js';
import { decodeBmp, decodeGif, encodePng } from '../../utils/rasterUtils.js';

/**
 * Megapixel ceiling for an embedded image. `embedPng`/`embedJpg` decode the full bitmap, so a
 * decompression-bomb image (a few header bytes declaring enormous dimensions) would force a multi-GB,
 * uncatchable allocation. A real image is far below this; a larger one degrades to its alt text.
 */
const MAX_IMAGE_PIXELS = 40_000_000;

/**
 * Code points WinAnsi (CP1252) encodes beyond Latin-1, which the Standard-14 fonts accept (smart
 * quotes, dashes, bullet, ellipsis, trademark, euro, etc.). Everything outside Latin-1 and this set
 * has no glyph in those fonts and makes pdf-lib throw, so it is replaced rather than drawn.
 */
const CP1252_EXTRA = new Set([0x20AC, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x017D, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x017E, 0x0178]);

/** Whether a Standard-14 font can draw a code point: the WinAnsi-encodable subset. */
function winAnsiCovers(cp: number): boolean {
    return (cp >= 0x20 && cp <= 0x7E) || (cp >= 0xA0 && cp <= 0xFF) || CP1252_EXTRA.has(cp);
}

/**
 * Characters that draw nothing of their own: zero-width spaces and joiners, the word joiner, the BOM,
 * the bidi marks and embeddings, and the variation selectors. A font that has one keeps it (a joiner
 * changes how its neighbours are shaped); otherwise it is dropped, and is no loss to report.
 */
function isInvisible(cp: number): boolean {
    return cp === 0x200B || cp === 0x200C || cp === 0x200D || cp === 0x2060 || cp === 0xFEFF
        || cp === 0x200E || cp === 0x200F || (cp >= 0x202A && cp <= 0x202E) || (cp >= 0x2066 && cp <= 0x2069)
        || (cp >= 0xFE00 && cp <= 0xFE0F);
}

/** Hebrew, Arabic and the other right-to-left blocks of the Basic Multilingual Plane. */
function isRtlChar(cp: number): boolean {
    return (cp >= 0x0590 && cp <= 0x08FF) || (cp >= 0xFB1D && cp <= 0xFDFF) || (cp >= 0xFE70 && cp <= 0xFEFF);
}

const LETTER_OR_DIGIT = /[\p{L}\p{N}]/u;
const OTHER_SPACE = /[\p{Zs}\u000B\u000C\u2028\u2029]/u;
const CONTROL = /\p{Cc}/u;
/** The most code points shaped as one string (see NativeLayout.segments). */
const MAX_SHAPED = 256;

const COMBINING_MARK = /\p{M}/u;

/** The direction a stretch of text is read in: right-to-left when it has any such letter, else left-to-right when it has a letter or digit. */
function directionOf(text: string): 'ltr' | 'rtl' | 'neutral' {
    let dir: 'ltr' | 'neutral' = 'neutral';
    for (const ch of text) {
        if (isRtlChar(ch.codePointAt(0)!)) return 'rtl';
        if (dir === 'neutral' && LETTER_OR_DIGIT.test(ch)) dir = 'ltr';
    }
    return dir;
}

/** Whether text is a right-to-left paragraph: its first letter is a right-to-left one. */
function startsRtl(text: string): boolean {
    for (const ch of text) {
        if (isRtlChar(ch.codePointAt(0)!)) return true;
        if (LETTER_OR_DIGIT.test(ch)) return false;
    }
    return false;
}

/** One thing on a line, for ordering: a word (or part of one) or the gap after it. */
interface LineItem { w: number; dir: 'ltr' | 'rtl' | 'neutral'; }

/**
 * The order a line's items are drawn in, left to right, given the order they are read in. This is the
 * Unicode bidirectional algorithm reduced to what a line of words needs: a right-to-left word has an
 * odd level, a left-to-right one an even level, and a gap or a word of punctuation takes the level of
 * its neighbours when they agree and the paragraph's otherwise; then each run at or above a level is
 * reversed, from the highest level down. A line with no right-to-left word is returned as it is.
 */
function visualOrder<T extends LineItem>(items: T[], baseRtl: boolean): T[] {
    if (!baseRtl && !items.some(it => it.dir === 'rtl')) return items;
    const base = baseRtl ? 1 : 0;
    const strong = (it: LineItem) => it.dir === 'rtl' ? 1 : it.dir === 'ltr' ? (baseRtl ? 2 : 0) : -1;
    const levels = items.map(strong);
    for (let i = 0; i < levels.length; i++) {
        if (levels[i] !== -1) continue;
        let j = i;
        while (j < levels.length && levels[j] === -1) j++;
        const before = i > 0 ? levels[i - 1] : base, after = j < levels.length ? levels[j] : base;
        for (let k = i; k < j; k++) levels[k] = before === after ? before : base;
        i = j - 1;
    }
    const order = items.map((_, i) => i);
    for (let level = Math.max(...levels); level >= 1; level--) {
        for (let i = 0; i < order.length; i++) {
            if (levels[order[i]] < level) continue;
            let j = i;
            while (j < order.length && levels[order[j]] >= level) j++;
            for (let a = i, b = j - 1; a < b; a++, b--) { const t = order[a]; order[a] = order[b]; order[b] = t; }
            i = j - 1;
        }
    }
    return order.map(i => items[i]);
}

/** Lazily loads `@pdf-lib/fontkit`, which pdf-lib embeds a font file with, throwing a typed, actionable error when it is not installed. */
async function loadFontkit(config: OfficeParserAST['config']): Promise<any> {
    try {
        // @ts-ignore - optional peer dependency, resolved at runtime
        const mod = await import('@pdf-lib/fontkit');
        return mod.default || mod;
    } catch (e) {
        throw getOfficeError(OfficeErrorType.PDF_GENERATION_FAILED, config,
            "Embedding fonts in the native PDF engine (pdfConfig.fonts, pdfConfig.fallbackFonts) requires the optional peer dependency '@pdf-lib/fontkit'. Install it with `npm install @pdf-lib/fontkit`, or pass no fonts to draw in the Standard-14 ones.");
    }
}

/** How many conversions are drawing with embedded fonts right now (see {@link borrowGeneratorRuntime}). */
let generatorRuntimeUsers = 0;

/**
 * `@pdf-lib/fontkit` 1.x shapes Indic scripts with a generator that was compiled to call a global
 * `regeneratorRuntime`, which the package does not bring: shaping a word of Hindi threw
 * "regeneratorRuntime is not defined". When the page has none, one is lent for as long as a conversion
 * draws with embedded fonts, and taken back when the last such conversion ends. It is the part of the
 * runtime that generator uses (a context with `prev`, `next` and `stop()`), not a general one, which is
 * why it does not stay. Returns the function that gives it back.
 */
function borrowGeneratorRuntime(): () => void {
    const scope = globalThis as any;
    if (generatorRuntimeUsers === 0 && scope.regeneratorRuntime !== undefined) return () => { /* the page has its own */ };
    if (generatorRuntimeUsers++ === 0) {
        scope.regeneratorRuntime = {
            mark: (fn: unknown) => fn,
            wrap: (step: (context: any) => unknown) => {
                const context = { prev: 0, next: 0, done: false, stop() { this.done = true; return undefined; } };
                const iterator = {
                    next() {
                        if (context.done) return { value: undefined, done: true };
                        const value = step(context);
                        return context.done ? { value: undefined, done: true } : { value, done: false };
                    },
                    [Symbol.iterator]() { return iterator; },
                };
                return iterator;
            },
        };
    }
    let returned = false;
    return () => {
        if (returned) return;
        returned = true;
        if (--generatorRuntimeUsers === 0) delete scope.regeneratorRuntime;
    };
}

/** Lazily loads `pdf-lib`, throwing a typed, actionable error when it is not installed. */
async function loadPdfLib(config: OfficeParserAST['config']): Promise<any> {
    try {
        // @ts-ignore - optional peer dependency, resolved at runtime
        return await import('pdf-lib');
    } catch (e) {
        throw getOfficeError(OfficeErrorType.PDF_GENERATION_FAILED, config,
            "The native PDF engine requires the optional peer dependency 'pdf-lib'. Install it with `npm install pdf-lib`, or use the default engine (pdfConfig.engine: 'html').");
    }
}

/** Resolves a numeric length from a pdf.js-style value that may be a number (px) or a unit string. */
function toPoints(v: string | number | undefined, fallback: number): number {
    if (typeof v === 'number' && Number.isFinite(v)) return v * 0.75; // css px -> pt
    if (typeof v === 'string') {
        const m = /^([\d.]+)\s*(pt|px|in|cm|mm)?$/i.exec(v.trim());
        if (m) {
            const n = parseFloat(m[1]);
            switch (m[2]?.toLowerCase()) {
                case 'in': return n * 72;
                case 'cm': return n * 28.3465;
                case 'mm': return n * 2.83465;
                case 'pt': return n;
                default: return n * 0.75; // px
            }
        }
    }
    return fallback;
}

/** The styles text is drawn in. */
type FaceKey = 'regular' | 'bold' | 'italic' | 'boldItalic' | 'mono';

/** A single word ready to place, carrying the style/size/color it should be drawn with. */
interface Piece { text: string; face: FaceKey; size: number; color: any; underline: boolean; }

/** One embedded font and the code points it has a glyph for. */
interface Face { font: any; covers: (cp: number) => boolean; }

/**
 * The fonts each style is drawn in, in the order they are tried for a character: the style's own font,
 * the regular one, the fallback fonts, then the Standard-14 font of the style (which comes first when
 * the style has no font given, see `embedFonts`).
 */
type Fonts = Record<FaceKey, Face[]>;

/** A heading, as the bookmark outline needs it. */
interface OutlineEntry { title: string; level: number; page: any; top: number; }

/**
 * Stateful layout cursor: owns the pdf-lib document, the current page and the y position, and knows
 * how to wrap text, paginate, and paint each node type.
 */
class NativeLayout {
    private page: any;
    /** Distance of the cursor from the top of the page, in points (converted to pdf-lib's y-up on draw). */
    private y = 0;
    /** Resolved image handling, shared with every other generator via {@link resolveImageMode}. */
    private readonly imageMode: ImageMode;
    /** `name -> attachment` index, so image lookups are O(1) rather than a scan per image node. */
    private readonly attachmentsByName = new Map<string, OfficeParserAST['attachments'][number]>();
    /** Each attachment's embedded image, or why it was not embedded: made once, drawn at every picture showing it. */
    private readonly embeddedImages = new Map<OfficeParserAST['attachments'][number], Promise<any>>();
    /** Footnote/endnote bodies gathered from `node.notes` during the walk, drawn at document end. */
    private readonly collectedNotes: OfficeContentNode[] = [];
    /** Sheets drawn so far: each after the first starts a page. */
    private sheetsDrawn = 0;
    private readonly collectedNoteSet = new Set<OfficeContentNode>();

    constructor(
        private readonly pdf: any,
        private readonly lib: any,
        private readonly fonts: Fonts,
        private readonly pageW: number,
        private readonly pageH: number,
        private readonly margin: { top: number; right: number; bottom: number; left: number },
        private readonly config: FullGeneratorConfig,
        private readonly ast: OfficeParserAST,
        // Reports a warning through the generator's own channel, so it lands in `result.messages`
        // (and reaches onWarning) rather than only firing onWarning directly.
        private readonly reportWarning: (type: OfficeWarningType, info?: any) => void,
        /** Whether any font was given: the Standard-14 ones are then the last resort, not the only fonts. */
        private readonly customFonts: boolean,
    ) {
        this.imageMode = resolveImageMode(this.config.includeImages);
        for (const a of this.ast.attachments || []) {
            if (a.name && !this.attachmentsByName.has(a.name)) this.attachmentsByName.set(a.name, a);
        }
        this.newPage();
    }

    private get contentWidth(): number { return this.pageW - this.margin.left - this.margin.right; }
    private get bottom(): number { return this.pageH - this.margin.bottom; }

    /** Starts a fresh page and resets the cursor to the top margin. */
    private newPage(): void {
        this.page = this.pdf.addPage([this.pageW, this.pageH]);
        this.y = this.margin.top;
    }

    /**
     * Ensures `h` points of vertical space remain, starting a new page if not. Never starts a page
     * when the cursor is already at the top margin: the block does not fit a whole page, so a blank
     * page in front of it would not help and would just pad the output. Callers that must not lose
     * content taller than a page (tables) split it across pages themselves.
     */
    private ensureSpace(h: number): void {
        if (this.y + h > this.bottom && this.y > this.margin.top) this.newPage();
    }

    /** Picks the style matching a run's formatting (mono for code, bold/italic variants otherwise). */
    private faceFor(fmt: TextFormatting | undefined, mono: boolean): FaceKey {
        if (mono) return 'mono';
        const bold = !!fmt?.bold, italic = !!fmt?.italic;
        if (bold && italic) return 'boldItalic';
        if (bold) return 'bold';
        if (italic) return 'italic';
        return 'regular';
    }

    /** Parses a `#rrggbb` (or `#rgb`) hex color into a pdf-lib rgb color, or null. */
    private color(hex: string | undefined): any {
        if (!hex || typeof hex !== 'string') return null;
        let h = hex.trim().replace(/^#/, '');
        if (h.length === 3) h = h.split('').map(c => c + c).join('');
        if (!/^[0-9a-f]{6}$/i.test(h)) return null;
        const n = parseInt(h, 16);
        return this.lib.rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255);
    }

    /** Whether the embed-as-link warning was given (once per document). */
    private embedWarned = false;
    /** Whether the chart-as-text warning was given (once per document). */
    private chartWarned = false;
    /** How many characters no font could draw, and the first few distinct ones. */
    private lostCharacters = 0;
    private readonly lostSample: string[] = [];
    /** The headings drawn, for the bookmark outline. */
    readonly outlineEntries: OutlineEntry[] = [];

    /**
     * Makes text drawable in a style before it is measured or drawn: a tab is four spaces, a line break
     * or a space the fonts lack a plain space (a drawn string cannot hold a line break), and any other
     * character none of the style's fonts has is a `?`, counted for the one warning
     * {@link reportLostCharacters} gives, unless `count` is false (text measured, not drawn).
     */
    private enc(text: string, key: FaceKey = 'regular', count = true): string {
        const chain = this.fonts[key];
        let out = '';
        for (const ch of text) {
            const cp = ch.codePointAt(0)!;
            if (cp === 0x09) out += '    ';
            else if (cp === 0x0A || cp === 0x0D) out += ' ';
            else if (cp === 0xAD) { /* a soft hyphen shows only where a line is broken at it, and none is */ }
            else if (chain.some(face => face.covers(cp))) out += ch;
            else if (isInvisible(cp)) { /* draws nothing: no loss */ }
            // A space of another width (an em space, a thin space, an ideographic space) or a line
            // separator the fonts lack is a space; any other control character draws nothing.
            else if (OTHER_SPACE.test(ch)) out += ' ';
            else if (CONTROL.test(ch)) { /* no glyph: no loss */ }
            else {
                out += '?';
                if (!count) continue;
                this.lostCharacters++;
                if (this.lostSample.length < 8 && !this.lostSample.includes(ch)) this.lostSample.push(ch);
            }
        }
        return out;
    }

    /** Warns, once, of the characters no font could draw: how many, and the first few distinct ones. */
    reportLostCharacters(): void {
        if (!this.lostCharacters) return;
        const feature = this.customFonts ? 'characters none of the given fonts has (native engine)' : 'non-WinAnsi characters (native engine)';
        this.reportWarning(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'pdf', feature, reason: 'unencodable-text', count: this.lostCharacters, sample: this.lostSample.join('') });
    }

    /**
     * Splits drawable text into stretches of one font each: every character goes to the first font of
     * the style that has it. A combining mark or a joiner stays with the font of the character before
     * it when that font has it, since a mark drawn in another font is not attached to its letter.
     */
    private segments(text: string, key: FaceKey): { text: string; face: Face }[] {
        const chain = this.fonts[key];
        const out: { text: string; face: Face }[] = [];
        let current: { text: string; face: Face; length: number } | undefined;
        for (const ch of text) {
            const cp = ch.codePointAt(0)!;
            const attached = isInvisible(cp) || COMBINING_MARK.test(ch);
            let face: Face | undefined;
            if (chain.length === 1) face = chain[0];
            else if (current && attached && current.face.covers(cp)) face = current.face;
            else face = chain.find(f => f.covers(cp));
            if (!face) { if (isInvisible(cp)) continue; face = chain[chain.length - 1]; }
            const drawn = chain.length === 1 || face.covers(cp) ? ch : '?';
            // A stretch is shaped as one string, in time that grows with the square of its length on a run
            // of marks (fontkit), so it is cut at MAX_SHAPED code points: before a character that starts a
            // cluster, or, in a cluster longer than that (a letter under thousands of marks), anywhere.
            const full = current !== undefined && current.length >= MAX_SHAPED && (!attached || current.length >= 2 * MAX_SHAPED);
            if (current && current.face === face && !full) { current.text += drawn; current.length++; }
            else { current = { text: drawn, face, length: 1 }; out.push(current); }
        }
        return out;
    }

    /** The width of drawable text in a style, each stretch measured in the font it is drawn in. */
    private widthOf(text: string, key: FaceKey, size: number): number {
        let w = 0;
        for (const seg of this.segments(text, key)) w += seg.face.font.widthOfTextAtSize(seg.text, size);
        return w;
    }

    /** Draws drawable text with its left edge at `x`, each stretch in its font. A right-to-left word's stretches are drawn last first. */
    private drawText(text: string, key: FaceKey, x: number, y: number, size: number, color: any): void {
        const segs = this.segments(text, key);
        if (segs.length > 1 && directionOf(text) === 'rtl') segs.reverse();
        for (const seg of segs) {
            this.page.drawText(seg.text, { x, y, size, font: seg.face.font, color });
            x += seg.face.font.widthOfTextAtSize(seg.text, size);
        }
    }

    /**
     * Draws one line of drawable text (a table cell's line, a line of code) inside `[x, x + width]`,
     * its words in the order they are seen: a right-to-left line starts at the right edge.
     */
    private drawLine(text: string, key: FaceKey, x: number, y: number, size: number, color: any, width: number): void {
        const dir = directionOf(text);
        if (dir !== 'rtl') { this.drawText(text, key, x, y, size, color); return; }
        const baseRtl = startsRtl(text);
        const items = text.split(/(\s+)/).filter(t => t !== '').map(t => ({ text: t, w: this.widthOf(t, key, size), dir: /^\s+$/.test(t) ? 'neutral' as const : directionOf(t) }));
        let cx = baseRtl ? x + Math.max(0, width - items.reduce((s, it) => s + it.w, 0)) : x;
        for (const it of visualOrder(items, baseRtl)) {
            if (it.text.trim()) this.drawText(it.text, key, cx, y, size, color);
            cx += it.w;
        }
    }

    /**
     * Flattens a node into inline text runs (text + formatting + link flag), recursing into children.
     *
     * The caller's `onNode` hook is applied to every descendant visited here - a text run is a node
     * like any other, and `CommonGeneratorConfig.onNode` promises every node - so `false` drops the run
     * (and its subtree) and a returned string replaces its text. The root `node` itself is NOT hooked:
     * whoever passes it here (render(), or the table/cell walk) has already applied the hook to it, so
     * hooking it again would fire twice for the same node.
     */
    private async collectRuns(node: OfficeContentNode): Promise<{ text: string; fmt: TextFormatting; link: boolean }[]> {
        const out: { text: string; fmt: TextFormatting; link: boolean }[] = [];
        const walk = async (n: OfficeContentNode, inherited: TextFormatting, isRoot: boolean): Promise<void> => {
            let text = n.text;
            let override: string | false | void = undefined;
            if (!isRoot) {
                override = await this.onNodeValue(n);
                if (override === false) return;
                if (typeof override === 'string') text = override;
            }
            const fmt = { ...inherited, ...(n.formatting || {}) };
            const link = !!(n.metadata as any)?.link;
            // A picture in a line the engine draws as text (a list item, heading or note; a paragraph
            // draws its pictures itself) is its alt or recognized text: it was dropped.
            if (n.type === 'image' && !isRoot) {
                const alt = this.inlineImageText(n, typeof override === 'string' ? override : undefined);
                if (alt) out.push({ text: alt, fmt, link });
                return;
            }
            // A break inside a line (`<br>`, Word's `w:br`, a column break) starts a new line, and a page
            // break inside a paragraph a new page; they were dropped, gluing the lines together.
            if (n.type === 'break' && !isRoot && typeof override !== 'string') {
                const kind = (n.metadata as any)?.breakType;
                if (kind === 'page') out.push({ text: '\f', fmt, link: false });
                else if (kind !== 'lastRenderedPage' && kind !== 'thematic') out.push({ text: '\n', fmt, link: false });
                return;
            }
            if (n.type === 'text' || (!n.children?.length && text)) {
                if (text) out.push({ text, fmt, link });
                return;
            }
            for (const c of n.children || []) await walk(c, fmt, false);
        };
        await walk(node, {}, true);
        if (!out.length && node.text) out.push({ text: node.text, fmt: node.formatting || {}, link: false });
        return out;
    }

    /**
     * The plain string a table cell (or any node the engine paints as flat text rather than as a block)
     * renders as. Cells built by the markup parsers carry their content only in `children` with no
     * `text` of their own, so reading `cell.text` alone painted every Markdown/HTML table empty.
     * Block children each own a line, which `wrapPlain` keeps as separate lines.
     */
    private async flatText(node: OfficeContentNode): Promise<string> {
        const children = node.children || [];
        // Inline-only children (the common Markdown cell) flatten to a single line; block children
        // (paragraphs, headings, lists, a nested table's rows) are each a line of their own. An image
        // in a cell counts too: the engine cannot embed a picture inside a cell, so it degrades to the
        // image's alt (or OCR) text rather than vanishing silently.
        if (children.some(c => BLOCK_CELL_TYPES.has(c.type) || c.type === 'image')) {
            const lines: string[] = [];
            for (const c of children) {
                // Each block child is a node in its own right, so it gets the hook here (collectRuns
                // only hooks the descendants of whatever root it is handed).
                const override = await this.onNodeValue(c);
                if (override === false) continue;
                let t: string;
                if (typeof override === 'string') t = override;
                else if (c.type === 'image') t = this.cellImageText(c);
                else t = (await this.collectRuns(c)).map(r => r.text).join('').replace(/\f/g, '\n');
                if (t !== '') lines.push(t);
            }
            if (lines.length) return lines.join('\n');
        }
        // A page break inside a cell is a line break there: a cell cannot start a page.
        const flat = (await this.collectRuns(node)).map(r => r.text).join('').replace(/\f/g, '\n');
        return flat || node.text || '';
    }

    private inlineImageWarned = false;
    /**
     * The text a picture in a line of text degrades to, as the image mode has it (none for 'none', the
     * recognized text for 'ocr-text-only', else the alt text or the recognized text), warning once that
     * the native engine draws such a picture as text. `override` is the onNode hook's replacement.
     */
    private inlineImageText(node: OfficeContentNode, override?: string): string {
        if (override !== undefined) return override;
        if (this.imageMode === 'none') return '';
        const ocr = (node.text || '').trim();
        const text = this.imageMode === 'ocr-text-only' ? ocr : (((node.metadata as any)?.altText || '').trim() || ocr);
        if (text && !this.inlineImageWarned) {
            this.inlineImageWarned = true;
            this.reportWarning(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'pdf', feature: 'image inside a list item, heading or note (native engine renders its alt/OCR text)', reason: 'image-in-text-line' });
        }
        return text;
    }

    private cellImageWarned = false;
    /** The text an image in a table cell degrades to (alt text, else OCR text), warning once that the
     * native engine cannot embed a picture inside a cell. */
    private cellImageText(node: OfficeContentNode): string {
        if (!this.cellImageWarned) {
            this.cellImageWarned = true;
            this.reportWarning(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'pdf', feature: 'image inside a table cell (native engine renders its alt/OCR text)', reason: 'image-in-table-cell' });
        }
        return ((node.metadata as any)?.altText || node.text || '').trim();
    }

    /**
     * Wraps a set of inline runs into lines within `[left, left+width]` and draws them, paginating as
     * needed. Returns nothing; advances the cursor to just below the last line.
     */
    private drawRuns(runs: { text: string; fmt: TextFormatting; link: boolean }[], left: number, width: number, mono = false): void {
        // Expand runs into per-word pieces (a trailing-space flag lets us re-insert inter-word gaps).
        // A piece is a word, or (`brk`) a forced break: a new line, or a new page.
        const pieces: { p: Piece; space: boolean; brk?: 'line' | 'page' }[] = [];
        // A paragraph whose first letter is a right-to-left one is read from the right edge.
        const baseRtl = startsRtl(runs.map(r => r.text).join(''));
        for (const run of runs) {
            const face = this.faceFor(run.fmt, mono);
            const size = parseFontSize(run.fmt.size) ?? (mono ? 10 : 11);
            const color = run.link ? this.lib.rgb(0.1, 0.32, 0.72) : (this.color(run.fmt.color) ?? this.lib.rgb(0.12, 0.12, 0.12));
            const underline = !!(run.fmt.underline || run.link);
            // A newline in a run's text is a line break, as the other writers have it (DOCX writes
            // `<w:br/>`): split there before encoding, which turns control characters into spaces.
            run.text.split(/(\r\n|[\n\r\f])/).forEach((segment, at) => {
                if (at % 2) { pieces.push({ p: { text: '', face, size, color, underline: false }, space: false, brk: segment === '\f' ? 'page' : 'line' }); return; }
                for (const w of this.enc(segment, face).split(/(\s+)/)) {
                    if (w === '') continue;
                    if (/^\s+$/.test(w)) { const last = pieces[pieces.length - 1]; if (last && !last.brk) last.space = true; continue; }
                    // Split a word wider than the whole line at character boundaries so it wraps instead
                    // of running off the page (a long URL, a hash, a base64 token).
                    for (const chunk of splitToWidth(w, t => this.widthOf(t, face, size), width)) {
                        pieces.push({ p: { text: chunk, face, size, color, underline }, space: false });
                    }
                }
            });
        }
        // Breaks at the very end draw nothing, as a browser shows a `<br>` closing a paragraph.
        while (pieces.length && pieces[pieces.length - 1].brk) pieces.pop();
        if (!pieces.length) return;

        const spaceWidth = (p: Piece) => this.widthOf(' ', p.face, p.size);
        let line: { p: Piece; space: boolean; w: number }[] = [];
        let lineWidth = 0;

        const flush = () => {
            if (!line.length) return;
            const lh = Math.max(...line.map(it => it.p.size)) * 1.35;
            this.ensureSpace(lh);
            const baseline = this.pageH - this.y - Math.max(...line.map(it => it.p.size));
            // Words and the gaps between them, in reading order, then in the order they are seen.
            const items: { p?: Piece; w: number; dir: 'ltr' | 'rtl' | 'neutral' }[] = [];
            line.forEach((it, i) => {
                items.push({ p: it.p, w: it.w, dir: directionOf(it.p.text) });
                if (it.space && i < line.length - 1) items.push({ w: spaceWidth(it.p), dir: 'neutral' });
            });
            let x = baseRtl ? left + Math.max(0, width - items.reduce((s, it) => s + it.w, 0)) : left;
            for (const it of visualOrder(items, baseRtl)) {
                if (it.p) {
                    this.drawText(it.p.text, it.p.face, x, baseline, it.p.size, it.p.color);
                    if (it.p.underline) {
                        this.page.drawLine({ start: { x, y: baseline - 1.5 }, end: { x: x + it.w, y: baseline - 1.5 }, thickness: 0.5, color: it.p.color });
                    }
                }
                x += it.w;
            }
            this.y += lh;
            line = [];
            lineWidth = 0;
        };

        for (const piece of pieces) {
            if (piece.brk) {
                // A line ended here: the next starts below it. Two in a row leave an empty line.
                if (line.length) flush();
                else { const lh = piece.p.size * 1.35; this.ensureSpace(lh); this.y += lh; }
                if (piece.brk === 'page') this.newPage();
                continue;
            }
            const w = this.widthOf(piece.p.text, piece.p.face, piece.p.size);
            // Only the previous piece's OWN trailing space counts toward width: a word split across two
            // runs (or a splitToWidth chunk) has space:false and draws with no gap, so measuring one
            // there wrapped such lines slightly early and disagreed with what flush() draws.
            const prev = line[line.length - 1];
            if (prev && lineWidth + (prev.space ? spaceWidth(prev.p) : 0) + w > width) flush();
            const prevAfter = line[line.length - 1];
            lineWidth += (prevAfter && prevAfter.space ? spaceWidth(prevAfter.p) : 0) + w;
            line.push({ p: piece.p, space: piece.space, w });
        }
        flush();
    }

    /**
     * The caller's `onNode` verdict for one node, normalized to `false` (drop it) / a replacement
     * string / undefined (render normally). Used for the nodes the engine paints as flat text - text
     * runs, table rows and cells - where a returned string is spliced in as text rather than drawn as
     * its own paragraph, so no "not representable" warning is owed.
     */
    private async onNodeValue(node: OfficeContentNode): Promise<false | string | undefined> {
        if (!this.config.onNode) return undefined;
        const override = await this.config.onNode(node);
        if (override === false) return false;
        return typeof override === 'string' ? override : undefined;
    }

    private onNodeWarned = false;
    /**
     * Applies the caller's `onNode` hook, matching {@link BaseGenerator.handleOnNode}: `false` skips the
     * node and its subtree; a returned string replaces the node's output. The native engine paints a
     * layout rather than emitting markup, so it cannot splice a returned string in structurally - it
     * draws it as a plain-text paragraph (warning once) so the override is honored rather than lost.
     * Returns true when the hook handled the node and default rendering should be skipped.
     */
    private async applyOnNode(node: OfficeContentNode): Promise<boolean> {
        if (!this.config.onNode) return false;
        const override = await this.config.onNode(node);
        if (override === false) return true;
        if (typeof override === 'string') {
            if (!this.onNodeWarned) {
                this.onNodeWarned = true;
                const info = { format: 'pdf', feature: 'onNode string override (native engine draws it as plain text)', reason: 'onnode-string-override' };
                this.reportWarning(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, info);
            }
            if (override) this.drawRuns([{ text: override, fmt: {}, link: false }], this.margin.left, this.contentWidth);
            this.y += 6;
            return true;
        }
        return false;
    }

    private prevPaginated: string | null = null;
    /** Entry point: renders any node, dispatching by type. `topLevel` drives inter-page pagination. */
    async render(node: OfficeContentNode, topLevel = false): Promise<void> {
        if (this.config.abortSignal?.aborted) throw getAbortError();
        // Preserve source pagination: start a fresh page between consecutive page (or slide) nodes.
        // Done BEFORE the hook so a page/slide the hook drops or replaces with a string still advances
        // the bookkeeping (the next page gets its break, and a replacement string lands on a fresh page)
        // instead of leaving prevPaginated stale.
        if (topLevel && (node.type === 'page' || node.type === 'slide')) {
            if (this.prevPaginated === node.type) this.newPage();
            this.prevPaginated = node.type;
        } else if (topLevel) {
            this.prevPaginated = null;
        }
        // The caller's onNode hook can skip a node/subtree or replace its output, exactly as the
        // text-based generators honor it; applied before any default rendering.
        if (await this.applyOnNode(node)) return;
        switch (node.type) {
            case 'page':
            case 'slide':
                for (const c of node.children || []) await this.render(c);
                // Speaker notes, when the PDF keeps them (pdfConfig.printSpeakerNotes), under their slide.
                for (const n of node.notes || []) await this.note(n);
                this.y += 4;
                return;
            case 'header':
            case 'footer':
                for (const c of node.children || []) await this.render(c);
                this.y += 4;
                return;
            // A spreadsheet is a grid: its rows/cells carry `col`, so lay it out as a table rather than
            // letting each cell fall through to `default` and render as its own stacked paragraph. Its
            // drawing images/charts sit as non-row children after the rows; render them after the grid
            // (as the HTML/DOCX/ODT generators do) rather than dropping them.
            case 'sheet': {
                // Each sheet on pages of its own, under its name when the workbook has more than one, as a
                // workbook prints: they ran on one after another with nothing between them.
                if (this.sheetsDrawn++ > 0 && this.y > this.margin.top) this.newPage();
                const sheetName = (node.metadata as any)?.sheetName;
                if (sheetName && this.ast.content.filter(n => n.type === 'sheet').length > 1) {
                    await this.heading({ type: 'heading', metadata: { level: 2 }, children: [{ type: 'text', text: String(sheetName) }] } as OfficeContentNode);
                }
                await this.table(node);
                for (const c of (node.children || [])) if (c.type !== 'row') await this.render(c);
                return;
            }
            case 'heading': return this.heading(node);
            case 'paragraph': return this.paragraph(node);
            case 'list': return this.listItem(node);
            case 'table': return this.table(node);
            case 'image': return this.image(node);
            case 'code': return this.code(node);
            case 'note': return this.note(node);
            case 'break': return this.breakNode(node);
            case 'embed': {
                // A video or framed page cannot play in a PDF: its URL is written, styled as a link (with
                // its label), as the DOCX and ODT generators write it. It was dropped.
                const meta = node.metadata as EmbedMetadata | undefined;
                const url = embedUrl(meta);
                const text = [meta?.label, url].filter(Boolean).join(': ');
                if (!text) return;
                if (!this.embedWarned) {
                    this.embedWarned = true;
                    this.reportWarning(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'pdf', feature: 'embed', reason: 'embed' });
                }
                this.drawRuns([{ text, fmt: {}, link: !!url }], this.margin.left, this.contentWidth);
                this.y += 6;
                return;
            }
            case 'chart':
                // includeCharts: false omits charts in every generator. The native engine cannot draw a
                // chart, so it otherwise renders the chart's data text (in node.text); skip it when off.
                if (this.config.includeCharts === false) return;
                if (node.text) {
                    if (!this.chartWarned) {
                        this.chartWarned = true;
                        this.reportWarning(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'pdf', feature: 'chart (native engine renders its data as text)', reason: 'chart' });
                    }
                    await this.paragraph(node);
                }
                return;
            default:
                if (node.children?.length) { for (const c of node.children) await this.render(c); }
                else if (node.text) await this.paragraph(node);
                return;
        }
    }

    private async heading(node: OfficeContentNode): Promise<void> {
        const level = Math.min(6, Math.max(1, (node.metadata as any)?.level || 1));
        const size = [24, 20, 16, 14, 12, 11][level - 1];
        // Keep-with-next: a heading should not be stranded as the last thing on a page. Reserve its own
        // line plus a couple of body lines, so a heading near the bottom moves to the next page with the
        // text it titles (the HTML engine expresses the same intent with `page-break-after: avoid`).
        // ensureSpace is a no-op at the top margin, so this never inserts a blank page before a heading.
        this.ensureSpace(size * 1.35 + 11 * 1.35 * 2);
        // Space-before, but not when the heading is already at the very top of a page (including one it
        // was just pushed onto by the reservation above): a gap there would float it below the margin.
        if (this.y > this.margin.top) this.y += size * 0.6;
        const runs = (await this.collectRuns(node)).map(r => ({ ...r, fmt: { ...r.fmt, bold: true, size: `${size}pt` } }));
        const title = (runs.length ? runs.map(r => r.text).join('') : node.text || '').replace(/\s+/g, ' ').trim();
        if (title) this.outlineEntries.push({ title, level, page: this.page, top: this.pageH - this.y });
        this.drawRuns(runs.length ? runs : [{ text: node.text || '', fmt: { bold: true, size: `${size}pt` }, link: false }], this.margin.left, this.contentWidth);
        this.y += size * 0.35;
    }

    private async paragraph(node: OfficeContentNode, indentLeft = 0): Promise<void> {
        // A picture in the paragraph (where Word, ODF and Markdown put one) is drawn where it stands,
        // between the lines of text around it: runs carry text only, and it was dropped.
        if (node.children?.some(child => child.type === 'image')) {
            let run: OfficeContentNode[] = [];
            const flush = async () => {
                const runs = run.length ? await this.collectRuns({ ...node, text: undefined, children: run }) : [];
                if (runs.some(r => r.text.trim())) {
                    this.drawRuns(runs, this.margin.left + indentLeft, this.contentWidth - indentLeft);
                    this.y += 6;
                }
                run = [];
            };
            for (const child of node.children) {
                if (child.type !== 'image') { run.push(child); continue; }
                await flush();
                await this.render(child);
            }
            await flush();
            return;
        }
        const runs = await this.collectRuns(node);
        if (!runs.length) { this.y += 6; return; }
        this.drawRuns(runs, this.margin.left + indentLeft, this.contentWidth - indentLeft);
        this.y += 6;
    }

    /** Draws one already-known plain string as a paragraph (an image's alt text, a note marker). */
    private plainParagraph(text: string): void {
        if (!text) return;
        this.drawRuns([{ text, fmt: {}, link: false }], this.margin.left, this.contentWidth);
        this.y += 6;
    }

    private async listItem(node: OfficeContentNode): Promise<void> {
        const meta = node.metadata as any;
        const level = Math.max(0, meta?.indentation || 0);
        const indent = 18 + level * 18;
        const ordered = meta?.listType === 'ordered';
        const marker = ordered ? `${(meta?.itemIndex ?? 0) + 1}.` : '•';
        // Body: the item's own text runs (its non-list children); nested list children render after.
        const bodyRuns = await this.collectRuns({ ...node, children: (node.children || []).filter(c => c.type !== 'list') });
        // Size the marker to the body text so their baselines line up (a body larger than the old fixed
        // 11pt would otherwise sit below its own marker), and reserve one line's height BEFORE drawing
        // the marker so a body that has to break to the next page carries its marker along instead of
        // orphaning it on this page.
        const bodySize = bodyRuns.length ? Math.max(...bodyRuns.map(r => parseFontSize(r.fmt.size) ?? 11)) : 11;
        const size = bodySize > 0 ? bodySize : 11;
        this.ensureSpace(size * 1.35);
        // Draw the marker, then the item body hanging-indented past it. Space is already reserved, so
        // drawRuns' first line stays on this page and draws at the same baseline as the marker.
        const markerX = this.margin.left + indent - 14;
        const baseline = this.pageH - this.y - size;
        this.drawText(this.enc(marker), 'regular', markerX, baseline, size, this.lib.rgb(0.12, 0.12, 0.12));
        // Advance for the marker line unless the body has real text to draw: a whitespace-only body
        // makes drawRuns place nothing and never advance `this.y`, so the next item would overprint.
        if (bodyRuns.some(r => r.text && r.text.trim())) this.drawRuns(bodyRuns, this.margin.left + indent, this.contentWidth - indent);
        else this.y += size * 1.35;
        // Route nested list items through render() (not listItem() directly) so the onNode hook and the
        // abort check see them like every other node.
        for (const c of (node.children || []).filter(c => c.type === 'list')) await this.render(c);
    }

    private code(node: OfficeContentNode): void {
        this.y += 2;
        this.drawMonospace(String(node.text || ''));
        this.y += 6;
    }

    /**
     * Draws text line-by-line in a monospace font, preserving leading indentation and column alignment
     * (word wrapping would collapse both). Used for code blocks and for OCR text, whose 2-D page layout
     * a flowing paragraph would destroy. Split BEFORE encoding: `enc` maps every newline to a space
     * (a drawn string cannot contain one), so encoding first would collapse the block onto one line.
     */
    private drawMonospace(text: string): void {
        const size = 10, lh = size * 1.35;
        for (const raw of text.split(/\r\n|\r|\n/)) {
            const expanded = this.enc(raw.replace(/\t/g, '    '), 'mono');
            const chunks = splitToWidth(expanded || ' ', t => this.widthOf(t, 'mono', size), this.contentWidth - 12);
            for (const chunk of (chunks.length ? chunks : [' '])) {
                this.ensureSpace(lh);
                this.drawLine(chunk, 'mono', this.margin.left + 6, this.pageH - this.y - size, size, this.lib.rgb(0.12, 0.12, 0.12), this.contentWidth - 12);
                this.y += lh;
            }
        }
    }

    /** Draws OCR text: multi-line keeps its 2-D layout in monospace, a single line flows as prose. */
    private async drawOcrText(text: string): Promise<void> {
        if (/[\r\n]/.test(text)) { this.y += 2; this.drawMonospace(text); this.y += 4; }
        else await this.plainParagraph(text);
    }

    private async note(node: OfficeContentNode): Promise<void> {
        const runs = (await this.collectRuns(node)).map(r => ({ ...r, fmt: { ...r.fmt, size: '9pt' } }));
        if (runs.length) this.drawRuns(runs, this.margin.left + 12, this.contentWidth - 12);
        this.y += 4;
    }

    private breakNode(node: OfficeContentNode): void {
        const t = (node.metadata as any)?.breakType;
        if (t === 'page') { this.newPage(); return; }
        if (t === 'thematic') {
            this.ensureSpace(12);
            this.y += 6;
            const yy = this.pageH - this.y;
            this.page.drawLine({ start: { x: this.margin.left, y: yy }, end: { x: this.pageW - this.margin.right, y: yy }, thickness: 0.5, color: this.lib.rgb(0.7, 0.7, 0.7) });
            this.y += 6;
        } else {
            this.y += 8;
        }
    }

    private async image(node: OfficeContentNode): Promise<void> {
        // Honor the same includeImages / imageMode contract as every other generator (HTML, DOCX, ODT):
        // 'none' drops the node entirely; 'ocr-text-only' emits only the recognized text; 'image-only'
        // and 'image+ocr-text' embed the picture, the latter also drawing its OCR text underneath.
        const mode = this.imageMode;
        if (mode === 'none') return;
        const ocr = (node.text || '').trim();
        if (mode === 'ocr-text-only') { if (ocr) await this.drawOcrText(ocr); return; }

        const meta = node.metadata as any;
        const name = meta?.attachmentName;
        // Nothing renderable: keep the alt text (or the OCR text) rather than drawing nothing, and
        // report the failure - the DOCX and ODT generators degrade and warn for the same input, and a
        // picture that silently vanishes gives the caller no way to find out.
        const fallback = () => this.plainParagraph(meta?.altText || ocr);
        const attachment = name ? this.attachmentsByName.get(name) : undefined;
        if (!attachment?.data) {
            if (name) this.reportWarning(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name, reason: attachment ? 'attachment has no data' : 'missing attachment' });
            fallback();
            return;
        }
        try {
            // Each attachment decoded and embedded once, and drawn wherever a picture shows it: embedded
            // again at each, one 3 MB picture shown 400 times (a 3 MB DOCX) made a 1.2 GB PDF.
            let embedded = this.embeddedImages.get(attachment);
            if (!embedded) {
                embedded = (async () => {
                    const bytes = base64ToBytes(attachment.data);
                    // Reject an image whose declared dimensions are absurd BEFORE decoding it: embedPng/embedJpg
                    // allocate the full bitmap, so a decompression bomb would OOM the process uncatchably.
                    const dim = sniffImageSize(bytes);
                    if (dim && dim.w * dim.h > MAX_IMAGE_PIXELS) return `image is too large to embed (${dim.w}x${dim.h} pixels)`;
                    const isJpg = /jpe?g/i.test(attachment.extension || '') || attachment.mimeType === 'image/jpeg';
                    if (isJpg) return await this.pdf.embedJpg(bytes);
                    // A GIF (its first frame) or an uncompressed BMP, which a PDF cannot hold, as a PNG: both
                    // were dropped. Decoded within the same pixel budget, checked before allocating.
                    const isGif = bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46;
                    const isBmp = bytes[0] === 0x42 && bytes[1] === 0x4D;
                    if (isGif || isBmp) {
                        const raster = isGif ? decodeGif(bytes, MAX_IMAGE_PIXELS) : decodeBmp(bytes, MAX_IMAGE_PIXELS);
                        const png = raster ? encodePng(raster) : null;
                        if (!png) return `unreadable ${isGif ? 'GIF' : 'BMP'} picture`;
                        return await this.pdf.embedPng(png);
                    }
                    return await this.pdf.embedPng(bytes);
                })();
                this.embeddedImages.set(attachment, embedded);
            }
            const img = await embedded;
            if (typeof img === 'string') {
                this.reportWarning(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name, reason: img });
                fallback();
                return;
            }
            // Resolve the draw size in points, by priority: an explicit ImageMetadata.width/height
            // (a length like '200px'/'3cm', a number in px, or a '%' of the content width), then PDF
            // page bounds (already points), then the intrinsic pixel size converted px->pt. Aspect ratio
            // is preserved from the intrinsic image when only a width is given.
            const ratio = img.width ? img.height / img.width : 1;
            const meta2 = node.metadata as any;
            const pctOf = (v: any): number | null => (typeof v === 'string' && /^\s*[\d.]+\s*%\s*$/.test(v)) ? parseFloat(v) / 100 : null;
            let w: number, h: number;
            if (meta2?.width != null) {
                const p = pctOf(meta2.width);
                w = p != null ? this.contentWidth * p : toPoints(meta2.width, img.width * 0.75);
                h = (meta2.height != null && pctOf(meta2.height) == null) ? toPoints(meta2.height, w * ratio) : w * ratio;
            } else if (node.bounds?.width) {
                w = node.bounds.width;
                h = node.bounds.height || w * ratio;
            } else {
                w = img.width * 0.75; // intrinsic pixels -> points (96dpi), not points 1:1
                h = img.height * 0.75;
            }
            const scale = Math.min(1, this.contentWidth / w);
            w *= scale; h *= scale;
            if (h > this.bottom - this.margin.top) { const s = (this.bottom - this.margin.top) / h; w *= s; h *= s; }
            this.ensureSpace(h + 6);
            this.page.drawImage(img, { x: this.margin.left, y: this.pageH - this.y - h, width: w, height: h });
            this.y += h + 6;
        } catch {
            // pdf-lib embeds PNG and JPEG (and a GIF or BMP is turned into PNG), so a WebP, TIFF or SVG
            // attachment lands here.
            this.reportWarning(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name, reason: 'unsupported image format (the native PDF engine embeds PNG, JPEG, GIF and BMP)' });
            fallback();
            return;
        }
        // 'image+ocr-text': draw the recognized text just below the successfully embedded image.
        if (mode === 'image+ocr-text' && ocr) await this.drawOcrText(ocr);
    }

    /**
     * Deep-collects footnote/endnote bodies in document order. Parsers attach `.notes` to the TEXT
     * RUN a marker sits in (e.g. WordParser), and a run is never passed to `render()` (its paragraph
     * renders it via `collectRuns`), so a render()-level sweep misses them. Walk the whole subtree.
     */
    collectNotes(node: OfficeContentNode): void {
        if (!node) return;
        // A slide's notes are its speaker's: drawn under the slide (see render), never at the end of the
        // document, where they printed on the last slide's page.
        if (node.type === 'slide') { for (const c of node.children || []) this.collectNotes(c); return; }
        // The notes a note's own text refers to follow it (they were left out); each note once.
        for (const note of node.notes || []) {
            if (this.collectedNoteSet.has(note)) continue;
            this.collectedNoteSet.add(note);
            this.collectedNotes.push(note);
            this.collectNotes(note);
        }
        for (const c of node.children || []) this.collectNotes(c);
    }

    /** Draws the footnote/endnote bodies gathered during the walk, under a short separating rule. */
    async flushNotes(): Promise<void> {
        if (!this.collectedNotes.length) return;
        this.y += 8;
        this.ensureSpace(14);
        const yy = this.pageH - this.y;
        this.page.drawLine({ start: { x: this.margin.left, y: yy }, end: { x: this.margin.left + Math.min(180, this.contentWidth), y: yy }, thickness: 0.5, color: this.lib.rgb(0.7, 0.7, 0.7) });
        this.y += 8;
        for (const n of this.collectedNotes) await this.note(n);
    }

    private async table(node: OfficeContentNode): Promise<void> {
        const rows = (node.children || []).filter(r => r.type === 'row');
        if (!rows.length) return;
        // Total columns = the widest grid extent any cell reaches (col + colSpan), falling back to
        // the largest raw cell count when cells carry no `col`.
        let cols = 0;
        for (const row of rows) {
            for (const cell of (row.children || []).filter(c => c.type === 'cell')) {
                const col = (cell.metadata as CellMeta)?.col;
                const span = (cell.metadata as CellMeta)?.colSpan || 1;
                cols = Math.max(cols, (typeof col === 'number' ? col : 0) + span);
            }
        }
        cols = Math.max(cols, ...rows.map(r => (r.children || []).filter(c => c.type === 'cell').length), 1);
        const pad = 4;
        const border = this.lib.rgb(0.6, 0.6, 0.6);

        // First the cells' texts and places, then the columns' widths from them, then the rows drawn.
        type Placed = { col: number; span: number; rowSpan: number; text: string; header: boolean; lines: string[] };
        const grid: Placed[][] = [];
        // Grid occupancy: column -> rows still covered by a rowspan from above, so a spanning cell
        // reserves its column(s) below instead of letting later rows slide left under it.
        const carry = new Map<number, number>();
        for (const row of rows) {
            const rowOverride = await this.onNodeValue(row);
            if (rowOverride === false) continue;
            const cells = (row.children || []).filter(c => c.type === 'cell');
            // Header detection via the shared heuristic every other generator uses (row/cell `style`,
            // `isHeader`, or an all-bold first row), so a bold first row from DOCX/Markdown bolds here
            // too - not only PDF's per-cell `style: 'header'`.
            const rowIsHeader = isHeaderRow(row, row === rows[0]);
            // Place each cell into a grid column: honour explicit `col`, otherwise the next free one.
            let cursor = 0;
            const placed: Placed[] = [];
            // The columns this row's earlier cells span: an empty cell there is one a merge covers.
            const spanned = new Set<number>();
            if (typeof rowOverride === 'string') {
                // An onNode string for a whole row: paint it across the full grid width, since the
                // engine has no markup to splice a replacement row into.
                placed.push({ col: 0, span: cols, rowSpan: 1, text: rowOverride, header: rowIsHeader, lines: [] });
            }
            for (const cell of typeof rowOverride === 'string' ? [] : cells) {
                const override = await this.onNodeValue(cell);
                if (override === false) continue;
                const meta = cell.metadata as CellMeta;
                const span = meta?.colSpan && meta.colSpan > 1 ? meta.colSpan : 1;
                let col = typeof meta?.col === 'number' ? meta.col : -1;
                if (col < 0) { while ((carry.get(cursor) || 0) > 0) cursor++; col = cursor; }
                // Cells built by the markup parsers carry their text only in `children`; reading
                // `cell.text` alone rendered every Markdown/HTML table as an empty grid.
                const text = typeof override === 'string' ? override : await this.flatText(cell);
                // An empty cell under another's span (a merged range's covered cells, as XLSX keeps them)
                // is not drawn: its box would cut through the merged one.
                if (!text.trim() && (spanned.has(col) || (carry.get(col) || 0) > 0)) continue;
                const header = rowIsHeader || meta?.style === 'header';
                placed.push({ col, span, rowSpan: Math.max(1, meta?.rowSpan || 1), text, header, lines: [] });
                for (let c = col + 1; c < col + span && c < cols; c++) spanned.add(c);
                cursor = col + span;
            }
            grid.push(placed);
            // Record this row's rowspans, then age the carries by one row. The carry reserves the
            // spanned columns below so later rows do not slide under a rowspan cell, and the cell's text
            // is placed correctly in its first row. Its border box, though, covers only that first row:
            // painting the continuation border correctly (a single tall box, or edges without internal
            // lines, and correct across a page break) needs pre-measured row heights, so a rowspan
            // cell's lower rows are left unbordered rather than mis-bordered.
            const newCarry = new Map<number, number>();
            for (const p of placed) {
                if (p.rowSpan > 1) for (let c = p.col; c < p.col + p.span; c++) newCarry.set(c, p.rowSpan - 1);
            }
            for (const [c, v] of [...carry]) { if (v > 1) carry.set(c, v - 1); else carry.delete(c); }
            for (const [c, v] of newCarry) carry.set(c, v);
        }

        // Each column as wide as its content asks, within the page: never narrower than its longest
        // word (so a number is never broken over two lines), wider for longer text, sharing what is
        // left. A table whose words alone are wider than the page is set smaller, down to 6pt, as a
        // printout fits a wide sheet to the page width; past that a word is broken where it must be.
        // Equal columns broke 104 into "10" over "4", and a narrow column's text a letter at a time.
        const { widths, size } = this.columnWidths(grid, cols, pad);
        const xs = [this.margin.left];
        for (const w of widths) xs.push(xs[xs.length - 1] + w);
        const spanWidth = (col: number, span: number) => xs[Math.min(col + span, cols)] - xs[Math.min(col, cols)];
        /** Wraps a cell's (un-encoded) text to the width of the columns it spans. */
        // A header cell is drawn bold, so it is wrapped at the bold font's widths.
        const wrap = (text: string, col: number, span: number, header = false) => {
            const face: FaceKey = header ? 'bold' : 'regular';
            return wrapPlain(text, t => this.widthOf(t, face, size), Math.max(10, spanWidth(col, span) - 2 * pad), t => this.enc(t, face));
        };

        this.y += 2;
        for (const placed of grid) {
            for (const p of placed) p.lines = wrap(p.text, p.col, p.span, p.header);
            const lineH = size * 1.35;
            const rowH = Math.max(size * 1.4, ...placed.map(p => p.lines.length * lineH)) + 2 * pad;
            // Draw the row in vertical bands so a row taller than a whole page is split across pages
            // rather than drawn off the bottom (which silently lost the overflowing text). A normal row
            // is a single band. Each band redraws every cell's border and its slice of the wrapped
            // lines, so the cell continues cleanly on the next page.
            const fullBand = Math.max(1, Math.floor((this.bottom - this.margin.top - 2 * pad) / lineH));
            const maxLines = Math.max(1, ...placed.map(p => p.lines.length));
            let lineOffset = 0;
            while (lineOffset < maxLines) {
                const remaining = maxLines - lineOffset;
                // Lines that fit below the cursor on the current page. If the whole remaining row fits,
                // draw it in one band; otherwise fill what is left of this page (>=1 line) before moving
                // on, so a tall row starting mid-page does not blank the rest of the page. A band is
                // never taller than a page (fitLines <= fullBand), so no line is ever drawn off-page.
                const fitLines = Math.max(0, Math.floor((this.bottom - this.y - 2 * pad) / lineH));
                let bandLines: number;
                if (fitLines >= remaining) bandLines = remaining;               // the rest fits here
                else if (this.y > this.margin.top && fitLines >= 1) bandLines = fitLines; // fill this page
                else if (this.y > this.margin.top) { this.newPage(); continue; } // nothing fits: fresh page
                else bandLines = Math.min(remaining, fullBand);                  // at page top: a full band
                const whole = lineOffset === 0 && bandLines === maxLines;
                const bandH = whole ? rowH : bandLines * lineH + 2 * pad;
                const topY = this.pageH - this.y;
                for (const p of placed) {
                    const x = xs[Math.min(p.col, cols)];
                    this.page.drawRectangle({ x, y: topY - bandH, width: spanWidth(p.col, p.span), height: bandH, borderColor: border, borderWidth: 0.5, color: undefined });
                    const face: FaceKey = p.header ? 'bold' : 'regular';
                    let ty = topY - pad - size;
                    for (let li = lineOffset; li < lineOffset + bandLines && li < p.lines.length; li++) {
                        this.drawLine(p.lines[li], face, x + pad, ty, size, this.lib.rgb(0.12, 0.12, 0.12), spanWidth(p.col, p.span) - 2 * pad); ty -= lineH;
                    }
                }
                this.y += bandH;
                lineOffset += bandLines;
            }
        }
        this.y += 6;
    }

    /**
     * The widths of a table's columns and the size its text is set in (see table): each column at least
     * as wide as its longest word, at most as wide as its longest line, at 10pt; set smaller, down to 6pt,
     * when its words alone are wider than the page. Measured from the cells that span one column.
     */
    private columnWidths(grid: { col: number; span: number; text: string; header: boolean }[][], cols: number, pad: number): { widths: number[]; size: number } {
        const base = 10, smallest = 6, room = this.contentWidth;
        const least = new Array<number>(cols).fill(2 * pad + 6), most = new Array<number>(cols).fill(2 * pad + 6);
        for (const placed of grid) {
            for (const p of placed) {
                if (p.span !== 1 || p.col >= cols) continue;
                const face: FaceKey = p.header ? 'bold' : 'regular';
                let word = 0, line = 0;
                for (const raw of p.text.split(/\r\n|\r|\n/)) {
                    const encoded = this.enc(raw, face, false);
                    line = Math.max(line, this.widthOf(encoded, face, base));
                    for (const w of encoded.split(/\s+/)) if (w) word = Math.max(word, this.widthOf(w, face, base));
                }
                // A point to spare, so a word as wide as its column is not broken by rounding.
                least[p.col] = Math.max(least[p.col], word + 1 + 2 * pad);
                most[p.col] = Math.max(most[p.col], line + 1 + 2 * pad);
            }
        }
        const sum = (a: number[]) => a.reduce((n, w) => n + w, 0);
        const leastSum = sum(least), mostSum = sum(most);
        if (mostSum <= room) return { widths: most.map(w => w * room / mostSum), size: base };
        if (leastSum <= room) {
            const spare = room - leastSum, give = mostSum - leastSum;
            return { widths: least.map((w, c) => w + (give > 0 ? spare * (most[c] - w) / give : spare / cols)), size: base };
        }
        // The words alone are wider than the page: smaller text, its widths scaled with it.
        const scale = Math.max(smallest / base, (room - 2 * pad * cols) / (leastSum - 2 * pad * cols));
        const scaled = least.map(w => (w - 2 * pad) * scale + 2 * pad);
        const total = sum(scaled);
        return { widths: scaled.map(w => w * room / total), size: base * scale };
    }
}

/** Minimal shape of `CellMetadata` the native engine reads. */
interface CellMeta { col?: number; colSpan?: number; rowSpan?: number; style?: string; }

/**
 * Node types that own a line of their own inside a table cell. A cell holding these is flattened one
 * line per child; a cell holding only inline content (the usual Markdown/HTML cell) flattens to one.
 */
const BLOCK_CELL_TYPES = new Set<string>(['paragraph', 'heading', 'list', 'code', 'table', 'row', 'admonition', 'note', 'page', 'slide', 'sheet']);

/** Parses a `TextFormatting.size` ("12pt", "14") into points, or null. */
function parseFontSize(size: string | undefined): number | null {
    if (size == null) return null;
    const s = String(size).trim();
    // Relative units resolve against the engine's 11pt default body size (HTML/Markdown sources store a
    // raw CSS `font-size`, so `1.2em` or `120%` must not be read as 1.2pt / 120pt).
    const rel = /^([\d.]+)\s*(em|rem|%)$/i.exec(s);
    if (rel) {
        const n = parseFloat(rel[1]);
        const pt = (rel[2].toLowerCase() === '%' ? n / 100 : n) * 11;
        return Number.isFinite(pt) && pt > 0 ? pt : null;
    }
    // An explicit absolute unit goes through the shared converter (px->pt at 96dpi, in/cm/mm too); a
    // bare number stays points, matching what the parsers write into `formatting.size`.
    if (/(pt|px|in|cm|mm)\s*$/i.test(s)) {
        const pt = lengthToPt(s);
        return pt != null && pt > 0 ? pt : null;
    }
    const n = parseFloat(s);
    return Number.isFinite(n) && n > 0 ? n : null;
}

/**
 * Splits a single token into chunks each no wider than `width`, at character boundaries. Used so a
 * word longer than the whole line (a long URL, hash or base64 blob) wraps instead of overflowing the
 * page. Returns the token unchanged when it already fits.
 */
function splitToWidth(text: string, measure: (text: string) => number, width: number): string[] {
    if (width <= 0 || measure(text) <= width) return [text];
    const chunks: string[] = [];
    let cur = '';
    // A running width, each character measured once: measuring the chunk so far again at every
    // character took time quadratic in its length, and zero-width characters (combining marks) never
    // fill a line, so 16 KB of them took seconds. A word broken at character boundaries is unshaped
    // across the break anyway, so the sum of its characters' widths is as exact as the split can be.
    let curWidth = 0;
    for (const ch of text) {
        const chWidth = measure(ch);
        if (cur && curWidth + chWidth > width) { chunks.push(cur); cur = ch; curWidth = chWidth; }
        else { cur += ch; curWidth += chWidth; }
    }
    if (cur) chunks.push(cur);
    return chunks.length ? chunks : [text];
}

/**
 * Greedy word-wrap of plain, UN-ENCODED text to a pixel width, returning the wrapped lines. The text
 * is split into source lines first and encoded one line at a time: `enc` maps every newline to
 * a space (a drawn string cannot contain one), so encoding up front would collapse a multi-line cell
 * onto a single line. `measure` gives the width of encoded text in the font it will be drawn in.
 */
function wrapPlain(text: string, measure: (text: string) => number, width: number, enc: (s: string) => string): string[] {
    const out: string[] = [];
    for (const raw of String(text ?? '').split(/\r\n|\r|\n/)) {
        const para = enc(raw);
        const words = para.split(/\s+/).filter(Boolean);
        let line = '';
        for (const w of words) {
            // A single word wider than the whole column (a URL, hash, long identifier) is broken at
            // character boundaries so it wraps instead of overflowing into the next cell or off the page.
            const pieces = width > 0 && measure(w) > width ? splitToWidth(w, measure, width) : [w];
            for (const piece of pieces) {
                const cand = line ? line + ' ' + piece : piece;
                if (line && measure(cand) > width) { out.push(line); line = piece; }
                else line = cand;
            }
        }
        out.push(line);
    }
    return out.length ? out : [''];
}

/** Decodes a base64 string to bytes in both Node and the browser. */
function base64ToBytes(b64: string): Uint8Array {
    if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(b64, 'base64'));
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
}

/**
 * Writes the document metadata (title, author, dates, keywords) onto the pdf-lib document. `m` is the
 * generator's already-resolved `effectiveMetadata` (overrides merged), so this does not re-merge.
 */
function applyMetadata(pdf: any, m: OfficeMetadata): void {
    const set = (fn: string, v: any) => { try { if (v != null && typeof pdf[fn] === 'function') pdf[fn](v); } catch { /* best effort */ } };
    const asDate = (v: unknown): Date | null => {
        if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
        if (typeof v === 'string' && v) { const d = new Date(v); return isNaN(d.getTime()) ? null : d; }
        return null;
    };
    set('setTitle', m.title);
    set('setAuthor', m.author);
    set('setSubject', m.subject);
    // The document language, as the HTML engine's `<html lang>` gives it (`/Lang`).
    if (m.language || m.nativeProperties?.language) set('setLanguage', documentLanguage(m));
    if (m.keywords) set('setKeywords', String(m.keywords).split(/[,;]\s*/).filter(Boolean));
    set('setCreator', 'officeParser (native engine)');
    set('setProducer', 'officeParser (pdf-lib)');
    const created = asDate(m.created); if (created) set('setCreationDate', created);
    const modified = asDate(m.modified); if (modified) set('setModificationDate', modified);
}

/**
 * A font as a caller gives it: the bytes of a TrueType or OpenType file. Anything else given (a file
 * name, an empty file) is reported and the default font used, rather than passed over without a word.
 */
function fontBytes(value: unknown, name: string, reportWarning: (type: OfficeWarningType, info?: any) => void): Uint8Array | null {
    if (value === undefined || value === null) return null;
    if (value instanceof Uint8Array && value.byteLength) return value;
    if (value instanceof ArrayBuffer && value.byteLength) return new Uint8Array(value);
    const shown = value instanceof Uint8Array || value instanceof ArrayBuffer ? 'an empty file' : typeof value === 'string' ? value : Object.prototype.toString.call(value);
    reportWarning(OfficeWarningType.INVALID_CONFIG_VALUE, { option: `pdfConfig.${name}`, value: shown, expected: 'the bytes of a TrueType or OpenType font file (a Uint8Array or ArrayBuffer)', fallback: undefined });
    return null;
}

/**
 * Embeds the fonts text is drawn in: the Standard-14 Helvetica and Courier faces, and before them the
 * fonts given in `pdfConfig.fonts` and `pdfConfig.fallbackFonts`. A style with no font of its own is
 * drawn in the regular one given, or in its Standard-14 font when that is not given either, the
 * fallback fonts drawing what it lacks; code keeps Courier unless a `mono` font is given. A font file that
 * cannot be read fails the conversion, naming which one: drawing in another font without a word would
 * be a different document.
 */
async function embedFonts(pdf: any, lib: any, pc: FullGeneratorConfig['pdfConfig'], config: OfficeParserAST['config'], reportWarning: (type: OfficeWarningType, info?: any) => void): Promise<{ fonts: Fonts; custom: boolean }> {
    const standard = (font: any): Face => ({ font, covers: winAnsiCovers });
    const std: Record<FaceKey, Face> = {
        regular: standard(await pdf.embedFont(lib.StandardFonts.Helvetica)),
        bold: standard(await pdf.embedFont(lib.StandardFonts.HelveticaBold)),
        italic: standard(await pdf.embedFont(lib.StandardFonts.HelveticaOblique)),
        boldItalic: standard(await pdf.embedFont(lib.StandardFonts.HelveticaBoldOblique)),
        mono: standard(await pdf.embedFont(lib.StandardFonts.Courier)),
    };
    const keys: FaceKey[] = ['regular', 'bold', 'italic', 'boldItalic', 'mono'];
    // `fonts` names a file for each style. Given as anything else (a font's bytes alone, a file name) it is
    // reported and no style has one; a style is read only as the object's own key.
    const fontMap: unknown = pc.fonts;
    const isMap = !!fontMap && typeof fontMap === 'object' && !Array.isArray(fontMap) && !ArrayBuffer.isView(fontMap) && !(fontMap instanceof ArrayBuffer);
    if (!isMap && fontMap !== undefined && fontMap !== null) {
        const shown = typeof fontMap === 'string' ? fontMap : Object.prototype.toString.call(fontMap);
        reportWarning(OfficeWarningType.INVALID_CONFIG_VALUE, { option: 'pdfConfig.fonts', value: shown, expected: 'an object giving a font file for each style it names (regular, bold, italic, boldItalic, mono)', fallback: undefined });
    }
    const styleFont = (key: FaceKey): unknown => isMap && Object.prototype.hasOwnProperty.call(fontMap, key) ? (fontMap as Record<string, unknown>)[key] : undefined;
    const given = keys.map(key => ({ name: `fonts.${key}`, key, bytes: fontBytes(styleFont(key), `fonts.${key}`, reportWarning) }));
    const fallbackList: unknown[] = Array.isArray(pc.fallbackFonts) ? pc.fallbackFonts : [];
    if (!Array.isArray(pc.fallbackFonts) && pc.fallbackFonts !== undefined && pc.fallbackFonts !== null) {
        reportWarning(OfficeWarningType.INVALID_CONFIG_VALUE, { option: 'pdfConfig.fallbackFonts', value: Object.prototype.toString.call(pc.fallbackFonts), expected: 'a list of font files (an array)', fallback: undefined });
    }
    const fallbackGiven = fallbackList.map((value, i) => ({ name: `fallbackFonts[${i}]`, bytes: fontBytes(value, `fallbackFonts[${i}]`, reportWarning) }));
    const custom = given.some(g => g.bytes) || fallbackGiven.some(g => g.bytes);
    if (!custom) {
        return { fonts: { regular: [std.regular], bold: [std.bold], italic: [std.italic], boldItalic: [std.boldItalic], mono: [std.mono] }, custom };
    }

    pdf.registerFontkit(await loadFontkit(config));
    // The same bytes given for two styles are embedded once: the same view, or views of the same part of
    // one buffer (an ArrayBuffer given twice is viewed anew each time).
    const embedded = new Map<ArrayBufferLike, Map<string, Face>>();
    const embed = async (name: string, bytes: Uint8Array | null): Promise<Face | undefined> => {
        if (!bytes) return undefined;
        const span = `${bytes.byteOffset}:${bytes.byteLength}`;
        let inBuffer = embedded.get(bytes.buffer);
        if (!inBuffer) embedded.set(bytes.buffer, inBuffer = new Map());
        let face = inBuffer.get(span);
        if (face) return face;
        try {
            const font = await pdf.embedFont(bytes, { subset: pc.subsetFonts !== false });
            const set = new Set<number>(font.getCharacterSet());
            face = { font, covers: cp => set.has(cp) };
        } catch (e: any) {
            throw getOfficeError(OfficeErrorType.PDF_GENERATION_FAILED, config, `pdfConfig.${name} could not be read as a TrueType or OpenType font (${e?.message || String(e)})`);
        }
        inBuffer.set(span, face);
        return face;
    };
    const own = {} as Record<FaceKey, Face | undefined>;
    for (const g of given) own[g.key] = await embed(g.name, g.bytes);
    const fallbacks: Face[] = [];
    for (const g of fallbackGiven) { const face = await embed(g.name, g.bytes); if (face) fallbacks.push(face); }

    const chain = (key: FaceKey): Face[] => {
        // A style is drawn in its own font, or in the regular one given. With neither (and for code,
        // which stays in a fixed-width font unless a `mono` font is given) the Standard-14 font of the
        // style comes first, and the fonts given draw only what it lacks.
        const primary = key === 'mono' ? own.mono : (own[key] ?? own.regular);
        const order = primary ? [primary, own.regular, ...fallbacks, std[key]] : [std[key], own.regular, ...fallbacks];
        return order.filter((face, i): face is Face => !!face && order.indexOf(face) === i);
    };
    return { fonts: { regular: chain('regular'), bold: chain('bold'), italic: chain('italic'), boldItalic: chain('boldItalic'), mono: chain('mono') }, custom };
}

/**
 * Writes the bookmark outline: one entry per heading, nested by heading level, each pointing at the
 * place on its page where the heading starts.
 */
function writeOutline(pdf: any, lib: any, entries: OutlineEntry[]): void {
    if (!entries.length) return;
    interface Item { entry?: OutlineEntry; ref: any; children: Item[]; }
    const context = pdf.context;
    const root: Item = { ref: context.nextRef(), children: [] };
    const open: { item: Item; level: number }[] = [{ item: root, level: 0 }];
    for (const entry of entries) {
        while (open[open.length - 1].level >= entry.level) open.pop();
        const item: Item = { entry, ref: context.nextRef(), children: [] };
        open[open.length - 1].item.children.push(item);
        open.push({ item, level: entry.level });
    }
    const descendants = (item: Item): number => item.children.reduce((n, child) => n + 1 + descendants(child), 0);
    const write = (parent: Item): void => {
        parent.children.forEach((item, i) => {
            const entry = item.entry!;
            const dict: Record<string, any> = {
                Title: lib.PDFHexString.fromText(entry.title),
                Parent: parent.ref,
                Dest: [entry.page.ref, 'XYZ', null, entry.top, null],
            };
            if (i > 0) dict.Prev = parent.children[i - 1].ref;
            if (i < parent.children.length - 1) dict.Next = parent.children[i + 1].ref;
            if (item.children.length) {
                dict.First = item.children[0].ref;
                dict.Last = item.children[item.children.length - 1].ref;
                dict.Count = descendants(item);
            }
            context.assign(item.ref, context.obj(dict));
            write(item);
        });
    };
    write(root);
    context.assign(root.ref, context.obj({
        Type: 'Outlines',
        First: root.children[0].ref,
        Last: root.children[root.children.length - 1].ref,
        Count: descendants(root),
    }));
    pdf.catalog.set(lib.PDFName.of('Outlines'), root.ref);
}

/**
 * Renders an AST to a PDF using the native pdf-lib layout engine.
 *
 * @returns the PDF as bytes.
 */
export async function renderNativePdf(ast: OfficeParserAST, config: FullGeneratorConfig, metadata: OfficeMetadata, reportWarning: (type: OfficeWarningType, info?: any) => void): Promise<Uint8Array> {
    const lib = await loadPdfLib(ast.config);
    // updateMetadata:false stops pdf-lib stamping CreationDate/ModDate with the current time: dates
    // come only from the document's own metadata (via applyMetadata), so a source with no dates
    // produces byte-identical output on every run, matching the reproducible DOCX/ODT/EPUB generators.
    const pdf = await lib.PDFDocument.create({ updateMetadata: false });
    applyMetadata(pdf, metadata);

    const pc = config.pdfConfig;
    const { fonts, custom: customFonts } = await embedFonts(pdf, lib, pc, ast.config, reportWarning);
    const returnGeneratorRuntime = customFonts ? borrowGeneratorRuntime() : () => { /* nothing lent */ };
    try {
        return await drawDocument(pdf, lib, fonts, customFonts, ast, config, reportWarning);
    } finally {
        returnGeneratorRuntime();
    }
}

/** Lays the document out on the pages of `pdf` and returns its bytes. */
async function drawDocument(pdf: any, lib: any, fonts: Fonts, customFonts: boolean, ast: OfficeParserAST, config: FullGeneratorConfig, reportWarning: (type: OfficeWarningType, info?: any) => void): Promise<Uint8Array> {
    const pc = config.pdfConfig;
    const paper = paperSizePt(pc.format);
    // Orient the FORMAT by swapping its dimensions when landscape (the same `landscape ? h : w` the
    // DOCX/ODT generators and Puppeteer use), so a paper that is already landscape - `ledger` (17x11) -
    // rotates to portrait under `landscape: true` instead of the old conditional swap that never
    // rotated a landscape sheet. Explicit width/height then override, and (like Puppeteer) are taken as
    // given rather than re-swapped by the flag.
    let w = pc.landscape ? paper.h : paper.w;
    let h = pc.landscape ? paper.w : paper.h;
    if (pc.width) w = toPoints(pc.width, w);
    if (pc.height) h = toPoints(pc.height, h);

    // Margin sides default to the '' unset sentinel; toPoints maps that (and any unparseable value) to
    // mdef so text is not glued to the sheet edge, while an explicit 0 parses to 0 and is honoured for
    // callers who genuinely want an edge-to-edge native PDF.
    const mdef = 48; // ~0.67in
    const mv = (v: string | number | undefined) => toPoints(v, mdef);
    const margin = { top: mv(pc.margin?.top), right: mv(pc.margin?.right), bottom: mv(pc.margin?.bottom), left: mv(pc.margin?.left) };

    const layout = new NativeLayout(pdf, lib, fonts, w, h, margin, config, ast, reportWarning, customFonts);
    // The native engine reflows a single stream and has no running page furniture, so master-page
    // headers/footers (ast.auxiliary, present only when parsing kept them) are drawn once - headers
    // before the body, footers after - rather than dropped. Footnote/endnote bodies gathered during the
    // walk are drawn last.
    for (const n of ast.auxiliary?.headers || []) await layout.render(n);
    // Gather footnote/endnote bodies from the whole content tree first (they attach to text runs, deep
    // in the tree, so a render()-level sweep would miss them), then render the body, then draw them.
    for (const node of ast.content) layout.collectNotes(node);
    for (const node of ast.content) await layout.render(node, true);
    for (const n of ast.auxiliary?.footers || []) await layout.render(n);
    await layout.flushNotes();
    layout.reportLostCharacters();
    if (pc.outline) writeOutline(pdf, lib, layout.outlineEntries);

    return await pdf.save();
}

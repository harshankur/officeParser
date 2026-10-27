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

import { fillPageNumber } from '../../utils/pageNumbers.js';
import { EmbedMetadata, FullGeneratorConfig, ImageMode, IndentationMetadata, OfficeContentNode, OfficeErrorType, OfficeMetadata, OfficeParserAST, OfficeWarningType, PageSetup, ParagraphSpacing, SheetMetadata, TextAlignment, TextFormatting, BreakMetadata, ChartData, HeaderFooterMetadata, NodeBounds, ShapeMetadata, SlideMetadata, PageNumberFormat } from '../../types.js';
import { hasPageSize, onDocumentPage } from '../../utils/pageLimits.js';
import { getAbortError, getOfficeError } from '../../utils/errorUtils.js';
import { documentLanguage, embedUrl, isHeaderRow, lengthToPt, paperSizePt, resolveImageMode, sniffImageSize, writableSpacing } from '../../utils/officeGenUtils.js';
import { decodeBmp, decodeGif, encodePng } from '../../utils/rasterUtils.js';
import { sanitizeOfficePackageUrl } from '../../utils/sanitize.js';
import { cssColorHex } from '../../utils/colorUtils.js';
import { formatListNumber } from '../../utils/listMarkers.js';
import { drawSvg, readSvg } from './svgDrawing.js';
import { colorTables, fontkitUnbrotli, nativeUnbrotli, sfntOf, substitutes, trueTypeOutlines, Unbrotli, UnreadableFont } from './fontFiles.js';

/** The largest of `values` and `floor`, without spreading them into one call (which a long list overflows). */
function maxOf(values: readonly number[], floor: number): number {
    let m = floor;
    for (const v of values) if (v > m) m = v;
    return m;
}

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

/**
 * Characters a font draws nothing for by design: spaces, controls and format characters, those Unicode
 * says are ignored where they are not supported (variation selectors, fillers), and U+2800 BRAILLE PATTERN
 * BLANK, the blank cell. Any other character a font maps to a glyph that draws nothing is not the font's
 * to draw, but for a combining mark drawn with its letter (see Face.has): a font after it that draws it
 * does, and one no font draws is lost (see embedFonts). Such a glyph is the font's defect (U+FC39 in one
 * of Fontsource's files of Noto Sans Arabic), or a color glyph's stand-in, drawn from color images or
 * layers a PDF does not hold.
 */
const BLANK = /[\p{Z}\p{Cc}\p{Cf}\p{Default_Ignorable_Code_Point}]/u;
function isBlank(cp: number): boolean {
    return isInvisible(cp) || cp === 0x2800 || BLANK.test(String.fromCodePoint(cp));
}

/** Hebrew, Arabic and the other right-to-left blocks of the Basic Multilingual Plane. */
function isRtlChar(cp: number): boolean {
    return (cp >= 0x0590 && cp <= 0x08FF) || (cp >= 0xFB1D && cp <= 0xFDFF) || (cp >= 0xFE70 && cp <= 0xFEFF);
}

const LETTER_OR_DIGIT = /[\p{L}\p{N}]/u;
const OTHER_SPACE = /[\p{Zs}\u000B\u000C\u2028\u2029]/u;
/** What text is cut into words at. */
const WHITESPACE = /\s/;
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
    for (let level = maxOf(levels, 0); level >= 1; level--) {
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

/** The most frames floating on one page whose room the text keeps (see floatingFrame). */
const MAX_FLOATS_PER_PAGE = 256;

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

/**
 * Readies a font given to the engine (embedded by pdf-lib) for the glyphs shaping gives, and returns the
 * number each glyph is shown by: kept in the font's subset, or, for a font embedded whole, written into
 * its widths and its text mapping. pdf-lib writes those for a whole font from the characters the font
 * maps alone, so a ligature, a conjunct, a joined or reordered form took a width of 1000 units and no
 * text; they are written for the glyphs drawn instead (and only those: a font of 60,000 glyphs wrote all).
 * An OpenType font of CFF outlines embedded whole is marked as one (`/Subtype /OpenType`), not as the bare
 * CFF it is not.
 */
function glyphCoder(font: any, lib: any): (glyph: any) => number {
    const embedder = font.embedder;
    if (embedder.subset) {
        return glyph => {
            const id = embedder.subset.includeGlyph(glyph);
            embedder.glyphs[id - 1] = glyph;
            embedder.glyphIdMap.set(glyph.id, id);
            embedder.glyphCache.invalidate();
            return id;
        };
    }
    const used = new Map<number, any>();
    // (A font no text was drawn in still writes a width: its first glyph's.)
    embedder.glyphCache.populate = () => used.size ? [...used.values()].sort((a, b) => a.id - b.id) : [embedder.font.getGlyph(0)];
    embedder.glyphCache.invalidate();
    const note = (glyph: any): number => {
        if (!used.has(glyph.id)) { used.set(glyph.id, glyph); embedder.glyphCache.invalidate(); }
        return glyph.id;
    };
    embedder.encodeText = (text: string) => {
        let hex = '';
        for (const glyph of embedder.font.layout(text, embedder.fontFeatures).glyphs) hex += note(glyph).toString(16).padStart(4, '0');
        return lib.PDFHexString.of(hex);
    };
    // (pdf-lib asks fontkit's `cff`, which it does not have: a CFF font was written as a TrueType one.)
    if (embedder.font['CFF '] || embedder.font['CFF2']) {
        embedder.isCFF = () => true;
        embedder.embedFontStream = async (context: any) => context.register(context.flateStream(await embedder.serializeFont(), { Subtype: 'OpenType' }));
    }
    return note;
}

/**
 * Makes a font given to the engine that cannot be embedded fail the conversion, with `fail`'s error
 * (which names the font), when the PDF is saved. First the outlines of the glyphs it is to hold are given
 * to `check` (see TrueTypeOutlines.check): a glyph that viewers cannot draw was embedded without a word.
 * Then fontkit writes it: pdf-lib takes a font's subset from `@pdf-lib/fontkit` 1.x as a stream that
 * fontkit writes on a later turn of the event loop, and a glyph it could not copy threw there, where no
 * caller could catch it, and the process ended while the conversion waited. The subset's writing is
 * wrapped so that what it throws is kept, and the font's embedding fails with it.
 */
function guardEmbedding(font: any, check: (glyphs: number[]) => void, fail: (error: unknown) => Error): void {
    const embedder = font?.embedder;
    if (!embedder || typeof embedder.embedIntoContext !== 'function') return;
    const subset = embedder.subset;
    let failure: { error: unknown } | undefined;
    if (subset && typeof subset.encode === 'function') {
        const encode = subset.encode;
        subset.encode = function (this: unknown, ...args: unknown[]) {
            try { return encode.apply(this, args); } catch (error) {
                failure ??= { error };
                return undefined;
            }
        };
    }
    const embedInto = embedder.embedIntoContext.bind(embedder);
    embedder.embedIntoContext = async (...args: unknown[]) => {
        let done: unknown;
        try {
            // (A subset holds the glyphs it was given, a font embedded whole writes the widths of those drawn.)
            const glyphs: unknown[] = subset ? subset.glyphs ?? [] : (embedder.glyphCache?.access?.() ?? []).map((glyph: any) => glyph?.id);
            check(glyphs.filter((id): id is number => typeof id === 'number'));
            done = await embedInto(...args);
        } catch (error) { throw fail(failure ? failure.error : error); }
        if (failure) throw fail(failure.error);
        return done;
    };
}

/** Thrown out of a font's state machine when it runs past its bound (see guardShaping). */
// (No message: it never leaves guardShaping, which catches it by identity.)
const SHAPING_RAN_AWAY = new Error();

/**
 * Keeps fontkit's shaping of a font from running forever. A font of Apple's layout tables (`morx`, which
 * fontkit prefers to OpenType's) is shaped by state machines that, on some input, never end: a glyph
 * insertion table inserting a glyph a step without moving on, on a vowel sign alone or after a space,
 * after an independent vowel (अा, the old spelling of आ), a doubled anusvara or virama. It filled the heap
 * and ended the process, which no caller can catch. Each of this font's state machines (one a subtable,
 * kept by fontkit for the font's processor) is bounded, in steps and in glyphs, by the glyphs it was given
 * (real text takes about two steps a glyph, and adds none); one past its bound stops, and the text is
 * shaped again a word at a time, then a word that runs away a grapheme cluster at a time, and a cluster
 * that runs away is laid out unshaped (each character its glyph, at its own advance). Other fonts, and
 * other users of fontkit, are as they were.
 */
export function guardShaping(font: any): void {
    const shaper = font?.embedder?.font;
    const processor = shaper?.layout ? shaper._layoutEngine?.engine?.morxProcessor : undefined;
    if (!processor?.getStateMachine) return;
    let bounded = 0;
    for (const chain of processor.morx?.chains ?? []) {
        for (const subtable of chain.subtables ?? []) {
            if (!subtable?.table?.stateTable) continue;
            let machine: any;
            try { machine = processor.getStateMachine(subtable); } catch { continue; }
            if (!machine || typeof machine.process !== 'function' || machine.bounded) continue;
            const run = machine.process.bind(machine);
            machine.process = (glyphs: unknown[], reverse: boolean, entry: (...args: unknown[]) => unknown) => {
                const given = glyphs.length, steps = { left: 64 * given + 256 };
                return run(glyphs, reverse, (...args: unknown[]) => {
                    if (--steps.left < 0 || glyphs.length > 8 * given + 64) throw SHAPING_RAN_AWAY;
                    return entry(...args);
                });
            };
            machine.bounded = true;
            bounded++;
        }
    }
    if (!bounded) return;
    const layout = shaper.layout.bind(shaper);
    const unshaped = (text: string) => {
        const glyphs: any[] = [], positions: any[] = [];
        for (const ch of text) {
            const glyph = shaper.glyphForCodePoint(ch.codePointAt(0)!);
            glyphs.push(glyph);
            positions.push({ xAdvance: glyph.advanceWidth, yAdvance: 0, xOffset: 0, yOffset: 0 });
        }
        return { glyphs, positions };
    };
    /** `text` laid out, or undefined where a state machine ran away on it. */
    const attempt = (text: string, rest: unknown[]): any => {
        try { return layout(text, ...rest); } catch (error) {
            if (error !== SHAPING_RAN_AWAY) throw error;
            return undefined;
        }
    };
    shaper.layout = (text: unknown, ...rest: unknown[]) => {
        if (typeof text !== 'string') return layout(text, ...rest);
        const whole = attempt(text, rest);
        if (whole) return whole;
        const glyphs: any[] = [], positions: any[] = [];
        let last: any;
        const add = (part: any) => { glyphs.push(...part.glyphs); positions.push(...part.positions); };
        // (Shaping never crosses a space: a word at a time keeps the font's shaping of every other word.)
        for (const word of text.split(/(\s+)/)) {
            if (!word) continue;
            const shaped = /^\s+$/.test(word) ? undefined : attempt(word, rest);
            if (shaped) { add(last = shaped); continue; }
            if (/^\s+$/.test(word)) { add(unshaped(word)); continue; }
            for (const cluster of clustersOf(word)) add(attempt(cluster, rest) ?? unshaped(cluster));
        }
        return { ...(last ?? {}), glyphs, positions };
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

/** Whether a picture's bytes are an SVG: by its type or name, or by its text starting `<svg` or `<?xml`. */
function isSvg(bytes: Uint8Array, attachment: { extension?: string; mimeType?: string }): boolean {
    if (/svg/i.test(attachment.mimeType || '') || /^svg$/i.test(attachment.extension || '')) return true;
    const head = new TextDecoder().decode(bytes.subarray(0, 512)).replace(/^\uFEFF/, '').trimStart();
    return /^<(?:\?xml[^>]*>\s*)?(?:<!--[\s\S]*?-->\s*)*(?:<!DOCTYPE[^>]*>\s*)?<svg[\s>]/i.test(head);
}

/** The styles text is drawn in. */
type FaceKey = 'regular' | 'bold' | 'italic' | 'boldItalic' | 'mono'
    | 'serif' | 'serifBold' | 'serifItalic' | 'serifBoldItalic' | 'monoBold' | 'monoItalic' | 'monoBoldItalic';

/** The styles `pdfConfig.fonts` gives a font for; the others are drawn in a stand-in of their kind. */
type GivenFaceKey = 'regular' | 'bold' | 'italic' | 'boldItalic' | 'mono';

/**
 * The kind of a font family, by its name: monospaced, serif or (anything else) sans-serif, for the
 * Standard-14 font that stands in for it. Serif text printed in Helvetica, and code inline in a line in
 * it too.
 */
function fontKind(family: string | undefined): 'mono' | 'serif' | 'sans' {
    if (!family) return 'sans';
    const name = family.toLowerCase().replace(/["']/g, '').trim();
    if (/mono|code|courier|consol|menlo|monaco|typewriter|terminal|fixed|^ocr|lucida console|andale|inconsolata|hack$|^pt mono/.test(name)) return 'mono';
    if (/sans/.test(name)) return 'sans';
    if (/serif|times|georgia|garamond|cambria|palatino|book ?antiqua|baskerville|bodoni|didot|caslon|minion|constantia|century|bookman|charter|cochin|hoefler|lora|merriweather|playfair|perpetua|rockwell|sylfaen|mincho|ming|song|batang|myeongjo|^roman|nimbus roman|tinos|crimson|cormorant|liberation serif|dejavu serif|noto serif|source serif|pt serif|droid serif|charis|gentium|linux libertine|libertinus|stix|cmu serif|computer modern|latin modern roman|lm roman/.test(name)) return 'serif';
    return 'sans';
}

/** A single word ready to place, carrying the style/size/color it should be drawn with. */
interface Piece { text: string; face: FaceKey; size: number; color: any; underline: boolean; link?: string; background?: string; rise?: number; note?: OfficeContentNode; picture?: { img: any; w: number; h: number }; /** Drawn invisibly (TextFormatting.invisible): in its place, there to select and search. */ invisible?: boolean; }

/** A stretch of inline text in one formatting, and the target it links to, if any. */
type Run = { text: string; fmt: TextFormatting; link?: string; note?: OfficeContentNode; picture?: { img: any; w: number; h: number } };

/**
 * How a paragraph's lines sit: its alignment, where its first line starts against the others, and the
 * distance from one line to the next (a multiple of the usual, or fixed, or at least a height).
 */
interface LineLayout { align?: TextAlignment; firstLine?: number; lineMultiple?: number; lineHeight?: number; lineAtLeast?: boolean }

/** One embedded font and the code points it has a glyph for. */
interface Face {
    font: any;
    /** Whether the font draws the character: its glyph draws something, or the character is blank by design (see BLANK). */
    covers: (cp: number) => boolean;
    /**
     * Whether the font may draw the character: it covers it, or its glyph for it is empty in a font that
     * substitutes glyphs as it shapes text (GSUB, morx). There a combining mark drawn with the letter
     * before it (an invisible virama that stacks the next letter, a mark composed into its letter) may
     * have an empty glyph of its own; a font that substitutes none draws an empty glyph as nothing.
     */
    has: (cp: number) => boolean;
    /** For a font given (not a Standard-14 one): the number a glyph is shown by, the glyph kept in the font (see glyphCoder). */
    code?: (glyph: any) => number;
}

/** A run of text as its font shapes it (fontkit): its glyphs, in the order drawn, and where each goes. */
interface ShapedRun { glyphs: any[]; positions: { xAdvance: number; xOffset: number; yOffset: number }[]; unitsPerEm: number }

/**
 * Text a font shapes beyond one glyph a character at its advance: marks (placed over or under their
 * letter), and the scripts of South and Southeast Asia, whose letters form conjuncts or reorder. Drawn
 * as shaped (see NativeLayout.drawShaped); other text, right-to-left text among it, as pdf-lib draws it
 * (its glyphs in the order they are seen, each mapped to its letters: readers put those back in the order
 * they are read, and read an `/ActualText` of right-to-left text, or glyphs moved by their offsets, apart
 * or backwards).
 */
const SHAPED_TEXT = /[\p{M}\u0900-\u0DFF\u0E00-\u0FFF\u1000-\u109F\u1780-\u17FF\u1A00-\u1BFF\uA800-\uABFF]/u;

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
        private pageW: number,
        private pageH: number,
        private margin: { top: number; right: number; bottom: number; left: number },
        private readonly config: FullGeneratorConfig,
        private readonly ast: OfficeParserAST,
        // Reports a warning through the generator's own channel, so it lands in `result.messages`
        // (and reaches onWarning) rather than only firing onWarning directly.
        private readonly reportWarning: (type: OfficeWarningType, info?: any) => void,
        /** Whether any font was given: the Standard-14 ones are then the last resort, not the only fonts. */
        private readonly customFonts: boolean,
        /**
         * Each section's pages from its PageSetup (pdfConfig.format 'auto'), the margins the caller left
         * unset taken from it; null when the caller's paper is used throughout.
         */
        private readonly sectionPage: ((setup: PageSetup | undefined) => { w: number; h: number; margin: { top: number; right: number; bottom: number; left: number } }) | null,
        /** The first section's distances of its header and footer from the page's edges, and whether its first page has its own. */
        private headerDistance?: number,
        private footerDistance?: number,
        private differentFirstPage?: boolean,
    ) {
        this.imageMode = resolveImageMode(this.config.includeImages);
        for (const a of this.ast.attachments || []) {
            if (a.name && !this.attachmentsByName.has(a.name)) this.attachmentsByName.set(a.name, a);
        }
        this.sectionSetups.set(0, this.ast.metadata?.pageSetup);
        this.newPage();
    }

    private get contentWidth(): number { return this.pageW - this.margin.left - this.margin.right; }
    /**
     * The lowest a line of the body goes: above the page's margin, and above its footnotes. On a page
     * drawn as it is (a header, a footer, a footnote, or one measured), as low as it needs.
     */
    private get bottom(): number {
        if (this.fixedPage) return Number.MAX_SAFE_INTEGER;
        const page = this.pageH - this.margin.bottom - this.footReserve;
        // Balanced columns end level; the last takes what is left of the page, should the measure have been short.
        return this.columnLimit !== null && this.columns && this.columnIndex < this.columns.count - 1 ? Math.min(this.columnLimit, page) : page;
    }

    // ── columns ──────────────────────────────────────────────────────────────────

    /**
     * The columns the section is laid out in (PageSetup.columns), which one the text is in, where they
     * start on this page, the lowest any has reached, the page's own margins (the text's are the
     * column's), and the bottom of columns balanced to end level (see setColumns).
     */
    private columns: NonNullable<PageSetup['columns']> | null = null;
    private columnIndex = 0;
    private columnTop = 0;
    private columnsBottom = 0;
    private pageMargin: { top: number; right: number; bottom: number; left: number } | null = null;
    private columnLimit: number | null = null;
    /**
     * A balanced section running past this page: the room its text takes (its height in the first column
     * times that column's width), and the room the columns filled so far took; its last page is balanced
     * on what is left (see balanceOnThisPage).
     */
    private balanceArea: number | null = null;
    private usedArea = 0;
    /** The document's top-level blocks, and which is being drawn: what a balanced section holds. */
    flow: OfficeContentNode[] = [];
    flowIndex = -1;

    /** The text's margins for the column it is in: the page's, narrowed to that column. */
    private applyColumn(): void {
        const pm = this.pageMargin, cols = this.columns;
        if (!pm || !cols) return;
        const room = this.pageW - pm.left - pm.right;
        const gap = Math.min(cols.gap, room / (2 * cols.count));
        const given = cols.widths?.length === cols.count && cols.widths.every(w => w > 0) ? cols.widths : undefined;
        const scale = given ? (room - gap * (cols.count - 1)) / given.reduce((a, b) => a + b, 0) : 1;
        const widths = given ? given.map(w => w * scale) : new Array<number>(cols.count).fill((room - gap * (cols.count - 1)) / cols.count);
        let left = pm.left;
        for (let i = 0; i < this.columnIndex; i++) left += widths[i] + gap;
        this.margin = { ...pm, left, right: this.pageW - left - widths[this.columnIndex] };
    }

    /** Ends the columns on this page: the lines between them, where the section draws them, to the lowest one's end. */
    private closeColumnsOnPage(): void {
        if (!this.columns || !this.pageMargin) return;
        this.columnsBottom = Math.max(this.columnsBottom, this.y);
        if (!this.columns.separator || this.columnIndex === 0) return;
        const saved = { margin: this.margin, index: this.columnIndex };
        for (let i = 1; i <= saved.index; i++) {
            this.columnIndex = i;
            this.applyColumn();
            const x = this.margin.left - Math.min(this.columns.gap, (this.pageW - this.pageMargin.left - this.pageMargin.right) / (2 * this.columns.count)) / 2;
            this.page.drawLine({ start: { x, y: this.pageH - this.columnTop }, end: { x, y: this.pageH - this.columnsBottom }, thickness: 0.5, color: this.lib.rgb(0, 0, 0) });
        }
        this.columnIndex = saved.index;
        this.margin = saved.margin;
    }

    /** The widths of all the text columns, as applyColumn lays them out. */
    private textColumnWidths(): number[] {
        const saved = { index: this.columnIndex, margin: this.margin };
        const out: number[] = [];
        for (let i = 0; i < (this.columns?.count ?? 0); i++) { this.columnIndex = i; this.applyColumn(); out.push(this.contentWidth); }
        this.columnIndex = saved.index;
        this.margin = saved.margin;
        return out;
    }

    /**
     * Balances the columns of this page on `area` (room as height times width): level at the height the
     * columns' widths together need for it, where that fits on the page. Whether it fit.
     */
    private balanceOnThisPage(area: number): boolean {
        const total = this.textColumnWidths().reduce((a, b) => a + b, 0);
        if (!(total > 0)) return false;
        const limit = this.columnTop + area / total;
        if (limit >= this.pageH - (this.pageMargin ?? this.margin).bottom - this.footReserve) return false;
        this.columnLimit = limit;
        return true;
    }

    /** The text runs on into the next column, or, from the last, onto the next page. */
    private nextColumnOrPage(): void {
        // The room the column took, for a balanced section's last page.
        if (this.columns && this.balanceArea !== null) this.usedArea += Math.max(0, this.y - this.columnTop) * this.contentWidth;
        if (this.columns && !this.fixedPage && this.columnIndex < this.columns.count - 1) {
            this.columnsBottom = Math.max(this.columnsBottom, this.y);
            this.columnIndex++;
            this.applyColumn();
            this.y = this.columnTop;
            return;
        }
        this.newPage();
    }

    /**
     * Lays the text out in `columns` from here on (none, or one, for the page's width), the columns
     * before ending at the lowest of them. A section that the next ends on the same page (a section
     * break that runs on) is balanced, as a word processor balances it: measured in a column's width
     * first, its columns end level, where they fit on this page.
     */
    async setColumns(columns: PageSetup['columns'] | undefined, balance: boolean): Promise<void> {
        if (this.columns) {
            this.closeColumnsOnPage();
            this.y = Math.max(this.columnsBottom, this.y);
            this.margin = this.pageMargin!;
            this.columns = null;
            this.pageMargin = null;
            this.columnLimit = null;
        }
        this.balanceArea = null;
        this.usedArea = 0;
        if (!columns || columns.count < 2) return;
        this.pageMargin = { ...this.margin };
        this.columns = { ...columns, count: Math.min(columns.count, 16) };
        this.columnIndex = 0;
        this.columnTop = this.y;
        this.columnsBottom = this.y;
        this.applyColumn();
        if (!balance) return;
        // The blocks up to the break that ends the section, measured in the column's width.
        const region: OfficeContentNode[] = [];
        for (let i = this.flowIndex + 1; i < this.flow.length; i++) {
            const meta = this.flow[i].type === 'break' ? this.flow[i].metadata as BreakMetadata | undefined : undefined;
            if (meta?.breakType === 'section' || (meta?.breakType === 'page' && typeof meta.section === 'number')) {
                if (meta.breakType !== 'section') return;
                break;
            }
            if (i === this.flow.length - 1) return;
            region.push(this.flow[i]);
        }
        // The room its text takes, measured in a column of the columns' mean width (measured in the first, a
        // narrow first column made the room too large to level the columns): on this page where it fits
        // (its columns, of whatever widths, ending level), else on the page it ends on, on what the pages
        // before left. (The last column has no limit: what the measure missed, it takes.)
        const pm = this.pageMargin, gap = Math.min(this.columns.gap, (this.pageW - pm.left - pm.right) / (2 * this.columns.count));
        const mean = Math.max(1, (this.pageW - pm.left - pm.right - gap * (this.columns.count - 1)) / this.columns.count);
        this.margin = { ...pm, right: this.pageW - pm.left - mean };
        let height: number;
        try { height = await this.measured(async () => { for (const node of region) await this.render(node); }); }
        finally { this.applyColumn(); }
        const area = height * mean;
        if (!this.balanceOnThisPage(area)) { this.balanceArea = area; this.usedArea = 0; }
    }

    // ── sections ─────────────────────────────────────────────────────────────────

    /** The section being laid out (see BreakMetadata.section), and whether the next page starts it. */
    private section = 0;
    private startingSection = true;
    /** Every page made: its section, whether it is that section's first, and its size and margins. */
    private readonly pageRecords: { page: any; section: number; firstOfSection: boolean; w: number; h: number; margin: { top: number; right: number; bottom: number; left: number }; headerDistance?: number; footerDistance?: number; differentFirstPage?: boolean; sourcePage?: number }[] = [];
    /** The page of the PDF being drawn whose content is being laid out (a PDF's `page` node's number). */
    private sourcePage: number | undefined;

    /** Notes the page just made, for its header and footer. */
    private recordPage(): void {
        this.pageRecords.push({ page: this.page, section: this.section, firstOfSection: this.startingSection, w: this.pageW, h: this.pageH, margin: { ...(this.pageMargin ?? this.margin) }, headerDistance: this.headerDistance, footerDistance: this.footerDistance, differentFirstPage: this.differentFirstPage, sourcePage: this.sourcePage });
        this.startingSection = false;
    }

    /**
     * Starts a section of the document: its pages' size and margins (pdfConfig.format 'auto'), on a new
     * page, or on this one when nothing is drawn on it yet (a page break before the section's own).
     */
    private startSection(section: number, setup: PageSetup | undefined): void {
        this.noteSection(section, setup);
        this.startingSection = true;
        this.differentFirstPage = setup?.differentFirstPage;
        this.onPagesOf(setup, true);
    }

    /** The section the text is in from here, and its setup (how its pages are numbered), for its pages. */
    private noteSection(section: number, setup: PageSetup | undefined): void {
        this.section = section;
        this.sectionSetups.set(section, setup);
    }

    /**
     * Lays the pages that follow out as `setup` has them (pdfConfig.format 'auto'), or with none, as the
     * document's first pages are (`orDocument`): on this page when nothing is drawn on it yet, else on a
     * new one. `always` starts a page even when the size is the same.
     */
    private onPagesOf(setup: PageSetup | undefined, always: boolean, orDocument = false): void {
        const before = { w: this.pageW, h: this.pageH, margin: this.margin };
        // (A section that sets no page is on the document's: its size, margins and header distances.)
        const pages = setup ? onDocumentPage(setup, this.ast.metadata?.pageSetup) : orDocument ? this.ast.metadata?.pageSetup : undefined;
        const next = this.sectionPage && (setup || orDocument) ? this.sectionPage(pages) : undefined;
        const fit = next ? this.withFurnitureRoom(next.margin, pages?.headerDistance, pages?.footerDistance) : undefined;
        const nextMargin = fit ? fit.margin : before.margin;
        const changed = !!next && (before.w !== next.w || before.h !== next.h || JSON.stringify(before.margin) !== JSON.stringify(nextMargin));
        if (!always && !changed) return;
        const record = this.pageRecords[this.pageRecords.length - 1];
        const reuse = !!record && record.page === this.page && this.y <= before.margin.top + 0.5 && !this.pageFootnotes.length;
        // The page that ends draws its footnotes in its own margins, before the next pages' take over.
        if (!reuse) this.drawPageFootnotes();
        if (next && fit) {
            this.pageW = next.w; this.pageH = next.h; this.margin = nextMargin;
            this.headerDistance = fit.headerDistance;
            this.footerDistance = fit.footerDistance;
        }
        if (reuse) {
            this.page.setSize(this.pageW, this.pageH);
            Object.assign(record, { section: this.section, firstOfSection: record.firstOfSection || this.startingSection, w: this.pageW, h: this.pageH, margin: { ...this.margin }, headerDistance: this.headerDistance, footerDistance: this.footerDistance, differentFirstPage: this.differentFirstPage });
            this.startingSection = false;
            this.y = this.margin.top;
            return;
        }
        this.newPage();
    }

    /**
     * A page added at the end of the document: on the document's page tree, where it is flat (as one this
     * engine makes is), the page put at its end directly. pdf-lib's addPage walks every page before it to
     * find the end, so a document of thousands of pages took time in their square.
     */
    private appendPage(w: number, h: number): any {
        // The frames floating on the page before are done with.
        this.floats = [];
        const pdf = this.pdf as any;
        const tree = pdf.catalog?.Pages?.();
        const treeRef = pdf.catalog?.get?.(this.lib.PDFName.of('Pages'));
        if (!tree || !treeRef || tree.Kids().size() !== tree.Count().asNumber() || !pdf.pageMap || !pdf.pageCache) return this.pdf.addPage([w, h]);
        const page = this.lib.PDFPage.create(this.pdf);
        page.setSize(w, h);
        tree.pushLeafNode(page.ref);
        page.node.setParent(treeRef);
        pdf.pageMap.set(page.node, page);
        pdf.pageCache.invalidate();
        pdf.pageCount = tree.Count().asNumber();
        return page;
    }

    /** Gives this page, while nothing is drawn on it, the size a page node of the source has. */
    private takePageSize(meta: { pageWidth?: number; pageHeight?: number } | undefined): void {
        const w = meta?.pageWidth, h = meta?.pageHeight;
        // Room for the margins and an inch of text, and no larger than a PDF page can be.
        if (!this.sectionPage || !w || !h || !(w > this.margin.left + this.margin.right + 72 && h > this.margin.top + this.margin.bottom + 72 && w <= 14400 && h <= 14400)) return;
        const record = this.pageRecords[this.pageRecords.length - 1];
        if (!record || record.page !== this.page || this.y > this.margin.top + 0.5 || (Math.abs(w - this.pageW) < 0.5 && Math.abs(h - this.pageH) < 0.5)) return;
        this.pageW = w; this.pageH = h;
        this.page.setSize(w, h);
        Object.assign(record, { w, h });
    }

    // ── slides ───────────────────────────────────────────────────────────────────

    /**
     * Draws a slide as it is laid out: its background over the page, each shape (ShapeMetadata) where
     * it is, each shape's text in its box (set from its top, middle or bottom, turned with it), each
     * picture over its frame, each table and chart in its box; the slide scaled to the page where the
     * page is another size. A deck was laid out as flowing text, its shapes, fills and places lost.
     */
    private async drawSlide(node: OfficeContentNode): Promise<void> {
        const meta = node.metadata as SlideMetadata | undefined;
        const children = node.children || [];
        // On a page of its own (after other content too), the slide's anchors at its top.
        if (this.y > this.margin.top + 0.5) this.newPage();
        if (this.pendingAnchors.length) this.placeAnchors();
        // The slide's size (the presentation's page), and its place on this page.
        const slideW = this.ast.metadata?.pageSetup?.width || this.pageW, slideH = this.ast.metadata?.pageSetup?.height || this.pageH;
        const scale = Math.min(this.pageW / slideW, this.pageH / slideH);
        const offsetX = (this.pageW - slideW * scale) / 2, offsetY = (this.pageH - slideH * scale) / 2;
        const place = (b: NodeBounds) => ({ x: offsetX + b.x * scale, y: offsetY + b.y * scale, width: b.width * scale, height: b.height * scale });
        const saved = { margin: this.margin, fixed: this.fixedPage, y: this.y };
        const pageNumber = this.pageRecords.length;
        if (meta?.background) this.page.drawRectangle({ x: offsetX, y: this.pageH - offsetY - slideH * scale, width: slideW * scale, height: slideH * scale, color: this.color(meta.background) });
        this.fixedPage = true;
        this.fieldPage = pageNumber;
        const flowing: OfficeContentNode[] = [];
        try {
            for (let i = 0; i < children.length; i++) {
                const child = children[i];
                if (!child.bounds) { flowing.push(child); continue; }
                const box = place(child.bounds);
                if (child.type === 'drawing') {
                    const shape = child.metadata as ShapeMetadata | undefined;
                    this.turned(box, shape, () => this.drawShape(box, shape));
                    // The shape's text: the nodes after it in its box.
                    const text: OfficeContentNode[] = [];
                    while (i + 1 < children.length && children[i + 1].type !== 'drawing' && sameBox(children[i + 1].bounds, child.bounds)) text.push(children[++i]);
                    if (text.length) await this.drawTextBox(text, box, shape);
                } else if (child.type === 'image') {
                    const picture = this.imageMode === 'image-only' || this.imageMode === 'image+ocr-text' ? await this.embedPicture(child) : null;
                    if (picture) this.drawPicture(picture.img, { x: box.x, y: this.pageH - box.y - box.height, width: box.width, height: box.height });
                    else await this.drawTextBox([child], box, undefined);
                } else if (child.type === 'table' || child.type === 'chart') {
                    this.margin = { top: box.y, left: box.x, right: this.pageW - box.x - box.width, bottom: 0 };
                    this.y = box.y;
                    this.chartRoom = child.type === 'chart' ? box.height : undefined;
                    try { await this.render(child); } finally { this.chartRoom = undefined; }
                } else {
                    const text: OfficeContentNode[] = [child];
                    while (i + 1 < children.length && children[i + 1].type !== 'drawing' && sameBox(children[i + 1].bounds, child.bounds)) text.push(children[++i]);
                    await this.drawTextBox(text, box, undefined);
                }
            }
        } finally {
            this.margin = saved.margin;
            this.fixedPage = saved.fixed;
            this.fieldPage = undefined;
            this.y = saved.y;
        }
        // What has no place on the slide, and its speaker notes (pdfConfig.printSpeakerNotes), follow it.
        if (flowing.length || node.notes?.length) {
            this.newPage();
            for (const c of flowing) await this.render(c);
            for (const n of node.notes || []) await this.note(n);
        } else {
            // The next slide starts a page of its own (see render's pagination): this one is full.
            this.y = this.bottom;
        }
    }

    /** Draws `draw` turned (and mirrored) about the box's centre, as the shape is. */
    private turned(box: { x: number; y: number; width: number; height: number }, shape: ShapeMetadata | undefined, draw: () => void): void {
        const rotation = shape?.rotation || 0;
        if (!rotation && !shape?.flipH && !shape?.flipV) { draw(); return; }
        const cx = box.x + box.width / 2, cy = this.pageH - box.y - box.height / 2;
        const angle = -rotation * Math.PI / 180, cos = Math.cos(angle), sin = Math.sin(angle);
        const sx = shape?.flipH ? -1 : 1, sy = shape?.flipV ? -1 : 1;
        // Translate to the centre, turn (clockwise on the page), mirror, and translate back.
        const a = cos * sx, b = sin * sx, c = -sin * sy, d = cos * sy;
        this.page.pushOperators(this.lib.pushGraphicsState(), this.lib.concatTransformationMatrix(a, b, c, d, cx - a * cx - c * cy, cy - b * cx - d * cy));
        try { draw(); } finally { this.page.pushOperators(this.lib.popGraphicsState()); }
    }

    /** Draws a shape's outline in its fill and line: a box, a rounded box, an ellipse, a line, a triangle or a diamond. */
    private drawShape(box: { x: number; y: number; width: number; height: number }, shape: ShapeMetadata | undefined): void {
        if (!shape?.fill && !shape?.line) return;
        const fill = shape.fill ? this.color(shape.fill) : undefined;
        const line = shape.line?.color ? { borderColor: this.color(shape.line.color), borderWidth: shape.line.width ?? 0.75 } : {};
        const bottom = this.pageH - box.y - box.height;
        switch (shape.geometry) {
            case 'ellipse':
            case 'circle':
                this.page.drawEllipse({ x: box.x + box.width / 2, y: bottom + box.height / 2, xScale: box.width / 2, yScale: box.height / 2, ...(fill ? { color: fill } : {}), ...line });
                return;
            case 'line':
            case 'straightConnector1':
                this.page.drawLine({ start: { x: box.x, y: bottom + box.height }, end: { x: box.x + box.width, y: bottom }, thickness: shape.line?.width ?? 0.75, color: (shape.line?.color && this.color(shape.line.color)) || fill || this.lib.rgb(0, 0, 0) });
                return;
            case 'triangle':
            case 'diamond': {
                const w = box.width, h = box.height;
                const path = shape.geometry === 'triangle' ? `M ${w / 2} 0 L ${w} ${h} L 0 ${h} Z` : `M ${w / 2} 0 L ${w} ${h / 2} L ${w / 2} ${h} L 0 ${h / 2} Z`;
                this.page.drawSvgPath(path, { x: box.x, y: this.pageH - box.y, ...(fill ? { color: fill } : {}), ...line });
                return;
            }
            default:
                // A box, a rounded box, and the outlines not drawn as themselves: their box, in their colours.
                this.page.drawRectangle({ x: box.x, y: bottom, width: box.width, height: box.height, ...(fill ? { color: fill } : {}), ...line });
        }
    }

    /**
     * The frames floating beside the text on their pages (see floatingFrame), with the room the text keeps
     * from them: lines beside one go round it (see drawRuns), other blocks go below it.
     */
    private floats: { page: unknown; top: number; bottom: number; left: number; right: number; wrap: 'beside' | 'aboveBelow' }[] = [];

    /**
     * Draws a frame floating beside a document's text (ShapeMetadata.float) where it stands, measured from
     * the paragraph after it (its anchor), the page's text area or the page: its box, and its blocks within
     * it. Its room on the page is kept for the text after it. A frame that does not fit below its anchor
     * on the page goes on the next, with its paragraph.
     */
    private async floatingFrame(node: OfficeContentNode): Promise<void> {
        const shape = node.metadata as ShapeMetadata;
        const float = shape.float!;
        const insets = shape.textInsets ?? { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 };
        const flow = async () => { for (const child of node.children || []) await this.render(child); };
        // (A frame of no width it can be drawn at, as only a hand-made one is: its blocks in the flow.)
        if (!(Number.isFinite(float.width) && float.width > 0)) return this.inFlow(flow);
        const width = Math.max(4, Math.min(float.width, this.pageW));
        // Its text, measured in a box of its width (its insets in it).
        const margin = this.margin;
        let textHeight = 0;
        try {
            textHeight = insets.top + insets.bottom + await this.measured(async () => {
                this.margin = { ...margin, left: insets.left, right: Math.max(0, this.pageW - width + insets.right), top: 0 };
                this.y = 0;
                await flow();
            });
        } finally { this.margin = margin; }
        // Text more than a page holds is set in the flow, all of it: a frame is one page's, and its text past
        // the page was lost (a word processor cuts it off there).
        if (textHeight > this.bottom - this.margin.top) return this.inFlow(flow);
        // As tall as it is set, else as its text.
        const height = Math.min(Number.isFinite(float.height) && float.height! > 0 ? float.height! : textHeight, this.pageH);
        let offset = Number.isFinite(float.y) ? float.y! : 0;
        let verticalFrom = float.verticalFrom;
        // Where it would stand over text already set on this page (measured from the page, or above its anchor),
        // it stands at its anchor instead: the text before it is laid out, and cannot go round it.
        const pageTop = (this.pageMargin ?? this.margin).top;
        const placed = (verticalFrom === 'page' ? 0 : verticalFrom === 'margin' ? pageTop : this.y) + offset;
        if (placed < this.y - 0.5 && placed + height > pageTop && this.y > pageTop + 0.5) { verticalFrom = 'paragraph'; offset = 0; }
        if (verticalFrom === 'paragraph') this.ensureSpace(Math.max(0, offset) + Math.min(height, this.bottom - this.margin.top));
        // Placed across the column (on the page or in the column it may have moved on to), the page's text area or the page.
        const page = this.pageMargin ?? this.margin;
        const across = float.horizontalFrom === 'page' ? { left: 0, width: this.pageW }
            : float.horizontalFrom === 'margin' ? { left: page.left, width: this.pageW - page.left - page.right }
            : { left: this.margin.left, width: this.contentWidth };
        const at = float.align === 'right' ? across.left + across.width - width : float.align === 'center' ? across.left + (across.width - width) / 2
            : float.align === 'left' ? across.left : across.left + (Number.isFinite(float.x) ? float.x! : 0);
        const x = Math.max(0, Math.min(at, this.pageW - width));
        const from = verticalFrom === 'page' ? 0 : verticalFrom === 'margin' ? page.top : this.y;
        const top = Math.max(0, Math.min(from + offset, this.pageH - height));
        const box = { x, y: top, width, height };
        // Drawn as a slide's text box is, on this page alone, with no frame around its own text.
        const saved = { margin: this.margin, fixed: this.fixedPage, y: this.y, floats: this.floats };
        this.fixedPage = true;
        this.floats = [];
        try {
            this.drawShape(box, shape);
            await this.drawTextBox(node.children || [], box, shape);
        } finally {
            this.margin = saved.margin;
            this.fixedPage = saved.fixed;
            this.y = saved.y;
            this.floats = saved.floats;
        }
        if (float.wrap !== 'none') {
            const gap = float.gap ?? { top: 0, right: 6, bottom: 0, left: 6 };
            // (Those of the pages before are done with. A page of more frames than a page holds keeps room
            // for its first ones: each line is measured against every one kept.)
            this.floats = this.floats.filter(f => f.page === this.page);
            if (this.floats.length < MAX_FLOATS_PER_PAGE) this.floats.push({ page: this.page, top: top - gap.top, bottom: top + height + gap.bottom, left: x - gap.left, right: x + width + gap.right, wrap: float.wrap });
        }
    }

    /**
     * Sets a frame's blocks in the flow, and the frames in them too: measured again at each level of frames
     * inside frames, a document of frames nested 1,600 deep took half a minute.
     */
    private framesInFlow = 0;
    private async inFlow(draw: () => Promise<void>): Promise<void> {
        this.framesInFlow++;
        try { await draw(); } finally { this.framesInFlow--; }
    }

    /**
     * Moves below the frames floating on this page that a block of this height starting here would reach (a
     * table, a picture, code, a chart: none of them goes round a frame).
     */
    private belowFloatsFor(height: number): void {
        if (this.furniture || this.fixedPage) return;
        for (const f of this.floats.filter(f => f.page === this.page).sort((a, b) => a.top - b.top)) {
            if (f.bottom > this.y && f.top < this.y + height) this.y = f.bottom;
        }
    }
    /** As belowFloatsFor, for a block measured as `draw` draws it, where a frame on this page is below here. */
    private async belowFloats(draw: () => Promise<void> | void): Promise<void> {
        if (this.furniture || this.fixedPage || !this.floats.some(f => f.page === this.page && f.bottom > this.y)) return;
        this.belowFloatsFor(await this.measured(draw));
    }

    /** Draws a shape's text in its box: within its insets, set from its top, middle or bottom, turned with it. */
    private async drawTextBox(nodes: OfficeContentNode[], box: { x: number; y: number; width: number; height: number }, shape: ShapeMetadata | undefined): Promise<void> {
        const insets = shape?.textInsets ?? { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 };
        const inner = { left: box.x + insets.left, right: this.pageW - (box.x + box.width - insets.right), top: box.y + insets.top, bottom: 0 };
        if (this.pageW - inner.left - inner.right < 4) { inner.left = box.x; inner.right = this.pageW - box.x - Math.max(4, box.width); }
        const drawAll = async () => { for (const n of nodes) await this.render(n); };
        // Measured from the top of a box of its width, then drawn from where its anchor sets it: its
        // first line at the top of the box it is given (a heading adds no space above it there).
        const height = await this.measured(async () => { this.margin = { ...inner, top: 0 }; this.y = 0; await drawAll(); });
        const room = box.height - insets.top - insets.bottom;
        const offset = shape?.textAnchor === 'middle' ? (room - height) / 2 : shape?.textAnchor === 'bottom' ? room - height : 0;
        this.margin = { ...inner, top: inner.top + offset };
        this.y = inner.top + offset;
        const rotation = shape?.rotation || 0;
        if (!rotation && !shape?.flipH && !shape?.flipV) { await drawAll(); return; }
        const cx = box.x + box.width / 2, cy = this.pageH - box.y - box.height / 2;
        const angle = -rotation * Math.PI / 180, cos = Math.cos(angle), sin = Math.sin(angle);
        this.page.pushOperators(this.lib.pushGraphicsState(), this.lib.concatTransformationMatrix(cos, sin, -sin, cos, cx - cos * cx + sin * cy, cy - sin * cx - cos * cy));
        try { await drawAll(); } finally { this.page.pushOperators(this.lib.popGraphicsState()); }
    }

    /**
     * Each section's setup, as its break gave it, the first section's the document's (and a document with
     * no section breaks all one section): its page number format.
     */
    private sectionSetups = new Map<number, PageSetup | undefined>();
    /** The format the page whose header and footer are being drawn numbers its pages in (its section's). */
    private fieldFormat?: PageNumberFormat;
    /** Where the part being drawn saved its page numbers: this page's section, or another's (see fillPageNumber). */
    private fieldSaved: { own: boolean; format?: PageNumberFormat } = { own: true };
    /** The page fields' numbers while a page's header and footer are drawn (see TextMetadata.field). */
    private fieldPage?: number;
    private fieldPages?: number;

    /**
     * Draws each page's header and footer, once every page is laid out (so a page count can be filled
     * in): the section's own, or the one before's where it has none, as Word carries them on; its
     * first page's where it has one, and its even pages' where it has those. They were drawn once,
     * headers before the body and footers after it, each section's and each kind's all together.
     */
    /** The headers or footers drawn on pages: shown ones, loose blocks (a reader giving one list) as one. */
    private furnitureNodes(kind: 'header' | 'footer'): OfficeContentNode[] {
        const nodes = kind === 'header' ? this.ast.auxiliary?.headers : this.ast.auxiliary?.footers;
        const wrapped = (nodes || []).filter(n => n.type === kind);
        const loose = (nodes || []).filter(n => n.type !== kind);
        if (loose.length) wrapped.push({ type: kind, children: loose, metadata: { type: 'default' } } as OfficeContentNode);
        const shown = wrapped.filter(n => !(n.metadata as HeaderFooterMetadata | undefined)?.hidden);
        // A header read from a PDF's page is one of that page's lines: one header a page, of its lines.
        const byPage = new Map<number, OfficeContentNode[]>();
        const kept = shown.filter(n => {
            const page = (n.metadata as HeaderFooterMetadata | undefined)?.page;
            if (page === undefined) return true;
            let lines = byPage.get(page);
            if (!lines) byPage.set(page, lines = []);
            lines.push({ ...n, type: 'paragraph', metadata: undefined });
            return false;
        });
        for (const [page, lines] of byPage) kept.push({ type: kind, children: lines, metadata: { type: 'default', page } } as OfficeContentNode);
        return kept;
    }

    /** The tallest header's and footer's heights (see fitRunningContent). */
    private headerRoom = 0;
    private footerRoom = 0;

    /**
     * Margins with room for the headers and footers below and above the body, as a word processor makes
     * it, and the distances they are drawn at (a third of the margin where the document gives none).
     */
    private withFurnitureRoom(margin: { top: number; right: number; bottom: number; left: number }, headerDistance: number | undefined, footerDistance: number | undefined) {
        const header = headerDistance || Math.max(12, margin.top / 3), footer = footerDistance || Math.max(12, margin.bottom / 3);
        return {
            margin: {
                ...margin,
                top: this.headerRoom ? Math.max(margin.top, header + this.headerRoom + 4) : margin.top,
                bottom: this.footerRoom ? Math.max(margin.bottom, footer + this.footerRoom + 4) : margin.bottom,
            },
            headerDistance: this.headerRoom ? header : headerDistance,
            footerDistance: this.footerRoom ? footer : footerDistance,
        };
    }

    /**
     * Measures the headers and footers before the body is laid out, and keeps the body clear of them: a
     * header taller than the top margin ran into the body, and a footer taller than the bottom one off
     * the page.
     */
    async fitRunningContent(): Promise<void> {
        for (const kind of ['header', 'footer'] as const) {
            let tallest = 0;
            for (const node of this.furnitureNodes(kind)) tallest = Math.max(tallest, await this.measured(async () => { await this.render(node); }));
            if (kind === 'header') this.headerRoom = Math.min(tallest, this.pageH / 3); else this.footerRoom = Math.min(tallest, this.pageH / 3);
        }
        if (!this.headerRoom && !this.footerRoom) return;
        const before = this.margin.top;
        const fit = this.withFurnitureRoom(this.margin, this.headerDistance, this.footerDistance);
        this.margin = fit.margin;
        this.headerDistance = fit.headerDistance;
        this.footerDistance = fit.footerDistance;
        if (this.y <= before + 0.5) this.y = this.margin.top;
        const record = this.pageRecords[this.pageRecords.length - 1];
        if (record) Object.assign(record, { margin: { ...this.margin }, headerDistance: this.headerDistance, footerDistance: this.footerDistance });
    }

    async drawRunningContent(): Promise<void> {
        const headers = this.furnitureNodes('header'), footers = this.furnitureNodes('footer');
        if (!headers.length && !footers.length) return;
        // The part of a type in force in a section: its own, else the latest section's before it (each list's
        // parts by type and section, in order, found once, then found by halving: every page scanning every
        // part took time in the pages times the sections).
        const byType = new Map<OfficeContentNode[], Map<string, { section: number; node: OfficeContentNode }[]>>();
        const pick = (list: OfficeContentNode[], section: number, type: string): OfficeContentNode | undefined => {
            let types = byType.get(list);
            if (!types) {
                types = new Map();
                for (const node of list) {
                    const meta = node.metadata as HeaderFooterMetadata | undefined;
                    const t = meta?.type ?? 'default', at = meta?.section ?? -1;
                    let parts = types.get(t);
                    if (!parts) types.set(t, parts = []);
                    parts.push({ section: at, node });
                }
                // (The first of a section in the list's order, as a scan kept it.)
                for (const parts of types.values()) {
                    parts.sort((a, b) => a.section - b.section);
                    const firsts = parts.filter((p, i) => i === 0 || parts[i - 1].section !== p.section);
                    parts.length = 0; for (const p of firsts) parts.push(p);
                }
                byType.set(list, types);
            }
            const parts = types.get(type);
            if (!parts) return undefined;
            let low = 0, high = parts.length - 1, best: OfficeContentNode | undefined;
            while (low <= high) { const mid = (low + high) >> 1; if (parts[mid].section <= section) { best = parts[mid].node; low = mid + 1; } else high = mid - 1; }
            return best;
        };
        // A section's title page: where its section has a different first page (or, not said, where a
        // first page's part of either kind is the section's own). It shows the first page's parts in force
        // alone, none of a kind with none, as Word shows it.
        const ownFirst = (list: OfficeContentNode[], section: number) => {
            const first = pick(list, section, 'first');
            const own = (first?.metadata as HeaderFooterMetadata | undefined)?.section;
            return !!first && (own === undefined || own === section);
        };
        const titled = (record: typeof this.pageRecords[number]) => record.firstOfSection
            && (record.differentFirstPage ?? (ownFirst(headers, record.section) || ownFirst(footers, record.section)));
        // Even pages' parts anywhere: then an even page shows its even part in force alone, as Word's even and odd pages' parts.
        const facing = [...headers, ...footers].some(n => (n.metadata as HeaderFooterMetadata | undefined)?.type === 'even');
        // A PDF's: that of the page whose content the page holds, else of the nearest page before it with
        // one (a PDF's page laid out on several, a page with none of its own).
        const byPage = new Map<OfficeContentNode[], { page: number; node: OfficeContentNode }[]>();
        const ofPage = (list: OfficeContentNode[], number: number) => {
            let sorted = byPage.get(list);
            if (!sorted) {
                sorted = list.flatMap(node => { const page = (node.metadata as HeaderFooterMetadata | undefined)?.page; return page === undefined ? [] : [{ page, node }]; }).sort((a, b) => a.page - b.page);
                byPage.set(list, sorted);
            }
            let low = 0, high = sorted.length - 1, best: OfficeContentNode | undefined;
            while (low <= high) {
                const mid = (low + high) >> 1;
                if (sorted[mid].page <= number) { best = sorted[mid].node; low = mid + 1; } else high = mid - 1;
            }
            return best;
        };
        const pagedLists = new Map<OfficeContentNode[], boolean>();
        const paged = (list: OfficeContentNode[]) => {
            let known = pagedLists.get(list);
            if (known === undefined) pagedLists.set(list, known = list.some(n => (n.metadata as HeaderFooterMetadata | undefined)?.page !== undefined));
            return known;
        };
        const choose = (list: OfficeContentNode[], record: typeof this.pageRecords[number], number: number) =>
            paged(list) ? ofPage(list, record.sourcePage ?? number)
            : titled(record) ? pick(list, record.section, 'first')
            : pick(list, record.section, facing && number % 2 === 0 ? 'even' : 'default');
        const saved = { page: this.page, y: this.y, w: this.pageW, h: this.pageH, margin: this.margin };
        this.fieldPages = this.pageRecords.length;
        this.fixedPage = true;
        this.furniture = true;
        // Each footer's height at each page width, measured once (a page number's own width aside).
        const footerHeights = new Map<OfficeContentNode, Map<number, number>>();
        try {
            for (const [index, record] of this.pageRecords.entries()) {
                this.fieldPage = index + 1;
                this.fieldFormat = this.sectionSetups.get(record.section)?.pageNumberFormat;
                this.pageW = record.w; this.pageH = record.h; this.margin = record.margin;
                const header = choose(headers, record, index + 1), footer = choose(footers, record, index + 1);
                // Where a part's page numbers were saved: in this page's section, or carried on from another's.
                const savedIn = (part: OfficeContentNode) => {
                    const section = (part.metadata as HeaderFooterMetadata | undefined)?.section;
                    this.fieldSaved = typeof section !== 'number' || section === record.section ? { own: true } : { own: false, format: this.sectionSetups.get(section)?.pageNumberFormat };
                };
                if (header) {
                    savedIn(header);
                    this.page = record.page;
                    this.y = record.headerDistance || Math.max(12, record.margin.top / 3);
                    await this.render(header);
                }
                if (footer) {
                    savedIn(footer);
                    // Measured first, on a page made for it and taken away: it ends at its distance from the bottom.
                    let heights = footerHeights.get(footer);
                    if (!heights) footerHeights.set(footer, heights = new Map());
                    let height = heights.get(record.w);
                    if (height === undefined) heights.set(record.w, height = await this.measured(async () => { await this.render(footer); }));
                    this.page = record.page;
                    // Ending at its distance from the bottom, but never above the body's margin, where the
                    // page's last lines and footnotes are (its height counts the space after its paragraphs).
                    this.y = Math.max(record.margin.top, record.h - record.margin.bottom + 2, record.h - (record.footerDistance || Math.max(12, record.margin.bottom / 3)) - height);
                    await this.render(footer);
                }
            }
        } finally {
            this.fixedPage = false;
            this.furniture = false;
            this.fieldPage = this.fieldPages = undefined;
            this.fieldFormat = undefined;
            this.fieldSaved = { own: true };
            Object.assign(this, { page: saved.page, y: saved.y, pageW: saved.w, pageH: saved.h, margin: saved.margin });
        }
    }

    /**
     * How tall something drawn is: drawn on a page of its own, outside the document's pages, and wiped
     * after, everything it left behind (links, anchors, characters counted as lost) taken back. The one
     * page is kept for every measure: a page added to the document and taken away again cost time in the
     * pages there were, so four thousand sections of their own footers took time in their square.
     */
    private scratchPage: any;
    private async measured(draw: () => Promise<void> | void): Promise<number> {
        const saved = { page: this.page, y: this.y, fixed: this.fixedPage, furniture: this.furniture, links: this.linkBoxes.length, lost: this.lostCharacters, sample: this.lostSample.length, pending: [...this.pendingAnchors], outline: this.outlineEntries.length };
        this.fixedPage = true;
        this.furniture = true;
        if (!this.scratchPage) this.scratchPage = this.lib.PDFPage.create(this.pdf);
        const scratch = this.scratchPage;
        try {
            scratch.setSize(this.pageW, this.pageH);
            this.page = scratch;
            this.y = 0;
            await draw();
            return this.y;
        } finally {
            // What was drawn on it taken out of the file (its contents' objects were still written, and a
            // footer's on every page doubled the file), and the page made ready for the next.
            const context = this.pdf.context;
            const name = this.lib.PDFName.of('Contents');
            const contents = scratch.node.get(name);
            const refs = contents instanceof this.lib.PDFArray ? contents.asArray() : contents ? [contents] : [];
            for (const ref of refs) if (ref instanceof this.lib.PDFRef) context.delete(ref);
            scratch.node.delete(name);
            scratch.contentStream = undefined;
            scratch.contentStreamRef = undefined;
            this.furniture = saved.furniture;
            this.fixedPage = saved.fixed;
            this.page = saved.page;
            this.y = saved.y;
            this.linkBoxes.length = saved.links;
            this.lostCharacters = saved.lost;
            this.lostSample.length = saved.sample;
            this.pendingAnchors = saved.pending;
            this.outlineEntries.length = saved.outline;
        }
    }

    /** The page measures are drawn on, out of the file once the document is drawn (see measured). */
    dropScratchPage(): void {
        if (!this.scratchPage) return;
        this.pdf.context.delete(this.scratchPage.ref);
        this.scratchPage = undefined;
    }

    /** Where each anchor id is (its page and the height on it), for the links into the document. */
    private readonly anchors = new Map<string, { page: any; top: number }>();
    /** Anchor ids met but not placed yet: they take the place of what is drawn next. */
    private pendingAnchors: string[] = [];
    /** The links drawn: a box on a page, and where it goes. */
    private readonly linkBoxes: { page: any; rect: [number, number, number, number]; target: string }[] = [];

    /** Notes a node's anchor ids, placed where the next thing drawn is (see placeAnchors). */
    private markAnchors(node: OfficeContentNode): void {
        // What is measured, or drawn on every page (a header), holds no place a link goes to.
        if (this.config.ignoreInternalLinks || this.furniture) return;
        const ids = (node.metadata as { anchorIds?: unknown } | undefined)?.anchorIds;
        if (Array.isArray(ids)) for (const id of ids) if (typeof id === 'string' && id && !this.anchors.has(id)) this.pendingAnchors.push(id);
    }

    /** Notes the anchor ids of a node and of what it holds (a note's, read before its page is known). */
    private markTreeAnchors(node: OfficeContentNode, depth = 0): void {
        this.markAnchors(node);
        if (depth < 32) for (const child of node.children || []) this.markTreeAnchors(child, depth + 1);
    }

    /** Places the anchors met so far at the cursor: called once the page the next thing goes on is known. */
    placeAnchors(): void {
        if (this.furniture) return;
        for (const id of this.pendingAnchors) if (!this.anchors.has(id)) this.anchors.set(id, { page: this.page, top: this.pageH - this.y });
        this.pendingAnchors = [];
    }

    /** Records a link's box, joined to the box before it when that is the same link on the same line. */
    private addLinkBox(target: string, x: number, bottom: number, width: number, height: number): void {
        const last = this.linkBoxes[this.linkBoxes.length - 1];
        if (last && last.target === target && last.page === this.page && last.rect[1] === bottom && x >= last.rect[2] - 0.01) {
            last.rect[2] = x + width;
            last.rect[3] = Math.max(last.rect[3], bottom + height);
            return;
        }
        this.linkBoxes.push({ page: this.page, rect: [x, bottom, x + width, bottom + height], target });
    }

    /**
     * Makes the links drawn clickable: a web, mail or phone address opens it, a `#id` goes to that anchor
     * in the document. Another target (a relative path, a script) stays drawn as a link that goes nowhere.
     */
    writeLinks(): void {
        this.placeAnchors();
        const context = this.pdf.context;
        for (const box of this.linkBoxes) {
            let action: Record<string, any> | undefined;
            let dest: any[] | undefined;
            if (box.target.startsWith('#')) {
                if (this.config.ignoreInternalLinks) continue;
                let id = box.target.slice(1);
                try { id = decodeURIComponent(id); } catch { /* kept as written */ }
                const anchor = this.anchors.get(id) ?? this.anchors.get(box.target.slice(1));
                if (!anchor) continue;
                dest = [anchor.page.ref, 'XYZ', null, anchor.top, null];
            } else {
                const url = sanitizeOfficePackageUrl(box.target);
                if (!/^(https?|mailto|tel):/i.test(url)) continue;
                // A URI is 7-bit text: anything else, and the string's own delimiters, percent-encoded.
                let ascii = '';
                for (const ch of url) ascii += /^[\x21-\x7E]$/.test(ch) && !'()\\'.includes(ch) ? ch : encodeURIComponent(ch).replace(/[()]/g, c => `%${c.charCodeAt(0).toString(16).toUpperCase()}`);
                action = { Type: 'Action', S: 'URI', URI: this.lib.PDFString.of(ascii) };
            }
            const annotation = context.register(context.obj({
                Type: 'Annot', Subtype: 'Link', Rect: box.rect, Border: [0, 0, 0],
                ...(action ? { A: action } : { Dest: dest }),
            }));
            box.page.node.addAnnot(annotation);
        }
    }

    /** Starts a fresh page and resets the cursor to the top margin. */
    private newPage(): void {
        // A page drawn as it is never runs on to another (see bottom).
        if (this.fixedPage) return;
        this.closeColumnsOnPage();
        // The page's own margins again, for its footnotes and the next page (the text's were a column's).
        if (this.pageMargin) this.margin = this.pageMargin;
        this.drawPageFootnotes();
        this.page = this.appendPage(this.pageW, this.pageH);
        this.y = this.margin.top;
        this.recordPage();
        // Columns run on onto the page, from its top, the first again; balanced ones end with the page before.
        if (this.columns) {
            this.columnIndex = 0;
            this.columnTop = this.columnsBottom = this.y;
            this.columnLimit = null;
            this.applyColumn();
            // A balanced section's last page: its columns end level on what is left of it.
            if (this.balanceArea !== null && this.balanceOnThisPage(Math.max(0, this.balanceArea - this.usedArea))) this.balanceArea = null;
        }
        // Footnotes whose reference ended the page before, without room left there, go on this one.
        const carried = this.carriedFootnotes;
        this.carriedFootnotes = [];
        for (const note of carried) this.reserveFootnote(note, true);
    }

    // ── footnotes ───────────────────────────────────────────────────────────────

    /** Each footnote's and endnote's number as its reference and its text show it (1, 2, 3; i, ii, iii). */
    private readonly noteLabels = new Map<OfficeContentNode, string>();
    private footnoteCount = 0;
    private endnoteCount = 0;
    /** Each footnote's text, ready to draw, and the height it takes (see measureFootnotes). */
    private readonly footnoteRuns = new Map<OfficeContentNode, Run[]>();
    private readonly footnoteHeights = new Map<OfficeContentNode, number>();
    /** The footnotes cited on this page, the room kept for them at its foot, and those for the next. */
    private pageFootnotes: OfficeContentNode[] = [];
    private footReserve = 0;
    private carriedFootnotes: OfficeContentNode[] = [];
    /** The footnotes given a page (the rest are drawn with the endnotes, at the end). */
    private readonly placedNotes = new Set<OfficeContentNode>();
    /** Set while footnotes are measured or drawn: no page is started, and no footnote placed. */
    private fixedPage = false;
    /** Set while something is measured, or a header or footer drawn: no anchor or outline entry is made. */
    private furniture = false;
    /** The rule over a page's footnotes and the space around it. */
    private static readonly FOOTNOTE_RULE = 12;

    /** Whether a note is a footnote (one with no kind counts as one) or an endnote, which get numbers. */
    private static isNumberedNote(note: OfficeContentNode): boolean {
        const kind = (note.metadata as { noteType?: string } | undefined)?.noteType;
        return kind === undefined || kind === 'footnote' || kind === 'endnote';
    }

    /**
     * Each footnote's text as it is drawn at the foot of a page (its number, then its text, at 9pt), and
     * its height there: measured once, on a page made for it and taken away, before the body is laid out.
     */
    async measureFootnotes(): Promise<void> {
        const notes = this.collectedNotes.filter(n => (n.metadata as { noteType?: string } | undefined)?.noteType !== 'endnote' && this.noteLabels.has(n));
        if (!notes.length) return;
        // At the narrowest text any section or sheet has, so a note never takes more room than is kept.
        let width = this.contentWidth;
        if (this.sectionPage) {
            for (const node of this.ast.content) {
                const setup = node.type === 'break' ? (node.metadata as BreakMetadata | undefined)?.pageSetup : node.type === 'sheet' ? (node.metadata as SheetMetadata | undefined)?.pageSetup : undefined;
                if (!setup) continue;
                const page = this.sectionPage(setup);
                width = Math.min(width, Math.max(36, page.w - page.margin.left - page.margin.right));
            }
        }
        for (const note of notes) {
            let runs: Run[] = [];
            const height = await this.measured(async () => {
                runs = [{ text: `${this.noteLabels.get(note)} `, fmt: { size: '9pt', superscript: true } } as Run,
                    ...(await this.collectRuns(note)).map(r => ({ ...r, fmt: { ...r.fmt, size: '9pt' } }))];
                this.drawRuns(runs, this.margin.left, width);
            });
            this.footnoteRuns.set(note, runs);
            this.footnoteHeights.set(note, height + 2);
        }
    }

    /**
     * Keeps room at this page's foot for a footnote cited on it, or, when the page has no more room, at
     * the next page's. A footnote taller than half a page is drawn with the endnotes instead.
     */
    private reserveFootnote(note: OfficeContentNode, fresh = false): void {
        const height = this.footnoteHeights.get(note);
        const body = this.pageH - this.margin.top - this.margin.bottom;
        if (height === undefined || height > body / 2) return;
        const needed = height + (this.pageFootnotes.length ? 0 : NativeLayout.FOOTNOTE_RULE);
        // No more than most of a page's body for its footnotes; past that (or past what is left below the
        // line citing it) a note goes to the next page, and from there on, as long as pages are full.
        const full = this.footReserve + needed > body * 0.6;
        // In columns, the note goes under the lowest line any column on the page has reached: the columns
        // before this one are drawn down to there already.
        const lowest = this.columns ? Math.max(this.y, this.columnsBottom) : this.y;
        const pageBottom = this.pageH - (this.pageMargin ?? this.margin).bottom - this.footReserve;
        if (full || (!fresh && (this.y + needed > this.bottom || lowest + needed > pageBottom))) { this.carriedFootnotes.push(note); this.placedNotes.add(note); return; }
        this.footReserve += needed;
        this.pageFootnotes.push(note);
        this.placedNotes.add(note);
    }

    /** A footnote cited in a line just drawn: given a page, once. */
    private citeFootnote(note: OfficeContentNode): void {
        if (this.fixedPage || this.placedNotes.has(note) || !this.footnoteRuns.has(note)) return;
        this.reserveFootnote(note);
    }

    /** Draws this page's footnotes in the room kept at its foot, under a short rule. */
    private drawPageFootnotes(): void {
        if (!this.pageFootnotes.length || !this.page) return;
        const notes = this.pageFootnotes, reserve = this.footReserve;
        this.pageFootnotes = [];
        this.footReserve = 0;
        this.fixedPage = true;
        try {
            this.y = this.pageH - this.margin.bottom - reserve + NativeLayout.FOOTNOTE_RULE / 2;
            const yy = this.pageH - this.y;
            this.page.drawLine({ start: { x: this.margin.left, y: yy }, end: { x: this.margin.left + Math.min(180, this.contentWidth), y: yy }, thickness: 0.5, color: this.lib.rgb(0.7, 0.7, 0.7) });
            this.y += NativeLayout.FOOTNOTE_RULE / 2;
            for (const note of notes) {
                // The anchors in the note (its runs were read before the pages were known) are here.
                this.markTreeAnchors(note);
                this.drawRuns(this.footnoteRuns.get(note)!, this.margin.left, this.contentWidth);
                this.y += 2;
            }
        } finally {
            this.fixedPage = false;
        }
    }

    /**
     * Ensures `h` points of vertical space remain, starting a new page if not. Never starts a page
     * when the cursor is already at the top margin: the block does not fit a whole page, so a blank
     * page in front of it would not help and would just pad the output. Callers that must not lose
     * content taller than a page (tables) split it across pages themselves.
     */
    private ensureSpace(h: number): void {
        if (this.y + h > this.bottom && this.y > this.margin.top && !this.fixedPage) this.nextColumnOrPage();
        if (this.pendingAnchors.length) this.placeAnchors();
    }

    /**
     * Picks the style matching a run's formatting: its font's kind (monospaced for code, serif or sans),
     * then its bold and italic.
     */
    private faceFor(fmt: TextFormatting | undefined, mono: boolean): FaceKey {
        const kind = mono ? 'mono' : fontKind(fmt?.font);
        const bold = !!fmt?.bold, italic = !!fmt?.italic;
        const variant = bold && italic ? 'BoldItalic' : bold ? 'Bold' : italic ? 'Italic' : '';
        if (kind === 'sans') return (variant ? variant[0].toLowerCase() + variant.slice(1) : 'regular') as FaceKey;
        return `${kind}${variant}` as FaceKey;
    }

    /**
     * A CSS colour (`#rrggbb`, `#rgb`, a name, `rgb()`, `hsl()`) as a pdf-lib rgb colour, or null. Read as
     * hex alone, an HTML document's named and `rgb()` colours were dropped.
     */
    private color(value: string | undefined): any {
        const h = cssColorHex(value);
        if (!h) return null;
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
     * character none of the style's fonts draws is a `?`, counted for the one warning
     * {@link reportLostCharacters} gives, unless `count` is false (text measured, not drawn). A combining
     * mark the font of the character before it has stays, drawn with it (see segments).
     */
    private enc(text: string, key: FaceKey = 'regular', count = true): string {
        const chain = this.fonts[key];
        let out = '';
        // (The font the character last written is drawn in, as segments finds it; none after a space, where
        // text is cut into words, so a mark there is one standing alone.)
        let previous: Face | undefined;
        const write = (s: string) => { out += s; previous = WHITESPACE.test(s) ? undefined : chain.find(face => face.covers(s.codePointAt(0)!)); };
        for (const ch of text) {
            const cp = ch.codePointAt(0)!;
            if (cp === 0x09) write('    ');
            else if (cp === 0x0A || cp === 0x0D) write(' ');
            else if (cp === 0xAD) { /* a soft hyphen shows only where a line is broken at it, and none is */ }
            else if ((isInvisible(cp) || COMBINING_MARK.test(ch)) && previous?.has(cp)) out += ch;
            else if (chain.some(face => face.covers(cp))) write(ch);
            else if (isInvisible(cp)) { /* draws nothing: no loss */ }
            // A space of another width (an em space, a thin space, an ideographic space) or a line
            // separator the fonts lack is a space; any other control character draws nothing.
            else if (OTHER_SPACE.test(ch)) write(' ');
            else if (CONTROL.test(ch)) { /* no glyph: no loss */ }
            else {
                write('?');
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
        const feature = this.customFonts ? 'characters none of the given fonts draws (native engine)' : 'non-WinAnsi characters (native engine)';
        this.reportWarning(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'pdf', feature, reason: 'unencodable-text', count: this.lostCharacters, sample: this.lostSample.join('') });
    }

    /**
     * Splits drawable text into stretches of one font each: every character goes to the first font of
     * the style that draws it. A combining mark or a joiner stays with the font of the character before
     * it when that font has it (see Face.has), since a mark drawn in another font is not attached to its
     * letter.
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
            else if (current && attached && current.face.has(cp)) face = current.face;
            // (A mark enc kept with its letter, here at the start of a stretch cut from it: in a font that has it.)
            else face = chain.find(f => f.covers(cp)) ?? (attached ? chain.find(f => f.has(cp)) : undefined);
            let drawn = ch;
            if (!face) { if (isInvisible(cp)) continue; face = chain[chain.length - 1]; drawn = '?'; }
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
        for (const seg of this.segments(text, key)) w += this.segmentWidth(seg, size);
        return w;
    }
    /** A stretch of one font's width: as shaped, where its font shapes it (see shapedRun). */
    private segmentWidth(seg: { text: string; face: Face }, size: number): number {
        const run = this.shapedRun(seg.face, seg.text);
        if (!run) return seg.face.font.widthOfTextAtSize(seg.text, size);
        let units = 0;
        for (const p of run.positions) units += p.xAdvance;
        return units * size / run.unitsPerEm;
    }

    /** Shaped runs by font and text (see shapedRun): a word is measured several times as lines are broken. */
    private shapedRuns = new Map<unknown, Map<string, ShapedRun | null>>();
    /**
     * `text` as a given font shapes it, where it is left-to-right text such a font shapes (SHAPED_TEXT): its
     * glyphs and their positions, the marks' offsets and advances among them, which pdf-lib leaves out (it
     * draws each glyph at its own advance). Null for other text, and for the Standard-14 fonts.
     */
    private shapedRun(face: Face, text: string): ShapedRun | null {
        const shaper = face.code ? face.font.embedder?.font : undefined;
        if (!shaper?.layout || !SHAPED_TEXT.test(text) || directionOf(text) === 'rtl') return null;
        let runs = this.shapedRuns.get(face.font);
        if (!runs) this.shapedRuns.set(face.font, runs = new Map());
        let run = runs.get(text);
        if (run === undefined) {
            const laid = shaper.layout(text, face.font.embedder.fontFeatures);
            run = { glyphs: laid.glyphs, positions: laid.positions, unitsPerEm: shaper.unitsPerEm || 1000 };
            if (runs.size >= 4096) runs.clear();
            runs.set(text, run);
        }
        return run;
    }

    /**
     * Draws a shaped run (see shapedRun) with each glyph where its font places it, and marks it as the text
     * it stands for (`/ActualText`): copied, a word reads in the order it is written, where the glyphs are
     * drawn in another (Hindi's ि before the letter it follows) or one glyph stands for several letters.
     * One string of glyphs, each moved from where the last one's width leaves off (`TJ`), so a reader
     * measures the word as drawn: drawn one by one, a mark of no advance that the font gives a width made
     * the word as long as that, and the space after it was lost.
     */
    private drawShaped(face: Face, run: ShapedRun, text: string, x: number, y: number, size: number, color: any): void {
        const lib = this.lib;
        const context = this.page.doc.context;
        this.useFont(face.font);
        const fontKey = (this.page as { fontKey?: unknown }).fontKey;
        face.font.modified = true;
        const thousandths = 1000 / run.unitsPerEm;
        const op = (name: string, ...args: any[]) => lib.PDFOperator.of(name, args);
        // (The span within the text object: closed after it, a reader took its text to stand where none does.)
        const ops = [
            lib.pushGraphicsState(), lib.setFillingColor(color), lib.beginText(), lib.setFontAndSize(fontKey, size),
            op('BDC', lib.PDFName.of('Span'), context.obj({ ActualText: lib.PDFHexString.fromText(text) })),
            lib.setTextMatrix(1, 0, 0, 1, x, y),
        ];
        let shown: any[] = [], pen = 0, cursor = 0, rise = 0;
        const flush = () => { if (shown.length) ops.push(op('TJ', context.obj(shown))); shown = []; };
        run.glyphs.forEach((glyph, i) => {
            const p = run.positions[i];
            // A mark raised or lowered: its rise, in a string of its own.
            if (p.yOffset !== rise) { flush(); rise = p.yOffset; ops.push(op('Ts', lib.PDFNumber.of(rise * size / run.unitsPerEm))); }
            const to = pen + p.xOffset;
            if (Math.abs(to - cursor) > 0.01) shown.push(-(to - cursor) * thousandths);
            shown.push(lib.PDFHexString.of(face.code!(glyph).toString(16).padStart(4, '0')));
            cursor = to + glyph.advanceWidth;
            pen += p.xAdvance;
        });
        flush();
        if (rise) ops.push(op('Ts', lib.PDFNumber.of(0)));
        ops.push(op('EMC'), lib.endText(), lib.popGraphicsState());
        this.page.pushOperators(...ops);
    }

    /**
     * Draws drawable text with its left edge at `x`, each stretch in its font. A right-to-left word's stretches are drawn last first.
     * Invisible text (TextFormatting.invisible) is drawn invisibly, in text rendering mode 3: where it would stand, as wide, there to
     * select and search, as a scan's recognised words lie over its picture.
     */
    private drawText(text: string, key: FaceKey, x: number, y: number, size: number, color: any, invisible = false): void {
        const segs = this.segments(text, key);
        if (segs.length > 1 && directionOf(text) === 'rtl') segs.reverse();
        // (The mode is the graphics state's, set around the text and its fallback drawing alike.)
        if (invisible) this.page.pushOperators(this.lib.pushGraphicsState(), this.lib.setTextRenderingMode(this.lib.TextRenderingMode.Invisible));
        try {
            for (const seg of segs) {
                const run = this.shapedRun(seg.face, seg.text);
                if (run) this.drawShaped(seg.face, run, seg.text, x, y, size, color);
                else {
                    this.useFont(seg.face.font);
                    this.page.drawText(seg.text, { x, y, size, color });
                }
                x += this.segmentWidth(seg, size);
            }
        } finally { if (invisible) this.page.pushOperators(this.lib.popGraphicsState()); }
    }

    /** Each page's resource name for each font it draws in, made once. */
    private readonly fontKeys = new WeakMap<object, Map<object, unknown>>();
    /**
     * Sets the page's font for the text drawn next. pdf-lib names a font anew in the page's resources each
     * time text is drawn with a font of its own, each name found by looking through all the names before:
     * a picture of 50,000 spans of text took time in their square.
     */
    private useFont(font: any): void {
        const page = this.page as { font?: unknown; fontKey?: unknown; setFont(font: unknown): void };
        if (page.font === font && page.fontKey) return;
        let keys = this.fontKeys.get(page);
        if (!keys) this.fontKeys.set(page, keys = new Map());
        const key = keys.get(font);
        if (key) { page.font = font; page.fontKey = key; return; }
        page.setFont(font);
        keys.set(font, page.fontKey);
    }

    /**
     * Draws one line of drawable text (a table cell's line, a line of code) inside `[x, x + width]`,
     * its words in the order they are seen: a right-to-left line starts at the right edge.
     */
    private drawLine(text: string, key: FaceKey, x: number, y: number, size: number, color: any, width: number, invisible = false): void {
        const dir = directionOf(text);
        if (dir !== 'rtl') { this.drawText(text, key, x, y, size, color, invisible); return; }
        const baseRtl = startsRtl(text);
        const items = text.split(/(\s+)/).filter(t => t !== '').map(t => ({ text: t, w: this.widthOf(t, key, size), dir: /^\s+$/.test(t) ? 'neutral' as const : directionOf(t) }));
        let cx = baseRtl ? x + Math.max(0, width - items.reduce((s, it) => s + it.w, 0)) : x;
        for (const it of visualOrder(items, baseRtl)) {
            if (it.text.trim()) this.drawText(it.text, key, cx, y, size, color, invisible);
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
    private async collectRuns(node: OfficeContentNode): Promise<Run[]> {
        const out: Run[] = [];
        // The runs of the page field being read: a field's result saved in several runs (its digits in two
        // looks) is one number, written once (each run drew the whole number: "Page 11 of 33"), in the runs
        // as a whole (see fillPageNumber).
        let field: { kind: string; format?: PageNumberFormat; runs: Run[] } | undefined;
        const endField = () => {
            if (!field) return;
            const value = field.kind === 'page' ? this.fieldPage! : this.fieldPages!;
            // (A field's own format writes its saved result as its own; the section's, from where it was saved.)
            const filledTexts = field.format ? fillPageNumber(field.runs.map(r => r.text), value, field.format)
                : fillPageNumber(field.runs.map(r => r.text), value, field.kind === 'page' ? this.fieldFormat : undefined, field.kind === 'page' ? this.fieldSaved : { own: true });
            field.runs.forEach((r, i) => { r.text = filledTexts[i]; });
            field = undefined;
        };
        const walk = async (n: OfficeContentNode, inherited: TextFormatting, inheritedLink: string | undefined, isRoot: boolean): Promise<void> => {
            let text = n.text;
            let override: string | false | void = undefined;
            // A picture drawn in its cell, or a chart after its table, is no text of the cell's (see table),
            // but for a picture's recognized text where that is drawn too.
            if (!isRoot && this.drawnApart.has(n)) { const ocr = this.ocrBesidePicture(n); if (ocr) out.push({ text: ocr, fmt: inherited }); return; }
            if (!isRoot) {
                override = await this.onNodeValue(n);
                if (override === false) return;
                if (typeof override === 'string') text = override;
            }
            const fmt = { ...inherited, ...(n.formatting || {}) };
            // A page field's number for the page being drawn (its header or footer), not the one saved.
            const meta = n.metadata as { field?: string; fieldFormat?: PageNumberFormat; link?: unknown } | undefined;
            const kind = meta?.field;
            const filled = typeof override !== 'string' && ((kind === 'page' && this.fieldPage !== undefined) || (kind === 'pages' && this.fieldPages !== undefined));
            // (A field's runs are the same field's text runs in a row: anything else between, a picture, a
            // break, another field, ends it; an empty run that draws nothing does not.)
            const sameField = filled && n.type === 'text' && kind === field?.kind;
            if (!isRoot && !sameField && !(n.type === 'text' && !text && !n.notes?.length)) endField();
            if (filled && n.type === 'text') {
                const own = meta?.link;
                const run: Run = { text: text ?? '', fmt, link: typeof own === 'string' && own ? own : inheritedLink };
                field ??= { kind: kind!, format: meta?.fieldFormat, runs: [] };
                field.runs.push(run);
                out.push(run);
                if (!isRoot) this.markAnchors(n);
                cite(n, fmt);
                return;
            }
            // A link's target, which the words inside it share: a link around formatted runs linked only
            // the run that carried it.
            const own = (n.metadata as { link?: unknown } | undefined)?.link;
            const link = typeof own === 'string' && own ? own : inheritedLink;
            if (!isRoot) this.markAnchors(n);
            // A picture in a line the engine draws as text (a list item, heading or note; a paragraph
            // draws its pictures itself) is its alt or recognized text: it was dropped.
            if (n.type === 'image' && !isRoot) {
                // A picture in the line, where it is drawn there (see paragraph).
                const picture = this.inlinePictures.get(n);
                if (picture && typeof override !== 'string') { out.push({ text: '', fmt, link, picture }); return; }
                const alt = this.inlineImageText(n, typeof override === 'string' ? override : undefined);
                if (alt) out.push({ text: alt, fmt, link });
                return;
            }
            // A break inside a line (`<br>`, Word's `w:br`, a column break) starts a new line, and a page
            // break inside a paragraph a new page; they were dropped, gluing the lines together.
            if (n.type === 'break' && !isRoot && typeof override !== 'string') {
                const kind = (n.metadata as any)?.breakType;
                if (kind === 'page') out.push({ text: '\f', fmt });
                else if (kind !== 'lastRenderedPage' && kind !== 'thematic') out.push({ text: '\n', fmt });
                return;
            }
            if (n.type === 'text' || (!n.children?.length && text)) {
                if (text) out.push({ text, fmt, link });
                cite(n, fmt);
                return;
            }
            for (const c of n.children || []) await walk(c, fmt, link, false);
            if (!isRoot) cite(n, fmt);
        };
        // A footnote's or endnote's number where it is cited, raised (it was drawn nowhere in the text).
        const cite = (n: OfficeContentNode, fmt: TextFormatting) => {
            for (const note of n.notes || []) {
                const label = this.noteLabels.get(note);
                if (label) out.push({ text: label, fmt: { ...fmt, superscript: true, subscript: false }, note });
            }
        };
        await walk(node, {}, undefined, true);
        endField();
        if (!out.length && node.text) out.push({ text: node.text, fmt: node.formatting || {} });
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
                else if (this.drawnApart.has(c)) t = this.ocrBesidePicture(c);
                else if (c.type === 'image') t = this.cellImageText(c);
                else t = (await this.collectRuns(c)).map(r => r.note ? this.raisedLabel(r.text) : r.text).join('').replace(/\f/g, '\n');
                if (t !== '') lines.push(t);
            }
            if (lines.length) return lines.join('\n');
        }
        // A page break inside a cell is a line break there: a cell cannot start a page.
        const flat = (await this.collectRuns(node)).map(r => r.note ? this.raisedLabel(r.text) : r.text).join('').replace(/\f/g, '\n');
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

    /**
     * What in a table's cells is drawn apart from their text: pictures, in their cells; charts, after
     * the table (a chart in a sheet's cell printed as its data's text, in the cell).
     */
    private drawnApart = new Set<OfficeContentNode>();

    /** A picture's recognized text, where the image mode draws it beside the picture ('image+ocr-text'). */
    private ocrBesidePicture(node: OfficeContentNode): string {
        return node.type === 'image' && this.imageMode === 'image+ocr-text' ? (node.text || '').trim() : '';
    }

    /**
     * A footnote's number in a cell's text, where formatting is not drawn: in superscript digits the font
     * has, else in brackets. It was glued to the text as plain digits ("100" and note 1 read "1001").
     */
    private raisedLabel(label: string): string {
        const raised = [...label].map(ch => '⁰¹²³⁴⁵⁶⁷⁸⁹'[Number(ch)] ?? '').join('');
        return raised.length === [...label].length && [...raised].every(ch => this.fonts.regular.some(face => face.covers(ch.codePointAt(0)!))) ? raised : `[${label}]`;
    }

    /** The footnotes a table cell's text cites (they are drawn at the page's foot when its row is drawn). */
    private notesCitedIn(cell: OfficeContentNode): OfficeContentNode[] {
        const found: OfficeContentNode[] = [];
        const walk = (n: OfficeContentNode, depth: number) => {
            if (depth > 16 || found.length > 64) return;
            for (const note of n.notes || []) if (this.noteLabels.has(note) && !found.includes(note)) found.push(note);
            for (const child of n.children || []) if (child.type !== 'table') walk(child, depth + 1);
        };
        walk(cell, 0);
        return found;
    }

    /** A cell's charts (its own and its blocks'), drawn after its table. */
    private chartsIn(cell: OfficeContentNode): OfficeContentNode[] {
        if (this.config.includeCharts === false) return [];
        const found: OfficeContentNode[] = [];
        const walk = (n: OfficeContentNode, depth: number) => {
            if (found.length >= 16 || depth > 6 || n.type === 'table') return;
            if (n.type === 'chart') { found.push(n); return; }
            for (const c of n.children || []) walk(c, depth + 1);
        };
        for (const c of cell.children || []) walk(c, 0);
        return found;
    }

    /** A cell's pictures, where pictures are drawn (the image mode embeds them): its own and its blocks'. */
    private picturesIn(cell: OfficeContentNode): OfficeContentNode[] {
        if (this.imageMode !== 'image-only' && this.imageMode !== 'image+ocr-text') return [];
        const found: OfficeContentNode[] = [];
        const walk = (n: OfficeContentNode, depth: number) => {
            if (found.length >= 16 || depth > 6) return;
            if (n.type === 'image') { found.push(n); return; }
            // A nested table's pictures are its own cells'.
            if (n.type === 'table') return;
            for (const c of n.children || []) walk(c, depth + 1);
        };
        for (const c of cell.children || []) walk(c, 0);
        return found;
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
    private drawRuns(runs: Run[], left: number, width: number, mono = false, layout: LineLayout = {}): void {
        // The column's left edge and width the paragraph starts at (see flush).
        let columnLeft = this.margin.left, columnWidth = this.contentWidth;
        // Expand runs into per-word pieces (a trailing-space flag lets us re-insert inter-word gaps).
        // A piece is a word, or (`brk`) a forced break: a new line, or a new page.
        const pieces: { p: Piece; space: boolean; brk?: 'line' | 'page' }[] = [];
        // A paragraph whose first letter is a right-to-left one is read from the right edge.
        const baseRtl = startsRtl(runs.map(r => r.text).join(''));
        for (const run of runs) {
            const face = this.faceFor(run.fmt, mono);
            // A superscript or subscript is smaller and raised or lowered (they were drawn as body text).
            const baseSize = parseFontSize(run.fmt.size) ?? (mono ? 10 : 11);
            const shifted = !!(run.fmt.superscript || run.fmt.subscript);
            const size = shifted ? baseSize * 0.65 : baseSize;
            const rise = run.fmt.superscript ? baseSize * 0.35 : run.fmt.subscript ? -baseSize * 0.12 : 0;
            const color = run.link ? this.lib.rgb(0.1, 0.32, 0.72) : (this.color(run.fmt.color) ?? this.lib.rgb(0.12, 0.12, 0.12));
            // (Invisible text has no line under it and no background: it is not seen.)
            const invisible = !!run.fmt.invisible;
            const underline = !invisible && !!(run.fmt.underline || run.link);
            // A highlight or a run's own background, drawn behind its words.
            const background = invisible ? undefined : cssColorHex(run.fmt.backgroundColor) ?? undefined;
            // A picture in the line: as wide and tall as it is drawn, no wider than the line, no taller
            // than most of a page.
            if (run.picture) {
                const scale = Math.min(1, width / run.picture.w, (this.bottom - this.margin.top) * 0.8 / run.picture.h);
                pieces.push({ p: { text: '', face, size, color, underline: false, link: run.link, picture: { img: run.picture.img, w: run.picture.w * scale, h: run.picture.h * scale } }, space: false });
                continue;
            }
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
                        pieces.push({ p: { text: chunk, face, size, color, underline, link: run.link, background, rise, note: run.note, ...(invisible ? { invisible } : {}) }, space: false });
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
        // The paragraph's first line starts `firstLine` points in from the others (out, when negative).
        let firstOfParagraph = true;
        const lineStart = () => firstOfParagraph ? (layout.firstLine || 0) : 0;
        // A line beside a frame floating on the page (see floatingFrame) starts and ends short of it: the room
        // it keeps at its left and right, found as its first word is placed.
        let besideLeft = 0, besideRight = 0, besidePage: unknown, besideColumn = columnLeft;
        const lineRoom = () => width - lineStart() - besideLeft - besideRight;

        /** Draws the line gathered; `last` when it ends the paragraph or a forced break ends it. */
        /** The distance from a line of text this size to the next, as the paragraph spaces its lines. */
        const lineAdvance = (size: number) => {
            const natural = size * 1.35;
            if (layout.lineMultiple) return natural * layout.lineMultiple;
            if (layout.lineHeight) return layout.lineAtLeast ? Math.max(natural, layout.lineHeight) : layout.lineHeight;
            return natural;
        };

        /**
         * The room a line of text this size starting here has beside the frames floating on the page: on the
         * side of each with more room, or, where one leaves too little or the text goes above and below it,
         * below it.
         */
        const fitBeside = (size: number) => {
            besideLeft = besideRight = 0;
            besidePage = this.page;
            besideColumn = columnLeft;
            // (A page's footnotes, a frame's or a slide's text: none goes round a frame.)
            if (this.furniture || this.fixedPage || !this.floats.length) return;
            const lh = lineAdvance(size);
            for (let guard = 0; guard < 64; guard++) {
                const from = left + lineStart(), to = left + width;
                let l = from, r = to, below = -Infinity, blocked = false;
                for (const f of this.floats) {
                    if (f.page !== this.page || f.top >= this.y + lh || f.bottom <= this.y || f.right <= from || f.left >= to) continue;
                    below = Math.max(below, f.bottom);
                    if (f.wrap === 'aboveBelow') { blocked = true; continue; }
                    if (f.left - from >= to - f.right) r = Math.min(r, f.left); else l = Math.max(l, f.right);
                }
                if (!blocked && r - l >= Math.min(36, (to - from) / 3)) { besideLeft = l - from; besideRight = to - r; return; }
                if (!(below > this.y)) return;
                this.y = below;
            }
        };

        /** Whether the text has run on into a column of another place or width: its lines then follow it. */
        const followColumn = (): boolean => {
            if (this.margin.left === columnLeft && this.contentWidth === columnWidth) return false;
            left += this.margin.left - columnLeft;
            width = Math.max(1, width + this.contentWidth - columnWidth);
            columnLeft = this.margin.left;
            columnWidth = this.contentWidth;
            return true;
        };

        const flush = (last: boolean) => {
            if (!line.length) return;
            // A picture in the line stands on its baseline, and the line is as tall as it needs.
            const tallest = maxOf(line.map(it => it.p.picture?.h ?? 0), 0);
            const textSize = maxOf(line.filter(it => !it.p.picture).map(it => it.p.size), tallest ? 0 : 11);
            const lh = Math.max(lineAdvance(textSize || 11), tallest + (textSize || 11) * 0.35);
            this.ensureSpace(lh);
            // Run on into another column: its lines move with it, as wide as that column lets them (its
            // indents kept), and the words of this line that no longer fit go on to the next.
            let leftover: typeof line = [];
            let moved = followColumn();
            // On another page or column, the frames beside it are that one's; below one past the foot, on to
            // the next column or page.
            if (moved || besidePage !== this.page || besideColumn !== columnLeft) {
                fitBeside(textSize || 11);
                for (let i = 0; i < 4 && this.y + lh > this.bottom && this.y > this.margin.top && !this.fixedPage; i++) {
                    this.ensureSpace(lh);
                    moved = followColumn() || moved;
                    fitBeside(textSize || 11);
                }
            }
            if (moved || besideLeft + besideRight > 0) {
                let used = 0, k = 0;
                for (; k < line.length; k++) {
                    const add = (k && line[k - 1].space ? spaceWidth(line[k - 1].p) : 0) + line[k].w;
                    if (k && used + add > lineRoom()) break;
                    used += add;
                }
                if (k < line.length) { leftover = line.slice(k); line = line.slice(0, k); last = false; }
            }
            const baseline = this.pageH - this.y - Math.max(textSize, tallest);
            // Words and the gaps between them, in reading order, then in the order they are seen.
            const items: { p?: Piece; w: number; dir: 'ltr' | 'rtl' | 'neutral' }[] = [];
            line.forEach((it, i) => {
                items.push({ p: it.p, w: it.w, dir: directionOf(it.p.text) });
                if (it.space && i < line.length - 1) items.push({ w: spaceWidth(it.p), dir: 'neutral' });
            });
            const room = lineRoom();
            const used = items.reduce((s, it) => s + it.w, 0);
            const free = Math.max(0, room - used);
            // A justified line's spaces share what is left, but for its last line (as a word processor sets it).
            const gaps = items.filter(it => !it.p).length;
            const stretch = layout.align === 'justify' && !last && gaps ? free / gaps : 0;
            const align = layout.align === 'center' || layout.align === 'right' ? layout.align : baseRtl ? 'right' : 'left';
            // (A justified line fills its width: its spaces take what is free, so it starts at the left.)
            let x = left + lineStart() + besideLeft + (stretch ? 0 : align === 'center' ? free / 2 : align === 'right' ? free : 0);
            // Backgrounds first, under the words: a word's and the gap after it, where the next word has
            // the same one, so a highlighted phrase is one band.
            const ordered = visualOrder(items, baseRtl);
            let bx = x;
            ordered.forEach((it, i) => {
                const w = it.w + (it.p ? 0 : stretch);
                const bg = it.p?.background ?? (!it.p && ordered[i - 1]?.p?.background && ordered[i - 1]?.p?.background === ordered[i + 1]?.p?.background ? ordered[i - 1].p!.background : undefined);
                const size = it.p?.size ?? ordered[i - 1]?.p?.size ?? 11;
                if (bg) this.page.drawRectangle({ x: bx, y: baseline - size * 0.25, width: w, height: size * 1.15, color: this.color(bg) });
                bx += w;
            });
            for (const it of ordered) {
                if (it.p?.picture) {
                    this.drawPicture(it.p.picture.img, { x, y: baseline, width: it.p.picture.w, height: it.p.picture.h });
                    if (it.p.link) this.addLinkBox(it.p.link, x, baseline, it.p.picture.w, it.p.picture.h);
                } else if (it.p) {
                    this.drawText(it.p.text, it.p.face, x, baseline + (it.p.rise || 0), it.p.size, it.p.color, it.p.invisible);
                    if (it.p.underline) {
                        this.page.drawLine({ start: { x, y: baseline - 1.5 }, end: { x: x + it.w, y: baseline - 1.5 }, thickness: 0.5, color: it.p.color });
                    }
                    if (it.p.link) this.addLinkBox(it.p.link, x, baseline - it.p.size * 0.25, it.w, it.p.size * 1.15);
                }
                x += it.w + (it.p ? 0 : stretch);
            }
            this.y += lh;
            // The footnotes this line cites, given room at this page's foot.
            for (const it of line) if (it.p.note) this.citeFootnote(it.p.note);
            line = [];
            lineWidth = 0;
            firstOfParagraph = false;
            if (leftover.length) {
                for (const it of leftover) place(it.p, it.space, it.w);
                if (last) flush(true);
            }
        };
        /** Adds a word to the line, drawing the line first where the word does not fit after it. */
        const place = (p: Piece, space: boolean, w: number) => {
            // Only the previous piece's OWN trailing space counts toward width: a word split across two
            // runs (or a splitToWidth chunk) has space:false and draws with no gap, so measuring one
            // there wrapped such lines slightly early and disagreed with what flush() draws.
            // Again after each line drawn: a line drawn may run on into a narrower column, the words that no
            // longer fit it put back on the next, and this word must fit after them too (it overran the column).
            for (let prev = line[line.length - 1]; prev && lineWidth + (prev.space ? spaceWidth(prev.p) : 0) + w > lineRoom(); prev = line[line.length - 1]) flush(false);
            // A new line's room beside the frames floating on the page.
            if (!line.length) fitBeside(p.size);
            // A word wider than the column it has run on into (one cut to the paragraph's first column's
            // width): cut again to this one's, each piece a word of its own.
            if (!line.length && !p.picture && w > lineRoom() + 0.5 && p.text.length > 1) {
                const chunks = splitToWidth(p.text, t => this.widthOf(t, p.face, p.size), lineRoom());
                if (chunks.length > 1) {
                    chunks.forEach((text, i) => place({ ...p, text, ...(i < chunks.length - 1 ? { note: undefined } : {}) }, i === chunks.length - 1 ? space : false, this.widthOf(text, p.face, p.size)));
                    return;
                }
            }
            const prevAfter = line[line.length - 1];
            lineWidth += (prevAfter && prevAfter.space ? spaceWidth(prevAfter.p) : 0) + w;
            line.push({ p, space, w });
        };

        for (const piece of pieces) {
            if (piece.brk) {
                // A line ended here: the next starts below it. Two in a row leave an empty line.
                if (line.length) flush(true);
                else { const lh = lineAdvance(piece.p.size); this.ensureSpace(lh); followColumn(); this.y += lh; firstOfParagraph = false; }
                if (piece.brk === 'page') this.newPage();
                continue;
            }
            place(piece.p, piece.space, piece.p.picture ? piece.p.picture.w : this.widthOf(piece.p.text, piece.p.face, piece.p.size));
        }
        flush(true);
    }

    /** Space above a paragraph, as its document sets it: none at the top of a page, as a word processor sets it. */
    private spaceAbove(points: number | undefined): void {
        if (points && this.y > this.margin.top) this.y += Math.min(points, (this.bottom - this.margin.top) / 2);
    }

    /** Space below a paragraph, as its document sets it, or the engine's own. */
    private spaceBelow(points: number | undefined, otherwise: number): number {
        return points !== undefined ? Math.min(points, (this.bottom - this.margin.top) / 2) : otherwise;
    }

    /**
     * Where a paragraph or heading sits: its left edge and width within the page's text column, from its
     * indents (in twips, as Word gives them; out into the margin when negative), and its lines' layout.
     */
    private placement(node: OfficeContentNode, indentLeft = 0): { left: number; width: number; layout: LineLayout; spacing?: ParagraphSpacing } {
        const meta = node.metadata as { alignment?: TextAlignment; paragraphIndentation?: IndentationMetadata } | undefined;
        const pt = (twips: unknown) => typeof twips === 'number' && Number.isFinite(twips) ? twips / 20 : 0;
        const ind = meta?.paragraphIndentation;
        let left = this.margin.left + indentLeft + pt(ind?.left);
        let right = this.margin.left + this.contentWidth - pt(ind?.right);
        // Never past the page's edges, and never narrower than an inch: an indent that leaves less is
        // the source's own mistake, or a unit this engine does not read.
        left = Math.max(0, left);
        right = Math.min(this.pageW, right);
        if (right - left < 72) { left = this.margin.left + indentLeft; right = this.margin.left + this.contentWidth; }
        const firstLine = ind?.firstLine ? pt(ind.firstLine) : ind?.hanging ? -pt(ind.hanging) : 0;
        const width = right - left;
        const align = meta?.alignment && ['left', 'center', 'right', 'justify'].includes(meta.alignment) ? meta.alignment : undefined;
        const spacing = writableSpacing(meta);
        // A first line set in at most to leave an inch (none in a column narrower than that), and set out at
        // most to the page's edge, or in columns to the column's own: it ran into the gutter and over the
        // column before.
        const outmost = (this.columns?.count ?? 1) > 1 ? Math.min(0, this.margin.left - left) : -left;
        const layout: LineLayout = { align, firstLine: Math.max(outmost, Math.min(firstLine, Math.max(0, width - 72))) };
        // A line no closer to the next than a quarter of its usual distance, nor farther than a page.
        if (spacing?.lineMultiple) layout.lineMultiple = Math.min(Math.max(spacing.lineMultiple, 0.25), 10);
        else if (spacing?.lineHeight) { layout.lineHeight = Math.min(Math.max(spacing.lineHeight, 2), this.bottom - this.margin.top); layout.lineAtLeast = spacing.lineAtLeast; }
        return { left, width, layout, spacing };
    }

    /**
     * The caller's `onNode` verdict for one node, normalized to `false` (drop it) / a replacement
     * string / undefined (render normally). Used for the nodes the engine paints as flat text - text
     * runs, table rows and cells - where a returned string is spliced in as text rather than drawn as
     * its own paragraph, so no "not representable" warning is owed.
     */
    /** onNode's verdict on each node, asked once: a node laid out twice (measured, then drawn) is not asked again. */
    private readonly onNodeVerdicts = new WeakMap<OfficeContentNode, false | string | undefined>();
    private async onNodeValue(node: OfficeContentNode): Promise<false | string | undefined> {
        if (!this.config.onNode) return undefined;
        if (this.onNodeVerdicts.has(node)) return this.onNodeVerdicts.get(node);
        const override = await this.config.onNode(node);
        const verdict = override === false ? false : typeof override === 'string' ? override : undefined;
        this.onNodeVerdicts.set(node, verdict);
        return verdict;
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
        const override = await this.onNodeValue(node);
        if (override === false) return true;
        if (typeof override === 'string') {
            if (!this.onNodeWarned) {
                this.onNodeWarned = true;
                const info = { format: 'pdf', feature: 'onNode string override (native engine draws it as plain text)', reason: 'onnode-string-override' };
                this.reportWarning(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, info);
            }
            if (override) this.drawRuns([{ text: override, fmt: {} }], this.margin.left, this.contentWidth);
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
        this.markAnchors(node);
        switch (node.type) {
            case 'page':
            case 'slide':
                // A slide whose shapes have their places is drawn as it is laid out (see drawSlide).
                if (node.type === 'slide' && (node.children || []).some(c => c.bounds)) { await this.drawSlide(node); return; }
                // A PDF's page keeps its own size (pdfConfig.format 'auto'), as each section's does.
                if (topLevel && node.type === 'page') {
                    this.takePageSize(node.metadata as { pageWidth?: number; pageHeight?: number } | undefined);
                    // Its headers are drawn on the pages its content is laid out on (see drawRunningContent).
                    const number = (node.metadata as { pageNumber?: number } | undefined)?.pageNumber;
                    if (typeof number === 'number') {
                        this.sourcePage = number;
                        const record = this.pageRecords[this.pageRecords.length - 1];
                        if (record && record.page === this.page && this.y <= this.margin.top + 0.5) record.sourcePage = number;
                    }
                }
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
                // Printed on its own pages' paper, orientation and margins (pdfConfig.format 'auto').
                this.onPagesOf((node.metadata as SheetMetadata | undefined)?.pageSetup, false, true);
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
            case 'table': await this.belowFloats(() => this.table(node)); return this.table(node);
            case 'image': return this.image(node);
            case 'code': await this.belowFloats(() => this.code(node)); return this.code(node);
            case 'drawing':
                // A frame floating beside the text, where it stands; any other shape as any other block.
                if ((node.metadata as ShapeMetadata | undefined)?.float && !this.fixedPage && !this.framesInFlow) return this.floatingFrame(node);
                if (node.children?.length) { for (const c of node.children) await this.render(c); }
                else if (node.text) await this.paragraph(node);
                return;
            case 'note': return this.note(node);
            case 'break': return await this.breakNode(node);
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
                this.drawRuns([{ text, fmt: {}, link: url || undefined }], this.margin.left, this.contentWidth);
                this.y += 6;
                return;
            }
            case 'chart':
                // includeCharts: false omits charts in every generator. The native engine cannot draw a
                // chart, so it otherwise renders the chart's data text (in node.text); skip it when off.
                if (this.config.includeCharts === false) return;
                await this.belowFloats(async () => { if (!this.drawChart(node) && node.text) await this.paragraph(node); });
                // Drawn from its data, as HTML output draws it (bars for each label, one colour a series).
                if (this.drawChart(node)) return;
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

    /** The height a chart has to be drawn in (its box on a slide), or undefined for a chart in the flow. */
    private chartRoom?: number;

    /**
     * Draws a chart from its data (ChartData on its attachment) as a bar chart, as the HTML engine draws
     * one: a bar for each series at each label, a value axis with its ticks, the title, the labels and a
     * legend. False, nothing drawn, for a chart without numbers to draw or with more than can be read on
     * a page; it is then written as its data's text. Charts printed as text only.
     */
    private drawChart(node: OfficeContentNode): boolean {
        const name = (node.metadata as { attachmentName?: string } | undefined)?.attachmentName;
        const data = (name ? this.attachmentsByName.get(name) as { chartData?: ChartData } | undefined : undefined)?.chartData;
        if (!data || !Array.isArray(data.dataSets) || !Array.isArray(data.labels)) return false;
        const series = data.dataSets.slice(0, 12).map(set => (Array.isArray(set.values) ? set.values : []).map(v => { const n = Number(String(v).replace(/,/g, '')); return Number.isFinite(n) ? n : NaN; }));
        const groups = maxOf(series.map(v => v.length), data.labels.length);
        if (!groups || groups > 120 || data.dataSets.length > 12 || !series.some(v => v.some(n => !Number.isNaN(n)))) return false;
        const ink = this.lib.rgb(0.12, 0.12, 0.12), grid = this.lib.rgb(0.85, 0.85, 0.85);
        const palette = [[255, 99, 132], [54, 162, 235], [255, 206, 86], [75, 192, 192], [153, 102, 255], [255, 159, 64]].map(([r, g, b]) => this.lib.rgb(r / 255, g / 255, b / 255));
        const values = series.flat().filter(n => !Number.isNaN(n));
        let low = 0, high = 0;
        for (const v of values) { if (v < low) low = v; if (v > high) high = v; }
        if (low === high) high = low + 1;
        // Ticks at a round step: 1, 2 or 5 times a power of ten, about five of them. Values whose range no
        // number holds, or too small to step through, are written as text (they crashed the conversion).
        const rough = (high - low) / 5, power = 10 ** Math.floor(Math.log10(rough));
        const step = [1, 2, 5, 10].map(m => m * power).find(t => t >= rough) ?? rough;
        if (!(Number.isFinite(rough) && rough > 0 && Number.isFinite(step) && step > 0 && Number.isFinite(low / step) && Number.isFinite(high / step))) return false;
        low = Math.floor(low / step) * step; high = Math.ceil(high / step) * step;
        const tickText = (v: number) => this.enc(String(+v.toFixed(6)), 'regular', false);
        const ticks: number[] = [];
        for (let v = low; v <= high + step / 2 && ticks.length < 50; v += step) ticks.push(v);
        const legendRows = data.dataSets.length > 1 ? Math.ceil(data.dataSets.length / 4) : 0;
        const titleH = data.title ? 20 : 0, labelH = 24, legendH = legendRows * 14;
        // As tall as its box on a slide allows, else 180.
        const plotH = this.chartRoom !== undefined ? Math.max(40, this.chartRoom - titleH - labelH - legendH - (data.xAxisTitle ? 14 : 0) - 10) : 180;
        const height = titleH + plotH + labelH + legendH + (data.xAxisTitle ? 14 : 0) + 10;
        this.ensureSpace(height);
        if (this.pendingAnchors.length) this.placeAnchors();
        const top = this.pageH - this.y;
        if (data.title) {
            const title = this.enc(String(data.title), 'bold');
            const w = this.widthOf(title, 'bold', 12);
            this.drawText(title, 'bold', this.margin.left + Math.max(0, (this.contentWidth - w) / 2), top - 14, 12, ink);
        }
        const axisW = maxOf(ticks.map(t => this.widthOf(tickText(t), 'regular', 8)), -Infinity) + 6 + (data.yAxisTitle ? 14 : 0);
        const left = this.margin.left + axisW, width = this.contentWidth - axisW;
        const plotTop = top - titleH, plotBottom = plotTop - plotH;
        const yOf = (v: number) => plotBottom + (v - low) / (high - low) * plotH;
        for (const t of ticks) {
            const y = yOf(t);
            this.page.drawLine({ start: { x: left, y }, end: { x: left + width, y }, thickness: 0.4, color: grid });
            const text = tickText(t);
            this.drawText(text, 'regular', left - 4 - this.widthOf(text, 'regular', 8), y - 3, 8, ink);
        }
        this.page.drawLine({ start: { x: left, y: plotBottom }, end: { x: left, y: plotTop }, thickness: 0.6, color: ink });
        this.page.drawLine({ start: { x: left, y: yOf(0) }, end: { x: left + width, y: yOf(0) }, thickness: 0.6, color: ink });
        if (data.yAxisTitle) {
            this.useFont(this.fonts.regular[0].font);
            this.page.drawText(this.enc(String(data.yAxisTitle), 'regular').slice(0, 60), { x: this.margin.left + 8, y: plotBottom + plotH / 2 - this.widthOf(this.enc(String(data.yAxisTitle), 'regular').slice(0, 60), 'regular', 8) / 2, size: 8, color: ink, rotate: this.lib.degrees(90) });
        }
        const groupW = width / groups, barW = groupW * 0.8 / series.length;
        series.forEach((values, s) => values.forEach((v, g) => {
            if (Number.isNaN(v)) return;
            const x = left + g * groupW + groupW * 0.1 + s * barW;
            const y0 = yOf(0), y1 = yOf(v);
            this.page.drawRectangle({ x, y: Math.min(y0, y1), width: Math.max(0.5, barW - 1), height: Math.max(0.5, Math.abs(y1 - y0)), color: palette[s % palette.length], opacity: 0.75, borderColor: palette[s % palette.length], borderWidth: 0.5 });
        }));
        // Each label under its bars, cut to their width; every few, where they are too many to be read.
        const every = Math.max(1, Math.ceil(groups * 18 / width));
        data.labels.slice(0, groups).forEach((label, g) => {
            if (g % every) return;
            const text = splitToWidth(this.enc(String(label ?? ''), 'regular'), t => this.widthOf(t, 'regular', 7), groupW * every - 2)[0] ?? '';
            this.drawText(text, 'regular', left + g * groupW + Math.max(0, (groupW - this.widthOf(text, 'regular', 7)) / 2), plotBottom - 10, 7, ink);
        });
        let below = plotBottom - labelH;
        if (data.xAxisTitle) {
            const text = this.enc(String(data.xAxisTitle), 'regular');
            this.drawText(text, 'regular', left + Math.max(0, (width - this.widthOf(text, 'regular', 8)) / 2), below + 6, 8, ink);
            below -= 14;
        }
        data.dataSets.forEach((set, s) => {
            if (!legendRows) return;
            const x = this.margin.left + (s % 4) * (this.contentWidth / 4), y = below - Math.floor(s / 4) * 14;
            this.page.drawRectangle({ x, y, width: 8, height: 8, color: palette[s % palette.length], opacity: 0.75 });
            const text = splitToWidth(this.enc(set.name || `Series ${s + 1}`, 'regular'), t => this.widthOf(t, 'regular', 8), this.contentWidth / 4 - 14)[0] ?? '';
            this.drawText(text, 'regular', x + 11, y + 1, 8, ink);
        });
        this.y += height;
        return true;
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
        const { left, width, layout, spacing } = this.placement(node);
        if (spacing?.before !== undefined) this.spaceAbove(spacing.before);
        else if (this.y > this.margin.top) this.y += size * 0.6;
        // Bold, at the level's size unless its runs have their own (a slide's title, a style's size).
        const runs = (await this.collectRuns(node)).map(r => ({ ...r, fmt: { ...r.fmt, bold: true, size: r.fmt.size ?? `${size}pt` } }));
        const title = (runs.length ? runs.map(r => r.text).join('') : node.text || '').replace(/\s+/g, ' ').trim();
        if (title && !this.furniture) this.outlineEntries.push({ title, level, page: this.page, top: this.pageH - this.y });
        this.drawRuns(runs.length ? runs : [{ text: node.text || '', fmt: { bold: true, size: `${size}pt` } }], left, width, false, layout);
        this.y += this.spaceBelow(spacing?.after, size * 0.35);
    }

    private async paragraph(node: OfficeContentNode, indentLeft = 0): Promise<void> {
        // A picture in the paragraph (where Word, ODF and Markdown put one) is drawn where it stands,
        // between the lines of text around it: runs carry text only, and it was dropped.
        // A picture among the paragraph's words is drawn in its line, where it can be embedded (it broke
        // the paragraph in two, as a picture of its own between them).
        const images = (node.children || []).filter(child => child.type === 'image');
        if (images.length && (node.children || []).some(child => child.type !== 'image' && (child.text || child.children?.length))
            && (this.imageMode === 'image-only' || this.imageMode === 'image+ocr-text')) {
            for (const image of images) {
                const picture = await this.embedPicture(image);
                if (picture) this.inlinePictures.set(image, picture);
            }
            if (images.every(image => this.inlinePictures.has(image))) {
                const runs = await this.collectRuns(node);
                const { left, width, layout, spacing } = this.placement(node, indentLeft);
                this.spaceAbove(spacing?.before);
                this.drawRuns(runs, left, width, false, layout);
                this.y += this.spaceBelow(spacing?.after, 6);
                // 'image+ocr-text': the pictures' recognized text after the paragraph, as after a picture of its own.
                for (const image of images) { const ocr = this.ocrBesidePicture(image); if (ocr) await this.drawOcrText(ocr); }
                return;
            }
        }
        if (node.children?.some(child => child.type === 'image')) {
            let run: OfficeContentNode[] = [];
            const flush = async () => {
                const runs = run.length ? await this.collectRuns({ ...node, text: undefined, children: run }) : [];
                if (runs.some(r => r.text.trim())) {
                    const { left, width, layout, spacing } = this.placement(node, indentLeft);
                    this.spaceAbove(spacing?.before);
                    this.drawRuns(runs, left, width, false, layout);
                    this.y += this.spaceBelow(spacing?.after, 6);
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
        const { left, width, layout, spacing } = this.placement(node, indentLeft);
        this.spaceAbove(spacing?.before);
        this.drawRuns(runs, left, width, false, layout);
        this.y += this.spaceBelow(spacing?.after, 6);
    }

    /** The pictures drawn in their lines (see paragraph), embedded. */
    private readonly inlinePictures = new Map<OfficeContentNode, { img: any; w: number; h: number }>();

    /** Draws one already-known plain string as a paragraph (an image's alt text, a note marker). */
    private plainParagraph(text: string): void {
        if (!text) return;
        this.drawRuns([{ text, fmt: {} }], this.margin.left, this.contentWidth);
        this.y += 6;
    }

    private async listItem(node: OfficeContentNode): Promise<void> {
        const meta = node.metadata as any;
        const level = Math.max(0, meta?.indentation || 0);
        const indent = 18 + level * 18;
        const ordered = meta?.listType === 'ordered';
        // The document's own marker (`a)`, `iv.`, `◦`), or the item's number and a dot, or a bullet.
        const marker = typeof meta?.marker === 'string' ? meta.marker : ordered ? `${(meta?.itemIndex ?? 0) + 1}.` : '•';
        // Body: the item's own text runs (its non-list children); nested list children render after.
        const bodyRuns = await this.collectRuns({ ...node, children: (node.children || []).filter(c => c.type !== 'list') });
        // Size the marker to the body text so their baselines line up (a body larger than the old fixed
        // 11pt would otherwise sit below its own marker), and reserve one line's height BEFORE drawing
        // the marker so a body that has to break to the next page carries its marker along instead of
        // orphaning it on this page.
        const bodySize = bodyRuns.length ? maxOf(bodyRuns.map(r => parseFontSize(r.fmt.size) ?? 11), -Infinity) : 11;
        const size = bodySize > 0 ? bodySize : 11;
        this.ensureSpace(size * 1.35);
        // Draw the marker, then the item body hanging-indented past it. Space is already reserved, so
        // drawRuns' first line stays on this page and draws at the same baseline as the marker.
        const markerX = this.margin.left + indent - 14;
        const baseline = this.pageH - this.y - size;
        // In its text's colour (white bullets on a dark slide were drawn dark); an invisible item's invisibly, as its text is.
        const ink = this.color(bodyRuns.find(r => r.text.trim())?.fmt.color) ?? this.lib.rgb(0.12, 0.12, 0.12);
        const invisibleItem = bodyRuns.some(r => r.text.trim()) && bodyRuns.every(r => !r.text.trim() || r.fmt.invisible);
        if (marker && (invisibleItem || !this.drawBulletShape(marker, markerX, baseline, size, ink))) this.drawText(this.enc(marker), 'regular', markerX, baseline, size, ink, invisibleItem);
        // The body starts past a marker wider than the usual gap (`1.1.1.`, `(iii)`), as a tab stop
        // past it would set it.
        const markerWidth = marker ? this.widthOf(this.enc(marker, 'regular', false), 'regular', size) : 0;
        const bodyIndent = Math.min(Math.max(indent, indent - 14 + markerWidth + 4), this.contentWidth / 2);
        // Advance for the marker line unless the body has real text to draw: a whitespace-only body
        // makes drawRuns place nothing and never advance `this.y`, so the next item would overprint.
        if (bodyRuns.some(r => r.text && r.text.trim())) this.drawRuns(bodyRuns, this.margin.left + bodyIndent, this.contentWidth - bodyIndent);
        else this.y += size * 1.35;
        // Route nested list items through render() (not listItem() directly) so the onNode hook and the
        // abort check see them like every other node.
        for (const c of (node.children || []).filter(c => c.type === 'list')) await this.render(c);
    }

    /**
     * Draws a bullet no font given has as the shape it is (a circle, a square, a diamond), the
     * Standard-14 fonts having none but the round bullet. False when the marker is not such a bullet,
     * or a font draws it.
     */
    private drawBulletShape(marker: string, x: number, baseline: number, size: number, color: any): boolean {
        const cp = marker.codePointAt(0)!;
        if ([...marker].length !== 1 || this.fonts.regular.some(face => face.covers(cp))) return false;
        const r = size * 0.18, cx = x + r + 1, cy = baseline + size * 0.3;
        switch (marker) {
            case '◦': case '○': case '❍':
                this.page.drawCircle({ x: cx, y: cy, size: r, borderColor: color, borderWidth: 0.6 });
                return true;
            case '●': case '⚫':
                this.page.drawCircle({ x: cx, y: cy, size: r, color });
                return true;
            case '▪': case '■': case '◼': case '▫': case '□': case '◻': case '❑': {
                const filled = '▪■◼'.includes(marker);
                this.page.drawRectangle({ x: cx - r, y: cy - r, width: 2 * r, height: 2 * r, ...(filled ? { color } : { borderColor: color, borderWidth: 0.6 }) });
                return true;
            }
            case '◆': case '♦': case '❖': case '◇': {
                const d = r * 1.3;
                // Drawn about its centre: an SVG path's y runs down the page, which a diamond does not mind.
                this.page.drawSvgPath(`M 0 ${-d} L ${d} 0 L 0 ${d} L ${-d} 0 Z`, { x: cx, y: cy, ...(marker === '◇' ? { borderColor: color, borderWidth: 0.6 } : { color }) });
                return true;
            }
            default:
                // Another symbol (an arrow, a check mark): a round bullet, as the item is still one.
                if (/\p{So}|\p{Sm}/u.test(marker) || (cp >= 0x2190 && cp <= 0x27BF)) { this.drawText('•', 'regular', x, baseline, size, color); return true; }
                return false;
        }
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
            // Each line cut to the width of the column it lands in (a line may run on into a narrower one).
            // (Cut again only where the width changes, so a long line is cut in time in proportion to it.)
            const expanded = this.enc(raw.replace(/\t/g, '    '), 'mono') || ' ';
            let at = 0, cutFor = NaN, chunks: string[] = [], next = 0;
            while (at < expanded.length) {
                this.ensureSpace(lh);
                if (this.contentWidth !== cutFor) { cutFor = this.contentWidth; chunks = splitToWidth(expanded.slice(at), t => this.widthOf(t, 'mono', size), cutFor - 12); next = 0; }
                const chunk = chunks[next++] || expanded.slice(at);
                this.drawLine(chunk, 'mono', this.margin.left + 6, this.pageH - this.y - size, size, this.lib.rgb(0.12, 0.12, 0.12), this.contentWidth - 12);
                this.y += lh;
                at += chunk.length;
            }
        }
    }

    /** Draws OCR text: multi-line keeps its 2-D layout in monospace, a single line flows as prose. */
    private async drawOcrText(text: string): Promise<void> {
        if (/[\r\n]/.test(text)) { this.y += 2; this.drawMonospace(text); this.y += 4; }
        else await this.plainParagraph(text);
    }

    private async note(node: OfficeContentNode): Promise<void> {
        const label = this.noteLabels.get(node);
        const runs = [...(label ? [{ text: `${label} `, fmt: { size: '9pt', superscript: true } } as Run] : []),
            ...(await this.collectRuns(node)).map(r => ({ ...r, fmt: { ...r.fmt, size: '9pt' } }))];
        if (runs.length) this.drawRuns(runs, this.margin.left + 12, this.contentWidth - 12);
        this.y += 4;
    }

    private async breakNode(node: OfficeContentNode): Promise<void> {
        const meta = node.metadata as BreakMetadata | undefined;
        const t = meta?.breakType;
        // A section that starts on this page: its columns from here (balanced where the next ends it here).
        if (t === 'section') {
            // (One on pages of another size, or turned, starts on a page of its own, as a word processor sets it:
            // it went on on the page before, of its size.)
            const next = meta?.pageSetup && this.sectionPage ? this.sectionPage(meta.pageSetup) : undefined;
            if (next && typeof meta?.section === 'number' && (Math.abs(next.w - this.pageW) > 0.5 || Math.abs(next.h - this.pageH) > 0.5)) {
                await this.setColumns(undefined, false);
                this.startSection(meta.section, meta.pageSetup);
                await this.setColumns(meta.pageSetup?.columns, true);
                return;
            }
            if (typeof meta?.section === 'number') this.noteSection(meta.section, meta.pageSetup);
            await this.setColumns(meta?.pageSetup?.columns, true);
            return;
        }
        if (t === 'page' && typeof meta?.section === 'number' && meta.section > 0) {
            await this.setColumns(undefined, false);
            this.startSection(meta.section, meta.pageSetup);
            await this.setColumns(meta.pageSetup?.columns, true);
            return;
        }
        if (t === 'page') { this.newPage(); return; }
        // A column break: the next column, where the text is set in columns.
        if (t === 'column' && this.columns) { if (!this.fixedPage) this.nextColumnOrPage(); return; }
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
        // Nothing renderable: keep the alt text (or the OCR text) rather than drawing nothing (the
        // failure is reported): the DOCX and ODT generators degrade and warn for the same input.
        const picture = await this.embedPicture(node);
        if (!picture) { this.plainParagraph(meta?.altText || ocr); return; }
        let { w, h } = picture;
        const scale = Math.min(1, this.contentWidth / w);
        w *= scale; h *= scale;
        if (h > this.bottom - this.margin.top) { const s = (this.bottom - this.margin.top) / h; w *= s; h *= s; }
        this.belowFloatsFor(h + 6);
        this.ensureSpace(h + 6);
        this.drawPicture(picture.img, { x: this.margin.left, y: this.pageH - this.y - h, width: w, height: h });
        const target = (node.metadata as { link?: unknown } | undefined)?.link;
        if (typeof target === 'string' && target) this.addLinkBox(target, this.margin.left, this.pageH - this.y - h, w, h);
        this.y += h + 6;
        // 'image+ocr-text': draw the recognized text just below the successfully embedded image.
        if (mode === 'image+ocr-text' && ocr) await this.drawOcrText(ocr);
    }

    /**
     * A picture embedded in the PDF (each attachment once, wherever it shows) and the size it is drawn at
     * in points: its ImageMetadata width and height (a length, pixels, or a '%' of the text's width),
     * else its bounds on a PDF page, else its pixels at 96 per inch. Null, the failure reported, when it
     * cannot be: no data, too large, or a format a PDF holds none of.
     */
    /** Draws an embedded picture: a raster image, or an SVG's page (a form). */
    private drawPicture(img: any, box: { x: number; y: number; width: number; height: number }): void {
        if (img instanceof this.lib.PDFEmbeddedPage) this.page.drawPage(img, box);
        else this.page.drawImage(img, box);
    }

    /**
     * An SVG picture drawn on a page of its own, in this document (so its text is in the document's
     * fonts), embedded as a form and the page taken away; what it could not draw is reported. Its
     * `width` and `height` are its size in CSS pixels, as a raster picture's are.
     */
    private async svgPicture(bytes: Uint8Array): Promise<any> {
        const svg = readSvg(new TextDecoder().decode(bytes), this.ast.config);
        if (!svg) return 'unreadable SVG picture';
        const savedPage = this.page;
        // (A page outside the document's: one added and taken away again cost time in the pages there were.)
        const scratch = this.lib.PDFPage.create(this.pdf);
        scratch.setSize(svg.width * 0.75, svg.height * 0.75);
        this.page = scratch;
        let left: string[];
        try {
            left = await drawSvg(svg, {
                lib: this.lib, page: scratch,
                text: (text, face, x, y, size, color) => this.drawText(this.enc(text, face as FaceKey), face as FaceKey, x, y, size, color),
                width: (text, face, size) => this.widthOf(this.enc(text, face as FaceKey, false), face as FaceKey, size),
                image: async (data, kind) => { try { return kind === 'png' ? await this.pdf.embedPng(data) : await this.pdf.embedJpg(data); } catch { return undefined; } },
            });
        } finally {
            this.page = savedPage;
        }
        if (left.length) this.reportWarning(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { reason: 'svg', feature: `SVG ${left.join(', ')}`, format: 'pdf' });
        return await this.pdf.embedPage(scratch);
    }

    private async embedPicture(node: OfficeContentNode): Promise<{ img: any; w: number; h: number } | null> {
        const meta = node.metadata as any;
        const name = meta?.attachmentName;
        const attachment = name ? this.attachmentsByName.get(name) : undefined;
        if (!attachment?.data) {
            if (name) this.reportWarning(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name, reason: attachment ? 'attachment has no data' : 'missing attachment' });
            return null;
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
                    // An SVG drawn as vector graphics, on a page embedded as a form (see svgDrawing).
                    if (isSvg(bytes, attachment)) return await this.svgPicture(bytes);
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
                return null;
            }
            // Its size in CSS pixels: a raster picture's pixels, an SVG's page (in points) as pixels.
            const svgPage = img instanceof this.lib.PDFEmbeddedPage;
            const pxW = svgPage ? img.width / 0.75 : img.width, pxH = svgPage ? img.height / 0.75 : img.height;
            const ratio = pxW ? pxH / pxW : 1;
            const pctOf = (v: any): number | null => (typeof v === 'string' && /^\s*[\d.]+\s*%\s*$/.test(v)) ? parseFloat(v) / 100 : null;
            let w: number, h: number;
            if (meta?.width != null) {
                const p = pctOf(meta.width);
                w = p != null ? this.contentWidth * p : toPoints(meta.width, pxW * 0.75);
                h = (meta.height != null && pctOf(meta.height) == null) ? toPoints(meta.height, w * ratio) : w * ratio;
            } else if (node.bounds?.width) {
                w = node.bounds.width;
                h = node.bounds.height || w * ratio;
            } else {
                w = pxW * 0.75; // intrinsic pixels -> points (96dpi), not points 1:1
                h = pxH * 0.75;
            }
            return { img, w: Math.max(1, w), h: Math.max(1, h) };
        } catch {
            // pdf-lib embeds PNG and JPEG (and a GIF or BMP is turned into PNG), so a WebP, TIFF or SVG
            // attachment lands here.
            this.reportWarning(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name, reason: 'unsupported image format (the native PDF engine embeds PNG, JPEG, GIF, BMP and SVG)' });
            return null;
        }
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
            // Numbered in the order they are cited: footnotes 1, 2, 3, endnotes i, ii, iii, as Word does.
            if (NativeLayout.isNumberedNote(note)) {
                const endnote = (note.metadata as { noteType?: string } | undefined)?.noteType === 'endnote';
                const count = endnote ? ++this.endnoteCount : ++this.footnoteCount;
                this.noteLabels.set(note, endnote ? formatListNumber(count, 'lowerRoman') : String(count));
            }
            this.collectNotes(note);
        }
        for (const c of node.children || []) this.collectNotes(c);
    }

    /** Draws the footnote/endnote bodies gathered during the walk, under a short separating rule. */
    async flushNotes(): Promise<void> {
        // The last page's footnotes at its foot; then the endnotes, and any footnote no page took.
        const carried = this.carriedFootnotes;
        this.carriedFootnotes = [];
        for (const note of carried) this.placedNotes.delete(note);
        this.drawPageFootnotes();
        const rest = this.collectedNotes.filter(n => !this.placedNotes.has(n));
        if (!rest.length) return;
        this.y += 8;
        this.ensureSpace(14);
        const yy = this.pageH - this.y;
        this.page.drawLine({ start: { x: this.margin.left, y: yy }, end: { x: this.margin.left + Math.min(180, this.contentWidth), y: yy }, thickness: 0.5, color: this.lib.rgb(0.7, 0.7, 0.7) });
        this.y += 8;
        for (const n of rest) await this.note(n);
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
        type Placed = { col: number; span: number; rowSpan: number; text: string; header: boolean; lines: string[]; fill?: string; align?: TextAlignment; ink?: string; invisible?: boolean; pictures?: { img: any; w: number; h: number }[]; notes?: OfficeContentNode[] };
        const grid: Placed[][] = [];
        const chartsAfter: OfficeContentNode[] = [];
        const picturesAfter: { img: any; w: number; h: number }[] = [];
        // Grid occupancy: column -> rows still covered by a rowspan from above, so a spanning cell
        // reserves its column(s) below instead of letting later rows slide left under it.
        const carry = new Map<number, number>();
        /**
         * A cell's fill, alignment and text colour: its own, or what every text run in it shares (a sheet's
         * cell style rides on its runs). Shading was not drawn, and every cell's text sat at its left in grey
         * (white text on a dark fill was unreadable once the fill was drawn).
         */
        const cellLook = (cell: OfficeContentNode): { fill?: string; align?: TextAlignment; ink?: string; invisible?: boolean } => {
            const meta = cell.metadata as { backgroundColor?: string; align?: TextAlignment } | undefined;
            const runs: OfficeContentNode[] = [];
            const gather = (n: OfficeContentNode, depth: number) => {
                if (runs.length > 64 || depth > 8) return;
                if (n.type === 'text') runs.push(n);
                else for (const c of n.children || []) gather(c, depth + 1);
            };
            for (const c of cell.children || []) gather(c, 0);
            const shared = (pick: (n: OfficeContentNode) => string | undefined) => {
                const values = runs.filter(n => (n.text || '').trim()).map(pick);
                return values.length && values.every(v => v && v === values[0]) ? values[0] : undefined;
            };
            const fill = cssColorHex(meta?.backgroundColor) ?? cssColorHex(shared(n => n.formatting?.backgroundColor)) ?? undefined;
            const blocks = (cell.children || []).filter(c => c.type === 'paragraph' || c.type === 'heading');
            const blockAlign = blocks.length && blocks.every(b => (b.metadata as any)?.alignment === (blocks[0].metadata as any)?.alignment) ? (blocks[0].metadata as any)?.alignment : undefined;
            const align = meta?.align ?? blockAlign ?? shared(n => n.formatting?.alignment);
            const ink = cssColorHex(shared(n => n.formatting?.color)) ?? undefined;
            // A cell whose text is all invisible (a scan's recognised table over its picture) drawn invisibly.
            const invisible = shared(n => n.formatting?.invisible ? 'invisible' : undefined) === 'invisible';
            return { ...(fill ? { fill } : {}), ...(align === 'center' || align === 'right' ? { align } : {}), ...(ink ? { ink } : {}), ...(invisible ? { invisible } : {}) };
        };
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
                // A cell's pictures are drawn in it, under its text (each one that can be embedded; another
                // is its alt text): only their alt text was.
                const pictures: { img: any; w: number; h: number }[] = [];
                if (typeof override !== 'string') {
                    for (const picture of this.picturesIn(cell)) {
                        const embedded = await this.embedPicture(picture);
                        if (embedded) { pictures.push(embedded); this.drawnApart.add(picture); }
                    }
                    for (const chart of this.chartsIn(cell)) { chartsAfter.push(chart); this.drawnApart.add(chart); }
                }
                // Cells built by the markup parsers carry their text only in `children`; reading
                // `cell.text` alone rendered every Markdown/HTML table as an empty grid.
                const text = typeof override === 'string' ? override : await this.flatText(cell);
                // An empty cell under another's span (a merged range's covered cells, as XLSX keeps them)
                // is not drawn: its box would cut through the merged one.
                if (!text.trim() && (spanned.has(col) || (carry.get(col) || 0) > 0)) continue;
                const header = rowIsHeader || meta?.style === 'header';
                const notes = this.notesCitedIn(cell);
                placed.push({ col, span, rowSpan: Math.max(1, meta?.rowSpan || 1), text, header, lines: [], ...cellLook(cell), ...(pictures.length ? { pictures } : {}), ...(notes.length ? { notes } : {}) });
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
        let { size } = this.columnWidths(grid, cols, pad);
        let xs: number[] = [];
        // A row run on into another column is drawn in it (its cells' edges move with the column's), and
        // in a column of another width the table is laid out again at that width (see relayout).
        let tableLeft = this.margin.left, layoutWidth = this.contentWidth;
        const place = (widths: number[]) => { xs = [this.margin.left]; for (const w of widths) xs.push(xs[xs.length - 1] + w); };
        place(this.columnWidths(grid, cols, pad).widths);
        /** Lays the table out again where the text has moved into a column of another width; whether it did. */
        const relayout = (): boolean => {
            if (this.contentWidth === layoutWidth) return false;
            const again = this.columnWidths(grid, cols, pad);
            size = again.size;
            place(again.widths);
            tableLeft = this.margin.left;
            layoutWidth = this.contentWidth;
            return true;
        };
        const spanWidth = (col: number, span: number) => xs[Math.min(col + span, cols)] - xs[Math.min(col, cols)];
        /** Wraps a cell's (un-encoded) text to the width of the columns it spans. */
        // A header cell is drawn bold, so it is wrapped at the bold font's widths.
        const wrap = (text: string, col: number, span: number, header = false) => {
            const face: FaceKey = header ? 'bold' : 'regular';
            return wrapPlain(text, t => this.widthOf(t, face, size), Math.max(10, spanWidth(col, span) - 2 * pad), t => this.enc(t, face));
        };

        // A table whose text is all invisible (a scan's recognised table) is drawn without its rules, as its text is unseen.
        const unseen = grid.some(r => r.some(p => p.text.trim())) && grid.every(r => r.every(p => !p.text.trim() || p.invisible));
        const ruled = unseen ? {} : { borderColor: border, borderWidth: 0.5 };
        this.y += 2;
        for (const placed of grid) {
            relayout();
            for (const p of placed) p.lines = wrap(p.text, p.col, p.span, p.header);
            const lineH = size * 1.35;
            // A row with pictures is drawn whole, on a page of its own when it does not fit this one: each
            // cell's text, then its pictures, scaled to its width (and to half a page at most).
            // A row with a cell whose text alone takes most of a page has its pictures drawn after the table,
            // and is drawn in bands as any other (drawn whole, it ran off the page).
            const pageBody = this.pageH - this.margin.top - this.margin.bottom - this.footReserve - 2 * pad;
            if (placed.some(p => p.lines.length * lineH > pageBody - 40)) for (const p of placed) if (p.pictures?.length) { picturesAfter.push(...p.pictures); p.pictures = undefined; }
            if (placed.some(p => p.pictures?.length)) {
                const sized = placed.map(p => {
                    const pics = (p.pictures || []).map(pic => {
                        const scale = Math.min(1, Math.max(1, spanWidth(p.col, p.span) - 2 * pad) / pic.w);
                        return { img: pic.img, w: pic.w * scale, h: pic.h * scale };
                    });
                    // All of a cell's pictures, and its text, on one page: scaled down together where they are taller.
                    const room = pageBody - p.lines.length * lineH - 4 * pics.length;
                    const total = pics.reduce((sum, pic) => sum + pic.h, 0);
                    const fit = total > room && total > 0 ? Math.max(0.01, room / total) : 1;
                    return pics.map(pic => ({ img: pic.img, w: pic.w * fit, h: pic.h * fit }));
                });
                const heightOf = (i: number) => placed[i].lines.length * lineH + sized[i].reduce((sum, pic) => sum + pic.h + 4, 0);
                const rowH = Math.max(size * 1.4, ...placed.map((_, i) => heightOf(i))) + 2 * pad;
                // (The table's top gap is not something drawn: a row starting there needs no new page.)
                if (this.y + rowH > this.bottom && this.y > this.margin.top + 3) this.nextColumnOrPage();
                // (Moved into a column of another width, the row is drawn at the table's old width there:
                // its pictures were sized for it. A row of pictures is rare in columns of unequal widths.)
                if (this.pendingAnchors.length) this.placeAnchors();
                const topY = this.pageH - this.y;
                placed.forEach((p, i) => {
                    const x = xs[Math.min(p.col, cols)] + this.margin.left - tableLeft;
                    const room = spanWidth(p.col, p.span) - 2 * pad;
                    if (!unseen || p.fill) this.page.drawRectangle({ x, y: topY - rowH, width: spanWidth(p.col, p.span), height: rowH, ...ruled, color: p.fill ? this.color(p.fill) : undefined });
                    const face: FaceKey = p.header ? 'bold' : 'regular';
                    let ty = topY - pad - size;
                    for (const line of p.lines) {
                        const width = this.widthOf(line, face, size);
                        const free = p.align ? Math.max(0, room - width) : 0;
                        // An aligned line is placed here: drawn in its own width, a right-to-left one is not set right again.
                        this.drawLine(line, face, x + pad + (p.align === 'center' ? free / 2 : p.align === 'right' ? free : 0), ty, size, (p.ink && this.color(p.ink)) || this.lib.rgb(0.12, 0.12, 0.12), p.align ? width : room, p.invisible);
                        ty -= lineH;
                    }
                    let py = topY - pad - p.lines.length * lineH;
                    for (const pic of sized[i]) {
                        const free = Math.max(0, room - pic.w);
                        this.drawPicture(pic.img, { x: x + pad + (p.align === 'center' ? free / 2 : p.align === 'right' ? free : 0), y: py - pic.h, width: pic.w, height: pic.h });
                        py -= pic.h + 4;
                    }
                });
                this.y += rowH;
                for (const p of placed) for (const note of p.notes || []) this.citeFootnote(note);
                continue;
            }
            let rowH = Math.max(size * 1.4, ...placed.map(p => p.lines.length * lineH)) + 2 * pad;
            // Draw the row in vertical bands so a row taller than a whole page is split across pages
            // rather than drawn off the bottom (which silently lost the overflowing text). A normal row
            // is a single band. Each band redraws every cell's border and its slice of the wrapped
            // lines, so the cell continues cleanly on the next page.
            const fullBand = Math.max(1, Math.floor((this.bottom - this.margin.top - 2 * pad) / lineH));
            let maxLines = Math.max(1, ...placed.map(p => p.lines.length));
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
                else if (this.y > this.margin.top) {
                    // Nothing fits: the next column or page. In a column of another width, what is left of the
                    // row's cells is wrapped again at the widths the table has there.
                    this.nextColumnOrPage();
                    if (relayout()) {
                        for (const p of placed) p.lines = [...p.lines.slice(0, lineOffset), ...wrap(p.lines.slice(lineOffset).join(' '), p.col, p.span, p.header)];
                        maxLines = Math.max(1, ...placed.map(p => p.lines.length));
                        rowH = Math.max(size * 1.4, ...placed.map(p => p.lines.length * lineH)) + 2 * pad;
                    }
                    continue;
                }
                else bandLines = Math.min(remaining, fullBand);                  // at page top: a full band
                const whole = lineOffset === 0 && bandLines === maxLines;
                const bandH = whole ? rowH : bandLines * lineH + 2 * pad;
                if (this.pendingAnchors.length) this.placeAnchors();
                const topY = this.pageH - this.y;
                for (const p of placed) {
                    const x = xs[Math.min(p.col, cols)] + this.margin.left - tableLeft;
                    if (!unseen || p.fill) this.page.drawRectangle({ x, y: topY - bandH, width: spanWidth(p.col, p.span), height: bandH, ...ruled, color: p.fill ? this.color(p.fill) : undefined });
                    const face: FaceKey = p.header ? 'bold' : 'regular';
                    const room = spanWidth(p.col, p.span) - 2 * pad;
                    let ty = topY - pad - size;
                    for (let li = lineOffset; li < lineOffset + bandLines && li < p.lines.length; li++) {
                        const width = this.widthOf(p.lines[li], face, size);
                        const free = p.align ? Math.max(0, room - width) : 0;
                        const lx = x + pad + (p.align === 'center' ? free / 2 : p.align === 'right' ? free : 0);
                        this.drawLine(p.lines[li], face, lx, ty, size, (p.ink && this.color(p.ink)) || this.lib.rgb(0.12, 0.12, 0.12), p.align ? width : room, p.invisible); ty -= lineH;
                    }
                }
                this.y += bandH;
                lineOffset += bandLines;
            }
            // The footnotes its cells cite, given room at the foot of the page the row ends on.
            for (const p of placed) for (const note of p.notes || []) this.citeFootnote(note);
        }
        this.y += 6;
        for (const picture of picturesAfter) {
            const scale = Math.min(1, this.contentWidth / picture.w, (this.pageH - this.margin.top - this.margin.bottom) * 0.9 / picture.h);
            const w = picture.w * scale, h = picture.h * scale;
            this.ensureSpace(h + 6);
            this.drawPicture(picture.img, { x: this.margin.left, y: this.pageH - this.y - h, width: w, height: h });
            this.y += h + 6;
        }
        for (const chart of chartsAfter) await this.render(chart);
    }

    /**
     * The widths of a table's columns and the size its text is set in (see table): each column at least
     * as wide as its longest word, at most as wide as its longest line, at 10pt; set smaller, down to 6pt,
     * when its words alone are wider than the page. Measured from the cells that span one column.
     */
    private columnWidths(grid: { col: number; span: number; text: string; header: boolean; pictures?: { w: number }[] }[][], cols: number, pad: number): { widths: number[]; size: number } {
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
                // A picture asks for its width (to most of the page's), and is scaled to what it gets.
                for (const picture of p.pictures || []) most[p.col] = Math.max(most[p.col], Math.min(picture.w, room * 0.6) + 2 * pad);
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

/** Whether two boxes are the same (a shape's, which its text nodes share). */
function sameBox(a: NodeBounds | undefined, b: NodeBounds | undefined): boolean {
    return !!a && !!b && (a === b || (a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height));
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
/** A text's grapheme clusters (its code points where they cannot be told, or it has no marks). */
function clustersOf(text: string): Iterable<string> {
    const Segmenter = (Intl as unknown as { Segmenter?: new (locale?: string, options?: object) => { segment(t: string): Iterable<{ segment: string }> } }).Segmenter;
    if (!Segmenter || !/\p{M}/u.test(text)) return text;
    const segments = new Segmenter(undefined, { granularity: 'grapheme' }).segment(text);
    return { *[Symbol.iterator]() { for (const { segment } of segments) yield segment; } };
}

function splitToWidth(text: string, measure: (text: string) => number, width: number): string[] {
    if (width <= 0 || measure(text) <= width) return [text];
    const chunks: string[] = [];
    let cur = '';
    // A running width, each character measured once: measuring the chunk so far again at every
    // character took time quadratic in its length, and zero-width characters (combining marks) never
    // fill a line, so 16 KB of them took seconds. A word broken at character boundaries is unshaped
    // across the break anyway, so the sum of its characters' widths is as exact as the split can be.
    let curWidth = 0;
    // Cut between the characters a reader sees (grapheme clusters), not inside one: a syllable cut after its
    // letter left its vowel sign to start the next line alone.
    for (const ch of clustersOf(text)) {
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
 * A font as a caller gives it: the bytes of a TrueType, OpenType, WOFF or WOFF2 file. Anything else
 * given (a file name, an empty file) is reported and the default font used, rather than passed over
 * without a word.
 */
function fontBytes(value: unknown, name: string, reportWarning: (type: OfficeWarningType, info?: any) => void): Uint8Array | null {
    if (value === undefined || value === null) return null;
    if (value instanceof Uint8Array && value.byteLength) return value;
    if (value instanceof ArrayBuffer && value.byteLength) return new Uint8Array(value);
    const shown = value instanceof Uint8Array || value instanceof ArrayBuffer ? 'an empty file' : typeof value === 'string' ? value : Object.prototype.toString.call(value);
    reportWarning(OfficeWarningType.INVALID_CONFIG_VALUE, { option: `pdfConfig.${name}`, value: shown, expected: 'the bytes of a TrueType, OpenType, WOFF or WOFF2 font file (a Uint8Array or ArrayBuffer)', fallback: undefined });
    return null;
}

/**
 * Embeds the fonts text is drawn in: the Standard-14 Helvetica and Courier faces, and before them the
 * fonts given in `pdfConfig.fonts` and `pdfConfig.fallbackFonts`. A style with no font of its own is
 * drawn in the regular one given, or in its Standard-14 font when that is not given either, the
 * fallback fonts drawing what it lacks; code keeps Courier unless a `mono` font is given. A font file that
 * cannot be read fails the conversion, naming which one: drawing in another font without a word would
 * be a different document. So does one fontkit cannot write into the PDF, when it is saved (see
 * guardEmbedding). A color font none of whose characters has an outline is reported, and not used.
 */
async function embedFonts(pdf: any, lib: any, pc: FullGeneratorConfig['pdfConfig'], config: OfficeParserAST['config'], reportWarning: (type: OfficeWarningType, info?: any) => void): Promise<{ fonts: Fonts; custom: boolean }> {
    const standard = (font: any): Face => ({ font, covers: winAnsiCovers, has: winAnsiCovers });
    const std: Record<FaceKey, Face> = {
        regular: standard(await pdf.embedFont(lib.StandardFonts.Helvetica)),
        bold: standard(await pdf.embedFont(lib.StandardFonts.HelveticaBold)),
        italic: standard(await pdf.embedFont(lib.StandardFonts.HelveticaOblique)),
        boldItalic: standard(await pdf.embedFont(lib.StandardFonts.HelveticaBoldOblique)),
        mono: standard(await pdf.embedFont(lib.StandardFonts.Courier)),
        monoBold: standard(await pdf.embedFont(lib.StandardFonts.CourierBold)),
        monoItalic: standard(await pdf.embedFont(lib.StandardFonts.CourierOblique)),
        monoBoldItalic: standard(await pdf.embedFont(lib.StandardFonts.CourierBoldOblique)),
        serif: standard(await pdf.embedFont(lib.StandardFonts.TimesRoman)),
        serifBold: standard(await pdf.embedFont(lib.StandardFonts.TimesRomanBold)),
        serifItalic: standard(await pdf.embedFont(lib.StandardFonts.TimesRomanItalic)),
        serifBoldItalic: standard(await pdf.embedFont(lib.StandardFonts.TimesRomanBoldItalic)),
    };
    const allKeys = Object.keys(std) as FaceKey[];
    const keys: GivenFaceKey[] = ['regular', 'bold', 'italic', 'boldItalic', 'mono'];
    // `fonts` names a file for each style. Given as anything else (a font's bytes alone, a file name) it is
    // reported and no style has one; a style is read only as the object's own key.
    const fontMap: unknown = pc.fonts;
    const isMap = !!fontMap && typeof fontMap === 'object' && !Array.isArray(fontMap) && !ArrayBuffer.isView(fontMap) && !(fontMap instanceof ArrayBuffer);
    if (!isMap && fontMap !== undefined && fontMap !== null) {
        const shown = typeof fontMap === 'string' ? fontMap : Object.prototype.toString.call(fontMap);
        reportWarning(OfficeWarningType.INVALID_CONFIG_VALUE, { option: 'pdfConfig.fonts', value: shown, expected: 'an object giving a font file for each style it names (regular, bold, italic, boldItalic, mono)', fallback: undefined });
    }
    const styleFont = (key: GivenFaceKey): unknown => isMap && Object.prototype.hasOwnProperty.call(fontMap, key) ? (fontMap as Record<string, unknown>)[key] : undefined;
    const given = keys.map(key => ({ name: `fonts.${key}`, key, bytes: fontBytes(styleFont(key), `fonts.${key}`, reportWarning) }));
    const fallbackList: unknown[] = Array.isArray(pc.fallbackFonts) ? pc.fallbackFonts : [];
    if (!Array.isArray(pc.fallbackFonts) && pc.fallbackFonts !== undefined && pc.fallbackFonts !== null) {
        reportWarning(OfficeWarningType.INVALID_CONFIG_VALUE, { option: 'pdfConfig.fallbackFonts', value: Object.prototype.toString.call(pc.fallbackFonts), expected: 'a list of font files (an array)', fallback: undefined });
    }
    const fallbackGiven = fallbackList.map((value, i) => ({ name: `fallbackFonts[${i}]`, bytes: fontBytes(value, `fallbackFonts[${i}]`, reportWarning) }));
    const custom = given.some(g => g.bytes) || fallbackGiven.some(g => g.bytes);
    if (!custom) return { fonts: Object.fromEntries(allKeys.map(key => [key, [std[key]]])) as Fonts, custom };

    const fontkit = await loadFontkit(config);
    pdf.registerFontkit(fontkit);
    // (Brotli, for a WOFF2 font: Node's where it has one, else the decoder fontkit carries.)
    let unbrotli: Unbrotli | undefined;
    const brotli = () => unbrotli ??= nativeUnbrotli() ?? fontkitUnbrotli(fontkit);
    // The same bytes given for two styles are embedded once: the same view, or views of the same part of
    // one buffer (an ArrayBuffer given twice is viewed anew each time).
    const embedded = new Map<ArrayBufferLike, Map<string, Face | null>>();
    const embed = async (name: string, bytes: Uint8Array | null): Promise<Face | undefined> => {
        if (!bytes) return undefined;
        const span = `${bytes.byteOffset}:${bytes.byteLength}`;
        let inBuffer = embedded.get(bytes.buffer);
        if (!inBuffer) embedded.set(bytes.buffer, inBuffer = new Map());
        const known = inBuffer.get(span);
        if (known !== undefined) return known ?? undefined;
        let face: Face;
        try {
            // A WOFF or WOFF2 font is embedded as the TrueType or OpenType font it holds: fontkit's subset
            // of a WOFF2 font, and either embedded as it is, are fonts viewers do not draw.
            const sfnt = sfntOf(bytes, brotli);
            // An OpenType font of CFF outlines (`OTTO`) is embedded whole: its subset (fontkit's) is not one
            // viewers draw (each glyph another's, or none).
            const cff = sfnt.length >= 4 && sfnt[0] === 0x4F && sfnt[1] === 0x54 && sfnt[2] === 0x54 && sfnt[3] === 0x4F;
            const outlines = trueTypeOutlines(sfnt);
            const color = colorTables(sfnt);
            // How a font draws a character: `drawn` when its glyph draws something, or the character is
            // blank by design (see BLANK); `empty` when its glyph draws nothing in a font that substitutes
            // glyphs, which only a combining mark drawn with its letter is drawn with (see Face.has). Any
            // other character goes to a font after it that draws it, or is lost: drawn blank, it showed
            // nothing while it was counted as drawn. A color font's empty glyph stands in for a color one,
            // drawn from tables a PDF does not hold, so it is not the font's at all, nor is a glyph that
            // cannot be read.
            const shapes = !color.length && substitutes(sfnt);
            const drawing = (fontkitFont: any, cp: number): 'drawn' | 'empty' | undefined => {
                if (isBlank(cp)) return 'drawn';
                let glyph: any;
                try { glyph = fontkitFont.glyphForCodePoint(cp); } catch { return undefined; }
                if (outlines) return outlines.draws(glyph?.id) ? 'drawn' : shapes && outlines.reads(glyph?.id) ? 'empty' : undefined;
                // (A CFF font's outlines, as fontkit reads them.)
                let draws: boolean;
                try { draws = glyph.path.commands.some((c: any) => c.command !== 'moveTo' && c.command !== 'closePath'); } catch { return undefined; }
                return draws ? 'drawn' : shapes ? 'empty' : undefined;
            };
            // A font of color glyphs alone draws none of its characters: it is reported, and not used.
            if (color.length) {
                const probe = fontkit.create(sfnt);
                if (!probe.characterSet.some((cp: number) => !isBlank(cp) && drawing(probe, cp) === 'drawn')) {
                    reportWarning(OfficeWarningType.INVALID_CONFIG_VALUE, { option: `pdfConfig.${name}`, value: `a color font (${color.map(t => t.trim()).join(', ')})`, expected: "a font of outlines: the native engine draws a font's outlines, and this font's glyphs are color images or layers", fallback: undefined });
                    inBuffer.set(span, null);
                    return undefined;
                }
            }
            // (fontkit subsets no font of color layers, COLR, whose glyphs it reads as layers alone.)
            const whole = cff || color.includes('COLR');
            const font = await pdf.embedFont(sfnt, { subset: pc.subsetFonts !== false && !whole });
            // Every glyph is placed by its width (hmtx, read with hhea and maxp): a font fontkit cannot read
            // them in failed halfway through the layout, naming no font.
            if (!Number.isFinite(font.embedder?.font?.getGlyph?.(0)?.advanceWidth)) throw new UnreadableFont('its glyph widths (hmtx, hhea, maxp) cannot be read');
            const set = new Set<number>(font.getCharacterSet());
            const drawings = new Map<number, 'drawn' | 'empty' | undefined>();
            const drawingOf = (cp: number): 'drawn' | 'empty' | undefined => {
                if (!set.has(cp)) return undefined;
                if (!drawings.has(cp)) drawings.set(cp, drawing(font.embedder.font, cp));
                return drawings.get(cp);
            };
            guardShaping(font);
            guardEmbedding(font, glyphs => outlines?.check(glyphs), error => (error as any)?.officeIssue ? error as Error
                : getOfficeError(OfficeErrorType.PDF_GENERATION_FAILED, config, `pdfConfig.${name} could not be embedded (${(error as any)?.message || String(error)})`));
            face = { font, covers: cp => drawingOf(cp) === 'drawn', has: cp => drawingOf(cp) !== undefined, code: glyphCoder(font, lib) };
        } catch (e: any) {
            throw getOfficeError(OfficeErrorType.PDF_GENERATION_FAILED, config, `pdfConfig.${name} could not be read as a TrueType, OpenType, WOFF or WOFF2 font (${e?.message || String(e)})`);
        }
        inBuffer.set(span, face);
        return face;
    };
    const own = {} as Record<GivenFaceKey, Face | undefined>;
    for (const g of given) own[g.key] = await embed(g.name, g.bytes);
    const fallbacks: Face[] = [];
    for (const g of fallbackGiven) { const face = await embed(g.name, g.bytes); if (face) fallbacks.push(face); }

    const chain = (key: FaceKey): Face[] => {
        // A style is drawn in its own font, or in the regular one given. With neither (and for code,
        // which stays in a fixed-width font unless a `mono` font is given) the Standard-14 font of the
        // style comes first, and the fonts given draw only what it lacks. Serif text, which no font
        // given is for, is drawn in Times, and in the fonts given where Times lacks a character.
        const variant = key.replace(/^(serif|mono)/, '');
        const sansKey = (variant ? variant[0].toLowerCase() + variant.slice(1) : 'regular') as GivenFaceKey;
        if (key.startsWith('serif')) {
            const order = [std[key], own[sansKey], own.regular, ...fallbacks];
            return order.filter((face, i): face is Face => !!face && order.indexOf(face) === i);
        }
        const given = key.startsWith('mono') ? own.mono : (own[sansKey] ?? own.regular);
        const order = given ? [given, own.regular, ...fallbacks, std[key]] : [std[key], own.regular, ...fallbacks];
        return order.filter((face, i): face is Face => !!face && order.indexOf(face) === i);
    };
    return { fonts: Object.fromEntries(allKeys.map(key => [key, chain(key)])) as Fonts, custom };
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
export async function renderNativePdf(ast: OfficeParserAST, config: FullGeneratorConfig, metadata: OfficeMetadata, report: (type: OfficeWarningType, info?: any) => void): Promise<Uint8Array> {
    // Each warning once: what is laid out twice (a balanced section, measured, then drawn) met the same
    // problem twice.
    const reported = new Set<string>();
    const reportWarning = (type: OfficeWarningType, info?: any) => {
        let key: string;
        try { key = `${type}|${JSON.stringify(info ?? null)}`; } catch { key = `${type}|${String(info)}`; }
        if (reported.has(key)) return;
        reported.add(key);
        report(type, info);
    };
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
    // 'auto' (the default): the document's own pages, each section's, where it sets them; else A4. A
    // width or height given is the caller's paper.
    const auto = String(pc.format ?? 'auto').toLowerCase() === 'auto' && !pc.width && !pc.height;
    const paper = paperSizePt(auto ? 'a4' : pc.format);
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
    // A side the caller left unset is the document's (under 'auto'), else the default.
    const unset = (v: string | number | undefined) => v === undefined || v === null || v === '';
    const marginFor = (setup: PageSetup | undefined) => {
        const side = (v: string | number | undefined, own: number | undefined) => unset(v) && auto && own !== undefined ? own : toPoints(v, mdef);
        return { top: side(pc.margin?.top, setup?.margin?.top), right: side(pc.margin?.right, setup?.margin?.right), bottom: side(pc.margin?.bottom, setup?.margin?.bottom), left: side(pc.margin?.left, setup?.margin?.left) };
    };
    // A section's page: its own size (turned to landscape when the caller asks), its own margins; one that
    // sets no page (a section that only numbers its pages its own way) is on the document's. Held as the
    // readers hold a page: a size that is no page (an inch to two hundred inches each way) is the
    // document's, a margin at most a third of the page: a section of a hundred million points by a
    // twentieth of one lost its text.
    const isPage = hasPageSize;
    const documentPage = auto && isPage(ast.metadata?.pageSetup) ? ast.metadata!.pageSetup! : undefined;
    const sectionPage = (given: PageSetup | undefined) => {
        const setup = onDocumentPage(given, documentPage), own = isPage(setup);
        let pw = own ? setup.width : w, ph = own ? setup.height : h;
        if (own && pc.landscape && pw < ph) [pw, ph] = [ph, pw];
        const most = Math.min(pw, ph) / 3, margins = marginFor(setup);
        const held = (v: number) => Math.min(Number.isFinite(v) && v > 0 ? v : 0, most);
        return { w: pw, h: ph, margin: { top: held(margins.top), right: held(margins.right), bottom: held(margins.bottom), left: held(margins.left) } };
    };
    const first = auto ? ast.metadata?.pageSetup : undefined;
    const firstPage = first ? sectionPage(first) : { w, h, margin: marginFor(undefined) };

    const layout = new NativeLayout(pdf, lib, fonts, firstPage.w, firstPage.h, firstPage.margin, config, ast, reportWarning, customFonts,
        auto ? sectionPage : null, first?.headerDistance, first?.footerDistance, ast.metadata?.pageSetup?.differentFirstPage);
    // Gather footnote/endnote bodies from the whole content tree first (they attach to text runs, deep
    // in the tree, so a render()-level sweep would miss them), then render the body, then draw them.
    await layout.fitRunningContent();
    for (const node of ast.content) layout.collectNotes(node);
    await layout.measureFootnotes();
    // The document's columns, its first section's, then its blocks (each section's own as it starts).
    layout.flow = ast.content;
    await layout.setColumns(ast.metadata?.pageSetup?.columns, true);
    for (const [index, node] of ast.content.entries()) { layout.flowIndex = index; await layout.render(node, true); }
    await layout.setColumns(undefined, false);
    await layout.flushNotes();
    // The anchors last met take their place on the last page, before the headers are drawn on the pages.
    layout.placeAnchors();
    // Each page's header and footer, once the pages are known (ast.auxiliary, where parsing kept them).
    await layout.drawRunningContent();
    layout.reportLostCharacters();
    layout.writeLinks();
    layout.dropScratchPage();
    if (pc.outline) writeOutline(pdf, lib, layout.outlineEntries);

    return await pdf.save();
}

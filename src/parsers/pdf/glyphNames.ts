/**
 * The text of a simple font's glyph by its name, where pdf.js reads none: a font whose `/Encoding`
 * `/Differences` name a code's glyph with a name pdf.js does not know ("c.sc", a small capital; "f_l", a
 * ligature), and whose `/ToUnicode` CMap has no entry for the code, reads in pdf.js as the code's own
 * character, so "nicht" in a form's subset font came out as "ni!ht". The name says what the glyph is, by the
 * Adobe Glyph List's rules, as Poppler reads it: what is after a period is a variant's suffix, an
 * underscore joins the components of a ligature, and each component is a name of the list, `uniXXXX` or
 * `uXXXX`.
 *
 * pdf.js keeps no font's encoding or CMap where its text is read (it can be asked to, `fontExtraProperties`,
 * but then sends every font's whole CMap, a CJK font's a sparse array billions long), so they are read here
 * from the file (see pdfFile), for the simple fonts (Type 1, Type 3, TrueType; never a composite font) of a
 * page that may have such a code, and each matched to pdf.js's font by the order the page sets them in:
 * the n-th `Tf` here is the n-th font pdf.js sets. A glyph is read by its name only where pdf.js read it
 * as its code's own character (see textGlyphs' collectGlyphs and measureItems).
 *
 * @module parsers/pdf/glyphNames
 */

import { ContentOp, contentStreams, formOf, Lexer, MAX_FORM_DEPTH, PdfDict, PdfFile, PdfName, PdfReader, PdfString, PdfValue, isStream } from './pdfFile.js';

/**
 * The names of the Adobe Glyph List a glyph's name is made of, beyond a single letter and a letter with an
 * accent (see ACCENTS), with their characters as the list gives them: the ASCII and Latin-1 marks, digits,
 * ligatures, typographic marks and the Greek letters.
 */
const GLYPH_LIST: Record<string, number> = {
    space: 0x20, exclam: 0x21, quotedbl: 0x22, numbersign: 0x23, dollar: 0x24, percent: 0x25, ampersand: 0x26,
    quotesingle: 0x27, parenleft: 0x28, parenright: 0x29, asterisk: 0x2a, plus: 0x2b, comma: 0x2c, hyphen: 0x2d,
    period: 0x2e, slash: 0x2f, zero: 0x30, one: 0x31, two: 0x32, three: 0x33, four: 0x34, five: 0x35, six: 0x36,
    seven: 0x37, eight: 0x38, nine: 0x39, colon: 0x3a, semicolon: 0x3b, less: 0x3c, equal: 0x3d, greater: 0x3e,
    question: 0x3f, at: 0x40, bracketleft: 0x5b, backslash: 0x5c, bracketright: 0x5d, asciicircum: 0x5e,
    underscore: 0x5f, grave: 0x60, braceleft: 0x7b, bar: 0x7c, braceright: 0x7d, asciitilde: 0x7e,
    exclamdown: 0xa1, cent: 0xa2, sterling: 0xa3, currency: 0xa4, yen: 0xa5, brokenbar: 0xa6, section: 0xa7,
    dieresis: 0xa8, copyright: 0xa9, ordfeminine: 0xaa, guillemotleft: 0xab, logicalnot: 0xac, registered: 0xae,
    macron: 0xaf, degree: 0xb0, plusminus: 0xb1, twosuperior: 0xb2, threesuperior: 0xb3, acute: 0xb4, mu: 0xb5,
    paragraph: 0xb6, periodcentered: 0xb7, cedilla: 0xb8, onesuperior: 0xb9, ordmasculine: 0xba,
    guillemotright: 0xbb, onequarter: 0xbc, onehalf: 0xbd, threequarters: 0xbe, questiondown: 0xbf, AE: 0xc6,
    Eth: 0xd0, multiply: 0xd7, Oslash: 0xd8, Thorn: 0xde, germandbls: 0xdf, ae: 0xe6, eth: 0xf0, divide: 0xf7,
    oslash: 0xf8, thorn: 0xfe, dotlessi: 0x131, Lslash: 0x141, lslash: 0x142, OE: 0x152, oe: 0x153,
    florin: 0x192, IJ: 0x132, ij: 0x133, Ldotaccent: 0x13f, ldotaccent: 0x140, dmacron: 0x111,
    circumflex: 0x2c6, caron: 0x2c7, breve: 0x2d8, dotaccent: 0x2d9, ring: 0x2da, ogonek: 0x2db, tilde: 0x2dc,
    hungarumlaut: 0x2dd, endash: 0x2013, emdash: 0x2014, quoteleft: 0x2018, quoteright: 0x2019,
    quotesinglbase: 0x201a, quotereversed: 0x201b, quotedblleft: 0x201c, quotedblright: 0x201d,
    quotedblbase: 0x201e, dagger: 0x2020, daggerdbl: 0x2021, bullet: 0x2022, ellipsis: 0x2026,
    perthousand: 0x2030, minute: 0x2032, second: 0x2033, guilsinglleft: 0x2039, guilsinglright: 0x203a,
    fraction: 0x2044, Euro: 0x20ac, trademark: 0x2122, minus: 0x2212, ff: 0xfb00, fi: 0xfb01, fl: 0xfb02,
    ffi: 0xfb03, ffl: 0xfb04, alpha: 0x3b1, beta: 0x3b2, gamma: 0x3b3, delta: 0x3b4, epsilon: 0x3b5,
    zeta: 0x3b6, eta: 0x3b7, theta: 0x3b8, iota: 0x3b9, kappa: 0x3ba, lambda: 0x3bb, nu: 0x3bd, xi: 0x3be,
    omicron: 0x3bf, pi: 0x3c0, rho: 0x3c1, sigma: 0x3c3, sigma1: 0x3c2, tau: 0x3c4, upsilon: 0x3c5, phi: 0x3c6,
    chi: 0x3c7, psi: 0x3c8, omega: 0x3c9, Alpha: 0x391, Beta: 0x392, Gamma: 0x393, Epsilon: 0x395, Zeta: 0x396,
    Eta: 0x397, Theta: 0x398, Iota: 0x399, Kappa: 0x39a, Lambda: 0x39b, Mu: 0x39c, Nu: 0x39d, Xi: 0x39e,
    Omicron: 0x39f, Pi: 0x3a0, Rho: 0x3a1, Sigma: 0x3a3, Tau: 0x3a4, Upsilon: 0x3a5, Phi: 0x3a6, Chi: 0x3a7,
    Psi: 0x3a8,
};

/**
 * The accents the list names a letter with ("eacute", "Scaron"), as combining marks: such a name is its
 * letter and its accent composed, where they compose to one character and the list has the name.
 */
const ACCENTS: Record<string, number> = {
    acute: 0x301, grave: 0x300, circumflex: 0x302, dieresis: 0x308, tilde: 0x303, ring: 0x30a, cedilla: 0x327,
    caron: 0x30c, macron: 0x304, breve: 0x306, ogonek: 0x328, dotaccent: 0x307, hungarumlaut: 0x30b,
};
/** A letter and an accent that compose, but whose name the list does not have. */
const NOT_LISTED = new Set(['Adotaccent', 'adotaccent', 'Ecedilla', 'ecedilla', 'Hcaron', 'hcaron', 'Ngrave', 'ngrave', 'Odotaccent', 'odotaccent', 'Ymacron', 'ymacron']);

/** One component of a glyph's name as text: a name of the list, `uniXXXX` (one or more), `uXXXX` to `uXXXXXX`; else undefined. */
function componentText(part: string): string | undefined {
    if (Object.prototype.hasOwnProperty.call(GLYPH_LIST, part)) return String.fromCodePoint(GLYPH_LIST[part]);
    if (/^[A-Za-z]$/.test(part)) return part;
    const accented = /^([A-Za-z])([a-z]+)$/.exec(part);
    if (accented && Object.prototype.hasOwnProperty.call(ACCENTS, accented[2]) && !NOT_LISTED.has(part)) {
        const composed = (accented[1] + String.fromCharCode(ACCENTS[accented[2]])).normalize('NFC');
        if (composed.length === 1) return composed;
    }
    // (Not a surrogate: the list's rules leave a name of one out.)
    const scalar = (v: number) => v <= 0x10ffff && !(v >= 0xd800 && v <= 0xdfff);
    const uni = /^uni((?:[0-9A-F]{4})+)$/.exec(part);
    if (uni) {
        let out = '';
        for (let i = 0; i < uni[1].length; i += 4) {
            const v = parseInt(uni[1].slice(i, i + 4), 16);
            if (!scalar(v)) return undefined;
            out += String.fromCharCode(v);
        }
        return out;
    }
    const u = /^u([0-9A-F]{4,6})$/.exec(part);
    if (u) { const v = parseInt(u[1], 16); return scalar(v) ? String.fromCodePoint(v) : undefined; }
    return undefined;
}

/** The longest glyph name read (a name of a font is short; a longer one is no glyph's). */
const MAX_NAME = 256;

/**
 * A glyph's name as text, by the Adobe Glyph List's rules: up to its first period (a variant's suffix,
 * "c.sc", "one.osf", left out), its components joined by underscores ("f_l") each read (see
 * componentText). Undefined where a component cannot be read, or nothing is left ".notdef".
 */
export function glyphNameText(name: string): string | undefined {
    if (name.length > MAX_NAME) return undefined;
    const base = name.split('.')[0];
    if (!base) return undefined;
    let out = '';
    for (const part of base.split('_')) {
        const text = componentText(part);
        if (text === undefined) return undefined;
        out += text;
    }
    return out || undefined;
}

/** The simple fonts, whose codes are one byte and whose encoding can have `/Differences`. */
const SIMPLE_FONTS = new Set(['Type1', 'MMType1', 'TrueType', 'Type3']);
/** At most this many tokens of one `/ToUnicode` CMap are read; a font whose CMap is longer is left as pdf.js reads it. */
const MAX_CMAP_TOKENS = 2_000_000;
/** At most this many fonts a page sets, read here: a page setting more is left as pdf.js reads it. */
const MAX_FONT_SETS = 500_000;

/** A font's codes read by their glyphs' names (code to text), with a key of them, equal for fonts that read alike. */
interface CodeTexts { texts: ReadonlyMap<number, string>; key: string }
const NONE: CodeTexts = { texts: new Map(), key: '' };

/**
 * The fonts of a document's pages whose codes pdf.js reads as their own characters where their glyphs'
 * names say otherwise, read from the file as each page asks (see codeTextsOf): each font, each resources
 * dictionary, and each form's fonts read once, within the file's bounds.
 */
export class GlyphNames {
    private readonly fonts = new Map<PdfDict, CodeTexts>();
    /** The codes each `/ToUnicode` CMap has an entry for (see covered), read once whatever fonts share it. */
    private readonly cmaps = new Map<PdfValue, Uint8Array | undefined>();
    /** The codes some font a resources dictionary reaches (its own, its forms') reads by name. */
    private readonly reached = new Map<PdfDict, ReadonlySet<number>>();
    /** A form's fonts in the order it sets them, with the resources it was last drawn with. */
    private readonly forms = new Map<PdfValue, { resources: PdfDict | undefined; fonts: (PdfDict | undefined)[] }>();

    constructor(private readonly file: PdfFile) {}

    /**
     * The codes of a page's fonts (index 0 the first page's) read by their glyphs' names: pdf.js's id of a
     * font (as `setFont` and a text item name it) to its codes' texts. Read only where a glyph of the page
     * is drawn with a code some font of it reads by name, and pdf.js read that glyph as the code's own
     * character; a font pdf.js sets where the page sets fonts that read differently is left out. Undefined
     * where there is none, or the page cannot be read here (see pdfFile), or its fonts are set otherwise
     * than pdf.js sets them.
     *
     * @param ops - the page's operator list, as read (see PdfParser's readOperatorList).
     */
    codeTextsOf(index: number, ops: any, OPS: any): Map<string, ReadonlyMap<number, string>> | undefined {
        try {
            const fnArray: number[] | undefined = ops?.fnArray, argsArray: any[] | undefined = ops?.argsArray;
            if (!fnArray || !argsArray) return undefined;
            const at = this.file.page(index);
            if (!at) return undefined;
            const { reader, page } = at;
            const codes = page.resources ? this.codesReached(reader, page.resources, 0) : undefined;
            if (!codes?.size) return undefined;
            // pdf.js's fonts in the order the page sets them, and whether a glyph is drawn with such a code and read as
            // its own character (an annotation's appearance, drawn after the page and set from its own resources, apart).
            const set: string[] = [];
            let annotation = 0, drawn = false;
            for (let i = 0; i < fnArray.length; i++) {
                const fn = fnArray[i];
                if (fn === OPS.beginAnnotation) annotation++;
                else if (fn === OPS.endAnnotation) annotation = Math.max(0, annotation - 1);
                else if (annotation) continue;
                else if (fn === OPS.setFont) set.push(typeof argsArray[i]?.[0] === 'string' ? argsArray[i][0] : '');
                else if (fn === OPS.showText && !drawn) {
                    const glyphs: unknown[] = Array.isArray(argsArray[i]?.[0]) ? argsArray[i][0] : [];
                    for (const g of glyphs) {
                        const glyph = g as { unicode?: unknown; originalCharCode?: unknown } | null;
                        if (glyph && typeof glyph === 'object' && typeof glyph.originalCharCode === 'number' && codes.has(glyph.originalCharCode)
                            && glyph.unicode === String.fromCharCode(glyph.originalCharCode)) { drawn = true; break; }
                    }
                }
            }
            if (!drawn) return undefined;
            const fonts = this.fontsSet(reader, page);
            if (!fonts || fonts.length !== set.length) return undefined;
            // Each of pdf.js's fonts to its codes' texts; one set where the page sets fonts that read differently, none.
            const byId = new Map<string, CodeTexts | null>();
            for (let k = 0; k < set.length; k++) {
                const texts = fonts[k] ? this.codeTexts(reader, fonts[k]!) : NONE;
                const known = byId.get(set[k]);
                if (known === undefined) byId.set(set[k], texts);
                else if (known && known.key !== texts.key) byId.set(set[k], null);
            }
            const out = new Map<string, ReadonlyMap<number, string>>();
            for (const [id, texts] of byId) if (id && texts?.texts.size) out.set(id, texts.texts);
            return out.size ? out : undefined;
        } catch {
            return undefined;
        }
    }

    /** The codes the fonts a resources dictionary reaches (its own, and its forms', `depth` forms in) read by name. */
    private codesReached(reader: PdfReader, resources: PdfDict, depth: number): ReadonlySet<number> {
        const known = this.reached.get(resources);
        if (known) return known;
        const codes = new Set<number>();
        // (Set before its forms are read: a form whose resources are its drawer's, or that draws itself, adds none.)
        this.reached.set(resources, codes);
        const fonts = reader.resolve(resources.get('Font'));
        if (fonts instanceof Map) {
            for (const value of fonts.values()) {
                const font = reader.resolve(value);
                if (font instanceof Map) for (const code of this.codeTexts(reader, font).texts.keys()) codes.add(code);
            }
        }
        const xobjects = depth < MAX_FORM_DEPTH ? reader.resolve(resources.get('XObject')) : undefined;
        if (xobjects instanceof Map) {
            for (const name of xobjects.keys()) {
                const drawn = formOf(reader, resources, name);
                if (drawn?.resources && drawn.resources !== resources) for (const code of this.codesReached(reader, drawn.resources, depth + 1)) codes.add(code);
            }
        }
        return codes;
    }

    /**
     * The fonts a page sets, in order, as pdf.js meets them (forms drawn with `Do` in place): each `Tf`'s
     * font dictionary, none where its name is no font of the resources. Undefined where a stream cannot be
     * read, or the page sets more than MAX_FONT_SETS.
     */
    private fontsSet(reader: PdfReader, page: { dict: PdfDict; resources: PdfDict | undefined }): (PdfDict | undefined)[] | undefined {
        let complete = true;
        const play = (ops: ContentOp[], resources: PdfDict | undefined, depth: number, active: Set<PdfValue>, out: (PdfDict | undefined)[]) => {
            for (const o of ops) {
                if (out.length > MAX_FONT_SETS) { complete = false; return; }
                if (o.op === 'Tf') {
                    const fonts = resources ? reader.resolve(resources.get('Font')) : undefined;
                    const font = fonts instanceof Map && o.name !== undefined ? reader.resolve(fonts.get(o.name)) : undefined;
                    out.push(font instanceof Map ? font : undefined);
                } else if (o.op === 'Do' && depth < MAX_FORM_DEPTH) {
                    const drawnForm = formOf(reader, resources, o.name);
                    if (!drawnForm || active.has(drawnForm.form)) continue;
                    const { form, resources: res } = drawnForm;
                    // Read once for the resources it is drawn with, and repeated where it is drawn again (see MarkedContent).
                    const cached = this.forms.get(form);
                    let drawn = cached && cached.resources === res ? cached.fonts : undefined;
                    if (!drawn) {
                        // (A form that cannot be read sets fonts pdf.js reads: the page's are not known.)
                        const body = this.file.program(reader, form);
                        if (!body) { complete = false; return; }
                        drawn = [];
                        active.add(form);
                        play(body, res, depth + 1, active, drawn);
                        active.delete(form);
                        if (!complete) return;
                        this.forms.set(form, { resources: res, fonts: drawn });
                    }
                    for (let k = 0; k < drawn.length && out.length <= MAX_FONT_SETS; k++) out.push(drawn[k]);
                }
            }
        };
        const out: (PdfDict | undefined)[] = [];
        for (const st of contentStreams(reader, page.dict)) {
            const ops = this.file.program(reader, st);
            if (!ops) return undefined;
            play(ops, page.resources, 0, new Set(), out);
        }
        return complete && out.length <= MAX_FONT_SETS ? out : undefined;
    }

    /**
     * A font's codes read by their glyphs' names: a simple font's codes its `/Differences` name with a name
     * that reads as text other than the code's own character, and that its `/ToUnicode` CMap has no entry
     * for. None for a composite font, for one whose CMap is a named one (pdf.js reads every code by it) or
     * cannot be read here; and none for a code whose own
     * character is a blank or a format character (pdf.js reads its glyph as a space, or leaves it out).
     */
    private codeTexts(reader: PdfReader, font: PdfDict): CodeTexts {
        const known = this.fonts.get(font);
        if (known) return known;
        this.fonts.set(font, NONE);
        const subtype = reader.resolve(font.get('Subtype'));
        if (!(subtype instanceof PdfName) || !SIMPLE_FONTS.has(subtype.name)) return NONE;
        const encoding = reader.resolve(font.get('Encoding'));
        const differences = encoding instanceof Map ? reader.resolve(encoding.get('Differences')) : undefined;
        if (!Array.isArray(differences)) return NONE;
        // The names, as pdf.js reads them: a number sets the next code, each name is the next code's (a later
        // name of a code its own). Every entry is walked, as the file's reader holds them (see pdfFile's Lexer),
        // and only the last name of each one-byte code kept.
        const named = new Map<number, string>();
        let code = 0;
        for (const entry of differences) {
            const value = reader.resolve(entry);
            if (typeof value === 'number') code = value;
            else if (value instanceof PdfName) {
                if (Number.isInteger(code) && code >= 0 && code <= 0xff) named.set(code, value.name);
                code++;
            }
        }
        const texts = new Map<number, string>();
        for (const [c, name] of named) {
            const own = String.fromCharCode(c);
            const text = name === '.notdef' ? undefined : glyphNameText(name);
            if (text !== undefined && text !== own && !/^[\s\p{Cf}\p{Mn}]$/u.test(own)) texts.set(c, text);
        }
        if (!texts.size) return NONE;
        const toUnicode = font.get('ToUnicode');
        if (toUnicode !== undefined) {
            const cmap = reader.resolve(toUnicode);
            if (cmap instanceof PdfName) return NONE;
            if (cmap && isStream(cmap)) {
                if (!this.cmaps.has(cmap)) this.cmaps.set(cmap, this.covered(reader, cmap));
                const covered = this.cmaps.get(cmap);
                if (!covered) return NONE;
                for (const c of [...texts.keys()]) if (covered[c]) texts.delete(c);
            }
        }
        const result = texts.size ? { texts, key: JSON.stringify([...texts]) } : NONE;
        this.fonts.set(font, result);
        return result;
    }

    /**
     * Which one-byte codes a `/ToUnicode` CMap has an entry for (`bfchar`, `bfrange`, and their `cid` forms,
     * as pdf.js reads them: a range of an array of destinations as far as the array reaches); undefined where
     * it cannot be read here (another filter, past the file's bounds, another CMap used, an entry not of
     * codes), so its font is left as pdf.js reads it. Each entry is a range of codes, counted in a difference
     * array: a CMap of a million ranges each over every code is read in time in its length.
     */
    private covered(reader: PdfReader, cmap: Parameters<PdfReader['decode']>[0]): Uint8Array | undefined {
        const data = reader.decode(cmap);
        if (!data) return undefined;
        const allowed = this.file.takeTokens(MAX_CMAP_TOKENS);
        const lx = new Lexer(data);
        const starts = new Int32Array(257);
        const codeOf = (v: unknown): number | undefined => {
            if (!(v instanceof PdfString) || !v.bytes.length || v.bytes.length > 4) return undefined;
            let n = 0;
            for (const b of v.bytes) n = n * 256 + b;
            return n;
        };
        const cover = (lo: number, hi: number) => {
            if (hi < lo || lo > 0xff || hi < 0) return;
            starts[Math.max(0, lo)]++;
            starts[Math.min(0xff, hi) + 1]--;
        };
        let mode: 'char' | 'range' | undefined, operands: unknown[] = [], tokens = 0, readable = true;
        for (; tokens < allowed && readable; tokens++) {
            const t = lx.next();
            if (t === undefined) break;
            if (t && typeof t === 'object' && 'op' in t) {
                const op = t.op;
                if (op === 'usecmap') readable = false;
                else if (op === 'beginbfchar' || op === 'begincidchar') mode = 'char';
                else if (op === 'beginbfrange' || op === 'begincidrange') mode = 'range';
                else if (mode && (op === 'endbfchar' || op === 'endcidchar' || op === 'endbfrange' || op === 'endcidrange')) {
                    if (operands.length) readable = false;
                    mode = undefined;
                } else if (mode) readable = false;
                operands = [];
                continue;
            }
            if (!mode) continue;
            operands.push(t);
            if (mode === 'char' && operands.length === 2) {
                const c = codeOf(operands[0]);
                if (c === undefined) readable = false; else cover(c, c);
                operands = [];
            } else if (mode === 'range' && operands.length === 3) {
                const lo = codeOf(operands[0]), hi = codeOf(operands[1]), dst = operands[2];
                if (lo === undefined || hi === undefined) readable = false;
                else cover(lo, Array.isArray(dst) ? Math.min(hi, lo + dst.length - 1) : hi);
                operands = [];
            }
        }
        this.file.returnTokens(allowed - tokens);
        if (!readable || tokens >= allowed) return undefined;
        const covered = new Uint8Array(256);
        let level = 0;
        for (let c = 0; c < 256; c++) { level += starts[c]; covered[c] = level > 0 ? 1 : 0; }
        return covered;
    }
}

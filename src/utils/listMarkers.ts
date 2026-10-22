/**
 * List markers as a document shows them: a number in its format (`a`, `iv`, `07`) and a bullet as the
 * character it draws. The readers record an item's marker (ListMetadata.marker) where it is other than
 * a plain `1.` or `•`, for the outputs that print it.
 */

/** How a list level writes its numbers. */
export type NumberFormat = 'decimal' | 'decimalZero' | 'lowerLetter' | 'upperLetter' | 'lowerRoman' | 'upperRoman' | 'none';

/** At most this many letters or numerals a number is written in: past it, it is written in digits. */
const MAX_NUMBER = 100_000;

/** A number in a list format: 1 is `a` or `i`, 27 is `aa` (as Word and LibreOffice letter them). */
export function formatListNumber(n: number, format: NumberFormat): string {
    if (format === 'none') return '';
    if (!Number.isFinite(n) || n < 1 || n > MAX_NUMBER) return String(Math.trunc(Number.isFinite(n) ? n : 0));
    switch (format) {
        case 'decimalZero': return n < 10 ? `0${n}` : String(n);
        case 'lowerLetter': case 'upperLetter': {
            // a..z, then aa..zz, then aaa..: the letter repeated, as both word processors do, as far as
            // ten letters; past that, digits (a long list's markers grew with its length, past any output).
            if (n > 260) return String(n);
            const letter = String.fromCharCode(97 + ((n - 1) % 26)).repeat(Math.ceil(n / 26));
            return format === 'upperLetter' ? letter.toUpperCase() : letter;
        }
        case 'lowerRoman': case 'upperRoman': {
            if (n >= 4000) return String(n);
            const numerals: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
            let rest = n, out = '';
            for (const [value, numeral] of numerals) while (rest >= value) { out += numeral; rest -= value; }
            return format === 'upperRoman' ? out.toUpperCase() : out;
        }
        default: return String(n);
    }
}

/** Word's `w:numFmt` values as a NumberFormat (others are written in digits). */
export function wordNumberFormat(numFmt: string | undefined): NumberFormat {
    switch (numFmt) {
        case 'lowerLetter': case 'upperLetter': case 'lowerRoman': case 'upperRoman': case 'decimalZero': case 'none': return numFmt;
        default: return 'decimal';
    }
}

/** ODF's `style:num-format` (`1`, `a`, `A`, `i`, `I`, or empty for none) as a NumberFormat. */
export function odfNumberFormat(numFormat: string | null | undefined): NumberFormat {
    switch (numFormat) {
        case 'a': return 'lowerLetter';
        case 'A': return 'upperLetter';
        case 'i': return 'lowerRoman';
        case 'I': return 'upperRoman';
        case '': case null: case undefined: return 'none';
        default: return 'decimal';
    }
}

/**
 * CSS `list-style-type` (or an `<ol type>`) as a NumberFormat, or a bullet's character, or a marker
 * written as a string (`'a) '`, with CSS escapes, as HTML output writes an item's own marker).
 */
export function cssListStyle(type: string): { format: NumberFormat } | { bullet: string } | undefined {
    const quoted = /^(['"])(.*)\1$/s.exec(type.trim());
    if (quoted) {
        const text = quoted[2].replace(/\\([0-9a-f]{1,6})\s?|\\(.)/gi, (_, hex: string | undefined, ch: string | undefined) => {
            if (!hex) return ch ?? '';
            const cp = parseInt(hex, 16);
            return cp > 0 && cp <= 0x10FFFF && !(cp >= 0xD800 && cp <= 0xDFFF) ? String.fromCodePoint(cp) : '';
        });
        return { bullet: text.trim() };
    }
    switch (type) {
        case '1': case 'decimal': return { format: 'decimal' };
        case 'decimal-leading-zero': return { format: 'decimalZero' };
        case 'a': case 'lower-alpha': case 'lower-latin': return { format: 'lowerLetter' };
        case 'A': case 'upper-alpha': case 'upper-latin': return { format: 'upperLetter' };
        case 'i': case 'lower-roman': return { format: 'lowerRoman' };
        case 'I': case 'upper-roman': return { format: 'upperRoman' };
        case 'disc': return { bullet: '•' };
        case 'circle': return { bullet: '◦' };
        case 'square': return { bullet: '▪' };
        case 'none': return { bullet: '' };
        default: return undefined;
    }
}

/**
 * The character a bullet draws. Word and LibreOffice draw most bullets in a symbol font, at a private-use
 * code point (U+F0B7 in Symbol is a round bullet, U+F0A7 in Wingdings a square), and a circle as `o` in Courier
 * New: each is the Unicode character it looks like. Another private-use character is a round bullet.
 */
export function bulletCharacter(char: string, font?: string): string {
    const cp = char.codePointAt(0);
    if (cp === undefined) return '';
    const f = (font || '').toLowerCase().trim();
    // A symbol font's character, at its own code or in the private-use row Word maps it to. (OpenSymbol,
    // LibreOffice's, holds Unicode's own characters: only its private-use ones are looked at below.)
    const symbolCode = cp >= 0xF000 && cp <= 0xF0FF ? cp - 0xF000 : cp < 0x100 ? cp : undefined;
    if (symbolCode !== undefined && f.startsWith('wingdings')) {
        const wingdings: Record<number, string> = { 0x6C: '●', 0x6E: '■', 0x71: '❑', 0x75: '◆', 0x76: '❖', 0xA7: '▪', 0xA8: '◻', 0xD8: '➢', 0xFC: '✓', 0xFB: '✗', 0x9F: '•' };
        return wingdings[symbolCode] ?? '•';
    }
    if (symbolCode !== undefined && f === 'symbol') return symbolCode === 0xA8 ? '♦' : '•';
    if (char === 'o' && /courier/.test(f)) return '◦';
    if (cp >= 0xE000 && cp <= 0xF8FF) return '•';
    return char;
}

/** How a list level writes its markers: a number in a format between a prefix and a suffix, or a bullet. */
export interface MarkerPattern {
    kind: 'ordered' | 'unordered';
    format: NumberFormat;
    /** Text before the number (`(` in `(a)`). */
    prefix: string;
    /** Text after the number (`)` in `a)`, `.` in `1.`). */
    suffix: string;
    /** The number follows its parent levels' (`1.2.3.`): each writer writes those as its own fields. */
    withParents: boolean;
    /** The parent levels' numbers as the marker shows them (`1.2.`), for a writer whose list numbers no parents. */
    parentText?: string;
    /** The bullet's character (`''` for none). */
    bullet: string;
}

/** Formats tried, in order, for the number in a marker. */
const PATTERN_FORMATS: NumberFormat[] = ['decimal', 'decimalZero', 'lowerRoman', 'upperRoman', 'lowerLetter', 'upperLetter'];

/**
 * A list level's pattern from one item's marker (ListMetadata.marker), its number (`itemIndex` + 1) and
 * its level: the format whose writing of the number stands in the marker as a word of its own, with the
 * text around it. An item without a marker numbers as `1.` or bullets as `•`; a marker in which the
 * number cannot be found numbers as `1.`.
 */
export function markerPattern(meta: { listType?: string; itemIndex?: number; marker?: unknown } | undefined, level: number): MarkerPattern {
    const ordered = meta?.listType === 'ordered';
    const plain: MarkerPattern = { kind: ordered ? 'ordered' : 'unordered', format: 'decimal', prefix: '', suffix: '.', withParents: false, bullet: '•' };
    const marker = typeof meta?.marker === 'string' ? meta.marker.slice(0, 64) : undefined;
    if (marker === undefined) return plain;
    if (!ordered) return { ...plain, bullet: marker };
    if (marker === '') return { ...plain, format: 'none', suffix: '' };
    const n = (typeof meta?.itemIndex === 'number' && Number.isFinite(meta.itemIndex) ? Math.max(0, Math.trunc(meta.itemIndex)) : 0) + 1;
    for (const format of PATTERN_FORMATS) {
        const written = formatListNumber(n, format);
        const at = marker.lastIndexOf(written);
        if (at < 0) continue;
        const before = marker[at - 1] ?? '', after = marker[at + written.length] ?? '';
        if (/[\p{L}\p{N}]/u.test(before) || /[\p{L}\p{N}]/u.test(after)) continue;
        const prefix = marker.slice(0, at), suffix = marker.slice(at + written.length);
        const withParents = level > 0 && new RegExp(`^(?:\\d+\\.){${level}}$`).test(prefix);
        return { kind: 'ordered', format, prefix: withParents ? '' : prefix, suffix, withParents, ...(withParents ? { parentText: prefix } : {}), bullet: '' };
    }
    return plain;
}

/** A NumberFormat as ODF's `style:num-format` (ODF has no zero-padded decimal: it is written plain). */
export function odfNumFormat(format: NumberFormat): string {
    switch (format) {
        case 'lowerLetter': return 'a';
        case 'upperLetter': return 'A';
        case 'lowerRoman': return 'i';
        case 'upperRoman': return 'I';
        case 'none': return '';
        default: return '1';
    }
}

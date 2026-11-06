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

/** What a marker's word can stand for, as a letter and as a numeral, or null (both when it can be either). */
export interface MarkerWord {
    letter: number | null;
    roman: number | null;
    /** Its case's formats. */
    upper: boolean;
}

/**
 * The word of a marker (`l`, `IV`, `aa`, the `l` of `(l)`) as what it can stand for: a letter or a letter
 * repeated (as `formatListNumber` letters), or a numeral written as `formatListNumber` writes it. In one
 * case only and written so, or it is neither: `Ill` and `Mix` are words, not numerals.
 */
export function markerWord(word: string): MarkerWord | null {
    const lower = word.toLowerCase();
    const upper = word === word.toUpperCase();
    if (!upper && word !== lower) return null;
    if (!/^[a-z]{1,10}$/.test(lower)) return null;
    const same = [...lower].every(ch => ch === lower[0]);
    const letter = same ? (lower.charCodeAt(0) - 96) + 26 * (lower.length - 1) : null;
    let roman: number | null = null;
    if (/^[ivxlcdm]+$/.test(lower)) {
        const values: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
        let total = 0, prev = 0;
        for (let i = lower.length - 1; i >= 0; i--) {
            const v = values[lower[i]];
            total += v < prev ? -v : v;
            prev = Math.max(prev, v);
        }
        if (total > 0 && formatListNumber(total, 'lowerRoman') === lower) roman = total;
    }
    return letter === null && roman === null ? null : { letter, roman, upper };
}

/** Numerals read with nothing around them to say so: past this a word is more likely a word (`mi`, `cc`). */
const LONE_NUMERAL_LIMIT = 40;

/** Where a run of markers of one kind has reached, as letters and as numerals, and which was read last. */
export interface MarkerRun {
    letter: number;
    roman: number;
    last: 'letter' | 'roman' | null;
}

/**
 * What a marker's word stands for, from the markers of its kind around it (`run`, where those before it
 * reached, and `next`, the word after it): `l` after `k` is 12, in `xlix`, `l`, `li` it is 50. Where they do
 * not say, a lone letter is a letter (`c` is 3, not 100), but `i` starts numerals, as does a short numeral.
 * A word that is only something past that (`cc`, `aa`, `mi`) needs a neighbour to say so, or is null.
 * `run` is moved on past the word.
 */
export function readMarkerWord(word: MarkerWord, run: MarkerRun, next: MarkerWord | null | undefined): number | null {
    const followed = word.roman !== null && next?.roman === word.roman + 1;
    const letterFits = word.letter !== null && (word.letter === run.letter + 1 || next?.letter === word.letter + 1);
    const romanFits = word.roman !== null && word.roman === run.roman + 1;
    const as = (reading: 'letter' | 'roman'): number => {
        const n = (reading === 'letter' ? word.letter : word.roman)!;
        run[reading] = n;
        run.last = reading;
        return n;
    };
    if (followed) return as('roman');
    if (letterFits && romanFits) return as(run.last ?? 'letter');
    if (letterFits) return as('letter');
    if (romanFits) return as('roman');
    if (word.letter !== null && word.roman === null && word.letter <= 26) return as('letter');
    if (word.roman === 1) return as('roman');
    if (word.letter !== null && word.letter <= 26) return as('letter');
    if (word.roman !== null && word.roman < LONE_NUMERAL_LIMIT) return as('roman');
    return null;
}

/** A marker's word out of its label (`(l)`, `iv.`, `a)`), or null where it has none of letters alone. */
export function markerLabelWord(label: string): string | null {
    return /^[(\[]*([a-zA-Z]{1,10})[.)\]]*$/.exec(label.trim())?.[1] ?? null;
}

/**
 * The numbers of a list's labels in order (null where a label has no letter or numeral word), each kind
 * (its brackets and case) read as its own run, from `readMarkerWord`.
 */
export function readMarkerLabels(labels: string[]): (number | null)[] {
    const words = labels.map(label => { const w = markerLabelWord(label); return w ? markerWord(w) : null; });
    const kinds = labels.map((label, i) => words[i] ? `${label.trim().replace(/[a-zA-Z]+/, '')}|${words[i]!.upper}` : '');
    const nexts = new Array<MarkerWord | null>(labels.length).fill(null);
    const seen = new Map<string, MarkerWord>();
    for (let i = labels.length - 1; i >= 0; i--) {
        if (!words[i]) continue;
        nexts[i] = seen.get(kinds[i]) ?? null;
        seen.set(kinds[i], words[i]!);
    }
    const runs = new Map<string, MarkerRun>();
    return labels.map((_, i) => {
        if (!words[i]) return null;
        let run = runs.get(kinds[i]);
        if (!run) runs.set(kinds[i], run = { letter: 0, roman: 0, last: null });
        return readMarkerWord(words[i]!, run, nexts[i]);
    });
}

/** Word's `w:numFmt` values as a NumberFormat (others are written in digits). */
export function wordNumberFormat(numFmt: string | undefined): NumberFormat {
    switch (numFmt) {
        case 'lowerLetter': case 'upperLetter': case 'lowerRoman': case 'upperRoman': case 'decimalZero': case 'none': return numFmt;
        default: return 'decimal';
    }
}

/**
 * The most letters a note's mark is written in with `style:num-letter-sync` (LibreOffice writes the letter
 * n / 26 times, rounded up, so its marks grow with their number): as LibreOffice to 64 letters (note 1664),
 * digits past that, so a start value up to 999999 cannot make each mark thousands of characters long.
 */
const MAX_SYNCED_NOTE_LETTERS = 64;

/**
 * A note's number as LibreOffice writes it in an ODF `style:num-format` (`1`, `a`, `A`, `i`, `I`; empty
 * is `1`), or undefined for any other format, which only the writing application knows. Letters run
 * a..z, aa, ab, .. (as a spreadsheet's columns), or with `style:num-letter-sync` a..z, aa, bb, .., the
 * letter written once more for each 26 (`oooooooooooo` for 301), up to MAX_SYNCED_NOTE_LETTERS letters;
 * numerals past 3999 repeat the thousands (`mmmm`). Past MAX_NUMBER, digits: a mark stays short whatever
 * the start value.
 */
export function odfNoteNumber(n: number, numFormat: string, letterSync: boolean): string | undefined {
    if (numFormat === '' || numFormat === '1') return String(n);
    if (!/^[aAiI]$/.test(numFormat)) return undefined;
    if (!Number.isInteger(n) || n < 1 || n > MAX_NUMBER) return String(n);
    const upper = numFormat === 'A' || numFormat === 'I';
    if (numFormat === 'i' || numFormat === 'I') {
        const text = 'm'.repeat(Math.floor(n / 1000)) + (n % 1000 ? formatListNumber(n % 1000, 'lowerRoman') : '');
        return upper ? text.toUpperCase() : text;
    }
    if (letterSync) {
        const count = Math.ceil(n / 26);
        if (count > MAX_SYNCED_NOTE_LETTERS) return String(n);
        const text = String.fromCharCode(97 + ((n - 1) % 26)).repeat(count);
        return upper ? text.toUpperCase() : text;
    }
    let text = '';
    for (let rest = n; rest > 0; rest = Math.floor((rest - 1) / 26)) text = String.fromCharCode(97 + ((rest - 1) % 26)) + text;
    return upper ? text.toUpperCase() : text;
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

/** How much of a marker's label an output writes: one label can stand for every item of a large list. */
export const MARKER_LABEL_LIMIT = 255;

/** A marker's label as an output writes it. */
export function writtenMarkerLabel(marker: string): string {
    return [...marker.replace(/\s+/g, ' ').trim().slice(0, MARKER_LABEL_LIMIT * 2)].slice(0, MARKER_LABEL_LIMIT).join('');
}

/**
 * A marker as CSS `list-style-type`: a string in which anything other than a letter, a digit or a space is
 * escaped, so nothing in it can end the string or an attribute around it.
 */
export function cssListMarker(marker: string): string {
    const escaped = [...marker].slice(0, MARKER_LABEL_LIMIT).map(ch => /^[\p{L}\p{N} ]$/u.test(ch) ? ch : `\\${ch.codePointAt(0)!.toString(16)} `).join('');
    return `'${escaped}${marker ? ' ' : ''}'`;
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

/**
 * Page numbers as documents write them: the format a page field or a section sets (read from Word's
 * field switches and `w:pgNumType`, RTF's switches and section words, ODF's `style:num-format`), and a
 * number written in one (see TextMetadata.fieldFormat, PageSetup.pageNumberFormat).
 */
import type { PageNumberFormat } from '../types.js';
import { formatListNumber } from './listMarkers.js';

const ONES = ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'eleven', 'twelve',
    'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen'];
const TENS = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];

/** A number below a million in English words, as Word's CardText writes it (`one hundred twenty-three`). */
function cardinal(n: number): string {
    if (n < 20) return ONES[n];
    if (n < 100) return TENS[Math.floor(n / 10)] + (n % 10 ? `-${ONES[n % 10]}` : '');
    if (n < 1000) return `${ONES[Math.floor(n / 100)]} hundred${n % 100 ? ` ${cardinal(n % 100)}` : ''}`;
    return `${cardinal(Math.floor(n / 1000))} thousand${n % 1000 ? ` ${cardinal(n % 1000)}` : ''}`;
}

/** The ordinal of a cardinal in words: its last word made an ordinal (`twenty-first`, `one hundredth`). */
function ordinalWords(words: string): string {
    const irregular: Record<string, string> = { one: 'first', two: 'second', three: 'third', five: 'fifth', eight: 'eighth', nine: 'ninth', twelve: 'twelfth' };
    return words.replace(/[a-z]+$/, word => irregular[word] ?? (word.endsWith('y') ? `${word.slice(0, -1)}ieth` : `${word}th`));
}

/** `n` written in a page number format: digits where the format cannot write it (0, negatives, a million or more). */
export function pageNumberText(n: number, format: PageNumberFormat | undefined): string {
    const whole = Number.isFinite(n) ? Math.trunc(n) : 0;
    if (!format || format === 'decimal' || whole < 1) return String(whole);
    switch (format) {
        case 'lowerRoman': case 'upperRoman': case 'lowerLetter': case 'upperLetter': return formatListNumber(whole, format);
        case 'decimalDash': return `- ${whole} -`;
        case 'hex': return whole.toString(16).toUpperCase();
        case 'ordinal': {
            const tens = whole % 100;
            return `${whole}${tens >= 11 && tens <= 13 ? 'th' : ['th', 'st', 'nd', 'rd'][whole % 10] ?? 'th'}`;
        }
        case 'cardinalText': return whole < 1_000_000 ? cardinal(whole) : String(whole);
        case 'ordinalText': return whole < 1_000_000 ? ordinalWords(cardinal(whole)) : String(whole);
    }
}

/**
 * The format a field's instruction sets for its number (`PAGE \* roman`, `NUMPAGES \* ALPHABETIC`), as
 * Word and RTF write it: `roman` and `alphabetic` lower case, `ROMAN` and `ALPHABETIC` upper; the other
 * switches in any case. Undefined where it sets none (or one of none of these: `\* MERGEFORMAT`).
 */
export function fieldSwitchFormat(instruction: string): PageNumberFormat | undefined {
    for (const [, name] of instruction.matchAll(/\\\*\s*([A-Za-z]+)/g)) {
        if (name === 'roman') return 'lowerRoman';
        if (name === 'ROMAN') return 'upperRoman';
        if (name === 'alphabetic') return 'lowerLetter';
        if (name === 'ALPHABETIC') return 'upperLetter';
        switch (name.toLowerCase()) {
            case 'arabic': return 'decimal';
            case 'arabicdash': return 'decimalDash';
            case 'ordinal': return 'ordinal';
            case 'cardtext': return 'cardinalText';
            case 'ordtext': return 'ordinalText';
            case 'hex': return 'hex';
        }
    }
    return undefined;
}

/**
 * A Word section's page number format (`w:pgNumType w:fmt`): decimal where it gives none, undefined for
 * one of none of these.
 */
export function wordPageNumberFormat(fmt: string | null | undefined): PageNumberFormat | undefined {
    switch (fmt) {
        case null: case undefined: case '': case 'decimal': return 'decimal';
        case 'lowerRoman': case 'upperRoman': case 'lowerLetter': case 'upperLetter': case 'ordinal': case 'cardinalText': case 'ordinalText': case 'hex':
            return fmt;
        case 'numberInDash': return 'decimalDash';
        default: return undefined;
    }
}

/**
 * An ODF number format (`style:num-format`: `1`, `i`, `I`, `a`, `A`): decimal where it gives none,
 * undefined for one of none of these.
 */
export function odfPageNumberFormat(numFormat: string | null | undefined): PageNumberFormat | undefined {
    switch (numFormat) {
        case null: case undefined: case '': case '1': return 'decimal';
        case 'i': return 'lowerRoman';
        case 'I': return 'upperRoman';
        case 'a': return 'lowerLetter';
        case 'A': return 'upperLetter';
        default: return undefined;
    }
}

/** A roman numeral as a numeral is written: thousands, hundreds, tens and ones, each in its order. */
const ROMAN = /^m{0,4}(cm|cd|d?c{0,3})(xc|xl|l?x{0,3})(ix|iv|v?i{0,3})$/i;

/**
 * A page field's saved result, split into its runs' texts, with the number for the page being printed
 * written in: in `format` where the field or its section gives one, else in the form the saved number is
 * in. The saved number is found in the runs as a whole (a number saved in two looks is one number), and
 * the new one takes its place in the run its first character was in; what is around it stays (`- 3 -`,
 * `Page`). Read from the saved text, a number is digits of any script (written back in that script), a
 * roman numeral (in its case; a letter alone only as i, v or x), or a letter repeated (Word's
 * `ALPHABETIC`: 27 is `AA`); a result saved
 * empty is the number, and one of none of these (`three`, a format the field has and the reader did not
 * read) is replaced whole by the number in digits: the right page in another form, not the saved page on
 * every page.
 */
export function fillPageNumber(texts: string[], value: number, format: PageNumberFormat | undefined): string[] {
    const saved = texts.join('');
    let start = -1, end = -1, inferred: PageNumberFormat | undefined, zero: number | undefined;
    const digits = /\p{Nd}+/u.exec(saved);
    if (digits) {
        start = digits.index; end = start + digits[0].length;
        // The script's zero: Unicode's digits come in whole runs of ten from a zero, some runs one after
        // another (the mathematical digits, five), so the zero is ten in a row back from the stretch's start.
        const cp = digits[0].codePointAt(0)!;
        let first = cp;
        while (first > cp - 64 && /\p{Nd}/u.test(String.fromCodePoint(first - 1))) first--;
        zero = cp - ((cp - first) % 10);
        if (zero === 0x30) zero = undefined;
    } else {
        for (const word of saved.matchAll(/\p{L}+/gu)) {
            const w = word[0];
            // (A letter alone is a roman numeral only as i, v or x, a front matter's first pages: c, d, l
            // and m are far likelier pages 3, 4, 12 and 13 in letters than 100, 500, 50 and 1000.)
            if (w.length === 1 && !/[ivx]/i.test(w)) continue;
            if (ROMAN.test(w) && (w === w.toLowerCase() || w === w.toUpperCase())) {
                start = word.index!; end = start + w.length;
                inferred = w === w.toLowerCase() ? 'lowerRoman' : 'upperRoman';
                break;
            }
        }
        for (const word of start < 0 ? saved.matchAll(/\p{L}+/gu) : []) {
            const w = word[0];
            if (!/^([A-Za-z])\1*$/.test(w)) continue;
            start = word.index!; end = start + w.length;
            inferred = w === w.toLowerCase() ? 'lowerLetter' : 'upperLetter';
            break;
        }
        // Saved empty, or as nothing read as a number: the whole result is the number.
        if (start < 0) { start = 0; end = saved.length; }
    }
    let written = pageNumberText(value, format ?? inferred);
    if (zero !== undefined && !format) written = written.replace(/\d/g, d => String.fromCodePoint(zero! + Number(d)));
    // Each run its text outside the saved number, the number in the run its first character was in.
    const out: string[] = [];
    let at = 0, placed = false;
    texts.forEach((text, i) => {
        const a = at, b = at + text.length;
        at = b;
        const before = text.slice(0, Math.max(0, Math.min(text.length, start - a)));
        const after = text.slice(Math.max(0, Math.min(text.length, end - a)));
        const holds = !placed && (start < b || (start === b && i === texts.length - 1)) && start >= a;
        if (holds) placed = true;
        out.push(before + (holds ? written : '') + after);
    });
    if (!placed) out[out.length - 1] = (out[out.length - 1] ?? '') + written;
    return out;
}

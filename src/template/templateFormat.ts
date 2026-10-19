/**
 * Formatting a template's values with pictures, as Word's mail merge does: a number picture
 * (`$#,##0.00`, Word's `\#` switch, Excel's number formats) or a date picture (`MMMM d, yyyy`, Word's
 * `\@` switch). The formatting is this library's own, not the runtime's `Intl`, so a template renders
 * the same on every Node version and in every browser; the separators and the month and day names are
 * the caller's to set.
 *
 * @module template/templateFormat
 */

import { DOUBLE_QUOTES, SINGLE_QUOTES } from './templateTags.js';

/** The names a date picture writes, and the separators a number picture writes. */
export interface PictureSettings {
    decimal: string;
    group: string;
    months: readonly string[];
    monthsShort: readonly string[];
    days: readonly string[];
    daysShort: readonly string[];
    am: string;
    pm: string;
    /** Minutes added to UTC before a date is written (0: dates are written in UTC). */
    utcOffsetMinutes: number;
}

/** A number picture's piece: literal text, a digit placeholder, or the decimal point. */
type NumberToken = { t: 'text'; text: string } | { t: 'digit'; zero: boolean } | { t: 'point' };

/** One section of a number picture (positive; negative; zero). */
interface NumberSection { tokens: NumberToken[]; grouping: boolean; percents: number; fractionDigits: number }

/** A date picture's piece: literal text, or a part of the date. */
type DateToken = { t: 'text'; text: string } | { t: 'part'; part: DatePart };
type DatePart = 'd' | 'dd' | 'ddd' | 'dddd' | 'M' | 'MM' | 'MMM' | 'MMMM' | 'yy' | 'yyyy'
    | 'h' | 'hh' | 'H' | 'HH' | 'm' | 'mm' | 's' | 'ss' | 'AM/PM' | 'am/pm' | 'A/P' | 'a/p';

/** A picture, read once when the template is compiled. */
export type Picture =
    | { kind: 'number'; sections: NumberSection[] }
    | { kind: 'date'; tokens: DateToken[] };

/** A picture read, or why it cannot be. */
export type PictureReading = { ok: true; picture: Picture } | { ok: false; reason: string };

/** The most sections a number picture has: positive, negative, zero. */
const MAX_SECTIONS = 3;

/** Significant digits a number is written with, as Excel shows them: past 15, a double's digits are noise. */
const SIGNIFICANT_DIGITS = 15;

/**
 * The literal text a picture holds at `at`: `'quoted text'`, a `"quoted"` one, or one character after
 * `\`. Quotes may be straight or curly, as Word's AutoCorrect writes them inside a picture too; a quote
 * closes at the next of its family. Returns the text and where reading goes on, or null when `at` starts
 * no literal; `false` when a quote is never closed.
 */
function readLiteral(picture: string, at: number): { text: string; next: number } | null | false {
    const c = picture[at];
    if (c === '\\') return at + 1 < picture.length ? { text: picture[at + 1], next: at + 2 } : { text: '', next: at + 1 };
    const family = DOUBLE_QUOTES.includes(c) ? DOUBLE_QUOTES : SINGLE_QUOTES.includes(c) ? SINGLE_QUOTES : null;
    if (family) {
        let close = at + 1;
        while (close < picture.length && !family.includes(picture[close])) close++;
        if (close >= picture.length) return false;
        return { text: picture.slice(at + 1, close), next: close + 1 };
    }
    return null;
}

/** Whether a picture holds a digit placeholder (`0` or `#`) outside its literal text: a number picture. */
function isNumberPicture(picture: string): boolean {
    for (let at = 0; at < picture.length;) {
        const literal = readLiteral(picture, at);
        if (literal === false) return false;
        if (literal) { at = literal.next; continue; }
        if (picture[at] === '0' || picture[at] === '#') return true;
        at++;
    }
    return false;
}

function readNumberSection(text: string): NumberSection | string {
    const tokens: NumberToken[] = [];
    let grouping = false, percents = 0, fractionDigits = 0, seenPoint = false;
    // A comma after the last integer digit is Excel's thousands scaling (`0,,` for millions), which a Word
    // picture does not have: read as grouping, it wrote 1,234,567 where Excel writes 1.
    let commaAfterDigit = false;
    const scaling = 'has a comma after its last digit, Excel\'s thousands scaling, which a picture does not take';
    const literal = (s: string): void => {
        const last = tokens[tokens.length - 1];
        if (last?.t === 'text') last.text += s; else tokens.push({ t: 'text', text: s });
    };
    for (let at = 0; at < text.length;) {
        const lit = readLiteral(text, at);
        if (lit === false) return 'has a quote that is never closed';
        if (lit) { literal(lit.text); at = lit.next; continue; }
        const c = text[at++];
        if (c === '0' || c === '#') {
            commaAfterDigit = false;
            tokens.push({ t: 'digit', zero: c === '0' });
            if (seenPoint) fractionDigits++;
        } else if (c === '.' && !seenPoint) {
            if (commaAfterDigit) return scaling;
            seenPoint = true;
            tokens.push({ t: 'point' });
        } else if (c === ',' && !seenPoint) {
            grouping = true;
            if (tokens.some(t => t.t === 'digit')) commaAfterDigit = true;
        } else if (c === ',' && tokens.some(t => t.t === 'digit')) {
            // After the point, a comma that no digit follows scales too (`0.00,`); one between digits is text.
            commaAfterDigit = true;
            literal(c);
        } else if (c === '%') {
            percents++;
            literal('%');
        } else if (c === '[') {
            return 'uses [...], a colour or condition, which a picture does not take';
        } else if ((c === 'E' || c === 'e') && (text[at] === '+' || text[at] === '-') && tokens.some(t => t.t === 'digit')) {
            return 'uses scientific notation (E+ or E-), which a picture does not take';
        } else {
            literal(c);
        }
    }
    if (commaAfterDigit) return scaling;
    return { tokens, grouping, percents, fractionDigits };
}

/** Reads a date picture: Word's `\@` parts (`M` a month, `m` a minute; `d`, `y`, `s` in either case). */
function readDatePicture(picture: string): DateToken[] | string {
    const tokens: DateToken[] = [];
    const literal = (s: string): void => {
        const last = tokens[tokens.length - 1];
        if (last?.t === 'text') last.text += s; else tokens.push({ t: 'text', text: s });
    };
    for (let at = 0; at < picture.length;) {
        const lit = readLiteral(picture, at);
        if (lit === false) return 'has a quote that is never closed';
        if (lit) { literal(lit.text); at = lit.next; continue; }
        const marker = ['AM/PM', 'am/pm', 'A/P', 'a/p'].find(m => picture.startsWith(m, at)) as DatePart | undefined;
        if (marker) { tokens.push({ t: 'part', part: marker }); at += marker.length; continue; }
        const c = picture[at];
        const unit = c === 'D' ? 'd' : c === 'Y' ? 'y' : c === 'S' ? 's' : c;
        if (!'dMyhHms'.includes(unit)) { literal(c); at++; continue; }
        let run = 1;
        while (at + run < picture.length && picture[at + run] === c) run++;
        at += run;
        const part: DatePart =
            unit === 'd' ? (['d', 'dd', 'ddd', 'dddd'] as const)[Math.min(run, 4) - 1]
                : unit === 'M' ? (['M', 'MM', 'MMM', 'MMMM'] as const)[Math.min(run, 4) - 1]
                    : unit === 'y' ? (run <= 2 ? 'yy' : 'yyyy')
                        : (run === 1 ? unit : unit + unit) as DatePart;
        tokens.push({ t: 'part', part });
    }
    return tokens.some(t => t.t === 'part') ? tokens : 'is a picture of neither a number (0, #) nor a date (d, M, y, h, m, s)';
}

/** Reads a picture once, when the template is compiled. */
export function readPicture(picture: string): PictureReading {
    if (!picture) return { ok: false, reason: 'is empty' };
    if (isNumberPicture(picture)) {
        const sections: NumberSection[] = [];
        let start = 0;
        // Sections are split at `;` outside literal text.
        for (let at = 0; at <= picture.length;) {
            const lit = at < picture.length ? readLiteral(picture, at) : null;
            if (lit === false) return { ok: false, reason: 'has a quote that is never closed' };
            if (lit) { at = lit.next; continue; }
            if (at === picture.length || picture[at] === ';') {
                const section = readNumberSection(picture.slice(start, at));
                if (typeof section === 'string') return { ok: false, reason: section };
                sections.push(section);
                if (sections.length > MAX_SECTIONS) return { ok: false, reason: `has more than ${MAX_SECTIONS} sections (positive;negative;zero)` };
                start = at + 1;
            }
            at++;
        }
        return { ok: true, picture: { kind: 'number', sections } };
    }
    const tokens = readDatePicture(picture);
    return typeof tokens === 'string' ? { ok: false, reason: tokens } : { ok: true, picture: { kind: 'date', tokens } };
}

/** A number's digits as written: the integer part without leading zeros, and the fraction part. */
function decimalDigits(value: number): { integer: string; fraction: string } {
    // `toPrecision` is exactly specified, so the digits are the same on every runtime; 15 significant
    // digits drop a double's binary noise (0.1 + 0.2 is 0.3, 1.005 is 1.005).
    const text = value.toPrecision(SIGNIFICANT_DIGITS);
    const [mantissa, exponentText] = text.split('e');
    const exponent = exponentText ? Number(exponentText) : 0;
    const [whole, part = ''] = mantissa.split('.');
    let digits = whole + part;
    let point = whole.length + exponent;
    if (point < 0) { digits = '0'.repeat(-point) + digits; point = 0; }
    if (point > digits.length) digits += '0'.repeat(point - digits.length);
    // Trailing zeros cut by index: `/0+$/` retried from every zero of a tiny number's fraction (5e-324
    // has 323 before its digit), time quadratic in its length.
    let end = digits.length;
    while (end > point && digits.charCodeAt(end - 1) === 0x30) end--;
    return { integer: digits.slice(0, point).replace(/^0+/, ''), fraction: digits.slice(point, end) };
}

/** Rounds decimal digits to `places` fraction digits, half away from zero, as Word and Excel do. */
function roundDigits(integer: string, fraction: string, places: number): { integer: string; fraction: string } {
    if (fraction.length <= places) return { integer, fraction: fraction.padEnd(places, '0') };
    const kept = (integer + fraction.slice(0, places)).split('');
    if (fraction.charCodeAt(places) >= 0x35) { // the first digit dropped is 5 or more
        let i = kept.length - 1;
        while (i >= 0 && kept[i] === '9') kept[i--] = '0';
        if (i >= 0) kept[i] = String(Number(kept[i]) + 1); else kept.unshift('1');
    }
    const integerLength = kept.length - places;
    return { integer: kept.slice(0, integerLength).join('').replace(/^0+/, ''), fraction: kept.slice(integerLength).join('') };
}

/** Writes a number with a number picture; null when it is too large to write (past a double's range). */
export function formatNumber(value: number, picture: Extract<Picture, { kind: 'number' }>, settings: PictureSettings): string | null {
    const { sections } = picture;
    // The section: the second for a negative number, the third for zero, else the first (which writes a
    // negative number's sign itself).
    let section = sections[0];
    let sign = value < 0;
    if (value < 0 && sections.length >= 2) { section = sections[1]; sign = false; }
    else if (value === 0 && sections.length === 3) section = sections[2];
    // A section of literal text only (`0;'negative'`) writes just that: no digits, no sign.
    if (!section.tokens.some(t => t.t === 'digit')) return section.tokens.map(t => t.t === 'text' ? t.text : '').join('');
    let scaled = Math.abs(value);
    for (let i = 0; i < section.percents; i++) scaled *= 100;
    if (!Number.isFinite(scaled)) return null;
    const exact = decimalDigits(scaled);
    const { integer, fraction } = roundDigits(exact.integer, exact.fraction, section.fractionDigits);

    const pointAt = section.tokens.findIndex(t => t.t === 'point');
    let integerTokens = pointAt < 0 ? section.tokens : section.tokens.slice(0, pointAt);
    // A picture with no integer placeholder (`.00`) still writes the integer digits, before its point.
    if (!integerTokens.some(t => t.t === 'digit')) integerTokens = [...integerTokens, { t: 'digit', zero: false }];
    const fractionTokens = pointAt < 0 ? [] : section.tokens.slice(pointAt + 1);

    // The integer part, written right to left: each placeholder takes a digit, the leftmost all that
    // remain; a `0` with no digit left writes 0, a `#` nothing. Groups of three are separated.
    const placeholders = integerTokens.filter(t => t.t === 'digit') as Extract<NumberToken, { t: 'digit' }>[];
    const firstZero = placeholders.findIndex(p => p.zero);
    const minDigits = firstZero < 0 ? 0 : placeholders.length - firstZero;
    const digits = integer.padStart(minDigits, '0');
    let remaining = digits.length;
    let written = 0;
    const integerOut: string[] = [];
    let placeholdersLeft = placeholders.length;
    for (let i = integerTokens.length - 1; i >= 0; i--) {
        const token = integerTokens[i];
        if (token.t === 'text') { integerOut.push(token.text); continue; }
        if (token.t !== 'digit') continue;
        placeholdersLeft--;
        const take = placeholdersLeft === 0 ? remaining : Math.min(1, remaining);
        for (let k = 0; k < take; k++) {
            if (section.grouping && written > 0 && written % 3 === 0) integerOut.push(settings.group);
            integerOut.push(digits[remaining - 1]);
            remaining--;
            written++;
        }
    }
    const integerText = integerOut.reverse().join('');

    // The fraction part: digits in order; a `#` drops a trailing zero.
    const fractionPlaceholders = fractionTokens.filter(t => t.t === 'digit') as Extract<NumberToken, { t: 'digit' }>[];
    let shown = fraction.length;
    while (shown > 0 && !fractionPlaceholders[shown - 1].zero && fraction[shown - 1] === '0') shown--;
    let digitIndex = 0;
    let fractionText = '';
    for (const token of fractionTokens) {
        if (token.t === 'text') fractionText += token.text;
        else if (token.t === 'digit') { if (digitIndex < shown) fractionText += fraction[digitIndex]; digitIndex++; }
    }
    const point = pointAt >= 0 && (shown > 0 || fractionTokens.some(t => t.t === 'text')) ? settings.decimal : '';
    // No sign on a number that rounds to nothing shown.
    const showsDigit = /[1-9]/.test(integer) || /[1-9]/.test(fraction.slice(0, shown));
    return (sign && showsDigit ? '-' : '') + integerText + point + fractionText;
}

/** Writes a date with a date picture, in UTC moved by `utcOffsetMinutes`. */
export function formatDate(date: Date, picture: Extract<Picture, { kind: 'date' }>, settings: PictureSettings): string | null {
    const shifted = new Date(date.getTime() + settings.utcOffsetMinutes * 60_000);
    // A date at the edge of what a Date holds, moved past it by the offset, is no date to write.
    if (Number.isNaN(shifted.getTime())) return null;
    const year = shifted.getUTCFullYear(), month = shifted.getUTCMonth(), day = shifted.getUTCDate();
    const weekday = shifted.getUTCDay(), hour = shifted.getUTCHours(), minute = shifted.getUTCMinutes(), second = shifted.getUTCSeconds();
    const two = (n: number): string => String(n).padStart(2, '0');
    const hour12 = hour % 12 === 0 ? 12 : hour % 12;
    let out = '';
    for (const token of picture.tokens) {
        if (token.t === 'text') { out += token.text; continue; }
        switch (token.part) {
            case 'd': out += day; break;
            case 'dd': out += two(day); break;
            case 'ddd': out += settings.daysShort[weekday]; break;
            case 'dddd': out += settings.days[weekday]; break;
            case 'M': out += month + 1; break;
            case 'MM': out += two(month + 1); break;
            case 'MMM': out += settings.monthsShort[month]; break;
            case 'MMMM': out += settings.months[month]; break;
            case 'yy': out += two(Math.abs(year) % 100); break;
            // The year's four digits, its sign before them (a Date can hold a year before 1).
            case 'yyyy': out += (year < 0 ? '-' : '') + String(Math.abs(year)).padStart(4, '0'); break;
            case 'h': out += hour12; break;
            case 'hh': out += two(hour12); break;
            case 'H': out += hour; break;
            case 'HH': out += two(hour); break;
            case 'm': out += minute; break;
            case 'mm': out += two(minute); break;
            case 's': out += second; break;
            case 'ss': out += two(second); break;
            case 'AM/PM': out += hour < 12 ? settings.am : settings.pm; break;
            case 'am/pm': out += (hour < 12 ? settings.am : settings.pm).toLowerCase(); break;
            case 'A/P': out += (hour < 12 ? settings.am : settings.pm).charAt(0); break;
            case 'a/p': out += (hour < 12 ? settings.am : settings.pm).charAt(0).toLowerCase(); break;
        }
    }
    return out;
}

/** The largest UTC offset ISO 8601 text may carry, in hours. */
const MAX_OFFSET_HOURS = 18;

/** Text that is a number, as form data and CSV carry numbers: `1250`, `-3.5`, `+1e3`, `.5`. */
// `\d+(?:\.\d*)?`, not `\d+\.?\d*`: with the point optional between two digit runs, a long run of
// digits that is no number was split every way before failing, in time quadratic in its length.
const NUMBER_TEXT = /^\s*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?\s*$/;

/** The number a value is for a number picture: a number, or text that is one; null when it is neither. */
export function numberOf(value: unknown): number | null {
    // A value past a double's range (1e400, a bigint of 400 digits) is no number a picture can write.
    const n = typeof value === 'number' ? value
        : typeof value === 'bigint' ? Number(value)
            : typeof value === 'string' && NUMBER_TEXT.test(value) ? Number(value) : NaN;
    return Number.isFinite(n) ? n : null;
}

/** ISO 8601 text: a date, or a date and time with an optional offset (read as UTC without one). */
const ISO_TEXT = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,9}))?)?(Z|[+-]\d{2}:?\d{2})?)?$/;

/**
 * The date a value is for a date picture: a valid Date, or ISO 8601 text, read by this library (the
 * runtime reads a date and time without an offset as local time, which differs between machines; here
 * it is UTC). Null when it is neither.
 */
export function dateOf(value: unknown): Date | null {
    if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
    if (typeof value !== 'string') return null;
    const m = ISO_TEXT.exec(value.trim());
    if (!m) return null;
    const [, y, mo, d, h = '0', mi = '0', s = '0', frac = '', zone] = m;
    const month = Number(mo), day = Number(d), hour = Number(h), minute = Number(mi), second = Number(s);
    if (month < 1 || month > 12 || day < 1 || hour > 23 || minute > 59 || second > 59) return null;
    // setUTCFullYear, not Date.UTC, which reads years 0 to 99 as 1900 to 1999.
    const date = new Date(0);
    date.setUTCFullYear(Number(y), month - 1, day);
    date.setUTCHours(hour, minute, second, Number((frac + '000').slice(0, 3)));
    if (date.getUTCDate() !== day) return null; // 2026-02-30
    let time = date.getTime();
    if (zone && zone !== 'Z') {
        const sign = zone[0] === '-' ? -1 : 1;
        const digits = zone.slice(1).replace(':', '');
        const offsetHours = Number(digits.slice(0, 2)), offsetMinutes = Number(digits.slice(2));
        if (offsetHours > MAX_OFFSET_HOURS || offsetMinutes > 59) return null;
        time -= sign * (offsetHours * 60 + offsetMinutes) * 60_000;
    }
    return new Date(time);
}

/**
 * The text a spreadsheet shows for a number through its cell's number format (`[$€-407]#,##0.00` shows
 * 120 as `€120.00`, `0%` shows 0.19 as `19%`, `dd.mm.yyyy` a day count as `01.10.2026`), for the outputs
 * that print a sheet as it is seen (see CellMetadata.shownText). Written with the template engine's
 * picture formatting, which reads Excel's number formats as Word's pictures are read, after what only
 * Excel's codes hold is turned into those: the section for the value's sign, a currency's `[$€-407]` as
 * its symbol, colours and locales dropped, `_x` padding as a space, `*x` fill as nothing, thousands
 * scaling (`0,`), and a date's `m` read as a month or, next to an hour or a second, a minute. A code it
 * cannot write the way Excel does (General, scientific notation, fractions, conditions, elapsed time,
 * text sections) shows nothing of its own: the value is shown as read.
 */
import { DEFAULT_TEMPLATE_CONFIG } from '../defaults.js';
import { formatDate, formatNumber, PictureSettings, readPicture } from '../template/templateFormat.js';

/** Excel's built-in number formats that are not dates (ECMA-376 Part 1, 18.8.30), by id. Its built-in
 * date formats show the machine's own date order, so their dates keep the ISO text they are read as. */
export const BUILTIN_NUMBER_FORMATS: ReadonlyMap<string, string> = new Map([
    ['1', '0'], ['2', '0.00'], ['3', '#,##0'], ['4', '#,##0.00'], ['9', '0%'], ['10', '0.00%'],
    ['37', '#,##0 ;(#,##0)'], ['38', '#,##0 ;[Red](#,##0)'], ['39', '#,##0.00;(#,##0.00)'], ['40', '#,##0.00;[Red](#,##0.00)'],
]);

/** Excel's own separators and English names, as its codes are written. */
const SETTINGS: PictureSettings = {
    decimal: '.', group: ',', ...DEFAULT_TEMPLATE_CONFIG.dateNames, utcOffsetMinutes: 0,
};

/** Pictures read, by their text: a sheet's cells share a few formats, read once each, not once a cell. */
const READ_PICTURES = new Map<string, ReturnType<typeof readPicture>>();
const MAX_READ_PICTURES = 1024;
function readPictureOnce(picture: string): ReturnType<typeof readPicture> {
    let reading = READ_PICTURES.get(picture);
    if (!reading) {
        reading = readPicture(picture);
        if (READ_PICTURES.size >= MAX_READ_PICTURES) READ_PICTURES.clear();
        READ_PICTURES.set(picture, reading);
    }
    return reading;
}

/** At most this long a code is read: longer, it is no format a spreadsheet writes. */
const MAX_CODE_LENGTH = 512;

/** 9999-12-31 as a day count in the 1900 and 1904 date systems. */
const LAST_DAY_1900 = 2958465;
const LAST_DAY_1904 = 2957003;

/** A code's sections, split at `;` outside quoted text, `[...]` and escaped characters. */
function sections(code: string): string[] {
    const out: string[] = [];
    let start = 0;
    for (let i = 0; i < code.length; i++) {
        const c = code[i];
        if (c === '\\' || c === '_' || c === '*') i++;
        else if (c === '"') { const end = code.indexOf('"', i + 1); if (end < 0) break; i = end; }
        else if (c === '[') { const end = code.indexOf(']', i + 1); if (end < 0) break; i = end; }
        else if (c === ';') { out.push(code.slice(start, i)); start = i + 1; }
    }
    out.push(code.slice(start));
    return out;
}

/**
 * A section as a picture: `[$€-407]` as `"€"`, colours and locales dropped, `_x` as a space and `*x` as
 * nothing. Undefined for one with a condition (`[>=100]`), elapsed time (`[h]`) or a text placeholder.
 */
function asPicture(section: string): string | undefined {
    let out = '';
    for (let i = 0; i < section.length; i++) {
        const c = section[i];
        if (c === '\\') { out += section.slice(i, i + 2); i++; continue; }
        if (c === '_') { out += ' '; i++; continue; }
        if (c === '*') { i++; continue; }
        if (c === '@') return undefined;
        if (c === '"') { const end = section.indexOf('"', i + 1); if (end < 0) return undefined; out += section.slice(i, end + 1); i = end; continue; }
        if (c === '[') {
            const end = section.indexOf(']', i + 1);
            if (end < 0) return undefined;
            const inner = section.slice(i + 1, end);
            if (inner.startsWith('$')) {
                // A currency and locale: the symbol before the `-`, written as literal text.
                const symbol = inner.slice(1).split('-')[0];
                if (symbol) out += `"${symbol.replace(/"/g, '')}"`;
            } else if (/^[<>=]/.test(inner) || /^[hms]+$/i.test(inner)) {
                return undefined;
            }
            // Anything else in brackets is a colour (`[Red]`, `[Color 10]`): it is not text.
            i = end;
            continue;
        }
        out += c;
    }
    return out;
}

/** Whether a picture holds a date or time part (outside quoted text and escapes), and so is a date's. */
function isDatePicture(picture: string): boolean {
    for (let i = 0; i < picture.length; i++) {
        const c = picture[i];
        if (c === '\\') { i++; continue; }
        if (c === '"') { const end = picture.indexOf('"', i + 1); if (end < 0) return false; i = end; continue; }
        if ('yYdDhHsS'.includes(c) || ((c === 'm' || c === 'M') && !/[0#?]/.test(picture))) return true;
    }
    return false;
}

/**
 * An Excel date picture as a Word one: `m` a month (`M`) unless it follows an hour or comes before a
 * second, `h` a 24-hour clock unless the picture has AM/PM. Undefined for fractions of a second.
 */
function wordDatePicture(picture: string): string | undefined {
    const twelve = /AM\/PM|A\/P/i.test(picture);
    const tokens: { kind: 'text' | 'part'; text: string }[] = [];
    for (let i = 0; i < picture.length;) {
        const c = picture[i];
        if (c === '\\') { tokens.push({ kind: 'text', text: picture.slice(i, i + 2) }); i += 2; continue; }
        if (c === '"') { const end = picture.indexOf('"', i + 1); tokens.push({ kind: 'text', text: picture.slice(i, end + 1) }); i = end + 1; continue; }
        const marker = ['AM/PM', 'am/pm', 'A/P', 'a/p'].find(m => picture.startsWith(m, i));
        if (marker) { tokens.push({ kind: 'part', text: marker }); i += marker.length; continue; }
        const lower = c.toLowerCase();
        if ('ymdhs'.includes(lower)) {
            let run = 1;
            while (i + run < picture.length && picture[i + run].toLowerCase() === lower) run++;
            tokens.push({ kind: 'part', text: lower.repeat(run) });
            i += run;
            continue;
        }
        if (c === '0' && tokens.length && /^s+$/.test(tokens[tokens.length - 1].text)) return undefined;
        tokens.push({ kind: 'text', text: c });
        i++;
    }
    const parts = tokens.filter(t => t.kind === 'part');
    return tokens.map(t => {
        if (t.kind === 'text') return t.text;
        if (t.text[0] === 'm') {
            const at = parts.indexOf(t);
            const minute = parts[at - 1]?.text[0] === 'h' || parts[at + 1]?.text[0] === 's';
            return minute ? t.text : 'M'.repeat(t.text.length);
        }
        if (t.text[0] === 'h') return twelve ? t.text : 'H'.repeat(t.text.length);
        return t.text;
    }).join('');
}

/**
 * The text a cell shows for `value` (its stored number) through the number format `code`, or undefined
 * when the code is one written another way than Excel writes it, or the value is no number.
 */
export function excelShownText(value: string, code: string, date1904: boolean): string | undefined {
    if (!code || code.length > MAX_CODE_LENGTH || /^general$/i.test(code.trim())) return undefined;
    let number = Number(value);
    if (value.trim() === '' || !Number.isFinite(number)) return undefined;
    const parts = sections(code);
    let section = parts[0];
    if (number < 0 && parts.length >= 2) { section = parts[1]; number = -number; }
    else if (number === 0 && parts.length >= 3) section = parts[2];
    const picture = asPicture(section);
    if (picture === undefined || picture.trim() === '') return undefined;
    if (isDatePicture(picture)) {
        const word = wordDatePicture(picture);
        if (!word || number < 0) return undefined;
        const seconds = Math.round(number * 86400);
        const days = Math.floor(seconds / 86400);
        if (days > (date1904 ? LAST_DAY_1904 : LAST_DAY_1900) || (!date1904 && days === 60)) return undefined;
        const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, days < 60 ? 31 : 30);
        const reading = readPictureOnce(word);
        if (!reading.ok || reading.picture.kind !== 'date') return undefined;
        return formatDate(new Date(epoch + seconds * 1000), reading.picture, SETTINGS) ?? undefined;
    }
    const bare = picture.replace(/"[^"]*"|\\./g, '');
    // A section of literal text alone (`"-"` for zero) shows that text.
    if (!/[0#?]/.test(bare)) return picture.replace(/"([^"]*)"|\\(.)/g, (_, quoted: string | undefined, escaped: string | undefined) => quoted ?? escaped ?? '').trim() || undefined;
    if (/[?/]|[eE][+-]/.test(bare)) return undefined;
    // Thousands scaling: each comma right after the last digit placeholder divides by 1,000.
    const scaled = picture.replace(/([0#])(,+)(?=[^0#]*$|\.)/, (_, digit: string, commas: string) => { number /= 1000 ** commas.length; return digit; });
    const reading = readPictureOnce(scaled);
    if (!reading.ok || reading.picture.kind !== 'number') return undefined;
    // Padding (`_)`, `_€`) at the edges lines numbers up in a column; the text shown is without it.
    return formatNumber(number, reading.picture, SETTINGS)?.trim() ?? undefined;
}

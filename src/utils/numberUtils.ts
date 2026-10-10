/**
 * Small shared numeric helpers used across the PDF layout/geometry code, the OCR reconstructor, and
 * the text generator, so the same statistic is computed one way everywhere.
 *
 * @module numberUtils
 */

/**
 * Upper median of a numeric list: the value at `floor(n/2)` of the sorted list (0 for empty). Not
 * the mean of the two middle values - the geometry heuristics want an actual observed value (a real
 * line pitch, column width or character width), and the extra branch buys nothing here.
 */
export function median(values: number[]): number {
    if (!values.length) return 0;
    const s = [...values].sort((a, b) => a - b);
    return s[Math.floor(s.length / 2)];
}

/**
 * Clamps a document-derived repeat count to a safe range before it reaches `String.prototype.repeat`,
 * so a hostile file cannot turn a tiny attribute (an ODF `text:c`, a list `ilvl`/indentation) into a
 * multi-gigabyte string, and a negative/NaN count cannot throw. Returns 0 for NaN/negative and caps at
 * `max`. No real document repeats a character thousands of times or nests thousands deep.
 */
export function clampRepeat(count: number, max = 10000): number {
    return Number.isFinite(count) && count > 0 ? Math.min(Math.floor(count), max) : 0;
}

/**
 * A whole number from a document or AST value, within `[min, max]`: `fallback` for anything that is
 * not a finite number (a string, an array), so a value written into markup (a heading level, a list
 * number, an RTF control word's parameter) can only ever be digits.
 */
export function clampInt(value: unknown, min: number, max: number, fallback: number): number {
    return typeof value === 'number' && Number.isFinite(value) ? Math.min(Math.max(Math.floor(value), min), max) : fallback;
}

/**
 * The most columns and rows one table cell spans, as browsers bound HTML's `colspan` and `rowspan`:
 * every reader holds a cell's spans to these, and to at least 1. A document stating a span of
 * billions (or a negative one) made writers fill, or walk back over, a grid no memory holds.
 */
export const MAX_COL_SPAN = 1000;
export const MAX_ROW_SPAN = 65534;

/** A span as a document states it (`"3"`), held to 1 through `max`; 1 when it states none a number reads from. */
export function cellSpan(value: string | null | undefined, max: number): number {
    return clampInt(parseInt(value ?? '', 10), 1, max, 1);
}

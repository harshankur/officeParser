/**
 * How the PDF parser joins a paragraph's lines, and how a writer that keeps those lines tells where
 * it joined a word divided by a hyphen (see `pdfParserConfig.mergeHyphenatedWords`).
 */
import { NodeBounds, OfficeContentNode } from '../types.js';

/** Chinese and Japanese characters (Han, kana, their punctuation and full-width forms): text set without spaces between words. */
export const CJK_CHAR = /[\u3000-\u303F\u3040-\u30FF\u31F0-\u31FF\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF]/;

/**
 * Whether a line read as `line` ends in a hyphen dividing a word that goes on on the next line, read as
 * `next`: a soft hyphen after a letter, or a hyphen after a letter before a lower-case letter. (A dash after
 * a space, or a second hyphen, divides no word: "four -" and "nothing--" keep theirs.)
 */
export function endsInWordHyphen(line: string, next: string): boolean {
    return /\p{L}\u00AD$/u.test(line) || (/\p{L}-$/u.test(line) && /^\p{Ll}/u.test(next));
}

/**
 * Whether the PDF parser joined the line read as `before` to the one read as `after` at a word hyphen
 * (endsInWordHyphen) it took off: the line ends in a letter and the next starts with one, with nothing
 * between. Every other join leaves a space, except Chinese and Japanese text, and a stack of a glyph or
 * two a line (a day "1" over "1"), which it joins without one.
 */
export function joinedAtWordHyphen(before: string, after: string): boolean {
    if (!/\p{L}$/u.test(before) || !/^\p{L}/u.test(after)) return false;
    if (CJK_CHAR.test(before.slice(-1)) || (CJK_CHAR.test(before) && !/\s/.test(before.trim()))) return false;
    return !(before.trim().length <= 2 && after.trim().length <= 2);
}

/**
 * Of a paragraph's runs in reading order, each with its box on the page, the ones that end a line where
 * the PDF parser joined a word divided across the break and took its hyphen off (joinedAtWordHyphen):
 * a writer that keeps the page's lines writes the hyphen after each again. A line starts at a run set
 * below the run before it and back to its left (one raised or lowered after it is on the same line).
 * Only text runs count, and runs without a box tell no lines: none is given.
 */
export function droppedHyphenRuns(runs: readonly OfficeContentNode[]): Set<number> {
    const found = new Set<number>();
    const lines: { text: string; last: number }[] = [];
    let box: NodeBounds | undefined;
    for (let i = 0; i < runs.length; i++) {
        const { type, text, bounds: b } = runs[i];
        if (type !== 'text' || typeof text !== 'string' || !text) continue;
        if (!b || ![b.x, b.y, b.width, b.height].every(Number.isFinite)) return found;
        if (!box || (b.y > box.y + box.height * 0.6 && b.x < box.x + box.width)) lines.push({ text: '', last: i });
        const line = lines[lines.length - 1];
        line.text += text;
        line.last = i;
        box = b;
    }
    for (let k = 1; k < lines.length; k++) if (joinedAtWordHyphen(lines[k - 1].text, lines[k].text)) found.add(lines[k - 1].last);
    return found;
}

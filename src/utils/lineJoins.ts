/**
 * How the PDF parser joins a paragraph's lines, and which runs it took a word's hyphen off where it
 * joined the word across a line (see `pdfParserConfig.mergeHyphenatedWords`).
 */
import { OfficeContentNode } from '../types.js';

/** Chinese and Japanese characters (Han, kana, their punctuation and full-width forms): text set without spaces between words. */
export const CJK_CHAR = /[\u3000-\u303F\u3040-\u30FF\u31F0-\u31FF\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF]/;

/**
 * Whether a line read as `line` ends in a hyphen dividing a word that goes on on the next line, read as
 * `next`: a soft hyphen after a letter, or a hyphen after a letter before a lower-case letter. A letter
 * keeps the marks set on it, an accent written as a character of its own ("e" and U+0301) among them.
 * (A dash after a space, or a second hyphen, divides no word: "four -" and "nothing--" keep theirs.)
 */
export function endsInWordHyphen(line: string, next: string): boolean {
    // (Only the line's end is read: a letter under more marks than that ends no word.)
    const end = line.slice(-32);
    return /\p{L}\p{M}*\u00AD$/u.test(end) || (/\p{L}\p{M}*-$/u.test(end) && /^\p{Ll}/u.test(next));
}

/**
 * The text runs the PDF parser took a word's hyphen off, where it joined the word across a line (a hyphen
 * that was a run of its own leaves that run empty). A writer that keeps the page's lines writes the hyphen
 * after each again. Held beside the AST, not in it: a copy of a run (a JSON round trip) is no run marked,
 * and is written as it reads.
 */
const DROPPED_HYPHEN = new WeakSet<OfficeContentNode>();

/** Marks `run` as one the PDF parser took a word's hyphen off (see DROPPED_HYPHEN). */
export function markDroppedHyphen(run: OfficeContentNode): void {
    DROPPED_HYPHEN.add(run);
}

/** Whether the PDF parser took a word's hyphen off `run` where it joined the word across a line (see DROPPED_HYPHEN). */
export function droppedHyphen(run: OfficeContentNode): boolean {
    return DROPPED_HYPHEN.has(run);
}

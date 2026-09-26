/**
 * Trimming by a set of characters, in one scan from the end being trimmed.
 *
 * A pattern such as `/\n+$/` or `/[ \t]+$/` tries every run of those characters inside the text, fails
 * where the run is not at the end, and retries from the next character, so a long run in the middle of
 * a document's text costs time quadratic in its length. These cost time linear in what they trim.
 * (For whitespace in general, `String.prototype.trimEnd`/`trimStart` are linear already.)
 */

/** `text` without the characters of `chars` at its end. */
export function trimEndChars(text: string, chars: string): string {
    let end = text.length;
    while (end > 0 && chars.includes(text[end - 1])) end--;
    return end === text.length ? text : text.slice(0, end);
}

/** `text` without the characters of `chars` at its start. */
export function trimStartChars(text: string, chars: string): string {
    let start = 0;
    while (start < text.length && chars.includes(text[start])) start++;
    return start === 0 ? text : text.slice(start);
}

/**
 * The whitespace Markdown strips from the edges of a block, heading, list item or cell, and reads
 * between a marker and its content: spaces, tabs and line endings. A no-break space (U+00A0) and the
 * other Unicode spaces are content there, though `String.prototype.trim` and `\s` take them too.
 */
export const ASCII_WHITESPACE = ' \t\n\r\f\v';

/** `text` without {@link ASCII_WHITESPACE} at either end. */
export function trimAsciiWhitespace(text: string): string {
    return trimEndChars(trimStartChars(text, ASCII_WHITESPACE), ASCII_WHITESPACE);
}

/** `text` without whole repetitions of `unit` (such as `\r\n`) at either end. */
export function trimRepeated(text: string, unit: string): string {
    if (!unit) return text;
    let start = 0;
    while (text.startsWith(unit, start)) start += unit.length;
    let end = text.length;
    while (end - unit.length >= start && text.endsWith(unit, end)) end -= unit.length;
    return start === 0 && end === text.length ? text : text.slice(start, end);
}

/** How many of its last characters a {@link TextBuilder} keeps at hand. */
const TEXT_BUILDER_TAIL = 4;

/**
 * Text built piece by piece that knows how it ends without reading itself back. Reading the end of
 * a string grown with `+=` (`endsWith`, `slice(-1)`, an index) makes V8 copy the whole string into
 * one piece first, and again after every append: a writer checking how its output ends after each
 * block would take time quadratic in the length of the document.
 */
export class TextBuilder {
    private readonly parts: string[] = [];
    private tail = '';

    /** Appends `text`. */
    append(text: string): void {
        if (!text) return;
        this.parts.push(text);
        this.tail = (text.length >= TEXT_BUILDER_TAIL ? text : this.tail + text).slice(-TEXT_BUILDER_TAIL);
    }

    /** Whether the text ends with `suffix`, of at most four characters. */
    endsWith(suffix: string): boolean {
        return this.tail.endsWith(suffix);
    }

    /** The last character, or '' while the text is empty. */
    lastChar(): string {
        return this.tail.slice(-1);
    }

    /** Whether nothing has been appended. */
    isEmpty(): boolean {
        return this.parts.length === 0;
    }

    toString(): string {
        return this.parts.join('');
    }
}

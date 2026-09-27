import { AdmonitionMetadata, BreakMetadata, CodeMetadata, CommentMetadata, EmbedMetadata, FullOfficeParserConfig, HeadingMetadata, ImageMetadata, ListMetadata, OfficeAttachment, OfficeContentNode, OfficeMetadata, OfficeParserAST, TextFormatting, TextMetadata } from '../types.js';
import { anchorMark, resolveAnchorMarks } from '../utils/anchorUtils.js';
import { createAST } from '../utils/astUtils.js';
import { parseHtml } from './HtmlParser.js';
import { isSourceComment } from '../utils/commentUtils.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { decodeCharacterReference, decodeCharacterReferences } from '../utils/htmlEntities.js';
import { iframeAllowed } from '../utils/sanitize.js';
import { ASCII_WHITESPACE, trimAsciiWhitespace, trimEndChars, trimStartChars } from '../utils/textUtils.js';
import { lookupTable, setOwn } from '../utils/lookupUtils.js';

// Sentinel node type for a standalone bookmark-anchor block (e.g. `<a id="x"></a>` on its
// own line). A post-parse pass folds these into the following node's anchorIds so they
// round-trip as real anchors rather than being escaped to visible text on regeneration.
const ANCHOR_PLACEHOLDER = '__anchorPlaceholder__';

/**
 * Splits the inner content of a YAML flow array (`a, "b, c", d`) on top-level commas,
 * ignoring commas inside single- or double-quoted items.
 */
const splitFlowArrayItems = (inner: string): string[] => {
    const items: string[] = [];
    let current = '';
    let quote: '"' | '\'' | null = null;
    for (const ch of inner) {
        if (quote) {
            current += ch;
            if (ch === quote) quote = null;
        } else if (ch === '"' || ch === '\'') {
            quote = ch;
            current += ch;
        } else if (ch === ',') {
            items.push(current.trim());
            current = '';
        } else {
            current += ch;
        }
    }
    if (current.trim() !== '') items.push(current.trim());
    return items;
};

/**
 * Maps every accepted-on-import admonition type spelling (GitHub's five plus GLFM's
 * `danger`) to the canonical AdmonitionMetadata type. Per MARKDOWN_DIALECT.md's
 * Decisions, `danger` folds into `caution` - there is no separate danger type.
 */
const ADMONITION_TYPE_MAP: Record<string, AdmonitionMetadata['admonitionType']> = lookupTable({
    note: 'note',
    tip: 'tip',
    important: 'important',
    warning: 'warning',
    caution: 'caution',
    danger: 'caution'
});

/** Where the next `-->` is, remembered between calls so that scanning many `<!--` stays linear. */
interface CommentCloseCache { at: number; from: number }

/**
 * The comment opening at `open` (where `text` has `<!--`): its raw body and the index just past it, or
 * null when it never closes. `<!-->` and `<!--->` are complete, empty comments, as in HTML and
 * CommonMark; otherwise the comment ends at the first `-->`. `closes` must be shared by calls made in
 * increasing `open` order over the same text: each `-->` search resumes from the last one, so a text
 * full of unclosed `<!--` costs time linear in its length rather than quadratic.
 */
function commentAt(text: string, open: number, closes: CommentCloseCache): { body: string; end: number } | null {
    if (text.startsWith('<!-->', open)) return { body: '', end: open + 5 };
    if (text.startsWith('<!--->', open)) return { body: '', end: open + 6 };
    const from = open + 4;
    if (!(from >= closes.from && (closes.at === -1 || closes.at >= from))) {
        closes.at = text.indexOf('-->', from);
        closes.from = from;
    }
    return closes.at === -1 ? null : { body: text.slice(from, closes.at), end: closes.at + 3 };
}

/**
 * A function returning the first `\n` at or after a position of `text` (-1 when there is none), for
 * positions asked in increasing order: it resumes from its last answer, so any number of questions
 * costs one scan of the text.
 */
function newlineCursor(text: string): (from: number) => number {
    let at = -2;
    return from => {
        if (at === -2 || (at !== -1 && at < from)) at = text.indexOf('\n', from);
        return at;
    };
}

/**
 * Paragraph lines, with the lines a code span runs across joined into one, each line end in it read
 * as a space, as CommonMark reads a line end inside a code span, so a span an editor wrapped
 * (`` x `a `` then `` b` y ``) is code rather than two backticks of text. A backtick run not escaped
 * with a backslash opens a span that the next run of as many backticks closes, on its line or a
 * later one (codeSpanCloser, so linear); a line joined earlier (for a comment or display math) stays
 * whole.
 */
function joinCodeSpanLines(lines: string[]): string[] {
    if (lines.length < 2 || !lines.some(line => line.includes('`'))) return lines;
    const text = lines.join('\n');
    // Where each line after the first starts, less one: the newline joining it to the previous.
    const breaks: number[] = [];
    for (let i = 0, at = -1; i < lines.length - 1; i++) breaks.push(at += lines[i].length + 1);
    const close = codeSpanCloser(text);
    const joined = new Set<number>();
    let b = 0;
    for (let i = 0; i < text.length; i++) {
        if (text[i] === '\\' && i + 1 < text.length && text[i + 1] !== '\n') { i++; continue; }
        if (text[i] !== '`') continue;
        let end = i + 1;
        while (text[end] === '`') end++;
        const closeAt = close(end, end - i);
        if (closeAt === -1) { i = end - 1; continue; }
        while (b < breaks.length && breaks[b] < i) b++;
        while (b < breaks.length && breaks[b] < closeAt) joined.add(b++);
        i = closeAt + (end - i) - 1;
    }
    if (joined.size === 0) return lines;
    const out: string[] = [];
    let current = lines[0];
    for (let k = 1; k < lines.length; k++) {
        if (joined.has(k - 1)) current += ` ${lines[k]}`;
        else { out.push(current); current = lines[k]; }
    }
    out.push(current);
    return out;
}

/**
 * Paragraph lines, with the lines a comment spans joined back into one (with their line breaks), so a
 * comment that opens on one line and closes on a later one is parsed as one comment rather than as
 * visible text. An opener inside a code span on its line (after an odd number of backticks) is not a
 * comment. Linear: each line is scanned once, and `-->` and newline searches each resume from their
 * last answer.
 */
function joinCommentLines(lines: string[]): string[] {
    if (!lines.some(l => l.includes('<!--'))) return lines;
    const text = lines.join('\n');
    const closes: CommentCloseCache = { at: -1, from: Number.MAX_SAFE_INTEGER };
    const newlineFrom = newlineCursor(text);
    const joins = new Set<number>(); // newline offsets inside a comment
    let pos = 0; // everything before pos has been accounted for
    let ticks = 0; // backticks between the start of pos's line and pos
    for (;;) {
        const open = text.indexOf('<!--', pos);
        if (open === -1) break;
        for (let k = pos; k < open; k++) {
            if (text[k] === '\n') ticks = 0;
            else if (text[k] === '`') ticks++;
        }
        pos = open + 4;
        if (ticks % 2 === 1) continue;
        const found = commentAt(text, open, closes);
        if (!found) continue;
        let spansLines = false;
        for (let k = newlineFrom(open); k !== -1 && k < found.end; k = newlineFrom(k + 1)) {
            joins.add(k);
            spansLines = true;
        }
        if (spansLines) ticks = 0;
        pos = found.end;
    }
    if (joins.size === 0) return lines;
    const out: string[] = [];
    let offset = 0;
    for (const line of lines) {
        if (out.length && joins.has(offset - 1)) out[out.length - 1] += `\n${line}`;
        else out.push(line);
        offset += line.length + 1;
    }
    return out;
}

/**
 * Text that is not parsed as inline Markdown but is still Markdown text (image alt text, link and
 * image titles, abbreviation definitions), read as CommonMark reads it: backslash-escaped
 * punctuation and character references decoded, in one pass, so an escaped `\&` stays a literal
 * `&` rather than starting a reference.
 */
/** A YAML double-quoted scalar's value: decoded as JSON when it is valid JSON, else its quotes removed. */
const decodeDoubleQuoted = (quoted: string): string => {
    try {
        const value = JSON.parse(quoted);
        if (typeof value === 'string') return value;
    } catch { /* not JSON: keep the text between the quotes */ }
    return quoted.slice(1, -1);
};

const decodeMarkdownText = (text: string): string =>
    text.replace(/\\([!-/:-@[-`{-~])|&(#\d+|#[xX][0-9a-fA-F]+|[a-zA-Z][a-zA-Z0-9]*);/g,
        (full: string, escaped: string | undefined, ref: string | undefined) => (escaped !== undefined ? escaped : decodeCharacterReference(ref!) ?? full));

/** Whether the character at `i` is escaped: preceded by an odd number of backslashes. */
const isEscapedAt = (text: string, i: number): boolean => {
    let slashes = 0;
    while (i - slashes - 1 >= 0 && text[i - slashes - 1] === '\\') slashes++;
    return slashes % 2 === 1;
};

/**
 * A Markdown link destination `url "title"` (also `'title'` or `(title)`) split into its URL and
 * optional title, both decoded as CommonMark decodes them. A title may escape its own delimiter
 * (`"say \"hi\""`), and holds no unescaped closing delimiter. It is found from the end: the
 * destination's last character closes it, and it opens at the leftmost unescaped opening delimiter
 * that follows whitespace with no unescaped closing one after it (a pattern scanning from the start
 * retried every position of a long run of spaces, quadratic). With no title, the whole is the URL.
 */
function splitUrlTitle(raw: string): { url: string; title?: string } {
    const text = raw.trim();
    const close = text[text.length - 1];
    const open = close === ')' ? '(' : close;
    if (text.length >= 3 && (close === '"' || close === '\'' || close === ')') && !isEscapedAt(text, text.length - 1)) {
        let opener = -1;
        for (let i = text.length - 2; i >= 1; i--) {
            if (text[i] === close && close !== open && !isEscapedAt(text, i)) break; // a `)` the title cannot hold
            if (text[i] === open && !isEscapedAt(text, i)) {
                if (/\s/.test(text[i - 1])) opener = i;
                if (close === open) break; // a quote title opens at the nearest unescaped quote
            }
        }
        const url = opener === -1 ? '' : text.slice(0, opener).trimEnd();
        if (opener !== -1 && url && !url.includes('\n')) {
            return { url: decodeMarkdownText(url), title: decodeMarkdownText(text.slice(opener + 1, -1)) };
        }
    }
    // The target is decoded too, as a renderer decodes it (the generator escapes what it must).
    return { url: decodeMarkdownText(text) };
}

/**
 * Appends `items` to `target` one by one. `target.push(...items)` passes every item as an argument,
 * which throws once there are more than the engine allows (some hundred thousand nodes, from one
 * long line); a document is only as long as its author made it.
 */
function appendAll<T>(target: T[], items: readonly T[]): void {
    for (const item of items) target.push(item);
}

/** ASCII punctuation, the characters a backslash escapes in CommonMark. */
const ASCII_PUNCTUATION = /[!-/:-@[-`{-~]/;

/** A letter, digit or other character that is neither whitespace nor ASCII punctuation (a word character, for emphasis). */
const isWordCharacter = (char: string | undefined): boolean => char !== undefined && !/\s/.test(char) && !ASCII_PUNCTUATION.test(char);

/**
 * For one text, a function giving where the code span opened by a run of `length` backticks ending
 * at `from` closes (the index of the closing run), or -1. As CommonMark reads it, a code span closes
 * at the next run of exactly as many backticks (runs are maximal, and a backslash is literal inside a
 * code span). The text's runs are listed by length once, and each length's cursor only moves
 * forward, as the tokenizer's openers do, so a text full of openers that never close costs time
 * linear in its length.
 */
function codeSpanCloser(text: string): (from: number, length: number) => number {
    let runs: Map<number, number[]> | undefined;
    const cursors = new Map<number, number>();
    return (from, length) => {
        if (!runs) {
            runs = new Map();
            for (let start = text.indexOf('`'); start !== -1;) {
                let end = start + 1;
                while (text[end] === '`') end++;
                const starts = runs.get(end - start);
                if (starts) starts.push(start); else runs.set(end - start, [start]);
                start = text.indexOf('`', end);
            }
        }
        const starts = runs.get(length);
        if (!starts) return -1;
        let k = cursors.get(length) ?? 0;
        while (k < starts.length && starts[k] < from) k++;
        cursors.set(length, k);
        return k < starts.length ? starts[k] : -1;
    };
}

/**
 * For one text, a function giving where the underscore emphasis opened by a run of `length`
 * underscores ending at `from` closes (the index of the closing run), or -1. As CommonMark has it
 * for `_`, the closer is a run of exactly as many underscores later on the same line, not preceded
 * by whitespace and not followed by a word character, so an underscore inside a word (snake_case)
 * never closes emphasis. A line found to hold no closer for a length is remembered, so a line full
 * of openers costs time linear in its length.
 */
function underscoreCloser(text: string): (from: number, length: number) => number {
    const newlineFrom = newlineCursor(text);
    const noCloserBefore = [0, 0, 0, 0];
    return (from, length) => {
        if (from < noCloserBefore[length]) return -1;
        const newline = newlineFrom(from);
        const lineEnd = newline === -1 ? text.length : newline;
        for (let start = text.indexOf('_', from + 1); start !== -1 && start < lineEnd; start = text.indexOf('_', start + 1)) {
            let end = start + 1;
            while (text[end] === '_') end++;
            if (end - start === length && !/\s/.test(text[start - 1]) && !isWordCharacter(text[end])) return start;
            start = end - 1;
        }
        noCloserBefore[length] = lineEnd;
        return -1;
    };
}

/**
 * For one text, a function giving where the emphasis (`*`), strikethrough (`~~`) or highlight (`==`)
 * opened by a run of `char` ending at `from` closes: the start of the first run of at least `length`
 * of them later on its line, after at least one character, that is not escaped with a backslash (so
 * `**a\***` is bold `a*`), or -1. (An opener is a run not followed by whitespace, so `5 * 3 * 2` holds
 * none. A closer may follow whitespace, which CommonMark does not allow, because this library's
 * earlier versions wrote a run's trailing space inside its delimiters, `**Note: **body`, and files
 * saved that way keep their bold.) As for underscores, a line found to hold no closer for a length is
 * remembered, so a line full of openers that never close costs time linear in it.
 */
function delimiterCloser(text: string, char: string): (from: number, length: number) => number {
    const newlineFrom = newlineCursor(text);
    const noCloserBefore: number[] = [];
    return (from, length) => {
        if (from < (noCloserBefore[length] ?? -1)) return -1;
        const newline = newlineFrom(from);
        const lineEnd = newline === -1 ? text.length : newline;
        for (let i = from; i < lineEnd; i++) {
            if (text[i] === '\\') { i++; continue; }
            if (text[i] !== char) continue;
            let end = i + 1;
            while (end < lineEnd && text[end] === char) end++;
            if (end - i >= length && i > from) return i;
            i = end - 1;
        }
        noCloserBefore[length] = lineEnd;
        return -1;
    };
}

/**
 * The cells of a pipe-table row, as GFM splits it: at each `|` not escaped with a backslash, with an
 * escaped `\|` read as `|` (in code spans too); the pipes at either end bound the row rather than
 * delimit empty cells. Other backslash escapes are left for inline parsing.
 */
function tableRowCells(line: string): string[] {
    const cells: string[] = [];
    let cell = '';
    for (let i = 0; i < line.length; i++) {
        // `\|` is a pipe in the cell whatever stands before it, as GFM reads it (a backslash before it is
        // not an escape of its own here): the generator writes a code span's `\|` as `\\|`, which, read
        // as an escaped backslash and a pipe, split the cell and broke the span.
        if (line[i] === '\\' && line[i + 1] === '|') { cell += '|'; i++; }
        else if (line[i] === '|') { cells.push(cell); cell = ''; }
        else cell += line[i];
    }
    cells.push(cell);
    if (cells.length > 1 && !trimAsciiWhitespace(cells[0]) && trimStartChars(line, ASCII_WHITESPACE).startsWith('|')) cells.shift();
    const end = trimEndChars(line, ASCII_WHITESPACE);
    if (cells.length > 1 && !trimAsciiWhitespace(cells[cells.length - 1]) && end.endsWith('|') && end[end.length - 2] !== '\\') cells.pop();
    return cells;
}

/**
 * Whether a block holds a pipe table's delimiter row: a line after its first made only of `-`, `:`,
 * `|`, spaces and tabs, with a `-` in it (it may end the block, as in a table with no body rows).
 * Tested line by line with one character class, where one pattern holding several overlapping runs
 * backtracked for minutes over a line of spaces.
 */
function hasTableDelimiterRow(block: string): boolean {
    const lines = block.split('\n');
    for (let i = 1; i < lines.length; i++) {
        if (lines[i].includes('-') && /^[-:| \t]+$/.test(lines[i])) return true;
    }
    return false;
}

/**
 * A heading's text and the id written after it (`## Title {#id}`): the id runs from the first ` {#`
 * after any earlier `}` to a `}` ending the heading. Found by index searches, not a backtracking
 * pattern, which took minutes on a heading holding many unclosed `{#`.
 */
function splitHeadingAnchor(rest: string): { text: string; anchor?: string } {
    const text = trimEndChars(rest, ASCII_WHITESPACE);
    if (text.endsWith('}')) {
        const previousClose = text.lastIndexOf('}', text.length - 2);
        for (let open = text.indexOf('{#', previousClose + 1); open !== -1 && open < text.length - 3; open = text.indexOf('{#', open + 1)) {
            // (At the very start, the id is an empty heading's: `## {#x}`.)
            if (open === 0 || /[ \t]/.test(text[open - 1])) return { text: trimEndChars(text.slice(0, open), ASCII_WHITESPACE), anchor: text.slice(open + 2, -1) };
        }
    }
    return { text };
}

/**
 * MDX components removed, their content kept (parse-only: MDX is never written back). A component is
 * a tag whose name starts with a capital, as React and MDX tell it from HTML. `<Component ... />` goes,
 * and `<Component ...>inner</Component>` becomes `inner`, nested ones too. A tag inside a code span on its line (after an odd
 * number of backticks) is code, and one with no matching closing tag is text. One scan finds the tags
 * and pairs each closing tag with the latest open one of its name, so the time is linear in the text.
 */
function stripMdxComponents(text: string): string {
    const token = /(`)|(\n)|<(\/?)([A-Z][A-Za-z0-9]*)(?:\s[^<>]*?)?(\/?)>/g;
    const open = new Map<string, { start: number; end: number }[]>();
    const removed: { start: number; end: number }[] = [];
    let ticks = 0;
    for (let match = token.exec(text); match; match = token.exec(text)) {
        if (match[1]) { ticks++; continue; }
        if (match[2]) { ticks = 0; continue; }
        if (ticks % 2 === 1) continue;
        const [whole, , , closing, name, selfClosing] = match;
        const span = { start: match.index, end: match.index + whole.length };
        if (selfClosing && !closing) removed.push(span);
        else if (!closing) (open.get(name) ?? open.set(name, []).get(name)!).push(span);
        else {
            const opener = open.get(name)?.pop();
            if (opener) removed.push(opener, span);
        }
    }
    if (!removed.length) return text;
    removed.sort((a, b) => a.start - b.start);
    let out = '';
    let pos = 0;
    for (const span of removed) {
        out += text.slice(pos, span.start);
        pos = span.end;
    }
    return out + text.slice(pos);
}

/** The MarkdownGenerator's alignment wrapper in a table cell, `<div style="text-align: X">`. */
const ALIGN_DIV_OPEN = /<div\s+style="text-align:\s*(left|center|right|justify);?"\s*>/gi;

/**
 * A cell's text with each alignment wrapper (`<div style="text-align: X">...</div>`) taken off its
 * content, and the alignment the last one gave. Each wrapper ends at the first `</div>` after it; once
 * none follows, no later wrapper can end either, so the search stops rather than looking for one from
 * every opener to the end of the cell.
 */
function unwrapAlignDivs(text: string): { text: string; align?: 'left' | 'center' | 'right' | 'justify' } {
    let out = '';
    let pos = 0;
    let align: 'left' | 'center' | 'right' | 'justify' | undefined;
    ALIGN_DIV_OPEN.lastIndex = 0;
    for (let open = ALIGN_DIV_OPEN.exec(text); open; open = ALIGN_DIV_OPEN.exec(text)) {
        const close = text.indexOf('</div>', ALIGN_DIV_OPEN.lastIndex);
        if (close === -1) break;
        align = open[1].toLowerCase() as typeof align;
        out += text.slice(pos, open.index) + text.slice(ALIGN_DIV_OPEN.lastIndex, close);
        pos = ALIGN_DIV_OPEN.lastIndex = close + 6;
    }
    return { text: pos === 0 ? text : out + text.slice(pos), align };
}

/** Nodes the generator writes anchors for: waiting anchors go to the next of these. */
const ANCHOR_HOLDERS = new Set<string>(['paragraph', 'heading', 'list', 'image', 'table', 'sheet', 'slide', 'page', 'code', 'break', 'admonition', 'definitionList']);

/**
 * A line starting an HTML block that can hold a table (CommonMark's HTML block start 6: a block-level
 * tag, opening or closing), up to three spaces in: `<table>` itself, or a wrapper such as `<center>`,
 * `<p align="center">`, `<figure>` or `<div>` before it on the line.
 */
const HTML_BLOCK_LINE = /(?:^|\n) {0,3}<\/?(?:address|article|aside|blockquote|body|caption|center|dd|details|dialog|dir|div|dl|dt|fieldset|figcaption|figure|footer|form|h[1-6]|header|html|legend|li|main|menu|nav|ol|p|section|summary|table|tbody|td|tfoot|th|thead|tr|ul)(?=[\s/>]|$)/i;

/** A table's own tags: a piece holding one is HTML, not Markdown in a cell (see splitIntoBlocks). */
const TABLE_PART_TAG = /<\/?(?:table|thead|tbody|tfoot|tr|td|th|caption|colgroup|col)\b/i;

/** How many blank-line-separated pieces an HTML table may run across while Markdown in its cells joins it (see splitIntoBlocks). */
const MAX_TABLE_PIECES = 64;

/** How deeply nested an HTML table that blank lines run through may be, for its pieces to join (see splitIntoBlocks). */
const MAX_JOINED_TABLE_DEPTH = 16;

/**
 * The change in how many HTML tables are open over `text`: its `<table` tags less its `</table>` tags.
 */
function htmlTableDepthChange(text: string): number {
    if (!/<\/?table/i.test(text)) return 0;
    return (text.match(/<table\b/gi)?.length ?? 0) - (text.match(/<\/table\s*>/gi)?.length ?? 0);
}

/** The column a line's text starts at, a tab advancing to the next multiple of four. */
function indentColumn(line: string): number {
    let column = 0;
    for (const ch of line) {
        if (ch === ' ') column++;
        else if (ch === '\t') column += 4 - (column % 4);
        else break;
    }
    return column;
}

/** `line` without up to `column` columns of its leading spaces and tabs. */
function dedentLine(line: string, column: number): string {
    let at = 0;
    let i = 0;
    while (i < line.length && at < column && (line[i] === ' ' || line[i] === '\t')) {
        at = line[i] === '\t' ? at + 4 - (at % 4) : at + 1;
        i++;
    }
    return line.slice(i);
}

/**
 * The columns the content of `block`'s list items starts at (CommonMark's item content column: past
 * the marker and up to four spaces after it, or one space when the item is empty or its content is
 * indented code), added to `columns`.
 */
function addItemContentColumns(block: string, columns: Set<number>): void {
    for (const line of block.split('\n')) {
        if (!LIST_ITEM_START.test(line)) continue;
        const item = /^([ \t]*)([-*+]|\d{1,9}[.)])([ \t]*)/.exec(line)!;
        const indent = indentColumn(item[1]);
        // The space after the marker counted from where the marker ends: a tab there reaches the next
        // multiple of four from that column, not from the line's start.
        const markerEnd = indent + item[2].length;
        const spaces = indentColumn(' '.repeat(markerEnd) + item[3]) - markerEnd;
        const empty = !line.slice(item[0].length).trim();
        columns.add(indent + item[2].length + (empty || spaces > 4 ? 1 : spaces));
    }
}

/** Whether every line of `block` with content is indented four columns: an indented code block. */
function isIndentedCodeBlock(block: string): boolean {
    let content = false;
    for (const line of block.split('\n')) {
        if (!trimAsciiWhitespace(line)) continue;
        if (!/^(?: {4}|\t)/.test(line)) return false;
        content = true;
    }
    return content;
}

/** A thematic break: three or more `-`, `*` or `_`, spaces between allowed, indented at most three spaces. */
const THEMATIC_BREAK = /^ {0,3}([-*_])(?:[ \t]*\1){2,}[ \t]*$/;

/** A setext heading's underline, under paragraph text: `=` or `-` only, indented at most three spaces. */
const SETEXT_UNDERLINE = /^ {0,3}(?:=+|-+)[ \t]*$/;

/** The start of an ATX heading (after any empty anchors): up to three spaces, one to six `#`, then a space or the end. */
const ATX_HEADING_START = /^(?:<a[^>]*><\/a>)* {0,3}#{1,6}(?:[ \t]|$)/;

/**
 * A list item's marker line: the indentation, the marker, and the first character of its content
 * (a no-break space counts: only spaces and tabs are Markdown whitespace), which is missing for an
 * empty item (a marker alone on its line).
 */
const LIST_ITEM_START = /^([ \t]*)([-*+]|\d{1,9}[.)])(?:[ \t]+([^ \t])|[ \t]*$)/;

/** An empty anchor (a bookmark target, `<a id="x"></a>`) where the scan stands, after any whitespace. */
const EMPTY_ANCHOR = /\s*<a\s[^>]*>\s*<\/a>/iy;

/**
 * The run of empty anchors (whitespace between them allowed) at the start of `text`: where it ends,
 * and the ids its anchors name. Read anchor by anchor, so each character is looked at a bounded number
 * of times: a pattern for "anchors and nothing else" tried at every `<a` would scan from each one to
 * the next `>`, the end of the document when none follows.
 */
function emptyAnchorRun(text: string): { end: number; ids: string[] } {
    const ids: string[] = [];
    let end = 0;
    for (;;) {
        EMPTY_ANCHOR.lastIndex = end;
        const anchor = EMPTY_ANCHOR.exec(text);
        if (!anchor) return { end, ids };
        const id = /<a\s[^>]*\b(?:name|id)="([^"]*)"/i.exec(anchor[0])?.[1];
        if (id) ids.push(id);
        end = EMPTY_ANCHOR.lastIndex;
    }
}

/**
 * A code span's content as CommonMark reads it: line endings become spaces, and when it both begins
 * and ends with a space (and is not all spaces) one is stripped from each end, so the padding that
 * lets a span start or end with a backtick (`` `` `x` `` ``) is not part of the code.
 */
function codeSpanContent(raw: string): string {
    const text = raw.replace(/\n/g, ' ');
    return text.length > 2 && text.startsWith(' ') && text.endsWith(' ') && /[^ ]/.test(text) ? text.slice(1, -1) : text;
}

/** Link text: characters, backslash escapes, and bracket pairs nested up to three deep. */
const LINK_TEXT = (() => {
    const unit = String.raw`[^\[\]\\\n]|\\[^\n]`;
    let text = `(?:${unit})*`;
    for (let depth = 0; depth < 3; depth++) text = `(?:${unit}|\\[${text}\\])*`;
    return text;
})();

/** A link target: characters and parenthesis pairs nested up to three deep (a Wikipedia `Foo_(bar)`). */
/** A footnote reference, `[^id]`, anywhere in a text. */
const FOOTNOTE_REFERENCE = /\[\^[^\[\]\n]+\]/;

const LINK_DESTINATION = (() => {
    const unit = String.raw`[^()\n\]]|\](?!\()`;
    let destination = `(?:${unit})*`;
    for (let depth = 0; depth < 3; depth++) destination = `(?:${unit}|\\(${destination}\\))*`;
    return destination;
})();

/**
 * The inline HTML elements Markdown text may hold, as the formatting each gives what it holds: raw
 * HTML a renderer shows as formatting, where it was escaped into visible tags on the first save
 * (`Press <kbd>Ctrl</kbd>` became `Press &lt;kbd>Ctrl&lt;/kbd>`). `code` marks the elements whose
 * content is code (read as written, references decoded); an element with no formatting of its own
 * (`<span>`, `<small>`) is its content. `<a>`, `<abbr>`, `<q>` and `<img>` are read in inlineHtml.
 */
const HTML_INLINE_ELEMENTS: Readonly<Record<string, TextFormatting | 'code'>> = lookupTable({
    b: { bold: true }, strong: { bold: true },
    i: { italic: true }, em: { italic: true }, cite: { italic: true }, var: { italic: true }, dfn: { italic: true },
    s: { strikethrough: true }, strike: { strikethrough: true }, del: { strikethrough: true },
    u: { underline: true }, ins: { underline: true },
    sub: { subscript: true }, sup: { superscript: true },
    mark: { backgroundColor: '#ffff00' },
    code: 'code', kbd: 'code', samp: 'code', tt: 'code',
    span: {}, small: {}, big: {}, font: {}, bdi: {}, bdo: {}, time: {}, data: {}, a: {}, abbr: {}, q: {},
} as Record<string, TextFormatting | 'code'>);

/** An HTML tag's attributes (`name="value"`, `name='value'`, `name=value`, `name`), names lowercased and values decoded. */
function htmlAttributes(source: string): Map<string, string> {
    const out = new Map<string, string>();
    for (const m of source.matchAll(/([^\s"'<>\/=]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/g)) {
        const name = m[1].toLowerCase();
        if (!out.has(name)) out.set(name, decodeCharacterReferences(m[2] ?? m[3] ?? m[4] ?? ''));
    }
    return out;
}

/** The closing tag of each inline element, matched in any case (`</B>`, `</b >`). */
const HTML_CLOSERS = new Map<string, RegExp>();
const htmlCloser = (tag: string): RegExp => {
    let closer = HTML_CLOSERS.get(tag);
    if (!closer) HTML_CLOSERS.set(tag, closer = new RegExp(`</${tag}\\s*>`, 'gi'));
    return closer;
};

/**
 * The inline tokenizer's alternatives, as named groups (so adding one never renumbers the dispatch):
 *
 * - `esc`: a backslash-escaped ASCII punctuation character. Listed first, as only a backslash starts
 *   it; a backslash inside a code span is never offered to it, since the span is consumed whole.
 * - `imgBang`/`imgAlt`/`imgUrl`/`imgAttrs`: an inline image or link, `[text](target "title")`, with an
 *   optional attribute list (`{width=50%}`). The target may hold parenthesis pairs (LINK_DESTINATION)
 *   and stops at its first unpaired `)` (or at a following `](`), except that a quoted title after a
 *   target without spaces may hold `(`, `)` and brackets.
 * - `stars` (`*`, `**`, `***`), `tildes` (`~~`) and `equals` (`==`): the opener of emphasis,
 *   strikethrough or a highlight, a run not followed by whitespace; see delimiterCloser.
 * - `underscores`: the opener of `_italic_`, `__bold__` or `___both___`, a run of one to three
 *   underscores that is not inside a word (CommonMark's rule for `_`); see underscoreCloser.
 * - `codeFence`: a backtick run, the opener of a code span; see codeSpanCloser.
 * - `underline`/`subscript`/`superscript`/`spanStyle`+`spanContent`: HTML-style inline formatting,
 *   each ending before another opening tag of its kind.
 * - `htmlTag`/`htmlAttrs`: any other opening tag (`<kbd>`, `<a href>`, `<img>`); one that is not an
 *   inline element HTML_INLINE_ELEMENTS knows, or that is never closed, is text (see inlineHtml).
 * - `lineBreak` (`<br>`), `anchorTag` (an empty `<a id="x"></a>`, a bookmark target: the id of the
 *   picture right after it, else of the line's paragraph, heading, item or cell; an empty `<a>`
 *   naming no id is text), `htmlComment` (the `<!--` of an inline source comment; see commentAt).
 * - `footnoteId` (`[^id]`), `citationKey` (`[@key]`), `wikiPage`/`wikiAlias` (`[[page|alias]]`),
 *   `refBang`/`refText`/`refId` (`[text][ref]`, `[text][]`), and `shortBang`/`shortText` (`[text]`,
 *   the most generic bracket pattern, so last among those starting with `[`).
 * - `autolinkUrl` (`<https://...>`), `autolinkEmail` (`<me@example.com>`, a mail link), `mathDisplay` (`$$...$$`, before `mathInline`, which would
 *   otherwise match its inner `$...$`), and `mathInline` (`$...$`, with no whitespace just inside
 *   either `$`, the Pandoc/KaTeX heuristic that keeps "$5 and $10" as text).
 *
 * Every alternative stops scanning early, so a paragraph costs time in proportion to its length
 * however it is written: link text holds no stray `[`, a target stops at its first `)`, labels hold
 * no `[`, an HTML-style span stops at another opening tag of its kind, an attribute list is at most
 * 1000 characters, and the closers of code spans and underscore emphasis are looked up rather than
 * scanned for. Nothing else is capped, so a long link or span is still a link or span.
 */
const INLINE_TOKENS = [
    String.raw`\\(?<esc>[!-\/:-@\[-\x60{-~])`,
    String.raw`(?<imgBang>!?)\[(?<imgAlt>${LINK_TEXT})\]\((?<imgUrl>[^\s()\[\]]*\s+(?:"(?:[^"\\\n]|\\[^\n])*"|'(?:[^'\\\n]|\\[^\n])*')\s*|${LINK_DESTINATION})\)(?:\{(?<imgAttrs>[^}\n]{0,1000})\})?`,
    String.raw`(?<stars>\*{1,3})(?![\s*])`,
    String.raw`(?<!(?:^|[^\\])_)(?<![^\s!-\/:-@\[-\x60{-~])(?<underscores>_{1,3})(?![\s_])`,
    String.raw`(?<tildes>~~)(?![\s~])`,
    String.raw`(?<equals>==)(?![\s=])`,
    String.raw`(?<codeFence>\x60+)`,
    String.raw`<u>(?<underline>(?:(?!<u>).)+?)<\/u>`,
    String.raw`<sub>(?<subscript>(?:(?!<sub>).)+?)<\/sub>`,
    String.raw`<sup>(?<superscript>(?:(?!<sup>).)+?)<\/sup>`,
    String.raw`(?<lineBreak><br\s*\/?>)`,
    String.raw`(?<anchorTag><a\s[^<>\n]*>[ \t]*<\/a>)`,
    String.raw`(?<htmlComment><!--)`,
    String.raw`<span\s+style="(?<spanStyle>[^"\n]*)">(?<spanContent>(?:(?!<span[\s>]).)+?)<\/span>`,
    String.raw`<(?<htmlTag>[a-zA-Z][a-zA-Z0-9]*)(?<htmlAttrs>\s[^<>]{0,1000})?>`,
    String.raw`\[\^(?<footnoteId>[^\[\]\n]+)\]`,
    String.raw`\[@(?<citationKey>[a-zA-Z0-9_:.-]+)\]`,
    String.raw`\[\[(?<wikiPage>[^\[\]|\n]+)(?:\|(?<wikiAlias>[^\[\]\n]+))?\]\]`,
    String.raw`(?<refBang>!?)\[(?<refText>[^\[\]\n]*)\]\[(?<refId>[^\[\]\n]*)\]`,
    String.raw`(?<shortBang>!?)\[(?<shortText>[^\[\]\n]+)\]`,
    String.raw`<(?<autolinkUrl>(?:https?|mailto):[^\s<>]+)>`,
    String.raw`<(?<autolinkEmail>[\w.!#$%&'*+\/=?^\x60{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*)>`,
    String.raw`\$\$(?!\$)(?<mathDisplay>(?:\\[\s\S]|[^$\\])+?)\$\$`,
    String.raw`\$(?!\s)(?<mathInline>[^$\n]+?)(?<!\s)\$`,
].join('|');

/** The inline tokenizer, compiled once. Its `lastIndex` is set before every search (see parseInline). */
const INLINE_TOKEN_REGEX = new RegExp(INLINE_TOKENS, 'g');

/**
 * Replaces each fenced code block in `text` with what `replace` returns for it, read as CommonMark
 * reads a fence: three or more backticks or tildes, whose info string's first word is the language
 * (`c++`, `objective-c`, `js title="a.js"`; a backtick fence's info string holds no backtick), closed
 * by a fence of the same character at least as long (so a `~~~` block isn't closed by a stray ```
 * inside it). A fence may be indented up to three spaces past the start of what holds it: the line,
 * or, inside a list item, the column the item's content starts at (so the four-space-indented fence
 * under `- item` is a fence, not indented code); an item's fence closes within the item. Content
 * lines lose up to the opening fence's indentation. A fence that is never closed is left as text.
 */
function liftFencedCode(text: string, replace: (lang: string, code: string) => string): string {
    const lines = text.split('\n');
    const out: string[] = [];
    const closingFence = /^( *)(`{3,}|~{3,})[ \t]*$/;
    const indentOf = (line: string) => /^ */.exec(line)![0].length;
    // Per fence character, the longest closing fence indented at most three spaces on each line or
    // any line after it (0 for none), so a top-level opener that is never closed is known at once
    // instead of by scanning to the end: any number of unclosed fences costs time linear in the text.
    const longestAfter: Record<string, Int32Array> = { '`': new Int32Array(lines.length + 1), '~': new Int32Array(lines.length + 1) };
    for (let i = lines.length - 1; i >= 0; i--) {
        longestAfter['`'][i] = longestAfter['`'][i + 1];
        longestAfter['~'][i] = longestAfter['~'][i + 1];
        const close = closingFence.exec(lines[i]);
        if (close && close[1].length <= 3) longestAfter[close[2][0]][i] = Math.max(longestAfter[close[2][0]][i], close[2].length);
    }
    // The list items the scan is inside, innermost last: the column each one's content starts at.
    const items: { column: number; start: number }[] = [];
    // Every closing fence, in order, and for each line the next line at or after it holding content:
    // an item's end is found by stepping over its content lines only, not its blank ones.
    const closers: { at: number; indent: number; char: string; length: number }[] = [];
    const nextContent = new Int32Array(lines.length + 1).fill(lines.length);
    for (let i = lines.length - 1; i >= 0; i--) {
        nextContent[i] = trimAsciiWhitespace(lines[i]) ? i : nextContent[i + 1];
    }
    for (let i = 0; i < lines.length; i++) {
        const close = closingFence.exec(lines[i]);
        if (close) closers.push({ at: i, indent: close[1].length, char: close[2][0], length: close[2].length });
    }
    /** The index in `list` (ascending `at`) of the first entry after line `line`. */
    const firstAfter = (list: { at: number }[], line: number) => {
        let lo = 0, hi = list.length;
        while (lo < hi) { const mid = (lo + hi) >> 1; if (list[mid].at <= line) lo = mid + 1; else hi = mid; }
        return lo;
    };
    // For a list item holding a fence, computed once: the closing fences indented as the item's fences
    // are (up to three spaces past its content) among its lines, where a line with content indented
    // less than its content ends it, and per fence character the longest of them from each one on. An
    // opener in the item is then decided at once, as a top-level one is. Only those fences are kept,
    // not a value for each of the item's lines: items nested in items, each holding an unclosed fence
    // and ending in blank lines, took time and memory for every line of every one of them.
    const itemClosers = new Map<number, { list: { at: number }[]; longest: Record<string, Int32Array> }>();
    const closersIn = (item: { column: number; start: number }) => {
        let found = itemClosers.get(item.start);
        if (!found) {
            let end = nextContent[item.start + 1] ?? lines.length;
            while (end < lines.length && indentOf(lines[end]) >= item.column) end = nextContent[end + 1];
            const list: { at: number; char: string; length: number }[] = [];
            for (let k = firstAfter(closers, item.start); k < closers.length && closers[k].at < end; k++) {
                const close = closers[k];
                if (close.indent >= item.column && close.indent - item.column <= 3) list.push(close);
            }
            const longest = { '`': new Int32Array(list.length + 1), '~': new Int32Array(list.length + 1) };
            for (let k = list.length - 1; k >= 0; k--) {
                longest['`'][k] = longest['`'][k + 1];
                longest['~'][k] = longest['~'][k + 1];
                longest[list[k].char as '`' | '~'][k] = Math.max(longest[list[k].char as '`' | '~'][k], list[k].length);
            }
            found = { list, longest };
            itemClosers.set(item.start, found);
        }
        return found;
    };
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const indent = indentOf(line);
        const item = /^( *)(?:[-*+]|\d{1,9}[.)])( {1,4})(?=\S)/.exec(line);
        while (items.length && (item ? item[1].length < items[items.length - 1].column : trimAsciiWhitespace(line) && indent < items[items.length - 1].column)) items.pop();
        if (item) items.push({ column: item[0].length, start: i });
        const container = items.length && !item ? items[items.length - 1] : undefined;
        const base = container?.column ?? 0;
        const open = /^( *)(`{3,}|~{3,})(.*)$/.exec(line);
        const char = open?.[2][0];
        if (open && char && indent - base <= 3 && !(char === '`' && open[3].includes('`'))) {
            const [, indentStr, fence, info] = open;
            const closes = (candidate: string) => {
                const close = closingFence.exec(candidate);
                return !!close && close[2][0] === char && close[2].length >= fence.length && close[1].length >= base && close[1].length - base <= 3;
            };
            const inItem = container && closersIn(container);
            const closed = inItem ? inItem.longest[char][firstAfter(inItem.list, i)] >= fence.length : longestAfter[char][i + 1] >= fence.length;
            if (closed) {
                let j = i + 1;
                while (!closes(lines[j])) j++;
                const dedent = new RegExp(`^ {0,${indentStr.length}}`);
                out.push(replace(info.trim().split(/\s+/)[0] || '', lines.slice(i + 1, j).map(content => content.replace(dedent, '')).join('\n')));
                i = j;
                continue;
            }
        }
        out.push(line);
    }
    return out.join('\n');
}

/** Plain text of parsed inline nodes, leaving out source comments: a hidden note is not text. */
const plainTextOf = (nodes: OfficeContentNode[]): string => nodes.map(n => (isSourceComment(n) ? '' : n.text || '')).join('');

export const parseMarkdown = async (buffer: Buffer, config: FullOfficeParserConfig): Promise<OfficeParserAST> => {
    // Honour cancellation requests before the line-by-line Markdown scanning loop begins.
    // Markdown parsing is entirely synchronous and CPU-bound, so failing fast avoids
    // processing content whose result will be discarded anyway.
    checkAbortSignal(config.abortSignal);

    let textStr = buffer.toString('utf-8');
    textStr = textStr.replace(/\r\n/g, '\n');

    const content: OfficeContentNode[] = [];
    const metadata: OfficeMetadata = {};
    const attachments: OfficeAttachment[] = [];

    // Parse YAML Front Matter
    if (/^---\n---[ \t]*(?:\n|$)/.test(textStr)) {
        // Empty frontmatter block: strip it so `---\n---` isn't misread as a setext `## ---`
        // heading (empty metadata used to emit exactly this shape, and other producers do too).
        textStr = textStr.replace(/^---\n---[ \t]*(?:\n|$)/, '');
    } else if (textStr.startsWith('---\n') && !/^---\n[ \t]*(?:\n|$)/.test(textStr)) {
        // (Not a `---` followed by a blank line: that is a rule, as Pandoc reads it, and what the
        // generator writes for a rule or a page, slide or sheet boundary starting the document.
        // Read as front matter, everything to the next rule was lost.)
        const endIdx = textStr.indexOf('\n---\n', 4);
        // Front matter is YAML: each line a `key: value` (a list item, an indented or a comment line
        // after one). Otherwise `---`, text and `---` are a rule and a setext heading, as Pandoc reads
        // them; read as front matter, the text was lost.
        const isYaml = endIdx !== -1 && textStr.substring(4, endIdx).split('\n').every(line =>
            !line.trim() || /^[^\s:#][^:]*:(?:\s|$)/.test(line) || /^\s+\S|^-\s|^#/.test(line));
        if (endIdx !== -1 && isYaml) {
            const frontMatter = textStr.substring(4, endIdx);
            textStr = textStr.substring(endIdx + 5);

            const lines = frontMatter.split('\n');
            const customProps: Record<string, any> = {};
            const nativeProps: Record<string, any> = {};

            for (const line of lines) {
                const match = line.match(/^([^:]+):\s*(.*)$/);
                if (match) {
                    const key = match[1].trim();
                    const rawVal = match[2].trim();
                    // A quoted scalar is explicitly a string in YAML: strip the quotes but never
                    // coerce it, so `version: "123"` / `flag: "true"` keep their string-ness across
                    // a save/reload cycle instead of silently degrading to a number/boolean on the
                    // next parse (which the generator would then re-emit unquoted, losing the type
                    // permanently). Only bare, unquoted scalars coerce.
                    const isQuoted = /^"(.*)"$/.test(rawVal) || /^'(.*)'$/.test(rawVal);
                    // A double-quoted scalar's escapes (`\"`, `\\`, `\u003c`, as the generator writes
                    // them) are decoded as JSON decodes them, which YAML's double-quoted form extends;
                    // in a single-quoted one, `''` is a quote.
                    const val = /^"(.*)"$/.test(rawVal) ? decodeDoubleQuoted(rawVal)
                        : /^'(.*)'$/.test(rawVal) ? rawVal.slice(1, -1).replace(/''/g, "'") : rawVal;

                    let parsedVal: any = val;
                    if (!isQuoted && rawVal.startsWith('[') && rawVal.endsWith(']')) {
                        // Flow-array (`tags: [a, b]`) or JSON-array (`tags: ["a","b"]`) value -
                        // parse into a real array instead of storing the literal bracket string,
                        // so it round-trips symmetrically with MarkdownGenerator's frontmatter output.
                        try {
                            const jsonParsed = JSON.parse(rawVal);
                            parsedVal = Array.isArray(jsonParsed) ? jsonParsed : val;
                        } catch {
                            const inner = rawVal.slice(1, -1).trim();
                            // Quoted items decoded as quoted scalars are (above).
                            parsedVal = inner === '' ? [] : splitFlowArrayItems(inner).map(item => (/^"(.*)"$/.test(item) ? decodeDoubleQuoted(item)
                                : /^'(.*)'$/.test(item) ? item.slice(1, -1).replace(/''/g, "'") : item));
                        }
                    } else if (isQuoted) parsedVal = val;
                    else if (val === 'true') parsedVal = true;
                    else if (val === 'false') parsedVal = false;
                    else if (!isNaN(Number(val)) && val !== '') parsedVal = Number(val);

                    setOwn(nativeProps, key, parsedVal);

                    if (key === 'title') metadata.title = val;
                    else if (key === 'author') metadata.author = val;
                    else if (key === 'created') metadata.created = new Date(val);
                    else if (key === 'modified') metadata.modified = new Date(val);
                    else if (key === 'description') metadata.description = val;
                    else {
                        setOwn(customProps, key, parsedVal);
                    }
                }
            }
            if (Object.keys(customProps).length > 0) metadata.customProperties = customProps;
            if (Object.keys(nativeProps).length > 0) metadata.nativeProperties = nativeProps;
        }
    }
    // Blocks lifted out before block splitting (code, math, admonitions, comments) are replaced by a
    // placeholder line. The placeholder is built from private-use characters the document does not
    // contain, so literal text such as `__CODE_BLOCK_0__` in the document is never mistaken for one.
    let sentinel = '\uE000';
    while (textStr.includes(sentinel)) sentinel += '\uE001';
    const placeholder = (kind: string, n: number) => `${sentinel}${kind}_${n}${sentinel}`;
    const placeholderIndex = (block: string, kind: string): number | null => {
        const m = block.startsWith(sentinel + kind + '_') && block.endsWith(sentinel) ? block.slice(sentinel.length + kind.length + 1, -sentinel.length) : null;
        return m !== null && /^\d+$/.test(m) ? parseInt(m, 10) : null;
    };

    // Extract fenced code blocks first to protect their contents (see liftFencedCode).
    const codeBlocks: string[] = [];
    const liftCode = (text: string): string => liftFencedCode(text, (lang, code) => {
        const id = placeholder('CODE_BLOCK', codeBlocks.length);
        codeBlocks.push(JSON.stringify({ lang, code }));
        return `\n\n${id}\n\n`;
    });
    /** The code node a lifted block's placeholder stands for, or null for any other text. */
    const liftedCodeNode = (part: string): OfficeContentNode | null => {
        const index = placeholderIndex(part.trim(), 'CODE_BLOCK');
        if (index === null || index >= codeBlocks.length) return null;
        const data = JSON.parse(codeBlocks[index]);
        return { type: 'code', text: data.code, metadata: { language: data.lang } as CodeMetadata };
    };
    textStr = liftCode(textStr);

    // MDX components (`<Callout>...</Callout>`, `<Chart />`) are stripped and their inner Markdown kept,
    // after code blocks are lifted, so a component in a code sample stays code (see stripMdxComponents).
    textStr = stripMdxComponents(textStr);

    // Extract block math ($$\n...\n$$) before block splitting, mirroring the code-block
    // pre-pass above - its body may contain blank lines that would otherwise fragment it.
    // Inline math ($...$) is handled directly in parseInline below.
    const mathBlocks: string[] = [];
    textStr = textStr.replace(/^\$\$\n([\s\S]*?)\n\$\$$/gm, (_match, latex: string) => {
        const id = placeholder('MATH_BLOCK', mathBlocks.length);
        // The generator indents a content line of `$$` (which would close the block) by one space,
        // and one already indented likewise: one space comes off each.
        mathBlocks.push(latex.replace(/^ (?= *\$\$$)/gm, ''));
        return `\n\n${id}\n\n`;
    });
    // Single-line `$$...$$` occupying its own line is display (block) math too. Without this it
    // falls through to the inline `$...$` tokenizer, which matches the INNER `$\int$` and leaks
    // the outer pair as two stray literal `$`. Runs after the multi-line pass, whose placeholders
    // carry no `$$` and so can't be re-matched. `(?!\$)` rejects `$$$...`/empty `$$$$`.
    textStr = textStr.replace(/^\$\$(?!\$)([^\n]+?)\$\$[ \t]*$/gm, (_match, latex) => {
        const id = placeholder('MATH_BLOCK', mathBlocks.length);
        mathBlocks.push(latex);
        return `\n\n${id}\n\n`;
    });

    // Extract GLFM-style fenced-div admonitions (`:::note ... :::`) before block splitting,
    // since their body may itself contain blank lines that would otherwise fragment them.
    // The `> [!NOTE]` GitHub form doesn't need this - it's detected inline in the blockquote
    // branch below, since a `>`-prefixed block never contains a real blank line.
    // A `:::type` line opens one and the next `:::` line closes it; one of an unrecognised type is left
    // as literal text, its body included. Lines are scanned once, with each line's next closing line
    // found in advance (a lazy pattern rescanned to the end for every line that opened one unclosed).
    const admonitionBlocks: string[] = [];
    {
        const lines = textStr.split('\n');
        const nextClose = new Int32Array(lines.length + 1).fill(lines.length);
        for (let i = lines.length - 1; i >= 0; i--) nextClose[i] = /^:::[ \t]*$/.test(lines[i]) ? i : nextClose[i + 1];
        const out: string[] = [];
        for (let i = 0; i < lines.length; i++) {
            const open = /^:::(\w+)[ \t]*$/.exec(lines[i]);
            const close = open ? nextClose[i + 1] : lines.length;
            if (open && close < lines.length) {
                const admonitionType = ADMONITION_TYPE_MAP[open[1].toLowerCase()];
                if (admonitionType) {
                    const id = placeholder('ADMONITION', admonitionBlocks.length);
                    admonitionBlocks.push(JSON.stringify({ admonitionType, body: lines.slice(i + 1, close).join('\n') }));
                    out.push('', '', id, '', '');
                } else {
                    appendAll(out, lines.slice(i, close + 1));
                }
                i = close;
                continue;
            }
            out.push(lines[i]);
        }
        textStr = out.join('\n');
    }

    // Extract source comments that stand on their own lines (`<!-- ... -->`, possibly spanning several
    // lines, blank ones included) before block splitting, mirroring the code/math/admonition pre-passes -
    // otherwise the blank-line split tears a multi-line comment apart and its pieces reparse as visible
    // text. Runs after code-block extraction, so a comment inside a fence stays code. The comment must
    // open at column 0 (an indented one inside a list item is left to the inline path, so it doesn't
    // split the list) and its closing `-->` must end the line; the body can't contain `-->`, so the match
    // always ends at the comment's real close. The raw body is kept verbatim for re-emission.
    // A scan with indexOf rather than a regex: a lazy regex body re-scans to the end of the text for
    // every unclosed `<!--`, which is quadratic on hostile input.
    const htmlComments: string[] = [];
    {
        const closes: CommentCloseCache = { at: -1, from: Number.MAX_SAFE_INTEGER };
        let out = '';
        let pos = 0;
        let search = 0;
        for (;;) {
            const open = textStr.indexOf('<!--', search);
            if (open === -1) break;
            search = open + 1;
            if (open > 0 && textStr[open - 1] !== '\n') continue;
            const found = commentAt(textStr, open, closes);
            if (!found) continue;
            let end = found.end;
            while (textStr[end] === ' ' || textStr[end] === '\t') end++;
            if (end < textStr.length && textStr[end] !== '\n') continue;
            out += `${textStr.slice(pos, open)}\n\n${placeholder('HTML_COMMENT', htmlComments.length)}\n\n`;
            htmlComments.push(found.body);
            pos = search = end;
        }
        textStr = out + textStr.slice(pos);
    }

    // Extract footnote definitions (`[^id]: text`) before block splitting, since
    // definitions conventionally live at the end of the document, after every place
    // they're referenced - inline parsing below needs the full map upfront. The first line
    // may be followed by continuation lines indented one level (4 spaces or a tab), which are
    // dedented and joined onto the definition (Pandoc/GFM). A 4-space-indented block right after
    // a definition is therefore read as its continuation rather than as a standalone code block.
    // Supported (lossless) shape: contiguous continuation - the indented lines follow the
    // definition with no blank line between them. Known limitation (6.E.1): a continuation
    // separated from the definition by a BLANK line is not folded in - the regex below stops at
    // the blank line, and the indented block after it re-parses as a fenced/indented code block on
    // save. Multi-paragraph footnotes should therefore use the contiguous form.
    const footnoteDefinitions = new Map<string, string>();
    // Every id a `[^id]` reference consumes, so definitions that are never referenced can be
    // detected at the end and preserved rather than silently dropped (see the orphan sweep below).
    const referencedFootnoteIds = new Set<string>();
    // One reused note node per referenced id. Repeated `[^id]` references are a single shared
    // footnote in Markdown, so they must not each materialise a full copy of the body - the
    // generators would otherwise renumber them to [^1]/[^2] and duplicate the definition. Office
    // notes reach the generators as distinct objects even when they share a numeric id, so those
    // stay separate; only genuinely shared Markdown references collapse.
    const footnoteNodesById = new Map<string, OfficeContentNode>();
    // A label here (as below, for abbreviations and link references) holds at most 999 characters,
    // CommonMark's limit for a link label. The bound also bounds the work: a label may run across
    // lines, so without it every line starting with `[` would be scanned to the end of the document.
    // A definition continues on lines indented four columns, blank lines between them allowed, so a
    // note may hold several paragraphs or a list (as Pandoc and the generator write one).
    textStr = textStr.replace(/^\[\^([^\]]{1,999})\]:[ \t]*(.*(?:\n(?:[ \t]*\n)*(?: {4}|\t).*)*)$/gm, (_match, id, definition) => {
        const dedented = String(definition)
            .split('\n')
            .map((line: string, i: number) => i === 0 ? line : line.replace(/^(?: {4}|\t)/, ''))
            .join('\n');
        footnoteDefinitions.set(id, trimAsciiWhitespace(dedented));
        return '';
    });

    // Extract Markdown Extra abbreviation definitions (`*[HTML]: Hypertext Markup Language`)
    // before block splitting, for the same reason as footnotes: they conventionally live
    // at the end of the document.
    const abbreviationDefinitions = new Map<string, string>();
    textStr = textStr.replace(/^\*\[([^\]]{1,999})\]:[ \t]*(.*)$/gm, (_match, abbr, definition) => {
        // Decoded, as the text they are matched against is.
        abbreviationDefinitions.set(decodeMarkdownText(abbr), decodeMarkdownText(trimAsciiWhitespace(definition)));
        return '';
    });

    // Extract link/image reference definitions (`[ref]: /url "title"`) before block
    // splitting, for the same reason as footnotes/abbreviations: they conventionally
    // live at the end of the document, after every place they're referenced. Keyed by
    // trimmed/lowercased label, matching CommonMark's case-insensitive reference matching.
    // A label ends at its first unescaped `]` (`[Step \[2\]: configure](url)`, a link, was read as the
    // label `Step \[2\` and the paragraph lost), and a definition does not interrupt a paragraph: one
    // on the line after text (`- item` then `[x]: y`) is that text's continuation.
    const linkDefinitions = new Map<string, { url: string; title?: string }>();
    let lastDefinitionEnd = -1;
    textStr = textStr.replace(/^\[((?:[^\]\\]|\\[\s\S]){1,999})\]:[ \t]*(\S+)(?:[ \t]+"((?:[^"\\]|\\.)*)")?[ \t]*$/gm, (match: string, label: string, url: string, title: string | undefined, offset: number, whole: string) => {
        if (offset > 0 && lastDefinitionEnd !== offset - 1) {
            const previousLine = whole.slice(whole.lastIndexOf('\n', offset - 2) + 1, offset - 1);
            if (trimAsciiWhitespace(previousLine)) return match;
        }
        lastDefinitionEnd = offset + match.length;
        linkDefinitions.set(label.trim().toLowerCase(), { url: decodeMarkdownText(url), title: title === undefined ? undefined : decodeMarkdownText(title) });
        return '';
    });

    // Parses a Pandoc-style attribute list body (the part inside `{...}`), e.g.
    // `width=50% .centered` or `align=right`. Per MARKDOWN_DIALECT.md §15's Decisions,
    // the vocabulary matches ImageMetadata/TableMetadata's own width/align fields;
    // several class-name spellings are accepted on import for compatibility with
    // hand-written content, but the generator only ever emits canonical `align=value`.
    const parseAttributeList = (attrStr: string): { width?: string; align?: 'left' | 'center' | 'right' } => {
        const result: { width?: string; align?: 'left' | 'center' | 'right' } = {};
        for (const token of attrStr.trim().split(/\s+/).filter(Boolean)) {
            const kv = token.match(/^([a-zA-Z-]+)=(.+)$/);
            if (kv) {
                if (kv[1] === 'width') result.width = kv[2];
                else if (kv[1] === 'align' && ['left', 'center', 'right'].includes(kv[2])) result.align = kv[2] as any;
            } else if (token.startsWith('.')) {
                const cls = token.slice(1).toLowerCase();
                if (cls === 'left' || cls === 'align-left') result.align = 'left';
                else if (cls === 'center' || cls === 'centered' || cls === 'align-center') result.align = 'center';
                else if (cls === 'right' || cls === 'align-right') result.align = 'right';
            }
        }
        return result;
    };

    // Attribute list for an embed leaf directive `{id=... src=... width=... height=... align=...}`.
    // Superset of parseAttributeList (adds id/src/height); space-separated `k=v` tokens.
    const parseEmbedDirectiveAttrs = (attrStr: string): { id?: string; src?: string; width?: string; height?: string; align?: 'left' | 'center' | 'right' } => {
        const result: { id?: string; src?: string; width?: string; height?: string; align?: 'left' | 'center' | 'right' } = {};
        for (const token of attrStr.trim().split(/\s+/).filter(Boolean)) {
            const kv = token.match(/^([a-zA-Z-]+)=(.+)$/);
            if (!kv) continue;
            const [, key, val] = kv;
            if (key === 'id') result.id = val;
            else if (key === 'src') result.src = val;
            else if (key === 'width') result.width = val;
            else if (key === 'height') result.height = val;
            else if (key === 'align' && ['left', 'center', 'right'].includes(val)) result.align = val as 'left' | 'center' | 'right';
        }
        return result;
    };

    // Extracts a YouTube video id from any of its URL shapes (watch?v=, youtu.be/, /embed/,
    // img.youtube.com/vi/). Returns undefined for a non-YouTube URL. Used only by the opt-in
    // folk-form import (embedFolkForms).
    const extractYoutubeId = (url: string): string | undefined => {
        if (!url || !/(?:youtu\.be|youtube(?:-nocookie)?\.com)/.test(url)) return undefined;
        const m = url.match(/(?:youtu\.be\/|\/embed\/|[?&]v=|\/vi\/)([A-Za-z0-9_-]+)/);
        return m ? m[1] : undefined;
    };

    // The note for a footnote reference to `noteId`, marking the definition as used. The same note
    // object serves every reference to the id (see the map's declaration): the first reference builds
    // the body, the rest share it, so the generators assign one key and emit one definition.
    //
    // A note's body is read after the document (readQueuedNoteBodies), from a queue rather than while
    // the reference is read, so notes referring to notes cost no stack (a note referring to itself
    // recursed without end). A note may refer to another note only if that one refers to none: nesting
    // stays one level deep and never loops back, so the AST is a tree of bounded depth, which writers
    // walk and JSON serializes. Any other reference is null (the caller keeps it as text), and a note
    // referred to only that way stays an unreferenced definition.
    const footnoteNode = (noteId: string): OfficeContentNode | null => {
        if (readingNote && FOOTNOTE_REFERENCE.test(footnoteDefinitions.get(noteId) ?? '')) return null;
        referencedFootnoteIds.add(noteId);
        let noteNode = footnoteNodesById.get(noteId);
        if (!noteNode) {
            noteNode = { type: 'note', text: '', children: [], metadata: { noteType: 'footnote', noteId } };
            footnoteNodesById.set(noteId, noteNode);
            notesToRead.push({ note: noteNode, definition: footnoteDefinitions.get(noteId) ?? '' });
        }
        return noteNode;
    };
    // Bodies waiting to be read, and the note whose body is being read.
    const notesToRead: { note: OfficeContentNode; definition: string }[] = [];
    let readingNote: OfficeContentNode | undefined;

    // `displayMath` is true only for the lines of a paragraph, where display math (`$$...$$` written
    // inside the text) becomes a block the paragraph is split around (see splitAtDisplayMath).
    // Everywhere else (a heading, list item, table cell, quote, note, or inside emphasis) a block
    // cannot go, so it is inline math there.
    const parseInline = (text: string, currentFormatting: TextFormatting = {}, displayMath = false): OfficeContentNode[] => {
        const nodes: OfficeContentNode[] = [];
        // Text of this call's own, with its character references decoded (`&amp;` is `&`). Once: what a
        // nested call returns (emphasis, a link's text) it decoded already, and decoding that again
        // turned a literal `&amp;quot;` inside `**...**` into `"`.
        const plainText = (t: string): OfficeContentNode => ({ type: 'text', text: currentFormatting.font === 'monospace' ? t : decodeCharacterReferences(t), formatting: Object.keys(currentFormatting).length > 0 ? { ...currentFormatting } : undefined });

        // Builds the same image/link node shape regardless of whether the URL came from
        // an inline `(url)` or a resolved reference definition - shared by the inline
        // image/link branch and the two reference-style branches below.
        // A picture: an attachment when its source is data (with extractAttachments), else a reference.
        const imageNode = (url: string, altText: string, title: string | undefined, attrs?: object): OfficeContentNode => {
            if (url.startsWith('data:')) {
                const dataMatch = url.match(/^data:([^;]+);base64,(.*)$/);
                if (dataMatch && config.extractAttachments) {
                    const mimeType = dataMatch[1] as any;
                    const data = dataMatch[2];
                    const name = `image_${attachments.length + 1}.${mimeType.split('/')[1]}`;
                    attachments.push({
                        type: 'image',
                        mimeType,
                        data,
                        name,
                        extension: mimeType.split('/')[1]
                    });
                    return { type: 'image', metadata: { attachmentName: name, altText, title, ...attrs } as ImageMetadata };
                }
            }
            return { type: 'image', metadata: { url, altText, title, ...attrs } as ImageMetadata };
        };
        const buildLinkOrImageNodes = (isImage: boolean, altText: string, rawUrl: string, attrsStr?: string): OfficeContentNode[] => {
            const { url, title } = splitUrlTitle(rawUrl);
            if (isImage) {
                // Alt text is plain text, decoded as a Markdown renderer decodes it; a Pandoc-style
                // attribute list may follow the image, e.g. {width=50% .centered}.
                return [imageNode(url, decodeMarkdownText(altText), title, attrsStr !== undefined ? parseAttributeList(attrsStr) : undefined)];
            }
            return applyLink(parseInline(altText, currentFormatting), url, title);
        };
        // `linkNodes` linked to `url`: its runs, and a picture among them (a badge) carries the link itself.
        const applyLink = (linkNodes: OfficeContentNode[], url: string, title: string | undefined): OfficeContentNode[] => {
            // A link with no text (`[](url)`) keeps its target, on an empty run.
            if (linkNodes.length === 0) linkNodes.push(plainText(''));
            // A target in this document (`#id`) is internal, as the other parsers read it.
            const linkType = url.startsWith('#') ? 'internal' : 'external';
            linkNodes.forEach(n => {
                if (n.type === 'text') {
                    n.metadata = { link: url, linkType, title } as TextMetadata;
                } else if (n.type === 'image') {
                    // A linked image (a badge, `[![alt](src)](target)`) carries the link itself.
                    n.metadata = { ...n.metadata, link: url, linkType, ...(title !== undefined && { linkTitle: title }) } as ImageMetadata;
                }
            });
            return linkNodes;
        };

        // Raw inline HTML (see HTML_INLINE_ELEMENTS): what an element holds, with its formatting; a
        // code element's content as written; `<a href>` a link (`<a id>` an anchor before its content);
        // `<abbr title>` an abbreviation; `<q>` its content in quotation marks; `<img>` a picture.
        const inlineHtml = (tag: string, attrsSource: string, inner: string): OfficeContentNode[] => {
            const attrs = htmlAttributes(attrsSource);
            if (tag === 'img') {
                const width = attrs.get('width');
                const align = attrs.get('align')?.toLowerCase();
                return [imageNode(attrs.get('src')!, attrs.get('alt') ?? '', attrs.get('title') || undefined, {
                    ...(width && { width }),
                    ...((align === 'left' || align === 'center' || align === 'right') && { align }),
                })];
            }
            const element = HTML_INLINE_ELEMENTS[tag];
            if (element === 'code') return [{ type: 'text', text: decodeCharacterReferences(inner), formatting: { ...currentFormatting, font: 'monospace' } }];
            const content = parseInline(inner, { ...currentFormatting, ...element });
            if (tag === 'a') {
                const href = attrs.get('href');
                if (href !== undefined) return applyLink(content, href, attrs.get('title') || undefined);
                const id = attrs.get('id') || attrs.get('name');
                return id ? [anchorMark([id]), ...content] : content;
            }
            const title = attrs.get('title');
            if (tag === 'abbr' && title) {
                for (const n of content) if (n.type === 'text') n.metadata = { ...(n.metadata as object), abbreviationTitle: title } as TextMetadata;
            }
            return tag === 'q' ? [plainText('“'), ...content, plainText('”')] : content;
        };
        // Where each inline element's next closing tag is, from where this text's scan stands: looked
        // up once while it lies ahead, and never again once none follows, so any number of elements
        // never closed costs time linear in the text.
        const htmlClosersAt = new Map<string, { at: number; length: number } | null>();
        const closeHtml = (tag: string, from: number): { at: number; length: number } | null => {
            const known = htmlClosersAt.get(tag);
            if (known === null || (known && known.at >= from)) return known;
            const closer = htmlCloser(tag);
            closer.lastIndex = from;
            const m = closer.exec(text);
            const found = m ? { at: m.index, length: m[0].length } : null;
            htmlClosersAt.set(tag, found);
            return found;
        };

        // The tokenizer (see INLINE_TOKENS) finds the next inline construct; text between constructs
        // is plain text. A backtick run, a run of underscores and a comment's `<!--` are matched as
        // openers only: where each closes is looked up (codeSpanCloser, underscoreCloser, commentAt),
        // in time linear in the text however many never close. An opener with no closer is ordinary
        // text, left in its run. The tokenizer is shared by every call (compiling it for each piece of
        // text doubled the time of a document of short paragraphs); each search sets where it starts,
        // so the nested calls made while handling a match cannot disturb this one.
        const closeCodeSpan = codeSpanCloser(text);
        const closeUnderscores = underscoreCloser(text);
        const closeStars = delimiterCloser(text, '*');
        const closeTildes = delimiterCloser(text, '~');
        const closeEquals = delimiterCloser(text, '=');
        let lastIndex = 0; // end of the text already emitted
        let next = 0; // where the next search starts
        const closes: CommentCloseCache = { at: -1, from: Number.MAX_SAFE_INTEGER };

        for (;;) {
            INLINE_TOKEN_REGEX.lastIndex = next;
            const match = INLINE_TOKEN_REGEX.exec(text);
            if (!match) break;
            next = INLINE_TOKEN_REGEX.lastIndex;
            const g = match.groups!;
            let comment: { body: string; end: number } | null = null;
            if (g.htmlComment !== undefined) {
                comment = commentAt(text, match.index, closes);
                if (!comment) { next = match.index + 4; continue; }
                next = comment.end;
            }
            let anchorIds: string[] = [];
            if (g.anchorTag !== undefined) {
                const id = /\sid="([^"]*)"/i.exec(g.anchorTag)?.[1] || /\sname="([^"]*)"/i.exec(g.anchorTag)?.[1];
                if (!id) { next = match.index + 2; continue; }
                anchorIds = [id];
            }
            let closeAt = -1;
            // The stars of a run that open emphasis: as many as a closing run allows, from three down,
            // the rest are text (`***a**` is a star, then bold `a`).
            let starLength = 0;
            if (g.codeFence !== undefined || g.underscores !== undefined) {
                const run = g.codeFence ?? g.underscores;
                closeAt = g.codeFence !== undefined ? closeCodeSpan(next, run.length) : closeUnderscores(next, run.length);
                if (closeAt === -1) continue;
                next = closeAt + run.length;
            } else if (g.stars !== undefined) {
                for (starLength = g.stars.length; starLength > 0; starLength--) {
                    closeAt = closeStars(next, starLength);
                    if (closeAt !== -1) break;
                }
                if (closeAt === -1) continue;
                next = closeAt + starLength;
            } else if (g.tildes !== undefined || g.equals !== undefined) {
                closeAt = (g.tildes !== undefined ? closeTildes : closeEquals)(next, 2);
                if (closeAt === -1) continue;
                next = closeAt + 2;
            } else if (g.htmlTag !== undefined) {
                // Not an element it reads, a picture with no source, or an element never closed: text.
                const tag = g.htmlTag.toLowerCase();
                if (tag === 'img' ? !htmlAttributes(g.htmlAttrs ?? '').get('src') : HTML_INLINE_ELEMENTS[tag] === undefined) { next = match.index + 1; continue; }
                if (tag !== 'img') {
                    const close = closeHtml(tag, next);
                    if (!close) { next = match.index + 1; continue; }
                    closeAt = close.at;
                    next = close.at + close.length;
                }
            }
            const constructStart = g.stars !== undefined ? match.index + g.stars.length - starLength : match.index;
            if (constructStart > lastIndex) {
                nodes.push(plainText(text.substring(lastIndex, constructStart)));
            }

            if (g.esc !== undefined) { // Backslash-escaped punctuation
                nodes.push(plainText(g.esc));
            } else if (g.imgAlt !== undefined) { // Image or Link
                appendAll(nodes, buildLinkOrImageNodes(g.imgBang === '!', g.imgAlt, g.imgUrl, g.imgAttrs));
            } else if (g.stars !== undefined) { // Italic (*), bold (**), or both (***)
                const formatting: TextFormatting = { ...currentFormatting };
                if (starLength >= 2) formatting.bold = true;
                if (starLength !== 2) formatting.italic = true;
                appendAll(nodes, parseInline(text.slice(match.index + g.stars.length, closeAt), formatting));
            } else if (g.underscores !== undefined) { // Italic (_), bold (__), or both (___)
                const formatting: TextFormatting = { ...currentFormatting };
                if (g.underscores.length !== 2) formatting.italic = true;
                if (g.underscores.length >= 2) formatting.bold = true;
                appendAll(nodes, parseInline(text.slice(match.index + g.underscores.length, closeAt), formatting));
            } else if (g.tildes !== undefined) { // Strikethrough
                appendAll(nodes, parseInline(text.slice(match.index + 2, closeAt), { ...currentFormatting, strikethrough: true }));
            } else if (g.equals !== undefined) { // ==highlight== (Obsidian/extended); additive on import
                appendAll(nodes, parseInline(text.slice(match.index + 2, closeAt), { ...currentFormatting, backgroundColor: '#ffff00' }));
            } else if (g.codeFence !== undefined) { // Inline code, closed by a run of as many backticks
                nodes.push({ type: 'text', text: codeSpanContent(text.slice(match.index + g.codeFence.length, closeAt)), formatting: { ...currentFormatting, font: 'monospace' } });
            } else if (g.underline !== undefined) { // Underline
                appendAll(nodes, parseInline(g.underline, { ...currentFormatting, underline: true }));
            } else if (g.subscript !== undefined) { // Subscript
                appendAll(nodes, parseInline(g.subscript, { ...currentFormatting, subscript: true }));
            } else if (g.superscript !== undefined) { // Superscript
                appendAll(nodes, parseInline(g.superscript, { ...currentFormatting, superscript: true }));
            } else if (g.htmlTag !== undefined) { // Raw inline HTML: an element, or a picture (see inlineHtml)
                appendAll(nodes, inlineHtml(g.htmlTag.toLowerCase(), g.htmlAttrs ?? '', closeAt === -1 ? '' : text.slice(match.index + match[0].length, closeAt)));
            } else if (g.anchorTag !== undefined) { // An empty anchor: an id (see resolveAnchorMarks)
                nodes.push(anchorMark(anchorIds));
            } else if (g.lineBreak !== undefined) { // Raw inline <br>/<br/>/<br /> - a hard line break.
                // MarkdownGenerator emits a raw <br> for a line break inside a table cell (a GFM pipe
                // cell can't hold a newline), so the parser must read it back symmetrically as a break
                // node instead of escaping it to literal `&lt;br&gt;` text and destroying it.
                nodes.push({ type: 'break', metadata: { breakType: 'carriageReturn' } as BreakMetadata });
            } else if (g.spanContent !== undefined) { // Inline styled span: color / highlight / font-size
                const style = g.spanStyle || '';
                const styled: TextFormatting = { ...currentFormatting };
                // Anchor each property to a declaration boundary so `color` doesn't match inside
                // `background-color`.
                const prop = (name: string): string | undefined => {
                    const m = style.match(new RegExp(`(?:^|;)\\s*${name}\\s*:\\s*([^;]+)`, 'i'));
                    return m ? m[1].trim() : undefined;
                };
                const color = prop('color');
                if (color) styled.color = color;
                const background = prop('background-color');
                if (background) styled.backgroundColor = background;
                const size = prop('font-size');
                if (size) styled.size = size;
                appendAll(nodes, parseInline(g.spanContent, styled));
            } else if (g.footnoteId !== undefined) { // Footnote reference
                const noteId = g.footnoteId;
                // ignoreNotes drops footnotes at parse time (as in DOCX/ODT/PDF): swallow the marker and
                // attach nothing. Advance lastIndex past the marker (so the gap text is not re-emitted)
                // before skipping. The orphan sweep below is likewise skipped.
                if (config.ignoreNotes) { lastIndex = next; continue; }
                const noteNode = footnoteNode(noteId);
                // Notes attach to the preceding text node (matches WordParser's convention);
                // fall back to an empty text node if the reference opens the inline run. A reference
                // that would make a note contain itself stays text.
                if (!noteNode) {
                    nodes.push(plainText(match[0]));
                } else if (nodes.length > 0) {
                    const target = nodes[nodes.length - 1];
                    if (!target.notes) target.notes = [];
                    target.notes.push(noteNode);
                } else {
                    nodes.push({ type: 'text', text: '', notes: [noteNode] });
                }
            } else if (g.citationKey !== undefined) { // Citation reference
                nodes.push({ type: 'text', text: g.citationKey, metadata: { citationKey: g.citationKey } as TextMetadata });
            } else if (g.wikiPage !== undefined) { // Wikilink
                const page = g.wikiPage.trim();
                const alias = g.wikiAlias?.trim();
                nodes.push({ type: 'text', text: alias || page, metadata: { link: page, linkType: 'internal', wikilink: true } as TextMetadata });
            } else if (g.refText !== undefined) { // Explicit/collapsed reference link or image: [text][ref] / [text][]
                const isImage = g.refBang === '!';
                const label = g.refText;
                const refId = (g.refId || label).trim().toLowerCase();
                const def = linkDefinitions.get(refId);
                if (def) {
                    nodes.push(...buildLinkOrImageNodes(isImage, label, def.url));
                } else {
                    // Not a known reference - preserve the literal bracketed text unchanged.
                    nodes.push(plainText(text.substring(match.index, match.index + match[0].length)));
                }
            } else if (g.shortText !== undefined) { // Shortcut reference: [text]
                const isImage = g.shortBang === '!';
                const label = g.shortText;
                const def = linkDefinitions.get(label.trim().toLowerCase());
                if (def) {
                    nodes.push(...buildLinkOrImageNodes(isImage, label, def.url));
                } else {
                    // Not a known reference - ordinary bracketed prose, preserve unchanged.
                    nodes.push(plainText(`${g.shortBang}[${label}]`));
                }
            } else if (g.autolinkUrl !== undefined) { // <url> autolink
                // Its references decoded in the target too, as in the text (CommonMark reads them in URLs):
                // left in the target, `&amp;` was written back as `&amp;amp;`.
                const url = decodeCharacterReferences(g.autolinkUrl);
                nodes.push({ type: 'text', text: url, formatting: Object.keys(currentFormatting).length > 0 ? { ...currentFormatting } : undefined, metadata: { link: url, linkType: 'external' } as TextMetadata });
            } else if (g.autolinkEmail !== undefined) { // <address> autolink: a mail link (CommonMark)
                nodes.push({ type: 'text', text: g.autolinkEmail, formatting: Object.keys(currentFormatting).length > 0 ? { ...currentFormatting } : undefined, metadata: { link: `mailto:${g.autolinkEmail}`, linkType: 'external' } as TextMetadata });
            } else if (g.mathDisplay !== undefined) { // Display math inside the text
                // As KaTeX, MathJax, GitLab and Pandoc read it: display math, wherever it is written.
                if (!g.mathDisplay.trim()) nodes.push(plainText(match[0]));
                else nodes.push({ type: 'code', text: g.mathDisplay.trim(), metadata: { math: displayMath ? 'block' : 'inline' } as CodeMetadata });
            } else if (g.mathInline !== undefined) { // Inline math
                nodes.push({ type: 'code', text: g.mathInline, metadata: { math: 'inline' } as CodeMetadata });
            } else if (comment) { // Inline source comment: a hidden note, kept verbatim
                nodes.push({ type: 'comment', text: comment.body, metadata: { sourceSyntax: 'html' } as CommentMetadata });
            }

            lastIndex = next;
        }

        if (lastIndex < text.length) {
            nodes.push(plainText(text.substring(lastIndex)));
        }

        return applyAbbreviations(nodes);
    };

    const escapeRegExpChars = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

    // Splits abbreviation occurrences out of plain text nodes so they carry
    // TextMetadata.abbreviationTitle, rendered as <abbr title> in HTML/editor output.
    const applyAbbreviations = (nodes: OfficeContentNode[]): OfficeContentNode[] => {
        if (abbreviationDefinitions.size === 0) return nodes;
        const pattern = new RegExp(`\\b(${[...abbreviationDefinitions.keys()].map(escapeRegExpChars).join('|')})\\b`, 'g');

        const result: OfficeContentNode[] = [];
        for (const node of nodes) {
            if (node.type !== 'text' || !node.text || node.metadata) {
                result.push(node);
                continue;
            }

            let lastIndex = 0;
            let match: RegExpExecArray | null;
            let matched = false;
            pattern.lastIndex = 0;
            while ((match = pattern.exec(node.text)) !== null) {
                matched = true;
                if (match.index > lastIndex) {
                    result.push({ type: 'text', text: node.text.substring(lastIndex, match.index), formatting: node.formatting });
                }
                result.push({
                    type: 'text',
                    text: match[0],
                    formatting: node.formatting,
                    metadata: { abbreviationTitle: abbreviationDefinitions.get(match[0]) } as TextMetadata
                });
                lastIndex = pattern.lastIndex;
            }

            if (!matched) {
                result.push(node);
                continue;
            }
            if (lastIndex < node.text.length) {
                result.push({ type: 'text', text: node.text.substring(lastIndex), formatting: node.formatting });
            }
        }
        return result;
    };

    // Splits a paragraph-shaped block's internal lines into inline-parsed content,
    // inserting a real 'break' node for a hard line break (a line ending in 2+ trailing
    // spaces or a trailing backslash) instead of collapsing it to a space. A plain single
    // newline with no such marker is still a soft break and collapses to a space,
    // unchanged from before - CommonMark itself renders a soft break as a space/newline.
    const splitParagraphLines = (block: string): OfficeContentNode[] => {
        // A continuation line's leading spaces and tabs are not part of the text (CommonMark); a
        // comment's lines, joined first, keep theirs.
        const lines = joinCodeSpanLines(joinDisplayMathLines(joinCommentLines(block.split('\n'))
            .map((line, i) => (i === 0 ? line : trimStartChars(line, ' \t')))));
        const children: OfficeContentNode[] = [];
        lines.forEach((line, i) => {
            // Two or more trailing spaces, or a trailing backslash that is not itself escaped (a
            // line ending in an escaped `\\` has no break), found from the end: an end-anchored
            // pattern retried every run of spaces in the line, quadratic in a long one.
            const spaces = line.length - trimEndChars(line, ' ').length;
            // On the last line a backslash has no line end to break, and is text.
            const backslash = spaces === 0 && i < lines.length - 1 && line.endsWith('\\') && !isEscapedAt(line, line.length - 1);
            const hardBreak = spaces >= 2 || backslash;
            appendAll(children, parseInline(spaces >= 2 ? line.slice(0, -spaces) : backslash ? line.slice(0, -1) : line, {}, true));
            if (i < lines.length - 1) {
                if (hardBreak) {
                    children.push({ type: 'break', metadata: { breakType: 'carriageReturn' } as BreakMetadata });
                } else {
                    children.push({ type: 'text', text: ' ' });
                }
            }
        });
        return children;
    };

    // A `$$` opened on one line of a paragraph and closed on a later one is one display equation:
    // those lines are rejoined so the inline tokenizer sees the whole of it. An unescaped `$$` count
    // that stays odd to the end of the paragraph leaves the lines as they are (literal text).
    const joinDisplayMathLines = (lines: string[]): string[] => {
        // Counted outside code spans, where a `$$` is literal.
        const opens = (line: string) => ((line.replace(/(`+)[^`]*?\1/g, '').match(/(?<!\\)\$\$/g) || []).length % 2) === 1;
        const out: string[] = [];
        for (let i = 0; i < lines.length; i++) {
            if (!opens(lines[i])) { out.push(lines[i]); continue; }
            let j = i + 1;
            while (j < lines.length && !opens(lines[j])) j++;
            if (j >= lines.length) { out.push(lines[i]); continue; }
            out.push(lines.slice(i, j + 1).join('\n'));
            i = j;
        }
        return out;
    };

    // A paragraph's children with its display math lifted out as blocks: the text before it, the
    // equation on its own, the text after it (as the LaTeX parser treats display math, which sits in
    // its paragraph in TeX and is its own block in the AST). Whitespace at each cut is trimmed, and a
    // part left with nothing but whitespace or breaks is dropped.
    const splitAtDisplayMath = (children: OfficeContentNode[], makeParagraph: (children: OfficeContentNode[]) => OfficeContentNode): OfficeContentNode[] => {
        if (!children.some(c => c.type === 'code' && (c.metadata as CodeMetadata)?.math === 'block')) return [makeParagraph(children)];
        const out: OfficeContentNode[] = [];
        let run: OfficeContentNode[] = [];
        const isBlank = (n: OfficeContentNode) => n.type === 'break' || (n.type === 'text' && !trimAsciiWhitespace(n.text ?? '') && !n.notes?.length && !n.comments?.length);
        const flush = () => {
            while (run.length && isBlank(run[0])) run.shift();
            while (run.length && isBlank(run[run.length - 1])) run.pop();
            if (run.length) {
                const first = run[0], last = run[run.length - 1];
                if (first.type === 'text') run[0] = { ...first, text: trimStartChars(first.text ?? '', ASCII_WHITESPACE) };
                if (last.type === 'text') run[run.length - 1] = { ...run[run.length - 1], text: trimEndChars(run[run.length - 1].text ?? '', ASCII_WHITESPACE) };
                out.push(makeParagraph(run));
            }
            run = [];
        };
        for (const child of children) {
            if (child.type === 'code' && (child.metadata as CodeMetadata)?.math === 'block') { flush(); out.push(child); }
            else run.push(child);
        }
        flush();
        return out;
    };

    // Builds an admonition node from its raw body text, read as blocks (paragraphs, lists, headings,
    // tables, code), as the document is.
    const buildAdmonitionNode = async (admonitionType: AdmonitionMetadata['admonitionType'], body: string, sourceSyntax: 'github' | 'gitlab'): Promise<OfficeContentNode> => {
        // A fenced block in the body (dequoted, so the top-level pass could not see it) is a code child.
        const children: OfficeContentNode[] = [];
        await parseBlocks(splitIntoBlocks(liftCode(body)), children);
        foldAnchorPlaceholders(children);
        return {
            type: 'admonition',
            metadata: { admonitionType, sourceSyntax } as AdmonitionMetadata,
            children
        };
    };

    // Fold standalone anchor placeholders into the anchorIds of the next node that can hold them (one
    // the generator writes anchors for: a comment or embed is passed over) so a bookmark
    // target emitted on its own line round-trips as a real anchor. Anchors no node follows are an
    // empty paragraph holding them, where they stand, as the generator writes a bookmark on an empty
    // paragraph (they joined the node before them, which could not always hold them).
    const foldAnchorPlaceholders = (content: OfficeContentNode[]): void => {
        if (!content.some(n => (n.type as any) === ANCHOR_PLACEHOLDER)) return;
        const merged: OfficeContentNode[] = [];
        let carried: string[] = [];
        for (const node of content) {
            if ((node.type as any) === ANCHOR_PLACEHOLDER) {
                carried.push(...(((node.metadata as any)?.anchorIds as string[]) || []));
                continue;
            }
            if (carried.length > 0 && ANCHOR_HOLDERS.has(node.type)) {
                const meta: any = node.metadata || (node.metadata = {} as any);
                meta.anchorIds = [...carried, ...((meta.anchorIds as string[]) || [])];
                carried = [];
            }
            merged.push(node);
        }
        if (carried.length > 0) merged.push({ type: 'paragraph', metadata: { anchorIds: carried } as any, children: [] });
        content.length = 0;
        appendAll(content, merged);
    };

    // A text's blocks: split at blank lines, then at headings, lists and HTML divs that start
    // without one. Used for the document and, recursively, for what a quote or note holds.
    // Markdown found in an HTML table's cells between blank lines (see splitIntoBlocks), each standing
    // in the table's HTML for a token (private-use characters and a per-document nonce, so no text
    // the document holds can be one) and read as Markdown into its cell after the HTML is.
    const markdownPieces: string[] = [];
    const pieceNonce = Math.random().toString(36).slice(2);
    const PIECE_TOKEN = new RegExp(`\uE000${pieceNonce}:(\\d+)\uE001`, 'g');
    const isInlineNode = (node: OfficeContentNode) => node.type === 'text' || node.type === 'image' || node.type === 'break'
        || (node.type === 'code' && (node.metadata as CodeMetadata | undefined)?.math === 'inline');
    const restoreMarkdownPieces = async (nodes: OfficeContentNode[]): Promise<OfficeContentNode[]> => {
        let changed = false;
        const out: OfficeContentNode[] = [];
        for (const node of nodes) {
            if (node.children?.length) {
                const restored = await restoreMarkdownPieces(node.children);
                if (restored !== node.children) {
                    node.children = restored;
                    // A paragraph a piece's blocks landed in is those blocks, and its runs of text.
                    if (node.type === 'paragraph' && restored.some(child => !isInlineNode(child))) {
                        changed = true;
                        let run: OfficeContentNode[] = [];
                        const flush = () => { if (run.some(child => child.type !== 'text' || child.text?.trim())) out.push({ ...node, children: run }); run = []; };
                        for (const child of restored) {
                            if (isInlineNode(child)) run.push(child);
                            else { flush(); out.push(child); }
                        }
                        flush();
                        continue;
                    }
                }
            }
            if (node.type !== 'text' || !node.text?.includes('\uE000')) {
                out.push(node);
                continue;
            }
            // Each piece is a block of its own, as blank lines around it make it (a paragraph beside the
            // text of the cell): two pieces ran together as one line.
            changed = true;
            let last = 0;
            for (const match of node.text.matchAll(PIECE_TOKEN)) {
                const before = node.text.slice(last, match.index);
                if (before.trim()) out.push({ ...node, text: before });
                const parsed: OfficeContentNode[] = [];
                await parseBlocks(splitIntoBlocks(markdownPieces[Number(match[1])] ?? ''), parsed);
                foldAnchorPlaceholders(parsed);
                appendAll(out, parsed);
                last = match.index! + match[0].length;
            }
            const after = node.text.slice(last);
            if (after.trim()) out.push({ ...node, text: after });
        }
        return changed ? out : nodes;
    };
    // A piece that landed where no Markdown is read (the text of a `<pre>`, an attribute such as alt
    // text) is that text, as written: its token was left there.
    const restoreRawPieces = (nodes: OfficeContentNode[]): void => {
        const raw = (value: string) => value.replace(PIECE_TOKEN, (_, index: string) => markdownPieces[Number(index)] ?? '');
        for (const node of nodes) {
            if (typeof node.text === 'string' && node.text.includes('\uE000')) node.text = raw(node.text);
            for (const record of [node.metadata as Record<string, unknown> | undefined, (node as any).htmlAttributes as Record<string, unknown> | undefined]) {
                if (!record) continue;
                for (const [key, value] of Object.entries(record)) if (typeof value === 'string' && value.includes('\uE000')) record[key] = raw(value);
            }
            if (node.children?.length) restoreRawPieces(node.children);
            if (node.notes?.length) restoreRawPieces(node.notes);
        }
    };

    // An HTML block, read with the HTML parser as a renderer shows it: its Markdown pieces restored,
    // its pictures joining this document's attachments, and its footnote references taking their
    // `[^id]:` definitions from this document.
    // Its comments are kept, as Markdown's are (the HTML parser drops them unless asked, and a
    // comment in an HTML table or block was lost).
    const readHtmlBlock = async (block: string): Promise<OfficeParserAST> => {
        const html = await parseHtml(Buffer.from(block), { ...config, htmlParserConfig: { ...config.htmlParserConfig, preserveComments: true } });
        if (markdownPieces.length) {
            html.content = await restoreMarkdownPieces(html.content);
            restoreRawPieces(html.content);
        }
        const renamed = new Map<string, string>();
        for (const attachment of html.attachments) {
            const name = `image_${attachments.length + 1}.${attachment.extension || 'png'}`;
            renamed.set(attachment.name, name);
            attachments.push({ ...attachment, name });
        }
        const adopt = (nodes: OfficeContentNode[]) => {
            for (const node of nodes) {
                const image = node.type === 'image' ? node.metadata as ImageMetadata | undefined : undefined;
                if (image?.attachmentName && renamed.has(image.attachmentName)) image.attachmentName = renamed.get(image.attachmentName)!;
                if (node.notes?.length) {
                    node.notes = node.notes.map(note => {
                        const noteId = (note.metadata as any)?.noteId;
                        return (noteId !== undefined && footnoteDefinitions.has(noteId) ? footnoteNode(noteId) : null) ?? note;
                    });
                }
                if (node.children) adopt(node.children);
            }
        };
        adopt(html.content);
        return html;
    };

    const splitIntoBlocks = (text: string): string[] => {
    // A line of only spaces and tabs is a blank line too (CommonMark). Each block keeps the blank
    // lines before it, which an indented code block holds (see below).
    const parts = text.split(/(\n(?:[ \t]*\n)+)/);
    const rawBlocks: string[] = [];
    const rawSeparators: string[] = [];
    // An HTML table a blank line runs through (`<table>`, rows, a blank line, more rows) is one
    // table: a renderer passes each piece through as HTML, and a browser builds one table of them.
    // Pieces join while a table the first piece opened is open and each piece starts with a tag, up
    // to MAX_JOINED_TABLE_DEPTH tables deep: pieces that only open tables, never closing one, would
    // otherwise join into one block nested past what the HTML parser reads. Where the table closes
    // within MAX_TABLE_PIECES pieces, pieces of Markdown between (the GitHub way of writing a cell:
    // `<td>`, a blank line, `Some *text*`, a blank line, `</td>`) join it too, read as Markdown into
    // their cell (see restoreMarkdownPieces); they were a paragraph, and the rest escaped text.
    let tableDepth = 0;
    let joinMarkdownUntil = -1;
    for (let i = 0; i < parts.length; i += 2) {
        const part = parts[i];
        const separator = i > 0 ? parts[i - 1] : '';
        const startsWithTag = /^[ \t]*</.test(part);
        if (tableDepth > 0 && tableDepth <= MAX_JOINED_TABLE_DEPTH && (startsWithTag || i <= joinMarkdownUntil)) {
            const markdown = !startsWithTag && !TABLE_PART_TAG.test(part);
            if (markdown) markdownPieces.push(part);
            rawBlocks[rawBlocks.length - 1] += separator + (markdown ? `\uE000${pieceNonce}:${markdownPieces.length - 1}\uE001` : part);
            tableDepth += markdown ? 0 : htmlTableDepthChange(part);
            continue;
        }
        rawBlocks.push(part);
        rawSeparators.push(separator);
        tableDepth = HTML_BLOCK_LINE.exec(part)?.index === 0 ? htmlTableDepthChange(part) : 0;
        joinMarkdownUntil = -1;
        for (let j = i + 2, depth = tableDepth, n = 0; tableDepth > 0 && j < parts.length && n < MAX_TABLE_PIECES; j += 2, n++) {
            depth += htmlTableDepthChange(parts[j]);
            if (depth <= 0) { joinMarkdownUntil = j; break; }
        }
    }
    const blocks: string[] = [];
    const separators: string[] = [];

    // Sub-split blocks that contain headings or lists without double newlines
    for (let r = 0; r < rawBlocks.length; r++) {
        const rawBlock = rawBlocks[r];
        if (!trimAsciiWhitespace(rawBlock)) continue;
        const firstBlock = blocks.length;

        // Match headings or lists that might be joined with other text via single newline
        const lines = rawBlock.split('\n');
        let currentSubBlock: string[] = [];
        // Tracks whether we're currently "inside" a list (a list-item line, or an
        // indented continuation line right after one) so a continuation line doesn't
        // itself get treated as the boundary that splits the list into a new block -
        // see the "Lists" block dispatch below, which merges such a line into the
        // previous item's content instead of dropping it.
        let inList: boolean = false;
        // Whether the sub-block's second line is a table's delimiter row, found when that line is added:
        // tested before every line, a long second line was read again for each line after it.
        let inTable = false;
        const add = (line: string) => {
            currentSubBlock.push(line);
            if (currentSubBlock.length === 2) inTable = line.includes('-') && /^[-:| \t]+$/.test(line);
        };
        const flush = () => {
            if (currentSubBlock.length > 0) {
                separators.push(blocks.length === firstBlock ? rawSeparators[r] : '\n');
                blocks.push(currentSubBlock.join('\n'));
            }
            currentSubBlock = [];
            inTable = false;
        };
        for (const line of lines) {
            checkAbortSignal(config.abortSignal);
            // Paragraph text is open when the sub-block holds lines that are neither a list's nor a
            // table's: what follows it may continue it rather than start a block (CommonMark).
            const inParagraph: boolean = currentSubBlock.length > 0 && !inList && !inTable;
            // A line of `=` or `-` under paragraph text is a setext heading's underline; otherwise
            // three or more `-`, `*` or `_` (spaces between allowed) are a thematic break, which
            // is never a list item.
            // (Not under a quote's paragraph: an underline cannot continue a quote lazily.)
            const isSetextUnderline: boolean = inParagraph && !/^ {0,3}>/.test(currentSubBlock[0]) && SETEXT_UNDERLINE.test(line);
            const isRule: boolean = !isSetextUnderline && THEMATIC_BREAK.test(line);
            const isHeading: boolean = !isSetextUnderline && ATX_HEADING_START.test(line);
            const item: RegExpExecArray | null = isSetextUnderline || isRule ? null : LIST_ITEM_START.exec(line);
            const itemIndent: number = item ? item[1].replace(/\t/g, '    ').length : 0;
            // An item starts a list where a block may start (four columns of indentation make that
            // a code block, unless the line is nested in a list), and interrupts a paragraph only if
            // it has content and, when ordered, starts at 1.
            const isList: boolean = !!item && (inList || itemIndent < 4)
                && !(inParagraph && (item[3] === undefined || (/\d/.test(item[2]) && parseInt(item[2], 10) !== 1)));
            const isHtmlTag = !!line.match(/^<\/?div[^>]*>$/i);
            // A non-list, non-blank, indented (>=2 columns or a tab) line encountered
            // while already inside a list is a continuation of the current item, not a
            // new construct. Scoped to a single such line at a time (no nested
            // code/blockquote/sub-list/multi-paragraph items - those require un-splitting
            // already-separated raw blocks, out of scope here).
            const isContinuation: boolean = !isList && inList && /^(?: {2,}|\t)/.test(line) && trimAsciiWhitespace(line).length > 0;
            const staysInListMode: boolean = isList || isContinuation;

            if (isSetextUnderline) {
                // The underline ends the heading's block, so the heading is its last line.
                add(line);
                flush();
                inList = false;
                continue;
            }

            // Split if:
            // 1. Current line is a heading or a thematic break
            // 2. Current line enters or leaves "list mode" relative to the previous line
            // 3. Current line is an HTML tag (div)
            if (isHeading || isRule || isHtmlTag || staysInListMode !== inList) flush();

            add(line);
            inList = staysInListMode;

            // Headings, thematic breaks and HTML tags are single-line blocks for our state machine
            if (isHeading || isRule || isHtmlTag) {
                flush();
                inList = false;
            }
        }
        flush();
    }

    // Re-join a list block that a blank line tore away from its parent. The generator's older
    // loose output (`- a\n\n\n    - a1`) and foreign editors both split a nested item into its
    // own block; left apart, the child's leading indent is stripped by the per-block `trim()`
    // below and it reparses as a flat top-level item under a fresh listId. Merge a block back
    // into the preceding one only when the previous block is itself a list (its FIRST line is a
    // marker - the sub-splitter guarantees such a block holds only marker/continuation lines) and
    // the current block OPENS with an INDENTED marker. An unindented `- b` after a blank line is
    // deliberately left split (a flat loose list keeps its own listId), and anything that is not
    // an indented marker (continuation text, indented code, placeholders) never triggers a merge.
    // (Two spaces, then any more: ` {2,}[ \t]*` tried every split of a long run of them.)
    const indentedMarkerStart = /^(?: {2}|\t)[ \t]*(?:[-*+]|\d+[.)])(?:[ \t]|$)/;
    const mergedBlocks: string[] = [];
    // Whether the last merged block is an indented code block, and a list: a list takes an indented
    // marker after it, and an indented code block the one after it (see below).
    let previousIsList = false;
    let previousIsCode = false;
    // The content columns of the last list's items (see below), and them in order, sorted when needed.
    const itemColumns = new Set<number>();
    let sortedColumns: number[] | undefined;
    for (let i = 0; i < blocks.length; i++) {
        let block = blocks[i];
        if (previousIsList && indentedMarkerStart.test(block.split('\n', 1)[0])) {
            mergedBlocks[mergedBlocks.length - 1] += `\n${block}`;
            addItemContentColumns(block, itemColumns);
            sortedColumns = undefined;
            continue;
        }
        // A block after a list, indented to one of its items' content columns, is that item's content
        // (a paragraph, a table, a code block indented four more): the AST's items hold one line, so it
        // is read as a block of its own after the list, without the item's indentation. Read as
        // indented code, a table or paragraph under an item came out as a code block.
        if (itemColumns.size) {
            const indent = indentColumn(block);
            sortedColumns ??= [...itemColumns].sort((a, b) => a - b);
            let column = 0;
            for (let lo = 0, hi = sortedColumns.length - 1; lo <= hi;) {
                const mid = (lo + hi) >> 1;
                if (sortedColumns[mid] <= indent) { column = sortedColumns[mid]; lo = mid + 1; } else hi = mid - 1;
            }
            if (column > 0) block = block.split('\n').map(line => dedentLine(line, column)).join('\n');
            else itemColumns.clear();
        }
        const isCode = isIndentedCodeBlock(block);
        // An indented code block runs on across blank lines (CommonMark): the next indented block
        // after one is part of it, with the blank lines between.
        if (isCode && previousIsCode && separators[i] !== '\n') {
            mergedBlocks[mergedBlocks.length - 1] += separators[i] + block;
            continue;
        }
        mergedBlocks.push(block);
        const continuation = block !== blocks[i];
        previousIsList = !continuation && LIST_ITEM_START.test(block.split('\n', 1)[0]);
        previousIsCode = isCode;
        if (previousIsList) {
            itemColumns.clear();
            addItemContentColumns(block, itemColumns);
            sortedColumns = undefined;
        } else if (!continuation) {
            itemColumns.clear();
        }
    }
    return mergedBlocks;
    };

    let listIdCounter = 1;

    // Parses blocks (see splitIntoBlocks) into nodes appended to `content`: the document's, or those
    // of a quote or note being read.
    const parseBlocks = async (blocks: string[], content: OfficeContentNode[]): Promise<void> => {
    let currentAlignment: 'left' | 'center' | 'right' | 'justify' | undefined = undefined;

    for (let block of blocks) {
        // Empty anchors (bookmark targets) before a block's content, on its first line or on a line of
        // their own, as the generator writes them before a paragraph, image or table: they are the
        // anchor ids of the node the rest of the block becomes (folded in below), not visible text.
        // A block of nothing but anchors is handled further down.
        const leadingAnchors = /^\s*(?:<a\s[^>]*>\s*<\/a>[ \t]*)+\n?/i.exec(block);
        // Those on the line of the block's first content (not a line of their own), which a definition
        // list gives its first term (see there).
        let firstLineAnchors: OfficeContentNode | undefined;
        if (leadingAnchors && trimAsciiWhitespace(block.slice(leadingAnchors[0].length))) {
            const anchorIds = [...leadingAnchors[0].matchAll(/<a\s[^>]*\b(?:name|id)="([^"]*)"/gi)].map(m => m[1]).filter(Boolean);
            if (anchorIds.length > 0) {
                const placeholder = { type: ANCHOR_PLACEHOLDER as any, metadata: { anchorIds } as any, children: [] };
                content.push(placeholder);
                if (!leadingAnchors[0].endsWith('\n')) firstLineAnchors = placeholder;
                block = block.slice(leadingAnchors[0].length);
            }
        }
        // Preserved before the generic trim() below, which strips the first line's
        // leading indentation - the indented-code-block check further down needs every
        // line's original indentation, including the first.
        const untrimmedBlock = block;
        block = trimAsciiWhitespace(block);
        if (!block) continue;

        // Standalone anchor-only block: one or more empty `<a name|id="…"></a>` tags on their
        // own line (bookmark targets the MarkdownGenerator emits just before a heading/paragraph).
        // Capture them as a placeholder so the post-loop pass can re-attach them to the following
        // node's anchorIds; otherwise the tag-opening `<` is escaped and they render as visible text.
        const anchorsOnly = /<a\s/i.test(block) ? emptyAnchorRun(block) : null;
        if (anchorsOnly && !trimAsciiWhitespace(block.slice(anchorsOnly.end))) {
            if (anchorsOnly.ids.length > 0) {
                content.push({ type: ANCHOR_PLACEHOLDER as any, metadata: { anchorIds: anchorsOnly.ids } as any, children: [] });
                continue;
            }
        }

        // Check for alignment wrapper start/end
        const alignStartMatch = block.match(/^<div\s+(?:style="text-align:\s*(left|center|right|justify);?"|align="(left|center|right|justify)")>$/i);
        if (alignStartMatch) {
            currentAlignment = (alignStartMatch[1] || alignStartMatch[2]).toLowerCase() as any;
            continue;
        }
        if (block.match(/^<\/div>$/i)) {
            currentAlignment = undefined;
            continue;
        }

        let alignment = currentAlignment;
        // Check for single-line alignment wrapper (for compatibility)
        const alignMatch = block.slice(-6).toLowerCase() === '</div>'
            ? /^<div\s+(?:style="text-align:\s*(left|center|right|justify);?"|align="(left|center|right|justify)")>/i.exec(block)
            : null;
        if (alignMatch) {
            alignment = (alignMatch[1] || alignMatch[2]).toLowerCase() as any;
            block = trimAsciiWhitespace(block.slice(alignMatch[0].length, -6));
        }

        // Embed leaf directive (remark-directive family): `::youtube[Label]{id=... width=... align=...}`
        // or `::embed[Label]{src=... width=... height=... align=...}`. Only these two names are
        // recognised; any other `::name` stays literal text (no catch-all). `::youtube` renders from
        // a validated id via a fixed template, so it is unconditional; `::embed` carries an arbitrary
        // src, so it is gated behind `preserveIframes` (the trust input) exactly like a raw <iframe>,
        // and stays literal text otherwise. New input only; nothing that parsed before changes.
        const embedDirectiveMatch = block.match(/^::(youtube|embed)(?:\[((?:[^\]\\\n]|\\.)*)\])?\{([^}]*)\}$/);
        if (embedDirectiveMatch) {
            const kind = embedDirectiveMatch[1];
            // The label is Markdown text, decoded as the generator escapes it.
            const label = decodeMarkdownText((embedDirectiveMatch[2] || '').trim()) || undefined;
            const attrs = parseEmbedDirectiveAttrs(embedDirectiveMatch[3]);
            if (kind === 'youtube' && attrs.id) {
                const embedUrl = `https://www.youtube.com/watch?v=${attrs.id}`;
                content.push({
                    type: 'embed',
                    text: embedUrl,
                    metadata: { embedType: 'youtube', videoId: attrs.id, url: embedUrl, width: attrs.width, align: attrs.align, label } as EmbedMetadata
                });
                continue;
            }
            if (kind === 'embed' && attrs.src && iframeAllowed(attrs.src, config.htmlParserConfig?.preserveIframes)) {
                content.push({
                    type: 'embed',
                    text: attrs.src,
                    metadata: { embedType: 'iframe', url: attrs.src, width: attrs.width, height: attrs.height, align: attrs.align, label } as EmbedMetadata
                });
                continue;
            }
            // Recognised name but not a usable/allowed directive: fall through so the line becomes
            // ordinary text rather than being dropped.
        }

        // Ambiguous "folk" embed forms, imported only under the opt-in (embedFolkForms), since
        // auto-upgrading an image/link to an embed is a heuristic that could mangle a genuine image
        // link. Both become a safe youtube embed (rendered from the validated id). A standalone line
        // only; anything not matching falls through to ordinary image/link parsing.
        if (config.htmlParserConfig?.embedFolkForms) {
            // Clickable thumbnail: [![alt](thumb)](watch), youtube when either URL is a youtube link.
            const thumbMatch = block.match(/^\[!\[((?:[^\]\\\n]|\\.)*)\]\(([^)\s]+)\)\]\(([^)\s]+)\)$/);
            if (thumbMatch) {
                const fid = extractYoutubeId(thumbMatch[2]) || extractYoutubeId(thumbMatch[3]);
                if (fid) {
                    const embedUrl = `https://www.youtube.com/watch?v=${fid}`;
                    content.push({ type: 'embed', text: embedUrl, metadata: { embedType: 'youtube', videoId: fid, url: embedUrl, label: decodeMarkdownText(thumbMatch[1].trim()) || undefined } as EmbedMetadata });
                    continue;
                }
            }
            // Obsidian-style: a standalone image whose URL is a youtube link.
            const obsMatch = block.match(/^!\[((?:[^\]\\\n]|\\.)*)\]\(([^)\s]+)\)$/);
            if (obsMatch) {
                const fid = extractYoutubeId(obsMatch[2]);
                if (fid) {
                    const embedUrl = `https://www.youtube.com/watch?v=${fid}`;
                    content.push({ type: 'embed', text: embedUrl, metadata: { embedType: 'youtube', videoId: fid, url: embedUrl, label: decodeMarkdownText(obsMatch[1].trim()) || undefined } as EmbedMetadata });
                    continue;
                }
            }
        }

        // YouTube embed fallback: MarkdownGenerator's 'embed' case emits a single-line
        // <div data-youtube-video="ID" data-width="…" data-align="…"></div> when fallbackToHtml
        // is on; recognise it here so a saved-then-reopened .md keeps the video.
        const youtubeMatch = block.match(/^<div\s+data-youtube-video="([^"]*)"([^>]*)>\s*<\/div>$/i);
        if (youtubeMatch) {
            const videoId = youtubeMatch[1];
            const attrsStr = youtubeMatch[2];
            const widthMatch = attrsStr.match(/data-width="([^"]*)"/i);
            const youtubeAlignMatch = attrsStr.match(/data-align="([^"]*)"/i);
            const youtubeLabelMatch = attrsStr.match(/data-embed-label="([^"]*)"/i);
            const embedAlign = youtubeAlignMatch && (['left', 'center', 'right'] as const).includes(youtubeAlignMatch[1] as any) ? youtubeAlignMatch[1] as 'left' | 'center' | 'right' : undefined;
            const embedUrl = videoId ? `https://www.youtube.com/watch?v=${videoId}` : undefined;
            content.push({
                type: 'embed',
                // Childless nodes need .text so generic AST consumers (text/chunking generators)
                // don't silently drop them.
                text: embedUrl,
                metadata: {
                    embedType: 'youtube',
                    videoId,
                    url: embedUrl,
                    // Attribute values, HTML-escaped by the generator.
                    width: widthMatch ? decodeCharacterReferences(widthMatch[1]) : undefined,
                    align: embedAlign,
                    label: youtubeLabelMatch ? decodeCharacterReferences(youtubeLabelMatch[1]) : undefined
                } as EmbedMetadata
            });
            continue;
        }

        // Generic iframe fallback: MarkdownGenerator's 'embed' case emits a single-line
        // <iframe src="…"></iframe> for a preserved iframe when fallbackToHtml is on. Recognise
        // it only when the caller opted into iframe preservation, so default parsing is unchanged.
        const iframeMatch = block.match(/^<iframe(\s[^>]*?)?\/?>(?:\s*<\/iframe>)?$/i);
        if (iframeMatch) {
            const attrsStr = iframeMatch[1] ?? '';
            // The emitted <iframe> HTML-escapes its attribute values (sanitizeUrl -> escapeHtml), so
            // decode them back; otherwise the src double-escapes (`&amp;` -> `&amp;amp;`) and its
            // query string is corrupted a little more on every save/reload cycle. One pass, the exact
            // inverse of that escaping, so a genuinely double-escaped value unwinds one level per parse.
            const decodeAttr = (s: string | undefined) => decodeCharacterReferences(s || '');
            const src = decodeAttr(attrsStr.match(/\bsrc="([^"]*)"/i)?.[1]);
            const width = attrsStr.match(/\bwidth="([^"]*)"/i)?.[1];
            const height = attrsStr.match(/\bheight="([^"]*)"/i)?.[1];
            // A YouTube iframe is recognised the same way the HTML parser does it (host + id
            // capture), BEFORE and INDEPENDENT of the preserveIframes gate, so the same iframe
            // yields the same 'youtube' embed whichever parser sees it. Only a generic (non-YouTube)
            // iframe is gated behind preserveIframes and kept as an 'iframe' embed.
            const ytMatch = src && /youtube(?:-nocookie)?\.com/.test(src) ? src.match(/(?:embed\/|v=)([^&?/\s]+)/) : null;
            if (ytMatch) {
                const embedUrl = `https://www.youtube.com/watch?v=${ytMatch[1]}`;
                content.push({
                    type: 'embed',
                    text: embedUrl,
                    metadata: {
                        embedType: 'youtube',
                        videoId: ytMatch[1],
                        url: embedUrl,
                        width: width !== undefined ? decodeAttr(width) : undefined,
                        height: height !== undefined ? decodeAttr(height) : undefined
                    } as EmbedMetadata
                });
                continue;
            }
            if (src && iframeAllowed(src, config.htmlParserConfig?.preserveIframes)) {
                content.push({
                    type: 'embed',
                    text: src,
                    metadata: {
                        embedType: 'iframe',
                        url: src,
                        width: width !== undefined ? decodeAttr(width) : undefined,
                        height: height !== undefined ? decodeAttr(height) : undefined
                    } as EmbedMetadata
                });
                continue;
            }
        }

        // Source comment on its own lines, extracted to a placeholder above: a hidden note, kept verbatim
        const commentIndex = placeholderIndex(block, 'HTML_COMMENT');
        if (commentIndex !== null && commentIndex < htmlComments.length) {
            content.push({
                type: 'comment',
                text: htmlComments[commentIndex],
                metadata: { sourceSyntax: 'html' } as CommentMetadata
            });
            continue;
        }

        // Code Block
        const codeIndex = placeholderIndex(block, 'CODE_BLOCK');
        if (codeIndex !== null && codeIndex < codeBlocks.length) {
            const data = JSON.parse(codeBlocks[codeIndex]);
            content.push({
                type: 'code',
                text: data.code,
                metadata: { language: data.lang } as CodeMetadata
            });
            continue;
        }

        // GLFM-style fenced-div admonition, extracted to a placeholder above
        const admonitionIndex = placeholderIndex(block, 'ADMONITION');
        if (admonitionIndex !== null && admonitionIndex < admonitionBlocks.length) {
            const data = JSON.parse(admonitionBlocks[admonitionIndex]);
            content.push(await buildAdmonitionNode(data.admonitionType, data.body, 'gitlab'));
            continue;
        }

        // Block math ($$...$$), extracted to a placeholder above
        const mathIndex = placeholderIndex(block, 'MATH_BLOCK');
        if (mathIndex !== null && mathIndex < mathBlocks.length) {
            content.push({
                type: 'code',
                text: mathBlocks[mathIndex],
                metadata: { math: 'block' } as CodeMetadata
            });
            continue;
        }

        // Indented code block (4-space or tab indent on every non-blank line), before any other
        // construct: four columns of indentation make a code block of what would otherwise read as
        // a heading, quote, list or table (CommonMark). A nested list item indented under its
        // parent never gets here alone: the splitter keeps it in its list's block. A partially
        // indented block (some lines indented, some not) falls through to the other branches.
        {
            const codeLines = untrimmedBlock.split('\n');
            const nonBlankLines = codeLines.filter(l => trimAsciiWhitespace(l).length > 0);
            if (nonBlankLines.length > 0 && nonBlankLines.every(l => /^(?: {4}|\t)/.test(l))) {
                // Blank lines before and after the code are not part of it (CommonMark): the one
                // ending the document came through as a trailing line break.
                const codeText = codeLines.map(l => l.replace(/^(?: {4}|\t)/, ''));
                let first = 0, last = codeText.length;
                while (first < last && /^[ \t]*$/.test(codeText[first])) first++;
                while (last > first && /^[ \t]*$/.test(codeText[last - 1])) last--;
                content.push({ type: 'code', text: codeText.slice(first, last).join('\n') });
                continue;
            }
        }

        // Hr - a thematic break (horizontal rule), not a page break, so it survives a save as
        // `---` rather than collapsing to a bare newline. Three or more `-`, `*` or `_`, spaces
        // between allowed; before lists, as `* * *` and `- - -` are breaks and not list items.
        if (THEMATIC_BREAK.test(block)) {
            content.push({ type: 'break', metadata: { breakType: 'thematic' } });
            continue;
        }

        // Heading (allowing for leading HTML anchors and trailing {#anchor}); `#` alone is an empty one.
        const headingMatch = block.match(/^((?:<a[^>]*><\/a>)*)[ \t]*(#{1,6})(?:[ \t]+([\s\S]*))?$/);
        if (headingMatch) {
            const leadingAnchorsRaw = headingMatch[1];
            const { text: rawText, anchor: explicitAnchor } = splitHeadingAnchor(headingMatch[3] ?? '');

            const anchorIds: string[] = [];
            if (leadingAnchorsRaw) {
                const idMatches = leadingAnchorsRaw.matchAll(/<a\s[^>]*\b(?:name|id)="([^"]+)"/gi);
                for (const m of idMatches) anchorIds.push(m[1]);
            }
            if (explicitAnchor) anchorIds.push(explicitAnchor);

            const children = parseInline(rawText);
            content.push({
                type: 'heading',
                text: plainTextOf(children),
                metadata: {
                    level: headingMatch[2].length,
                    alignment,
                    anchorIds: anchorIds.length > 0 ? anchorIds : undefined
                } as HeadingMetadata,
                children
            });
            continue;
        }

        // Setext heading (Text\n===  or  Text\n---): a line of text immediately followed
        // by a lone `=`/`-` underline with no blank line between them. By the time a
        // block reaches this point, the sub-splitter above has already separated out any
        // genuinely blank-line-preceded thematic break into its own isolated block (which
        // has no preceding text line to combine with here), so this only fires for the
        // ambiguous "text directly above a dash/equals-only line" shape setext needs.
        // Scoped to a single line immediately above the underline becoming the heading
        // text; multi-line setext text (CommonMark's "Foo\nbar\n===" merging into one
        // heading) is an explicitly out-of-scope simplification - any earlier lines in
        // the block are pushed as a separate paragraph first.
        // (Not a quote's: an underline cannot continue a quote lazily; see the quote branch.)
        const setextMatch = block.startsWith('>') ? null : block.match(/^([\s\S]*)\n([=]+|-+)[ \t]*$/);
        if (setextMatch) {
            const lines = setextMatch[1].split('\n');
            const headingLine = lines[lines.length - 1];
            const earlierLines = trimAsciiWhitespace(lines.slice(0, -1).join('\n'));
            if (earlierLines) {
                // Split around display math as every other paragraph is, so no block equation sits in one.
                appendAll(content, splitAtDisplayMath(splitParagraphLines(earlierLines),
                    parts => ({ type: 'paragraph', metadata: { alignment } as any, children: parts })));
            }
            const children = parseInline(headingLine);
            content.push({
                type: 'heading',
                text: plainTextOf(children),
                metadata: { level: setextMatch[2][0] === '=' ? 1 : 2, alignment } as HeadingMetadata,
                children
            });
            continue;
        }

        // Blockquote: `>` starts it, a space after it optional. The markers of every nesting level
        // come off each line in one pass (a nested quote reads as part of this one; stripping one
        // level per pass made a deeply nested line quadratic), and a line without one is a lazy
        // continuation of the quote's paragraph.
        if (block.startsWith('>')) {
            // A line without a marker is a lazy continuation of the quote's paragraph: a line of `=`
            // or `-` there is text, not a heading's underline, so it is escaped as it is dequoted.
            const dequoted = block.split('\n').map(line => {
                const markers = /^(?: {0,3}>[ \t]?)+/.exec(line);
                if (markers) return line.slice(markers[0].length);
                return SETEXT_UNDERLINE.test(line) ? `\\${trimStartChars(line, ' ')}` : line;
            }).join('\n');

            // GitHub-style admonition: `> [!NOTE]` on the first quoted line.
            const admonitionHeaderMatch = dequoted.match(/^\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\][ \t]*\n?([\s\S]*)$/i);
            if (admonitionHeaderMatch) {
                const admonitionType = admonitionHeaderMatch[1].toLowerCase() as AdmonitionMetadata['admonitionType'];
                content.push(await buildAdmonitionNode(admonitionType, admonitionHeaderMatch[2], 'github'));
                continue;
            }

            // What the quote holds is read as blocks. Its paragraphs are quote paragraphs (a quote is a
            // paragraph style in the AST); a list, heading, table or code block in it cannot sit
            // inside a paragraph, so it stands between them, keeping its structure. (A fenced block
            // in it is lifted here: the top-level pass could not see it behind the markers.)
            const quoted: OfficeContentNode[] = [];
            await parseBlocks(splitIntoBlocks(liftCode(dequoted)), quoted);
            for (const node of quoted) {
                if (node.type === 'paragraph') node.metadata = { ...node.metadata, style: 'Quote' } as any;
                content.push(node);
            }
            continue;
        }

        // Definition list (Markdown Extra / Pandoc / Kramdown): a term line followed by one or more
        // ": definition" lines, any number of such groups in the block, e.g.:
        //   Term
        //   : Definition of the term.
        //   Another term
        //   : Its definition.
        {
            const lines = block.split('\n');
            const isDefinition = (line: string) => /^:[ \t]+\S/.test(line);
            let valid = lines.length >= 2;
            for (let i = 0; valid && i < lines.length; i++) {
                // A term does not start with `:`, and a definition follows it.
                if (!isDefinition(lines[i])) valid = !lines[i].startsWith(':') && i + 1 < lines.length && isDefinition(lines[i + 1]);
            }
            if (valid) {
                const children = lines.map((line): OfficeContentNode => (isDefinition(line)
                    ? { type: 'definitionDescription', children: parseInline(line.replace(/^:[ \t]+/, '')) }
                    : { type: 'definitionTerm', children: parseInline(line) }));
                // Anchors starting the first term's line are the term's, as the generator writes a term's
                // ids; the list's stand on a line of their own before it.
                if (firstLineAnchors && content[content.length - 1] === firstLineAnchors) {
                    content.pop();
                    children[0].metadata = { anchorIds: (firstLineAnchors.metadata as any).anchorIds };
                }
                content.push({ type: 'definitionList', children });
                continue;
            }
        }

        // Lists
        if (LIST_ITEM_START.test(block.split('\n', 1)[0])) {
            const lines = block.split('\n');
            let listId = `md-list-${listIdCounter++}`;
            const listCounters = new Map<number, number>();
            // The marker each level's list uses (its bullet, or an ordered list's `.` or `)`): another
            // starts a new list, as CommonMark reads it (`1. one` after `- b` counted on from the bullets).
            const levelMarkers = new Map<number, string>();
            // Relative indent stack (not a fixed-width divisor) so nesting level is
            // computed from what indentation actually appeared in this block, rather
            // than assuming a specific indent width. This makes the parser agnostic to
            // 2-space (hand-written), 4-space (this generator's own output), or
            // tab-indented (normalized to a 4-column stop) nested lists.
            const indentStack: number[] = [];
            // The most recently pushed list-item node, so a following indented
            // continuation line (see the sub-splitter above) can be merged into it
            // instead of being silently dropped.
            let lastListNode: OfficeContentNode | undefined;
            // Items that took continuation lines; their text is set once, after the block.
            const continued = new Set<OfficeContentNode>();

            for (const line of lines) {
                // A marker alone on its line is an empty item.
                const match = line.match(/^([ \t]*)([-*+]|\d+[.)])(?:[ \t]+(.*))?$/);
                if (match) {
                    const rawIndent = match[1].replace(/\t/g, '    ').length;
                    while (indentStack.length > 0 && rawIndent <= indentStack[indentStack.length - 1]) {
                        indentStack.pop();
                    }
                    const level = indentStack.length;
                    indentStack.push(rawIndent);

                    // Purge any deeper levels' counters now that we're back at this
                    // level - otherwise a nested sub-list under a later sibling item
                    // would incorrectly continue a previous sibling's child numbering
                    // instead of restarting at 0.
                    for (const key of [...listCounters.keys()]) {
                        if (key > level) listCounters.delete(key);
                    }

                    const marker = match[2];
                    const isOrdered = !!marker.match(/\d+[.)]/);
                    const listType: 'ordered' | 'unordered' = isOrdered ? 'ordered' : 'unordered';
                    const markerKind = isOrdered ? marker.slice(-1) : marker;
                    const previousKind = levelMarkers.get(level);
                    if (previousKind !== undefined && previousKind !== markerKind) {
                        listCounters.delete(level);
                        if (level === 0) listId = `md-list-${listIdCounter++}`;
                    }
                    levelMarkers.set(level, markerKind);
                    for (const key of [...levelMarkers.keys()]) if (key > level) levelMarkers.delete(key);

                    if (listCounters.get(level) === undefined) {
                        if (isOrdered) {
                            const startNum = parseInt(marker, 10);
                            listCounters.set(level, isNaN(startNum) ? 0 : startNum - 1);
                        } else {
                            listCounters.set(level, 0);
                        }
                    } else {
                        listCounters.set(level, listCounters.get(level)! + 1);
                    }

                    let itemText = match[3] ?? '';
                    let isTask: boolean | undefined;
                    let checked: boolean | undefined;
                    const taskMatch = itemText.match(/^\[([ xX])\](?:[ \t]+(.*))?$/);
                    if (taskMatch) {
                        isTask = true;
                        checked = taskMatch[1].toLowerCase() === 'x';
                        itemText = taskMatch[2] ?? '';
                    }
                    // Empty anchors right after the marker (as the generator writes an item's
                    // bookmark targets) are the item's anchor ids, not visible text.
                    const itemAnchors = /^(?:<a\s[^>]*>\s*<\/a>\s*)+/i.exec(itemText);
                    const itemAnchorIds = itemAnchors ? [...itemAnchors[0].matchAll(/<a\s[^>]*\b(?:name|id)="([^"]*)"/gi)].map(m => m[1]).filter(Boolean) : [];
                    if (itemAnchorIds.length > 0) itemText = itemText.slice(itemAnchors![0].length);

                    const children = parseInline(itemText);
                    const listNode: OfficeContentNode = {
                        type: 'list',
                        text: plainTextOf(children),
                        metadata: {
                            listType,
                            indentation: level,
                            alignment: alignment || 'left',
                            listId,
                            itemIndex: listCounters.get(level),
                            isTask,
                            checked,
                            ...(itemAnchorIds.length > 0 && { anchorIds: itemAnchorIds })
                        } as ListMetadata,
                        children
                    };
                    content.push(listNode);
                    lastListNode = listNode;
                } else if (lastListNode && trimAsciiWhitespace(line).length > 0 && /^(?: {2,}|\t)/.test(line)) {
                    // Indented continuation line: merge its inline content into the
                    // previous item rather than dropping it. Scoped to a single such
                    // line (no nested code/blockquote/sub-list/multi-paragraph items).
                    // Appended in place, and the item's text set once below: rebuilding the children
                    // and their text for every line made an item of many lines quadratic.
                    const children = lastListNode.children ?? (lastListNode.children = []);
                    children.push({ type: 'text', text: ' ' });
                    appendAll(children, parseInline(trimAsciiWhitespace(line)));
                    continued.add(lastListNode);
                }
            }
            for (const item of continued) item.text = plainTextOf(item.children ?? []);
            continue;
        }

        // An HTML block (CommonMark's): a line starting with a block-level tag (`<table>`, `<p
        // align="center">`, `<details>`, `<div>`, `<h1>`), or a `<pre>` closed in the block, starts
        // one, which runs to the blank line. It is read with the HTML parser, as a renderer shows it:
        // it was text, its tags escaped into view on the first save. `<table>` in a code span or a
        // sentence is text. The block starts at that line: text before it is Markdown (a paragraph
        // the block interrupts).
        const htmlStart = HTML_BLOCK_LINE.exec(block) ?? (/^ {0,3}<pre[\s>]/i.test(block) && /<\/pre\s*>/i.test(block) ? { index: 0 } : null);
        if (htmlStart && htmlStart.index > 0) {
            const cut = htmlStart!.index + 1;
            await parseBlocks([block.slice(0, htmlStart!.index)], content);
            block = block.slice(cut);
        }
        if (htmlStart || (block.includes('|') && hasTableDelimiterRow(block))) {
            // Pandoc-style trailing attribute list (`{align=right}`) immediately after the
            // table, or Kramdown's `{: align=right}` on its own following line - both land
            // in this same raw block since there's no blank line separating them.
            let tableAlign: 'left' | 'center' | 'right' | undefined;
            const tableAttrLineMatch = block.match(/\n\{:?([^}\n]*)\}[ \t]*$/);
            if (tableAttrLineMatch) {
                tableAlign = parseAttributeList(tableAttrLineMatch[1]).align;
                block = block.slice(0, tableAttrLineMatch.index);
            }

            if (htmlStart) {
                // HTML is read with the HTML parser: a table's cells keep their formatting, links,
                // pictures, math, line breaks and nested tables, and text around it in the block is
                // kept too (see readHtmlBlock).
                const html = await readHtmlBlock(block);
                for (const node of html.content) {
                    if (node.type === 'table' && tableAlign) node.metadata = { ...node.metadata, align: tableAlign };
                    content.push(node);
                }
                continue;
            } else {
                const lines = trimAsciiWhitespace(block).split('\n');
                const rows: OfficeContentNode[] = [];
                // Pre-scan the separator row for per-column GFM alignment (`:--` left, `:-:` center,
                // `--:` right; a bare `--` column has none), so every cell can carry its column's
                // alignment on CellMetadata.align (the header row precedes the separator, so a
                // per-cell pass alone could not see it).
                // The delimiter row is the first such line after the header: a body row of `-` or `:`
                // cells (`| - | - |`) is a row, which was dropped as another delimiter row.
                const sepIndex = lines.findIndex((l, i) => i >= 1 && /^\|?\s*:?-+:?\s*(?:\|\s*:?-+:?\s*)*\|?$/.test(l));
                const sepLine = sepIndex >= 0 ? lines[sepIndex] : undefined;
                const columnAligns: ('left' | 'center' | 'right' | null)[] = sepLine
                    ? sepLine.replace(/^\||\|$/g, '').split('|').map(c => {
                        const t = c.trim();
                        const l = t.startsWith(':'), r = t.endsWith(':');
                        return (l && r) ? 'center' : r ? 'right' : l ? 'left' : null;
                    })
                    : [];
                for (let i = 0; i < lines.length; i++) {
                    if (i === sepIndex) continue; // Separator row (per-cell `:?-+:?`, GFM-style; accepts short cells like `|-|-|`)

                    const cellsStr = tableRowCells(lines[i]);
                    const cells: OfficeContentNode[] = cellsStr.map((c, colIdx) => {
                        // Recognize the MarkdownGenerator's own cell-alignment fallback,
                        // `<div style="text-align: X">…</div>`, and lift it into an aligned
                        // paragraph so it round-trips as alignment instead of being escaped to
                        // visible text on regeneration. Unwrap wherever it sits (e.g. inside **…**).
                        const { text: cellText, align: cellAlign } = unwrapAlignDivs(trimAsciiWhitespace(c));
                        const inline = parseInline(cellText, i === 0 ? { bold: true } : {});
                        const colAlign = columnAligns[colIdx] ?? undefined;
                        const cellMeta = colAlign ? { col: colIdx, align: colAlign } as any : undefined;
                        if (cellAlign && cellAlign !== 'left') {
                            return {
                                type: 'cell',
                                metadata: cellMeta,
                                children: [{ type: 'paragraph', metadata: { alignment: cellAlign } as any, children: inline }]
                            } as OfficeContentNode;
                        }
                        return { type: 'cell', metadata: cellMeta, children: inline } as OfficeContentNode;
                    });
                    rows.push({ type: 'row', children: cells });
                }
                // If every explicitly-aligned column agrees, also expose it as the table-level align,
                // so an editor that models one alignment per table (and HTML data-align) round-trips.
                const explicitAligns = columnAligns.filter((a): a is 'left' | 'center' | 'right' => a !== null);
                const uniformAlign = explicitAligns.length > 0 && explicitAligns.every(a => a === explicitAligns[0]) ? explicitAligns[0] : undefined;
                const resolvedTableAlign = tableAlign || uniformAlign;
                content.push({ type: 'table', metadata: resolvedTableAlign ? { align: resolvedTableAlign } : undefined, children: rows });
                continue;
            }
        }

        // Paragraph (with any display math written in it lifted out between its parts)
        appendAll(content, splitAtDisplayMath(splitParagraphLines(block), parts => ({
            type: 'paragraph',
            metadata: { alignment } as any,
            children: parts
        })));
    }
    };

    await parseBlocks(splitIntoBlocks(textStr), content);
    foldAnchorPlaceholders(content);

    // Note bodies, read now that the document has been: a one-line definition is the note's text, one
    // of several lines is read as blocks (paragraphs, a list, code). Reading one can refer to further
    // notes, whose bodies join the queue.
    let notesRead = 0;
    const readQueuedNoteBodies = async (): Promise<void> => {
        for (; notesRead < notesToRead.length; notesRead++) {
            const { note, definition } = notesToRead[notesRead];
            readingNote = note;
            if (definition.includes('\n')) {
                const children: OfficeContentNode[] = [];
                await parseBlocks(splitIntoBlocks(liftCode(definition)), children);
                foldAnchorPlaceholders(children);
                note.children = children;
                note.text = children.map(child => child.text ?? plainTextOf(child.children ?? [])).join('\n');
            } else {
                note.children = parseInline(definition);
                note.text = plainTextOf(note.children);
            }
            readingNote = undefined;
        }
    };
    await readQueuedNoteBodies();

    // Orphan footnote definitions (defined but never referenced) would otherwise vanish entirely -
    // a user who deletes a `[^x]` reference but keeps its `[^x]: ...` definition loses the
    // definition on the next save. Preserve them as trailing note nodes, marked `unreferenced` so
    // the generators route them into their footnotes section (not inline) and emit no citation
    // marker or dangling back-link. Both generators still emit the definition (md: a `[^x]:` line;
    // html: a `div[data-footnote-id]` inside `section[data-footnotes]`, which re-parses on import).
    for (const [id, definition] of footnoteDefinitions) {
        if (config.ignoreNotes) break; // ignoreNotes drops footnotes, orphan definitions included
        // Referenced from no text, but perhaps from an unreferenced note read below: its body is read
        // at once, so a note it refers to counts as referenced, and is the same note as any other
        // reference to that id.
        if (referencedFootnoteIds.has(id)) continue;
        const note: OfficeContentNode = { type: 'note', text: '', children: [], metadata: { noteType: 'footnote', noteId: id, unreferenced: true } };
        footnoteNodesById.set(id, note);
        notesToRead.push({ note, definition });
        await readQueuedNoteBodies();
        content.push(note);
    }

    return createAST('md', metadata, resolveAnchorMarks(content), attachments, config, undefined);
};

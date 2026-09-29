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
import { appendAll } from '../utils/nodeListUtils.js';

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

/** A link label as CommonMark matches it: without its outer whitespace, each inner run of whitespace one space, in lower case. */
const normalizeLabel = (label: string): string => trimAsciiWhitespace(label).replace(/\s+/g, ' ').toLowerCase();

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
            return { url: decodeMarkdownText(unbracketTarget(url)), title: decodeMarkdownText(text.slice(opener + 1, -1)) };
        }
    }
    // The target is decoded too, as a renderer decodes it (the generator escapes what it must).
    return { url: decodeMarkdownText(unbracketTarget(text)) };
}

/** A link target without the angle brackets that let it hold spaces (`<./my docs/a.md>`), which are not part of it. */
const unbracketTarget = (target: string): string => (/^<[^<>\n]*>$/.test(target) ? target.slice(1, -1) : target);

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

/** Whether a character next to a delimiter run is whitespace, as emphasis reads it: the text's edge (undefined) is. */
const isFlankingSpace = (char: string | undefined): boolean => char === undefined || /\s/.test(char);

/** Whether a character next to a delimiter run is punctuation or a symbol (Unicode P or S), as CommonMark reads it. */
const isFlankingPunctuation = (char: string | undefined): boolean => char !== undefined && /[\p{P}\p{S}]/u.test(char);

/** The character (a whole surrogate pair) ending at index `i` of `text`, or undefined at its start. */
const characterBefore = (text: string, i: number): string | undefined => {
    if (i <= 0) return undefined;
    return i >= 2 && /[\uDC00-\uDFFF]/.test(text[i - 1]) && /[\uD800-\uDBFF]/.test(text[i - 2]) ? text.slice(i - 2, i) : text[i - 1];
};

/** The character (a whole surrogate pair) starting at index `i` of `text`, or undefined at its end. */
const characterAt = (text: string, i: number): string | undefined => {
    const code = text.codePointAt(i);
    return code === undefined ? undefined : String.fromCodePoint(code);
};

/** What a matched pair of delimiter runs makes of the text between them. */
type EmphasisFlag = 'bold' | 'italic' | 'strikethrough' | 'highlight';

/**
 * A run of `*` or `_`, or a `~~` or `==`, that may open or close emphasis, strikethrough or a
 * highlight (see matchEmphasis). `remaining` of its characters are still text once it is matched;
 * `opened` and `closed` list, in the order they were matched, what it opens (with characters from its
 * end) and closes (with characters from its start). `previous` and `next` link the runs not removed.
 */
interface DelimiterRun {
    char: string;
    length: number;
    remaining: number;
    canOpen: boolean;
    canClose: boolean;
    /** It cannot close or open only for what stands beside it, so it may in the second pass (see matchEmphasis). */
    closesLate: boolean;
    opensLate: boolean;
    /** Punctuation on both sides (`**` in ``**`a`****.**``), where the rule of three is not applied (see matchEmphasis). */
    enclosed: boolean;
    opened: EmphasisFlag[];
    closed: EmphasisFlag[];
    previous: number;
    next: number;
    removed: boolean;
}

/**
 * The delimiter run `run` found at `start` in `text`, with what it can do by CommonMark's flanking
 * rules, or null when it can do nothing (it is text). A run is left-flanking when it is not followed
 * by whitespace and, if followed by punctuation, is preceded by whitespace or punctuation, and
 * right-flanking the other way about. `*`, `~~` and `==` open when left-flanking and close when
 * right-flanking; `_` opens only when it is also not right-flanking or is preceded by punctuation, and
 * closes likewise, so an underscore inside a word (snake_case) does neither. `~` and `=` form runs of
 * exactly two.
 */
function delimiterRun(text: string, start: number, run: string): DelimiterRun | null {
    const char = run[0];
    if ((char === '~' || char === '=') && run.length !== 2) return null;
    const before = characterBefore(text, start);
    const after = characterAt(text, start + run.length);
    const left = !isFlankingSpace(after) && (!isFlankingPunctuation(after) || isFlankingSpace(before) || isFlankingPunctuation(before));
    const right = !isFlankingSpace(before) && (!isFlankingPunctuation(before) || isFlankingSpace(after) || isFlankingPunctuation(after));
    const canOpen = char === '_' ? left && (!right || isFlankingPunctuation(before)) : left;
    const canClose = char === '_' ? right && (!left || isFlankingPunctuation(after)) : right;
    // (Not a `_` run: the earlier versions wrote underscores around trimmed text only.)
    const closesLate = char !== '_' && !canClose && before !== undefined;
    const opensLate = char !== '_' && !canOpen && !isFlankingSpace(after);
    if (!canOpen && !canClose && !closesLate && !opensLate) return null;
    const enclosed = isFlankingPunctuation(before) && isFlankingPunctuation(after);
    return { char, length: run.length, remaining: run.length, canOpen, canClose, closesLate, opensLate, enclosed, opened: [], closed: [], previous: -1, next: -1, removed: false };
}

/**
 * Whether `text` from `start` to `end` holds a delimiter run of one of the characters in `openers`
 * that may close (in either pass of matchEmphasis). Runs are read whole, and an escaped character is
 * no delimiter.
 */
function closesEmphasisWithin(text: string, start: number, end: number, openers: Set<string>): boolean {
    for (let i = start; i < end; i++) {
        if (text[i] === '\\') { i++; continue; }
        if (!openers.has(text[i])) continue;
        let runEnd = i + 1;
        while (text[runEnd] === text[i]) runEnd++;
        const run = delimiterRun(text, i, text.slice(i, runEnd));
        if (run && (run.canClose || run.closesLate)) return true;
        i = runEnd - 1;
    }
    return false;
}

/**
 * Pairs delimiter runs (listed in text order) into emphasis, strikethrough and highlights, as
 * CommonMark's "process emphasis" does. Each run that can close, in turn, closes the nearest earlier
 * run of its character that can open, unless the rule of three rules the pair out (when either can
 * both open and close, their lengths may not sum to a multiple of three unless both are multiples of
 * three). That rule keeps `*a**b**c*` italic around bold; it is not applied to a run with punctuation
 * on both sides, which only ever closes one run and opens the next (``**`a`****.**``, as this library
 * wrote bold code followed by bold text, which was read so before). A pair uses two characters of each run (bold, strikethrough, a highlight) or one (italic),
 * a closer with characters left closes again, and the runs between the two are text: `*a **b** c*`
 * is italic with bold inside. Where a closer finds no opener, later closers of its kind stop their
 * search below it, and a run is removed at most once, so the time is linear in the number of runs.
 *
 * Then the same again over the runs left, where a run that is not left-flanking only for what stands
 * beside it opens too, and one not right-flanking closes too (`**Note: **body` is bold `Note: `, and
 * `*Source:*Data` italic `Source:`): this library's earlier versions wrote a run's edge whitespace
 * and punctuation inside its delimiters, and read them back so, and files saved so keep their
 * formatting. It is a second pass rather than a looser rule so that what CommonMark pairs, it still
 * pairs as CommonMark does (`*a *b*` is `a`, then italic `b`).
 */
function matchEmphasis(runs: DelimiterRun[]): void {
    runs.forEach((run, i) => { run.previous = i - 1; run.next = i + 1 < runs.length ? i + 1 : -1; });
    const remove = (i: number) => {
        const run = runs[i];
        if (run.previous !== -1) runs[run.previous].next = run.next;
        if (run.next !== -1) runs[run.next].previous = run.previous;
        run.removed = true;
    };
    for (const late of [false, true]) {
        // Per kind of closer (its character, whether it can also open, its length modulo three): the
        // run at or below which no opener for it was found.
        const floors = new Map<string, number>();
        for (let c = 0; c < runs.length;) {
            const closer = runs[c];
            if (closer.removed || !(closer.canClose || (late && closer.closesLate))) { c++; continue; }
            const emphasis = closer.char === '*' || closer.char === '_';
            const closerOpens = closer.canOpen || (late && closer.opensLate);
            const kind = `${closer.char}${closerOpens ? 1 : 0}${closer.length % 3}`;
            const floor = floors.get(kind) ?? -1;
            let o = closer.previous;
            for (; o > floor; o = runs[o].previous) {
                const opener = runs[o];
                if (!(opener.canOpen || (late && opener.opensLate)) || opener.char !== closer.char) continue;
                const either = (closerOpens && !closer.enclosed) || ((opener.canClose || (late && opener.closesLate)) && !opener.enclosed);
                if (emphasis && either && (opener.length + closer.length) % 3 === 0 && !(opener.length % 3 === 0 && closer.length % 3 === 0)) continue;
                break;
            }
            if (o <= floor) {
                floors.set(kind, closer.previous);
                // A run that can open stays, and in the first pass every run, for the second.
                if (late && !closer.canOpen && !closer.opensLate) remove(c);
                c++;
                continue;
            }
            const opener = runs[o];
            const use = emphasis && (opener.remaining < 2 || closer.remaining < 2) ? 1 : 2;
            const flag: EmphasisFlag = closer.char === '~' ? 'strikethrough' : closer.char === '=' ? 'highlight' : use === 2 ? 'bold' : 'italic';
            opener.remaining -= use;
            opener.opened.push(flag);
            closer.remaining -= use;
            closer.closed.push(flag);
            for (let k = closer.previous; k !== o;) {
                const previous = runs[k].previous;
                remove(k);
                k = previous;
            }
            if (opener.remaining === 0) remove(o);
            if (closer.remaining === 0) { remove(c); c++; }
        }
    }
}

/** The units abbreviations are matched in: a run of letters, marks, digits and connectors (a word), or any other one character. */
const ABBREVIATION_UNIT = /[\p{L}\p{M}\p{N}\p{Pc}]+|[\s\S]/gu;

/**
 * A function finding the abbreviations (Markdown Extra's `*[HTML]: ...`) named by `keys` in a text:
 * where one stands as whole words (never starting or ending inside a run of letters or digits), the
 * longest starting at the leftmost place, then the same in the rest after it. The keys are read into
 * an Aho-Corasick automaton of their words backwards, once per document, which gives the longest key
 * starting at each word in one backward scan, so a text costs time linear in its length however many
 * keys there are (a pattern alternating every key tried each of them at every place of every text,
 * built again for each: 10,000 of them over 10,000 lines took 12 seconds).
 */
function abbreviationFinder(keys: Iterable<string>): (text: string) => { start: number; end: number }[] {
    const unitIds = new Map<string, number>();
    // Per unit, the node it leads to from each node that has such a child.
    const children = new Map<number, Map<number, number>>();
    const parents: number[] = [0];
    const units: number[] = [-1];
    const depths: number[] = [0];
    // The number of units of the key a node completes (0 for none).
    const keyUnits: number[] = [0];
    for (const key of keys) {
        const words = key.match(ABBREVIATION_UNIT) ?? [];
        let node = 0;
        for (let w = words.length - 1; w >= 0; w--) {
            let unit = unitIds.get(words[w]);
            if (unit === undefined) unitIds.set(words[w], unit = unitIds.size);
            const edges = children.get(unit) ?? children.set(unit, new Map()).get(unit)!;
            let child = edges.get(node);
            if (child === undefined) {
                edges.set(node, child = parents.length);
                parents.push(node);
                units.push(unit);
                depths.push(depths[node] + 1);
                keyUnits.push(0);
            }
            node = child;
        }
        if (words.length) keyUnits[node] = words.length;
    }
    // Failure links and, per node, the longest key its match ends with, in order of depth (a node's
    // failure is shallower than it).
    const byDepth = [...parents.keys()].sort((a, b) => depths[a] - depths[b]);
    const failure = new Int32Array(parents.length);
    const longest = new Int32Array(parents.length);
    for (const node of byDepth) {
        if (node === 0) continue;
        const edges = children.get(units[node])!;
        let fallback = failure[parents[node]];
        while (parents[node] !== 0 && fallback !== 0 && !edges.has(fallback)) fallback = failure[fallback];
        failure[node] = parents[node] === 0 ? 0 : edges.get(fallback) ?? 0;
        longest[node] = keyUnits[node] || longest[failure[node]];
    }
    return text => {
        const words = [...text.matchAll(ABBREVIATION_UNIT)];
        // The number of units of the longest key starting at each unit.
        const starting = new Int32Array(words.length);
        let state = 0;
        for (let w = words.length - 1; w >= 0; w--) {
            const unit = unitIds.get(words[w][0]);
            const edges = unit === undefined ? undefined : children.get(unit);
            if (!edges) { state = 0; continue; }
            while (state !== 0 && !edges.has(state)) state = failure[state];
            state = edges.get(state) ?? 0;
            starting[w] = longest[state];
        }
        const found: { start: number; end: number }[] = [];
        for (let w = 0; w < words.length;) {
            if (!starting[w]) { w++; continue; }
            const last = words[w + starting[w] - 1];
            found.push({ start: words[w].index!, end: last.index! + last[0].length });
            w += starting[w];
        }
        return found;
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

/** A reference definition's title, `"..."`, `'...'` or `(...)` (each may escape its delimiter), then nothing else on its line. */
const DEFINITION_TITLE = /^[ \t]*(?:"((?:[^"\\]|\\.)*)"|'((?:[^'\\]|\\.)*)'|\(((?:[^()\\]|\\.)*)\))[ \t]*$/;

/**
 * The link reference definition (`[label]: target "title"`) starting at line `i` of `lines`, as
 * CommonMark reads one, or null: up to three spaces in, a label of at most 999 characters (up to its
 * first unescaped `]`) holding something besides whitespace, then its target, `<in angle brackets>`
 * or with no spaces, on the same line or the next (`[x]:` then `  https://...`), then an optional
 * title, on the target's line or the next. A title line that is no title ends the definition before
 * it; anything else after the target makes the lines no definition. `lines` says how many it takes.
 * Its target and title are decoded as CommonMark decodes them.
 */
function readLinkDefinition(lines: string[], i: number): { label: string; url: string; title?: string; lines: number } | null {
    const head = /^ {0,3}\[((?:[^\]\\]|\\.){1,999})\]:(.*)$/.exec(lines[i]);
    if (!head || !trimAsciiWhitespace(head[1])) return null;
    let rest = head[2];
    let used = 1;
    if (!trimAsciiWhitespace(rest)) {
        if (i + 1 >= lines.length || !trimAsciiWhitespace(lines[i + 1])) return null;
        rest = lines[i + 1];
        used = 2;
    }
    const target = /^[ \t]*(?:<((?:[^<>\\]|\\.)*)>|([^\s<]\S*))(?=[ \t]|$)/.exec(rest);
    if (!target) return null;
    const url = decodeMarkdownText(target[1] ?? target[2]);
    const after = rest.slice(target[0].length);
    const title = trimAsciiWhitespace(after) ? DEFINITION_TITLE.exec(after) : i + used < lines.length ? DEFINITION_TITLE.exec(lines[i + used]) : null;
    if (trimAsciiWhitespace(after) && !title) return null;
    if (!title) return { label: head[1], url, lines: used };
    return { label: head[1], url, title: decodeMarkdownText(title[1] ?? title[2] ?? title[3]), lines: trimAsciiWhitespace(after) ? used : used + 1 };
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
 * The names of HTML's elements (current and obsolete ones a browser still renders): a tag of one of
 * them written in capitals (`<IMG SRC>`, `<B>`, `<BR/>`, older HTML's way) is HTML, not a component.
 */
const HTML_ELEMENT_NAMES = new Set([
    'a', 'abbr', 'acronym', 'address', 'area', 'article', 'aside', 'audio', 'b', 'base', 'basefont', 'bdi', 'bdo', 'big', 'blink',
    'blockquote', 'body', 'br', 'button', 'canvas', 'caption', 'center', 'cite', 'code', 'col', 'colgroup', 'data', 'datalist', 'dd',
    'del', 'details', 'dfn', 'dialog', 'dir', 'div', 'dl', 'dt', 'em', 'embed', 'fieldset', 'figcaption', 'figure', 'font', 'footer',
    'form', 'frame', 'frameset', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'head', 'header', 'hgroup', 'hr', 'html', 'i', 'iframe', 'img',
    'input', 'ins', 'kbd', 'label', 'legend', 'li', 'link', 'main', 'map', 'mark', 'marquee', 'menu', 'meta', 'meter', 'nav', 'nobr',
    'noscript', 'object', 'ol', 'optgroup', 'option', 'output', 'p', 'param', 'picture', 'pre', 'progress', 'q', 'rp', 'rt', 'ruby', 's',
    'samp', 'script', 'search', 'section', 'select', 'slot', 'small', 'source', 'span', 'strike', 'strong', 'style', 'sub', 'summary',
    'sup', 'table', 'tbody', 'td', 'template', 'textarea', 'tfoot', 'th', 'thead', 'time', 'title', 'tr', 'track', 'tt', 'u', 'ul',
    'var', 'video', 'wbr',
]);

/**
 * MDX components removed, their content kept (parse-only: MDX is never written back). A component is
 * a tag whose name starts with a capital, as React and MDX tell it from HTML, and that is not an HTML
 * element's name in capitals (see HTML_ELEMENT_NAMES: `<IMG SRC="a.png">` is a picture and `<B>` bold;
 * they were taken out, the picture with them). `<Component ... />` goes, and
 * `<Component ...>inner</Component>` becomes `inner`, nested ones too. A tag inside a code span on its
 * line (after an odd number of backticks) is code, and one with no matching closing tag is text. One
 * scan finds the tags and pairs each closing tag with the latest open one of its name, so the time is
 * linear in the text.
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
        if (!/[a-z]/.test(name) && HTML_ELEMENT_NAMES.has(name.toLowerCase())) continue;
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

/** Link text: characters (line ends too), backslash escapes, and bracket pairs nested up to three deep. */
const LINK_TEXT = (() => {
    const unit = String.raw`[^\[\]\\]|\\[\s\S]`;
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
 * - `stars`, `underscores`, `tildes` and `equals`: a run of `*` or `_`, or a `~~` or `==`, which may
 *   open or close emphasis, strikethrough or a highlight. They are paired once the whole text is read
 *   (see delimiterRun and matchEmphasis), as CommonMark pairs them.
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
 * A paragraph is read as one text, its line ends in it, so any of these may run across lines (a
 * target, a label's id, a title and inline math may not). Every alternative stops scanning early, so a
 * paragraph costs time in proportion to its length however it is written: link text holds no stray
 * `[`, a target stops at its first `)`, labels hold no `[`, an HTML-style span stops at another opening
 * tag of its kind, an attribute list is at most 1000 characters, and the closers of code spans are
 * looked up rather than scanned for. Nothing else is capped, so a long link or span is still a link or
 * span.
 */
const INLINE_TOKENS = [
    String.raw`\\(?<esc>[!-\/:-@\[-\x60{-~])`,
    String.raw`(?<imgBang>!?)\[(?<imgAlt>${LINK_TEXT})\]\((?<imgUrl>[^\s()\[\]]*\s+(?:"(?:[^"\\\n]|\\[^\n])*"|'(?:[^'\\\n]|\\[^\n])*')\s*|${LINK_DESTINATION})\)(?:\{(?<imgAttrs>[^}\n]{0,1000})\})?`,
    String.raw`(?<stars>\*+)`,
    String.raw`(?<underscores>_+)`,
    String.raw`(?<tildes>~+)`,
    String.raw`(?<equals>=+)`,
    String.raw`(?<codeFence>\x60+)`,
    String.raw`<u>(?<underline>(?:(?!<u>)[\s\S])+?)<\/u>`,
    String.raw`<sub>(?<subscript>(?:(?!<sub>)[\s\S])+?)<\/sub>`,
    String.raw`<sup>(?<superscript>(?:(?!<sup>)[\s\S])+?)<\/sup>`,
    String.raw`(?<lineBreak><[bB][rR]\s*\/?>)`,
    String.raw`(?<anchorTag><[aA]\s[^<>\n]*>[ \t]*<\/[aA]>)`,
    String.raw`(?<htmlComment><!--)`,
    String.raw`<span\s+style="(?<spanStyle>[^"\n]*)">(?<spanContent>(?:(?!<span[\s>])[\s\S])+?)<\/span>`,
    String.raw`<(?<htmlTag>[a-zA-Z][a-zA-Z0-9]*)(?<htmlAttrs>\s[^<>]{0,1000})?>`,
    String.raw`\[\^(?<footnoteId>[^\[\]\n]+)\]`,
    String.raw`\[@(?<citationKey>[a-zA-Z0-9_:.-]+)\]`,
    String.raw`\[\[(?<wikiPage>[^\[\]|\n]+)(?:\|(?<wikiAlias>[^\[\]\n]+))?\]\]`,
    String.raw`(?<refBang>!?)\[(?<refText>[^\[\]]*)\]\[(?<refId>[^\[\]]*)\]`,
    String.raw`(?<shortBang>!?)\[(?<shortText>[^\[\]]+)\]`,
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
    // A byte order mark starting the file is no text (as a renderer reads it): left in, it stood before
    // the first block's marker, and a heading or list on the first line was read as a paragraph.
    if (textStr.charCodeAt(0) === 0xFEFF) textStr = textStr.slice(1);
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
    // normalized label, matching CommonMark's case-insensitive reference matching (see
    // readLinkDefinition). A definition does not interrupt a paragraph: one on the line after text
    // (`- item` then `[x]: y`) is that text's continuation. After a heading, a rule or another
    // definition it is one (a heading then a list of definitions lost them all).
    const linkDefinitions = new Map<string, { url: string; title?: string }>();
    // The characters reference definitions (a link's or picture's target, an abbreviation's title) may
    // repeat into the document beyond their first use: each use is written out in full by every
    // writer, so a long target used many times (100 KB, 200 times) made output growing with the
    // square of the input. Past this, a reference stays its text, as an unknown one does.
    let referenceBudget = 16 * 1024 * 1024;
    const usedDefinitions = new Set<string>();
    const expandReference = (key: string, size: number): boolean => {
        if (!usedDefinitions.has(key)) { usedDefinitions.add(key); return true; }
        if (size > referenceBudget) return false;
        referenceBudget -= size;
        return true;
    };
    {
        const lines = textStr.split('\n');
        const kept: string[] = [];
        // Whether the line before is paragraph text: any line with content but a heading, a rule or a
        // heading's underline (a definition's lines are taken out).
        let inParagraph = false;
        for (let i = 0; i < lines.length; i++) {
            const definition = inParagraph ? null : readLinkDefinition(lines, i);
            if (definition) {
                linkDefinitions.set(normalizeLabel(definition.label), definition.title === undefined ? { url: definition.url } : { url: definition.url, title: definition.title });
                for (let k = 0; k < definition.lines; k++) kept.push('');
                i += definition.lines - 1;
                continue;
            }
            const line = lines[i];
            kept.push(line);
            inParagraph = !!trimAsciiWhitespace(line) && !ATX_HEADING_START.test(line) && !THEMATIC_BREAK.test(line) && !SETEXT_UNDERLINE.test(line);
        }
        textStr = kept.join('\n');
    }

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
        // Text of this call's own, with its character references decoded (`&amp;` is `&`). Once: what a
        // nested call returns (a link's text, an element's content) it decoded already, and decoding
        // that again turned a literal `&amp;quot;` inside `<b>...</b>` into `"`.
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
        const buildLinkOrImageNodes = (isImage: boolean, altText: string, { url, title }: { url: string; title?: string }, attrsStr?: string): OfficeContentNode[] => {
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
                if (href !== undefined) {
                    // Whitespace at either end of what the link holds (the line ends around a logo
                    // written `<a href>`, `<img>`, `</a>` on lines of their own) stands beside it: a
                    // linked space was a link of its own on the next save.
                    // (Only around other content, and not from a run a note is on.)
                    const movable = (node: OfficeContentNode | undefined): node is OfficeContentNode & { text: string } =>
                        content.length > 1 && node?.type === 'text' && typeof node.text === 'string' && !node.notes?.length;
                    let before = '', after = '';
                    const first = content[0];
                    if (movable(first)) {
                        before = first.text.slice(0, first.text.length - trimStartChars(first.text, ASCII_WHITESPACE).length);
                        if (before === first.text) content.shift(); else if (before) first.text = first.text.slice(before.length);
                    }
                    const last = content[content.length - 1];
                    if (movable(last)) {
                        after = last.text.slice(trimEndChars(last.text, ASCII_WHITESPACE).length);
                        if (after === last.text) content.pop(); else if (after) last.text = last.text.slice(0, -after.length);
                    }
                    return [...(before ? [plainText(before)] : []), ...applyLink(content, href, attrs.get('title') || undefined), ...(after ? [plainText(after)] : [])];
                }
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
        // is plain text. A backtick run and a comment's `<!--` are matched as openers only: where each
        // closes is looked up (codeSpanCloser, commentAt), in time linear in the text however many
        // never close. An opener with no closer is ordinary text, left in its run. A run of emphasis
        // delimiters is set aside and paired once the whole text is read (matchEmphasis), so emphasis
        // may hold links, code and other emphasis, and run across a line end. The tokenizer is shared by
        // every call (compiling it for each piece of text doubled the time of a document of short
        // paragraphs); each search sets where it starts, so the nested calls made while handling a match
        // cannot disturb this one.
        //
        // What the scan finds, in order: stretches of plain text (as offsets into `text`), finished
        // nodes, delimiter runs, and footnotes (which go to the node before them).
        const items: ({ start: number; end: number } | { node: OfficeContentNode } | { run: DelimiterRun } | { note: OfficeContentNode })[] = [];
        const runs: DelimiterRun[] = [];
        // The characters of emphasis that may be open where the scan stands: per character, the runs so
        // far that only open less those that only close (not below none), a count rather than a pairing.
        const openCounts = new Map<string, number>();
        const openers = new Set<string>();
        const push = (found: OfficeContentNode | OfficeContentNode[]) => {
            if (Array.isArray(found)) for (const node of found) items.push({ node });
            else items.push({ node: found });
        };
        const closeCodeSpan = codeSpanCloser(text);
        let lastIndex = 0; // end of the text already taken
        let next = 0; // where the next search starts
        const closes: CommentCloseCache = { at: -1, from: Number.MAX_SAFE_INTEGER };

        for (;;) {
            INLINE_TOKEN_REGEX.lastIndex = next;
            const match = INLINE_TOKEN_REGEX.exec(text);
            if (!match) break;
            next = INLINE_TOKEN_REGEX.lastIndex;
            const g = match.groups!;
            const delimiters = g.stars ?? g.underscores ?? g.tildes ?? g.equals;
            if (delimiters !== undefined) {
                // A run that can neither open nor close is text, left in its stretch.
                const run = delimiterRun(text, match.index, delimiters);
                if (!run) continue;
                if (match.index > lastIndex) items.push({ start: lastIndex, end: match.index });
                items.push({ run });
                runs.push(run);
                const count = openCounts.get(run.char) ?? 0;
                const opens = run.canOpen || run.opensLate, closes = run.canClose || run.closesLate;
                const now = opens && !closes ? count + 1 : closes && !opens ? Math.max(0, count - 1) : count;
                openCounts.set(run.char, now);
                if (now > 0) openers.add(run.char); else openers.delete(run.char);
                lastIndex = next;
                continue;
            }
            // Math does not run past the closer of emphasis opened before it (`_$data_: ... [$x`): a
            // `$` read as the start of math there swallowed the closer and the text after it, where
            // the text is emphasis. Its `$` is then text. (Emphasis closed before it does not count:
            // `*a* and $x*y$` is math.)
            if ((g.mathInline !== undefined || g.mathDisplay !== undefined) && openers.size && closesEmphasisWithin(text, match.index + 1, next - 1, openers)) {
                next = match.index + 1;
                continue;
            }
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
            // A reference to no definition is text: its brackets, and what they hold read as any text
            // is (`[see *this*]` keeps its emphasis).
            if (g.refText !== undefined || g.shortText !== undefined) {
                const label = normalizeLabel(g.refText !== undefined ? (g.refId || g.refText) : g.shortText);
                const def = linkDefinitions.get(label);
                if (!def || !expandReference(`link:${label}`, def.url.length + (def.title?.length ?? 0))) {
                    next = match.index + (g.refBang || g.shortBang ? 2 : 1);
                    continue;
                }
            }
            let closeAt = -1;
            if (g.codeFence !== undefined) {
                closeAt = closeCodeSpan(next, g.codeFence.length);
                if (closeAt === -1) continue;
                next = closeAt + g.codeFence.length;
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
            if (match.index > lastIndex) items.push({ start: lastIndex, end: match.index });

            if (g.esc !== undefined) { // Backslash-escaped punctuation
                push(plainText(g.esc));
            } else if (g.imgAlt !== undefined) { // Image or Link
                push(buildLinkOrImageNodes(g.imgBang === '!', g.imgAlt, splitUrlTitle(g.imgUrl), g.imgAttrs));
            } else if (g.codeFence !== undefined) { // Inline code, closed by a run of as many backticks
                push({ type: 'text', text: codeSpanContent(text.slice(match.index + g.codeFence.length, closeAt)), formatting: { ...currentFormatting, font: 'monospace' } });
            } else if (g.underline !== undefined) { // Underline
                push(parseInline(g.underline, { ...currentFormatting, underline: true }));
            } else if (g.subscript !== undefined) { // Subscript
                push(parseInline(g.subscript, { ...currentFormatting, subscript: true }));
            } else if (g.superscript !== undefined) { // Superscript
                push(parseInline(g.superscript, { ...currentFormatting, superscript: true }));
            } else if (g.htmlTag !== undefined) { // Raw inline HTML: an element, or a picture (see inlineHtml)
                push(inlineHtml(g.htmlTag.toLowerCase(), g.htmlAttrs ?? '', closeAt === -1 ? '' : text.slice(match.index + match[0].length, closeAt)));
            } else if (g.anchorTag !== undefined) { // An empty anchor: an id (see resolveAnchorMarks)
                push(anchorMark(anchorIds));
            } else if (g.lineBreak !== undefined) { // Raw inline <br>/<br/>/<br /> - a hard line break.
                // MarkdownGenerator emits a raw <br> for a line break inside a table cell (a GFM pipe
                // cell can't hold a newline), so the parser must read it back symmetrically as a break
                // node instead of escaping it to literal `&lt;br&gt;` text and destroying it.
                push({ type: 'break', metadata: { breakType: 'carriageReturn' } as BreakMetadata });
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
                push(parseInline(g.spanContent, styled));
            } else if (g.footnoteId !== undefined) { // Footnote reference
                const noteId = g.footnoteId;
                // ignoreNotes drops footnotes at parse time (as in DOCX/ODT/PDF): swallow the marker and
                // attach nothing. Advance lastIndex past the marker (so the gap text is not re-emitted)
                // before skipping. The orphan sweep below is likewise skipped.
                if (config.ignoreNotes) { lastIndex = next; continue; }
                const noteNode = footnoteNode(noteId);
                // Notes attach to the node before them (matches WordParser's convention; see below). A
                // reference that would make a note contain itself stays text.
                if (noteNode) items.push({ note: noteNode });
                else push(plainText(match[0]));
            } else if (g.citationKey !== undefined) { // Citation reference
                push({ type: 'text', text: g.citationKey, metadata: { citationKey: g.citationKey } as TextMetadata });
            } else if (g.wikiPage !== undefined) { // Wikilink
                const page = g.wikiPage.trim();
                const alias = g.wikiAlias?.trim();
                push({ type: 'text', text: alias || page, metadata: { link: page, linkType: 'internal', wikilink: true } as TextMetadata });
            } else if (g.refText !== undefined || g.shortText !== undefined) { // A reference link or picture: [text][ref], [text][], [text]
                const def = linkDefinitions.get(normalizeLabel(g.refText !== undefined ? (g.refId || g.refText) : g.shortText))!;
                push(buildLinkOrImageNodes((g.refBang || g.shortBang) === '!', g.refText ?? g.shortText, def));
            } else if (g.autolinkUrl !== undefined) { // <url> autolink
                // Its references decoded in the target too, as in the text (CommonMark reads them in URLs):
                // left in the target, `&amp;` was written back as `&amp;amp;`.
                const url = decodeCharacterReferences(g.autolinkUrl);
                push({ type: 'text', text: url, formatting: Object.keys(currentFormatting).length > 0 ? { ...currentFormatting } : undefined, metadata: { link: url, linkType: 'external' } as TextMetadata });
            } else if (g.autolinkEmail !== undefined) { // <address> autolink: a mail link (CommonMark)
                push({ type: 'text', text: g.autolinkEmail, formatting: Object.keys(currentFormatting).length > 0 ? { ...currentFormatting } : undefined, metadata: { link: `mailto:${g.autolinkEmail}`, linkType: 'external' } as TextMetadata });
            } else if (g.mathDisplay !== undefined) { // Display math inside the text
                // As KaTeX, MathJax, GitLab and Pandoc read it: display math, wherever it is written.
                if (!g.mathDisplay.trim()) push(plainText(match[0]));
                else push({ type: 'code', text: g.mathDisplay.trim(), metadata: { math: displayMath ? 'block' : 'inline' } as CodeMetadata });
            } else if (g.mathInline !== undefined) { // Inline math
                push({ type: 'code', text: g.mathInline, metadata: { math: 'inline' } as CodeMetadata });
            } else if (comment) { // Inline source comment: a hidden note, kept verbatim
                push({ type: 'comment', text: comment.body, metadata: { sourceSyntax: 'html' } as CommentMetadata });
            }

            lastIndex = next;
        }
        if (lastIndex < text.length) items.push({ start: lastIndex, end: text.length });

        matchEmphasis(runs);

        // The nodes, in order: each run's matched characters open or close the formatting they give
        // what is between them (bold, italic, strikethrough, highlight, counted, as they nest), and its
        // other characters are text. Plain text under one formatting is one run of text.
        const nodes: OfficeContentNode[] = [];
        const open: Record<EmphasisFlag, number> = { bold: 0, italic: 0, strikethrough: 0, highlight: 0 };
        let emphasis: TextFormatting | undefined; // what the open runs give, when any is open
        const updateEmphasis = () => {
            emphasis = open.bold || open.italic || open.strikethrough || open.highlight ? {
                ...(open.bold && { bold: true }), ...(open.italic && { italic: true }),
                ...(open.strikethrough && { strikethrough: true }), ...(open.highlight && { backgroundColor: '#ffff00' }),
            } : undefined;
        };
        let pending = '';
        const flush = () => {
            if (!pending) return;
            appendAll(nodes, textNodes(pending, emphasis ? { ...currentFormatting, ...emphasis } : currentFormatting));
            pending = '';
        };
        for (const item of items) {
            if ('start' in item) {
                pending += text.slice(item.start, item.end);
            } else if ('run' in item) {
                const { run } = item;
                if (run.closed.length) {
                    flush();
                    for (const flag of run.closed) open[flag]--;
                    updateEmphasis();
                }
                pending += run.char.repeat(run.remaining);
                if (run.opened.length) {
                    flush();
                    for (const flag of run.opened) open[flag]++;
                    updateEmphasis();
                }
            } else if ('note' in item) {
                // After the text before it: on that text's last node, else on an empty run of its own.
                flush();
                const target = nodes[nodes.length - 1];
                if (target) (target.notes ??= []).push(item.note);
                else nodes.push({ type: 'text', text: '', notes: [item.note] });
            } else {
                flush();
                const { node } = item;
                if (emphasis && node.type === 'text') node.formatting = { ...node.formatting, ...emphasis };
                // Display math in emphasis is inline: a block cannot go there.
                else if (emphasis && node.type === 'code' && (node.metadata as CodeMetadata | undefined)?.math === 'block') node.metadata = { ...node.metadata, math: 'inline' } as CodeMetadata;
                nodes.push(node);
            }
        }
        flush();

        return applyAbbreviations(nodes);
    };

    /**
     * Text read under `formatting` (its character references decoded, unless it is code), as nodes:
     * a line end in it is a soft break, a space (with the spaces and tabs around it), or, after two or
     * more spaces or a backslash that is not itself escaped, a line break. (A line's leading spaces and
     * tabs are taken off before; see splitParagraphLines.)
     */
    const textNodes = (raw: string, formatting: TextFormatting): OfficeContentNode[] => {
        const node = (t: string): OfficeContentNode => ({ type: 'text', text: formatting.font === 'monospace' ? t : decodeCharacterReferences(t), formatting: Object.keys(formatting).length > 0 ? { ...formatting } : undefined });
        if (!raw.includes('\n')) return [node(raw)];
        const out: OfficeContentNode[] = [];
        let current = '';
        let start = 0;
        for (let end = raw.indexOf('\n'); ; end = raw.indexOf('\n', start)) {
            const line = raw.slice(start, end === -1 ? raw.length : end);
            if (end === -1) { current += line; break; }
            // Found from the end: an end-anchored pattern retried every run of spaces in the line,
            // quadratic in a long one.
            const spaces = line.length - trimEndChars(line, ' ').length;
            if (spaces >= 2 || (spaces === 0 && line.endsWith('\\') && !isEscapedAt(line, line.length - 1))) {
                current += line.slice(0, spaces >= 2 ? -spaces : -1);
                if (current) out.push(node(current));
                current = '';
                out.push({ type: 'break', metadata: { breakType: 'carriageReturn' } as BreakMetadata });
            } else {
                current += `${trimEndChars(line, ' \t')} `;
            }
            start = end + 1;
        }
        if (current) out.push(node(current));
        return out;
    };

    // The finder of abbreviations, built once for the document when a text is first read (the
    // definitions are all read before any text is).
    let findAbbreviations: ((text: string) => { start: number; end: number }[]) | undefined;

    // Splits abbreviation occurrences out of plain text nodes so they carry
    // TextMetadata.abbreviationTitle, rendered as <abbr title> in HTML/editor output.
    const applyAbbreviations = (nodes: OfficeContentNode[]): OfficeContentNode[] => {
        if (abbreviationDefinitions.size === 0) return nodes;
        findAbbreviations ??= abbreviationFinder(abbreviationDefinitions.keys());

        const result: OfficeContentNode[] = [];
        for (const node of nodes) {
            const found = node.type !== 'text' || !node.text || node.metadata ? [] : findAbbreviations(node.text);
            if (found.length === 0) {
                result.push(node);
                continue;
            }
            const text = node.text!;
            let lastIndex = 0;
            for (const { start, end } of found) {
                if (start > lastIndex) result.push({ type: 'text', text: text.substring(lastIndex, start), formatting: node.formatting });
                const abbreviation = text.substring(start, end);
                const title = abbreviationDefinitions.get(abbreviation) ?? '';
                result.push({
                    type: 'text',
                    text: abbreviation,
                    formatting: node.formatting,
                    ...(expandReference(`abbr:${abbreviation}`, title.length) && { metadata: { abbreviationTitle: title } as TextMetadata }),
                });
                lastIndex = end;
            }
            if (lastIndex < text.length) result.push({ type: 'text', text: text.substring(lastIndex), formatting: node.formatting, ...(node.notes && { notes: node.notes }) });
            else if (node.notes) result[result.length - 1].notes = node.notes;
        }
        return result;
    };

    // A paragraph's inline content, read as one text across its lines as CommonMark reads it, so
    // emphasis, a link, a code span or an inline element that an editor wrapped onto the next line is
    // still that construct: a line end is a soft break or a line break (see textNodes). A continuation
    // line's leading spaces and tabs are not part of the text (CommonMark); a comment's lines, joined
    // first, keep theirs.
    const splitParagraphLines = (block: string): OfficeContentNode[] =>
        parseInline(joinCommentLines(block.split('\n')).map((line, i) => (i === 0 ? line : trimStartChars(line, ' \t'))).join('\n'), {}, true);

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
                appendAll(carried, (((node.metadata as any)?.anchorIds as string[]) || []));
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
    // The last list read: its last item, and its id and each level's counter, marker and indentation,
    // for a list block after it to go on with (see the list branch).
    let previousList: { last: OfficeContentNode; listId: string; counters: Map<number, number>; markers: Map<number, string>; indents: number[] } | undefined;

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
            // A list block right after a list (blank lines between) whose first item is of the kind of
            // its top-level items goes on with it: a loose list, as CommonMark reads it (`1. a`, a blank
            // line, `1. b` is items 1 and 2 of one list, where each started a list and was numbered 1).
            const first = /^([ \t]*)([-*+]|\d+[.)])/.exec(lines[0])!;
            const firstKind = /\d/.test(first[2]) ? first[2].slice(-1) : first[2];
            const continues = previousList && content[content.length - 1] === previousList.last
                && previousList.markers.get(0) === firstKind && first[1].replace(/\t/g, '    ').length <= (previousList.indents[0] ?? 0)
                ? previousList : undefined;
            let listId = continues?.listId ?? `md-list-${listIdCounter++}`;
            const listCounters = continues?.counters ?? new Map<number, number>();
            // The marker each level's list uses (its bullet, or an ordered list's `.` or `)`): another
            // starts a new list, as CommonMark reads it (`1. one` after `- b` counted on from the bullets).
            const levelMarkers = continues?.markers ?? new Map<number, string>();
            // Relative indent stack (not a fixed-width divisor) so nesting level is
            // computed from what indentation actually appeared in this block, rather
            // than assuming a specific indent width. This makes the parser agnostic to
            // 2-space (hand-written), 4-space (this generator's own output), or
            // tab-indented (normalized to a 4-column stop) nested lists.
            const indentStack: number[] = continues?.indents ?? [];
            // The item being read and its lines: its text and indented continuation lines (see the
            // sub-splitter above), read as one text when the item ends, as a paragraph's lines are,
            // so emphasis or a link may run from one line to the next.
            let item: { node: OfficeContentNode; lines: string[] } | undefined;
            const endItem = () => {
                if (!item) return;
                const children = parseInline(trimEndChars(item.lines.join('\n'), ASCII_WHITESPACE));
                item.node.children = children;
                item.node.text = plainTextOf(children);
                item = undefined;
            };

            for (const line of lines) {
                // A marker alone on its line is an empty item.
                const match = line.match(/^([ \t]*)([-*+]|\d+[.)])(?:[ \t]+(.*))?$/);
                if (match) {
                    endItem();
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

                    const listNode: OfficeContentNode = {
                        type: 'list',
                        text: '',
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
                        children: []
                    };
                    content.push(listNode);
                    item = { node: listNode, lines: [itemText] };
                } else if (item && trimAsciiWhitespace(line).length > 0 && /^(?: {2,}|\t)/.test(line)) {
                    // Indented continuation line: part of the item's text (only lines of text: no
                    // nested code, quote or paragraph).
                    item.lines.push(trimStartChars(line, ' \t'));
                }
            }
            previousList = { last: content[content.length - 1], listId, counters: listCounters, markers: levelMarkers, indents: indentStack };
            endItem();
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

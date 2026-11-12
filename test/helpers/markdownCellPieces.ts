/**
 * Markdown in an HTML table's cells (GitHub's way of writing one: `<td>`, a blank line, Markdown, a blank line,
 * `</td>`) is read, and no token or placeholder standing for a piece of it while the table is read reaches the AST.
 * Each document is a table whose cell holds a Markdown piece (`wN *wN*` between blank lines) beside block starts
 * of every kind: CommonMark's seven HTML block starts (raw-text elements, comments, processing instructions,
 * declarations, CDATA, block-level tags, other tags), some across pieces and some closing in a piece what an
 * earlier one opened (a quoted value, a tag, a comment, a script); lifted blocks (fences, math, admonitions);
 * setext underlines, rules, headings, list items, quotes, indented code, pipe tables, definitions, embeds and
 * anchors; then a tail after `</table>` (text, an underline, a rule, a heading, a list, a quote, a block).
 *
 * Every marker word (`w0`, `w1`, ...) must be a word of the AST's text, metadata strings or notes, and none of
 * those may hold a private-use character (the reader's tokens and placeholders are made of them) or a
 * placeholder's name. Text the HTML reader reads as nothing wherever it stands, the same HTML written on one line
 * included, is the word `hidden` and no marker: a tag's attribute names, an attribute it does not keep (`title` on
 * a cell), the text of a `<script>` or `<style>`, a processing instruction and a declaration. That is the HTML
 * reader's reading of such markup in any input, not something this join loses. Tables whose tags do not balance
 * are not drawn: a table that never closes is not joined, so no token can arise in it.
 *
 * The documents: every block start with every tail, each separator (a line break or a blank line) and on each
 * side of the piece, in one opening (systematic); then random ones of one to four block starts with every opening
 * and tail, from a full-period linear congruential generator (Hull-Dobell: c odd, a - 1 divisible by 4, modulus
 * 2^31), the distinct ones counted.
 *
 * The quote part (see quoteDocument): a joined table whose cell holds Markdown pieces beside HTML that leaves a
 * quote, a value, a comment or CDATA open, or holds a quote the HTML reader's tag scan pairs only with a later one
 * (an unquoted value holding an apostrophe, a value missing its closing quote), and a line before the table in its
 * block holding the same (in a code span, in prose, in a wrapper's tag). The pieces hold apostrophes, quotes, `>`,
 * `-->` and `]]>` as prose does. Every marker word is kept: the join may write a piece as text only where the HTML
 * reader reads it as the join does, and keeps its token elsewhere. Drawn so that only the HTML reader's own reading
 * of such markup in any input could hide a word, and then it is the word `hidden`: one quote that opens in markup
 * per document (two pair with each other in its tag scan, wherever they stand), nothing but pieces after a value or
 * comment left open to the end, and no construct closed only by a piece after a quote whose tag a `>` inside it may
 * end (where that tag ends is the reader's to choose, see the Markdown reader's writeEndingPieces). Drawn from the
 * same generator, carried on, the distinct ones counted.
 */
import type { OfficeContentNode, OfficeParserAST } from '../../src/types';

/** What checkMarkdownCellPieces found. */
export interface MarkdownCellPieces {
    /** Documents read: the systematic ones, the random ones drawn and how many of those were distinct. */
    systematic: number; random: number; distinct: number;
    /** Documents of the quote part drawn, and how many of those were distinct. */
    quotes: number; quotesDistinct: number;
    /** Documents whose AST shows a token or placeholder, and documents losing a marker word. */
    leaked: number; lost: number;
    /** The block start kinds the random part never drew, and the quote part's kinds it never drew. */
    undrawn: string[];
    /** The first few failures, each with its document and what its AST holds. */
    failures: string[];
}

/** A block start that can stand in a cell beside a Markdown piece, by its kind; `w` gives the next marker word. */
type BlockStart = [kind: string, make: (w: () => string) => string];

const BLOCK_STARTS: BlockStart[] = [
    // CommonMark's HTML block start 1: <pre>, <script>, <style>, <textarea>.
    ['pre', w => `<pre>\n${w()}\n</pre>`], ['pre-fence', w => `<pre>${w()}\n\n\`\`\`\n${w()}\n\`\`\`\n\n</pre>`],
    ['pre-admonition', w => `<pre>${w()}\n\n:::note\n${w()}\n:::\n\n</pre>`], ['textarea', w => `<textarea>${w()}</textarea>`],
    ['textarea-pieces', w => `<textarea>\n\n${w()}\n\n</textarea>`], ['script', () => '<script>x = 1</script>'],
    ['script-pieces', () => '<script>\n\nhidden\n\n</script>'], ['style-pieces', () => '<style>\n\nhidden\n\n</style>'],
    ['script-closed-in-piece', w => `<script>\n\nhidden </script> ${w()}`],
    // 2: comments.
    ['comment', w => `<!-- ${w()} -->`], ['comment-lines', w => `<!--\n${w()}\n-->`], ['comment-pieces', w => `<!-- x\n\n${w()}\n\n-->`],
    ['comment-fence', w => `<!--\n\`\`\`\n${w()}\n\`\`\`\n-->`], ['comment-admonition', w => `<!--\n:::note\n${w()}\n:::\n-->`],
    ['comment-closed-in-piece', w => `<!-- x\n\n${w()} --> ${w()}`], ['comment-closed-by-bang', w => `<!-- x\n\n${w()} --!> ${w()}`],
    // 3, 4 and 5: a processing instruction, a declaration, CDATA.
    ['processing-instruction', () => '<?php echo 1; ?>'], ['declaration', () => '<!X decl>'], ['cdata', w => `<![CDATA[${w()}]]>`],
    ['cdata-pieces', w => `<![CDATA[x\n\n${w()}\n\n]]>`],
    // 6: block-level tags, opening, closing and holding text.
    ['div', () => '<div>'], ['div-close', () => '</div>'], ['div-align', () => '<div align="center">'], ['div-text', w => `<div>${w()}</div>`],
    ['p', w => `<p>${w()}</p>`], ['p-open', () => '<p>'], ['p-close', () => '</p>'], ['center', () => '<center>'], ['center-close', () => '</center>'],
    ['details', w => `<details>\n<summary>${w()}</summary>`], ['details-close', () => '</details>'], ['summary', w => `<summary>${w()}</summary>`],
    ['h2', w => `<h2>${w()}</h2>`], ['ul', w => `<ul><li>${w()}</li></ul>`], ['blockquote', w => `<blockquote>${w()}</blockquote>`],
    ['nested-table', w => `<table><tr><td>${w()}</td></tr></table>`], ['nested-table-pieces', w => `<table><tr><td>\n\n${w()} *${w()}*\n\n</td></tr></table>`],
    ['next-cell', () => '</td><td>'], ['next-row', () => '</td></tr><tr><td>'],
    // 7: other tags, alone and with text, and a link around a piece.
    ['span', () => '<span>'], ['span-text', w => `<span>${w()}</span>`], ['kbd', w => `<kbd>${w()}</kbd> ${w()}`], ['em-close', () => '</em>'],
    ['bold-open', () => '<b>'], ['bold-close', () => '</b>'], ['bold-text', w => `<b>${w()}</b> *${w()}*`],
    ['link-around', w => `<a href="http://x/${w()}">\n\n${w()} *${w()}*\n\n</a>`], ['template-pieces', w => `<template>\n\n${w()}\n\n</template>`],
    // Pieces in a quoted value or a tag, which a later piece closes.
    ['alt-pieces', w => `<img alt="\n\n${w()} *${w()}*\n\n">`], ['alt-admonition', w => `<img alt="\n\n:::note\n${w()}\n:::\n\n">`],
    ['tag-pieces', w => `<span\n\nhidden\n\n>${w()}</span>`], ['title-pieces', w => `<td title="\n\nhidden\n\n">${w()}`],
    // Lifted blocks.
    ['fence', w => `\`\`\`\n${w()}\n\`\`\``], ['fence-tilde', w => `~~~js\n${w()}\n~~~`], ['fence-unclosed', () => '```'], ['math', w => `$$\n${w()}\n$$`],
    ['admonition', w => `:::note\n${w()}\n:::`], ['admonition-pandoc', w => `::: {.warning}\n${w()}\n:::`], ['admonition-open', () => ':::note'],
    ['admonition-close', () => ':::'], ['admonition-unknown', w => `:::foo\n${w()}\n:::`],
    // Setext underlines, rules and headings.
    ['setext-equals', () => '==='], ['setext-dash', () => '---'], ['setext-text', w => `${w()}\n---`],
    ['rule-stars', () => '***'], ['rule-underscores', () => '___'], ['rule-spaced', () => '- - -'], ['atx', w => `# ${w()}`], ['atx-empty', () => '#'],
    // List items, quotes, indentation and the rest.
    ['item-dash', w => `- ${w()}`], ['item-ordered', w => `1. ${w()}`], ['item-paren', w => `2) ${w()}`], ['item-star', w => `* ${w()}`], ['item-empty', () => '-'],
    ['quote', w => `> ${w()}`], ['quote-empty', () => '>'], ['quote-div', w => `> <div>${w()}</div>`],
    ['indented-code', w => `    ${w()}`], ['tab-code', w => `\t${w()}`], ['indent-two', w => `  ${w()}`],
    ['pipe-table', w => `| ${w()} | b |\n| - | - |\n| ${w()} | 2 |`], ['definition', w => `${w()}\n: ${w()}`], ['embed', () => '::youtube{id=dQw4w9WgXcQ}'],
    ['anchor', w => `<a id="${w()}"></a>`], ['wrapper-close-text', w => `</div>${w()}`], ['less-than-text', w => `<3 ${w()}`], ['code-span-td', w => `\`<td>\` ${w()}`],
];

/** What follows `</table>`. */
const TAILS: ((w: () => string) => string)[] = [() => '', w => `\n${w()}`, () => '\n---', () => '\n===', w => `\n${w()}\n===`, () => '\n***', w => `\n# ${w()}`,
    w => `\n- ${w()}`, w => `\n> ${w()}`, w => `\n<div>${w()}</div>`, w => `\n\n${w()}`, w => `\n    ${w()}`, w => `\n:::note\n${w()}\n:::`];

/** What stands between the parts of a random cell: a line break, a blank line, or a line of a space. */
const SEPARATORS = ['\n', '\n\n', '\n\n', '\n \n'];

/** How the table opens (an aligned `<div>` around it is closed after it). */
const OPENINGS = ['<table><tr><td>', '<table>\n<tr>\n<td>', '<table><tr><th>', '<table><tbody><tr><td>', '<div align="center">\n<table><tr><td>', '<table>\n<caption>c</caption>\n<tr><td>'];

/**
 * A table whose cell holds the Markdown piece `wN *wN*` beside `starts` (each after its separator but the first),
 * `beside` between the two, `lead` after the opening and `trail` before `</td>`, then `tail` (see the file comment).
 */
function cellDocument(opening: string, starts: { make: BlockStart[1]; separator: string }[], pieceFirst: boolean,
    space: { lead: string; beside: string; trail: string }, tail: (w: () => string) => string): string {
    let n = 0;
    const w = () => `w${n++}`;
    const body = starts.map((start, k) => (k ? start.separator : '') + start.make(w)).join('');
    const piece = `${w()} *${w()}*`;
    const cell = pieceFirst ? piece + space.beside + body : body + space.beside + piece;
    const wrapped = opening.startsWith('<div');
    return opening + space.lead + cell + space.trail + '</td></tr></table>' + tail(w) + (wrapped ? '\n</div>' : '');
}

const wordsOf = (s: string): string[] => s.match(/(?<![A-Za-z0-9])w\d+(?![0-9])/g) ?? [];

/** The text, metadata strings (in arrays too) and notes of an AST's nodes. */
function stringsOf(nodes: OfficeContentNode[] | undefined, out: string[]): string[] {
    for (const node of nodes ?? []) {
        if (typeof node.text === 'string') out.push(node.text);
        for (const value of Object.values((node.metadata ?? {}) as Record<string, unknown>)) {
            if (typeof value === 'string') out.push(value);
            else if (Array.isArray(value)) for (const item of value) if (typeof item === 'string') out.push(item);
        }
        stringsOf(node.children, out);
        stringsOf(node.notes, out);
    }
    return out;
}

/** A token or placeholder: a private-use character, or a placeholder's name. */
const LEAK = /[\uE000-\uF8FF]|(?:CODE_BLOCK|MATH_BLOCK|ADMONITION|HTML_COMMENT|RAW_HTML|LIFTED_TEXT|LIFTED_BLOCKS|DISCLOSURE|JOINED_TABLE)_\d/;

/** An item of the quote part, by its kind; `w` gives the next marker word. */
type QuoteItem = [kind: string, make: (w: () => string) => string[]];
const QUOTE_PIECES: QuoteItem[] = [
    ['piece', w => [`${w()} *${w()}*`]], ['piece-apostrophe', w => [`${w()}'s **${w()}**`]], ['piece-quotes', w => [`It's ${w()} "${w()}" *${w()}*`]],
    ['piece-odd-quote', w => [`${w()} it's "${w()}`]], ['piece-gt', w => [`${w()} > ${w()}`]], ['piece-list', w => [`- ${w()}'s\n- ${w()}`]],
    ['piece-link-title', w => [`[${w()}](http://x/ "t's")`]], ['piece-comment-end', w => [`${w()} --> ${w()}`]], ['piece-cdata-end', w => [`${w()} ]]> ${w()}`]],
];
/** HTML holding a quote that opens: `uncertain` when a `>` stands inside it (the reader may end its tag there), `open` when it never closes. */
type QuoteOpener = [kind: string, make: (w: () => string) => string[], how?: 'uncertain' | 'open'];
const QUOTE_OPENERS: QuoteOpener[] = [
    ['unquoted-apostrophe', w => [`<img alt=Bob's src=a.png> ${w()}`], 'uncertain'], ['cell-apostrophe', w => [`<td title=it's> ${w()}`], 'uncertain'],
    ['missing-close-quote', w => [`<span class="x>${w()}</span>`], 'uncertain'], ['apostrophe-last', w => [`<b>${w()}</b><img src=a.png alt=it's>`], 'uncertain'],
    ['alt-open', () => ['<img alt="'], 'open'], ['alt-closed-by-tag', w => ['<img alt="', `<i>">${w()}`]],
    ['alt-closed-by-piece', w => ['<img alt="', `">${w()} *${w()}*`]], ['alt-text-closed', w => [`<img alt="${w()}`, `<i>"> ${w()}`]],
];
/** HTML that opens nothing it leaves open. */
const QUOTE_PLAIN: QuoteItem[] = [
    ['comment', w => [`<!-- ${w()} -->`]], ['cdata', w => [`<![CDATA[${w()}]]>`]], ['bold', w => [`<b>${w()}</b>`]], ['next-cell', () => ['</td><td>']],
    ['dashes', w => [`<i>${w()}</i> -- ${w()}`]],
];
/** A comment or CDATA a later item closes: a piece, HTML, or nothing (open to the table's end, its own text the word `hidden`). */
const QUOTE_SPANS: [kind: string, open: string, closers: [kind: string, make: ((w: () => string) => string) | null][]][] = [
    ['comment-span', '<!-- x', [['by-piece', w => `${w()} --> ${w()}`], ['by-tag', w => `<i>${w()} -->`], ['by-bang', w => `<i>--!> ${w()}`], ['open', null]]],
    ['cdata-span', '<![CDATA[x', [['by-piece', w => `${w()} ]]> ${w()}`], ['by-tag', w => `<i>]]> ${w()}`], ['open', null]]],
];
/** What stands before the table in its block; `uncertain` for a wrapper whose tag holds an unpaired quote and a `>` in it. */
const QUOTE_LEADS: [kind: string, make: (w: () => string) => string, uncertain?: boolean][] = [
    ['none', () => ''], ['code-span-comment', w => `Use \`<!--\` here ${w()}\n`], ['code-span-cdata', w => `Use \`<![CDATA[\` here ${w()}\n`],
    ['code-span-quote', w => `It's \`"\` ${w()}\n`], ['wrapper-missing-quote', () => '<p class="x>\n', true], ['wrapper-apostrophe', () => "<div title=it's>\n", true],
    ['paragraph-quotes', w => `${w()}'s "${w()}\n\n`], ['comment-line', w => `<!-- ${w()} -->\n`],
];

/**
 * One document of the quote part (see the file comment), drawn with `rnd`: a lead, then a table whose cell holds one
 * to five items joined by blank lines, then a marker paragraph. `drawn` collects the kinds it used.
 */
function quoteDocument(rnd: () => number, drawn: Set<string>): string {
    let n = 0;
    const w = () => `w${n++}`;
    const pick = <T>(a: T[]): T => a[Math.floor(rnd() * a.length)];
    const [leadKind, lead, leadUncertain] = pick(QUOTE_LEADS);
    drawn.add(`lead:${leadKind}`);
    const cell: string[] = [];
    // Whether a quote whose tag a `>` in it may end was drawn, whether a quote opened in markup, and whether
    // something is left open to the table's end (only pieces follow it).
    let uncertain = !!leadUncertain;
    let quoted = uncertain;
    let open = false;
    const pieces = (count: number) => {
        for (let j = 0; j < count; j++) {
            const [kind, make] = pick(QUOTE_PIECES);
            drawn.add(kind);
            cell.push(...make(w));
        }
    };
    const head = lead(w);
    for (let k = 1 + Math.floor(rnd() * 5); k > 0 && !open; k--) {
        const r = rnd();
        if (r < 0.45) pieces(1);
        else if (r < 0.6 && !quoted) {
            const [kind, make, how] = pick(QUOTE_OPENERS);
            drawn.add(kind);
            quoted = true;
            uncertain ||= how === 'uncertain';
            open = how === 'open';
            cell.push(...make(w));
        } else if (r < 0.8) {
            const [kind, make] = pick(QUOTE_PLAIN);
            drawn.add(kind);
            cell.push(...make(w));
        } else {
            const [kind, opener, closers] = pick(QUOTE_SPANS);
            const [closerKind, closer] = pick(closers.filter(([name]) => !uncertain || name !== 'by-piece'));
            drawn.add(`${kind}:${closerKind}`);
            cell.push(opener.replace('x', closer ? w() : 'hidden'));
            pieces(Math.floor(rnd() * 3));
            if (closer) cell.push(closer(w));
            else open = true;
        }
    }
    if (open) pieces(Math.floor(rnd() * 3));
    return head + '<table><tr><td>\n\n' + cell.join('\n\n') + '\n\n</td></tr></table>\n\n' + w();
}
/** Every kind the quote part can draw. */
const QUOTE_KINDS = [...QUOTE_LEADS.map(([kind]) => `lead:${kind}`), ...[...QUOTE_PIECES, ...QUOTE_OPENERS, ...QUOTE_PLAIN].map(([kind]) => kind),
    ...QUOTE_SPANS.flatMap(([kind, , closers]) => closers.map(([closer]) => `${kind}:${closer}`))];

/** Reads the documents (see the file comment) with `read`, the random ones drawn from `seed`, and counts the failures. */
export async function checkMarkdownCellPieces(read: (markdown: string) => Promise<OfficeParserAST>, seed: number): Promise<MarkdownCellPieces> {
    const result: MarkdownCellPieces = { systematic: 0, random: 0, distinct: 0, quotes: 0, quotesDistinct: 0, leaked: 0, lost: 0, undrawn: [], failures: [] };
    const check = async (source: string): Promise<void> => {
        const all = stringsOf((await read(source)).content, []).join('\n');
        const leaked = LEAK.test(all);
        const have = new Set(wordsOf(all));
        const missing = wordsOf(source).filter(word => !have.has(word));
        if (leaked) result.leaked++;
        if (missing.length) result.lost++;
        if ((leaked || missing.length) && result.failures.length < 3) {
            result.failures.push(`${leaked ? 'a token or placeholder shown' : `lost ${missing.join(' ')}`} reading ${JSON.stringify(source)}: ${JSON.stringify(all.replace(/[\uE000-\uF8FF]/g, '\u00A7'))}`);
        }
    };
    for (const [, make] of BLOCK_STARTS) {
        for (const tail of TAILS) {
            for (const separator of ['\n', '\n\n']) {
                for (const pieceFirst of [false, true]) {
                    await check(cellDocument(OPENINGS[0], [{ make, separator: '' }], pieceFirst, { lead: '\n\n', beside: separator, trail: '\n\n' }, tail));
                    result.systematic++;
                }
            }
        }
    }
    const rnd = (): number => { seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff; return seed / 0x80000000; };
    const pick = <T>(a: T[]): T => a[Math.floor(rnd() * a.length)];
    const seen = new Set<string>();
    const drawn = new Set<string>();
    for (let k = 0; k < 3000; k++) {
        const opening = pick(OPENINGS);
        const starts = Array.from({ length: 1 + Math.floor(rnd() * 4) }, () => {
            const [kind, make] = pick(BLOCK_STARTS);
            drawn.add(kind);
            return { make, separator: pick(SEPARATORS) };
        });
        const source = cellDocument(opening, starts, rnd() < 0.5, { lead: pick(SEPARATORS), beside: '\n\n', trail: pick(SEPARATORS) }, pick(TAILS));
        result.random++;
        if (seen.has(source)) continue;
        seen.add(source);
        await check(source);
    }
    result.distinct = seen.size;
    const quoteSeen = new Set<string>();
    const quoteDrawn = new Set<string>();
    for (let k = 0; k < 3000; k++) {
        const source = quoteDocument(rnd, quoteDrawn);
        result.quotes++;
        if (quoteSeen.has(source)) continue;
        quoteSeen.add(source);
        await check(source);
    }
    result.quotesDistinct = quoteSeen.size;
    result.undrawn = [...BLOCK_STARTS.map(([kind]) => kind).filter(kind => !drawn.has(kind)), ...QUOTE_KINDS.filter(kind => !quoteDrawn.has(kind))];
    return result;
}

import { AdmonitionMetadata, CellMetadata, CodeMetadata, CommentMetadata, EmbedMetadata, FullOfficeParserConfig, HeadingMetadata, ImageMetadata, ListMetadata, OfficeAttachment, OfficeContentNode, OfficeErrorType, OfficeMetadata, OfficeParserAST, ParagraphMetadata, TableMetadata, TextFormatting, TextMetadata } from '../types.js';
import { anchorMark, isAnchorMark, resolveAnchorMarks } from '../utils/anchorUtils.js';
import { createAST } from '../utils/astUtils.js';
import { isSourceComment } from '../utils/commentUtils.js';
import { checkAbortSignal, getOfficeError } from '../utils/errorUtils.js';
import { CHARACTER_REFERENCE, decodeCharacterReference } from '../utils/htmlEntities.js';
import { isEmptyMath, MathNode, mathmlTreeToLatex } from '../utils/mathUtils.js';
import { isSafeHtmlAttributeName, iframeAllowed } from '../utils/sanitize.js';
import { setOwn } from '../utils/lookupUtils.js';
import { cellSpan, MAX_COL_SPAN, MAX_ROW_SPAN } from '../utils/numberUtils.js';
import { appendAll } from '../utils/nodeListUtils.js';

/**
 * Maximum element nesting depth accepted from an HTML/XHTML source before the parser gives up
 * with a typed error rather than letting the recursion overflow the call stack. See the guard in
 * `parseNode` for why this value and not a larger one.
 */
const MAX_HTML_NESTING_DEPTH = 256;

interface HtmlNode {
    /** `comment` = an HTML comment, kept only under `HtmlParserConfig.preserveComments`; `text` is its body. */
    type: 'element' | 'text' | 'comment';
    tagName?: string;
    attributes?: Record<string, string>;
    text?: string;
    children: HtmlNode[];
    parent?: HtmlNode;
    /** Elements between this one and the root, set while the tree is built. */
    depth?: number;
    /** Whether the element is open while the tree is built (see parseHtmlTree's formatting elements). */
    open?: boolean;
    /** Whether opening it started a new run of formatting to carry (a table cell, a caption, an object). */
    marksFormatting?: boolean;
}

/** What the tree builder gives: the tree, its <head> element, and the attributes of <html> and <body>. */
interface HtmlTree {
    /**
     * The document: what a browser puts in <body>, and <head> among it. The <html> and <body> start tags
     * build no element (a browser has one of each, whatever the markup), so content before or after
     * them, which a browser moves into <body>, is not lost outside it.
     */
    root: HtmlNode;
    head?: HtmlNode;
    htmlAttributes: Record<string, string>;
    bodyAttributes: Record<string, string>;
}

/**
 * HTML's "special" elements: an end tag of another element that would close one of these is ignored,
 * and a list item or definition ends the one before it only when no other of them stands between.
 */
const SPECIAL_ELEMENTS = new Set(['address', 'applet', 'area', 'article', 'aside', 'base', 'basefont', 'bgsound', 'blockquote', 'body', 'br',
    'button', 'caption', 'center', 'col', 'colgroup', 'dd', 'details', 'dialog', 'dir', 'div', 'dl', 'dt', 'embed', 'fieldset', 'figcaption',
    'figure', 'footer', 'form', 'frame', 'frameset', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'head', 'header', 'hgroup', 'hr', 'html', 'iframe',
    'img', 'input', 'keygen', 'li', 'link', 'listing', 'main', 'marquee', 'menu', 'meta', 'nav', 'noembed', 'noframes', 'noscript', 'object',
    'ol', 'p', 'param', 'plaintext', 'pre', 'script', 'search', 'section', 'select', 'source', 'style', 'summary', 'table', 'tbody', 'td',
    'template', 'textarea', 'tfoot', 'th', 'thead', 'title', 'tr', 'track', 'ul', 'wbr', 'xmp']);

/** Special elements a list item or definition is ended across (`<li><div><li>` ends the first item). */
const ITEM_TRANSPARENT = new Set(['address', 'div', 'p']);

/**
 * The elements an end tag (or an implied end) does not reach past: HTML's scopes. An end tag whose
 * element has one of these open inside it is ignored, so `</b>` in a table cell does not end a `<b>`
 * outside the table, and `</td>` in a table nested in a cell does not end the outer cell.
 */
const DEFAULT_SCOPE = ['applet', 'caption', 'marquee', 'object', 'table', 'td', 'template', 'th'];
const BUTTON_SCOPE = [...DEFAULT_SCOPE, 'button'];
const LIST_ITEM_SCOPE = [...DEFAULT_SCOPE, 'ol', 'ul'];
const TABLE_SCOPE = ['table', 'template'];

/** Elements whose end tag is read in table scope (it does not reach past a table nested inside). */
const TABLE_SCOPED_ENDS = new Set(['caption', 'colgroup', 'table', 'tbody', 'td', 'tfoot', 'th', 'thead', 'tr']);

/** A table's parts: outside a table, a browser reads their tags as nothing (what they hold stays). */
const TABLE_PARTS = new Set(['caption', 'col', 'colgroup', 'tbody', 'td', 'tfoot', 'th', 'thead', 'tr']);

/** Start tags that end an open paragraph (`<p>a<div>` is a paragraph, then a division). */
const CLOSES_PARAGRAPH = new Set(['address', 'article', 'aside', 'blockquote', 'center', 'details', 'dialog', 'dir', 'div', 'dl', 'fieldset',
    'figcaption', 'figure', 'footer', 'form', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'header', 'hgroup', 'hr', 'li', 'dd', 'dt', 'listing', 'main',
    'menu', 'nav', 'ol', 'p', 'pre', 'search', 'section', 'summary', 'table', 'ul', 'xmp']);

/** What <head> holds: any other start tag (or text) ends it, as a browser reads it. */
const HEAD_CONTENT = new Set(['base', 'basefont', 'bgsound', 'link', 'meta', 'noframes', 'noscript', 'script', 'style', 'template', 'title']);

/** Elements that hold nothing, whether or not their tag is written self-closed. */
const VOID_ELEMENTS = new Set(['area', 'base', 'basefont', 'bgsound', 'br', 'col', 'embed', 'hr', 'img', 'input', 'keygen', 'link', 'meta',
    'param', 'source', 'track', 'wbr']);

/**
 * Elements whose content is text up to their end tag, never markup: raw text (`<script>`, `<style>`,
 * `<xmp>`...; `<noscript>` too, as a browser running scripts reads it) and, with its character
 * references, `<title>` and `<textarea>`. Read as markup, `<textarea><b>x</b></textarea>` was bold text.
 */
const RAW_TEXT_ELEMENTS = new Set(['iframe', 'noembed', 'noframes', 'noscript', 'script', 'style', 'xmp']);
const ESCAPABLE_RAW_TEXT_ELEMENTS = new Set(['textarea', 'title']);

/**
 * HTML's formatting elements. One a paragraph's end closes is opened again where the text goes on, as
 * a browser does (`<p><b>bold<p>still bold</b>`): it is carried, until its end tag, within the cell,
 * caption or object it was opened in.
 */
const FORMATTING_ELEMENTS = new Set(['a', 'b', 'big', 'code', 'em', 'font', 'i', 'nobr', 's', 'small', 'strike', 'strong', 'tt', 'u']);
const FORMATTING_BOUNDARIES = new Set(['applet', 'caption', 'marquee', 'object', 'td', 'template', 'th']);

/**
 * The most formatting elements carried at once (the oldest is let go past it): each is opened again
 * wherever text follows its implied end, so the number bounds the elements that a few bytes of markup
 * make. Browsers carry any number; no document keeps more than a few open.
 */
const MAX_CARRIED_FORMATTING = 16;

/** Start tags before which carried formatting is not opened again (blocks, table parts, head content, raw text). */
const OPENS_NO_FORMATTING = new Set([...CLOSES_PARAGRAPH, ...TABLE_PARTS, ...HEAD_CONTENT, 'frame', 'frameset', 'head', 'iframe', 'noembed',
    'param', 'plaintext', 'source', 'textarea', 'track']);

/**
 * What HtmlGenerator writes in a picture's caption: nothing, or the picture's attachment name (a file
 * name, whatever its extension: `image1.png`, `image1.tmp`). Anything else in a caption is the
 * document's own text, and kept.
 */
const WRITER_CAPTION = /^(?:[^\s<>]*[^\s<>.]\.[A-Za-z0-9]{1,8})?$/;

/**
 * The cells HtmlGenerator adds around a sheet's data, as a spreadsheet shows them: the row numbers,
 * the column letters, and the corner between them. They are not the sheet's content.
 */
const SHEET_CHROME_CLASSES = new Set(['excel-row-num', 'excel-row-num-header', 'excel-col-header']);
const isSheetChrome = (node: HtmlNode): boolean => node.type === 'element' && (node.tagName === 'td' || node.tagName === 'th')
    && (node.attributes?.class || '').split(/\s+/).some(name => SHEET_CHROME_CLASSES.has(name));

/** Node types that are blocks, wherever they stand: a caption holding one is not wrapped in a paragraph. */
const BLOCK_NODE_TYPES = new Set<string>(['paragraph', 'heading', 'list', 'table', 'chart', 'embed', 'admonition', 'definitionList']);

/** Whether `node` is a block: one of BLOCK_NODE_TYPES, a code block or display equation, or a rule or page break. */
const isBlockNode = (node: OfficeContentNode): boolean => BLOCK_NODE_TYPES.has(node.type)
    || (node.type === 'code' && (node.metadata as CodeMetadata | undefined)?.math !== 'inline')
    || (node.type === 'break' && ['thematic', 'page'].includes((node.metadata as { breakType?: string } | undefined)?.breakType ?? ''));

/** Elements that are blocks in HTML's layout: what one holds is never part of the text around it. */
const BLOCK_LEVEL_TAGS = new Set(['address', 'article', 'aside', 'blockquote', 'center', 'details', 'dialog', 'dd', 'div', 'dl', 'dt', 'fieldset',
    'figcaption', 'figure', 'footer', 'form', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'header', 'hgroup', 'hr', 'li', 'main', 'menu', 'nav', 'ol', 'p',
    'pre', 'section', 'summary', 'table', 'ul']);

/**
 * Decode the character references in a text node. Text nodes are kept in their raw escaped form
 * during parsing, so any branch that lifts text into AST content has to decode first: `&lt;` inside
 * a code/math body is a less-than operator, not markup. Attribute values are decoded once, as they
 * are read (see `parseAttributes`).
 *
 * One pass, so it is the exact inverse of `escapeHtml`: chained replaces decoded `&amp;` before
 * `&quot;`/`&#39;`, turning the escaped literal text `&amp;quot;` into `"` instead of `&quot;`. Every
 * numeric reference and HTML 4's named references decode (`&rsquo;`, `&#8217;`, `&copy;`), not only
 * the few `escapeHtml` writes. A number past the last code point (`&#1114112;`) is U+FFFD, as HTML
 * reads it (NUL and surrogates are too, see decodeCharacterReference).
 */
const decodeEntities = (text: string): string =>
    text.replace(CHARACTER_REFERENCE, (full: string, body: string) => decodeCharacterReference(body) ?? (body[0] === '#' ? '\uFFFD' : full));

/** An element's raw child text as the source had it (code/math bodies). A comment is not text. */
const rawChildText = (node: HtmlNode): string => node.children.map(c => (c.type === 'comment' ? '' : c.text || '')).join('');

/**
 * The raw text of preformatted content: its text at every depth (the `<span>`s a syntax highlighter
 * wraps each token in) and a line break for each `<br>`. Reading the direct text alone dropped every
 * highlighted token and ran the lines together.
 */
const preformattedText = (node: HtmlNode): string => {
    let out = '';
    const walk = (n: HtmlNode, depth: number): void => {
        for (const c of n.children) {
            if (c.type === 'comment') continue;
            if (c.type === 'element') {
                if (c.tagName === 'br') out += '\n';
                else if (depth < MAX_HTML_NESTING_DEPTH) walk(c, depth + 1);
            } else {
                out += c.text || '';
            }
        }
    };
    walk(node, 0);
    return out;
};

/**
 * Elements not laid out inline: blocks, a table's parts, and what is not shown (the document's head,
 * scripts). Every other element is inline (`<math>`, `<button>`, `<nobr>`, Word's `<o:p>`, a custom
 * element): whitespace between two inline elements, or between one and text, is a visible space, where
 * whitespace between blocks is only layout. A list of the inline elements instead left out every
 * element it did not name, and ran `<em>word</em> <math>` together.
 */
const NOT_INLINE_ELEMENTS = new Set([...BLOCK_LEVEL_TAGS, ...TABLE_PARTS, ...HEAD_CONTENT, 'root', 'head', 'dir', 'frame', 'frameset',
    'iframe', 'legend', 'listing', 'noembed', 'optgroup', 'option', 'param', 'plaintext', 'search', 'source', 'template', 'track', 'xmp']);

/** Elements never shown (the document's head, scripts, templates): read as nothing wherever they stand. */
const HIDDEN_ELEMENTS = new Set([...HEAD_CONTENT, 'head', 'noembed']);

/** Whether `node` is an element laid out inline (see NOT_INLINE_ELEMENTS); display MathML is a block. */
const isInlineElement = (node: HtmlNode): boolean => node.type === 'element' && !NOT_INLINE_ELEMENTS.has(node.tagName ?? '')
    && !(node.tagName === 'math' && (node.attributes?.display === 'block' || node.attributes?.mode === 'display'));

/**
 * Collapses a text node's whitespace as HTML renders it: each run of ASCII whitespace becomes one
 * space. A no-break space (`&nbsp;`) is content, not layout, and stays.
 */
const collapseWhitespace = (text: string): string => text.replace(/[ \t\n\f\r]+/g, ' ');

/**
 * Inline content with each space that follows another space (in the text node before it) removed,
 * as HTML collapses whitespace across element boundaries: `<b>a </b> b` shows one space, not
 * three. A text node left empty is dropped unless it carries something (a note, a comment).
 */
const collapseSpacesAcrossNodes = (nodes: OfficeContentNode[]): OfficeContentNode[] => {
    const out: OfficeContentNode[] = [];
    let previous: OfficeContentNode | undefined;
    for (const node of nodes) {
        let current = node;
        if (current.type === 'text' && current.text?.startsWith(' ') && previous?.type === 'text' && previous.text?.endsWith(' ')) {
            current = { ...current, text: current.text.slice(1) };
            if (!current.text && !current.notes?.length && !current.comments?.length) continue;
        }
        out.push(current);
        // A named anchor takes no room: the spaces either side of it are one.
        if (!isAnchorMark(current)) previous = current;
    }
    return out;
};

/**
 * The text of `parent.children[i]` (a text node) as HTML shows it: its whitespace collapsed, the space
 * at a line's end or start around a `<br>` dropped, and whitespace alone kept only between two pieces of
 * inline content (`<code>a</code> <code>b</code>` is two spans with a space, not one), since anywhere
 * else it is layout. That holds for a text of only no-break spaces too (a spacer paragraph's
 * `&nbsp;`); between inline content it is kept as it is. Null when none of it shows.
 */
function visibleText(parent: HtmlNode, i: number): string | null {
    // The sibling next to children[i] in direction `step`, passing over comments and what is not shown
    // (`word <script>...</script> <b>more</b>` shows one space between the words, not none).
    const sibling = (step: number): HtmlNode | undefined => {
        let j = i + step;
        while (parent.children[j] && (parent.children[j].type === 'comment' || HIDDEN_ELEMENTS.has(parent.children[j].tagName ?? ''))) j += step;
        return parent.children[j];
    };
    const isBreak = (sib: HtmlNode | undefined) => sib?.type === 'element' && sib.tagName === 'br';
    // Inline content: text, an inline element, or (at the edge of an inline element such as `<b> </b>`)
    // whatever lies beyond it.
    const isInline = (sib: HtmlNode | undefined) => sib === undefined ? isInlineElement(parent) : sib.type === 'text' || isInlineElement(sib);
    let text = collapseWhitespace(decodeEntities(parent.children[i].text || ''));
    if (isBreak(sibling(-1))) text = text.replace(/^ /, '');
    if (isBreak(sibling(1))) text = text.replace(/ $/, '');
    if (!text || (!text.trim() && !(isInline(sibling(-1)) && isInline(sibling(1))))) return null;
    return text;
}

/**
 * A block's inline content without the space at its start and at its end, which HTML does not show
 * (a line starts and ends where its text does). Only the text at either edge is touched: an empty
 * text node that carries a note or comment is passed over, and one left empty is dropped.
 */
const trimBlockEdges = (nodes: OfficeContentNode[]): OfficeContentNode[] => {
    const out = nodes.slice();
    const trimmed = new Set<number>();
    for (const [start, step, edge] of [[0, 1, /^ /], [out.length - 1, -1, / $/]] as const) {
        for (let i = start; i >= 0 && i < out.length; i += step) {
            const node = out[i];
            if (isAnchorMark(node)) continue;
            if (node.type !== 'text') break;
            if (!node.text) continue;
            if (edge.test(node.text)) {
                out[i] = { ...node, text: node.text.replace(edge, '') };
                trimmed.add(i);
            }
            break;
        }
    }
    return out.filter((node, i) => !trimmed.has(i) || node.text || node.notes?.length || node.comments?.length);
};

/** trimBlockEdges on `nodes` itself: its edges are all it touches, where a copy of the list costs its whole length. */
const trimBlockEdgesInPlace = (nodes: OfficeContentNode[]): void => {
    for (const [start, step, edge] of [[0, 1, /^ /], [-1, -1, / $/]] as const) {
        for (let i = start < 0 ? nodes.length - 1 : start; i >= 0 && i < nodes.length; i += step) {
            const node = nodes[i];
            if (isAnchorMark(node)) continue;
            if (node.type !== 'text') break;
            if (!node.text) continue;
            if (edge.test(node.text)) {
                const trimmed = { ...node, text: node.text.replace(edge, '') };
                if (trimmed.text || trimmed.notes?.length || trimmed.comments?.length) nodes[i] = trimmed;
                else nodes.splice(i, 1);
            }
            break;
        }
    }
};

/**
 * `nodes` with a named anchor for `ids` before them: `ids` joins a mark already first (an element
 * inside marked it) rather than the whole list being copied behind a new one at every level.
 */
const markedBefore = (ids: string[], nodes: OfficeContentNode[]): OfficeContentNode[] => {
    if (ids.length === 0) return nodes;
    const first = nodes[0];
    if (first && isAnchorMark(first)) {
        const meta = first.metadata as { anchorIds?: string[] };
        meta.anchorIds = [...ids, ...(meta.anchorIds ?? [])];
        return nodes;
    }
    nodes.unshift(anchorMark(ids));
    return nodes;
};

/** Holds a list item's place at the start of its nested lists until the item itself is built. */
const selfNodePlaceholder: OfficeContentNode = { type: 'text', text: '' };

/** Plain text of parsed content nodes, leaving out source comments: a hidden note is not text. */
const plainTextOf = (nodes: OfficeContentNode[]): string => nodes.map(n => (isSourceComment(n) ? '' : n.text || '')).join('');

/** plainTextOf at every depth: the text of a note holding paragraphs (GitHub's), which has none of its own. */
const nodesText = (nodes: OfficeContentNode[], depth = 0): string => nodes.map(n => (isSourceComment(n) ? ''
    : n.children?.length && depth < MAX_HTML_NESTING_DEPTH ? nodesText(n.children, depth + 1) : n.text || '')).join('');

/**
 * Presents an `HtmlNode` as a `MathNode` for the shared MathML converter.
 *
 * The shapes already line up field for field; the one thing that must happen here is entity
 * decoding, since this parser keeps text nodes in their raw escaped form and `&lt;` inside an
 * `<mo>` is a less-than operator, not markup.
 */
const toMathNode = (node: HtmlNode): MathNode => ({
    tagName: node.tagName,
    attributes: node.attributes,
    text: node.text === undefined ? undefined : decodeEntities(node.text),
    children: (node.children || []).map(toMathNode),
});

const parseAttributes = (attrString: string): Record<string, string> => {
    // Null-prototype: keyed by the document's attribute names, as every such map is.
    const attrs: Record<string, string> = Object.create(null);
    // Attribute names follow the HTML5 rule - any character except whitespace and
    // " ' > / = - rather than a hand-picked allowlist. The previous class
    // ([a-zA-Z0-9\-:]) silently split a legal name on any character outside it, so
    // `data_foo="x"` produced TWO attributes: `data` (empty) and an invented
    // `foo="x"` that was never in the source. Harmless while nothing read unknown
    // attributes; not harmless once they can be replayed into generated output.
    const regex = /([^\s"'>/=]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+)))?/g;
    let match;
    while ((match = regex.exec(attrString)) !== null) {
        const name = match[1].toLowerCase();
        const value = match[2] !== undefined ? match[2] : (match[3] !== undefined ? match[3] : (match[4] || ''));
        // Decoded here, once, so every reader gets the value the author meant (`href="?a=1&amp;b=2"`
        // is `?a=1&b=2`, `alt="Tom &amp; Jerry"` is `Tom & Jerry`), and none decodes it again.
        attrs[name] = decodeEntities(value);
    }
    return attrs;
};

/**
 * Splits an inline `style` attribute into a property -> value map.
 *
 * Replaces substring matching (`styleAttr.includes('font-weight: bold')`) and unanchored regexes
 * (`/color:\s*([^;]+)/`), which were wrong in both directions:
 *   - false positives: `color:` matched inside `background-color:`, so
 *     `"background-color: red; color: blue"` yielded color=red - the *wrong* value, not merely a
 *     spurious one - and `width:` matched inside `max-width:`, so the ubiquitous responsive-image
 *     style `max-width: 100%` was read as an explicit width.
 *   - false negatives: `font-weight:bold` without a space, `font-weight: 700`, and `bolder` were
 *     all missed, as was `line-through` inside `text-decoration: underline line-through`.
 *
 * Splitting is quote- and paren-aware so a semicolon inside `url(data:image/png;base64,...)` or a
 * quoted font stack doesn't shatter the declaration. `!important` is stripped from values, since
 * substring matching used to tolerate it and exact comparison otherwise would not - dropping it
 * would be a silent regression rather than the intended fix.
 */
const parseStyleDeclarations = (styleAttr: string): Map<string, string> => {
    const decls = new Map<string, string>();
    if (!styleAttr) return decls;

    let depth = 0;
    let quote: '"' | '\'' | null = null;
    let current = '';
    const chunks: string[] = [];
    for (const ch of styleAttr) {
        if (quote) {
            if (ch === quote) quote = null;
        } else if (ch === '"' || ch === '\'') {
            quote = ch;
        } else if (ch === '(') {
            depth++;
        } else if (ch === ')') {
            if (depth > 0) depth--;
        } else if (ch === ';' && depth === 0) {
            chunks.push(current);
            current = '';
            continue;
        }
        current += ch;
    }
    chunks.push(current);

    for (const chunk of chunks) {
        const idx = chunk.indexOf(':');
        if (idx === -1) continue;
        const prop = chunk.slice(0, idx).trim().toLowerCase();
        if (!prop) continue;
        let value = chunk.slice(idx + 1).trim();
        // A trailing `!important` is dropped, found from the end (a pattern searching from the start
        // retried every run of whitespace in a long value).
        if (value.toLowerCase().endsWith('important')) {
            let bang = value.length - 'important'.length;
            while (bang > 0 && /\s/.test(value[bang - 1])) bang--;
            if (value[bang - 1] === '!') value = value.slice(0, bang - 1).trim();
        }
        if (value) decls.set(prop, value);
    }
    return decls;
};

/**
 * Reads a declaration, also accepting the `-webkit-`/`-moz-`/`-ms-`/`-o-` prefixed spelling so a
 * vendor-prefixed property keeps matching (substring matching used to catch those by accident).
 */
const getDeclaration = (decls: Map<string, string>, prop: string): string | undefined =>
    decls.get(prop)
    ?? decls.get(`-webkit-${prop}`)
    ?? decls.get(`-moz-${prop}`)
    ?? decls.get(`-ms-${prop}`)
    ?? decls.get(`-o-${prop}`);

/**
 * Returns the first family from a `font-family` stack, respecting quotes so a quoted family name
 * containing a comma (`'Fira, A', serif`) isn't split through the middle of its own name.
 */
const firstFontFamily = (fontFamily: string): string => {
    let quote: '"' | '\'' | null = null;
    let first = '';
    for (const ch of fontFamily) {
        if (quote) {
            if (ch === quote) { quote = null; continue; }
        } else if (ch === '"' || ch === '\'') {
            quote = ch;
            continue;
        } else if (ch === ',') {
            break;
        }
        first += ch;
    }
    return first.trim();
};

/** Whether `ch` is an ASCII letter: after `<` (or `</`) it starts a tag; anything else is text or a declaration. */
const isAsciiLetter = (ch: string | undefined): boolean => ch !== undefined && ((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z'));

/** `attributes` without `id` and `name`, for an element opened again in another place (its ids are the first one's). */
const withoutIds = (attributes: Record<string, string>): Record<string, string> => {
    if (!('id' in attributes) && !('name' in attributes)) return attributes;
    const rest: Record<string, string> = Object.create(null);
    for (const key in attributes) if (key !== 'id' && key !== 'name') rest[key] = attributes[key];
    return rest;
};

/**
 * Builds the element tree of an HTML (or XHTML) document as a browser does, in one pass:
 *  - markup declarations and processing instructions (`<!DOCTYPE html>`, `<?xml ...?>`, Word's
 *    conditional markers `<![if !supportLists]>` and `<![endif]>`) are read as nothing, and a `<` that
 *    starts no tag (`a < b`) is text. A tag's name is everything up to a space, `/` or `>`, so Word's
 *    `<o:p>` and an EPUB's `<m:math>` are elements (their names were read as text);
 *  - a comment ends where HTML ends one (`<!-->`, `<!--->`, `--!>`); CDATA is text, as XHTML reads it;
 *  - `<script>`, `<style>`, `<xmp>`, `<iframe>`, `<noscript>`... hold text up to their end tag, and
 *    `<title>` and `<textarea>` text with its character references, never markup;
 *  - end tags a browser leaves out are implied: a block ends a paragraph, an item the item before it,
 *    a heading an open heading, a cell the cell before it (never one of a table outside it), and a cell
 *    written directly in a table is given its row;
 *  - an end tag does not reach past the scope it stands in (`</b>` in a table cell leaves a `<b>` outside
 *    the table open; `</span>` does not end a `<div>` opened inside it);
 *  - a formatting element (b, i, a, font...) ended with the paragraph it stood in is opened again where
 *    text goes on, until its own end tag (`<p><b>bold<p>still bold</b>`);
 *  - `<html>` and `<body>` build no element: their attributes are kept, and what a browser moves into
 *    the body (content before `<body>` or after `</body>`) stays in the document. `<head>` ends at the
 *    first element or text that is not head content.
 */
const parseHtmlTree = (html: string, config: FullOfficeParserConfig, preserveComments: boolean = false): HtmlTree => {
    const root: HtmlNode = { type: 'element', tagName: 'root', children: [], attributes: {}, depth: 0, open: true };
    const htmlAttributes: Record<string, string> = Object.create(null);
    const bodyAttributes: Record<string, string> = Object.create(null);
    let head: HtmlNode | undefined;
    // Once content is read (or <body> opens), a later <head> is no head.
    let bodyStarted = false;
    let current = root;
    let cursor = 0;
    // Where the next end tag of each raw-text element starts at or after the cursor (-1: nowhere).
    const closeTagAt = new Map<string, number>();
    // Text joins the text node before it, when nothing but markup read as nothing (a comment, an end
    // tag closing nothing, a declaration) stands between them, as a browser shows it: a node each, `a`
    // and a comment repeated 7 million times (62 KB of EPUB) made 14 million nodes, which the element
    // budget, counting markup, did not bound.
    // The open elements of each name (the ancestors of `current`, and it), outermost first. An end tag
    // or an implied end finds the element it closes from these, where each walked up the open elements
    // (up to 256): 16 million end tags closing nothing (63 KB of EPUB) took 17 seconds, and 2 million
    // paragraphs under 250 divs, 55. Every move of `current` goes through `closeTo` or `open`, so each
    // element is pushed and popped once. The open special elements, and those that stop a list item's
    // implied end, are kept the same way.
    const openByName = new Map<string, HtmlNode[]>();
    const openSpecial: HtmlNode[] = [];
    const openItemStops: HtmlNode[] = [];
    // The formatting elements to carry (see FORMATTING_ELEMENTS): each with the attributes it is opened
    // again with; null marks where a cell, caption or object began, which nothing before it crosses.
    const formatting: ({ node: HtmlNode; attributes: Record<string, string> } | null)[] = [];
    // How many elements carrying formatting may still be opened again: a quarter of the document's
    // length, about what its own markup could make, so hostile markup cannot multiply them.
    let reopenLeft = (html.length >> 2) + 64;

    const stackTop = (name: string): HtmlNode | undefined => {
        const stack = openByName.get(name);
        return stack?.[stack.length - 1];
    };
    const closeTo = (ancestor: HtmlNode): void => {
        for (let n: HtmlNode | undefined = current; n && n !== ancestor; n = n.parent) {
            openByName.get(n.tagName!)?.pop();
            if (SPECIAL_ELEMENTS.has(n.tagName!)) {
                openSpecial.pop();
                if (!ITEM_TRANSPARENT.has(n.tagName!)) openItemStops.pop();
            }
            n.open = false;
            // Formatting carried within a cell (or caption, object) ends with it.
            if (n.marksFormatting) while (formatting.length && formatting.pop() !== null);
        }
        current = ancestor;
    };
    const open = (node: HtmlNode): void => {
        let stack = openByName.get(node.tagName!);
        if (!stack) openByName.set(node.tagName!, stack = []);
        stack.push(node);
        if (SPECIAL_ELEMENTS.has(node.tagName!)) {
            openSpecial.push(node);
            if (!ITEM_TRANSPARENT.has(node.tagName!)) openItemStops.push(node);
        }
        node.open = true;
        current = node;
    };
    const insert = (tagName: string, attributes: Record<string, string>): HtmlNode => {
        const node: HtmlNode = { type: 'element', tagName, attributes, children: [], parent: current, depth: (current.depth ?? 0) + 1 };
        // Nesting deeper than the parser reads is refused here, with the typed error, before any
        // walk of the tree could run out of stack on it.
        if (node.depth! > MAX_HTML_NESTING_DEPTH) throw getOfficeError(OfficeErrorType.MAX_NESTING_DEPTH_EXCEEDED, config);
        current.children.push(node);
        return node;
    };
    /** The innermost open element named in `names`, if any. */
    const innermostOf = (names: Iterable<string>): HtmlNode | undefined => {
        let found: HtmlNode | undefined;
        for (const name of names) {
            const top = stackTop(name);
            if (top && (!found || top.depth! > found.depth!)) found = top;
        }
        return found;
    };
    /** Whether no element of `scope` is open inside `target` (an end tag or implied end reaches it). */
    const inScope = (target: HtmlNode, scope: readonly string[]): boolean => {
        for (const name of scope) {
            const top = stackTop(name);
            if (top && top !== target && top.depth! > target.depth!) return false;
        }
        return true;
    };
    /** The open paragraph in button scope, which a block's start tag ends. */
    const closeParagraph = (): void => {
        const paragraph = stackTop('p');
        if (paragraph && inScope(paragraph, BUTTON_SCOPE)) closeTo(paragraph.parent!);
    };
    /** Where in `formatting` the last `tagName` carried in the current cell (or the document) is, else -1. */
    const carriedIndex = (tagName: string): number => {
        for (let k = formatting.length - 1; k >= 0 && formatting[k] !== null; k--) if (formatting[k]!.node.tagName === tagName) return k;
        return -1;
    };
    const carry = (node: HtmlNode): void => {
        let count = 0;
        let oldest = -1;
        for (let k = formatting.length - 1; k >= 0 && formatting[k] !== null; k--) { count++; oldest = k; }
        if (count >= MAX_CARRIED_FORMATTING) formatting.splice(oldest, 1);
        formatting.push({ node, attributes: withoutIds(node.attributes ?? {}) });
    };
    /** The carried formatting elements that are no longer open, opened again at `current` (HTML's reconstruction). */
    const reopenFormatting = (): void => {
        let k = formatting.length - 1;
        if (k < 0 || formatting[k] === null || formatting[k]!.node.open) return;
        while (k > 0 && formatting[k - 1] !== null && !formatting[k - 1]!.node.open) k--;
        for (; k < formatting.length && reopenLeft > 0; k++, reopenLeft--) {
            const entry = formatting[k]!;
            const again = insert(entry.node.tagName!, entry.attributes);
            open(again);
            formatting[k] = { node: again, attributes: entry.attributes };
        }
    };
    const inForeignContent = (): boolean => !!stackTop('svg') || !!stackTop('math');

    // A quote-aware scan that finds no '>' outside quotes runs to the end of the document. The quote
    // state each failed scan was in (none, `"` or `'`), carried forward a character at a time, tells a
    // later scan starting in the same state, at the same place, that it would fail too: two states never
    // become one at a character, so at most three scans run to the end. Scanned again from each tag, a
    // document of `<br ">"` (each '>' between quotes) took two minutes for 2 KB of EPUB.
    let failingFrom = 0;
    const failingStates: string[] = [];
    const scanFailsFrom = (at: number): boolean => {
        if (failingStates.length === 0) return false;
        for (; failingFrom < at; failingFrom++) {
            const ch = html[failingFrom];
            for (let k = 0; k < failingStates.length; k++) {
                const state = failingStates[k];
                if (state) { if (ch === state) failingStates[k] = ''; }
                else if (ch === '"' || ch === '\'') failingStates[k] = ch;
            }
        }
        return failingStates.includes('');
    };
    const pushText = (text: string): void => {
        // Text ends the head (a browser moves it into the body), and goes on in the formatting carried to it.
        if (/[^\t\n\f\r ]/.test(text) && !stackTop('template')) {
            if (head?.open) closeTo(head.parent!);
            bodyStarted = true;
        }
        reopenFormatting();
        const last = current.children[current.children.length - 1];
        if (last?.type === 'text') last.text += text;
        else current.children.push({ type: 'text', text, children: [], parent: current });
    };
    /** The end of a comment starting at `start` (`<!--`) and where reading goes on, as HTML ends one. */
    const commentEnd = (start: number): { end: number; next: number } => {
        const body = start + 4;
        // `<!-->` and `<!--->` are empty comments.
        if (html[body] === '>') return { end: body, next: body + 1 };
        if (html.startsWith('->', body)) return { end: body, next: body + 2 };
        let at = html.indexOf('--', body);
        while (at !== -1 && html[at + 2] !== '>' && !(html[at + 2] === '!' && html[at + 3] === '>')) at = html.indexOf('--', at + 1);
        if (at === -1) return { end: -1, next: html.length };
        return { end: at, next: at + (html[at + 2] === '>' ? 3 : 4) };
    };
    /**
     * The text of a raw-text element opened at the cursor, up to its end tag (`</script`, whatever its
     * case, followed by a space, `/` or `>`). The end tag found is kept while it lies ahead, and one never
     * found is not looked for again: each unclosed tag searched the rest of the document, which took
     * seconds for a few hundred kilobytes of them. Left open when there is none (what follows is read as
     * markup, not swallowed).
     */
    const readRawText = (node: HtmlNode, escapable: boolean): void => {
        const tagName = node.tagName!;
        const known = closeTagAt.get(tagName);
        let closeAt = known;
        if (known === undefined || (known !== -1 && known < cursor)) {
            const closeRe = new RegExp(`</${tagName}(?=[\\t\\n\\f\\r />])`, 'gi');
            closeRe.lastIndex = cursor;
            closeAt = closeRe.exec(html)?.index ?? -1;
            closeTagAt.set(tagName, closeAt);
        }
        if (closeAt === undefined || closeAt === -1) return;
        const text = html.substring(cursor, closeAt);
        // Raw text keeps its `&` as written (`<xmp>&amp;</xmp>` shows `&amp;`): text nodes are decoded when read.
        if (text) node.children.push({ type: 'text', text: escapable ? text : text.replace(/&/g, '&amp;'), children: [], parent: node });
        const end = html.indexOf('>', closeAt);
        cursor = end === -1 ? html.length : end + 1;
        closeTo(node.parent!);
    };

    const startTag = (rawName: string, attrString: string, selfClosing: boolean): void => {
        const foreign = inForeignContent();
        const tagName = rawName === 'image' && !foreign ? 'img' : rawName;
        // One <html> and one <body>, whatever the markup: their attributes are kept, the first of each name.
        if (tagName === 'html' || tagName === 'body') {
            const into = tagName === 'html' ? htmlAttributes : bodyAttributes;
            const attributes = parseAttributes(attrString);
            for (const key in attributes) if (!(key in into)) into[key] = attributes[key];
            if (tagName === 'body') {
                if (head?.open) closeTo(head.parent!);
                bodyStarted = true;
            }
            return;
        }
        if (tagName === 'head') {
            if (!head && !bodyStarted && current === root) open(head = insert('head', parseAttributes(attrString)));
            return;
        }
        // (A template's content is its own: it ends nothing around it.)
        if (!HEAD_CONTENT.has(tagName) && !stackTop('template')) {
            if (head?.open) closeTo(head.parent!);
            bodyStarted = true;
        }
        const table = stackTop('table');
        if (TABLE_PARTS.has(tagName) && !foreign) {
            // Outside a table a table's part is read as nothing, as a browser reads it: what it holds stays.
            if (!table) return;
            const section = innermostOf(['tbody', 'thead', 'tfoot']);
            const tableSection = section && section.depth! > table.depth! ? section : table;
            if (tagName === 'td' || tagName === 'th') {
                // A cell goes in its table's open row; one written directly in a table (or a section) is given
                // a row, as a browser gives it. The row is never one of a table outside this one.
                const row = stackTop('tr');
                if (row && row.depth! > table.depth!) closeTo(row);
                else {
                    closeTo(tableSection);
                    open(insert('tr', Object.create(null)));
                }
            } else if (tagName === 'tr') {
                closeTo(tableSection);
            } else if (tagName === 'col') {
                const group = stackTop('colgroup');
                closeTo(group && group.depth! > table.depth! ? group : table);
            } else {
                closeTo(table);
            }
        } else if (tagName === 'table' && table && !foreign) {
            // A table started in a table, not in one of its cells, ends that table first.
            const cell = innermostOf(['td', 'th', 'caption']);
            if (!cell || cell.depth! < table.depth!) closeTo(table.parent!);
            else closeParagraph();
        } else if (CLOSES_PARAGRAPH.has(tagName) && !foreign) {
            if (tagName === 'li' || tagName === 'dd' || tagName === 'dt') {
                // An item ends the item before it, unless a block other than a division or paragraph stands between.
                const stop = openItemStops[openItemStops.length - 1];
                if (stop && (tagName === 'li' ? stop.tagName === 'li' : stop.tagName === 'dd' || stop.tagName === 'dt')) closeTo(stop.parent!);
            }
            closeParagraph();
            // A heading ends a heading it would otherwise be written in (`<h1>A<h2>B` is two headings).
            if (/^h[1-6]$/.test(tagName) && /^h[1-6]$/.test(current.tagName ?? '')) closeTo(current.parent!);
        } else if (tagName === 'option' || tagName === 'optgroup') {
            const option = stackTop('option');
            if (option && current === option) closeTo(option.parent!);
            const group = stackTop('optgroup');
            if (tagName === 'optgroup' && group && current === group) closeTo(group.parent!);
        } else if (tagName === 'button') {
            const button = stackTop('button');
            if (button && inScope(button, DEFAULT_SCOPE)) closeTo(button.parent!);
        } else if (tagName === 'a') {
            // A link cannot hold a link: an open one ends where another starts (so no text is given its
            // links' metadata again at every level of nesting).
            const k = carriedIndex('a');
            if (k !== -1) {
                const link = formatting[k]!.node;
                formatting.splice(k, 1);
                if (link.open) closeTo(link.parent!);
            }
        }
        if (!OPENS_NO_FORMATTING.has(tagName)) reopenFormatting();
        const node = insert(tagName, parseAttributes(attrString));
        if (selfClosing || VOID_ELEMENTS.has(tagName)) return;
        open(node);
        if (FORMATTING_BOUNDARIES.has(tagName)) {
            node.marksFormatting = true;
            formatting.push(null);
        }
        if (FORMATTING_ELEMENTS.has(tagName)) carry(node);
        // In SVG or MathML only a script or a style is raw text (an SVG <title> is an element).
        if (RAW_TEXT_ELEMENTS.has(tagName) && (!foreign || tagName === 'script' || tagName === 'style')) readRawText(node, false);
        else if (ESCAPABLE_RAW_TEXT_ELEMENTS.has(tagName) && !foreign) readRawText(node, true);
    };

    const endTag = (tagName: string): void => {
        // The one <html> and <body> are never ended: what follows their end tags is still the document's.
        if (tagName === 'html' || tagName === 'body') return;
        if (tagName === 'br') return startTag('br', '', true);
        if (tagName === 'head') {
            if (head?.open) closeTo(head.parent!);
            return;
        }
        if (/^h[1-6]$/.test(tagName)) {
            // Any heading's end tag ends the open heading, whatever its level.
            const heading = innermostOf(['h1', 'h2', 'h3', 'h4', 'h5', 'h6']);
            if (heading && inScope(heading, DEFAULT_SCOPE)) closeTo(heading.parent!);
            return;
        }
        if (FORMATTING_ELEMENTS.has(tagName)) {
            const k = carriedIndex(tagName);
            if (k !== -1) {
                const element = formatting[k]!.node;
                if (!element.open) {
                    formatting.splice(k, 1);
                } else if (inScope(element, DEFAULT_SCOPE)) {
                    formatting.splice(k, 1);
                    closeTo(element.parent!);
                }
                return;
            }
        }
        const target = stackTop(tagName);
        if (!target) return;
        if (TABLE_SCOPED_ENDS.has(tagName)) {
            if (!inScope(target, TABLE_SCOPE)) return;
        } else if (SPECIAL_ELEMENTS.has(tagName)) {
            if (!inScope(target, tagName === 'p' ? BUTTON_SCOPE : tagName === 'li' ? LIST_ITEM_SCOPE : DEFAULT_SCOPE)) return;
        } else {
            // Any other element's end tag ends it only when no special element (a block) is open inside it.
            const special = openSpecial[openSpecial.length - 1];
            if (special && special.depth! > target.depth!) return;
        }
        closeTo(target.parent!);
    };

    while (cursor < html.length) {
        const tagStart = html.indexOf('<', cursor);

        if (tagStart === -1) {
            const text = html.substring(cursor);
            if (text) pushText(text);
            break;
        }

        if (tagStart > cursor) {
            const text = html.substring(cursor, tagStart);
            if (text) pushText(text);
        }
        cursor = tagStart;

        if (html.startsWith('<!--', tagStart)) {
            const { end, next: resume } = commentEnd(tagStart);
            // Kept as a node only when asked (HtmlParserConfig.preserveComments). A conditional comment
            // (`<!--[if ...]>`, `<!--<![endif]-->`) is an Office/IE directive, not an authored note: never kept.
            if (preserveComments && end !== -1) {
                const body = html.substring(tagStart + 4, end);
                if (!/^\[if\b/i.test(body) && !/<!\[endif\]$/i.test(body)) {
                    current.children.push({ type: 'comment', text: body, children: [], parent: current });
                }
            }
            cursor = resume;
            continue;
        }
        const next = html[tagStart + 1];
        if (next === '!' || next === '?') {
            if (html.startsWith('<![CDATA[', tagStart)) {
                // A CDATA section is its text, as XHTML reads it (an EPUB's chapters are XHTML).
                const end = html.indexOf(']]>', tagStart + 9);
                const text = html.substring(tagStart + 9, end === -1 ? html.length : end);
                if (text) pushText(text.replace(/&/g, '&amp;'));
                cursor = end === -1 ? html.length : end + 3;
                continue;
            }
            // A declaration or processing instruction (`<!DOCTYPE html>`, `<?xml ...?>`, Word's conditional
            // markers `<![if !supportLists]>` and `<![endif]>`): markup read as nothing, as a browser reads
            // it, and what stands between the markers is kept.
            const end = html.indexOf('>', tagStart + 2);
            cursor = end === -1 ? html.length : end + 1;
            continue;
        }
        if (next === '/') {
            const after = html[tagStart + 2];
            if (!isAsciiLetter(after)) {
                // `</>` is nothing; `</` and anything but a letter is a declaration to its `>`, as a browser reads it.
                const end = after === '>' ? tagStart + 2 : html.indexOf('>', tagStart + 2);
                cursor = end === -1 ? html.length : end + 1;
                continue;
            }
        } else if (!isAsciiLetter(next)) {
            // A `<` that starts no tag (`a < b`, `<5`) is text.
            pushText('<');
            cursor = tagStart + 1;
            continue;
        }

        // Scan for the tag's closing '>', skipping any that appear inside a quoted attribute value.
        // Browsers do NOT escape '>' inside attribute values on serialization, so a literal '>' there
        // (e.g. a mermaid diagram's `-->` in data-mermaid) must not be read as the tag end. The scan
        // is linear in the tag's length and the cursor never rewinds, so the parse stays O(n) overall
        // (no substring().match allocation per '<').
        let tagEndIdx = -1;
        if (!scanFailsFrom(tagStart + 1)) {
            let attrQuote = '';
            for (let i = tagStart + 1; i < html.length; i++) {
                const ch = html[i];
                if (attrQuote) {
                    if (ch === attrQuote) attrQuote = '';
                } else if (ch === '"' || ch === '\'') {
                    attrQuote = ch;
                } else if (ch === '>') {
                    tagEndIdx = i;
                    break;
                }
            }
            if (tagEndIdx === -1) {
                if (failingStates.length === 0) failingFrom = tagStart + 1;
                failingStates.push('');
            }
        }
        if (tagEndIdx === -1) {
            // The quote-aware scan ran to the end without closing the tag: an unbalanced quote in a tag.
            // Retry naively for the next literal '>', so a malformed tag degrades to a tag instead of
            // swallowing the rest of the document into one text node.
            tagEndIdx = html.indexOf('>', tagStart);
        }
        if (tagEndIdx === -1) {
            pushText(html.substring(tagStart));
            break;
        }

        const tagContent = html.substring(tagStart + 1, tagEndIdx);
        cursor = tagEndIdx + 1;

        const isClosing = tagContent.startsWith('/');
        const isSelfClosing = !isClosing && tagContent.endsWith('/');
        const tagCore = tagContent.replace(/^\/|\/$/g, '').trim();
        // The name runs to a space or `/`: any other character is part of it (`o:p`, `m:math`, `my-element`).
        const nameEnd = tagCore.search(/[\t\n\f\r /]/);
        const tagName = (nameEnd === -1 ? tagCore : tagCore.substring(0, nameEnd)).toLowerCase();
        const attrString = nameEnd === -1 ? '' : tagCore.substring(nameEnd);

        if (isClosing) endTag(tagName);
        else startTag(tagName, attrString, isSelfClosing);
    }

    return { root, head, htmlAttributes, bodyAttributes };
};

/** What a container whose parts are HTML (an EPUB's chapters) tells the reading of each part. */
export interface HtmlPartContext {
    /**
     * The name of the attachment a picture's `src` shows, when the container holds that picture as
     * a file of its own; undefined for any other source. The container makes the attachment, once
     * however many pictures show it: an EPUB's images were written into each chapter as data, a copy
     * for every `<img>`, so a small book showing one large picture many times took gigabytes.
     */
    imageAttachment?: (src: string) => string | undefined;
}



export const parseHtml = async (buffer: Buffer, config: FullOfficeParserConfig, part: HtmlPartContext = {}): Promise<OfficeParserAST> => {
    // Honour cancellation requests before the HTML tree is built and traversed.
    // The custom recursive HTML parser can be expensive for large documents;
    // rejecting early here prevents both the parsing and the subsequent AST construction.
    checkAbortSignal(config.abortSignal);

    const { root, head } = parseHtmlTree(buffer.toString('utf-8'), config, config.htmlParserConfig?.preserveComments === true);
    // What a browser shows: the document, its <head> passed over (the tree builds no <html> or <body>).
    const body = root;

    const metadata: OfficeMetadata = {};
    const attachments: OfficeAttachment[] = [];

    // The head's <meta> elements.
    const metaElements = (head?.children ?? []).filter(child => child.tagName === 'meta');
    if (head) {
        const titleNode = head.children.find(child => child.tagName === 'title');
        if (titleNode && titleNode.children.length > 0 && titleNode.children[0].text) {
            metadata.title = titleNode.children[0].text;
        }

        metadata.nativeProperties = {};
        for (const child of metaElements) {
            const name = child.attributes?.name || child.attributes?.property || child.attributes?.['http-equiv'];
            if (name) {
                setOwn(metadata.nativeProperties, name, child.attributes?.content || '');
            }
        }

        const extractMeta = (name: string): string | undefined => {
            for (const child of metaElements) {
                if (child.attributes?.name === name || child.attributes?.property === name) {
                    return child.attributes?.content;
                }
            }
            return undefined;
        };

        const author = extractMeta('author');
        if (author) metadata.author = author;
        const desc = extractMeta('description');
        if (desc) metadata.description = desc;

        const created = extractMeta('dcterms.created');
        if (created) metadata.created = new Date(created);
        const modified = extractMeta('dcterms.modified');
        if (modified) metadata.modified = new Date(modified);
        const lastMod = extractMeta('lastModifiedBy');
        if (lastMod) metadata.lastModifiedBy = lastMod;

        // Custom properties
        const customProps: Record<string, string | number | boolean | Date> = {};
        for (const child of metaElements) {
            if (child.attributes?.name?.startsWith('custom:')) {
                const key = child.attributes.name.substring(7);
                const val = child.attributes.content || '';
                // Try to infer type
                if (val === 'true') setOwn(customProps, key, true);
                else if (val === 'false') setOwn(customProps, key, false);
                else if (!isNaN(Number(val)) && val.trim() !== '') setOwn(customProps, key, Number(val));
                else if (!isNaN(Date.parse(val)) && val.includes(':')) setOwn(customProps, key, new Date(val));
                else setOwn(customProps, key, val);
            }
        }
        if (Object.keys(customProps).length > 0) metadata.customProperties = customProps;
    }

    const content: OfficeContentNode[] = [];
    let htmlListIdCounter = 1;

    interface ListContext {
        listId: string;
        type: 'ordered' | 'unordered';
        level: number;
        counters: Record<number, number>;
        isTask?: boolean;
    }

    // Finds the checked state from a nested <input type="checkbox"> (GFM task-list items
    // nest it inside a <label>, so it isn't a direct child of the <li>).
    const findNestedCheckboxChecked = (n: HtmlNode): boolean | undefined => {
        if (n.tagName === 'input' && (n.attributes?.type || '').toLowerCase() === 'checkbox') {
            return 'checked' in (n.attributes || {});
        }
        for (const child of n.children) {
            const found = findNestedCheckboxChecked(child);
            if (found !== undefined) return found;
        }
        return undefined;
    };

    // Populated from the document's footnote definitions (found and parsed before the main body loop,
    // since references can appear anywhere earlier in the document; see readFootnotes below) and
    // consulted by parseChildren's footnote-reference handling below.
    const footnoteDefinitions = new Map<string, OfficeContentNode[]>();
    // Each definition's note id (GitHub's `user-content-fn-1` is note `1`) and kind, by key.
    const footnoteIds = new Map<string, string>();
    const footnoteTypes = new Map<string, 'footnote' | 'endnote'>();
    // The elements definitions were read from, and the footnotes sections that held them: left out of the text.
    const footnoteElements = new Set<HtmlNode>();
    // Keys a reference actually consumed, so definitions in the section that no reference points at
    // (orphans) can be recovered at the end instead of silently dropped.
    const referencedFootnoteKeys = new Set<string>();
    // One note node per key, shared by every reference to it: a node (and its text) built per reference
    // let a small file refer to one large note thousands of times and fill the heap.
    const noteNodesByKey = new Map<string, OfficeContentNode>();
    // Set while a definition is read: a back-link in it (to its reference) is plumbing, not the note's text.
    let readingFootnote = 0;
    const hasToken = (value: string | undefined, token: string): boolean => !!value && value.split(/\s+/).includes(token);
    /** A note's back-link to its reference: this library's, GitHub's, Pandoc's and markdown-it's. */
    const isFootnoteBackLink = (node: HtmlNode): boolean => node.tagName === 'a' && (
        (node.attributes?.href || '').startsWith('#footnote-ref-') || node.attributes?.['data-footnote-backref'] !== undefined
        || node.attributes?.role === 'doc-backlink' || hasToken(node.attributes?.class, 'footnote-back') || hasToken(node.attributes?.class, 'footnote-backref'));
    /** Whether `link` is marked as a reference to a note (`inMarkedSup`: it is all a `<sup class="footnote-ref">` holds). */
    const isNoteReferenceLink = (link: HtmlNode, inMarkedSup: boolean): boolean => link.tagName === 'a' && (inMarkedSup
        || link.attributes?.['data-footnote-ref'] !== undefined || link.attributes?.role === 'doc-noteref'
        || hasToken(link.attributes?.['epub:type'], 'noteref') || hasToken(link.attributes?.class, 'footnote-ref'));
    /** The one element `node` holds, with nothing but whitespace and comments beside it. */
    const soleElementChild = (node: HtmlNode): HtmlNode | undefined => {
        let sole: HtmlNode | undefined;
        for (const child of node.children) {
            if (child.type === 'comment' || (child.type === 'text' && !/[^\t\n\f\r ]/.test(child.text || ''))) continue;
            if (child.type !== 'element' || sole) return undefined;
            sole = child;
        }
        return sole;
    };
    /** The link a footnote reference is, if `node` is one of the link forms (bare, or all a `<sup>` holds). */
    const referenceLinkOf = (node: HtmlNode): HtmlNode | undefined => {
        const link = node.tagName === 'a' ? node : node.tagName === 'sup' ? soleElementChild(node) : undefined;
        return link && isNoteReferenceLink(link, link !== node && hasToken(node.attributes?.class, 'footnote-ref')) ? link : undefined;
    };
    /**
     * The key of the note `node` refers to, when it is a footnote reference: this library's `<sup
     * data-footnote-ref="KEY">`, or a link marked as one (see isNoteReferenceLink) to a definition read.
     * A link to no definition stays a link.
     */
    const footnoteReferenceKey = (node: HtmlNode): string | undefined => {
        if (node.type !== 'element') return undefined;
        if (node.tagName === 'sup' && node.attributes?.['data-footnote-ref'] !== undefined) return node.attributes['data-footnote-ref'];
        const href = referenceLinkOf(node)?.attributes?.href;
        return href?.startsWith('#') && footnoteDefinitions.has(href.slice(1)) ? href.slice(1) : undefined;
    };
    /** The note a reference to `key` shares (see noteNodesByKey). */
    const noteFor = (key: string): OfficeContentNode => {
        let noteNode = noteNodesByKey.get(key);
        if (!noteNode) {
            const definition = footnoteDefinitions.get(key);
            noteNode = {
                type: 'note',
                text: nodesText(definition || []),
                children: definition || [],
                metadata: { noteType: footnoteTypes.get(key) ?? 'footnote', noteId: footnoteIds.get(key) ?? key }
            };
            noteNodesByKey.set(key, noteNode);
        }
        return noteNode;
    };

    // --- Generic attribute pass-through (htmlParserConfig.preserveAttributes) ---------------
    // Captures attributes no typed metadata field consumed, so they can be replayed on
    // generation. Everything here is a *defence-in-depth* filter: HtmlGenerator sanitizes again
    // on the way out, because an AST can be built programmatically rather than parsed.
    const preserveAttributes = config.htmlParserConfig?.preserveAttributes === true;

    // `style` is already consumed wholesale into TextFormatting/metadata above, and `id` is
    // consumed into anchorIds and re-emitted by the generator - carrying either would duplicate
    // an attribute the generator composes itself. `class` is deliberately NOT excluded: the
    // generator's class attribute is built purely from style-mapping and never from a parsed
    // `class`, so without this a plain `<p class="lead">` loses "lead" entirely. The generator
    // merges it into that attribute rather than emitting a second one.
    const GENERATOR_OWNED_ATTRS = new Set(['id', 'style']);

    /**
     * Captures the attributes of `node` that `consumed` didn't claim.
     * Returns undefined when nothing survives, so the field stays absent rather than `{}`.
     */
    const collectHtmlAttributes = (node: HtmlNode, consumed: string[]): Record<string, string> | undefined => {
        if (!preserveAttributes || !node.attributes) return undefined;
        const consumedSet = new Set(consumed.map(c => c.toLowerCase()));
        const bag: Record<string, string> = {};
        for (const [rawKey, value] of Object.entries(node.attributes)) {
            const key = rawKey.toLowerCase();
            if (consumedSet.has(key) || GENERATOR_OWNED_ATTRS.has(key)) continue;
            // Event handlers are never carried, at any layer, with no opt-in.
            if (/^on/i.test(key)) continue;
            // srcdoc holds a whole HTML document; it cannot be safely escaped into an attribute.
            if (key === 'srcdoc') continue;
            // Reject anything that isn't a plain attribute name outright - a key containing a
            // quote or '=' is the shape an attribute-injection payload takes.
            if (!isSafeHtmlAttributeName(key)) continue;
            setOwn(bag, key, value);
        }
        return Object.keys(bag).length > 0 ? bag : undefined;
    };

    const parseNode = (node: HtmlNode, currentFormatting: TextFormatting = {}, listContext?: ListContext, depth: number = 0): OfficeContentNode | OfficeContentNode[] | null => {
        // Guard against a maliciously deep element tree (e.g. tens of thousands of nested
        // <div>) recursing until the call stack overflows.
        //
        // The previous limit of 1000 could never fire: measured overflow is around 800 and
        // varies run to run (796/862/796 on three identical runs), so the RangeError always
        // arrived first and the typed error this guard exists to produce never did. Failure was
        // still graceful - it surfaces as a wrapped Error, not a crash - which is why this was a
        // dead guard rather than a denial of service.
        //
        // 256 is chosen to hold across engines rather than tuned to one. It is far below the
        // lowest overflow observed here and leaves room for a smaller frame budget on older V8
        // (the supported floor is Node 18), while sitting orders of magnitude above real
        // content: the bundled HTML and EPUB fixtures reach an AST depth of 8.
        // Per node, alongside the depth guard: the two together are what make a hostile
        // document both bounded and cancellable rather than only bounded.
        checkAbortSignal(config.abortSignal);
        if (depth > MAX_HTML_NESTING_DEPTH) {
            throw getOfficeError(OfficeErrorType.MAX_NESTING_DEPTH_EXCEEDED, config);
        }
        if (node.type === 'comment') {
            // A preserved HTML comment: the author's hidden note, kept verbatim (never entity-decoded -
            // entities are not interpreted inside a comment).
            return { type: 'comment', text: node.text || '', metadata: { sourceSyntax: 'html' } as CommentMetadata };
        }

        if (node.type === 'text') {
            let decodedText = decodeEntities(node.text || '');

            if (!config.preserveXmlWhitespace) {
                decodedText = collapseWhitespace(decodedText);
            }
            if (!decodedText.trim() && !config.preserveXmlWhitespace) return null;

            const textNode: OfficeContentNode = {
                type: 'text',
                text: decodedText,
                formatting: Object.keys(currentFormatting).length > 0 ? { ...currentFormatting } : undefined
            };

            if (config.includeRawContent && node.text) {
                // For text nodes in this manual parser, we just use the decoded text as raw content
                // as we don't have accurate locators for the original source slice
                textNode.rawContent = node.text;
            }

            return textNode;
        }

        if (node.type === 'element' && node.tagName) {
            const tagName = node.tagName;
            // What is never shown: the head (read for the metadata), a <title> wherever it stands, scripts,
            // styles, a <template>'s inert content and <noscript> (read as a browser running scripts reads it).
            if (HIDDEN_ELEMENTS.has(tagName)) return null;
            const newFormatting = { ...currentFormatting };

            if (tagName === 'b' || tagName === 'strong') newFormatting.bold = true;
            if (tagName === 'i' || tagName === 'em') newFormatting.italic = true;
            if (tagName === 'u') newFormatting.underline = true;
            if (tagName === 'strike' || tagName === 's' || tagName === 'del') newFormatting.strikethrough = true;
            if (tagName === 'sub') newFormatting.subscript = true;
            if (tagName === 'sup') newFormatting.superscript = true;
            if (tagName === 'code') newFormatting.font = 'monospace';
            if (tagName === 'mark') {
                // <mark> is a highlight. Use its data-color when present (an inline
                // background-color style, read below, still wins); a bare <mark> falls back to the
                // conventional yellow so it round-trips as a highlight rather than plain text.
                newFormatting.backgroundColor = node.attributes?.['data-color'] || '#ffff00';
            }

            const styleAttr = node.attributes?.style || '';
            const alignAttr = node.attributes?.align || '';
            if (styleAttr || alignAttr) {
                const decls = parseStyleDeclarations(styleAttr);

                // `bold`, `bolder`, and any weight >= 600 are all bold; the old substring check
                // only ever saw the literal "font-weight: bold".
                const weight = getDeclaration(decls, 'font-weight');
                if (weight) {
                    const numericWeight = parseInt(weight, 10);
                    if (weight === 'bold' || weight === 'bolder' || (!isNaN(numericWeight) && numericWeight >= 600)) {
                        newFormatting.bold = true;
                    }
                }

                if (getDeclaration(decls, 'font-style') === 'italic') newFormatting.italic = true;

                // text-decoration is a shorthand that can carry several keywords at once, so
                // "underline line-through" has to set both flags rather than only the first.
                const decoration = getDeclaration(decls, 'text-decoration') ?? getDeclaration(decls, 'text-decoration-line');
                if (decoration) {
                    const parts = decoration.split(/\s+/);
                    if (parts.includes('underline')) newFormatting.underline = true;
                    if (parts.includes('line-through')) newFormatting.strikethrough = true;
                }

                const color = decls.get('color');
                if (color) newFormatting.color = color;

                const background = getDeclaration(decls, 'background-color');
                if (background) newFormatting.backgroundColor = background;

                const size = getDeclaration(decls, 'font-size');
                if (size) newFormatting.size = size;

                const fontFamily = getDeclaration(decls, 'font-family');
                if (fontFamily) newFormatting.font = firstFontFamily(fontFamily);

                const textAlign = getDeclaration(decls, 'text-align')?.toLowerCase();
                if (textAlign && ['left', 'center', 'right', 'justify'].includes(textAlign)) {
                    newFormatting.alignment = textAlign as any;
                } else if (alignAttr) {
                    const align = alignAttr.toLowerCase();
                    if (['left', 'center', 'right', 'justify'].includes(align)) {
                        newFormatting.alignment = align as any;
                    }
                }
            }

            const anchorIds = node.attributes?.id ? [node.attributes.id] : [];

            // An element's content: its text, and what each child element gives (a node, or a list it
            // passes up, which is already collapsed as its own content). The spaces between pieces are
            // collapsed where two pieces meet, and a block's edges trimmed in place: collapsed and trimmed
            // again at every level, and copied up into each, a list under 250 wrappers was read 250
            // times (2 million paragraphs, 6.7 KB of EPUB, took 55 seconds). The first list a child
            // passes up is taken over rather than copied; later ones are appended in bulk.
            const parseChildren = (n: HtmlNode, fmt: TextFormatting, lCtx?: any): OfficeContentNode[] => {
                let kids: OfficeContentNode[] = [];
                const collapse = !config.preserveXmlWhitespace;
                // The last node kept that takes room (a named anchor does not), for collapsing spaces.
                let previous: OfficeContentNode | undefined;
                // Adds one node, its leading space dropped when the node before ends in one (as
                // collapseSpacesAcrossNodes does); false when it is dropped.
                const add = (node: OfficeContentNode): boolean => {
                    let current = node;
                    if (collapse && current.type === 'text' && current.text?.startsWith(' ') && previous?.type === 'text' && previous.text?.endsWith(' ')) {
                        current = { ...current, text: current.text.slice(1) };
                        if (!current.text && !current.notes?.length && !current.comments?.length) return false;
                    }
                    kids.push(current);
                    if (!isAnchorMark(current)) previous = current;
                    return current === node;
                };
                const addList = (nodes: OfficeContentNode[]): void => {
                    // The nodes taken as they are: all of them into an empty list, else those after the
                    // first kept unchanged (only the ones before it can meet the space before them).
                    let bulkFrom = 0;
                    if (kids.length === 0) {
                        kids = nodes;
                    } else {
                        while (bulkFrom < nodes.length) {
                            const node = nodes[bulkFrom++];
                            if (add(node) && !isAnchorMark(node)) break;
                        }
                        for (let from = bulkFrom; from < nodes.length; from += 32768) {
                            Array.prototype.push.apply(kids, nodes.slice(from, from + 32768));
                        }
                    }
                    for (let j = nodes.length - 1; j >= bulkFrom; j--) if (!isAnchorMark(nodes[j])) { previous = nodes[j]; break; }
                };
                for (let i = 0; i < n.children.length; i++) {
                    const child = n.children[i];
                    if (child.type === 'text' && !config.preserveXmlWhitespace) {
                        const text = visibleText(n, i);
                        if (text === null) continue;
                        const textNode: OfficeContentNode = { type: 'text', text, formatting: Object.keys(fmt).length > 0 ? { ...fmt } : undefined };
                        if (config.includeRawContent && child.text) textNode.rawContent = child.text;
                        add(textNode);
                        continue;
                    }
                    // Footnote/endnote reference: attach as .notes on the preceding node
                    // instead of inserting a visible node, matching WordParser's convention.
                    const key = footnoteReferenceKey(child);
                    if (key !== undefined) {
                        referencedFootnoteKeys.add(key);
                        // ignoreNotes drops footnotes at parse time (as in DOCX/ODT/PDF): skip the marker
                        // and attach nothing. The orphan sweep below is likewise skipped.
                        if (config.ignoreNotes) continue;
                        const noteNode = noteFor(key);
                        if (kids.length > 0) {
                            const target = kids[kids.length - 1];
                            if (!target.notes) target.notes = [];
                            target.notes.push(noteNode);
                        } else {
                            add({ type: 'text', text: '', notes: [noteNode] });
                        }
                        continue;
                    }

                    const parsed = parseNode(child, fmt, lCtx, depth + 1);
                    if (parsed) {
                        if (Array.isArray(parsed)) addList(parsed);
                        else add(parsed);
                    }
                }
                if (collapse && !isInlineElement(n)) trimBlockEdgesInPlace(kids);
                return kids;
            };

            // Source-comment element (HtmlGenerator's `sourceAttributes` shape, emitted for editors whose DOM
            // parser discards real comment nodes): `<span data-html-comment="raw text">`. Always read, like
            // the gated embed below - it is this library's own round-trip shape. The text is data only.
            // Only an empty element: that is the shape the generator writes, and an element with content
            // is content (its text must not vanish into a hidden note).
            if ((tagName === 'span' || tagName === 'div') && node.attributes?.['data-html-comment'] !== undefined && node.children.length === 0) {
                return {
                    type: 'comment',
                    text: node.attributes['data-html-comment'],
                    metadata: { sourceSyntax: 'html' } as CommentMetadata
                };
            }

            // Gated generic-iframe embed (HtmlGenerator's `gatedEmbeds` shape): an inert
            // click-to-load placeholder that never auto-loads its src. Read it back to the same
            // `embed` node unconditionally - capturing metadata is safe (the src is scheme-checked
            // again on any re-emit); it is the trusted, already-gated counterpart to a raw <iframe>.
            if (tagName === 'div' && node.attributes?.['data-embed-gated'] !== undefined) {
                const gatedSrc = node.attributes?.['data-embed-src'] || '';
                if (!gatedSrc) return null;
                const gatedAlignAttr = node.attributes?.['data-embed-align'];
                const gatedAlign = (['left', 'center', 'right'] as const).includes(gatedAlignAttr as any) ? gatedAlignAttr as 'left' | 'center' | 'right' : undefined;
                const gatedNode: OfficeContentNode = {
                    type: 'embed',
                    text: gatedSrc,
                    metadata: {
                        embedType: 'iframe',
                        url: gatedSrc,
                        width: node.attributes?.['data-embed-width'],
                        height: node.attributes?.['data-embed-height'],
                        align: gatedAlign,
                        label: node.attributes?.['data-embed-label']
                    } as EmbedMetadata
                };
                if (config.includeRawContent) gatedNode.rawContent = '<div data-embed-gated>...</div>';
                return gatedNode;
            }

            // YouTube embeds: attribute-driven editors render
            // <div data-youtube-video="ID" data-width="…" data-align="…">…<iframe…></div>.
            // Recognise both the wrapper div and a bare iframe so externally-authored HTML
            // (and a saved-then-reopened .md that fell back to raw HTML) both round-trip.
            if (tagName === 'div' && node.attributes?.['data-youtube-video'] !== undefined) {
                const videoId = node.attributes['data-youtube-video'] || '';
                const width = node.attributes?.['data-width'];
                const embedAlignAttr = node.attributes?.['data-align'];
                const embedAlign = (['left', 'center', 'right'] as const).includes(embedAlignAttr as any) ? embedAlignAttr as 'left' | 'center' | 'right' : undefined;
                const embedUrl = videoId ? `https://www.youtube.com/watch?v=${videoId}` : undefined;
                const embedNode: OfficeContentNode = {
                    type: 'embed',
                    // Childless nodes need .text so generic AST consumers (text/chunking generators)
                    // don't silently drop them.
                    text: embedUrl,
                    metadata: {
                        embedType: 'youtube',
                        videoId,
                        url: embedUrl,
                        width,
                        align: embedAlign,
                        label: node.attributes?.['data-embed-label']
                    } as EmbedMetadata
                };
                if (config.includeRawContent) embedNode.rawContent = '<div data-youtube-video>...</div>';
                return embedNode;
            }
            if (tagName === 'iframe') {
                const src = node.attributes?.src || '';
                const ytMatch = /youtube(?:-nocookie)?\.com/.test(src) ? src.match(/(?:embed\/|v=)([^&?/\s]+)/) : null;
                if (ytMatch) {
                    const embedUrl = `https://www.youtube.com/watch?v=${ytMatch[1]}`;
                    const embedNode: OfficeContentNode = {
                        type: 'embed',
                        text: embedUrl,
                        // Carry the iframe's own width/height so a YouTube iframe's dimensions are not
                        // dropped (metadata is unified across embedTypes; align has no source here).
                        metadata: {
                            embedType: 'youtube',
                            videoId: ytMatch[1],
                            url: embedUrl,
                            width: node.attributes?.width,
                            height: node.attributes?.height
                        } as EmbedMetadata
                    };
                    if (config.includeRawContent) embedNode.rawContent = '<iframe>...</iframe>';
                    return embedNode;
                }
                // Non-YouTube iframes are dropped by default (a deliberate security posture).
                // preserveIframes opts back in, keeping the src as a generic 'iframe' embed; the
                // src is scheme-checked again on generation, so this only widens what is retained.
                // The src is decoded (attribute values are, as they are read), so the generator's
                // re-escaping does not double-escape it and corrupt a query string.
                const decodedSrc = src;
                if (iframeAllowed(decodedSrc, config.htmlParserConfig?.preserveIframes)) {
                    const iframeNode: OfficeContentNode = {
                        type: 'embed',
                        text: decodedSrc,
                        metadata: {
                            embedType: 'iframe',
                            url: decodedSrc,
                            width: node.attributes?.width,
                            height: node.attributes?.height
                        } as EmbedMetadata
                    };
                    if (config.includeRawContent) iframeNode.rawContent = '<iframe>...</iframe>';
                    return iframeNode;
                }
                return null;
            }

            // A footnotes section definitions were read from, and a note read as one (an EPUB's cited
            // <aside>): their notes were read up front (see readFootnotes below), so they are skipped here
            // wherever they appear in the tree (a non-standalone HtmlGenerator output wraps its section in a
            // <div>). A section that gave no definition is content.
            if (footnoteElements.has(node)) return null;
            // A note's back-link to its reference, read with the note, is plumbing, not its text.
            if (readingFootnote > 0 && isFootnoteBackLink(node)) return null;

            // Math. Two accepted shapes, disambiguated by the `data-math` value:
            //   1. This library's own output - `data-math="inline|block"` names the mode, and the
            //      LaTeX is the visible ($-delimited, escaped) text content.
            //   2. Attribute-driven producers that put the raw LaTeX in `data-math` and signal the
            //      mode through the class (`math-inline`/`math-block`) or the tag.
            // Anything whose `data-math` is exactly `inline`/`block` takes path 1 unchanged; every
            // other value is treated as LaTeX (path 2). LaTeX literally equal to `inline`/`block`
            // is the only ambiguous input, and its text content wins there anyway.
            if ((tagName === 'div' || tagName === 'span') && node.attributes?.['data-math'] !== undefined) {
                const dataMath = node.attributes['data-math'];
                const modeIsExplicit = dataMath === 'inline' || dataMath === 'block';
                const classTokens = (node.attributes?.class || '').split(/\s+/);
                const rawText = decodeEntities(rawChildText(node));
                // Prefer the text content; fall back to the attribute value (path 2 producers may
                // emit an empty body).
                const source = rawText || (modeIsExplicit ? '' : dataMath);
                // Strip whichever `$`/`$$` delimiters are actually present, independent of the
                // resolved mode - a `$`-delimited body inside a <div> must not keep its delimiters.
                // The delimiter also disambiguates the mode when neither an explicit `data-math` nor
                // a `math-inline`/`math-block` class settles it (so `<div data-math="x">$x$</div>`
                // reads as inline, not block-via-tag).
                let latex = source;
                let delimiterMode: 'inline' | 'block' | undefined;
                if (source.length >= 4 && source.startsWith('$$') && source.endsWith('$$')) {
                    latex = source.slice(2, -2);
                    delimiterMode = 'block';
                } else if (source.length >= 2 && source.startsWith('$') && source.endsWith('$')) {
                    latex = source.slice(1, -1);
                    delimiterMode = 'inline';
                }
                const mathMode: 'inline' | 'block' = modeIsExplicit
                    ? (dataMath as 'inline' | 'block')
                    : classTokens.includes('math-block') ? 'block'
                        : classTokens.includes('math-inline') ? 'inline'
                            : delimiterMode ?? (tagName === 'div' ? 'block' : 'inline');
                return {
                    type: 'code',
                    text: latex,
                    metadata: { math: mathMode, anchorIds: anchorIds.length > 0 ? anchorIds : undefined } as CodeMetadata
                };
            }

            // Native MathML. This is what a real-world page and every EPUB3 uses (EpubParser
            // routes each spine item through here), as opposed to the `data-math` round-trip
            // contract above, which only ever appears in this library's own HTML output. Without
            // it, a `<math>` element fell through to the generic element handling below, which
            // concatenates descendant text: `<mfrac><mn>1</mn><mn>2</mn></mfrac>` became "12".
            if (tagName === 'math' || tagName.endsWith(':math')) {
                // `display="block"` is MathML's own attribute for a display equation; the legacy
                // `mode="display"` means the same thing and is still emitted by older producers.
                const isBlock = node.attributes?.['display'] === 'block'
                    || node.attributes?.['mode'] === 'display';
                const latex = mathmlTreeToLatex(toMathNode(node));
                if (isEmptyMath(latex)) return null;
                return {
                    type: 'code',
                    text: latex,
                    metadata: { math: isBlock ? 'block' : 'inline', anchorIds: anchorIds.length > 0 ? anchorIds : undefined } as CodeMetadata
                };
            }

            // Admonition: attribute-driven editors render
            // <div class="admonition admonition-note" data-type="note">…children…</div>.
            if (tagName === 'div' && (node.attributes?.class || '').split(/\s+/).includes('admonition')) {
                // Its type from `data-type`, else from a class naming one (`admonition warning`,
                // `admonition-warning`), else a note.
                const admonitionTypes = ['note', 'tip', 'important', 'warning', 'caution'] as const;
                const admonitionTypeAttr = node.attributes?.['data-type'];
                const fromClass = (node.attributes?.class || '').split(/\s+/).map(name => name.replace(/^admonition-/, '')).find(name => (admonitionTypes as readonly string[]).includes(name));
                const admonitionType = (admonitionTypes as readonly string[]).includes(admonitionTypeAttr as any)
                    ? admonitionTypeAttr as AdmonitionMetadata['admonitionType']
                    : (fromClass as AdmonitionMetadata['admonitionType'] | undefined) ?? 'note';
                const admonitionNode: OfficeContentNode = {
                    type: 'admonition',
                    metadata: { admonitionType, anchorIds: anchorIds.length > 0 ? anchorIds : undefined } as AdmonitionMetadata,
                    children: parseChildren(node, newFormatting, listContext)
                };
                if (config.includeRawContent) admonitionNode.rawContent = '<div class="admonition">...</div>';
                return admonitionNode;
            }

            // Blockquote. Previously dropped entirely (its children were lifted out unquoted), so a
            // <blockquote> lost its `> ` on the Markdown hop. Mark each block child with the 'Quote'
            // style the styleMapper maps back to a blockquote; loose inline content is wrapped in one
            // Quote-styled paragraph so it isn't emitted as an ordinary line.
            if (tagName === 'blockquote') {
                const kids = parseChildren(node, newFormatting, listContext);
                const isBlock = (t: string) => t === 'paragraph' || t === 'heading' || t === 'list';
                if (!kids.some(k => isBlock(k.type))) {
                    return { type: 'paragraph', metadata: { style: 'Quote', ...(anchorIds.length > 0 && { anchorIds }) } as any, children: kids };
                }
                kids.forEach(k => {
                    // A block quoted already (a quote in a quote) is not given the style again.
                    if (isBlock(k.type) && (k.metadata as any)?.style !== 'Quote') k.metadata = { ...(k.metadata as any), style: 'Quote' };
                });
                // The quote's id is the first of what it holds, where HtmlGenerator writes a quoted
                // paragraph's id (on its <blockquote>).
                return markedBefore(anchorIds, kids);
            }

            // Mermaid diagrams. Attribute-driven producers render a
            // <div class="mermaid" data-mermaid="<code>"> with the code also as text content.
            // Map either shape to a fenced code node with language `mermaid`, so it round-trips as
            // a ```mermaid block. Previously this div fell through to generic handling and its
            // code flattened to paragraph text.
            if (tagName === 'div' && (node.attributes?.['data-mermaid'] !== undefined || (node.attributes?.class || '').split(/\s+/).includes('mermaid'))) {
                const code = decodeEntities(rawChildText(node)).trim()
                    || (node.attributes?.['data-mermaid'] || '');
                // Only claim this as a mermaid code node when there is actual diagram source.
                // A bare `class="mermaid"` div with nested elements (a mermaid.js-rendered <svg>,
                // or a div merely reusing the class for styling) has no direct text and no
                // data-mermaid; fall through to generic handling so its content is not dropped.
                if (code) {
                    const mermaidNode: OfficeContentNode = {
                        type: 'code',
                        text: code,
                        metadata: { language: 'mermaid' } as CodeMetadata
                    };
                    if (config.includeRawContent) mermaidNode.rawContent = '<div data-mermaid>...</div>';
                    return mermaidNode;
                }
            }

            // The caption HtmlGenerator writes under a picture (its file name, which the image node
            // already holds) is a label, not document text: read as text, it grew a paragraph on every
            // save. Any other caption, a <figcaption> and a table's <caption> (which the table's branch puts
            // before it), is a block of its own, which must not run into the picture before it.
            const isClass = (name: string) => (node.attributes?.class || '').split(/\s+/).includes(name);
            if (tagName === 'figcaption' || tagName === 'caption' || (tagName === 'div' && isClass('caption'))) {
                const caption = parseChildren(node, newFormatting, listContext);
                if (!caption.some(c => c.type !== 'text' || c.text?.trim())) return [];
                const writersLabel = tagName === 'div' && node.parent?.tagName === 'div' && (node.parent.attributes?.class || '').split(/\s+/).includes('image-container')
                    && caption.every(c => c.type === 'text') && WRITER_CAPTION.test(plainTextOf(caption).trim());
                if (writersLabel) return [];
                // Blocks in it stay blocks (a <p> cannot hold a <p>); each run of inline content (text, a
                // line break, inline math, a picture) is a paragraph.
                if (!caption.some(isBlockNode)) return { type: 'paragraph', children: caption };
                const parts: OfficeContentNode[] = [];
                let run: OfficeContentNode[] = [];
                const flush = () => {
                    // Each run's edges trimmed, as a block's are (a space before the block it stood by stayed).
                    const children = config.preserveXmlWhitespace ? run : trimBlockEdges(collapseSpacesAcrossNodes(run));
                    if (children.some(c => c.type !== 'text' || c.text?.trim())) parts.push({ type: 'paragraph', children });
                    run = [];
                };
                for (const child of caption) {
                    if (isBlockNode(child)) { flush(); parts.push(child); } else run.push(child);
                }
                flush();
                return parts;
            }
            // Skip structural containers produced by HtmlGenerator to avoid deep AST nesting
            if (tagName === 'div' && (
                node.attributes?.class === 'container' ||
                node.attributes?.class === 'spreadsheet-container' ||
                node.attributes?.class === 'presentation-container' ||
                node.attributes?.class === 'pdf-container' ||
                node.attributes?.class === 'metadata-summary' ||
                node.attributes?.class === 'image-container' ||
                node.attributes?.class === 'chart-container' ||
                node.attributes?.class === 'table-container' ||
                node.attributes?.class === 'sheet' ||
                node.attributes?.class === 'page' ||
                node.attributes?.class === 'slide' ||
                node.attributes?.class === 'note-content'
            )) {
                return parseChildren(node, newFormatting, listContext);
            }
            if (tagName === 'article') {
                return parseChildren(node, newFormatting, listContext);
            }

            if (tagName === 'p' || tagName === 'div') {
                const children = parseChildren(node, newFormatting, listContext);

                // If it's a div and contains block elements, return children directly
                const hasBlockElements = children.some(c => ['paragraph', 'table', 'heading', 'list', 'image', 'chart', 'code', 'embed', 'admonition', 'definitionList'].includes(c.type));
                // A div of blocks is read through; its id marks its first block (HtmlGenerator puts a
                // picture's there, and a page's section or sheet is linked to by it).
                if (tagName === 'div' && hasBlockElements) {
                    return markedBefore(anchorIds, children);
                }

                // Flatten nested paragraphs to avoid deep AST nesting (e.g. from notes)
                const flattenedChildren: OfficeContentNode[] = [];
                for (const child of children) {
                    if (child.type === 'paragraph' && child.children) {
                        appendAll(flattenedChildren, child.children);
                    } else {
                        flattenedChildren.push(child);
                    }
                }

                const pNode: OfficeContentNode = {
                    type: 'paragraph',
                    metadata: { alignment: newFormatting.alignment, anchorIds: anchorIds.length > 0 ? anchorIds : undefined } as ParagraphMetadata,
                    children: flattenedChildren,
                    htmlAttributes: collectHtmlAttributes(node, ['align'])
                };

                if (config.includeRawContent) {
                    // Note: Since this is a manual parser without locators, we can't easily get the original source slice.
                    // We'll skip rawContent for structural nodes here unless we want to implement index tracking in parseHtmlTree.
                }

                return pNode;
            }
            if (tagName.match(/^h[1-6]$/)) {
                const level = parseInt(tagName.substring(1));
                const hNode: OfficeContentNode = {
                    type: 'heading',
                    metadata: { level, alignment: newFormatting.alignment, anchorIds: anchorIds.length > 0 ? anchorIds : undefined } as HeadingMetadata,
                    children: parseChildren(node, newFormatting, listContext),
                    htmlAttributes: collectHtmlAttributes(node, ['align'])
                };
                return hNode;
            }
            if (tagName === 'dl') {
                return {
                    type: 'definitionList',
                    ...(anchorIds.length > 0 && { metadata: { anchorIds } }),
                    children: parseChildren(node, newFormatting, listContext)
                };
            }
            if (tagName === 'dt') {
                return {
                    type: 'definitionTerm',
                    ...(anchorIds.length > 0 && { metadata: { anchorIds } }),
                    children: parseChildren(node, newFormatting, listContext)
                };
            }
            if (tagName === 'dd') {
                return {
                    type: 'definitionDescription',
                    ...(anchorIds.length > 0 && { metadata: { anchorIds } }),
                    children: parseChildren(node, newFormatting, listContext)
                };
            }
            if (tagName === 'abbr') {
                const title = node.attributes?.title;
                const children = parseChildren(node, newFormatting, listContext);
                if (title) {
                    // The innermost abbreviation's title is the one a reader sees on its text: an outer
                    // one gives it only to text without one (and allocates nothing for the rest).
                    children.forEach(c => {
                        if (c.type === 'text' && !(c.metadata as TextMetadata | undefined)?.abbreviationTitle) {
                            c.metadata = { ...c.metadata, abbreviationTitle: title } as TextMetadata;
                        }
                    });
                }
                return children;
            }
            if (tagName === 'cite' && node.attributes?.['data-citation-key'] !== undefined) {
                const citationKey = node.attributes['data-citation-key'];
                return {
                    type: 'text',
                    text: citationKey,
                    formatting: Object.keys(newFormatting).length > 0 ? { ...newFormatting } : undefined,
                    metadata: { citationKey } as TextMetadata
                };
            }
            // Attribute-driven citation shape: a <span> carrying the `citation` class token (among
            // any others) and a non-empty data-key. Produces the same bare-key text node as the
            // <cite> form above. An empty/absent data-key falls through to generic span handling,
            // so the span's visible text still survives.
            if (tagName === 'span'
                && (node.attributes?.class || '').split(/\s+/).includes('citation')
                && node.attributes?.['data-key']) {
                const citationKey = node.attributes['data-key'];
                return {
                    type: 'text',
                    text: citationKey,
                    formatting: Object.keys(newFormatting).length > 0 ? { ...newFormatting } : undefined,
                    metadata: { citationKey } as TextMetadata
                };
            }
            if (tagName === 'ul' || tagName === 'ol') {
                const isNewTopLevel = !listContext;
                const newListContext: ListContext = {
                    listId: isNewTopLevel ? `html-list-${htmlListIdCounter++}` : listContext!.listId,
                    type: tagName === 'ol' ? 'ordered' : 'unordered',
                    level: isNewTopLevel ? 0 : listContext!.level + 1,
                    counters: isNewTopLevel ? {} : { ...listContext!.counters }, // Clone to avoid side effects on parent levels
                    isTask: node.attributes?.['data-type'] === 'taskList'
                };

                // Initialize counter for this level
                if (tagName === 'ol' && node.attributes?.start) {
                    const start = parseInt(node.attributes.start, 10);
                    newListContext.counters[newListContext.level] = isNaN(start) ? 0 : start - 1;
                } else {
                    newListContext.counters[newListContext.level] = 0;
                }

                // A comment between two items (preserveComments) ends the item before it, where it stood: left
                // between them, it parted the list in two in every writer.
                const items = parseChildren(node, currentFormatting, newListContext);
                if (!items.some(isSourceComment)) return items;
                const kept: OfficeContentNode[] = [];
                let lastItem: OfficeContentNode | undefined;
                for (const item of items) {
                    if (isSourceComment(item) && lastItem) (lastItem.children ??= []).push(item);
                    else kept.push(item);
                    if (item.type === 'list') lastItem = item;
                }
                return kept;
            }
            if (tagName === 'li') {
                if (listContext) {
                    if (node.attributes?.value) {
                        const val = parseInt(node.attributes.value, 10);
                        if (!isNaN(val)) listContext.counters[listContext.level] = val;
                    } else {
                        listContext.counters[listContext.level]++;
                    }
                }

                const children = parseChildren(node, newFormatting, listContext);
                const nestedLists: OfficeContentNode[] = [selfNodePlaceholder];
                const selfChildren: OfficeContentNode[] = [];
                for (const c of children) (c.type === 'list' ? nestedLists : selfChildren).push(c);

                let isTask: boolean | undefined;
                let checked: boolean | undefined;
                if (listContext?.isTask) {
                    isTask = true;
                    const dataChecked = node.attributes?.['data-checked'];
                    checked = dataChecked !== undefined ? dataChecked === 'true' : (findNestedCheckboxChecked(node) ?? false);
                }

                const selfNode: OfficeContentNode = {
                    type: 'list',
                    text: plainTextOf(selfChildren),
                    metadata: {
                        listType: listContext?.type || 'unordered',
                        indentation: listContext?.level || 0,
                        alignment: newFormatting.alignment || 'left',
                        listId: listContext?.listId || 'html-list-none',
                        itemIndex: (listContext?.counters[listContext.level] ?? 1) - 1,
                        anchorIds: anchorIds.length > 0 ? anchorIds : undefined,
                        isTask,
                        checked
                    } as ListMetadata,
                    children: selfChildren
                };

                // The item, then the lists nested in it, in the one list built for them (a copy of the nested
                // lists at every level of nesting cost their whole length again).
                nestedLists[0] = selfNode;
                return nestedLists;
            }
            if (tagName === 'table') {
                // Attribute-driven editors render data-align on the <table> itself.
                const tableAlignAttr = node.attributes?.['data-align'];
                const tableAlign = (['left', 'center', 'right'] as const).includes(tableAlignAttr as any) ? tableAlignAttr as 'left' | 'center' | 'right' : undefined;

                // What a table holds outside its rows and cells (its <caption>, text or an element written
                // between rows, a comment) stands before it, where a browser shows it: left among the rows, a
                // caption became a header row in Markdown and the rest was lost from plain text.
                const rows: OfficeContentNode[] = [];
                const before: OfficeContentNode[] = [];
                for (const child of parseChildren(node, newFormatting, listContext)) (child.type === 'row' ? rows : before).push(child);
                const tableNode: OfficeContentNode = {
                    type: 'table',
                    metadata: { anchorIds: anchorIds.length > 0 ? anchorIds : undefined, align: tableAlign } as TableMetadata,
                    children: rows,
                    htmlAttributes: collectHtmlAttributes(node, ['data-align', 'align'])
                };
                if (config.includeRawContent) {
                    tableNode.rawContent = '<table>...</table>';
                }
                if (!before.length) return tableNode;
                before.push(tableNode);
                return before;
            }
            if (tagName === 'tr') {
                // What a row holds outside its cells goes before the table (see the table's branch).
                const cells: OfficeContentNode[] = [];
                const before: OfficeContentNode[] = [];
                for (const child of parseChildren(node, newFormatting, listContext)) (child.type === 'cell' ? cells : before).push(child);
                // The row of column letters over a sheet (HtmlGenerator's) is no row of the sheet.
                if (!cells.length && node.children.some(isSheetChrome)) return before;
                const rowNode: OfficeContentNode = {
                    type: 'row',
                    children: cells,
                    htmlAttributes: collectHtmlAttributes(node, [])
                };
                if (config.includeRawContent) {
                    rowNode.rawContent = '<tr>...</tr>';
                }
                if (!before.length) return rowNode;
                before.push(rowNode);
                return before;
            }
            if (isSheetChrome(node)) return [];
            if (tagName === 'td' || tagName === 'th') {
                // Merged cells, their spans held as a browser holds them (colspan to 1000, rowspan to
                // 65534, and both to at least 1: a `rowspan="-3"` came through as -3).
                const colSpan = cellSpan(node.attributes?.colspan, MAX_COL_SPAN);
                const rowSpan = cellSpan(node.attributes?.rowspan, MAX_ROW_SPAN);

                // Per-column GFM alignment: read the cell's own `text-align` (or a legacy `align=`
                // attribute) into `CellMetadata.align`, so the `:---`/`:---:`/`---:` markers survive
                // AST -> HTML -> AST. `justify` has no pipe-table marker, so it is not a cell align.
                // The table-level `<table data-align>` form is read separately in the `table` branch.
                const cellTextAlign = (getDeclaration(parseStyleDeclarations(node.attributes?.style || ''), 'text-align')
                    || node.attributes?.align || '').toLowerCase();
                const cellAlign = (['left', 'center', 'right'] as const).includes(cellTextAlign as any)
                    ? cellTextAlign as 'left' | 'center' | 'right'
                    : undefined;

                const cellNode: OfficeContentNode = {
                    type: 'cell',
                    metadata: {
                        colSpan: colSpan > 1 ? colSpan : undefined,
                        rowSpan: rowSpan > 1 ? rowSpan : undefined,
                        align: cellAlign,
                        anchorIds: anchorIds.length > 0 ? anchorIds : undefined
                    } as CellMetadata,
                    children: parseChildren(node, newFormatting, listContext),
                    htmlAttributes: collectHtmlAttributes(node, ['colspan', 'rowspan', 'align'])
                };
                if (config.includeRawContent) {
                    cellNode.rawContent = '<td>...</td>';
                }
                return cellNode;
            }
            if (tagName === 'img') {
                const src = node.attributes?.src;
                const alt = node.attributes?.alt;

                // Attribute-driven editors render data-width/data-align, falling back to
                // parsing the inline style for consumers that only emit the CSS.
                const imgDecls = parseStyleDeclarations(node.attributes?.style || '');
                // Exact lookup, so `max-width: 100%` - the standard responsive-image style, and by
                // far the most common inline style on an <img> - is no longer read as a declared
                // width. It constrains the rendered size; it is not an author-specified width.
                // The `width` attribute too (`<img width="200">`, as READMEs size a logo), which was dropped.
                const width = node.attributes?.['data-width'] || getDeclaration(imgDecls, 'width') || node.attributes?.width || undefined;

                // Alignment is inferred from which auto margin is present. Comparing the parsed
                // value rather than substring-matching "margin-left: 0" stops `margin-left: 0.5rem`
                // from being read as left-aligned, and lets the `margin: 0 auto` centering
                // shorthand be recognised at all.
                const marginLeft = getDeclaration(imgDecls, 'margin-left');
                const marginRight = getDeclaration(imgDecls, 'margin-right');
                const marginShorthand = getDeclaration(imgDecls, 'margin');
                const isZero = (v: string | undefined) => v !== undefined && /^0(?:[a-z%]*)$/.test(v);
                const shorthandParts = marginShorthand ? marginShorthand.split(/\s+/) : [];
                const shorthandCentres = shorthandParts.length > 1
                    && shorthandParts[shorthandParts.length - 1] === 'auto'
                    && shorthandParts[1] === 'auto';

                const alignAttr = node.attributes?.['data-align']
                    ?? (shorthandCentres ? 'center'
                        : (isZero(marginLeft) && !isZero(marginRight) ? 'left'
                            : (isZero(marginRight) && !isZero(marginLeft) ? 'right' : undefined)));
                const align = (['left', 'center', 'right'] as const).includes(alignAttr as any) ? alignAttr as 'left' | 'center' | 'right' : undefined;

                let imageNode: OfficeContentNode;
                const heldAttachment = src && !src.startsWith('data:') ? part.imageAttachment?.(src) : undefined;
                if (heldAttachment) {
                    imageNode = {
                        type: 'image',
                        metadata: {
                            attachmentName: heldAttachment,
                            altText: alt,
                            title: node.attributes?.title,
                            anchorIds: anchorIds.length > 0 ? anchorIds : undefined,
                            width,
                            align
                        } as ImageMetadata
                    };
                } else if (src?.startsWith('data:')) {
                    const match = src.match(/^data:([^;]+);base64,(.*)$/);
                    if (match && config.extractAttachments) {
                        const mimeType = match[1] as any;
                        const data = match[2];
                        const name = `image_${attachments.length + 1}.${mimeType.split('/')[1]}`;
                        attachments.push({
                            type: 'image',
                            mimeType,
                            data,
                            name,
                            extension: mimeType.split('/')[1]
                        });
                        imageNode = {
                            type: 'image',
                            metadata: {
                                attachmentName: name,
                                altText: alt,
                                title: node.attributes?.title,
                                anchorIds: anchorIds.length > 0 ? anchorIds : undefined,
                                width,
                                align
                            } as ImageMetadata
                        };
                    } else {
                        imageNode = {
                            type: 'image',
                            metadata: {
                                url: src,
                                altText: alt,
                                title: node.attributes?.title,
                                width,
                                align
                            } as ImageMetadata
                        };
                    }
                } else {
                    imageNode = {
                        type: 'image',
                        metadata: {
                            url: src,
                            altText: alt,
                            title: node.attributes?.title,
                            anchorIds: anchorIds.length > 0 ? anchorIds : undefined,
                            width,
                            align
                        } as ImageMetadata
                    };
                }

                if (config.includeRawContent) {
                    imageNode.rawContent = '<img>';
                }
                return imageNode;
            }
            if (tagName === 'a') {
                const href = node.attributes?.href;
                const wikilinkPage = node.attributes?.['data-wikilink-page'];
                const children = parseChildren(node, newFormatting, listContext);
                // A named anchor with no target (`<a id="x"></a>`, `<a name="sec">Title</a>`) marks a
                // bookmark: its id goes to the node it stands at (see resolveAnchorMarks).
                const anchorName = node.attributes?.id || node.attributes?.name;
                if (!href && wikilinkPage === undefined && node.attributes?.['data-wikilink'] === undefined && anchorName) {
                    return markedBefore([anchorName], children);
                }
                if (wikilinkPage !== undefined) {
                    children.forEach(c => {
                        if (c.type === 'text' && !(c.metadata as TextMetadata | undefined)?.link) {
                            c.metadata = { ...c.metadata, link: wikilinkPage, linkType: 'internal', wikilink: true } as TextMetadata;
                        }
                    });
                } else if (node.attributes?.['data-wikilink'] !== undefined) {
                    // Attribute-driven wikilink shape: the page lives in data-target, the display
                    // text is the anchor's own content (or data-alias/data-target when the anchor
                    // is empty). data-wikilink-page above keeps precedence over this form.
                    const page = node.attributes['data-target'] || '';
                    if (!children.some(c => c.type === 'text')) {
                        children.push({
                            type: 'text',
                            text: node.attributes['data-alias'] || node.attributes['data-target'] || '',
                        });
                    }
                    children.forEach(c => {
                        if (c.type === 'text' && !(c.metadata as TextMetadata | undefined)?.link) {
                            c.metadata = { ...c.metadata, link: page, linkType: 'internal', wikilink: true } as TextMetadata;
                        }
                    });
                } else if (href) {
                    const linkType = href.startsWith('#') ? 'internal' : 'external';
                    const linkTitle = node.attributes?.title;
                    // A link inside this one (which a browser allows across a table cell or an object) is
                    // the one its text follows: this link goes only to text without one, and allocates
                    // nothing for the rest at every level of nesting.
                    children.forEach(c => {
                        if ((c.metadata as TextMetadata | undefined)?.link) return;
                        if (c.type === 'text') {
                            c.metadata = { ...c.metadata, link: href, linkType, title: linkTitle } as TextMetadata;
                        } else if (c.type === 'image') {
                            // A linked image (`<a href><img></a>`) carries the link itself.
                            c.metadata = { ...c.metadata, link: href, linkType, ...(linkTitle !== undefined && { linkTitle }) } as ImageMetadata;
                        }
                    });
                }
                return children;
            }
            if (tagName === 'br') {
                // A <br> is a hard line break: `carriageReturn` so the Markdown generator emits a
                // hard break (`  \n`, or a `<br>` inside a table cell) that re-imports as a <br>.
                // `textWrapping` emitted a bare `\n` in a paragraph, which re-imports as a space.
                const brNode: OfficeContentNode = { type: 'break', metadata: { breakType: 'carriageReturn' } };
                if (config.includeRawContent) {
                    brNode.rawContent = '<br/>';
                }
                return brNode;
            }
            if (tagName === 'hr') {
                // A horizontal rule is a thematic break. This library tags an office page break
                // as <hr class="page-break"> on emission, so that variant round-trips back to a
                // page break; every other <hr> is thematic. Previously <hr> was dropped entirely.
                const isPageBreak = (node.attributes?.class || '').split(/\s+/).includes('page-break');
                const hrNode: OfficeContentNode = { type: 'break', metadata: { breakType: isPageBreak ? 'page' : 'thematic', anchorIds: anchorIds.length > 0 ? anchorIds : undefined } };
                if (config.includeRawContent) {
                    hrNode.rawContent = '<hr/>';
                }
                return hrNode;
            }
            // (`<xmp>` and `<listing>` are the older spellings of a preformatted block; an xmp holds its text as written.)
            if (tagName === 'pre' || tagName === 'xmp' || tagName === 'listing') {
                const codeNode = node.children.find(c => c.tagName === 'code');
                let language;
                let codeText = '';
                if (codeNode) {
                    const classAttr = codeNode.attributes?.class || '';
                    const langMatch = classAttr.split(' ').find((c: string) => c.startsWith('language-'));
                    if (langMatch) language = langMatch.replace('language-', '');
                    // Decode entities: the code body is stored raw, so `&lt;`/`&gt;`/`&amp;` (e.g. a
                    // mermaid `-->` arrow, or `a < b` in a snippet) must be turned back into text.
                }
                // All of the block's text, not the first <code>'s alone (a prompt in a <span> before it,
                // a second <code> after it, were dropped), without the line break that HTML drops
                // right after the <pre> start tag.
                codeText = decodeEntities(preformattedText(node));
                if (node.children[0]?.type === 'text' && /^\r?\n/.test(node.children[0].text || '')) codeText = codeText.replace(/^\r?\n/, '');
                // A `mermaid` class token (on the <pre> or its <code>) names the language when no
                // explicit language-* class is present - some producers emit <pre class="mermaid">.
                if (!language && (
                    (node.attributes?.class || '').split(/\s+/).includes('mermaid') ||
                    (codeNode?.attributes?.class || '').split(/\s+/).includes('mermaid')
                )) {
                    language = 'mermaid';
                }

                const preNode: OfficeContentNode = {
                    type: 'code',
                    text: codeText,
                    metadata: { language, anchorIds: anchorIds.length > 0 ? anchorIds : undefined } as CodeMetadata
                };
                if (config.includeRawContent) {
                    preNode.rawContent = '<pre>...</pre>';
                }
                return preNode;
            }

            return parseChildren(node, newFormatting, listContext);
        }

        return null;
    };

    // Footnotes, in the markup of each writer that gives them, read up front so their definitions are
    // there for references anywhere before them:
    //  - this library's: `<sup data-footnote-ref="KEY">` citing `<div data-footnote-id="KEY">` in a
    //    `<section data-footnotes>`;
    //  - GitHub's, Pandoc's and markdown-it's: a link marked as a reference (`<a data-footnote-ref
    //    href="#user-content-fn-1">`, `<a class="footnote-ref" href="#fn1">`, `role="doc-noteref"`), bare
    //    or all a `<sup>` holds, citing an item of a footnotes section's list (`<section data-footnotes>`,
    //    `class="footnotes"`, `role="doc-endnotes"`), `<li id="user-content-fn-1">`, at any depth;
    //  - EPUB 3's: `<a epub:type="noteref" href="#n1">` citing an `<aside epub:type="footnote" id="n1">`
    //    (or `role="doc-footnote"`), wherever it stands: read as a note when a reference cites it.
    // A footnotes section is left out of the text only when definitions were read from it: GitHub's was
    // left out whatever it held (its items carry no data-footnote-id), and every note in it was lost.
    const readFootnotes = (): void => {
        const isFootnotesSection = (n: HtmlNode): boolean => n.attributes?.['data-footnotes'] !== undefined || n.attributes?.role === 'doc-endnotes'
            || (['section', 'div', 'aside', 'ol'].includes(n.tagName!) && hasToken(n.attributes?.class, 'footnotes'));
        const noteKind = (n: HtmlNode): 'footnote' | 'endnote' | undefined => {
            if (!n.attributes?.id) return undefined;
            const types = (n.attributes['epub:type'] ?? '').split(/\s+/);
            if (types.includes('footnote') || n.attributes.role === 'doc-footnote') return 'footnote';
            if (types.includes('endnote') || types.includes('rearnote') || n.attributes.role === 'doc-endnote') return 'endnote';
            return undefined;
        };
        const sections: HtmlNode[] = [];
        const notes: HtmlNode[] = [];
        const cited = new Set<string>();
        const scan = (n: HtmlNode, depth: number, inSection: boolean): void => {
            for (const child of n.children) {
                if (child.type !== 'element') continue;
                const href = referenceLinkOf(child)?.attributes?.href;
                if (href?.startsWith('#')) cited.add(href.slice(1));
                let within = inSection;
                if (!inSection && isFootnotesSection(child)) { sections.push(child); within = true; }
                else if (!inSection && noteKind(child)) notes.push(child);
                if (depth < MAX_HTML_NESTING_DEPTH) scan(child, depth + 1, within);
            }
        };
        scan(body, 0, false);

        const read = (key: string, item: HtmlNode, kind: 'footnote' | 'endnote', id: string): void => {
            if (footnoteDefinitions.has(key)) return;
            // With whitespace kept as written, the one space written before this library's back-link is the
            // writer's, not the note's: left in, it grew by one each save.
            const children = item.children.slice();
            const backLinkAt = children.findIndex(isFootnoteBackLink);
            const beforeBackLink = backLinkAt > 0 ? children[backLinkAt - 1] : undefined;
            if (config.preserveXmlWhitespace && beforeBackLink?.type === 'text' && beforeBackLink.text?.endsWith(' ')) {
                children[backLinkAt - 1] = { ...beforeBackLink, text: beforeBackLink.text.slice(0, -1) };
            }
            const contentNodes: OfficeContentNode[] = [];
            readingFootnote++;
            try {
                for (const child of children) {
                    const parsed = parseNode(child);
                    if (parsed) {
                        if (Array.isArray(parsed)) appendAll(contentNodes, parsed);
                        else contentNodes.push(parsed);
                    }
                }
            } finally {
                readingFootnote--;
            }
            // The space written before the back-link ended the note's text once the link was left out,
            // and grew by one each save: the definition's edges are trimmed as a block's are.
            footnoteDefinitions.set(key, config.preserveXmlWhitespace ? contentNodes : trimBlockEdges(collapseSpacesAcrossNodes(contentNodes)));
            footnoteIds.set(key, id);
            footnoteTypes.set(key, kind);
            footnoteElements.add(item);
        };
        for (const section of sections) {
            // A section's definitions: each outermost element naming its key (`data-footnote-id`) and each
            // outermost list item with an id (GitHub's `user-content-fn-1` is note `1`, Pandoc's `fn1` too).
            let found = false;
            const collect = (n: HtmlNode, depth: number): void => {
                for (const child of n.children) {
                    if (child.type !== 'element') continue;
                    const key = child.attributes?.['data-footnote-id'] || (child.tagName === 'li' ? child.attributes?.id : undefined);
                    if (key) {
                        const id = child.attributes?.['data-footnote-id'] ? key : key.replace(/^user-content-/, '').replace(/^fn-?(?=.)/, '');
                        read(key, child, 'footnote', id);
                        found = true;
                    } else if (depth < MAX_HTML_NESTING_DEPTH) {
                        collect(child, depth + 1);
                    }
                }
            };
            collect(section, 0);
            if (found) footnoteElements.add(section);
        }
        for (const note of notes) {
            const key = note.attributes!.id;
            if (cited.has(key)) read(key, note, noteKind(note)!, key);
        }
    };
    readFootnotes();
    // Inline content written directly in the body (text, and inline elements such as <b> or <a>) is one
    // paragraph per run between blocks, as a browser lays it out in an anonymous block: not one
    // paragraph for each piece of text.
    let inlineRun: OfficeContentNode[] = [];
    const flushInlineRun = () => {
        const isBlank = (n: OfficeContentNode) => n.type === 'text' && !n.text?.trim() && !n.notes?.length && !n.comments?.length;
        // Named anchors ending a run of text stand before the block after it (`text <a id="x"></a><h2>`),
        // which they mark, as HtmlGenerator writes a block's anchors: they are not the run's.
        let end = inlineRun.length;
        while (end > 0 && (isAnchorMark(inlineRun[end - 1]) || isBlank(inlineRun[end - 1]))) end--;
        const trailingMarks = end > 0 ? inlineRun.slice(end).filter(isAnchorMark) : [];
        if (trailingMarks.length) inlineRun = inlineRun.slice(0, end);
        const children = config.preserveXmlWhitespace ? inlineRun : trimBlockEdges(collapseSpacesAcrossNodes(inlineRun));
        // Pictures and named anchors with nothing else in the run stand among the blocks, as before
        // (an anchor then marks the next block); beside text they are part of its paragraph.
        if (children.some(n => n.type === 'image' || isAnchorMark(n)) && children.every(n => n.type === 'image' || isAnchorMark(n) || isBlank(n))) {
            for (const n of children) if (!isBlank(n)) content.push(n);
        } else if (children.some(n => !isBlank(n))) content.push({ type: 'paragraph', children });
        // (A run of whitespace alone, between two blocks, lays out as nothing, whitespace kept as
        // written or not: made a paragraph, each save added an empty one.)
        for (const mark of trailingMarks) content.push(mark);
        inlineRun = [];
    };
    // A line break, inline math and a picture sit in a line of text; a rule or page break (<hr>) is a block.
    const isInlineNode = (n: OfficeContentNode) => n.type === 'text' || n.type === 'image' || isAnchorMark(n)
        || (n.type === 'break' && !['thematic', 'page'].includes((n.metadata as { breakType?: string } | undefined)?.breakType ?? ''))
        || (n.type === 'code' && (n.metadata as CodeMetadata | undefined)?.math === 'inline');
    for (let i = 0; i < body.children.length; i++) {
        const child = body.children[i];
        if (child.type === 'text' && !config.preserveXmlWhitespace) {
            const text = visibleText(body, i);
            if (text !== null) inlineRun.push({ type: 'text', text, ...(config.includeRawContent && child.text ? { rawContent: child.text } : {}) });
            continue;
        }
        const parsed = parseNode(child);
        // What a block element holds never joins the inline content around it (a container read
        // through, such as <figure> or <section>, can hand back bare text).
        const isBlockElement = child.type === 'element' && BLOCK_LEVEL_TAGS.has(child.tagName!);
        if (isBlockElement) flushInlineRun();
        for (const node of parsed ? (Array.isArray(parsed) ? parsed : [parsed]) : []) {
            if (isInlineNode(node)) inlineRun.push(node);
            else { flushInlineRun(); content.push(node); }
        }
        if (isBlockElement) flushInlineRun();
    }
    flushInlineRun();

    // Orphan footnote definitions: a `<section data-footnotes>` entry that no `<sup
    // data-footnote-ref>` consumed would otherwise be dropped (it is skipped in the body walk and
    // only materialised via a reference). Recover them as trailing `unreferenced` note nodes, the
    // same shape MarkdownParser produces, so md -> html -> md preserves the definition instead of
    // turning it into junk text with a dead back-link.
    for (const [key, definition] of footnoteDefinitions) {
        if (config.ignoreNotes) break; // ignoreNotes drops footnotes, orphan definitions included
        if (referencedFootnoteKeys.has(key)) continue;
        content.push({
            type: 'note',
            text: nodesText(definition || []),
            children: definition || [],
            metadata: { noteType: footnoteTypes.get(key) ?? 'footnote', noteId: footnoteIds.get(key) ?? key, unreferenced: true },
        });
    }

    return createAST('html', metadata, resolveAnchorMarks(content), attachments, config, undefined);
};

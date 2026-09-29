import { AdmonitionMetadata, AdmonitionSyntax, AttributeListSyntax, BreakMetadata, CitationSyntax, CodeMetadata, ConversionResult, DefinitionListSyntax, DeprecatedAdmonitionFlavor, EmbedMetadata, EmbedSyntax, FallbackToHtmlConfig, FootnoteSyntax, GeneratorConfig, HeadingMetadata, HighlightSyntax, ImageMetadata, ListMetadata, MarkdownDialectConfig, MarkdownDialectPreset, NoteMetadata, OfficeContentNode, OfficeParserAST, OfficeWarningType, StrikethroughSyntax, TableMetadata, TextFormatting, TextMetadata, WikilinkSyntax } from '../types.js';
import { escapeHtml, markdownEscapeInline, markdownEscapePlain, markdownEscapeTags, markdownEscapeText, sanitizeCommentText, sanitizeCssValue, sanitizeImageUrl, sanitizeMarkdownUrl, sanitizeUrl } from '../utils/sanitize.js';
import { isSourceComment } from '../utils/commentUtils.js';
import { base64ByteLength, resolveEmbed } from '../utils/officeGenUtils.js';
import { clampInt, clampRepeat } from '../utils/numberUtils.js';
import { BaseGenerator } from './BaseGenerator.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { TextBuilder, trimAsciiWhitespace, trimEndChars, trimStartChars } from '../utils/textUtils.js';
import { appendAll } from '../utils/nodeListUtils.js';

/**
 * A fully-resolved dialect: every capability collapsed to its canonical syntax variant (or `'none'`).
 * Deprecated boolean/flavor inputs never reach this shape - `resolveDialect` coerces them first - so
 * generator call sites test capabilities with `=== 'none'`, not truthiness (`'none'` is truthy).
 */
type ResolvedMarkdownDialect = {
    admonitions: AdmonitionSyntax;
    definitionLists: DefinitionListSyntax;
    footnotes: FootnoteSyntax;
    citations: CitationSyntax;
    wikilinks: WikilinkSyntax;
    math: 'dollar' | 'none';
    attributeLists: AttributeListSyntax;
    strikethrough: StrikethroughSyntax;
    highlight: HighlightSyntax;
    bulletListMarker: '-' | '*' | '+';
    orderedListMarker: '.' | ')';
    emphasisMarker: 'asterisk' | 'underscore';
    tables: 'native' | 'html';
};
type ResolvedFallbackToHtml = Required<FallbackToHtmlConfig>;

/**
 * Values accepted for an attribute-list `align=`. Matches what `MarkdownParser`'s own
 * `parseAttributeList` allowlists on import (plus `justify`, which HTML sources can supply),
 * so this is lossless for anything the parser produced.
 */
const MD_ALIGN_VALUES = new Set(['left', 'center', 'right', 'justify']);

/** A CSS length or percentage - the only shape `width=` legitimately carries. */
const MD_LENGTH_PATTERN = /^\d+(?:\.\d+)?(?:px|%|em|rem|pt|pc|in|cm|mm|ex|ch|vw|vh)?$/;

/** Admonition kinds, mirroring the union declared on `AdmonitionMetadata` in types.ts. */
const MD_ADMONITION_TYPES = new Set(['note', 'tip', 'important', 'warning', 'caution']);

/**
 * Folds line breaks to spaces.
 *
 * Used on values that sit inside a single-line construct (an abbreviation definition, an
 * admonition's bold title). A raw newline there does not merely look wrong: it terminates the
 * construct and exposes whatever follows as document-level Markdown.
 */
const foldLines = (value: unknown): string => String(value ?? '').replace(/[\r\n]+/g, ' ');

/**
 * A link or image title as the `"..."` of `[text](url "title")`: on one line, with its double quotes
 * as `&quot;` and its backslashes kept (the parser decodes titles, as CommonMark does), so a quote
 * cannot end the title early and spill the rest into the URL.
 */
const markdownTitle = (title: unknown): string => ` "${markdownEscapePlain(foldLines(title)).replace(/"/g, '&quot;')}"`;

/**
 * Blocks whose Markdown starts with a separator newline, so they stand apart even when they follow
 * inline content. After output that already ends in a blank line, that newline would make a second
 * blank line; {@link appendBlock} drops it.
 */
const SEPARATED_BLOCK_TYPES = new Set(['code', 'table', 'sheet', 'slide', 'page', 'embed', 'break']);

/** Nodes that are blocks of their own in Markdown (the parts of a definition list are laid out by it). */
const BLOCK_NODE_TYPES = new Set(['paragraph', 'heading', 'list', 'table', 'sheet', 'slide', 'page', 'admonition', 'definitionList', 'embed', 'header', 'footer', 'chart', 'drawing', 'slideMaster']);

/**
 * `text` as a code span: fenced by one backtick more than its longest run of them, and padded with a
 * space where it touches a backtick or has a space at both ends (which a reader strips from a span).
 */
const codeSpan = (text: string): string => {
    const longestRun = Math.max(0, ...(text.match(/`+/g) || []).map(run => run.length));
    const fence = '`'.repeat(longestRun + 1);
    const pad = text.startsWith('`') || text.endsWith('`') || (/^ [\s\S]* $/.test(text) && /[^ ]/.test(text)) ? ' ' : '';
    return `${fence}${pad}${text}${pad}${fence}`;
};

/** Whether a character is punctuation or a symbol (Unicode P or S), as CommonMark's flanking rule reads it. */
const isPunctuationCharacter = (char: string | undefined): boolean => !!char && /[\p{P}\p{S}]/u.test(char);

/** Whether a character is a word character for CommonMark's flanking rule: neither whitespace nor punctuation. */
const isWordCharacter = (char: string | undefined): boolean => !!char && !/\s/.test(char) && !isPunctuationCharacter(char);

/** The character (a whole surrogate pair) starting at index `i` of `text`, or undefined at its end. */
const characterAt = (text: string, i: number): string | undefined => {
    const code = text.codePointAt(i);
    return code === undefined ? undefined : String.fromCodePoint(code);
};

/** The character (a whole surrogate pair) ending at index `i` of `text`, or undefined at its start. */
const characterBefore = (text: string, i: number): string | undefined => {
    if (i <= 0) return undefined;
    return i >= 2 && /[\uDC00-\uDFFF]/.test(text[i - 1]) && /[\uD800-\uDBFF]/.test(text[i - 2]) ? text.slice(i - 2, i) : text[i - 1];
};

/** Nodes whose children are a line of text (in Markdown, a line break goes only in one of these, or in a node holding text directly). */
const LINE_HOLDERS = new Set(['paragraph', 'heading', 'list', 'cell', 'definitionTerm', 'definitionDescription']);

/**
 * A line break's type: a carriage return, or a text-wrapping break (a DOCX `w:br`, an ODF
 * `text:line-break`; the type a break without one has). Written as a bare newline, a text-wrapping
 * break joined its lines with a space.
 */
const isLineBreakType = (breakType: BreakMetadata['breakType'] | undefined): boolean =>
    breakType === undefined || breakType === 'carriageReturn' || breakType === 'textWrapping';

const isLineBreak = (node: OfficeContentNode): boolean =>
    node.type === 'break' && isLineBreakType((node.metadata as BreakMetadata | undefined)?.breakType);

/** A text node of whitespace alone, carrying nothing. */
const isBlankRun = (node: OfficeContentNode): boolean =>
    node.type === 'text' && !node.metadata && !node.notes?.length && !node.comments?.length && !trimAsciiWhitespace(node.text ?? '');

/** A block written in the middle of a paragraph's children: a rule, a code block or display math. */
const isBlockInLine = (node: OfficeContentNode): boolean =>
    (node.type === 'break' && (node.metadata as BreakMetadata | undefined)?.breakType === 'thematic')
    || (node.type === 'code' && (node.metadata as CodeMetadata | undefined)?.math !== 'inline');

/**
 * A paragraph's or heading's children without the line breaks (and blank runs among them) that end it
 * or stand next to a block in it: Markdown has none there, as a line ends where a block does. A
 * renderer drops trailing spaces, shows a backslash ending the last line as text, and a paragraph of
 * breaks alone as a lone backslash.
 */
const withoutBreaksAtBlocks = (children: OfficeContentNode[]): OfficeContentNode[] => {
    let out: OfficeContentNode[] | undefined;
    for (let i = 0; i < children.length;) {
        if (!isLineBreak(children[i]) && !isBlankRun(children[i])) {
            if (out) out.push(children[i]);
            i++;
            continue;
        }
        let end = i;
        let broke = false;
        while (end < children.length && (isLineBreak(children[end]) || isBlankRun(children[end]))) broke = isLineBreak(children[end++]) || broke;
        const after = children[end];
        if (broke && (after === undefined || isBlockInLine(after) || (i > 0 && isBlockInLine(children[i - 1])))) {
            out ??= children.slice(0, i);
        } else if (out) {
            for (let k = i; k < end; k++) out.push(children[k]);
        }
        i = end;
    }
    return out ?? children;
};

/** A footnote or endnote standing among blocks rather than attached to text: a definition no text refers to. */
const isStandingNote = (node: OfficeContentNode): boolean => {
    const noteType = node.type === 'note' ? (node.metadata as NoteMetadata | undefined)?.noteType : undefined;
    return noteType === 'footnote' || noteType === 'endnote';
};

/** Whether `node` is a block (a code node is one unless it is inline math). */
const isBlockNode = (node: OfficeContentNode): boolean =>
    BLOCK_NODE_TYPES.has(node.type) || (node.type === 'code' && (node.metadata as CodeMetadata | undefined)?.math !== 'inline');

/**
 * Collapses each run of blank lines (whitespace-only lines count) to one empty line, outside fenced
 * code and `$$` math blocks, whose blank lines are content. However blocks were joined (a page break,
 * a table, a slide note or a rule each brought its own separator), they end up one blank line apart:
 * a renderer shows the same either way, but the file and every diff of it stay tidy.
 */
function collapseBlankLines(markdown: string): string {
    const out: string[] = [];
    let fence: { char: string; length: number } | null = null;
    let inMath = false;
    for (const line of markdown.split('\n')) {
        if (fence) {
            out.push(line);
            const close = /^[ \t]*(`{3,}|~{3,})[ \t]*$/.exec(line);
            if (close && close[1][0] === fence.char && close[1].length >= fence.length) fence = null;
            continue;
        }
        if (inMath) {
            out.push(line);
            if (line === '$$') inMath = false;
            continue;
        }
        const open = /^[ \t]*(`{3,}|~{3,})/.exec(line);
        if (open) fence = { char: open[1][0], length: open[1].length };
        // Only a bare `$$` line opens or closes a math block, as the parser reads it (the math
        // writer indents a content line that is itself `$$`).
        else if (line === '$$') inMath = true;
        else if (line.trim() === '') {
            if (out.length > 0 && out[out.length - 1] !== '') out.push('');
            continue;
        }
        out.push(line);
    }
    return out.join('\n');
}

/**
 * Markdown lines joined into one with `br`, for a list item or table cell, which is a single line: a
 * hard break in either form (trailing spaces, or a backslash that is not itself escaped) becomes
 * `br` like any other line end.
 */
const joinLines = (markdown: string, br: string): string => {
    let out = '';
    let i = 0;
    for (;;) {
        const newline = markdown.indexOf('\n', i);
        if (newline === -1) return out + markdown.slice(i);
        const line = markdown.slice(i, newline);
        let slashes = 0;
        while (slashes < line.length && line[line.length - 1 - slashes] === '\\') slashes++;
        i = newline + 1;
        if (slashes % 2 === 1) {
            // A hard break written as a backslash: it and its line end are one `br`.
            out += line.slice(0, -1) + br;
        } else {
            // A line end, with the spaces before it and any blank lines after it, is one `br`.
            out += trimEndChars(line, ' \t') + br;
            while (markdown[i] === '\n') i++;
        }
    }
};

/**
 * A block's content without the spaces and tabs at its start and end. Markdown readers drop them
 * (CommonMark strips a paragraph's leading and trailing whitespace), so writing them only made the
 * next save differ from this one.
 */
const trimBlockEdges = (markdown: string): string => trimEndChars(trimStartChars(markdown, ' \t'), ' \t');

/** Appends a node's Markdown, leaving exactly one blank line before a separated block. */
const appendBlock = (output: TextBuilder, node: { type: string }, next: string): void =>
    output.append(SEPARATED_BLOCK_TYPES.has(node.type) && output.endsWith('\n\n') ? trimStartChars(next, '\n') : next);

/**
 * Named Markdown dialect presets. `extended` reproduces this library's historical output
 * exactly (every feature on, GitHub-style admonitions) - the backward-compatibility anchor.
 */
const MARKDOWN_DIALECT_PRESETS: Record<MarkdownDialectPreset, ResolvedMarkdownDialect> = {
    extended: { admonitions: 'blockquote', definitionLists: 'colon', footnotes: 'caret', citations: 'at', wikilinks: 'double-bracket', math: 'dollar', attributeLists: 'brace', strikethrough: 'tilde', highlight: 'equals', bulletListMarker: '-', orderedListMarker: '.', emphasisMarker: 'asterisk', tables: 'native' },
    github: { admonitions: 'blockquote', definitionLists: 'none', footnotes: 'caret', citations: 'none', wikilinks: 'none', math: 'dollar', attributeLists: 'none', strikethrough: 'tilde', highlight: 'none', bulletListMarker: '-', orderedListMarker: '.', emphasisMarker: 'asterisk', tables: 'native' },
    gitlab: { admonitions: 'fence', definitionLists: 'none', footnotes: 'caret', citations: 'none', wikilinks: 'none', math: 'dollar', attributeLists: 'none', strikethrough: 'tilde', highlight: 'none', bulletListMarker: '-', orderedListMarker: '.', emphasisMarker: 'asterisk', tables: 'native' },
    obsidian: { admonitions: 'blockquote', definitionLists: 'none', footnotes: 'caret', citations: 'none', wikilinks: 'double-bracket', math: 'dollar', attributeLists: 'none', strikethrough: 'tilde', highlight: 'equals', bulletListMarker: '-', orderedListMarker: '.', emphasisMarker: 'asterisk', tables: 'native' },
    pandoc: { admonitions: 'fence-attribute', definitionLists: 'colon', footnotes: 'caret', citations: 'at', wikilinks: 'none', math: 'dollar', attributeLists: 'brace', strikethrough: 'tilde', highlight: 'none', bulletListMarker: '-', orderedListMarker: '.', emphasisMarker: 'asterisk', tables: 'native' },
    commonmark: { admonitions: 'none', definitionLists: 'none', footnotes: 'none', citations: 'none', wikilinks: 'none', math: 'none', attributeLists: 'none', strikethrough: 'none', highlight: 'none', bulletListMarker: '-', orderedListMarker: '.', emphasisMarker: 'asterisk', tables: 'html' },
};

/**
 * Normalizes `MdGeneratorConfig.dialect` into a fully-resolved preset. A string names a preset
 * directly; an object's `extends` field (default `'extended'`) names the base preset that any
 * omitted field falls back to - NOT "whatever preset was ambient before", since config merging
 * replaces the whole `dialect` field rather than layering an object on top of a prior string.
 */
/**
 * Coerces a per-capability field to its canonical syntax variant. An omitted value inherits `base`;
 * a deprecated boolean maps `true` -> `onValue` and `false` -> `'none'` (the two are the only legacy
 * inputs, dropped next major); an explicit syntax string passes through unchanged.
 */
function resolveToggle<T extends string>(value: T | boolean | undefined, onValue: T, base: T): T {
    if (value === undefined) return base;
    if (value === true) return onValue;
    if (value === false) return 'none' as T;
    return value;
}

/** Maps the deprecated admonition flavor aliases to their syntax names; passes syntax names through. */
function resolveAdmonitions(value: AdmonitionSyntax | DeprecatedAdmonitionFlavor | undefined, base: AdmonitionSyntax): AdmonitionSyntax {
    switch (value) {
        case undefined: return base;
        case 'github': return 'blockquote';
        case 'gitlab': return 'fence';
        case 'pandoc': return 'fence-attribute';
        default: return value;
    }
}

function resolveDialect(dialect: MarkdownDialectPreset | MarkdownDialectConfig | undefined): ResolvedMarkdownDialect {
    if (dialect === undefined) return MARKDOWN_DIALECT_PRESETS.extended;
    if (typeof dialect === 'string') return MARKDOWN_DIALECT_PRESETS[dialect] ?? MARKDOWN_DIALECT_PRESETS.extended;

    const base = MARKDOWN_DIALECT_PRESETS[dialect.extends ?? 'extended'] ?? MARKDOWN_DIALECT_PRESETS.extended;
    return {
        admonitions: resolveAdmonitions(dialect.admonitions, base.admonitions),
        definitionLists: resolveToggle(dialect.definitionLists, 'colon', base.definitionLists),
        footnotes: resolveToggle(dialect.footnotes, 'caret', base.footnotes),
        citations: resolveToggle(dialect.citations, 'at', base.citations),
        wikilinks: resolveToggle(dialect.wikilinks, 'double-bracket', base.wikilinks),
        math: dialect.math ?? base.math,
        attributeLists: resolveToggle(dialect.attributeLists, 'brace', base.attributeLists),
        strikethrough: resolveToggle(dialect.strikethrough, 'tilde', base.strikethrough),
        highlight: dialect.highlight ?? base.highlight,
        bulletListMarker: dialect.bulletListMarker ?? base.bulletListMarker,
        orderedListMarker: dialect.orderedListMarker ?? base.orderedListMarker,
        emphasisMarker: dialect.emphasisMarker ?? base.emphasisMarker,
        tables: dialect.tables ?? base.tables,
    };
}

/**
 * Normalizes `MdGeneratorConfig.fallbackToHtml` into a fully resolved object, mirroring
 * `HtmlGenerator`'s `resolveStandalone()` pattern: `true`/undefined turns every part on; `false`
 * turns every part off; an object's omitted fields default to on.
 */
function resolveFallbackToHtml(fallbackToHtml: boolean | FallbackToHtmlConfig | undefined): ResolvedFallbackToHtml {
    const uniform = (on: boolean): ResolvedFallbackToHtml => ({
        // inlineFormatting is opt-in only: it is never enabled by the boolean form, since it changes
        // default output. Every other field follows the boolean.
        textFormatting: on, alignment: on, anchors: on, tables: on, embeds: on, cellLineBreaks: on,
        itemLineBreaks: on,
        inlineFormatting: false,
    });
    if (fallbackToHtml === undefined || typeof fallbackToHtml === 'boolean') return uniform(fallbackToHtml ?? true);
    const on = uniform(true);
    return {
        textFormatting: fallbackToHtml.textFormatting ?? on.textFormatting,
        alignment: fallbackToHtml.alignment ?? on.alignment,
        anchors: fallbackToHtml.anchors ?? on.anchors,
        tables: fallbackToHtml.tables ?? on.tables,
        embeds: fallbackToHtml.embeds ?? on.embeds,
        cellLineBreaks: fallbackToHtml.cellLineBreaks ?? on.cellLineBreaks,
        itemLineBreaks: fallbackToHtml.itemLineBreaks ?? on.itemLineBreaks,
        inlineFormatting: fallbackToHtml.inlineFormatting ?? on.inlineFormatting,
    };
}

/**
 * Generates Markdown from an AST.
 * 
 * DESIGN PRINCIPLES:
 * 1. **Strict Native Preference**: Always utilize native Markdown syntax for features that 
 *    are natively supported (headings, lists, bold/italic, etc.). HTML tags should NEVER 
 *    be used for these features.
 * 
 * 2. **Fidelity vs. Purity (The `fallbackToHtml` Principle)**:
 *    - When a given `fallbackToHtml` part is TRUE: The generator prioritizes high-fidelity
 *      document conversion for that part. It will use HTML tags for features that Markdown
 *      cannot natively represent (e.g., `<u>` for underline, `<div>` for alignment, `<table>`
 *      for nested structures or merged cells).
 *    - When FALSE: The generator prioritizes "pure" Markdown for that part.
 *      Unsupported features are either:
 *      - **Skipped**: Non-essential formatting like underline, subscript, superscript,
 *        or text alignment is omitted.
 *      - **Simplified/Hoisted**: Complex structures like nested tables are hoisted out
 *        of their parent cells and rendered as separate sequential tables to maintain
 *        valid Markdown syntax.
 *
 * 3. **Consistency**: All similar structural or formatting ideological problems must be
 *    resolved using these same rules to ensure predictable output.
 *
 * 4. **Dialect (`MdGeneratorConfig.dialect`)**: A second, independent axis from `fallbackToHtml` -
 *    which *native* Markdown syntax to emit for constructs with more than one real-world
 *    convention (admonitions, definition lists, footnotes, citations, wikilinks, math, list/
 *    emphasis markers, tables). See `resolveDialect()` and `MARKDOWN_DIALECT_PRESETS` above.
 */
export class MarkdownGenerator extends BaseGenerator<'md'> {
    private isInsideTable = false;
    /**
     * Set while rendering the children of a heading, or the cells of a table's header row.
     *
     * Markdown already conveys "this is a heading" with `#` and "this is a header row" with the
     * separator line, so a run inside one that also carries bold - the normal case for ODF, whose
     * heading and header-row paragraph styles are bold and are now inherited by their runs - would
     * render as `# **Heading**` and `| **ITEM** |`. That is redundant rather than wrong, but it
     * also round-trips back into bold text nodes nested inside a heading, so the noise compounds
     * on every parse/generate cycle. Emphasis the node type already implies is dropped; every
     * other formatting flag still comes through.
     */
    private inImplicitBold = false;
    /** Rendering a pipe-table cell, where a fenced block cannot go: code there stays an inline span. */
    private inPipeTableCell = false;
    /**
     * Anchors of an empty paragraph (a bookmark on an empty line), written with those of the next block
     * that writes anchors: where the parser gives an anchor standing alone, so a save reads back as itself.
     */
    private pendingAnchorIds: string[] = [];
    /** How many list items are being written: an item is one line of Markdown (see the `code` case). */
    private inListItem = 0;
    /** How many definition terms or descriptions are being written: each is one line too. */
    private inDefinition = 0;
    /** How many nodes holding a line of text (a paragraph, heading, item, cell, definition) are being written. */
    private lineDepth = 0;
    /** For each definition list being written, innermost last: whether it is in one line (see inOneLine). */
    private definitionListsInLine: boolean[] = [];
    /** Each list item's written depth (see assignListDepths). */
    private listDepths = new WeakMap<OfficeContentNode, number>();
    /**
     * The node being rendered starts a line of a paragraph (or of the document): text written there
     * must not read as a block marker, and a hard break there takes the backslash form. Set before
     * each child is rendered.
     */
    private atLineStart = false;
    /** Rendering the cells of an HTML table (the fallback for merged cells, or the `html` table dialect): their content is HTML. */
    private inHtmlTable = 0;
    /** The last character written before the node being rendered, within its parent ('' for none). */
    private previousOutputChar = '';
    /** The sibling after the node being rendered, if any. */
    private nextSibling: OfficeContentNode | undefined;
    private hoistedContent: string[] = [];
    private collectedAbbreviations = new Map<string, string>();
    private resolvedDialect: ResolvedMarkdownDialect;
    private resolvedFallbackToHtml: ResolvedFallbackToHtml;
    private resolvedEmbeds: EmbedSyntax;

    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'md'>) {
        super('md', ast, config);
        this.resolvedDialect = resolveDialect(this.config.mdConfig.dialect);
        this.resolvedFallbackToHtml = resolveFallbackToHtml(this.config.mdConfig.fallbackToHtml);
        // `dialect.embeds` is the authority for embed form. It lives in a different config object
        // than the deprecated `fallbackToHtml.embeds` boolean, and an explicit boolean `false` must
        // still win over the preset default, so it is resolved here rather than in `resolveDialect`:
        // an embeds value set on the dialect OBJECT wins; otherwise the boolean maps (`true`/unset ->
        // `'html'`, `false` -> `'link'`); otherwise the default `'html'`.
        const dialectCfg = this.config.mdConfig.dialect;
        const explicitEmbeds = (dialectCfg && typeof dialectCfg === 'object') ? dialectCfg.embeds : undefined;
        this.resolvedEmbeds = explicitEmbeds ?? (this.resolvedFallbackToHtml.embeds ? 'html' : 'link');
    }

    /**
     * Renders anchor tags if HTML fallback is allowed.
     */
    private renderAnchors(metadata: any, takePending = true): string {
        if (!this.resolvedFallbackToHtml.anchors || this.config.ignoreInternalLinks) return '';
        const ids = [...(takePending ? this.pendingAnchorIds.splice(0) : []), ...(metadata?.anchorIds || [])];
        return this.anchorSlugs(ids).map(id => `<a id="${id}"></a>`).join('');
    }

    /**
     * `ids` as written: slugified, each once (the last time it appears), and without those that slugify
     * to nothing, which are no id (`<a id="">` is read back as text).
     */
    private anchorSlugs(ids: string[]): string[] {
        const seen = new Set<string>();
        const slugs: string[] = [];
        for (let i = ids.length - 1; i >= 0; i--) {
            const slug = this.slugify(ids[i]);
            if (slug && !seen.has(slug)) {
                seen.add(slug);
                slugs.push(slug);
            }
        }
        return slugs.reverse();
    }

    /**
     * A node's ids in a table written as HTML: the first as the element's `id`, the rest as empty
     * anchors (inside a container element, before a picture or equation), slugified as the Markdown
     * anchors are. They were left out, and links to them went nowhere.
     */
    private htmlIds(metadata: any): { attr: string; extra: string } {
        if (!this.resolvedFallbackToHtml.anchors || this.config.ignoreInternalLinks) return { attr: '', extra: '' };
        const ids = this.anchorSlugs(metadata?.anchorIds || []);
        if (!ids.length) return { attr: '', extra: '' };
        return { attr: ` id="${escapeHtml(ids[0])}"`, extra: ids.slice(1).map(id => `<a id="${escapeHtml(id)}"></a>`).join('') };
    }

    /**
     * Anchors of a node in a line the parser gives them to its container: a paragraph's inline math, or
     * anything in a list item, cell or definition (one line of Markdown). An anchor there reads back as
     * the container's, so it is written where the container's are, at the line's start (they wait in
     * pendingAnchorIds, which the container's own anchors take), rather than moving there on the next
     * save.
     */
    private deferAnchors(metadata: any): string {
        if (this.resolvedFallbackToHtml.anchors && !this.config.ignoreInternalLinks && metadata?.anchorIds?.length) appendAll(this.pendingAnchorIds, metadata.anchorIds);
        return '';
    }

    /**
     * A block's anchors (and those waiting, see pendingAnchorIds) on a line of their own before it: a
     * blank line either side, so they neither join the text before nor make it a heading's underline.
     */
    private anchorsBefore(metadata: any): string {
        const anchors = this.renderAnchors(metadata);
        return anchors ? `\n\n${anchors}\n\n` : '';
    }

    /** Writing a list item, pipe-table cell or definition: one line of Markdown, where no block can go. */
    private get inOneLine(): boolean {
        return this.inPipeTableCell || this.inListItem > 0 || this.inDefinition > 0;
    }

    /**
     * Serializes a frontmatter array as a YAML flow sequence (e.g. `[a, b]`), matching
     * MarkdownParser's frontmatter array handling. Plain strings are left bare; anything
     * that would break flow-array syntax (or isn't a string) falls back to JSON encoding.
     */
    private serializeFrontmatterArray(arr: any[]): string {
        // An item holding `<` is JSON-encoded with it written `\u003c`, as scalars are (a tag stays text).
        const items = arr.map(item =>
            (typeof item === 'string' && item.trim() === item && !/[,[\]<"]/.test(item))
                ? item
                : JSON.stringify(item).replace(/</g, '\\u003c')
        );
        return `[${items.join(', ')}]`;
    }

    /**
     * Renders a Pandoc-style attribute list (e.g. `{width=50% align=left}`) from
     * ImageMetadata/TableMetadata's width/align fields - the canonical form is always
     * `key=value`, matching MarkdownParser's own vocabulary (MARKDOWN_DIALECT.md §15).
     */
    private renderAttributeList(meta: { width?: string; align?: string } | undefined, options: { skipAlign?: boolean } = {}): string {
        if (this.resolvedDialect.attributeLists === 'none') return '';
        if (!meta) return '';
        const align = options.skipAlign ? undefined : meta.align;
        if (!meta.width && !align) return '';
        const parts: string[] = [];
        // Allowlist, not escape. These land in `metadata.width`/`align` on reparse, which the
        // parser does NOT entity-decode, so encoding here would not round-trip - and stripping
        // alone is not enough: the previous `[{}\s]+` guard removed whitespace, which stops
        // `<img src=x onerror=…>` but not the slash-separated `<img/src=x/onerror=…>`.
        // Both values have a small, fully-known shape, so matching that shape is both safer and
        // lossless for anything a parser can produce.
        //
        // (`isValidContainerWidth` in utils/configUtils.ts is a near-identical regex, but it is a
        // config validator that also accepts 'auto' and numbers; importing configUtils here for
        // one pattern would be a worse coupling than this local constant.)
        if (meta.width && MD_LENGTH_PATTERN.test(String(meta.width).trim())) {
            parts.push(`width=${String(meta.width).trim()}`);
        }
        if (align && MD_ALIGN_VALUES.has(String(align).trim().toLowerCase())) {
            parts.push(`align=${String(align).trim().toLowerCase()}`);
        }
        if (parts.length === 0) return '';
        return `{${parts.join(' ')}}`;
    }

    /**
     * A picture wrapped in the link it is (a badge, `[![alt](src)](target "title")`), by the rules a
     * linked run of text follows: an internal link is dropped under `ignoreInternalLinks` and points at
     * the slugged id where ids are generated, and the target is scheme-checked.
     */
    private linkedImage(meta: ImageMetadata | undefined, picture: string): string {
        if (!meta?.link) return picture;
        const isInternal = meta.linkType !== 'external';
        if (this.config.ignoreInternalLinks && isInternal) return picture;
        let link = meta.link;
        if (isInternal && link.startsWith('#') && (this.config.generateIds || this.resolvedFallbackToHtml.anchors)) link = '#' + this.slugify(link.substring(1));
        return `[${picture}](${sanitizeMarkdownUrl(link)}${meta.linkTitle ? markdownTitle(meta.linkTitle) : ''})`;
    }

    /**
     * Whether `_` emphasis can close where the node being rendered ends: after its own trailing
     * whitespace, or before a next sibling that starts with whitespace or punctuation (a leading `_`
     * of its text is escaped) and has no emphasis of its own, whose delimiter would join the closing
     * run. A sibling written as a code span, link or tag starts with punctuation.
     */
    private nextSiblingStartsCleanly(trail: string): boolean {
        const next = this.nextSibling;
        if (trail || !next || next.type !== 'text') return true;
        const formatting = next.formatting;
        if (formatting?.bold || formatting?.italic) return false;
        if (formatting?.font === 'monospace' || (next.metadata as TextMetadata | undefined)?.link) return true;
        const first = (next.text || '')[0];
        return first === undefined || /[\s!-/:-@[-`{-~]/.test(first);
    }

    /**
     * Whether what the next sibling writes starts with a word character (neither whitespace nor
     * punctuation), right where the node being rendered ends with no whitespace of its own: plain text
     * starting with one. A sibling written formatted, linked or as anything but text starts with a
     * delimiter, a bracket or a tag, which are punctuation.
     */
    private nextSiblingStartsWithWord(trail: string): boolean {
        const next = this.nextSibling;
        const meta = next?.metadata as TextMetadata | undefined;
        if (trail || !next || next.type !== 'text' || meta?.link || meta?.citationKey) return false;
        if (this.config.includeFormatting && next.formatting && this.markdownFormattingKey(next.formatting) !== this.markdownFormattingKey({})) return false;
        return isWordCharacter(characterAt(next.text || '', 0));
    }

    /** Converts a document-supplied date to an ISO string, or '' if invalid
     *  (a malformed date would otherwise throw a RangeError and abort generation). */
    private toIsoDate(value: unknown): string {
        if (value === undefined || value === null || value === '') return '';
        const d = new Date(value as any);
        return isNaN(d.getTime()) ? '' : d.toISOString();
    }

    /**
     * Generates Markdown string from the provided AST.
     * 
     * @returns A Markdown string
     */
    /**
     * Source comments in a paragraph's or heading's line: a line break in one ends the line's block
     * when what follows could start a block of its own (a blank line, `<div>`, `>`, `#`), and the rest
     * of the comment was then Markdown, or live HTML, in a renderer. See inlineCommentText.
     */
    private readonly inlineComments = new Set<OfficeContentNode>();

    private inlineCommentText(node: OfficeContentNode): string {
        const text = node.text || '';
        // (Parsers never give an inline comment such a line: only a hand-built AST can.)
        return this.inlineComments.has(node) && /\n[ \t]*(?:\n|$)|\n {0,3}(?:[<>#=+*_~`|:-]|\d{1,9}[.)])/.test(text) ? text.replace(/\r?\n/g, ' ') : text;
    }

    async generate(): Promise<ConversionResult<'md'>> {
        let output = '';
        {
            const seen = new Set<OfficeContentNode>();
            const stack: OfficeContentNode[] = [...this.ast.content];
            while (stack.length) {
                const node = stack.pop()!;
                if (!node || seen.has(node)) continue;
                seen.add(node);
                if (node.type === 'paragraph' || node.type === 'heading') for (const child of node.children ?? []) if (isSourceComment(child)) this.inlineComments.add(child);
                for (const list of [node.children, node.notes, node.comments]) if (list) appendAll(stack, list);
            }
        }

        // Add Metadata (YAML Front Matter)
        const meta = this.effectiveMetadata;
        if (meta) {
            // Build the field lines first. JSON-encode scalar values so a title/author/description
            // containing a quote or newline can't break out of the YAML string and inject arbitrary
            // front-matter keys. (JSON.stringify of a benign value yields the same `"..."` form as
            // before, so normal output is unchanged.)
            let fields = '';
            // JSON-encoded, with `<` written `\u003c`: a renderer that does not know front matter shows
            // it as text, and a tag in a title would otherwise be live there.
            const scalar = (value: unknown) => JSON.stringify(value).replace(/[<\u2028\u2029]/g, c => `\\u${c.charCodeAt(0).toString(16).padStart(4, '0')}`);
            if (meta.title) fields += `title: ${scalar(meta.title)}\n`;
            if (meta.author) fields += `author: ${scalar(meta.author)}\n`;
            const createdIso = this.toIsoDate(meta.created);
            if (createdIso) fields += `created: ${createdIso}\n`;
            const modifiedIso = this.toIsoDate(meta.modified);
            if (modifiedIso) fields += `modified: ${modifiedIso}\n`;
            if (meta.description) fields += `description: ${scalar(meta.description)}\n`;
            if (meta.subject) fields += `subject: ${scalar(meta.subject)}\n`;
            if (meta.keywords) fields += `keywords: ${scalar(meta.keywords)}\n`;

            if (meta.customProperties) {
                for (const [key, val] of Object.entries(meta.customProperties)) {
                    // Strip newlines/colons from the key so it can't inject a new mapping.
                    // A key is a name, where `<` carries no meaning: dropped rather than encoded.
                    const safeKey = String(key).replace(/[\r\n:<]+/g, ' ').trim();
                    fields += `${safeKey}: ${Array.isArray(val) ? this.serializeFrontmatterArray(val) : scalar(val)}\n`;
                }
            }

            // Only emit the frontmatter fence when at least one field is present. Empty metadata
            // (a bare Tiptap/HTML fragment with no <head>) would otherwise emit `---\n---`, which
            // reparses as a setext `## ---` heading and corrupts the document on every save/reload.
            if (fields) {
                output += `---\n${fields}---\n\n`;
            }
        }

        const processor = async (node: OfficeContentNode, childrenOutput: string): Promise<string> => {
            // Handle Style Mapping for Markdown using the semantic mapping helper
            const mapping = this.getSemanticMapping(node);
            if (mapping) {
                // Map common HTML tags to Markdown equivalents
                // A quote cannot go in one line (a list item, cell or definition): its text is written.
                if (mapping.tag === 'blockquote') return this.inOneLine
                    ? `${this.deferAnchors(node.metadata)}${trimBlockEdges(childrenOutput)}\n\n`
                    : `> ${this.renderAnchors(node.metadata)}${trimBlockEdges(childrenOutput)}\n\n`;
                if (mapping.tag === 'code') return `\`${childrenOutput}\` `;
                if (mapping.tag === 'pre') return `\`\`\`\n${childrenOutput}\n\`\`\`\n\n`;

                const hMatch = mapping.tag.match(/^h([1-6])$/);
                if (hMatch) {
                    const level = parseInt(hMatch[1]);
                    return `${'#'.repeat(level)} ${trimBlockEdges(childrenOutput)}\n\n`;
                }
            }

            switch (node.type) {
                case 'text': {
                    // Escaped so the text reads back as itself: its Markdown characters get a
                    // backslash (a block marker too, where it starts a line), and a tag-opening `<`
                    // is entity-encoded, so document text can't inject a raw HTML tag (e.g.
                    // <script>) when the Markdown is rendered to HTML.
                    // Where a paragraph line starts, its leading spaces are dropped, as a reader drops
                    // them (four would make it a code block).
                    // So are the spaces ending a line before a block in the paragraph (a code block, a
                    // rule) or a line break (a reader takes them as part of the break).
                    let source = this.atLineStart ? (node.text || '').replace(/^[ \t]+/, '') : node.text || '';
                    if (this.nextSibling && (isBlockInLine(this.nextSibling) || isLineBreak(this.nextSibling))) source = trimEndChars(source, ' \t');
                    let text = markdownEscapeInline(source, this.atLineStart);
                    if (this.config.includeFormatting && node.formatting) {
                        // Inline code: re-wrap the RAW text in backticks. The content is literal
                        // inside a code span, so the entity-escaped form above must not show through.
                        // The fence is one backtick longer than the longest embedded run so an inner
                        // backtick can't close the span early, padded when the content touches a
                        // backtick. Done before emphasis so bold/italic wrap the span (`**`code`**`).
                        // Previously a monospace text node emitted its bare text, dropping the code.
                        if (node.formatting.font === 'monospace') {
                            // A reader strips one space from each end of a span that has one at both
                            // (and is not all spaces), so such content is padded too, to keep its spaces.
                            text = codeSpan(node.text || '');
                        }
                        // Delimiters go around the text without its surrounding whitespace, which
                        // stays outside: a delimiter next to a space cannot open or close emphasis
                        // (CommonMark's flanking rule), so `**Note: **` would not be bold at all.
                        // Whitespace alone takes no delimiters.
                        const core = node.formatting.font === 'monospace' ? text : text.trim();
                        const lead = core === text ? '' : text.slice(0, text.length - text.trimStart().length);
                        const trail = core === text || !core ? '' : text.slice(text.trimEnd().length);
                        // A backslash that ended in a space (so was left as it is) now stands before the
                        // closing delimiter, which it would escape: it is escaped itself.
                        text = core !== text && /(?:^|[^\\])(?:\\\\)*\\$/.test(core) ? `${core}\\` : core;
                        // `_` emphasis must start and end at a word boundary, and must not touch
                        // another underscore run (the two would read as one run): intraword emphasis,
                        // or emphasis next to other emphasis, takes the `*` form instead.
                        const underscoresFit = !!core
                            && !/[^\s!-/:-@[-^`{-~]/.test(lead ? ' ' : this.previousOutputChar || ' ')
                            && this.nextSiblingStartsCleanly(trail);
                        const emphasisAsterisk = this.resolvedDialect.emphasisMarker === 'asterisk' || !underscoresFit;
                        // `==text==` highlight, in dialects that define it (Obsidian/extended). A plain
                        // highlight (the default yellow) always becomes `==text==`; a highlight carrying
                        // a SPECIFIC colour stays a background-color <span> when `inlineFormatting` is on,
                        // so its exact colour survives. With `inlineFormatting` off (no span to hold it)
                        // even a coloured highlight degrades to `==` rather than being dropped. In
                        // GFM/CommonMark `==` is literal, so a highlight falls through to the <span> path.
                        const isDefaultHighlight = node.formatting.backgroundColor === '#ffff00';
                        const emitHighlightMark = !!node.formatting.backgroundColor && this.resolvedDialect.highlight !== 'none'
                            && (isDefaultHighlight || !this.resolvedFallbackToHtml.inlineFormatting);
                        const emphasis = !!core && ((node.formatting.bold && !this.inImplicitBold) || !!node.formatting.italic);
                        const strike = !!core && !!node.formatting.strikethrough && this.resolvedDialect.strikethrough !== 'none';
                        const highlight = !!core && emitHighlightMark;
                        // CommonMark reads a delimiter run with punctuation on its inner side only where
                        // whitespace or punctuation stands on its outer side (`*Source:*Data` and
                        // ``a**`b`**`` hold no emphasis), and reads a run right after the previous
                        // run's closing one of its character as one run with it (``**`a`****.**`` is
                        // bold with `****` in it). There, where the fallback allows it, the formatting is
                        // written as the HTML elements that read back the same. (The inner side of an
                        // outer run is the run inside it; an HTML wrapper is punctuation.)
                        const runs = (emphasis ? 1 : 0) + (strike ? 1 : 0) + (highlight ? 1 : 0);
                        const wrapped = this.resolvedFallbackToHtml.textFormatting && !!(node.formatting.underline || node.formatting.subscript || node.formatting.superscript);
                        const outerDelimiter = highlight ? '=' : strike ? '~' : emphasisAsterisk ? '*' : '_';
                        const asElements = runs > 0 && !wrapped && this.resolvedFallbackToHtml.textFormatting
                            && ((!lead && this.previousOutputChar === outerDelimiter)
                                || ((runs > 1 || isPunctuationCharacter(characterAt(core, 0))) && isWordCharacter(lead ? ' ' : this.previousOutputChar))
                                || ((runs > 1 || isPunctuationCharacter(characterBefore(core, core.length))) && this.nextSiblingStartsWithWord(trail)));
                        if (core && node.formatting.bold && !this.inImplicitBold) text = asElements ? `<strong>${text}</strong>` : emphasisAsterisk ? `**${text}**` : `__${text}__`;
                        if (core && node.formatting.italic) text = asElements ? `<em>${text}</em>` : emphasisAsterisk ? `*${text}*` : `_${text}_`;
                        if (strike) text = asElements ? `<del>${text}</del>` : `~~${text}~~`;
                        if (highlight) text = asElements ? `<mark>${text}</mark>` : `==${text}==`;
                        text = lead + text + trail;

                        // Use HTML tags for formatting not natively supported by standard Markdown
                        if (this.resolvedFallbackToHtml.textFormatting) {
                            if (node.formatting.underline) text = `<u>${text}</u>`;
                            if (node.formatting.subscript) text = `<sub>${text}</sub>`;
                            if (node.formatting.superscript) text = `<sup>${text}</sup>`;
                        }

                        // Inline color / highlight / font size have no Markdown syntax; emit a styled
                        // <span> (outermost, so the inner Markdown markers survive) only when opted in,
                        // so default output is unchanged. Values are CSS-sanitized against injection.
                        if (this.resolvedFallbackToHtml.inlineFormatting) {
                            const styles: string[] = [];
                            const pushStyle = (prop: string, val: string | undefined) => {
                                if (!val) return;
                                const safe = sanitizeCssValue(val); // drops url()/expression()/<>/quotes
                                if (safe) styles.push(`${prop}: ${safe}`);
                            };
                            pushStyle('color', node.formatting.color);
                            // Skip the background-color only when it was already emitted as `==text==`
                            // above; a specific-colour highlight in a highlight dialect still keeps its
                            // exact colour here.
                            if (!emitHighlightMark) pushStyle('background-color', node.formatting.backgroundColor);
                            pushStyle('font-size', node.formatting.size);
                            if (styles.length) text = `<span style="${styles.join('; ')}">${text}</span>`;
                        }
                    }
                    const meta = node.metadata as TextMetadata;
                    if (meta?.wikilink && this.resolvedDialect.wikilinks !== 'none') {
                        // Obsidian syntax: bare page name, or page|alias when the display
                        // text differs from the page name. Strip the `[]|`/newline chars
                        // that would break out of the `[[...]]` wrapper.
                        // The alias must be built from the ESCAPED text, not from raw node.text.
                        // Rebuilding from the raw value here discarded the markdownEscapeText()
                        // applied above, so a wikilink was the one place document text reached
                        // the output unescaped. Escaping is lossless for the alias specifically,
                        // because it lands back in a text node, which the parser entity-decodes.
                        const alias = markdownEscapeText(node.text || '').replace(/[[\]|\r\n]+/g, '');
                        // `page` lands in metadata.link, which is NOT entity-decoded on reparse,
                        // so it gets `<` dropped rather than encoded - a page name is an
                        // identifier, and `<` carries no meaning in one.
                        const page = (meta.link || '').replace(/[[\]|<\r\n]+/g, '');
                        text = (node.text && node.text !== (meta.link || '')) ? `[[${page}|${alias}]]` : `[[${page}]]`;
                    } else if (meta?.link) {
                        const isInternal = meta.linkType !== 'external';
                        if (!this.config.ignoreInternalLinks || !isInternal) {
                            let link = meta.link;
                            // Slugify internal link targets to match heading IDs if generating IDs
                            if (isInternal && link.startsWith('#') && (this.config.generateIds || this.resolvedFallbackToHtml.anchors)) {
                                const target = link.substring(1);
                                link = '#' + this.slugify(target);
                            }
                            // Reject javascript:/data: schemes and encode `()`/whitespace so the
                            // URL can't break out of `](...)` or inject a script link. An advisory
                            // title follows as `"title"`, in the form the parser reads back.
                            const linkTitle = meta.title ? markdownTitle(meta.title) : '';
                            text = `[${text}](${sanitizeMarkdownUrl(link)}${linkTitle})`;
                        }
                    }
                    if (meta?.abbreviationTitle) {
                        // Markdown Extra's abbreviation syntax has no inline marker - the bare
                        // word round-trips as-is, with its expansion collected at the document
                        // end via `*[abbr]: title`.
                        this.collectedAbbreviations.set(node.text || '', meta.abbreviationTitle);
                    }
                    if (meta?.citationKey) {
                        // Allowlist to exactly the character class MarkdownParser's own citation
                        // recognizer accepts, so this is provably lossless for anything it
                        // produced - while fully neutralizing a key arriving from HtmlParser's
                        // `data-citation-key`, which accepts any string. Like the wikilink above,
                        // this branch also replaces `text` wholesale, so a strip that left `<`
                        // behind discarded the escaping applied earlier.
                        const key = String(meta.citationKey).replace(/[^a-zA-Z0-9_:.-]/g, '');
                        text = this.resolvedDialect.citations !== 'none' ? `[@${key}]` : `[${key}]`;
                    }
                    return text;
                }

                case 'heading': {
                    const meta = node.metadata as HeadingMetadata;
                    const level = Math.min(Math.max(meta?.level || 1, 1), 6);
                    let id = '';
                    let remainingAnchors: string[] = [];

                    // (An id that slugifies to nothing is not written: `{#}` is no id, and read as text.)
                    // Anchors waiting from an empty paragraph come first (see pendingAnchorIds), so the
                    // heading's own id, or the generated one, stays its id.
                    const waiting = this.resolvedFallbackToHtml.anchors && !this.config.ignoreInternalLinks ? this.pendingAnchorIds.splice(0) : [];
                    const generated = this.config.generateIds && !meta?.anchorIds?.length ? this.slugify(this.getNodeText(node)) : '';
                    if (!this.config.ignoreInternalLinks && (meta?.anchorIds?.length || waiting.length)) {
                        // Slugified, as a Markdown identifier must be; the last is the heading's own.
                        const ids = this.anchorSlugs([...waiting, ...(meta?.anchorIds ?? []), ...(generated ? [generated] : [])]);
                        const lastId = ids.pop();
                        if (lastId) id = ` {#${lastId}}`;
                        remainingAnchors = ids;
                    } else if (generated) {
                        id = ` {#${generated}}`;
                    }
                    // An empty heading is its hashes alone (with its id, if any): `#` ends the line. A
                    // heading is one line, so its line breaks are joined as a list item's are (a line
                    // break ended it, and the rest, its id with it, was read back as a paragraph).
                    const headingText = joinLines(trimBlockEdges(childrenOutput), this.resolvedFallbackToHtml.itemLineBreaks ? '<br>' : ' ');
                    const prefix = '#'.repeat(level) + (headingText || id ? ' ' : '');

                    const anchors = this.resolvedFallbackToHtml.anchors
                        ? remainingAnchors.map(aid => `<a name="${aid}"></a>`).join('')
                        : '';
                    // In a list item, cell or definition (one line of Markdown) a heading cannot be one:
                    // its text is written, with its ids, rather than a `#` read back as text.
                    if (this.inOneLine) {
                        this.deferAnchors({ anchorIds: [...remainingAnchors, ...(id ? [id.slice(3, -1)] : [])] });
                        return `${trimBlockEdges(childrenOutput)}\n\n`;
                    }
                    let content = `${prefix}${headingText}${headingText ? id : id.trimStart()}`;

                    // Alignment fallback via HTML div/p
                    if (this.resolvedFallbackToHtml.alignment && meta?.alignment && meta.alignment !== 'left') {
                        // Use extra newlines to ensure Markdown inside the div is parsed
                        content = `<div style="text-align: ${sanitizeCssValue(meta.alignment)}">\n\n${content}\n\n</div>`;
                    }

                    return `${anchors}${anchors ? '\n' : ''}${content}\n\n`;
                }

                case 'paragraph': {
                    const meta = node.metadata as any;
                    let content = trimBlockEdges(childrenOutput);
                    // An empty paragraph (a bookmark on an empty line) keeps its anchors for the block
                    // that follows (see pendingAnchorIds).
                    if (!content) {
                        if (meta?.anchorIds?.length && this.resolvedFallbackToHtml.anchors && !this.config.ignoreInternalLinks) appendAll(this.pendingAnchorIds, meta.anchorIds);
                        return '';
                    }
                    const anchors = this.inOneLine ? this.deferAnchors(meta) : this.renderAnchors(meta);

                    // Alignment fallback via HTML div/p
                    if (this.resolvedFallbackToHtml.alignment && meta?.alignment && meta.alignment !== 'left') {
                        content = `<div style="text-align: ${sanitizeCssValue(meta.alignment)}">${content}</div>`;
                    }

                    return `${anchors}${content}\n\n`;
                }

                case 'list': {
                    const meta = node.metadata as ListMetadata;
                    const indentSpaces = ' '.repeat(4);
                    // Clamp the nesting depth: `indentation` derives from an uncapped document `ilvl`,
                    // so a hostile value would otherwise repeat the indent into a multi-GB string.
                    const indent = indentSpaces.repeat(this.listDepths.get(node) ?? clampRepeat(meta?.indentation || 0, 64));
                    const bullet = `${this.resolvedDialect.bulletListMarker} `;
                    const marker = meta?.isTask
                        ? (meta.checked ? `${bullet}[x] ` : `${bullet}[ ] `)
                        : (meta?.listType === 'ordered' ? `${clampInt(meta.itemIndex, 0, 999_999_998, 0) + 1}${this.resolvedDialect.orderedListMarker} ` : bullet);
                    const anchors = this.renderAnchors(meta);
                    // A list item is a single Markdown line. HTML-origin items carry `paragraph`
                    // children (e.g. `<li><p>a</p><ul>...`), whose renderer appends `\n\n`; dumped
                    // verbatim that produces `- a\n\n\n    - a1`, whose blank line splits the list
                    // apart and whose indent is then stripped on reparse, flattening the nesting.
                    // Collapse the item's internal breaks the same way table cells do (see the
                    // `cellLineBreaks` handling in renderMarkdownTable): join with `<br>` when the
                    // fallback is on, a space when off. Block children (code fences, tables) inside
                    // an item degrade under this join, exactly as they do inside a cell.
                    const br = this.resolvedFallbackToHtml.itemLineBreaks ? '<br>' : ' ';
                    // Only the whitespace a reader strips: a no-break space starting the item is text.
                    const content = joinLines(trimAsciiWhitespace(childrenOutput), br);
                    // An empty item is its marker alone, with no trailing space. In a definition (one line)
                    // an item cannot be one, and its marker is text there, escaped as text is.
                    if (this.inDefinition > 0) return `${markdownEscapeInline(marker, true)}${anchors}${content}\n`;
                    return `${indent}${anchors || content ? marker : marker.trimEnd()}${anchors}${content}\n`;
                }

                case 'image': {
                    const mode = this.imageMode();
                    if (mode === 'none') return '';
                    const meta = node.metadata as ImageMetadata;
                    const anchors = this.renderAnchors(meta, false);
                    // On the image's line, as a paragraph's are: a line break after them is a line break
                    // in the paragraph holding the image. On a line of their own only before a fenced
                    // block of recognized text, which must start a line.
                    const withAnchors = (markdown: string) => (anchors && markdown.startsWith('`') ? `${anchors}\n` : anchors) + markdown;
                    const ocr = (node.text || '').trim();

                    // OCR text from a scanned page carries meaning in its whitespace and line breaks
                    // (columns, indentation, aligned rows). Regular Markdown collapses runs of spaces and
                    // joins single line breaks, which destroys that layout, so render multi-line OCR
                    // verbatim inside a fenced block (one backtick longer than any embedded run so an
                    // inner ``` can't close it early) and keep single-line OCR as ordinary prose.
                    let ocrMd = '';
                    if (ocr) {
                        if (/[\r\n]/.test(ocr)) {
                            const longestRun = Math.max(0, ...(ocr.match(/`+/g) || []).map(s => s.length));
                            const fence = '`'.repeat(Math.max(3, longestRun + 1));
                            ocrMd = `${fence}\n${ocr}\n${fence}`;
                        } else {
                            ocrMd = markdownEscapeInline(ocr, true);
                        }
                    }

                    // ocr-text-only: emit just the recognized text, no image markup.
                    if (mode === 'ocr-text-only') return ocr ? withAnchors(ocrMd) : '';

                    // Build the image markup: inline as a data URI when small, otherwise reference by
                    // name (never inline a multi-MB image, which would emit a single line that overflows
                    // downstream Markdown parsers).
                    let src = meta?.url || meta?.attachmentName || '';
                    if (!meta?.url && meta?.attachmentName && this.ast) {
                        const attachment = this.getAttachment(meta.attachmentName);
                        if (attachment) {
                            const bytes = base64ByteLength(attachment.data);
                            if (bytes <= this.config.maxInlineImageBytes) {
                                if (this.inlineWithinBudget(bytes, meta.attachmentName)) src = `data:${attachment.mimeType || 'image/png'};base64,${attachment.data}`;
                            } else {
                                // Over the cap, so the picture itself cannot travel in the output
                                // (inlining a multi-MB image would overflow downstream Markdown
                                // parsers). Surface it, then keep the image's recognized text when it
                                // has any - that is the readable content of a scanned page, and it is
                                // what maxInlineImageBytes documents - falling back to the compact
                                // name reference when there is none.
                                // ('image+ocr-text' already emits both, so only 'image-only' changes.)
                                this.warn(OfficeWarningType.IMAGE_NOT_INLINED, { name: meta.attachmentName, bytes, limit: this.config.maxInlineImageBytes });
                                if (ocr && mode === 'image-only') return withAnchors(ocrMd);
                            }
                        }
                    }
                    // Alt text on one line, escaped as text (a `]` would close the `![...]`, and a
                    // renderer reads markup in alt text), and the URL scheme neutralized.
                    const safeAlt = markdownEscapeInline(foldLines(meta?.altText || 'image'));
                    const safeSrc = sanitizeMarkdownUrl(src, { allowDataImage: true });
                    const imgTitle = meta?.title ? markdownTitle(meta.title) : '';
                    const imageMd = withAnchors(this.linkedImage(meta, `![${safeAlt}](${safeSrc}${imgTitle})${this.renderAttributeList(meta)}`));

                    // image+ocr-text: the image, then its recognized text.
                    if (mode === 'image+ocr-text' && ocr) return `${imageMd}\n\n${ocrMd}`;
                    return imageMd;
                }

                case 'table': {
                    const anchors = this.renderAnchors(node.metadata);
                    const tableOutput = await this.renderMarkdownTable(node, processor);
                    // The HTML-fallback path (merged cells/nested tables, or a dialect that forces
                    // HTML tables outright) already carries data-align on the <table> tag directly -
                    // only the plain pipe-table form needs the attribute-list syntax for alignment.
                    const usedHtmlFallback = this.resolvedDialect.tables === 'html' ||
                        (this.resolvedFallbackToHtml.tables && (this.hasNestedTable(node) || this.hasColspanOrRowspan(node)));
                    const attrList = usedHtmlFallback ? '' : this.renderAttributeList(node.metadata as TableMetadata, { skipAlign: true });
                    if (attrList) {
                        // Must glue directly below the last row with no blank line, or
                        // MarkdownParser's block splitter won't see it as part of the same block.
                        return `${anchors}${anchors ? '\n' : ''}${tableOutput.replace(/(?<!\n)\n+$/, '\n')}${attrList}\n`;
                    }
                    return `${anchors}${anchors ? '\n' : ''}${tableOutput}`;
                }

                case 'row':
                    // Handled in the 'table' case above
                    return childrenOutput;
                case 'cell':
                    // Its ids start its line, where the parser reads them back as its own.
                    return `${this.renderAnchors(node.metadata)}${childrenOutput}`;

                case 'break': {
                    // A hard line break (CommonMark: two trailing spaces before the newline)
                    // round-trips back to a distinct 'break' node on reparse. A thematic break
                    // emits `---` as its own block (the top-level loop supplies the surrounding
                    // blank lines), so a Markdown `---` / HTML `<hr>` survives a save instead of
                    // collapsing to whitespace. Every other breakType - notably 'page', which
                    // Markdown has no syntax for - keeps emitting a bare newline, unchanged.
                    const meta = node.metadata as BreakMetadata | undefined;
                    // Where the break starts a line (after another break), two trailing spaces would
                    // make a whitespace-only line, which ends the paragraph: a backslash does not.
                    // Among blocks (no line of text around it) a line break is nothing Markdown can show.
                    if (isLineBreakType(meta?.breakType)) return this.lineDepth === 0 ? '' : this.atLineStart ? '\\\n' : '  \n';
                    // A rule is a block, blank lines around it: `a---b` in a paragraph was text. In a list
                    // item, cell or definition (one line of Markdown) it cannot be one, and is a line break.
                    if (meta?.breakType === 'thematic') return this.inOneLine ? `${this.deferAnchors(meta)}\n` : `${this.anchorsBefore(meta) || '\n\n'}---\n\n`;
                    return '\n';
                }

                case 'code': {
                    const meta = node.metadata as CodeMetadata;
                    // Math content reached the output completely raw, which mattered most under
                    // `math: 'none'` (the commonmark preset), where there is no `$` wrapper at
                    // all and the text lands directly in the document body.
                    //
                    // Encode rather than drop: `$a < b$` is ordinary LaTeX, and dropping `<`
                    // would silently corrupt real formulae. markdownEscapeTags only touches `<`
                    // followed by a letter/`/`/`!`/`?`, which is not idiomatic math, and it is
                    // idempotent - so output is stable across repeated round-trips even though
                    // the first cycle shifts an anomalous `<img` to `&lt;img`. Not the text
                    // escaping: the parser decodes no entities in math, so an `&` stays as written
                    // (`a &= b` is an alignment). (Fully lossless would mean teaching
                    // MarkdownParser.decodeHtmlEntities to cover math `code` nodes; that is a
                    // parser behaviour change with its own baseline consequences.)
                    // A list item and a pipe-table cell are one line of Markdown, where no block can go:
                    // block math is written as inline math, its line ends as spaces (whitespace to TeX).
                    const inLine = this.inOneLine;
                    // A block's anchors stand on a line before it; in one line, or on inline math, they
                    // are the line's container's (see deferAnchors).
                    const anchors = inLine || meta?.math === 'inline' ? this.deferAnchors(meta) : '';
                    const anchorsBefore = inLine || meta?.math === 'inline' ? '' : this.anchorsBefore(meta);
                    if (meta?.math === 'block' && inLine) {
                        const mathInline = markdownEscapeTags(node.text || '').replace(/[$]+/g, '').replace(/(?<!\s)\s*[\r\n]+\s*/g, ' ').trim();
                        return anchors + (this.resolvedDialect.math === 'dollar' ? `$${mathInline}$` : mathInline);
                    }
                    if (meta?.math === 'block') {
                        // A content line of exactly `$$` would close the block early, so it is
                        // indented by one space, as is one already indented (the parser removes one
                        // space from each, so every such line reads back as written).
                        const mathBlock = markdownEscapeTags(node.text || '')
                            .split('\n').map(l => (/^ *\$\$$/.test(l) ? ` ${l}` : l)).join('\n');
                        return this.resolvedDialect.math === 'dollar' ? `${anchorsBefore || '\n'}$$\n${mathBlock}\n$$\n\n` : `${anchorsBefore || '\n'}${mathBlock}\n\n`;
                    }
                    if (meta?.math === 'inline') {
                        // Dropping `$` and newlines is lossless here: the parser's own inline-math
                        // recognizer is `\$(?!\s)([^$\n]+?)(?<!\s)\$`, which can never capture either.
                        const mathInline = markdownEscapeTags(node.text || '').replace(/[$\r\n]+/g, '');
                        return anchors + (this.resolvedDialect.math === 'dollar' ? `$${mathInline}$` : mathInline);
                    }
                    // No language name holds `<` or `>`: without them, an info string cannot open a tag
                    // in a renderer that writes it into its HTML unescaped.
                    const lang = (meta?.language || '').replace(/[\r\n`<>]+/g, '');
                    // A `code` node is always block-level: genuinely inline code is a monospace
                    // text node, never a `code` node. So it is a fenced block, whatever its length
                    // and whether or not it names a language: a one-line block (a shell command, a
                    // one-line `mermaid` diagram) written as an inline span came back as inline
                    // code, losing its block-ness and any language. Only a pipe-table cell, where a
                    // fence cannot go, keeps a one-line block without a language as an inline span.
                    // (Testing `[\r\n]`, not just `\n`, still routes a CR-only body to the fenced
                    // branch there, where a renderer that normalizes `\r` would kill an inline span.)
                    if (!inLine) {
                        // Fence with one more backtick than the longest run inside the content
                        // so an embedded ``` can't close the block early and inject markup.
                        const longestRun = Math.max(0, ...((node.text || '').match(/`+/g) || []).map(s => s.length));
                        const fence = '`'.repeat(Math.max(3, longestRun + 1));
                        return `${anchorsBefore || '\n'}${fence}${lang}\n${node.text || ''}\n${fence}\n\n`;
                    }
                    // In a list item or pipe-table cell each line is a code span, and the item's or
                    // cell's line breaks (`<br>`) keep them apart: a fence there was joined into
                    // `` ```<br>a<br>``` ``, which read back as one span holding the tags.
                    return `${anchors}${(node.text || '').split(/\r?\n/).map(line => (line ? codeSpan(line) : '')).join('\n')} `;
                }

                case 'sheet': {
                    const anchors = this.renderAnchors(node.metadata);
                    const tableOutput = await this.renderMarkdownTable(node, processor);
                    return `\n---\n\n${anchors}${anchors ? '\n' : ''}${tableOutput}\n\n`;
                }

                case 'slide': {
                    const anchors = this.renderAnchors(node.metadata);
                    return `\n---\n\n${anchors}${anchors ? '\n' : ''}${childrenOutput}\n\n`;
                }
                case 'page': {
                    const anchors = this.renderAnchors(node.metadata);
                    return `\n---\n\n${anchors}${anchors ? '\n' : ''}${childrenOutput}\n\n`;
                }
                case 'note': {
                    const meta = node.metadata as NoteMetadata;
                    if (meta?.noteType === 'footnote' || meta?.noteType === 'endnote') {
                        if (this.resolvedDialect.footnotes === 'none') {
                            // Dialect has no footnote syntax - the caller inlines this bare body
                            // as a parenthetical at the reference point instead of collecting it
                            // into an end-of-document "### Notes" section under a [^id] marker.
                            return `${this.deferAnchors(meta)}${trimAsciiWhitespace(childrenOutput)}`;
                        }
                        // Indent continuation lines one level so a multi-line body re-parses as a
                        // single definition (a bare newline would end it). Single-line bodies, the
                        // common case, are unaffected. The note's ids start its body.
                        return `[^${this.getFootnoteKey(node)}]: ${this.renderAnchors(meta)}${trimAsciiWhitespace(childrenOutput).replace(/\n/g, '\n    ')}\n\n`;
                    }
                    return `> ${this.renderAnchors(meta)}**Note:** ${trimAsciiWhitespace(childrenOutput)}\n\n`;
                }

                case 'embed': {
                    // Markdown has no native embed syntax. `this.resolvedEmbeds` (from
                    // `dialect.embeds`, honoring the deprecated `fallbackToHtml.embeds` boolean)
                    // selects the form: 'html' (the single-line block this library has always
                    // emitted and re-recognises), 'directive' (a remark-directive leaf), 'link'
                    // (a plain link), 'thumbnail' (YouTube-only clickable preview).
                    const meta = node.metadata as EmbedMetadata;
                    // In a list item, cell or definition (one line of Markdown) an embed's block cannot
                    // go: it is a link there (its HTML block was read back as text).
                    const mode = this.inOneLine ? 'link' : this.resolvedEmbeds;
                    // A directive label sits inside `::name[...]`: on one line, escaped (its brackets
                    // too, which would end it), as the parser decodes it. A link label is link text,
                    // escaped as text. An attribute value sits inside `{...}`; percent-encode the
                    // space/brace chars that would break out (widths/aligns/ids never contain them,
                    // but a src can).
                    const dirLabel = markdownEscapePlain(foldLines(meta?.label || '').trim()).replace(/[[\]]/g, '\\$&');
                    const linkLabel = (fallback: string) => markdownEscapeInline(foldLines(meta?.label || fallback));
                    const dirUrl = (u: string) => sanitizeMarkdownUrl(u).replace(/[{}\s]/g, c => '%' + c.charCodeAt(0).toString(16).toUpperCase().padStart(2, '0'));
                    const attrList = (pairs: Array<[string, string | undefined]>) => {
                        const kv = pairs.filter(([, v]) => v !== undefined && v !== '').map(([k, v]) => `${k}=${v}`);
                        return kv.length ? `{${kv.join(' ')}}` : '';
                    };

                    // An embed naming no type (built by hand), or a YouTube one by its URL alone, is
                    // what it carries: its URL was lost in an empty YouTube block.
                    const embed = resolveEmbed(meta);
                    if (embed?.kind === 'iframe') {
                        const rawUrl = embed.url;
                        if (mode === 'directive') {
                            const src = dirUrl(rawUrl);
                            if (!src) return '';
                            const lbl = dirLabel ? `[${dirLabel}]` : '';
                            return `::embed${lbl}${attrList([['src', src], ['width', meta?.width], ['height', meta?.height], ['align', meta?.align]])}\n\n`;
                        }
                        if (mode === 'html') {
                            // sanitizeUrl scheme-checks and HTML-escapes the src (hostile schemes drop
                            // the node). The single-line <iframe> is what MarkdownParser recognises on
                            // reimport, gated there on preserveIframes.
                            const safe = sanitizeUrl(rawUrl);
                            if (!safe) return '';
                            const w = meta?.width ? ` width="${escapeHtml(meta.width)}"` : '';
                            const h = meta?.height ? ` height="${escapeHtml(meta.height)}"` : '';
                            return `\n<iframe src="${safe}"${w}${h}></iframe>\n\n`;
                        }
                        // 'link' and 'thumbnail' (thumbnail is YouTube-only, so a generic iframe
                        // degrades to a link) both emit a plain link.
                        const safe = sanitizeMarkdownUrl(rawUrl);
                        return safe ? `[${linkLabel('Embed')}](${safe})\n\n` : '';
                    }

                    const id = embed?.kind === 'youtube' ? embed.videoId : '';
                    if (mode === 'directive') {
                        const lbl = dirLabel ? `[${dirLabel}]` : '';
                        return `::youtube${lbl}${attrList([['id', id], ['width', meta?.width], ['align', meta?.align]])}\n\n`;
                    }
                    if (mode === 'html') {
                        const width = meta?.width ? ` data-width="${escapeHtml(meta.width)}"` : '';
                        const align = meta?.align ? ` data-align="${escapeHtml(meta.align)}"` : '';
                        const lbl = meta?.label ? ` data-embed-label="${escapeHtml(meta.label)}"` : '';
                        return `\n<div data-youtube-video="${escapeHtml(id)}"${width}${align}${lbl}></div>\n\n`;
                    }
                    if (mode === 'thumbnail' && id) {
                        const watch = sanitizeMarkdownUrl(`https://www.youtube.com/watch?v=${id}`);
                        const thumb = sanitizeMarkdownUrl(`https://img.youtube.com/vi/${id}/hqdefault.jpg`);
                        return `[![${linkLabel('YouTube')}](${thumb})](${watch})\n\n`;
                    }
                    // 'link' (and 'thumbnail' with no id): a plain link.
                    const url = meta?.url || (id ? `https://youtu.be/${id}` : '');
                    return url ? `[${linkLabel('YouTube')}](${sanitizeMarkdownUrl(url)})\n\n` : '';
                }

                case 'admonition': {
                    const meta = node.metadata as AdmonitionMetadata;
                    // `admonitionType` is a closed union in types.ts and both parsers already
                    // allowlist on import, so enforcing it here is a no-op for any conforming
                    // AST - it closes the gap for a programmatically-built one, where the type is
                    // interpolated straight into `:::TYPE` / `::: {.TYPE}` / `[!TYPE]`.
                    const rawType = String(meta?.admonitionType || 'note').toLowerCase();
                    const type = MD_ADMONITION_TYPES.has(rawType) ? rawType : 'note';
                    const label = type.toUpperCase();
                    // A newline in the title would close the `**...**` and, in the fenced-div
                    // branches, could emit a stray `:::` line. `title` is never parser-set, so
                    // there is no round-trip to preserve and escaping is free.
                    const title = meta?.title ? markdownEscapeInline(foldLines(meta.title)) : '';
                    const body = trimAsciiWhitespace(childrenOutput);
                    // In a list item, cell or definition (one line of Markdown) no quote can go: its
                    // label, bold, then its text (the marker was read back as text).
                    if (this.inOneLine) return `${this.deferAnchors(meta)}**${title || label.charAt(0) + label.slice(1).toLowerCase()}:** ${body}\n\n`;
                    const anchors = this.anchorsBefore(meta);

                    switch (this.resolvedDialect.admonitions) {
                        case 'fence':
                            // GLFM fenced-div: no dedicated title syntax, so a custom title (if
                            // any) is folded into the body as a bold first line.
                            return `${anchors}:::${type}\n${title ? `**${title}**\n\n` : ''}${body}\n:::\n\n`;
                        case 'fence-attribute':
                            // Pandoc's own fenced-div-with-class syntax; same title handling as fence.
                            return `${anchors}::: {.${type}}\n${title ? `**${title}**\n\n` : ''}${body}\n:::\n\n`;
                        case 'none': {
                            // Degrade to a plain bold-labeled blockquote, no special marker.
                            const quotedLines = body.split('\n').map(l => l.length > 0 ? `> ${l}` : '>').join('\n');
                            const heading = title || label.charAt(0) + label.slice(1).toLowerCase();
                            return `${anchors}> **${heading}:**\n${quotedLines}\n\n`;
                        }
                        case 'blockquote':
                        default: {
                            // Canonical GitHub blockquote form. No dedicated title syntax either
                            // (matches this library's historical output).
                            const quotedLines = body.split('\n').map(l => l.length > 0 ? `> ${l}` : '>').join('\n');
                            return `${anchors}> [!${label}]\n${quotedLines}\n\n`;
                        }
                    }
                }

                case 'definitionList':
                    // In a list item, cell or definition (one line of Markdown) it is its lines (see below).
                    if (this.inOneLine) return `${this.deferAnchors(node.metadata)}${childrenOutput}\n`;
                    return `${this.anchorsBefore(node.metadata)}${childrenOutput}\n`;

                case 'definitionTerm':
                case 'definitionDescription': {
                    // One line each, as the parser reads a term and its `: ` definitions (a term that
                    // holds a paragraph, from HTML's <dt><p>, was followed by a blank line, which ended
                    // the list): line breaks inside join as a list item's do.
                    const line = joinLines(trimAsciiWhitespace(childrenOutput), this.resolvedFallbackToHtml.itemLineBreaks ? '<br>' : ' ');
                    // In a list that is itself in one line (a list item, cell or definition), no `: ` line
                    // can start: the term is bold and the description plain, lines of that line.
                    if (this.definitionListsInLine[this.definitionListsInLine.length - 1]) {
                        return `${this.deferAnchors(node.metadata)}${node.type === 'definitionTerm' ? `**${line}**` : line}\n`;
                    }
                    // Its ids start its line, where the parser reads them back as its own.
                    const anchors = this.renderAnchors(node.metadata);
                    // A term's leading `:` is escaped: a term line starting with one (`:root`) is none.
                    if (node.type === 'definitionTerm') return this.resolvedDialect.definitionLists === 'none' ? `${anchors}**${line}**\n\n` : `${anchors}${line.startsWith(':') ? '\\' : ''}${line}\n`;
                    return this.resolvedDialect.definitionLists === 'none' ? `${anchors}${line}\n\n` : `: ${anchors}${line}\n`;
                }

                case 'comment':
                    // A source comment (`<!-- ... -->`) is re-emitted verbatim - inline it sits in its run,
                    // at top level the block loop below separates it like any other block. A review
                    // comment keeps its existing rendering (its children).
                    if (isSourceComment(node)) return `<!--${sanitizeCommentText(this.inlineCommentText(node))}-->`;
                    // A comment of text alone (a CSV comment line) is that text: it was dropped.
                    if (!node.children?.length && node.text) return `${markdownEscapeInline(node.text, this.atLineStart)}\n\n`;
                    return childrenOutput;

                case 'chart':
                case 'drawing':
                case 'header':
                case 'footer':
                case 'slideMaster':
                    return childrenOutput;
            }
        };

        const optimizedContent = this.optimizeNodes(this.ast.content);
        this.assignListDepths(optimizedContent);
        const body = new TextBuilder();
        body.append(output);
        for (let i = 0; i < optimizedContent.length; i++) {
            const node = optimizedContent[i];

            // A top-level footnote/endnote note is an orphan definition (unreferenced `[^id]: ...`
            // the MarkdownParser recovered). Collect it so it's emitted with the other definitions
            // at the document end rather than inline before them.
            if (isStandingNote(node)) {
                this.collectedNotes.push(node);
                continue;
            }
            let following = i + 1;
            while (following < optimizedContent.length && isStandingNote(optimizedContent[following])) following++;
            const nextNode = optimizedContent[following];

            this.atLineStart = body.isEmpty() || body.endsWith('\n');
            this.previousOutputChar = '';
            this.nextSibling = undefined;
            let result = await this.processNodeRecursive(node, processor);

            // Ensure lists and other block elements are separated from non-similar content by a blank line
            if (nextNode) {
                const isBothLists = node.type === 'list' && nextNode.type === 'list';
                if (!isBothLists) {
                    if (!result.endsWith('\n\n')) {
                        if (result.endsWith('\n')) result += '\n';
                        else result += '\n\n';
                    }
                }
            }

            appendBlock(body, node, result);
        }
        // Anchors no block took (an empty paragraph ending the document) are written at its end.
        if (this.pendingAnchorIds.length > 0) output = `${trimEndChars(body.toString(), '\n')}\n\n${this.renderAnchors({})}\n`;
        else output = body.toString();

        if (this.collectedNotes.length > 0) {
            // No decorative `---\n\n### Notes` preamble: `[^id]:` definitions are valid on their own
            // (GitHub/Pandoc render the footnotes section and its rule automatically), and the
            // literal heading round-tripped as a real `###` node - so every save/reload re-emitted
            // the parsed heading AND a fresh one, growing the document unbounded. Emitting the bare
            // definitions makes the cycle byte-stable. Behaviour change, noted in the changelog.
            // Collapse the preceding block's trailing blank lines so exactly one blank line separates
            // the body from the definitions (rather than the doubled `\n\n\n\n` the concatenation
            // would otherwise leave).
            output = trimEndChars(output, '\n');
            let notesMd = '\n\n';
            // De-duplicate by node identity: a footnote referenced more than once shares a single
            // note object (see MarkdownParser), pushed here once per reference. Emit its definition
            // just once. Distinct notes - even two office notes that happen to share a numeric id -
            // are separate objects and are all kept.
            // Writing a note collects the notes its own text refers to, which are written after it.
            const written = new Set<OfficeContentNode>();
            for (let i = 0; i < this.collectedNotes.length; i++) {
                const note = this.collectedNotes[i];
                if (written.has(note)) continue;
                written.add(note);
                notesMd += await this.processNodeRecursive(note, processor);
            }
            output += notesMd;
        }

        if (this.collectedAbbreviations.size > 0) {
            // One blank line before the definitions, as before the notes above.
            output = trimEndChars(output, '\n') + '\n\n';
            for (const [abbr, title] of this.collectedAbbreviations) {
                output += `*[${markdownEscapePlain(String(abbr).replace(/[[\]\r\n]+/g, ''))}]: ${markdownEscapePlain(foldLines(title))}\n`;
            }
        }

        // Only a run of literal "\n" at either end is ever a generator artifact here: block
        // separators, the notes/abbreviations sections, the unconditional '\n\n' before
        // hoistedContent (added even when hoistedContent is empty), and renderMarkdownTable's
        // HTML-fallback branches, which unconditionally wrap in a leading+trailing '\n' as
        // separators from whatever precedes/follows (in practice this rarely surfaces at the very
        // start of `output` today since frontmatter's own "---" almost always precedes real
        // content first - see the type doc on `ast.metadata` - but the strip is correct regardless
        // of what precedes it). Nothing else at either end is a generator artifact: not leading
        // whitespace, and not any other kind of trailing whitespace, both of which would be real
        // document content. See the identical reasoning in TextGenerator.generate().
        return {
            value: trimEndChars(trimStartChars(collapseBlankLines(output + '\n\n' + this.hoistedContent.join('\n\n')), '\n'), '\n'),
            messages: this.messages
        };
    }

    /**
     * Recursively processes nodes and builds output.
     * Overridden to provide AST optimization (merging adjacent text nodes).
     */
    protected override async processNodeRecursive(
        node: OfficeContentNode,
        processor: (node: OfficeContentNode, childrenOutput: string) => string | Promise<string>
    ): Promise<string> {
        // Mirrors the check in BaseGenerator.processNodeRecursive. This override replaces that
        // method entirely, so without repeating the check here the signal would be silently
        // inert for this generator - which is exactly how it was missed.
        checkAbortSignal(this.config.abortSignal);
        // Allow user to completely override rendering or skip via onNode
        const override = await this.handleOnNode(node);
        if (override === false) {
            return '';
        }
        if (typeof override === 'string') {
            return override;
        }

        const walkedByProcessor = node.type === 'table' || node.type === 'sheet';
        const wasInImplicitBold = this.inImplicitBold;
        if (node.type === 'heading' && this.hasUniformFormatting(node, f => f?.bold === true)) this.inImplicitBold = true;
        let childrenOutput = '';
        if (!walkedByProcessor && node.children && node.children.length > 0) {
            // Optimization: Merge adjacent text nodes with identical formatting
            const optimizedChildren = this.optimizeNodes(node.type === 'paragraph' || node.type === 'heading' ? withoutBreaksAtBlocks(node.children) : node.children);
            this.assignListDepths(optimizedChildren);
            // What a paragraph, list item or definition holds starts a line, where a block marker at
            // its start must be escaped (`- # x` is an item holding a heading); a heading or cell holds
            // inline content only.
            const startsLines = (node.type === 'paragraph' || node.type === 'list' || node.type === 'definitionTerm' || node.type === 'definitionDescription') && !this.inPipeTableCell;
            // So does any other node holding text directly (an admonition's or a note's inline content).
            const holdsLine = LINE_HOLDERS.has(node.type) || optimizedChildren.some(child => child.type === 'text');
            if (holdsLine) this.lineDepth++;
            if (node.type === 'definitionList') this.definitionListsInLine.push(this.inOneLine);
            if (node.type === 'list') this.inListItem++;
            const definition = node.type === 'definitionTerm' || node.type === 'definitionDescription';
            if (definition) this.inDefinition++;
            const children = new TextBuilder();
            for (let i = 0; i < optimizedChildren.length; i++) {
                const child = optimizedChildren[i];
                // A footnote or endnote standing in a container (a page's endnote no text refers to) is
                // an unreferenced definition: written with the others at the end, as at the top level.
                if (isStandingNote(child)) {
                    this.collectedNotes.push(child);
                    continue;
                }
                // A paragraph's lines stay lines; every other container puts its content after a
                // marker on one line (a heading, list item or cell), so nothing in it starts one. A
                // pipe-table cell holds inline content only, so nothing in it begins a block either.
                this.atLineStart = startsLines && (children.isEmpty() || children.endsWith('\n'));
                this.previousOutputChar = children.lastChar();
                let following = i + 1;
                while (following < optimizedChildren.length && isStandingNote(optimizedChildren[following])) following++;
                const next = this.nextSibling = optimizedChildren[following];
                let childOutput = await this.processNodeRecursive(child, processor);
                // Blocks in a container (a page, a slide, a cell) stand apart as top-level blocks do:
                // one blank line where a block ends or begins, except between a list's items. A heading
                // after a picture would otherwise share its line and read back as text, and a paragraph
                // on the line after a list item as part of the item. Runs of inline content stay joined.
                if (next && childOutput && (isBlockNode(child) || isBlockNode(next)) && !(child.type === 'list' && next.type === 'list') && !childOutput.endsWith('\n\n')) {
                    childOutput += childOutput.endsWith('\n') ? '\n' : '\n\n';
                }
                appendBlock(children, child, childOutput);
            }
            childrenOutput = children.toString();
            if (holdsLine) this.lineDepth--;
            if (node.type === 'definitionList') this.definitionListsInLine.pop();
            if (node.type === 'list') this.inListItem--;
            if (definition) this.inDefinition--;
        }

        this.inImplicitBold = wasInImplicitBold;

        // When the dialect has no footnote syntax, a footnote/endnote is inlined right at its
        // reference point instead (see below) - so it must not also be collected into the
        // end-of-document "### Notes" section, or its content would be duplicated.
        const isInlinedFootnote = (note: OfficeContentNode): boolean => {
            const meta = note.metadata as NoteMetadata;
            return (meta?.noteType === 'footnote' || meta?.noteType === 'endnote') && this.resolvedDialect.footnotes === 'none';
        };

        this.collectNotesFrom(node);

        let result = await processor(node, childrenOutput);

        if (node.type === 'slide' && node.notes && node.notes.length > 0) {
            for (const note of node.notes) {
                result += await this.processNodeRecursive(note, processor);
            }
        } else if (node.notes && node.notes.length > 0) {
            for (const note of node.notes) {
                const meta = note.metadata as NoteMetadata;
                if (meta?.noteType !== 'footnote' && meta?.noteType !== 'endnote') continue;
                if (isInlinedFootnote(note)) {
                    // Markdown-specific degrade (not RTF/plain-text's "drop the marker, just
                    // append at the end" convention): inline the note's rendered body as a
                    // parenthetical right where it's referenced, since Markdown readers benefit
                    // from an inline association those simpler formats don't need in the same way.
                    const body = await this.processNodeRecursive(note, processor);
                    result += ` (Note: ${body})`;
                } else {
                    // Emit the [^id] reference marker at the point of reference. Without this,
                    // a footnote/endnote would only ever show up in the collected ### Notes
                    // section at the end, with no indication of where it was originally cited.
                    // In an HTML table cell the reference is HTML too (as HtmlGenerator writes it), which
                    // the parser ties to its `[^id]:` definition.
                    const key = this.getFootnoteKey(note);
                    result += this.inHtmlTable ? `<sup data-footnote-ref="${escapeHtml(key)}">${escapeHtml(key)}</sup>` : `[^${key}]`;
                }
            }
        }

        return result;
    }

    /**
     * Sets the written depth of each list item among `siblings`: at most one deeper than the item before
     * it, and 0 for the first of a run. Markdown nests an item only under an item, so an item starting a
     * list deeper, or nested two levels below its parent, was indented four columns more than an item
     * can be and read as a code block.
     */
    private assignListDepths(siblings: OfficeContentNode[]): void {
        let previous = -1;
        for (const node of siblings) {
            if (node.type === 'list') {
                const depth = Math.min(clampRepeat((node.metadata as ListMetadata | undefined)?.indentation || 0, 64), previous + 1);
                this.listDepths.set(node, depth);
                previous = depth;
            } else if (!isStandingNote(node)) {
                previous = -1;
            }
        }
    }

    /**
     * Merges adjacent text nodes with identical formatting and metadata.
     */
    private optimizeNodes(nodes: OfficeContentNode[]): OfficeContentNode[] {
        if (nodes.length <= 1) return nodes;

        const result: OfficeContentNode[] = [];
        let current: OfficeContentNode | null = null;

        for (const node of nodes) {
            if (node.type === 'text' && current && current.type === 'text' &&
                // A note anchors to the end of its text run and its `[^id]` marker is emitted there;
                // merging a following run onto a note-carrying run would slide the marker past it
                // (`Body[^1].` -> `Body.[^1]`). Keep such runs separate so the marker stays put and
                // matches where HtmlGenerator emits it. (A note-carrying run can join the run before
                // it: its note still ends the merged run.)
                (!current.notes || current.notes.length === 0) &&
                // Review comments stay on their own run (merging dropped the second run's).
                !current.comments?.length && !node.comments?.length &&
                this.areFormattingEqual(node.formatting, current.formatting) &&
                JSON.stringify(node.metadata) === JSON.stringify(current.metadata)) {
                current.text = (current.text || '') + (node.text || '');
                if (current.rawContent && node.rawContent) current.rawContent += node.rawContent;
                if (node.notes && node.notes.length > 0) {
                    if (!current.notes) current.notes = [];
                    appendAll(current.notes, node.notes);
                }
            } else if (node.type === 'text') {
                current = { ...node }; // Clone: the runs merged into it are appended to its text
                if (node.notes) {
                    current.notes = [...node.notes];
                }
                result.push(current);
            } else {
                // Any other node is kept itself: a note standing among blocks must stay the note
                // object text elsewhere refers to, or it is written twice under two keys.
                current = node;
                result.push(node);
            }
        }
        return result;
    }

    /**
     * Whether two runs' formatting is written the same way in Markdown, so the runs can be written as
     * one. A font, size or colour Markdown does not write (without the inline-HTML fallback) does not
     * keep apart runs that would otherwise come out side by side as two identical spans: a link's space
     * in the link colour was written `[ ](u)[text](u)`, and read back as one link.
     */
    private areFormattingEqual(f1: TextFormatting | undefined, f2: TextFormatting | undefined): boolean {
        return f1 === f2 || this.markdownFormattingKey(f1) === this.markdownFormattingKey(f2);
    }

    /** The part of a run's formatting the text writer writes (see the `text` case of generate). */
    private markdownFormattingKey(f: TextFormatting | undefined): string {
        if (!f || !this.config.includeFormatting) return '';
        const html = this.resolvedFallbackToHtml;
        return JSON.stringify([
            f.font === 'monospace', !!f.bold, !!f.italic, !!f.strikethrough, f.backgroundColor ?? null,
            html.textFormatting && [!!f.underline, !!f.subscript, !!f.superscript],
            html.inlineFormatting && [f.color ?? null, f.size ?? null],
        ]);
    }

    private async renderMarkdownTable(node: OfficeContentNode, processor: any): Promise<string> {
        if (!node.children || node.children.length === 0) return '';

        // A dialect that has no native table syntax at all (e.g. strict CommonMark) always
        // renders as HTML, regardless of complexity - this is a separate axis from the
        // nested/merged-cell HTML fallback below, which only applies to otherwise-native tables.
        if (this.resolvedDialect.tables === 'html') {
            return '\n' + await this.renderTableAsHtml(node) + '\n';
        }

        // If table is complex, nested, or uses merges, fallback to HTML for high fidelity if allowed
        const isComplex = this.hasNestedTable(node) || this.hasColspanOrRowspan(node);
        if (this.resolvedFallbackToHtml.tables && isComplex) {
            return '\n' + await this.renderTableAsHtml(node) + '\n';
        }

        // Handle nested tables in pure Markdown by hoisting them out
        if (this.isInsideTable && !this.resolvedFallbackToHtml.tables) {
            const wasInside = this.isInsideTable;
            this.isInsideTable = false; // Reset to allow rendering the hoisted table correctly
            const hoistedId = this.hoistedContent.length + 1;
            const tableOutput = await this.renderMarkdownTableInternal(node, processor);
            this.hoistedContent.push(`**Table ${hoistedId} (Hoisted from cell content):**\n${tableOutput}`);
            this.isInsideTable = wasInside;
            return `*(See Table ${hoistedId} below)*`;
        }

        this.isInsideTable = true;
        const result = await this.renderMarkdownTableInternal(node, processor);
        this.isInsideTable = false;
        return result;
    }

    private collectNotesFrom(node: OfficeContentNode) {
        if (!node.notes || node.notes.length === 0) return;
        if (node.type === 'slide') return;
        const isInlinedFootnote = (note: OfficeContentNode): boolean => {
            const meta = note.metadata as NoteMetadata;
            return (meta?.noteType === 'footnote' || meta?.noteType === 'endnote') && this.resolvedDialect.footnotes === 'none';
        };
        appendAll(this.collectedNotes, node.notes.filter(note => !isInlinedFootnote(note)));
    }

    private async renderMarkdownTableInternal(node: OfficeContentNode, processor: any): Promise<string> {
        let tableOutput = '';
        let maxCols = 0;

        // First pass: Process rows and determine max columns (accounting for colspans)
        const processedRows: string[][] = [];
        for (const rowNode of (node.children ?? [])) {
            // Content among the rows that is no row (a CSV comment line, a picture) is a row of one
            // cell where it stands: it was an empty row, its content gone.
            if (rowNode.type !== 'row') {
                const wasInPipeTableCell = this.inPipeTableCell;
                this.inPipeTableCell = true;
                // A cell's text starts no line of Markdown: a `#` there is text as it is.
                this.atLineStart = false;
                let content: string;
                try {
                    content = await this.processNodeRecursive(rowNode, processor);
                } finally {
                    this.inPipeTableCell = wasInPipeTableCell;
                }
                content = joinLines(trimAsciiWhitespace(content), this.resolvedFallbackToHtml.cellLineBreaks ? '<br>' : ' ').replace(/\|/g, '\\|');
                if (content) {
                    processedRows.push([content]);
                    maxCols = Math.max(maxCols, 1);
                }
                continue;
            }
            // The first row becomes the header row - the `| --- |` separator emitted below marks
            // it as such - so bold inside it is already implied.
            const wasInImplicitBold = this.inImplicitBold;
            if (processedRows.length === 0 && this.hasUniformFormatting(rowNode, f => f?.bold === true)) this.inImplicitBold = true;
            try {
                const override = await this.handleOnNode(rowNode);
                if (override === false) continue;
                if (typeof override === 'string') {
                    processedRows.push([override]);
                    continue;
                }
                // After the override checks, not before: a row the caller skipped via `onNode` must not
                // still contribute its footnote to the end-of-document Notes section, where it would
                // appear with no `[^id]` marker anywhere in the document pointing at it.
                // `renderTableAsHtml` gets this right by returning early, so collecting here keeps the
                // pipe and HTML paths agreeing on what a skipped row means.
                this.collectNotesFrom(rowNode);

                const rowCells: string[] = [];
                let lastCol = -1;

                if (rowNode.children) {
                    const cellNodes = rowNode.children.filter(c => c.type === 'cell');
                    for (const cellNode of cellNodes) {
                        const currentCol = (cellNode.metadata as any)?.col ?? (lastCol + 1);

                        // Fill gaps with empty cells
                        while (lastCol < currentCol - 1) {
                            rowCells.push('');
                            lastCol++;
                        }

                        // Process cell content
                        const wasInPipeTableCell = this.inPipeTableCell;
                        this.inPipeTableCell = true;
                        let cellContent: string;
                        try {
                            cellContent = await this.processNodeRecursive(cellNode, processor);
                        } finally {
                            this.inPipeTableCell = wasInPipeTableCell;
                        }
                        // Use <br> fallback only if allowed, otherwise space
                        const br = this.resolvedFallbackToHtml.cellLineBreaks ? '<br>' : ' ';
                        // Consume any trailing spaces before the newline(s) too, so a hard-break's
                        // `  \n` collapses to a single `<br>` instead of leaving `  <br>` in the cell.
                        cellContent = joinLines(trimAsciiWhitespace(cellContent), br).replace(/\|/g, '\\|');
                        rowCells.push(cellContent);

                        // Handle colspan by adding empty cells
                        const colSpan = (cellNode.metadata as any)?.colSpan || 1;
                        for (let i = 1; i < colSpan; i++) {
                            rowCells.push('');
                        }
                        lastCol = currentCol + colSpan - 1;
                    }
                }
                processedRows.push(rowCells);
                maxCols = Math.max(maxCols, rowCells.length);
            } finally {
                this.inImplicitBold = wasInImplicitBold;
            }
        }

        // Second pass: Build table string with separator. The separator carries standard GFM
        // per-column alignment (`:---`/`:---:`/`---:`) from columnAlignments, or the single table-level
        // align applied to every column, rather than a non-standard trailing `{align}` attribute list.
        const tableMeta = node.metadata as TableMetadata | undefined;
        // Column alignment lives on each cell (CellMetadata.align); read it off the header row.
        // Fall back to the single table-level align (an editor's data-align) for every column.
        const headerCells = (node.children?.[0]?.children || []).filter(c => c.type === 'cell');
        const alignMarker = (i: number): string => {
            const a = (headerCells[i]?.metadata as any)?.align ?? tableMeta?.align;
            return a === 'center' ? ':---:' : a === 'left' ? ':---' : a === 'right' ? '---:' : '---';
        };
        for (let i = 0; i < processedRows.length; i++) {
            const row = processedRows[i];
            // Pad row with empty cells if it has fewer than maxCols: the header row always (its cells and
            // the delimiter row's must match), the others within the document's budget (a short row
            // reads as ending in empty cells).
            const fill = i === 0 ? maxCols - row.length : this.padWithinBudget(maxCols - row.length);
            for (let k = 0; k < fill; k++) row.push('');

            tableOutput += `| ${row.join(' | ')} |\n`;

            if (i === 0) {
                // Header separator
                tableOutput += `| ${Array.from({ length: maxCols }, (_, i) => alignMarker(i)).join(' | ')} |\n`;
            }
        }

        return `\n${tableOutput}\n`;
    }

    private hasNestedTable(node: OfficeContentNode): boolean {
        if (!node.children) return false;
        for (const child of node.children) {
            if (child.type === 'table') return true;
            if (this.hasNestedTable(child)) return true;
        }
        return false;
    }

    private hasColspanOrRowspan(node: OfficeContentNode): boolean {
        if (!node.children) return false;
        for (const row of node.children) {
            if (row.type === 'row' && row.children) {
                for (const cell of row.children) {
                    if (cell.type === 'cell') {
                        const meta = cell.metadata as any;
                        if ((meta?.colSpan && meta.colSpan > 1) || (meta?.rowSpan && meta.rowSpan > 1)) {
                            return true;
                        }
                    }
                }
            }
        }
        return false;
    }

    /**
     * Renders a complex table as HTML since Markdown doesn't support nested tables or rowspans.
     */
    /**
     * One node of an HTML table cell's content, as HtmlGenerator writes it: text with its formatting,
     * link, citation or wikilink; a line break; code, and math in the markup the HTML parser reads back
     * as math; a picture (inlined as a `data:` URI up to `maxInlineImageBytes`, like a Markdown image)
     * with its link; a list item; a comment; a nested table.
     */
    private async htmlCellNode(n: OfficeContentNode, co: string): Promise<string> {
        // A Markdown HTML block ends at a blank line, so a line end in a cell's text, code, math or an
        // attribute is written as the character reference HTML reads back as the same character.
        const esc = (value: string) => escapeHtml(value).replace(/\r?\n/g, '&#10;');
        const link = (href: string, linkType: string | undefined, title: string | undefined, inner: string) => {
            const internal = linkType !== 'external';
            if (internal && this.config.ignoreInternalLinks) return inner;
            const target = internal && href.startsWith('#') && (this.config.generateIds || this.resolvedFallbackToHtml.anchors) ? `#${this.slugify(href.slice(1))}` : href;
            return `<a href="${sanitizeUrl(target)}"${title ? ` title="${esc(title)}"` : ''}>${inner}</a>`;
        };
        switch (n.type) {
            case 'text': {
                const meta = n.metadata as TextMetadata | undefined;
                if (meta?.abbreviationTitle) this.collectedAbbreviations.set(n.text || '', meta.abbreviationTitle);
                let text = esc(n.text || '');
                const f = n.formatting;
                if (this.config.includeFormatting && f) {
                    if (f.font === 'monospace') text = `<code>${text}</code>`;
                    if (f.bold && !this.inImplicitBold) text = `<b>${text}</b>`;
                    if (f.italic) text = `<i>${text}</i>`;
                    if (f.underline) text = `<u>${text}</u>`;
                    if (f.strikethrough) text = `<s>${text}</s>`;
                    if (f.subscript) text = `<sub>${text}</sub>`;
                    if (f.superscript) text = `<sup>${text}</sup>`;
                    if (f.backgroundColor === '#ffff00') text = `<mark>${text}</mark>`;
                    const styles = [['color', f.color], ['background-color', f.backgroundColor !== '#ffff00' ? f.backgroundColor : undefined], ['font-size', f.size]]
                        .map(([prop, value]) => [prop, value ? sanitizeCssValue(value) : ''])
                        .filter(([, value]) => value).map(([prop, value]) => `${prop}: ${value}`);
                    if (styles.length) text = `<span style="${esc(styles.join('; '))}">${text}</span>`;
                }
                if (meta?.citationKey) return `<cite data-citation-key="${esc(meta.citationKey)}">[@${esc(meta.citationKey)}]</cite>`;
                if (meta?.wikilink) return `<a href="#${esc(this.slugify(meta.link || ''))}" data-wikilink-page="${esc(meta.link || '')}">${text}</a>`;
                return meta?.link ? link(meta.link, meta.linkType, meta.title, text) : text;
            }
            case 'break':
                return (n.metadata as BreakMetadata | undefined)?.breakType === 'page' || (n.metadata as BreakMetadata | undefined)?.breakType === 'thematic' ? '<hr>' : '<br>';
            case 'code': {
                const meta = n.metadata as CodeMetadata | undefined;
                const code = n.text || '';
                const { attr, extra } = this.htmlIds(meta);
                if (meta?.math === 'inline') return `${extra}<span class="math math-inline" data-math="inline"${attr}>${esc(`$${code}$`)}</span>`;
                if (meta?.math === 'block') return `${extra}<div class="math math-block" data-math="block"${attr}>${esc(`$$${code}$$`)}</div>`;
                const lang = meta?.language ? ` class="language-${esc(meta.language)}"` : '';
                return `${extra}<pre${attr}><code${lang}>${esc(code)}</code></pre>`;
            }
            case 'image': {
                const mode = this.imageMode();
                const meta = n.metadata as ImageMetadata | undefined;
                const ocr = esc((n.text || '').trim());
                if (mode === 'none') return '';
                if (mode === 'ocr-text-only') return ocr;
                let src = meta?.url || meta?.attachmentName || '';
                const attachment = !meta?.url && meta?.attachmentName ? this.getAttachment(meta.attachmentName) : undefined;
                if (attachment) {
                    const bytes = base64ByteLength(attachment.data);
                    if (bytes <= this.config.maxInlineImageBytes) { if (this.inlineWithinBudget(bytes, meta!.attachmentName)) src = `data:${attachment.mimeType || 'image/png'};base64,${attachment.data}`; }
                    else this.warn(OfficeWarningType.IMAGE_NOT_INLINED, { name: meta!.attachmentName, bytes, limit: this.config.maxInlineImageBytes });
                }
                const { attr, extra } = this.htmlIds(meta);
                let img = `<img src="${sanitizeImageUrl(src)}" alt="${esc(meta?.altText || '')}"${meta?.title ? ` title="${esc(meta.title)}"` : ''}${attr}>`;
                if (meta?.link) img = link(meta.link, meta.linkType, meta.linkTitle, img);
                return extra + (mode === 'image+ocr-text' && ocr ? `${img}<br>${ocr}` : img);
            }
            case 'paragraph': {
                const { attr, extra } = this.htmlIds(n.metadata);
                return `<p${attr}>${extra}${co}</p>`;
            }
            case 'heading': {
                const level = Math.min(Math.max(Number((n.metadata as any)?.level) || 1, 1), 6);
                const { attr, extra } = this.htmlIds(n.metadata);
                return `<h${level}${attr}>${extra}${co}</h${level}>`;
            }
            case 'list': {
                const ordered = (n.metadata as ListMetadata | undefined)?.listType === 'ordered';
                const { attr, extra } = this.htmlIds(n.metadata);
                return ordered ? `<ol><li${attr}>${extra}${co}</li></ol>` : `<ul><li${attr}>${extra}${co}</li></ul>`;
            }
            case 'table': return await this.renderTableAsHtml(n);
            case 'embed': {
                const meta = n.metadata as EmbedMetadata | undefined;
                const url = meta?.url || (meta?.videoId ? `https://youtu.be/${meta.videoId}` : '');
                return url ? link(url, 'external', undefined, esc(meta?.label || url)) : co;
            }
            case 'comment':
                if (isSourceComment(n)) return `<!--${sanitizeCommentText(n.text || '')}-->`;
                return !n.children?.length && n.text ? esc(n.text) : co;
            case 'chart':
            case 'drawing':
            case 'slide':
            case 'note':
            case 'sheet':
            case 'row':
            case 'cell':
            case 'page':
            case 'header':
            case 'footer':
            case 'slideMaster':
            case 'admonition':
            case 'definitionList':
            case 'definitionTerm':
            case 'definitionDescription':
                return co;
        }
    }

    private async renderTableAsHtml(node: OfficeContentNode, override?: string | false | void): Promise<string> {
        if (override === false) return '';
        if (typeof override === 'string') {
            if (node.type === 'row') return `  <tr><td colspan="100%">${override}</td></tr>\n`;
            if (node.type === 'cell') return `<td>${override}</td>`;
            return override;
        }

        // A sheet is written as a table is (it was written as nothing: every sheet under a dialect
        // without pipe tables, and one with merged cells).
        if (node.type === 'table' || node.type === 'sheet') {
            let rows = '';
            this.inHtmlTable++;
            try {
                if (node.children) {
                    for (const row of node.children) {
                        // Content among the rows that is no row is a row of one cell, where it stands.
                        rows += row.type === 'row' || row.type === 'cell'
                            ? await this.renderTableAsHtml(row, await this.handleOnNode(row))
                            : `  <tr>\n${await this.renderTableAsHtml({ type: 'cell', children: [row] }, undefined)}  </tr>\n`;
                    }
                }
            } finally {
                this.inHtmlTable--;
            }
            // Carry table-layout alignment through the HTML fallback so it isn't lost
            // just because the table also needed HTML for merged cells.
            const tableMeta = node.metadata as any;
            const alignAttr = tableMeta?.align ? ` data-align="${escapeHtml(tableMeta.align)}"` : '';
            return `<table${alignAttr}>\n${rows}</table>\n`;
        } else if (node.type === 'row') {
            this.collectNotesFrom(node);
            let cells = '';
            if (node.children) {
                for (const cell of node.children) {
                    cells += await this.renderTableAsHtml(cell, await this.handleOnNode(cell));
                }
            }
            return `  <tr>\n${cells}  </tr>\n`;
        } else if (node.type === 'cell') {
            this.collectNotesFrom(node);
            const meta = node.metadata as any;
            const rs = meta?.rowSpan > 1 ? ` rowspan="${meta.rowSpan}"` : '';
            const cs = meta?.colSpan > 1 ? ` colspan="${meta.colSpan}"` : '';

            let content = '';
            if (node.children) {
                // Cell content as HTML, which is what the cell is (a Markdown renderer shows an HTML
                // block as it is, and the parser reads it with the HTML parser): the markup
                // HtmlGenerator writes for each inline construct, so nothing in a cell is dropped.
                for (const child of this.optimizeNodes(node.children)) {
                    content += await this.processNodeRecursive(child, (n, co) => this.htmlCellNode(n, co));
                }
            }
            const { attr, extra } = this.htmlIds(meta);
            return `    <td${rs}${cs}${attr}>${extra}${content}</td>\n`;
        }
        return '';
    }
}

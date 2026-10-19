/**
 * DOCX templating: fills a Word document's tags with caller data, preserving all of the document's
 * formatting and structure, and returns a new `.docx`. A tag is a field (`{{name}}`, `{{ a.b }}`) or a
 * block tag: `{{ if cond }}` ... `{{ else if cond }}` ... `{{ else }}` ... `{{ endif }}` keeps one
 * branch, `{{ for item in list }}` ... `{{ endfor }}` repeats its content once per item.
 *
 * A template is compiled once and rendered once per data entry:
 *
 * 1. Tags in text. Word freely splits a run of text across several `<w:r>`/`<w:t>` elements (a
 *    spell-check boundary, an rsid change), so a tag typed as `{{name}}` can arrive as `{{`, `nam`,
 *    `e}}`. Per paragraph we join its `<w:t>` texts, find the tags in the joined text, and put a marker
 *    where each tag began (and an empty one in every other `<w:t>` it spanned), so a field's value
 *    later adopts the formatting of the run its tag began in. Paragraphs nest (a text box's `<w:p>`
 *    lives inside a run of the outer `<w:p>`), so the text between paragraph-boundary tags is the unit:
 *    no tag spans a paragraph boundary.
 * 2. Blocks. In a part holding block tags, one forward scan of the XML records where every marker
 *    sits (its ancestors). Tags pair into blocks, and each block's scope follows from where its tags
 *    sit: in one paragraph, the inline text between them; in paragraphs side by side, the content
 *    between them (a paragraph holding nothing but tags is removed with them, so no blank line
 *    remains); in cells of one table, whole rows. A block is exactly the content between its tags, cut
 *    where they share a parent, so the output is always balanced XML; tags that cannot enclose content
 *    together are a typed error, never broken XML.
 * 3. Render. The compiled tree is walked with the data: fields filled, branches chosen, loop content
 *    copied with fresh ids where Word requires unique ones, and what repetition adds kept within
 *    `maxRepeatedContent`. A last pass keeps every table with a row and every cell ending in a paragraph,
 *    as Word requires, when a loop or a branch left one without.
 *
 * @module template/docxTemplate
 */

import { zipSync, Zippable } from 'fflate';
import { DecompressionLimits, OfficeErrorType, OfficeParserConfig, TemplateData, TemplateDataDetails, TemplateFormatter, TemplateSyntaxDetails, TemplateValue } from '../types.js';
import { escapeXml, stripInvalidXmlChars } from '../utils/sanitize.js';
import { extractFiles } from '../utils/zipUtils.js';
import { dateOf, formatDate, formatNumber, numberOf, Picture, PictureSettings, readPicture } from './templateFormat.js';
import { conditionSize, ConditionValue, evaluateCondition, Incomparable, kindOfValue, LOOP_DETAILS_NAME, MISSING, readTag, TemplateExpression, TemplatePath, TemplateTag } from './templateTags.js';

/** The `word/*.xml` parts that carry body text a tag could live in. */
const TEXT_PART = /^word\/(document\d*|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/;

/**
 * Parts whose bytes are already compressed, so deflating them again buys nothing and only costs CPU,
 * once per document in a batch. Everything else (the XML parts, the relationship files, an embedded
 * uncompressed metafile or font) still deflates: those are plain text or highly redundant binary and
 * compress several-fold, which matters more than the CPU on parts that are typically tens of KB.
 */
const STORED_PART = /\.(png|jpe?g|jfif|gif|webp|avif|heic|heif|emz|wmz|woff2?|mp[34]|m4[av]|aac|ogg|oga|ogv|opus|flac|mov|avi|wmv|webm|zip|gz|docx|xlsx|pptx|odt|ods|odp)$/i;

/** The element names the engine reads. */
const W = {
    p: 'w:p', r: 'w:r', t: 'w:t', pPr: 'w:pPr', rPr: 'w:rPr', tbl: 'w:tbl', tr: 'w:tr', tc: 'w:tc',
    sdt: 'w:sdt', sdtContent: 'w:sdtContent', customXml: 'w:customXml', sectPr: 'w:sectPr',
    bookmarkStart: 'w:bookmarkStart', bookmarkEnd: 'w:bookmarkEnd', txbxContent: 'w:txbxContent', fldChar: 'w:fldChar',
} as const;

/** The start tag a rewritten `<w:t>` gets, keeping its spaces. */
const TEXT_START_TAG = '<w:t xml:space="preserve">';

/** What a paragraph may hold and still count as holding nothing but tags (besides its properties). */
const TAG_ONLY_ALLOWED = new Set(['w:pPr', 'w:rPr', 'w:r', 'w:t', 'w:proofErr', 'w:lastRenderedPageBreak']);

/** The bookmark Word keeps at the last edit; dropping it with a tag-only paragraph loses nothing. */
const GO_BACK_BOOKMARK = '_GoBack';

/**
 * How many elements a cut may pass through (a run, its text, and wrappers such as hyperlinks): each is
 * closed and opened again at every tag, so an unbounded depth would cost tags times depth.
 */
const MAX_CUT_DEPTH = 16;

/** Elements a block may be cut through: closed before the cut and opened again after it. */
const SPLITTABLE = new Set(['w:t', 'w:r', 'w:hyperlink', 'w:smartTag', 'w:customXml', 'w:dir', 'w:bdo']);

/** Elements between a table and its rows that a row block may widen to. */
const ROW_WRAPPERS = new Set<string>([W.sdt, W.sdtContent, W.customXml]);

/** Containers Word requires to hold at least one paragraph (or table). */
const NEEDS_BLOCK = new Set(['w:tc', 'w:txbxContent', 'w:footnote', 'w:endnote', 'w:comment', 'w:hdr', 'w:ftr']);

/** The block-level children of a container. */
const BLOCK_ELEMENTS = new Set<string>([W.p, W.tbl, W.sdt]);

/** The character that ends an element's name in a tag. */
const NAME_END = /[\s/>]/;

/** How a tracked change in a block's path is named in an error, with what to do about it. */
const TRACKED_CHANGE = new Set(['w:ins', 'w:del', 'w:moveFrom', 'w:moveTo']);

/** Id attributes Word requires to be unique, and so renumbers in each repeated copy. */
type IdKind = 'docPr' | 'bookmarkStart' | 'bookmarkEnd' | 'bookmarkName' | 'sdt' | 'paraId' | 'textId'
    | 'footnoteRef' | 'endnoteRef' | 'commentStart' | 'commentEnd' | 'commentRef';

/** The longest bookmark name Word keeps. */
const MAX_BOOKMARK_NAME = 40;

/** Where a note a reference points at lives, so a repeated reference can point at a copy of it. */
const NOTE_PARTS = {
    footnoteRef: { part: 'word/footnotes.xml', element: 'w:footnote' },
    endnoteRef: { part: 'word/endnotes.xml', element: 'w:endnote' },
    comment: { part: 'word/comments.xml', element: 'w:comment' },
} as const;

/**
 * What evaluating a field, a loop or a branch costs in the repetition budget, in characters, besides
 * the length of its path (or its condition's size): about what such a tag writes, so a render over values
 * that write nothing, or branches that hold nothing, is bounded like one that writes. A path is read
 * by comparing its text with loop names and data keys, which costs as much as its length, at most once
 * per enclosing loop (64 at most).
 */
const NODE_COST = 4;

/** The longest string the runtime can make (V8's limit): a part that would grow past it is refused, typed. */
const MAX_PART_LENGTH = 2 ** 29 - 24;

/** How deeply blocks may nest: rendering walks one level of the tree per block. */
const MAX_BLOCK_DEPTH = 64;

/** The largest paragraph/text id Word accepts (they are below 0x80000000). */
const MAX_PARA_ID = 0x7FFFFFFF;

/**
 * Escapes only the characters Word actually escapes inside a `<w:t>` text node: `&`, `<`, `>`. Unlike
 * the general {@link escapeXml} (which also encodes `"`/`'` as `&quot;`/`&apos;`), this matches the
 * literal quotes Word leaves in text, so a custom delimiter such as `«"»` still matches the stored XML.
 */
function escapeXmlText(s: string): string {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Decodes the XML entities of `<w:t>` text, so a tag's condition reads `<` where Word stored `&lt;`. */
function decodeXmlText(s: string): string {
    return s.replace(/&(lt|gt|amp|quot|apos|#\d{1,7}|#x[0-9a-fA-F]{1,6});/g, (entity, name: string) => {
        if (name === 'lt') return '<';
        if (name === 'gt') return '>';
        if (name === 'amp') return '&';
        if (name === 'quot') return '"';
        if (name === 'apos') return '\'';
        const code = name[1] === 'x' ? parseInt(name.slice(2), 16) : parseInt(name.slice(1), 10);
        return code <= 0x10FFFF ? String.fromCodePoint(code) : entity;
    });
}

/** A scalar's text: dates as ISO, null and undefined as nothing. */
function scalarText(value: string | number | boolean | bigint | symbol | Date | null | undefined): string {
    if (value == null) return '';
    if (value instanceof Date) return isNaN(value.getTime()) ? '' : value.toISOString();
    return String(value);
}

/** What a text's characters grow by when written into a `<w:t>`: escapes, and a `<w:br/>` per line break. */
const RUN_BREAK_GROWTH = '</w:t><w:br/>'.length + TEXT_START_TAG.length;
function runXmlGrowth(text: string): number {
    let growth = 0;
    for (let i = 0; i < text.length; i++) {
        switch (text.charCodeAt(i)) {
            case 0x26: growth += 4; break;                     // &amp;
            case 0x3c: case 0x3e: growth += 3; break;          // &lt; &gt;
            case 0x22: case 0x27: growth += 5; break;          // &quot; &apos;
            case 0x0a: case 0x0d: case 0x2028: case 0x2029: growth += RUN_BREAK_GROWTH; break;
        }
    }
    return growth;
}

/** A text as the inner XML of a `<w:t>`, every line separator a `<w:br/>`. */
function textToRunXml(text: string): string {
    // Strip XML-illegal control chars (DB-sourced data often carries them) BEFORE escaping, exactly
    // as the DOCX generator does; then every line separator (CRLF, lone CR/LF, and the Unicode line/
    // paragraph separators U+2028/U+2029) becomes a real break: close this <w:t>, emit <w:br/>, reopen.
    return escapeXml(stripInvalidXmlChars(text)).replace(/\r\n|[\r\n\u2028\u2029]/g, '</w:t><w:br/>' + TEXT_START_TAG);
}

/** Whether a value has own properties a path can step into (an object or a list, not a date). */
function isContainer(value: unknown): value is Record<string, unknown> {
    return typeof value === 'object' && value !== null && !(value instanceof Date);
}

function hasOwn(target: object, key: string): boolean {
    return Object.prototype.hasOwnProperty.call(target, key);
}

// ─── Options and errors ─────────────────────────────────────────────────────

/** What the caller decides, and how errors are raised (as the caller's typed errors). */
export interface DocxTemplateOptions {
    start: string;
    end: string;
    onMissing: 'keep' | 'empty' | 'error';
    mtime: Date;
    onFieldMissing: (key: string) => never;
    /** What the document's loops and values may add by repetition (see render), and what to throw past it. */
    maxRepeatedContent: number;
    onRepeatLimit: () => never;
    onSyntax: (details: TemplateSyntaxDetails) => never;
    onDataInvalid: (details: TemplateDataDetails) => never;
    /** Blocks nest deeper than {@link MAX_BLOCK_DEPTH}. */
    onNestingLimit: () => never;
    /** A part would grow past what one string can hold (with data that large, not by repetition). */
    onOutputLimit: () => never;
    /** The template holds more elements than `decompressionLimits.maxXmlElements`, the parser's own bound. */
    maxXmlElements: number;
    onElementLimit: () => never;
    /** What a picture writes: separators, month and day names, the UTC offset. */
    pictures: PictureSettings;
    /** The caller's own filters, by name. */
    formatters: Record<string, TemplateFormatter>;
}

/** The built-in filters. */
const BUILT_IN_FILTERS = ['format', 'upper', 'lower'] as const;

/** A filter, read once when the template is compiled. */
type CompiledFilter =
    | { kind: 'format'; picture: Picture }
    | { kind: 'upper' | 'lower' }
    | { kind: 'custom'; name: string; run: TemplateFormatter; args: (string | number)[] };

// ─── Compiled template ──────────────────────────────────────────────────────

/** A tag found in the text: what it means, how it was typed, and its text in each `<w:t>` it spanned. */
interface TagRecord {
    tag: TemplateTag;
    /** Its raw text between the delimiters, from which an error message writes it as typed (see {@link typedOf}). */
    inner: string;
    /** Its raw length, for what a value outgrows it by. */
    length: number;
    /** A field's filters, compiled (none when it has none). */
    filters?: CompiledFilter[];
}

/** Literal XML, written as is (with its unique ids renumbered in a repeated copy). */
interface LiteralNode { kind: 'literal'; xml: string; id: number; slots: IdSlot[] | null }
/**
 * One `<w:t>` of a field: the value in its first, nothing (or the kept text) in the rest. `id` tracks
 * its first write (a kept field written again is a repeat), `cost` what evaluating it is charged.
 */
interface FieldNode { kind: 'field'; record: TagRecord; piece: number; pieceText: string; path: TemplatePath; text: string; id: number; cost: number }
/** Each branch's `cost` is what testing it is charged: its condition's size. */
interface IfNode { kind: 'if'; branches: { condition: TemplateExpression | null; record: TagRecord; cost: number }[]; sections: TemplateNode[][]; kept: KeptSpan }
interface ForNode { kind: 'for'; name: string; list: TemplatePath; listText: string; record: TagRecord; body: TemplateNode[]; kept: KeptSpan; cost: number }

/**
 * What a block writes when it is kept as typed: its stretch of the marked part, every marker written
 * back as the text it stands for. Made only when a block is kept, never stored: made up front, blocks
 * nested 64 deep would each hold a copy of everything inside them.
 */
interface KeptSpan { from: number; to: number; open: string; close: string; xml: string; ctx: PartContext; id: number; lastId: number; inLoop: boolean; firstTag: number; lastTag: number }
type TemplateNode = LiteralNode | FieldNode | IfNode | ForNode;

/** An id in a literal that a repeated copy renumbers. */
/** An id in a start tag: where its value starts and ends, and where its tag starts (`<`) and ends (`>`), as visitIds read them. */
interface IdSlot { start: number; end: number; kind: IdKind; value: string; tag: number; gt: number }

/**
 * A note compiled on its own, to be written per reference: its nodes (null with no tag), its text, and
 * the ranges across its own loops, written once in each copy of it.
 */
interface NoteTemplate { nodes: TemplateNode[] | null; xml: string; rangesAcross: Set<string> }

/** A part, compiled: its nodes, or null when it holds no tag (copied byte for byte). */
interface CompiledPart { name: string; nodes: TemplateNode[] | null; hasBlocks: boolean }

/** A whole template, compiled once and rendered once per data entry. */
export interface CompiledDocxTemplate {
    entries: Record<string, Uint8Array>;
    parts: Map<string, CompiledPart>;
    /** Where fresh ids start, when a loop repeats content holding ids (null when no loop does). */
    ids: IdSeeds | null;
    literalCount: number;
    /** Each note of a notes part compiled on its own, by part and id: rendered per reference in a loop. */
    noteTemplates: Map<string, Map<string, NoteTemplate>>;
    /**
     * A notes part with its notes that only loops refer to cut out, by part: such a note is written only
     * as its copies, each rendered in a pass's scope (rendered at the top level, its `{{ item }}` would
     * read nothing). The part is compiled from this text.
     */
    notesPartXml: Map<string, string>;
    /**
     * Bookmark and comment ranges (`bookmark:<id>`, `comment:<id>`) whose start and end lie on two sides of
     * a loop's edge: written once, under their own ids (see {@link rangesAcrossLoops}).
     */
    rangesAcrossLoops: Set<string>;
}

// ─── Phase 1: tags in text ──────────────────────────────────────────────────

/** What a template may hold, counted as it is read (see PartContext.elements). */
interface ElementBudget { left: number; textRuns: number }

/** Per-part compile state. */
interface PartContext {
    part: string;
    opts: DocxTemplateOptions;
    /** The delimiters as Word stores them in `<w:t>` text (`<<` as `&lt;&lt;`). */
    startText: string;
    endText: string;
    tags: TagRecord[];
    /**
     * Each marker's tag, which `<w:t>` of the tag it is (0 for the one where the tag began, which stands
     * for the tag), and the tag's raw text in that `<w:t>` (what a tag kept as typed writes back), by
     * marker number: three flat lists rather than an object per marker.
     */
    markerTag: number[];
    markerPiece: number[];
    pieceText: string[];
    /** The two characters a marker is written between, chosen to occur nowhere in the part. */
    open: string;
    close: string;
    hasBlocks: boolean;
    /**
     * What is left of the template's element budget, shared by its parts. The template's bound on elements (`left`: every element scanStructure reads, and each tag by its
     * weight), and, apart, on the text runs markChunk reads (`textRuns`), which it makes objects for
     * before scanStructure counts them.
     */
    elements: ElementBudget;
    /** Markers in comments, CDATA and declarations, which hold no content: written back as they were. */
    inert: { at: number; length: number; marker: number }[];
    /** Text that starts like a block tag but is not one: its first word, and the number of tags found before it. */
    unreadable: { inner: string; reason: string; beforeTag: number; word: string }[];
}

/**
 * Two characters that occur nowhere in `xml`, to write markers between: Unicode noncharacters first
 * (legal in XML, never typed), then the private use area.
 */
function chooseMarkerChars(xml: string): [string, string] | null {
    const ranges: [number, number][] = [[0xFDD0, 0xFDEF], [0xE000, 0xF8FF]];
    const present = new Set<number>();
    for (let i = 0; i < xml.length; i++) {
        const c = xml.charCodeAt(i);
        if ((c >= 0xFDD0 && c <= 0xFDEF) || (c >= 0xE000 && c <= 0xF8FF)) present.add(c);
    }
    const free: string[] = [];
    for (const [from, to] of ranges) {
        for (let c = from; c <= to && free.length < 2; c++) if (!present.has(c)) free.push(String.fromCharCode(c));
    }
    return free.length === 2 ? [free[0], free[1]] : null;
}

/**
 * The tags in a paragraph's joined text: each run from a start delimiter to the next end delimiter,
 * starting at the last start delimiter before that end (`{{ a {{b}}` holds the tag `{{b}}`), read as it
 * is found; what is returned is the tags, and the text that starts like a block tag but cannot be read.
 * One forward pass: a start delimiter with no end after it ends the search, so a run of unclosed `{{`
 * costs nothing.
 */
function findTagSpans(joined: string, ctx: PartContext): TagSpan[] {
    const spans: TagSpan[] = [];
    const { startText, endText } = ctx;
    for (let pos = 0; pos < joined.length;) {
        const s = joined.indexOf(startText, pos);
        if (s < 0) break;
        const e = joined.indexOf(endText, s + startText.length);
        if (e < 0) break;
        let start = joined.lastIndexOf(startText, e - startText.length);
        if (start < s) start = s;
        // Each tag becomes nodes, as an element does, and reading it makes an object for each of its
        // tokens (operands, filters, arguments, picture parts), so it is counted against the same bound,
        // by its weight, before any of them is made: one text run can hold millions of tags, or one tag
        // millions of tokens, and their objects exhausted the heap.
        const inner = joined.slice(start + startText.length, e);
        ctx.elements.left -= tagWeight(inner);
        if (ctx.elements.left < 0) ctx.opts.onElementLimit();
        const span: TagSpan = { start, end: e + endText.length, inner, tag: null, unreadable: null };
        span.tag = tagOfSpan(span);
        // A block keyword misspelt as another template language spells it is a mistake, named at once.
        if (span.misspelled) ctx.opts.onSyntax({ reason: 'unreadable', tag: typedTag(decodeXmlText(span.inner), ctx.opts), part: ctx.part, detail: span.unreadable ?? '' });
        // `name | word` where the word is no filter is prose, as text between the delimiters always was.
        if (span.tag?.kind === 'field' && span.tag.filters?.some(f => !isFilter(f.name, ctx))) span.tag = null;
        if (span.tag || span.unreadable) spans.push(span);
        // A tag consumes its end delimiter. Text that is not one does not: its end delimiter may begin the
        // next tag's start (`$5 for $name$` with `$` delimiters), so the search resumes inside it, as the
        // pattern this replaced retried. Still forward-only: each search overlaps the last by less than a
        // delimiter.
        pos = span.tag ? e + endText.length : Math.max(start + 1, e - startText.length + 1);
    }
    return spans;
}

/**
 * At least as many as the objects reading a tag's text can make: one for the tag, and one for each
 * character that is not white space, except a letter repeating the one before it (a long name is one
 * token, and a date picture's `dddd` one part, while a number picture's `0000` is four). Linear.
 */
function tagWeight(inner: string): number {
    let weight = 1;
    for (let i = 0; i < inner.length; i++) {
        const c = inner.charCodeAt(i);
        if (c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d || c === 0xa0) continue;
        const letter = (c >= 0x41 && c <= 0x5a) || (c >= 0x61 && c <= 0x7a);
        if (!letter || i === 0 || c !== inner.charCodeAt(i - 1)) weight++;
    }
    return weight;
}

/**
 * Reads a span's text as a tag. Text that starts like a block tag but cannot be read as one
 * (`{{ If you have questions, call us }}`) stays literal text, as prose between the delimiters always
 * has; it is remembered, so that when the part's blocks then fail to pair, the error names it and why.
 */
function tagOfSpan(span: TagSpan): TemplateTag | null {
    const reading = readTag(decodeXmlText(span.inner));
    if (reading === null) return null;
    if (!reading.ok) {
        span.unreadable = reading.reason;
        span.misspelled = reading.misspelled === true;
        return null;
    }
    return reading.tag;
}

/** A run of text between delimiters, and what it was read as. */
interface TagSpan { start: number; end: number; inner: string; tag: TemplateTag | null; unreadable: string | null; misspelled?: boolean }

/** Whether a name is a filter: a built-in one, or one of the caller's `formatters`. */
function isFilter(name: string, ctx: PartContext): boolean {
    return (BUILT_IN_FILTERS as readonly string[]).includes(name) || Object.prototype.hasOwnProperty.call(ctx.opts.formatters, name);
}

/**
 * A field's filters, read once: `format` takes one picture (read here, so a picture that cannot be read
 * is a typed error before anything is rendered), `upper` and `lower` take nothing, and any other name
 * must be one of the caller's `formatters`.
 */
function compileFilters(filters: { name: string; args: (string | number)[] }[], record: TagRecord, ctx: PartContext): CompiledFilter[] {
    const unreadable = (detail: string): never => ctx.opts.onSyntax({ reason: 'unreadable', tag: typedOf(record, ctx.opts), part: ctx.part, detail });
    return filters.map((filter): CompiledFilter => {
        const custom = Object.prototype.hasOwnProperty.call(ctx.opts.formatters, filter.name) ? ctx.opts.formatters[filter.name] : undefined;
        if (custom) return { kind: 'custom', name: filter.name, run: custom, args: filter.args };
        switch (filter.name) {
            case 'format': {
                const [picture, ...extra] = filter.args;
                if (typeof picture !== 'string' || extra.length) return unreadable('formats with something other than one quoted picture, as in format "#,##0.00" or format "MMMM d, yyyy"');
                const reading = readPicture(picture);
                if (!reading.ok) return unreadable(`formats with the picture "${picture}", which ${reading.reason}`);
                return { kind: 'format', picture: reading.picture };
            }
            case 'upper':
            case 'lower':
                if (filter.args.length) return unreadable(`writes something after ${filter.name}, which takes nothing`);
                return { kind: filter.name };
            default:
                // Not reached: a tag naming no filter is prose (see findTagSpans).
                return unreadable(`uses ${filter.name}, which is not a filter`);
        }
    });
}

/** How much of a long tag an error message shows. */
const MAX_TYPED_TAG = 120;

/** A tag as typed, for messages (long tags shortened). */
function typedTag(decodedInner: string, opts: DocxTemplateOptions): string {
    const inner = decodedInner.length > MAX_TYPED_TAG ? `${decodedInner.slice(0, MAX_TYPED_TAG)}...` : decodedInner;
    return `${opts.start}${inner}${opts.end}`;
}

/** A recorded tag as typed, for an error message: made only when an error is raised. */
function typedOf(record: TagRecord, opts: DocxTemplateOptions): string {
    return typedTag(decodeXmlText(record.inner.length > MAX_TYPED_TAG * 8 ? record.inner.slice(0, MAX_TYPED_TAG * 8) : record.inner), opts);
}

/**
 * Puts markers for the tags in one chunk of run-level XML (the content between two paragraph-boundary
 * tags). Returns the chunk unchanged when it holds no tag.
 */
function markChunk(chunk: string, ctx: PartContext): string {
    // A `<w:t>` is either self-closing (`<w:t/>`, emitted by the OpenXML SDK / POI / docx4j for an
    // empty run) or a normal `<w:t ...>text</w:t>`; capture the inner text of the latter.
    //
    // This is a forward-only scan rather than a `<w:t\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:t>)/g` exec loop:
    // the template input is arbitrary caller bytes, and that pattern rescans to the end of the part
    // from every `<w:t` token, so a run of unclosed tokens costs O(n^2) (80k tokens: ~6s). Here every
    // byte is looked at once. Well-formed XML, where `<w:t>` never nests, parses identically.
    // A chunk without the delimiter's first character holds no tag, however its text is split.
    if (!chunk.includes(ctx.startText[0])) return chunk;
    const segs: { start: number; end: number; inner: string }[] = [];
    const counted = (): void => { if (--ctx.elements.textRuns < 0) ctx.opts.onElementLimit(); };
    // One walk from `<` to `<`: a comment (or CDATA section, or processing instruction) is passed whole,
    // a `<w:t` or quote in it being none, and nothing is searched for past the next `<`.
    for (let s = chunk.indexOf('<'); s >= 0; s = chunk.indexOf('<', s + 1)) {
        const inert = inertEnd(chunk, s);
        if (inert !== null) { if (inert < 0) break; s = inert; continue; }
        // The `\b` after the tag name: `<w:tab` and `<w:tabs` are not `<w:t` elements.
        if (!chunk.startsWith('<w:t', s) || /\w/.test(chunk[s + 4] ?? '')) continue;
        const tagEnd = startTagEnd(chunk, s + 4);
        if (tagEnd < 0) break;                       // unterminated tag: nothing further can match
        counted();
        if (chunk.charCodeAt(tagEnd - 1) === 0x2f) { // `<w:t/>`
            segs.push({ start: s, end: tagEnd + 1, inner: '' });
            s = tagEnd;
            continue;
        }
        const close = chunk.indexOf('</w:t>', tagEnd + 1);
        if (close < 0) break;                        // no closing tag left in this chunk
        segs.push({ start: s, end: close + 6, inner: chunk.slice(tagEnd + 1, close) });
        s = close + 5;
    }
    if (!segs.length) return chunk;

    const joined = segs.map(s => s.inner).join('');
    if (!joined.includes(ctx.startText)) return chunk; // no tag here: leave it exactly as-is
    const matches = findTagSpans(joined, ctx);
    let tagCount = 0;
    for (const span of matches) if (span.tag) tagCount++;
    // Text that could not be read as a block tag is remembered in document order, among the tags.
    const remember = (span: TagSpan): void => {
        if (span.unreadable) ctx.unreadable.push({ inner: span.inner, reason: span.unreadable, beforeTag: ctx.tags.length, word: (/\p{L}+/u.exec(decodeXmlText(span.inner))?.[0] ?? '').toLowerCase() });
    };
    if (!tagCount) { matches.forEach(remember); return chunk; }

    // Rebuild each segment's inner text. Ordinary characters stay in their own segment; a tag becomes a
    // marker in each segment it spanned (the first one standing for the tag). Built from slices of the
    // joined text, cut where segments end: a character at a time (with a segment index per character),
    // one paragraph of 300 million characters ended the process out of memory.
    const offsets: number[] = [0];
    for (const seg of segs) offsets.push(offsets[offsets.length - 1] + seg.inner.length);
    const pieces: string[][] = segs.map(() => []);
    // The segment holding position `at` of the joined text; asked of positions that never decrease.
    let segment = 0;
    const segmentAt = (at: number): number => {
        while (segment < segs.length - 1 && offsets[segment + 1] <= at) segment++;
        return segment;
    };
    // The joined text from `from` to `to`, each part to the segment it came from.
    const copy = (from: number, to: number): void => {
        for (let at = from; at < to;) {
            const si = segmentAt(at);
            const end = Math.min(to, offsets[si + 1]);
            pieces[si].push(joined.slice(at, end));
            at = end;
        }
    };
    let copied = 0;
    for (const mt of matches) {
        if (!mt.tag) { remember(mt); continue; }
        copy(copied, mt.start);
        const record: TagRecord = { tag: mt.tag, inner: mt.inner, length: mt.end - mt.start };
        if (mt.tag.kind === 'field' && mt.tag.filters?.length) record.filters = compileFilters(mt.tag.filters, record, ctx);
        const tagIndex = ctx.tags.push(record) - 1;
        if (mt.tag.kind !== 'field') ctx.hasBlocks = true;
        // Each segment the tag overlaps gets one marker, holding that segment's part of the tag.
        for (let at = mt.start, piece = 0; at < mt.end; piece++) {
            const si = segmentAt(at);
            const end = Math.min(mt.end, offsets[si + 1]);
            const markerIndex = ctx.markerTag.push(tagIndex) - 1;
            ctx.markerPiece.push(piece);
            ctx.pieceText.push(joined.slice(at, end));
            pieces[si].push(`${ctx.open}${markerIndex}${ctx.close}`);
            at = end;
        }
        copied = mt.end;
    }
    copy(copied, joined.length);

    // Rebuild the chunk in a single forward pass: copy the text between `<w:t>` segments verbatim and
    // drop each rewritten segment in place. Splicing last-to-first would recopy the growing tail on
    // every step (quadratic; ~40s for a paragraph of 100k runs).
    const rebuilt: string[] = [];
    let cursor = 0;
    for (let si = 0; si < segs.length; si++) {
        rebuilt.push(chunk.slice(cursor, segs[si].start), TEXT_START_TAG, pieces[si].join(''), '</w:t>');
        cursor = segs[si].end;
    }
    rebuilt.push(chunk.slice(cursor));
    return rebuilt.join('');
}

/**
 * Marks every tag in a part. The part is read between paragraph-boundary tags (`<w:p ...>`, `</w:p>`,
 * `<w:p/>`): each stretch between two is one paragraph's own run content (nested paragraphs, e.g. text
 * boxes, are their own stretches), so joining a stretch's `<w:t>` never crosses a paragraph boundary.
 * Boundaries are found by a forward scan (a pattern reading a tag to its `>` read the rest of the part
 * from every `<w:p` that has none: 570 bytes of zipped `<w:p ` took 24 seconds), and only the stretches
 * a tag changes are collected, so a part of many paragraphs holds no list of them.
 */
function markPart(xml: string, ctx: PartContext): string {
    const out: string[] = [];
    let copied = 0;
    let stretchStart = 0;
    const mark = (end: number): void => {
        const stretch = xml.slice(stretchStart, end);
        const marked = markChunk(stretch, ctx);
        if (marked === stretch) return;
        out.push(xml.slice(copied, stretchStart), marked);
        copied = end;
    };
    for (let at = xml.indexOf('<', 0); at !== -1; at = xml.indexOf('<', at + 1)) {
        const inert = inertEnd(xml, at);
        if (inert !== null) { if (inert < 0) break; at = inert; continue; } // a comment's `<w:p` is none
        let tagEnd = -1;
        if (xml.startsWith('</w:p>', at)) tagEnd = at + 6;
        else if (xml.startsWith('<w:p', at) && !/[\w]/.test(xml.charAt(at + 4))) {
            const gt = startTagEnd(xml, at + 4);
            // No `>` left: no tag can end in the rest of the part.
            if (gt === -1) break;
            tagEnd = gt + 1;
        }
        if (tagEnd === -1) continue;
        mark(at);
        stretchStart = tagEnd;
        at = tagEnd - 1;
    }
    mark(xml.length);
    if (!out.length) return xml;
    out.push(xml.slice(copied));
    return out.join('');
}

// ─── Phase 2: where the markers sit ─────────────────────────────────────────

/** An element, as the structure scan saw it. Frames a marker sits in are kept (through `parent`). */
interface Frame {
    name: string;
    /** Offset of its `<`, just past its start tag, and just past its end (-1 until closed). */
    start: number;
    tagEnd: number;
    end: number;
    parent: Frame | null;
    depth: number;
    /** The paragraph it is in (itself, for a paragraph). */
    para: Frame | null;
    /** Inside paragraph or run properties. */
    props: boolean;
    /** A run's `<w:rPr>`, copied when the run is split. */
    rPr: [number, number] | null;
    /** A paragraph holds something besides block tags (text, a field, a picture, a bookmark...). */
    content: boolean;
    /** A paragraph's block tags (their first markers). */
    blockMarks: Occurrence[];
    goBack: Set<string> | null;
    tagOnly: boolean;
    /** Block-level children, and those that are tag-only paragraphs (the last of them kept). */
    kids: number;
    tagOnlyKids: number;
    lastTagOnly: Frame | null;
    /** Set on a `<w:rPr>` that is its run's properties. */
    isRunProps: boolean;
}

/** A marker in the XML: where it is, and the `<w:t>` holding it. */
interface Occurrence {
    marker: number; at: number; length: number; frame: Frame; block: Block | null; tagIndex: number;
    /** Which part of which complex field (a run of w:fldChar begin, separate, end) the tag is in; 0 outside any. */
    field: number;
}

function newFrame(name: string, start: number, tagEnd: number, parent: Frame | null): Frame {
    return {
        name, start, tagEnd, end: -1, parent, depth: parent ? parent.depth + 1 : 0,
        // A text box's content starts afresh: its paragraphs are not inside the paragraph holding the box.
        para: name === W.p || name === W.txbxContent ? null : parent?.para ?? null, props: false, rPr: null, content: false,
        blockMarks: [], goBack: null, tagOnly: false, kids: 0, tagOnlyKids: 0, lastTagOnly: null, isRunProps: false,
    };
}

/** An attribute's value in a start tag (bounded to the tag). */
function attributeOf(startTag: string, name: string): string | null {
    let nameEnd = 1;
    while (nameEnd < startTag.length && !NAME_END.test(startTag[nameEnd])) nameEnd++;
    const value = readStartTag(startTag, nameEnd, [name])?.attributes.get(name);
    return value ? startTag.slice(value[0], value[1]) : null;
}

/**
 * Where a comment, CDATA section, declaration or processing instruction starting at `at` ends (the index
 * of its last character), -1 for one never closed, or null when none starts there. Text in one that looks
 * like a tag is none.
 */
function inertEnd(xml: string, at: number): number | null {
    const c = xml.charCodeAt(at + 1);
    if (xml.charCodeAt(at) !== 0x3c || (c !== 0x21 && c !== 0x3f)) return null;
    // Looked for past the opener: `<!-->` opens a comment, it does not close one.
    const [close, from] = xml.startsWith('<!--', at) ? ['-->', at + 4] : xml.startsWith('<![CDATA[', at) ? [']]>', at + 9] : c === 0x3f ? ['?>', at + 2] : ['>', at + 2];
    const e = xml.indexOf(close, from);
    return e < 0 ? -1 : e + close.length - 1;
}

/** Where a start tag ends: its first `>` outside a quoted attribute value, looking from `from` (in the tag, outside a value), or -1. */
function startTagEnd(xml: string, from: number): number {
    for (let i = from; i < xml.length; i++) {
        const c = xml.charCodeAt(i);
        if (c === 0x22 || c === 0x27) {
            const q = xml.indexOf(c === 0x22 ? '"' : '\'', i + 1);
            if (q < 0) return -1;
            i = q;
        } else if (c === 0x3e) return i;
    }
    return -1;
}

/**
 * One forward scan of a marked part: every element's frame (its ancestors), every marker's place,
 * and what each paragraph holds. Raises `malformed-xml` for XML whose elements do not nest.
 */
function scanStructure(xml: string, ctx: PartContext): Occurrence[] {
    const occurrences: Occurrence[] = [];
    const fields: number[] = [];
    let fieldStates = 0;
    const root = newFrame('#document', 0, 0, null);
    let top = root;
    function malformed(detail: string): never {
        return ctx.opts.onSyntax({ reason: 'malformed-xml', tag: firstBlockTag(ctx), part: ctx.part, detail });
    }

    // The next marker at or after a position, asked of positions that never decrease: one forward
    // search for the whole part, not one per text node (which reads ahead to the next marker each time).
    let nextMarker = xml.indexOf(ctx.open);
    const markerFrom = (from: number): number => {
        if (nextMarker !== -1 && nextMarker < from) nextMarker = xml.indexOf(ctx.open, from);
        return nextMarker;
    };
    const onText = (from: number, to: number): void => {
        if (top.name !== W.t) return;
        const para = top.para;
        let at = from;
        while (at < to) {
            const m = markerFrom(at);
            const textEnd = m < 0 || m >= to ? to : m;
            if (para && !para.content && hasNonSpace(xml, at, textEnd)) para.content = true;
            if (textEnd >= to) break;
            const close = xml.indexOf(ctx.close, m + 1);
            const marker = Number(xml.slice(m + 1, close));
            const tagIndex = ctx.markerTag[marker];
            const tag = ctx.tags[tagIndex].tag;
            const occurrence: Occurrence = { marker, at: m, length: close + 1 - m, frame: top, block: null, tagIndex, field: fields[fields.length - 1] ?? 0 };
            occurrences.push(occurrence);
            if (para) {
                if (tag.kind === 'field') para.content = true;
                else if (ctx.markerPiece[marker] === 0) para.blockMarks.push(occurrence);
            }
            at = close + 1;
        }
    };
    const onElement = (frame: Frame, startTag: string): void => {
        const parent = frame.parent!;
        // A complex field is runs, not an element: its begin, separate and end marks are followed in
        // document order, each part of each field a state of its own.
        if (frame.name === W.fldChar) {
            const type = attributeOf(startTag, 'w:fldCharType');
            if (type === 'begin') fields.push(++fieldStates);
            else if (type === 'separate' && fields.length) fields[fields.length - 1] = ++fieldStates;
            else if (type === 'end') fields.pop();
        }
        frame.props = parent.props || frame.name === W.pPr || frame.name === W.rPr;
        if (frame.name === W.rPr && parent.name === W.r && !parent.rPr) frame.isRunProps = true;
        if (BLOCK_ELEMENTS.has(frame.name)) parent.kids++;
        const para = parent.para;
        if (frame.name === W.p) frame.para = frame;
        if (!para || frame.name === W.p) return;
        if (parent.props) {
            if (frame.name === W.sectPr) para.content = true;
        } else if (TAG_ONLY_ALLOWED.has(frame.name)) {
            // allowed in a paragraph that holds nothing but tags
        } else if (frame.name === W.bookmarkStart && attributeOf(startTag, 'w:name') === GO_BACK_BOOKMARK) {
            (para.goBack ??= new Set()).add(attributeOf(startTag, 'w:id') ?? '');
        } else if (frame.name === W.bookmarkEnd && para.goBack?.has(attributeOf(startTag, 'w:id') ?? '')) {
            // the end of Word's last-edit bookmark
        } else {
            para.content = true;
        }
    };
    const onClose = (frame: Frame, end: number): void => {
        frame.end = end;
        if (frame.isRunProps) frame.parent!.rPr = [frame.start, end];
        if (frame.name === W.p) {
            frame.tagOnly = !frame.content && frame.blockMarks.length > 0;
            if (frame.tagOnly && frame.parent) {
                frame.parent.tagOnlyKids++;
                frame.parent.lastTagOnly = frame;
            }
        }
    };

    // A `<w:t>` in a comment (or CDATA, or a declaration) was marked like any other; it is not content.
    const inert = (from: number, to: number): void => {
        for (let m = markerFrom(from); m >= 0 && m < to; m = markerFrom(m + 1)) {
            const close = xml.indexOf(ctx.close, m + 1);
            ctx.inert.push({ at: m, length: close + 1 - m, marker: Number(xml.slice(m + 1, close)) });
        }
    };
    let pos = 0;
    const n = xml.length;
    while (pos < n) {
        const lt = xml.indexOf('<', pos);
        const textEnd = lt < 0 ? n : lt;
        if (textEnd > pos) onText(pos, textEnd);
        if (lt < 0) break;
        if (xml.startsWith('<!--', lt)) {
            const e = xml.indexOf('-->', lt + 4);
            if (e < 0) malformed('a comment is not closed');
            inert(lt, e);
            pos = e + 3;
            continue;
        }
        if (xml.startsWith('<![CDATA[', lt)) {
            const e = xml.indexOf(']]>', lt + 9);
            if (e < 0) malformed('a CDATA section is not closed');
            inert(lt, e);
            pos = e + 3;
            continue;
        }
        if (xml.startsWith('<?', lt) || xml.startsWith('<!', lt)) {
            const e = xml.indexOf('>', lt + 2);
            if (e < 0) malformed('a declaration is not closed');
            inert(lt, e);
            pos = e + 1;
            continue;
        }
        if (xml.charCodeAt(lt + 1) === 0x2f) { // `</`
            const gt = xml.indexOf('>', lt + 2);
            if (gt < 0) malformed('an end tag is not closed');
            const name = xml.slice(lt + 2, gt).trim();
            if (top === root || name !== top.name) malformed(`</${name.slice(0, 40)}> does not close the element open there`);
            onClose(top, gt + 1);
            top = top.parent!;
            pos = gt + 1;
            continue;
        }
        let nameEnd = lt + 1;
        while (nameEnd < n && !NAME_END.test(xml[nameEnd])) nameEnd++;
        // The start tag's end: the first `>` outside a quoted attribute value.
        const gt = startTagEnd(xml, nameEnd);
        if (gt < 0) malformed('a start tag is not closed');
        const name = xml.slice(lt + 1, nameEnd);
        // Every element, a text run too: nested runs (malformed, but read) and runs where no tag can be
        // were counted nowhere, and each kept a frame alive.
        if (--ctx.elements.left < 0) ctx.opts.onElementLimit();
        const frame = newFrame(name, lt, gt + 1, top);
        onElement(frame, xml.slice(lt, gt + 1));
        if (xml.charCodeAt(gt - 1) === 0x2f) onClose(frame, gt + 1);
        else top = frame;
        pos = gt + 1;
    }
    if (top !== root) malformed(`<${top.name.slice(0, 40)}> is not closed`);
    return occurrences;
}

/** Whether `xml` holds anything but whitespace between `from` and `to` (without slicing it). */
function hasNonSpace(xml: string, from: number, to: number): boolean {
    for (let i = from; i < to; i++) {
        const c = xml.charCodeAt(i);
        if (c !== 0x20 && c !== 0x09 && c !== 0x0A && c !== 0x0D && c !== 0xA0 && !(c >= 0x2000 && c <= 0x200A) && c !== 0x3000) return true;
    }
    return false;
}

/** The first block tag of a part, for an error about the part as a whole. */
function firstBlockTag(ctx: PartContext): string {
    const record = ctx.tags.find(t => t.tag.kind !== 'field');
    return record ? typedOf(record, ctx.opts) : '';
}

// ─── Phase 3: blocks ────────────────────────────────────────────────────────

/** Where a tag cuts its block: content before `a` gets `close`, content after `b` gets `open`; `[a, b)` is dropped. */
interface Cut { a: number; b: number; close: string; open: string }

interface Block {
    kind: 'if' | 'for';
    /** Opening tag, any `else if`/`else`, closing tag. */
    tags: Occurrence[];
    mode: 'inline' | 'paragraph' | 'row';
    /** Row mode: the row (or row wrapper) each tag widens to. */
    units: Frame[];
    cuts: Cut[];
    order: number;
}

/** Pairs a part's block tags into blocks, in document order. */
function pairBlocks(occurrences: Occurrence[], ctx: PartContext): Block[] {
    const blocks: Block[] = [];
    const open: Block[] = [];
    /**
     * Raises a pairing error, or rather the error of a tag that could not be read where it would have
     * made the blocks pair: one whose first word would have closed the open block, or opened the block
     * the failing tag closes, inside the stretch where it would have done so. Prose elsewhere is not
     * blamed.
     */
    function syntax(reason: TemplateSyntaxDetails['reason'], occurrence: Occurrence, detail?: string, top?: Block): never {
        let words: string[] = [];
        let after = -1;
        let before = Infinity;
        const tag = ctx.tags[occurrence.tagIndex].tag;
        const closers = (kind: Block['kind']) => kind === 'if' ? ['endif', 'else'] : ['endfor'];
        const opener = tag.kind === 'endfor' ? 'for' : 'if';
        if (reason === 'unclosed') { words = closers(top!.kind); after = occurrence.tagIndex; }
        else if (reason === 'unopened' || reason === 'mismatched') {
            words = top ? [opener, ...closers(top.kind)] : [opener];
            after = top ? top.tags[0].tagIndex : -1;
            before = occurrence.tagIndex;
        }
        const unreadable = words.length
            ? ctx.unreadable.filter(u => u.beforeTag > after && u.beforeTag <= before && words.includes(u.word)).pop()
            : undefined;
        if (unreadable) return ctx.opts.onSyntax({ reason: 'unreadable', tag: typedTag(decodeXmlText(unreadable.inner), ctx.opts), part: ctx.part, detail: unreadable.reason });
        // A block already closed that crosses a boundary explains the failure better: a text box holds
        // its paragraphs twice (Word's drawing and its VML fallback), so a block closed inside one closes
        // with the first copy and its second copy belongs to nothing.
        resolveScopes(blocks.filter(b => !open.includes(b)), ctx);
        return ctx.opts.onSyntax({ reason, tag: typedOf(ctx.tags[occurrence.tagIndex], ctx.opts), part: ctx.part, detail });
    }
    for (const occurrence of occurrences) {
        if (ctx.markerPiece[occurrence.marker] !== 0) continue;
        const tag = ctx.tags[occurrence.tagIndex].tag;
        if (tag.kind === 'field') continue;
        const top = open[open.length - 1];
        switch (tag.kind) {
            case 'if':
            case 'for': {
                const block: Block = { kind: tag.kind, tags: [occurrence], mode: 'inline', units: [], cuts: [], order: blocks.length };
                blocks.push(block);
                if (open.push(block) > MAX_BLOCK_DEPTH) ctx.opts.onNestingLimit();
                occurrence.block = block;
                break;
            }
            case 'elseif':
            case 'else': {
                if (!top || top.kind !== 'if') syntax('unopened', occurrence, top ? `${ctx.opts.start} endfor ${ctx.opts.end}` : `${ctx.opts.start} if ${ctx.opts.end}`, top);
                const last = ctx.tags[top.tags[top.tags.length - 1].tagIndex].tag;
                if (last.kind === 'else') syntax('else-after-else', occurrence);
                top.tags.push(occurrence);
                occurrence.block = top;
                break;
            }
            case 'endif':
            case 'endfor': {
                const wanted = tag.kind === 'endif' ? 'if' : 'for';
                if (!top) syntax('unopened', occurrence, `${ctx.opts.start} ${wanted} ${ctx.opts.end}`);
                if (top.kind !== wanted) syntax('mismatched', occurrence, `${ctx.opts.start} end${top.kind} ${ctx.opts.end}`, top);
                top.tags.push(occurrence);
                occurrence.block = top;
                open.pop();
                break;
            }
        }
    }
    const unclosed = open[open.length - 1];
    if (unclosed) syntax('unclosed', unclosed.tags[0], `${ctx.opts.start} end${unclosed.kind} ${ctx.opts.end}`, unclosed);
    return blocks;
}

/**
 * The deepest frame holding all of `frames`, found from them upward: the deepest frames are lifted to
 * their parents, level by level, merging as they meet, until one is left. Only the frames between the
 * tags and that ancestor are visited, never the chain above it (walking it to the root for each block
 * took time quadratic in nesting, with many blocks under deep content controls).
 */
function commonAncestorOf(frames: Frame[]): Frame {
    const frontier = new Set(frames);
    if (frontier.size === 1) return frames[0];
    const byDepth = new Map<number, Frame[]>();
    let deepest = 0;
    for (const f of frontier) {
        let level = byDepth.get(f.depth);
        if (!level) byDepth.set(f.depth, level = []);
        level.push(f);
        if (f.depth > deepest) deepest = f.depth;
    }
    for (let depth = deepest; frontier.size > 1 && depth > 0; depth--) {
        for (const f of byDepth.get(depth) ?? []) {
            if (frontier.size === 1) break;
            frontier.delete(f);
            const parent = f.parent!;
            if (frontier.has(parent)) continue;
            frontier.add(parent);
            let level = byDepth.get(depth - 1);
            if (!level) byDepth.set(depth - 1, level = []);
            level.push(parent);
        }
        byDepth.delete(depth);
    }
    return frontier.values().next().value!;
}

/** What a crossing is, for its error. */
function describeCrossing(units: Frame[], container: Frame): string {
    if (container.para) return 'one is inside a text box and another is not';
    if (units.some(u => u.name === W.tbl)) return 'one is inside a table and another is outside it, or in another table';
    if (units.some(u => u.name === W.sdt)) return 'one is inside a content control and another is not';
    return 'they are in places no single stretch of content joins';
}

/**
 * Decides each block's scope from where its tags sit (inline, paragraphs, or table rows), raising
 * `crossing` or `row-branch` for tags that cannot enclose content together.
 */
function resolveScopes(blocks: Block[], ctx: PartContext): void {
    function syntax(reason: TemplateSyntaxDetails['reason'], occurrence: Occurrence, detail?: string): never {
        return ctx.opts.onSyntax({ reason, tag: typedOf(ctx.tags[occurrence.tagIndex], ctx.opts), part: ctx.part, detail });
    }
    for (const block of blocks) {
        // A field's parts are runs: a block whose tags are in different ones would cut the field apart,
        // leaving its begin without its end, or its result without its instruction.
        const stray = block.tags.find(t => t.field !== block.tags[0].field);
        if (stray) syntax('crossing', stray, 'one is inside a Word field (a hyperlink, a page number, a date) or a part of one that another is not in');
        const paragraphs = block.tags.map(t => t.frame.para);
        const missing = block.tags.find(t => !t.frame.para);
        if (missing) syntax('crossing', missing, 'it is not inside a paragraph');
        const paras = paragraphs as Frame[];
        if (paras.every(p => p === paras[0])) { block.mode = 'inline'; continue; }
        const container = commonAncestorOf(paras);
        // Paragraphs that share a paragraph: one is in a text box inside the other.
        if (container.para) syntax('crossing', block.tags[Math.max(paras.indexOf(container), 0)], describeCrossing([], container));
        // What a frame below the container sees above it, each frame worked out once for the block (a block
        // can hold any number of tags under one deep path: walking up for each of them took time quadratic in
        // the template): the topmost row between it and the container, its ancestor that is a child of the
        // container, and whether everything above it is a row wrapper.
        const above = new Map<Frame, { row: Frame | null; child: Frame; wrappers: boolean }>();
        const view = (start: Frame): { row: Frame | null; child: Frame; wrappers: boolean } => {
            const chain: Frame[] = [];
            let f: Frame | null = start;
            while (f && f !== container && !above.has(f)) { chain.push(f); f = f.parent; }
            let next = f && f !== container ? above.get(f)! : null;
            for (let k = chain.length - 1; k >= 0; k--) {
                const g = chain[k];
                const own = g.name === W.tr ? g : null;
                next = next
                    ? { row: next.row ?? own, child: next.child, wrappers: next.wrappers && ROW_WRAPPERS.has(g.name) }
                    : { row: own, child: g, wrappers: ROW_WRAPPERS.has(g.name) };
                above.set(g, next);
            }
            return above.get(start)!;
        };
        const childOfContainer = (p: Frame): Frame => view(p).child;
        // Rows: every tag in a row of the same table (or of row wrappers inside it).
        const rows = paras.map(p => p.parent && p.parent !== container ? view(p.parent).row : null);
        const rowLevel = container.name === W.tr || rows.every(row => {
            if (!row) return false;
            if (row.parent && row.parent !== container && !view(row.parent).wrappers) return false;
            return container.name === W.tbl || ROW_WRAPPERS.has(container.name);
        });
        if (rowLevel) {
            block.mode = 'row';
            block.units = container.name === W.tr ? paras.map(() => container) : paras.map(childOfContainer);
            const last = block.tags.length - 1;
            for (let i = 1; i < last; i++) {
                if (block.units[i].start <= block.units[i - 1].start) syntax('row-branch', block.tags[i]);
            }
            if (last > 1 && block.units[last].start < block.units[last - 1].start) syntax('row-branch', block.tags[last]);
            continue;
        }
        const units = paras.map(childOfContainer);
        const strayIndex = units.findIndex((u, i) => u !== paras[i]);
        if (strayIndex >= 0) syntax('crossing', block.tags[strayIndex], describeCrossing(units, container));
        block.mode = 'paragraph';
    }
}

/**
 * Which tag-only paragraphs are removed with their tags. A paragraph holding nothing but tags of
 * paragraph or row blocks is removed, unless it is the last of a container that holds only such
 * paragraphs (a cell keeps a paragraph); and a paragraph block's tags are removed with their
 * paragraphs all together or not at all, since a cut at a paragraph's edge and a cut inside a paragraph
 * do not balance. Each paragraph turns from removed to kept at most once and each block is settled at
 * most once, so this takes time linear in the number of tags.
 */
function removableParagraphs(blocks: Block[]): Set<Frame> {
    const removable = new Set<Frame>();
    const candidates = new Set<Frame>();
    for (const block of blocks) for (const tag of block.tags) candidates.add(tag.frame.para!);
    for (const p of candidates) {
        if (!p.tagOnly) continue;
        if (p.blockMarks.some(m => m.block!.mode === 'inline')) continue;
        const parent = p.parent!;
        if (NEEDS_BLOCK.has(parent.name) && parent.kids === parent.tagOnlyKids && parent.lastTagOnly === p) continue;
        removable.add(p);
    }
    // For each paragraph block, how many of its tags sit in paragraphs still to be removed; a block
    // with some but not all is settled once, by keeping all of its paragraphs, which may unsettle others.
    const removedTags = new Map<Block, number>();
    const queue: Block[] = [];
    const queued = new Set<Block>();
    const consider = (block: Block): void => {
        const count = removedTags.get(block)!;
        if (count > 0 && count < block.tags.length && !queued.has(block)) { queued.add(block); queue.push(block); }
    };
    for (const block of blocks) {
        if (block.mode !== 'paragraph') continue;
        removedTags.set(block, block.tags.filter(t => removable.has(t.frame.para!)).length);
        consider(block);
    }
    while (queue.length) {
        const block = queue.pop()!;
        for (const tag of block.tags) {
            const p = tag.frame.para!;
            if (!removable.delete(p)) continue;
            for (const m of p.blockMarks) {
                if (m.block!.mode !== 'paragraph') continue;
                removedTags.set(m.block!, removedTags.get(m.block!)! - 1);
                consider(m.block!);
            }
        }
    }
    return removable;
}

/** The closing and opening halves of cuts, by the `<w:t>` cut and where the cut stops: tags in one run share them. */
type CutCache = Map<Frame, Map<Frame, { close: string; open: string }>>;

/** A cut through the elements between a marker's `<w:t>` and `stop`: each closed before and opened after. */
function splitCut(occurrence: Occurrence, stop: Frame, xml: string, ctx: PartContext, cache: CutCache): Cut {
    let byStop = cache.get(occurrence.frame);
    if (!byStop) cache.set(occurrence.frame, byStop = new Map());
    let halves = byStop.get(stop);
    if (!halves) byStop.set(stop, halves = cutHalves(occurrence, stop, xml, ctx));
    return { a: occurrence.at, b: occurrence.at + occurrence.length, close: halves.close, open: halves.open };
}

function cutHalves(occurrence: Occurrence, stop: Frame, xml: string, ctx: PartContext): { close: string; open: string } {
    const chain: Frame[] = [];
    for (let f: Frame | null = occurrence.frame; f && f !== stop; f = f.parent) {
        if (chain.length >= MAX_CUT_DEPTH) {
            ctx.opts.onSyntax({ reason: 'crossing', tag: typedOf(ctx.tags[occurrence.tagIndex], ctx.opts), part: ctx.part, detail: `it sits inside more than ${MAX_CUT_DEPTH} nested elements, where a block can neither begin nor end` });
        }
        if (!SPLITTABLE.has(f.name)) {
            const detail = TRACKED_CHANGE.has(f.name)
                ? `it sits inside a tracked change (${f.name}); accept or reject the change first`
                : `it sits inside ${f.name}, where a block can neither begin nor end`;
            ctx.opts.onSyntax({ reason: 'crossing', tag: typedOf(ctx.tags[occurrence.tagIndex], ctx.opts), part: ctx.part, detail });
        }
        chain.push(f);
    }
    const close = chain.map(f => `</${f.name}>`).join('');
    const open = chain.reverse().map(f => xml.slice(f.start, f.tagEnd) + (f.name === W.r && f.rPr ? xml.slice(f.rPr[0], f.rPr[1]) : '')).join('');
    return { close, open };
}

/** Each block's cuts, from its scope. */
function placeCuts(blocks: Block[], removable: Set<Frame>, xml: string, ctx: PartContext): void {
    const cache: CutCache = new Map();
    for (const block of blocks) {
        if (block.mode === 'inline') {
            let stop = commonAncestorOf(block.tags.map(t => t.frame));
            while (stop.name === W.t || stop.name === W.r) stop = stop.parent!;
            block.cuts = block.tags.map(t => splitCut(t, stop, xml, ctx, cache));
        } else if (block.mode === 'paragraph') {
            block.cuts = block.tags.map(t => {
                const p = t.frame.para!;
                return removable.has(p) ? { a: p.start, b: p.end, close: '', open: '' } : splitCut(t, p, xml, ctx, cache);
            });
        } else {
            const last = block.tags.length - 1;
            block.cuts = block.units.map((unit, i) => {
                const at = i === last ? unit.end : unit.start;
                return { a: at, b: at, close: '', open: '' };
            });
        }
    }
}

// ─── Phase 4: the node tree ─────────────────────────────────────────────────

type TreeEvent =
    | { kind: 'cut'; at: number; rank: number; seq: number; cut: Cut; block: Block; index: number }
    | { kind: 'delete'; at: number; rank: number; seq: number; to: number }
    | { kind: 'field'; at: number; rank: number; seq: number; occurrence: Occurrence }
    | { kind: 'restore'; at: number; rank: number; seq: number; to: number; text: string };

/**
 * Builds nodes, numbering literals as they are made. A cut that closes a run right after it was opened
 * (a tag at the start or end of a run, or two tags side by side) cancels the opening instead of
 * writing an empty run.
 */
class NodeBuilder {
    count = 0;
    /** The cut whose opening a sequence ends with, until anything else is added to it. */
    private readonly opened = new WeakMap<TemplateNode[], Cut>();
    /** `copies` is shared by every part of the template: what cuts copy is budgeted for the template as a whole. */
    constructor(private readonly ctx: PartContext, private readonly literalBase: number, private readonly copies: { left: number }) { }
    literal(xml: string): LiteralNode { return { kind: 'literal', xml, id: this.reserve(), slots: null }; }
    /** A fresh literal id. */
    reserve(): number { return this.literalBase + this.count++; }
    /** The last id reserved. */
    lastId(): number { return this.literalBase + this.count - 1; }
    push(seq: TemplateNode[], xml: string): void {
        if (!xml) return;
        const opened = this.opened.get(seq);
        if (opened) {
            this.opened.delete(seq);
            if (opened.close && xml.startsWith(opened.close) && this.dropEnding(seq, opened.open)) xml = xml.slice(opened.close.length);
            if (!xml) return;
        }
        const last = seq[seq.length - 1];
        if (last?.kind === 'literal') last.xml += xml;
        else seq.push(this.literal(xml));
    }
    node(seq: TemplateNode[], node: TemplateNode): void {
        this.opened.delete(seq);
        seq.push(node);
    }
    /** Writes the closing half of a cut: or, when the sequence ends with the cut's own opening, drops that. */
    close(seq: TemplateNode[], cut: Cut): void {
        this.opened.delete(seq);
        if (!cut.close) return;
        if (!this.dropEnding(seq, cut.open)) { this.copied(cut.close.length); this.push(seq, cut.close); }
    }
    open(seq: TemplateNode[], cut: Cut): void {
        this.copied(cut.open.length);
        this.push(seq, cut.open);
        if (cut.open) this.opened.set(seq, cut);
    }
    /**
     * What cuts copy into the template: each opening half repeats its run's start tag and properties,
     * so a run of many tags with long properties would grow the compiled template by tags x properties.
     * Charged against `maxRepeatedContent` as it is made.
     */
    private copied(amount: number): void {
        if ((this.copies.left -= amount) < 0) this.ctx.opts.onRepeatLimit();
    }
    /** Removes `ending` from the end of the sequence's last literal, if it ends with it. */
    private dropEnding(seq: TemplateNode[], ending: string): boolean {
        const last = seq[seq.length - 1];
        if (!ending || last?.kind !== 'literal' || !last.xml.endsWith(ending)) return false;
        last.xml = last.xml.slice(0, last.xml.length - ending.length);
        if (!last.xml) seq.pop();
        return true;
    }
    field(markerIndex: number): FieldNode {
        const record = this.ctx.tags[this.ctx.markerTag[markerIndex]];
        const tag = record.tag as Extract<TemplateTag, { kind: 'field' }>;
        return {
            kind: 'field', record, piece: this.ctx.markerPiece[markerIndex], pieceText: this.ctx.pieceText[markerIndex],
            path: tag.path, text: tag.text, id: this.reserve(), cost: NODE_COST + (record.filters ? record.inner.length : tag.text.length),
        };
    }
}

/**
 * A block's stretch of the marked XML as typed, every marker written back: a field's part in each
 * `<w:t>` it spanned, so its runs keep their text; a block tag of this block or one inside it whole
 * where it began (its later parts may lie past the stretch, in runs the block does not own); and
 * nothing for a tag of another block, which can share a paragraph holding only tags with this one.
 */
function restoreMarkers(xml: string, ctx: PartContext, firstTag: number, lastTag: number): string {
    let out = '';
    let at = 0;
    for (;;) {
        const m = xml.indexOf(ctx.open, at);
        if (m < 0) break;
        const close = xml.indexOf(ctx.close, m + 1);
        const marker = Number(xml.slice(m + 1, close));
        const tagIndex = ctx.markerTag[marker];
        const record = ctx.tags[tagIndex];
        let text = '';
        if (record.tag.kind === 'field') text = ctx.pieceText[marker];
        else if (tagIndex >= firstTag && tagIndex <= lastTag && ctx.markerPiece[marker] === 0) text = ctx.startText + record.inner + ctx.endText;
        out += xml.slice(at, m) + text;
        at = close + 1;
    }
    return out + xml.slice(at);
}

/** The nodes of a part holding fields but no block: literals and fields, nothing to scan for structure. */
function buildFlatNodes(xml: string, ctx: PartContext, builder: NodeBuilder): TemplateNode[] {
    const nodes: TemplateNode[] = [];
    let at = 0;
    for (;;) {
        const m = xml.indexOf(ctx.open, at);
        if (m < 0) break;
        const close = xml.indexOf(ctx.close, m + 1);
        builder.push(nodes, xml.slice(at, m));
        const marker = Number(xml.slice(m + 1, close));
        builder.node(nodes, builder.field(marker));
        at = close + 1;
    }
    builder.push(nodes, xml.slice(at));
    return nodes;
}

/** The nodes of a part holding blocks: a walk over cuts, deletions and fields in document order. */
function buildTree(xml: string, occurrences: Occurrence[], blocks: Block[], removable: Set<Frame>, ctx: PartContext, builder: NodeBuilder): TemplateNode[] {
    const events: TreeEvent[] = [];
    for (const block of blocks) {
        block.cuts.forEach((cut, index) => events.push({ kind: 'cut', at: cut.a, rank: 0, seq: block.tags[index].at, cut, block, index }));
    }
    const cutMarkers = new Set<number>();
    for (const block of blocks) if (block.mode !== 'row') for (const t of block.tags) cutMarkers.add(t.marker);
    for (const occurrence of occurrences) {
        if (ctx.tags[occurrence.tagIndex].tag.kind === 'field') events.push({ kind: 'field', at: occurrence.at, rank: 1, seq: occurrence.at, occurrence });
        else if (!cutMarkers.has(occurrence.marker)) events.push({ kind: 'delete', at: occurrence.at, rank: 1, seq: occurrence.at, to: occurrence.at + occurrence.length });
    }
    for (const p of removable) events.push({ kind: 'delete', at: p.start, rank: 1, seq: p.start, to: p.end });
    for (const m of ctx.inert) events.push({ kind: 'restore', at: m.at, rank: 1, seq: m.at, to: m.at + m.length, text: ctx.pieceText[m.marker] });
    events.sort((x, y) => x.at - y.at || x.rank - y.rank || x.seq - y.seq);

    const root: TemplateNode[] = [];
    const stack: { seq: TemplateNode[]; block: Block | null; node: IfNode | ForNode | null }[] = [{ seq: root, block: null, node: null }];
    let cursor = 0;
    const current = () => stack[stack.length - 1];
    for (const event of events) {
        if (event.kind !== 'cut' && event.at < cursor) continue; // inside a stretch already dropped
        if (event.at > cursor) builder.push(current().seq, xml.slice(cursor, event.at));
        if (event.kind === 'delete') { cursor = Math.max(cursor, event.to); continue; }
        if (event.kind === 'restore') { builder.push(current().seq, event.text); cursor = event.to; continue; }
        if (event.kind === 'field') {
            builder.node(current().seq, builder.field(event.occurrence.marker));
            cursor = event.at + event.occurrence.length;
            continue;
        }
        const { block, cut, index } = event;
        builder.close(current().seq, cut);
        if (index === 0) {
            const node = blockNode(block, xml, ctx, builder);
            builder.node(current().seq, node);
            const seq = node.kind === 'if' ? node.sections[0] : node.body;
            stack.push({ seq, block, node });
        } else {
            const frame = current();
            if (frame.block !== block) {
                const inner = frame.block ?? block;
                ctx.opts.onSyntax({ reason: 'overlap', tag: typedOf(ctx.tags[inner.tags[0].tagIndex], ctx.opts), part: ctx.part });
            }
            if (index === block.tags.length - 1) {
                // Every literal and field id of the block lies from its kept id to here.
                frame.node!.kept.lastId = builder.lastId();
                stack.pop();
            }
            else {
                const node = frame.node as IfNode;
                const section: TemplateNode[] = [];
                node.sections.push(section);
                frame.seq = section;
            }
        }
        builder.open(current().seq, cut);
        cursor = Math.max(cursor, cut.b);
    }
    builder.push(current().seq, xml.slice(cursor));
    return root;
}

/** A block's node, with what it writes when it is kept as typed. */
function blockNode(block: Block, xml: string, ctx: PartContext, builder: NodeBuilder): IfNode | ForNode {
    const first = block.cuts[0];
    const last = block.cuts[block.cuts.length - 1];
    const kept: KeptSpan = {
        from: first.a, to: last.b, open: first.open, close: last.close, xml, ctx, id: builder.reserve(), lastId: -1, inLoop: false,
        firstTag: block.tags[0].tagIndex, lastTag: block.tags[block.tags.length - 1].tagIndex,
    };
    const tags = block.tags.map(t => ctx.tags[t.tagIndex]);
    const open = tags[0].tag;
    if (open.kind === 'for') return { kind: 'for', name: open.name, list: open.list, listText: open.listText, record: tags[0], body: [], kept, cost: NODE_COST + open.listText.length };
    const branches = tags.slice(0, -1).map(record => {
        const tag = record.tag;
        const condition = tag.kind === 'if' || tag.kind === 'elseif' ? tag.condition : null;
        return { condition, record, cost: NODE_COST + (condition ? conditionSize(condition) : 0) };
    });
    return { kind: 'if', branches, sections: [[]], kept };
}

// ─── Unique ids in repeated content ─────────────────────────────────────────

/** Where fresh ids start: past every id of each kind in the document; and the bookmark names it has. */
interface IdSeeds { docPr: number; annotation: number; sdt: number; para: number; paraIds: Set<number>; bookmarkNames: Set<string> }

/** The unique ids in a start tag: each attribute, and its kind. */
function idAttributes(name: string): [string, IdKind][] {
    switch (name) {
        case 'wp:docPr': return [['id', 'docPr']];
        case W.bookmarkStart: return [['w:id', 'bookmarkStart'], ['w:name', 'bookmarkName']];
        case W.bookmarkEnd: return [['w:id', 'bookmarkEnd']];
        case 'w:id': return [['w:val', 'sdt']];
        case 'w:footnoteReference': return [['w:id', 'footnoteRef']];
        case 'w:endnoteReference': return [['w:id', 'endnoteRef']];
        case 'w:commentRangeStart': return [['w:id', 'commentStart']];
        case 'w:commentRangeEnd': return [['w:id', 'commentEnd']];
        case 'w:commentReference': return [['w:id', 'commentRef']];
        default: return [];
    }
}

/**
 * Where a start tag's attributes named in `wanted` start and end, and where the tag ends: one pass that
 * reads quoted values as values, so a `>` or ` id="` inside another attribute's value (alt text such as
 * `Badge id=`) is never taken for the tag's end or for an attribute. Only the wanted few are kept, so a
 * tag of millions of attributes holds no more than they do. Null for a tag that never ends.
 */
function readStartTag(xml: string, nameEnd: number, wanted: readonly string[]): { gt: number; attributes: Map<string, [number, number]> } | null {
    const attributes = new Map<string, [number, number]>();
    let at = nameEnd;
    for (;;) {
        while (at < xml.length && /\s/.test(xml[at])) at++;
        if (at >= xml.length) return null;
        if (xml[at] === '>') return { gt: at, attributes };
        if (xml[at] === '/' && xml[at + 1] === '>') return { gt: at + 1, attributes };
        const nameStart = at;
        while (at < xml.length && !/[\s=>/]/.test(xml[at])) at++;
        const name = wanted.find(w => w.length === at - nameStart && xml.startsWith(w, nameStart));
        while (at < xml.length && /\s/.test(xml[at])) at++;
        if (xml[at] !== '=') { if (at === nameStart) at++; continue; } // a name with no value (not XML): skipped
        at++;
        while (at < xml.length && /\s/.test(xml[at])) at++;
        const quote = xml[at];
        if (quote !== '"' && quote !== "'") continue;
        const close = xml.indexOf(quote, at + 1);
        if (close < 0) return null;
        if (name !== undefined && !attributes.has(name)) attributes.set(name, [at + 1, close]);
        at = close + 1;
    }
}

/** Every unique id in a stretch of XML, in order, by a forward scan of its start tags. */
function findIds(xml: string): IdSlot[] {
    const slots: IdSlot[] = [];
    visitIds(xml, slot => { slots.push(slot); });
    return slots;
}

/** Each unique id in a stretch of XML, in order, handed to `visit` as it is found, so no list is kept. */
function visitIds(xml: string, visit: (slot: IdSlot) => void): void {
    for (let lt = xml.indexOf('<'); lt >= 0 && lt < xml.length; lt = xml.indexOf('<', lt + 1)) {
        const c = xml.charCodeAt(lt + 1);
        if (c === 0x2f) continue; // an end tag
        const inert = inertEnd(xml, lt); // a comment, CDATA section, declaration or processing instruction, skipped whole
        if (inert !== null) { if (inert < 0) break; lt = inert; continue; }
        let nameEnd = lt + 1;
        while (nameEnd < xml.length && xml[nameEnd] !== '>' && !NAME_END.test(xml[nameEnd])) nameEnd++;
        // Any element can carry paragraph ids; the others belong to the element they are named for.
        const wanted = [...idAttributes(xml.slice(lt + 1, nameEnd)), ['w14:paraId', 'paraId'], ['w14:textId', 'textId']] as [string, IdKind][];
        const tag = readStartTag(xml, nameEnd, wanted.map(([attribute]) => attribute));
        if (!tag) break;
        const inTag: IdSlot[] = [];
        for (const [attribute, kind] of wanted) {
            const value = tag.attributes.get(attribute);
            if (!value) continue;
            inTag.push({ start: value[0], end: value[1], kind, value: xml.slice(value[0], value[1]), tag: lt, gt: tag.gt });
        }
        // A tag's attributes come in any order; the slots are written in the order they stand.
        for (const slot of inTag.sort((x, y) => x.start - y.start)) visit(slot);
        lt = tag.gt;
    }
}

/** An annotation's id attribute (bookmarks, comments, notes, tracked changes), quoted either way. */
const ANNOTATION_ID = /\sw:id\s*=\s*["'](\d{1,15})["']/g;

/** Fresh ids start past every id the template already has, in every part. */
function seedIds(entries: Record<string, Uint8Array>, chargeFor: (part: string) => (() => void) | undefined): IdSeeds {
    const seeds: IdSeeds = { docPr: 1, annotation: 1, sdt: 1, para: 1, paraIds: new Set(), bookmarkNames: new Set() };
    for (const [name, bytes] of Object.entries(entries)) {
        if (!TEXT_PART.test(name)) continue;
        const xml = Buffer.from(bytes).toString('utf8');
        // Bookmarks, comments and notes are numbered among the document's annotations, so a fresh one
        // starts past every `w:id` attribute in the document.
        for (const m of xml.matchAll(ANNOTATION_ID)) seeds.annotation = Math.max(seeds.annotation, Number(m[1]) + 1);
        const charge = chargeFor(name);
        visitIds(xml, slot => {
            charge?.();
            if (slot.kind === 'bookmarkName') { seeds.bookmarkNames.add(slot.value); return; }
            if (slot.kind === 'paraId' || slot.kind === 'textId') {
                const id = parseInt(slot.value, 16);
                if (Number.isFinite(id)) { seeds.paraIds.add(id); seeds.para = Math.max(seeds.para, id + 1); }
                return;
            }
            const id = parseInt(slot.value, 10);
            if (!Number.isFinite(id)) return;
            if (slot.kind === 'docPr') seeds.docPr = Math.max(seeds.docPr, id + 1);
            else if (slot.kind === 'sdt') seeds.sdt = Math.max(seeds.sdt, id + 1);
        });
    }
    return seeds;
}

/** Gives every literal inside a loop its id slots. */
function markRepeatedIds(nodes: TemplateNode[], inLoop: boolean): boolean {
    let found = false;
    for (const node of nodes) {
        switch (node.kind) {
            case 'literal':
                if (inLoop) {
                    const slots = findIds(node.xml);
                    if (slots.length) { node.slots = slots; found = true; }
                }
                break;
            case 'field': break;
            case 'if':
                for (const section of node.sections) found = markRepeatedIds(section, inLoop) || found;
                // A kept block's ids are found when it is written (see renderKept); it may hold some.
                if (inLoop) node.kept.inLoop = found = true;
                break;
            case 'for':
                found = markRepeatedIds(node.body, true) || found;
                if (inLoop) node.kept.inLoop = found = true;
                break;
        }
    }
    return found;
}

/** A note a repeated reference points at, to be copied under a fresh id. */
/** A note to write: a copy (`from` written again under `to`), or (`kept`) a note written under its own id. */
interface NoteCopy { part: string; element: string; from: string; to: string; scope: Scope | null; kept?: true }

/**
 * Per render: which ids have been written, and the next fresh one of each kind. The first write of an
 * id keeps it and each later write gets a fresh one. A bookmark's end, and a comment's range end and
 * reference, follow the start written before them. A note referred to inside a loop (a footnote,
 * endnote or comment) gets a copy of its own at every pass, rendered in that pass's scope, so its
 * `{{ item }}` reads the pass's item (see {@link copyNotes}).
 */
class IdState {
    readonly copies: NoteCopy[] = [];
    private readonly keptNotes = new Set<string>();
    private across = new Set<string>();
    private readonly seen = new Map<string, Set<string>>();
    private readonly pending = new Map<string, string>();
    private readonly madeNames = new Set<string>();
    /** The comment ids copies were made under: a note's own loop gives its comments these as it renders. */
    private readonly madeComments = new Set<string>();
    private docPr: number;
    private annotation: number;
    private sdt: number;
    private para: number;
    private nameNumber = 2;
    constructor(private readonly seeds: IdSeeds, private rangesAcross: Set<string> = new Set()) {
        this.docPr = seeds.docPr;
        this.annotation = seeds.annotation;
        this.sdt = seeds.sdt;
        this.para = seeds.para;
    }
    /**
     * Renders a note written on its own (a copy, or one written back as typed): the ranges across its own
     * loops are those, each written once in this note, apart from the part's and every other copy's.
     */
    inNote<T>(rangesAcross: Set<string>, render: () => T): T {
        const saved: [Set<string>, Set<string>] = [this.rangesAcross, this.across];
        this.rangesAcross = rangesAcross;
        this.across = new Set();
        try { return render(); } finally { [this.rangesAcross, this.across] = saved; }
    }
    /**
     * For an element of a range across a loop's edge (see rangesAcrossLoops): whether this kind of its
     * elements was written already, so this one is left out. Its first is written, under its own id.
     */
    writtenAcross(range: string, kind: IdKind): boolean {
        if (!this.rangesAcross.has(range)) return false;
        const key = `${range}:${kind}`;
        if (this.across.has(key)) return true;
        this.across.add(key);
        return false;
    }
    take(kind: IdKind, value: string, scope: Scope | null): string {
        const range = kind === 'bookmarkStart' || kind === 'bookmarkEnd' ? `bookmark:${value}`
            : kind === 'commentStart' || kind === 'commentEnd' || kind === 'commentRef' ? `comment:${value}` : null;
        if (range && this.rangesAcross.has(range)) return value;
        switch (kind) {
            case 'bookmarkStart': {
                const id = this.first('bookmark', value) ? value : String(this.annotation++);
                this.follow('bookmarkEnd', value, id);
                return id;
            }
            // An end with no start in this pass (a branch skipped it) gets an id of its own, matching
            // nothing: under the last pass's id it closed that pass's range a second time.
            case 'bookmarkEnd': return this.followed('bookmarkEnd', value) ?? String(this.annotation++);
            case 'bookmarkName': return this.first('bookmarkName', value) ? value : this.freshName(value);
            case 'commentStart': {
                const id = this.copyNote('comment', value, scope);
                this.follow('commentEnd', value, id);
                this.follow('commentRef', value, id);
                return id;
            }
            case 'commentEnd': return this.followed('commentEnd', value) ?? String(this.annotation++);
            case 'commentRef': return this.followed('commentRef', value) ?? this.copyNote('comment', value, scope);
            case 'footnoteRef':
            case 'endnoteRef': return this.copyNote(kind, value, scope);
            case 'docPr': return this.first(kind, value) ? value : this.freshDocPr();
            case 'sdt': return this.first(kind, value) ? value : String(this.sdt++);
            case 'paraId':
            case 'textId': return this.first(kind, value) ? value : this.freshPara();
        }
    }
    /**
     * A note reference written as typed, under the note's own id, in a block kept as typed outside any
     * loop. A note only loops refer to was cut from its part when the template was compiled, so it is
     * written back, as typed (see {@link copyNotes}).
     */
    keep(kind: IdKind, value: string): void {
        const note = kind === 'footnoteRef' || kind === 'endnoteRef' ? kind : kind === 'commentStart' || kind === 'commentRef' ? 'comment' : null;
        // A comment id a copy was made under (in a note's own loop) is that copy's, written already.
        if (!note || this.keptNotes.has(`${note}:${value}`) || (note === 'comment' && this.madeComments.has(value))) return;
        this.keptNotes.add(`${note}:${value}`);
        this.copies.push({ ...NOTE_PARTS[note], from: value, to: value, scope: null, kept: true });
    }
    freshDocPr(): string { return String(this.docPr++); }
    /**
     * Fresh ids for everything in a copied note: pictures, paragraphs, content controls, and bookmarks
     * and comment anchors (their starts, ends and references renumbered together within the copy). A
     * comment anchor a loop inside the note already gave a copy of its own keeps that copy's id: taken
     * again, it named a copy of a comment that does not exist, and the filled comment lost its anchor.
     */
    renumberCopy(note: string, scope: Scope | null): string {
        const slots = findIds(note);
        if (!slots.length) return note;
        const local = new Map<string, string>();
        const together = (space: string, value: string, make: () => string): string => {
            const key = `${space}:${value}`;
            let id = local.get(key);
            if (id === undefined) local.set(key, id = make());
            return id;
        };
        const pieces: string[] = [];
        let at = 0;
        for (const slot of slots) {
            let id: string;
            switch (slot.kind) {
                case 'docPr': id = this.freshDocPr(); break;
                case 'paraId':
                case 'textId': id = this.freshPara(); break;
                case 'sdt': id = String(this.sdt++); break;
                case 'bookmarkName': id = this.freshName(slot.value); break;
                case 'bookmarkStart':
                case 'bookmarkEnd': id = together('bookmark', slot.value, () => String(this.annotation++)); break;
                case 'commentStart':
                case 'commentEnd':
                case 'commentRef': id = this.madeComments.has(slot.value) ? slot.value : together('comment', slot.value, () => this.copyNote('comment', slot.value, scope)); break;
                case 'footnoteRef':
                case 'endnoteRef': id = slot.value; break;
            }
            pieces.push(note.slice(at, slot.start), id);
            at = slot.end;
        }
        pieces.push(note.slice(at));
        return pieces.join('');
    }
    freshPara(): string {
        if (this.para > MAX_PARA_ID) this.para = 1;
        while (this.seeds.paraIds.has(this.para)) this.para++;
        return (this.para++).toString(16).toUpperCase().padStart(8, '0');
    }
    /** Whether `value` is written for the first time in its space (and marks it written). */
    private first(space: string, value: string): boolean {
        let seen = this.seen.get(space);
        if (!seen) this.seen.set(space, seen = new Set());
        if (seen.has(value)) return false;
        seen.add(value);
        return true;
    }
    /** The id the range opened last under `value` goes by: its end and reference follow the nearest start before them. */
    private follow(kind: string, value: string, fresh: string): void {
        this.pending.set(`${kind}:${value}`, fresh);
    }
    /** The id an end or reference follows, used once: the next pass's follows its own start. */
    private followed(kind: string, value: string): string | undefined {
        const key = `${kind}:${value}`;
        const id = this.pending.get(key);
        this.pending.delete(key);
        return id;
    }
    private copyNote(kind: keyof typeof NOTE_PARTS, from: string, scope: Scope | null): string {
        const to = String(this.annotation++);
        this.copies.push({ ...NOTE_PARTS[kind], from, to, scope });
        if (kind === 'comment') this.madeComments.add(to);
        return to;
    }
    /** A bookmark name the document does not have: the name with a number, within Word's 40 characters. */
    freshName(name: string): string {
        // Cut the name as its characters, not as the XML text it is written in (an entity, or a character
        // outside the BMP, is one character), and write it back escaped.
        const characters = Array.from(decodeXmlText(name));
        for (;;) {
            const suffix = `_${this.nameNumber++}`;
            const candidate = escapeXml(characters.slice(0, MAX_BOOKMARK_NAME - suffix.length).join('') + suffix);
            if (!this.seeds.bookmarkNames.has(candidate) && !this.madeNames.has(candidate)) {
                this.madeNames.add(candidate);
                return candidate;
            }
        }
    }
}

/**
 * The notes of a notes part by id: where each element starts and ends. One forward scan; `charge` is
 * called for each note before it is kept.
 */
function noteSpans(xml: string, element: string, charge?: () => void): Map<string, [number, number]> {
    const notes = new Map<string, [number, number]>();
    const open = `<${element}`;
    const close = `</${element}>`;
    // From `<` to `<`, a comment (or CDATA section, or processing instruction) passed whole: a note's
    // tag in one is none.
    for (let at = xml.indexOf('<'); at >= 0; at = xml.indexOf('<', at + 1)) {
        const inert = inertEnd(xml, at);
        if (inert !== null) { if (inert < 0) break; at = inert; continue; }
        if (!xml.startsWith(open, at) || !NAME_END.test(xml.charAt(at + open.length))) continue; // <w:footnotes>, <w:footnoteReference>, ...
        const gt = startTagEnd(xml, at + open.length);
        if (gt < 0) break;
        const id = attributeOf(xml.slice(at, gt + 1), 'w:id');
        const end = xml.charCodeAt(gt - 1) === 0x2f ? gt + 1 : xml.indexOf(close, gt) + close.length;
        if (end < gt) break;
        if (id !== null && !notes.has(id)) { charge?.(); notes.set(id, [at, end]); }
        at = end - 1;
    }
    return notes;
}

/** Where the last `close` tag outside a comment (or CDATA section, or processing instruction) starts, or -1. */
function containerEndOf(xml: string, close: string): number {
    let found = -1;
    for (let at = xml.indexOf('<'); at >= 0; at = xml.indexOf('<', at + 1)) {
        const inert = inertEnd(xml, at);
        if (inert !== null) { if (inert < 0) break; at = inert; continue; }
        if (xml.startsWith(close, at)) found = at;
    }
    return found;
}

/** A note with the id in its own start tag (quoted either way) replaced by `id`. */
function withNoteId(note: string, id: string): string {
    let nameEnd = 1;
    while (nameEnd < note.length && !NAME_END.test(note[nameEnd])) nameEnd++;
    const value = readStartTag(note, nameEnd, ['w:id'])?.attributes.get('w:id');
    return value ? note.slice(0, value[0]) + id + note.slice(value[1]) : note;
}

/**
 * Adds to a notes part (footnotes, endnotes, comments) a copy of each note a reference inside a loop
 * points at, under the reference's fresh id: the note rendered in the scope of the pass that refers to
 * it, so its fields read that pass's item, with fresh ids of its own throughout. (A note only loops
 * refer to was cut from its part when the template was compiled: it is written only as these copies.)
 * Each copy is charged to the repetition budget before it is made.
 */
function copyNotes(xml: string, part: string, element: string, copies: NoteCopy[], compiled: CompiledDocxTemplate, state: RenderState): string {
    const notes = noteSpans(xml, element);
    const containerEnd = containerEndOf(xml, `</${element}s>`);
    if (containerEnd < 0) return xml;
    const templates = compiled.noteTemplates.get(part);
    const made: string[] = [];
    for (const copy of copies) {
        const template = templates?.get(copy.from);
        const span = notes.get(copy.from);
        if (copy.kept) {
            // A note a block kept as typed refers to: written as typed when the template's compile cut it
            // (only loops referred to it), else it is already in the part.
            if (span || !template) continue;
            const pieces: string[] = [];
            const nodes = template.nodes;
            if (nodes) state.ids!.inNote(template.rangesAcross, () => renderNodes(nodes, null, state, pieces));
            // Repaired as its part was (that ran before the notes were written): a range its loop or branch
            // left open is closed, and a note left with no paragraph gets one.
            const text = nodes ? repairStructure(pieces.join(''), state.opts.onOutputLimit) : template.xml;
            // The comments it anchors were cut with it, so they are written as typed too: the note came
            // back pointing at comments that were not.
            const ids = state.ids!;
            visitIds(text, slot => ids.keep(slot.kind, slot.value));
            made.push(text);
            continue;
        }
        if (!template && !span) continue;
        // Charged before it is made: each copy is a repeat of its note.
        state.charge(template ? template.xml.length : span![1] - span![0]);
        let note: string;
        if (template?.nodes) {
            const pieces: string[] = [];
            const nodes = template.nodes;
            state.ids!.inNote(template.rangesAcross, () => renderNodes(nodes, copy.scope, state, pieces));
            note = repairStructure(pieces.join(''), state.opts.onOutputLimit);
        } else {
            note = template ? template.xml : xml.slice(span![0], span![1]);
        }
        made.push(state.ids!.renumberCopy(withNoteId(note, copy.to), copy.scope));
    }
    if (!made.length) return xml;
    const added = made.join('');
    if (xml.length + added.length > MAX_PART_LENGTH) state.opts.onOutputLimit();
    return xml.slice(0, containerEnd) + added + xml.slice(containerEnd);
}

// ─── Compile ────────────────────────────────────────────────────────────────

/** Compiles one part: null nodes when it holds no tag. */
function compilePart(name: string, xml: string, opts: DocxTemplateOptions, literalBase: number, copies: { left: number }, elements: ElementBudget): { part: CompiledPart; literals: number } {
    const markerChars = chooseMarkerChars(xml);
    // A part using every noncharacter and private-use character cannot be marked; it holds no tags we
    // could tell from its own text, so it is copied as is (only a crafted part reaches this).
    if (!markerChars) return { part: { name, nodes: null, hasBlocks: false }, literals: 0 };
    const ctx: PartContext = {
        part: name, opts, startText: escapeXmlText(opts.start), endText: escapeXmlText(opts.end),
        tags: [], markerTag: [], markerPiece: [], pieceText: [], open: markerChars[0], close: markerChars[1], hasBlocks: false, unreadable: [], inert: [], elements,
    };
    const marked = markPart(xml, ctx);
    if (!ctx.tags.length) return { part: { name, nodes: null, hasBlocks: false }, literals: 0 };
    const builder = new NodeBuilder(ctx, literalBase, copies);
    if (!ctx.hasBlocks) return { part: { name, nodes: buildFlatNodes(marked, ctx, builder), hasBlocks: false }, literals: builder.count };
    const occurrences = scanStructure(marked, ctx);
    const blocks = pairBlocks(occurrences, ctx);
    resolveScopes(blocks, ctx);
    const removable = removableParagraphs(blocks);
    placeCuts(blocks, removable, marked, ctx);
    const nodes = buildTree(marked, occurrences, blocks, removable, ctx, builder);
    return { part: { name, nodes, hasBlocks: true }, literals: builder.count };
}

/** Compiles a template's parts once, for every data entry of a batch. Raises the tag errors up front. */
export function compileDocxTemplate(entries: Record<string, Uint8Array>, opts: DocxTemplateOptions): CompiledDocxTemplate {
    const parts = new Map<string, CompiledPart>();
    let literalCount = 0;
    let repeatsIds = false;
    const copies = { left: opts.maxRepeatedContent };
    // Each element read costs memory, as when parsing: the template as a whole is held to the parser's bound.
    const elements: ElementBudget = { left: opts.maxXmlElements, textRuns: opts.maxXmlElements };
    for (const [name, bytes] of Object.entries(entries)) {
        if (!TEXT_PART.test(name)) continue;
        const { part, literals } = compilePart(name, Buffer.from(bytes).toString('utf8'), opts, literalCount, copies, elements);
        literalCount += literals;
        if (part.nodes && part.hasBlocks) repeatsIds = markRepeatedIds(part.nodes, false) || repeatsIds;
        parts.set(name, part);
    }
    // Only a part with blocks is read element by element (scanStructure), its elements counted. In any
    // other (no tag, or fields only) the ids and notes a scan finds are counted to the same bound, each
    // scan on its own, as each holds what it finds (note references, fresh-id seeds, notes).
    const unreadScan = (): ((part: string) => (() => void) | undefined) => {
        const left = { found: opts.maxXmlElements };
        return part => parts.get(part)?.hasBlocks ? undefined : () => { if (--left.found < 0) opts.onElementLimit(); };
    };
    const noteTemplates = new Map<string, Map<string, NoteTemplate>>();
    const notesPartXml = new Map<string, string>();
    const noteElements: ElementBudget = { left: opts.maxXmlElements, textRuns: opts.maxXmlElements };
    if (repeatsIds) {
        // Each note on its own, to be rendered in the scope of a reference to it inside a loop. Its
        // elements were counted against maxXmlElements with its part when that was read element by
        // element (this reading is a second); otherwise each note is counted as it is found.
        const chargeNotes = unreadScan();
        for (const { part, element } of Object.values(NOTE_PARTS)) {
            if (!entries[part]) continue;
            const xml = Buffer.from(entries[part]).toString('utf8');
            const templates = new Map<string, NoteTemplate>();
            for (const [id, [start, end]] of noteSpans(xml, element, chargeNotes(part))) {
                const noteXml = xml.slice(start, end);
                const compiledNote = compilePart(part, noteXml, opts, literalCount, copies, noteElements);
                literalCount += compiledNote.literals;
                if (compiledNote.part.nodes && compiledNote.part.hasBlocks) markRepeatedIds(compiledNote.part.nodes, false);
                templates.set(id, { nodes: compiledNote.part.nodes, xml: noteXml, rangesAcross: compiledNote.part.nodes ? rangesAcrossLoops(compiledNote.part.nodes) : new Set() });
            }
            noteTemplates.set(part, templates);
        }
        // Notes referred to inside loops, and those referred to outside them: a note's references in a part,
        // less those in its loops' content.
        const inLoops = new Set<string>();
        const outside = new Set<string>();
        const chargeUnread = unreadScan();
        const chargeCut = unreadScan();
        // What a note only loops refer to holds (a comment anchored in a footnote) is written only with the
        // note's copies, one per pass, as if in the loop: so the notes parts are counted after the parts
        // that refer to them, footnotes and endnotes before the comments they can hold.
        const rank = (name: string): number => name === NOTE_PARTS.comment.part ? 2 : name === NOTE_PARTS.footnoteRef.part || name === NOTE_PARTS.endnoteRef.part ? 1 : 0;
        const textParts = Object.keys(entries).filter(name => TEXT_PART.test(name)).sort((a, b) => rank(a) - rank(b));
        for (const name of textParts) {
            const counts = new Map<string, number>();
            const countOne = (slot: IdSlot, by: number): void => {
                const part = slot.kind === 'footnoteRef' ? NOTE_PARTS.footnoteRef.part : slot.kind === 'endnoteRef' ? NOTE_PARTS.endnoteRef.part
                    : slot.kind === 'commentStart' || slot.kind === 'commentRef' ? NOTE_PARTS.comment.part : null;
                if (!part) return;
                const key = `${part}:${slot.value}`;
                counts.set(key, (counts.get(key) ?? 0) + by);
                if (by < 0) inLoops.add(key);
            };
            const charge = chargeUnread(name);
            const xml = Buffer.from(entries[name]).toString('utf8');
            const element = Object.values(NOTE_PARTS).find(note => note.part === name)?.element;
            const cut = element ? [...noteSpans(xml, element, chargeCut(name))].filter(([id]) => inLoops.has(`${name}:${id}`) && !outside.has(`${name}:${id}`))
                .map(([, span]) => span).sort((a, b) => a[0] - b[0]) : [];
            let next = 0;
            visitIds(xml, slot => {
                charge?.();
                while (next < cut.length && cut[next][1] <= slot.start) next++;
                countOne(slot, next < cut.length && cut[next][0] <= slot.start ? -1 : 1);
            });
            const compiledPart = parts.get(name);
            if (compiledPart?.nodes) {
                // A range across a loop's edge is written once, under its own id (see rangesAcrossLoops):
                // its comment is referred to as from outside the loops, not cut to be copied per pass.
                const across = rangesAcrossLoops(compiledPart.nodes);
                forEachLoopSlots(compiledPart.nodes, slots => {
                    for (const slot of slots) {
                        const range = rangeOf(slot);
                        if (!range || !across.has(range)) countOne(slot, -1);
                    }
                });
            }
            for (const [key, n] of counts) if (n > 0) outside.add(key);
        }
        // Each notes part without the notes only loops refer to, compiled again from that text.
        for (const { part, element } of Object.values(NOTE_PARTS)) {
            if (!entries[part]) continue;
            const xml = Buffer.from(entries[part]).toString('utf8');
            const cut = [...noteSpans(xml, element)].filter(([id]) => inLoops.has(`${part}:${id}`) && !outside.has(`${part}:${id}`)).map(([, span]) => span).sort((a, b) => a[0] - b[0]);
            if (!cut.length) continue;
            const kept: string[] = [];
            let at = 0;
            for (const [start, end] of cut) { kept.push(xml.slice(at, start)); at = end; }
            kept.push(xml.slice(at));
            const stripped = kept.join('');
            notesPartXml.set(part, stripped);
            const recompiled = compilePart(part, stripped, opts, literalCount, copies, noteElements);
            literalCount += recompiled.literals;
            if (recompiled.part.nodes && recompiled.part.hasBlocks) markRepeatedIds(recompiled.part.nodes, false);
            parts.set(part, recompiled.part);
        }
    }
    const rangesAcross = new Set<string>();
    if (repeatsIds) for (const part of parts.values()) if (part.nodes) for (const key of rangesAcrossLoops(part.nodes)) rangesAcross.add(key);
    return { entries, parts, ids: repeatsIds ? seedIds(entries, unreadScan()) : null, literalCount, noteTemplates, notesPartXml, rangesAcrossLoops: rangesAcross };
}

/** The range a slot is an end of (a bookmark's or a comment's), as `bookmark:<id>` or `comment:<id>`, or null. */
function rangeOf(slot: IdSlot): string | null {
    if (slot.kind === 'bookmarkStart' || slot.kind === 'bookmarkEnd') return `bookmark:${slot.value}`;
    if (slot.kind === 'commentStart' || slot.kind === 'commentEnd' || slot.kind === 'commentRef') return `comment:${slot.value}`;
    return null;
}

/**
 * The ranges of a part whose start and end are in different loops, or one in a loop and one outside: a
 * bookmark or comment that opens before `{{ for }}` and closes inside it, or the other way round. Repeated
 * with the loop, its end (or start) would be written once per pass under one id, or each pass would open a
 * range of its own that nothing closes. Such a range is written once instead: the first pass writes its end
 * in the loop, under its own id, and later passes leave it out.
 */
function rangesAcrossLoops(nodes: TemplateNode[]): Set<string> {
    const where = new Map<string, Set<ForNode | null>>();
    const visit = (list: TemplateNode[], loop: ForNode | null): void => {
        for (const node of list) {
            if (node.kind === 'literal') {
                for (const slot of node.slots ?? (loop ? [] : findIds(node.xml))) {
                    const range = rangeOf(slot);
                    if (!range) continue;
                    let loops = where.get(range);
                    if (!loops) where.set(range, loops = new Set());
                    loops.add(loop);
                }
            } else if (node.kind === 'if') node.sections.forEach(section => visit(section, loop));
            else if (node.kind === 'for') visit(node.body, node);
        }
    };
    visit(nodes, null);
    return new Set([...where].filter(([, loops]) => loops.size > 1).map(([range]) => range));
}

/**
 * Where an empty element starting at `from` (its start tag ending at `tagEnd`) ends: after its `/>`, or
 * after the close tag right behind it. -1 when it holds content: such an element is not left out, as
 * finding its end means searching past its content (and a range's ends are empty as Word writes them).
 */
function emptyElementEnd(xml: string, from: number, tagEnd: number): number {
    if (tagEnd < 0) return -1;
    if (xml.charCodeAt(tagEnd - 1) === 0x2f) return tagEnd + 1;
    const name = /^<([^\s>/]+)/.exec(xml.slice(from, tagEnd + 1))?.[1];
    const close = `</${name}>`;
    return name && xml.startsWith(close, tagEnd + 1) ? tagEnd + 1 + close.length : -1;
}

/**
 * Writes `xml` with its id slots filled in for this pass: each id as the render's IdState gives it, and an
 * element that ends a range across a loop's edge left out after its first pass.
 */
function writeSlots(xml: string, slots: IdSlot[], scope: Scope | null, state: RenderState, out: string[]): void {
    const ids = state.ids!;
    let at = 0;
    let element = -1;
    for (let i = 0; i < slots.length; i++) {
        const slot = slots[i];
        if (slot.start < at) continue; // in an element already left out
        // Whether to leave the element out is decided at its first slot, before any of it is written: a
        // bookmark's name may come before its id (Word writes both orders). A range's ids are in its own
        // start tag, so only that tag's slots are looked at, and only an empty element is left out, its end
        // found next to its start: looking through elements to their ends took time quadratic in nested
        // elements, and leaving out the start of one with content wrote unbalanced XML.
        // The tag's bounds as the id scan read them: read again from the slot, a quote outside any value
        // could pair otherwise, and every later slot was looked through for each one.
        const from = slot.tag;
        if (from !== element) {
            element = from;
            const tagEnd = slot.gt;
            let leaveOut = false;
            for (let k = i; k < slots.length && slots[k].tag === from; k++) {
                const range = rangeOf(slots[k]);
                if (range) { leaveOut = ids.writtenAcross(range, slots[k].kind); break; }
            }
            const to = leaveOut ? emptyElementEnd(xml, from, tagEnd) : -1;
            if (to >= 0) {
                state.emit(out, xml.slice(at, from));
                at = to;
                continue;
            }
        }
        state.emit(out, xml.slice(at, slot.start));
        state.emit(out, ids.take(slot.kind, slot.value, scope));
        at = slot.end;
    }
    state.emit(out, xml.slice(at));
}

/** Calls `visit` with the id slots of every literal inside a loop (the slots markRepeatedIds gave them). */
function forEachLoopSlots(nodes: TemplateNode[], visit: (slots: IdSlot[]) => void): void {
    for (const node of nodes) {
        if (node.kind === 'literal') { if (node.slots) visit(node.slots); }
        else if (node.kind === 'if') node.sections.forEach(section => forEachLoopSlots(section, visit));
        else if (node.kind === 'for') forEachLoopSlots(node.body, visit);
    }
}

// ─── Render ─────────────────────────────────────────────────────────────────

/** A loop's own details, read as `loop.index` and so on. */
interface LoopDetails { index: number; number: number; first: boolean; last: boolean; length: number }

/** The names in scope: a loop's item and details, innermost first. */
interface Scope { name: string; item: TemplateValue; list: readonly TemplateValue[]; index: number; loop: LoopDetails; parent: Scope | null }

/** A value found in the data, with the object and key it was read from (for what a repeat adds). */
interface Found { value: TemplateValue; container: object; key: string }

/** Per-document render state. */
class RenderState {
    repeatLeft: number;
    readonly written: Uint8Array;
    readonly valueXml = new WeakMap<object, Map<string, string>>();
    readonly valueUsed = new WeakMap<object, Set<string>>();
    readonly ids: IdState | null;
    part = '';
    constructor(readonly data: TemplateData, readonly opts: DocxTemplateOptions, compiled: CompiledDocxTemplate) {
        this.repeatLeft = opts.maxRepeatedContent;
        this.written = new Uint8Array(compiled.literalCount);
        this.ids = compiled.ids ? new IdState(compiled.ids, compiled.rangesAcrossLoops) : null;
    }
    /** What the part being rendered has written so far. */
    partLength = 0;
    /** Writes to the part, refusing (typed) a part past what the runtime can hold in one string. */
    emit(out: string[], text: string): void {
        if (!text) return;
        if ((this.partLength += text.length) > MAX_PART_LENGTH) this.opts.onOutputLimit();
        out.push(text);
    }
    /** Counts what repetition adds, within `maxRepeatedContent`. */
    charge(amount: number): void {
        if (amount > 0 && (this.repeatLeft -= amount) < 0) this.opts.onRepeatLimit();
    }
}

/**
 * Reads a path: a loop's item or details when its first step names one (innermost loop first), else
 * the data, where a key holding the whole dotted text is read as typed before the path is followed.
 * Only own properties are read, so `__proto__`, `constructor` and `toString` are never found.
 */
function lookup(path: TemplatePath, text: string, scope: Scope | null, data: TemplateData): Found | typeof MISSING {
    let found: Found | null = null;
    for (let s = scope; s && !found; s = s.parent) {
        if (path[0] === s.name) found = { value: s.item, container: s.list, key: String(s.index) };
        else if (path[0] === LOOP_DETAILS_NAME) found = { value: s.loop as unknown as TemplateValue, container: s.loop, key: '' };
    }
    let step = 1;
    if (!found) {
        if (hasOwn(data, text)) return { value: data[text], container: data, key: text };
        found = { value: data as unknown as TemplateValue, container: data, key: '' };
        step = 0;
    }
    for (; step < path.length; step++) {
        const base: TemplateValue = found.value;
        if (!isContainer(base) || !hasOwn(base, path[step])) return MISSING;
        found = { value: (base as Record<string, TemplateValue>)[path[step]], container: base, key: path[step] };
    }
    return found;
}

function renderNodes(nodes: TemplateNode[], scope: Scope | null, state: RenderState, out: string[]): void {
    for (const node of nodes) {
        switch (node.kind) {
            case 'literal': renderLiteral(node, scope, state, out); break;
            case 'field': renderField(node, scope, state, out); break;
            case 'if': renderIf(node, scope, state, out); break;
            case 'for': renderFor(node, scope, state, out); break;
        }
    }
}

/** Writes a block as typed: made now, charged as a repeat when it was written before. */
function renderKept(span: KeptSpan, scope: Scope | null, state: RenderState, out: string[]): void {
    const xml = span.open + restoreMarkers(span.xml.slice(span.from, span.to), span.ctx, span.firstTag, span.lastTag) + span.close;
    // Free only the first time anything of the block is written: its content written as typed, or
    // rendered, is the template's own once. After that, charged by what it reads or writes, whichever is
    // more (its tags and fields are written back at full length; tags of other blocks it passes are read
    // and written as nothing), before writing it. Writing it marks all of the block written, so blocks
    // nested inside it are repeats from then on. Scanning the ids costs no more than the charge or, the
    // one time it is free, once per id.
    let repeat = false;
    for (let id = span.id; id <= span.lastId && !repeat; id++) repeat = state.written[id] === 1;
    if (repeat) state.charge(Math.max(xml.length, span.to - span.from));
    state.written.fill(1, span.id, span.lastId + 1);
    if (!state.ids) { state.emit(out, xml); return; }
    if (!span.inLoop) {
        // Written once, under the ids it was typed with; a note it refers to is written with it.
        const ids = state.ids;
        visitIds(xml, slot => ids.keep(slot.kind, slot.value));
        state.emit(out, xml);
        return;
    }
    writeSlots(xml, findIds(xml), scope, state, out);
}

function renderLiteral(node: LiteralNode, scope: Scope | null, state: RenderState, out: string[]): void {
    if (state.written[node.id]) state.charge(node.xml.length);
    else state.written[node.id] = 1;
    if (!node.slots || !state.ids) { state.emit(out, node.xml); return; }
    writeSlots(node.xml, node.slots, scope, state, out);
}

function renderField(node: FieldNode, scope: Scope | null, state: RenderState, out: string[]): void {
    state.charge(node.cost);
    const found = lookup(node.path, node.text, scope, state.data);
    if (found === MISSING) {
        if (state.opts.onMissing === 'error') state.opts.onFieldMissing(node.text);
        // 'keep': each piece of the tag stays in its own run, so a tag split across runs keeps each
        // run's formatting (written again, as in a loop, it is a repeat); 'empty': nothing.
        if (state.opts.onMissing === 'keep') {
            if (state.written[node.id]) state.charge(node.pieceText.length);
            else state.written[node.id] = 1;
            state.emit(out, node.pieceText);
        }
        return;
    }
    if (node.piece !== 0) return;
    const { container } = found;
    let { value, key } = found;
    // A filtered value is written per filter chain: the same value formatted two ways is two values.
    if (node.record.filters) key = `${key}\u0000${node.record.inner}`;
    // A formatter of the caller's own may answer differently each time it is asked (a counter, a lookup),
    // so its chain is run at each use; any other chain once per value, its XML kept: formatting reads the
    // whole value, and reading it again at every pass of a loop took time the budget did not see.
    const custom = node.record.filters?.some(filter => filter.kind === 'custom') ?? false;
    let cache = custom ? undefined : state.valueXml.get(container);
    if (!custom && !cache) state.valueXml.set(container, cache = new Map());
    let xml = cache?.get(key);
    if (xml === undefined) {
        if (node.record.filters) value = applyFilters(value, node.record.filters, node, state);
        // A list, an object, a function or a symbol has no text of its own (a function's would be its source).
        if (isContainer(value) || typeof value === 'function' || typeof value === 'symbol') {
            state.opts.onDataInvalid({ reason: 'not-text', tag: typedOf(node.record, state.opts), part: state.part, kind: kindOfValue(value) });
        }
    }
    // Each value's run-XML, made once; a value's first use is the caller's own data, and each later
    // use adds what it outgrows its placeholder by, within `maxRepeatedContent`, as a parse repeats a
    // document's shared content: 1,000,000 placeholders of one 100-character value (9.6 KB of template)
    // made 450 MB of XML, and a longer value failed untyped.
    // A value escaping or breaking into far more than its length is measured before it is built, so one
    // past what a part can hold is refused, typed, instead of failing in the runtime.
    const build = (): string => {
        const text = scalarText(value as string | number | boolean | Date | null | undefined);
        if (state.partLength + text.length + runXmlGrowth(text) > MAX_PART_LENGTH) state.opts.onOutputLimit();
        return textToRunXml(text);
    };
    if (xml === undefined) {
        xml = build();
        cache?.set(key, xml);
    }
    let used = state.valueUsed.get(container);
    if (!used) state.valueUsed.set(container, used = new Set());
    if (used.has(key)) state.charge(xml.length - node.record.length);
    else used.add(key);
    state.emit(out, xml);
}

/**
 * A value through a field's filters, in order. `format` writes a number with a number picture and a date
 * with a date picture (taking text that is a number, or a date in ISO 8601 form, as form data carries
 * them); `upper` and `lower` change case; the caller's own get the value as the data holds it. An empty
 * value (null or undefined) stays empty through the built-in filters.
 */
function applyFilters(value: TemplateValue, filters: CompiledFilter[], node: FieldNode, state: RenderState): TemplateValue {
    const unformattable = (filter: string, takes: string, v: TemplateValue): never => state.opts.onDataInvalid({
        reason: 'unformattable', tag: typedOf(node.record, state.opts), part: state.part, kind: kindOfValue(v), filter, takes,
    });
    let current = value;
    for (const filter of filters) {
        if (filter.kind === 'custom') { current = filter.run(current, ...filter.args); continue; }
        if (current === null || current === undefined) continue;
        if (filter.kind === 'format') {
            // Reading a number or a date from text reads the whole text.
            if (typeof current === 'string') state.charge(current.length);
            const { picture } = filter;
            if (picture.kind === 'number') {
                const n = numberOf(current);
                const written = n === null ? null : formatNumber(n, picture, state.opts.pictures);
                if (written === null) unformattable('format', 'number', current);
                current = written;
            } else {
                const date = dateOf(current);
                if (!date) unformattable('format', 'date', current);
                const written = formatDate(date as Date, picture, state.opts.pictures);
                if (written === null) unformattable('format', 'date', current);
                current = written;
            }
            continue;
        }
        // A list, an object, a function or a symbol has no text to change the case of.
        if (isContainer(current) || typeof current === 'function' || typeof current === 'symbol') unformattable(filter.kind, 'text', current);
        const text = scalarText(current as string | number | boolean | Date);
        // Changing case reads the whole text.
        state.charge(text.length);
        current = filter.kind === 'upper' ? text.toUpperCase() : text.toLowerCase();
    }
    return current;
}

/** A condition's or a list's value, with `onMissing` applied: MISSING only for 'keep'. */
function blockValue(path: TemplatePath, text: string, scope: Scope | null, state: RenderState): ConditionValue {
    const found = lookup(path, text, scope, state.data);
    if (found !== MISSING) return found.value;
    if (state.opts.onMissing === 'error') state.opts.onFieldMissing(text);
    return state.opts.onMissing === 'empty' ? null : MISSING;
}

function renderIf(node: IfNode, scope: Scope | null, state: RenderState, out: string[]): void {
    let branch = 0;
    const read = (path: TemplatePath, text: string): ConditionValue => blockValue(path, text, scope, state);
    // Comparing two strings reads them: as long as the shorter one.
    const compared = (length: number): void => state.charge(length);
    const incomparable = (why: Incomparable): never => state.opts.onDataInvalid({ reason: 'incomparable', tag: typedOf(node.branches[branch].record, state.opts), part: state.part, kind: why.left, otherKind: why.right });
    for (; branch < node.branches.length; branch++) {
        const { condition, cost } = node.branches[branch];
        state.charge(cost);
        const holds = condition === null ? true : evaluateCondition(condition, read, incomparable, compared);
        if (holds === MISSING) { renderKept(node.kept, scope, state, out); return; }
        if (holds) { renderNodes(node.sections[branch], scope, state, out); return; }
    }
}

function renderFor(node: ForNode, scope: Scope | null, state: RenderState, out: string[]): void {
    state.charge(node.cost);
    const list = blockValue(node.list, node.listText, scope, state);
    if (list === MISSING) { renderKept(node.kept, scope, state, out); return; }
    if (list === null || list === undefined) return;
    if (!Array.isArray(list)) state.opts.onDataInvalid({ reason: 'not-a-list', tag: typedOf(node.record, state.opts), part: state.part, kind: kindOfValue(list) });
    const items = list as readonly TemplateValue[];
    for (let index = 0; index < items.length; index++) {
        // Each pass costs at least one, so nested loops over long lists stay bounded even with nothing to write.
        state.charge(1);
        const loop: LoopDetails = { index, number: index + 1, first: index === 0, last: index === items.length - 1, length: items.length };
        renderNodes(node.body, { name: node.name, item: items[index], list: items, index, loop, parent: scope }, state, out);
    }
}

// ─── Structure repair ───────────────────────────────────────────────────────

/**
 * Keeps a rendered part as Word requires after loops and branches removed content: a table left with
 * no row is removed, a cell that no longer ends in a paragraph gets an empty one, and a text box,
 * footnote, endnote, comment, header, footer or block content control left with no paragraph gets an
 * empty one. A bookmark or comment range whose end went with the content (a loop of no passes, a
 * branch not taken) is closed where it starts. One forward scan; returns the part unchanged when
 * nothing needs it.
 */
function repairStructure(xml: string, limit: () => never): string {
    interface Open { name: string; start: number; rows: number; last: string | null; blocks: number; needsBlock: boolean; blockLevel: boolean; rowLevel: boolean }
    const edits: { from: number; to: number; insert: string }[] = [];
    const stack: Open[] = [];
    // The innermost table or row-level wrapper open: a row counts for it, and a wrapper's rows count for
    // what holds it when it closes.
    const holders: Open[] = [];
    const open = (name: string, start: number): Open => {
        const parent = stack[stack.length - 1];
        // A content control is block-level (holds paragraphs and tables) when it stands where they do:
        // in a block container, or in the content of a block-level control. One in a paragraph holds runs,
        // and one in a table holds rows.
        const inBlocks = !!parent && (BLOCK_HOLDERS.has(parent.name) || ((parent.name === W.sdtContent || parent.name === W.customXml) && parent.blockLevel));
        const blockLevel = !!parent && (
            ((name === W.sdt || name === W.customXml) && inBlocks)
            || (name === W.sdtContent && parent.name === W.sdt && parent.blockLevel));
        const needsBlock = NEEDS_BLOCK.has(name) || ((name === W.sdtContent || name === W.customXml) && blockLevel);
        // A content control (or custom XML) among a table's rows holds rows; left with none, it goes.
        const rowLevel = !!parent && (
            ((name === W.sdt || name === W.customXml) && (parent.name === W.tbl || ((parent.name === W.sdtContent || parent.name === W.customXml) && parent.rowLevel)))
            || (name === W.sdtContent && parent.name === W.sdt && parent.rowLevel));
        return { name, start, rows: 0, last: null, blocks: 0, needsBlock, blockLevel, rowLevel };
    };
    // `slash` is where a self-closing element's `/>` begins (-1 for one with an end tag): a paragraph it
    // needs goes inside it, `<w:tc/>` becoming `<w:tc><w:p/></w:tc>`.
    const close = (frame: Open, closeStart: number, end: number, slash: number): void => {
        const parent = stack[stack.length - 1];
        let removed = false;
        const addParagraph = (): void => {
            edits.push(slash >= 0 ? { from: slash, to: end, insert: `><w:p/></${frame.name}>` } : { from: closeStart, to: closeStart, insert: '<w:p/>' });
        };
        if ((frame.name === W.tbl || (frame.rowLevel && frame.name !== W.sdtContent)) && frame.rows === 0) { edits.push({ from: frame.start, to: end, insert: '' }); removed = true; }
        else if (frame.name === W.tc && frame.last !== W.p && frame.last !== W.sdt && frame.last !== W.customXml) addParagraph();
        else if (frame.needsBlock && frame.name !== W.tc && frame.blocks === 0) addParagraph();
        const isBlock = frame.name === W.p || frame.name === W.tbl || ((frame.name === W.sdt || frame.name === W.customXml) && frame.blockLevel);
        if (parent && isBlock && !removed) { parent.last = frame.name; parent.blocks++; }
    };
    // Ranges opened and not yet closed, by kind and id: where each start tag ends, to close it there.
    const RANGE_ENDS: Record<string, [string, string]> = { [W.bookmarkStart]: ['bookmark', W.bookmarkEnd], 'w:commentRangeStart': ['comment', 'w:commentRangeEnd'] };
    const openRanges = new Map<string, { at: number; end: string; id: string }>();
    let pos = 0;
    while (pos < xml.length) {
        const lt = xml.indexOf('<', pos);
        if (lt < 0) break;
        if (xml.startsWith('<!--', lt)) { const e = xml.indexOf('-->', lt + 4); if (e < 0) break; pos = e + 3; continue; }
        if (xml.startsWith('<![CDATA[', lt)) { const e = xml.indexOf(']]>', lt + 9); if (e < 0) break; pos = e + 3; continue; }
        if (xml.startsWith('<?', lt) || xml.startsWith('<!', lt)) { const e = xml.indexOf('>', lt + 2); if (e < 0) break; pos = e + 1; continue; }
        if (xml.charCodeAt(lt + 1) === 0x2f) {
            const gt = xml.indexOf('>', lt + 2);
            if (gt < 0) break;
            const frame = stack.pop();
            if (frame && (frame.name === W.tbl || frame.rowLevel)) {
                holders.pop();
                if (frame.rowLevel && frame.rows && holders.length) holders[holders.length - 1].rows += frame.rows;
            }
            if (frame) close(frame, lt, gt + 1, -1);
            pos = gt + 1;
            continue;
        }
        let nameEnd = lt + 1;
        while (nameEnd < xml.length && !NAME_END.test(xml[nameEnd])) nameEnd++;
        const gt = startTagEnd(xml, nameEnd);
        if (gt < 0) break;
        const frame = open(xml.slice(lt + 1, nameEnd), lt);
        if (frame.name === W.tr && holders.length) holders[holders.length - 1].rows++;
        const opens = RANGE_ENDS[frame.name];
        if (opens || frame.name === W.bookmarkEnd || frame.name === 'w:commentRangeEnd') {
            const id = attributeOf(xml.slice(lt, gt + 1), 'w:id');
            if (id !== null) {
                if (opens) openRanges.set(`${opens[0]}:${id}`, { at: gt + 1, end: opens[1], id });
                else openRanges.delete(`${frame.name === W.bookmarkEnd ? 'bookmark' : 'comment'}:${id}`);
            }
        }
        if (xml.charCodeAt(gt - 1) === 0x2f) close(frame, gt + 1, gt + 1, gt - 1);
        else {
            stack.push(frame);
            if (frame.name === W.tbl || frame.rowLevel) holders.push(frame);
        }
        pos = gt + 1;
    }
    for (const range of openRanges.values()) {
        // Closed right after its start, which Word writes empty; a start holding content (no Word file has
        // one) is left as it is.
        const closeAt = xml.charCodeAt(range.at - 2) === 0x2f ? range.at : -1;
        // Written in double quotes: an id read from single quotes may hold one.
        if (closeAt >= 0) edits.push({ from: closeAt, to: closeAt, insert: `<${range.end} w:id="${range.id.replace(/"/g, '&quot;')}"/>` });
    }
    if (!edits.length) return xml;
    edits.sort((x, y) => x.from - y.from || x.to - y.to);
    const out: string[] = [];
    let at = 0;
    let length = 0;
    for (const edit of edits) {
        if (edit.from < at) continue; // inside a stretch already removed
        out.push(xml.slice(at, edit.from), edit.insert);
        length += edit.from - at + edit.insert.length;
        at = edit.to;
    }
    if (length + xml.length - at > MAX_PART_LENGTH) limit();
    out.push(xml.slice(at));
    return out.join('');
}

/** Containers of paragraphs and tables, where a content control is block-level. */
const BLOCK_HOLDERS = new Set(['w:body', 'w:tc', 'w:txbxContent', 'w:footnote', 'w:endnote', 'w:comment', 'w:hdr', 'w:ftr', 'w:docPartBody']);

// ─── Render a document ──────────────────────────────────────────────────────

/** Renders one `.docx` from a compiled template and a single data map. */
export function renderDocxTemplate(compiled: CompiledDocxTemplate, data: TemplateData, opts: DocxTemplateOptions): Uint8Array {
    const state = new RenderState(data, opts, compiled);
    const rendered = new Map<string, string>();
    const renderPart = (name: string): void => {
        const part = compiled.parts.get(name);
        if (!part?.nodes) {
            const cut = compiled.notesPartXml.get(name);
            if (cut !== undefined) rendered.set(name, cut);
            return;
        }
        state.part = name;
        state.partLength = 0;
        const pieces: string[] = [];
        renderNodes(part.nodes, null, state, pieces);
        const xml = pieces.join('');
        rendered.set(name, part.hasBlocks ? repairStructure(xml, opts.onOutputLimit) : xml);
    };
    // The parts referring to notes first, then the notes, footnotes and endnotes before comments: a note
    // referred to in a loop is copied once the parts referring to it are rendered, and a footnote can
    // anchor a comment.
    const noteParts = Object.values(NOTE_PARTS).map(n => n.part) as string[];
    for (const name of compiled.parts.keys()) if (!noteParts.includes(name)) renderPart(name);
    for (const { part, element } of Object.values(NOTE_PARTS)) {
        renderPart(part);
        const copies = state.ids?.copies.filter(c => c.part === part) ?? [];
        if (!copies.length || !compiled.entries[part]) continue;
        state.part = part;
        rendered.set(part, copyNotes(rendered.get(part) ?? Buffer.from(compiled.entries[part]).toString('utf8'), part, element, copies, compiled, state));
    }
    const out: Zippable = {};
    for (const [name, bytes] of Object.entries(compiled.entries)) {
        const xml = rendered.get(name);
        if (xml !== undefined) {
            out[name] = [Buffer.from(xml, 'utf8'), { mtime: opts.mtime }];
        } else if (STORED_PART.test(name)) {
            // Already-compressed media (PNG/JPEG images, audio/video, WOFF fonts): deflate yields next
            // to nothing, so store them (`level: 0`) instead of paying to re-deflate the same bytes on
            // every document in a batch. Much less CPU, negligible size.
            out[name] = [bytes, { mtime: opts.mtime, level: 0 }];
        } else {
            // Everything else, including a text part with no tag (copied byte for byte, so a part that
            // is not valid UTF-8 is never corrupted by a decode/encode round-trip it did not need):
            // styles.xml, numbering.xml, settings.xml, theme/*.xml, the .rels parts deflate normally.
            out[name] = [bytes, { mtime: opts.mtime }];
        }
    }
    return zipSync(out);
}

/**
 * Unzips a docx under decompression limits and confirms it is one (has `word/document.xml`); returns
 * null when it is not a docx. Uses the parser's streaming `extractFiles` guard so a zip-bomb template
 * is bounded exactly as an ordinary parsed document is.
 */
export async function openDocx(bytes: Uint8Array, limits?: DecompressionLimits, config?: OfficeParserConfig): Promise<Record<string, Uint8Array> | null> {
    const src = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let files;
    try { files = await extractFiles(src, () => true, limits ?? {}, config); }
    catch (e) {
        // A decompression-limit breach or a truncated archive carries a typed officeIssue and was
        // already reported once; surface it so a zip-bomb or corrupt template is not silently
        // misreported as "not a .docx" (and not reported a second time). ZIP_NO_ENTRIES_FOUND and any
        // raw (non-typed) fflate error mean the input simply is not a zip -> not a docx -> null.
        const code = (e as { officeIssue?: { code?: OfficeErrorType } })?.officeIssue?.code;
        if (code && code !== OfficeErrorType.ZIP_NO_ENTRIES_FOUND) throw e;
        return null;
    }
    const map: Record<string, Uint8Array> = {};
    for (const f of files) map[f.path] = f.content;
    return map['word/document.xml'] ? map : null;
}

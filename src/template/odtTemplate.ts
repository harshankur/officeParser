/**
 * OpenDocument text templates (`.odt`, and `.ott`, written as `.odt`): the dialect the template engine
 * reads an ODT with (see {@link module:template/templateEngine}), the ids it keeps unique in copies, the
 * repair after loops and branches, and the package it writes.
 *
 * - Text is the character data of a paragraph (`text:p`, `text:h`) and the spans and links in it,
 *   with `<text:s/>` read as the spaces it stands for. A tag split across spans (as LibreOffice splits
 *   text at every change of attribute or edit) is one tag. Text in fields (`text:date`), in a note's
 *   citation, in a comment's author and date, and deleted text kept for tracked changes is not read.
 * - Tags are filled in `content.xml` and in `styles.xml`, whose master pages hold the headers and footers.
 *   Notes (`text:note`), comments (`office:annotation`) and text boxes (`draw:text-box`) live inside the
 *   paragraph that holds them, their paragraphs a content of their own.
 * - Blocks are inline, over paragraphs, over list items (a paragraph's list items below the block's
 *   container), and over table rows (in `table:table-rows` and row groups too, and inside the heading
 *   rows, `table:table-header-rows`, which a block never crosses).
 *
 * @module template/odtTemplate
 */

import { TemplateData } from '../types.js';
import { escapeXml, stripInvalidXmlChars } from '../utils/sanitize.js';
import {
    attributeOf, CompiledTemplate, Frame, IdAttribute, IdKind, IdSlot, inertEnd, markRuns, NAME_END, PartContext, renderParts,
    Scope, startTagEnd, TemplateDialect, TemplateIds, TemplateOptions, TextRun, visitIds, zipRendered,
} from './templateEngine.js';

/** The media types of an OpenDocument text and of its template. */
export const ODT_MEDIA_TYPE = 'application/vnd.oasis.opendocument.text';
export const OTT_MEDIA_TYPE = 'application/vnd.oasis.opendocument.text-template';

/** The parts that carry text a tag could be in: the body, and the master pages' headers and footers. */
const TEXT_PARTS = new Set(['content.xml', 'styles.xml']);

const PARAGRAPHS = new Set(['text:p', 'text:h']);

/** Elements whose own character data is a paragraph's text. */
const TEXT_HOLDERS = new Set(['text:p', 'text:h', 'text:span', 'text:a', 'text:meta']);

/** Elements whose text is never read: deleted text kept for tracked changes. */
const UNREAD = new Set(['text:tracked-changes']);

/** The element standing for spaces (`text:c` of them, one when it has none). */
const SPACE = 'text:s';

/** The most spaces one `<text:s/>` is read as: past it, it is read as none (no tag holds a run that long). */
const MAX_SPACE_RUN = 256;

/** List elements a paragraph of a paragraph block may sit in, below the block's container. */
const LIST_WRAPPERS = new Set(['text:list', 'text:list-item', 'text:list-header']);

/** The heading rows, a group of rows a block may lie wholly inside and never cross out of. */
const HEADER_ROWS = 'table:table-header-rows';

/** Groups of rows a row block widens to, as it widens to Word's content controls around rows. */
const ROW_WRAPPERS = new Set(['table:table-rows', 'table:table-row-group']);

/** Containers that must keep a paragraph: a cell, a header or footer, a note's body, a text box, a comment. */
const NEEDS_BLOCK = new Set([
    'table:table-cell', 'style:header', 'style:footer', 'style:header-left', 'style:footer-left', 'style:header-first', 'style:footer-first',
    'text:note-body', 'draw:text-box', 'office:annotation',
]);

/** The block-level children of a container (list items among a list's). */
const BLOCK_ELEMENTS = new Set([
    'text:p', 'text:h', 'table:table', 'text:list', 'text:list-item', 'text:list-header', 'text:section', 'text:numbered-paragraph',
    'text:table-of-content', 'text:illustration-index', 'text:table-index', 'text:object-index', 'text:user-index', 'text:alphabetical-index', 'text:bibliography',
]);

/** What a paragraph may hold and still count as holding nothing but tags: spans, spaces, page breaks LibreOffice marks. */
const TAG_ONLY_ALLOWED = new Set(['text:span', 'text:s', 'text:soft-page-break']);

/** Elements a block may be cut through, closed before the cut and opened again after it. */
const SPLITTABLE = new Set(['text:span', 'text:a', 'text:meta']);

/** The paragraphs of these start afresh: they are not inside the paragraph holding the element. */
const FRESH_CONTENT = new Set(['draw:text-box', 'text:note-body', 'office:annotation']);

/** A run of spaces as OpenDocument keeps it: `<text:s/>` for one, with its count for more. */
function spacesXml(count: number): string {
    return count === 1 ? '<text:s/>' : `<text:s text:c="${count}"/>`;
}

/** How a marker's piece of a `<text:s/>` is written back: as the spaces it stood for. */
const ODT_SPACES = { spacesXml: (spaces: string): string => spaces ? spacesXml(spaces.length) : '' };

/** Whether a part may hold a tag: it has the start delimiter's first character. */
function mayHoldTag(xml: string, ctx: PartContext): boolean {
    return xml.includes(ctx.startText[0]);
}

/**
 * Finds and marks the tags in an OpenDocument part: one forward scan with the elements open, collecting
 * each paragraph's runs of text (the character data of the paragraph and of its spans and links, and
 * each `<text:s/>`) between two paragraph-boundary tags, as for Word: a note's, a comment's or a text
 * box's paragraphs are their own, and no tag spans them. Every element and run is counted against the
 * template's bound on elements. Returns the part unchanged when it holds no tag.
 */
function markOdtPart(xml: string, ctx: PartContext): string {
    if (!mayHoldTag(xml, ctx)) return xml;
    const out: string[] = [];
    let copied = 0;
    const open: string[] = [];
    let unread = 0;
    let runs: TextRun[] = [];
    const counted = (): void => { if (--ctx.elements.textRuns < 0) ctx.opts.onElementLimit(); };
    const flush = (): void => {
        if (!runs.length) return;
        const pieces = markRuns(runs, ctx, ODT_SPACES);
        if (pieces) {
            runs.forEach((run, i) => {
                // A `<text:s/>` with a marker in it is written as its markers, and the spaces around them as `<text:s/>`.
                const text = run.spaces ? pieces[i].map(piece => piece.startsWith(ctx.open) ? piece : ODT_SPACES.spacesXml(piece)).join('') : pieces[i].join('');
                out.push(xml.slice(copied, run.start), text);
                copied = run.end;
            });
        }
        runs = [];
    };
    const inText = (): boolean => unread === 0 && TEXT_HOLDERS.has(open[open.length - 1]);
    let pos = 0;
    for (;;) {
        const lt = xml.indexOf('<', pos);
        const textEnd = lt < 0 ? xml.length : lt;
        if (textEnd > pos && inText()) {
            counted();
            runs.push({ start: pos, end: textEnd, inner: xml.slice(pos, textEnd) });
        }
        if (lt < 0) break;
        // A comment, CDATA section or processing instruction is passed whole: a tag in it is none.
        const inert = inertEnd(xml, lt);
        if (inert !== null) { if (inert < 0) break; pos = inert + 1; continue; }
        if (xml.charCodeAt(lt + 1) === 0x2f) {
            const gt = xml.indexOf('>', lt + 2);
            if (gt < 0) break;
            const name = xml.slice(lt + 2, gt).trim();
            if (PARAGRAPHS.has(name)) flush();
            // An end tag closing nothing open is passed over (the part is checked for balance where blocks are).
            if (open[open.length - 1] === name) {
                open.pop();
                if (UNREAD.has(name)) unread--;
            }
            pos = gt + 1;
            continue;
        }
        let nameEnd = lt + 1;
        while (nameEnd < xml.length && !NAME_END.test(xml[nameEnd])) nameEnd++;
        const gt = startTagEnd(xml, nameEnd);
        if (gt < 0) break;
        counted();
        const name = xml.slice(lt + 1, nameEnd);
        const selfClosing = xml.charCodeAt(gt - 1) === 0x2f;
        if (PARAGRAPHS.has(name)) flush();
        if (name === SPACE && selfClosing && inText()) {
            const count = spaceCount(xml.slice(lt, gt + 1));
            if (count <= MAX_SPACE_RUN) runs.push({ start: lt, end: gt + 1, inner: ' '.repeat(count), spaces: true });
        }
        if (!selfClosing) {
            open.push(name);
            if (UNREAD.has(name)) unread++;
        }
        pos = gt + 1;
    }
    flush();
    if (!out.length) return xml;
    out.push(xml.slice(copied));
    return out.join('');
}

/** How many spaces a `<text:s>` stands for: its `text:c`, one when it has none or one that is no count. */
function spaceCount(startTag: string): number {
    const given = attributeOf(startTag, 'text:c');
    if (given === null || !/^\d{1,6}$/.test(given.trim())) return 1;
    return Math.max(1, Number(given.trim()));
}

/** Line separators a value may hold, each written as a `<text:line-break/>`. */
const LINE_BREAK = /\r\n|[\r\n\u2028\u2029]/;

/**
 * A value's text as XML in a paragraph: escaped, XML-illegal characters stripped, each line break a
 * `<text:line-break/>`, each tab a `<text:tab/>`, and each space OpenDocument would fold into the one
 * before it (a run of spaces, or spaces where a line starts) a `<text:s/>`, so the value reads as typed.
 */
function odtTextXml(text: string): string {
    return writeOdtText(text, null);
}

/** How much longer a value's XML is than its text (see {@link odtTextXml}), counted without writing it. */
function odtTextGrowth(text: string): number {
    const counter = { growth: 0 };
    writeOdtText(text, counter);
    return counter.growth;
}

/** Writes a value's text as XML, or (given `counter`) only counts what the XML adds. One forward pass. */
function writeOdtText(text: string, counter: { growth: number } | null): string {
    const clean = stripInvalidXmlChars(text);
    const out: string[] = [];
    let lineStart = true;
    for (let at = 0; at < clean.length;) {
        const c = clean[at];
        if (c === ' ') {
            let end = at;
            while (end < clean.length && clean[end] === ' ') end++;
            const run = end - at;
            // The first space of a run is kept as text, save where a line starts; the rest are `<text:s/>`.
            const xml = lineStart ? spacesXml(run) : run === 1 ? ' ' : ` ${spacesXml(run - 1)}`;
            if (counter) counter.growth += xml.length - run; else out.push(xml);
            at = end;
            lineStart = false;
            continue;
        }
        if (c === '\t') {
            if (counter) counter.growth += '<text:tab/>'.length - 1; else out.push('<text:tab/>');
            at++;
            lineStart = false;
            continue;
        }
        const pair = c === '\r' && clean[at + 1] === '\n';
        if (pair || LINE_BREAK.test(c)) {
            if (counter) counter.growth += '<text:line-break/>'.length - (pair ? 2 : 1); else out.push('<text:line-break/>');
            at += pair ? 2 : 1;
            lineStart = true;
            continue;
        }
        // A stretch of other characters, escaped as a whole.
        let end = at + 1;
        while (end < clean.length && clean[end] !== ' ' && clean[end] !== '\t' && !LINE_BREAK.test(clean[end])) end++;
        const stretch = clean.slice(at, end);
        if (counter) counter.growth += escapeGrowth(stretch); else out.push(escapeXml(stretch));
        at = end;
        lineStart = false;
    }
    return out.join('');
}

/** What escaping adds to a text. */
function escapeGrowth(text: string): number {
    let growth = 0;
    for (let i = 0; i < text.length; i++) {
        switch (text.charCodeAt(i)) {
            case 0x26: growth += 4; break;                     // &amp;
            case 0x3c: case 0x3e: growth += 3; break;          // &lt; &gt;
            case 0x22: case 0x27: growth += 5; break;          // &quot; &apos;
        }
    }
    return growth;
}

// ─── Unique ids in repeated content ─────────────────────────────────────────

/** The elements whose start and end name a range, by the names their ids are unique among. */
const RANGE_STARTS: Record<string, [string, string]> = {
    'text:bookmark-start': ['text:name', 'bookmark'],
    'text:reference-mark-start': ['text:name', 'reference'],
    'office:annotation': ['office:name', 'annotation'],
    'text:toc-mark-start': ['text:id', 'index'],
    'text:alphabetical-index-mark-start': ['text:id', 'index'],
    'text:user-index-mark-start': ['text:id', 'index'],
};
const RANGE_ENDS: Record<string, [string, string]> = {
    'text:bookmark-end': ['text:name', 'bookmark'],
    'text:reference-mark-end': ['text:name', 'reference'],
    'office:annotation-end': ['office:name', 'annotation'],
    'text:toc-mark-end': ['text:id', 'index'],
    'text:alphabetical-index-mark-end': ['text:id', 'index'],
    'text:user-index-mark-end': ['text:id', 'index'],
};

/** Elements named by an attribute unique among others of their kind. */
const NAMED: Record<string, [string, string]> = {
    'text:bookmark': ['text:name', 'bookmark'],
    'text:reference-mark': ['text:name', 'reference'],
    'table:table': ['table:name', 'table'],
    'text:section': ['text:name', 'section'],
    'text:sequence': ['text:ref-name', 'sequence'],
};

/** The ids in an OpenDocument start tag, by its name (`xml:id`, which any element may carry, apart). */
function odtIdAttributes(name: string): IdAttribute[] {
    const start = RANGE_STARTS[name];
    if (start) return [[start[0], 'odfRangeStart', start[1]]];
    const end = RANGE_ENDS[name];
    if (end) return [[end[0], 'odfRangeEnd', end[1]]];
    const named = NAMED[name];
    if (named) return [[named[0], 'odfName', named[1]]];
    if (name === 'text:note') return [['text:id', 'odfNoteId', 'note']];
    // A drawing's name is unique among drawings, and its `draw:id` is an id, as `xml:id` is (the two the same where both are given).
    if (name.startsWith('draw:')) return [['draw:name', 'odfName', 'draw'], ['draw:id', 'odfId', 'id']];
    return [];
}

/** The range an OpenDocument slot is an end of (`bookmark:<name>`), or null. */
function odtRangeOf(slot: IdSlot): string | null {
    return slot.kind === 'odfRangeStart' || slot.kind === 'odfRangeEnd' ? `${slot.space}:${slot.value}` : null;
}

/** Every id and name the template holds, by the names it is unique among: fresh ones are made past them. */
type OdtSeeds = Map<string, Set<string>>;

function seedOdtIds(entries: Record<string, Uint8Array>, chargeFor: (part: string) => (() => void) | undefined): OdtSeeds {
    const seeds: OdtSeeds = new Map();
    for (const [name, bytes] of Object.entries(entries)) {
        if (!TEXT_PARTS.has(name)) continue;
        const charge = chargeFor(name);
        visitIds(Buffer.from(bytes).toString('utf8'), ODT_DIALECT, slot => {
            charge?.();
            const space = slot.space ?? 'id';
            let values = seeds.get(space);
            if (!values) seeds.set(space, values = new Set());
            values.add(slot.value);
        });
    }
    return seeds;
}

/**
 * Per render of an OpenDocument: which ids and names have been written, by the names each is unique
 * among. The first write keeps its value; each later one gets a fresh value, the typed one with a number
 * (`Image1_2`), past every value the template holds. A range's end follows its start; one with no start
 * before it gets a value of its own, matching nothing, and the repair removes it. The `xml:id` and
 * `draw:id` of one element stay one value.
 */
class OdtIds implements TemplateIds {
    private readonly seen = new Map<string, Set<string>>();
    private readonly made = new Map<string, Set<string>>();
    private readonly pending = new Map<string, string>();
    private across = new Set<string>();
    private number = 2;
    /** The last start tag an id was taken in, and what its ids were written as, so its `xml:id` and `draw:id` stay one. */
    private lastSlot: IdSlot | null = null;
    private readonly inTag = new Map<string, string>();
    constructor(private readonly seeds: OdtSeeds, private readonly rangesAcross: Set<string>) { }
    writtenAcross(range: string, kind: IdKind): boolean {
        if (!this.rangesAcross.has(range)) return false;
        const key = `${range}:${kind}`;
        if (this.across.has(key)) return true;
        this.across.add(key);
        return false;
    }
    take(kind: IdKind, value: string, _scope: Scope | null, slot: IdSlot): string {
        const space = slot.space ?? 'id';
        if (kind === 'odfRangeStart' || kind === 'odfRangeEnd') {
            const range = `${space}:${value}`;
            if (this.rangesAcross.has(range)) return value;
            if (kind === 'odfRangeStart') {
                const id = this.first(space, value) ? value : this.fresh(space, value);
                this.pending.set(range, id);
                return id;
            }
            const id = this.pending.get(range);
            this.pending.delete(range);
            return id ?? this.fresh(space, value);
        }
        if (kind === 'odfId') {
            const last = this.lastSlot;
            if (!last || last.tag !== slot.tag || last.gt !== slot.gt || last.start >= slot.start) this.inTag.clear();
            this.lastSlot = slot;
            const known = this.inTag.get(value);
            if (known !== undefined) return known;
            const id = this.first(space, value) ? value : this.fresh(space, value);
            this.inTag.set(value, id);
            return id;
        }
        return this.first(space, value) ? value : this.fresh(space, value);
    }
    /** Ids written as typed, in a block kept as typed outside any loop: copies made after them are fresh. */
    keep(kind: IdKind, value: string, slot: IdSlot): void {
        if (kind !== 'odfRangeEnd') this.first(slot.space ?? 'id', value);
    }
    /** Whether `value` is written for the first time among its names (and marks it written). */
    private first(space: string, value: string): boolean {
        let seen = this.seen.get(space);
        if (!seen) this.seen.set(space, seen = new Set());
        if (seen.has(value)) return false;
        seen.add(value);
        return true;
    }
    /** A value the template does not hold and this render has not made: the typed one with a number. */
    private fresh(space: string, value: string): string {
        let made = this.made.get(space);
        if (!made) this.made.set(space, made = new Set());
        const held = this.seeds.get(space);
        for (;;) {
            const candidate = `${value}_${this.number++}`;
            if (!held?.has(candidate) && !made.has(candidate) && !this.seen.get(space)?.has(candidate)) {
                made.add(candidate);
                return candidate;
            }
        }
    }
}

// ─── Structure repair ───────────────────────────────────────────────────────

/** Groups of rows, which hold rows and are left out when a loop or branch left them none. */
const ROW_HOLDERS = new Set(['table:table', HEADER_ROWS, 'table:table-rows', 'table:table-row-group']);

/**
 * Keeps a rendered OpenDocument part valid after loops and branches removed content: a table, or a group
 * of its rows, left with no row is removed; a cell, header, footer, note, text box or comment left with
 * no paragraph gets an empty one (an empty header also hides its left-page header in LibreOffice); a
 * bookmark, reference or index range whose end went with the content is closed after its start; and an
 * end whose start went is removed. One forward scan; the part is returned unchanged when nothing needs it.
 */
function repairOdt(xml: string, limit: () => never): string {
    interface Open { name: string; start: number; rows: number; blocks: number }
    const edits: { from: number; to: number; insert: string }[] = [];
    const stack: Open[] = [];
    const openRanges = new Map<string, { at: number; end: string; attribute: string; value: string }>();
    const close = (frame: Open, closeStart: number, end: number, slash: number): void => {
        const parent = stack[stack.length - 1];
        let removed = false;
        if (ROW_HOLDERS.has(frame.name) && frame.rows === 0) {
            edits.push({ from: frame.start, to: end, insert: '' });
            removed = true;
        } else if (NEEDS_BLOCK.has(frame.name) && frame.blocks === 0) {
            edits.push(slash >= 0 ? { from: slash, to: end, insert: `><text:p/></${frame.name}>` } : { from: closeStart, to: closeStart, insert: '<text:p/>' });
        }
        if (!parent || removed) return;
        if (BLOCK_ELEMENTS.has(frame.name)) parent.blocks++;
        // A group's rows count for what holds it.
        if (frame.name === 'table:table-row') parent.rows++;
        else if (ROW_HOLDERS.has(frame.name) && frame.name !== 'table:table') parent.rows += frame.rows;
    };
    let pos = 0;
    while (pos < xml.length) {
        const lt = xml.indexOf('<', pos);
        if (lt < 0) break;
        const inert = inertEnd(xml, lt);
        if (inert !== null) { if (inert < 0) break; pos = inert + 1; continue; }
        if (xml.charCodeAt(lt + 1) === 0x2f) {
            const gt = xml.indexOf('>', lt + 2);
            if (gt < 0) break;
            const frame = stack.pop();
            if (frame) close(frame, lt, gt + 1, -1);
            pos = gt + 1;
            continue;
        }
        let nameEnd = lt + 1;
        while (nameEnd < xml.length && !NAME_END.test(xml[nameEnd])) nameEnd++;
        const gt = startTagEnd(xml, nameEnd);
        if (gt < 0) break;
        const name = xml.slice(lt + 1, nameEnd);
        const selfClosing = xml.charCodeAt(gt - 1) === 0x2f;
        const start = RANGE_STARTS[name];
        const end = RANGE_ENDS[name];
        if (start || end) {
            const [attribute, space] = (start ?? end)!;
            const value = attributeOf(xml.slice(lt, gt + 1), attribute);
            if (value !== null) {
                const key = `${space}:${value}`;
                if (start) {
                    // A comment is a range only with its end; one left without it is a comment on a point, as typed.
                    if (name !== 'office:annotation') openRanges.set(key, { at: selfClosing ? gt + 1 : -1, end: name.replace(/-start$/, '-end'), attribute, value });
                    else openRanges.set(key, { at: -1, end: '', attribute, value });
                } else if (openRanges.has(key)) {
                    openRanges.delete(key);
                } else if (selfClosing) {
                    // An end whose start went with a branch or a loop of no passes closes nothing.
                    edits.push({ from: lt, to: gt + 1, insert: '' });
                }
            }
        }
        const frame: Open = { name, start: lt, rows: 0, blocks: 0 };
        if (selfClosing) close(frame, gt + 1, gt + 1, gt - 1);
        else stack.push(frame);
        pos = gt + 1;
    }
    for (const range of openRanges.values()) {
        if (range.at >= 0) edits.push({ from: range.at, to: range.at, insert: `<${range.end} ${range.attribute}="${range.value}"/>` });
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

/** The longest string the runtime can make (V8's limit), as the engine holds a part to. */
const MAX_PART_LENGTH = 2 ** 29 - 24;

/** What a crossing is in an OpenDocument, for its error. */
function describeOdtCrossing(units: Frame[], container: Frame): string {
    if (container.para) return 'one is inside a text box, a note or a comment and another is not';
    if (units.some(u => u.name === HEADER_ROWS)) return 'one is in a table\'s heading rows and another is not';
    if (units.some(u => u.name === 'table:table')) return 'one is inside a table and another is outside it, or in another table';
    if (units.some(u => LIST_WRAPPERS.has(u.name))) return 'one is inside a list and another is not, or they are at different levels of a list';
    if (units.some(u => u.name === 'text:section')) return 'one is inside a section and another is not';
    return 'they are in places no single stretch of content joins';
}

/** OpenDocument's names and rules (see {@link TemplateDialect}). */
export const ODT_DIALECT: TemplateDialect = {
    isTextPart: name => TEXT_PARTS.has(name),
    markPart: markOdtPart,
    paragraphs: PARAGRAPHS,
    textHolders: TEXT_HOLDERS,
    freshContent: FRESH_CONTENT,
    table: 'table:table',
    row: 'table:table-row',
    rowWrappers: ROW_WRAPPERS,
    rowGroups: new Set([HEADER_ROWS]),
    paragraphWrappers: LIST_WRAPPERS,
    tagOnlyAllowed: TAG_ONLY_ALLOWED,
    splittable: SPLITTABLE,
    liftedFromInline: new Set(),
    needsBlock: NEEDS_BLOCK,
    blockElements: BLOCK_ELEMENTS,
    // A tracked change is marked by points (`text:change-start`), which a cut does not pass through.
    trackedChanges: new Set(),
    propertyElements: new Set(),
    runProps: null,
    fieldChar: null,
    sectionProps: null,
    describeCrossing: describeOdtCrossing,
    textXml: odtTextXml,
    textGrowth: odtTextGrowth,
    idAttributes: odtIdAttributes,
    everyElementIds: [['xml:id', 'odfId', 'id']],
    rangeOf: odtRangeOf,
    seedIds: seedOdtIds,
    newIds: (seeds, rangesAcross) => new OdtIds(seeds as OdtSeeds, rangesAcross),
    repair: repairOdt,
    noteParts: null,
};

// ─── The package ────────────────────────────────────────────────────────────

const MANIFEST = 'META-INF/manifest.xml';

/**
 * An OpenDocument text template's entries, or null when the package is not one (its `mimetype`,
 * or else its manifest's root, names a text or a text template, and it has `content.xml`). A template
 * (`.ott`) is made a document (`.odt`): its `mimetype` and its manifest's root media type are switched.
 * A template that was decrypted (`decrypted`) loses its manifest's encryption data, as the document
 * written is plain. Nothing else changes.
 */
export function odtEntries(entries: Record<string, Uint8Array>, decrypted: boolean): Record<string, Uint8Array> | null {
    if (!entries['content.xml']) return null;
    const manifest = entries[MANIFEST] ? Buffer.from(entries[MANIFEST]).toString('utf8') : null;
    const mimetype = entries['mimetype'] ? Buffer.from(entries['mimetype']).toString('utf8').trim() : manifest ? rootMediaType(manifest) : null;
    if (mimetype !== ODT_MEDIA_TYPE && mimetype !== OTT_MEDIA_TYPE) return null;
    if (mimetype === OTT_MEDIA_TYPE && entries['mimetype']) entries['mimetype'] = Buffer.from(ODT_MEDIA_TYPE, 'utf8');
    if (manifest !== null) {
        let written = mimetype === OTT_MEDIA_TYPE ? withRootMediaType(manifest, ODT_MEDIA_TYPE) : manifest;
        if (decrypted) written = withoutEncryptionData(written);
        if (written !== manifest) entries[MANIFEST] = Buffer.from(written, 'utf8');
    }
    return entries;
}

/** The manifest's file entry for the package's root (`manifest:full-path="/"`): where its start tag starts and ends. */
function rootEntry(manifest: string): [number, number] | null {
    const open = '<manifest:file-entry';
    for (let at = manifest.indexOf(open); at >= 0; at = manifest.indexOf(open, at + 1)) {
        const gt = startTagEnd(manifest, at + open.length);
        if (gt < 0) return null;
        if (attributeOf(manifest.slice(at, gt + 1), 'manifest:full-path') === '/') return [at, gt + 1];
        at = gt;
    }
    return null;
}

function rootMediaType(manifest: string): string | null {
    const entry = rootEntry(manifest);
    return entry ? attributeOf(manifest.slice(entry[0], entry[1]), 'manifest:media-type') : null;
}

/** The manifest with its root's media type switched. */
function withRootMediaType(manifest: string, mediaType: string): string {
    const entry = rootEntry(manifest);
    if (!entry) return manifest;
    const tag = manifest.slice(entry[0], entry[1]);
    const switched = tag.replace(/(manifest:media-type\s*=\s*)(["'])[^"']*\2/, `$1"${mediaType}"`);
    return manifest.slice(0, entry[0]) + switched + manifest.slice(entry[1]);
}

/** The manifest without its `manifest:encryption-data` elements. One forward pass. */
function withoutEncryptionData(manifest: string): string {
    const open = '<manifest:encryption-data';
    const close = '</manifest:encryption-data>';
    const out: string[] = [];
    let at = 0;
    for (let start = manifest.indexOf(open); start >= 0; start = manifest.indexOf(open, at)) {
        const gt = startTagEnd(manifest, start + open.length);
        if (gt < 0) break;
        const end = manifest.charCodeAt(gt - 1) === 0x2f ? gt + 1 : manifest.indexOf(close, gt);
        if (end < 0) break;
        out.push(manifest.slice(at, start));
        at = manifest.charCodeAt(gt - 1) === 0x2f ? end : end + close.length;
    }
    if (!out.length) return manifest;
    out.push(manifest.slice(at));
    return out.join('');
}

/** Renders one `.odt` from a compiled template and a single data map: `mimetype` first and stored, as OpenDocument requires. */
export function renderOdtTemplate(compiled: CompiledTemplate, data: TemplateData, opts: TemplateOptions): Uint8Array {
    return zipRendered(compiled, renderParts(compiled, data, opts), opts.mtime, 'mimetype');
}

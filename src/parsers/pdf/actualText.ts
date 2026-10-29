/**
 * The `/ActualText` of a PDF's marked content: the text a span of glyphs stands for, as a producer
 * writes it where the glyphs alone do not say it (a ligature, a conjunct of an Indic script, a word in
 * glyphs that map to nothing). Chromium and LibreOffice wrap every such cluster in
 * `/Span << /ActualText (...) >> BDC ... EMC`, and map its glyphs to nothing, or all to the first: read
 * from the glyphs, "हिन्दी" came out as "ह ी". pdf.js keeps only a marked-content item's tag and MCID,
 * so the page's content streams are read here, for their marked-content operators alone, in the order
 * pdf.js meets them (forms drawn with `Do` in place): the n-th begin here is the n-th begin pdf.js gives.
 *
 * The streams are read by the small reader of the file in pdfFile: past its bounds, or in a file it
 * cannot read, a page has no events and its glyphs' text stands.
 *
 * @module parsers/pdf/actualText
 */

import { ContentOp, contentStreams, formOf, MAX_FORM_DEPTH, PdfDict, PdfFile, PdfName, PdfReader, PdfString, PdfValue, textOf } from './pdfFile.js';
import { measureItems, PageGlyphs } from './textGlyphs.js';

/** One marked-content event: a begin (with its ActualText, where it has one) or an end. */
export type MarkedEvent = { begin: true; actualText?: string } | { begin: false };

/** At most this many events a page, and for the whole document. */
const MAX_EVENTS = 500_000;
const MAX_DOCUMENT_EVENTS = 10_000_000;

/**
 * The marked-content events of a document's pages, read from the file when a page asks for them (see
 * eventsOf): the structure, each stream's operators and each form's events read once, every read within
 * the document's bounds (bytes inflated, tokens, events), past which a page has none and its glyphs'
 * text stands.
 */
export class MarkedContent {
    /** A form's events, with the resources it was last drawn with (pages of resources of their own each keep none). */
    private readonly forms = new Map<PdfValue, { resources: PdfDict | undefined; events: MarkedEvent[] }>();
    private eventsLeft = MAX_DOCUMENT_EVENTS;

    constructor(private readonly file: PdfFile) {}

    /**
     * A page's events (index 0 the first page's), in the order pdf.js meets them; undefined for a file
     * whose structure cannot be read here (encrypted, damaged, in a filter not read), or once the
     * document's bounds are spent. A page whose content cannot be read has none.
     */
    eventsOf(index: number): MarkedEvent[] | undefined {
        try {
            const at = this.file.page(index);
            if (!at || this.eventsLeft <= 0) return undefined;
            const events = this.pageEvents(at.reader, at.page);
            this.eventsLeft -= events.length;
            return events;
        } catch {
            return [];
        }
    }

    /** A page's events: its content streams' operators with its resources, forms drawn in place. */
    private pageEvents(reader: PdfReader, page: { dict: PdfDict; resources: PdfDict | undefined }): MarkedEvent[] {
        /** The events of operators drawn with `resources` (EMC included even at no depth, see below). */
        const play = (ops: ContentOp[], resources: PdfDict | undefined, depth: number, active: Set<PdfValue>, events: MarkedEvent[]) => {
            for (const o of ops) {
                if (events.length >= MAX_EVENTS) return;
                if (o.op === 'BMC') events.push({ begin: true });
                else if (o.op === 'EMC') events.push({ begin: false });
                else if (o.op === 'BDC') {
                    let props = o.props;
                    if (props instanceof PdfName) {
                        const named = resources ? reader.resolve(resources.get('Properties')) : undefined;
                        props = named instanceof Map ? reader.resolve(named.get(props.name)) : undefined;
                    }
                    const actual = props instanceof Map ? reader.resolve(props.get('ActualText')) : undefined;
                    events.push(actual instanceof PdfString ? { begin: true, actualText: textOf(actual) } : { begin: true });
                } else if (o.op === 'Do' && depth < MAX_FORM_DEPTH) {
                    const drawnForm = formOf(reader, resources, o.name);
                    if (!drawnForm || active.has(drawnForm.form)) continue;
                    const { form, resources: res } = drawnForm;
                    // A form's events are read once for the resources it is drawn with, and repeated where it is
                    // drawn again: a form drawn ten times in each of eight forms nested would be read a hundred
                    // million times.
                    const cached = this.forms.get(form);
                    let drawn = cached && cached.resources === res ? cached.events : undefined;
                    if (!drawn) {
                        drawn = [];
                        const body = this.file.program(reader, form);
                        active.add(form);
                        if (body) play(body, res, depth + 1, active, drawn);
                        active.delete(form);
                        this.forms.set(form, { resources: res, events: drawn });
                    }
                    for (let k = 0; k < drawn.length && events.length < MAX_EVENTS; k++) events.push(drawn[k]);
                }
            }
        };
        const raw: MarkedEvent[] = [];
        // The streams are one content stream, as a page's are.
        for (const st of contentStreams(reader, page.dict)) {
            const ops = this.file.program(reader, st);
            if (!ops) return [];
            play(ops, page.resources, 0, new Set(), raw);
        }
        // An end with nothing open is dropped, as pdf.js drops it.
        const events: MarkedEvent[] = [];
        let level = 0;
        for (const e of raw) {
            if (e.begin) { level++; events.push(e); }
            else if (level > 0) { level--; events.push(e); }
        }
        return events;
    }
}

/** A pdf.js text content item, as far as reading it needs (`measured` and `invisible`: see measureItems). */
interface TextContentItem { type?: string; str?: string; transform?: number[]; width?: number; measured?: boolean; invisible?: boolean }

/**
 * A page's text items as they read: each span with an ActualText (see readMarkedContent) as one item of
 * that text over the span's width (its glyphs' own text, in it, left out), and the control characters a
 * font's lost mapping gives (U+0002, U+0005) taken out of every item. The events are used only where
 * they are as many as pdf.js's marked-content begins (else the page is as pdf.js reads it). With
 * `normalize` (pdfParserConfig.normalizeText), an ActualText's marks are composed with their letters, as
 * pdf.js composes the glyphs' text: "e" and a combining acute is "é". With the page's `glyphs` (see
 * textGlyphs), the items are first placed where their glyphs are drawn (see measureItems).
 */
export function readableItems<T extends TextContentItem>(items: T[], events: MarkedEvent[] | undefined, normalize = false, glyphs?: PageGlyphs): T[] {
    return withoutCoveredSpaces(advanced(spans(measureItems(items, glyphs), events, normalize)));
}

/**
 * Items of whitespace alone drawn inside a glyph of their line (a space pdf.js put where a glyph's width
 * was lost, over the glyphs after it) left out: no gap is there. A space lies inside a glyph when most of
 * it is in the glyph's box, or it starts before a glyph before it ends; a space whose end runs a little
 * into the next word (as one a producer sets between words of right-to-left text, or at a font's change)
 * is a word's space and stays. Looked for among the few items drawn around each.
 */
function withoutCoveredSpaces<T extends TextContentItem>(items: T[]): T[] {
    const level = (it: T) => it.transform && it.transform.length >= 6 && Math.abs(it.transform[1]) < 1e-6 && Math.abs(it.transform[2]) < 1e-6
        && it.transform.every(Number.isFinite);
    return items.filter((item, i) => {
        if (typeof item.str !== 'string' || item.str.trim() || !item.str || !item.width || !Number.isFinite(item.width) || !level(item)) return true;
        const x0 = item.transform![4], x1 = x0 + item.width, y = item.transform![5];
        for (let k = Math.max(0, i - 4); k < Math.min(items.length, i + 5); k++) {
            const o = items[k];
            if (k === i || typeof o.str !== 'string' || !o.str.trim() || !level(o) || !(Math.abs(o.transform![5] - y) <= 0.5)) continue;
            const ox0 = o.transform![4], ox1 = ox0 + (o.width || 0);
            const overlap = Math.min(x1, ox1) - Math.max(x0, ox0);
            // (One reaching a fifth of an em past the glyph is that glyph's lost width and a word space after it.)
            const em = Math.abs(item.transform![0]) || Math.abs(o.transform![0]) || 1;
            if ((overlap > 0.5 * (x1 - x0) || (ox0 < x0 - 0.5 && ox1 > x0 + 0.5)) && !(x1 - ox1 >= 0.2 * em)) return false;
        }
        return true;
    });
}

/**
 * The user-perceived characters in a text: its grapheme clusters (or its letters where they cannot be
 * told), a conjunct joined by a virama counting once ("க்ஷ", "ஸ்ரீ", "क्ष"): counted apart, a conjunct's lost
 * width was taken for twice its own, over the space after it.
 */
function clusters(text: string): number {
    const Segmenter = (Intl as unknown as { Segmenter?: new (locale?: string, options?: object) => { segment(t: string): Iterable<{ segment: string }> } }).Segmenter;
    if (!Segmenter) return Math.max(1, (text.match(/\p{L}/gu) ?? []).length);
    let n = 0, joined = false;
    for (const { segment } of new Segmenter(undefined, { granularity: 'grapheme' }).segment(text)) {
        if (!joined) n++;
        joined = VIRAMA.test(segment[segment.length - 1] ?? '');
    }
    return Math.max(1, n);
}

/** The viramas (halants) of the Indic scripts, which join the consonant after them into one conjunct. */
const VIRAMA = /[\u094D\u09CD\u0A4D\u0ACD\u0B4D\u0BCD\u0C4D\u0CCD\u0D4D\u0DCA\u1039\u17D2]/;

/**
 * Items pdf.js gave no width (a glyph whose text holds a virama or another mark it takes for a
 * diacritic, as an Indic conjunct's does: its own width lost, and the text after it on its line drawn
 * that much too far left, so read out of order) given one (nine tenths of an em a cluster), and the items after
 * them on their line moved by it, in the order they are drawn. A mark alone keeps none. The space
 * pdf.js puts before the text after such an item (the gap its lost width left) is no space: "क्ष त्रिय"
 * for "क्षत्रिय". An item measured from its glyphs (see measureItems) is where it is drawn, and is left as it is.
 */
function advanced<T extends TextContentItem>(items: T[]): T[] {
    // (The line's baseline: none before its first item, so the shift starts afresh on every line, the first too.)
    // (`lost`: the item without its width last given one, until the next item on its line shows whether pdf.js
    // already left the room: one drawn at least a third of an em past it was, and is not moved.)
    let shift = 0, lineY: number | undefined, afterLost = false, lost: { x: number; width: number; em: number } | undefined;
    return items.map(item => {
        const t = item.transform;
        if (typeof item.str !== 'string' || !t || t.length < 6 || !t.every(Number.isFinite) || Math.abs(t[1]) > 1e-6 || Math.abs(t[2]) > 1e-6 || item.measured) return item;
        if (lineY === undefined || !(Math.abs(t[5] - lineY) <= 0.5)) { shift = 0; lineY = t[5]; afterLost = false; lost = undefined; }
        else if (lost && item.str.trim()) {
            // (A blank decides nothing: pdf.js puts one over the lost glyph, as wide as it and the word space after it.)
            // (Room left: the space before this item, if any, is the word's, not the lost width's gap.)
            if (t[4] >= lost.x + lost.em / 3) { shift -= lost.width; afterLost = false; }
            lost = undefined;
        }
        let moved = shift ? { ...item, transform: [t[0], t[1], t[2], t[3], t[4] + shift, t[5]] } : item;
        if (afterLost && /^\s/.test(item.str) && item.str.trim()) moved = { ...moved, str: item.str.replace(/^\s+/, '') };
        if (item.str.trim()) afterLost = false;
        if (!item.width && /\p{L}/u.test(item.str) && /[\p{Mn}\p{Mc}]/u.test(item.str)) {
            afterLost = true;
            // Nine tenths of an em a cluster (more than less: text drawn over its neighbour still reads in
            // order, where a gap left would read as a space, or a column's edge).
            const width = 0.9 * Math.abs(t[0]) * clusters(item.str.trim());
            shift += width;
            lost = { x: t[4], width, em: Math.abs(t[0]) || 1 };
            return { ...moved, width };
        }
        return moved;
    });
}

function spans<T extends TextContentItem>(items: T[], events: MarkedEvent[] | undefined, normalize: boolean): T[] {
    const controls = /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g;
    const clean = (item: T): T => typeof item.str === 'string' && controls.test(item.str) ? { ...item, str: item.str.replace(controls, '') } : item;
    const isBegin = (item: T) => item.type === 'beginMarkedContent' || item.type === 'beginMarkedContentProps';
    const begins = events?.filter(e => e.begin) as ({ begin: true; actualText?: string })[] | undefined;
    if (!begins?.some(e => e.actualText !== undefined) || items.filter(isBegin).length !== begins.length) return items.map(clean);
    const out: T[] = [];
    let at = 0, depth = 0, text = '';
    let target: T | undefined, minX = Infinity, maxX = -Infinity, level = true;
    // Whether an item of the span was measured from its glyphs, and how many of its glyph items there are and are invisible.
    let measured = false, glyphItems = 0, invisibleItems = 0;
    // The span's text with the space before it (where pdf.js put one among its items, before its glyphs):
    // its text, which way it reads, the first glyph's place, and whether it has that space yet.
    let said = '', rtl = false, glyphX: number | undefined, spaced = false;
    const finish = () => {
        if (target && level && Number.isFinite(minX)) {
            target.transform = [...target.transform!];
            target.transform[4] = minX;
            target.width = maxX - minX;
        }
        if (target) {
            if (measured) target.measured = true;
            if (glyphItems && invisibleItems === glyphItems) target.invisible = true;
            else delete target.invisible;
        }
        target = undefined;
    };
    for (const item of items) {
        if (isBegin(item)) {
            const event = begins[at++];
            out.push(item);
            if (depth) depth++;
            else if (event?.actualText !== undefined) {
                depth = 1; text = event.actualText; target = undefined; minX = Infinity; maxX = -Infinity; level = true;
                measured = false; glyphItems = 0; invisibleItems = 0;
            }
            continue;
        }
        if (item.type === 'endMarkedContent') {
            if (depth && --depth === 0) finish();
            out.push(item);
            continue;
        }
        if (!depth) { out.push(clean(item)); continue; }
        // A span of no text (a halo drawn under its label, the label drawn after it): nothing of it is text, its
        // blanks neither. As the span's text over its glyphs' width, a blank before them lay over the label's own
        // glyphs and was dropped with the label's word space.
        if (!text.replace(controls, '')) continue;
        // In a span with an ActualText: the first glyph's item holds it, over every glyph's width.
        const t = item.transform;
        if (t && t.length >= 6) {
            if (Math.abs(t[1]) > 1e-6 || Math.abs(t[2]) > 1e-6) level = false;
            minX = Math.min(minX, t[4]);
            maxX = Math.max(maxX, t[4] + (item.width || 0));
        }
        const blank = !(item.str ?? '').trim();
        if (item.measured) measured = true;
        if (!blank) { glyphItems++; if (item.invisible) invisibleItems++; }
        if (!target) {
            // A space pdf.js put before the glyphs (a ligature's own width lost) stays on that side of the text:
            // before it, or after it for right-to-left text, which is read from the other side.
            spaced = /^\s/.test(item.str ?? '');
            said = text.replace(controls, '');
            if (normalize) said = said.normalize('NFC');
            rtl = /[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/.test(said);
            glyphX = blank || !t ? undefined : t[4];
            target = { ...item, str: spaced ? (rtl ? said + ' ' : ' ' + said) : said };
            out.push(target);
        } else if (!spaced && blank && /\s/.test(item.str ?? '') && t && t.length >= 6 && glyphX !== undefined && t[4] + (item.width || 0) <= glyphX + 0.5) {
            // One pdf.js put after a glyph whose width it lost (a conjunct: its text holds a virama), but drawn
            // before the span's glyphs: the space before the word, which was lost with the span's own blanks.
            spaced = true;
            target.str = rtl ? said + ' ' : ' ' + said;
        } else if (!blank && glyphX === undefined && t && t.length >= 6) glyphX = t[4];
    }
    finish();
    return out;
}

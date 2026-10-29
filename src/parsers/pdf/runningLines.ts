/**
 * A PDF's running headers and footers where it does not mark them (an untagged PDF, or one whose tags are
 * not trusted): text repeated in one place at the top or the bottom of its pages, found by comparing the
 * pages (see findRunningLines). And the page numbers a page's header or footer prints, marked as the page
 * fields they are (see markPageFields), so a writer that lays out pages numbers each of its own.
 */
import type { OfficeContentNode, PageNumberFormat, TextMetadata } from '../../types.js';
import { formatListNumber } from '../../utils/listMarkers.js';
import { rotateBoundsToRendered } from './geometry.js';
import type { PageExtract, PdfLayoutConfig, RawRun } from './pdfTypes.js';
import { buildLines, lineRuns } from './textLayout.js';

/**
 * The share of a page's height, at its top and at its bottom, a running header or footer is read in, as
 * the page is shown (a tagged PDF's artifacts too, see PdfParser's splitArtifacts).
 */
export const FURNITURE_BAND = 0.15;

/** Lines this close to the page's edge (points) are in one place, as are lines whose left edges are this close. */
const EDGE_TOLERANCE = 2;
const LEFT_TOLERANCE = 4;
/** A line is running when it is on at least this share of the pages read, and on at least this many. */
const RUNNING_SHARE = 0.4;
const RUNNING_PAGES = 3;
/** A page's own number is running on this many pages: a two page document numbers both. */
const PAGE_NUMBER_PAGES = 2;
/** A running line is no larger than the body's text by more than this (points). */
const SIZE_SLACK = 0.5;
/**
 * A running header stands apart from the text below it (a running footer from the text above) by at least
 * this many of its own heights: one closer reads with that text, as a table's head row, repeated at the top
 * of each page the table goes on to, reads with the rows under it.
 */
const GAP_HEIGHTS = 1;
/** As much of a line's text as is compared (with its length): a line of a megabyte is not kept twice. */
const KEY_TEXT = 120;
/**
 * A line whose numbers are not the page's (a heading's "Question 3", a chapter's "Chapter 3") is running only
 * where its text, as printed, is on at least this many of the pages of its key, or where most of its key's
 * pages are so: a chapter's head repeated on its pages is, and a short chapter's on its one page among them;
 * a question numbered once on each page is not.
 */
const SAME_TEXT_PAGES = 2;
/** The longest line a page number is looked for in (characters). */
const NUMBERED_TEXT = 200;

/** The words a page number is printed with ("Page 3", "Seite 3 von 4"), lower case. */
const PAGE_WORD = String.raw`(?:page|seite|pagina|página|pág\.?|pag\.?|p\.|s\.|blatt|side|sida|sivu|strona|str\.|стр\.?|страница)`;
const OF_WORD = String.raw`(?:/|of|von|de|sur|di|van|af|av|z|из)`;
/** Dashes around a page number ("- 3 -"). */
const DASH = String.raw`[-\u2010-\u2015\u2212]`;
/** A line that is a page number and nothing else: "3", "- 3 -", "3/4", "Page 3 of 4", "Seite 3 von 4", "iii". */
const PAGE_LINE = new RegExp(String.raw`^${DASH}?\s*(?:${PAGE_WORD}\s*)?(\d+|[ivxlcdm]+)(?:\s*${OF_WORD}\s*(\d+))?\s*${DASH}?$`, 'iu');
/** In a header's or footer's text: "3/4", "3 of 4", "Seite 3 von 4" (the page's number, then the pages'). */
const NUMBER_OF = new RegExp(String.raw`(?<![\p{L}\p{N}])(\d+)\s*${OF_WORD}\s*(\d+)(?![\p{L}\p{N}])`, 'giu');
/** "Page 3", "Seite 3", "Page iii". */
const WORD_NUMBER = new RegExp(String.raw`(?<![\p{L}\p{N}])${PAGE_WORD}\s*(\d+|[ivxlcdm]+)(?![\p{L}\p{N}])`, 'giu');
/** A header or footer that is its number alone: "3", "- 3 -", "iii". */
const BARE_NUMBER = new RegExp(String.raw`^\s*${DASH}?\s*(\d+|[ivxlcdm]+)\s*${DASH}?\s*$`, 'iu');

/** The page's number as a roman numeral, in the case `token` is written in, where that is what it is. */
function romanOf(token: string, pageNumber: number): PageNumberFormat | undefined {
    if (token === token.toLowerCase() && formatListNumber(pageNumber, 'lowerRoman') === token) return 'lowerRoman';
    if (token === token.toUpperCase() && formatListNumber(pageNumber, 'upperRoman') === token) return 'upperRoman';
    return undefined;
}

/**
 * Where a line is the page's own number and nothing else, the line with its number as '#' (as its digits are
 * in its key, so "- iii -" and "- iv -" are one line): its number in digits, or as its printed label says it
 * (a front matter's "iii", where the document's page labels number the page so). Null where it is not.
 */
function ownNumberKey(text: string, pageNumber: number, label: string | null | undefined): string | null {
    const m = PAGE_LINE.exec(text);
    if (!m) return null;
    const n = m[1];
    if (!(/^\d+$/.test(n) && Number(n) === pageNumber) && !(label && n.toLowerCase() === label.toLowerCase())) return null;
    const at = m[0].lastIndexOf(n, m[2] === undefined ? undefined : m[0].length - m[2].length - 1);
    return `${text.slice(0, at)}#${text.slice(at + n.length)}`.replace(/\p{Nd}+/gu, '#');
}

/** As much of a text as a key holds: its start and its length, past KEY_TEXT. */
function keyOf(text: string): string {
    return text.length > KEY_TEXT ? `${text.slice(0, KEY_TEXT)}|${text.length}` : text;
}

/**
 * A line's text with the number it prints as a page number ("3", "- 3 -", "3/4", "Page 3 of 4", "Seite 3 von
 * 4", "iii") as '#' and how far it is from the page's own: lines that number the pages the same way, a
 * document's own or one numbered from elsewhere (a form placed after a scan, an article numbered as in its
 * journal), are one text however their numbers run. Its other numbers are as printed. Null where the line
 * prints no page number.
 */
function pageNumbered(text: string, pageNumber: number, label: string | null | undefined): string | null {
    if (text.length > NUMBERED_TEXT) return null;
    // The line alone, then a number of the pages' ("3 of 4") in it, then a page's word and its number.
    const line = PAGE_LINE.exec(text);
    let at = -1, token = '';
    if (line) {
        token = line[1];
        at = line[0].lastIndexOf(token, line[2] === undefined ? undefined : line[0].length - line[2].length - 1);
    } else {
        const of = text.matchAll(NUMBER_OF).next().value;
        const word = of ? undefined : text.matchAll(WORD_NUMBER).next().value;
        const m = of ?? word;
        if (!m) return null;
        token = m[1];
        at = m.index! + (of ? m[0].indexOf(token) : m[0].lastIndexOf(token));
    }
    const own = !!label && token.toLowerCase() === label.toLowerCase();
    // (A roman numeral is a page's number where its page label says so: "Part II" is no page's number.)
    if (!own && !/^\d+$/.test(token)) return null;
    const from = own ? 'label' : String(Number(token) - pageNumber);
    return `${text.slice(0, at)}#${from}#${text.slice(at + token.length)}`;
}

/** One line in a page's header or footer band, and what it is compared by. */
interface BandLine {
    runs: readonly RawRun[];
    band: 'header' | 'footer';
    /** Distance from the page's edge it is nearer (rendered space), to its own nearer edge. */
    edge: number;
    /** Its left edge (rendered space). */
    left: number;
    height: number;
    fontSize: number;
    entry: KeyEntry;
    /** Its text as printed, and as its page number numbers the pages (see pageNumbered), as the entry counts them. */
    printed: string;
    numberedAs: string | null;
    ownNumber: boolean;
    /** Whether the page has another line of its text (its numbers as one) at its left edge: a table's column. */
    repeated: boolean;
}

/** One running candidate: the pages a line of this text in this place is on. */
interface KeyEntry {
    pages: number;
    odd: number;
    even: number;
    /** Of those, the pages on which it is the page's own number. */
    numbered: number;
    /** The last page counted (a page counts once). */
    last: number;
    /**
     * The pages of each of its texts as printed, and as a page number numbers the pages: its key holds its
     * numbers as one, which a line's text is only where they are the pages' own or the line is printed alike.
     */
    texts: Map<string, { pages: number; last: number }>;
    /** Of its pages, those on which it is printed alike on another (see SAME_TEXT_PAGES), and the last counted. */
    alike: number;
    alikeLast: number;
}

/** The runs of each page that are its running header and its running footer. */
export interface RunningRuns {
    header: Set<RawRun>;
    footer: Set<RawRun>;
}

/**
 * The lines repeated at the top and the bottom of the pages, each page's that are its running header and
 * footer.
 *
 * A line in a page's top or bottom band (FURNITURE_BAND, as the page is shown, so a turned page's too) is
 * compared with the other pages' by its band, its distance from the page's edge, its left edge, and its
 * text with its numbers as one (so "1/4" and "2/4" are one line), where its numbers are the page's (they number
 * the pages alike, see pageNumbered) or it is printed alike on another page of them: "Question 3" at the top of
 * one page and "Question 4" at the top of the next are two lines, a chapter's head on each of its pages is one.
 * It is running when it is on at least two
 * in five of the pages read, and three of them; or two in five of the odd or of the even pages (and three),
 * so a book's heads, one on the left of even pages and another on the right of odd ones, are found when its
 * first page or a chapter's opening page has none; or, as the page's own number ("3", "3/4", "Seite 3 von
 * 4", a roman numeral the page labels number it in), on two pages.
 *
 * From each edge of a page, its lines are taken while each line on that row is running and no larger than
 * the body's text (by SIZE_SLACK), stopping at the first row that is not: a heading repeated at the top of
 * each page, larger than the text, stays the page's, and so does all below it. Last, a header that stands
 * close to the text under it (GAP_HEIGHTS) reads with it and stays: a table's head row, repeated at the top of
 * each page the table goes on to, is the table's.
 *
 * One pass over the band lines and one over the pages: a map of the lines' keys for the document, linear in
 * its lines.
 */
export function findRunningLines(extracts: readonly PageExtract[], cfg: PdfLayoutConfig, bodyFontSize: number, pageLabels: readonly (string | null)[] | null): Map<PageExtract, RunningRuns> {
    const out = new Map<PageExtract, RunningRuns>();
    if (extracts.length < 2) return out;

    // The keys, by band and text, then by place: a line joins a key of its text within the tolerances of
    // its place, so a line a point lower on one page than on the next is the same line.
    const keys = new Map<string, Map<string, KeyEntry>>();
    const entryFor = (band: string, text: string, edge: number, left: number): KeyEntry => {
        let places = keys.get(`${band}|${text}`);
        if (!places) keys.set(`${band}|${text}`, places = new Map());
        const e = Math.floor(edge / EDGE_TOLERANCE), x = Math.floor(left / LEFT_TOLERANCE);
        for (const de of [0, -1, 1]) for (const dx of [0, -1, 1]) {
            const found = places.get(`${e + de}|${x + dx}`);
            if (found) return found;
        }
        const entry: KeyEntry = { pages: 0, odd: 0, even: 0, numbered: 0, last: -1, texts: new Map(), alike: 0, alikeLast: -1 };
        places.set(`${e}|${x}`, entry);
        return entry;
    };

    const pagesLines: BandLine[][] = [];
    let oddPages = 0, evenPages = 0;
    for (const extract of extracts) {
        const odd = extract.pageNumber % 2 === 1;
        if (odd) oddPages++; else evenPages++;
        const turned = extract.rotation % 180 !== 0;
        const shownH = turned ? extract.authoredW : extract.authoredH;
        const shown = (x: number, y: number, width: number, height: number) => rotateBoundsToRendered({ x, y, width, height }, extract.rotation, extract.authoredW, extract.authoredH);
        // The page's own text in its bands: not a tagged artifact (read as one already), nor a form field's
        // value (the field's, not the page's), nor turned text (read after the page's). Its spaces too, drawn
        // as runs of their own: a line taken without them ran its words together ("Page2").
        const inBands = extract.runs.filter(r => {
            if (r.inArtifact || r.formField || r.angle !== 0 || !r.text) return false;
            const y = shown(r.x, r.yTop, r.width, r.height).y;
            return y < FURNITURE_BAND * shownH || y > (1 - FURNITURE_BAND) * shownH;
        });
        const lines: BandLine[] = [];
        const label = pageLabels?.[extract.pageNumber - 1];
        // The page's lines of each text at each left edge, for a line's text repeated down the band.
        const atLeft = new Map<string, number>();
        const lefts: { left: string; x: number }[] = [];
        for (const line of inBands.length ? buildLines(inBands, cfg) : []) {
            const box = shown(line.x, line.yTop, line.width, line.height);
            const band = box.y < FURNITURE_BAND * shownH ? 'header' : box.y > (1 - FURNITURE_BAND) * shownH ? 'footer' : null;
            if (!band) continue;
            const text = line.fragments.map(f => f.text).join('').replace(/\s+/g, ' ').trim();
            if (!text) continue;
            const numbered = text.length <= 40 ? ownNumberKey(text, extract.pageNumber, label) : null;
            const ownNumber = numbered !== null;
            const keyText = keyOf(numbered ?? text.replace(/\p{Nd}+/gu, '#'));
            const edge = band === 'header' ? box.y : shownH - (box.y + box.height);
            const entry = entryFor(band, keyText, edge, box.x);
            if (entry.last !== extract.pageNumber) {
                entry.last = extract.pageNumber;
                entry.pages++;
                if (odd) entry.odd++; else entry.even++;
                if (ownNumber) entry.numbered++;
            }
            const printed = keyOf(text), numberedAs = pageNumbered(text, extract.pageNumber, label);
            for (const t of numberedAs === null ? [printed] : [printed, keyOf(numberedAs)]) {
                let seen = entry.texts.get(t);
                if (!seen) entry.texts.set(t, seen = { pages: 0, last: -1 });
                if (seen.last !== extract.pageNumber) { seen.last = extract.pageNumber; seen.pages++; }
            }
            const left = `${band}|${keyText}|`, x = Math.floor(box.x / LEFT_TOLERANCE);
            atLeft.set(left + x, (atLeft.get(left + x) ?? 0) + 1);
            lefts.push({ left, x });
            lines.push({ runs: lineRuns(line), band, edge, left: box.x, height: box.height, fontSize: line.fontSize, entry, printed, numberedAs: numberedAs === null ? null : keyOf(numberedAs), ownNumber, repeated: false });
        }
        // (At its own left edge, and a tolerance either side.)
        lines.forEach((l, k) => {
            const { left, x } = lefts[k];
            l.repeated = (atLeft.get(left + (x - 1)) ?? 0) + (atLeft.get(left + x) ?? 0) + (atLeft.get(left + (x + 1)) ?? 0) > 1;
        });
        pagesLines.push(lines);
    }

    const need = (pages: number) => Math.max(RUNNING_PAGES, Math.ceil(RUNNING_SHARE * pages));
    const needAll = need(extracts.length), needOdd = need(oddPages), needEven = need(evenPages);
    // (A line whose text the page repeats down its band at its left edge is a table's column or a list's, not
    // furniture: "Item 46" over "Item 47" at the top of each page a table goes on to.)
    // (Its numbers are as one in its key, so "Page 3" is "Page 4"; but a line whose numbers are not the page's
    // is running only where it is printed alike on other pages: "Question 3", at the top of a page as
    // "Question 4" is on the next, is the page's.)
    const sameText = (l: BandLine) => (l.entry.texts.get(l.printed)?.pages ?? 0) >= SAME_TEXT_PAGES
        || (l.numberedAs !== null && (l.entry.texts.get(l.numberedAs)?.pages ?? 0) >= SAME_TEXT_PAGES);
    extracts.forEach((extract, i) => {
        for (const l of pagesLines[i]) {
            if (l.entry.alikeLast === extract.pageNumber || !sameText(l)) continue;
            l.entry.alikeLast = extract.pageNumber;
            l.entry.alike++;
        }
    });
    const alike = (l: BandLine) => sameText(l) || 2 * l.entry.alike > l.entry.pages;
    const running = (l: BandLine) => !l.repeated && alike(l) && (l.entry.pages >= needAll || l.entry.odd >= needOdd || l.entry.even >= needEven
        || (l.ownNumber && l.entry.numbered >= PAGE_NUMBER_PAGES));
    const small = (l: BandLine) => l.fontSize <= bodyFontSize + SIZE_SLACK;

    extracts.forEach((extract, i) => {
        const lines = pagesLines[i];
        if (!lines.length) return;
        const turned = extract.rotation % 180 !== 0;
        const shownH = turned ? extract.authoredW : extract.authoredH;
        const found: RunningRuns = { header: new Set(), footer: new Set() };
        let any = false;
        for (const band of ['header', 'footer'] as const) {
            // The band's rows from the page's edge: lines on one row (a footer's text on the left, its
            // number on the right) are taken together or not at all.
            const own = lines.filter(l => l.band === band).sort((a, b) => a.edge - b.edge);
            const rows: BandLine[][] = [];
            for (const l of own) {
                const row = rows[rows.length - 1];
                const centre = l.edge + l.height / 2;
                if (row && Math.abs(centre - (row[0].edge + row[0].height / 2)) <= Math.min(l.height, row[0].height) / 2) row.push(l);
                else rows.push([l]);
            }
            let taken = 0;
            while (taken < rows.length && rows[taken].every(l => running(l) && small(l))) taken++;
            if (!taken) continue;
            // The text nearest the taken rows on the page's side of them: a header close above it is its.
            const inner = (row: BandLine[]) => row.reduce((m, l) => Math.max(m, l.edge + l.height), -Infinity);
            const rowHeight = (row: BandLine[]) => row.reduce((m, l) => Math.max(m, l.height), 0);
            const takenRuns = new Set<RawRun>();
            for (let k = 0; k < taken; k++) for (const l of rows[k]) for (const r of l.runs) takenRuns.add(r);
            // (Inward of the rows taken: past the middle of the last one.)
            const from = inner(rows[taken - 1]) - rowHeight(rows[taken - 1]) / 2;
            const rest: { near: number; height: number; left: number }[] = [];
            let nearest = Infinity;
            for (const r of extract.runs) {
                if (takenRuns.has(r) || !r.text.trim()) continue;
                const box = rotateBoundsToRendered({ x: r.x, y: r.yTop, width: r.width, height: r.height }, extract.rotation, extract.authoredW, extract.authoredH);
                const near = band === 'header' ? box.y : shownH - (box.y + box.height);
                if (near + box.height / 2 <= from) continue;
                rest.push({ near, height: box.height, left: box.x });
                if (near < nearest) nearest = near;
            }
            // The left edges of the text nearest them (its row), as places a tolerance wide.
            const placesOf = (lefts: number[]) => new Set(lefts.map(x => Math.floor(x / LEFT_TOLERANCE)));
            let nextLefts = placesOf(rest.filter(t => t.near - nearest <= t.height / 2).map(t => t.left));
            const under = (x: number) => { const k = Math.floor(x / LEFT_TOLERANCE); return nextLefts.has(k - 1) || nextLefts.has(k) || nextLefts.has(k + 1); };
            while (taken > 0) {
                const row = rows[taken - 1];
                // (A page's number is no table's head: a footer's last line, "Seite 1 von 2", close under its
                // address, is the page's number still.)
                if (row.every(l => l.ownNumber)) break;
                // A row of cells over a row whose text starts where two of them do is a table's head, however
                // far apart its rows are.
                const aligned = row.length > 1 && row.filter(l => under(l.left)).length > 1;
                if (!aligned && nearest - inner(row) >= GAP_HEIGHTS * rowHeight(row)) break;
                // Read with what follows: the row stays the page's, and is what follows the row before it.
                nearest = row.reduce((m, l) => Math.min(m, l.edge), nearest);
                nextLefts = placesOf(row.map(l => l.left));
                taken--;
            }
            for (let k = 0; k < taken; k++) for (const l of rows[k]) for (const r of l.runs) { found[band].add(r); any = true; }
        }
        if (any) out.set(extract, found);
    });
    return out;
}

/** A page's header or footer nodes, as read, for markPageFields. */
export interface PageFurniture {
    pageNumber: number;
    kind: 'header' | 'footer';
    nodes: OfficeContentNode[];
    /** The page's rotation: a number's box is cut from its text's along the line only on an upright page. */
    rotation: number;
}

/** A page field found in a node's text: where its number is, and which field it is. */
interface FieldSpan { start: number; end: number; field: 'page' | 'pages'; format?: PageNumberFormat }

/** The text leaves of a node, each with where its text starts in theirs joined. */
function textLeaves(node: OfficeContentNode): { parent: OfficeContentNode; leaf: OfficeContentNode; at: number }[] {
    const out: { parent: OfficeContentNode; leaf: OfficeContentNode; at: number }[] = [];
    let at = 0;
    // (A header's paragraph is a few levels deep at most: a list's item's text.)
    const walk = (parent: OfficeContentNode, depth: number) => {
        if (depth > 4) return;
        for (const child of parent.children ?? []) {
            if (child.type === 'text' && !child.children?.length) { out.push({ parent, leaf: child, at }); at += (child.text ?? '').length; }
            else walk(child, depth + 1);
        }
    };
    walk(node, 0);
    return out;
}

/**
 * The page fields in a header's or footer's text, where its numbers are the page's and the document's pages'
 * (the pages' only where `allPages`: the output holds all of them).
 */
function fieldsIn(text: string, pageNumber: number, numPages: number, allPages: boolean): FieldSpan[] | null {
    for (const m of text.matchAll(NUMBER_OF)) {
        if (Number(m[1]) !== pageNumber || Number(m[2]) !== numPages) continue;
        const start = m.index!, own = start + m[0].indexOf(m[1]);
        const total = start + m[0].lastIndexOf(m[2]);
        const page: FieldSpan = { start: own, end: own + m[1].length, field: 'page' };
        return allPages ? [page, { start: total, end: total + m[2].length, field: 'pages' }] : [page];
    }
    for (const m of text.matchAll(WORD_NUMBER)) {
        const format = /^\d+$/.test(m[1]) ? undefined : romanOf(m[1], pageNumber);
        if (/^\d+$/.test(m[1]) ? Number(m[1]) !== pageNumber : !format) continue;
        const own = m.index! + m[0].lastIndexOf(m[1]);
        return [{ start: own, end: own + m[1].length, field: 'page', ...(format ? { format } : {}) }];
    }
    const bare = BARE_NUMBER.exec(text);
    if (bare) {
        const token = bare[1], start = bare.index + bare[0].indexOf(token);
        if (/^\d+$/.test(token)) return Number(token) === pageNumber ? [{ start, end: start + token.length, field: 'page' }] : null;
        const format = romanOf(token, pageNumber);
        if (format) return [{ start, end: start + token.length, field: 'page', format }];
    }
    return null;
}

/**
 * Marks the page numbers in the pages' headers and footers as the page fields they are (TextMetadata.field:
 * `page` for the page's number, `pages` for the document's pages), so a writer that lays out pages (Word, ODT,
 * RTF, the PDF writer) prints each page's own rather than the first page's on every page.
 *
 * A number is the page's where it is the page's own (its number in the document, or in roman numerals as
 * written) and, in "3/4", "3 of 4", "Seite 3 von 4", where the pages' is the document's. A header or footer
 * is marked only where every page's of the same text (its numbers as one) is: "1/4" printed on every page
 * of a form is the form's number, not the first page's.
 *
 * The furniture is of the document's pages from its first, one after another (a writer numbers its own pages
 * from one); `allPages` where it is of all of them, else the pages' count ("of 4") is left as its text, no
 * longer the output's.
 */
export function markPageFields(furniture: readonly PageFurniture[], numPages: number, allPages = true): void {
    const groups = new Map<string, { node: OfficeContentNode; leaves: ReturnType<typeof textLeaves>; spans: FieldSpan[] | null; rotation: number }[]>();
    for (const page of furniture) {
        for (const node of page.nodes) {
            const leaves = textLeaves(node);
            if (!leaves.length) continue;
            const text = leaves.map(l => l.leaf.text ?? '').join('');
            if (!/[\divxlcdm]/i.test(text) || text.length > 200) continue;
            const key = `${page.kind}|${text.replace(/\p{Nd}+/gu, '#').trim()}`;
            const spans = fieldsIn(text, page.pageNumber, numPages, allPages);
            let group = groups.get(key);
            if (!group) groups.set(key, group = []);
            group.push({ node, leaves, spans, rotation: page.rotation });
        }
    }
    for (const group of groups.values()) {
        if (group.some(g => !g.spans)) continue;
        for (const { leaves, spans, rotation } of group) {
            // From the last, so a leaf cut for the pages' number still holds the page's.
            for (const span of [...spans!].sort((a, b) => b.start - a.start)) {
                const hit = leaves.find(l => l.at <= span.start && span.end <= l.at + (l.leaf.text ?? '').length);
                if (!hit) continue;
                cutField(hit, span, rotation);
            }
        }
    }
}

/** Cuts a page field's number out of its text leaf, into a text node of its own marked as the field. */
function cutField(hit: { parent: OfficeContentNode; leaf: OfficeContentNode; at: number }, span: FieldSpan, rotation: number): void {
    const { parent, leaf } = hit;
    const children = parent.children!;
    const index = children.indexOf(leaf);
    if (index < 0) return;
    const text = leaf.text ?? '';
    const from = span.start - hit.at, to = span.end - hit.at;
    const piece = (s: number, e: number, field?: FieldSpan): OfficeContentNode => {
        const node: OfficeContentNode = { ...leaf, text: text.slice(s, e) };
        if (leaf.formatting) node.formatting = { ...leaf.formatting };
        const meta = { ...(leaf.metadata as TextMetadata | undefined) };
        if (field) { meta.field = field.field; if (field.format) meta.fieldFormat = field.format; }
        if (Object.keys(meta).length) node.metadata = meta; else delete node.metadata;
        // Its share of the leaf's box along an upright line, else the leaf's box.
        if (leaf.bounds) {
            const b = leaf.bounds;
            node.bounds = rotation === 0 && text.length
                ? { ...b, x: b.x + (b.width * s) / text.length, width: (b.width * (e - s)) / text.length }
                : { ...b };
        }
        return node;
    };
    const pieces: OfficeContentNode[] = [];
    if (from > 0) pieces.push(piece(0, from));
    pieces.push(piece(from, to, span));
    if (to < text.length) pieces.push(piece(to, text.length));
    children.splice(index, 1, ...pieces);
    // The later spans of this leaf (an earlier number in the text) now find the piece before.
    if (from > 0) { hit.leaf = pieces[0]; }
}

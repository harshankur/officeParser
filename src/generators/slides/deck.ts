/**
 * A document as a deck of slides, for the presentation writers (PPTX, ODP).
 *
 * - A deck's slides (PPTX, ODP) as they are: each shape (a `drawing` node) with the text in its box,
 *   each text box, picture and table at its place (`bounds`), the slide's background, its speaker notes.
 * - A PDF's pages, a slide each at the page's size: each paragraph a text box exactly where it stood
 *   (no insets, no fitting, its runs' fonts, sizes and colours), pictures and tables in their places,
 *   and the page's background picture where one is given (`backgrounds`: the page drawn without its
 *   text, as a caller renders it), under them. A deck has one slide size: the pages' most common, and
 *   a page of another size scaled into it, centred.
 * - A flowing document: a title slide, then a slide for each heading (of the first two levels) with its
 *   text as bullets, carried on to slides of the same title where it does not fit, each table and
 *   picture on a slide of its own (a long table over several, its header repeated), and its notes on a
 *   slide at the end.
 *
 * @module generators/slides/deck
 */

import { HeadingMetadata, ListMetadata, NodeBounds, OfficeContentNode, OfficeMetadata, ShapeMetadata, SlideMetadata, TextFormatting, TextMetadata } from '../../types.js';
import { lengthToPt } from '../../utils/officeGenUtils.js';
import { layoutTableRows } from '../../utils/tableLayout.js';

/** A box on a slide, in points from its top left. */
export interface Box { x: number; y: number; w: number; h: number }

export interface DeckRun { text: string; bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean; size?: number; color?: string; font?: string; link?: string; superscript?: boolean; subscript?: boolean; lineBreak?: boolean }
export interface DeckParagraph { runs: DeckRun[]; align?: 'left' | 'center' | 'right' | 'justify'; bullet?: { level: number; char?: string; numbered?: boolean; /** A numbered item's marker as it is shown (`2.`, `1.1.`, `iv.`), written as text: each item is a paragraph, often a box, of its own. */ text?: string } ; mono?: boolean; /** The first line's indent past the others, in points (negative: hanging). */ firstIndent?: number; /** From baseline to baseline, in points, where the source measured it. */ lineSpacing?: number }

export type DeckItem =
    | { kind: 'text'; box: Box; paragraphs: DeckParagraph[]; anchor?: 'top' | 'middle' | 'bottom'; insets?: { top: number; right: number; bottom: number; left: number }; exact?: boolean; shape?: ShapeMetadata; rotation?: number; role?: 'title' | 'body' }
    | { kind: 'picture'; box: Box; attachment: string; rotation?: number; /** Its alternative text. */ alt?: string }
    | { kind: 'table'; box: Box; rows: { cells: { paragraphs: DeckParagraph[]; colSpan: number; rowSpan: number; bold: boolean; fill?: string; col?: number }[] }[]; widths: number[]; heights?: number[] }
    | { kind: 'shape'; box: Box; shape: ShapeMetadata };

export interface DeckSlide { items: DeckItem[]; background?: string; picture?: Uint8Array; notes?: DeckParagraph[] }
export interface Deck { width: number; height: number; slides: DeckSlide[]; unrepresentable: Set<string> }

/** At most this many slides are made. */
const MAX_SLIDES = 10_000;
/** A flowing document's slides: 16:9, its title's box and its body's. */
const FLOW_W = 720, FLOW_H = 405;
const TITLE_BOX: Box = { x: 36, y: 22, w: 648, h: 56 };
const BODY_BOX: Box = { x: 36, y: 90, w: 648, h: 290 };
const TITLE_SIZE = 28, BODY_SIZE = 18;

/** What the deck needs of a writer: whether a node is written (onNode, a string its text instead). */
export interface DeckSource {
    keep(node: OfficeContentNode): Promise<boolean | string>;
    /** Whether charts are written (includeCharts): a chart left out is not reported. */
    includeCharts: boolean;
    /** A chart's data as a table (its labels down, its series across), or undefined where it has none. */
    chartTable(node: OfficeContentNode): OfficeContentNode | undefined;
}

/**
 * A font's family as a slide names it: a PDF's font name without its subset prefix, its style
 * (`ABCDEF+Ubuntu-Bold` is Ubuntu, `TimesNewRomanPSMT` Times New Roman, `Arial,BoldItalic` Arial).
 */
export function fontFamily(name: string): string {
    const bare = name.replace(/^[A-Z]{6}\+/, '').split(/[-,]/)[0].replace(/(?:PS)?MT$|PS$/, '');
    return /\s/.test(bare) ? bare : bare.replace(/([a-z])([A-Z])/g, '$1 $2');
}

const sizeOf = (fmt: TextFormatting | undefined): number | undefined => {
    const pt = fmt?.size ? lengthToPt(String(fmt.size)) : null;
    return pt != null && pt > 0 && pt < 4000 ? Math.round(pt * 100) / 100 : undefined;
};

/** A link a slide can open: a web, mail or phone address. */
const linkOf = (meta: TextMetadata | undefined): string | undefined => {
    const url = (meta as { link?: string } | undefined)?.link;
    return typeof url === 'string' && /^(https?:|mailto:|tel:)/i.test(url.trim()) ? url.trim() : undefined;
};

/** A block's runs: its text with their looks, line breaks kept, notes cited by their numbers. */
async function runsOf(node: OfficeContentNode, source: DeckSource, notes: { list: DeckParagraph[]; seen: Map<OfficeContentNode, number> }, inherited: TextFormatting = {}, depth = 0): Promise<DeckRun[]> {
    const runs: DeckRun[] = [];
    for (const child of node.children ?? []) {
        const kept = await source.keep(child);
        if (kept === false) continue;
        if (typeof kept === 'string') { runs.push({ text: kept }); continue; }
        const fmt: TextFormatting = { ...inherited, ...(child.formatting ?? {}) };
        if (child.type === 'text' || child.type === 'code') {
            if (child.text) runs.push({
                text: child.text, ...(fmt.bold ? { bold: true } : {}), ...(fmt.italic ? { italic: true } : {}), ...(fmt.underline ? { underline: true } : {}),
                ...(fmt.strikethrough ? { strike: true } : {}), ...(fmt.superscript ? { superscript: true } : {}), ...(fmt.subscript ? { subscript: true } : {}),
                ...(sizeOf(fmt) ? { size: sizeOf(fmt) } : {}), ...(fmt.color ? { color: fmt.color } : {}), ...(fmt.font ? { font: fontFamily(fmt.font) } : child.type === 'code' ? { font: 'Courier New' } : {}),
                ...(linkOf(child.metadata as TextMetadata) ? { link: linkOf(child.metadata as TextMetadata) } : {}),
            });
        } else if (child.type === 'break') {
            const kind = (child.metadata as { breakType?: string } | undefined)?.breakType;
            if (kind !== 'section' && kind !== 'page' && kind !== 'lastRenderedPage') runs.push({ text: '', lineBreak: true });
        } else if (depth < 32) {
            for (const r of await runsOf(child, source, notes, fmt, depth + 1)) runs.push(r);
        }
        // A note cited here: its number, and its text on the notes slide.
        for (const note of child.notes ?? []) {
            let n = notes.seen.get(note);
            if (n === undefined) {
                n = notes.list.length + 1;
                notes.seen.set(note, n);
                const text = (await runsOf(note, source, notes, {}, depth + 1)).map(r => r.text).join('').trim();
                notes.list.push({ runs: [{ text: `${n}. ${text}`, size: 14 }] });
            }
            runs.push({ text: String(n), superscript: true });
        }
    }
    return runs;
}

/**
 * A page paragraph's lines as the page set them, each its runs and its box (where every run of text has
 * its place; else none): a line starts where a run sits below the one before it.
 */
function placedLines(node: OfficeContentNode, runs: DeckRun[]): { runs: DeckRun[]; box: NodeBounds }[] | undefined {
    const placed: NodeBounds[] = [];
    const walk = (n: OfficeContentNode, d: number) => { for (const c of n.children ?? []) { if ((c.type === 'text' || c.type === 'code') && c.text) { if (c.bounds) placed.push(c.bounds); else placed.length = 0; } else if (d < 16) walk(c, d + 1); } };
    walk(node, 0);
    // (A note's number, which no text node gave, is no placed run.)
    const texts = runs.filter(r => !r.lineBreak && r.text && !(r.superscript && /^\d+$/.test(r.text)));
    if (!texts.length || placed.length !== texts.length) return undefined;
    const lines: { runs: DeckRun[]; box: NodeBounds }[] = [];
    let pending: DeckRun[] = [], at = 0, last: NodeBounds | undefined;
    for (const run of runs) {
        if (run !== texts[at]) { pending.push(run); continue; }
        const b = placed[at++];
        if (!last || b.y > last.y + last.height * 0.6) {
            // The line before ends without its trailing space or break; this one starts without its leading space.
            const prev = lines[lines.length - 1];
            if (prev) { while (pending.length && pending[pending.length - 1].lineBreak) pending.pop(); prev.runs.push(...pending); pending = []; }
            lines.push({ runs: [], box: { ...b } });
            pending.push(prev ? { ...run, text: run.text.trimStart() } : run);
        } else {
            const box = lines[lines.length - 1].box;
            const right = Math.max(box.x + box.width, b.x + b.width), bottom = Math.max(box.y + box.height, b.y + b.height);
            box.x = Math.min(box.x, b.x); box.y = Math.min(box.y, b.y); box.width = right - box.x; box.height = bottom - box.y;
            pending.push(run);
        }
        lines[lines.length - 1].runs.push(...pending);
        pending = [];
        last = b;
    }
    lines[lines.length - 1].runs.push(...pending);
    for (const line of lines.slice(0, -1)) {
        const end = line.runs.length - 1;
        if (end >= 0 && !line.runs[end].lineBreak) line.runs[end] = { ...line.runs[end], text: line.runs[end].text.replace(/\s+$/, '') };
    }
    return lines;
}

/** A page paragraph's lines joined with a break where each ended on the page. */
const joinLines = (lines: { runs: DeckRun[] }[]): DeckRun[] => lines.flatMap((line, i) => i ? [{ text: '', lineBreak: true }, ...line.runs] : line.runs);

/** The distance from one line to the next, where lines were measured on the page and are evenly apart enough to say. */
function pitchOf(lines: { box: NodeBounds }[]): number | undefined {
    if (lines.length < 2) return undefined;
    const pitch = (lines[lines.length - 1].box.y - lines[0].box.y) / (lines.length - 1);
    return pitch > 0 && Number.isFinite(pitch) ? Math.round(pitch * 100) / 100 : undefined;
}

/** A paragraph's sizes (indent, line spacing) scaled with its page. */
const scaledParagraph = (p: DeckParagraph, scale: number): DeckParagraph => scale === 1 ? p
    : { ...p, ...(p.firstIndent ? { firstIndent: p.firstIndent * scale } : {}), ...(p.lineSpacing ? { lineSpacing: p.lineSpacing * scale } : {}) };

/**
 * A page paragraph's lines in groups that share a left edge (the first line's own indent kept as one), each
 * group a box of its own: a paragraph set beside a floating table or picture, its lines narrowed, then
 * full, keeps both. Undefined where its lines share one edge or have no places.
 */
function placedGroups(node: OfficeContentNode, p: DeckParagraph): { paragraph: DeckParagraph; box: NodeBounds }[] | undefined {
    if (p.bullet || (p.align && p.align !== 'left' && p.align !== 'justify')) return undefined;
    const lines = placedLines(node, p.runs);
    if (!lines || lines.length < 2) return undefined;
    const pitch = pitchOf(lines);
    const groups: { lines: typeof lines; left: number; first: number }[] = [];
    for (const [i, line] of lines.entries()) {
        const g = groups[groups.length - 1];
        if (g && Math.abs(line.box.x - g.left) <= 3) g.lines.push(line);
        // The paragraph's first line set in from (or out past) the rest: its indent.
        else if (g && i === 1) { g.left = line.box.x; g.lines.push(line); }
        else groups.push({ lines: [line], left: line.box.x, first: line.box.x });
    }
    if (groups.length < 2 && Math.abs(groups[0].first - groups[0].left) <= 3) return undefined;
    return groups.map(g => {
        let right = -Infinity, top = Infinity, bottom = -Infinity;
        for (const l of g.lines) { right = Math.max(right, l.box.x + l.box.width); top = Math.min(top, l.box.y); bottom = Math.max(bottom, l.box.y + l.box.height); }
        const indent = g.first - g.left;
        const own = pitchOf(g.lines) ?? pitch;
        const paragraph: DeckParagraph = { ...p, runs: joinLines(g.lines), ...(Math.abs(indent) > 3 ? { firstIndent: indent } : {}), ...(own ? { lineSpacing: own } : {}) };
        return { paragraph, box: { x: Math.min(g.left, g.first), y: top, width: right - Math.min(g.left, g.first), height: bottom - top } };
    });
}

/** The pictures in a node, with their alternative text. */
function picturesIn(node: OfficeContentNode): { attachment: string; alt?: string }[] {
    const out: { attachment: string; alt?: string }[] = [];
    const walk = (n: OfficeContentNode, d: number) => {
        if (n.type === 'image') {
            const m = n.metadata as { attachmentName?: string; altText?: string } | undefined;
            if (m?.attachmentName) out.push({ attachment: m.attachmentName, ...(typeof m.altText === 'string' && m.altText.trim() ? { alt: m.altText.trim() } : {}) });
        }
        if (d < 16) for (const c of n.children ?? []) walk(c, d + 1);
    };
    walk(node, 0);
    return out;
}

/** A block as a paragraph: its runs, its alignment, a list item's bullet and level. */
async function paragraphOf(node: OfficeContentNode, source: DeckSource, notes: { list: DeckParagraph[]; seen: Map<OfficeContentNode, number> }, kept: true | string = true, asPlaced = false): Promise<DeckParagraph> {
    const read = typeof kept === 'string' ? [{ text: kept }] : await runsOf(node, source, notes);
    // A page's paragraph: its lines as the page broke them, as far apart as they were.
    const lines = asPlaced && typeof kept !== 'string' ? placedLines(node, read) : undefined;
    const pitch = lines && pitchOf(lines);
    const meta = node.metadata as (ListMetadata & { alignment?: string }) | undefined;
    const align = meta?.alignment === 'center' || meta?.alignment === 'right' || meta?.alignment === 'justify' ? meta.alignment : undefined;
    const p: DeckParagraph = { runs: lines ? joinLines(lines) : read, ...(align ? { align } : {}), ...(pitch ? { lineSpacing: pitch } : {}) };
    if (node.type === 'list') {
        const index = typeof meta?.itemIndex === 'number' && Number.isFinite(meta.itemIndex) ? Math.max(0, Math.trunc(meta.itemIndex)) : 0;
        p.bullet = { level: Math.max(0, Math.min(8, meta?.indentation ?? 0)), ...(meta?.listType === 'ordered'
            ? { numbered: true, text: typeof meta?.marker === 'string' ? meta.marker.replace(/\s+/g, ' ').trim().slice(0, 64) : `${index + 1}.` }
            : { char: typeof meta?.marker === 'string' && meta.marker ? meta.marker : '•' }) };
    }
    if (node.type === 'code') p.mono = true;
    return p;
}

/**
 * A table's rows and cells as the slide's table holds them: laid out as DOCX and ODT lay tables out (see
 * layoutTableRows: a sparse sheet's cells at their own columns, the columns between them empty, merges
 * covering what they span), with its columns' widths, and its rows' heights where its source measured them.
 */
async function tableOf(node: OfficeContentNode, source: DeckSource, notes: { list: DeckParagraph[]; seen: Map<OfficeContentNode, number> }) {
    type DeckCell = { paragraphs: DeckParagraph[]; colSpan: number; rowSpan: number; bold: boolean; fill?: string; col: number };
    // The rows and cells kept, laid out.
    const kept: OfficeContentNode[] = [];
    const keptText = new Map<OfficeContentNode, true | string>();
    for (const row of node.children ?? []) {
        if (row.type !== 'row' || await source.keep(row) === false) continue;
        const cells: OfficeContentNode[] = [];
        for (const cell of row.children ?? []) {
            if (cell.type !== 'cell') continue;
            const k = await source.keep(cell);
            if (k === false) continue;
            keptText.set(cell, k);
            cells.push(cell);
        }
        kept.push({ ...row, children: cells });
    }
    const layout = layoutTableRows(kept);
    const rows: { cells: DeckCell[] }[] = [];
    const placed: { row: number; col: number; colSpan: number; rowSpan: number; len: number; bounds?: NodeBounds }[] = [];
    let cols = 1;
    for (const [ri, slots] of (layout?.rows ?? []).entries()) {
        const cells: DeckCell[] = [];
        for (const slot of slots) {
            if (slot.kind === 'covered') continue;
            if (slot.kind === 'gap') { cells.push({ paragraphs: [], colSpan: 1, rowSpan: 1, bold: false, col: slot.col }); cols = Math.max(cols, slot.col + 1); continue; }
            const cell = slot.cell;
            const meta = cell.metadata as { style?: string; backgroundColor?: string } | undefined;
            const paragraphs: DeckParagraph[] = [];
            if ((cell.children ?? []).some(c => c.type === 'paragraph' || c.type === 'list' || c.type === 'heading' || c.type === 'table')) {
                for (const child of cell.children ?? []) {
                    // A table inside the cell: its rows as lines of the cell, their cells apart by tabs.
                    if (child.type === 'table') { for (const r of child.children ?? []) if (r.type === 'row') paragraphs.push({ runs: [{ text: (r.children ?? []).map(x => (x.text || '').trim()).join('\t') }] }); continue; }
                    paragraphs.push(await paragraphOf(child, source, notes));
                }
            } else paragraphs.push(await paragraphOf(cell, source, notes, keptText.get(cell) ?? true));
            cells.push({ paragraphs, colSpan: slot.colSpan, rowSpan: slot.rowSpan, bold: meta?.style === 'header', ...(meta?.backgroundColor ? { fill: meta.backgroundColor } : {}), col: slot.col });
            placed.push({ row: ri, col: slot.col, colSpan: slot.colSpan, rowSpan: slot.rowSpan, len: paragraphs.map(p => p.runs.map(x => x.text).join('')).join(' ').length, bounds: cell.bounds });
            cols = Math.max(cols, slot.col + slot.colSpan);
        }
        rows.push({ cells });
    }
    // Columns (and rows) as the source measured them, where every one has a cell of its own with bounds;
    // else columns as wide as their longest text, in proportion.
    const measured = (count: number, at: (p: typeof placed[number]) => number, span: (p: typeof placed[number]) => number, size: (b: NodeBounds) => number) => {
        const out = new Array<number>(count).fill(0);
        for (const p of placed) if (span(p) === 1 && p.bounds && at(p) < count) out[at(p)] = Math.max(out[at(p)], size(p.bounds));
        return out.every(v => v > 0) ? out : undefined;
    };
    const longest = new Array<number>(cols).fill(4);
    for (const p of placed) if (p.colSpan === 1 && p.col < cols) longest[p.col] = Math.max(longest[p.col], Math.min(40, p.len));
    const widths = measured(cols, p => p.col, p => p.colSpan, b => b.width) ?? longest;
    const heights = measured(rows.length, p => p.row, p => p.rowSpan, b => b.height);
    return { rows, widths, ...(heights ? { heights } : {}) };
}

/** The bounds a node, or the first of its descendants, has. */
const boxOf = (b: NodeBounds | undefined, scale = 1, dx = 0, dy = 0): Box | undefined => b && b.width >= 0 && b.height >= 0 ? { x: dx + b.x * scale, y: dy + b.y * scale, w: Math.max(1, b.width * scale), h: Math.max(1, b.height * scale) } : undefined;

/** Builds the deck of a document (see the module's description). */
export async function buildDeck(content: OfficeContentNode[], metadata: OfficeMetadata, source: DeckSource, options: { backgrounds?: (Uint8Array | undefined | null)[] }): Promise<Deck> {
    const unrepresentable = new Set<string>();
    const notes = { list: [] as DeckParagraph[], seen: new Map<OfficeContentNode, number>() };
    const slideNodes = content.filter(n => n.type === 'slide');
    const pageNodes = content.filter(n => n.type === 'page');
    const placed = (nodes: OfficeContentNode[]) => nodes.some(n => (n.children ?? []).some(c => c.bounds));

    // A deck's slides, or a PDF's pages, laid out as they are.
    if ((slideNodes.length && placed(slideNodes)) || (pageNodes.length && placed(pageNodes))) {
        const isSlides = slideNodes.length > 0;
        const nodes = isSlides ? slideNodes : pageNodes;
        // The slide size: the deck's own, or the pages' most common.
        const sizes = new Map<string, number>();
        for (const n of nodes) { const m = n.metadata as { pageWidth?: number; pageHeight?: number } | undefined; if (m?.pageWidth && m.pageHeight) { const k = `${Math.round(m.pageWidth)}x${Math.round(m.pageHeight)}`; sizes.set(k, (sizes.get(k) ?? 0) + 1); } }
        const common = [...sizes].sort((a, b) => b[1] - a[1])[0]?.[0];
        const [width, height] = isSlides && metadata.pageSetup ? [metadata.pageSetup.width, metadata.pageSetup.height]
            : common ? common.split('x').map(Number) as [number, number] : [metadata.pageSetup?.width ?? FLOW_W, metadata.pageSetup?.height ?? FLOW_H];
        const slides: DeckSlide[] = [];
        const printed = new Set<number>();
        for (const [index, node] of nodes.entries()) {
            if (slides.length >= MAX_SLIDES) { unrepresentable.add('slides past the deck\'s limit'); break; }
            if (await source.keep(node) === false) continue;
            const pm = node.metadata as { pageWidth?: number; pageHeight?: number } | undefined;
            // A page of another size, scaled into the slide, centred.
            const scale = !isSlides && pm?.pageWidth && pm.pageHeight ? Math.min(width / pm.pageWidth, height / pm.pageHeight) : 1;
            const dx = !isSlides && pm?.pageWidth ? (width - pm.pageWidth * scale) / 2 : 0, dy = !isSlides && pm?.pageHeight ? (height - pm.pageHeight * scale) / 2 : 0;
            const slide: DeckSlide = { items: [] };
            const background = (node.metadata as SlideMetadata | undefined)?.background;
            if (background) slide.background = background;
            const picture = options.backgrounds?.[index];
            if (picture) slide.picture = picture;
            // A page's block in its place: its lines kept as the page broke and spaced them, a paragraph whose
            // lines start at more than one edge (beside a floating table or picture) a box for each group of them.
            const placeBlock = async (block: OfficeContentNode, kept: true | string) => {
                const rotation = (block.metadata as { rotation?: number } | undefined)?.rotation;
                if (!rotation && kept === true && (block.type === 'paragraph' || block.type === 'heading')) {
                    const groups = placedGroups(block, await paragraphOf(block, source, notes));
                    if (groups) {
                        for (const g of groups) { const at = boxOf(g.box, scale, dx, dy); if (at) slide.items.push({ kind: 'text', box: at, paragraphs: [scaledParagraph(g.paragraph, scale)], exact: true }); }
                        return;
                    }
                }
                const p = await paragraphOf(block, source, notes, kept, true);
                if (!p.runs.some(r => r.text.trim())) return;
                const at = boxOf(block.bounds, scale, dx, dy);
                if (at) slide.items.push({ kind: 'text', box: at, paragraphs: [scaledParagraph(p, scale)], exact: true, ...(rotation ? { rotation } : {}) });
                else slide.items.push({ kind: 'text', box: { ...BODY_BOX, x: width * 0.05, w: width * 0.9, y: height * 0.1 + slide.items.length * 24, h: 24 }, paragraphs: [p] });
            };
            // A table's cells' text placed block by block, where every block of text in it has its place.
            const cellsPlaced = async (table: OfficeContentNode): Promise<boolean> => {
                const blocks: OfficeContentNode[] = [];
                const walk = (n: OfficeContentNode, d: number): boolean => {
                    for (const c of n.children ?? []) {
                        // A cell of runs (no blocks of its own) is its one block.
                        if (c.type === 'cell' && !(c.children ?? []).some(x => ['paragraph', 'heading', 'list', 'code', 'table'].includes(x.type))) {
                            if (!(c.text || '').trim()) continue;
                            // Placed where its text is (a cell's box is wider than a centred or right-aligned number).
                            const texts = (c.children ?? []).filter(x => (x.text || '').trim());
                            const placed = texts.length && texts.every(x => x.bounds) ? texts.map(x => x.bounds!) : undefined;
                            const box = placed ? placed.reduce((u, b) => { const x = Math.min(u.x, b.x), y = Math.min(u.y, b.y); return { x, y, width: Math.max(u.x + u.width, b.x + b.width) - x, height: Math.max(u.y + u.height, b.y + b.height) - y }; }) : c.bounds;
                            if (!box) return false;
                            blocks.push({ ...c, bounds: box });
                            continue;
                        }
                        if (c.type === 'row' || c.type === 'cell' || c.type === 'table') { if (d >= 16 || !walk(c, d + 1)) return false; continue; }
                        if (!['paragraph', 'heading', 'list', 'code'].includes(c.type)) { if ((c.text || '').trim() || c.type === 'image') return false; continue; }
                        if (!c.bounds) return false;
                        blocks.push(c);
                    }
                    return true;
                };
                if (!walk(table, 0)) return false;
                for (const b of blocks) { const k = await source.keep(b); if (k !== false) await placeBlock(b, k); }
                return true;
            };
            const children = node.children ?? [];
            for (let i = 0; i < children.length; i++) {
                const child = children[i];
                const kept = await source.keep(child);
                if (kept === false) continue;
                const box = boxOf(child.bounds, scale, dx, dy);
                if (child.type === 'drawing') {
                    // A shape, and the text in its box (the nodes after it with its bounds).
                    const shape = (child.metadata ?? {}) as ShapeMetadata;
                    const texts: OfficeContentNode[] = [];
                    while (i + 1 < children.length && ['paragraph', 'heading', 'list'].includes(children[i + 1].type) && sameBox(children[i + 1].bounds, child.bounds)) texts.push(children[++i]);
                    if (!box) continue;
                    if (texts.length) {
                        const paragraphs = [];
                        for (const t of texts) { const k = await source.keep(t); if (k !== false) paragraphs.push(await paragraphOf(t, source, notes, k)); }
                        slide.items.push({ kind: 'text', box, paragraphs, shape, ...(shape.textAnchor ? { anchor: shape.textAnchor } : {}), ...(shape.textInsets ? { insets: shape.textInsets } : {}), ...(shape.rotation ? { rotation: shape.rotation } : {}) });
                    } else slide.items.push({ kind: 'shape', box, shape });
                    continue;
                }
                if (['paragraph', 'heading', 'list', 'code'].includes(child.type)) {
                    // A deck's text box holds the paragraphs that share its bounds; a page's paragraph is its own.
                    const group: OfficeContentNode[] = [child];
                    while (isSlides && i + 1 < children.length && ['paragraph', 'heading', 'list'].includes(children[i + 1].type) && sameBox(children[i + 1].bounds, child.bounds)) group.push(children[++i]);
                    if (!isSlides) { await placeBlock(child, kept); continue; }
                    const rotation = (child.metadata as { rotation?: number } | undefined)?.rotation;
                    const paragraphs = [];
                    for (const g of group) { const k = g === child ? kept : await source.keep(g); if (k !== false) paragraphs.push(await paragraphOf(g, source, notes, k)); }
                    if (!paragraphs.some(p => p.runs.some(r => r.text.trim()))) continue;
                    if (box) slide.items.push({ kind: 'text', box, paragraphs, ...(rotation ? { rotation } : {}) });
                    else slide.items.push({ kind: 'text', box: { ...BODY_BOX, x: width * 0.05, w: width * 0.9, y: height * 0.1 + slide.items.length * 24, h: 24 }, paragraphs });
                    continue;
                }
                if (child.type === 'image') {
                    const [picture] = picturesIn(child);
                    if (picture && box) slide.items.push({ kind: 'picture', box, ...picture });
                    else if (picture) unrepresentable.add('pictures without a place');
                    continue;
                }
                if (child.type === 'table') {
                    // On a page drawn behind its slide, the table's rules are in the picture: its cells' text in
                    // their places over it, where every cell's text has a place.
                    if (slide.picture && !isSlides && await cellsPlaced(child)) continue;
                    const t = await tableOf(child, source, notes);
                    if (box) slide.items.push({ kind: 'table', box, ...t });
                    if (picturesIn(child).length) unrepresentable.add('pictures inside table cells');
                    continue;
                }
                if (child.type === 'chart') {
                    if (!source.includeCharts) continue;
                    // A chart as a table of its data, in its place.
                    const table = source.chartTable(child);
                    if (table && box) { const t = await tableOf(table, source, notes); slide.items.push({ kind: 'table', box, rows: t.rows, widths: t.widths }); }
                    else unrepresentable.add('charts without their data');
                    continue;
                }
                // Any other block that holds text (a page's notes, printed at its foot): its blocks in their places.
                if (child.type !== 'break' && (child.text || child.children?.length)) {
                    const blocks: OfficeContentNode[] = [];
                    const walk = (n: OfficeContentNode, d: number) => { for (const c of n.children ?? []) { if (['paragraph', 'heading', 'list', 'code'].includes(c.type)) blocks.push(c); else if (d < 16) walk(c, d + 1); } };
                    walk(child, 0);
                    if (!blocks.length && (child.text || '').trim()) blocks.push(child);
                    for (const b of blocks) { const k = b === child ? kept : await source.keep(b); if (k !== false) await placeBlock(b, k); }
                }
            }
            // The notes cited here that the page printed (their blocks placed): in their places, numbered.
            const cited: OfficeContentNode[] = [];
            const find = (n: OfficeContentNode, d: number) => { if (n !== node) cited.push(...(n.notes ?? [])); if (d < 16) for (const c of n.children ?? []) find(c, d + 1); };
            find(node, 0);
            for (const note of cited) {
                const n = notes.seen.get(note);
                const blocks = note.children ?? [];
                if (n === undefined || printed.has(n) || !blocks.length || !blocks.every(b => b.bounds)) continue;
                printed.add(n);
                for (const [k, block] of blocks.entries()) {
                    const p = await paragraphOf(block, source, notes, true, true);
                    const box = boxOf(block.bounds, scale, dx, dy);
                    if (box) slide.items.push({ kind: 'text', box, paragraphs: [k ? p : { ...p, runs: [{ text: String(n), superscript: true }, { text: ' ' }, ...p.runs] }], exact: true });
                }
            }
            // A slide's speaker notes.
            const own: DeckParagraph[] = [];
            for (const note of node.notes ?? []) { if (await source.keep(note) === false) continue; for (const block of note.children ?? []) own.push(await paragraphOf(block, source, notes)); }
            if (own.length) slide.notes = own;
            slides.push(slide);
        }
        slides.push(...notesSlides(notes.list.filter((_, i) => !printed.has(i + 1)), width, height));
        return { width, height, slides, unrepresentable };
    }

    // A flowing document: slides of its headings.
    const deck = await flowDeck(content, metadata, source, notes, unrepresentable);
    // Pictures behind the slides, one a slide in order.
    deck.slides.forEach((slide, i) => { const picture = options.backgrounds?.[i]; if (picture) slide.picture = picture; });
    return deck;
}

/** Whether two bounds are one box. */
function sameBox(a: NodeBounds | undefined, b: NodeBounds | undefined): boolean {
    return !!a && !!b && Math.abs(a.x - b.x) < 0.5 && Math.abs(a.y - b.y) < 0.5 && Math.abs(a.width - b.width) < 0.5 && Math.abs(a.height - b.height) < 0.5;
}

/** The slides that list the notes cited on the slides before them, twelve a slide. */
function notesSlides(list: DeckParagraph[], width: number, height: number): DeckSlide[] {
    const out: DeckSlide[] = [];
    for (let at = 0; at < list.length; at += 12) out.push(notesSlide(list.slice(at, at + 12), width, height));
    return out;
}

/** A slide that lists notes. */
function notesSlide(list: DeckParagraph[], width: number, height: number): DeckSlide {
    return { items: [
        { kind: 'text', box: { x: width * 0.05, y: height * 0.05, w: width * 0.9, h: height * 0.14 }, paragraphs: [{ runs: [{ text: 'Notes', bold: true, size: TITLE_SIZE }] }], role: 'title' },
        { kind: 'text', box: { x: width * 0.05, y: height * 0.22, w: width * 0.9, h: height * 0.72 }, paragraphs: list, role: 'body' },
    ] };
}

/** How many lines a paragraph takes in a body of `width` at `size` (half an em a character, as an estimate). */
function linesOf(p: DeckParagraph, width: number, size: number): number {
    const indent = p.bullet ? 22 * (p.bullet.level + 1) : 0;
    const perLine = Math.max(8, Math.floor((width - indent) / (size * 0.5)));
    let lines = 0;
    for (const part of p.runs.map(r => r.lineBreak ? '\n' : r.text).join('').split('\n')) lines += Math.max(1, Math.ceil(part.length / perLine));
    return lines;
}

/** A paragraph cut into pieces of at most `lines` lines each, at its words. */
function cutParagraph(p: DeckParagraph, lines: number, width: number, size: number): DeckParagraph[] {
    const perLine = Math.max(8, Math.floor((width - (p.bullet ? 22 * (p.bullet.level + 1) : 0)) / (size * 0.5)));
    const room = Math.max(1, lines) * perLine;
    const out: DeckParagraph[] = [];
    let current: DeckRun[] = [], used = 0;
    for (const run of p.runs) {
        const words = run.lineBreak ? [''] : run.text.split(/(?<=\s)/);
        for (const w of words) {
            if (used + w.length > room && current.length) { out.push({ ...p, runs: current }); current = []; used = 0; }
            if (run.lineBreak) current.push(run); else if (w) { const last = current[current.length - 1]; if (last && !last.lineBreak && sameLook(last, run)) last.text += w; else current.push({ ...run, text: w }); }
            used += run.lineBreak ? perLine : w.length;
        }
    }
    if (current.length) out.push({ ...p, runs: current });
    return out.map((q, i) => i ? { ...q, bullet: q.bullet ? { ...q.bullet, char: '' } : undefined } : q);
}

const sameLook = (a: DeckRun, b: DeckRun) => a.bold === b.bold && a.italic === b.italic && a.underline === b.underline && a.size === b.size && a.color === b.color && a.font === b.font && a.link === b.link && a.superscript === b.superscript && a.subscript === b.subscript && a.strike === b.strike;

/** A flowing document's deck (see the module's description). */
async function flowDeck(content: OfficeContentNode[], metadata: OfficeMetadata, source: DeckSource, notes: { list: DeckParagraph[]; seen: Map<OfficeContentNode, number> }, unrepresentable: Set<string>): Promise<Deck> {
    const slides: DeckSlide[] = [];
    const lineH = BODY_SIZE * 1.25;
    const capacity = Math.floor(BODY_BOX.h / lineH);
    let title: DeckParagraph | undefined;
    let body: DeckParagraph[] = [];
    let used = 0;
    const titleItem = (p: DeckParagraph): DeckItem => ({ kind: 'text', box: TITLE_BOX, paragraphs: [{ ...p, runs: p.runs.map(r => ({ ...r, size: TITLE_SIZE, bold: true })) }], role: 'title', anchor: 'bottom' });
    // Whether the current title is on a slide yet: a heading with nothing under it has a slide of its own,
    // one whose table or picture follows is on theirs.
    let shown = false;
    const put = (items: DeckItem[]) => {
        if (slides.length >= MAX_SLIDES) { unrepresentable.add('slides past the deck\'s limit'); return; }
        slides.push({ items: [...(title ? [titleItem(title)] : []), ...items] });
        shown = true;
    };
    const flush = () => {
        if (!body.length && (!title || shown)) return;
        put(body.length ? [{ kind: 'text', box: title ? BODY_BOX : { ...BODY_BOX, y: TITLE_BOX.y, h: BODY_BOX.h + BODY_BOX.y - TITLE_BOX.y }, paragraphs: body, role: 'body' }] : []);
        body = []; used = 0;
    };
    const flushBody = () => { if (body.length) flush(); };
    // A table's slides: at most ten rows a slide, the first row (its header) repeated.
    const tableSlides = async (node: OfficeContentNode) => {
        flushBody();
        const t = await tableOf(node, source, notes);
        // The header: the first row and the rows its merges reach into (a header of two rows), at most three.
        let head = t.rows.length ? 1 : 0;
        for (let r = 0; r < head && r < t.rows.length; r++) for (const c of t.rows[r].cells) head = Math.max(head, Math.min(t.rows.length, r + c.rowSpan));
        if (head > 3) head = 0;
        const header = t.rows.slice(0, head), rest = t.rows.slice(head);
        // A merge running on past a slide's last row: its place on the next slide is an empty cell, so the
        // cells after it keep their columns (open: the merges begun on earlier slides, with the row each began in).
        let open: { row: number; cell: typeof rest[number]['cells'][number] }[] = [];
        for (let at = 0; at === 0 || at < rest.length; at += 10) {
            if (slides.length >= MAX_SLIDES) { unrepresentable.add('slides past the deck\'s limit'); break; }
            const rows = rest.slice(at, at + 10).map((row, i) => {
                const covered = open.filter(o => o.row + o.cell.rowSpan > at + i);
                if (!covered.length) return row;
                const filler = covered.map(o => ({ paragraphs: [], colSpan: o.cell.colSpan, rowSpan: 1, bold: false, col: o.cell.col ?? 0 }));
                return { cells: [...row.cells, ...filler].sort((a, b) => (a.col ?? 0) - (b.col ?? 0)) };
            });
            put([{ kind: 'table', box: BODY_BOX, rows: [...header, ...rows], widths: t.widths }]);
            const end = at + 10;
            open = open.filter(o => o.row + o.cell.rowSpan > end);
            rest.slice(at, end).forEach((row, i) => { for (const cell of row.cells) if (at + i + cell.rowSpan > end) open.push({ row: at + i, cell }); });
        }
        // The pictures in its cells, which a slide's table cannot hold, on slides of their own after it.
        for (const picture of picturesIn(node)) put([{ kind: 'picture', box: BODY_BOX, ...picture }]);
    };
    const add = (p: DeckParagraph) => {
        const size = BODY_SIZE;
        const lines = linesOf(p, BODY_BOX.w, size);
        if (used + lines > capacity && body.length) flush();
        if (lines > capacity) {
            for (const piece of cutParagraph(p, capacity - used, BODY_BOX.w, size)) { const n = linesOf(piece, BODY_BOX.w, size); if (used + n > capacity && body.length) flush(); body.push(piece); used += n; }
            return;
        }
        body.push(p);
        used += lines;
    };
    // The title slide: the document's title (its first top heading, or its metadata's), with what follows it until the next heading.
    let first = true;
    const visit = async (nodes: OfficeContentNode[], depth: number) => {
        for (const node of nodes) {
            const kept = await source.keep(node);
            if (kept === false) continue;
            if (node.type === 'heading') {
                const level = (node.metadata as HeadingMetadata | undefined)?.level ?? 1;
                const p = await paragraphOf(node, source, notes, kept);
                if (level <= 2) { flush(); title = p; shown = false; first = false; continue; }
                add({ ...p, runs: p.runs.map(r => ({ ...r, bold: true })) });
                continue;
            }
            if (node.type === 'sheet') {
                // A sheet: its name the title, its rows a table, then its charts and pictures.
                flush();
                const name = (node.metadata as { sheetName?: string } | undefined)?.sheetName;
                title = name ? { runs: [{ text: name }] } : title;
                shown = false; first = false;
                if ((node.children ?? []).some(c => c.type === 'row')) await tableSlides(node);
                if (depth < 16) await visit((node.children ?? []).filter(c => c.type !== 'row'), depth + 1);
                flush();
                continue;
            }
            if (first && !title && metadata.title) { title = { runs: [{ text: metadata.title }] }; }
            first = false;
            if (['paragraph', 'list', 'code', 'definitionTerm', 'definitionDescription', 'admonition'].includes(node.type)) {
                if (node.type === 'admonition' && depth < 16) { await visit(node.children ?? [], depth + 1); continue; }
                const p = await paragraphOf(node, source, notes, kept);
                if (p.runs.some(r => r.text.trim())) add(p);
                // Its pictures on slides of their own.
                for (const picture of picturesIn(node)) { flushBody(); put([{ kind: 'picture', box: BODY_BOX, ...picture }]); }
                continue;
            }
            if (node.type === 'image') {
                for (const picture of picturesIn(node)) { flushBody(); put([{ kind: 'picture', box: BODY_BOX, ...picture }]); }
                continue;
            }
            if (node.type === 'table') {
                await tableSlides(node);
                continue;
            }
            if (node.type === 'chart') {
                if (!source.includeCharts) continue;
                const table = source.chartTable(node);
                if (!table) { unrepresentable.add('charts without their data'); continue; }
                flushBody();
                const t = await tableOf(table, source, notes);
                put([{ kind: 'table', box: BODY_BOX, rows: t.rows, widths: t.widths }]);
                continue;
            }
            if (node.type === 'slide' || node.type === 'page') {
                flush();
                const before = slides.length;
                if (depth < 16) await visit(node.children ?? [], depth + 1);
                flush();
                // A slide's speaker notes go with the last slide its content made (one of their own if it made none).
                const own: DeckParagraph[] = [];
                for (const note of node.notes ?? []) { if (await source.keep(note) === false) continue; for (const block of note.children ?? []) own.push(await paragraphOf(block, source, notes)); }
                if (own.length) {
                    if (slides.length === before) put([]);
                    const last = slides[slides.length - 1];
                    if (last) last.notes = [...(last.notes ?? []), ...own];
                }
                continue;
            }
            if (depth < 16 && node.children?.length) await visit(node.children, depth + 1);
        }
    };
    await visit(content, 0);
    flush();
    if (!slides.length) slides.push({ items: [titleItem({ runs: [{ text: metadata.title ?? '' }] })] });
    slides.push(...notesSlides(notes.list, FLOW_W, FLOW_H));
    return { width: FLOW_W, height: FLOW_H, slides, unrepresentable };
}

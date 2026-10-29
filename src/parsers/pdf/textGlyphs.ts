/**
 * A page's glyphs, read from its operator list: where each is drawn, how far it truly advances, the
 * rendering mode it is drawn in and the colour it is filled with, and the text items pdf.js made of them
 * (see measureItems).
 *
 * pdf.js's text items are not always where their glyphs are. Its text extraction gives no advance to a
 * glyph whose text holds a nonspacing mark (its test for a "zero-width diacritic" matches a mark anywhere
 * in the glyph's text, so a conjunct "स्" or a cluster "ते" counts), nor to a format character: such a
 * glyph's item ends short, the items after it in the same show are drawn too far left, and the gap pdf.js
 * measures to the next glyph placed apart (a cluster LibreOffice places with its own `Tm`) is taken from
 * the stale point, so a blank item is put over glyphs the page draws: "नमस् ते" for "नमस्ते". Here every
 * glyph is placed twice, as pdf.js's text extraction places it (to know which item it is) and as it is
 * drawn (to know where it is).
 *
 * The graphics state is mirrored as the colour reader mirrored it: `q`/`Q`, the page's `cm`s and a form
 * XObject's `/Matrix` compose the CTM, and the text state (font, size, spacing, scaling, rise, rendering
 * mode) and the colours are saved and restored with it.
 *
 * @module parsers/pdf/textGlyphs
 */

import { annotationMatrix, identityMatrix, mulMatrix as mul, toMatrix6 as toMat6 } from './geometry.js';

/** What the walk needs to know of a font: its matrix (glyph space to text space) and its writing. */
export interface GlyphFont {
    fontMatrix?: ArrayLike<number>;
    vertical?: boolean;
    /**
     * Its codes read by their glyphs' names (see glyphNames), as the page's text items are normalized: a
     * glyph of such a code that pdf.js read as the code's own character is read so.
     */
    codeTexts?: ReadonlyMap<number, string>;
}

/** One show of text: its start point in user space, the rendering mode and the colour it paints in. */
export interface TextShow {
    x: number;
    y: number;
    /** The colour the glyphs are painted with (the stroke colour for stroke-only modes), if one is known. */
    color?: string;
    /** The text rendering mode (`Tr`): 3 and 7 draw nothing. */
    mode: number;
}

/** A glyph is a space (its text is blank). */
const SPACE = 1;
/** pdf.js's text extraction gives the glyph no advance (see the module's note). */
const LOST = 2;
/** The glyph is drawn invisibly (rendering mode 3 or 7). */
const INVISIBLE = 4;

/**
 * A page's glyphs, in the order they are drawn, as parallel arrays (a page may hold many thousands). All
 * positions are in user space, the frame pdf.js gives its text items' `transform` in.
 */
export interface PageGlyphs {
    count: number;
    /** Where pdf.js's text extraction places the glyph's origin. */
    px: number[];
    py: number[];
    /** Where the glyph's origin is drawn, and where its own advance ends. */
    x0: number[];
    y0: number[];
    x1: number[];
    y1: number[];
    /** The glyph's em in user space (its font size through the text and current matrices). */
    em: number[];
    flags: number[];
    /** The text pdf.js's extraction reads of the glyph: none for one it reads as a blank or skips. */
    text: string[];
    /** The glyphs read by their names (see GlyphFont's codeTexts), by their index, with that text. */
    named: Map<number, string>;
    /** Every show, for the colour of the text (see pdfColor). */
    shows: TextShow[];
}

/**
 * A page with more glyphs than this keeps the first of them only (its later text is read as pdf.js reads
 * it): a crafted content stream that shows millions of glyphs cannot make the walk hold them all.
 */
const MAX_GLYPHS = 500_000;

/** pdf.js's own test of a glyph's text (fonts_utils' SpecialCharRegExp): a blank, a nonspacing mark, a format character. */
const SPECIAL_CHAR = /^(\s)|(\p{Mn})|(\p{Cf})$/u;

/** Letters of the right-to-left scripts: a text holding one pdf.js may have reordered. */
const RTL = /[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/;

/** The control characters, a glyph's text where its font lost what it is. */
const CONTROL = /[\u0000-\u001f\u007f]/;

/** Graphics state the walk keeps across `q`/`Q` and forms. */
interface State {
    ctm: number[];
    fill: string | undefined;
    stroke: string | undefined;
    mode: number;
    charSpacing: number;
    wordSpacing: number;
    hScale: number;
    leading: number;
    rise: number;
    font: GlyphFont | undefined;
    fontSize: number;
}

/** Applies the text-space translation `[x, y]` to a matrix, as pdf.js's `TextState` does. */
function translate(m: number[], x: number, y: number): number[] {
    return [m[0], m[1], m[2], m[3], m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}

/**
 * Walks an already-fetched operator list and returns its glyphs and shows (see PageGlyphs).
 *
 * @param fontOf - the font a `setFont` names (by pdf.js's id for it), or undefined when it is not known:
 *   its glyphs are then placed with the usual matrix of a thousand units an em.
 */
export function collectGlyphs(ops: any, OPS: any, fontOf: (id: string) => GlyphFont | undefined): PageGlyphs {
    const out: PageGlyphs = { count: 0, px: [], py: [], x0: [], y0: [], x1: [], y1: [], em: [], flags: [], text: [], named: new Map(), shows: [] };
    const fnArray: number[] | undefined = ops?.fnArray;
    const argsArray: any[] | undefined = ops?.argsArray;
    if (!fnArray || !argsArray) return out;

    let s: State = {
        ctm: identityMatrix(), fill: undefined, stroke: undefined, mode: 0,
        charSpacing: 0, wordSpacing: 0, hScale: 1, leading: 0, rise: 0, font: undefined, fontSize: 0,
    };
    const stack: State[] = [];
    // The text matrix as the page draws it, and as pdf.js's text extraction moves it: the two part only
    // within a show, at a glyph whose advance pdf.js leaves out, and meet again at each positioning operator.
    let tm = identityMatrix(), tlm = identityMatrix(), qtm = identityMatrix();
    const setLine = (m: number[]) => { tlm = m; tm = m.slice(); qtm = m.slice(); };
    const num = (v: unknown, fallback = 0) => { const n = +(v as number); return Number.isFinite(n) ? n : fallback; };
    // A point of text space (x along the line, the rise above it) in user space.
    const userX = (m: number[], rise: number) => s.ctm[0] * (m[2] * rise + m[4]) + s.ctm[2] * (m[3] * rise + m[5]) + s.ctm[4];
    const userY = (m: number[], rise: number) => s.ctm[1] * (m[2] * rise + m[4]) + s.ctm[3] * (m[3] * rise + m[5]) + s.ctm[5];

    for (let i = 0; i < fnArray.length; i++) {
        const a = argsArray[i];
        switch (fnArray[i]) {
            case OPS.save: stack.push({ ...s }); break;
            case OPS.restore: { const r = stack.pop(); if (r) s = r; break; }
            case OPS.transform: { const cm = toMat6(a); if (cm) s.ctm = mul(s.ctm, cm); break; }
            // A form XObject composes its /Matrix and localizes the graphics state.
            case OPS.paintFormXObjectBegin: {
                stack.push({ ...s });
                const fm = toMat6(a?.[0]);
                if (fm) s.ctm = mul(s.ctm, fm);
                break;
            }
            case OPS.paintFormXObjectEnd: { const r = stack.pop(); if (r) s = r; break; }
            // An annotation's appearance, in its own place on the page (see annotationMatrix).
            case OPS.beginAnnotation: stack.push({ ...s }); s = { ...s, ctm: annotationMatrix(a) }; break;
            case OPS.endAnnotation: { const r = stack.pop(); if (r) s = r; break; }

            // Fill / stroke color: all spaces arrive pre-normalized to an *RGBColor hex string.
            case OPS.setFillRGBColor: s.fill = typeof a?.[0] === 'string' ? a[0] : undefined; break;
            case OPS.setFillColorN:
            case OPS.setFillTransparent: s.fill = undefined; break;
            case OPS.setStrokeRGBColor: s.stroke = typeof a?.[0] === 'string' ? a[0] : undefined; break;
            case OPS.setStrokeColorN:
            case OPS.setStrokeTransparent: s.stroke = undefined; break;

            // Text state (mirror of pdf.js's TextState).
            case OPS.setFont: s.font = typeof a?.[0] === 'string' ? fontOf(a[0]) : undefined; s.fontSize = num(a?.[1]); break;
            case OPS.setCharSpacing: s.charSpacing = num(a?.[0]); break;
            case OPS.setWordSpacing: s.wordSpacing = num(a?.[0]); break;
            case OPS.setHScale: s.hScale = num(a?.[0], 100) / 100; break;
            case OPS.setLeading: s.leading = num(a?.[0]); break;
            case OPS.setTextRise: s.rise = num(a?.[0]); break;
            case OPS.setTextRenderingMode: s.mode = num(a?.[0]); break;
            case OPS.beginText: setLine(identityMatrix()); break;
            case OPS.setTextMatrix: { const tmx = toMat6(a?.[0]) ?? toMat6(a); if (tmx) setLine(tmx); break; }
            case OPS.moveText: setLine(translate(tlm, num(a?.[0]), num(a?.[1]))); break;
            case OPS.setLeadingMoveText: s.leading = -num(a?.[1]); setLine(translate(tlm, num(a?.[0]), num(a?.[1]))); break;
            case OPS.nextLine: setLine(translate(tlm, 0, -s.leading)); break;

            case OPS.showText: {
                const glyphs: unknown[] = Array.isArray(a?.[0]) ? a[0] : [];
                const invisible = s.mode === 3 || s.mode === 7;
                const color = (s.mode === 1 || s.mode === 5) ? s.stroke : s.fill;
                out.shows.push({ x: userX(tm, s.rise), y: userY(tm, s.rise), ...(color ? { color } : {}), mode: s.mode });
                if (s.font?.vertical) break; // (placed down the page: not measured here)
                const fm0 = num(s.font?.fontMatrix?.[0], 0.001) || 0.001;
                const size = s.fontSize, h = s.hScale;
                // The em through the text and current matrices, for every glyph of the show alike.
                const m = mul(s.ctm, tm);
                const em = Math.abs(size) * Math.hypot(m[2], m[3]);
                // What the last item was: a glyph pdf.js skips whole (a format character), any other glyph, or none.
                let last: 'skipped' | 'glyph' | 'none' = 'none';
                for (const g of glyphs) {
                    if (typeof g === 'number') {
                        if (!Number.isFinite(g) || g === 0) continue;
                        const extra = -g * size / 1000 * h;
                        tm = translate(tm, extra, 0);
                        // (pdf.js adds a number to the glyph before it; one with no glyph before it in the show,
                        // or after another number, comes with the character spacing once more; one after a
                        // glyph it skips whole is lost.)
                        if (last === 'none') qtm = translate(qtm, extra + s.charSpacing * h, 0);
                        else if (last === 'glyph') qtm = translate(qtm, extra, 0);
                        last = 'none';
                        continue;
                    }
                    const glyph = g as { unicode?: unknown; width?: unknown; isSpace?: unknown; originalCharCode?: unknown } | null;
                    if (!glyph || typeof glyph !== 'object') continue;
                    const unicode = typeof glyph.unicode === 'string' ? glyph.unicode : '';
                    const width = num(glyph.width) * fm0 * size;
                    const spacing = s.charSpacing + (glyph.isSpace ? s.wordSpacing : 0);
                    const special = SPECIAL_CHAR.exec(unicode);
                    // (A space: blank text that is no control character, which is a glyph whose text a font lost, drawn
                    // as any other: pdf.js reads a vertical tab as a blank, LibreOffice's "त्र" was one.)
                    const blank = !!unicode && !unicode.trim() && !CONTROL.test(unicode);
                    if (out.count < MAX_GLYPHS) {
                        out.px.push(userX(qtm, s.rise)); out.py.push(userY(qtm, s.rise));
                        out.x0.push(userX(tm, s.rise)); out.y0.push(userY(tm, s.rise));
                        const end = translate(tm, width * h, 0);
                        out.x1.push(userX(end, s.rise)); out.y1.push(userY(end, s.rise));
                        out.em.push(em);
                        out.flags.push((blank ? SPACE : 0) | (special?.[2] || special?.[3] ? LOST : 0) | (invisible ? INVISIBLE : 0));
                        out.text.push(special?.[1] || special?.[3] ? '' : unicode);
                        const code = glyph.originalCharCode;
                        const named = typeof code === 'number' ? s.font?.codeTexts?.get(code) : undefined;
                        if (named !== undefined && unicode === String.fromCharCode(code as number)) out.named.set(out.count, named);
                        out.count++;
                    }
                    tm = translate(tm, (width + spacing) * h, 0);
                    if (special?.[3]) { last = 'skipped'; continue; }
                    const wordSpacing = glyph.originalCharCode === 0x20 ? s.wordSpacing : 0;
                    qtm = translate(qtm, ((special?.[2] ? 0 : width) + s.charSpacing + wordSpacing) * h, 0);
                    last = 'glyph';
                }
                break;
            }
        }
    }
    return out;
}

/** A pdf.js text content item, as far as measuring it needs. */
export interface MeasuredItem {
    type?: string;
    str?: string;
    dir?: string;
    transform?: number[];
    width?: number;
    /** Its place and width are its glyphs' as the page draws them (see measureItems). */
    measured?: boolean;
    /** Every glyph of it is drawn invisibly (rendering mode 3 or 7). */
    invisible?: boolean;
}

/** How near (in user space) pdf.js's place for an item is to the glyph it starts at. */
const NEAR = 0.05;

/**
 * A page's text items with what their glyphs say of them: each item matched to the glyph it starts at
 * (the first, in the order drawn, at the place pdf.js gave the item, after the glyphs of the items before
 * it), and then
 *
 * - an item of level text is given the place and width its glyphs are drawn over, where pdf.js's are not
 *   theirs (a glyph in it pdf.js gave no advance, or one before it in its show): marked `measured`, as
 *   every item matched is, so no later guess at a lost width moves it;
 * - an item whose glyphs are all drawn invisibly is marked `invisible`;
 * - an item with a glyph read by its name (see GlyphFont's codeTexts) reads that glyph so, where its text
 *   is its glyphs' in order (see named): "ni!ht" is "nicht";
 * - a blank item pdf.js put before a glyph, for the gap it measured to it, is left out where the gap is
 *   none (the glyph drawn before it ends within a tenth of an em of it, and no space glyph is between):
 *   the gap was measured from where pdf.js lost the widths of the glyphs between. Kept, it is placed over
 *   the gap as the page has it.
 *
 * Items not matched (no operator list, a page past its glyphs' bound, a place the walk does not give as
 * pdf.js does) are left as they are.
 */
export function measureItems<T extends MeasuredItem>(items: T[], glyphs: PageGlyphs | undefined): T[] {
    if (!glyphs || !glyphs.count) return items;
    const { px, py, x0, y0, x1, em, flags } = glyphs;
    let text = glyphs.text;
    const key = (x: number, y: number) => `${Math.round(x / NEAR)}:${Math.round(y / NEAR)}`;
    const at = new Map<string, number[]>();
    for (let k = 0; k < glyphs.count; k++) {
        const id = key(px[k], py[k]);
        const list = at.get(id);
        if (list) list.push(k); else at.set(id, [k]);
    }
    const isText = (it: T) => typeof it.str === 'string' && Array.isArray(it.transform) && it.transform.length >= 6 && it.transform.every(Number.isFinite);
    const near = (k: number, x: number, y: number) => Math.abs(px[k] - x) <= NEAR && Math.abs(py[k] - y) <= NEAR;
    /**
     * The first glyph at an item's place from `from` on, else the first there at all (not taken in order).
     * A few of each place's glyphs are looked at (a page drawing a glyph many times at one place is read in
     * time in its length, not its square).
     */
    const find = (x: number, y: number, from: number): { k: number; inOrder: boolean } | undefined => {
        let best = -1, any = -1;
        const cx = Math.round(x / NEAR), cy = Math.round(y / NEAR);
        for (let dx = -1; dx <= 1; dx++) for (let dy = -1; dy <= 1; dy++) {
            const list = at.get(`${cx + dx}:${cy + dy}`);
            if (!list) continue;
            let lo = 0, hi = list.length;
            while (lo < hi) { const mid = (lo + hi) >> 1; if (list[mid] < from) lo = mid + 1; else hi = mid; }
            for (let j = lo; j < Math.min(list.length, lo + 4); j++) if (near(list[j], x, y)) { if (best < 0 || list[j] < best) best = list[j]; break; }
            for (let j = 0; j < Math.min(lo, 4); j++) if (near(list[j], x, y)) { if (any < 0 || list[j] < any) any = list[j]; break; }
        }
        return best >= 0 ? { k: best, inOrder: true } : any >= 0 ? { k: any, inOrder: false } : undefined;
    };

    // Each item of text (not a blank) matched to its first glyph, and whether in order (after the glyphs of
    // the items before it).
    const first = new Array<number>(items.length).fill(-1);
    const inOrder = new Array<boolean>(items.length).fill(false);
    let next = 0;
    items.forEach((it, i) => {
        if (!isText(it) || !it.str!.trim()) return;
        const hit = find(it.transform![4], it.transform![5], next);
        if (!hit) return;
        first[i] = hit.k;
        inOrder[i] = hit.inOrder;
        if (hit.inOrder) next = hit.k + 1;
    });
    // Where each item matched in order has its glyphs end: at the next one's first glyph (one matched out of
    // order is its first glyph alone).
    const ends = new Array<number>(items.length).fill(glyphs.count);
    let after = glyphs.count;
    for (let i = items.length - 1; i >= 0; i--) {
        if (first[i] < 0) continue;
        if (!inOrder[i]) { ends[i] = first[i] + 1; continue; }
        ends[i] = after;
        after = first[i];
    }

    // The glyphs read by their names, in the items' text (and in what each glyph reads, for the spacing below).
    if (glyphs.named.size) {
        text = text.slice();
        items = items.map((it, i) => {
            if (first[i] < 0) return it;
            const read = named(it, first[i], ends[i]);
            if (!read) return it;
            for (const k of read.glyphs) text[k] = glyphs.named.get(k)!;
            return { ...it, str: read.str };
        });
    }

    const level = (t: number[]) => t[0] > 0 && Math.abs(t[1]) <= 1e-6 * t[0] && Math.abs(t[2]) <= 1e-6 * t[0];
    const out: T[] = [];
    items.forEach((it, i) => {
        if (!isText(it)) { out.push(it); return; }
        const t = it.transform!;
        if (!it.str!.trim()) {
            const blank = blankOver(it, i);
            if (blank) out.push(blank);
            return;
        }
        const g = first[i];
        if (g < 0) { out.push(it); return; }
        // The item's glyphs: from its first, on its line, as far as pdf.js's width for it reaches (a glyph it
        // gives no advance is not placed by it: it is the item's, wherever the page draws it).
        const reach = t[4] + Math.max(0, it.width || 0) + NEAR;
        let lo = Infinity, hi = -Infinity, lost = false, seen = 0, hidden = 0, last = g;
        for (let k = g; k < ends[i] && k < g + 4096; k++) {
            if (Math.abs(py[k] - t[5]) > NEAR || (!(flags[k] & LOST) && (px[k] > reach || px[k] < t[4] - NEAR))) break;
            last = k;
            if (flags[k] & LOST) lost = true;
            if (flags[k] & SPACE) continue;
            seen++;
            if (flags[k] & INVISIBLE) hidden++;
            lo = Math.min(lo, x0[k]);
            hi = Math.max(hi, x1[k]);
        }
        const moved = Math.abs(x0[g] - t[4]) > NEAR || Math.abs(y0[g] - t[5]) > NEAR;
        // (Its glyphs reach past its width: a glyph pdf.js read as a blank, its advance left out.)
        const short = hi > t[4] + Math.max(0, it.width || 0) + NEAR;
        const item: T = { ...it, measured: true };
        if (seen && hidden === seen) item.invisible = true;
        // (One matched out of order is known by its first glyph alone: whether it is seen.)
        if (!inOrder[i]) { delete item.measured; out.push(item); return; }
        // Seen and unseen glyphs in one item (pdf.js does not part its items where the rendering mode changes):
        // parted where it does, so invisible text is never read as a visible run's. (Left whole, and seen,
        // where its text cannot be told glyph by glyph.)
        if (hidden && hidden < seen) {
            const parts = level(t) ? partedByVisibility(item, g, last) : undefined;
            if (parts) { out.push(...parts); return; }
            delete item.invisible;
        }
        // Its spaces where the page has them, where pdf.js measured a gap from a glyph whose advance it lost.
        if ((lost || short || /^\s/.test(it.str!)) && /\s/.test(it.str!) && level(t)) {
            const spaced = respaced(it, g, last);
            if (spaced !== undefined) item.str = spaced;
        }
        // Placed anew where pdf.js lost a glyph's advance, placed it elsewhere, or left it short.
        if (level(t) && (lost || moved || short) && Number.isFinite(lo) && hi >= lo) {
            item.transform = [t[0], t[1], t[2], t[3], lo, y0[g]];
            item.width = hi - lo;
        }
        out.push(item);
    });
    return out;

    /**
     * An item's text with its glyphs read by their names (see GlyphFont's codeTexts): its glyphs from `g`
     * (before `end`) walked along its text, each glyph's text found where the text is at (as pdf.js reads it,
     * or normalized as pdf.js normalizes an item's), its blanks kept, and a glyph read by its name read so.
     * Undefined where none is, or its text is not its glyphs' in order (read right to left, composed by
     * normalizing, or another glyph's than the walk meets): the item is left as pdf.js reads it.
     */
    function named(it: T, g: number, end: number): { str: string; glyphs: number[] } | undefined {
        const str = it.str!;
        if (it.dir === 'rtl' || RTL.test(str)) return undefined;
        let out = '', p = 0;
        const read: number[] = [];
        for (let k = g; k < end && p < str.length; k++) {
            const own = text[k].replace(/\s/g, '');
            if (!own) continue;
            while (p < str.length && /\s/.test(str[p])) out += str[p++];
            if (p >= str.length) break;
            let found = str.startsWith(own, p) ? own : undefined;
            if (found === undefined) { const normal = own.normalize('NFKC'); if (normal && str.startsWith(normal, p)) found = normal; }
            if (found === undefined) return undefined;
            const name = glyphs!.named.get(k);
            if (name !== undefined) read.push(k);
            out += name ?? found;
            p += found.length;
        }
        if (!read.length || str.slice(p).trim()) return undefined;
        return { str: out + str.slice(p), glyphs: read };
    }

    /** The glyph drawn before glyph `g` on its line, past any space glyph (-1 for none), and whether one was passed. */
    function before(g: number): { p: number; spaced: boolean } {
        let spaced = false;
        for (let k = g - 1; k >= 0 && k >= g - 64; k--) {
            if (Math.abs(y0[k] - y0[g]) > NEAR) break;
            if (flags[k] & SPACE) { spaced = true; continue; }
            return { p: k, spaced };
        }
        return { p: -1, spaced };
    }

    /** Whether the page has a space before glyph `g`: a space glyph, a gap of a tenth of an em, or no glyph before it on its line. */
    function spaceBefore(g: number): boolean {
        const { p, spaced } = before(g);
        return spaced || p < 0 || x0[g] - x1[p] >= 0.1 * em[g];
    }

    /** Whether the page has a space between glyphs `p` and `n` (`p` before `n` on their line): a space glyph or a gap of a tenth of an em. */
    function spaceBetween(p: number, n: number): boolean {
        for (let k = p + 1; k < n; k++) if (flags[k] & SPACE) return true;
        return x0[n] - x1[p] >= 0.1 * em[n];
    }

    /**
     * An item's text with its spaces where the page has them (each blank of it, between two of its glyphs or
     * before the first, kept where a space glyph or a gap is there): the glyphs `g` to `last` read as pdf.js
     * reads them, in order. Undefined where its text is not theirs in that order (read right to left, or
     * composed by normalizing), or no blank of it moves.
     */
    function respaced(it: T, g: number, last: number): string | undefined {
        const str = it.str!;
        const counted = charsBefore(str, it.dir, g, last);
        if (!counted) return undefined;
        const before = counted.slice(0, -1);
        let out = '', read = 0, changed = false, j = 0;
        for (const part of str.split(/(\s+)/)) {
            if (!part) continue;
            if (/\s/.test(part)) {
                // The glyph whose text follows (pdf.js puts the blank before the text of the glyph after it), and
                // the glyph drawn before that one, past space glyphs. (The glyphs are walked once, as the text is.)
                while (j < before.length && (before[j] < read || (before[j] === read && !text[g + j].trim()))) j++;
                const at = j < before.length && before[j] === read ? j : -1;
                let keep = true;
                if (at === 0) keep = spaceBefore(g);
                else if (at > 0) {
                    let p = g + at - 1;
                    while (p > g && flags[p] & SPACE) p--;
                    keep = spaceBetween(p, g + at);
                }
                if (keep) out += part; else changed = true;
                continue;
            }
            out += part;
            read += part.length;
        }
        return changed ? out : undefined;
    }

    /**
     * The characters (not blanks) of an item's text read before each of its glyphs `g` to `last`, and in all;
     * undefined where its text is not its glyphs' in order (see respaced).
     */
    function charsBefore(str: string, dir: string | undefined, g: number, last: number): number[] | undefined {
        if (dir === 'rtl' || RTL.test(str)) return undefined;
        const before: number[] = [];
        let n = 0;
        for (let k = g; k <= last; k++) { before.push(n); n += text[k].replace(/\s/g, '').length; }
        before.push(n);
        return n === str.replace(/\s/g, '').length ? before : undefined;
    }

    /**
     * An item whose glyphs are partly drawn invisibly, as one item a run of glyphs drawn alike, each placed
     * over its own glyphs and marked `invisible` where they are. Undefined where its text cannot be parted
     * glyph by glyph.
     */
    function partedByVisibility(item: T, g: number, last: number): T[] | undefined {
        const str = item.str!;
        const before = charsBefore(str, item.dir, g, last);
        if (!before) return undefined;
        // Where each run of glyphs drawn alike starts (a space glyph goes with the run it is in).
        const starts: number[] = [g];
        let state = -1;
        for (let k = g; k <= last; k++) {
            if (flags[k] & SPACE) continue;
            const hiddenHere = flags[k] & INVISIBLE ? 1 : 0;
            if (state >= 0 && hiddenHere !== state) starts.push(k);
            state = hiddenHere;
        }
        // The text's index of each character (not a blank) read, so a run's text is cut before the blanks ahead of it.
        const index: number[] = [];
        for (let c = 0; c < str.length; c++) if (!/\s/.test(str[c])) index.push(c);
        index.push(str.length);
        const parts: T[] = [];
        for (let r = 0; r < starts.length; r++) {
            const from = starts[r], to = r + 1 < starts.length ? starts[r + 1] - 1 : last;
            const a = r === 0 ? 0 : index[before[from - g]], b = r + 1 < starts.length ? index[before[starts[r + 1] - g]] : str.length;
            const piece = str.slice(a, b);
            if (!piece) continue;
            let lo = Infinity, hi = -Infinity, unseen = false;
            for (let k = from; k <= to; k++) {
                if (flags[k] & SPACE) continue;
                lo = Math.min(lo, x0[k]); hi = Math.max(hi, x1[k]);
                unseen = !!(flags[k] & INVISIBLE);
            }
            if (!Number.isFinite(lo)) continue;
            const t = item.transform!;
            const part: T = { ...item, str: piece, transform: [t[0], t[1], t[2], t[3], lo, y0[from]], width: hi - lo, measured: true };
            if (unseen) part.invisible = true; else delete part.invisible;
            parts.push(part);
        }
        return parts.length ? parts : undefined;
    }

    /** A blank item: left out, or placed over the gap before the glyph it is put before (see above). */
    function blankOver(it: T, i: number): T | undefined {
        const t = it.transform!;
        let j = i + 1;
        while (j < items.length && !isText(items[j])) j++;
        const g = j < items.length ? first[j] : -1;
        if (g < 0 || !level(t)) return it;
        const { p } = before(g);
        if (p < 0) return it;
        if (!spaceBefore(g)) return undefined;
        return { ...it, transform: [t[0], t[1], t[2], t[3], x1[p], t[5]], width: Math.max(0, x0[g] - x1[p]), measured: true };
    }
}

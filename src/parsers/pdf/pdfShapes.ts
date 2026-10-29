/**
 * The lines and shapes a page draws, from its operator list, for what text alone does not show: a list
 * item's bullet drawn as a small circle or square (as Chromium draws `<ul>` markers), and a table's
 * ruling lines and cell borders. pdf.js gives each path as one `constructPath` operator: its paint
 * operator (fill, stroke, both, or none), its data in the user space of the moment (`DrawOPS`: move,
 * line, cubic and quadratic curves, close), and its box. The current transformation is tracked as
 * `collectGlyphs` (textGlyphs) tracks it (`q`/`Q`, `cm`, form XObjects), so a shape lands in the same viewport
 * space as the text runs. Bounded: at most MAX_SHAPES shapes a page, and MAX_SEGMENTS straight edges.
 *
 * @module parsers/pdf/pdfShapes
 */

import { annotationMatrix, identityMatrix, mulMatrix as mul, toMatrix6 as toMat6 } from './geometry.js';
import type { RawRun } from './pdfTypes.js';

/** At most this many shapes are kept on a page, and this many straight edges in all. */
const MAX_SHAPES = 20_000;
const MAX_SEGMENTS = 50_000;
/** At most this many straight edges are kept for one shape (a table's grid can be one path). */
const MAX_SHAPE_SEGMENTS = 4_096;

/** A drawn shape, in viewport space (y down). */
export interface PdfShape {
    /** Its box. */
    x: number;
    y: number;
    width: number;
    height: number;
    /** Painted by filling, stroking, or both. */
    filled: boolean;
    stroked: boolean;
    /** The colours it is painted in (`#rrggbb`), where known. */
    fill?: string;
    stroke?: string;
    /** The stroke's width, in viewport units. */
    lineWidth: number;
    /** Whether it has curves (a disc, a ring), and how many straight edges. */
    curved: boolean;
    /** Its straight edges, [x1, y1, x2, y2], at most MAX_SHAPE_SEGMENTS. */
    segments: [number, number, number, number][];
    /**
     * The box of the clipping path it was painted through, where the page clipped it (the boxes of the
     * paths it set as clips, and of the forms it was drawn in, met): what it paints is inside it. A browser
     * draws a box's left border by clipping to the border's strip and filling the whole box.
     */
    clip?: { x: number; y: number; width: number; height: number };
    /** Filled see-through: at a fill opacity under one (`ca`), or through a soft mask (a browser's shadow). */
    seeThrough?: boolean;
}

/** A clipping box, [minX, minY, maxX, maxY] in viewport space, or none (the whole page). */
type ClipBox = [number, number, number, number] | undefined;

/** Where two clipping boxes meet (an empty box where they do not). */
function meet(a: ClipBox, b: [number, number, number, number]): [number, number, number, number] {
    if (!a) return b;
    const minX = Math.max(a[0], b[0]), minY = Math.max(a[1], b[1]);
    return [minX, minY, Math.max(minX, Math.min(a[2], b[2])), Math.max(minY, Math.min(a[3], b[3]))];
}

/** A form's box ([x0, y0, x1, y1] in its space) in viewport space, through `m`: the box around its corners. */
function boxOf(m: number[], bbox: ArrayLike<number>): [number, number, number, number] {
    const xs: number[] = [], ys: number[] = [];
    for (const [x, y] of [[+bbox[0], +bbox[1]], [+bbox[2], +bbox[1]], [+bbox[0], +bbox[3]], [+bbox[2], +bbox[3]]]) { xs.push(m[0] * x + m[2] * y + m[4]); ys.push(m[1] * x + m[3] * y + m[5]); }
    return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
}

/** Walks an operator list fetched with drawing operators kept, and returns the shapes it paints. */
export function collectShapes(ops: any, viewportTransform: number[], OPS: any): PdfShape[] {
    const shapes: PdfShape[] = [];
    const fnArray: number[] | undefined = ops?.fnArray;
    const argsArray: any[] | undefined = ops?.argsArray;
    if (!fnArray || !argsArray || OPS.constructPath === undefined) return shapes;
    let ctm = identityMatrix();
    let fill: string | undefined, stroke: string | undefined, lineWidth = 1;
    // The clipping box, and whether the next path sets a clip (as `W` before its `n` or its paint); whether
    // a fill is see-through (its opacity, a soft mask); and whether what is drawn is a soft mask's own
    // contents (its group form, `/G`), which shape what is painted through the mask and paint nothing.
    let clip: ClipBox, clipNext = false, alpha = 1, masked = false, inMask = false;
    const stack: { ctm: number[]; fill?: string; stroke?: string; lineWidth: number; clip: ClipBox; alpha: number; masked: boolean; inMask: boolean }[] = [];
    let segmentsLeft = MAX_SEGMENTS;
    const fills = new Set([OPS.fill, OPS.eoFill, OPS.fillStroke, OPS.eoFillStroke, OPS.closeFillStroke, OPS.closeEOFillStroke]);
    const strokes = new Set([OPS.stroke, OPS.closeStroke, OPS.fillStroke, OPS.eoFillStroke, OPS.closeFillStroke, OPS.closeEOFillStroke]);
    for (let i = 0; i < fnArray.length && shapes.length < MAX_SHAPES; i++) {
        const a = argsArray[i];
        switch (fnArray[i]) {
            case OPS.save: stack.push({ ctm, fill, stroke, lineWidth, clip, alpha, masked, inMask }); break;
            case OPS.restore: { const s = stack.pop(); if (s) ({ ctm, fill, stroke, lineWidth, clip, alpha, masked, inMask } = s); break; }
            case OPS.transform: { const cm = toMat6(a); if (cm) ctm = mul(ctm, cm); break; }
            case OPS.paintFormXObjectBegin: {
                stack.push({ ctm, fill, stroke, lineWidth, clip, alpha, masked, inMask });
                const fm = toMat6(a?.[0]);
                if (fm) ctm = mul(ctm, fm);
                // (A form is clipped to its box.)
                const bbox = a?.[1];
                if (bbox && bbox.length === 4 && [0, 1, 2, 3].every(k => Number.isFinite(+bbox[k]))) clip = meet(clip, boxOf(mul(viewportTransform, ctm), bbox));
                break;
            }
            // A transparency group's form: pdf.js gives its box here, not with the form's matrix (which it
            // gives as the form's, next), and a soft mask's group its mask.
            case OPS.beginGroup: {
                stack.push({ ctm, fill, stroke, lineWidth, clip, alpha, masked, inMask });
                const group = a?.[0];
                if (group?.smask) inMask = true;
                const gm = toMat6(group?.matrix);
                const bbox = group?.bbox;
                if (bbox && bbox.length === 4 && [0, 1, 2, 3].every(k => Number.isFinite(+bbox[k]))) clip = meet(clip, boxOf(mul(viewportTransform, gm ? mul(ctm, gm) : ctm), bbox));
                break;
            }
            case OPS.endGroup: { const s = stack.pop(); if (s) ({ ctm, fill, stroke, lineWidth, clip, alpha, masked, inMask } = s); break; }
            case OPS.paintFormXObjectEnd: { const s = stack.pop(); if (s) ({ ctm, fill, stroke, lineWidth, clip, alpha, masked, inMask } = s); break; }
            // An annotation's appearance, in its own place on the page (see annotationMatrix).
            case OPS.beginAnnotation: stack.push({ ctm, fill, stroke, lineWidth, clip, alpha, masked, inMask }); ctm = annotationMatrix(a); break;
            case OPS.endAnnotation: { const s = stack.pop(); if (s) ({ ctm, fill, stroke, lineWidth, clip, alpha, masked, inMask } = s); break; }
            case OPS.clip: case OPS.eoClip: clipNext = true; break;
            case OPS.setGState: {
                for (const entry of Array.isArray(a?.[0]) ? a[0] : []) {
                    if (!Array.isArray(entry)) continue;
                    if (entry[0] === 'ca' && Number.isFinite(+entry[1])) alpha = +entry[1];
                    else if (entry[0] === 'SMask') masked = !!entry[1];
                }
                break;
            }
            case OPS.setFillRGBColor: fill = typeof a?.[0] === 'string' ? a[0] : undefined; break;
            case OPS.setFillColorN: case OPS.setFillTransparent: fill = undefined; break;
            case OPS.setStrokeRGBColor: stroke = typeof a?.[0] === 'string' ? a[0] : undefined; break;
            case OPS.setStrokeColorN: case OPS.setStrokeTransparent: stroke = undefined; break;
            case OPS.setLineWidth: { const w = +a?.[0]; if (Number.isFinite(w) && w >= 0) lineWidth = w; break; }
            case OPS.constructPath: {
                const paint = a?.[0];
                const filled = fills.has(paint), stroked = strokes.has(paint);
                const clips = clipNext;
                clipNext = false;
                if (!filled && !stroked && !clips) break;
                const data = a?.[1]?.[0];
                if (!data || typeof data.length !== 'number') break;
                // The path's points in viewport space: user space, then the CTM, then the viewport.
                const m = mul(viewportTransform, ctm);
                const px = (x: number, y: number) => m[0] * x + m[2] * y + m[4], py = (x: number, y: number) => m[1] * x + m[3] * y + m[5];
                let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
                const take = (x: number, y: number) => { if (x < minX) minX = x; if (x > maxX) maxX = x; if (y < minY) minY = y; if (y > maxY) maxY = y; };
                const segments: [number, number, number, number][] = [];
                let curved = false, cx = 0, cy = 0, sx = 0, sy = 0;
                const line = (x: number, y: number) => {
                    // (A clipping path's edges are not kept: its box is.)
                    if ((filled || stroked) && segments.length < MAX_SHAPE_SEGMENTS && segmentsLeft > 0) { segments.push([cx, cy, x, y]); segmentsLeft--; }
                    cx = x; cy = y; take(x, y);
                };
                for (let k = 0; k < data.length;) {
                    const op = data[k++];
                    if (op === 0) { cx = sx = px(data[k], data[k + 1]); cy = sy = py(data[k], data[k + 1]); take(cx, cy); k += 2; }
                    else if (op === 1) { line(px(data[k], data[k + 1]), py(data[k], data[k + 1])); k += 2; }
                    else if (op === 2) {
                        curved = true;
                        for (let j = 0; j < 6; j += 2) take(px(data[k + j], data[k + j + 1]), py(data[k + j], data[k + j + 1]));
                        cx = px(data[k + 4], data[k + 5]); cy = py(data[k + 4], data[k + 5]); k += 6;
                    } else if (op === 3) {
                        curved = true;
                        for (let j = 0; j < 4; j += 2) take(px(data[k + j], data[k + j + 1]), py(data[k + j], data[k + j + 1]));
                        cx = px(data[k + 2], data[k + 3]); cy = py(data[k + 2], data[k + 3]); k += 4;
                    } else if (op === 4) { if (cx !== sx || cy !== sy) line(sx, sy); }
                    else break;
                }
                if (!Number.isFinite(minX) || !Number.isFinite(minY)) break;
                if ((filled || stroked) && !inMask) {
                    // A stroke's width in viewport space, scaled as the CTM scales (its mean).
                    const scale = Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1;
                    shapes.push({
                        x: minX, y: minY, width: maxX - minX, height: maxY - minY, filled, stroked,
                        ...(filled && fill ? { fill } : {}), ...(stroked && stroke ? { stroke } : {}),
                        lineWidth: lineWidth * scale, curved, segments,
                        ...(clip ? { clip: { x: clip[0], y: clip[1], width: clip[2] - clip[0], height: clip[3] - clip[1] } } : {}),
                        ...(filled && (alpha < 1 || masked) ? { seeThrough: true } : {}),
                    });
                }
                // (The path clips what is painted after it, not itself.)
                if (clips) clip = meet(clip, [minX, minY, maxX, maxY]);
                break;
            }
        }
    }
    return shapes;
}

/**
 * A list item's bullet drawn as a shape (Chromium draws `<ul>` markers as a small disc, ring or square,
 * not as a glyph) as a run of its glyph: `•` a filled round one, `◦` a ring, `▪` a filled square, `□` an
 * outlined one. A shape is a bullet where it is about the size of a bullet (a sixth to three fifths of
 * the text's size, as wide as it is high), its middle within the x-height of a line of text that starts
 * within three of its ems to its right. The run takes that line's place, so the line reads `• Apple`, as
 * a bullet glyph's would: the bullets were lost, and the items read as paragraphs.
 */
export function shapeBullets(shapes: readonly PdfShape[] | undefined, runs: readonly RawRun[]): RawRun[] {
    if (!shapes?.length) return [];
    // Runs by their baseline, in buckets of 4 units, so each shape looks at the few lines near it.
    const buckets = new Map<number, RawRun[]>();
    for (const r of runs) {
        if (r.angle !== 0 || !r.text.trim()) continue;
        const key = Math.floor(r.yBaseline / 4);
        const list = buckets.get(key);
        if (list) list.push(r); else buckets.set(key, [r]);
    }
    // Each marker-sized shape before a line of text, with whether text ends just before it on its line.
    const found: { s: PdfShape; next: RawRun; textBefore: boolean; line: number }[] = [];
    // Every marker-sized shape on a line of text, by the line.
    const perLine = new Map<number, number>();
    for (const s of shapes) {
        if (s.width <= 0 || s.height <= 0 || s.width > 24 || s.height > 24) continue;
        if (Math.abs(s.width - s.height) > 0.25 * Math.max(s.width, s.height)) continue;
        const cx = s.x + s.width / 2, cy = s.y + s.height / 2, right = s.x + s.width;
        let next: RawRun | undefined, counted = false, over = false, textBefore = false;
        for (let key = Math.floor(cy / 4) - 1; key <= Math.floor((cy + 24) / 4) + 1 && !over; key++) {
            for (const r of buckets.get(key) ?? []) {
                const fs = r.fontSize || 12;
                if (s.width < 0.15 * fs || s.width > 0.6 * fs) continue;
                if (cy < r.yBaseline - 0.8 * fs || cy > r.yBaseline - 0.05 * fs) continue;
                if (!counted) { counted = true; const k = Math.round(r.yBaseline); perLine.set(k, (perLine.get(k) ?? 0) + 1); }
                // On the shape's line, a run reaching over it means the shape is in the text, not before it.
                if (r.x < cx && r.x + r.width > cx) { over = true; break; }
                if (r.x + r.width <= s.x + 0.5 && s.x - (r.x + r.width) < 4 * fs) textBefore = true;
                if (r.x < right - 0.5 || r.x - right > 3 * fs) continue;
                if (!next || r.x < next.x) next = r;
            }
        }
        if (over || !next || /^\s*[•◦▪‣·○●■□∙]/.test(next.text)) continue;
        found.push({ s, next, textBefore, line: Math.round(next.yBaseline) });
    }
    // A list's markers stand one over another, at one x on the lines of its items (in a column of a page
    // too, beside the text of the column before it). A shape with none over or under it is a marker where
    // it stands alone on its line with no text just before it: shapes in a row with their labels between
    // them (a chart legend's keys, "Series 1 ▪ Series 2") are none.
    const byX = new Map<number, typeof found>();
    for (const f of found) { const k = Math.round(f.s.x); for (const key of [k - 1, k, k + 1]) { const list = byX.get(key); if (list) list.push(f); else byX.set(key, [f]); } }
    const stacked = (f: typeof found[number]) => {
        const em = f.next.fontSize || 12;
        return (byX.get(Math.round(f.s.x)) ?? []).some(o => o !== f && Math.abs(o.s.x - f.s.x) <= 1.5 && o.line !== f.line && Math.abs(o.line - f.line) <= 4 * em);
    };
    const out: RawRun[] = [];
    for (const f of found) {
        const alone = (perLine.get(f.line - 1) ?? 0) + (perLine.get(f.line) ?? 0) + (perLine.get(f.line + 1) ?? 0) === 1;
        if (!stacked(f) && (f.textBefore || !alone)) continue;
        const { s, next } = f;
        const glyph = s.curved ? (s.filled ? '•' : '◦') : (s.filled ? '▪' : '□');
        out.push({
            text: glyph, x: s.x, width: s.width, yTop: next.yTop, yBaseline: next.yBaseline, height: next.height, fontSize: next.fontSize,
            dir: 'ltr', angle: 0, turn: 0, mcid: next.mcid, inArtifact: next.inArtifact, formatting: {},
        });
    }
    return out;
}

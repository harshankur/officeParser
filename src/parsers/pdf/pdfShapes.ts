/**
 * The lines and shapes a page draws, from its operator list, for what text alone does not show: a list
 * item's bullet drawn as a small circle or square (as Chromium draws `<ul>` markers), and a table's
 * ruling lines and cell borders. pdf.js gives each path as one `constructPath` operator: its paint
 * operator (fill, stroke, both, or none), its data in the user space of the moment (`DrawOPS`: move,
 * line, cubic and quadratic curves, close), and its box. The current transformation is tracked as
 * `collectColorMarks` tracks it (`q`/`Q`, `cm`, form XObjects), so a shape lands in the same viewport
 * space as the text runs. Bounded: at most MAX_SHAPES shapes a page, and MAX_SEGMENTS straight edges.
 *
 * @module parsers/pdf/pdfShapes
 */

import { identityMatrix, mulMatrix as mul, toMatrix6 as toMat6 } from './geometry.js';
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
}

/** Walks an operator list fetched with drawing operators kept, and returns the shapes it paints. */
export function collectShapes(ops: any, viewportTransform: number[], OPS: any): PdfShape[] {
    const shapes: PdfShape[] = [];
    const fnArray: number[] | undefined = ops?.fnArray;
    const argsArray: any[] | undefined = ops?.argsArray;
    if (!fnArray || !argsArray || OPS.constructPath === undefined) return shapes;
    let ctm = identityMatrix();
    let fill: string | undefined, stroke: string | undefined, lineWidth = 1;
    const stack: { ctm: number[]; fill?: string; stroke?: string; lineWidth: number }[] = [];
    let segmentsLeft = MAX_SEGMENTS;
    const fills = new Set([OPS.fill, OPS.eoFill, OPS.fillStroke, OPS.eoFillStroke, OPS.closeFillStroke, OPS.closeEOFillStroke]);
    const strokes = new Set([OPS.stroke, OPS.closeStroke, OPS.fillStroke, OPS.eoFillStroke, OPS.closeFillStroke, OPS.closeEOFillStroke]);
    for (let i = 0; i < fnArray.length && shapes.length < MAX_SHAPES; i++) {
        const a = argsArray[i];
        switch (fnArray[i]) {
            case OPS.save: stack.push({ ctm, fill, stroke, lineWidth }); break;
            case OPS.restore: { const s = stack.pop(); if (s) ({ ctm, fill, stroke, lineWidth } = s); break; }
            case OPS.transform: { const cm = toMat6(a); if (cm) ctm = mul(ctm, cm); break; }
            case OPS.paintFormXObjectBegin: {
                stack.push({ ctm, fill, stroke, lineWidth });
                const fm = toMat6(a?.[0]);
                if (fm) ctm = mul(ctm, fm);
                break;
            }
            case OPS.paintFormXObjectEnd: { const s = stack.pop(); if (s) ({ ctm, fill, stroke, lineWidth } = s); break; }
            case OPS.setFillRGBColor: fill = typeof a?.[0] === 'string' ? a[0] : undefined; break;
            case OPS.setFillColorN: case OPS.setFillTransparent: fill = undefined; break;
            case OPS.setStrokeRGBColor: stroke = typeof a?.[0] === 'string' ? a[0] : undefined; break;
            case OPS.setStrokeColorN: case OPS.setStrokeTransparent: stroke = undefined; break;
            case OPS.setLineWidth: { const w = +a?.[0]; if (Number.isFinite(w) && w >= 0) lineWidth = w; break; }
            case OPS.constructPath: {
                const paint = a?.[0];
                const filled = fills.has(paint), stroked = strokes.has(paint);
                if (!filled && !stroked) break;
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
                    if (segments.length < MAX_SHAPE_SEGMENTS && segmentsLeft > 0) { segments.push([cx, cy, x, y]); segmentsLeft--; }
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
                // A stroke's width in viewport space, scaled as the CTM scales (its mean).
                const scale = Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1;
                shapes.push({
                    x: minX, y: minY, width: maxX - minX, height: maxY - minY, filled, stroked,
                    ...(filled && fill ? { fill } : {}), ...(stroked && stroke ? { stroke } : {}),
                    lineWidth: lineWidth * scale, curved, segments,
                });
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
    const out: RawRun[] = [];
    for (const s of shapes) {
        if (s.width <= 0 || s.height <= 0 || s.width > 24 || s.height > 24) continue;
        if (Math.abs(s.width - s.height) > 0.25 * Math.max(s.width, s.height)) continue;
        const cx = s.x + s.width / 2, cy = s.y + s.height / 2, right = s.x + s.width;
        let next: RawRun | undefined;
        for (let key = Math.floor(cy / 4) - 1; key <= Math.floor((cy + 24) / 4) + 1; key++) {
            for (const r of buckets.get(key) ?? []) {
                const fs = r.fontSize || 12;
                if (s.width < 0.15 * fs || s.width > 0.6 * fs) continue;
                if (cy < r.yBaseline - 0.8 * fs || cy > r.yBaseline - 0.05 * fs) continue;
                // On the shape's line, a run reaching over it means the shape is in the text, not before it.
                if (r.x < cx && r.x + r.width > cx) { next = undefined; key = Infinity; break; }
                if (r.x < right - 0.5 || r.x - right > 3 * fs) continue;
                if (!next || r.x < next.x) next = r;
            }
        }
        if (!next || /^\s*[•◦▪‣·○●■□∙]/.test(next.text)) continue;
        const glyph = s.curved ? (s.filled ? '•' : '◦') : (s.filled ? '▪' : '□');
        out.push({
            text: glyph, x: s.x, width: s.width, yTop: next.yTop, yBaseline: next.yBaseline, height: next.height, fontSize: next.fontSize,
            dir: 'ltr', angle: 0, turn: 0, mcid: next.mcid, inArtifact: next.inArtifact, formatting: {},
        });
    }
    return out;
}

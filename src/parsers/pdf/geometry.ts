/**
 * Pure geometry helpers for the PDF pipeline: affine-matrix and bounding-box arithmetic plus run
 * normalization. No pdf.js objects appear here; callers pass in plain arrays of numbers.
 *
 * @module parsers/pdf/geometry
 */

import { NodeBounds } from '../../types.js';
import { RunAngle } from './pdfTypes.js';

/** Upper bound on a run's font size in points. Guards against a hostile text matrix (see computeRunBox). */
const MAX_FONT_SIZE_PT = 10000;

/** The identity affine matrix, as a fresh array the caller may mutate. */
export function identityMatrix(): number[] {
    return [1, 0, 0, 1, 0, 0];
}

/** Composes two affine matrices, `pdfjs.Util.transform(a, b)`, inlined to keep pdf.js types out. */
export function mulMatrix(a: number[], b: number[]): number[] {
    return [
        a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1],
        a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3],
        a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5],
    ];
}

/**
 * Coerces a pdf.js matrix operand to a plain `number[6]`, or null if it is not a 6-element
 * array-like. pdf.js passes some matrices (`setTextMatrix`, `paintFormXObjectBegin`) as a single
 * argument that is a `Float32Array`, so `Array.isArray` is false and its six values sit at `[0]`.
 */
export function toMatrix6(m: any): number[] | null {
    if (!m || typeof m.length !== 'number' || m.length < 6) return null;
    return [+m[0], +m[1], +m[2], +m[3], +m[4], +m[5]];
}

/** Rounds to 2 decimals, collapsing -0 to 0. */
export function round2(n: number): number {
    const r = Math.round(n * 100) / 100;
    return r === 0 ? 0 : r;
}

/** Returns a copy of the box with all fields rounded to 2 decimals (its baseline too, where it has one). */
export function roundBounds(b: NodeBounds): NodeBounds {
    const out: NodeBounds = { x: round2(b.x), y: round2(b.y), width: round2(b.width), height: round2(b.height) };
    if (b.baseline !== undefined && Number.isFinite(b.baseline)) out.baseline = round2(b.baseline);
    return out;
}

/** Smallest box containing both inputs. */
export function unionBounds(a: NodeBounds, b: NodeBounds): NodeBounds {
    const x = Math.min(a.x, b.x);
    const y = Math.min(a.y, b.y);
    const right = Math.max(a.x + a.width, b.x + b.width);
    const bottom = Math.max(a.y + a.height, b.y + b.height);
    return { x, y, width: right - x, height: bottom - y };
}

/** Smallest box containing every input, or undefined for an empty list (a box of boxes: no baseline, see NodeBounds). */
export function unionAll(boxes: (NodeBounds | undefined)[]): NodeBounds | undefined {
    let acc: NodeBounds | undefined;
    for (const b of boxes) {
        if (!b) continue;
        acc = acc ? unionBounds(acc, b) : { x: b.x, y: b.y, width: b.width, height: b.height };
    }
    return acc;
}

/**
 * Snaps an angle in radians to the nearest quarter turn when within ~0.5 degrees, else -1; text within
 * 15 degrees of level (a scan's skewed text layer, a stamp's slight tilt) is level text, read in its
 * lines: read apart as turned text, it was taken out of the paragraphs it is part of.
 */
export function snapAngle(radians: number): RunAngle | -1 {
    let deg = (radians * 180) / Math.PI;
    deg = ((deg % 360) + 360) % 360;
    for (const cand of [0, 90, 180, 270] as const) {
        if (Math.abs(deg - cand) <= 0.5 || Math.abs(deg - (cand + 360)) <= 0.5) return cand;
    }
    if (deg <= 15 || deg >= 345) return 0;
    return -1;
}

/** Geometry derived from a single normalized text run. */
export interface RunBox {
    x: number;
    yTop: number;
    yBaseline: number;
    width: number;
    height: number;
    fontSize: number;
    angle: RunAngle | -1;
}

/**
 * Normalizes one text item into top-left-origin viewport space.
 *
 * @param m - the composed matrix `Util.transform(viewport.transform, item.transform)`
 * @param itemWidth - pdf.js `TextItem.width` (already in viewport units for a scale-1 viewport)
 * @param ascent - font ascent as a fraction of font size (from `textContent.styles`)
 * @param descent - font descent as a fraction of font size (negative)
 */
export function computeRunBox(m: number[], itemWidth: number, ascent: number, descent: number): RunBox {
    // Numbers a box can be made of: a font's metrics and an advance that are none (NaN, infinities) are a
    // font's usual ones and no advance.
    if (!Number.isFinite(ascent) || ascent <= 0) ascent = 0.8;
    if (!Number.isFinite(descent) || descent > 0) descent = -0.2;
    if (!Number.isFinite(itemWidth) || itemWidth < 0) itemWidth = 0;
    // Clamp to a sane maximum: the size comes from the document's own (untrusted) text matrix, and a
    // pathological `Tf`/scale can make `hypot(m2,m3)` astronomically large, which would then drive a
    // huge box height and, via the color lookup's y-window, an unbounded loop. No real font is anywhere
    // near this cap, so clamping never affects a genuine document.
    const size = Math.hypot(m[2], m[3]) || Math.abs(m[3]);
    const fontSize = Math.min(Number.isFinite(size) && size > 0 ? size : 12, MAX_FONT_SIZE_PT);
    const angle = snapAngle(Math.atan2(m[1], m[0]));
    const originX = m[4];
    const yBaseline = m[5];
    const height = (ascent - descent) * fontSize;

    if (angle === 0) {
        return { x: originX, yTop: yBaseline - ascent * fontSize, yBaseline, width: itemWidth, height, fontSize, angle };
    }
    // Non-horizontal runs: report an axis-aligned box around the advance so bounds stay meaningful.
    // The horizontal line builder skips these; PdfParser's rotatedTextNodes recovers 90/180/270 runs
    // as trailing paragraphs so their text is not lost.
    if (angle === 180) {
        // A 180-degree run is upside down: its glyph body extends BELOW the baseline in viewport space,
        // so the box spans [yBaseline + descent*fs, yBaseline + ascent*fs] (descent is negative), not the
        // upright [yBaseline - ascent*fs, ...]. Height is unchanged.
        return { x: originX - itemWidth, yTop: yBaseline + descent * fontSize, yBaseline, width: itemWidth, height, fontSize, angle };
    }
    if (angle === 90 || angle === 270) {
        // Advance is vertical; approximate a box tall by itemWidth and wide by height.
        const top = angle === 90 ? yBaseline : yBaseline - itemWidth;
        return { x: yBaseline === top ? originX : originX - height, yTop: top, yBaseline, width: height, height: itemWidth, fontSize, angle };
    }
    // Arbitrary angle (-1): fall back to a small box at the origin.
    return { x: originX, yTop: yBaseline - ascent * fontSize, yBaseline, width: itemWidth, height, fontSize, angle };
}

/**
 * Maps a box from authored (rotation-0) page space into the page's rendered space for a given
 * `/Rotate`. Reading order is computed in authored space; this is applied only at emit time so the
 * public `NodeBounds` match what pdf.js renders. For rotation 0 it is the identity; on a turned page the
 * box keeps no baseline (its text is turned on the page as it is seen, see NodeBounds.baseline).
 *
 * @param authoredW - authored page width (rotation-0 viewport width)
 * @param authoredH - authored page height
 */
export function rotateBoundsToRendered(b: NodeBounds, rotation: number, authoredW: number, authoredH: number): NodeBounds {
    const r = ((rotation % 360) + 360) % 360;
    if (r === 0) return b;
    const corners = [
        { x: b.x, y: b.y },
        { x: b.x + b.width, y: b.y + b.height },
    ].map(p => {
        if (r === 90) return { x: authoredH - p.y, y: p.x };
        if (r === 180) return { x: authoredW - p.x, y: authoredH - p.y };
        return { x: p.y, y: authoredW - p.x }; // 270
    });
    const x = Math.min(corners[0].x, corners[1].x);
    const y = Math.min(corners[0].y, corners[1].y);
    return { x, y, width: Math.abs(corners[0].x - corners[1].x), height: Math.abs(corners[0].y - corners[1].y) };
}

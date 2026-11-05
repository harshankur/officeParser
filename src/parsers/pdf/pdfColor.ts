/**
 * Per-run text fill color from a page's operator list (pdf.js exposes no color on `getTextContent`).
 *
 * pdf.js's worker normalizes every fill-color operator into a single `OPS.setFillRGBColor` whose one
 * argument is an already-computed lowercase `#rrggbb` string (gray, CMYK, indexed, ICC and separation
 * spaces are all pre-converted for us), and every text-show operator into `OPS.showText`. So we never
 * touch color spaces: the page's shows, each with the colour it paints in and its start point in user
 * space (see textGlyphs, which mirrors the graphics and text state as `getTextContent` does, so both
 * share one frame), are taken into viewport space here, where a run looks up its color by position.
 * Pattern/shading/transparent fills yield no color, so a run over them simply keeps its default (never
 * a wrong color).
 *
 * @module parsers/pdf/pdfColor
 */

import { TextShow } from './textGlyphs.js';

/** One text-show occurrence: its start point in viewport space plus the active fill color. */
export interface ColorMark {
    x: number;
    y: number;
    color: string;
}

/**
 * The viewport-space start point and fill color of every show of a page, in paint order. A show is
 * left out when its fill is a pattern, shading or transparent, or when the text is invisible (render
 * modes 3/7); stroke-render text (modes 1/5) has the stroke color.
 *
 * @param viewportTransform - `viewport.transform` for the page (scale-1 layout viewport).
 */
export function collectColorMarks(shows: readonly TextShow[], viewportTransform: number[]): ColorMark[] {
    const marks: ColorMark[] = [];
    for (const show of shows) {
        if (show.mode === 3 || show.mode === 7 || !show.color) continue;
        const vx = viewportTransform[0] * show.x + viewportTransform[2] * show.y + viewportTransform[4];
        const vy = viewportTransform[1] * show.x + viewportTransform[3] * show.y + viewportTransform[5];
        marks.push({ x: vx, y: vy, color: show.color });
    }
    return marks;
}

/**
 * A run's color lookup. Given the run's start point `(vx, vy)`, its font size and its width, returns
 * the fill color, or undefined. See {@link makeColorLookup}.
 */
export type ColorLookup = (vx: number, vy: number, fontSize: number, width: number) => string | undefined;

/**
 * Builds a color lookup over a page's marks, bucketed by rounded viewport-y. A run is coloured only
 * when the marks lying on its baseline *within its own horizontal extent* `[vx, vx+width]` agree on a
 * single colour: that colour is returned. If no mark falls inside the run it keeps its default (so a
 * default-black run never borrows a neighbour's colour), and if the marks disagree - a colour change
 * mid-run that pdf.js merged into one text item - it returns undefined rather than confidently
 * picking the wrong one.
 */
/**
 * A page with more colour marks than this is not coloured at all (the lookup returns undefined for
 * every run). A real page has hundreds to low thousands of coloured pieces; this ceiling is far above
 * any genuine document and exists so a crafted content stream that emits millions of marks cannot make
 * the per-page build (a sort of every bucket) or the queries do unbounded work.
 */
const MAX_COLOR_MARKS = 300_000;

export function makeColorLookup(marks: ColorMark[]): ColorLookup {
    if (marks.length > MAX_COLOR_MARKS) return () => undefined;

    // Bucket by rounded viewport-y, then sort each bucket by x once and precompute, per bucket, a
    // prefix count of colour changes along x. A run's query then answers "do the marks in its
    // horizontal window agree, and on which colour?" with two binary searches and one prefix
    // subtraction - O(log n) per bucket - instead of scanning every mark. Without this, a page that
    // crowds N marks and N runs on one baseline is O(N^2): a CPU-hang DoS, because colour is default-on.
    interface Bucket { xs: number[]; colors: string[]; changePrefix: Int32Array; }
    const rows = new Map<number, ColorMark[]>();
    for (const mk of marks) {
        const k = Math.round(mk.y);
        const bucket = rows.get(k);
        if (bucket) bucket.push(mk); else rows.set(k, [mk]);
    }
    const buckets = new Map<number, Bucket>();
    for (const [k, arr] of rows) {
        arr.sort((a, b) => a.x - b.x);
        const xs = new Array<number>(arr.length);
        const colors = new Array<string>(arr.length);
        // changePrefix[i] = number of colour changes in colors[0..i]. colors[a..b] are all one colour
        // iff changePrefix[b] === changePrefix[a], so a range's homogeneity is an O(1) comparison.
        const changePrefix = new Int32Array(arr.length);
        for (let i = 0; i < arr.length; i++) {
            xs[i] = arr[i].x;
            colors[i] = arr[i].color;
            changePrefix[i] = i === 0 ? 0 : changePrefix[i - 1] + (colors[i] !== colors[i - 1] ? 1 : 0);
        }
        buckets.set(k, { xs, colors, changePrefix });
    }
    // First index with xs[i] >= target.
    const lowerBound = (xs: number[], target: number): number => {
        let lo = 0, hi = xs.length;
        while (lo < hi) { const mid = (lo + hi) >> 1; if (xs[mid] < target) lo = mid + 1; else hi = mid; }
        return lo;
    };
    // Last index with xs[i] <= target, or -1 when none.
    const upperIndex = (xs: number[], target: number): number => {
        let lo = 0, hi = xs.length;
        while (lo < hi) { const mid = (lo + hi) >> 1; if (xs[mid] <= target) lo = mid + 1; else hi = mid; }
        return lo - 1;
    };
    return (vx: number, vy: number, fontSize: number, width: number): string | undefined => {
        // Cap the vertical window independent of font size: a colour mark on the run's baseline sits
        // within about half the font size, but the loop steps one integer y per row, so an unclamped
        // `fontSize` (from a hostile text matrix) would iterate billions of empty rows. A colour on a
        // real line is always within 64pt; this bounds the loop to ~129 buckets. The horizontal pad is
        // clamped the same way so the window never spans thousands of points on a giant font.
        const yTol = Math.min(Math.max(2, fontSize * 0.5), 64);
        const pad = Math.min(Math.max(1, fontSize * 0.5), 64);
        const lo = vx - pad, hi = vx + width + pad;
        let found: string | undefined;
        for (let k = Math.round(vy - yTol); k <= Math.round(vy + yTol); k++) {
            const b = buckets.get(k);
            if (!b) continue;
            const a = lowerBound(b.xs, lo);
            const z = upperIndex(b.xs, hi);
            if (a > z) continue; // no mark in this bucket within the run's horizontal window
            // A colour change inside the run's window means pdf.js merged two colours into one item:
            // do not guess.
            if (b.changePrefix[z] !== b.changePrefix[a]) return undefined;
            const c = b.colors[a];
            if (found === undefined) found = c;
            else if (found !== c) return undefined; // two baselines within the window disagree
        }
        return found;
    };
}

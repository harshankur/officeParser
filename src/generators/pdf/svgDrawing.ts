/**
 * An SVG picture drawn as PDF vector graphics, for the native PDF engine, which embeds pictures as PNG
 * and JPEG and dropped an SVG (IMAGE_PROCESSING_FAILED). The picture is drawn onto a page of its own
 * (the engine embeds that page as a form, and draws it wherever the picture stands): its paths (every
 * command, arcs as curves), basic shapes, groups and their transforms, `use` of a defined element or
 * symbol, nested `svg`, text (with its anchor, size, family, weight and style), PNG and JPEG pictures
 * given as data, fills, strokes (width, caps, joins, dashes), the even-odd rule and opacity, styled by
 * presentation attributes, `style` and the picture's own `<style>` rules of simple selectors (a type,
 * a class, an id). A gradient is drawn in the mean of its stops' colours; filters, masks, clip paths and
 * patterns are left out, and named in what the drawing returns, for the engine's warning. Everything is
 * bounded: the elements drawn, a path's commands, the depth of groups and of `use`.
 */
import { parseCssColor } from '../../utils/colorUtils.js';
import { parseXmlString } from '../../utils/xmlUtils.js';
import type { OfficeParserConfig } from '../../types.js';

/** At most this many elements are drawn, `use` copies counted. */
const MAX_ELEMENTS = 50_000;
/** At most this many path commands in all. */
const MAX_PATH_COMMANDS = 500_000;
/** At most this deep groups (and nested `svg`) are followed. */
const MAX_DEPTH = 64;
/** At most this deep `use` refers to `use`. */
const MAX_USE_DEPTH = 8;
/** At most this many stylesheet rules are read. */
const MAX_RULES = 2_000;
/** At most this long the stylesheet text read. */
const MAX_CSS_LENGTH = 256 * 1024;
/** A picture larger than this, each way, in CSS pixels, is drawn at this size. */
const MAX_SIDE_PX = 14_400;

/** The faces the engine draws SVG text in. */
export type SvgFace = 'regular' | 'bold' | 'italic' | 'boldItalic' | 'serif' | 'serifBold' | 'serifItalic' | 'serifBoldItalic' | 'mono' | 'monoBold' | 'monoItalic' | 'monoBoldItalic';

/** What the engine gives the drawing: its pdf-lib, the page to draw on, and how it draws text and pictures. */
export interface SvgCanvas {
    lib: any;
    page: any;
    /** Draws text with its left end at (x, y), upright, in the current transformation. */
    text(text: string, face: SvgFace, x: number, y: number, size: number, color: any): void;
    /** The width of text in a face at a size. */
    width(text: string, face: SvgFace, size: number): number;
    /** A PNG or JPEG picture embedded for drawing, or undefined where it cannot be. */
    image(bytes: Uint8Array, kind: 'png' | 'jpeg'): Promise<any | undefined>;
}

/** An SVG read: its size in CSS pixels, and its tree, to draw. */
export interface SvgPicture {
    width: number;
    height: number;
    root: Element;
    viewBox?: [number, number, number, number];
    preserve: string;
    rules: Rule[];
    ids: Map<string, Element>;
}

type Matrix = [number, number, number, number, number, number];
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

interface Rule { tag?: string; classes: string[]; id?: string; specificity: number; order: number; declarations: Map<string, string> }

/** The style an element draws with, inherited from its parents where SVG inherits it. */
interface Style {
    fill: string; stroke: string; strokeWidth: number; fillOpacity: number; strokeOpacity: number; opacity: number;
    fillRule: string; lineCap: string; lineJoin: string; dash: string; fontSize: number; fontFamily: string;
    fontWeight: string; fontStyle: string; textAnchor: string; color: string; visibility: string;
}

const INITIAL_STYLE: Style = {
    fill: 'black', stroke: 'none', strokeWidth: 1, fillOpacity: 1, strokeOpacity: 1, opacity: 1, fillRule: 'nonzero',
    lineCap: 'butt', lineJoin: 'miter', dash: 'none', fontSize: 16, fontFamily: 'sans-serif', fontWeight: 'normal',
    fontStyle: 'normal', textAnchor: 'start', color: 'black', visibility: 'visible',
};

/** The properties read from attributes, `style` and rules. */
const PROPERTIES = new Set(['fill', 'stroke', 'stroke-width', 'fill-opacity', 'stroke-opacity', 'opacity', 'fill-rule',
    'stroke-linecap', 'stroke-linejoin', 'stroke-dasharray', 'font-size', 'font-family', 'font-weight', 'font-style',
    'text-anchor', 'color', 'visibility', 'display', 'stop-color', 'stop-opacity', 'font', 'filter', 'mask', 'clip-path']);

const localName = (el: Element): string => (el.localName || el.tagName || '').replace(/^.*:/, '');
const childElements = (el: Element): Element[] => {
    const out: Element[] = [];
    for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as Element);
    return out;
};

/** A number, or NaN. */
const num = (v: string | null | undefined): number => {
    if (v == null) return NaN;
    const m = /^\s*([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)/.exec(v);
    return m ? parseFloat(m[1]) : NaN;
};

/** A length in user units (CSS pixels): px, pt, pc, mm, cm, in, em, ex, or a percentage of `whole`. */
function length(v: string | null | undefined, whole: number, fontSize = 16): number {
    if (v == null) return NaN;
    const m = /^\s*([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)\s*(px|pt|pc|mm|cm|in|em|ex|%)?\s*$/i.exec(v);
    if (!m) return NaN;
    const n = parseFloat(m[1]);
    switch ((m[2] || '').toLowerCase()) {
        case 'pt': return n * 4 / 3;
        case 'pc': return n * 16;
        case 'mm': return n * 96 / 25.4;
        case 'cm': return n * 96 / 2.54;
        case 'in': return n * 96;
        case 'em': return n * fontSize;
        case 'ex': return n * fontSize / 2;
        case '%': return n * whole / 100;
        default: return n;
    }
}

/** Numbers in a list (`10 20,30`), at most `max`. */
function numbers(v: string | null | undefined, max = 1 << 20): number[] {
    const out: number[] = [];
    if (!v) return out;
    const re = /[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/g;
    for (let m = re.exec(v); m && out.length < max; m = re.exec(v)) out.push(parseFloat(m[0]));
    return out;
}

const multiply = (a: Matrix, b: Matrix): Matrix => [
    a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1],
    a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3],
    a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5],
];

/** A `transform` attribute as one matrix. */
function transformOf(v: string | null): Matrix {
    let m: Matrix = IDENTITY;
    if (!v) return m;
    const re = /(matrix|translate|scale|rotate|skewX|skewY)\s*\(([^)]*)\)/g;
    let count = 0;
    for (let t = re.exec(v); t && count < 64; t = re.exec(v), count++) {
        const a = numbers(t[2], 6);
        let next: Matrix = IDENTITY;
        switch (t[1]) {
            case 'matrix': if (a.length === 6) next = a as Matrix; break;
            case 'translate': next = [1, 0, 0, 1, a[0] || 0, a[1] || 0]; break;
            case 'scale': next = [a[0] ?? 1, 0, 0, a[1] ?? a[0] ?? 1, 0, 0]; break;
            case 'rotate': {
                const r = (a[0] || 0) * Math.PI / 180, c = Math.cos(r), s = Math.sin(r), cx = a[1] || 0, cy = a[2] || 0;
                next = [c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy];
                break;
            }
            case 'skewX': next = [1, 0, Math.tan((a[0] || 0) * Math.PI / 180), 1, 0, 0]; break;
            case 'skewY': next = [1, Math.tan((a[0] || 0) * Math.PI / 180), 0, 1, 0, 0]; break;
        }
        m = multiply(m, next);
    }
    return m.every(Number.isFinite) ? m : IDENTITY;
}

/** Declarations (`fill: red; stroke: blue`), the properties drawing reads. */
function declarations(text: string, into = new Map<string, string>()): Map<string, string> {
    for (const part of text.split(';')) {
        const colon = part.indexOf(':');
        if (colon < 0) continue;
        const name = part.slice(0, colon).trim().toLowerCase();
        if (!PROPERTIES.has(name)) continue;
        const value = part.slice(colon + 1).replace(/!\s*important\s*$/i, '').trim();
        if (value && value.length <= 256) into.set(name, value);
    }
    return into;
}

/** The picture's `<style>` rules of simple selectors: a type, classes, an id, or any element. */
function readRules(css: string): Rule[] {
    const rules: Rule[] = [];
    const text = css.slice(0, MAX_CSS_LENGTH).replace(/\/\*[\s\S]*?(?:\*\/|$)/g, ' ');
    let at = 0;
    while (at < text.length && rules.length < MAX_RULES) {
        const open = text.indexOf('{', at);
        if (open < 0) break;
        const close = text.indexOf('}', open);
        const prelude = text.slice(at, open).trim(), body = text.slice(open + 1, close < 0 ? text.length : close);
        at = close < 0 ? text.length : close + 1;
        if (!prelude || prelude.startsWith('@')) continue;
        const decls = declarations(body);
        if (!decls.size) continue;
        for (const selector of prelude.split(',')) {
            const m = /^\s*([a-zA-Z][\w-]*|\*)?((?:[.#][\w-]+)*)\s*$/.exec(selector);
            if (!m || (!m[1] && !m[2])) continue;
            const classes: string[] = [];
            let id: string | undefined;
            for (const part of m[2].match(/[.#][\w-]+/g) ?? []) part[0] === '.' ? classes.push(part.slice(1)) : (id = part.slice(1));
            const tag = m[1] && m[1] !== '*' ? m[1] : undefined;
            rules.push({ tag, classes, id, specificity: (id ? 10_000 : 0) + classes.length * 100 + (tag ? 1 : 0), order: rules.length, declarations: decls });
            if (rules.length >= MAX_RULES) break;
        }
    }
    return rules.sort((a, b) => a.specificity - b.specificity || a.order - b.order);
}

/**
 * Reads an SVG's text: undefined for one that is no SVG, or that is too large to draw. Its size is its
 * `width` and `height` (in CSS pixels), else its `viewBox`'s, else 300 by 150, as a browser sizes it.
 */
export function readSvg(text: string, config?: OfficeParserConfig): SvgPicture | undefined {
    let doc: Document;
    try { doc = parseXmlString(text, { config }); } catch { return undefined; }
    const root = doc?.documentElement;
    if (!root || localName(root) !== 'svg') return undefined;
    const vb = numbers(root.getAttribute('viewBox'), 4);
    const viewBox = vb.length === 4 && vb[2] > 0 && vb[3] > 0 ? vb as [number, number, number, number] : undefined;
    let width = length(root.getAttribute('width'), viewBox?.[2] ?? 300), height = length(root.getAttribute('height'), viewBox?.[3] ?? 150);
    if (!(width > 0)) width = viewBox && height > 0 ? height * viewBox[2] / viewBox[3] : viewBox?.[2] ?? 300;
    if (!(height > 0)) height = viewBox ? width * viewBox[3] / viewBox[2] : 150;
    const scale = Math.min(1, MAX_SIDE_PX / width, MAX_SIDE_PX / height);
    width *= scale; height *= scale;
    const ids = new Map<string, Element>();
    let css = '';
    // Every element's id, and the stylesheets, in one walk (bounded like the drawing).
    const stack: Element[] = [root];
    let seen = 0;
    while (stack.length && seen < MAX_ELEMENTS * 2) {
        const el = stack.pop()!;
        seen++;
        const id = el.getAttribute('id');
        if (id && !ids.has(id)) ids.set(id, el);
        if (localName(el) === 'style') { if (css.length < MAX_CSS_LENGTH) css += (el.textContent || '') + '\n'; continue; }
        const kids = childElements(el);
        for (let i = kids.length - 1; i >= 0; i--) stack.push(kids[i]);
    }
    return { width, height, root, viewBox, preserve: root.getAttribute('preserveAspectRatio') || 'xMidYMid meet', rules: readRules(css), ids };
}

/** The matrix that places a viewBox in a viewport of `w` by `h`, as `preserveAspectRatio` says. */
function viewBoxMatrix(viewBox: [number, number, number, number] | undefined, preserve: string, w: number, h: number): Matrix {
    if (!viewBox) return IDENTITY;
    const [x, y, vw, vh] = viewBox;
    let sx = w / vw, sy = h / vh;
    const [align = 'xMidYMid', meetOrSlice = 'meet'] = preserve.trim().split(/\s+/);
    if (align !== 'none') {
        const s = meetOrSlice === 'slice' ? Math.max(sx, sy) : Math.min(sx, sy);
        sx = sy = s;
    }
    const free = (whole: number, used: number, where: string) => where === 'Mid' ? (whole - used) / 2 : where === 'Max' ? whole - used : 0;
    const ax = align === 'none' ? 0 : free(w, vw * sx, /^xM(id|ax)/.exec(align)?.[1] ?? 'Min');
    const ay = align === 'none' ? 0 : free(h, vh * sy, /YM(id|ax)$/.exec(align)?.[1] ?? 'Min');
    return [sx, 0, 0, sy, ax - x * sx, ay - y * sy];
}

/** One path segment, absolute: move, line, cubic curve, close. */
type Segment = ['M', number, number] | ['L', number, number] | ['C', number, number, number, number, number, number] | ['Z'];

/** An arc (SVG's endpoint form) as cubic curves. */
function arcToCurves(x1: number, y1: number, rx: number, ry: number, angle: number, large: boolean, sweep: boolean, x2: number, y2: number): Segment[] {
    if (rx === 0 || ry === 0) return [['L', x2, y2]];
    if (x1 === x2 && y1 === y2) return [];
    rx = Math.abs(rx); ry = Math.abs(ry);
    const phi = angle * Math.PI / 180, cos = Math.cos(phi), sin = Math.sin(phi);
    const dx = (x1 - x2) / 2, dy = (y1 - y2) / 2;
    const x1p = cos * dx + sin * dy, y1p = -sin * dx + cos * dy;
    const lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
    if (lambda > 1) { rx *= Math.sqrt(lambda); ry *= Math.sqrt(lambda); }
    const num2 = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
    const den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
    let coef = den ? Math.sqrt(Math.max(0, num2 / den)) : 0;
    if (large === sweep) coef = -coef;
    const cxp = coef * rx * y1p / ry, cyp = -coef * ry * x1p / rx;
    const cx = cos * cxp - sin * cyp + (x1 + x2) / 2, cy = sin * cxp + cos * cyp + (y1 + y2) / 2;
    const angleOf = (ux: number, uy: number, vx: number, vy: number) => {
        const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
        return a;
    };
    const theta1 = angleOf(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
    let delta = angleOf((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
    if (!sweep && delta > 0) delta -= 2 * Math.PI;
    if (sweep && delta < 0) delta += 2 * Math.PI;
    const parts = Math.min(8, Math.max(1, Math.ceil(Math.abs(delta) / (Math.PI / 2))));
    const step = delta / parts, k = 4 / 3 * Math.tan(step / 4);
    const out: Segment[] = [];
    const point = (t: number) => [Math.cos(t), Math.sin(t)];
    const place = (ex: number, ey: number): [number, number] => [cx + rx * ex * cos - ry * ey * sin, cy + rx * ex * sin + ry * ey * cos];
    let t = theta1;
    for (let i = 0; i < parts; i++) {
        const [c1, s1] = point(t), [c2, s2] = point(t + step);
        const p1 = place(c1 - k * s1, s1 + k * c1), p2 = place(c2 + k * s2, s2 - k * c2), p3 = place(c2, s2);
        out.push(['C', p1[0], p1[1], p2[0], p2[1], p3[0], p3[1]]);
        t += step;
    }
    return out;
}

/** Path data (`d`) as absolute segments, read once forward; at most `budget.left` commands. */
function parsePath(d: string, budget: { left: number }): Segment[] {
    const out: Segment[] = [];
    const n = d.length;
    let i = 0;
    const NUMBER = /[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/y;
    const skip = () => { while (i < n && (d[i] === ' ' || d[i] === ',' || d[i] === '\n' || d[i] === '\t' || d[i] === '\r')) i++; };
    const readNumber = (): number => { skip(); NUMBER.lastIndex = i; const m = NUMBER.exec(d); if (!m) return NaN; i = NUMBER.lastIndex; return parseFloat(m[0]); };
    const readFlag = (): boolean | undefined => { skip(); const c = d[i]; if (c === '0' || c === '1') { i++; return c === '1'; } return undefined; };
    let x = 0, y = 0, sx = 0, sy = 0, cmd = '', lastCtrl: [number, number] | null = null, lastQuad: [number, number] | null = null;
    while (i < n && budget.left > 0) {
        skip();
        if (i >= n) break;
        const c = d[i];
        if (/[MmLlHhVvCcSsQqTtAaZz]/.test(c)) { cmd = c; i++; }
        else if (!cmd || cmd === 'Z' || cmd === 'z') break;
        else if (!/[0-9+\-.]/.test(c)) break;
        budget.left--;
        const rel = cmd === cmd.toLowerCase();
        const ox = rel ? x : 0, oy = rel ? y : 0;
        const before = i;
        switch (cmd.toUpperCase()) {
            case 'M': {
                const a = readNumber(), b = readNumber();
                if (isNaN(a) || isNaN(b)) return out;
                x = ox + a; y = oy + b; sx = x; sy = y;
                out.push(['M', x, y]);
                // Pairs after a move are lines.
                cmd = rel ? 'l' : 'L';
                lastCtrl = lastQuad = null;
                break;
            }
            case 'L': {
                const a = readNumber(), b = readNumber();
                if (isNaN(a) || isNaN(b)) return out;
                x = ox + a; y = oy + b; out.push(['L', x, y]); lastCtrl = lastQuad = null;
                break;
            }
            case 'H': { const a = readNumber(); if (isNaN(a)) return out; x = (rel ? x : 0) + a; out.push(['L', x, y]); lastCtrl = lastQuad = null; break; }
            case 'V': { const a = readNumber(); if (isNaN(a)) return out; y = (rel ? y : 0) + a; out.push(['L', x, y]); lastCtrl = lastQuad = null; break; }
            case 'C': {
                const v = [readNumber(), readNumber(), readNumber(), readNumber(), readNumber(), readNumber()];
                if (v.some(isNaN)) return out;
                const [x1, y1, x2, y2] = [ox + v[0], oy + v[1], ox + v[2], oy + v[3]];
                x = ox + v[4]; y = oy + v[5];
                out.push(['C', x1, y1, x2, y2, x, y]); lastCtrl = [x2, y2]; lastQuad = null;
                break;
            }
            case 'S': {
                const v = [readNumber(), readNumber(), readNumber(), readNumber()];
                if (v.some(isNaN)) return out;
                const [x1, y1] = lastCtrl ? [2 * x - lastCtrl[0], 2 * y - lastCtrl[1]] : [x, y];
                const [x2, y2] = [ox + v[0], oy + v[1]];
                x = ox + v[2]; y = oy + v[3];
                out.push(['C', x1, y1, x2, y2, x, y]); lastCtrl = [x2, y2]; lastQuad = null;
                break;
            }
            case 'Q': case 'T': {
                let qx: number, qy: number, ex: number, ey: number;
                if (cmd.toUpperCase() === 'Q') {
                    const v = [readNumber(), readNumber(), readNumber(), readNumber()];
                    if (v.some(isNaN)) return out;
                    [qx, qy, ex, ey] = [ox + v[0], oy + v[1], ox + v[2], oy + v[3]];
                } else {
                    const v = [readNumber(), readNumber()];
                    if (v.some(isNaN)) return out;
                    [qx, qy] = lastQuad ? [2 * x - lastQuad[0], 2 * y - lastQuad[1]] : [x, y];
                    [ex, ey] = [ox + v[0], oy + v[1]];
                }
                out.push(['C', x + 2 / 3 * (qx - x), y + 2 / 3 * (qy - y), ex + 2 / 3 * (qx - ex), ey + 2 / 3 * (qy - ey), ex, ey]);
                x = ex; y = ey; lastQuad = [qx, qy]; lastCtrl = null;
                break;
            }
            case 'A': {
                const rx = readNumber(), ry = readNumber(), angle = readNumber(), large = readFlag(), sweep = readFlag(), a = readNumber(), b = readNumber();
                if ([rx, ry, angle, a, b].some(isNaN) || large === undefined || sweep === undefined) return out;
                const ex = ox + a, ey = oy + b;
                for (const s of arcToCurves(x, y, rx, ry, angle, large, sweep, ex, ey)) out.push(s);
                x = ex; y = ey; lastCtrl = lastQuad = null;
                break;
            }
            case 'Z': out.push(['Z']); x = sx; y = sy; lastCtrl = lastQuad = null; break;
        }
        // Nothing read (a command letter with no numbers after it): stop rather than loop.
        if (i === before && cmd.toUpperCase() !== 'Z') break;
    }
    return out;
}

/** An ellipse as four curves. */
function ellipse(cx: number, cy: number, rx: number, ry: number): Segment[] {
    const k = 0.5522847498;
    return [['M', cx + rx, cy], ['C', cx + rx, cy + k * ry, cx + k * rx, cy + ry, cx, cy + ry], ['C', cx - k * rx, cy + ry, cx - rx, cy + k * ry, cx - rx, cy],
        ['C', cx - rx, cy - k * ry, cx - k * rx, cy - ry, cx, cy - ry], ['C', cx + k * rx, cy - ry, cx + rx, cy - k * ry, cx + rx, cy], ['Z']];
}

/** A rectangle, its corners rounded by rx and ry. */
function rect(x: number, y: number, w: number, h: number, rx: number, ry: number): Segment[] {
    if (!(rx > 0) && !(ry > 0)) return [['M', x, y], ['L', x + w, y], ['L', x + w, y + h], ['L', x, y + h], ['Z']];
    rx = Math.min(rx > 0 ? rx : ry, w / 2); ry = Math.min(ry > 0 ? ry : rx, h / 2);
    const k = 0.5522847498;
    return [['M', x + rx, y], ['L', x + w - rx, y], ['C', x + w - rx + k * rx, y, x + w, y + ry - k * ry, x + w, y + ry], ['L', x + w, y + h - ry],
        ['C', x + w, y + h - ry + k * ry, x + w - rx + k * rx, y + h, x + w - rx, y + h], ['L', x + rx, y + h],
        ['C', x + rx - k * rx, y + h, x, y + h - ry + k * ry, x, y + h - ry], ['L', x, y + ry], ['C', x, y + ry - k * ry, x + rx - k * rx, y, x + rx, y], ['Z']];
}

/** Draws an SVG read with {@link readSvg} onto `canvas.page`, which is its size in points (pixels times 0.75). Returns what it left out. */
export async function drawSvg(svg: SvgPicture, canvas: SvgCanvas): Promise<string[]> {
    const { lib, page } = canvas;
    const left = new Set<string>();
    const budget = { elements: MAX_ELEMENTS, path: { left: MAX_PATH_COMMANDS } };
    const W = svg.width, H = svg.height;
    // Page space (points, y up) from the picture's (pixels, y down), then its viewBox.
    const base: Matrix = multiply([0.75, 0, 0, -0.75, 0, H * 0.75], viewBoxMatrix(svg.viewBox, svg.preserve, W, H));
    const viewport = svg.viewBox ? { w: svg.viewBox[2], h: svg.viewBox[3] } : { w: W, h: H };
    const op = (name: string, ...args: any[]) => lib.PDFOperator.of(name, args);
    const states = new Map<string, any>();
    const opacityState = (fill: number, stroke: number): any => {
        const key = `${fill.toFixed(3)},${stroke.toFixed(3)}`;
        let name = states.get(key);
        if (!name) {
            const dict = page.doc.context.obj({ Type: 'ExtGState', ca: Math.max(0, Math.min(1, fill)), CA: Math.max(0, Math.min(1, stroke)) });
            name = page.node.newExtGState('GS', dict);
            states.set(key, name);
        }
        return name;
    };

    /** The style an element draws with: its parent's inherited, then rules, attributes, `style`. */
    const styleOf = (el: Element, parent: Style): { style: Style; own: Map<string, string> } => {
        const own = new Map<string, string>();
        const classes = (el.getAttribute('class') || '').split(/\s+/).filter(Boolean);
        const id = el.getAttribute('id') || undefined, tag = localName(el);
        // Presentation attributes first, then the stylesheet's rules over them, then `style` over both.
        for (const name of PROPERTIES) { const v = el.getAttribute(name); if (v != null && v.length <= 256) own.set(name, v.trim()); }
        for (const rule of svg.rules) {
            if ((rule.tag && rule.tag !== tag) || (rule.id && rule.id !== id) || !rule.classes.every(c => classes.includes(c))) continue;
            for (const [k, v] of rule.declarations) own.set(k, v);
        }
        const inline = el.getAttribute('style');
        if (inline) declarations(inline.slice(0, 4096), own);
        const style: Style = { ...parent, opacity: 1 };
        const get = (k: string) => { const v = own.get(k); return v === undefined || v === 'inherit' ? undefined : v; };
        const font = get('font');
        if (font) {
            const size = /(\d*\.?\d+(?:px|pt|em|%)?)\s*(?:\/\S+\s+)?(.+)$/.exec(font);
            if (size) { own.set('font-size', own.get('font-size') ?? size[1]); own.set('font-family', own.get('font-family') ?? size[2]); }
            if (/\bbold\b|\b[6-9]00\b/.test(font) && !own.has('font-weight')) own.set('font-weight', 'bold');
            if (/\bitalic\b|\boblique\b/.test(font) && !own.has('font-style')) own.set('font-style', 'italic');
        }
        const v = (k: string) => get(k);
        if (v('fill')) style.fill = v('fill')!;
        if (v('stroke')) style.stroke = v('stroke')!;
        if (v('color')) style.color = v('color')!;
        const sw = length(v('stroke-width'), Math.hypot(viewport.w, viewport.h) / Math.SQRT2, parent.fontSize);
        if (sw >= 0) style.strokeWidth = sw;
        const unit = (k: string, current: number) => { const n = num(v(k)); return Number.isFinite(n) ? Math.max(0, Math.min(1, /%\s*$/.test(v(k)!) ? n / 100 : n)) : current; };
        style.fillOpacity = unit('fill-opacity', parent.fillOpacity);
        style.strokeOpacity = unit('stroke-opacity', parent.strokeOpacity);
        style.opacity = unit('opacity', 1);
        if (v('fill-rule')) style.fillRule = v('fill-rule')!;
        if (v('stroke-linecap')) style.lineCap = v('stroke-linecap')!;
        if (v('stroke-linejoin')) style.lineJoin = v('stroke-linejoin')!;
        if (v('stroke-dasharray')) style.dash = v('stroke-dasharray')!;
        const fs = length(v('font-size'), parent.fontSize, parent.fontSize);
        if (fs > 0) style.fontSize = fs;
        if (v('font-family')) style.fontFamily = v('font-family')!;
        if (v('font-weight')) style.fontWeight = v('font-weight')!;
        if (v('font-style')) style.fontStyle = v('font-style')!;
        if (v('text-anchor')) style.textAnchor = v('text-anchor')!;
        if (v('visibility')) style.visibility = v('visibility')!;
        return { style, own };
    };

    /** A paint (`fill`, `stroke`) as a colour and its alpha, or undefined for none. */
    const paint = (value: string, style: Style): { color: any; alpha: number } | undefined => {
        const v = value.trim();
        if (!v || v === 'none' || v === 'transparent') return undefined;
        if (v === 'currentColor' || v === 'currentcolor') return paint(style.color, { ...style, color: 'black' });
        const ref = /^url\(\s*['"]?#([^'")]+)['"]?\s*\)\s*(.*)$/.exec(v);
        if (ref) {
            const target = svg.ids.get(ref[1]);
            const kind = target ? localName(target) : '';
            if (kind === 'linearGradient' || kind === 'radialGradient') {
                // A gradient in the mean of its stops' colours (its own, else those of the one it refers to).
                left.add('gradients (drawn in one colour)');
                let stops: Element[] = [];
                for (let g: Element | undefined = target, hops = 0; g && !stops.length && hops < 8; hops++) {
                    stops = childElements(g).filter(s => localName(s) === 'stop');
                    const href = g.getAttribute('href') || g.getAttribute('xlink:href');
                    g = href?.startsWith('#') ? svg.ids.get(href.slice(1)) : undefined;
                }
                const colours = stops.map(s => {
                    const own = declarations(s.getAttribute('style') || '');
                    return { c: parseCssColor(own.get('stop-color') ?? s.getAttribute('stop-color') ?? 'black'), a: num(own.get('stop-opacity') ?? s.getAttribute('stop-opacity') ?? '1') };
                }).filter(s => s.c);
                if (!colours.length) return undefined;
                const mean = [0, 2, 4].map(i => colours.reduce((sum, s) => sum + parseInt(s.c!.hex.slice(i, i + 2), 16), 0) / colours.length / 255);
                const alpha = colours.reduce((sum, s) => sum + (Number.isFinite(s.a) ? s.a : 1) * s.c!.alpha, 0) / colours.length;
                return { color: lib.rgb(mean[0], mean[1], mean[2]), alpha };
            }
            if (kind === 'pattern') left.add('patterns');
            return ref[2] ? paint(ref[2], style) : undefined;
        }
        const c = parseCssColor(v);
        if (!c) return undefined;
        const n = parseInt(c.hex, 16);
        return { color: lib.rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255), alpha: Number.isFinite(c.alpha) ? c.alpha : 1 };
    };

    /** Draws path segments with the element's fill and stroke. */
    const drawPath = (segments: Segment[], style: Style, opacity: number) => {
        if (!segments.length || style.visibility === 'hidden' || style.visibility === 'collapse') return;
        const fill = paint(style.fill, style), stroke = style.strokeWidth > 0 ? paint(style.stroke, style) : undefined;
        if (!fill && !stroke) return;
        const ops: any[] = [lib.pushGraphicsState()];
        const fa = (fill?.alpha ?? 1) * style.fillOpacity * opacity, sa = (stroke?.alpha ?? 1) * style.strokeOpacity * opacity;
        if (fa < 1 || sa < 1) ops.push(op('gs', opacityState(fa, sa)));
        if (fill) ops.push(lib.setFillingRgbColor(fill.color.red, fill.color.green, fill.color.blue));
        if (stroke) {
            ops.push(lib.setStrokingRgbColor(stroke.color.red, stroke.color.green, stroke.color.blue), lib.setLineWidth(style.strokeWidth));
            ops.push(op('J', lib.asPDFNumber(style.lineCap === 'round' ? 1 : style.lineCap === 'square' ? 2 : 0)));
            ops.push(op('j', lib.asPDFNumber(style.lineJoin === 'round' ? 1 : style.lineJoin === 'bevel' ? 2 : 0)));
            const dash = style.dash !== 'none' ? numbers(style.dash, 32).filter(n => n >= 0) : [];
            if (dash.length && dash.some(n => n > 0)) ops.push(lib.setDashPattern(dash.length % 2 ? [...dash, ...dash] : dash, 0));
        }
        for (const s of segments) {
            if (s[0] === 'M') ops.push(lib.moveTo(s[1], s[2]));
            else if (s[0] === 'L') ops.push(lib.lineTo(s[1], s[2]));
            else if (s[0] === 'C') ops.push(lib.appendBezierCurve(s[1], s[2], s[3], s[4], s[5], s[6]));
            else ops.push(lib.closePath());
        }
        const evenOdd = style.fillRule === 'evenodd';
        ops.push(fill && stroke ? op(evenOdd ? 'B*' : 'B') : fill ? op(evenOdd ? 'f*' : 'f') : lib.stroke());
        ops.push(lib.popGraphicsState());
        page.pushOperators(...ops);
    };

    const faceOf = (style: Style): SvgFace => {
        const family = style.fontFamily.toLowerCase();
        const kind = /mono|courier|consol|menlo|code/.test(family) ? 'mono' : /serif/.test(family) && !/sans/.test(family) || /times|georgia|garamond|cambria/.test(family) ? 'serif' : '';
        const bold = /bold|bolder|[6-9]00/.test(style.fontWeight), italic = /italic|oblique/.test(style.fontStyle);
        const suffix = bold && italic ? 'BoldItalic' : bold ? 'Bold' : italic ? 'Italic' : '';
        if (!kind) return (suffix ? suffix[0].toLowerCase() + suffix.slice(1) : 'regular') as SvgFace;
        return (kind + suffix) as SvgFace;
    };

    /** Draws a text element: its runs (`tspan`s) one after another from its position, each at its anchor. */
    const drawText = (el: Element, style: Style, opacity: number) => {
        let x = length((el.getAttribute('x') || '').split(/[\s,]+/)[0], viewport.w, style.fontSize) || 0;
        let y = length((el.getAttribute('y') || '').split(/[\s,]+/)[0], viewport.h, style.fontSize) || 0;
        x += length((el.getAttribute('dx') || '').split(/[\s,]+/)[0], viewport.w, style.fontSize) || 0;
        y += length((el.getAttribute('dy') || '').split(/[\s,]+/)[0], viewport.h, style.fontSize) || 0;
        // The runs, each with its style and any position of its own.
        const runs: { text: string; style: Style; x?: number; y?: number }[] = [];
        const collect = (node: Element, s: Style, depth: number) => {
            for (let n = node.firstChild; n; n = n.nextSibling) {
                if (n.nodeType === 3 || n.nodeType === 4) {
                    const text = (n.nodeValue || '').replace(/\s+/g, ' ');
                    if (text) runs.push({ text, style: s });
                } else if (n.nodeType === 1 && localName(n as Element) === 'tspan' && depth < 8 && budget.elements-- > 0) {
                    const t = n as Element;
                    const { style: ts } = styleOf(t, s);
                    const tx = length((t.getAttribute('x') || '').split(/[\s,]+/)[0], viewport.w, ts.fontSize);
                    const ty = length((t.getAttribute('y') || '').split(/[\s,]+/)[0], viewport.h, ts.fontSize);
                    const at = runs.length;
                    collect(t, ts, depth + 1);
                    if (runs[at]) {
                        if (Number.isFinite(tx)) runs[at].x = tx;
                        if (Number.isFinite(ty)) runs[at].y = ty;
                        const dx = length((t.getAttribute('dx') || '').split(/[\s,]+/)[0], viewport.w, ts.fontSize), dy = length((t.getAttribute('dy') || '').split(/[\s,]+/)[0], viewport.h, ts.fontSize);
                        if (Number.isFinite(dx)) runs[at].x = (runs[at].x ?? NaN) + dx;
                        if (Number.isFinite(dy)) runs[at].y = (runs[at].y ?? NaN) + dy;
                    }
                }
            }
        };
        collect(el, style, 0);
        if (!runs.length) return;
        runs[0].text = runs[0].text.replace(/^ /, '');
        runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/ $/, '');
        // A chunk (runs from one position to the next) is anchored as a whole.
        const chunks: { x: number; y: number; runs: typeof runs }[] = [];
        for (const run of runs) {
            const own = run.x !== undefined || run.y !== undefined;
            if (!chunks.length || own) {
                const prev = chunks[chunks.length - 1];
                const endX = prev ? prev.x + prev.runs.reduce((w, r) => w + canvas.width(r.text, faceOf(r.style), r.style.fontSize), 0) : x;
                const rx = run.x === undefined ? endX : Number.isFinite(run.x) ? run.x : endX + (run.x ?? 0);
                const ry = run.y === undefined ? (prev ? prev.y : y) : Number.isFinite(run.y) ? run.y : (prev ? prev.y : y);
                chunks.push({ x: rx, y: ry, runs: [run] });
            } else chunks[chunks.length - 1].runs.push(run);
        }
        for (const chunk of chunks) {
            const width = chunk.runs.reduce((w, r) => w + canvas.width(r.text, faceOf(r.style), r.style.fontSize), 0);
            let cx = chunk.x - (style.textAnchor === 'middle' ? width / 2 : style.textAnchor === 'end' ? width : 0);
            for (const run of chunk.runs) {
                const fill = paint(run.style.fill, run.style);
                const face = faceOf(run.style);
                const w = canvas.width(run.text, face, run.style.fontSize);
                if (fill && run.text.trim()) {
                    const alpha = fill.alpha * run.style.fillOpacity * opacity;
                    page.pushOperators(lib.pushGraphicsState(), ...(alpha < 1 ? [op('gs', opacityState(alpha, alpha))] : []),
                        lib.concatTransformationMatrix(1, 0, 0, -1, cx, chunk.y));
                    try { canvas.text(run.text, face, 0, 0, run.style.fontSize, fill.color); } finally { page.pushOperators(lib.popGraphicsState()); }
                }
                cx += w;
            }
        }
    };

    /** Draws a picture given as data (`data:image/png;base64,...`) in its box. */
    const drawImage = async (el: Element, opacity: number) => {
        const href = el.getAttribute('href') || el.getAttribute('xlink:href') || '';
        const m = /^data:image\/(png|jpe?g);base64,([A-Za-z0-9+/=\s]*)$/i.exec(href.trim());
        if (!m) { left.add('pictures that are not PNG or JPEG data'); return; }
        const w = length(el.getAttribute('width'), viewport.w), h = length(el.getAttribute('height'), viewport.h);
        if (!(w > 0) || !(h > 0)) return;
        let bytes: Uint8Array;
        try { bytes = Uint8Array.from(Buffer.from(m[2].replace(/\s+/g, ''), 'base64')); } catch { return; }
        const img = await canvas.image(bytes, m[1].toLowerCase() === 'png' ? 'png' : 'jpeg');
        if (!img) { left.add('pictures that could not be read'); return; }
        const x = length(el.getAttribute('x'), viewport.w) || 0, y = length(el.getAttribute('y'), viewport.h) || 0;
        page.pushOperators(lib.pushGraphicsState(), ...(opacity < 1 ? [op('gs', opacityState(opacity, opacity))] : []),
            lib.concatTransformationMatrix(w, 0, 0, -h, x, y + h));
        try { page.drawImage(img, { x: 0, y: 0, width: 1, height: 1 }); } finally { page.pushOperators(lib.popGraphicsState()); }
    };

    const NOT_DRAWN = new Set(['defs', 'style', 'title', 'desc', 'metadata', 'symbol', 'linearGradient', 'radialGradient', 'pattern', 'clipPath', 'mask', 'filter', 'marker', 'script', 'foreignObject']);

    /** Draws an element and what it holds. */
    const draw = async (el: Element, parent: Style, depth: number, uses: number, active: Set<Element>): Promise<void> => {
        if (budget.elements-- <= 0 || depth > MAX_DEPTH) return;
        const tag = localName(el);
        if (NOT_DRAWN.has(tag)) { if (tag === 'foreignObject') left.add('foreign objects'); return; }
        const { style, own } = styleOf(el, parent);
        if (own.get('display') === 'none') return;
        if (own.has('filter') && own.get('filter') !== 'none') left.add('filters');
        if (own.has('mask') && own.get('mask') !== 'none') left.add('masks');
        if (own.has('clip-path') && own.get('clip-path') !== 'none') left.add('clip paths');
        const opacity = style.opacity * ((parent as Style & { inherited?: number }).inherited ?? 1);
        let matrix = transformOf(el.getAttribute('transform'));
        if (tag === 'svg' && depth > 0) {
            // A nested viewport: placed at x, y, its own viewBox scaled into its size.
            const x = length(el.getAttribute('x'), viewport.w) || 0, y = length(el.getAttribute('y'), viewport.h) || 0;
            const w = length(el.getAttribute('width'), viewport.w), h = length(el.getAttribute('height'), viewport.h);
            const vb = numbers(el.getAttribute('viewBox'), 4);
            matrix = multiply(matrix, [1, 0, 0, 1, x, y]);
            if (vb.length === 4 && vb[2] > 0 && vb[3] > 0 && w > 0 && h > 0) matrix = multiply(matrix, viewBoxMatrix(vb as [number, number, number, number], el.getAttribute('preserveAspectRatio') || 'xMidYMid meet', w, h));
        }
        const moved = matrix !== IDENTITY && matrix.some((v, i) => v !== IDENTITY[i]);
        if (moved) page.pushOperators(lib.pushGraphicsState(), lib.concatTransformationMatrix(...matrix));
        try {
            const inherited = { ...style, inherited: opacity } as Style;
            const L = (k: string, whole: number) => length(el.getAttribute(k), whole, style.fontSize);
            switch (tag) {
                case 'svg': case 'g': case 'a': case 'switch':
                    for (const child of childElements(el)) await draw(child, inherited, depth + 1, uses, active);
                    break;
                case 'path': drawPath(parsePath(el.getAttribute('d') || '', budget.path), style, opacity); break;
                case 'rect': {
                    const w = L('width', viewport.w), h = L('height', viewport.h);
                    if (w > 0 && h > 0) drawPath(rect(L('x', viewport.w) || 0, L('y', viewport.h) || 0, w, h, L('rx', viewport.w), L('ry', viewport.h)), style, opacity);
                    break;
                }
                case 'circle': {
                    const r = L('r', Math.hypot(viewport.w, viewport.h) / Math.SQRT2);
                    if (r > 0) drawPath(ellipse(L('cx', viewport.w) || 0, L('cy', viewport.h) || 0, r, r), style, opacity);
                    break;
                }
                case 'ellipse': {
                    const rx = L('rx', viewport.w), ry = L('ry', viewport.h);
                    if (rx > 0 && ry > 0) drawPath(ellipse(L('cx', viewport.w) || 0, L('cy', viewport.h) || 0, rx, ry), style, opacity);
                    break;
                }
                case 'line':
                    drawPath([['M', L('x1', viewport.w) || 0, L('y1', viewport.h) || 0], ['L', L('x2', viewport.w) || 0, L('y2', viewport.h) || 0]], { ...style, fill: 'none' }, opacity);
                    break;
                case 'polyline': case 'polygon': {
                    const p = numbers(el.getAttribute('points'), Math.min(2 * budget.path.left, 1 << 20));
                    budget.path.left -= p.length / 2;
                    const segs: Segment[] = [];
                    for (let i = 0; i + 1 < p.length; i += 2) segs.push([i ? 'L' : 'M', p[i], p[i + 1]] as Segment);
                    if (tag === 'polygon' && segs.length) segs.push(['Z']);
                    drawPath(segs, style, opacity);
                    break;
                }
                case 'text': drawText(el, style, opacity); break;
                case 'image': await drawImage(el, opacity); break;
                case 'use': {
                    const href = el.getAttribute('href') || el.getAttribute('xlink:href') || '';
                    const target = href.startsWith('#') ? svg.ids.get(href.slice(1)) : undefined;
                    if (!target || uses >= MAX_USE_DEPTH || active.has(target)) break;
                    const x = L('x', viewport.w) || 0, y = L('y', viewport.h) || 0;
                    page.pushOperators(lib.pushGraphicsState(), lib.concatTransformationMatrix(1, 0, 0, 1, x, y));
                    active.add(target);
                    try {
                        if (localName(target) === 'symbol') {
                            // A symbol is drawn as a group, its own viewBox scaled into the `use`'s size.
                            const vb = numbers(target.getAttribute('viewBox'), 4);
                            const w = L('width', viewport.w), h = L('height', viewport.h);
                            const fit = vb.length === 4 && vb[2] > 0 && vb[3] > 0 && w > 0 && h > 0;
                            if (fit) page.pushOperators(lib.pushGraphicsState(), lib.concatTransformationMatrix(...viewBoxMatrix(vb as [number, number, number, number], target.getAttribute('preserveAspectRatio') || 'xMidYMid meet', w, h)));
                            try { for (const child of childElements(target)) await draw(child, inherited, depth + 1, uses + 1, active); }
                            finally { if (fit) page.pushOperators(lib.popGraphicsState()); }
                        } else await draw(target, inherited, depth + 1, uses + 1, active);
                    } finally {
                        active.delete(target);
                        page.pushOperators(lib.popGraphicsState());
                    }
                    break;
                }
                default: break;
            }
        } finally {
            if (moved) page.pushOperators(lib.popGraphicsState());
        }
    };

    page.pushOperators(lib.pushGraphicsState(), lib.concatTransformationMatrix(...base));
    try { await draw(svg.root, INITIAL_STYLE, 0, 0, new Set()); }
    finally { page.pushOperators(lib.popGraphicsState()); }
    if (budget.elements <= 0 || budget.path.left <= 0) left.add('what was past the picture\'s limits');
    return [...left];
}

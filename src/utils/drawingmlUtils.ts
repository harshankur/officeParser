/**
 * DrawingML, as PowerPoint writes a slide's shapes: where a shape is (`a:xfrm`, through the groups it is
 * in), what it looks like (its outline, fill and line, its theme's colours) and where its text sits.
 * Lengths are EMU (12,700 to a point) in the file and points here.
 */
import { getDirectChildren, isElement } from './xmlUtils.js';

/** EMU in a point. */
export const EMU_PER_POINT = 12700;

/** A box in points, from the slide's top left corner. */
export interface Box { x: number; y: number; width: number; height: number }

/** How a group places what it holds: a box in its own space to a box on the slide. */
export type Transform = (box: Box) => Box;

/** No group: a box is where it says. */
export const IDENTITY: Transform = box => box;

/** The largest a slide's length is (56 inches), and so the largest a box in it is read to be. */
const MAX_POINTS = 4032 * 4;

/** An EMU length attribute in points, or undefined when it is none or past any slide. */
function points(element: Element | undefined, name: string): number | undefined {
    const raw = element?.getAttribute(name);
    if (!raw) return undefined;
    const value = Number(raw) / EMU_PER_POINT;
    return Number.isFinite(value) && Math.abs(value) <= MAX_POINTS ? value : undefined;
}

/** A shape's place (its `a:xfrm`, or a graphic frame's `p:xfrm`) and its turn, in points. */
export function xfrmOf(holder: Element | undefined): (Box & { rotation?: number; flipH?: boolean; flipV?: boolean }) | undefined {
    const xfrm = holder ? (getDirectChildren(holder, 'a:xfrm')[0] ?? getDirectChildren(holder, 'p:xfrm')[0]) : undefined;
    if (!xfrm) return undefined;
    const off = getDirectChildren(xfrm, 'a:off')[0], ext = getDirectChildren(xfrm, 'a:ext')[0];
    const x = points(off, 'x'), y = points(off, 'y'), width = points(ext, 'cx'), height = points(ext, 'cy');
    if (x === undefined || y === undefined || width === undefined || height === undefined || width < 0 || height < 0) return undefined;
    const rot = Number(xfrm.getAttribute('rot')) / 60000;
    return {
        x, y, width, height,
        ...(Number.isFinite(rot) && rot % 360 !== 0 ? { rotation: ((rot % 360) + 360) % 360 } : {}),
        ...(xfrm.getAttribute('flipH') === '1' ? { flipH: true } : {}),
        ...(xfrm.getAttribute('flipV') === '1' ? { flipV: true } : {}),
    };
}

/**
 * The transform of a group (`p:grpSp`): its children's space (`a:chOff`, `a:chExt`) mapped onto its
 * box (`a:off`, `a:ext`), then by the groups it is in.
 */
export function groupTransform(grpSpPr: Element | undefined, outer: Transform): Transform {
    const xfrm = grpSpPr ? getDirectChildren(grpSpPr, 'a:xfrm')[0] : undefined;
    if (!xfrm) return outer;
    const off = getDirectChildren(xfrm, 'a:off')[0], ext = getDirectChildren(xfrm, 'a:ext')[0];
    const chOff = getDirectChildren(xfrm, 'a:chOff')[0], chExt = getDirectChildren(xfrm, 'a:chExt')[0];
    const x = points(off, 'x') ?? 0, y = points(off, 'y') ?? 0;
    const cx = points(chOff, 'x') ?? x, cy = points(chOff, 'y') ?? y;
    const w = points(ext, 'cx'), h = points(ext, 'cy'), cw = points(chExt, 'cx'), ch = points(chExt, 'cy');
    const sx = w && cw ? w / cw : 1, sy = h && ch ? h / ch : 1;
    return box => outer({ x: x + (box.x - cx) * sx, y: y + (box.y - cy) * sy, width: box.width * sx, height: box.height * sy });
}

/** A box rounded to a hundredth of a point, as the AST keeps boxes. */
export function roundBox(box: Box): Box {
    const round = (v: number) => Math.round(v * 100) / 100;
    return { x: round(box.x), y: round(box.y), width: round(box.width), height: round(box.height) };
}

/** A theme's colours (`a:clrScheme`) by name (`dk1`, `lt1`, `accent1`, ...), as six hex digits. */
export function themeColors(theme: Document | Element | undefined): Map<string, string> {
    const colors = new Map<string, string>();
    const root = theme && 'documentElement' in theme ? theme.documentElement : theme;
    const scheme = root ? findFirst(root, 'a:clrScheme', 4) : undefined;
    if (!scheme) return colors;
    for (const entry of Array.from(scheme.childNodes).filter(isElement)) {
        const name = entry.tagName.replace(/^a:/, '');
        const hex = rgbOf(Array.from(entry.childNodes).filter(isElement)[0], colors);
        if (hex) colors.set(name, hex);
    }
    return colors;
}

/** The first element named `tag` within `depth` levels of `root`. */
function findFirst(root: Element, tag: string, depth: number): Element | undefined {
    if (root.tagName === tag) return root;
    if (depth <= 0) return undefined;
    for (const child of Array.from(root.childNodes).filter(isElement)) {
        const found = findFirst(child, tag, depth - 1);
        if (found) return found;
    }
    return undefined;
}

/** The scheme names a shape's colours use, as the theme names them (the default colour map). */
const SCHEME_NAMES: Record<string, string> = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };

/** The few preset colours (`a:prstClr`) a shape is likely to use. */
const PRESET_COLORS: Record<string, string> = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080' };

/**
 * A colour element's six hex digits (`a:srgbClr`, `a:schemeClr` through the theme, `a:sysClr`'s last
 * colour, `a:prstClr`), with its luminance, tint and shade applied.
 */
export function rgbOf(color: Element | undefined, theme: Map<string, string>): string | undefined {
    if (!color) return undefined;
    let hex: string | undefined;
    const val = color.getAttribute('val') || '';
    switch (color.tagName) {
        case 'a:srgbClr': hex = val; break;
        case 'a:schemeClr': hex = theme.get(SCHEME_NAMES[val] ?? val); break;
        case 'a:sysClr': hex = color.getAttribute('lastClr') || (val === 'window' ? 'FFFFFF' : val === 'windowText' ? '000000' : undefined); break;
        case 'a:prstClr': hex = PRESET_COLORS[val]; break;
        case 'a:scrgbClr': {
            const pct = (name: string) => Math.round(Math.min(100000, Math.max(0, Number(color.getAttribute(name)) || 0)) / 100000 * 255);
            hex = [pct('r'), pct('g'), pct('b')].map(v => v.toString(16).padStart(2, '0')).join('');
            break;
        }
    }
    if (!hex || !/^[0-9a-f]{6}$/i.test(hex)) return undefined;
    let [h, s, l] = toHsl(parseInt(hex, 16));
    let [r, g, b] = [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
    for (const modifier of Array.from(color.childNodes).filter(isElement)) {
        const amount = Number(modifier.getAttribute('val')) / 100000;
        if (!Number.isFinite(amount)) continue;
        switch (modifier.tagName) {
            case 'a:lumMod': [h, s, l] = toHsl((r << 16) | (g << 8) | b); l = Math.min(1, l * amount); [r, g, b] = fromHsl(h, s, l); break;
            case 'a:lumOff': [h, s, l] = toHsl((r << 16) | (g << 8) | b); l = Math.min(1, Math.max(0, l + amount)); [r, g, b] = fromHsl(h, s, l); break;
            case 'a:tint': [r, g, b] = [r, g, b].map(v => Math.round(255 - (255 - v) * amount)) as [number, number, number]; break;
            case 'a:shade': [r, g, b] = [r, g, b].map(v => Math.round(v * amount)) as [number, number, number]; break;
        }
    }
    return [r, g, b].map(v => Math.max(0, Math.min(255, v)).toString(16).padStart(2, '0')).join('').toUpperCase();
}

function toHsl(rgb: number): [number, number, number] {
    const r = ((rgb >> 16) & 255) / 255, g = ((rgb >> 8) & 255) / 255, b = (rgb & 255) / 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2;
    if (max === min) return [0, 0, l];
    const d = max - min, s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    return [h / 6, s, l];
}

function fromHsl(h: number, s: number, l: number): [number, number, number] {
    if (s === 0) return [l, l, l].map(v => Math.round(v * 255)) as [number, number, number];
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    const hue = (t: number) => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
    return [hue(h + 1 / 3), hue(h), hue(h - 1 / 3)].map(v => Math.round(v * 255)) as [number, number, number];
}

/** The colour a fill element gives (`a:solidFill`, a gradient's first stop), or `null` for none (`a:noFill`). */
function fillColor(holder: Element | undefined, theme: Map<string, string>): string | null | undefined {
    if (!holder) return undefined;
    if (getDirectChildren(holder, 'a:noFill').length) return null;
    const solid = getDirectChildren(holder, 'a:solidFill')[0];
    if (solid) return rgbOf(Array.from(solid.childNodes).filter(isElement)[0], theme) ?? undefined;
    const gradient = getDirectChildren(holder, 'a:gradFill')[0];
    const stop = gradient ? getDirectChildren(getDirectChildren(gradient, 'a:gsLst')[0] ?? gradient, 'a:gs')[0] : undefined;
    if (stop) return rgbOf(Array.from(stop.childNodes).filter(isElement)[0], theme) ?? undefined;
    return undefined;
}

/**
 * What a shape looks like: its outline (`a:prstGeom`), its fill and its line, each its own (`p:spPr`)
 * or else its style's (`p:style`'s `a:fillRef` and `a:lnRef`, in the colour they name).
 */
export function shapeLook(shape: Element, theme: Map<string, string>): { geometry?: string; fill?: string; line?: { color?: string; width?: number } } {
    const spPr = getDirectChildren(shape, 'p:spPr')[0];
    const style = getDirectChildren(shape, 'p:style')[0];
    const geometry = spPr ? (getDirectChildren(spPr, 'a:prstGeom')[0]?.getAttribute('prst') || (getDirectChildren(spPr, 'a:custGeom').length ? 'custom' : undefined)) : undefined;
    const styleRef = (tag: string) => {
        const ref = style ? getDirectChildren(style, tag)[0] : undefined;
        return ref && Number(ref.getAttribute('idx')) > 0 ? rgbOf(Array.from(ref.childNodes).filter(isElement)[0], theme) : undefined;
    };
    const ownFill = fillColor(spPr, theme);
    const fill = ownFill === null ? undefined : ownFill ?? styleRef('a:fillRef');
    const ln = spPr ? getDirectChildren(spPr, 'a:ln')[0] : undefined;
    const ownLine = fillColor(ln, theme);
    const lineColor = ownLine === null ? undefined : ownLine ?? styleRef('a:lnRef');
    const lineWidth = points(ln, 'w');
    return {
        ...(geometry ? { geometry } : {}),
        ...(fill ? { fill: `#${fill}` } : {}),
        // (A width of none is the thinnest line, as PowerPoint draws it.)
        ...(lineColor ? { line: { color: `#${lineColor}`, width: lineWidth || (lineWidth === 0 ? 0.25 : 0.75) } } : {}),
    };
}

/** Where a shape's text sits in it (`a:bodyPr`): its anchor, and its insets in points (DrawingML's defaults 7.2 and 3.6). */
export function textFrameOf(bodyPr: Element | undefined): { anchor?: 'top' | 'middle' | 'bottom'; insets: { top: number; right: number; bottom: number; left: number }; vertical?: boolean } {
    const anchor = bodyPr?.getAttribute('anchor');
    const vert = bodyPr?.getAttribute('vert');
    return {
        ...(anchor === 'ctr' ? { anchor: 'middle' as const } : anchor === 'b' ? { anchor: 'bottom' as const } : anchor === 't' ? { anchor: 'top' as const } : {}),
        insets: { top: points(bodyPr, 'tIns') ?? 3.6, right: points(bodyPr, 'rIns') ?? 7.2, bottom: points(bodyPr, 'bIns') ?? 3.6, left: points(bodyPr, 'lIns') ?? 7.2 },
        ...(vert === 'vert' || vert === 'eaVert' || vert === 'vert270' ? { vertical: true } : {}),
    };
}

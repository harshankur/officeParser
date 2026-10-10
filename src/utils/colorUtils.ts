import { lookupTable } from './lookupUtils.js';

/** The CSS named colours (CSS Color Module Level 4), as six hex digits. */
const NAMED_COLORS: Record<string, string> = lookupTable({
    aliceblue: 'F0F8FF', antiquewhite: 'FAEBD7', aqua: '00FFFF', aquamarine: '7FFFD4', azure: 'F0FFFF',
    beige: 'F5F5DC', bisque: 'FFE4C4', black: '000000', blanchedalmond: 'FFEBCD', blue: '0000FF',
    blueviolet: '8A2BE2', brown: 'A52A2A', burlywood: 'DEB887', cadetblue: '5F9EA0', chartreuse: '7FFF00',
    chocolate: 'D2691E', coral: 'FF7F50', cornflowerblue: '6495ED', cornsilk: 'FFF8DC', crimson: 'DC143C',
    cyan: '00FFFF', darkblue: '00008B', darkcyan: '008B8B', darkgoldenrod: 'B8860B', darkgray: 'A9A9A9',
    darkgreen: '006400', darkgrey: 'A9A9A9', darkkhaki: 'BDB76B', darkmagenta: '8B008B', darkolivegreen: '556B2F',
    darkorange: 'FF8C00', darkorchid: '9932CC', darkred: '8B0000', darksalmon: 'E9967A', darkseagreen: '8FBC8F',
    darkslateblue: '483D8B', darkslategray: '2F4F4F', darkslategrey: '2F4F4F', darkturquoise: '00CED1', darkviolet: '9400D3',
    deeppink: 'FF1493', deepskyblue: '00BFFF', dimgray: '696969', dimgrey: '696969', dodgerblue: '1E90FF',
    firebrick: 'B22222', floralwhite: 'FFFAF0', forestgreen: '228B22', fuchsia: 'FF00FF', gainsboro: 'DCDCDC',
    ghostwhite: 'F8F8FF', gold: 'FFD700', goldenrod: 'DAA520', gray: '808080', green: '008000',
    greenyellow: 'ADFF2F', grey: '808080', honeydew: 'F0FFF0', hotpink: 'FF69B4', indianred: 'CD5C5C',
    indigo: '4B0082', ivory: 'FFFFF0', khaki: 'F0E68C', lavender: 'E6E6FA', lavenderblush: 'FFF0F5',
    lawngreen: '7CFC00', lemonchiffon: 'FFFACD', lightblue: 'ADD8E6', lightcoral: 'F08080', lightcyan: 'E0FFFF',
    lightgoldenrodyellow: 'FAFAD2', lightgray: 'D3D3D3', lightgreen: '90EE90', lightgrey: 'D3D3D3', lightpink: 'FFB6C1',
    lightsalmon: 'FFA07A', lightseagreen: '20B2AA', lightskyblue: '87CEFA', lightslategray: '778899', lightslategrey: '778899',
    lightsteelblue: 'B0C4DE', lightyellow: 'FFFFE0', lime: '00FF00', limegreen: '32CD32', linen: 'FAF0E6',
    magenta: 'FF00FF', maroon: '800000', mediumaquamarine: '66CDAA', mediumblue: '0000CD', mediumorchid: 'BA55D3',
    mediumpurple: '9370DB', mediumseagreen: '3CB371', mediumslateblue: '7B68EE', mediumspringgreen: '00FA9A', mediumturquoise: '48D1CC',
    mediumvioletred: 'C71585', midnightblue: '191970', mintcream: 'F5FFFA', mistyrose: 'FFE4E1', moccasin: 'FFE4B5',
    navajowhite: 'FFDEAD', navy: '000080', oldlace: 'FDF5E6', olive: '808000', olivedrab: '6B8E23',
    orange: 'FFA500', orangered: 'FF4500', orchid: 'DA70D6', palegoldenrod: 'EEE8AA', palegreen: '98FB98',
    paleturquoise: 'AFEEEE', palevioletred: 'DB7093', papayawhip: 'FFEFD5', peachpuff: 'FFDAB9', peru: 'CD853F',
    pink: 'FFC0CB', plum: 'DDA0DD', powderblue: 'B0E0E6', purple: '800080', rebeccapurple: '663399',
    red: 'FF0000', rosybrown: 'BC8F8F', royalblue: '4169E1', saddlebrown: '8B4513', salmon: 'FA8072',
    sandybrown: 'F4A460', seagreen: '2E8B57', seashell: 'FFF5EE', sienna: 'A0522D', silver: 'C0C0C0',
    skyblue: '87CEEB', slateblue: '6A5ACD', slategray: '708090', slategrey: '708090', snow: 'FFFAFA',
    springgreen: '00FF7F', steelblue: '4682B4', tan: 'D2B48C', teal: '008080', thistle: 'D8BFD8',
    tomato: 'FF6347', turquoise: '40E0D0', violet: 'EE82EE', wheat: 'F5DEB3', white: 'FFFFFF',
    whitesmoke: 'F5F5F5', yellow: 'FFFF00', yellowgreen: '9ACD32',
});

/** A number from 0 to 255 as two hex digits. */
const hexByte = (n: number): string => Math.round(Math.min(255, Math.max(0, n))).toString(16).padStart(2, '0').toUpperCase();

/** A CSS number or percentage (of `whole`) as a number, NaN when it is neither. */
const cssNumber = (part: string | undefined, whole: number): number => {
    if (part === undefined || !/^[+-]?(?:\d+\.?\d*|\.\d+)%?$/.test(part)) return NaN;
    return part.endsWith('%') ? parseFloat(part) * whole / 100 : parseFloat(part);
};

/**
 * A colour as the AST holds it (a CSS colour: an HTML or Markdown document's own, or `#RRGGBB` from the
 * office formats) as six upper-case hex digits: `#rgb`, `#rrggbb` (or without the `#`), `#rgba` and
 * `#rrggbbaa` (the alpha left out), a named colour, `rgb()`/`rgba()` and `hsl()`/`hsla()`. Null for none
 * of those (`transparent`, `currentColor`, `inherit`, ...): a writer then gives the text no colour of
 * its own. Read as `#RRGGBB` alone, a named colour was dropped by the ODT writer and wrote `NaN` into an
 * RTF colour table, which RTF readers refuse whole.
 */
export function cssColorHex(value: unknown): string | null {
    return parseCssColor(value)?.hex ?? null;
}

/**
 * Whether a background colour is the one a highlight has when it names none (Markdown's `==text==`, a
 * bare `<mark>`): opaque yellow, however it is written (`#ffff00`, `#FF0`, `yellow`,
 * `rgb(255, 255, 0)`). Compared as the text `#ffff00` alone, the same highlight come back from an
 * editor as `rgb(255, 255, 0)`, or read from a Word file as `#FFFF00`, was taken for a colour of its
 * own. A yellow that is partly see-through (`rgba(255, 255, 0, 0.3)`, `#ffff0080`) is a colour of its
 * own: written as the default, it came out solid.
 */
export function isDefaultHighlight(color: unknown): boolean {
    const parsed = parseCssColor(color);
    return parsed?.hex === 'FFFF00' && parsed.alpha === 1;
}

/**
 * A CSS colour as six upper-case hex digits and its alpha, from 0 to 1 (1 for a colour that names
 * none, NaN for an alpha that is not a number or a percentage). Null for what is not a colour (see
 * cssColorHex).
 */
function parseCssColor(value: unknown): { hex: string, alpha: number } | null {
    if (typeof value !== 'string') return null;
    const v = value.trim().toLowerCase();
    // Longer than any colour: not read, so a hostile value costs nothing.
    if (!v || v.length > 64) return null;
    const hex = /^#?([0-9a-f]{3,8})$/.exec(v);
    if (hex) {
        const h = hex[1];
        if (h.length === 3 || h.length === 4) {
            return { hex: (h[0] + h[0] + h[1] + h[1] + h[2] + h[2]).toUpperCase(), alpha: h.length === 4 ? parseInt(h[3] + h[3], 16) / 255 : 1 };
        }
        if (h.length !== 6 && h.length !== 8) return null;
        return { hex: h.slice(0, 6).toUpperCase(), alpha: h.length === 8 ? parseInt(h.slice(6), 16) / 255 : 1 };
    }
    const named = NAMED_COLORS[v];
    if (named) return { hex: named, alpha: 1 };
    const fn = /^(rgba?|hsla?)\(([^()]*)\)$/.exec(v);
    if (!fn) return null;
    const parts = fn[2].split(/[\s,/]+/).filter(Boolean);
    const alpha = parts[3] === undefined ? 1 : Math.min(1, Math.max(0, cssNumber(parts[3], 1)));
    if (fn[1].startsWith('rgb')) {
        const [r, g, b] = [0, 1, 2].map(i => cssNumber(parts[i], 255));
        return [r, g, b].some(isNaN) ? null : { hex: hexByte(r) + hexByte(g) + hexByte(b), alpha };
    }
    const hue = parseFloat(parts[0] ?? '');
    const saturation = cssNumber(parts[1], 1), lightness = cssNumber(parts[2], 1);
    if (isNaN(hue) || isNaN(saturation) || isNaN(lightness)) return null;
    const h = (((hue % 360) + 360) % 360) / 360;
    const s = Math.min(1, Math.max(0, saturation)), l = Math.min(1, Math.max(0, lightness));
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    const channel = (t: number) => {
        const u = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
        return 255 * (u < 1 / 6 ? p + (q - p) * 6 * u : u < 1 / 2 ? q : u < 2 / 3 ? p + (q - p) * (2 / 3 - u) * 6 : p);
    };
    return { hex: hexByte(channel(h + 1 / 3)) + hexByte(channel(h)) + hexByte(channel(h - 1 / 3)), alpha };
}

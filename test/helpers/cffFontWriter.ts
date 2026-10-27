/**
 * A CFF font writer for tests: it writes a CFF table of the tests' own Type 2 charstrings (Adobe's
 * Technical Notes #5176 and #5177), CID-keyed or not, and gives it to a TrueType font in place of its
 * outlines, as an OpenType font (`OTTO`), so the native PDF engine's reading of CFF glyphs is tested on
 * Liberation Sans (shipped with pdfjs-dist) rather than on font files committed for it.
 */

/** Type 2 charstring operators by name: an escaped one as its two bytes. */
const OPERATORS: Record<string, number[]> = {
    hstem: [1], vstem: [3], vmoveto: [4], rlineto: [5], hlineto: [6], vlineto: [7], rrcurveto: [8], callsubr: [10],
    return: [11], endchar: [14], hstemhm: [18], hintmask: [19], cntrmask: [20], rmoveto: [21], hmoveto: [22],
    vstemhm: [23], rcurveline: [24], rlinecurve: [25], vvcurveto: [26], hhcurveto: [27], callgsubr: [29],
    vhcurveto: [30], hvcurveto: [31], dotsection: [12, 0], hflex: [12, 34], flex: [12, 35], hflex1: [12, 36],
    flex1: [12, 37], dup: [12, 27], drop: [12, 18], exch: [12, 28], add: [12, 10],
};

/** An integer as a Type 2 charstring operand. */
function operand(v: number): number[] {
    if (v >= -107 && v <= 107) return [v + 139];
    if (v >= 108 && v <= 1131) return [((v - 108) >> 8) + 247, (v - 108) & 0xFF];
    if (v >= -1131 && v <= -108) return [((-v - 108) >> 8) + 251, (-v - 108) & 0xFF];
    return [28, (v >> 8) & 0xFF, v & 0xFF];
}

/**
 * A Type 2 charstring: each number an operand, each name an operator, each array of numbers bytes as
 * they are (a hint mask's).
 */
export function charString(...parts: (number | string | number[])[]): number[] {
    const out: number[] = [];
    for (const part of parts) {
        if (typeof part === 'number') out.push(...operand(part));
        else if (typeof part === 'string') {
            const op = OPERATORS[part];
            if (!op) throw new Error(`no charstring operator ${part}`);
            out.push(...op);
        } else out.push(...part);
    }
    return out;
}

/** A glyph that draws a box, and one that draws nothing. */
export const BOX = charString(50, 0, 'rmoveto', 400, 0, 'rlineto', 0, 600, 'rlineto', -400, 0, 'rlineto', 'endchar');
export const EMPTY = charString('endchar');

/** An INDEX of these items: their count, offsets four bytes each, and data. */
function index(items: ArrayLike<number>[]): number[] {
    if (!items.length) return [0, 0];
    const out = [items.length >> 8, items.length & 0xFF, 4];
    let offset = 1;
    const u32 = (v: number) => out.push((v >>> 24) & 0xFF, (v >>> 16) & 0xFF, (v >>> 8) & 0xFF, v & 0xFF);
    u32(offset);
    for (const item of items) u32(offset += item.length);
    for (const item of items) for (let i = 0; i < item.length; i++) out.push(item[i]);
    return out;
}

/** A DICT: each operator with its integer operands, each written in five bytes (so a DICT's length does not depend on its offsets). */
function dict(entries: [number[], number[]][]): number[] {
    const out: number[] = [];
    for (const [op, operands] of entries) {
        for (const v of operands) out.push(29, (v >>> 24) & 0xFF, (v >>> 16) & 0xFF, (v >>> 8) & 0xFF, v & 0xFF);
        out.push(...op);
    }
    return out;
}

const bytesOf = (text: string) => [...text].map(c => c.charCodeAt(0));

export interface CffSpec {
    /** Each glyph's charstring, by glyph id. */
    charStrings: number[][];
    globalSubrs?: number[][];
    /** The local subroutines of a font that is not CID-keyed. */
    subrs?: number[][];
    /** The standard string (SID) each of these glyphs is named by (seac finds its glyphs by these names); the others are named g1, g2 and on. */
    names?: Map<number, number>;
    /** A CID-keyed font: each font dictionary's local subroutines, and the FDSelect ranges, each its first glyph and font dictionary. */
    cid?: { fds: number[][][]; ranges: [number, number][] };
}

/** A CFF table (version 1) of these glyphs. */
export function cffTable(spec: CffSpec): Uint8Array {
    const count = spec.charStrings.length;
    const strings: string[] = [];
    const sid = (text: string) => 391 + strings.push(text) - 1;
    let charset: number[];
    if (spec.cid) charset = [2, 0, 1, (count - 2) >> 8, (count - 2) & 0xFF]; // CIDs 1 to count - 1, in order
    else {
        charset = [0];
        for (let g = 1; g < count; g++) { const s = spec.names?.get(g) ?? sid(`g${g}`); charset.push(s >> 8, s & 0xFF); }
    }
    const ros = spec.cid ? [sid('Adobe'), sid('Identity'), 0] : [];
    const privates = spec.cid ? spec.cid.fds : [spec.subrs ?? []];
    // (Laid out once with offsets of zero, then again with the offsets found: each is five bytes either way.)
    const layout = (at: Record<string, number>) => {
        const top: [number[], number[]][] = [[[15], [at.charset]], [[17], [at.charStrings]]];
        if (spec.cid) {
            top.unshift([[12, 30], ros]);
            top.push([[12, 36], [at.fdArray]], [[12, 37], [at.fdSelect]]);
        } else top.push([[18], [at.privateSize0, at.private0]]);
        const head = [1, 0, 4, 4, ...index([bytesOf('Test')]), ...index([dict(top)]), ...index(strings.map(bytesOf)), ...index(spec.globalSubrs ?? [])];
        const parts: number[] = [...head];
        // (Byte by byte: a long charstring is more bytes than one call takes arguments.)
        const place = (name: string, bytes: number[]) => { at[name] = parts.length; for (const byte of bytes) parts.push(byte); };
        place('charset', charset);
        place('charStrings', index(spec.charStrings));
        privates.forEach((subrs, i) => {
            // (A Private DICT naming its Subrs, which follow it, offset from its start.)
            const priv = subrs.length ? dict([[[19], [6]]]) : [];
            at[`privateSize${i}`] = priv.length;
            place(`private${i}`, [...priv, ...(subrs.length ? index(subrs) : [])]);
        });
        if (spec.cid) {
            place('fdArray', index(privates.map((_, i) => dict([[[18], [at[`privateSize${i}`], at[`private${i}`]]]]))));
            const ranges = spec.cid.ranges;
            place('fdSelect', [3, ranges.length >> 8, ranges.length & 0xFF, ...ranges.flatMap(([first, fd]) => [first >> 8, first & 0xFF, fd]), count >> 8, count & 0xFF]);
        }
        return parts;
    };
    const at: Record<string, number> = {};
    layout(at);
    return Uint8Array.from(layout(at));
}

/** The tables of a TrueType font that are its outlines, or work on them. */
const TRUETYPE_TABLES = new Set(['glyf', 'loca', 'cvt ', 'fpgm', 'prep', 'gasp', 'hdmx', 'LTSH', 'VDMX']);

/** A TrueType font of as many glyphs as the CFF table has, with this CFF table in place of its outlines: an OpenType font of CFF outlines. */
export function withCffOutlines(font: Uint8Array, cff: Uint8Array): Uint8Array {
    const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
    const tables = new Map<string, Uint8Array>();
    for (let i = 0; i < view.getUint16(4); i++) {
        const at = 12 + i * 16;
        const tag = String.fromCharCode(...font.subarray(at, at + 4));
        if (!TRUETYPE_TABLES.has(tag)) tables.set(tag, font.slice(view.getUint32(at + 8), view.getUint32(at + 8) + view.getUint32(at + 12)));
    }
    // (A CFF font's maxp is of version 0.5: its number of glyphs alone.)
    const maxp = tables.get('maxp')!;
    tables.set('maxp', Uint8Array.from([0, 0, 0x50, 0, maxp[4], maxp[5]]));
    tables.set('CFF ', cff);
    const tags = [...tables.keys()].sort();
    let offset = 12 + 16 * tags.length;
    const out = new Uint8Array(offset + tags.reduce((sum, tag) => sum + ((tables.get(tag)!.length + 3) & ~3), 0));
    const o = new DataView(out.buffer);
    o.setUint32(0, 0x4F54544F); o.setUint16(4, tags.length);
    let power = 1, log = 0;
    while (power * 2 <= tags.length) { power *= 2; log++; }
    o.setUint16(6, power * 16); o.setUint16(8, log); o.setUint16(10, tags.length * 16 - power * 16);
    tags.forEach((tag, i) => {
        const data = tables.get(tag)!;
        let sum = 0;
        for (let p = 0; p < data.length; p += 4) sum = (sum + (((data[p] << 24) | ((data[p + 1] ?? 0) << 16) | ((data[p + 2] ?? 0) << 8) | (data[p + 3] ?? 0)) >>> 0)) >>> 0;
        out.set([...tag].map(c => c.charCodeAt(0)), 12 + i * 16);
        o.setUint32(16 + i * 16, sum); o.setUint32(20 + i * 16, offset); o.setUint32(24 + i * 16, data.length);
        out.set(data, offset);
        offset += (data.length + 3) & ~3;
    });
    return out;
}

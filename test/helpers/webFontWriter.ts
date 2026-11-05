/**
 * A WOFF and WOFF2 writer for tests: it turns a TrueType or OpenType font into the web's formats, so the
 * native PDF engine's unpacking of them is tested on the fonts already in the repository (Liberation
 * Sans, shipped with pdfjs-dist) rather than on font files committed for it.
 *
 * WOFF compresses each table with zlib. WOFF2 compresses all of them together with Brotli (Node's
 * `zlib`), and first writes the outlines (`glyf` and `loca`), and the left side bearings of `hmtx` where
 * they are the glyphs' left edges, in the smaller forms of the W3C WOFF2 specification, section 5, as
 * the reference encoder does: each point's movement in the shortest of its 128 triplet encodings, a
 * bounding box only where it is not the points' own.
 */

import * as zlib from 'zlib';
import { zlibSync } from 'fflate';

/** The tags a WOFF2 table directory names by number, in that order. */
const KNOWN_TAGS = [
    'cmap', 'head', 'hhea', 'hmtx', 'maxp', 'name', 'OS/2', 'post', 'cvt ', 'fpgm', 'glyf', 'loca', 'prep',
    'CFF ', 'VORG', 'EBDT', 'EBLC', 'gasp', 'hdmx', 'kern', 'LTSH', 'PCLT', 'VDMX', 'vhea', 'vmtx', 'BASE',
    'GDEF', 'GPOS', 'GSUB', 'EBSC', 'JSTF', 'MATH', 'CBDT', 'CBLC', 'COLR', 'CPAL', 'SVG ', 'sbix', 'acnt',
    'avar', 'bdat', 'bloc', 'bsln', 'cvar', 'fdsc', 'feat', 'fmtx', 'fvar', 'gvar', 'hsty', 'just', 'lcar',
    'mort', 'morx', 'opbd', 'prop', 'trak', 'Zapf', 'Silf', 'Glat', 'Gloc', 'Feat', 'Sill',
];

/** Bytes written big-endian, growing as needed. */
class Out {
    bytes: number[] = [];
    u8(v: number): this { this.bytes.push(v & 0xFF); return this; }
    u16(v: number): this { return this.u8(v >> 8).u8(v); }
    u32(v: number): this { return this.u16(v >>> 16).u16(v & 0xFFFF); }
    put(data: ArrayLike<number>): this { for (let i = 0; i < data.length; i++) this.bytes.push(data[i]); return this; }
    pad(): this { while (this.bytes.length % 4) this.bytes.push(0); return this; }
    base128(v: number): this {
        const groups: number[] = [];
        do { groups.unshift((v & 0x7F) | (groups.length ? 0x80 : 0)); v = Math.floor(v / 128); } while (v);
        return this.put(groups);
    }
    u255(v: number): this {
        if (v < 253) return this.u8(v);
        if (v < 506) return this.u8(255).u8(v - 253);
        if (v < 762) return this.u8(254).u8(v - 506);
        return this.u8(253).u16(v);
    }
    done(): Uint8Array { return Uint8Array.from(this.bytes); }
}

/** The tables of an sfnt, by tag, and its flavor. */
function readSfnt(font: Uint8Array): { flavor: number; tables: Map<string, Uint8Array> } {
    const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
    const tables = new Map<string, Uint8Array>();
    for (let i = 0; i < view.getUint16(4); i++) {
        const at = 12 + i * 16;
        tables.set(String.fromCharCode(...font.subarray(at, at + 4)), font.subarray(view.getUint32(at + 8), view.getUint32(at + 8) + view.getUint32(at + 12)));
    }
    return { flavor: view.getUint32(0), tables };
}

function checksum(data: Uint8Array): number {
    let sum = 0;
    for (let i = 0; i < data.length; i += 4) sum = (sum + (((data[i] << 24) | ((data[i + 1] ?? 0) << 16) | ((data[i + 2] ?? 0) << 8) | (data[i + 3] ?? 0)) >>> 0)) >>> 0;
    return sum;
}

/** A WOFF file of a font: each table zlib-compressed where that makes it smaller. */
export function toWoff(font: Uint8Array): Uint8Array {
    const { flavor, tables } = readSfnt(font);
    const tags = [...tables.keys()].sort();
    const stored = tags.map(tag => {
        const data = tables.get(tag)!;
        const packed = zlibSync(data, { level: 9 });
        return packed.length < data.length ? packed : data;
    });
    const out = new Out();
    let offset = 44 + 20 * tags.length;
    const offsets = stored.map(s => { const at = offset; offset += (s.length + 3) & ~3; return at; });
    const sfntSize = 12 + 16 * tags.length + tags.reduce((n, tag) => n + ((tables.get(tag)!.length + 3) & ~3), 0);
    out.u32(0x774F4646).u32(flavor).u32(offset).u16(tags.length).u16(0).u32(sfntSize).u16(1).u16(0).u32(0).u32(0).u32(0).u32(0).u32(0);
    tags.forEach((tag, i) => {
        out.put([...tag].map(c => c.charCodeAt(0))).u32(offsets[i]).u32(stored[i].length).u32(tables.get(tag)!.length).u32(checksum(tables.get(tag)!));
    });
    for (const s of stored) out.put(s).pad();
    return out.done();
}

/**
 * The movement to a point in the shortest of the WOFF2 triplet encodings: the flag (its top bit set
 * for a point off the curve) and the bytes that follow it in the glyph stream.
 */
function triplet(dx: number, dy: number, onCurve: boolean): { flag: number; bytes: number[] } {
    const ax = Math.abs(dx), ay = Math.abs(dy);
    const on = onCurve ? 0 : 128, xs = dx < 0 ? 0 : 1, ys = dy < 0 ? 0 : 1, xys = xs + 2 * ys;
    if (dx === 0 && ay < 1280) return { flag: on + ((ay & 0xF00) >> 7) + ys, bytes: [ay & 0xFF] };
    if (dy === 0 && ax < 1280) return { flag: on + 10 + ((ax & 0xF00) >> 7) + xs, bytes: [ax & 0xFF] };
    if (ax < 65 && ay < 65) return { flag: on + 20 + ((ax - 1) & 0x30) + (((ay - 1) & 0x30) >> 2) + xys, bytes: [(((ax - 1) & 0xF) << 4) | ((ay - 1) & 0xF)] };
    if (ax < 769 && ay < 769) return { flag: on + 84 + 12 * (((ax - 1) & 0x300) >> 8) + (((ay - 1) & 0x300) >> 6) + xys, bytes: [(ax - 1) & 0xFF, (ay - 1) & 0xFF] };
    if (ax < 4096 && ay < 4096) return { flag: on + 120 + xys, bytes: [ax >> 4, ((ax & 0xF) << 4) | (ay >> 8), ay & 0xFF] };
    return { flag: on + 124 + xys, bytes: [ax >> 8, ax & 0xFF, ay >> 8, ay & 0xFF] };
}

/** The transformed `glyf` of a font (section 5.1), and each glyph's left edge. */
function transformGlyf(tables: Map<string, Uint8Array>): { data: Uint8Array; xMins: number[] } {
    const head = new DataView(tables.get('head')!.buffer, tables.get('head')!.byteOffset);
    const maxp = new DataView(tables.get('maxp')!.buffer, tables.get('maxp')!.byteOffset);
    const glyf = tables.get('glyf')!, locaBytes = tables.get('loca')!;
    const loca = new DataView(locaBytes.buffer, locaBytes.byteOffset, locaBytes.byteLength);
    const glyfView = new DataView(glyf.buffer, glyf.byteOffset, glyf.byteLength);
    const indexFormat = head.getInt16(50), numGlyphs = maxp.getUint16(4);
    const offset = (i: number) => indexFormat ? loca.getUint32(i * 4) : loca.getUint16(i * 2) * 2;
    const nContours = new Out(), nPoints = new Out(), flags = new Out(), glyphs = new Out(), composites = new Out(), bboxes = new Out(), instructions = new Out();
    const bboxBitmap = new Uint8Array(4 * ((numGlyphs + 31) >> 5));
    const xMins: number[] = [];
    for (let g = 0; g < numGlyphs; g++) {
        const start = offset(g), end = offset(g + 1);
        if (start === end) { nContours.u16(0); xMins.push(0); continue; }
        let at = start;
        const i16 = () => { const v = glyfView.getInt16(at); at += 2; return v; };
        const u16 = () => { const v = glyfView.getUint16(at); at += 2; return v; };
        const contours = i16();
        const box = [i16(), i16(), i16(), i16()];
        xMins.push(box[0]);
        nContours.u16(contours & 0xFFFF);
        if (contours < 0) {
            const from = at;
            let flag: number, instructed = false;
            do {
                flag = u16();
                at += 2 + (flag & 1 ? 4 : 2) + (flag & 8 ? 2 : flag & 0x40 ? 4 : flag & 0x80 ? 8 : 0);
                if (flag & 0x100) instructed = true;
            } while (flag & 0x20);
            composites.put(glyf.subarray(from, at));
            if (instructed) { const n = u16(); glyphs.u255(n); instructions.put(glyf.subarray(at, at + n)); }
            bboxBitmap[g >> 3] |= 0x80 >> (g & 7);
            for (const v of box) bboxes.u16(v);
            continue;
        }
        const ends: number[] = [];
        for (let c = 0; c < contours; c++) ends.push(u16());
        const instructionLength = u16();
        const program = glyf.subarray(at, at + instructionLength);
        at += instructionLength;
        const points = contours ? ends[contours - 1] + 1 : 0;
        const pointFlags: number[] = [];
        while (pointFlags.length < points) {
            const f = glyf[at++];
            pointFlags.push(f);
            if (f & 8) for (let r = glyf[at++]; r > 0; r--) pointFlags.push(f);
        }
        const xs: number[] = [], ys: number[] = [];
        let x = 0, y = 0;
        for (const f of pointFlags) { x += f & 2 ? (f & 0x10 ? glyf[at++] : -glyf[at++]) : f & 0x10 ? 0 : i16(); xs.push(x); }
        for (const f of pointFlags) { y += f & 4 ? (f & 0x20 ? glyf[at++] : -glyf[at++]) : f & 0x20 ? 0 : i16(); ys.push(y); }
        let last = -1;
        for (const e of ends) { nPoints.u255(e - last); last = e; }
        let px = 0, py = 0;
        pointFlags.forEach((f, p) => {
            const t = triplet(xs[p] - px, ys[p] - py, !!(f & 1));
            px = xs[p]; py = ys[p];
            flags.u8(t.flag);
            glyphs.put(t.bytes);
        });
        glyphs.u255(instructionLength);
        instructions.put(program);
        const computed = points ? [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)] : [0, 0, 0, 0];
        if (computed.some((v, i) => v !== box[i])) {
            bboxBitmap[g >> 3] |= 0x80 >> (g & 7);
            for (const v of box) bboxes.u16(v);
        }
    }
    const streams = [nContours, nPoints, flags, glyphs, composites, new Out().put(bboxBitmap).put(bboxes.bytes), instructions].map(s => s.done());
    const out = new Out().u16(0).u16(0).u16(numGlyphs).u16(indexFormat);
    for (const s of streams) out.u32(s.length);
    for (const s of streams) out.put(s);
    return { data: out.done(), xMins };
}

/**
 * A WOFF2 file of a font. Its outlines are transformed unless `transformGlyf` is false (then written as
 * they are, transform 3), and its `hmtx` too when `transformHmtx` is set and every left side bearing is
 * its glyph's left edge. Gives which tables it transformed.
 */
export function toWoff2(font: Uint8Array, options: { transformGlyf?: boolean; transformHmtx?: boolean } = {}): { bytes: Uint8Array; transformed: string[] } {
    const { flavor, tables } = readSfnt(font);
    const tags = [...tables.keys()].sort();
    const transformed: string[] = [];
    const data = new Map<string, Uint8Array>();
    const glyfTransform = tables.has('glyf') && options.transformGlyf !== false ? transformGlyf(tables) : undefined;
    if (glyfTransform) {
        data.set('glyf', glyfTransform.data);
        data.set('loca', new Uint8Array(0));
        transformed.push('glyf', 'loca');
        if (options.transformHmtx) {
            const hmtx = tables.get('hmtx')!;
            const view = new DataView(hmtx.buffer, hmtx.byteOffset, hmtx.byteLength);
            const hhea = tables.get('hhea')!;
            const metrics = (hhea[34] << 8) | hhea[35];
            const lsb = (g: number) => g < metrics ? view.getInt16(g * 4 + 2) : view.getInt16(metrics * 4 + (g - metrics) * 2);
            if (glyfTransform.xMins.every((x, g) => lsb(g) === x)) {
                const out = new Out().u8(metrics < glyfTransform.xMins.length ? 3 : 1);
                for (let g = 0; g < metrics; g++) out.u16(view.getUint16(g * 4));
                data.set('hmtx', out.done());
                transformed.push('hmtx');
            }
        }
    }
    const directory = new Out();
    const stream = new Out();
    for (const tag of tags) {
        const known = KNOWN_TAGS.indexOf(tag);
        const isTransformed = transformed.includes(tag);
        // glyf and loca are transformed at version 0 and written as they are at 3; hmtx is transformed at 1.
        const version = tag === 'glyf' || tag === 'loca' ? (isTransformed ? 0 : 3) : isTransformed ? 1 : 0;
        directory.u8((known >= 0 ? known : 63) | (version << 6));
        if (known < 0) directory.put([...tag].map(c => c.charCodeAt(0)));
        directory.base128(tables.get(tag)!.length);
        if (isTransformed) directory.base128(data.get(tag)!.length);
        stream.put(data.get(tag) ?? tables.get(tag)!);
    }
    const compressed = new Uint8Array(zlib.brotliCompressSync(stream.done(), { params: { [zlib.constants.BROTLI_PARAM_QUALITY]: 11 } }));
    const sfntSize = 12 + 16 * tags.length + tags.reduce((n, tag) => n + ((tables.get(tag)!.length + 3) & ~3), 0);
    const length = (48 + directory.bytes.length + compressed.length + 3) & ~3;
    const out = new Out().u32(0x774F4632).u32(flavor).u32(length).u16(tags.length).u16(0).u32(sfntSize).u32(compressed.length)
        .u16(1).u16(0).u32(0).u32(0).u32(0).u32(0).u32(0);
    out.put(directory.bytes).put(compressed).pad();
    return { bytes: out.done(), transformed };
}

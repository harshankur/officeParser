/**
 * The font files given to the native PDF engine, read before pdf-lib and fontkit see them.
 *
 * Fonts in the web's formats, WOFF and WOFF2, are unpacked to the TrueType or OpenType file (an sfnt)
 * each holds, which is what a PDF embeds and what fontkit subsets. A WOFF file is the sfnt's tables each
 * compressed with zlib; a WOFF2 file is them compressed together with Brotli, its glyph outlines (`glyf`
 * and `loca`) and, optionally, its horizontal metrics (`hmtx`) written in a smaller form first (the W3C
 * WOFF2 specification, section 5), which is rebuilt here. fontkit reads both formats, but neither its
 * subset of a WOFF2 font nor a WOFF or WOFF2 file embedded whole is a font viewers draw.
 *
 * The TrueType outlines of the glyphs drawn are checked before they are embedded, and a character whose
 * glyph draws nothing (empty, or a color glyph's stand-in) is not counted as one the font draws, unless it
 * is blank by design, or a mark a font that substitutes glyphs draws with its letter (see TrueTypeOutlines,
 * and the engine's BLANK).
 *
 * @module generators/pdf/fontFiles
 */

import { Unzlib } from 'fflate';

/**
 * What is wrong with a font that cannot be unpacked or drawn. A class of its own (rather than
 * `new Error`), as `DecryptionError` is, keeps this module free of the generator's error machinery: the
 * engine reports it as the font's `PDF_GENERATION_FAILED`.
 */
export class UnreadableFont extends Error {
    constructor(message: string) {
        super(message);
        this.name = 'UnreadableFont';
    }
}

/** Brotli-decompresses `data`, which holds exactly `size` bytes. Throws when it does not. */
export type Unbrotli = (data: Uint8Array, size: number) => Uint8Array;

/**
 * The most bytes a font's tables come to once unpacked. A real font is far below this (Arial Unicode,
 * of some 50,000 glyphs, is 23 MB); a file declaring more is refused rather than allocated.
 */
export const MAX_UNPACKED_FONT = 64 * 1024 * 1024;

/** The tags a WOFF2 table directory names by number (the specification's table 2), in that order. */
const KNOWN_TAGS = [
    'cmap', 'head', 'hhea', 'hmtx', 'maxp', 'name', 'OS/2', 'post', 'cvt ', 'fpgm', 'glyf', 'loca', 'prep',
    'CFF ', 'VORG', 'EBDT', 'EBLC', 'gasp', 'hdmx', 'kern', 'LTSH', 'PCLT', 'VDMX', 'vhea', 'vmtx', 'BASE',
    'GDEF', 'GPOS', 'GSUB', 'EBSC', 'JSTF', 'MATH', 'CBDT', 'CBLC', 'COLR', 'CPAL', 'SVG ', 'sbix', 'acnt',
    'avar', 'bdat', 'bloc', 'bsln', 'cvar', 'fdsc', 'feat', 'fmtx', 'fvar', 'gvar', 'hsty', 'just', 'lcar',
    'mort', 'morx', 'opbd', 'prop', 'trak', 'Zapf', 'Silf', 'Glat', 'Gloc', 'Feat', 'Sill',
];

/** The web font format of a file, by its signature, or undefined for any other file. */
export function webFontFormat(bytes: Uint8Array): 'woff' | 'woff2' | undefined {
    if (bytes.length < 4 || bytes[0] !== 0x77 || bytes[1] !== 0x4F || bytes[2] !== 0x46) return undefined;
    return bytes[3] === 0x46 ? 'woff' : bytes[3] === 0x32 ? 'woff2' : undefined;
}

/**
 * The sfnt of a font file: a TrueType or OpenType file as it is, the one a WOFF or WOFF2 file holds
 * unpacked. Throws, saying what is wrong, for a collection of fonts (`.ttc`, which a PDF does not embed
 * whole and pdf-lib reads no font of), and for a web font that cannot be unpacked: damaged, cut short,
 * or larger than {@link MAX_UNPACKED_FONT}.
 */
export function sfntOf(bytes: Uint8Array, unbrotli: () => Unbrotli): Uint8Array {
    const format = webFontFormat(bytes);
    if (format === 'woff') return unpackWoff(bytes);
    if (format === 'woff2') return unpackWoff2(bytes, unbrotli());
    if (bytes.length >= 4 && bytes[0] === 0x74 && bytes[1] === 0x74 && bytes[2] === 0x63 && bytes[3] === 0x66) throw new UnreadableFont(COLLECTION);
    return bytes;
}

/** Why a collection of fonts is not read. */
const COLLECTION = 'it is a collection of fonts (ttcf), not one font: give one of its fonts';

/** Reads big-endian numbers from a part of a file, throwing when a read runs past its end. */
class Reader {
    pos: number;
    constructor(readonly bytes: Uint8Array, start = 0, readonly end = bytes.length, readonly what = 'the file') { this.pos = start; }
    private need(n: number): number {
        const at = this.pos;
        if (n < 0 || at + n > this.end) throw new UnreadableFont(`${this.what} ends before its data does`);
        this.pos += n;
        return at;
    }
    u8(): number { return this.bytes[this.need(1)]; }
    u16(): number { const at = this.need(2); return (this.bytes[at] << 8) | this.bytes[at + 1]; }
    i16(): number { const v = this.u16(); return v & 0x8000 ? v - 0x10000 : v; }
    u32(): number { const at = this.need(4); return ((this.bytes[at] << 24) | (this.bytes[at + 1] << 16) | (this.bytes[at + 2] << 8) | this.bytes[at + 3]) >>> 0; }
    take(n: number): Uint8Array { const at = this.need(n); return this.bytes.subarray(at, at + n); }
    tag(): string { return String.fromCharCode(...this.take(4)); }
    /** A WOFF2 `UIntBase128`: up to five bytes of seven bits each, most significant first. */
    base128(): number {
        let value = 0;
        for (let i = 0; i < 5; i++) {
            const byte = this.u8();
            if (i === 0 && byte === 0x80) throw new UnreadableFont('a WOFF2 number starts with a zero');
            if (value > 0x1FFFFFF) throw new UnreadableFont('a WOFF2 number is larger than 32 bits');
            value = value * 128 + (byte & 0x7F);
            if (!(byte & 0x80)) return value;
        }
        throw new UnreadableFont('a WOFF2 number runs past five bytes');
    }
    /** A WOFF2 `255UInt16`: one byte, or a marker byte and one or two more. */
    u255(): number {
        const code = this.u8();
        if (code === 253) return this.u16();
        if (code === 255) return this.u8() + 253;
        if (code === 254) return this.u8() + 506;
        return code;
    }
}

/** Writes big-endian numbers into a buffer that grows as it is written. */
class Writer {
    bytes = new Uint8Array(256);
    length = 0;
    private room(n: number): number {
        if (this.length + n > this.bytes.length) {
            const grown = new Uint8Array(Math.max(this.bytes.length * 2, this.length + n));
            grown.set(this.bytes.subarray(0, this.length));
            this.bytes = grown;
        }
        const at = this.length;
        this.length += n;
        return at;
    }
    // (Each finds its room before it reads `bytes`, which finding room may replace.)
    u8(v: number): void { const at = this.room(1); this.bytes[at] = v; }
    u16(v: number): void { const at = this.room(2); this.bytes[at] = v >> 8; this.bytes[at + 1] = v; }
    u32(v: number): void { const at = this.room(4); this.bytes[at] = v >>> 24; this.bytes[at + 1] = v >>> 16; this.bytes[at + 2] = v >>> 8; this.bytes[at + 3] = v; }
    put(data: Uint8Array): void { const at = this.room(data.length); this.bytes.set(data, at); }
    /** Zeros up to the next multiple of four bytes. */
    pad(): void { while (this.length % 4) this.u8(0); }
    done(): Uint8Array { return this.bytes.subarray(0, this.length); }
}

/** The sum of a table's bytes as big-endian 32-bit numbers, the last one padded with zeros (an sfnt's checksum). */
function checksum(data: Uint8Array): number {
    let sum = 0;
    const whole = data.length & ~3;
    for (let i = 0; i < whole; i += 4) sum = (sum + ((data[i] << 24) | (data[i + 1] << 16) | (data[i + 2] << 8) | data[i + 3])) >>> 0;
    if (whole < data.length) {
        let last = 0;
        for (let i = 0; i < 4; i++) last = (last << 8) | (whole + i < data.length ? data[whole + i] : 0);
        sum = (sum + (last >>> 0)) >>> 0;
    }
    return sum;
}

/**
 * An sfnt of these tables: its header, a record for each table (in the order of their tags), and the
 * tables, each starting on a four-byte boundary, with their checksums and the font's
 * (`head.checkSumAdjustment`) written anew.
 */
function writeSfnt(flavor: number, tables: Map<string, Uint8Array>): Uint8Array {
    const tags = [...tables.keys()].sort((a, b) => a < b ? -1 : a > b ? 1 : 0);
    const n = tags.length;
    let power = 1, log = 0;
    while (power * 2 <= n) { power *= 2; log++; }
    const out = new Writer();
    out.u32(flavor);
    out.u16(n);
    out.u16(power * 16);
    out.u16(log);
    out.u16(n * 16 - power * 16);
    let offset = 12 + n * 16;
    const head = tables.get('head');
    // (The font's checksum is of the font with this field zero.)
    if (head && head.length >= 12) head.fill(0, 8, 12);
    for (const tag of tags) {
        const data = tables.get(tag)!;
        for (let i = 0; i < 4; i++) out.u8(tag.charCodeAt(i));
        out.u32(checksum(data));
        out.u32(offset);
        out.u32(data.length);
        offset += (data.length + 3) & ~3;
    }
    let headAt = -1;
    for (const tag of tags) {
        if (tag === 'head') headAt = out.length;
        out.put(tables.get(tag)!);
        out.pad();
    }
    const font = out.done();
    if (headAt >= 0 && head && head.length >= 12) {
        const adjustment = (0xB1B0AFBA - checksum(font)) >>> 0;
        font[headAt + 8] = adjustment >>> 24; font[headAt + 9] = adjustment >>> 16; font[headAt + 10] = adjustment >>> 8; font[headAt + 11] = adjustment;
    }
    return font;
}

/** The sfnt of a WOFF file: each table inflated (zlib) where it is stored compressed. */
function unpackWoff(bytes: Uint8Array): Uint8Array {
    const header = new Reader(bytes, 4, bytes.length, 'the WOFF file');
    const flavor = header.u32();
    header.u32(); // length
    const numTables = header.u16();
    header.pos = 44;
    if (flavor === 0x74746366) throw new UnreadableFont(COLLECTION);
    const tables = new Map<string, Uint8Array>();
    let total = 0;
    for (let i = 0; i < numTables; i++) {
        const tag = header.tag();
        const offset = header.u32(), compLength = header.u32(), origLength = header.u32();
        header.u32(); // origChecksum
        if (tables.has(tag)) throw new UnreadableFont(`it has two tables '${tag}'`);
        total += origLength;
        if (total > MAX_UNPACKED_FONT) throw new UnreadableFont(`its tables come to more than ${MAX_UNPACKED_FONT / 1048576} MB`);
        const stored = new Reader(bytes, offset, bytes.length, `the WOFF table '${tag}'`).take(compLength);
        let data: Uint8Array;
        if (compLength === origLength) data = stored.slice();
        else if (compLength < origLength) data = inflateTable(stored, origLength, tag);
        else throw new UnreadableFont(`the WOFF table '${tag}' is stored larger than it is`);
        tables.set(tag, data);
    }
    return writeSfnt(flavor, tables);
}

/**
 * How much of a WOFF table's zlib data is inflated at a time. Data inflating past the length its table
 * declares is refused at the step it does, after at most some 1,032 times this (deflate's largest ratio),
 * 16 MB: inflated whole, a few kilobytes inflated to gigabytes before they were found too long.
 */
const INFLATE_STEP = 16 * 1024;

/** A WOFF table's zlib data inflated, which must come to exactly `origLength` bytes. */
function inflateTable(stored: Uint8Array, origLength: number, tag: string): Uint8Array {
    const out = new Uint8Array(origLength);
    let length = 0;
    const inflater = new Unzlib((chunk: Uint8Array) => {
        if (length + chunk.length > origLength) throw new UnreadableFont(`the WOFF table '${tag}' inflates to more than the ${origLength} bytes it declares`);
        out.set(chunk, length);
        length += chunk.length;
    });
    try {
        for (let at = 0; at < stored.length; at += INFLATE_STEP) inflater.push(stored.subarray(at, at + INFLATE_STEP), at + INFLATE_STEP >= stored.length);
    } catch (e: any) {
        if (e instanceof UnreadableFont) throw e;
        throw new UnreadableFont(`the WOFF table '${tag}' does not inflate (${e?.message || String(e)})`);
    }
    if (length !== origLength) throw new UnreadableFont(`the WOFF table '${tag}' inflates to ${length} bytes, not the ${origLength} it declares`);
    return out;
}

/** One table of a WOFF2 file, as its directory gives it. */
interface Woff2Table { tag: string; version: number; origLength: number; transformLength?: number; data?: Uint8Array }

/** The sfnt of a WOFF2 file: its tables Brotli-decompressed, then its outlines and metrics rebuilt. */
function unpackWoff2(bytes: Uint8Array, unbrotli: Unbrotli): Uint8Array {
    const header = new Reader(bytes, 4, bytes.length, 'the WOFF2 file');
    const flavor = header.u32();
    header.u32(); // length
    const numTables = header.u16();
    header.u16(); // reserved
    header.u32(); // totalSfntSize
    const compressedSize = header.u32();
    header.pos = 48;
    if (flavor === 0x74746366) throw new UnreadableFont(COLLECTION);
    if (!numTables) throw new UnreadableFont('it has no tables');
    const directory: Woff2Table[] = [];
    const tags = new Set<string>();
    let total = 0;
    for (let i = 0; i < numTables; i++) {
        const flags = header.u8();
        const tag = (flags & 0x3F) === 0x3F ? header.tag() : KNOWN_TAGS[flags & 0x3F];
        const version = flags >> 6;
        const origLength = header.base128();
        // glyf and loca are transformed unless their version is 3 (none); any other table is unless it is 0.
        const outlines = tag === 'glyf' || tag === 'loca';
        if (outlines ? version === 1 || version === 2 : version !== 0 && !(tag === 'hmtx' && version === 1)) throw new UnreadableFont(`its table '${tag}' is in an unknown form (transform ${version})`);
        const transformed = outlines ? version === 0 : version !== 0;
        const transformLength = transformed ? header.base128() : undefined;
        if (tags.has(tag)) throw new UnreadableFont(`it has two tables '${tag}'`);
        tags.add(tag);
        directory.push({ tag, version, origLength, transformLength });
        total += transformLength ?? origLength;
        if (total > MAX_UNPACKED_FONT) throw new UnreadableFont(`its tables come to more than ${MAX_UNPACKED_FONT / 1048576} MB`);
    }
    let data: Uint8Array;
    try { data = unbrotli(header.take(compressedSize), total); } catch (e: any) { throw new UnreadableFont(`its tables do not decompress (${e?.message || String(e)})`); }
    if (data.length !== total) throw new UnreadableFont(`its tables decompress to ${data.length} bytes, not the ${total} its directory declares`);
    let at = 0;
    for (const table of directory) {
        const length = table.transformLength ?? table.origLength;
        table.data = data.subarray(at, at + length);
        at += length;
    }
    const byTag = new Map(directory.map(t => [t.tag, t]));
    const tables = new Map<string, Uint8Array>();
    const glyf = byTag.get('glyf'), loca = byTag.get('loca');
    if ((glyf?.transformLength !== undefined) !== (loca?.transformLength !== undefined)) throw new UnreadableFont('its glyf and loca tables are not transformed together');
    let xMins: Int16Array | undefined;
    if (glyf?.transformLength !== undefined && loca) {
        const rebuilt = rebuildGlyf(glyf.data!);
        tables.set('glyf', rebuilt.glyf);
        tables.set('loca', rebuilt.loca);
        xMins = rebuilt.xMins;
        const head = byTag.get('head')?.data;
        if (!head || head.length < 54) throw new UnreadableFont('it has no head table');
        // The offsets are written in the size the outlines' header gives; the head table says which.
        tables.set('head', head.slice());
        tables.get('head')![51] = rebuilt.longOffsets ? 1 : 0;
        tables.get('head')![50] = 0;
    }
    for (const table of directory) {
        if (tables.has(table.tag)) continue;
        if (table.tag === 'hmtx' && table.transformLength !== undefined) {
            const hhea = byTag.get('hhea')?.data;
            if (!xMins || !hhea || hhea.length < 36) throw new UnreadableFont('its hmtx table is transformed without the outlines or hhea it is rebuilt from');
            tables.set('hmtx', rebuildHmtx(table.data!, xMins, (hhea[34] << 8) | hhea[35]));
        } else tables.set(table.tag, table.data!.slice());
    }
    return writeSfnt(flavor, tables);
}

/** The bit of a glyph in a bitmap of one bit a glyph, the first glyph the top bit of the first byte. */
function bitOf(bitmap: Uint8Array, glyph: number): boolean {
    return (bitmap[glyph >> 3] & (0x80 >> (glyph & 7))) !== 0;
}

/**
 * The movement to a glyph's next point, from its flag (a byte of the flag stream) and one to four bytes
 * of the glyph stream: the specification's table of 128 triplet encodings (section 5.2), as the
 * reference decoder computes it. The top bit of the flag says the point is off the curve.
 */
function readTriplet(flag: number, glyphs: Reader): [number, number] {
    const f = flag & 0x7F;
    const signed = (bit: number, value: number) => bit & 1 ? value : -value;
    if (f < 10) return [0, signed(f, ((f & 14) << 7) + glyphs.u8())];
    if (f < 20) return [signed(f, (((f - 10) & 14) << 7) + glyphs.u8()), 0];
    if (f < 84) {
        const b0 = f - 20, b1 = glyphs.u8();
        return [signed(f, 1 + (b0 & 0x30) + (b1 >> 4)), signed(f >> 1, 1 + ((b0 & 0x0C) << 2) + (b1 & 0x0F))];
    }
    if (f < 120) {
        const b0 = f - 84, b1 = glyphs.u8(), b2 = glyphs.u8();
        return [signed(f, 1 + (Math.floor(b0 / 12) << 8) + b1), signed(f >> 1, 1 + (((b0 % 12) >> 2) << 8) + b2)];
    }
    if (f < 124) {
        const b1 = glyphs.u8(), b2 = glyphs.u8(), b3 = glyphs.u8();
        return [signed(f, (b1 << 4) + (b2 >> 4)), signed(f >> 1, ((b2 & 0x0F) << 8) + b3)];
    }
    const b1 = glyphs.u8(), b2 = glyphs.u8(), b3 = glyphs.u8(), b4 = glyphs.u8();
    return [signed(f, (b1 << 8) + b2), signed(f >> 1, (b3 << 8) + b4)];
}

// The flags of a TrueType simple glyph's points, and of a composite glyph's components.
const ON_CURVE = 0x01, X_SHORT = 0x02, Y_SHORT = 0x04, REPEAT = 0x08, X_SAME = 0x10, Y_SAME = 0x20, OVERLAP_SIMPLE = 0x40;
const ARGS_ARE_WORDS = 0x0001, HAVE_SCALE = 0x0008, MORE_COMPONENTS = 0x0020, HAVE_X_Y_SCALE = 0x0040, HAVE_TWO_BY_TWO = 0x0080, HAVE_INSTRUCTIONS = 0x0100;

/**
 * The `glyf` and `loca` tables of a WOFF2 file's transformed outlines (section 5.1): each glyph's
 * contours, points, flags, coordinates, components, bounding box and instructions read from the
 * streams they were split into, and written as a TrueType glyph, each starting on a four-byte boundary.
 * Also gives each glyph's left edge (`xMin`), from which a transformed `hmtx` is rebuilt.
 */
function rebuildGlyf(table: Uint8Array): { glyf: Uint8Array; loca: Uint8Array; xMins: Int16Array; longOffsets: boolean } {
    const head = new Reader(table, 0, table.length, "the WOFF2 table 'glyf'");
    head.u16(); // reserved
    const options = head.u16();
    const numGlyphs = head.u16();
    const indexFormat = head.u16();
    const sizes = [head.u32(), head.u32(), head.u32(), head.u32(), head.u32(), head.u32(), head.u32()];
    const streams: Reader[] = [];
    const names = ['nContour', 'nPoints', 'flag', 'glyph', 'composite', 'bbox', 'instruction'];
    let at = head.pos;
    for (let i = 0; i < sizes.length; i++) {
        if (at + sizes[i] > table.length) throw new UnreadableFont(`the WOFF2 outlines' ${names[i]} stream runs past their table`);
        streams.push(new Reader(table, at, at + sizes[i], `the WOFF2 outlines' ${names[i]} stream`));
        at += sizes[i];
    }
    const [nContours, nPoints, flags, glyphs, composites, bboxes, instructions] = streams;
    const bboxBitmap = bboxes.take(4 * ((numGlyphs + 31) >> 5));
    const overlapBitmap = options & 1 ? new Reader(table, at, table.length, "the WOFF2 outlines' overlap bitmap").take((numGlyphs + 7) >> 3) : undefined;

    const glyf = new Writer();
    const offsets = new Uint32Array(numGlyphs + 1);
    const xMins = new Int16Array(numGlyphs);
    for (let g = 0; g < numGlyphs; g++) {
        // (Outlines rebuilt come to half as much again as their transformed bytes: the cap holds for them too.)
        if (glyf.length > MAX_UNPACKED_FONT) throw new UnreadableFont(`its outlines come to more than ${MAX_UNPACKED_FONT / 1048576} MB`);
        offsets[g] = glyf.length;
        const contours = nContours.i16();
        const hasBbox = bitOf(bboxBitmap, g);
        if (contours === 0) {
            if (hasBbox) throw new UnreadableFont(`the WOFF2 outlines give the empty glyph ${g} a bounding box`);
            continue;
        }
        if (contours === -1) {
            if (!hasBbox) throw new UnreadableFont(`the WOFF2 outlines give the composite glyph ${g} no bounding box`);
            const box = [bboxes.i16(), bboxes.i16(), bboxes.i16(), bboxes.i16()];
            const start = composites.pos;
            let flag: number, instructed = false;
            do {
                flag = composites.u16();
                composites.u16(); // glyph index
                composites.take((flag & ARGS_ARE_WORDS ? 4 : 2) + (flag & HAVE_SCALE ? 2 : flag & HAVE_X_Y_SCALE ? 4 : flag & HAVE_TWO_BY_TWO ? 8 : 0));
                if (flag & HAVE_INSTRUCTIONS) instructed = true;
            } while (flag & MORE_COMPONENTS);
            glyf.u16(0xFFFF);
            for (const v of box) glyf.u16(v & 0xFFFF);
            glyf.put(table.subarray(start, composites.pos));
            if (instructed) {
                const length = glyphs.u255();
                glyf.u16(length);
                glyf.put(instructions.take(length));
            }
            xMins[g] = box[0];
            glyf.pad();
            continue;
        }
        if (contours < 0) throw new UnreadableFont(`the WOFF2 outlines give glyph ${g} ${contours} contours`);
        const ends: number[] = [];
        let points = 0;
        for (let c = 0; c < contours; c++) {
            points += nPoints.u255();
            ends.push(points - 1);
        }
        if (points > 0xFFFF) throw new UnreadableFont(`the WOFF2 outlines give glyph ${g} more points than a glyph holds`);
        const xs = new Int32Array(points), ys = new Int32Array(points), on = new Uint8Array(points);
        let x = 0, y = 0;
        for (let p = 0; p < points; p++) {
            const flag = flags.u8();
            const [dx, dy] = readTriplet(flag, glyphs);
            x += dx; y += dy;
            xs[p] = x; ys[p] = y; on[p] = flag & 0x80 ? 0 : 1;
        }
        const instructionLength = glyphs.u255();
        let box: number[];
        if (hasBbox) box = [bboxes.i16(), bboxes.i16(), bboxes.i16(), bboxes.i16()];
        else {
            box = points ? [xs[0], ys[0], xs[0], ys[0]] : [0, 0, 0, 0];
            for (let p = 1; p < points; p++) {
                if (xs[p] < box[0]) box[0] = xs[p];
                if (ys[p] < box[1]) box[1] = ys[p];
                if (xs[p] > box[2]) box[2] = xs[p];
                if (ys[p] > box[3]) box[3] = ys[p];
            }
        }
        glyf.u16(contours);
        for (const v of box) glyf.u16(v & 0xFFFF);
        for (const end of ends) glyf.u16(end);
        glyf.u16(instructionLength);
        glyf.put(instructions.take(instructionLength));
        // The points' flags (a run of one flag written once, with its count), then their x and y
        // movements, each in a byte where it fits one (its sign in the flag), or left out where it is 0.
        const xBytes = new Writer(), yBytes = new Writer();
        let last = -1, repeats = 0, lastX = 0, lastY = 0;
        for (let p = 0; p < points; p++) {
            let flag = on[p] ? ON_CURVE : 0;
            if (p === 0 && overlapBitmap && bitOf(overlapBitmap, g)) flag |= OVERLAP_SIMPLE;
            const dx = xs[p] - lastX, dy = ys[p] - lastY;
            lastX = xs[p]; lastY = ys[p];
            if (dx === 0) flag |= X_SAME;
            else if (dx > -256 && dx < 256) { flag |= X_SHORT | (dx > 0 ? X_SAME : 0); xBytes.u8(Math.abs(dx)); } else xBytes.u16(dx & 0xFFFF);
            if (dy === 0) flag |= Y_SAME;
            else if (dy > -256 && dy < 256) { flag |= Y_SHORT | (dy > 0 ? Y_SAME : 0); yBytes.u8(Math.abs(dy)); } else yBytes.u16(dy & 0xFFFF);
            if (flag === last && repeats !== 255) {
                glyf.bytes[glyf.length - 1] |= REPEAT;
                repeats++;
            } else {
                if (repeats) glyf.u8(repeats);
                glyf.u8(flag);
                repeats = 0;
            }
            last = flag;
        }
        if (repeats) glyf.u8(repeats);
        glyf.put(xBytes.done());
        glyf.put(yBytes.done());
        xMins[g] = box[0];
        glyf.pad();
    }
    offsets[numGlyphs] = glyf.length;
    // Short offsets are halves, and hold up to 128 kB of outlines; past that the long form is written.
    const longOffsets = indexFormat !== 0 || glyf.length > 0x1FFFE;
    const loca = new Writer();
    for (const offset of offsets) longOffsets ? loca.u32(offset) : loca.u16(offset / 2);
    return { glyf: glyf.done(), loca: loca.done(), xMins, longOffsets };
}

/**
 * The `hmtx` table of a WOFF2 file's transformed metrics (section 5.4): the advance widths as they are,
 * and the left side bearings left out of it taken from the glyphs' left edges, as they are equal to.
 */
function rebuildHmtx(table: Uint8Array, xMins: Int16Array, numberOfHMetrics: number): Uint8Array {
    const data = new Reader(table, 0, table.length, "the WOFF2 table 'hmtx'");
    const flags = data.u8();
    const numGlyphs = xMins.length;
    if (numberOfHMetrics < 1 || numberOfHMetrics > numGlyphs) throw new UnreadableFont(`its hhea table gives ${numberOfHMetrics} metrics for ${numGlyphs} glyphs`);
    const advances: number[] = [];
    for (let i = 0; i < numberOfHMetrics; i++) advances.push(data.u16());
    const bearings: number[] = [];
    for (let i = 0; i < numberOfHMetrics; i++) bearings.push(flags & 1 ? xMins[i] : data.i16());
    for (let i = numberOfHMetrics; i < numGlyphs; i++) bearings.push(flags & 2 ? xMins[i] : data.i16());
    const out = new Writer();
    for (let i = 0; i < numGlyphs; i++) {
        if (i < numberOfHMetrics) out.u16(advances[i]);
        out.u16(bearings[i] & 0xFFFF);
    }
    return out.done();
}

/**
 * The Brotli decoder fontkit carries (it reads WOFF2 with it), reached through what both its builds,
 * `@pdf-lib/fontkit` and `@cantoo/fontkit`, do: the compressed data is wrapped in a WOFF2 file of one
 * table, as long as the data decompresses to, and fontkit decompresses it to read that table. The
 * table is `maxp`, and the bytes are those fontkit decompressed into (its font's `stream`, which both
 * builds replace with them), taken as they are; where a build does not, the table is the font
 * program (`fpgm`) and the bytes its instructions, which cost some fifteen times their length (fontkit
 * reads them into an array of numbers): 64 MB took a gigabyte.
 */
export function fontkitUnbrotli(fontkit: any): Unbrotli {
    return (data, size) => {
        const maxp = fontkit.create(woff2OfOneTable(4, data, size));
        const before = maxp.stream;
        try { void maxp.maxp; } catch { /* (the bytes are no maxp table: they are decompressed all the same) */ }
        const out = maxp.stream !== before ? maxp.stream?.buffer : undefined;
        if (out instanceof Uint8Array) {
            // (Data decompressing to more than its table's length is no font's: fontkit's decoder does not stop there.)
            if (out.length !== size) throw new UnreadableFont(`it decompresses to ${out.length} bytes, not the ${size} it declares`);
            return out;
        }
        const program = fontkit.create(woff2OfOneTable(9, data, size)).fpgm?.instructions;
        if (!program || program.length !== size) throw new UnreadableFont('fontkit did not decompress it');
        return Uint8Array.from(program);
    };
}

/** A WOFF2 file of one table, given by its tag's number (4, maxp; 9, fpgm), not transformed, `size` bytes long, its data Brotli's `data`. */
function woff2OfOneTable(tagNumber: number, data: Uint8Array, size: number): Uint8Array {
    // The table's entry: its tag by number, then its length (UIntBase128).
    const entry: number[] = [];
    let rest = size;
    do { entry.unshift((rest & 0x7F) | (entry.length ? 0x80 : 0)); rest = Math.floor(rest / 128); } while (rest);
    entry.unshift(tagNumber);
    const file = new Writer();
    file.u32(0x774F4632); // wOF2
    file.u32(0x00010000);
    file.u32(48 + entry.length + data.length);
    file.u16(1); // numTables
    file.u16(0);
    file.u32(12 + 16 + ((size + 3) & ~3)); // totalSfntSize
    file.u32(data.length); // totalCompressedSize
    // (Its version, then its metadata's and private data's offsets and lengths: none.)
    file.u16(0);
    file.u16(0);
    for (let i = 0; i < 5; i++) file.u32(0);
    for (const byte of entry) file.u8(byte);
    file.put(data);
    return file.done();
}

/**
 * Node's Brotli decoder (`zlib`), where the runtime gives its built-in modules without an import (Node
 * 20.16 and 22.3 on), so that no build for the browser carries it; undefined elsewhere.
 */
export function nativeUnbrotli(): Unbrotli | undefined {
    let zlib: any;
    try { zlib = (globalThis as any).process?.getBuiltinModule?.('zlib'); } catch { return undefined; }
    if (typeof zlib?.brotliDecompressSync !== 'function') return undefined;
    return (data, size) => {
        const out: Uint8Array = zlib.brotliDecompressSync(data, { maxOutputLength: Math.max(size, 1) });
        return new Uint8Array(out.buffer, out.byteOffset, out.byteLength);
    };
}

/** A table of an sfnt, or undefined when it has none; throws for one that lies outside the file. */
function sfntTable(sfnt: Uint8Array, tag: string): Uint8Array | undefined {
    const directory = new Reader(sfnt, 4, sfnt.length, 'the font file');
    const numTables = directory.u16();
    directory.pos = 12;
    for (let i = 0; i < numTables; i++) {
        const name = directory.tag();
        directory.u32(); // checksum
        const offset = directory.u32(), length = directory.u32();
        if (name !== tag) continue;
        if (offset + length > sfnt.length) throw new UnreadableFont(`its table '${tag}' lies outside the file`);
        return sfnt.subarray(offset, offset + length);
    }
    return undefined;
}

/** How deep a composite glyph's components may nest (FreeType's own limit is far below it). */
const MAX_COMPONENT_DEPTH = 32;

/** The color tables a font may draw its glyphs from instead of their outlines, which a PDF does not hold. */
const COLOR_TABLES = ['sbix', 'CBDT', 'COLR', 'SVG '];

/** The color tables of an sfnt (see COLOR_TABLES). */
export function colorTables(sfnt: Uint8Array): string[] {
    return COLOR_TABLES.filter(tag => sfntTable(sfnt, tag));
}

/** The tables a font substitutes glyphs by as it shapes text: OpenType's, and Apple's. */
const SUBSTITUTION_TABLES = ['GSUB', 'morx', 'mort'];

/** Whether an sfnt substitutes glyphs as it shapes text (see SUBSTITUTION_TABLES); not one whose table lies outside the file. */
export function substitutes(sfnt: Uint8Array): boolean {
    try { return SUBSTITUTION_TABLES.some(tag => sfntTable(sfnt, tag)); } catch { return false; }
}

/**
 * The TrueType outlines (`glyf`) of a font, read as viewers read them, or undefined for a font of other
 * outlines (CFF).
 */
export function trueTypeOutlines(sfnt: Uint8Array): TrueTypeOutlines | undefined {
    const glyf = sfntTable(sfnt, 'glyf');
    if (!glyf) return undefined;
    const loca = sfntTable(sfnt, 'loca'), head = sfntTable(sfnt, 'head');
    if (!loca || !head || head.length < 54) throw new UnreadableFont('its outlines have no loca or head table');
    return new TrueTypeOutlines(glyf, loca, head[51] === 1, sfntTable(sfnt, 'maxp'));
}

/** A glyph's outline as it was read: the number of its contours' points, or the glyphs it is made of. */
type Outline = { contours: number[] } | { components: number[] };

export class TrueTypeOutlines {
    readonly numGlyphs: number;
    private readonly read = new Map<number, Outline>();
    private readonly drawing = new Map<number, boolean>();
    /** The glyphs checked, each with how deep its components nest below it (0 for a glyph of none). */
    private readonly checkedDepths = new Map<number, number>();
    constructor(private readonly glyf: Uint8Array, private readonly loca: Uint8Array, private readonly long: boolean, maxp: Uint8Array | undefined) {
        const entries = Math.floor(loca.length / (long ? 4 : 2));
        this.numGlyphs = Math.max(0, Math.min(maxp && maxp.length >= 6 ? (maxp[4] << 8) | maxp[5] : entries - 1, entries - 1));
    }

    private offset(i: number): number {
        const loca = this.loca;
        return this.long
            ? ((loca[i * 4] << 24) | (loca[i * 4 + 1] << 16) | (loca[i * 4 + 2] << 8) | loca[i * 4 + 3]) >>> 0
            : ((loca[i * 2] << 8) | loca[i * 2 + 1]) * 2;
    }

    /**
     * A glyph's outline: where `loca` says, inside `glyf`, its contours, instructions, flags and
     * coordinates inside its own bytes, or its components. Throws, naming the glyph, where it is not.
     */
    private outline(id: number): Outline {
        const known = this.read.get(id);
        if (known) return known;
        if (!Number.isInteger(id) || id < 0 || id >= this.numGlyphs) throw new UnreadableFont(`it has no glyph ${id} (it has ${this.numGlyphs})`);
        const start = this.offset(id), end = this.offset(id + 1);
        if (start > end || end > this.glyf.length) throw new UnreadableFont(`the outline of its glyph ${id} lies outside its glyf table`);
        let outline: Outline = { contours: [] };
        if (start < end) {
            const data = new Reader(this.glyf, start, end, `the outline of its glyph ${id}`);
            const contours = data.i16();
            data.take(8); // bounding box
            if (contours >= 0) {
                const points: number[] = [];
                let last = -1;
                for (let c = 0; c < contours; c++) {
                    const endPoint = data.u16();
                    if (endPoint <= last) throw new UnreadableFont(`the contours of its glyph ${id} end out of order`);
                    points.push(endPoint - last);
                    last = endPoint;
                }
                data.take(data.u16()); // instructions
                let left = last + 1, xBytes = 0, yBytes = 0;
                while (left > 0) {
                    const flag = data.u8();
                    const count = 1 + (flag & REPEAT ? data.u8() : 0);
                    if (count > left) throw new UnreadableFont(`the flags of its glyph ${id} run past its points`);
                    left -= count;
                    xBytes += count * (flag & X_SHORT ? 1 : flag & X_SAME ? 0 : 2);
                    yBytes += count * (flag & Y_SHORT ? 1 : flag & Y_SAME ? 0 : 2);
                }
                data.take(xBytes + yBytes);
                outline = { contours: points };
            } else if (contours === -1) {
                const components: number[] = [];
                let flag: number;
                do {
                    flag = data.u16();
                    components.push(data.u16());
                    data.take((flag & ARGS_ARE_WORDS ? 4 : 2) + (flag & HAVE_SCALE ? 2 : flag & HAVE_X_Y_SCALE ? 4 : flag & HAVE_TWO_BY_TWO ? 8 : 0));
                } while (flag & MORE_COMPONENTS);
                if (flag & HAVE_INSTRUCTIONS) data.take(data.u16());
                outline = { components };
            } else throw new UnreadableFont(`its glyph ${id} has ${contours} contours`);
        }
        this.read.set(id, outline);
        return outline;
    }

    /**
     * Checks that the outline of each of these glyphs, and of each glyph they are made of, can be read as
     * viewers read it (see outline), its components glyphs of the font that are not made of themselves.
     * Throws, naming the glyph, for one that cannot be: a font embedded with it drew nothing there, and
     * fontkit, reading a glyph made of itself, filled the heap and ended the process.
     */
    check(glyphs: Iterable<number>): void {
        const open = new Set<number>();
        // (Each glyph is read once for the font: a glyph checked before is not read again, only how deep
        // its components nest is added to where it is met. Checked anew for each glyph, one glyph of many
        // components that many glyphs were made of was read again for each.)
        const check = (id: number, depth: number): number => {
            const below = this.checkedDepths.get(id);
            if (below !== undefined) {
                if (depth + below > MAX_COMPONENT_DEPTH) throw new UnreadableFont(`its glyph ${id} nests components more than ${MAX_COMPONENT_DEPTH} deep`);
                return below;
            }
            if (open.has(id)) throw new UnreadableFont(`its glyph ${id} is made of itself`);
            if (depth > MAX_COMPONENT_DEPTH) throw new UnreadableFont(`its glyph ${id} nests components more than ${MAX_COMPONENT_DEPTH} deep`);
            open.add(id);
            const outline = this.outline(id);
            let deepest = 0;
            if ('components' in outline) for (const component of outline.components) deepest = Math.max(deepest, 1 + check(component, depth + 1));
            open.delete(id);
            this.checkedDepths.set(id, deepest);
            return deepest;
        };
        for (const id of glyphs) check(id, 0);
    }

    /** Whether a glyph's outline, and its components', can be read as viewers read them (see check). */
    reads(id: number): boolean {
        try { this.check([id]); return true; } catch { return false; }
    }

    /**
     * Whether a glyph draws anything: a contour of three points or more, its own or a component's. An
     * empty glyph draws nothing (a font's blank, or its defect), nor does the glyph of a character in a
     * color font (Apple Color Emoji, Noto Color Emoji), drawn from the color tables, whose outline is
     * empty or a placeholder of single points; one that cannot be read draws nothing either.
     */
    draws(id: number): boolean {
        const known = this.drawing.get(id);
        if (known !== undefined) return known;
        let draws: boolean;
        try { this.check([id]); draws = this.drawsChecked(id); } catch { draws = false; }
        this.drawing.set(id, draws);
        return draws;
    }

    /** Whether a glyph that was checked (so is made of no glyph made of itself) draws anything, each component once. */
    private drawsChecked(id: number): boolean {
        const known = this.drawing.get(id);
        if (known !== undefined) return known;
        const outline = this.outline(id);
        const draws = 'contours' in outline ? outline.contours.some(points => points >= 3) : outline.components.some(component => this.drawsChecked(component));
        this.drawing.set(id, draws);
        return draws;
    }
}

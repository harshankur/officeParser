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
 * CffOutlines for a font of CFF outlines, and the engine's BLANK).
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

/** What a CFF table is called where it cannot be read. */
const CFF_TABLE = 'its CFF table';

/**
 * An INDEX of a CFF table (the CFF specification's section 5): its number of items, and where each lies,
 * read only for the item asked for. Throws where the INDEX does not fit in the table.
 */
class CffIndex {
    readonly count: number;
    /** Where the INDEX ends, and the structure after it starts. */
    readonly end: number;
    private readonly offSize: number = 0;
    private readonly offsets: number = 0;
    /** The byte before the items' data, from which their offsets count (the first is 1). */
    private readonly data: number = 0;
    constructor(private readonly table: Uint8Array, at: number, cff2: boolean) {
        const r = new Reader(table, at, table.length, CFF_TABLE);
        this.count = cff2 ? r.u32() : r.u16();
        if (!this.count) { this.end = r.pos; return; }
        this.offSize = r.u8();
        if (this.offSize < 1 || this.offSize > 4) throw new UnreadableFont(`${CFF_TABLE} has an INDEX of offsets ${this.offSize} bytes long`);
        this.offsets = r.pos;
        r.take((this.count + 1) * this.offSize);
        this.data = r.pos - 1;
        this.end = this.data + this.offset(this.count);
        if (this.end < r.pos || this.end > table.length) throw new UnreadableFont(`${CFF_TABLE} has an INDEX that runs past its end`);
    }

    private offset(i: number): number {
        let value = 0;
        for (let b = 0, at = this.offsets + i * this.offSize; b < this.offSize; b++) value = value * 256 + this.table[at + b];
        return value;
    }

    /** Where an item lies, its start and end in the table, or undefined for one not inside the INDEX's data. */
    item(i: number): [number, number] | undefined {
        if (!Number.isInteger(i) || i < 0 || i >= this.count) return undefined;
        const start = this.data + this.offset(i), end = this.data + this.offset(i + 1);
        return start > this.data && start <= end && end <= this.end ? [start, end] : undefined;
    }
}

/**
 * A DICT of a CFF table (the CFF specification's section 4): each operator's operands, an escaped
 * operator (12 n) as 1200 + n. Throws for one that cannot be read.
 */
function cffDict(table: Uint8Array, start: number, end: number): Map<number, number[]> {
    if (start < 0 || start > end || end > table.length) throw new UnreadableFont(`${CFF_TABLE} has a DICT that lies outside it`);
    const dict = new Map<number, number[]>();
    let operands: number[] = [];
    const r = new Reader(table, start, end, `a DICT of ${CFF_TABLE}`);
    while (r.pos < end) {
        const b0 = r.u8();
        if (b0 === 12) { dict.set(1200 + r.u8(), operands); operands = []; }
        else if (b0 <= 27) { dict.set(b0, operands); operands = []; }
        else if (b0 === 28) operands.push(r.i16());
        else if (b0 === 29) operands.push(r.u32() | 0);
        else if (b0 === 30) {
            // (A real number: a nibble for each digit, point, exponent or sign, ended by 0xF.)
            let text = '';
            for (let ended = false; !ended;) {
                const byte = r.u8();
                for (const nibble of [byte >> 4, byte & 0xF]) {
                    if (nibble === 0xF) { ended = true; break; }
                    text += nibble <= 9 ? String(nibble) : ['.', 'E', 'E-', '', '-'][nibble - 10];
                }
            }
            operands.push(Number(text));
        }
        else if (b0 >= 32 && b0 <= 246) operands.push(b0 - 139);
        else if (b0 >= 247 && b0 <= 250) operands.push((b0 - 247) * 256 + r.u8() + 108);
        else if (b0 >= 251 && b0 <= 254) operands.push(-(b0 - 251) * 256 - r.u8() - 108);
        else throw new UnreadableFont(`a DICT of ${CFF_TABLE} holds the reserved byte ${b0}`);
    }
    return dict;
}

/**
 * The CFF outlines (`CFF `, or a variable font's `CFF2`) of a font, or undefined for a font of other
 * outlines (TrueType). Throws for a CFF table whose glyphs cannot be found in it.
 */
export function cffOutlines(sfnt: Uint8Array): CffOutlines | undefined {
    // (Of a font with both, fontkit draws the CFF2 table's.)
    const table = sfntTable(sfnt, 'CFF2') ?? sfntTable(sfnt, 'CFF ');
    return table && new CffOutlines(table);
}

/** How a CFF glyph draws: something, nothing, or what cannot be read. */
type CffDrawing = 'drawn' | 'empty' | 'unreadable';

/** The operands of one line or curve of each Type 2 operator that draws them (an escaped one, 12 n, as 1200 + n). */
const CFF_SEGMENT_OPERANDS = new Map<number, number>([
    [5, 2], [6, 1], [7, 1], [8, 6], [24, 2], [25, 6], [26, 4], [27, 4], [30, 4], [31, 4], // rlineto to hvcurveto
    [1234, 7], [1235, 13], [1236, 9], [1237, 11], // hflex, flex, hflex1, flex1
]);

/** The operands each arithmetic operator of a CFF (not CFF2) charstring takes (12 n, by n), and how many it leaves more or fewer. */
const CFF_ARITHMETIC = new Map<number, [number, number]>([
    [3, [2, -1]], [4, [2, -1]], [5, [1, 0]], [9, [1, 0]], [10, [2, -1]], [11, [2, -1]], [12, [2, -1]], [14, [1, 0]], [15, [2, -1]],
    [18, [1, -1]], [20, [2, -2]], [21, [1, 0]], [22, [4, -3]], [23, [0, 1]], [24, [2, -1]], [26, [1, 0]], [27, [1, 1]], [28, [2, 0]],
    [29, [1, 0]], [30, [2, -2]],
]);

/**
 * The CFF outlines of a font: whether each glyph draws something, read from its Type 2 charstring (Adobe's
 * Technical Note #5177, and the OpenType CFF2 table's) without running it.
 *
 * fontkit's outline of a CFF glyph is not read for this. @pdf-lib/fontkit 1.1.1 runs the first glyph of each
 * range of a CID-keyed font's FDSelect (Noto Sans KR, SC, TC and JP) with the font dictionary of the range
 * before it, so with its subroutines, and the outline comes back empty, or not at all; it draws nothing
 * of an accented glyph made of two others (seac); and it runs subroutines nested to any depth, in time
 * that doubles with each level of a subroutine calling the next twice.
 *
 * A glyph draws something when its charstring draws a line or a curve, calls a subroutine (none is run:
 * one called is taken to draw), or ends as an accented glyph of two others (endchar with four operands,
 * or five with its width). It is empty when it ends (at endchar, or at its end) having done none of these.
 * Its stem hints are counted, so a hint mask's bytes are passed over, not read as operators. A charstring
 * that cannot be read (an operator reserved or unknown, operands missing or too many, data past its end)
 * draws nothing. Each glyph is read once, in time linear in its charstring.
 */
export class CffOutlines {
    private readonly cff2: boolean;
    private readonly topDict: Map<number, number[]>;
    private readonly charStrings: CffIndex;
    /** The most operands a charstring's stack holds: 48 in CFF, a CFF2 font's maxstack (193 unless it says) in CFF2. */
    private readonly maxStack: number;
    private readonly drawing = new Map<number, CffDrawing>();
    /** A CFF2 font's, read when a glyph first blends: the regions of each of its variation store's data, and its font dictionaries. */
    private regionCounts?: number[];
    private fontDicts?: Map<number, number[]>[];

    constructor(private readonly table: Uint8Array) {
        const header = new Reader(table, 0, table.length, CFF_TABLE);
        const major = header.u8();
        header.u8(); // minor version
        const headerSize = header.u8();
        if (major === 1) {
            this.cff2 = false;
            const names = new CffIndex(table, headerSize, false);
            const top = new CffIndex(table, names.end, false).item(0);
            if (!top) throw new UnreadableFont(`${CFF_TABLE} has no Top DICT`);
            this.topDict = cffDict(table, top[0], top[1]);
            this.maxStack = 48;
        } else if (major === 2) {
            this.cff2 = true;
            this.topDict = cffDict(table, headerSize, headerSize + header.u16());
            this.maxStack = Math.min(this.topDict.get(25)?.[0] ?? 193, 513);
        } else throw new UnreadableFont(`${CFF_TABLE} is of version ${major}`);
        const type = this.topDict.get(1206)?.[0] ?? 2;
        if (type !== 2) throw new UnreadableFont(`${CFF_TABLE}'s glyphs are charstrings of type ${type}, not 2`);
        const at = this.topDict.get(17)?.[0];
        if (at === undefined) throw new UnreadableFont(`${CFF_TABLE} has no CharStrings`);
        this.charStrings = new CffIndex(table, at, this.cff2);
    }

    /** Whether a glyph draws something (see the class). */
    draws(id: number): boolean {
        return this.drawingOf(id) === 'drawn';
    }

    /** Whether a glyph's charstring can be read: one that cannot draws nothing. */
    reads(id: number): boolean {
        return this.drawingOf(id) !== 'unreadable';
    }

    private drawingOf(id: number): CffDrawing {
        let known = this.drawing.get(id);
        if (known === undefined) {
            try { known = this.scan(id); } catch { known = 'unreadable'; }
            this.drawing.set(id, known);
        }
        return known;
    }

    /** Reads a glyph's charstring, operator by operator, until it draws or ends (see the class). */
    private scan(id: number): CffDrawing {
        const where = this.charStrings.item(id);
        if (!where) return 'unreadable';
        const b = this.table, cff2 = this.cff2;
        let at = where[0];
        const end = where[1];
        // (The operands on the stack: each number pushed, or NaN for one worked out. Only CFF2's vsindex
        // and blend read them, and a CFF2 charstring works none out: it has no arithmetic.)
        const stack: number[] = [];
        let stems = 0;
        let regions: number | undefined;
        while (at < end) {
            const op = b[at++];
            if (op >= 32 || op === 28) {
                let value: number;
                if (op === 28) { if (at + 2 > end) return 'unreadable'; value = ((b[at] << 24) >> 16) | b[at + 1]; at += 2; }
                else if (op <= 246) value = op - 139;
                else if (op === 255) { if (at + 4 > end) return 'unreadable'; value = ((b[at] << 24) | (b[at + 1] << 16) | (b[at + 2] << 8) | b[at + 3]) / 65536; at += 4; }
                else { if (at >= end) return 'unreadable'; value = op <= 250 ? (op - 247) * 256 + b[at] + 108 : -(op - 251) * 256 - b[at] - 108; at++; }
                if (stack.length >= this.maxStack) return 'unreadable';
                stack.push(value);
                continue;
            }
            let key = op;
            if (op === 12) { if (at >= end) return 'unreadable'; key = 1200 + b[at++]; }
            const segment = CFF_SEGMENT_OPERANDS.get(key);
            if (segment !== undefined) {
                // (Too few operands for one line or curve draws none.)
                if (stack.length >= segment) return 'drawn';
                stack.length = 0;
                continue;
            }
            switch (key) {
                case 10: case 29: return 'drawn'; // callsubr, callgsubr
                case 1: case 3: case 18: case 23: // hstem, vstem, hstemhm, vstemhm: two operands a stem (a width first, where odd)
                    stems += stack.length >> 1;
                    stack.length = 0;
                    break;
                case 19: case 20: // hintmask, cntrmask: the vstems left before it, then a bit for each stem
                    stems += stack.length >> 1;
                    stack.length = 0;
                    at += (stems + 7) >> 3;
                    if (at > end) return 'unreadable';
                    break;
                case 4: case 21: case 22: // vmoveto, rmoveto, hmoveto
                    stack.length = 0;
                    break;
                case 14: // endchar, with the two glyphs of an accented one (seac) where it has four operands, or five
                    if (cff2) return 'unreadable';
                    return stack.length === 4 || stack.length === 5 ? 'drawn' : 'empty';
                case 11: // return, which ends a charstring outside a subroutine
                    return cff2 ? 'unreadable' : 'empty';
                case 1200: // dotsection, which does nothing
                    if (cff2) return 'unreadable';
                    stack.length = 0;
                    break;
                case 15: // vsindex: the variation store data blends are by
                    if (!cff2 || !stack.length) return 'unreadable';
                    regions = this.regionCount(stack.pop()!);
                    if (regions === undefined) return 'unreadable';
                    break;
                case 16: { // blend: n values, a delta for each region of each, then n; n values left
                    if (!cff2 || !stack.length) return 'unreadable';
                    regions ??= this.regionCount(this.defaultVsindex(id));
                    const n = stack.pop()!;
                    if (regions === undefined || !Number.isInteger(n) || n < 0 || stack.length < n * (regions + 1)) return 'unreadable';
                    stack.length -= n * regions;
                    stack.fill(NaN, stack.length - n);
                    break;
                }
                default: {
                    const arithmetic = !cff2 && key > 1200 ? CFF_ARITHMETIC.get(key - 1200) : undefined;
                    if (!arithmetic) return 'unreadable';
                    const [takes, leaves] = arithmetic;
                    if (stack.length < takes || stack.length + leaves > this.maxStack) return 'unreadable';
                    if (leaves > 0) stack.push(NaN);
                    else stack.length += leaves;
                }
            }
        }
        // (A CFF2 charstring ends at its end, and so does a CFF one that runs out before endchar.)
        return 'empty';
    }

    /** How many regions a CFF2 font's variation store gives its data of this index, or undefined where it has none. */
    private regionCount(vsindex: number): number | undefined {
        if (!this.regionCounts) {
            this.regionCounts = [];
            const at = this.topDict.get(24)?.[0];
            if (at !== undefined) {
                try {
                    // (Its length, then an ItemVariationStore, whose data lie at offsets from its start.)
                    const store = at + 2;
                    const r = new Reader(this.table, store, this.table.length, CFF_TABLE);
                    r.u16(); r.u32(); // format, the region list's offset
                    const count = r.u16();
                    const offsets: number[] = [];
                    for (let i = 0; i < count; i++) offsets.push(r.u32());
                    for (const offset of offsets) {
                        const data = new Reader(this.table, store + offset, this.table.length, CFF_TABLE);
                        data.u16(); data.u16(); // item count, word delta count
                        this.regionCounts.push(data.u16());
                    }
                } catch { this.regionCounts = []; }
            }
        }
        return Number.isInteger(vsindex) ? this.regionCounts[vsindex] : undefined;
    }

    /** The variation store data a CFF2 glyph blends by unless it says: its font dictionary's Private DICT's vsindex, or 0. */
    private defaultVsindex(id: number): number {
        try {
            if (!this.fontDicts) {
                this.fontDicts = [];
                const at = this.topDict.get(1236)?.[0];
                if (at !== undefined) {
                    const fds = new CffIndex(this.table, at, true);
                    for (let i = 0; i < fds.count; i++) {
                        const item = fds.item(i);
                        this.fontDicts.push(item ? cffDict(this.table, item[0], item[1]) : new Map());
                    }
                }
            }
            const priv = this.fontDicts[this.fdOf(id)]?.get(18);
            if (!priv || priv.length < 2) return 0;
            return cffDict(this.table, priv[1], priv[1] + priv[0]).get(22)?.[0] ?? 0;
        } catch { return 0; }
    }

    /** A CFF2 glyph's font dictionary, by the font's FDSelect (of format 0, 3 or 4), or the first where it has none. */
    private fdOf(id: number): number {
        const at = this.topDict.get(1237)?.[0];
        if (at === undefined) return 0;
        const r = new Reader(this.table, at, this.table.length, CFF_TABLE);
        const format = r.u8();
        if (format === 0) { r.take(id); return r.u8(); }
        if (format !== 3 && format !== 4) return 0;
        // (Ranges in order of their first glyph, each up to the next one's first: the last range whose
        // first glyph is at or before this one is its range.)
        const wide = format === 4, size = wide ? 6 : 3;
        const ranges = wide ? r.u32() : r.u16();
        const records = r.pos;
        r.take(ranges * size);
        const firstOf = (i: number) => new Reader(this.table, records + i * size, this.table.length, CFF_TABLE)[wide ? 'u32' : 'u16']();
        if (!ranges || id < firstOf(0)) return 0;
        let lo = 0, hi = ranges - 1;
        while (lo < hi) {
            const mid = (lo + hi + 1) >> 1;
            if (firstOf(mid) <= id) lo = mid; else hi = mid - 1;
        }
        const fd = new Reader(this.table, records + lo * size + (wide ? 4 : 2), this.table.length, CFF_TABLE);
        return wide ? fd.u16() : fd.u8();
    }
}

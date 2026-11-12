/**
 * Windows metafiles (EMF and WMF): their media types, their sizes from their own headers, WMF's
 * placeable header, which DOCX and ODT packages hold and RTF leaves out, and the bitmap a metafile draws
 * (readMetafile), which the native PDF engine draws it from.
 *
 * Pure byte functions. Each reads a fixed number of bytes, or a bounded run of records, per picture.
 *
 * @module metafileUtils
 */

import { pictureSignature } from './imageUtils.js';
import { lookupTable } from './lookupUtils.js';
import { sniffImageSize } from './officeGenUtils.js';
import { encodePng, MAX_RASTER_PIXELS, tiffSize, type BitmapTransparency } from './rasterUtils.js';

/** The extension of a metafile packaged in DOCX or ODT output, by its media type (the RFC 7903 names and the older `x-` ones). */
export const METAFILE_EXT: Readonly<Record<string, 'emf' | 'wmf'>> = lookupTable({
    'image/emf': 'emf', 'image/x-emf': 'emf',
    'image/wmf': 'wmf', 'image/x-wmf': 'wmf',
});

/** The content type a DOCX or ODT package gives a metafile part: what Word and LibreOffice write. */
export const METAFILE_CONTENT_TYPE: Readonly<Record<'emf' | 'wmf', string>> = lookupTable({
    emf: 'image/x-emf',
    wmf: 'image/x-wmf',
});

/** The key a placeable WMF starts with (0x9AC6CDD7, little-endian). */
const PLACEABLE_KEY = [0xD7, 0xCD, 0xC6, 0x9A];
/** The size of a placeable WMF's header. */
const PLACEABLE_HEADER_BYTES = 22;
/** How many of a WMF's first records are looked through for its window. */
const MAX_WINDOW_RECORDS = 64;
/** The largest whole number RTF takes as a parameter. */
const MAX_RTF_PARAM = 2147483647;

const u16 = (b: Uint8Array, at: number) => b[at] | (b[at + 1] << 8);
const i16 = (b: Uint8Array, at: number) => (u16(b, at) << 16) >> 16;
const u32 = (b: Uint8Array, at: number) => (b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24)) >>> 0;
const i32 = (b: Uint8Array, at: number) => b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24);

/** Whether `bytes` start with a placeable WMF's header. */
const isPlaceable = (bytes: Uint8Array) => bytes.length >= PLACEABLE_HEADER_BYTES && PLACEABLE_KEY.every((v, i) => bytes[i] === v);

/** Whether `bytes` are an EMF: an EMR_HEADER (type 1) whose signature at byte 40 is " EMF". */
const isEmf = (bytes: Uint8Array) => bytes.length >= 44 && u32(bytes, 0) === 1
    && bytes[40] === 0x20 && bytes[41] === 0x45 && bytes[42] === 0x4D && bytes[43] === 0x46;

/** A size kept only when both sides are positive and finite. */
const sized = (w: number, h: number): { w: number; h: number } | null =>
    Number.isFinite(w) && Number.isFinite(h) && w > 0 && h > 0 ? { w, h } : null;

/**
 * A metafile's size at 96 pixels an inch, from its own header: an EMF's frame (rclFrame, hundredths of a
 * millimetre, offset 24) or a placeable WMF's bounding box and units per inch. Null for anything else, a
 * headerless WMF included, or a size that is not positive and finite.
 */
export function metafileSize(bytes: Uint8Array): { w: number; h: number } | null {
    if (isEmf(bytes) && bytes.length >= 40) {
        const w = i32(bytes, 32) - i32(bytes, 24), h = i32(bytes, 36) - i32(bytes, 28);
        return sized(w / 2540 * 96, h / 2540 * 96);
    }
    if (isPlaceable(bytes)) {
        const inch = u16(bytes, 14);
        if (!inch) return null;
        const w = Math.abs(i16(bytes, 10) - i16(bytes, 6)), h = Math.abs(i16(bytes, 12) - i16(bytes, 8));
        return sized(w / inch * 96, h / inch * 96);
    }
    return null;
}

/** Twips, HIMETRIC and the other mapping modes' units an inch (`\wmetafileN`'s N), for `\picw`. */
const MAP_MODE_UNITS_PER_INCH: Readonly<Record<number, number>> = {
    1: 96,      // MM_TEXT: pixels, at 96 an inch
    2: 254,     // MM_LOMETRIC: 0.1 mm
    3: 2540,    // MM_HIMETRIC: 0.01 mm
    4: 100,     // MM_LOENGLISH: 0.01 in
    5: 1000,    // MM_HIENGLISH: 0.001 in
    6: 1440,    // MM_TWIPS
    7: 2540,    // MM_ISOTROPIC: \picw in HIMETRIC
    8: 2540,    // MM_ANISOTROPIC: \picw in HIMETRIC
};

/** A WMF's window (the first META_SETWINDOWORG and META_SETWINDOWEXT among its first records), or null. */
function wmfWindow(data: Uint8Array): { x: number; y: number; w: number; h: number } | null {
    if (data.length < 18) return null;
    let at = u16(data, 2) * 2;
    if (at < 18) return null;
    let org: { x: number; y: number } | undefined, ext: { w: number; h: number } | undefined;
    for (let record = 0; record < MAX_WINDOW_RECORDS && at + 6 <= data.length; record++) {
        const words = u32(data, at), fn = u16(data, at + 4);
        if (fn === 0x0000 || words < 3) break;
        if (fn === 0x020B && at + 10 <= data.length && !org) org = { y: i16(data, at + 6), x: i16(data, at + 8) };
        else if (fn === 0x020C && at + 10 <= data.length && !ext) ext = { h: i16(data, at + 6), w: i16(data, at + 8) };
        if (org && ext) break;
        if (words > (data.length - at) / 2) break;
        at += words * 2;
    }
    return ext ? { x: org?.x ?? 0, y: org?.y ?? 0, w: ext.w, h: ext.h } : null;
}

/**
 * RTF's `\wmetafileN` data (a METAFILE with no placeable header, MS-WMF 2.3.2.3) as a standalone .wmf file,
 * as DOCX and ODT packages hold one: the 22-byte placeable header prepended. Its bounding box is the
 * metafile's window (the first META_SETWINDOWORG 0x020B and META_SETWINDOWEXT 0x020C among its first 64
 * records), its units per inch the window's width over the picture's width in inches: `\picwgoal` (twips) when
 * positive, else `\picw` in the mapping mode's units (HIMETRIC for modes 7 and 8; 0.1 mm, 0.01 mm, 0.01 in,
 * 0.001 in, twips for 2 to 6; pixels at 96 an inch for 1). Returned unchanged when it already has the header,
 * when no window or no width is found, or when the units per inch fall outside 1 to 65535.
 */
export function placeableWmf(data: Uint8Array, size: { mapMode?: number; picw?: number; pich?: number; goalWidth?: number; goalHeight?: number }): Uint8Array {
    if (isPlaceable(data)) return data;
    const window = wmfWindow(data);
    if (!window || window.w === 0 || window.h === 0) return data;
    const perInch = MAP_MODE_UNITS_PER_INCH[size.mapMode ?? 8] ?? 2540;
    const inches = size.goalWidth !== undefined && size.goalWidth > 0 ? size.goalWidth / 1440
        : size.picw !== undefined && size.picw > 0 ? size.picw / perInch : 0;
    if (!(inches > 0) || !Number.isFinite(inches)) return data;
    const inch = Math.round(Math.abs(window.w) / inches);
    if (!(inch >= 1 && inch <= 65535)) return data;
    const left = Math.min(window.x, window.x + window.w), right = Math.max(window.x, window.x + window.w);
    const top = Math.min(window.y, window.y + window.h), bottom = Math.max(window.y, window.y + window.h);
    if ([left, right, top, bottom].some(v => v < -32768 || v > 32767)) return data;
    const words = [0xCDD7, 0x9AC6, 0, left & 0xFFFF, top & 0xFFFF, right & 0xFFFF, bottom & 0xFFFF, inch, 0, 0];
    const out = new Uint8Array(PLACEABLE_HEADER_BYTES + data.length);
    let checksum = 0;
    words.forEach((word, i) => { out[i * 2] = word & 0xFF; out[i * 2 + 1] = word >>> 8; checksum ^= word; });
    out[20] = checksum & 0xFF; out[21] = checksum >>> 8;
    out.set(data, PLACEABLE_HEADER_BYTES);
    return out;
}

/** A length in inches as a whole RTF parameter of `perInch` units, from 1 to 2147483647. */
export const rtfLength = (inches: number, perInch: number) => Math.min(MAX_RTF_PARAM, Math.max(1, Math.round(inches * perInch)));

/**
 * The other way, for RTF output: a WMF's METAFILE without its placeable header, and the picture's size in
 * HIMETRIC (`\picw`, `\pich`) and twips (`\picwgoal`, `\pichgoal`) from the header, each a whole number
 * from 1 to 2147483647; the sizes absent for a headerless WMF.
 */
export function rtfWmf(bytes: Uint8Array): { data: Uint8Array; picw?: number; pich?: number; goalWidth?: number; goalHeight?: number } {
    if (!isPlaceable(bytes)) return { data: bytes };
    const data = bytes.subarray(PLACEABLE_HEADER_BYTES);
    const inch = u16(bytes, 14);
    const w = Math.abs(i16(bytes, 10) - i16(bytes, 6)) / inch, h = Math.abs(i16(bytes, 12) - i16(bytes, 8)) / inch;
    if (!inch || !(w > 0) || !(h > 0)) return { data };
    return { data, picw: rtfLength(w, 2540), pich: rtfLength(h, 2540), goalWidth: rtfLength(w, 1440), goalHeight: rtfLength(h, 1440) };
}

// ─── What a metafile holds that a PDF can show ───────────────────────────────

/** What a metafile holds that a PDF can show: the bitmap drawn over its picture, and what else it draws. */
export interface MetafileReading {
    /** The metafile's size in CSS pixels: an EMF's frame (rclFrame, 0.01 mm), a placeable WMF's box and inch. */
    frame?: { width: number; height: number };
    /**
     * The bitmap drawn with the largest area: a PNG or JPEG file as held (BI_PNG, BI_JPEG, an EMF+ compressed bitmap),
     * a device-independent bitmap given a BMP file header, or an EMF+ bitmap of pixels as a PNG. `crop` is the part
     * of it drawn, in its pixels from its top left, when that is not all of it.
     */
    bitmap?: {
        format: 'PNG' | 'JPEG' | 'GIF' | 'TIFF' | 'BMP'; bytes: Uint8Array; crop?: { x: number; y: number; width: number; height: number };
        /**
         * How the record draws it over what is under it, when not opaque: EMR_ALPHABLEND's per-pixel alpha (its 32-bit
         * pixels premultiplied, AC_SRC_ALPHA) and constant alpha (0 to 255), EMR_TRANSPARENTBLT's colour drawn transparent.
         */
        transparency?: BitmapTransparency;
    };
    /** What it draws besides that bitmap, for a warning: 'lines', 'shapes', 'text', 'other bitmaps'. */
    alsoDraws: string[];
    /** Why no bitmap was read, when none was. */
    reason?: string;
}

/** What a drawing record draws, for {@link MetafileReading.alsoDraws}. */
type Drawn = 'lines' | 'shapes' | 'text';

/** EMF drawing records ([MS-EMF] 2.3) by what they draw; records inside a path's definition are not drawn by themselves. */
const EMF_DRAWS: Readonly<Record<number, Drawn>> = lookupTable({
    2: 'lines', 4: 'lines', 5: 'lines', 6: 'lines', 7: 'lines', 45: 'lines', 54: 'lines', 55: 'lines', 56: 'lines',
    64: 'lines', 72: 'lines', 85: 'lines', 87: 'lines', 88: 'lines', 89: 'lines', 90: 'lines', 92: 'lines',
    3: 'shapes', 8: 'shapes', 15: 'shapes', 42: 'shapes', 43: 'shapes', 44: 'shapes', 46: 'shapes', 47: 'shapes',
    53: 'shapes', 62: 'shapes', 63: 'shapes', 71: 'shapes', 73: 'shapes', 74: 'shapes', 86: 'shapes', 91: 'shapes', 118: 'shapes',
    83: 'text', 84: 'text', 96: 'text', 97: 'text', 108: 'text',
} as Record<number, Drawn>);

/** EMF records that draw what a path defines (EMR_FILLPATH, EMR_STROKEANDFILLPATH, EMR_STROKEPATH): drawn even when they follow its definition. */
const EMF_PATH_DRAWS = new Set([62, 63, 64]);

/**
 * EMF bitmap records and where their fields are ([MS-EMF] 2.3.1): the destination's size (or, for EMR_PLGBLT, its
 * three points), the source rectangle, and the DIB's header and bits (offsets from the record's start). `dib`: the
 * source rectangle is in the DIB's own rows, counted from its foot when it is stored bottom up; the others' sources
 * are a device context's, counted from the top.
 */
const EMF_BITMAP_RECORDS: Readonly<Record<number, { dest: [number, number] | 'points'; src: [number, number, number, number]; bmi: number; dib?: true }>> = lookupTable({
    76: { dest: [32, 36], src: [44, 48, 32, 36], bmi: 84 },          // EMR_BITBLT: the source as large as the destination
    77: { dest: [32, 36], src: [44, 48, 100, 104], bmi: 84 },        // EMR_STRETCHBLT
    78: { dest: [32, 36], src: [44, 48, 32, 36], bmi: 84 },          // EMR_MASKBLT
    79: { dest: 'points', src: [48, 52, 56, 60], bmi: 96 },          // EMR_PLGBLT
    80: { dest: [40, 44], src: [32, 36, 40, 44], bmi: 48, dib: true }, // EMR_SETDIBITSTODEVICE
    81: { dest: [72, 76], src: [32, 36, 40, 44], bmi: 48, dib: true }, // EMR_STRETCHDIBITS
    114: { dest: [32, 36], src: [44, 48, 100, 104], bmi: 84 },       // EMR_ALPHABLEND
    116: { dest: [32, 36], src: [44, 48, 100, 104], bmi: 84 },       // EMR_TRANSPARENTBLT
} as Record<number, { dest: [number, number] | 'points'; src: [number, number, number, number]; bmi: number; dib?: true }>);

/** WMF drawing records ([MS-WMF] 2.3) by what they draw. */
const WMF_DRAWS: Readonly<Record<number, Drawn>> = lookupTable({
    0x0213: 'lines', 0x0325: 'lines', 0x0817: 'lines', 0x0429: 'lines',
    0x0324: 'shapes', 0x0538: 'shapes', 0x041B: 'shapes', 0x061C: 'shapes', 0x0418: 'shapes', 0x0830: 'shapes', 0x081A: 'shapes',
    0x0228: 'shapes', 0x012B: 'shapes', 0x012A: 'shapes', 0x0419: 'shapes', 0x0548: 'shapes', 0x041F: 'shapes', 0x061D: 'shapes',
    0x0521: 'text', 0x0A32: 'text',
} as Record<number, Drawn>);

/** WMF records holding a device-independent bitmap: the DIB's offset, destination size and source rectangle (offsets from the record's start); `own` as an EMF record's `dib`. */
const WMF_DIB_RECORDS: Readonly<Record<number, { dib: number; dest: [number, number]; src: [number, number, number, number]; own?: true }>> = lookupTable({
    0x0F43: { dib: 28, dest: [22, 20], src: [18, 16, 14, 12], own: true }, // META_STRETCHDIB
    0x0B41: { dib: 26, dest: [20, 18], src: [16, 14, 12, 10] },      // META_DIBSTRETCHBLT
    0x0940: { dib: 22, dest: [16, 14], src: [12, 10, 16, 14] },      // META_DIBBITBLT: the source as large as the destination
    0x0D33: { dib: 24, dest: [18, 16], src: [14, 12, 18, 16], own: true }, // META_SETDIBTODEV
} as Record<number, { dib: number; dest: [number, number]; src: [number, number, number, number]; own?: true }>);

/** WMF records of device-dependent bitmaps (META_BITBLT, META_STRETCHBLT): other bitmaps, or a filled rectangle without one. */
const WMF_DDB_RECORDS = new Set([0x0922, 0x0B23]);

/** EMF+ record types ([MS-EMFPLUS] 2.1.1.1) read: an object, the two that draw an image, and GetDC. */
const EMFPLUS = { object: 0x4008, drawImage: 0x401A, drawImagePoints: 0x401B, getDc: 0x4004, header: 0x4001, endOfFile: 0x4002 } as const;

/** EMF+ drawing records that are not images, by what they draw. */
const EMFPLUS_DRAWS: Readonly<Record<number, Drawn>> = lookupTable({
    0x4009: 'shapes', 0x400A: 'shapes', 0x400B: 'lines', 0x400C: 'shapes', 0x400D: 'lines', 0x400E: 'shapes', 0x400F: 'lines',
    0x4010: 'shapes', 0x4011: 'lines', 0x4012: 'lines', 0x4013: 'shapes', 0x4014: 'shapes', 0x4015: 'lines', 0x4016: 'shapes',
    0x4017: 'lines', 0x4018: 'lines', 0x4019: 'lines', 0x401C: 'text', 0x4036: 'text',
} as Record<number, Drawn>);

/** EMF+ pixel formats read from an uncompressed bitmap: bytes a pixel, and whether it has alpha (premultiplied or not). */
const EMFPLUS_PIXEL_FORMATS: Readonly<Record<number, { bytes: 3 | 4; alpha: 'none' | 'straight' | 'premultiplied' }>> = lookupTable({
    0x00021808: { bytes: 3, alpha: 'none' },          // PixelFormat24bppRGB
    0x00022009: { bytes: 4, alpha: 'none' },          // PixelFormat32bppRGB
    0x0026200A: { bytes: 4, alpha: 'straight' },      // PixelFormat32bppARGB
    0x000E200B: { bytes: 4, alpha: 'premultiplied' }, // PixelFormat32bppPARGB
} as Record<number, { bytes: 3 | 4; alpha: 'none' | 'straight' | 'premultiplied' }>);

const f32 = (b: Uint8Array, at: number) => new DataView(b.buffer, b.byteOffset + at, 4).getFloat32(0, true);

/** A bitmap found, with the area it is drawn over. */
interface Found { bitmap: NonNullable<MetafileReading['bitmap']>; area: number }

/**
 * A device-independent bitmap (its header, colour table and bits) as a picture file: a BMP file header put before
 * it, or the PNG or JPEG file it holds (BI_PNG 5, BI_JPEG 4). `bits` is where its pixels start when they are apart
 * from its header (EMF), else they follow its colour table (WMF). Null when its header is not one of a DIB or runs
 * past the record.
 */
function dibAsFile(record: Uint8Array, bmi: number, bmiSize: number | undefined, bits: number | undefined, bitsSize: number | undefined): { format: 'PNG' | 'JPEG' | 'BMP'; bytes: Uint8Array } | null {
    if (bmi < 0 || bmi + 12 > record.length) return null;
    const headerSize = u32(record, bmi);
    if (headerSize !== 12 && !(headerSize >= 40 && headerSize <= 124)) return null;
    const core = headerSize === 12;
    const bitCount = core ? u16(record, bmi + 10) : u16(record, bmi + 14);
    const compression = core ? 0 : u32(record, bmi + 16);
    const used = core ? 0 : u32(record, bmi + 32);
    // The colour table: palette entries for 8 bits or fewer, three masks for BI_BITFIELDS with a 40-byte header.
    const table = bitCount <= 8 ? Math.min(used || (1 << bitCount), 256) * (core ? 3 : 4) : compression === 3 && headerSize === 40 ? 12 : 0;
    const infoSize = bmiSize ?? headerSize + table;
    const bitsAt = bits ?? bmi + headerSize + (compression === 4 || compression === 5 ? 0 : table);
    const bitsEnd = bitsSize !== undefined ? bitsAt + bitsSize : record.length;
    if (bmi + infoSize > record.length || bitsAt < 0 || bitsAt > bitsEnd || bitsEnd > record.length) return null;
    const pixels = record.subarray(bitsAt, bitsEnd);
    if (compression === 4 || compression === 5) return { format: compression === 5 ? 'PNG' : 'JPEG', bytes: pixels.slice() }; // (a copy: a view into the metafile reads as the whole of it to pdf-lib)
    const file = new Uint8Array(14 + infoSize + pixels.length);
    file[0] = 0x42; file[1] = 0x4D;
    const view = new DataView(file.buffer);
    view.setUint32(2, file.length, true);
    view.setUint32(10, 14 + infoSize, true);
    file.set(record.subarray(bmi, bmi + infoSize), 14);
    file.set(pixels, 14 + infoSize);
    return { format: 'BMP', bytes: file };
}

/** The crop a source rectangle makes of a bitmap of `width` by `height` pixels: undefined when it is all of it or not inside it. */
function cropOf(x: number, y: number, w: number, h: number, width: number, height: number): { x: number; y: number; width: number; height: number } | undefined {
    if (!(width > 0 && height > 0) || w <= 0 || h <= 0 || x < 0 || y < 0 || x + w > width || y + h > height) return undefined;
    if (x === 0 && y === 0 && w === width && h === height) return undefined;
    return { x, y, width: w, height: h };
}

/** A DIB's size in pixels from its header, and whether its rows are stored from the foot up. */
function dibSize(record: Uint8Array, bmi: number): { width: number; height: number; bottomUp: boolean } {
    if (bmi + 16 > record.length) return { width: 0, height: 0, bottomUp: false };
    if (u32(record, bmi) === 12) return { width: u16(record, bmi + 4), height: u16(record, bmi + 6), bottomUp: true };
    const height = i32(record, bmi + 8);
    return { width: Math.abs(i32(record, bmi + 4)), height: Math.abs(height), bottomUp: height > 0 };
}

/** A DIB's source rectangle as a crop from its top left: one in its own rows counted from its foot when it is stored bottom up. */
const dibCrop = (src: number[], dims: { width: number; height: number; bottomUp: boolean }, own: boolean | undefined) =>
    cropOf(src[0], own && dims.bottomUp ? dims.height - src[1] - src[3] : src[1], src[2], src[3], dims.width, dims.height);

/** The words for what a metafile draws besides its bitmap, in a fixed order. */
const ALSO_ORDER = ['lines', 'shapes', 'text', 'other bitmaps'];

/**
 * What a metafile holds that a PDF can show: the bitmap drawn over the largest area, and what else it draws. Its
 * records are walked once, by their sizes, each field checked inside its record; an EMF+ object continued across
 * records is joined only up to the metafile's own size. Never throws.
 */
export function readMetafile(bytes: Uint8Array, kind: 'EMF' | 'WMF'): MetafileReading {
    const size = metafileSize(bytes);
    const frame = size ? { width: size.w, height: size.h } : undefined;
    const drawn = new Set<string>();
    const found: Found[] = [];
    let reason: string | undefined;
    if (kind === 'EMF') reason = readEmf(bytes, drawn, found);
    else reason = readWmf(bytes, drawn, found);
    found.sort((a, b) => b.area - a.area);
    if (found.length > 1) drawn.add('other bitmaps');
    const alsoDraws = ALSO_ORDER.filter(w => drawn.has(w));
    if (!found.length) return { frame, alsoDraws, ...(reason ? { reason } : {}) };
    return { frame, bitmap: found[0].bitmap, alsoDraws };
}

/** An EMF's records walked: what they draw into `drawn`, their bitmaps into `found`; why it stopped early. */
function readEmf(bytes: Uint8Array, drawn: Set<string>, found: Found[]): string | undefined {
    // EMF+ records (in EMR_COMMENT) are classified apart: when they draw, they are the picture, and the GDI records
    // the fallback copy of it (but for those after an EmfPlusGetDC, which are drawn in the EMF+ picture).
    const gdiDrawn = new Set<string>(), gdiFound: Found[] = [];
    const plusDrawn = new Set<string>(), plusFound: Found[] = [];
    let plusDraws = false, afterGetDc = false, inPath = false;
    const objects = new Map<number, NonNullable<MetafileReading['bitmap']>>();
    let pending: { id: number; total: number; parts: Uint8Array[]; got: number } | undefined;
    let damaged: string | undefined;
    for (let at = 0; at + 8 <= bytes.length;) {
        const type = u32(bytes, at), size = u32(bytes, at + 4);
        if (size < 8 || size % 4 !== 0 || at + size > bytes.length) { damaged = 'its EMF records are damaged'; break; }
        const record = bytes.subarray(at, at + size);
        at += size;
        if (type === 14) break; // EMR_EOF
        if (type === 70 && size >= 16 && u32(record, 12) === 0x2B464D45) {
            // EMR_COMMENT holding EMF+ records ("EMF+"), each Type, Flags, Size, DataSize, data.
            const end = Math.min(record.length, 12 + u32(record, 8));
            afterGetDc = false;
            for (let p = 16; p + 12 <= end;) {
                const plusType = u16(record, p), flags = u16(record, p + 2), plusSize = u32(record, p + 4), dataSize = u32(record, p + 8);
                if (plusSize < 12 || p + plusSize > end || dataSize > plusSize - 12) break;
                const data = record.subarray(p + 12, p + 12 + dataSize);
                p += plusSize;
                if (plusType === EMFPLUS.getDc) { afterGetDc = true; continue; }
                if (plusType === EMFPLUS.object) {
                    const id = flags & 0xFF, objectType = (flags >> 8) & 0x7F;
                    if (objectType !== 5) continue;
                    if (flags & 0x8000) {
                        // Continued: the whole object's size first, then a part of it; joined only up to the metafile's size.
                        if (data.length < 4) continue;
                        const total = u32(data, 0);
                        if (!pending || pending.id !== id) pending = total <= bytes.length ? { id, total, parts: [], got: 0 } : undefined;
                        if (!pending) continue;
                        const part = data.subarray(4);
                        pending.parts.push(part);
                        pending.got += part.length;
                        if (pending.got >= pending.total || pending.got > bytes.length) {
                            const whole = new Uint8Array(Math.min(pending.got, pending.total));
                            let o = 0;
                            for (const piece of pending.parts) { const take = Math.min(piece.length, whole.length - o); whole.set(piece.subarray(0, take), o); o += take; }
                            const image = emfPlusImage(whole);
                            if (image) objects.set(id, image);
                            pending = undefined;
                        }
                    } else {
                        pending = undefined;
                        const image = emfPlusImage(data);
                        if (image) objects.set(id, image); else objects.delete(id);
                    }
                    continue;
                }
                if (plusType === EMFPLUS.drawImage || plusType === EMFPLUS.drawImagePoints) {
                    plusDraws = true;
                    const image = objects.get(flags & 0xFF);
                    if (!image) { plusDrawn.add('other bitmaps'); continue; }
                    // The source rectangle (floats, in the image's units: pixels for a bitmap) and the destination.
                    const src = data.length >= 24 ? [f32(data, 8), f32(data, 12), f32(data, 16), f32(data, 20)] : undefined;
                    let area = 0;
                    const compressed = (flags & 0x4000) !== 0;
                    if (plusType === EMFPLUS.drawImage) {
                        if (compressed && data.length >= 32) area = Math.abs(i16(data, 28) * i16(data, 30));
                        else if (data.length >= 40) area = Math.abs(f32(data, 32) * f32(data, 36));
                    } else if (data.length >= 28) {
                        const pts: number[] = [];
                        for (let k = 0; k < 6; k++) pts.push(compressed ? (data.length >= 28 + 12 ? i16(data, 28 + k * 2) : 0) : (data.length >= 28 + 24 ? f32(data, 28 + k * 4) : 0));
                        area = Math.abs((pts[2] - pts[0]) * (pts[5] - pts[1]) - (pts[3] - pts[1]) * (pts[4] - pts[0]));
                    }
                    const dims = rasterSize(image);
                    const crop = src && dims ? cropOf(Math.round(src[0]), Math.round(src[1]), Math.round(src[2]), Math.round(src[3]), dims.width, dims.height) : undefined;
                    plusFound.push({ bitmap: crop ? { ...image, crop } : image, area: Number.isFinite(area) ? area : 0 });
                    continue;
                }
                const what = EMFPLUS_DRAWS[plusType];
                if (what) { plusDraws = true; plusDrawn.add(what); }
            }
            continue;
        }
        const target = afterGetDc ? plusDrawn : gdiDrawn, targetFound = afterGetDc ? plusFound : gdiFound;
        if (type === 59) { inPath = true; continue; }    // EMR_BEGINPATH
        if (type === 60 || type === 68) { inPath = false; continue; } // EMR_ENDPATH, EMR_ABORTPATH
        const bitmapRecord = EMF_BITMAP_RECORDS[type];
        if (bitmapRecord) {
            const bmiOffset = bitmapRecord.bmi;
            if (bmiOffset + 16 > record.length) { target.add('shapes'); continue; }
            const offBmi = u32(record, bmiOffset), cbBmi = u32(record, bmiOffset + 4), offBits = u32(record, bmiOffset + 8), cbBits = u32(record, bmiOffset + 12);
            // (A blit with no bitmap fills a rectangle with the brush.)
            if (!cbBmi) { target.add('shapes'); continue; }
            const file = dibAsFile(record, offBmi, cbBmi, offBits, cbBits);
            if (!file) { target.add('other bitmaps'); continue; }
            let area: number;
            if (bitmapRecord.dest === 'points') {
                const p = [24, 28, 32, 36, 40, 44].map(o => i32(record, o));
                area = Math.abs((p[2] - p[0]) * (p[5] - p[1]) - (p[3] - p[1]) * (p[4] - p[0]));
            } else area = Math.abs(i32(record, bitmapRecord.dest[0]) * i32(record, bitmapRecord.dest[1]));
            const crop = dibCrop(bitmapRecord.src.map(o => i32(record, o)), dibSize(record, offBmi), bitmapRecord.dib);
            const transparency = blitTransparency(type, record);
            targetFound.push({ bitmap: { ...file, ...(crop ? { crop } : {}), ...(transparency ? { transparency } : {}) }, area: Number.isFinite(area) ? area : 0 });
            continue;
        }
        const what = EMF_DRAWS[type];
        if (what && (!inPath || EMF_PATH_DRAWS.has(type))) target.add(what);
    }
    const useful = plusDraws ? { d: plusDrawn, f: plusFound } : { d: gdiDrawn, f: gdiFound };
    for (const d of useful.d) drawn.add(d);
    found.push(...useful.f);
    return damaged;
}

/**
 * How an EMF bitmap record draws its bitmap when not opaque ([MS-EMF] 2.3.1.1, 2.3.1.8): EMR_ALPHABLEND's
 * BLENDFUNCTION (offset 40: SourceConstantAlpha at 42, AlphaFormat at 43, AC_SRC_ALPHA 1 for 32-bit pixels of
 * premultiplied alpha), EMR_TRANSPARENTBLT's TransparentColor (a ColorRef at 40: red, green, blue). Undefined when it
 * changes nothing.
 */
function blitTransparency(type: number, record: Uint8Array): BitmapTransparency | undefined {
    if (record.length < 44) return undefined;
    if (type === 114) {
        const pixelAlpha = (record[43] & 1) === 1, constantAlpha = record[42];
        return pixelAlpha || constantAlpha < 255 ? { ...(pixelAlpha ? { pixelAlpha } : {}), ...(constantAlpha < 255 ? { constantAlpha } : {}) } : undefined;
    }
    if (type === 116) return { transparentColor: [record[40], record[41], record[42]] };
    return undefined;
}

/** An EMF+ Image object's bitmap ([MS-EMFPLUS] 2.2.1.4): a compressed one's file, or one of pixels as a PNG; null for a metafile image or one not read. */
function emfPlusImage(data: Uint8Array): NonNullable<MetafileReading['bitmap']> | null {
    // Version, Type (1 bitmap), then Width, Height, Stride, PixelFormat, Type (0 pixels, 1 compressed), data.
    if (data.length < 28 || u32(data, 4) !== 1) return null;
    const width = i32(data, 8), height = i32(data, 12), stride = i32(data, 16), format = u32(data, 20), type = u32(data, 24);
    const body = data.subarray(28);
    if (type === 1) {
        const signature = pictureSignature(body);
        return signature === 'PNG' || signature === 'JPEG' || signature === 'GIF' || signature === 'TIFF' || signature === 'BMP' ? { format: signature, bytes: body.slice() } : null;
    }
    const pixel = EMFPLUS_PIXEL_FORMATS[format];
    if (type !== 0 || !pixel || !(width > 0 && height > 0) || width * height > MAX_RASTER_PIXELS || Math.abs(stride) < width * pixel.bytes || Math.abs(stride) * height > body.length) return null;
    const rgba = new Uint8Array(width * height * 4);
    for (let y = 0; y < height; y++) {
        // (A negative stride stores the rows from the foot up.)
        const row = stride > 0 ? y * stride : (height - 1 - y) * -stride;
        for (let x = 0; x < width; x++) {
            const p = row + x * pixel.bytes, o = (y * width + x) * 4;
            const a = pixel.alpha === 'none' ? 255 : body[p + 3];
            let r = body[p + 2], g = body[p + 1], b = body[p];
            if (pixel.alpha === 'premultiplied' && a > 0 && a < 255) { r = Math.min(255, Math.round(r * 255 / a)); g = Math.min(255, Math.round(g * 255 / a)); b = Math.min(255, Math.round(b * 255 / a)); }
            rgba[o] = r; rgba[o + 1] = g; rgba[o + 2] = b; rgba[o + 3] = a;
        }
    }
    const png = encodePng({ width, height, rgba });
    return png ? { format: 'PNG', bytes: png } : null;
}

/** A bitmap file's size in pixels from its header, as its picture is drawn (a TIFF's as its Orientation turns it). */
function rasterSize(bitmap: NonNullable<MetafileReading['bitmap']>): { width: number; height: number } | undefined {
    const b = bitmap.bytes;
    if (bitmap.format === 'PNG' && b.length >= 24) return { width: (b[16] << 24 | b[17] << 16 | b[18] << 8 | b[19]) >>> 0, height: (b[20] << 24 | b[21] << 16 | b[22] << 8 | b[23]) >>> 0 };
    if (bitmap.format === 'BMP' && b.length >= 26) return dibSize(b, 14);
    if (bitmap.format === 'GIF' && b.length >= 10) return { width: u16(b, 6), height: u16(b, 8) };
    if (bitmap.format === 'JPEG') { const size = sniffImageSize(b); return size ? { width: size.w, height: size.h } : undefined; }
    if (bitmap.format === 'TIFF') return tiffSize(b);
    return undefined;
}

/** A WMF's records walked (after its placeable header, if any): what they draw into `drawn`, their bitmaps into `found`; why it stopped early. */
function readWmf(bytes: Uint8Array, drawn: Set<string>, found: Found[]): string | undefined {
    const start = isPlaceable(bytes) ? PLACEABLE_HEADER_BYTES : 0;
    if (start + 18 > bytes.length) return 'its WMF header is damaged';
    const headerWords = u16(bytes, start + 2);
    if (headerWords < 9) return 'its WMF header is damaged';
    for (let at = start + headerWords * 2; at + 6 <= bytes.length;) {
        const words = u32(bytes, at), fn = u16(bytes, at + 4);
        if (fn === 0x0000) break; // META_EOF
        if (words < 3 || words > (bytes.length - at) / 2) return 'its WMF records are damaged';
        const record = bytes.subarray(at, at + words * 2);
        at += words * 2;
        const dibRecord = WMF_DIB_RECORDS[fn];
        if (dibRecord) {
            // (META_DIBSTRETCHBLT and META_DIBBITBLT with no bitmap fill a rectangle: their size is the parameters' words alone.)
            if ((fn === 0x0B41 || fn === 0x0940) && words === (fn >> 8) + 3) { drawn.add('shapes'); continue; }
            const file = dibAsFile(record, dibRecord.dib, undefined, undefined, undefined);
            if (!file) { drawn.add('other bitmaps'); continue; }
            const area = Math.abs(i16(record, dibRecord.dest[0]) * i16(record, dibRecord.dest[1]));
            const crop = dibCrop(dibRecord.src.map(o => i16(record, o)), dibSize(record, dibRecord.dib), dibRecord.own);
            found.push({ bitmap: crop ? { ...file, crop } : file, area });
            continue;
        }
        if (WMF_DDB_RECORDS.has(fn)) { drawn.add(words === (fn >> 8) + 3 ? 'shapes' : 'other bitmaps'); continue; }
        const what = WMF_DRAWS[fn];
        if (what) drawn.add(what);
    }
    return undefined;
}

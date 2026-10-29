/**
 * A picture the PDF stores as a JPEG, kept as one. pdf.js decodes every picture to pixels, and a page
 * scan that was a 1 MB JPEG came out as a 14 MB PNG, so every export of it was many times the file's
 * size. Where the JPEG needs no decoding to show (the viewer-independent cases: an 8-bit gray or RGB
 * picture with no mask, no `/Decode` and no colour transform to apply), the stream's own bytes are the
 * picture. Read from the file by the picture's object number (see pdfFile); whatever is not certain to
 * show the same as pdf.js's pixels (another filter, CMYK, an EXIF orientation a viewer applies, a header
 * that does not say what the dictionary says) is left to the PNG pdf.js's pixels make.
 *
 * @module parsers/pdf/pdfJpeg
 */

import { isStream, PdfDict, PdfFile, PdfName, PdfReader, PdfValue } from './pdfFile.js';

/** What a JPEG's header says of its picture. */
export interface JpegHeader {
    width: number;
    height: number;
    components: number;
    /** The sample precision in bits. */
    precision: number;
    /** The EXIF orientation (1 to 8) of the first EXIF segment that has one; undefined where there is none. */
    orientation?: number;
    /** An Adobe APP14 marker's colour transform (0 none, 1 YCbCr, 2 YCCK); undefined where there is no marker. */
    adobeTransform?: number;
}

/** The most JPEG bytes kept for a whole document: past it, a picture is the PNG pdf.js's pixels make. */
const MAX_KEPT_BYTES = 128 * 1024 * 1024;

/** Whether a marker has no length (and no segment): the start of the image, a restart, a temporary. */
const isStandalone = (marker: number): boolean => marker === 0x01 || marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7);

/**
 * Reads a JPEG's header: the frame (size, components, precision) and the markers that change how a
 * viewer shows it, up to the start of its scan, one marker segment at a time (linear in the header).
 * Undefined for bytes that are not a JPEG, a frame a browser does not decode (lossless, arithmetic-coded,
 * hierarchical), two frames, a header cut short, or a picture with no frame before its scan (or no scan).
 */
export function readJpegHeader(bytes: Uint8Array): JpegHeader | undefined {
    if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return undefined;
    const u16 = (at: number): number => (bytes[at] << 8) | bytes[at + 1];
    let header: JpegHeader | undefined;
    let orientation: number | undefined, adobeTransform: number | undefined;
    let at = 2, scan = false;
    while (at + 2 <= bytes.length) {
        // Any number of 0xff fill bytes may come before a marker.
        if (bytes[at] !== 0xff) return undefined;
        while (at + 1 < bytes.length && bytes[at + 1] === 0xff) at++;
        const marker = bytes[at + 1];
        if (marker === undefined || marker === 0x00) return undefined;
        at += 2;
        if (isStandalone(marker)) continue;
        if (marker === 0xd9) return undefined;
        if (at + 2 > bytes.length) return undefined;
        const length = u16(at);
        if (length < 2 || at + length > bytes.length) return undefined;
        const body = at + 2, end = at + length;
        if (marker === 0xda) { scan = true; break; }
        if (marker === 0xc0 || marker === 0xc1 || marker === 0xc2) {
            if (header || length < 8) return undefined;
            const components = bytes[body + 5];
            if (components < 1 || length < 8 + 3 * components) return undefined;
            header = { precision: bytes[body], height: u16(body + 1), width: u16(body + 3), components };
        } else if (marker >= 0xc3 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
            // Lossless, hierarchical or arithmetic-coded: not a picture a browser shows.
            return undefined;
        } else if (marker === 0xe1 && orientation === undefined) {
            orientation = exifOrientation(bytes, body, end);
        } else if (marker === 0xee && adobeTransform === undefined && end - body >= 12 && bytes[body] === 0x41 && bytes[body + 1] === 0x64 && bytes[body + 2] === 0x6f && bytes[body + 3] === 0x62 && bytes[body + 4] === 0x65) {
            adobeTransform = bytes[body + 11];
        }
        at = end;
    }
    if (!header || !scan) return undefined;
    if (orientation !== undefined) header.orientation = orientation;
    if (adobeTransform !== undefined) header.adobeTransform = adobeTransform;
    return header;
}

/** The orientation an EXIF segment (the bytes from `from` to `end`) holds in its first directory; undefined where it has none. */
function exifOrientation(bytes: Uint8Array, from: number, end: number): number | undefined {
    // "Exif" and two zero bytes, then a TIFF header: the byte order, 42, the first directory's offset.
    if (end - from < 14 || bytes[from] !== 0x45 || bytes[from + 1] !== 0x78 || bytes[from + 2] !== 0x69 || bytes[from + 3] !== 0x66 || bytes[from + 4] !== 0 || bytes[from + 5] !== 0) return undefined;
    const tiff = from + 6;
    const little = bytes[tiff] === 0x49 && bytes[tiff + 1] === 0x49;
    if (!little && !(bytes[tiff] === 0x4d && bytes[tiff + 1] === 0x4d)) return undefined;
    const u16 = (at: number): number => little ? bytes[at] | (bytes[at + 1] << 8) : (bytes[at] << 8) | bytes[at + 1];
    const u32 = (at: number): number => little
        ? (bytes[at] | (bytes[at + 1] << 8) | (bytes[at + 2] << 16) | (bytes[at + 3] << 24)) >>> 0
        : ((bytes[at] << 24) | (bytes[at + 1] << 16) | (bytes[at + 2] << 8) | bytes[at + 3]) >>> 0;
    if (u16(tiff + 2) !== 42) return undefined;
    const dir = tiff + u32(tiff + 4);
    if (dir + 2 > end) return undefined;
    const count = u16(dir);
    for (let i = 0; i < count; i++) {
        const entry = dir + 2 + i * 12;
        if (entry + 12 > end) return undefined;
        if (u16(entry) === 0x0112) return u16(entry + 8);
    }
    return undefined;
}

/**
 * The JPEGs a document keeps as they are stored, found in the file by the object number pdf.js gives
 * a picture, within the document's bound of bytes kept.
 */
export class KeptJpegs {
    private left = MAX_KEPT_BYTES;

    constructor(private readonly file: PdfFile) {}

    /**
     * A copy of the JPEG that a picture is, where its bytes show what pdf.js's pixels do: `ref` is the
     * picture's reference as pdf.js writes it ("4R", "12R1"), and `width` and `height` the pixels pdf.js
     * decoded. Undefined for a picture that has no reference (an inline one), is in a file whose
     * streams cannot be read here (encrypted, damaged), or is any other kind of picture.
     */
    keep(ref: unknown, width: number, height: number): Buffer | undefined {
        try {
            const match = typeof ref === 'string' ? /^(\d{1,9})R\d{0,9}$/.exec(ref) : null;
            const reader = match ? this.file.objects() : undefined;
            if (!match || !reader) return undefined;
            const stream = reader.get(Number(match[1]));
            if (!stream || !isStream(stream)) return undefined;
            const found = this.describe(reader, stream.dict);
            if (!found) return undefined;
            // The stream's bytes: what the dictionary, the header and pdf.js's pixels all say the same of.
            if (found.width !== width || found.height !== height) return undefined;
            // Up to its end-of-image marker: a stream whose length the file does not give is read to its
            // `endstream`, so the line end before it is not the picture's, and a JPEG cut short is no picture.
            const written = stream.src.subarray(stream.start, stream.start + stream.length);
            const bytes = written.subarray(0, endOfImage(written));
            if (bytes.length === 0 || bytes.length > this.left) return undefined;
            const header = readJpegHeader(bytes);
            if (!header || header.precision !== 8 || header.width !== width || header.height !== height || header.components !== found.components) return undefined;
            if (header.orientation !== undefined && header.orientation !== 1) return undefined;
            if (header.adobeTransform === 2) return undefined;
            this.left -= bytes.length;
            return Buffer.from(bytes);
        } catch {
            return undefined;
        }
    }

    /**
     * What a picture's dictionary says, where it is one that is shown as its JPEG bytes are: the only
     * filter DCTDecode with no colour transform of its own, 8 bits of DeviceGray, DeviceRGB or ICCBased
     * with one or three components, and nothing that changes the pixels (`/Decode`, a mask, a soft mask).
     */
    private describe(reader: PdfReader, dict: PdfDict): { width: number; height: number; components: number } | undefined {
        const get = (key: string): PdfValue | undefined => reader.resolve(dict.get(key));
        if ((get('Subtype') as PdfName | undefined)?.name !== 'Image') return undefined;
        for (const key of ['Decode', 'SMask', 'Mask', 'F', 'FFilter']) if (dict.has(key)) return undefined;
        if (get('ImageMask') === true) return undefined;
        const filter = get('Filter');
        const filters = (Array.isArray(filter) ? filter : [filter]).map(f => reader.resolve(f as PdfValue | undefined));
        if (filters.length !== 1 || !(filters[0] instanceof PdfName) || (filters[0].name !== 'DCTDecode' && filters[0].name !== 'DCT')) return undefined;
        const parms = get('DecodeParms');
        for (const p of Array.isArray(parms) ? parms : [parms]) {
            const one = reader.resolve(p as PdfValue | undefined);
            if (one instanceof Map && one.has('ColorTransform')) return undefined;
        }
        if (get('BitsPerComponent') !== 8) return undefined;
        const width = get('Width'), height = get('Height');
        if (typeof width !== 'number' || typeof height !== 'number') return undefined;
        const space = get('ColorSpace');
        let components: number | undefined;
        if (space instanceof PdfName) components = space.name === 'DeviceGray' ? 1 : space.name === 'DeviceRGB' ? 3 : undefined;
        else if (Array.isArray(space) && space.length === 2 && reader.resolve(space[0]) instanceof PdfName && (reader.resolve(space[0]) as PdfName).name === 'ICCBased') {
            const profile = reader.resolve(space[1]);
            const n = profile && isStream(profile) ? reader.resolve(profile.dict.get('N')) : undefined;
            components = n === 1 || n === 3 ? n : undefined;
        }
        return components ? { width, height, components } : undefined;
    }
}

/** Where a JPEG's bytes end, after its end-of-image marker (only white space or zero padding may follow it); 0 where they do not end with one. */
function endOfImage(bytes: Uint8Array): number {
    let end = bytes.length;
    while (end > 2 && bytes.length - end < 32 && (bytes[end - 1] === 0 || bytes[end - 1] === 0x0a || bytes[end - 1] === 0x0d || bytes[end - 1] === 0x20)) end--;
    return end >= 2 && bytes[end - 2] === 0xff && bytes[end - 1] === 0xd9 ? end : 0;
}

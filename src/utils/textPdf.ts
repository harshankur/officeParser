/**
 * Images carried inside a `.tex` file, as all-ASCII PDFs.
 *
 * LaTeX has no inline image syntax: `\includegraphics` reads a file. The kernel's `filecontents*`
 * environment writes a file while the document compiles, but only text, a line at a time. So the
 * LaTeX generator carries each PNG or JPEG as a one-page PDF whose every byte is printable ASCII
 * ({@link imageToTextPdf}): the image keeps its own compression behind an ASCII85 layer, and the
 * cross-reference table is a stream in ASCIIHex (a classic table's entries end in a space, which TeX
 * strips from every line it writes). Every engine includes a PDF: pdfTeX, LuaTeX and XeTeX directly,
 * and dvipdfmx for the DVI engines. The byte offsets hold because TeX writes the lines exactly as
 * given, ending each with LF (TeX Live on every platform, and MiKTeX since 2.9.7200).
 *
 * The LaTeX parser reads such a PDF back into the image it holds ({@link imageFromPdf}), which also
 * recovers the picture from any single-image PDF page a LaTeX project includes.
 */
import { unzlibSync, zlibSync } from 'fflate';
import { crc32 } from './imageUtils.js';

/** Largest image, in pixels, decoded to be re-encoded; a larger one is left as it is. */
const MAX_DECODED_PIXELS = 40_000_000;
/** Largest page content stream read when recognizing an image-only page. */
const MAX_CONTENT_STREAM_BYTES = 64 * 1024;
/**
 * Deflate expands at most about 1032:1, so data claiming more than this per compressed byte is
 * refused before any buffer for it is allocated.
 */
const MAX_INFLATE_RATIO = 1100;
/** Characters per line of encoded data. */
const LINE_LENGTH = 76;

/** FNV-1a over the bytes, as 8 hex digits: a short, stable name for identical content. */
export function contentHash(bytes: Uint8Array): string {
    let h = 0x811C9DC5;
    for (let i = 0; i < bytes.length; i++) h = Math.imul(h ^ bytes[i], 0x01000193);
    return (h >>> 0).toString(16).padStart(8, '0');
}

const u16 = (b: Uint8Array, i: number) => (b[i] << 8) | b[i + 1];
const u32 = (b: Uint8Array, i: number) => ((b[i] << 24) >>> 0) + (b[i + 1] << 16) + (b[i + 2] << 8) + b[i + 3];

/** Bytes as a string of the same code units (Latin-1), in chunks so a large input cannot overflow the call stack. */
function latin1(bytes: Uint8Array): string {
    let out = '';
    for (let i = 0; i < bytes.length; i += 0x8000) out += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    return out;
}

function concat(parts: Uint8Array[]): Uint8Array {
    const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
    let o = 0;
    for (const p of parts) { out.set(p, o); o += p.length; }
    return out;
}

// ── writing ─────────────────────────────────────────────────────────────────────────────────────

/** ASCII85 in lines of {@link LINE_LENGTH}, ending with the `~>` marker. */
function ascii85(bytes: Uint8Array): string {
    const out = new Uint8Array(Math.ceil(bytes.length / 4) * 5 + 2);
    let o = 0;
    const digits = [0, 0, 0, 0, 0];
    for (let i = 0; i < bytes.length; i += 4) {
        const rem = Math.min(4, bytes.length - i);
        let v = 0;
        for (let j = 0; j < 4; j++) v = v * 256 + (j < rem ? bytes[i + j] : 0);
        if (v === 0 && rem === 4) { out[o++] = 0x7A; continue; }
        for (let k = 4; k >= 0; k--) { digits[k] = v % 85; v = Math.floor(v / 85); }
        for (let k = 0; k <= rem; k++) out[o++] = 33 + digits[k];
    }
    out[o++] = 0x7E; out[o++] = 0x3E;
    const text = latin1(out.subarray(0, o));
    const lines: string[] = [];
    for (let i = 0; i < text.length; i += LINE_LENGTH) lines.push(text.slice(i, i + LINE_LENGTH));
    return lines.join('\n');
}

/** A number as PDF writes it: at most four decimals, no trailing zeros. */
function num(n: number): string {
    return String(Math.round(n * 10000) / 10000);
}

interface XObject { dict: string; data: Uint8Array; width: number; height: number }

interface JpegInfo { width: number; height: number; components: number; adobe: boolean }

/** Frame size and components of a JPEG PDF's DCTDecode can take (8-bit, Huffman-coded), or null. */
function jpegInfo(b: Uint8Array): JpegInfo | null {
    if (b.length < 4 || b[0] !== 0xFF || b[1] !== 0xD8) return null;
    let adobe = false;
    for (let i = 2; i + 3 < b.length;) {
        if (b[i] !== 0xFF) return null;
        const m = b[i + 1];
        if (m === 0xFF) { i++; continue; }
        if (m === 0x01 || (m >= 0xD0 && m <= 0xD8)) { i += 2; continue; }
        if (m === 0xD9 || m === 0xDA) return null;
        const len = u16(b, i + 2);
        if (len < 2 || i + 2 + len > b.length) return null;
        if (m === 0xEE && len >= 7 && latin1(b.subarray(i + 4, i + 9)) === 'Adobe') adobe = true;
        if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) {
            // Baseline, extended and progressive Huffman only: PDF readers do not all decode the rest.
            if (m > 0xC2 || len < 8) return null;
            const precision = b[i + 4], height = u16(b, i + 5), width = u16(b, i + 7), components = b[i + 9];
            if (precision !== 8 || !width || !height || ![1, 3, 4].includes(components)) return null;
            return { width, height, components, adobe };
        }
        i += 2 + len;
    }
    return null;
}

function jpegXObjects(bytes: Uint8Array): XObject[] | null {
    const info = jpegInfo(bytes);
    if (!info) return null;
    const cs = info.components === 1 ? '/DeviceGray' : info.components === 4 ? '/DeviceCMYK' : '/DeviceRGB';
    // Adobe writes CMYK JPEGs inverted, which readers undo only when told.
    const decode = info.components === 4 && info.adobe ? ' /Decode [1 0 1 0 1 0 1 0]' : '';
    return [{
        dict: `/Type /XObject /Subtype /Image /Width ${info.width} /Height ${info.height} /ColorSpace ${cs} /BitsPerComponent 8${decode} /Filter [/ASCII85Decode /DCTDecode]`,
        data: bytes, width: info.width, height: info.height,
    }];
}

interface PngInfo {
    width: number; height: number; depth: number; colorType: number; interlaced: boolean;
    palette: Uint8Array | null; trns: Uint8Array | null; idat: Uint8Array;
}

const PNG_SIGNATURE = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
const PNG_DEPTHS: Record<number, number[]> = { 0: [1, 2, 4, 8, 16], 2: [8, 16], 3: [1, 2, 4, 8], 4: [8, 16], 6: [8, 16] };
const PNG_CHANNELS: Record<number, number> = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 };
/** Adam7 passes: first column, first row, column step, row step. */
const ADAM7 = [[0, 0, 8, 8], [4, 0, 8, 8], [0, 4, 4, 8], [2, 0, 4, 4], [0, 2, 2, 4], [1, 0, 2, 2], [0, 1, 1, 2]];

function readPng(b: Uint8Array): PngInfo | null {
    if (b.length < 33 || PNG_SIGNATURE.some((v, i) => b[i] !== v)) return null;
    let ihdr: Uint8Array | null = null, palette: Uint8Array | null = null, trns: Uint8Array | null = null;
    const idat: Uint8Array[] = [];
    for (let i = 8; i + 12 <= b.length;) {
        const len = u32(b, i), type = latin1(b.subarray(i + 4, i + 8));
        if (i + 12 + len > b.length) return null;
        const data = b.subarray(i + 8, i + 8 + len);
        if (type === 'IHDR') ihdr = data;
        else if (type === 'PLTE') palette = data;
        else if (type === 'tRNS') trns = data;
        else if (type === 'IDAT') idat.push(data);
        else if (type === 'IEND') break;
        i += 12 + len;
    }
    if (!ihdr || ihdr.length !== 13 || idat.length === 0) return null;
    const width = u32(ihdr, 0), height = u32(ihdr, 4), depth = ihdr[8], colorType = ihdr[9];
    if (!width || !height || !PNG_DEPTHS[colorType]?.includes(depth) || ihdr[10] !== 0 || ihdr[11] !== 0 || ihdr[12] > 1) return null;
    if (colorType === 3 && (!palette || palette.length % 3 !== 0 || palette.length === 0 || palette.length > 768)) return null;
    return { width, height, depth, colorType, interlaced: ihdr[12] === 1, palette, trns, idat: concat(idat) };
}

function hexString(bytes: Uint8Array): string {
    return Array.from(bytes, v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** PNG's prediction for filter `f` from the byte to the left (`a`), above (`up`) and above-left (`c`). */
function predict(f: number, a: number, up: number, c: number): number {
    if (f === 1) return a;
    if (f === 2) return up;
    if (f === 3) return (a + up) >> 1;
    if (f === 4) { const q = a + up - c, pa = Math.abs(q - a), pb = Math.abs(q - up), pc = Math.abs(q - c); return pa <= pb && pa <= pc ? a : pb <= pc ? up : c; }
    return 0;
}

/** Undoes PNG row filters, row by row; `bpp` is bytes per complete pixel (at least 1). Null for an unknown filter. */
function unfilterRows(raw: Uint8Array, offset: number, rows: number, stride: number, bpp: number): Uint8Array | null {
    const out = new Uint8Array(rows * stride);
    for (let y = 0; y < rows; y++) {
        const f = raw[offset + y * (stride + 1)];
        if (f > 4) return null;
        const src = offset + y * (stride + 1) + 1, dst = y * stride, prev = dst - stride;
        for (let x = 0; x < stride; x++) {
            const a = x >= bpp ? out[dst + x - bpp] : 0;
            const up = y > 0 ? out[prev + x] : 0;
            const c = x >= bpp && y > 0 ? out[prev + x - bpp] : 0;
            out[dst + x] = (raw[src + x] + predict(f, a, up, c)) & 0xFF;
        }
    }
    return out;
}

/** The sample at `index` of a row packed at `bits` per sample. */
function sampleAt(row: Uint8Array, start: number, index: number, bits: number): number {
    if (bits === 8) return row[start + index];
    if (bits === 16) return (row[start + index * 2] << 8) | row[start + index * 2 + 1];
    const bit = index * bits, byte = row[start + (bit >> 3)];
    return (byte >> (8 - bits - (bit & 7))) & ((1 << bits) - 1);
}

interface Pixels { width: number; height: number; channels: 1 | 3; color: Uint8Array; alpha: Uint8Array | null }

/** A PNG decoded to 8-bit gray or RGB plus an alpha plane (null when fully opaque), or null. */
function decodePng(p: PngInfo): Pixels | null {
    const { width, height, depth, colorType } = p;
    if (width * height > MAX_DECODED_PIXELS) return null;
    const spp = PNG_CHANNELS[colorType], bitsPerPixel = spp * depth, bpp = Math.max(1, bitsPerPixel >> 3);
    const passes = p.interlaced ? ADAM7 : [[0, 0, 1, 1]];
    const sizes = passes.map(([x0, y0, dx, dy]) => {
        const pw = Math.max(0, Math.ceil((width - x0) / dx)), ph = Math.max(0, Math.ceil((height - y0) / dy));
        return { pw, ph, stride: Math.ceil(pw * bitsPerPixel / 8) };
    });
    const expected = sizes.reduce((n, s) => n + (s.pw && s.ph ? s.ph * (s.stride + 1) : 0), 0);
    if (expected > p.idat.length * MAX_INFLATE_RATIO + 1024) return null;
    let raw: Uint8Array;
    try { raw = unzlibSync(p.idat, { out: new Uint8Array(expected) }); } catch { return null; }
    if (raw.length < expected) return null;

    const channels: 1 | 3 = colorType === 0 || colorType === 4 ? 1 : 3;
    const color = new Uint8Array(width * height * channels);
    const alpha = new Uint8Array(width * height).fill(255);
    let hasAlpha = false;
    const maxSample = (1 << depth) - 1;
    const to8 = (v: number) => depth === 16 ? v >> 8 : depth === 8 ? v : Math.round(v * 255 / maxSample);
    const key = p.trns && colorType === 0 && p.trns.length >= 2 ? [u16(p.trns, 0)]
        : p.trns && colorType === 2 && p.trns.length >= 6 ? [u16(p.trns, 0), u16(p.trns, 2), u16(p.trns, 4)] : null;

    let offset = 0;
    for (let pass = 0; pass < passes.length; pass++) {
        const [x0, y0, dx, dy] = passes[pass];
        const { pw, ph, stride } = sizes[pass];
        if (!pw || !ph) continue;
        const rows = unfilterRows(raw, offset, ph, stride, bpp);
        if (!rows) return null;
        offset += ph * (stride + 1);
        for (let py = 0; py < ph; py++) {
            for (let px = 0; px < pw; px++) {
                const i = (y0 + py * dy) * width + (x0 + px * dx);
                const s = (k: number) => sampleAt(rows, py * stride, px * spp + k, depth);
                if (colorType === 3) {
                    const idx = s(0), pal = p.palette!;
                    if (idx * 3 + 2 < pal.length) color.set(pal.subarray(idx * 3, idx * 3 + 3), i * 3);
                    if (p.trns && idx < p.trns.length) alpha[i] = p.trns[idx];
                } else if (channels === 1) {
                    const g = s(0);
                    color[i] = to8(g);
                    if (colorType === 4) alpha[i] = to8(s(1));
                    else if (key && g === key[0]) alpha[i] = 0;
                } else {
                    const r = s(0), g = s(1), b = s(2);
                    color[i * 3] = to8(r); color[i * 3 + 1] = to8(g); color[i * 3 + 2] = to8(b);
                    if (colorType === 6) alpha[i] = to8(s(3));
                    else if (key && r === key[0] && g === key[1] && b === key[2]) alpha[i] = 0;
                }
                if (alpha[i] !== 255) hasAlpha = true;
            }
        }
    }
    return { width, height, channels, color, alpha: hasAlpha ? alpha : null };
}

/**
 * Rows prefixed with a PNG filter byte, each filter picked by the usual heuristic (least sum of
 * absolute differences), for PNG output and for PDF streams read with `/Predictor 15`.
 */
function filterRows(data: Uint8Array, width: number, height: number, channels: number): Uint8Array {
    const stride = width * channels, out = new Uint8Array(height * (stride + 1));
    const trial = new Uint8Array(stride);
    for (let y = 0; y < height; y++) {
        const row = y * stride, prev = row - stride;
        let best = 0, bestScore = Infinity;
        for (let f = 0; f <= 4; f++) {
            let score = 0;
            for (let x = 0; x < stride; x++) {
                const a = x >= channels ? data[row + x - channels] : 0;
                const up = y > 0 ? data[prev + x] : 0;
                const c = x >= channels && y > 0 ? data[prev + x - channels] : 0;
                const v = (data[row + x] - predict(f, a, up, c)) & 0xFF;
                score += v < 128 ? v : 256 - v;
            }
            if (score < bestScore) { bestScore = score; best = f; }
        }
        for (let x = 0; x < stride; x++) {
            const a = x >= channels ? data[row + x - channels] : 0;
            const up = y > 0 ? data[prev + x] : 0;
            const c = x >= channels && y > 0 ? data[prev + x - channels] : 0;
            trial[x] = (data[row + x] - predict(best, a, up, c)) & 0xFF;
        }
        out[y * (stride + 1)] = best;
        out.set(trial, y * (stride + 1) + 1);
    }
    return out;
}

/** An image stream compressed with PNG-style predictors, which PDF's `/Predictor 15` undoes. */
function flateImage(width: number, height: number, cs: string, data: Uint8Array, channels: number, extra = ''): XObject {
    return {
        dict: `/Type /XObject /Subtype /Image /Width ${width} /Height ${height} /ColorSpace ${cs} /BitsPerComponent 8${extra} /Filter [/ASCII85Decode /FlateDecode] /DecodeParms [null << /Predictor 15 /Colors ${channels} /BitsPerComponent 8 /Columns ${width} >>]`,
        data: zlibSync(filterRows(data, width, height, channels)), width, height,
    };
}

function pngXObjects(bytes: Uint8Array): XObject[] | null {
    const p = readPng(bytes);
    if (!p) return null;
    const { width, height, depth, colorType } = p;
    const paletteMasked = colorType === 3 && !!p.trns;
    if (!p.interlaced && (colorType === 0 || colorType === 2 || (colorType === 3 && !paletteMasked))) {
        // PDF reads PNG's own compressed rows (Predictor 15), so the data goes in unchanged.
        const channels = colorType === 2 ? 3 : 1;
        const cs = colorType === 0 ? '/DeviceGray' : colorType === 2 ? '/DeviceRGB'
            : `[/Indexed /DeviceRGB ${p.palette!.length / 3 - 1} <${hexString(p.palette!)}>]`;
        let mask = '';
        if (p.trns && colorType === 0 && p.trns.length >= 2) { const v = u16(p.trns, 0); mask = ` /Mask [${v} ${v}]`; }
        if (p.trns && colorType === 2 && p.trns.length >= 6) { const [r, g, b] = [u16(p.trns, 0), u16(p.trns, 2), u16(p.trns, 4)]; mask = ` /Mask [${r} ${r} ${g} ${g} ${b} ${b}]`; }
        return [{
            dict: `/Type /XObject /Subtype /Image /Width ${width} /Height ${height} /ColorSpace ${cs} /BitsPerComponent ${depth}${mask} /Filter [/ASCII85Decode /FlateDecode] /DecodeParms [null << /Predictor 15 /Colors ${channels} /BitsPerComponent ${depth} /Columns ${width} >>]`,
            data: p.idat, width, height,
        }];
    }
    // Transparency, a masked palette or interlacing: decode, and give the alpha its own soft mask.
    const px = decodePng(p);
    if (!px) return null;
    const cs = px.channels === 1 ? '/DeviceGray' : '/DeviceRGB';
    if (!px.alpha) return [flateImage(width, height, cs, px.color, px.channels)];
    return [flateImage(width, height, '/DeviceGray', px.alpha, 1), flateImage(width, height, cs, px.color, px.channels, ' /SMask {MASK}')];
}

/**
 * A PNG or JPEG as a one-page PDF of printable ASCII lines (none ending in a space): what a
 * `filecontents*` block can write for `\includegraphics` to read. The page is the image at
 * `pointsPerPixel` (0.75 draws it at 96 pixels per inch), and `bbox` is that page as a graphicx
 * `bb` value. Null for any other data, or an image too damaged or too large to carry.
 */
export function imageToTextPdf(bytes: Uint8Array, pointsPerPixel: number): { pdf: string; bbox: string } | null {
    let images: XObject[] | null;
    try { images = jpegXObjects(bytes) ?? pngXObjects(bytes); } catch { return null; }
    if (!images || !(pointsPerPixel > 0)) return null;
    const w = num(images[0].width * pointsPerPixel), h = num(images[0].height * pointsPerPixel);
    const imageNum = 4 + images.length;
    const content = `q ${w} 0 0 ${h} 0 0 cm /Im0 Do Q`;
    const bodies = [
        '<< /Type /Catalog /Pages 2 0 R >>',
        '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${w} ${h}] /Resources << /XObject << /Im0 ${imageNum} 0 R >> >> /Contents 4 0 R >>`,
        `<< /Length ${content.length} >>\nstream\n${content}\nendstream`,
        ...images.map(img => {
            const data = ascii85(img.data);
            return `<< ${img.dict.replace('{MASK}', `${imageNum - 1} 0 R`)} /Length ${data.length} >>\nstream\n${data}\nendstream`;
        }),
    ];
    let pdf = '%PDF-1.5\n';
    const offsets: number[] = [];
    bodies.forEach((body, i) => { offsets.push(pdf.length); pdf += `${i + 1} 0 obj\n${body}\nendobj\n`; });
    const xrefNum = bodies.length + 1;
    offsets.push(pdf.length);
    const hex = (n: number, bytes: number) => n.toString(16).toUpperCase().padStart(bytes * 2, '0');
    const rows = ['00' + hex(0, 4) + 'FFFF', ...offsets.map(o => '01' + hex(o, 4) + '0000')].join('\n') + '>';
    pdf += `${xrefNum} 0 obj\n<< /Type /XRef /Size ${xrefNum + 1} /W [1 4 2] /Root 1 0 R /Filter /ASCIIHexDecode /Length ${rows.length} >>\nstream\n${rows}\nendstream\nendobj\n`;
    pdf += `startxref\n${offsets[offsets.length - 1]}\n%%EOF`;
    return { pdf, bbox: `0 0 ${w} ${h}` };
}

// ── reading ─────────────────────────────────────────────────────────────────────────────────────

type PdfValue = number | boolean | null | PdfName | PdfRef | PdfString | PdfValue[] | PdfDict;
interface PdfName { name: string }
interface PdfRef { ref: number }
interface PdfString { str: Uint8Array }
interface PdfDict { dict: Map<string, PdfValue> }

const isName = (v: PdfValue | undefined): v is PdfName => !!v && typeof v === 'object' && 'name' in v;
const isRef = (v: PdfValue | undefined): v is PdfRef => !!v && typeof v === 'object' && 'ref' in v;
const isDict = (v: PdfValue | undefined): v is PdfDict => !!v && typeof v === 'object' && 'dict' in v;
const isStr = (v: PdfValue | undefined): v is PdfString => !!v && typeof v === 'object' && 'str' in v;

const WHITESPACE = new Set([0, 9, 10, 12, 13, 32]);
const DELIMITERS = new Set('()<>[]{}/%'.split('').map(c => c.charCodeAt(0)));

/** A minimal reader of PDF objects: enough to find a page, its content and the image it draws. */
class PdfReader {
    private objects = new Map<number, { value: PdfValue; stream: Uint8Array | null }>();

    constructor(private b: Uint8Array) {
        const text = latin1(b);
        const re = /(\d+)\s+(\d+)\s+obj\b/g;
        for (let m: RegExpExecArray | null; (m = re.exec(text));) {
            try {
                const parsed = this.value(m.index + m[0].length, 0);
                if (!parsed) continue;
                let stream: Uint8Array | null = null;
                let at = this.skip(parsed.end);
                if (text.startsWith('stream', at) && isDict(parsed.value)) {
                    at += 6;
                    if (b[at] === 13) at++;
                    if (b[at] === 10) at++;
                    const declared = parsed.value.dict.get('Length');
                    const end = typeof declared === 'number' && at + declared <= b.length && text.startsWith('endstream', this.skip(at + declared))
                        ? at + declared : text.indexOf('endstream', at);
                    if (end < 0) continue;
                    stream = b.subarray(at, end);
                    re.lastIndex = end;
                }
                this.objects.set(Number(m[1]), { value: parsed.value, stream });
            } catch { /* an unreadable object is skipped */ }
        }
    }

    get(ref: PdfValue | undefined): PdfValue | undefined {
        for (let i = 0; i < 8 && isRef(ref); i++) ref = this.objects.get(ref.ref)?.value;
        return ref;
    }

    stream(ref: PdfValue | undefined): { dict: Map<string, PdfValue>; data: Uint8Array } | null {
        if (!isRef(ref)) return null;
        const obj = this.objects.get(ref.ref);
        return obj?.stream && isDict(obj.value) ? { dict: obj.value.dict, data: obj.stream } : null;
    }

    all(): Array<{ value: PdfValue; stream: Uint8Array | null }> { return [...this.objects.values()]; }

    private skip(i: number): number {
        const b = this.b;
        while (i < b.length) {
            if (WHITESPACE.has(b[i])) i++;
            else if (b[i] === 37) { while (i < b.length && b[i] !== 10 && b[i] !== 13) i++; }
            else break;
        }
        return i;
    }

    private word(i: number): string {
        let j = i;
        while (j < this.b.length && !WHITESPACE.has(this.b[j]) && !DELIMITERS.has(this.b[j])) j++;
        return latin1(this.b.subarray(i, j));
    }

    private value(start: number, depth: number): { value: PdfValue; end: number } | null {
        if (depth > 32) return null;
        const b = this.b;
        let i = this.skip(start);
        if (i >= b.length) return null;
        const c = b[i];
        if (c === 0x3C && b[i + 1] === 0x3C) {
            const dict = new Map<string, PdfValue>();
            i += 2;
            for (;;) {
                i = this.skip(i);
                if (b[i] === 0x3E && b[i + 1] === 0x3E) return { value: { dict }, end: i + 2 };
                if (b[i] !== 0x2F) return null;
                const key = this.word(i + 1);
                const v = this.value(i + 1 + key.length, depth + 1);
                if (!v) return null;
                dict.set(key, v.value);
                i = v.end;
            }
        }
        if (c === 0x5B) {
            const arr: PdfValue[] = [];
            i++;
            for (;;) {
                i = this.skip(i);
                if (b[i] === 0x5D) return { value: arr, end: i + 1 };
                const v = this.value(i, depth + 1);
                if (!v) return null;
                arr.push(v.value);
                i = v.end;
            }
        }
        if (c === 0x2F) { const w = this.word(i + 1); return { value: { name: w }, end: i + 1 + w.length }; }
        if (c === 0x3C) {
            const close = b.indexOf(0x3E, i);
            if (close < 0) return null;
            return { value: { str: decodeAsciiHex(b.subarray(i + 1, close + 1)) }, end: close + 1 };
        }
        if (c === 0x28) {
            const out: number[] = [];
            let level = 1;
            for (i++; i < b.length && level > 0; i++) {
                const ch = b[i];
                if (ch === 0x5C) {
                    const n = b[++i];
                    const esc: Record<number, number> = { 0x6E: 10, 0x72: 13, 0x74: 9, 0x62: 8, 0x66: 12 };
                    if (n >= 0x30 && n <= 0x37) {
                        let v = 0, k = 0;
                        while (k < 3 && b[i] >= 0x30 && b[i] <= 0x37) { v = v * 8 + b[i] - 0x30; i++; k++; }
                        i--;
                        out.push(v & 0xFF);
                    } else if (n !== 10 && n !== 13) out.push(esc[n] ?? n);
                    continue;
                }
                if (ch === 0x28) level++;
                if (ch === 0x29 && --level === 0) break;
                out.push(ch);
            }
            return { value: { str: Uint8Array.from(out) }, end: i + 1 };
        }
        const w = this.word(i);
        if (!w) return null;
        if (w === 'true' || w === 'false') return { value: w === 'true', end: i + w.length };
        if (w === 'null') return { value: null, end: i + w.length };
        const n = Number(w);
        if (!Number.isFinite(n)) return null;
        // `N G R` is a reference.
        const after = this.skip(i + w.length);
        const gen = this.word(after);
        if (/^\d+$/.test(w) && /^\d+$/.test(gen)) {
            const r = this.skip(after + gen.length);
            if (b[r] === 0x52 && (r + 1 >= b.length || WHITESPACE.has(b[r + 1]) || DELIMITERS.has(b[r + 1]))) return { value: { ref: n }, end: r + 1 };
        }
        return { value: n, end: i + w.length };
    }
}

function decodeAsciiHex(b: Uint8Array): Uint8Array {
    const out: number[] = [];
    let hi = -1;
    for (const c of b) {
        if (c === 0x3E) break;
        const v = c >= 0x30 && c <= 0x39 ? c - 0x30 : c >= 0x41 && c <= 0x46 ? c - 55 : c >= 0x61 && c <= 0x66 ? c - 87 : -1;
        if (v < 0) continue;
        if (hi < 0) hi = v; else { out.push(hi * 16 + v); hi = -1; }
    }
    if (hi >= 0) out.push(hi * 16);
    return Uint8Array.from(out);
}

function decodeAscii85(b: Uint8Array): Uint8Array | null {
    let zeros = 0;
    for (const c of b) if (c === 0x7A) zeros++;
    const out = new Uint8Array(Math.ceil((b.length - zeros) * 4 / 5) + zeros * 4 + 4);
    let o = 0, n = 0, v = 0;
    let i = b[0] === 0x3C && b[1] === 0x7E ? 2 : 0;
    for (; i < b.length; i++) {
        const c = b[i];
        if (c === 0x7E) break;
        if (WHITESPACE.has(c)) continue;
        if (c === 0x7A && n === 0) { o += 4; continue; }
        if (c < 33 || c > 117) return null;
        v = v * 85 + (c - 33);
        if (++n === 5) {
            if (v > 0xFFFFFFFF) return null;
            out[o++] = v >>> 24; out[o++] = (v >>> 16) & 0xFF; out[o++] = (v >>> 8) & 0xFF; out[o++] = v & 0xFF;
            n = 0; v = 0;
        }
    }
    if (n === 1) return null;
    if (n > 0) {
        for (let k = n; k < 5; k++) v = v * 85 + 84;
        for (let k = 0; k < n - 1; k++) out[o++] = (v >>> (24 - 8 * k)) & 0xFF;
    }
    return out.subarray(0, o);
}

const asArray = (v: PdfValue | undefined): PdfValue[] => Array.isArray(v) ? v : v === undefined || v === null ? [] : [v];

/**
 * Applies a stream's ASCII filters, then Flate (inflating at most `maxBytes`). Returns the data
 * still behind a remaining image filter (DCTDecode) along with that filter's name.
 */
function decodeStream(reader: PdfReader, dict: Map<string, PdfValue>, data: Uint8Array, maxBytes: number): { data: Uint8Array; remaining: string | null; parms: PdfDict | null } | null {
    const filters = asArray(reader.get(dict.get('Filter'))).map(f => isName(f) ? f.name : '');
    const parmsList = asArray(reader.get(dict.get('DecodeParms')));
    let out: Uint8Array = data;
    for (let k = 0; k < filters.length; k++) {
        const f = filters[k];
        if (f === 'ASCIIHexDecode' || f === 'AHx') out = decodeAsciiHex(out);
        else if (f === 'ASCII85Decode' || f === 'A85') { const d = decodeAscii85(out); if (!d) return null; out = d; }
        else if (f === 'FlateDecode' || f === 'Fl') {
            if (k !== filters.length - 1) return null;
            const parms = reader.get(parmsList[k]);
            return { data: out, remaining: 'FlateDecode', parms: isDict(parms) ? parms : null };
        } else if (f === 'DCTDecode' || f === 'DCT') {
            return k === filters.length - 1 ? { data: out, remaining: 'DCTDecode', parms: null } : null;
        } else return null;
        if (out.length > maxBytes) return null;
    }
    return { data: out, remaining: null, parms: null };
}

function inflate(data: Uint8Array, size: number): Uint8Array | null {
    if (size > data.length * MAX_INFLATE_RATIO + 1024) return null;
    try {
        const out = unzlibSync(data, { out: new Uint8Array(size) });
        return out.length >= size ? out.subarray(0, size) : null;
    } catch { return null; }
}

function pngChunk(type: string, body: Uint8Array): Uint8Array {
    const out = new Uint8Array(12 + body.length);
    const view = new DataView(out.buffer);
    view.setUint32(0, body.length);
    for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
    out.set(body, 8);
    view.setUint32(8 + body.length, crc32(out.subarray(4, 8 + body.length)));
    return out;
}

function png(width: number, height: number, depth: number, colorType: number, idat: Uint8Array, extra: Uint8Array[] = []): Uint8Array {
    const ihdr = new Uint8Array(13);
    const view = new DataView(ihdr.buffer);
    view.setUint32(0, width); view.setUint32(4, height);
    ihdr[8] = depth; ihdr[9] = colorType;
    return concat([Uint8Array.from(PNG_SIGNATURE), pngChunk('IHDR', ihdr), ...extra, pngChunk('IDAT', idat), pngChunk('IEND', new Uint8Array(0))]);
}

interface ImageSpec { width: number; height: number; bpc: number; colors: number; palette: Uint8Array | null }

/** Width, height, bits per component and colour model of an image XObject PNG can hold, or null. */
function imageSpec(reader: PdfReader, dict: Map<string, PdfValue>): ImageSpec | null {
    const width = reader.get(dict.get('Width')), height = reader.get(dict.get('Height'));
    const bpc = reader.get(dict.get('BitsPerComponent')) ?? 8;
    if (typeof width !== 'number' || typeof height !== 'number' || typeof bpc !== 'number') return null;
    if (!(width > 0 && height > 0) || width * height > MAX_DECODED_PIXELS || ![1, 2, 4, 8, 16].includes(bpc)) return null;
    if (reader.get(dict.get('ImageMask')) === true || dict.has('Decode')) return null;
    const cs = reader.get(dict.get('ColorSpace'));
    const family = (v: PdfValue | undefined): number | null => {
        const n = isName(v) ? v.name : null;
        if (n === 'DeviceGray' || n === 'CalGray' || n === 'G') return 1;
        if (n === 'DeviceRGB' || n === 'CalRGB' || n === 'RGB') return 3;
        return null;
    };
    if (Array.isArray(cs) && isName(cs[0]) && (cs[0].name === 'Indexed' || cs[0].name === 'I')) {
        const hival = reader.get(cs[2]);
        const base = reader.get(cs[1]);
        const lookup = reader.get(cs[3]);
        const lookupStream = reader.stream(cs[3]);
        let table: Uint8Array | null = isStr(lookup) ? lookup.str : null;
        if (!table && lookupStream) {
            const d = decodeStream(reader, lookupStream.dict, lookupStream.data, 768);
            table = d?.remaining === 'FlateDecode' ? inflate(d.data, 768) : d?.remaining === null ? d.data : null;
        }
        if (typeof hival !== 'number' || family(base) !== 3 || !table || bpc > 8 || table.length < (hival + 1) * 3) return null;
        return { width, height, bpc, colors: 1, palette: table.subarray(0, (hival + 1) * 3) };
    }
    if (Array.isArray(cs) && isName(cs[0]) && cs[0].name === 'ICCBased') {
        const icc = reader.stream(cs[1]);
        const n = icc ? reader.get(icc.dict.get('N')) : undefined;
        return n === 1 || n === 3 ? { width, height, bpc, colors: n, palette: null } : null;
    }
    const colors = family(cs);
    return colors ? { width, height, bpc, colors, palette: null } : null;
}

/** An image XObject's samples as packed rows (predictors undone), or null. */
function imageRows(reader: PdfReader, dict: Map<string, PdfValue>, data: Uint8Array, spec: ImageSpec): Uint8Array | null {
    const stride = Math.ceil(spec.width * spec.colors * spec.bpc / 8);
    const size = stride * spec.height;
    const d = decodeStream(reader, dict, data, size * 2 + 1024);
    if (!d || d.remaining === 'DCTDecode') return null;
    const predictor = d.parms ? reader.get(d.parms.dict.get('Predictor')) : undefined;
    if (d.remaining !== 'FlateDecode') return d.data.length >= size ? d.data.subarray(0, size) : null;
    if (typeof predictor === 'number' && predictor >= 10) {
        const raw = inflate(d.data, size + spec.height);
        return raw && unfilterRows(raw, 0, spec.height, stride, Math.max(1, (spec.colors * spec.bpc) >> 3));
    }
    if (typeof predictor === 'number' && predictor > 1) return null;
    return inflate(d.data, size);
}

/** The image a PDF's page draws, when the page draws nothing else. */
function pageImage(reader: PdfReader): { dict: Map<string, PdfValue>; data: Uint8Array } | null {
    const pages = reader.all().filter(o => isDict(o.value) && isName(o.value.dict.get('Type')) && (o.value.dict.get('Type') as PdfName).name === 'Page');
    if (pages.length !== 1) return null;
    const page = (pages[0].value as PdfDict).dict;
    let resources = reader.get(page.get('Resources'));
    let parent = reader.get(page.get('Parent'));
    for (let i = 0; !resources && isDict(parent) && i < 16; i++) {
        resources = reader.get(parent.dict.get('Resources'));
        parent = reader.get(parent.dict.get('Parent'));
    }
    const xobjects = isDict(resources) ? reader.get(resources.dict.get('XObject')) : undefined;
    if (!isDict(xobjects)) return null;

    let content = '';
    const contents = page.get('Contents');
    for (const ref of isRef(contents) && reader.stream(contents) ? [contents] : asArray(reader.get(contents))) {
        const s = reader.stream(ref);
        if (!s) return null;
        const d = decodeStream(reader, s.dict, s.data, MAX_CONTENT_STREAM_BYTES);
        if (!d || d.remaining === 'DCTDecode') return null;
        const bytes = d.remaining === 'FlateDecode' ? (() => { try { return unzlibSync(d.data, { out: new Uint8Array(MAX_CONTENT_STREAM_BYTES) }); } catch { return null; } })() : d.data;
        if (!bytes) return null;
        content += latin1(bytes) + '\n';
    }
    // Only placement and one image: anything else drawn means the page is more than the picture.
    const tokens = content.split(/[\s]+/).filter(Boolean);
    const draws = tokens.filter(t => t === 'Do').length;
    if (draws !== 1 || tokens.some(t => !/^(?:q|Q|cm|Do|-?[\d.]+|\/[^\s/]+)$/.test(t))) return null;
    const name = tokens[tokens.indexOf('Do') - 1];
    if (!name?.startsWith('/')) return null;
    const ref = xobjects.dict.get(name.slice(1));
    const s = reader.stream(ref);
    if (!s || !isName(s.dict.get('Subtype')) || (s.dict.get('Subtype') as PdfName).name !== 'Image') return null;
    return s;
}

/**
 * The picture a single-page PDF shows, as PNG or JPEG bytes: the page must draw one image and
 * nothing else (as {@link imageToTextPdf} writes, and as an image saved as PDF usually is). A JPEG
 * comes back byte for byte; other images become a PNG with the same pixels, transparency included.
 * Null for any other PDF (a vector figure, several pages, an unsupported encoding), which stays a PDF.
 */
export function imageFromPdf(bytes: Uint8Array): { data: Uint8Array; mimeType: 'image/png' | 'image/jpeg' } | null {
    try {
        if (latin1(bytes.subarray(0, 1024)).indexOf('%PDF-') < 0) return null;
        const reader = new PdfReader(bytes);
        const img = pageImage(reader);
        if (!img) return null;
        const d = decodeStream(reader, img.dict, img.data, MAX_DECODED_PIXELS * 4);
        if (!d) return null;
        if (d.remaining === 'DCTDecode') return jpegInfo(d.data) ? { data: d.data, mimeType: 'image/jpeg' } : null;

        const spec = imageSpec(reader, img.dict);
        if (!spec) return null;
        const smask = reader.stream(img.dict.get('SMask'));
        const colorKey = reader.get(img.dict.get('Mask'));
        const predictor = d.parms ? reader.get(d.parms.dict.get('Predictor')) : undefined;
        const colorType = spec.palette ? 3 : spec.colors === 1 ? 0 : 2;
        const plte = spec.palette ? [pngChunk('PLTE', spec.palette)] : [];

        // PNG's own compressed rows, as the generator passes them through: they become the PNG as they are.
        const PARM_DEFAULTS: Record<string, number> = { Colors: 1, BitsPerComponent: 8, Columns: 1 };
        const parmsMatch = (key: string, expected: number) => (d.parms ? reader.get(d.parms.dict.get(key)) ?? PARM_DEFAULTS[key] : PARM_DEFAULTS[key]) === expected;
        if (!smask && colorKey === undefined && d.remaining === 'FlateDecode' && predictor === 15
            && parmsMatch('Colors', spec.colors) && parmsMatch('BitsPerComponent', spec.bpc) && parmsMatch('Columns', spec.width)) {
            return { data: png(spec.width, spec.height, spec.bpc, colorType, d.data, plte), mimeType: 'image/png' };
        }

        const rows = imageRows(reader, img.dict, img.data, spec);
        if (!rows) return null;
        const stride = Math.ceil(spec.width * spec.colors * spec.bpc / 8);
        const n = spec.width * spec.height;
        const channels = spec.palette ? 3 : spec.colors;
        const color = new Uint8Array(n * channels);
        const alpha = new Uint8Array(n).fill(255);
        const max = (1 << spec.bpc) - 1;
        const keys = Array.isArray(colorKey) && colorKey.every(k => typeof k === 'number') ? colorKey as number[] : null;
        let hasAlpha = false;
        for (let y = 0; y < spec.height; y++) {
            for (let x = 0; x < spec.width; x++) {
                const i = y * spec.width + x;
                const raw = Array.from({ length: spec.colors }, (_, k) => sampleAt(rows, y * stride, x * spec.colors + k, spec.bpc));
                if (spec.palette) color.set(spec.palette.subarray(raw[0] * 3, raw[0] * 3 + 3), i * 3);
                else raw.forEach((v, k) => { color[i * channels + k] = spec.bpc === 16 ? v >> 8 : spec.bpc === 8 ? v : Math.round(v * 255 / max); });
                if (keys && raw.every((v, k) => v >= keys[k * 2] && v <= keys[k * 2 + 1])) { alpha[i] = 0; hasAlpha = true; }
            }
        }
        if (smask) {
            const maskSpec = imageSpec(reader, smask.dict);
            if (!maskSpec || maskSpec.colors !== 1 || maskSpec.width !== spec.width || maskSpec.height !== spec.height) return null;
            const maskRows = imageRows(reader, smask.dict, smask.data, maskSpec);
            if (!maskRows) return null;
            const maskStride = Math.ceil(maskSpec.width * maskSpec.bpc / 8), maskMax = (1 << maskSpec.bpc) - 1;
            for (let y = 0; y < spec.height; y++) {
                for (let x = 0; x < spec.width; x++) {
                    const v = sampleAt(maskRows, y * maskStride, x, maskSpec.bpc);
                    const a = maskSpec.bpc === 16 ? v >> 8 : maskSpec.bpc === 8 ? v : Math.round(v * 255 / maskMax);
                    alpha[y * spec.width + x] = Math.min(alpha[y * spec.width + x], a);
                    if (a !== 255) hasAlpha = true;
                }
            }
        }
        if (!hasAlpha) {
            return { data: png(spec.width, spec.height, 8, channels === 1 ? 0 : 2, zlibSync(filterRows(color, spec.width, spec.height, channels))), mimeType: 'image/png' };
        }
        const withAlpha = new Uint8Array(n * (channels + 1));
        for (let i = 0; i < n; i++) {
            withAlpha.set(color.subarray(i * channels, (i + 1) * channels), i * (channels + 1));
            withAlpha[i * (channels + 1) + channels] = alpha[i];
        }
        return { data: png(spec.width, spec.height, 8, channels === 1 ? 4 : 6, zlibSync(filterRows(withAlpha, spec.width, spec.height, channels + 1))), mimeType: 'image/png' };
    } catch {
        return null;
    }
}

/**
 * Pictures on their way to output: data that is not base64, a picture's format read from its bytes, TIFF pictures
 * refused by the native PDF engine and RTF, metafile bitmaps in the native PDF engine, pictures drawn in a heading's, list item's or note's line, equations
 * written as their LaTeX source reported, and ODF chart points that are blank.
 */
import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import * as zlib from 'zlib';
import { strToU8, unzlibSync, zipSync, zlibSync } from 'fflate';
import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFRawStream, PDFString } from 'pdf-lib';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { RtfGenerator } from '../../src/generators/RtfGenerator';
import { OfficeParser } from '../../src/OfficeParser';
import type { OfficeContentNode, OfficeParserAST } from '../../src/types';
import { decodeBmp, decodeGif, encodePng } from '../../src/utils/rasterUtils';
import { readMetafile } from '../../src/utils/metafileUtils';
import { withinTime } from '../helpers/timing';
import { allNodes, bytesOf, docxOf, emfBytes, lcg, para, pictureRun, placeableWmf, run, zipTexts } from './wordRtfContent';

// ─── Shared helpers ──────────────────────────────────────────────────────────

/** Every output format, the PDF one through the native engine. */
const EVERY_FORMAT = ['text', 'md', 'html', 'pdf', 'csv', 'rtf', 'chunks', 'epub', 'docx', 'odt', 'tex', 'xlsx', 'ods', 'pptx', 'odp'] as const;

/** An AST of `content` and `attachments`. */
export const astOf = (content: OfficeContentNode[], attachments: OfficeParserAST['attachments'] = []): OfficeParserAST =>
    ({ type: 'docx', metadata: {}, content, attachments } as unknown as OfficeParserAST);

/** Writes `ast` as `format` (PDF through the native engine), collecting its warnings. */
export const writeWithWarnings = async (ast: OfficeParserAST, format: string, config: Record<string, unknown> = {}) => {
    const warnings: any[] = [];
    const pdfConfig = { engine: 'native', ...(config.pdfConfig as object | undefined) };
    const value = (await OfficeGenerator.generate(ast, format as any, { ...config, ...(format === 'pdf' ? { pdfConfig } : {}), onWarning: (w: any) => warnings.push(w) } as any)).value as any;
    return { value, warnings };
};

// ─── Data that is not base64 ─────────────────────────────────────────────

/**
 * A picture whose attachment's data is not base64, shown twice (once with alternative text), to every output
 * format: none throws; PPTX, ODP, RTF, DOCX, ODT and LaTeX report it as data that could not be decoded, and the
 * decks and RTF write its alternative text where it stood.
 */
export async function testUndecodablePictures(): Promise<void> {
    console.log('\n=== Pictures whose data is not base64 ===');
    const ast = astOf([
        { type: 'paragraph', children: [{ type: 'text', text: 'Before the picture.' }] },
        { type: 'paragraph', children: [{ type: 'image', text: '', metadata: { attachmentName: 'bad.png' } }] },
        { type: 'paragraph', children: [{ type: 'image', text: '', metadata: { attachmentName: 'bad.png', altText: 'ALTWORDS' } }] },
    ] as OfficeContentNode[], [{ type: 'image', name: 'bad.png', extension: 'png', mimeType: 'image/png', data: '!!!not base64!!!' }] as any);
    const undecodable = (w: any) => w.code === 'IMAGE_PROCESSING_FAILED' && w.details?.name === 'bad.png' && w.details?.reason === 'its data could not be decoded';
    for (const format of EVERY_FORMAT) {
        let result: Awaited<ReturnType<typeof writeWithWarnings>>;
        try { result = await writeWithWarnings(ast, format); } catch (err: any) { throw new Error(`Undecodable picture: ${format} output threw (${err?.message ?? err})`); }
        if (['pptx', 'odp', 'rtf', 'docx', 'odt', 'tex'].includes(format)) {
            assert.ok(result.warnings.some(undecodable), `Undecodable picture: ${format} reports its data could not be decoded (${JSON.stringify(result.warnings.map(w => w.details))})`);
        }
        if (format === 'pptx') {
            const files = zipTexts(result.value);
            const slides = Object.keys(files).filter(k => /^ppt\/slides\/[^/]+\.xml$/.test(k)).map(k => files[k]).join('');
            assert.ok(/<a:t>ALTWORDS<\/a:t>/.test(slides), 'Undecodable picture: PPTX writes its alt text in a text box');
            assert.ok(!Object.keys(files).some(k => k.startsWith('ppt/media/')), 'Undecodable picture: PPTX packages no media for it');
        }
        if (format === 'odp') {
            const files = zipTexts(result.value);
            assert.ok(/<draw:text-box>.*ALTWORDS.*<\/draw:text-box>/s.test(files['content.xml']), 'Undecodable picture: ODP writes its alt text in a text box');
        }
        if (format === 'rtf') assert.ok(String(result.value).includes('ALTWORDS') && !String(result.value).includes('\\pict'), 'Undecodable picture: RTF writes its alt text in the text, and no picture');
    }
    console.log(`✓ Undecodable picture: ${EVERY_FORMAT.length} formats, none throws; the six that embed pictures report it`);
}

// ─── Pictures embedded by what their bytes are ───────────────────────────

/** The image XObjects a PDF holds (their soft masks left out), by size, and how many form XObjects. */
export const pdfPictures = async (value: unknown): Promise<{ images: { w: number; h: number }[]; forms: number }> => {
    const doc = await PDFDocument.load(bytesOf(value));
    const masks = new Set<string>();
    const images: { ref: string; w: number; h: number }[] = [];
    let forms = 0;
    for (const [ref, object] of doc.context.enumerateIndirectObjects()) {
        if (!(object instanceof PDFRawStream)) continue;
        const subtype = object.dict.get(PDFName.of('Subtype'));
        if (subtype === PDFName.of('Form')) forms++;
        if (subtype !== PDFName.of('Image')) continue;
        const mask = object.dict.get(PDFName.of('SMask'));
        if (mask) masks.add(mask.toString());
        images.push({ ref: ref.toString(), w: Number(object.dict.get(PDFName.of('Width'))?.toString()), h: Number(object.dict.get(PDFName.of('Height'))?.toString()) });
    }
    return { images: images.filter(i => !masks.has(i.ref)).map(({ w, h }) => ({ w, h })), forms };
};

/** A PNG of `w` by `h` pixels, each `(x, y)` coloured by `color`. */
export const pngOf = (w: number, h: number, color: (x: number, y: number) => [number, number, number, number] = (x, y) => [x * 7 & 255, y * 11 & 255, (x + y) & 255, 255]): Uint8Array => {
    const rgba = new Uint8Array(w * h * 4);
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) rgba.set(color(x, y), (y * w + x) * 4);
    return encodePng({ width: w, height: h, rgba })!;
};

/** An uncompressed 24-bit BMP of `w` by `h` pixels, coloured as {@link pngOf} colours them. */
export const bmpOf = (w: number, h: number): Uint8Array => {
    const stride = Math.ceil(w * 3 / 4) * 4;
    const out = new Uint8Array(54 + stride * h);
    const view = new DataView(out.buffer);
    out[0] = 0x42; out[1] = 0x4D;
    view.setUint32(2, out.length, true); view.setUint32(10, 54, true); view.setUint32(14, 40, true);
    view.setInt32(18, w, true); view.setInt32(22, h, true); view.setUint16(26, 1, true); view.setUint16(28, 24, true);
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
        const at = 54 + (h - 1 - y) * stride + x * 3;
        out[at] = (x + y) & 255; out[at + 1] = y * 11 & 255; out[at + 2] = x * 7 & 255;
    }
    return out;
};

/** A PNG's pixels as RGBA (8-bit grey, RGB or RGBA, any filter), for a picture read back. */
export const pngPixels = (png: Uint8Array): { width: number; height: number; rgba: Uint8Array } => {
    const view = new DataView(png.buffer, png.byteOffset, png.byteLength);
    const width = view.getUint32(16), height = view.getUint32(20), channels = png[25] === 6 ? 4 : png[25] === 2 ? 3 : 1;
    const parts: Uint8Array[] = [];
    for (let at = 8; at < png.length;) {
        const length = view.getUint32(at), type = String.fromCharCode(...png.subarray(at + 4, at + 8));
        if (type === 'IDAT') parts.push(png.subarray(at + 8, at + 8 + length));
        at += 12 + length;
    }
    const joined = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
    parts.reduce((at, p) => (joined.set(p, at), at + p.length), 0);
    const raw = unzlibSync(joined), stride = width * channels, rgba = new Uint8Array(width * height * 4);
    let previous = new Uint8Array(stride);
    for (let y = 0; y < height; y++) {
        const filter = raw[y * (stride + 1)], line = raw.slice(y * (stride + 1) + 1, (y + 1) * (stride + 1));
        for (let i = 0; i < stride; i++) {
            const a = i >= channels ? line[i - channels] : 0, b = previous[i], c = i >= channels ? previous[i - channels] : 0;
            const p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c);
            line[i] = (line[i] + (filter === 1 ? a : filter === 2 ? b : filter === 3 ? (a + b) >> 1 : filter === 4 ? (pa <= pb && pa <= pc ? a : pb <= pc ? b : c) : 0)) & 255;
        }
        for (let x = 0; x < width; x++) {
            const px = channels === 4 ? Array.from(line.subarray(x * 4, x * 4 + 4)) : channels === 3 ? [...line.subarray(x * 3, x * 3 + 3), 255] : [line[x], line[x], line[x], 255];
            rgba.set(px, (y * width + x) * 4);
        }
        previous = line;
    }
    return { width, height, rgba };
};

/** A JPEG's markers to its frame header (192 by 48 pixels), which is all a PDF needs to embed one. */
export const JPEG_192x48 = new Uint8Array([0xFF, 0xD8, 0xFF, 0xE0, 0, 16, 0x4A, 0x46, 0x49, 0x46, 0, 1, 1, 0, 0, 1, 0, 1, 0, 0, 0xFF, 0xC0, 0, 17, 8, 0, 48, 0, 192, 3, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1, 0xFF, 0xD9]);
/** A 1x1 GIF. */
export const GIF_1x1 = new Uint8Array(Buffer.from('R0lGODlhAQABAIAAAP///wAAACH5BAEAAAAALAAAAAABAAEAAAICRAEAOw==', 'base64'));
/** A WebP's RIFF header (its picture data left out). */
const WEBP = new Uint8Array([0x52, 0x49, 0x46, 0x46, 26, 0, 0, 0, 0x57, 0x45, 0x42, 0x50, 0x56, 0x50, 0x38, 0x4C, ...new Array(18).fill(0)]);
/** An SVG of 40 by 20 pixels. */
const SVG_40x20 = strToU8('<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"><rect width="40" height="20" fill="#3060c0"/></svg>');

/** One paragraph holding a picture of `bytes`, its attachment named `name` and typed `mimeType`. */
export const onePicture = (bytes: Uint8Array, name: string, mimeType: string, altText?: string): OfficeParserAST => astOf(
    [{ type: 'paragraph', children: [{ type: 'image', text: '', metadata: { attachmentName: name, ...(altText ? { altText } : {}) } }] }] as OfficeContentNode[],
    [{ type: 'image', name, extension: name.includes('.') ? name.split('.').pop() : undefined, mimeType, data: Buffer.from(bytes).toString('base64') }] as any);

/**
 * Native PDF output embeds a picture by its bytes, whatever its name or type says: a JPEG named and typed as a
 * PNG, a PNG typed JPEG or named `.svg`, an SVG typed PNG, a GIF named `.jpg` and a BMP typed PNG are each drawn
 * at their size, with no warning. Bytes that are no picture, XML, a WebP and a damaged PNG are reported by what
 * they are.
 */
export async function testNativePdfPictureSignatures(): Promise<void> {
    console.log('\n=== Native PDF: pictures embedded by their bytes ===');
    const png = pngOf(40, 20);
    const drawn: [string, Uint8Array, string, string, { w: number; h: number } | 'form'][] = [
        ['a JPEG named and typed PNG', JPEG_192x48, 'photo.png', 'image/png', { w: 192, h: 48 }],
        ['a PNG typed JPEG', png, 'chart.jpg', 'image/jpeg', { w: 40, h: 20 }],
        ['a PNG named .svg', png, 'logo.svg', 'image/svg+xml', { w: 40, h: 20 }],
        ['an SVG typed PNG', SVG_40x20, 'logo.png', 'image/png', 'form'],
        ['a GIF named .jpg', GIF_1x1, 'dot.jpg', 'image/jpeg', { w: 1, h: 1 }],
        ['a BMP typed PNG', bmpOf(5, 3), 'tile.png', 'image/png', { w: 5, h: 3 }],
    ];
    for (const [label, bytes, name, mimeType, size] of drawn) {
        const { value, warnings } = await writeWithWarnings(onePicture(bytes, name, mimeType), 'pdf');
        assert.deepStrictEqual(warnings.map(w => w.details), [], `Picture signatures: ${label} warns nothing`);
        const { images, forms } = await pdfPictures(value);
        if (size === 'form') assert.ok(images.length === 0 && forms >= 1, `Picture signatures: ${label} is drawn as a form (${images.length} images, ${forms} forms)`);
        else assert.deepStrictEqual(images, [size], `Picture signatures: ${label} is drawn at its size`);
    }
    const reasonOf = async (bytes: Uint8Array, name: string, mimeType: string) => {
        const { warnings } = await writeWithWarnings(onePicture(bytes, name, mimeType, 'ALT'), 'pdf');
        return warnings.filter(w => w.code === 'IMAGE_PROCESSING_FAILED').map(w => w.details?.reason);
    };
    const noise = new Uint8Array(64).map((_, i) => (i * 37 + 11) & 255);
    assert.deepStrictEqual(await reasonOf(noise, 'noise.png', 'image/png'), ['its data is not a picture of a format the native PDF engine knows (its type or name says PNG)'], 'Picture signatures: bytes that are no picture, named by their type');
    assert.deepStrictEqual(await reasonOf(noise, 'noise', ''), ['the native PDF engine does not embed pictures of its kind (only PNG, JPEG, GIF, BMP, SVG, and bitmaps in EMF and WMF)'], 'Picture signatures: bytes that are no picture, nameless');
    assert.deepStrictEqual(await reasonOf(strToU8('<root><a/></root>'), 'x.png', 'image/png'), ['its data is XML, not a picture'], 'Picture signatures: XML');
    assert.deepStrictEqual(await reasonOf(WEBP, 'a.png', 'image/png'), ['the native PDF engine does not embed WebP pictures (only PNG, JPEG, GIF, BMP, SVG, and bitmaps in EMF and WMF)'], 'Picture signatures: a WebP typed PNG, named as WebP');
    assert.deepStrictEqual(await reasonOf(png.slice(0, 40), 'b.jpg', 'image/jpeg'), ['its PNG data could not be read'], 'Picture signatures: a damaged PNG typed JPEG, named as PNG');
    console.log(`  ✓ ${drawn.length} pictures labelled as another format drawn by their bytes; five refusals named by what the bytes are`);
}

/**
 * RTF output writes each picture by its bytes: a GIF and a BMP as a PNG of their pixels (read back by the RTF
 * reader to the same pixels), an EMF as `\emfblip`, a WMF as `\wmetafile8`, a JPEG typed PNG as `\jpegblip`; an
 * SVG, and bytes that are no picture, are their alt text and reported.
 */
export async function testRtfPictureSignatures(): Promise<void> {
    console.log('\n=== RTF output: pictures by their bytes ===');
    const rtfOf = async (bytes: Uint8Array, name: string, mimeType: string) => {
        const { value, warnings } = await writeWithWarnings(onePicture(bytes, name, mimeType, 'ALTWORDS'), 'rtf');
        return { rtf: String(value), warnings };
    };
    const readBack = async (rtf: string) => (await OfficeParser.parseOffice(Buffer.from(rtf, 'latin1'), { extractAttachments: true })).attachments;
    const tiffSpec: TiffSpec = { width: 6, height: 4, photometric: 2, bits: 8, compression: 5, little: true, samples: Array.from({ length: 72 }, (_, i) => i * 3) };
    const cases: [string, Uint8Array, string, string, number, number][] = [['GIF', GIF_1x1, 'a.gif', 'image/gif', 1, 1], ['BMP', bmpOf(7, 5), 'a.bmp', 'image/bmp', 7, 5], ['BMP typed JPEG', bmpOf(3, 2), 'a.jpg', 'image/jpeg', 3, 2]];
    for (const [label, bytes, name, mime, w, h] of cases) {
        const { rtf, warnings } = await rtfOf(bytes, name, mime);
        assert.ok(rtf.includes(`{\\pict\\pngblip\\picw${w}\\pich${h}`) && !rtf.includes('ALTWORDS'), `RTF signatures: a ${label} is written as a PNG`);
        assert.deepStrictEqual(warnings, [], `RTF signatures: a ${label} warns nothing`);
        const [back] = await readBack(rtf);
        const pixels = pngPixels(new Uint8Array(Buffer.from(back.data, 'base64')));
        const source = label === 'GIF' ? decodeGif(GIF_1x1, 100)! : pngPixels(pngOf(w, h));
        assert.deepStrictEqual([pixels.width, pixels.height, Buffer.from(pixels.rgba).equals(Buffer.from(source.rgba))], [w, h, true], `RTF signatures: the ${label} read back holds its pixels`);
    }
    assert.ok((await rtfOf(JPEG_192x48, 'a.png', 'image/png')).rtf.includes('{\\pict\\jpegblip\\picw192\\pich48'), 'RTF signatures: a JPEG typed PNG is \\jpegblip');
    assert.ok((await rtfOf(emfBytes(), 'a.png', 'image/png')).rtf.includes('{\\pict\\emfblip'), 'RTF signatures: an EMF typed PNG is \\emfblip');
    assert.ok((await rtfOf(placeableWmf({ left: 0, top: 0, right: 1440, bottom: 720 }, 1440), 'a.bin', 'application/octet-stream')).rtf.includes('{\\pict\\wmetafile8'), 'RTF signatures: a WMF named .bin is \\wmetafile8');
    const svg = await rtfOf(SVG_40x20, 'a.svg', 'image/svg+xml');
    assert.ok(!svg.rtf.includes('\\pict') && svg.rtf.includes('ALTWORDS'), 'RTF signatures: an SVG is its alt text');
    assert.deepStrictEqual(svg.warnings.map(w => w.details?.reason), ['RTF output does not take SVG pictures (it takes PNG, JPEG, EMF and WMF, and GIF and BMP as PNG)'], 'RTF signatures: an SVG reported');
    // A TIFF is not converted: it is its alt text, whatever its pages or compression, and reported as an SVG is.
    for (const [label, spec] of [['TIFF', tiffSpec], ['TIFF of two pages', { ...tiffSpec, pages: 2 }], ['fax TIFF', { ...tiffSpec, compression: 4 }]] as const) {
        const refused = await rtfOf(tiffOf(spec), 'scan.tif', 'image/tiff');
        assert.ok(!refused.rtf.includes('\\pict') && refused.rtf.includes('ALTWORDS'), `RTF signatures: a ${label} is its alt text`);
        assert.deepStrictEqual(refused.warnings.map(w => [w.code, w.details?.name, w.details?.reason]), [['IMAGE_PROCESSING_FAILED', 'scan.tif', 'RTF output does not take TIFF pictures (it takes PNG, JPEG, EMF and WMF, and GIF and BMP as PNG)']], `RTF signatures: a ${label} reported as not taken`);
    }
    const noise = await rtfOf(new Uint8Array(64).map((_, i) => (i * 37 + 11) & 255), 'n.png', 'image/png');
    assert.ok(!noise.rtf.includes('\\pict') && noise.rtf.includes('ALTWORDS') && noise.warnings.length === 1, 'RTF signatures: bytes that are no picture, typed PNG, are their alt text and reported, not \\pngblip');

    // Each picture charged to the document's picture budget what it is written as: a GIF written as a PNG is charged
    // the PNG, at each of its showings; a PNG its own bytes; a TIFF, written nowhere, is reported at each showing.
    const charged = async (bytes: Uint8Array, name: string, mimeType: string, times: number) => {
        const warnings: any[] = [];
        const ast = astOf(Array.from({ length: times }, () => ({ type: 'paragraph', children: [{ type: 'image', text: '', metadata: { attachmentName: name } }] })) as OfficeContentNode[],
            [{ type: 'image', name, extension: name.split('.').pop(), mimeType, data: Buffer.from(bytes).toString('base64') }] as any);
        const generator = new RtfGenerator(ast, { onWarning: (w: any) => warnings.push(w) } as any);
        const rtf = String((await generator.generate()).value);
        const written = [...rtf.matchAll(/\{\\pict\\pngblip[^\n]*\n([0-9a-f\n]*)\n\}/g)].map(m => m[1].replace(/\n/g, '').length / 2);
        return { bytes: (generator as any).inlinedImageBytes as number, written, warnings };
    };
    const gifShown = await charged(GIF_1x1, 'g.gif', 'image/gif', 3);
    assert.ok(gifShown.written.length === 3 && gifShown.bytes === gifShown.written[0] * 3 && gifShown.warnings.length === 0, `RTF signatures: a GIF written as a PNG is charged the PNG at each showing (${gifShown.bytes} charged, ${JSON.stringify(gifShown.written)} written)`);
    const pngShown = await charged(pngOf(6, 4), 'p.png', 'image/png', 2);
    assert.strictEqual(pngShown.bytes, 2 * pngOf(6, 4).length, 'RTF signatures: a PNG shown twice is charged twice its bytes');
    const tiffShown = await charged(tiffOf(tiffSpec), 'scan.tif', 'image/tiff', 2);
    assert.deepStrictEqual([tiffShown.written, tiffShown.warnings.map(w => w.code)], [[], ['IMAGE_PROCESSING_FAILED', 'IMAGE_PROCESSING_FAILED']], 'RTF signatures: a TIFF is written nowhere, reported at each showing');
    console.log('  ✓ GIF and BMP written as PNG with their pixels; TIFF refused; JPEG, EMF, WMF by their bytes; SVG and noise reported; each picture charged what it is written as');
}

// ─── TIFF pictures drawn ─────────────────────────────────────────────────

/** `run(n)` timed at n and at n/4: the larger within 8 times the smaller plus 250 ms (measured again once when over). */
export const linearTime = async (label: string, run: (n: number) => unknown | Promise<unknown>, n: number): Promise<void> => {
    const quarter = Math.floor(n / 4);
    const time = async (count: number) => { const started = performance.now(); await run(count); return performance.now() - started; };
    let small = await time(quarter), large = await time(n);
    if (large > 8 * small + 250) { small = Math.min(small, await time(quarter)); large = Math.min(large, await time(n)); }
    assert.ok(large <= 8 * small + 250, `${label}: linear (${Math.round(small)} ms at ${quarter}, ${Math.round(large)} ms at ${n})`);
    console.log(`  ✓ ${label}: ${Math.round(small)} ms at ${quarter}, ${Math.round(large)} ms at ${n}`);
};

/** What a TIFF built by {@link tiffOf} holds. */
export interface TiffSpec {
    width: number; height: number;
    /** 0 WhiteIsZero, 1 BlackIsZero, 2 RGB, 3 palette. */
    photometric: 0 | 1 | 2 | 3;
    bits: 1 | 2 | 4 | 8 | 16;
    /** An extra alpha sample: 1 premultiplied, 2 not. */
    alpha?: 1 | 2;
    /** 1 none, 5 LZW, 8 Deflate, 32773 PackBits; any other number is written as given (with no data change). */
    compression: number;
    /** fflate's level for Deflate blocks; 0 writes stored blocks. */
    deflateLevel?: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;
    predictor?: 1 | 2;
    little: boolean;
    /** Tiles of this size, else strips of `rowsPerStrip` rows. */
    tile?: { w: number; h: number };
    rowsPerStrip?: number;
    orientation?: number;
    /** The samples as stored, row by row, each pixel's samples in turn. */
    samples: number[];
    /** A palette's 3 x 2^bits 16-bit values, reds then greens then blues. */
    palette?: number[];
    /** How many pages (each IFD the same image). */
    pages?: number;
    /** Fields written as given, over the built ones: [tag, type, values]. */
    extra?: [number, number, number[]][];
}

/** The number of samples a pixel of `spec` has. */
const samplesOf = (spec: TiffSpec) => (spec.photometric === 2 ? 3 : 1) + (spec.alpha ? 1 : 0);

/** TIFF's LZW, as libtiff writes it: a clear code first, widths changed one code early, cleared again before the table fills. */
export const tiffLzwEncode = (data: Uint8Array): Uint8Array => {
    const out: number[] = [];
    let acc = 0, held = 0, width = 9;
    const emit = (code: number) => { acc = (acc << width) | code; held += width; while (held >= 8) { out.push((acc >> (held - 8)) & 0xFF); held -= 8; } acc &= (1 << held) - 1; };
    let table = new Map<string, number>(), next = 258;
    emit(256);
    let w = '';
    for (const byte of data) {
        const wc = w + String.fromCharCode(byte);
        if (w === '' || table.has(wc) || wc.length === 1) { w = wc; continue; }
        emit(w.length === 1 ? w.charCodeAt(0) : table.get(w)!);
        table.set(wc, next++);
        if (next >= 4094) { emit(256); table = new Map(); next = 258; width = 9; }
        else if (next > (1 << width) - 1 && width < 12) width++;
        w = String.fromCharCode(byte);
    }
    if (w) emit(w.length === 1 ? w.charCodeAt(0) : table.get(w)!);
    // (The end code is read at the width the decoder has reached after the last code: one more entry.)
    if (next + 1 > (1 << width) - 1 && width < 12) width++;
    emit(257);
    if (held) out.push((acc << (8 - held)) & 0xFF);
    return new Uint8Array(out);
};

/** PackBits: runs of three or more equal bytes as repeats, the rest as literals of up to 128. */
export const packBits = (data: Uint8Array): Uint8Array => {
    const out: number[] = [];
    for (let i = 0; i < data.length;) {
        let run = 1;
        while (i + run < data.length && run < 128 && data[i + run] === data[i]) run++;
        if (run >= 3) { out.push(257 - run, data[i]); i += run; continue; }
        const start = i;
        while (i < data.length && i - start < 128 && !(i + 2 < data.length && data[i] === data[i + 1] && data[i] === data[i + 2])) i++;
        out.push(i - start - 1, ...data.subarray(start, i));
    }
    return new Uint8Array(out);
};

/** A TIFF of `spec`: its blocks packed, predicted and compressed as it says, its IFD (or IFDs) after them. */
export const tiffOf = (spec: TiffSpec): Uint8Array => {
    const { width, height, bits, little } = spec;
    const spp = samplesOf(spec);
    const blockW = spec.tile?.w ?? width, blockH = spec.tile?.h ?? (spec.rowsPerStrip ?? height);
    const across = Math.ceil(width / blockW), down = Math.ceil(height / blockH);
    const rowBytes = Math.ceil(blockW * spp * bits / 8);
    const blocks: Uint8Array[] = [];
    for (let by = 0; by < down; by++) for (let bx = 0; bx < across; bx++) {
        const rows = spec.tile ? blockH : Math.min(blockH, height - by * blockH);
        const raw = new Uint8Array(rowBytes * rows);
        for (let y = 0; y < rows; y++) {
            const sy = by * blockH + y;
            const values: number[] = [];
            for (let x = 0; x < blockW; x++) for (let s = 0; s < spp; s++) {
                const sx = bx * blockW + x;
                values.push(sx < width && sy < height ? spec.samples[(sy * width + sx) * spp + s] : 0);
            }
            if (spec.predictor === 2) for (let i = values.length - 1; i >= spp; i--) values[i] = (values[i] - values[i - spp]) & (bits === 16 ? 0xFFFF : 0xFF);
            const row = y * rowBytes;
            values.forEach((v, i) => {
                if (bits === 16) { raw[row + i * 2] = little ? v & 0xFF : v >> 8; raw[row + i * 2 + 1] = little ? v >> 8 : v & 0xFF; }
                else if (bits === 8) raw[row + i] = v;
                else { const bit = i * bits; raw[row + (bit >> 3)] |= v << (8 - bits - (bit & 7)); }
            });
        }
        blocks.push(spec.compression === 5 ? tiffLzwEncode(raw) : spec.compression === 8 ? zlibSync(raw, { level: spec.deflateLevel ?? 6 }) : spec.compression === 32773 ? packBits(raw) : raw);
    }
    const out: number[] = [];
    const w16 = (v: number) => little ? [v & 0xFF, (v >> 8) & 0xFF] : [(v >> 8) & 0xFF, v & 0xFF];
    const w32 = (v: number) => little ? [v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF] : [(v >>> 24) & 0xFF, (v >>> 16) & 0xFF, (v >>> 8) & 0xFF, v & 0xFF];
    out.push(...(little ? [0x49, 0x49] : [0x4D, 0x4D]), ...w16(42), ...w32(0));
    const offsets: number[] = [];
    for (const b of blocks) { offsets.push(out.length); for (const v of b) out.push(v); if (out.length & 1) out.push(0); }
    const fields = new Map<number, [number, number[]]>([
        [256, [4, [width]]], [257, [4, [height]]], [258, [3, new Array(spp).fill(bits)]], [259, [3, [spec.compression]]],
        [262, [3, [spec.photometric]]], [277, [3, [spp]]], [284, [3, [1]]],
    ]);
    if (spec.tile) fields.set(322, [3, [blockW]]).set(323, [3, [blockH]]).set(324, [4, offsets]).set(325, [4, blocks.map(b => b.length)]);
    else fields.set(273, [4, offsets]).set(278, [4, [blockH]]).set(279, [4, blocks.map(b => b.length)]);
    if (spec.predictor) fields.set(317, [3, [spec.predictor]]);
    if (spec.orientation) fields.set(274, [3, [spec.orientation]]);
    if (spec.alpha) fields.set(338, [3, [spec.alpha]]);
    if (spec.palette) fields.set(320, [3, spec.palette]);
    for (const [tag, type, values] of spec.extra ?? []) fields.set(tag, [type, values]);
    const sorted = [...fields.entries()].sort((a, b) => a[0] - b[0]);
    const pages = spec.pages ?? 1;
    const ifdSize = 2 + sorted.length * 12 + 4;
    // Each page's IFD, its long values after it.
    let firstIfd = 0;
    for (let page = 0; page < pages; page++) {
        const ifdAt = out.length;
        if (!page) firstIfd = ifdAt;
        let valuesAt = ifdAt + ifdSize;
        const entries: number[] = [], valueBytes: number[] = [];
        entries.push(...w16(sorted.length));
        for (const [tag, [type, values]] of sorted) {
            const size = type === 3 ? 2 : 4;
            const bytes = values.flatMap(v => size === 2 ? w16(v) : w32(v));
            entries.push(...w16(tag), ...w16(type), ...w32(values.length));
            if (bytes.length <= 4) entries.push(...bytes, ...new Array(4 - bytes.length).fill(0));
            else { entries.push(...w32(valuesAt + valueBytes.length)); valueBytes.push(...bytes); if (valueBytes.length & 1) valueBytes.push(0); }
        }
        const nextIfd = page + 1 < pages ? valuesAt + valueBytes.length : 0;
        entries.push(...w32(nextIfd));
        out.push(...entries, ...valueBytes);
    }
    const bytes = new Uint8Array(out);
    new DataView(bytes.buffer).setUint32(4, firstIfd, little);
    return bytes;
};

/**
 * A little-endian TIFF: `data` at byte 8, then one IFD of `fields` ([tag, type 3 SHORT or 4 LONG, values]; an offset
 * into the data is 8 plus its place in it), each value longer than four bytes after the IFD.
 */
export const rawTiff = (fields: [number, 3 | 4, number[]][], data: Uint8Array): Uint8Array => {
    const sorted = [...fields].sort((a, b) => a[0] - b[0]);
    const ifdAt = 8 + data.length + (data.length & 1);
    let longAt = ifdAt + 2 + sorted.length * 12 + 4;
    const longBytes = sorted.reduce((n, [, type, values]) => n + ((type === 3 ? 2 : 4) * values.length > 4 ? (type === 3 ? 2 : 4) * values.length : 0), 0);
    const out = new Uint8Array(longAt + longBytes);
    const view = new DataView(out.buffer);
    out.set([0x49, 0x49, 42, 0]);
    view.setUint32(4, ifdAt, true);
    out.set(data, 8);
    view.setUint16(ifdAt, sorted.length, true);
    sorted.forEach(([tag, type, values], k) => {
        const entry = ifdAt + 2 + k * 12, size = type === 3 ? 2 : 4;
        view.setUint16(entry, tag, true);
        view.setUint16(entry + 2, type, true);
        view.setUint32(entry + 4, values.length, true);
        let at = entry + 8;
        if (size * values.length > 4) { view.setUint32(entry + 8, longAt, true); at = longAt; longAt += size * values.length; }
        values.forEach((v, i) => { if (size === 2) view.setUint16(at + i * 2, v, true); else view.setUint32(at + i * 4, v, true); });
    });
    return out;
};

/**
 * TIFF pictures are refused by every output: the native PDF engine and RTF each name TIFF in the warning, whatever the
 * picture's pages or compression, and draw or convert nothing.
 */
export async function testTiffPicturesRefused(): Promise<void> {
    console.log('\n=== TIFF pictures ===');
    const grey = (extra: TiffSpec['extra'], more: Partial<TiffSpec> = {}): Uint8Array => tiffOf({ width: 4, height: 2, photometric: 1, bits: 8, compression: 1, little: true, samples: new Array(8).fill(9), extra, ...more });

    // The native PDF engine refuses a TIFF whatever its pages or compression, and says it does not embed TIFF.
    const rgbSpec: TiffSpec = { width: 30, height: 10, photometric: 2, bits: 8, compression: 5, little: false, samples: Array.from({ length: 900 }, (_, i) => i * 7 & 255) };
    const refusal = 'the native PDF engine does not embed TIFF pictures (only PNG, JPEG, GIF, BMP, SVG, and bitmaps in EMF and WMF)';
    for (const [file, bytes, alt] of [['scan.tif', tiffOf(rgbSpec), undefined], ['scan.tiff', tiffOf({ ...rgbSpec, pages: 3 }), undefined], ['fax.tif', grey([], { compression: 3 }), 'FAXALT']] as const) {
        const refused = await writeWithWarnings(onePicture(bytes, file, 'image/tiff', alt), 'pdf');
        assert.deepStrictEqual(refused.warnings.map(w => [w.code, w.details?.name, w.details?.reason]), [['IMAGE_PROCESSING_FAILED', file, refusal]], `TIFF: ${file} refused by the native PDF engine`);
        assert.deepStrictEqual((await pdfPictures(refused.value)).images, [], `TIFF: ${file} not drawn`);
    }
    // RTF writes its alt text, as for an SVG, and says it does not take TIFF.
    const rtfRefused = await writeWithWarnings(onePicture(tiffOf(rgbSpec), 'scan.tif', 'image/tiff', 'SCANALT'), 'rtf');
    assert.deepStrictEqual(rtfRefused.warnings.map(w => w.code), ['IMAGE_PROCESSING_FAILED'], 'TIFF: RTF refuses it');
    assert.ok(/does not take TIFF pictures/.test(String(rtfRefused.warnings[0].details?.reason)), 'TIFF: RTF says it does not take TIFF');
}

// ─── EMF and WMF drawn from the bitmap they hold ─────────────────────────

const le16 = (v: number) => [v & 0xFF, (v >>> 8) & 0xFF];
const le32 = (v: number) => [v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF];
const f32le = (v: number) => { const b = new Uint8Array(4); new DataView(b.buffer).setFloat32(0, v, true); return [...b]; };
/** `bytes` padded with zeros to a multiple of `n`. */
const padTo = (bytes: number[], n: number) => { while (bytes.length % n) bytes.push(0); return bytes; };

/** A device-independent bitmap: a BITMAPINFOHEADER and its bits, 24-bit rows stored from the foot up (pixels as {@link pngOf} colours them). */
export const dib24 = (w: number, h: number): { header: number[]; bits: number[] } => {
    const stride = Math.ceil(w * 3 / 4) * 4;
    const bits = new Array(stride * h).fill(0);
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
        const at = (h - 1 - y) * stride + x * 3;
        bits[at] = (x + y) & 255; bits[at + 1] = y * 11 & 255; bits[at + 2] = x * 7 & 255;
    }
    return { header: [...le32(40), ...le32(w), ...le32(h), ...le16(1), ...le16(24), ...le32(0), ...le32(bits.length), ...le32(2835), ...le32(2835), ...le32(0), ...le32(0)], bits };
};

/** A DIB holding a PNG or JPEG file (BI_PNG 5, BI_JPEG 4). */
export const dibOfFile = (file: Uint8Array, compression: 4 | 5, w: number, h: number) => ({ header: [...le32(40), ...le32(w), ...le32(h), ...le16(1), ...le16(0), ...le32(compression), ...le32(file.length), ...le32(0), ...le32(0), ...le32(0), ...le32(0)], bits: [...file] });

/** A 32-bit DIB of one row of BGRA `pixels` (four bytes each, as stored). */
const dib32 = (pixels: number[]) => ({ header: [...le32(40), ...le32(pixels.length / 4), ...le32(1), ...le16(1), ...le16(32), ...le32(0), ...le32(pixels.length), ...le32(0), ...le32(0), ...le32(0), ...le32(0)], bits: pixels });

/** An EMF record of `type` holding `body` after its type and size. */
const emfRecord = (type: number, body: number[]) => padTo([...le32(type), ...le32(8 + padTo([...body], 4).length), ...padTo([...body], 4)], 4);

/** EMR_STRETCHDIBITS drawing `dib` over `destW` x `destH`, from its source rectangle (all of it unless given). */
export const stretchDibits = (dib: { header: number[]; bits: number[] }, destW: number, destH: number, src?: [number, number, number, number]) => {
    const w = dib.header[4] | dib.header[5] << 8, h = dib.header[8] | dib.header[9] << 8;
    const [sx, sy, sw, sh] = src ?? [0, 0, w, h];
    const fixed = [...le32(0), ...le32(0), ...le32(destW), ...le32(destH), ...le32(0), ...le32(0), ...le32(sx), ...le32(sy), ...le32(sw), ...le32(sh),
        ...le32(80), ...le32(dib.header.length), ...le32(80 + dib.header.length), ...le32(dib.bits.length), ...le32(0), ...le32(0x00CC0020), ...le32(destW), ...le32(destH)];
    return emfRecord(81, [...fixed, ...dib.header, ...dib.bits]);
};

/**
 * EMR_ALPHABLEND (114) or EMR_TRANSPARENTBLT (116) drawing all of `dib` over `destW` x `destH`; `blend` the four bytes
 * at offset 40: a BLENDFUNCTION (operation, flags, SourceConstantAlpha, AlphaFormat), or a TransparentColor (red,
 * green, blue, 0).
 */
const blendBlt = (type: 114 | 116, dib: { header: number[]; bits: number[] }, destW: number, destH: number, blend: number[]) => {
    const w = dib.header[4] | dib.header[5] << 8, h = dib.header[8] | dib.header[9] << 8;
    const fixed = [...new Array(16).fill(0), ...le32(0), ...le32(0), ...le32(destW), ...le32(destH), ...blend, ...le32(0), ...le32(0), ...new Array(24).fill(0),
        ...le32(0), ...le32(0), ...le32(108), ...le32(dib.header.length), ...le32(108 + dib.header.length), ...le32(dib.bits.length), ...le32(w), ...le32(h)];
    return emfRecord(type, [...fixed, ...dib.header, ...dib.bits]);
};

/** An EMF: its header (frame `frameW` x `frameH` hundredths of a millimetre), `records`, EMR_EOF. */
export const emfOf = (records: number[][], frameW = 2540, frameH = 1270): Uint8Array => {
    const body = records.flat();
    const eof = emfRecord(14, [...le32(0), ...le32(16), ...le32(20)]);
    const header = [...le32(1), ...le32(108), ...le32(0), ...le32(0), ...le32(95), ...le32(47), ...le32(0), ...le32(0), ...le32(frameW), ...le32(frameH),
        0x20, 0x45, 0x4D, 0x46, ...le32(0x10000), ...le32(108 + body.length + eof.length), ...le32(records.length + 2), ...le16(1), ...le16(0),
        ...le32(0), ...le32(0), ...le32(0), ...le32(1920), ...le32(1080), ...le32(508), ...le32(286), ...le32(0), ...le32(0), ...le32(0), ...le32(508000), ...le32(286000)];
    return new Uint8Array([...header, ...body, ...eof]);
};

/** EMR_COMMENT holding EMF+ records, each [type, flags, data]. */
export const emfPlus = (records: [number, number, number[]][]) => {
    const plus = records.flatMap(([type, flags, data]) => [...le16(type), ...le16(flags), ...le32(12 + data.length), ...le32(data.length), ...data]);
    return emfRecord(70, [...le32(4 + plus.length), ...le32(0x2B464D45), ...plus]);
};
/** An EMF+ header record (dual). */
export const plusHeader: [number, number, number[]] = [0x4001, 1, [...le32(0xDBC01002), ...le32(0), ...le32(96), ...le32(96)]];
/** An EMF+ Image object `id`: a compressed bitmap holding `file`, or `pixels` of a format. */
export const plusImage = (id: number, w: number, h: number, file?: Uint8Array, pixels?: { format: number; stride: number; data: number[] }): [number, number, number[]] =>
    [0x4008, id | 5 << 8, [...le32(0xDBC01002), ...le32(1), ...le32(w), ...le32(h), ...le32(pixels?.stride ?? 0), ...le32(pixels?.format ?? 0), ...le32(file ? 1 : 0), ...(file ? [...file] : pixels!.data)]];
/** EmfPlusDrawImage of object `id`, its source the whole `w` x `h` and its destination `dw` x `dh`. */
export const plusDrawImage = (id: number, w: number, h: number, dw: number, dh: number): [number, number, number[]] =>
    [0x401A, id, [...le32(0), ...le32(2), ...f32le(0), ...f32le(0), ...f32le(w), ...f32le(h), ...f32le(0), ...f32le(0), ...f32le(dw), ...f32le(dh)]];

/** A placeable WMF (box `w` x `h` at 1440 an inch) of `records` (each [function, parameter bytes]). */
export const wmfOf = (records: [number, number[]][], w = 1440, h = 720): Uint8Array => {
    const body = records.flatMap(([fn, params]) => { const p = padTo([...params], 2); return [...le32(3 + p.length / 2), ...le16(fn), ...p]; });
    const eof = [...le32(3), ...le16(0)];
    const words = (18 + body.length + eof.length) / 2;
    const meta = [...le16(1), ...le16(9), ...le16(0x0300), ...le32(words), ...le16(0), ...le32(64), ...le16(0)];
    const head = [0xCDD7, 0x9AC6, 0, 0, 0, w, h, 1440, 0, 0];
    const checksum = head.reduce((a, v) => a ^ v, 0);
    return new Uint8Array([...head.flatMap(le16), ...le16(checksum), ...meta, ...body, ...eof]);
};
/** META_STRETCHDIB drawing `dib` over `dw` x `dh`. */
export const metaStretchDib = (dib: { header: number[]; bits: number[] }, w: number, h: number, dw: number, dh: number): [number, number[]] =>
    [0x0F43, [...le32(0x00CC0020), ...le16(0), ...le16(h), ...le16(w), ...le16(0), ...le16(0), ...le16(dh), ...le16(dw), ...le16(0), ...le16(0), ...dib.header, ...dib.bits]];
/** META_DIBBITBLT drawing `dib` at its size. */
const metaDibBitBlt = (dib: { header: number[]; bits: number[] }, w: number, h: number): [number, number[]] =>
    [0x0940, [...le32(0x00CC0020), ...le16(0), ...le16(0), ...le16(h), ...le16(w), ...le16(0), ...le16(0), ...dib.header, ...dib.bits]];

/** The pictures of a native PDF read back: each image node's size, in points. */
const drawnSizes = async (value: unknown) => {
    const back = await OfficeParser.parseOffice(Buffer.from(bytesOf(value)), { extractAttachments: true });
    const sizes: { w: number; h: number }[] = [];
    const walk = (n: OfficeContentNode) => { if (n.type === 'image' && n.bounds) sizes.push({ w: Math.round(n.bounds.width), h: Math.round(n.bounds.height) }); n.children?.forEach(walk); };
    back.content.forEach(walk);
    return sizes;
};

/**
 * EMF and WMF pictures holding a bitmap are drawn from it at the metafile's frame: EMR_STRETCHDIBITS of a 24-bit DIB
 * (with state records before it), of a PNG (BI_PNG) and of a JPEG (BI_JPEG); an EMF+ dual metafile with a compressed
 * PNG and DrawImage (its GDI copy, text and a polygon, not reported); an EMF+ bitmap of 32-bit ARGB pixels; a
 * placeable WMF of META_STRETCHDIB and of META_DIBBITBLT. A bitmap with text beside it is drawn and the text
 * reported; a metafile of lines and text, with no bitmap, is reported; a cropped DIB draws its part. A bitmap drawn
 * with EMR_ALPHABLEND keeps its per-pixel and constant alpha, and one drawn with EMR_TRANSPARENTBLT its transparent
 * colour; a PNG drawn so is drawn opaque and said so. An EMF+ JPEG drawn in part is drawn whole and said so, and a TIFF
 * (its part read in its oriented size) is refused, as no output draws a TIFF. Generated
 * record sequences are read as what they hold, and none throws.
 */
export async function testMetafileBitmaps(): Promise<void> {
    console.log('\n=== Metafiles drawn from their bitmaps ===');
    const dib = dib24(20, 10);
    const state = [emfRecord(9, [...le32(100), ...le32(50)]), emfRecord(10, [...le32(0), ...le32(0)]), emfRecord(17, [...le32(8)])]; // SETWINDOWEXTEX, SETWINDOWORGEX, SETMAPMODE
    const png = pngOf(12, 6);
    const argb = Array.from({ length: 4 * 3 * 2 }, (_, i) => i % 4 === 3 ? 128 : i * 9 & 255);
    const drawn: [string, Uint8Array, string, { w: number; h: number }][] = [
        ['an EMF of a 24-bit DIB', emfOf([...state, stretchDibits(dib, 100, 50)]), 'a.emf', { w: 72, h: 36 }],
        ['an EMF of a BI_PNG DIB', emfOf([stretchDibits(dibOfFile(png, 5, 12, 6), 100, 50)], 5080, 2540), 'b.emf', { w: 144, h: 72 }],
        ['an EMF of a BI_JPEG DIB', emfOf([stretchDibits(dibOfFile(JPEG_192x48, 4, 192, 48), 100, 50)]), 'c.emf', { w: 72, h: 36 }],
        ['an EMF+ dual of a compressed PNG', emfOf([emfPlus([plusHeader, plusImage(0, 12, 6, png), plusDrawImage(0, 12, 6, 100, 50)]), emfRecord(84, new Array(40).fill(0)), emfRecord(86, new Array(20).fill(0))]), 'd.emf', { w: 72, h: 36 }],
        ['an EMF+ bitmap of ARGB pixels', emfOf([emfPlus([plusHeader, plusImage(3, 3, 2, undefined, { format: 0x0026200A, stride: 12, data: argb }), plusDrawImage(3, 3, 2, 30, 20)])]), 'e.emf', { w: 72, h: 36 }],
        ['a WMF of META_STRETCHDIB', wmfOf([metaStretchDib(dib, 20, 10, 1440, 720)]), 'f.wmf', { w: 72, h: 36 }],
        ['a WMF of META_DIBBITBLT', wmfOf([metaDibBitBlt(dib, 20, 10)], 2880, 1440), 'g.wmf', { w: 144, h: 72 }],
    ];
    for (const [label, bytes, name, size] of drawn) {
        const { value, warnings } = await writeWithWarnings(onePicture(bytes, name, name.endsWith('emf') ? 'image/emf' : 'image/wmf'), 'pdf');
        assert.deepStrictEqual(warnings.map(w => w.details), [], `Metafiles: ${label} warns nothing`);
        // (The JPEG here is its headers alone, which a PDF reader cannot show: its image is counted instead.)
        if (label.includes('JPEG')) assert.deepStrictEqual((await pdfPictures(value)).images, [{ w: 192, h: 48 }], `Metafiles: ${label} drawn`);
        else assert.deepStrictEqual(await drawnSizes(value), [size], `Metafiles: ${label} drawn at its frame's size`);
    }
    // The ARGB pixels kept with their alpha.
    const argbReading = readMetafile(drawn[4][1], 'EMF');
    const argbPixels = pngPixels(argbReading.bitmap!.bytes);
    assert.deepStrictEqual([argbPixels.rgba[0], argbPixels.rgba[1], argbPixels.rgba[2], argbPixels.rgba[3]], [argb[2], argb[1], argb[0], 128], 'Metafiles: an ARGB pixel kept, its alpha too');

    // A bitmap and text: drawn, the text reported; lines and text alone: reported, nothing drawn.
    const withText = await writeWithWarnings(onePicture(emfOf([stretchDibits(dib, 100, 50), emfRecord(84, new Array(40).fill(0))]), 'image1.emf', 'image/emf'), 'pdf');
    assert.deepStrictEqual(withText.warnings.map(w => [w.code, w.details]), [['CONTENT_NOT_REPRESENTABLE', { format: 'pdf', feature: "EMF picture 'image1.emf': its text (the native engine draws only the bitmap it holds)", reason: 'metafile-vector', name: 'image1.emf' }]], 'Metafiles: text beside the bitmap reported');
    assert.strictEqual((await drawnSizes(withText.value)).length, 1, 'Metafiles: the bitmap beside text drawn');
    const vectorEmf = await writeWithWarnings(onePicture(emfOf([emfRecord(86, new Array(20).fill(0)), emfRecord(87, new Array(20).fill(0)), emfRecord(84, new Array(40).fill(0))]), 'pic.emf', 'image/emf', 'ALT'), 'pdf');
    assert.deepStrictEqual(vectorEmf.warnings.map(w => w.details?.reason), ['the native PDF engine draws an EMF picture only from a bitmap it holds, and this one holds none (it is drawn with lines, shapes and text)'], 'Metafiles: a vector EMF reported');
    const vectorWmf = await writeWithWarnings(onePicture(wmfOf([[0x0324, [...le16(3), ...new Array(12).fill(0)]], [0x0521, [...le16(2), 0x41, 0x42, ...le16(0), ...le16(0)]]]), 'pic.wmf', 'image/wmf', 'ALT'), 'pdf');
    assert.deepStrictEqual(vectorWmf.warnings.map(w => w.details?.reason), ['the native PDF engine draws a WMF picture only from a bitmap it holds, and this one holds none (it is drawn with shapes and text)'], 'Metafiles: a vector WMF reported');
    const damaged = emfOf([stretchDibits(dib, 100, 50)]);
    new DataView(damaged.buffer).setUint32(108 + 4, 7, true);
    assert.deepStrictEqual(readMetafile(damaged, 'EMF'), { frame: { width: 96, height: 48 }, alsoDraws: [], reason: 'its EMF records are damaged' }, 'Metafiles: a record of 7 bytes stops the walk');

    // A cropped DIB (its bottom-up rows): the part drawn is its source rectangle, counted from its foot.
    const cropped = readMetafile(emfOf([stretchDibits(dib, 100, 50, [5, 2, 10, 4])]), 'EMF');
    assert.deepStrictEqual(cropped.bitmap?.crop, { x: 5, y: 4, width: 10, height: 4 }, 'Metafiles: a DIB source rectangle counted from its foot');
    const croppedPdf = await writeWithWarnings(onePicture(emfOf([stretchDibits(dib, 100, 50, [5, 2, 10, 4])]), 'h.emf', 'image/emf'), 'pdf');
    assert.deepStrictEqual((await pdfPictures(croppedPdf.value)).images, [{ w: 10, h: 4 }], 'Metafiles: the part of the DIB drawn');
    const croppedPng = await writeWithWarnings(onePicture(emfOf([stretchDibits(dibOfFile(png, 5, 12, 6), 100, 50, [0, 0, 6, 6])]), 'i.emf', 'image/emf'), 'pdf');
    assert.deepStrictEqual(croppedPng.warnings.map(w => w.details?.feature), ["EMF picture 'i.emf': the part of its bitmap outside the picture (the native engine draws only the bitmap it holds)"], 'Metafiles: a part of a PNG drawn whole, and said so');

    // A bitmap drawn with transparency: EMR_ALPHABLEND's premultiplied pixels and constant alpha, EMR_TRANSPARENTBLT's colour.
    const decodedAs = (emf: Uint8Array) => { const reading = readMetafile(emf, 'EMF'); return Array.from(decodeBmp(reading.bitmap!.bytes, 1_000_000, reading.bitmap!.transparency)!.rgba); };
    const alphaOf = (rgba: number[]) => rgba.filter((_, i) => (i & 3) === 3);
    // (A transparent pixel, and red at half alpha, premultiplied, stored blue, green, red, alpha.)
    const halfRed = dib32([0, 0, 0, 0, 0, 0, 128, 128]);
    const alphaBlend = emfOf([blendBlt(114, halfRed, 20, 10, [0, 0, 255, 1])]);
    assert.deepStrictEqual(readMetafile(alphaBlend, 'EMF').bitmap?.transparency, { pixelAlpha: true }, 'Metafiles: EMR_ALPHABLEND with AC_SRC_ALPHA read');
    assert.deepStrictEqual(decodedAs(alphaBlend), [0, 0, 0, 0, 255, 0, 0, 128], 'Metafiles: EMR_ALPHABLEND pixels keep their alpha, red un-premultiplied');
    assert.deepStrictEqual(alphaOf(decodedAs(emfOf([blendBlt(114, halfRed, 20, 10, [0, 0, 128, 1])]))), [0, 64], 'Metafiles: EMR_ALPHABLEND per-pixel alpha times its constant alpha');
    assert.deepStrictEqual(alphaOf(decodedAs(emfOf([blendBlt(114, dib24(4, 2), 20, 10, [0, 0, 128, 0])]))), new Array(8).fill(128), 'Metafiles: EMR_ALPHABLEND of a 24-bit DIB at constant alpha 128');
    // (Magenta, then green, stored blue, green, red, in a row of eight bytes.)
    const keyed = { header: [...le32(40), ...le32(2), ...le32(1), ...le16(1), ...le16(24), ...le32(0), ...le32(8), ...le32(0), ...le32(0), ...le32(0), ...le32(0)], bits: [255, 0, 255, 0, 200, 0, 0, 0] };
    assert.deepStrictEqual(decodedAs(emfOf([blendBlt(116, keyed, 20, 10, [255, 0, 255, 0])])), [255, 0, 255, 0, 0, 200, 0, 255], 'Metafiles: EMR_TRANSPARENTBLT draws its colour transparent');
    // Through native PDF: drawn with a soft mask, no warning; a 32-bit EMR_STRETCHDIBITS stays opaque.
    const softMasks = async (value: unknown) => {
        const doc = await PDFDocument.load(bytesOf(value));
        return doc.context.enumerateIndirectObjects().filter(([, o]) => o instanceof PDFRawStream && o.dict.get(PDFName.of('Subtype')) === PDFName.of('Image') && !!o.dict.get(PDFName.of('SMask'))).length;
    };
    const blended = await writeWithWarnings(onePicture(alphaBlend, 'blend.emf', 'image/emf'), 'pdf');
    assert.deepStrictEqual([blended.warnings.length, await softMasks(blended.value)], [0, 1], 'Metafiles: an EMR_ALPHABLEND bitmap drawn with a soft mask, no warning');
    const stretched32 = emfOf([stretchDibits(halfRed, 20, 10)]);
    assert.strictEqual(readMetafile(stretched32, 'EMF').bitmap?.transparency, undefined, 'Metafiles: EMR_STRETCHDIBITS draws its bitmap opaque');
    assert.deepStrictEqual(alphaOf(decodedAs(stretched32)), [255, 255], 'Metafiles: a 32-bit EMR_STRETCHDIBITS bitmap opaque');
    const opaque32 = await writeWithWarnings(onePicture(stretched32, 'opaque.emf', 'image/emf'), 'pdf');
    assert.deepStrictEqual([opaque32.warnings.length, await softMasks(opaque32.value)], [0, 0], 'Metafiles: a 32-bit EMR_STRETCHDIBITS bitmap drawn opaque, no warning');
    const blendedPng = await writeWithWarnings(onePicture(emfOf([blendBlt(114, dibOfFile(png, 5, 12, 6), 100, 50, [0, 0, 128, 0])]), 'p.emf', 'image/emf'), 'pdf');
    assert.deepStrictEqual(blendedPng.warnings.map(w => [w.details?.reason, w.details?.feature]), [['metafile-vector', "EMF picture 'p.emf': the transparency it is drawn with (the native engine draws only the bitmap it holds)"]], 'Metafiles: a PNG drawn at constant alpha drawn opaque, and said so');
    assert.strictEqual((await pdfPictures(blendedPng.value)).images.length, 1, 'Metafiles: the PNG drawn at constant alpha drawn');

    // EMF+ JPEG and TIFF bitmaps drawn in part: a JPEG drawn whole and said so; a TIFF refused.
    const plusPart = (file: Uint8Array, w: number, h: number, partW: number, partH: number) => emfOf([emfPlus([plusHeader, plusImage(1, w, h, file), plusDrawImage(1, partW, partH, 50, 50)])]);
    const halfJpeg = plusPart(JPEG_192x48, 192, 48, 96, 48);
    assert.deepStrictEqual(readMetafile(halfJpeg, 'EMF').bitmap?.crop, { x: 0, y: 0, width: 96, height: 48 }, 'Metafiles: an EMF+ JPEG drawn in half has its part');
    const halfJpegPdf = await writeWithWarnings(onePicture(halfJpeg, 'j.emf', 'image/emf'), 'pdf');
    assert.deepStrictEqual(halfJpegPdf.warnings.map(w => w.details?.feature), ["EMF picture 'j.emf': the part of its bitmap outside the picture (the native engine draws only the bitmap it holds)"], 'Metafiles: an EMF+ JPEG drawn in part is drawn whole, and said so');
    const tiff4x2: TiffSpec = { width: 4, height: 2, photometric: 2, bits: 8, compression: 5, little: true, samples: Array.from({ length: 24 }, (_, i) => i * 10) };
    // The metafile reader still finds a TIFF bitmap and its part (read in its oriented size), but no TIFF is drawn:
    // the engine refuses it, naming the picture and that it does not embed TIFF.
    const tiffRefusal = ['IMAGE_PROCESSING_FAILED', 'the bitmap it holds cannot be drawn: the native PDF engine does not embed TIFF pictures (only PNG, JPEG, GIF, BMP, SVG, and bitmaps in EMF and WMF)'];
    const refusedTiff = async (file: Uint8Array, name: string) => {
        const out = await writeWithWarnings(onePicture(file, name, 'image/emf'), 'pdf');
        assert.deepStrictEqual(out.warnings.map(w => [w.code, w.details?.reason]), [tiffRefusal], `Metafiles: a TIFF in ${name} refused`);
        assert.deepStrictEqual((await pdfPictures(out.value)).images, [], `Metafiles: a TIFF in ${name} not drawn`);
    };
    await refusedTiff(plusPart(tiffOf(tiff4x2), 4, 2, 2, 2), 't.emf');
    // (Orientation 6 shows it 2 wide and 4 tall: a part 2 x 3 is inside that, not inside the 4 x 2 it is stored as.)
    const turned = plusPart(tiffOf({ ...tiff4x2, orientation: 6 }), 2, 4, 2, 3);
    assert.deepStrictEqual(readMetafile(turned, 'EMF').bitmap?.crop, { x: 0, y: 0, width: 2, height: 3 }, 'Metafiles: an EMF+ TIFF part read in its oriented size');
    await refusedTiff(turned, 'o.emf');
    await refusedTiff(plusPart(tiffOf({ ...tiff4x2, pages: 2 }), 4, 2, 4, 2), 'x.emf');

    // Generated record sequences: state, bitmaps, vector records and damaged sizes; what each holds is what is read.
    const rnd = lcg(0xE3F5);
    const keys = new Set<string>();
    const vectorTypes: [number, string][] = [[3, 'shapes'], [4, 'lines'], [42, 'shapes'], [43, 'shapes'], [84, 'text'], [86, 'shapes'], [87, 'lines'], [108, 'text']];
    for (let round = 0; round < 1500; round++) {
        const records: number[][] = [];
        const expected = new Set<string>();
        const bitmaps: number[] = [];
        const count = Math.floor(rnd() * 7);
        let broken = false;
        const shape: string[] = [];
        for (let i = 0; i < count; i++) {
            const roll = rnd();
            if (roll < 0.25) { records.push(emfRecord(9 + Math.floor(rnd() * 3), [...le32(1), ...le32(1)])); shape.push('s'); }
            else if (roll < 0.5) { const w = 1 + Math.floor(rnd() * 6), h = 1 + Math.floor(rnd() * 6); bitmaps.push(w * h); records.push(stretchDibits(dib24(w, h), w * 10, h * 10)); shape.push(`b${w}x${h}`); }
            else if (roll < 0.9) { const [type, what] = vectorTypes[Math.floor(rnd() * vectorTypes.length)]; records.push(emfRecord(type, new Array(12).fill(0))); expected.add(what); shape.push(`v${type}`); }
            else if (!broken) { const r = emfRecord(10, [...le32(0), ...le32(0)]); r[4] = 3; records.push(r); broken = true; shape.push('x'); break; }
        }
        keys.add(shape.join(','));
        const reading = readMetafile(emfOf(records), 'EMF');
        if (broken) {
            assert.ok(reading.bitmap || reading.reason === 'its EMF records are damaged', `Metafiles: case ${round} (${shape}) stops at the damaged record`);
            continue;
        }
        if (bitmaps.length > 1) expected.add('other bitmaps');
        assert.deepStrictEqual(reading.alsoDraws, ['lines', 'shapes', 'text', 'other bitmaps'].filter(w => expected.has(w)), `Metafiles: case ${round} (${shape}) draws what it holds`);
        assert.strictEqual(!!reading.bitmap, bitmaps.length > 0, `Metafiles: case ${round} (${shape}) holds a bitmap or not`);
        if (bitmaps.length) {
            const largest = Math.max(...bitmaps);
            const bmp = reading.bitmap!.bytes;
            assert.strictEqual(bmp[18] * bmp[22], largest, `Metafiles: case ${round} (${shape}) the largest bitmap read`);
        }
    }
    // Random bytes after a header, as EMF and WMF: never a throw.
    for (let round = 0; round < 500; round++) {
        const noise = Array.from({ length: Math.floor(rnd() * 400) }, () => Math.floor(rnd() * 256));
        readMetafile(new Uint8Array([...emfOf([]).subarray(0, 108), ...noise]), 'EMF');
        readMetafile(new Uint8Array([...wmfOf([]).subarray(0, 40), ...noise]), 'WMF');
    }
    console.log(`  ✓ 7 metafiles drawn from their bitmaps at their frames; text and vector pictures reported; transparency kept, or reported for a PNG; EMF+ JPEG part, and a TIFF refused; ${keys.size} distinct generated record sequences read as they hold; 1,000 noisy files read without a throw`);
}

// ─── Pictures in headings, list items and notes drawn in their line ─────

/** A native PDF read back: its pages, and on each its text items and pictures with their bounds (top-left origin, points). */
const pdfLayout = async (value: unknown) => {
    const back = await OfficeParser.parseOffice(Buffer.from(bytesOf(value)), { extractAttachments: true });
    const pages: { texts: { text: string; x: number; y: number; bottom: number; baseline: number }[]; images: { x: number; y: number; w: number; h: number; bottom: number }[] }[] = [];
    for (const page of back.content) {
        const entry = { texts: [] as any[], images: [] as any[] };
        const walk = (n: OfficeContentNode) => {
            const b = n.bounds as any;
            if (n.type === 'text' && b) entry.texts.push({ text: n.text ?? '', x: b.x, y: b.y, bottom: b.y + b.height, baseline: b.baseline ?? b.y + b.height });
            if (n.type === 'image' && b) entry.images.push({ x: b.x, y: b.y, w: b.width, h: b.height, bottom: b.y + b.height });
            n.children?.forEach(walk);
        };
        walk(page);
        pages.push(entry);
    }
    return { pages, text: pages.flatMap(p => p.texts.map(t => t.text)).join('\n') };
};

/** The c5 document: a numbered list item, a Heading 1 and a footnote, each holding a 40x20 PNG, then a plain paragraph's. */
export const c5Lines = (alt: string): Buffer => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>`;
    const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const p = (inner: string, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${inner}</w:p>`;
    return docxOf(
        p(run('LISTITEM ') + pictureRun('rP', 1, alt && `${alt}LIST`), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')
        + p(run('HEADINGTEXT ') + pictureRun('rP', 2, alt && `${alt}HEAD`), '<w:pStyle w:val="Heading1"/>')
        + p(run('BODYWITHNOTE') + '<w:r><w:footnoteReference w:id="1"/></w:r>')
        + p(run('PLAINPARA ') + pictureRun('rP', 3, alt && `${alt}PLAIN`)),
        {
            rels: [['rP', 'image', 'media/image1.png'], ['rS', 'styles', 'styles.xml'], ['rN', 'numbering', 'numbering.xml']],
            media: { 'image1.png': pngOf(40, 20) },
            extra: { 'word/styles.xml': styles, 'word/numbering.xml': numbering },
            footnotes: `<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r>${run(' NOTETEXT ')}${pictureRun('rFP', 4, alt && `${alt}NOTE`)}</w:p></w:footnote>`,
            footRels: [['rFP', 'image', 'media/image1.png']],
        });
};

/**
 * Native PDF draws a picture in a heading, a list item and a footnote in its line, as in a paragraph: each on its
 * line's baseline, right of its text, the footnote's below the body; with alt text the picture, not the words; a
 * picture with no data in a list item is reported (its alt text printed where it has one); a list item's marker on
 * the baseline of a line made tall by a picture, and moved with it to the next page; 'image+ocr-text' a heading
 * picture's recognized text below the heading; a picture in a heading in a nested table's cell, drawn as flat text
 * with nothing to stand in for it, reported; n list items with a picture each laid out in linear time. A pictured list
 * item moved to the next column takes its marker there; a pictured heading or item moved to the next page takes its
 * bookmark and its link target there, and a text-only heading keeps both where it starts; a picture in a list inside
 * an item's paragraph is drawn.
 */
export async function testNativePdfLinePictures(): Promise<void> {
    console.log('\n=== Native PDF: pictures in headings, list items and notes drawn in their line ===');
    for (const alt of ['', 'ALTTEXT']) {
        const ast = await OfficeParser.parseOffice(c5Lines(alt), { extractAttachments: true });
        const { value, warnings } = await writeWithWarnings(ast, 'pdf');
        assert.deepStrictEqual(warnings.map(w => w.details), [], `Line pictures${alt ? ' with alt text' : ''}: no warning`);
        const { pages, text } = await pdfLayout(value);
        const page = pages[0];
        assert.strictEqual(page.images.length, 4, `Line pictures${alt ? ' with alt text' : ''}: four pictures drawn (${page.images.length})`);
        for (const word of ['LISTITEM', 'HEADINGTEXT', 'PLAINPARA']) {
            const label = page.texts.find(t => t.text.includes(word))!;
            const picture = page.images.find(i => Math.abs(i.bottom - label.baseline) < 4 && i.x > label.x);
            assert.ok(picture, `Line pictures: the ${word} picture stands on its line's baseline, right of its text (${JSON.stringify(label)} ${JSON.stringify(page.images)})`);
        }
        const body = Math.max(...page.texts.filter(t => /PLAINPARA|BODYWITHNOTE/.test(t.text)).map(t => t.bottom));
        const note = page.texts.find(t => t.text.includes('NOTETEXT'))!;
        assert.ok(note.y > body && page.images.some(i => i.y > body && Math.abs(i.bottom - note.baseline) < 4), 'Line pictures: the footnote\'s picture in its line, below the body');
        assert.ok(!text.includes('ALTTEXT'), 'Line pictures: the alt text is not printed');
    }

    // A picture with no data in a list item: reported; its alt text printed where it has one.
    const noData = (altText?: string) => astOf([{ type: 'list', metadata: { listType: 'ordered', itemIndex: 0 }, children: [{ type: 'text', text: 'ITEMWORDS ' }, { type: 'image', text: '', metadata: { attachmentName: 'gone.png', ...(altText ? { altText } : {}) } }] }] as any,
        [{ type: 'image', name: 'gone.png', extension: 'png', mimeType: 'image/png', data: '' }] as any);
    for (const altText of [undefined, 'ALTGONE']) {
        const { value, warnings } = await writeWithWarnings(noData(altText), 'pdf');
        assert.ok(warnings.some(w => w.code === 'IMAGE_PROCESSING_FAILED' && w.details?.name === 'gone.png'), 'Line pictures: a picture with no data in a list item reported');
        assert.strictEqual((await pdfLayout(value)).text.includes('ALTGONE'), !!altText, 'Line pictures: its alt text printed where it has one');
    }

    // A list item whose picture is 200pt tall: its marker on that line's baseline; near the page's foot, on the next page with it.
    const tall = pngOf(10, 20);
    const tallItem = { type: 'list', metadata: { listType: 'ordered', itemIndex: 0 }, children: [{ type: 'text', text: 'TALLITEM ' }, { type: 'image', text: '', metadata: { attachmentName: 't.png', width: '100pt', height: '200pt' } }] };
    const tallAst = (before: number) => astOf([
        ...Array.from({ length: before }, (_, i) => ({ type: 'paragraph', children: [{ type: 'text', text: `FILLER${i}` }] })), tallItem,
    ] as any, [{ type: 'image', name: 't.png', extension: 'png', mimeType: 'image/png', data: Buffer.from(tall).toString('base64') }] as any);
    // (After 30 paragraphs the item, without its picture, still starts on the first page, near its foot.)
    const plain = astOf([...tallAst(30).content.slice(0, 30), { ...tallItem, children: [tallItem.children[0]] }] as any);
    const plainPages = (await pdfLayout((await writeWithWarnings(plain, 'pdf')).value)).pages;
    assert.ok(plainPages[0].texts.some(t => t.text.trim().startsWith('1.')) && plainPages[0].texts.at(-1)!.bottom > 600, 'Line pictures: the item without its picture starts near the first page\'s foot');
    for (const before of [0, 30]) {
        const { pages } = await pdfLayout((await writeWithWarnings(tallAst(before), 'pdf')).value);
        const at = pages.findIndex(p => p.images.length);
        const picture = pages[at].images[0];
        const marker = pages[at].texts.find(t => t.text.trim().startsWith('1.'));
        assert.ok(marker && Math.abs(marker.baseline - picture.bottom) < 2 && Math.round(picture.h) === 200, `Line pictures: the marker on the tall picture's baseline (${before} paragraphs before; ${JSON.stringify(marker)} ${JSON.stringify(picture)})`);
        if (before) assert.ok(at > 0 && !pages.slice(0, at).some(p => p.texts.some(t => t.text.trim().startsWith('1.'))), 'Line pictures: an item near the foot moves to the next page with its marker');
    }

    // 'image+ocr-text': a heading's picture's recognized text below the heading.
    const ocrAst = astOf([{ type: 'heading', metadata: { level: 1 }, children: [{ type: 'text', text: 'HEADWORDS ' }, { type: 'image', text: 'OCRWORDS', metadata: { attachmentName: 'p.png' } }] }] as any,
        [{ type: 'image', name: 'p.png', extension: 'png', mimeType: 'image/png', data: Buffer.from(pngOf(40, 20)).toString('base64') }] as any);
    const ocrLayout = await pdfLayout((await writeWithWarnings(ocrAst, 'pdf', { includeImages: 'image+ocr-text' })).value);
    const head = ocrLayout.pages[0].texts.find(t => t.text.includes('HEADWORDS'))!, ocr = ocrLayout.pages[0].texts.find(t => t.text.includes('OCRWORDS'));
    assert.ok(ocr && ocr.y > head.bottom && ocrLayout.pages[0].images.length === 1, 'Line pictures: a heading picture\'s recognized text below the heading');

    // A heading in a nested table's cell (flat text, where no picture is embedded) holding a picture with no alt text.
    const cell = (children: any[]) => ({ type: 'cell', metadata: { row: 0, col: 0 }, children });
    const nested = astOf([{ type: 'table', children: [{ type: 'row', children: [cell([{ type: 'table', children: [{ type: 'row', children: [cell([{ type: 'heading', metadata: { level: 2 }, children: [{ type: 'text', text: 'CELLHEAD ' }, { type: 'image', text: '', metadata: { attachmentName: 'image1.png' } }] }])] }] }])] }] }] as any,
        [{ type: 'image', name: 'image1.png', extension: 'png', mimeType: 'image/png', data: Buffer.from(pngOf(4, 4)).toString('base64') }] as any);
    const nestedOut = await writeWithWarnings(nested, 'pdf');
    assert.ok(nestedOut.warnings.some(w => w.code === 'CONTENT_NOT_REPRESENTABLE' && w.details?.reason === 'image-in-text-line-dropped' && w.details?.name === 'image1.png'
        && w.details?.feature === "picture 'image1.png' in a line of text (no alternative text to draw in its place)"), `Line pictures: a picture left out of a nested cell's heading reported (${JSON.stringify(nestedOut.warnings.map(w => w.details))})`);

    // A numbered item holding a 30 x 300 pt picture after column 1 is nearly full, in two columns: the line moves to
    // column 2, and its marker with it, on the picture's baseline.
    const tallPng = { type: 'image', name: 'tall.png', extension: 'png', mimeType: 'image/png', data: Buffer.from(pngOf(40, 400)).toString('base64') };
    const filler = (n: number) => Array.from({ length: n }, (_, i) => ({ type: 'paragraph', children: [{ type: 'text', text: `Filler line ${i}` }] }));
    const pictured = (word: string, anchor: string, kind: 'list' | 'heading') => ({ type: kind, metadata: kind === 'list' ? { listType: 'ordered', itemIndex: 0, indentation: 0, anchorIds: [anchor] } : { level: 1, anchorIds: [anchor] }, children: [{ type: 'text', text: `${word} ` }, { type: 'image', text: '', metadata: { attachmentName: 'tall.png' } }] });
    const columns = await writeWithWarnings(astOf([{ type: 'break', metadata: { breakType: 'section', section: 0, pageSetup: { columns: { count: 2, gap: 36 } } } }, ...filler(25), pictured('COLITEM', 'c1', 'list')] as any, [tallPng] as any), 'pdf');
    const columnsPage = (await pdfLayout(columns.value)).pages.find(p => p.images.length)!;
    const pageWidth = (await PDFDocument.load(bytesOf(columns.value))).getPage(0).getWidth();
    const colMarker = columnsPage?.texts.find(t => t.text.trim() === '1.' || t.text.trim().startsWith('1. '));
    const colPicture = columnsPage?.images[0];
    // (Two columns 36 apart, the margins alike: column 2 starts 18 right of the page's middle.)
    assert.ok(columns.warnings.length === 0 && colMarker && colPicture && colMarker.x >= pageWidth / 2 + 18 - 0.5 && colMarker.x < colPicture.x && Math.abs(colMarker.baseline - colPicture.bottom) < 4,
        `Line pictures: a pictured item moved to column 2 takes its marker there (${JSON.stringify(colMarker)} ${JSON.stringify(colPicture)}, page ${pageWidth} wide)`);

    // A heading, and a list item, holding the picture after 25 lines: moved to page 2, with its bookmark and the
    // target of a link to it; a text-only heading at a page's foot keeps both where it starts (the next page's top).
    const targets = async (value: unknown) => {
        const doc = await PDFDocument.load(bytesOf(value));
        const pageOf = new Map(doc.getPages().map((p, i) => [p.ref.toString(), i + 1]));
        const at = (dest: PDFArray) => ({ page: pageOf.get(dest.get(0).toString()), top: Number(dest.get(3).toString()) });
        const outline: { title: string; page?: number; top: number }[] = [];
        const outlines = doc.catalog.lookupMaybe(PDFName.of('Outlines'), PDFDict);
        for (let item = outlines?.lookupMaybe(PDFName.of('First'), PDFDict); item; item = item.lookupMaybe(PDFName.of('Next'), PDFDict)) {
            const title = item.lookup(PDFName.of('Title'));
            outline.push({ title: title instanceof PDFHexString || title instanceof PDFString ? title.decodeText() : '', ...at(item.lookup(PDFName.of('Dest'), PDFArray)) });
        }
        const links = doc.getPages().flatMap(page => (page.node.Annots()?.asArray() ?? []).map(ref => doc.context.lookup(ref, PDFDict)).filter(a => a.get(PDFName.of('Subtype')) === PDFName.of('Link') && a.get(PDFName.of('Dest'))).map(a => at(a.lookup(PDFName.of('Dest'), PDFArray))));
        return { outline, links, height: doc.getPage(0).getHeight() };
    };
    const linkTo = (anchor: string) => ({ type: 'paragraph', children: [{ type: 'text', text: 'GO THERE', metadata: { link: `#${anchor}` } }] });
    for (const kind of ['heading', 'list'] as const) {
        const word = kind === 'heading' ? 'BOTTOMHEAD' : 'BOTTOMITEM', anchor = kind === 'heading' ? 'h1' : 'li1';
        const moved = await writeWithWarnings(astOf([linkTo(anchor), ...filler(24), pictured(word, anchor, kind), ...filler(3)] as any, [tallPng] as any), 'pdf', { pdfConfig: { outline: true } });
        const layout = (await pdfLayout(moved.value)).pages;
        const pictureAt = layout.findIndex(p => p.images.length) + 1, picture = layout[pictureAt - 1].images[0];
        const found = await targets(moved.value);
        const lineTop = found.height - picture.y;
        assert.ok(pictureAt === 2 && layout[1].texts.some(t => t.text.includes(word)), `Line pictures: the pictured ${kind} drawn on page 2 (page ${pictureAt})`);
        assert.ok(found.links.length === 1 && found.links[0].page === 2 && Math.abs(found.links[0].top - lineTop) <= 2, `Line pictures: a link to the pictured ${kind} goes to page 2, its line's top (${JSON.stringify(found.links)}, top ${lineTop})`);
        if (kind === 'heading') assert.ok(found.outline.length === 1 && found.outline[0].title === word && found.outline[0].page === 2 && Math.abs(found.outline[0].top - lineTop) <= 2, `Line pictures: the pictured heading's bookmark on page 2, its line's top (${JSON.stringify(found.outline)}, top ${lineTop})`);
    }
    const edge = await writeWithWarnings(astOf([linkTo('h2'), ...filler(35), { type: 'heading', metadata: { level: 1, anchorIds: ['h2'] }, children: [{ type: 'text', text: 'EDGEHEAD ' }] }, ...filler(3)] as any), 'pdf', { pdfConfig: { outline: true } });
    const edgeFound = await targets(edge.value);
    const edgePage = (await pdfLayout(edge.value)).pages.findIndex(p => p.texts.some(t => t.text.includes('EDGEHEAD'))) + 1;
    // (Its place as before pictures were drawn in headings: the top margin of the page it moved to, 48 points.)
    assert.deepStrictEqual([edgePage, edgeFound.outline, edgeFound.links], [2, [{ title: 'EDGEHEAD', page: 2, top: edgeFound.height - 48 }], [{ page: 2, top: edgeFound.height - 48 }]], 'Line pictures: a text-only heading at a page\'s foot keeps its bookmark and link target at the next page\'s top');

    // A picture in a list inside a list item's paragraph (no alt text): drawn, unreported.
    const listInList = await writeWithWarnings(astOf([{ type: 'list', metadata: { listType: 'unordered' }, children: [{ type: 'paragraph', children: [{ type: 'text', text: 'OUT ' }, { type: 'list', metadata: { listType: 'unordered' }, children: [{ type: 'text', text: 'IN ' }, { type: 'image', text: '', metadata: { attachmentName: 'p.png' } }] }] }] }] as any,
        [{ type: 'image', name: 'p.png', extension: 'png', mimeType: 'image/png', data: Buffer.from(pngOf(2, 2)).toString('base64') }] as any), 'pdf');
    assert.deepStrictEqual([listInList.warnings.map(w => w.details), (await pdfPictures(listInList.value)).images.length], [[], 1], 'Line pictures: a picture in a list inside an item\'s paragraph drawn, unreported');

    // Linear: n list items each with the one picture.
    const items = (n: number) => astOf(Array.from({ length: n }, (_, i) => ({ type: 'list', metadata: { listType: 'unordered', itemIndex: i }, children: [{ type: 'text', text: `ITEM${i} ` }, { type: 'image', text: '', metadata: { attachmentName: 'p.png' } }] })) as any,
        [{ type: 'image', name: 'p.png', extension: 'png', mimeType: 'image/png', data: Buffer.from(pngOf(8, 8)).toString('base64') }] as any);
    const built = new Map<number, OfficeParserAST>();
    await linearTime('Line pictures: list items with a picture each', n => writeWithWarnings(built.get(n) ?? built.set(n, items(n)).get(n)!, 'pdf'), 4000);
}

// ─── Each equation written as its LaTeX source reported ──────────────────

/** The ODF namespaces an in-memory package's content uses. */
const ODF_NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"';

/** An ODT of `body` (office:text's content) and other parts. */
export const odtOf = (body: string, parts: Record<string, string> = {}): Buffer => Buffer.from(zipSync({
    mimetype: [strToU8('application/vnd.oasis.opendocument.text'), { level: 0 }],
    'content.xml': strToU8(`<?xml version="1.0"?><office:document-content ${ODF_NS}><office:body><office:text>${body}</office:text></office:body></office:document-content>`),
    ...Object.fromEntries(Object.entries(parts).map(([k, v]) => [k, strToU8(v)])),
}));

/** The math-as-source warnings of a native PDF of `ast`, as [latex, kind], and all its warnings. */
const mathWarnings = async (ast: OfficeParserAST, config: Record<string, unknown> = {}) => {
    const { warnings } = await writeWithWarnings(ast, 'pdf', config);
    const reported = warnings.filter(w => w.code === 'CONTENT_NOT_REPRESENTABLE' && w.details?.reason === 'math-as-source');
    for (const w of reported) {
        assert.strictEqual(w.details.format, 'pdf', 'Equations reported: for pdf');
        assert.ok(w.details.feature.startsWith('equation \'') && w.details.feature.endsWith('\' (drawn as its LaTeX source)'), `Equations reported: the feature names the equation (${w.details.feature})`);
    }
    return { reported: reported.map(w => [w.details.latex, w.details.math] as [string, string]), warnings };
};

/** The equations (math code nodes) of an AST, as [LaTeX, kind]. */
const equationsOf = (ast: OfficeParserAST) => allNodes(ast).filter(n => n.type === 'code' && (n.metadata as any)?.math).map(n => [n.text ?? '', (n.metadata as any).math] as [string, string]);

/**
 * Native PDF reports each equation it writes as its LaTeX source (CONTENT_NOT_REPRESENTABLE, reason
 * `math-as-source`, `details.latex` and `details.math`): DOCX OMML display and inline, ODT MathML, Markdown inline and
 * display, LaTeX; in a table cell, a heading and a footnote. The same equation twice gives one; onNode's string for
 * it gives none; one of 5,000 characters is named by 80 of them; one whose picture cannot be embedded is reported by
 * that failure alone, in a paragraph, heading, list item or footnote. LibreOffice's RTF fixture, whose equations are
 * drawn as their own pictures, gives none, and one each with includeImages 'none'.
 */
export async function testNativePdfEquationsReported(): Promise<void> {
    console.log('\n=== Native PDF: each equation written as its LaTeX source reported ===');
    const MATH = '<m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r><m:r><m:t>=</m:t></m:r><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara>';
    const INLINE = '<m:oMath><m:r><m:t>y</m:t></m:r><m:sSup><m:e><m:r><m:t>z</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>';
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>`;
    const cellMath = '<m:oMath><m:r><m:t>c</m:t></m:r></m:oMath>', headMath = '<m:oMath><m:r><m:t>h</m:t></m:r></m:oMath>', noteMath = '<m:oMath><m:r><m:t>n</m:t></m:r></m:oMath>';
    const docx = docxOf(
        para(run('BEFOREMATH')) + `<w:p>${MATH}</w:p>` + para(run('INLINE '), INLINE, run(' AFTERMATH'))
        + `<w:tbl><w:tr><w:tc><w:p>${cellMath}</w:p></w:tc></w:tr></w:tbl>`
        + `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${run('HEAD ')}${headMath}</w:p>`
        + para(run('NOTED'), '<w:r><w:footnoteReference w:id="1"/></w:r>'),
        { rels: [['rS', 'styles', 'styles.xml']], extra: { 'word/styles.xml': styles }, footnotes: `<w:footnote w:id="1"><w:p>${run('NOTE ')}${noteMath}</w:p></w:footnote>` });
    const odt = odtOf('<text:p>BEFORE <draw:frame text:anchor-type="as-char"><draw:object xlink:href="./Object 1"/></draw:frame> AFTER</text:p>',
        { 'Object 1/content.xml': '<?xml version="1.0"?><math xmlns="http://www.w3.org/1998/Math/MathML"><semantics><mrow><msqrt><mi>x</mi></msqrt></mrow></semantics></math>' });
    const sources: [string, OfficeParserAST][] = [
        ['DOCX', await OfficeParser.parseOffice(docx, { extractAttachments: true })],
        ['ODT', await OfficeParser.parseOffice(odt, { extractAttachments: true })],
        ['Markdown', await OfficeParser.parseOffice(Buffer.from('Inline $a^2$ here.\n\n$$\\int x$$\n\n| col |\n| --- |\n| $q$ |\n'), { fileType: 'md' } as any)],
        ['LaTeX', await OfficeParser.parseOffice(Buffer.from('\\documentclass{article}\\begin{document}\nSome \\(x\\) text.\n\\section{Head $s$}\n\\end{document}\n'), { fileType: 'tex' } as any)],
    ];
    let total = 0;
    for (const [label, ast] of sources) {
        const equations = equationsOf(ast);
        assert.ok(equations.length >= 1, `Equations reported: ${label} has equations (${JSON.stringify(equations)})`);
        const distinct = [...new Map(equations.map(e => [e.join('\u0000'), e])).values()];
        const { reported } = await mathWarnings(ast);
        assert.deepStrictEqual(reported.map(r => r.join('|')).sort(), distinct.map(e => e.join('|')).sort(), `Equations reported: ${label}, one warning for each distinct equation, its LaTeX and kind`);
        total += reported.length;
    }
    // DOCX: the display and inline equations, and those in a cell, a heading and a footnote, each reported.
    const docxReported = (await mathWarnings(sources[0][1])).reported.map(r => r[0]);
    assert.deepStrictEqual([...docxReported].sort(), ['c', 'h', 'n', 'x=\\frac{a}{b}', 'yz^2'].sort(), `Equations reported: DOCX display, inline, cell, heading and footnote (${JSON.stringify(docxReported)})`);

    // The same equation twice: one warning. onNode's string for it: none. One of 5,000 characters: named by 80.
    const twice = astOf([{ type: 'paragraph', children: [{ type: 'text', text: 'A ' }, { type: 'code', text: 'e=mc^2', metadata: { math: 'inline' } }] }, { type: 'paragraph', children: [{ type: 'code', text: 'e=mc^2', metadata: { math: 'inline' } }] }] as any);
    assert.deepStrictEqual((await mathWarnings(twice)).reported, [['e=mc^2', 'inline']], 'Equations reported: the same equation twice, once');
    const replaced = await mathWarnings(twice, { onNode: (n: OfficeContentNode) => n.type === 'code' ? 'EQUATIONWORDS' : undefined });
    assert.deepStrictEqual(replaced.reported, [], 'Equations reported: onNode\'s string for an equation, none');
    const long = 'x+'.repeat(2500);
    const longOut = await writeWithWarnings(astOf([{ type: 'paragraph', children: [{ type: 'code', text: long, metadata: { math: 'block' } }] }, { type: 'code', text: `${long}y`, metadata: { math: 'block' } }] as any), 'pdf');
    const longWarnings = longOut.warnings.filter(w => w.details?.reason === 'math-as-source');
    assert.strictEqual(longWarnings.length, 2, 'Equations reported: a long equation in a paragraph and a block one');
    for (const w of longWarnings) {
        const named = /^equation '(.*)' \(drawn as its LaTeX source\)$/s.exec(w.details.feature)![1];
        assert.ok(named.length <= 80 && long.startsWith(named.replace(/\.\.\.$/, '')), `Equations reported: a 5,000-character equation named by at most 80 of them (${named.length})`);
        assert.ok(w.details.latex.length === 1000 && w.details.length > 1000, 'Equations reported: its source held to 1,000 characters, its length given');
    }

    // An equation whose picture is no picture, in a paragraph, a heading, a list item and a footnote: its picture's
    // failure reported once, and its LaTeX written in the line, not reported again.
    const badPicture = [{ type: 'image', name: 'bad.png', extension: 'png', mimeType: 'image/png', data: Buffer.from('not a picture at all').toString('base64') }];
    const eq = () => ({ type: 'code', text: 'x^2', metadata: { math: 'inline', fallbackAttachmentName: 'bad.png' } });
    const failedShapes: [string, any[]][] = [
        ['paragraph', [{ type: 'paragraph', children: [{ type: 'text', text: 'P ' }, eq()] }]],
        ['heading', [{ type: 'heading', metadata: { level: 1 }, children: [{ type: 'text', text: 'H ' }, eq()] }]],
        ['list item', [{ type: 'list', metadata: { listType: 'ordered' }, children: [{ type: 'text', text: 'L ' }, eq()] }]],
        ['footnote', [{ type: 'paragraph', children: [{ type: 'text', text: 'CITED', notes: [{ type: 'note', metadata: { noteType: 'footnote', noteId: '1' }, children: [{ type: 'paragraph', children: [{ type: 'text', text: 'N ' }, eq()] }] }] }] }]],
    ];
    for (const [label, content] of failedShapes) {
        const { value, warnings } = await writeWithWarnings(astOf(content, badPicture as any), 'pdf');
        assert.deepStrictEqual(warnings.map(w => w.code), ['IMAGE_PROCESSING_FAILED'], `Equations reported: one whose picture fails, in a ${label}, warned once (${JSON.stringify(warnings.map(w => w.details?.reason))})`);
        assert.ok((await pdfLayout(value)).text.includes('x^2'), `Equations reported: one whose picture fails, in a ${label}, written as its LaTeX`);
    }

    // LibreOffice's RTF: its equations drawn as their pictures, unreported; without pictures, each reported.
    const lo = await OfficeParser.parseOffice(fs.readFileSync(path.join(__dirname, '..', 'files', 'objects', 'equations-libreoffice.rtf')), { extractAttachments: true });
    const loDrawn = await mathWarnings(lo);
    assert.deepStrictEqual(loDrawn.reported, [], 'Equations reported: LibreOffice\'s equations drawn as their pictures, none reported');
    assert.strictEqual((await pdfPictures((await writeWithWarnings(lo, 'pdf')).value)).images.length, 2, 'Equations reported: LibreOffice\'s two pictures drawn');
    const loLatex = await mathWarnings(lo, { includeImages: 'none' });
    assert.deepStrictEqual(loLatex.reported.map(r => r[0]).sort(), equationsOf(lo).map(e => e[0]).sort(), 'Equations reported: with includeImages none, each of LibreOffice\'s equations reported');
    console.log(`  ✓ ${total} equations from DOCX, ODT, Markdown and LaTeX each reported once; repeats, onNode strings and pictures not; long ones named by 80 characters`);
}

// ─── ODF chart blank categories and values keep their positions ──

/** A chart's embedded table as LibreOffice writes one: a header row of series names, then a row per category. */
const odfChartTable = (series: string[], rows: [string, (string | null)[]][]) => {
    const cell = (text: string) => `<table:table-cell office:value-type="string"><text:p>${text}</text:p></table:table-cell>`;
    // (A blank value as LibreOffice writes it: a float of NaN.)
    const value = (v: string | null) => v === null ? '<table:table-cell office:value-type="float" office:value="NaN"><text:p>NaN</text:p></table:table-cell>' : `<table:table-cell office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>`;
    return `<table:table table:name="local-table"><table:table-header-columns><table:table-column/></table:table-header-columns><table:table-columns><table:table-column table:number-columns-repeated="${series.length}"/></table:table-columns>`
        + `<table:table-header-rows><table:table-row><table:table-cell><text:p/></table:table-cell>${series.map(cell).join('')}</table:table-row></table:table-header-rows>`
        + `<table:table-rows>${rows.map(([label, values]) => `<table:table-row>${label ? cell(label) : '<table:table-cell office:value-type="string"><text:p/></table:table-cell>'}${values.map(value).join('')}</table:table-row>`).join('')}</table:table-rows></table:table>`;
};

/** An ODT holding a bar chart (as an object frame) of `table`. */
const odtChartOf = (table: string) => odtOf('<text:p>BEFORE</text:p><text:p><draw:frame text:anchor-type="paragraph" svg:width="4in" svg:height="2in"><draw:object xlink:href="./Object 1"/></draw:frame></text:p><text:p>AFTER</text:p>', {
    'Object 1/content.xml': `<?xml version="1.0"?><office:document-content ${ODF_NS}><office:body><office:chart><chart:chart chart:class="chart:bar"><chart:title><text:p>BlankChart</text:p></chart:title><chart:plot-area><chart:axis chart:dimension="x" chart:name="primary-x"/><chart:axis chart:dimension="y" chart:name="primary-y"/><chart:series/><chart:series/></chart:plot-area>${table}</chart:chart></office:chart></office:body></office:document-content>`,
});

/**
 * An ODF chart's embedded table with a blank category and a blank value (LibreOffice's NaN): each label and value
 * keeps its point's position (`labelIndexes`, `valueIndexes`, `pointCount`), so the chart's table pairs every value
 * with its own category and shows a blank where a value is blank, never "NaN"; native PDF draws it, with no NaN.
 */
export async function testOdfChartBlankPoints(): Promise<void> {
    console.log('\n=== ODF charts: blank categories and values keep their positions ===');
    const odt = odtChartOf(odfChartTable(['SeriesA', 'SeriesB'], [['Jan', ['10', '1']], ['', ['20', '2']], ['Mar', [null, '3']], ['Apr', ['40', '4']]]));
    const ast = await OfficeParser.parseOffice(odt, { extractAttachments: true });
    const data = ast.attachments.find(a => a.type === 'chart')?.chartData as any;
    assert.ok(data, 'ODF blank points: the chart read');
    assert.deepStrictEqual([data.labels, data.labelIndexes, data.pointCount], [['Jan', 'Mar', 'Apr'], [0, 2, 3], 4], 'ODF blank points: the labels at their positions');
    assert.deepStrictEqual(data.dataSets.map((s: any) => [s.name, s.values, s.valueIndexes]), [['SeriesA', ['10', '20', '40'], [0, 1, 3]], ['SeriesB', ['1', '2', '3', '4'], [0, 1, 2, 3]]], 'ODF blank points: the values at their positions, the blank left out');
    assert.ok(!JSON.stringify(data).includes('NaN'), 'ODF blank points: no NaN in the chart\'s data');

    // The chart's table (Markdown writes it as a pipe table): each row its category and its own values.
    const md = String((await writeWithWarnings(ast, 'md')).value);
    const rows = md.split('\n').filter(l => l.startsWith('|') && !/^\|[\s:|-]+\|$/.test(l)).map(l => l.slice(1, -1).split('|').map(c => c.trim()));
    assert.deepStrictEqual(rows, [['', 'SeriesA', 'SeriesB'], ['Jan', '10', '1'], ['', '20', '2'], ['Mar', '', '3'], ['Apr', '40', '4']], `ODF blank points: the chart's table pairs each value with its category (${JSON.stringify(rows)})`);
    const pdf = await writeWithWarnings(ast, 'pdf');
    const back = await OfficeParser.parseOffice(Buffer.from(bytesOf(pdf.value)));
    const text = allNodes(back).filter(n => !n.children?.length).map(n => n.text ?? '').join('\n');
    assert.ok(text.includes('BlankChart') && text.includes('Apr') && !text.includes('NaN'), `ODF blank points: native PDF draws the chart with no NaN (${JSON.stringify(text)})`);

    // A chart with no blank point: every position present, as before.
    const full = await OfficeParser.parseOffice(odtChartOf(odfChartTable(['S'], [['A', ['1']], ['B', ['2']]])), { extractAttachments: true });
    const fullData = full.attachments.find(a => a.type === 'chart')?.chartData as any;
    assert.deepStrictEqual([fullData.labels, fullData.labelIndexes, fullData.pointCount, fullData.dataSets[0].values, fullData.dataSets[0].valueIndexes], [['A', 'B'], [0, 1], 2, ['1', '2'], [0, 1]], 'ODF blank points: a chart with none keeps every position');
    console.log('  ✓ blank categories and NaN values are gaps at their points, in the data, the chart\'s table and native PDF');
}

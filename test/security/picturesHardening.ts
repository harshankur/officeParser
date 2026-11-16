/**
 * Hostile pictures in the native PDF engine's readers: a hostile TIFF in a DOCX (many strips sharing their bytes, a
 * tile far larger than its image) refused by native PDF and RTF quickly and within its bytes, a metafile of a million
 * records, an EMF+ object whose parts claim 4 GB, and pictures whose reading must stay linear in their size.
 */
import * as fs from 'fs';
import * as path from 'path';
import * as zlib from 'zlib';
import { strToU8, zlibSync } from 'fflate';
import { OfficeParser } from '../../src/OfficeParser';
import { readMetafile } from '../../src/utils/metafileUtils';
import { crc32 } from '../../src/utils/imageUtils';
import { scaledMs, withinTime } from '../helpers/timing';
import { bmpOf, dibOfFile, emfOf, emfPlus, GIF_1x1, JPEG_192x48, metaStretchDib, onePicture, pdfPictures, plusDrawImage, plusHeader, plusImage, rawTiff, stretchDibits, tiffOf, wmfOf, writeWithWarnings } from '../parser/picturesAndLines';
import { docxOf, para, pictureRun, run } from '../parser/wordRtfContent';

type Check = (name: string, cond: boolean, detail?: string) => void;

/** The memory in use: the heap and the buffers outside it (a typed array's bytes are not in the heap). */
const memoryInUse = () => { const m = process.memoryUsage(); return m.heapUsed + m.arrayBuffers; };

/** An 8-bit grey TIFF `width` pixels wide (64 unless given) of `n` one-row strips compressed as `compression`, strip i at `offset(i)` in `data`, `count` bytes long. */
const STRIP_W = 64;
const stripsTiff = (compression: number, data: Uint8Array, n: number, offset: (i: number) => number, count: number, width = STRIP_W): Uint8Array =>
    rawTiff([[256, 3, [width]], [257, 4, [n]], [258, 3, [8]], [259, 3, [compression]], [262, 3, [1]], [273, 4, Array.from({ length: n }, (_, i) => 8 + offset(i))], [277, 3, [1]], [278, 4, [1]], [279, 4, new Array(n).fill(count)]], data);

/** `bytes` of data that make no pixels in `compression`: PackBits no-ops, LZW clear codes, empty stored Deflate blocks after a zlib header, or zeros. */
const noProgress = (compression: number, bytes: number): Uint8Array => {
    const data = new Uint8Array(bytes);
    if (compression === 32773) data.fill(0x80);
    // (9-bit clear codes, 256, most significant bit first.)
    else if (compression === 5) for (let bit = 0; bit + 9 <= bytes * 8; bit += 9) data[bit >> 3] |= 0x80 >> (bit & 7);
    else if (compression === 8) { data.set([0x78, 0x01]); for (let i = 2; i + 5 <= bytes; i += 5) data.set([0, 0, 0, 0xFF, 0xFF], i); }
    return data;
};

/** Writes a zlib stream's bits after its two-byte header: `put` a value's low bits least significant first, `code` a Huffman code most significant first, `align` to the next byte. */
interface DeflateWriter { out: number[]; put(value: number, n: number): void; code(value: number, n: number): void; align(): void }

/** A zlib stream's header, and a writer for its blocks. */
const deflateWriter = (): DeflateWriter => {
    const out: number[] = [0x78, 0x01];
    let acc = 0, held = 0;
    const w: DeflateWriter = {
        out,
        put: (value, n) => { for (let i = 0; i < n; i++) { acc |= ((value >> i) & 1) << held; if (++held === 8) { out.push(acc); acc = 0; held = 0; } } },
        code: (value, n) => { for (let i = n - 1; i >= 0; i--) w.put((value >> i) & 1, 1); },
        align: () => { if (held) { out.push(acc); acc = 0; held = 0; } },
    };
    return w;
};

/** The stream `w` wrote, ended: to the next byte, then the Adler-32 of `zeros` zero bytes. */
const zlibEnd = (w: DeflateWriter, zeros: number): Uint8Array => {
    w.align();
    const adler = ((zeros % 65521) * 65536 + 1) >>> 0;
    w.out.push(adler >>> 24, (adler >>> 16) & 0xFF, (adler >>> 8) & 0xFF, adler & 0xFF);
    return new Uint8Array(w.out);
};

/**
 * A zlib stream of `count` Deflate blocks `block` writes (none of them the last), then a last fixed block: empty, or
 * (given `row`) writing `row` zero bytes, a literal and then matches at distance 1.
 */
const zlibOf = (block: (w: DeflateWriter) => void, count: number, row = 0): Uint8Array => {
    const w = deflateWriter();
    for (let i = 0; i < count; i++) block(w);
    // The last block, fixed codes: literal 0 (8 bits, 0x30), then lengths 258 (code 285) or 3 to 10 (codes 257 to
    // 264, 7 bits) at distance code 0, then its end code (seven zero bits).
    w.put(1, 1); w.put(1, 2);
    let left = row;
    if (left) { w.code(0x30, 8); left--; }
    while (left > 0) {
        if (left >= 258) { w.code(0xC5, 8); w.code(0, 5); left -= 258; }
        else if (left >= 3) { const length = Math.min(left, 10); w.code(length - 2, 7); w.code(0, 5); left -= length; }
        else { w.code(0x30, 8); left--; }
    }
    w.code(0, 7);
    return zlibEnd(w, row);
};

/**
 * A dynamic Deflate block whose header is as costly as one can be: 288 literal and 32 distance
 * lengths, every one 9 (the code length code gives 9 and 16 one bit each; a 9, then repeats of it by code 16), and
 * only its end code: about 26 bytes that define 320 lengths and build three codes.
 */
const headerBlock = (w: DeflateWriter) => {
    // Not the last block, dynamic; 288 literals, 32 distances, 7 code length code lengths.
    w.put(0, 1); w.put(2, 2); w.put(31, 5); w.put(31, 5); w.put(7 - 4, 4);
    // In its order 16, 17, 18, 0, 8, 7, 9: 16 and 9 one bit each (9 is code 0, 16 code 1).
    w.put(1, 3); w.put(0, 3); w.put(0, 3); w.put(0, 3); w.put(0, 3); w.put(0, 3); w.put(1, 3);
    // A 9, then 319 more by repeats of 3 to 6 (and single 9s for what is left under 3).
    w.code(0, 1);
    for (let filled = 1; filled < 320;) {
        const repeat = Math.min(6, 320 - filled);
        if (repeat < 3) { w.code(0, 1); filled++; } else { w.code(1, 1); w.put(repeat - 3, 2); filled += repeat; }
    }
    // The end code: the 257th of 288 literals of 9 bits.
    w.code(256, 9);
};

/** A zlib stream of `count` header-heavy dynamic blocks (headerBlock), then a last empty block: no pixels. */
const headerBlocks = (count: number): Uint8Array => zlibOf(headerBlock, count);

/** A DOCX of one paragraph showing a TIFF (`image1.tiff`). */
const docxWithTiff = (tiff: Uint8Array) => docxOf(para(run('Before the picture')) + para(pictureRun('rT', 1)), { rels: [['rT', 'image', 'media/image1.tiff']], media: { 'image1.tiff': tiff } });

/** `run(n)` timed at n and at n/4: whether the larger is within 8 times the smaller plus 250 ms (measured again once when over). */
const linearIn = async (run: (n: number) => unknown, n: number): Promise<{ ok: boolean; detail: string }> => {
    const quarter = Math.floor(n / 4);
    const time = (count: number) => { const started = performance.now(); run(count); return performance.now() - started; };
    let small = time(quarter), large = time(n);
    if (large > 8 * small + await scaledMs(250)) { small = Math.min(small, time(quarter)); large = Math.min(large, time(n)); }
    return { ok: large <= 8 * small + await scaledMs(250), detail: `${Math.round(small)} ms at ${quarter}, ${Math.round(large)} ms at ${n}` };
};

/** What each output says of a TIFF picture: none draws or converts one, so a hostile TIFF is refused by its signature, unread. */
const TIFF_REFUSED = {
    pdf: 'the native PDF engine does not embed TIFF pictures (only PNG, JPEG, GIF, BMP, SVG, and bitmaps in EMF and WMF)',
    rtf: 'RTF output does not take TIFF pictures (it takes PNG, JPEG, EMF and WMF, and GIF and BMP as PNG)',
} as const;

export async function picturesHardeningTests(check: Check): Promise<void> {
    console.log('- Hostile TIFF pictures in a DOCX are refused by their signature, unread and quickly...');
    const zeros = new Uint8Array(zlib.deflateSync(Buffer.alloc(4_000_000), { level: 9 }));

    // End to end: a 2 KB DOCX whose TIFF has 20,000 strips sharing 400 KB of PackBits no-ops, to native PDF and RTF.
    const overlapped = docxWithTiff((() => { const data = noProgress(32773, 400_000); return stripsTiff(32773, data, 20_000, () => 0, data.length); })());
    for (const format of ['pdf', 'rtf'] as const) {
        const warnings: any[] = [];
        const written = await withinTime(2000, async () => {
            warnings.length = 0;
            const ast = await OfficeParser.parseOffice(overlapped, { extractAttachments: true });
            return ast.to(format, { pdfConfig: { engine: 'native' }, onWarning: (w: any) => warnings.push(w) } as any);
        });
        const refusals = warnings.filter(w => w.code === 'IMAGE_PROCESSING_FAILED');
        check(`TIFF strips sharing their bytes, in a DOCX to ${format}: refused as a TIFF, quickly`, written.within && refusals.length === 1 && refusals[0].details?.reason === TIFF_REFUSED[format], `${written.ms} ms, ${JSON.stringify(warnings.map(w => w.details?.reason))}`);
    }

    // End to end: a DOCX of about 5 KB whose TIFF (4 MB) has 500,000 one-pixel strips sharing the stream of zeros, to
    // native PDF and RTF: refused as a TIFF (none is drawn), unread, quickly.
    const bomb500k = docxWithTiff(stripsTiff(8, zeros, 500_000, () => 0, zeros.length, 1));
    for (const format of ['pdf', 'rtf'] as const) {
        const warnings: any[] = [];
        const written = await withinTime(5000, async () => {
            warnings.length = 0;
            const ast = await OfficeParser.parseOffice(bomb500k, { extractAttachments: true });
            return ast.to(format, { pdfConfig: { engine: 'native' }, onWarning: (w: any) => warnings.push(w) } as any);
        });
        const refusals = warnings.filter(w => w.code === 'IMAGE_PROCESSING_FAILED');
        const images = format === 'pdf' ? (await pdfPictures((written.result as any).value)).images : [];
        check(`TIFF of 500,000 strips sharing a stream of 4 MB of zeros, in a ${bomb500k.length}-byte DOCX to ${format}: refused as a TIFF, not drawn, quickly`, written.within && refusals.length === 1 && refusals[0].details?.reason === TIFF_REFUSED[format] && images.length === 0, `${written.ms} ms, ${JSON.stringify(warnings.map(w => w.details?.reason))}${format === 'pdf' ? `, ${JSON.stringify(images)}` : ''}`);
    }

    // 625,000 one-pixel strips sharing 157 header-heavy dynamic blocks, then 45 MB of zeros no strip reads (about 50 KB as a DOCX).
    const headers = headerBlocks(157);
    const padded = (megabytes: number) => { const data = new Uint8Array(headers.length + megabytes * 1_000_000); data.set(headers); return stripsTiff(8, data, 625_000, () => 0, headers.length, 1); };

    // Refused as a TIFF, quickly, once, to native PDF and RTF.
    {
        const mh45 = docxWithTiff(padded(45));
        for (const format of ['pdf', 'rtf'] as const) {
            const warnings: any[] = [];
            const written = await withinTime(3000, async () => {
                warnings.length = 0;
                const ast = await OfficeParser.parseOffice(mh45, { extractAttachments: true });
                return ast.to(format, { pdfConfig: { engine: 'native' }, onWarning: (w: any) => warnings.push(w) } as any);
            });
            const refusals = warnings.filter(w => w.code === 'IMAGE_PROCESSING_FAILED');
            check(`TIFF of 625,000 strips sharing header-heavy blocks, padded by 45 MB, in a ${mh45.length}-byte DOCX to ${format}: refused as a TIFF, quickly, once`, written.within && refusals.length === 1 && refusals[0].details?.reason === TIFF_REFUSED[format], `${written.ms} ms, ${JSON.stringify(warnings.map(w => w.details?.reason))}`);
        }
    }

    // A 1x1 picture of 16 samples of 16 bits in a 6000x6000 tile, in a 1.5 KB DOCX, to native PDF and RTF.
    const deep = (width: number, height: number, tile?: number): Uint8Array => {
        const blocks = tile ? Math.ceil(width / tile) * Math.ceil(height / tile) : 1;
        const fields: [number, 3 | 4, number[]][] = [[256, 4, [width]], [257, 4, [height]], [258, 3, new Array(16).fill(16)], [259, 3, [1]], [262, 3, [1]], [277, 3, [16]]];
        if (tile) fields.push([322, 4, [tile]], [323, 4, [tile]], [324, 4, new Array(blocks).fill(8)], [325, 4, new Array(blocks).fill(4)]);
        else fields.push([273, 4, [8]], [279, 4, [4]]);
        return rawTiff(fields, new Uint8Array([1, 2, 3, 4]));
    };
    const bigTile = docxWithTiff(deep(1, 1, 6000));
    for (const format of ['pdf', 'rtf'] as const) {
        const before = memoryInUse();
        const written = await withinTime(1000, async () => (await OfficeParser.parseOffice(bigTile, { extractAttachments: true })).to(format, { pdfConfig: { engine: 'native' } } as any));
        const grown = memoryInUse() - before;
        check(`TIFF: a picture in a tile far larger than it, in a DOCX to ${format}: quickly and within its bytes`, written.within && grown < 64 * 1024 * 1024, `${written.ms} ms, ${Math.round(grown / 1048576)} MB`);
    }

    console.log('- Hostile metafiles read once, by their records, within the heap...');
    // An EMF of a million records of 8 bytes (EMR_SAVEDC, which draws nothing): walked once.
    const records = (n: number) => { const out = new Uint8Array(8 * n); const v = new DataView(out.buffer); for (let i = 0; i < n; i++) { v.setUint32(i * 8, 33, true); v.setUint32(i * 8 + 4, 8, true); } return out; };
    const withRecords = (n: number) => { const head = emfOf([]); const body = records(n); const out = new Uint8Array(head.length - 20 + body.length + 20); out.set(head.subarray(0, 108)); out.set(body, 108); out.set(head.subarray(108), 108 + body.length); return out; };
    const million = withRecords(1_000_000);
    const walked = await withinTime(2000, () => readMetafile(million, 'EMF'));
    check('EMF: a million records walked quickly', walked.within && !walked.result.bitmap && walked.result.alsoDraws.length === 0, `${walked.ms} ms`);
    const built = new Map<number, Uint8Array>([[250_000, withRecords(250_000)], [1_000_000, million]]);
    const emfLinear = await linearIn(n => readMetafile(built.get(n)!, 'EMF'), 1_000_000);
    check('EMF: records read in time linear in their number', emfLinear.ok, emfLinear.detail);

    // An EMF+ Image object continued over records that claim it is 4 GB: never joined past the metafile's size.
    const u32 = (v: number) => [v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF];
    const part = (n: number): [number, number, number[]] => [0x4008, 0x8000 | 5 << 8, [...u32(0xFFFFFFF0), ...new Array(n).fill(7)]];
    const claims = emfOf(Array.from({ length: 200 }, () => emfPlus([part(4000)])));
    const heapBefore2 = memoryInUse();
    const claimed = await withinTime(1000, () => readMetafile(claims, 'EMF'));
    const grown2 = memoryInUse() - heapBefore2;
    check('EMF+: an object claiming 4 GB is not joined, nothing drawn', !claimed.result.bitmap && claimed.within && grown2 < 64 * 1024 * 1024, `${claimed.ms} ms, ${Math.round(grown2 / 1048576)} MB`);
    // The same parts claiming what they hold, together more than the file could: still bounded by the file.
    const honest = emfOf(Array.from({ length: 50 }, (_, i) => emfPlus([[0x4008, 0x8000 | 5 << 8, [...u32(50 * 4000), ...new Array(4000).fill(i)]]])));
    const joined = await withinTime(1000, () => readMetafile(honest, 'EMF'));
    check('EMF+: an object joined from its parts, read without a throw', joined.within, `${joined.ms} ms`);

    console.log('- Every picture the engine embeds within the pixel budget...');
    // Each picture declares 20,000 x 20,000 pixels in a few hundred bytes, wherever it comes from: an attachment, a
    // bitmap an EMF, WMF or EMF+ record holds, a picture in an SVG. Refused before it is decoded, reported once.
    const huge = 20_000;
    const be32 = (v: number) => [(v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255];
    const pngChunk = (type: string, data: number[]) => { const typed = [...type].map(c => c.charCodeAt(0)).concat(data); return [...be32(data.length), ...typed, ...be32(crc32(new Uint8Array(typed)))]; };
    const ihdr = (w: number, h: number) => pngChunk('IHDR', [...be32(w), ...be32(h), 8, 6, 0, 0, 0]);
    const pngFrom = (...chunks: number[][]) => new Uint8Array([0x89, 0x50, 0x4E, 0x47, 13, 10, 26, 10, ...chunks.flat(), ...pngChunk('IEND', [])]);
    const idat = pngChunk('IDAT', [...zlibSync(new Uint8Array(64))]);
    const hugePng = pngFrom(ihdr(huge, huge), idat);
    // (A JPEG's frame size at bytes 25 to 28 of its frame header.)
    const jpegOf = (w: number, h: number) => { const jpeg = JPEG_192x48.slice(); jpeg.set([h >> 8, h & 255, w >> 8, w & 255], 25); return jpeg; };
    const hugeJpeg = jpegOf(huge, huge);
    const hugeGif = GIF_1x1.slice(); hugeGif.set([huge & 255, huge >> 8, huge & 255, huge >> 8], 6);
    const hugeBmp = bmpOf(1, 1); new DataView(hugeBmp.buffer).setInt32(18, huge, true); new DataView(hugeBmp.buffer).setInt32(22, huge, true);
    const hugeTiff = rawTiff([[256, 4, [huge]], [257, 4, [huge]], [258, 3, [8]], [259, 3, [1]], [262, 3, [1]], [273, 4, [8]], [277, 3, [1]], [279, 4, [4]]], new Uint8Array(4));
    const plusOf = (file: Uint8Array) => emfOf([emfPlus([plusHeader, plusImage(1, huge, huge, file), plusDrawImage(1, huge, huge, 100, 50)])]);
    const tooLarge = `image is too large to embed (${huge}x${huge} pixels)`;
    const pictures: [string, Uint8Array, string, string][] = [
        ['an EMF of a BI_PNG bitmap', emfOf([stretchDibits(dibOfFile(hugePng, 5, huge, huge), 100, 50)]), 'a.emf', 'image/emf'],
        ['an EMF of a BI_JPEG bitmap', emfOf([stretchDibits(dibOfFile(hugeJpeg, 4, huge, huge), 100, 50)]), 'b.emf', 'image/emf'],
        ['a WMF of a BI_PNG bitmap', wmfOf([metaStretchDib(dibOfFile(hugePng, 5, huge, huge), huge, huge, 100, 50)]), 'c.wmf', 'image/wmf'],
        ['an EMF+ PNG', plusOf(hugePng), 'd.emf', 'image/emf'],
        ['an EMF+ JPEG', plusOf(hugeJpeg), 'e.emf', 'image/emf'],
        ['an EMF+ GIF', plusOf(hugeGif), 'f.emf', 'image/emf'],
        ['an EMF+ BMP', plusOf(hugeBmp), 'g.emf', 'image/emf'],
    ];
    /** Writes one picture to native PDF: its warnings, time and memory grown (heap and buffers). */
    const embedded = async (bytes: Uint8Array, name: string, mimeType: string) => {
        const before = memoryInUse();
        const written = await withinTime(1500, () => writeWithWarnings(onePicture(bytes, name, mimeType, 'ALT'), 'pdf'));
        return { ...written, grown: memoryInUse() - before };
    };
    const withinBounds = (r: { within: boolean; grown: number }) => r.within && r.grown < 128 * 1024 * 1024;
    const shown = (r: Awaited<ReturnType<typeof embedded>>) => `${r.ms} ms, ${Math.round(r.grown / 1048576)} MB, ${JSON.stringify(r.result.warnings.map((w: any) => w.details?.reason ?? w.details?.feature))}`;
    for (const [label, bytes, name, mimeType] of pictures) {
        const r = await embedded(bytes, name, mimeType);
        const failures = r.result.warnings.filter((w: any) => w.code === 'IMAGE_PROCESSING_FAILED');
        check(`Pixel budget: ${label} declaring ${huge}x${huge} is refused before decoding`, withinBounds(r) && r.result.warnings.length === 1 && failures.length === 1 && String(failures[0].details?.reason).includes(tooLarge), shown(r));
    }
    // A TIFF declaring 20,000 x 20,000 in an EMF+ record is refused as a TIFF (no output draws one), unread.
    const hugeTiffRun = await embedded(plusOf(hugeTiff), 'h.emf', 'image/emf');
    check('Pixel budget: an EMF+ TIFF declaring 20000x20000 is refused as a TIFF, unread', withinBounds(hugeTiffRun) && JSON.stringify(hugeTiffRun.result.warnings.map((w: any) => w.details?.reason)) === JSON.stringify([`the bitmap it holds cannot be drawn: ${TIFF_REFUSED.pdf}`]), shown(hugeTiffRun));
    // An SVG's picture: one of its pictures that could not be read, in the SVG's own warning.
    const svg = strToU8(`<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"><image href="data:image/png;base64,${Buffer.from(hugePng).toString('base64')}" width="40" height="20"/></svg>`);
    const svgRun = await embedded(svg, 's.svg', 'image/svg+xml');
    check('Pixel budget: a PNG inside an SVG declaring 20000x20000 is not decoded, reported as the SVG\'s', withinBounds(svgRun) && svgRun.result.warnings.length === 1 && svgRun.result.warnings[0].details?.reason === 'svg' && String(svgRun.result.warnings[0].details?.feature).includes('pictures that could not be read'), shown(svgRun));
    // A PNG whose second IHDR (the one pdf-lib's decoder takes) declares 20,000 x 20,000.
    const twoHeaders = await embedded(pngFrom(ihdr(10, 10), ihdr(huge, huge), idat), 'two.png', 'image/png');
    check('Pixel budget: a PNG with a second, larger IHDR is refused', withinBounds(twoHeaders) && JSON.stringify(twoHeaders.result.warnings.map((w: any) => w.details?.reason)) === '["its PNG data could not be read"]', shown(twoHeaders));
    // An animated PNG of 6,000 x 6,000 and 50 frames, each of which pdf-lib's decoder would decode.
    const frames: number[][] = [];
    for (let i = 0; i < 50; i++) frames.push(pngChunk('fcTL', [...be32(2 * i), ...be32(1), ...be32(1), ...be32(0), ...be32(0), 0, 1, 0, 1, 0, 0]), pngChunk('fdAT', [...be32(2 * i + 1), ...zlibSync(new Uint8Array(5))]));
    const animated = await embedded(pngFrom(ihdr(6000, 6000), pngChunk('acTL', [...be32(50), ...be32(0)]), idat, ...frames), 'anim.png', 'image/png');
    check('Pixel budget: an animated PNG is refused before its frames are decoded', withinBounds(animated) && JSON.stringify(animated.result.warnings.map((w: any) => w.details?.reason)) === '["its PNG picture is animated, which the native PDF engine does not embed"]', shown(animated));
    // A JPEG with an FFCC segment before its frame header: pdf-lib takes FFCC as its frame, so its size decides.
    const withFfcc = (w: number, h: number) => { const jpeg = JPEG_192x48; return new Uint8Array([...jpeg.subarray(0, 20), 0xFF, 0xCC, 0, 8, 8, h >> 8, h & 255, w >> 8, w & 255, 3, ...jpeg.subarray(20)]); };
    const ffccHuge = await embedded(withFfcc(huge, huge), 'cc.jpg', 'image/jpeg');
    check('Pixel budget: a JPEG whose FFCC segment reads as a 20000x20000 frame is refused', withinBounds(ffccHuge) && JSON.stringify(ffccHuge.result.warnings.map((w: any) => w.details?.reason)) === JSON.stringify([tooLarge]), shown(ffccHuge));
    const ffccSmall = await embedded(withFfcc(192, 48), 'cc.jpg', 'image/jpeg');
    const drawnSmall = ffccSmall.result.warnings.length === 0 ? (await pdfPictures(ffccSmall.result.value)).images : [];
    check('Pixel budget: the same JPEG with a small FFCC frame is drawn at that size', JSON.stringify(drawnSmall) === '[{"w":192,"h":48}]', shown(ffccSmall));
    // No path past the budget: pdf-lib's PNG and JPEG embedding is called only inside embedRaster.
    const engine = fs.readFileSync(path.join(__dirname, '..', '..', 'src', 'generators', 'pdf', 'nativePdfEngine.ts'), 'utf8');
    const start = engine.indexOf('private async embedRaster('), end = engine.indexOf('\n    }\n', start);
    const calls = [...engine.matchAll(/embed(?:Png|Jpg)\(/g)].map(m => m.index!);
    check('Pixel budget: embedPng( and embedJpg( are called only in embedRaster', start > 0 && calls.length === 2 && calls.every(at => at > start && at < end), `${calls.length} calls`);
}

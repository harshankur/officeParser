/**
 * Word and RTF content kept on the way to output: EMF and WMF pictures typed and written, RTF pictures,
 * tracked deletions, equations and embedded objects, Word content-control prompts and embedded objects.
 */
import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { HtmlGenerator } from '../../src/generators/HtmlGenerator';
import { OfficeParser } from '../../src/OfficeParser';
import type { OfficeContentNode, OfficeParserAST, OfficeParserConfig } from '../../src/types';
import { getMimeFromBytes, pictureKindOf } from '../../src/utils/imageUtils';
import { ocrTestHooks, terminateOcr } from '../../src/utils/ocrUtils';
import { metafileSize } from '../../src/utils/metafileUtils';
import { embeddedObjectKind } from '../../src/utils/oleUtils';
import { ommlToLatex } from '../../src/utils/mathUtils';
import { parseXmlString } from '../../src/utils/xmlUtils';

// ─── Shared fixtures ─────────────────────────────────────────────────────────

/** Every node of an AST, content, notes and comments, in document order. */
export const allNodes = (ast: OfficeParserAST): OfficeContentNode[] => {
    const out: OfficeContentNode[] = [];
    const pending = [...ast.content].reverse();
    while (pending.length) {
        const node = pending.pop()!;
        out.push(node);
        for (const list of [node.comments, node.notes, node.children]) if (list) for (let i = list.length - 1; i >= 0; i--) pending.push(list[i]);
    }
    return out;
};

/** Little-endian writers for building metafiles. */
const u16 = (v: number) => [v & 0xFF, (v >>> 8) & 0xFF];
const u32 = (v: number) => [v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF];

/**
 * A small EMF: an EMR_HEADER whose frame (rclFrame, hundredths of a millimetre) is `frameW` by `frameH`,
 * and an EMR_EOF. `seed` varies one reserved byte of the description so pictures differ.
 */
export const emfBytes = (frameW = 2540, frameH = 1270, seed = 0): Uint8Array => {
    const header = [
        ...u32(1), ...u32(108),                 // EMR_HEADER, its size
        ...u32(0), ...u32(0), ...u32(95), ...u32(47), // rclBounds (device units)
        ...u32(0), ...u32(0), ...u32(frameW), ...u32(frameH), // rclFrame
        0x20, 0x45, 0x4D, 0x46,                 // " EMF"
        ...u32(0x10000), ...u32(108 + 20), ...u32(2), ...u16(1), ...u16(0), // version, bytes, records, handles, reserved
        ...u32(0), ...u32(0), ...u32(0),        // nDescription, offDescription, nPalEntries
        ...u32(1920), ...u32(1080), ...u32(508), ...u32(286), // szlDevice, szlMillimeters
        ...u32(0), ...u32(0), ...u32(0),        // cbPixelFormat, offPixelFormat, bOpenGL
        ...u32(508000), ...u32(286000),         // szlMicrometers
        ...u32(seed),                            // (padding to 108)
    ];
    const eof = [...u32(14), ...u32(20), ...u32(0), ...u32(16), ...u32(20)];
    return new Uint8Array([...header, ...eof]);
};

/**
 * A WMF METAFILE with no placeable header, as RTF holds one: a METAHEADER, META_SETMAPMODE (when given),
 * META_SETWINDOWORG and META_SETWINDOWEXT (when `window` is given), a small rectangle, META_EOF.
 */
export const headerlessWmf = (window?: { x: number; y: number; w: number; h: number }, mapMode?: number): Uint8Array => {
    const records: number[] = [];
    if (mapMode !== undefined) records.push(...u32(4), ...u16(0x0103), ...u16(mapMode));
    if (window) records.push(...u32(5), ...u16(0x020B), ...u16(window.y), ...u16(window.x), ...u32(5), ...u16(0x020C), ...u16(window.h), ...u16(window.w));
    records.push(...u32(7), ...u16(0x041B), ...u16(10), ...u16(10), ...u16(0), ...u16(0)); // META_RECTANGLE
    records.push(...u32(3), ...u16(0x0000));
    const words = (18 + records.length) / 2;
    return new Uint8Array([...u16(1), ...u16(9), ...u16(0x0300), ...u32(words), ...u16(0), ...u32(7), ...u16(0), ...records]);
};

/** A placeable WMF: the 22-byte header (key, bounding box, units per inch, checksum) and the METAFILE. */
export const placeableWmf = (box: { left: number; top: number; right: number; bottom: number }, inch: number, metafile = headerlessWmf({ x: box.left, y: box.top, w: box.right - box.left, h: box.bottom - box.top })): Uint8Array => {
    const words = [0xCDD7, 0x9AC6, 0, box.left & 0xFFFF, box.top & 0xFFFF, box.right & 0xFFFF, box.bottom & 0xFFFF, inch, 0, 0];
    const checksum = words.reduce((a, w) => a ^ w, 0);
    return new Uint8Array([...words.flatMap(u16), ...u16(checksum), ...metafile]);
};

/** A 1x1 PNG. */
export const TINY_PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 'base64');

const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
export const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';

/** A DOCX of `body`, its relationships (`[id, type, target]`), media under word/media and other parts. */
export const docxOf = (body: string, opts: { rels?: [string, string, string][]; media?: Record<string, Uint8Array>; extra?: Record<string, string | Uint8Array>; footnotes?: string; footRels?: [string, string, string][] } = {}): Buffer => {
    const relXml = (rels: [string, string, string][]) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.map(([id, type, target]) => `<Relationship Id="${id}" Type="${RT}${type}" Target="${target}"/>`).join('')}</Relationships>`;
    const rels = [...(opts.footnotes ? [['rIdF', 'footnotes', 'footnotes.xml'] as [string, string, string]] : []), ...(opts.rels ?? [])];
    const files: Record<string, Uint8Array> = {
        '[Content_Types].xml': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="emf" ContentType="image/x-emf"/><Default Extension="wmf" ContentType="image/x-wmf"/><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'),
        '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RT}officeDocument" Target="word/document.xml"/></Relationships>`),
        'word/document.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W_NS}><w:body>${body}</w:body></w:document>`),
        'word/_rels/document.xml.rels': strToU8(relXml(rels)),
    };
    if (opts.footnotes) {
        files['word/footnotes.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${W_NS}>${opts.footnotes}</w:footnotes>`);
        if (opts.footRels?.length) files['word/_rels/footnotes.xml.rels'] = strToU8(relXml(opts.footRels));
    }
    for (const [name, bytes] of Object.entries(opts.media ?? {})) files['word/media/' + name] = bytes;
    for (const [name, value] of Object.entries(opts.extra ?? {})) files[name] = typeof value === 'string' ? strToU8(value) : value;
    return Buffer.from(zipSync(files));
};

/** A paragraph of runs. */
export const para = (...runs: string[]) => `<w:p>${runs.join('')}</w:p>`;
/** A run of text. */
export const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** A picture run (DrawingML inline) showing the image the relationship `rid` names. */
export const pictureRun = (rid: string, id = 1, descr = '') =>
    `<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="${id}" name="Picture ${id}"${descr ? ` descr="${descr}"` : ''}/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="p${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

/** A DOCX showing an EMF and a WMF picture (c3-pictures' shape). */
export const metafilePictures = (emf = emfBytes(), wmf = placeableWmf({ left: 0, top: 0, right: 1440, bottom: 720 }, 1440)): Buffer => docxOf(
    para(run('LABELEMF')) + para(pictureRun('rE', 1)) + para(run('LABELWMF')) + para(pictureRun('rW', 2)),
    { rels: [['rE', 'image', 'media/image1.emf'], ['rW', 'image', 'media/image2.wmf']], media: { 'image1.emf': emf, 'image2.wmf': wmf } });

/**
 * An embedded object run as Word writes one (`w:object`: the VML preview shape, its `v:imagedata`, and
 * `o:OLEObject` naming the ProgID), the c5-objects shape.
 */
export const oleRun = (shapeRid: string | undefined, oleRid: string, progId: string, alt?: string) =>
    `<w:r><w:object w:dxaOrig="1440" w:dyaOrig="720"><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:72pt;height:36pt" o:ole=""${alt !== undefined ? ` alt="${alt}"` : ''}>${shapeRid ? `<v:imagedata r:id="${shapeRid}" o:title=""/>` : ''}</v:shape><o:OLEObject Type="Embed" ProgID="${progId}" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1234" r:id="${oleRid}"/></w:object></w:r>`;

/** The c5-objects document: an Excel sheet with a PNG preview, Equation 3.0 with a WMF one, MathType with an EMF one. */
export const c5Objects = (): Buffer => docxOf(
    para(run('BEFOREOLE')) + para(oleRun('rIp', 'rO1', 'Excel.Sheet.12')) + para(run('MIDOLE')) + para(oleRun('rIw', 'rO2', 'Equation.3')) + para(oleRun('rIe', 'rO3', 'Equation.DSMT4')) + para(run('AFTEROLE')),
    {
        rels: [['rIp', 'image', 'media/image1.png'], ['rIw', 'image', 'media/image2.wmf'], ['rIe', 'image', 'media/image3.emf'], ['rO1', 'package', 'embeddings/sheet.xlsx'], ['rO2', 'oleObject', 'embeddings/oleObject1.bin'], ['rO3', 'oleObject', 'embeddings/oleObject2.bin']],
        media: { 'image1.png': TINY_PNG, 'image2.wmf': placeableWmf({ left: 0, top: 0, right: 1440, bottom: 720 }, 1440), 'image3.emf': emfBytes() },
        extra: { 'word/embeddings/sheet.xlsx': new Uint8Array([80, 75, 5, 6, ...new Array(18).fill(0)]), 'word/embeddings/oleObject1.bin': new Uint8Array(16), 'word/embeddings/oleObject2.bin': new Uint8Array(16) },
    });

/** Parses `input`, collecting the warnings' codes and the warnings themselves. */
export const parseWithWarnings = async (input: Buffer | string, config: OfficeParserConfig = {}) => {
    const warnings: any[] = [];
    const ast = await OfficeParser.parseOffice(typeof input === 'string' ? Buffer.from(input, 'latin1') : input, { ...config, onWarning: (w: any) => warnings.push(w) });
    return { ast, warnings, codes: warnings.map(w => String(w.code)) };
};

/** A generator's output as bytes. */
export const bytesOf = (value: unknown): Uint8Array => value instanceof Uint8Array ? value : Buffer.isBuffer(value) ? new Uint8Array(value) : new Uint8Array(value as ArrayBuffer);

/** A zip's entries as text, by name. */
export const zipTexts = (value: unknown): Record<string, string> => Object.fromEntries(Object.entries(unzipSync(bytesOf(value))).map(([k, v]) => [k, strFromU8(v)]));

// ─── EMF and WMF typed ───────────────────────────────────────────────────

/** EMF and WMF attachments are typed image/emf and image/wmf in every reader, and are never OCR'd. */
export async function testMetafileTyping(): Promise<void> {
    console.log('\n=== Metafiles typed image/emf and image/wmf ===');
    const types = (ast: OfficeParserAST) => ast.attachments.map(a => `${a.name}|${a.mimeType}`).join(', ');

    // The c5-objects shape: the package's media, typed by their bytes.
    const docx = await OfficeParser.parseOffice(c5Objects(), { extractAttachments: true });
    assert.strictEqual(types(docx), 'image1.png|image/png, image2.wmf|image/wmf, image3.emf|image/emf', `Metafiles: DOCX media typed (${types(docx)})`);

    // A PPTX holding an EMF and a WMF picture (written from a DOCX's pictures, read back).
    const pictures = await OfficeParser.parseOffice(metafilePictures(), { extractAttachments: true });
    assert.strictEqual(types(pictures), 'image1.emf|image/emf, image2.wmf|image/wmf', `Metafiles: DOCX pictures typed (${types(pictures)})`);
    const deck = await OfficeParser.parseOffice(Buffer.from(bytesOf((await pictures.to('pptx')).value)), { extractAttachments: true });
    const deckTypes = deck.attachments.map(a => a.mimeType).sort();
    assert.ok(deckTypes.includes('image/emf') && deckTypes.includes('image/wmf'), `Metafiles: PPTX media typed (${types(deck)})`);

    // By bytes: an EMF, a placeable WMF, a headerless WMF; and not a WMF: four bytes of its start and a
    // version that is none.
    assert.strictEqual(getMimeFromBytes(Buffer.from(emfBytes())), 'image/emf', 'Metafiles: EMF by its header');
    assert.strictEqual(getMimeFromBytes(Buffer.from(placeableWmf({ left: 0, top: 0, right: 100, bottom: 100 }, 1440))), 'image/wmf', 'Metafiles: placeable WMF by its key');
    assert.strictEqual(getMimeFromBytes(Buffer.from(headerlessWmf())), 'image/wmf', 'Metafiles: headerless WMF by its METAHEADER');
    assert.strictEqual(getMimeFromBytes(Buffer.from([1, 0, 9, 0])), undefined, 'Metafiles: four bytes of a METAHEADER are not a WMF');
    assert.strictEqual(getMimeFromBytes(Buffer.from([1, 0, 9, 0, 0, 7, 0, 0])), undefined, 'Metafiles: a METAHEADER of no version is not a WMF');
    assert.strictEqual(getMimeFromBytes(Buffer.from([1, 0, 0, 0, ...new Array(40).fill(0)])), undefined, 'Metafiles: a record of type 1 without " EMF" is not an EMF');
    assert.strictEqual(pictureKindOf(emfBytes()), 'EMF', 'Metafiles: named EMF');
    assert.strictEqual(pictureKindOf(headerlessWmf()), 'WMF', 'Metafiles: named WMF');
    // By extension, where the bytes say nothing.
    const named = await OfficeParser.parseOffice(docxOf(para(run('x')), { media: { 'image9.emf': new Uint8Array([0, 1, 2, 3]), 'image8.wmf': new Uint8Array([0, 1, 2, 3]) } }), { extractAttachments: true });
    assert.strictEqual(types(named), 'image9.emf|image/emf, image8.wmf|image/wmf', `Metafiles: typed by extension (${types(named)})`);

    // OCR is never given a metafile: no recognition, no OCR_FAILED; the PNG is still read.
    const recognized: number[] = [];
    ocrTestHooks.createWorker = async () => ({
        recognize: async (image: any) => { recognized.push(Buffer.from(image).length); return { data: { text: 'READ' } }; },
        reinitialize: async () => { },
        terminate: async () => { },
    });
    try {
        const { ast: ocred, codes } = await parseWithWarnings(c5Objects(), { ocr: true, extractAttachments: true });
        assert.ok(!codes.includes('OCR_FAILED'), `Metafiles: no OCR_FAILED (${codes})`);
        assert.deepStrictEqual(recognized, [TINY_PNG.length], `Metafiles: only the PNG is recognized (${recognized})`);
        assert.ok(ocred.attachments.find(a => a.name === 'image1.png')?.ocrText === 'READ' && !ocred.attachments.find(a => a.name === 'image3.emf')?.ocrText, 'Metafiles: the PNG has its OCR text, the EMF none');
    } finally {
        ocrTestHooks.createWorker = undefined;
        await terminateOcr();
    }
    console.log('  ✓ EMF and WMF typed by bytes and extension, never OCR\'d');
}

// ─── EMF and WMF written into DOCX and ODT ───────────────────────────────

/** DOCX and ODT output package EMF and WMF pictures as such, typed as Word and LibreOffice type them, sized by their headers. */
export async function testMetafileWriters(): Promise<void> {
    console.log('\n=== Metafiles written into DOCX and ODT ===');
    const emf = emfBytes(5080, 2540), wmf = placeableWmf({ left: 0, top: 0, right: 2880, bottom: 1440 }, 1440);
    assert.deepStrictEqual(metafileSize(emf), { w: 192, h: 96 }, 'Metafiles: an EMF sized by its frame');
    assert.deepStrictEqual(metafileSize(wmf), { w: 192, h: 96 }, 'Metafiles: a placeable WMF sized by its box and units per inch');
    assert.strictEqual(metafileSize(headerlessWmf({ x: 0, y: 0, w: 100, h: 100 })), null, 'Metafiles: a headerless WMF has no size of its own');
    assert.strictEqual(metafileSize(emfBytes(0, 100)), null, 'Metafiles: an EMF of no width has no size');
    assert.strictEqual(metafileSize(placeableWmf({ left: 0, top: 0, right: 100, bottom: 100 }, 0)), null, 'Metafiles: a WMF of no units per inch has no size');

    const { ast } = await parseWithWarnings(metafilePictures(emf, wmf), { extractAttachments: true });
    const warned: string[] = [];
    const docx = zipTexts((await ast.to('docx', { onWarning: (w: any) => warned.push(String(w.code)) } as any)).value);
    const docxBytes = unzipSync(bytesOf((await ast.to('docx')).value));
    assert.ok(!warned.includes('IMAGE_PROCESSING_FAILED'), `Metafiles: DOCX takes them without a warning (${warned})`);
    assert.ok(Buffer.from(docxBytes['word/media/image1.emf']).equals(Buffer.from(emf)), 'Metafiles: DOCX packages the EMF as it is');
    assert.ok(Buffer.from(docxBytes['word/media/image2.wmf']).equals(Buffer.from(wmf)), 'Metafiles: DOCX packages the WMF as it is');
    assert.ok(docx['[Content_Types].xml'].includes('<Default Extension="emf" ContentType="image/x-emf"/>') && docx['[Content_Types].xml'].includes('<Default Extension="wmf" ContentType="image/x-wmf"/>'), 'Metafiles: DOCX content types as Word writes them');
    assert.ok(/<wp:extent cx="1828800" cy="914400"\/>/.test(docx['word/document.xml']), 'Metafiles: DOCX sizes the picture by its header (2 by 1 inches)');
    const back = await OfficeParser.parseOffice(Buffer.from(bytesOf((await ast.to('docx')).value)), { extractAttachments: true });
    assert.deepStrictEqual(back.attachments.map(a => a.mimeType), ['image/emf', 'image/wmf'], 'Metafiles: DOCX output read back keeps both');

    const odtWarned: string[] = [];
    const odtValue = (await ast.to('odt', { onWarning: (w: any) => odtWarned.push(String(w.code)) } as any)).value;
    const odt = zipTexts(odtValue), odtBytes = unzipSync(bytesOf(odtValue));
    assert.ok(!odtWarned.includes('IMAGE_PROCESSING_FAILED'), `Metafiles: ODT takes them without a warning (${odtWarned})`);
    assert.ok(Buffer.from(odtBytes['Pictures/image1.emf']).equals(Buffer.from(emf)) && Buffer.from(odtBytes['Pictures/image2.wmf']).equals(Buffer.from(wmf)), 'Metafiles: ODT packages both as they are');
    assert.ok(odt['META-INF/manifest.xml'].includes('manifest:full-path="Pictures/image1.emf" manifest:media-type="image/x-emf"') && odt['META-INF/manifest.xml'].includes('manifest:full-path="Pictures/image2.wmf" manifest:media-type="image/x-wmf"'), 'Metafiles: ODT manifest types them as LibreOffice does');
    assert.ok(/svg:width="2in" svg:height="1in"|svg:width="144pt" svg:height="72pt"/.test(odt['content.xml']), `Metafiles: ODT sizes the picture by its header (${/svg:width="[^"]*" svg:height="[^"]*"/.exec(odt['content.xml'])?.[0]})`);
    const odtBack = await OfficeParser.parseOffice(Buffer.from(bytesOf(odtValue)), { extractAttachments: true });
    assert.deepStrictEqual(odtBack.attachments.map(a => a.mimeType).sort(), ['image/emf', 'image/wmf'], 'Metafiles: ODT output read back keeps both');
    console.log('  ✓ DOCX and ODT keep EMF and WMF, typed and sized');
}

// ─── RTF EMF and WMF pictures read, unread pictures reported ─────────────

/** Bytes as RTF picture hex. */
export const hexOf = (bytes: Uint8Array) => Buffer.from(bytes).toString('hex');
/** An RTF document of `body`. */
export const rtfOf = (body: string) => `{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Times New Roman;}}${body}}`;

/**
 * A shape `build(n)` read at n and at n/4: the larger within 8 times the smaller plus 250 ms (measured
 * again once when over), so the reading is linear in the shape's size.
 */
export const linear = async (label: string, build: (n: number) => Buffer | string, n: number, config: OfficeParserConfig = {}, check?: (ast: OfficeParserAST, n: number) => void): Promise<void> => {
    const quarter = Math.floor(n / 4);
    const large = build(n), small = build(quarter);
    const time = async (input: Buffer | string, count: number) => {
        const started = performance.now();
        const ast = await OfficeParser.parseOffice(typeof input === 'string' ? Buffer.from(input, 'latin1') : input, config);
        const ms = performance.now() - started;
        check?.(ast, count);
        return ms;
    };
    let smallMs = await time(small, quarter), largeMs = await time(large, n);
    if (largeMs > 8 * smallMs + 250) { smallMs = Math.min(smallMs, await time(small, quarter)); largeMs = Math.min(largeMs, await time(large, n)); }
    assert.ok(largeMs <= 8 * smallMs + 250, `${label}: linear (${Math.round(smallMs)} ms at ${quarter}, ${Math.round(largeMs)} ms at ${n})`);
    console.log(`  ✓ ${label}: ${Math.round(smallMs)} ms at ${quarter}, ${Math.round(largeMs)} ms at ${n}`);
};

/** The 16-bit words of a placeable WMF header. */
const placeableWords = (bytes: Uint8Array) => Array.from({ length: 11 }, (_, i) => bytes[i * 2] | bytes[i * 2 + 1] << 8);

/** RTF's EMF and WMF pictures are read (a WMF given its placeable header); a picture that is not read is reported. */
export async function testRtfPictures(): Promise<void> {
    console.log('\n=== RTF pictures: EMF, WMF and those not read ===');
    const pictures = (ast: OfficeParserAST) => allNodes(ast).filter(n => n.type === 'image');
    const read = (body: string, config: OfficeParserConfig = {}) => parseWithWarnings(rtfOf(body), { extractAttachments: true, ...config });
    const bytesOfAttachment = (ast: OfficeParserAST, i = 0) => Buffer.from(ast.attachments[i].data, 'base64');

    // An EMF, as it is.
    const emf = emfBytes(2010, 952);
    const emfDoc = await read(`\\pard LABELEMF\\par\\pard {\\pict\\emfblip\\picw120\\pich57\\picwgoal1140\\pichgoal540 ${hexOf(emf)}}\\par`);
    assert.strictEqual(pictures(emfDoc.ast).length, 1, 'RTF pictures: the EMF is a picture');
    assert.ok(emfDoc.ast.attachments.length === 1 && emfDoc.ast.attachments[0].mimeType === 'image/emf' && emfDoc.ast.attachments[0].name === 'image_1.emf' && bytesOfAttachment(emfDoc.ast).equals(Buffer.from(emf)), 'RTF pictures: the EMF attachment is the RTF\'s bytes, typed image/emf');
    assert.deepStrictEqual(emfDoc.codes, [], 'RTF pictures: an EMF warns nothing');

    // A WMF as RTF holds one (no placeable header): given one, its box the metafile's window, its units
    // per inch making the box \picwgoal wide.
    const window = { x: 10, y: 20, w: 2000, h: 1000 };
    const meta = headerlessWmf(window, 8);
    const wmfDoc = await read(`{\\pict\\wmetafile8\\picw2540\\pich1270\\picwgoal1440\\pichgoal720 ${hexOf(meta)}}`);
    const wmf = bytesOfAttachment(wmfDoc.ast);
    const words = placeableWords(wmf);
    assert.ok(wmfDoc.ast.attachments[0].mimeType === 'image/wmf' && wmfDoc.ast.attachments[0].name === 'image_1.wmf', 'RTF pictures: the WMF typed image/wmf');
    assert.deepStrictEqual(words.slice(0, 2), [0xCDD7, 0x9AC6], 'RTF pictures: the WMF has the placeable key');
    assert.strictEqual(words.slice(0, 10).reduce((a, w) => a ^ w, 0), words[10], 'RTF pictures: the placeable checksum is right');
    assert.deepStrictEqual(words.slice(3, 7), [10, 20, 2010, 1020], 'RTF pictures: the box is the metafile\'s window');
    assert.ok(Math.abs((words[5] - words[3]) * 1440 / words[7] - 1440) <= 1, `RTF pictures: the box is \\picwgoal wide at its units per inch (${words[7]})`);
    assert.ok(wmf.subarray(22).equals(Buffer.from(meta)), 'RTF pictures: the METAFILE follows the header unchanged');

    // The width from \picw where no goal is given, in the mapping mode's units; LibreOffice's \picw0 with goals.
    const inchOf = async (sizeWords: string, mapMode: number) => placeableWords(bytesOfAttachment((await read(`{\\pict\\wmetafile${mapMode}${sizeWords} ${hexOf(headerlessWmf(window, mapMode))}}`)).ast))[7];
    assert.strictEqual(await inchOf('\\picw2540\\pich1270', 8), 2000, 'RTF pictures: \\picw in HIMETRIC (mapping mode 8)');
    assert.strictEqual(await inchOf('\\picw96\\pich48', 1), 2000, 'RTF pictures: \\picw in pixels (mapping mode 1)');
    assert.strictEqual(await inchOf('\\picw1440\\pich720', 6), 2000, 'RTF pictures: \\picw in twips (mapping mode 6)');
    assert.strictEqual(await inchOf('\\picw0\\pich0\\picwgoal720\\pichgoal360', 8), 4000, 'RTF pictures: \\picw0 with goals, from the goals');
    // No window among its records: kept headerless, still a WMF. Already placeable: unchanged.
    const bare = headerlessWmf(undefined, 8);
    const bareDoc = await read(`{\\pict\\wmetafile8\\picwgoal1440\\pichgoal720 ${hexOf(bare)}}`);
    assert.ok(bytesOfAttachment(bareDoc.ast).equals(Buffer.from(bare)) && bareDoc.ast.attachments[0].mimeType === 'image/wmf', 'RTF pictures: a WMF without a window kept as it is, a WMF');
    const placed = placeableWmf({ left: 0, top: 0, right: 1440, bottom: 720 }, 1440);
    const placedDoc = await read(`{\\pict\\wmetafile8\\picwgoal1140\\pichgoal540 ${hexOf(placed)}}`);
    assert.ok(bytesOfAttachment(placedDoc.ast).equals(Buffer.from(placed)), 'RTF pictures: a placeable WMF given in RTF is unchanged');

    // LibreOffice's pair: the EMF in \shppict and a WMF of the same picture in \nonshppict; one picture, the EMF.
    const pair = await read(`\\pard {\\*\\shppict{\\pict\\emfblip ${hexOf(emf)}}}{\\nonshppict{\\pict\\wmetafile8 ${hexOf(meta)}}}\\par`);
    assert.ok(pictures(pair.ast).length === 1 && pair.ast.attachments.length === 1 && pair.ast.attachments[0].mimeType === 'image/emf', 'RTF pictures: \\shppict and \\nonshppict are one picture, the EMF');

    // Pictures not read: reported once each, no picture.
    for (const [body, part] of [
        ['{\\pict\\macpict\\picw10\\pich10 0011223344}', 'a QuickDraw picture, \\macpict'],
        ['{\\pict\\pmmetafile8 0011223344}', 'an OS/2 metafile, \\pmmetafile'],
        ['{\\pict\\picw10\\pich10 0011223344}', 'a picture of no format RTF names'],
        ['{\\pict\\pngblip\\bin4 PNGS}', 'a picture held as binary data, \\bin'],
    ] as const) {
        const { ast, warnings } = await read(`\\pard BEFORE\\par ${body}\\pard AFTER\\par`);
        assert.strictEqual(pictures(ast).length, 0, `RTF pictures: no picture for ${part}`);
        assert.deepStrictEqual(warnings.map(w => [w.code, w.details?.part]), [['CONTENT_PART_NOT_READ', part]], `RTF pictures: ${part} reported once (${JSON.stringify(warnings.map(w => w.message))})`);
        assert.ok(warnings[0].message.includes(part) && warnings[0].message.includes('not read'), `RTF pictures: the message names it (${warnings[0].message})`);
        assert.strictEqual((await ast.to('text')).value.replace(/\s+/g, ' ').trim(), 'BEFORE AFTER', `RTF pictures: the text around ${part} is read`);
        // Pictures not asked for are not reported.
        assert.deepStrictEqual((await read(`\\pard BEFORE\\par ${body}`, { extractAttachments: false })).codes, [], `RTF pictures: ${part} not reported without extractAttachments`);
    }
    // A list's bullet picture is not content, and not reported; an empty picture loses nothing.
    const bullet = await read('{\\*\\listtable{\\list\\listtemplateid1{\\listlevel\\levelnfc23{\\leveltext \\\'01*;}}{\\*\\listpicture{\\pict\\macpict 0011}}\\listid1}}\\pard x\\par{\\pict\\pngblip}');
    assert.deepStrictEqual(bullet.codes, [], `RTF pictures: a list picture and an empty picture are not reported (${bullet.codes})`);

    // Linear: many small EMF pictures.
    const small = hexOf(emfBytes());
    await linear('RTF pictures: 20,000 EMF pictures', n => rtfOf(`{\\pict\\emfblip ${small}}`.repeat(n)), 20000, { extractAttachments: true }, (ast, n) => assert.strictEqual(ast.attachments.length, n, 'RTF pictures: every EMF read'));
    console.log('  ✓ RTF EMF and WMF read, unread pictures reported');
}

// ─── RTF output writes each picture in its own format ───────────────────

/** A 1x1 GIF. */
const TINY_GIF = Buffer.from('R0lGODlhAQABAIAAAP///wAAACH5BAEAAAAALAAAAAABAAEAAAICRAEAOw==', 'base64');

/** RTF output writes PNG, JPEG, EMF and WMF pictures as such, and reports any other. */
export async function testRtfPictureWriters(): Promise<void> {
    console.log('\n=== RTF output writes pictures in their own formats ===');
    const window = { x: 0, y: 0, w: 2000, h: 1000 };
    const meta = headerlessWmf(window, 8);
    const emf = emfBytes(5080, 2540);
    const source = rtfOf(`\\pard A\\par{\\pict\\emfblip ${hexOf(emf)}}\\pard B\\par{\\pict\\wmetafile8\\picw2540\\pich1270\\picwgoal1440\\pichgoal720 ${hexOf(meta)}}\\pard C\\par`);
    const { ast } = await parseWithWarnings(source, { extractAttachments: true });

    // RTF to RTF and back: the EMF as it was, the WMF's metafile as it was, its size from its header.
    const warned: string[] = [];
    const rtf = String((await ast.to('rtf', { onWarning: (w: any) => warned.push(String(w.code)) } as any)).value);
    assert.ok(rtf.includes('\\pict\\emfblip') && rtf.includes('\\pict\\wmetafile8\\picw2540\\pich1270\\picwgoal1440\\pichgoal720') && !rtf.includes('jpegblip'), `RTF writer: \\emfblip and \\wmetafile8, sized (${/\\pict[^\n]*/g.exec(rtf)?.[0]})`);
    assert.deepStrictEqual(warned, [], 'RTF writer: EMF and WMF warn nothing');
    const again = await OfficeParser.parseOffice(Buffer.from(rtf, 'latin1'), { extractAttachments: true });
    assert.deepStrictEqual(again.attachments.map(a => a.mimeType), ['image/emf', 'image/wmf'], 'RTF writer: read back, both pictures');
    assert.ok(Buffer.from(again.attachments[0].data, 'base64').equals(Buffer.from(emf)), 'RTF writer: the EMF read back is identical');
    const wmfBack = Buffer.from(again.attachments[1].data, 'base64');
    assert.ok(wmfBack.subarray(22).equals(Buffer.from(meta)) && wmfBack.equals(Buffer.from(ast.attachments[1].data, 'base64')), 'RTF writer: the WMF read back is identical, its header too');

    // DOCX to RTF for an EMF and a placeable WMF: the WMF written without its header, sized by it.
    const placed = placeableWmf({ left: 0, top: 0, right: 2880, bottom: 1440 }, 1440);
    const fromDocx = await OfficeParser.parseOffice(metafilePictures(emf, placed), { extractAttachments: true });
    const docxRtf = String((await fromDocx.to('rtf')).value);
    assert.ok(docxRtf.includes('\\pict\\wmetafile8\\picw5080\\pich2540\\picwgoal2880\\pichgoal1440'), 'RTF writer: a placeable WMF\'s size from its header (2 by 1 inches)');
    const docxBack = await OfficeParser.parseOffice(Buffer.from(docxRtf, 'latin1'), { extractAttachments: true });
    assert.ok(Buffer.from(docxBack.attachments[0].data, 'base64').equals(Buffer.from(emf)), 'RTF writer: DOCX to RTF keeps the EMF');
    assert.ok(Buffer.from(docxBack.attachments[1].data, 'base64').subarray(22).equals(Buffer.from(placed).subarray(22)), 'RTF writer: DOCX to RTF keeps the WMF\'s metafile');

    // RTF to DOCX and ODT: packaged as metafiles, sized by their headers.
    const docx = zipTexts((await ast.to('docx')).value);
    assert.ok(Object.keys(docx).includes('word/media/image1.emf') && docx['[Content_Types].xml'].includes('<Default Extension="emf" ContentType="image/x-emf"/>'), 'RTF writer: RTF to DOCX packages the EMF');
    assert.ok(docx['word/document.xml'].includes('<wp:extent cx="1828800" cy="914400"/>'), 'RTF writer: RTF to DOCX sizes the EMF by its frame');
    const odt = zipTexts((await ast.to('odt')).value);
    assert.ok(Object.keys(odt).includes('Pictures/image1.emf') && Object.keys(odt).includes('Pictures/image2.wmf') && odt['META-INF/manifest.xml'].includes('image/x-wmf'), 'RTF writer: RTF to ODT packages both');

    // A GIF is written as a PNG of its pixels, never as JPEG data; a kind RTF does not hold (WebP) is reported and
    // written as its alt text.
    const oneAst = (name: string, mimeType: string, bytes: Uint8Array) => ({ type: 'docx', metadata: {}, content: [{ type: 'paragraph', text: '', children: [{ type: 'image', text: '', metadata: { attachmentName: name, altText: 'A picture' } }] }], attachments: [{ type: 'image', name, mimeType, extension: name.split('.').pop(), data: Buffer.from(bytes).toString('base64') }] } as any);
    const gifWarnings: any[] = [];
    const gifRtf = String((await OfficeGenerator.generate(oneAst('a.gif', 'image/gif', TINY_GIF), 'rtf', { onWarning: (w: any) => gifWarnings.push(w) } as any)).value);
    assert.ok(gifRtf.includes('{\\pict\\pngblip\\picw1\\pich1') && !gifRtf.includes('jpegblip') && !gifRtf.includes('A picture'), 'RTF writer: a GIF is written as a PNG');
    assert.deepStrictEqual(gifWarnings, [], 'RTF writer: a GIF warns nothing');
    const webp = new Uint8Array([0x52, 0x49, 0x46, 0x46, 26, 0, 0, 0, 0x57, 0x45, 0x42, 0x50, 0x56, 0x50, 0x38, 0x4C, ...new Array(18).fill(0)]);
    const webpWarnings: any[] = [];
    const webpRtf = String((await OfficeGenerator.generate(oneAst('a.webp', 'image/webp', webp), 'rtf', { onWarning: (w: any) => webpWarnings.push(w) } as any)).value);
    assert.ok(!webpRtf.includes('\\pict') && webpRtf.includes('A picture'), 'RTF writer: a WebP is written as its alt text');
    assert.ok(webpWarnings.some(w => w.code === 'IMAGE_PROCESSING_FAILED' && /RTF output does not take WebP pictures/.test(w.message)), `RTF writer: a WebP reported (${webpWarnings.map(w => w.message)})`);

    // Each picture at the size its own header gives, cut to the width available.
    const pictureAst = (bytes: Uint8Array, mimeType: string, cells = 0): OfficeParserAST => {
        const picture = { type: 'paragraph', text: '', children: [{ type: 'image', text: '', metadata: { attachmentName: 'p' } }] };
        const content = cells ? [{ type: 'table', children: [{ type: 'row', children: Array.from({ length: cells }, (_, i) => ({ type: 'cell', children: i ? [{ type: 'paragraph', text: 'X', children: [{ type: 'text', text: 'X' }] }] : [picture] })) }] }] : [picture];
        return { type: 'docx', metadata: {}, content, attachments: [{ type: 'image', name: 'p', mimeType, data: Buffer.from(bytes).toString('base64') }] } as any;
    };
    const sizeWords = async (bytes: Uint8Array, mimeType: string, cells = 0) => {
        const out = String((await OfficeGenerator.generate(pictureAst(bytes, mimeType, cells), 'rtf', { onWarning: () => { } } as any)).value);
        return { out, words: /\\pict\\[a-z0-9]+((?:\\[a-z]+-?\d+)*)/.exec(out)?.[1] ?? '' };
    };
    /** A PNG's signature and IHDR of the size given (the rest of the picture left out). */
    const pngOf = (w: number, h: number) => new Uint8Array([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, ...[w >>> 24, w >>> 16, w >>> 8, w].map(b => b & 0xFF), ...[h >>> 24, h >>> 16, h >>> 8, h].map(b => b & 0xFF), 8, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82]);
    assert.strictEqual((await sizeWords(emfBytes(2540, 1270), 'image/emf')).words, '\\picw2540\\pich1270\\picwgoal1440\\pichgoal720', 'RTF writer: an EMF of a 1 by 0.5 inch frame, sized by it');
    assert.strictEqual((await sizeWords(TINY_PNG, 'image/png')).words, '\\picw1\\pich1\\picwgoal15\\pichgoal15', 'RTF writer: a PNG of one pixel, at 96 an inch');
    assert.strictEqual((await sizeWords(pngOf(4000, 100), 'image/png')).words, '\\picw4000\\pich100\\picwgoal9000\\pichgoal225', 'RTF writer: a wide PNG cut to the page\'s width, its shape kept');
    assert.ok((await sizeWords(pngOf(4000, 100), 'image/png', 2)).words.includes('\\picwgoal4500\\pichgoal113'), 'RTF writer: in a cell of a two-cell table, cut to the cell\'s width');
    const jpeg = new Uint8Array([0xFF, 0xD8, 0xFF, 0xE0, 0, 16, 0x4A, 0x46, 0x49, 0x46, 0, 1, 1, 0, 0, 1, 0, 1, 0, 0, 0xFF, 0xC0, 0, 17, 8, 0, 48, 0, 192, 3, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1, 0xFF, 0xD9]);
    assert.strictEqual((await sizeWords(jpeg, 'image/jpeg')).words, '\\picw192\\pich48\\picwgoal2880\\pichgoal720', 'RTF writer: a JPEG sized by its frame header');
    assert.strictEqual((await sizeWords(pngOf(40, 20).subarray(0, 20), 'image/png')).words, '\\picwgoal4320\\pichgoal3240', 'RTF writer: a PNG whose header is cut short, the default goals');
    // Property: EMF frames anywhere in int32 and PNG sizes anywhere in uint31: every size word a whole number from 1 to
    // 2147483647, the goals within the width available, the output read back with one picture.
    const sizeRnd = lcg(0x512E);
    const int32 = () => (Math.floor(sizeRnd() * 4294967296) - 2147483648) | 0;
    const sizes = new Set<string>();
    for (let round = 0; round < 1200 && sizes.size < 1000; round++) {
        const emf = sizeRnd() < 0.5;
        const values = emf ? [int32(), int32(), int32(), int32()] : [Math.floor(sizeRnd() * 2147483648), Math.floor(sizeRnd() * 2147483648)];
        // (Small sizes too, where most real pictures are.)
        if (sizeRnd() < 0.3) for (let i = 0; i < values.length; i++) values[i] = Math.floor(sizeRnd() * 5000) - (emf ? 1000 : 0);
        const key = `${emf}:${values}`;
        if (sizes.has(key)) continue;
        sizes.add(key);
        let bytes: Uint8Array;
        if (emf) { bytes = emfBytes(); const view = new DataView(bytes.buffer); values.forEach((v, i) => view.setInt32(24 + i * 4, v, true)); }
        else bytes = pngOf(values[0], values[1]);
        const { out, words } = await sizeWords(bytes, emf ? 'image/emf' : 'image/png');
        const params = Array.from(words.matchAll(/\\(picw|pich|picwgoal|pichgoal)(-?\d+)/g), m => [m[1], Number(m[2])] as const);
        assert.ok(params.length >= 2 && params.every(([, v]) => Number.isInteger(v) && v >= 1 && v <= 2147483647), `RTF writer: size words in range for ${key} (${words})`);
        assert.ok(params.filter(([k]) => k === 'picwgoal').every(([, v]) => v <= 9000), `RTF writer: the goal within the page for ${key} (${words})`);
        const back = await OfficeParser.parseOffice(Buffer.from(out, 'latin1'), { extractAttachments: true });
        assert.strictEqual(back.attachments.length, 1, `RTF writer: read back with one picture for ${key}`);
    }
    assert.ok(sizes.size >= 1000, `RTF writer: ${sizes.size} distinct picture sizes`);
    console.log(`  ✓ ${sizes.size} distinct picture sizes written within range`);
    console.log('  ✓ RTF output writes PNG, JPEG, EMF and WMF as such, a GIF as PNG, reports the rest');
}

// ─── RTF tracked deletions ───────────────────────────────────────────────

/** A full-period LCG mod 2^32 (a = 1664525, c = 1013904223), seeded: the same cases each run. */
export const lcg = (seed: number) => {
    let state = seed >>> 0;
    return (): number => (state = (Math.imul(state, 1664525) + 1013904223) >>> 0) / 4294967296;
};

/** The text of a parse: its blocks' text, one a line. */
const blockTexts = (ast: OfficeParserAST) => ast.content.map(n => n.text ?? '').join('\n');

/** RTF text a tracked change deleted (`\deleted`) is not read, as DOCX's `w:del` is not; insertions are. */
export async function testRtfTrackedChanges(): Promise<void> {
    console.log('\n=== RTF tracked deletions ===');
    const read = async (body: string, config: OfficeParserConfig = {}) => (await parseWithWarnings(rtfOf(body), { extractAttachments: true, ...config })).ast;

    // The shapes Word and LibreOffice write.
    const hand = await read('\\pard KEEPONE {\\deleted\\revauthdel1 DELETEDWORD} {\\revised\\revauth1 INSERTEDWORD} KEEPTWO\\par');
    assert.strictEqual(blockTexts(hand), 'KEEPONE  INSERTEDWORD KEEPTWO', `RTF deletions: Word's shape (${blockTexts(hand)})`);
    const lo = await read('{\\*\\revtbl {Unknown;}{A;}}\\pard\\plain KEEPONE {\\deleted\\revauthdel1\\revdttmdel-2015295488 \\loch DELETEDWORD}{\\revised\\revauth1\\revdttm-2015295488 \\loch  INSERTEDWORD} KEEPTWO\\par');
    assert.strictEqual(blockTexts(lo), 'KEEPONE  INSERTEDWORD KEEPTWO', `RTF deletions: LibreOffice's shape (${blockTexts(lo)})`);

    // Where a deletion ends: \deleted0, \plain, the end of its group; a nested group inherits it.
    const cases: [string, string][] = [
        ['A{\\deleted B\\deleted0 C}D', 'ACD'],
        ['A{\\deleted B\\plain C}D', 'ACD'],
        ['A{\\deleted B{\\b C}D}E', 'AE'],
        ['A{\\deleted B}{C}D', 'ACD'],
        ['A\\deleted B\\par\\pard C', 'A'],
        ['A{\\deleted \\tab\\line\\u8364?\\emdash X}B', 'AB'],
    ];
    for (const [body, expected] of cases) {
        const ast = await read(`\\pard ${body}\\par`);
        assert.strictEqual(blockTexts(ast), expected, `RTF deletions: ${body} reads ${JSON.stringify(expected)} (${JSON.stringify(blockTexts(ast))})`);
    }

    // What a deletion holds is skipped whole: a picture, a footnote, a link's text, a cell's text (the cell kept).
    const png = TINY_PNG.toString('hex');
    const picture = await read(`\\pard A{\\deleted {\\pict\\pngblip ${png}}{\\*\\shppict{\\pict\\pngblip ${png}}}}B\\par`);
    assert.ok(picture.attachments.length === 0 && !allNodes(picture).some(n => n.type === 'image') && blockTexts(picture) === 'AB', 'RTF deletions: a deleted picture is neither node nor attachment');
    const note = await read('\\pard A{\\deleted X{\\footnote\\pard NOTETEXT}}B\\par');
    assert.ok(!allNodes(note).some(n => n.type === 'note') && blockTexts(note) === 'AB', 'RTF deletions: a deleted footnote is no note');
    const link = await read('\\pard A{\\field{\\*\\fldinst HYPERLINK "https://example.com/"}{\\fldrslt {\\deleted LINKTEXT}}}B\\par');
    assert.ok(!allNodes(link).some(n => (n.metadata as any)?.link) && blockTexts(link) === 'AB', `RTF deletions: a deleted link's text is not read (${blockTexts(link)})`);
    const table = await read('\\trowd\\cellx1000\\cellx2000\\pard\\intbl {\\deleted GONE}\\cell\\pard\\intbl KEPT\\cell\\row\\pard AFTER\\par');
    const cells = allNodes(table).filter(n => n.type === 'cell');
    assert.deepStrictEqual(cells.map(c => c.text), ['', 'KEPT'], `RTF deletions: a deleted cell's text, the cell kept (${JSON.stringify(cells.map(c => c.text))})`);
    // An annotation in a deletion is no comment.
    const comment = await read('\\pard A{\\deleted {\\*\\atnid X}{\\*\\atnauthor X}\\chatn{\\*\\annotation\\pard COMMENTTEXT}}B\\par');
    assert.ok(!allNodes(comment).some(n => n.type === 'comment'), 'RTF deletions: a deleted annotation is no comment');
    // (The same shapes, not deleted, are read: the note, the comment and the picture.)
    const kept = await read(`\\pard A{X{\\footnote\\pard NOTETEXT}}{{\\*\\atnid X}\\chatn{\\*\\annotation\\pard COMMENTTEXT}}{\\pict\\pngblip ${png}}B\\par`);
    assert.ok(['note', 'comment', 'image'].every(type => allNodes(kept).some(n => n.type === type)), 'RTF deletions: not deleted, the note, comment and picture are read');

    // Property: random runs of text, groups, \deleted, \deleted0, \plain and \b read as a reference walker reads them.
    const rnd = lcg(0x5EED);
    const seen = new Set<string>();
    for (let round = 0; round < 2400 && seen.size < 2000; round++) {
        let rtf = '', expected = '';
        const stack: boolean[] = [];
        let del = false, depth = 0;
        const length = 1 + Math.floor(rnd() * 14);
        for (let i = 0; i < length; i++) {
            const pick = Math.floor(rnd() * 7);
            if (pick === 0 || pick === 1) {
                const word = 'abcdefgh'.slice(0, 1 + Math.floor(rnd() * 3)).split('').map(() => 'abcdefgh'[Math.floor(rnd() * 8)]).join('');
                rtf += word; if (!del) expected += word;
            } else if (pick === 2) { rtf += '{'; stack.push(del); depth++; }
            else if (pick === 3 && depth > 0) { rtf += '}'; del = stack.pop()!; depth--; }
            else if (pick === 4) { rtf += '\\deleted '; del = true; }
            else if (pick === 5) { rtf += rnd() < 0.5 ? '\\deleted0 ' : '\\plain '; del = false; }
            else { rtf += '\\b '; }
        }
        rtf += '}'.repeat(depth);
        if (seen.has(rtf)) continue;
        seen.add(rtf);
        const ast = await read(`\\pard ${rtf}\\par`);
        assert.strictEqual(blockTexts(ast), expected, `RTF deletions: ${rtf} reads ${JSON.stringify(expected)} (${JSON.stringify(blockTexts(ast))})`);
    }
    assert.ok(seen.size >= 2000, `RTF deletions: ${seen.size} distinct documents`);
    console.log(`  ✓ ${seen.size} distinct documents read as the reference reads them`);

    await linear('RTF deletions: 200,000 deleted groups', n => rtfOf('\\pard ' + '{\\deleted x}y'.repeat(n) + '\\par'), 200000, {}, (ast, n) => assert.strictEqual(ast.content[0]?.text?.length, n, 'RTF deletions: every kept character read'));
    console.log('  ✓ RTF tracked deletions left out, insertions read');
}

// ─── RTF equations ───────────────────────────────────────────────────────

/** A random equation, written both as OMML and as RTF's control words for it. */
type MathTree = { omml: string; rtf: string };

/** Text as RTF writes it: `\`, `{` and `}` escaped, characters past ASCII as `\uN?`. */
const rtfText = (text: string) => Array.from(text).map(ch => {
    const code = ch.charCodeAt(0);
    if (ch === '\\' || ch === '{' || ch === '}') return '\\' + ch;
    return code > 126 ? `\\u${code > 32767 ? code - 65536 : code}?` : ch;
}).join('');
const xmlText = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** A builder of random OMML trees and their RTF, from `rnd`. */
const mathTrees = (rnd: () => number) => {
    const pick = <T>(list: readonly T[]) => list[Math.floor(rnd() * list.length)];
    const chars = ['a', 'x', '1', '2', ' ', '<', '&', '"', '>', '\\', '{', '}', '%', '$', '_', '^', '~', '#', String.fromCharCode(0x3B1), String.fromCharCode(0x2211), String.fromCharCode(0x221E)];
    const textOf = () => Array.from({ length: 1 + Math.floor(rnd() * 3) }, () => pick(chars)).join('');
    const el = (name: string, inner: MathTree): MathTree => ({ omml: `<m:${name}>${inner.omml}</m:${name}>`, rtf: `{\\m${name} ${inner.rtf}}` });
    const val = (name: string, value: string): MathTree => ({ omml: `<m:${name} m:val="${xmlText(value)}"/>`, rtf: `{\\m${name} ${rtfText(value)}}` });
    const cat = (...parts: MathTree[]): MathTree => ({ omml: parts.map(p => p.omml).join(''), rtf: parts.map(p => p.rtf).join('') });
    const node = (depth: number): MathTree => {
        const kind = depth > 3 ? 'r' : pick(['r', 'r', 'f', 'sSup', 'sSub', 'sSubSup', 'd', 'rad', 'nary', 'func', 'acc', 'bar', 'm'] as const);
        const arg = () => seq(depth + 1);
        switch (kind) {
            case 'r': { const t = textOf(); return { omml: `<m:r><m:t xml:space="preserve">${xmlText(t)}</m:t></m:r>`, rtf: `{\\mr ${rtfText(t)}}` }; }
            case 'f': return el('f', cat(rnd() < 0.5 ? el('fPr', val('type', 'lin')) : cat(), el('num', arg()), el('den', arg())));
            case 'sSup': return el('sSup', cat(el('e', arg()), el('sup', arg())));
            case 'sSub': return el('sSub', cat(el('e', arg()), el('sub', arg())));
            case 'sSubSup': return el('sSubSup', cat(el('e', arg()), el('sub', arg()), el('sup', arg())));
            case 'd': return el('d', cat(el('dPr', cat(...(rnd() < 0.6 ? [val('begChr', pick(['[', '{', '|', '<', String.fromCharCode(0x27E8)]))] : []), ...(rnd() < 0.6 ? [val('endChr', pick([']', '}', '|', '>', String.fromCharCode(0x27E9)]))] : []))), el('e', arg()), ...(rnd() < 0.3 ? [el('e', arg())] : [])));
            case 'rad': return el('rad', rnd() < 0.5 ? cat(el('deg', arg()), el('e', arg())) : cat(el('radPr', val('degHide', '1')), el('deg', cat()), el('e', arg())));
            case 'nary': return el('nary', cat(...(rnd() < 0.7 ? [el('naryPr', val('chr', pick([String.fromCharCode(0x2211), String.fromCharCode(0x220F), String.fromCharCode(0x222B), '*'])))] : []), el('sub', arg()), el('sup', arg()), el('e', arg())));
            case 'func': return el('func', cat(el('fName', arg()), el('e', arg())));
            case 'acc': return el('acc', cat(...(rnd() < 0.5 ? [el('accPr', val('chr', pick([String.fromCharCode(0x304), String.fromCharCode(0x20D7), String.fromCharCode(0x307)])))] : []), el('e', arg())));
            case 'bar': return el('bar', cat(...(rnd() < 0.5 ? [el('barPr', val('pos', pick(['top', 'bot'])))] : []), el('e', arg())));
            case 'm': { const rows = 1 + Math.floor(rnd() * 2), cols = 1 + Math.floor(rnd() * 2); return el('m', cat(...Array.from({ length: rows }, () => el('mr', cat(...Array.from({ length: cols }, () => el('e', arg()))))))); }
        }
    };
    const seq = (depth: number): MathTree => cat(...Array.from({ length: 1 + Math.floor(rnd() * 2) }, () => node(depth)));
    return seq;
};

/** RTF equations (`\mmath`) are read as math, as DOCX's are; one that cannot be read is reported. */
export async function testRtfEquations(): Promise<void> {
    console.log('\n=== RTF equations ===');
    const read = (body: string, config: OfficeParserConfig = {}) => parseWithWarnings(rtfOf(body), { extractAttachments: true, ...config });
    const maths = (ast: OfficeParserAST) => allNodes(ast).filter(n => n.type === 'code' && (n.metadata as any)?.math);
    const png = hexOf(TINY_PNG);

    // Parity: random equations read from RTF as their OMML reads.
    const rnd = lcg(0xE0A710);
    const tree = mathTrees(rnd);
    const seen = new Set<string>();
    for (let round = 0; round < 3600 && seen.size < 3000; round++) {
        const t = tree(0);
        const block = rnd() < 0.5;
        if (seen.has(t.rtf + block)) continue;
        seen.add(t.rtf + block);
        const omml = block ? `<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath>${t.omml}</m:oMath></m:oMathPara>` : `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${t.omml}</m:oMath>`;
        const expected = ommlToLatex(parseXmlString(omml).documentElement as any);
        const rtf = block ? `{\\mmath{\\*\\moMathPara{\\*\\moMath ${t.rtf}}}}` : `{\\mmath{\\*\\moMath ${t.rtf}}}`;
        const { ast } = await read(`\\pard ${rtf}\\par`);
        const found = maths(ast);
        if (!expected.trim()) { assert.strictEqual(found.length, 0, `RTF equations: an empty one is no node (${rtf})`); continue; }
        assert.strictEqual(found.length, 1, `RTF equations: one math node for ${rtf}`);
        assert.strictEqual(found[0].text, expected, `RTF equations: ${rtf} reads as its OMML does`);
        assert.strictEqual((found[0].metadata as any).math, block ? 'block' : 'inline', `RTF equations: ${block ? 'a display' : 'an inline'} equation`);
    }
    assert.ok(seen.size >= 3000, `RTF equations: ${seen.size} distinct equations`);
    console.log(`  ✓ ${seen.size} distinct equations read as their OMML reads`);

    // LibreOffice's shape: the equation in its paragraph, its own picture named, no picture node.
    const loMath = '{\\mmath {\\*\\moMath {\\mr x}{\\mr =}{\\mf {\\mnum {\\mr a}}{\\mden {\\mr b}}}}{\\mmathPict {\\*\\shppict{\\pict{\\*\\picprop{\\sp{\\sn wzDescription}{\\sv }}}\\picw896\\pich998\\picwgoal508\\pichgoal566\\pngblip ' + png + '}}{\\nonshppict{\\pict\\wmetafile8\\picwgoal508\\pichgoal566 ' + hexOf(headerlessWmf({ x: 0, y: 0, w: 100, h: 100 })) + '}}}}';
    const lo = await read(`\\pard BEFOREMATH ${loMath} AFTERMATH\\par`);
    const loNodes = maths(lo.ast);
    assert.ok(loNodes.length === 1 && loNodes[0].text === 'x=\\frac{a}{b}', `RTF equations: LibreOffice's equation (${loNodes.map(n => n.text)})`);
    assert.strictEqual(lo.ast.content[0].text, 'BEFOREMATH x=\\frac{a}{b} AFTERMATH', 'RTF equations: the LaTeX in the paragraph\'s text');
    assert.ok(lo.ast.content[0].children!.includes(loNodes[0]), 'RTF equations: the node is a child of its paragraph');
    const fallback = (loNodes[0].metadata as any).fallbackAttachmentName;
    assert.ok(fallback && lo.ast.attachments.find(a => a.name === fallback)?.mimeType === 'image/png', `RTF equations: fallbackAttachmentName names the PNG (${fallback})`);
    assert.ok(!allNodes(lo.ast).some(n => n.type === 'image') && lo.ast.attachments.length === 1, 'RTF equations: no picture node, one attachment');
    assert.deepStrictEqual(lo.codes, [], 'RTF equations: nothing reported');
    // Without attachments: the equation alone.
    const bare = await read(`\\pard ${loMath}\\par`, { extractAttachments: false });
    assert.ok(maths(bare.ast).length === 1 && !(maths(bare.ast)[0].metadata as any).fallbackAttachmentName && bare.ast.attachments.length === 0, 'RTF equations: without extractAttachments, no fallbackAttachmentName');

    // LibreOffice's empty equation (one it could not convert) with its picture: the picture, as before.
    const empty = await read(`\\pard {\\mmath {\\*\\moMath }{\\mmathPict {\\*\\shppict{\\pict\\pngblip ${png}}}{\\nonshppict{\\pict\\wmetafile8 00}}}}\\par`);
    assert.ok(maths(empty.ast).length === 0 && allNodes(empty.ast).filter(n => n.type === 'image').length === 1 && empty.codes.length === 0, `RTF equations: an empty one with a picture is the picture (${empty.codes})`);
    // An equation of no text and no picture loses what it holds: reported. One that holds nothing is not.
    const lost = await read('\\pard {\\mmath{\\*\\moMath{\\mctrlPr\\i\\f3 }}}\\par');
    assert.deepStrictEqual(lost.warnings.map(w => [w.code, w.details?.part]), [['CONTENT_PART_NOT_READ', 'an equation, \\moMath']], `RTF equations: no text, no picture reported (${lost.codes})`);
    assert.deepStrictEqual((await read('\\pard {\\mmath{\\*\\moMath }}\\par')).codes, [], 'RTF equations: an empty equation is not reported');
    const lostPictured = await read(`\\pard {\\mmath{\\*\\moMath{\\mctrlPr\\i }}{\\mmathPict{\\pict\\pngblip ${png}}}}\\par`, { extractAttachments: false });
    assert.deepStrictEqual(lostPictured.codes, ['CONTENT_PART_NOT_READ'], 'RTF equations: no text, its picture not extracted, reported');

    // An \mmath with no \moMath of its own is its own \moMath: its text is math, not dropped.
    for (const [body, latex] of [['A{\\mmath x+y}B', 'x+y'], ['A{\\mmath{\\mr x}}B', 'x']] as const) {
        const own = await read(`\\pard ${body}\\par`);
        assert.ok(maths(own.ast).length === 1 && maths(own.ast)[0].text === latex && (maths(own.ast)[0].metadata as any).math === 'inline', `RTF equations: ${body} is the equation ${latex} (${maths(own.ast).map(n => n.text)})`);
        assert.ok(own.ast.content[0].text === `A${latex}B` && own.codes.length === 0, `RTF equations: ${body} reads A${latex}B, nothing reported (${own.ast.content[0].text}; ${own.codes})`);
    }
    const ownPictured = await read(`\\pard A{\\mmath{\\mr x}{\\mmathPict{\\*\\shppict{\\pict\\pngblip ${png}}}}}B\\par`);
    assert.ok(maths(ownPictured.ast).length === 1 && maths(ownPictured.ast)[0].text === 'x' && (maths(ownPictured.ast)[0].metadata as any).fallbackAttachmentName === 'image_1.png', 'RTF equations: an \\mmath of its own text and a picture names the picture');
    assert.ok(ownPictured.ast.attachments.length === 1 && !allNodes(ownPictured.ast).some(n => n.type === 'image') && !allNodes(ownPictured.ast).some(n => (n.text ?? '').includes(png.slice(0, 16))), 'RTF equations: its picture one attachment, no picture node, its data in no text');
    const onlyPicture = await read(`\\pard A{\\mmath{\\mmathPict{\\*\\shppict{\\pict\\pngblip ${png}}}}}B\\par`);
    assert.ok(maths(onlyPicture.ast).length === 0 && allNodes(onlyPicture.ast).filter(n => n.type === 'image').length === 1, 'RTF equations: an \\mmath of a picture alone is the picture');
    for (const body of ['{\\mmath}', '{\\mmath }']) {
        const nothing = await read(`\\pard A${body}B\\par`);
        assert.ok(maths(nothing.ast).length === 0 && nothing.ast.content[0].text === 'AB' && nothing.codes.length === 0, `RTF equations: ${body} is nothing, not reported`);
    }
    const ownLost = await read('\\pard {\\mmath{\\mctrlPr\\i }}\\par', { extractAttachments: false });
    assert.deepStrictEqual(ownLost.warnings.map(w => [w.code, w.details?.part]), [['CONTENT_PART_NOT_READ', 'an equation, \\moMath']], 'RTF equations: an \\mmath of no text reported as the \\moMath form is');
    assert.strictEqual(maths((await read('\\pard A{\\deleted {\\mmath x+y}}B\\par')).ast).length, 0, 'RTF equations: a deleted \\mmath of its own text is not read');
    // Property: random equations read the same with and without their \moMath.
    const ownRnd = lcg(0x0A11A7);
    const ownTree = mathTrees(ownRnd);
    const ownSeen = new Set<string>();
    for (let round = 0; round < 2000 && ownSeen.size < 1500; round++) {
        const t = ownTree(0);
        if (ownSeen.has(t.rtf)) continue;
        ownSeen.add(t.rtf);
        const wrapped = maths((await read(`\\pard {\\mmath{\\*\\moMath ${t.rtf}}}\\par`)).ast).map(n => n.text);
        const bareOwn = maths((await read(`\\pard {\\mmath ${t.rtf}}\\par`)).ast).map(n => n.text);
        assert.deepStrictEqual(bareOwn, wrapped, `RTF equations: {\\mmath ${t.rtf}} reads as its \\moMath form`);
    }
    assert.ok(ownSeen.size >= 1500, `RTF equations: ${ownSeen.size} distinct equations without \\moMath`);
    console.log(`  ✓ ${ownSeen.size} distinct equations read the same without their \\moMath`);
    // Such an \mmath reads only its math as math: a picture, a note, a field or an ignorable group in it is read as it
    // is read anywhere, after the equation (a picture's data was the equation's text, a note's text too).
    const textsOf = (ast: OfficeParserAST) => allNodes(ast).map(n => n.text ?? '');
    const bareImage = await read(`\\pard A{\\mmath{\\pict\\pngblip ${png}}}B\\par`);
    assert.ok(allNodes(bareImage.ast).filter(n => n.type === 'image').length === 1 && bareImage.ast.attachments.length === 1 && maths(bareImage.ast).length === 0
        && !textsOf(bareImage.ast).some(t => t.includes(png.slice(0, 16))) && bareImage.codes.length === 0, `RTF equations: an \\mmath of a bare picture is the picture, its data in no text (${textsOf(bareImage.ast)}; ${bareImage.codes})`);
    const bareUnextracted = await read(`\\pard A{\\mmath{\\pict\\pngblip ${png}}}B\\par`, { extractAttachments: false });
    assert.ok(!allNodes(bareUnextracted.ast).some(n => n.type === 'image') && !textsOf(bareUnextracted.ast).some(t => t.includes(png.slice(0, 16))), 'RTF equations: not extracted, the bare picture is in no node and no text');
    const olderPicture = await read('\\pard A{\\mmath x{\\nonshppict{\\pict\\wmetafile8 0100090000}}}B\\par');
    assert.ok(maths(olderPicture.ast).length === 1 && maths(olderPicture.ast)[0].text === 'x' && !allNodes(olderPicture.ast).some(n => n.type === 'image')
        && !textsOf(olderPicture.ast).some(t => t.includes('0100090000')) && olderPicture.codes.length === 0, `RTF equations: \\nonshppict in an \\mmath is no math (${maths(olderPicture.ast).map(n => n.text)}; ${olderPicture.codes})`);
    const shapePicture = await read(`\\pard A{\\mmath{\\mr x}{\\*\\shppict{\\pict\\pngblip ${png}}}}B\\par`);
    assert.ok(maths(shapePicture.ast).length === 1 && maths(shapePicture.ast)[0].text === 'x' && !(maths(shapePicture.ast)[0].metadata as any).fallbackAttachmentName
        && allNodes(shapePicture.ast).filter(n => n.type === 'image').length === 1 && shapePicture.codes.length === 0, `RTF equations: a \\shppict in an \\mmath is a picture beside the equation, not its fallback (${shapePicture.codes})`);
    const noteInMath = await read('\\pard A{\\mmath x{\\footnote\\pard FN\\par}}B\\par');
    const mathNotes = allNodes(noteInMath.ast).filter(n => n.type === 'note');
    assert.ok(maths(noteInMath.ast).length === 1 && maths(noteInMath.ast)[0].text === 'x' && mathNotes.length === 1 && mathNotes[0].text === 'FN'
        && noteInMath.ast.content[0].text === 'AxB' && noteInMath.codes.length === 0, `RTF equations: a footnote in an \\mmath is a note (${maths(noteInMath.ast).map(n => n.text)}; ${mathNotes.map(n => n.text)}; ${noteInMath.ast.content[0].text})`);
    const linkInMath = await read('\\pard A{\\mmath{\\field{\\*\\fldinst HYPERLINK "http://e.com"}{\\fldrslt link}}}B\\par');
    const links = allNodes(linkInMath.ast).filter(n => n.type === 'text' && n.text === 'link');
    assert.ok(maths(linkInMath.ast).length === 0 && links.length === 1 && (links[0].metadata as any)?.link === 'http://e.com' && linkInMath.codes.length === 0, `RTF equations: a field in an \\mmath is its link text (${JSON.stringify(links.map(n => n.metadata))}; ${linkInMath.codes})`);
    const macpict = await read('\\pard A{\\mmath x{\\macpict 0102}}B\\par');
    assert.ok(maths(macpict.ast).length === 1 && maths(macpict.ast)[0].text === 'x', `RTF equations: \\macpict is no math word (${maths(macpict.ast).map(n => n.text)})`);
    // Property: an \mmath of random parts reads as {\mmath{\*\moMath <its math parts>}<its \mmathPict>}<its other parts>.
    const partRnd = lcg(0x11AA7);
    const partPick = (n: number) => Math.floor(partRnd() * n);
    const mathWord = () => 'abcxyz'[partPick(6)] + 'pqrs'[partPick(4)];
    const partKinds: (() => { rtf: string; math: boolean })[] = [
        () => ({ rtf: mathWord(), math: true }),
        () => ({ rtf: `{\\mr ${mathWord()}}`, math: true }),
        () => ({ rtf: `{\\i ${mathWord()}}`, math: true }),
        () => ({ rtf: `{\\pict\\pngblip ${png}}`, math: false }),
        () => ({ rtf: `{\\*\\shppict{\\pict\\pngblip ${png}}}`, math: false }),
        () => ({ rtf: '{\\nonshppict{\\pict\\wmetafile8 0100090000}}', math: false }),
        () => ({ rtf: `{\\footnote\\pard N${mathWord()}\\par}`, math: false }),
        () => ({ rtf: `{\\*\\unknowndest J${mathWord()}}`, math: false }),
        () => ({ rtf: `{\\field{\\*\\fldinst HYPERLINK "http://e.com/${mathWord()}"}{\\fldrslt L${mathWord()}}}`, math: false }),
    ];
    const shapeOf = (ast: OfficeParserAST) => {
        const out: string[] = [];
        const walk = (n: OfficeContentNode) => {
            const m = n.metadata as any;
            out.push(`${n.type}:${n.text ?? ''}${m?.math ? `/${m.math}${m.fallbackAttachmentName ? '+picture' : ''}` : ''}${m?.link ? `@${m.link}` : ''}`);
            for (const c of n.children ?? []) walk(c);
            for (const note of n.notes ?? []) { out.push('note['); walk(note); out.push(']'); }
        };
        for (const n of ast.content) walk(n);
        return `${out.join('|')} attachments=${ast.attachments.length}`;
    };
    const partSeen = new Set<string>();
    for (let round = 0; round < 6000 && partSeen.size < 1500; round++) {
        const parts = Array.from({ length: 1 + partPick(6) }, () => partKinds[partPick(partKinds.length)]());
        const own = partPick(3) === 0 ? `{\\mmathPict{\\*\\shppict{\\pict\\pngblip ${png}}}}` : '';
        const body = parts.map(p => p.rtf).join('') + own;
        if (partSeen.has(body)) continue;
        partSeen.add(body);
        const got = (await read(`\\pard A{\\mmath ${body}}B\\par`)).ast;
        const expected = (await read(`\\pard A{\\mmath{\\*\\moMath ${parts.filter(p => p.math).map(p => p.rtf).join('')}}${own}}${parts.filter(p => !p.math).map(p => p.rtf).join('')}B\\par`)).ast;
        assert.strictEqual(shapeOf(got), shapeOf(expected), `RTF equations: {\\mmath ${body.split(png).join('PNG')}} reads as its math and its other parts`);
        assert.ok(!textsOf(got).some(t => /89504e47|0100090000|J[a-z][p-s]/.test(t)), `RTF equations: {\\mmath ${body.split(png).join('PNG')}} holds no picture data or ignorable text in any text`);
    }
    assert.ok(partSeen.size >= 1500, `RTF equations: ${partSeen.size} distinct \\mmath groups of mixed parts`);
    console.log(`  ✓ ${partSeen.size} distinct \\mmath groups without \\moMath read as their math and their other parts`);
    await linear('RTF equations: an \\mmath of 4,000 footnotes', n => rtfOf(`\\pard A{\\mmath x${'{\\footnote\\pard N\\par}'.repeat(n)}}B\\par`), 4000, { extractAttachments: true }, (ast, n) => assert.strictEqual(allNodes(ast).filter(x => x.type === 'note').length, n, 'RTF equations: every note read'));
    await linear('RTF equations: an \\mmath of 4,000 bare pictures', n => rtfOf(`\\pard A{\\mmath x${`{\\pict\\pngblip ${png}}`.repeat(n)}}B\\par`), 4000, { extractAttachments: true }, (ast, n) => assert.strictEqual(ast.attachments.length, n, 'RTF equations: every picture read'));

    // Word's shape: \moMathPara, a value as a parameter, the picture an ignorable group after \mmath.
    const word = await read(`\\pard {\\mmath{\\*\\moMathPara{\\*\\moMath{\\md{\\mdPr{\\mbegChr [}{\\mendChr ]}}{\\me{\\mr a}}}{\\mrad{\\mradPr\\mdegHide1 }{\\mdeg}{\\me{\\mr x}}}}}}{\\*\\mmathPict{\\pict\\pngblip ${png}}}\\par`);
    const wordNodes = maths(word.ast);
    assert.ok(wordNodes.length === 1 && wordNodes[0].text === '[a]\\sqrt{x}' && (wordNodes[0].metadata as any).math === 'block', `RTF equations: Word's shape (${wordNodes.map(n => n.text)})`);
    assert.ok((wordNodes[0].metadata as any).fallbackAttachmentName === 'image_1.png' && !allNodes(word.ast).some(n => n.type === 'image'), 'RTF equations: Word\'s picture after the equation is its fallback, no picture node');
    // In a cell and in a footnote.
    const placed = await read('\\trowd\\cellx2000\\pard\\intbl {\\mmath{\\*\\moMath{\\mr y}}}\\cell\\row\\pard N{\\footnote\\pard {\\mmath{\\*\\moMath{\\mr z}}}}\\par');
    const cell = allNodes(placed.ast).find(n => n.type === 'cell')!, note = allNodes(placed.ast).find(n => n.type === 'note')!;
    assert.ok(maths({ content: [cell] } as any).some(n => n.text === 'y') && maths({ content: [note] } as any).some(n => n.text === 'z'), 'RTF equations: in a cell and in a footnote');
    // The header's math properties are no equation.
    const header = await read('{\\*\\mmathPr\\mmathFont34\\mbrkBin0\\mbrkBinSub0\\msmallFrac0\\mdispDef1\\mlMargin0\\mrMargin0\\mdefJc1\\mwrapIndent1440\\mintLim0\\mnaryLim1}{\\mmathPr\\mmathFont34}\\pard TEXT\\par');
    assert.ok(maths(header.ast).length === 0 && header.ast.content.map(n => n.text).join('') === 'TEXT', `RTF equations: \\mmathPr yields nothing (${header.ast.content.map(n => n.text)})`);
    // A tracked deletion holding an equation: nothing.
    const deletedMath = await read(`\\pard A{\\deleted {\\mmath{\\*\\moMath{\\mr q}}}{\\*\\mmathPict{\\pict\\pngblip ${png}}}}B\\par`);
    assert.ok(maths(deletedMath.ast).length === 0 && deletedMath.ast.attachments.length === 0 && deletedMath.ast.content[0].text === 'AB', 'RTF equations: a deleted equation is not read');

    // An RTF chunk of a DOCX: its equation's picture renamed with the chunk's attachments.
    const chunk = rtfOf(`\\pard ${loMath}\\par`);
    const docx = docxOf(para(run('DOC')) + '<w:altChunk r:id="rIdC"/>' + para(pictureRun('rIp')), {
        rels: [['rIdC', 'aFChunk', 'chunk.rtf'], ['rIp', 'image', 'media/image_1.png']],
        media: { 'image_1.png': TINY_PNG },
        extra: { 'word/chunk.rtf': chunk },
    });
    const chunked = await OfficeParser.parseOffice(docx, { extractAttachments: true });
    const chunkMath = maths(chunked)[0];
    const renamed = (chunkMath?.metadata as any)?.fallbackAttachmentName;
    assert.ok(renamed && renamed !== 'image_1.png' && chunked.attachments.some(a => a.name === renamed && a.mimeType === 'image/png'), `RTF equations: the chunk's picture renamed and followed (${renamed}; ${chunked.attachments.map(a => a.name)})`);
    // To DOCX: the equation drawn as the chunk's picture, beside the document's own.
    const chunkedDocx = zipTexts((await chunked.to('docx')).value);
    assert.strictEqual(Object.keys(chunkedDocx).filter(k => k.startsWith('word/media/')).length, 2, 'RTF equations: the chunk\'s equation is its picture in DOCX output, beside the document\'s');
    assert.ok(chunkedDocx['word/document.xml'].includes('descr="x=\\frac{a}{b}"'), 'RTF equations: its LaTeX the picture\'s alt text');

    // Linear: many equations; one nested to the tokenizer's limit reads, its depth cut short.
    await linear('RTF equations: 20,000 equations', n => rtfOf('\\pard ' + '{\\mmath{\\*\\moMath{\\mf{\\mnum{\\mr a}}{\\mden{\\mr b}}}}}'.repeat(n) + '\\par'), 20000, {}, (ast, n) => assert.strictEqual(maths(ast).length, n, 'RTF equations: every equation read'));
    const deep = 240;
    const nested = await read(`\\pard {\\mmath{\\*\\moMath ${'{\\me '.repeat(deep)}{\\mr x}${'}'.repeat(deep)}}}\\par`);
    assert.ok(maths(nested.ast).length === 1 && maths(nested.ast)[0].text!.includes('\\ldots'), `RTF equations: nested ${deep} deep, read with its depth cut (${maths(nested.ast)[0]?.text?.slice(0, 40)})`);
    console.log('  ✓ RTF equations read as math, their pictures named, the unreadable reported');
}

// ─── RTF equations written as their own pictures where math cannot be typeset ───

/** LibreOffice's RTF of a display equation `x=\frac{a}{b}` (image_1.png) and an inline `yz^2` (image_2.png). */
const EQUATIONS_LO = path.join(__dirname, '../files/objects/equations-libreoffice.rtf');
const LO_LATEX = ['x=\\frac{a}{b}', 'yz^2'];

/** The group starting at `start` in RTF `text` (its braces balanced, escaped braces skipped). */
const rtfGroupAt = (text: string, start: number): string => {
    let depth = 0;
    for (let i = start; i < text.length; i++) {
        const ch = text[i];
        if (ch === '\\') { i++; continue; }
        if (ch === '{') depth++;
        else if (ch === '}' && --depth === 0) return text.slice(start, i + 1);
    }
    throw new Error('unbalanced group');
};

/** Every value of attribute `attr` on elements named `tag` in `xml`, as an XML reader reads it. */
const attributeValues = (xml: string, tag: string, attr: string): string[] =>
    Array.from((parseXmlString(xml).getElementsByTagName(tag) as any) as ArrayLike<any>).map(e => e.getAttribute(attr) as string).filter(v => v !== null && v !== '');
/** The text of every element named `tag` in `xml`, as an XML reader reads it. */
const elementTexts = (xml: string, tag: string): string[] =>
    Array.from((parseXmlString(xml).getElementsByTagName(tag) as any) as ArrayLike<any>).map(e => e.textContent as string);

/**
 * An equation that carries its document's own picture (CodeMetadata.fallbackAttachmentName) is written as that
 * picture, its LaTeX the alt text, by DOCX, ODT, RTF, PPTX, ODP and PDF output; HTML, Markdown, LaTeX, text and
 * chunks keep the LaTeX. onNode is asked about the equation, never the picture.
 */
export async function testRtfEquationWriters(): Promise<void> {
    console.log('\n=== RTF equations written as their own pictures ===');
    const source = fs.readFileSync(EQUATIONS_LO);
    const ast = await OfficeParser.parseOffice(source, { extractAttachments: true });
    const mathNodes = allNodes(ast).filter(n => n.type === 'code' && (n.metadata as any)?.math);
    assert.deepStrictEqual(mathNodes.map(n => [n.text, (n.metadata as any).fallbackAttachmentName]), [[LO_LATEX[0], 'image_1.png'], [LO_LATEX[1], 'image_2.png']], 'Equation pictures: the fixture\'s two equations, each naming its picture');
    const write = async (doc: OfficeParserAST, format: string, config: any = {}) => {
        const warnings: any[] = [];
        const value = (await OfficeGenerator.generate(doc, format as any, { ...config, ...(format === 'pdf' ? { pdfConfig: { engine: 'native' } } : {}), onWarning: (w: any) => warnings.push(w) } as any)).value as any;
        return { value, warnings, codes: warnings.map(w => String(w.code)) };
    };
    const media = (files: Record<string, string>, prefix: string) => Object.keys(files).filter(k => k.startsWith(prefix)).length;
    const pdfBack = async (value: unknown) => {
        const back = await OfficeParser.parseOffice(Buffer.from(bytesOf(value)), { extractAttachments: true });
        return { images: back.attachments.filter(a => a.type === 'image').length, text: allNodes(back).filter(n => !n.children?.length).map(n => n.text ?? '').join('\n') };
    };
    // The document's text in each package, alt text (descr, svg:title, svg:desc) left out.
    const docxText = (files: Record<string, string>) => elementTexts(files['word/document.xml'], 'w:t').join('');
    const odfText = (files: Record<string, string>) => files['content.xml'].replace(/<svg:title>[^<]*<\/svg:title>|<svg:desc>[^<]*<\/svg:desc>/g, '').replace(/<[^>]+>/g, '');
    const pptxText = (files: Record<string, string>) => Object.keys(files).filter(k => /^ppt\/(slides|notesSlides)\/[^/]+\.xml$/.test(k)).map(k => elementTexts(files[k], 'a:t').join('')).join('\n');
    const noMathWarning = (codes: string[], format: string) => assert.ok(!codes.includes('CONTENT_NOT_REPRESENTABLE'), `Equation pictures: ${format} reports no math (${codes})`);

    // 1. Every writer that cannot typeset math writes the two pictures, their LaTeX the alt text.
    const docx = await write(ast, 'docx');
    const docxFiles = zipTexts(docx.value);
    assert.strictEqual(media(docxFiles, 'word/media/'), 2, 'Equation pictures: DOCX packages both pictures');
    assert.deepStrictEqual(attributeValues(docxFiles['word/document.xml'], 'wp:docPr', 'descr'), LO_LATEX, 'Equation pictures: DOCX alt text is each LaTeX');
    assert.ok(!docxText(docxFiles).includes('frac') && !docxText(docxFiles).includes('yz^2'), `Equation pictures: DOCX writes no LaTeX as text (${docxText(docxFiles)})`);
    const inlineParagraph = docxFiles['word/document.xml'].split('</w:p>').find(p => p.includes('INLINE'))!;
    assert.ok(inlineParagraph.includes('<w:drawing>') && inlineParagraph.includes(' AFTERMATH') && inlineParagraph.indexOf('<w:drawing>') > inlineParagraph.indexOf('INLINE '), 'Equation pictures: DOCX writes the inline equation in its line');
    noMathWarning(docx.codes, 'DOCX');

    const odt = await write(ast, 'odt');
    const odtFiles = zipTexts(odt.value);
    assert.strictEqual(media(odtFiles, 'Pictures/'), 2, 'Equation pictures: ODT packages both pictures');
    assert.deepStrictEqual(elementTexts(odtFiles['content.xml'], 'svg:title'), LO_LATEX, 'Equation pictures: ODT alt text is each LaTeX');
    assert.ok(!odfText(odtFiles).includes('frac') && !odfText(odtFiles).includes('yz^2'), 'Equation pictures: ODT writes no LaTeX as text');
    const odtInline = odtFiles['content.xml'].split('</text:p>').find(p => p.includes('INLINE'))!;
    assert.ok(/INLINE .*<draw:frame [^>]*text:anchor-type="as-char".*AFTERMATH/.test(odtInline), `Equation pictures: ODT writes the inline equation as a character in its line (${odtInline})`);
    noMathWarning(odt.codes, 'ODT');

    const rtf = await write(ast, 'rtf');
    const rtfOut = String(rtf.value);
    assert.strictEqual(rtfOut.split('{\\pict\\pngblip').length - 1, 2, 'Equation pictures: RTF writes both pictures');
    const between = rtfOut.slice(rtfOut.indexOf('INLINE '), rtfOut.indexOf(' AFTERMATH'));
    assert.ok(between.includes('{\\pict\\pngblip') && !between.includes('\\par'), 'Equation pictures: RTF writes the inline equation in its paragraph');
    assert.ok(!rtfOut.includes('frac') && !rtfOut.includes('yz^2') && !rtfOut.includes('{\\f2 '), 'Equation pictures: RTF writes no LaTeX run');
    assert.ok(rtfOut.includes('\\pngblip\\picw53\\pich59\\picwgoal795\\pichgoal885') && rtfOut.includes('\\pngblip\\picw38\\pich31\\picwgoal570\\pichgoal465'), 'Equation pictures: RTF writes each picture at its own size');
    noMathWarning(rtf.codes, 'RTF');

    const pptx = await write(ast, 'pptx');
    const pptxFiles = zipTexts(pptx.value);
    assert.strictEqual(media(pptxFiles, 'ppt/media/'), 2, 'Equation pictures: PPTX packages both pictures');
    assert.deepStrictEqual(Object.keys(pptxFiles).filter(k => /^ppt\/slides\/[^/]+\.xml$/.test(k)).flatMap(k => attributeValues(pptxFiles[k], 'p:cNvPr', 'descr')).sort(), [...LO_LATEX].sort(), 'Equation pictures: PPTX alt text is each LaTeX');
    assert.ok(!pptxText(pptxFiles).includes('frac') && !pptxText(pptxFiles).includes('yz^2') && pptxText(pptxFiles).includes('AFTERMATH'), 'Equation pictures: PPTX writes no LaTeX run, the text kept');
    const odp = await write(ast, 'odp');
    const odpFiles = zipTexts(odp.value);
    assert.strictEqual(media(odpFiles, 'Pictures/'), 2, 'Equation pictures: ODP packages both pictures');
    assert.deepStrictEqual(elementTexts(odpFiles['content.xml'], 'svg:desc').sort(), [...LO_LATEX].sort(), 'Equation pictures: ODP alt text is each LaTeX');
    assert.ok(!odfText(odpFiles).includes('frac') && !odfText(odpFiles).includes('yz^2') && odfText(odpFiles).includes('AFTERMATH'), 'Equation pictures: ODP writes no LaTeX run, the text kept');

    const pdf = await pdfBack((await write(ast, 'pdf')).value);
    assert.ok(pdf.images === 2 && !pdf.text.includes('frac') && pdf.text.includes('AFTERMATH'), `Equation pictures: native PDF draws both pictures, no LaTeX (${pdf.images}; ${JSON.stringify(pdf.text)})`);

    // The HTML the default PDF engine prints draws the pictures; HTML output keeps the LaTeX.
    const printed = String((await new HtmlGenerator(ast, { htmlConfig: { standalone: true } } as any, { drawsMathPictures: true }).generate()).value);
    const alts = Array.from(printed.matchAll(/<img\b[^>]*\balt="([^"]*)"/g), m => m[1]);
    assert.deepStrictEqual(alts, LO_LATEX, `Equation pictures: the printed HTML draws both pictures, their LaTeX the alt text (${alts})`);
    assert.ok(!printed.includes('class="math') && !/empty-paragraph[^>]*>\s*<br>/.test(printed.slice(printed.indexOf('BEFOREMATH'), printed.indexOf('INLINE'))), 'Equation pictures: the printed HTML writes no math, and the display equation\'s paragraph is not empty');
    const html = String((await write(ast, 'html')).value);
    assert.ok(html.includes('math math-inline') && html.includes('yz^2') && !html.includes('<img'), 'Equation pictures: HTML output keeps the LaTeX');
    const epub = zipTexts((await write(ast, 'epub')).value);
    assert.ok(Object.keys(epub).some(k => /\.xhtml$/.test(k) && epub[k].includes('math math-inline') && epub[k].includes('yz^2')), 'Equation pictures: EPUB keeps the LaTeX');
    assert.ok(String((await write(ast, 'md')).value).includes('$yz^2$'), 'Equation pictures: Markdown keeps the LaTeX');
    assert.ok(String((await write(ast, 'tex')).value).includes('\\frac{a}{b}'), 'Equation pictures: LaTeX output keeps the LaTeX');
    assert.ok(String((await write(ast, 'text')).value).includes('x=\\frac{a}{b}'), 'Equation pictures: text keeps the LaTeX');
    assert.ok(((await write(ast, 'chunks')).value as { text: string }[]).some(c => c.text.includes('yz^2')), 'Equation pictures: chunks keep the LaTeX');
    for (const format of ['xlsx', 'ods']) {
        const files = zipTexts((await write(ast, format)).value);
        assert.ok(Object.values(files).some(x => x.includes('x=\\frac{a}{b}')) && !Object.keys(files).some(k => /^(xl\/media|Pictures)\//.test(k)), `Equation pictures: ${format} writes the LaTeX in its cell, no picture`);
    }

    // 2. Image modes: no pictures written, the LaTeX as before; 'image+ocr-text' as the default. Not extracted: the LaTeX.
    const bare = await OfficeParser.parseOffice(source, { extractAttachments: false });
    for (const [label, doc, config] of [['includeImages: none', ast, { includeImages: 'none' }], ['includeImages: ocr-text-only', ast, { includeImages: 'ocr-text-only' }], ['extractAttachments: false', bare, {}]] as const) {
        const d = await write(doc, 'docx', config), o = await write(doc, 'odt', config), r = await write(doc, 'rtf', config);
        const df = zipTexts(d.value), of = zipTexts(o.value);
        assert.ok(media(df, 'word/media/') === 0 && docxText(df).includes('x=\\frac{a}{b}') && df['word/document.xml'].includes('CodeChar'), `Equation pictures: ${label}, DOCX writes the LaTeX`);
        assert.ok(media(of, 'Pictures/') === 0 && odfText(of).includes('x=\\frac{a}{b}'), `Equation pictures: ${label}, ODT writes the LaTeX`);
        assert.ok(!String(r.value).includes('\\pict') && String(r.value).includes('{\\f2 yz^2}') && r.warnings.some(w => w.code === 'CONTENT_NOT_REPRESENTABLE' && w.details?.feature === 'math'), `Equation pictures: ${label}, RTF writes the LaTeX, reported`);
        const p = zipTexts((await write(doc, 'pptx', config)).value), q = zipTexts((await write(doc, 'odp', config)).value);
        assert.ok(media(p, 'ppt/media/') === 0 && pptxText(p).includes('yz^2') && media(q, 'Pictures/') === 0 && odfText(q).includes('yz^2'), `Equation pictures: ${label}, PPTX and ODP write the LaTeX`);
        const n = await pdfBack((await write(doc, 'pdf', config)).value);
        assert.ok(n.images === 0 && n.text.includes('yz^2'), `Equation pictures: ${label}, native PDF draws the LaTeX`);
    }
    const both = await write(ast, 'docx', { includeImages: 'image+ocr-text' });
    assert.ok(media(zipTexts(both.value), 'word/media/') === 2 && !docxText(zipTexts(both.value)).includes('frac'), 'Equation pictures: image+ocr-text writes the pictures, no text after them');

    // 3. onNode is asked about the equation: a string for it is written in its place, false leaves it out; never asked about a picture.
    for (const format of ['docx', 'odt', 'rtf', 'pptx', 'odp', 'pdf']) {
        const seen = new Set<string>();
        const replaced = await write(ast, format, { onNode: (n: OfficeContentNode) => { seen.add(n.type); return n.type === 'code' ? 'EQUATION' : undefined; } });
        const dropped = await write(ast, format, { onNode: (n: OfficeContentNode) => n.type === 'code' ? false : undefined });
        let count: number, pictures: number, latex: boolean, droppedPictures: number;
        if (format === 'rtf') {
            const v = String(replaced.value), x = String(dropped.value);
            count = v.split('EQUATION').length - 1; pictures = v.split('\\pict').length - 1; droppedPictures = x.split('\\pict').length - 1; latex = x.includes('frac') || x.includes('yz^2');
        } else if (format === 'pdf') {
            const v = await pdfBack(replaced.value), x = await pdfBack(dropped.value);
            count = v.text.split('EQUATION').length - 1; pictures = v.images; droppedPictures = x.images; latex = x.text.includes('frac') || x.text.includes('yz^2');
        } else {
            const v = zipTexts(replaced.value), x = zipTexts(dropped.value);
            const prefix = format === 'docx' ? 'word/media/' : format === 'pptx' ? 'ppt/media/' : 'Pictures/';
            // (A package writer writes onNode's string as it is, as its markup.)
            count = Object.values(v).join('').split('EQUATION').length - 1; pictures = media(v, prefix); droppedPictures = media(x, prefix);
            latex = Object.values(x).some(t => t.includes('frac') || t.includes('yz^2'));
        }
        assert.ok(count === 2 && pictures === 0, `Equation pictures: ${format}, onNode's string for each equation written in its place, no picture (${count}, ${pictures})`);
        assert.deepStrictEqual([...seen].sort(), ['code', 'paragraph', 'text'], `Equation pictures: ${format}, onNode never asked about a picture`);
        assert.ok(droppedPictures === 0 && !latex, `Equation pictures: ${format}, onNode's false leaves out the picture and the LaTeX`);
    }

    // 4. Where the equation stands: in every container, each equation written once, as its picture or as its LaTeX.
    const fixtureText = source.toString('latin1');
    const group = rtfGroupAt(fixtureText, fixtureText.indexOf('{\\mmath '));
    const styles = '{\\stylesheet{\\s0 Normal;}{\\s1\\outlinelevel0 heading 1;}}';
    const listTable = `{\\*\\listtable{\\list\\listtemplateid1{\\listlevel\\levelnfc23{\\leveltext\\'01\\u8226 ?;}{\\levelnumbers;}}\\listid1}}{\\*\\listoverridetable{\\listoverride\\listid1\\ls1}}`;
    const inRow = (cell: string) => `\\trowd\\cellx3000\\cellx6000 ${cell} \\pard\\intbl C\\cell\\row\\pard\\par`;
    const wordBlock = `{\\mmath{\\*\\moMathPara{\\*\\moMath{\\mf{\\mnum{\\mr a}}{\\mden{\\mr b}}}}}}{\\*\\mmathPict{\\*\\shppict{\\pict\\pngblip ${hexOf(TINY_PNG)}}}}`;
    // [body, the place's equations]
    const places: Record<string, [string, number]> = {
        cell: [inRow(`\\pard\\intbl A${group}B\\cell`), 1],
        cellAlone: [inRow(`\\pard\\intbl ${group}\\cell`), 1],
        cellSecondParagraph: [inRow(`\\pard\\intbl C1\\par ${group}\\cell`), 1],
        cellTwo: [inRow(`\\pard\\intbl ${group}${group}\\cell`), 2],
        cellWordBlock: [inRow(`\\pard\\intbl ${wordBlock}\\cell`), 1],
        cellHeading: [styles + inRow(`\\pard\\intbl\\s1\\outlinelevel0 ${group}\\cell`), 1],
        cellList: [listTable + inRow(`\\pard\\intbl\\ls1\\ilvl0{\\listtext\\u8226 ?\\tab}${group}\\cell`), 1],
        cellNote: [inRow(`\\pard\\intbl C{\\super\\chftn}{\\footnote\\pard\\plain{\\super\\chftn} ${group}\\par}\\cell`), 1],
        nestedTable: [inRow(`\\pard\\intbl\\itap1 X\\nestcell{\\*\\nesttableprops\\trowd\\cellx3000\\nestrow}{\\pard\\intbl\\itap2 ${group}\\nestcell{\\*\\nesttableprops\\trowd\\cellx3000\\nestrow}}\\cell`), 1],
        footnote: [`\\pard A{\\super\\chftn}{\\footnote\\pard\\plain N ${group}\\par}B\\par`, 1],
        list: [`${listTable}\\pard\\ls1 L${group}M\\par`, 1],
        heading: [`${styles}\\pard\\s1\\outlinelevel0 H${group}\\par`, 1],
        textBox: [`\\pard before{\\shp{\\*\\shpinst\\shpleft0\\shptop0\\shpright3000\\shpbottom1000{\\sp{\\sn shapeType}{\\sv 202}}{\\shptxt\\pard ${group}\\par}}}after\\par`, 1],
    };
    // Where each writer writes the LaTeX instead of the picture: notes are text on slides, and native PDF's nested
    // tables' cells are flat text. Everywhere else, the picture (in native PDF a list item's, heading's or note's in its line).
    const asLatex: Record<string, string[]> = {
        pptx: ['footnote', 'cellNote'],
        odp: ['footnote', 'cellNote'],
        pdf: ['nestedTable'],
    };
    const occurrences = (text: string, word: string) => text.split(word).length - 1;
    for (const [place, [body, n]] of Object.entries(places)) {
        const doc = await OfficeParser.parseOffice(Buffer.from(rtfOf(body), 'latin1'), { extractAttachments: true });
        assert.strictEqual(allNodes(doc).filter(x => x.type === 'code' && (x.metadata as any)?.fallbackAttachmentName).length, n, `Equation pictures: in a ${place}, ${n} equation(s) read with their pictures`);
        for (const format of ['docx', 'odt', 'rtf', 'pptx', 'odp', 'pdf', 'html']) {
            const drawn = !asLatex[format]?.includes(place);
            let pictures: number, latexCount: number;
            if (format === 'html') {
                // The HTML the default PDF engine prints: its pictures with their LaTeX alt; no LaTeX outside alt and id (a heading's slug) attributes.
                const printed = String((await new HtmlGenerator(doc, { htmlConfig: { standalone: true } } as any, { drawsMathPictures: true }).generate()).value);
                const printedBody = printed.slice(printed.indexOf('<body'));
                pictures = Array.from(printedBody.matchAll(/<img\b[^>]*\balt="([^"]*)"/g)).filter(m => m[1].includes('frac')).length;
                latexCount = occurrences(printedBody.replace(/\s(alt|id)="[^"]*"/g, ''), 'frac');
            } else {
                const out = await write(doc, format);
                if (format === 'rtf') { pictures = occurrences(String(out.value), '\\pict'); latexCount = occurrences(String(out.value), 'frac'); }
                else if (format === 'pdf') { const back = await pdfBack(out.value); pictures = back.images; latexCount = occurrences(back.text, 'frac'); }
                else {
                    const files = zipTexts(out.value);
                    pictures = media(files, format === 'docx' ? 'word/media/' : format === 'pptx' ? 'ppt/media/' : 'Pictures/');
                    // (Text only: a heading's bookmark is named for its text, the LaTeX included.)
                    latexCount = Object.entries(files).filter(([k]) => /\.xml$/.test(k)).reduce((sum, [, t]) => sum + occurrences(format === 'docx' || format === 'pptx' ? elementTexts(t, format === 'docx' ? 'w:t' : 'a:t').join('') : t.replace(/<svg:title>[^<]*<\/svg:title>|<svg:desc>[^<]*<\/svg:desc>/g, '').replace(/<[^>]+>/g, ''), 'frac'), 0);
                }
            }
            assert.ok(pictures === (drawn ? n : 0) && latexCount === (drawn ? 0 : n), `Equation pictures: in a ${place}, ${format} writes ${drawn ? 'the picture' : 'the LaTeX'} of each of ${n}, once (${pictures} pictures, LaTeX ${latexCount} times)`);
        }
    }
    console.log(`  ✓ Equation pictures: ${Object.keys(places).length} places, each equation written once, as its picture or as its LaTeX`);
    // (3, in a cell: native PDF draws an equation alone in a cell as its picture whatever onNode says of it, and writes
    // no LaTeX there either way. What it draws for onNode's verdict is not settled here.)
    const aloneDoc = await OfficeParser.parseOffice(Buffer.from(rtfOf(places.cellAlone[0]), 'latin1'), { extractAttachments: true });
    for (const verdict of [false, 'EQUATIONX'] as const) {
        const back = await pdfBack((await write(aloneDoc, 'pdf', { onNode: (x: OfficeContentNode) => x.type === 'code' ? verdict : undefined })).value);
        assert.ok(!back.text.includes('frac'), `Equation pictures: native PDF, onNode ${JSON.stringify(verdict)} for an equation alone in a cell, no LaTeX (${JSON.stringify(back.text)})`);
    }

    // What a cell draws apart in native PDF is not written again from the text its parents hold: a chart (its
    // paragraph's and cell's text its data) and a picture whose recognized text its paragraph and cell hold.
    const chartDoc = await OfficeParser.parseOffice(fs.readFileSync(path.join(__dirname, '../files/charts/inline-chart.odt')), { extractAttachments: true });
    const chart = allNodes(chartDoc).find(x => x.type === 'chart')!;
    assert.ok(chart?.text?.includes('Revenue and margin'), 'Equation pictures: the chart fixture\'s title in its text');
    const cellOf = (child: OfficeContentNode, text: string): OfficeContentNode => ({ type: 'cell', text, metadata: { row: 0, col: 0 }, children: [{ type: 'paragraph', text, children: [child] }] } as any);
    const tableAround = (cell: OfficeContentNode): OfficeContentNode[] => [
        { type: 'table', text: cell.text, children: [{ type: 'row', text: cell.text, children: [cell, { type: 'cell', text: 'Z', metadata: { row: 0, col: 1 }, children: [{ type: 'paragraph', text: 'Z', children: [{ type: 'text', text: 'Z' }] }] }] }] } as any,
        { type: 'paragraph', text: 'after', children: [{ type: 'text', text: 'after' }] },
    ];
    const chartTable = { ...chartDoc, content: tableAround(cellOf(chart, chart.text!)) } as OfficeParserAST;
    const chartBack = await pdfBack((await write(chartTable, 'pdf')).value);
    assert.strictEqual(occurrences(chartBack.text, 'Revenue and margin'), 1, `Equation pictures: native PDF writes a cell's chart once, not again from its cell's text (${JSON.stringify(chartBack.text)})`);
    const ocrImage = { type: 'image', text: 'OCRWORDS', metadata: { attachmentName: 'p.png', altText: 'ALTWORDS' } } as any as OfficeContentNode;
    const pictureTable = { type: 'odt', metadata: {}, attachments: [{ type: 'image', name: 'p.png', mimeType: 'image/png', extension: 'png', data: TINY_PNG.toString('base64') }], content: tableAround(cellOf(ocrImage, 'OCRWORDS')) } as any as OfficeParserAST;
    for (const [mode, times] of [['image-only', 0], ['image+ocr-text', 1]] as const) {
        const back = await pdfBack((await write(pictureTable, 'pdf', { includeImages: mode })).value);
        assert.ok(back.images === 1 && occurrences(back.text, 'OCRWORDS') === times, `Equation pictures: native PDF, ${mode}, a cell's picture drawn once, its recognized text ${times} time(s) (${back.images}; ${JSON.stringify(back.text)})`);
    }
    console.log('  ✓ Native PDF writes what a cell draws apart once, not again from the text its parents hold');

    // PPTX and ODP: a nested table's cell is a line of its table's cell, and its equations, drawn as their pictures after
    // the table, are not in that line; without pictures, the line holds their LaTeX.
    const second = rtfGroupAt(fixtureText, fixtureText.indexOf('{\\mmath ', fixtureText.indexOf(group) + group.length));
    const nestedRow = (inner: string) => `\\pard\\intbl\\itap1 X\\nestcell{\\*\\nesttableprops\\trowd\\cellx3000\\nestrow}{\\pard\\intbl\\itap2 ${inner}\\nestcell{\\*\\nesttableprops\\trowd\\cellx3000\\nestrow}}\\cell`;
    const nestedDoc = await OfficeParser.parseOffice(Buffer.from(rtfOf(inRow(nestedRow(`N1 ${second} N2\\par N3 ${group}`))), 'latin1'), { extractAttachments: true });
    const slideTexts = (files: Record<string, string>) => Object.keys(files).filter(k => /^ppt\/slides\/[^/]+\.xml$/.test(k)).flatMap(k => elementTexts(files[k], 'a:t'));
    const nestedPptx = zipTexts((await write(nestedDoc, 'pptx')).value), nestedOdp = zipTexts((await write(nestedDoc, 'odp')).value);
    assert.ok(media(nestedPptx, 'ppt/media/') === 2 && slideTexts(nestedPptx).includes('X') && slideTexts(nestedPptx).includes('N1  N2\nN3') && !slideTexts(nestedPptx).join('').includes('frac') && !slideTexts(nestedPptx).join('').includes('yz^2'),
        `Equation pictures: PPTX draws a nested cell's two equations, its line without their LaTeX (${JSON.stringify(slideTexts(nestedPptx))})`);
    assert.ok(media(nestedOdp, 'Pictures/') === 2 && odfText(nestedOdp).includes('N1') && odfText(nestedOdp).includes('N3') && !odfText(nestedOdp).includes('frac') && !odfText(nestedOdp).includes('yz^2'),
        `Equation pictures: ODP draws a nested cell's two equations, its line without their LaTeX (${odfText(nestedOdp)})`);
    const noPictures = zipTexts((await write(nestedDoc, 'pptx', { includeImages: 'none' })).value);
    const noPicturesOdp = zipTexts((await write(nestedDoc, 'odp', { includeImages: 'none' })).value);
    assert.ok(media(noPictures, 'ppt/media/') === 0 && occurrences(slideTexts(noPictures).join('\n'), 'x=\\frac{a}{b}') === 1 && occurrences(slideTexts(noPictures).join('\n'), 'yz^2') === 1,
        `Equation pictures: PPTX without pictures, a nested cell's line holds each LaTeX once (${JSON.stringify(slideTexts(noPictures))})`);
    assert.ok(media(noPicturesOdp, 'Pictures/') === 0 && occurrences(odfText(noPicturesOdp), 'x=\\frac{a}{b}') === 1 && occurrences(odfText(noPicturesOdp), 'yz^2') === 1,
        `Equation pictures: ODP without pictures, a nested cell's line holds each LaTeX once (${odfText(noPicturesOdp)})`);
    // The equation under k wrappers in its nested cell: a picture or its LaTeX in the line, once, at every depth (the
    // line leaves out exactly what the pictures' walk draws).
    const depths: string[] = [];
    for (let k = 0; k <= 14; k++) {
        const doc = await OfficeParser.parseOffice(Buffer.from(rtfOf(places.nestedTable[0]), 'latin1'), { extractAttachments: true });
        const holder = allNodes(doc).find(x => x.type === 'paragraph' && x.children?.some(c => c.type === 'code'))!;
        let wrapped: OfficeContentNode = holder.children!.find(c => c.type === 'code')!;
        for (let i = 0; i < k; i++) wrapped = { type: 'paragraph', text: wrapped.text, children: [wrapped] };
        holder.children = [wrapped];
        const files = zipTexts((await write(doc, 'pptx')).value);
        const pictures = media(files, 'ppt/media/'), latex = occurrences(slideTexts(files).join('\n'), 'frac');
        depths.push(`${k}:${pictures}+${latex}`);
        assert.strictEqual(pictures + latex, 1, `Equation pictures: PPTX, a nested cell's equation under ${k} wrappers written once (${pictures} pictures, LaTeX ${latex} times)`);
    }
    console.log(`  ✓ PPTX and ODP leave a nested cell's equations drawn as pictures out of its line (depths ${depths.join(' ')})`);
    // Linear: a nested table of n rows, each an equation with its own picture, through PPTX.
    const nestedRows = (n: number) => OfficeParser.parseOffice(Buffer.from(rtfOf(`\\trowd\\cellx6000\\pard\\intbl\\itap1 X\\nestcell{\\*\\nesttableprops\\trowd\\cellx3000\\nestrow}${Array.from({ length: n }, (_, i) => `{\\pard\\intbl\\itap2 R${i} ${second}\\nestcell{\\*\\nesttableprops\\trowd\\cellx3000\\nestrow}}`).join('')}\\cell\\row\\pard after\\par`), 'latin1'), { extractAttachments: true });
    const fewRows = await nestedRows(150), manyRows = await nestedRows(600);
    const timeDeck = async (doc: OfficeParserAST) => { const started = performance.now(); await write(doc, 'pptx'); return performance.now() - started; };
    let fewMs = await timeDeck(fewRows), manyMs = await timeDeck(manyRows);
    if (manyMs > 8 * fewMs + 250) { fewMs = Math.min(fewMs, await timeDeck(fewRows)); manyMs = Math.min(manyMs, await timeDeck(manyRows)); }
    assert.ok(manyMs <= 8 * fewMs + 250, `Equation pictures: PPTX nested table of equations linear (${Math.round(fewMs)} ms at 150, ${Math.round(manyMs)} ms at 600)`);
    console.log(`  ✓ PPTX, a nested table of equations: ${Math.round(fewMs)} ms at 150 rows, ${Math.round(manyMs)} ms at 600`);

    // 5. A picture that is missing: the LaTeX, nothing reported. One of a kind no writer takes: the image path's alt text, reported.
    const equationAst = (attachments: any[]): OfficeParserAST => ({
        type: 'rtf', metadata: {}, attachments,
        content: [{ type: 'paragraph', text: 'A yz^2 B', children: [{ type: 'text', text: 'A ' }, { type: 'code', text: 'yz^2', metadata: { math: 'inline', fallbackAttachmentName: 'p.pict' } }, { type: 'text', text: ' B' }] }],
    } as any);
    const missing = equationAst([]);
    for (const format of ['docx', 'odt', 'rtf', 'pptx', 'odp']) {
        const out = await write(missing, format);
        const text = format === 'rtf' ? String(out.value) : Object.values(zipTexts(out.value)).join('');
        assert.ok(text.includes('yz^2') && !out.codes.includes('IMAGE_PROCESSING_FAILED'), `Equation pictures: ${format}, a missing picture is the LaTeX, not reported (${out.codes})`);
    }
    const unknown = equationAst([{ type: 'image', name: 'p.pict', mimeType: 'image/x-pict', extension: 'pict', data: Buffer.from(Array.from({ length: 16 }, (_, i) => 0x11 * i % 251)).toString('base64') }]);
    for (const format of ['docx', 'odt', 'rtf', 'pptx', 'odp']) {
        const out = await write(unknown, format);
        const text = format === 'rtf' ? String(out.value) : format === 'docx' ? docxText(zipTexts(out.value)) : format === 'pptx' ? pptxText(zipTexts(out.value)) : odfText(zipTexts(out.value));
        assert.ok(text.includes('yz^2') && out.warnings.some(w => w.code === 'IMAGE_PROCESSING_FAILED' && /output does not take/.test(w.message)), `Equation pictures: ${format}, a picture it does not take is its alt text, reported (${out.warnings.map(w => w.message)})`);
    }

    // 7. Property: equations with and without their pictures, in paragraphs and one-cell tables, through DOCX, ODT and RTF.
    const rnd = lcg(0xE9A7);
    const chars = ['a', 'x', '1', '2', ' ', '<', '&', '"', '>', '\\', '{', '}', '%', '$', '_', '^', '~', '#', String.fromCharCode(0x3B1), String.fromCharCode(0x2211)];
    const png = hexOf(TINY_PNG);
    const seen = new Set<string>();
    let equations = 0;
    for (let round = 0; round < 2600 && seen.size < 2000; round++) {
        let body = '', count = 0;
        const pictured: string[] = [], plain: string[] = [];
        for (let p = 1 + Math.floor(rnd() * 6); p > 0; p--) {
            let line = '';
            for (let items = 1 + Math.floor(rnd() * 4); items > 0; items--) {
                if (rnd() < 0.5) { line += ` W${Math.floor(rnd() * 90)} `; continue; }
                const marker = `Q${count++}Z`;
                const more = Array.from({ length: Math.floor(rnd() * 4) }, () => chars[Math.floor(rnd() * chars.length)]).join('');
                const withPicture = rnd() < 0.5;
                line += `{\\mmath {\\*\\moMath {\\mr ${marker}}${more ? `{\\mr ${rtfText(more)}}` : ''}}${withPicture ? `{\\mmathPict {\\*\\shppict{\\pict\\picw1\\pich1\\pngblip ${png}}}}` : ''}}`;
                (withPicture ? pictured : plain).push(marker);
            }
            body += rnd() < 0.2 ? `\\trowd\\cellx3000\\pard\\intbl ${line}\\cell\\row` : `\\pard ${line}\\par`;
        }
        if (!count || seen.has(body)) continue;
        seen.add(body);
        equations += count;
        const doc = await OfficeParser.parseOffice(Buffer.from(rtfOf(body), 'latin1'), { extractAttachments: true });
        const latexOf = new Map(allNodes(doc).filter(n => n.type === 'code').map(n => [/Q\d+Z/.exec(n.text ?? '')![0], n.text!]));
        assert.strictEqual(latexOf.size, count, `Equation pictures: every equation read (${body})`);
        for (const format of ['docx', 'odt', 'rtf']) {
            const value = (await doc.to(format as any, { onWarning: () => { } } as any)).value;
            const back = await OfficeParser.parseOffice(typeof value === 'string' ? Buffer.from(value, 'latin1') : Buffer.from(bytesOf(value)), { extractAttachments: true });
            const images = allNodes(back).filter(n => n.type === 'image');
            const text = allNodes(back).filter(n => n.type === 'paragraph' || n.type === 'cell').map(n => n.text ?? '').join('\n');
            assert.strictEqual(images.length, pictured.length, `Equation pictures: ${format} writes one picture an equation with one (${body})`);
            for (const marker of plain) assert.ok(text.includes(latexOf.get(marker)!), `Equation pictures: ${format} writes ${JSON.stringify(latexOf.get(marker))} as its LaTeX (${body})`);
            for (const marker of pictured) assert.ok(!text.includes(marker), `Equation pictures: ${format} writes no LaTeX for ${marker} (${body})`);
            if (format !== 'rtf') assert.deepStrictEqual(images.map(n => (n.metadata as any)?.altText).sort(), pictured.map(m => latexOf.get(m)!).sort(), `Equation pictures: ${format} alt texts are the pictured equations' LaTeX (${body})`);
        }
    }
    assert.ok(seen.size >= 2000, `Equation pictures: ${seen.size} distinct documents`);
    console.log(`  ✓ ${seen.size} distinct documents (${equations} equations) written as their pictures or their LaTeX`);

    // 8. Linear: many LibreOffice-shaped equations, each its own picture.
    const many = (n: number) => OfficeParser.parseOffice(Buffer.from(rtfOf(`\\pard ${`A${group}B `.repeat(n)}\\par`), 'latin1'), { extractAttachments: true });
    const small = await many(1000), large = await many(4000);
    assert.strictEqual(large.attachments.length, 4000, 'Equation pictures: each equation its own picture');
    for (const format of ['docx', 'rtf', 'pptx', 'pdf']) {
        const time = async (doc: OfficeParserAST) => { const started = performance.now(); await write(doc, format); return performance.now() - started; };
        let smallMs = await time(small), largeMs = await time(large);
        if (largeMs > 8 * smallMs + 250) { smallMs = Math.min(smallMs, await time(small)); largeMs = Math.min(largeMs, await time(large)); }
        assert.ok(largeMs <= 8 * smallMs + 250, `Equation pictures: ${format} linear (${Math.round(smallMs)} ms at 1000, ${Math.round(largeMs)} ms at 4000)`);
        console.log(`  ✓ Equation pictures, ${format}: ${Math.round(smallMs)} ms at 1000, ${Math.round(largeMs)} ms at 4000`);
    }
    console.log('  ✓ RTF equations written as their own pictures where math cannot be typeset');
}

// ─── RTF embedded objects ────────────────────────────────────────────────

/** RTF embedded objects (`\object`) are read as the picture or text they show (`\result`). */
export async function testRtfEmbeddedObjects(): Promise<void> {
    console.log('\n=== RTF embedded objects ===');
    const read = (body: string, config: OfficeParserConfig = {}) => parseWithWarnings(rtfOf(body), { extractAttachments: true, ...config });
    const pictures = (ast: OfficeParserAST) => allNodes(ast).filter(n => n.type === 'image');
    const png = hexOf(TINY_PNG);
    const object = (progId: string, result: string) => `{\\object\\objemb{\\*\\objclass ${progId}}\\objw1440\\objh720{\\*\\objdata 01050000020000000b000000}${result}}`;

    // The kinds: listed ProgIDs by name, any equation as Equation, an unlisted one as written.
    for (const [progId, kind] of [['Excel.Sheet.12', 'Excel worksheet'], ['Excel.SheetMacroEnabled.12', 'Excel worksheet'], ['Excel.Chart.8', 'Excel chart'], ['Word.Document.12', 'Word document'], ['PowerPoint.Show.12', 'PowerPoint presentation'], ['PowerPoint.Slide.12', 'PowerPoint slide'], ['Equation.3', 'Equation'], ['Equation.DSMT4', 'Equation'], ['Visio.Drawing.15', 'Visio drawing'], ['MSGraph.Chart.8', 'Chart'], ['AcroExch.Document.DC', 'PDF document'], ['Package', 'Embedded file'], ['Paint.Picture', 'Picture'], ['PBrush', 'Picture'], ['Acme.Widget.2', 'Acme.Widget.2'], ['__proto__', '__proto__'], ['constructor', 'constructor']] as const) {
        assert.strictEqual(embeddedObjectKind(progId), kind, `RTF objects: ${progId} is ${kind}`);
    }
    assert.strictEqual(embeddedObjectKind(undefined), undefined, 'RTF objects: no ProgID, no kind');

    // Word's shape: the result a \shppict PNG and a \nonshppict WMF of it; one picture, named by its kind.
    const word = await read(`\\pard BEFOREOBJ ${object('Excel.Sheet.12', `{\\result{\\*\\shppict{\\pict\\pngblip ${png}}}{\\nonshppict{\\pict\\wmetafile8 ${hexOf(headerlessWmf())}}}}`)} AFTER\\par`);
    assert.strictEqual(pictures(word.ast).length, 1, 'RTF objects: one picture');
    assert.strictEqual((pictures(word.ast)[0].metadata as any).altText, 'Excel worksheet', 'RTF objects: its alt text the object\'s kind');
    assert.ok(word.ast.attachments.length === 1 && word.ast.attachments[0].mimeType === 'image/png' && word.codes.length === 0, 'RTF objects: the PNG, nothing reported');
    assert.ok((await word.ast.to('text')).value.includes('[Image: Excel worksheet]'), 'RTF objects: text output names the object');

    // A text result is text, in its paragraph; r4-hand's shape.
    const text = await read(`\\pard MIDOBJ ${object('Equation.3', '{\\result OBJRESULTTEXT}')} AFTEROBJ\\par`);
    assert.strictEqual(text.ast.content.map(n => n.text).join('|'), 'MIDOBJ OBJRESULTTEXT AFTEROBJ', `RTF objects: a text result read in place (${text.ast.content.map(n => n.text)})`);
    const textBare = await read(`\\pard MIDOBJ ${object('Equation.3', '{\\result OBJRESULTTEXT}')} AFTEROBJ\\par`, { extractAttachments: false });
    assert.ok(textBare.ast.content.map(n => n.text).join('|') === 'MIDOBJ OBJRESULTTEXT AFTEROBJ' && textBare.codes.length === 0, 'RTF objects: a text result read without extractAttachments, nothing reported');

    // No result: a picture of its kind alone, reported with its ProgID.
    const none = await read(`\\pard A ${object('Excel.Sheet.12', '')} B\\par`);
    assert.ok(pictures(none.ast).length === 1 && (pictures(none.ast)[0].metadata as any).altText === 'Excel worksheet' && !(pictures(none.ast)[0].metadata as any).attachmentName, 'RTF objects: no result, an alt-only picture');
    assert.deepStrictEqual(none.warnings.map(w => [w.code, w.details?.part, w.details?.progId]), [['CONTENT_PART_NOT_READ', 'the embedded Excel.Sheet.12 object', 'Excel.Sheet.12']], 'RTF objects: no result reported');
    assert.ok(none.warnings[0].message.includes('the embedded Excel.Sheet.12 object') && none.warnings[0].message.includes('no \\result'), `RTF objects: the message (${none.warnings[0].message})`);
    const noClass = await read('\\pard {\\object\\objemb\\objw10\\objh10}\\par');
    assert.ok(pictures(noClass.ast).length === 1 && !pictures(noClass.ast)[0].metadata && noClass.warnings[0]?.details?.part === 'an embedded object', 'RTF objects: no ProgID, "an embedded object"');
    const noneBare = await read(`\\pard A ${object('Excel.Sheet.12', '')} B\\par`, { extractAttachments: false });
    assert.ok(pictures(noneBare.ast).length === 0 && noneBare.codes.length === 0, 'RTF objects: no result without extractAttachments, nothing (pictures not asked for)');
    const pictureBare = await read(`\\pard ${object('Excel.Sheet.12', `{\\result{\\pict\\pngblip ${png}}}`)}\\par`, { extractAttachments: false });
    assert.ok(pictures(pictureBare.ast).length === 0 && pictureBare.codes.length === 0, 'RTF objects: a picture result without extractAttachments, nothing');

    // In a tracked deletion: nothing. A ProgID of __proto__, and one of 10,000 characters cut to 128.
    const gone = await read(`\\pard A{\\deleted ${object('Excel.Sheet.12', `{\\result{\\pict\\pngblip ${png}}}`)}}B\\par`);
    assert.ok(pictures(gone.ast).length === 0 && gone.ast.attachments.length === 0 && gone.codes.length === 0 && gone.ast.content[0].text === 'AB', 'RTF objects: a deleted object is nothing');
    const proto = await read(`\\pard ${object('__proto__', `{\\result{\\pict\\pngblip ${png}}}`)}\\par`);
    assert.strictEqual((pictures(proto.ast)[0].metadata as any).altText, '__proto__', 'RTF objects: a ProgID of __proto__ is just unlisted');
    const long = await read(`\\pard ${object('X'.repeat(10000), '')}\\par`);
    assert.ok((pictures(long.ast)[0].metadata as any).altText === 'X'.repeat(128) && long.warnings[0].details.progId === 'X'.repeat(128), 'RTF objects: a long ProgID cut to 128 characters');
    console.log('  ✓ RTF embedded objects read as the picture or text they show');
}

// ─── Content-control prompts in every writer, found in linear time ───────

const PROMPT = 'Click or tap here to enter text.';
/** A content control showing its placeholder, the glossary building block `docPart` its prompt. */
const promptControl = (content: string, docPart = 'DefaultPlaceholder_1') =>
    `<w:sdt><w:sdtPr><w:placeholder><w:docPart w:val="${docPart}"/></w:placeholder><w:showingPlcHdr/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
/** A glossary whose building blocks are `parts` (name and content of its paragraph). */
const glossaryOf = (parts: [string, string][]) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:glossaryDocument ${W_NS}><w:docParts>${parts.map(([name, content]) => `<w:docPart><w:docPartPr><w:name w:val="${name}"/></w:docPartPr><w:docPartBody><w:p>${content}</w:p></w:docPartBody></w:docPart>`).join('')}</w:docParts></w:glossaryDocument>`;

/** Every writer leaves a content control's prompt out, and writes it with includePlaceholderText; the native PDF engine too. */
export async function testWordPlaceholderWriters(): Promise<void> {
    console.log('\n=== Content-control prompts in every writer ===');
    const body = para(run('Name: '), promptControl(run(PROMPT)), run(' (required)')) + promptControl(para(run(PROMPT))) + para(run('Applicant: Ada'))
        + `<w:tbl><w:tr><w:tc>${para(run('CELL1'))}</w:tc>${promptControl(`<w:tc>${para(run(PROMPT))}</w:tc>`)}</w:tr></w:tbl>`;
    const ast = await OfficeParser.parseOffice(docxOf(body, { extra: { 'word/glossary/document.xml': glossaryOf([['DefaultPlaceholder_1', run(PROMPT)]]) } }));
    const textOf = async (format: string, config: any = {}): Promise<string> => {
        const value = (await ast.to(format as any, { onWarning: () => { }, ...config })).value as any;
        if (format === 'chunks') return (value as { text: string }[]).map(c => c.text).join('\n');
        if (typeof value === 'string') return value;
        // A binary output read back as text.
        const back = await OfficeParser.parseOffice(Buffer.from(bytesOf(value)), {});
        return (await back.to('text')).value as string;
    };
    const formats: [string, any][] = [['html', {}], ['md', {}], ['docx', {}], ['odt', {}], ['rtf', {}], ['tex', {}], ['pptx', {}], ['odp', {}], ['xlsx', {}], ['ods', {}], ['text', {}], ['chunks', {}], ['pdf', { pdfConfig: { engine: 'native' } }]];
    for (const [format, config] of formats) {
        const out = await textOf(format, config);
        assert.ok(!out.includes('Click or tap'), `Prompts: ${format} leaves the prompt out`);
        assert.ok(out.includes('Name:') && out.includes('(required)') && out.includes('Applicant: Ada') && out.includes('CELL1'), `Prompts: ${format} keeps the text around it (${JSON.stringify(out.slice(0, 200))})`);
        const shown = await textOf(format, { ...config, includePlaceholderText: true });
        // (A writer may write an inline prompt's text twice only where it writes a paragraph's text twice, as
        // a spreadsheet cell's or a slide's notes would: counted once a place.)
        assert.strictEqual(shown.replace(/\s+/g, ' ').split(PROMPT).length - 1, 3, `Prompts: ${format} with includePlaceholderText writes each prompt once (${JSON.stringify(shown.slice(0, 300))})`);
    }
    // A caller's onNode string for a prompt wins, in HTML and in the native PDF engine.
    const onNode = (node: OfficeContentNode) => (node.metadata as any)?.placeholder ? 'FILLEDIN' : undefined;
    const html = String((await ast.to('html', { onNode } as any)).value);
    assert.ok(html.includes('FILLEDIN') && !html.includes('Click or tap'), 'Prompts: onNode\'s string for a prompt is written in HTML');
    const pdf = await OfficeParser.parseOffice(Buffer.from(bytesOf((await ast.to('pdf', { onNode, onWarning: () => { }, pdfConfig: { engine: 'native' } } as any)).value)));
    const pdfText = (await pdf.to('text')).value as string;
    assert.ok(pdfText.includes('FILLEDIN') && !pdfText.includes('Click or tap'), `Prompts: onNode's string for a prompt is drawn in native PDF (${JSON.stringify(pdfText.slice(0, 200))})`);

    // Chunks: a prompt below a node's children (a footnote's, an endnote's paragraph) is written with the option too.
    const placed = para(run('BODY1 '), promptControl(run(PROMPT)), run(' BODY2'))
        + `<w:tbl><w:tr><w:tc>${para(run('CELLA'))}</w:tc><w:tc>${para(promptControl(run(PROMPT)))}</w:tc></w:tr></w:tbl>`
        + `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${run('HEAD ')}${promptControl(run(PROMPT))}</w:p>`
        + para(run('FN'), '<w:r><w:footnoteReference w:id="1"/></w:r>')
        + para(run('EN'), '<w:r><w:endnoteReference w:id="1"/></w:r>')
        + para(`<w:hyperlink r:id="rL">${promptControl(run(PROMPT))}${run('LINKTEXT')}</w:hyperlink>`)
        + '<w:sectPr><w:headerReference w:type="default" r:id="rH"/></w:sectPr>';
    const placedAst = await OfficeParser.parseOffice(docxOf(placed, {
        rels: [['rH', 'header', 'header1.xml'], ['rL', 'hyperlink', 'http://example.com'], ['rE', 'endnotes', 'endnotes.xml']],
        footnotes: `<w:footnote w:id="1">${para(run('NOTE '), promptControl(run(PROMPT)))}</w:footnote>`,
        extra: {
            'word/glossary/document.xml': glossaryOf([['DefaultPlaceholder_1', run(PROMPT)]]),
            'word/header1.xml': `<?xml version="1.0"?><w:hdr ${W_NS}>${para(run('HDR '), promptControl(run(PROMPT)))}</w:hdr>`,
            'word/endnotes.xml': `<?xml version="1.0"?><w:endnotes ${W_NS}><w:endnote w:id="1">${para(run('ENDNOTE '), promptControl(run(PROMPT)))}</w:endnote></w:endnotes>`,
        },
    }));
    const chunkTexts = async (config: any) => ((await placedAst.to('chunks', config)).value as { text: string }[]).map(c => c.text);
    assert.deepStrictEqual(await chunkTexts({ includePlaceholderText: true }), [`BODY1 ${PROMPT} BODY2`, `| CELLA | ${PROMPT} |`, `HEAD ${PROMPT}`, `FN\nNOTE ${PROMPT}`, `EN\nENDNOTE ${PROMPT}`, `${PROMPT}LINKTEXT`], 'Prompts: chunks write each prompt with includePlaceholderText, a footnote\'s and an endnote\'s too');
    assert.deepStrictEqual(await chunkTexts({}), ['BODY1  BODY2', '| CELLA |  |', 'HEAD', 'FN\nNOTE', 'EN\nENDNOTE', 'LINKTEXT'], 'Prompts: chunks leave every prompt out by default');

    // A glossary that is not well-formed XML is no glossary: the prompt read as text, and reported once.
    // Over the element budget, the parse still ends.
    const prompted = para(run('Name: '), promptControl(run(PROMPT)));
    for (const [label, glossary] of [['corrupt', '<<<not xml'], ['truncated', `<?xml version="1.0"?><w:glossaryDocument ${W_NS}><w:docParts><w:docPart>`]] as const) {
        const { ast: read, warnings } = await parseWithWarnings(docxOf(prompted, { extra: { 'word/glossary/document.xml': glossary } }));
        assert.strictEqual(read.content[0].text, `Name: ${PROMPT}`, `Prompts: a ${label} glossary, the prompt read as text (${read.content[0].text})`);
        assert.ok(!allNodes(read).some(n => (n.metadata as any)?.placeholder), `Prompts: a ${label} glossary, no prompt marked`);
        assert.deepStrictEqual(warnings.map(w => [w.code, w.details?.part]), [['CONTENT_PART_NOT_READ', 'the glossary, word/glossary/document.xml']], `Prompts: a ${label} glossary reported once (${warnings.map(w => w.message)})`);
        assert.ok(/not well-formed XML/.test(warnings[0].message) && /read as text/.test(warnings[0].message), `Prompts: the message says why and what follows (${warnings[0].message})`);
    }
    const big = `<w:glossaryDocument ${W_NS}><w:docParts>${'<w:p/>'.repeat(200000)}</w:docParts></w:glossaryDocument>`;
    await assert.rejects(OfficeParser.parseOffice(docxOf(prompted, { extra: { 'word/glossary/document.xml': big } }), { decompressionLimits: { maxXmlElements: 100000 } } as any), (e: any) => e?.officeIssue?.code === 'XML_ELEMENT_LIMIT_EXCEEDED' || /XML element limit/i.test(String(e?.message)), 'Prompts: a glossary over the element budget still ends the parse');
    console.log('  ✓ every writer leaves prompts out, and writes them when asked');
}

/** Content-control prompts are found in linear time: nested controls, plain and flagged with a glossary entry. */
export async function testWordPlaceholdersLinear(): Promise<void> {
    console.log('\n=== Content-control prompts found in linear time ===');
    const glossary = { 'word/glossary/document.xml': glossaryOf([['P', run('PROMPT')]]) };
    const nested = (flagged: boolean) => (n: number) => {
        const open = flagged ? '<w:sdt><w:sdtPr><w:placeholder><w:docPart w:val="P"/></w:placeholder><w:showingPlcHdr/></w:sdtPr><w:sdtContent>' : '<w:sdt><w:sdtContent>';
        return docxOf(`${open}${para(run('x'))}`.repeat(n) + '</w:sdtContent></w:sdt>'.repeat(n), { extra: glossary });
    };
    for (const flagged of [false, true]) {
        await linear(`Prompts: 8,000 nested controls${flagged ? ', flagged' : ''}`, nested(flagged), 8000, {}, (ast, n) => assert.strictEqual(ast.content.length, n, 'Prompts: every nested paragraph read'));
    }
    // Many prompts side by side, and one prompt control of a long text that is not its prompt.
    await linear('Prompts: 20,000 prompt controls', n => docxOf(para(promptControl(run('PROMPT'), 'P')).repeat(n) + para(run('END')), { extra: glossary }), 20000, {}, ast => assert.strictEqual(ast.content.filter(p => p.text === 'END').length, 1, 'Prompts: the text after them read'));
    const long = await OfficeParser.parseOffice(docxOf(para(promptControl(run('PROMPT' + 'y'.repeat(1_000_000)), 'P')), { extra: glossary }));
    assert.strictEqual(long.content[0].text?.length, 1_000_006, 'Prompts: a control holding more than its prompt is read as text');
    const exact = await OfficeParser.parseOffice(docxOf(para(run('A'), promptControl(run('PROMPT'), 'P'), run('B')), { extra: glossary }));
    assert.strictEqual(exact.content[0].text, 'AB', 'Prompts: a control holding its prompt is still found');
    // Chunks with the option, for the nested flagged controls: each node's prompts found once.
    const nestedAt = async (n: number) => OfficeParser.parseOffice(nested(true)(n));
    const [smallDoc, largeDoc] = [await nestedAt(2000), await nestedAt(8000)];
    const chunkTime = async (doc: OfficeParserAST) => { const started = performance.now(); await doc.to('chunks', { includePlaceholderText: true } as any); return performance.now() - started; };
    let smallMs = await chunkTime(smallDoc), largeMs = await chunkTime(largeDoc);
    if (largeMs > 8 * smallMs + 250) { smallMs = Math.min(smallMs, await chunkTime(smallDoc)); largeMs = Math.min(largeMs, await chunkTime(largeDoc)); }
    assert.ok(largeMs <= 8 * smallMs + 250, `Prompts: chunks with includePlaceholderText linear (${Math.round(smallMs)} ms at 2000, ${Math.round(largeMs)} ms at 8000)`);
    console.log(`  ✓ Prompts: chunks of 8,000 nested controls with includePlaceholderText: ${Math.round(smallMs)} ms at 2000, ${Math.round(largeMs)} ms at 8000`);
    console.log('  ✓ prompts found in linear time');
}

// ─── DOCX embedded objects ───────────────────────────────────────────────

/** DOCX embedded objects (`w:object`) are read as their preview pictures, their alt text naming their kind. */
export async function testWordEmbeddedObjects(): Promise<void> {
    console.log('\n=== DOCX embedded objects ===');
    const read = (docx: Buffer, config: OfficeParserConfig = {}) => parseWithWarnings(docx, { extractAttachments: true, ...config });
    const pictures = (nodes: OfficeContentNode[]) => nodes.filter(n => n.type === 'image');

    // The c5-objects shape: an Excel sheet (PNG preview), Equation 3.0 (WMF) and MathType (EMF).
    const { ast, codes } = await read(c5Objects());
    const holding = ast.content.filter(p => pictures(p.children ?? []).length);
    assert.strictEqual(holding.length, 3, 'Objects: three paragraphs hold a picture');
    const meta = holding.map(p => pictures(p.children!)[0].metadata as any);
    assert.deepStrictEqual(meta.map(m => [m.attachmentName, m.altText]), [['image1.png', 'Excel worksheet'], ['image2.wmf', 'Equation'], ['image3.emf', 'Equation']], `Objects: each its preview, named by its kind (${JSON.stringify(meta)})`);
    assert.ok(holding.every(p => p.children!.length === 1), 'Objects: one picture in each paragraph, inline');
    assert.deepStrictEqual(codes, [], 'Objects: nothing reported');
    assert.deepStrictEqual(ast.attachments.map(a => a.mimeType), ['image/png', 'image/wmf', 'image/emf'], 'Objects: attachments typed');

    // Alt text: the shape's own wins over the kind; in the w:drawing and w:objectEmbed form, docPr's.
    const alt = await read(docxOf(para(oleRun('rIp', 'rO1', 'Excel.Sheet.12', 'Budget')) + para(`<w:r><w:object>${pictureRun('rIp', 7, 'Quarterly table').replace(/^<w:r>|<\/w:r>$/g, '')}<w:objectEmbed w:progId="Excel.Sheet.12" r:id="rO1"/></w:object></w:r>`), {
        rels: [['rIp', 'image', 'media/image1.png'], ['rO1', 'package', 'embeddings/sheet.xlsx']], media: { 'image1.png': TINY_PNG } }));
    assert.deepStrictEqual(pictures(allNodes(alt.ast)).map(n => (n.metadata as any).altText), ['Budget', 'Quarterly table'], 'Objects: the shape\'s alt text, then docPr\'s, win over the kind');

    // No preview, or a preview whose relationship is missing: a picture of the alt text alone, reported.
    const missing = await read(docxOf(para(oleRun(undefined, 'rO1', 'Excel.Sheet.12')) + para(oleRun('rGone', 'rO2', 'Equation.3')), {
        rels: [['rO1', 'package', 'embeddings/sheet.xlsx'], ['rO2', 'oleObject', 'embeddings/oleObject1.bin']] }));
    const missingPictures = pictures(allNodes(missing.ast));
    assert.deepStrictEqual(missingPictures.map(n => [(n.metadata as any)?.attachmentName, (n.metadata as any)?.altText]), [[undefined, 'Excel worksheet'], [undefined, 'Equation']], 'Objects: alt-only pictures');
    assert.deepStrictEqual(missing.warnings.map(w => [w.code, w.details?.part, w.details?.progId, w.details?.target]), [
        ['CONTENT_PART_NOT_READ', 'the embedded Excel.Sheet.12 object', 'Excel.Sheet.12', 'embeddings/sheet.xlsx'],
        ['CONTENT_PART_NOT_READ', 'the embedded Equation.3 object', 'Equation.3', 'embeddings/oleObject1.bin'],
    ], `Objects: each reported (${JSON.stringify(missing.warnings.map(w => w.details))})`);
    assert.ok(missing.warnings[0].message.includes('A part of the document (the embedded Excel.Sheet.12 object) was not read: it has no preview picture, and the object itself is not read.'), `Objects: the message (${missing.warnings[0].message})`);
    // An older producer's object in a w:pict, and one with no ProgID.
    const pict = await read(docxOf(para('<w:r><w:pict><v:shape id="s"><v:imagedata r:id="rIp"/></v:shape><o:OLEObject Type="Embed" ProgID="Word.Document.12" r:id="rO1"/></w:pict></w:r>') + para('<w:r><w:object><v:shape id="t"/></w:object></w:r>'), {
        rels: [['rIp', 'image', 'media/image1.png'], ['rO1', 'package', 'embeddings/doc.docx']], media: { 'image1.png': TINY_PNG } }));
    assert.deepStrictEqual(pictures(allNodes(pict.ast)).map(n => (n.metadata as any)?.altText), ['Word document', undefined], 'Objects: an object in w:pict named by its kind; one of no ProgID has no alt text');
    assert.deepStrictEqual(pict.warnings.map(w => w.details?.part), ['an embedded object'], 'Objects: one of no ProgID reported as "an embedded object"');

    // An ActiveX control (`w:control`) is not an embedded object: the picture it shows, or nothing, unreported.
    const controls = await read(docxOf(
        para('<w:r><w:object><v:shape id="s2"><v:imagedata r:id="rIp"/></v:shape><w:control r:id="rX" w:name="CheckBox1" w:shapeid="s2"/></w:object></w:r>')
        + para(run('BEFORECTL'), '<w:r><w:object><v:shape id="s3"/><w:control r:id="rX" w:name="TextBox1" w:shapeid="s3"/></w:object></w:r>', run('AFTERCTL')), {
        rels: [['rIp', 'image', 'media/image1.png'], ['rX', 'control', 'activeX/activeX1.xml']], media: { 'image1.png': TINY_PNG } }));
    assert.deepStrictEqual(pictures(allNodes(controls.ast)).map(n => n.metadata), [{ attachmentName: 'image1.png', altText: '' }], 'Objects: a control showing a picture is that picture; one showing none is no node');
    assert.ok(controls.ast.content[1].text === 'BEFORECTLAFTERCTL' && controls.codes.length === 0, `Objects: the text around a control kept, nothing reported (${controls.ast.content[1].text}; ${controls.codes})`);
    // A control beside an embedded object's own element stays an object: no preview, reported.
    const controlObject = await read(docxOf(para('<w:r><w:object><v:shape id="s4"/><w:control r:id="rX" w:name="X"/><o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" r:id="rO1"/></w:object></w:r>'), {
        rels: [['rX', 'control', 'activeX/activeX1.xml'], ['rO1', 'package', 'embeddings/sheet.xlsx']] }));
    assert.deepStrictEqual(pictures(allNodes(controlObject.ast)).map(n => (n.metadata as any)?.altText), ['Excel worksheet'], 'Objects: a control with o:OLEObject is still the object');
    assert.deepStrictEqual(controlObject.warnings.map(w => [w.code, w.details?.progId]), [['CONTENT_PART_NOT_READ', 'Excel.Sheet.12']], 'Objects: and reported as one with no preview');

    // Placed where they stand: a footnote, a table cell, the fallback of alternate content.
    const placed = await read(docxOf(
        para(run('NOTE'), '<w:r><w:footnoteReference w:id="1"/></w:r>')
        + `<w:tbl><w:tr><w:tc>${para(oleRun('rIp', 'rO1', 'Excel.Sheet.12'))}</w:tc></w:tr></w:tbl>`
        + para(`<w:r><mc:AlternateContent><mc:Choice Requires="x99"><w:t>NEVER</w:t></mc:Choice><mc:Fallback>${oleRun('rIp', 'rO1', 'Equation.3').replace(/^<w:r>|<\/w:r>$/g, '')}</mc:Fallback></mc:AlternateContent></w:r>`), {
        rels: [['rIp', 'image', 'media/image1.png'], ['rO1', 'package', 'embeddings/sheet.xlsx']], media: { 'image1.png': TINY_PNG },
        footnotes: `<w:footnote w:id="1">${para(oleRun('rFp', 'rFo', 'Equation.DSMT4'))}</w:footnote>`, footRels: [['rFp', 'image', 'media/image1.png'], ['rFo', 'oleObject', 'embeddings/o.bin']] }));
    const note = allNodes(placed.ast).find(n => n.type === 'note')!;
    const cell = allNodes(placed.ast).find(n => n.type === 'cell')!;
    assert.strictEqual((pictures(allNodes({ content: [note] } as any))[0]?.metadata as any)?.altText, 'Equation', 'Objects: in a footnote');
    assert.strictEqual((pictures(allNodes({ content: [cell] } as any))[0]?.metadata as any)?.altText, 'Excel worksheet', 'Objects: in a table cell');
    assert.ok(pictures(allNodes(placed.ast)).some(n => (n.metadata as any)?.altText === 'Equation' && !allNodes({ content: [note] } as any).includes(n)), 'Objects: in the fallback of alternate content');

    // Without attachments: no pictures, nothing reported, as for every picture.
    const bare = await read(c5Objects(), { extractAttachments: false });
    assert.ok(pictures(allNodes(bare.ast)).length === 0 && bare.codes.length === 0, 'Objects: without extractAttachments, nothing');
    const bareMissing = await read(docxOf(para(oleRun(undefined, 'rO1', 'Excel.Sheet.12'))), { extractAttachments: false });
    assert.deepStrictEqual(bareMissing.codes, [], 'Objects: without extractAttachments, one with no preview is not reported');

    // Writers.
    const md = String((await ast.to('md')).value), text = String((await ast.to('text')).value), html = String((await ast.to('html')).value);
    const chunks = ((await ast.to('chunks')).value as { text: string }[]).map(c => c.text).join('\n');
    assert.ok(md.includes('![Excel worksheet]('), 'Objects: Markdown names the sheet');
    assert.ok(text.includes('[Image: Equation]') && chunks.includes('[Image: Equation]'), 'Objects: text and chunks name the equation');
    assert.ok(html.includes('alt="Excel worksheet"'), 'Objects: HTML alt text');
    const docx = unzipSync(bytesOf((await ast.to('docx')).value));
    assert.ok(Object.keys(docx).some(k => /^word\/media\/.*\.wmf$/.test(k)) && Object.keys(docx).some(k => /^word\/media\/.*\.emf$/.test(k)), 'Objects: DOCX output packages the WMF and EMF previews');

    // Linear: many objects in one paragraph.
    await linear('Objects: 20,000 in one paragraph', n => docxOf(para(oleRun('rIp', 'rO1', 'Excel.Sheet.12').repeat(n)), { rels: [['rIp', 'image', 'media/image1.png'], ['rO1', 'package', 'embeddings/sheet.xlsx']], media: { 'image1.png': TINY_PNG } }), 20000, { extractAttachments: true }, (doc, n) => assert.strictEqual(pictures(doc.content[0].children ?? []).length, n, 'Objects: every object a picture'));
    console.log('  ✓ DOCX embedded objects read as their previews');
}

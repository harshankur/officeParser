/**
 * CVE-2026-111497: on a runtime whose Brotli decoder has the flaw, a PDF's Brotli stream and a WOFF2 font
 * are not given to it, what they hold is still read, and a warning says so (see src/utils/brotliFlaw.ts).
 * The streams here are ordinary ones; what is tested is who decodes them. The flaw is set by a hook, so
 * this runs the same on a fixed Node as on one without the fix.
 */
import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import * as zlib from 'zlib';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParser } from '../../src/OfficeParser';
import { brotliFlawHooks, nodeHasBrotliFlaw, runtimeHasBrotliFlaw } from '../../src/utils/brotliFlaw';
import { toWoff2 } from '../helpers/webFontWriter';

const CODE = 'BROTLI_DECODER_AVOIDED';

function pdfOf(objects: (string | Buffer)[]): Buffer {
    const parts = [Buffer.from('%PDF-2.0\n')];
    const offsets: number[] = [];
    let position = parts[0].length;
    objects.forEach((object, i) => {
        const part = Buffer.concat([Buffer.from(`${i + 1} 0 obj\n`), Buffer.isBuffer(object) ? object : Buffer.from(object), Buffer.from('\nendobj\n')]);
        offsets.push(position);
        parts.push(part);
        position += part.length;
    });
    const rows = offsets.map(offset => `${String(offset).padStart(10, '0')} 00000 n \n`).join('');
    parts.push(Buffer.from(`xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${rows}trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${position}\n%%EOF\n`));
    return Buffer.concat(parts);
}
const stream = (dict: string, data: Buffer) => Buffer.concat([Buffer.from(`<< ${dict} /Length ${data.length} >>\nstream\n`), data, Buffer.from('\nendstream')]);
const pageOf = (content: Buffer, resources = '', extra: (string | Buffer)[] = []) => pdfOf([
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> ${resources} >> >>`,
    content,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    ...extra,
]);
const TEXT = 'BT /F1 12 Tf 20 100 Td (Brotli page) Tj ET';
const brotliContent = () => pageOf(stream('/Filter /BrotliDecode', zlib.brotliCompressSync(Buffer.from(TEXT))));
const brotliImage = () => pageOf(stream('', Buffer.from(`${TEXT} q 50 0 0 50 20 20 cm /I1 Do Q`)), '/XObject << /I1 6 0 R >>', [
    stream('/Type /XObject /Subtype /Image /Width 2 /Height 2 /BitsPerComponent 8 /ColorSpace /DeviceRGB /Filter /BrotliDecode', zlib.brotliCompressSync(Buffer.alloc(12, 200))),
]);
const plain = () => pageOf(stream('', Buffer.from(TEXT)));

export async function testBrotliGuard(): Promise<void> {
    const before = brotliFlawHooks.flawed;
    // A process of its own for each setting of the hook: a process keeps the word it was started with, and
    // one is reused only for the same memory limit.
    let memoryMb = 777;
    const read = async (bytes: Buffer, config: any = {}, separateProcess = true) => {
        const codes: string[] = [];
        const ast = await OfficeParser.parseOffice(bytes, { fileType: 'pdf', onWarning: (w: any) => codes.push(w.code), ...config, pdfParserConfig: { separateProcess, processMemoryMb: memoryMb } } as any);
        return { ast, codes, text: (await ast.to('text')).value as string };
    };
    try {
        // The versions: the first fixed release of each line, the one before it, and the lines that got no fix.
        for (const [version, flawed] of [
            ['22.13.0', true], ['22.23.3', true], ['22.23.4', false], ['22.24.0', false], ['23.11.0', true], ['24.7.0', true], ['24.22.0', true], ['24.22.1', false],
            ['24.23.0', false], ['25.9.0', true], ['26.0.0', true], ['26.12.0', true], ['26.12.1', false], ['26.13.0', false], ['27.0.0', false], ['20.17.0', true], ['v24.22.1', false],
        ] as const) assert.strictEqual(nodeHasBrotliFlaw(version), flawed, `Node ${version}`);

        for (const separateProcess of [true, false]) {
            const where = separateProcess ? 'separate process' : 'in process';
            memoryMb++;
            brotliFlawHooks.flawed = true;

            // What the stream holds is read, by pdf.js's own decoder, and the warning says so once.
            const content = await read(brotliContent(), {}, separateProcess);
            assert.ok(content.text.includes('Brotli page'), `${where}: a Brotli content stream is read: ${JSON.stringify(content.text)}`);
            assert.strictEqual(content.codes.filter(code => code === CODE).length, 1, `${where}: warned once: ${content.codes}`);
            const image = await read(brotliImage(), { extractAttachments: true }, separateProcess);
            assert.strictEqual(image.ast.attachments.length, 1, `${where}: a Brotli image is read`);
            assert.strictEqual(image.codes.filter(code => code === CODE).length, 1, `${where}: warned once for the image: ${image.codes}`);

            // A PDF with no Brotli stream changes nothing, and says nothing.
            const ordinary = await read(plain(), {}, separateProcess);
            assert.ok(ordinary.text.includes('Brotli page'), `${where}: an ordinary PDF is read`);
            assert.ok(!ordinary.codes.includes(CODE), `${where}: no warning for an ordinary PDF: ${ordinary.codes}`);

            // A stream no decoder can read is not read, and the parse goes on.
            const broken = await read(pageOf(stream('/Filter /BrotliDecode', Buffer.from([0xff, 0xff, 0xff, 0xff, 1, 2, 3]))), {}, separateProcess).catch(e => e);
            assert.ok(!(broken instanceof Error) || !/process ended|Brotli error/.test(broken.message), `${where}: a broken Brotli stream does not end the parse: ${broken}`);

            // On a fixed runtime: the same text, no warning.
            memoryMb++;
            brotliFlawHooks.flawed = false;
            const fixed = await read(brotliContent(), {}, separateProcess);
            assert.ok(fixed.text.includes('Brotli page'), `${where}: read on a fixed runtime`);
            assert.ok(!fixed.codes.includes(CODE), `${where}: no warning on a fixed runtime: ${fixed.codes}`);
            assert.strictEqual(fixed.text, content.text, `${where}: the same text from either decoder`);
        }

        // The host's own stream is its own, even with pdf.js in this process guarded.
        brotliFlawHooks.flawed = true;
        await read(brotliContent(), {}, false);
        const { readable, writable } = new (globalThis as any).DecompressionStream('brotli');
        const writer = writable.getWriter();
        writer.write(zlib.brotliCompressSync(Buffer.from('the host'))).catch(() => undefined);
        writer.close().catch(() => undefined);
        const chunks: Uint8Array[] = [];
        for await (const chunk of readable) chunks.push(chunk);
        assert.strictEqual(Buffer.concat(chunks).toString(), 'the host', 'the host can still use the Brotli stream');
        assert.ok(new (globalThis as any).DecompressionStream('gzip'), 'and any other');

        // A WOFF2 font: fontkit's decoder, and the warning, once for all its styles.
        const liberation = new Uint8Array(fs.readFileSync(path.join(path.dirname(require.resolve('pdfjs-dist/package.json')), 'standard_fonts', 'LiberationSans-Regular.ttf')));
        const woff2 = toWoff2(liberation).bytes;
        const written = async () => {
            const codes: string[] = [];
            const doc: any = { type: 'pdf', metadata: {}, attachments: [], content: [{ type: 'paragraph', text: 'Font', children: [{ type: 'text', text: 'Font' }] }], getImages: () => [] };
            const result = await OfficeGenerator.generate(doc, 'pdf', { pdfConfig: { engine: 'native', fonts: { regular: woff2, bold: woff2 } }, onWarning: (w: any) => codes.push(w.code) } as any);
            return { codes, bytes: (result.value as Uint8Array).length };
        };
        const avoided = await written();
        assert.strictEqual(avoided.codes.filter(code => code === CODE).length, 1, `a WOFF2 font warns once: ${avoided.codes}`);
        brotliFlawHooks.flawed = false;
        const native = await written();
        assert.ok(!native.codes.includes(CODE), `no warning for a font on a fixed runtime: ${native.codes}`);
        assert.strictEqual(avoided.bytes, native.bytes, 'the same PDF from either decoder');

        // The real answer for this runtime is a yes or a no, and a browser, with no process, has no flaw.
        brotliFlawHooks.flawed = undefined;
        assert.strictEqual(typeof runtimeHasBrotliFlaw(), 'boolean');
    } finally {
        brotliFlawHooks.flawed = before;
    }
}

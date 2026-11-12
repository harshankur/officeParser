/**
 * Hostile shapes of the Word and RTF content readers and writers: metafile headers whose sizes are out of
 * range, RTF equations whose text holds markup, RTF pictures whose sizes no RTF parameter takes, and equations
 * written as their pictures whose LaTeX (the alt text) holds markup.
 */
import { unzipSync } from 'fflate';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { HtmlGenerator } from '../../src/generators/HtmlGenerator';
import { parseXmlString } from '../../src/utils/xmlUtils';
import { OfficeParser } from '../../src/OfficeParser';
import type { OfficeParserAST } from '../../src/types';
import { metafileSize, placeableWmf } from '../../src/utils/metafileUtils';

type Check = (name: string, cond: boolean, detail?: string) => void;

const u16 = (v: number) => [v & 0xFF, (v >>> 8) & 0xFF];
const u32 = (v: number) => [v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF];

/** An EMF header whose frame is the four values given (any int32), and EMR_EOF. */
const emfFramed = (left: number, top: number, right: number, bottom: number) => new Uint8Array([
    ...u32(1), ...u32(88), ...u32(0), ...u32(0), ...u32(10), ...u32(10),
    ...u32(left), ...u32(top), ...u32(right), ...u32(bottom),
    0x20, 0x45, 0x4D, 0x46, ...u32(0x10000), ...u32(108), ...u32(2), ...u16(1), ...u16(0),
    ...u32(0), ...u32(0), ...u32(0), ...u32(1920), ...u32(1080), ...u32(508), ...u32(286),
    ...u32(14), ...u32(20), ...u32(0), ...u32(16), ...u32(20),
]);

const astOf = (attachment: { name: string; mimeType: string; bytes: Uint8Array }): OfficeParserAST => ({
    type: 'docx', metadata: {}, content: [{ type: 'paragraph', text: '', children: [{ type: 'image', text: '', metadata: { attachmentName: attachment.name } }] }],
    attachments: [{ type: 'image', name: attachment.name, mimeType: attachment.mimeType, extension: attachment.name.split('.').pop(), data: Buffer.from(attachment.bytes).toString('base64') }],
} as any);

export async function wordRtfContentHardeningTests(check: Check): Promise<void> {
    console.log('- Metafile sizes, RTF equations and RTF picture sizes held to their ranges...');
    const silent = { onWarning: () => { } };

    // An EMF frame that is negative, empty or overflows: no size, so the writer's default size is used.
    for (const [label, frame] of [
        ['negative', [100, 100, -2000, -2000]],
        ['empty', [0, 0, 0, 0]],
        ['overflowing', [-2147483648, -2147483648, 2147483647, 2147483647]],
    ] as const) {
        const emf = emfFramed(frame[0], frame[1], frame[2], frame[3]);
        const size = metafileSize(emf);
        check(`metafile: an EMF of a ${label} frame has no size, or a finite positive one`, size === null || (Number.isFinite(size.w) && size.w > 0 && Number.isFinite(size.h) && size.h > 0), JSON.stringify(size));
        const docx = unzipSync((await OfficeGenerator.generate(astOf({ name: 'x.emf', mimeType: 'image/emf', bytes: emf }), 'docx', silent as any)).value as Uint8Array);
        const xml = new TextDecoder().decode(docx['word/document.xml']);
        const ext = /<wp:extent cx="(-?\d+)" cy="(-?\d+)"\/>/.exec(xml);
        check(`metafile: DOCX extent of a ${label} EMF frame is whole and positive`, !!ext && Number(ext[1]) > 0 && Number(ext[2]) > 0 && Number.isSafeInteger(Number(ext[1])), ext?.[0] ?? 'no extent');
    }

    // A WMF window of extreme values, records claiming more words than the file holds, zero-sized
    // records: placeableWmf never throws, and gives a header only within int16 and 1 to 65535 units an inch.
    const wmf = (records: number[]) => new Uint8Array([...u16(1), ...u16(9), ...u16(0x0300), ...u32(0), ...u16(0), ...u32(0), ...u16(0), ...records]);
    const shapes: [string, Uint8Array, any][] = [
        ['a window of the largest extent at a goal of one twip', wmf([...u32(5), ...u16(0x020C), ...u16(32767), ...u16(32767), ...u32(3), ...u16(0)]), { goalWidth: 1 }],
        ['a window at the edge of int16', wmf([...u32(5), ...u16(0x020B), ...u16(32767), ...u16(32767), ...u32(5), ...u16(0x020C), ...u16(100), ...u16(100), ...u32(3), ...u16(0)]), { goalWidth: 1440 }],
        ['a record claiming four billion words', wmf([...u32(0xFFFFFFFF), ...u16(0x0001)]), { goalWidth: 1440 }],
        ['a record of no words', wmf([...u32(0), ...u16(0x020C), ...u16(10), ...u16(10)]), { goalWidth: 1440 }],
        ['a truncated record', wmf([...u32(5), ...u16(0x020C), 1]), { goalWidth: 1440 }],
        ['a header size past the file', new Uint8Array([...u16(1), ...u16(0xFFFF), ...u16(0x0300)]), { goalWidth: 1440 }],
        ['a negative goal and a NaN size', wmf([...u32(5), ...u16(0x020C), ...u16(10), ...u16(10), ...u32(3), ...u16(0)]), { goalWidth: -5, picw: NaN }],
    ];
    for (const [label, data, size] of shapes) {
        let out: Uint8Array | undefined, error: unknown;
        try { out = placeableWmf(data, size); } catch (e) { error = e; }
        const placeable = !!out && out[0] === 0xD7 && out[1] === 0xCD;
        const inch = out && placeable ? out[14] | out[15] << 8 : 1;
        check(`metafile: placeableWmf of ${label} does not throw and keeps its ranges`, !error && !!out && inch >= 1 && inch <= 65535 && (placeable || out === data), String(error ?? inch));
    }

    // A parse of such a WMF in RTF: the document reads.
    const hex = Buffer.from(shapes[2][1]).toString('hex');
    const rtf = await OfficeParser.parseOffice(Buffer.from(`{\\rtf1{\\pict\\wmetafile8\\picwgoal1440 ${hex}}}`, 'latin1'), { extractAttachments: true, onWarning: () => { } });
    check('metafile: an RTF WMF of a hostile record is read', rtf.attachments.length === 1 && rtf.attachments[0].mimeType === 'image/wmf', JSON.stringify(rtf.attachments.map(a => a.mimeType)));

    // An RTF equation's text that looks like markup, an entity, a CDATA end, a NUL or a lone surrogate:
    // read as its characters (escaped LaTeX), never as OMML, and never throwing.
    for (const [label, text, literal] of [
        ['OMML markup', '</m:t><m:f><m:num>', '</m:t><m:f><m:num>'],
        ['an entity', '&amp;', '\\&amp;'],
        ['a CDATA end', ']]>', ']]>'],
        ['a NUL', 'a\\\'00b', 'ab'],
        ['a lone surrogate', 'c\\u-10240?d', 'cd'],
    ] as const) {
        let latex: string | undefined, error: unknown;
        try {
            const ast = await OfficeParser.parseOffice(Buffer.from(`{\\rtf1\\pard {\\mmath{\\*\\moMath{\\mr ${text}}{\\mf{\\mnum{\\mr 1}}{\\mden{\\mr 2}}}}}\\par}`, 'latin1'), { onWarning: () => { } });
            latex = ast.content[0]?.children?.find(n => n.type === 'code')?.text;
        } catch (e) { error = e; }
        check(`RTF equation: ${label} in its text is read as characters`, !error && latex === `${literal}\\frac{1}{2}`, String(error ?? latex));
    }
    // A backslash and braces in the text are escaped: the text's own command is never a command.
    const command = await OfficeParser.parseOffice(Buffer.from('{\\rtf1\\pard {\\mmath{\\*\\moMath{\\mr \\\\frac\\{1\\}\\{2\\}}}}\\par}', 'latin1'), { onWarning: () => { } });
    const commandText = command.content[0]?.children?.find(n => n.type === 'code')?.text ?? '';
    check('RTF equation: a LaTeX command in its text is escaped', commandText.startsWith('\\textbackslash') && !commandText.includes('\\frac{'), commandText);

    // RTF output of a placeable WMF 32767 units wide at one unit an inch, or 1 unit wide at 65535: its
    // \picw, \pich and goals whole numbers from 1 to 2147483647.
    const header = (right: number, bottom: number, inch: number) => {
        const words = [0xCDD7, 0x9AC6, 0, 0, 0, right, bottom, inch, 0, 0];
        return [...words.flatMap(u16), ...u16(words.reduce((a, w) => a ^ w, 0))];
    };
    for (const [label, right, bottom, inch] of [['32767 wide at one an inch', 32767, 32767, 1], ['one wide at 65535 an inch', 1, 1, 65535], ['of no units an inch', 100, 100, 0]] as const) {
        const bytes = new Uint8Array([...header(right, bottom, inch), ...wmf([...u32(3), ...u16(0)])]);
        const out = String((await OfficeGenerator.generate(astOf({ name: 'x.wmf', mimeType: 'image/wmf', bytes }), 'rtf', silent as any)).value);
        const words = /\\pict\\wmetafile8\\picw(\d+)\\pich(\d+)\\picwgoal(\d+)\\pichgoal(\d+)/.exec(out);
        check(`metafile: RTF sizes of a WMF ${label} are whole numbers in range`, !!words && words.slice(1).every(v => /^\d+$/.test(v) && Number(v) >= 1 && Number(v) <= 2147483647), words?.[0] ?? out.slice(0, 200));
    }

    // An equation drawn as its picture: its LaTeX, which holds markup, is the picture's alt text, escaped in each
    // package (read back, exactly the text); RTF writes none of it. Without the picture, the LaTeX path escapes it.
    const hostile = '"><img src=x onerror=1>]]></w:t>\\par}{';
    const TINY_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
    const equation = (pictured: boolean): OfficeParserAST => ({
        type: 'rtf', metadata: {},
        content: [{ type: 'paragraph', text: hostile, children: [{ type: 'code', text: hostile, metadata: { math: 'inline', fallbackAttachmentName: 'p.png' } }] }],
        attachments: pictured ? [{ type: 'image', name: 'p.png', mimeType: 'image/png', extension: 'png', data: TINY_PNG }] : [],
    } as any);
    const read = async (format: string, ast: OfficeParserAST) => {
        const files = unzipSync((await OfficeGenerator.generate(ast, format as any, silent as any)).value as Uint8Array);
        return Object.fromEntries(Object.entries(files).map(([k, v]) => [k, new TextDecoder().decode(v)]));
    };
    const values = (xml: string, tag: string, attr?: string) => Array.from(parseXmlString(xml).getElementsByTagName(tag) as any as ArrayLike<any>).map(e => attr ? e.getAttribute(attr) : e.textContent);
    for (const [format, part, tag, attr] of [['docx', 'word/document.xml', 'wp:docPr', 'descr'], ['odt', 'content.xml', 'svg:title', undefined], ['pptx', 'ppt/slides/slide1.xml', 'p:cNvPr', 'descr'], ['odp', 'content.xml', 'svg:desc', undefined]] as const) {
        const files = await read(format, equation(true));
        const found = Object.keys(files).filter(k => k === part || (format === 'pptx' && /^ppt\/slides\/slide\d+\.xml$/.test(k))).flatMap(k => values(files[k], tag, attr));
        check(`equation picture: ${format} alt text of markup is escaped, read back exactly`, found.includes(hostile) && !Object.values(files).some(x => x.includes('<img src=x')), JSON.stringify(found));
    }
    const printed = String((await new HtmlGenerator(equation(true), { htmlConfig: { standalone: true } } as any, { drawsMathPictures: true }).generate()).value);
    const alt = /<img\b[^>]*\balt="([^"]*)"/.exec(printed)?.[1];
    const unescaped = alt?.replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
    check('equation picture: the printed HTML\'s alt text of markup is escaped', unescaped === hostile && !printed.includes('<img src=x'), alt ?? printed.slice(0, 200));
    const rtfOut = String((await OfficeGenerator.generate(equation(true), 'rtf', silent as any)).value);
    check('equation picture: RTF writes the picture and none of its LaTeX', rtfOut.includes('\\pict\\pngblip') && !rtfOut.includes('onerror') && !rtfOut.includes('w:t'), rtfOut.slice(0, 300));
    // Without the picture: the LaTeX as text, escaped (read back from DOCX and ODT exactly).
    for (const [format, part, tag] of [['docx', 'word/document.xml', 'w:t'], ['odt', 'content.xml', 'text:span']] as const) {
        const files = await read(format, equation(false));
        check(`equation picture: ${format} writes the LaTeX of markup as text when the picture is missing`, values(files[part], tag).join('').includes(hostile), values(files[part], tag).join('|'));
    }
}

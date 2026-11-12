/**
 * Hostile chart shapes through the writers: chart text a small chart part makes a writer write again and again (an
 * axis title over each series' columns, a chart framed many times), held to the repeated content budget.
 */
import * as fs from 'fs';
import * as path from 'path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParser } from '../../src/OfficeParser';
import type { OfficeParserAST } from '../../src/types';
import { repeatedContentOfAst } from '../../src/utils/repeatUtils';
import { withinTime } from '../helpers/timing';
import { parseXmlString } from '../../src/utils/xmlUtils';

type Check = (name: string, cond: boolean, detail?: string) => void;

const FILES = path.join(__dirname, '..', 'files');
const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const title = (text: string) => `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></c:rich></c:tx></c:title>`;
const numRef = (values: string[]) => `<c:numRef><c:f>S!$B$2</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef>`;
const chartPart = (plotArea: string, chartTitle = '') => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace ${C}><c:chart>${chartTitle ? title(chartTitle) : ''}<c:plotArea>${plotArea}</c:plotArea></c:chart></c:chartSpace>`;

/** What `format` writes of `ast`, in characters (a zip's parts unzipped), and the codes of its warnings; a throw is its message. */
const writtenSize = async (ast: OfficeParserAST, format: string): Promise<{ size: number; codes: string[]; error?: string }> => {
    const codes: string[] = [];
    try {
        const out = await ast.to(format as any, { htmlConfig: { chartJsSrc: '' }, pdfConfig: { engine: 'native' }, onWarning: (w: any) => codes.push(w.code) } as any);
        const value = out.value as any;
        const size = typeof value === 'string' ? value.length : format === 'pdf' ? value.length : Object.values(unzipSync(new Uint8Array(value))).reduce((sum, part) => sum + part.length, 0);
        return { size, codes };
    } catch (error: any) {
        return { size: Infinity, codes, error: String(error?.message ?? error).slice(0, 120) };
    }
};

export async function chartWritersHardeningTests(check: Check): Promise<void> {
    console.log('- Chart text written again held to the repeated content budget...');
    const silent = { extractAttachments: true, onWarning: () => { } };

    // The review's shape: 3.3 KB of chart XML (a scatter chart of 499 one-point series sharing two axes, each axis
    // titled 400,000 characters) in test.xlsx made 400 MB of HTML, Markdown and RTF: each axis title over every
    // series' column. Now each title is whole once and its other copies within the repeat budget.
    const T = 400_000;
    const sers = Array.from({ length: 499 }, (_, i) => `<c:ser><c:idx val="${i}"/><c:xVal>${numRef(['1'])}</c:xVal><c:yVal>${numRef(['2'])}</c:yVal></c:ser>`).join('');
    const scatter = (t: number) => chartPart(`<c:scatterChart>${sers}<c:axId val="1"/><c:axId val="2"/></c:scatterChart><c:valAx><c:axId val="1"/><c:axPos val="b"/>${title('A'.repeat(t))}<c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:axPos val="l"/>${title('B'.repeat(t))}<c:crossAx val="1"/></c:valAx>`, 'T');
    const xlsx = (t: number) => Buffer.from(zipSync({ ...unzipSync(new Uint8Array(fs.readFileSync(path.join(FILES, 'test.xlsx')))), 'xl/charts/chart1.xml': strToU8(scatter(t)) }));
    const amplified = await OfficeParser.parseOffice(xlsx(T), silent as any) as OfficeParserAST;
    const plain = await OfficeParser.parseOffice(xlsx(1), silent as any) as OfficeParserAST;
    for (const format of ['md', 'rtf', 'html', 'docx']) {
        const base = await writtenSize(plain, format);
        const { size, codes, error } = await writtenSize(amplified, format);
        const bound = base.size + 2 * T + repeatedContentOfAst(amplified) + 499 * 2 * 3 * 70;
        check(`charts: axis titles over 499 series (${format}) are written within the repeat budget`, !error && size <= bound, error ?? `${size} > ${bound}`);
        check(`charts: axis titles over 499 series (${format}) are reported past the budget`, codes.includes('REPEATED_CONTENT_LIMIT_EXCEEDED'), codes.join(','));
    }

    // Shapes of one 20,000-value chart shown 1,000 times: a paragraph of word-charts.docx repeated (65 KB;
    // 348 MB of RTF), a frame of test.pptx repeated (143 KB; RTF past what a string holds, DOCX too before the budget). Each
    // copy of the table is charged what it writes, past the budget a reference: every writer within what one chart
    // node writes, plus the budget, plus a few KB a node (a slide or sheet of its own in a deck or workbook).
    const values = Array.from({ length: 20_000 }, () => '1234567890');
    const bar = chartPart(`<c:barChart><c:ser><c:idx val="0"/><c:tx><c:strRef><c:f>S!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:strRef></c:tx><c:val>${numRef(values)}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>`);
    const docx = (n: number) => {
        const base = unzipSync(new Uint8Array(fs.readFileSync(path.join(FILES, 'charts', 'word-charts.docx'))));
        const doc = strFromU8(base['word/document.xml']);
        const para = doc.match(/<w:p><w:pPr>(?:(?!<w:p>).)*?r:id="rId2"\/>.*?<\/w:p>/s)![0];
        return Buffer.from(zipSync({ ...base, 'word/charts/chart1.xml': strToU8(bar), 'word/document.xml': strToU8(doc.replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${para.repeat(n)}</w:body>`)) }, { level: 9 }));
    };
    const pptx = (n: number) => {
        const base = unzipSync(new Uint8Array(fs.readFileSync(path.join(FILES, 'test.pptx'))));
        const slide = strFromU8(base['ppt/slides/slide9.xml']);
        const frame = slide.match(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/)![0];
        return Buffer.from(zipSync({ ...base, 'ppt/charts/chart1.xml': strToU8(bar), 'ppt/slides/slide9.xml': strToU8(slide.replace(frame, frame.repeat(n))) }, { level: 9 }));
    };
    // (DOCX, ODT and LaTeX write a chart standing in a line, as Word places one, as its data's lines of text, not a
    // table, each copy charged what they write: 1,000 paragraphs of a 20,000-value chart made 358 MB of DOCX XML.)
    const writers: [string, (n: number) => Buffer, string[]][] = [
        ['DOCX paragraphs', docx, ['md', 'rtf', 'html', 'docx', 'odt', 'tex', 'pptx', 'odp', 'xlsx', 'ods', 'epub']],
        ['PPTX frames', pptx, ['md', 'rtf', 'html', 'docx', 'odt', 'tex', 'pptx', 'odp', 'xlsx', 'ods', 'epub']],
    ];
    for (const [shape, file, formats] of writers) {
        const one = await OfficeParser.parseOffice(file(1), silent as any) as OfficeParserAST;
        const many = await OfficeParser.parseOffice(file(1000), silent as any) as OfficeParserAST;
        for (const format of formats) {
            const first = await writtenSize(one, format);
            const { size, error } = await writtenSize(many, format);
            const bound = first.size + repeatedContentOfAst(many) + 3000 * 1000;
            check(`charts: ${shape} of one chart shown 1,000 times (${format}) within the repeat budget`, !error && size <= bound, error ?? `${size} > ${bound}`);
        }
    }

    // Native PDF: a chart framed 200 times whose title and series name are 400,000 characters each (the review's
    // 15 seconds at 200,000) is drawn in the time and size of short ones: each string cut to its line first.
    const long = 'S'.repeat(400_000);
    const framed = (name: string, chartTitle: string) => {
        const part = chartPart(`<c:barChart><c:ser><c:idx val="0"/><c:tx><c:strRef><c:f>S!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>S!$A$2</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val>${numRef(['1', '2'])}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>`, chartTitle);
        return pptxFramed(part, 200);
    };
    const pptxFramed = (part: string, n: number) => {
        const base = unzipSync(new Uint8Array(fs.readFileSync(path.join(FILES, 'test.pptx'))));
        const slide = strFromU8(base['ppt/slides/slide9.xml']);
        const frame = slide.match(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/)![0];
        return Buffer.from(zipSync({ ...base, 'ppt/charts/chart1.xml': strToU8(part), 'ppt/slides/slide9.xml': strToU8(slide.replace(frame, frame.repeat(n))) }, { level: 9 }));
    };
    const short = await writtenSize(await OfficeParser.parseOffice(framed('S', 'T'), silent as any) as OfficeParserAST, 'pdf');
    const longAst = await OfficeParser.parseOffice(framed(long, 'T'.repeat(400_000)), silent as any) as OfficeParserAST;
    const { result: drawn, ms, within } = await withinTime(20_000, () => writtenSize(longAst, 'pdf'));
    check('charts: native PDF of a chart framed 200 times with 400,000-character titles is the size of short ones', !drawn.error && drawn.size <= 1.5 * short.size, drawn.error ?? `${drawn.size} against ${short.size}`);
    check('charts: and drawn in seconds', within, `${ms} ms`);

    // Native PDF: a paragraph of word-charts.docx holding a chart of 20,000 one-character values (one it cannot draw,
    // written as its lines), framed 1,000 times: 200 frames took 96 seconds. Each copy of its text is charged what it
    // lays out, past the budget its reference: the default budget's worth of lines, in seconds.
    const ones = chartPart(`<c:barChart><c:ser><c:idx val="0"/><c:tx><c:strRef><c:f>S!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:strRef></c:tx><c:val>${numRef(Array.from({ length: 20_000 }, () => '1'))}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>`);
    const inLines = await OfficeParser.parseOffice((() => {
        const base = unzipSync(new Uint8Array(fs.readFileSync(path.join(FILES, 'charts', 'word-charts.docx'))));
        const doc = strFromU8(base['word/document.xml']);
        const para = doc.match(/<w:p><w:pPr>(?:(?!<w:p>).)*?r:id="rId2"\/>.*?<\/w:p>/s)![0];
        return Buffer.from(zipSync({ ...base, 'word/charts/chart1.xml': strToU8(ones), 'word/document.xml': strToU8(doc.replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${para.repeat(1000)}</w:body>`)) }, { level: 9 }));
    })(), silent as any) as OfficeParserAST;
    const { result: lined, ms: linedMs, within: linedWithin } = await withinTime(40_000, () => writtenSize(inLines, 'pdf'));
    check('charts: native PDF of a 20,000-value chart in 1,000 paragraphs is written within the repeat budget', !lined.error && lined.codes.includes('REPEATED_CONTENT_LIMIT_EXCEEDED'), lined.error ?? lined.codes.join(','));
    check('charts: and in under 40 seconds', linedWithin, `${linedMs} ms`);

    // The round 2 review's shapes of text a writer escapes to many bytes: an 89 KB test.pptx whose chart's 50 labels are
    // 2,000 `^` each, framed 200 times, ran LaTeX (`\textasciicircum{}` a character) out of a 2 GB heap; 1,000 emoji
    // each, framed 300 times, made 204 MB of RTF (`\uN` a code unit). The round 3 review's: 700 `'a  '` each, framed
    // 400 times, made 117 MB of ODP (a run of two spaces is a space and a `<text:s text:c="1"/>`, 22 bytes for 3
    // characters, which ODP's charge left out). Each copy is charged what the writer writes of it.
    const escaped = (unit: string, length: number) => {
        const label = unit.repeat(length);
        return chartPart(`<c:barChart><c:ser><c:idx val="0"/><c:cat><c:strRef><c:f>S!$A$2</c:f><c:strCache><c:ptCount val="50"/>${Array.from({ length: 50 }, (_, i) => `<c:pt idx="${i}"><c:v>${label}</c:v></c:pt>`).join('')}</c:strCache></c:strRef></c:cat><c:val>${numRef(Array.from({ length: 50 }, () => '1'))}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>`);
    };
    const unitNames: Record<string, string> = { '^': '^', '\u{1F600}': 'emoji', 'a  ': "'a  '" };
    for (const [unit, length, frames, format] of [['^', 2000, 200, 'tex'], ['\u{1F600}', 1000, 300, 'rtf'], ['a  ', 700, 400, 'odp']] as const) {
        const first = await writtenSize(await OfficeParser.parseOffice(pptxFramed(escaped(unit, length), 1), silent as any) as OfficeParserAST, format);
        const framedAst = await OfficeParser.parseOffice(pptxFramed(escaped(unit, length), frames), silent as any) as OfficeParserAST;
        const { size, error } = await writtenSize(framedAst, format);
        const bound = 1.1 * (first.size + repeatedContentOfAst(framedAst)) + frames * 3000;
        check(`charts: 50 labels of ${length} ${unitNames[unit]} framed ${frames} times (${format}) within the repeat budget as written`, !error && size <= bound, error ?? `${size} > ${bound}`);
    }

    // EPUB writes a chart's title, labels, series names and values into its chapter as a table's text: markup, an XML
    // section's end, an entity and a control character in them are text, and the chapter is XML a reading system opens.
    console.log('- Chart text in an EPUB chapter...');
    const hostile = ['</td><script>alert(1)</script>', ']]>', '&amp;x', `a${String.fromCharCode(1)}b`];
    const hostileChart = {
        type: 'docx', metadata: {},
        content: [{ type: 'chart', text: '', metadata: { attachmentName: 'c' } }],
        attachments: [{ type: 'chart', name: 'c', mimeType: 'application/xml', extension: 'xml', data: '', chartData: { title: hostile[0], labels: hostile, dataSets: [{ name: hostile[1], values: ['1', '2', '3', '4'] }], rawTexts: [] } }],
    } as any as OfficeParserAST;
    const book = (await OfficeGenerator.generate(hostileChart, 'epub', { onWarning: () => { } } as any)).value as Uint8Array;
    const chapter = strFromU8(unzipSync(book)['OEBPS/chapter1.xhtml']);
    let wellFormed = true;
    try { parseXmlString(chapter); } catch { wellFormed = false; }
    check('charts: EPUB chapter of a chart of hostile text is well-formed XML', wellFormed);
    check('charts: EPUB chapter of a chart of hostile text has no script', !/<script/i.test(chapter), chapter.match(/<script[^>]*>/i)?.[0] ?? '');
    const readBack = String((await (await OfficeParser.parseOffice(Buffer.from(book), { fileType: 'epub', onWarning: () => { } } as any)).to('text')).value);
    const shown = ['</td><script>alert(1)</script>', ']]>', '&amp;x', 'ab'];
    check('charts: EPUB reads back a chart\'s hostile text as text (the control character left out)', shown.every(text => readBack.includes(text)), JSON.stringify(readBack.slice(0, 200)));

    // A chart without its data is written as its text (HTML and EPUB wrote nothing for it): markup in it is text.
    console.log('- Chart without its data written as its text...');
    const markup = '<img src=x onerror=alert(1)>\n</p><script>alert(2)</script>';
    const dataless = { type: 'docx', metadata: {}, attachments: [], content: [{ type: 'chart', text: markup, metadata: { attachmentName: 'c' } }] } as any as OfficeParserAST;
    for (const format of ['html', 'epub'] as const) {
        const value = (await OfficeGenerator.generate(dataless, format, { onWarning: () => { } } as any)).value;
        const page = format === 'html' ? String(value) : strFromU8(unzipSync(value as unknown as Uint8Array)['OEBPS/chapter1.xhtml']);
        check(`charts: ${format} writes a chart without its data's markup as text (no <img>, no <script>)`, !/<img/i.test(page) && !/<script/i.test(page) && page.includes('&lt;img src=x onerror=alert(1)&gt;'), page.slice(0, 300));
        if (format === 'epub') {
            let ok = true;
            try { parseXmlString(page); } catch { ok = false; }
            check('charts: EPUB chapter of a chart without its data of markup is well-formed XML', ok);
        }
    }
}

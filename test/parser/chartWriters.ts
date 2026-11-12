/**
 * Charts as the writers write them: chart text a writer writes again (a title at every chart showing it, an axis
 * title over each series' columns, a table at every chart showing it) charged to the repeated content budget.
 */
import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import * as vm from 'vm';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParser } from '../../src/OfficeParser';
import type { ChartData, OfficeParserAST } from '../../src/types';
import { repeatedContentOfAst, repeatPreview } from '../../src/utils/repeatUtils';
import { parseXmlString } from '../../src/utils/xmlUtils';

// ─── Shared fixtures ─────────────────────────────────────────────────────────

const FILES = path.join(__dirname, '..', 'files');

/** The writers that write a chart as its data's table. */
export const TABLE_WRITERS = ['md', 'rtf', 'html', 'docx', 'odt', 'tex', 'pptx', 'odp', 'xlsx', 'ods'] as const;
type TableWriter = typeof TABLE_WRITERS[number];

/** How many times `part` occurs in `text`, not overlapping. */
export const occurrences = (text: string, part: string): number => text.split(part).length - 1;

/** A document of `nodes` chart nodes each showing `chartData`, built in code, of the repeat budget given (none: the default). */
export const chartAst = (chartData: unknown, nodes = 1, maxRepeatedContent?: number): OfficeParserAST => ({
    type: 'docx', metadata: {},
    ...(maxRepeatedContent !== undefined ? { config: { decompressionLimits: { maxRepeatedContent } } } : {}),
    content: Array.from({ length: nodes }, () => ({ type: 'chart', text: '', metadata: { attachmentName: 'c' } })),
    attachments: [{ type: 'chart', name: 'c', mimeType: 'application/xml', extension: 'xml', data: '', chartData }],
}) as any as OfficeParserAST;

/**
 * What `format` writes of `ast` as text (a zip's XML and XHTML parts joined), and the codes of its warnings. ODS sheet names
 * (`table:name`, a chart's sheet named for its title, at most 255 characters a chart) are left out.
 */
export const writtenText = async (ast: OfficeParserAST, format: string, config: Record<string, unknown> = {}): Promise<{ text: string; bytes: number; warnings: string[] }> => {
    const warnings: string[] = [];
    const out = await OfficeGenerator.generate(ast, format as any, { htmlConfig: { chartJsSrc: '' }, ...config, onWarning: (w: any) => warnings.push(w.code) } as any);
    let text: string;
    if (typeof out.value === 'string') text = out.value;
    else {
        const parts = unzipSync(new Uint8Array(out.value as any));
        text = Object.entries(parts).filter(([name]) => /\.(xml|rels|xhtml)$/.test(name)).map(([, data]) => strFromU8(data)).join('\n');
    }
    const bytes = text.length;
    if (format === 'ods') text = text.replace(/table:name="[^"]*"/g, '');
    return { text, bytes, warnings };
};

/** A chart part of `plotArea` (OpenXML), titled `chartTitle` where given. */
const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const title = (text: string) => `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></c:rich></c:tx></c:title>`;
const pts = (values: string[]) => values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('');
const numRef = (values: string[]) => `<c:numRef><c:f>S!$B$2</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${pts(values)}</c:numCache></c:numRef>`;
const strRef = (values: string[]) => `<c:strRef><c:f>S!$A$2</c:f><c:strCache><c:ptCount val="${values.length}"/>${pts(values)}</c:strCache></c:strRef>`;
const chartPart = (plotArea: string, chartTitle = '') => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace ${C}><c:chart>${chartTitle ? title(chartTitle) : ''}<c:plotArea>${plotArea}</c:plotArea></c:chart></c:chartSpace>`;

/** test.xlsx with its chart part `chart`. */
const xlsxWith = (chart: string) => Buffer.from(zipSync({ ...unzipSync(new Uint8Array(fs.readFileSync(path.join(FILES, 'test.xlsx')))), 'xl/charts/chart1.xml': strToU8(chart) }));
/** test.pptx with its chart part `chart`, and the chart's frame on slide 9 repeated `frames` times. */
const pptxWith = (chart: string, frames = 1) => {
    const base = unzipSync(new Uint8Array(fs.readFileSync(path.join(FILES, 'test.pptx'))));
    const slide = strFromU8(base['ppt/slides/slide9.xml']);
    const frame = slide.match(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/)![0];
    return Buffer.from(zipSync({ ...base, 'ppt/charts/chart1.xml': strToU8(chart), 'ppt/slides/slide9.xml': strToU8(slide.replace(frame, frame.repeat(frames))) }, { level: 9 }));
};

/** A scatter chart of `series` one-point series sharing two value axes, titled `xTitle` and `yTitle` (the review's shape). */
const sharedAxesScatter = (series: number, xTitle: string, yTitle: string) => chartPart(
    `<c:scatterChart>${Array.from({ length: series }, (_, i) => `<c:ser><c:idx val="${i}"/><c:xVal>${numRef(['1'])}</c:xVal><c:yVal>${numRef(['2'])}</c:yVal></c:ser>`).join('')}<c:axId val="1"/><c:axId val="2"/></c:scatterChart>`
    + `<c:valAx><c:axId val="1"/><c:axPos val="b"/>${title(xTitle)}<c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:axPos val="l"/>${title(yTitle)}<c:crossAx val="1"/></c:valAx>`, 'T');

/** Times `run` at n and 4n (the shape given), held to `large <= 8 * small + 250` ms, measured again once when over. */
const linearTime = async (name: string, run: (scale: 1 | 4) => Promise<unknown>): Promise<void> => {
    const time = async (scale: 1 | 4) => { const started = performance.now(); await run(scale); return performance.now() - started; };
    let small = await time(1), large = await time(4);
    if (large > 8 * small + 250) { small = await time(1); large = Math.min(large, await time(4)); }
    assert.ok(large <= 8 * small + 250, `${name}: 4n in ${large.toFixed(0)} ms, n in ${small.toFixed(0)} ms`);
};

// ─── Chart titles and axis titles written again ──────────────────────────

/**
 * A chart's title written at each chart node showing it, and an axis title over each series' columns of a scatter or
 * bubble chart's table: the first whole, the others charged to the repeated content budget, past it their start.
 */
export async function testChartTextRepeats(): Promise<void> {
    const A100 = 'A'.repeat(100), B100 = 'B'.repeat(100);
    const xy = (a: string, b: string): ChartData => ({
        labels: [], rawTexts: [],
        axes: [{ id: '1', kind: 'value', title: a }, { id: '2', kind: 'value', title: b }],
        dataSets: [0, 1, 2].map(i => ({ name: `S${i}`, values: ['1', '2'], xValues: ['3', '4'], axisIds: ['1', '2'] })),
    } as any);
    for (const format of TABLE_WRITERS) {
        // A budget of 0: each axis title whole over its first column, its start over the other two.
        const { text, warnings } = await writtenText(chartAst(xy(A100, B100), 1, 0), format);
        for (const [letter, whole] of [['A', A100], ['B', B100]] as const) {
            assert.strictEqual(occurrences(text, whole), 1, `${format}: the ${letter} axis title is written whole once`);
            if (format !== 'xlsx') assert.strictEqual(occurrences(text, letter.repeat(64)), 3, `${format}: and its start (64 characters) over the other two columns`);
        }
        assert.strictEqual(warnings.filter(w => w === 'REPEATED_CONTENT_LIMIT_EXCEEDED').length, 1, `${format}: reported once; got ${warnings}`);
        // Titles of 64 characters cost nothing: whole over every column, no warning.
        const A64 = 'A'.repeat(64), B64 = 'B'.repeat(64);
        const short = await writtenText(chartAst(xy(A64, B64), 1, 0), format);
        if (format !== 'xlsx') assert.ok(occurrences(short.text, A64) === 3 && occurrences(short.text, B64) === 3, `${format}: titles of 64 characters are whole over all six columns`);
        assert.ok(!short.warnings.includes('REPEATED_CONTENT_LIMIT_EXCEEDED'), `${format}: titles of 64 characters are not reported`);
        // The default budget: whole over all six.
        const roomy = await writtenText(chartAst(xy(A100, B100), 1), format);
        if (format !== 'xlsx') assert.ok(occurrences(roomy.text, A100) === 3 && occurrences(roomy.text, B100) === 3, `${format}: within the budget each axis title is whole over its three columns`);
        assert.ok(!roomy.warnings.includes('REPEATED_CONTENT_LIMIT_EXCEEDED'), `${format}: within the budget nothing is reported`);
    }

    // One bar chart at three chart nodes: its title whole at the first, its start at the others past the budget.
    const T100 = 'T'.repeat(100);
    const bar = { title: T100, labels: ['a', 'b'], dataSets: [{ name: 'Series', values: ['1', '2'] }], rawTexts: [] };
    for (const format of TABLE_WRITERS) {
        const { text, warnings } = await writtenText(chartAst(bar, 3, 0), format);
        assert.strictEqual(occurrences(text, T100), 1, `${format}: the title is written whole at the first chart`);
        assert.ok(occurrences(text, 'T'.repeat(64)) >= 3, `${format}: and its start at the other two`);
        assert.strictEqual(warnings.filter(w => w === 'REPEATED_CONTENT_LIMIT_EXCEEDED').length, 1, `${format}: reported once`);
        // Within the default budget: whole at all three, nothing reported (a writer that built one chart's table twice
        // would show its own first table's title cut).
        const roomy = await writtenText(chartAst(bar, 3), format);
        assert.strictEqual(occurrences(roomy.text, T100), 3, `${format}: within the budget the title is whole at all three`);
        assert.ok(!roomy.warnings.includes('REPEATED_CONTENT_LIMIT_EXCEEDED'), `${format}: within the budget nothing is reported`);
    }

    // The review's shape: a scatter chart of 499 one-point series sharing two axes, titled T characters each, from
    // test.xlsx. Each writer's output is at most what the same chart with one-character titles writes, plus each
    // title whole once, plus the repeat budget (1 MiB here, plus 16 a byte of the file), plus 70 a cell; and n to 4n
    // adds the titles' own growth alone.
    const parse = async (t: number) => await OfficeParser.parseOffice(xlsxWith(sharedAxesScatter(499, 'A'.repeat(t), 'B'.repeat(t))), { extractAttachments: true, decompressionLimits: { maxRepeatedContent: 1024 * 1024 }, onWarning: () => { } } as any) as OfficeParserAST;
    const base = await parse(1), small = await parse(20_000), large = await parse(80_000);
    const cells = 499 * 2 * 3;
    for (const format of TABLE_WRITERS) {
        const baseBytes = (await writtenText(base, format)).bytes;
        const at = async (ast: OfficeParserAST, t: number) => {
            const { bytes, warnings } = await writtenText(ast, format);
            assert.ok(bytes <= baseBytes + 2 * t + repeatedContentOfAst(ast) + cells * 70, `${format}: titles of ${t} write ${bytes} characters, the budget bounds them`);
            assert.ok(!warnings.includes('OUTPUT_TOO_LARGE'), `${format}: within what a string holds`);
            return bytes;
        };
        const n = await at(small, 20_000), n4 = await at(large, 80_000);
        assert.ok(n4 - n <= 2 * 3 * 20_000 + 1000, `${format}: titles four times as long add their own growth alone (${n} to ${n4})`);
        await linearTime(`${format}: axis titles over 499 series`, scale => writtenText(scale === 1 ? small : large, format));
    }
}

// ─── A chart's table written again ───────────────────────────────────────

/** Each table writer's class, to read what it charges (BaseGenerator.chartTextWeight) and records (chartTables). */
const generatorOf = async (format: TableWriter): Promise<any> => {
    const classes: Record<TableWriter, () => Promise<any>> = {
        md: async () => (await import('../../src/generators/MarkdownGenerator')).MarkdownGenerator,
        rtf: async () => (await import('../../src/generators/RtfGenerator')).RtfGenerator,
        html: async () => (await import('../../src/generators/HtmlGenerator')).HtmlGenerator,
        docx: async () => (await import('../../src/generators/DocxGenerator')).DocxGenerator,
        odt: async () => (await import('../../src/generators/OdtGenerator')).OdtGenerator,
        tex: async () => (await import('../../src/generators/LatexGenerator')).LatexGenerator,
        pptx: async () => (await import('../../src/generators/PptxGenerator')).PptxGenerator,
        odp: async () => (await import('../../src/generators/OdpGenerator')).OdpGenerator,
        xlsx: async () => (await import('../../src/generators/XlsxGenerator')).XlsxGenerator,
        ods: async () => (await import('../../src/generators/OdsGenerator')).OdsGenerator,
    };
    return await classes[format]();
};

/** The configuration the tests write each format with (HTML: a chart as its table, no Chart.js). */
const configOf = (format: string, config: Record<string, unknown> = {}) => ({ ...(format === 'html' ? { htmlConfig: { chartJsSrc: '' } } : {}), ...config, onWarning: () => { } });

/**
 * What a copy of `data`'s table weighs as `format` charges it (the chart's record, BaseGenerator.chartTables): a deck's
 * or workbook's estimate, a flow writer's first table as written. Read from the writer itself, writing two chart nodes.
 */
const recordedWeight = async (format: TableWriter, data: unknown, config: Record<string, unknown> = {}): Promise<number> => {
    const Generator = await generatorOf(format);
    const generator = new Generator(chartAst(data, 2, 1024 * 1024 * 1024), configOf(format, config));
    await generator.generate();
    const record = generator.chartTables.get(data);
    return (record.plain ?? record.caption).weight;
};

/** What `format` writes of `ast` in UTF-8 bytes (a zip's parts unzipped), its text (a zip's XML and XHTML parts), and the codes of its warnings. */
export const writtenBytes = async (ast: OfficeParserAST, format: string, config: Record<string, unknown> = {}): Promise<{ bytes: number; text: string; warnings: string[] }> => {
    const warnings: string[] = [];
    const out = await OfficeGenerator.generate(ast, format as any, { ...configOf(format, config), onWarning: (w: any) => warnings.push(w.code) } as any);
    if (typeof out.value === 'string') return { bytes: Buffer.byteLength(out.value, 'utf8'), text: out.value, warnings };
    const parts = unzipSync(new Uint8Array(out.value as any));
    const xml = Object.entries(parts).filter(([name]) => /\.(xml|rels|xhtml)$/.test(name));
    return { bytes: Object.values(parts).reduce((sum, part) => sum + part.length, 0), text: xml.map(([, data]) => strFromU8(data)).join('\n'), warnings };
};

/**
 * What a chart node past the budget writes, at most: its caption and its reference (a title cut to 65 characters in
 * each), a paragraph each, in a flow writer; a slide or a sheet of its own in a deck or workbook, as every chart's
 * table gets one (an ODS sheet named for its chart's title, at most 255 characters).
 */
const PER_NODE: Record<TableWriter, number> = { md: 600, rtf: 600, html: 600, docx: 600, odt: 600, tex: 600, pptx: 2500, odp: 800, xlsx: 1200, ods: 1200 };

/** word-charts.docx of one paragraph holding its chart (its part `chart`), the paragraph `n` times (the review's 65 KB shape). */
const docxFramed = (chart: string, n: number) => {
    const base = unzipSync(new Uint8Array(fs.readFileSync(path.join(FILES, 'charts', 'word-charts.docx'))));
    const doc = strFromU8(base['word/document.xml']);
    const para = doc.match(/<w:p><w:pPr>(?:(?!<w:p>).)*?r:id="rId2"\/>.*?<\/w:p>/s)![0];
    return Buffer.from(zipSync({ ...base, 'word/charts/chart1.xml': strToU8(chart), 'word/document.xml': strToU8(doc.replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${para.repeat(n)}</w:body>`)) }, { level: 9 }));
};
const barOf = (values: string[], categories?: string[], seriesName = 'A', chartTitle = '') => chartPart(`<c:barChart><c:ser><c:idx val="0"/><c:tx>${strRef([seriesName])}</c:tx>${categories ? `<c:cat>${strRef(categories)}</c:cat>` : ''}<c:val>${numRef(values)}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>`, chartTitle);

/** A full-period generator of 32-bit values (Math.imul), for property tests. */
export const randomOf = (seed: number) => {
    let state = seed >>> 0;
    return () => { state = (Math.imul(state, 1664525) + 1013904223) >>> 0; return state / 2 ** 32; };
};

/** The classes of text a chart's strings are drawn from: plain, escaped by some writer, past ASCII, line and space breaks. */
export const TEXT_CLASSES: Record<string, string> = {
    plain: 'a', caret: '^', ampersand: '&', quote: '"', less: '<', backslash: '\\', cjk: '中', emoji: '\u{1F600}',
    combining: 'é', newline: 'a\n', tab: 'a\t', spaces: '  ', alternating: 'a  ', spacesAndBreaks: ' a\t  \n',
};

/**
 * Every class of text a writer's charge is checked against (see testChartTextWeights): each character the writers
 * escape, strip or write past ASCII; pure breaks, one and two spaces; the alternating and leading or trailing mixes of
 * letters, spaces, tabs and breaks; Markdown's and LaTeX's markers beside a space; and TEXT_CLASSES.
 */
export const WEIGHT_CLASSES: Record<string, string> = (() => {
    const classes: Record<string, string> = {
        a: 'a', caret: '^', bs: '\\', lt: '<', gt: '>', amp: '&', quot: '"', apos: "'", lbrace: '{', rbrace: '}', dollar: '$', pct: '%', hash: '#', us: '_', tilde: '~', pipe: '|',
        star: '*', btick: '`', lbr: '[', rbr: ']', bang: '!', dash: '-', plus: '+', dot: '.', colon: ':', eq: '=',
        cjk: '\u4e2d', emoji: '\u{1F600}', comb: 'e\u0301', rtl: '\u05e9', c01: 'a\u0001', c07: 'a\u0007', c1f: 'a\u001f', c7f: 'a\u007f', zwsp: '\u200b', nbsp: '\u00a0',
        bom: '\ufeff', ls: '\u2028', nel: '\u0085', lone: '\ud800', fffe: '\ufffe', xesc: '_x0041_', ideo: '\u3000',
        nl: '\n', tab: '\t', cr: '\r', crlf: '\r\n', sp: ' ', sp2: '  ',
        anl: 'a\n', atab: 'a\t', acr: 'a\r', acrlf: 'a\r\n', asp: 'a ', asp2: 'a  ', asp3: 'a   ', spa: ' a', sp2a: '  a', sp3a: '   a',
        asptab: 'a \t', atabsp: 'a\t ', aspnl: 'a \n', anlsp: 'a\n ', asp2nl: 'a  \n', atabsp2: 'a\t  ', sptab: ' \t', tabsp: '\t ', spnl: ' \n', nlsp: '\n ',
        sp2nl: '  \n', nlsp2: '\n  ', spcrlf: ' \r\n', tabnl: '\t\n', nltab: '\n\t', tabtab: 'a\t\t', nlnl: 'a\n\n',
        ampsp2: '&  ', lttab: '<\t', cjksp2: '\u4e2d  ', emojisp: '\u{1F600} ', caretsp: '^ ', bssp2: '\\  ', quotnl: '"\n', nbspsp: 'a\u00a0 ', spideo: ' \u3000',
        pipesp: '| ', starsp: '* ', hashsp: '# ', dashsp: '- ', usp: '_ ', btsp: '` ', bangbr: '![', gtsp: '> ', dotsp: '1. ',
    };
    const units = new Set(Object.values(classes));
    for (const [name, unit] of Object.entries(TEXT_CLASSES)) if (!units.has(unit)) classes[name] = unit;
    return classes;
})();

/** The writers that decide a chart's copy at layout and write it later, charging an estimate (see BaseGenerator.chartTable). */
const ESTIMATING_WRITERS = ['pptx', 'odp', 'xlsx', 'ods'] as const;

/**
 * A chart's table at a chart node after the first showing the same data: the first table again while the repeat
 * budget lasts, each copy charged what it writes (a flow writer measures it, a deck or workbook writer estimates it);
 * past it a reference to the chart, `[Chart: <title>]`.
 */
export async function testChartTableCopies(): Promise<void> {
    // A deck's or workbook's copy writes no more than the weight it is charged: each class of text (100 labels of 400
    // characters, the title and series name the same), and two scatter charts (300 one-point series, their headers
    // spanning columns; 100 series of names, titles and axis titles of many lines). Third chart node against the second.
    const xyShapes: Record<string, unknown> = {
        'scatter of 300 series': { title: 'T', labels: [], xAxisTitle: 'X title', yAxisTitle: 'Y title', dataSets: Array.from({ length: 300 }, (_, i) => ({ name: `series ${i}`, values: ['1'], xValues: ['2'] })), rawTexts: [] },
        'scatter of lines': { title: 'T\n'.repeat(50), labels: [], xAxisTitle: 'X\n'.repeat(30), yAxisTitle: 'Y', dataSets: Array.from({ length: 100 }, () => ({ name: 'n\n'.repeat(20), values: ['1'], xValues: ['2'] })), rawTexts: [] },
    };
    const shapes: [string, unknown][] = [
        ...Object.entries(TEXT_CLASSES).map(([name, unit]): [string, unknown] => {
            const label = unit.repeat(Math.ceil(400 / unit.length));
            return [name, { title: label, labels: Array.from({ length: 100 }, () => label), dataSets: [{ name: label, values: Array.from({ length: 100 }, () => '1') }], rawTexts: [] }];
        }),
        ...Object.entries(xyShapes),
    ];
    for (const format of ESTIMATING_WRITERS) {
        for (const [name, data] of shapes) {
            const weight = await recordedWeight(format, data);
            const copy = (await writtenBytes(chartAst(data, 3, 1024 * 1024 * 1024), format)).bytes - (await writtenBytes(chartAst(data, 2, 1024 * 1024 * 1024), format)).bytes;
            if (process.env.SHOW_WEIGHTS) console.log(`  ${format}, ${name}: ${(copy / weight).toFixed(2)}`);
            assert.ok(copy <= weight, `${format}, ${name}: a copy writes ${copy} bytes, charged ${weight}`);
        }
    }

    // Past the budget (0 here): the first table whole, then the chart's reference, reported once.
    const sales = (chartTitle?: string) => ({ ...(chartTitle !== undefined ? { title: chartTitle } : {}), labels: Array.from({ length: 10 }, (_, i) => `L${i}`), dataSets: [{ name: 'Units', values: Array.from({ length: 10 }, (_, i) => String(1000 + i)) }], rawTexts: [] });
    const T100 = 'T'.repeat(100);
    for (const format of TABLE_WRITERS) {
        for (const [chartTitle, reference] of [['Sales', 'Chart: Sales'], [undefined, 'Chart: c'], [T100, `Chart: ${'T'.repeat(64)}`]] as const) {
            const { text, warnings } = await writtenText(chartAst(sales(chartTitle), 3, 0), format);
            assert.strictEqual(occurrences(text, 'L9'), 1, `${format}: the first table is whole, the others not written (${chartTitle})`);
            assert.strictEqual(occurrences(text, reference), 2, `${format}: the two copies past the budget are the chart's reference "${reference}"`);
            assert.strictEqual(warnings.filter(w => w === 'REPEATED_CONTENT_LIMIT_EXCEEDED').length, 1, `${format}: reported once`);
            if (chartTitle === T100) assert.strictEqual(occurrences(text, 'T'.repeat(65)), 1, `${format}: a long title is whole once, cut in the copies' captions and references`);
        }
    }
    // Read back: one table, and the two references as text.
    for (const format of ['docx', 'odt', 'md', 'rtf', 'html'] as const) {
        const out = await OfficeGenerator.generate(chartAst(sales('Sales'), 3, 0), format, { htmlConfig: { chartJsSrc: '' }, onWarning: () => { } } as any);
        const back = await OfficeParser.parseOffice(Buffer.from(out.value as any), { fileType: format, onWarning: () => { } } as any) as OfficeParserAST;
        const nodes: any[] = [];
        const walk = (list: any[] | undefined) => { for (const node of list ?? []) { nodes.push(node); walk(node.children); } };
        walk(back.content);
        assert.strictEqual(nodes.filter(n => n.type === 'table').length, 1, `${format}: reads back as one table`);
        assert.strictEqual(occurrences(String((await back.to('text')).value), '[Chart: Sales]'), 2, `${format}: and the two references as text`);
    }

    // Within the budget nothing changes: a 50-value chart at 20 chart nodes is 20 whole tables in every writer.
    const fifty = { title: 'Fifty', labels: Array.from({ length: 50 }, (_, i) => `L${i}`), dataSets: [{ name: 'Units', values: Array.from({ length: 50 }, (_, i) => String(5000 + i)) }], rawTexts: [] };
    for (const format of TABLE_WRITERS) {
        const { text, warnings } = await writtenText(chartAst(fifty, 20), format);
        // (ODS writes a number twice, as its value and as the text it shows.)
        for (const value of ['5000', '5049', 'L49']) assert.strictEqual(occurrences(text, value), format === 'ods' && value !== 'L49' ? 40 : 20, `${format}: ${value} is in all 20 tables`);
        assert.ok(!warnings.includes('REPEATED_CONTENT_LIMIT_EXCEEDED'), `${format}: within the budget nothing is reported`);
    }
    // test.pptx's chart framed 20 times: 20 whole tables.
    const pptx20 = (() => {
        const base = unzipSync(new Uint8Array(fs.readFileSync(path.join(FILES, 'test.pptx'))));
        const slide = strFromU8(base['ppt/slides/slide9.xml']);
        const frame = slide.match(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/)![0];
        return Buffer.from(zipSync({ ...base, 'ppt/slides/slide9.xml': strToU8(slide.replace(frame, frame.repeat(20))) }));
    })();
    const framed = await OfficeParser.parseOffice(pptx20, { extractAttachments: true, onWarning: () => { } } as any) as OfficeParserAST;
    const framedData = framed.attachments.find(a => a.chartData)!.chartData!;
    const lastValue = framedData.dataSets[0].values[framedData.dataSets[0].values.length - 1];
    for (const format of ['md', 'rtf', 'docx', 'odt'] as const) {
        const { text, warnings } = await writtenText(framed, format);
        assert.ok(occurrences(text, framedData.dataSets[0].name!) >= 20 && occurrences(text, lastValue) >= 20, `${format}: test.pptx's chart framed 20 times is 20 whole tables`);
        assert.ok(!warnings.includes('REPEATED_CONTENT_LIMIT_EXCEEDED'), `${format}: and nothing is reported`);
    }

    // The review's shapes, at N and 4N chart nodes (here 2,000 values and a repeat budget of 256 KiB): a paragraph of
    // one chart repeated in word-charts.docx, a frame repeated in test.pptx, with categories and without, and one of a
    // long title and series name. Each output is at most what one chart node writes, plus the budget, plus a reference
    // a node; 4N adds references (and what its larger budget holds) alone. (DOCX, ODT and LaTeX write a chart standing in
    // a line, as Word places one, as its data's lines of text, not as a table, each copy charged what they write: see
    // testChartLineCopies.)
    const values = Array.from({ length: 2000 }, () => '1234567890');
    const files: [string, (n: number) => Buffer, TableWriter[]][] = [
        ['DOCX paragraphs', n => docxFramed(barOf(values), n), []],
        ['PPTX frames', n => pptxWith(barOf(values), n), []],
        ['PPTX frames with categories', n => pptxWith(barOf(values, values), n), []],
        ['PPTX frames, long title and series name', n => pptxWith(barOf(['1', '2'], ['a', 'b'], 'S'.repeat(20_000), 'T'.repeat(20_000)), n), []],
    ];
    const parse = async (file: Buffer) => await OfficeParser.parseOffice(file, { extractAttachments: true, decompressionLimits: { maxRepeatedContent: 256 * 1024 }, onWarning: () => { } } as any) as OfficeParserAST;
    for (const [shape, file, asLines] of files) {
        const one = await parse(file(1)), few = await parse(file(50)), many = await parse(file(200));
        for (const format of TABLE_WRITERS.filter(f => !asLines.includes(f))) {
            const first = (await writtenText(one, format)).bytes;
            const at = async (ast: OfficeParserAST, n: number) => {
                const { bytes, warnings } = await writtenText(ast, format);
                assert.ok(bytes <= first + repeatedContentOfAst(ast) + PER_NODE[format] * n, `${shape}, ${format}: ${n} chart nodes write ${bytes}, one writes ${first}`);
                assert.ok(!warnings.includes('OUTPUT_TOO_LARGE'), `${shape}, ${format}: within what a string holds`);
                return bytes;
            };
            const n = await at(few, 50), n4 = await at(many, 200);
            // (The larger file's budget is larger, 16 a byte, and what one budget leaves unspent is under a copy: what
            // that lets through besides.)
            const grown = repeatedContentOfAst(many) - repeatedContentOfAst(few) + first;
            assert.ok(n4 - n <= 150 * PER_NODE[format] + grown, `${shape}, ${format}: 150 more chart nodes add ${n4 - n}, their references and what the larger budget holds alone`);
            await linearTime(`${shape}, ${format}`, scale => writtenText(scale === 1 ? few : many, format));
        }
    }

    // Property: random charts at random numbers of chart nodes and budgets; the budget is spent in order, a whole
    // table a copy while it lasts, and the output stays within one node's output, the budget and a reference a node.
    const random = randomOf(0x8b1d);
    const seen = new Set<string>();
    for (let round = 0; round < 24; round++) {
        const valueCount = Math.max(1, Math.round(Math.exp(random() * Math.log(2000))));
        const seriesCount = 1 + Math.floor(random() * 8);
        const nodes = 1 + Math.floor(random() * 40);
        const budget = [0, 64 * 1024, 1024 * 1024][Math.floor(random() * 3)];
        seen.add(`${valueCount}/${seriesCount}/${nodes}/${budget}`);
        const data = {
            labels: Array.from({ length: valueCount }, (_, i) => i === valueCount - 1 ? 'Lend' : `L${i}`),
            dataSets: Array.from({ length: seriesCount }, (_, s) => ({ name: `S${s}`, values: Array.from({ length: valueCount }, (_, i) => String((i * 7 + s) % 1000)) })),
            rawTexts: [],
        };
        for (const format of TABLE_WRITERS) {
            // (Each copy weighs what the first did: the same table written again.)
            const weight = await recordedWeight(format, data);
            const expected = Math.min(nodes, 1 + Math.floor(budget / weight));
            const first = (await writtenText(chartAst(data, 1, budget), format)).bytes;
            const { text, bytes } = await writtenText(chartAst(data, nodes, budget), format);
            assert.strictEqual(occurrences(text, 'Lend'), expected, `${format}: ${valueCount} values of ${seriesCount} series at ${nodes} nodes, budget ${budget}: ${expected} whole tables`);
            assert.ok(bytes <= first + budget + PER_NODE[format] * nodes, `${format}: ${valueCount}x${seriesCount} at ${nodes} nodes, budget ${budget}: ${bytes} within one node's ${first}, the budget and a reference a node`);
        }
    }
    console.log(`  chart table copies: ${seen.size} distinct random cases`);
}

// ─── Chart copies charged what each writer writes ────────────────────────────

/** Each writer a chart's copies are bounded in, as the tests write it: the table writers, and Markdown without HTML. */
const BOUND_WRITERS: [string, TableWriter, Record<string, unknown>][] = [
    ['docx', 'docx', {}], ['odt', 'odt', {}], ['tex', 'tex', {}], ['md', 'md', {}], ['md without HTML', 'md', { mdConfig: { fallbackToHtml: false } }],
    ['rtf', 'rtf', {}], ['html', 'html', {}], ['pptx', 'pptx', {}], ['odp', 'odp', {}], ['xlsx', 'xlsx', {}], ['ods', 'ods', {}],
];

/**
 * What a chart node past the budget may add in `format`: PER_NODE, and its title's start (see repeatPreview) twice as
 * the writer writes it, in its caption and in its reference; in ODS also the sheet its table stands on, named for the
 * title (at most 255 characters, escaped).
 */
const perNodeOf = async (format: TableWriter, chartTitle: string | undefined): Promise<number> => {
    const Generator = await generatorOf(format);
    const generator = new Generator(chartAst({ labels: [], dataSets: [], rawTexts: [] }), configOf(format));
    const title = chartTitle?.trim() ?? '';
    const preview = title ? repeatPreview(title) : '';
    return PER_NODE[format] + 2 * generator.chartTextWeight(preview) + (format === 'ods' ? generator.chartTextWeight(title.slice(0, 255)) : 0);
};

/** A chart of one series of 50 points, each label `unit` repeated to about `length` characters (the review's esc.mts shape). */
const labelsChart = (unit: string, length: number) => {
    const label = unit.repeat(Math.max(1, Math.round(length / unit.length))).replace(/&/g, '&amp;').replace(/</g, '&lt;');
    return barOf(Array.from({ length: 50 }, () => '1'), Array.from({ length: 50 }, () => label));
};

/**
 * A chart's copies (its table, title and axis titles written again) together weigh at most the repeat budget as each
 * writer writes them, escapes and markup included, and each chart node past it a reference: the review's shapes (test.pptx
 * of a chart of 50 labels of 200 characters framed N times, LaTeX 18 times its charge for `^`, RTF 22 times for an
 * emoji; a title of 2,000 characters; 50 series under axis titles of 2,000) in every class of text, at N = 50 and 200.
 */
export async function testChartCopyBound(): Promise<void> {
    const budget = 100_000;
    const parse = async (file: Buffer) => await OfficeParser.parseOffice(file, { extractAttachments: true, decompressionLimits: { maxRepeatedContent: budget }, onWarning: () => { } } as any) as OfficeParserAST;
    const titled = (title: string) => ({ title, labels: ['a'], dataSets: [{ name: 'S', values: ['1'] }], rawTexts: [] });
    const axisTitled = (axisTitle: string) => ({ title: 'T', labels: [], xAxisTitle: axisTitle, yAxisTitle: axisTitle, dataSets: Array.from({ length: 50 }, () => ({ name: 's', values: ['1'], xValues: ['2'] })), rawTexts: [] });
    for (const [className, unit] of Object.entries(TEXT_CLASSES)) {
        const long = unit.repeat(Math.ceil(2000 / unit.length));
        const shapes: [string, (n: number) => Promise<OfficeParserAST>, string | undefined][] = [
            ['labels', async n => await parse(pptxWith(labelsChart(unit, 200), n)), undefined],
            ['title', async n => chartAst(titled(long), n, budget), long],
            ['axis titles', async n => chartAst(axisTitled(long), n, budget), 'T'],
        ];
        for (const [shape, astAt, chartTitle] of shapes) {
            const one = await astAt(1), few = await astAt(50), many = await astAt(200);
            for (const [name, format, config] of BOUND_WRITERS) {
                const perNode = await perNodeOf(format, chartTitle);
                const first = (await writtenBytes(one, format, config)).bytes;
                const at = async (ast: OfficeParserAST, n: number) => {
                    const { bytes, warnings } = await writtenBytes(ast, format, config);
                    assert.ok(bytes <= first + repeatedContentOfAst(ast) + (n - 1) * perNode, `${name}, ${className} ${shape}: ${n} chart nodes write ${bytes} bytes, one ${first}, the budget ${repeatedContentOfAst(ast)}`);
                    return { bytes, spent: warnings.includes('REPEATED_CONTENT_LIMIT_EXCEEDED') };
                };
                const n50 = await at(few, 50), n200 = await at(many, 200);
                // Once 50 nodes spend the budget, 150 more add their references alone. (A parsed file's budget grows with
                // it, 16 a byte: what the larger one holds, and what the smaller left unspent, under a copy, besides.)
                const larger = repeatedContentOfAst(many) - repeatedContentOfAst(few);
                const grown = larger > 0 ? larger + first : 0;
                if (n50.spent) assert.ok(n200.bytes - n50.bytes <= 150 * perNode + grown, `${name}, ${className} ${shape}: 150 more chart nodes add ${n200.bytes - n50.bytes} bytes, ${perNode} a node`);
                await linearTime(`${name}, ${className} ${shape}`, scale => writtenBytes(scale === 1 ? few : many, format, config));
            }
        }
    }

    // Budget 0, a chart at three places, each in its container (Markdown and RTF write a line holding a chart in parts):
    // the first table whole, its reference at the other two, reported once. A writer that wrote a table and threw it
    // away would make the kept one a copy, refused.
    const sales = { title: 'Sales', labels: Array.from({ length: 10 }, (_, i) => `L${i}`), dataSets: [{ name: 'Units', values: Array.from({ length: 10 }, (_, i) => String(1000 + i)) }], rawTexts: [] };
    const line = (): Run[] => [{ type: 'text', text: 'before ' }, { type: 'chart', text: '', metadata: { attachmentName: 'c' } }, { type: 'text', text: ' after' }];
    const noHtml = { mdConfig: { fallbackToHtml: false } };
    const cases: [string, Record<string, unknown>, ('md' | 'rtf')[]][] = [
        ...['paragraph', 'heading', 'list item of runs', 'list item of paragraphs', 'cell of runs', 'cell of paragraphs', 'term', 'description'].map((name): [string, Record<string, unknown>, ('md' | 'rtf')[]] => [name, {}, CONTAINER_FORMATS]),
        ['cell of runs', noHtml, ['md']], ['list item of runs', noHtml, ['md']], ['term', noHtml, ['md']],
    ];
    for (const [name, config, formats] of cases) {
        const ast = { ...chartAst(sales, 0, 0), content: [0, 1, 2].flatMap(() => CONTAINERS[name](line())) } as any as OfficeParserAST;
        for (const format of formats) {
            const label = `${name}${config === noHtml ? ', no HTML' : ''} (${format})`;
            const { text, warnings } = await writtenBytes(ast, format, config);
            assert.strictEqual(occurrences(text, 'L9'), 1, `${label}: the first table whole, no copy; got ${JSON.stringify(text)}`);
            assert.strictEqual(occurrences(text, 'Chart: Sales'), 2, `${label}: the reference at the other two`);
            assert.strictEqual(warnings.filter(w => w === 'REPEATED_CONTENT_LIMIT_EXCEEDED').length, 1, `${label}: reported once`);
            const hoisted = [...text.matchAll(/\*\*Table (\d+) \(Hoisted/g)].map(m => Number(m[1]));
            assert.deepStrictEqual(hoisted, hoisted.map((_, i) => i + 1), `${label}: hoisted tables numbered with no gap; got ${hoisted}`);
        }
    }

    // A copy that weighs more than the first did (hoisted from a list item with its heading), written and then refused,
    // is taken back with its hoisted table and number: the next table hoisted is Table 1.
    const other = { title: 'Other', labels: ['M0', 'M1'], dataSets: [{ name: 'Units', values: ['1', '2'] }], rawTexts: [] };
    const firstWeight = await recordedWeight('md', sales, noHtml);
    const attachment = (name: string, chartData: unknown) => ({ type: 'chart', name, mimeType: 'application/xml', extension: 'xml', data: '', chartData });
    const item = (text: string, name: string) => ({ type: 'list', metadata: { listType: 'unordered', indentation: 0 }, children: [{ type: 'text', text }, { type: 'chart', text: '', metadata: { attachmentName: name } }] });
    const undone = {
        type: 'docx', metadata: {}, config: { decompressionLimits: { maxRepeatedContent: firstWeight + 8 } },
        attachments: [attachment('c', sales), attachment('d', other)],
        content: [{ type: 'paragraph', children: [{ type: 'chart', text: '', metadata: { attachmentName: 'c' } }] }, item('item ', 'c'), item('next ', 'd')],
    } as any as OfficeParserAST;
    const md = (await writtenBytes(undone, 'md', noHtml)).text;
    assert.strictEqual(occurrences(md, 'L9'), 1, `a refused copy's hoisted table is taken back; got ${JSON.stringify(md)}`);
    assert.strictEqual(occurrences(md, 'Chart: Sales'), 1, 'its reference stands in its line');
    assert.ok(occurrences(md, '**Table 1 (Hoisted from a list item):**') === 1 && !md.includes('Table 2') && occurrences(md, 'See Table 1 below') === 1 && md.includes('M1'), `the next chart's table is Table 1; got ${JSON.stringify(md)}`);
}

/** What a generator's output weighs in UTF-8 bytes (a zip's parts unzipped). */
const bytesOf = (value: unknown): number => typeof value === 'string' ? Buffer.byteLength(value, 'utf8')
    : Object.values(unzipSync(new Uint8Array(value as ArrayBuffer))).reduce((sum, part) => sum + part.length, 0);

/**
 * What keeping a chart's copies makes `format` write over writing their references (its output at a budget of 1 GiB,
 * every copy kept, less its output at 0), and what the writer charged them (read off its class after the 1 GiB run).
 */
const keptAndCharged = async (format: TableWriter, config: Record<string, unknown>, data: unknown, nodes: number): Promise<{ written: number; charged: number }> => {
    const Generator = await generatorOf(format);
    const all = new Generator(chartAst(data, nodes, 1024 * 1024 * 1024), configOf(format, config));
    const kept = bytesOf((await all.generate()).value);
    const none = bytesOf((await new Generator(chartAst(data, nodes, 0), configOf(format, config)).generate()).value);
    return { written: kept - none, charged: all.chartRepeatLimit - all.chartRepeatsLeft };
};

/**
 * Property: random charts (bar or scatter, 1 to 8 series, 1 to 200 points, strings of 1 to 500 characters drawn from
 * every class of text, titled or not) at 2 to 12 chart nodes with a budget of 0, 2,000, 50,000 or 1 GiB, through each
 * writer: the output within one node's, the budget and a reference a node, and the first table whole; and what its
 * copies write over their references within what the writer charged them (and the 64 bytes a title or heading copy is
 * charged less, the place of its preview). A writer that estimates a copy (a deck or a workbook) draws a round's chart
 * of one class, three units in four, its strings' lengths uniform (not log-uniform) up to a longest uniform up to 500:
 * each such writer meets every class in turn, as a bar chart's labels and then as a scatter chart's names, in text long
 * enough that a charge short of what the writer writes of the class shows past what its cells' estimate leaves over.
 */
export async function testChartCopyBoundProperty(): Promise<void> {
    const random = randomOf(0xc0b1);
    const units = Object.values(TEXT_CLASSES);
    const upTo = (most: number) => Math.max(1, Math.round(Math.exp(random() * Math.log(most))));
    let main: string | undefined;
    const text = (most: number) => {
        let s = '';
        for (const length = main === undefined ? upTo(most) : 1 + Math.floor(random() * most); s.length < length;) s += main !== undefined && random() < 0.75 ? main : units[Math.floor(random() * units.length)];
        return s;
    };
    const seen = new Set<string>();
    for (let round = 0; round < 300; round++) {
        const [name, format, config] = BOUND_WRITERS[round % BOUND_WRITERS.length];
        // (An estimating writer's turn: each class of text in order, as a bar chart's labels, then as a scatter chart's names.)
        const turn = Math.floor(round / BOUND_WRITERS.length);
        main = (ESTIMATING_WRITERS as readonly string[]).includes(format) ? units[turn % units.length] : undefined;
        const drawnXy = random() < 0.5, xy = main === undefined ? drawnXy : Math.floor(turn / units.length) % 2 === 1, seriesCount = 1 + Math.floor(random() * 8), points = upTo(200), drawnLongest = upTo(500), hasTitle = random() < 0.5;
        const longest = main === undefined ? drawnLongest : 1 + Math.floor(random() * 500);
        const nodes = 2 + Math.floor(random() * 11), budget = [0, 2000, 50_000, 1024 * 1024 * 1024][Math.floor(random() * 4)];
        seen.add(`${name}/${JSON.stringify(main)}/${xy}/${seriesCount}/${points}/${longest}/${hasTitle}/${nodes}/${budget}`);
        // The first table's last label (its last value, for a scatter chart) a mark found as it is in every writer.
        const values = (s: number) => Array.from({ length: points }, (_, i) => xy && s === seriesCount - 1 && i === points - 1 ? '424242' : String(i + 1));
        const data = {
            ...(hasTitle ? { title: text(longest) } : {}), rawTexts: [],
            labels: xy ? [] : Array.from({ length: points }, (_, i) => i === points - 1 ? 'Lend' : text(longest)),
            ...(xy ? { xAxisTitle: text(longest), yAxisTitle: text(longest) } : {}),
            dataSets: Array.from({ length: seriesCount }, (_, s) => ({ name: text(longest), values: values(s), ...(xy ? { xValues: Array.from({ length: points }, (_, i) => String(i * 2)) } : {}) })),
        };
        const mark = xy ? '424242' : 'Lend';
        const perNode = await perNodeOf(format, (data as { title?: string }).title);
        const first = (await writtenBytes(chartAst(data, 1, budget), format, config)).bytes;
        const { bytes, text: written } = await writtenBytes(chartAst(data, nodes, budget), format, config);
        const label = `${name}, round ${round} (xy ${xy}, ${seriesCount} series of ${points}, strings to ${longest}, title ${hasTitle}, ${nodes} nodes, budget ${budget})`;
        assert.ok(bytes <= first + budget + (nodes - 1) * perNode, `${label}: ${bytes} bytes, one node ${first}, ${perNode} a node`);
        assert.ok(occurrences(written, mark) >= 1, `${label}: the first table is whole`);
        // (A copy is kept only while its charge fits the budget, so a writer charged short shows here, every copy kept.)
        const { written: kept, charged } = await keptAndCharged(format, config, data, nodes);
        assert.ok(kept <= charged + 64 * (nodes + 3 * seriesCount), `${label}: keeping the copies writes ${kept} bytes, charged ${charged}`);
    }
    console.log(`  chart copy bound: ${seen.size} distinct random cases`);
}

/**
 * Each writer charges a chart's copy at least what it writes of it, in every class of text (WEIGHT_CLASSES): a title of
 * 1,000 characters at two chart nodes, and an X axis title of 1,000 characters over 20 series' columns; for the writers
 * that estimate a copy (a deck or a workbook) also a table of 40 labels of 100 characters (the series' name and values
 * the same) at two nodes, and 30 series named 100 characters. What keeping the copies writes over writing their
 * references is at most what they were charged, and the 64 bytes each title or heading copy is charged less (the place
 * of its preview). (ODP charged a run of two spaces, 22 bytes written for `'a  '`, as its 3 characters.) And DOCX and
 * LaTeX charge a title copy what their caption writes, not a cell's form of its line breaks (up to 16 times as much).
 */
export async function testChartTextWeights(): Promise<void> {
    const of = (unit: string, length: number) => unit.repeat(Math.ceil(length / unit.length)).slice(0, length);
    const shapes: Record<string, { data: (unit: string) => unknown; nodes: number; copies: number; writers: readonly TableWriter[] }> = {
        title: { data: unit => ({ title: `X${of(unit, 1000)}Y`, labels: [], dataSets: [{ name: 'S', values: ['1'] }], rawTexts: [] }), nodes: 2, copies: 1, writers: TABLE_WRITERS },
        xy: {
            data: unit => ({ title: 'T', labels: [], xAxisTitle: `X${of(unit, 1000)}Y`, yAxisTitle: 'Z', dataSets: Array.from({ length: 20 }, () => ({ name: 's', values: ['1'], xValues: ['2'] })), rawTexts: [] }),
            nodes: 1, copies: 19, writers: TABLE_WRITERS,
        },
        tbl: {
            data: unit => { const label = of(unit, 100); return { title: '', labels: Array.from({ length: 40 }, (_, i) => label + i), dataSets: [{ name: label, values: Array.from({ length: 40 }, () => label) }], rawTexts: [] }; },
            nodes: 2, copies: 0, writers: ESTIMATING_WRITERS,
        },
        xyt: {
            data: unit => { const label = of(unit, 100); return { title: '', labels: [], dataSets: Array.from({ length: 30 }, () => ({ name: label, values: [label], xValues: [label] })), rawTexts: [] }; },
            nodes: 2, copies: 0, writers: ESTIMATING_WRITERS,
        },
    };
    const TITLE_BREAKS = ['nl', 'crlf', 'anl', 'acr', 'acrlf', 'asp2nl'];
    const worst = new Map<string, [number, string]>();
    for (const [shape, { data, nodes, copies, writers }] of Object.entries(shapes)) {
        for (const [name, format, config] of BOUND_WRITERS.filter(([, format]) => writers.includes(format))) {
            for (const [className, unit] of Object.entries(WEIGHT_CLASSES)) {
                const { written, charged } = await keptAndCharged(format, config, data(unit), nodes);
                assert.ok(written <= charged + 64 * copies, `${name}, ${className} ${shape}: keeping the copies writes ${written} bytes, charged ${charged} (and ${copies} copies of 64)`);
                // DOCX and LaTeX write a title in their caption, its lines as they are (DOCX) or spaces (LaTeX), not as a
                // cell's: charged that, not a multiple of it.
                if (shape === 'title' && (format === 'docx' || format === 'tex') && TITLE_BREAKS.includes(className)) {
                    assert.ok(written >= 0.8 * charged, `${name}, ${className} title: keeping the copy writes ${written} bytes, charged ${charged}: no more than its caption writes`);
                }
                const ratio = written / Math.max(1, charged + 64 * copies), key = `${name}, ${shape}`;
                if (!worst.has(key) || ratio > worst.get(key)![0]) worst.set(key, [ratio, className]);
            }
        }
    }
    if (process.env.SHOW_WEIGHTS) for (const [key, [ratio, className]] of worst) console.log(`  ${key}: ${ratio.toFixed(3)} (${className})`);
}

/** The writers that write a chart standing in a line as its data's lines, in the line (see BaseGenerator.chartLineWritten). */
const LINE_WRITERS = ['docx', 'odt', 'tex'] as const;

/**
 * A chart standing in a line, which DOCX, ODT and LaTeX write as its data's lines: each copy charged what the writer
 * writes (a paragraph of word-charts.docx holding a chart of 20,000 one-character values, framed 1,000 times, made
 * 358 MB of DOCX XML: the parse charged the text's characters, the writer wrote 30 bytes a line). Within the budget
 * every copy whole, as before; past it the chart's reference in the line.
 */
export async function testChartLineCopies(): Promise<void> {
    const budget = 1_000_000;
    const xmlEscaped = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;');
    const parse = async (file: Buffer, maxRepeatedContent?: number) => await OfficeParser.parseOffice(file, { extractAttachments: true, ...(maxRepeatedContent !== undefined ? { decompressionLimits: { maxRepeatedContent } } : {}), onWarning: () => { } } as any) as OfficeParserAST;
    for (const unit of ['1', '^', '&', '"', '中', '\u{1F600}']) {
        const chart = barOf(Array.from({ length: 20_000 }, () => xmlEscaped(unit)));
        const one = await parse(docxFramed(chart, 1), budget), few = await parse(docxFramed(chart, 50), budget), many = await parse(docxFramed(chart, 200), budget);
        for (const format of LINE_WRITERS) {
            const first = (await writtenBytes(one, format)).bytes;
            const at = async (ast: OfficeParserAST, n: number) => {
                const { bytes, warnings } = await writtenBytes(ast, format);
                assert.ok(bytes <= first + repeatedContentOfAst(ast) + n * 200, `${format}, values of ${unit}: ${n} paragraphs write ${bytes} bytes, one ${first}, the budget ${repeatedContentOfAst(ast)}`);
                assert.ok(warnings.includes('REPEATED_CONTENT_LIMIT_EXCEEDED'), `${format}, values of ${unit}: past the budget, reported`);
                return bytes;
            };
            const n50 = await at(few, 50), n200 = await at(many, 200);
            assert.ok(n200 - n50 <= 150 * 200, `${format}, values of ${unit}: 150 more paragraphs add ${n200 - n50} bytes, their references alone`);
            await linearTime(`${format}, a chart in a line, values of ${unit}`, scale => writtenBytes(scale === 1 ? few : many, format));
        }
    }

    // Within the budget (the default, three paragraphs): every copy's 20,000 lines whole, nothing reported, and the output
    // the same as with a budget no document reaches.
    const parsed = await parse(docxFramed(barOf(Array.from({ length: 20_000 }, () => '1')), 3));
    // (A fixed modified date: a package is stamped with the time it is written when the document gives none, and two
    // written either side of a second's turn differed.)
    const three = { ...parsed, metadata: { ...parsed.metadata, modified: new Date('2026-01-02T03:04:05Z') } } as OfficeParserAST;
    const unbounded = { ...three, config: { ...three.config, decompressionLimits: { ...three.config?.decompressionLimits, maxRepeatedContent: 1024 * 1024 * 1024 } } } as OfficeParserAST;
    const breaks: Record<typeof LINE_WRITERS[number], string> = { docx: '<w:br/>', odt: '<text:line-break/>', tex: '\\hfil\\break{}' };
    for (const format of LINE_WRITERS) {
        const out = await OfficeGenerator.generate(three, format, { onWarning: () => { } } as any);
        const { text, warnings } = await writtenBytes(three, format);
        assert.ok(occurrences(text, breaks[format]) >= 3 * 20_000, `${format}: three charts in a line, each of its 20,000 lines; got ${occurrences(text, breaks[format])} line breaks`);
        assert.ok(!warnings.includes('REPEATED_CONTENT_LIMIT_EXCEEDED'), `${format}: within the budget nothing is reported`);
        const same = await OfficeGenerator.generate(unbounded, format, { onWarning: () => { } } as any);
        assert.ok(Buffer.from(out.value as any).equals(Buffer.from(same.value as any)), `${format}: within the budget the output is what no budget writes`);
    }

    // Past the budget (0), a chart at three places, each in a line of its container: its lines whole in the first, its
    // reference in the line at the other two, reported once; read back, each line's words around the chart in place.
    const sales = { title: 'Sales', labels: ['L0', 'L9'], dataSets: [{ name: 'Units', values: ['1000', '1009'] }], rawTexts: [] };
    const line = (i: number): Run[] => [{ type: 'text', text: `before${i} ` }, { type: 'chart', text: 'L0\nL9\nUnits\n1000\n1009', metadata: { attachmentName: 'c' } }, { type: 'text', text: ` after${i}` }];
    for (const name of ['paragraph', 'heading', 'list item of runs', 'cell of runs', 'term', 'description'] as const) {
        const ast = { ...chartAst(sales, 0, 0), content: [0, 1, 2].flatMap(i => CONTAINERS[name](line(i))) } as any as OfficeParserAST;
        for (const format of LINE_WRITERS) {
            const { text: written, warnings } = await writtenBytes(ast, format);
            // (A LaTeX section's short title and label are made from the heading's whole text: not its line.)
            const text = format === 'tex' ? written.replace(/\\section\[\{[^\]]*\}\]/g, '\\section').replace(/\\label\{[^}]*\}/g, '') : written;
            assert.strictEqual(occurrences(text, 'L9'), 1, `${name} (${format}): the chart's lines whole in the first line, no copy; got ${JSON.stringify(text.slice(0, 4000))}`);
            assert.strictEqual(occurrences(text, 'Chart: Sales'), 2, `${name} (${format}): its reference in the line at the other two`);
            assert.strictEqual(warnings.filter(w => w === 'REPEATED_CONTENT_LIMIT_EXCEEDED').length, 1, `${name} (${format}): reported once`);
            if (format === 'tex') continue;
            const out = await OfficeGenerator.generate(ast, format, { onWarning: () => { } } as any);
            const back = String((await (await OfficeParser.parseOffice(Buffer.from(out.value as any), { fileType: format, onWarning: () => { } } as any)).to('text')).value);
            const order = ['before0', 'L9', 'after0', 'before1', 'Chart: Sales', 'after1', 'before2', 'Chart: Sales', 'after2'];
            let at = -1;
            for (const word of order) {
                const next = back.indexOf(word, at + 1);
                assert.ok(next > at, `${name} (${format}): read back, "${word}" in its place; got ${JSON.stringify(back)}`);
                at = next;
            }
        }
    }
}

// ─── Native PDF: a chart's text within its line ──────────────────────────────

/** The text of a PDF as the PDF reader reads it. */
export const pdfText = async (pdf: Uint8Array): Promise<string> => String((await (await OfficeParser.parseOffice(Buffer.from(pdf), { fileType: 'pdf', onWarning: () => { } } as any)).to('text')).value);

/**
 * A chart's title, axis titles, labels and legend drawn within their line: the native engine encoded and measured
 * each whole at every chart node, and drew a title far past the page's edge (a chart framed 200 times with a title of
 * 200,000 characters took 15 seconds and 1.9 MB).
 */
export async function testNativePdfChartLines(): Promise<void> {
    const shape = (length: number) => pptxWith(barOf(['1', '2'], ['a', 'b'], 'S'.repeat(length), `Head${'T'.repeat(length)}`), 50);
    const pdfOf = async (file: Buffer) => {
        const ast = await OfficeParser.parseOffice(file, { extractAttachments: true, onWarning: () => { } } as any) as OfficeParserAST;
        const started = performance.now();
        const pdf = (await ast.to('pdf', { pdfConfig: { engine: 'native' }, onWarning: () => { } } as any)).value as Uint8Array;
        return { pdf, ms: performance.now() - started };
    };
    const small = await pdfOf(shape(20_000)), large = await pdfOf(shape(80_000));
    assert.ok(large.pdf.length <= 1.5 * small.pdf.length, `the PDF does not grow with the title: ${small.pdf.length} bytes, then ${large.pdf.length}`);
    let ms = [small.ms, large.ms];
    if (ms[1] > 1.5 * ms[0] + 250) ms = [(await pdfOf(shape(20_000))).ms, Math.min(ms[1], (await pdfOf(shape(80_000))).ms)];
    assert.ok(ms[1] <= 1.5 * ms[0] + 250, `nor its time: ${ms[0].toFixed(0)} ms, then ${ms[1].toFixed(0)} ms`);
    const text = await pdfText(large.pdf);
    assert.ok(text.includes('HeadTTTT'), 'the title is drawn, its start on its line');
    assert.ok(!text.includes('T'.repeat(1000)), 'and no more of it than the line holds');

    // A chart of a short title, labels and legend draws them as before.
    const plain = { title: 'Quarterly', labels: ['Qone', 'Qtwo'], xAxisTitle: 'Quarter', yAxisTitle: 'Units', dataSets: [{ name: 'SeriesAlpha', values: ['11', '22'] }, { name: 'SeriesBeta', values: ['5', '9'] }], rawTexts: [] };
    const drawn = await pdfText((await OfficeGenerator.generate(chartAst(plain), 'pdf', { pdfConfig: { engine: 'native' }, onWarning: () => { } } as any)).value as Uint8Array);
    for (const word of ['Quarterly', 'Qone', 'Qtwo', 'Quarter', 'Units', 'SeriesAlpha', 'SeriesBeta']) assert.ok(drawn.includes(word), `${word} is drawn; got ${JSON.stringify(drawn)}`);
}

// ─── A chart in a heading, list item, definition or cell ─────────────

/** The writers that write a chart in a line as a table between the text before and after it. */
const CONTAINER_FORMATS: ('md' | 'rtf')[] = ['md', 'rtf'];

/** A Word chart drawing (inline) of relationship `rid`, as a run. */
const wordChart = (rid: string) => `<w:r><w:drawing><wp:inline><wp:extent cx="2000000" cy="1500000"/><wp:docPr id="9" name="C"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rid}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

/** word-charts.docx of a Heading 1, a list item and a hyperlink paragraph each holding one of its charts (the review's hl.docx). */
export const headingAndItemDocx = (): Buffer => {
    const base = unzipSync(new Uint8Array(fs.readFileSync(path.join(FILES, 'charts', 'word-charts.docx'))));
    const doc = strFromU8(base['word/document.xml']);
    const body = `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Head </w:t></w:r>${wordChart('rId2')}</w:p>`
        + `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item </w:t></w:r>${wordChart('rId3')}</w:p>`
        + `<w:p><w:hyperlink w:anchor="x"><w:r><w:t>link </w:t></w:r>${wordChart('rId4')}</w:hyperlink></w:p>`
        + `<w:p><w:r><w:t>plain</w:t></w:r></w:p>`;
    const numbering = `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const rels = strFromU8(base['word/_rels/document.xml.rels']).replace('</Relationships>', '<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>');
    return Buffer.from(zipSync({ ...base, 'word/document.xml': strToU8(doc.replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${body}</w:body>`)), 'word/numbering.xml': strToU8(numbering), 'word/_rels/document.xml.rels': strToU8(rels) }));
};

/**
 * word-charts.docx of a paragraph, a table cell, a text box (3.1 by 2.1 inches) and a footnote each holding one of
 * its charts (the review's nested.docx), the text box's chart its combo chart, the footnote's its scatter chart; the
 * text box `boxHeight` EMU high.
 */
export const nestedChartsDocx = (boxHeight = 2000000): Buffer => {
    const base = unzipSync(new Uint8Array(fs.readFileSync(path.join(FILES, 'charts', 'word-charts.docx'))));
    const doc = strFromU8(base['word/document.xml']);
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    const textbox = `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="3000000" cy="${boxHeight}"/><wp:docPr id="20" name="TB"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>In box:</w:t></w:r>${wordChart('rId4')}</w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox><w:txbxContent><w:p><w:r><w:t>In fallback box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
    const body = `<w:p><w:r><w:t>Top para with chart </w:t></w:r>${wordChart('rId2')}<w:r><w:t> after.</w:t></w:r></w:p>`
        + `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r>${wordChart('rId3')}</w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`
        + `<w:p>${textbox}</w:p>`
        + `<w:p><w:r><w:t>Note ref</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>`;
    const footnotes = `<?xml version="1.0"?><w:footnotes ${W}><w:footnote w:id="1"><w:p><w:r><w:t>Footnote chart:</w:t></w:r>${wordChart('rId1')}</w:p></w:footnote></w:footnotes>`;
    const rels = strFromU8(base['word/_rels/document.xml.rels']).replace('</Relationships>', '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>');
    const partRels = (target: string) => strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="${target}"/></Relationships>`);
    const types = strFromU8(base['[Content_Types].xml']).replace('</Types>', '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>');
    return Buffer.from(zipSync({ ...base, 'word/document.xml': strToU8(doc.replace('<w:document ', '<w:document xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ').replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${body}</w:body>`)), 'word/footnotes.xml': strToU8(footnotes), 'word/_rels/document.xml.rels': strToU8(rels), 'word/_rels/footnotes.xml.rels': partRels('charts/chart4.xml'), '[Content_Types].xml': strToU8(types) }));
};

type Run = { type: string; text?: string; formatting?: object; metadata?: object; children?: Run[] };

/** The containers a chart can stand in, each given the runs (text and charts) of its line. */
const CONTAINERS: Record<string, (line: Run[]) => Run[]> = {
    heading: line => [{ type: 'heading', metadata: { level: 1 }, children: line }],
    'list item of runs': line => [{ type: 'list', metadata: { listType: 'unordered', indentation: 0 }, children: line }, { type: 'list', metadata: { listType: 'unordered', indentation: 1 }, children: [{ type: 'text', text: 'nestedword' }] }],
    'list item of paragraphs': line => [{ type: 'list', metadata: { listType: 'ordered', indentation: 0 }, children: [{ type: 'paragraph', children: line }] }, { type: 'list', metadata: { listType: 'ordered', indentation: 1 }, children: [{ type: 'paragraph', children: [{ type: 'text', text: 'nestedword' }] }] }],
    quote: line => [{ type: 'paragraph', metadata: { style: 'Quote' }, children: line }],
    'cell of paragraphs': line => [{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [{ type: 'paragraph', children: line }] }, { type: 'cell', children: [{ type: 'paragraph', children: [{ type: 'text', text: 'otherword' }] }] }] }] }],
    'cell of runs': line => [{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: line }, { type: 'cell', children: [{ type: 'text', text: 'otherword' }] }] }] }],
    note: line => [{ type: 'paragraph', children: [{ type: 'text', text: 'refword' }, { type: 'note', metadata: { noteType: 'footnote', noteId: '1' }, children: [{ type: 'paragraph', children: line }] } as Run] }],
    term: line => [{ type: 'definitionList', children: [{ type: 'definitionTerm', children: line }, { type: 'definitionDescription', children: [{ type: 'text', text: 'otherword' }] }] }],
    description: line => [{ type: 'definitionList', children: [{ type: 'definitionTerm', children: [{ type: 'text', text: 'otherword' }] }, { type: 'definitionDescription', children: line }] }],
    paragraph: line => [{ type: 'paragraph', children: line }],
};

/** The read-back's nodes in document order (notes after their node), with whether each stands in a table or heading. */
const readBackNodes = (ast: OfficeParserAST) => {
    const out: { node: any; inTable: boolean; inHeading: boolean }[] = [];
    const walk = (list: any[] | undefined, inTable: boolean, inHeading: boolean) => {
        for (const node of list ?? []) {
            out.push({ node, inTable, inHeading });
            walk(node.children, inTable || node.type === 'table', inHeading || node.type === 'heading');
            walk(node.notes, inTable, inHeading);
        }
    };
    walk(ast.content, false, false);
    return out;
};

/** A table's rows as the text of each cell, trimmed, its empty cells left out. */
const rowsOf = (table: any): string[][] => {
    const textOf = (node: any): string => node.children?.length ? node.children.map(textOf).join('') : String(node.text ?? '');
    return (table.children ?? []).filter((row: any) => row.type === 'row').map((row: any) => (row.children ?? []).map((cell: any) => textOf(cell).trim()).filter(Boolean));
};

/** Whether `rtf`'s groups are balanced (braces not escaped). */
const bracesBalanced = (rtf: string): boolean => {
    let depth = 0;
    for (let i = 0; i < rtf.length; i++) {
        const c = rtf[i];
        if (c === '\\') { i++; continue; }
        if (c === '{') depth++;
        else if (c === '}' && --depth < 0) return false;
    }
    return depth === 0;
};

/**
 * Writes `content` (charts of one two-point series, `Sales`) to `format` and reads it back: every word of `words`
 * once and in order outside tables, one table a chart (with its rows, `a 1` and `b 2`) and none in a heading,
 * bold only on the words given bold and the chart's title.
 */
const checkContainer = async (name: string, content: Run[], format: 'md' | 'rtf', words: string[], charts: number, boldWords: Set<string>, config: Record<string, unknown> = {}): Promise<string> => {
    const data = { title: 'Sales', labels: ['a', 'b'], dataSets: [{ name: 'S1', values: ['1', '2'] }], rawTexts: [] };
    const ast = { type: 'docx', metadata: {}, content, attachments: [{ type: 'chart', name: 'c', mimeType: 'application/xml', extension: 'xml', data: '', chartData: data }] } as any as OfficeParserAST;
    const written = String((await OfficeGenerator.generate(ast, format, { ...config, onWarning: () => { } } as any)).value);
    if (format === 'rtf') assert.ok(bracesBalanced(written), `${name} (${format}): its groups are balanced`);
    const back = await OfficeParser.parseOffice(Buffer.from(written), { fileType: format, onWarning: () => { } } as any) as OfficeParserAST;
    const nodes = readBackNodes(back);
    const tables = nodes.filter(n => n.node.type === 'table' && rowsOf(n.node).some(r => r.join(' ') === 'S1'));
    assert.strictEqual(tables.length, charts, `${name} (${format}): ${charts} chart tables read back; got ${written}`);
    for (const table of tables) {
        const rows = rowsOf(table.node).map(r => r.join(' '));
        assert.ok(rows.includes('a 1') && rows.includes('b 2'), `${name} (${format}): the chart's rows; got ${JSON.stringify(rows)}`);
    }
    assert.ok(!nodes.some(n => n.node.type === 'table' && n.inHeading), `${name} (${format}): no table in a heading`);
    const inCharts = new Set<any>();
    const mark = (node: any) => { inCharts.add(node); for (const child of node.children ?? []) mark(child); };
    for (const table of tables) mark(table.node);
    const texts = nodes.filter(n => n.node.type === 'text' && !inCharts.has(n.node));
    const line = texts.map(n => String(n.node.text ?? '')).join(' ').replace(/\s+/g, ' ');
    let at = -1;
    for (const word of words) {
        assert.strictEqual(occurrences(line, word), 1, `${name} (${format}): "${word}" reads back once; got ${JSON.stringify(line)} from ${JSON.stringify(written)}`);
        const next = line.indexOf(word);
        assert.ok(next > at, `${name} (${format}): "${word}" in order; got ${JSON.stringify(line)}`);
        at = next;
    }
    // (A heading's runs are left out; RTF writes a term bold.)
    if (!name.startsWith('heading') && !(format === 'rtf' && name.startsWith('term'))) for (const n of texts) {
        const text = String(n.node.text ?? '').trim();
        if (!text) continue;
        const shouldBe = text === 'Sales' || [...boldWords, ...(format === 'rtf' && name.startsWith('description') ? ['otherword'] : [])].some(word => text.includes(word));
        assert.strictEqual(!!n.node.formatting?.bold, shouldBe, `${name} (${format}): "${text}" is ${shouldBe ? '' : 'not '}bold`);
    }
    return written;
};

/**
 * A chart standing in a heading, a list item, a definition, a table cell, a quote, a note or around formatted runs:
 * in Markdown and RTF its table stands between the text before and after it, never joined into the line or the
 * heading, and the line's own text, formatting and nesting read back as they were.
 */
export async function testChartContainersMarkdownRtf(): Promise<void> {
    const chart = (): Run => ({ type: 'chart', text: 'a\nb', metadata: { attachmentName: 'c' } });
    const t = (text: string, formatting?: object): Run => ({ type: 'text', text, ...(formatting ? { formatting } : {}) });
    for (const format of CONTAINER_FORMATS) {
        for (const [name, container] of Object.entries(CONTAINERS)) {
            const words = name === 'note' ? ['refword', 'beforeword', 'afterword'] : name === 'description' ? ['otherword', 'beforeword', 'afterword'] : ['beforeword', 'afterword'];
            await checkContainer(name, container([t('beforeword '), chart(), t(' afterword')]), format, words, 1, new Set());
        }
        // The nested item after a list item holding a chart stays nested under it.
        const listed = await OfficeGenerator.generate({ type: 'docx', metadata: {}, attachments: [{ type: 'chart', name: 'c', mimeType: 'application/xml', extension: 'xml', data: '', chartData: { title: 'Sales', labels: ['a'], dataSets: [{ name: 'S1', values: ['1'] }], rawTexts: [] } }], content: CONTAINERS['list item of runs']([t('beforeword '), chart(), t(' afterword')]) } as any, format, { onWarning: () => { } } as any);
        const listBack = await OfficeParser.parseOffice(Buffer.from(String(listed.value)), { fileType: format, onWarning: () => { } } as any) as OfficeParserAST;
        const items = readBackNodes(listBack).filter(n => n.node.type === 'list');
        const nested = items.find(n => JSON.stringify(n.node).includes('nestedword'));
        assert.ok(nested && (nested.node.metadata?.indentation ?? 0) > (items[0].node.metadata?.indentation ?? 0), `${format}: the item after keeps its nesting; got ${JSON.stringify(items.map(i => i.node.metadata))}`);
    }

    // Markdown without HTML tables: a chart in a list item or definition is hoisted after the document, its title and
    // a reference to it in the line.
    for (const name of ['list item of runs', 'term', 'description'] as const) {
        const ast = { type: 'docx', metadata: {}, attachments: [{ type: 'chart', name: 'c', mimeType: 'application/xml', extension: 'xml', data: '', chartData: { title: 'Sales', labels: ['a', 'b'], dataSets: [{ name: 'S1', values: ['1', '2'] }], rawTexts: [] } }], content: CONTAINERS[name]([t('beforeword '), chart(), t(' afterword')]) } as any;
        const md = String((await OfficeGenerator.generate(ast, 'md', { mdConfig: { fallbackToHtml: false }, onWarning: () => { } } as any)).value);
        const from = name === 'list item of runs' ? 'a list item' : 'a definition';
        assert.ok(new RegExp(`beforeword \\*\\*Sales\\*\\* \\*\\(See Table 1 below\\)\\*\\s*afterword`).test(md), `md, ${name}: its title and a reference in the line; got ${JSON.stringify(md)}`);
        assert.ok(new RegExp(`\\*\\*Table 1 \\(Hoisted from ${from}\\):\\*\\*\\n+\\|[^\\n]*S1[^\\n]*\\|\\n\\| --- \\| --- \\|\\n\\| a \\| 1 \\|\\n\\| b \\| 2 \\|$`).test(md), `md, ${name}: the table after the document; got ${JSON.stringify(md)}`);
    }

    // The review's documents: every chart of a heading, list item and hyperlink paragraph, and of a paragraph, cell,
    // text box and footnote, reads back as a table.
    for (const [shape, file] of [['heading, item and link', headingAndItemDocx()], ['paragraph, cell, text box and footnote', nestedChartsDocx()]] as const) {
        const ast = await OfficeParser.parseOffice(file, { extractAttachments: true, onWarning: () => { } } as any) as OfficeParserAST;
        const charts = readBackNodes(ast).filter(n => n.node.type === 'chart' && ast.attachments.some(a => a.name === n.node.metadata?.attachmentName && a.chartData)).length;
        for (const format of CONTAINER_FORMATS) {
            const back = await OfficeParser.parseOffice(Buffer.from(String((await ast.to(format, { onWarning: () => { } } as any)).value)), { fileType: format, onWarning: () => { } } as any) as OfficeParserAST;
            const tables = readBackNodes(back).filter(n => n.node.type === 'table' && !JSON.stringify(n.node.children ?? []).includes('"type":"table"'));
            assert.ok(tables.length >= charts && charts >= 3, `${shape} (${format}): its ${charts} charts read back as tables; got ${tables.length}`);
            assert.ok(!readBackNodes(back).some(n => n.node.type === 'table' && n.inHeading), `${shape} (${format}): none in a heading`);
        }
    }

    // Property: every container, a chart at the start, middle, end or twice, runs plain, bold or linked, charts on and
    // off: every word once and in order, a table a chart (none with includeCharts false).
    const random = randomOf(0x3c41);
    const word = () => Array.from({ length: 8 }, () => String.fromCharCode(97 + Math.floor(random() * 26))).join('');
    const seen = new Set<string>();
    for (const [name, container] of Object.entries(CONTAINERS)) {
        for (const position of ['start', 'middle', 'end', 'twice'] as const) {
            for (const look of ['plain', 'bold', 'link'] as const) {
                for (const includeCharts of [true, false]) {
                    const words = [word(), word(), word()];
                    const run = (w: string): Run => ({ type: 'text', text: ` ${w} `, ...(look === 'bold' ? { formatting: { bold: true } } : {}), ...(look === 'link' ? { metadata: { link: 'https://example.test/' } } : {}) });
                    const line = position === 'start' ? [chart(), run(words[0])] : position === 'middle' ? [run(words[0]), chart(), run(words[1])] : position === 'end' ? [run(words[0]), chart()] : [run(words[0]), chart(), run(words[1]), chart(), run(words[2])];
                    const used = position === 'twice' ? words : position === 'middle' ? words.slice(0, 2) : words.slice(0, 1);
                    const all = name === 'note' ? ['refword', ...used] : name === 'description' ? ['otherword', ...used] : used;
                    const charts = includeCharts ? (position === 'twice' ? 2 : 1) : 0;
                    for (const format of CONTAINER_FORMATS) {
                        seen.add(`${name}/${position}/${look}/${includeCharts}/${format}`);
                        await checkContainer(`${name}, ${position}, ${look}${includeCharts ? '' : ', no charts'}`, container(line), format, all, charts, new Set(look === 'bold' ? used : []), { includeCharts });
                    }
                }
            }
        }
    }
    console.log(`  charts in containers: ${seen.size} distinct cases`);

    // Markdown: each part of a heading split at its chart is identified by its own text (the review's wrap-sdt.docx, a
    // Heading 1 `Head <chart> htail`, the chart in a content control): both parts had the whole heading's id.
    const base = unzipSync(new Uint8Array(fs.readFileSync(path.join(FILES, 'charts', 'word-charts.docx'))));
    const heading = `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Head </w:t></w:r><w:sdt><w:sdtPr/><w:sdtContent>${wordChart('rId2')}</w:sdtContent></w:sdt><w:r><w:t> htail</w:t></w:r></w:p>`;
    const sdt = Buffer.from(zipSync({ ...base, 'word/document.xml': strToU8(strFromU8(base['word/document.xml']).replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${heading}</w:body>`)) }));
    const sdtAst = await OfficeParser.parseOffice(sdt, { extractAttachments: true, onWarning: () => { } } as any) as OfficeParserAST;
    const sdtMd = String((await sdtAst.to('md', { generateIds: true, onWarning: () => { } } as any)).value);
    assert.deepStrictEqual(sdtMd.match(/^#.*$/gm), ['# Head {#head}', '# htail {#htail}'], `md: each heading line its own id; got ${JSON.stringify(sdtMd)}`);
    const sdtHtml = String((await (await OfficeParser.parseOffice(Buffer.from(sdtMd), { fileType: 'md', onWarning: () => { } } as any)).to('html', { onWarning: () => { } } as any)).value);
    const ids = [...sdtHtml.matchAll(/<h1[^>]* id="([^"]*)"/g)].map(m => m[1]);
    assert.deepStrictEqual(ids, ['head', 'htail'], `md read back as HTML: no id twice; got ${JSON.stringify(ids)}`);
    // A heading of its own anchor: the first part keeps it, the second its own id. A paragraph whose `.text` holds the
    // whole line (as readers set it), split in an HTML list item: each part its own text.
    const chartData = { title: 'Sales', labels: ['a', 'b'], dataSets: [{ name: 'S1', values: ['1', '2'] }], rawTexts: [] };
    const anchored = { ...chartAst(chartData, 0), content: [{ type: 'heading', text: 'Head  htail', metadata: { level: 2, anchorIds: ['own'] }, children: [{ type: 'text', text: 'Head ' }, { type: 'chart', text: 'a\nb', metadata: { attachmentName: 'c' } }, { type: 'text', text: ' htail' }] }] } as any as OfficeParserAST;
    const anchoredMd = String((await OfficeGenerator.generate(anchored, 'md', { generateIds: true, onWarning: () => { } } as any)).value);
    assert.deepStrictEqual(anchoredMd.match(/^#.*$/gm), ['## Head {#own}', '## htail {#htail}'], `md: the first part keeps the heading's anchor, the second its own id; got ${JSON.stringify(anchoredMd)}`);
    const inItem = { ...chartAst(chartData, 0), content: [{ type: 'list', metadata: { listType: 'unordered', indentation: 0 }, children: [{ type: 'paragraph', text: 'P before  after', children: [{ type: 'text', text: 'P before ' }, { type: 'chart', text: 'a\nb', metadata: { attachmentName: 'c' } }, { type: 'text', text: ' after' }] }] }] } as any as OfficeParserAST;
    const itemMd = String((await OfficeGenerator.generate(inItem, 'md', { onWarning: () => { } } as any)).value);
    assert.ok(occurrences(itemMd, 'P before') === 1 && occurrences(itemMd, 'after') === 1, `md: a paragraph split in an HTML list item, each part its own text; got ${JSON.stringify(itemMd)}`);
}

// ─── Native PDF, a chart in a text box, a heading, a list item or a note ──

/** The native engine's PDF of `ast` read back: its text nodes with their bounds (top down), its text and warnings. */
const nativePdf = async (ast: OfficeParserAST, config: Record<string, unknown> = {}) => {
    const warnings: any[] = [];
    const out = await OfficeGenerator.generate(ast, 'pdf', { pdfConfig: { engine: 'native' }, ...config, onWarning: (w: any) => warnings.push(w) } as any);
    const back = await OfficeParser.parseOffice(Buffer.from(out.value as Uint8Array), { fileType: 'pdf', ocr: false, pdfParserConfig: { separateProcess: false }, onWarning: () => { } } as any) as OfficeParserAST;
    const texts = readBackNodes(back).map(n => n.node).filter(n => n.type === 'text' && n.bounds);
    return { texts, text: String((await back.to('text')).value), warnings };
};

/**
 * A chart in a Word text box drawn within the box (it was drawn 180 pt high whatever the box, over what follows), and
 * a chart in a heading, list item or note (lines the engine lays out as runs) its data's lines, on lines of their
 * own, reported once (they ran on from the word before them, unreported).
 */
export async function testNativePdfChartRoom(): Promise<void> {
    const parse = async (file: Buffer) => await OfficeParser.parseOffice(file, { extractAttachments: true, onWarning: () => { } } as any) as OfficeParserAST;
    const nested = await parse(nestedChartsDocx());
    const { texts, warnings } = await nativePdf(nested);
    const find = (part: string) => texts.find((n: any) => String(n.text).includes(part)) as any;
    const isLegend = (n: any) => String(n.text).trim() === 'Margin';
    const legend = texts.find(isLegend) as any, inBox = find('In box:');
    assert.ok(legend && inBox, `the text box's chart is drawn; got ${JSON.stringify(texts.map((n: any) => n.text))}`);
    // The box is 2,000,000 EMU (157.5 pt) high, from its first line (drawn 180 pt high, its chart's legend was below it).
    const boxFoot = inBox.bounds.y + 157.5;
    assert.ok(legend.bounds.y + legend.bounds.height <= boxFoot + 1, `the chart's legend is inside the box: legend at ${legend.bounds.y}, box from ${inBox.bounds.y}`);
    // The footnote's scatter chart is its data's lines, each on its own.
    const footnote = find('Footnote chart:');
    assert.ok(footnote && !String(footnote.text).includes('150'), `the footnote's chart starts a line of its own; got ${JSON.stringify(footnote?.text)}`);
    assert.ok(texts.some((n: any) => String(n.text).trim() === '150'), 'each of its data lines on a line of its own');
    const lineCharts = (list: any[]) => list.filter(w => w.details?.reason === 'chart-in-text-line').length;
    assert.strictEqual(lineCharts(warnings), 1, 'reported once');
    // includeCharts: false: neither its lines nor the warning.
    const off = await nativePdf(nested, { includeCharts: false });
    assert.ok(!off.text.includes('Group A') && off.text.includes('Footnote chart:'), 'includeCharts: false leaves the footnote\'s chart out');
    assert.strictEqual(lineCharts(off.warnings), 0, 'and reports nothing of it');

    // A text box anchored at its foot: measured and drawn alike, the chart still inside it.
    const bottom = JSON.parse(JSON.stringify(nested));
    const box = readBackNodes(bottom).map(n => n.node).find((n: any) => n.type === 'drawing' && n.metadata?.float);
    assert.ok(box, 'the text box is a drawing of a float');
    box.metadata.textAnchor = 'bottom';
    const anchored = await nativePdf(bottom);
    const bottomLegend = anchored.texts.find(isLegend) as any;
    assert.ok(bottomLegend && bottomLegend.bounds.y + bottomLegend.bounds.height <= boxFoot + 1, `anchored at its foot, the chart stays inside the box: legend at ${bottomLegend?.bounds.y}`);

    // A chart in a tall text box is drawn as one in the flow (at most 180 pt of plot).
    const tall = await nativePdf(await parse(nestedChartsDocx(8_000_000)));
    const flowAst = { type: 'docx', metadata: {}, attachments: nested.attachments, content: [{ type: 'paragraph', children: [readBackNodes(nested).map(n => n.node).find((n: any) => n.type === 'chart' && n.metadata?.attachmentName === 'chart3.xml')] }] } as any as OfficeParserAST;
    const flow = await nativePdf(flowAst);
    const span = (list: any[]) => { const title = list.find((n: any) => String(n.text).includes('Revenue and margin')), mark = list.find((n: any) => String(n.text).includes('Margin') && !String(n.text).includes('margin')); return title && mark ? mark.bounds.y - title.bounds.y : NaN; };
    assert.ok(Math.abs(span(tall.texts) - span(flow.texts)) < 1, `a chart in a tall text box is as tall as one in the flow: ${span(tall.texts)} and ${span(flow.texts)}`);

    // A chart in a heading and in a list item: its data on lines of their own, reported once.
    const lines = await nativePdf(await parse(headingAndItemDocx()));
    const head = lines.texts.find((n: any) => String(n.text).startsWith('Head')) as any, item = lines.texts.find((n: any) => String(n.text).includes('Item')) as any;
    assert.ok(head && !String(head.text).includes('Jan'), `the heading's chart starts a line of its own; got ${JSON.stringify(head?.text)}`);
    assert.ok(item && !/Q1|2025/.test(String(item.text)), `the list item's too; got ${JSON.stringify(item?.text)}`);
    assert.ok(lines.text.includes('Jan') && lines.text.includes('2025'), 'their data is written');
    assert.strictEqual(lineCharts(lines.warnings), 1, 'reported once');
}

/**
 * Native PDF: a chart the engine writes as text (one it cannot draw, or one in a heading, list item or note) written
 * again within the chart-copy budget, charged what it lays out, its reference past it: a paragraph of a 20,000-value
 * chart framed 200 times took 96 seconds and 40 MB. Measured and drawn alike (the text box's and the note's).
 */
export async function testNativePdfChartTextCopies(): Promise<void> {
    const native = { pdfConfig: { engine: 'native' }, onWarning: () => { } } as any;
    const chart = barOf(Array.from({ length: 20_000 }, () => '1'), undefined, 'A', 'T');
    const parse = async (n: number) => await OfficeParser.parseOffice(docxFramed(chart, n), { extractAttachments: true, decompressionLimits: { maxRepeatedContent: 1_000_000 }, onWarning: () => { } } as any) as OfficeParserAST;
    const few = await parse(25), many = await parse(100);
    const pdfOf = async (ast: OfficeParserAST) => {
        const started = performance.now();
        const pdf = (await ast.to('pdf', native)).value as Uint8Array;
        return { pdf, ms: performance.now() - started };
    };
    const small = await pdfOf(few), large = await pdfOf(many);
    assert.ok(large.pdf.length <= 1.2 * small.pdf.length, `four times the paragraphs, the PDF about the same: ${small.pdf.length} bytes, then ${large.pdf.length}`);
    let ms = [small.ms, large.ms];
    if (ms[1] > 1.5 * ms[0] + 500) ms = [(await pdfOf(few)).ms, Math.min(ms[1], (await pdfOf(many)).ms)];
    assert.ok(ms[1] <= 1.5 * ms[0] + 500, `and its time: ${ms[0].toFixed(0)} ms, then ${ms[1].toFixed(0)} ms`);
    const text = await pdfText(large.pdf);
    assert.ok(occurrences(text, '1') >= 20_000 && text.includes('[Chart: T]'), `the first chart's lines, then its reference; got ${occurrences(text, '1')} lines and ${occurrences(text, '[Chart: T]')} references`);

    // In a heading and a list item (lines the engine lays out as runs), budget 0: the first whole, the others its reference.
    const labels = Array.from({ length: 200 }, (_, i) => `L${i}`);
    const big = { title: 'Big', labels, dataSets: [{ name: 'S', values: labels.map((_, i) => String(i)) }], rawTexts: [] };
    const bigText = ['Big', ...labels, 'S', ...labels.map((_, i) => String(i))].join('\n');
    const chartNode = () => ({ type: 'chart', text: bigText, metadata: { attachmentName: 'c' } });
    const lines = {
        ...chartAst(big, 0, 0),
        content: [
            { type: 'heading', metadata: { level: 1 }, children: [{ type: 'text', text: 'Head ' }, chartNode()] },
            { type: 'list', metadata: { listType: 'unordered', indentation: 0 }, children: [{ type: 'text', text: 'Item ' }, chartNode()] },
            { type: 'paragraph', children: [{ type: 'text', text: 'Para ' }, chartNode()] },
        ],
    } as any as OfficeParserAST;
    const linesOut = await nativePdf(lines);
    assert.strictEqual(occurrences(linesOut.text, 'L199'), 1, `in a heading, a list item and a paragraph: the chart's lines once; got ${JSON.stringify(linesOut.text.slice(0, 600))}`);
    assert.strictEqual(occurrences(linesOut.text, '[Chart: Big]'), 2, 'its reference at the other two');
    assert.strictEqual(linesOut.warnings.filter((w: any) => w.code === 'REPEATED_CONTENT_LIMIT_EXCEEDED').length, 1, 'reported once');

    // Measured, then drawn: the review's nested.docx (a paragraph, a cell, a text box and a footnote each holding a
    // chart), every chart the one 200-point chart the engine writes as text, budget 0: its lines once, its reference at
    // the others, and no line drawn over another (a note or a box measured with one text and drawn with the other would).
    // (Its text box kept clear of the text, wrapTopAndBottom: without a wrap element it floats over the lines below, as
    // Word draws it.)
    const wrapped = unzipSync(new Uint8Array(nestedChartsDocx()));
    wrapped['word/document.xml'] = strToU8(strFromU8(wrapped['word/document.xml']).replace(/(<wp:anchor>\s*<wp:extent [^>]*\/>)/, '$1<wp:wrapTopAndBottom/>'));
    const nested = await OfficeParser.parseOffice(Buffer.from(zipSync(wrapped)), { extractAttachments: true, onWarning: () => { } } as any) as OfficeParserAST;
    const charts = readBackNodes(nested).map(n => n.node).filter((n: any) => n.type === 'chart');
    for (const node of charts) Object.assign(node, { text: bigText, metadata: { ...node.metadata, attachmentName: 'big' } });
    for (const note of (nested as any).content.flatMap(function walk(n: any): any[] { return [...(n.notes ?? []), ...(n.children ?? []).flatMap(walk)]; })) {
        for (const node of readBackNodes({ content: [note] } as any).map(n => n.node)) if (node.type === 'chart') Object.assign(node, { text: bigText, metadata: { ...node.metadata, attachmentName: 'big' } });
    }
    const measured = { ...nested, config: { ...nested.config, decompressionLimits: { maxRepeatedContent: 0 } }, attachments: [...nested.attachments, { type: 'chart', name: 'big', mimeType: 'application/xml', extension: 'xml', data: '', chartData: big }] } as any as OfficeParserAST;
    const shown = await nativePdf(measured);
    assert.strictEqual(occurrences(shown.text, 'L199'), 1, `nested: the chart's lines once; got ${JSON.stringify(shown.text.slice(0, 800))}`);
    assert.ok(occurrences(shown.text, '[Chart: Big]') >= 3, `nested: its reference in the cell, the box and the note; got ${occurrences(shown.text, '[Chart: Big]')}`);
    // (Each page's lines, read back: a page is a node of the PDF's content.)
    const out = await OfficeGenerator.generate(measured, 'pdf', native);
    const back = await OfficeParser.parseOffice(Buffer.from(out.value as Uint8Array), { fileType: 'pdf', ocr: false, pdfParserConfig: { separateProcess: false }, onWarning: () => { } } as any) as OfficeParserAST;
    for (const page of back.content) {
        const texts = readBackNodes({ content: [page] } as any).map(n => n.node).filter((n: any) => n.type === 'text' && n.bounds?.width > 0 && n.bounds?.height > 0);
        for (let i = 0; i < texts.length; i++) for (let j = i + 1; j < texts.length; j++) {
            const a = texts[i].bounds, b = texts[j].bounds;
            const overlap = Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x) > 2 && Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y) > 2;
            assert.ok(!overlap, `nested: no line drawn over another; "${texts[i].text}" and "${texts[j].text}" at ${JSON.stringify(a)} and ${JSON.stringify(b)}`);
        }
    }
}

// ─── Chart.js, a bubble series of sizes and no X values ──────────────────

/** Runs an HTML page's Chart.js scripts in a context of their own with a stand-in Chart; the configs they give it. */
export const chartScriptConfigs = (html: string): any[] => {
    const json = new Map<string, string>();
    for (const m of html.matchAll(/<script type="application\/json" id="([^"]+)">([\s\S]*?)<\/script>/g)) json.set(m[1], m[2]);
    const configs: any[] = [];
    for (const [, code] of html.matchAll(/<script>([\s\S]*?)<\/script>/g)) {
        if (!code.includes('initChart')) continue;
        const document = { readyState: 'complete', getElementById: (id: string) => json.has(id) ? { textContent: json.get(id) } : { getContext: () => ({}) } };
        vm.runInNewContext(code, { document, window: { addEventListener() { } }, setTimeout, console: { error: (...args: unknown[]) => { throw new Error(args.map(String).join(' ')); } }, Map, JSON, Math, Number, Array, Object, String, Chart: class { constructor(_: unknown, config: unknown) { configs.push(config); } } });
    }
    return configs;
};

/**
 * A bubble series with sizes and no X values (`c:xVal` is optional in a bubble series) is drawn by the Chart.js
 * script as bubbles, each at its point's place, as every other writer treats it: it was drawn as bars.
 */
export async function testChartScriptBubbles(): Promise<void> {
    const part = chartPart(`<c:bubbleChart><c:ser><c:idx val="0"/><c:tx>${strRef(['B1'])}</c:tx><c:yVal>${numRef(['1', '2'])}</c:yVal><c:bubbleSize>${numRef(['3', '40'])}</c:bubbleSize></c:ser><c:axId val="1"/><c:axId val="2"/></c:bubbleChart><c:valAx><c:axId val="1"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>`);
    const ast = await OfficeParser.parseOffice(xlsxWith(part), { extractAttachments: true, onWarning: () => { } } as any) as OfficeParserAST;
    const name = ast.attachments.find(a => a.chartData)!.name;
    const only = { ...ast, content: [{ type: 'chart', text: 'B1', metadata: { attachmentName: name } }] } as any as OfficeParserAST;
    const [config] = chartScriptConfigs(String((await OfficeGenerator.generate(only, 'html', { onWarning: () => { } } as any)).value));
    assert.strictEqual(config?.type, 'bubble', `drawn as a bubble chart; got ${JSON.stringify(config?.type)}`);
    // (Made in the script's own context: compared as JSON.)
    assert.deepStrictEqual(JSON.parse(JSON.stringify(config.data.datasets)).map((set: any) => set.data.map((p: any) => [p.x, p.y, +p.r.toFixed(3)])), [[[1, 1, 4.275], [2, 2, 20]]], 'each point at its place, its radius by its size');
}

// ─── Text and chunk output, a chart's text on lines of its own ───────────

/**
 * Text and chunk output start a chart's text on a line of its own: it ran on from the word before it in text
 * (`In box:45658`), and chunks of a paragraph, cell or note holding a chart left the chart's text out.
 */
export async function testChartTextLines(): Promise<void> {
    const ast = await OfficeParser.parseOffice(nestedChartsDocx(), { extractAttachments: true, onWarning: () => { } } as any) as OfficeParserAST;
    const text = String((await ast.to('text')).value);
    assert.ok(/In box:\n45658\n/.test(text) && /Top para with chart\s*\nJan\n/.test(text) && /\n40\n\s*after\./.test(text), `text: a chart's lines stand apart from the words around it; got ${JSON.stringify(text)}`);
    assert.ok(/\bcell\s+Q1\b/.test(text) && /Footnote chart:\n150\n/.test(text), `text: in a cell and a note too; got ${JSON.stringify(text)}`);
    const chunks = ((await ast.to('chunks')).value as any[]).map(chunk => String(chunk.text));
    const all = chunks.join('\n');
    for (const [before, after] of [['Top para with chart', 'Jan'], ['40', 'after.'], ['In box:', '45658'], ['Footnote chart:', '150']]) {
        assert.ok(new RegExp(`${before.replace(/[.]/g, '\\.')}\\s*\\n\\s*${after.replace(/[.]/g, '\\.')}`).test(all), `chunks: "${before}" and "${after}" stand on lines of their own; got ${JSON.stringify(chunks)}`);
    }
    assert.ok(chunks.some(chunk => /\| cell Q1 Q2 .*1650 \| B \|/.test(chunk)), `chunks: a cell's chart is in its row, apart from the cell's word; got ${JSON.stringify(chunks)}`);
    // includeCharts: false: no chart text in either.
    const off = await OfficeParser.parseOffice(nestedChartsDocx(), { extractAttachments: true, onWarning: () => { } } as any) as OfficeParserAST;
    assert.ok(!String((await off.to('text', { includeCharts: false } as any)).value).includes('45658'), 'text: includeCharts false leaves charts out');
    assert.ok(!((await off.to('chunks', { includeCharts: false } as any)).value as any[]).some(chunk => String(chunk.text).includes('45658')), 'chunks: includeCharts false leaves charts out');
    // A chart alone in a paragraph: its lines, as before.
    const alone = chartAst({ title: 'T', labels: ['Lone'], dataSets: [{ name: 'S', values: ['7'] }], rawTexts: [] });
    (alone.content[0] as any).text = 'Lone\nS\n7';
    const aloneAst = { ...alone, content: [{ type: 'paragraph', children: alone.content }] } as any as OfficeParserAST;
    assert.strictEqual(String((await OfficeGenerator.generate(aloneAst, 'text', { onWarning: () => { } } as any)).value).trim(), 'Lone\nS\n7', 'text: a chart alone in a paragraph is its lines');
    assert.deepStrictEqual(((await OfficeGenerator.generate(aloneAst, 'chunks', { onWarning: () => { } } as any)).value as any[]).map(chunk => String(chunk.text).trim()), ['Lone\nS\n7'], 'chunks: likewise');
}

// ─── Property: chart text written again, budget 0 ─────────────────────────

/**
 * Random charts (1 to 40 series, axes shared or each its own, bar or scatter, 0 to 3 outer category levels, strings
 * of 10, 64, 65, 300 or 5,000 random letters, 1 to 6 chart nodes) with a repeat budget of 0: every distinct string
 * longer than 64 characters is written whole exactly once, and every shorter one at least once.
 */
export async function testChartTextRepeatsProperty(): Promise<void> {
    const random = randomOf(0x51c3);
    const lengths = [10, 64, 65, 300, 5000];
    let serial = 0;
    // A string of `length` random letters, made distinct by a serial number at its start.
    const text = (length: number) => { const id = `q${serial++}z`; return id + Array.from({ length: Math.max(0, length - id.length) }, () => String.fromCharCode(97 + Math.floor(random() * 26))).join(''); };
    const pickLength = () => lengths[Math.floor(random() * lengths.length)];
    const seen = new Set<string>();
    const formats = ['md', 'rtf', 'html', 'tex', 'docx', 'odt', 'pptx', 'xlsx'] as const;
    for (let round = 0; round < 30; round++) {
        const seriesCount = 1 + Math.floor(random() * 40), xy = random() < 0.5, shared = random() < 0.5, levels = Math.floor(random() * 4), nodes = 1 + Math.floor(random() * 6), points = 1 + Math.floor(random() * 3);
        seen.add(`${seriesCount}/${xy}/${shared}/${levels}/${nodes}/${points}`);
        // (Axes of their own for at most 31 series, and one pair the others share: a chart keeps 64 axes, MAX_CHART_AXES.)
        const axisOf = (s: number) => shared || s >= 31 ? 0 : s + 1;
        const axes = Array.from({ length: shared ? 1 : Math.min(seriesCount, 31) + 1 }, (_, a) => [{ id: `x${a}`, kind: 'value', title: text(pickLength()) }, { id: `y${a}`, kind: 'value', title: text(pickLength()) }]).flat();
        const labels = Array.from({ length: points }, () => text(pickLength()));
        const data: any = {
            title: text(pickLength()), labels, rawTexts: [], axes,
            dataSets: Array.from({ length: seriesCount }, (_, s) => ({
                name: text(pickLength()), values: Array.from({ length: points }, (_, i) => String(i + 1)),
                axisIds: [`x${axisOf(s)}`, `y${axisOf(s)}`],
                ...(xy ? { xValues: Array.from({ length: points }, (_, i) => String(i * 2)) } : {}),
            })),
            ...(levels && !xy ? { labelLevels: [{ labels }, ...Array.from({ length: levels }, () => ({ labels: Array.from({ length: points }, () => text(pickLength())) }))] } : {}),
        };
        // The strings a table holds: the title, series names, labels and outer labels; a scatter chart's axis titles.
        const strings: string[] = [data.title, ...data.dataSets.map((set: any) => set.name), ...(xy ? axes.filter(axis => data.dataSets.some((set: any) => set.axisIds.includes(axis.id))).map(axis => axis.title) : [...labels, ...(data.labelLevels ?? []).slice(1).flatMap((level: any) => level.labels)])];
        for (const format of formats) {
            const { text: written } = await writtenText(chartAst(data, nodes, 0), format);
            for (const s of strings) {
                const count = occurrences(written, s);
                if (s.length > 64) assert.strictEqual(count, 1, `${format}, round ${round} (${seriesCount} series, xy ${xy}, shared ${shared}, levels ${levels}, nodes ${nodes}, points ${points}): string ${strings.indexOf(s)} of ${s.length} characters (${s.slice(0, 8)}) is whole once; got ${count}`);
                else assert.ok(count >= 1, `${format}, round ${round} (${seriesCount} series, xy ${xy}, shared ${shared}, levels ${levels}, nodes ${nodes}, points ${points}): string ${strings.indexOf(s)} of ${s.length} characters (${s.slice(0, 8)}) is written`);
            }
        }
    }
    console.log(`  chart text repeats: ${seen.size} distinct random cases`);
}

/** Markdown: a definition list holding a chart, written as HTML, is given to onNode once, and its string replaces it. */
export async function testChartDefinitionListOnNode(): Promise<void> {
    const list = { type: 'definitionList', children: [{ type: 'definitionTerm', children: [{ type: 'text', text: 'term ' }, { type: 'chart', text: 'a', metadata: { attachmentName: 'c' } }] }, { type: 'definitionDescription', children: [{ type: 'text', text: 'desc' }] }] };
    const ast = { ...chartAst({ title: 'Sales', labels: ['a'], dataSets: [{ name: 'S1', values: ['1'] }], rawTexts: [] }), content: [list] } as any as OfficeParserAST;
    let calls = 0;
    const md = String((await OfficeGenerator.generate(ast, 'md', { onNode: (node: any) => { if (node === ast.content[0]) calls++; }, onWarning: () => { } } as any)).value);
    assert.strictEqual(calls, 1, `the definition list is given to onNode once; got ${calls}`);
    assert.ok(md.includes('<dl>') && md.includes('<table>'), `and written as HTML with its chart's table; got ${JSON.stringify(md)}`);
    const replaced = String((await OfficeGenerator.generate(ast, 'md', { onNode: (node: any) => node === ast.content[0] ? 'REPLACED' : undefined, onWarning: () => { } } as any)).value);
    assert.strictEqual(replaced.trim(), 'REPLACED', 'a string onNode gives it replaces it');
}

/**
 * Markdown and RTF ask onNode about a line holding a chart (a paragraph, heading, list item, cell, term or description)
 * before they write it in parts: `false` leaves out the line with its text and its chart's table, a string stands for
 * it, and the parts made of it are never given to the hook (it never saw the line: what it dropped was written).
 */
export async function testChartLineOnNode(): Promise<void> {
    const data = { title: 'ChartTitle', labels: ['lblA', 'lblB'], dataSets: [{ name: 'S1', values: ['1', '2'] }], rawTexts: [] };
    const chart = () => ({ type: 'chart', text: 'lblA\nlblB', metadata: { attachmentName: 'c' } });
    const runs = () => [{ type: 'text', text: 'SECRET before ' }, chart(), { type: 'text', text: ' AFTERWORD' }];
    const item = (children: unknown[]) => ({ type: 'list', metadata: { listType: 'unordered', indentation: 0 }, children });
    const table = (cell: unknown) => ({ type: 'table', children: [{ type: 'row', children: [cell, { type: 'cell', children: [{ type: 'text', text: 'othercell' }] }] }] });
    // Each container: its content, and the line in it holding the chart.
    type Shape = () => { content: any[]; line: any };
    const shapes: Record<string, Shape> = {
        paragraph: () => { const line = { type: 'paragraph', children: runs() }; return { content: [line], line }; },
        'anchored paragraph, chart first': () => { const line = { type: 'paragraph', metadata: { anchorIds: ['pa'] }, children: [chart(), { type: 'text', text: ' SECRET before AFTERWORD' }] }; return { content: [line], line }; },
        heading: () => { const line = { type: 'heading', metadata: { level: 1 }, children: runs() }; return { content: [line], line }; },
        'list item of runs': () => { const line = item(runs()); return { content: [line], line }; },
        'list item of paragraphs': () => { const line = { type: 'paragraph', children: runs() }; return { content: [item([line])], line }; },
        'cell of runs': () => { const line = { type: 'cell', children: runs() }; return { content: [table(line)], line }; },
        'cell of paragraphs': () => { const line = { type: 'paragraph', children: runs() }; return { content: [table({ type: 'cell', children: [line] })], line }; },
        'paragraph in an HTML cell': () => { const line = { type: 'paragraph', children: runs() }; return { content: [table({ type: 'cell', metadata: { colSpan: 2 }, children: [line] })], line }; },
        term: () => { const line = { type: 'definitionTerm', children: runs() }; return { content: [{ type: 'definitionList', children: [line, { type: 'definitionDescription', children: [{ type: 'text', text: 'otherdesc' }] }] }], line }; },
        description: () => { const line = { type: 'definitionDescription', children: runs() }; return { content: [{ type: 'definitionList', children: [{ type: 'definitionTerm', children: [{ type: 'text', text: 'otherterm' }] }, line] }], line }; },
        quote: () => { const line = { type: 'paragraph', metadata: { style: 'Quote' }, children: runs() }; return { content: [line], line }; },
    };
    const noHtml = { mdConfig: { fallbackToHtml: false } };
    const cases: [string, 'md' | 'rtf', Record<string, unknown>][] = [
        ...Object.keys(shapes).flatMap(name => CONTAINER_FORMATS.map((format): [string, 'md' | 'rtf', Record<string, unknown>] => [name, format, {}])),
        ['list item of runs', 'md', noHtml], ['term', 'md', noHtml], ['description', 'md', noHtml],
    ];
    const lineText = ['SECRET', 'lblA', 'lblB', 'ChartTitle', 'Hoisted', 'AFTERWORD'];
    for (const [name, format, config] of cases) {
        const build = () => {
            const { content, line } = shapes[name]();
            const ast = { ...chartAst(data, 0), content: [{ type: 'paragraph', children: [{ type: 'text', text: 'FIRSTWORD' }] }, ...content, { type: 'paragraph', children: [{ type: 'text', text: 'LASTWORD' }] }] } as any as OfficeParserAST;
            return { ast, line };
        };
        const label = `${name} (${format}${config === noHtml ? ', no HTML' : ''})`;
        const plain = String((await OfficeGenerator.generate(build().ast, format, { ...config, onWarning: () => { } } as any)).value);
        for (const verdict of [false, 'REPL', undefined] as const) {
            const { ast, line } = build();
            const originals = new Set<any>(readBackNodes(ast).map(n => n.node));
            const lineChildren = new Set<any>(line.children);
            const given: any[] = [];
            const out = String((await OfficeGenerator.generate(ast, format, { ...config, onNode: (node: any) => { given.push(node); return node === line ? verdict : undefined; }, onWarning: () => { } } as any)).value);
            assert.strictEqual(given.filter(n => n === line).length, 1, `${label}, ${verdict}: the line is given to onNode once`);
            const made = given.filter(n => n !== line && ((n.type === 'chart' && !originals.has(n)) || (n.children ?? []).some((child: any) => lineChildren.has(child))));
            assert.strictEqual(made.length, 0, `${label}, ${verdict}: no object made from the line is given to onNode; got ${JSON.stringify(made.map(n => n.type))}`);
            assert.ok(out.includes('FIRSTWORD') && out.includes('LASTWORD'), `${label}, ${verdict}: the rest of the document is written`);
            if (verdict === false) for (const word of lineText) assert.ok(!out.includes(word), `${label}: false leaves out "${word}"; got ${JSON.stringify(out.slice(-1500))}`);
            if (verdict === 'REPL') {
                assert.strictEqual(occurrences(out, 'REPL'), 1, `${label}: the string stands for the line once`);
                for (const word of lineText) assert.ok(!out.includes(word), `${label}: and none of the line is written ("${word}")`);
            }
            if (verdict === undefined) assert.strictEqual(out, plain, `${label}: a hook leaving everything alone changes nothing`);
        }
    }

    // The review's hl.docx (a heading, a list item and a link paragraph each holding a chart): dropped by the hook, none
    // of their text or charts is written, as in DOCX and HTML.
    const parsed = await OfficeParser.parseOffice(headingAndItemDocx(), { extractAttachments: true, onWarning: () => { } } as any) as OfficeParserAST;
    const dropped = new Set(parsed.content.slice(0, 3));
    for (const format of CONTAINER_FORMATS) {
        const out = String((await parsed.to(format, { onNode: (node: any) => dropped.has(node) ? false : undefined, onWarning: () => { } } as any)).value);
        assert.ok(out.includes('plain'), `hl.docx (${format}): the paragraph kept is written`);
        for (const word of ['Head', 'Item', 'link', 'Jan', 'Q1', '2025']) assert.ok(!out.includes(word), `hl.docx (${format}): "${word}" of a dropped line is not written; got ${JSON.stringify(out.slice(-1500))}`);
    }
}

/**
 * A document of every container a line holding a chart stands in (the top level, a quote, a heading, list items of runs
 * and of paragraphs, a nested item, cells of runs and of paragraphs, an HTML cell, a definition, a footnote), twice,
 * in three mixes of characters: its chart's title and two labels each of the mix.
 */
const orderHolders = (mix: string, mixIndex: number): OfficeParserAST => {
    const data = { title: `Ti ${mix}`, labels: [`${mix}1`, `${mix}2`], dataSets: [{ name: `S ${mix}`, values: ['1', '2'] }], rawTexts: [] };
    const t = (text: string) => ({ type: 'text', text });
    const chart = () => ({ type: 'chart', text: data.labels.join('\n'), metadata: { attachmentName: 'c' } });
    const item = (children: unknown[], indentation = 0) => ({ type: 'list', metadata: { listType: 'unordered', indentation }, children });
    const holders = (i: number): unknown[] => [
        chart(),
        { type: 'paragraph', children: [t(`b${i} `), chart(), t(` a${i}`)] },
        { type: 'paragraph', metadata: { anchorIds: [`bm${i}`] }, children: [chart(), t(` lead${i}`)] },
        { type: 'paragraph', metadata: { quote: true, quoteId: 'q' }, children: [t(`q${i} `), chart(), t(' qa')] },
        { type: 'paragraph', metadata: { quote: true, quoteId: 'q' }, children: [chart(), t(` qb${i}`)] },
        { type: 'heading', metadata: { level: 2 }, children: [t(`H${i} `), chart(), t(` Ha${i}`)] },
        { type: 'heading', metadata: { level: 2, anchorIds: [`hb${i}`] }, children: [chart(), t(` Hb${i}`)] },
        item([t(`item${i} `), chart(), t(' tail')]),
        item([{ type: 'paragraph', children: [t(`op${i}`), chart()] }, { type: 'paragraph', children: [t(`oq${i}`)] }]),
        item([t('nested')]), item([{ type: 'paragraph', children: [t(`np${i}`), chart(), t('nz')] }], 1),
        { type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [t(`c${i}`), chart()] }, { type: 'cell', children: [{ type: 'paragraph', children: [t(`cp${i}`), chart(), t('z')] }] }] }] },
        { type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [item([t('li')]), { type: 'paragraph', children: [t(`hc${i}`), chart(), t('hz')] }] }] }] },
        { type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [{ type: 'paragraph', children: [t(`fig${i}`), { type: 'image', metadata: { attachmentName: 'p' }, children: [] }, t('after')] }, { type: 'code', text: 'x=1', children: [] }] }] }] },
        { type: 'definitionList', children: [{ type: 'definitionTerm', children: [t(`T${i} `), chart()] }, { type: 'definitionDescription', children: [{ type: 'paragraph', children: [t(`D${i}`), chart(), t('dz')] }] }] },
        { type: 'paragraph', children: [{ type: 'text', text: `ref${i}`, notes: [{ type: 'note', metadata: { noteType: 'footnote', noteId: String(mixIndex * 10 + i) }, children: [{ type: 'paragraph', children: [t(`n${i} `), chart()] }] }] }] },
        { type: 'paragraph', children: [t(`plain${i}`)] },
    ];
    return { ...chartAst(data, 0), content: [...holders(0), ...holders(1)] } as any as OfficeParserAST;
};

/**
 * The order `format` asks onNode about the paragraphs and headings of `ast` (each by its place in a walk of the
 * content's children, notes left out), each counted the first time it is given; and how many times each line holding a
 * chart is given.
 */
const askingOrder = async (ast: OfficeParserAST, format: string, config: Record<string, unknown> = {}): Promise<{ order: number[]; lines: number; lineAskings: number[] }> => {
    const index = new Map<unknown, number>();
    const lines = new Set<unknown>();
    const walk = (node: any) => {
        if (node.type === 'paragraph' || node.type === 'heading') {
            index.set(node, index.size);
            if ((node.children ?? []).some((child: any) => child.type === 'chart')) lines.add(node);
        }
        for (const child of node.children ?? []) walk(child);
    };
    for (const node of ast.content) walk(node);
    const order: number[] = [];
    const askings = new Map<unknown, number>();
    await OfficeGenerator.generate(ast, format as any, {
        ...configOf(format, config),
        onNode: (node: any) => {
            const at = index.get(node);
            if (at !== undefined && !askings.has(node)) order.push(at);
            if (at !== undefined) askings.set(node, (askings.get(node) ?? 0) + 1);
        },
    } as any);
    return { order, lines: index.size, lineAskings: [...lines].map(line => askings.get(line) ?? 0) };
};

/**
 * Markdown asks onNode about a line holding a chart at its turn, in document order as every writer asks, not while the
 * container holding it is prepared (before what stands before it in the container is written): a hook whose answer
 * depends on what it has seen (one dropping a section from its heading on) answered every such line with the state at
 * the container's start, so a section's chart line was written, or one before it dropped. The line is asked once and
 * split afterwards, from the line as the hook left it.
 */
export async function testChartLineOnNodeOrder(): Promise<void> {
    const noHtml = { mdConfig: { fallbackToHtml: false } };
    const writers: [string, string, Record<string, unknown>][] = [
        ['md', 'md', {}], ['md without HTML', 'md', noHtml], ['md with ids', 'md', { generateIds: true }],
        ['rtf', 'rtf', {}], ['html', 'html', {}], ['docx', 'docx', {}], ['odt', 'odt', {}],
    ];

    // Every container, every writer: each paragraph and heading asked, in document order, the same sequence in all; each
    // line holding a chart given once in Markdown and RTF.
    for (const [mixIndex, mix] of ['a', '^&<"\'{}\\|*_`$%#~', 'x\ty  z\r\nw\u0001'].entries()) {
        let expected: number[] | undefined;
        for (const [name, format, config] of writers) {
            const { order, lines, lineAskings } = await askingOrder(orderHolders(mix, mixIndex), format, config);
            const label = `${name}, mix ${mixIndex}`;
            assert.ok(order.every((at, k) => k === 0 || at > order[k - 1]), `${label}: paragraphs and headings asked in document order; got ${order}`);
            assert.strictEqual(order.length, lines, `${label}: every paragraph and heading asked (${order.length} of ${lines})`);
            expected ??= order;
            assert.deepStrictEqual(order, expected, `${label}: the same sequence as Markdown's`);
            if (format === 'md' || format === 'rtf') assert.ok(lineAskings.every(count => count === 1), `${label}: each line holding a chart given once; got ${lineAskings}`);
        }
    }

    // Without charts: every paragraph and heading of the fixtures, in document order.
    for (const file of ['test.docx', 'test.odt', 'test.md', 'test.html', 'test.rtf', 'test.pptx', 'test.epub']) {
        const ast = await OfficeParser.parseOffice(fs.readFileSync(path.join(FILES, file)), { extractAttachments: true, fileType: path.extname(file).slice(1), onWarning: () => { } } as any) as OfficeParserAST;
        for (const [name, config] of [['md', {}], ['md without HTML', noHtml]] as const) {
            const { order, lines } = await askingOrder(ast, 'md', config);
            assert.ok(order.every((at, k) => k === 0 || at > order[k - 1]), `${file} (${name}): paragraphs and headings asked in document order`);
            assert.strictEqual(order.length, lines, `${file} (${name}): every paragraph and heading asked (${order.length} of ${lines})`);
        }
    }

    const t = (text: string) => ({ type: 'text', text });
    const chart = () => ({ type: 'chart', text: 'a', metadata: { attachmentName: 'c' } });

    // The review's shape: P1, P2 holding a chart, P3, and a hook dropping what follows P1 until P3. Markdown asked about P2
    // first and wrote it; now P1, P2, P3 as every writer, and P2 dropped.
    const skipped = (): OfficeParserAST => ({ ...chartAst({ title: 'T', labels: ['a'], dataSets: [{ name: 'S', values: ['1'] }], rawTexts: [] }, 0), content: [{ type: 'paragraph', children: [t('P1')] }, { type: 'paragraph', children: [t('P2 '), chart()] }, { type: 'paragraph', children: [t('P3')] }] } as any as OfficeParserAST);
    for (const [name, format, config] of [['md', 'md', {}], ['md without HTML', 'md', noHtml], ['rtf', 'rtf', {}], ['html', 'html', {}], ['docx', 'docx', {}]] as const) {
        const order: string[] = [];
        let skipping = false;
        const onNode = (node: any) => {
            if (node.type !== 'paragraph') return undefined;
            const text = String(node.children?.[0]?.text ?? '').trim();
            order.push(text);
            if (text === 'P1') skipping = true;
            if (text === 'P3') skipping = false;
            return skipping && text !== 'P1' ? false : undefined;
        };
        const { text } = await writtenText(skipped(), format, { ...config, onNode });
        assert.deepStrictEqual(order, ['P1', 'P2', 'P3'], `${name}: P1, P2 and P3 asked in order`);
        assert.ok(!text.includes('P2') && text.includes('P1') && text.includes('P3'), `${name}: P2 dropped, P1 and P3 written; got ${JSON.stringify(text.slice(-600))}`);
    }

    // A hook dropping a section (from the heading "Confidential" to the next): none of it written, its chart line
    // standing at the top level or in a container, in every writer.
    const payroll = { title: 'Payroll', labels: ['Alice', 'Bob'], dataSets: [{ name: 'Salary', values: ['91000', '87000'] }], rawTexts: [] };
    const heading = (text: string) => ({ type: 'heading', metadata: { level: 1 }, children: [t(text)] });
    const sections: Record<string, (line: unknown) => unknown> = {
        'the top level': line => line,
        'a list item of paragraphs': line => ({ type: 'list', metadata: { listType: 'unordered', indentation: 0 }, children: [line] }),
        'a cell of paragraphs': line => ({ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [line] }] }] }),
        'an HTML cell': line => ({ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [{ type: 'list', metadata: { listType: 'unordered', indentation: 0 }, children: [t('li')] }, line] }] }] }),
        'a definition description': line => ({ type: 'definitionList', children: [{ type: 'definitionTerm', children: [t('term')] }, { type: 'definitionDescription', children: [line] }] }),
        'a quote paragraph': () => ({ type: 'paragraph', metadata: { quote: true, quoteId: 'q' }, children: [t('Secret plan: see chart '), { type: 'chart', text: 'Alice\nBob', metadata: { attachmentName: 'c' } }] }),
    };
    const confidential = (holder: (line: unknown) => unknown): OfficeParserAST => ({
        ...chartAst(payroll, 0),
        content: [
            heading('Public'), { type: 'paragraph', children: [t('Intro.')] },
            heading('Confidential'), holder({ type: 'paragraph', children: [t('Secret plan: see chart '), { type: 'chart', text: 'Alice\nBob', metadata: { attachmentName: 'c' } }] }),
            { type: 'paragraph', children: [t('Secret footnote text.')] },
            heading('Appendix'), { type: 'paragraph', children: [t('End.')] },
        ],
    } as any as OfficeParserAST);
    const redacting = () => {
        let skipping = false;
        return (node: any) => {
            if (node.type === 'heading') { skipping = (node.children?.[0]?.text ?? node.text) === 'Confidential'; return skipping ? false : undefined; }
            if (skipping && node.type === 'paragraph') return false;
            return undefined;
        };
    };
    const everyWriter: [string, string, Record<string, unknown>][] = [...writers, ['tex', 'tex', {}], ['text', 'text', {}]];
    for (const [where, holder] of Object.entries(sections)) {
        for (const [name, format, config] of where === 'the top level' ? everyWriter : writers.slice(0, 2)) {
            const { text } = await writtenText(confidential(holder), format, { ...config, onNode: redacting() });
            for (const word of ['Secret plan', '91000', 'Secret footnote']) assert.ok(!text.includes(word), `${name}, the section's chart line in ${where}: "${word}" not written; got ${JSON.stringify(text.slice(-800))}`);
            for (const word of ['Intro.', 'End.']) assert.ok(text.includes(word), `${name}, the section's chart line in ${where}: "${word}" written`);
        }
    }

    // word-charts.docx under a heading "Confidential", the hook dropping it and every paragraph, table and list after:
    // none of its charts' labels or values written, as DOCX, HTML and RTF.
    const words = await OfficeParser.parseOffice(fs.readFileSync(path.join(FILES, 'charts', 'word-charts.docx')), { extractAttachments: true, onWarning: () => { } } as any) as OfficeParserAST;
    // (Its charts' titles, series names and values, none of them elsewhere in an output.)
    const shown = ['Monthly sales', 'Quarters by year', 'Revenue', 'Height and weight', 'Group A', '1250000', '1650'];
    for (const word of shown) assert.ok(JSON.stringify(words.attachments).includes(word), `word-charts.docx: its charts hold "${word}"`);
    for (const [name, format, config] of [['md', 'md', {}], ['md without HTML', 'md', noHtml], ['docx', 'docx', {}], ['html', 'html', {}], ['rtf', 'rtf', {}]] as const) {
        const ast = { ...words, content: [heading('Confidential'), ...words.content] } as OfficeParserAST;
        let skipping = false;
        const onNode = (node: any) => {
            if (node.type === 'heading') { skipping = node.children?.[0]?.text === 'Confidential'; return skipping ? false : undefined; }
            if (skipping && (node.type === 'paragraph' || node.type === 'table' || node.type === 'list')) return false;
            return undefined;
        };
        const { text } = await writtenText(ast, format, { ...config, onNode });
        for (const word of shown) assert.ok(!text.includes(word), `word-charts.docx (${name}): "${word}" of a dropped section not written`);
    }

    // A hook that changes a line holding a chart and lets it be written: the line is written as the hook left it (a run
    // it removed not written; children replaced by text alone written as that text, no table), each line asked once.
    const changed = (): OfficeParserAST => {
        const line = (words: string, ...rest: unknown[]) => [t(words), { type: 'chart', text: 'L1\nL2', metadata: { attachmentName: 'c' } }, ...rest];
        const item = (children: unknown[]) => ({ type: 'list', metadata: { listType: 'unordered', indentation: 0 }, children });
        return {
            ...chartAst({ title: 'Ti', labels: ['L1', 'L2'], dataSets: [{ name: 'S', values: ['1', '2'] }], rawTexts: [] }, 0),
            content: [
                { type: 'paragraph', children: line('SECRET top ', t(' after')) },
                { type: 'heading', metadata: { level: 2 }, children: line('SECRET head ', t(' tail')) },
                item([{ type: 'paragraph', children: line('SECRET item ') }]),
                { type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [item([t('li')]), { type: 'paragraph', children: line('SECRET html ') }] }] }] },
                { type: 'paragraph', children: line('NOCHART ') },
            ],
        } as any as OfficeParserAST;
    };
    for (const [name, config] of [['md', {}], ['md without HTML', noHtml]] as const) {
        const asked = new Map<unknown, number>();
        const onNode = (node: any) => {
            if ((node.type !== 'paragraph' && node.type !== 'heading') || !node.children?.some((child: any) => child.type === 'chart')) return undefined;
            asked.set(node, (asked.get(node) ?? 0) + 1);
            if (node.children[0].text === 'NOCHART ') { node.children = [t('chart removed by the hook')]; return undefined; }
            node.children = node.children.filter((child: any) => !String(child.text ?? '').startsWith('SECRET'));
            return undefined;
        };
        const md = String((await OfficeGenerator.generate(changed(), 'md', { ...config, onNode, onWarning: () => { } } as any)).value);
        assert.ok(!md.includes('SECRET'), `${name}: a run the hook removed is not written; got ${JSON.stringify(md)}`);
        assert.ok(md.includes('chart removed by the hook'), `${name}: a line the hook made text is written as that text`);
        assert.strictEqual(occurrences(md, '| L1 |') + occurrences(md, '<td>L1</td>'), 4, `${name}: the four lines keeping their chart write its table, the fifth none; got ${JSON.stringify(md)}`);
        assert.deepStrictEqual([...asked.values()], [1, 1, 1, 1, 1], `${name}: each line asked once`);
    }
}

/**
 * XLSX and ODS output find each chart node's data by its attachment's name through a table made once: a scan of the
 * attachments a chart node took time of charts times attachments (four times as many of each, sixteen times the time).
 * Ten attachments a chart node, so the scan, not the sheet each chart's table stands on, is what the time measures. Each
 * is found by its name, a chart node of no name or an empty one too.
 */
export async function testChartLookupLinear(): Promise<void> {
    const astOf = (n: number) => ({
        type: 'docx', metadata: {},
        attachments: Array.from({ length: 10 * n }, (_, i) => ({ type: 'chart', name: `c${i}`, mimeType: 'application/xml', extension: 'xml', data: '', chartData: { title: 'T', labels: ['lblA'], dataSets: [{ name: 'S', values: ['1'] }], rawTexts: [] } })),
        content: Array.from({ length: n }, () => ({ type: 'chart', text: '', metadata: { attachmentName: `c${10 * n - 1}` } })),
    }) as any as OfficeParserAST;
    const small = astOf(2000), large = astOf(8000);
    // Found by name (`attachments.find(a => a?.name === name)`): a chart node with no name the first attachment
    // with none, an empty name the first named so, the first of two of a name, names an object's prototype holds, a number.
    const chartOf = (key: string) => ({ title: `T${key}`, labels: [`L${key}`], dataSets: [{ name: 'S', values: ['1'] }], rawTexts: [] });
    const attachment = (name: unknown, key: string) => ({ type: 'chart', name, mimeType: 'application/xml', extension: 'xml', data: '', chartData: chartOf(key) });
    const node = (name?: unknown) => ({ type: 'chart', text: 'x', metadata: name === undefined ? {} : { attachmentName: name } });
    const lookups: [string, unknown[], unknown[], string[], string[]][] = [
        ['no name', [node()], [attachment(undefined, 'A'), attachment('n', 'B')], ['LA'], ['LB']],
        ['an empty name', [node('')], [attachment('', 'E'), attachment('n', 'B')], ['LE'], ['LB']],
        ['two of a name', [node('n')], [attachment('n', 'First'), attachment('n', 'Second')], ['LFirst'], ['LSecond']],
        ['__proto__ and constructor', [node('__proto__'), node('constructor')], [attachment('__proto__', 'P'), attachment('constructor', 'C')], ['LP', 'LC'], []],
        ['a number', [node(5)], [attachment(5, 'Num')], ['LNum'], []],
    ];
    for (const [name, content, attachments, found, notFound] of lookups) {
        for (const format of ['xlsx', 'ods'] as const) {
            const { text } = await writtenBytes({ type: 'docx', metadata: {}, content, attachments } as any as OfficeParserAST, format);
            for (const label of found) assert.ok(text.includes(`>${label}<`), `${format}, ${name}: the chart's table is written (${label})`);
            for (const label of notFound) assert.ok(!text.includes(`>${label}<`), `${format}, ${name}: not another chart's (${label})`);
        }
    }
    for (const format of ['xlsx', 'ods'] as const) {
        assert.ok((await writtenBytes(astOf(3), format)).text.includes('lblA'), `${format}: the chart named is found and its table written`);
        await linearTime(`${format}: 8,000 chart nodes among 80,000 attachments`, scale => OfficeGenerator.generate(scale === 1 ? small : large, format, { onWarning: () => { } } as any));
    }
}

// ─── Charts in EPUB ──────────────────────────────────────────────────────────

/** The fixtures holding charts, parsed with their data. */
const CHART_FIXTURES = ['test.xlsx', 'test.pptx', 'charts/word-charts.docx', 'charts/inline-chart.odt', 'test.ods', 'test.odp'];

/** Every node of `content`, depth first. */
const allNodes = (content: any[] | undefined): any[] => {
    const nodes: any[] = [];
    const walk = (list: any[] | undefined) => { for (const node of list ?? []) { nodes.push(node); walk(node.children); } };
    walk(content);
    return nodes;
};

/** A cell's text as read back. */
const cellTextOf = (node: any): string => node.text ?? (node.children ?? []).map(cellTextOf).join('');

/** What a chart's data holds that a reader is shown: its title, labels, series names and values (as shown). */
const chartStrings = (data: ChartData): string[] => {
    const strings: string[] = [];
    if (typeof data.title === 'string' && data.title) strings.push(data.title);
    for (const label of (data as any).shownLabels ?? data.labels ?? []) if (label) strings.push(label);
    for (const set of data.dataSets ?? []) {
        if (set.name) strings.push(set.name);
        for (const value of (set as any).shownValues ?? set.values ?? []) if (value) strings.push(value);
    }
    return strings;
};

/** An EPUB's first chapter. */
const chapterOf = (value: unknown): string => strFromU8(unzipSync(new Uint8Array(value as any))['OEBPS/chapter1.xhtml']);

/** Each chart's caption and table as HTML writes them with no chart script (`<p><strong>` title, then its table). */
const captionedTables = (html: string): string[] => [...html.matchAll(/<p><strong>[^<]*<\/strong><\/p>(?:<div class="table-container">)?<table>[\s\S]*?<\/table>/g)].map(m => m[0]);

/**
 * EPUB writes each chart as its data's table under its title, as HTML writes one with no chart script: a reading
 * system runs no script, and the Chart.js canvas EPUB wrote showed nothing (its data, in a script, was stripped). Not
 * reported: it is the form a chart takes in EPUB. HTML is as before: a canvas a chart node, and with `chartJsSrc: ''`
 * the same table, reported at each. EPUB's result lists what its HTML reported (the chart copies past the budget).
 */
export async function testChartsInEpub(): Promise<void> {
    for (const file of CHART_FIXTURES) {
        const ast = await OfficeParser.parseOffice(path.join(FILES, file), { extractAttachments: true, onWarning: () => { } } as any) as OfficeParserAST;
        const charts = allNodes(ast.content).filter(n => n.type === 'chart');
        assert.ok(charts.length > 0, `${file}: holds charts`);
        const expected = new Set<string>();
        for (const chart of charts) {
            const data = ast.attachments.find(a => a.name === chart.metadata?.attachmentName)?.chartData;
            assert.ok(data, `${file}: each chart has its data`);
            for (const text of chartStrings(data!)) expected.add(text);
        }

        const heard: any[] = [];
        const out = await OfficeGenerator.generate(ast, 'epub', { onWarning: (w: any) => heard.push(w) } as any);
        const chapter = chapterOf(out.value);
        assert.ok(!chapter.includes('<canvas'), `${file}: EPUB writes no canvas`);
        assert.deepStrictEqual(heard.map(w => w.code), [], `${file}: EPUB reports nothing for its charts`);
        assert.deepStrictEqual(out.messages.map(m => m.code), [], `${file}: nor lists anything`);
        const back = await OfficeParser.parseOffice(Buffer.from(out.value as any), { fileType: 'epub', onWarning: () => { } } as any) as OfficeParserAST;
        const cells = new Set(allNodes(back.content).filter(n => n.type === 'cell').map(n => cellTextOf(n).trim()));
        const text = String((await back.to('text')).value);
        const missing = [...expected].filter(t => !cells.has(t.trim()) && !text.includes(t));
        assert.deepStrictEqual(missing, [], `${file}: EPUB reads back every chart's title, labels, series and values`);

        // HTML as before: a canvas a chart node, unreported; with no chart script the caption and table, reported at each.
        const drawnWarnings: any[] = [];
        const drawn = String((await OfficeGenerator.generate(ast, 'html', { onWarning: (w: any) => drawnWarnings.push(w) } as any)).value);
        assert.strictEqual(occurrences(drawn, '<canvas'), charts.length, `${file}: HTML draws each chart node on a canvas`);
        assert.ok(!drawnWarnings.some(w => w.code === 'CONTENT_NOT_REPRESENTABLE'), `${file}: and reports none`);
        const tabledWarnings: any[] = [];
        const tabled = String((await OfficeGenerator.generate(ast, 'html', { htmlConfig: { chartJsSrc: '', standalone: false }, onWarning: (w: any) => tabledWarnings.push(w) } as any)).value);
        assert.ok(!tabled.includes('<canvas'), `${file}: HTML with no chart script writes no canvas`);
        assert.strictEqual(tabledWarnings.filter(w => w.code === 'CONTENT_NOT_REPRESENTABLE' && w.details?.reason === 'chart').length, charts.length, `${file}: and reports each chart node`);
        const tables = captionedTables(tabled);
        assert.strictEqual(tables.length, charts.length, `${file}: HTML with no chart script writes a caption and table a chart node`);
        for (const table of tables) assert.ok(chapter.includes(table), `${file}: EPUB writes the chart's caption and table as HTML does: ${table.slice(0, 80)}`);
    }

    // The chart copies past the budget: reported once, heard once, and listed in EPUB's result.
    const sales = { title: 'Sales', labels: Array.from({ length: 10 }, (_, i) => `L${i}`), dataSets: [{ name: 'Units', values: Array.from({ length: 10 }, (_, i) => String(1000 + i)) }], rawTexts: [] };
    const heard: string[] = [];
    const past = await OfficeGenerator.generate(chartAst(sales, 3, 0), 'epub', { onWarning: (w: any) => heard.push(w.code) } as any);
    assert.deepStrictEqual(heard, ['REPEATED_CONTENT_LIMIT_EXCEEDED'], 'EPUB: the copies past the budget are heard once');
    assert.deepStrictEqual(past.messages.map(m => m.code), ['REPEATED_CONTENT_LIMIT_EXCEEDED'], 'EPUB: and listed once in its result');
    const pastChapter = chapterOf(past.value);
    assert.strictEqual(occurrences(pastChapter, 'L9'), 1, 'EPUB: the first table whole');
    assert.strictEqual(occurrences(pastChapter, '[Chart: Sales]'), 2, 'EPUB: the reference at the other two');
    // A configuration problem is resolved before the chapter is written: listed once, as often as it is heard.
    const configHeard: string[] = [];
    const bogus = await OfficeGenerator.generate(chartAst(sales), 'epub', { mdConfig: { dialect: 'bogus' }, htmlConfig: { standalone: { styles: 'nope' } }, onWarning: (w: any) => configHeard.push(w.code) } as any);
    const listed = bogus.messages.filter(m => m.code === 'INVALID_CONFIG_VALUE');
    assert.ok(listed.length > 0, 'EPUB: a configuration problem is listed');
    assert.strictEqual(listed.length, configHeard.filter(code => code === 'INVALID_CONFIG_VALUE').length, 'EPUB: as often as it is heard');
    assert.strictEqual(new Set(listed.map(m => m.message)).size, listed.length, 'EPUB: each configuration problem once');

    // A chart in a heading, a list item, a paragraph and a cell: the chapter is XML a reading system opens.
    const chart = { type: 'chart', text: 'Q1\nQ2', metadata: { attachmentName: 'c' } };
    const held = {
        ...chartAst({ title: 'Sales & <co>', labels: ['Q1', 'Q2'], dataSets: [{ name: 'Units', values: ['10', '20'] }], rawTexts: [] }, 0),
        content: [
            { type: 'heading', metadata: { level: 2 }, children: [{ type: 'text', text: 'H ' }, chart] },
            { type: 'list', metadata: { listType: 'unordered', indentation: 0 }, children: [{ type: 'text', text: 'item ' }, chart] },
            { type: 'paragraph', children: [{ type: 'text', text: 'para ' }, chart, { type: 'text', text: ' after' }] },
            { type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [chart] }] }] },
        ],
    } as any as OfficeParserAST;
    const heldChapter = chapterOf((await OfficeGenerator.generate(held, 'epub', { onWarning: () => { } } as any)).value);
    assert.doesNotThrow(() => parseXmlString(heldChapter), 'EPUB: a chart in a heading, list item, paragraph and cell is well-formed XML');
    assert.strictEqual(occurrences(heldChapter, 'Sales &amp; &lt;co&gt;'), 4, 'EPUB: each chart\'s title, escaped, at each');
    assert.ok(!heldChapter.includes('<canvas'), 'EPUB: and no canvas');
}

/** test.xlsx with its chart part `chart`, the drawing on its second sheet relating it `n` times (one chart node each). */
const xlsxRelated = (chart: string, n: number) => {
    const base = unzipSync(new Uint8Array(fs.readFileSync(path.join(FILES, 'test.xlsx'))));
    const rels = strFromU8(base['xl/drawings/_rels/drawing2.xml.rels']);
    const rel = rels.match(/<Relationship [^>]*chart1[^>]*\/>/)![0];
    return Buffer.from(zipSync({ ...base, 'xl/charts/chart1.xml': strToU8(chart), 'xl/drawings/_rels/drawing2.xml.rels': strToU8(rels.replace(rel, Array.from({ length: n }, (_, i) => rel.replace('rId1', `rIdC${i}`)).join(''))) }, { level: 9 }));
};

/** inline-chart.odt with its chart's table `rows` rows long and the paragraph holding the chart `n` times. */
const odtFramed = (rows: number, n: number) => {
    const base = unzipSync(new Uint8Array(fs.readFileSync(path.join(FILES, 'charts', 'inline-chart.odt'))));
    const cell = (value: number) => `<table:table-cell office:value-type="float" office:value="${value}"><text:p>${value}</text:p></table:table-cell>`;
    const object = strFromU8(base['Object 2/content.xml']).replace(/<table:table-rows>[\s\S]*<\/table:table-rows>/, `<table:table-rows>${Array.from({ length: rows }, (_, i) => `<table:table-row>${cell(i)}${cell(1000 + i)}${cell(0.5)}</table:table-row>`).join('')}</table:table-rows>`);
    const content = strFromU8(base['content.xml']);
    const para = content.match(/<text:p text:style-name="Standard">Revenue and margin[\s\S]*?<\/text:p>/)![0];
    return Buffer.from(zipSync({ ...base, 'Object 2/content.xml': strToU8(object), 'content.xml': strToU8(content.replace(para, para.repeat(n))) }, { level: 9 }));
};

/**
 * EPUB's chart tables charged what they write: each chart file framed 50 and 200 times (a 2,000-value chart, a repeat
 * budget of 256 KiB) within one node's output, the budget and a reference a node, the 200 reported (in the result
 * too), in linear time; and for every class of text, what keeping a chart's copies writes in the chapter at most what
 * HTML (charts as tables) writes for them, which charges exactly what it writes.
 */
export async function testEpubChartCopies(): Promise<void> {
    const values = Array.from({ length: 2000 }, () => '1234567890');
    const shapes: [string, (n: number) => Buffer][] = [
        ['XLSX relationships', n => xlsxRelated(barOf(values, values, 'A', 'Sales'), n)],
        ['PPTX frames', n => pptxWith(barOf(values, values, 'A', 'Sales'), n)],
        ['DOCX paragraphs', n => docxFramed(barOf(values, undefined, 'A', 'Sales'), n)],
        ['ODT frames', n => odtFramed(700, n)],
    ];
    const parse = async (file: Buffer) => await OfficeParser.parseOffice(file, { extractAttachments: true, decompressionLimits: { maxRepeatedContent: 256 * 1024 }, onWarning: () => { } } as any) as OfficeParserAST;
    for (const [shape, file] of shapes) {
        const one = await parse(file(1)), few = await parse(file(50)), many = await parse(file(200));
        assert.strictEqual(allNodes(many.content).filter(n => n.type === 'chart').length, 200, `${shape}: 200 chart nodes`);
        const first = (await writtenBytes(one, 'epub')).bytes;
        const at = async (ast: OfficeParserAST, n: number) => {
            const heard: string[] = [];
            const out = await OfficeGenerator.generate(ast, 'epub', { onWarning: (w: any) => heard.push(w.code) } as any);
            const bytes = bytesOf(out.value);
            assert.ok(bytes <= first + repeatedContentOfAst(ast) + 600 * n, `${shape}: ${n} chart nodes write ${bytes} bytes of EPUB, one ${first}`);
            const spent = heard.includes('REPEATED_CONTENT_LIMIT_EXCEEDED');
            assert.strictEqual(out.messages.some(m => m.code === 'REPEATED_CONTENT_LIMIT_EXCEEDED'), spent, `${shape}: ${n} chart nodes: the result lists what was heard`);
            return { bytes, spent };
        };
        const n = await at(few, 50), n4 = await at(many, 200);
        assert.ok(n4.spent, `${shape}: 200 chart nodes spend the budget, reported`);
        const grown = repeatedContentOfAst(many) - repeatedContentOfAst(few) + first;
        if (n.spent) assert.ok(n4.bytes - n.bytes <= 150 * 600 + grown, `${shape}: 150 more chart nodes add ${n4.bytes - n.bytes} bytes, their references and what the larger budget holds alone`);
        await linearTime(`${shape}, epub`, scale => OfficeGenerator.generate(scale === 1 ? few : many, 'epub', { onWarning: () => { } } as any));
    }

    // Every class of text: a table of 40 labels of 100 characters at two nodes, a title of 1,000 at two, an X axis title
    // of 1,000 over 20 series, 30 series named 100 characters at two.
    const of = (unit: string, length: number) => unit.repeat(Math.ceil(length / unit.length));
    const classShapes: Record<string, [(unit: string) => unknown, number]> = {
        tbl: [unit => ({ title: of(unit, 100), labels: Array.from({ length: 40 }, (_, i) => of(unit, 100) + i), dataSets: [{ name: of(unit, 100), values: Array.from({ length: 40 }, () => of(unit, 100)) }], rawTexts: [] }), 2],
        title: [unit => ({ title: `X${of(unit, 1000)}Y`, labels: ['a'], dataSets: [{ name: 'S', values: ['1'] }], rawTexts: [] }), 2],
        xy: [unit => ({ title: 'T', labels: [], xAxisTitle: `X${of(unit, 1000)}Y`, yAxisTitle: 'Y', dataSets: Array.from({ length: 20 }, () => ({ name: 's', values: ['1'], xValues: ['2'] })), rawTexts: [] }), 1],
        xyt: [unit => ({ title: 'T', labels: [], dataSets: Array.from({ length: 30 }, () => ({ name: of(unit, 100), values: ['1'], xValues: ['2'] })), rawTexts: [] }), 2],
    };
    const G = 1024 * 1024 * 1024;
    // (The chapter alone: the book's identifier holds the chapter's length.)
    const chapterBytes = async (data: unknown, nodes: number, budget: number) => Buffer.byteLength(chapterOf((await OfficeGenerator.generate(chartAst(data, nodes, budget), 'epub', { onWarning: () => { } } as any)).value), 'utf8');
    const htmlBytes = async (data: unknown, nodes: number, budget: number) => bytesOf((await OfficeGenerator.generate(chartAst(data, nodes, budget), 'html', { htmlConfig: { chartJsSrc: '', standalone: false }, onWarning: () => { } } as any)).value);
    let cases = 0;
    for (const [shape, [make, nodes]] of Object.entries(classShapes)) {
        for (const [className, unit] of Object.entries(WEIGHT_CLASSES)) {
            const data = make(unit);
            const kept = await chapterBytes(data, nodes, G) - await chapterBytes(data, nodes, 0);
            const html = await htmlBytes(data, nodes, G) - await htmlBytes(data, nodes, 0);
            assert.ok(kept <= html, `${className} ${shape}: keeping the copies writes ${kept} bytes in EPUB, ${html} in HTML`);
            cases++;
        }
    }
    console.log(`  EPUB chart copies against HTML: ${cases} cases`);
}

// ─── Charts in CSV ───────────────────────────────────────────────────────────

/** `content` with every chart node taken out. */
const withoutCharts = (content: any[]): any[] => content.filter(node => node.type !== 'chart').map(node => node.children ? { ...node, children: withoutCharts(node.children) } : node);

/** CSV output as text: a string, or a zip's files by name. */
const csvText = (value: unknown): string => typeof value === 'string' ? value
    : Object.entries(unzipSync(new Uint8Array(value as any))).map(([name, data]) => `[${name}]\n${strFromU8(data)}`).join('\n');

/**
 * CSV writes sheets and tables, not charts: each chart it leaves out is reported, once a document, naming how many and
 * the first (its title, else its attachment's name), in a document of sheets and tables beside them and in one of
 * charts alone (the report that it has no table names them). What CSV writes is as before, byte for byte. Charts left
 * out by the caller (includeCharts: false) or with the sheets a range does not select are not reported. onNode false
 * for a chart, or for what holds it, leaves it out unreported; a string is written in a cell, and elsewhere the chart is
 * reported.
 */
export async function testChartsInCsvReported(): Promise<void> {
    const charted = (w: any) => w.code === 'CONTENT_NOT_REPRESENTABLE' && w.details?.reason === 'chart';
    const csv = async (ast: OfficeParserAST, config: Record<string, unknown> = {}) => {
        const heard: any[] = [];
        const out = await OfficeGenerator.generate(ast, 'csv', { ...config, onWarning: (w: any) => heard.push(w) } as any);
        assert.deepStrictEqual(out.messages.map(m => m.message), heard.map(w => w.message), 'CSV lists what it reports');
        return { text: csvText(out.value), heard };
    };
    for (const file of ['test.xlsx', 'test.pptx', 'test.ods', 'test.odp']) {
        for (const extractAttachments of [true, false]) {
            const ast = await OfficeParser.parseOffice(path.join(FILES, file), { extractAttachments, onWarning: () => { } } as any) as OfficeParserAST;
            const plain = { ...ast, content: withoutCharts(ast.content) } as OfficeParserAST;
            const label = `${file}${extractAttachments ? '' : ' (no chart data)'}`;
            // (An XLSX read without its attachments has no chart node: nothing to report.)
            const charts = allNodes(ast.content).filter(n => n.type === 'chart').length;
            assert.strictEqual(charts, extractAttachments || file !== 'test.xlsx' ? 1 : 0, `${label}: its chart nodes`);
            for (const config of [{}, { csvConfig: { mergeSheets: false } }, { csvConfig: { columnDelimiter: ';' } }]) {
                const { text, heard } = await csv(ast, config);
                const before = await csv(plain, config);
                assert.strictEqual(text, before.text, `${label} ${JSON.stringify(config)}: CSV writes what it wrote before, byte for byte`);
                assert.deepStrictEqual(before.heard.map(w => w.code), [], `${label}: without its chart, nothing is reported`);
                if (!charts) { assert.deepStrictEqual(heard, [], `${label}: no chart, no report`); continue; }
                assert.strictEqual(heard.length, 1, `${label} ${JSON.stringify(config)}: one report; got ${heard.map(w => w.message)}`);
                assert.ok(charted(heard[0]) && heard[0].details.count === 1, `${label}: it is the chart left out`);
                if (extractAttachments) assert.ok(heard[0].message.includes('a chart ("Chart Title")') && heard[0].details.name === 'Chart Title', `${label}: named by its title; got ${heard[0].message}`);
                else assert.ok(typeof heard[0].details.name === 'string' && heard[0].details.name && heard[0].message.includes(`"${heard[0].details.name}"`), `${label}: named by its attachment; got ${heard[0].message}`);
            }
            const dropped = await csv(ast, { includeCharts: false });
            assert.deepStrictEqual(dropped.heard, [], `${label}: charts left out by the caller (includeCharts: false) are not reported`);
            if (!extractAttachments) continue;
            // onNode false for its charts (an XLSX chart a sheet's child, an ODS chart in a cell, a slide's chart outside
            // the tables): left out unreported, the text as without them. A string: written in an ODS cell, unreported;
            // elsewhere CSV has no place for it, and the chart is reported, the text unchanged.
            const plainText = (await csv(plain)).text;
            const refused = await csv(ast, { onNode: (n: any) => n.type === 'chart' ? false : undefined });
            assert.deepStrictEqual(refused.heard, [], `${label}: a chart onNode leaves out is not reported; got ${refused.heard.map(w => w.message)}`);
            assert.strictEqual(refused.text, plainText, `${label}: and the text is as without the chart`);
            const replaced = await csv(ast, { onNode: (n: any) => n.type === 'chart' ? 'CHARTSTR' : undefined });
            if (file === 'test.ods') {
                assert.ok(replaced.text.includes('CHARTSTR'), `${label}: a string for a chart in a cell is written`);
                assert.deepStrictEqual(replaced.heard, [], `${label}: and the chart is not reported`);
            } else {
                assert.strictEqual(replaced.heard.filter(charted).length, 1, `${label}: a string for a chart CSV has no place for: the chart reported once; got ${replaced.heard.map(w => w.message)}`);
                assert.strictEqual(replaced.text, plainText, `${label}: and the text unchanged`);
            }
        }
    }

    // A document of charts alone: '' and the report that it has no table, naming them.
    for (const [file, count, name] of [['charts/word-charts.docx', 5, 'Monthly sales, a month blank'], ['charts/inline-chart.odt', 1, 'Revenue and margin']] as const) {
        const ast = await OfficeParser.parseOffice(path.join(FILES, file), { extractAttachments: true, onWarning: () => { } } as any) as OfficeParserAST;
        const { text, heard } = await csv(ast);
        assert.strictEqual(text, '', `${file}: CSV writes nothing`);
        assert.strictEqual(heard.length, 1, `${file}: one report`);
        const shown = count === 1 ? `a chart ("${name}")` : `${count} charts (the first "${name}")`;
        assert.ok(charted(heard[0]) && heard[0].details.count === count && heard[0].details.name === name && heard[0].message.includes(`a document with no spreadsheet or table, only ${shown}`), `${file}: it has no table and holds ${shown}; got ${heard[0].message}`);
        const dropped = await csv(ast, { includeCharts: false });
        assert.ok(dropped.heard.length === 1 && dropped.heard[0].details.feature === 'a document with no spreadsheet or table' && !charted(dropped.heard[0]), `${file}: with includeCharts false, that it has no table`);
    }

    // Where charts stand: outside the tables (before, between, after), in a cell, beside a table's rows (alone, or in a
    // drawing), in a row beside its cells. Counted and the first named in document order; a range counts those in what it
    // selects; onNode false for a chart, or for what holds it, leaves it out unreported; a string is written in a cell,
    // and elsewhere the chart is reported. onNode is asked of a node once at most, and not of a drawing beside the rows
    // that holds no chart.
    const chartData = (title: string) => ({ title, labels: ['Q1'], dataSets: [{ name: 'U', values: ['1'] }], rawTexts: [] });
    const chart = (name: string) => ({ type: 'chart', text: '', metadata: { attachmentName: name } });
    const cell = (children: any[]) => ({ type: 'cell', children });
    const text = (t: string) => ({ type: 'text', text: t });
    const table = (name: string, rows: any[], beside: any[] = []) => ({ type: 'table', metadata: { sheetName: name }, children: [...beside, ...rows] });
    const names = ['outside1', 'cell2', 'beside2', 'drawn2', 'row2', 'outside3', 'cell3'];
    const picture = { type: 'image', metadata: { attachmentName: 'pic' } };
    const pictureDrawing = { type: 'drawing', children: [picture] };
    const placed = {
        type: 'docx', metadata: {},
        attachments: names.map(name => ({ type: 'chart', name, mimeType: 'application/xml', extension: 'xml', data: '', chartData: chartData(`T ${name}`) })),
        content: [
            { type: 'paragraph', children: [text('before')] },
            table('A', [{ type: 'row', children: [cell([text('a1')])] }]),
            chart('outside1'),
            table('B', [{ type: 'row', children: [cell([text('b1')]), cell([chart('cell2')])] }, { type: 'row', children: [chart('row2'), cell([text('b2')])] }], [chart('beside2'), { type: 'drawing', children: [chart('drawn2')] }, pictureDrawing]),
            { type: 'paragraph', children: [chart('outside3')] },
            table('C', [{ type: 'row', children: [cell([{ type: 'paragraph', children: [chart('cell3')] }])] }]),
        ],
    } as any as OfficeParserAST;
    const plainPlaced = { ...placed, content: withoutCharts(placed.content) } as OfficeParserAST;
    const chartNamed = (chartName: string) => (n: any) => n.type === 'chart' && n.metadata?.attachmentName === chartName;
    const falseFor = (test: (n: any) => boolean) => (n: any) => test(n) ? false : undefined;
    const stringFor = (test: (n: any) => boolean) => (n: any) => test(n) ? 'CHARTSTR' : undefined;
    // What the CSV is: as the chart-free AST's ('plain'; with the same onNode where it drops only charts or what holds
    // them), holding the string ('string'), or not compared (onNode false for a row drops the row, which no chart-free
    // AST asks the same).
    const cases: [string, Record<string, unknown>, number, string | undefined, 'plain' | 'string' | 'other'][] = [
        ['every table', {}, 7, 'T outside1', 'plain'],
        ['tables 2 and 3', { csvConfig: { sheets: '2-3' } }, 5, 'T beside2', 'plain'],
        ['table 3', { csvConfig: { sheets: '3' } }, 1, 'T cell3', 'plain'],
        ['table 1', { csvConfig: { sheets: '1' } }, 0, undefined, 'plain'],
        ['onNode false for a chart', { onNode: falseFor(chartNamed('outside1')) }, 6, 'T beside2', 'plain'],
        ['onNode false for a row holding charts', { onNode: falseFor(n => n.type === 'row' && n.children?.some((c: any) => c.type === 'chart')) }, 6, 'T outside1', 'other'],
        ['onNode false for a chart beside the rows', { onNode: falseFor(chartNamed('beside2')) }, 6, 'T outside1', 'plain'],
        ['onNode false for a chart in a row beside its cells', { onNode: falseFor(chartNamed('row2')) }, 6, 'T outside1', 'plain'],
        ['onNode false for a chart in a drawing beside the rows', { onNode: falseFor(chartNamed('drawn2')) }, 6, 'T outside1', 'plain'],
        ['onNode false for the drawing holding a chart', { onNode: falseFor(n => n.type === 'drawing' && n.children?.some(chartNamed('drawn2'))) }, 6, 'T outside1', 'plain'],
        ['onNode a string for a chart in a cell', { onNode: stringFor(chartNamed('cell2')) }, 6, 'T outside1', 'string'],
        ['onNode a string for a chart outside the tables', { onNode: stringFor(chartNamed('outside1')) }, 7, 'T outside1', 'plain'],
        ['onNode a string for a chart beside the rows', { onNode: stringFor(chartNamed('beside2')) }, 7, 'T outside1', 'plain'],
        ['onNode a string for a chart in a row beside its cells', { onNode: stringFor(chartNamed('row2')) }, 7, 'T outside1', 'plain'],
        ['onNode a string for a chart in a drawing beside the rows', { onNode: stringFor(chartNamed('drawn2')) }, 7, 'T outside1', 'plain'],
    ];
    for (const [name, config, count, first, expected] of cases) {
        // Every node's askings counted.
        const asked = new Map<object, number>();
        const onNode = config.onNode as ((n: any) => unknown) | undefined;
        const { text: written, heard } = await csv(placed, { ...config, onNode: (n: any) => { asked.set(n, (asked.get(n) ?? 0) + 1); return onNode?.(n); } });
        for (const [node, times] of asked) assert.strictEqual(times, 1, `${name}: onNode asked of a ${(node as any).type} ${times} times`);
        assert.ok(!asked.has(pictureDrawing) && !asked.has(picture), `${name}: onNode not asked of a drawing beside the rows holding no chart, nor of its picture`);
        if (expected === 'plain') assert.strictEqual(written, (await csv(plainPlaced, config)).text, `${name}: CSV writes what it wrote before`);
        if (expected === 'string') assert.ok(written.includes('CHARTSTR'), `${name}: the string is written`);
        else assert.ok(!written.includes('CHARTSTR'), `${name}: no string is written`);
        const reports = heard.filter(charted);
        assert.strictEqual(reports.length, count ? 1 : 0, `${name}: ${count ? 'one report' : 'no report'}; got ${heard.map(w => w.message)}`);
        if (count) assert.ok(reports[0].details.count === count && reports[0].details.name === first, `${name}: ${count} charts, the first "${first}"; got ${reports[0].message}`);
    }

    // The name as a report shows it: a title of many lines on one, cut to 80 characters; no title, the attachment's name.
    const named = (title: unknown, attachmentName = 'chart9.xml') => ({
        type: 'docx', metadata: {},
        attachments: [{ type: 'chart', name: attachmentName, mimeType: 'application/xml', extension: 'xml', data: '', chartData: { ...chartData(''), title } }],
        content: [table('A', [{ type: 'row', children: [cell([text('a')])] }]), chart(attachmentName)],
    }) as any as OfficeParserAST;
    const long = (await csv(named(`line one\nline two ${'x'.repeat(200)}`))).heard[0];
    assert.ok(long.details.name.length <= 80 && long.details.name.startsWith('line one line two x') && !/\n/.test(long.message), `a long title of two lines is shown on one, cut; got ${JSON.stringify(long.details.name)}`);
    const untitled = (await csv(named(undefined))).heard[0];
    assert.ok(untitled.details.name === 'chart9.xml' && untitled.message.includes('a chart ("chart9.xml")'), `an untitled chart is named by its attachment; got ${untitled.message}`);

    // Linear time: n and 4n charts in a table's cells and beside it.
    const many = (n: number) => ({
        type: 'docx', metadata: {}, attachments: [],
        content: [table('A', Array.from({ length: n }, () => ({ type: 'row', children: [cell([text('a')]), cell([chart('c')])] })), Array.from({ length: n }, () => chart('c'))), ...Array.from({ length: n }, () => chart('c'))],
    }) as any as OfficeParserAST;
    const counted = (await csv(many(5000))).heard.filter(charted);
    assert.ok(counted.length === 1 && counted[0].details.count === 15000, `15,000 charts in one report; got ${counted.map(w => w.message)}`);
    await linearTime('CSV counting 15,000 charts', scale => OfficeGenerator.generate(many(5000 * scale), 'csv', { onWarning: () => { } } as any));
}

// ─── A chart without its data ────────────────────────────────────────────────

/** The writers that write a chart without its data (a parse without extractAttachments) as its text. */
const TEXT_WRITERS = ['html', 'epub', 'docx', 'odt', 'tex', 'md', 'rtf'] as const;

/** A document of `nodes` chart nodes of `text` and no attachment (a parse without extractAttachments), of the repeat budget given. */
const dataless = (text: string, nodes: number, maxRepeatedContent?: number): OfficeParserAST => ({
    type: 'docx', metadata: {}, attachments: [],
    ...(maxRepeatedContent !== undefined ? { config: { decompressionLimits: { maxRepeatedContent } } } : {}),
    content: Array.from({ length: nodes }, () => ({ type: 'chart', text, metadata: { attachmentName: 'c' } })),
}) as any as OfficeParserAST;

/** What `format` writes of `ast` (under `config`), read back as plain text, and the codes of its warnings. */
const readBackText = async (ast: OfficeParserAST, format: string, config: Record<string, unknown> = {}): Promise<{ text: string; warnings: string[]; value: unknown }> => {
    const warnings: string[] = [];
    const out = await OfficeGenerator.generate(ast, format as any, { ...config, onWarning: (w: any) => warnings.push(w.code) } as any);
    const back = await OfficeParser.parseOffice(Buffer.from(out.value as any), { fileType: format, onWarning: () => { } } as any) as OfficeParserAST;
    return { text: String((await back.to('text')).value), warnings, value: out.value };
};

/**
 * A chart without its data (a parse without extractAttachments gives a chart node its text and no data) is written as
 * its text by HTML, EPUB, DOCX, ODT, LaTeX, Markdown and RTF, as plain text writes it: HTML and EPUB wrote nothing for it,
 * DOCX, ODT and LaTeX only its name. Not reported: it is the text the parse gave of it. Written again at each chart node
 * of its name within the chart-copy budget, past it the chart's reference, reported once; a chart of no text is its
 * reference. So it is in plain text, chunks and native PDF (a DOCX or PPTX chart parsed without its attachments has no
 * text), which wrote nothing for it: in every place it stands, once each, none with includeCharts: false.
 */
export async function testChartsWithoutData(): Promise<void> {
    const X100 = 'x'.repeat(100);
    const text = `Q1\nQ2\nUnits\n10\n${X100}`;
    for (const format of TEXT_WRITERS) {
        const one = await readBackText(dataless(text, 1), format);
        assert.ok(occurrences(one.text, 'Q2') === 1 && occurrences(one.text, X100) === 1, `${format}: a chart without its data is its text; got ${JSON.stringify(one.text)}`);
        assert.deepStrictEqual(one.warnings, [], `${format}: not reported`);
        const past = await readBackText(dataless(text, 3, 0), format);
        assert.strictEqual(occurrences(past.text, X100), 1, `${format}: at a budget of 0 its text is written once`);
        assert.strictEqual(occurrences(past.text, '[Chart: c]'), 2, `${format}: and its reference at the other two`);
        assert.deepStrictEqual(past.warnings, ['REPEATED_CONTENT_LIMIT_EXCEEDED'], `${format}: reported once`);
        const roomy = await readBackText(dataless(text, 3), format);
        assert.strictEqual(occurrences(roomy.text, X100), 3, `${format}: within the budget its text at all three`);
        assert.deepStrictEqual(roomy.warnings, [], `${format}: within the budget nothing is reported`);
        const empty = await readBackText(dataless('', 1), format);
        assert.ok(empty.text.includes('[Chart: c]'), `${format}: a chart of no text is its reference; got ${JSON.stringify(empty.text)}`);
    }
    // Plain text writes it at each chart node, as before.
    const plain = String((await OfficeGenerator.generate(dataless(text, 3, 0), 'text', { onWarning: () => { } } as any)).value);
    assert.strictEqual(occurrences(plain, X100), 3, 'text: written at each chart node, as before');

    // A chart of no text standing in a line (every DOCX chart parsed without its attachments: Word places a chart in a
    // paragraph's run) is its reference in every place, as at the top: DOCX, ODT and LaTeX wrote nothing for it there.
    // None with includeCharts: false; nothing reported. (A footnote as readers give one: a text run's notes.)
    const lineRun = (t: string) => ({ type: 'text', text: t });
    const lineParagraph = (...children: any[]) => ({ type: 'paragraph', children });
    const inLine: Record<string, (chart: () => any) => any[]> = {
        'alone in a paragraph': chart => [lineParagraph(chart())],
        'in a paragraph of text': chart => [lineParagraph(lineRun('before '), chart(), lineRun(' after'))],
        'in a heading': chart => [{ type: 'heading', metadata: { level: 2 }, children: [lineRun('H '), chart()] }],
        'in a list item': chart => [{ type: 'list', metadata: { listType: 'unordered' }, children: [{ type: 'paragraph', metadata: { listId: 'l1', indentation: 0 }, children: [lineRun('item '), chart()] }] }],
        'in a cell': chart => [{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [lineParagraph(chart())] }] }] }],
        'in a footnote': chart => [lineParagraph({ type: 'text', text: 'x', notes: [{ type: 'note', metadata: { noteType: 'footnote', noteId: '1' }, children: [lineParagraph(chart())] }] })],
        'in a text box': chart => [{ type: 'drawing', children: [lineParagraph(lineRun('box '), chart())] }],
        'two in a paragraph': chart => [lineParagraph(chart(), lineRun(' mid '), chart())],
    };
    const inLineDocument = (place: string, chart: () => any, attachments: any[] = []) => ({ type: 'docx', metadata: {}, attachments, content: inLine[place](chart) }) as any as OfficeParserAST;
    const lineChart = (chartText: string, name = 'c') => () => ({ type: 'chart', text: chartText, metadata: { attachmentName: name } });
    for (const format of TEXT_WRITERS) {
        for (const place of Object.keys(inLine)) {
            const times = place === 'two in a paragraph' ? 2 : 1;
            const shown = await readBackText(inLineDocument(place, lineChart('')), format);
            assert.strictEqual(occurrences(shown.text, '[Chart: c]'), times, `${format}, a chart of no text ${place}: its reference ${times} time(s); got ${JSON.stringify(shown.text)}`);
            assert.deepStrictEqual(shown.warnings, [], `${format}, a chart of no text ${place}: not reported`);
            if (place === 'in a paragraph of text') {
                const at = ['before', '[Chart: c]', 'after'].map(part => shown.text.indexOf(part));
                assert.ok(at.every((where, i) => where >= 0 && (i === 0 || where > at[i - 1])), `${format}, ${place}: the paragraph's parts and the reference in order; got ${JSON.stringify(shown.text)}`);
            }
            const left = await readBackText(inLineDocument(place, lineChart('')), format, { includeCharts: false });
            assert.strictEqual(occurrences(left.text, '[Chart:'), 0, `${format}, a chart of no text ${place}, includeCharts: false: no reference; got ${JSON.stringify(left.text)}`);
            assert.deepStrictEqual(left.warnings, [], `${format}, a chart of no text ${place}, includeCharts: false: not reported`);
        }
    }
    // In a cell Markdown writes as HTML (a list beside the chart), the same reference; none with includeCharts: false.
    const cellList = { type: 'list', metadata: { listType: 'unordered' }, children: [{ type: 'paragraph', metadata: { listId: 'l1', indentation: 0 }, children: [lineRun('item')] }] };
    const htmlCell = { type: 'docx', metadata: {}, attachments: [], content: [{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [cellList, lineParagraph(lineChart('')())] }] }] }] } as any as OfficeParserAST;
    const htmlCellMd = String((await OfficeGenerator.generate(htmlCell, 'md', { onWarning: () => { } } as any)).value);
    assert.ok(/<td>[\s\S]*<ul>[\s\S]*<\/ul>[\s\S]*\[Chart: c\][\s\S]*<\/td>/.test(htmlCellMd) && occurrences(htmlCellMd, '[Chart: c]') === 1, `md: a chart of no text in a cell written as HTML is its reference once; got ${JSON.stringify(htmlCellMd)}`);
    const htmlCellLeft = String((await OfficeGenerator.generate(htmlCell, 'md', { includeCharts: false, onWarning: () => { } } as any)).value);
    assert.ok(htmlCellLeft.includes('item') && !htmlCellLeft.includes('[Chart:'), `md: includeCharts: false, the cell without a reference; got ${JSON.stringify(htmlCellLeft)}`);
    // A chart with data and no text in a line (a built AST): its reference by its title, as plain text writes it.
    const titled = [{ type: 'chart', name: 's', mimeType: 'application/xml', extension: 'xml', data: '', chartData: { title: 'Sales', labels: ['Q1'], dataSets: [{ name: 'Units', values: ['10'] }], rawTexts: [] } }];
    for (const format of LINE_WRITERS) {
        const shown = await readBackText(inLineDocument('in a paragraph of text', lineChart('', 's'), titled), format);
        assert.strictEqual(occurrences(shown.text, '[Chart: Sales]'), 1, `${format}: a chart with data and no text in a paragraph is its reference by its title; got ${JSON.stringify(shown.text)}`);
    }
    // A chart of text in a line: its text, as before, never its reference.
    for (const format of LINE_WRITERS) {
        for (const place of Object.keys(inLine)) {
            const times = place === 'two in a paragraph' ? 2 : 1;
            const shown = await readBackText(inLineDocument(place, lineChart('Q1 5')), format);
            assert.ok(occurrences(shown.text, 'Q1 5') === times && !shown.text.includes('[Chart:'), `${format}, a chart of text ${place}: its text ${times} time(s), no reference; got ${JSON.stringify(shown.text)}`);
        }
    }
    // Charged as any first write and copy: n paragraphs each holding a chart of no text naming one chart, a reference a
    // node within one node's output, the budget and 600 bytes a node, in linear time; at a budget of 0 reported once,
    // as the same charts standing alone are.
    const lineCharts = (n: number, budget?: number) => ({
        type: 'docx', metadata: {}, attachments: [],
        ...(budget !== undefined ? { config: { decompressionLimits: { maxRepeatedContent: budget } } } : {}),
        content: Array.from({ length: n }, () => lineParagraph(lineRun('a '), lineChart('')())),
    }) as any as OfficeParserAST;
    for (const format of LINE_WRITERS) {
        for (const budget of [0, undefined]) {
            const first = (await writtenBytes(lineCharts(1, budget), format)).bytes;
            for (const n of [2000, 8000]) {
                const ast = lineCharts(n, budget);
                const { bytes, text: written, warnings } = await writtenBytes(ast, format);
                assert.strictEqual(occurrences(written, 'Chart: c'), n, `${format}: ${n} charts of no text in a line (budget ${budget ?? 'default'}), a reference each`);
                assert.ok(bytes <= first + repeatedContentOfAst(ast) + 600 * n, `${format}: ${n} charts of no text in a line (budget ${budget ?? 'default'}) write ${bytes} bytes, one ${first}`);
                assert.deepStrictEqual(warnings, budget === 0 ? ['REPEATED_CONTENT_LIMIT_EXCEEDED'] : [], `${format}: ${n} charts of no text in a line (budget ${budget ?? 'default'}): ${budget === 0 ? 'reported once' : 'nothing reported'}`);
                assert.deepStrictEqual(warnings, (await writtenBytes(dataless('', n, budget), format)).warnings, `${format}: ${n} charts of no text in a line reported as the same charts standing alone`);
            }
            await linearTime(`${format}: charts of no text in a line (budget ${budget ?? 'default'})`, scale => writtenBytes(lineCharts(2000 * scale, budget), format));
        }
    }

    // In a paragraph: HTML and EPUB write the paragraph's parts and the chart's text between them, the chapter XML.
    const inParagraph = {
        type: 'docx', metadata: {}, attachments: [],
        content: [{ type: 'paragraph', children: [{ type: 'text', text: 'para ' }, { type: 'chart', text: 'Q1\nQ2', metadata: { attachmentName: 'c' } }, { type: 'text', text: ' after' }] }],
    } as any as OfficeParserAST;
    for (const format of ['html', 'epub'] as const) {
        const { text: back, value } = await readBackText(inParagraph, format);
        const at = ['para', 'Q1', 'Q2', 'after'].map(part => back.indexOf(part));
        assert.ok(at.every((place, i) => place >= 0 && (i === 0 || place > at[i - 1])), `${format}: the paragraph's parts and the chart's text in order; got ${JSON.stringify(back)}`);
        if (format === 'epub') assert.doesNotThrow(() => parseXmlString(chapterOf(value)), 'epub: the chapter is well-formed XML');
    }

    // n and 4n chart nodes of a 2,000-line text (a repeat budget of 256 KiB): within one node's output, the budget and a
    // reference a node, in linear time.
    const long = Array.from({ length: 2000 }, (_, i) => `line ${i}`).join('\n');
    for (const format of TEXT_WRITERS) {
        const first = (await writtenBytes(dataless(long, 1, 256 * 1024), format)).bytes;
        for (const n of [200, 800]) {
            const ast = dataless(long, n, 256 * 1024);
            const { bytes, warnings } = await writtenBytes(ast, format);
            assert.ok(bytes <= first + repeatedContentOfAst(ast) + 600 * n, `${format}: ${n} chart nodes without their data write ${bytes} bytes, one ${first}`);
            assert.ok(warnings.includes('REPEATED_CONTENT_LIMIT_EXCEEDED'), `${format}: ${n} chart nodes spend the budget, reported`);
        }
        await linearTime(`${format}: chart nodes without their data`, scale => writtenBytes(dataless(long, 200 * scale, 256 * 1024), format));
    }

    // A copy is charged what HTML and EPUB each write of it: a line break is `<br/>` in both (EPUB self-closes void
    // tags, so `<br>` charged 4 bytes was written as 5). Blank lines, a digit a line, CRLF: the copies EPUB keeps within
    // a budget of 1 GiB, past those at 0, are the bytes HTML writes for them.
    const GIB = 1024 * 1024 * 1024;
    const epubChapterBytes = async (ast: OfficeParserAST) => Buffer.byteLength(chapterOf((await OfficeGenerator.generate(ast, 'epub', { onWarning: () => { } } as any)).value), 'utf8');
    const htmlBytes = async (ast: OfficeParserAST) => Buffer.byteLength(String((await OfficeGenerator.generate(ast, 'html', { htmlConfig: { standalone: false }, onWarning: () => { } } as any)).value), 'utf8');
    for (const [name, lines] of [['blank lines', 'a' + '\n'.repeat(2000) + 'b'], ['a digit a line', Array.from({ length: 2000 }, (_, i) => String(i % 10)).join('\n')], ['CRLF', 'a' + '\r\n'.repeat(2000) + 'b']]) {
        const epubKept = await epubChapterBytes(dataless(lines, 10, GIB)) - await epubChapterBytes(dataless(lines, 10, 0));
        const htmlKept = await htmlBytes(dataless(lines, 10, GIB)) - await htmlBytes(dataless(lines, 10, 0));
        assert.ok(htmlKept > 0, `${name}: copies are kept within the budget`);
        assert.strictEqual(epubKept, htmlKept, `${name}: EPUB writes the copies it keeps as HTML does, what each is charged`);
    }
    // Short lines (one byte a line): n chart nodes within one node's output, the budget and a reference a node.
    const short = 'a' + '\n'.repeat(1000) + 'b';
    for (const format of TEXT_WRITERS) {
        const first = (await writtenBytes(dataless(short, 1, 1024 * 1024), format)).bytes;
        for (const n of [300, 1200]) {
            const ast = dataless(short, n, 1024 * 1024);
            const { bytes } = await writtenBytes(ast, format);
            assert.ok(bytes <= first + repeatedContentOfAst(ast) + 600 * n, `${format}: ${n} chart nodes of 1,000 short lines write ${bytes} bytes, one ${first}`);
        }
    }

    // Plain text, chunks and native PDF: a chart of no text is its reference, in every place it stands, as the flow
    // writers write it; none with includeCharts: false.
    const noText = (name = 'c') => ({ type: 'chart', text: '', metadata: { attachmentName: name } });
    const run = (t: string) => ({ type: 'text', text: t });
    const places: Record<string, (chart: any) => any[]> = {
        'at the top': chart => [chart],
        'in a paragraph': chart => [{ type: 'paragraph', children: [run('before '), chart, run(' after')] }],
        'in a heading': chart => [{ type: 'heading', metadata: { level: 2 }, children: [run('H '), chart] }],
        'in a list item': chart => [{ type: 'paragraph', metadata: { listId: 'l1', indentation: 0 }, children: [run('item '), chart] }],
        'in a cell': chart => [{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [{ type: 'paragraph', children: [chart] }] }] }] }],
        'in a footnote': chart => [{ type: 'paragraph', children: [run('x'), { type: 'note', metadata: { noteType: 'footnote', noteId: '1' }, children: [{ type: 'paragraph', children: [chart] }] }] }],
        'on a slide': chart => [{ type: 'slide', metadata: { slideNumber: 1 }, children: [chart] }],
    };
    const documentOf = (place: string, chart: any, attachments: any[] = []) => ({ type: place === 'on a slide' ? 'pptx' : 'docx', metadata: {}, attachments, content: places[place](chart) }) as any as OfficeParserAST;
    const STRATEGIES = ['document-structure', 'fixed-size', 'semantic'] as const;
    const embeddingFunction = async () => [1, 0];
    const chunkTexts = async (ast: OfficeParserAST, strategy: string, config: Record<string, unknown> = {}) => ((await OfficeGenerator.generate(ast, 'chunks', { ...config, chunksConfig: { strategy, embeddingFunction }, onWarning: () => { } } as any)).value as any[]).map(chunk => String(chunk.text)).join('\n');
    const plainText = async (ast: OfficeParserAST, config: Record<string, unknown> = {}) => String((await OfficeGenerator.generate(ast, 'text', { ...config, onWarning: () => { } } as any)).value);
    for (const place of Object.keys(places)) {
        const ast = documentOf(place, noText());
        for (const includeCharts of [true, false]) {
            const expected = includeCharts ? 1 : 0;
            const config = includeCharts ? {} : { includeCharts: false };
            assert.strictEqual(occurrences(await plainText(ast, config), '[Chart: c]'), expected, `text, a chart of no text ${place}${includeCharts ? '' : ', includeCharts: false'}: its reference ${expected} times`);
            for (const strategy of STRATEGIES) assert.strictEqual(occurrences(await chunkTexts(ast, strategy, config), '[Chart: c]'), expected, `chunks (${strategy}), a chart of no text ${place}${includeCharts ? '' : ', includeCharts: false'}: its reference ${expected} times`);
            const printed = await nativePdf(ast, config);
            assert.strictEqual(occurrences(printed.text, '[Chart: c]'), expected, `native PDF, a chart of no text ${place}${includeCharts ? '' : ', includeCharts: false'}: its reference ${expected} times; got ${JSON.stringify(printed.text)}`);
            const charted = printed.warnings.filter(w => w.code === 'CONTENT_NOT_REPRESENTABLE');
            if (includeCharts) assert.ok(charted.length === 1 && ['chart', 'chart-in-text-line'].includes(charted[0].details?.reason), `native PDF, ${place}: written as text, reported once; got ${charted.map(w => w.details?.reason)}`);
            else assert.deepStrictEqual(charted, [], `native PDF, ${place}, includeCharts: false: nothing reported`);
        }
    }
    // A chart with data and no text (a built AST): its reference by its title. (Native PDF draws a chart it can from its
    // data: one in a heading, and one whose values it cannot draw, are written as text.)
    const sales = [{ type: 'chart', name: 's', mimeType: 'application/xml', extension: 'xml', data: '', chartData: { title: 'Sales', labels: ['Q1'], dataSets: [{ name: 'Units', values: ['10'] }], rawTexts: [] } }];
    const undrawn = [{ ...sales[0], chartData: { ...sales[0].chartData, dataSets: [{ name: 'Units', values: ['n/a'] }] } }];
    assert.strictEqual(occurrences(await plainText(documentOf('at the top', noText('s'), sales)), '[Chart: Sales]'), 1, 'text: a chart with data and no text is its reference by its title');
    for (const strategy of STRATEGIES) assert.strictEqual(occurrences(await chunkTexts(documentOf('at the top', noText('s'), sales), strategy), '[Chart: Sales]'), 1, `chunks (${strategy}): a chart with data and no text is its reference by its title`);
    assert.strictEqual(occurrences((await nativePdf(documentOf('in a heading', noText('s'), sales))).text, '[Chart: Sales]'), 1, 'native PDF: a chart with data and no text in a heading is its reference by its title');
    assert.strictEqual(occurrences((await nativePdf(documentOf('at the top', noText('s'), undrawn))).text, '[Chart: Sales]'), 1, 'native PDF: a chart with data it cannot draw and no text is its reference by its title');
    assert.ok(!(await nativePdf(documentOf('at the top', noText('s'), sales))).text.includes('[Chart:'), 'native PDF: a chart drawn from its data is drawn, as before');

    // Parsed without attachments: a DOCX or PPTX chart has no text, and is its reference in text, chunks and native PDF.
    for (const [file, count] of [['charts/word-charts.docx', 5], ['test.pptx', 1]] as const) {
        const ast = await OfficeParser.parseOffice(path.join(FILES, file), { onWarning: () => { } } as any) as OfficeParserAST;
        const charts = allNodes(ast.content).filter(n => n.type === 'chart');
        assert.ok(charts.length === count && charts.every(n => !(n.text ?? '').trim()), `${file}: ${count} chart nodes of no text without attachments`);
        const references = charts.map(n => `[Chart: ${n.metadata.attachmentName}]`);
        for (const [writer, written] of [['text', await plainText(ast)], ['chunks', await chunkTexts(ast, 'document-structure')], ['native PDF', (await nativePdf(ast)).text]] as const) {
            assert.strictEqual(occurrences(written, '[Chart: '), count, `${file}, ${writer}: ${count} references`);
            for (const reference of new Set(references)) assert.strictEqual(occurrences(written, reference), references.filter(r => r === reference).length, `${file}, ${writer}: ${reference}, the chart's attachment's name`);
        }
        // Word places each chart in a paragraph's run: DOCX, ODT and LaTeX write its reference in its line (read back).
        if (file.endsWith('.docx')) {
            for (const format of LINE_WRITERS) {
                const { text: back, warnings } = await readBackText(ast, format);
                assert.strictEqual(occurrences(back, '[Chart: '), count, `${file}, ${format}: ${count} references; got ${JSON.stringify(back)}`);
                for (let n = 1; n <= count; n++) assert.strictEqual(occurrences(back, `[Chart: chart${n}.xml]`), 1, `${file}, ${format}: [Chart: chart${n}.xml] once`);
                assert.deepStrictEqual(warnings, [], `${file}, ${format}: nothing reported`);
            }
        }
    }

    // n chart nodes of no text naming one chart: native PDF within one node's output, the budget and a reference a node
    // (at a budget of 0 and the default); plain text and chunks a reference a node; in linear time.
    const pdfBytes = async (ast: OfficeParserAST) => ((await OfficeGenerator.generate(ast, 'pdf', { pdfConfig: { engine: 'native' }, onWarning: () => { } } as any)).value as Uint8Array).length;
    for (const budget of [0, undefined]) {
        const first = await pdfBytes(dataless('', 1, budget));
        for (const n of [200, 800]) {
            const ast = dataless('', n, budget);
            const bytes = await pdfBytes(ast);
            assert.ok(bytes <= first + repeatedContentOfAst(ast) + 600 * n, `native PDF: ${n} chart nodes of no text (budget ${budget ?? 'default'}) write ${bytes} bytes, one ${first}`);
        }
        await linearTime(`native PDF: chart nodes of no text (budget ${budget ?? 'default'})`, scale => pdfBytes(dataless('', 200 * scale, budget)));
    }
    const perReference = Buffer.byteLength('[Chart: c]', 'utf8') + 2;
    const flowBytes: Record<string, (ast: OfficeParserAST) => Promise<number>> = {
        text: async ast => Buffer.byteLength(await plainText(ast), 'utf8'),
        chunks: async ast => Buffer.byteLength(await chunkTexts(ast, 'document-structure'), 'utf8'),
    };
    for (const [writer, bytesOf] of Object.entries(flowBytes)) {
        const first = await bytesOf(dataless('', 1));
        for (const n of [2000, 8000]) {
            const bytes = await bytesOf(dataless('', n));
            assert.ok(bytes <= n * perReference + first, `${writer}: ${n} chart nodes of no text write ${bytes} bytes, one ${first}`);
        }
        await linearTime(`${writer}: chart nodes of no text`, scale => bytesOf(dataless('', 2000 * scale)));
    }
}

/**
 * Document shapes kept across readers and writers: XHTML read as HTML, list types per level, note
 * marks, quotes of several paragraphs, figure captions, disclosures and keyboard input, and sheets
 * written as document tables.
 */
import * as assert from 'assert';
import * as fs from 'fs';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParser } from '../../src/OfficeParser';
import type { ImageMetadata, NoteMetadata, OfficeContentNode, OfficeParserAST, ParagraphMetadata } from '../../src/types';
import { MARKDOWN_FINGERPRINTS_FILE, markdownFingerprints, markdownWithoutDisclosures } from '../helpers/markdownFingerprints';

/** An XML document whose root is `html` (XHTML, as office suites export HTML) is read as HTML. */
export async function testXhtmlReadAsHtml(): Promise<void> {
    const xhtml = '<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">\n'
        + '<html xmlns="http://www.w3.org/1999/xhtml"><head><title>t</title></head><body><p>Hello</p></body></html>';
    for (const config of [{ fileType: 'html' as const }, {}]) {
        const ast = await OfficeParser.parseOffice(Buffer.from(xhtml), config);
        assert.strictEqual(ast.type, 'html', 'XHTML: read as HTML');
        assert.deepStrictEqual(ast.warnings.map(w => w.code), [], `XHTML: no warning (${JSON.stringify(config)})`);
        assert.strictEqual((await ast.to('text')).value.trim(), 'Hello', 'XHTML: its text');
    }
    // A prefixed root, a comment and a DOCTYPE with an internal subset holding `>`.
    const prefixed = '﻿<?xml version="1.0"?>\n<!-- c --><!DOCTYPE x [<!ENTITY a ">">]>'
        + '<x:html xmlns:x="http://www.w3.org/1999/xhtml"><x:body><x:p>Hi</x:p></x:body></x:html>';
    const prefixedAst = await OfficeParser.parseOffice(Buffer.from(prefixed));
    assert.strictEqual(prefixedAst.type, 'html', 'XHTML: a prefixed root is HTML');
    assert.strictEqual((await prefixedAst.to('text')).value.trim(), 'Hi', 'XHTML: the internal subset is no text');
    // Other XML is unchanged: refused with no type, a mismatch when called HTML.
    const other = Buffer.from('<?xml version="1.0"?><root><html/></root>');
    await assert.rejects(OfficeParser.parseOffice(other), 'XML: other XML is still not a supported type');
    const otherAsHtml = await OfficeParser.parseOffice(other, { fileType: 'html' });
    assert.ok(otherAsHtml.warnings.some(w => w.code === 'BUFFER_TYPE_MISMATCH'), 'XML: other XML called HTML still warns');
    console.log('✓ XHTML read as HTML');
}

/** An RTF list is typed per level by its definition (`\levelnfc` of the `\listlevel` its `\ilvl` names). */
export async function testRtfListTypesPerLevel(): Promise<void> {
    const rtf = String.raw`{\rtf1\ansi{\fonttbl{\f0 Times;}}`
        + String.raw`{\*\listtable{\list\listtemplateid2`
        + String.raw`{\listlevel\levelnfc23\leveljc0\levelstartat1{\leveltext \'01舦 ?;}{\levelnumbers;}\fi-360\li720}`
        + String.raw`{\listlevel\levelnfc0\leveljc0\levelstartat1{\leveltext \'02\'01.;}{\levelnumbers\'01;}\fi-360\li1080}`
        + String.raw`{\listlevel\levelnfc22\leveljc0\levelstartat1{\leveltext \'02\'02.;}{\levelnumbers\'01;}\fi-360\li1440}\listid2}`
        + String.raw`{\list\listtemplateid3{\listlevel\levelnfc0\leveljc0\levelstartat1{\leveltext \'02\'00.;}{\levelnumbers\'01;}\fi-360\li720}`
        + String.raw`{\listlevel\levelnfc23\leveljc0\levelstartat1{\leveltext \'01舦 ?;}{\levelnumbers;}\fi-360\li1080}\listid3}}`
        + String.raw`{\listoverridetable{\listoverride\listid2\listoverridecount0\ls2}{\listoverride\listid3\listoverridecount0\ls3}}`
        + String.raw`\pard{\listtext\pard\plain 舦\'95\tab}\ilvl0\ls2 \li720\fi-360 North\par`
        + String.raw`\pard{\listtext\pard\plain 1.\tab}\ls2\ilvl1 \li1080\fi-360 Detail\par`
        + String.raw`\pard{\listtext\pard\plain 01.\tab}\ls2\ilvl2 \li1440\fi-360 Deeper\par`
        + String.raw`\pard{\listtext\pard\plain 舦\'95\tab}\ilvl0\ls2 \li720\fi-360 South\par`
        + String.raw`\pard Between\par`
        + String.raw`\pard{\listtext\pard\plain 1.\tab}\ilvl0\ls3 \li720\fi-360 Hire\par`
        + String.raw`\pard{\listtext\pard\plain 舦\'95\tab}\ilvl1\ls3 \li1080\fi-360 Two\par}`;
    const ast = await OfficeParser.parseOffice(Buffer.from(rtf), { fileType: 'rtf' });
    const items = ast.content.filter(n => n.type === 'list').map(n => [n.text, (n.metadata as any).listType, (n.metadata as any).indentation]);
    assert.deepStrictEqual(items, [
        ['North', 'unordered', 0], ['Detail', 'ordered', 1], ['Deeper', 'ordered', 2], ['South', 'unordered', 0],
        ['Hire', 'ordered', 0], ['Two', 'unordered', 1],
    ], 'RTF: each level typed by its own definition');
    const html = (await ast.to('html', { htmlConfig: { standalone: false } })).value;
    assert.ok(/<ul>\s*<li[^>]*>\s*North/.test(html), `RTF: the bullet list writes as <ul>: ${html}`);
    assert.ok(/<ol>\s*<li[^>]*>\s*Hire/.test(html), 'RTF: the numbered list writes as <ol>');
    console.log('✓ RTF list types per level');
}

/** The generate options of an editor-bound HTML fragment and of a Markdown save, as a host round-tripping documents sets them. */
const EDITOR_HTML = { maxInlineImageBytes: Infinity, htmlConfig: { sourceAttributes: true, standalone: { document: false, styles: 'none' as const }, gatedEmbeds: true, omitDefaultTextColor: true } };
const SAVED_MD = { generateIds: false, maxInlineImageBytes: Infinity, mdConfig: { dialect: { extends: 'extended' as const, embeds: 'directive' as const }, fallbackToHtml: { inlineFormatting: true } } };

/** Markdown to editor HTML and back to Markdown, `passes` times; each pass's HTML and Markdown. */
async function markdownThroughHtml(markdown: string, passes = 3): Promise<{ html: string[]; md: string[] }> {
    const html: string[] = [];
    const md: string[] = [];
    let current = markdown;
    for (let pass = 0; pass < passes; pass++) {
        const ast = await OfficeParser.parseOffice(Buffer.from(current), { fileType: 'md', extractAttachments: true, includeBreakNodes: true, ocr: false, htmlParserConfig: { preserveIframes: true } });
        html.push((await ast.to('html', EDITOR_HTML)).value);
        const back = await OfficeParser.parseOffice(Buffer.from(html[pass]), { fileType: 'html' });
        current = (await back.to('md', SAVED_MD)).value;
        md.push(current);
    }
    return { html, md };
}

/** A quote of several paragraphs is one quote: one <blockquote> of paragraphs in HTML, `>` lines between them in Markdown. */
export async function testQuoteParagraphs(): Promise<void> {
    const source = 'Before\n\n> a\n>\n> b *c*\n\n> d\n\nAfter';
    const { html, md } = await markdownThroughHtml(source);
    assert.ok(html[0].includes('<blockquote><p>a</p><p>b <i>c</i></p></blockquote>'), `Quote: one <blockquote> of two paragraphs: ${html[0]}`);
    assert.ok(html[0].includes('<blockquote>d</blockquote>'), 'Quote: a quote of one paragraph is written as before');
    assert.deepStrictEqual(md, [source, source, source], 'Quote: Markdown to HTML to Markdown is stable over three passes');
    const ast = await OfficeParser.parseOffice(Buffer.from(source), { fileType: 'md' });
    const quotes = ast.content.filter(n => (n.metadata as ParagraphMetadata | undefined)?.style === 'Quote').map(n => (n.metadata as ParagraphMetadata).quoteId);
    assert.ok(quotes.length === 3 && quotes[0] === quotes[1] && quotes[1] !== quotes[2], `Quote: the paragraphs of one quote share an id: ${quotes}`);
    // The HTML reader reads one <blockquote> of paragraphs as one quote; two quotes stay two.
    const fromHtml = await OfficeParser.parseOffice(Buffer.from('<blockquote><p>a</p><p>b</p></blockquote><blockquote><p>c</p></blockquote>'), { fileType: 'html' });
    assert.strictEqual((await fromHtml.to('md')).value, '> a\n>\n> b\n\n> c', 'Quote: HTML quotes to Markdown');
    // LaTeX: one quote environment, read back as one quote.
    const tex = String((await ast.to('tex')).value);
    assert.ok(tex.includes('\\begin{quote}\na\n\nb \\textit{c}\n\\end{quote}'), `Quote: one quote environment: ${tex}`);
    const fromTex = await OfficeParser.parseOffice(Buffer.from(tex), { fileType: 'tex' });
    assert.strictEqual((await fromTex.to('md')).value, source, 'Quote: LaTeX reads a quote environment of paragraphs as one quote');
    // EPUB (XHTML) keeps it too.
    const fromEpub = await OfficeParser.parseOffice(Buffer.from((await ast.to('epub')).value as Uint8Array));
    assert.ok(((await fromEpub.to('md')).value).endsWith(source), 'Quote: EPUB keeps one quote of two paragraphs');
    console.log('✓ Quote paragraphs');
}

/** A one-pixel PNG, as base64. */
const PIXEL_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

/**
 * A `<figure>`'s caption is its picture's (ImageMetadata.caption): a figure again in HTML, EPUB and
 * Markdown, a caption paragraph after the picture in every other format, never a stray paragraph.
 */
export async function testFigureCaptions(): Promise<void> {
    const figure = `<p>Before</p><figure><img src="data:image/png;base64,${PIXEL_PNG}" alt="A map"><figcaption>Survey  area &amp; map</figcaption></figure><p>After</p>`;
    const ast = await OfficeParser.parseOffice(Buffer.from(figure), { fileType: 'html', extractAttachments: true });
    assert.deepStrictEqual(ast.content.map(n => n.type), ['paragraph', 'image', 'paragraph'], 'Figure: a picture, its caption no paragraph of its own');
    assert.strictEqual((ast.content[1].metadata as ImageMetadata).caption, 'Survey area & map', 'Figure: the caption is the picture\'s');
    const html = (await ast.to('html', { htmlConfig: { standalone: false } })).value;
    assert.ok(/<figure><img [^>]*alt="A map"[^>]*><figcaption>Survey area &amp; map<\/figcaption><\/figure>/.test(html), `Figure: HTML writes a figure: ${html}`);
    // Markdown: an HTML <figure> block, read back as the same captioned picture, stable through an editor's HTML.
    const md = (await ast.to('md')).value;
    assert.ok(/\n\n<figure><img [^>]*><figcaption>Survey area &amp; map<\/figcaption><\/figure>\n\nAfter/.test(md), `Figure: Markdown writes an HTML figure block: ${md}`);
    const { md: passes } = await markdownThroughHtml(md);
    assert.deepStrictEqual(passes, [md, md, md], 'Figure: Markdown to HTML to Markdown is stable over three passes');
    // EPUB keeps the figure.
    const epub = await OfficeParser.parseOffice(Buffer.from((await ast.to('epub')).value as Uint8Array), { extractAttachments: true });
    assert.strictEqual((epub.content.find(n => n.type === 'image')?.metadata as ImageMetadata | undefined)?.caption, 'Survey area & map', 'Figure: EPUB keeps the caption');
    // DOCX and ODT: a caption paragraph in the format's own caption style, after the picture.
    const docx = new TextDecoder().decode(unzipSync((await ast.to('docx')).value as Uint8Array)['word/document.xml']);
    assert.ok(/<w:drawing>[\s\S]*<\/w:p><w:p><w:pPr><w:pStyle w:val="Caption"\/><\/w:pPr><w:r><w:t xml:space="preserve">Survey area &amp; map<\/w:t>/.test(docx), 'Figure: DOCX writes a Caption paragraph after the picture');
    const odt = new TextDecoder().decode(unzipSync((await ast.to('odt')).value as Uint8Array)['content.xml']);
    assert.ok(/<text:p text:style-name="Caption">Survey area &amp; map<\/text:p>/.test(odt), 'Figure: ODT writes a Caption paragraph');
    const tex = String((await ast.to('tex')).value);
    assert.ok(tex.includes('\\captionof{figure}{Survey area \\& map}'), 'Figure: LaTeX numbers the caption as the figure\'s');
    for (const format of ['rtf', 'docx', 'odt', 'pptx'] as const) {
        const back = await OfficeParser.parseOffice(Buffer.from((await ast.to(format)).value as unknown as Uint8Array));
        const text = (await back.to('text')).value;
        assert.ok(/Before[\s\S]*Survey area & map[\s\S]*After/.test(text), `Figure: ${format} keeps the caption after the picture: ${text}`);
    }
    assert.ok(/\[Image: A map\]\nSurvey area & map\nAfter/.test((await ast.to('text')).value), 'Figure: text writes the caption after the picture');
    const pdf = await OfficeParser.parseOffice(Buffer.from((await ast.to('pdf', { pdfConfig: { engine: 'native' } })).value as Uint8Array), { fileType: 'pdf', ocr: false, pdfParserConfig: { separateProcess: false } });
    assert.ok(((await pdf.to('text')).value).includes('Survey area & map'), 'Figure: the native PDF engine draws the caption');
    // In a list item, where no figure goes, the caption follows the picture, reported.
    const inItem = await OfficeParser.parseOffice(Buffer.from(`<ul><li><figure><img src="data:image/png;base64,${PIXEL_PNG}" alt="Dot"><figcaption>The dot</figcaption></figure></li></ul>`), { fileType: 'html', extractAttachments: true });
    const itemMd = await inItem.to('md');
    assert.ok(/^- !\[Dot\]\([^)]*\) The dot$/m.test(itemMd.value), `Figure: in a list item the caption follows the picture: ${itemMd.value}`);
    assert.ok(itemMd.messages.some(m => m.code === 'CONTENT_NOT_REPRESENTABLE'), 'Figure: and that is reported');
    // A caption of more than text stays a paragraph after the picture, nothing of it lost.
    const rich = await OfficeParser.parseOffice(Buffer.from(`<figure><img src="data:image/png;base64,${PIXEL_PNG}" alt="A"><figcaption>See <a href="https://example.com">the source</a></figcaption></figure>`), { fileType: 'html', extractAttachments: true });
    assert.deepStrictEqual(rich.content.map(n => n.type), ['image', 'paragraph'], 'Figure: a rich caption stays a paragraph');
    console.log('✓ Figure captions');
}

/**
 * `<details>`/`<summary>` is a disclosure node, written back as `<details>` in HTML, EPUB and Markdown
 * (an HTML block holding Markdown) and as a bold summary line and its body elsewhere; `<kbd>` is keyboard
 * input, `<kbd>` in HTML and Markdown and code elsewhere.
 */
export async function testDisclosuresAndKeyboard(): Promise<void> {
    const source = 'Before\n\n<details open>\n<summary>Click me</summary>\n\nSome *body* text.\n\n- a\n- b\n\n```js\nconst x = 1;\n\nconst y = 2;\n```\n\n'
        + '<details>\n<summary>Inner</summary>\n\nNested.\n\n</details>\n\n</details>\n\nPress <kbd>Ctrl</kbd>+<kbd>*C*</kbd> to copy.\n\nAfter';
    const ast = await OfficeParser.parseOffice(Buffer.from(source), { fileType: 'md' });
    const outer = ast.content[1];
    assert.strictEqual(outer.type, 'disclosure', 'Disclosure: a details block is a disclosure');
    assert.deepStrictEqual(outer.metadata, { summary: 'Click me', open: true }, 'Disclosure: its summary and open state');
    assert.deepStrictEqual(outer.children!.map(c => c.type), ['paragraph', 'list', 'list', 'code', 'disclosure'], 'Disclosure: its body read as Markdown, a nested one too');
    const keys = ast.content[2].children!.filter(c => c.formatting?.keyboard).map(c => [c.text, c.formatting?.font]);
    assert.deepStrictEqual(keys, [['Ctrl', 'monospace'], ['*C*', 'monospace']], 'Keyboard: <kbd> is keyboard input, monospace too');
    // Markdown to editor HTML and back, three times: the same Markdown, kbd's Markdown characters as references.
    const expected = source.replace('<kbd>*C*</kbd>', '<kbd>&#42;C&#42;</kbd>');
    const { html, md } = await markdownThroughHtml(source);
    assert.deepStrictEqual(md, [expected, expected, expected], 'Disclosure: Markdown to HTML to Markdown is stable over three passes');
    assert.ok(html[0].includes('<details open=""><summary>Click me</summary><p>Some <i>body</i> text.</p>'), `Disclosure: HTML writes <details>: ${html[0]}`);
    assert.ok(html[0].includes('<details><summary>Inner</summary><p>Nested.</p>'), 'Disclosure: a nested one, closed');
    assert.ok(html[0].includes('Press <kbd>Ctrl</kbd>+<kbd>*C*</kbd> to copy.'), 'Keyboard: HTML writes <kbd>');
    // The HTML reader: loose text in a disclosure is a paragraph; one with no summary has none.
    const loose = await OfficeParser.parseOffice(Buffer.from('<details><summary>S</summary>Loose <b>text</b></details><details><p>x</p></details>'), { fileType: 'html' });
    assert.deepStrictEqual(loose.content.map(n => [n.type, n.metadata, n.children?.map(c => c.type)]), [
        ['disclosure', { summary: 'S' }, ['paragraph']], ['disclosure', {}, ['paragraph']],
    ], 'Disclosure: HTML read as disclosures');
    assert.ok((await loose.to('html', { htmlConfig: { standalone: false } })).value.includes('<details><p>x</p>'), 'Disclosure: none written where none was given');
    // EPUB: XHTML's open attribute, read back.
    const epub = await OfficeParser.parseOffice(Buffer.from((await ast.to('epub')).value as Uint8Array));
    const epubOuter = epub.content.find(n => n.type === 'disclosure');
    assert.deepStrictEqual(epubOuter?.metadata, { summary: 'Click me', open: true }, 'Disclosure: EPUB keeps it');
    // Elsewhere: the summary as a bold line, then the body, nothing hidden or lost.
    const text = (await ast.to('text')).value;
    assert.ok(text.includes('Click me\nSome body text.\n- a\n- b\nconst x = 1;\n\nconst y = 2;\nInner\nNested.'), `Disclosure: text writes it open: ${text}`);
    assert.ok(/Details\nx/.test((await loose.to('text')).value), 'Disclosure: one with no summary is "Details", as a browser shows it');
    for (const format of ['docx', 'odt', 'rtf', 'tex', 'pptx'] as const) {
        const back = await OfficeParser.parseOffice(Buffer.from((await ast.to(format)).value as unknown as Uint8Array), format === 'tex' ? { fileType: 'tex' } : {});
        const backMd = (await back.to('md')).value;
        assert.ok(/\*\*Click me\*\*[\s\S]*Some \*body\* text\.[\s\S]*\*\*Inner\*\*[\s\S]*Nested\./.test(backMd), `Disclosure: ${format} writes its summary in bold and its body: ${backMd}`);
        if (format === 'docx' || format === 'odt' || format === 'tex') assert.ok(backMd.includes('`Ctrl`'), `Keyboard: ${format} writes keyboard input as code: ${backMd}`);
    }
    assert.ok(/\{\\f2 Ctrl\}/.test((await ast.to('rtf')).value) && (await ast.to('rtf')).value.includes('\\f2\\fmodern\\fcharset0 Courier New;'), 'Keyboard: RTF writes keyboard input in the monospace font');
    const pdf = await OfficeParser.parseOffice(Buffer.from((await ast.to('pdf', { pdfConfig: { engine: 'native' } })).value as Uint8Array), { fileType: 'pdf', ocr: false, pdfParserConfig: { separateProcess: false } });
    assert.ok(/Click me[\s\S]*Nested\./.test((await pdf.to('text')).value), 'Disclosure: the native PDF engine draws it open');
    // In a list item, where no block goes: its summary in bold, then its body, reported.
    const inItem = await OfficeParser.parseOffice(Buffer.from('<ul><li><details><summary>S</summary><p>Body</p></details></li></ul>'), { fileType: 'html' });
    const itemMd = await inItem.to('md');
    assert.ok(/^- \*\*S\*\* Body/m.test(itemMd.value), `Disclosure: in a list item, its summary in bold: ${itemMd.value}`);
    assert.ok(itemMd.messages.some(m => m.code === 'CONTENT_NOT_REPRESENTABLE'), 'Disclosure: and that is reported');
    console.log('✓ Disclosures and keyboard input');
}

/**
 * htmlConfig.sheetLayout 'tables': each sheet a heading and a plain table (merged cells as spans, empty
 * rows kept, a header only where the source marks one, not the workbook's font on every cell); the
 * default 'viewer' unchanged.
 */
export async function testSheetsAsTables(): Promise<void> {
    const run = (text: string, extra: Record<string, unknown> = {}) => ({ type: 'text', text, formatting: { font: 'Arial', size: '10pt', ...extra } });
    const cell = (row: number, col: number, text: string, extra: Record<string, unknown> = {}, meta: Record<string, unknown> = {}) => ({ type: 'cell', metadata: { row, col, ...meta }, children: [run(text, extra)] });
    const ast = {
        type: 'xlsx', metadata: {}, attachments: [], content: [
            { type: 'sheet', metadata: { sheetName: 'Data <1>' }, children: [
                { type: 'row', children: [cell(0, 0, 'Region', { bold: true }), cell(0, 1, 'Visits', { bold: true })] },
                { type: 'row', children: [cell(1, 0, 'Both', {}, { colSpan: 2 })] },
                { type: 'row', children: [cell(3, 0, 'North'), cell(3, 1, 'Big', { font: 'Georgia', size: '14pt' })] },
            ] },
            { type: 'sheet', metadata: { sheetName: 'Notes' }, children: [{ type: 'row', children: [cell(0, 0, 'Plain'), cell(0, 2, 'Far')] }] },
        ],
    } as unknown as OfficeParserAST;
    const tables = (await OfficeGenerator.generate(ast, 'html', { htmlConfig: { standalone: { document: false, styles: 'none' }, sheetLayout: 'tables' } })).value;
    assert.ok(/<h2 id="sheet-0">Data &lt;1&gt;<\/h2>/.test(tables), `Sheets as tables: the sheet's name as a heading: ${tables}`);
    assert.ok(/<thead><tr><th><b>Region<\/b><\/th><th><b>Visits<\/b><\/th><\/tr><\/thead>/.test(tables), 'Sheets as tables: a bold first row is the header');
    assert.ok(tables.includes('<td colspan="2">Both</td>'), 'Sheets as tables: merged cells as spans');
    assert.ok(/<tr><td><\/td><td><\/td><\/tr>/.test(tables), 'Sheets as tables: an empty row kept in its place');
    assert.ok(/<td>North<\/td><td><span style="font-size: 14pt; font-family: Georgia[^"]*">Big<\/span><\/td>/.test(tables), 'Sheets as tables: a cell set apart keeps its own font');
    assert.ok(!/Arial|10pt|excel-|spreadsheet-|<script|sheet-tab/.test(tables), 'Sheets as tables: no viewer, no workbook font on every cell');
    assert.ok(/<h2 id="sheet-1">Notes<\/h2>\s*<table><tr><td>Plain<\/td><td><\/td><td>Far<\/td><\/tr>/.test(tables), 'Sheets as tables: no header unless marked, cells keep their columns');
    // The default is the viewer, as before.
    const viewer = (await OfficeGenerator.generate(ast, 'html', { htmlConfig: { standalone: { document: false, styles: 'none' } } })).value;
    assert.ok(viewer.includes('excel-col-header') && viewer.includes('spreadsheet-tab'), 'Sheets as tables: the viewer stays the default');
    // A workbook read from XLSX and ODS, and a bad value reported.
    for (const format of ['xlsx', 'ods'] as const) {
        const back = await OfficeParser.parseOffice(Buffer.from((await OfficeGenerator.generate(ast, format)).value as unknown as Uint8Array));
        const html = (await back.to('html', { htmlConfig: { standalone: false, sheetLayout: 'tables' } })).value;
        assert.ok(html.includes('Data &lt;1&gt;</h2>') && html.includes('colspan="2"') && !html.includes('excel-col-header'), `Sheets as tables: from ${format}: ${html}`);
    }
    const bad = await OfficeGenerator.generate(ast, 'html', { htmlConfig: { standalone: false, sheetLayout: 'grid' as any } });
    assert.ok(bad.messages.some(m => m.code === 'INVALID_CONFIG_VALUE') && bad.value.includes('excel-col-header'), 'Sheets as tables: an unknown layout is reported, the viewer used');
    console.log('✓ Sheets as tables');
}

/** Every note in an AST, in reference order, once each. */
function notesOf(ast: OfficeParserAST): OfficeContentNode[] {
    const out: OfficeContentNode[] = [];
    const walk = (nodes: OfficeContentNode[] | undefined) => {
        for (const n of nodes ?? []) {
            for (const note of n.notes ?? []) if (!out.includes(note)) out.push(note);
            walk(n.children);
        }
    };
    walk(ast.content);
    return out;
}

/** The marks the references in generated HTML show, with the keys their links go to. */
function referenceMarks(html: string): string[][] {
    return [...html.matchAll(/<sup data-footnote-ref="([^"]*)"[^>]*><a href="#footnote-([^"]*)">([^<]*)<\/a><\/sup>/g)].map(m => [m[1], m[3]]);
}

/**
 * A note's reference shows its mark (the one its source records, else its ordinal within its kind), never
 * its internal key, which stays in ids and links; a custom mark is kept through DOCX and ODT.
 */
export async function testNoteMarks(): Promise<void> {
    const md = await OfficeParser.parseOffice(Buffer.from('A[^a] B[^b] C[^c] D[^d].\n\n[^a]: First.\n[^b]: Second.\n[^c]: Third.\n[^d]: Fourth.\n'), { fileType: 'md' });
    const plain = (await md.to('html', { htmlConfig: { standalone: false } })).value;
    assert.deepStrictEqual(referenceMarks(plain), [['a', '1'], ['b', '2'], ['c', '3'], ['d', '4']], 'Notes: keys in links, ordinals shown');
    // One endnote among the footnotes, and one footnote with a mark its author chose.
    const [, b, c] = notesOf(md);
    (b.metadata as NoteMetadata).noteType = 'endnote';
    Object.assign(c.metadata as NoteMetadata, { mark: '*', customMark: true });
    const expected = ['1', 'i', '*', '2'];
    const shown = (html: string) => referenceMarks(html).map(r => r[1]);
    assert.deepStrictEqual(shown((await md.to('html', { htmlConfig: { standalone: false } })).value), expected, 'Notes: footnotes and endnotes counted apart, a custom mark kept');
    // Through the formats that number notes themselves, and back.
    for (const format of ['docx', 'odt', 'epub'] as const) {
        const written = (await md.to(format)).value as unknown as Uint8Array;
        const back = await OfficeParser.parseOffice(Buffer.from(written));
        const html = (await back.to('html', { htmlConfig: { standalone: false } })).value;
        assert.deepStrictEqual(shown(html), expected, `Notes: ${format} keeps the marks`);
        const custom = notesOf(back)[2];
        assert.strictEqual((custom.text ?? '').trim(), 'Third.', `Notes: ${format} keeps the custom mark out of the note's text`);
        if (format !== 'epub') assert.strictEqual((custom.metadata as NoteMetadata).customMark, true, `Notes: ${format} reads the custom mark back as one`);
    }
    const odt = new TextDecoder().decode(unzipSync((await md.to('odt')).value as Uint8Array)['content.xml']);
    assert.ok(odt.includes('<text:note-citation text:label="*">*</text:note-citation>'), 'Notes: ODT writes a custom mark as its label');
    assert.ok(odt.includes('<text:note text:id="edn1" text:note-class="endnote"><text:note-citation>i</text:note-citation>'), 'Notes: ODT caches an endnote citation as LibreOffice shows it');
    const docx = new TextDecoder().decode(unzipSync((await md.to('docx')).value as Uint8Array)['word/document.xml']);
    assert.ok(/<w:footnoteReference w:customMarkFollows="1" w:id="\d+"\/><w:t xml:space="preserve">\*<\/w:t>/.test(docx), 'Notes: DOCX writes a custom mark after its reference');
    // The native PDF engine and slides draw the same marks.
    const pdf = (await md.to('pdf', { pdfConfig: { engine: 'native' } })).value as Uint8Array;
    const pdfBack = await OfficeParser.parseOffice(Buffer.from(pdf), { fileType: 'pdf', ocr: false, pdfParserConfig: { separateProcess: false } });
    const pdfText = (await pdfBack.to('text')).value;
    assert.ok(/\*\s*Third\./.test(pdfText) && /i\s*Second\./.test(pdfText), `Notes: native PDF draws the marks: ${pdfText}`);
    const pptx = await OfficeParser.parseOffice(Buffer.from((await md.to('pptx')).value as Uint8Array));
    const pptxText = (await pptx.to('text')).value;
    assert.ok(/\*\.\s*Third\./.test(pptxText) && /i\.\s*Second\./.test(pptxText), `Notes: slides list the notes by their marks: ${pptxText}`);
    // An EPUB 3 note whose text repeats its mark as a link back to the reference: the mark is no text.
    const epubHtml = '<p>Text<sup id="calledF1"><a epub:type="noteref" href="#dataF1">1</a></sup>.</p>'
        + '<aside epub:type="footnote" id="dataF1"><p><sup><a href="#calledF1">1</a></sup>Figures are rounded.</p></aside>';
    const epubAst = await OfficeParser.parseOffice(Buffer.from(epubHtml), { fileType: 'html' });
    const epubNote = notesOf(epubAst)[0];
    assert.strictEqual((epubNote.text ?? '').trim(), 'Figures are rounded.', 'Notes: a link back to the reference is the back-link, not text');
    assert.strictEqual((epubNote.metadata as NoteMetadata).mark, '1', 'Notes: the reference text is the mark');
    assert.deepStrictEqual(shown((await epubAst.to('html', { htmlConfig: { standalone: false } })).value), ['1'], 'Notes: and shown again');
    console.log('✓ Note marks');
}

/** Milliseconds `run` takes. */
async function timed(run: () => Promise<unknown>): Promise<number> {
    const start = performance.now();
    await run();
    return performance.now() - start;
}

/** `text` read as `fileType` and written back as `format`, `passes` times: each pass's output. */
async function passesOf(text: string, fileType: 'html' | 'md', format: 'html' | 'md', passes = 3): Promise<string[]> {
    const out: string[] = [];
    let current = text;
    for (let pass = 0; pass < passes; pass++) {
        const ast = await OfficeParser.parseOffice(Buffer.from(current), { fileType: pass === 0 ? fileType : format });
        current = (await ast.to(format, { htmlConfig: { standalone: false } })).value;
        out.push(current);
    }
    return out;
}

/**
 * Nested figures are read once each: in time linear in their depth, each picture decoded once and each
 * quote numbered once, and a picture a figure inside captioned keeps that caption (the outer caption
 * stays a paragraph).
 */
export async function testNestedFiguresReadOnce(): Promise<void> {
    const depth = 250;
    const picture = `<img src="data:image/png;base64,${PIXEL_PNG}">`;
    const figures = '<figure><b>c</b><figcaption><b>x</b></figcaption>'.repeat(depth) + picture + '</figure>'.repeat(depth);
    const divs = '<div><b>c</b><div><b>x</b></div>'.repeat(depth) + picture + '</div>'.repeat(depth);
    let ast: OfficeParserAST | undefined;
    for (const source of [figures, divs]) await OfficeParser.parseOffice(Buffer.from(source), { fileType: 'html', extractAttachments: true });
    const figureTime = await timed(async () => { ast = await OfficeParser.parseOffice(Buffer.from(figures), { fileType: 'html', extractAttachments: true }); });
    const divTime = await timed(() => OfficeParser.parseOffice(Buffer.from(divs), { fileType: 'html', extractAttachments: true }));
    assert.ok(figureTime < 3 * divTime + 200, `Nested figures: ${depth} levels in ${figureTime.toFixed(0)} ms, the same divs in ${divTime.toFixed(0)} ms`);
    assert.strictEqual(ast!.attachments.length, 1, 'Nested figures: the picture is decoded once');
    // Through Markdown, whose HTML blocks the HTML reader reads.
    const mdFigures = '<figure><figcaption>x</figcaption>'.repeat(depth) + '<p>t</p>' + '</figure>'.repeat(depth) + '\n';
    const mdDivs = '<div><div>x</div>'.repeat(depth) + '<p>t</p>' + '</div>'.repeat(depth) + '\n';
    // Each read once first, so neither time holds the first read's compiling.
    for (const source of [mdFigures, mdDivs]) await OfficeParser.parseOffice(Buffer.from(source), { fileType: 'md' });
    const mdFigureTime = await timed(() => OfficeParser.parseOffice(Buffer.from(mdFigures), { fileType: 'md' }));
    const mdDivTime = await timed(() => OfficeParser.parseOffice(Buffer.from(mdDivs), { fileType: 'md' }));
    assert.ok(mdFigureTime < 3 * mdDivTime + 200, `Nested figures: Markdown, ${depth} levels in ${mdFigureTime.toFixed(0)} ms, the same divs in ${mdDivTime.toFixed(0)} ms`);
    // The inner figure's caption stays its picture's; the outer caption is a paragraph, nothing lost.
    const nested = '<figure><figure><img src="a.png" alt="pic"><figcaption>Inner caption</figcaption></figure><figcaption>Outer caption</figcaption></figure><p>After</p>';
    const nestedAst = await OfficeParser.parseOffice(Buffer.from(nested), { fileType: 'html' });
    assert.deepStrictEqual(nestedAst.content.map(n => [n.type, (n.metadata as ImageMetadata | undefined)?.caption ?? nodeText(n)]),
        [['image', 'Inner caption'], ['paragraph', 'Outer caption'], ['paragraph', 'After']], 'Nested figures: the inner caption kept, the outer one a paragraph');
    assert.strictEqual((await nestedAst.to('text')).value.trim(), '[Image: pic]\nInner caption\nOuter caption\nAfter', 'Nested figures: text');
    const html = await passesOf(nested, 'html', 'html');
    assert.ok(html[0].includes('<figcaption>Inner caption</figcaption></figure>') && html[0].includes('<p>Outer caption</p>'), `Nested figures: HTML: ${html[0]}`);
    assert.deepStrictEqual(html, [html[0], html[0], html[0]], 'Nested figures: HTML to HTML is stable over three passes');
    const md = await passesOf((await nestedAst.to('md')).value, 'md', 'md');
    assert.deepStrictEqual(md, [md[0], md[0], md[0]], 'Nested figures: Markdown to Markdown is stable over three passes');
    // A quote in a captioned figure is the document's first quote, numbered once.
    const quoted = await OfficeParser.parseOffice(Buffer.from('<figure><blockquote>Q</blockquote><figcaption>Who</figcaption></figure>'), { fileType: 'html' });
    assert.strictEqual((quoted.content[0].metadata as ParagraphMetadata).quoteId, 'html-quote-1', 'Nested figures: a quote in a figure is numbered once');
    console.log('✓ Nested figures read once');
}

/** A node's text, from its runs where it has none of its own. */
function nodeText(node: OfficeContentNode): string {
    return node.text ?? (node.children ?? []).map(nodeText).join('');
}

/** The length of every `text` in a tree of nodes. */
function textLength(nodes: OfficeContentNode[] | undefined): number {
    let length = 0;
    for (const node of nodes ?? []) length += (node.text?.length ?? 0) + textLength(node.children);
    return length;
}

/**
 * A disclosure carries no text of its own: its summary is its metadata and its body its children, so a
 * disclosure with no body is not written with its summary as content, and nested ones copy nothing.
 */
export async function testDisclosureCarriesNoText(): Promise<void> {
    const empty = '<details><summary>Only summary</summary></details><p>x</p>';
    const ast = await OfficeParser.parseOffice(Buffer.from(empty), { fileType: 'html' });
    assert.deepStrictEqual(ast.content[0], { type: 'disclosure', metadata: { summary: 'Only summary' }, children: [] }, 'Disclosure: no text of its own');
    assert.strictEqual((await ast.to('text')).value.trim(), 'Only summary\nx', 'Disclosure: text shows its summary once');
    const html = await passesOf(empty, 'html', 'html');
    assert.ok(html[0].includes('<details><summary>Only summary</summary></details>'), `Disclosure: an empty one written empty: ${html[0]}`);
    assert.deepStrictEqual(html, [html[0], html[0], html[0]], 'Disclosure: HTML to HTML is stable over three passes');
    const md = await passesOf((await ast.to('md')).value, 'md', 'md');
    assert.deepStrictEqual(md, [md[0], md[0], md[0]], 'Disclosure: Markdown to Markdown is stable over three passes');
    // Nested disclosures hold their body's text once, as nested divs do.
    const levels = 200;
    const body = '<p>' + 'word '.repeat(50_000) + '</p>';
    const disclosures = '<details><summary>s</summary>'.repeat(levels) + body + '</details>'.repeat(levels);
    const divs = '<div><p>s</p>'.repeat(levels) + body + '</div>'.repeat(levels);
    let nested: OfficeParserAST | undefined;
    let flat: OfficeParserAST | undefined;
    const nestedTime = await timed(async () => { nested = await OfficeParser.parseOffice(Buffer.from(disclosures), { fileType: 'html' }); });
    const flatTime = await timed(async () => { flat = await OfficeParser.parseOffice(Buffer.from(divs), { fileType: 'html' }); });
    assert.ok(textLength(nested!.content) < 2 * textLength(flat!.content), `Disclosure: nested ones hold ${textLength(nested!.content)} characters of text, divs ${textLength(flat!.content)}`);
    assert.ok(nestedTime < 3 * flatTime + 200, `Disclosure: ${levels} levels in ${nestedTime.toFixed(0)} ms, divs in ${flatTime.toFixed(0)} ms`);
    console.log('✓ Disclosure carries no text');
}

/**
 * A declaration ends at its first `>`, as HTML has it; only the document's DOCTYPE, in its prolog, may
 * carry an XML internal subset, read when well formed. The reading is linear whatever the declarations.
 */
export async function testDoctypeAndDeclarations(): Promise<void> {
    const textOf = async (source: string, config: { fileType?: 'html' } = { fileType: 'html' }) => {
        const ast = await OfficeParser.parseOffice(Buffer.from(source), config);
        return [ast.type, (await ast.to('text')).value.trim()];
    };
    assert.deepStrictEqual(await textOf('<!DOCTYPE html [><html><body><p>Visible</p><p>Item a[0] is ok</p><p>Tail</p></body></html>'),
        ['html', 'Visible\nItem a[0] is ok\nTail'], 'DOCTYPE: a `[` that opens no subset ends at the first `>`');
    assert.deepStrictEqual(await textOf('<!DOCTYPE html PUBLIC "x[y" "z"><p>Pub</p><p>b[1]</p>'), ['html', 'Pub\nb[1]'], 'DOCTYPE: a `[` in a quoted public id opens no subset');
    // XHTML: a subset whose entity values, comments and instructions hold `>`, and a parameter entity.
    assert.deepStrictEqual(await textOf('<?xml version="1.0"?>\n<!DOCTYPE html [<!ENTITY a ">"> <!-- c > --> <?pi > ?> %pe; ]>\n'
        + '<html xmlns="http://www.w3.org/1999/xhtml"><body><p>X</p></body></html>', {}), ['html', 'X'], 'DOCTYPE: an XHTML subset read whole');
    assert.deepStrictEqual(await textOf('<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "]>">]><x:html xmlns:x="http://www.w3.org/1999/xhtml"><x:body><x:p>Q</x:p></x:body></x:html>', {}),
        ['html', 'Q'], 'DOCTYPE: an entity value holding `]>` leaves the file XHTML');
    // A DOCTYPE after content is a declaration to its first `>`; Word's conditional markers are nothing.
    assert.deepStrictEqual(await textOf('<p>Before</p><!DOCTYPE html [<!ENTITY a ">">]><p>After</p>'), ['html', 'Before\n">]>\nAfter'], 'DOCTYPE: one after content ends at its first `>`');
    assert.deepStrictEqual(await textOf('<!DOCTYPE html>\n<p>Plain</p><![if !supportLists]>1.<![endif]> item'), ['html', 'Plain\n1. item'], 'DOCTYPE: conditional markers unchanged');
    // Linear: each against plain paragraphs of the same size, in the same run.
    const n = 160_000;
    const plain = (length: number) => '<p>' + 'a'.repeat(length) + '</p>';
    for (const [name, source] of [
        ['declarations', '<p>x</p>' + '<!a>'.repeat(2 * n) + '<p>y</p>'],
        ['repeated DOCTYPEs', '<!DOCTYPE a [<!ENTITY x "> '.repeat(n) + '<p>y</p>'],
        ['an unclosed subset', '<!DOCTYPE a [' + '<!ENTITY x "a">'.repeat(n) + '<p>y</p>'],
        ['brackets in text', '<!DOCTYPE html><p>' + 'a[b] <!x[> '.repeat(n) + '</p>'],
    ]) {
        const time = await timed(() => OfficeParser.parseOffice(Buffer.from(source), { fileType: 'html' }));
        const baseline = await timed(() => OfficeParser.parseOffice(Buffer.from(plain(source.length)), { fileType: 'html' }));
        assert.ok(time < 3 * baseline + 200, `DOCTYPE: ${name} (${source.length} bytes) in ${time.toFixed(0)} ms, plain text in ${baseline.toFixed(0)} ms`);
    }
    console.log('✓ DOCTYPE and declarations');
}

/** The quote ids of a document's quote paragraphs, in order. */
function quoteIds(ast: OfficeParserAST): (string | undefined)[] {
    return ast.content.filter(n => (n.metadata as ParagraphMetadata | undefined)?.style === 'Quote').map(n => (n.metadata as ParagraphMetadata).quoteId);
}

/**
 * The quotes of a document read as several HTML parts (Markdown's HTML blocks, an EPUB's chapters, a
 * DOCX's imported HTML) are numbered across the document: two quotes of two parts stay two quotes.
 */
export async function testQuoteIdsAcrossParts(): Promise<void> {
    const md = await OfficeParser.parseOffice(Buffer.from('<blockquote>alpha</blockquote>\n\n<blockquote>beta</blockquote>\n'), { fileType: 'md' });
    assert.deepStrictEqual(quoteIds(md), ['md-quote-1', 'md-quote-2'], 'Quote ids: Markdown\'s HTML blocks numbered with the document');
    assert.strictEqual((await md.to('md')).value, '> alpha\n\n> beta', 'Quote ids: two quotes in Markdown');
    assert.strictEqual((await md.to('html', { htmlConfig: { standalone: false } })).value.match(/<blockquote>/g)?.length, 2, 'Quote ids: two quotes in HTML');
    const mixed = await OfficeParser.parseOffice(Buffer.from('> one\n\n<blockquote>two</blockquote>\n\n> three\n'), { fileType: 'md' });
    assert.strictEqual(new Set(quoteIds(mixed)).size, 3, `Quote ids: Markdown and HTML quotes share one numbering: ${quoteIds(mixed)}`);
    // An EPUB whose first chapter ends and second starts with a quote.
    const chapter = (body: string) => `<?xml version="1.0"?><html xmlns="http://www.w3.org/1999/xhtml"><head><title>c</title></head><body>${body}</body></html>`;
    const epub = Buffer.from(zipSync({
        mimetype: strToU8('application/epub+zip'),
        'META-INF/container.xml': strToU8('<?xml version="1.0"?><container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles></container>'),
        'OEBPS/content.opf': strToU8('<?xml version="1.0"?><package xmlns="http://www.idpf.org/2007/opf" version="3.0"><metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>T</dc:title></metadata>'
            + '<manifest><item id="c1" href="c1.xhtml" media-type="application/xhtml+xml"/><item id="c2" href="c2.xhtml" media-type="application/xhtml+xml"/></manifest><spine><itemref idref="c1"/><itemref idref="c2"/></spine></package>'),
        'OEBPS/c1.xhtml': strToU8(chapter('<p>One</p><blockquote><p>First</p></blockquote>')),
        'OEBPS/c2.xhtml': strToU8(chapter('<blockquote><p>Second</p></blockquote><p>Two</p>')),
    }));
    const book = await OfficeParser.parseOffice(epub);
    assert.deepStrictEqual(quoteIds(book), ['html-quote-1', 'html-quote-2'], 'Quote ids: an EPUB\'s chapters numbered across the book');
    assert.ok((await book.to('md')).value.includes('> First\n\n> Second'), 'Quote ids: two quotes from two chapters');
    // A DOCX with two HTML chunks, each one quote.
    const relationship = (id: string, target: string) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="${target}"/>`;
    const docx = Buffer.from(zipSync({
        '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
        'word/document.xml': strToU8('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            + '<w:body><w:altChunk r:id="c1"/><w:altChunk r:id="c2"/></w:body></w:document>'),
        'word/_rels/document.xml.rels': strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationship('c1', 'a.htm')}${relationship('c2', 'b.htm')}</Relationships>`),
        'word/a.htm': strToU8('<blockquote><p>First</p></blockquote>'),
        'word/b.htm': strToU8('<blockquote><p>Second</p></blockquote>'),
    }));
    const word = await OfficeParser.parseOffice(docx);
    assert.deepStrictEqual(quoteIds(word), ['html-quote-1', 'html-quote-2'], 'Quote ids: a DOCX\'s HTML chunks numbered across the document');
    assert.strictEqual((await word.to('md')).value, '> First\n\n> Second', 'Quote ids: two quotes from two chunks');
    console.log('✓ Quote ids across parts');
}

/** A Markdown document's blocks, each its type and its children's types. */
async function markdownShape(markdown: string): Promise<[string, string[] | undefined][]> {
    const ast = await OfficeParser.parseOffice(Buffer.from(markdown), { fileType: 'md' });
    return ast.content.map(n => [n.type, n.children?.map(c => c.type)]);
}

/**
 * A `<details>` block's body in Markdown: Markdown but for pieces CommonMark starts an HTML block with;
 * a disclosure's tags in code spans and indented code are text; it holds everything up to its close
 * however long; tables are joined as before.
 */
export async function testMarkdownDisclosureBodies(): Promise<void> {
    for (const body of ['Use the `<details>` element.', 'Use the `<summary>` element.', 'Plain *body*.']) {
        const source = `<details>\n<summary>Help</summary>\n\n${body}\n\n</details>\n\nAfter\n`;
        const ast = await OfficeParser.parseOffice(Buffer.from(source), { fileType: 'md' });
        assert.deepStrictEqual(ast.content.map(n => [n.type, n.children?.map(c => c.type)]), [['disclosure', ['paragraph']], ['paragraph', ['text']]], `Markdown disclosure: ${body} is one paragraph in it`);
        const html = (await ast.to('html', { htmlConfig: { standalone: false } })).value;
        if (body.includes('`')) assert.ok(/<code>&lt;(details|summary)&gt;<\/code>/.test(html), `Markdown disclosure: the code span kept: ${html}`);
    }
    // Our own Markdown of a disclosure mentioning its tag, in a code span and in code.
    const help = await OfficeParser.parseOffice(Buffer.from('<details><summary>Help</summary><p>Use <code>&lt;details&gt;</code> to fold.</p><pre><code>x\n&lt;/details&gt;\ny</code></pre><p>Inner end</p></details><p>After</p>'), { fileType: 'html' });
    const helpMd = (await help.to('md')).value;
    assert.deepStrictEqual(await markdownShape(helpMd), [['disclosure', ['paragraph', 'code', 'paragraph']], ['paragraph', ['text']]], `Markdown disclosure: code mentioning its tag stays in it: ${helpMd}`);
    // A body starting with an inline tag is a paragraph; indented code is code; nested code spans are text.
    const keys = await OfficeParser.parseOffice(Buffer.from('<details>\n<summary>Keys</summary>\n\n<kbd>Ctrl</kbd> and *bold*.\n\n</details>\n\nAfter\n'), { fileType: 'md' });
    const paragraph = keys.content[0].children![0];
    assert.ok(paragraph.type === 'paragraph' && paragraph.children!.some(c => c.formatting?.keyboard) && paragraph.children!.some(c => c.formatting?.italic), 'Markdown disclosure: <kbd> and emphasis in a paragraph');
    assert.deepStrictEqual(await markdownShape('<details>\n<summary>Keys</summary>\n\n    <details> in indented code\n\n</details>\n\nAfter\n'), [['disclosure', ['code']], ['paragraph', ['text']]], 'Markdown disclosure: indented code is code');
    assert.deepStrictEqual(await markdownShape('<details>\n<summary>S</summary>\n\nA ``x `<details>` y`` b and `</details>` c.\n\nMore\n\n</details>\n\nAfter\n'),
        [['disclosure', ['paragraph', 'paragraph']], ['paragraph', ['text']]], 'Markdown disclosure: tags in code spans are text');
    // A closing tag on a paragraph's last line closes it, as a browser has it.
    assert.deepStrictEqual(await markdownShape('<details>\n<summary>S</summary>\n\n*Body*\n</details>\n\nAfter\n'), [['disclosure', ['paragraph']], ['paragraph', ['text']]], 'Markdown disclosure: an inline closing tag closes it');
    // Through HTML and Markdown: stable over three passes.
    const kbdHtml = await OfficeParser.parseOffice(Buffer.from('<details><summary>K</summary><p><kbd>Ctrl</kbd> then <b>bold</b></p></details>'), { fileType: 'html' });
    const kbdPasses = await passesOf((await kbdHtml.to('md')).value, 'md', 'md');
    assert.ok(kbdPasses[0].includes('<kbd>Ctrl</kbd> then **bold**'), `Markdown disclosure: <kbd> written: ${kbdPasses[0]}`);
    assert.deepStrictEqual(kbdPasses, [kbdPasses[0], kbdPasses[0], kbdPasses[0]], 'Markdown disclosure: <kbd> in a body is stable over three passes');
    // A long body, whatever its length.
    for (const count of [70, 500]) {
        const long = await OfficeParser.parseOffice(Buffer.from('<details><summary>S</summary>' + Array.from({ length: count }, (_, i) => `<p>para ${i}</p>`).join('') + '</details><p>After</p>'), { fileType: 'html' });
        const longMd = (await long.to('md')).value;
        const back = await OfficeParser.parseOffice(Buffer.from(longMd), { fileType: 'md' });
        assert.deepStrictEqual(back.content.map(n => [n.type, n.children?.length]), [['disclosure', count], ['paragraph', 1]], `Markdown disclosure: ${count} paragraphs stay in it`);
        const passes = await passesOf(longMd, 'md', 'md');
        assert.deepStrictEqual(passes, [longMd, longMd, longMd], `Markdown disclosure: ${count} paragraphs stable over three passes`);
    }
    // Tables as before: Markdown between a cell's tags is the cell's.
    assert.deepStrictEqual(await markdownShape('<table><tr><td>\n\ncell *x*\n\n</td></tr></table>\n'), [['table', ['row']]], 'Markdown disclosure: a table joined as before');
    // Linear: each against plain paragraphs of the same size, in the same run.
    const n = 4000;
    const plain = (length: number) => 'Some *plain* text here.\n\n'.repeat(Math.ceil(length / 25));
    for (const [name, source] of [
        ['openings', '<details>\n<summary>s</summary>\n\n'.repeat(n)],
        ['code spans', '<details>\n<summary>s</summary>\n\n' + ('`<details>` ``` x\n\n').repeat(2 * n) + '</details>\n'],
        ['unclosed code spans', '<details>\n\n' + 'a ```` <details> '.repeat(n) + '\n\n</details>\n'],
        ['lone tag lines', '<details>\n\n' + ('<a' + ' b'.repeat(2000) + '\n\n').repeat(25) + '</details>\n'],
        ['closers of the other kind', '<details>\n\n' + '<table></details>\n\n'.repeat(2 * n) + '</table>\n'],
        ['inline closers', '<details>\n\n' + 'a `x` </details> <details>\n\n'.repeat(2 * n) + '</details>\n'],
    ]) {
        const time = await timed(() => OfficeParser.parseOffice(Buffer.from(source), { fileType: 'md' }));
        const baseline = await timed(() => OfficeParser.parseOffice(Buffer.from(plain(source.length)), { fileType: 'md' }));
        assert.ok(time < 3 * baseline + 200, `Markdown disclosure: ${name} (${source.length} bytes) in ${time.toFixed(0)} ms, paragraphs in ${baseline.toFixed(0)} ms`);
    }
    console.log('✓ Markdown disclosure bodies');
}

/** The marks of an ODT's notes, in reference order: each `kind:mark`. */
async function odtNoteMarks(odt: Uint8Array): Promise<string[]> {
    const ast = await OfficeParser.parseOffice(Buffer.from(odt), { fileType: 'odt' });
    return notesOf(ast).map(note => `${(note.metadata as NoteMetadata).noteType}:${(note.metadata as NoteMetadata).mark}`);
}

/**
 * An ODT note's mark is its number as the notes configuration gives it (format and start value), as
 * LibreOffice shows it, not the citation a writer cached; a note with a label keeps it and takes no
 * number; footnotes numbered per page or chapter keep the cached citation.
 */
export async function testOdtNoteMarksFromConfiguration(): Promise<void> {
    const sample = new Uint8Array(fs.readFileSync('test/files/test.odt'));
    assert.deepStrictEqual(await odtNoteMarks(sample), ['footnote:1', 'endnote:i'], 'ODT note marks: an endnote is numbered i, as LibreOffice shows it');
    const sampleAst = await OfficeParser.parseOffice(Buffer.from(sample), { fileType: 'odt' });
    assert.deepStrictEqual(referenceMarks((await sampleAst.to('html', { htmlConfig: { standalone: false } })).value).map(r => r[1]), ['1', 'i'], 'ODT note marks: and HTML shows them');
    // Our own ODT of two footnotes, an endnote and a custom mark, its citations cached as 9 and a configuration added.
    const md = await OfficeParser.parseOffice(Buffer.from('A[^a] B[^b] C[^c] D[^d].\n\n[^a]: First.\n[^b]: Second.\n[^c]: Third.\n[^d]: Fourth.\n'), { fileType: 'md' });
    const [, b, c] = notesOf(md);
    (b.metadata as NoteMetadata).noteType = 'endnote';
    Object.assign(c.metadata as NoteMetadata, { mark: '*', customMark: true });
    const files = unzipSync((await md.to('odt')).value as Uint8Array);
    const decode = (name: string) => new TextDecoder().decode(files[name]);
    const content = decode('content.xml').replace(/<text:note-citation>[^<]*<\/text:note-citation>/g, '<text:note-citation>9</text:note-citation>');
    const configured = (footnote: string, endnote: string) => {
        const styles = decode('styles.xml').replace('</office:styles>', `<text:notes-configuration text:note-class="footnote" ${footnote}/><text:notes-configuration text:note-class="endnote" ${endnote}/></office:styles>`);
        return zipSync({ ...files, 'content.xml': strToU8(content), 'styles.xml': strToU8(styles) });
    };
    assert.deepStrictEqual(await odtNoteMarks(zipSync({ ...files, 'content.xml': strToU8(content) })), ['footnote:1', 'endnote:i', 'footnote:*', 'footnote:2'],
        'ODT note marks: with no configuration, footnotes 1, 2 and endnotes i; a label takes no number');
    // Each as LibreOffice 26 numbers the same file (saved again, its citations are these).
    assert.deepStrictEqual(await odtNoteMarks(configured('style:num-format="a" text:start-value="2" text:start-numbering-at="document"', 'text:start-value="4"')), ['footnote:c', 'endnote:5', 'footnote:*', 'footnote:d'],
        'ODT note marks: the format and the start value (an offset); a configuration with no format numbers 1, 2, 3');
    assert.deepStrictEqual(await odtNoteMarks(configured('text:start-numbering-at="document" text:start-value="1"', 'style:num-format="a"')), ['footnote:2', 'endnote:a', 'footnote:*', 'footnote:3'],
        'ODT note marks: footnotes numbered from 2, endnotes by letter');
    assert.deepStrictEqual(await odtNoteMarks(configured('style:num-format="A" text:start-numbering-at="page"', 'style:num-format="I"')), ['footnote:9', 'endnote:I', 'footnote:*', 'footnote:9'],
        'ODT note marks: footnotes numbered per page keep the cached citation');
    assert.deepStrictEqual(await odtNoteMarks(configured('style:num-format="a"', 'style:num-format="i"')), ['footnote:9', 'endnote:i', 'footnote:*', 'footnote:9'],
        'ODT note marks: a footnote configuration naming no restart restarts per page, as LibreOffice reads it');
    console.log('✓ ODT note marks from the notes configuration');
}

/**
 * Whether `run` takes under three times what `baseline` (plain input of the same size) takes, plus 200 ms:
 * measured again when over, the faster kept, so a pause of the machine is not taken for the limit failing
 * (see helpers/timing). Both times, for the message.
 */
async function linearAgainst(run: () => Promise<unknown>, baseline: () => Promise<unknown>): Promise<{ within: boolean; time: number; plain: number }> {
    let plain = await timed(baseline);
    let time = await timed(run);
    if (time >= 3 * plain + 200) {
        plain = Math.max(plain, await timed(baseline));
        time = Math.min(time, await timed(run));
    }
    return { within: time < 3 * plain + 200, time, plain };
}

/** `text` read as `fileType` and written as `format`, then read back and written again, `passes` times: each pass's output and text. */
async function passesWithText(text: string, fileType: 'html' | 'md', format: 'html' | 'md', passes = 3): Promise<{ out: string[]; text: string[] }> {
    const out: string[] = [];
    const texts: string[] = [];
    let ast = await OfficeParser.parseOffice(Buffer.from(text), { fileType });
    const first = (await ast.to('text')).value;
    for (let pass = 0; pass < passes; pass++) {
        out.push((await ast.to(format, { htmlConfig: { standalone: false } })).value);
        ast = await OfficeParser.parseOffice(Buffer.from(out[pass]), { fileType: format });
        texts.push((await ast.to('text')).value);
    }
    return { out, text: [first, ...texts] };
}

/** A tree's shape: each node's type, a disclosure's summary in brackets, and its children's in parentheses. */
function treeShape(nodes: OfficeContentNode[] | undefined): string {
    return (nodes ?? []).map(n => n.type + (n.type === 'disclosure' ? `[${(n.metadata as { summary?: string } | undefined)?.summary ?? ''}]` : '')
        + (n.type !== 'text' && n.children?.length ? `(${treeShape(n.children)})` : '')).join(',');
}

/** Every block kind in a disclosure's body, from Markdown. */
const EVERY_KIND_MD = (() => {
    const body = [
        '# H1 {#h1}', '## H2', '###### H6',
        '<div style="text-align: center">Centered **bold**</div>',
        '<div style="text-align: right">\n\n## Right heading\n\n</div>',
        '<div align="center">\n\n*centered*\n\n</div>',
        'Plain *em* **strong** `code` [link](https://e.com) and math $x^2$.',
        '| A | B |\n| :-- | --: |\n| 1 | *2* |',
        '<table><tr><td>\n\ncell *md*\n\n</td></tr></table>',
        '- a\n- b\n  - b1\n- [x] done',
        '1. one\n2. two',
        '> quote one\n>\n> quote two',
        '```js\nconst x = "<details>";\n```',
        '    indented code',
        '$$\n\\int_0^1 x\\,dx\n$$',
        '![alt](a.png)',
        '<figure><img src="a.png" alt="p"><figcaption>Cap</figcaption></figure>',
        'Text with a note[^n1].',
        '---',
        '<details open>\n<summary>Inner</summary>\n\nInner *body*\n\n<details>\n<summary>Deepest</summary>\n\nDeep\n\n</details>\n\n</details>',
    ].join('\n\n');
    return `Before\n\n<details>\n<summary>Every kind</summary>\n\n${body}\n\n</details>\n\nAfter\n\n[^n1]: The *note*.\n`;
})();

/** Every block kind in a disclosure, from HTML. */
const EVERY_KIND_HTML = '<p>Before</p><details><summary>Every kind</summary><h1>H1</h1><h2 style="text-align:center">C2</h2><p style="text-align:right">Right <b>b</b></p><p style="text-align:justify">Just</p>'
    + '<table><tr><th>A</th><th>B</th></tr><tr><td>1</td><td><i>2</i></td></tr></table><ul><li>a</li><li>b<ul><li>b1</li></ul></li></ul><ol><li>one</li></ol>'
    + '<blockquote><p>q1</p><p>q2</p></blockquote><pre><code>x &lt;/details&gt;\ny</code></pre><p><img src="a.png" alt="i"></p><figure><img src="a.png" alt="p"><figcaption>Cap</figcaption></figure>'
    + '<details open><summary>Inner</summary><p>In <kbd>K</kbd></p><details><summary>D</summary><p>Deep</p></details></details><hr></details><p>After</p>';

/**
 * A Markdown disclosure's body is a Markdown document of its own, read by the same reading as the
 * document: an aligned wrapper, a one-line aligned div, a `<div>` holding Markdown and an inline
 * `<details>` in it read as at top level, no piece token or escaped Markdown in any output, every block
 * kind stable over three passes, nesting bounded with nothing dropped, in time linear in the text.
 */
export async function testMarkdownDisclosureSubDocuments(): Promise<void> {
    const token = /[a-z0-9]{6,}:\d/;
    // The editor's HTML of a centred paragraph and of a centred heading in a disclosure, to Markdown and back three times.
    for (const [html, expected, text] of [
        ['<details><summary>S</summary><p style="text-align:center">Centered <b>bold</b></p></details>',
            '<details>\n<summary>S</summary>\n\n<div style="text-align: center">Centered **bold**</div>\n\n</details>', 'S\nCentered bold'],
        ['<details><summary>S</summary><h2 style="text-align:center">Heading</h2><p>Plain body</p></details><p>After</p>',
            '<details>\n<summary>S</summary>\n\n<div style="text-align: center">\n\n## Heading {#heading}\n\n</div>\n\nPlain body\n\n</details>\n\nAfter', 'S\nHeading\nPlain body\nAfter'],
    ]) {
        let ast = await OfficeParser.parseOffice(Buffer.from(html), { fileType: 'html' });
        for (let pass = 0; pass < 3; pass++) {
            const md = (await ast.to('md')).value;
            assert.strictEqual(md, expected, `Disclosure sub-document: pass ${pass} of ${html}`);
            ast = await OfficeParser.parseOffice(Buffer.from(md), { fileType: 'md' });
            const shown = (await ast.to('text')).value;
            assert.strictEqual(shown, text, `Disclosure sub-document: pass ${pass} text, no token or escaped mark`);
            const out = shown + (await ast.to('html', { htmlConfig: { standalone: false } })).value;
            assert.ok(!token.test(out) && !out.includes('\\*'), `Disclosure sub-document: no piece token or backslash in pass ${pass}: ${out}`);
        }
    }
    // Those Markdown sources read as at top level.
    const textOf = async (md: string) => (await (await OfficeParser.parseOffice(Buffer.from(md), { fileType: 'md' })).to('text')).value;
    assert.strictEqual(await textOf('<details>\n<summary>S</summary>\n\n<div style="text-align: center">\n\n## Heading {#heading}\n\n</div>\n\nPlain body\n\n</details>\n\nAfter'),
        'S\nHeading\nPlain body\nAfter', 'Disclosure sub-document: an aligned wrapper\'s heading, and the body after it, stay in the disclosure');
    assert.strictEqual(await textOf('<details>\n<summary>S</summary>\n\n<div style="text-align: center">Centered **bold**</div>\n\n</details>'), 'S\nCentered bold',
        'Disclosure sub-document: a one-line aligned div\'s Markdown is read');
    // A <div> holding Markdown, an aligned one, and an aligned paragraph tag around a picture.
    const shapes: [string, unknown][] = [
        ['<details>\n<summary>S</summary>\n\n<div>\n\n*md in div*\n\n</div>\n\nBody\n\n</details>\n\nAfter\n', [
            { type: 'disclosure', metadata: { summary: 'S' }, children: [
                { type: 'paragraph', metadata: {}, children: [{ type: 'text', text: 'md in div', formatting: { italic: true } }] },
                { type: 'paragraph', metadata: {}, children: [{ type: 'text', text: 'Body' }] }] },
            { type: 'paragraph', metadata: {}, children: [{ type: 'text', text: 'After' }] }]],
        ['<details>\n<summary>S</summary>\n\n<div align="center">\n\n*centered*\n\n</div>\n\n</details>\n', [
            { type: 'disclosure', metadata: { summary: 'S' }, children: [
                { type: 'paragraph', metadata: { alignment: 'center' }, children: [{ type: 'text', text: 'centered', formatting: { italic: true } }] }] }]],
        ['<details>\n<summary>S</summary>\n\n<p align="center">\n\n![x](a.png)\n\n</p>\n\nBody\n\n</details>\n', [
            { type: 'disclosure', metadata: { summary: 'S' }, children: [
                { type: 'paragraph', metadata: { alignment: 'center' }, children: [{ type: 'image', metadata: { url: 'a.png', altText: 'x' } }] },
                { type: 'paragraph', metadata: {}, children: [{ type: 'text', text: 'Body' }] }] }]],
    ];
    for (const [source, expected] of shapes) {
        assert.deepStrictEqual(JSON.parse(JSON.stringify((await OfficeParser.parseOffice(Buffer.from(source), { fileType: 'md' })).content)), expected, `Disclosure sub-document: ${JSON.stringify(source)}`);
    }
    // An inline <details> in a body paragraph opens a disclosure there; the outer body after it stays in the outer one.
    const inline = await OfficeParser.parseOffice(Buffer.from('<details>\n<summary>Outer</summary>\n\nIntro text <details><summary>Inner</summary>\n\nInner body\n\n</details>\n\nOuter body continues\n\n</details>\n\nAfter\n'), { fileType: 'md' });
    assert.strictEqual(treeShape(inline.content), 'disclosure[Outer](paragraph(text),disclosure[Inner](paragraph(text)),paragraph(text)),paragraph(text)', 'Disclosure sub-document: an inline opener nests');
    assert.strictEqual((await inline.to('text')).value, 'Outer\nIntro text\nInner\nInner body\nOuter body continues\nAfter', 'Disclosure sub-document: an inline opener\'s text');
    // Every block kind, and the edges of a body: the same Markdown, HTML and text over three passes.
    const cases: [string, 'md' | 'html', string, string][] = [
        ['every kind from Markdown', 'md', EVERY_KIND_MD, ''],
        ['every kind from HTML', 'html', EVERY_KIND_HTML, ''],
        ['an aligned wrapper', 'md', '<details>\n<summary>S</summary>\n\n<div style="text-align: center">\n\n## Heading {#heading}\n\n</div>\n\nPlain body\n\n</details>\n\nAfter\n', 'disclosure[S](heading,paragraph),paragraph'],
        ['a one-line aligned div', 'md', '<details>\n<summary>S</summary>\n\n<div style="text-align: center">Centered **bold**</div>\n\n</details>\n', 'disclosure[S](paragraph)'],
        ['a summary after a blank line', 'md', '<details>\n\n<summary>Later</summary>\n\nBody\n\n</details>\n\nAfter\n', 'disclosure[Later](paragraph),paragraph'],
        ['no summary', 'md', '<details>\n\nBody *x*\n\n</details>\n\nAfter\n', 'disclosure[](paragraph),paragraph'],
        ['raw HTML after the summary', 'md', '<details>\n<summary>S</summary>\nRaw *not md* line\n\nMd *yes*\n\n</details>\n\nAfter\n', 'disclosure[S](paragraph,paragraph),paragraph'],
        ['one block', 'md', '<details><summary>S</summary><p>x</p></details>\n\nAfter\n', 'disclosure[S](paragraph),paragraph'],
        ['text after the closing tag', 'md', '<details>\n<summary>S</summary>\n\nBody\n\n</details> tail *t*\n\nAfter\n', 'disclosure[S](paragraph),paragraph,paragraph'],
        ['an inline opener', 'md', '<details>\n<summary>Outer</summary>\n\nIntro text <details><summary>Inner</summary>\n\nInner body\n\n</details>\n\nOuter body continues\n\n</details>\n\nAfter\n', ''],
        ['never closed', 'md', '<details>\n<summary>S</summary>\n\n# All\n\nthe rest\n', 'disclosure[S](heading,paragraph)'],
        ['an anchor before it', 'md', '<a id="d1"></a>\n<details>\n<summary>S</summary>\n\nBody\n\n</details>\n\nAfter\n', 'disclosure[S](paragraph),paragraph'],
    ];
    const blockShape = (nodes: OfficeContentNode[] | undefined): string => (nodes ?? []).filter(n => n.type !== 'text')
        .map(n => n.type + (n.type === 'disclosure' ? `[${(n.metadata as { summary?: string }).summary ?? ''}](${blockShape(n.children)})` : '')).join(',');
    for (const [name, fileType, source, shape] of cases) {
        if (shape) assert.strictEqual(blockShape((await OfficeParser.parseOffice(Buffer.from(source), { fileType })).content), shape, `Disclosure sub-document: ${name}, its blocks`);
        // HTML: a pipe table's HTML gains white space between cells on its second pass (8.6.0 the same), so only Markdown for it.
        for (const format of name === 'every kind from Markdown' ? ['md' as const] : ['md' as const, 'html' as const]) {
            const { out, text } = await passesWithText(source, fileType, format);
            assert.deepStrictEqual(out, [out[0], out[0], out[0]], `Disclosure sub-document: ${name}, ${format} stable over three passes`);
            assert.deepStrictEqual(text, text.map(() => text[0]), `Disclosure sub-document: ${name}, the same text through ${format}`);
            assert.ok(!token.test(out[0]), `Disclosure sub-document: ${name}, no piece token in ${format}: ${out[0]}`);
        }
    }
    const raw = (await OfficeParser.parseOffice(Buffer.from(cases[6][2]), { fileType: 'md' })).content[0].children!;
    assert.strictEqual(nodeText(raw[0]), 'Raw *not md* line', 'Disclosure sub-document: raw HTML after the summary keeps its Markdown literal, as GitHub shows it');
    // 20 disclosures nested: 16 read as nested, deeper ones a disclosure of their summary followed by their body; all text kept, stable.
    let nested = '';
    for (let i = 0; i < 20; i++) nested += `<details>\n<summary>L${i}</summary>\n\nText ${i}\n\n`;
    nested += '</details>\n\n'.repeat(20) + 'After\n';
    const nestedAst = await OfficeParser.parseOffice(Buffer.from(nested), { fileType: 'md' });
    let level = nestedAst.content[0];
    for (let i = 0; i < 16; i++) {
        assert.ok(level.type === 'disclosure' && (level.metadata as { summary?: string }).summary === `L${i}` && nodeText(level.children![0]) === `Text ${i}`, `Disclosure sub-document: level ${i} nested`);
        level = level.children![level.children!.length - 1];
    }
    const allText = Array.from({ length: 20 }, (_, i) => `L${i}\nText ${i}`).join('\n') + '\nAfter';
    assert.strictEqual((await nestedAst.to('text')).value, allText, 'Disclosure sub-document: past 16 levels nothing dropped');
    const nestedMd = (await nestedAst.to('md')).value;
    const { out: nestedPasses, text: nestedTexts } = await passesWithText(nestedMd, 'md', 'md');
    assert.deepStrictEqual(nestedPasses, [nestedMd, nestedMd, nestedMd], 'Disclosure sub-document: 20 nested stable over three passes');
    assert.ok(nestedTexts.every(t => t === allText), 'Disclosure sub-document: 20 nested keep their text over three passes');
    // Linear: each against plain paragraphs of the same size, in the same run.
    const n = 4000;
    const plain = (length: number) => 'Some *plain* text here.\n\n'.repeat(Math.ceil(length / 25));
    for (const [name, source] of [
        ['inline closers and openers', ('<details>\n\n' + 'a </details> b\n\n').repeat(n)],
        ['nested blocks', '<details><summary>s</summary>\n\nx\n\n'.repeat(n) + '</details>\n\n'.repeat(n)],
        ['openers in quotes', '<details>\n\n' + '> <details>\n>\n> x\n\n'.repeat(n)],
        ['tables of Markdown cells', '<details>\n\n' + '<table><tr><td>\n\nc *x*\n\n</td></tr></table>\n\n'.repeat(n) + '</details>\n'],
    ]) {
        const { within, time, plain: baseline } = await linearAgainst(() => OfficeParser.parseOffice(Buffer.from(source), { fileType: 'md' }),
            () => OfficeParser.parseOffice(Buffer.from(plain(source.length)), { fileType: 'md' }));
        assert.ok(within, `Disclosure sub-document: ${name} (${source.length} bytes) in ${time.toFixed(0)} ms, paragraphs in ${baseline.toFixed(0)} ms`);
    }
    console.log('✓ Markdown disclosure as a sub-document');
}

/**
 * Markdown with no disclosure reads and writes as before disclosures were read as sub-documents: the
 * Markdown test files and a seeded fuzz of every other block kind, tree, metadata, Markdown, HTML and
 * text, against the fingerprints kept in helpers/markdownFingerprints.json.
 */
export async function testMarkdownWithoutDisclosuresUnchanged(): Promise<void> {
    const documents = markdownWithoutDisclosures();
    assert.ok(documents.every(document => !/<details/i.test(document)), 'Markdown without disclosures: no document holds one');
    const fingerprints = await markdownFingerprints((buffer, config) => OfficeParser.parseOffice(buffer, config), documents);
    if (process.env.WRITE_MARKDOWN_FINGERPRINTS) {
        fs.writeFileSync(MARKDOWN_FINGERPRINTS_FILE, JSON.stringify({ from: 'WRITE_MARKDOWN_FINGERPRINTS', fingerprints }, null, 1) + '\n');
        console.log(`✓ Markdown without disclosures: ${fingerprints.length} fingerprints written`);
        return;
    }
    const kept: string[] = JSON.parse(fs.readFileSync(MARKDOWN_FINGERPRINTS_FILE, 'utf8')).fingerprints;
    assert.strictEqual(fingerprints.length, kept.length, 'Markdown without disclosures: as many documents as fingerprints');
    const changed = fingerprints.map((fingerprint, i) => fingerprint === kept[i] ? -1 : i).filter(i => i >= 0);
    assert.deepStrictEqual(changed, [], `Markdown without disclosures: read as before, but ${changed.length} differ, the first ${JSON.stringify(documents[changed[0]]?.slice(0, 300))}`);
    console.log(`✓ Markdown without disclosures unchanged (${documents.length} documents)`);
}

/** Markdown read and written back as Markdown, and its tree's shape (see treeShape). */
async function markdownRead(markdown: string): Promise<{ shape: string; md: string; text: string }> {
    const ast = await OfficeParser.parseOffice(Buffer.from(markdown), { fileType: 'md' });
    return { shape: treeShape(ast.content), md: (await ast.to('md')).value, text: (await ast.to('text')).value };
}

/**
 * A `<pre>` block in a disclosure is HTML from its first line, as CommonMark's condition 1 has it, whether
 * or not the block closes it: a `</details>` in it closes the disclosure there, as a browser has it, and
 * nothing it holds is read as Markdown.
 */
export async function testMarkdownDisclosurePreBlock(): Promise<void> {
    const open = await markdownRead('<details>\n<summary>S</summary>\n\n<pre>\n</details>\n\nstill pre\n</pre>\n\nBody\n\n</details>\n\nAfter\n');
    assert.strictEqual(open.shape, 'disclosure[S](code),paragraph(text),paragraph(text),paragraph(text)', 'Disclosure <pre>: a closing tag in a <pre> it does not close closes the disclosure');
    assert.strictEqual(open.md, '<details>\n<summary>S</summary>\n\n```\n\n```\n\n</details>\n\nstill pre &lt;/pre>\n\nBody\n\nAfter', 'Disclosure <pre>: what follows stands after it');
    const closed = await markdownRead('<details>\n<summary>S</summary>\n\n<pre>\nline\n</pre>\n\n</details>\n\nAfter\n');
    assert.strictEqual(closed.md, '<details>\n<summary>S</summary>\n\n```\nline\n\n```\n\n</details>\n\nAfter', 'Disclosure <pre>: a closed <pre> reads as before');
    console.log('✓ Markdown disclosure <pre> block');
}

/**
 * A disclosure's or summary's tag is read whole, an attribute's quoted value holding `>` included, as a
 * browser reads HTML and CommonMark reads raw HTML in a paragraph; a tag never closed holds the rest of
 * its HTML block. Linear in the number of unclosed values.
 */
export async function testMarkdownDisclosureTagAttributes(): Promise<void> {
    const attribute = await markdownRead('<details title="a>b">\n<summary>S</summary>\n\nBody\n\n</details>\n\nAfter\n');
    assert.strictEqual(attribute.shape, 'disclosure[S](paragraph(text)),paragraph(text)', 'Disclosure tag: a quoted value with > in a block tag');
    assert.strictEqual(attribute.text, 'S\nBody\nAfter', 'Disclosure tag: no attribute text read as body');
    const inline = await markdownRead('Text <details title="a>b"> inline\n\nBody\n\n</details> tail\n');
    assert.strictEqual(inline.shape, 'paragraph(text),disclosure[](paragraph(text),paragraph(text)),paragraph(text)', 'Disclosure tag: a quoted value with > inline');
    assert.strictEqual(inline.text, 'Text\nDetails\ninline\nBody\ntail', 'Disclosure tag: inline, its text in order');
    const upper = await markdownRead('<DETAILS OPEN>\n<SUMMARY>S</SUMMARY>\n\nBody\n\n</DETAILS >\n\nAfter\n');
    assert.strictEqual(upper.shape, 'disclosure[S](paragraph(text)),paragraph(text)', 'Disclosure tag: upper case');
    const n = 4000;
    const plain = (length: number) => 'Some *plain* text here.\n\n'.repeat(Math.ceil(length / 25));
    for (const [name, source] of [
        ['unclosed values in HTML', '<details title="'.repeat(n) + '\n\nx\n'],
        ['unclosed values inline', 'x ' + '<details a="'.repeat(n) + '\n\nx\n'],
        ['values with >', '<details title="a>b">\n\nbody\n\n</details>\n\n'.repeat(n)],
    ]) {
        const { within, time, plain: baseline } = await linearAgainst(() => OfficeParser.parseOffice(Buffer.from(source), { fileType: 'md' }),
            () => OfficeParser.parseOffice(Buffer.from(plain(source.length)), { fileType: 'md' }));
        assert.ok(within, `Disclosure tag: ${name} (${source.length} bytes) in ${time.toFixed(0)} ms, paragraphs in ${baseline.toFixed(0)} ms`);
    }
    console.log('✓ Markdown disclosure tags with attributes');
}

/** ODF namespaces of a note test file. */
const ODF_NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/"';

/**
 * An ODT of notes: `styles` the document's `office:styles` (its notes configurations), `body` its text,
 * `auto` its automatic styles (section styles), `header` the page style's header. A page style is given,
 * as LibreOffice needs one to number by a configuration.
 */
function notesOdt(styles: string, body: string, auto = '', header = ''): Buffer {
    return Buffer.from(zipSync({
        mimetype: [strToU8('application/vnd.oasis.opendocument.text'), { level: 0 }],
        'META-INF/manifest.xml': strToU8('<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/></manifest:manifest>'),
        'styles.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${ODF_NS} office:version="1.2"><office:styles>${styles}</office:styles><office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1">${header}</style:master-page></office:master-styles></office:document-styles>`),
        'content.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ODF_NS} office:version="1.2"><office:automatic-styles>${auto}</office:automatic-styles><office:body><office:text>${body}</office:text></office:body></office:document-content>`),
    }));
}

/** Every note's `kind:mark` in reading order, a note's own notes after it. */
function noteMarksOf(nodes: OfficeContentNode[] | undefined, out: string[] = []): string[] {
    for (const node of nodes ?? []) {
        for (const note of node.notes ?? []) {
            out.push(`${(note.metadata as NoteMetadata).noteType?.[0]}:${(note.metadata as NoteMetadata).mark}`);
            noteMarksOf(note.children, out);
        }
        noteMarksOf(node.children, out);
    }
    return out;
}

/**
 * An ODT note's mark is its number as LibreOffice shows it (LibreOffice 7.3 converting each file to HTML,
 * one at a time): only `office:styles`' configurations are the document's, an empty format is decimal,
 * a section whose own configuration gives a start value numbers its notes on its own (sections inside it
 * inheriting), letters run aa, ab (aa, bb with letter sync, the letter repeated as far as 64 letters,
 * digits past that), numerals repeat `m` past 3999. A note in a text box, drawing, header, footer,
 * comment or note keeps its cached citation and takes no number from the text's. Linear in the notes.
 */
export async function testOdtNoteMarksAsLibreOffice(): Promise<void> {
    let k = 0;
    const note = (kind: 'footnote' | 'endnote', citation = '9', label?: string) => `<text:note text:id="n${++k}" text:note-class="${kind}"><text:note-citation${label === undefined ? '' : ` text:label="${label}"`}>${citation}</text:note-citation><text:note-body><text:p>${kind[0].toUpperCase()}${k}</text:p></text:note-body></text:note>`;
    const fn = (citation?: string, label?: string) => note('footnote', citation, label);
    const en = (citation?: string, label?: string) => note('endnote', citation, label);
    const cfg = (kind: string, attrs: string) => `<text:notes-configuration text:note-class="${kind}" ${attrs}/>`;
    const p = (text: string) => `<text:p>${text}</text:p>`;
    const sect = (name: string, inner: string, style = 'Sect1') => `<text:section text:style-name="${style}" text:name="${name}">${inner}</text:section>`;
    const sectStyle = (configs: string, name = 'Sect1') => `<style:style style:name="${name}" style:family="section"><style:section-properties>${configs}</style:section-properties></style:style>`;
    const many = (count: number, make: () => string) => p(Array.from({ length: count }, (_, i) => `w${i}${make()}`).join(''));
    const letters = (from: number, count: number, upper = false) => Array.from({ length: count }, (_, i) => {
        let text = '';
        for (let rest = from + i; rest > 0; rest = Math.floor((rest - 1) / 26)) text = String.fromCharCode(97 + ((rest - 1) % 26)) + text;
        return upper ? text.toUpperCase() : text;
    });
    const repeated = (n: number, upper = false) => String.fromCharCode((upper ? 65 : 97) + ((n - 1) % 26)).repeat(Math.ceil(n / 26));
    // Each: its files' parts, and the marks LibreOffice shows (where it shows none, what is kept and why).
    const cases: [string, () => Buffer, string[]][] = [
        // Section configurations (in the section style's properties).
        ['secCollect', () => notesOdt('', p(`a${en()}b${fn()}`) + sect('S1', p(`c${en()}d${fn()}`)) + p(`e${en()}f${fn()}`), sectStyle(cfg('footnote', '') + cfg('endnote', ''))), ['e:i', 'f:1', 'e:ii', 'f:2', 'e:iii', 'f:3']],
        ['secPlain', () => notesOdt('', p(`a${en()}b${fn()}`) + sect('S1', p(`c${en()}d${fn()}`)) + p(`e${en()}f${fn()}`), '<style:style style:name="Sect1" style:family="section"><style:section-properties fo:background-color="#eeeeee"/></style:style>'), ['e:i', 'f:1', 'e:ii', 'f:2', 'e:iii', 'f:3']],
        ['sectionEn', () => notesOdt(cfg('endnote', 'style:num-format="i" text:start-value="0"'), p(`a${en()}`) + sect('S1', p(`b${en()}c${en()}`)) + p(`d${en()}`), sectStyle(cfg('endnote', 'style:num-format="A" text:start-value="0"'))), ['e:i', 'e:A', 'e:B', 'e:ii']],
        ['secRestart', () => notesOdt('', p(`a${en()}`) + sect('S1', p(`b${en()}c${en()}`)) + p(`d${en()}`), sectStyle(cfg('endnote', 'text:start-value="0"'))), ['e:i', 'e:i', 'e:ii', 'e:ii']],
        ['secRestart1', () => notesOdt('', p(`a${en()}`) + sect('S1', p(`b${en()}c${en()}`)) + p(`d${en()}`), sectStyle(cfg('endnote', 'text:start-value="1"'))), ['e:i', 'e:i', 'e:ii', 'e:ii']],
        ['secRestart2', () => notesOdt('', p(`a${en()}`) + sect('S1', p(`b${en()}c${en()}`)) + p(`d${en()}`), sectStyle(cfg('endnote', 'text:start-value="2"'))), ['e:i', 'e:ii', 'e:iii', 'e:ii']],
        ['secRestart3', () => notesOdt('', p(`a${en()}`) + sect('S1', p(`b${en()}c${en()}`)) + p(`d${en()}`), sectStyle(cfg('endnote', 'text:start-value="3"'))), ['e:i', 'e:iii', 'e:iv', 'e:ii']],
        ['secFmtOnly', () => notesOdt('', p(`a${en()}`) + sect('S1', p(`b${en()}c${en()}`)) + p(`d${en()}`), sectStyle(cfg('endnote', 'style:num-format="A"'))), ['e:i', 'e:ii', 'e:iii', 'e:iv']],
        ['secPrefixOnly', () => notesOdt('', p(`a${en()}`) + sect('S1', p(`b${en()}c${en()}`)) + p(`d${en()}`), sectStyle(cfg('endnote', 'style:num-prefix="["'))), ['e:i', 'e:ii', 'e:iii', 'e:iv']],
        ['secDocFmt', () => notesOdt(cfg('endnote', 'style:num-format="I"'), p(`a${en()}`) + sect('S1', p(`b${en()}c${en()}`)) + p(`d${en()}`), sectStyle(cfg('endnote', 'text:start-value="0"'))), ['e:I', 'e:I', 'e:II', 'e:II']],
        ['secFnOwn', () => notesOdt(cfg('footnote', 'style:num-format="1" text:start-numbering-at="document"'), p(`a${fn()}`) + sect('S1', p(`b${fn()}c${fn()}`)) + p(`d${fn()}`), sectStyle(cfg('footnote', 'style:num-format="A" text:start-value="0"'))), ['f:1', 'f:A', 'f:B', 'f:2']],
        ['secFnRestart', () => notesOdt(cfg('footnote', 'style:num-format="1" text:start-numbering-at="document"'), p(`a${fn()}`) + sect('S1', p(`b${fn()}c${fn()}`)) + p(`d${fn()}`), sectStyle(cfg('footnote', 'text:start-value="0"'))), ['f:1', 'f:1', 'f:2', 'f:2']],
        // Footnotes numbered per page: a section numbering its own is numbered, the rest keep their cached citations (LibreOffice: 0 a b 0 on one page).
        ['secFnPage', () => notesOdt(cfg('footnote', 'style:num-format="1"'), p(`a${fn('7')}`) + sect('S1', p(`b${fn()}c${fn()}`)) + p(`d${fn('8')}`), sectStyle(cfg('footnote', 'style:num-format="a" text:start-value="0"'))), ['f:7', 'f:a', 'f:b', 'f:8']],
        ['secNested', () => notesOdt('', p(`a${en()}`) + sect('S1', p(`b${en()}`) + sect('S2', p(`c${en()}`), 'Sect2') + p(`d${en()}`)) + p(`e${en()}`), sectStyle(cfg('endnote', 'style:num-format="A" text:start-value="0"')) + sectStyle('', 'Sect2')), ['e:i', 'e:A', 'e:A', 'e:B', 'e:ii']],
        ['secNestedOwn', () => notesOdt('', p(`a${en()}`) + sect('S1', p(`b${en()}`) + sect('S2', p(`c${en()}`), 'Sect2') + p(`d${en()}`)) + p(`e${en()}`), sectStyle(cfg('endnote', 'style:num-format="A" text:start-value="0"')) + sectStyle(cfg('endnote', 'style:num-format="a" text:start-value="0"'), 'Sect2')), ['e:i', 'e:A', 'e:a', 'e:B', 'e:ii']],
        ['secNestedCollect', () => notesOdt('', p(`a${en()}`) + sect('S1', p(`b${en()}`) + sect('S2', p(`c${en()}`), 'Sect2') + p(`d${en()}`)) + p(`e${en()}`), sectStyle(cfg('endnote', 'style:num-format="A" text:start-value="0"')) + sectStyle(cfg('endnote', ''), 'Sect2')), ['e:i', 'e:A', 'e:B', 'e:C', 'e:ii']],
        ['secTwo', () => notesOdt('', p(`a${en()}`) + sect('S1', p(`b${en()}c${en()}`)) + p(`d${en()}`) + sect('S3', p(`e${en()}`)) + p(`f${en()}`), sectStyle(cfg('endnote', 'style:num-format="A" text:start-value="0"'))), ['e:i', 'e:A', 'e:B', 'e:ii', 'e:A', 'e:iii']],
        ['secAlpha', () => notesOdt('', p(`a${en()}`) + sect('S1', many(28, () => en())) + p(`d${en()}`), sectStyle(cfg('endnote', 'style:num-format="a" text:start-value="0"'))), ['e:i', ...letters(1, 28).map(m => `e:${m}`), 'e:ii']],
        // The document's configuration.
        ['fa2', () => notesOdt(cfg('footnote', 'style:num-format="a" text:start-value="2" text:start-numbering-at="document"') + cfg('endnote', 'style:num-format="A" text:start-value="0"'), p(`a${fn()}b${en()}c${fn()}d${en()}`)), ['f:c', 'e:A', 'f:d', 'e:B']],
        ['fa2b', () => notesOdt('<text:notes-configuration text:note-class="footnote" text:citation-style-name="FootnoteReference" text:citation-body-style-name="FootnoteReference" text:start-value="2" style:num-format="a" text:start-numbering-at="document" text:footnotes-position="page"/><text:notes-configuration text:note-class="endnote" text:citation-style-name="EndnoteReference" text:citation-body-style-name="EndnoteReference" text:start-value="0" style:num-format="A" text:start-numbering-at="document" text:footnotes-position="document"/>', p(`a${fn()}b${en()}c${fn()}d${en()}`)), ['f:c', 'e:A', 'f:d', 'e:B']],
        ['fa2c', () => notesOdt(cfg('footnote', 'style:num-format="a" text:start-value="2" text:start-numbering-at="document" text:footnotes-position="page"') + cfg('endnote', 'style:num-format="A" text:start-value="0" text:footnotes-position="document"'), p(`a${fn()}b${en()}c${fn()}d${en()}`)), ['f:c', 'e:A', 'f:d', 'e:B']],
        ['none', () => notesOdt('', p(`a${fn()}b${en()}c${fn()}d${en()}e${fn()}`)), ['f:1', 'e:i', 'f:2', 'e:ii', 'f:3']],
        ['enNoFmt', () => notesOdt(cfg('endnote', ''), p(`a${en()}b${en()}`)), ['e:1', 'e:2']],
        ['label', () => notesOdt('', p(`a${fn()}b${fn('*', '*')}c${fn()}d${en('x', 'x')}e${en()}`)), ['f:1', 'f:*', 'f:2', 'e:x', 'e:i']],
        ['prefix', () => notesOdt(cfg('footnote', 'style:num-format="1" style:num-prefix="[" style:num-suffix="]" text:start-numbering-at="document"'), p(`a${fn()}b${fn()}`)), ['f:1', 'f:2']],
        ['startVal1', () => notesOdt(cfg('footnote', 'style:num-format="1" text:start-value="1" text:start-numbering-at="document"') + cfg('endnote', 'style:num-format="I" text:start-value="4"'), p(`a${fn()}b${fn()}c${en()}`)), ['f:2', 'f:3', 'e:V']],
        ['emptyFmt', () => notesOdt(cfg('footnote', 'style:num-format="" text:start-numbering-at="document"'), p(`a${fn('5')}b${fn('6')}`)), ['f:1', 'f:2']],
        ['emptyFmtEn', () => notesOdt(cfg('endnote', 'style:num-format=""'), p(`a${en('5')}b${en('6')}`)), ['e:1', 'e:2']],
        ['table', () => notesOdt('', p(`a${fn()}`) + `<table:table table:name="T"><table:table-column/><table:table-row><table:table-cell>${p(`c${fn()}`)}</table:table-cell></table:table-row></table:table>` + p(`d${fn()}`)), ['f:1', 'f:2', 'f:3']],
        ['heading', () => notesOdt('', `<text:h text:outline-level="1">H${fn()}</text:h><text:list><text:list-item>${p(`l${fn()}`)}</text:list-item></text:list>` + p(`d${fn()}`)), ['f:1', 'f:2', 'f:3']],
        // Letters and numerals.
        ['alpha30', () => notesOdt(cfg('footnote', 'style:num-format="a" text:start-value="0" text:start-numbering-at="document"'), many(30, () => fn())), letters(1, 30).map(m => `f:${m}`)],
        ['upper30', () => notesOdt(cfg('footnote', 'style:num-format="A" text:start-numbering-at="document"'), many(30, () => fn())), letters(1, 30, true).map(m => `f:${m}`)],
        ['enAlpha30', () => notesOdt(cfg('endnote', 'style:num-format="a"'), many(30, () => en())), letters(1, 30).map(m => `e:${m}`)],
        ['alpha700', () => notesOdt(cfg('footnote', 'style:num-format="a" text:start-value="700" text:start-numbering-at="document"'), p(`a${fn()}b${fn()}`)), ['f:zy', 'f:zz']],
        ['alphaSync', () => notesOdt(cfg('footnote', 'style:num-format="a" style:num-letter-sync="true" text:start-value="0" text:start-numbering-at="document"'), many(30, () => fn())), Array.from({ length: 30 }, (_, i) => `f:${repeated(i + 1)}`)],
        ['alphaSync300', () => notesOdt(cfg('footnote', 'style:num-format="a" style:num-letter-sync="true" text:start-value="300" text:start-numbering-at="document"'), p(`a${fn()}b${fn()}`)), ['f:oooooooooooo', 'f:pppppppppppp']],
        ['alphaSync1661', () => notesOdt(cfg('footnote', 'style:num-format="a" style:num-letter-sync="true" text:start-value="1661" text:start-numbering-at="document"'), p(`a${fn()}b${fn()}c${fn()}`)), ['f:' + 'x'.repeat(64), 'f:' + 'y'.repeat(64), 'f:' + 'z'.repeat(64)]],
        // LibreOffice writes the letter 65 times for note 1665; past 64 letters the mark is digits.
        ['alphaSync1662', () => notesOdt(cfg('footnote', 'style:num-format="A" style:num-letter-sync="true" text:start-value="1662" text:start-numbering-at="document"'), p(`a${fn()}b${fn()}c${fn()}`)), ['f:' + 'Y'.repeat(64), 'f:' + 'Z'.repeat(64), 'f:1665']],
        ['romanBig', () => notesOdt(cfg('footnote', 'style:num-format="i" text:start-value="3998" text:start-numbering-at="document"'), p(`a${fn()}b${fn()}c${fn()}`)), ['f:mmmcmxcix', 'f:mmmm', 'f:mmmmi']],
        ['roman5k', () => notesOdt(cfg('footnote', 'style:num-format="I" text:start-value="4998" text:start-numbering-at="document"'), p(`a${fn()}b${fn()}`)), ['f:MMMMCMXCIX', 'f:MMMMM']],
        // Notes outside the text LibreOffice numbers: LibreOffice drops a text box's note (1 2) and opens no
        // file with a note in a header, comment or note; theirs keep the cached citation, nothing lost.
        ['frame', () => notesOdt('', p(`a${fn()}<draw:frame text:anchor-type="as-char" svg:width="3cm" svg:height="1cm"><draw:text-box>${p(`t${fn()}`)}</draw:text-box></draw:frame>z${fn()}`)), ['f:1', 'f:9', 'f:2']],
        ['framePara', () => notesOdt('', p(`a${fn()}`) + p(`<draw:frame text:anchor-type="paragraph" svg:width="3cm" svg:height="1cm"><draw:text-box>${p(`t${fn()}`)}</draw:text-box></draw:frame>m`) + p(`z${fn()}`)), ['f:1', 'f:9', 'f:2']],
        ['header', () => notesOdt('', p(`a${fn()}b${fn()}`), '', `<style:header>${p(`h${fn()}`)}</style:header>`), ['f:1', 'f:2']],
        ['annot', () => notesOdt('', p(`a${fn()}<office:annotation><dc:creator>x</dc:creator>${p(`c${fn()}`)}</office:annotation>b${fn()}`)), ['f:1', 'f:2']],
        ['noteInNote', () => notesOdt('', p(`a<text:note text:id="nn" text:note-class="footnote"><text:note-citation>9</text:note-citation><text:note-body><text:p>outer${fn()}</text:p></text:note-body></text:note>b${fn()}`)), ['f:1', 'f:9', 'f:2']],
    ];
    for (const [name, make, expected] of cases) {
        k = 0;
        const ast = await OfficeParser.parseOffice(make(), { fileType: 'odt' });
        assert.deepStrictEqual(noteMarksOf(ast.content), expected, `ODT note marks as LibreOffice: ${name}`);
    }
    // The text box's note is kept, with its cached citation and its text.
    k = 0;
    const frame = await OfficeParser.parseOffice(cases.find(c => c[0] === 'framePara')![1](), { fileType: 'odt' });
    assert.ok(/--- Notes ---\nF1\nF2\nF3/.test((await frame.to('text')).value), 'ODT note marks as LibreOffice: a text box\'s note kept');
    // Linear: many notes in nested sections and in text boxes, against as many in one paragraph.
    const many_ = 8000;
    const endnotes = (count: number) => Array.from({ length: count }, (_, i) => `w<text:note text:id="n${i}" text:note-class="endnote"><text:note-citation>9</text:note-citation><text:note-body><text:p>E</text:p></text:note-body></text:note>`).join('');
    const ownNumbering = sectStyle(cfg('endnote', 'text:start-value="0" style:num-format="a"'), 'S');
    for (const [name, body] of [
        ['nested sections', '<text:section text:style-name="S" text:name="s">'.repeat(100) + p(endnotes(many_)) + '</text:section>'.repeat(100)],
        ['text boxes', p(Array.from({ length: many_ }, (_, i) => `<draw:frame svg:width="1cm"><draw:text-box><text:p>x<text:note text:id="n${i}" text:note-class="endnote"><text:note-citation>9</text:note-citation><text:note-body><text:p>E</text:p></text:note-body></text:note></text:p></draw:text-box></draw:frame>`).join(''))],
    ]) {
        const file = notesOdt('', body, ownNumbering);
        const plain = notesOdt('', p(endnotes(many_)) + p('x'.repeat(body.length - endnotes(many_).length)));
        const { within, time, plain: baseline } = await linearAgainst(() => OfficeParser.parseOffice(file, { fileType: 'odt' }), () => OfficeParser.parseOffice(plain, { fileType: 'odt' }));
        assert.ok(within, `ODT note marks as LibreOffice: ${many_} notes in ${name} in ${time.toFixed(0)} ms, in one paragraph ${baseline.toFixed(0)} ms`);
    }
    console.log('✓ ODT note marks as LibreOffice shows them');
}

/**
 * The XHTML sniffer reads the start of an XML document, twice as much each time its prolog runs past
 * it, so its answer is the whole document's: a comment, instruction or internal subset longer than the
 * first read, or a root name across it, still finds an XHTML page; an `<html>` in an entity value of a
 * long subset is no root. Linear in the prolog's length.
 */
export async function testXhtmlSniffedPastFirstRead(): Promise<void> {
    const K = 64 * 1024;
    const decl = '<?xml version="1.0"?>';
    const page = '<html xmlns="http://www.w3.org/1999/xhtml"><body><p>A</p></body></html>';
    const pad = (prefix: string, filler: (n: number) => string, suffix: string) => prefix + filler(Math.max(0, K - prefix.length - 3)) + suffix;
    const readAs = async (source: string): Promise<string> => {
        try {
            const ast = await OfficeParser.parseOffice(Buffer.from(source));
            return `${ast.type}:${(await ast.to('text')).value.trim().split('\n')[0]}`;
        } catch (error: any) {
            return /supports/.test(error.message) ? 'refused' : `error ${error.message}`;
        }
    };
    for (const [name, source, expected] of [
        ['a comment across the first read', decl + '<!--' + 'c'.repeat(K) + '-->' + page, 'html:A'],
        ['an instruction across it', decl + '<?pi ' + 'p'.repeat(K) + '?>' + page, 'html:A'],
        ['a subset across it', decl + '<!DOCTYPE html [<!ENTITY e "' + 'x'.repeat(K) + '">]>' + page, 'html:A'],
        ['an <html> in an entity value of a long subset', decl + '<!DOCTYPE r [ <!ENTITY x "a>  <html>"><!ENTITY e "' + 'x'.repeat(70000) + '"> ]><r>data</r>', 'refused'],
        ['the same, the entity last', decl + '<!DOCTYPE r [<!ENTITY e "' + 'x'.repeat(70000) + '"><!ENTITY x "a> <html>"> ]><r>data</r>', 'refused'],
        ['a root name across it', pad(decl, n => ' '.repeat(n), page), 'html:A'],
        ['white space to it, then another root', pad(decl, n => ' '.repeat(n + 10), '<r>data</r>'), 'refused'],
        ['a plain DOCTYPE and a long body', decl + '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "x.dtd">' + page.replace('<p>A</p>', '<p>A</p>' + '<p>b</p>'.repeat(20000)), 'html:A'],
        ['a malformed subset, then the page', decl + '<!DOCTYPE html [ junk ' + 'x'.repeat(K) + ' >' + page, 'html:A'],
        ['an unclosed comment', decl + '<!--' + 'c'.repeat(2 * K), 'refused'],
    ]) {
        assert.strictEqual(await readAs(source), expected, `XHTML sniffing: ${name}`);
    }
    // Linear: a prolog four times as long takes about four times as long, not sixteen.
    const prolog = (bytes: number) => Buffer.from(decl + '<?a b?>'.repeat(Math.floor(bytes / 7)) + '<r/>');
    const subset = (bytes: number) => Buffer.from(decl + '<!DOCTYPE r [' + '<!ENTITY a "x">'.repeat(Math.floor(bytes / 15)));
    for (const [name, make] of [['instructions', prolog], ['an open subset', subset]] as const) {
        const small = make(2 * 1024 * 1024);
        const big = make(8 * 1024 * 1024);
        const refuse = (buffer: Buffer) => OfficeParser.parseOffice(buffer).then(() => assert.fail(`XHTML sniffing: ${name} read`), () => undefined);
        const { within, time, plain } = await linearAgainst(() => refuse(big), async () => { for (let i = 0; i < 4; i++) await refuse(small); });
        assert.ok(within, `XHTML sniffing: ${name}, 8 MB in ${time.toFixed(0)} ms, 2 MB four times in ${plain.toFixed(0)} ms`);
    }
    console.log('✓ XHTML sniffed past its first read');
}

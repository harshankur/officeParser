/**
 * Document shapes kept across readers and writers: XHTML read as HTML, list types per level, note
 * marks, quotes of several paragraphs, figure captions, disclosures and keyboard input, and sheets
 * written as document tables.
 */
import * as assert from 'assert';
import { unzipSync } from 'fflate';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParser } from '../../src/OfficeParser';
import type { ImageMetadata, NoteMetadata, OfficeContentNode, OfficeParserAST, ParagraphMetadata } from '../../src/types';

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

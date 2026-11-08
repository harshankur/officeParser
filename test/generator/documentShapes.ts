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
import { referenceHasDisclosureElement } from '../helpers/disclosureReference';
import { runFigureLineGenerator } from '../helpers/figureLineGenerator';
import { PICTURE_MATRIX_FINGERPRINTS_FILE, type PictureMatrixIo, runPictureMatrix } from '../helpers/pictureMatrixGenerator';

/** The text of a node: a line's runs joined, blocks set apart by a space. */
const wordsOfNode = (node: OfficeContentNode): string => node.children?.length
    ? node.children.map(wordsOfNode).join(['paragraph', 'heading', 'text'].includes(node.type) ? '' : ' ')
    : node.text ?? '';

/** A disclosure's summary as text (its first child, a `summary` node), white space collapsed; '' where it has none. */
const summaryText = (node: OfficeContentNode): string => (node.children?.[0]?.type === 'summary' ? wordsOfNode(node.children[0]).replace(/\s+/g, ' ').trim() : '');

/** A disclosure's body: its children after its summary. */
const bodyOf = (node: OfficeContentNode): OfficeContentNode[] => (node.children?.[0]?.type === 'summary' ? node.children.slice(1) : node.children ?? []);

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
    assert.deepStrictEqual(outer.metadata, { open: true }, 'Disclosure: its open state');
    assert.deepStrictEqual(JSON.parse(JSON.stringify(outer.children![0])), { type: 'summary', children: [{ type: 'paragraph', children: [{ type: 'text', text: 'Click me' }] }] }, 'Disclosure: its summary, its first child');
    assert.deepStrictEqual(outer.children!.map(c => c.type), ['summary', 'paragraph', 'list', 'list', 'code', 'disclosure'], 'Disclosure: its body read as Markdown, a nested one too');
    const keys = ast.content[2].children!.filter(c => c.formatting?.keyboard).map(c => [c.text, c.formatting?.font]);
    assert.deepStrictEqual(keys, [['Ctrl', 'monospace'], ['*C*', 'monospace']], 'Keyboard: <kbd> is keyboard input, monospace too');
    // Markdown to editor HTML and back, three times: the same Markdown, kbd's Markdown characters as references.
    const expected = source.replace('<kbd>*C*</kbd>', '<kbd>&#42;C&#42;</kbd>');
    const { html, md } = await markdownThroughHtml(source);
    assert.deepStrictEqual(md, [expected, expected, expected], 'Disclosure: Markdown to HTML to Markdown is stable over three passes');
    assert.ok(/<details open=""><summary>Click me<\/summary>\s*<p>Some <i>body<\/i> text\.<\/p>/.test(html[0]), `Disclosure: HTML writes <details>: ${html[0]}`);
    assert.ok(/<details><summary>Inner<\/summary>\s*<p>Nested\.<\/p>/.test(html[0]), 'Disclosure: a nested one, closed');
    assert.ok(html[0].includes('Press <kbd>Ctrl</kbd>+<kbd>*C*</kbd> to copy.'), 'Keyboard: HTML writes <kbd>');
    // The HTML reader: loose text in a disclosure is a paragraph; one with no summary has none.
    const loose = await OfficeParser.parseOffice(Buffer.from('<details><summary>S</summary>Loose <b>text</b></details><details><p>x</p></details>'), { fileType: 'html' });
    assert.deepStrictEqual(loose.content.map(n => [n.type, n.metadata, n.children?.map(c => c.type)]), [
        ['disclosure', {}, ['summary', 'paragraph']], ['disclosure', {}, ['paragraph']],
    ], 'Disclosure: HTML read as disclosures');
    assert.strictEqual(summaryText(loose.content[0]), 'S', 'Disclosure: HTML, its summary');
    assert.ok((await loose.to('html', { htmlConfig: { standalone: false } })).value.includes('<details><p>x</p>'), 'Disclosure: none written where none was given');
    // EPUB: XHTML's open attribute, read back.
    const epub = await OfficeParser.parseOffice(Buffer.from((await ast.to('epub')).value as Uint8Array));
    const epubOuter = epub.content.find(n => n.type === 'disclosure');
    assert.deepStrictEqual([epubOuter?.metadata, summaryText(epubOuter!)], [{ open: true }, 'Click me'], 'Disclosure: EPUB keeps it');
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
    // In a list item, whose one line of Markdown holds no block: the item written as an HTML list, read back the same.
    const inItem = await OfficeParser.parseOffice(Buffer.from('<ul><li>x<details><summary>S <b>b</b></summary><p>Body</p></details></li></ul>'), { fileType: 'html' });
    const itemMd = await inItem.to('md');
    assert.strictEqual(itemMd.value, '<ul><li>x<details><summary>S <b>b</b></summary>\n<p>Body</p></details></li></ul>', 'Disclosure: in a list item, the item as HTML');
    assert.deepStrictEqual((await OfficeParser.parseOffice(Buffer.from(itemMd.value), { fileType: 'md' })).content, inItem.content, 'Disclosure: in a list item, read back the same');
    // In a definition, where no block goes either: its summary in bold, then its body, reported.
    const inDefinition = await OfficeGenerator.generate({ type: 'md', metadata: {}, attachments: [], content: [{ type: 'definitionList', children: [{ type: 'definitionTerm', children: [{ type: 'text', text: 'T' }] },
        { type: 'definitionDescription', children: [{ type: 'disclosure', children: [{ type: 'summary', children: [{ type: 'paragraph', children: [{ type: 'text', text: 'S' }] }] }, { type: 'paragraph', children: [{ type: 'text', text: 'Body' }] }] }] }] }] } as unknown as OfficeParserAST, 'md');
    assert.ok(/\*\*S\*\*(?: |<br>)Body/.test(String(inDefinition.value)), `Disclosure: in a definition, its summary in bold: ${inDefinition.value}`);
    assert.ok(inDefinition.messages.some(m => m.code === 'CONTENT_NOT_REPRESENTABLE'), 'Disclosure: and that is reported');
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
 * A disclosure carries no text of its own: its summary is its first child and its body the rest, so a
 * disclosure with no body is not written with its summary as content, and nested ones copy nothing.
 */
export async function testDisclosureCarriesNoText(): Promise<void> {
    const empty = '<details><summary>Only summary</summary></details><p>x</p>';
    const ast = await OfficeParser.parseOffice(Buffer.from(empty), { fileType: 'html' });
    assert.deepStrictEqual(JSON.parse(JSON.stringify(ast.content[0])), { type: 'disclosure', metadata: {}, children: [{ type: 'summary', children: [{ type: 'paragraph', children: [{ type: 'text', text: 'Only summary' }] }] }] }, 'Disclosure: no text of its own');
    assert.strictEqual((await ast.to('text')).value.trim(), 'Only summary\nx', 'Disclosure: text shows its summary once');
    const html = await passesOf(empty, 'html', 'html');
    assert.ok(/<details><summary>Only summary<\/summary>\s*<\/details>/.test(html[0]), `Disclosure: an empty one written empty: ${html[0]}`);
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
        assert.deepStrictEqual(ast.content.map(n => [n.type, n.children?.map(c => c.type)]), [['disclosure', ['summary', 'paragraph']], ['paragraph', ['text']]], `Markdown disclosure: ${body} is one paragraph in it`);
        const html = (await ast.to('html', { htmlConfig: { standalone: false } })).value;
        if (body.includes('`')) assert.ok(/<code>&lt;(details|summary)&gt;<\/code>/.test(html), `Markdown disclosure: the code span kept: ${html}`);
    }
    // Our own Markdown of a disclosure mentioning its tag, in a code span and in code.
    const help = await OfficeParser.parseOffice(Buffer.from('<details><summary>Help</summary><p>Use <code>&lt;details&gt;</code> to fold.</p><pre><code>x\n&lt;/details&gt;\ny</code></pre><p>Inner end</p></details><p>After</p>'), { fileType: 'html' });
    const helpMd = (await help.to('md')).value;
    assert.deepStrictEqual(await markdownShape(helpMd), [['disclosure', ['summary', 'paragraph', 'code', 'paragraph']], ['paragraph', ['text']]], `Markdown disclosure: code mentioning its tag stays in it: ${helpMd}`);
    // A body starting with an inline tag is a paragraph; indented code is code; nested code spans are text.
    const keys = await OfficeParser.parseOffice(Buffer.from('<details>\n<summary>Keys</summary>\n\n<kbd>Ctrl</kbd> and *bold*.\n\n</details>\n\nAfter\n'), { fileType: 'md' });
    const paragraph = keys.content[0].children![1];
    assert.ok(paragraph.type === 'paragraph' && paragraph.children!.some(c => c.formatting?.keyboard) && paragraph.children!.some(c => c.formatting?.italic), 'Markdown disclosure: <kbd> and emphasis in a paragraph');
    assert.deepStrictEqual(await markdownShape('<details>\n<summary>Keys</summary>\n\n    <details> in indented code\n\n</details>\n\nAfter\n'), [['disclosure', ['summary', 'code']], ['paragraph', ['text']]], 'Markdown disclosure: indented code is code');
    assert.deepStrictEqual(await markdownShape('<details>\n<summary>S</summary>\n\nA ``x `<details>` y`` b and `</details>` c.\n\nMore\n\n</details>\n\nAfter\n'),
        [['disclosure', ['summary', 'paragraph', 'paragraph']], ['paragraph', ['text']]], 'Markdown disclosure: tags in code spans are text');
    // A closing tag on a paragraph's last line closes it, as a browser has it.
    assert.deepStrictEqual(await markdownShape('<details>\n<summary>S</summary>\n\n*Body*\n</details>\n\nAfter\n'), [['disclosure', ['summary', 'paragraph']], ['paragraph', ['text']]], 'Markdown disclosure: an inline closing tag closes it');
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
        assert.deepStrictEqual(back.content.map(n => [n.type, n.children?.length]), [['disclosure', count + 1], ['paragraph', 1]], `Markdown disclosure: ${count} paragraphs stay in it, after its summary`);
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

/** A tree's shape: each node's type, a disclosure's summary's text in brackets, and its children's (a disclosure's body's) in parentheses. */
function treeShape(nodes: OfficeContentNode[] | undefined): string {
    return (nodes ?? []).map(n => {
        const children = n.type === 'disclosure' ? bodyOf(n) : n.children;
        return n.type + (n.type === 'disclosure' ? `[${summaryText(n)}]` : '') + (n.type !== 'text' && children?.length ? `(${treeShape(children)})` : '');
    }).join(',');
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
    const SUMMARY_S = { type: 'summary', children: [{ type: 'paragraph', children: [{ type: 'text', text: 'S' }] }] };
    const shapes: [string, unknown][] = [
        ['<details>\n<summary>S</summary>\n\n<div>\n\n*md in div*\n\n</div>\n\nBody\n\n</details>\n\nAfter\n', [
            { type: 'disclosure', metadata: {}, children: [SUMMARY_S,
                { type: 'paragraph', metadata: {}, children: [{ type: 'text', text: 'md in div', formatting: { italic: true } }] },
                { type: 'paragraph', metadata: {}, children: [{ type: 'text', text: 'Body' }] }] },
            { type: 'paragraph', metadata: {}, children: [{ type: 'text', text: 'After' }] }]],
        ['<details>\n<summary>S</summary>\n\n<div align="center">\n\n*centered*\n\n</div>\n\n</details>\n', [
            { type: 'disclosure', metadata: {}, children: [SUMMARY_S,
                { type: 'paragraph', metadata: { alignment: 'center' }, children: [{ type: 'text', text: 'centered', formatting: { italic: true } }] }] }]],
        ['<details>\n<summary>S</summary>\n\n<p align="center">\n\n![x](a.png)\n\n</p>\n\nBody\n\n</details>\n', [
            { type: 'disclosure', metadata: {}, children: [SUMMARY_S,
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
        .map(n => n.type + (n.type === 'disclosure' ? `[${summaryText(n)}](${blockShape(bodyOf(n))})` : '')).join(',');
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
    const raw = bodyOf((await OfficeParser.parseOffice(Buffer.from(cases[6][2]), { fileType: 'md' })).content[0]);
    assert.strictEqual(nodeText(raw[0]), 'Raw *not md* line', 'Disclosure sub-document: raw HTML after the summary keeps its Markdown literal, as GitHub shows it');
    // 20 disclosures nested: 16 read as nested, deeper ones flattened to their summary and body as text; all text kept, stable.
    let nested = '';
    for (let i = 0; i < 20; i++) nested += `<details>\n<summary>L${i}</summary>\n\nText ${i}\n\n`;
    nested += '</details>\n\n'.repeat(20) + 'After\n';
    const nestedAst = await OfficeParser.parseOffice(Buffer.from(nested), { fileType: 'md' });
    let level = nestedAst.content[0];
    for (let i = 0; i < 16; i++) {
        assert.ok(level.type === 'disclosure' && summaryText(level) === `L${i}` && nodeText(bodyOf(level)[0]) === `Text ${i}`, `Disclosure sub-document: level ${i} nested`);
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

/**
 * Markdown with no disclosure reads the same whatever names stand where no tag is read: documents of
 * fragments holding `details` in other tags' attribute values, comments, code spans, code, link tails,
 * escapes, wrappers and lone `<summary>` elements (none a disclosure as GitHub shows it), read again with
 * `details` renamed `detailx`, give the same tree, Markdown and text but for that name. A name read as a
 * tag where a browser reads none, or a reading hanging on it, would part the two.
 */
export async function testMarkdownDisclosureNamesRenamed(): Promise<void> {
    const FRAGMENTS = ['The <abbr title="the <details> element">disclosure</abbr> widget.', '<span title="</details>">end</span> key', '<p title="</details>">x</p>',
        "<span data-x='<details open>'>s</span> t", '- item <abbr title="<details>">a</abbr>\n- two', '> q <b title="</details>">b</b>', '## H <i title="<details>">i</i>',
        '<summary>s</summary>', '<summary>a</summary><summary>b</summary>', 'x <!-- <details> --> y', '<!-- <details> -->', '`<details>` code', '\\<details> esc',
        '<div style="text-align: center">Use `</details>` here</div>', '<div align="center">\n\n*c*\n\n</div>', '| a | <b title="<details>">b</b> |\n| - | - |\n| 1 | 2 |',
        '[x](</details>) link', '[y](u "</details>") t', '<table><tr><td>c <abbr title="<details>">d</abbr></td></tr></table>', '```\n<details>\n```',
        '<div>\n<p title="<details>">p</p>\n</div>', 'Para *em* `c`', '1. one\n2. two', 'Note[^1]', '[^1]: N', 'Title\n===', '---',
        '<section data-x="</details>">s</section>', '<details-like>x</details-like>', '<summary title="<details>">s</summary>', '> quoted\n<summary>lazy</summary>'];
    let seed = 99;
    const rnd = (n: number): number => { seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff; return seed % n; };
    const rename = (text: string): string => text.replace(/(detail)(s)/gi, (_, name: string, s: string) => name + (s === 'S' ? 'X' : 'x'));
    const read = async (markdown: string): Promise<string> => {
        const ast = await OfficeParser.parseOffice(Buffer.from(markdown), { fileType: 'md' });
        return [JSON.stringify(ast.content).replace(/,?"quoteId":"md-quote-\d+"/g, ''), (await ast.to('md')).value, (await ast.to('text')).value].join('\n---\n');
    };
    let documents = 0;
    for (let k = 0; k < 400; k++) {
        const doc = Array.from({ length: 2 + rnd(9) }, () => FRAGMENTS[rnd(FRAGMENTS.length)]).map(p => p + (rnd(3) ? '\n\n' : '\n')).join('');
        // (Joined tight, two fragments can make a disclosure as GitHub shows it: left out.)
        if (referenceHasDisclosureElement(doc)) continue;
        documents++;
        const [named, renamed] = [await read(doc), await read(rename(doc))];
        assert.strictEqual(rename(named), renamed, `Disclosure names renamed: ${JSON.stringify(doc)}`);
    }
    assert.ok(documents > 300, `Disclosure names renamed: ${documents} documents`);
    console.log(`✓ Markdown disclosure names renamed read the same (${documents} documents)`);
}

/**
 * Round 9's findings, each read and written back: a summary of more than text keeps its content (2, 4);
 * a tag's name in another tag's quoted value is no tag (3); the writer's aligned wrapper reads as written
 * (5); text cut after a tag stays the paragraph text it was (6); an enclosing element's end tag closes a
 * disclosure opened in it (7); a summary after a disclosure's first is a block of its own, and a
 * `</summary>` closes a disclosure opened in its summary (8). Then a lone `<summary>` line ends no quote or
 * list item in Markdown with no disclosure, as 8.6.0 read it.
 */
export async function testMarkdownDisclosureRound9Findings(): Promise<void> {
    const settled = async (source: string, fileType: 'md' | 'html', shape: string, name: string): Promise<OfficeParserAST> => {
        const ast = await OfficeParser.parseOffice(Buffer.from(source), { fileType });
        assert.strictEqual(treeShape(ast.content), shape, `Round 9: ${name}`);
        let md = (await ast.to('md')).value;
        for (let pass = 0; pass < 2; pass++) {
            const back = await OfficeParser.parseOffice(Buffer.from(md), { fileType: 'md' });
            assert.strictEqual(treeShape(back.content), shape, `Round 9: ${name}, the same tree after pass ${pass + 1}: ${md}`);
            const again = (await back.to('md')).value;
            if (pass) assert.strictEqual(again, md, `Round 9: ${name}, the Markdown settled`);
            md = again;
        }
        return ast;
    };
    const summaryOf = (ast: OfficeParserAST): OfficeContentNode => ast.content.find(n => n.type === 'disclosure')!.children![0];
    // 2: what a summary holds, block elements and all.
    const blocks = await settled('<details><summary><b>Title</b><div>sub</div></summary><p>body</p></details>', 'html', 'disclosure[Title sub](paragraph(text))', 'a summary holding a block');
    assert.deepStrictEqual(summaryOf(blocks).children!.map(n => n.type), ['paragraph', 'paragraph'], 'Round 9: a summary holding a block, its blocks');
    const heading = await settled('<details><summary><h2>Installation</h2></summary><p>body</p></details>', 'html', 'disclosure[Installation](paragraph(text))', 'a heading summary');
    assert.deepStrictEqual(summaryOf(heading).children!.map(n => n.type), ['heading'], 'Round 9: a heading summary is a heading');
    const table = await settled('<details><summary>A<table><tr><td>t1</td></tr></table></summary><p>body</p></details>', 'html', 'disclosure[A t1](paragraph(text))', 'a table in a summary');
    assert.deepStrictEqual(summaryOf(table).children!.map(n => n.type), ['paragraph', 'table'], 'Round 9: a table in a summary');
    // 4: a link, a picture, a line break and a note in a Markdown summary.
    const rich = await settled('<details>\n<summary>\n\nSee [docs](https://ex.com/docs) ![logo](logo.png) one  \ntwo[^1]\n\n</summary>\n\nbody\n</details>\n\n[^1]: the note\n', 'md',
        'disclosure[See docs onetwo](paragraph(text))', 'a summary of a link, a picture, a break and a note');
    const line = summaryOf(rich).children![0];
    assert.ok(line.children!.some(n => (n.metadata as { link?: string } | undefined)?.link === 'https://ex.com/docs'), 'Round 9: the summary\'s link');
    assert.ok(line.children!.some(n => n.type === 'image') && line.children!.some(n => n.type === 'break'), 'Round 9: the summary\'s picture and line break');
    assert.ok(line.children!.some(n => n.notes?.some(note => nodeText(note) === 'the note')), 'Round 9: the summary\'s note');
    assert.ok(/one\ntwo/.test((await rich.to('text')).value), 'Round 9: the summary\'s line break in its text');
    // 3: a tag's name in a quoted value.
    await settled('The <abbr title="the <details> element">disclosure</abbr> widget folds text.\n\nNext para.\n', 'md', 'paragraph(text),paragraph(text)', 'a name in an inline tag\'s value');
    await settled('<details>\n<summary>Notes</summary>\n\n- Close it with <kbd title="</details>">end</kbd> tag\n- second item\n\nStill inside.\n\n</details>\n\nAfter.\n', 'md',
        'disclosure[Notes](list(text),list(text),paragraph(text)),paragraph(text)', 'a name in a value in a list item');
    // 5: the aligned wrapper, read as the writer writes it.
    await settled('<details><summary>S</summary><p style="text-align:center">Use <code>&lt;/details&gt;</code> w5</p><p>w6</p></details><p>after</p>', 'html',
        'disclosure[S](paragraph(text,text,text),paragraph(text)),paragraph(text)', 'a code span closing tag in an aligned paragraph');
    // 6: the text after a tag, a paragraph's text, its code span and link tail text.
    const cut = await settled('See <details><summary>more</summary><p>Use `</details>` to fold</p> w9\n\nw10\n\n</details>\n\nend\n', 'md',
        'paragraph(text),disclosure[more](paragraph(text,text,text),paragraph(text)),paragraph(text)', 'text after a tag');
    assert.ok((await cut.to('text')).value.includes('Use </details> to fold'), 'Round 9: text after a tag, its code span kept');
    await settled('<details>\n<summary>S</summary>\n\n[x](</details>) w2\n\n</details>\n\nafter\n', 'md', 'disclosure[S](paragraph(text,text)),paragraph(text)', 'a closing tag in a link\'s destination');
    // 7: an enclosing element's end tag.
    await settled('<div>\n<details>\n<summary>S</summary>\nx\n</div>\n\nafter\n\n## Next\n', 'md', 'disclosure[S](paragraph(text)),paragraph(text),heading(text)', 'a </div> closing a disclosure opened in its div');
    await settled('<div align="center">\n<details><summary>S</summary>\n\nbody\n\n</div>\n\nafter\n', 'md', 'disclosure[S](paragraph(text)),paragraph(text)', 'a </div> in a later block');
    // 8: summaries after the first, and a disclosure opened in a summary.
    await settled('<details><summary>A</summary><summary>w7</summary><summary>w8</summary></details>', 'html', 'disclosure[A](paragraph(text),paragraph(text))', 'summaries side by side (HTML)');
    await settled('<details>\n<summary>A</summary>\n<summary>w7</summary><summary>w8</summary>\n</details>\n', 'md', 'disclosure[A](paragraph(text),paragraph(text))', 'summaries side by side (Markdown)');
    await settled('<details>\n<summary>S <details>x</summary>\n\nbody\n', 'md', 'disclosure[S x](paragraph(text))', 'a disclosure opened in a summary, closed by its </summary>');
    // A lone <summary> line after a quote's or a list item's text is their lazy text, as 8.6.0 read it.
    assert.strictEqual(treeShape((await OfficeParser.parseOffice(Buffer.from('> q\n<summary>s</summary>\n'), { fileType: 'md' })).content), 'paragraph(text),paragraph(text)', 'Round 9: a <summary> line after a quote');
    const quoted = (await OfficeParser.parseOffice(Buffer.from('> q\n<summary>s</summary>\n'), { fileType: 'md' })).content;
    assert.deepStrictEqual(quoted.map(n => (n.metadata as ParagraphMetadata | undefined)?.style), ['Quote', 'Quote'], 'Round 9: a <summary> line after a quote stays in it');
    const afterQuoted = (await OfficeParser.parseOffice(Buffer.from('> <summary>s</summary>\nw\n'), { fileType: 'md' })).content;
    assert.deepStrictEqual(afterQuoted.map(n => (n.metadata as ParagraphMetadata | undefined)?.style), ['Quote', 'Quote'], 'Round 9: a line after a quote\'s <summary> line stays in it');
    const item = (await OfficeParser.parseOffice(Buffer.from('1.  a\n  <summary>s</summary>\n'), { fileType: 'md' })).content;
    assert.deepStrictEqual(item.map(n => [n.type, nodeText(n)]), [['list', 'a <summary>s</summary>']], 'Round 9: a <summary> line after a list item is its text');
    // (A disclosure's tags still end them.)
    assert.strictEqual(treeShape((await OfficeParser.parseOffice(Buffer.from('> q\n<details>\n<summary>s</summary>\n\nb\n\n</details>\n'), { fileType: 'md' })).content),
        'paragraph(text),disclosure[s](paragraph(text))', 'Round 9: a <details> line after a quote ends it');
    console.log('✓ Markdown disclosure round 9 findings');
}

/**
 * A disclosure in a list item or table cell, whose one line of Markdown holds no block, is written as HTML
 * (the item as an HTML list, nested items after it too; the table as an HTML table) and read back the same:
 * its summary, rich or plain, its body, its item's depth, number and task box. In a pipe-table cell with
 * the HTML table fallback off, or a definition, its summary in bold and its body, reported.
 */
export async function testMarkdownDisclosureInItemsAndCells(): Promise<void> {
    const same = async (html: string, name: string, expectedMd?: string): Promise<string> => {
        const ast = await OfficeParser.parseOffice(Buffer.from(html), { fileType: 'html' });
        const md = (await ast.to('md')).value;
        if (expectedMd !== undefined) assert.strictEqual(md, expectedMd, `In items and cells: ${name}, its Markdown`);
        const back = await OfficeParser.parseOffice(Buffer.from(md), { fileType: 'md' });
        // (Which list an item is in, and a bullet's place in it, which Markdown restarts after an HTML block, aside.)
        const keep = (nodes: OfficeContentNode[]) => JSON.stringify(nodes, function (this: any, key, value) {
            return key === 'listId' || (key === 'itemIndex' && this.listType === 'unordered') ? undefined : value;
        });
        assert.strictEqual(keep(back.content), keep(ast.content), `In items and cells: ${name}, read back the same: ${md}`);
        assert.strictEqual((await back.to('md')).value, md, `In items and cells: ${name}, the Markdown settled`);
        return md;
    };
    await same('<ul><li>x <details><summary>s <b>b</b></summary><p>body</p></details></li><li>y<ul><li>z <details open></details></li><li>after</li></ul></li></ul><ol start="3"><li>n <details><summary>t</summary><p>u</p></details></li></ol>',
        'list items', '<ul><li>x <details><summary>s <b>b</b></summary>\n<p>body</p></details></li></ul>\n\n- y\n\n<ul><ul><li>z <details open></details></li></ul></ul>\n\n<ul><ul><li>after</li></ul></ul>\n\n<ol start="3"><li>n <details><summary>t</summary>\n<p>u</p></details></li></ol>');
    await same('<ul data-type="taskList"><li data-checked="true">done <details><summary><h3>S</h3></summary><p>b</p></details></li></ul>', 'a task item');
    await same('<table><tr><td>h</td></tr><tr><td>c <details><summary><h3>S</h3><p>more</p></summary><p>b1</p><ul><li>li</li></ul></details></td></tr></table>', 'a table cell',
        '<table>\n  <tr>\n    <td>h</td>\n  </tr>\n  <tr>\n    <td>c <details><summary><h3>S</h3>\n<p>more</p></summary>\n<p>b1</p>\n<ul><li>li</li></ul></details></td>\n  </tr>\n</table>');
    await same('<table><tr><td>w17 <details></td></tr></table>', 'an empty disclosure in a cell', '<table>\n  <tr>\n    <td>w17 <details></details></td>\n  </tr>\n</table>');
    // With the HTML table fallback off: a pipe table, its summary in bold, reported.
    const cell = await OfficeParser.parseOffice(Buffer.from('<table><tr><td>c <details><summary>S</summary><p>b</p></details></td></tr></table>'), { fileType: 'html' });
    const pipe = await cell.to('md', { mdConfig: { fallbackToHtml: { tables: false } } });
    assert.ok(/^\| c<br>\*\*S\*\*(?: |<br>)b \|/m.test(pipe.value) && pipe.messages.some(m => m.code === 'CONTENT_NOT_REPRESENTABLE'), `In items and cells: a pipe-table cell, its summary in bold, reported: ${pipe.value}`);
    console.log('✓ Markdown disclosures in list items and cells');
}

/**
 * Round 10's findings, each read and written back. Markdown: no line after a closed quote's blank line
 * and an empty list item is lost, nor after a leading blank line, an indented first line or anchors
 * before an item (S2); an anchor before a disclosure stays on it (L2). A list item or table cell written
 * as HTML keeps a definition list, an admonition (its kind and title), a disclosure's id, a stray summary,
 * an embed and a paragraph holding a figure apart and whole (S3, M1, L3, L5).
 */
export async function testMarkdownDisclosureRound10Findings(): Promise<void> {
    const wordsIn = (s: string): string[] => s.match(/(?<![A-Za-z0-9])w\d+(?![0-9])/g) ?? [];
    const read = (source: string, fileType: 'md' | 'html' = 'md') => OfficeParser.parseOffice(Buffer.from(source), { fileType });
    const textOf = async (ast: OfficeParserAST): Promise<string> => (await ast.to('text')).value;
    const mdOf = async (ast: OfficeParserAST): Promise<string> => (await OfficeGenerator.generate(ast, 'md')).value as string;
    const astOf = (content: any[]): OfficeParserAST => ({ type: 'md', metadata: {}, attachments: [], content } as unknown as OfficeParserAST);
    const t = (text: string, extra: object = {}): any => ({ type: 'text', text, ...extra });
    const p = (...children: any[]): any => ({ type: 'paragraph', children });
    const disclosure = (summary: string, ...body: any[]): any => ({ type: 'disclosure', children: [{ type: 'summary', children: [p(t(summary))] }, ...body] });
    const item = (...children: any[]): any => ({ type: 'list', metadata: { listType: 'unordered', indentation: 0, listId: 'L' }, children });
    const cellTable = (...children: any[]): any => ({ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children }] }] });
    /** Writes `ast` as Markdown and reads it back `passes` times: each pass's Markdown, and the last tree. */
    const passes = async (ast: OfficeParserAST, count = 3): Promise<{ md: string[]; last: OfficeParserAST }> => {
        const md: string[] = [];
        let current = ast;
        for (let pass = 0; pass < count; pass++) {
            md.push(await mdOf(current));
            current = await read(md[pass]);
        }
        return { md, last: current };
    };
    const allWords = async (source: string, name: string): Promise<void> => {
        const text = await textOf(await read(source));
        assert.deepStrictEqual(wordsIn(source).filter(w => !wordsIn(text).includes(w)), [], `Round 10: ${name}, every word kept: ${JSON.stringify(source)} read ${JSON.stringify(text)}`);
    };

    // S2 (q13, q20): the text after a disclosure closed in a quote, its blank quote line and an empty item.
    const q13 = await textOf(await read('<details>\n<summary>Notes</summary>\n\n> quoted </details>\n>\n1.\nThis paragraph is lost.\nSo is this line.\n\nNext paragraph kept.\n'));
    assert.ok(q13.includes('This paragraph is lost. So is this line.') && q13.includes('Next paragraph kept.'), `Round 10: the lines after an empty item kept: ${JSON.stringify(q13)}`);
    for (const [source, words] of [['<details>\n\n> a </details>\n>\n-\nlost line\n', 'lost line'], ['<details>\n\n> a </details>\n>\n> - \nlazy line\n', 'lazy line'],
        ['<div>\n\n<details>\n\n> a </div>\n>\n1.\nlast words\n', 'last words']]) {
        assert.ok((await textOf(await read(source))).includes(words), `Round 10: ${JSON.stringify(words)} kept after an empty item`);
    }
    for (const source of ['<details>\n\n> a </details>\n>\n> *\n> w1 w2\n', '<div>\n\n> a </div>\n>\n> 1.\n> w1\n', '<details>\n\n> a\n> </details>\n>\n> 1.\n> w1\n',
        '<details>\n\n> > a </details>\n> >\n> > 1.\n> > w1\n', '<details>\n\n- a </details>\n\n  *\n  w1\n']) await allWords(source, 'a quoted empty item');
    // A leading blank line, an indented first line and anchors before an item: every word kept, and with a
    // leading blank line the reading of the text without it.
    for (const [source, without] of [['\n1.\nw1', '1.\nw1'], [' \n1.\nw1', '1.\nw1'], ['>\n1.\nw1', '1.\nw1'], ['\n2. a\nw1', '2. a\nw1']]) {
        await allWords(source, 'a leading blank line');
        const [ast, plain] = [await read(source), await read(without)];
        assert.strictEqual(treeShape(ast.content), treeShape(plain.content), `Round 10: ${JSON.stringify(source)} read as ${JSON.stringify(without)}`);
        assert.strictEqual(await textOf(ast), await textOf(plain), `Round 10: ${JSON.stringify(source)}, the text of ${JSON.stringify(without)}`);
    }
    for (const source of ['\t1.\nw2', '    1. a\nw2', '<a id="x"></a>\n1.\nw2']) await allWords(source, 'an item read after its indentation or anchors');

    // L2 (q17, q18): an anchor before a disclosure stays on it, and HTML gets the disclosure's id.
    for (const [source, name] of [['<blockquote><p>x</p><details id="w5"><summary>s</summary><p>y</p></details></blockquote><h2>h</h2>', 'after a quote'],
        ['<details id="w5"><summary>s</summary><p>y</p></details><h2>h</h2>', 'before a heading'], ['<p><a href="#w5">go</a></p><details id="w5"><summary>s</summary><p>y</p></details>', 'linked to']]) {
        const { md, last } = await passes(await read(source, 'html'));
        assert.ok(md.every(m => m.includes('<a id="w5"></a>\n\n<details>')), `Round 10: an anchor before a disclosure ${name} stays before it: ${JSON.stringify(md)}`);
        assert.strictEqual(new Set(md).size, 1, `Round 10: an anchor before a disclosure ${name}, the Markdown settled`);
        assert.ok((await last.to('html', { htmlConfig: { standalone: false } })).value.includes('<details id="w5">'), `Round 10: an anchor before a disclosure ${name}, the disclosure's id in HTML`);
    }

    // S3 (q1, q21): a definition list in an item or cell written as HTML, its words apart.
    const q1 = await mdOf(await read('<ul><li>alpha<dl><dt>beta</dt><dd>gamma</dd></dl><details><summary>s</summary><p>delta</p></details></li></ul>', 'html'));
    assert.ok(q1.includes('<dl><dt>beta</dt><dd>gamma</dd></dl>'), `Round 10: a definition list in an HTML item: ${q1}`);
    assert.ok(/alpha\s+beta\s+gamma\s+s\s+delta/.test(await textOf(await read(q1))), 'Round 10: a definition list in an HTML item, its words apart');
    const q1Cell = await mdOf(await read('<table><tr><td>alpha<dl><dt>beta</dt><dd>gamma</dd></dl></td><td colspan="2">x</td></tr></table>', 'html'));
    assert.ok(/alpha beta gamma/.test(await textOf(await read(q1Cell))), `Round 10: a definition list in an HTML cell, its words apart: ${q1Cell}`);
    const q21 = await mdOf(await read('<table><tr><th>h</th></tr><tr><td><details><summary>s</summary><dl><dt>alpha</dt><dd>beta</dd></dl></details></td></tr></table>', 'html'));
    assert.ok(/s alpha beta/.test(await textOf(await read(q21))), `Round 10: a definition list in a disclosure in a cell: ${q21}`);

    // M1 (q2): a disclosure's id in an item and in a cell, so a link to it resolves.
    const anchoredDisclosure = { ...disclosure('s', p(t('w3'))), metadata: { anchorIds: ['tgt'] } };
    const link = p(t('go', { metadata: { link: '#tgt', linkType: 'internal' } }));
    for (const [content, name] of [[[link, item(t('w1'), anchoredDisclosure)], 'item'], [[link, cellTable(p(t('w1')), anchoredDisclosure)], 'cell']] as [any[], string][]) {
        const md = await mdOf(astOf(content));
        assert.ok(md.includes('<details id="tgt">'), `Round 10: a disclosure's id in an HTML ${name}: ${md}`);
        const found: OfficeContentNode[] = [];
        const walk = (nodes: OfficeContentNode[] | undefined) => { for (const n of nodes ?? []) { if (n.type === 'disclosure') found.push(n); walk(n.children); } };
        walk((await read(md)).content);
        assert.deepStrictEqual(found.map(n => (n.metadata as { anchorIds?: string[] } | undefined)?.anchorIds), [['tgt']], `Round 10: a disclosure's id in an HTML ${name}, read back`);
    }

    // L3 (l1): an admonition in an item holding a disclosure keeps its kind, and settles.
    const admonitionItem = await passes(astOf([item(t('w1'), { type: 'admonition', metadata: { admonitionType: 'note', sourceSyntax: 'github' }, children: [disclosure('s', p(t('w3')))] })]), 4);
    assert.strictEqual(treeShape(admonitionItem.last.content), 'list(text,admonition(disclosure[s](paragraph(text))))', `Round 10: an admonition in an HTML item: ${admonitionItem.md[0]}`);
    assert.strictEqual(admonitionItem.md[2], admonitionItem.md[3], 'Round 10: an admonition in an HTML item, the Markdown settled');
    // Its custom title in a cell, a bold first line, written as it reads back.
    const titled = await passes(astOf([cellTable(p(t('w1')), disclosure('s', p(t('w2'))), { type: 'admonition', metadata: { admonitionType: 'warning', title: 'w8' }, children: [p(t('w9'))] })]), 2);
    assert.ok(titled.md[0].includes('data-type="warning"><p><b>w8</b></p>\n<p>w9</p></div>'), `Round 10: an admonition's title in an HTML cell: ${titled.md[0]}`);
    assert.strictEqual(titled.md[1], titled.md[0], 'Round 10: an admonition\'s title in an HTML cell, the Markdown settled');
    assert.ok(/w8 w9/.test(await textOf(titled.last)), 'Round 10: an admonition\'s title in an HTML cell, its words apart');
    // An admonition ending in a table: no line of a lone </div>, and the rest of the item stays in it.
    const tableEnd = await passes(astOf([item(t('w1'), disclosure('s', p(t('w2'))), { type: 'admonition', metadata: { admonitionType: 'note' }, children: [p(t('w3')), cellTable(p(t('w4')))] }, p(t('w5')))]), 2);
    assert.ok(!/^<\/div>$/m.test(tableEnd.md[0]), `Round 10: an admonition ending in a table in an HTML item, no lone </div> line: ${tableEnd.md[0]}`);
    assert.strictEqual(treeShape(tableEnd.last.content), 'list(text,disclosure[s](paragraph(text)),admonition(paragraph(text),table(row(cell(paragraph(text))))),paragraph(text))',
        'Round 10: an admonition ending in a table in an HTML item, the item whole');
    assert.strictEqual(tableEnd.md[1], tableEnd.md[0], 'Round 10: an admonition ending in a table in an HTML item, the Markdown settled');

    // L5 (s2): a summary that is no disclosure's, set apart in a one-line item and in an HTML item.
    const stray = await mdOf(astOf([item(t('w0'), { type: 'summary', children: [p(t('w1'))] })]));
    assert.strictEqual(stray, '- w0<br>**w1**', 'Round 10: a stray summary in a one-line item');
    assert.ok(/w0\s+w1/.test(await textOf(await read(stray))), 'Round 10: a stray summary in a one-line item, its words apart');
    const strayHtml = await mdOf(astOf([item(t('w0'), { type: 'summary', children: [p(t('w1'))] }, disclosure('s', p(t('w2'))))]));
    assert.ok(strayHtml.includes('<div><p>w1</p></div>') && /w0\s+w1\s+s\s+w2/.test(await textOf(await read(strayHtml))), `Round 10: a stray summary in an HTML item, a block: ${strayHtml}`);

    // An embed after an item's text, in a paragraph of its own.
    const embed = await passes(astOf([item(t('w5'), { type: 'embed', metadata: { url: 'https://e.x/v', label: 'w6' } }, disclosure('s', p(t('w7'))))]), 2);
    assert.ok(/w5\s+w6/.test(await textOf(embed.last)), `Round 10: an embed after an HTML item's text, apart: ${embed.md[0]}`);
    // A paragraph holding a figure in a disclosure in a cell: split at it, no empty paragraph added on a save.
    const figure = await passes(astOf([cellTable(p(t('w1')), disclosure('s', p(t('w2'), { type: 'image', metadata: { url: 'https://e.x/f.png', altText: 'alt', caption: 'w3' } }, t('w4'))))]), 3);
    assert.ok(figure.md[0].includes('<p>w2</p>\n<figure>') && figure.md[0].includes('</figure>\n<p>w4</p>'), `Round 10: a paragraph holding a figure in an HTML cell, split: ${figure.md[0]}`);
    assert.strictEqual(new Set(figure.md).size, 1, `Round 10: a paragraph holding a figure in an HTML cell, the Markdown settled: ${JSON.stringify(figure.md)}`);
    console.log('✓ Markdown disclosure round 10 findings');
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

/**
 * A disclosure's summary is the `<summary>` first in it (white space, blank lines too, between, a
 * `</summary>` with none open ignored), to its `</summary>` wherever that falls: across blank lines and
 * blocks, with disclosure tags in it (its text). Its text runs inline nodes together and sets blocks
 * apart; the Markdown written settles.
 */
export async function testMarkdownDisclosureSummaries(): Promise<void> {
    const cases: [string, string, string][] = [
        ['<details>\n<summary>\n\n**Bold** summary\n\n</summary>\n\nBody\n\n</details>\n\nAfter\n', 'disclosure[Bold summary](paragraph(text)),paragraph(text)', 'a Markdown summary across blank lines'],
        ['<details><summary>\n\n## Heading summary\n\n</summary>\n\nBody\n\n</details>\n\nAfter\n', 'disclosure[Heading summary](paragraph(text)),paragraph(text)', 'a heading as the summary'],
        ['<details>\n<summary>\n\n### Install on Linux\n\n</summary>\n\nRun make.\n\n</details>\n\nAfter\n', 'disclosure[Install on Linux](paragraph(text)),paragraph(text)', 'a heading summary, then the body'],
        ['<details>\n<summary>S</summary>\n<p>raw</p>\n\nBody *md*\n\n</details>\n\nAfter\n', 'disclosure[S](paragraph(text),paragraph(text,text)),paragraph(text)', 'raw HTML after a one-line summary'],
        ['<details>\n\n\n<summary>S</summary>\n\nBody\n\n</details>\n\nAfter\n', 'disclosure[S](paragraph(text)),paragraph(text)', 'a summary after blank lines'],
        ['<details>\n<summary>S <details> x</summary>\n\nBody\n\n</details>\n\nAfter\n', 'disclosure[S x](paragraph(text)),paragraph(text)', 'a disclosure tag in the summary'],
        ['<DETAILS OPEN>\n<SUMMARY>S</SUMMARY>\n\nBody\n\n</DETAILS >\n\nAfter\n', 'disclosure[S](paragraph(text)),paragraph(text)', 'upper case'],
        ['Text <details>\n\nBody\n\n</details> tail\n\nAfter\n', 'paragraph(text),disclosure[](paragraph(text)),paragraph(text),paragraph(text)', 'no summary'],
        ['<details>\n<summary>a<b>b</b>c <kbd>K</kbd>d</summary>\n\nx\n\n</details>\n', 'disclosure[abc Kd](paragraph(text))', 'inline formatting in a one-line summary'],
        ['> <details>\n> </summary>\n> <summary>S</summary> body\n', 'disclosure[S](paragraph(text))', 'a stray </summary> before the summary'],
        ['<details>\n<summary>S <details><summary>inner</summary> x</details> y</summary>\n\nB\n\n</details>\n', 'disclosure[S inner x y](paragraph(text))', 'a disclosure in the summary, its summary kept'],
    ];
    for (const [source, shape, name] of cases) {
        const read = await markdownRead(source);
        assert.strictEqual(read.shape, shape, `Disclosure summary: ${name}`);
        const again = await markdownRead(read.md);
        assert.strictEqual(again.md, read.md, `Disclosure summary: ${name}, the Markdown written settles`);
        assert.strictEqual(again.shape, shape, `Disclosure summary: ${name}, the same tree read back`);
    }
    const open = await OfficeParser.parseOffice(Buffer.from('<details open>\n<summary>\n\nS\n\n</summary>\n\nB\n\n</details>\n'), { fileType: 'md' });
    assert.strictEqual(JSON.stringify(open.content[0].metadata), '{"open":true}', 'Disclosure summary: its metadata as the HTML reader writes it');
    assert.strictEqual(JSON.stringify(open.content[0].children![0]), '{"type":"summary","children":[{"type":"paragraph","metadata":{},"children":[{"type":"text","text":"S"}]}]}', 'Disclosure summary: a summary node, its first child');
    console.log('✓ Markdown disclosure summaries');
}

/** Block kinds a disclosure's closing tag follows in its body, tight or after a blank line. */
const DISCLOSURE_NEIGHBOURS: Record<string, string> = {
    para: 'Para *x*', head: '## Head', setext: 'Title\n===', list: '- a\n- b', olist: '1. a', listIndented: '- a\n\n  more', quote: '> q', quoteLazy: '> q\nlazy',
    fence: '```\nc\n```', indent: '    code', pipe: '| a |\n| - |\n| b |', hr: '***', div: '<div align="center">\n\n*c*\n\n</div>', pdiv: '<p align="center">x</p>',
    table: '<table><tr><td>\n\n*m*\n\n</td></tr></table>', img: '![a](x.png)', fn: 'N[^z]\n\n[^z]: fnote', def: 'Term\n: def', comment: '<!-- c -->', math: '$$\nx^2\n$$',
    adm: '> [!NOTE]\n> text', task: '- [ ] todo', kbd: '<kbd>K</kbd> key',
};

/**
 * A disclosure's or summary's tag starting a line is an HTML block (CommonMark's condition 6): it ends a
 * quote, admonition or list item it is not marked or indented into. A closing tag in a quote, list item
 * or heading closes that too, as a browser has it: what follows it there stands after the disclosure,
 * a list's later items a block of their own; a `</summary>` in one closes a summary open outside.
 * Tags in a heading's text, a code span or an HTML table in it close nothing outside it.
 */
export async function testMarkdownDisclosureContainerCloses(): Promise<void> {
    const note = await markdownRead('<details>\n<summary>Note</summary>\n\n> [!NOTE]\n> Read this.\n</details>\n\n## Install\n\nRun it.\n');
    assert.strictEqual(note.shape, 'disclosure[Note](admonition(paragraph(text))),heading(text),paragraph(text)', 'Disclosure close: a closing line after an admonition ends it');
    assert.strictEqual(note.md, '<details>\n<summary>Note</summary>\n\n> [!NOTE]\n> Read this.\n\n</details>\n\n## Install {#install}\n\nRun it.', 'Disclosure close: after an admonition, its Markdown');
    const item = await markdownRead('<details>\n<summary>Log</summary>\n\n- step one\n  </details>\n\n## Next\n');
    assert.strictEqual(item.shape, 'disclosure[Log](list(text)),heading(text)', 'Disclosure close: a closing line in a list item');
    assert.strictEqual(item.md, '<details>\n<summary>Log</summary>\n\n- step one\n\n</details>\n\n## Next {#next}', 'Disclosure close: in a list item, no tag kept as its text');
    // Every block kind before the closing tag in the body, and before the disclosure, tight and after a blank line.
    for (const [name, block] of Object.entries(DISCLOSURE_NEIGHBOURS)) for (const join of ['\n', '\n\n']) {
        const inBody = (await OfficeParser.parseOffice(Buffer.from(`<details>\n<summary>S</summary>\n\n${block}${join}</details>\n\nAFTER\n`), { fileType: 'md' })).content;
        const body = JSON.stringify(inBody[0]);
        assert.ok(inBody.length === 2 && inBody[0].type === 'disclosure' && JSON.stringify(inBody[1]).includes('AFTER') && !body.includes('AFTER') && !JSON.stringify(inBody).includes('</details>'),
            `Disclosure close: after ${name} (${JSON.stringify(join)}) the disclosure ends: ${treeShape(inBody)}`);
        const before = (await OfficeParser.parseOffice(Buffer.from(`${block}${join}<details>\n<summary>S</summary>\n\nBODY\n\n</details>\n\nAFTER\n`), { fileType: 'md' })).content;
        const disclosure = before.find(node => node.type === 'disclosure');
        assert.ok(disclosure && JSON.stringify(disclosure).includes('BODY') && !JSON.stringify(disclosure).includes('AFTER') && JSON.stringify(before[before.length - 1]).includes('AFTER'),
            `Disclosure close: a disclosure after ${name} (${JSON.stringify(join)}) holds its body: ${treeShape(before)}`);
    }
    for (const [name, block] of Object.entries({ liCont: '- a\n  ', quoteGt: '> q\n> ', quoteNested: '> > q\n', admLazy: '> [!TIP]\n> t\n', quote2: 'Para\n\n> q\n', quoteThenPara: '> q\n\nP\n' })) {
        const content = (await OfficeParser.parseOffice(Buffer.from(`<details>\n<summary>S</summary>\n\n${block}</details>\n\nAFTER\n`), { fileType: 'md' })).content;
        assert.ok(content.length === 2 && content[0].type === 'disclosure' && JSON.stringify(content[1]).includes('AFTER') && !JSON.stringify(content[0]).includes('AFTER') && !JSON.stringify(content).includes('</details>'),
            `Disclosure close: a closing tag marked or indented into ${name} closes it: ${treeShape(content)}`);
    }
    const cases: [string, string, string][] = [
        ['<details>\n\n> q\n> </details>\n> tail\n\nout\n', 'disclosure[](paragraph(text)),paragraph(text),paragraph(text)', 'a marked closing line: the quote\'s later lines after it'],
        ['<details open>\n\n1. a </details>\n2. <details >\n', 'disclosure[](list(text)),list(text)', 'a list\'s later items a list of their own'],
        ['<details>\n\n## h <summary>s</summary> `</details> x`\n\nafter\n\n</details>\n\nout\n', 'disclosure[](heading(text,text),paragraph(text)),paragraph(text)', 'a heading\'s text is inline: a tag in its code span closes nothing'],
        ['<details>\n\n> q\n> <table><tr><td>a </details></td></tr></table>\n\nin\n\n</details>\n\nout\n', 'disclosure[](paragraph(text),table(row(cell(text))),paragraph(text)),paragraph(text)', 'a table in a quote holds its closing tag'],
        ['<details open>\n<summary>\n\n> quoted\n> </summary>\n\nBody\n\n</details>\n', 'disclosure[quoted](paragraph(text))', 'a summary closed in a quote'],
        ['<details>\n\n> <details>\n> inner\n\nout of the quote\n\n</details>\n\nafter\n', 'disclosure[](disclosure[](paragraph(text)),paragraph(text)),paragraph(text)', 'a disclosure opened in a quote ends with it'],
    ];
    for (const [source, shape, name] of cases) {
        const read = await markdownRead(source);
        assert.strictEqual(read.shape, shape, `Disclosure close: ${name}`);
        assert.ok(name.includes('code span') || !read.text.includes('</details>'), `Disclosure close: ${name}, no closing tag in its text: ${read.text}`);
    }
    console.log('✓ Markdown disclosure closed in quotes, list items and headings');
}

/**
 * Past 16 levels of nesting a Markdown disclosure is flattened, inline or as a block: its tags go, its
 * summary and body stay as text in order, with no tag written as text, and the text and Markdown are
 * the same over three round trips.
 */
export async function testMarkdownDisclosureFlattenedPastDepth(): Promise<void> {
    for (const levels of [16, 17, 20]) for (const inline of [true, false]) {
        let source = '';
        for (let i = 0; i < levels; i++) source += inline ? `T${i} <details><summary>S${i}</summary> B${i} *x*\n\n` : `<details>\n<summary>S${i}</summary>\n\nB${i} *x*\n\n`;
        source += 'CORE\n\n';
        for (let i = levels - 1; i >= 0; i--) source += `</details>\n\nA${i}\n\n`;
        const name = `${levels} levels${inline ? ' inline' : ''}`;
        let ast = await OfficeParser.parseOffice(Buffer.from(source), { fileType: 'md' });
        const text = (await ast.to('text')).value;
        const words = source.match(/\b[TSBA]\d+\b|CORE/g)!;
        assert.deepStrictEqual(words.filter(word => !new RegExp(`\\b${word}\\b`).test(text)), [], `Disclosure depth: ${name}, nothing dropped`);
        assert.ok(!/<\/?(?:details|summary)|&lt;/i.test(text), `Disclosure depth: ${name}, no tag as text: ${text.slice(text.indexOf('T15'), text.indexOf('T15') + 120)}`);
        const markdown: string[] = [];
        for (let pass = 0; pass < 3; pass++) {
            markdown.push((await ast.to('md')).value);
            ast = await OfficeParser.parseOffice(Buffer.from(markdown[pass]), { fileType: 'md' });
        }
        assert.strictEqual((await ast.to('text')).value, text, `Disclosure depth: ${name}, the same text over three round trips`);
        assert.strictEqual(markdown[2], markdown[1], `Disclosure depth: ${name}, the Markdown settles`);
        let depth = 0;
        const walk = (nodes: OfficeContentNode[] | undefined, level: number): void => {
            for (const node of nodes ?? []) {
                if (node.type === 'disclosure') depth = Math.max(depth, level + 1);
                walk(node.children, level + (node.type === 'disclosure' ? 1 : 0));
            }
        };
        walk(ast.content, 0);
        assert.strictEqual(depth, 16, `Disclosure depth: ${name}, 16 levels nested`);
    }
    console.log('✓ Markdown disclosures flattened past 16 levels');
}

/** Hostile Markdown shapes for disclosures, each `n` times over: the review's and the fix design's. */
const DISCLOSURE_SHAPES: Record<string, (n: number) => string> = {
    unclosedBlank: n => '<details>\n<summary>S</summary>\n\nbody\n\n'.repeat(n),
    strayClose: n => 'p\n\n</details>\n\n'.repeat(n),
    inlineOneBlock: n => 'x <details><summary>s</summary> y '.repeat(n),
    htmlOneBlock: n => '<details><summary>s\n'.repeat(n),
    htmlOneBlockClosed: n => '<details><summary>s</summary>\n'.repeat(n) + '</details>\n'.repeat(n),
    siblings: n => '<details>\n<summary>S</summary>\n\nBody *x*\n\n</details>\n\n'.repeat(n),
    longBody: n => '<details>\n<summary>S</summary>\n\n' + 'para *x* `c`\n\n'.repeat(n) + '</details>\n',
    interleaved: n => '<details>\n\n<table><tr><td>\n\n</details>\n\n</td></tr></table>\n\n'.repeat(n),
    codeSpans: n => ('`<details>` ' + '`'.repeat(3) + ' a ').repeat(n) + '\n\n</details>\n',
    backtickRuns: n => '<details>\n\n' + Array.from({ length: n }, (_, i) => '`'.repeat(1 + (i % 50)) + ' <details> ').join('') + '\n\n</details>',
    noGt: n => '<details\n'.repeat(n) + '\n\n' + 'x'.repeat(n),
    summaryNoClose: n => ('<details><summary>' + 'a'.repeat(50) + '\n').repeat(n),
    summaryBlankLead: n => '<details>\n\n<summary>s\n\nb\n\n'.repeat(n),
    nested16x: n => ('<details>\n\n'.repeat(16) + 'b\n\n' + '</details>\n\n'.repeat(16)).repeat(Math.max(1, n / 32)),
    deepNest: n => '<details>\n\nb\n\n'.repeat(n) + '</details>\n\n'.repeat(n),
    deepInline: n => 'p <details> q\n\n'.repeat(n) + 'p </details> q\n\n'.repeat(n),
    quoteSwallow: n => '<details>\n\n> q\n</details>\n\n'.repeat(n),
    escaped: n => '\\'.repeat(n) + '<details>\n\nx\n\n</details>',
    quoteClose: n => '<details>\n\n> a </details>\n\n'.repeat(n),
    quoteLineClose: n => '<details>\n\n> a\n> </details>\n> b\n\n'.repeat(n),
    listInQuoteClose: n => '<details>\n\n> - a\n>   </details>\n> - b\n\n'.repeat(n),
    listClose: n => '<details>\n\n- a\n  </details>\n- b\n\n'.repeat(n),
    headingClose: n => '<details>\n\n## h </details> t\n\n'.repeat(n),
    manyOpenOneQuote: n => '<details>\n\n'.repeat(n) + '> ' + '</details> '.repeat(n) + '\n',
    manyOpenQuoteLines: n => '<details>\n\n'.repeat(n) + '> x\n> </details>\n'.repeat(n),
    nestedContainers: n => '<details>\n\n'.repeat(16) + '> - '.repeat(Math.min(n, 2000)) + 'x </details>\n',
    quoteOpensConfined: n => '<details>\n\n' + '> <details> a\n'.repeat(n) + '> </details>\n',
    listItemsConfined: n => '<details>\n\n' + '- <details> a\n'.repeat(n) + '- </details>\n',
    unclosedAttrHtml: n => '<details title="'.repeat(n) + '\n\nx\n',
    unclosedAttrInline: n => 'x ' + '<details a="'.repeat(n) + '\n\nx\n',
    unclosedAttrQuote: n => '<details>\n\n> ' + '<details a="'.repeat(n) + '\n',
    attrWithGt: n => '<details title="a>b">\n<summary title="c>d">s</summary>\n\nbody\n\n</details>\n\n'.repeat(n),
    summaryOpenAcross: n => '<details><summary>\n\nx\n\n'.repeat(n),
    summaryFarClose: n => '<details><summary>\n\n' + '<details>\n\n'.repeat(n) + '</summary>\n\nbody\n\n</details>\n',
    summaryAcrossSiblings: n => '<details>\n<summary>\n\n### S\n\n</summary>\n\nBody\n\n</details>\n\n'.repeat(n),
    lazyAfterHtmlQuote: n => '<details>\n\n' + '> <details>\nlazy\n'.repeat(n),
    summaryInQuoteClose: n => '<details>\n<summary>\n\n> s\n> </summary>\n\nbody\n\n</details>\n\n'.repeat(n),
    // Round 9's review: a close after each open in one block, in quotes, list items and headings.
    inlinePairs: n => 'a <details>x</details> '.repeat(n) + '\n',
    inlinePairsSum: n => 'a <details><summary>s</summary>x</details> '.repeat(n) + '\n',
    headingPairs: n => '## h ' + '<details>x</details> '.repeat(n) + '\n',
    quoteCloses: n => '<details>\n\n' + '> a </details> <details>\n'.repeat(n),
    quoteLines: n => '<details>\n\n> q\n' + '> </details>\n> <details>\n> x\n'.repeat(n),
    listItems: n => '<details>\n\n' + '- a </details> <details>\n'.repeat(n),
    listNested: n => '<details>\n\n' + '- a\n  - b\n    </details><details>\n'.repeat(n),
    sumAcross: n => '<details>\n<summary>\n\n' + 'p\n\n'.repeat(n) + '</summary>\n\nbody\n</details>\n',
    sumUnclosed: n => '<details><summary>s\n\n'.repeat(n) + '</details>\n'.repeat(n),
    strayCloses: n => '<details>' + ' </summary>'.repeat(n) + '\n\nbody\n',
    attrQuotes: n => '<details>\n\n' + 'x <details title="a '.repeat(n) + '\n',
    attrQuotes2: n => 'x <span title="' + '<details a='.repeat(n) + '\n',
    quoteFence: n => '<details>\n\n' + '> ```\n> </details>\n> ```\n'.repeat(n),
    quoteHtml: n => '<details>\n\n' + '> <div>\n> </details>\n> <details>\n>\n'.repeat(n),
    // Round 9's fix design: every path of the tag scanner, the tree builder and the cut pass.
    attrTagNames: n => 'The <abbr title="the <details> element">x</abbr> w '.repeat(n) + '\n',
    attrTagNamesHtml: n => '<p title="</details>">x</p>\n'.repeat(n),
    commentsUnclosedInline: n => 'x ' + '<!-- <details> '.repeat(n) + '\n',
    commentsUnclosedHtml: n => '<div>\n' + '<!-- <details> '.repeat(n) + '\n',
    declarationsInline: n => 'x ' + '<!a <details> '.repeat(n) + '\n',
    rawTextUnclosed: n => '<div>\n' + '<script> <details> '.repeat(n) + '\n',
    linkTails: n => '<details>\n\n' + '[a](</details> [b](<x '.repeat(n) + '\n',
    linkTitles: n => '<details>\n\n' + '[a](u "</details>) '.repeat(n) + '\n',
    backtickRunsInBody: n => '<details>\n\n' + Array.from({ length: n }, (_, i) => '`'.repeat(1 + (i % 40)) + ' <details> ').join('') + '\n',
    divCloses: n => '<div>\n\n' + '<details>\n\nx\n\n</div>\n\n<div>\n\n'.repeat(n),
    divEnclosingInline: n => '<div>\n\n' + 'a <details> b </div> <div> '.repeat(n) + '\n',
    wrapperCode: n => '<details>\n\n' + '<div style="text-align: center">Use `</details>` w</div>\n\n'.repeat(n) + '</details>\n',
    wrapperCloses: n => '<details>\n\n<div style="text-align: center">' + 'a <details> b </div> '.repeat(n) + '</div>\n',
    summaryBlocksMany: n => '<details>\n<summary>\n\n' + '### h\n\n| a | b |\n| - | - |\n| c | d |\n\n'.repeat(n) + '</summary>\n\nbody\n\n</details>\n',
    summariesSideBySide: n => '<details>\n<summary>A</summary>' + '<summary>w</summary>'.repeat(n) + '\n</details>\n',
    listDeepEscapes: n => '<details>\n\n'.repeat(16) + Array.from({ length: Math.min(n, 400) }, (_, i) => '  '.repeat(i) + '- a </details> <details>').join('\n') + '\n',
    listEscapesLong: n => '<details>\n\n' + '- a\n  more text here\n  </details> b <details>\n'.repeat(n),
    quoteListEscapes: n => '<details>\n\n' + '> - a </details> <details>\n> - b\n'.repeat(n),
    itemHtmlRunOn: n => '<details>\n\n- <div>\n' + '  - x </details> <details>\n'.repeat(n),
    inlineAfterTagCode: n => 'See <details><summary>more</summary>' + 'Use `</details>` and \\</details> x '.repeat(n) + '</details>\n',
    nested16Inline: n => ('a <details> '.repeat(20) + 'b ' + '</details> '.repeat(20)).repeat(Math.max(1, n / 20)) + '\n',
    openMany: n => '<details>\n\n'.repeat(n) + 'x\n',
    closeManyOneQuote: n => '<details>\n\n'.repeat(n) + '> ' + '</details> '.repeat(n) + '\n',
    summaryUnclosedNested: n => '<details><summary>s <details>x\n\n'.repeat(n) + 'body\n',
    // Disclosures in list items and table cells, which the Markdown writer writes as HTML lists and tables.
    itemsHoldingDisclosures: n => '<ul>' + '<li>x <details><summary>s <b>b</b></summary>y</details><ul><li>z <details>w</details></li></ul></li>'.repeat(n) + '</ul>\n',
    cellsHoldingDisclosures: n => '<table>' + '<tr><td>x <details><summary>s</summary>y</details></td><td>z</td></tr>'.repeat(n) + '</table>\n',
};

/**
 * Markdown disclosures are read and written back in time linear in the text, on every hostile shape the
 * reviews and the fix designs found: closing tags in quotes, list items and headings, many opened, values
 * never closed, summaries across blocks, tags in attributes, comments and link tails, enclosing elements,
 * disclosures in list items and cells. Each against plain paragraphs of the same size, and at n against
 * 4n (at most 8 times the time: a quadratic reading takes 16); a shape the HTML reader's nesting limit
 * refuses (one HTML block of a thousand open disclosures) is held to the same time.
 */
export async function testMarkdownDisclosureShapesLinear(): Promise<void> {
    const n = 4000;
    const plain = (length: number) => 'Some *plain* text here.\n\n'.repeat(Math.ceil(length / 25));
    const readAndWrite = async (source: string): Promise<void> => {
        try {
            await (await OfficeParser.parseOffice(Buffer.from(source), { fileType: 'md' })).to('md');
        } catch (error) {
            if (!/nesting depth/i.test((error as Error).message)) throw error;
        }
    };
    for (const [name, shape] of Object.entries(DISCLOSURE_SHAPES)) {
        const source = shape(n);
        const { within, time, plain: baseline } = await linearAgainst(() => readAndWrite(source), () => readAndWrite(plain(source.length)));
        assert.ok(within, `Disclosure shapes: ${name} (${source.length} bytes) in ${time.toFixed(0)} ms, paragraphs in ${baseline.toFixed(0)} ms`);
        // n against 4n: the time at 4000 against the time at 1000, taken again once where it is over.
        const quarter = shape(n / 4);
        let small = await timed(() => readAndWrite(quarter));
        let large = time;
        if (large > 8 * small + 250) {
            small = Math.max(small, await timed(() => readAndWrite(quarter)));
            large = Math.min(large, await timed(() => readAndWrite(source)));
        }
        assert.ok(large <= 8 * small + 250, `Disclosure shapes: ${name} at ${n} in ${large.toFixed(0)} ms, at ${n / 4} in ${small.toFixed(0)} ms`);
    }
    console.log(`✓ Markdown disclosure shapes in linear time (${Object.keys(DISCLOSURE_SHAPES).length} shapes, n against 4n)`);
}

/** Hostile shapes of summaries, disclosures and Markdown lines, each `n` times over: round 10's review and fix design. */
const SUMMARY_SHAPES: Record<string, ['html' | 'md', (n: number) => string, number?]> = {
    // (At 32,000 for summaries after content: at 4,000, a quadratic reading of them still took under 8 times its time at 1,000.)
    textThenSummaries: ['html', n => '<details>' + 'a<br>'.repeat(n) + '<summary>s</summary>'.repeat(n) + '</details>', 32000],
    inlineThenSummaries: ['html', n => '<details>' + '<i>x</i>'.repeat(n) + '<summary>s</summary>'.repeat(n) + '</details>', 32000],
    blocksThenSummaries: ['html', n => '<details>' + '<p>a</p>'.repeat(n) + '<summary>s</summary>'.repeat(n) + '</details>', 32000],
    deepSummaries: ['html', n => '<details>a' + '<div>'.repeat(200) + 'b<br>'.repeat(n) + '<summary>s</summary>'.repeat(n) + '</details>', 32000],
    nestedSummaryRuns: ['html', n => '<details><summary>t</summary>' + ('<summary>'.repeat(50) + 'x' + '</summary>'.repeat(50)).repeat(n / 50) + '</details>'],
    manyDetails: ['html', n => '<details><summary>a</summary><summary>b</summary>x</details>'.repeat(n)],
    nestedDetails: ['html', n => ('<details>' + 'a<br>'.repeat(n / 100) + '<summary>s</summary>'.repeat(n / 100)).repeat(100) + '</details>'.repeat(100)],
    listItems: ['html', n => '<ul>' + '<li>i <details><summary>s</summary>x</details></li>'.repeat(n) + '</ul>'],
    tableCells: ['html', n => '<table><tr>' + '<td>c<details><summary>s</summary>x</details></td>'.repeat(n) + '</tr></table>'],
    itemAdmonitions: ['html', n => '<ul><li>i<details><summary>s</summary>x</details>'
        + '<div class="admonition" data-type="note"><dl><dt>t</dt><dd>d<figure><img src="x.png"><figcaption>c</figcaption></figure></dd></dl></div>'.repeat(n) + '</li></ul>'],
    deepCellNesting: ['html', n => ('<table><tr><td colspan="2"><div class="admonition" data-type="note">' + 'x '.repeat(n / 50)).repeat(50) + '</div></td></tr></table>'.repeat(50)],
    mdSummaries: ['md', n => '<details>\n' + 'a\n\n'.repeat(n) + '<summary>s</summary>\n\n'.repeat(n) + '</details>\n'],
    mdQuoted: ['md', n => '> <details>\n> <summary>s</summary>\n> x </details>\n>\n'.repeat(n)],
    mdSpansThenCloses: ['md', n => '<details>\n\n' + '<span>'.repeat(n) + '</details>'.repeat(n) + '\n'],
    mdSummaryRuns: ['md', n => '<details>\n' + '<summary>s</summary>\n'.repeat(n) + 'x\n</details>\n'],
    mdQuoteBlankItems: ['md', n => '<details>\n\n> q </details>\n>\n1.\nw\n'.repeat(n)],
    mdLeadingBlankLists: ['md', n => '\t1.\nw\n# h\n'.repeat(n)],
    mdLazyLines: ['md', n => '\t1.\n' + 'w\n'.repeat(n)],
    mdItems: ['md', n => '- <details><summary>s</summary>x</details>\n'.repeat(n)],
};

/**
 * Summaries and disclosures are read and written in time linear in the text, on round 10's shapes: extra
 * summaries after text, inline runs, blocks and 200 nested divs (each found its disclosure's first summary
 * by a scan of its children), nested summary runs, many and nested disclosures, disclosures in list items
 * and cells, admonitions, definitions and figures in an HTML item, deep cell nesting, and Markdown
 * summaries, quotes, closing runs, quote-blank-item runs, leading-blank lists and lazy lines. Each read and
 * written as Markdown, HTML and text, at n against 4n (at most 8 times the time: a quadratic reading takes 16;
 * n is 4,000, and 32,000 for the summaries after content, whose reading was quadratic).
 */
export async function testDisclosureSummaryShapesLinear(): Promise<void> {
    const readAndWrite = async (fileType: 'html' | 'md', source: string): Promise<void> => {
        try {
            const ast = await OfficeParser.parseOffice(Buffer.from(source), { fileType });
            await ast.to('md');
            await ast.to('html', { htmlConfig: { standalone: false } });
            await ast.to('text');
        } catch (error) {
            if (!/nesting depth/i.test((error as Error).message)) throw error;
        }
    };
    for (const [name, [fileType, shape, n = 4000]] of Object.entries(SUMMARY_SHAPES)) {
        const [source, quarter] = [shape(n), shape(n / 4)];
        let small = await timed(() => readAndWrite(fileType, quarter));
        let large = await timed(() => readAndWrite(fileType, source));
        if (large > 8 * small + 250) {
            small = Math.max(small, await timed(() => readAndWrite(fileType, quarter)));
            large = Math.min(large, await timed(() => readAndWrite(fileType, source)));
        }
        assert.ok(large <= 8 * small + 250, `Summary shapes: ${name} at ${n} (${source.length} bytes) in ${large.toFixed(0)} ms, at ${n / 4} in ${small.toFixed(0)} ms`);
    }
    console.log(`✓ Summary and disclosure shapes in linear time (${Object.keys(SUMMARY_SHAPES).length} shapes, n against 4n)`);
}

/**
 * A captioned picture in a line (a list item, pipe-table cell, definition or heading) ends the line it is
 * on: the words on both sides stay apart, no `<figure>` markup is written into a heading, and Markdown
 * settles at the second save.
 */
export async function testFigureEndsItsLine(): Promise<void> {
    const figure = '<figure><img src="a.png" alt="x"><figcaption>w1</figcaption></figure>';
    const cases: Array<[string, string, RegExp]> = [
        ['list item', `<ol><li>${figure}<a href="https://e.x">w2</a></li></ol>`, /^1\. !\[x\]\(a\.png\) w1<br>\[w2\]\(https:\/\/e\.x\)$/m],
        ['list item, bold after', `<ol><li>${figure}<b>w2</b></li></ol>`, /^1\. !\[x\]\(a\.png\) w1<br>\*\*w2\*\*$/m],
        ['list item, code after', `<ol><li>${figure}<code>w2</code></li></ol>`, /^1\. !\[x\]\(a\.png\) w1<br>`w2`$/m],
        ['list item, text on both sides', `<ol><li>w0${figure}w2</li></ol>`, /^1\. w0<br>!\[x\]\(a\.png\) w1<br>w2( \{#[^}]*\})?$/m],
        ['pipe cell', `<table><tr><td>${figure}<b>w2</b></td><td>w3</td></tr></table>`, /!\[x\]\(a\.png\) w1<br>\*\*w2\*\*/],
        ['definition', `<dl><dt>w0</dt><dd>${figure}<b>w2</b></dd></dl>`, /!\[x\]\(a\.png\) w1<br>\*\*w2\*\*/],
        ['heading', `<h2>w0${figure}w2</h2>`, /^## w0<br>!\[x\]\(a\.png\) w1<br>w2( \{#[^}]*\})?$/m],
    ];
    for (const [name, html, expected] of cases) {
        const ast = await OfficeParser.parseOffice(Buffer.from(html), { fileType: 'html' });
        const md = await ast.to('md');
        assert.ok(expected.test(md.value), `Figure line: ${name}: ${JSON.stringify(md.value)}`);
        assert.ok(!/<figure|&lt;/.test(md.value), `Figure line: ${name} writes no figure markup: ${JSON.stringify(md.value)}`);
        assert.ok(md.messages.some(m => m.code === 'CONTENT_NOT_REPRESENTABLE'), `Figure line: ${name} is reported`);
        const back = await OfficeParser.parseOffice(Buffer.from(md.value), { fileType: 'md' });
        const text = (await back.to('text')).value;
        assert.ok(!/<\/?[a-z]|&lt;/i.test(text), `Figure line: ${name} reads back with no markup as text: ${JSON.stringify(text)}`);
        const tokens: string[] = String(text).match(/[A-Za-z0-9]+/g) ?? [];
        for (const word of ['w1', 'w2'].filter(x => html.includes(`>${x}<`) || html.includes(`${x}</`))) assert.ok(tokens.includes(word), `Figure line: ${name} keeps ${word} apart: ${JSON.stringify(text)}`);
        assert.strictEqual((await (await OfficeParser.parseOffice(Buffer.from(md.value), { fileType: 'md' })).to('md')).value, md.value, `Figure line: ${name} is stable at the second save`);
    }
    console.log(`✓ A captioned picture ends its line (${cases.length} shapes)`);
}

/** Generated figures and inline kinds in every line container: no joined, lost or markup words, settled by pass 3. */
export async function testFigureLinesGenerated(): Promise<void> {
    for (const [label, count, seed] of [['fast', 1500, 777], ['full', 12000, 31337]] as const) {
        const started = Date.now();
        const result = await runFigureLineGenerator(count, seed);
        console.log(`  figure lines ${label}: docs=${result.docs} distinct ASTs=${result.distinctAsts} Markdown=${result.distinctMarkdown} kind pairs=${result.distinctKindPairs} config x container=${result.distinctConfigContainers} in ${((Date.now() - started) / 1000).toFixed(1)}s`);
        assert.deepStrictEqual(result.failures, {}, `Figure lines (${label}, seed ${seed}): ${JSON.stringify(result.examples)}`);
        assert.ok(result.distinctAsts > count * 0.5 && result.distinctConfigContainers === 54, `Figure lines (${label}) cover the shapes: ${JSON.stringify(result)}`);
    }
    console.log('✓ Figure lines generated, no joined, lost or markup words, settled by pass 3');
}

/** A Markdown round trip of `content`: the first save, the text read back from it, and the second and third saves. */
async function pictureRoundTrip(content: unknown[], config?: Record<string, unknown>, attachments: unknown[] = []): Promise<{ md: string; text: string; md2: string; md3: string }> {
    const ast = { type: 'md', metadata: {}, attachments, content } as unknown as OfficeParserAST;
    const md = (await OfficeGenerator.generate(ast, 'md', config as any)).value as string;
    const back = await OfficeParser.parseOffice(Buffer.from(md), { fileType: 'md' });
    const md2 = (await OfficeGenerator.generate(back, 'md', config as any)).value as string;
    const md3 = (await OfficeGenerator.generate(await OfficeParser.parseOffice(Buffer.from(md2), { fileType: 'md' }), 'md', config as any)).value as string;
    return { md, text: String((await back.to('text')).value), md2, md3 };
}

/** Each of `words` is a word of its own in `text`, none joined to another, and no markup is shown as text. */
function assertWordsApart(text: string, words: string[], name: string): void {
    const tokens: string[] = text.match(/[A-Za-z0-9]+/g) ?? [];
    for (const word of words) assert.ok(tokens.includes(word), `${name}: keeps ${word} apart: ${JSON.stringify(text)}`);
    assert.ok(!/<\/?[a-z]|&lt;|&#/i.test(text), `${name}: no markup read as text: ${JSON.stringify(text)}`);
}

const r12Text = (value: string, extra: Record<string, unknown> = {}) => ({ type: 'text', text: value, ...extra });
const r12Figure = (ocr?: string, metadata: Record<string, unknown> = { url: 'a.png' }) => ({ type: 'image', text: ocr, metadata: { altText: 'alt1', caption: 'c1 c2', ...metadata } });
const r12Item = (children: unknown[]) => ({ type: 'list', metadata: { listType: 'unordered', indentation: 0, itemIndex: 0, listId: 'L' }, children });

/**
 * A figure's caption is written in every image mode, and the figure stands apart from the words beside it:
 * over maxInlineImageBytes its recognized text takes the picture's place and the caption follows it (it was
 * lost in a line); under 'ocr-text-only' and 'none' the caption is written (it was dropped, and the words
 * around the picture were glued to its text or split by nothing).
 */
export async function testFigureCaptionInEveryImageMode(): Promise<void> {
    const big = [{ type: 'image', name: 'p.png', mimeType: 'image/png', data: Buffer.alloc(3000, 7).toString('base64') }];
    const oversized = r12Figure('o1 o2', { attachmentName: 'p.png' });
    const run = [r12Text('w1 '), oversized, r12Text(' w2')];
    const cases: Array<[string, unknown[], Record<string, unknown>, unknown[], RegExp]> = [
        ['oversized, paragraph', [{ type: 'paragraph', children: run }], { maxInlineImageBytes: 100 }, big, /^w1\n\no1 o2\n\nc1 c2\n\nw2$/],
        ['oversized, list item', [r12Item(run)], { maxInlineImageBytes: 100 }, big, /^- w1<br>o1 o2<br>c1 c2<br>w2$/],
        ['oversized, heading', [{ type: 'heading', metadata: { level: 2 }, children: run }], { maxInlineImageBytes: 100 }, big, /^## w1<br>o1 o2<br>c1 c2<br> ?w2/],
        ['ocr-text-only, paragraph', [{ type: 'paragraph', children: [r12Text('w1 '), r12Figure('o1 o2'), r12Text(' w2')] }], { includeImages: 'ocr-text-only' }, [], /^w1\n\no1 o2\n\nc1 c2\n\nw2$/],
        ['ocr-text-only, bold then a space, paragraph', [{ type: 'paragraph', children: [r12Text('w1', { formatting: { bold: true } }), r12Text(' '), r12Figure('o1 o2'), r12Text(' w2')] }], { includeImages: 'ocr-text-only' }, [], /^\*\*w1\*\*\n\no1 o2\n\nc1 c2\n\nw2$/],
        ['ocr-text-only, bold then a space, list item', [r12Item([r12Text('w1', { formatting: { bold: true } }), r12Text(' '), r12Figure('o1 o2'), r12Text(' w2')])], { includeImages: 'ocr-text-only' }, [], /^- \*\*w1\*\*<br>o1 o2<br>c1 c2<br>w2$/],
        ['image-only, bold then a space, list item', [r12Item([r12Text('w1', { formatting: { bold: true } }), r12Text(' '), r12Figure(), r12Text('w2', { formatting: { bold: true } })])], {}, [], /^- \*\*w1\*\*<br>!\[alt1\]\(a\.png\) c1 c2<br>\*\*w2\*\*$/],
        ['ocr-text-only, two lines after a word', [{ type: 'paragraph', children: [r12Text('w1'), r12Figure('o1\no2')] }], { includeImages: 'ocr-text-only' }, [], /^w1\n\no1 {2}\no2\n\nc1 c2$/],
        ['image+ocr-text, two lines, list item', [r12Item([r12Text('w1 '), r12Figure('o1\no2'), r12Text(' w2')])], { includeImages: 'image+ocr-text' }, [], /^- w1<br>!\[alt1\]\(a\.png\) c1 c2<br>o1<br>o2<br>w2$/],
        ['none, paragraph', [{ type: 'paragraph', children: [r12Text('w1 '), r12Figure(), r12Text(' w2')] }], { includeImages: 'none' }, [], /^w1\n\nc1 c2\n\nw2$/],
        ['none, heading', [{ type: 'heading', metadata: { level: 2 }, children: [r12Text('w1 '), r12Figure(), r12Text(' w2')] }], { includeImages: 'none' }, [], /^## w1<br>c1 c2<br> ?w2/],
    ];
    for (const [name, content, config, attachments, expected] of cases) {
        const { md, text, md2, md3 } = await pictureRoundTrip(content, config, attachments);
        assert.ok(expected.test(md.trim()), `Caption in every mode: ${name}: ${JSON.stringify(md)}`);
        const words = ['w1', 'w2', 'c1', 'c2', ...(/o1/.test(md) ? ['o1', 'o2'] : [])].filter(word => JSON.stringify(content).includes(word));
        assertWordsApart(text, words, `Caption in every mode: ${name}`);
        assert.strictEqual(md3, md2, `Caption in every mode: ${name} settles by the third save`);
    }
    console.log(`✓ A figure's caption in every image mode (${cases.length} shapes)`);
}

/**
 * A footnote whose body is a figure alone closes the figure on a continuation line, so the definition is
 * read as blocks: written on one line, it was read back as the note's text, the figure's markup with it.
 */
export async function testFigureAloneInFootnote(): Promise<void> {
    const configs: Array<[string, Record<string, unknown> | undefined]> = [['default', undefined], ['fallbackToHtml false', { mdConfig: { fallbackToHtml: false } }],
        ['line breaks off', { mdConfig: { fallbackToHtml: { itemLineBreaks: false, cellLineBreaks: false } } }]];
    for (const [name, config] of configs) {
        const note = { type: 'note', metadata: { noteType: 'footnote', noteId: 'n1' }, children: [{ type: 'paragraph', children: [r12Figure()] }] };
        const { md, text, md2 } = await pictureRoundTrip([{ type: 'paragraph', children: [{ ...r12Text('w1'), notes: [note] }, r12Text(' w2')] }], config);
        assert.ok(md.includes('[^n1]: <figure><img src="a.png" alt="alt1"><figcaption>c1 c2</figcaption>\n    </figure>'), `Footnote figure: ${name}: ${JSON.stringify(md)}`);
        assertWordsApart(text, ['w1', 'w2', 'alt1', 'c1', 'c2'], `Footnote figure: ${name}`);
        assert.strictEqual(md2, md, `Footnote figure: ${name} is stable at the second save`);
    }
    console.log(`✓ A figure alone in a footnote reads back as a figure (${configs.length} configurations)`);
}

/**
 * A figure in a heading inside an HTML cell is written as the heading's lines (`<br>`), as a Markdown
 * heading writes it: a `<figure>` in the `<h2>` was read with its caption after the heading's words.
 */
export async function testFigureInHtmlCellHeading(): Promise<void> {
    const heading = { type: 'heading', metadata: { level: 2 }, children: [r12Text('w1 '), r12Figure(), r12Text(' w2')] };
    const cases: Array<[string, unknown, Record<string, unknown> | undefined]> = [
        ['commonmark cell', { type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [heading] }, { type: 'cell', children: [r12Text('w8')] }] }] }, { mdConfig: { dialect: 'commonmark' } }],
        ['spanned cell', { type: 'table', children: [{ type: 'row', children: [{ type: 'cell', metadata: { colSpan: 2 }, children: [heading] }] }, { type: 'row', children: [{ type: 'cell', children: [r12Text('w5')] }, { type: 'cell', children: [r12Text('w6')] }] }] }, undefined],
    ];
    for (const [name, table, config] of cases) {
        const { md, text, md2, md3 } = await pictureRoundTrip([{ type: 'paragraph', children: [r12Text('top')] }, table, { type: 'paragraph', children: [r12Text('end')] }], config);
        assert.ok(md.includes('<br><img src="a.png" alt="alt1"> c1 c2<br>') && !/<h2>[^]*<figure/.test(md), `HTML cell heading: ${name}: ${JSON.stringify(md)}`);
        assertWordsApart(text, ['w1', 'alt1', 'c1', 'c2', 'w2'], `HTML cell heading: ${name}`);
        assert.ok(/w1\b.*\bc1 c2\b.*\bw2\b/s.test(text), `HTML cell heading: ${name} keeps the caption before w2: ${JSON.stringify(text)}`);
        assert.strictEqual(md3, md2, `HTML cell heading: ${name} settles by the third save`);
    }
    console.log(`✓ A figure in an HTML cell's heading stays in its place (${cases.length} shapes)`);
}

/** A figure in an admonition's inline content gets one blank quote line before it, not three. */
export async function testFigureInAdmonitionLine(): Promise<void> {
    const content = [{ type: 'admonition', metadata: { admonitionType: 'note' }, children: [r12Text('w1 '), r12Figure(), r12Text(' w2')] }];
    for (const [name, config] of [['default', undefined], ['commonmark', { mdConfig: { dialect: 'commonmark' } }]] as const) {
        const { md, text } = await pictureRoundTrip(content, config);
        assert.ok(md.includes('> w1\n>\n> <figure><img src="a.png" alt="alt1"><figcaption>c1 c2</figcaption></figure>\n>\n>'), `Admonition figure: ${name}: ${JSON.stringify(md)}`);
        assert.ok(!/\n>\n>\n/.test(md), `Admonition figure: ${name} has no run of blank quote lines: ${JSON.stringify(md)}`);
        assertWordsApart(text, ['w1', 'alt1', 'c1', 'c2', 'w2'], `Admonition figure: ${name}`);
    }
    console.log('✓ A figure in an admonition line, one blank line before it');
}

/**
 * HTML and EPUB write a figure's caption in every image mode, as 8.6.0 did: a <figure> with its
 * <figcaption> where the picture or its recognized text is written, the caption alone as a paragraph (a
 * span in a heading) carrying the picture's ids where nothing else of it is. A picture without a caption
 * still writes nothing under 'none'.
 */
export async function testHtmlFigureCaptionInEveryImageMode(): Promise<void> {
    const figure = (ocr?: string) => ({ type: 'image', text: ocr, metadata: { url: 'a.png', altText: 'alt1', caption: 'c1 c2', anchorIds: ['x9'] } });
    const paragraph = (picture: unknown) => [{ type: 'paragraph', children: [r12Text('w1 '), picture, r12Text(' w2')] }];
    const heading = (picture: unknown) => [{ type: 'heading', metadata: { level: 2 }, children: [r12Text('w1 '), picture, r12Text(' w2')] }];
    const item = (picture: unknown) => [r12Item([r12Text('w1 '), picture, r12Text(' w2')])];
    const cases: Array<[string, string, unknown[], string]> = [
        ['none, paragraph', 'none', paragraph(figure('o1 o2')), '<p id="x9">c1 c2</p>'],
        ['none, list item', 'none', item(figure()), '<p id="x9">c1 c2</p>'],
        ['none, heading', 'none', heading(figure()), '<a id="x9" name="x9"></a><span class="caption">c1 c2</span>'],
        ['ocr-text-only, no text, paragraph', 'ocr-text-only', paragraph(figure()), '<p id="x9">c1 c2</p>'],
        ['ocr-text-only, paragraph', 'ocr-text-only', paragraph(figure('o1 o2')), '<figure><pre class="ocr-text" id="x9">o1 o2</pre><figcaption>c1 c2</figcaption></figure>'],
        ['ocr-text-only, list item', 'ocr-text-only', item(figure('o1 o2')), '<figure><pre class="ocr-text" id="x9">o1 o2</pre><figcaption>c1 c2</figcaption></figure>'],
        ['ocr-text-only, heading', 'ocr-text-only', heading(figure('o1 o2')), '<pre class="ocr-text" id="x9">o1 o2</pre> <span class="caption">c1 c2</span>'],
        ['image-only, paragraph', 'image-only', paragraph(figure('o1 o2')), '<figure><img src="a.png" alt="alt1" id="x9"><figcaption>c1 c2</figcaption></figure>'],
        ['image+ocr-text, paragraph', 'image+ocr-text', paragraph(figure('o1 o2')), '<figure><img src="a.png" alt="alt1" id="x9"><pre class="ocr-text">o1 o2</pre><figcaption>c1 c2</figcaption></figure>'],
    ];
    for (const [name, mode, content, expected] of cases) {
        const ast = { type: 'md', metadata: {}, attachments: [], content } as unknown as OfficeParserAST;
        const html = (await OfficeGenerator.generate(ast, 'html', { includeImages: mode as any, htmlConfig: { standalone: false } })).value as string;
        assert.ok(html.includes(expected), `HTML caption in every mode: ${name}: ${JSON.stringify(html)}`);
        const back = await OfficeParser.parseOffice(Buffer.from(html), { fileType: 'html' });
        assertWordsApart(String((await back.to('text')).value), ['w1', 'c1', 'c2', 'w2'], `HTML caption in every mode: ${name}`);
        assert.ok(JSON.stringify(back.content).includes('"x9"'), `HTML caption in every mode: ${name} keeps the picture's id`);
        if (mode === 'none' || mode === 'ocr-text-only') {
            const epub = (await OfficeGenerator.generate(ast, 'epub', { includeImages: mode as any })).value as Uint8Array;
            const epubText = String((await (await OfficeParser.parseOffice(Buffer.from(epub), { fileType: 'epub' })).to('text')).value);
            assertWordsApart(epubText, ['w1', 'c1', 'c2', 'w2'], `EPUB caption in every mode: ${name}`);
        }
    }
    const uncaptioned = { type: 'md', metadata: {}, attachments: [], content: paragraph({ type: 'image', text: 'o1', metadata: { url: 'a.png', altText: 'alt1' } }) } as unknown as OfficeParserAST;
    const none = (await OfficeGenerator.generate(uncaptioned, 'html', { includeImages: 'none', htmlConfig: { standalone: false } })).value as string;
    assert.ok(!/caption|<figure|o1/.test(none), `HTML caption in every mode: a picture without a caption writes nothing under 'none': ${JSON.stringify(none)}`);
    console.log(`✓ HTML and EPUB write a figure's caption in every image mode (${cases.length} shapes)`);
}

/** The picture matrix's writer and reader: this tree's. */
const pictureMatrixIo: PictureMatrixIo = {
    generate: async (ast, format, config) => (await OfficeGenerator.generate(ast, format, config)).value as string,
    parse: markdown => OfficeParser.parseOffice(Buffer.from(markdown), { fileType: 'md' }),
};

/**
 * Generated documents over the picture matrix (64 image and Markdown configurations, every container,
 * captioned pictures in every image mode): no joined, lost or markup words, settled by pass 3.
 */
export async function testPictureMatrixGenerated(): Promise<void> {
    for (const [label, count, seed] of [['fast', 2000, 4242], ['full', 10000, 99]] as const) {
        const started = Date.now();
        const result = await runPictureMatrix(pictureMatrixIo, { count, seed, captions: true, clean: true });
        console.log(`  picture matrix ${label}: docs=${result.docs} distinct ASTs=${result.distinctAsts} Markdown=${result.distinctMarkdown} configurations=${result.configCombos} documents with figures=${result.figureDocs} figures=${result.figures} data: picture ids (ledger)=${result.readerDataImgId} in ${((Date.now() - started) / 1000).toFixed(1)}s`);
        assert.deepStrictEqual(result.failures, {}, `Picture matrix (${label}, seed ${seed}): ${JSON.stringify(result.examples)}`);
        assert.ok(result.configCombos === 64 && result.distinctAsts > count * 0.95 && result.figures > count, `Picture matrix (${label}) covers the shapes: ${JSON.stringify({ ...result, fingerprints: undefined })}`);
    }
    console.log('✓ Picture matrix generated, no joined, lost or markup words, settled by pass 3');
}

/**
 * Generated documents over the picture matrix with no captions write the same Markdown (three passes) and
 * HTML as before the round 12 change: the fingerprints in helpers/pictureMatrixFingerprints.json were made
 * by runPictureMatrix with c56459ec's writer and reader.
 */
export async function testPictureMatrixCaptionFreeUnchanged(): Promise<void> {
    const kept: { count: number; seed: number; fingerprints: string[] } = JSON.parse(fs.readFileSync(PICTURE_MATRIX_FINGERPRINTS_FILE, 'utf8'));
    const result = await runPictureMatrix(pictureMatrixIo, { count: kept.count, seed: kept.seed, captions: false, clean: false });
    assert.strictEqual(result.figures, 0, 'Picture matrix without captions: no figure');
    assert.deepStrictEqual(result.failures, {}, `Picture matrix without captions: ${JSON.stringify(result.examples)}`);
    const changed = result.fingerprints.map((fingerprint, i) => fingerprint === kept.fingerprints[i] ? -1 : i).filter(i => i >= 0);
    assert.strictEqual(result.fingerprints.length, kept.fingerprints.length, 'Picture matrix without captions: as many documents as fingerprints');
    assert.deepStrictEqual(changed, [], `Picture matrix without captions: written as before, but ${changed.length} differ, the first document ${changed[0]}`);
    console.log(`✓ Picture matrix without captions written as before (docs=${result.docs} distinct ASTs=${result.distinctAsts} Markdown=${result.distinctMarkdown} configurations=${result.configCombos})`);
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

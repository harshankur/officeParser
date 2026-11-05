/**
 * Document shapes kept across readers and writers: XHTML read as HTML, list types per level, note
 * marks, quotes of several paragraphs, figure captions, disclosures and keyboard input, and sheets
 * written as document tables.
 */
import * as assert from 'assert';
import { unzipSync } from 'fflate';
import { OfficeParser } from '../../src/OfficeParser';
import type { NoteMetadata, OfficeContentNode, OfficeParserAST, ParagraphMetadata } from '../../src/types';

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
    const tex = (await ast.to('tex')).value;
    assert.ok(tex.includes('\\begin{quote}\na\n\nb \\textit{c}\n\\end{quote}'), `Quote: one quote environment: ${tex}`);
    const fromTex = await OfficeParser.parseOffice(Buffer.from(tex), { fileType: 'tex' });
    assert.strictEqual((await fromTex.to('md')).value, source, 'Quote: LaTeX reads a quote environment of paragraphs as one quote');
    // EPUB (XHTML) keeps it too.
    const fromEpub = await OfficeParser.parseOffice(Buffer.from((await ast.to('epub')).value as Uint8Array));
    assert.ok(((await fromEpub.to('md')).value).endsWith(source), 'Quote: EPUB keeps one quote of two paragraphs');
    console.log('✓ Quote paragraphs');
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
        const written = (await md.to(format)).value as Uint8Array;
        const back = await OfficeParser.parseOffice(Buffer.from(written));
        const html = (await back.to('html', { htmlConfig: { standalone: false } })).value;
        assert.deepStrictEqual(shown(html), expected, `Notes: ${format} keeps the marks`);
        const custom = notesOf(back)[2];
        assert.strictEqual(custom.text.trim(), 'Third.', `Notes: ${format} keeps the custom mark out of the note's text`);
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
    assert.strictEqual(epubNote.text.trim(), 'Figures are rounded.', 'Notes: a link back to the reference is the back-link, not text');
    assert.strictEqual((epubNote.metadata as NoteMetadata).mark, '1', 'Notes: the reference text is the mark');
    assert.deepStrictEqual(shown((await epubAst.to('html', { htmlConfig: { standalone: false } })).value), ['1'], 'Notes: and shown again');
    console.log('✓ Note marks');
}

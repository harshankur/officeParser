/**
 * Document shapes kept across readers and writers: XHTML read as HTML, list types per level, note
 * marks, quotes of several paragraphs, figure captions, disclosures and keyboard input, and sheets
 * written as document tables.
 */
import * as assert from 'assert';
import { OfficeParser } from '../../src/OfficeParser';

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

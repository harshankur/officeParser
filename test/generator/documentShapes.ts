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

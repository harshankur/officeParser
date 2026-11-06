/**
 * Word templates whose rendered parts must stay what officeParser 8.6.0 wrote: every shape the template
 * engine's later changes pass through (text that looks like a tag, runs of spaces, tags across runs,
 * blocks cut through runs and links, ids in loops, notes, headers), each with the data it renders.
 *
 * `test/files/template/docx-template-digests.json` holds 8.6.0's digest of each case's rendering
 * (made by rendering these cases with a v8.6.0 tree). A case added here needs its digest made the same
 * way; a case's digest changes only when the output it pins changes.
 */
import { createHash } from 'crypto';
import { strToU8, unzipSync, zipSync } from 'fflate';
import type { TemplateData } from '../../src/types';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
const r = (text: string, props = '') => `<w:r>${props}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (...runs: string[]) => `<w:p>${runs.map(run => run.startsWith('<') ? run : r(run)).join('')}</w:p>`;
const bold = '<w:rPr><w:b/></w:rPr>';

function docx(body: string, parts: Record<string, string> = {}): Uint8Array {
    const types = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';
    return zipSync({
        '[Content_Types].xml': strToU8(types),
        '_rels/.rels': strToU8('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
        'word/document.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`),
        ...Object.fromEntries(Object.entries(parts).map(([name, xml]) => [name, strToU8(xml)])),
    });
}

const items = [{ name: 'One & <1>', v: 2 }, { name: 'Two', v: '3' }, { name: 'Three', v: 0 }];
const data: TemplateData = { name: 'Ann & <Bob>', x: 1, a: true, b: false, items, v: 'V', total: 1250.5, empty: [] };

/** Each case: a name, the template, and the data it renders. */
export function docxTemplateCases(): { name: string; bytes: Uint8Array; data: TemplateData }[] {
    const cases: [string, string, Record<string, string>?][] = [
        ['fields and spaces', p('Dear {{ name }},   ', r('{{ total }}', bold), r('     '), '{{ v }}')],
        ['a tag across runs', p('{{', ' na', 'me }}', r(' and {{ v', bold), ' }}')],
        ['look-alikes beside a field', p('{{ v }} {a}} '.repeat(300) + '{{x}, {y}} {{ name } {a: {b}} {"a": {"b": 1}} f(){ {{x} \uFF5B\uFF5Bnom\uFF5D\uFF5D {{first name}} {{{v}}} {{ v }}}')],
        ['look-alikes only', p('{{ name } and { name }}') + p('{a}}'.repeat(500))],
        ['braces of every kind', p('{ } {} {{}} {{ }} }} {{ {{{ }}} a {{ b c }} d }}') + p('{{ v | nope }} {{ v | upper }}')],
        ['an inline if cut through a bold run', p(r('a{{ if a }}b', bold), 'c{{ endif }}d', r('{{ if b }}x{{ else }}y{{ endif }}', bold))],
        ['an inline loop in a link', p(`<w:hyperlink r:id="rId9">${r('L{{ for i in items }}', bold)}${r('[{{ i.name }}]')}</w:hyperlink>`, '{{ endfor }}e')],
        ['a paragraph loop with ids', p('{{ for i in items }}') + `<w:p w14:paraId="1A2B3C4D" w14:textId="77777777">${r('<w:bookmarkStart w:id="1" w:name="bm"/>{{ i.name }}<w:bookmarkEnd w:id="1"/>')}</w:p>` + p('{{ endfor }}') + p('after')],
        ['a bookmark across a loop edge', p('<w:bookmarkStart w:id="5" w:name="edge"/>start') + p('{{ for i in items }}') + p('{{ i.v }}<w:bookmarkEnd w:id="5"/>') + p('{{ endfor }}')],
        ['a comment range across a loop edge', `<w:p><w:commentRangeStart w:id="0"/>${r('c')}</w:p>` + p('{{ for i in items }}') + `<w:p>${r('{{ i.v }}')}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>` + p('{{ endfor }}'),
            { 'word/comments.xml': `<?xml version="1.0"?><w:comments ${W}><w:comment w:id="0" w:author="A">${p('note {{ name }}')}</w:comment></w:comments>` }],
        ['a footnote in a loop', p('{{ for i in items }}') + `<w:p>${r('{{ i.name }}')}<w:r><w:footnoteReference w:id="2"/></w:r></w:p>` + p('{{ endfor }}'),
            { 'word/footnotes.xml': `<?xml version="1.0"?><w:footnotes ${W}><w:footnote w:id="2">${p('fn {{ i.name }} {a}}')}</w:footnote></w:footnotes>` }],
        ['a table row loop', `<w:tbl><w:tr w14:paraId="00000001">${`<w:tc>${p('{{ for i in items }}{{ i.name }}')}</w:tc><w:tc>${p('{{ i.v }}{{ endfor }}')}</w:tc>`}</w:tr></w:tbl>` + p('')],
        ['a picture in a loop', p('{{ for i in items }}') + p(`<w:r><w:drawing><wp:inline><wp:docPr id="7" name="Pic"/></wp:inline></w:drawing></w:r>`) + p('{{ endfor }}')],
        ['a header and footer', p('{{ name }}'), { 'word/header1.xml': `<?xml version="1.0"?><w:hdr ${W}>${p('H {{ v }} {{ v }')}</w:hdr>`, 'word/footer1.xml': `<?xml version="1.0"?><w:ftr ${W}>${p('{{ if a }}F{{ endif }}')}</w:ftr>` }],
        ['a kept block', p('{{ if missing }}k {a}} {{ v }}{{ endif }}') + p('{{ for z in nothing }}z{{ endfor }}')],
        ['an empty loop and a missing field', p('{{ for i in empty }}x{{ endfor }}{{ nope }}')],
    ];
    return cases.map(([name, body, parts]) => ({ name, bytes: docx(body, parts), data: { ...data } }));
}

/** A rendering's digest: each entry's name and content, in order, so the zip's compression is not pinned. */
export function digestOf(rendered: Uint8Array): string {
    const hash = createHash('sha256');
    for (const [name, content] of Object.entries(unzipSync(rendered))) {
        hash.update(name);
        hash.update('\u0000');
        hash.update(content);
        hash.update('\u0000');
    }
    return hash.digest('hex');
}

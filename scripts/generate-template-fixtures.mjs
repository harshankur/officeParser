/**
 * Generates the DOCX template fixture for the templating tests: `test/files/template/invoice.docx`.
 *
 * Purpose-built to exercise the hard cases of run-aware placeholder substitution:
 *  - `{{name}}` is deliberately SPLIT across three `<w:t>` runs (`{{` | `na` | `me}}`), the Word quirk
 *    the substitution has to survive.
 *  - `{{amount}}` sits in a BOLD run, to prove a value adopts the formatting of its placeholder's run.
 *  - `{{item}}` lives in a TABLE CELL, `{{company}}` in a HEADER part, `{{note}}` receives a multiline
 *    value (-> `<w:br/>`), and `{{missing}}` has no data field (for the onMissing modes).
 *  - Block tags in each of their scopes: an `{{ if paid }}` ... `{{ else }}` ... `{{ endif }}` over whole
 *    paragraphs (each tag alone in its paragraph, so the tag paragraphs are removed), a line-items table
 *    whose second row repeats once per `{{ for line in lines }}` item, a discount row an `{{ if }}`
 *    keeps or drops, and an inline loop holding a condition (`{{ if not loop.last }}`) inside one paragraph.
 *  - Values formatted with pictures: `{{ total | format "$#,##0.00" }}` and `{{ dueDate | format "MMMM d, yyyy" }}`.
 *
 * And the Template Studio's conditions sample, `test/files/template/conditions.docx`: a condition that
 * compares without case through a filter (`{{ if (answer | lower) == "yes" }}`), one ordering two number
 * texts (`{{ if score > target }}`, which spreadsheet data writes as text), one with `<>`, and an amount
 * written in either of two ways (`12.50`, `1.234,50`) through a number picture. Deterministic (a fixed zip
 * time), so running it again writes the same bytes.
 *
 * Run: `node scripts/generate-template-fixtures.mjs`
 */
import { zipSync } from 'fflate';
import fs from 'fs';
import path from 'path';

const OUT = path.join('test', 'files', 'template');
const S = s => Buffer.from(s, 'utf8');

const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`;
const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`;

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const r = (t, bold) => `<w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`;

// {{name}} split across three runs on purpose.
const greeting = `<w:p>${r('Dear {{')}${r('na')}${r('me}}, ')}${r('welcome.')}</w:p>`;
const balance = `<w:p>${r('Your balance is ')}${r('{{amount}}', true)}${r(' as of {{date}}.')}</w:p>`;
const table = `<w:tbl><w:tr><w:tc><w:p>${r('Item')}</w:p></w:tc><w:tc><w:p>${r('{{item}}')}</w:p></w:tc></w:tr></w:tbl>`;
const missing = `<w:p>${r('Ref: {{missing}}')}</w:p>`;
const note = `<w:p>${r('Note: {{note}}')}</w:p>`;
const para = (...runs) => `<w:p>${runs.join('')}</w:p>`;
// Paragraph-scope conditional: each tag alone in its paragraph.
const paid = para(r('{{ if paid }}')) + para(r('Paid in full. Thank you!', true)) + para(r('{{ else }}'))
    + para(r('Please pay {{amount}} by {{due}}.')) + para(r('{{ endif }}'));
// Row-scope loop and conditional: a header row, one row repeated per line item, a discount row.
const cell = (...runs) => `<w:tc><w:p>${runs.join('')}</w:p></w:tc>`;
const lines = '<w:tbl>'
    + `<w:tr>${cell(r('Description', true))}${cell(r('Qty', true))}${cell(r('Price', true))}</w:tr>`
    + `<w:tr>${cell(r('{{ for line in lines }}{{ line.description }}'))}${cell(r('{{ line.qty }}'))}${cell(r('{{ line.price }}{{ endfor }}'))}</w:tr>`
    + `<w:tr>${cell(r('{{ if discount }}Discount'))}${cell(r(''))}${cell(r('-{{ discount }}{{ endif }}'))}</w:tr>`
    + '</w:tbl>';
// Values formatted with Word-style pictures: a number and a date.
const totals = para(r('Total due: {{ total | format "$#,##0.00" }} by {{ dueDate | format "MMMM d, yyyy" }}.'));
// Inline loop with a condition inside one paragraph.
const contacts = para(r('Contacts: {{ for c in contacts }}{{ c.name }}{{ if not loop.last }}, {{ endif }}{{ endfor }}'));
const sect = `<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/></w:sectPr>`;
const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${greeting}${balance}${table}${missing}${note}${paid}${lines}${totals}${contacts}${sect}</w:body></w:document>`;
const header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}><w:p>${r('{{company}} - Confidential')}</w:p></w:hdr>`;

fs.mkdirSync(OUT, { recursive: true });
fs.writeFileSync(path.join(OUT, 'invoice.docx'), Buffer.from(zipSync({
    '[Content_Types].xml': S(ct),
    '_rels/.rels': S(rels),
    'word/_rels/document.xml.rels': S(docRels),
    'word/document.xml': S(document),
    'word/header1.xml': S(header),
})));

// The conditions sample: one paragraph per condition, its label bold (text written as XML: `<>` as `&lt;&gt;`).
const MTIME = new Date(2001, 0, 1, 0, 0, 0);
const conditionsCt = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
const heading = t => `<w:p><w:pPr><w:spacing w:after="240"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const line = (label, text) => para(r(label, true), r(text));
const conditionsBody = heading('Event reply for {{ name }}')
    + line('Reply: ', '{{ if (answer | lower) == "yes" }}Attending{{ else }}Not attending{{ endif }}')
    + line('Score: ', '{{ score }} against a target of {{ target }}: {{ if score > target }}Above target{{ else }}At or below target{{ endif }}')
    + line('Amount: ', '{{ amount | format "#,##0.00" }}{{ if status &lt;&gt; "paid" }} (Payment due){{ endif }}');
const conditions = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${conditionsBody}<w:sectPr/></w:body></w:document>`;
fs.writeFileSync(path.join(OUT, 'conditions.docx'), Buffer.from(zipSync({
    '[Content_Types].xml': S(conditionsCt),
    '_rels/.rels': S(rels),
    'word/document.xml': S(conditions),
}, { mtime: MTIME })));
console.log('wrote', fs.readdirSync(OUT).join(', '));

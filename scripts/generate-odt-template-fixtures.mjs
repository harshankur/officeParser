/**
 * Generates the OpenDocument template fixtures for the templating tests, written from the ODF 1.3 spec
 * (part 3, the schema; part 2, the package), not by officeParser's own writer:
 *
 * - `test/files/template/invoice.odt`: the same template as `invoice.docx`, as LibreOffice keeps text.
 *   `{{name}}` is split across a span (`{{` | `na` | `me}}`), `{{amount}}` sits in a bold span,
 *   `{{item}}` in a table cell, `{{company}}` in the master page's header (styles.xml), `{{note}}` takes
 *   a value of two lines, and `{{missing}}` has no data. Blocks: an if/else over whole paragraphs, a
 *   line-items table whose heading row is in `table:table-header-rows` and whose second row repeats per
 *   item, a discount row an if keeps or drops, values through pictures, and an inline loop with a
 *   condition inside it.
 * - `test/files/template/invoice.ott`: the same, as a template (`application/vnd.oasis.opendocument.text-template`).
 * - `test/files/template/shapes.odt`: every other place a tag can be and what a copy must renumber: a
 *   tag with `<text:s/>` in it, a heading, a loop over list items, a loop over paragraphs holding a
 *   bookmark range, a paragraph `xml:id`, a text box in a named frame, a footnote and a comment range;
 *   a loop over a section; a loop inside a table's heading rows; a branch in the header, a field in the
 *   left-page header, a loop in the footer; deleted text kept for tracked changes (never read).
 * - `test/files/template/near-tags.odt`: text that looks like a tag and is not one, in the body, a
 *   table, a heading, a list, a footnote, a comment, a text box, the header and the footer; and braces
 *   in prose and JSON that are none.
 *
 * Deterministic (a fixed zip time), so running it again writes the same bytes.
 * Run: `node scripts/generate-odt-template-fixtures.mjs`
 */
import { zipSync } from 'fflate';
import fs from 'fs';
import path from 'path';

const OUT = path.join('test', 'files', 'template');
const S = s => Buffer.from(s, 'utf8');
const MTIME = new Date(2001, 0, 1, 0, 0, 0);

const ODT = 'application/vnd.oasis.opendocument.text';
const OTT = 'application/vnd.oasis.opendocument.text-template';

const NS = [
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
    'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
    'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
    'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
    'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
    'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
    'xmlns:dc="http://purl.org/dc/elements/1.1/"',
    'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"',
].join(' ');

/** Automatic styles: a bold and an italic span, a bulleted list. */
const AUTOMATIC = '<office:automatic-styles>'
    + '<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>'
    + '<style:style style:name="T2" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>'
    + '<style:style style:name="fr1" style:family="graphic"><style:graphic-properties fo:border="0.5pt solid #000000"/></style:style>'
    + '<text:list-style style:name="L1"><text:list-level-style-bullet text:level="1" text:bullet-char="&#x2022;"><style:list-level-properties text:space-before="0.635cm" text:min-label-width="0.635cm"/></text:list-level-style-bullet></text:list-style>'
    + '</office:automatic-styles>';

const content = body => `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.3">${AUTOMATIC}<office:body><office:text>${body}</office:text></office:body></office:document-content>`;

/** Styles: a page layout, and the Standard master page with a header (and a left-page header) and a footer. */
const styles = ({ header, headerLeft, footer }) => `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS} office:version="1.3">`
    + '<office:styles><style:style style:name="Standard" style:family="paragraph" style:class="text"/>'
    + '<style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:default-outline-level="1"><style:text-properties fo:font-size="16pt" fo:font-weight="bold"/></style:style></office:styles>'
    + '<office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm"/>'
    + '<style:header-style><style:header-footer-properties fo:min-height="0.5cm" fo:margin-bottom="0.3cm"/></style:header-style>'
    + '<style:footer-style><style:header-footer-properties fo:min-height="0.5cm" fo:margin-top="0.3cm"/></style:footer-style></style:page-layout></office:automatic-styles>'
    + '<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1">'
    + `<style:header>${header}</style:header>`
    + (headerLeft ? `<style:header-left>${headerLeft}</style:header-left>` : '')
    + (footer ? `<style:footer>${footer}</style:footer>` : '')
    + '</style:master-page></office:master-styles></office:document-styles>';

const meta = `<?xml version="1.0" encoding="UTF-8"?><office:document-meta ${NS} office:version="1.3"><office:meta><meta:generator>officeParser template fixtures</meta:generator><dc:title>Template fixture</dc:title></office:meta></office:document-meta>`;

const manifest = mediaType => '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">'
    + `<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${mediaType}"/>`
    + '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
    + '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>'
    + '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>'
    + '</manifest:manifest>';

/** A package: `mimetype` first and stored, as ODF part 2 (3.3) requires. */
const odf = (mediaType, body, styleParts) => Buffer.from(zipSync({
    mimetype: [S(mediaType), { level: 0, mtime: MTIME }],
    'META-INF/manifest.xml': [S(manifest(mediaType)), { mtime: MTIME }],
    'content.xml': [S(content(body)), { mtime: MTIME }],
    'styles.xml': [S(styles(styleParts)), { mtime: MTIME }],
    'meta.xml': [S(meta), { mtime: MTIME }],
}));

const p = (inner, style = 'Standard') => `<text:p text:style-name="${style}">${inner}</text:p>`;
const bold = text => `<text:span text:style-name="T1">${text}</text:span>`;
const cell = inner => `<table:table-cell office:value-type="string">${p(inner)}</table:table-cell>`;
const row = (...cells) => `<table:table-row>${cells.join('')}</table:table-row>`;

// ─── invoice ───
// {{name}} split across a span: `{{` | `na` | `me}}`.
const greeting = p('Dear {{<text:span text:style-name="T2">na</text:span>me}}, welcome.');
const balance = p(`Your balance is ${bold('{{amount}}')} as of {{date}}.`);
const items = `<table:table table:name="Items"><table:table-column table:number-columns-repeated="2"/>${row(cell('Item'), cell('{{item}}'))}</table:table>`;
const missing = p('Ref: {{missing}}');
const note = p('Note: {{note}}');
const paid = p('{{ if paid }}') + p(bold('Paid in full. Thank you!')) + p('{{ else }}') + p('Please pay {{amount}} by {{due}}.') + p('{{ endif }}');
const lines = '<table:table table:name="Lines"><table:table-column table:number-columns-repeated="3"/>'
    + `<table:table-header-rows>${row(cell(bold('Description')), cell(bold('Qty')), cell(bold('Price')))}</table:table-header-rows>`
    + row(cell('{{ for line in lines }}{{ line.description }}'), cell('{{ line.qty }}'), cell('{{ line.price }}{{ endfor }}'))
    + row(cell('{{ if discount }}Discount'), cell(''), cell('-{{ discount }}{{ endif }}'))
    + '</table:table>';
const totals = p('Total due: {{ total | format &quot;$#,##0.00&quot; }} by {{ dueDate | format &quot;MMMM d, yyyy&quot; }}.');
const contacts = p('Contacts: {{ for c in contacts }}{{ c.name }}{{ if not loop.last }}, {{ endif }}{{ endfor }}');
const invoiceBody = greeting + balance + items + missing + note + paid + lines + totals + contacts;
const invoiceStyles = { header: p('{{company}} - Confidential') };

// ─── shapes ───
const frame = (name, inner) => `<draw:frame draw:style-name="fr1" draw:name="${name}" text:anchor-type="as-char" svg:width="5cm" svg:height="1cm" draw:z-index="0"><draw:text-box>${inner}</draw:text-box></draw:frame>`;
const footnote = (id, inner) => `<text:note text:id="${id}" text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body>${p(inner)}</text:note-body></text:note>`;
const comment = (name, inner) => `<office:annotation office:name="${name}"><dc:creator>Reviewer</dc:creator><dc:date>2026-01-01T00:00:00</dc:date>${p(inner)}</office:annotation>`;
const shapesBody = '<text:tracked-changes><text:changed-region xml:id="ct1" text:id="ct1"><text:deletion><office:change-info><dc:creator>Editor</dc:creator><dc:date>2026-01-01T00:00:00</dc:date></office:change-info>'
    + p('Deleted {{ deleted }}') + '</text:deletion></text:changed-region></text:tracked-changes>'
    + '<text:h text:style-name="Heading_20_1" text:outline-level="1">Report for {{ client }}</text:h>'
    // A tag holding spaces LibreOffice keeps as <text:s/>.
    + p('Hello {{<text:s/>who<text:s text:c="2"/>}}!<text:change text:change-id="ct1"/>')
    // A loop over list items: the tag-only items go with their tags.
    + '<text:list text:style-name="L1">'
    + `<text:list-item>${p('{{ for t in tasks }}')}</text:list-item>`
    + `<text:list-item>${p('Task: {{ t }}')}</text:list-item>`
    + `<text:list-item>${p('{{ endfor }}')}</text:list-item>`
    + '</text:list>'
    // A loop over a paragraph holding everything a copy renumbers.
    + p('{{ for person in people }}')
    + `<text:p text:style-name="Standard" xml:id="person-para"><text:bookmark-start text:name="person"/>{{ person.name }}<text:bookmark-end text:name="person"/>`
    + ` ${frame('Badge', p('Badge: {{ person.name }}'))}${footnote('ftn1', 'About {{ person.name }}')}`
    + `${comment('c1', 'Check {{ person.name }}')}seen<office:annotation-end office:name="c1"/></text:p>`
    + p('{{ endfor }}')
    // A loop over a section.
    + p('{{ for part in parts }}')
    + `<text:section text:name="Part">${p('Part {{ loop.number }}: {{ part }}')}</text:section>`
    + p('{{ endfor }}')
    // A loop inside a table's heading rows, and a row loop below them.
    + '<table:table table:name="Grid"><table:table-column table:number-columns-repeated="2"/>'
    + `<table:table-header-rows>${row(cell('{{ for h in heads }}{{ h.a }}'), cell('{{ h.b }}{{ endfor }}'))}</table:table-header-rows>`
    + row(cell('{{ for r in rows }}{{ r.a }}'), cell('{{ r.b }}{{ endfor }}'))
    + '</table:table>';
const shapesStyles = {
    header: p('{{ if draft }}DRAFT {{ endif }}{{ company }}'),
    headerLeft: p('Left: {{ company }}'),
    footer: p('Authors: {{ for a in authors }}{{ a }}{{ if not loop.last }}, {{ endif }}{{ endfor }}'),
};

// ─── near-tags ───
const nearBody = p('Dear {{first name}},')
    + p('Your order {{ order }is ready.')
    + p('Total: {{ total | fromat &quot;0.00&quot; }}')
    + p('Bonjour \uFF5B\uFF5Bnom\uFF5D\uFF5D')
    + '<text:h text:style-name="Heading_20_1" text:outline-level="1">Report { client }} and {{ ok }}</text:h>'
    + `<table:table table:name="Near"><table:table-column/>${row(cell('{{ cell value }}'))}</table:table>`
    + `<text:list text:style-name="L1"><text:list-item>${p('{{ list item }}')}</text:list-item></text:list>`
    + p(`See${footnote('ftn1', '{{ note name }}')} and ${frame('Box', p('{{ box name }}'))}${comment('c1', '{{ comment name }}')}here<office:annotation-end office:name="c1"/>`)
    // Braces that are none: a single brace, a word in braces, JSON, a set in code, an empty pair.
    + p('Use { and } freely: {x}, {"a": {"b": 1}}, f(x) {return x;}, {{}} and {{ ok }}.');
const nearStyles = { header: p('{{header name}}'), footer: p('{{ footer }') };

fs.mkdirSync(OUT, { recursive: true });
fs.writeFileSync(path.join(OUT, 'invoice.odt'), odf(ODT, invoiceBody, invoiceStyles));
fs.writeFileSync(path.join(OUT, 'invoice.ott'), odf(OTT, invoiceBody, invoiceStyles));
fs.writeFileSync(path.join(OUT, 'shapes.odt'), odf(ODT, shapesBody, shapesStyles));
fs.writeFileSync(path.join(OUT, 'near-tags.odt'), odf(ODT, nearBody, nearStyles));
console.log('wrote', ['invoice.odt', 'invoice.ott', 'shapes.odt', 'near-tags.odt'].join(', '));

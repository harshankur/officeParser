/**
 * Generates the PDF fixtures for the PDF reader's table detection: tables as a browser and LibreOffice
 * print them, with and without tags, ruled and set by spacing alone. The sources are the HTML pages below;
 * the PDFs are committed, as the producers' output differs from one version to the next.
 *
 *  - `tables_borderless_chrome.pdf`: untagged, no rules, a page a table: an invoice whose header cells
 *    are set closer than a wide gap, a header over two lines, labels and their values (two columns), a
 *    calendar whose days are set close, and a header over sub-columns.
 *  - `tables_tagged_chrome.pdf`: tagged: a calendar whose first and last weeks leave out their empty
 *    cells, a ruled table of merged cells, and labels and values with no borders (tagged as no table).
 *  - `tables_borderless_libreoffice.pdf`: untagged, no rules, full width: a header over two lines, merged
 *    cells set in the middle of the rows they cover, and labels and their values.
 *  - `tables_tagged_libreoffice.pdf`: tagged, ruled: merged cells.
 *  - `two_columns_not_tables_chrome.pdf`: untagged: two lists of links side by side, code with comments
 *    in a monospaced font, a bulleted list. None is a table.
 *
 * Run: `node scripts/generate-table-pdf-fixtures.mjs` (needs puppeteer's browser and LibreOffice; set
 * `CHROME` to another Chromium's path, and `SOFFICE` to LibreOffice's `soffice` where it is not at
 * its usual place).
 */
import { execFileSync } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';
import puppeteer from 'puppeteer';

const OUT = path.join('test', 'files', 'pdf');
const SOFFICE = process.env.SOFFICE || (process.platform === 'darwin' ? '/Applications/LibreOffice.app/Contents/MacOS/soffice' : 'soffice');

const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;');
/** A cell: a string, or { t, cs, rs } for one spanning columns or rows. */
const cell = (x, tag) => {
    const c = typeof x === 'string' ? { t: x } : x;
    const span = `${c.cs > 1 ? ` colspan="${c.cs}"` : ''}${c.rs > 1 ? ` rowspan="${c.rs}"` : ''}`;
    return `<${tag}${span}>${esc(c.t).replace(/\n/g, '<br>')}</${tag}>`;
};
/** A table: its rows, how many are headers, ruled or not, and full width or as wide as its text. */
const table = (rows, { header = 1, ruled = false, wide = false } = {}) =>
    `<table${wide ? ' width="100%"' : ''} ${ruled ? 'border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse"' : 'border="0" cellpadding="6" cellspacing="0"'}>\n` +
    rows.map((r, i) => `<tr>${r.map(x => cell(x, i < header ? 'th' : 'td')).join('')}</tr>`).join('\n') + '\n</table>';
const page = (title, body) => `<h1>${esc(title)}</h1>\n<p>The table below lists the figures as they were reported at the end of the month, and the notes after it explain how they were gathered by the office.</p>\n${body}\n<p>Figures are final unless a correction is issued in writing. Questions about this page go to the records office, which answers them within ten working days.</p>`;
const doc = pages => `<!doctype html><html><head><meta charset="utf-8"><title>Tables</title><style>body{font-family:Helvetica,Arial,sans-serif;font-size:11pt;margin:40px} th,td{text-align:left;vertical-align:top} section+section{break-before:page}</style></head><body>\n${pages.map(p => `<section>${p}</section>`).join('\n')}\n</body></html>`;

const invoice = [['Date', 'Description', 'Quantity', 'Unit price', 'Amount'], ...Array.from({ length: 6 }, (_, i) => [`2026-10-0${i + 1}`, ['Consulting', 'Travel', 'Licence', 'Hosting'][i % 4], String(i % 5 + 1), (40 + i * 2.5).toFixed(2), ((i % 5 + 1) * (40 + i * 2.5)).toFixed(2)])];
const twoLine = [['Item\ncode', 'Quantity\nordered', 'Unit price\n(EUR)', 'Delivery\ndate'], ['A-100', '12', '4.50', '2026-11-02'], ['B-220', '3', '19.99', '2026-11-05'], ['C-310', '40', '0.75', '2026-11-09']];
const labels = [['Name', 'Harriet Example'], ['Date of birth', '1984-03-12'], ['Address', '12 High Street, Springfield'], ['Phone', '+44 20 7946 0000'], ['Membership', 'Gold']];
// October 2026: the 1st is a Thursday; weeks start on Monday.
const calendar = [['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'], ...Array.from({ length: 5 }, (_, w) => Array.from({ length: 7 }, (_, d) => { const day = w * 7 + d - 2; return day >= 1 && day <= 31 ? String(day) : ''; }))];
const spans = [[{ t: 'Region', rs: 2 }, { t: 'Sales', cs: 2 }, { t: 'Cost', cs: 2 }], ['Q1', 'Q2', 'Q1', 'Q2'], [{ t: 'North', rs: 2 }, '120', '130', '80', '85'], ['140', '150', '90', '95'], ['South', '95', '100', '70', '72']];

const chromeBorderless = doc([
    page('Invoice 2026-17', table(invoice)),
    page('Purchase order', table(twoLine)),
    page('Applicant details', table(labels, { header: 0 })),
    page('October 2026', table(calendar)),
    page('Regional figures', table(spans, { header: 2 })),
]);
const chromeTagged = doc([
    page('October 2026', table(calendar)),
    page('Regional figures', table(spans, { header: 2, ruled: true })),
    page('Applicant details', table(labels, { header: 0 })),
]);
const libreBorderless = doc([
    page('Purchase order', table(twoLine, { wide: true })),
    page('Regional figures', table(spans, { header: 2, wide: true })),
    page('Applicant details', table(labels, { header: 0, wide: true })),
]);
const libreTagged = doc([page('Regional figures', table(spans, { header: 2, ruled: true, wide: true }))]);
const notTables = `<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:Helvetica,Arial,sans-serif;font-size:11pt;margin:40px} .cols{display:grid;grid-template-columns:1fr 1fr} .cols div p{margin:4px 0} pre{font-size:10pt}</style></head><body>
<h1>Site map</h1>
<div class="cols"><div><p><b>CORE APIS</b></p><p>OfficeConverter</p><p>OfficeGenerator</p><p>OfficeTemplate</p></div><div><p><b>CONFIGURATION</b></p><p>Global Config</p><p>Generator Config</p></div></div>
<h2>Code</h2>
<pre>BT                      % Begin a text object
/F1 12 Tf               % Select font F1 at 12pt
1 0 0 1 72 710 Tm       % Text matrix: position
(The) Tj                % Show the glyphs
ET                      % End the text object</pre>
<h2>Steps</h2>
<ul><li>Open the file you want to read.</li><li>Choose the pages to keep.</li><li>Save the result under a new name.</li></ul>
</body></html>`;

const browser = await puppeteer.launch({ headless: true, ...(process.env.CHROME ? { executablePath: process.env.CHROME } : {}) });
const tab = await browser.newPage();
for (const [name, html, tagged] of [['tables_borderless_chrome', chromeBorderless, false], ['tables_tagged_chrome', chromeTagged, true], ['two_columns_not_tables_chrome', notTables, false]]) {
    await tab.setContent(html);
    fs.writeFileSync(path.join(OUT, `${name}.pdf`), await tab.pdf({ format: 'A4', tagged }));
}
await browser.close();

// LibreOffice opens the HTML as a text document; a profile of its own says whether to tag the PDF.
const work = fs.mkdtempSync(path.join(os.tmpdir(), 'table-fixtures-'));
try {
    for (const [name, html, tagged] of [['tables_borderless_libreoffice', libreBorderless, false], ['tables_tagged_libreoffice', libreTagged, true]]) {
        const profile = path.join(work, `profile-${tagged}`);
        fs.mkdirSync(path.join(profile, 'user'), { recursive: true });
        fs.writeFileSync(path.join(profile, 'user', 'registrymodifications.xcu'), `<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Common/Filter/PDF/Export"><prop oor:name="UseTaggedPDF" oor:op="fuse"><value>${tagged}</value></prop></item>
</oor:items>
`);
        const src = path.join(work, `${name}.html`);
        fs.writeFileSync(src, html);
        execFileSync(SOFFICE, [`-env:UserInstallation=file://${profile}`, '--headless', '--infilter=HTML (StarWriter)', '--convert-to', 'pdf:writer_pdf_Export', '--outdir', work, src], { stdio: 'inherit' });
        fs.copyFileSync(path.join(work, `${name}.pdf`), path.join(OUT, `${name}.pdf`));
    }
} finally {
    fs.rmSync(work, { recursive: true, force: true });
}
console.log('Wrote the table fixtures to', OUT);

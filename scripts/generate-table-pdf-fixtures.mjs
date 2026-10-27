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
 *  - `prose_not_tables_chrome.pdf`: untagged, a page each: names in two columns, two poems side by side,
 *    two addresses side by side, and justified articles in three narrow columns. None is a table.
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

const prose = `<!doctype html><html><head><meta charset="utf-8"><style>body{margin:40px} section+section{break-before:page}</style></head><body>
<section style="font-family: sans-serif; font-size: 11pt">
<p>The members of the committee for this year, in the order they joined it, are listed below by their full names.</p>
<div style="column-count:2; width: 500px">
Alice Brown<br>Bob Carter<br>Carol Davis<br>Dan Evans<br>Erin Foster<br>Frank Green<br>Grace Hill<br>Henry Irwin<br>Ivy Jones<br>Jack King
</div>
<p>Each member serves for two years and may stand again once, after which a year must pass before standing again.</p>
</section>
<section style="font-family: serif; font-size: 12pt">
<p>Two short poems set side by side, as an anthology prints them on facing halves of the page.</p>
<div style="display:flex;gap:120px">
<div>The rain came down<br>upon the town<br>and all was grey<br>the livelong day</div>
<div>The sun rose bright<br>with morning light<br>and birds took wing<br>to greet the spring</div>
</div>
<p>Both are by the same hand, written a year apart, and they were first printed together in a small magazine.</p>
</section>
<section style="font-family: sans-serif; font-size: 11pt">
<div style="display:flex;gap:200px">
<div>John Smith<br>12 Main Street<br>Springfield 12345<br>United States</div>
<div>Jane Doe<br>45 Elm Road, Flat 3<br>Shelbyville<br>Canada</div>
</div>
<p>Dear Jane, thank you for your letter of last week. I was glad to hear that the move went well and that the new place suits you all.</p>
</section>
<section style="font-family:serif;font-size:10pt"><h2>Local news</h2><p>An introduction that runs across the whole width of the page, above the columns of the story below it.</p><div style="column-count:3;width:520px;text-align:justify;column-gap:14px"><p>River played farmers their mill past their river about spring rain across arrived from from wars wheat across snow old where across hills across merchants evening past talked kings each letters kings autumn hills distant slowly were while winter bridges from valleys letters stories winter hills.</p><p>Summer of and harvest with prices prices merchants oats of merchants farmers planted where past rain from in valleys built were parents hills about distant arrived.</p><p>Old oats hills letters valleys evening old roads children oats across roads and merchants spices river weather merchants mill towns about while markets cloth summer ships river talked prices across the stories children markets kings the.</p><p>Arrived storms where towns letters harvest slowly rain evening rain old stories stories were weather each brought farmers hills rain evening while weather cloth prices roads valleys cloth oats planted letters.</p><p>And built letters wheat brought with past markets of bridges across cloth in mill mill hills planted wheat spring valleys news old barley cloth summer markets of storms each roads spices slowly hills of spices children kings letters the across prices arrived slowly markets mill.</p><p>Slowly about planted wars summer harvest the autumn brought news letters farmers arrived rain stories wheat distant and wheat evening summer slowly and slowly across and news cloth mill rain barley and slowly prices snow where kings about weather barley past ships beyond.</p><p>Cloth old weather hills parents arrived prices in towns arrived letters ran their summer spring built rain children kings old mill mill wheat were brought prices the cloth rain prices wheat past arrived valleys summer summer spring mill harvest snow where children planted autumn in summer about bridges distant about.</p><p>While across slowly built farmers while spices kings of while hills across were harvest with talked beyond slowly mill about harvest each valleys farmers evening planted river ships ships towns played with.</p><p>Children ships where oats storms towns beyond of winter in news played parents news oats oats talked from of bridges mill where mill the distant kings with distant roads parents built arrived of winter roads the bridges arrived about built.</p><p>News summer distant old prices barley their harvest storms summer beyond in spring children river talked kings each oats winter river built slowly evening storms old bridges snow beyond played hills across across slowly harvest markets roads cloth old hills each harvest where.</p></div></section>
<section style="font-family:serif;font-size:10pt"><h2>Local news</h2><div style="column-count:3;width:540px;text-align:justify;column-gap:14px"><p>Built storms from summer wars across were ships evening towns prices storms merchants kings markets ran valleys stories of storms farmers stories rain rain evening farmers barley towns arrived.</p><p>Letters played mill oats past slowly evening ships kings arrived winter letters spring parents across stories stories each spices storms hills bridges cloth were old talked storms mill spring news towns slowly.</p><p>Letters merchants farmers ships merchants spring beyond played snow oats summer distant their old markets talked bridges prices merchants played wars storms arrived announced were wheat rain harvest in and bridges.</p><p>Cloth bridges committee of arrived brought bridges brought letters built their beyond talked farmers their ships markets snow of storms summer of spring yesterday distant stories past beyond children committee cloth committee brought evening weather old wheat autumn stories rain snow towns valleys about hills built rain committee.</p><p>Their and arrived while were summer parents stories hills played each ships cloth of parents the mill talked towns harvest each children children wheat each the hills harvest and harvest.</p><p>Arrived hills farmers talked merchants while built weather spring farmers built while river with while oats barley spring barley with kings spring their roads farmers parents about mill committee summer winter barley across towns were spring planted announced government snow built cloth played the about harvest hills parents.</p><p>Roads merchants slowly committee snow talked across summer announced old bridges children cloth children spring summer harvest slowly wheat summer valleys evening brought barley past slowly while were committee snow.</p><p>Played arrived across ships beyond cloth farmers rain old announced planted news of played with snow barley with built talked barley played across the about markets yesterday harvest winter.</p><p>Each government cloth announced while old mill planted wars planted while announced arrived river about letters cloth wars markets ships summer cloth about harvest where winter slowly rain planted slowly evening announced about wars winter mill barley slowly slowly where children planted.</p><p>Were their snow each wheat old slowly evening brought yesterday beyond parents bridges rain hills played barley news planted cloth harvest from about announced arrived talked their cloth were their beyond where ran slowly were news from spices of slowly valleys mill ships past wheat.</p></div></section>
<section style="font-family:serif;font-size:10pt"><h2>Local news</h2><div style="column-count:3;width:540px;text-align:justify;column-gap:14px"><p>From markets barley where committee the weather stories wheat kings news prices summer autumn from summer planted rain each where were beyond from government summer autumn merchants storms towns spring with spices rain brought snow built parents across evening storms weather announced snow.</p><p>Snow about past government and in storms committee children from in valleys planted built rain evening yesterday spring yesterday in farmers government beyond announced of rain letters kings planted bridges bridges spices in storms letters towns markets prices announced rain barley.</p><p>Slowly winter children kings hills from prices across about kings played snow children wars cloth autumn barley beyond brought from announced autumn markets wheat the wars while their past storms brought evening and planted valleys merchants arrived wheat markets letters autumn.</p><p>Were ran old slowly bridges old from kings yesterday weather harvest were talked roads government distant hills weather mill brought barley children snow where spices barley summer merchants brought storms.</p><p>Mill talked planted ran prices rain their talked planted children brought distant evening slowly stories their government storms spring harvest harvest brought children spices spring built committee spring stories of rain cloth planted the ships.</p><p>Announced slowly each planted their kings slowly past roads farmers arrived autumn brought autumn and planted talked were wheat from from distant the while and roads towns news roads oats.</p><p>Parents talked about snow mill oats harvest old wheat were in their snow announced ran built ran wars mill from farmers committee mill built spring and planted spring.</p><p>Arrived arrived cloth roads cloth evening cloth the river bridges in played weather government arrived valleys evening evening kings talked old slowly river distant their towns rain ships built ran valleys hills mill of arrived yesterday farmers planted farmers spring arrived letters old barley towns past.</p><p>Storms oats from storms ships across kings built with wheat autumn of kings roads where winter valleys evening slowly the where government stories of yesterday kings roads old built river cloth spices merchants letters about harvest wars with and where spices letters evening were were slowly government yesterday.</p><p>About slowly their towns from prices news river snow news their announced cloth about markets hills beyond with spring hills stories and brought and old spring storms the talked old weather oats wheat valleys yesterday merchants towns cloth brought cloth oats about.</p></div></section>
</body></html>`;

const browser = await puppeteer.launch({ headless: true, ...(process.env.CHROME ? { executablePath: process.env.CHROME } : {}) });
const tab = await browser.newPage();
for (const [name, html, tagged] of [['tables_borderless_chrome', chromeBorderless, false], ['tables_tagged_chrome', chromeTagged, true], ['two_columns_not_tables_chrome', notTables, false], ['prose_not_tables_chrome', prose, false]]) {
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

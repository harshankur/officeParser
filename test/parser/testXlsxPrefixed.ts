/** Regression test for XLSX files whose SpreadsheetML elements carry a namespace prefix. */
import { strToU8, zipSync } from 'fflate';
import { OfficeParser } from '../../src/OfficeParser';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

async function main() {
    const file = Buffer.from(zipSync({
        '[Content_Types].xml': strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
        '_rels/.rels': strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
        'xl/workbook.xml': strToU8(`<x:workbook xmlns:x="${NS}" xmlns:r="${REL}"><x:sheets><x:sheet name="Data" sheetId="1" r:id="rId1"/></x:sheets></x:workbook>`),
        'xl/_rels/workbook.xml.rels': strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
        'xl/worksheets/sheet1.xml': strToU8(`<x:worksheet xmlns:x="${NS}"><x:sheetData><x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>Hello</x:t></x:is></x:c><x:c r="B1"><x:v>42</x:v></x:c></x:row></x:sheetData></x:worksheet>`),
    }));

    const ast = await OfficeParser.parseOffice(file);
    const out = JSON.stringify(ast.content);
    if (!out.includes('"sheetName":"Data"') || !out.includes('Hello') || !out.includes('42')) {
        throw new Error(`prefixed sheet not read: ${out}`);
    }
    console.log('ok');
}

main().catch(err => { console.error(err); process.exit(1); });

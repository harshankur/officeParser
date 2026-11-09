/** Regression tests for XLSX cell text that the fixture baselines don't cover. */
import { strToU8, zipSync } from 'fflate';
import { OfficeParser } from '../../src/OfficeParser';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function xlsx(workbook: string, sheet: string): Buffer {
    return Buffer.from(zipSync({
        '[Content_Types].xml': strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
        '_rels/.rels': strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
        'xl/workbook.xml': strToU8(workbook),
        'xl/_rels/workbook.xml.rels': strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
        'xl/worksheets/sheet1.xml': strToU8(sheet),
    }));
}

async function main() {
    const workbook = `<workbook xmlns="${NS}" xmlns:r="${REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`;
    const sheet = `<worksheet xmlns="${NS}"><sheetData><row r="1"><c r="A1" t="str"><f>"AT"&amp;"&amp;T"</f><v>AT&amp;T</v></c></row></sheetData></worksheet>`;
    const ast = await OfficeParser.parseOffice(xlsx(workbook, sheet));
    const out = JSON.stringify(ast.content);
    if (!out.includes('AT&T') || out.includes('&amp;')) {
        throw new Error(`formula text not decoded: ${out}`);
    }
    console.log('ok');
}

main().catch(err => { console.error(err); process.exit(1); });

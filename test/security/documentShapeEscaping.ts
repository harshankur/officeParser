/**
 * Escaping of the document-supplied text the newer shapes write: a note's mark, a picture's caption, a
 * disclosure's summary, keyboard input and a sheet's name as a heading. Each lands in HTML, Markdown or
 * XML, and none may open markup there.
 */
import { unzipSync } from 'fflate';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParser } from '../../src/OfficeParser';
import type { OfficeParserAST } from '../../src/types';
import { parseXmlString } from '../../src/utils/xmlUtils';

type Check = (name: string, cond: boolean, detail?: string) => void;

const HOSTILE = '</sup></a><script>alert(1)</script><img src=x onerror=alert(1)>]]>&"\'\n\n# h';

const astOf = (content: any[]): OfficeParserAST => ({ type: 'docx', metadata: {}, attachments: [], content } as any);

/** No live tag the hostile text could open: a <script>, or an element with an event handler. */
const live = (html: string): boolean => /<script\b/i.test(html) || /<[a-z][^>]*\son[a-z]+\s*=/i.test(html);

export async function documentShapeEscapingTests(check: Check): Promise<void> {
    console.log('- Note marks, captions, disclosures, keyboard input and sheet headings escaped...');
    const silent = { onWarning: () => {} };
    // A note's mark: HTML's reference, Markdown's HTML table cell, DOCX's and ODT's custom mark.
    const note = { type: 'note', metadata: { noteType: 'footnote', mark: HOSTILE.slice(0, 60), customMark: true }, children: [{ type: 'paragraph', children: [{ type: 'text', text: 'n' }] }] };
    const noted = astOf([{ type: 'paragraph', children: [{ type: 'text', text: 'a', notes: [note] }] },
        { type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [{ type: 'text', text: 'c', notes: [note] }] }, { type: 'cell', metadata: { colSpan: 2 }, children: [{ type: 'text', text: 'd' }] }] }] }]);
    const noteHtml = String((await OfficeGenerator.generate(noted, 'html', { htmlConfig: { standalone: false }, ...silent } as any)).value);
    check('note mark: HTML reference escaped', !live(noteHtml), noteHtml);
    const noteMd = String((await OfficeGenerator.generate(noted, 'md', silent as any)).value);
    check('note mark: Markdown HTML cell escaped', !live(noteMd), noteMd);
    for (const format of ['docx', 'odt'] as const) {
        const files = unzipSync((await OfficeGenerator.generate(noted, format, silent as any)).value as Uint8Array);
        const part = new TextDecoder().decode(files[format === 'docx' ? 'word/document.xml' : 'content.xml']);
        let wellFormed = true;
        try { parseXmlString(part); } catch { wellFormed = false; }
        check(`note mark: ${format} custom mark well-formed`, wellFormed && !part.includes('<script'));
    }
    // A picture's caption: HTML's <figcaption>, Markdown's <figure> block.
    const pictured = astOf([{ type: 'image', metadata: { url: 'https://example.com/a.png', altText: 'a', caption: HOSTILE } }]);
    const figureHtml = String((await OfficeGenerator.generate(pictured, 'html', { htmlConfig: { standalone: false }, ...silent } as any)).value);
    check('caption: HTML figcaption escaped', !live(figureHtml), figureHtml);
    const figureMd = String((await OfficeGenerator.generate(pictured, 'md', silent as any)).value);
    check('caption: Markdown figure block escaped, and on lines of its own', !live(figureMd) && !/\n\n# h/.test(figureMd), figureMd);
    const figureBack = await OfficeParser.parseOffice(Buffer.from(figureMd), { fileType: 'md' });
    check('caption: Markdown figure reads back as the same text', (figureBack.content[0]?.metadata as any)?.caption === HOSTILE.replace(/\s+/g, ' ').trim());
    // A disclosure's summary: HTML's and Markdown's <summary>.
    const disclosed = astOf([{ type: 'disclosure', metadata: { summary: HOSTILE, open: true }, children: [{ type: 'paragraph', children: [{ type: 'text', text: 'body' }] }] }]);
    // An open state that is not a boolean (an AST built in code) writes nothing of itself.
    const badOpen = String((await OfficeGenerator.generate(astOf([{ type: 'disclosure', metadata: { open: '"><script>' }, children: [] }]), 'html', { htmlConfig: { standalone: false }, ...silent } as any)).value);
    check('disclosure: a non-boolean open state writes no markup', !live(badOpen), badOpen);
    const detailsHtml = String((await OfficeGenerator.generate(disclosed, 'html', { htmlConfig: { standalone: false }, ...silent } as any)).value);
    check('summary: HTML escaped', !live(detailsHtml), detailsHtml);
    const detailsMd = String((await OfficeGenerator.generate(disclosed, 'md', silent as any)).value);
    check('summary: Markdown escaped, the block not ended early', !live(detailsMd) && /^<details open>\n<summary>[^\n]*<\/summary>\n\nbody\n\n<\/details>/.test(detailsMd), detailsMd);
    const detailsBack = await OfficeParser.parseOffice(Buffer.from(detailsMd), { fileType: 'md' });
    check('summary: Markdown reads back as the same text', (detailsBack.content[0]?.metadata as any)?.summary === HOSTILE.replace(/\s+/g, ' ').trim());
    // Keyboard input: HTML's and Markdown's <kbd>.
    const keyed = astOf([{ type: 'paragraph', children: [{ type: 'text', text: HOSTILE, formatting: { keyboard: true, font: 'monospace' } }] }]);
    const kbdHtml = String((await OfficeGenerator.generate(keyed, 'html', { htmlConfig: { standalone: false }, ...silent } as any)).value);
    check('keyboard: HTML escaped', !live(kbdHtml), kbdHtml);
    const kbdMd = String((await OfficeGenerator.generate(keyed, 'md', silent as any)).value);
    check('keyboard: Markdown escaped, on one line', !live(kbdMd) && !kbdMd.includes('\n# h'), kbdMd);
    // A sheet's name as a heading, with sheets as tables.
    const sheets = { type: 'xlsx', metadata: {}, attachments: [], content: [{ type: 'sheet', metadata: { sheetName: HOSTILE }, children: [{ type: 'row', children: [{ type: 'cell', metadata: { row: 0, col: 0 }, children: [{ type: 'text', text: 'v' }] }] }] }] } as any;
    const sheetHtml = String((await OfficeGenerator.generate(sheets, 'html', { htmlConfig: { standalone: false, sheetLayout: 'tables' }, ...silent } as any)).value);
    check('sheet name: heading escaped with sheets as tables', !live(sheetHtml), sheetHtml);
    const kbdBack = await OfficeParser.parseOffice(Buffer.from(kbdMd), { fileType: 'md' });
    const kbdRun = kbdBack.content[0]?.children?.find(c => c.formatting?.keyboard);
    check('keyboard: Markdown reads back as keyboard input', !!kbdRun && kbdRun.text === HOSTILE.replace(/[\r\n]+/g, ' '), JSON.stringify(kbdRun?.text));
}

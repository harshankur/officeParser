/**
 * Spaces next to tags in an ODT template: a paragraph whose runs of spaces are `<text:s text:c="N"/>`
 * (read short past a limit) must render the same text as the same paragraph with those spaces typed out
 * (read whole, the ground truth), once `<text:s/>` is expanded in both outputs. Random delimiters whose
 * edges hold 0 to 12 spaces (and starts of spaces only), gaps of spaces between tags holding the edges of
 * the tags beside them and up to 300 more, as one run or as typed spaces and runs mixed, tags sometimes
 * split across a span, filled, filled empty, or kept.
 */
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"';
const MEDIA = 'application/vnd.oasis.opendocument.text';
const odt = (body: string): Uint8Array => zipSync({
    mimetype: [strToU8(MEDIA), { level: 0 }],
    'META-INF/manifest.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="${MEDIA}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>`),
    'content.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.3"><office:body><office:text>${body}</office:text></office:body></office:document-content>`),
});
const bodyOf = (bytes: Uint8Array): string => {
    const content = strFromU8(unzipSync(bytes)['content.xml']);
    return content.slice(content.indexOf('<office:text>') + 13, content.indexOf('</office:text>'));
};
const expanded = (xml: string): string => xml.replace(/<text:s text:c="(\d+)"\/>/g, (_, n) => ' '.repeat(Number(n))).replace(/<text:s\/>/g, ' ');

/** Renders an ODT template's bytes with a config (`OfficeTemplate.render`). */
export type OdtRender = (bytes: Uint8Array, config: any) => Promise<Uint8Array>;

/** Runs `cases` random paragraphs; returns the failures (at most 8), described. */
export async function checkOdtSpaceRuns(render: OdtRender, cases: number, seed = 1): Promise<string[]> {
    // The C library generator, modulo 2^31 in integer arithmetic, which has its full period: multiplied
    // in floating point the product passed 2^53, lost its low bits, and every seed fell into one cycle
    // of 10,466 values. Divided by 2^31, a draw is under 1, so a pick stays inside its list.
    const rnd = (): number => { seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff; return seed / 0x80000000; };
    const pick = <T>(xs: readonly T[]): T => xs[Math.floor(rnd() * xs.length)];
    const int = (n: number): number => Math.floor(rnd() * n);
    type Piece = { xml: string; typed: string };
    const run = (n: number): Piece => ({ xml: n === 1 ? '<text:s/>' : `<text:s text:c="${n}"/>`, typed: ' '.repeat(n) });
    const literal = (t: string): Piece => ({ xml: t, typed: t });
    const spaces = (k: number): Piece[] => {
        if (k > 0 && rnd() < 0.5) return [run(k)];
        const out: Piece[] = [];
        for (let left = k; left > 0;) { const n = 1 + int(left); out.push(rnd() < 0.3 ? literal(' '.repeat(n)) : run(n)); left -= n; }
        return out;
    };
    const escape = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const failures: string[] = [];
    for (let c = 0; c < cases && failures.length < 8; c++) {
        const edge = (): number => pick([0, 0, 1, 2, 3, 7, 8, 9, 12]);
        const [o, k] = pick([['[[', ']]'], ['{{', '}}'], ['<<', '>>']]);
        const delimiters = rnd() < 0.1 ? { start: ' '.repeat(1 + int(3)), end: '}}' } : { start: ' '.repeat(edge()) + o, end: k + ' '.repeat(edge()) };
        const startCore = escape(delimiters.start.trimStart()), endCore = escape(delimiters.end.trimEnd());
        const lead = delimiters.start.length - delimiters.start.trimStart().length, trail = delimiters.end.length - delimiters.end.trimEnd().length;
        const pieces: Piece[] = [literal(pick(['x', 'x ']))];
        const n = 1 + int(4);
        for (let j = 0; j <= n; j++) {
            const need = (j > 0 ? trail : 0) + (j < n ? lead : 0);
            const gap = rnd() < 0.85 ? need + pick([0, 0, 1, 2, 5, 16, 40, 300]) : int(need + 1);
            if (rnd() < 0.2) pieces.push(literal('y'));
            pieces.push(...spaces(gap));
            if (j === n) break;
            const core = `${startCore} ${pick(['a', 'b'])} ${endCore}`;
            if (rnd() < 0.3) { const cut = 1 + int(core.length - 1); pieces.push(literal(core.slice(0, cut)), literal(`<text:span>${core.slice(cut)}</text:span>`)); }
            else pieces.push(literal(core));
        }
        pieces.push(literal('z'));
        const fill = pick([{ data: { a: 'A', b: 'B' } }, { data: { a: '', b: 'B' } }, { data: {}, onMissing: 'keep' }]);
        const rendered = async (body: string) => {
            try { return expanded(bodyOf(await render(odt(`<text:p>${body}</text:p>`), { delimiters, ...fill, onWarning: () => { } }))); }
            catch (e: any) { return `error ${e?.officeIssue?.code ?? e}`; }
        };
        const template = pieces.map(x => x.xml).join('');
        const got = await rendered(template), want = await rendered(pieces.map(x => x.typed).join(''));
        if (got !== want) failures.push(`${JSON.stringify(delimiters)} ${JSON.stringify(fill)} ${template}: want ${JSON.stringify(want)}, got ${JSON.stringify(got)}`);
    }
    return failures;
}

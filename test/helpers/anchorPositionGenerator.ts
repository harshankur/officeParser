/**
 * A generator of HTML documents with named anchors and element ids in every position, used by the anchor
 * round trip tests.
 *
 * Each document holds one to three blocks of 16 kinds (paragraph, heading, quote, bulleted and numbered
 * items, a table of a data and a header cell, a definition list, a disclosure, a disclosure holding stray
 * summaries, a figure of a picture and a caption of a line or of blocks, a div, an admonition, preformatted
 * text, a rule, a picture), nested three deep. Anchors come in five forms (`<a id name>`, `<a name>`,
 * `<a id>`, an empty `<span id>`, `<a id name>` holding a space) before a block, as its first and last
 * content, between its parts, after it, or as its only content, and a quarter of the elements carry an id of
 * their own. A disclosure's summary is none, a line, blocks, anchors only or empty, with anchors before it
 * and a body of anchors alone or of blocks.
 *
 * Each document is read, then written and read back three times as HTML, as EPUB and as Markdown. Every id
 * the first read holds must still be held after pass 3, and every word kept. No id may be read onto a
 * disclosure's summary, which takes none.
 *
 * One shape is counted apart (`excludedLost`), as it was lost the same way by 8.6.0 (next-release ledger):
 * in Markdown, the ids of cells and headings in a table that holds a nested table, inside a list item, which
 * Markdown writes as HTML.
 *
 * The random numbers come from a full-period LCG mod 2^32 (Math.imul, a = 1664525, c = 1013904223: c odd,
 * a - 1 divisible by 4), so a seed never repeats within 2^32 draws.
 */
import { strFromU8, unzipSync } from 'fflate';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParser } from '../../src/OfficeParser';

type Format = 'html' | 'epub' | 'md';

/** A reader and writer to run the documents through (this tree's, or another's to compare with). */
export interface AnchorRoundTripIo {
    parse: (value: string | Uint8Array, format: Format) => Promise<any>;
    generate: (ast: any, format: Format, config: Record<string, unknown>) => Promise<string | Uint8Array>;
}

/** This tree's reader and writer. */
export const anchorRoundTripIo: AnchorRoundTripIo = {
    parse: (value, format) => OfficeParser.parseOffice(typeof value === 'string' ? Buffer.from(value) : Buffer.from(value), { fileType: format }),
    generate: async (ast, format, config) => (await OfficeGenerator.generate(ast, format, config as any)).value as string | Uint8Array,
};

/** The counts and failures of one run. */
export interface AnchorPositionResult {
    docs: number;
    distinctSources: number;
    /** Distinct pass-1 outputs by format. */
    distinctOutputs: Record<Format, number>;
    /** Ids the first read holds, by the type of the node holding them. */
    holders: Record<string, number>;
    /** Documents losing an id by pass 3, by format. */
    lostDocs: Record<Format, number>;
    /** Ids lost in the shape counted apart (see above), Markdown only. */
    excludedLost: number;
    failures: Record<string, number>;
    examples: Record<string, string[]>;
}

const FORMATS: Format[] = ['html', 'epub', 'md'];

/** An EPUB's XHTML and HTML files, in name order, its generated UUID masked. */
function epubXhtml(value: Uint8Array): string {
    const files = unzipSync(value);
    return Object.keys(files).sort().filter(name => /\.x?html$/.test(name)).map(name => `${name}\n${strFromU8(files[name]).replace(/urn:uuid:[0-9a-f-]+/g, 'UUID')}`).join('\n---\n');
}

/** Every id in a tree, with the node holding it and the types of that node's ancestors (outermost first). */
function idsOf(content: any): Map<string, { holder: string; ancestors: string[]; inTableHoldingTable: boolean }> {
    const out = new Map<string, { holder: string; ancestors: string[]; inTableHoldingTable: boolean }>();
    const holdsTable = (node: any): boolean => (node.children ?? []).some((child: any) => child?.type === 'table' || holdsTable(child));
    const walk = (node: any, ancestors: string[], inTableHoldingTable: boolean): void => {
        if (!node || typeof node !== 'object' || ancestors.length > 200) return;
        if (Array.isArray(node)) { node.forEach(child => walk(child, ancestors, inTableHoldingTable)); return; }
        for (const id of node.metadata?.anchorIds ?? []) if (!out.has(id)) out.set(id, { holder: node.type, ancestors, inTableHoldingTable });
        const inner = inTableHoldingTable || (node.type === 'table' && ancestors.includes('list') && holdsTable(node));
        walk(node.children, [...ancestors, node.type], inner);
        walk(node.notes, [...ancestors, node.type], inner);
        walk(node.comments, [...ancestors, node.type], inner);
    };
    walk(content, [], false);
    return out;
}

/** Runs `count` documents from `seed` through `io`. Failures: an id lost, an id on a summary, a word lost, throws. */
export async function runAnchorPositionGenerator(count: number, seed: number, io: AnchorRoundTripIo = anchorRoundTripIo): Promise<AnchorPositionResult> {
    let state = seed >>> 0;
    const rnd = (): number => (state = (Math.imul(state, 1664525) + 1013904223) >>> 0) / 4294967296;
    const pick = <T>(xs: readonly T[]): T => xs[Math.floor(rnd() * xs.length)];
    const result: AnchorPositionResult = {
        docs: 0, distinctSources: 0, distinctOutputs: { html: 0, epub: 0, md: 0 }, holders: {}, lostDocs: { html: 0, epub: 0, md: 0 }, excludedLost: 0, failures: {}, examples: {},
    };
    const fail = (kind: string, message: string): void => {
        result.failures[kind] = (result.failures[kind] ?? 0) + 1;
        const list = (result.examples[kind] ??= []);
        if (list.length < 3) list.push(message.slice(0, 1600));
    };
    const sources = new Set<string>();
    const outputs: Record<Format, Set<string>> = { html: new Set(), epub: new Set(), md: new Set() };

    let idNo = 0, wordNo = 0;
    const newId = (): string => `x${++idNo}`;
    const anchor = (): string => {
        const id = newId();
        return pick([`<a id="${id}" name="${id}"></a>`, `<a name="${id}"></a>`, `<a id="${id}"></a>`, `<span id="${id}"></span>`, `<a id="${id}" name="${id}"> </a>`]);
    };
    const anchors = (): string => (rnd() < 0.5 ? '' : rnd() < 0.7 ? anchor() : anchor() + (rnd() < 0.5 ? ' ' : '') + anchor());
    const own = (): string => (rnd() < 0.25 ? ` id="${newId()}"` : '');
    const word = (): string => `w${++wordNo}`;
    const line = (): string => {
        const n = Math.floor(rnd() * 3);
        let s = '';
        for (let i = 0; i < n; i++) s += pick([word() + ' ', `<b>${word()}</b> `, anchor(), `<a href="#x1">${word()}</a> `, '<br>']);
        return s;
    };
    // A container's content: anchors (or none) at its start, at its end and between its parts; or nothing but anchors.
    const around = (parts: string[]): string => (rnd() < 0.12 ? anchors() || anchor() : anchors() + parts.map(part => part + anchors()).join(''));
    const KINDS = ['p', 'h', 'quote', 'ul', 'ol', 'table', 'dl', 'details', 'strayDetails', 'figure', 'figureBlocks', 'div', 'admonition', 'pre', 'hr', 'img'] as const;
    const block = (depth: number): string => {
        const kind = depth > 2 ? pick(['p', 'h', 'img', 'hr', 'pre'] as const) : pick(KINDS);
        const kids = (): string[] => Array.from({ length: Math.floor(rnd() * 3) }, () => (rnd() < 0.4 ? line() : block(depth + 1)));
        switch (kind) {
            case 'p': return `<p${own()}>${around([line()])}</p>`;
            case 'h': { const level = 1 + Math.floor(rnd() * 3); return `<h${level}${own()}>${around([line()])}</h${level}>`; }
            case 'quote': return `<blockquote${own()}>${around(kids())}</blockquote>`;
            case 'ul': case 'ol': return `<${kind}${own()}>${Array.from({ length: 1 + Math.floor(rnd() * 2) }, () => `${anchors()}<li${own()}>${around(rnd() < 0.5 ? [line()] : kids())}</li>`).join('')}</${kind}>`;
            case 'table': return `<table${own()}><tr${own()}><td${own()}>${around(rnd() < 0.5 ? [line()] : kids())}</td><th${own()}>${around([line()])}</th></tr></table>`;
            case 'dl': return `<dl${own()}><dt${own()}>${around([line()])}</dt><dd${own()}>${around([line()])}</dd></dl>`;
            case 'details': {
                const summary = pick(['none', 'line', 'blocks', 'anchorsOnly', 'empty'] as const);
                const s = summary === 'none' ? '' : summary === 'line' ? `<summary${own()}>${around([line() || word()])}</summary>` : summary === 'blocks' ? `<summary>${around(kids())}</summary>`
                    : summary === 'anchorsOnly' ? `<summary>${anchor()}${anchors()}</summary>` : '<summary></summary>';
                const before = rnd() < 0.3 ? anchors() : '';
                const body = rnd() < 0.4 ? anchors() : around(kids());
                return `<details${own()}${rnd() < 0.3 ? ' open' : ''}>${before}${s}${body}</details>`;
            }
            case 'strayDetails': return `<details><summary>${word()}</summary>${anchors()}<summary>${around(kids())}</summary>${anchors()}</details>`;
            case 'figure': return `<figure${own()}>${anchors()}<img src="a.png" alt="${word()}"${own()}>${anchors()}<figcaption>${around([word() + ' ' + word()])}</figcaption>${anchors()}</figure>`;
            case 'figureBlocks': return `<figure><img src="a.png" alt="${word()}"><figcaption>${around(kids())}</figcaption></figure>`;
            case 'div': return `<div${own()}>${around(kids())}</div>`;
            case 'admonition': return `<div class="admonition admonition-note" data-type="note"${own()}>${around(kids())}</div>`;
            case 'pre': return `<pre${own()}>${word()}</pre>`;
            case 'hr': return `<hr${own()}>`;
            case 'img': return `<img src="a.png" alt="${word()}"${own()}>`;
        }
    };
    const wordsOf = async (ast: any): Promise<string> => (String((await ast.to('text')).value).match(/w\d+/g) ?? []).sort().join(',');

    for (let i = 0; i < count; i++) {
        idNo = 0;
        wordNo = 0;
        const html = `<html><body><p><a href="#x1">go</a></p>${Array.from({ length: 1 + Math.floor(rnd() * 3) }, () => anchors() + block(0)).join('')}${anchors()}</body></html>`;
        sources.add(html);
        result.docs++;
        let first: any;
        try { first = await io.parse(html, 'html'); } catch (error) { fail('read:throw', `${(error as Error).message}\n${html}`); continue; }
        const ids = idsOf(first.content);
        for (const [id, { holder }] of ids) {
            result.holders[holder] = (result.holders[holder] ?? 0) + 1;
            if (holder === 'summary') fail('read:idOnSummary', `${id}\n${html}`);
        }
        const words = await wordsOf(first);
        for (const format of FORMATS) {
            let current = first;
            const outs: string[] = [];
            try {
                for (let pass = 0; pass < 3; pass++) {
                    const value = await io.generate(current, format, format === 'html' ? { htmlConfig: { standalone: false } } : {});
                    outs.push(format === 'epub' ? epubXhtml(value as Uint8Array) : String(value));
                    current = await io.parse(value, format);
                }
            } catch (error) { fail(`${format}:throw`, `${(error as Error).message}\n${html}`); continue; }
            outputs[format].add(outs[0]);
            const after = idsOf(current.content);
            const lost = [...ids.keys()].filter(id => !after.has(id));
            const excluded = format === 'md' ? lost.filter(id => { const at = ids.get(id)!; return at.inTableHoldingTable && (at.holder === 'cell' || at.holder === 'heading'); }) : [];
            result.excludedLost += excluded.length;
            const counted = lost.filter(id => !excluded.includes(id));
            if (counted.length) {
                result.lostDocs[format]++;
                fail(`${format}:idLost`, `${counted.map(id => `${id}@${ids.get(id)!.holder}`).join(',')}\n${html}\nout1=${outs[0]}\nout3=${outs[2]}`);
            }
            if ((await wordsOf(current)) !== words) fail(`${format}:wordsChanged`, `${html}\nout1=${outs[0]}`);
        }
    }
    result.distinctSources = sources.size;
    result.distinctOutputs = { html: outputs.html.size, epub: outputs.epub.size, md: outputs.md.size };
    return result;
}

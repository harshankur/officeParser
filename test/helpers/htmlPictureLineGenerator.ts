/**
 * Generators of HTML and EPUB round trips for the pictures and headings written in a line.
 *
 * runHtmlPictureMatrix: captioned pictures (a caption of one or two words; url, small attachment or one over
 * the 1000-byte cap; alt text, link, anchors; recognized text of none, one line or two lines) under every
 * image mode (8 configurations, three with `maxInlineImageBytes: 1000`) in 17 containers (heading, Quote
 * paragraphs, paragraph, a disclosure's summary and a heading in its body, inline and paragraph list items,
 * cells of inline content, a paragraph or a heading, definition terms and descriptions, admonitions of
 * inline content or a heading, a heading whose picture is in a comment's range, a footnote) at 7 positions
 * (alone, at the start, between words, at the end, two pictures, a line break before or after it), the
 * words beside it plain, bold, linked or inline math. Each document is written as HTML and as EPUB and read
 * back four times: no word joined to another, no word lost (after pass 1 and pass 4), and each pass's
 * output recorded for when it settles.
 *
 * runHtmlPictureCommentMatrix: the same documents with a source comment (a comment, or under sourceAttributes
 * an empty `data-html-comment` span) at the start of each picture's line, right before or after the picture
 * or at the line's end, read back five times: no joined or lost words, no line break written for a comment
 * (after a line's start or another break with only hidden markup between), each pass recorded for settling.
 * It takes a reader and writer (HtmlRoundTripIo), so another tree's can be compared.
 *
 * runDisclosureHeadingGenerator: disclosures holding headings (in the body, the summary, a nested
 * disclosure, after loose text, two headings; ids none, the generated slug, another, two), as ASTs written
 * as HTML and EPUB with and without generateIds, and as HTML a person writes (a named anchor before the
 * heading or none), each read and written five times: the output settles at pass 1 (the first output is
 * the fifth), with no empty paragraph.
 *
 * The random numbers come from a full-period LCG mod 2^32 (Math.imul, a = 1664525, c = 1013904223: c odd,
 * a - 1 divisible by 4), so a seed never repeats within 2^32 draws.
 */
import { strFromU8, unzipSync } from 'fflate';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParser } from '../../src/OfficeParser';

type Node = Record<string, any>;

/** The counts and failures of one run. */
export interface HtmlLineGeneratorResult {
    docs: number;
    distinctSources: number;
    distinctHtml: number;
    distinctEpub: number;
    /** Distinct `mode/container/position` combinations covered (picture matrix only). */
    distinctCombos: number;
    /** `format:pass` (0 when the four or five outputs never settle) to the number of round trips. */
    settledAtPass: Record<string, number>;
    /** Round trips settled after pass 3 or never, by `format:container/position` (picture matrix only). */
    unsettledShapes: Record<string, number>;
    failures: Record<string, number>;
    examples: Record<string, string[]>;
}

const SMALL = Buffer.alloc(150, 9).toString('base64');
const BIG = Buffer.alloc(3000, 7).toString('base64');

/** An EPUB's XHTML and HTML files, in name order, its generated UUID masked. */
function epubXhtml(value: Uint8Array): string {
    const files = unzipSync(value);
    return Object.keys(files).sort().filter(name => /\.x?html$/.test(name)).map(name => `${name}\n${strFromU8(files[name]).replace(/urn:uuid:[0-9a-f-]+/g, 'UUID')}`).join('\n---\n');
}

/** A reader and writer to run the documents through (this tree's, or another's to compare with). */
export interface HtmlRoundTripIo {
    parse: (value: string | Uint8Array, format: 'html' | 'epub') => Promise<any>;
    generate: (ast: any, format: 'html' | 'epub', config: Record<string, unknown>) => Promise<string | Uint8Array>;
}

/** This tree's reader and writer. */
export const htmlRoundTripIo: HtmlRoundTripIo = {
    parse: (value, format) => OfficeParser.parseOffice(typeof value === 'string' ? Buffer.from(value) : Buffer.from(value), { fileType: format }),
    generate: async (ast, format, config) => (await OfficeGenerator.generate(ast, format, config as any)).value as string | Uint8Array,
};

/** Writes `ast` (or reads `source` first) and reads it back `passes` times: each pass's output and text. */
async function roundTrip(ast: Node | undefined, format: 'html' | 'epub', config: Record<string, unknown>, passes: number, source?: string, io: HtmlRoundTripIo = htmlRoundTripIo): Promise<{ outs: string[]; texts: string[] }> {
    const outs: string[] = [];
    const texts: string[] = [];
    let current: any = source === undefined ? ast : await io.parse(source, format);
    for (let pass = 0; pass < passes; pass++) {
        const settings = format === 'html' ? { ...config, htmlConfig: { standalone: false, ...(config.htmlConfig as object | undefined) } } : config;
        const value = await io.generate(current, format, settings);
        outs.push(format === 'epub' ? epubXhtml(value as Uint8Array) : String(value));
        current = await io.parse(value, format);
        texts.push(String((await current.to('text')).value));
    }
    return { outs, texts };
}

/** The first pass from which every output is the same (1 when all are), or 0 when the last two differ. */
const settledAt = (outs: string[]): number => {
    for (let pass = 0; pass + 1 < outs.length; pass++) if (outs.slice(pass).every(out => out === outs[pass])) return pass + 1;
    return 0;
};
const withoutData = (out: string): string => out.replace(/data:image\/png;base64,[A-Za-z0-9+/=]+/g, 'DATA');
const tokensOf = (text: string): string[] => text.match(/[A-Za-z0-9]+/g) ?? [];

function newResult(): HtmlLineGeneratorResult {
    return { docs: 0, distinctSources: 0, distinctHtml: 0, distinctEpub: 0, distinctCombos: 0, settledAtPass: {}, unsettledShapes: {}, failures: {}, examples: {} };
}
function failer(result: HtmlLineGeneratorResult): (kind: string, message: string) => void {
    return (kind, message) => {
        result.failures[kind] = (result.failures[kind] ?? 0) + 1;
        const list = (result.examples[kind] ??= []);
        if (list.length < 3) list.push(message.slice(0, 1600));
    };
}

const IMAGE_MODES: Array<Record<string, unknown>> = [
    {}, { includeImages: 'image+ocr-text' }, { includeImages: 'ocr-text-only' }, { includeImages: 'none' },
    { maxInlineImageBytes: 1000 }, { includeImages: 'image+ocr-text', maxInlineImageBytes: 1000 }, { includeImages: 'image-only', maxInlineImageBytes: 1000 }, { includeImages: 'ocr-text-only', maxInlineImageBytes: 1000 },
];
const POSITIONS = ['alone', 'start', 'middle', 'end', 'pair', 'breakBefore', 'breakAfter'] as const;

/** One captioned-picture document of the matrix: its AST, image mode, container, position and the words a reader sees. */
export interface PictureMatrixDoc {
    i: number;
    ast: { type: string; metadata: Node; attachments: Node[]; content: Node[] };
    mode: Record<string, unknown>;
    container: string;
    position: string;
    ocr: string | undefined;
    /** The words a reader sees: text, every caption, alt text where the image is written, recognized text where it is. */
    expected: Set<string>;
}

/** The matrix's `count` documents from `seed` (the draws runHtmlPictureMatrix has always made). */
export function* htmlPictureMatrixDocs(count: number, seed: number): Generator<PictureMatrixDoc> {
    let state = seed >>> 0;
    const rnd = (): number => (state = (Math.imul(state, 1664525) + 1013904223) >>> 0) / 4294967296;
    const pick = <T>(xs: readonly T[]): T => xs[Math.floor(rnd() * xs.length)];

    const text = (value: string, extra: Node = {}): Node => ({ type: 'text', text: value, ...extra });
    const paragraph = (children: Node[], metadata?: Node): Node => ({ type: 'paragraph', children, ...(metadata ? { metadata } : {}) });
    const heading = (children: Node[]): Node => ({ type: 'heading', metadata: { level: 2 }, children });
    const item = (children: Node[]): Node => ({ type: 'list', metadata: { listType: 'unordered', indentation: 0, itemIndex: 0, listId: 'L' }, children });
    const table = (cell: Node[]): Node => ({ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: cell }, { type: 'cell', children: [text('z0')] }] }] });
    const CONTAINERS: Record<string, (line: Node[]) => Node[]> = {
        heading: line => [heading(line)],
        quote: line => [paragraph(line, { style: 'Quote', quoteId: 'q1' })],
        quote2: line => [paragraph(line, { style: 'Quote', quoteId: 'q2' }), paragraph([text('z1')], { style: 'Quote', quoteId: 'q2' })],
        paragraph: line => [paragraph(line)],
        summary: line => [{ type: 'disclosure', metadata: {}, children: [{ type: 'summary', children: [paragraph(line)] }, paragraph([text('z1')])] }],
        discHeading: line => [{ type: 'disclosure', metadata: {}, children: [{ type: 'summary', children: [paragraph([text('z1')])] }, heading(line)] }],
        itemInline: line => [item(line)],
        itemPara: line => [item([paragraph(line)])],
        cellInline: line => [table(line)],
        cellPara: line => [table([paragraph(line)])],
        cellHeading: line => [table([heading(line)])],
        term: line => [{ type: 'definitionList', children: [{ type: 'definitionTerm', children: line }, { type: 'definitionDescription', children: [text('z1')] }] }],
        desc: line => [{ type: 'definitionList', children: [{ type: 'definitionTerm', children: [text('z1')] }, { type: 'definitionDescription', children: line }] }],
        admonition: line => [{ type: 'admonition', metadata: { admonitionType: 'note' }, children: line }],
        admonitionHeading: line => [{ type: 'admonition', metadata: { admonitionType: 'note' }, children: [heading(line)] }],
        commentInHeading: line => [heading(line.map(node => node.type === 'image' ? { type: 'comment', metadata: { author: 'x' }, children: [node] } : node))],
        footnote: line => [paragraph([text('z1'), { type: 'text', text: 'z2', notes: [{ type: 'note', metadata: { noteType: 'footnote', noteId: 'n1' }, children: [paragraph(line)] }] }])],
    };
    const containerNames = Object.keys(CONTAINERS);

    for (let i = 0; i < count; i++) {
        let wordNo = 0;
        const w = (): string => `w${++wordNo}`;
        const attachments: Node[] = [];
        const neighbour = (): Node => {
            const r = rnd();
            const word = w();
            return r < 0.5 ? text(word) : r < 0.7 ? text(word, { formatting: { bold: true } }) : r < 0.85 ? text(word, { metadata: { link: 'https://e.x', linkType: 'external' } }) : { type: 'code', text: word, metadata: { math: 'inline' } };
        };
        const ocr = pick([undefined, 'line', 'lines'] as const);
        const picture = (): Node => {
            const metadata: Node = { altText: rnd() < 0.8 ? w() : '' };
            if (rnd() < 0.35) metadata.url = 'a.png';
            else {
                const name = `p${attachments.length}.png`;
                attachments.push({ type: 'image', name, mimeType: 'image/png', extension: 'png', data: rnd() < 0.6 ? SMALL : BIG });
                metadata.attachmentName = name;
            }
            metadata.caption = rnd() < 0.5 ? w() : `${w()} ${w()}`;
            if (rnd() < 0.15) metadata.anchorIds = ['id' + w()];
            if (rnd() < 0.1) metadata.link = 'https://e.y';
            const node: Node = { type: 'image', metadata };
            if (ocr === 'line') node.text = `${w()} ${w()}`;
            else if (ocr === 'lines') node.text = `${w()}\n${w()}`;
            return node;
        };
        const space = (): Node[] => (rnd() < 0.5 ? [text(' ')] : []);
        const mode = IMAGE_MODES[i % IMAGE_MODES.length];
        const container = pick(containerNames);
        const position = pick(POSITIONS);
        const line: Node[] = [];
        if (position !== 'start' && position !== 'alone' && position !== 'pair') line.push(neighbour(), ...space());
        if (position === 'breakBefore') line.push({ type: 'break', metadata: { breakType: 'carriageReturn' } });
        line.push(picture());
        if (position === 'pair') line.push(...space(), picture());
        if (position === 'breakAfter') line.push({ type: 'break', metadata: { breakType: 'carriageReturn' } });
        if (position !== 'end' && position !== 'alone') line.push(...space(), neighbour());
        const ast = { type: 'md', metadata: {}, attachments, content: [paragraph([text('z9')]), ...CONTAINERS[container](line), paragraph([text('z8')])] };

        const imageMode = (mode.includeImages as string | undefined) ?? 'image-only';
        const expected = new Set<string>();
        const add = (value?: string): void => (value?.match(/w\d+/g) ?? []).forEach(word => expected.add(word));
        const walk = (node: any): void => {
            if (!node || typeof node !== 'object') return;
            if (Array.isArray(node)) { node.forEach(walk); return; }
            if (node.type === 'image') {
                add(node.metadata.caption);
                if (imageMode === 'image-only' || imageMode === 'image+ocr-text') add(node.metadata.altText);
                if (imageMode === 'image+ocr-text' || imageMode === 'ocr-text-only') add(node.text);
                return;
            }
            if (typeof node.text === 'string') add(node.text);
            walk(node.children);
            walk(node.notes);
        };
        walk(ast.content);
        yield { i, ast, mode, container, position, ocr, expected };
    }
}

/** Joined and lost words of a pass's text (see runHtmlPictureMatrix). */
function wordFailures(text: string, expected: Set<string>): { joined: string[]; lost: string[] } {
    const tokens = tokensOf(text);
    const present = new Set(tokens);
    const joined = tokens.filter(token => /w\d/.test(token) && !/^w\d+$/.test(token) && !/^idw\d+$/.test(token));
    const lost = [...expected].filter(word => !present.has(word) && !joined.some(j => j.includes(word)));
    return { joined, lost };
}

/** Runs `count` captioned-picture documents from `seed`, each as HTML and EPUB. Failures: joined or lost words, throws. */
export async function runHtmlPictureMatrix(count: number, seed: number): Promise<HtmlLineGeneratorResult> {
    const result = newResult();
    const fail = failer(result);
    const sources = new Set<string>(), htmls = new Set<string>(), epubs = new Set<string>(), combos = new Set<string>();

    for (const { i, ast, mode, container, position, ocr, expected } of htmlPictureMatrixDocs(count, seed)) {
        combos.add(`${i % IMAGE_MODES.length}/${container}/${position}`);
        const json = JSON.stringify(ast);
        sources.add(json);
        result.docs++;

        for (const format of ['html', 'epub'] as const) {
            const tag = `${format} ${JSON.stringify(mode)} ${container} ${position} ocr=${ocr}`;
            let trip: { outs: string[]; texts: string[] };
            try { trip = await roundTrip(JSON.parse(json), format, mode, 4); } catch (error) { fail(`${format}:throw`, `${tag}: ${(error as Error).message}`); continue; }
            (format === 'html' ? htmls : epubs).add(trip.outs[0]);
            for (const pass of [0, 3]) {
                const { joined, lost } = wordFailures(trip.texts[pass], expected);
                if (joined.length) fail(`${format}:joined`, `${tag} pass ${pass + 1} joined=${joined.join(',')}\nout1=${JSON.stringify(withoutData(trip.outs[0]))}\ntext=${JSON.stringify(trip.texts[pass])}`);
                if (lost.length) fail(`${format}:lost`, `${tag} pass ${pass + 1} lost=${lost.join(',')}\nout1=${JSON.stringify(withoutData(trip.outs[0]))}\ntext=${JSON.stringify(trip.texts[pass])}`);
            }
            const at = settledAt(trip.outs.map(withoutData));
            result.settledAtPass[`${format}:${at}`] = (result.settledAtPass[`${format}:${at}`] ?? 0) + 1;
            if (!at || at > 3) {
                const shape = `${format}:${container}/${position}`;
                result.unsettledShapes[shape] = (result.unsettledShapes[shape] ?? 0) + 1;
            }
        }
    }
    result.distinctSources = sources.size;
    result.distinctHtml = htmls.size;
    result.distinctEpub = epubs.size;
    result.distinctCombos = combos.size;
    return result;
}

/**
 * A line break written for a captioned picture (one followed by the picture, not by the white space a line
 * break node writes after its own) right after a line's start or another line break, with hidden markup
 * between, a source comment among it: the comment was counted as line content.
 */
const BREAK_AFTER_HIDDEN = /(?:<(?:h\d|summary|blockquote|p|li|td|th|dt|dd|div)\b[^>]*>|<br>)(?:\s|<!--(?:[^-]|-(?!->))*-->|<span data-html-comment="[^"]*"><\/span>|<a id="[^"]*" name="[^"]*"><\/a>)*(?:<!--(?:[^-]|-(?!->))*-->|<span data-html-comment="[^"]*"><\/span>)(?:\s|<!--(?:[^-]|-(?!->))*-->|<span data-html-comment="[^"]*"><\/span>|<a id="[^"]*" name="[^"]*"><\/a>)*<br>(?!\s)/;

/**
 * Runs the picture matrix's `count` documents from `seed` with a source comment (`<!--src-->`, or under
 * `sourceAttributes` in HTML an empty `data-html-comment` span) put into each line holding a picture: at its
 * start, right before or after the picture (inside a comment's range, when the picture is in one), or at its
 * end. Each is written as HTML and EPUB and read back five times. Failures: joined or lost words (pass 1 and
 * 5), a line break after a line's start or another line break with only hidden markup between, throws.
 * Round trips settling after pass 3 are counted by shape in `unsettledShapes`.
 */
export async function runHtmlPictureCommentMatrix(count: number, seed: number, io: HtmlRoundTripIo = htmlRoundTripIo): Promise<HtmlLineGeneratorResult> {
    let state = (seed ^ 0x5bd1e995) >>> 0;
    const rnd = (): number => (state = (Math.imul(state, 1664525) + 1013904223) >>> 0) / 4294967296;
    const result = newResult();
    const fail = failer(result);
    const sources = new Set<string>(), htmls = new Set<string>(), epubs = new Set<string>(), combos = new Set<string>();
    const comment = (): Node => ({ type: 'comment', text: 'src', metadata: { sourceSyntax: 'html' } });
    const withComments = (node: any): any => {
        if (!node || typeof node !== 'object') return node;
        if (Array.isArray(node)) return node.map(withComments);
        const out = { ...node };
        if (node.children) {
            const kids = node.children.map(withComments);
            const at = kids.findIndex((child: Node) => child.type === 'image');
            if (at >= 0) {
                const r = rnd();
                kids.splice(r < 0.3 ? 0 : r < 0.55 ? at : r < 0.8 ? at + 1 : kids.length, 0, comment());
            }
            out.children = kids;
        }
        if (node.notes) out.notes = node.notes.map(withComments);
        return out;
    };

    for (const { i, ast, mode, container, position, ocr, expected } of htmlPictureMatrixDocs(count, seed)) {
        const doc = { ...ast, content: withComments(ast.content) };
        const sourceAttributes = rnd() < 0.5;
        const json = JSON.stringify(doc);
        sources.add(json);
        combos.add(`${i % IMAGE_MODES.length}/${container}/${position}`);
        result.docs++;
        for (const format of ['html', 'epub'] as const) {
            const config = format === 'html' && sourceAttributes ? { ...mode, htmlConfig: { sourceAttributes: true } } : mode;
            const tag = `${format} ${JSON.stringify(config)} ${container} ${position} ocr=${ocr}`;
            let trip: { outs: string[]; texts: string[] };
            try { trip = await roundTrip(JSON.parse(json), format, config, 5, undefined, io); } catch (error) { fail(`${format}:throw`, `${tag}: ${(error as Error).message}`); continue; }
            (format === 'html' ? htmls : epubs).add(trip.outs[0]);
            for (const pass of [0, 4]) {
                const { joined, lost } = wordFailures(trip.texts[pass], expected);
                if (joined.length) fail(`${format}:joined`, `${tag} pass ${pass + 1} joined=${joined.join(',')}\nout1=${JSON.stringify(withoutData(trip.outs[0]))}`);
                if (lost.length) fail(`${format}:lost`, `${tag} pass ${pass + 1} lost=${lost.join(',')}\nout1=${JSON.stringify(withoutData(trip.outs[0]))}`);
            }
            const broken = trip.outs.findIndex(out => BREAK_AFTER_HIDDEN.test(out));
            if (broken >= 0) fail(`${format}:breakAfterComment`, `${tag} pass ${broken + 1}: ${JSON.stringify(withoutData(trip.outs[broken]).match(BREAK_AFTER_HIDDEN)?.[0])}\nout=${JSON.stringify(withoutData(trip.outs[broken]))}`);
            const at = settledAt(trip.outs.map(withoutData));
            result.settledAtPass[`${format}:${at}`] = (result.settledAtPass[`${format}:${at}`] ?? 0) + 1;
            if (!at || at > 3) {
                const shape = `${format}:${container}/${position}`;
                result.unsettledShapes[shape] = (result.unsettledShapes[shape] ?? 0) + 1;
            }
        }
    }
    result.distinctSources = sources.size;
    result.distinctHtml = htmls.size;
    result.distinctEpub = epubs.size;
    result.distinctCombos = combos.size;
    return result;
}

/**
 * Runs `count` disclosures holding headings from `seed`, each as an AST (HTML and EPUB, generateIds on and
 * off) and as hand-written HTML (generateIds on and off). Failures: unsettled after pass 1, an empty
 * paragraph in any output, output growing, throws.
 */
export async function runDisclosureHeadingGenerator(count: number, seed: number): Promise<HtmlLineGeneratorResult> {
    let state = seed >>> 0;
    const rnd = (): number => (state = (Math.imul(state, 1664525) + 1013904223) >>> 0) / 4294967296;
    const pick = <T>(xs: readonly T[]): T => xs[Math.floor(rnd() * xs.length)];
    const result = newResult();
    const fail = failer(result);
    const sources = new Set<string>(), htmls = new Set<string>(), epubs = new Set<string>(), combos = new Set<string>();
    const text = (value: string): Node => ({ type: 'text', text: value });
    const paragraph = (value: string): Node => ({ type: 'paragraph', children: [text(value)] });
    const words = ['hello', 'hello there', 'Version 2.0', 'a-b c', 'x'];
    let wordNo = 0;

    for (let i = 0; i < count; i++) {
        const heading = pick(words) + (rnd() < 0.5 ? ` w${++wordNo}` : '');
        const idKind = pick(['none', 'slug', 'other', 'two'] as const);
        const slug = heading.toLowerCase().replace(/[^a-z0-9 -]/g, '').replace(/ /g, '-');
        const ids = idKind === 'none' ? undefined : idKind === 'slug' ? [slug] : idKind === 'other' ? [`x${i}`] : [`x${i}`, `y${i}`];
        const level = 1 + Math.floor(rnd() * 6);
        const where = pick(['body', 'summary', 'nested', 'afterText', 'twoHeadings'] as const);
        const headingNode = (value = heading, anchorIds = ids): Node => ({ type: 'heading', metadata: { level, ...(anchorIds ? { anchorIds } : {}) }, children: [text(value)] });
        const disclosure = (body: Node[], summary: Node[] = [paragraph(`s${i}`)]): Node => ({ type: 'disclosure', metadata: rnd() < 0.3 ? { open: true } : {}, children: [{ type: 'summary', children: summary }, ...body] });
        const content = where === 'body' ? [disclosure([headingNode(), paragraph(`b${i}`)])]
            : where === 'summary' ? [disclosure([paragraph(`b${i}`)], [headingNode()])]
                : where === 'nested' ? [disclosure([disclosure([headingNode()])])]
                    : where === 'afterText' ? [disclosure([paragraph(`b${i}`), headingNode()])]
                        : [disclosure([headingNode(), headingNode(`${heading} two`, ids?.map(id => `${id}b`))])];
        const ast = { type: 'md', metadata: {}, attachments: [], content };
        // The same as HTML a person writes: an id on the heading, a named anchor before it or none.
        const anchor = rnd() < 0.4 ? `<a id="a${i}" name="a${i}"></a>` : '';
        const h = `${anchor}<h${level}${ids ? ` id="${ids[0]}"` : ''}>${heading}</h${level}>`;
        const html = where === 'summary' ? `<details><summary>${h}</summary><p>b${i}</p></details>`
            : where === 'nested' ? `<details><summary>s</summary><details><summary>t</summary>${h}</details></details>`
                : `<details><summary>s</summary>${where === 'afterText' ? 'loose text ' : ''}${h}<p>b${i}</p></details>`;
        const cases: Array<[string, Node | undefined, string | undefined]> = [[`ast:${where}/${idKind}`, ast, undefined], [`html:${where}/${idKind}/${anchor ? 'anchor' : 'noanchor'}`, undefined, html]];
        for (const [name, source, markup] of cases) {
            sources.add(markup ?? JSON.stringify(source));
            result.docs++;
            combos.add(name);
            for (const format of ['html', 'epub'] as const) {
                if (markup !== undefined && format === 'epub') continue;
                for (const config of [{}, { generateIds: false }]) {
                    const tag = `${name} ${format} ${JSON.stringify(config)}`;
                    let trip: { outs: string[]; texts: string[] };
                    try { trip = await roundTrip(source && JSON.parse(JSON.stringify(source)), format, config, 5, markup); } catch (error) { fail(`${format}:throw`, `${tag}: ${(error as Error).message}`); continue; }
                    (format === 'html' ? htmls : epubs).add(trip.outs[0]);
                    const at = settledAt(trip.outs);
                    result.settledAtPass[`${format}:${at}`] = (result.settledAtPass[`${format}:${at}`] ?? 0) + 1;
                    const lengths = trip.outs.map(out => out.length).join(' ');
                    if (at !== 1) fail(`${format}:unsettled`, `${tag} lengths=${lengths}\nout1=${JSON.stringify(trip.outs[0])}\nout2=${JSON.stringify(trip.outs[1])}`);
                    if (trip.outs[4].length > trip.outs[0].length) fail(`${format}:grows`, `${tag} lengths=${lengths}`);
                    if (trip.outs.some(out => out.includes('empty-paragraph'))) fail(`${format}:emptyParagraph`, `${tag}\nout1=${JSON.stringify(trip.outs[0])}`);
                }
            }
        }
    }
    result.distinctSources = sources.size;
    result.distinctHtml = htmls.size;
    result.distinctEpub = epubs.size;
    result.distinctCombos = combos.size;
    return result;
}

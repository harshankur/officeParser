/**
 * A generator of Markdown holding a lone carriage return, a line separator (U+2028), a paragraph separator
 * (U+2029) or a CRLF inside a line, used by the Markdown line separator tests.
 *
 * Reader: each of the separators in one of 26 block kinds (paragraph, ATX and setext headings, bullet,
 * ordered, `)` and task items, nested items, an item's continuation and `<br>` lines, a quote and a quoted
 * item, a table, a definition, a footnote, an admonition, a disclosure's summary and body, emphasis, a
 * link, a code span, a fence, an HTML block, an item in an item), the words around it with or without a
 * space or `<br>` between them: no word is lost.
 *
 * Writer then reader: text or a picture's caption holding `\r` or `\r\n` in 11 containers (paragraph,
 * heading, inline and paragraph list items, task items, a cell, a Quote paragraph, a definition term and
 * description, an admonition, a disclosure's summary) under 4 image modes, written as Markdown and read
 * back: no word is lost.
 *
 * The random numbers come from a full-period LCG mod 2^32 (Math.imul, a = 1664525, c = 1013904223: c odd,
 * a - 1 divisible by 4), so a seed never repeats within 2^32 draws.
 */
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParser } from '../../src/OfficeParser';

type Node = Record<string, any>;

/** The counts and failures of one run. */
export interface LineSeparatorResult {
    readerDocs: number;
    distinctReaderSources: number;
    distinctReaderKinds: number;
    writerDocs: number;
    distinctWriterAsts: number;
    distinctWriterMarkdown: number;
    distinctWriterContainers: number;
    failures: Record<string, number>;
    examples: Record<string, string[]>;
}

const LINE_SEPARATOR = String.fromCharCode(0x2028);
const PARAGRAPH_SEPARATOR = String.fromCharCode(0x2029);
const SEPARATORS: Record<string, string> = { CR: '\r', LS: LINE_SEPARATOR, PS: PARAGRAPH_SEPARATOR, CRLF: '\r\n' };
const BLOCK_KINDS: Record<string, (line: string) => string> = {
    paragraph: a => a, atx: a => `## ${a}`, setext: a => `${a}\n---`, bullet: a => `- ${a}`, ordered: a => `1. ${a}`, paren: a => `3) ${a}`,
    task: a => `- [ ] ${a}`, taskDone: a => `- [x] ${a}`, nested: a => `- z1\n  - ${a}`, continuation: a => `- z1\n  ${a}`, itemBr: a => `- z1<br>${a}`,
    quote: a => `> ${a}`, quoteItem: a => `> - ${a}`, table: a => `| z1 |\n| --- |\n| ${a} |`, definition: a => `z1\n: ${a}`, footnote: a => `z1[^1]\n\n[^1]: ${a}`,
    admonition: a => `> [!NOTE]\n> ${a}`, details: a => `<details><summary>${a}</summary>\n\nz1\n\n</details>`, detailsBody: a => `<details><summary>z1</summary>\n\n${a}\n\n</details>`,
    emphasis: a => `*${a}*`, link: a => `[${a}](https://e.x)`, codeSpan: a => `\`${a}\``, fence: a => '```\n' + a + '\n```', html: a => `<div>${a}</div>`,
    itemInQuote: a => `> 1. ${a}`, itemInItem: a => `- - ${a}`,
};
const IMAGE_MODES: Array<Record<string, unknown>> = [{}, { includeImages: 'image+ocr-text' }, { includeImages: 'ocr-text-only' }, { includeImages: 'none' }];

const tokensOf = (text: string): Set<string> => new Set(text.match(/[A-Za-z0-9]+/g) ?? []);

/** Runs `count` reader documents and `count` writer documents from `seed`. Failures: a word lost, throws. */
export async function runLineSeparatorGenerator(count: number, seed: number): Promise<LineSeparatorResult> {
    let state = seed >>> 0;
    const rnd = (): number => (state = (Math.imul(state, 1664525) + 1013904223) >>> 0) / 4294967296;
    const pick = <T>(xs: readonly T[]): T => xs[Math.floor(rnd() * xs.length)];
    const result: LineSeparatorResult = { readerDocs: 0, distinctReaderSources: 0, distinctReaderKinds: 0, writerDocs: 0, distinctWriterAsts: 0, distinctWriterMarkdown: 0, distinctWriterContainers: 0, failures: {}, examples: {} };
    const fail = (kind: string, message: string): void => {
        result.failures[kind] = (result.failures[kind] ?? 0) + 1;
        const list = (result.examples[kind] ??= []);
        if (list.length < 3) list.push(message.slice(0, 1200));
    };
    const readerSources = new Set<string>(), readerKinds = new Set<string>();
    const kinds = Object.entries(BLOCK_KINDS);
    const separators = Object.entries(SEPARATORS);

    for (let i = 0; i < count; i++) {
        const [kindName, kind] = i < kinds.length ? kinds[i] : pick(kinds);
        const [separatorName, separator] = pick(separators);
        const a = `w${2 * i + 1}`, b = `w${2 * i + 2}`;
        const source = `z0\n\n${kind(`${a}${pick(['', ' '])}${separator}${pick(['', ' ', '<br>', '  '])}${b}`)}\n\nz9`;
        readerSources.add(source);
        readerKinds.add(kindName);
        result.readerDocs++;
        try {
            const text = String((await (await OfficeParser.parseOffice(Buffer.from(source), { fileType: 'md' })).to('text')).value);
            const present = tokensOf(text);
            const lost = [a, b, 'z0', 'z9'].filter(word => !present.has(word));
            if (lost.length) fail(`reader-lost:${kindName}`, `${separatorName} ${JSON.stringify(source)} => ${JSON.stringify(text)}`);
        } catch (error) { fail('reader-throw', `${kindName}/${separatorName}: ${(error as Error).message}`); }
    }

    const text = (value: string): Node => ({ type: 'text', text: value });
    const paragraph = (children: Node[], metadata?: Node): Node => ({ type: 'paragraph', children, ...(metadata ? { metadata } : {}) });
    const item = (children: Node[], extra: Node = {}): Node => ({ type: 'list', metadata: { listType: pick(['ordered', 'unordered']), indentation: 0, itemIndex: 0, listId: 'L', ...extra }, children });
    const CONTAINERS: Record<string, (line: Node[]) => Node> = {
        paragraph: line => paragraph(line), heading: line => ({ type: 'heading', metadata: { level: 2 }, children: line }),
        item: line => item(line), itemPara: line => item([paragraph(line)]), task: line => item(line, { listType: 'unordered', isTask: true, checked: false }),
        cell: line => ({ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: line }] }] }), quote: line => paragraph(line, { style: 'Quote', quoteId: 'q' }),
        term: line => ({ type: 'definitionList', children: [{ type: 'definitionTerm', children: line }, { type: 'definitionDescription', children: [text('z1')] }] }),
        desc: line => ({ type: 'definitionList', children: [{ type: 'definitionTerm', children: [text('z1')] }, { type: 'definitionDescription', children: line }] }),
        admonition: line => ({ type: 'admonition', metadata: { admonitionType: 'note' }, children: [paragraph(line)] }),
        summary: line => ({ type: 'disclosure', metadata: {}, children: [{ type: 'summary', children: [paragraph(line)] }, paragraph([text('z1')])] }),
    };
    const containers = Object.entries(CONTAINERS);
    const asts = new Set<string>(), markdowns = new Set<string>(), writerContainers = new Set<string>();
    for (let i = 0; i < count; i++) {
        const [containerName, container] = i < containers.length ? containers[i] : pick(containers);
        const mode = pick(IMAGE_MODES);
        const separator = pick(['\r', '\r\n']);
        const a = `w${2 * i + 1}`, b = `w${2 * i + 2}`;
        const inCaption = rnd() < 0.5;
        const line = inCaption ? [text('z2 '), { type: 'image', metadata: { url: 'a.png', caption: `${a}${separator}${b}` } }, text(' z3')] : [text(`z2 ${a}${separator}${b} z3`)];
        const ast = { type: 'md', metadata: {}, attachments: [], content: [paragraph([text('z0')]), container(line), paragraph([text('z9')])] };
        const json = JSON.stringify(ast);
        asts.add(json);
        writerContainers.add(`${containerName}${inCaption ? '/caption' : ''}`);
        result.writerDocs++;
        try {
            const md = String((await OfficeGenerator.generate(JSON.parse(json), 'md', mode as any)).value);
            markdowns.add(md);
            const back = String((await (await OfficeParser.parseOffice(Buffer.from(md), { fileType: 'md' })).to('text')).value);
            const present = tokensOf(back);
            const lost = [a, b, 'z0', 'z2', 'z3', 'z9'].filter(word => !present.has(word));
            if (lost.length) fail(`writer-reader-lost:${containerName}${inCaption ? '/caption' : ''}`, `${JSON.stringify(mode)} ${JSON.stringify(md)} => ${JSON.stringify(back)}`);
        } catch (error) { fail('writer-throw', `${containerName}: ${(error as Error).message}`); }
    }
    result.distinctReaderSources = readerSources.size;
    result.distinctReaderKinds = readerKinds.size;
    result.distinctWriterAsts = asts.size;
    result.distinctWriterMarkdown = markdowns.size;
    result.distinctWriterContainers = writerContainers.size;
    return result;
}

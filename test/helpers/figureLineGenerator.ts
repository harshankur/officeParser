/**
 * A generator of documents that put captioned pictures and every inline kind next to each other in every
 * container that holds a line (paragraph, heading, list item, pipe-table cell, definition), written as
 * Markdown and read back for four passes. Used by the figure-in-a-line tests.
 *
 * The random numbers come from a full-period LCG mod 2^32 (a = 1664525, c odd, a - 1 divisible by 4), so
 * a seed never repeats within 2^32 draws.
 */
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParser } from '../../src/OfficeParser';

type Node = Record<string, any>;

/** The counts and failures of one run. */
export interface FigureLineResult {
    docs: number;
    distinctAsts: number;
    distinctMarkdown: number;
    distinctKindPairs: number;
    distinctConfigContainers: number;
    failures: Record<string, number>;
    examples: Record<string, string[]>;
}

const text = (value: string, extra: Node = {}): Node => ({ type: 'text', text: value, ...extra });

const CONFIGS: Array<Record<string, unknown> | undefined> = [
    undefined,
    { mdConfig: { dialect: 'commonmark' } },
    { mdConfig: { dialect: 'github' } },
    { mdConfig: { dialect: 'pandoc' } },
    { mdConfig: { fallbackToHtml: { itemLineBreaks: false, cellLineBreaks: false } } },
    { mdConfig: { fallbackToHtml: false } },
];
const CONTAINERS = ['paragraph', 'heading', 'item', 'itemPara', 'cell', 'cellPara', 'term', 'definition', 'itemInDefinition'];

/** Runs `count` documents from `seed`. Failures: joined words, lost words, markup read as text, Markdown unsettled by pass 3, throws. */
export async function runFigureLineGenerator(count: number, seed: number): Promise<FigureLineResult> {
    let state = seed >>> 0;
    const rnd = (): number => (state = (Math.imul(state, 1664525) + 1013904223) >>> 0) / 4294967296;
    const pick = <T>(xs: readonly T[]): T => xs[Math.floor(rnd() * xs.length)];
    let wordNo = 0;
    const w = (): string => `w${++wordNo}`;

    // `tight` kinds may sit with no space beside a neighbour without that being one word in the source.
    const KINDS: Record<string, { tight: boolean; make: () => Node }> = {
        plain: { tight: false, make: () => text(w()) },
        bold: { tight: false, make: () => text(w(), { formatting: { bold: true } }) },
        italic: { tight: false, make: () => text(w(), { formatting: { italic: true } }) },
        strike: { tight: false, make: () => text(w(), { formatting: { strikethrough: true } }) },
        underline: { tight: false, make: () => text(w(), { formatting: { underline: true } }) },
        mono: { tight: false, make: () => text(w(), { formatting: { font: 'monospace' } }) },
        kbd: { tight: false, make: () => text(w(), { formatting: { font: 'monospace', keyboard: true } }) },
        highlight: { tight: false, make: () => text(w(), { formatting: { backgroundColor: '#ffff00' } }) },
        link: { tight: false, make: () => text(w(), { metadata: { link: 'https://e.x', linkType: 'external' } }) },
        math: { tight: false, make: () => ({ type: 'code', text: w(), metadata: { math: 'inline' } }) },
        image: { tight: true, make: () => ({ type: 'image', metadata: { url: 'a.png', altText: w() } }) },
        linkedImage: { tight: true, make: () => ({ type: 'image', metadata: { url: 'a.png', altText: w(), link: 'https://e.y' } }) },
        figure: { tight: true, make: () => ({ type: 'image', metadata: { url: 'a.png', altText: w(), caption: `${w()} ${w()}` } }) },
        figure1: { tight: true, make: () => ({ type: 'image', metadata: { url: 'b.png', altText: w(), caption: w() } }) },
        linkedFigure: { tight: true, make: () => ({ type: 'image', metadata: { url: 'a.png', altText: w(), caption: w(), link: 'https://e.y' } }) },
        break: { tight: true, make: () => ({ type: 'break', metadata: { breakType: 'carriageReturn' } }) },
        rule: { tight: true, make: () => ({ type: 'break', metadata: { breakType: 'thematic' } }) },
        para: { tight: true, make: () => ({ type: 'paragraph', children: [text(w())] }) },
    };
    const kindNames = Object.keys(KINDS);
    const blocky = new Set(['para', 'rule']); // kinds a paragraph or heading does not hold

    const run = (container: string): { nodes: Node[]; kinds: string[] } => {
        const n = 1 + Math.floor(rnd() * 4);
        const allowed = kindNames.filter(k => !((container === 'paragraph' || container === 'heading') && blocky.has(k)));
        const kinds = Array.from({ length: n }, () => pick(allowed));
        const nodes: Node[] = [];
        kinds.forEach((k, i) => {
            if (i > 0 && (!(KINDS[k].tight || KINDS[kinds[i - 1]].tight) || rnd() < 0.3)) nodes.push(text(' '));
            nodes.push(KINDS[k].make());
        });
        return { nodes, kinds };
    };

    const list = (id: string, listType: string, children: Node[]): Node => ({ type: 'list', metadata: { listType, indentation: 0, itemIndex: 0, listId: id }, children });
    const build = (container: string, nodes: Node[]): Node => {
        switch (container) {
            case 'paragraph': return { type: 'paragraph', children: nodes };
            case 'heading': return { type: 'heading', metadata: { level: 2 }, children: nodes };
            case 'item': return list('l1', 'ordered', nodes);
            case 'itemPara': return list('l2', 'unordered', [{ type: 'paragraph', children: nodes }]);
            case 'cell': case 'cellPara': {
                const content = container === 'cell' ? nodes : [{ type: 'paragraph', children: nodes }];
                return { type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: content }, { type: 'cell', children: [text(w())] }] }] };
            }
            case 'term': return { type: 'definitionList', children: [{ type: 'definitionTerm', children: nodes }, { type: 'definitionDescription', children: [text(w())] }] };
            case 'definition': return { type: 'definitionList', children: [{ type: 'definitionTerm', children: [text(w())] }, { type: 'definitionDescription', children: nodes }] };
            default: return { type: 'definitionList', children: [{ type: 'definitionTerm', children: [text(w())] }, { type: 'definitionDescription', children: [list('l3', 'unordered', nodes)] }] };
        }
    };

    const wordsOf = (ast: Node): string[] => {
        const out: string[] = [];
        const walk = (n: any): void => {
            if (!n || typeof n !== 'object') return;
            if (Array.isArray(n)) { n.forEach(walk); return; }
            for (const v of [n.text, n.metadata?.altText, n.metadata?.caption]) if (typeof v === 'string') out.push(...(v.match(/w\d+/g) ?? []));
            walk(n.children);
        };
        walk(ast.content);
        return [...new Set(out)];
    };

    const result: FigureLineResult = { docs: count, distinctAsts: 0, distinctMarkdown: 0, distinctKindPairs: 0, distinctConfigContainers: 0, failures: {}, examples: {} };
    const fail = (kind: string, message: string): void => {
        result.failures[kind] = (result.failures[kind] ?? 0) + 1;
        const examples = (result.examples[kind] ??= []);
        if (examples.length < 4) examples.push(message.slice(0, 600));
    };
    const asts = new Set<string>(), markdowns = new Set<string>(), pairs = new Set<string>(), configContainers = new Set<string>();
    for (let i = 0; i < count; i++) {
        wordNo = 0;
        const container = CONTAINERS[i % CONTAINERS.length];
        const configNo = Math.floor(i / CONTAINERS.length) % CONFIGS.length;
        const { nodes, kinds } = run(container);
        kinds.forEach((k, j) => { if (j) pairs.add(`${container}:${kinds[j - 1]}>${k}`); });
        const ast = { type: 'md', metadata: {}, attachments: [], content: [{ type: 'paragraph', children: [text(w())] }, build(container, nodes), { type: 'paragraph', children: [text(w())] }] };
        asts.add(JSON.stringify(ast));
        const words = wordsOf(ast);
        const passes: string[] = [];
        let current: any = ast;
        try {
            for (let pass = 0; pass < 4; pass++) {
                const md = (await OfficeGenerator.generate(current, 'md', CONFIGS[configNo] as any)).value as string;
                passes.push(md);
                current = await OfficeParser.parseOffice(Buffer.from(md), { fileType: 'md' });
            }
        } catch (error) { fail('throw', `${container} ${kinds.join(',')}: ${(error as Error).message}`); continue; }
        markdowns.add(passes[0]);
        configContainers.add(`${configNo}${container}`);
        const out = (await current.to('text')).value as string;
        const tokens = out.match(/[A-Za-z0-9]+/g) ?? [];
        const joined = tokens.filter(t => /w\d/.test(t) && !/^w\d+$/.test(t));
        const present = new Set(tokens);
        const lost = words.filter(x => !present.has(x));
        const tag = `cfg${configNo} ${container} [${kinds.join(',')}] md1=${JSON.stringify(passes[0])}`;
        if (joined.length) fail('joined', `${tag} joined=${joined.join(',')}`);
        if (/<\/?[a-z]|&lt;|&#|&amp;/i.test(out)) fail('markupAsText', `${tag} text=${JSON.stringify(out)}`);
        if (lost.length) fail('lost', `${tag} lost=${lost.join(',')} text=${JSON.stringify(out)}`);
        if (!(passes[0] === passes[1] || passes[1] === passes[2] || passes[2] === passes[3])) fail('unsettled', `${tag} md2=${JSON.stringify(passes[1])} md3=${JSON.stringify(passes[2])}`);
    }
    result.distinctAsts = asts.size;
    result.distinctMarkdown = markdowns.size;
    result.distinctKindPairs = pairs.size;
    result.distinctConfigContainers = configContainers.size;
    return result;
}

/**
 * List items and table cells written as HTML in Markdown keep what they hold: random ASTs of list items
 * each holding a disclosure and of tables each with a merged cell (the two that Markdown writes as HTML),
 * holding at depth up to 3 every node kind (see HTML_PATH_KINDS: paragraphs, headings, formatting, links,
 * breaks, code, inline and display math, pictures, figures, nested and ordered lists, task items, nested
 * tables, admonitions with and without a title, definition lists, disclosures with a summary, without one,
 * open and with a rich summary, a summary standing alone, notes, quotes, embeds, anchored paragraphs),
 * anchor ids on some, written as Markdown and read back three times. After every pass every word is in the
 * text output and set apart (not run into the text before it), every anchor id is in the tree; a fourth
 * write gives the third's Markdown.
 *
 * The generator is a full-period linear congruential generator (Hull-Dobell: c odd, a - 1 divisible by 4,
 * modulus 2^31), so seeds give distinct sequences; the distinct ASTs drawn are counted.
 */
import type { OfficeContentNode, OfficeParserAST } from '../../src/types';

/** Every node kind the ASTs hold. */
export const HTML_PATH_KINDS = ['paragraph', 'heading', 'formatted', 'link', 'break', 'code', 'mathInline', 'mathBlock', 'image', 'figure', 'list', 'nestedList',
    'task', 'table', 'admonition', 'admonitionTitle', 'definitionList', 'disclosure', 'disclosureNoSummary', 'disclosureRich', 'straySummary', 'note', 'quote',
    'embed', 'anchoredParagraph'] as const;

/** One drawn AST: its words and anchor ids, in order, and the kinds it holds. */
export interface HtmlPathAst { ast: OfficeParserAST; words: string[]; anchors: string[]; kinds: Set<string> }

/** What checkHtmlPathRoundTrip found. */
export interface HtmlPathRoundTrip {
    checked: number; distinct: number;
    /** The kinds the distinct ASTs held, of HTML_PATH_KINDS. */
    kinds: number;
    lost: number; joined: number; anchorsLost: number; unsettled: number;
    /** The first few failures, each with its first Markdown. */
    failures: string[];
}

/** A random AST drawer (see the file comment). Each call of the result draws the next AST. */
export function htmlPathAstGenerator(seed: number): () => HtmlPathAst {
    const rnd = (): number => { seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff; return seed / 0x80000000; };
    const pick = <T>(a: readonly T[]): T => a[Math.floor(rnd() * a.length)];
    const chance = (p: number): boolean => rnd() < p;
    const SHALLOW = HTML_PATH_KINDS.filter(k => !/disclosure|table|List|list|admonition/.test(k));
    return () => {
        let wc = 0, ac = 0, nc = 0;
        const W = (): string => `w${wc++}`;
        const A = (): string => `a${ac++}`;
        const kinds = new Set<string>();
        const t = (text = W(), extra: object = {}): any => ({ type: 'text', text, ...extra });
        const p = (...children: any[]): any => ({ type: 'paragraph', children: children.length ? children : [t()] });
        const ids = (): object => (chance(0.4) ? { anchorIds: [A()] } : {});
        const block = (d: number): any => {
            const k = pick(d > 2 ? SHALLOW : HTML_PATH_KINDS);
            kinds.add(k);
            switch (k) {
                case 'paragraph': return p(t(), t(' ' + W()));
                case 'anchoredParagraph': return { type: 'paragraph', metadata: { anchorIds: [A()] }, children: [t()] };
                case 'heading': return { type: 'heading', metadata: { level: 1 + Math.floor(rnd() * 3), ...ids() }, children: [t()] };
                case 'formatted': return p(t(W(), { formatting: { bold: true } }), t(' '), t(W(), { formatting: { italic: true } }));
                case 'link': return p(t(W(), { metadata: { link: 'https://e.x/l' + ac, linkType: 'external' } }));
                case 'break': return p(t(), { type: 'break' }, t());
                case 'code': return { type: 'code', text: W() + '\n' + W(), metadata: { language: 'js', ...ids() } };
                case 'mathInline': return p(t(W() + ' '), { type: 'code', text: 'x + ' + W(), metadata: { math: 'inline' } });
                case 'mathBlock': return { type: 'code', text: 'y + ' + W(), metadata: { math: 'block' } };
                case 'image': return p(t(), { type: 'image', metadata: { url: 'https://e.x/i.png', altText: 'alt' } });
                case 'figure': return { type: 'image', metadata: { url: 'https://e.x/f.png', altText: 'alt', caption: W() } };
                case 'list': return item(d + 1, 0, 'unordered');
                case 'nestedList': return item(d + 1, 1, 'ordered');
                case 'task': return { type: 'list', metadata: { listType: 'unordered', indentation: 0, isTask: true, checked: chance(0.5), listId: 'T' }, children: [t()] };
                case 'table': return table(d + 1);
                case 'admonition': return { type: 'admonition', metadata: { admonitionType: pick(['note', 'tip', 'warning']), ...ids() }, children: blocks(d + 1, 1 + Math.floor(rnd() * 2)) };
                case 'admonitionTitle': return { type: 'admonition', metadata: { admonitionType: 'important', title: W() }, children: [p()] };
                case 'definitionList': return {
                    type: 'definitionList', metadata: ids(),
                    children: [{ type: 'definitionTerm', children: [t()] }, { type: 'definitionDescription', children: chance(0.5) ? [t()] : blocks(d + 1, 1 + Math.floor(rnd() * 2)) }],
                };
                case 'disclosure': return { type: 'disclosure', metadata: { ...(chance(0.3) ? { open: true } : {}), ...ids() }, children: [{ type: 'summary', children: [p()] }, ...blocks(d + 1, Math.floor(rnd() * 3))] };
                case 'disclosureNoSummary': return { type: 'disclosure', metadata: ids(), children: blocks(d + 1, 1) };
                case 'disclosureRich': return { type: 'disclosure', children: [{ type: 'summary', children: [p(t(W(), { formatting: { bold: true } }), t(' ' + W()))] }, ...blocks(d + 1, 1)] };
                case 'straySummary': return { type: 'summary', children: [p()] };
                case 'note': return p(t(W(), { notes: [{ type: 'note', metadata: { noteType: 'footnote', noteId: 'n' + (nc++) }, children: [p()] }] }));
                case 'quote': return { type: 'paragraph', metadata: { style: 'Quote' }, children: [t()] };
                case 'embed': return { type: 'embed', metadata: { url: 'https://e.x/v', label: W() } };
            }
        };
        const blocks = (d: number, n: number): any[] => Array.from({ length: n }, () => block(d));
        const item = (d: number, indentation: number, listType: string): any =>
            ({ type: 'list', metadata: { listType, indentation, listId: 'L' + d, ...ids() }, children: [t(), ...blocks(d, 1 + Math.floor(rnd() * 3))] });
        const table = (d: number): any => ({
            type: 'table',
            children: [0, 1].map(() => ({ type: 'row', children: [0, 1].map(() => ({ type: 'cell', metadata: chance(0.15) ? { colSpan: 2 } : {}, children: chance(0.3) ? [t()] : blocks(d, 1 + Math.floor(rnd() * 2)) })) })),
        });
        const content: any[] = [];
        const n = 1 + Math.floor(rnd() * 3);
        for (let i = 0; i < n; i++) {
            if (rnd() < 0.55) {
                // List items each holding a disclosure, which Markdown writes as an HTML list.
                const items = 1 + Math.floor(rnd() * 3);
                for (let j = 0; j < items; j++) {
                    const it = item(0, j === 0 ? 0 : Math.floor(rnd() * 2), chance(0.5) ? 'ordered' : 'unordered');
                    it.metadata.listId = 'X' + i;
                    it.metadata.itemIndex = j;
                    it.children.splice(1 + Math.floor(rnd() * it.children.length), 0, { type: 'disclosure', children: [{ type: 'summary', children: [p()] }, ...blocks(1, Math.floor(rnd() * 2))] });
                    content.push(it);
                }
            } else {
                // A table with a merged cell, which Markdown writes as an HTML table.
                const tb = table(0);
                tb.children[0].children[0].metadata = { colSpan: 2 };
                content.push(tb);
            }
        }
        const words = Array.from({ length: wc }, (_, i) => `w${i}`);
        const anchors = Array.from({ length: ac }, (_, i) => `a${i}`);
        // Links to every anchor, so the anchors are written.
        if (ac) content.unshift(p(...anchors.map(a => t(a + 'link ', { metadata: { link: '#' + a, linkType: 'internal' } }))));
        return { ast: { type: 'md', metadata: {}, attachments: [], content } as unknown as OfficeParserAST, words, anchors, kinds };
    };
}

const wordsOf = (s: string): string[] => s.match(/(?<![A-Za-z0-9])w\d+(?![0-9])/g) ?? [];

/** Every anchor id in a tree, its notes' too. */
function anchorsIn(nodes: OfficeContentNode[] | undefined, out = new Set<string>()): Set<string> {
    for (const node of nodes ?? []) {
        for (const id of (node.metadata as { anchorIds?: string[] } | undefined)?.anchorIds ?? []) out.add(id);
        anchorsIn(node.children, out);
        anchorsIn(node.notes, out);
    }
    return out;
}

/**
 * Draws `count` ASTs from `seed`, writes each as Markdown with `write`, reads it back with `read` three
 * times, and counts what changed (see the file comment).
 */
export async function checkHtmlPathRoundTrip(
    write: (ast: OfficeParserAST) => Promise<string>,
    read: (markdown: string) => Promise<OfficeParserAST>,
    text: (ast: OfficeParserAST) => Promise<string>,
    count: number, seed: number,
): Promise<HtmlPathRoundTrip> {
    const draw = htmlPathAstGenerator(seed);
    const result: HtmlPathRoundTrip = { checked: 0, distinct: 0, kinds: 0, lost: 0, joined: 0, anchorsLost: 0, unsettled: 0, failures: [] };
    const seen = new Set<string>();
    const kinds = new Set<string>();
    for (let k = 0; k < count; k++) {
        const { ast, words, anchors, kinds: held } = draw();
        const key = JSON.stringify(ast.content);
        if (seen.has(key)) continue;
        seen.add(key);
        for (const kind of held) kinds.add(kind);
        result.checked++;
        const outputs: string[] = [];
        let current = ast;
        let bad = '';
        for (let pass = 0; pass < 3 && !bad; pass++) {
            outputs.push(await write(current));
            current = await read(outputs[pass]);
            const out = await text(current);
            const have = new Set(wordsOf(out));
            const missed = words.filter(w => !have.has(w));
            const gone = missed.filter(w => !new RegExp(`${w}(?![0-9])`).test(out));
            const anchored = anchorsIn(current.content);
            const unanchored = anchors.filter(a => !anchored.has(a));
            if (gone.length) { result.lost++; bad = `pass ${pass + 1} lost ${gone.join(' ')}`; }
            else if (missed.length) { result.joined++; bad = `pass ${pass + 1} joined ${missed.join(' ')}`; }
            else if (unanchored.length) { result.anchorsLost++; bad = `pass ${pass + 1} lost anchors ${unanchored.join(' ')}`; }
        }
        if (!bad && await write(current) !== outputs[2]) { result.unsettled++; bad = 'unsettled after 3 passes'; }
        if (bad && result.failures.length < 3) result.failures.push(`${bad}\n  first ${JSON.stringify(outputs[0])}`);
    }
    result.distinct = seen.size;
    result.kinds = kinds.size;
    return result;
}

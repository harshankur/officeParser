/**
 * Writer and reader agree on disclosures: random ASTs holding disclosures (summaries of every kind of
 * content, bodies of every block kind the Markdown writer writes, nested disclosures, disclosures in list
 * items and table cells, tag text in code, code spans and plain text, aligned paragraphs, notes, breaks)
 * written out and read back, three passes (or more). After the first pass and every pass after, the
 * disclosure tree (who holds which word: summaries, bodies, and the list item or cell a disclosure stands
 * in) is the AST's, every word is kept, every note in a summary or body is kept, and the output is settled
 * (the last two passes write the same). HTML is given four passes: an AST's list item has no alignment,
 * which the HTML reader gives it (`text-align: left`), and a paragraph in a disclosure in the item takes
 * it from the item a pass later.
 *
 * The generator is a full-period linear congruential generator (Hull-Dobell: c odd, a - 1 divisible by 4,
 * modulus 2^31), so seeds give distinct sequences; the distinct ASTs drawn are counted.
 */
import type { OfficeContentNode, OfficeParserAST } from '../../src/types';

/** What checkDisclosureAstRoundTrip found. */
export interface DisclosureRoundTrip {
    checked: number; distinct: number;
    treeDiffers: number; wordsLost: number; notesLost: number; unsettled: number;
    /** The first few failures, each with its tree and output. */
    failures: string[];
}

/** A random AST drawer (see the file comment). Each call of the result draws the next AST. */
export function disclosureAstGenerator(seed: number): () => OfficeParserAST {
    const rnd = (): number => { seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff; return seed / 0x80000000; };
    const pick = <T>(a: T[]): T => a[Math.floor(rnd() * a.length)];
    let wc = 0;
    let notes = 0;
    let lists = 0;
    const W = (): string => `w${wc++}`;
    const TAGS = ['<details>', '</details>', '<summary>', '</summary>', '<details open>', '</DETAILS>'];
    const text = (t: string, extra: object = {}): any => ({ type: 'text', text: t, ...extra });
    const inline = (rich: boolean): any[] => {
        const out: any[] = [];
        const n = 1 + Math.floor(rnd() * 3);
        for (let i = 0; i < n; i++) {
            if (i) out.push(text(' '));
            const r = rnd();
            if (r < 0.4 || !rich) out.push(text(W()));
            else if (r < 0.5) out.push(text(W(), { formatting: { bold: true } }));
            else if (r < 0.58) out.push(text(W(), { formatting: { italic: true } }));
            else if (r < 0.68) out.push(text(`${pick(TAGS)} ${W()}`, { formatting: { font: 'monospace' } }));
            else if (r < 0.76) out.push(text(`${W()} ${pick(TAGS)} x`));
            else if (r < 0.83) out.push(text(W(), { metadata: { link: 'http://e.x/' + pick(['a', '</details>']), linkType: 'external' } }));
            else if (r < 0.9) { notes++; out.push(text(W(), { notes: [{ type: 'note', text: 'a note', children: [text('a note')], metadata: { noteType: 'footnote', noteId: `n${notes}` } }] })); }
            else { out.push(text(W())); out.push({ type: 'break', metadata: { breakType: 'carriageReturn' } }); out.push(text(W())); }
        }
        return out;
    };
    // A list item's or cell's line, a disclosure after it at times (a block the line of Markdown cannot hold).
    const line = (depth: number): any[] => [...inline(true), ...(depth < 3 && rnd() < 0.3 ? [disclosure(depth + 1)] : [])];
    const para = (rich = true, extra: object = {}): any => ({ type: 'paragraph', metadata: { ...extra }, children: inline(rich) });
    const block = (depth: number): any[] => {
        const r = rnd();
        if (r < 0.28) return [para()];
        if (r < 0.34) return [para(true, { alignment: pick(['center', 'right']) })];
        if (r < 0.41) return [{ type: 'heading', metadata: { level: pick([2, 3]) }, children: inline(false) }];
        if (r < 0.52) {
            const ordered = rnd() < 0.5;
            const task = !ordered && rnd() < 0.3;
            const id = `l${++lists}`;
            return Array.from({ length: 1 + Math.floor(rnd() * 3) }, (_, i) => ({
                type: 'list',
                metadata: { listType: ordered ? 'ordered' : 'unordered', indentation: i > 0 && rnd() < 0.3 ? 1 : 0, listId: id, itemIndex: i, ...(task ? { isTask: true, checked: rnd() < 0.5 } : {}) },
                children: line(depth),
            }));
        }
        if (r < 0.6) {
            const cell = (h: boolean): any => ({ type: 'cell', ...(h ? { metadata: { style: 'header' } } : {}), children: h ? inline(true) : line(depth) });
            return [{ type: 'table', children: [{ type: 'row', children: [cell(true), cell(true)] }, { type: 'row', children: [cell(false), cell(false)] }] }];
        }
        if (r < 0.65) return [{ type: 'code', text: `${W()}\n${pick(TAGS)}\n${W()}`, metadata: { language: pick(['', 'html']) } }];
        if (r < 0.7) return [para(true, { style: 'Quote' })];
        if (r < 0.74) return [{ type: 'break', metadata: { breakType: 'thematic' } }];
        if (r < 0.78) return [{ type: 'admonition', metadata: { admonitionType: 'note', sourceSyntax: 'github' }, children: [para()] }];
        if (depth < 3) return [disclosure(depth + 1)];
        return [para()];
    };
    const summary = (depth: number): any[] | undefined => {
        const r = rnd();
        if (r < 0.15) return undefined;
        if (r < 0.45) return [para(false)];
        if (r < 0.65) return [para(true)];
        if (r < 0.72) return [{ type: 'heading', metadata: { level: 3 }, children: inline(true) }];
        if (r < 0.8) return [para(true), para(true)];
        if (r < 0.86) return [...block(3), para(false)];
        if (r < 0.92) return [para(true, { alignment: 'center' })];
        if (depth < 3) return [para(false), disclosure(depth + 1)];
        return [para(false)];
    };
    const disclosure = (depth: number): any => {
        const s = summary(depth);
        const body: any[] = [];
        const n = Math.floor(rnd() * 4);
        for (let i = 0; i < n; i++) body.push(...block(depth));
        return { type: 'disclosure', metadata: rnd() < 0.3 ? { open: true } : {}, children: [...(s ? [{ type: 'summary', children: s }] : []), ...body] };
    };
    return () => {
        wc = 0; notes = 0; lists = 0;
        const content: any[] = [];
        const n = 1 + Math.floor(rnd() * 4);
        for (let i = 0; i < n; i++) content.push(...(rnd() < 0.6 ? [disclosure(1)] : block(0)));
        return { type: 'md', metadata: {}, attachments: [], content } as unknown as OfficeParserAST;
    };
}

const wordsOf = (s: string): string[] => s.match(/(?<![A-Za-z0-9])w\d+(?![0-9])/g) ?? [];

/**
 * A tree's disclosures with the words each holds, in order: `D[summary](body)` for each, the list item
 * (`L(...)`) or cell (`C(...)`) one stands in kept, every other node its words.
 */
export function disclosureTreeOf(nodes: OfficeContentNode[] | undefined): string {
    return (nodes ?? []).map(n => {
        if (n.type === 'disclosure') {
            const kids = n.children ?? [];
            const summary = kids[0]?.type === 'summary' ? kids[0] : undefined;
            return `D[${disclosureTreeOf(summary?.children)}](${disclosureTreeOf(summary ? kids.slice(1) : kids)})`;
        }
        if (n.children?.length) {
            const inner = disclosureTreeOf(n.children);
            return n.type === 'list' && inner.includes('D[') ? `L(${inner})` : n.type === 'cell' && inner.includes('D[') ? `C(${inner})` : inner;
        }
        return wordsOf(n.text ?? '').join(' ');
    }).filter(Boolean).join(' ');
}

/** How many notes disclosures (summaries and bodies) hold. */
function notesIn(nodes: OfficeContentNode[] | undefined, inside = false): number {
    return (nodes ?? []).reduce((sum, n) => sum + (inside ? n.notes?.length ?? 0 : 0) + notesIn(n.children, inside || n.type === 'disclosure'), 0);
}

/** The words of a tree's leaves. */
function leafWords(nodes: OfficeContentNode[] | undefined): string[] {
    return (nodes ?? []).flatMap(n => (n.children?.length ? leafWords(n.children) : wordsOf(n.text ?? '')));
}

/**
 * Draws `count` ASTs from `seed`, writes each with `write`, reads it back with `read` (`passes` passes),
 * and counts what changed (see the file comment).
 */
export async function checkDisclosureAstRoundTrip(
    write: (ast: OfficeParserAST) => Promise<string>,
    read: (output: string) => Promise<OfficeParserAST>,
    text: (ast: OfficeParserAST) => Promise<string>,
    count: number, seed: number, passes = 3,
): Promise<DisclosureRoundTrip> {
    const draw = disclosureAstGenerator(seed);
    const result: DisclosureRoundTrip = { checked: 0, distinct: 0, treeDiffers: 0, wordsLost: 0, notesLost: 0, unsettled: 0, failures: [] };
    const seen = new Set<string>();
    for (let k = 0; k < count; k++) {
        const ast = draw();
        seen.add(JSON.stringify(ast.content));
        const expected = disclosureTreeOf(ast.content);
        const words = leafWords(ast.content).sort().join(' ');
        const notes = notesIn(ast.content);
        const outputs: string[] = [];
        let current = ast;
        let bad = '';
        for (let pass = 0; pass < passes; pass++) {
            outputs.push(await write(current));
            current = await read(outputs[pass]);
            const tree = disclosureTreeOf(current.content);
            if (!bad && tree !== expected) bad = `tree after pass ${pass + 1}: ${tree}`;
            if (!bad && wordsOf(await text(current)).sort().join(' ') !== words) bad = `words after pass ${pass + 1}`;
            if (!bad && notesIn(current.content) !== notes) bad = `notes after pass ${pass + 1}: ${notesIn(current.content)} of ${notes}`;
        }
        result.checked++;
        if (bad.startsWith('tree')) result.treeDiffers++;
        else if (bad.startsWith('words')) result.wordsLost++;
        else if (bad.startsWith('notes')) result.notesLost++;
        else if (outputs[passes - 2] !== outputs[passes - 1]) { result.unsettled++; bad = 'unsettled'; }
        if (bad && result.failures.length < 3) result.failures.push(`${bad}\n  expected ${expected}\n  first ${JSON.stringify(outputs[0])}\n  last ${JSON.stringify(outputs[passes - 1])}`);
    }
    result.distinct = seen.size;
    return result;
}

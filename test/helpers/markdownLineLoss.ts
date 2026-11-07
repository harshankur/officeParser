/**
 * No line of Markdown is lost: random texts of quotes (nested, and with list markers in them), list
 * markers ordered and unordered (`)` and a start of 2 too), empty items with and without a trailing
 * space, blank and quoted blank lines, lines indented by spaces and tabs, leading blank lines, disclosure
 * tags opening and closing (inline and alone on a line), summaries, `<div>` tags, headings and setext
 * lines, each read, its words looked for in the text output, then written as Markdown and read back. A
 * word missing from the text is lost; one there but run into the text before it is joined (counted apart:
 * a `<summary>` outside any disclosure runs into the text around it, as it did in 8.6.0).
 *
 * The generator is a full-period linear congruential generator (Hull-Dobell: c odd, a - 1 divisible by 4,
 * modulus 2^31), so seeds give distinct sequences; the distinct texts drawn are counted.
 */
import type { OfficeParserAST } from '../../src/types';

/** What checkMarkdownLineLoss found. */
export interface MarkdownLineLoss {
    checked: number; distinct: number;
    /** Texts losing a word on reading, and after the Markdown round trip. */
    lost: number; lostAfterRoundTrip: number;
    /** Texts with a word run into the text before it (none lost). */
    joined: number;
    /** The first few failures, each with its text and output. */
    failures: string[];
}

/** A random Markdown text drawer (see the file comment). Each call of the result draws the next text. */
export function markdownLineLossGenerator(seed: number): () => string {
    const rnd = (): number => { seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff; return seed / 0x80000000; };
    const pick = <T>(a: T[]): T => a[Math.floor(rnd() * a.length)];
    let wc = 0;
    const W = (): string => `w${wc++}`;
    const CONTENT: (() => string)[] = [
        () => W(), () => `${W()} ${W()}`, () => '', () => '', () => pick(['1.', '2.', '1)', '-', '*', '+', '1. ', '- ']),
        () => `${pick(['1.', '2.', '-', '*', '1)'])} ${W()}`, () => `${W()} </details>`, () => '</details>', () => '<details>',
        () => `<summary>${W()}</summary>`, () => '<summary>', () => `${W()}</summary>`, () => `<details>\n<summary>${W()}</summary>`,
        () => '</div>', () => '<div>', () => '---', () => '===', () => `# ${W()}`, () => `    ${W()}`,
    ];
    const PREFIXES = ['', '', '', '> ', '>', '> > ', '  ', '   ', '    ', '> - ', '- > ', '\t'];
    return () => {
        wc = 0;
        const lines: string[] = [];
        const n = 2 + Math.floor(rnd() * 9);
        for (let i = 0; i < n; i++) {
            const prefix = pick(PREFIXES);
            for (const line of pick(CONTENT)().split('\n')) lines.push(line || rnd() < 0.5 ? prefix + line : prefix.trimEnd());
        }
        return pick(['', '', '\n', ' \n', '>\n', '<details>\n\n', '<details>\n<summary>s</summary>\n\n']) + lines.join('\n')
            + pick(['\n', '', '\n\n' + W(), '\n</details>\n\n' + W()]);
    };
}

const wordsOf = (s: string): string[] => s.match(/(?<![A-Za-z0-9])w\d+(?![0-9])/g) ?? [];

/** The words of `want` not in `text`, and of those the ones not there even run into the text before them. */
function missing(want: Set<string>, text: string): { missed: string[]; gone: string[] } {
    const have = new Set(wordsOf(text));
    const missed = [...want].filter(w => !have.has(w));
    return { missed, gone: missed.filter(w => !new RegExp(`${w}(?![0-9])`).test(text)) };
}

/**
 * Draws `count` texts from `seed`, reads each with `read`, takes its text with `text` and its Markdown with
 * `write`, reads that back, and counts the texts losing words (see the file comment).
 */
export async function checkMarkdownLineLoss(
    read: (markdown: string) => Promise<OfficeParserAST>,
    text: (ast: OfficeParserAST) => Promise<string>,
    write: (ast: OfficeParserAST) => Promise<string>,
    count: number, seed: number,
): Promise<MarkdownLineLoss> {
    const draw = markdownLineLossGenerator(seed);
    const result: MarkdownLineLoss = { checked: 0, distinct: 0, lost: 0, lostAfterRoundTrip: 0, joined: 0, failures: [] };
    const seen = new Set<string>();
    for (let k = 0; k < count; k++) {
        const source = draw();
        if (seen.has(source)) continue;
        seen.add(source);
        result.checked++;
        const want = new Set(wordsOf(source));
        const ast = await read(source);
        const first = missing(want, await text(ast));
        if (first.gone.length) {
            result.lost++;
            if (result.failures.length < 3) result.failures.push(`lost ${first.gone.join(' ')} reading ${JSON.stringify(source)}`);
            continue;
        }
        if (first.missed.length) result.joined++;
        const markdown = await write(ast);
        const back = missing(want, await text(await read(markdown)));
        if (back.gone.length) {
            result.lostAfterRoundTrip++;
            if (result.failures.length < 3) result.failures.push(`lost ${back.gone.join(' ')} after the round trip of ${JSON.stringify(source)}: ${JSON.stringify(markdown)}`);
        }
    }
    result.distinct = seen.size;
    return result;
}

/**
 * Markdown with no disclosure (`<details>`), read and written as before the disclosure reading changed:
 * the Markdown test files and a seeded fuzz of every other block kind (tables with Markdown cells,
 * aligned wrappers, lists, quotes, code, math, figures, notes, admonitions), each fingerprinted by its
 * tree, metadata, Markdown, HTML and text. `markdownFingerprints.json` holds the fingerprints 8.7.0's
 * reading gave before (15a4c7ec); write it again only for a change to Markdown reading that is meant:
 * `WRITE_MARKDOWN_FINGERPRINTS=1 npx tsx test/testExhaustive.ts`.
 */
import { createHash } from 'crypto';
import * as fs from 'fs';
import * as path from 'path';

/** Where the fingerprints are kept. */
export const MARKDOWN_FINGERPRINTS_FILE = path.join(__dirname, 'markdownFingerprints.json');

/** The pieces the fuzz joins, with one blank line or one line break after each. */
const FRAGMENTS = ['# H', '## H {#x}', 'Para *em* `c`', '<div align="center">', '</div>', '<div style="text-align: center">Centered **b**</div>', '<table><tr><td>', '</td></tr></table>', 'cell *x*',
    '| a | b |\n| - | - |\n| 1 | 2 |', '- a\n- b\n  - c', '1. x', '> q\n>\n> r', '```\ncode\n```', '    ind', '$$\nx\n$$', '![i](a.png)', '<figure><img src="a.png"><figcaption>c</figcaption></figure>',
    'Note[^1]', '[^1]: N', '<p align="center">', '</p>', '<kbd>K</kbd> x', 'Title\n===', '---', '<!-- c -->', ':::note\nbody\n:::', '<center>', '</center>', '<summary>s</summary>', 'text <b>b</b>', '<div>', '<pre>\nx\n</pre>'];

/** How many fuzz documents are fingerprinted. */
const FUZZ_DOCUMENTS = 400;

/** The documents: the Markdown test files, then the fuzz. None holds `<details`. */
export function markdownWithoutDisclosures(): string[] {
    const root = path.join(__dirname, '..', 'files');
    const documents = [fs.readFileSync(path.join(root, 'test.md'), 'utf8'), fs.readFileSync(path.join(root, 'exhaustive', 'markdown.md'), 'utf8')];
    let seed = 7;
    const rnd = (n: number) => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed % n; };
    for (let k = 0; k < FUZZ_DOCUMENTS; k++) {
        const parts = Array.from({ length: 3 + rnd(10) }, () => FRAGMENTS[rnd(FRAGMENTS.length)]);
        documents.push(parts.map(part => part + (rnd(3) ? '\n\n' : '\n')).join(''));
    }
    return documents;
}

/** Reads a buffer as Markdown (OfficeParser.parseOffice of the tree under test). */
export type ParseMarkdown = (buffer: Buffer, config: { fileType: 'md' }) => Promise<{ content: unknown; metadata?: unknown; to: (format: any, options?: any) => Promise<{ value: unknown }> }>;

/**
 * Each document's fingerprint. The reader's one random value (the token standing for a Markdown piece of
 * an HTML table, which a few fuzz documents show) is fixed while it reads, so the fingerprint is the same
 * in every run.
 */
export async function markdownFingerprints(parse: ParseMarkdown, documents = markdownWithoutDisclosures()): Promise<string[]> {
    const out: string[] = [];
    const random = Math.random;
    for (const document of documents) {
        let ast: Awaited<ReturnType<ParseMarkdown>>;
        Math.random = () => 0.7071067811865476;
        try {
            ast = await parse(Buffer.from(document), { fileType: 'md' });
        } finally {
            Math.random = random;
        }
        const parts = [JSON.stringify(ast.content), JSON.stringify(ast.metadata ?? {})];
        for (const format of ['md', 'html', 'text']) parts.push(String((await ast.to(format, { htmlConfig: { standalone: false } })).value));
        out.push(createHash('sha256').update(parts.join('\u0000')).digest('hex').slice(0, 16));
    }
    return out;
}

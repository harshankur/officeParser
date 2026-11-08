/**
 * A generator of documents over the picture matrix: 8 image configurations (every `includeImages` mode,
 * three with `maxInlineImageBytes: 1000`) x 8 Markdown configurations (every dialect, `fallbackToHtml:
 * false`, item and cell line breaks off), cycled, so 64 combinations; pictures (url, small attachment
 * inlined, big attachment over the 1000-byte cap; alt text, link, title, anchors; a caption of one or two
 * words; recognized text of none, one line, two lines, two lines with a blank between) in every container
 * (paragraph, Quote paragraph, heading, ordered, unordered and task items, nested, HTML items through a
 * disclosure, pipe and HTML tables, cells of inline content, paragraphs, headings or items, definition
 * terms and descriptions, admonitions, disclosures, caret and inlined footnotes, a footnote whose body is
 * a picture alone). Each document is written as Markdown and read back three times.
 *
 * With captions, the checks are: no joined words (a `w\d` token that is not a word alone), no lost words
 * (text; a caption always; alt text where the image is written; recognized text where it is written),
 * no markup read as text (after pass 1 and pass 3), and Markdown settled by pass 3 (md2 === md3). The
 * 8.6.0 constructs in the next-release ledger are left out (`clean`): multi-line recognized text on a
 * picture without a caption, a table in a definition, a footnote in a table cell, a space at the end of
 * link text. A pass difference that is only a `data:` picture's `<img>` id (the HTML reader drops it,
 * 8.6.0, ledger) is counted apart, not as unsettled.
 *
 * Without captions, each document's three passes of Markdown and its HTML fragment are fingerprinted, to
 * be compared with the fingerprints of the writers before the round 12 change (pictureMatrixFingerprints.json).
 *
 * The random numbers come from a full-period LCG mod 2^32 (Math.imul, a = 1664525, c = 1013904223: c odd,
 * a - 1 divisible by 4), so a seed never repeats within 2^32 draws.
 */
import { createHash } from 'crypto';
import * as path from 'path';

/** Where the caption-free fingerprints are kept (made with c56459ec's writer and reader passed as the io). */
export const PICTURE_MATRIX_FINGERPRINTS_FILE = path.join(__dirname, 'pictureMatrixFingerprints.json');

type Node = Record<string, any>;

/** The writer and reader under test (OfficeGenerator.generate, OfficeParser.parseOffice of Markdown). */
export interface PictureMatrixIo {
    generate: (ast: any, format: 'md' | 'html', config: any) => Promise<string>;
    parse: (markdown: string) => Promise<any>;
}

/** The counts and failures of one run. */
export interface PictureMatrixResult {
    docs: number;
    distinctAsts: number;
    distinctMarkdown: number;
    configCombos: number;
    figureDocs: number;
    figures: number;
    /** Pass differences that are only a `data:` picture's `<img>` id (ledger), not counted as unsettled. */
    readerDataImgId: number;
    failures: Record<string, number>;
    examples: Record<string, string[]>;
    /** Without captions: each document's fingerprint (its three passes of Markdown and its HTML). */
    fingerprints: string[];
}

const IMAGE_CONFIGS: Array<Record<string, unknown>> = [
    {}, { includeImages: 'image+ocr-text' }, { includeImages: 'ocr-text-only' }, { includeImages: 'none' },
    { maxInlineImageBytes: 1000 }, { includeImages: 'image+ocr-text', maxInlineImageBytes: 1000 }, { includeImages: 'image-only', maxInlineImageBytes: 1000 }, { includeImages: true },
];
const MARKDOWN_CONFIGS: Array<Record<string, unknown>> = [
    {}, { mdConfig: { dialect: 'commonmark' } }, { mdConfig: { dialect: 'github' } }, { mdConfig: { dialect: 'gitlab' } },
    { mdConfig: { dialect: 'obsidian' } }, { mdConfig: { dialect: 'pandoc' } }, { mdConfig: { fallbackToHtml: false } },
    { mdConfig: { fallbackToHtml: { itemLineBreaks: false, cellLineBreaks: false } } },
];
const BIG = Buffer.alloc(3000, 7).toString('base64');
const SMALL = Buffer.alloc(150, 9).toString('base64');

/**
 * Runs `count` documents from `seed`. `captions`: pictures carry captions (60%) and are checked; else none
 * do and each document is fingerprinted. `clean`: the ledger's 8.6.0 constructs are left out.
 */
export async function runPictureMatrix(io: PictureMatrixIo, options: { count: number; seed: number; captions: boolean; clean: boolean }): Promise<PictureMatrixResult> {
    const { count, captions, clean } = options;
    let state = options.seed >>> 0;
    const rnd = (): number => (state = (Math.imul(state, 1664525) + 1013904223) >>> 0) / 4294967296;
    const pick = <T>(xs: readonly T[]): T => xs[Math.floor(rnd() * xs.length)];
    let wordNo = 0;
    const w = (): string => `w${++wordNo}`;
    const text = (value: string, extra: Node = {}): Node => ({ type: 'text', text: value, ...extra });

    let cellDepth = 0;
    let listNo = 0;
    let attachments: Node[] = [];
    const image = (): Node => {
        const meta: Node = { altText: rnd() < 0.8 ? w() : '' };
        const where = rnd();
        if (where < 0.4) meta.url = rnd() < 0.5 ? 'a.png' : 'b c.png';
        else {
            const name = `p${attachments.length}.png`;
            attachments.push({ type: 'image', name, mimeType: 'image/png', extension: 'png', data: where < 0.7 ? SMALL : BIG });
            meta.attachmentName = name;
        }
        if (captions && rnd() < 0.6) meta.caption = rnd() < 0.5 ? w() : `${w()} ${w()}`;
        if (rnd() < 0.15) meta.link = pick(['https://e.y', '#w0']);
        if (rnd() < 0.1) meta.title = w();
        if (rnd() < 0.08) meta.anchorIds = ['id' + w()];
        const node: Node = { type: 'image', metadata: meta };
        const ocr = rnd();
        if (ocr < 0.3) node.text = `${w()} ${w()}`;
        else if (ocr < 0.5) node.text = `${w()}\n${w()}`;
        else if (ocr < 0.55) node.text = `${w()}\n\n${w()}`;
        if (clean && !meta.caption && /\n/.test(node.text ?? '')) node.text = node.text.replace(/\n+/g, ' ');
        return node;
    };
    const isFigure = (node: Node | undefined): boolean => node?.type === 'image' && !!node.metadata?.caption;
    const note = (depth: number): Node => {
        const r = rnd();
        // A body that is a picture alone, a run of inline content, or two paragraphs.
        const body = r < 0.3 ? [{ type: 'paragraph', children: [image()] }]
            : r < 0.8 ? [{ type: 'paragraph', children: inlines(depth + 1) }]
                : [{ type: 'paragraph', children: inlines(depth + 1) }, { type: 'paragraph', children: inlines(depth + 1) }];
        return { ...text(w()), notes: [{ type: 'note', metadata: { noteType: 'footnote', noteId: 'n' + w(), ...(rnd() < 0.1 ? { anchorIds: ['fn' + w()] } : {}) }, children: body }] };
    };
    const inline = (depth: number): Node => {
        const r = rnd();
        if (r < 0.25) return text(w());
        if (r < 0.33) return text(w(), { formatting: pick([{ bold: true }, { italic: true }, { strikethrough: true }, { font: 'monospace' }, { underline: true }]) });
        if (r < 0.37) return text(w(), { metadata: { link: 'https://e.x', linkType: 'external' } });
        if (r < 0.41) return { type: 'code', text: w(), metadata: { math: 'inline' } };
        if (r < 0.7) return image();
        if (r < 0.77) return { type: 'break', metadata: { breakType: 'carriageReturn' } };
        if (r < 0.85 && depth < 1 && !(clean && cellDepth)) return note(depth);
        return text(w(), { formatting: { bold: true, italic: true } });
    };
    // A line's runs: words apart by a space, a separate one or one inside the text; a figure or a line
    // break may touch its neighbours (each must set itself apart), a picture without a caption may not
    // (it runs in its line as in 8.6.0, where touching words join).
    const inlines = (depth: number): Node[] => {
        const n = 1 + Math.floor(rnd() * 4);
        const out: Node[] = [];
        if (rnd() < 0.1) out.push(text(' '));
        let previous: Node | undefined;
        const tight = (node: Node): boolean => isFigure(node) || node.type === 'break';
        for (let i = 0; i < n; i++) {
            const node = inline(depth);
            if (previous && (!(tight(node) || tight(previous)) || rnd() < 0.5)) {
                const last = out[out.length - 1];
                if (rnd() < 0.5 && last.type === 'text' && !last.formatting && !last.notes && !(clean && last.metadata)) out[out.length - 1] = { ...last, text: last.text + ' ' };
                else out.push(text(' '));
            }
            out.push(node);
            previous = node;
        }
        if (rnd() < 0.1) out.push(text(' '));
        return out;
    };
    const items = (depth: number, inCell: boolean): Node[] => {
        const id = 'L' + (++listNo);
        const out: Node[] = [];
        const k = 1 + Math.floor(rnd() * 3);
        let indentation = 0;
        const listType = pick(['ordered', 'unordered', 'task']);
        for (let i = 0; i < k; i++) {
            const meta: Node = { listType: listType === 'ordered' ? 'ordered' : 'unordered', indentation, itemIndex: i, listId: id };
            if (listType === 'task') { meta.isTask = true; meta.checked = rnd() < 0.5; }
            const r = rnd();
            // An item holding a disclosure is written as an HTML item.
            const children = r < 0.6 ? inlines(depth) : r < 0.85 || inCell ? [{ type: 'paragraph', children: inlines(depth) }]
                : [{ type: 'paragraph', children: inlines(depth) }, { type: 'disclosure', metadata: {}, children: [{ type: 'summary', children: [{ type: 'paragraph', children: [text(w())] }] }, { type: 'paragraph', children: inlines(depth) }] }];
            out.push({ type: 'list', metadata: meta, children });
            indentation = Math.min(indentation + (rnd() < 0.5 ? 1 : 0), 2);
        }
        return out;
    };
    const cellBlocks = (depth: number): Node[] => {
        const r = rnd();
        if (r < 0.4) return inlines(depth);
        if (r < 0.65) return [{ type: 'paragraph', children: inlines(depth) }];
        if (r < 0.85) return [{ type: 'heading', metadata: { level: 2 }, children: inlines(depth) }];
        return items(depth, true);
    };
    const block = (depth: number, inCell = false): Node[] => {
        const r = rnd();
        if (r < 0.2 || depth > 2) return [{ type: 'paragraph', children: inlines(depth) }];
        if (r < 0.27) return [{ type: 'paragraph', metadata: { style: 'Quote', quoteId: 'q' + w() }, children: inlines(depth) }];
        if (r < 0.37) return [{ type: 'heading', metadata: { level: 1 + Math.floor(rnd() * 3) }, children: inlines(depth) }];
        if (r < 0.52) return items(depth, inCell);
        if (r < 0.66 && !inCell) {
            const span = rnd() < 0.35;
            const cell = (): Node => { cellDepth++; try { return { type: 'cell', children: cellBlocks(depth) }; } finally { cellDepth--; } };
            const rows = [{ type: 'row', children: span ? [{ ...cell(), metadata: { colSpan: 2 } }] : [cell(), cell()] }, { type: 'row', children: [cell(), cell()] }];
            return [{ type: 'table', children: rows }];
        }
        if (r < 0.76) return [{ type: 'definitionList', children: [{ type: 'definitionTerm', children: inlines(depth) }, { type: 'definitionDescription', children: rnd() < 0.7 ? inlines(depth) : block(depth + 1, inCell || clean) }] }];
        if (r < 0.86) return [{ type: 'admonition', metadata: { admonitionType: pick(['note', 'tip', 'warning']) }, children: rnd() < 0.5 ? inlines(depth) : block(depth + 1, inCell) }];
        if (r < 0.94) return [{ type: 'disclosure', metadata: {}, children: [{ type: 'summary', children: [{ type: 'paragraph', children: [text(w())] }] }, ...block(depth + 1, inCell), ...block(depth + 1, inCell)] }];
        return [{ type: 'paragraph', children: [text(w()), text(' '), note(depth)] }];
    };

    // The words a document must keep: text; a picture's caption always, its alt text where its image is
    // written, its recognized text where written (the mode's, or in the image's place over the cap in
    // 'image-only'). An uncaptioned picture over the cap in an HTML cell keeps its name (8.6.0), so its
    // text in its place is not counted.
    const wordsOf = (ast: Node, config: Node): string[] => {
        const out = new Set<string>();
        const mode = config.includeImages === true || config.includeImages === undefined ? 'image-only' : config.includeImages;
        const cap = config.maxInlineImageBytes ?? Infinity;
        const add = (value?: string): void => (value?.match(/w\d+/g) ?? []).forEach(word => out.add(word));
        const walk = (node: any): void => {
            if (!node || typeof node !== 'object') return;
            if (Array.isArray(node)) { node.forEach(walk); return; }
            if (node.type === 'image') {
                const meta = node.metadata ?? {};
                const big = !meta.url && ast.attachments.find((a: Node) => a.name === meta.attachmentName)?.data === BIG && 3000 > cap;
                const replaced = mode === 'image-only' && big && !!node.text?.trim();
                if (meta.caption) add(meta.caption);
                if ((mode === 'image-only' || mode === 'image+ocr-text') && !replaced) add(meta.altText);
                if (mode === 'image+ocr-text' || mode === 'ocr-text-only' || (replaced && meta.caption)) add(node.text);
                return;
            }
            if (typeof node.text === 'string' && node.type !== 'note') add(node.text);
            walk(node.children);
            walk(node.notes);
        };
        walk(ast.content);
        return [...out];
    };

    const result: PictureMatrixResult = { docs: count, distinctAsts: 0, distinctMarkdown: 0, configCombos: 0, figureDocs: 0, figures: 0, readerDataImgId: 0, failures: {}, examples: {}, fingerprints: [] };
    const fail = (kind: string, message: string): void => {
        result.failures[kind] = (result.failures[kind] ?? 0) + 1;
        const examples = (result.examples[kind] ??= []);
        if (examples.length < 3) examples.push(message.slice(0, 1200));
    };
    const asts = new Set<string>(), markdowns = new Set<string>(), combos = new Set<string>();
    // The Markdown reader's one random value (a token standing for a Markdown piece of an HTML table) is
    // fixed, so a run is the same every time.
    const random = Math.random;
    Math.random = () => 0.7071067811865476;
    try {
        for (let i = 0; i < count; i++) {
            wordNo = 0; listNo = 0; attachments = [];
            const config = { ...IMAGE_CONFIGS[i % IMAGE_CONFIGS.length], ...MARKDOWN_CONFIGS[Math.floor(i / IMAGE_CONFIGS.length) % MARKDOWN_CONFIGS.length] };
            combos.add(JSON.stringify(config));
            const content = [{ type: 'paragraph', children: [text(w())] }, ...block(0), ...block(0), { type: 'paragraph', children: [text(w())] }];
            const json = JSON.stringify({ type: 'md', metadata: {}, attachments, content });
            asts.add(json);
            const figures = (json.match(/"caption"/g) ?? []).length;
            if (figures) result.figureDocs++;
            result.figures += figures;
            const md: string[] = [], texts: string[] = [];
            let ast: any = JSON.parse(json);
            try {
                for (let pass = 0; pass < 3; pass++) {
                    md.push(await io.generate(ast, 'md', config));
                    ast = await io.parse(md[pass]);
                    texts.push(String((await ast.to('text')).value));
                }
            } catch (error) { fail('throw', `doc ${i} ${JSON.stringify(config)}: ${(error as Error).stack?.slice(0, 500)}`); continue; }
            markdowns.add(md[0]);
            if (!captions) {
                const html = await io.generate(JSON.parse(json), 'html', { ...config, htmlConfig: { standalone: false } });
                result.fingerprints.push(createHash('sha256').update([...md, html].join('\u0000')).digest('hex').slice(0, 16));
                continue;
            }
            const tag = `doc ${i} ${JSON.stringify(config)} md1=${JSON.stringify(md[0].replace(/data:image\/png;base64,[A-Za-z0-9+/=]+/g, 'DATA'))}`;
            const expected = wordsOf(JSON.parse(json), config);
            for (const [pass, out] of [[1, texts[0]], [3, texts[2]]] as const) {
                const present = new Set(out.match(/[A-Za-z0-9]+/g) ?? []);
                // (Markup shown as text is reported as markup: its attributes are no joined words.)
                const joined = (out.replace(/<[^>]*>/g, ' ').match(/[A-Za-z0-9]+/g) ?? []).filter(x => /w\d/.test(x) && !/^w\d+$/.test(x));
                const lost = expected.filter(x => !present.has(x));
                if (joined.length) fail('joined', `${tag} pass ${pass} joined=${joined.join(',')} text=${JSON.stringify(out)}`);
                if (lost.length) fail('lost', `${tag} pass ${pass} lost=${lost.join(',')} text=${JSON.stringify(out)}`);
                if (/<\/?[a-z]|&lt;|&gt;|&#|&amp;/i.test(out)) fail('markupAsText', `${tag} pass ${pass} text=${JSON.stringify(out)}`);
            }
            if (md[1] !== md[2]) {
                const withoutDataImgIds = (m: string): string => m.replace(/(<img src="data:[^"]*"[^>]*?) id="[^"]*"/g, '$1');
                if (withoutDataImgIds(md[1]) === withoutDataImgIds(md[2])) result.readerDataImgId++;
                else fail('unsettled', `${tag}\nmd2=${JSON.stringify(md[1])}\nmd3=${JSON.stringify(md[2])}`);
            }
        }
    } finally {
        Math.random = random;
    }
    result.distinctAsts = asts.size;
    result.distinctMarkdown = markdowns.size;
    result.configCombos = combos.size;
    return result;
}

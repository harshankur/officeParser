/**
 * Unit tests for the pure PDF text-assembly helpers.
 *
 * `textLayout.ts` takes plain data in and returns AST nodes, so the layouts that break it (hanging
 * indents, tight lists, symbol-font bullets, scattered invoice rows, a caption under a body line)
 * can be built as synthetic runs and asserted directly, with no PDF and no OCR. That matters here:
 * the fast suite skips the PDF fixture, so without these the geometric path has no fast coverage at
 * all. Every case below is a layout that a real producer emits; the NEGATIVE ones are the shapes
 * that must NOT become lists or tables.
 *
 * Registered from `testOfficeParser.ts`, so it reports through the same table as everything else.
 */

import { blockToNodes, buildLines, computeDocContext, detectTables, PageContext, recoverTaggedGrids, segmentIntoBlocks } from '../../src/parsers/pdf/textLayout';
import { buildTaggedNodes } from '../../src/parsers/pdf/structTree';
import { makeColorLookup } from '../../src/parsers/pdf/pdfColor';
import { PdfLayoutConfig, RawRun } from '../../src/parsers/pdf/pdfTypes';
import { layoutOcrText, sniffImageMime } from '../../src/utils/ocrUtils';
import { OfficeContentNode } from '../../src/types';
import { DEFAULT_OFFICE_PARSER_CONFIG } from '../../src/defaults';

/** Mirrors `FeatureTest` in testOfficeParser.ts (structurally, so results can be pushed there). */
interface LayoutTest {
    category: string;
    feature: string;
    fileType: string;
    result: { status: 'PASS' | 'FAIL' | 'WARN' | 'SKIP'; expected: any; actual: any; details: string };
}

/** The shipped PDF layout defaults with deliberate test overrides. */
const CFG: PdfLayoutConfig = {
    ...DEFAULT_OFFICE_PARSER_CONFIG.pdfParserConfig,
    useTags: false,
    extractTextColor: false,
    includeBounds: false,
    separateProcess: false,
};

const PAGE: PageContext = { pageNumber: 1, authoredW: 612, authoredH: 792, rotation: 0 };

interface RunOpts { bold?: boolean; width?: number }

/** Builds one run at (x, baseline). Width defaults to a rough 0.5em per character. */
function run(text: string, x: number, yBaseline: number, fontSize = 12, opts: RunOpts = {}): RawRun {
    return {
        text, x, yBaseline,
        yTop: yBaseline - fontSize * 0.8,
        width: opts.width ?? text.length * fontSize * 0.5,
        height: fontSize * 1.2,
        fontSize,
        dir: 'ltr',
        angle: 0,
        mcid: null,
        inArtifact: false,
        formatting: opts.bold ? { bold: true } : {},
    };
}

/** Builds a paragraph NODE with authored-space bounds, as the tagged path emits before finalizeBounds. */
function paraNode(text: string, x: number, y: number, width: number, height: number): OfficeContentNode {
    const bounds = { x, y, width, height };
    return { type: 'paragraph', text, children: [{ type: 'text', text, bounds }], bounds };
}

/** Runs the whole geometric path over a page's runs and returns the emitted nodes. */
function nodesOf(runs: RawRun[]): OfficeContentNode[] {
    const doc = computeDocContext(runs, CFG, '\n');
    const out: OfficeContentNode[] = [];
    for (const block of segmentIntoBlocks(buildLines(runs, CFG), CFG)) out.push(...blockToNodes(block, PAGE, doc));
    return out;
}

/** Compact rendering of the emitted nodes, as `type:text` pairs. */
function shapeOf(nodes: OfficeContentNode[]): string[] {
    return nodes.map(n => `${n.type}:${(n.text ?? '').trim()}`);
}

function tablesOf(runs: RawRun[]): { node: OfficeContentNode; y: number }[] {
    const doc = computeDocContext(runs, CFG, '\n');
    return detectTables(buildLines(runs, CFG), PAGE, doc).tables;
}

/**
 * Builds a hanging-indent bullet list: marker at 72, item text at 90, item 2 wrapping onto a
 * continuation line in the hanging indent. `itemGap` is the baseline step between items (a tight
 * list uses the same 14.4 as the wrap, a spaced one leaves more).
 */
function bulletList(itemGap: number): RawRun[] {
    const F = 12, WRAP = 14.4;
    const runs: RawRun[] = [];
    let y = 100;
    const item = (text: string) => runs.push(run('•', 72, y, F, { width: 6 }), run(text, 90, y, F));
    item('First item is short');
    y += itemGap;
    item('Second item wraps onto a');
    y += WRAP;
    runs.push(run('second line under the hanging indent here.', 90, y, F));
    y += itemGap;
    item('Third item is short');
    return runs;
}

export async function testTextLayout(): Promise<LayoutTest[]> {
    const results: LayoutTest[] = [];
    const category = 'PDF Text Layout';
    const add = (feature: string, pass: boolean, expected: any, actual: any, details = '') =>
        results.push({ category, feature, fileType: 'pdf', result: { status: pass ? 'PASS' : 'FAIL', expected, actual, details } });

    // ── H6: hanging indents, tight lists, symbol-font bullets ───────────────
    {
        const shape = shapeOf(nodesOf(bulletList(14.4)));
        const ok = shape.length === 3 && shape.every(s => s.startsWith('list:'))
            && shape[1] === 'list:Second item wraps onto a second line under the hanging indent here.';
        add('Tight bullet list keeps its wrapped line', ok, '3 items, item 2 carries its wrap', shape.join(' | '));
    }
    {
        const shape = shapeOf(nodesOf(bulletList(25.9)));
        const ok = shape.length === 3 && shape.every(s => s.startsWith('list:'))
            && shape[1] === 'list:Second item wraps onto a second line under the hanging indent here.';
        add('Spaced bullet list keeps its wrapped line', ok, '3 items, item 2 carries its wrap', shape.join(' | '));
    }
    {
        // The bullet is drawn in a symbol font, so it is always its own fragment: the marker test has
        // to run on the whole line or every such list stays plain prose.
        const runs: RawRun[] = [];
        let y = 100;
        for (const t of ['Alpha item text', 'Beta item text', 'Gamma item text']) {
            runs.push(run('•', 72, y, 12, { bold: true, width: 6 }), run(t, 90, y, 12));
            y += 14.4;
        }
        const nodes = nodesOf(runs);
        const shape = shapeOf(nodes);
        const stripped = nodes.every(n => !(n.text ?? '').includes('•') && !(n.children ?? []).some(c => c.text === ''));
        add('Symbol-font bullets are detected', shape.length === 3 && shape.every(s => s.startsWith('list:')) && stripped,
            '3 items, markers stripped, no empty runs', shape.join(' | '));
    }
    {
        // Marker and text merge into one line here, so the numbered marker prefixes the line text.
        const F = 12, L = 14.4;
        const runs: RawRun[] = [];
        let y = 100;
        runs.push(run('1.', 72, y, F, { width: 12 }), run('Step one wraps onto a second line', 90, y, F)); y += L;
        runs.push(run('and continues on this line.', 90, y, F)); y += L;
        runs.push(run('2.', 72, y, F, { width: 12 }), run('Step two is short.', 90, y, F)); y += L;
        runs.push(run('3.', 72, y, F, { width: 12 }), run('Step three is short.', 90, y, F));
        const shape = shapeOf(nodesOf(runs));
        const ok = shape.length === 3 && shape.every(s => s.startsWith('list:'))
            && shape[0] === 'list:Step one wraps onto a second line and continues on this line.';
        add('Tight numbered list is recovered', ok, '3 ordered items, item 1 carries its wrap', shape.join(' | '));
    }
    {
        // Vertically stacked glyphs in a narrow cell: "Su" over "n" is one token ("Sun"), but "12"
        // over "Jan" is two ("12 Jan"). Fold only when BOTH pieces are 1-2 chars (the a475ec5 fix must
        // not glue longer stacked lines together).
        const sun = shapeOf(nodesOf([run('Su', 100, 100, 12, { width: 12 }), run('n', 100, 112, 12, { width: 6 })]));
        const janY = shapeOf(nodesOf([run('12', 100, 100, 12, { width: 12 }), run('Jan', 100, 112, 12, { width: 18 })]));
        add('Stacked short glyphs join, longer stacks stay split',
            sun.length === 1 && sun[0] === 'paragraph:Sun' && janY.length === 1 && janY[0] === 'paragraph:12 Jan',
            'Su/n -> "Sun", 12/Jan -> "12 Jan"', `${sun.join('|')} ;; ${janY.join('|')}`);
        // A digit stack still folds ("1"/"1" -> "11"), but two stacked upper-case abbreviations are two
        // real tokens and must keep their space ("US"/"UK" -> "US UK", not "USUK").
        const dd = shapeOf(nodesOf([run('1', 100, 100, 12, { width: 6 }), run('1', 100, 112, 12, { width: 6 })]));
        const abbr = shapeOf(nodesOf([run('US', 100, 100, 12, { width: 12 }), run('UK', 100, 112, 12, { width: 12 })]));
        add('Digit stacks fold but upper-case abbreviation stacks keep their space',
            dd.length === 1 && dd[0] === 'paragraph:11' && abbr.length === 1 && abbr[0] === 'paragraph:US UK',
            '1/1 -> "11", US/UK -> "US UK"', `${dd.join('|')} ;; ${abbr.join('|')}`);
    }
    {
        // A hostile PDF can put a run under an enormous text matrix; the colour lookup must stay bounded
        // regardless of the reported font size (its y-window is capped independent of fontSize).
        const lk = makeColorLookup([{ x: 105, y: 500, color: '#ff0000' }]);
        const t0 = performance.now();
        for (let i = 0; i < 2000; i++) lk(100, 500, 1e9, 50);
        const ms = performance.now() - t0;
        const normal = lk(100, 500, 12, 50);
        add('Colour lookup is bounded under a hostile font size', ms < 100 && normal === '#ff0000',
            'fast at fontSize 1e9 and still matches at 12pt', `${ms.toFixed(1)}ms, normal=${normal}`);
    }
    {
        // A crafted page can crowd N colour marks and N runs on one baseline. The lookup must stay near
        // linear (binary-searched bucket), not O(N x marks). Correctness: a uniform-colour window returns
        // that colour, a mixed one returns undefined, and over the mark cap it degrades to no colour.
        const N = 200_000;
        const marks = [];
        for (let i = 0; i < N; i++) marks.push({ x: i % 1000, y: 100, color: '#ff0000' });
        const t0 = performance.now();
        const lk = makeColorLookup(marks);
        let hits = 0;
        for (let i = 0; i < N; i++) if (lk(0, 100, 12, 1000)) hits++;
        const ms = performance.now() - t0;
        const mixed = makeColorLookup([{ x: 10, y: 50, color: '#00ff00' }, { x: 20, y: 50, color: '#0000ff' }])(5, 50, 12, 30);
        const overCap = makeColorLookup(new Array(300_001).fill({ x: 1, y: 1, color: '#111111' }))(0, 1, 12, 5);
        add('Colour lookup stays near-linear with many marks on one baseline',
            ms < 2000 && hits === N && mixed === undefined && overCap === undefined,
            'build + N queries over N marks is fast; uniform hit, mixed undefined, over-cap undefined',
            `${ms.toFixed(0)}ms, hits=${hits}, mixed=${mixed}, overCap=${overCap}`);
    }
    {
        // Two bibliography entries, each wrapping under a hanging indent.
        const F = 10, L = 12;
        const runs: RawRun[] = [];
        let y = 100;
        runs.push(run('Adams, J. (2019). A study of things. Journal of Things,', 72, y, F)); y += L;
        runs.push(run('12(3), 45-67. Publisher, City.', 90, y, F)); y += L;
        runs.push(run('Baker, K. (2021). Another study entirely. Journal of', 72, y, F)); y += L;
        runs.push(run('Other Things, 4(1), 8-19. Publisher, City.', 90, y, F));
        const shape = shapeOf(nodesOf(runs));
        const ok = shape.length === 2 && shape[0].startsWith('paragraph:Adams') && shape[1].startsWith('paragraph:Baker');
        add('Bibliography hanging indents group per entry', ok, '2 paragraphs, one per entry', shape.join(' | '));
    }
    {
        // NEGATIVE: ordinary prose whose first line is indented is one paragraph, not two.
        const F = 12, L = 14.4;
        const runs: RawRun[] = [];
        let y = 100;
        runs.push(run('This first line of the paragraph is indented by a tab.', 90, y, F)); y += L;
        runs.push(run('The rest of the paragraph flows back to the left margin', 72, y, F)); y += L;
        runs.push(run('and continues for several more lines of body prose here', 72, y, F)); y += L;
        runs.push(run('until the paragraph finally ends on this last line.', 72, y, F));
        const shape = shapeOf(nodesOf(runs));
        add('First-line-indented prose stays one paragraph', shape.length === 1 && shape[0].startsWith('paragraph:'), '1 paragraph', shape.join(' | '));
    }
    {
        // NEGATIVE: two such paragraphs still split at the indent, at the right boundary.
        const F = 12, L = 14.4;
        const runs: RawRun[] = [];
        let y = 100;
        runs.push(run('First paragraph opening line is indented here.', 90, y, F)); y += L;
        runs.push(run('It continues at the left margin for a while.', 72, y, F)); y += L;
        runs.push(run('Second paragraph opening line is indented too.', 90, y, F)); y += L;
        runs.push(run('And it also continues at the left margin.', 72, y, F));
        const shape = shapeOf(nodesOf(runs));
        const ok = shape.length === 2 && shape[0].startsWith('paragraph:First paragraph') && shape[1].startsWith('paragraph:Second paragraph');
        add('Two first-line-indented paragraphs split correctly', ok, '2 paragraphs', shape.join(' | '));
    }
    {
        // NEGATIVE: a two-column article of wrapped lines is prose, never a list and never a table.
        const F = 10, L = 12;
        const colText = [
            'Lorem ipsum dolor sit amet, consectetur',
            'adipiscing elit, sed do eiusmod tempor',
            'incididunt ut labore et dolore magna',
            'aliqua. Ut enim ad minim veniam, quis',
            'nostrud exercitation ullamco laboris.',
            'Duis aute irure dolor in reprehenderit',
            'in voluptate velit esse cillum dolore eu',
            'fugiat nulla pariatur excepteur sint.',
        ];
        const runs: RawRun[] = [];
        for (let c = 0; c < 2; c++) {
            let y = 100;
            for (const t of colText) { runs.push(run(t, 72 + c * 250, y, F)); y += L; }
        }
        const nodes = nodesOf(runs);
        const lists = nodes.filter(n => n.type === 'list').length;
        add('Multi-column article yields no lists', lists === 0, 0, lists, shapeOf(nodes).join(' | ').slice(0, 120));
        add('Multi-column article yields no tables', tablesOf(runs).length === 0, 0, tablesOf(runs).length);
    }

    // ── M4: tables need vertically adjacent rows ────────────────────────────
    {
        const runs: RawRun[] = [];
        const row = (y: number, cells: [string, number][]) => { for (const [t, x] of cells) runs.push(run(t, x, y, 10)); };
        row(100, [['Invoice No:', 72], ['12345', 220], ['Date:', 380]]);
        let y = 120;
        for (let i = 0; i < 6; i++) { runs.push(run('This is a line of ordinary prose that runs the full width of the page here.', 72, y, 10)); y += 14; }
        row(240, [['Bill To:', 72], ['ACME', 220], ['Ship To:', 380]]);
        y = 260;
        for (let i = 0; i < 6; i++) { runs.push(run('More ordinary prose filling the space between the two header rows below.', 72, y, 10)); y += 14; }
        row(500, [['Subtotal', 72], ['$10', 220], ['Tax', 380]]);
        const tables = tablesOf(runs);
        add('Scattered invoice rows are not a table', tables.length === 0, 0, tables.length, 'rows 100/240/500 with prose between');
    }
    {
        const runs: RawRun[] = [];
        const row = (y: number, cells: [string, number][]) => { for (const [t, x] of cells) runs.push(run(t, x, y, 10)); };
        row(100, [['Name', 72], ['Qty', 220], ['Price', 380]]);
        row(114, [['Widget', 72], ['2', 220], ['$4.00', 380]]);
        row(128, [['Gadget', 72], ['1', 220], ['$9.50', 380]]);
        row(142, [['Doohickey', 72], ['7', 220], ['$1.25', 380]]);
        const tables = tablesOf(runs);
        const rows = tables[0]?.node.children?.length ?? 0;
        add('Genuine grid is still a table', tables.length === 1 && rows === 4, '1 table of 4 rows', `${tables.length} table(s), ${rows} rows`);
    }

    // ── M5: hybrid grid recovery over tag-derived paragraphs ────────────────
    // Builds paragraph NODES (what the tagged path emits), one per grid cell, and asserts
    // recoverTaggedGrids folds the calendar into a table while leaving the title paragraph, and never
    // invents a table on prose or on a too-narrow layout.
    {
        // Seven columns centred at 100..460 (calendar day/number cells are centred over the column),
        // a header row and three week rows, plus a wide title whose centre falls off every column.
        const centers = [100, 160, 220, 280, 340, 400, 460];
        const nodes: OfficeContentNode[] = [];
        nodes.push(paraNode('December 2007', 250, 60, 120, 16)); // title: centre 310, off-grid
        const gridRow = (yTop: number, texts: string[]) => texts.forEach((t, c) => {
            const w = 8; nodes.push(paraNode(t, centers[c] - w / 2, yTop, w, 12));
        });
        gridRow(100, ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']);
        gridRow(140, ['1', '2', '3', '4', '5', '6', '7']);
        gridRow(180, ['8', '9', '10', '11', '12', '13', '14']);
        gridRow(220, ['15', '16', '17', '18', '19', '20', '21']);
        const out = recoverTaggedGrids(nodes);
        const table = out.find(n => n.type === 'table');
        const cols = table?.children?.[0]?.children?.length ?? 0;
        const rows = table?.children?.length ?? 0;
        const titleKept = out.some(n => n.type === 'paragraph' && (n.text || '').trim() === 'December 2007');
        add('Calendar paragraphs fold into a table', out.length === 2 && rows === 4 && cols === 7 && titleKept,
            'title paragraph + 4x7 table', `${out.length} nodes, ${rows}x${cols} table, titleKept=${titleKept}`);
    }
    {
        // Prose paragraphs (each far longer than a cell) must never become a table.
        const nodes: OfficeContentNode[] = [];
        for (let i = 0; i < 6; i++) nodes.push(paraNode('This is an ordinary sentence of prose that fills the line ' + i, 72, 100 + i * 16, 400, 12));
        const out = recoverTaggedGrids(nodes);
        add('Prose paragraphs are not a table', !out.some(n => n.type === 'table') && out.length === 6, 'no table, 6 paragraphs', `${out.filter(n => n.type === 'table').length} table(s), ${out.length} nodes`);
    }
    {
        // A two-column stack of short paragraphs is not a grid: fewer than three columns.
        const nodes: OfficeContentNode[] = [];
        for (let i = 0; i < 5; i++) { nodes.push(paraNode('L' + i, 100, 100 + i * 20, 8, 12), paraNode('R' + i, 300, 100 + i * 20, 8, 12)); }
        const out = recoverTaggedGrids(nodes);
        add('Two-column short paragraphs are not a table', !out.some(n => n.type === 'table'), 'no table', `${out.filter(n => n.type === 'table').length} table(s)`);
    }

    // ── Tagged TOC entries: a dot leader is one space, however many runs of dots it is split into ──
    {
        const entry = (texts: string[]) => {
            const runsByMcid = new Map<string, RawRun[]>();
            const kids = texts.map((t, i) => { runsByMcid.set(`m${i}`, [{ ...run(t, 72 + i * 60, 100), mcid: `m${i}` }]); return { type: 'content', id: `m${i}` }; });
            const all = [...runsByMcid.values()].flat();
            const { nodes } = buildTaggedNodes({ role: 'Root', children: [{ role: 'TOCI', children: kids }] }, runsByMcid, PAGE, computeDocContext(all, CFG, '\n'), { ignoreNotes: false, listCounter: { n: 0 } });
            return (nodes[0]?.text ?? '').replace(/\s+/g, ' ');
        };
        const split = entry(['Scope', '....', '....', '....', '12']);
        const joined = entry(['Scope ................ 12']);
        add('TOC dot leaders split into runs collapse to one space', split === 'Scope 12' && joined === 'Scope 12', 'Scope 12 (both)', `${JSON.stringify(split)} / ${JSON.stringify(joined)}`);
    }

    // ── L-SS: super/subscript merging vs a caption ──────────────────────────
    {
        const lines = buildLines([run('H', 72, 100, 12, { width: 8 }), run('2', 80, 101.2, 8, { width: 4 })], CFG);
        const text = lines.map(l => l.fragments.map(f => f.text).join('')).join('|');
        const sub = lines[0]?.fragments.some(f => f.formatting.subscript);
        add('Subscript folds into its line', lines.length === 1 && text === 'H2' && !!sub, 'H2 (subscript)', `${text} (${lines.length} line(s))`);
    }
    {
        const lines = buildLines([run('Here is a super', 72, 100, 12, { width: 90 }), run('script', 162, 96.5, 8, { width: 20 })], CFG);
        const text = lines.map(l => l.fragments.map(f => f.text).join('')).join('|');
        const sup = lines[0]?.fragments.some(f => f.formatting.superscript);
        add('Superscript folds into its line', lines.length === 1 && text === 'Here is a superscript' && !!sup,
            'superscript joined and flagged', `${text} (${lines.length} line(s), flagged: ${!!sup})`);
    }
    {
        // A small-font caption under a body line is a line of its own, not a subscript.
        const lines = buildLines([run('Body text line here.', 72, 100, 12, { width: 110 }), run('Figure 1', 72, 108, 8, { width: 34 })], CFG);
        const text = lines.map(l => l.fragments.map(f => f.text).join('')).join('|');
        add('Caption is not glued on as a subscript', lines.length === 2 && text === 'Body text line here.|Figure 1',
            'two separate lines', `${text} (${lines.length} line(s))`);
    }

    // ── H7: line building stays linear in the number of runs ────────────────
    {
        const runs: RawRun[] = [];
        for (let i = 0; i < 40000; i++) runs.push(run(`w${i}`, 72 + (i % 7) * 30, 100 + i * 3, 10));
        const t0 = Date.now();
        const lines = buildLines(runs, CFG);
        const ms = Date.now() - t0;
        // Scanning every cluster for every run took ~3.5 s here (and ~20 s at 100k runs, all of it
        // synchronous, with the abort signal only checked between pages); the bounded scan takes tens
        // of milliseconds. The line count is the one the old code produced, so nothing else changed.
        add('40k runs build in linear time', ms < 2000 && lines.length === 22861, '< 2000 ms, 22861 lines', `${ms} ms, ${lines.length} lines`);
    }

    // ── L-OCR: layout padding is clamped, Blob type follows the bytes ───────
    {
        const page = {
            text: 'A B',
            blocks: [{
                paragraphs: [{
                    lines: [{
                        bbox: { y0: 0 },
                        words: [
                            { text: 'A', bbox: { x0: 0, x1: 8, y0: 0 } },
                            { text: 'B', bbox: { x0: 10_000_000, x1: 10_000_008, y0: 0 } },
                        ],
                    }],
                }],
            }],
        };
        const out = layoutOcrText(page);
        add('OCR column padding is clamped', out.length <= 1100 && out.startsWith('A') && out.endsWith('B'),
            '<= 1100 chars, A ... B', `${out.length} chars`, 'word box 10M px from the left');
    }
    {
        // Two columns whose lines Tesseract splits into separate blocks at the same y must merge onto
        // one visual row (placed by x), not stair-step down the page; a single column stays line-per-line.
        const w = (t: string, x0: number, y0: number) => ({ text: t, bbox: { x0, x1: x0 + t.length * 10, y0 } });
        const ln = (words: any[], y0: number) => ({ bbox: { y0 }, words });
        const multi = { text: '', blocks: [
            { paragraphs: [{ lines: [ln([w('LEFT1', 0, 100)], 100), ln([w('LEFT2', 0, 130)], 130)] }] },
            { paragraphs: [{ lines: [ln([w('RIGHT1', 400, 100)], 100), ln([w('RIGHT2', 400, 130)], 130)] }] },
        ] };
        const rows = layoutOcrText(multi).split('\n');
        const single = { text: '', blocks: [{ paragraphs: [{ lines: [ln([w('one', 0, 100)], 100), ln([w('two', 0, 130)], 130), ln([w('three', 0, 160)], 160)] }] }] };
        const singleRows = layoutOcrText(single).split('\n');
        add('OCR merges same-row columns, keeps single-column lines',
            rows.length === 2 && /LEFT1/.test(rows[0]) && /RIGHT1/.test(rows[0]) && singleRows.length === 3,
            '2 merged rows; 3 single-column lines', `${rows.length} rows / ${singleRows.length} single`);
    }
    {
        const png = sniffImageMime(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]));
        const jpeg = sniffImageMime(Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0, 0, 0, 0]));
        const bmp = sniffImageMime(Buffer.from([0x42, 0x4d, 0, 0, 0, 0, 0, 0]));
        const ok = png === 'image/png' && jpeg === 'image/jpeg' && bmp === 'image/bmp';
        add('Image media type follows the bytes', ok, 'png/jpeg/bmp', `${png}/${jpeg}/${bmp}`);
    }

    return results;
}

/** Allows `npx tsx test/parser/testTextLayout.ts` for a quick standalone run. */
if (process.argv[1] && process.argv[1].endsWith('testTextLayout.ts')) {
    testTextLayout().then(rs => {
        for (const r of rs) console.log(`${r.result.status === 'PASS' ? 'PASS' : 'FAIL'}  ${r.feature}  expected=${JSON.stringify(r.result.expected)} actual=${JSON.stringify(r.result.actual)}`);
        const failed = rs.filter(r => r.result.status !== 'PASS').length;
        console.log(`\n${rs.length - failed}/${rs.length} passed`);
        if (failed) process.exit(1);
    });
}

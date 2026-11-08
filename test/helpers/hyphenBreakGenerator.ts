/**
 * A generator of PDFs whose paragraphs have lines ending in a word divided by a hyphen, used by the test of
 * the hyphen the layout text, PPTX and ODP writers write back at a line break (#145).
 *
 * Each line ends in a hyphenated split, a hyphen drawn as a run of its own (bold), a decomposed accent before
 * the hyphen (`e` and U+0301, from a ToUnicode map), a WinAnsi soft hyphen, a hyphen after a superscript run,
 * a hyphen inside a subscript run, a capital or a digit continuation, or no hyphen; its words mix
 * superscripts, subscripts, a superscript stacked over a subscript, kerned subscripts and stacks, TJ kerning
 * and bold runs. A page holds blocks of one column, two columns read across, a ruled 2 x 2 table, Hebrew
 * (right to left, from a ToUnicode map) and Han (Chinese, set without spaces).
 *
 * Checked per document against the parser's own record and the page:
 *  - the record: the runs marked as having lost a hyphen (droppedHyphen) are as many as the hyphens the
 *    parser took off the runs (the runs' hyphens with mergeHyphenatedWords false, less those with it on);
 *  - layout text, PPTX and ODP with mergeHyphenatedWords on are, line by line, what they are with it off,
 *    the page's own hyphens (whitespace runs collapsed). A line with a hyphen more is stray, one with a
 *    hyphen fewer is missing, any other difference is other. A paragraph a writer flows as one line (no
 *    placed lines) holds the word whole, as flat text does: counted apart, not as a difference.
 *
 * The random numbers come from a full-period LCG mod 2^32 (Math.imul, a = 1664525, c = 1013904223: c odd,
 * a - 1 divisible by 4), so a seed never repeats within 2^32 draws.
 */
import { createHash } from 'crypto';
import { strFromU8, unzipSync } from 'fflate';
import { OfficeParser } from '../../src/OfficeParser';
import { droppedHyphen } from '../../src/utils/lineJoins';
import type { OfficeContentNode, OfficeParserAST } from '../../src/types';

/** The counts and failures of one run. */
export interface HyphenBreakResult {
    docs: number;
    distinctPdfs: number;
    /** Merges the generator wrote: a word hyphen at a line end before a lower-case Latin line. */
    mergesWritten: number;
    /** Runs marked as having lost a hyphen, reachable in the returned ASTs. */
    marks: number;
    /** Hyphens the parser took off runs (mergeHyphenatedWords off against on). */
    hyphensRemoved: number;
    distinctMarkedShapes: number;
    /** Documents whose marks are not the hyphens taken off. */
    lostMarks: number;
    stray: number;
    missing: number;
    other: number;
    flowJoined: number;
    features: Record<string, number>;
    examples: string[];
}

/** A PDF of `pages` content streams. F1 Helvetica, F2 Helvetica-Bold, F3 Hebrew for A to Z, F4 Han for a to z, F5 `e` U+0301 for 0xE9. */
function pdfOf(pages: string[]): Buffer {
    const objs: string[] = [];
    const fontBase = 3 + pages.length * 2;
    objs.push('<< /Type /Catalog /Pages 2 0 R >>');
    objs.push(`<< /Type /Pages /Kids [${pages.map((_, i) => `${3 + i * 2} 0 R`).join(' ')}] /Count ${pages.length} >>`);
    const fonts = `/F1 ${fontBase} 0 R /F2 ${fontBase + 1} 0 R /F3 ${fontBase + 2} 0 R /F4 ${fontBase + 4} 0 R /F5 ${fontBase + 6} 0 R`;
    for (const c of pages) {
        objs.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << ${fonts} >> >> /Contents ${objs.length + 2} 0 R >>`);
        objs.push(`<< /Length ${Buffer.byteLength(c, 'latin1')} >>\nstream\n${c}\nendstream`);
    }
    objs.push('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>');
    objs.push('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>');
    const cmapFont = (entries: string) => {
        const cmap = `/CIDInit /ProcSet findresource begin 12 dict begin begincmap /CMapName /X def 1 begincodespacerange <00> <FF> endcodespacerange ${entries.split('\n').length} beginbfchar ${entries} endbfchar endcmap CMapName currentdict /CMap defineresource pop end end`;
        objs.push(`<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding /ToUnicode ${objs.length + 2} 0 R >>`);
        objs.push(`<< /Length ${cmap.length} >>\nstream\n${cmap}\nendstream`);
    };
    const hex = (n: number, width: number) => n.toString(16).toUpperCase().padStart(width, '0');
    cmapFont(Array.from({ length: 26 }, (_, i) => `<${hex(0x41 + i, 2)}> <${hex(0x05D0 + i, 4)}>`).join('\n'));
    cmapFont(Array.from({ length: 26 }, (_, i) => `<${hex(0x61 + i, 2)}> <${hex(0x4E00 + i * 7, 4)}>`).join('\n'));
    cmapFont('<E9> <00650301>');
    let pdf = '%PDF-1.4\n';
    const offsets = objs.map((o, i) => { const at = Buffer.byteLength(pdf, 'latin1'); pdf += `${i + 1} 0 obj\n${o}\nendobj\n`; return at; });
    const xref = Buffer.byteLength(pdf, 'latin1');
    pdf += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n${offsets.map(o => `${String(o).padStart(10, '0')} 00000 n \n`).join('')}`;
    pdf += `trailer\n<< /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
    return Buffer.from(pdf, 'latin1');
}

const WORDS = ['the', 'committee', 'reached', 'after', 'long', 'debate', 'over', 'budget', 'and', 'members', 'agreed', 'that', 'plan', 'would', 'go', 'ahead', 'with', 'some', 'changes', 'to', 'its', 'scope', 'for', 'next', 'year'];
const SPLITS = [['agree', 'ment'], ['develop', 'ment'], ['commit', 'tee'], ['discus', 'sion'], ['impor', 'tant'], ['infor', 'mation'], ['re', 'sults'], ['meet', 'ing']];
/** Helvetica widths (thousandths) of the letters set raised or lowered. */
const WIDTHS: Record<string, number> = { n: 556, i: 222, a: 556, k: 500, b: 556, d: 556 };
/** The line ends that are a word hyphen the parser joins when the next line starts lower-case. */
const MERGING_ENDS = ['hyphen', 'ownRun', 'decomposed', 'soft', 'supThenHyphen', 'subHyphen'];
type Script = 'latin' | 'heb' | 'han';

/** Layout text, PPTX or ODP as its lines, whitespace runs collapsed and blank lines left out. */
const norm = (t: string) => t.split('\n').map(l => l.replace(/\s+/g, ' ').trim()).filter(Boolean);

async function slidesText(ast: OfficeParserAST, format: 'pptx' | 'odp'): Promise<string> {
    const zip = unzipSync(new Uint8Array((await ast.to(format)).value as unknown as Uint8Array));
    if (format === 'pptx') {
        return Object.keys(zip).filter(k => /^ppt\/slides\/slide\d+\.xml$/.test(k)).sort((a, b) => parseInt(a.slice(16), 10) - parseInt(b.slice(16), 10))
            .map(k => strFromU8(zip[k]).replace(/<a:br[^>]*\/?>(<\/a:br>)?/g, '\n').replace(/<\/a:p>/g, '\n').replace(/<[^>]+>/g, '')).join('\n==\n');
    }
    return strFromU8(zip['content.xml']).replace(/<text:line-break\/>/g, '\n').replace(/<\/text:p>/g, '\n').replace(/<[^>]+>/g, '');
}

/** Runs `count` documents from `seed`. */
export async function runHyphenBreakGenerator(count: number, seed: number): Promise<HyphenBreakResult> {
    let state = seed >>> 0;
    const rnd = (): number => (state = (Math.imul(state, 1664525) + 1013904223) >>> 0) / 4294967296;
    const pick = <T>(xs: readonly T[]): T => xs[Math.floor(rnd() * xs.length)];
    const features = new Map<string, number>();
    const feat = (k: string) => features.set(k, (features.get(k) ?? 0) + 1);

    // One line's text-showing operators (inside BT and ET) and the kind of its end.
    const lineOps = (sz: number, script: Script, narrow: boolean): { ops: string; end: string } => {
        const word = () => script === 'heb' ? Array.from({ length: 3 + Math.floor(rnd() * 4) }, () => String.fromCharCode(65 + Math.floor(rnd() * 26))).join('')
            : script === 'han' ? Array.from({ length: 3 + Math.floor(rnd() * 4) }, () => String.fromCharCode(97 + Math.floor(rnd() * 26))).join('') : pick(WORDS);
        const font = script === 'heb' ? 'F3' : script === 'han' ? 'F4' : 'F1';
        let ops = `/${font} ${sz} Tf `;
        const words = narrow ? 1 + Math.floor(rnd() * 3) : 3 + Math.floor(rnd() * 5);
        for (let k = 0; k < words; k++) {
            const r = rnd();
            if (script === 'latin' && r < 0.12) { ops += `/F2 ${sz} Tf (${word()} ) Tj /${font} ${sz} Tf `; feat('bold'); continue; }
            if (script === 'latin' && r < 0.30) {
                // x with a superscript, a subscript, or both stacked (the subscript under the superscript), kerned or not
                const sup = pick(Object.keys(WIDTHS)), sub = pick(Object.keys(WIDTHS)), small = Math.round(sz * 0.7);
                const kind = pick(['sup', 'sub', 'stack', 'subKern', 'stackKern']);
                feat(kind);
                ops += `(${word()} x) Tj `;
                if (kind === 'subKern' || kind === 'stackKern') ops += '[40] TJ ';
                if (kind === 'sup' || kind === 'stack' || kind === 'stackKern') ops += `/F1 ${small} Tf ${(sz * 0.35).toFixed(1)} Ts (${sup}) Tj `;
                if (kind === 'stack' || kind === 'stackKern') ops += `[${WIDTHS[sup]}] TJ `;
                if (kind !== 'sup') ops += `/F1 ${small} Tf ${(-sz * 0.25).toFixed(1)} Ts (${sub}) Tj `;
                ops += `/${font} ${sz} Tf 0 Ts ( ) Tj `;
                continue;
            }
            if (r < 0.40) { const w = word(); ops += `[(${w.slice(0, 2)}) ${pick([-30, 20, 60])} (${w.slice(2)} )] TJ `; feat('kernTJ'); continue; }
            ops += `(${word()} ) Tj `;
        }
        const end = pick(script === 'latin'
            ? ['hyphen', 'hyphen', 'hyphen', 'ownRun', 'decomposed', 'soft', 'capital', 'digit', 'supThenHyphen', 'subHyphen', 'none', 'none']
            : ['hyphen', 'hyphen', 'none']);
        const [pre] = pick(SPLITS);
        const before = script === 'latin' ? pre : word();
        const small = Math.round(sz * 0.7);
        switch (end) {
            case 'hyphen': case 'capital': ops += `(${before}-) Tj `; break;
            case 'ownRun': ops += `(${before}) Tj /F2 ${sz} Tf (-) Tj `; break;
            case 'decomposed': ops += `/F5 ${sz} Tf (caf\\351-) Tj `; break;
            case 'soft': ops += `(${before}\\255) Tj `; break;
            case 'digit': ops += '(COVID-19-) Tj '; break;
            case 'supThenHyphen': ops += `(${before}) Tj /F1 ${small} Tf ${(sz * 0.35).toFixed(1)} Ts (n) Tj /F1 ${sz} Tf 0 Ts (-) Tj `; break;
            case 'subHyphen': ops += `(${before}) Tj /F1 ${small} Tf ${(-sz * 0.25).toFixed(1)} Ts (i-) Tj /F1 ${sz} Tf 0 Ts `; break;
            default: ops += `(${word()}) Tj `;
        }
        feat(`end:${end}${script === 'latin' ? '' : `:${script}`}`);
        return { ops, end };
    };

    // A paragraph of `lines` lines at (x, y); returns its operators, the merges it holds and where it ends.
    const paragraph = (x: number, y: number, sz: number, lines: number, script: Script, narrow = false): { ops: string; merges: number; bottom: number } => {
        let ops = '', merges = 0, nextStart = '';
        const lead = Math.round(sz * 1.25 * 10) / 10;
        for (let l = 0; l < lines; l++) {
            const nextScript = script === 'latin' ? 'latin' : (rnd() < 0.3 ? 'latin' : script);
            const { ops: lineOpsText, end } = lineOps(sz, script, narrow);
            ops += `BT ${x} ${(y - l * lead).toFixed(1)} Td ${nextStart}${lineOpsText}ET\n`;
            nextStart = '';
            if (l === lines - 1) continue;
            const [, suffix] = pick(SPLITS);
            if (end === 'capital') nextStart = `/F1 ${sz} Tf (Known ) Tj `;
            else if (end === 'none' || end === 'digit') nextStart = `/F1 ${sz} Tf (${pick(WORDS)} ) Tj `;
            else if (nextScript === 'latin') nextStart = `/F1 ${sz} Tf (${suffix} ) Tj `;
            else nextStart = `/${script === 'heb' ? 'F3' : 'F4'} ${sz} Tf (${script === 'heb' ? 'ABC' : 'abc'} ) Tj `;
            if (MERGING_ENDS.includes(end) && nextScript === 'latin') merges++;
        }
        return { ops, merges, bottom: y - lines * lead };
    };

    const parse = (bytes: Buffer, mergeHyphenatedWords: boolean) => OfficeParser.parseOffice(bytes, {
        fileType: 'pdf', ocr: false, onWarning: () => {}, pdfParserConfig: { separateProcess: false, mergeHyphenatedWords },
    } as any) as Promise<OfficeParserAST>;
    const nodesOf = (ast: OfficeParserAST): OfficeContentNode[] => {
        const all: OfficeContentNode[] = [];
        const walk = (n: OfficeContentNode) => { all.push(n); for (const c of n.children ?? []) walk(c); };
        ast.content.forEach(walk);
        return all;
    };
    const runHyphens = (nodes: OfficeContentNode[]) => nodes.reduce((sum, n) => sum + (n.type === 'text' && !n.children?.length ? (n.text ?? '').match(/[-\u00AD]/g)?.length ?? 0 : 0), 0);

    const result: HyphenBreakResult = { docs: 0, distinctPdfs: 0, mergesWritten: 0, marks: 0, hyphensRemoved: 0, distinctMarkedShapes: 0, lostMarks: 0, stray: 0, missing: 0, other: 0, flowJoined: 0, features: {}, examples: [] };
    const pdfs = new Set<string>(), shapes = new Set<string>();
    for (let d = 0; d < count; d++) {
        const pages: string[] = [];
        let merges = 0;
        const pageCount = rnd() < 0.2 ? 2 : 1;
        for (let p = 0; p < pageCount; p++) {
            let c = '', y = 740;
            const blocks = 2 + Math.floor(rnd() * 3);
            for (let b = 0; b < blocks && y > 200; b++) {
                const kind = pick(['single', 'single', 'columns', 'table', 'heb', 'han'] as const);
                feat(`block:${kind}`);
                const sz = pick([9, 10, 11, 12]);
                if (kind === 'single' || kind === 'heb' || kind === 'han') {
                    const para = paragraph(kind === 'heb' ? 200 : 72, y, sz, 2 + Math.floor(rnd() * 4), kind === 'single' ? 'latin' : kind);
                    c += para.ops; merges += para.merges; y = para.bottom - 2.5 * sz;
                } else if (kind === 'columns') {
                    const lines = 2 + Math.floor(rnd() * 4);
                    const left = paragraph(72, y, Math.min(sz, 10), lines, 'latin', true), right = paragraph(330, y, Math.min(sz, 10), lines, 'latin', true);
                    c += left.ops + right.ops; merges += left.merges + right.merges; y = Math.min(left.bottom, right.bottom) - 2.5 * sz;
                } else {
                    const top = y, h = 4 * 9 * 1.25 + 8;
                    c += `0.5 w 72 ${top - h} 250 ${h} re S 322 ${top - h} 250 ${h} re S 72 ${top - 2 * h} 250 ${h} re S 322 ${top - 2 * h} 250 ${h} re S\n`;
                    for (const [cx, cy] of [[76, top - 12], [326, top - 12], [76, top - h - 12], [326, top - h - 12]]) {
                        const para = paragraph(cx, cy, 7, 2, 'latin', true);
                        c += para.ops; merges += para.merges;
                    }
                    y = top - 2 * h - 30;
                }
            }
            pages.push(c);
        }
        const bytes = pdfOf(pages);
        pdfs.add(createHash('sha1').update(bytes).digest('hex'));
        result.docs++;
        result.mergesWritten += merges;

        const [on, off] = [await parse(bytes, true), await parse(bytes, false)];
        const onNodes = nodesOf(on);
        const marked = onNodes.filter(n => droppedHyphen(n));
        for (const n of marked) shapes.add(JSON.stringify([n.text ? 'run' : 'emptied', !!n.formatting?.superscript, !!n.formatting?.subscript, !!n.formatting?.bold]));
        const removed = runHyphens(nodesOf(off)) - runHyphens(onNodes);
        result.marks += marked.length;
        result.hyphensRemoved += removed;
        if (marked.length !== removed) { result.lostMarks++; if (result.examples.length < 6) result.examples.push(`seed ${seed} doc ${d}: ${marked.length} marks, ${removed} hyphens taken off`); }

        for (const format of ['text', 'pptx', 'odp'] as const) {
            const written = async (ast: OfficeParserAST) => norm(format === 'text' ? String((await ast.to('text')).value) : await slidesText(ast, format));
            const [a, b] = [await written(on), await written(off)];
            if (a.length !== b.length) { result.other++; if (result.examples.length < 6) result.examples.push(`seed ${seed} doc ${d} ${format}: ${a.length} lines, ${b.length} with mergeHyphenatedWords false`); continue; }
            for (let i = 0; i < a.length; i++) {
                if (a[i] === b[i]) continue;
                // A paragraph written flowing, as one line, holds the word whole, as flat text does.
                if (a[i] === b[i].replace(/(\p{L}\p{M}*)[-\u00AD] (?=\p{Ll})/gu, '$1')) { result.flowJoined++; continue; }
                const [ha, hb] = [a[i].match(/-/g)?.length ?? 0, b[i].match(/-/g)?.length ?? 0];
                const same = a[i].replace(/-/g, '') === b[i].replace(/-/g, '');
                if (same && ha > hb) result.stray++;
                else if (same && ha < hb) result.missing++;
                else result.other++;
                if (result.examples.length < 6) result.examples.push(`seed ${seed} doc ${d} ${format}: ${JSON.stringify(a[i])} against ${JSON.stringify(b[i])}`);
            }
        }
    }
    result.distinctPdfs = pdfs.size;
    result.distinctMarkedShapes = shapes.size;
    result.features = Object.fromEntries([...features].sort(([a], [b]) => a.localeCompare(b)));
    return result;
}

/**
 * Comprehensive Office Generator Test Suite
 *
 * Tests all supported output formats:
 * - Feature validation (structure, content, formatting)
 *
 * Known limitations reflected as WARN (not FAIL):
 * - xlsx/ods → csv: CsvGenerator returns a ZIP (Uint8Array) for multi-sheet files.
 *   The suite detects this and parses accordingly.
 * - xlsx/ods → chunks: Spreadsheet ASTs may produce 0 chunks (ChunkingGenerator
 *   currently focuses on paragraph/heading content, not raw cell values).
 * - rtf → rtf roundtrip: The RTF parser re-parsing a generated RTF may lose
 *   heading/list structure due to inherent RTF formatting heuristics.
 * - Config permutation testing
 * - Baseline snapshot comparison
 * - Roundtrip parity: same-format parse→generate→parse must have zero data loss
 */

import * as fs from 'fs';
import * as path from 'path';
import { strFromU8, unzipSync } from 'fflate';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParser } from '../../src/OfficeParser';
import {
    DeepRequired,
    GeneratorConfig,
    OfficeContentNode,
    OfficeContentNodeType,
    OfficeParserAST,
    OfficeWarningType,
    OfficeParserConfig,
} from '../../src/types';
import { resolveGeneratorConfig } from '../../src/utils/configUtils.js';
import { createAST } from '../../src/utils/astUtils.js';
import { isHeaderRow } from '../../src/utils/officeGenUtils.js';

// ============================================================================
// CONSTANTS & CONFIGURATION
// ============================================================================

/** Output formats tested (PDF excluded — slow/brittle in CI) */
const GENERATOR_FORMATS = ['html', 'md', 'text', 'rtf', 'csv', 'chunks', 'epub', 'docx', 'odt', 'tex'] as const;
type GeneratorFormat = typeof GENERATOR_FORMATS[number];

/** Formats that support roundtrip testing (parse → generate → re-parse) */
const ROUNDTRIP_FORMATS: GeneratorFormat[] = ['html', 'md', 'rtf', 'docx', 'odt', 'tex'];

/** Source baseline formats used to drive generation tests */
const SOURCE_FORMATS = {
    documents: ['docx', 'odt', 'pptx', 'odp', 'odg', 'pdf', 'rtf', 'html', 'md', 'epub', 'tex'] as const,
    spreadsheets: ['xlsx', 'ods', 'csv'] as const,
};

/**
 * Formats fast mode leaves out - see the matching block in `test/parser/testOfficeParser.ts` for
 * why PDF is the one that matters. The saving is smaller here (this suite excludes PDF *output*
 * already, so only the source-side parse is paid), but keeping the two suites on one flag means an
 * agent has a single command to remember rather than a rule about which suite takes the argument.
 */
const FAST_MODE_SKIPPED_FORMATS = ['pdf'];

/** True when invoked as `... fast`. Never applies to `baseline` - see the parser suite's note. */
const FAST_MODE = (process.argv[2] || '').toLowerCase() === 'fast';

/** Drops the fast-mode-skipped formats from a list of extensions when fast mode is on. */
const activeFormats = <T extends string>(formats: readonly T[]): T[] =>
    FAST_MODE ? formats.filter(f => !FAST_MODE_SKIPPED_FORMATS.includes(f)) : [...formats];

/** Parser config used to parse source files for generation */
const PARSER_CONFIG: DeepRequired<OfficeParserConfig> = {
    extractAttachments: true,
    ocr: false,
    ocrConfig: { language: 'eng', workerPath: '', corePath: '', langPath: '', preserveLayout: true, timeout: { autoTerminate: 10000, workerLoad: 60000, recognition: 30000 }, abortSignal: null },
    includeRawContent: false,
    ignoreNotes: false,
    newlineDelimiter: '\n',
    pdfWorkerSrc: '',
    serializeRawContent: true,
    preserveXmlWhitespace: false,
    includeBreakNodes: false,
    ignoreInternalLinks: false,
    csvDelimiter: ',',
    onWarning: () => { },
    fileType: null,
    ignoreComments: false,
    ignoreHeadersAndFooters: false,
    ignoreSlideMasters: false,
    abortSignal: null,
    decompressionLimits: {
        maxUncompressedBytes: 512 * 1024 * 1024,
        maxZipEntries: 10000,
        maxTableCells: 1000000
    },
    htmlParserConfig: { preserveAttributes: false, preserveIframes: false, preserveComments: true, embedFolkForms: false },
    ignorePageGeometry: false,
    password: '',
    onPassword: () => undefined,
    pdfParserConfig: {
        useTags: true,
        detectColumns: true,
        mergeHyphenatedWords: true,
        lineToleranceFactor: 0.35,
        spaceToleranceFactor: 0.25,
        headingDetection: 'auto',
        pageRange: '',
        normalizeText: true,
        extractTextColor: true
    },
    texParserConfig: { today: '' },
};

/**
 * Fixed modification instant for every generation in this suite.
 *
 * EPUB embeds a timestamp in two places - `dcterms:modified` in the OPF, and each zip entry's
 * mtime - and both default to the wall clock. Left unpinned, every `gen.*.to.epub.epub` baseline
 * differs on each run even when nothing changed, so the real signal (a content regression) is
 * indistinguishable from the noise and gets reverted along with it. Pinning makes the artifact
 * reproducible, which is what lets a byte diff of these baselines mean something.
 */
const BASELINE_MODIFIED = new Date('2024-01-01T00:00:00Z');

/**
 * Applied to the zip-packaged generators (EPUB, DOCX, ODT). Other formats embed
 * `ast.metadata.modified` straight from the fixture, which is already stable run to run, so pinning
 * them too would rewrite their baselines with a synthetic date and make the snapshots less faithful
 * to real output for no gain. The packaged formats are the ones that need it: fixtures without a
 * modification date of their own (CSV, HTML) otherwise fall through to the current time, and the zip
 * entry mtimes follow the same instant.
 */
const deterministicConfigFor = (destFmt: string): GeneratorConfig =>
    (destFmt === 'epub' || destFmt === 'docx' || destFmt === 'odt' ? { metadataOverrides: { modified: BASELINE_MODIFIED } } : {}) as GeneratorConfig;

/** Generator config permutations */
const GENERATOR_CONFIG_TESTS = [
    {
        id: 'G1',
        name: 'Defaults',
        config: {} as GeneratorConfig,
    },
    {
        id: 'G2',
        name: 'No Formatting',
        config: { includeFormatting: false } as GeneratorConfig,
    },
    {
        id: 'G3',
        name: 'With Metadata',
        config: { renderMetadata: true } as GeneratorConfig,
    },
    {
        id: 'G4',
        name: 'No Images',
        config: { includeImages: false } as GeneratorConfig,
    },
    {
        id: 'G5',
        name: 'No IDs',
        config: { generateIds: false } as GeneratorConfig,
    },
    {
        id: 'G6',
        name: 'HTML Non-Standalone',
        config: { htmlConfig: { standalone: false } } as GeneratorConfig,
        formats: ['html'] as GeneratorFormat[],
    },
    {
        id: 'G7',
        name: 'MD No HTML Fallback',
        config: { mdConfig: { fallbackToHtml: false } } as GeneratorConfig,
        formats: ['md'] as GeneratorFormat[],
    },
    {
        id: 'G8',
        name: 'CSV Merge Sheets',
        config: { csvConfig: { mergeSheets: true } } as GeneratorConfig,
        formats: ['csv'] as GeneratorFormat[],
    },
    {
        id: 'G9',
        name: 'CSV Custom Delimiter',
        config: { csvConfig: { columnDelimiter: ';' } } as GeneratorConfig,
        formats: ['csv'] as GeneratorFormat[],
    },
    {
        id: 'G10',
        name: 'Text Preserve Layout',
        config: { textConfig: { preserveLayout: true } } as GeneratorConfig,
        formats: ['text'] as GeneratorFormat[],
    },
    {
        id: 'G11',
        name: 'Chunks Fixed-Size',
        config: { chunksConfig: { strategy: 'fixed-size', chunkSize: 500, chunkOverlap: 50 } } as GeneratorConfig,
        formats: ['chunks'] as GeneratorFormat[],
    },
    {
        id: 'G12',
        name: 'Chunks Document-Structure',
        config: { chunksConfig: { strategy: 'document-structure', splitBy: 'heading', maxChunkSize: 800 } } as GeneratorConfig,
        formats: ['chunks'] as GeneratorFormat[],
    },
    {
        id: 'G13',
        name: 'HTML Custom Width String',
        config: { htmlConfig: { containerWidth: '950px' } } as GeneratorConfig,
        formats: ['html'] as GeneratorFormat[],
    },
    {
        id: 'G14',
        name: 'HTML Custom Width Number',
        config: { htmlConfig: { containerWidth: 1000 } } as GeneratorConfig,
        formats: ['html'] as GeneratorFormat[],
    },
    {
        id: 'G15',
        name: 'HTML Standalone Object (document:false only)',
        config: { htmlConfig: { standalone: { document: false } } } as GeneratorConfig,
        formats: ['html'] as GeneratorFormat[],
    },
    {
        id: 'G16',
        name: 'HTML Standalone Object (document:false, styles:scoped)',
        config: { htmlConfig: { standalone: { document: false, styles: 'scoped' } } } as GeneratorConfig,
        formats: ['html'] as GeneratorFormat[],
    },
    {
        id: 'G17',
        name: 'HTML Standalone Object (document:false, bodyInjections:true)',
        config: {
            htmlConfig: {
                standalone: { document: false, bodyInjections: true },
                injections: { bodyStart: '<!--OP_BODY_START_TEST-->', bodyEnd: '<!--OP_BODY_END_TEST-->' }
            }
        } as GeneratorConfig,
        formats: ['html'] as GeneratorFormat[],
    },
    {
        id: 'G18',
        name: 'HTML Standalone true (explicit)',
        config: { htmlConfig: { standalone: true } } as GeneratorConfig,
        formats: ['html'] as GeneratorFormat[],
    },
    {
        id: 'G-docx',
        name: 'DOCX Legal Landscape Margins',
        config: { docxConfig: { format: 'Legal', landscape: true, margin: { top: 36, right: 18, bottom: 36, left: 18 } } } as GeneratorConfig,
        formats: ['docx'] as GeneratorFormat[],
    },
    {
        id: 'G-odt',
        name: 'ODT Letter Landscape Margins',
        config: { odtConfig: { format: 'Letter', landscape: true, margin: { top: 36, right: 18, bottom: 36, left: 18 } } } as GeneratorConfig,
        formats: ['odt'] as GeneratorFormat[],
    },
    {
        id: 'G-tex',
        name: 'LaTeX report class, numbered sections, Letter landscape margins',
        config: { texConfig: { documentClass: 'report', numberSections: true, format: 'Letter', landscape: true, margin: { top: 36, right: 18, bottom: 36, left: 18 } } } as GeneratorConfig,
        formats: ['tex'] as GeneratorFormat[],
    },
    {
        id: 'G-tex-fragment',
        name: 'LaTeX body-only fragment',
        config: { texConfig: { standalone: false } } as GeneratorConfig,
        formats: ['tex'] as GeneratorFormat[],
    },
    {
        id: 'G-tex-bundle',
        name: 'LaTeX zip bundle (main.tex + images)',
        config: { texConfig: { bundle: true }, metadataOverrides: { modified: new Date('2024-01-01T00:00:00Z') } } as GeneratorConfig,
        formats: ['tex'] as GeneratorFormat[],
    },
];

/**
 * The LaTeX body with the parts TeX does not interpret as commands removed: verbatim/lstlisting
 * environments (their content is literal) and `%` comments (an unescaped `%` to the end of the
 * line). What remains is what a structural check must hold for.
 */
function liveLatex(tex: string): string {
    return tex
        .replace(/\\begin\{(verbatim|lstlisting)\}[\s\S]*?\\end\{\1\}/g, '')
        .replace(/(^|[^\\])((?:\\\\)*)%.*$/gm, '$1$2');
}

// ============================================================================
// TYPES
// ============================================================================

interface GenTestResult {
    status: 'PASS' | 'FAIL' | 'WARN' | 'SKIP';
    expected: any;
    actual: any;
    details: string;
    duration?: number;
}

interface GenFeatureTest {
    category: string;
    feature: string;
    sourceFormat: string;
    destFormat: string;
    result: GenTestResult;
}

interface GeneratedMetrics {
    outputLength: number;
    hasContent: boolean;
    lineCount: number;
    wordCount: number;
    // HTML/MD specific
    tagCounts?: Record<string, number>;
    headingCount?: number;
    tableCount?: number;
    listCount?: number;
    imageCount?: number;
    linkCount?: number;
    // Chunk specific
    chunkCount?: number;
    avgChunkLength?: number;
    // CSV specific
    rowCount?: number;
}

interface RoundtripMetrics {
    contentNodes: number;
    textLength: number;
    headings: number;
    tables: number;
    lists: number;
    images: number;
    // Fidelity dimensions. The metrics above count node *types* only, so a change to
    // header-cell/embed/highlight/caption fidelity moves none of them - which is exactly the
    // blind spot this harness exists to prevent. These make that class of regression visible.
    headerCells: number;
    embeds: number;
    highlightedText: number;
    captions: number;
}

// ============================================================================
// PATHS
// ============================================================================

function getSourcePath(ext: string): string {
    return path.join(__dirname, '..', 'files', `test.${ext}`);
}

function getBaselinePath(ext: string): string {
    return path.join(__dirname, '..', 'baseline', `test.${ext}.json`);
}

function getGeneratorBaselinePath(srcExt: string, destFmt: string): string {
    return path.join(__dirname, 'baseline', `gen.${srcExt}.to.${destFmt}.metrics.json`);
}

function getGeneratorOutputPath(srcExt: string, destFmt: string): string {
    return path.join(__dirname, 'output', `gen.${srcExt}.to.${destFmt}`);
}

function loadParserBaseline(ext: string): OfficeParserAST | null {
    const p = getBaselinePath(ext);
    return fs.existsSync(p) ? JSON.parse(fs.readFileSync(p, 'utf8')) : null;
}

function loadGeneratorBaseline(srcExt: string, destFmt: string): GeneratedMetrics | null {
    const p = getGeneratorBaselinePath(srcExt, destFmt);
    return fs.existsSync(p) ? JSON.parse(fs.readFileSync(p, 'utf8')) : null;
}

// ============================================================================
// METRICS EXTRACTION
// ============================================================================

async function extractRoundtripMetrics(ast: OfficeParserAST): Promise<RoundtripMetrics> {
    let contentNodes = 0, headings = 0, tables = 0, lists = 0, images = 0;
    let headerCells = 0, embeds = 0, highlightedText = 0, captions = 0;

    function traverse(nodes: OfficeContentNode[]) {
        for (const n of nodes) {
            contentNodes++;
            if (n.type === 'heading') headings++;
            if (n.type === 'table' || n.type === 'sheet') tables++;
            if (n.type === 'list') lists++;
            if (n.type === 'image') images++;
            if (n.type === 'embed') embeds++;
            // Tri-state: only an explicit `true` counts, so an absent flag never inflates this.
            if (n.type === 'cell' && (n.metadata as any)?.isHeader === true) headerCells++;
            // Cast: `highlight` lands on TextFormatting in a later commit; this metric is
            // introduced first so that commit's effect is visible in the diff rather than silent.
            if ((n.formatting as any)?.highlight === true) highlightedText++;
            if ((n.type === 'image' || n.type === 'table') && (n.metadata as any)?.caption) captions++;
            if (n.children) traverse(n.children);
        }
    }
    traverse(ast.content);

    const textLength = (await ast.to('text')).value.length;
    return { contentNodes, textLength, headings, tables, lists, images, headerCells, embeds, highlightedText, captions };
}

/** Metrics for a generated DOCX (binary zip): unzips and reads word/document.xml. */
function extractDocxMetrics(bytes: Uint8Array): GeneratedMetrics {
    const files = unzipSync(bytes);
    const doc = files['word/document.xml'] ? strFromU8(files['word/document.xml']) : '';
    const text = (doc.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g) || []).map(m => m.replace(/<[^>]+>/g, '')).join(' ');
    const words = text.split(/\s+/).filter(Boolean);
    return {
        outputLength: bytes.length,
        hasContent: bytes.length > 0 && words.length > 0,
        lineCount: (doc.match(/<w:p[ >/]/g) || []).length,
        wordCount: words.length,
        headingCount: (doc.match(/w:val="Heading/g) || []).length,
        tableCount: (doc.match(/<w:tbl[ >]/g) || []).length,
        listCount: (doc.match(/<w:numPr/g) || []).length,
        imageCount: (doc.match(/<w:drawing/g) || []).length,
        linkCount: (doc.match(/<w:hyperlink/g) || []).length,
    };
}

/** Metrics for a generated ODT (binary zip): unzips and reads content.xml. */
function extractOdtMetrics(bytes: Uint8Array): GeneratedMetrics {
    const files = unzipSync(bytes);
    const doc = files['content.xml'] ? strFromU8(files['content.xml']) : '';
    const text = doc.replace(/<[^>]+>/g, ' ');
    const words = text.split(/\s+/).filter(Boolean);
    return {
        outputLength: bytes.length,
        hasContent: bytes.length > 0 && words.length > 0,
        lineCount: (doc.match(/<text:p[ >/]/g) || []).length,
        wordCount: words.length,
        headingCount: (doc.match(/<text:h[ >]/g) || []).length,
        tableCount: (doc.match(/<table:table[ >]/g) || []).length,
        listCount: (doc.match(/<text:list-item[ >]/g) || []).length,
        imageCount: (doc.match(/<draw:image[ >/]/g) || []).length,
        linkCount: (doc.match(/<text:a[ >]/g) || []).length,
    };
}

function extractGeneratedMetrics(output: string, fmt: GeneratorFormat): GeneratedMetrics {
    const outputLength = output.length;
    const hasContent = outputLength > 0;
    const lineCount = output.split('\n').length;
    const wordCount = output.split(/\s+/).filter(w => w.length > 0).length;

    const metrics: GeneratedMetrics = { outputLength, hasContent, lineCount, wordCount };

    if (fmt === 'html') {
        const countTag = (tag: string) => (output.match(new RegExp(`<${tag}[\\s>]`, 'gi')) || []).length;
        metrics.headingCount = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'].reduce((s, t) => s + countTag(t), 0);
        metrics.tableCount = countTag('table');
        metrics.listCount = countTag('ul') + countTag('ol');
        metrics.imageCount = countTag('img');
        metrics.linkCount = countTag('a');
        metrics.tagCounts = { h1: countTag('h1'), h2: countTag('h2'), h3: countTag('h3'), table: countTag('table'), ul: countTag('ul'), ol: countTag('ol') };
    }

    if (fmt === 'md') {
        metrics.headingCount = (output.match(/^#{1,6} /gm) || []).length;
        metrics.tableCount = (output.match(/^\|.*\|.*\|/gm) || []).length > 0 ? 1 : 0;
        metrics.listCount = (output.match(/^[-*+] /gm) || []).length + (output.match(/^\d+\. /gm) || []).length;
        metrics.imageCount = (output.match(/!\[/g) || []).length;
        metrics.linkCount = (output.match(/\[.*?\]\(/g) || []).length;
    }

    if (fmt === 'csv') {
        const lines = output.split('\n').filter(l => l.trim() && !l.startsWith('#'));
        metrics.rowCount = lines.length;
    }

    if (fmt === 'tex') {
        const live = liveLatex(output);
        const count = (re: RegExp) => (live.match(re) || []).length;
        metrics.headingCount = count(/\\(?:chapter|section|subsection|subsubsection|paragraph|subparagraph)(?:\[|\{)/g) + count(/\\frametitle\{/g);
        metrics.tableCount = count(/\\begin\{(?:longtable|tabular)\}/g);
        metrics.listCount = count(/\\begin\{(?:itemize|enumerate|description)\}/g);
        metrics.imageCount = count(/\\includegraphics\[/g);
        metrics.linkCount = count(/\\(?:href|hyperref)[\[{]/g);
    }

    return metrics;
}

function extractChunkMetrics(chunks: any[]): GeneratedMetrics {
    const chunkCount = chunks.length;
    const totalLen = chunks.reduce((s, c) => s + (c.text?.length || 0), 0);
    const avgChunkLength = chunkCount > 0 ? Math.round(totalLen / chunkCount) : 0;
    return {
        outputLength: totalLen,
        hasContent: chunkCount > 0,
        lineCount: chunkCount,
        wordCount: chunks.reduce((s, c) => s + (c.text?.split(/\s+/).length || 0), 0),
        chunkCount,
        avgChunkLength,
    };
}

// ============================================================================
// COMPARISON HELPERS
// ============================================================================

function compareGeneratedMetrics(
    category: string,
    srcFmt: string,
    destFmt: string,
    expected: GeneratedMetrics,
    actual: GeneratedMetrics
): GenFeatureTest[] {
    const results: GenFeatureTest[] = [];

    const mk = (feature: string, exp: any, act: any, ok: boolean, details: string, warn = false): GenFeatureTest => ({
        category, feature, sourceFormat: srcFmt, destFormat: destFmt,
        result: { status: ok ? 'PASS' : (warn ? 'WARN' : 'FAIL'), expected: exp, actual: act, details }
    });

    // Must have content
    results.push(mk('Has Content', true, actual.hasContent, actual.hasContent, actual.hasContent ? 'Output is non-empty' : 'Output is empty!'));

    // Output length within 20% of baseline
    const lenRatio = expected.outputLength > 0 ? actual.outputLength / expected.outputLength : 1;
    const lenOk = lenRatio >= 0.80 && lenRatio <= 1.20;
    results.push(mk('Output Length', `${expected.outputLength} (±20%)`, actual.outputLength, lenOk,
        `${(lenRatio * 100).toFixed(1)}% of baseline length`, !lenOk));

    // Word count within 15%
    const wRatio = expected.wordCount > 0 ? actual.wordCount / expected.wordCount : 1;
    const wOk = wRatio >= 0.85 && wRatio <= 1.15;
    results.push(mk('Word Count', `${expected.wordCount} (±15%)`, actual.wordCount, wOk,
        `${(wRatio * 100).toFixed(1)}% of baseline word count`, !wOk));

    if (destFmt === 'html' || destFmt === 'md' || destFmt === 'docx' || destFmt === 'odt' || destFmt === 'tex') {
        if (expected.headingCount !== undefined && expected.headingCount > 0) {
            const hOk = actual.headingCount === expected.headingCount;
            results.push(mk('Headings', expected.headingCount, actual.headingCount, hOk,
                hOk ? `${actual.headingCount} headings` : `Expected ${expected.headingCount}, got ${actual.headingCount}`));
        }
        if (expected.tableCount !== undefined && expected.tableCount > 0) {
            const tOk = actual.tableCount === expected.tableCount;
            results.push(mk('Tables', expected.tableCount, actual.tableCount, tOk,
                tOk ? `${actual.tableCount} tables` : `Expected ${expected.tableCount}, got ${actual.tableCount}`));
        }
        if (expected.imageCount !== undefined) {
            const iOk = actual.imageCount === expected.imageCount;
            results.push(mk('Images', expected.imageCount, actual.imageCount, iOk,
                iOk ? `${actual.imageCount} images` : `Expected ${expected.imageCount}, got ${actual.imageCount}`, !iOk));
        }
    }

    if (destFmt === 'chunks') {
        if (expected.chunkCount !== undefined && expected.chunkCount > 0) {
            const cRatio = actual.chunkCount! / expected.chunkCount;
            const cOk = cRatio >= 0.85 && cRatio <= 1.15;
            results.push(mk('Chunk Count', `${expected.chunkCount} (±15%)`, actual.chunkCount, cOk,
                `${actual.chunkCount} chunks (${(cRatio * 100).toFixed(1)}% of baseline)`, !cOk));
        }
    }

    return results;
}

/**
 * Whether a roundtrip format has known structural re-parse limitations.
 * For these formats, structural checks (headings/lists/tables) are WARN not FAIL.
 */
const ROUNDTRIP_STRUCTURAL_WARN: Record<string, boolean> = {
    rtf: true,  // RTF re-parsing loses heading/list structure (heuristic detection)
};

function compareRoundtripASTs(
    srcFmt: string,
    destFmt: string,
    original: RoundtripMetrics,
    regenerated: RoundtripMetrics
): GenFeatureTest[] {
    const results: GenFeatureTest[] = [];
    const category = `Roundtrip (${srcFmt}→${destFmt}→${srcFmt})`;
    const structuralWarn = ROUNDTRIP_STRUCTURAL_WARN[srcFmt] ?? false;

    const mk = (feature: string, exp: any, act: any, ok: boolean, details: string, warn = false): GenFeatureTest => ({
        category, feature, sourceFormat: srcFmt, destFormat: destFmt,
        result: { status: ok ? 'PASS' : (warn ? 'WARN' : 'FAIL'), expected: exp, actual: act, details }
    });

    // Text length within 5% (cosmetic whitespace variance allowed; 10% for RTF to allow footnote body conversion)
    const tLimit = srcFmt === 'rtf' ? 0.10 : 0.05;
    const tRatio = original.textLength > 0 ? regenerated.textLength / original.textLength : 1;
    const tOk = tRatio >= (1 - tLimit) && tRatio <= (1 + tLimit);
    results.push(mk(`Text Length (±${tLimit * 100}%)`, `${original.textLength}`, regenerated.textLength,
        tOk, `${(tRatio * 100).toFixed(1)}% of original text length`));

    // Headings — WARN for RTF (re-parser uses heuristics)
    const hOk = original.headings === regenerated.headings;
    results.push(mk('Headings (exact)', original.headings, regenerated.headings,
        hOk, hOk ? `${regenerated.headings} headings` : `Expected ${original.headings}, got ${regenerated.headings}`,
        !hOk && structuralWarn));

    // Tables — WARN for RTF
    const tableOk = original.tables === regenerated.tables;
    results.push(mk('Tables (exact)', original.tables, regenerated.tables,
        tableOk, tableOk ? `${regenerated.tables} tables` : `Expected ${original.tables}, got ${regenerated.tables}`,
        !tableOk && structuralWarn));

    // Lists within ±1 — WARN for RTF
    const listDiff = Math.abs(original.lists - regenerated.lists);
    const listOk = listDiff <= 1;
    results.push(mk('Lists (±1)', original.lists, regenerated.lists,
        listOk, listOk ? `${regenerated.lists} lists` : `Expected ${original.lists}, got ${regenerated.lists}`,
        !listOk && (structuralWarn || listDiff <= 3)));

    // Content nodes within 10%
    const nRatio = original.contentNodes > 0 ? regenerated.contentNodes / original.contentNodes : 1;
    const nOk = nRatio >= 0.90 && nRatio <= 1.10;
    results.push(mk('Content Nodes (±10%)', original.contentNodes, regenerated.contentNodes,
        nOk, `${(nRatio * 100).toFixed(1)}% of original nodes`, !nOk));

    // --- Fidelity dimensions ---
    // These are the constructs a "round-trips fine" node-type count can't see. Each is compared
    // exactly, but only WARNs when it drops: several are legitimately unrepresentable in a given
    // target (a header cell has no RTF/CSV equivalent, GFM has one header row, plain text has no
    // highlight), so a FAIL would encode "lossy by design" as a failure. A drop still has to be
    // visible, which is the entire point of adding them.
    const fidelity: Array<[string, number, number]> = [
        ['Header Cells', original.headerCells, regenerated.headerCells],
        ['Embeds', original.embeds, regenerated.embeds],
        ['Highlighted Runs', original.highlightedText, regenerated.highlightedText],
        ['Captions', original.captions, regenerated.captions],
    ];
    for (const [label, before, after] of fidelity) {
        // Nothing to prove when the source had none of this construct.
        if (before === 0 && after === 0) continue;
        const ok = before === after;
        results.push(mk(`${label} (exact)`, before, after, ok,
            ok ? `${after} preserved` : `Expected ${before}, got ${after} — lossy for this target?`,
            !ok));
    }

    return results;
}


// ============================================================================
// TEST RUNNERS
// ============================================================================

/**
 * Test generation from one source AST to one destination format.
 * Compares against a snapshot baseline if available.
 */
async function testGeneration(
    srcFmt: string,
    ast: OfficeParserAST,
    destFmt: GeneratorFormat
): Promise<GenFeatureTest[]> {
    const results: GenFeatureTest[] = [];
    const category = `Generation (${srcFmt}→${destFmt})`;

    const skip = (feature: string, details: string): GenFeatureTest => ({
        category, feature, sourceFormat: srcFmt, destFormat: destFmt,
        result: { status: 'SKIP', expected: 'N/A', actual: 'N/A', details, duration: 0 }
    });

    const startTime = Date.now();

    // CSV only makes sense for spreadsheet sources
    if (destFmt === 'csv' && !['xlsx', 'ods', 'csv'].includes(srcFmt)) {
        return [skip('Generation', 'CSV output only tested for spreadsheet sources')];
    }

    // RTF only for document sources
    if (destFmt === 'rtf' && ['xlsx', 'ods', 'csv', 'pptx', 'odp'].includes(srcFmt)) {
        return [skip('Generation', 'RTF output only tested for document sources')];
    }

    try {
        const result = await OfficeGenerator.generate(ast as any, destFmt as any, deterministicConfigFor(destFmt) as any);

        // CSV generator returns a ZIP Uint8Array for multi-sheet sources.
        // Detect this and unpack the first CSV file for metric extraction.
        // EPUB generator always returns a Uint8Array (zip archive).
        let rawOutput: string = '';
        let isZipOutput = false;
        let isEpubOutput = false;
        let isDocxOutput = false;
        let isOdtOutput = false;
        if (destFmt === 'epub') {
            // EPUB is always a binary ZIP archive
            isZipOutput = true;
            isEpubOutput = true;
            rawOutput = '[EPUB archive]';
        } else if (destFmt === 'docx') {
            // DOCX is always a binary ZIP (OOXML) archive.
            isDocxOutput = true;
            rawOutput = '[DOCX archive]';
        } else if (destFmt === 'odt') {
            // ODT is always a binary ZIP (ODF) archive.
            isOdtOutput = true;
            rawOutput = '[ODT archive]';
        } else if (destFmt === 'csv') {
            if (result.value instanceof Uint8Array) {
                // ZIP archive — convert to base64 placeholder metrics
                isZipOutput = true;
                rawOutput = '[ZIP archive]';
            } else {
                rawOutput = result.value as string;
            }
        } else if (destFmt === 'chunks') {
            rawOutput = Array.isArray(result.value) ? JSON.stringify(result.value, null, 2) : result.value as string;
        } else {
            rawOutput = result.value as string;
        }

        let actualMetrics: GeneratedMetrics;
        if (destFmt === 'chunks') {
            const chunks = Array.isArray(result.value) ? result.value : JSON.parse(rawOutput);
            actualMetrics = extractChunkMetrics(chunks);
        } else if (isDocxOutput) {
            actualMetrics = extractDocxMetrics(result.value as Uint8Array);
        } else if (isOdtOutput) {
            actualMetrics = extractOdtMetrics(result.value as Uint8Array);
        } else if (isZipOutput) {
            // ZIP: treat as valid multi-sheet output, just validate non-empty archive
            actualMetrics = {
                outputLength: (result.value as Uint8Array).length,
                hasContent: (result.value as Uint8Array).length > 0,
                lineCount: 0,
                wordCount: 0,
                rowCount: 0,
            };
        } else {
            actualMetrics = extractGeneratedMetrics(rawOutput, destFmt);
        }

        // Save the actual output
        const outPath = getGeneratorOutputPath(srcFmt, destFmt);
        if (destFmt === 'chunks') {
            fs.writeFileSync(`${outPath}.json`, rawOutput, 'utf8');
        } else if (isEpubOutput) {
            fs.writeFileSync(`${outPath}.epub`, result.value as Uint8Array);
        } else if (isDocxOutput) {
            fs.writeFileSync(`${outPath}.docx`, result.value as Uint8Array);
        } else if (isOdtOutput) {
            fs.writeFileSync(`${outPath}.odt`, result.value as Uint8Array);
        } else if (isZipOutput) {
            fs.writeFileSync(`${outPath}.zip`, result.value as Uint8Array);
        } else {
            fs.writeFileSync(outPath, rawOutput, 'utf8');
        }

        // Save metrics for baseline
        const metricsPath = `${outPath}.metrics.json`;
        fs.writeFileSync(metricsPath, JSON.stringify(actualMetrics, null, 2), 'utf8');

        // Must have non-empty output.
        // Spreadsheet → chunks may return empty if chunker skips raw cell content — treat as WARN.
        const isSpreadsheetToChunks = destFmt === 'chunks' && ['xlsx', 'ods', 'csv'].includes(srcFmt);
        results.push({
            category, feature: 'Non-Empty Output', sourceFormat: srcFmt, destFormat: destFmt,
            result: {
                status: actualMetrics.hasContent ? 'PASS' : (isSpreadsheetToChunks ? 'WARN' : 'FAIL'),
                expected: 'Non-empty', actual: actualMetrics.outputLength,
                details: actualMetrics.hasContent
                    ? `Generated ${actualMetrics.outputLength} chars, ${actualMetrics.wordCount} words`
                    : (isSpreadsheetToChunks
                        ? 'Spreadsheet → chunks: chunker skips raw cell content (known limitation)'
                        : 'Generator produced empty output')
            }
        });

        // No fatal conversion errors
        const errors = result.messages.filter(m => m.type === 'error');
        results.push({
            category, feature: 'No Errors', sourceFormat: srcFmt, destFormat: destFmt,
            result: {
                status: errors.length === 0 ? 'PASS' : 'FAIL',
                expected: 0, actual: errors.length,
                details: errors.length === 0 ? 'No errors during generation' : errors.map(e => e.message).join('; ')
            }
        });

        // Compare against baseline if available (skip for ZIP outputs — no meaningful text baseline)
        if (!isZipOutput) {
            const baseline = loadGeneratorBaseline(srcFmt, destFmt);
            if (baseline) {
                results.push(...compareGeneratedMetrics(category, srcFmt, destFmt, baseline, actualMetrics));
            } else {
                results.push({
                    category, feature: 'Baseline', sourceFormat: srcFmt, destFormat: destFmt,
                    result: {
                        status: 'WARN',
                        expected: 'Baseline file', actual: 'None',
                        details: `No baseline yet. Run with 'baseline' arg to create one.`
                    }
                });
            }
        } else if (isEpubOutput) {
            // EPUB: validate it's a non-empty zip and starts with the PK magic bytes
            const epubBytes = result.value as Uint8Array;
            const hasPkMagic = epubBytes.length >= 4 && epubBytes[0] === 0x50 && epubBytes[1] === 0x4B && epubBytes[2] === 0x03 && epubBytes[3] === 0x04;
            results.push({
                category, feature: 'EPUB Archive Valid', sourceFormat: srcFmt, destFormat: destFmt,
                result: {
                    status: (actualMetrics.hasContent && hasPkMagic) ? 'PASS' : 'FAIL',
                    expected: 'Non-empty EPUB zip (PK magic)', actual: `${actualMetrics.outputLength} bytes, magic=${hasPkMagic}`,
                    details: `EPUB archive: ${actualMetrics.outputLength} bytes, PK magic: ${hasPkMagic}`
                }
            });
        } else {
            results.push({
                category, feature: 'CSV ZIP Archive', sourceFormat: srcFmt, destFormat: destFmt,
                result: {
                    status: actualMetrics.hasContent ? 'PASS' : 'FAIL',
                    expected: 'Non-empty ZIP', actual: `${actualMetrics.outputLength} bytes`,
                    details: `Multi-sheet CSV archive: ${actualMetrics.outputLength} bytes`
                }
            });
        }

        // Format-specific sanity checks
        if (destFmt === 'html' && result.value) {
            const html = result.value as string;
            results.push({
                category, feature: 'HTML Valid Structure', sourceFormat: srcFmt, destFormat: destFmt,
                result: {
                    status: html.includes('<!DOCTYPE html>') || html.includes('<div') ? 'PASS' : 'WARN',
                    expected: 'Valid HTML', actual: html.substring(0, 50),
                    details: 'Checking for DOCTYPE or div container'
                }
            });
        }

        if (destFmt === 'tex' && typeof result.value === 'string') {
            const live = liveLatex(result.value);
            const begins = (live.match(/\\begin\{[^}]+\}/g) || []).length;
            const ends = (live.match(/\\end\{[^}]+\}/g) || []).length;
            const shell = /\\documentclass\[[^\]]*\]\{\w+\}/.test(live) && (live.match(/\\begin\{document\}/g) || []).length === 1
                && (live.match(/\\end\{document\}/g) || []).length === 1;
            // Every brace the body writes is balanced once escaped braces are discounted.
            const braces = live.replace(/\\[{}]/g, '');
            const balanced = (braces.match(/\{/g) || []).length === (braces.match(/\}/g) || []).length;
            results.push({
                category, feature: 'LaTeX Valid Structure', sourceFormat: srcFmt, destFormat: destFmt,
                result: {
                    status: shell && begins === ends && balanced ? 'PASS' : 'FAIL',
                    expected: 'One document environment, matched \\begin/\\end, balanced braces',
                    actual: `shell=${shell}, begin=${begins}, end=${ends}, balanced=${balanced}`,
                    details: 'Checked outside verbatim environments and comments'
                }
            });
        }

        if (destFmt === 'md' && result.value) {
            const md = result.value as string;
            const hasMarkdownSyntax = /#{1,6} |^\*\*|\[.*?\]\(|^---/m.test(md);
            results.push({
                category, feature: 'Markdown Syntax Present', sourceFormat: srcFmt, destFormat: destFmt,
                result: {
                    status: hasMarkdownSyntax ? 'PASS' : 'WARN',
                    expected: 'Markdown syntax', actual: hasMarkdownSyntax ? 'Found' : 'Not found',
                    details: 'Checking for headings, bold, links, or HR'
                }
            });
        }

        if (destFmt === 'chunks') {
            let chunks: any[] = [];
            try { chunks = JSON.parse(rawOutput); } catch { chunks = []; }
            if (chunks.length > 0) {
                const allHaveText = chunks.every(c => typeof c.text === 'string' && c.text.length > 0);
                const allHaveMeta = chunks.every(c => c.metadata && typeof c.metadata === 'object');
                results.push({
                    category, feature: 'Chunks Have Text', sourceFormat: srcFmt, destFormat: destFmt,
                    result: {
                        status: allHaveText ? 'PASS' : 'FAIL',
                        expected: 'All chunks have text', actual: allHaveText ? 'Yes' : 'No',
                        details: `${chunks.length} chunks, all with text: ${allHaveText}`
                    }
                });
                results.push({
                    category, feature: 'Chunks Have Metadata', sourceFormat: srcFmt, destFormat: destFmt,
                    result: {
                        status: allHaveMeta ? 'PASS' : 'FAIL',
                        expected: 'All chunks have metadata', actual: allHaveMeta ? 'Yes' : 'No',
                        details: `All chunks have metadata object: ${allHaveMeta}`
                    }
                });
            } else if (isSpreadsheetToChunks) {
                results.push({
                    category, feature: 'Chunks Have Text', sourceFormat: srcFmt, destFormat: destFmt,
                    result: {
                        status: 'WARN', expected: 'Chunks', actual: '0 chunks',
                        details: 'Spreadsheet chunking produces 0 chunks (cell content not yet chunked)'
                    }
                });
            }
        }

        const duration = Date.now() - startTime;

    } catch (err: any) {
        const duration = Date.now() - startTime;
        results.push({
            category, feature: 'Generation Error', sourceFormat: srcFmt, destFormat: destFmt,
            result: { status: 'FAIL', expected: 'Success', actual: 'Error', details: err.message, duration }
        });
    }

    // Assign duration to all results
    const totalDuration = Date.now() - startTime;
    results.forEach(r => {
        if (r.result.duration === undefined) r.result.duration = totalDuration;
    });

    return results;
}

/**
 * Test config permutations for a given source AST and destination format.
 */
async function testGeneratorConfigs(
    srcFmt: string,
    ast: OfficeParserAST,
    destFmt: GeneratorFormat
): Promise<GenFeatureTest[]> {
    const results: GenFeatureTest[] = [];

    // CSV/RTF: only test with compatible sources
    if (destFmt === 'csv' && !['xlsx', 'ods', 'csv'].includes(srcFmt)) return [];
    if (destFmt === 'rtf' && ['xlsx', 'ods', 'csv', 'pptx', 'odp'].includes(srcFmt)) return [];

    for (const ct of GENERATOR_CONFIG_TESTS) {
        // Skip configs restricted to specific formats
        if (ct.formats && !ct.formats.includes(destFmt)) continue;

        const category = `Config (${srcFmt}→${destFmt})`;
        const startTime = Date.now();

        try {
            const result = await OfficeGenerator.generate(ast as any, destFmt as any, ct.config as any);
            const isSpreadsheetToChunks = destFmt === 'chunks' && ['xlsx', 'ods', 'csv'].includes(srcFmt);
            const isZipResult = result.value instanceof Uint8Array;
            const hasOutput = isZipResult
                ? (result.value as Uint8Array).length > 0
                : destFmt === 'chunks'
                    ? (() => {
                        const val = result.value;
                        if (Array.isArray(val)) return val.length > 0;
                        try {
                            const c = JSON.parse(val as string);
                            return Array.isArray(c) && c.length > 0;
                        } catch { return false; }
                    })()
                    : (result.value as string).length > 0;

            results.push({
                category, feature: `${ct.id}: ${ct.name}`, sourceFormat: srcFmt, destFormat: destFmt,
                result: {
                    status: hasOutput ? 'PASS' : (isSpreadsheetToChunks ? 'WARN' : 'FAIL'),
                    expected: 'Non-empty output', actual: hasOutput ? 'OK' : 'Empty',
                    details: hasOutput
                        ? `Generated successfully with config ${ct.id}`
                        : isSpreadsheetToChunks
                            ? `Spreadsheet → chunks: no chunks produced (known limitation)`
                            : `Config ${ct.id} produced empty output`
                }
            });

            // G2: No Formatting — output should be shorter than default (less markup)
            if (ct.id === 'G2' && (destFmt === 'html' || destFmt === 'md')) {
                const defaultResult = await OfficeGenerator.generate(ast as any, destFmt as any, {} as any);
                const actualLen = typeof result.value === 'string' ? result.value.length : 0;
                const defaultLen = typeof defaultResult.value === 'string' ? defaultResult.value.length : 0;
                const shorter = actualLen <= defaultLen;
                results.push({
                    category, feature: `${ct.id}: No Formatting shorter`, sourceFormat: srcFmt, destFormat: destFmt,
                    result: {
                        status: shorter ? 'PASS' : 'WARN',
                        expected: '≤ default length', actual: `${actualLen} vs ${defaultLen}`,
                        details: shorter ? 'No-formatting output is leaner' : 'No-formatting output unexpectedly larger'
                    }
                });
            }

            // G6: HTML Non-Standalone (standalone:false) — bare fragment: no DOCTYPE, no <style>,
            // no envelope-level <script> (Chart.js CDN loader / spreadsheet interactivity), but
            // still has the content container div. Note a per-chart *inline* init script (which
            // contains "new Chart(") is content, not envelope, and is intentionally NOT gated by
            // this flag - it survives regardless, same as EpubGenerator relies on.
            if (ct.id === 'G6' && destFmt === 'html') {
                const htmlVal = typeof result.value === 'string' ? result.value : '';
                const hasDoctype = htmlVal.includes('<!DOCTYPE');
                results.push({
                    category, feature: `${ct.id}: No DOCTYPE in fragment`, sourceFormat: srcFmt, destFormat: destFmt,
                    result: {
                        status: !hasDoctype ? 'PASS' : 'FAIL',
                        expected: 'No DOCTYPE', actual: hasDoctype ? 'Has DOCTYPE' : 'No DOCTYPE',
                        details: 'Non-standalone HTML should not include DOCTYPE'
                    }
                });

                const hasStyle = htmlVal.includes('<style>');
                const hasEnvelopeScript = htmlVal.includes('cdn.jsdelivr.net/npm/chart.js') || htmlVal.includes('initSpreadsheetResizing');
                const hasContainer = /<div class="[^"]*container[^"]*"/.test(htmlVal);
                results.push({
                    category, feature: `${ct.id}: Bare fragment (no style/envelope-script, has container)`, sourceFormat: srcFmt, destFormat: destFmt,
                    result: {
                        status: (!hasStyle && !hasEnvelopeScript && hasContainer) ? 'PASS' : 'FAIL',
                        expected: 'No <style>, no envelope <script>, has container div',
                        actual: `hasStyle=${hasStyle}, hasEnvelopeScript=${hasEnvelopeScript}, hasContainer=${hasContainer}`,
                        details: 'standalone:false turns every envelope part off (bare content fragment)'
                    }
                });
            }

            // G15: standalone:{document:false} — still a fragment (no DOCTYPE), but styled
            // (styles defaults to 'full' when omitted from the object).
            if (ct.id === 'G15' && destFmt === 'html') {
                const htmlVal = typeof result.value === 'string' ? result.value : '';
                const hasDoctype = htmlVal.includes('<!DOCTYPE');
                const hasStyle = htmlVal.includes('<style>');
                results.push({
                    category, feature: `${ct.id}: Styled fragment (document:false only)`, sourceFormat: srcFmt, destFormat: destFmt,
                    result: {
                        status: (!hasDoctype && hasStyle) ? 'PASS' : 'FAIL',
                        expected: 'No DOCTYPE, has <style>', actual: `hasDoctype=${hasDoctype}, hasStyle=${hasStyle}`,
                        details: 'Omitted StandaloneConfig fields default to their "on" (standalone) value'
                    }
                });
            }

            // G16: standalone:{document:false, styles:'scoped'} — CSS wrapped in @scope, and
            // the leak-prone :root/body selectors must not appear unscoped.
            if (ct.id === 'G16' && destFmt === 'html') {
                const htmlVal = typeof result.value === 'string' ? result.value : '';
                const hasScopeBlock = htmlVal.includes('@scope (.op-html-scope)');
                const hasBareBodyOrRoot = /(^|\n)\s*body\s*\{/.test(htmlVal) || /:root\s*\{/.test(htmlVal);
                results.push({
                    category, feature: `${ct.id}: Scoped styles via @scope`, sourceFormat: srcFmt, destFormat: destFmt,
                    result: {
                        status: (hasScopeBlock && !hasBareBodyOrRoot) ? 'PASS' : 'FAIL',
                        expected: 'Has @scope block, no bare body{}/:root{}',
                        actual: `hasScopeBlock=${hasScopeBlock}, hasBareBodyOrRoot=${hasBareBodyOrRoot}`,
                        details: 'styles:"scoped" must not leak global selectors onto a host page'
                    }
                });
            }

            // G17: standalone:{document:false, bodyInjections:true} + injections — body
            // injections must apply even to a bare content fragment (they wrap body content,
            // not the document shell).
            if (ct.id === 'G17' && destFmt === 'html') {
                const htmlVal = typeof result.value === 'string' ? result.value : '';
                const hasMarkers = htmlVal.includes('<!--OP_BODY_START_TEST-->') && htmlVal.includes('<!--OP_BODY_END_TEST-->');
                results.push({
                    category, feature: `${ct.id}: Body injections applied to fragment`, sourceFormat: srcFmt, destFormat: destFmt,
                    result: {
                        status: hasMarkers ? 'PASS' : 'FAIL',
                        expected: 'Both body injection markers present', actual: hasMarkers ? 'Present' : 'Missing',
                        details: 'bodyInjections apply regardless of document, unlike headInjections'
                    }
                });
            }

            // G18: standalone:true (explicit) — full document, must contain DOCTYPE.
            if (ct.id === 'G18' && destFmt === 'html') {
                const htmlVal = typeof result.value === 'string' ? result.value : '';
                const hasDoctype = htmlVal.includes('<!DOCTYPE');
                results.push({
                    category, feature: `${ct.id}: DOCTYPE present`, sourceFormat: srcFmt, destFormat: destFmt,
                    result: {
                        status: hasDoctype ? 'PASS' : 'FAIL',
                        expected: 'Has DOCTYPE', actual: hasDoctype ? 'Has DOCTYPE' : 'No DOCTYPE',
                        details: 'standalone:true must still produce a full standalone document'
                    }
                });
            }

            // G-docx: Legal + landscape + margins must reach word/document.xml's sectPr.
            if (ct.id === 'G-docx' && destFmt === 'docx' && result.value instanceof Uint8Array) {
                const doc = strFromU8(unzipSync(result.value)['word/document.xml']);
                const ok = /<w:pgSz w:w="20160"[^>]*w:orient="landscape"/.test(doc)
                    && /<w:pgMar w:top="720" w:right="360" w:bottom="720" w:left="360"/.test(doc);
                results.push({
                    category, feature: `${ct.id}: page size/orient/margins applied`, sourceFormat: srcFmt, destFormat: destFmt,
                    result: { status: ok ? 'PASS' : 'FAIL', expected: 'Legal landscape, 36/18pt margins', actual: doc.match(/<w:pgSz[^>]*>/)?.[0] ?? 'no pgSz', details: 'docxConfig must reach the section properties' }
                });
            }

            // G-odt: Letter + landscape + margins must reach styles.xml's page-layout.
            if (ct.id === 'G-odt' && destFmt === 'odt' && result.value instanceof Uint8Array) {
                const styles = strFromU8(unzipSync(result.value)['styles.xml']);
                const ok = /fo:page-width="11in"/.test(styles) && /fo:page-height="8.5in"/.test(styles)
                    && /style:print-orientation="landscape"/.test(styles)
                    && /fo:margin-top="36pt"/.test(styles) && /fo:margin-left="18pt"/.test(styles);
                results.push({
                    category, feature: `${ct.id}: page size/orient/margins applied`, sourceFormat: srcFmt, destFormat: destFmt,
                    result: { status: ok ? 'PASS' : 'FAIL', expected: 'Letter landscape, 36/18pt margins', actual: styles.match(/<style:page-layout-properties[^>]*>/)?.[0]?.slice(0, 120) ?? 'no page-layout', details: 'odtConfig must reach the page layout' }
                });
            }

            // G-tex: class, numbering and page geometry reach the preamble.
            if (ct.id === 'G-tex' && destFmt === 'tex' && typeof result.value === 'string') {
                const tex = result.value;
                const ok = /\\documentclass\[[^\]]*\]\{report\}/.test(tex) && !tex.includes('secnumdepth')
                    && /\\usepackage\[paperwidth=612pt,paperheight=792pt,landscape,top=36pt,bottom=36pt,left=18pt,right=18pt\]\{geometry\}/.test(tex);
                results.push({
                    category, feature: `${ct.id}: class/numbering/geometry applied`, sourceFormat: srcFmt, destFormat: destFmt,
                    result: { status: ok ? 'PASS' : 'FAIL', expected: 'report class, numbered, Letter landscape 36/18pt', actual: tex.slice(0, 160), details: 'texConfig must reach the preamble' }
                });
            }

            // G-tex-fragment: body only, headed by the package list, no document shell.
            if (ct.id === 'G-tex-fragment' && destFmt === 'tex' && typeof result.value === 'string') {
                const tex = result.value;
                const ok = !tex.includes('\\documentclass') && !tex.includes('\\begin{document}') && tex.startsWith('% LaTeX fragment generated by officeParser.');
                results.push({
                    category, feature: `${ct.id}: fragment has no document shell`, sourceFormat: srcFmt, destFormat: destFmt,
                    result: { status: ok ? 'PASS' : 'FAIL', expected: 'package comment header, no \\documentclass', actual: tex.slice(0, 120), details: 'standalone:false emits the body only' }
                });
            }

            // G-tex-bundle: a zip whose main.tex references exactly the images packaged beside it.
            if (ct.id === 'G-tex-bundle' && destFmt === 'tex') {
                let ok = false, detail = '';
                if (result.value instanceof Uint8Array) {
                    const files = unzipSync(result.value);
                    const main = files['main.tex'] ? strFromU8(files['main.tex']) : '';
                    const packaged = Object.keys(files).filter(n => n !== 'main.tex').sort();
                    const referenced = [...new Set([...main.matchAll(/\\includegraphics(?:\[[^\]]*\])?\{([^}]+)\}/g)].map(m => m[1]))];
                    // An image the source only named by path cannot be packaged; the bundle guards it with
                    // \IfFileExists so it still compiles. Every other reference must be in the zip.
                    const guarded = new Set([...main.matchAll(/\\IfFileExists\{([^}]+)\}\{\\includegraphics(?:\[[^\]]*\])?\{([^}]+)\}/g)].map(m => m[2]));
                    ok = !!main && packaged.every(n => /^images\/[A-Za-z0-9-]+\.[a-z]+$/.test(n)) && referenced.every(r => packaged.includes(r) || guarded.has(r));
                    detail = `${packaged.length} packaged, ${referenced.length} referenced, ${guarded.size} guarded`;
                } else {
                    detail = 'bundle:true did not return a Uint8Array';
                }
                results.push({
                    category, feature: `${ct.id}: zip holds main.tex and every referenced image`, sourceFormat: srcFmt, destFormat: destFmt,
                    result: { status: ok ? 'PASS' : 'FAIL', expected: 'main.tex + images/*', actual: detail, details: 'bundle must be self-contained' }
                });
            }

            // G3: Metadata rendered — output should contain metadata block
            if (ct.id === 'G3' && typeof result.value === 'string') {
                const output = result.value;
                const outputLower = output.toLowerCase();
                // Check for standard keys or any custom property keys
                const possibleKeys = ['title', 'author', 'created', 'modified', 'description', 'pages', 'metadata'];
                if (ast.metadata?.customProperties) {
                    possibleKeys.push(...Object.keys(ast.metadata.customProperties).map(k => k.toLowerCase()));
                }
                const hasMeta = possibleKeys.some(k => outputLower.includes(k)) || output.includes('---');

                results.push({
                    category, feature: `${ct.id}: Metadata In Output`, sourceFormat: srcFmt, destFormat: destFmt,
                    result: {
                        status: hasMeta ? 'PASS' : 'WARN',
                        expected: 'Metadata content', actual: hasMeta ? 'Found' : 'Not found',
                        details: hasMeta ? 'Metadata block detected' : 'renderMetadata=true should include metadata block'
                    }
                });
            }

        } catch (err: any) {
            results.push({
                category, feature: `${ct.id}: ${ct.name}`, sourceFormat: srcFmt, destFormat: destFmt,
                result: { status: 'FAIL', expected: 'Success', actual: 'Error', details: err.message, duration: Date.now() - startTime }
            });
        }

        // Assign duration to config results for this permutation
        const permutationDuration = Date.now() - startTime;
        results.filter(r => r.category === category && r.feature.startsWith(ct.id)).forEach(r => {
            if (r.result.duration === undefined) r.result.duration = permutationDuration;
        });
    }

    return results;
}

/**
 * Run roundtrip test: parse source file → generate to destFmt → re-parse → compare ASTs.
 * Only runs for formats where parse(generate(parse(src))) should be lossless.
 */
async function testRoundtrip(srcFmt: string): Promise<GenFeatureTest[]> {
    const results: GenFeatureTest[] = [];
    const destFmt = srcFmt as GeneratorFormat;

    if (!ROUNDTRIP_FORMATS.includes(destFmt)) {
        return [{
            category: `Roundtrip (${srcFmt}→${destFmt}→${srcFmt})`,
            feature: 'Roundtrip', sourceFormat: srcFmt, destFormat: destFmt,
            result: { status: 'SKIP', expected: 'Roundtrip format', actual: srcFmt, details: 'Format not in roundtrip list', duration: 0 }
        }];
    }

    const startTime = Date.now();

    const srcPath = getSourcePath(srcFmt);
    if (!fs.existsSync(srcPath)) {
        return [{
            category: `Roundtrip (${srcFmt}→${destFmt}→${srcFmt})`,
            feature: 'Source File', sourceFormat: srcFmt, destFormat: destFmt,
            result: { status: 'SKIP', expected: 'File exists', actual: 'Missing', details: `${srcPath} not found` }
        }];
    }

    try {
        // Step 1: Parse original file
        const ast1 = await OfficeParser.parseOffice(srcPath, PARSER_CONFIG);
        const metrics1 = await extractRoundtripMetrics(ast1);

        // Step 2: Generate to same format
        const genResult = await OfficeGenerator.generate(ast1 as any, destFmt as any);
        const genContent = genResult.value;

        // Step 3: Write generated file to temp, re-parse it (binary for zip-backed formats like docx)
        const tmpPath = path.join(__dirname, '..', 'files', `_roundtrip_tmp.${destFmt}`);
        if (genContent instanceof Uint8Array) fs.writeFileSync(tmpPath, genContent);
        else fs.writeFileSync(tmpPath, genContent as string, 'utf8');

        try {
            const ast2 = await OfficeParser.parseOffice(tmpPath, { ...PARSER_CONFIG, fileType: srcFmt as any });
            const metrics2 = await extractRoundtripMetrics(ast2);

            results.push(...compareRoundtripASTs(srcFmt, destFmt, metrics1, metrics2));
        } finally {
            if (fs.existsSync(tmpPath)) fs.unlinkSync(tmpPath);
        }
    } catch (err: any) {
        results.push({
            category: `Roundtrip (${srcFmt}→${destFmt}→${srcFmt})`,
            feature: 'Roundtrip Error', sourceFormat: srcFmt, destFormat: destFmt,
            result: { status: 'FAIL', expected: 'Success', actual: 'Error', details: err.message, duration: Date.now() - startTime }
        });
    }

    const totalDuration = Date.now() - startTime;
    results.forEach(r => {
        if (r.result.duration === undefined) r.result.duration = totalDuration;
    });

    return results;
}

/**
 * CSV roundtrip: csv source → generate csv (merge) → re-parse and compare table row counts.
 */
async function testCsvRoundtrip(): Promise<GenFeatureTest[]> {
    const results: GenFeatureTest[] = [];
    const category = 'Roundtrip (csv→csv→csv)';

    const srcPath = getSourcePath('csv');
    if (!fs.existsSync(srcPath)) {
        return [{
            category, feature: 'Source File', sourceFormat: 'csv', destFormat: 'csv',
            result: { status: 'SKIP', expected: 'File exists', actual: 'Missing', details: 'test.csv not found', duration: 0 }
        }];
    }

    const startTime = Date.now();

    try {
        const ast1 = await OfficeParser.parseOffice(srcPath, { ...PARSER_CONFIG, fileType: 'csv' });
        const text1 = (await ast1.to('text')).value;
        const lines1 = text1.split('\n').filter(l => l.trim() && !l.startsWith('#')).length;

        const genResult = await OfficeGenerator.generate(ast1 as any, 'csv', { csvConfig: { mergeSheets: true } } as any);
        const genCsv = genResult.value as string;
        const tmpPath = path.join(__dirname, '..', 'files', '_roundtrip_tmp.csv');
        fs.writeFileSync(tmpPath, genCsv, 'utf8');

        try {
            const ast2 = await OfficeParser.parseOffice(tmpPath, { ...PARSER_CONFIG, fileType: 'csv' });
            const text2 = (await ast2.to('text')).value;
            const lines2 = text2.split('\n').filter(l => l.trim() && !l.startsWith('#')).length;

            const textRatio = text1.length > 0 ? text2.length / text1.length : 1;
            const textOk = textRatio >= 0.95 && textRatio <= 1.05;
            results.push({
                category, feature: 'Text Length (±5%)', sourceFormat: 'csv', destFormat: 'csv',
                result: {
                    status: textOk ? 'PASS' : 'FAIL', expected: text1.length, actual: text2.length,
                    details: `${(textRatio * 100).toFixed(1)}% of original text`
                }
            });

            const lineRatio = lines1 > 0 ? lines2 / lines1 : 1;
            const lineOk = Math.abs(lineRatio - 1) <= 0.05;
            results.push({
                category, feature: 'Row Count (±5%)', sourceFormat: 'csv', destFormat: 'csv',
                result: {
                    status: lineOk ? 'PASS' : 'WARN', expected: lines1, actual: lines2,
                    details: `${(lineRatio * 100).toFixed(1)}% of original rows`
                }
            });
        } finally {
            if (fs.existsSync(tmpPath)) fs.unlinkSync(tmpPath);
        }
    } catch (err: any) {
        results.push({
            category, feature: 'CSV Roundtrip Error', sourceFormat: 'csv', destFormat: 'csv',
            result: { status: 'FAIL', expected: 'Success', actual: 'Error', details: err.message, duration: Date.now() - startTime }
        });
    }

    const totalDuration = Date.now() - startTime;
    results.forEach(r => {
        if (r.result.duration === undefined) r.result.duration = totalDuration;
    });

    return results;
}


// ============================================================================
// DUAL LOGGER
// ============================================================================

class DualLogger {
    private mdContent: string = '';
    log(message: string = '') { console.log(message); this.mdContent += message + '\n'; }
    getMarkdown(): string { return '```\n' + this.mdContent + '```\n'; }
}

// ============================================================================
// REPORTING
// ============================================================================

function generateReport(allResults: GenFeatureTest[], logger: DualLogger): number {
    const width = 150;
    const line = '═'.repeat(width);

    logger.log('┌' + '─'.repeat(width - 2) + '┐');
    logger.log('│' + ' '.repeat(width - 2) + '│');
    logger.log('│' + '    OFFICE GENERATOR COMPREHENSIVE TEST SUITE'.padEnd(width - 2) + '│');
    logger.log('│' + '    Generation, Config Permutations & Roundtrip Parity'.padEnd(width - 2) + '│');
    logger.log('│' + ' '.repeat(width - 2) + '│');
    logger.log('└' + '─'.repeat(width - 2) + '┘');
    logger.log('');

    const byCategory: Record<string, GenFeatureTest[]> = {};
    allResults.forEach(r => {
        if (!byCategory[r.category]) byCategory[r.category] = [];
        byCategory[r.category].push(r);
    });

    for (const [category, tests] of Object.entries(byCategory)) {
        logger.log(line);
        logger.log(category.toUpperCase());
        logger.log(line);
        logger.log('');
        logger.log('┌──────────┬──────────┬──────────────────────────────────────────────┬────────┬──────────┬─────────────────────────────────────────────────────┐');
        logger.log('│ Source   │ Dest     │ Feature                                      │ Status │ Time     │ Details                                             │');
        logger.log('├──────────┼──────────┼──────────────────────────────────────────────┼────────┼──────────┼─────────────────────────────────────────────────────┤');

        for (const test of tests) {
            const icon = { PASS: '✓', FAIL: '✗', WARN: '⚠', SKIP: '⊘' }[test.result.status];
            const src = test.sourceFormat.padEnd(8);
            const dst = test.destFormat.padEnd(8);
            const feat = test.feature.substring(0, 44).padEnd(44);
            const status = `${icon} ${test.result.status}`.padEnd(6);
            const time = test.result.duration !== undefined ? (test.result.duration >= 1000 ? `${(test.result.duration / 1000).toFixed(1)}s` : `${test.result.duration}ms`).padEnd(8) : 'N/A'.padEnd(8);
            const det = test.result.details.substring(0, 53).padEnd(53);
            logger.log(`│ ${src} │ ${dst} │ ${feat} │ ${status} │ ${time} │ ${det} │`);
        }
        logger.log('└──────────┴──────────┴──────────────────────────────────────────────┴────────┴──────────┴─────────────────────────────────────────────────────┘');
        logger.log('');
    }

    // Summary
    logger.log(line);
    logger.log('SUMMARY');
    logger.log(line);
    logger.log('');

    const passed = allResults.filter(r => r.result.status === 'PASS').length;
    const failed = allResults.filter(r => r.result.status === 'FAIL').length;
    const warned = allResults.filter(r => r.result.status === 'WARN').length;
    const skipped = allResults.filter(r => r.result.status === 'SKIP').length;
    const total = allResults.length;

    logger.log(`Total Tests: ${total}`);
    logger.log(`✓ Passed:  ${passed} (${((passed / total) * 100).toFixed(1)}%)`);
    logger.log(`✗ Failed:  ${failed} (${((failed / total) * 100).toFixed(1)}%)`);
    logger.log(`⚠ Warned:  ${warned} (${((warned / total) * 100).toFixed(1)}%)`);
    logger.log(`⊘ Skipped: ${skipped} (${((skipped / total) * 100).toFixed(1)}%)`);
    logger.log('');

    if (failed === 0) {
        logger.log('✓ All tests passed!');
    } else {
        logger.log(`✗ ${failed} test(s) FAILED`);
        logger.log('');
        logger.log('Failed tests:');
        allResults.filter(r => r.result.status === 'FAIL').forEach(r => {
            logger.log(`  ✗ [${r.sourceFormat}→${r.destFormat}] ${r.feature}: ${r.result.details}`);
        });
    }

    // Save reports
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const resultsDir = path.join(__dirname, 'results');
    if (!fs.existsSync(resultsDir)) fs.mkdirSync(resultsDir, { recursive: true });

    const jsonPath = path.join(resultsDir, `gen-test-results-${timestamp}.json`);
    const mdPath = path.join(resultsDir, `gen-test-results-${timestamp}.md`);

    fs.writeFileSync(jsonPath, JSON.stringify({
        timestamp: new Date().toISOString(),
        summary: { total, passed, failed, warned, skipped },
        results: allResults
    }, null, 2));

    const md = '# Office Generator Test Results\n\n' +
        `**Generated**: ${new Date().toLocaleString()}\n\n` +
        logger.getMarkdown();
    fs.writeFileSync(mdPath, md);

    logger.log('');
    logger.log(`Results saved to:`);
    logger.log(`  JSON: ${jsonPath}`);
    logger.log(`  Markdown: ${mdPath}`);

    return failed;
}

/**
 * Runs unit tests for generator configuration validation rules.
 */
function runConfigValidationTests(): GenFeatureTest[] {
    const results: GenFeatureTest[] = [];
    const category = 'Config Validation';
    const sourceFormat = 'docx';
    const destFormat = 'html';

    const mk = (feature: string, exp: any, act: any, ok: boolean, details: string): GenFeatureTest => ({
        category, feature, sourceFormat, destFormat,
        result: { status: ok ? 'PASS' : 'FAIL', expected: exp, actual: act, details, duration: 0 }
    });

    // Valid test cases
    const validWidths = ['auto', 900, '900px', '100%', '50vw', '.5rem', '  120  ', '80.5%'];
    validWidths.forEach(w => {
        try {
            const config = resolveGeneratorConfig('html', undefined, { htmlConfig: { containerWidth: w } });
            const resolved = config.htmlConfig.containerWidth;
            const ok = resolved === w || (typeof w === 'number' && resolved === w) || (typeof w === 'string' && resolved === w);
            results.push(mk(`Valid: ${w}`, 'Should not throw', resolved, ok, `Resolved to: ${resolved}`));
        } catch (e: any) {
            results.push(mk(`Valid: ${w}`, 'Should not throw', e.message, false, `Threw error unexpectedly: ${e.message}`));
        }
    });

    // Invalid test cases
    const invalidWidths = [0, -100, '-50px', 'abc', '', '  ', 'NaN', NaN, Infinity, -Infinity, {}, [], null];
    invalidWidths.forEach(w => {
        try {
            let warningTriggered = false;
            let warningCode = '';
            const config = resolveGeneratorConfig('html', undefined, {
                onWarning: (issue) => {
                    warningTriggered = true;
                    warningCode = issue.code;
                },
                htmlConfig: { containerWidth: w as any }
            });
            const resolved = config.htmlConfig.containerWidth;
            const ok = warningTriggered && warningCode === 'INVALID_CONTAINER_WIDTH' && resolved === 'auto';
            results.push(mk(`Invalid: ${JSON.stringify(w)}`, 'Should warn and fallback to auto', `Warned: ${warningTriggered}, Code: ${warningCode}, Width: ${resolved}`, ok, `Warning: ${warningCode}, Fallback: ${resolved}`));
        } catch (e: any) {
            results.push(mk(`Invalid: ${JSON.stringify(w)}`, 'Should warn and fallback to auto', e.message, false, `Threw error unexpectedly: ${e.message}`));
        }
    });

    return results;
}

/**
 * Tests for the Markdown dialect/fallbackToHtml config system: dialect presets, the `extends`
 * composition rule, the merge-erasure gotcha, and the `fallbackToHtml: boolean |
 * FallbackToHtmlConfig` shorthand-object pattern. Driven off `test/files/exhaustive/markdown.md`,
 * which covers every construct these options gate.
 */
async function runMarkdownDialectTests(): Promise<GenFeatureTest[]> {
    const results: GenFeatureTest[] = [];
    const category = 'Markdown Dialect';
    const sourceFormat = 'md';
    const destFormat = 'md';

    const mk = (feature: string, exp: any, act: any, ok: boolean, details: string): GenFeatureTest => ({
        category, feature, sourceFormat, destFormat,
        result: { status: ok ? 'PASS' : 'FAIL', expected: exp, actual: act, details, duration: 0 }
    });

    const fixturePath = path.join(__dirname, '..', 'files', 'exhaustive', 'markdown.md');
    const ast = await OfficeParser.parseOffice(fixturePath);

    const genMdConfig = async (mdConfig: any, onWarning?: (issue: any) => void): Promise<string> => {
        const result = await OfficeGenerator.generate(ast as any, 'md' as any, { mdConfig, onWarning } as any);
        return typeof result.value === 'string' ? result.value : '';
    };
    const genDialect = (dialect: any) => genMdConfig({ dialect });

    // --- Preset assertions ---
    const extended = await genDialect('extended');
    results.push(mk('extended: GitHub admonitions (default, backward-compatible)', 'contains "> [!"', extended.includes('> [!'), extended.includes('> [!'), 'omitting dialect entirely must reproduce historical GitHub-admonition output'));
    results.push(mk('extended: wikilinks native', 'contains "[["', extended.includes('[['), extended.includes('[['), 'extended preset keeps wikilinks native'));
    results.push(mk('extended: citations native', 'contains "[@"', extended.includes('[@'), extended.includes('[@'), 'extended preset keeps citations native'));
    results.push(mk('extended: definition lists native', 'contains ": Description"', extended.includes(': Description'), extended.includes(': Description'), 'extended preset keeps definition lists native'));

    const github = await genDialect('github');
    results.push(mk('github: admonitions', 'contains "> [!"', github.includes('> [!'), github.includes('> [!'), 'github dialect uses GitHub admonition syntax'));
    results.push(mk('github: wikilinks degrade to plain link', 'no "[[", has plain link to WikiPage', !github.includes('[[') && github.includes('(WikiPage)'), !github.includes('[[') && github.includes('(WikiPage)'), 'github dialect has no wikilinks - falls back to a plain link using the same target'));
    results.push(mk('github: citations degrade (no @)', 'no "[@", has "[smith2023]"', !github.includes('[@') && github.includes('[smith2023]'), !github.includes('[@') && github.includes('[smith2023]'), 'github dialect drops the @ but keeps the key'));

    const gitlab = await genDialect('gitlab');
    results.push(mk('gitlab: admonitions fenced-div', 'contains ":::note"', gitlab.includes(':::note'), gitlab.includes(':::note'), 'gitlab dialect uses :::type fenced-div admonitions'));
    results.push(mk('gitlab: not pandoc-style', 'no "::: {."', !gitlab.includes('::: {.'), !gitlab.includes('::: {.'), 'gitlab fenced-div has no class syntax'));

    const obsidian = await genDialect('obsidian');
    results.push(mk('obsidian: wikilinks native', 'contains "[["', obsidian.includes('[['), obsidian.includes('[['), 'obsidian dialect keeps wikilinks native'));
    results.push(mk('obsidian: admonitions github-style', 'contains "> [!"', obsidian.includes('> [!'), obsidian.includes('> [!'), 'obsidian callouts use the same > [!type] form as GitHub'));

    const pandoc = await genDialect('pandoc');
    results.push(mk('pandoc: admonitions fenced-div-with-class', 'contains "::: {.note}"', pandoc.includes('::: {.note}'), pandoc.includes('::: {.note}'), 'pandoc dialect uses its own ::: {.type} native fenced-div-with-class syntax'));
    results.push(mk('pandoc: citations native', 'contains "[@"', pandoc.includes('[@'), pandoc.includes('[@'), 'pandoc dialect keeps citations native'));
    results.push(mk('pandoc: definition lists native', 'contains ": Description"', pandoc.includes(': Description'), pandoc.includes(': Description'), 'pandoc dialect keeps definition lists native (its own signature feature)'));
    results.push(mk('pandoc: wikilinks degrade', 'no "[["', !pandoc.includes('[['), !pandoc.includes('[['), 'pandoc dialect has no wikilinks'));

    const commonmark = await genDialect('commonmark');
    results.push(mk('commonmark: tables forced HTML', 'contains "<table"', commonmark.includes('<table'), commonmark.includes('<table'), 'strict CommonMark has no table syntax of its own'));
    results.push(mk('commonmark: no admonition marker', 'no "> [!" and no ":::"', !commonmark.includes('> [!') && !commonmark.includes(':::'), !commonmark.includes('> [!') && !commonmark.includes(':::'), 'admonitions degrade to a plain blockquote'));
    results.push(mk('commonmark: no strikethrough marker', 'no "~~"', !commonmark.includes('~~'), !commonmark.includes('~~'), 'strikethrough is a GFM-only extension'));
    results.push(mk('commonmark: no citations/wikilinks', 'no "[@" and no "[["', !commonmark.includes('[@') && !commonmark.includes('[['), !commonmark.includes('[@') && !commonmark.includes('[['), 'citations/wikilinks are not part of base CommonMark'));
    results.push(mk('commonmark: no bare $ math', 'no "$E=mc"', !commonmark.includes('$E=mc'), !commonmark.includes('$E=mc'), 'math delimiters are not part of base CommonMark'));
    results.push(mk('commonmark: footnote inlined, not [^id]', 'no "[^fn1]", has inline "(Note: This is the footnote"', !commonmark.includes('[^fn1]') && commonmark.includes('(Note: This is the footnote'), !commonmark.includes('[^fn1]') && commonmark.includes('(Note: This is the footnote'), 'footnotes:false inlines the note body as a parenthetical instead of dropping it'));

    // --- extends composition: object form composes a named base preset with explicit overrides ---
    const extendsGithub = await genDialect({ extends: 'github', bulletListMarker: '*' });
    const usesStarBullets = /\n\* Unordered item A/.test('\n' + extendsGithub);
    results.push(mk('extends: github base + bulletListMarker override', 'contains "> [!" (github base) and "* " bullets (override)', extendsGithub.includes('> [!') && usesStarBullets, extendsGithub.includes('> [!') && usesStarBullets, 'object form composes a named base preset with explicit field overrides'));
    results.push(mk('extends: github base keeps other github fields', 'no "[[" (github has no wikilinks)', !extendsGithub.includes('[['), !extendsGithub.includes('[['), 'fields not explicitly overridden still come from the named base preset'));

    // --- merge-erasure gotcha: an object override with no `extends` must default to 'extended',
    // not silently drop to some other ambient default ---
    const noExtendsOverride = await genDialect({ bulletListMarker: '*' });
    results.push(mk('no extends: defaults to extended base', 'contains "> [!" and "[[" (extended defaults)', noExtendsOverride.includes('> [!') && noExtendsOverride.includes('[['), noExtendsOverride.includes('> [!') && noExtendsOverride.includes('[['), 'omitting extends must default to extended, not silently drop other features'));

    // --- fallbackToHtml: boolean | FallbackToHtmlConfig (same shorthand-object pattern as
    // HtmlGeneratorConfig.standalone/StandaloneConfig - no separate successor field) ---
    const fallbackFalse = await genMdConfig({ fallbackToHtml: false });
    results.push(mk('fallbackToHtml:false turns every part off', 'no "<u>"', !fallbackFalse.includes('<u>'), !fallbackFalse.includes('<u>'), 'the boolean shorthand still applies uniformly to every part'));

    const onlyTextFormattingOff = await genMdConfig({ fallbackToHtml: { textFormatting: false } });
    results.push(mk('fallbackToHtml:{textFormatting:false} only disables that part', 'no "<u>" but still has "<table"', !onlyTextFormattingOff.includes('<u>') && onlyTextFormattingOff.includes('<table'), !onlyTextFormattingOff.includes('<u>') && onlyTextFormattingOff.includes('<table'), 'omitted object fields (tables) still default to on'));

    const onlyTablesOff = await genMdConfig({ fallbackToHtml: { tables: false } });
    results.push(mk('fallbackToHtml:{tables:false} only disables that part', 'has "<u>" but no merged-cell "<table"', onlyTablesOff.includes('<u>') && !onlyTablesOff.includes('<table'), onlyTablesOff.includes('<u>') && !onlyTablesOff.includes('<table'), 'omitted object fields (textFormatting) still default to on'));

    return results;
}

/**
 * Tests for `metadataOverrides`, the common-config way to set the metadata embedded in generated
 * output without mutating the parsed AST.
 *
 * Covers the three properties that make it safe to use: it merges per field rather than replacing
 * wholesale, it leaves `ast.metadata` untouched so one generation can't leak into the next, and it
 * reaches every generator that embeds metadata rather than only the one it was built for.
 */
async function runMetadataOverrideTests(): Promise<GenFeatureTest[]> {
    const results: GenFeatureTest[] = [];
    const category = 'Metadata Overrides';
    const mk = (destFormat: string, feature: string, exp: any, act: any, ok: boolean, details: string): GenFeatureTest => ({
        category, feature, sourceFormat: 'synthetic', destFormat,
        result: { status: ok ? 'PASS' : 'FAIL', expected: exp, actual: act, details, duration: 0 }
    });

    const parserConfig = { ...PARSER_CONFIG };
    const sourceMeta = { title: 'Original Title', author: 'Original Author', description: 'Original Description' };
    const content: OfficeContentNode[] = [
        { type: 'paragraph', children: [{ type: 'text', text: 'Body' }] } as OfficeContentNode,
    ];
    const makeAst = () => createAST('docx', { ...sourceMeta }, content, [], parserConfig, undefined, () => 'Body');

    // --- Per-field merge: overriding one field must not blank the others ---
    const astMerge = makeAst();
    const { value: mergedHtml } = await astMerge.to('html', {
        metadataOverrides: { author: 'Override Author' },
    } as any);
    results.push(mk('html', 'overriding one field preserves the others',
        'title "Original Title" kept, author replaced', String(mergedHtml).match(/<meta name="author"[^>]*>/)?.[0] ?? '(none)',
        mergedHtml.includes('Original Title') && mergedHtml.includes('Override Author') && !mergedHtml.includes('Original Author'),
        'a partial override that blanked unset fields would make it unusable for a single-field change like a pinned timestamp'));

    // --- The AST must not be mutated: overrides are an output concern ---
    const astPure = makeAst();
    await astPure.to('html', { metadataOverrides: { title: 'Temporary' } } as any);
    results.push(mk('html', 'ast.metadata is not mutated by an override',
        sourceMeta.title, astPure.metadata?.title,
        astPure.metadata?.title === sourceMeta.title,
        'mutating the AST would leak one generation\'s metadata into every later use of the same AST'));

    // A second generation with no overrides must see the original values again.
    const { value: secondPass } = await astPure.to('html');
    results.push(mk('html', 'a later generation is unaffected by an earlier override',
        'Original Title', secondPass.includes('Original Title') ? 'Original Title' : '(missing)',
        secondPass.includes('Original Title') && !secondPass.includes('Temporary'),
        'the practical consequence of AST mutation, asserted directly'));

    // --- Reaches every generator that embeds metadata ---
    const overrides = { metadataOverrides: { title: 'Pinned Title', author: 'Pinned Author' } } as any;
    const { value: mdOut } = await makeAst().to('md', overrides);
    results.push(mk('md', 'override reaches YAML frontmatter',
        'title: "Pinned Title"', mdOut.split('\n').find((l: string) => l.startsWith('title:')) ?? '(none)',
        mdOut.includes('"Pinned Title"'),
        'Markdown embeds metadata as frontmatter'));

    const { value: rtfOut } = await makeAst().to('rtf', { ...overrides, renderMetadata: true });
    results.push(mk('rtf', 'override reaches the \\info group',
        '\\title Pinned Title', rtfOut.includes('Pinned Title') ? 'present' : '(missing)',
        rtfOut.includes('Pinned Title'),
        'RTF embeds metadata in \\info'));

    const { value: textOut } = await makeAst().to('text', { ...overrides, renderMetadata: true });
    results.push(mk('text', 'override reaches rendered metadata header',
        'Title: Pinned Title', textOut.split('\n')[0],
        textOut.includes('Pinned Title'),
        'renderMetadata writes metadata as visible content; it must reflect overrides too'));

    const epubBytes = (await makeAst().to('epub', overrides)).value as Uint8Array;
    const opf = strFromU8(unzipSync(epubBytes)['OEBPS/content.opf']);
    results.push(mk('epub', 'override reaches the OPF Dublin Core metadata',
        '<dc:title>Pinned Title</dc:title>', opf.match(/<dc:title>[^<]*<\/dc:title>/)?.[0] ?? '(none)',
        opf.includes('<dc:title>Pinned Title</dc:title>'),
        'EPUB embeds metadata in the OPF package document'));

    // --- language is representable in EPUB even though OfficeMetadata has no such field ---
    const epubLang = (await makeAst().to('epub', { metadataOverrides: { language: 'de-DE' } } as any)).value as Uint8Array;
    const langOpf = strFromU8(unzipSync(epubLang)['OEBPS/content.opf']);
    results.push(mk('epub', 'language override reaches dc:language',
        '<dc:language>de-DE</dc:language>', langOpf.match(/<dc:language>[^<]*<\/dc:language>/)?.[0] ?? '(none)',
        langOpf.includes('<dc:language>de-DE</dc:language>'),
        'language lives in nativeProperties rather than on OfficeMetadata; the override must still land'));

    // --- custom entries: written where representable ---
    const customCfg = { metadataOverrides: { custom: { department: 'Finance' } } } as any;
    const { value: customHtml } = await makeAst().to('html', customCfg);
    results.push(mk('html', 'custom entries are written as meta tags',
        'custom:department', customHtml.match(/<meta name="custom:department"[^>]*>/)?.[0] ?? '(none)',
        customHtml.includes('custom:department') && customHtml.includes('Finance'),
        'HTML meta tags are an open vocabulary'));

    // --- ...and reported, not silently dropped, where they are not ---
    for (const [fmt, label] of [['epub', 'EPUB'], ['rtf', 'RTF']] as const) {
        const warnings: any[] = [];
        await makeAst().to(fmt as any, { ...customCfg, renderMetadata: true, onWarning: (w: any) => warnings.push(w) });
        const warned = warnings.some(w => w.code === OfficeWarningType.METADATA_NOT_REPRESENTABLE);
        results.push(mk(fmt, `unrepresentable custom entries warn instead of vanishing`,
            'METADATA_NOT_REPRESENTABLE warning', warned ? 'warned' : `no such warning (got ${warnings.length})`,
            warned,
            `${label} has a fixed metadata vocabulary; a caller whose metadata silently never appears has no way to find out`));
    }

    // A caller who set no custom entries must not be warned at.
    const quietWarnings: any[] = [];
    await makeAst().to('epub', { metadataOverrides: { title: 'X' }, onWarning: (w: any) => quietWarnings.push(w) } as any);
    results.push(mk('epub', 'named-only overrides produce no warning',
        'no METADATA_NOT_REPRESENTABLE', quietWarnings.filter(w => w.code === OfficeWarningType.METADATA_NOT_REPRESENTABLE).length,
        !quietWarnings.some(w => w.code === OfficeWarningType.METADATA_NOT_REPRESENTABLE),
        'named fields map onto every metadata-bearing format, so warning about them would be noise'));

    return results;
}

/**
 * Regression tests for EPUB reproducibility.
 *
 * Generating the same AST twice used to produce archives that differed byte-for-byte, from two
 * independent wall-clock sources: `dcterms:modified` in the OPF, and - less obviously - the mtime
 * fflate stamps on every zip entry when none is supplied. Because DOS zip timestamps have
 * two-second granularity, back-to-back generation *looks* stable, so these tests deliberately
 * assert across a gap wider than that; a naive same-second test passes even when fully broken.
 *
 * This matters beyond tidiness: while the baselines churned on every run they were reverted
 * wholesale during review, which meant a real EPUB content regression would have been thrown out
 * with the noise.
 */
async function runEpubDeterminismTests(): Promise<GenFeatureTest[]> {
    const results: GenFeatureTest[] = [];
    const category = 'EPUB Determinism';
    const mk = (feature: string, exp: any, act: any, ok: boolean, details: string): GenFeatureTest => ({
        category, feature, sourceFormat: 'docx', destFormat: 'epub',
        result: { status: ok ? 'PASS' : 'FAIL', expected: exp, actual: act, details, duration: 0 }
    });
    const bytesEqual = (a: Uint8Array, b: Uint8Array) =>
        a.length === b.length && a.every((v, i) => v === b[i]);
    /** Longer than the two-second DOS timestamp quantum, or the assertion proves nothing. */
    const overOneZipTick = () => new Promise(r => setTimeout(r, 2500));

    const srcPath = getSourcePath('docx');
    if (!fs.existsSync(srcPath)) {
        return [{
            category, feature: 'byte-identical across runs', sourceFormat: 'docx', destFormat: 'epub',
            result: { status: 'SKIP', expected: '', actual: '', details: 'no docx source file', duration: 0 }
        }];
    }
    const ast = await OfficeParser.parseOffice(srcPath, PARSER_CONFIG);

    // Explicitly pinned: the reproducible-build contract callers rely on. The other hand-built zip
    // generators (DOCX, ODT) make the same reproducibility claim, so check them across the same tick.
    const pinned = { metadataOverrides: { modified: BASELINE_MODIFIED } } as any;
    const p1 = (await OfficeGenerator.generate(ast as any, 'epub', pinned)).value as Uint8Array;
    const dx1 = (await OfficeGenerator.generate(ast as any, 'docx', pinned)).value as Uint8Array;
    const ot1 = (await OfficeGenerator.generate(ast as any, 'odt', pinned)).value as Uint8Array;
    const tx1 = (await OfficeGenerator.generate(ast as any, 'tex', { ...pinned, texConfig: { bundle: true } } as any)).value as Uint8Array;
    await overOneZipTick();
    const p2 = (await OfficeGenerator.generate(ast as any, 'epub', pinned)).value as Uint8Array;
    const dx2 = (await OfficeGenerator.generate(ast as any, 'docx', pinned)).value as Uint8Array;
    const ot2 = (await OfficeGenerator.generate(ast as any, 'odt', pinned)).value as Uint8Array;
    const tx2 = (await OfficeGenerator.generate(ast as any, 'tex', { ...pinned, texConfig: { bundle: true } } as any)).value as Uint8Array;
    results.push(mk('pinned metadataOverrides.modified is byte-identical across runs',
        'identical bytes', bytesEqual(p1, p2) ? 'identical' : `differ (${p1.length} vs ${p2.length} bytes)`,
        bytesEqual(p1, p2),
        'an explicit modified timestamp must fully determine the archive, including zip entry mtimes'));
    for (const [dest, a, b] of [['docx', dx1, dx2], ['odt', ot1, ot2], ['tex', tx1, tx2]] as [string, Uint8Array, Uint8Array][]) {
        results.push({
            category, feature: `${dest} pinned modified is byte-identical across runs`, sourceFormat: 'docx', destFormat: dest,
            result: { status: bytesEqual(a, b) ? 'PASS' : 'FAIL', expected: 'identical bytes', actual: bytesEqual(a, b) ? 'identical' : `differ (${a.length} vs ${b.length})`, details: 'pinned mtimes + deterministic ids must fully determine the package', duration: 0 }
        });
    }

    // Unpinned but the document carries its own modification date: must also be reproducible,
    // since that is what every baseline regeneration relies on.
    const d1 = (await OfficeGenerator.generate(ast as any, 'epub')).value as Uint8Array;
    await overOneZipTick();
    const d2 = (await OfficeGenerator.generate(ast as any, 'epub')).value as Uint8Array;
    results.push(mk('ast.metadata.modified makes output reproducible without config',
        'identical bytes', bytesEqual(d1, d2) ? 'identical' : `differ (${d1.length} vs ${d2.length} bytes)`,
        bytesEqual(d1, d2),
        'a document with its own modification date should not need generator config to be reproducible'));

    // The pinned value must actually reach the OPF, not merely be accepted and ignored.
    const opf = strFromU8(unzipSync(p1)['OEBPS/content.opf']);
    const expectedIso = '2024-01-01T00:00:00Z';
    results.push(mk('pinned timestamp is written to dcterms:modified',
        expectedIso, opf.match(/dcterms:modified">([^<]*)</)?.[1] ?? '(absent)',
        opf.includes(`<meta property="dcterms:modified">${expectedIso}</meta>`),
        'EPUB 3 requires dcterms:modified; a config that silently failed to apply would still look deterministic'));

    // A date outside zip's representable 1980-2099 window must clamp, not throw: fflate rejects
    // it outright, and an epoch-zero mtime on a source document is an ordinary occurrence.
    let outOfRangeOk = true;
    let outOfRangeDetail = 'generated without throwing';
    try {
        const ancient = { metadataOverrides: { modified: new Date('1601-01-01T00:00:00Z') } } as any;
        const bytes = (await OfficeGenerator.generate(ast as any, 'epub', ancient)).value as Uint8Array;
        outOfRangeOk = bytes.length > 0;
    } catch (err: any) {
        outOfRangeOk = false;
        outOfRangeDetail = `threw: ${err?.message}`;
    }
    results.push(mk('a pre-1980 modified date clamps instead of throwing',
        'generates successfully', outOfRangeDetail, outOfRangeOk,
        'zip DOS timestamps cannot represent dates outside 1980-2099 and fflate throws rather than clamping'));

    return results;
}

/**
 * Regression tests for a reported bug: `.to('text')`/`.to('md')` stripped a document's genuine
 * leading whitespace via an unconditional `.trim()` on the whole accumulated output. Fixed by
 * trimming only the trailing-newline artifact the block-joining logic produces (TextGenerator.ts /
 * MarkdownGenerator.ts), not the whole string. These tests assert the resulting `.to('text')`
 * behavior directly.
 */
async function runWhitespaceFidelityTests(): Promise<GenFeatureTest[]> {
    const results: GenFeatureTest[] = [];
    const category = 'Whitespace Fidelity';

    const mk = (destFormat: string, feature: string, exp: any, act: any, ok: boolean, details: string): GenFeatureTest => ({
        category, feature, sourceFormat: 'synthetic', destFormat,
        result: { status: ok ? 'PASS' : 'FAIL', expected: exp, actual: act, details, duration: 0 }
    });

    const leadingWhitespaceText = '   Leading spaces from an intentionally-indented opening line';
    const content: OfficeContentNode[] = [
        { type: 'paragraph', children: [{ type: 'text', text: leadingWhitespaceText }] } as OfficeContentNode,
    ];
    const parserConfig: OfficeParserConfig = { newlineDelimiter: '\n' };
    const ast = createAST('docx', {}, content, [], parserConfig, undefined);

    // --- .to('text') emits the paragraph's text verbatim, leading whitespace and all ---
    const { value: newText } = await ast.to('text');
    results.push(mk('text', '.to(text) emits the block text verbatim', JSON.stringify(leadingWhitespaceText), JSON.stringify(newText), newText === leadingWhitespaceText, 'a single-paragraph document should render to exactly its text with no added or stripped characters'));
    results.push(mk('text', '.to(text) preserves leading whitespace', 'starts with "   "', newText.startsWith('   '), newText.startsWith('   '), 'a full trim() would silently strip this as if it were a generation artifact'));

    // --- .to('md') writes the paragraph's text byte-for-byte, without the spaces before it: a Markdown
    // reader drops a paragraph's leading spaces (and four or more would make it a code block), so
    // writing them only made the next save differ. Plain text, above, keeps them. ---
    const { value: mdValue } = await ast.to('md', { renderMetadata: false } as any);
    const mdText = leadingWhitespaceText.replace(/^ +/, '');
    results.push(mk('md', '.to(md) writes the paragraph text without its leading spaces', JSON.stringify(mdText), JSON.stringify(mdValue), mdValue === mdText, 'the text itself survives byte-for-byte; the spaces before it are not written, as a Markdown reader drops them'));

    // --- Trailing-newline cleanup must still work (the actual original purpose of the trim) ---
    const noTrailingJunkContent: OfficeContentNode[] = [
        { type: 'paragraph', children: [{ type: 'text', text: 'No trailing junk please' }] } as OfficeContentNode,
    ];
    const cleanAst = createAST('docx', {}, noTrailingJunkContent, [], parserConfig, undefined);
    const { value: cleanText } = await cleanAst.to('text');
    results.push(mk('text', '.to(text) still strips the block-joiner\'s trailing newline', 'no trailing newline', !cleanText.endsWith('\n'), !cleanText.endsWith('\n'), 'trimEnd() must still clean up the artifact trim() was originally added for'));

    // --- Genuine trailing whitespace that ISN'T the delimiter (e.g. authored trailing spaces) is
    // real content and must survive too - only a run of the exact delimiter is a safe strip ---
    const trailingSpacesContent: OfficeContentNode[] = [
        { type: 'paragraph', children: [{ type: 'text', text: 'Trailing spaces are content   ' }] } as OfficeContentNode,
    ];
    const trailingSpacesAst = createAST('docx', {}, trailingSpacesContent, [], parserConfig, undefined);
    const { value: trailingSpacesValue } = await trailingSpacesAst.to('text');
    results.push(mk('text', '.to(text) preserves genuine trailing spaces', 'ends with "   "', trailingSpacesValue.endsWith('   '), trailingSpacesValue.endsWith('   '), 'only a run of the exact newline delimiter is a generator artifact - other trailing whitespace is real content'));

    // --- Issue #102: a whitespace-ONLY block is content, not an empty block ---
    // The reported symptom was ".to('text') strips leading whitespace" for a Word document that
    // begins with whitespace. In Word that is almost always a leading blank-but-not-empty
    // paragraph, which is a different code path from an indented first line: the generator
    // discarded any block whose children trimmed to '', when it should keep any block whose text is
    // not the empty string. Both readings of the report are covered here.
    const expectedText = (content: OfficeContentNode[]) => content
        .map(c => (c.children || []).map(k => k.text || '').join(''))
        .filter(t => t != '')
        .join('\n');

    const issue102Cases: Array<[string, OfficeContentNode[]]> = [
        ['leading whitespace-only paragraph', [
            { type: 'paragraph', children: [{ type: 'text', text: '   ' }] } as OfficeContentNode,
            { type: 'paragraph', children: [{ type: 'text', text: 'Hello' }] } as OfficeContentNode,
        ]],
        ['whitespace-only paragraph mid-document', [
            { type: 'paragraph', children: [{ type: 'text', text: 'A' }] } as OfficeContentNode,
            { type: 'paragraph', children: [{ type: 'text', text: '  ' }] } as OfficeContentNode,
            { type: 'paragraph', children: [{ type: 'text', text: 'B' }] } as OfficeContentNode,
        ]],
        // A genuinely empty paragraph must still be dropped - that is the behaviour the
        // trimmed-emptiness check was there for, and both paths already agreed on it.
        ['genuinely empty paragraph still dropped', [
            { type: 'paragraph', children: [{ type: 'text', text: '' }] } as OfficeContentNode,
            { type: 'paragraph', children: [{ type: 'text', text: 'Hello' }] } as OfficeContentNode,
        ]],
    ];

    for (const [label, content] of issue102Cases) {
        const caseAst = createAST('docx', {}, content, [], parserConfig, undefined);
        const expected = expectedText(content);
        const { value: modern } = await caseAst.to('text');
        results.push(mk('text', `#102: ${label}`, JSON.stringify(expected), JSON.stringify(modern), modern === expected,
            'a whitespace-only block is content and must survive; a genuinely empty block must still be dropped'));
    }

    // --- Nodes that carry their content in `text` with no children must not vanish ---
    // TextGenerator rendered only `childrenOutput`, so a `chart` (whose entire data series lives in
    // `text`) and a CSV `comment` were emitted as nothing at all. The fix reads node.text as a
    // fallback, generically rather than per-type, so assert it that way: any leaf-with-text must
    // survive, not just the two we happened to find.
    const leafTextTypes: Array<OfficeContentNodeType> = ['chart', 'comment'];
    for (const leafType of leafTextTypes) {
        const marker = `SERIES-${leafType.toUpperCase()}-42`;
        const leafAst = createAST('docx', {}, [{ type: leafType, text: marker } as OfficeContentNode],
            [], parserConfig, undefined);
        const { value: leafOut } = await leafAst.to('text');
        results.push(mk('text', `leaf '${leafType}' node text is not dropped`,
            marker, JSON.stringify(leafOut), leafOut.includes(marker),
            'a node carrying its content in text with no children must not be rendered as empty'));
    }
    // Children still win when both are present - the fallback must not double-render.
    const bothAst = createAST('docx', {}, [{
        type: 'chart', text: 'RAW-FALLBACK',
        children: [{ type: 'text', text: 'CHILD-TEXT' }],
    } as OfficeContentNode], [], parserConfig, undefined);
    const { value: bothOut } = await bothAst.to('text');
    results.push(mk('text', 'node text is a fallback, not an addition',
        'CHILD-TEXT only', JSON.stringify(bothOut),
        bothOut.includes('CHILD-TEXT') && !bothOut.includes('RAW-FALLBACK'),
        'rendered children are the richer representation; emitting both would duplicate the content'));

    // --- Cells must be separated, not concatenated ---
    // Without an explicit separator "ITEM" + "NEEDED" renders as "ITEMNEEDED" and the cell boundary
    // is gone. This was invisible for XLSX/ODS only because their cell values carry a trailing
    // non-breaking space in the source data - the delimiter came from the document, not from us,
    // so any format without that trailing whitespace (MD, HTML) collided outright.
    const cellRowContent: OfficeContentNode[] = [
        {
            type: 'table',
            children: [{
                type: 'row',
                children: [
                    { type: 'cell', children: [{ type: 'text', text: 'ITEM' }] },
                    { type: 'cell', children: [{ type: 'text', text: 'NEEDED' }] },
                ],
            } as OfficeContentNode],
        } as OfficeContentNode,
    ];
    const cellAst = createAST('docx', undefined as any, cellRowContent, [], parserConfig, undefined, () => '');
    const { value: flatCells } = await cellAst.to('text', { textConfig: { preserveLayout: false } } as any);
    results.push(mk('text', 'cells are separated when not rendering a grid',
        'ITEM and NEEDED separated', JSON.stringify(flatCells),
        !flatCells.includes('ITEMNEEDED') && /ITEM\s+NEEDED/.test(flatCells),
        'preserveLayout:false must still delimit cells; concatenating them destroys the cell boundary'));
    // No trailing delimiter: it is an artifact of appending one per cell, not content.
    results.push(mk('text', 'no trailing cell separator on a row',
        'row does not end with a tab', JSON.stringify(flatCells),
        !flatCells.split('\n').some(l => l.endsWith('\t')),
        'the separator after the final cell has nothing to separate from'));
    // The aligned-grid path must be untouched by the separator change.
    const { value: gridCells } = await cellAst.to('text', { textConfig: { preserveLayout: true } } as any);
    results.push(mk('text', 'grid rendering unaffected by the cell separator',
        'renders | ITEM | NEEDED |', JSON.stringify(gridCells),
        gridCells.includes('| ITEM') && gridCells.includes('| NEEDED') && !gridCells.includes('ITEM\t'),
        'preserveLayout:true routes through renderTable, which does its own column alignment'));

    // --- A table as the only/first top-level node: renderTable (text) and the HTML-fallback path
    // of renderMarkdownTable (md) both unconditionally wrap in a leading+trailing newline as
    // separators from siblings - with no preceding/following sibling, that leaks as a spurious
    // leading AND trailing artifact unless stripped from both ends, not just the end ---
    const tableOnlyContent: OfficeContentNode[] = [
        {
            type: 'table',
            children: [{ type: 'row', children: [{ type: 'cell', children: [{ type: 'text', text: 'Cell' }] }] } as OfficeContentNode],
        } as OfficeContentNode,
    ];
    const tableOnlyAst = createAST('docx', undefined as any, tableOnlyContent, [], parserConfig, undefined, () => '');
    const { value: tableOnlyText } = await tableOnlyAst.to('text', { textConfig: { preserveLayout: true } } as any);
    results.push(mk('text', '.to(text) strips renderTable\'s leading separator artifact too', 'does not start with "\\n"', !tableOnlyText.startsWith('\n'), !tableOnlyText.startsWith('\n'), 'renderTable seeds a leading newline as a separator from a preceding sibling - with none, it must not leak through'));

    const { value: tableOnlyMd } = await tableOnlyAst.to('md', { mdConfig: { dialect: 'commonmark' } } as any);
    results.push(mk('md', '.to(md) strips the HTML-fallback table\'s leading separator artifact too', 'does not start with "\\n"', typeof tableOnlyMd === 'string' && !tableOnlyMd.startsWith('\n'), typeof tableOnlyMd === 'string' && !tableOnlyMd.startsWith('\n'), 'commonmark forces tables through the HTML-fallback branch, which has the same leading-separator convention as TextGenerator\'s renderTable'));

    return results;
}

// ============================================================================
// PRE-MERGE REVIEW REGRESSIONS (v8.0.0)
// ============================================================================

/** Concatenated text of a node's whole subtree, for asserting on a re-parsed document. */
function subtreeText(node: OfficeContentNode | undefined): string {
    if (!node) return '';
    let out = node.type === 'text' ? (node.text || '') : '';
    if (!node.children?.length && node.text && node.type !== 'text') out += node.text;
    for (const c of node.children || []) out += subtreeText(c);
    return out;
}

/** The first `table` node anywhere in a parsed document. */
function firstTable(nodes: OfficeContentNode[] | undefined): OfficeContentNode | null {
    for (const n of nodes || []) {
        if (n.type === 'table') return n;
        const found = firstTable(n.children);
        if (found) return found;
    }
    return null;
}

/** The rendered cell strings of a table, row by row. */
const tableGrid = (table: OfficeContentNode | null): string[][] =>
    (table?.children || []).filter(r => r.type === 'row').map(r => (r.children || []).filter(c => c.type === 'cell').map(subtreeText));

/** A 1x1 GIF: a real image pdf-lib cannot embed (it handles PNG and JPEG only). */
const GIF_1PX = 'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';
/** A 1x1 PNG. */
const PNG_1PX = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const synthAst = (type: string, content: OfficeContentNode[], attachments: any[] = []) =>
    createAST(type as any, {}, content, attachments, { newlineDelimiter: '\n' }, undefined, () => '');

/**
 * Regressions for the native PDF engine (`pdfConfig.engine: 'native'`), which lays a PDF out itself
 * with pdf-lib instead of printing HTML through a headless browser. Each case generates a PDF, parses
 * it back, and asserts on the text that comes out - the engine paints pixels, so the only honest way
 * to test it is to read the result back.
 */
async function runNativePdfEngineTests(): Promise<GenFeatureTest[]> {
    const results: GenFeatureTest[] = [];
    const category = 'Native PDF Engine';
    const mk = (feature: string, exp: any, act: any, ok: boolean, details: string): GenFeatureTest => ({
        category, feature, sourceFormat: 'synthetic', destFormat: 'pdf',
        result: { status: ok ? 'PASS' : 'FAIL', expected: exp, actual: act, details, duration: 0 }
    });
    const nativeText = async (ast: any, config: any = {}): Promise<{ text: string; messages: any[] }> => {
        const res = await OfficeGenerator.generate(ast, 'pdf' as any, { ...config, pdfConfig: { engine: 'native' } } as any);
        const back = await OfficeParser.parseOffice(Buffer.from(res.value as Uint8Array), { ...PARSER_CONFIG, ocr: false });
        return { text: String((await (back as any).to('text')).value), messages: (res as any).messages || [] };
    };

    // --- Cell text comes from the cell's children, not only from `cell.text` ---
    // Markdown and HTML build cells as `{ type:'cell', children:[...] }` with no `text` of their own,
    // so an engine reading `cell.text` painted every table from those sources as an empty grid. The
    // DOCX/ODT/XLSX/PDF parsers all set `text`, which is why no fixture caught it.
    const mdCellAst = synthAst('md', [{
        type: 'table', children: [
            { type: 'row', children: [{ type: 'cell', children: [{ type: 'text', text: 'HeadOne' }] }, { type: 'cell', children: [{ type: 'text', text: 'HeadTwo' }] }] },
            { type: 'row', children: [{ type: 'cell', children: [{ type: 'paragraph', children: [{ type: 'text', text: 'CellOne' }] }] }, { type: 'cell', children: [{ type: 'text', text: 'CellTwo' }] }] },
        ],
    } as OfficeContentNode]);
    const cellOut = await nativeText(mdCellAst);
    results.push(mk('table cells with no `text` (Markdown/HTML shape) still render',
        'HeadOne/CellOne/CellTwo present', JSON.stringify(cellOut.text).slice(0, 160),
        ['HeadOne', 'HeadTwo', 'CellOne', 'CellTwo'].every(t => cellOut.text.includes(t)),
        'a Markdown table converted with engine:native produced a table of empty boxes'));

    // --- A code block keeps one line per source line, with its indentation ---
    // The WinAnsi mapping replaces every newline with a space, so splitting AFTER encoding collapsed
    // the whole block onto one wrapped line.
    const codeAst = synthAst('md', [{ type: 'code', text: 'line one\n    line two indented\nline three' } as OfficeContentNode]);
    const codeOut = await nativeText(codeAst);
    const codeLines = codeOut.text.split('\n').filter(l => l.trim());
    results.push(mk('a fenced code block keeps its lines',
        '3 lines', `${codeLines.length}: ${JSON.stringify(codeOut.text)}`,
        codeLines.length === 3 && codeOut.text.includes('line three'),
        'encoding before splitting turned every newline into a space, so the block rendered as one line'));
    results.push(mk('a code line keeps its leading indentation',
        'indented second line', JSON.stringify(codeLines[1] ?? ''),
        /^\s+line two indented/.test(codeLines[1] ?? ''),
        'leading whitespace is the whole point of drawing code lines verbatim in a monospace font'));

    // --- A multi-line cell keeps its lines too (same collapse, via wrapPlain) ---
    const multiLineCell = synthAst('docx', [{
        type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [
            { type: 'paragraph', children: [{ type: 'text', text: 'CellLineA' }] },
            { type: 'paragraph', children: [{ type: 'text', text: 'CellLineB' }] },
        ] }] }],
    } as OfficeContentNode]);
    const multiOut = await nativeText(multiLineCell);
    results.push(mk('a cell with two paragraphs renders two lines',
        'CellLineA and CellLineB on separate lines', JSON.stringify(multiOut.text),
        /CellLineA[\s\S]*\n[\s\S]*CellLineB/.test(multiOut.text),
        'block children of a cell each own a line; joining them would run the words together'));

    // --- A row taller than a page is split across pages, not drawn off the bottom (data loss) ---
    // A cell of 120 short paragraphs is far taller than one page. The engine used to draw the whole
    // row from the top down with no bottom check, so every line past the margin was silently lost.
    const tallRows = Array.from({ length: 120 }, (_, i) => ({ type: 'paragraph', children: [{ type: 'text', text: `ROWLINE${i}` }] }));
    const tallTable = synthAst('docx', [{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: tallRows }] }] } as OfficeContentNode]);
    const tallOut = await nativeText(tallTable);
    const survived = Array.from({ length: 120 }, (_, i) => `ROWLINE${i}`).filter(l => tallOut.text.includes(l)).length;
    results.push(mk('an over-tall table row is paginated, not truncated',
        'all 120 cell lines survive', `${survived}/120 survived`,
        survived === 120,
        'a row taller than the page was drawn off the bottom, silently losing the overflow text'));

    // --- onNode reaches text runs, which is where nearly all the text lives ---
    let runCalls = 0;
    const hookAst = synthAst('md', [{ type: 'paragraph', children: [{ type: 'text', text: 'Original words' }] } as OfficeContentNode]);
    const hookOut = await nativeText(hookAst, { onNode: (n: OfficeContentNode) => { if (n.type === 'text') { runCalls++; return 'OVERRIDDEN'; } return undefined; } });
    results.push(mk('onNode fires for text runs and its string override is drawn',
        '1 call, output contains OVERRIDDEN', `${runCalls} call(s), ${JSON.stringify(hookOut.text)}`,
        runCalls === 1 && hookOut.text.includes('OVERRIDDEN') && !hookOut.text.includes('Original'),
        'CommonGeneratorConfig.onNode promises every node; the engine flattened runs without ever asking'));

    // --- An image the engine cannot embed degrades to its alt text, and says so ---
    const gifAst = synthAst('docx', [{ type: 'image', metadata: { attachmentName: 'pic.gif', altText: 'A described picture' } } as OfficeContentNode],
        [{ name: 'pic.gif', extension: 'gif', mimeType: 'image/gif', data: GIF_1PX, size: 40 }]);
    const gifOut = await nativeText(gifAst);
    results.push(mk('an unembeddable image (GIF) falls back to its alt text',
        'alt text drawn', JSON.stringify(gifOut.text).slice(0, 120),
        gifOut.text.includes('A described picture'),
        'pdf-lib embeds PNG/JPEG only; a bare catch drew nothing at all for every other format'));
    results.push(mk('an unembeddable image warns (IMAGE_PROCESSING_FAILED)',
        'IMAGE_PROCESSING_FAILED', JSON.stringify(gifOut.messages.map((m: any) => m.code)),
        gifOut.messages.some((m: any) => m.code === OfficeWarningType.IMAGE_PROCESSING_FAILED),
        'DOCX and ODT warn and degrade for the same input; a silently vanishing picture is unfindable'));

    const missingAst = synthAst('docx', [{ type: 'image', metadata: { attachmentName: 'gone.png', altText: 'Missing picture alt' } } as OfficeContentNode]);
    const missingOut = await nativeText(missingAst);
    results.push(mk('a missing attachment falls back to alt text and warns',
        'alt text + IMAGE_PROCESSING_FAILED', `${missingOut.text.includes('Missing picture alt')}, ${JSON.stringify(missingOut.messages.map((m: any) => m.code))}`,
        missingOut.text.includes('Missing picture alt') && missingOut.messages.some((m: any) => m.code === OfficeWarningType.IMAGE_PROCESSING_FAILED),
        'an unresolvable attachment name is the other half of the same silent-drop path'));

    return results;
}

/**
 * Regressions for table geometry: the sparse grids a spreadsheet parser emits, and header rows.
 *
 * `ExcelParser` emits only the non-empty cells (each carrying its absolute `{ row, col }`) and only
 * the non-empty rows, so a writer that places cells at a running cursor silently moves the data.
 * Header rows have the mirror problem: they are written in a form the reader then does not read back.
 */
async function runTableGeometryTests(): Promise<GenFeatureTest[]> {
    const results: GenFeatureTest[] = [];
    const category = 'Table Geometry';
    const mk = (destFormat: string, feature: string, exp: any, act: any, ok: boolean, details: string): GenFeatureTest => ({
        category, feature, sourceFormat: 'synthetic', destFormat,
        result: { status: ok ? 'PASS' : 'FAIL', expected: exp, actual: act, details, duration: 0 }
    });
    const roundTrip = async (ast: any, fmt: 'docx' | 'odt') => {
        const res = await OfficeGenerator.generate(ast, fmt as any, deterministicConfigFor(fmt) as any);
        const tmp = path.join(__dirname, 'output', `_regression_tmp.${fmt}`);
        fs.writeFileSync(tmp, res.value as Uint8Array);
        try {
            return await OfficeParser.parseOffice(tmp, PARSER_CONFIG);
        } finally {
            if (fs.existsSync(tmp)) fs.unlinkSync(tmp);
        }
    };

    // --- A sparse sheet (values in A1 and D1, nothing in row 2) keeps its coordinates ---
    const sparseCell = (row: number, col: number, text: string) =>
        ({ type: 'cell', text, metadata: { row, col }, children: [{ type: 'text', text }] } as OfficeContentNode);
    const sparseAst = () => synthAst('xlsx', [{
        type: 'sheet', metadata: { sheetName: 'Sheet1' }, children: [
            { type: 'row', children: [sparseCell(0, 0, 'A1VAL'), sparseCell(0, 3, 'D1VAL')] },
            { type: 'row', children: [sparseCell(2, 0, 'A3VAL')] },
        ],
    } as OfficeContentNode]);

    for (const fmt of ['docx', 'odt'] as const) {
        const grid = tableGrid(firstTable((await roundTrip(sparseAst(), fmt)).content));
        const firstRow = grid[0] || [];
        results.push(mk(fmt, 'a value in column D stays in column D',
            'A1VAL at index 0, D1VAL at index 3', JSON.stringify(firstRow),
            firstRow[0] === 'A1VAL' && firstRow[3] === 'D1VAL',
            'cells were placed at a running cursor, so a sparse row collapsed into adjacent columns'));
        results.push(mk(fmt, 'an empty spreadsheet row is not dropped',
            '3 rows with A3VAL last', `${grid.length} rows: ${JSON.stringify(grid.map(r => r[0]))}`,
            grid.length === 3 && (grid[2] || [])[0] === 'A3VAL',
            'only non-empty rows reach the AST; dropping the gap moves every row below it up'));
    }

    // --- A header row survives the library's own round trip, in both packaged formats ---
    const headerCell = (t: string) => ({ type: 'cell', text: t, children: [{ type: 'text', text: t, formatting: { bold: true } }] } as OfficeContentNode);
    const bodyCell = (t: string) => ({ type: 'cell', text: t, children: [{ type: 'text', text: t }] } as OfficeContentNode);
    const headerAst = () => synthAst('docx', [{
        type: 'table', children: [
            { type: 'row', children: [headerCell('HeadOne'), headerCell('HeadTwo')] },
            { type: 'row', children: [bodyCell('bodyOne'), bodyCell('bodyTwo')] },
        ],
    } as OfficeContentNode]);

    for (const fmt of ['docx', 'odt'] as const) {
        const table = firstTable((await roundTrip(headerAst(), fmt)).content);
        const grid = tableGrid(table);
        const rows = (table?.children || []).filter(r => r.type === 'row');
        results.push(mk(fmt, 'a header row survives the round trip',
            '2 rows, header first', JSON.stringify(grid),
            grid.length === 2 && (grid[0] || [])[0] === 'HeadOne',
            fmt === 'odt'
                ? 'the generator wraps header rows in <table:table-header-rows>, which the reader skipped entirely - losing the row, including from real LibreOffice files'
                : 'the header row must not be dropped or reordered by the <w:tblHeader/> marker'));
        results.push(mk(fmt, 'the re-parsed header row is still recognized as a header',
            'isHeaderRow(row 0) true, row 1 false', `${isHeaderRow(rows[0], true)}, ${isHeaderRow(rows[1], false)}`,
            !!rows[0] && !!rows[1] && isHeaderRow(rows[0], true) && !isHeaderRow(rows[1], false),
            'the header marker each format writes (<w:tblHeader/>, <table:table-header-rows>) has to be read back, or the header is only ever recovered when it happens to be all-bold'));
    }

    // --- onNode fires exactly once per node in HTML, including in the table shapes that
    // re-render their rows (a bold-header table, a rowspan grid, a spreadsheet sheet) ---
    const countOnNode = async (content: OfficeContentNode[]): Promise<Record<string, number>> => {
        const counts: Record<string, number> = {};
        await OfficeGenerator.generate(synthAst('docx', content) as any, 'html' as any,
            { onNode: (n: OfficeContentNode) => { counts[n.type] = (counts[n.type] || 0) + 1; return undefined; } } as any);
        return counts;
    };
    const plainCounts = await countOnNode([{ type: 'paragraph', children: [{ type: 'text', text: 'Plain' }] } as OfficeContentNode]);
    results.push(mk('html', 'onNode fires once per node (plain paragraph)',
        '{paragraph:1, text:1}', JSON.stringify(plainCounts),
        plainCounts.paragraph === 1 && plainCounts.text === 1,
        'the pre-evaluated verdict was passed down as a bare value, and "no override" IS undefined - so the hook was asked again'));

    const headerCounts = await countOnNode([{
        type: 'table', children: [
            { type: 'row', children: [headerCell('H1'), headerCell('H2')] },
            { type: 'row', children: [bodyCell('b1'), bodyCell('b2')] },
        ],
    } as OfficeContentNode]);
    results.push(mk('html', 'onNode fires once per node (bold-header table)',
        '{table:1, row:2, cell:4, text:4}', JSON.stringify(headerCounts),
        headerCounts.table === 1 && headerCounts.row === 2 && headerCounts.cell === 4 && headerCounts.text === 4,
        'the header path re-renders the rows, so the generic children pass was pure waste and fired the hook a second time for every cell'));

    const rowspanCounts = await countOnNode([{
        type: 'table', children: [
            { type: 'row', children: [{ type: 'cell', metadata: { rowSpan: 2 }, children: [{ type: 'text', text: 'spanning' }] } as OfficeContentNode, bodyCell('x')] },
            { type: 'row', children: [bodyCell('y')] },
        ],
    } as OfficeContentNode]);
    results.push(mk('html', 'onNode fires once per node (rowspan table)',
        '{table:1, row:2, cell:3, text:3}', JSON.stringify(rowspanCounts),
        rowspanCounts.table === 1 && rowspanCounts.row === 2 && rowspanCounts.cell === 3 && rowspanCounts.text === 3,
        'the rowspan grid builds each <tr> itself, so the row nodes have to be offered to the hook there'));

    // --- A header row dropped by onNode must not leave a body row wrapped in <thead> ---
    // A rowspan in a later row forces the grid-occupancy path; the all-bold first row would normally
    // split off as <thead>. When onNode drops that header row, trs[0] shifts to a body row, which must
    // land in <tbody>, not <thead>.
    const dropHeaderAst = synthAst('docx', [{
        type: 'table', children: [
            { type: 'row', metadata: { drop: true } as any, children: [headerCell('DROPHEAD'), headerCell('H2')] },
            { type: 'row', children: [{ type: 'cell', metadata: { rowSpan: 2 }, children: [{ type: 'text', text: 'span' }] } as OfficeContentNode, bodyCell('bodyX')] },
            { type: 'row', children: [bodyCell('bodyY')] },
        ],
    } as OfficeContentNode]);
    const keepRes = await OfficeGenerator.generate(dropHeaderAst as any, 'html' as any, {} as any);
    const keepHtml = typeof keepRes.value === 'string' ? keepRes.value : '';
    const headerInThead = /<thead>[\s\S]*DROPHEAD[\s\S]*<\/thead>/.test(keepHtml);
    const dropRes = await OfficeGenerator.generate(dropHeaderAst as any, 'html' as any,
        { onNode: (n: OfficeContentNode) => (n.type === 'row' && (n.metadata as any)?.drop) ? false : undefined } as any);
    const dropHtml = typeof dropRes.value === 'string' ? dropRes.value : '';
    const headerDropped = !dropHtml.includes('DROPHEAD');
    const noThead = !dropHtml.includes('<thead>');
    results.push(mk('html', 'onNode-dropped header row does not wrap a body row in <thead>',
        'kept: header in <thead>; dropped: header gone, no <thead>',
        `headerInThead=${headerInThead}, headerDropped=${headerDropped}, noThead=${noThead}`,
        headerInThead && headerDropped && noThead,
        'renderRowsWithRowspans skips a dropped row, so trs[0] shifts to a body row; wrapping it in <thead> put a <td> body row in the header'));

    const sheetCounts = await countOnNode([{
        type: 'sheet', metadata: { sheetName: 'S' }, children: [
            { type: 'row', children: [sparseCell(0, 0, 'A1'), sparseCell(0, 1, 'B1')] },
        ],
    } as OfficeContentNode]);
    results.push(mk('html', 'onNode fires once per node (spreadsheet sheet)',
        '{sheet:1, row:1, cell:2, text:2}', JSON.stringify(sheetCounts),
        sheetCounts.sheet === 1 && sheetCounts.row === 1 && sheetCounts.cell === 2 && sheetCounts.text === 2,
        'the sheet branch rebuilds the grid from the cell indices and never reads the generic children output'));

    return results;
}

/**
 * Regressions for how an image, a note and OCR text degrade when the target format cannot carry them
 * whole: an image too large to inline, a picture the format cannot embed, a note that opens with a
 * table, and multi-line recognized text.
 */
async function runDegradeFidelityTests(): Promise<GenFeatureTest[]> {
    const results: GenFeatureTest[] = [];
    const category = 'Degrade Fidelity';
    const mk = (destFormat: string, feature: string, exp: any, act: any, ok: boolean, details: string): GenFeatureTest => ({
        category, feature, sourceFormat: 'synthetic', destFormat,
        result: { status: ok ? 'PASS' : 'FAIL', expected: exp, actual: act, details, duration: 0 }
    });

    // --- An image over maxInlineImageBytes renders its recognized text, per the documented contract ---
    const overCapAst = (ocr: string) => synthAst('docx',
        [{ type: 'image', text: ocr, metadata: { attachmentName: 'big.png', altText: 'Alt' } } as OfficeContentNode],
        [{ name: 'big.png', extension: 'png', mimeType: 'image/png', data: 'A'.repeat(4000), size: 3000 }]);
    const OCR_TEXT = 'Recognized line one\nRecognized line two';

    const overCapCfg = { maxInlineImageBytes: 100 } as any;
    const mdWithOcr = String((await OfficeGenerator.generate(overCapAst(OCR_TEXT) as any, 'md' as any, overCapCfg)).value);
    results.push(mk('md', 'an over-cap image renders its OCR text',
        'fenced OCR block, no ![...](big.png)', JSON.stringify(mdWithOcr),
        mdWithOcr.includes('Recognized line one') && mdWithOcr.includes('Recognized line two') && !mdWithOcr.includes('](big.png)'),
        'an image that cannot be inlined has no picture to show; its recognized text is the content that is left'));
    results.push(mk('md', 'multi-line OCR text keeps its layout (fenced)',
        'wrapped in ```', JSON.stringify(mdWithOcr.slice(0, 8)),
        mdWithOcr.trimStart().startsWith('```'),
        'OCR layout is column-aligned with spaces, which Markdown collapses outside a fence'));

    const mdNoOcr = String((await OfficeGenerator.generate(overCapAst('') as any, 'md' as any, overCapCfg)).value);
    results.push(mk('md', 'an over-cap image with no OCR text keeps the name reference',
        '![Alt](big.png)', JSON.stringify(mdNoOcr),
        mdNoOcr.includes('](big.png)'),
        'with no recognized text the compact reference is all there is to emit'));

    const capWarnings: any[] = [];
    await OfficeGenerator.generate(overCapAst(OCR_TEXT) as any, 'md' as any, { ...overCapCfg, onWarning: (w: any) => capWarnings.push(w) } as any);
    results.push(mk('md', 'the over-cap degrade still warns (IMAGE_NOT_INLINED)',
        'IMAGE_NOT_INLINED', JSON.stringify(capWarnings.map(w => w.code)),
        capWarnings.some(w => w.code === OfficeWarningType.IMAGE_NOT_INLINED),
        'the warning is how a caller finds out the picture did not travel with the output'));

    const textWithOcr = String((await OfficeGenerator.generate(overCapAst(OCR_TEXT) as any, 'text' as any, overCapCfg)).value);
    results.push(mk('text', 'an over-cap image renders its OCR text',
        'OCR text, no [Image: ...] placeholder', JSON.stringify(textWithOcr),
        textWithOcr.includes('Recognized line one') && !textWithOcr.includes('[Image:'),
        'for a scanned page the recognized text is the entire content; a placeholder line loses it'));
    const textNoOcr = String((await OfficeGenerator.generate(overCapAst('') as any, 'text' as any, overCapCfg)).value);
    results.push(mk('text', 'an over-cap image with no OCR text keeps the placeholder',
        '[Image: Alt]', JSON.stringify(textNoOcr),
        textNoOcr.includes('[Image: Alt]'),
        'the placeholder is the fallback, not the default'));

    // Under the cap nothing changes: the image still inlines (md) / is still named (text).
    const smallAst = synthAst('docx',
        [{ type: 'image', text: 'Small ocr', metadata: { attachmentName: 'small.png', altText: 'Alt' } } as OfficeContentNode],
        [{ name: 'small.png', extension: 'png', mimeType: 'image/png', data: PNG_1PX, size: 70 }]);
    const smallMd = String((await OfficeGenerator.generate(smallAst as any, 'md' as any, {} as any)).value);
    const smallText = String((await OfficeGenerator.generate(smallAst as any, 'text' as any, {} as any)).value);
    results.push(mk('md', 'an under-cap image still inlines as a data URI',
        'data:image/png;base64', smallMd.slice(0, 40),
        smallMd.includes('data:image/png;base64'),
        'the OCR fallback must apply only past the cap'));
    results.push(mk('text', 'an under-cap image still renders its placeholder',
        '[Image: Alt]', JSON.stringify(smallText),
        smallText.includes('[Image: Alt]'),
        'the OCR fallback must apply only past the cap'));

    // --- RTF: multi-line OCR text needs \line, not a raw newline ---
    const rtfOcrAst = synthAst('docx',
        [{ type: 'image', text: 'OCR line A\nOCR line B', metadata: { attachmentName: 'small.png' } } as OfficeContentNode],
        [{ name: 'small.png', extension: 'png', mimeType: 'image/png', data: PNG_1PX, size: 70 }]);
    const rtfOut = String((await OfficeGenerator.generate(rtfOcrAst as any, 'rtf' as any, { includeImages: 'image+ocr-text' } as any)).value);
    results.push(mk('rtf', 'multi-line OCR text keeps its line breaks',
        'OCR line A\\line OCR line B', JSON.stringify(rtfOut.slice(rtfOut.indexOf('OCR line A') - 2, rtfOut.indexOf('OCR line A') + 32)),
        rtfOut.includes('OCR line A\\line OCR line B'),
        'a raw newline is whitespace to an RTF reader, so a column-aligned OCR block collapsed to one run'));

    // --- DOCX: a note whose body starts with a table still opens with a paragraph ---
    const noteWithTable: OfficeContentNode = {
        type: 'note', metadata: { noteType: 'footnote' },
        children: [{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [{ type: 'text', text: 'in-note cell' }] }] }] }],
    } as OfficeContentNode;
    const noteAst = synthAst('docx', [{
        type: 'paragraph', children: [{ type: 'text', text: 'Body', notes: [noteWithTable] } as OfficeContentNode],
    } as OfficeContentNode]);
    const noteDocx = (await OfficeGenerator.generate(noteAst as any, 'docx' as any, deterministicConfigFor('docx') as any)).value as Uint8Array;
    const footnotesXml = strFromU8(unzipSync(noteDocx)['word/footnotes.xml']);
    const noteBody = /<w:footnote w:id="2">([\s\S]*?)<\/w:footnote>/.exec(footnotesXml)?.[1] || '';
    results.push(mk('docx', 'a footnote opening with a table starts with a paragraph',
        'starts with <w:p>', noteBody.slice(0, 48),
        noteBody.startsWith('<w:p>'),
        'a bare <w:r> is not valid block content under <w:footnote>; the marker needs a paragraph of its own'));

    // --- DOCX: an image that is itself a link keeps the link (ODT already did) ---
    const linkedImageAst = synthAst('docx',
        [{ type: 'image', metadata: { attachmentName: 'small.png', link: 'https://example.com/target', altText: 'Linked' } } as OfficeContentNode],
        [{ name: 'small.png', extension: 'png', mimeType: 'image/png', data: PNG_1PX, size: 70 }]);
    const linkedDocx = (await OfficeGenerator.generate(linkedImageAst as any, 'docx' as any, deterministicConfigFor('docx') as any)).value as Uint8Array;
    const linkedFiles = unzipSync(linkedDocx);
    const documentXml = strFromU8(linkedFiles['word/document.xml']);
    const relsXml = strFromU8(linkedFiles['word/_rels/document.xml.rels']);
    results.push(mk('docx', 'a linked image keeps its hyperlink',
        '<w:hyperlink> wrapping the drawing + an external rel', `${documentXml.includes('<w:hyperlink')}, ${relsXml.includes('example.com/target')}`,
        /<w:hyperlink r:id="[^"]+"><w:r><w:drawing>/.test(documentXml) && relsXml.includes('https://example.com/target'),
        'ODT wraps the frame in <draw:a>; DOCX dropped ImageMetadata.link entirely'));

    return results;
}

// ============================================================================
// MAIN RUNNERS
// ============================================================================

async function runAllTests(): Promise<void> {
    const startTime = Date.now();
    console.log('Starting OfficeGenerator comprehensive test suite...\n');

    const allResults: GenFeatureTest[] = [];

    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

    if (FAST_MODE) {
        console.log(`⚡ FAST MODE - skipping: ${FAST_MODE_SKIPPED_FORMATS.join(', ')}`);
        console.log('   Run \'npm run test:generator\' for full coverage before considering work done.\n');
    }

    // All source formats to test generation from
    const allSources = activeFormats([
        ...SOURCE_FORMATS.documents,
        ...SOURCE_FORMATS.spreadsheets,
    ]);

    // 1. Generation + Config tests
    console.log('Running generation and config permutation tests...');
    for (const srcFmt of allSources) {
        const srcPath = getSourcePath(srcFmt);
        if (!fs.existsSync(srcPath)) {
            console.log(`  Skipping ${srcFmt}: source file not found`);
            continue;
        }

        let ast: OfficeParserAST;
        try {
            ast = await OfficeParser.parseOffice(srcPath, PARSER_CONFIG);
        } catch (err: any) {
            console.error(`  Failed to parse ${srcFmt}: ${err.message}`);
            continue;
        }

        for (const destFmt of GENERATOR_FORMATS) {
            process.stdout.write(`  ${srcFmt} → ${destFmt}...`);
            const genResults = await testGeneration(srcFmt, ast, destFmt);
            const cfgResults = await testGeneratorConfigs(srcFmt, ast, destFmt);
            allResults.push(...genResults, ...cfgResults);
            const fails = [...genResults, ...cfgResults].filter(r => r.result.status === 'FAIL').length;
            console.log(fails > 0 ? ` ✗ (${fails} fail)` : ' ✓');
        }
    }

    // 2. Roundtrip tests
    console.log('\nRunning roundtrip tests...');
    for (const fmt of ROUNDTRIP_FORMATS) {
        process.stdout.write(`  ${fmt}→${fmt}...`);
        const rt = await testRoundtrip(fmt);
        allResults.push(...rt);
        const fails = rt.filter(r => r.result.status === 'FAIL').length;
        console.log(fails > 0 ? ` ✗ (${fails} fail)` : ' ✓');
    }

    // CSV roundtrip
    process.stdout.write('  csv→csv...');
    const csvRt = await testCsvRoundtrip();
    allResults.push(...csvRt);
    console.log(csvRt.filter(r => r.result.status === 'FAIL').length > 0 ? ' ✗' : ' ✓');

    // Configuration Validation tests
    process.stdout.write('  config validation...');
    const validationFails = runConfigValidationTests();
    allResults.push(...validationFails);
    console.log(validationFails.filter(r => r.result.status === 'FAIL').length > 0 ? ' ✗' : ' ✓');

    // Markdown dialect/fallbackToHtml config tests
    process.stdout.write('  markdown dialect config...');
    const dialectResults = await runMarkdownDialectTests();
    allResults.push(...dialectResults);
    console.log(dialectResults.filter(r => r.result.status === 'FAIL').length > 0 ? ' ✗' : ' ✓');

    // Whitespace fidelity tests (.to('text') / .to('md') parity)
    process.stdout.write('  whitespace fidelity...');
    const whitespaceResults = await runWhitespaceFidelityTests();
    allResults.push(...whitespaceResults);
    console.log(whitespaceResults.filter(r => r.result.status === 'FAIL').length > 0 ? ' ✗' : ' ✓');

    // Metadata overrides (common generator config)
    process.stdout.write('  metadata overrides...');
    const metaResults = await runMetadataOverrideTests();
    allResults.push(...metaResults);
    console.log(metaResults.filter(r => r.result.status === 'FAIL').length > 0 ? ' ✗' : ' ✓');

    // Native PDF engine (pdf-lib layout path; generates and re-parses small PDFs)
    process.stdout.write('  native pdf engine...');
    const nativePdfResults = await runNativePdfEngineTests();
    allResults.push(...nativePdfResults);
    console.log(nativePdfResults.filter(r => r.result.status === 'FAIL').length > 0 ? ' ✗' : ' ✓');

    // Table geometry: sparse spreadsheet grids, header rows, onNode call counts
    process.stdout.write('  table geometry...');
    const tableGeometryResults = await runTableGeometryTests();
    allResults.push(...tableGeometryResults);
    console.log(tableGeometryResults.filter(r => r.result.status === 'FAIL').length > 0 ? ' ✗' : ' ✓');

    // Degrade fidelity: over-cap images, unembeddable pictures, notes, OCR line breaks
    process.stdout.write('  degrade fidelity...');
    const degradeResults = await runDegradeFidelityTests();
    allResults.push(...degradeResults);
    console.log(degradeResults.filter(r => r.result.status === 'FAIL').length > 0 ? ' ✗' : ' ✓');

    // EPUB reproducibility (deliberately slow: asserts across the 2s zip-timestamp quantum)
    process.stdout.write('  epub determinism...');
    const epubDetResults = await runEpubDeterminismTests();
    allResults.push(...epubDetResults);
    console.log(epubDetResults.filter(r => r.result.status === 'FAIL').length > 0 ? ' ✗' : ' ✓');

    // 3. Report
    console.log('\n');
    const logger = new DualLogger();
    const failedCount = generateReport(allResults, logger);

    const duration = ((Date.now() - startTime) / 1000).toFixed(2);
    logger.log(`\nTotal duration: ${duration}s`);

    if (failedCount > 0) process.exit(1);
}

/**
 * Generate baselines: run generators against all parser baselines and save metrics.
 */
async function generateBaselines(): Promise<void> {
    console.log('Generating OfficeGenerator baselines...\n');

    const allSources = [...SOURCE_FORMATS.documents, ...SOURCE_FORMATS.spreadsheets];
    const baselineDir = path.join(__dirname, 'baseline');
    if (!fs.existsSync(baselineDir)) fs.mkdirSync(baselineDir, { recursive: true });

    for (const srcFmt of allSources) {
        const srcPath = getSourcePath(srcFmt);
        if (!fs.existsSync(srcPath)) { console.log(`  Skipping ${srcFmt}: no source file`); continue; }

        let ast: OfficeParserAST;
        try {
            ast = await OfficeParser.parseOffice(srcPath, PARSER_CONFIG);
        } catch (err: any) {
            console.error(`  Failed to parse ${srcFmt}: ${err.message}`);
            continue;
        }

        for (const destFmt of GENERATOR_FORMATS) {
            if (destFmt === 'csv' && !['xlsx', 'ods', 'csv'].includes(srcFmt)) continue;
            if (destFmt === 'rtf' && ['xlsx', 'ods', 'csv', 'pptx', 'odp'].includes(srcFmt)) continue;

            try {
                const result = await OfficeGenerator.generate(ast as any, destFmt as any, deterministicConfigFor(destFmt) as any);
                let metrics: GeneratedMetrics;
                if (destFmt === 'docx') {
                    metrics = extractDocxMetrics(result.value as Uint8Array);
                } else if (destFmt === 'odt') {
                    metrics = extractOdtMetrics(result.value as Uint8Array);
                } else if (result.value instanceof Uint8Array) {
                    // ZIP archive (multi-sheet CSV)
                    metrics = { outputLength: result.value.length, hasContent: result.value.length > 0, lineCount: 0, wordCount: 0 };
                } else if (destFmt === 'chunks') {
                    const chunks = Array.isArray(result.value) ? result.value : JSON.parse(result.value as string);
                    metrics = extractChunkMetrics(chunks);
                } else {
                    metrics = extractGeneratedMetrics(result.value as string, destFmt);
                }

                const baselineMetricsPath = getGeneratorBaselinePath(srcFmt, destFmt);
                fs.writeFileSync(baselineMetricsPath, JSON.stringify(metrics, null, 2), 'utf8');

                // Also save the actual file as a baseline
                const fileBaselinePath = path.join(baselineDir, `gen.${srcFmt}.to.${destFmt}`);
                if (destFmt === 'epub') {
                    fs.writeFileSync(`${fileBaselinePath}.epub`, result.value as Uint8Array);
                } else if (destFmt === 'docx') {
                    fs.writeFileSync(`${fileBaselinePath}.docx`, result.value as Uint8Array);
                } else if (destFmt === 'odt') {
                    fs.writeFileSync(`${fileBaselinePath}.odt`, result.value as Uint8Array);
                } else if (result.value instanceof Uint8Array) {
                    fs.writeFileSync(`${fileBaselinePath}.zip`, result.value);
                } else if (destFmt === 'chunks') {
                    const chunkData = Array.isArray(result.value) ? JSON.stringify(result.value, null, 2) : result.value as string;
                    fs.writeFileSync(`${fileBaselinePath}.json`, chunkData, 'utf8');
                } else {
                    fs.writeFileSync(fileBaselinePath, result.value as string, 'utf8');
                }

                console.log(`  ✓ Baseline saved: ${srcFmt} → ${destFmt} (${metrics.outputLength} chars)`);
            } catch (err: any) {
                console.error(`  ✗ Failed ${srcFmt} → ${destFmt}: ${err.message}`);
            }
        }
    }

    console.log('\nBaseline generation complete.');
}

// ============================================================================
// ENTRY POINT
// ============================================================================

const args = process.argv.slice(2);

if (args[0] === 'baseline') {
    generateBaselines().catch(err => { console.error(err); process.exit(1); });
} else {
    runAllTests().catch(err => { console.error(err); process.exit(1); });
}

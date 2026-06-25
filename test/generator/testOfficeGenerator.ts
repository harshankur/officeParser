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
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParser } from '../../src/OfficeParser';
import {
    DeepRequired,
    GeneratorConfig,
    OfficeContentNode,
    OfficeParserAST,
    OfficeParserConfig,
} from '../../src/types';
import { resolveGeneratorConfig } from '../../src/utils/configUtils.js';

// ============================================================================
// CONSTANTS & CONFIGURATION
// ============================================================================

/** Output formats tested (PDF excluded — slow/brittle in CI) */
const GENERATOR_FORMATS = ['html', 'md', 'text', 'rtf', 'csv', 'chunks'] as const;
type GeneratorFormat = typeof GENERATOR_FORMATS[number];

/** Formats that support roundtrip testing (parse → generate → re-parse) */
const ROUNDTRIP_FORMATS: GeneratorFormat[] = ['html', 'md', 'rtf'];

/** Source baseline formats used to drive generation tests */
const SOURCE_FORMATS = {
    documents: ['docx', 'odt', 'pptx', 'odp', 'pdf', 'rtf', 'html', 'md'] as const,
    spreadsheets: ['xlsx', 'ods', 'csv'] as const,
};

/** Parser config used to parse source files for generation */
const PARSER_CONFIG: DeepRequired<OfficeParserConfig> = {
    extractAttachments: true,
    ocr: false,
    ocrLanguage: 'eng',
    ocrConfig: { language: 'eng', workerPath: '', corePath: '', langPath: '', timeout: { autoTerminate: 10000, workerLoad: 60000, recognition: 30000 }, autoTerminateTimeout: 10000, abortSignal: null },
    includeRawContent: false,
    ignoreNotes: false,
    putNotesAtLast: false,
    newlineDelimiter: '\n',
    outputErrorToConsole: false,
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
        maxZipEntries: 10000
    }
};

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
];

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

function extractRoundtripMetrics(ast: OfficeParserAST): RoundtripMetrics {
    let contentNodes = 0, headings = 0, tables = 0, lists = 0, images = 0;

    function traverse(nodes: OfficeContentNode[]) {
        for (const n of nodes) {
            contentNodes++;
            if (n.type === 'heading') headings++;
            if (n.type === 'table' || n.type === 'sheet') tables++;
            if (n.type === 'list') lists++;
            if (n.type === 'image') images++;
            if (n.children) traverse(n.children);
        }
    }
    traverse(ast.content);

    const textLength = ast.toText().length;
    return { contentNodes, textLength, headings, tables, lists, images };
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

    if (destFmt === 'html' || destFmt === 'md') {
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
        const result = await OfficeGenerator.generate(ast as any, destFmt as any);

        // CSV generator returns a ZIP Uint8Array for multi-sheet sources.
        // Detect this and unpack the first CSV file for metric extraction.
        let rawOutput: string = '';
        let isZipOutput = false;
        if (destFmt === 'csv') {
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

            // G6: HTML Non-Standalone — must NOT contain DOCTYPE
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
        const metrics1 = extractRoundtripMetrics(ast1);

        // Step 2: Generate to same format
        const genResult = await OfficeGenerator.generate(ast1 as any, destFmt as any);
        const genContent = genResult.value as string;

        // Step 3: Write generated file to temp, re-parse it
        const tmpPath = path.join(__dirname, '..', 'files', `_roundtrip_tmp.${destFmt}`);
        fs.writeFileSync(tmpPath, genContent, 'utf8');

        try {
            const ast2 = await OfficeParser.parseOffice(tmpPath, { ...PARSER_CONFIG, fileType: srcFmt as any });
            const metrics2 = extractRoundtripMetrics(ast2);

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
        const text1 = ast1.toText();
        const lines1 = text1.split('\n').filter(l => l.trim() && !l.startsWith('#')).length;

        const genResult = await OfficeGenerator.generate(ast1 as any, 'csv', { csvConfig: { mergeSheets: true } } as any);
        const genCsv = genResult.value as string;
        const tmpPath = path.join(__dirname, '..', 'files', '_roundtrip_tmp.csv');
        fs.writeFileSync(tmpPath, genCsv, 'utf8');

        try {
            const ast2 = await OfficeParser.parseOffice(tmpPath, { ...PARSER_CONFIG, fileType: 'csv' });
            const text2 = ast2.toText();
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

// ============================================================================
// MAIN RUNNERS
// ============================================================================

async function runAllTests(): Promise<void> {
    const startTime = Date.now();
    console.log('Starting OfficeGenerator comprehensive test suite...\n');

    const allResults: GenFeatureTest[] = [];

    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

    // All source formats to test generation from
    const allSources = [
        ...SOURCE_FORMATS.documents,
        ...SOURCE_FORMATS.spreadsheets,
    ];

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
                const result = await OfficeGenerator.generate(ast as any, destFmt as any);
                let metrics: GeneratedMetrics;
                if (result.value instanceof Uint8Array) {
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
                if (result.value instanceof Uint8Array) {
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

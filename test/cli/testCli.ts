import * as fs from 'fs';
import * as path from 'url';
import * as fsPath from 'path';
import * as child_process from 'child_process';
import { fileURLToPath } from 'url';
import { createRequire } from 'module';
import { unzipSync, strFromU8 } from 'fflate';

const __filename = fileURLToPath(import.meta.url);
const __dirname = fsPath.dirname(__filename);
const ROOT = fsPath.join(__dirname, '..', '..');

// Resolve tsx's own CLI entry so the CLI can be launched as `node <tsx-cli> <script>` directly.
// Spawning `npx` (a `.cmd` shim on Windows) through child_process fails under modern Node's
// batch-file spawn hardening, which left `test:cli` - the last stage of `npm test` - failing on
// Windows even after the rest of #111 was fixed. Going straight through node needs no shell and
// no batch file, so it is portable.
const require = createRequire(import.meta.url);
const TSX_CLI = require.resolve('tsx/cli');

/**
 * Formats fast mode leaves out - see the matching block in `test/parser/testOfficeParser.ts` for
 * the full reasoning. Here it is one test rather than seven, but that one test shells out to the
 * CLI, which parses the PDF with OCR on: 27.0s against 1.8s for the next slowest format, making it
 * the single largest cost in `npm run test:fast` once the parser and generator suites were fixed.
 */
const FAST_MODE_SKIPPED_FORMATS = ['pdf'];

/** True when invoked as `... fast`. This suite has no `baseline` mode, so nothing else to gate. */
const FAST_MODE = (process.argv[2] || '').toLowerCase() === 'fast';

const CLI_SRC = fsPath.join(ROOT, 'src', 'cli.ts');
const SAMPLE_HTML = fsPath.join(ROOT, 'test', 'files', 'test.html');
const RESULTS_DIR = fsPath.join(__dirname, 'results');

if (!fs.existsSync(RESULTS_DIR)) {
    fs.mkdirSync(RESULTS_DIR, { recursive: true });
}

interface TestResult {
    name: string;
    status: 'PASS' | 'FAIL' | 'WARN' | 'SKIP';
    details: string;
    duration: number;
}

const results: TestResult[] = [];

class DualLogger {
    private mdContent: string = '';

    log(message: string = '') {
        console.log(message);
        this.mdContent += message + '\n';
    }

    getMarkdown(): string {
        return '```\n' + this.mdContent + '```\n';
    }

    clear() {
        this.mdContent = '';
    }
}

/** Default per-command wall-clock budget. Flag tests finish in well under 2s. */
const CLI_TIMEOUT_MS = 30000;
/**
 * A generous budget for the OCR-heavy PDF parity run, which takes ~27-30s (and more on a cold
 * cache) - the 30s default clipped it, killing the child mid-run. See `normalizeCliResult` for why
 * that kill used to masquerade as a content mismatch instead of the timeout it is.
 */
const CLI_OCR_TIMEOUT_MS = 120000;

type CliResult = { stdout: string; stderr: string; status: number; timedOut: boolean };

/**
 * Normalizes a `spawnSync` result. When `timeout` elapses, `spawnSync` kills the child with SIGTERM
 * and returns `status: null` (with `error.code === 'ETIMEDOUT'`). Coercing that null straight to 0 -
 * as `status ?? 0` did - makes a timed-out run read as a clean exit 0 with empty stdout, so the
 * parity check reports `0.0% similarity ... Got: 0 words` (a phantom content bug) rather than the
 * timeout it actually was. Surface it instead: a distinct non-zero status (124, the conventional
 * timeout code) and a `timedOut` flag callers can report on.
 */
function normalizeCliResult(result: child_process.SpawnSyncReturns<string>): CliResult {
    const timedOut = (result.error as NodeJS.ErrnoException | undefined)?.code === 'ETIMEDOUT'
        || (result.status === null && result.signal === 'SIGTERM');
    return {
        stdout: result.stdout || '',
        stderr: result.stderr || '',
        status: result.status ?? (timedOut ? 124 : 0),
        timedOut
    };
}

function runCli(args: string[], timeout: number = CLI_TIMEOUT_MS): CliResult {
    return normalizeCliResult(child_process.spawnSync(process.execPath, [TSX_CLI, CLI_SRC, SAMPLE_HTML, ...args], {
        encoding: 'utf8',
        timeout,
    }));
}

function runCliRaw(args: string[], timeout: number = CLI_TIMEOUT_MS): CliResult {
    return normalizeCliResult(child_process.spawnSync(process.execPath, [TSX_CLI, CLI_SRC, ...args], {
        encoding: 'utf8',
        timeout,
    }));
}

function assertContains(output: string, expected: string, testName: string, duration: number) {
    if (output.includes(expected)) {
        results.push({ name: testName, status: 'PASS', details: `Output contains "${expected}"`, duration });
    } else {
        results.push({
            name: testName,
            status: 'FAIL',
            details: `Expected output to contain "${expected}", but it didn't.`,
            duration
        });
    }
}

function generateReport(allResults: TestResult[], logger: DualLogger): number {
    const width = 146;
    const line = '═'.repeat(width);

    logger.log('┌' + '─'.repeat(width - 2) + '┐');
    logger.log('│' + ' '.repeat(width - 2) + '│');
    logger.log('│' + '    OFFICE PARSER CLI TEST SUITE'.padEnd(width - 2) + '│');
    logger.log('│' + '    Flag & Option Specification Validation'.padEnd(width - 2) + '│');
    logger.log('│' + ' '.repeat(width - 2) + '│');
    logger.log('└' + '─'.repeat(width - 2) + '┘');
    logger.log('');

    logger.log(line);
    logger.log('CLI FLAG VERIFICATION');
    logger.log(line);
    logger.log('');

    // Table header
    logger.log('┌' + '─'.repeat(63) + '┬' + '─'.repeat(10) + '┬' + '─'.repeat(12) + '┬' + '─'.repeat(57) + '┐');
    logger.log('│ ' + 'CLI Option / Feature'.padEnd(61) + ' │ ' + 'Status'.padEnd(8) + ' │ ' + 'Time'.padEnd(10) + ' │ ' + 'Details'.padEnd(55) + ' │');
    logger.log('├' + '─'.repeat(63) + '┼' + '─'.repeat(10) + '┼' + '─'.repeat(12) + '┼' + '─'.repeat(57) + '┤');

    for (const res of allResults) {
        const statusIcon = {
            'PASS': '✓',
            'FAIL': '✗',
            'WARN': '⚠',
            'SKIP': '⊘'
        }[res.status];

        const feature = res.name.substring(0, 61).padEnd(61);
        const status = `${statusIcon} ${res.status}`.padEnd(8);
        const time = (res.duration >= 1000 ? `${(res.duration / 1000).toFixed(1)}s` : `${res.duration}ms`).padEnd(10);
        const details = res.details.substring(0, 55).padEnd(55);

        logger.log(`│ ${feature} │ ${status} │ ${time} │ ${details} │`);
    }

    logger.log('└' + '─'.repeat(63) + '┴' + '─'.repeat(10) + '┴' + '─'.repeat(12) + '┴' + '─'.repeat(57) + '┘');
    logger.log('');

    // Summary
    logger.log(line);
    logger.log('SUMMARY');
    logger.log(line);
    logger.log('');

    const passed = allResults.filter(r => r.status === 'PASS').length;
    const failed = allResults.filter(r => r.status === 'FAIL').length;
    const warned = allResults.filter(r => r.status === 'WARN').length;
    const skipped = allResults.filter(r => r.status === 'SKIP').length;
    const total = allResults.length;

    logger.log(`Total Tests: ${total}`);
    logger.log(`✓ Passed:  ${passed} (${((passed / total) * 100).toFixed(1)}%)`);
    logger.log(`✗ Failed:  ${failed} (${((failed / total) * 100).toFixed(1)}%)`);
    logger.log(`⚠ Warned:  ${warned} (${((warned / total) * 100).toFixed(1)}%)`);
    logger.log(`⊘ Skipped: ${skipped} (${((skipped / total) * 100).toFixed(1)}%)`);
    logger.log('');

    if (failed === 0) {
        logger.log('✓ All CLI tests passed!');
    } else {
        logger.log(`✗ ${failed} test(s) failed - CLI parser needs improvement`);
    }

    // Save reports
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const jsonPath = fsPath.join(RESULTS_DIR, `cli-test-results-${timestamp}.json`);
    const mdPath = fsPath.join(RESULTS_DIR, `cli-test-results-${timestamp}.md`);

    fs.writeFileSync(jsonPath, JSON.stringify({
        timestamp: new Date().toISOString(),
        summary: { total, passed, failed, warned, skipped },
        results: allResults
    }, null, 2));

    const markdown = '# Office Parser CLI Test Results\n\n' +
        `**Generated**: ${new Date().toLocaleString()}\n\n` +
        logger.getMarkdown();
    fs.writeFileSync(mdPath, markdown);

    logger.log('');
    logger.log(`Detailed results saved to:`);
    logger.log(`  JSON: ${jsonPath}`);
    logger.log(`  Markdown: ${mdPath}`);

    return failed;
}

async function runTests() {
    console.log('Starting CLI Flag Verification Tests...');

    // 1. Default JSON AST output
    console.log('Test 1: Default JSON AST output');
    const t1 = Date.now();
    const res1 = runCli([]);
    const d1 = Date.now() - t1;
    try {
        const json = JSON.parse(res1.stdout);
        if (json.type === 'html' && Array.isArray(json.content)) {
            results.push({ name: 'Default JSON AST', status: 'PASS', details: 'Valid JSON AST returned', duration: d1 });
        } else {
            results.push({ name: 'Default JSON AST', status: 'FAIL', details: `Valid JSON but not an AST. Type: ${json.type}`, duration: d1 });
        }
    } catch (e) {
        results.push({ name: 'Default JSON AST', status: 'FAIL', details: 'Output is not valid JSON.', duration: d1 });
    }

    // 2. --format=text
    console.log('Test 2: --format=text');
    const t2 = Date.now();
    const res2 = runCli(['--format=text']);
    const d2 = Date.now() - t2;
    assertContains(res2.stdout, 'Demonstration of DOCX support', 'Format: text', d2);
    assertContains(res2.stdout, 'Inline formatting', 'Format: text (inline section)', d2);

    // 3. --format=md
    console.log('Test 3: --format=md');
    const t3 = Date.now();
    const res3 = runCli(['--format=md']);
    const d3 = Date.now() - t3;
    assertContains(res3.stdout, 'Demonstration of DOCX support', 'Format: md (heading)', d3);
    assertContains(res3.stdout, 'bold', 'Format: md (bold)', d3);
    assertContains(res3.stdout, 'italic', 'Format: md (italic)', d3);

    // 4. --format=html
    console.log('Test 4: --format=html');
    const t4 = Date.now();
    const res4 = runCli(['--format=html']);
    const d4 = Date.now() - t4;
    const hasHeading = res4.stdout.includes('Demonstration of DOCX support') && res4.stdout.includes('<h1');
    if (hasHeading) {
        results.push({ name: 'Format: html (heading)', status: 'PASS', details: 'Found H1 heading', duration: d4 });
    } else {
        results.push({ name: 'Format: html (heading)', status: 'FAIL', details: 'Expected <h1> heading.', duration: d4 });
    }

    const hasBold = res4.stdout.includes('bold') && (res4.stdout.includes('<b>') || res4.stdout.includes('<strong>') || res4.stdout.includes('span'));
    if (hasBold) {
        results.push({ name: 'Format: html (bold)', status: 'PASS', details: 'Found bold tag', duration: d4 });
    } else {
        results.push({ name: 'Format: html (bold)', status: 'FAIL', details: 'Expected bold text.', duration: d4 });
    }

    // 5. --format=chunks
    console.log('Test 5: --format=chunks');
    const t5 = Date.now();
    const res5 = runCli(['--format=chunks']);
    const d5 = Date.now() - t5;
    try {
        const chunks = JSON.parse(res5.stdout);
        if (Array.isArray(chunks) && chunks.length > 0 && chunks[0].text) {
            results.push({ name: 'Format: chunks', status: 'PASS', details: 'Valid chunks array returned', duration: d5 });
        } else {
            results.push({ name: 'Format: chunks', status: 'FAIL', details: 'Output is not a valid chunks array', duration: d5 });
        }
    } catch (e) {
        results.push({ name: 'Format: chunks', status: 'FAIL', details: 'Output is not valid JSON', duration: d5 });
    }

    // 6. --output flag
    console.log('Test 6: --output');
    const t6 = Date.now();
    const outputPath = fsPath.join(RESULTS_DIR, 'output_test.md');
    if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
    const res6 = runCli(['--format=md', `--output=${outputPath}`]);
    const d6 = Date.now() - t6;
    if (fs.existsSync(outputPath)) {
        const content = fs.readFileSync(outputPath, 'utf8');
        if (content.includes('Demonstration of DOCX support')) {
            results.push({ name: 'Output to file', status: 'PASS', details: 'File created with correct contents', duration: d6 });
        } else {
            results.push({ name: 'Output to file', status: 'FAIL', details: 'File created but content is incorrect', duration: d6 });
        }
    } else {
        results.push({ name: 'Output to file', status: 'FAIL', details: 'Output file was not created.', duration: d6 });
    }

    // 7. --verbose=true
    console.log('Test 7: --verbose=true');
    const t7 = Date.now();
    const res7 = runCli(['--verbose=true']);
    const d7 = Date.now() - t7;
    if (res7.status === 0) {
        results.push({ name: 'Verbose flag', status: 'PASS', details: 'CLI exited successfully', duration: d7 });
    } else {
        results.push({ name: 'Verbose flag', status: 'FAIL', details: `CLI failed with exit code ${res7.status}`, duration: d7 });
    }

    // 8. Custom config flag --ignoreNotes=true
    console.log('Test 8: Custom config flag --ignoreNotes=true');
    const t8 = Date.now();
    const res8 = runCli(['--ignoreNotes=true']);
    const d8 = Date.now() - t8;
    if (res8.status === 0) {
        results.push({ name: 'Custom config flag', status: 'PASS', details: 'CLI exited successfully', duration: d8 });
    } else {
        results.push({ name: 'Custom config flag', status: 'FAIL', details: `CLI failed with exit code ${res8.status}`, duration: d8 });
    }

    // 9. Every flag removed in v8 errors clearly and names its replacement. Silently accepting one
    //    is worse than rejecting it: --ocrLanguage=deu would quietly run English OCR.
    console.log('Test 9: Removed v8 flags');
    const removedFlags: Array<[string, RegExp]> = [
        ['--toText=true', /--to=text/],
        ['--ocrLanguage=deu', /--ocrConfig\.language/],
        ['--putNotesAtLast', /node\.notes/],
        ['--outputErrorToConsole', /--verbose/],
    ];
    for (const [flag, hint] of removedFlags) {
        const tR = Date.now();
        const resR = runCli([flag]);
        const dR = Date.now() - tR;
        const name = `Removed flag errors: ${flag.split('=')[0]}`;
        if (resR.status !== 0 && /removed in v8/i.test(resR.stderr) && hint.test(resR.stderr)) {
            results.push({ name, status: 'PASS', details: 'Exited nonzero with a migration hint', duration: dR });
        } else {
            results.push({ name, status: 'FAIL', details: `status ${resR.status}, stderr: ${resR.stderr.slice(0, 80)}`, duration: dR });
        }
    }

    // 10. Help / Usage output
    console.log('Test 10: Usage output (no args)');
    const t10 = Date.now();
    const res10 = child_process.spawnSync(process.execPath, [TSX_CLI, CLI_SRC], { encoding: 'utf8' });
    const d10 = Date.now() - t10;
    assertContains(res10.stdout, 'Usage: officeparser <file>', 'Usage output', d10);

    // 11. File not found
    console.log('Test 11: File not found error');
    const t11 = Date.now();
    const res11 = child_process.spawnSync(process.execPath, [TSX_CLI, CLI_SRC, 'non_existent.docx'], { encoding: 'utf8' });
    const d11 = Date.now() - t11;
    assertContains(res11.stderr, 'Error parsing file', 'File not found error', d11);

    // 12. Invalid flag value (boolean)
    console.log('Test 12: Invalid boolean flag value');
    const t12 = Date.now();
    const res12 = runCli(['--ocr=maybe']);
    const d12 = Date.now() - t12;
    assertContains(res12.stderr, 'Invalid boolean value for --ocr', 'Invalid boolean flag warning', d12);

    // 13. --includeRawContent=true
    console.log('Test 13: --includeRawContent=true');
    const t13 = Date.now();
    const res13 = runCli(['--includeRawContent=true']);
    const d13 = Date.now() - t13;
    try {
        const json = JSON.parse(res13.stdout);
        if (json.config.includeRawContent === true) {
            results.push({ name: 'Config: includeRawContent passed', status: 'PASS', details: 'Reflected in AST config', duration: d13 });
        } else {
            results.push({ name: 'Config: includeRawContent passed', status: 'FAIL', details: 'Flag not reflected in AST config', duration: d13 });
        }
        const hasRaw = JSON.stringify(json.content).includes('"rawContent"');
        if (hasRaw) {
            results.push({ name: 'AST: rawContent present', status: 'PASS', details: 'rawContent key present in content nodes', duration: d13 });
        } else {
            results.push({ name: 'AST: rawContent present', status: 'FAIL', details: 'No nodes found with rawContent', duration: d13 });
        }
    } catch (e) {
        results.push({ name: 'Config: includeRawContent check', status: 'FAIL', details: 'Failed to parse JSON', duration: d13 });
    }

    // 14. --extractAttachments=true
    console.log('Test 14: --extractAttachments=true');
    const t14 = Date.now();
    const res14 = runCli(['--extractAttachments=true']);
    const d14 = Date.now() - t14;
    try {
        const json = JSON.parse(res14.stdout);
        if (json.config.extractAttachments === true) {
            results.push({ name: 'Config: extractAttachments passed', status: 'PASS', details: 'Reflected in AST config', duration: d14 });
        } else {
            results.push({ name: 'Config: extractAttachments passed', status: 'FAIL', details: 'Flag not reflected in AST config', duration: d14 });
        }
        if (json.attachments && json.attachments.length > 0) {
            results.push({ name: 'AST: attachments present', status: 'PASS', details: 'Reflected in AST attachments array', duration: d14 });
        } else {
            results.push({ name: 'AST: attachments present', status: 'FAIL', details: 'No attachments found in AST', duration: d14 });
        }
    } catch (e) {
        results.push({ name: 'Config: extractAttachments check', status: 'FAIL', details: 'Failed to parse JSON', duration: d14 });
    }

    // 15. --format=rtf
    console.log('Test 15: --format=rtf');
    const t15 = Date.now();
    const res15 = runCli(['--format=rtf']);
    const d15 = Date.now() - t15;
    assertContains(res15.stdout, '{\\rtf1', 'Format: rtf', d15);

    // 16. --format=csv
    console.log('Test 16: --format=csv');
    const t16 = Date.now();
    const res16 = runCli(['--format=csv']);
    const d16 = Date.now() - t16;
    assertContains(res16.stdout, 'ITEM', 'Format: csv', d16);

    // 17. Flag combination
    console.log('Test 17: Flag combination');
    const t17 = Date.now();
    const res17 = runCli(['--format=text', '--includeRawContent=true', '--verbose=true']);
    const d17 = Date.now() - t17;
    assertContains(res17.stdout, 'Demonstration of DOCX support', 'Flag combination: text content', d17);
    try {
        JSON.parse(res17.stdout);
        results.push({ name: 'Flag combination: text format', status: 'FAIL', details: 'Should not output JSON', duration: d17 });
    } catch (e) {
        results.push({ name: 'Flag combination: text format', status: 'PASS', details: 'Correctly outputs plaintext', duration: d17 });
    }

    // 18. --preserveXmlWhitespace=true
    console.log('Test 18: --preserveXmlWhitespace=true');
    const t18 = Date.now();
    const res18 = runCli(['--preserveXmlWhitespace=true']);
    const d18 = Date.now() - t18;
    try {
        const json = JSON.parse(res18.stdout);
        if (json.config.preserveXmlWhitespace === true) {
            results.push({ name: 'Config: preserveXmlWhitespace passed', status: 'PASS', details: 'Reflected in AST config', duration: d18 });
        } else {
            results.push({ name: 'Config: preserveXmlWhitespace passed', status: 'FAIL', details: 'Flag not reflected in AST config', duration: d18 });
        }
    } catch (e) {
        results.push({ name: 'Config: preserveXmlWhitespace check', status: 'FAIL', details: 'Failed to parse JSON', duration: d18 });
    }

    // 19. Remaining config flags passthrough
    console.log('Test 19: Remaining config flags passthrough');
    const t19 = Date.now();
    const res19 = runCli([
        '--serializeRawContent=false',
        '--includeBreakNodes=true'
    ]);
    const d19 = Date.now() - t19;
    try {
        const json = JSON.parse(res19.stdout);
        const c = json.config;
        if (c.serializeRawContent === false) results.push({ name: 'Config: serializeRawContent passed', status: 'PASS', details: 'serializeRawContent reflected in AST config', duration: d19 });
        else results.push({ name: 'Config: serializeRawContent passed', status: 'FAIL', details: `serializeRawContent was ${c.serializeRawContent}`, duration: d19 });

        if (c.includeBreakNodes === true) results.push({ name: 'Config: includeBreakNodes passed', status: 'PASS', details: 'includeBreakNodes reflected in AST config', duration: d19 });
        else results.push({ name: 'Config: includeBreakNodes passed', status: 'FAIL', details: `includeBreakNodes was ${c.includeBreakNodes}`, duration: d19 });
    } catch (e) {
        results.push({ name: 'Config: multi-flag check', status: 'FAIL', details: 'Failed to parse JSON', duration: d19 });
    }

    // 20. Space-separated format option (--to md)
    console.log('Test 20: Space-separated target format --to md');
    const t20 = Date.now();
    const res20 = runCli(['--to', 'md']);
    const d20 = Date.now() - t20;
    assertContains(res20.stdout, 'Demonstration of DOCX support', 'Format: --to md', d20);

    // 21. Bare boolean flag (--ocr)
    console.log('Test 21: Bare boolean flag --ocr');
    const t21 = Date.now();
    const res21 = runCli(['--ocr']);
    const d21 = Date.now() - t21;
    try {
        const json = JSON.parse(res21.stdout);
        if (json.config.ocr === true) results.push({ name: 'Flag: bare --ocr is true', status: 'PASS', details: 'ocr resolved to true', duration: d21 });
        else results.push({ name: 'Flag: bare --ocr is true', status: 'FAIL', details: `Expected config.ocr to be true, got ${json.config.ocr}`, duration: d21 });
    } catch (e) {
        results.push({ name: 'Flag: bare --ocr check', status: 'FAIL', details: 'Failed to parse JSON', duration: d21 });
    }

    // 22. Negated boolean flag (--no-ocr)
    console.log('Test 22: Negated boolean flag --no-ocr');
    const t22 = Date.now();
    const res22 = runCli(['--no-ocr']);
    const d22 = Date.now() - t22;
    try {
        const json = JSON.parse(res22.stdout);
        if (json.config.ocr === false) results.push({ name: 'Flag: negated --no-ocr is false', status: 'PASS', details: 'ocr resolved to false', duration: d22 });
        else results.push({ name: 'Flag: negated --no-ocr is false', status: 'FAIL', details: `Expected config.ocr to be false, got ${json.config.ocr}`, duration: d22 });
    } catch (e) {
        results.push({ name: 'Flag: negated --no-ocr check', status: 'FAIL', details: 'Failed to parse JSON', duration: d22 });
    }

    // 23. Positional argument swallowing fix (--ocr followed by input file path)
    console.log('Test 23: Positional argument swallowing fix (--ocr SAMPLE_HTML)');
    const t23 = Date.now();
    const res23 = runCliRaw(['--ocr', SAMPLE_HTML]);
    const d23 = Date.now() - t23;
    try {
        const json = JSON.parse(res23.stdout);
        if (json.type === 'html' && json.config.ocr === true) {
            results.push({ name: 'CLI: positional argument not swallowed by bare flag', status: 'PASS', details: 'Parsed successfully with correct config', duration: d23 });
        } else {
            results.push({ name: 'CLI: positional argument not swallowed by bare flag', status: 'FAIL', details: `Expected HTML AST, got type: ${json.type}, ocr: ${json.config.ocr}`, duration: d23 });
        }
    } catch (e) {
        results.push({ name: 'CLI: positional argument swallowing check', status: 'FAIL', details: `CLI failed to output valid JSON when flag precedes filename.`, duration: d23 });
    }

    // 24. Nested config option (--ocrConfig.language fra)
    console.log('Test 24: Nested config option --ocrConfig.language fra');
    const t24 = Date.now();
    const res24 = runCli(['--ocrConfig.language', 'fra']);
    const d24 = Date.now() - t24;
    try {
        const json = JSON.parse(res24.stdout);
        if (json.config.ocrConfig && json.config.ocrConfig.language === 'fra') {
            results.push({ name: 'Config: nested --ocrConfig.language passed', status: 'PASS', details: 'ocrConfig.language set correctly', duration: d24 });
        } else {
            results.push({ name: 'Config: nested --ocrConfig.language passed', status: 'FAIL', details: 'ocrConfig.language not set correctly', duration: d24 });
        }
    } catch (e) {
        results.push({ name: 'Config: nested --ocrConfig.language check', status: 'FAIL', details: 'Failed to parse JSON', duration: d24 });
    }

    // 25. Nested config option dot-notation with equals (--htmlConfig.containerWidth=800px)
    console.log('Test 25: Nested config option --htmlConfig.containerWidth=800px');
    const t25 = Date.now();
    const res25 = runCli(['--to', 'html', '--htmlConfig.containerWidth=800px']);
    const d25 = Date.now() - t25;
    assertContains(res25.stdout, '--container-width: 800px', 'Config: nested --htmlConfig.containerWidth passed', d25);

    // 26. Deprecated options warning logging
    console.log('Test 26: Deprecated options warning logging');
    const t26 = Date.now();
    const res26 = runCli(['--format=text']);
    const d26 = Date.now() - t26;
    if (res26.stderr.includes('Warning: --format is deprecated')) {
        results.push({ name: 'CLI: deprecated option prints warning', status: 'PASS', details: 'Warning message found in stderr', duration: d26 });
    } else {
        results.push({ name: 'CLI: deprecated option prints warning', status: 'FAIL', details: 'Expected warning message in stderr not found', duration: d26 });
    }

    // 27-38: Parity comparison against baseline parser outputs for all 12 formats
    if (FAST_MODE) {
        console.log(`\n⚡ FAST MODE - skipping parity for: ${FAST_MODE_SKIPPED_FORMATS.join(', ')}`);
        console.log('   Run \'npm run test:cli\' for full coverage before considering work done.\n');
    }
    const extensions = ['docx', 'odt', 'xlsx', 'ods', 'pptx', 'odp', 'odg', 'pdf', 'rtf', 'csv', 'html', 'md', 'epub'];
    for (let idx = 0; idx < extensions.length; idx++) {
        const ext = extensions[idx];
        const testNum = 27 + idx;
        // Skipped in place rather than filtered out of `extensions`, so the numbering the console
        // output and the report use stays stable between a fast run and a full one.
        if (FAST_MODE && FAST_MODE_SKIPPED_FORMATS.includes(ext)) {
            console.log(`Test ${testNum}: CLI Parity for ${ext.toUpperCase()} - SKIPPED (fast mode)`);
            results.push({
                name: `CLI text parity: ${ext}`,
                status: 'SKIP',
                details: `Fast mode: ${ext} omitted (OCR-heavy). Run 'npm run test:cli' for full coverage.`,
                duration: 0
            });
            continue;
        }
        console.log(`Test ${testNum}: CLI Parity against Parser Baseline for ${ext.toUpperCase()}`);
        const tStart = Date.now();
        const testFile = fsPath.join(ROOT, 'test', 'files', `test.${ext}`);
        
        // Check plain text parity via --to=text. OCR (pdf) can run ~30s on a cold cache, so this
        // one call gets the larger budget rather than the 30s default.
        const resText = runCliRaw([
            testFile,
            '--to=text',
            '--ocr=true',
            '--includeBreakNodes=true',
            '--includeRawContent=true',
            '--extractAttachments=true'
        ], CLI_OCR_TIMEOUT_MS);
        const dText = Date.now() - tStart;

        if (resText.status !== 0) {
            results.push({
                name: `CLI text parity: ${ext}`,
                status: 'FAIL',
                details: resText.timedOut
                    ? `CLI command timed out (killed after ${CLI_OCR_TIMEOUT_MS}ms). This reads as an empty output, not a parser regression.`
                    : `CLI command failed with exit code ${resText.status}. Error: ${resText.stderr}`,
                duration: dText
            });
            continue;
        }

        const baselineTextPath = fsPath.join(ROOT, 'test', 'parser', 'baseline', `test.${ext}.txt`);
        if (fs.existsSync(baselineTextPath)) {
            const baselineText = fs.readFileSync(baselineTextPath, 'utf8');
            const baselineWords = baselineText.split(/\s+/).filter(w => w.length > 0);
            const cliWords = resText.stdout.split(/\s+/).filter(w => w.length > 0);
            
            let similarity = 100;
            if (baselineWords.length > 0 || cliWords.length > 0) {
                const wordCountDiff = Math.abs(baselineWords.length - cliWords.length);
                const wordCountSimilarity = 1 - (wordCountDiff / Math.max(baselineWords.length, cliWords.length));
                similarity = wordCountSimilarity * 100;
            }
            
            const isPassing = similarity === 100;
            if (isPassing) {
                results.push({
                    name: `CLI text parity: ${ext}`,
                    status: 'PASS',
                    details: `Output similarity is ${similarity.toFixed(1)}% (expected 100.0%)`,
                    duration: dText
                });
            } else {
                results.push({
                    name: `CLI text parity: ${ext}`,
                    status: 'FAIL',
                    details: `Output similarity is ${similarity.toFixed(1)}% (expected 100.0%). Expected: ${baselineWords.length} words, Got: ${cliWords.length} words`,
                    duration: dText
                });
            }
        } else {
            results.push({
                name: `CLI text parity: ${ext}`,
                status: 'SKIP',
                details: 'No text baseline found',
                duration: dText
            });
        }
    }

    // 39. Markdown dialect flag (--mdConfig.dialect=commonmark)
    console.log('Test 39: Markdown dialect flag --mdConfig.dialect=commonmark');
    const t39 = Date.now();
    const res39 = runCli(['--to', 'md', '--mdConfig.dialect=commonmark']);
    const d39 = Date.now() - t39;
    assertContains(res39.stdout, '<table', 'Config: --mdConfig.dialect=commonmark forces HTML tables', d39);

    // 40. Markdown fallbackToHtml flag (--mdConfig.fallbackToHtml=false)
    console.log('Test 40: Markdown fallbackToHtml flag --mdConfig.fallbackToHtml=false');
    const t40 = Date.now();
    const res40 = runCli(['--to', 'md', '--mdConfig.fallbackToHtml=false']);
    const d40 = Date.now() - t40;
    if (!res40.stdout.includes('<u>')) {
        results.push({ name: 'Config: --mdConfig.fallbackToHtml=false disables HTML fallback', status: 'PASS', details: 'No <u> tag in output', duration: d40 });
    } else {
        results.push({ name: 'Config: --mdConfig.fallbackToHtml=false disables HTML fallback', status: 'FAIL', details: 'Found <u> tag despite fallbackToHtml=false', duration: d40 });
    }

    // 41. Binary DOCX generation to a file (--to docx --output), with a docxConfig flag threaded through
    console.log('Test 41: DOCX generation --to docx --output');
    const t41 = Date.now();
    const docxOut = fsPath.join(RESULTS_DIR, 'output_test.docx');
    if (fs.existsSync(docxOut)) fs.unlinkSync(docxOut);
    const res41 = runCli(['--to', 'docx', '--docxConfig.format=Letter', `--output=${docxOut}`]);
    const d41 = Date.now() - t41;
    if (res41.status === 0 && fs.existsSync(docxOut)) {
        const bytes = new Uint8Array(fs.readFileSync(docxOut));
        let ok = false, detail = '';
        try {
            const files = unzipSync(bytes);
            const doc = files['word/document.xml'] ? strFromU8(files['word/document.xml']) : '';
            // Valid OOXML package (PK + required parts) and Letter page size honored (12240 x 15840 twips).
            ok = bytes[0] === 0x50 && bytes[1] === 0x4B
                && !!files['[Content_Types].xml'] && !!files['word/document.xml']
                && /<w:pgSz w:w="12240"/.test(doc);
            detail = ok ? 'PK zip with document.xml and Letter pgSz' : `pgSz/parts check failed: ${doc.match(/<w:pgSz[^>]*>/)?.[0] ?? 'no pgSz'}`;
        } catch (e: any) {
            detail = `output is not a valid zip: ${e.message}`;
        }
        results.push({ name: 'CLI: --to docx writes a valid Word package', status: ok ? 'PASS' : 'FAIL', details: detail, duration: d41 });
    } else {
        results.push({ name: 'CLI: --to docx writes a valid Word package', status: 'FAIL', details: `exit ${res41.status}, file exists: ${fs.existsSync(docxOut)}. stderr: ${res41.stderr.slice(0, 120)}`, duration: d41 });
    }

    // 42. Binary ODT generation to a file (--to odt --output), with an odtConfig flag threaded through
    console.log('Test 42: ODT generation --to odt --output');
    const t42 = Date.now();
    const odtOut = fsPath.join(RESULTS_DIR, 'output_test.odt');
    if (fs.existsSync(odtOut)) fs.unlinkSync(odtOut);
    const res42 = runCli(['--to', 'odt', '--odtConfig.format=Letter', `--output=${odtOut}`]);
    const d42 = Date.now() - t42;
    if (res42.status === 0 && fs.existsSync(odtOut)) {
        const bytes = new Uint8Array(fs.readFileSync(odtOut));
        let ok = false, detail = '';
        try {
            const files = unzipSync(bytes);
            const styles = files['styles.xml'] ? strFromU8(files['styles.xml']) : '';
            // Valid ODF package: PK zip, mimetype first + STORED, content.xml present, Letter width in styles.
            const mimetypeFirst = strFromU8(bytes.slice(30, 38)) === 'mimetype' && (bytes[8] | (bytes[9] << 8)) === 0;
            ok = bytes[0] === 0x50 && bytes[1] === 0x4B && mimetypeFirst
                && !!files['content.xml'] && /fo:page-width="8.5in"/.test(styles);
            detail = ok ? 'ODF zip: mimetype-first/STORED, content.xml, Letter page-width' : `checks failed: ${styles.match(/fo:page-width="[^"]*"/)?.[0] ?? 'no page-width'}, mimetypeFirst=${mimetypeFirst}`;
        } catch (e: any) {
            detail = `output is not a valid zip: ${e.message}`;
        }
        results.push({ name: 'CLI: --to odt writes a valid ODF package', status: ok ? 'PASS' : 'FAIL', details: detail, duration: d42 });

        // End-to-end: feed the generated ODT back through the CLI and confirm text re-extracts.
        const back = runCliRaw([odtOut, '--to=text']);
        const textOk = back.status === 0 && back.stdout.includes('Demonstration of DOCX support');
        results.push({ name: 'CLI: generated ODT re-parses through the CLI', status: textOk ? 'PASS' : 'FAIL', details: textOk ? 'round-tripped text via --to=text' : `exit ${back.status}, stdout: ${back.stdout.slice(0, 120)}`, duration: 0 });
    } else {
        results.push({ name: 'CLI: --to odt writes a valid ODF package', status: 'FAIL', details: `exit ${res42.status}, file exists: ${fs.existsSync(odtOut)}. stderr: ${res42.stderr.slice(0, 120)}`, duration: d42 });
    }

    // 43. A renamed option reaches ast.warnings even though no onWarning was supplied, and the CLI
    //     prints it without --verbose. Both halves regressed silently before: the collector that
    //     fills ast.warnings used to be installed after config resolution (so the warning went to
    //     the default no-op handler), and the CLI only printed warnings under --verbose.
    console.log('Test 43: Renamed option surfaces in ast.warnings and on stderr');
    const t43 = Date.now();
    const res43 = runCli(['--ocrConfig.autoTerminateTimeout=5000']);
    const d43 = Date.now() - t43;
    try {
        const json = JSON.parse(res43.stdout);
        const warn = (json.warnings || []).find((w: any) => w.code === 'UNRECOGNIZED_CONFIG_OPTION');
        if (warn && /ocrConfig\.autoTerminateTimeout/.test(warn.message) && /ocrConfig\.timeout\.autoTerminate/.test(warn.message)) {
            results.push({ name: 'Warnings: renamed key reaches ast.warnings without onWarning', status: 'PASS', details: 'Warning names both the old key and its replacement', duration: d43 });
        } else {
            results.push({ name: 'Warnings: renamed key reaches ast.warnings without onWarning', status: 'FAIL', details: `Expected UNRECOGNIZED_CONFIG_OPTION naming ocrConfig.timeout.autoTerminate, got: ${JSON.stringify(json.warnings)?.slice(0, 90)}`, duration: d43 });
        }
    } catch (e) {
        results.push({ name: 'Warnings: renamed key reaches ast.warnings without onWarning', status: 'FAIL', details: 'Failed to parse JSON', duration: d43 });
    }
    if (res43.stderr.includes('UNRECOGNIZED_CONFIG_OPTION') && res43.stderr.includes('ocrConfig.timeout.autoTerminate')) {
        results.push({ name: 'CLI: unrecognized option printed without --verbose', status: 'PASS', details: 'Warning found on stderr', duration: 0 });
    } else {
        results.push({ name: 'CLI: unrecognized option printed without --verbose', status: 'FAIL', details: `stderr: ${res43.stderr.slice(0, 80)}`, duration: 0 });
    }

    // 44. v8 parser flags reach the resolved config: --password, --ignorePageGeometry and the
    //     dotted --pdfParserConfig.* family (a boolean and a string in the same run).
    console.log('Test 44: v8 parser flags passthrough');
    const t44 = Date.now();
    const res44 = runCli([
        '--password=secret',
        '--ignorePageGeometry=true',
        '--pdfParserConfig.useTags=false',
        '--pdfParserConfig.pageRange=1-3'
    ]);
    const d44 = Date.now() - t44;
    try {
        const c = JSON.parse(res44.stdout).config;
        const checks: Array<[string, boolean, string]> = [
            ['Config: --password passed', c.password === 'secret', `password was ${JSON.stringify(c.password)}`],
            ['Config: --ignorePageGeometry passed', c.ignorePageGeometry === true, `ignorePageGeometry was ${c.ignorePageGeometry}`],
            ['Config: --pdfParserConfig.useTags passed', c.pdfParserConfig?.useTags === false, `useTags was ${c.pdfParserConfig?.useTags}`],
            ['Config: --pdfParserConfig.pageRange passed', c.pdfParserConfig?.pageRange === '1-3', `pageRange was ${JSON.stringify(c.pdfParserConfig?.pageRange)}`],
        ];
        for (const [name, ok, detail] of checks) {
            results.push({ name, status: ok ? 'PASS' : 'FAIL', details: ok ? 'Reflected in AST config' : detail, duration: d44 });
        }
    } catch (e) {
        results.push({ name: 'Config: v8 parser flags check', status: 'FAIL', details: 'Failed to parse JSON', duration: d44 });
    }

    // 45. --includeImages takes an ImageMode in both the `=` and the space-separated form. The
    //     space form used to read as a bare boolean, leaving the mode behind as a stray positional
    //     that the CLI then mistook for the input file.
    console.log('Test 45: --includeImages=<mode> and --includeImages <mode>');
    const t45 = Date.now();
    const res45eq = runCli(['--to=md', '--includeImages=none']);
    const res45sp = runCli(['--to=md', '--includeImages', 'none']);
    const res45default = runCli(['--to=md']);
    const d45 = Date.now() - t45;
    if (res45sp.status === 0 && res45sp.stdout === res45eq.stdout && res45eq.stdout !== res45default.stdout) {
        results.push({ name: 'CLI: --includeImages accepts a mode in both forms', status: 'PASS', details: 'Space form matches = form and changes the output', duration: d45 });
    } else {
        results.push({ name: 'CLI: --includeImages accepts a mode in both forms', status: 'FAIL', details: `space exit ${res45sp.status}, forms equal: ${res45sp.stdout === res45eq.stdout}, mode had effect: ${res45eq.stdout !== res45default.stdout}`, duration: d45 });
    }

    // 46. LaTeX generation: `--to tex` prints a complete document, `latex` is an alias for it, a
    //     texConfig flag reaches the generator, and a bare `--texConfig.bundle` (a known boolean, so
    //     it must not swallow the output path) writes a zip holding main.tex and its images.
    console.log('Test 46: LaTeX generation --to tex / --to latex / --texConfig.*');
    const t46 = Date.now();
    const res46 = runCli(['--to', 'tex', '--texConfig.numberSections']);
    const res46alias = runCli(['--to=latex']);
    const texBundleOut = fsPath.join(RESULTS_DIR, 'output_test.zip');
    if (fs.existsSync(texBundleOut)) fs.unlinkSync(texBundleOut);
    const res46bundle = runCli(['--extractAttachments', '--to', 'tex', '--texConfig.bundle', `--output=${texBundleOut}`]);
    const d46 = Date.now() - t46;
    const texOk = res46.status === 0 && /^% Generated by officeParser\.\n\\documentclass\[[^\]]*\]\{article\}/.test(res46.stdout)
        && res46.stdout.includes('\\begin{document}') && res46.stdout.includes('\\end{document}')
        && res46.stdout.includes('\\section') && !res46.stdout.includes('secnumdepth');
    results.push({ name: 'CLI: --to tex prints a LaTeX document (numberSections honored)', status: texOk ? 'PASS' : 'FAIL', details: texOk ? 'preamble, body, sections, numbering left on' : `exit ${res46.status}, head: ${res46.stdout.slice(0, 100)}`, duration: d46 });
    const aliasOk = res46alias.status === 0 && res46alias.stdout.includes('\\documentclass');
    results.push({ name: 'CLI: --to latex is an alias of tex', status: aliasOk ? 'PASS' : 'FAIL', details: aliasOk ? 'alias resolved' : `exit ${res46alias.status}, stderr: ${res46alias.stderr.slice(0, 120)}`, duration: 0 });
    let bundleOk = false, bundleDetail = '';
    if (res46bundle.status === 0 && fs.existsSync(texBundleOut)) {
        try {
            const files = unzipSync(new Uint8Array(fs.readFileSync(texBundleOut)));
            const main = files['main.tex'] ? strFromU8(files['main.tex']) : '';
            const images = Object.keys(files).filter(n => n.startsWith('images/'));
            bundleOk = main.includes('\\includegraphics') && images.length > 0 && images.every(p => main.includes(p));
            bundleDetail = bundleOk ? `main.tex + ${images.length} image(s), all referenced` : `main.tex: ${!!main}, images: ${images.join(',')}`;
        } catch (e: any) { bundleDetail = `output is not a valid zip: ${e.message}`; }
    } else {
        bundleDetail = `exit ${res46bundle.status}, file exists: ${fs.existsSync(texBundleOut)}. stderr: ${res46bundle.stderr.slice(0, 120)}`;
    }
    results.push({ name: 'CLI: --texConfig.bundle writes main.tex plus its images', status: bundleOk ? 'PASS' : 'FAIL', details: bundleDetail, duration: 0 });

    // 47. LaTeX input: a .tex file parses and converts (here to DOCX), and a LaTeX project zip
    //     named .zip is routed by what it holds.
    console.log('Test 47: LaTeX input --to docx / project .zip');
    const t47 = Date.now();
    const texIn = fsPath.join(ROOT, 'test', 'files', 'test.tex');
    const texDocx = fsPath.join(RESULTS_DIR, 'from_tex.docx');
    if (fs.existsSync(texDocx)) fs.unlinkSync(texDocx);
    const res47 = runCliRaw([texIn, '--to=docx', `--output=${texDocx}`]);
    let docxOk = false;
    if (res47.status === 0 && fs.existsSync(texDocx)) {
        const doc = strFromU8(unzipSync(new Uint8Array(fs.readFileSync(texDocx)))['word/document.xml'] ?? new Uint8Array());
        docxOk = doc.includes('Demonstration of DOCX support in calibre') && /w:val="Heading1"/.test(doc);
    }
    results.push({ name: 'CLI: a .tex file converts to DOCX', status: docxOk ? 'PASS' : 'FAIL', details: docxOk ? 'headings and text carried over' : `exit ${res47.status}, stderr: ${res47.stderr.slice(0, 160)}`, duration: Date.now() - t47 });
    const res47zip = runCliRaw([texBundleOut, '--to=text']);
    const zipOk = res47zip.status === 0 && res47zip.stdout.includes('Demonstration of DOCX support');
    results.push({ name: 'CLI: a LaTeX project .zip is parsed as LaTeX', status: zipOk ? 'PASS' : 'FAIL', details: zipOk ? 'routed by archive contents' : `exit ${res47zip.status}, stderr: ${res47zip.stderr.slice(0, 160)}`, duration: 0 });

    // Print summary report
    const logger = new DualLogger();
    const failedCount = generateReport(results, logger);

    if (failedCount > 0) {
        process.exit(1);
    }
}

runTests();

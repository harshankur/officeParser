import * as fs from 'fs';
import * as path from 'url';
import * as fsPath from 'path';
import * as child_process from 'child_process';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = fsPath.dirname(__filename);
const ROOT = fsPath.join(__dirname, '..', '..');

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

function runCli(args: string[]): { stdout: string; stderr: string; status: number } {
    const result = child_process.spawnSync('npx', ['tsx', CLI_SRC, SAMPLE_HTML, ...args], {
        encoding: 'utf8',
        timeout: 30000,
    });
    return {
        stdout: result.stdout || '',
        stderr: result.stderr || '',
        status: result.status ?? 0
    };
}

function runCliRaw(args: string[]): { stdout: string; stderr: string; status: number } {
    const result = child_process.spawnSync('npx', ['tsx', CLI_SRC, ...args], {
        encoding: 'utf8',
        timeout: 30000,
    });
    return {
        stdout: result.stdout || '',
        stderr: result.stderr || '',
        status: result.status ?? 0
    };
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

    // 9. Legacy --toText=true
    console.log('Test 9: Legacy --toText=true');
    const t9 = Date.now();
    const res9 = runCli(['--toText=true']);
    const d9 = Date.now() - t9;
    assertContains(res9.stdout, 'Demonstration of DOCX support', 'Legacy --toText', d9);

    // 10. Help / Usage output
    console.log('Test 10: Usage output (no args)');
    const t10 = Date.now();
    const res10 = child_process.spawnSync('npx', ['tsx', CLI_SRC], { encoding: 'utf8' });
    const d10 = Date.now() - t10;
    assertContains(res10.stdout, 'Usage: officeparser <file>', 'Usage output', d10);

    // 11. File not found
    console.log('Test 11: File not found error');
    const t11 = Date.now();
    const res11 = child_process.spawnSync('npx', ['tsx', CLI_SRC, 'non_existent.docx'], { encoding: 'utf8' });
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
        '--ocrLanguage=fra',
        '--putNotesAtLast=true',
        '--serializeRawContent=false',
        '--includeBreakNodes=true'
    ]);
    const d19 = Date.now() - t19;
    try {
        const json = JSON.parse(res19.stdout);
        const c = json.config;
        if (c.ocrLanguage === 'fra') results.push({ name: 'Config: ocrLanguage passed', status: 'PASS', details: 'ocrLanguage reflected in AST config', duration: d19 });
        else results.push({ name: 'Config: ocrLanguage passed', status: 'FAIL', details: `ocrLanguage was ${c.ocrLanguage}`, duration: d19 });

        if (c.putNotesAtLast === true) results.push({ name: 'Config: putNotesAtLast passed', status: 'PASS', details: 'putNotesAtLast reflected in AST config', duration: d19 });
        else results.push({ name: 'Config: putNotesAtLast passed', status: 'FAIL', details: `putNotesAtLast was ${c.putNotesAtLast}`, duration: d19 });

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
    const extensions = ['docx', 'odt', 'xlsx', 'ods', 'pptx', 'odp', 'pdf', 'rtf', 'csv', 'html', 'md', 'epub'];
    for (let idx = 0; idx < extensions.length; idx++) {
        const ext = extensions[idx];
        const testNum = 27 + idx;
        console.log(`Test ${testNum}: CLI Parity against Parser Baseline for ${ext.toUpperCase()}`);
        const tStart = Date.now();
        const testFile = fsPath.join(ROOT, 'test', 'files', `test.${ext}`);
        
        // Check plain text parity via --to=text
        const resText = runCliRaw([
            testFile,
            '--to=text',
            '--ocr=true',
            '--includeBreakNodes=true',
            '--includeRawContent=true',
            '--extractAttachments=true'
        ]);
        const dText = Date.now() - tStart;
        
        if (resText.status !== 0) {
            results.push({
                name: `CLI text parity: ${ext}`,
                status: 'FAIL',
                details: `CLI command failed with exit code ${resText.status}. Error: ${resText.stderr}`,
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

    // Print summary report
    const logger = new DualLogger();
    const failedCount = generateReport(results, logger);

    if (failedCount > 0) {
        process.exit(1);
    }
}

runTests();

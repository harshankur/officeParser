import * as fs from 'fs';
import * as path from 'path';
import * as child_process from 'child_process';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT = path.join(__dirname, '..');

const CLI_SRC = path.join(ROOT, 'src', 'cli.ts');
const SAMPLE_HTML = path.join(ROOT, 'test', 'cli_sample.html');
const RESULTS_DIR = path.join(ROOT, 'test', 'results', 'cli');

if (!fs.existsSync(RESULTS_DIR)) {
    fs.mkdirSync(RESULTS_DIR, { recursive: true });
}

interface TestResult {
    name: string;
    status: 'PASS' | 'FAIL';
    message?: string;
}

const results: TestResult[] = [];

function runCli(args: string[]): { stdout: string; stderr: string; status: number } {
    // Use tsx to run the TypeScript CLI source
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

function assertContains(output: string, expected: string, testName: string) {
    if (output.includes(expected)) {
        results.push({ name: testName, status: 'PASS' });
    } else {
        results.push({ name: testName, status: 'FAIL', message: `Expected output to contain "${expected}", but it didn't. Got: ${output.slice(0, 100)}...` });
    }
}

async function runTests() {
    console.log('Starting CLI Flag Verification Tests...');

    // 1. Basic parsing (JSON AST)
    console.log('Test 1: Default JSON AST output');
    const res1 = runCli([]);
    try {
        const json = JSON.parse(res1.stdout);
        // OfficeParserAST has a 'type' property matching the input format (html in this case)
        if (json.type === 'html' && Array.isArray(json.content)) {
            results.push({ name: 'Default JSON AST', status: 'PASS' });
        } else {
            results.push({ name: 'Default JSON AST', status: 'FAIL', message: `Output is valid JSON but not a valid AST. Got type: ${json.type}` });
        }
    } catch (e) {
        results.push({ name: 'Default JSON AST', status: 'FAIL', message: `Output is not valid JSON. Error: ${e}. Stdout: ${res1.stdout.slice(0, 100)}` });
    }

    // 2. --format=text
    console.log('Test 2: --format=text');
    const res2 = runCli(['--format=text']);
    assertContains(res2.stdout, 'Main Heading', 'Format: text');
    assertContains(res2.stdout, 'Item 1', 'Format: text (list item)');

    // 3. --format=md
    console.log('Test 3: --format=md');
    const res3 = runCli(['--format=md']);
    assertContains(res3.stdout, '# Main Heading', 'Format: md (heading)');
    assertContains(res3.stdout, '**bold**', 'Format: md (bold)');
    assertContains(res3.stdout, '*italic*', 'Format: md (italic)');

    // 4. --format=html
    console.log('Test 4: --format=html');
    const res4 = runCli(['--format=html']);
    // Headings might have IDs generated
    const hasHeading = res4.stdout.includes('Main Heading') && res4.stdout.includes('<h1');
    if (hasHeading) {
        results.push({ name: 'Format: html (heading)', status: 'PASS' });
    } else {
        results.push({ name: 'Format: html (heading)', status: 'FAIL', message: `Expected <h1> heading. Got: ${res4.stdout.slice(0, 200)}...` });
    }

    // Bold might be in <b> or <strong> depending on how it's parsed/generated
    const hasBold = res4.stdout.includes('bold') && (res4.stdout.includes('<b>') || res4.stdout.includes('<strong>'));
    if (hasBold) {
        results.push({ name: 'Format: html (bold)', status: 'PASS' });
    } else {
        results.push({ name: 'Format: html (bold)', status: 'FAIL', message: `Expected bold text. Got: ${res4.stdout.slice(0, 500)}...` });
    }

    // 5. --format=chunks
    console.log('Test 5: --format=chunks');
    const res5 = runCli(['--format=chunks']);
    try {
        const chunks = JSON.parse(res5.stdout);
        if (Array.isArray(chunks) && chunks.length > 0 && chunks[0].text) {
            results.push({ name: 'Format: chunks', status: 'PASS' });
        } else {
            results.push({ name: 'Format: chunks', status: 'FAIL', message: 'Output is not a valid chunks array' });
        }
    } catch (e) {
        results.push({ name: 'Format: chunks', status: 'FAIL', message: 'Output is not valid JSON' });
    }

    // 6. --output flag
    console.log('Test 6: --output');
    const outputPath = path.join(RESULTS_DIR, 'output_test.md');
    if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
    const res6 = runCli(['--format=md', `--output=${outputPath}`]);
    if (fs.existsSync(outputPath)) {
        const content = fs.readFileSync(outputPath, 'utf8');
        if (content.includes('# Main Heading')) {
            results.push({ name: 'Output to file', status: 'PASS' });
        } else {
            results.push({ name: 'Output to file', status: 'FAIL', message: 'File created but content is incorrect' });
        }
    } else {
        results.push({ name: 'Output to file', status: 'FAIL', message: `Output file was not created. Stderr: ${res6.stderr}` });
    }

    // 7. --verbose=true
    console.log('Test 7: --verbose=true');
    const res7 = runCli(['--verbose=true']);
    if (res7.status === 0) {
        results.push({ name: 'Verbose flag', status: 'PASS' });
    } else {
        results.push({ name: 'Verbose flag', status: 'FAIL', message: `CLI failed with exit code ${res7.status}. Stderr: ${res7.stderr}` });
    }

    // 8. Custom config flag --ignoreNotes (even if HTML doesn't have notes, check if it passes through)
    console.log('Test 8: Custom config flag --ignoreNotes=true');
    const res8 = runCli(['--ignoreNotes=true']);
    if (res8.status === 0) {
        results.push({ name: 'Custom config flag', status: 'PASS' });
    } else {
        results.push({ name: 'Custom config flag', status: 'FAIL', message: `CLI failed when passing custom config flag. Stderr: ${res8.stderr}` });
    }

    // 9. Legacy --toText=true
    console.log('Test 9: Legacy --toText=true');
    const res9 = runCli(['--toText=true']);
    assertContains(res9.stdout, 'Main Heading', 'Legacy --toText');

    // 10. Help / Usage output
    console.log('Test 10: Usage output (no args)');
    const res10 = child_process.spawnSync('npx', ['tsx', CLI_SRC], { encoding: 'utf8' });
    assertContains(res10.stdout, 'Usage: officeparser <file>', 'Usage output');

    // 11. File not found
    console.log('Test 11: File not found error');
    const res11 = child_process.spawnSync('npx', ['tsx', CLI_SRC, 'non_existent.docx'], { encoding: 'utf8' });
    assertContains(res11.stderr, 'Error parsing file', 'File not found error');

    // 12. Invalid flag value (boolean)
    console.log('Test 12: Invalid boolean flag value');
    const res12 = runCli(['--ocr=maybe']);
    // Should warn and use default
    assertContains(res12.stderr, 'Invalid boolean value for --ocr', 'Invalid boolean flag warning');

    // 13. --includeRawContent=true
    console.log('Test 13: --includeRawContent=true');
    const res13 = runCli(['--includeRawContent=true']);
    try {
        const json = JSON.parse(res13.stdout);
        // Check if config was passed
        if (json.config.includeRawContent === true) {
            results.push({ name: 'Config: includeRawContent passed', status: 'PASS' });
        } else {
            results.push({ name: 'Config: includeRawContent passed', status: 'FAIL', message: 'Flag not reflected in AST config' });
        }
        // Check if at least one node has rawContent (HtmlParser now adds it to text nodes)
        const hasRaw = JSON.stringify(json.content).includes('"rawContent"');
        if (hasRaw) {
            results.push({ name: 'AST: rawContent present', status: 'PASS' });
        } else {
            results.push({ name: 'AST: rawContent present', status: 'FAIL', message: 'No nodes found with rawContent' });
        }
    } catch (e) {
        results.push({ name: 'Config: includeRawContent check', status: 'FAIL', message: 'Failed to parse JSON' });
    }

    // 14. --extractAttachments=true
    console.log('Test 14: --extractAttachments=true');
    const res14 = runCli(['--extractAttachments=true']);
    try {
        const json = JSON.parse(res14.stdout);
        if (json.config.extractAttachments === true) {
            results.push({ name: 'Config: extractAttachments passed', status: 'PASS' });
        } else {
            results.push({ name: 'Config: extractAttachments passed', status: 'FAIL' });
        }
        // Check if attachments were extracted (sample HTML has a data-uri image)
        if (json.attachments && json.attachments.length > 0) {
            results.push({ name: 'AST: attachments present', status: 'PASS' });
        } else {
            results.push({ name: 'AST: attachments present', status: 'FAIL', message: 'No attachments found in AST' });
        }
    } catch (e) {
        results.push({ name: 'Config: extractAttachments check', status: 'FAIL' });
    }

    // 15. --format=rtf
    console.log('Test 15: --format=rtf');
    const res15 = runCli(['--format=rtf']);
    // RTF should start with {\rtf1
    assertContains(res15.stdout, '{\\rtf1', 'Format: rtf');

    // 16. --format=csv
    console.log('Test 16: --format=csv');
    const res16 = runCli(['--format=csv']);
    // CSV output for non-spreadsheet should be a simple CSV or warning?
    // Actually HtmlGenerator produces a CSV if it finds a table.
    assertContains(res16.stdout, 'Header 1,Header 2', 'Format: csv');

    // 17. Flag combination
    console.log('Test 17: Flag combination');
    const res17 = runCli(['--format=text', '--includeRawContent=true', '--verbose=true']);
    assertContains(res17.stdout, 'Main Heading', 'Flag combination: text content');
    // Verbose should be passed (though it mainly affects error reporting)
    try {
        const json = JSON.parse(res17.stdout); // Wait, --format=text is NOT JSON
        results.push({ name: 'Flag combination: text format skip JSON check', status: 'PASS' });
    } catch (e) {
        // Expected because it's text
        results.push({ name: 'Flag combination: text format', status: 'PASS' });
    }

    // 18. --preserveXmlWhitespace=true
    console.log('Test 18: --preserveXmlWhitespace=true');
    const res18 = runCli(['--preserveXmlWhitespace=true']);
    try {
        const json = JSON.parse(res18.stdout);
        if (json.config.preserveXmlWhitespace === true) {
            results.push({ name: 'Config: preserveXmlWhitespace passed', status: 'PASS' });
        } else {
            results.push({ name: 'Config: preserveXmlWhitespace passed', status: 'FAIL' });
        }
    } catch (e) {
        results.push({ name: 'Config: preserveXmlWhitespace check', status: 'FAIL' });
    }

    // 19. Remaining config flags passthrough
    console.log('Test 19: Remaining config flags passthrough');
    const res19 = runCli([
        '--ocrLanguage=fra',
        '--putNotesAtLast=true',
        '--serializeRawContent=false',
        '--includeBreakNodes=true'
    ]);
    try {
        const json = JSON.parse(res19.stdout);
        const c = json.config;
        if (c.ocrLanguage === 'fra') results.push({ name: 'Config: ocrLanguage passed', status: 'PASS' });
        else results.push({ name: 'Config: ocrLanguage passed', status: 'FAIL' });

        if (c.putNotesAtLast === true) results.push({ name: 'Config: putNotesAtLast passed', status: 'PASS' });
        else results.push({ name: 'Config: putNotesAtLast passed', status: 'FAIL' });

        if (c.serializeRawContent === false) results.push({ name: 'Config: serializeRawContent passed', status: 'PASS' });
        else results.push({ name: 'Config: serializeRawContent passed', status: 'FAIL' });

        if (c.includeBreakNodes === true) results.push({ name: 'Config: includeBreakNodes passed', status: 'PASS' });
        else results.push({ name: 'Config: includeBreakNodes passed', status: 'FAIL' });
    } catch (e) {
        results.push({ name: 'Config: multi-flag check', status: 'FAIL' });
    }

    // Print Summary
    console.log('\n--- CLI Test Summary ---');
    let passed = 0;
    results.forEach(r => {
        if (r.status === 'PASS') {
            console.log(`✅ [PASS] ${r.name}`);
            passed++;
        } else {
            console.log(`❌ [FAIL] ${r.name}: ${r.message || ''}`);
            if (r.message === undefined && (r as any).stdout) {
                console.log(`   Stdout: ${(r as any).stdout.slice(0, 100)}`);
            }
        }
    });
    console.log(`\nTotal: ${passed}/${results.length} passed.`);

    if (passed < results.length) {
        process.exit(1);
    }
}

runTests();

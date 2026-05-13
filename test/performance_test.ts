import * as fs from 'fs';
import * as path from 'path';
import { OfficeParser } from '../src/OfficeParser';
import { OfficeParserConfig } from '../src/types';

const TEST_FILES_DIR = path.join(__dirname, 'files');
const ITERATIONS = 3; // To get a slightly more stable average, though "one-shot" might mean 1.

const CONFIG: OfficeParserConfig = {
    extractAttachments: true,
    ocr: false, // Disabling OCR for baseline performance
    includeRawContent: true,
};

async function runPerformanceTest() {
    const files = fs.readdirSync(TEST_FILES_DIR)
        .filter(file => !file.includes('.actual.') && !file.startsWith('.'))
        .sort();

    console.log('# OfficeParser Performance Test Results');
    console.log(`\nTesting ${files.length} files with ${ITERATIONS} iterations each.`);
    console.log('Config: OCR=false, Attachments=true, RawContent=true\n');
    
    console.log('| File | Extension | Size (KB) | Iteration 1 (ms) | Iteration 2 (ms) | Iteration 3 (ms) | Average (ms) |');
    console.log('| :--- | :--- | :--- | :--- | :--- | :--- | :--- |');

    for (const file of files) {
        const filePath = path.join(TEST_FILES_DIR, file);
        const stats = fs.statSync(filePath);
        const sizeKb = (stats.size / 1024).toFixed(2);
        const ext = path.extname(file).slice(1);

        const times: number[] = [];
        for (let i = 0; i < ITERATIONS; i++) {
            const start = performance.now();
            try {
                await OfficeParser.parseOffice(filePath, CONFIG);
                const end = performance.now();
                times.push(end - start);
            } catch (err) {
                console.error(`Error parsing ${file}:`, err);
                times.push(-1);
            }
        }

        const avg = times.filter(t => t >= 0).reduce((a, b) => a + b, 0) / times.length;
        const timeCols = times.map(t => t.toFixed(2)).join(' | ');
        console.log(`| ${file} | ${ext} | ${sizeKb} | ${timeCols} | ${avg.toFixed(2)} |`);
    }
}

runPerformanceTest().catch(console.error);

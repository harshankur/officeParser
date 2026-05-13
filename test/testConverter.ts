import { OfficeConverter } from '../src/OfficeConverter.js';
import * as fs from 'fs';
import * as path from 'path';

async function runTests() {
    console.log('--- OfficeConverter Verification Tests ---');

    const testFilesDir = path.join(process.cwd(), 'test', 'files');
    const outputDir = path.join(process.cwd(), 'test', 'results', 'converter');

    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    try {
        // 1. DOCX to MD
        console.log('Testing: DOCX -> MD...');
        const docxPath = path.join(testFilesDir, 'test.docx');
        const docxResult = await OfficeConverter.convert(docxPath, 'md', {
            generatorConfig: {
                includeImages: true,
            },
            onWarning: (msg) => console.warn(`[DOCX->MD Warning]: ${msg}`)
        });
        fs.writeFileSync(path.join(outputDir, 'test_docx.md'), docxResult.value);
        console.log(`Success! Saved to ${path.join(outputDir, 'test_docx.md')}`);
        console.log(`Messages: ${docxResult.messages.length}`);

        // 2. PDF to HTML
        console.log('\nTesting: PDF -> HTML...');
        const pdfPath = path.join(testFilesDir, 'test.pdf');
        const pdfResult = await OfficeConverter.convert(pdfPath, 'html', {
            generatorConfig: {
                includeImages: true,
            },
            onWarning: (msg) => console.warn(`[PDF->HTML Warning]: ${msg}`)
        });
        fs.writeFileSync(path.join(outputDir, 'test_pdf.html'), pdfResult.value);
        console.log(`Success! Saved to ${path.join(outputDir, 'test_pdf.html')}`);

        // 3. Buffer test with fileType hint (CSV -> Text)
        console.log('\nTesting: Buffer (CSV) -> Text with fileType hint...');
        const csvPath = path.join(testFilesDir, 'test.csv');
        const csvBuffer = fs.readFileSync(csvPath);
        const csvResult = await OfficeConverter.convert(csvBuffer, 'text', {
            parseConfig: {
                fileType: 'csv',
                csvDelimiter: ','
            },
            generatorConfig: {
                textConfig: {
                    preserveLayout: true
                }
            }
        });
        fs.writeFileSync(path.join(outputDir, 'test_csv.txt'), csvResult.value);
        console.log(`Success! Saved to ${path.join(outputDir, 'test_csv.txt')}`);

        // 4. Test filtering (DOCX -> MD without images)
        console.log('\nTesting: DOCX -> MD (includeImages: false)...');
        const docxNoImagesResult = await OfficeConverter.convert(docxPath, 'md', {
            generatorConfig: {
                includeImages: false
            }
        });
        const hasImages = docxNoImagesResult.value.includes('![image]');
        console.log(`Success! Images present: ${hasImages} (expected: false)`);
        fs.writeFileSync(path.join(outputDir, 'test_docx_no_images.md'), docxNoImagesResult.value);

    } catch (error) {
        console.error('Test failed:', error);
        process.exit(1);
    }
}

runTests();

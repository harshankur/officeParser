import * as fs from 'fs';
import * as path from 'path';
import { OfficeParser } from '../src/OfficeParser';

async function analyze(ext: 'html' | 'md' | 'csv') {
    const filePath = path.join(__dirname, 'files', `test.${ext}`);
    console.log(`\n--- Analyzing ${ext.toUpperCase()} ---`);
    
    if (!fs.existsSync(filePath)) {
        console.error(`File not found: ${filePath}`);
        return;
    }

    const original = fs.readFileSync(filePath, 'utf8');
    
    try {
        const ast = await OfficeParser.parseOffice(filePath, {
            includeRawContent: true,
            extractAttachments: true
        });

        const result = await ast.to(ext as any, ext === 'csv' ? { csvConfig: { mergeSheets: true } } : {});
        const generated = result.value as string;

        const outputDir = path.join(__dirname, 'results', 'roundtrip');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }
        
        const outputPath = path.join(outputDir, `test.generated.${ext}`);
        fs.writeFileSync(outputPath, generated);
        
        console.log(`Original size: ${original.length} chars`);
        console.log(`Generated size: ${generated.length} chars`);
        
        if (original === generated) {
            console.log("✅ Perfect match!");
        } else {
            console.log("❌ Differences found.");
            const originalLines = original.split('\n');
            const generatedLines = generated.split('\n');
            console.log(`Original lines: ${originalLines.length}`);
            console.log(`Generated lines: ${generatedLines.length}`);
            
            // Compare first 5 lines
            console.log("\n--- First 5 lines comparison ---");
            for (let i = 0; i < 5; i++) {
                const orig = (originalLines[i] || '').trim();
                const gen = (generatedLines[i] || '').trim();
                if (orig !== gen) {
                    console.log(`L${i+1} DIFF:`);
                    console.log(`  Orig: "${orig}"`);
                    console.log(`  Gen:  "${gen}"`);
                } else {
                    console.log(`L${i+1} OK`);
                }
            }
        }
        
        console.log(`Generated file saved to: ${outputPath}`);

    } catch (error) {
        console.error(`Error processing ${ext}:`, error);
    }
}

async function main() {
    await analyze('html');
    await analyze('md');
    await analyze('csv');
}

main();

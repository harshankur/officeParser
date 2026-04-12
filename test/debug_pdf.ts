
import * as fs from 'fs';
import * as path from 'path';
import { parsePdf } from '../src/parsers/PdfParser';

async function main() {
    const pdfPath = path.join(__dirname, 'files', 'test.pdf');
    const buffer = fs.readFileSync(pdfPath);
    console.log("Starting PDF parse for debug...");
    
    try {
        const ast = await parsePdf(buffer, {
            extractAttachments: true,
            ocr: true,
            outputErrorToConsole: true
        });
        console.log("Parse complete!");
        console.log("Content nodes:", ast.content.length);
        console.log("Links found:", ast.content.flatMap(c => c.children || []).filter(c => c.metadata?.link).length);
    } catch (e) {
        console.error("Parse failed:", e);
    }
}

main();

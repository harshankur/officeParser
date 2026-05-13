import { OfficeParser } from '../../src/OfficeParser';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import * as path from 'path';
import * as fs from 'fs';

async function testXlsxChunking() {
    console.log('Testing XLSX chunking...');
    const xlsxPath = path.join(__dirname, '..', 'files', 'test.xlsx');
    
    if (!fs.existsSync(xlsxPath)) {
        console.error('XLSX file not found at:', xlsxPath);
        process.exit(1);
    }

    const ast = await OfficeParser.parseOffice(xlsxPath);
    console.log(`Parsed XLSX. AST content nodes: ${ast.content.length}`);

    // Strategy: document-structure (default)
    const result = await OfficeGenerator.generate(ast as any, 'chunks');
    const chunks = result.value as any[];
    
    console.log(`Generated ${chunks.length} chunks.`);
    if (chunks.length > 0) {
        console.log('Sample chunk:', JSON.stringify(chunks[0], null, 2));
        console.log('✓ PASS: XLSX produced chunks.');
    } else {
        console.error('✗ FAIL: XLSX produced 0 chunks.');
        console.log('Messages:', result.messages);
    }

    // Strategy: fixed-size
    const resultFixed = await OfficeGenerator.generate(ast as any, 'chunks', { 
        chunksConfig: { strategy: 'fixed-size', chunkSize: 100 } 
    });
    const chunksFixed = resultFixed.value as any[];
    console.log(`Fixed-size produced ${chunksFixed.length} chunks.`);
    if (chunksFixed.length > 0) {
        console.log('✓ PASS: XLSX produced fixed-size chunks.');
    } else {
        console.error('✗ FAIL: XLSX produced 0 fixed-size chunks.');
    }
}

testXlsxChunking().catch(console.error);

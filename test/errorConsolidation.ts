import { OfficeParser } from '../src/OfficeParser';
import { OfficeIssue, OfficeWarningType } from '../src/types';
import * as fs from 'fs';
import * as path from 'path';

async function testErrorConsolidation() {
    console.log('Testing Error Consolidation...');

    // 1. Test parsing-phase warning accumulation
    // We'll use a file that might trigger a warning, or we'll mock one.
    // For now, let's use a dummy PDF and expect a worker fallback warning in Node if not found.
    const dummyPdf = Buffer.from('%PDF-1.4\n1 0 obj\n<<>>\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF');
    
    let callbackIssue: OfficeIssue | undefined;
    try {
        const ast = await OfficeParser.parseOffice(dummyPdf, {
            fileType: 'pdf',
            onWarning: (issue) => {
                console.log('  [Callback Received Issue]:', issue.code, '-', issue.message);
                callbackIssue = issue;
            }
        });
        console.log('  [AST Warnings Count]:', ast.warnings.length);
    } catch (err) {
        console.log('  [Caught Expected Error]:', (err as Error).message);
    }

    if (callbackIssue) {
        console.log('  [Callback Issue Verified]: type =', callbackIssue.type, ', code =', callbackIssue.code);
    } else {
        console.error('  [FAILED]: No issue received in callback');
    }

    // 1.5 Test non-fatal warning accumulation
    console.log('\nTesting Non-fatal Warning Accumulation...');
    const pdfBuffer = Buffer.from('%PDF-1.4\n1 0 obj\n<<>>\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF');
    const astWithWarning = await OfficeParser.parseOffice(pdfBuffer, {
        fileType: 'md' // Intentional mismatch to trigger BUFFER_TYPE_MISMATCH
    });
    console.log('  [AST Warnings Count]:', astWithWarning.warnings.length);
    if (astWithWarning.warnings.length > 0) {
        console.log('  [Captured Warning]:', astWithWarning.warnings[0].code, '-', astWithWarning.warnings[0].message);
    }

    // 2. Test generation-phase warning
    console.log('\nTesting Generation Warning...');
    const validAst = await OfficeParser.parseOffice(Buffer.from('# Valid Doc'), { fileType: 'md' });
    const result = await validAst.to('md', {
        onWarning: (issue) => {
            console.log('  [Generator Callback Issue]:', issue.code, '-', issue.message);
        }
    });
    console.log('  [Result Messages Count]:', result.messages.length);

    console.log('\nError Consolidation Test Completed.');
}

testErrorConsolidation().catch(err => {
    console.error('Test Failed:', err);
    process.exit(1);
});

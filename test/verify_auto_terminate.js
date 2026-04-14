const { OfficeParser } = require('./dist/OfficeParser.js');
const path = require('path');

const testFile = path.join(__dirname, 'test', 'files', 'test.pdf');

console.log('--- OCR Auto-Termination Verification ---');
console.log('Processing PDF with OCR (2s auto-terminate)...');

async function run() {
    const start = Date.now();
    try {
        await OfficeParser.parseOffice(testFile, {
            ocr: true,
            ocrConfig: {
                autoTerminateTimeout: 2000 // 2 seconds
            }
        });
        console.log('Parsing complete.');
        console.log('Waiting for auto-termination (should take ~2s)...');
        
        // We don't call process.exit(). If it exits, it means the workers matched the timer.
    } catch (err) {
        console.error('Error during parsing:', err);
    }
}

run();

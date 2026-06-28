const puppeteer = require('puppeteer');
const http = require('http');
const fs = require('fs');
const path = require('path');
const child_process = require('child_process');

const ROOT = path.join(__dirname, '..');

const MIMES = {
    '.html': 'text/html',
    '.js': 'application/javascript',
    '.mjs': 'application/javascript',
    '.css': 'text/css',
    '.json': 'application/json',
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.pdf': 'application/pdf'
};

// Start static file server
const server = http.createServer((req, res) => {
    const urlPath = req.url.split('?')[0];

    // Serve virtual pages for tests
    if (urlPath === '/test-iife') {
        res.writeHead(200, {
            'Content-Type': 'text/html',
            'Cross-Origin-Opener-Policy': 'same-origin',
            'Cross-Origin-Embedder-Policy': 'require-corp'
        });
        res.end(`
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <title>IIFE Integration Test</title>
                <script src="/dist/officeparser.browser.iife.js"></script>
            </head>
            <body>
                <div id="status">RUNNING</div>
                <script>
                    // Polyfill Promise.try for environments lacking ES2024 native support
                    if (typeof Promise !== 'undefined' && !Promise.try) {
                        Promise.try = function(fn, ...args) {
                            return new Promise((resolve, reject) => {
                                try {
                                    resolve(fn(...args));
                                } catch (err) {
                                    reject(err);
                                }
                            });
                        };
                    }

                    // Polyfill Map getOrInsert / getOrInsertComputed for modern PDF.js (v5.x)
                    if (typeof Map !== 'undefined') {
                        if (!Map.prototype.getOrInsertComputed) {
                            Map.prototype.getOrInsertComputed = function(key, callback) {
                                if (this.has(key)) {
                                    return this.get(key);
                                }
                                const value = callback(key);
                                this.set(key, value);
                                return value;
                            };
                        }
                        if (!Map.prototype.getOrInsert) {
                            Map.prototype.getOrInsert = function(key, value) {
                                if (this.has(key)) {
                                    return this.get(key);
                                }
                                this.set(key, value);
                                return value;
                            };
                        }
                    }

                    async function run() {
                        try {
                            const results = [];
                            
                            // 1. Test DOCX
                            const docxRes = await fetch('/test/files/test.docx');
                            const docxBuf = await docxRes.arrayBuffer();
                            const docxAst = await officeParser.parseOffice(new Uint8Array(docxBuf));
                            const docxText = (await docxAst.to('text')).value;
                            if (docxText.includes('Demonstration of DOCX support')) {
                                results.push('DOCX: PASS');
                            } else {
                                results.push('DOCX: FAIL (text missing: ' + docxText.slice(0, 100) + ')');
                            }

                            // 2. Test XLSX
                            const xlsxRes = await fetch('/test/files/test.xlsx');
                            const xlsxBuf = await xlsxRes.arrayBuffer();
                            const xlsxAst = await officeParser.parseOffice(new Uint8Array(xlsxBuf));
                            const xlsxText = (await xlsxAst.to('text')).value;
                            if (xlsxText.includes('ITEM') || xlsxText.includes('Revenue')) {
                                results.push('XLSX: PASS');
                            } else {
                                results.push('XLSX: FAIL (text missing: ' + xlsxText.slice(0, 100) + ')');
                            }

                            // 3. Test PDF (uses local polyfilled web worker)
                            const pdfRes = await fetch('/test/files/test.pdf');
                            const pdfBuf = await pdfRes.arrayBuffer();
                            const pdfAst = await officeParser.parseOffice(new Uint8Array(pdfBuf), { pdfWorkerSrc: '/dist/pdf.worker.mjs' });
                            const pdfText = (await pdfAst.to('text')).value;
                            if (pdfText.includes('Demonstration of DOCX support') || pdfText.includes('calibre')) {
                                results.push('PDF: PASS');
                            } else {
                                results.push('PDF: FAIL (text missing: ' + pdfText.slice(0, 100) + ')');
                            }

                            document.getElementById('status').textContent = results.join(' | ');
                        } catch (err) {
                            console.error('TEST ERROR STACK:', err.stack || err.message);
                            document.getElementById('status').textContent = 'ERROR: ' + err.message;
                        }
                    }
                    run();
                </script>
            </body>
            </html>
        `);
        return;
    }

    if (urlPath === '/test-esm') {
        res.writeHead(200, {
            'Content-Type': 'text/html',
            'Cross-Origin-Opener-Policy': 'same-origin',
            'Cross-Origin-Embedder-Policy': 'require-corp'
        });
        res.end(`
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <title>ESM Integration Test</title>
            </head>
            <body>
                <div id="status">RUNNING</div>
                <script type="module">
                    // Polyfill Promise.try for environments lacking ES2024 native support
                    if (typeof Promise !== 'undefined' && !Promise.try) {
                        Promise.try = function(fn, ...args) {
                            return new Promise((resolve, reject) => {
                                try {
                                    resolve(fn(...args));
                                } catch (err) {
                                    reject(err);
                                }
                            });
                        };
                    }

                    // Polyfill Map getOrInsert / getOrInsertComputed for modern PDF.js (v5.x)
                    if (typeof Map !== 'undefined') {
                        if (!Map.prototype.getOrInsertComputed) {
                            Map.prototype.getOrInsertComputed = function(key, callback) {
                                if (this.has(key)) {
                                    return this.get(key);
                                }
                                const value = callback(key);
                                this.set(key, value);
                                return value;
                            };
                        }
                        if (!Map.prototype.getOrInsert) {
                            Map.prototype.getOrInsert = function(key, value) {
                                if (this.has(key)) {
                                    return this.get(key);
                                }
                                this.set(key, value);
                                return value;
                            };
                        }
                    }

                    import { parseOffice } from '/dist/officeparser.browser.mjs';

                    async function run() {
                        try {
                            const results = [];
                            
                            // 1. Test DOCX
                            const docxRes = await fetch('/test/files/test.docx');
                            const docxBuf = await docxRes.arrayBuffer();
                            const docxAst = await parseOffice(new Uint8Array(docxBuf));
                            const docxText = (await docxAst.to('text')).value;
                            if (docxText.includes('Demonstration of DOCX support')) {
                                results.push('DOCX: PASS');
                            } else {
                                results.push('DOCX: FAIL (text missing: ' + docxText.slice(0, 100) + ')');
                            }

                            // 2. Test XLSX
                            const xlsxRes = await fetch('/test/files/test.xlsx');
                            const xlsxBuf = await xlsxRes.arrayBuffer();
                            const xlsxAst = await parseOffice(new Uint8Array(xlsxBuf));
                            const xlsxText = (await xlsxAst.to('text')).value;
                            if (xlsxText.includes('ITEM') || xlsxText.includes('Revenue')) {
                                results.push('XLSX: PASS');
                            } else {
                                results.push('XLSX: FAIL (text missing: ' + xlsxText.slice(0, 100) + ')');
                            }

                            // 3. Test PDF (uses local polyfilled web worker)
                            const pdfRes = await fetch('/test/files/test.pdf');
                            const pdfBuf = await pdfRes.arrayBuffer();
                            const pdfAst = await parseOffice(new Uint8Array(pdfBuf), { pdfWorkerSrc: '/dist/pdf.worker.mjs' });
                            const pdfText = (await pdfAst.to('text')).value;
                            if (pdfText.includes('Demonstration of DOCX support') || pdfText.includes('calibre')) {
                                results.push('PDF: PASS');
                            } else {
                                results.push('PDF: FAIL (text missing: ' + pdfText.slice(0, 100) + ')');
                            }

                            document.getElementById('status').textContent = results.join(' | ');
                        } catch (err) {
                            console.error('TEST ERROR STACK:', err.stack || err.message);
                            document.getElementById('status').textContent = 'ERROR: ' + err.message;
                        }
                    }
                    run();
                </script>
            </body>
            </html>
        `);
        return;
    }

    if (urlPath === '/test-iife-slim') {
        res.writeHead(200, {
            'Content-Type': 'text/html',
            'Cross-Origin-Opener-Policy': 'same-origin',
            'Cross-Origin-Embedder-Policy': 'require-corp'
        });
        res.end(`
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <title>IIFE Slim Integration Test</title>
                <script src="/dist/officeparser.browser.slim.iife.js"></script>
            </head>
            <body>
                <div id="status">RUNNING</div>
                <script>
                    if (typeof Promise !== 'undefined' && !Promise.try) {
                        Promise.try = function(fn, ...args) {
                            return new Promise((resolve, reject) => {
                                try { resolve(fn(...args)); } catch (err) { reject(err); }
                            });
                        };
                    }
                    if (typeof Map !== 'undefined') {
                        if (!Map.prototype.getOrInsertComputed) {
                            Map.prototype.getOrInsertComputed = function(key, callback) {
                                if (this.has(key)) return this.get(key);
                                const value = callback(key);
                                this.set(key, value);
                                return value;
                            };
                        }
                        if (!Map.prototype.getOrInsert) {
                            Map.prototype.getOrInsert = function(key, value) {
                                if (this.has(key)) return this.get(key);
                                this.set(key, value);
                                return value;
                            };
                        }
                    }

                    async function run() {
                        try {
                            const results = [];
                            
                            // 1. Test DOCX
                            const docxRes = await fetch('/test/files/test.docx');
                            const docxBuf = await docxRes.arrayBuffer();
                            const docxAst = await officeParser.parseOffice(new Uint8Array(docxBuf));
                            const docxText = (await docxAst.to('text')).value;
                            if (docxText.includes('Demonstration of DOCX support')) {
                                results.push('DOCX: PASS');
                            } else {
                                results.push('DOCX: FAIL');
                            }

                            // 2. Test PDF (without worker)
                            let pdfError = null;
                            try {
                                const pdfRes = await fetch('/test/files/test.pdf');
                                const pdfBuf = await pdfRes.arrayBuffer();
                                await officeParser.parseOffice(new Uint8Array(pdfBuf));
                            } catch (err) {
                                pdfError = err;
                            }
                            if (pdfError !== null) {
                                results.push('PDF_DEFAULT_MISSING: PASS');
                            } else {
                                results.push('PDF_DEFAULT_MISSING: FAIL');
                            }

                            // 3. Test PDF (with worker)
                            const pdfRes2 = await fetch('/test/files/test.pdf');
                            const pdfBuf2 = await pdfRes2.arrayBuffer();
                            const pdfAst = await officeParser.parseOffice(new Uint8Array(pdfBuf2), { pdfWorkerSrc: '/dist/pdf.worker.mjs' });
                            const pdfText = (await pdfAst.to('text')).value;
                            if (pdfText.includes('Demonstration of DOCX support')) {
                                results.push('PDF_WITH_WORKER: PASS');
                            } else {
                                results.push('PDF_WITH_WORKER: FAIL');
                            }

                            // 4. Test OCR Stub
                            const docxRes2 = await fetch('/test/files/test.docx');
                            const docxBuf2 = await docxRes2.arrayBuffer();
                            const docxAst2 = await officeParser.parseOffice(new Uint8Array(docxBuf2), { ocr: true, extractAttachments: true });
                            const hasOcrStubWarning = docxAst2.warnings && docxAst2.warnings.some(w => 
                                w.code === 'OCR_FAILED' && 
                                w.details && 
                                w.details.message && 
                                w.details.message.includes('disabled in the slim browser bundle')
                            );
                            if (hasOcrStubWarning) {
                                results.push('OCR_STUB: PASS');
                            } else {
                                results.push('OCR_STUB: FAIL (warnings: ' + JSON.stringify(docxAst2.warnings || []) + ')');
                            }

                            document.getElementById('status').textContent = results.join(' | ');
                        } catch (err) {
                            console.error(err);
                            document.getElementById('status').textContent = 'ERROR: ' + err.message;
                        }
                    }
                    run();
                </script>
            </body>
            </html>
        `);
        return;
    }

    if (urlPath === '/test-esm-slim') {
        res.writeHead(200, {
            'Content-Type': 'text/html',
            'Cross-Origin-Opener-Policy': 'same-origin',
            'Cross-Origin-Embedder-Policy': 'require-corp'
        });
        res.end(`
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <title>ESM Slim Integration Test</title>
            </head>
            <body>
                <div id="status">RUNNING</div>
                <script type="module">
                    if (typeof Promise !== 'undefined' && !Promise.try) {
                        Promise.try = function(fn, ...args) {
                            return new Promise((resolve, reject) => {
                                try { resolve(fn(...args)); } catch (err) { reject(err); }
                            });
                        };
                    }
                    if (typeof Map !== 'undefined') {
                        if (!Map.prototype.getOrInsertComputed) {
                            Map.prototype.getOrInsertComputed = function(key, callback) {
                                if (this.has(key)) return this.get(key);
                                const value = callback(key);
                                this.set(key, value);
                                return value;
                            };
                        }
                        if (!Map.prototype.getOrInsert) {
                            Map.prototype.getOrInsert = function(key, value) {
                                if (this.has(key)) return this.get(key);
                                this.set(key, value);
                                return value;
                            };
                        }
                    }

                    import { parseOffice } from '/dist/officeparser.browser.slim.mjs';

                    async function run() {
                        try {
                            const results = [];
                            
                            // 1. Test DOCX
                            const docxRes = await fetch('/test/files/test.docx');
                            const docxBuf = await docxRes.arrayBuffer();
                            const docxAst = await parseOffice(new Uint8Array(docxBuf));
                            const docxText = (await docxAst.to('text')).value;
                            if (docxText.includes('Demonstration of DOCX support')) {
                                results.push('DOCX: PASS');
                            } else {
                                results.push('DOCX: FAIL');
                            }

                            // 2. Test PDF (without worker)
                            let pdfError = null;
                            try {
                                const pdfRes = await fetch('/test/files/test.pdf');
                                const pdfBuf = await pdfRes.arrayBuffer();
                                await parseOffice(new Uint8Array(pdfBuf));
                            } catch (err) {
                                pdfError = err;
                            }
                            if (pdfError !== null) {
                                results.push('PDF_DEFAULT_MISSING: PASS');
                            } else {
                                results.push('PDF_DEFAULT_MISSING: FAIL');
                            }

                            // 3. Test PDF (with worker)
                            const pdfRes2 = await fetch('/test/files/test.pdf');
                            const pdfBuf2 = await pdfRes2.arrayBuffer();
                            const pdfAst = await parseOffice(new Uint8Array(pdfBuf2), { pdfWorkerSrc: '/dist/pdf.worker.mjs' });
                            const pdfText = (await pdfAst.to('text')).value;
                            if (pdfText.includes('Demonstration of DOCX support')) {
                                results.push('PDF_WITH_WORKER: PASS');
                            } else {
                                results.push('PDF_WITH_WORKER: FAIL');
                            }

                            // 4. Test OCR Stub
                            const docxRes2 = await fetch('/test/files/test.docx');
                            const docxBuf2 = await docxRes2.arrayBuffer();
                            const docxAst2 = await parseOffice(new Uint8Array(docxBuf2), { ocr: true, extractAttachments: true });
                            const hasOcrStubWarning = docxAst2.warnings && docxAst2.warnings.some(w => 
                                w.code === 'OCR_FAILED' && 
                                w.details && 
                                w.details.message && 
                                w.details.message.includes('disabled in the slim browser bundle')
                            );
                            if (hasOcrStubWarning) {
                                results.push('OCR_STUB: PASS');
                            } else {
                                results.push('OCR_STUB: FAIL (warnings: ' + JSON.stringify(docxAst2.warnings || []) + ')');
                            }

                            document.getElementById('status').textContent = results.join(' | ');
                        } catch (err) {
                            console.error(err);
                            document.getElementById('status').textContent = 'ERROR: ' + err.message;
                        }
                    }
                    run();
                </script>
            </body>
            </html>
        `);
        return;
    }

    if (urlPath === '/dist/pdf.worker.mjs') {
        const workerPath = path.join(ROOT, 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.worker.mjs');
        if (fs.existsSync(workerPath)) {
            const polyfill = `
if (typeof Promise !== 'undefined' && !Promise.try) {
    Promise.try = function(fn, ...args) {
        return new Promise((resolve, reject) => {
            try {
                resolve(fn(...args));
            } catch (err) {
                reject(err);
            }
        });
    };
}

if (typeof Map !== 'undefined') {
    if (!Map.prototype.getOrInsertComputed) {
        Map.prototype.getOrInsertComputed = function(key, callback) {
            if (this.has(key)) {
                return this.get(key);
            }
            const value = callback(key);
            this.set(key, value);
            return value;
        };
    }
    if (!Map.prototype.getOrInsert) {
        Map.prototype.getOrInsert = function(key, value) {
            if (this.has(key)) {
                return this.get(key);
            }
            this.set(key, value);
            return value;
        };
    }
}
`;
            const content = fs.readFileSync(workerPath, 'utf8');
            res.writeHead(200, {
                'Content-Type': 'application/javascript',
                'Cross-Origin-Opener-Policy': 'same-origin',
                'Cross-Origin-Embedder-Policy': 'require-corp'
            });
            res.end(polyfill + '\n' + content);
            return;
        }
    }

    const safePath = path.normalize(urlPath).replace(/^(\.\.[\/\\])+/, '');
    const filePath = path.join(ROOT, safePath);

    fs.stat(filePath, (err, stats) => {
        if (err || !stats.isFile()) {
            res.writeHead(404);
            res.end('Not found');
            return;
        }
        const ext = path.extname(filePath).toLowerCase();
        const mime = MIMES[ext] || 'application/octet-stream';
        res.writeHead(200, {
            'Content-Type': mime,
            'Cross-Origin-Opener-Policy': 'same-origin',
            'Cross-Origin-Embedder-Policy': 'require-corp'
        });
        fs.createReadStream(filePath).pipe(res);
    });
});

async function main() {
    let failed = false;

    // Start server
    await new Promise((resolve) => {
        server.listen(0, resolve);
    });
    const port = server.address().port;
    console.log(`Test server started at http://localhost:${port}`);

    // Launch Puppeteer
    console.log('Launching browser...');
    const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    const page = await browser.newPage();

    // Enable console & error logs inside the browser page
    page.on('console', msg => console.log('PAGE LOG [' + msg.type() + ']:', msg.text()));
    page.on('pageerror', err => console.error('PAGE ERROR:', err.stack || err.message));
    page.on('requestfailed', request => {
        console.error('REQUEST FAILED:', request.url(), request.failure()?.errorText || '');
    });

    // -----------------------------------------------------------------------
    // PART 1: IIFE Browser Bundle Tests
    // -----------------------------------------------------------------------
    try {
        console.log('\n--- Testing IIFE Browser Bundle ---');
        await page.goto(`http://localhost:${port}/test-iife`, { waitUntil: 'networkidle2' });

        await page.waitForFunction(
            () => document.getElementById('status').textContent !== 'RUNNING',
            { timeout: 15000 }
        );

        const status = await page.$eval('#status', el => el.textContent);
        console.log(`IIFE Status: ${status}`);

        if (status.includes('FAIL') || status.includes('ERROR')) {
            console.error('❌ IIFE Bundle Integration Tests Failed!');
            failed = true;
        } else {
            console.log('✅ IIFE Bundle Integration Tests Passed!');
        }
    } catch (err) {
        console.error(`❌ IIFE Test execution failed: ${err.message}`);
        failed = true;
    }

    // -----------------------------------------------------------------------
    // PART 2: ESM Browser Bundle Tests
    // -----------------------------------------------------------------------
    try {
        console.log('\n--- Testing ESM Browser Bundle ---');
        await page.goto(`http://localhost:${port}/test-esm`, { waitUntil: 'networkidle2' });

        await page.waitForFunction(
            () => document.getElementById('status').textContent !== 'RUNNING',
            { timeout: 15000 }
        );

        const status = await page.$eval('#status', el => el.textContent);
        console.log(`ESM Status: ${status}`);

        if (status.includes('FAIL') || status.includes('ERROR')) {
            console.error('❌ ESM Bundle Integration Tests Failed!');
            failed = true;
        } else {
            console.log('✅ ESM Bundle Integration Tests Passed!');
        }
    } catch (err) {
        console.error(`❌ ESM Test execution failed: ${err.message}`);
        failed = true;
    }

    // -----------------------------------------------------------------------
    // PART 2b: IIFE Slim Browser Bundle Tests
    // -----------------------------------------------------------------------
    try {
        console.log('\n--- Testing IIFE Slim Browser Bundle ---');
        await page.goto(`http://localhost:${port}/test-iife-slim`, { waitUntil: 'networkidle2' });

        await page.waitForFunction(
            () => document.getElementById('status').textContent !== 'RUNNING',
            { timeout: 15000 }
        );

        const status = await page.$eval('#status', el => el.textContent);
        console.log(`IIFE Slim Status: ${status}`);

        if (status.includes('FAIL') || status.includes('ERROR')) {
            console.error('❌ IIFE Slim Bundle Integration Tests Failed!');
            failed = true;
        } else {
            console.log('✅ IIFE Slim Bundle Integration Tests Passed!');
        }
    } catch (err) {
        console.error(`❌ IIFE Slim Test execution failed: ${err.message}`);
        failed = true;
    }

    // -----------------------------------------------------------------------
    // PART 2c: ESM Slim Browser Bundle Tests
    // -----------------------------------------------------------------------
    try {
        console.log('\n--- Testing ESM Slim Browser Bundle ---');
        await page.goto(`http://localhost:${port}/test-esm-slim`, { waitUntil: 'networkidle2' });

        await page.waitForFunction(
            () => document.getElementById('status').textContent !== 'RUNNING',
            { timeout: 15000 }
        );

        const status = await page.$eval('#status', el => el.textContent);
        console.log(`ESM Slim Status: ${status}`);

        if (status.includes('FAIL') || status.includes('ERROR')) {
            console.error('❌ ESM Slim Bundle Integration Tests Failed!');
            failed = true;
        } else {
            console.log('✅ ESM Slim Bundle Integration Tests Passed!');
        }
    } catch (err) {
        console.error(`❌ ESM Slim Test execution failed: ${err.message}`);
        failed = true;
    }

    await browser.close();
    server.close();

    // -----------------------------------------------------------------------
    // PART 3: Compiled CLI Standalone Binary Tests
    // -----------------------------------------------------------------------
    console.log('\n--- Testing Compiled CLI Standalone Binary ---');
    const binaryPath = path.join(ROOT, 'dist', 'bin', 'officeparser-macos-arm64');

    if (!fs.existsSync(binaryPath)) {
        console.warn(`⚠️ Warning: Compiled macOS binary not found at ${binaryPath}. Skipping CLI binary tests.`);
    } else {
        try {
            // Test DOCX parsing
            console.log('Running CLI Binary on Word Document...');
            const docxResult = child_process.spawnSync(binaryPath, ['test/files/test.docx', '--to=text'], { encoding: 'utf8' });
            if (docxResult.status !== 0) {
                console.error(`❌ CLI DOCX test failed with exit code ${docxResult.status}. Stderr: ${docxResult.stderr}`);
                failed = true;
            } else if (!docxResult.stdout.includes('Demonstration of DOCX support')) {
                console.error(`❌ CLI DOCX output does not match expectation. Stdout: ${docxResult.stdout.slice(0, 300)}`);
                failed = true;
            } else {
                console.log('✅ CLI DOCX Parsing Passed!');
            }

            // Test PDF parsing (verifying warning removal)
            console.log('Running CLI Binary on PDF Document (Verifying warning removal)...');
            const pdfResult = child_process.spawnSync(binaryPath, ['test/files/test.pdf', '--to=json'], { encoding: 'utf8' });
            if (pdfResult.status !== 0) {
                console.error(`❌ CLI PDF test failed with exit code ${pdfResult.status}. Stderr: ${pdfResult.stderr}`);
                failed = true;
            } else {
                try {
                    const json = JSON.parse(pdfResult.stdout);
                    const warnings = json.warnings || [];
                    const pdfWorkerWarnings = warnings.filter(w => w.code === 'PDF_WORKER_FALLBACK');
                    
                    if (pdfWorkerWarnings.length > 0) {
                        console.error('❌ CLI PDF warning verification failed! Found unexpected PDF_WORKER_FALLBACK warning:');
                        console.error(JSON.stringify(pdfWorkerWarnings, null, 2));
                        failed = true;
                    } else if (warnings.length > 0) {
                        console.log(`⚠️ CLI PDF returned warnings, but none were PDF_WORKER_FALLBACK:`, JSON.stringify(warnings, null, 2));
                        console.log('✅ CLI PDF Warning Resolution Passed!');
                    } else {
                        console.log('✅ CLI PDF Warning Resolution Passed (No warnings found at all)!');
                    }
                } catch (e) {
                    console.error(`❌ CLI PDF JSON parse failed: ${e.message}. Raw stdout: ${pdfResult.stdout.slice(0, 300)}`);
                    failed = true;
                }
            }
        } catch (err) {
            console.error(`❌ CLI binary tests execution failed: ${err.message}`);
            failed = true;
        }
    }

    if (failed) {
        console.error('\n❌ Integration Verification FAILED!');
        process.exit(1);
    } else {
        console.log('\n🎉 All Integration Verification Tests Passed Successfully!');
        process.exit(0);
    }
}

main().catch((err) => {
    console.error('Test execution error:', err);
    process.exit(1);
});

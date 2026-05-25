const puppeteer = require('puppeteer');
const http = require('http');
const fs = require('fs');
const path = require('path');

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

// Start a lightweight static file server on a dynamic port
const server = http.createServer((req, res) => {
    let urlPath = req.url.split('?')[0];
    if (urlPath === '/' || urlPath === '') {
        urlPath = '/docs/index.html';
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

server.listen(0, async () => {
    const port = server.address().port;
    console.log(`Temp server running at http://localhost:${port}`);

    console.log('Launching Puppeteer browser...');
    const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    const page = await browser.newPage();

    try {
        console.log('Navigating to visualizer page...');
        await page.goto(`http://localhost:${port}/docs/index.html`, { waitUntil: 'networkidle2' });

        console.log('Waiting for sample picker...');
        await page.waitForSelector('#sample-picker-grid', { timeout: 5000 });

        console.log('Locating test.xlsx button...');
        const xlsxButton = await page.evaluateHandle(() => {
            const buttons = document.querySelectorAll('.sample-btn');
            return Array.from(buttons).find(btn => btn.textContent.includes('test.xlsx'));
        });

        if (!xlsxButton) {
            throw new Error('test.xlsx button not found in sample picker grid');
        }

        console.log('Clicking test.xlsx button...');
        await xlsxButton.click();

        console.log('Waiting for parsing to complete...');
        await page.waitForFunction(
            () => {
                const el = document.getElementById('status-msg');
                return el && el.textContent.includes('Successfully parsed test.xlsx');
            },
            { timeout: 15000 }
        );

        console.log('Waiting for HTML preview iframe...');
        const iframeElement = await page.waitForSelector('#visual-html iframe');
        const frame = await iframeElement.contentFrame();

        console.log('Checking spreadsheet tabs container presence...');
        const tabsContainer = await frame.waitForSelector('.spreadsheet-tabs', { timeout: 5000 });
        if (!tabsContainer) {
            throw new Error('Spreadsheet tabs container (.spreadsheet-tabs) was not rendered!');
        }

        // --- Bounding Box & Visibility Check (Counter-Measure) ---
        console.log('Verifying that spreadsheet tabs are fully within the visible viewport (not cut off)...');
        const tabVisibilityInfo = await frame.evaluate(() => {
            const container = document.querySelector('.spreadsheet-tabs');
            if (!container) return { exists: false };

            const rect = container.getBoundingClientRect();
            const viewportHeight = window.innerHeight;
            const viewportWidth = window.innerWidth;

            // Bounding box checks
            const isWithinBounds = (
                rect.top >= 0 &&
                rect.left >= 0 &&
                rect.bottom <= viewportHeight &&
                rect.right <= viewportWidth
            );

            const style = window.getComputedStyle(container);
            const isNotHidden = (
                style.display !== 'none' &&
                style.visibility !== 'hidden' &&
                parseFloat(style.opacity) > 0 &&
                rect.width > 0 &&
                rect.height > 0
            );

            return {
                exists: true,
                rect: {
                    top: rect.top,
                    bottom: rect.bottom,
                    left: rect.left,
                    right: rect.right,
                    width: rect.width,
                    height: rect.height
                },
                viewport: {
                    height: viewportHeight,
                    width: viewportWidth
                },
                isWithinBounds,
                isNotHidden
            };
        });

        console.log('Visibility metrics:', JSON.stringify(tabVisibilityInfo, null, 2));

        if (!tabVisibilityInfo.isWithinBounds) {
            throw new Error(
                `Spreadsheet tab selector is out of viewport bounds! ` +
                `Bottom position is ${tabVisibilityInfo.rect.bottom}px but viewport height is only ${tabVisibilityInfo.viewport.height}px.`
            );
        }

        if (!tabVisibilityInfo.isNotHidden) {
            throw new Error('Spreadsheet tab selector is hidden via display, visibility, opacity or has 0 size!');
        }

        console.log('SUCCESS: Spreadsheet tab selector is fully visible and correctly rendered inside the viewport bounds!');

        // --- Verify Enlarge Modal ---
        console.log('Clicking Enlarge button on HTML preview window...');
        const enlargeBtn = await page.waitForSelector('#window-html .btn-maximize');
        await enlargeBtn.click();

        console.log('Waiting for modal to be active...');
        await page.waitForSelector('#preview-modal.active', { timeout: 3000 });

        console.log('Checking modal filename content...');
        const modalFilenameText = await page.$eval('#modal-filename', el => el.textContent);
        console.log(`Modal filename shows: "${modalFilenameText}"`);
        if (modalFilenameText !== 'test.xlsx') {
            throw new Error(`Expected modal filename to be "test.xlsx", got "${modalFilenameText}"`);
        }

        console.log('Clicking modal close button...');
        const modalCloseBtn = await page.waitForSelector('#modal-close-btn');
        await modalCloseBtn.click();

        console.log('Verifying that modal is closed...');
        await page.waitForFunction(
            () => !document.getElementById('preview-modal').classList.contains('active'),
            { timeout: 3000 }
        );
        console.log('SUCCESS: Enlarge modal opens, displays correct filename, and closes successfully!');

        process.exit(0);

    } catch (err) {
        console.error('TEST FAILED:', err.message);
        process.exit(1);
    } finally {
        await browser.close();
        server.close();
    }
});

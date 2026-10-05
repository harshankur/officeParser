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

        // --- The LaTeX preview is the page LaTeX typesets, not a web page ---
        console.log('Loading test.docx and switching the LaTeX window to its preview...');
        const docxButton = await page.evaluateHandle(() => Array.from(document.querySelectorAll('.sample-btn')).find(btn => btn.textContent.includes('test.docx')));
        await docxButton.click();
        await page.waitForFunction(() => document.getElementById('status-msg')?.textContent.includes('Successfully parsed test.docx'), { timeout: 30000 });
        await page.$eval('#toggle-tex .toggle-btn[data-state="preview"]', button => button.click());
        const texFrameElement = await page.waitForSelector('#visual-tex iframe', { timeout: 15000 });
        const texFrame = await texFrameElement.contentFrame();
        await texFrame.waitForSelector('.tex-sheet', { timeout: 15000 });

        // What the generated LaTeX asks for, read from the source the window made: A4 by its name and
        // an inch of margin all round, which is what the window's defaults give.
        const texSource = await page.evaluate(async () => (await window.currentAst.to('tex')).value);
        const geometry = /\\usepackage\[([^\]]*)\]\{geometry\}/.exec(texSource)[1];
        if (geometry !== 'a4paper,margin=1in') throw new Error(`The generated LaTeX's page is not the default one this test reads: ${geometry}`);
        const asked = name => ({ paperwidth: 595.28, top: 72, right: 72, bottom: 72, left: 72 })[name];
        const classSize = Number(/\\documentclass\[[^\]]*?(\d+)pt/.exec(texSource)[1]);
        const bodySize = { 10: 10, 11: 10.95, 12: 12 }[classSize];

        const typeset = await texFrame.evaluate(async () => {
            await document.fonts.ready;
            const sheet = document.querySelector('.tex-sheet');
            const style = getComputedStyle(sheet);
            const zoom = Number(document.documentElement.style.getPropertyValue('--tex-zoom')) || 1;
            const points = px => parseFloat(px) * 0.75;
            const title = document.querySelector('.tex-title');
            const paragraph = sheet.querySelector('article > p:not([class])');
            return {
                zoom,
                family: style.fontFamily,
                loaded: [...document.fonts].filter(face => face.status === 'loaded').map(face => face.family),
                size: points(style.fontSize),
                width: points(style.width),
                padding: [style.paddingTop, style.paddingRight, style.paddingBottom, style.paddingLeft].map(points),
                align: getComputedStyle(paragraph).textAlign,
                titleWeight: title ? getComputedStyle(title).fontWeight : null,
                titleAlign: title ? getComputedStyle(title).textAlign : null,
                titleSize: title ? points(getComputedStyle(title).fontSize) : null,
                overflows: document.documentElement.scrollWidth > document.documentElement.clientWidth + 1,
            };
        });
        console.log('Typeset preview:', JSON.stringify(typeset));
        const near = (value, wanted) => Math.abs(value - wanted) < 0.2;
        if (!/TeX Roman/.test(typeset.family)) throw new Error(`The LaTeX preview is not set in LaTeX's font: ${typeset.family}`);
        if (!near(typeset.size, bodySize)) throw new Error(`The LaTeX preview's text is ${typeset.size}pt; the source's class sets ${bodySize}pt`);
        if (!near(typeset.width, asked('paperwidth'))) throw new Error(`The LaTeX preview's paper is ${typeset.width}pt wide; the source asks for ${asked('paperwidth')}pt`);
        const margins = ['top', 'right', 'bottom', 'left'].map(asked);
        if (!typeset.padding.every((value, i) => near(value, margins[i]))) throw new Error(`The LaTeX preview's margins are ${typeset.padding}; the source asks for ${margins}`);
        if (typeset.align !== 'justify') throw new Error(`The LaTeX preview's paragraphs are not justified: ${typeset.align}`);
        if (typeset.titleWeight !== '400' || typeset.titleAlign !== 'center' || !near(typeset.titleSize, 17.28)) {
            throw new Error(`The LaTeX preview's title is not set as \\maketitle sets it: ${typeset.titleSize}pt, weight ${typeset.titleWeight}, ${typeset.titleAlign}`);
        }
        if (typeset.overflows) throw new Error('The LaTeX preview scrolls sideways: the page is wider than its window');
        console.log('SUCCESS: The LaTeX preview is set in LaTeX\'s font, at the size, on the paper and inside the margins the generated source asks for!');

        process.exit(0);

    } catch (err) {
        console.error('TEST FAILED:', err.message);
        process.exit(1);
    } finally {
        await browser.close();
        server.close();
    }
});

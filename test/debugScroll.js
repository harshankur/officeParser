// Scroll verification script for officeParser visualizer.
// Run using a custom npm script or node.
const puppeteer = require('puppeteer');
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');

const MIMES = {
    '.html': 'text/html',
    '.js': 'application/javascript',
    '.css': 'text/css',
    '.json': 'application/json',
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
};

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
    console.log(`Debug server running at http://localhost:${port}`);

    const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    const page = await browser.newPage();

    try {
        await page.goto(`http://localhost:${port}/docs/index.html`, { waitUntil: 'networkidle2' });

        const docxBtn = await page.evaluateHandle(() => {
            const buttons = document.querySelectorAll('.sample-btn');
            return Array.from(buttons).find(btn => btn.textContent.includes('test.docx'));
        });
        await docxBtn.click();

        await page.waitForFunction(
            () => {
                const el = document.getElementById('status-msg');
                return el && el.textContent.includes('Successfully parsed test.docx');
            },
            { timeout: 15000 }
        );

        const iframeElement = await page.waitForSelector('#visual-html iframe');
        
        await page.evaluate((el) => {
            el.scrollIntoView({ block: 'center' });
        }, iframeElement);
        
        await new Promise(r => setTimeout(r, 500));
        
        await page.evaluate(() => {
            const container = document.getElementById('visual-html');
            container.scrollTop = 0;
        });
        
        let box = await iframeElement.boundingBox();
        let hoverX = box.x + box.width / 2;
        let hoverY = box.y + 100;
        await page.mouse.move(hoverX, hoverY);
        await page.mouse.wheel({ deltaY: 300 });
        await new Promise(r => setTimeout(r, 500));

        const postParentScroll = await page.evaluate(() => {
            return document.getElementById('visual-html').scrollTop;
        });
        console.log('Post Parent Scroll:', postParentScroll);

        const enlargeBtn = await page.waitForSelector('#window-html .btn-maximize');
        await enlargeBtn.click();

        const modalIframeEl = await page.waitForSelector('#modal-body-content iframe');
        const modalFrame = await modalIframeEl.contentFrame();

        const modalBox = await modalIframeEl.boundingBox();
        const modalHoverX = modalBox.x + modalBox.width / 2;
        const modalHoverY = modalBox.y + modalBox.height / 2;
        await page.mouse.move(modalHoverX, modalHoverY);

        await page.mouse.wheel({ deltaY: 300 });
        await new Promise(r => setTimeout(r, 500));

        const postModalScroll = await modalFrame.evaluate(() => {
            return window.pageYOffset || document.documentElement.scrollTop || document.body.scrollTop;
        });
        console.log('Post Modal Scroll:', postModalScroll);

        if (postParentScroll > 0 && postModalScroll > 0) {
            console.log('SUCCESS: All scroll tests passed!');
            process.exit(0);
        } else {
            console.error('FAILURE: Scroll tests failed.');
            process.exit(1);
        }

    } catch (err) {
        console.error('ERROR:', err);
        process.exit(1);
    } finally {
        await browser.close();
        server.close();
    }
});

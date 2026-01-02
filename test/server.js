const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 3000;
const ROOT = path.join(__dirname, '..'); // Serve from root to access dist/ and test/

const MIMES = {
    '.html': 'text/html',
    '.js': 'application/javascript',
    '.mjs': 'application/javascript',
    '.css': 'text/css',
    '.json': 'application/json',
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.pdf': 'application/pdf'
};

const server = http.createServer((req, res) => {
    console.log(`${req.method} ${req.url}`);

    // Map root URL directly to the visualizer
    let urlPath = req.url;
    if (urlPath === '/') {
        urlPath = '/public/index.html';
    }

    // Sanitize path (prevent directory traversal)
    const safePath = path.normalize(urlPath).replace(/^(\.\.[\/\\])+/, '');
    let filePath = path.join(ROOT, safePath);

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
            'Access-Control-Allow-Origin': '*', // Enable CORS for local testing
            'Cross-Origin-Opener-Policy': 'same-origin',
            'Cross-Origin-Embedder-Policy': 'require-corp'
        });

        const stream = fs.createReadStream(filePath);
        stream.pipe(res);
    });
});

server.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}/`);
    console.log(`Root mapped to: http://localhost:${PORT}/public/index.html`);
});

#!/usr/bin/env node
/**
 * Copies the repository's own documents (the readme, the changelog, how to contribute, the security
 * policy, where to get support, the code of conduct and the licence) into the site, docs/repo/, where
 * the home page opens them in a reader of its own instead of sending a visitor to GitHub. The licence,
 * which has no extension, is served as LICENSE.txt, so it comes back as text.
 *
 * Run with --check to compare instead of copy: the site's copies must be the repository's documents as
 * they are, and a copy left behind (the changelog of the last release on a site that shows this one) is
 * an error, as `npm run test:site` reports it.
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const TARGET = path.join(ROOT, 'docs', 'repo');
/** Each document, and the name the site serves it by. */
const DOCUMENTS = [
    ['README.md', 'README.md'],
    ['CHANGELOG.md', 'CHANGELOG.md'],
    ['CONTRIBUTING.md', 'CONTRIBUTING.md'],
    ['SECURITY.md', 'SECURITY.md'],
    ['SUPPORT.md', 'SUPPORT.md'],
    ['CODE_OF_CONDUCT.md', 'CODE_OF_CONDUCT.md'],
    ['LICENSE', 'LICENSE.txt'],
];

const check = process.argv.includes('--check');
const stale = [];
if (!check) fs.mkdirSync(TARGET, { recursive: true });
for (const [source, served] of DOCUMENTS) {
    const from = fs.readFileSync(path.join(ROOT, source));
    const to = path.join(TARGET, served);
    if (check) {
        if (!fs.existsSync(to) || !fs.readFileSync(to).equals(from)) stale.push(`docs/repo/${served}`);
    } else {
        fs.writeFileSync(to, from);
    }
}

if (check) {
    if (stale.length) {
        console.error(`✗ The site's copies of the repository's documents are not the documents (run npm run sync:site): ${stale.join(', ')}`);
        process.exit(1);
    }
    console.log(`✓ The site's ${DOCUMENTS.length} repository documents are the repository's`);
} else {
    console.log(`Copied ${DOCUMENTS.length} repository documents into docs/repo/`);
}

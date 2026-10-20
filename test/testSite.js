/**
 * Site Pages Validation
 *
 * Every whole page of the site (docs/: the home page and the blog's pages, not the spec fragments the
 * home page loads, the sample documents in docs/test/ or the bundles in docs/dist/ and docs/vendor/)
 * loads the view counter's consent prompt (consent.js) once, and that prompt names who runs the site
 * (data-operator). Consent is only informed when the visitor knows who is asking, and who runs the
 * site comes first in what a consent request has to say. The blog's pages are built from
 * blog/templates/layout.html, so the template is checked too: a page rebuilt from it keeps the name.
 * It reads files only: no browser is needed.
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const SITE_DIR = path.join(ROOT, 'docs');
/** Folders of docs/ that hold files the site serves, not pages of it. */
const NOT_PAGES = ['test', 'dist', 'vendor', 'specs'];
const LAYOUT = path.join(ROOT, 'blog', 'templates', 'layout.html');
/** Who runs the site, as the shared legal notice names him. */
const OPERATOR = 'Harsh Ankur';
const CONSENT_SCRIPT = /<script\b[^>]*\bsrc="https:\/\/views\.harshankur\.com\/consent\.js"[^>]*>/g;

let passed = 0;
let failed = 0;
const check = (name, ok, detail = '') => {
    if (ok) { passed++; return; }
    failed++;
    console.log(`  ✗ FAIL: ${name}${detail ? ` (${detail})` : ''}`);
};

/** Every whole page under a folder of the site, by its path from the repository's root. */
function pagesIn(dir, top = true) {
    return fs.readdirSync(dir, { withFileTypes: true }).flatMap(entry => {
        const file = path.join(dir, entry.name);
        if (entry.isDirectory()) return top && NOT_PAGES.includes(entry.name) ? [] : pagesIn(file, false);
        return entry.name.endsWith('.html') && /<html[\s>]/i.test(fs.readFileSync(file, 'utf8')) ? [file] : [];
    });
}

/** The value of an attribute in a tag, or null when the tag has none. */
const attribute = (tag, name) => new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1] ?? null;

const pages = pagesIn(SITE_DIR).map(file => path.relative(ROOT, file)).sort();
check('the site has its home page and the blog\'s pages', pages.includes('docs/index.html') && pages.includes(path.join('docs', 'blog', 'index.html')), pages.join(', '));

for (const file of [...pages, path.relative(ROOT, LAYOUT)]) {
    const tags = fs.readFileSync(path.join(ROOT, file), 'utf8').match(CONSENT_SCRIPT) || [];
    check(`${file}: loads consent.js once`, tags.length === 1, `${tags.length} times`);
    for (const tag of tags) {
        check(`${file}: consent.js names who runs the site (data-operator="${OPERATOR}")`, attribute(tag, 'data-operator') === OPERATOR, String(attribute(tag, 'data-operator')));
    }
}

console.log(failed === 0 ? `✓ Site pages: ${passed} passed, 0 failed` : `✗ Site pages: ${passed} passed, ${failed} failed`);
process.exit(failed === 0 ? 0 : 1);

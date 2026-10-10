/**
 * Site Version Stamp
 *
 * The home page (docs/index.html) is written by hand, with one thing in it that is the build's: the
 * package's version, shown as a badge in the hero. It is written here from package.json, between the
 * two `package-version` comments, so the page never names a version the package is not at.
 *
 *   node scripts/sync-site-version.js           writes the badge (a step of `npm run build`)
 *   node scripts/sync-site-version.js --check   fails when the page's badge is not package.json's
 *                                               version: the page was not built, or not committed,
 *                                               after the version changed (`npm run test:site`)
 *   node scripts/sync-site-version.js --stage   stages the page for the commit being made when the
 *                                               badge is all that changed in it (the pre-commit hook,
 *                                               after the build): a version bump takes its badge along
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PAGE = path.join(ROOT, 'docs', 'index.html');

/** The comments the badge stands between. What is between them is replaced whole. */
const START = '<!-- package-version:start (written from package.json by scripts/sync-site-version.js: do not edit) -->';
const END = '<!-- package-version:end -->';
/** A version as npm writes one (`8.1.1`, `9.0.0-beta.2`): nothing in it needs escaping in HTML. */
const VERSION = /^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/;

/** The package's version, from package.json. */
function packageVersion() {
    const { version } = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
    if (typeof version !== 'string' || !VERSION.test(version)) throw new Error(`package.json: "${version}" is not a version`);
    return version;
}

/** The badge for a version: the version, linked to what changed in it. */
function versionBadge(version, indent) {
    return [
        START,
        `<a class="badge-link" href="https://github.com/harshankur/officeParser/blob/master/CHANGELOG.md" target="_blank"`,
        `    rel="noopener" aria-label="officeParser version ${version}: see what changed">`,
        `    <div class="badge badge-version" id="version-badge">v${version}</div>`,
        `</a>`,
        END,
    ].join(`\n${indent}`);
}

/**
 * The page with its badge at `version`. The two comments are in the page once, in order: a page
 * without them has no place for the badge, which is an error and not a page left as it is.
 */
function stampVersion(html, version) {
    const start = html.indexOf(START), end = html.indexOf(END);
    if (start === -1 || end < start || html.indexOf(START, start + 1) !== -1 || html.indexOf(END, end + 1) !== -1) {
        throw new Error('docs/index.html: the package-version comments are not there once each, in order');
    }
    const indent = /[ \t]*$/.exec(html.slice(0, start))[0];
    return html.slice(0, start) + versionBadge(version, indent) + html.slice(end + END.length);
}

/**
 * Stages the page when what waits unstaged in it is the badge alone: the page as it is staged, with
 * its badge at the package's version, is then the page on disk. A page with other changes waiting is
 * left for whoever is making them to stage: they are no part of this commit unless they say so.
 */
function stageBadge(version) {
    const git = (...args) => execFileSync('git', args, { cwd: ROOT, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
    const onDisk = fs.readFileSync(PAGE, 'utf8');
    let staged;
    try { staged = git('show', ':docs/index.html'); } catch { return 'docs/index.html is not in the repository yet'; }
    if (staged === onDisk) return null;
    let stamped;
    try { stamped = stampVersion(staged, version); } catch { return 'docs/index.html has other changes that are not staged; its version badge was left with them'; }
    if (stamped !== onDisk) return 'docs/index.html has other changes that are not staged; its version badge was left with them';
    git('add', 'docs/index.html');
    return `docs/index.html staged: its version badge is now v${version}`;
}

module.exports = { PAGE, packageVersion, stampVersion };

if (require.main === module) {
    try {
        const version = packageVersion();
        const html = fs.readFileSync(PAGE, 'utf8');
        const stamped = stampVersion(html, version);
        if (process.argv.includes('--stage')) {
            const said = stageBadge(version);
            if (said) console.log(`Site version: ${said}`);
        } else if (process.argv.includes('--check')) {
            if (stamped !== html) {
                console.error(`✗ Site: the home page's version badge is not package.json's version (${version}). Run \`npm run build\` and commit docs/index.html.`);
                process.exit(1);
            }
            console.log(`✓ Site: the home page shows version ${version}`);
        } else {
            if (stamped !== html) fs.writeFileSync(PAGE, stamped);
            console.log(`Site version → docs/index.html (v${version})`);
        }
    } catch (error) {
        console.error(`Site version failed: ${error.message}`);
        process.exit(1);
    }
}

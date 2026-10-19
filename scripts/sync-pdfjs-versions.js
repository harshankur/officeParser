const fs = require('fs');
const path = require('path');

/**
 * Synchronization Script for PDF.js Version
 *
 * This script ensures that all documentation, examples, and type definitions
 * use the same version of pdfjs-dist as specified in package.json.
 *
 * PIN RATIONALE (do not bump pdfjs-dist casually): pdfjs-dist is pinned to an EXACT version, not a
 * range, because its output feeds every committed PDF baseline and its tagged-structure API is
 * behavior-sensitive. 6.3.289 was tried and REVERTED, but the reason is narrower than it first
 * looks: on 6.3.289 `getMarkInfo()` returns `{}` for a marked PDF (6.2.108 returns
 * `{ Marked: true }`), so our tagged-detection (which keys off `getMarkInfo().Marked` in
 * PdfParser.ts) collapses to the geometry fallback and six PDF smoke assertions fail. The tagged
 * STRUCTURE itself is not lost: `getStructTree()` on 6.3.289 is byte-identical to 6.2.108 (verified
 * on test.pdf: same 1092 struct nodes, same H1/P/Table/TR/TH/TD/L/LI/TOC/Note/Figure/Link roles).
 * So the durable fix is to derive tagged-ness from the presence of a real struct tree (keeping
 * `getMarkInfo().Suspects` for the trust gate when present) instead of from `getMarkInfo().Marked`;
 * do that decoupling in PdfParser, THEN bump. Note also that 6.3.289 is currently the newest release
 * (no upstream fix exists) and is a lateral move plus a known regression, so staying on 6.2.108 is
 * correct until a genuinely newer pdfjs ships. Before raising the version: decouple detection as
 * above, run `npm run test:parser` (full, not fast), confirm the tagged-PDF assertions still pass,
 * and regenerate the PDF baselines. Bump with: `node scripts/sync-pdfjs-versions.js <version>`.
 */

function syncVersions() {
    const args = process.argv.slice(2);
    const versionFromArg = args.find(arg => !arg.startsWith('--'));

    const packageJsonPath = path.join(__dirname, '../package.json');
    if (!fs.existsSync(packageJsonPath)) {
        console.error('Error: package.json not found');
        process.exit(1);
    }

    let packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));

    // If version is provided as an argument, update package.json first
    if (versionFromArg) {
        console.log(`Updating package.json pdfjs-dist dependency to: ${versionFromArg}`);
        packageJson.dependencies['pdfjs-dist'] = versionFromArg;
        fs.writeFileSync(packageJsonPath, JSON.stringify(packageJson, null, 2) + '\n', 'utf8');
    }

    const pdfJsDependency = packageJson.dependencies['pdfjs-dist'];

    if (!pdfJsDependency) {
        console.error('Error: pdfjs-dist not found in dependencies');
        process.exit(1);
    }

    // Strip characters like ^, ~ from version
    const pdfJsVersion = pdfJsDependency.replace(/[\^~]/g, '');
    const versionRegex = /pdfjs-dist@[\d.]+/g;
    const targetVersionString = `pdfjs-dist@${pdfJsVersion}`;

    console.log(`Syncing PDF.js version: ${pdfJsVersion}`);

    const filesToUpdate = [
        'src/types.ts',
        'src/defaults.ts',
        'docs/index.html',
        'docs/specs/debugging_fragment.html',
        'README.md'
    ];

    let updatedCount = 0;
    const handledFiles = new Set(filesToUpdate.map(f => path.normalize(f)));

    // 1. Perform Sync
    filesToUpdate.forEach(relativePath => {
        const filePath = path.join(__dirname, '..', relativePath);
        if (fs.existsSync(filePath)) {
            const content = fs.readFileSync(filePath, 'utf8');
            
            // Standard replacement (pdfjs-dist@x.y.z)
            let newContent = content.replace(versionRegex, targetVersionString);
            
            // Special replacement for src/defaults.ts (PDFJS_VERSION constant)
            if (relativePath === 'src/defaults.ts') {
                const rawVersionRegex = /const PDFJS_VERSION = '[\d.]+';/;
                const targetRawVersionString = `const PDFJS_VERSION = '${pdfJsVersion}';`;
                newContent = newContent.replace(rawVersionRegex, targetRawVersionString);
            }

            if (content !== newContent) {
                fs.writeFileSync(filePath, newContent, 'utf8');
                console.log(`  ✓ Updated ${relativePath}`);
                updatedCount++;
            } else {
                console.log(`  - ${relativePath} is already up to date`);
            }
        } else {
            console.warn(`  ! Warning: File not found: ${relativePath}`);
        }
    });

    // The site loads the pdf.js worker from a copy vendored under docs/vendor/pdfjs-dist-<version>/: the
    // page names that folder, and pdf.js refuses a worker of another version than its own. So the page is
    // pointed at the folder for this version, and the build stops when that folder is not vendored yet.
    const vendored = path.join(__dirname, '..', 'docs', 'vendor', `pdfjs-dist-${pdfJsVersion}`, 'pdf.worker.min.mjs');
    const sitePage = path.join(__dirname, '..', 'docs', 'index.html');
    if (fs.existsSync(sitePage)) {
        const page = fs.readFileSync(sitePage, 'utf8');
        const pointed = page.replace(/vendor\/pdfjs-dist-[\d.]+\//g, `vendor/pdfjs-dist-${pdfJsVersion}/`);
        if (pointed !== page) {
            fs.writeFileSync(sitePage, pointed, 'utf8');
            console.log(`  ✓ Pointed docs/index.html at vendor/pdfjs-dist-${pdfJsVersion}/`);
            updatedCount++;
        }
    }
    if (!fs.existsSync(vendored)) {
        console.error(`\n[!] docs/vendor/pdfjs-dist-${pdfJsVersion}/pdf.worker.min.mjs is missing: copy node_modules/pdfjs-dist/build/pdf.worker.min.mjs`
            + ` (and its LICENSE) there, and remove the old version's folder. The site would otherwise load a worker of another version.`);
        process.exitCode = 1;
    }

    // 2. Scan for unhandled occurrences
    console.log('\nScanning for unhandled PDF.js version occurrences...');
    const rootDir = path.join(__dirname, '..');
    const ignoreDirs = ['node_modules', 'dist', '.git', 'test/results', 'release_assets', 'scratch'];
    // Files that legitimately carry a pdfjs version string but must NOT be synced. CHANGELOG.md
    // records the version shipped with each release as a historical fact, so rewriting it on a future
    // bump would falsify past entries. (release_assets/ and scratch/ are excluded wholesale above:
    // the former are frozen release artifacts pinned to a past release's pdfjs, the latter is scratch.)
    const ignoreFiles = new Set(['CHANGELOG.md'].map(f => path.normalize(f)));
    const unhandledFiles = [];

    function scanDir(currentDir) {
        const files = fs.readdirSync(currentDir);
        for (const file of files) {
            const fullPath = path.join(currentDir, file);
            const relPath = path.relative(rootDir, fullPath);

            if (fs.statSync(fullPath).isDirectory()) {
                if (!ignoreDirs.includes(file)) {
                    scanDir(fullPath);
                }
                continue;
            }

            // Only scan text-like files
            if (!/\.(ts|js|html|md|json|txt|css)$/.test(file)) continue;

            // Skip the package files and the script itself
            if (relPath === 'package.json' || relPath === 'package-lock.json') continue;
            if (relPath === 'scripts/sync-pdfjs-versions.js') continue;
            // Skip files that intentionally record a historical/pinned pdfjs version.
            if (ignoreFiles.has(path.normalize(relPath))) continue;

            try {
                const content = fs.readFileSync(fullPath, 'utf8');
                // Use a fresh regex for test to avoid lastIndex issues
                if (/pdfjs-dist@[\d.]+/.test(content) && !handledFiles.has(path.normalize(relPath))) {
                    unhandledFiles.push(relPath);
                }
            } catch (err) {
                console.warn(`  ! Warning: Could not read ${relPath}: ${err.message}`);
            }
        }
    }

    try {
        scanDir(rootDir);
    } catch (err) {
        console.error(`Error during scan: ${err.message}`);
    }

    if (unhandledFiles.length > 0) {
        console.warn('\n[!] WARNING: Found PDF.js version strings in files not handled by this script:');
        unhandledFiles.forEach(f => console.warn(`    - ${f}`));
        console.warn('Please add these files to the "filesToUpdate" list in scripts/sync-pdfjs-versions.js if they need synchronization.\n');
    } else {
        console.log('  ✓ No unhandled occurrences found.');
    }

    console.log(`\nSync complete. ${updatedCount} file(s) updated.`);
}

syncVersions();

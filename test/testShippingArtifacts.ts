/**
 * Shipping Artifact Validation Suite
 *
 * Validates every artifact that officeparser ships to ensure they are correctly
 * formed and consumable before the package is published. This runs as part of
 * `npm test` (after the build step, before the parser tests).
 *
 * Checks:
 *  - Node.js CJS package   (dist/index.js)
 *  - Node.js ESM package   (dist/index.mjs) — via spawned subprocess
 *  - CLI entry             (dist/cli.js)
 *  - Browser IIFE bundle   (dist/officeparser.browser.iife.js)
 *  - Browser ESM bundle    (dist/officeparser.browser.mjs)
 *  - Browser type decls    (dist/officeparser.browser.d.ts)
 *  - package.json paths    (all "exports", "main", "module", etc.)
 */

import * as fs from 'fs';
import * as path from 'path';
import * as child_process from 'child_process';
import { fileURLToPath, pathToFileURL } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT = path.join(__dirname, '..');

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

interface CheckResult {
    name: string;
    status: 'PASS' | 'FAIL' | 'SKIP';
    detail: string;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function pass(name: string, detail = ''): CheckResult {
    return { name, status: 'PASS', detail };
}

function fail(name: string, detail: string): CheckResult {
    return { name, status: 'FAIL', detail };
}

function fileExists(relPath: string): boolean {
    return fs.existsSync(path.join(ROOT, relPath));
}

function readFile(relPath: string): string {
    return fs.readFileSync(path.join(ROOT, relPath), 'utf8');
}

function fileSize(relPath: string): number {
    return fs.statSync(path.join(ROOT, relPath)).size;
}

// ---------------------------------------------------------------------------
// Section 1: Node.js CJS Package
// ---------------------------------------------------------------------------

function checkCjs(): CheckResult[] {
    const results: CheckResult[] = [];
    const distPath = path.join(ROOT, 'dist', 'index.js');

    // Existence
    if (!fileExists('dist/index.js')) {
        return [fail('CJS: dist/index.js exists', 'File not found')];
    }
    results.push(pass('CJS: dist/index.js exists'));

    if (!fileExists('dist/index.d.ts')) {
        results.push(fail('CJS: dist/index.d.ts exists', 'File not found'));
    } else {
        results.push(pass('CJS: dist/index.d.ts exists'));
    }

    // No shebang in library entry (shebang should only be in cli.js)
    const content = readFile('dist/index.js');
    if (content.startsWith('#!')) {
        results.push(fail('CJS: dist/index.js has no shebang', 'Found shebang — library should not have one'));
    } else {
        results.push(pass('CJS: dist/index.js has no shebang'));
    }

    // require() resolves
    let mod: any;
    try {
        mod = require(distPath);
        results.push(pass('CJS: require() resolves'));
    } catch (e) {
        results.push(fail('CJS: require() resolves', String(e)));
        return results;
    }

    // Named exports
    if (typeof mod.OfficeParser === 'function' || typeof mod.OfficeParser === 'object') {
        results.push(pass('CJS: named export OfficeParser', typeof mod.OfficeParser));
    } else {
        results.push(fail('CJS: named export OfficeParser', `Got: ${typeof mod.OfficeParser}`));
    }

    if (typeof mod.parseOffice === 'function') {
        results.push(pass('CJS: named export parseOffice', 'function'));
    } else {
        results.push(fail('CJS: named export parseOffice', `Got: ${typeof mod.parseOffice}`));
    }

    if (typeof mod.OfficeGenerator === 'function' || typeof mod.OfficeGenerator === 'object') {
        results.push(pass('CJS: named export OfficeGenerator', typeof mod.OfficeGenerator));
    } else {
        results.push(fail('CJS: named export OfficeGenerator', `Got: ${typeof mod.OfficeGenerator}`));
    }

    if (mod.OfficeGenerator && typeof mod.OfficeGenerator.generate === 'function') {
        results.push(pass('CJS: OfficeGenerator.generate is a function'));
    } else {
        results.push(fail('CJS: OfficeGenerator.generate is a function', `Got: ${typeof mod.OfficeGenerator?.generate}`));
    }

    // Default export
    const defaultExport = mod.default ?? mod;
    if (defaultExport && typeof defaultExport.parseOffice === 'function') {
        results.push(pass('CJS: OfficeParser.parseOffice is a function'));
    } else {
        results.push(fail('CJS: OfficeParser.parseOffice is a function', `Got: ${typeof defaultExport?.parseOffice}`));
    }

    return results;
}

// ---------------------------------------------------------------------------
// Section 2: Node.js ESM Package (spawned subprocess)
// ---------------------------------------------------------------------------

function checkEsm(): CheckResult[] {
    const results: CheckResult[] = [];

    if (!fileExists('dist/index.mjs')) {
        return [fail('ESM: dist/index.mjs exists', 'File not found')];
    }
    results.push(pass('ESM: dist/index.mjs exists'));

    // Spawn the ESM test helper as a true ESM subprocess
    const helperPath = path.join(__dirname, 'esm-test.mjs');
    if (!fs.existsSync(helperPath)) {
        results.push(fail('ESM: subprocess helper exists', `test/esm-test.mjs not found`));
        return results;
    }

    let stdout = '';
    let exitCode = 0;
    try {
        const result = child_process.spawnSync(process.execPath, [helperPath], {
            encoding: 'utf8',
            timeout: 30000,
        });
        stdout = result.stdout ?? '';
        exitCode = result.status ?? 1;

        if (result.error) {
            results.push(fail('ESM: subprocess ran', String(result.error)));
            return results;
        }
    } catch (e) {
        results.push(fail('ESM: subprocess ran', String(e)));
        return results;
    }

    // Parse subprocess results
    try {
        const subResults: Array<{ name: string; status: string; detail: string }> = JSON.parse(stdout);
        for (const r of subResults) {
            results.push({
                name: `ESM: ${r.name}`,
                status: r.status === 'PASS' ? 'PASS' : 'FAIL',
                detail: r.detail ?? '',
            });
        }
    } catch {
        // If JSON parse fails, still report exit code
        if (exitCode !== 0) {
            results.push(fail('ESM: subprocess output', `Non-zero exit (${exitCode}), stdout: ${stdout.slice(0, 200)}`));
        }
    }

    return results;
}

// ---------------------------------------------------------------------------
// Section 3: CLI Entry
// ---------------------------------------------------------------------------

function checkCli(): CheckResult[] {
    const results: CheckResult[] = [];

    if (!fileExists('dist/cli.js')) {
        return [fail('CLI: dist/cli.js exists', 'File not found')];
    }
    results.push(pass('CLI: dist/cli.js exists'));

    const content = readFile('dist/cli.js');

    // Must have shebang
    if (content.startsWith('#!/usr/bin/env node')) {
        results.push(pass('CLI: has shebang #!/usr/bin/env node'));
    } else {
        results.push(fail('CLI: has shebang #!/usr/bin/env node', `First line: ${content.split('\n')[0].slice(0, 60)}`));
    }

    // Invoke CLI with no args — should print usage and exit 0
    try {
        const result = child_process.spawnSync(process.execPath, [path.join(ROOT, 'dist', 'cli.js')], {
            encoding: 'utf8',
            timeout: 10000,
        });
        if (result.stdout.includes('Usage') || result.stdout.includes('officeparser')) {
            results.push(pass('CLI: prints usage when invoked without args'));
        } else {
            results.push(fail('CLI: prints usage when invoked without args', `stdout: ${result.stdout.slice(0, 200)}`));
        }
    } catch (e) {
        results.push(fail('CLI: invocation', String(e)));
    }

    return results;
}

// ---------------------------------------------------------------------------
// Section 4: Browser IIFE Bundle
// ---------------------------------------------------------------------------

function checkBrowserIife(isSlim = false): CheckResult[] {
    const results: CheckResult[] = [];
    const suffix = isSlim ? '.slim' : '';
    const relPath = `dist/officeparser.browser${suffix}.iife.js`;
    const label = isSlim ? 'IIFE Slim' : 'IIFE';

    if (!fileExists(relPath)) {
        return [fail(`${label}: ${relPath} exists`, 'File not found')];
    }
    results.push(pass(`${label}: exists`));

    const content = readFile(relPath);
    const size = fileSize(relPath);

    // No shebang
    if (content.startsWith('#!')) {
        results.push(fail(`${label}: no shebang`, 'Bundle starts with shebang — Vite will throw SyntaxError'));
    } else {
        results.push(pass(`${label}: no shebang`));
    }

    // Has module.exports (UMD footer)
    if (content.includes('module.exports')) {
        results.push(pass(`${label}: has module.exports (UMD footer)`));
    } else {
        results.push(fail(`${label}: has module.exports (UMD footer)`, 'Missing — Vite __commonJS wrapper will get empty object'));
    }

    // Has IIFE assignment
    if (content.includes('officeParser')) {
        results.push(pass(`${label}: has globalName officeParser`));
    } else {
        results.push(fail(`${label}: has globalName officeParser`, 'IIFE global not found'));
    }

    // Has OfficeGenerator
    if (content.includes('OfficeGenerator')) {
        results.push(pass(`${label}: contains OfficeGenerator export`));
    } else {
        results.push(fail(`${label}: contains OfficeGenerator export`, 'OfficeGenerator not found in bundle'));
    }

    // Has @vite-ignore
    if (content.includes('@vite-ignore')) {
        results.push(pass(`${label}: has @vite-ignore for pdfjs dynamic import`));
    } else {
        results.push(fail(`${label}: has @vite-ignore for pdfjs dynamic import`, 'Missing — Vite will warn about unanalyzable dynamic import'));
    }

    // Reasonable size (must be > 100KB and < 10MB)
    const sizeMb = (size / 1024 / 1024).toFixed(2);
    if (size > 100 * 1024 && size < 10 * 1024 * 1024) {
        results.push(pass(`${label}: size is reasonable`, `${sizeMb} MB`));
    } else {
        results.push(fail(`${label}: size is reasonable`, `${sizeMb} MB — expected between 100KB and 10MB`));
    }

    return results;
}

// ---------------------------------------------------------------------------
// Section 5: Browser ESM Bundle
// ---------------------------------------------------------------------------

function checkBrowserEsm(isSlim = false): CheckResult[] {
    const results: CheckResult[] = [];
    const suffix = isSlim ? '.slim' : '';
    const relPath = `dist/officeparser.browser${suffix}.mjs`;
    const label = isSlim ? 'Browser ESM Slim' : 'Browser ESM';

    if (!fileExists(relPath)) {
        return [fail(`${label}: ${relPath} exists`, 'File not found')];
    }
    results.push(pass(`${label}: exists`));

    const content = readFile(relPath);
    const size = fileSize(relPath);

    // No shebang
    if (content.startsWith('#!')) {
        results.push(fail(`${label}: no shebang`, 'Bundle starts with shebang'));
    } else {
        results.push(pass(`${label}: no shebang`));
    }

    // Has export statements (ESM)
    if (/\bexport\b/.test(content)) {
        results.push(pass(`${label}: has export statements`));
    } else {
        results.push(fail(`${label}: has export statements`, 'No export keyword found — not a valid ESM module'));
    }

    // Does NOT have module.exports (should be ESM, not CJS)
    if (content.includes('module.exports')) {
        results.push(fail(`${label}: no module.exports`, 'Found module.exports in an ESM bundle'));
    } else {
        results.push(pass(`${label}: no module.exports`));
    }

    // Has @vite-ignore
    if (content.includes('@vite-ignore')) {
        results.push(pass(`${label}: has @vite-ignore for pdfjs dynamic import`));
    } else {
        results.push(fail(`${label}: has @vite-ignore for pdfjs dynamic import`, 'Missing — Vite will warn about unanalyzable dynamic import'));
    }

    // Reasonable size
    const sizeMb = (size / 1024 / 1024).toFixed(2);
    if (size > 100 * 1024 && size < 10 * 1024 * 1024) {
        results.push(pass(`${label}: size is reasonable`, `${sizeMb} MB`));
    } else {
        results.push(fail(`${label}: size is reasonable`, `${sizeMb} MB — expected between 100KB and 10MB`));
    }

    return results;
}

/**
 * A string that only `file-type` contributes, used to prove it was inlined into a bundle.
 *
 * The browser build has no module resolver at runtime, so `file-type` has to be baked in at
 * build time. That only happens because `moduleLoader.ts` writes the specifier as a literal
 * `import('file-type')` that esbuild can see, while the Node path deliberately hides the same
 * string inside `String(...)` so it is left alone. Making the browser specifier computed, for
 * instance by hoisting it into a variable, silently drops the module from the bundle and type
 * detection fails at runtime rather than at build time.
 *
 * The marker is deliberately something no office format mentions: our own source names OOXML
 * content types, so any of those would also match our code and prove nothing.
 */
const FILE_TYPE_INLINE_MARKER = 'META-INF/mozilla.rsa';

/**
 * Node built-ins that must not survive as import specifiers in a browser bundle.
 *
 * A consumer's bundler resolves every specifier it can see, whether or not the surrounding code
 * would ever run in a browser, so a leftover Node built-in fails their build even though our
 * code paths are gated on running under Node.
 */
const NODE_ONLY_SPECIFIERS = ['child_process', 'fs', 'url', 'puppeteer'];

/**
 * Exercises the dynamic-import classifier against the shapes it has to tell apart.
 *
 * The bundle checks below run the same classifier the build uses to place its annotations, so
 * on their own they cannot see a classifier that is wrong in both places at once: that is
 * exactly how issue #108 shipped, with a rule that mistook an interpolated template literal for
 * a static string. These pin the rule itself against fixed inputs, independently of any bundle.
 */
function checkDynamicImportClassifier(): CheckResult[] {
    const { findUnanalyzableDynamicImports } = require('../scripts/dynamicImports.js');
    const results: CheckResult[] = [];

    const cases: Array<{ source: string, unanalyzable: boolean, why: string }> = [
        { source: 'await import("file-type")', unanalyzable: false, why: 'a quoted specifier is resolvable' },
        { source: "await import('pdfjs-dist')", unanalyzable: false, why: 'single quotes too' },
        { source: 'await import(`pdfjs-dist`)', unanalyzable: false, why: 'a template with no substitution is still fixed' },
        // The shape that broke webpack: a template that interpolates cannot be resolved.
        { source: 'await import(`${base}${this._noWasmFilename}`)', unanalyzable: true, why: 'an interpolated template is not resolvable' },
        { source: 'await import(`${dir}/worker.mjs`)', unanalyzable: true, why: 'interpolation anywhere makes it dynamic' },
        { source: 'await import(specifier)', unanalyzable: true, why: 'a bare identifier is not resolvable' },
        { source: 'await import(getName(a, b))', unanalyzable: true, why: 'a call expression is not resolvable' },
        { source: 'await import(/* webpackIgnore: true */ `${x}`)', unanalyzable: false, why: 'already annotated' },
        { source: 'await import(/* @vite-ignore */ specifier)', unanalyzable: false, why: 'already annotated' },
    ];

    for (const { source, unanalyzable, why } of cases) {
        const found = findUnanalyzableDynamicImports(source).length > 0;
        const label = `Dynamic import rule: ${source}`;
        if (found === unanalyzable) results.push(pass(label, why));
        else results.push(fail(label, `expected ${unanalyzable ? 'unanalyzable' : 'resolvable'} (${why})`));
    }

    return results;
}

/**
 * Guards the two ways a browser bundle can break a consumer's build (issue #108).
 *
 * A dynamic import a bundler cannot resolve is not skipped by webpack: it builds a context
 * module over the whole directory the expression might reach, which for a published package
 * means every file in `dist/`, Node-only ones included. And a bare Node built-in specifier is
 * reported as a missing module. Neither shows up in our own tests, since we do not bundle the
 * bundles; only a consumer sees it.
 */
function checkBundlerCompatibility(relPath: string, label: string): CheckResult[] {
    if (!fileExists(relPath)) {
        return [fail(`${label}: ${relPath} exists`, 'File not found')];
    }

    const results: CheckResult[] = [];
    const content = readFile(relPath);

    // Same classifier the build uses to add the annotations, so the two cannot disagree.
    const { findUnanalyzableDynamicImports } = require('../scripts/dynamicImports.js');
    const unannotated = findUnanalyzableDynamicImports(content);
    if (unannotated.length === 0) {
        results.push(pass(`${label}: every dynamic import is bundler-safe`));
    } else {
        results.push(fail(`${label}: every dynamic import is bundler-safe`,
            `${unannotated.length} unannotated dynamic import(s), first: ${unannotated[0].snippet}`));
    }

    const leaked = NODE_ONLY_SPECIFIERS.filter(name =>
        content.includes(`import("${name}")`) || content.includes(`import('${name}')`)
        || content.includes(`require("${name}")`) || content.includes(`require('${name}')`));
    if (leaked.length === 0) {
        results.push(pass(`${label}: no Node-only module specifiers`));
    } else {
        results.push(fail(`${label}: no Node-only module specifiers`,
            `bundle still references ${leaked.join(', ')} — a consumer's bundler cannot resolve these`));
    }

    return results;
}

function checkFileTypeInlined(relPath: string, label: string): CheckResult[] {
    if (!fileExists(relPath)) {
        return [fail(`${label}: ${relPath} exists`, 'File not found')];
    }

    if (readFile(relPath).includes(FILE_TYPE_INLINE_MARKER)) {
        return [pass(`${label}: file-type is inlined for buffer detection`)];
    }
    return [fail(`${label}: file-type is inlined for buffer detection`,
        `Marker ${JSON.stringify(FILE_TYPE_INLINE_MARKER)} not found — the browser bundle cannot detect a file type from a buffer`)];
}

/**
 * Smoke-imports a browser ESM bundle in a real Node subprocess (4.F.6 / 6.E.3). Every other bundle
 * check here is static string matching, which cannot catch a bundle that parses but throws on
 * evaluation - a bad top-level reference, or a dynamic-import annotation the runtime rejects.
 * Importing it for real and confirming the documented entry points are present closes that gap.
 * (A browser bundle is still valid Node ESM: its browser-only code paths are gated at call time,
 * so module evaluation and export access work under Node.)
 */
function checkBrowserEsmExecutes(isSlim = false): CheckResult[] {
    const suffix = isSlim ? '.slim' : '';
    const relPath = `dist/officeparser.browser${suffix}.mjs`;
    const label = isSlim ? 'Browser ESM Slim' : 'Browser ESM';

    if (!fileExists(relPath)) {
        return [fail(`${label}: smoke-imports as an ESM module`, `${relPath} not found`)];
    }

    const url = pathToFileURL(path.join(ROOT, relPath)).href;
    const script =
        `import(${JSON.stringify(url)}).then(m => {` +
        `const ok = typeof m.parseOffice === 'function' && typeof m.OfficeGenerator?.generate === 'function';` +
        `process.stdout.write(ok ? 'OK' : 'MISSING:' + typeof m.parseOffice + ',' + typeof m.OfficeGenerator);` +
        `}).catch(e => { process.stdout.write('THREW:' + ((e && e.message) || e)); process.exit(3); });`;

    let r: ReturnType<typeof child_process.spawnSync>;
    try {
        r = child_process.spawnSync(process.execPath, ['--input-type=module', '-e', script], { encoding: 'utf8', timeout: 30000 });
    } catch (e) {
        return [fail(`${label}: smoke-imports as an ESM module`, String(e))];
    }
    const out = String(r.stdout ?? '').trim();
    if (r.status === 0 && out === 'OK') {
        return [pass(`${label}: smoke-imports and exposes parseOffice + OfficeGenerator.generate`)];
    }
    return [fail(`${label}: smoke-imports and exposes parseOffice + OfficeGenerator.generate`,
        out || String(r.stderr ?? '').slice(0, 200) || `exit ${r.status}`)];
}

// ---------------------------------------------------------------------------
// Section 6: Browser Type Declarations
// ---------------------------------------------------------------------------

function checkBrowserTypes(isSlim = false): CheckResult[] {
    const results: CheckResult[] = [];
    const suffix = isSlim ? '.slim' : '';
    const relPath = `dist/officeparser.browser${suffix}.d.ts`;
    const label = isSlim ? 'Browser types Slim' : 'Browser types';

    if (!fileExists(relPath)) {
        return [fail(`${label}: ${relPath} exists`, 'File not found')];
    }
    results.push(pass(`${label}: exists`));

    const content = readFile(relPath);

    if (content.includes('OfficeParser')) {
        results.push(pass(`${label}: contains OfficeParser declaration`));
    } else {
        results.push(fail(`${label}: contains OfficeParser declaration`, 'OfficeParser not found in .d.ts'));
    }

    if (content.includes('parseOffice')) {
        results.push(pass(`${label}: contains parseOffice declaration`));
    } else {
        results.push(fail(`${label}: contains parseOffice declaration`, 'parseOffice not found in .d.ts'));
    }

    if (content.includes('OfficeGenerator')) {
        results.push(pass(`${label}: contains OfficeGenerator declaration`));
    } else {
        results.push(fail(`${label}: contains OfficeGenerator declaration`, 'OfficeGenerator not found in .d.ts'));
    }

    if (content.includes('GeneratorConfig')) {
        results.push(pass(`${label}: contains GeneratorConfig declaration`));
    } else {
        results.push(fail(`${label}: contains GeneratorConfig declaration`, 'GeneratorConfig not found in .d.ts'));
    }

    // OfficeError is the documented way to read a thrown error's typed code
    // (`err.officeIssue.code`). Being an interface it has no runtime export to check, so the
    // declaration bundle is the only place a TypeScript consumer can lose access to it.
    if (content.includes('OfficeError')) {
        results.push(pass(`${label}: contains OfficeError declaration`));
    } else {
        results.push(fail(`${label}: contains OfficeError declaration`, 'OfficeError not found in .d.ts'));
    }

    return results;
}

// ---------------------------------------------------------------------------
// Section 7: package.json Paths Validation
// ---------------------------------------------------------------------------

function checkPackageJson(): CheckResult[] {
    const results: CheckResult[] = [];
    const pkgPath = path.join(ROOT, 'package.json');

    let pkg: any;
    try {
        pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
        results.push(pass('package.json: parseable'));
    } catch (e) {
        return [fail('package.json: parseable', String(e))];
    }

    const checkField = (label: string, relPath: string | undefined) => {
        if (!relPath) {
            results.push(fail(`package.json: ${label}`, 'Field is missing'));
            return;
        }
        // Normalise leading ./
        const normalised = relPath.replace(/^\.\//, '');
        if (fileExists(normalised)) {
            results.push(pass(`package.json: ${label}`, relPath));
        } else {
            results.push(fail(`package.json: ${label}`, `${relPath} → file not found`));
        }
    };

    checkField('"main"', pkg.main);
    checkField('"module"', pkg.module);
    checkField('"types"', pkg.types);
    checkField('"browser"', pkg.browser);
    checkField('"bin.officeparser"', pkg.bin?.officeparser);

    // Exports map
    const exp = pkg.exports?.['.'];
    if (!exp) {
        results.push(fail('package.json: exports["."]', 'Missing exports map'));
    } else {
        checkField('exports["."].types', exp.types);
        checkField('exports["."].browser', exp.browser);
        checkField('exports["."].import', exp.import);
        checkField('exports["."].require', exp.require);
    }

    const slimExp = pkg.exports?.['./slim'];
    if (!slimExp) {
        results.push(fail('package.json: exports["./slim"]', 'Missing exports map for slim'));
    } else {
        checkField('exports["./slim"].types', slimExp.types);
        checkField('exports["./slim"].browser', slimExp.browser);
        checkField('exports["./slim"].import', slimExp.import);
    }

    return results;
}

// ---------------------------------------------------------------------------
// Runner & Reporter
// ---------------------------------------------------------------------------

function printSection(title: string, results: CheckResult[]): { passed: number; failed: number } {
    console.log(`\n${'─'.repeat(70)}`);
    console.log(`  ${title}`);
    console.log('─'.repeat(70));

    let passed = 0;
    let failed = 0;

    for (const r of results) {
        const icon = r.status === 'PASS' ? '✅' : r.status === 'FAIL' ? '❌' : '⏭';
        const detail = r.detail ? `  (${r.detail})` : '';
        console.log(`  ${icon} ${r.name}${detail}`);
        if (r.status === 'PASS') passed++;
        if (r.status === 'FAIL') failed++;
    }

    console.log(`\n  Passed: ${passed}/${results.length}${failed > 0 ? `   Failed: ${failed}` : ''}`);
    return { passed, failed };
}

async function main() {
    console.log('═'.repeat(70));
    console.log('  SHIPPING ARTIFACT VALIDATION');
    console.log('═'.repeat(70));

    const sections: Array<{ title: string; fn: () => CheckResult[] }> = [
        { title: 'Node.js CJS Package', fn: checkCjs },
        { title: 'Node.js ESM Package', fn: checkEsm },
        { title: 'CLI Entry (dist/cli.js)', fn: checkCli },
        { title: 'Browser IIFE Bundle (dist/officeparser.browser.iife.js)', fn: () => checkBrowserIife(false) },
        { title: 'Browser IIFE Slim Bundle (dist/officeparser.browser.slim.iife.js)', fn: () => checkBrowserIife(true) },
        { title: 'Browser ESM Bundle (dist/officeparser.browser.mjs)', fn: () => checkBrowserEsm(false) },
        { title: 'Browser ESM Slim Bundle (dist/officeparser.browser.slim.mjs)', fn: () => checkBrowserEsm(true) },
        { title: 'Browser Type Declarations', fn: () => checkBrowserTypes(false) },
        { title: 'Browser Slim Type Declarations', fn: () => checkBrowserTypes(true) },
        { title: 'Dynamic import classification', fn: () => checkDynamicImportClassifier() },
        {
            title: 'Browser bundles are safe for consumer bundlers', fn: () => [
                ...checkBundlerCompatibility('dist/officeparser.browser.mjs', 'Browser ESM'),
                ...checkBundlerCompatibility('dist/officeparser.browser.slim.mjs', 'Browser ESM Slim'),
                ...checkBundlerCompatibility('dist/officeparser.browser.iife.js', 'Browser IIFE'),
                ...checkBundlerCompatibility('dist/officeparser.browser.slim.iife.js', 'Browser IIFE Slim'),
            ]
        },
        {
            title: 'Buffer type detection is bundled for the browser', fn: () => [
                ...checkFileTypeInlined('dist/officeparser.browser.mjs', 'Browser ESM'),
                ...checkFileTypeInlined('dist/officeparser.browser.slim.mjs', 'Browser ESM Slim'),
                ...checkFileTypeInlined('dist/officeparser.browser.iife.js', 'Browser IIFE'),
                ...checkFileTypeInlined('dist/officeparser.browser.slim.iife.js', 'Browser IIFE Slim'),
            ]
        },
        {
            title: 'Browser ESM bundles execute (smoke import)', fn: () => [
                ...checkBrowserEsmExecutes(false),
                ...checkBrowserEsmExecutes(true),
            ]
        },
        { title: 'package.json Path Integrity', fn: checkPackageJson },
    ];

    let totalPassed = 0;
    let totalFailed = 0;

    for (const { title, fn } of sections) {
        const results = fn();
        const { passed, failed } = printSection(title, results);
        totalPassed += passed;
        totalFailed += failed;
    }

    console.log(`\n${'═'.repeat(70)}`);
    console.log(`  SUMMARY: ${totalPassed} passed, ${totalFailed} failed`);
    console.log('═'.repeat(70));

    if (totalFailed > 0) {
        console.log('\n❌ Artifact validation FAILED — fix the issues above before publishing.\n');
        process.exit(1);
    } else {
        console.log('\n✅ All shipping artifacts are valid.\n');
    }
}

main().catch(err => {
    console.error('Artifact test runner error:', err);
    process.exit(1);
});

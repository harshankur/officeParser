/**
 * Browser Bundle Builder
 *
 * Produces two browser-targeted bundles from src/index.ts:
 *
 *  1. dist/officeparser.browser.mjs  — ESM, for Vite/webpack/bundlers
 *  2. dist/officeparser.browser.iife.js — IIFE + UMD footer, for <script> tags
 *
 * Both bundles:
 *  - Are fully self-contained (all deps bundled in)
 *  - Polyfill Node.js built-ins for browser compatibility
 *  - Inject a `/* @vite-ignore *\/` comment before `import(this.workerSrc)`
 *    in the bundled pdfjs-dist code to suppress Vite's unanalyzable dynamic
 *    import warning.
 *  - Evaluate no code from a string as they load, so a page with a strict
 *    Content Security Policy can load them (see scripts/loadTimeEval.js).
 */

const esbuild = require('esbuild');
const { nodeModulesPolyfillPlugin } = require('esbuild-plugins-node-modules-polyfill');
const fs = require('fs');
const path = require('path');
const { annotateDynamicImports } = require('./scripts/dynamicImports.js');
const { findLoadTimeEvals, replaceLoadTimeEvals } = require('./scripts/loadTimeEval.js');

// ---------------------------------------------------------------------------
// Config Generator
// ---------------------------------------------------------------------------

function getBrowserConfig(isSlim, externalPdfLib = false) {
    const config = {
        entryPoints: ['src/index.ts'],
        bundle: true,
        platform: 'browser',
        target: ['es2020'],
        sourcemap: false,
        minify: true,
        // Remap node built-ins used in our own source to empty stubs.
        // The browser bundle is Buffer-in/Buffer-out; nobody reads from disk.
        alias: {
            'fs': path.resolve(__dirname, 'scripts/browser-stubs/fs.js'),
            'fs/promises': path.resolve(__dirname, 'scripts/browser-stubs/fs.js'),
            'puppeteer': path.resolve(__dirname, 'scripts/browser-stubs/puppeteer.js'),
            // pdf-lib (native PDF engine) is an optional peer dep: keep it out of the prebuilt
            // bundle. A self-bundling consumer with pdf-lib installed resolves the real package.
            'pdf-lib': path.resolve(__dirname, 'scripts/browser-stubs/pdf-lib.js'),
            // Left unresolved, this reaches the output as a bare Node built-in and a consumer's
            // bundler reports it as missing, even though the only code path that imports it is
            // gated on running under Node.
            'child_process': path.resolve(__dirname, 'scripts/browser-stubs/child_process.js'),
            'url': path.resolve(__dirname, 'scripts/browser-stubs/url.js'),
        },
        define: {
            'process.env.NODE_ENV': '"production"',
            'global': 'window',
            'import.meta.url': '""',
            '__SLIM__': isSlim ? 'true' : 'false',
        },
        inject: ['./scripts/browser-shims.js'],
        banner: {
            js: `
// officeparser browser bundle
// Shim for setImmediate (not available in all browsers)
if (typeof setImmediate === 'undefined') {
  window.setImmediate = function(callback) { return setTimeout(callback, 0); };
}
`.trim(),
        },
        plugins: [
            // Polyfill Node.js built-ins for the browser.
            // fs: 'empty' — browser bundle never reads from disk (callers pass Buffer directly)
            // Other modules are polyfilled with working browser equivalents.
            nodeModulesPolyfillPlugin({
                modules: {
                    zlib: true,
                    crypto: true,
                    stream: true,
                    buffer: true,
                    util: true,
                    events: true,
                    timers: true,
                    path: true,
                    os: true,
                    assert: true,
                    url: true,
                    vm: true,
                    http: true,
                    https: true,
                    string_decoder: true,
                },
            }),
            // Post-process: inject /* @vite-ignore */ into dynamic imports with variables
            // to suppress Vite's unanalyzable dynamic import warning.
            viteIgnoreDynamicImportsPlugin(),
            // Post-process: what a polyfill evaluates from a string at load is written out in its
            // place, so the bundle loads under a Content Security Policy without 'unsafe-eval'.
            loadTimeEvalPlugin(),
        ],
    };

    if (isSlim) {
        config.alias['tesseract.js'] = path.resolve(__dirname, 'scripts/browser-stubs/tesseract.js');
    }

    // Dedicated native-PDF entry: leave pdf-lib EXTERNAL instead of aliasing it to the throwing stub,
    // so a self-bundling consumer that installs pdf-lib gets the real native PDF engine in the browser
    // (the dynamic `import('pdf-lib')` survives for their bundler to resolve). The default bundles keep
    // the stub so they stay self-contained and never force pdf-lib on a consumer that does not want it.
    if (externalPdfLib) {
        delete config.alias['pdf-lib'];
        config.external = ['pdf-lib'];
    }

    return config;
}

// ---------------------------------------------------------------------------
// Bundler ignore annotations for dynamic imports (see scripts/dynamicImports.js)
// ---------------------------------------------------------------------------

function viteIgnoreDynamicImportsPlugin() {
    return {
        name: 'vite-ignore-dynamic-imports',
        setup(build) {
            build.onEnd(result => {
                if (result.errors.length > 0) return;

                const outfile = build.initialOptions.outfile;
                if (!outfile || !fs.existsSync(outfile)) return;

                const content = fs.readFileSync(outfile, 'utf8');
                const { output, annotated } = annotateDynamicImports(content);

                if (annotated > 0) {
                    // This edits an already-built bundle by offset, so confirm the result still
                    // parses before it is written. A miscounted offset would otherwise ship a
                    // syntactically broken bundle that only fails in a consumer's build.
                    try {
                        esbuild.transformSync(output, { loader: 'js', format: 'esm' });
                    } catch (err) {
                        throw new Error(`Annotating dynamic imports broke ${path.basename(outfile)}: ${err.message}`);
                    }

                    fs.writeFileSync(outfile, output, 'utf8');
                    console.log(`  → annotated ${annotated} dynamic import(s) in ${path.basename(outfile)}`);
                }
            });
        },
    };
}

// ---------------------------------------------------------------------------
// Nothing evaluated from a string at load (see scripts/loadTimeEval.js)
// ---------------------------------------------------------------------------

function loadTimeEvalPlugin() {
    return {
        name: 'load-time-eval',
        setup(build) {
            build.onEnd(result => {
                if (result.errors.length > 0) return;

                const outfile = build.initialOptions.outfile;
                if (!outfile || !fs.existsSync(outfile)) return;

                const { output, replaced } = replaceLoadTimeEvals(fs.readFileSync(outfile, 'utf8'));

                // A probe written another way (a dependency changed) is not one the replacement
                // knows: fail here rather than ship a bundle that breaks a strict policy again.
                const left = findLoadTimeEvals(output);
                if (left.length > 0) {
                    throw new Error(`${path.basename(outfile)} still evaluates code at load: ${left[0].snippet} (add it to scripts/loadTimeEval.js)`);
                }

                if (replaced > 0) {
                    // As for the dynamic imports: an edited bundle is parsed before it is written.
                    try {
                        esbuild.transformSync(output, { loader: 'js', format: 'esm' });
                    } catch (err) {
                        throw new Error(`Replacing load-time evaluation broke ${path.basename(outfile)}: ${err.message}`);
                    }

                    fs.writeFileSync(outfile, output, 'utf8');
                    console.log(`  → replaced ${replaced} load-time evaluation(s) in ${path.basename(outfile)}`);
                }
            });
        },
    };
}

// ---------------------------------------------------------------------------
// Build 1: ESM bundle (for Vite, webpack, Angular, etc.)
// ---------------------------------------------------------------------------

async function buildEsm(isSlim = false) {
    const suffix = isSlim ? '.slim' : '';
    console.log(`Building ESM browser bundle → dist/officeparser.browser${suffix}.mjs`);
    await esbuild.build({
        ...getBrowserConfig(isSlim),
        outfile: `dist/officeparser.browser${suffix}.mjs`,
        format: 'esm',
    });
    console.log(`  ✓ dist/officeparser.browser${suffix}.mjs`);
}

// ---------------------------------------------------------------------------
// Build 2: IIFE bundle (for <script> tags and backward compat)
// ---------------------------------------------------------------------------

async function buildIife(isSlim = false) {
    const suffix = isSlim ? '.slim' : '';
    console.log(`Building IIFE browser bundle → dist/officeparser.browser${suffix}.iife.js`);
    await esbuild.build({
        ...getBrowserConfig(isSlim),
        outfile: `dist/officeparser.browser${suffix}.iife.js`,
        format: 'iife',
        globalName: 'officeParser',
        // UMD-style footer: set module.exports so Vite's __commonJS wrapper
        // picks up the IIFE result when consumers still import the IIFE file.
        footer: {
            js: 'if(typeof module!=="undefined")module.exports=officeParser;',
        },
    });
    console.log(`  ✓ dist/officeparser.browser${suffix}.iife.js`);
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Build 3: ESM bundle with pdf-lib left external (client-side native PDF export)
// ---------------------------------------------------------------------------

async function buildNativePdf() {
    console.log('Building ESM browser bundle (pdf-lib external) → dist/officeparser.browser.native-pdf.mjs');
    await esbuild.build({
        ...getBrowserConfig(false, true),
        outfile: 'dist/officeparser.browser.native-pdf.mjs',
        format: 'esm',
    });
    console.log('  ✓ dist/officeparser.browser.native-pdf.mjs');
}

async function main() {
    try {
        await buildEsm(false);
        await buildIife(false);
        await buildEsm(true);
        await buildIife(true);
        await buildNativePdf();
        console.log('\nBrowser bundles built successfully.');
    } catch (err) {
        console.error('Build failed:', err);
        process.exit(1);
    }
}

main();

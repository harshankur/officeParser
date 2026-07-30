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
 */

const esbuild = require('esbuild');
const { nodeModulesPolyfillPlugin } = require('esbuild-plugins-node-modules-polyfill');
const fs = require('fs');
const path = require('path');
const { annotateDynamicImports } = require('./scripts/dynamicImports.js');

// ---------------------------------------------------------------------------
// Config Generator
// ---------------------------------------------------------------------------

function getBrowserConfig(isSlim) {
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
        ],
    };

    if (isSlim) {
        config.alias['tesseract.js'] = path.resolve(__dirname, 'scripts/browser-stubs/tesseract.js');
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

async function main() {
    try {
        await buildEsm(false);
        await buildIife(false);
        await buildEsm(true);
        await buildIife(true);
        console.log('\nBrowser bundles built successfully.');
    } catch (err) {
        console.error('Build failed:', err);
        process.exit(1);
    }
}

main();

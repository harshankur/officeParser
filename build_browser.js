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

// ---------------------------------------------------------------------------
// Shared configuration
// ---------------------------------------------------------------------------

const sharedConfig = {
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
    },
    define: {
        'process.env.NODE_ENV': '"production"',
        'global': 'window',
        'import.meta.url': '""',
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
        // Post-process: inject /* @vite-ignore */ before `import(this.workerSrc)`
        // inside bundled pdfjs-dist to suppress Vite's unanalyzable dynamic import warning.
        viteIgnorePdfjsWorkerPlugin(),
    ],
};

// ---------------------------------------------------------------------------
// @vite-ignore plugin for pdfjs-dist workerSrc dynamic import
// ---------------------------------------------------------------------------

function viteIgnorePdfjsWorkerPlugin() {
    return {
        name: 'vite-ignore-pdfjs-worker',
        setup(build) {
            build.onEnd(result => {
                if (result.errors.length > 0) return;

                const outfile = build.initialOptions.outfile;
                if (!outfile || !fs.existsSync(outfile)) return;

                let content = fs.readFileSync(outfile, 'utf8');

                // Replace `import(this.workerSrc)` with the Vite-ignore annotated form.
                // We match both `import(this.workerSrc)` and its minified variants.
                const pattern = /\bimport\((this\.workerSrc)\)/g;
                const replaced = content.replace(
                    pattern,
                    'import(/* @vite-ignore */ $1)'
                );

                if (replaced !== content) {
                    fs.writeFileSync(outfile, replaced, 'utf8');
                    console.log(`  → injected @vite-ignore into ${path.basename(outfile)}`);
                } else {
                    console.error(`\nERROR: Could not find \`import(this.workerSrc)\` in ${path.basename(outfile)} to inject @vite-ignore. Vite bundle will contain unanalyzable dynamic imports!\n`);
                    process.exit(1);
                }
            });
        },
    };
}

// ---------------------------------------------------------------------------
// Build 1: ESM bundle (for Vite, webpack, Angular, etc.)
// ---------------------------------------------------------------------------

async function buildEsm() {
    console.log('Building ESM browser bundle → dist/officeparser.browser.mjs');
    await esbuild.build({
        ...sharedConfig,
        outfile: 'dist/officeparser.browser.mjs',
        format: 'esm',
    });
    console.log('  ✓ dist/officeparser.browser.mjs');
}

// ---------------------------------------------------------------------------
// Build 2: IIFE bundle (for <script> tags and backward compat)
// ---------------------------------------------------------------------------

async function buildIife() {
    console.log('Building IIFE browser bundle → dist/officeparser.browser.iife.js');
    await esbuild.build({
        ...sharedConfig,
        outfile: 'dist/officeparser.browser.iife.js',
        format: 'iife',
        globalName: 'officeParser',
        // UMD-style footer: set module.exports so Vite's __commonJS wrapper
        // picks up the IIFE result when consumers still import the IIFE file.
        footer: {
            js: 'if(typeof module!=="undefined")module.exports=officeParser;',
        },
    });
    console.log('  ✓ dist/officeparser.browser.iife.js');
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

async function main() {
    try {
        await buildEsm();
        await buildIife();
        console.log('\nBrowser bundles built successfully.');
    } catch (err) {
        console.error('Build failed:', err);
        process.exit(1);
    }
}

main();

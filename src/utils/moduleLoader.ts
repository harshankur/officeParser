/**
 * Module Loader Utility
 * 
 * Centralizes dynamic imports for ESM-only packages (like file-type and pdfjs-dist)
 * to maintain CommonJS compatibility while supporting modern ESM dependencies.
 * 
 * This approach prevents TypeScript from transpiling dynamic import() 
 * into require() when targeting CommonJS.
 */

import { isBrowser, ensureEnvPolyfills } from './envUtils.js';
import type * as FileTypeModule from 'file-type' with { 'resolution-mode': 'import' };

async function loadNodeEsmModule<T>(specifier: string): Promise<T> {
    // In Node.js, we resolve the specifier to an absolute file URL.
    // This ensures that the dynamic import() call
    // always finds the correct module regardless of the caller's context.
    // This is especially important in Node 18 for sub-paths of packages.
    try {
        const { pathToFileURL } = await import('url');
        // @ts-ignore - require.resolve is available in Node.js CJS context
        const absolutePath = require.resolve(specifier);
        const fileUrl = pathToFileURL(absolutePath).href;
        return import(fileUrl);
    } catch (e) {
        // Fallback for cases where require.resolve might fail (e.g. non-file specifiers)
        return import(specifier);
    }
}

/**
 * Returns true if require.resolve is available in this runtime context.
 * It is NOT available in native ESM (e.g. the .mjs wrapper) or browser environments.
 * Checking this guards against a ReferenceError that would silently trigger
 * the bundled fallback path for non-bundled ESM consumers.
 */
function isRequireAvailable(): boolean {
    // @ts-ignore - require may not exist in ESM context
    return typeof require !== 'undefined' && typeof require.resolve === 'function';
}

/** 
 * Specialized loader for file-type 
 */
export async function loadFileType(): Promise<typeof FileTypeModule> {
    if (!isBrowser) {
        // Ensure environment polyfills for Node.js 18 support
        ensureEnvPolyfills();

        if (isRequireAvailable()) {
            // Check if node_modules is available at runtime.
            // @ts-ignore - require.resolve is available in Node.js CJS context
            require.resolve(String('file-type'));
            // node_modules present: use the path-resolved loader for Node 18 ESM compatibility
            return loadNodeEsmModule<typeof FileTypeModule>('file-type');
        }
    }
    // Covers three cases with one return:
    //   1. Node.js standalone/bundled (SEA or bundled CJS): bundler inlines the module at build time.
    //   2. Node.js native ESM (no require available): runtime resolves the bare specifier.
    //   3. Browser: bundler (esbuild/Vite) handles the static-looking dynamic import().
    // Note: bypasses the Node 18 sub-path ESM fix in loadNodeEsmModule, but bundlers handle it.
    return import('file-type');
}

/** 
 * Specialized loader for pdfjs-dist 
 */
export async function loadPdfJs(): Promise<any> {
    if (!isBrowser) {
        // Ensure environment polyfills for Node.js 18 support
        ensureEnvPolyfills();

        if (isRequireAvailable()) {
            // @ts-ignore - require.resolve is available in Node.js CJS context
            const pkgExists = (() => { try { require.resolve(String('pdfjs-dist')); return true; } catch { return false; } })();

            if (pkgExists) {
                // node_modules present: try the legacy build path first for stability with
                // ESM-only main, then fall back to the package root.
                // Errors from these loaders are NOT swallowed into the bundled-fallback path.
                try {
                    return await loadNodeEsmModule('pdfjs-dist/legacy/build/pdf.mjs');
                } catch {
                    return await loadNodeEsmModule('pdfjs-dist');
                }
            }
        }
        // Standalone/bundled fallback for Node.js (SEA, bundled CJS without node_modules, or native ESM context):
        // Load both PDF.js and its worker, and register the worker on globalThis to enable the offline fake-worker.
        // @ts-ignore - Ignore type check for local .mjs files in node_modules/packaging
        const [pdfjs, pdfjsWorker] = await Promise.all([
            // @ts-ignore - mjs imports may not have types
            import('pdfjs-dist/legacy/build/pdf.mjs'),
            // @ts-ignore - mjs imports may not have types
            import('pdfjs-dist/legacy/build/pdf.worker.mjs')
        ]);
        (globalThis as any).pdfjsWorker = pdfjsWorker;
        return pdfjs;
    }
    // Browser environment: bundler (esbuild/Vite) handles the standard dynamic import()
    return import('pdfjs-dist');
}

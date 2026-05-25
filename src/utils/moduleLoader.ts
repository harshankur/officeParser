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
        // @ts-ignore - require.resolve is available in Node.js
        const absolutePath = require.resolve(specifier);
        const fileUrl = pathToFileURL(absolutePath).href;
        return import(fileUrl);
    } catch (e) {
        // Fallback for cases where require.resolve might fail (e.g. non-file specifiers)
        return import(specifier);
    }
}

/** 
 * Specialized loader for file-type 
 */
export async function loadFileType(): Promise<typeof FileTypeModule> {
    if (!isBrowser) {
        // Ensure environment polyfills for Node.js 18 support
        ensureEnvPolyfills();
        // Node.js path: Use dynamic import wrapper for CJS compatibility
        return loadNodeEsmModule<typeof FileTypeModule>('file-type');
    }
    // Browser path: standard dynamic import() is handled by bundlers (e.g. esbuild/Vite)
    return import('file-type');
}

/** 
 * Specialized loader for pdfjs-dist 
 */
export async function loadPdfJs(): Promise<any> {
    if (!isBrowser) {
        // Ensure environment polyfills for Node.js 18 support
        ensureEnvPolyfills();
        // Node.js environment: require legacy build for stability with ESM-only main
        try {
            return await loadNodeEsmModule('pdfjs-dist/legacy/build/pdf.mjs');
        } catch {
            return await loadNodeEsmModule('pdfjs-dist');
        }
    }
    // Browser environment: esbuild handles standard static-looking dynamic import()
    return import('pdfjs-dist');
}

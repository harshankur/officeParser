/**
 * Module Loader Utility
 * 
 * Centralizes dynamic imports for ESM-only packages (like file-type and pdfjs-dist)
 * to maintain CommonJS compatibility while supporting modern ESM dependencies.
 * 
 * This approach prevents TypeScript from transpiling dynamic import() 
 * into require() when targeting CommonJS.
 */

import { isBrowser, ensureDomMatrix } from './envUtils.js';
import type * as FileTypeModule from 'file-type' with { 'resolution-mode': 'import' };

/**
 * Dynamically loads an ESM module in a Node.js CJS context.
 * 
 * @param specifier - The module specifier to load
 * @returns The loaded module
 */
async function loadNodeEsmModule<T>(specifier: string): Promise<T> {
    // We use 'new Function' to bypass static analysis of tsc and some bundlers.
    // This ensures that Node.js sees a real 'import()' at runtime, even in a CJS file.
    return new Function('s', 'return import(s)')(specifier);
}

/** 
 * Specialized loader for file-type 
 */
export async function loadFileType(): Promise<typeof FileTypeModule> {
    if (!isBrowser) {
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
        // Ensure DOMMatrix polyfill for Node.js 18 support
        ensureDomMatrix();
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

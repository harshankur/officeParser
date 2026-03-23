/**
 * Module Loader Utility
 * 
 * Centralizes dynamic imports for ESM-only packages (like file-type and pdfjs-dist)
 * to maintain CommonJS compatibility while supporting modern ESM dependencies.
 * 
 * This approach prevents TypeScript from transpiling dynamic import() 
 * into require() when targeting CommonJS.
 */

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
export async function loadFileType(): Promise<typeof import('file-type')> {
    if (typeof window === 'undefined') {
        // Node.js path
        return loadNodeEsmModule<typeof import('file-type')>('file-type');
    }
    // Browser path: esbuild handles standard dynamic import()
    return import('file-type');
}

/** 
 * Specialized loader for pdfjs-dist 
 */
export async function loadPdfJs(): Promise<any> {
    if (typeof window === 'undefined') {
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

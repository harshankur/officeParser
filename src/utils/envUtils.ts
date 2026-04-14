/**
 * Environment detection and safe utility wrappers.
 */

/**
 * Detect if we are running in a browser environment.
 */
export const isBrowser = typeof window !== 'undefined' && typeof window.document !== 'undefined';

/**
 * Supported Node.js-only features that require explicit guarding for browser compatibility.
 */
export type NodeFeature = 'fs' | 'path-parsing' | 'pdf-worker-auto-resolution';

/**
 * Human-readable descriptions for Node-only features.
 */
const readableFeatures: Record<NodeFeature, string> = {
    'fs': 'direct file system access',
    'path-parsing': 'parsing from file path string',
    'pdf-worker-auto-resolution': 'automatic PDF worker resolution from node_modules'
};

/**
 * Throws an error if attempted to use Node.js-specific features in the browser.
 * 
 * @param feature - The Node.js feature being accessed
 * @throws {Error} Clear error message directing browser users to use Buffers
 */
export function assertNode(feature: NodeFeature): void {
    if (isBrowser) {
        throw new Error(`officeparser: '${readableFeatures[feature]}' is not supported in the browser. Browser users must pass file content as Buffer or ArrayBuffer directly.`);
    }
}

/**
 * Polyfills DOMMatrix if not available globally (required for Node.js < 20).
 * This shim provides enough properties for pdfjs-dist 5.x to calculate
 * text coordinates and transformations.
 */
export function ensureDomMatrix(): void {
    if (typeof global !== 'undefined' && !(global as any).DOMMatrix) {
        (global as any).DOMMatrix = class DOMMatrix {
            a: number; b: number; c: number; d: number; e: number; f: number;

            constructor(init?: number[] | string) {
                if (Array.isArray(init) && init.length >= 6) {
                    this.a = init[0]; this.b = init[1];
                    this.c = init[2]; this.d = init[3];
                    this.e = init[4]; this.f = init[5];
                } else {
                    this.a = this.d = 1;
                    this.b = this.c = this.e = this.f = 0;
                }
            }

            // Map standard matrix properties for compatibility
            get m11() { return this.a; }
            get m12() { return this.b; }
            get m21() { return this.c; }
            get m22() { return this.d; }
            get m41() { return this.e; }
            get m42() { return this.f; }
        };
    }
}
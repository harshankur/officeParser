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
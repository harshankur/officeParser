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

import { OfficeErrorType } from '../types.js';
import { getOfficeError } from './errorUtils.js';

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
        throw getOfficeError(OfficeErrorType.FEATURE_NOT_SUPPORTED_IN_BROWSER, undefined, readableFeatures[feature]);
    }
}

/**
 * Polyfills environment features not available globally (required for Node.js < 20).
 * This shim provides enough properties for pdfjs-dist 5.x to calculate
 * text coordinates and transformations, and guards against modern RegExp flags.
 */
export function ensureEnvPolyfills(): void {
    if (typeof global !== 'undefined') {
        // 1. DOMMatrix Polyfill (Node.js < 20)
        if (!(global as any).DOMMatrix) {
            (global as any).DOMMatrix = class DOMMatrix {
                a: number; b: number; c: number; d: number; e: number; f: number;

                constructor(init?: number[] | string | DOMMatrix) {
                    if (Array.isArray(init) && init.length >= 6) {
                        this.a = init[0]; this.b = init[1];
                        this.c = init[2]; this.d = init[3];
                        this.e = init[4]; this.f = init[5];
                    } else if (typeof init === 'object' && init !== null) {
                        this.a = (init as any).a; this.b = (init as any).b;
                        this.c = (init as any).c; this.d = (init as any).d;
                        this.e = (init as any).e; this.f = (init as any).f;
                    } else {
                        this.a = this.d = 1;
                        this.b = this.c = this.e = this.f = 0;
                    }
                }

                // Standard matrix property aliases for compatibility
                get m11() { return this.a; }
                get m12() { return this.b; }
                get m21() { return this.c; }
                get m22() { return this.d; }
                get m41() { return this.e; }
                get m42() { return this.f; }

                multiply(other: any) {
                    return new DOMMatrix([
                        this.a * other.a + this.c * other.b,
                        this.b * other.a + this.d * other.b,
                        this.a * other.c + this.c * other.d,
                        this.b * other.c + this.d * other.d,
                        this.a * other.e + this.c * other.f + this.e,
                        this.b * other.e + this.d * other.f + this.f
                    ]);
                }

                inverse() {
                    const det = this.a * this.d - this.b * this.c;
                    if (det === 0) return new DOMMatrix();
                    return new DOMMatrix([
                        this.d / det,
                        -this.b / det,
                        -this.c / det,
                        this.a / det,
                        (this.c * this.f - this.d * this.e) / det,
                        (this.b * this.e - this.a * this.f) / det
                    ]);
                }

                transformPoint(point?: { x: number, y: number }) {
                    const x = point?.x ?? 0;
                    const y = point?.y ?? 0;
                    return {
                        x: x * this.a + y * this.c + this.e,
                        y: x * this.b + y * this.d + this.f
                    };
                }
            };
        }

        // 2. ImageData Polyfill (Node.js < 20)
        if (!(global as any).ImageData) {
            (global as any).ImageData = class ImageData {
                width: number;
                height: number;
                data: Uint8ClampedArray;
                constructor(data: Uint8ClampedArray, width: number, height: number) {
                    this.data = data;
                    this.width = width;
                    this.height = height;
                }
            };
        }

        // 3. RegExp 'v' flag Polyfill (Node.js < 20)
        // Modern dependencies (file-type 22+, pdfjs-dist 5+) might use the 'v' flag.
        // We wrap the constructor to downgrade 'v' to 'u' in Node 18.
        const OriginalRegExp = global.RegExp;
        try {
            new OriginalRegExp('', 'v');
        } catch (e) {
            // 'v' flag is not supported (Node 18), apply fallback wrapper
            const RegExpWrapper = function (pattern: string | RegExp, flags?: string) {
                if (typeof flags === 'string' && flags.includes('v')) {
                    // Fallback 'v' to 'u' (Unicode sets to Unicode)
                    return new OriginalRegExp(pattern, flags.replace('v', 'u'));
                }
                return new OriginalRegExp(pattern, flags);
            };

            // Maintain prototype chain and static methods
            RegExpWrapper.prototype = OriginalRegExp.prototype;
            Object.setPrototypeOf(RegExpWrapper, OriginalRegExp);
            
            (global as any).RegExp = RegExpWrapper;
        }
    }
}

/**
 * Backward-compatible alias for ensureEnvPolyfills.
 * @deprecated Use ensureEnvPolyfills instead.
 */
export const ensureDomMatrix = ensureEnvPolyfills;
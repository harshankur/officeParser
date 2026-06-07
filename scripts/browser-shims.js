/**
 * Browser-side shims for Node.js built-ins.
 */

import { Buffer } from 'buffer';
import process from 'process';

// Robust global assignment
const globals = [
    typeof globalThis !== 'undefined' ? globalThis : null,
    typeof window !== 'undefined' ? window : null,
    typeof self !== 'undefined' ? self : null,
    typeof global !== 'undefined' ? global : null,
].filter(Boolean);

globals.forEach(g => {
    if (!g.Buffer) g.Buffer = Buffer;
    if (!g.process) g.process = process;
});

if (typeof process !== 'undefined' && !process.env) {
    process.env = { NODE_ENV: 'production' };
}

// Polyfill Promise.try for environments/browsers lacking native ES2024 Promise.try support (e.g. older Puppeteer/Chrome)
if (typeof Promise !== 'undefined' && !Promise.try) {
    Promise.try = function(fn, ...args) {
        return new Promise((resolve, reject) => {
            try {
                resolve(fn(...args));
            } catch (err) {
                reject(err);
            }
        });
    };
}

// Polyfill Map getOrInsert / getOrInsertComputed for compatibility with modern PDF.js (v5.x)
if (typeof Map !== 'undefined') {
    if (!Map.prototype.getOrInsertComputed) {
        Map.prototype.getOrInsertComputed = function(key, callback) {
            if (this.has(key)) {
                return this.get(key);
            }
            const value = callback(key);
            this.set(key, value);
            return value;
        };
    }
    if (!Map.prototype.getOrInsert) {
        Map.prototype.getOrInsert = function(key, value) {
            if (this.has(key)) {
                return this.get(key);
            }
            this.set(key, value);
            return value;
        };
    }
}

// Minimal util shim for debuglog if dependencies like @jspm/core leak it
if (typeof globals[0] !== 'undefined') {
    const g = globals[0];
    if (!g.util) {
        g.util = {
            debuglog: () => () => {},
            inspect: (obj) => JSON.stringify(obj),
        };
    }
}

export { Buffer, process };

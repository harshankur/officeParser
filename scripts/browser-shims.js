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

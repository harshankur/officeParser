/**
 * ESM Import Test Helper
 *
 * This script is spawned as a subprocess by testShippingArtifacts.ts to validate
 * that the ESM wrapper (dist/index.mjs) can be imported and provides the expected
 * named exports. Results are printed to stdout as JSON.
 *
 * Usage (internal): spawned via `node test/esm-test.mjs`
 *
 * Exit codes:
 *   0 = all checks passed
 *   1 = one or more checks failed
 */

import { createRequire } from 'module';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import { existsSync } from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const ROOT = join(__dirname, '..');

const results = [];
let failed = false;

function pass(name, detail = '') {
    results.push({ name, status: 'PASS', detail });
}

function fail(name, detail = '') {
    results.push({ name, status: 'FAIL', detail });
    failed = true;
}

// --- Check 1: dist/index.mjs exists ---
const mjsPath = join(ROOT, 'dist', 'index.mjs');
if (!existsSync(mjsPath)) {
    fail('ESM wrapper exists', `dist/index.mjs not found at ${mjsPath}`);
    console.log(JSON.stringify(results));
    process.exit(1);
} else {
    pass('ESM wrapper exists');
}

// --- Check 2: dynamic import resolves ---
let mod;
try {
    mod = await import(mjsPath);
    pass('ESM import resolves');
} catch (e) {
    fail('ESM import resolves', String(e));
    console.log(JSON.stringify(results));
    process.exit(1);
}

// --- Check 3: OfficeParser named export ---
if (typeof mod.OfficeParser === 'function' || typeof mod.OfficeParser === 'object') {
    pass('Named export: OfficeParser', typeof mod.OfficeParser);
} else {
    fail('Named export: OfficeParser', `Got: ${typeof mod.OfficeParser}`);
}

// --- Check 4: parseOffice named export ---
if (typeof mod.parseOffice === 'function') {
    pass('Named export: parseOffice', 'function');
} else {
    fail('Named export: parseOffice', `Got: ${typeof mod.parseOffice}`);
}

// --- Check 5: default export ---
const defaultExport = mod.default;
if (defaultExport && (typeof defaultExport === 'function' || typeof defaultExport === 'object')) {
    pass('Default export exists', typeof defaultExport);
} else {
    fail('Default export exists', `Got: ${typeof defaultExport}`);
}

// --- Check 6: parseOffice is a function on OfficeParser class ---
const OfficeParser = mod.OfficeParser ?? mod.default;
if (OfficeParser && typeof OfficeParser.parseOffice === 'function') {
    pass('OfficeParser.parseOffice is a function');
} else {
    fail('OfficeParser.parseOffice is a function', `Got: ${typeof OfficeParser?.parseOffice}`);
}

console.log(JSON.stringify(results));
process.exit(failed ? 1 : 0);

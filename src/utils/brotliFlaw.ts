/**
 * Node's Brotli decoder writes out of bounds on crafted input (CVE-2026-111497, "Node2Shell"): a process
 * decompressing a hostile stream can crash, leak memory or run the attacker's code. The flaw is in the
 * runtime, and only a fixed runtime removes it, but a document or a font must not be the way in. A PDF can
 * name `/Filter /BrotliDecode`, and pdf.js gives such a stream to the runtime's
 * `DecompressionStream('brotli')` where it has one; a WOFF2 font is Brotli too. On a runtime with the flaw
 * both go to a decoder written in JavaScript instead (pdf.js's own, fontkit's), and a warning says so.
 */

/** The first fixed release of each line that got a fix. The lines between them (23, 25), and those before 22, got none. */
const FIXED_NODE: Record<number, [minor: number, patch: number]> = { 22: [23, 4], 24: [22, 1], 26: [12, 1] };
/** The first release of the line after the last one above, all of whose releases carry the fix. */
const FIRST_LINE_ALWAYS_FIXED = 27;
const FIXED_BUN: [major: number, minor: number, patch: number] = [1, 4, 4];

/** What the warning tells a reader to upgrade to. */
export const BROTLI_FIXED_RELEASES = 'Node.js 22.23.4, 24.22.1, 26.12.1 or later';

/** For tests: says whether the runtime has the flaw, whatever its version. */
export const brotliFlawHooks: { flawed?: boolean } = {};

const numbers = (version: string): number[] => version.replace(/^v/, '').split('.').map(part => parseInt(part, 10));

/** Whether a version of Node has the flaw. */
export function nodeHasBrotliFlaw(version: string): boolean {
    const [major, minor, patch] = numbers(version);
    if (major >= FIRST_LINE_ALWAYS_FIXED) return false;
    const fixed = FIXED_NODE[major];
    if (!fixed) return true;
    return minor < fixed[0] || (minor === fixed[0] && patch < fixed[1]);
}

/** Whether this runtime's Brotli decoder has the flaw. Never in a browser, whose decoder is its own. */
export function runtimeHasBrotliFlaw(): boolean {
    if (brotliFlawHooks.flawed !== undefined) return brotliFlawHooks.flawed;
    const versions = (globalThis as any).process?.versions;
    if (typeof versions?.bun === 'string') {
        const [major, minor, patch] = numbers(versions.bun);
        return major < FIXED_BUN[0] || (major === FIXED_BUN[0] && (minor < FIXED_BUN[1] || (minor === FIXED_BUN[1] && patch < FIXED_BUN[2])));
    }
    if (typeof versions?.deno === 'string') return false;
    return typeof versions?.node === 'string' && nodeHasBrotliFlaw(versions.node);
}

/** The runtime as the warning names it. */
export function runtimeName(): string {
    const versions = (globalThis as any).process?.versions;
    if (typeof versions?.bun === 'string') return `Bun ${versions.bun}`;
    return typeof versions?.node === 'string' ? `Node.js ${versions.node}` : 'this runtime';
}

/** Whether pdf.js is what asks: the method of pdf.js's streams that makes a `DecompressionStream`, by name, in any of its builds. */
const askedByPdfjs = (): boolean => {
    const limit = Error.stackTraceLimit;
    Error.stackTraceLimit = 8;
    const stack = new Error().stack;
    Error.stackTraceLimit = limit;
    return String(stack).includes('asyncGetBytesFromDecompressionStream');
};

let refusedHere: (() => void) | undefined;
let installedHere = false;

/**
 * For pdf.js running in this process (a separate process has its own guard, in pdfProcess): the runtime's
 * Brotli stream is refused when pdf.js asks for it, and pdf.js then decodes the stream itself, as it does
 * wherever the stream fails. pdf.js reads the global, which is the host's too, so only pdf.js is refused:
 * the host's own `DecompressionStream`, of any format, works as it did.
 */
export function avoidBrotliInThisProcess(): void {
    const Native = (globalThis as any).DecompressionStream;
    if (installedHere || typeof Native !== 'function') return;
    installedHere = true;
    (globalThis as any).DecompressionStream = class DecompressionStream extends Native {
        constructor(format: unknown, ...rest: unknown[]) {
            if (format === 'brotli' && runtimeHasBrotliFlaw() && askedByPdfjs()) {
                refusedHere?.();
                throw new TypeError('The Brotli decoder of this runtime is not used (CVE-2026-111497)');
            }
            super(format, ...rest);
        }
    };
}

/** Who is told when pdf.js in this process was refused: the parse whose document is open (one at a time, see takeInProcessTurn). */
export function onBrotliRefused(sink: (() => void) | undefined): void {
    refusedHere = sink;
}

/**
 * Browser-side stub for Node.js 'url'.
 *
 * The library reaches for this module only to turn a resolved file path into a `file://` URL
 * for the PDF.js worker, on code paths that require Node and are already inside a try/catch.
 * Left unresolved it reaches the output as a bare Node built-in, which a consumer's bundler
 * reports as a missing module in a browser build (issue #108).
 *
 * `URL` and `URLSearchParams` are forwarded to the platform's own implementations rather than
 * stubbed out, since browsers provide both and any consumer reaching for them through this
 * module should keep working.
 */

export const URL = globalThis.URL;
export const URLSearchParams = globalThis.URLSearchParams;

export const pathToFileURL = (path) => new globalThis.URL(`file://${path}`);

export const fileURLToPath = (url) => {
    const href = typeof url === 'string' ? url : url?.href;
    if (typeof href === 'string' && href.startsWith('file://')) return href.slice('file://'.length);
    throw new Error("officeparser: 'url.fileURLToPath' requires a file:// URL.");
};

export default { URL, URLSearchParams, pathToFileURL, fileURLToPath };

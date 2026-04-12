/**
 * Browser-side stub for Node.js 'fs' module.
 * 
 * officeparser is designed to work in the browser by accepting 
 * Buffer or Uint8Array inputs instead of file paths.
 */

function throwBrowserError(fnName) {
    throw new Error(`officeparser: 'fs.${fnName}' is not supported in the browser. Callers must pass file content as Buffer or Uint8Array directly to the parser.`);
}

export const readFileSync = (path) => throwBrowserError('readFileSync');
export const readFile = (path, cb) => throwBrowserError('readFile');
export const existsSync = (path) => false;
export const statSync = (path) => throwBrowserError('statSync');
export const lstatSync = (path) => throwBrowserError('lstatSync');
export const readdirSync = (path) => throwBrowserError('readdirSync');
export const mkdirSync = (path) => throwBrowserError('mkdirSync');
export const writeFileSync = (path, data) => throwBrowserError('writeFileSync');

export default {
    readFileSync,
    readFile,
    existsSync,
    statSync,
    lstatSync,
    readdirSync,
    mkdirSync,
    writeFileSync,
    promises: {
        readFile: (path) => Promise.reject(new Error("officeparser: 'fs.promises.readFile' is not supported in the browser.")),
        writeFile: (path, data) => Promise.reject(new Error("officeparser: 'fs.promises.writeFile' is not supported in the browser.")),
        readdir: (path) => Promise.reject(new Error("officeparser: 'fs.promises.readdir' is not supported in the browser.")),
    }
};

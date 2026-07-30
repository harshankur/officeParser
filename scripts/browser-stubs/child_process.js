/**
 * Browser-side stub for Node.js 'child_process'.
 *
 * Only reached from host-inspection code that never runs in a browser: the Rosetta check in
 * PdfGenerator shells out to `sysctl`, and is guarded by `process.platform === 'darwin'`. The
 * stub exists so bundlers have something to resolve, rather than reporting an unresolvable
 * Node built-in in a browser build (issue #108).
 */

function throwBrowserError(fnName) {
    throw new Error(`officeparser: 'child_process.${fnName}' is not supported in the browser.`);
}

export const execSync = () => throwBrowserError('execSync');
export const exec = () => throwBrowserError('exec');
export const spawn = () => throwBrowserError('spawn');
export const spawnSync = () => throwBrowserError('spawnSync');

export default { execSync, exec, spawn, spawnSync };

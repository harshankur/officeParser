/**
 * Code a browser bundle evaluates from a string as it loads.
 *
 * The Node `crypto` polyfill the browser bundles carry (for password-protected documents) includes
 * `is-generator-function`, which finds the generator function constructor by running
 * `Function("return function*() {}")()` once, when the module loads. A page whose Content Security
 * Policy has no `'unsafe-eval'` blocks that call. It sits in a `try`, so the bundle went on working,
 * but the browser reported a `script-src` violation on every load, and an app that collects those
 * reports could not tell it from a real one.
 *
 * The call is replaced after bundling by the value it evaluates to, which is the same function with
 * nothing evaluated. The rule lives here, rather than inside the build script, so the test that
 * guards the shipped bundles applies exactly the same one.
 */

/** Each call that runs at load, and the expression it evaluates to. */
const LOAD_TIME_EVALS = [
    { call: 'Function("return function*() {}")()', value: '(function*() {})' },
];

/**
 * A generator function made by the Function constructor, however the call is spelled. One the list
 * above does not name (a dependency that changed its quotes or its spacing) is still found by this,
 * so it fails the build and the artifact test instead of shipping.
 */
const GENERATOR_PROBE = /Function\(\s*(["'`])return function\s*\*/g;

/**
 * The load-time evaluations left in a bundle's source.
 *
 * @param {string} source
 * @returns {{ index: number, snippet: string }[]}
 */
function findLoadTimeEvals(source) {
    const found = [];
    for (const match of source.matchAll(GENERATOR_PROBE)) {
        found.push({ index: match.index, snippet: source.slice(match.index, match.index + 48) });
    }
    return found;
}

/**
 * Replaces every known load-time evaluation by its value.
 *
 * @param {string} source
 * @returns {{ output: string, replaced: number }}
 */
function replaceLoadTimeEvals(source) {
    let output = source;
    let replaced = 0;
    for (const { call, value } of LOAD_TIME_EVALS) {
        const parts = output.split(call);
        replaced += parts.length - 1;
        output = parts.join(value);
    }
    return { output, replaced };
}

module.exports = { findLoadTimeEvals, replaceLoadTimeEvals };

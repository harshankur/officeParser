/**
 * Dynamic-import analysis for the browser bundles.
 *
 * Bundlers try to resolve `import(...)` at build time. Given a specifier they cannot evaluate,
 * webpack does not skip it: it creates a *context module* covering the whole directory the
 * expression could resolve to, and pulls every file in that directory into the build. For a
 * published package that means `dist/`, so a consumer bundling the browser build ends up with
 * the Node-only files too and their build fails on `child_process` and friends (issue #108).
 * Vite is less drastic but warns.
 *
 * Both bundlers offer an opt-out via a leading comment, so any specifier that is not statically
 * analyzable gets annotated after bundling. The classifier lives here, rather than inside the
 * build script, so the test that guards the shipped bundles applies exactly the same rule.
 */

/** Annotations that tell each bundler to leave a dynamic import alone. */
const IGNORE_ANNOTATION = '/* @vite-ignore */ /* webpackIgnore: true */ ';

/** Length of the `import(` token the scan anchors on. */
const IMPORT_CALL = 'import(';

/**
 * Whether a template literal starting at `start` interpolates anything.
 *
 * A template with no substitutions is as static as a quoted string, so it needs no annotation.
 * One that interpolates cannot be resolved at build time and does.
 *
 * @param {string} source
 * @param {number} start - Index of the opening backtick
 * @returns {boolean}
 */
function templateInterpolates(source, start) {
    for (let i = start + 1; i < source.length; i++) {
        const char = source[i];
        if (char === '\\') { i++; continue; }
        if (char === '`') return false;
        if (char === '$' && source[i + 1] === '{') return true;
    }
    return false;
}

/**
 * Finds dynamic imports a bundler cannot resolve statically and that carry no ignore annotation.
 *
 * @param {string} source - Bundle contents
 * @returns {Array<{ index: number, snippet: string }>} One entry per offending import
 */
function findUnanalyzableDynamicImports(source) {
    const found = [];
    let index = source.indexOf(IMPORT_CALL);

    while (index !== -1) {
        // Only a genuine call: `import(` preceded by an identifier character would be part of a
        // longer name, and `.import(` a method call on something else.
        const before = index === 0 ? '' : source[index - 1];
        if (!/[\w$.]/.test(before)) {
            let cursor = index + IMPORT_CALL.length;
            while (cursor < source.length && /\s/.test(source[cursor])) cursor++;
            const rest = source.slice(cursor);

            const alreadyAnnotated = rest.startsWith('/* @vite-ignore */') || rest.startsWith('/* webpackIgnore: true */');
            const quotedLiteral = rest[0] === '"' || rest[0] === "'";
            const staticTemplate = rest[0] === '`' && !templateInterpolates(source, cursor);

            if (!alreadyAnnotated && !quotedLiteral && !staticTemplate) {
                found.push({ index, snippet: source.slice(index, index + 80).split('\n')[0] });
            }
        }
        index = source.indexOf(IMPORT_CALL, index + IMPORT_CALL.length);
    }

    return found;
}

/**
 * Adds ignore annotations to every dynamic import a bundler cannot resolve statically.
 *
 * The annotation goes immediately after `import(`, so the specifier expression is never parsed
 * or rewritten. An earlier version tried to capture the whole expression and put it back, which
 * could not survive nested parentheses and skipped template literals outright, leaving exactly
 * the import that broke webpack builds.
 *
 * @param {string} source - Bundle contents
 * @returns {{ output: string, annotated: number }}
 */
function annotateDynamicImports(source) {
    const targets = findUnanalyzableDynamicImports(source);
    if (targets.length === 0) return { output: source, annotated: 0 };

    // Applied back to front so earlier offsets stay valid.
    let output = source;
    for (let i = targets.length - 1; i >= 0; i--) {
        const at = targets[i].index + IMPORT_CALL.length;
        output = output.slice(0, at) + IGNORE_ANNOTATION + output.slice(at);
    }
    return { output, annotated: targets.length };
}

module.exports = { findUnanalyzableDynamicImports, annotateDynamicImports, IGNORE_ANNOTATION };

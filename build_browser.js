const esbuild = require('esbuild');
const { polyfillNode } = require('esbuild-plugin-polyfill-node');

esbuild.build({
    entryPoints: ['src/index.ts'],
    bundle: true,
    outfile: 'dist/officeparser.browser.js',
    format: 'iife',
    globalName: 'officeParser',
    platform: 'browser',
    target: ['es2020'],
    sourcemap: true,
    minify: true,
    define: {
        'process.env.NODE_ENV': '"production"',
        'global': 'window',
        'import.meta.url': '""'  // Shim import.meta.url to avoid SyntaxError in classic script
    },
    banner: {
        js: `
// Shim for setImmediate
if (typeof setImmediate === 'undefined') {
  window.setImmediate = function(callback) {
    return setTimeout(callback, 0);
  };
}
`
    },
    plugins: [
        polyfillNode({
            polyfills: {
                fs: true,
                stream: true,
                buffer: true,
                events: true
            }
        })
    ]
}).catch(() => process.exit(1));

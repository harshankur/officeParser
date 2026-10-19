/**
 * Browser-side stub for '@pdf-lib/fontkit'.
 *
 * fontkit is an optional peer dependency used only by the native PDF engine, and only when fonts are
 * given (pdfConfig.fonts, pdfConfig.fallbackFonts). Like pdf-lib it is deliberately NOT bundled into
 * the prebuilt browser bundles: the `officeparser/browser-native-pdf` entry leaves both external for
 * the consumer's bundler to resolve. In the other bundles the native engine is unavailable anyway
 * (see ./pdf-lib.js), so this is never reached; it throws a clear error if it ever is.
 */

const unavailable = () => {
    throw new Error("officeparser: embedding fonts in the native PDF engine requires '@pdf-lib/fontkit', which is not included in the prebuilt browser bundle. Install pdf-lib and @pdf-lib/fontkit and import 'officeparser/browser-native-pdf'.");
};

export const create = unavailable;

export default { create };

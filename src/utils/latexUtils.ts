import { lookupTable } from './lookupUtils.js';

/**
 * Character-coverage helpers for LaTeX output.
 *
 * pdfLaTeX reads UTF-8 through `inputenc`, which only knows the characters the loaded font
 * encodings (T1, TS1, OT1, OMS) declare; any other character is a fatal "Unicode character ... not
 * set up for use with LaTeX" error. XeLaTeX and LuaLaTeX accept every character but draw nothing for
 * one the font lacks, and the default Latin Modern has no dingbats, arrows past the basic four, or
 * most math symbols. Office documents use exactly those characters (check marks, bullets, arrows,
 * emoji, unusual spaces), so the generator declares a fallback for each one the document contains
 * rather than leaving the output to fail or go blank.
 *
 * @module latexUtils
 */

/**
 * Code points pdfLaTeX can typeset with `\usepackage[T1]{fontenc}` and the default text encodings,
 * as inclusive ranges. Derived from TeX Live 2026's `t1enc.dfu`, `ts1enc.dfu`, `ot1enc.dfu` and
 * `omsenc.dfu` plus the encoding-independent entries of `utf8enc.dfu` (`\textellipsis` and the two
 * spacing accents). Printable ASCII is handled separately.
 */
const PDFTEX_TEXT_RANGES: ReadonlyArray<readonly [number, number]> = [
    [0x00A0, 0x0125], [0x0128, 0x0137], [0x0139, 0x013E], [0x0141, 0x0148], [0x014A, 0x0165], [0x0168, 0x017E],
    [0x0192, 0x0192], [0x01C4, 0x01D4], [0x01E2, 0x01E3], [0x01E6, 0x01EB], [0x01F0, 0x01F0], [0x01F4, 0x01F5],
    [0x0218, 0x021B], [0x0232, 0x0233], [0x0237, 0x0237], [0x02C6, 0x02C7], [0x02D8, 0x02D9], [0x02DB, 0x02DD],
    [0x0E3F, 0x0E3F], [0x1E02, 0x1E03], [0x1E0D, 0x1E0D], [0x1E1E, 0x1E21], [0x1E25, 0x1E25], [0x1E30, 0x1E31],
    [0x1E37, 0x1E37], [0x1E43, 0x1E43], [0x1E45, 0x1E45], [0x1E47, 0x1E47], [0x1E5B, 0x1E5B], [0x1E63, 0x1E63],
    [0x1E6D, 0x1E6D], [0x1E8E, 0x1E91], [0x1E9E, 0x1E9E], [0x1EF2, 0x1EF3], [0x200C, 0x200C], [0x2010, 0x2016],
    [0x2018, 0x201A], [0x201C, 0x201E], [0x2020, 0x2022], [0x2026, 0x2026], [0x2030, 0x2031], [0x2039, 0x203B],
    [0x203D, 0x203D], [0x2044, 0x2044], [0x204E, 0x204E], [0x2052, 0x2052], [0x20A1, 0x20A1], [0x20A4, 0x20A4],
    [0x20A6, 0x20A6], [0x20A9, 0x20A9], [0x20AB, 0x20AC], [0x20B1, 0x20B1], [0x2103, 0x2103], [0x2116, 0x2117],
    [0x211E, 0x211E], [0x2120, 0x2120], [0x2122, 0x2122], [0x2126, 0x2127], [0x212E, 0x212E], [0x2190, 0x2193],
    [0x2329, 0x232A], [0x2422, 0x2423], [0x25E6, 0x25E6], [0x25EF, 0x25EF], [0x266A, 0x266A], [0x27E8, 0x27E9],
    [0x3008, 0x3009], [0xFB00, 0xFB06], [0xFEFF, 0xFEFF],
];

/** Whether pdfLaTeX can typeset a code point with the generator's font setup. */
export function isPdfTexSupported(codePoint: number): boolean {
    if (codePoint === 0x09 || codePoint === 0x0A || (codePoint >= 0x20 && codePoint <= 0x7E)) return true;
    for (const [lo, hi] of PDFTEX_TEXT_RANGES) {
        if (codePoint < lo) return false;
        if (codePoint <= hi) return true;
    }
    return false;
}

/** A replacement for one character and the package it needs, if any. */
interface LatexFallback { tex: string; pkg?: 'amssymb' | 'pifont' }

const math = (cmd: string): LatexFallback => ({ tex: `\\ensuremath{${cmd}}`, pkg: 'amssymb' });
const ding = (n: number): LatexFallback => ({ tex: `\\ding{${n}}`, pkg: 'pifont' });
const space = (cmd: string): LatexFallback => ({ tex: cmd });

/**
 * Engine-independent replacements for symbols and spaces the default fonts lack, declared with
 * `\newunicodechar` so the body keeps the original characters. Covers the typographic spaces, the
 * common math and arrow symbols, and the dingbats office documents use for check marks, bullets
 * and stars.
 */
const SYMBOL_FALLBACKS: ReadonlyMap<number, LatexFallback> = new Map<number, LatexFallback>([
    // Spaces (U+2000 to U+200A, narrow no-break, medium math, ideographic) and invisible joiners.
    [0x2000, space('\\enspace')], [0x2001, space('\\quad')], [0x2002, space('\\enspace')], [0x2003, space('\\quad')],
    [0x2004, space('\\;')], [0x2005, space('\\:')], [0x2006, space('\\,')], [0x2007, space('\\enspace')],
    [0x2008, space('\\,')], [0x2009, space('\\,')], [0x200A, space('\\,')], [0x202F, space('\\,')],
    [0x205F, space('\\:')], [0x3000, space('\\quad')], [0x200D, space('{}')], [0x2060, space('{}')],
    // Math and relations.
    [0x2212, math('-')], [0x2213, math('\\mp')], [0x2264, math('\\le')], [0x2265, math('\\ge')], [0x2260, math('\\neq')],
    [0x2248, math('\\approx')], [0x2261, math('\\equiv')], [0x221E, math('\\infty')], [0x2219, math('\\bullet')],
    [0x22C5, math('\\cdot')], [0x221A, math('\\surd')], [0x2211, math('\\sum')], [0x220F, math('\\prod')],
    [0x222B, math('\\int')], [0x2202, math('\\partial')], [0x2206, math('\\Delta')], [0x2207, math('\\nabla')],
    [0x2208, math('\\in')], [0x2209, math('\\notin')], [0x2282, math('\\subset')], [0x2283, math('\\supset')],
    [0x2286, math('\\subseteq')], [0x2287, math('\\supseteq')], [0x2229, math('\\cap')], [0x222A, math('\\cup')],
    [0x2200, math('\\forall')], [0x2203, math('\\exists')], [0x2205, math('\\emptyset')], [0x2227, math('\\wedge')],
    [0x2228, math('\\vee')], [0x2295, math('\\oplus')], [0x2297, math('\\otimes')], [0x2234, math('\\therefore')],
    [0x221D, math('\\propto')], [0x2220, math('\\angle')], [0x22A5, math('\\perp')], [0x2225, math('\\parallel')],
    [0x2032, math('{}^{\\prime}')], [0x2033, math('{}^{\\prime\\prime}')],
    // Arrows past the four TS1 has.
    [0x2194, math('\\leftrightarrow')], [0x2195, math('\\updownarrow')], [0x2196, math('\\nwarrow')], [0x2197, math('\\nearrow')],
    [0x2198, math('\\searrow')], [0x2199, math('\\swarrow')], [0x21A6, math('\\mapsto')], [0x21D0, math('\\Leftarrow')],
    [0x21D2, math('\\Rightarrow')], [0x21D4, math('\\Leftrightarrow')], [0x21D1, math('\\Uparrow')], [0x21D3, math('\\Downarrow')],
    [0x27F5, math('\\longleftarrow')], [0x27F6, math('\\longrightarrow')], [0x27F9, math('\\Longrightarrow')],
    // Shapes, check marks and dingbats.
    [0x2713, math('\\checkmark')], [0x2714, ding(52)], [0x2717, ding(55)], [0x2718, ding(56)], [0x2605, ding(72)],
    [0x2606, ding(73)], [0x25CF, ding(108)], [0x25A0, ding(110)], [0x25A1, math('\\square')], [0x25AA, math('\\blacksquare')],
    [0x25AB, math('\\square')], [0x25C6, ding(117)], [0x25C7, math('\\diamond')], [0x25B6, ding(229)], [0x25BA, ding(229)],
    [0x25B2, math('\\blacktriangle')], [0x25BC, math('\\blacktriangledown')], [0x25B3, math('\\triangle')],
    [0x25BD, math('\\triangledown')], [0x25CB, math('\\circ')], [0x2610, math('\\square')], [0x2611, math('\\boxtimes')],
    [0x2612, math('\\boxtimes')], [0x260E, ding(37)], [0x2709, ding(41)], [0x2702, ding(34)], [0x2708, ding(40)],
    [0x261B, ding(42)], [0x261E, ding(43)], [0x2660, math('\\spadesuit')], [0x2663, math('\\clubsuit')],
    [0x2665, math('\\heartsuit')], [0x2666, math('\\diamondsuit')], [0x2764, ding(170)], [0x2756, ding(118)],
    [0x27A2, ding(226)], [0x2794, ding(212)], [0x279C, ding(220)],
]);

/** Greek letters pdfLaTeX can draw as math symbols (upright capitals that look Latin use the Latin letter). */
const GREEK_LOWER = ['alpha', 'beta', 'gamma', 'delta', 'epsilon', 'zeta', 'eta', 'theta', 'iota', 'kappa', 'lambda', 'mu', 'nu', 'xi', 'o', 'pi', 'rho', 'varsigma', 'sigma', 'tau', 'upsilon', 'phi', 'chi', 'psi', 'omega'];
const GREEK_UPPER = ['A', 'B', 'Gamma', 'Delta', 'E', 'Z', 'H', 'Theta', 'I', 'K', 'Lambda', 'M', 'N', 'Xi', 'O', 'Pi', 'P', '', 'Sigma', 'T', 'Upsilon', 'Phi', 'X', 'Psi', 'Omega'];

/** pdfLaTeX-only replacement for a Greek letter, or undefined. */
function greekFallback(cp: number): LatexFallback | undefined {
    const lower = cp - 0x03B1;
    if (lower >= 0 && lower < GREEK_LOWER.length) {
        const name = GREEK_LOWER[lower];
        return name === 'o' ? { tex: 'o' } : math(`\\${name}`);
    }
    const upper = cp - 0x0391;
    if (upper >= 0 && upper < GREEK_UPPER.length && GREEK_UPPER[upper]) {
        const name = GREEK_UPPER[upper];
        return name.length === 1 ? { tex: name } : math(`\\${name}`);
    }
    return undefined;
}

/**
 * Scripts a text uses that Latin Modern, the default font under XeLaTeX and LuaLaTeX, has no glyphs
 * for, so those engines need other fonts for them: Greek and Cyrillic, and the CJK scripts (Han,
 * kana, Hangul). `cjkLanguage` is the language its CJK characters are mostly in (Japanese when kana
 * are a good part of them, Korean when Hangul outnumbers the rest, else Chinese), whose font sets
 * them; `cjkDominant` says CJK characters outnumber Latin letters, which decides whether the
 * punctuation both share (curly quotes, dashes, the ellipsis) takes the CJK or the Latin font.
 */
export interface LatexScripts {
    greekCyrillic: boolean;
    han: boolean;
    kana: boolean;
    hangul: boolean;
    cjkLanguage?: 'ja' | 'ko' | 'zh';
    cjkDominant: boolean;
}

/**
 * The preamble declarations a piece of LaTeX needs for its non-ASCII characters.
 *
 * - `unicodeEngines`: `\newunicodechar` fallbacks under XeLaTeX and LuaLaTeX, for the symbols and
 *   spaces the default fonts lack;
 * - `eightBitEngines`: `\DeclareUnicodeCharacter` lines for pdfLaTeX (and upLaTeX/pLaTeX, for the
 *   characters they do not read as CJK; `\newunicodechar` fails there on those they do): the same
 *   fallbacks, and replacements for characters pdfLaTeX cannot typeset at all (Greek becomes math
 *   letters, anything else a visible `[U+XXXX]` marker). XeLaTeX and LuaLaTeX keep the real
 *   character there, set in a font that covers its script (see `scripts`);
 * - `scripts`: the scripts XeLaTeX and LuaLaTeX need fonts beyond Latin Modern for;
 * - `packages`: the packages the fallbacks draw on.
 */
export interface LatexUnicodePlan {
    unicodeEngines: string[];
    eightBitEngines: string[];
    scripts: LatexScripts;
    packages: Set<'amssymb' | 'pifont' | 'newunicodechar'>;
}

/** Whether a code point is Greek or Cyrillic (with their extended and supplementary blocks). */
const isGreekCyrillic = (cp: number) => (cp >= 0x0370 && cp <= 0x052F) || (cp >= 0x1C80 && cp <= 0x1C8F) || (cp >= 0x1F00 && cp <= 0x1FFF)
    || (cp >= 0x2DE0 && cp <= 0x2DFF) || (cp >= 0xA640 && cp <= 0xA69F);
/** Han ideographs (unified, extensions, compatibility). */
const isHan = (cp: number) => (cp >= 0x3400 && cp <= 0x4DBF) || (cp >= 0x4E00 && cp <= 0x9FFF) || (cp >= 0xF900 && cp <= 0xFAFF)
    || (cp >= 0x20000 && cp <= 0x3134F);
/** Hiragana, katakana and their extensions (half-width katakana included). */
const isKana = (cp: number) => (cp >= 0x3040 && cp <= 0x30FF) || (cp >= 0x31F0 && cp <= 0x31FF) || (cp >= 0xFF65 && cp <= 0xFF9F);
/** Hangul syllables and jamo. */
const isHangul = (cp: number) => (cp >= 0xAC00 && cp <= 0xD7FF) || (cp >= 0x1100 && cp <= 0x11FF) || (cp >= 0x3130 && cp <= 0x318F) || (cp >= 0xA960 && cp <= 0xA97F);
/** CJK punctuation and full-width forms, which count toward a text being CJK. */
const isCjkPunctuation = (cp: number) => (cp >= 0x3000 && cp <= 0x303F) || (cp >= 0xFF00 && cp <= 0xFF64);

/** A code point as `\DeclareUnicodeCharacter` takes it: four or more uppercase hex digits. */
const hexOf = (cp: number) => cp.toString(16).toUpperCase().padStart(4, '0');

/** Plans the fallbacks for every character of `text` that some engine cannot typeset as is. */
export function planLatexUnicode(text: string): LatexUnicodePlan {
    const scripts: LatexScripts = { greekCyrillic: false, han: false, kana: false, hangul: false, cjkDominant: false };
    const plan: LatexUnicodePlan = { unicodeEngines: [], eightBitEngines: [], scripts, packages: new Set() };
    const seen = new Set<number>();
    // Letters of the text itself, not of the commands around it, decide which script dominates.
    let latin = 0, cjk = 0, han = 0, kana = 0, hangul = 0;
    for (const ch of text.replace(/\\[A-Za-z@]+/g, '')) {
        const cp = ch.codePointAt(0)!;
        if ((cp >= 0x41 && cp <= 0x5A) || (cp >= 0x61 && cp <= 0x7A) || (cp >= 0xC0 && cp <= 0x24F)) latin++;
        else if (cp >= 0x0370) {
            if (isGreekCyrillic(cp)) scripts.greekCyrillic = true;
            else if (isHan(cp)) { han++; cjk++; }
            else if (isKana(cp)) { kana++; cjk++; }
            else if (isHangul(cp)) { hangul++; cjk++; }
            else if (isCjkPunctuation(cp)) cjk++;
        }
    }
    Object.assign(scripts, { han: han > 0, kana: kana > 0, hangul: hangul > 0, cjkDominant: cjk > latin });
    if (han + kana + hangul > 0) scripts.cjkLanguage = hangul >= han + kana ? 'ko' : kana * 5 >= han ? 'ja' : 'zh';
    for (const ch of text) {
        const cp = ch.codePointAt(0)!;
        if (cp < 0x80 || seen.has(cp)) continue;
        seen.add(cp);
        const symbol = SYMBOL_FALLBACKS.get(cp);
        if (symbol) {
            plan.unicodeEngines.push(`\\newunicodechar{${ch}}{${symbol.tex}}`);
            plan.eightBitEngines.push(`\\DeclareUnicodeCharacter{${hexOf(cp)}}{${symbol.tex}}`);
            plan.packages.add('newunicodechar');
            if (symbol.pkg) plan.packages.add(symbol.pkg);
            continue;
        }
        if (isPdfTexSupported(cp)) continue;
        const hex = hexOf(cp);
        const greek = greekFallback(cp);
        if (greek?.pkg) plan.packages.add(greek.pkg);
        plan.eightBitEngines.push(`\\DeclareUnicodeCharacter{${hex}}{${greek ? greek.tex : `{\\ttfamily[U+${hex}]}`}}`);
    }
    plan.unicodeEngines.sort();
    plan.eightBitEngines.sort();
    return plan;
}

/**
 * Code-block languages `listings` ships a definition for, keyed by the names documents use. A
 * language `listings` does not know is an error ("Couldn't load requested language"), so anything
 * missing here is written as plain `verbatim`.
 */
export const LISTINGS_LANGUAGES: Record<string, string> = lookupTable({
    python: 'Python', py: 'Python', java: 'Java', c: 'C', h: 'C', cpp: 'C++', 'c++': 'C++', cc: 'C++', cxx: 'C++', hpp: 'C++',
    csharp: '[Sharp]C', cs: '[Sharp]C', 'c#': '[Sharp]C', ruby: 'Ruby', rb: 'Ruby', php: 'PHP', perl: 'Perl', pl: 'Perl',
    sql: 'SQL', bash: 'bash', sh: 'sh', shell: 'bash', zsh: 'bash', ksh: 'ksh', csh: 'csh', html: 'HTML', xml: 'XML',
    xslt: 'XSLT', tex: 'TeX', latex: '[LaTeX]TeX', r: 'R', matlab: 'Matlab', octave: 'Octave', haskell: 'Haskell', hs: 'Haskell',
    lisp: 'Lisp', elisp: 'Lisp', fortran: 'Fortran', pascal: 'Pascal', delphi: 'Delphi', erlang: 'erlang', scilab: 'Scilab',
    ocaml: '[Objective]Caml', ml: 'ML', make: 'make', makefile: 'make', awk: 'Awk', tcl: 'tcl', vbscript: 'VBScript',
    verilog: 'Verilog', vhdl: 'VHDL', gnuplot: 'Gnuplot', prolog: 'Prolog', cobol: 'Cobol', ada: 'Ada', mathematica: 'Mathematica',
    sparql: 'SPARQL', postscript: 'PostScript', ps: 'PostScript', lua: '[5.3]Lua', go: 'Go',
});

/**
 * The name documents conventionally use for each `listings` language, for reading a `lstlisting`
 * back: the first key mapped to that language in {@link LISTINGS_LANGUAGES} (so `Python` reads
 * as `python`, `[Sharp]C` as `csharp`). Keyed case-insensitively.
 */
export const LISTINGS_LANGUAGE_NAMES: ReadonlyMap<string, string> = (() => {
    const map = new Map<string, string>();
    for (const [name, listings] of Object.entries(LISTINGS_LANGUAGES)) {
        const key = listings.toLowerCase();
        if (!map.has(key)) map.set(key, name);
    }
    return map;
})();

/**
 * The character each engine-independent fallback in {@link SYMBOL_FALLBACKS} stands for, keyed by
 * the replacement's LaTeX (`\ding{170}` reads back as the heart), so a document using those
 * commands directly parses to the same characters.
 */
export const LATEX_SYMBOL_CHARACTERS: ReadonlyMap<string, string> = (() => {
    const map = new Map<string, string>();
    for (const [cp, fallback] of SYMBOL_FALLBACKS) {
        if (!map.has(fallback.tex)) map.set(fallback.tex, String.fromCodePoint(cp));
    }
    return map;
})();

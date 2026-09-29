import { MATHTOOLS_ENVIRONMENTS } from './sanitize.js';
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


/**
 * The commands a formula may use once the LaTeX generator's preamble has loaded amsmath and amssymb:
 * what TeX Live 2026 defines under pdfLaTeX in the kernel's plain, math, spacing, font, text-symbol,
 * box, reference and tabular parts (`latex.ltx`), in `fontmath.ltx`, and in amsmath and amssymb
 * (with amsfonts, amsopn, amsbsy and amstext), less those `sanitizeLatexMath` refuses and the
 * kernel's package-writing commands. Found by reading those files' command names and keeping the
 * ones `\ifcsname` finds defined. A formula using anything else needs a package of
 * {@link MATH_PACKAGE_COMMANDS}, or is a command no package the output loads defines.
 */
const LATEX_MATH_COMMANDS: ReadonlySet<string> = new Set(`
    AA AE Acute Alph AmS AmSfont And Arrowvert AssignSocketPlug Bar Bbb Bbbk Big Bigg Biggl Biggm Biggr Bigl Bigm Bigr
    Box Breve Bumpeq Cap Check CheckEncodingSubset Cup DOTSB DOTSI DOTSX Ddot Delta Diamond Dot Doteq Downarrow Finv
    Game Gamma Grave H Hat IJ IfBooleanF IfBooleanTF Im Join L LaTeX LaTeXe Lambda Leftarrow Leftrightarrow Lleftarrow
    Longleftarrow Longleftrightarrow Longrightarrow Lsh MakeUppercase MultiIntegral O OE Omega P Phi Pi Pr Psi Re Ref
    Relbar Rightarrow Roman Rrightarrow Rsh S SS Sigma Subset Supset TeX TextOrMath TextSymbolUnavailable Theta Tilde
    UndeclareTextCommand Uparrow Updownarrow Upsilon UseHookWithArguments UseLegacyTextSymbols UseSocket
    UseTaggingSocket UseTextAccent UseTextSymbol Vdash Vec Vert Vvdash Xi a aa above abovedisplayshortskip
    abovedisplayskip abovewithdelims accent acute addpenalty addvspace ae aleph align aligned alignedat allowbreak
    allowdisplaybreaks alph alpha amalg angle approx approxeq arabic arccos arcsin arctan arg array arraycolsep
    arrayrulewidth arraystretch arrowvert ast asymp atop atopwithdelims b backepsilon backprime backsim backsimeq
    backslash bar barwedge because begin begingroup belowdisplayshortskip belowdisplayskip beta beth between bfseries
    bgroup big bigcap bigcirc bigcup bigg biggl biggm biggr bigl bigm bigodot bigoplus bigotimes bigr bigskip
    bigskipamount bigsqcup bigstar bigtriangledown bigtriangleup biguplus bigvee bigwedge binom blacklozenge blacksquare
    blacktriangle blacktriangledown blacktriangleleft blacktriangleright bmod bold boldmath boldsymbol bordermatrix bot
    bowtie box boxdot boxed boxmaxdepth boxminus boxplus boxtimes brace braceld bracelu bracerd braceru bracevert brack
    break breve buildrel bullet bumpeq c cap capitalacute capitalbreve capitalcaron capitalcedilla capitalcircumflex
    capitaldieresis capitaldotaccent capitalgrave capitalhungarumlaut capitalmacron capitalnewtie capitalogonek
    capitalring capitaltie capitaltilde cases cdot cdotp cdots centerdot centerline cfrac check checkmark chi choose
    circ circeq circlearrowleft circlearrowright circledR circledS circledast circledcirc circleddash clap cline
    clubsuit colon complement cong coprod copy copyright cos cosh cot coth csc cup curlyeqprec curlyeqsucc curlyvee
    curlywedge currentgrouptype curvearrowleft curvearrowright d dag dagger daleth dasharrow dashleftarrow
    dashrightarrow dashv dbinom ddag ddagger ddddot dddot ddot ddots deg delimiter delimiterfactor delimitershortfall
    delta det dfrac diagdown diagup diamond diamondsuit digamma dim dimen discretionary displaybreak displayindent
    displaylimits displaylines displaymath displaystyle displaywidowpenalty displaywidth div divide divideontimes dot
    doteq doteqdot dotfill dotplus dots dotsb dotsc dotsi dotsm dotso doublebarwedge doublecap doublecup
    doublehyphendemerits doublerulesep downarrow downbracefill downdownarrows downharpoonleft downharpoonright dp egroup
    ell em emph empty emptyset end endalign endaligned endarray enddisplaymath endeqnarray endequation endgather endgraf
    endgroup endline endmath endmathdisplay endmatrix endminipage endmultline endsplit endsubarray endtabbing endtabular
    endtrivlist enskip enspace ensuremath epsilon eqcirc eqnarray eqno eqref eqsim eqslantgtr eqslantless equation equiv
    eta eth exhyphenpenalty exists exp extracolsep fallingdotseq fbox fboxrule fboxsep fill finalhyphendemerits flat
    fmtversion fnsymbol fontchardp fontcharht fontdimen fontencoding fontfamily fontsize footnotesize forall frac frak
    frame framebox frenchspacing frown gamma gather gcd ge genfrac geq geqq geqslant gets gg ggg gggtr gimel glossary
    gnapprox gneq gneqq gnsim grave gtrapprox gtrdot gtreqless gtreqqless gtrless gtrsim gvertneqq halign hat hbadness
    hbar hbox hdots hdotsfor heartsuit hfil hfill hfilneg hfuzz hglue hideskip hidewidth hline hom hookleftarrow
    hookrightarrow hphantom hrule hrulefill hsize hskip hslash hspace hss ht hyphenpenalty i ialign idotsint iff
    ignorespaces ignorespacesafterend iiiint iiint iint ij imath impliedby implies in indent index inf infty injlim int
    intercal interdisplaylinepenalty interfootnotelinepenalty interlinepenalty intertext intop iota item itshape j jmath
    joinrel jot k kappa ker kern kill l lVert label labelformat lambda land langle lbrace lbrack lceil ldotp ldots le
    leaders leadsto leavevmode left leftarrow leftarrowfill leftarrowtail lefteqn leftharpoondown leftharpoonup
    leftleftarrows leftline leftrightarrow leftrightarrows leftrightharpoons leftrightsquigarrow leftroot leftskip
    leftthreetimes legacyoldstylenums leq leqno leqq leqslant lessapprox lessdot lesseqgtr lesseqqgtr lessgtr lesssim
    lfloor lg lgroup lhd lhook lim liminf limits limsup linepenalty lineskip lineskiplimit linewidth ll llap llcorner
    lll llless lmoustache ln lnapprox lneq lneqq lnot lnsim log longleftarrow longleftrightarrow longmapsto
    longrightarrow lor lower lowercase lozenge lq lrcorner ltimes lvert lvertneqq magstep magstephalf makebox maltese
    mapsto mapstochar math mathaccent mathaccentV mathalpha mathbb mathbf mathbin mathcal mathchar mathchoice mathclose
    mathdisplay mathdollar mathellipsis mathfrak mathgroup mathhexbox mathinner mathit mathnormal mathop mathopen
    mathord mathpalette mathparagraph mathpunct mathrel mathring mathrm mathsection mathsf mathsterling mathstrut
    mathsurround mathtt mathunderscore mathversion matrix max maxdepth maxdimen mbox mdseries measuredangle medmuskip
    medskip medskipamount medspace mho mid middle min minalignsep minipage mintagsep mkern mod models moveleft mp mskip
    mspace mu multicolumn multimap multiply multispan multline multlinegap multlinetaggap muskip nLeftarrow
    nLeftrightarrow nRightarrow nVDash nVdash nabla narrower natural ncong ne nearrow neg negmedspace negthickspace
    negthinspace neq nexists ngeq ngeqq ngeqslant ngtr ni nleftarrow nleftrightarrow nleq nleqq nleqslant nless nmid
    nobreak nobreakdash nobreakdashes nobreakspace nocorr nocorrlist noindent nointerlineskip nolimits nonfrenchspacing
    nonscript nonumber normalbaselines normalbaselineskip normalcolor normalfont normallineskip normallineskiplimit
    normalsize not notag notin nparallel nprec npreceq nrightarrow nshortmid nshortparallel nsim nsubseteq nsubseteqq
    nsucc nsucceq nsupseteq nsupseteqq ntriangleleft ntrianglelefteq ntriangleright ntrianglerighteq nu null
    nulldelimiterspace nvDash nvdash nwarrow o oalign odot oe offinterlineskip oint ointop oldstylenums omega ominus
    ooalign openup operatorfont operatorname operatornamewithlimits oplus oslash otimes over overbrace overfullrule
    overleftarrow overleftrightarrow overline overrightarrow overset overunderset overwithdelims owns pageref parallel
    parbox parfillskip parindent parshape parskip partial penalty perp phantom phi pi pitchfork pm pmatrix pmb pmod pod
    poptabs postdisplaypenalty pounds prec precapprox preccurlyeq preceq precnapprox precneqq precnsim precsim
    predisplaypenalty pretolerance prevdepth prime primfrac prod projlim propto protect psi pushtabs qopname qquad quad
    r rVert raise raisebox raisetag rangle rbrace rbrack rceil ref relax relbar relpenalty removelastskip restriction
    rfloor rgroup rhd rho rhook right rightarrow rightarrowfill rightarrowtail rightharpoondown rightharpoonup
    rightleftarrows rightleftharpoons rightline rightrightarrows rightskip rightsquigarrow rightthreetimes risingdotseq
    rlap rmdefault rmfamily rmoustache rmsubstdefault roman root rootbox rq rtimes rule rvert sb scriptfont
    scriptscriptfont scriptscriptstyle scriptspace scriptstyle scshape searrow sec selectfont setminus sf sffamily
    sfsubstdefault sharp shortmid shortparallel shoveleft shoveright sideset sigma sim simeq sin sinh skew skewchar skip
    slash slshape smallfrown smallint smallsetminus smallskip smallskipamount smallsmile smash smile sp space
    spacefactor spaceskip spadesuit sphericalangle split splitmaxdepth splittopskip sqcap sqcup sqrt sqrtsign sqsubset
    sqsubseteq sqsupset sqsupseteq square ss sscshape stackrel star stretch strut strutbox subarray subset subseteq
    subseteqq subsetneq subsetneqq substack succ succapprox succcurlyeq succeq succnapprox succneqq succnsim succsim sum
    sup supset supseteq supseteqq supsetneq supsetneqq surd swarrow swshape symAMSa symAMSb symbol symletters
    symoperators t tabbing tabbingsep tabcolsep tabskip tabular tag tan tanh tau tbinom text textacutedbl
    textascendercompwordmark textasciiacute textasciibreve textasciicaron textasciicircum textasciidieresis
    textasciigrave textasciimacron textasciitilde textasteriskcentered textbackslash textbaht textbar textbardbl textbf
    textbigcircle textblank textborn textbraceleft textbraceright textbrokenbar textbullet textcapitalcompwordmark
    textcelsius textcent textcentoldstyle textcircled textcircledP textcolonmonetary textcommaabove textcommabelow
    textcompsubstdefault textcompwordmark textcopyleft textcopyright textcurrency textdagger textdaggerdbl textdblhyphen
    textdblhyphenchar textdegree textdied textdiscount textdiv textdivorced textdollar textdollaroldstyle textdong
    textdownarrow texteightoldstyle textellipsis textemdash textendash textestimated texteuro textexclamdown
    textfiveoldstyle textflorin textfont textfouroldstyle textfractionsolidus textgravedbl textgreater textguarani
    textinterrobang textinterrobangdown textit textlangle textlbrackdbl textleaf textleftarrow
    textlegacyasteriskcentered textlegacybardbl textlegacybullet textlegacydagger textlegacydaggerdbl
    textlegacyparagraph textlegacyperiodcentered textlegacysection textless textlira textlnot textlquill textmarried
    textmd textmho textminus textmu textmusicalnote textnaira textnineoldstyle textnormal textnumero textohm textonehalf
    textoneoldstyle textonequarter textonesuperior textopenbullet textordfeminine textordmasculine textparagraph
    textperiodcentered textpertenthousand textperthousand textpeso textpilcrow textpm textquestiondown textquotedblleft
    textquotedblright textquoteleft textquoteright textquotesingle textquotestraightbase textquotestraightdblbase
    textrangle textrbrackdbl textrecipe textreferencemark textregistered textrightarrow textrm textrquill textsc
    textsection textservicemark textsevenoldstyle textsf textsixoldstyle textsl textssc textsterling textstyle
    textsuperscript textsurd textsw textthreeoldstyle textthreequarters textthreequartersemdash textthreesuperior
    texttildelow texttimes texttrademark texttt texttwelveudash texttwooldstyle texttwosuperior textulc textunderscore
    textup textuparrow textvisiblespace textwidth textwon textyen textzerooldstyle tfrac theequation thempfn
    thempfootnote thepage theparentequation therefore theta thetag thickapprox thickmuskip thicksim thickspace
    thinmuskip thinspace tilde times tmspace to toks tolerance top topskip triangle triangledown triangleleft
    trianglelefteq triangleq triangleright trianglerighteq trivlist ttfamily ttsubstdefault twoheadleftarrow
    twoheadrightarrow u ulcorner ulcshape underbar underbrace underleftarrow underleftrightarrow underline
    underrightarrow underset unitlength unlhd unpenalty unrhd unskip uparrow upbracefill updownarrow upharpoonleft
    upharpoonright uplus uppercase uproot upshape upsilon upuparrows urcorner usefont v vDash vadjust valign value
    varDelta varGamma varLambda varOmega varPhi varPi varPsi varSigma varTheta varUpsilon varXi varbigtriangledown
    varbigtriangleup varepsilon varinjlim varkappa varliminf varlimsup varnothing varphi varpi varprojlim varpropto
    varrho varsigma varsubsetneq varsubsetneqq varsupsetneq varsupsetneqq vartheta vartriangle vartriangleleft
    vartriangleright vbadness vbox vcenter vdash vdots vec vee veebar veqno vert vfil vfilneg vfuzz vglue vline vphantom
    vrule vskip vspace vss vtop wd wedge widehat widetilde widowpenalty wp wr xi xleaders xleftarrow xrightarrow yen
    zeta
`.split(/\s+/).filter(Boolean));

/** What mathCommandsOf records for a formula opening one of mathtools' environments: no control word has this name. */
const MATHTOOLS_ENVIRONMENT = 'begin-mathtools-environment';

/**
 * Packages with math commands a formula may use (so a document parsed from LaTeX that loaded them
 * keeps working), in the order they are loaded, each with its load options and its commands. `bm`
 * comes last, as its manual asks; `xcolor` and `graphicx` are loaded by the generator itself.
 */
const MATH_PACKAGES: ReadonlyArray<{ name: string; options?: string; commands: string }> = [
    {
        name: 'mathtools', commands: 'coloneqq Coloneqq coloneq Coloneq eqqcolon Eqqcolon eqcolon Eqcolon vcentcolon dblcolon approxcolon Approxcolon '
            + 'colonapprox Colonapprox simcolon Simcolon colonsim Colonsim ordinarycolon mathclap mathllap mathrlap mathmbox mathmakebox cramped '
            + 'crampedclap crampedllap crampedrlap smashoperator adjustlimits prescript underbracket overbracket xmapsto xhookrightarrow '
            + 'xhookleftarrow xLeftarrow xRightarrow xleftrightarrow xLeftrightarrow xleftharpoonup xrightharpoonup xleftharpoondown '
            + 'xrightharpoondown xleftrightharpoons xrightleftharpoons splitfrac splitdfrac MoveEqLeft ArrowBetweenLines shortintertext '
            + `lparen rparen Aboxed ${MATHTOOLS_ENVIRONMENT}`,
    },
    { name: 'esint', commands: 'oiint oiiint sqint sqiint ointclockwise ointctrclockwise varointclockwise varointctrclockwise fint landupint landdownint' },
    { name: 'stmaryrd', commands: 'llbracket rrbracket llparenthesis rrparenthesis mapsfrom Mapsfrom Mapsto longmapsfrom Longmapsfrom Longmapsto lightning' },
    { name: 'mathrsfs', commands: 'mathscr' },
    {
        name: 'upgreek', commands: 'upalpha upbeta upgamma updelta upepsilon upvarepsilon upzeta upeta uptheta upvartheta upiota upkappa uplambda upmu '
            + 'upnu upxi uppi upvarpi uprho upvarrho upsigma upvarsigma uptau upupsilon upphi upvarphi upchi uppsi upomega Upgamma Updelta Uptheta '
            + 'Uplambda Upxi Uppi Upsigma Upupsilon Upphi Uppsi Upomega',
    },
    { name: 'dsfont', commands: 'mathds' },
    { name: 'bbm', commands: 'mathbbm mathbbmss mathbbmtt' },
    { name: 'cancel', commands: 'cancel bcancel xcancel cancelto' },
    { name: 'xfrac', commands: 'sfrac' },
    { name: 'nicefrac', commands: 'nicefrac' },
    { name: 'relsize', commands: 'mathlarger mathsmaller' },
    { name: 'mathdots', commands: 'iddots' },
    { name: 'gensymb', commands: 'degree celsius perthousand ohm micro' },
    { name: 'braket', commands: 'bra ket braket Bra Ket Braket' },
    { name: 'siunitx', commands: 'SI si num ang unit qty numrange qtyrange SIrange numlist qtylist' },
    { name: 'mhchem', options: 'version=4', commands: 'ce pu' },
    { name: 'xcolor', commands: 'color textcolor colorbox fcolorbox' },
    { name: 'graphicx', commands: 'scalebox rotatebox reflectbox resizebox' },
    { name: 'bm', commands: 'bm bmmax hmmax' },
];

/**
 * siunitx's units and prefixes, which it defines inside its unit arguments (`\SI{3}{\metre}`): known
 * commands when a formula loads it.
 */
const SIUNITX_UNITS: ReadonlySet<string> = new Set(('ampere candela kelvin kilogram gram metre meter mole second becquerel degreeCelsius coulomb '
    + 'farad gray hertz henry joule katal lumen lux newton ohm pascal radian siemens sievert steradian tesla volt watt weber '
    + 'astronomicalunit bel dalton day decibel degree electronvolt hectare hour litre liter arcminute minute arcsecond neper tonne '
    + 'percent quecto ronto yocto zepto atto femto pico nano micro milli centi deci deca deka hecto kilo mega giga tera peta exa '
    + 'zetta yotta ronna quetta square squared cubic cubed per tothe raiseto of cancel highlight').split(' '));

/** The package of each command in {@link MATH_PACKAGES}. */
const MATH_PACKAGE_OF: Record<string, string> = lookupTable(Object.fromEntries(MATH_PACKAGES.flatMap(p => p.commands.split(' ').map(c => [c, p.name]))));

/**
 * Macros KaTeX and MathJax define that LaTeX does not (Markdown and HTML math is written for them),
 * and the LaTeX each stands for.
 */
const MATH_DIALECT_MACROS: Record<string, string> = lookupTable({
    R: '\\mathbb{R}', N: '\\mathbb{N}', Z: '\\mathbb{Z}', Q: '\\mathbb{Q}', C: '\\mathbb{C}', Reals: '\\mathbb{R}', reals: '\\mathbb{R}',
    natnums: '\\mathbb{N}', cnums: '\\mathbb{C}', Complex: '\\mathbb{C}', lang: '\\langle', rang: '\\rangle', larr: '\\leftarrow',
    rarr: '\\rightarrow', lArr: '\\Leftarrow', rArr: '\\Rightarrow', Larr: '\\Leftarrow', Rarr: '\\Rightarrow', harr: '\\leftrightarrow',
    hArr: '\\Leftrightarrow', Harr: '\\Leftrightarrow', uarr: '\\uparrow', darr: '\\downarrow', uArr: '\\Uparrow', dArr: '\\Downarrow',
    Uarr: '\\Uparrow', Darr: '\\Downarrow', isin: '\\in', plusmn: '\\pm', sdot: '\\cdot', sub: '\\subset', sube: '\\subseteq',
    supe: '\\supseteq', alef: '\\aleph', alefsym: '\\aleph', weierp: '\\wp', image: '\\Im', real: '\\Re', infin: '\\infty',
    clubs: '\\clubsuit', diamonds: '\\diamondsuit', hearts: '\\heartsuit', spades: '\\spadesuit', thetasym: '\\vartheta',
    Alpha: '\\mathrm{A}', Beta: '\\mathrm{B}', Epsilon: '\\mathrm{E}', Zeta: '\\mathrm{Z}', Eta: '\\mathrm{H}', Iota: '\\mathrm{I}',
    Kappa: '\\mathrm{K}', Mu: '\\mathrm{M}', Nu: '\\mathrm{N}', Omicron: '\\mathrm{O}', Rho: '\\mathrm{P}', Tau: '\\mathrm{T}',
    Chi: '\\mathrm{X}', omicron: 'o',
});

/**
 * Calls `visit` with each control word of a formula (its name, and where it starts and ends), read
 * as TeX reads them: `\\alpha` is a line break, then letters.
 */
function eachMathCommand(latex: string, visit: (name: string, start: number, end: number) => void): void {
    for (let i = latex.indexOf('\\'); i >= 0 && i < latex.length; i = latex.indexOf('\\', i)) {
        let j = i + 1;
        while (j < latex.length && /[A-Za-z]/.test(latex[j])) j++;
        if (j > i + 1) { visit(latex.slice(i + 1, j), i, j); i = j; } else i += 2;
    }
}

/** Adds the control words of a formula to `into`, and MATHTOOLS_ENVIRONMENT when it opens one of mathtools' environments. */
export function mathCommandsOf(latex: string, into: Set<string>): void {
    eachMathCommand(latex, (name, _start, end) => {
        into.add(name);
        if (name !== 'begin') return;
        const env = /^[ \t\n]*\{([A-Za-z]+\*?)\}/.exec(latex.slice(end, end + 48));
        if (env && MATHTOOLS_ENVIRONMENTS.has(env[1])) into.add(MATHTOOLS_ENVIRONMENT);
    });
}

/**
 * A formula with the KaTeX and MathJax macros LaTeX lacks written as what they stand for (`\R` as
 * `\mathbb{R}`), so it means in LaTeX what its source meant. Only whole control words are replaced.
 */
export function withLatexMathMacros(latex: string): string {
    const parts: string[] = [];
    let from = 0;
    eachMathCommand(latex, (name, start, end) => {
        const macro = MATH_DIALECT_MACROS[name];
        if (!macro) return;
        parts.push(latex.slice(from, start), macro);
        from = end;
    });
    if (from === 0) return latex;
    parts.push(latex.slice(from));
    return parts.join('');
}

/**
 * What the math a document holds needs beyond amsmath and amssymb: the packages its commands come
 * from (in load order, with their options), and the commands nothing the output loads defines, with
 * a definition for each that prints its own name, so the document compiles and shows where it is.
 */
export interface LatexMathPlan {
    packages: { name: string; options?: string }[];
    definitions: string[];
    undefinedCommands: string[];
}

/** Plans what the commands a document's formulas use need (see {@link LatexMathPlan}). */
export function planLatexMath(commands: Iterable<string>): LatexMathPlan {
    const names = [...new Set(commands)].sort();
    const needed = new Set<string>();
    for (const name of names) {
        const pkg = MATH_PACKAGE_OF[name];
        if (pkg) needed.add(pkg);
    }
    const undefinedCommands = names.filter(name => !MATH_PACKAGE_OF[name] && !LATEX_MATH_COMMANDS.has(name) && !(needed.has('siunitx') && SIUNITX_UNITS.has(name)));
    return {
        packages: MATH_PACKAGES.filter(p => needed.has(p.name)).map(p => ({ name: p.name, ...(p.options ? { options: p.options } : {}) })),
        definitions: undefinedCommands.map(name => `\\providecommand{\\${name}}{\\texttt{\\textbackslash ${name}}}`),
        undefinedCommands,
    };
}

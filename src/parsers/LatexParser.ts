import { AdmonitionMetadata, CellMetadata, CodeMetadata, CommentMetadata, FullOfficeParserConfig, HeadingMetadata, ImageMetadata, ListMetadata, NoteMetadata, OfficeAttachment, OfficeAuxiliaryContent, OfficeContentNode, OfficeMetadata, OfficeParserAST, OfficeWarningType, ParagraphMetadata, TextAlignment, TextFormatting, TextMetadata } from '../types.js';
import { attachmentLookup, repeatPreview, takeRepeats } from '../utils/repeatUtils.js';
import { trimEndChars } from '../utils/textUtils.js';
import { createAST } from '../utils/astUtils.js';
import { checkAbortSignal, logWarning } from '../utils/errorUtils.js';
import { isSourceComment } from '../utils/commentUtils.js';
import { createAttachment } from '../utils/imageUtils.js';
import { imageFromPdf, newDecodeBudget } from '../utils/textPdf.js';
import { LATEX_SYMBOL_CHARACTERS, LISTINGS_LANGUAGE_NAMES } from '../utils/latexUtils.js';
import { ADMONITION_COLOR } from '../utils/officeGenUtils.js';
import { ocrDuringParse } from '../utils/ocrUtils.js';
import { extractFiles } from '../utils/zipUtils.js';
import { setOwn } from '../utils/lookupUtils.js';

// ── limits ──────────────────────────────────────────────────────────────────────────────────────

/**
 * The input is a program: a handful of macro definitions can request exponential expansion
 * (`\def\a{\b\b}\def\b{\c\c}...`), and `\input` can include files in a cycle. These bound the
 * work any document can demand; past them expansion stops with `LATEX_EXPANSION_LIMIT_REACHED`.
 */
const MAX_MACRO_EXPANSIONS = 200000;
const MAX_EXPANDED_CHARS = 5000000;
const MAX_INPUT_FRAMES = 128;
const MAX_INCLUDE_DEPTH = 16;
/** Deepest nesting of groups and environments interpreted; deeper content is read as plain text. */
const MAX_NESTING_DEPTH = 128;
/** Deepest nesting of tables (a table inside a cell) interpreted; deeper tables are read as plain text. */
const MAX_TABLE_NESTING_DEPTH = 32;
/** Longest column specification interpreted, after `*{n}{...}` repetition. */
const MAX_COLUMN_SPEC_CHARS = 100000;
/** How often (in tokens) the parse loop checks the abort signal. */
const ABORT_CHECK_INTERVAL = 4096;

/** Extensions tried, in order, for an `\includegraphics` path given without one. */
const IMAGE_EXTENSIONS = ['', '.pdf', '.png', '.jpg', '.jpeg', '.eps', '.gif', '.svg', '.bmp', '.tif', '.tiff', '.webp'];

// ── character tables ────────────────────────────────────────────────────────────────────────────

/** Text-mode symbol commands and the characters they print. */
const SYMBOLS: Record<string, string> = {
    textbackslash: '\\', textasciitilde: '~', textasciicircum: '^', textless: '<', textgreater: '>', textbar: '|',
    textasciigrave: '`', textquotedbl: '"', textquotesingle: '\'', textbullet: '•', textperiodcentered: '·',
    textdegree: '°', textcopyright: '©', copyright: '©', textregistered: '®', texttrademark: '™',
    textellipsis: '…', ldots: '…', dots: '…', textendash: '–', textemdash: '—',
    textquoteleft: '‘', textquoteright: '’', textquotedblleft: '“', textquotedblright: '”',
    textexclamdown: '¡', textquestiondown: '¿', textsection: '§', S: '§', textparagraph: '¶',
    P: '¶', dag: '†', textdagger: '†', ddag: '‡', textdaggerdbl: '‡', pounds: '£',
    textsterling: '£', euro: '€', texteuro: '€', textyen: '¥', textcent: '¢',
    textonehalf: '½', textonequarter: '¼', textthreequarters: '¾', texttimes: '×', textdiv: '÷',
    textpm: '±', textmu: 'µ', textminus: '−', textrightarrow: '→', textleftarrow: '←',
    textuparrow: '↑', textdownarrow: '↓', guillemotleft: '«', guillemetleft: '«',
    guillemotright: '»', guillemetright: '»', guilsinglleft: '‹', guilsinglright: '›',
    quotedblbase: '„', quotesinglbase: '‚', ss: 'ß', SS: 'SS', ae: 'æ', AE: 'Æ', oe: 'œ',
    OE: 'Œ', o: 'ø', O: 'Ø', aa: 'å', AA: 'Å', l: 'ł', L: 'Ł', i: 'ı', j: 'ȷ',
    DH: 'Ð', dh: 'ð', TH: 'Þ', th: 'þ', NG: 'Ŋ', ng: 'ŋ', checkmark: '✓',
    LaTeX: 'LaTeX', TeX: 'TeX', LaTeXe: 'LaTeX2ε', XeTeX: 'XeTeX', XeLaTeX: 'XeLaTeX', LuaTeX: 'LuaTeX',
    LuaLaTeX: 'LuaLaTeX', BibTeX: 'BibTeX', textunderscore: '_', textbraceleft: '{', textbraceright: '}',
    textdollar: '$', textnumero: '№', textordfeminine: 'ª', textordmasculine: 'º', textbrokenbar: '¦',
    textcurrency: '¤', textlnot: '¬', textperthousand: '‰', textreferencemark: '※',
    textinterrobang: '‽', textmusicalnote: '♪', textvisiblespace: '␣', textopenbullet: '◦',
    textbigcircle: '◯', textlangle: '〈', textrangle: '〉', slash: '/', textfraction: '⁄',
    '#': '#', '$': '$', '%': '%', '&': '&', '_': '_', '{': '{', '}': '}', ' ': ' ', '-': '­', ',': ' ',
    ';': ' ', ':': ' ', '!': '', '/': '', '@': '', quad: ' ', qquad: '  ', enspace: ' ',
    thinspace: ' ', enskip: ' ', space: ' ', nobreakspace: ' ',
};

/** Accent commands and the combining mark each adds to its argument. */
const ACCENTS: Record<string, string> = {
    '\'': '́', '`': '̀', '^': '̂', '"': '̈', '~': '̃', '=': '̄', '.': '̇',
    u: '̆', v: '̌', H: '̋', c: '̧', k: '̨', r: '̊', d: '̣', b: '̱', t: '͡',
};

/** Named xcolor colours usable without a definition. */
const NAMED_COLORS: Record<string, string> = {
    red: 'FF0000', green: '00FF00', blue: '0000FF', cyan: '00FFFF', magenta: 'FF00FF', yellow: 'FFFF00', black: '000000',
    white: 'FFFFFF', gray: '808080', darkgray: '404040', lightgray: 'BFBFBF', brown: 'BF8040', lime: 'BFFF00',
    olive: '808000', orange: 'FF8000', pink: 'FFBFBF', purple: 'BF0040', teal: '008080', violet: '800080',
};

/** Size switches, in points, per class body size (the standard classes' size tables). */
const SIZE_TABLE: Record<number, Record<string, number>> = {
    10: { tiny: 5, scriptsize: 7, footnotesize: 8, small: 9, normalsize: 10, large: 12, Large: 14.4, LARGE: 17.28, huge: 20.74, Huge: 24.88 },
    11: { tiny: 6, scriptsize: 8, footnotesize: 9, small: 10, normalsize: 10.95, large: 12, Large: 14.4, LARGE: 17.28, huge: 20.74, Huge: 24.88 },
    12: { tiny: 6, scriptsize: 8, footnotesize: 10, small: 10.95, normalsize: 12, large: 14.4, Large: 17.28, LARGE: 20.74, huge: 24.88, Huge: 24.88 },
};

/**
 * Commands a preamble may redefine for typesetting (the LaTeX generator itself redefines
 * `\paragraph`) whose meaning for the document's structure does not change: a redefinition of
 * one of these is not expanded in place of the built-in reading.
 */
const PROTECTED_COMMANDS = new Set(['part', 'chapter', 'section', 'subsection', 'subsubsection', 'paragraph', 'subparagraph',
    'item', 'footnote', 'endnote', 'footnotemark', 'footnotetext', 'caption', 'label', 'ref', 'cite', 'href', 'url', 'hyperref',
    'textbf', 'textit', 'emph', 'underline', 'uline', 'sout', 'texttt', 'textsuperscript', 'textsubscript', 'textcolor', 'colorbox',
    'includegraphics', 'maketitle', 'title', 'author', 'date', 'thanks', 'frametitle', 'note', 'tableofcontents', 'LaTeX', 'TeX']);

/** Sectioning commands in rank order; the top rank present in the document becomes heading level 1. */
const SECTIONING = ['part', 'chapter', 'section', 'subsection', 'subsubsection', 'paragraph', 'subparagraph'];

/** Display-math environments, read as block math. Numbered or multi-line ones keep their environment in the LaTeX. */
const DISPLAY_MATH_ENVS = new Set(['equation', 'equation*', 'align', 'align*', 'gather', 'gather*', 'multline', 'multline*',
    'flalign', 'flalign*', 'alignat', 'alignat*', 'eqnarray', 'eqnarray*', 'displaymath', 'math', 'dmath', 'dmath*']);
/** Of those, the ones that are just `\[...\]`: their body alone is the formula. */
const PLAIN_DISPLAY_MATH_ENVS = new Set(['equation*', 'displaymath', 'dmath*']);

/** Environments whose body is literal text. */
const VERBATIM_ENVS = new Set(['verbatim', 'verbatim*', 'Verbatim', 'Verbatim*', 'BVerbatim', 'LVerbatim', 'lstlisting', 'minted', 'alltt', 'comment']);

/** Table environments and the arguments before their column specification. */
const TABLE_ENVS: Record<string, string> = {
    tabular: 'o', 'tabular*': 'mo', tabularx: 'm', tabulary: 'm', longtable: 'o', xltabular: 'm', supertabular: '',
    'longtable*': 'o', tblr: 'o', longtblr: 'o', array: 'o',
};

/** Environments that only draw pictures: their content is skipped rather than read as text. */
const DRAWING_ENVS = new Set(['tikzpicture', 'pgfpicture', 'picture', 'forest', 'axis', 'circuitikz', 'pspicture', 'tikzcd']);

/**
 * `src` without its `filecontents` bodies, which are file data rather than document source. Scanned
 * forward once: a pattern looking for each start's end read the rest of the source again for every
 * unclosed one (a megabyte of them took seconds). An unclosed body is left in place, as the pattern did.
 */
function withoutFilecontents(src: string): string {
    const begin = /\\begin\s*\{filecontents\*?\}/g;
    const end = /\\end\s*\{filecontents\*?\}/g;
    let out = '', from = 0;
    for (let open = begin.exec(src); open; open = begin.exec(src)) {
        end.lastIndex = open.index + open[0].length;
        const close = end.exec(src);
        if (!close) break;
        out += src.slice(from, open.index);
        from = begin.lastIndex = close.index + close[0].length;
    }
    return out + src.slice(from);
}

/**
 * Environments with no meaning of their own for the AST (layout wrappers): their content is read
 * in place. The value is the argument signature consumed after `\begin{name}` (see {@link readArgs}).
 */
const TRANSPARENT_ENVS: Record<string, string> = {
    minipage: 'oomm', columns: 'o', column: 'om', multicols: 'mo', 'multicols*': 'mo', adjustbox: 'm', landscape: '',
    spacing: 'm', singlespace: '', onehalfspace: '', doublespace: '', small: '', footnotesize: '', scriptsize: '',
    large: '', Large: '', sloppypar: '', frontmatter: '', mainmatter: '', appendices: '', subequations: '',
    samepage: '', onlyenv: 'O', overprint: 'o', uncoverenv: 'O',
    visibleenv: 'O', actionenv: 'O', titlepage: '', subfigure: 'om', subtable: 'om', framed: '', mdframed: 'o', tcolorbox: 'o',
    hyphenrules: 'm', CJK: 'mm', 'CJK*': 'mm', example: 'o', remark: 'o',
};

/**
 * Conditionals with a fixed value: `\iftrue`, `\iffalse`, and the engine tests of `iftex` (and the
 * older `ifxetex`/`ifluatex`/`ifpdf` packages). The parser reads a document as pdfLaTeX, the usual
 * engine (Overleaf's and arXiv's default), would compile it, so a document written for several
 * engines yields one branch's text.
 */
const FIXED_CONDITIONALS: Record<string, boolean> = {
    ifPDFTeX: true, ifpdftex: true, ifpdf: true, ifeTeX: true, ifetex: true, iftrue: true, iffalse: false,
    ifXeTeX: false, ifxetex: false, ifLuaTeX: false, ifluatex: false, ifLuaHBTeX: false, ifluahbtex: false, iftutex: false,
    ifpTeX: false, ifptex: false, ifupTeX: false, ifuptex: false, ifVTeX: false, ifvtex: false, ifAlephTeX: false, ifalephtex: false,
};

/** TeX's other conditionals: decided only in the forms {@link LatexReader.conditional} knows, but always matched with their `\fi`. */
const PRIMITIVE_CONDITIONALS = new Set(['if', 'ifcat', 'ifnum', 'ifdim', 'ifodd', 'ifvmode', 'ifhmode', 'ifmmode', 'ifinner', 'ifvoid',
    'ifhbox', 'ifvbox', 'ifx', 'ifeof', 'ifcase', 'ifdefined', 'ifcsname', 'iffontchar', 'ifincsname', 'ifabsnum', 'ifabsdim',
    'ifpdfprimitive', 'ifprimitive']);

/** Commands undefined under pdfLaTeX that documents test for: other engines' primitives and converters' markers. */
const UNDEFINED_UNDER_PDFLATEX = new Set(['directlua', 'luatexversion', 'luaescapestring', 'XeTeXversion', 'XeTeXrevision',
    'XeTeXinterchartokenstate', 'kanjiskip', 'ptexversion', 'uptexversion', 'HCode', 'Configure', 'NoValue']);

/**
 * `\if...` commands that take arguments instead of ending in `\fi` (etoolbox, ifthen, babel), and the
 * math symbol `\iff`: never counted as conditionals. Any other `\if...` word not followed by a brace is
 * taken for a conditional a package defined (such as `\ifdraft`), so its `\else` and `\fi` stay its own.
 */
const ARGUMENT_TESTS = new Set(['iff', 'ifthenelse', 'iflanguage', 'ifdef', 'ifundef', 'ifcsdef', 'ifcsundef', 'ifdefmacro',
    'ifcsmacro', 'ifdefparam', 'ifcsparam', 'ifdefprefix', 'ifcsprefix', 'ifdefprotected', 'ifcsprotected', 'ifdefltxprotect',
    'ifcsltxprotect', 'ifdefempty', 'ifcsempty', 'ifdefvoid', 'ifcsvoid', 'ifdefequal', 'ifcsequal', 'ifdefstring', 'ifcsstring',
    'ifdefstrequal', 'ifcsstrequal', 'ifdefcounter', 'ifcscounter', 'ifltxcounter', 'ifdeflength', 'ifcslength', 'ifdefdimen',
    'ifcsdimen', 'ifstrequal', 'ifstrempty', 'ifblank', 'ifnumcomp', 'ifnumequal', 'ifnumgreater', 'ifnumless', 'ifnumodd',
    'ifdimcomp', 'ifdimequal', 'ifdimgreater', 'ifdimless', 'ifboolexpr', 'ifboolexpe', 'ifbool', 'iftoggle', 'ifinlist',
    'ifinlistcs', 'ifrmnum', 'ifpatchable']);

/** pdfTeX's integer registers that documents test, with their values under pdfLaTeX. */
const PDFTEX_REGISTERS: Record<string, number> = { pdfoutput: 1, pdftexversion: 140, pdfshellescape: 0 };

/** Document classes with chapters (where `\chapter` is defined). */
const CHAPTER_CLASSES = new Set(['book', 'report', 'memoir', 'scrbook', 'scrreprt', 'amsbook', 'extbook', 'extreport', 'ctexbook',
    'ctexrep', 'jsbook', 'jbook', 'jreport', 'ujbook', 'ujreport', 'tbook', 'treport', 'thesis', 'ociamthesis', 'Dissertate']);

/** Theorem-like environments' styles in the classes that provide them without `\newtheorem`. */
const DEFAULT_THEOREM_STYLES: Record<string, string> = {
    definition: 'definition', example: 'definition', exercise: 'definition', problem: 'definition', question: 'definition',
    solution: 'definition', remark: 'remark', claim: 'remark', note: 'remark', case: 'remark', fact: 'plain',
    definitions: 'definition', examples: 'definition',
};

/**
 * babel/polyglossia language names and their BCP 47 tags: a language switch keeps only its text,
 * and the document's main language is recorded in `metadata.language`.
 */
const LANGUAGE_CODES: Record<string, string> = {
    english: 'en', american: 'en-US', USenglish: 'en-US', british: 'en-GB', UKenglish: 'en-GB', canadian: 'en-CA',
    australian: 'en-AU', newzealand: 'en-NZ', french: 'fr', francais: 'fr', acadian: 'fr-CA', canadien: 'fr-CA',
    german: 'de', ngerman: 'de', austrian: 'de-AT', naustrian: 'de-AT', swissgerman: 'de-CH', nswissgerman: 'de-CH',
    spanish: 'es', mexican: 'es-MX', italian: 'it', portuguese: 'pt', portuges: 'pt', brazil: 'pt-BR', brazilian: 'pt-BR',
    dutch: 'nl', afrikaans: 'af', russian: 'ru', ukrainian: 'uk', belarusian: 'be', polish: 'pl', czech: 'cs',
    slovak: 'sk', slovene: 'sl', slovenian: 'sl', croatian: 'hr', serbian: 'sr', serbianc: 'sr', bulgarian: 'bg',
    macedonian: 'mk', greek: 'el', polutonikogreek: 'el', swedish: 'sv', danish: 'da', norsk: 'nb', norwegian: 'nb',
    bokmal: 'nb', nynorsk: 'nn', finnish: 'fi', icelandic: 'is', estonian: 'et', latvian: 'lv', lithuanian: 'lt',
    hungarian: 'hu', magyar: 'hu', romanian: 'ro', turkish: 'tr', catalan: 'ca', basque: 'eu', galician: 'gl',
    irish: 'ga', scottish: 'gd', welsh: 'cy', breton: 'br', latin: 'la', hebrew: 'he', arabic: 'ar', persian: 'fa',
    farsi: 'fa', urdu: 'ur', hindi: 'hi', bengali: 'bn', marathi: 'mr', sanskrit: 'sa', tamil: 'ta', telugu: 'te',
    thai: 'th', vietnamese: 'vi', indonesian: 'id', bahasa: 'id', bahasai: 'id', malay: 'ms', bahasam: 'ms',
    japanese: 'ja', chinese: 'zh', korean: 'ko', esperanto: 'eo', albanian: 'sq', armenian: 'hy', georgian: 'ka',
};

/** Titles of the theorem-like environments classes provide without `\newtheorem`. */
const DEFAULT_THEOREM_TITLES: Record<string, string> = {
    theorem: 'Theorem', lemma: 'Lemma', corollary: 'Corollary', proposition: 'Proposition', definition: 'Definition',
    example: 'Example', remark: 'Remark', conjecture: 'Conjecture', claim: 'Claim', exercise: 'Exercise',
    problem: 'Problem', property: 'Property', question: 'Question', solution: 'Solution', note: 'Note', case: 'Case',
    fact: 'Fact', definitions: 'Definitions', examples: 'Examples',
};

/** Springer's classes, which define every theorem-like environment above, each numbered with its own counter. */
const SPRINGER_CLASSES = new Set(['llncs', 'svjour', 'svjour2', 'svjour3', 'svmono', 'svmult', 'svproc']);
/** The theorem-like environments beamer defines (as blocks, unnumbered). */
const BEAMER_THEOREMS = new Set(['theorem', 'corollary', 'fact', 'lemma', 'problem', 'solution', 'definition', 'definitions', 'example', 'examples']);
/**
 * Names that are theorem-like in any document, whose definition may be out of sight (in a package or
 * a class the parser does not read): headed by their title, unnumbered, since their numbering is unknown.
 */
const GENERIC_THEOREMS = new Set(['theorem', 'lemma', 'corollary', 'proposition', 'definition', 'conjecture']);

/** LaTeX for the special characters of a verbatim (`v`) argument, which print as themselves. */
const VERBATIM_CHARS: Record<string, string> = {
    '\\': '\\textbackslash{}', '{': '\\{', '}': '\\}', $: '\\$', '&': '\\&', '#': '\\#', '^': '\\textasciicircum{}', _: '\\_',
    '%': '\\%', '~': '\\textasciitilde{}',
};

/** What an optional `xparse` argument that was not given holds (and prints), as in LaTeX. */
const NO_VALUE = '-NoValue-';

/**
 * Legacy `inputenc` encodings and the WHATWG encoding each decodes with. A `.tex` file is UTF-8
 * unless it declares one of these (or is not valid UTF-8, when Windows-1252 is assumed).
 */
const INPUTENC_ENCODINGS: Record<string, string> = {
    utf8: 'utf-8', latin1: 'windows-1252', latin9: 'iso-8859-15', latin2: 'iso-8859-2', latin3: 'iso-8859-3',
    latin4: 'iso-8859-4', latin5: 'windows-1254', latin10: 'iso-8859-16', cp1250: 'windows-1250', cp1251: 'windows-1251',
    cp1252: 'windows-1252', cp1257: 'windows-1257', ansinew: 'windows-1252', applemac: 'macintosh', koi8r: 'koi8-r',
    'koi8-r': 'koi8-r', 'x-mac-cyrillic': 'x-mac-cyrillic', cp866: 'ibm866',
};

/** Commands that carry no content for the AST, with the argument signature to consume. */
const IGNORED_COMMANDS: Record<string, string> = {
    usepackage: 'om', RequirePackage: 'om', documentclass: 'om', vspace: 'sm', vskip: '', hskip: '', vfill: '', hfill: '',
    smallskip: '', medskip: '', bigskip: '', noindent: '', indent: '', nopagebreak: 'o', nolinebreak: 'o',
    tableofcontents: '', listoffigures: '', listoftables: '', frontmatter: '', mainmatter: '', backmatter: '',
    appendix: '', pagestyle: 'm', thispagestyle: 'm', pagenumbering: 'm', setlength: 'mm', addtolength: 'mm',
    settowidth: 'mm', newlength: 'm', addtocounter: 'mm', stepcounter: 'm', refstepcounter: 'm', newcounter: 'mo',
    selectfont: '', usefont: 'mmmm', fontfamily: 'm', fontseries: 'm', fontshape: 'm', linespread: 'm', strut: '',
    relax: '', protect: '', leavevmode: '', null: '', ignorespaces: '', unskip: '', vphantom: 'm', index: 'm',
    glossary: 'm', nocite: 'm', bibliographystyle: 'm', bibliography: 'm', printbibliography: 'o', addbibresource: 'om',
    hypersetup: 'm', lstset: 'm', setbeamertemplate: 'mo', setbeamercolor: 'mm', setbeamerfont: 'mm', usetheme: 'om',
    usecolortheme: 'om', usefonttheme: 'om', useinnertheme: 'om', useoutertheme: 'om', setbeameroption: 'm', pause: 'o',
    titlepage: '', graphicspath: 'm', DeclareGraphicsExtensions: 'm', geometry: 'm', newgeometry: 'm', restoregeometry: '',
    makeatletter: '', makeatother: '', newunicodechar: 'mm', DeclareUnicodeCharacter: 'mm', setmainfont: 'om',
    setsansfont: 'om', setmonofont: 'om', setmathfont: 'om', newfontfamily: 'mom', babelprovide: 'om', selectlanguage: 'm',
    theendnotes: '', printindex: '', makeindex: '', frenchspacing: '', nonfrenchspacing: '', raggedbottom: '',
    flushbottom: '', onecolumn: '', twocolumn: 'o', centerline: '', hline: '', cline: 'm', toprule: 'o', midrule: 'o',
    bottomrule: 'o', addlinespace: 'o', endhead: '', endfirsthead: '', endfoot: '', endlastfoot: '', fancyhf: 'm',
    renewcommand: '', footnotesize: '', thanks: 'm', setcounter: 'mm', normalsize: '', ding: 'm', captionsetup: 'om',
    AtBeginDocument: 'm', AtEndDocument: 'm', listfiles: '', hyphenation: 'm',
    enlargethispage: 'sm', newpage: '', color: '', label: '', column: 'om', textwidth: '', linewidth: '',
    columnwidth: '', paperwidth: '', textheight: '', paperheight: '', baselineskip: '', parindent_: '', tabcolsep: '',
    arraybackslash: '', arrayrulewidth: '', dimexpr: '', fboxsep: '', setbeamersize: 'm', logo: 'm', institute: 'om',
    documentstyle: 'om', NeedsTeXFormat: 'mo', ProvidesPackage: 'mo', ProvidesClass: 'mo', ProvidesFile: 'mo',
    PassOptionsToPackage: 'mm', PassOptionsToClass: 'mm', setCJKmainfont: 'omo', setCJKsansfont: 'omo', setCJKmonofont: 'omo',
    setCJKfamilyfont: 'momo', setCJKfallbackfamilyfont: 'mom', xeCJKsetup: 'm', xeCJKDeclareCharClass: 'mm', setmainjfont: 'omo',
    setsansjfont: 'omo', setmonojfont: 'omo', ltjsetparameter: 'm', babelfont: 'omom', CJKfamily: 'm',
    setotherlanguage: 'om', setotherlanguages: 'm', newtheoremstyle: 'mmmmmmmmm', IEEEpeerreviewmaketitle: '',
    IEEEoverridecommandlockouts: '', IEEEdisplaynontitleabstractindextext: '', IEEEaftertitletext: 'm', ExplSyntaxOff: '',
    BooleanTrue: '', BooleanFalse: '', swapnumbers: '',
};

/**
 * TeX's and LaTeX's length and integer parameters: an assignment to one (`\parskip=6pt`,
 * plain TeX's `\magnification=1200`) is not text.
 */
const TEX_REGISTERS = new Set(['magnification', 'hsize', 'vsize', 'hoffset', 'voffset', 'baselineskip', 'lineskip', 'lineskiplimit',
    'topskip', 'maxdepth', 'tolerance', 'pretolerance', 'looseness', 'emergencystretch', 'hbadness', 'vbadness', 'hfuzz', 'vfuzz',
    'clubpenalty', 'widowpenalty', 'displaywidowpenalty', 'brokenpenalty', 'interlinepenalty', 'exhyphenpenalty', 'hyphenpenalty',
    'abovedisplayskip', 'belowdisplayskip', 'abovedisplayshortskip', 'belowdisplayshortskip', 'textwidth', 'textheight',
    'linewidth', 'columnwidth', 'paperwidth', 'paperheight', 'oddsidemargin', 'evensidemargin', 'topmargin', 'headheight',
    'headsep', 'footskip', 'marginparwidth', 'marginparsep', 'columnsep', 'tabcolsep', 'arrayrulewidth', 'fboxsep', 'fboxrule',
    'overfullrule', 'parfillskip', 'spaceskip', 'xspaceskip', 'mathsurround', 'nulldelimiterspace',
    'scriptspace', 'unitlength', 'itemsep', 'parsep', 'topsep', 'partopsep', 'labelsep', 'labelwidth', 'leftmargin',
    'rightmargin', 'itemindent', 'listparindent', 'footnotesep', 'skip', 'dimen', 'count', 'toks']);

/** Commands whose last mandatory argument is ordinary content, with the signature before it. */
const WRAPPER_COMMANDS: Record<string, string> = {
    mbox: '', makebox: 'oo', fbox: '', framebox: 'oo', parbox: 'oom', raisebox: 'moo', scalebox: 'mo', resizebox: 'sm',
    rotatebox: 'om', centerline: '', textnormal: '', textup: '', textmd: '', textsc: '', textulc: '', textbf_: '',
    boxed: '', text: '', only: 'O', uncover: 'O', visible: 'O', invisible: 'O', alt: 'Om', temporal: 'Omm', onslide: 'O',
    structure: 'O', alert: 'O', emph_: '', newblock: '', tcbox: 'o', adjustbox: 'm', shortstack: 'o', underbrace: '',
    foreignlanguage: 'om', textlang: 'om', IEEEauthorblockN: '', IEEEauthorblockA: '', IEEEmembership: '',
    IEEEtitleabstractindextext: '', dedicatory: '', line: '', leftline: '', rightline: '',
};

/** Whether a table has its own entry `key` (not one inherited from `Object.prototype`, such as `constructor`). */
const own = (table: object, key: string): boolean => Object.prototype.hasOwnProperty.call(table, key);

// ── scanner ─────────────────────────────────────────────────────────────────────────────────────

interface Frame { s: string; i: number; file?: string; }

type Tok =
    | { t: 'cs'; name: string }
    | { t: 'text'; v: string }
    | { t: 'space' }
    | { t: 'par' }
    | { t: 'bgroup' }
    | { t: 'egroup' }
    | { t: 'dollar' }
    | { t: 'amp' }
    | { t: 'comment'; v: string }
    | { t: 'eof' };

interface Snapshot { frames: Frame[]; idx: number[]; afterCs: boolean; lineStart: boolean; }

/**
 * Reads LaTeX source the way TeX tokenizes it: control words (a backslash and letters, with the
 * spaces and single line end after them skipped), control symbols, groups, math shifts, comments,
 * and whitespace, where a blank line is a paragraph break. Source is a stack of frames, so macro
 * expansions and included files are read in place, and a raw argument can span frames.
 */
class Scanner {
    private frames: Frame[];
    private afterCs = false;
    private lineStart = true;

    constructor(src: string, file?: string) {
        this.frames = [{ s: src, i: 0, file }];
    }

    get depth(): number { return this.frames.length; }

    /** Files currently open on the frame stack (for include-cycle detection). */
    openFiles(): string[] { return this.frames.map(f => f.file).filter((f): f is string => !!f); }

    /** The current reading position (top frame and offset), for remembering where a look-ahead ended. */
    position(): { frame: Frame; i: number } | undefined {
        const f = this.frames[this.frames.length - 1];
        return f ? { frame: f, i: f.i } : undefined;
    }

    push(src: string, file?: string): void {
        this.frames.push({ s: src, i: 0, file });
    }

    save(): Snapshot {
        return { frames: this.frames.slice(), idx: this.frames.map(f => f.i), afterCs: this.afterCs, lineStart: this.lineStart };
    }

    restore(snap: Snapshot): void {
        this.frames = snap.frames.slice();
        snap.frames.forEach((f, k) => { f.i = snap.idx[k]; });
        this.afterCs = snap.afterCs;
        this.lineStart = snap.lineStart;
    }

    peekCh(): string | undefined {
        for (;;) {
            const f = this.frames[this.frames.length - 1];
            if (!f) return undefined;
            if (f.i < f.s.length) return f.s[f.i];
            if (this.frames.length === 1) return undefined;
            this.frames.pop();
        }
    }

    nextCh(): string | undefined {
        const c = this.peekCh();
        if (c !== undefined) this.frames[this.frames.length - 1].i++;
        return c;
    }

    private isLetter(c: string | undefined, atLetter: boolean): boolean {
        return c !== undefined && (/[A-Za-z]/.test(c) || (atLetter && c === '@'));
    }

    /**
     * The letters of a control word, taken from the frame it starts in only: a macro body or an
     * included file ends a name, as TeX tokenizes them apart (`\def\x{\textbf}\x b` is `\textbf` then `b`).
     */
    private readName(atLetter: boolean): string {
        this.peekCh();
        const f = this.frames[this.frames.length - 1];
        let name = '';
        while (f && f.i < f.s.length && this.isLetter(f.s[f.i], atLetter)) name += f.s[f.i++];
        return name;
    }

    next(atLetter = false): Tok {
        for (;;) {
            const c = this.nextCh();
            if (c === undefined) return { t: 'eof' };
            if (c === '\\') {
                const n = this.peekCh();
                if (n === undefined) { this.afterCs = false; return { t: 'text', v: '\\' }; }
                if (this.isLetter(n, atLetter)) {
                    const name = this.readName(atLetter);
                    while (this.peekCh() === ' ' || this.peekCh() === '\t') this.nextCh();
                    this.afterCs = true;
                    this.lineStart = false;
                    return { t: 'cs', name };
                }
                this.nextCh();
                this.afterCs = n === ' ';
                this.lineStart = false;
                return { t: 'cs', name: n };
            }
            if (c === '%') {
                let v = '';
                while (this.peekCh() !== undefined && this.peekCh() !== '\n') v += this.nextCh();
                this.nextCh();
                while (this.peekCh() === ' ' || this.peekCh() === '\t') this.nextCh();
                this.lineStart = true;
                return { t: 'comment', v };
            }
            if (c === ' ' || c === '\t' || c === '\n' || c === '\r') {
                let newlines = c === '\n' ? 1 : 0;
                while (this.peekCh() === ' ' || this.peekCh() === '\t' || this.peekCh() === '\n' || this.peekCh() === '\r') {
                    if (this.nextCh() === '\n') newlines++;
                }
                const par = newlines >= 2 || (this.lineStart && newlines >= 1);
                const skip = this.afterCs && !par;
                this.afterCs = false;
                this.lineStart = false;
                if (skip) continue;
                return par ? { t: 'par' } : { t: 'space' };
            }
            this.afterCs = false;
            this.lineStart = false;
            switch (c) {
                case '{': return { t: 'bgroup' };
                case '}': return { t: 'egroup' };
                case '$': return { t: 'dollar' };
                case '&': return { t: 'amp' };
                case '~': return { t: 'text', v: ' ' };
                default: return { t: 'text', v: c };
            }
        }
    }

    /** Skips whitespace and comments (not paragraph breaks' meaning: used only before arguments). */
    skipBlanks(): void {
        for (;;) {
            const c = this.peekCh();
            if (c === ' ' || c === '\t' || c === '\n' || c === '\r') { this.nextCh(); continue; }
            if (c === '%') { while (this.peekCh() !== undefined && this.peekCh() !== '\n') this.nextCh(); continue; }
            return;
        }
    }

    /**
     * The raw source of the next argument: a brace group's contents, or else the next single
     * token (TeX's undelimited argument). Comments inside are kept, since they may carry comment
     * annotations the parser reads.
     */
    /** A raw read ends any "just after a control word" state: a space after `}` is a real space. */
    private endRaw<T>(v: T): T {
        this.afterCs = false;
        this.lineStart = false;
        return v;
    }

    readRawGroup(): string | null {
        this.skipBlanks();
        const c = this.peekCh();
        if (c === undefined) return null;
        if (c !== '{') {
            if (c === '}' ) return null;
            this.nextCh();
            if (c === '\\') {
                const name = /[A-Za-z]/.test(this.peekCh() ?? '') ? this.readName(true) : this.nextCh() ?? '';
                return this.endRaw('\\' + name);
            }
            return this.endRaw(c);
        }
        this.nextCh();
        return this.endRaw(this.readBalanced('{', '}'));
    }

    private readBalanced(open: string, close: string): string {
        let depth = 1;
        let out = '';
        for (;;) {
            const c = this.nextCh();
            if (c === undefined) return out;
            if (c === '\\') {
                const n = this.nextCh();
                out += c + (n ?? '');
                continue;
            }
            if (c === '%') {
                out += c;
                while (this.peekCh() !== undefined && this.peekCh() !== '\n') out += this.nextCh();
                continue;
            }
            if (c === open) depth++;
            else if (c === close && --depth === 0) return out;
            if (open !== '{' && c === '{') {
                out += c + this.readBalanced('{', '}') + '}';
                continue;
            }
            out += c;
        }
    }

    /** `[...]` at the current position (after blanks), or null without consuming anything. */
    readRawOptional(open = '[', close = ']'): string | null {
        const snap = this.save();
        this.skipBlanks();
        if (this.peekCh() !== open) { this.restore(snap); return null; }
        this.nextCh();
        return this.endRaw(this.readBalanced(open, close));
    }

    readStar(): boolean {
        return this.readChar('*');
    }

    /** Consumes `ch` if it is the next non-blank character (an `xparse` `t` argument, a star). */
    readChar(ch: string): boolean {
        const snap = this.save();
        this.skipBlanks();
        if (this.peekCh() === ch) { this.nextCh(); return true; }
        this.restore(snap);
        return false;
    }

    /**
     * Skips a conditional branch that is not taken: raw source up to the `\else` (with `toElse`)
     * or `\fi` that belongs to it, counting the conditionals opened inside it. Consumes that word
     * and returns it, or null when the input ends first.
     */
    skipBranch(isConditional: (name: string) => boolean, toElse: boolean): 'else' | 'fi' | null {
        let depth = 0;
        for (;;) {
            const c = this.nextCh();
            if (c === undefined) return null;
            if (c === '%') { while (this.peekCh() !== undefined && this.peekCh() !== '\n') this.nextCh(); continue; }
            if (c !== '\\') continue;
            const name = this.readName(true);
            if (!name) { this.nextCh(); continue; }
            if (name === 'fi') {
                if (depth === 0) return this.endWord('fi');
                depth--;
            } else if (name === 'else') {
                if (depth === 0 && toElse) return this.endWord('else');
            } else if (isConditional(name)) depth++;
        }
    }

    /** Ends a control word read by hand: the blanks after it are skipped, as TeX does. */
    private endWord<T>(v: T): T {
        while (this.peekCh() === ' ' || this.peekCh() === '\t') this.nextCh();
        this.afterCs = true;
        this.lineStart = false;
        return v;
    }

    /** Whether the next non-blank character is `ch` (nothing consumed). */
    nextIs(ch: string): boolean {
        const snap = this.save();
        this.skipBlanks();
        const r = this.peekCh() === ch;
        this.restore(snap);
        return r;
    }

    /** Raw source up to (and consuming) `marker`; the marker is not included. */
    readRawUntil(marker: string): { text: string; found: boolean } {
        // The body is collected in parts and only the last `marker.length` characters are compared,
        // so a long `\verb` or `\iffalse` body costs time linear in its length.
        const parts: string[] = [];
        let tail = '';
        for (;;) {
            const c = this.nextCh();
            if (c === undefined) return { text: parts.join(''), found: false };
            parts.push(c);
            tail = (tail + c).slice(-marker.length);
            if (tail === marker) return this.endRaw({ text: parts.join('').slice(0, -marker.length), found: true });
        }
    }

    /** Raw source up to the `\end{name}` matching an already-consumed `\begin{name}`. */
    readRawEnvBody(name: string): string {
        const begin = `\\begin{${name}}`;
        const end = `\\end{${name}}`;
        // As in readRawUntil: parts plus a short rolling tail, so a brace-heavy body stays linear.
        const width = Math.max(begin.length, end.length) + 1;
        let depth = 1;
        const parts: string[] = [];
        let tail = '';
        for (;;) {
            const c = this.nextCh();
            if (c === undefined) return parts.join('');
            parts.push(c);
            tail = (tail + c).slice(-width);
            if (c === '}') {
                if (tail.endsWith(end) && !tail.endsWith('\\' + end)) {
                    if (--depth === 0) return this.endRaw(parts.join('').slice(0, -end.length));
                } else if (tail.endsWith(begin)) depth++;
            }
        }
    }

    /** Raw math up to an unescaped closing delimiter (`$`, `$$`, `\)` or `\]`). */
    readMath(closer: string): string {
        let out = '';
        for (;;) {
            const c = this.nextCh();
            if (c === undefined) return out;
            if (c === '\\') {
                const n = this.nextCh();
                if (n === undefined) return out + c;
                if ((closer === '\\)' && n === ')') || (closer === '\\]' && n === ']')) return out;
                out += c + n;
                continue;
            }
            if (c === '%') { while (this.peekCh() !== undefined && this.peekCh() !== '\n') this.nextCh(); continue; }
            if (c === '$' && closer.startsWith('$')) {
                if (closer === '$$') { if (this.peekCh() === '$') this.nextCh(); return out; }
                return out;
            }
            out += c;
        }
    }
}

// ── state ───────────────────────────────────────────────────────────────────────────────────────

interface ParseState {
    fmt: TextFormatting;
    align?: TextAlignment;
    left?: number;
    right?: number;
    hang?: number;
    link?: { url: string; internal: boolean };
    /** Inside a moving context (a heading title): line breaks read as spaces. */
    inline?: boolean;
}

const cloneState = (s: ParseState): ParseState => ({ ...s, fmt: { ...s.fmt } });

/** One container being filled: the blocks it holds and the paragraph being built. */
class Flow {
    blocks: OfficeContentNode[] = [];
    inline: OfficeContentNode[] = [];
    pending = '';
    /**
     * The pending buffer's last character and whether it holds any non-space text. Tracked on
     * append because `pending` is built with `+=`: reading it (`endsWith`, `trim`) flattens the
     * whole string, and doing that per space would make a paragraph quadratic in its length.
     */
    pendingTail = '';
    pendingHasText = false;
    pendingState: ParseState | null = null;
    pendingMono = false;
    anchors: string[] = [];
    comments: OfficeContentNode[] = [];
    /** Anchors and comments met between blocks, for the next block. */
    nextAnchors: string[] = [];
    nextComments: OfficeContentNode[] = [];
    /** The block a following `\label` names (the heading, table or image just emitted). */
    labelTarget: OfficeContentNode | null = null;
    /** Alt text announced by a `% alt:` comment, for the next image. */
    pendingAlt: string | null = null;
    /** Paragraph style every paragraph in this flow gets (a quote, an abstract). */
    paragraphStyle?: string;
}

/** Why a flow stopped reading. */
type StopReason = 'eof' | 'egroup' | 'end' | 'item' | 'enddoc';

interface Stop {
    env?: string;
    group?: boolean;
    items?: string;
}

interface ListContext { listId: string; level: number; enumDepth: number; nextIndex?: number; }

/** One argument of an `xparse` (`\NewDocumentCommand`) signature. */
interface XArg { kind: 'm' | 'o' | 'O' | 's' | 't' | 'd' | 'D' | 'r' | 'R' | 'v' | 'b'; token?: string; open?: string; close?: string; value?: string; }
interface UserMacro { nargs: number; optDefault: string | null; body: string; spec?: XArg[]; }
interface UserEnv { nargs: number; optDefault: string | null; begin: string; end: string; spec?: XArg[]; }
/** A theorem-like environment: its printed title, `\theoremstyle`, and the counter it numbers with (none: unnumbered). */
interface TheoremDef { title: string; style: string; counter?: string; }

/** Files of a LaTeX project (a zip), addressed by normalized path. */
interface Project { files: Map<string, Buffer>; root: string; }

/** Normalizes a project-relative path, refusing one that leaves the project. */
function normalizeProjectPath(base: string, path: string): string | null {
    const cleaned = path.trim().replace(/\\/g, '/');
    if (!cleaned || cleaned.startsWith('/') || /^[A-Za-z]:/.test(cleaned)) return null;
    const parts: string[] = base ? base.split('/').filter(Boolean) : [];
    for (const seg of cleaned.split('/')) {
        if (!seg || seg === '.') continue;
        if (seg === '..') { if (parts.length === 0) return null; parts.pop(); continue; }
        parts.push(seg);
    }
    return parts.join('/');
}

/** `n` in lowercase roman numerals (LaTeX's `\roman`). */
function roman(n: number): string {
    if (n < 1 || n >= 4000) return String(n);
    const digits: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
    let out = '';
    for (const [v, s] of digits) while (n >= v) { out += s; n -= v; }
    return out;
}

/** `n` as a letter (LaTeX's `\alph`), or its digits past `z`. */
function alph(n: number): string {
    return n >= 1 && n <= 26 ? String.fromCharCode(96 + n) : String(n);
}

/** Applies TeX's input ligatures to a run of text (not in a typewriter font, where there are none). */
function ligatures(s: string): string {
    return s
        .replace(/---/g, '—').replace(/--/g, '–')
        .replace(/``/g, '“').replace(/''/g, '”')
        .replace(/!`/g, '¡').replace(/\?`/g, '¿')
        .replace(/`/g, '‘').replace(/'/g, '’');
}

/** Parses a `D:YYYYMMDDHHmmSS` PDF date (as written in `\hypersetup`) or an ISO-like date. */
function parseLatexDate(value: string): Date | undefined {
    const v = value.trim();
    const m = /^D:(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?/.exec(v);
    if (m) {
        const d = new Date(Date.UTC(+m[1], (+(m[2] ?? '1')) - 1, +(m[3] ?? '1'), +(m[4] ?? '0'), +(m[5] ?? '0'), +(m[6] ?? '0')));
        return isNaN(d.getTime()) ? undefined : d;
    }
    if (!/^\d{4}-\d{2}(-\d{2})?([T ].*)?$/.test(v)) return undefined;
    const d = new Date(v);
    return isNaN(d.getTime()) ? undefined : d;
}

/** Text of an inline node list. */
function textOf(nodes: OfficeContentNode[] | undefined): string {
    // A source comment is the author's hidden note, not text.
    return (nodes || []).map(n => (n.type === 'break' ? '\n' : isSourceComment(n) ? '' : n.text ?? textOf(n.children))).join('');
}

// ── the reader ──────────────────────────────────────────────────────────────────────────────────

/**
 * Interprets LaTeX document source into the shared AST.
 *
 * LaTeX is a macro language, so this reads it the way a document converter does rather than the
 * way TeX does: the document-level vocabulary (sectioning, lists, tables, floats, footnotes, math,
 * code, links, citations, the common formatting commands, beamer frames) is interpreted,
 * user-defined macros and environments are expanded, and anything else keeps its text and is
 * reported. Nothing reads the file system: included files come only from a project zip.
 */
class LatexReader {
    private state: ParseState = { fmt: {} };
    private stateStack: ParseState[] = [];
    private nesting = 0;
    private macros = new Map<string, UserMacro>();
    private envs = new Map<string, UserEnv>();
    private colors = new Map<string, string>();
    private expansions = 0;
    private expandedChars = 0;
    private expansionLimitHit = false;
    private nestingLimitHit = false;
    private tokens = 0;
    private atLetter = false;
    private finished = false;

    metadata: OfficeMetadata = {};
    private native: Record<string, any> = { packages: [] as string[] };
    /** Raw `\title`/`\subtitle`/`\author`/`\date` arguments, for the block `\maketitle` typesets. */
    private titleParts: { title?: string; subtitle?: string; author?: string; date?: string } = {};
    private titleTypeset = false;
    /** The last `\title` checked for text of its own (see typesetTitleBlock). */
    private titleChecked: { raw: string; empty: boolean } | undefined;
    /** The text of the first title block typeset, for the ones past the repeated-content budget (see typesetTitleBlock). */
    private firstTitleText: string | undefined;
    /** Theorem-like kinds already headed once (see theorem). */
    private readonly theoremsHeaded = new Set<TheoremDef>();
    /** Set while a title block is being built, so a `\maketitle` inside `\title` cannot re-enter it. */
    private inTitle = false;
    /** Depth of plain-text extraction (metadata, labels), which never typesets a title block. */
    private plainDepth = 0;
    /** Where the last `% <!--` look-ahead stopped without finding `-->` (see sourceComment). */
    private commentScanEnd: { frame: Frame; i: number } | undefined;
    /** Tables currently being built (see MAX_TABLE_NESTING_DEPTH). */
    private tableDepth = 0;
    /** Whether included files have been refused for exceeding the expansion budget. */
    private includeLimitHit = false;
    /** Metadata fields `\hypersetup` set explicitly (`pdftitle`, `pdfauthor`, `pdfkeywords`), which `\title`/`\author`/`\keywords` do not override. */
    private readonly pdfMetadata = new Set<'title' | 'author' | 'keywords'>();
    /** `\newif` switches (which `ifthen` and etoolbox booleans are too) with their current values. */
    private ifs = new Map<string, boolean>();
    /** etoolbox toggles with their current values. */
    private toggles = new Map<string, boolean>();
    /**
     * Conditionals open where reading is: `taken` (the true branch of a decided test, whose `\else`
     * part is skipped), `else` (the false branch of one), `both` (a test not decided: both are read).
     */
    private conds: ('taken' | 'else' | 'both')[] = [];
    /** Theorem-like environments (`\newtheorem`, `\declaretheorem`) by name. */
    private theorems = new Map<string, TheoremDef>();
    /** The `\theoremstyle` a new theorem-like environment takes. */
    private theoremStyle = 'plain';
    /** Theorem counters: the last number given, the counter numbered within, and its number then. */
    private theoremCounters = new Map<string, { value: number; within?: string; prefix: string }>();
    /** Paragraphs a theorem or proof head was put in, which a theorem around them does not merge into. */
    private theoremHeads = new WeakSet<OfficeContentNode>();
    /** The blocks a keywords command or environment printed, which an abstract's description leaves out. */
    private keywordBlocks = new WeakSet<OfficeContentNode>();
    /** Whether the current proof's end-of-proof mark was placed (`\qedhere`). */
    private qedPlaced = false;
    /** Languages babel loads (its options) and the one polyglossia or `\babelprovide` makes the main one. */
    private babelLanguages: string[] = [];
    private mainLanguage?: string;
    /** `\babeltags` tags, which name languages in `\text<tag>` and `<tag>` environments. */
    private languageTags = new Map<string, string>();
    /** xparse environments open, with the arguments their end code may use. */
    private openEnvs = new Map<string, string[][]>();
    /** Whether the document is ConTeXt (`\starttext`), read for its text as a fallback. */
    private context = false;
    /** ConTeXt lists open (`\startitemize`), with the LaTeX environment each is read as. */
    private contextLists: string[] = [];
    /** The legacy encoding the main file declares (`inputenc`), for included files that declare none. */
    inputEncoding: string | undefined;
    attachments: OfficeAttachment[] = [];
    private attachmentByPath = new Map<string, string>();
    /** What decoding PDF pictures into images may cost this document, across all of them. */
    private readonly decodeBudget = newDecodeBudget();
    private headerFields = new Map<string, string>();
    private footerFields = new Map<string, string>();
    private graphicsPaths: string[] = [];

    private classSize = 10;
    private beamer = false;
    private numbered = true;
    private sectionRanks = new Set<number>();
    private listCounter = 0;
    private slideCounter = 0;
    private noteCounter = 0;
    private lastSlide: OfficeContentNode | null = null;
    private pendingMarks: OfficeContentNode[] = [];
    private refs: { node: OfficeContentNode; label: string; kind: string }[] = [];
    private labelTargets = new Map<string, OfficeContentNode>();
    /** Every `\label` in reading order, so an environment can find the labels written inside it. */
    private labelLog: string[] = [];
    /** The number of the `enumerate` item being read at each enumerate depth, for `\ref`s to items. */
    private enumNumbers: number[] = [];
    private unknown = new Set<string>();
    private missingFiles = new Set<string>();
    private bodyStarted = false;
    private docFlow: Flow | null = null;

    constructor(private config: FullOfficeParserConfig, private project: Project | null, private mainDir: string) { }

    // ── entry ──

    parse(src: string, file?: string): OfficeContentNode[] {
        this.prescan(src);
        const flow = new Flow();
        this.docFlow = flow;
        this.parseFlow(new Scanner(src, file), flow, {});
        this.endParagraph(flow);
        // References resolve before runs merge: a reference without a link is merged into the text around it.
        this.resolveRefs();
        const content = this.finishBlocks(flow.blocks);
        // `\hypersetup{pdflang=...}` states the language outright; otherwise it is babel's or polyglossia's main one.
        const language = this.documentLanguage();
        if (language && !this.metadata.language) this.metadata.language = language;
        if (this.unknown.size) logWarning(OfficeWarningType.LATEX_CONSTRUCT_NOT_INTERPRETED, this.config, { constructs: [...this.unknown].sort() });
        if (this.missingFiles.size) logWarning(OfficeWarningType.LATEX_FILE_NOT_FOUND, this.config, { files: [...this.missingFiles] });
        this.metadata.nativeProperties = { ...this.native };
        // The class option is the document's body size; recording it lets a generator measure run
        // sizes against it (and write the same class size back).
        this.metadata.formatting = { ...(this.metadata.formatting ?? {}), size: `${this.classSize}pt` };
        return content;
    }

    /**
     * Settles what the whole document decides before reading it: which sectioning command is the
     * top level (a document using `\chapter` starts there, one using `\section` starts at
     * sections), whether sections are numbered, and the document class.
     */
    private prescan(src: string): void {
        // File data (an image carried as a PDF, say) is no evidence of the document's structure.
        src = withoutFilecontents(src);
        // `\documentstyle` is LaTeX 2.09's `\documentclass`.
        // Options and name stop at the next bracket or brace, and the whitespace after the options is
        // part of them: each scanning to the end, or trying every split of the spaces between them, took
        // time in the square of an unclosed `\documentclass[` repeated.
        const cls = /\\document(?:class|style)\s*(?:\[([^\][]*)\]\s*)?\{([^{}]*)\}/.exec(src);
        if (cls) {
            this.native.documentClass = cls[2].trim();
            // A macro among the options (the LaTeX generator's driver choice) cannot be expanded here: it is left out.
            this.native.classOptions = (cls[1] ?? '').replace(/\\[A-Za-z@]+\s*/g, ',').split(',').map(s => s.trim()).filter(Boolean);
            this.beamer = this.native.documentClass === 'beamer';
            const size = this.native.classOptions.map((o: string) => /^(10|11|12)pt$/.exec(o)).find(Boolean);
            if (size) this.classSize = +size[1];
        } else if (/\\starttext\b/.test(src)) {
            // ConTeXt is a different macro format: its text is read, with its sectioning, lists and
            // code, but little else of it is interpreted.
            this.context = true;
            this.unknown.add('ConTeXt (the document is not LaTeX)');
        }
        SECTIONING.forEach((name, rank) => {
            if (new RegExp(`\\\\${name}\\*?\\s*[[{]`).test(src) || (this.context && new RegExp(`\\\\start${name}\\b`).test(src))) this.sectionRanks.add(rank);
        });
        if (this.context) {
            // ConTeXt's unnumbered title levels read as the numbered ones they stand for.
            [['title', 'chapter'], ['subject', 'section'], ['subsubject', 'subsection'], ['subsubsubject', 'subsubsection']].forEach(([title, name]) => {
                if (new RegExp(`\\\\start${title}\\b`).test(src)) this.sectionRanks.add(SECTIONING.indexOf(name));
            });
        }
        if (/\\setcounter\s*\{secnumdepth\}\s*\{\s*-/.test(src)) this.numbered = false;
    }

    // ── state ──

    private pushState(): void {
        this.stateStack.push(this.state);
        this.state = cloneState(this.state);
        this.nesting++;
    }

    private popState(): void {
        const s = this.stateStack.pop();
        if (s) { this.state = s; this.nesting--; }
    }

    private withState<T>(patch: (s: ParseState) => void, fn: () => T): T {
        this.pushState();
        patch(this.state);
        try { return fn(); } finally { this.popState(); }
    }

    // ── text accumulation ──

    private sameRunState(a: ParseState, b: ParseState): boolean {
        return JSON.stringify(a.fmt) === JSON.stringify(b.fmt) && a.link?.url === b.link?.url;
    }

    private addText(flow: Flow, s: string, literal = false): void {
        if (!s) return;
        const mono = this.state.fmt.font === 'monospace';
        if (flow.pendingState && (!this.sameRunState(flow.pendingState, this.state) || literal !== flow.pendingMono)) this.flushText(flow);
        if (!flow.pendingState) { flow.pendingState = cloneState(this.state); flow.pendingMono = literal || mono; }
        flow.pending += s;
        flow.pendingTail = s[s.length - 1];
        if (!flow.pendingHasText && /\S/.test(s)) flow.pendingHasText = true;
        flow.labelTarget = null;
    }

    private addSpace(flow: Flow): void {
        if (flow.inline.length === 0 && flow.pending.length === 0) return;
        if (flow.pendingTail === ' ') return;
        this.addText(flow, ' ');
    }

    private clearPending(flow: Flow): void {
        flow.pending = '';
        flow.pendingTail = '';
        flow.pendingHasText = false;
        flow.pendingState = null;
    }

    private flushText(flow: Flow): void {
        if (flow.pending.length === 0 || !flow.pendingState) { this.clearPending(flow); return; }
        const st = flow.pendingState;
        const text = flow.pendingMono ? flow.pending : ligatures(flow.pending);
        const node: OfficeContentNode = { type: 'text', text };
        const fmt = Object.fromEntries(Object.entries(st.fmt).filter(([, v]) => v !== undefined && v !== false));
        if (Object.keys(fmt).length) node.formatting = fmt as TextFormatting;
        if (st.link && !this.config.ignoreInternalLinks || (st.link && !st.link.internal)) {
            node.metadata = { link: st.link!.url, linkType: st.link!.internal ? 'internal' : 'external' } as TextMetadata;
        }
        flow.inline.push(node);
        this.clearPending(flow);
    }

    /** The current run formatting, as a node's `formatting` (none when plain), for a run built outside {@link addText}. */
    private runFormatting(): { formatting?: TextFormatting } {
        const fmt = Object.fromEntries(Object.entries(this.state.fmt).filter(([, v]) => v !== undefined && v !== false));
        return Object.keys(fmt).length ? { formatting: fmt as TextFormatting } : {};
    }

    private addInline(flow: Flow, node: OfficeContentNode): void {
        this.flushText(flow);
        flow.inline.push(node);
        flow.labelTarget = null;
    }

    /** The node a note or comment met at this point attaches to: the last text run (or a new empty one). */
    private anchorRun(flow: Flow): OfficeContentNode {
        this.flushText(flow);
        const last = flow.inline[flow.inline.length - 1];
        if (last && last.type === 'text') return last;
        const empty: OfficeContentNode = { type: 'text', text: '' };
        flow.inline.push(empty);
        return empty;
    }

    // ── blocks ──

    private hasContent(nodes: OfficeContentNode[]): boolean {
        return nodes.some(n => n.type !== 'text' || (n.text ?? '').replace(/\s+/g, '') !== '' || n.notes?.length || n.comments?.length);
    }

    private endParagraph(flow: Flow): void {
        this.flushText(flow);
        const inline = flow.inline;
        flow.inline = [];
        if (!this.hasContent(inline)) {
            // Anchors/comments met in an empty paragraph belong to whatever comes next.
            flow.nextAnchors.push(...flow.anchors);
            flow.nextComments.push(...flow.comments);
            flow.anchors = [];
            flow.comments = [];
            return;
        }
        const children = this.trimInline(inline);
        const meaningful = children.filter(n => !(n.type === 'text' && !(n.text ?? '').trim() && !n.notes?.length && !n.comments?.length));
        // A paragraph that is only a rule is a thematic break; only an image is a block image.
        if (meaningful.length === 1 && meaningful[0].type === 'break' && (meaningful[0].metadata as any)?.breakType === 'thematic') {
            this.pushBlock(flow, meaningful[0]);
            return;
        }
        if (meaningful.length === 1 && meaningful[0].type === 'image') {
            const img = meaningful[0];
            if (this.state.align && this.state.align !== 'justify') (img.metadata as ImageMetadata).align = this.state.align === 'left' ? 'left' : this.state.align;
            this.pushBlock(flow, img);
            return;
        }
        // Typewriter lines with explicit breaks and nothing else are a code block written where a
        // verbatim environment cannot go (a table cell, a footnote).
        if (children.some(n => n.type === 'break') && children.every(n => n.type === 'break' || (n.type === 'text' && n.formatting?.font === 'monospace' && !n.metadata))) {
            const code = children.map(n => (n.type === 'break' ? '\n' : n.text)).join('');
            this.pushBlock(flow, { type: 'code', text: code, metadata: {} as CodeMetadata });
            return;
        }
        const meta: ParagraphMetadata = {};
        if (this.state.align && this.state.align !== 'justify') meta.alignment = this.state.align;
        if (this.state.left || this.state.right || this.state.hang) {
            meta.paragraphIndentation = {};
            if (this.state.left) meta.paragraphIndentation.left = Math.round(this.state.left * 20);
            if (this.state.right) meta.paragraphIndentation.right = Math.round(this.state.right * 20);
            if (this.state.hang) meta.paragraphIndentation.hanging = Math.round(this.state.hang * 20);
        }
        if (flow.paragraphStyle) meta.style = flow.paragraphStyle;
        const para: OfficeContentNode = { type: 'paragraph', text: textOf(children), children, metadata: meta };
        if (flow.anchors.length) { meta.anchorIds = [...flow.anchors]; flow.anchors = []; }
        if (flow.comments.length) { para.comments = [...flow.comments]; flow.comments = []; }
        this.pushBlock(flow, para);
        flow.labelTarget = null;
    }

    private trimInline(nodes: OfficeContentNode[]): OfficeContentNode[] {
        const out = nodes.slice();
        const first = out.find(n => n.type === 'text');
        if (first && out.indexOf(first) === 0) first.text = (first.text ?? '').replace(/^[ \t]+/, '');
        const last = [...out].reverse().find(n => n.type === 'text');
        if (last && out.indexOf(last) === out.length - 1) last.text = trimEndChars(last.text ?? '', ' \t');
        while (out.length && out[out.length - 1].type === 'break' && (out[out.length - 1].metadata as any)?.breakType === 'carriageReturn') out.pop();
        return out.filter(n => !(n.type === 'text' && n.text === '' && !n.notes?.length && !n.comments?.length));
    }

    /** Adds a block, giving it the anchors and comments waiting for the next block. */
    private pushBlock(flow: Flow, node: OfficeContentNode): void {
        if (flow.nextAnchors.length && !this.config.ignoreInternalLinks) {
            const meta = (node.metadata ??= {} as any) as any;
            meta.anchorIds = [...(meta.anchorIds ?? []), ...flow.nextAnchors];
        }
        flow.nextAnchors = [];
        if (flow.nextComments.length) node.comments = [...(node.comments ?? []), ...flow.nextComments];
        flow.nextComments = [];
        flow.blocks.push(node);
        flow.labelTarget = node;
    }

    private addBlock(flow: Flow, node: OfficeContentNode): void {
        this.endParagraph(flow);
        this.pushBlock(flow, node);
    }

    private addLabel(flow: Flow, label: string): void {
        if (!label) return;
        this.labelLog.push(label);
        // Without internal links a label is no anchor, but a `\ref` to it still reads its number.
        const links = !this.config.ignoreInternalLinks;
        const pendingText = flow.pendingHasText || this.hasContent(flow.inline);
        if (flow.labelTarget && !pendingText) {
            if (links) {
                const meta = (flow.labelTarget.metadata ??= {} as any) as any;
                meta.anchorIds = [...(meta.anchorIds ?? []), label];
            }
            this.labelTargets.set(label, flow.labelTarget);
        } else if (pendingText) {
            if (links) flow.anchors.push(label);
            this.labelTargets.set(label, { type: 'paragraph', text: '' } as OfficeContentNode);
        } else if (links) {
            flow.nextAnchors.push(label);
        }
    }

    // ── the main loop ──

    private parseFlow(sc: Scanner, flow: Flow, stop: Stop): StopReason {
        let localGroups = 0;
        const unwind = () => { while (localGroups-- > 0) { this.flushText(flow); this.popState(); } };
        for (;;) {
            if (this.finished) { unwind(); return 'enddoc'; }
            if (++this.tokens % ABORT_CHECK_INTERVAL === 0) checkAbortSignal(this.config.abortSignal);
            const snap = sc.save();
            const tok = sc.next(this.atLetter);
            switch (tok.t) {
                case 'eof': unwind(); return 'eof';
                case 'bgroup':
                    this.flushText(flow);
                    if (this.nesting >= MAX_NESTING_DEPTH) { this.tooDeep(flow, sc.readRawUntil('}').text); break; }
                    this.pushState();
                    localGroups++;
                    break;
                case 'egroup':
                    if (localGroups > 0) {
                        // The paragraph a group-scoped declaration applies to ends inside the group
                        // (`{\centering ...\par}`); anything still open takes the outer state.
                        this.flushText(flow);
                        this.popState();
                        localGroups--;
                    } else if (stop.group) { unwind(); return 'egroup'; }
                    break;
                case 'par': this.endParagraph(flow); break;
                case 'space': this.addSpace(flow); break;
                case 'text': this.addText(flow, tok.v); break;
                case 'amp': break;
                case 'dollar': this.readDollarMath(sc, flow); break;
                case 'comment': this.handleComment(flow, tok.v, sc); break;
                case 'cs': {
                    if (stop.items && tok.name === stop.items) { sc.restore(snap); unwind(); return 'item'; }
                    if (tok.name === 'end') {
                        const env = (sc.readRawGroup() ?? '').trim();
                        if (stop.env && env === stop.env) { unwind(); return 'end'; }
                        if (env === 'document') { this.endParagraph(flow); this.finished = true; unwind(); return 'enddoc'; }
                        const user = this.envs.get(env);
                        if (user) this.endUserEnv(sc, env, user);
                        break;
                    }
                    const r = this.command(sc, flow, tok.name, stop);
                    if (r) { unwind(); return r; }
                    break;
                }
            }
        }
    }

    /** Parses raw source into the current flow, in the current state. */
    private parseRawInto(flow: Flow, raw: string | null): StopReason {
        if (raw === null) return 'eof';
        if (this.nesting >= MAX_NESTING_DEPTH) { this.tooDeep(flow, raw); return 'eof'; }
        this.nesting++;
        try { return this.parseFlow(new Scanner(raw), flow, {}); } finally { this.nesting--; }
    }

    /** Parses raw source as the body of a separate container (a note, a cell) and returns its blocks. */
    private parseBlocksOf(raw: string, patch?: (f: Flow) => void, freshState = true): OfficeContentNode[] {
        const flow = new Flow();
        patch?.(flow);
        // An argument parsed into its own container (a note, a cell, a caption) is one level deeper.
        const run = () => {
            if (this.nesting >= MAX_NESTING_DEPTH) { this.tooDeep(flow, raw); this.endParagraph(flow); return; }
            this.nesting++;
            try { this.parseFlow(new Scanner(raw), flow, {}); } finally { this.nesting--; }
            this.endParagraph(flow);
        };
        if (freshState) {
            const savedState = this.state, savedStack = this.stateStack, savedNesting = this.nesting;
            this.state = { fmt: {} };
            this.stateStack = [];
            try { run(); } finally { this.state = savedState; this.stateStack = savedStack; this.nesting = savedNesting; }
        } else {
            run();
        }
        return flow.blocks;
    }

    /** Plain text of raw source (for metadata values and labels). */
    private plainText(raw: string | null): string {
        if (raw === null) return '';
        const saved = this.finished;
        this.plainDepth++;
        let blocks: OfficeContentNode[];
        try { blocks = this.parseBlocksOf(raw); } finally { this.plainDepth--; }
        this.finished = saved;
        return blocks.map(b => b.text ?? textOf(b.children)).join(' ').replace(/\s+/g, ' ').trim();
    }

    // ── macros ──

    private expand(sc: Scanner, src: string): boolean {
        if (this.expansionLimitHit) return false;
        this.expansions++;
        this.expandedChars += src.length;
        if (this.expansions > MAX_MACRO_EXPANSIONS || this.expandedChars > MAX_EXPANDED_CHARS || sc.depth >= MAX_INPUT_FRAMES) {
            this.expansionLimitHit = true;
            logWarning(OfficeWarningType.LATEX_EXPANSION_LIMIT_REACHED, this.config, { limit: 'macro expansion' });
            return false;
        }
        sc.push(src);
        return true;
    }

    /** Reads a user macro's arguments and returns its body with them substituted. */
    private substitute(sc: Scanner, def: UserMacro): string {
        return this.fill(def.body, this.readUserArgs(sc, def));
    }

    /**
     * Reads a user macro's or environment's arguments: by its `xparse` signature, or else `n`
     * mandatory ones, the first optional when it has a default. `env` names the environment whose
     * body a `b` argument takes.
     */
    private readUserArgs(sc: Scanner, def: { nargs: number; optDefault: string | null; spec?: XArg[] }, env?: string): string[] {
        if (def.spec) return def.spec.map(a => this.readXArg(sc, a, env));
        const args: string[] = [];
        for (let k = 0; k < def.nargs; k++) {
            if (k === 0 && def.optDefault !== null) args.push(sc.readRawOptional() ?? def.optDefault);
            else args.push(sc.readRawGroup() ?? '');
        }
        return args;
    }

    /** A definition's body with its parameters (`#1`...`#9`, `##`) replaced by the arguments. */
    private fill(body: string, args: string[]): string {
        // Size the expansion before building it: a short definition repeating a long argument can
        // otherwise allocate hundreds of megabytes before the expansion budget is checked.
        let size = 0;
        for (const m of body.matchAll(/##|#([1-9])/g)) size += m[0] === '##' ? 1 : (args[+m[1] - 1] ?? '').length;
        if (this.expandedChars + body.length + size > MAX_EXPANDED_CHARS) {
            if (!this.expansionLimitHit) logWarning(OfficeWarningType.LATEX_EXPANSION_LIMIT_REACHED, this.config, { limit: 'macro expansion' });
            this.expansionLimitHit = true;
            return '';
        }
        return body.replace(/##|#([1-9])/g, (m, n) => (m === '##' ? '#' : (args[+n - 1] ?? '')));
    }

    /** One `xparse` argument: an absent optional one is `-NoValue-` (or its default), a star or token a boolean. */
    private readXArg(sc: Scanner, a: XArg, env?: string): string {
        switch (a.kind) {
            case 'm': return sc.readRawGroup() ?? '';
            case 's': return sc.readStar() ? '\\BooleanTrue' : '\\BooleanFalse';
            case 't': return sc.readChar(a.token!) ? '\\BooleanTrue' : '\\BooleanFalse';
            case 'b': return env ? sc.readRawEnvBody(env) : '';
            case 'v': {
                // Verbatim: the characters themselves, so they are written back as text rather than read as LaTeX.
                let text: string;
                if (sc.nextIs('{')) text = sc.readRawGroup() ?? '';
                else {
                    sc.skipBlanks();
                    const delim = sc.nextCh();
                    text = delim === undefined ? '' : sc.readRawUntil(delim).text;
                }
                return text.replace(/[\\{}$&#^_%~]/g, c => VERBATIM_CHARS[c]);
            }
            default: {
                const open = a.open ?? '[', close = a.close ?? ']';
                // A delimiter that also closes (`d||`) cannot nest: the argument ends at its next occurrence.
                const value = open === close ? (sc.readChar(open) ? sc.readRawUntil(close).text : null) : sc.readRawOptional(open, close);
                return value ?? (a.kind === 'O' || a.kind === 'D' || a.kind === 'R' ? a.value ?? '' : NO_VALUE);
            }
        }
    }

    /**
     * Parses an `xparse` argument specification (`m`, `o`, `O{default}`, `s`, `t<token>`,
     * `d<open><close>`, `D..{default}`, `r..`, `R..{default}`, `v`, `b`, with the `+`, `!` and
     * `>{processor}` prefixes ignored), or returns null for one using anything else.
     */
    private parseXSpec(spec: string): XArg[] | null {
        const out: XArg[] = [];
        let i = 0;
        const blanks = () => { while (i < spec.length && /\s/.test(spec[i])) i++; };
        const group = (): string | null => {
            blanks();
            if (spec[i] !== '{') return null;
            let depth = 0;
            const start = i;
            for (; i < spec.length; i++) {
                if (spec[i] === '\\') { i++; continue; }
                if (spec[i] === '{') depth++;
                else if (spec[i] === '}' && --depth === 0) { i++; return spec.slice(start + 1, i - 1); }
            }
            return null;
        };
        const token = (): string | null => {
            blanks();
            const c = spec[i];
            if (c === undefined || c === '\\' || c === '{' || c === '}' || c === '#') return null;
            i++;
            return c;
        };
        for (;;) {
            blanks();
            if (i >= spec.length) return out.length <= 9 ? out : null;
            const k = spec[i++];
            if (k === '+' || k === '!') continue;
            if (k === '>') { if (group() === null) return null; continue; }
            if (k === 'm' || k === 'o' || k === 's' || k === 'v' || k === 'b') out.push({ kind: k });
            else if (k === 'O') { const value = group(); if (value === null) return null; out.push({ kind: k, value }); }
            else if (k === 't') { const t = token(); if (t === null) return null; out.push({ kind: k, token: t }); }
            else if (k === 'd' || k === 'r' || k === 'D' || k === 'R') {
                const open = token(), close = token();
                if (open === null || close === null) return null;
                const value = k === 'D' || k === 'R' ? group() : undefined;
                if (value === null) return null;
                out.push({ kind: k, open, close, ...(value !== undefined ? { value } : {}) });
            } else return null;
        }
    }

    /** Expands user macros inside a math formula (the formula is otherwise kept verbatim). */
    private expandMath(src: string): string {
        if (this.macros.size === 0) return src;
        let out = '';
        const sc = new Scanner(src);
        let steps = 0;
        for (;;) {
            if (++steps > MAX_EXPANDED_CHARS) break;
            const c = sc.nextCh();
            if (c === undefined) break;
            if (c !== '\\') { out += c; continue; }
            let name = '';
            if (/[A-Za-z]/.test(sc.peekCh() ?? '')) { while (/[A-Za-z]/.test(sc.peekCh() ?? '')) name += sc.nextCh(); }
            else { out += c + (sc.nextCh() ?? ''); continue; }
            const def = this.macros.get(name);
            if (!def || this.expansionLimitHit) { out += '\\' + name; continue; }
            const body = this.substitute(sc, def);
            if (!this.expand(sc, body)) { out += '\\' + name; continue; }
            if (/^[A-Za-z]/.test(sc.peekCh() ?? '')) sc.push(' ');
        }
        return out;
    }

    private defineMacro(sc: Scanner, redefine: boolean): void {
        sc.readStar();
        let name = (sc.readRawGroup() ?? '').trim().replace(/^\\/, '');
        const n = sc.readRawOptional();
        const opt = sc.readRawOptional();
        const body = sc.readRawGroup() ?? '';
        if (!name) return;
        name = name.replace(/\s+/g, '');
        if (!redefine && this.macros.has(name)) return;
        if (PROTECTED_COMMANDS.has(name)) return;
        this.macros.set(name, { nargs: Math.max(0, Math.min(9, parseInt(n ?? '0', 10) || 0)), optDefault: opt, body });
    }

    private defineTexMacro(sc: Scanner): void {
        const nameTok = sc.next(this.atLetter);
        if (nameTok.t !== 'cs') return;
        let params = '';
        while (sc.peekCh() !== undefined && sc.peekCh() !== '{') params += sc.nextCh();
        const body = sc.readRawGroup() ?? '';
        // Only undelimited parameters (#1#2...) are understood; a delimited-parameter macro is skipped.
        if (PROTECTED_COMMANDS.has(nameTok.name)) return;
        if (!/^(#[1-9])*\s*$/.test(params)) { this.unknown.add(`\\def\\${nameTok.name} (delimited parameters)`); return; }
        this.macros.set(nameTok.name, { nargs: (params.match(/#/g) || []).length, optDefault: null, body });
    }

    private defineEnv(sc: Scanner, redefine: boolean): void {
        sc.readStar();
        const name = (sc.readRawGroup() ?? '').trim();
        const n = sc.readRawOptional();
        const opt = sc.readRawOptional();
        const begin = sc.readRawGroup() ?? '';
        const end = sc.readRawGroup() ?? '';
        if (!name || (!redefine && this.envs.has(name))) return;
        this.envs.set(name, { nargs: Math.max(0, Math.min(9, parseInt(n ?? '0', 10) || 0)), optDefault: opt, begin, end });
    }

    /**
     * `\NewDocumentCommand` and its family (`xparse`, in the LaTeX kernel since 2020): `New` and
     * `Provide` keep a command already defined, `Renew` and `Declare` replace it.
     */
    private defineDocumentCommand(sc: Scanner, form: string): void {
        const name = (sc.readRawGroup() ?? '').trim().replace(/^\\/, '').replace(/\s+/g, '');
        const spec = sc.readRawGroup() ?? '';
        const body = sc.readRawGroup() ?? '';
        if (!name || PROTECTED_COMMANDS.has(name)) return;
        if (!/^(Renew|Declare)/.test(form) && this.macros.has(name)) return;
        const args = this.parseXSpec(spec);
        if (!args) { this.unknown.add(`\\${name} (argument specification ${spec.trim()})`); return; }
        this.macros.set(name, { nargs: args.length, optDefault: null, body, spec: args });
    }

    /** `\NewDocumentEnvironment` and its family: begin and end code both see the arguments. */
    private defineDocumentEnvironment(sc: Scanner, form: string): void {
        const name = (sc.readRawGroup() ?? '').trim();
        const spec = sc.readRawGroup() ?? '';
        const begin = sc.readRawGroup() ?? '';
        const end = sc.readRawGroup() ?? '';
        if (!name) return;
        if (!/^(Renew|Declare)/.test(form) && this.envs.has(name)) return;
        const args = this.parseXSpec(spec);
        if (!args) { this.unknown.add(`${name} environment (argument specification ${spec.trim()})`); return; }
        this.envs.set(name, { nargs: args.length, optDefault: null, begin, end, spec: args });
    }

    /** `\begin` of a user environment: its begin code, with its arguments, is read in place. */
    private beginUserEnv(sc: Scanner, env: string, user: UserEnv): void {
        const args = this.readUserArgs(sc, user, env);
        // An environment whose body is an argument (`b`) consumed its `\end` with it: its end code follows at once.
        if (user.spec?.some(a => a.kind === 'b')) { this.expand(sc, this.fill(user.begin, args) + this.fill(user.end, args)); return; }
        if (user.spec) {
            let open = this.openEnvs.get(env);
            if (!open) this.openEnvs.set(env, open = []);
            open.push(args);
        }
        if (!this.expand(sc, this.fill(user.begin, args))) sc.readRawEnvBody(env);
    }

    /** `\end` of a user environment: its end code is read in place (an xparse one's with its arguments). */
    private endUserEnv(sc: Scanner, env: string, user: UserEnv): void {
        if (!user.spec) { this.expand(sc, user.end); return; }
        // Kept per name: an `\end` with no matching `\begin` costs nothing to look up.
        const args = this.openEnvs.get(env)?.pop() ?? [];
        this.expand(sc, this.fill(user.end, args));
    }

    // ── conditionals ──

    /**
     * Whether a control word, just read from `sc`, is a TeX conditional (one that a skipped branch
     * must match with its `\fi`): a known one, or an `\if...` word a package defined (see ARGUMENT_TESTS).
     */
    private isConditional(name: string, sc: Scanner): boolean {
        if (own(FIXED_CONDITIONALS, name) || PRIMITIVE_CONDITIONALS.has(name) || this.ifs.has(name) || (name.startsWith('if') && name.includes('@'))) return true;
        return /^if[A-Za-z]+$/.test(name) && !ARGUMENT_TESTS.has(name) && !this.macros.has(name) && !sc.nextIs('{');
    }

    /**
     * Whether a command is defined when pdfLaTeX reads the document, or undefined when that is not
     * known here (a command of some package the parser does not know).
     */
    private isDefined(name: string): boolean | undefined {
        if (UNDEFINED_UNDER_PDFLATEX.has(name)) return false;
        if (this.macros.has(name) || this.ifs.has(name) || this.theorems.has(name) || this.envs.has(name)) return true;
        if (name === 'chapter') return CHAPTER_CLASSES.has(this.native.documentClass);
        if (own(PDFTEX_REGISTERS, name)) return true;
        if (own(SYMBOLS, name) || own(ACCENTS, name) || own(WRAPPER_COMMANDS, name) || own(IGNORED_COMMANDS, name) || own(FIXED_CONDITIONALS, name)
            || PROTECTED_COMMANDS.has(name) || PRIMITIVE_CONDITIONALS.has(name)) return true;
        return undefined;
    }

    /**
     * TeX conditionals, decided as pdfLaTeX would where the parser can: the fixed and engine tests
     * (`\ifXeTeX`), `\newif` switches, and `\ifdefined`, `\ifcsname` and `\ifx\cs\undefined` on
     * commands whose status is known. The branch not taken is skipped unread. A test it cannot decide
     * (`\ifnum`, `\ifx` in general) has both branches read and is reported. Returns false when
     * `name` is not a conditional's word.
     */
    private conditional(sc: Scanner, name: string): boolean {
        switch (name) {
            case 'else': case 'or': {
                const top = this.conds.pop();
                // The end of a taken branch: the rest, to the matching `\fi`, is the branch not taken.
                if (top === 'taken') { sc.skipBranch(n => this.isConditional(n, sc), false); return true; }
                if (top !== undefined) this.conds.push(top);
                return true;
            }
            case 'fi': this.conds.pop(); return true;
            case 'newif': {
                const tok = sc.next(this.atLetter);
                if (tok.t === 'cs') this.defineSwitch(tok.name);
                return true;
            }
            case 'unless': {
                const snap = sc.save();
                const tok = sc.next(true);
                if (tok.t === 'cs' && tok.name !== 'ifcase' && this.isConditional(tok.name, sc)) { this.branch(sc, tok.name, true); return true; }
                sc.restore(snap);
                return true;
            }
        }
        const flag = /^(.+)(true|false)$/.exec(name);
        if (flag && this.ifs.has(`if${flag[1]}`)) { this.ifs.set(`if${flag[1]}`, flag[2] === 'true'); return true; }
        if (!this.isConditional(name, sc)) return false;
        this.branch(sc, name, false);
        return true;
    }

    /** Opens conditional `name` (negated by `\unless`): decides it, and skips the branch not taken. */
    private branch(sc: Scanner, name: string, negate: boolean): void {
        let value = this.decide(sc, name);
        if (value === undefined) {
            this.conds.push('both');
            this.unknown.add(`\\${name}`);
            return;
        }
        if (negate) value = !value;
        if (value) { this.conds.push('taken'); return; }
        if (sc.skipBranch(n => this.isConditional(n, sc), true) === 'else') this.conds.push('else');
    }

    /** A conditional's value, reading its test, or undefined when it cannot be decided (the test is then left unread). */
    private decide(sc: Scanner, name: string): boolean | undefined {
        if (own(FIXED_CONDITIONALS, name)) return FIXED_CONDITIONALS[name];
        if (this.ifs.has(name)) return this.ifs.get(name);
        const snap = sc.save();
        if (name === 'ifdefined') {
            const a = sc.next(true);
            const d = a.t === 'cs' ? this.isDefined(a.name) : undefined;
            if (d === undefined) sc.restore(snap);
            return d;
        }
        if (name === 'ifx') {
            // `\ifx\cs\undefined` (or `\@undefined`, either way round): whether `\cs` is undefined.
            const a = sc.next(true);
            const b = sc.next(true);
            const marker = (t: Tok) => t.t === 'cs' && (t.name === 'undefined' || t.name === '@undefined');
            const d = a.t === 'cs' && b.t === 'cs' && (marker(b) || marker(a)) ? this.isDefined(marker(b) ? a.name : b.name) : undefined;
            if (d === undefined) { sc.restore(snap); return undefined; }
            return !d;
        }
        if (name === 'ifnum' || name === 'ifodd') {
            // Integer tests on literal numbers and on the pdfTeX registers documents test (`\pdfoutput` is 1).
            const operand = (): number | undefined => {
                sc.skipBlanks();
                if (sc.peekCh() === '\\') {
                    const t = sc.next(true);
                    return t.t === 'cs' && own(PDFTEX_REGISTERS, t.name) ? PDFTEX_REGISTERS[t.name] : undefined;
                }
                let digits = '';
                if (sc.peekCh() === '-' || sc.peekCh() === '+') digits += sc.nextCh();
                while (/\d/.test(sc.peekCh() ?? '') && digits.length < 12) digits += sc.nextCh();
                return /\d/.test(digits) ? parseInt(digits, 10) : undefined;
            };
            const a = operand();
            if (a === undefined) { sc.restore(snap); return undefined; }
            if (name === 'ifodd') { this.endNumber(sc); return Math.abs(a) % 2 === 1; }
            sc.skipBlanks();
            const rel = sc.nextCh();
            const b = rel === '<' || rel === '=' || rel === '>' ? operand() : undefined;
            if (b === undefined) { sc.restore(snap); return undefined; }
            this.endNumber(sc);
            return rel === '<' ? a < b : rel === '>' ? a > b : a === b;
        }
        if (name === 'ifcsname') {
            // The name is short: looking further for `\endcsname` (and re-reading on failure) would make
            // a run of unclosed `\ifcsname`s quadratic.
            let text = '';
            for (let c = sc.nextCh(); c !== undefined; c = sc.nextCh()) {
                text += c;
                if ((c === 'e' && text.endsWith('\\endcsname')) || text.length >= 256) break;
            }
            const m = /^\s*([A-Za-z@]+)\s*\\endcsname$/.exec(text);
            const d = m ? this.isDefined(m[1]) : undefined;
            if (d === undefined) sc.restore(snap);
            return d;
        }
        return undefined;
    }

    /** After a number TeX reads one optional space. */
    private endNumber(sc: Scanner): void {
        if (sc.peekCh() === ' ') sc.nextCh();
    }

    /**
     * The argument-style tests of LaTeX and its packages (`\@ifpackageloaded`, `\@ifundefined`,
     * `ifthen`'s `\ifthenelse` on booleans, etoolbox's `\ifdef` and toggles, xparse's `\IfBooleanTF`
     * and `\IfNoValueTF`): the branch taken is read in place. Returns false when `name` is not one.
     */
    private testCommand(sc: Scanner, name: string): boolean {
        let test: boolean | undefined;
        let branches = 'TF';
        switch (name) {
            case '@ifpackageloaded': case 'IfPackageLoadedTF': case '@ifclassloaded': case 'IfClassLoadedTF': {
                const pkg = (sc.readRawGroup() ?? '').trim();
                test = /class/i.test(name) ? this.native.documentClass === pkg : (this.native.packages as string[]).includes(pkg);
                break;
            }
            case '@ifundefined': {
                const d = this.isDefinedRaw(sc.readRawGroup());
                test = d === undefined ? undefined : !d;
                break;
            }
            case 'ifdef': case 'ifundef': case 'ifcsdef': case 'ifcsundef': {
                const d = this.isDefinedRaw(sc.readRawGroup());
                test = d === undefined ? undefined : name.endsWith('undef') ? !d : d;
                break;
            }
            case 'iftoggle': test = this.toggles.get((sc.readRawGroup() ?? '').trim()); break;
            case 'ifbool': test = this.ifs.get(`if${(sc.readRawGroup() ?? '').trim()}`); break;
            case 'ifthenelse': {
                const cond = (sc.readRawGroup() ?? '').trim();
                const bool = /^(\\NOT\s*)?\\boolean\s*\{([^{}]*)\}$/.exec(cond);
                const equal = /^\\equal\s*\{([^{}\\]*)\}\s*\{([^{}\\]*)\}$/.exec(cond);
                if (bool) { const v = this.ifs.get(`if${bool[2].trim()}`); test = v === undefined ? undefined : !!bool[1] !== v; }
                else if (equal) test = equal[1].trim() === equal[2].trim();
                break;
            }
            default: {
                const x = /^If(Boolean|NoValue|Value|Blank)(TF|T|F)$/.exec(name);
                if (!x) return false;
                const arg = (sc.readRawGroup() ?? '').trim();
                test = x[1] === 'Boolean' ? arg === '\\BooleanTrue' : x[1] === 'NoValue' ? arg === NO_VALUE : x[1] === 'Value' ? arg !== NO_VALUE : arg === '';
                branches = x[2];
            }
        }
        const yes = branches.includes('T') ? sc.readRawGroup() ?? '' : '';
        const no = branches.includes('F') ? sc.readRawGroup() ?? '' : '';
        if (test === undefined) {
            // Not decided: both branches are read, as before the test was understood.
            this.unknown.add(`\\${name}`);
            this.expand(sc, `${yes}${no}`);
            return true;
        }
        this.expand(sc, test ? yes : no);
        return true;
    }

    /** {@link isDefined} for a command named in an argument (`\cs` or its bare name). */
    private isDefinedRaw(raw: string | null): boolean | undefined {
        const m = /^\s*\\?([A-Za-z@]+)\s*$/.exec(raw ?? '');
        return m ? this.isDefined(m[1]) : undefined;
    }

    /** Defines `\newif` switch `name` (false), unless the name is one of TeX's own conditionals. */
    private defineSwitch(name: string): void {
        if (!/^if[A-Za-z@]+$/.test(name) || PRIMITIVE_CONDITIONALS.has(name) || own(FIXED_CONDITIONALS, name) || this.ifs.has(name)) return;
        this.ifs.set(name, false);
    }

    /**
     * Switches set by commands: `\newboolean`/`\setboolean` (ifthen), `\newbool`/`\setbool`/`\booltrue`
     * and `\newtoggle`/`\settoggle`/`\toggletrue` (etoolbox). Returns false when `name` is not one.
     */
    private setSwitch(sc: Scanner, name: string): boolean {
        const key = () => (sc.readRawGroup() ?? '').trim();
        switch (name) {
            case 'newboolean': case 'provideboolean': case 'newbool': case 'providebool': this.defineSwitch(`if${key()}`); return true;
            case 'setboolean': case 'setbool': {
                const k = `if${key()}`;
                const v = /^\s*true\s*$/i.test(sc.readRawGroup() ?? '');
                if (this.ifs.has(k)) this.ifs.set(k, v);
                return true;
            }
            case 'booltrue': case 'boolfalse': { const k = `if${key()}`; if (this.ifs.has(k)) this.ifs.set(k, name === 'booltrue'); return true; }
            case 'newtoggle': case 'providetoggle': { const k = key(); if (!this.toggles.has(k)) this.toggles.set(k, false); return true; }
            case 'settoggle': { const k = key(); this.toggles.set(k, /^\s*true\s*$/i.test(sc.readRawGroup() ?? '')); return true; }
            case 'toggletrue': case 'togglefalse': this.toggles.set(key(), name === 'toggletrue'); return true;
        }
        return false;
    }

    // ── comments ──

    private handleComment(flow: Flow, text: string, sc?: Scanner): void {
        const t = text.replace(/^ /, '');
        if (t.startsWith('<!--') && sc && this.sourceComment(flow, t, sc)) return;
        const alt = /^alt: (.*)$/.exec(t);
        if (alt) { flow.pendingAlt = alt[1]; return; }
        const m = /^Comment(?: \(([^)]*)\))?: ([\s\S]*)$/.exec(t);
        const last = flow.comments[flow.comments.length - 1] ?? flow.nextComments[flow.nextComments.length - 1]
            ?? this.lastInlineComment(flow);
        if (!m) {
            // A continuation line of the comment just read (a multi-line comment is one `%` line each).
            if (last && (last as any).__open) this.appendCommentLine(last, t);
            return;
        }
        if (last) delete (last as any).__open;
        if (this.config.ignoreComments) return;
        const meta: CommentMetadata = {};
        if (m[1]) {
            const parts = m[1].split(', ');
            const maybeDate = parts[parts.length - 1];
            if (parts.length > 1 && /^\d{4}-\d{2}-\d{2}/.test(maybeDate)) { meta.date = maybeDate; parts.pop(); }
            else if (parts.length === 1 && /^\d{4}-\d{2}-\d{2}/.test(parts[0])) { meta.date = parts[0]; parts.pop(); }
            if (parts.length) meta.author = parts.join(', ');
        }
        const node: OfficeContentNode = { type: 'comment', text: m[2], metadata: meta, children: [{ type: 'paragraph', text: m[2], children: [{ type: 'text', text: m[2] }] }] };
        (node as any).__open = true;
        if (flow.pendingHasText || this.hasContent(flow.inline)) {
            const run = this.anchorRun(flow);
            (run.comments ??= []).push(node);
        } else {
            flow.nextComments.push(node);
        }
    }

    /**
     * `% <!--body-->`, one `%` line per line of the body: a source comment (Markdown/HTML `<!-- ... -->`,
     * the author's hidden note) as the LaTeX generator writes it. Restored verbatim as the same node,
     * inline when the paragraph has text before it. It is not a review comment, so `ignoreComments` does
     * not govern it, and it takes no anchors meant for the block after it. A `<!--` whose `-->` does not
     * end one of the directly following comment lines is left an ordinary comment.
     */
    private sourceComment(flow: Flow, first: string, sc: Scanner): boolean {
        // A scan that found no `-->` also rules out every `<!--` line it read: none of them can close
        // before the same point. Remembering where it stopped keeps a file of unclosed openers linear.
        const at = sc.position();
        if (at && this.commentScanEnd && at.frame === this.commentScanEnd.frame && at.i < this.commentScanEnd.i) return false;
        const afterFirst = sc.save();
        const lines = [first.slice(4)];
        // `<!-->` and `<!--->` are complete, empty comments (as in HTML and CommonMark).
        const closed = () => {
            const last = lines[lines.length - 1];
            if (lines.length === 1 && (last === '>' || last === '->')) return true;
            return last.endsWith('-->') && (lines.length > 1 || last.length >= 3);
        };
        while (!closed()) {
            const beforeTok = sc.position();
            const tok = sc.next(this.atLetter);
            if (tok.t !== 'comment') {
                this.commentScanEnd = beforeTok;
                sc.restore(afterFirst);
                return false;
            }
            lines.push(tok.v.replace(/^ /, ''));
        }
        const joined = lines.join('\n');
        const body = lines.length === 1 && (joined === '>' || joined === '->') ? '' : joined.slice(0, -3);
        const node: OfficeContentNode = { type: 'comment', text: body, metadata: { sourceSyntax: 'html' } as CommentMetadata };
        // Inline when the paragraph has text before it, or text follows on the very next line (a comment
        // that opens a paragraph); a block comment stands between blank lines.
        const beforeNext = sc.save();
        const textFollows = sc.next(this.atLetter).t === 'text';
        sc.restore(beforeNext);
        if (flow.pendingHasText || this.hasContent(flow.inline) || textFollows) {
            this.addInline(flow, node);
        } else {
            this.endParagraph(flow);
            flow.blocks.push(node);
        }
        return true;
    }

    private lastInlineComment(flow: Flow): OfficeContentNode | undefined {
        for (let k = flow.inline.length - 1; k >= 0; k--) {
            const c = flow.inline[k].comments;
            if (c?.length) return c[c.length - 1];
        }
        return undefined;
    }

    private appendCommentLine(comment: OfficeContentNode, line: string): void {
        comment.text = `${comment.text}\n${line}`;
        const para = comment.children?.[0];
        if (para) { para.text = comment.text; para.children = [{ type: 'text', text: comment.text }]; }
    }

    // ── math ──

    private readDollarMath(sc: Scanner, flow: Flow): void {
        if (sc.peekCh() === '$') {
            sc.nextCh();
            this.addBlockMath(flow, sc.readMath('$$'));
            return;
        }
        this.addInlineMath(flow, sc.readMath('$'));
    }

    private addInlineMath(flow: Flow, raw: string): void {
        const latex = this.expandMath(raw).trim();
        if (!latex) return;
        this.addInline(flow, { type: 'code', text: latex, metadata: { math: 'inline' } as CodeMetadata });
    }

    private addBlockMath(flow: Flow, raw: string): void {
        const latex = this.expandMath(raw).trim();
        if (!latex) return;
        if (this.state.inline) { this.addInline(flow, { type: 'code', text: latex, metadata: { math: 'inline' } as CodeMetadata }); return; }
        // Display math sits inside its paragraph in TeX; in the AST it is its own block, and the
        // text after it continues as a new paragraph.
        this.addBlock(flow, { type: 'code', text: latex, metadata: { math: 'block' } as CodeMetadata });
    }

    // ── commands ──

    /**
     * Reads arguments by signature: `s` star, `o` optional `[..]`, `O` beamer overlay `<..>`,
     * `m` mandatory. Returns the raw values in order (a star as `'*'` or null).
     */
    private readArgs(sc: Scanner, spec: string): (string | null)[] {
        const out: (string | null)[] = [];
        for (const k of spec) {
            if (k === 's') out.push(sc.readStar() ? '*' : null);
            else if (k === 'o') out.push(sc.readRawOptional());
            else if (k === 'O') out.push(sc.readRawOptional('<', '>'));
            else out.push(sc.readRawGroup());
        }
        return out;
    }

    private command(sc: Scanner, flow: Flow, name: string, stop: Stop): StopReason | void {
        // User definitions win over built-ins, as in TeX.
        const macro = this.macros.get(name);
        if (macro) {
            const body = this.substitute(sc, macro);
            if (!this.expand(sc, body)) this.unknown.add(`\\${name}`);
            return;
        }
        if (this.conditional(sc, name) || this.setSwitch(sc, name) || this.testCommand(sc, name)) return;

        if (own(SYMBOLS, name)) {
            if (name === '-' || name === ',') this.addText(flow, SYMBOLS[name]);
            else if (SYMBOLS[name] === '' ) { /* spacing no-op */ }
            else this.addText(flow, SYMBOLS[name]);
            return;
        }
        if (own(ACCENTS, name)) {
            let arg = sc.readRawGroup() ?? '';
            if (arg === '\\i') arg = 'i';
            if (arg === '\\j') arg = 'j';
            const base = arg.startsWith('\\') ? (own(SYMBOLS, arg.slice(1)) ? SYMBOLS[arg.slice(1)] : '') : arg;
            this.addText(flow, (base + ACCENTS[name]).normalize('NFC'));
            return;
        }

        switch (name) {
            // ── document structure ──
            case 'begin': return this.beginEnv(sc, flow, stop);
            case 'part': case 'chapter': case 'section': case 'subsection': case 'subsubsection': case 'paragraph': case 'subparagraph':
            case 'addchap': case 'addsec':
                this.heading(sc, flow, name === 'addchap' ? 'chapter' : name === 'addsec' ? 'section' : name);
                return;
            case 'par': this.endParagraph(flow); return;
            case '\\': case 'newline': case 'linebreak': case 'break': case 'tabularnewline': case 'cr':
                if (name === '\\') { sc.readStar(); sc.readRawOptional(); }
                if (name === 'linebreak') sc.readRawOptional();
                if (this.state.inline) this.addSpace(flow);
                else this.addInline(flow, { type: 'break', metadata: { breakType: 'carriageReturn' } as any });
                return;
            case 'hfil': case 'hfill': case 'hss': return;
            case 'clearpage': case 'newpage': case 'cleardoublepage': case 'pagebreak': case 'include':
                if (name === 'pagebreak') sc.readRawOptional();
                if (name === 'include') { this.includeFile(sc, flow, sc.readRawGroup()); return; }
                this.endParagraph(flow);
                if (this.config.includeBreakNodes) this.pushBlock(flow, { type: 'break', metadata: { breakType: 'page' } as any });
                return;
            case 'input': case 'subfile': case 'include_':
                this.includeFile(sc, flow, sc.readRawGroup(), name === 'subfile');
                return;
            case 'import': case 'subimport': case 'inputfrom': case 'subinputfrom': {
                sc.readStar();
                const [dir, file] = this.readArgs(sc, 'mm');
                this.includeFile(sc, flow, `${(dir ?? '').replace(/\/?$/, '/')}${file ?? ''}`);
                return;
            }
            case 'item': case 'bibitem': {
                // Outside a list: read the item's content as plain text.
                sc.readRawOptional();
                if (name === 'bibitem') sc.readRawGroup();
                return;
            }
            case 'label': this.addLabel(flow, (sc.readRawGroup() ?? '').trim()); return;
            case 'caption': case 'captionof': {
                if (name === 'captionof') sc.readRawGroup();
                sc.readStar();
                sc.readRawOptional();
                const raw = sc.readRawGroup();
                this.endParagraph(flow);
                const blocks = this.parseBlocksOf(raw ?? '', f => { f.paragraphStyle = 'Caption'; }, true);
                for (const b of blocks) this.pushBlock(flow, b);
                return;
            }
            case 'frametitle': {
                sc.readRawOptional('<', '>');
                sc.readRawOptional();
                const raw = sc.readRawGroup();
                this.addBlock(flow, this.headingNode(raw ?? '', 1));
                return;
            }
            case 'framesubtitle': {
                sc.readRawOptional('<', '>');
                const raw = sc.readRawGroup();
                this.addBlock(flow, this.headingNode(raw ?? '', 2));
                return;
            }
            case 'note': {
                sc.readRawOptional('<', '>');
                sc.readRawOptional();
                const raw = sc.readRawGroup();
                if (this.config.ignoreNotes || !this.beamer) return;
                const blocks = this.parseBlocksOf(raw ?? '');
                const target = (flow as any).__slide ?? this.lastSlide;
                if (target) {
                    const slideNumber = (target.metadata as any)?.slideNumber;
                    (target.notes ??= []).push({ type: 'note', text: blocks.map(b => b.text ?? '').join('\n'), children: blocks, metadata: { slideNumber, noteId: `slide-note-${slideNumber}` } as NoteMetadata });
                }
                return;
            }
            case 'titlepage': (flow as any).__titlepage = true; this.typesetTitle(flow); return;
            case 'frame': {
                // `\frame{...}`, the command form of the frame environment.
                sc.readRawOptional('<', '>');
                const opts = sc.readRawOptional();
                const body = sc.readRawGroup() ?? '';
                this.expand(sc, `\\begin{frame}${opts !== null ? `[${opts}]` : ''}\\relax ${body}\\end{frame}`);
                return;
            }

            // ── notes, comments, citations, references ──
            case 'footnote': case 'endnote': case 'footnotetext': case 'footnotemark': {
                const num = sc.readRawOptional();
                const raw = name === 'footnotemark' ? null : sc.readRawGroup();
                if (this.config.ignoreNotes) return;
                const noteType = name === 'endnote' ? 'endnote' : 'footnote';
                if (name === 'footnotetext') {
                    const mark = this.pendingMarks.shift();
                    const blocks = this.parseBlocksOf(raw ?? '');
                    if (mark) { mark.children = blocks; mark.text = blocks.map(b => b.text ?? '').join('\n'); return; }
                    this.addBlock(flow, { type: 'note', text: blocks.map(b => b.text ?? '').join('\n'), children: blocks, metadata: { noteType, noteId: num ?? String(++this.noteCounter) } as NoteMetadata });
                    return;
                }
                const blocks = raw === null ? [] : this.parseBlocksOf(raw);
                const note: OfficeContentNode = { type: 'note', text: blocks.map(b => b.text ?? '').join('\n'), children: blocks, metadata: { noteType, noteId: num ?? String(++this.noteCounter) } as NoteMetadata };
                if (name === 'footnotemark') this.pendingMarks.push(note);
                (this.anchorRun(flow).notes ??= []).push(note);
                return;
            }
            case 'todo': case 'marginpar': case 'marginnote': case 'pdfcomment': case 'comment': {
                const opts = name === 'todo' || name === 'pdfcomment' ? sc.readRawOptional() : null;
                const raw = sc.readRawGroup();
                if (this.config.ignoreComments) return;
                const author = opts ? /author\s*=\s*\{?([^,}]*)\}?/.exec(opts)?.[1]?.trim() : undefined;
                const text = this.plainText(raw);
                const node: OfficeContentNode = { type: 'comment', text, metadata: author ? { author } : {}, children: [{ type: 'paragraph', text, children: [{ type: 'text', text }] }] };
                (this.anchorRun(flow).comments ??= []).push(node);
                return;
            }
            case 'cite': case 'citep': case 'citet': case 'parencite': case 'textcite': case 'autocite': case 'footcite':
            case 'Cite': case 'Citep': case 'Citet': case 'Parencite': case 'Textcite': case 'Autocite': case 'citealp': case 'citealt':
            case 'citeauthor': case 'citeyear': case 'supercite': case 'smartcite': case 'nocite_': {
                sc.readStar();
                sc.readRawOptional();
                sc.readRawOptional();
                const keys = (sc.readRawGroup() ?? '').split(',').map(k => k.trim()).filter(Boolean);
                keys.forEach((key, k) => {
                    if (k > 0) this.addText(flow, '; ');
                    this.addInline(flow, { type: 'text', text: key, ...this.runFormatting(), metadata: { citationKey: key } as TextMetadata });
                });
                return;
            }
            case 'ref': case 'autoref': case 'cref': case 'Cref': case 'eqref': case 'pageref': case 'nameref': case 'vref': case 'Autoref': {
                sc.readStar();
                const label = (sc.readRawGroup() ?? '').trim();
                const node: OfficeContentNode = { type: 'text', text: label, ...this.runFormatting() };
                if (!this.config.ignoreInternalLinks) node.metadata = { link: `#${label}`, linkType: 'internal' } as TextMetadata;
                this.refs.push({ node, label, kind: name });
                this.addInline(flow, node);
                return;
            }
            case 'hyperref': {
                const label = sc.readRawOptional();
                const raw = sc.readRawGroup();
                if (label === null) { this.parseRawInto(flow, raw); return; }
                this.withLink(flow, `#${label.trim()}`, true, raw);
                return;
            }
            case 'hyperlink': {
                const [target, raw] = this.readArgs(sc, 'mm');
                this.withLink(flow, `#${(target ?? '').trim()}`, true, raw);
                return;
            }
            case 'hypertarget': {
                const [target, raw] = this.readArgs(sc, 'mm');
                this.addLabel(flow, (target ?? '').trim());
                this.parseRawInto(flow, raw);
                return;
            }
            case 'href': {
                sc.readRawOptional();
                const [url, raw] = this.readArgs(sc, 'mm');
                const u = this.unescapeUrl(url ?? '');
                this.withLink(flow, u, u.startsWith('#'), raw);
                return;
            }
            case 'url': case 'nolinkurl': {
                const u = this.unescapeUrl(sc.readRawGroup() ?? '');
                if (name === 'nolinkurl') { this.addText(flow, u, true); return; }
                this.withState(s => { s.link = { url: u, internal: false }; }, () => this.addText(flow, u, true));
                return;
            }
            // An anchor at this spot: a \label after it names what follows, not the block before.
            case 'phantomsection': flow.labelTarget = null; return;

            // ── inline formatting (arguments) ──
            case 'textbf': case 'textit': case 'emph': case 'textsl': case 'underline': case 'uline': case 'uuline': case 'uwave':
            case 'sout': case 'st': case 'xout': case 'texttt': case 'textsf': case 'textrm': case 'textsuperscript': case 'textsubscript':
            case 'textnormal': case 'textup': case 'textmd': case 'textsc': case 'hl': case 'ul': case 'so': case 'lowercase': case 'uppercase':
            case 'MakeUppercase': case 'MakeLowercase': case 'textls': {
                const raw = sc.readRawGroup();
                this.withState(s => this.applyFormat(s, name), () => this.parseRawInto(flow, raw));
                return;
            }
            case 'textcolor': {
                const model = sc.readRawOptional();
                const [spec, raw] = this.readArgs(sc, 'mm');
                const hex = this.resolveColor(spec ?? '', model);
                this.withState(s => { if (hex) s.fmt.color = hex; }, () => this.parseRawInto(flow, raw));
                return;
            }
            case 'colorbox': case 'fcolorbox': {
                const model = sc.readRawOptional();
                if (name === 'fcolorbox') sc.readRawGroup();
                const [spec, raw] = this.readArgs(sc, 'mm');
                const hex = this.resolveColor(spec ?? '', model);
                this.withState(s => { if (hex) s.fmt.backgroundColor = hex; }, () => this.parseRawInto(flow, raw));
                return;
            }
            case 'color': {
                const model = sc.readRawOptional();
                const hex = this.resolveColor(sc.readRawGroup() ?? '', model);
                if (hex) this.state.fmt.color = hex;
                return;
            }
            case 'fontsize': {
                const [size] = this.readArgs(sc, 'mm');
                const pt = /^\s*([\d.]+)\s*(?:(pt)\s*)?$/.exec(size ?? '');
                if (pt) this.state.fmt.size = `${Math.round(parseFloat(pt[1]) * 100) / 100}pt`;
                return;
            }
            case 'tiny': case 'scriptsize': case 'footnotesize': case 'small': case 'normalsize': case 'large': case 'Large':
            case 'LARGE': case 'huge': case 'Huge': {
                if (name === 'normalsize') { delete this.state.fmt.size; return; }
                this.state.fmt.size = `${SIZE_TABLE[this.classSize][name]}pt`;
                return;
            }
            case 'bfseries': this.state.fmt.bold = true; return;
            case 'mdseries': delete this.state.fmt.bold; return;
            case 'itshape': case 'slshape': this.state.fmt.italic = true; return;
            case 'upshape': delete this.state.fmt.italic; return;
            case 'em': this.state.fmt.italic = !this.state.fmt.italic || undefined; return;
            case 'ttfamily': this.state.fmt.font = 'monospace'; return;
            case 'sffamily': this.state.fmt.font = 'sans-serif'; return;
            case 'rmfamily': delete this.state.fmt.font; return;
            case 'normalfont': delete this.state.fmt.bold; delete this.state.fmt.italic; delete this.state.fmt.font; return;
            case 'scshape': return;
            // LaTeX 2.09's font switches, which LaTeX2e keeps: each starts from the normal font.
            case 'bf': case 'it': case 'sl': case 'tt': case 'sf': case 'rm': case 'sc': {
                const fmt = this.state.fmt;
                delete fmt.bold; delete fmt.italic; delete fmt.font;
                if (name === 'bf') fmt.bold = true;
                else if (name === 'it' || name === 'sl') fmt.italic = true;
                else if (name === 'tt') fmt.font = 'monospace';
                else if (name === 'sf') fmt.font = 'sans-serif';
                return;
            }
            case 'centering': this.state.align = 'center'; return;
            case 'raggedleft': this.state.align = 'right'; return;
            case 'raggedright': this.state.align = 'left'; return;
            case 'leftskip': case 'rightskip': case 'hangindent': case 'parindent': case 'hangafter': case 'parskip': {
                const v = this.readDimenAssignment(sc);
                this.readGlueTail(sc);
                if (name === 'leftskip') this.state.left = v;
                else if (name === 'rightskip') this.state.right = v;
                else if (name === 'hangindent') this.state.hang = v;
                return;
            }
            case 'hspace': {
                const star = sc.readStar();
                const len = this.dimenPt(sc.readRawGroup() ?? '');
                if (star && flow.pending.length === 0 && !this.hasContent(flow.inline) && len) {
                    this.state.left = this.state.left ?? 0;
                    this.addText(flow, '');
                    (flow as any).__firstLine = len;
                } else this.addSpace(flow);
                return;
            }
            case 'phantom': case 'hphantom': {
                const width = this.plainText(sc.readRawGroup()).length;
                this.addText(flow, ' '.repeat(width), true);
                return;
            }
            case 'rule': {
                sc.readRawOptional();
                this.readArgs(sc, 'mm');
                this.addInline(flow, { type: 'break', metadata: { breakType: 'thematic' } as any });
                return;
            }
            case 'hrule': case 'hrulefill':
                this.addInline(flow, { type: 'break', metadata: { breakType: 'thematic' } as any });
                return;
            case 'ensuremath': this.addInlineMath(flow, sc.readRawGroup() ?? ''); return;
            case '(': this.addInlineMath(flow, sc.readMath('\\)')); return;
            case '[': this.addBlockMath(flow, sc.readMath('\\]')); return;
            case 'verb': case 'lstinline': case 'mintinline': case 'Verb': {
                if (name === 'mintinline') sc.readRawGroup();
                else if (name !== 'verb') sc.readRawOptional();
                if (name === 'verb') sc.readStar();
                const delim = sc.nextCh();
                const text = delim === '{' ? sc.readRawUntil('}').text : sc.readRawUntil(delim ?? '|').text;
                this.withState(s => { s.fmt.font = 'monospace'; }, () => this.addText(flow, text, true));
                return;
            }
            case 'ding': {
                const n = (sc.readRawGroup() ?? '').trim();
                const ch = LATEX_SYMBOL_CHARACTERS.get(`\\ding{${n}}`);
                if (ch) this.addText(flow, ch);
                return;
            }
            case 'includegraphics': this.image(sc, flow); return;

            // ── definitions and metadata ──
            case 'newcommand': case 'renewcommand': case 'providecommand': case 'DeclareRobustCommand':
                this.defineMacro(sc, name !== 'providecommand' && name !== 'newcommand'); return;
            case 'NewDocumentCommand': case 'RenewDocumentCommand': case 'ProvideDocumentCommand': case 'DeclareDocumentCommand':
            case 'NewExpandableDocumentCommand': case 'RenewExpandableDocumentCommand': case 'ProvideExpandableDocumentCommand':
            case 'DeclareExpandableDocumentCommand':
                this.defineDocumentCommand(sc, name); return;
            case 'NewDocumentEnvironment': case 'RenewDocumentEnvironment': case 'ProvideDocumentEnvironment': case 'DeclareDocumentEnvironment':
                this.defineDocumentEnvironment(sc, name); return;
            // expl3 code is a programming layer with its own syntax, not document content.
            case 'ExplSyntaxOn': sc.readRawUntil('\\ExplSyntaxOff'); return;
            case 'newtheorem': {
                const star = sc.readStar();
                const env = (sc.readRawGroup() ?? '').trim();
                const shared = sc.readRawOptional();
                const title = sc.readRawGroup() ?? '';
                const within = shared === null ? sc.readRawOptional() : null;
                this.defineTheorem(env, title, star ? undefined : (shared?.trim() || env), within?.trim());
                return;
            }
            case 'declaretheorem': {
                // thmtools: `\declaretheorem[options]{names}` (or options after the names).
                const before = sc.readRawOptional();
                const names = sc.readRawGroup() ?? '';
                const after = before === null ? sc.readRawOptional() : null;
                const kv = new Map(this.keyValues(before ?? after ?? ''));
                for (const env of names.split(',').map(n => n.trim()).filter(Boolean)) {
                    const title = kv.get('name') ?? kv.get('title') ?? env.charAt(0).toUpperCase() + env.slice(1);
                    const shared = (kv.get('sibling') ?? kv.get('numberlike') ?? kv.get('sharenumber'))?.trim();
                    const numbered = !/^\s*no\s*$/.test(kv.get('numbered') ?? '');
                    this.defineTheorem(env, title, numbered ? shared || env : undefined, (kv.get('numberwithin') ?? kv.get('within') ?? kv.get('parent'))?.trim(), kv.get('style')?.trim());
                }
                return;
            }
            case 'theoremstyle': this.theoremStyle = (sc.readRawGroup() ?? '').trim() || 'plain'; return;
            case 'qed': case 'qedhere':
                if (this.state.inline) return;
                this.addSpace(flow);
                this.addText(flow, '□');
                this.qedPlaced = true;
                return;
            case 'qedsymbol': this.addText(flow, '□'); return;
            case 'today': case 'DTMtoday': this.addText(flow, this.todayText()); return;
            case 'def': case 'gdef': case 'edef': case 'xdef': case 'long': case 'global':
                if (name === 'long' || name === 'global') return;
                this.defineTexMacro(sc); return;
            case 'let': {
                const a = sc.next(this.atLetter);
                if (sc.peekCh() === '=') sc.nextCh();
                sc.skipBlanks();
                const b = sc.next(this.atLetter);
                if (a.t === 'cs' && b.t === 'cs') {
                    const target = this.macros.get(b.name);
                    if (target) this.macros.set(a.name, target);
                    else if (!own(SYMBOLS, b.name)) this.macros.set(a.name, { nargs: 0, optDefault: null, body: `\\${b.name}` });
                    else this.macros.set(a.name, { nargs: 0, optDefault: null, body: SYMBOLS[b.name] });
                }
                return;
            }
            case 'newenvironment': case 'renewenvironment': this.defineEnv(sc, name === 'renewenvironment'); return;
            case 'definecolor': {
                const [cname, model, spec] = this.readArgs(sc, 'mmm');
                const hex = this.resolveColor(spec ?? '', model);
                if (cname && hex) this.colors.set(cname.trim(), hex);
                return;
            }
            case 'colorlet': {
                const [cname, spec] = this.readArgs(sc, 'mm');
                const hex = this.resolveColor(spec ?? '', null);
                if (cname && hex) this.colors.set(cname.trim(), hex);
                return;
            }
            case 'title': { sc.readRawOptional(); const raw = sc.readRawGroup(); this.titleParts.title = raw ?? undefined; const t = this.plainText(this.stripThanks(raw)); if (t && !this.pdfMetadata.has('title')) this.metadata.title = t; return; }
            case 'author': {
                sc.readRawOptional();
                // IEEEtran's author blocks: a name, then its affiliation lines, set apart in the title block.
                const raw = sc.readRawGroup()?.replace(/\\IEEEauthorblockA\b/g, ', \\IEEEauthorblockA') ?? null;
                this.titleParts.author = raw ?? undefined;
                const a = this.authorText(raw);
                if (a && !this.pdfMetadata.has('author')) this.metadata.author = a;
                return;
            }
            case 'date': { const raw = sc.readRawGroup(); this.titleParts.date = raw ?? undefined; const d = this.plainText(raw); if (d) this.native.date = d; return; }
            case 'subtitle': { sc.readRawOptional(); const raw = sc.readRawGroup(); this.titleParts.subtitle = raw ?? undefined; const t = this.plainText(raw); if (t) this.native.subtitle = t; return; }
            case 'maketitle': this.typesetTitle(flow); return;
            case 'keywords': case 'keyword': this.keywords(flow, sc.readRawGroup(), false); return;
            // amsart's author information, printed at the end of the article: kept as properties.
            case 'address': case 'curraddr': case 'email': case 'urladdr': {
                sc.readRawOptional();
                const text = this.plainText(sc.readRawGroup());
                const key = name === 'email' ? 'emails' : name === 'urladdr' ? 'urls' : 'addresses';
                // An `\email` inside `\author` is read again when the title block is typeset: record it once.
                if (text && !(this.native[key] ??= []).includes(text)) this.native[key].push(text);
                return;
            }
            case 'subjclass': {
                const year = sc.readRawOptional()?.trim();
                const codes = this.plainText(sc.readRawGroup());
                if (codes) this.native.subjectClassification = { codes, ...(year ? { scheme: `MSC${year}` } : {}) };
                return;
            }
            case 'IEEEPARstart': {
                const [first, rest] = this.readArgs(sc, 'mm');
                this.parseRawInto(flow, first);
                this.parseRawInto(flow, rest);
                return;
            }
            // KOMA-Script's `\minisec`: a small bold heading outside the document's structure.
            case 'minisec': {
                const raw = sc.readRawGroup();
                this.endParagraph(flow);
                this.withState(s => { s.fmt.bold = true; }, () => this.parseRawInto(flow, raw));
                this.endParagraph(flow);
                return;
            }
            // A quotation at the head of a chapter, its source set flush right (epigraph; KOMA-Script's `\dictum`).
            case 'epigraph': case 'dictum': {
                const author = name === 'dictum' ? sc.readRawOptional() : null;
                const text = sc.readRawGroup() ?? '';
                const source = name === 'epigraph' ? sc.readRawGroup() : author !== null ? `(${author})` : null;
                this.expand(sc, `\\begin{quote}${text}${source?.trim() ? `\\par{\\raggedleft ${source}\\par}` : ''}\\end{quote}`);
                return;
            }
            case 'setdefaultlanguage': case 'setmainlanguage': {
                const opts = sc.readRawOptional();
                const lang = (sc.readRawGroup() ?? '').trim();
                const variant = opts ? /variant\s*=\s*([A-Za-z]+)/.exec(opts)?.[1] : undefined;
                this.mainLanguage = variant && own(LANGUAGE_CODES, variant) ? variant : lang;
                return;
            }
            case 'babelprovide': {
                const opts = sc.readRawOptional() ?? '';
                const lang = (sc.readRawGroup() ?? '').trim();
                if (opts.split(',').some(o => o.trim() === 'main')) this.mainLanguage = lang;
                return;
            }
            case 'babeltags':
                for (const [tag, lang] of this.keyValues(sc.readRawGroup() ?? '')) if (tag && lang) this.languageTags.set(tag, lang.trim());
                return;
            // plain TeX's end of the document, and ConTeXt's.
            case 'bye': case 'stoptext':
                this.endParagraph(flow);
                this.finished = true;
                return 'enddoc';
            case 'starttext':
                if (!this.context) break;
                flow.blocks = [];
                flow.inline = [];
                this.clearPending(flow);
                this.bodyStarted = true;
                return;
            // plain TeX's section heading, its title delimited by the end of the paragraph.
            case 'beginsection': {
                let raw = '';
                for (;;) {
                    const snap = sc.save();
                    const tok = sc.next(this.atLetter);
                    if (tok.t === 'eof' || tok.t === 'par' || (tok.t === 'cs' && tok.name === 'par')) break;
                    if (tok.t === 'cs' && tok.name === 'end') { sc.restore(snap); break; }
                    raw += tok.t === 'cs' ? `\\${tok.name} ` : tok.t === 'text' ? tok.v : tok.t === 'space' ? ' ' : tok.t === 'bgroup' ? '{' : tok.t === 'egroup' ? '}' : tok.t === 'dollar' ? '$' : '';
                    if (raw.length > 10000) break;
                }
                this.addBlock(flow, this.headingNode(raw, this.headingLevel('section')));
                return;
            }
            // plain TeX glue and kerns: the amount is not text.
            case 'vskip': case 'hskip': case 'kern': case 'vglue': case 'hglue': case 'mskip': case 'mkern':
                this.readGlue(sc);
                return;
            case 'hypersetup': this.hypersetup(sc.readRawGroup() ?? ''); return;
            case 'usepackage': case 'RequirePackage': {
                const opts = sc.readRawOptional();
                const pkgs = (sc.readRawGroup() ?? '').split(',').map(p => p.trim()).filter(Boolean);
                (this.native.packages as string[]).push(...pkgs);
                if (pkgs.includes('babel')) this.babelLanguages.push(...(opts ?? '').split(',').map(o => o.trim()).filter(Boolean));
                return;
            }
            case 'graphicspath': {
                const raw = sc.readRawGroup() ?? '';
                this.graphicsPaths = [...raw.matchAll(/\{([^{}]*)\}/g)].map(m => m[1]);
                return;
            }
            case 'makeatletter': this.atLetter = true; return;
            case 'makeatother': this.atLetter = false; return;
            case 'setcounter': {
                const [counter, value] = this.readArgs(sc, 'mm');
                const c = (counter ?? '').trim();
                const v = parseInt((value ?? '').trim(), 10);
                if (c === 'secnumdepth' && v < 0) this.numbered = false;
                const list = (flow as any).__list as ListContext | undefined;
                if (list && /^enum(i|ii|iii|iv)$/.test(c) && Number.isFinite(v)) list.nextIndex = Math.max(0, v);
                return;
            }
            case 'fancyhead': case 'fancyfoot': case 'lhead': case 'chead': case 'rhead': case 'lfoot': case 'cfoot': case 'rfoot': {
                const pos = name.startsWith('fancy') ? (sc.readRawOptional() ?? 'LCR') : name[0].toUpperCase();
                const raw = sc.readRawGroup() ?? '';
                const target = name.includes('head') ? this.headerFields : this.footerFields;
                for (const p of ['L', 'C', 'R']) if (pos.toUpperCase().includes(p)) target.set(p, raw);
                return;
            }
            case 'IfFileExists': case 'InputIfFileExists': {
                // Nothing is read from disk, so the "file exists" branch is the document as written: it
                // is where the LaTeX generator puts an image the bundle could not include.
                const file = sc.readRawGroup();
                const then = sc.readRawGroup() ?? '';
                sc.readRawGroup();
                if (name === 'InputIfFileExists' && file !== null) this.includeFile(sc, flow, file);
                this.expand(sc, then);
                return;
            }
        }

        // babel's and polyglossia's `\text<language>{...}` (and `\text<tag>` for a `\babeltags` tag): the text alone.
        if (name.startsWith('text') && this.languageName(name.slice(4))) {
            sc.readRawOptional();
            this.parseRawInto(flow, sc.readRawGroup());
            return;
        }
        if (this.context && this.contextCommand(sc, flow, name)) return;
        if (own(WRAPPER_COMMANDS, name)) {
            this.readArgs(sc, WRAPPER_COMMANDS[name]);
            this.parseRawInto(flow, sc.readRawGroup());
            return;
        }
        if (TEX_REGISTERS.has(name)) {
            // An assignment (`=` or a value follows); a register named as a value elsewhere is left alone.
            const snap = sc.save();
            sc.skipBlanks();
            const next = sc.peekCh() ?? '';
            sc.restore(snap);
            if (next === '=' || /[\d.+-]/.test(next) || (/^(skip|dimen|count|toks)$/.test(name) && /\d/.test(next))) {
                if (/^(skip|dimen|count|toks)$/.test(name)) { sc.skipBlanks(); while (/\d/.test(sc.peekCh() ?? '')) sc.nextCh(); }
                this.readGlue(sc);
            }
            return;
        }
        if (own(IGNORED_COMMANDS, name)) {
            this.readArgs(sc, IGNORED_COMMANDS[name]);
            return;
        }
        // Unknown: keep what it would print in the common case (its braced arguments' text).
        if (/^[A-Za-z@]+$/.test(name)) this.unknown.add(`\\${name}`);
    }

    private applyFormat(s: ParseState, name: string): void {
        switch (name) {
            case 'textbf': s.fmt.bold = true; break;
            case 'textit': case 'textsl': s.fmt.italic = true; break;
            case 'emph': s.fmt.italic = !s.fmt.italic || undefined; break;
            case 'underline': case 'uline': case 'uuline': case 'uwave': case 'ul': s.fmt.underline = true; break;
            case 'sout': case 'st': case 'xout': case 'so': s.fmt.strikethrough = true; break;
            case 'texttt': s.fmt.font = 'monospace'; break;
            case 'textsf': s.fmt.font = 'sans-serif'; break;
            case 'textrm': case 'textnormal': delete s.fmt.font; break;
            case 'textup': delete s.fmt.italic; break;
            case 'textmd': delete s.fmt.bold; break;
            case 'textsuperscript': s.fmt.superscript = true; delete s.fmt.subscript; break;
            case 'textsubscript': s.fmt.subscript = true; delete s.fmt.superscript; break;
            case 'hl': s.fmt.backgroundColor = '#FFFF00'; break;
        }
    }

    private withLink(flow: Flow, url: string, internal: boolean, raw: string | null): void {
        if (internal && this.config.ignoreInternalLinks) { this.parseRawInto(flow, raw); return; }
        this.withState(s => { s.link = { url, internal }; }, () => this.parseRawInto(flow, raw));
    }

    private unescapeUrl(url: string): string {
        return url.trim().replace(/\\([#%&_~$^{}])/g, '$1').replace(/\s+/g, '');
    }

    /** Reads `=`? then a dimension after a TeX length register, in points. */
    private readDimenAssignment(sc: Scanner): number | undefined {
        sc.skipBlanks();
        if (sc.peekCh() === '=') sc.nextCh();
        sc.skipBlanks();
        let raw = '';
        while (/[-+\d.]/.test(sc.peekCh() ?? '')) raw += sc.nextCh();
        while (/[a-z]/.test(sc.peekCh() ?? '') && raw.length < 64) { raw += sc.nextCh(); if (/[a-z]{2}$/.test(raw)) break; }
        return this.dimenPt(raw);
    }

    private dimenPt(raw: string): number | undefined {
        const m = /^\s*([-+]?[\d.]+)\s*(?:(pt|bp|in|cm|mm|pc|em|ex|px|sp)\s*)?$/.exec(raw);
        if (!m) return undefined;
        const n = parseFloat(m[1]);
        const factor: Record<string, number> = { pt: 1, bp: 1.00375, in: 72.27, cm: 28.4528, mm: 2.84528, pc: 12, em: 10, ex: 4.3, px: 0.75, sp: 1 / 65536 };
        return Number.isFinite(n) ? Math.round(n * (factor[m[2] ?? 'pt'] ?? 1) * 100) / 100 : undefined;
    }

    // ── colours ──

    private resolveColor(spec: string, model: string | null): string | undefined {
        const s = spec.trim();
        const m = (model ?? '').trim();
        const hexOf = (r: number, g: number, b: number) => '#' + [r, g, b].map(v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase();
        if (m === 'HTML') return /^[0-9A-Fa-f]{6}$/.test(s) ? `#${s.toUpperCase()}` : undefined;
        if (m === 'rgb') { const v = s.split(',').map(Number); return v.length === 3 && v.every(Number.isFinite) ? hexOf(v[0] * 255, v[1] * 255, v[2] * 255) : undefined; }
        if (m === 'RGB') { const v = s.split(',').map(Number); return v.length === 3 && v.every(Number.isFinite) ? hexOf(v[0], v[1], v[2]) : undefined; }
        if (m === 'gray') { const v = Number(s); return Number.isFinite(v) ? hexOf(v * 255, v * 255, v * 255) : undefined; }
        if (m === 'cmyk') { const v = s.split(',').map(Number); return v.length === 4 && v.every(Number.isFinite) ? hexOf(255 * (1 - v[0]) * (1 - v[3]), 255 * (1 - v[1]) * (1 - v[3]), 255 * (1 - v[2]) * (1 - v[3])) : undefined; }
        // Named, possibly an xcolor mix `a!pct!b` (b defaults to white).
        const parts = s.split('!');
        const base = this.namedColor(parts[0]);
        if (!base) return undefined;
        if (parts.length === 1) return base;
        const pct = Math.max(0, Math.min(100, parseFloat(parts[1]) || 0)) / 100;
        const other = this.namedColor(parts[2] ?? 'white') ?? '#FFFFFF';
        const rgb = (h: string) => [1, 3, 5].map(k => parseInt(h.slice(k, k + 2), 16));
        const [a, b] = [rgb(base), rgb(other)];
        return hexOf(a[0] * pct + b[0] * (1 - pct), a[1] * pct + b[1] * (1 - pct), a[2] * pct + b[2] * (1 - pct));
    }

    private namedColor(name: string): string | undefined {
        const n = name.trim();
        const defined = this.colors.get(n);
        if (defined) return defined;
        const named = own(NAMED_COLORS, n) ? NAMED_COLORS[n] : undefined;
        return named ? `#${named}` : undefined;
    }

    // ── metadata helpers ──

    /**
     * The block `\maketitle` (or beamer's `\titlepage`) typesets, as content at that spot: a heading
     * styled `Title`, then `Subtitle`, `Author` and `Date` lines, the way a word processor's title
     * page reads. The values also stay in `ast.metadata`. `\thanks` become footnotes on the line that
     * carries them. A `\today` date reads as {@link todayText} says.
     */
    private typesetTitle(flow: Flow): void {
        // Only where content is being read: never while extracting plain text (metadata, a label),
        // and never from inside the title block itself (`\title{...\maketitle}`).
        if (this.inTitle || this.plainDepth > 0) return;
        this.inTitle = true;
        try { this.typesetTitleBlock(flow); } finally { this.inTitle = false; }
    }

    private typesetTitleBlock(flow: Flow): void {
        const { title, subtitle, author, date } = this.titleParts;
        // LaTeX refuses \maketitle without a \title, and empties the title after typesetting it once.
        if (!title || (this.titleTypeset && !this.beamer)) return;
        // beamer typesets the block at every \maketitle (a title frame per section is common), reading
        // it again from its source each time: after the first, each is charged to the document's
        // repeated-content budget (see repeatUtils) before it is read, past which it is the start of the
        // title's text. 2,000 title frames of a 100 KB title took 76 seconds and made 400 MB of HTML.
        if (this.titleTypeset) {
            const weight = title.length + (subtitle?.length ?? 0) + (author?.length ?? 0) + (date?.length ?? 0);
            if (takeRepeats(this.config, 1, weight) === 0) {
                this.endParagraph(flow);
                this.pushBlock(flow, { type: 'paragraph', text: repeatPreview(this.firstTitleText ?? ''), children: [{ type: 'text', text: repeatPreview(this.firstTitleText ?? '') }], metadata: { style: 'Title', alignment: 'center' } as ParagraphMetadata });
                return;
            }
        }
        // Whether the title has text of its own, found once per title (beamer asks at every \maketitle).
        if (title !== this.titleChecked?.raw) this.titleChecked = { raw: title, empty: !this.plainText(this.stripThanks(title)) };
        if (this.titleChecked.empty) return;
        this.titleTypeset = true;
        this.endParagraph(flow);
        const thanks = (raw: string) => raw.replace(/\\thanks\b/g, '\\footnote');
        const heading = this.headingNode(thanks(title), 1);
        Object.assign(heading.metadata as HeadingMetadata, { style: 'Title', alignment: 'center' });
        this.firstTitleText ??= heading.text;
        this.pushBlock(flow, heading);
        const line = (raw: string | undefined, style: string) => {
            if (raw === undefined) return;
            const node = this.headingNode(thanks(raw), 1);
            if (!node.text?.trim() && !node.children?.some(c => c.notes?.length)) return;
            this.pushBlock(flow, { type: 'paragraph', text: node.text, children: node.children, metadata: { style, alignment: 'center' } as ParagraphMetadata });
        };
        line(subtitle, 'Subtitle');
        // Authors are separated by \and; each one's own lines (affiliations) run together.
        line(author === undefined ? undefined : author.split(/\\and\b|\\AND\b/).map(a => a.trim()).filter(Boolean).join(', '), 'Author');
        line(date, 'Date');
        // A \label written right after \maketitle names the title.
        flow.labelTarget = heading;
    }

    /** The arguments of every `\name{...}` in raw source, in order. */
    private argumentsOf(raw: string, name: string): string[] {
        const sc = new Scanner(raw);
        const out: string[] = [];
        for (;;) {
            const tok = sc.next(this.atLetter);
            if (tok.t === 'eof') return out;
            if (tok.t === 'cs' && tok.name === name) out.push(sc.readRawGroup() ?? '');
        }
    }

    private stripThanks(raw: string | null): string | null {
        return raw === null ? null : raw.replace(/\\thanks\s*\{(?:[^{}]|\{[^{}]*\})*\}/g, '');
    }

    private authorText(raw: string | null): string {
        if (raw === null) return '';
        // With IEEEtran's author blocks, the authors are the names; the rest is their affiliations.
        if (/\\IEEEauthorblockN\b/.test(raw)) return this.argumentsOf(raw, 'IEEEauthorblockN').map(n => this.plainText(n.replace(/\\\\/g, ' '))).filter(Boolean).join(', ');
        return this.stripThanks(raw)!.split(/\\and\b|\\AND\b/).map(a => this.plainText(a.replace(/\\\\/g, ' '))).filter(Boolean).join(', ');
    }

    private hypersetup(raw: string): void {
        for (const [key, value] of this.keyValues(raw)) {
            // Read only for the keys that take it: `pdfinfo` reads its entries one by one, and reading
            // its whole value as well parsed each level of a nested one twice (twice as long per level).
            const v = ['pdftitle', 'pdfauthor', 'pdfsubject', 'pdfkeywords'].includes(key) ? this.plainText(value) : '';
            switch (key) {
                // The PDF metadata a document states outright is its metadata; \title and \author are
                // what the title block prints, and fill in only when it states none.
                case 'pdftitle': if (v) { this.metadata.title = v; this.pdfMetadata.add('title'); } break;
                case 'pdfauthor': if (v) { this.metadata.author = v; this.pdfMetadata.add('author'); } break;
                case 'pdfsubject': this.metadata.subject = v; break;
                case 'pdfkeywords': if (v) { this.metadata.keywords = v; this.pdfMetadata.add('keywords'); } break;
                case 'pdflang': this.metadata.language = value.trim(); break;
                case 'pdfcreationdate': { const d = parseLatexDate(value); if (d) this.metadata.created = d; break; }
                case 'pdfmoddate': { const d = parseLatexDate(value); if (d) this.metadata.modified = d; break; }
                case 'pdfinfo':
                    for (const [k, iv] of this.keyValues(value)) {
                        const text = this.plainText(iv);
                        if (k === 'Description') this.metadata.description = text;
                        else if (k === 'LastModifiedBy') this.metadata.lastModifiedBy = text;
                        else if (k === 'Title') this.metadata.title ??= text;
                        else if (k === 'Author') this.metadata.author ??= text;
                        else if (k) setOwn((this.metadata.customProperties ??= {}), k, text);
                    }
                    break;
            }
        }
    }

    /** Splits a `key=value, key={value}` list at top-level commas. */
    private keyValues(raw: string): [string, string][] {
        const out: [string, string][] = [];
        let depth = 0, cur = '';
        const push = () => {
            const eq = cur.indexOf('=');
            if (eq > 0) {
                let v = cur.slice(eq + 1).trim();
                if (v.startsWith('{') && v.endsWith('}')) v = v.slice(1, -1);
                out.push([cur.slice(0, eq).trim(), v]);
            }
            cur = '';
        };
        for (let i = 0; i < raw.length; i++) {
            const c = raw[i];
            if (c === '\\') { cur += c + (raw[++i] ?? ''); continue; }
            if (c === '{') depth++;
            if (c === '}') depth--;
            if (c === ',' && depth === 0) { push(); continue; }
            cur += c;
        }
        push();
        return out;
    }

    // ── headings ──

    private headingLevel(name: string): number {
        const rank = SECTIONING.indexOf(name);
        const top = Math.min(...[...this.sectionRanks].filter(r => r !== 0), rank);
        return Math.max(1, Math.min(6, rank - top + 1));
    }

    private headingNode(raw: string, level: number): OfficeContentNode {
        const flow = new Flow();
        this.withState(s => { s.inline = true; s.align = undefined; }, () => {
            if (this.nesting >= MAX_NESTING_DEPTH) this.tooDeep(flow, raw);
            else { this.nesting++; try { this.parseFlow(new Scanner(raw), flow, {}); } finally { this.nesting--; } }
            this.flushText(flow);
        });
        const children = this.trimInline(flow.inline);
        return { type: 'heading', text: textOf(children), children, metadata: { level } as HeadingMetadata };
    }

    private sectionNumbers: number[] = [];
    private floatNumbers: Record<string, number> = {};
    private equationNumber = 0;

    private heading(sc: Scanner, flow: Flow, name: string): void {
        const star = sc.readStar();
        sc.readRawOptional();
        const raw = sc.readRawGroup() ?? '';
        const level = this.headingLevel(name);
        const node = this.headingNode(raw, level);
        if (!star && this.numbered && name !== 'part' && level <= 3) {
            this.sectionNumbers = this.sectionNumbers.slice(0, level);
            while (this.sectionNumbers.length < level) this.sectionNumbers.push(0);
            this.sectionNumbers[level - 1]++;
            (node as any).__number = this.sectionNumbers.join('.');
        }
        this.addBlock(flow, node);
    }

    // ── includes and images ──

    private includeFile(sc: Scanner, flow: Flow, raw: string | null, subfile = false): void {
        void flow;
        const path = (raw ?? '').trim();
        if (!path) return;
        const resolved = this.project ? this.findProjectFile(path, ['', '.tex']) : null;
        if (!resolved) { this.missingFiles.add(path); return; }
        const open = sc.openFiles();
        if (open.includes(resolved) || open.length >= MAX_INCLUDE_DEPTH) {
            logWarning(OfficeWarningType.LATEX_EXPANSION_LIMIT_REACHED, this.config, { limit: 'file inclusion' });
            return;
        }
        let text = decodeTex(this.project!.files.get(resolved)!, this.inputEncoding).text.replace(/^﻿/, '').replace(/\r\n?/g, '\n');
        if (subfile) {
            // Looked up once each (a pattern tried every `\begin{document}` against the rest of the file).
            const start = text.indexOf('\\begin{document}');
            const end = start === -1 ? -1 : text.indexOf('\\end{document}', start + '\\begin{document}'.length);
            if (end !== -1) text = text.slice(start + '\\begin{document}'.length, end);
        }
        // Included text counts against the same budget as macro expansion, so including a large file
        // many times over (not only nesting includes) is bounded too.
        this.expandedChars += text.length;
        if (this.expandedChars > MAX_EXPANDED_CHARS) {
            if (!this.includeLimitHit) logWarning(OfficeWarningType.LATEX_EXPANSION_LIMIT_REACHED, this.config, { limit: 'file inclusion' });
            this.includeLimitHit = true;
            return;
        }
        sc.push(text + '\n', resolved);
    }

    private findProjectFile(path: string, extensions: string[], dirs: string[] = ['']): string | null {
        if (!this.project) return null;
        for (const dir of dirs) {
            for (const ext of extensions) {
                const p = normalizeProjectPath(this.mainDir, dir + path + ext);
                if (p !== null && this.project.files.has(p)) return p;
            }
        }
        return null;
    }

    /**
     * `filecontents`: compiling writes the body to a file (a file of that name already there is kept
     * unless the environment says `overwrite`), so the file joins the project for an `\input` or an
     * `\includegraphics` to read, as TeX would. The LaTeX generator carries images this way.
     */
    private fileContents(sc: Scanner, env: string): void {
        const options = (sc.readRawOptional() ?? '').split(',').map(o => o.trim());
        const name = (sc.readRawGroup() ?? '').trim();
        const body = sc.readRawEnvBody(env);
        const path = name ? normalizeProjectPath(this.mainDir, name) : null;
        if (path === null) return;
        // The file's lines are those after the \begin line, each written as TeX writes a line: without
        // trailing spaces, ending in LF.
        const lines = body.split('\n').slice(1, -1).map(line => trimEndChars(line, ' '));
        this.project ??= { files: new Map(), root: '' };
        if (!this.project.files.has(path) || options.includes('overwrite') || options.includes('force')) {
            this.project.files.set(path, Buffer.from(lines.map(line => line + '\n').join(''), 'utf8'));
        }
    }

    private image(sc: Scanner, flow: Flow): void {
        sc.readStar();
        const opts = sc.readRawOptional() ?? '';
        const path = (sc.readRawGroup() ?? '').trim().replace(/^"|"$/g, '');
        const meta: ImageMetadata = { attachmentName: '' };
        const kv = new Map(this.keyValues(opts));
        let width = kv.get('width');
        // The LaTeX generator's bounded form, `{\ifdim W>\linewidth\linewidth\else W\fi}` (natural width W,
        // capped at the line), is read as W.
        // (Each optional part carries the whitespace after it: two `\s*` with only something optional
        // between them tried every split of a run of spaces.)
        const bounded = width ? /^\s*(?:\{\s*)?\\ifdim\s*([-+]?[\d.]+(?:\s*[a-z]+)?)\s*>\s*\\linewidth\s*\\linewidth\s*\\else\s*\1\s*\\fi\s*(?:\}\s*)?$/.exec(width) : null;
        if (bounded) width = bounded[1];
        if (width) {
            const frac = /^\s*(?:([\d.]+)\s*)?\\(linewidth|textwidth|columnwidth|hsize)\s*$/.exec(width);
            if (frac) meta.width = `${Math.round((parseFloat(frac[1] || '1')) * 1000) / 10}%`;
            else { const pt = this.dimenPt(width); if (pt) meta.width = `${pt}pt`; }
        }
        const alt = kv.get('alt') ?? flow.pendingAlt;
        if (alt) meta.altText = this.plainText(alt);
        flow.pendingAlt = null;

        const resolved = this.project ? this.findProjectFile(path, IMAGE_EXTENSIONS, ['', ...this.graphicsPaths]) : null;
        if (resolved && this.config.extractAttachments) {
            let name = this.attachmentByPath.get(resolved);
            if (!name) {
                const bytes = this.project!.files.get(resolved)!;
                // A PDF that is only a picture (as the LaTeX generator carries images, or an image saved
                // as PDF) is taken as that picture, which every output format can show.
                const picture = /\.pdf$/i.test(resolved) ? imageFromPdf(bytes, this.decodeBudget) : null;
                name = resolved;
                if (picture) {
                    // Named for what it now is, and apart from any attachment already holding that name.
                    const stem = resolved.replace(/\.pdf$/i, ''), ext = picture.mimeType === 'image/png' ? '.png' : '.jpg';
                    name = stem + ext;
                    for (let n = 2; this.attachments.some(a => a.name === name); n++) name = `${stem}-${n}${ext}`;
                }
                this.attachments.push(createAttachment(name, picture ? Buffer.from(picture.data) : bytes));
                this.attachmentByPath.set(resolved, name);
            }
            meta.attachmentName = name;
        } else {
            if (!resolved) this.missingFiles.add(path);
            meta.url = path;
            delete (meta as any).attachmentName;
        }
        // A picture inside \href is that link, as a run of text inside it is.
        const link = this.state.link;
        if (link && (!link.internal || !this.config.ignoreInternalLinks)) {
            meta.link = link.url;
            meta.linkType = link.internal ? 'internal' : 'external';
        }
        this.addInline(flow, { type: 'image', metadata: meta });
    }

    // ── environments ──

    private beginEnv(sc: Scanner, flow: Flow, stop: Stop): StopReason | void {
        const env = (sc.readRawGroup() ?? '').trim();
        if (env === 'document') {
            // Everything before the body was preamble: definitions and metadata were taken from it,
            // any text it produced is not document content.
            flow.blocks = [];
            flow.inline = [];
            this.clearPending(flow);
            flow.nextAnchors = [];
            flow.nextComments = [];
            this.bodyStarted = true;
            return;
        }
        if (this.nesting >= MAX_NESTING_DEPTH) { this.tooDeep(flow, sc.readRawEnvBody(env)); return; }

        const user = this.envs.get(env);
        if (user) { this.beginUserEnv(sc, env, user); return; }
        const theorem = this.theoremDef(env);
        if (theorem) { this.theorem(sc, flow, env, theorem); return; }
        if (env === 'proof') { this.proof(sc, flow, env); return; }
        if (env === 'IEEEkeywords' || env === 'keywords' || env === 'keyword') { this.keywords(flow, sc.readRawEnvBody(env), env === 'IEEEkeywords'); return; }
        if (env === 'itemize' || env === 'enumerate' || env === 'compactitem' || env === 'compactenum' || env === 'inparaenum' || env === 'asparaenum') {
            sc.readRawOptional();
            this.endParagraph(flow);
            const parent = (flow as any).__list as ListContext | undefined;
            const ordered = env === 'enumerate' || env === 'compactenum' || env === 'inparaenum' || env === 'asparaenum';
            const ctx: ListContext = {
                listId: parent?.listId ?? `tex-list-${++this.listCounter}`,
                level: parent ? parent.level + 1 : 0,
                enumDepth: (parent?.enumDepth ?? 0) + (ordered ? 1 : 0),
            };
            return this.list(sc, flow, env, ordered, ctx);
        }
        if (env === 'description') { sc.readRawOptional(); return this.description(sc, flow, env); }
        if (env === 'thebibliography') { sc.readRawGroup(); return this.bibliography(sc, flow, env); }
        if (VERBATIM_ENVS.has(env)) { this.verbatim(sc, flow, env); return; }
        if (DISPLAY_MATH_ENVS.has(env)) {
            const body = sc.readRawEnvBody(env);
            // A numbered equation's labels resolve to its number for \ref/\eqref.
            if (!env.endsWith('*') && env !== 'displaymath' && env !== 'math') {
                const number = String(++this.equationNumber);
                for (const m of body.matchAll(/\\label\s*\{([^}]*)\}/g)) this.labelTargets.set(m[1].trim(), { type: 'code', __number: number } as any);
            }
            this.addBlockMath(flow, PLAIN_DISPLAY_MATH_ENVS.has(env) || env === 'math' || env === 'displaymath' ? body : `\\begin{${env}}${body}\\end{${env}}`);
            return;
        }
        if (own(TABLE_ENVS, env)) { this.table(sc, flow, env); return; }
        if (env === 'filecontents' || env === 'filecontents*') { this.fileContents(sc, env); return; }
        if (DRAWING_ENVS.has(env)) { sc.readRawEnvBody(env); this.unknown.add(`${env} environment`); return; }
        if (env === 'frame') return this.frame(sc, flow, stop);
        if (env === 'block' || env === 'alertblock' || env === 'exampleblock') {
            sc.readRawOptional('<', '>');
            const title = this.plainText(sc.readRawGroup());
            const type: AdmonitionMetadata['admonitionType'] = env === 'alertblock' ? 'warning' : env === 'exampleblock' ? 'tip' : 'note';
            const inner = this.subFlow(sc, env, flow);
            this.addBlock(flow, { type: 'admonition', children: inner, metadata: { admonitionType: type, ...(title ? { title } : {}) } as AdmonitionMetadata });
            return;
        }
        if (env === 'quote' || env === 'quotation' || env === 'verse') { this.quote(sc, flow, env); return; }
        if (env === 'abstract') {
            const inner = this.subFlow(sc, env, flow, f => { f.paragraphStyle = 'Abstract'; });
            // Keywords set inside the abstract (llncs) are the document's keywords, not part of its description.
            const text = inner.filter(b => !this.keywordBlocks.has(b)).map(b => b.text ?? '').join('\n').trim();
            if (text && !this.metadata.description) this.metadata.description = text;
            this.endParagraph(flow);
            for (const b of inner) this.pushBlock(flow, b);
            return;
        }
        if (env === 'center' || env === 'flushleft' || env === 'flushright') {
            this.endParagraph(flow);
            const align: TextAlignment = env === 'center' ? 'center' : env === 'flushright' ? 'right' : 'left';
            return this.withState(s => { s.align = align; }, () => this.envInto(sc, flow, env));
        }
        if (env === 'figure' || env === 'figure*' || env === 'table' || env === 'table*' || env === 'wrapfigure' || env === 'wraptable' || env === 'SCfigure') {
            if (env.startsWith('wrap')) this.readArgs(sc, 'oomom'.slice(0, 3));
            else sc.readRawOptional();
            this.float(sc, flow, env);
            return;
        }
        if (own(TRANSPARENT_ENVS, env)) {
            this.readArgs(sc, TRANSPARENT_ENVS[env]);
            return this.envInto(sc, flow, env);
        }
        // A language switch (babel's `otherlanguage`, polyglossia's `\begin{french}`, a `\babeltags` tag):
        // its text in place, within the paragraph around it.
        if (env === 'otherlanguage' || env === 'otherlanguage*' || this.languageName(env) || env === 'Arabic') {
            if (env.startsWith('otherlanguage')) sc.readRawGroup();
            else sc.readRawOptional();
            this.nesting++;
            try { return this.parseFlow(sc, flow, { env }) === 'enddoc' ? 'enddoc' : undefined; } finally { this.nesting--; }
        }
        this.unknown.add(`${env} environment`);
        return this.envInto(sc, flow, env);
    }

    /**
     * Content nested past {@link MAX_NESTING_DEPTH} is kept as its plain text (commands and braces
     * removed) rather than interpreted, so a hostile nesting depth costs no recursion.
     */
    private tooDeep(flow: Flow, raw: string): void {
        if (!this.nestingLimitHit) {
            this.nestingLimitHit = true;
            logWarning(OfficeWarningType.LATEX_EXPANSION_LIMIT_REACHED, this.config, { limit: 'nesting depth' });
        }
        this.addText(flow, raw.replace(/\\(begin|end)\s*\{[^{}]*\}/g, ' ').replace(/\\[A-Za-z@]+\*?|\\.|[{}]/g, ' ').replace(/\s+/g, ' '));
    }

    /** Reads an environment's body into the current flow. */
    private envInto(sc: Scanner, flow: Flow, env: string): StopReason | void {
        this.nesting++;
        try {
            const r = this.parseFlow(sc, flow, { env });
            // An environment ends the paragraph inside it, so a paragraph takes the alignment the
            // environment set (`center`, `flushright`) rather than the state outside it.
            this.endParagraph(flow);
            if (r === 'enddoc') return r;
        } finally { this.nesting--; }
    }

    /** Reads an environment's body as a separate container, returning its blocks. */
    private subFlow(sc: Scanner, env: string, parent: Flow, patch?: (f: Flow) => void): OfficeContentNode[] {
        this.endParagraph(parent);
        const flow = new Flow();
        (flow as any).__list = (parent as any).__list;
        patch?.(flow);
        this.nesting++;
        try {
            this.withState(() => { }, () => { this.parseFlow(sc, flow, { env }); this.endParagraph(flow); });
        } finally { this.nesting--; }
        return flow.blocks;
    }

    // ── theorems ──

    /** Defines theorem-like environment `env` titled `title`, numbered with `counter` (none: unnumbered). */
    private defineTheorem(env: string, title: string, counter: string | undefined, within?: string, style?: string): void {
        if (!env) return;
        this.theorems.set(env, { title: this.plainText(title), style: style || this.theoremStyle, counter });
        const c = counter ? this.theoremCounters.get(counter) : undefined;
        if (counter && !c) this.theoremCounters.set(counter, { value: 0, within, prefix: '' });
        else if (c && counter === env && within) c.within = within;
    }

    /**
     * The theorem-like environment `env` is, if it is one: defined by the document, provided by its
     * class (Springer's classes number each with its own counter; beamer's are blocks), or a name that
     * is theorem-like in any document (a theorem defined in a package out of sight), which is headed
     * but unnumbered, since its numbering is unknown.
     */
    private theoremDef(env: string): TheoremDef | undefined {
        const def = this.theorems.get(env);
        if (def || !own(DEFAULT_THEOREM_TITLES, env)) return def;
        const springer = SPRINGER_CLASSES.has(this.native.documentClass);
        if (!(springer || (this.beamer ? BEAMER_THEOREMS.has(env) : GENERIC_THEOREMS.has(env)))) return undefined;
        const provided: TheoremDef = { title: DEFAULT_THEOREM_TITLES[env], style: own(DEFAULT_THEOREM_STYLES, env) ? DEFAULT_THEOREM_STYLES[env] : 'plain', counter: springer ? env : undefined };
        this.theorems.set(env, provided);
        if (springer && !this.theoremCounters.has(env)) this.theoremCounters.set(env, { value: 0, prefix: '' });
        return provided;
    }

    /** The number of the section (or other sectioning unit) `unit` the reading is in, for counters numbered within it. */
    private unitNumber(unit: string): string {
        const rank = SECTIONING.indexOf(unit);
        if (rank < 0) return '';
        const level = this.headingLevel(unit);
        const parts = this.sectionNumbers.slice(0, level);
        while (parts.length < level) parts.push(0);
        return parts.join('.');
    }

    /**
     * A theorem-like environment: its body, headed by its title, number and note ("**Theorem 2.1**
     * (Note)**.**") as amsthm sets it, the body italic in the `plain` style. A `\ref` to a label in it
     * reads its number. Under beamer it is a titled block, as beamer sets it.
     */
    private theorem(sc: Scanner, flow: Flow, env: string, def: TheoremDef): void {
        const note = sc.readRawOptional();
        // Every theorem of a kind is headed by its title: after the first, each is charged to the
        // document's repeated-content budget (see repeatUtils), past which it is headed by the start
        // of it. One 100 KB title on 2,000 theorems made 200 MB of text.
        let title = def.title;
        if (!this.theoremsHeaded.has(def)) this.theoremsHeaded.add(def);
        else if (takeRepeats(this.config, 1, title.length) === 0) title = repeatPreview(title);
        this.endParagraph(flow);
        let number: string | undefined;
        const counter = def.counter ? this.theoremCounters.get(def.counter) : undefined;
        if (counter && !this.beamer) {
            const prefix = counter.within ? this.unitNumber(counter.within) : '';
            if (prefix !== counter.prefix) { counter.value = 0; counter.prefix = prefix; }
            number = `${prefix ? `${prefix}.` : ''}${++counter.value}`;
        }
        const noteRuns = note !== null && note.trim() ? this.headingNode(note, 1).children ?? [] : [];
        if (this.beamer) {
            const inner = this.subFlow(sc, env, flow);
            const blockTitle = `${title}${noteRuns.length ? ` (${textOf(noteRuns)})` : ''}`;
            const type: AdmonitionMetadata['admonitionType'] = env.startsWith('example') ? 'tip' : 'note';
            this.addBlock(flow, { type: 'admonition', children: inner, metadata: { admonitionType: type, title: blockTitle } as AdmonitionMetadata });
            return;
        }
        const labelsBefore = this.labelLog.length;
        const inner = this.withState(s => { if (def.style === 'plain') s.fmt.italic = true; else delete s.fmt.italic; }, () => this.subFlow(sc, env, flow));
        const head: TextFormatting = def.style === 'remark' ? { italic: true } : { bold: true };
        const runs: OfficeContentNode[] = [{ type: 'text', text: number ? `${title} ${number}` : title, formatting: head }];
        if (noteRuns.length) runs.push({ type: 'text', text: ' (' }, ...noteRuns, { type: 'text', text: ')' });
        runs.push({ type: 'text', text: '.', formatting: head });
        const first = this.headInto(inner, runs);
        if (number) {
            // Labels in the theorem name it: a `\ref` to one reads the theorem's number.
            // A label on something numbered of its own inside it (an equation, a nested theorem) keeps that.
            (first as any).__number = number;
            for (const id of this.labelLog.slice(labelsBefore)) {
                const t = this.labelTargets.get(id);
                if (!t || (t.type === 'paragraph' && !(t as any).__number)) this.labelTargets.set(id, first);
            }
        }
        for (const b of inner) this.pushBlock(flow, b);
    }

    /** amsthm's proof: "*Proof.*" (or its optional heading) before the body, and the end-of-proof mark □ after it. */
    private proof(sc: Scanner, flow: Flow, env: string): void {
        const heading = sc.readRawOptional();
        this.endParagraph(flow);
        this.qedPlaced = false;
        const inner = this.withState(s => { delete s.fmt.italic; }, () => this.subFlow(sc, env, flow));
        const runs = heading !== null
            ? this.withState(s => { s.fmt.italic = true; }, () => this.headingNode(heading, 1).children ?? [])
            : [{ type: 'text', text: 'Proof', formatting: { italic: true } } as OfficeContentNode];
        // amsthm adds the period unless the heading ends with punctuation.
        if (!/[.!?:]\s*$/.test(textOf(runs))) runs.push({ type: 'text', text: '.', formatting: { italic: true } });
        this.headInto(inner, runs);
        if (!this.qedPlaced) {
            const last = inner[inner.length - 1];
            if (last?.type === 'paragraph') {
                last.children = [...(last.children ?? []), { type: 'text', text: ' □' }];
                last.text = textOf(last.children);
            } else inner.push({ type: 'paragraph', text: '□', children: [{ type: 'text', text: '□' }], metadata: { alignment: 'right' } as ParagraphMetadata });
        }
        this.qedPlaced = false;
        for (const b of inner) this.pushBlock(flow, b);
    }

    /** Puts heading runs at the start of an environment's first paragraph (or in a paragraph of their own before it), returning that paragraph. */
    private headInto(blocks: OfficeContentNode[], runs: OfficeContentNode[]): OfficeContentNode {
        const first = blocks[0];
        // A nested theorem or proof opening the body keeps its own head, on its own line, as amsthm sets it.
        if (first?.type === 'paragraph' && !(first.metadata as ParagraphMetadata | undefined)?.style && !this.theoremHeads.has(first)) {
            first.children = [...runs, { type: 'text', text: ' ' }, ...(first.children ?? [])];
            first.text = textOf(first.children);
            this.theoremHeads.add(first);
            return first;
        }
        const para: OfficeContentNode = { type: 'paragraph', text: textOf(runs), children: runs, metadata: {} as ParagraphMetadata };
        blocks.unshift(para);
        this.theoremHeads.add(para);
        return para;
    }

    /**
     * Keywords (`\keywords`, and the `keywords`/`keyword` and IEEEtran `IEEEkeywords` environments): the
     * document's keywords metadata unless `\hypersetup` states them, printed where they stand as the
     * classes print them ("Keywords: a, b"; "Index Terms: a, b" for IEEEtran).
     */
    private keywords(flow: Flow, raw: string | null, ieee: boolean): void {
        if (raw === null) return;
        // llncs separates keywords with `\and`, elsarticle with `\sep`.
        // A match starts where whitespace does (not at each space of a run, each read to its end).
        const source = raw.replace(/(?<!\s)\s*\\(and|sep)\b\s*/g, ', ');
        // Parsed once, for the text printed and the metadata alike: parsed for each, keywords within
        // keywords took twice as long for each level.
        this.endParagraph(flow);
        const body = this.parseBlocksOf(source, undefined, false);
        const text = body.map(b => b.text ?? textOf(b.children)).join(' ').replace(/\s+/g, ' ').trim();
        if (!text) return;
        if (!this.pdfMetadata.has('keywords')) this.metadata.keywords = text;
        const label: OfficeContentNode = ieee
            ? { type: 'text', text: 'Index Terms: ', formatting: { bold: true, italic: true } }
            : { type: 'text', text: 'Keywords: ', formatting: { bold: true } };
        for (const b of body) this.keywordBlocks.add(b);
        const first = body[0];
        if (first?.type === 'paragraph') {
            first.children = [label, ...(first.children ?? [])];
            first.text = textOf(first.children);
        } else {
            const head: OfficeContentNode = { type: 'paragraph', text: label.text, children: [label], metadata: {} as ParagraphMetadata };
            this.keywordBlocks.add(head);
            body.unshift(head);
        }
        for (const b of body) this.pushBlock(flow, b);
    }

    /**
     * What `\today` prints: `texParserConfig.today` when set, else the date of the parse, as LaTeX
     * prints the date of the compile, written the way the document's language writes a date
     * ("September 25, 2026" in English).
     */
    private todayText(): string {
        const fixed = this.config.texParserConfig?.today;
        if (typeof fixed === 'string' && fixed) return fixed;
        const style: Intl.DateTimeFormatOptions = { year: 'numeric', month: 'long', day: 'numeric' };
        const now = new Date();
        try { return new Intl.DateTimeFormat(this.documentLanguage() ?? 'en-US', style).format(now); }
        catch { return new Intl.DateTimeFormat('en-US', style).format(now); }
    }

    // ── languages ──

    /** The babel/polyglossia language `name` (or `\babeltags` tag) names, if it names one. */
    private languageName(name: string): string | undefined {
        if (own(LANGUAGE_CODES, name)) return name;
        const tagged = this.languageTags.get(name);
        return tagged && own(LANGUAGE_CODES, tagged) ? tagged : undefined;
    }

    /**
     * The document's main language as a BCP 47 tag: polyglossia's `\setdefaultlanguage` or babel's
     * `main=` option, else the last language babel loads (from its options after the class's, as babel does).
     */
    private documentLanguage(): string | undefined {
        const pick = (name: string | undefined) => (name && own(LANGUAGE_CODES, name) ? LANGUAGE_CODES[name] : undefined);
        if (this.mainLanguage) return pick(this.mainLanguage);
        if (!(this.native.packages as string[]).includes('babel')) return undefined;
        const options = [...((this.native.classOptions ?? []) as string[]), ...this.babelLanguages];
        const main = options.map(o => /^main\s*=\s*(.+)$/.exec(o)?.[1]?.trim()).filter(Boolean).pop();
        return pick(main ?? options.filter(o => own(LANGUAGE_CODES, o)).pop());
    }

    // ── plain TeX and ConTeXt ──

    /** Skips plain TeX glue after `\vskip`/`\hskip`/`\kern`: a dimension with optional `plus` and `minus` parts. */
    private readGlue(sc: Scanner): void {
        this.skipDimen(sc);
        this.readGlueTail(sc);
    }

    /** Skips the `plus` and `minus` parts of glue after its natural size. */
    private readGlueTail(sc: Scanner): void {
        for (const word of ['plus', 'minus']) {
            const snap = sc.save();
            sc.skipBlanks();
            let w = '';
            while (/[a-z]/.test(sc.peekCh() ?? '') && w.length < word.length) w += sc.nextCh();
            if (w === word) this.skipDimen(sc);
            else sc.restore(snap);
        }
    }

    /** Skips an (optionally `=`-prefixed) dimension or number: digits and a unit, or a register or `\magstep<n>`. */
    private skipDimen(sc: Scanner): void {
        sc.skipBlanks();
        if (sc.peekCh() === '=') sc.nextCh();
        sc.skipBlanks();
        const register = () => {
            const t = sc.next(this.atLetter);
            if (t.t === 'cs' && t.name.startsWith('magstep')) while (/[\d.]/.test(sc.peekCh() ?? '')) sc.nextCh();
        };
        if (sc.peekCh() === '\\') { register(); return; }
        let n = 0;
        while (/[-+\d.,]/.test(sc.peekCh() ?? '') && n++ < 32) sc.nextCh();
        sc.skipBlanks();
        // A unit, including `fil`/`fill`/`filll` and `true` units; a register (`2\baselineskip`) ends the amount.
        if (sc.peekCh() === '\\') { register(); return; }
        const snap = sc.save();
        let unit = '';
        while (/[a-z]/.test(sc.peekCh() ?? '') && unit.length < 12) unit += sc.nextCh();
        if (!/^(true)?(pt|pc|in|bp|cm|mm|dd|cc|sp|em|ex|mu|px|fil+)$/.test(unit)) sc.restore(snap);
    }

    /**
     * ConTeXt's commands, read for a ConTeXt document's text: `\start<name>`/`\stop<name>` pairs as
     * the LaTeX environments they match (sectioning, itemize, typing, formula, quotation) or else as
     * transparent ones, and its `\setup...`/`\use...`/`\define...` configuration skipped. Returns
     * false for a command it does not know.
     */
    private contextCommand(sc: Scanner, flow: Flow, name: string): boolean {
        const options = () => { while (sc.nextIs('[')) sc.readRawOptional(); };
        if (/^(setup|use|define|enable|disable)[A-Za-z]+$/.test(name) || name === 'mainlanguage' || name === 'language') {
            options();
            while (sc.nextIs('{')) sc.readRawGroup();
            return true;
        }
        const start = /^start([A-Za-z]+)$/.exec(name);
        const stop = /^stop([A-Za-z]+)$/.exec(name);
        const what = (start ?? stop)?.[1];
        if (!what) return false;
        const sectioning: Record<string, string> = { part: 'part', chapter: 'chapter', section: 'section', subsection: 'subsection', subsubsection: 'subsubsection', title: 'chapter*', subject: 'section*', subsubject: 'subsection*', subsubsubject: 'subsubsection*' };
        if (own(sectioning, what)) {
            if (stop) return true;
            const opts = sc.readRawOptional() ?? '';
            const kv = new Map(this.keyValues(opts));
            const title = kv.get('title') ?? '';
            const [cmd, star] = sectioning[what].endsWith('*') ? [sectioning[what].slice(0, -1), '*'] : [sectioning[what], ''];
            this.expand(sc, `\\${cmd}${star}{${title}}${kv.get('reference') ? `\\label{${kv.get('reference')}}` : ''}`);
            return true;
        }
        if (what === 'typing') { sc.readRawOptional(); this.expand(sc, `\\begin{verbatim}${sc.readRawUntil('\\stoptyping').text}\\end{verbatim}`); return true; }
        if (what === 'formula') { sc.readRawOptional(); this.addBlockMath(flow, sc.readRawUntil('\\stopformula').text); return true; }
        if (what === 'itemize') {
            if (stop) { const env = this.contextLists.pop(); if (env) this.expand(sc, `\\end{${env}}`); return true; }
            const opts = (sc.readRawOptional() ?? '').split(',').map(o => o.trim());
            const env = opts.some(o => /^(n|a|A|r|R|numbers?|characters|Characters|romannumerals|Romannumerals)$/.test(o)) ? 'enumerate' : 'itemize';
            this.contextLists.push(env);
            this.expand(sc, `\\begin{${env}}`);
            return true;
        }
        if (what === 'quotation' || what === 'quote' || what === 'blockquote') { if (start) options(); this.expand(sc, start ? '\\begin{quote}' : '\\end{quote}'); return true; }
        if (start) options();
        return true;
    }

    private verbatim(sc: Scanner, flow: Flow, env: string): void {
        let language: string | undefined;
        if (env === 'lstlisting') {
            const opts = sc.readRawOptional('[', ']');
            const lang = opts ? /language\s*=\s*\{?([^,}]*)\}?/.exec(opts)?.[1]?.trim() : undefined;
            if (lang) language = LISTINGS_LANGUAGE_NAMES.get(lang.toLowerCase()) ?? lang.toLowerCase();
        } else if (env === 'minted') {
            sc.readRawOptional();
            language = (sc.readRawGroup() ?? '').trim().toLowerCase() || undefined;
        } else if (/^(Verbatim|BVerbatim|LVerbatim)/.test(env)) {
            sc.readRawOptional();
        }
        const { text } = sc.readRawUntil(`\\end{${env}}`);
        if (env === 'comment') return;
        const code = text.replace(/^[ \t]*\n/, '').replace(/\n[ \t]*$/, '');
        const meta: CodeMetadata = {};
        if (language) meta.language = language;
        this.addBlock(flow, { type: 'code', text: code, metadata: meta });
    }

    private quote(sc: Scanner, flow: Flow, env: string): void {
        const inner = this.subFlow(sc, env, flow, f => { f.paragraphStyle = 'Quote'; });
        // A quote that opens with a bold coloured label naming an admonition type (the shape the
        // LaTeX generator writes admonitions in) is that admonition.
        const first = inner[0];
        const label = first?.type === 'paragraph' && first.children?.length === 1 && first.children[0].type === 'text' ? first.children[0] : null;
        if (label?.formatting?.bold && label.formatting.color) {
            const color = label.formatting.color.replace('#', '').toUpperCase();
            const byColor = Object.entries(ADMONITION_COLOR).find(([, c]) => c.toUpperCase() === color)?.[0];
            const byName = (label.text ?? '').trim().toLowerCase();
            const type = (byColor ?? (own(ADMONITION_COLOR, byName) ? byName : undefined)) as AdmonitionMetadata['admonitionType'] | undefined;
            if (type) {
                const title = (label.text ?? '').trim();
                const body = inner.slice(1).map(b => { if (b.type === 'paragraph') delete (b.metadata as any)?.style; return b; });
                const meta: AdmonitionMetadata = { admonitionType: type };
                if (title && title.toLowerCase() !== type) meta.title = title;
                this.addBlock(flow, { type: 'admonition', children: body, metadata: meta });
                return;
            }
        }
        for (const b of inner) this.pushBlock(flow, b);
    }

    private float(sc: Scanner, flow: Flow, env: string): void {
        const labelsBefore = this.labelLog.length;
        const inner = this.subFlow(sc, env, flow);
        // Labels in a float name its figure or table, wherever in the float they were written.
        const main = inner.find(b => b.type === 'image' || b.type === 'table');
        const kind = env.startsWith('table') || env === 'wraptable' ? 'table' : 'figure';
        const caption = inner.find(b => (b.metadata as any)?.style === 'Caption');
        if (caption) {
            const n = String(this.floatNumbers[kind] = (this.floatNumbers[kind] ?? 0) + 1);
            if (main) (main as any).__number = n;
            else {
                // A float with nothing the AST can hold (a TikZ drawing, which is omitted) keeps its
                // caption, which then carries the float's number: a label in the float reads it.
                (caption as any).__number = n;
                for (const id of this.labelLog.slice(labelsBefore)) {
                    const t = this.labelTargets.get(id);
                    if (!t || (inner.includes(t) && (t.metadata as any)?.style === 'Caption') || (t.type === 'paragraph' && !(t as any).__number && !t.text)) this.labelTargets.set(id, caption);
                }
            }
        }
        if (main) {
            const ids: string[] = [];
            for (const b of inner) {
                if (b === main) continue;
                const a = (b.metadata as any)?.anchorIds as string[] | undefined;
                if (a?.length && (b.metadata as any)?.style === 'Caption') { ids.push(...a); delete (b.metadata as any).anchorIds; }
            }
            if (ids.length) {
                const meta = (main.metadata ??= {} as any) as any;
                meta.anchorIds = [...(meta.anchorIds ?? []), ...ids];
                for (const id of ids) this.labelTargets.set(id, main);
            }
            // Without internal links there are no anchors to move, but a label on the caption still names the float.
            for (const id of this.labelLog.slice(labelsBefore)) {
                const t = this.labelTargets.get(id);
                if (t && t !== main && inner.includes(t) && (t.metadata as any)?.style === 'Caption') this.labelTargets.set(id, main);
            }
        }
        for (const b of inner) this.pushBlock(flow, b);
    }

    // ── lists ──

    private list(sc: Scanner, flow: Flow, env: string, ordered: boolean, ctx: ListContext): StopReason | void {
        let index = 0;
        // Anything before the first \item (typically \setcounter{enumi}{N}).
        const pre = new Flow();
        (pre as any).__list = ctx;
        let r = this.parseFlow(sc, pre, { env, items: 'item' });
        if (ctx.nextIndex !== undefined) { index = ctx.nextIndex; ctx.nextIndex = undefined; }
        while (r === 'item') {
            sc.next(this.atLetter); // \item
            sc.readRawOptional('<', '>');
            const label = sc.readRawOptional();
            const itemFlow = new Flow();
            (itemFlow as any).__list = ctx;
            const labelsBefore = this.labelLog.length;
            const numbered = ordered && label === null;
            if (numbered) this.enumNumbers[ctx.enumDepth - 1] = index + 1;
            this.nesting++;
            try { r = this.withState(() => { }, () => this.parseFlow(sc, itemFlow, { env, items: 'item' })); } finally { this.nesting--; }
            this.endParagraph(itemFlow);
            if (numbered) {
                // A label in a numbered item reads as the item's reference ("2", "2a", "2(a)iii"), as in LaTeX.
                const n = this.enumNumbers;
                const d = ctx.enumDepth;
                const number = d <= 1 ? `${n[0]}` : d === 2 ? `${n[0]}${alph(n[1])}`
                    : `${n[0]}(${alph(n[1])})${roman(n[2])}${d >= 4 ? alph(n[3]).toUpperCase() : ''}`;
                for (const id of this.labelLog.slice(labelsBefore)) {
                    const t = this.labelTargets.get(id);
                    if (!t || (t.type === 'paragraph' && !(t as any).__number)) this.labelTargets.set(id, { type: 'list', __number: number } as any);
                }
            }
            const blocks = itemFlow.blocks;
            let isTask: boolean | undefined, checked: boolean | undefined;
            let prefix: OfficeContentNode[] = [];
            if (label !== null) {
                const l = label.trim();
                if (/\\boxtimes|☒|☑|\\checkmark|✓|✔|\[x\]/i.test(l)) { isTask = true; checked = true; }
                else if (/\\square|☐|□|\[ \]/.test(l)) { isTask = true; checked = false; }
                else if (l) {
                    const t = this.plainText(l);
                    if (t) prefix = [{ type: 'text', text: `${t} ` }];
                }
            }
            // Leading paragraphs are the item's text; anything after them (a nested list, a
            // table) follows the item as its own blocks.
            const paras: OfficeContentNode[] = [];
            let k = 0;
            while (k < blocks.length && blocks[k].type === 'paragraph') paras.push(blocks[k++]);
            const rest = blocks.slice(k);
            const children: OfficeContentNode[] = [...prefix];
            paras.forEach((p, pi) => {
                if (pi > 0) children.push({ type: 'break', metadata: { breakType: 'carriageReturn' } as any });
                children.push(...(p.children ?? []));
            });
            const anchorIds = paras.flatMap(p => ((p.metadata as any)?.anchorIds as string[] | undefined) ?? []);
            const comments = paras.flatMap(p => p.comments ?? []);
            // An empty item with an empty label is scaffolding for a deeper list, not an item.
            if (children.length || label === null || isTask) {
                const meta: ListMetadata = {
                    listType: ordered ? 'ordered' : 'unordered',
                    indentation: ctx.level,
                    alignment: 'left',
                    listId: ctx.listId,
                    itemIndex: index++,
                };
                if (isTask) { meta.isTask = true; meta.checked = checked; }
                if (anchorIds.length) meta.anchorIds = anchorIds;
                const node: OfficeContentNode = { type: 'list', text: textOf(children), children, metadata: meta };
                if (comments.length) node.comments = comments;
                this.pushBlock(flow, node);
            }
            for (const b of rest) this.pushBlock(flow, b);
            if (ctx.nextIndex !== undefined) { index = ctx.nextIndex; ctx.nextIndex = undefined; }
        }
        if (r === 'enddoc') return r;
    }

    private description(sc: Scanner, flow: Flow, env: string): StopReason | void {
        this.endParagraph(flow);
        const children: OfficeContentNode[] = [];
        let r = this.parseFlow(sc, new Flow(), { env, items: 'item' });
        while (r === 'item') {
            sc.next(this.atLetter);
            const label = sc.readRawOptional();
            const itemFlow = new Flow();
            this.nesting++;
            try { r = this.withState(() => { }, () => this.parseFlow(sc, itemFlow, { env, items: 'item' })); } finally { this.nesting--; }
            this.endParagraph(itemFlow);
            if (label !== null && label.trim()) {
                const term = this.headingNode(label.replace(/^\{([\s\S]*)\}$/, '$1'), 1);
                children.push({ type: 'definitionTerm', text: term.text, children: term.children });
            }
            const desc = itemFlow.blocks.flatMap((b, bi) => (b.type === 'paragraph' ? [...(bi > 0 ? [{ type: 'break', metadata: { breakType: 'carriageReturn' } } as OfficeContentNode] : []), ...(b.children ?? [])] : []));
            if (desc.length) children.push({ type: 'definitionDescription', text: textOf(desc), children: desc });
        }
        if (children.length) this.pushBlock(flow, { type: 'definitionList', children });
        if (r === 'enddoc') return r;
    }

    private bibliography(sc: Scanner, flow: Flow, env: string): StopReason | void {
        this.endParagraph(flow);
        this.pushBlock(flow, { type: 'heading', text: 'References', children: [{ type: 'text', text: 'References' }], metadata: { level: 1 } as HeadingMetadata });
        const listId = `tex-list-${++this.listCounter}`;
        let index = 0;
        let r = this.parseFlow(sc, new Flow(), { env, items: 'bibitem' });
        while (r === 'bibitem' as any || r === 'item') {
            sc.next(this.atLetter);
            sc.readRawOptional();
            const key = (sc.readRawGroup() ?? '').trim();
            const itemFlow = new Flow();
            r = this.parseFlow(sc, itemFlow, { env, items: 'bibitem' });
            this.endParagraph(itemFlow);
            const children = itemFlow.blocks.flatMap(b => b.children ?? []);
            const meta: ListMetadata = { listType: 'ordered', indentation: 0, alignment: 'left', listId, itemIndex: index++ };
            if (key && !this.config.ignoreInternalLinks) meta.anchorIds = [key];
            this.pushBlock(flow, { type: 'list', text: textOf(children), children, metadata: meta });
        }
        if (r === 'enddoc') return r;
    }

    // ── beamer ──

    private frame(sc: Scanner, flow: Flow, stop: Stop): StopReason | void {
        void stop;
        sc.readRawOptional('<', '>');
        sc.readRawOptional();
        sc.readRawOptional('<', '>');
        let title: string | null = null;
        let subtitle: string | null = null;
        if (sc.nextIs('{')) {
            title = sc.readRawGroup();
            if (sc.nextIs('{')) subtitle = sc.readRawGroup();
        }
        this.endParagraph(flow);
        const slideNumber = ++this.slideCounter;
        const slide: OfficeContentNode = { type: 'slide', children: [], metadata: { slideNumber } as any };
        const inner = new Flow();
        (inner as any).__slide = slide;
        if (title !== null && title.trim()) inner.blocks.push(this.headingNode(title, 1));
        if (subtitle !== null && subtitle.trim()) inner.blocks.push(this.headingNode(subtitle, 2));
        this.nesting++;
        let r: StopReason;
        try { r = this.withState(() => { }, () => this.parseFlow(sc, inner, { env: 'frame' })); } finally { this.nesting--; }
        this.endParagraph(inner);
        slide.children = inner.blocks;
        if ((inner as any).__titlepage && !inner.blocks.length) { this.slideCounter--; return r === 'enddoc' ? r : undefined; }
        if (!this.config.ignoreNotes && slide.notes?.length === 0) delete slide.notes;
        this.pushBlock(flow, slide);
        this.lastSlide = slide;
        if (r === 'enddoc') return r;
    }

    // ── tables ──

    private table(sc: Scanner, flow: Flow, env: string): void {
        const pre = TABLE_ENVS[env];
        this.readArgs(sc, pre);
        const spec = sc.readRawGroup() ?? '';
        const body = sc.readRawEnvBody(env);
        // A table nested in a cell costs far more stack per level than a group, so tables have their
        // own, tighter depth bound (well past any real document), the same on every JavaScript engine.
        if (this.tableDepth >= MAX_TABLE_NESTING_DEPTH) { this.tooDeep(flow, body); return; }
        const aligns = this.columnAligns(spec);
        this.endParagraph(flow);
        this.tableDepth++;
        let node: ReturnType<LatexReader['buildTable']>;
        try { node = this.buildTable(body, aligns, env); } finally { this.tableDepth--; }
        if (node) this.pushBlock(flow, node.table);
        if (node?.caption.length) for (const c of node.caption) this.pushBlock(flow, c);
    }

    /** Column alignments from a column specification (`l`, `c`, `r`, `p{}` with `>{\centering}`, `*{n}{...}`). */
    private columnAligns(spec: string): (('left' | 'center' | 'right') | undefined)[] {
        let s = spec;
        for (let guard = 0; guard < 8 && /\*\{\s*(\d+)\s*\}\{/.test(s); guard++) {
            // `*{n}{spec}` repeats `spec`; the repetition is bounded by column count and by length, since
            // a long `spec` repeated 1000 times would otherwise allocate a huge string.
            s = s.replace(/\*\{\s*(\d+)\s*\}\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}/, (_m, n, inner: string) =>
                inner.repeat(Math.max(0, Math.min(1000, +n, Math.floor(MAX_COLUMN_SPEC_CHARS / Math.max(1, inner.length))))));
            if (s.length > MAX_COLUMN_SPEC_CHARS) s = s.slice(0, MAX_COLUMN_SPEC_CHARS);
        }
        const out: (('left' | 'center' | 'right') | undefined)[] = [];
        let pendingAlign: 'left' | 'center' | 'right' | undefined;
        for (let i = 0; i < s.length; i++) {
            const c = s[i];
            if (c === '>' && s[i + 1] === '{') {
                const end = this.matchBrace(s, i + 1);
                const inner = s.slice(i + 2, end);
                if (/\\centering/.test(inner)) pendingAlign = 'center';
                else if (/\\raggedleft/.test(inner)) pendingAlign = 'right';
                else if (/\\raggedright/.test(inner)) pendingAlign = 'left';
                i = end;
                continue;
            }
            if ((c === '<' || c === '@' || c === '!') && s[i + 1] === '{') { i = this.matchBrace(s, i + 1); continue; }
            if (c === '{') { i = this.matchBrace(s, i); continue; }
            if ('lcrpmbXLCRJS'.includes(c)) {
                const a = c === 'c' || c === 'C' ? 'center' : c === 'r' || c === 'R' ? 'right' : c === 'l' || c === 'L' ? 'left' : undefined;
                out.push(pendingAlign ?? a);
                pendingAlign = undefined;
                if ('pmbX'.includes(c) && s[i + 1] === '{') i = this.matchBrace(s, i + 1);
            }
        }
        return out;
    }

    private matchBrace(s: string, open: number): number {
        let depth = 0;
        for (let i = open; i < s.length; i++) {
            if (s[i] === '\\') { i++; continue; }
            if (s[i] === '{') depth++;
            else if (s[i] === '}' && --depth === 0) return i;
        }
        return s.length - 1;
    }

    /** Splits a table body at top-level `&` and `\\`, skipping nested groups, environments and math. */
    private splitTable(body: string): string[][] {
        const rows: string[][] = [];
        let cells: string[] = [];
        let cur = '';
        let depth = 0, envDepth = 0, inMath = false;
        for (let i = 0; i < body.length; i++) {
            const c = body[i];
            if (c === '\\') {
                const rest = body.slice(i);
                const word = /^\\([A-Za-z]+|.)/.exec(rest)?.[1] ?? '';
                if (word === 'begin') envDepth++;
                if (word === 'end') envDepth--;
                if (depth === 0 && envDepth === 0 && !inMath && (word === '\\' || word === 'tabularnewline')) {
                    cells.push(cur); rows.push(cells); cells = []; cur = '';
                    i += word.length;
                    // `\\*` and `\\[len]`
                    if (body[i + 1] === '*') i++;
                    if (body[i + 1] === '[') { const e = body.indexOf(']', i + 1); if (e > 0) i = e; }
                    continue;
                }
                cur += '\\' + word;
                i += word.length;
                continue;
            }
            if (c === '%') { const e = body.indexOf('\n', i); cur += body.slice(i, e < 0 ? body.length : e); i = e < 0 ? body.length : e - 1; continue; }
            if (c === '$') inMath = !inMath;
            if (c === '{') depth++;
            if (c === '}') depth--;
            if (c === '&' && depth === 0 && envDepth === 0 && !inMath) { cells.push(cur); cur = ''; continue; }
            cur += c;
        }
        cells.push(cur);
        rows.push(cells);
        return rows;
    }

    private static RULES = /^\s*(?:\\(?:hline|toprule|midrule|bottomrule|endhead|endfirsthead|endfoot|endlastfoot|hhline\s*\{[^}]*\}|cline\s*\{[^}]*\}|cmidrule\s*(?:\([^)]*\)\s*)?(?:\[[^\]]*\]\s*)?\{[^}]*\}|addlinespace(?:\s*\[[^\]]*\])?|noalign\s*\{[^}]*\}|rowcolor\s*(?:\[[^\]]*\]\s*)?\{[^}]*\}|specialrule\s*\{[^}]*\}\{[^}]*\}\{[^}]*\})(?:\[[^\]]*\])?\s*)+/;

    private buildTable(body: string, aligns: (('left' | 'center' | 'right') | undefined)[], env: string): { table: OfficeContentNode; caption: OfficeContentNode[] } | null {
        void env;
        const raw = this.splitTable(body);
        type RawRow = { cells: string[]; header?: boolean; rowColor?: string };
        const rows: RawRow[] = [];
        const caption: OfficeContentNode[] = [];
        let seenFirstHead = false, seenHead = false, seenFoot = false;
        let headStart = -1;
        for (let rawIndex = 0; rawIndex < raw.length; rawIndex++) {
            const cells = raw[rawIndex];
            let first = cells[0] ?? '';
            const prefix = LatexReader.RULES.exec(first)?.[0] ?? '';
            first = first.slice(prefix.length);
            // Row markers from longtable apply to the rows before them.
            if (/\\endfirsthead/.test(prefix)) { seenFirstHead = true; rows.forEach(r => { r.header = true; }); headStart = rows.length; }
            if (/\\endhead/.test(prefix)) {
                if (seenFirstHead) rows.splice(headStart);
                else rows.forEach(r => { r.header = true; });
                seenHead = true;
            }
            if (/\\endfoot/.test(prefix) && !/\\endlastfoot/.test(prefix)) { seenFoot = true; }
            if (/\\midrule/.test(prefix) && rows.length && !seenHead && rows.length <= 2 && !rows.some(r => r.header)) rows.forEach(r => { r.header = true; });
            const rowColor = /\\rowcolor\s*(?:\[([^\]]*)\]\s*)?\{([^}]*)\}/.exec(prefix);
            const cellsNow = [first, ...cells.slice(1)];
            // What follows the last `\\` (usually just a closing rule) or a line holding only a rule
            // is not a row; an empty row written with its `&`s is.
            if (cellsNow.length === 1 && !first.trim() && (rawIndex === raw.length - 1 || prefix)) continue;
            // A caption row (longtable) is the table's caption, not a row.
            const cap = /^\s*\\caption\*?\s*(?:\[[^\]]*\]\s*)?\{([\s\S]*)\}\s*$/.exec(first);
            if (cellsNow.length === 1 && cap) {
                caption.push(...this.parseBlocksOf(cap[1], f => { f.paragraphStyle = 'Caption'; }));
                continue;
            }
            rows.push({ cells: cellsNow, rowColor: rowColor ? this.resolveColor(rowColor[2], rowColor[1] ?? null) : undefined });
            void seenFoot;
        }
        if (!rows.length) return null;

        const active = new Map<number, number>();
        const rowNodes: OfficeContentNode[] = [];
        rows.forEach((row, ri) => {
            const cellNodes: OfficeContentNode[] = [];
            let col = 0;
            for (const rawCell of row.cells) {
                checkAbortSignal(this.config.abortSignal);
                let src = rawCell.trim();
                let colSpan = 1, rowSpan = 1;
                let align = aligns[col];
                let bg: string | undefined = row.rowColor;
                const mc = /^\\multicolumn\s*\{\s*(\d+)\s*\}\s*\{/.exec(src);
                if (mc) {
                    colSpan = Math.max(1, Math.min(1000, +mc[1]));
                    const specStart = mc[0].length - 1;
                    const specEnd = this.matchBrace(src, specStart);
                    const cellAlign = this.columnAligns(src.slice(specStart + 1, specEnd))[0];
                    if (cellAlign) align = cellAlign;
                    const contentStart = src.indexOf('{', specEnd + 1);
                    src = contentStart >= 0 ? src.slice(contentStart + 1, this.matchBrace(src, contentStart)) : '';
                }
                src = src.trim();
                const cc = /^\\cellcolor\s*(?:\[([^\]]*)\]\s*)?\{([^}]*)\}/.exec(src);
                if (cc) { bg = this.resolveColor(cc[2], cc[1] ?? null); src = src.slice(cc[0].length).trim(); }
                // (The whitespace after each optional argument is part of it, here and in RULES: two
                // `\s*` with an optional argument between them tried every split of a run of spaces.)
                const mr = /^\\multirow\s*(?:\[[^\]]*\]\s*)?\{\s*(\d+)\s*\}\s*(?:\[[^\]]*\]\s*)?\{[^}]*\}\s*(?:\[[^\]]*\]\s*)?\{/.exec(src);
                if (mr) {
                    rowSpan = Math.max(1, Math.min(rows.length - ri, +mr[1]));
                    const open = mr[0].length - 1;
                    src = src.slice(open + 1, this.matchBrace(src, open));
                }
                const covered = (active.get(col) ?? 0) > 0;
                if (covered && !src.trim()) { col += colSpan; continue; }
                const blocks = this.parseBlocksOf(src);
                const meta: CellMetadata = { row: ri, col };
                if (colSpan > 1) meta.colSpan = colSpan;
                if (rowSpan > 1) { meta.rowSpan = rowSpan; for (let k = 0; k < colSpan; k++) active.set(col + k, rowSpan); }
                if (align === 'center' || align === 'right') meta.align = align;
                if (bg) meta.backgroundColor = bg;
                if (row.header) meta.style = 'header';
                cellNodes.push({ type: 'cell', text: blocks.map(b => b.text ?? '').join('\n'), children: blocks, metadata: meta });
                col += colSpan;
            }
            for (const [c, n] of [...active]) { if (n <= 1) active.delete(c); else active.set(c, n - 1); }
            rowNodes.push({ type: 'row', children: cellNodes });
        });
        return { table: { type: 'table', children: rowNodes }, caption };
    }

    // ── finishing ──

    private finishBlocks(blocks: OfficeContentNode[]): OfficeContentNode[] {
        const clean = (nodes: OfficeContentNode[]): OfficeContentNode[] => {
            const out: OfficeContentNode[] = [];
            for (const n of nodes) {
                delete (n as any).__open;
                if (n.children) n.children = n.type === 'paragraph' || n.type === 'heading' || n.type === 'list' || n.type === 'definitionTerm' || n.type === 'definitionDescription'
                    ? this.mergeRuns(clean(n.children)) : clean(n.children);
                for (const c of n.comments ?? []) delete (c as any).__open;
                if (n.notes) n.notes = clean(n.notes);
                out.push(n);
            }
            return out;
        };
        return clean(blocks);
    }

    /**
     * Merges neighbouring runs with identical formatting and metadata, and lets a whitespace run
     * between two runs sharing a highlight take that highlight (a highlight boxed word by word
     * reads back as one run).
     */
    private mergeRuns(nodes: OfficeContentNode[]): OfficeContentNode[] {
        const key = (n: OfficeContentNode) => JSON.stringify([n.formatting ?? {}, n.metadata ?? null]);
        const keyWithoutHighlight = (n: OfficeContentNode) => { const { backgroundColor: _bg, ...rest } = n.formatting ?? {}; return JSON.stringify([rest, n.metadata ?? null]); };
        for (let i = 1; i + 1 < nodes.length; i++) {
            const [a, b, c] = [nodes[i - 1], nodes[i], nodes[i + 1]];
            if (a.type === 'text' && b.type === 'text' && c.type === 'text' && !(b.text ?? '').trim() && !b.notes?.length && !b.comments?.length
                && a.formatting?.backgroundColor && key(a) === key(c) && keyWithoutHighlight(b) === keyWithoutHighlight(a)) b.formatting = { ...a.formatting };
        }
        const out: OfficeContentNode[] = [];
        for (const n of nodes) {
            const prev = out[out.length - 1];
            if (prev && prev.type === 'text' && n.type === 'text' && !prev.notes?.length && !prev.comments?.length && key(prev) === key(n)
                && !(n.metadata as any)?.citationKey) {
                prev.text = (prev.text ?? '') + (n.text ?? '');
                if (n.notes) prev.notes = n.notes;
                if (n.comments) prev.comments = n.comments;
                continue;
            }
            out.push(n);
        }
        return out;
    }

    /** Fills in the text of `\ref`s now that every label is known. */
    private resolveRefs(): void {
        // A reference repeats its target's name or number: after the first to a label, each is charged
        // to the document's repeated-content budget (see repeatUtils), past which it reads the start of
        // it. `\nameref` to one 100 KB heading 2,000 times made 200 MB of text.
        const referenced = new Set<string>();
        for (const { node, label, kind } of this.refs) {
            const target = this.labelTargets.get(label);
            const number = (target as any)?.__number as string | undefined;
            const title = target?.type === 'heading' ? target.text : undefined;
            let text = kind === 'nameref' ? (title ?? label) : (number ?? title ?? label);
            if (!referenced.has(label)) referenced.add(label);
            else if (takeRepeats(this.config, 1, text.length) === 0) text = repeatPreview(text);
            if (kind === 'eqref') text = `(${text})`;
            node.text = text;
        }
        // A resolved reference changes the text of whatever holds it: derive each enclosing node's text
        // again, with the formula it was built with, so `node.text` (what text output and chunks read)
        // says "Section 1" rather than the label.
        const refNodes = new Set(this.refs.map(r => r.node));
        const refresh = (n: OfficeContentNode): boolean => {
            let dirty = refNodes.has(n);
            for (const c of n.children ?? []) if (refresh(c)) dirty = true;
            for (const c of n.notes ?? []) refresh(c);
            for (const c of n.comments ?? []) refresh(c);
            if (dirty && !refNodes.has(n) && n.text !== undefined) {
                n.text = n.type === 'note' || n.type === 'cell' ? (n.children ?? []).map(b => b.text ?? '').join('\n') : textOf(n.children);
            }
            return dirty;
        };
        if (refNodes.size && this.docFlow) for (const b of this.docFlow.blocks) refresh(b);
        const strip = (nodes: OfficeContentNode[]) => { for (const n of nodes) { delete (n as any).__number; if (n.children) strip(n.children); } };
        if (this.docFlow) strip(this.docFlow.blocks);
    }

    /** Page header/footer content from `fancyhdr`, as auxiliary header/footer nodes. */
    auxiliary(): OfficeAuxiliaryContent | undefined {
        if (this.config.ignoreHeadersAndFooters) return undefined;
        // The fields were collected from the preamble; read them even though the body has ended.
        this.finished = false;
        const build = (fields: Map<string, string>, type: 'header' | 'footer'): OfficeContentNode[] => {
            const paragraphs: OfficeContentNode[] = [];
            for (const pos of ['L', 'C', 'R']) {
                const raw = fields.get(pos);
                if (!raw || !raw.replace(/\\thepage|\\hfill|\s/g, '')) continue;
                const align: TextAlignment = pos === 'C' ? 'center' : pos === 'R' ? 'right' : 'left';
                const blocks = this.parseBlocksOf(raw);
                for (const b of blocks) {
                    if (b.type === 'paragraph') (b.metadata ??= {} as any) && ((b.metadata as any).alignment = align);
                    paragraphs.push(b);
                }
            }
            return paragraphs.length ? [{ type, metadata: { type: 'default' } as any, children: paragraphs }] : [];
        };
        const headers = build(this.headerFields, 'header');
        const footers = build(this.footerFields, 'footer');
        if (!headers.length && !footers.length) return undefined;
        return { ...(headers.length ? { headers } : {}), ...(footers.length ? { footers } : {}) };
    }
}

// ── project loading ─────────────────────────────────────────────────────────────────────────────

/**
 * The legacy encoding a `.tex` file declares, as a WHATWG encoding label: `inputenc`'s option (the
 * last one it knows), or a `% !TEX encoding =` line (TeXShop, TeXstudio). A declaration is ASCII,
 * so it is found in the bytes read as Latin-1, before the file is decoded.
 */
function declaredEncoding(head: string): string | undefined {
    // A commented-out declaration declares nothing.
    const live = head.replace(/(^|[^\\])%.*$/gm, '$1');
    // An option list or argument stops at the next bracket or brace (scanning on to the end from each
    // start took time in the square of an unclosed `\usepackage[` repeated).
    const inputenc = /\\usepackage\s*\[([^\][]*)\]\s*\{inputenc\}/.exec(live) ?? /\\inputencoding\s*\{([^{}]*)\}/.exec(live);
    if (inputenc) {
        const option = inputenc[1].split(',').map(o => o.trim()).reverse().find(o => own(INPUTENC_ENCODINGS, o));
        if (option) return INPUTENC_ENCODINGS[option];
    }
    // Line by line, the name trimmed rather than matched lazily before trailing spaces (which tried every
    // split of a long run of them).
    const magicLine = /^%[ \t]*!TEX[ \t]+encoding[ \t]*=([^\n]*)$/im.exec(head)?.[1].trim();
    const magic = magicLine && /^[A-Za-z0-9_ -]+$/.test(magicLine) ? magicLine.toLowerCase().replace(/[\s_-]+/g, '') : undefined;
    if (!magic) return undefined;
    if (/^utf8/.test(magic)) return 'utf-8';
    if (own(MAGIC_ENCODINGS, magic)) return MAGIC_ENCODINGS[magic];
    try { return new TextDecoder(magic).encoding; } catch { return undefined; }
}

/**
 * The encoding names a `% !TEX encoding` line uses (TeXShop's and TeXstudio's), lowercased without
 * spaces, dashes or underscores, as WHATWG encoding labels. Any other name is tried as a label itself.
 */
const MAGIC_ENCODINGS: Record<string, string> = {
    isolatin: 'windows-1252', isolatin1: 'windows-1252', latin1: 'windows-1252', iso88591: 'windows-1252', windowslatin1: 'windows-1252',
    isolatin2: 'iso-8859-2', latin2: 'iso-8859-2', iso88592: 'iso-8859-2', isolatin5: 'windows-1254', isolatin9: 'iso-8859-15',
    latin9: 'iso-8859-15', iso885915: 'iso-8859-15', isolatingreek: 'iso-8859-7', macosroman: 'macintosh', macroman: 'macintosh',
    windowscentraleuropean: 'windows-1250', windowscyrillic: 'windows-1251',
    maccyrillic: 'x-mac-cyrillic', doscyrillic: 'ibm866', dosrussian: 'ibm866', koi8r: 'koi8-r', gbk: 'gbk', gb2312: 'gbk',
    gb18030: 'gb18030', eucjp: 'euc-jp', sjis: 'shift_jis', sjisx0213: 'shift_jis', dosjapanese: 'shift_jis', macjapanese: 'shift_jis',
    maccyrillicukrainian: 'x-mac-cyrillic', mackorean: 'euc-kr', euckr: 'euc-kr', big5: 'big5', macchinesetraditional: 'big5',
    macchinesesimplified: 'gbk', doschinesetraditional: 'big5', doschinesesimplified: 'gbk',
};

/**
 * Decodes a `.tex` file. UTF-8 (LaTeX's default input encoding since 2018) when it is valid UTF-8;
 * otherwise the legacy encoding the file declares (or `inherited`, the main file's, for an included
 * file that declares none), or else Windows-1252, the usual 8-bit one. A file that is mostly UTF-8
 * with a few damaged bytes, and declares nothing else, stays UTF-8.
 */
function decodeTex(buf: Buffer, inherited?: string): { text: string; declared?: string } {
    const declared = declaredEncoding(buf.subarray(0, 1 << 20).toString('latin1'));
    try {
        return { text: new TextDecoder('utf-8', { fatal: true }).decode(buf), declared };
    } catch { /* not valid UTF-8 */ }
    // A declared UTF-8 the bytes contradict is judged by the bytes, like an undeclared file.
    let label = declared === 'utf-8' ? undefined : declared ?? (inherited === 'utf-8' ? undefined : inherited);
    if (!label) {
        // Valid multi-byte sequences outnumbering the damaged ones mean UTF-8 text with some damage.
        const lenient = buf.toString('utf8');
        let replaced = 0, multibyte = 0;
        for (const ch of lenient) { if (ch === '\uFFFD') replaced++; else if (ch.charCodeAt(0) > 0x7F) multibyte++; }
        label = multibyte > replaced ? 'utf-8' : 'windows-1252';
    }
    if (label === 'utf-8') return { text: buf.toString('utf8'), declared };
    try { return { text: new TextDecoder(label).decode(buf), declared }; }
    catch { return { text: new TextDecoder('windows-1252').decode(buf), declared }; }
}

const ZIP_MAGIC = [0x50, 0x4b];

/** Picks the main file of a project: `main.tex`, else a top-level `.tex` with `\documentclass`, else the shallowest one. */
function findMainFile(files: Map<string, Buffer>): string | null {
    if (files.has('main.tex')) return 'main.tex';
    const tex = [...files.keys()].filter(p => p.toLowerCase().endsWith('.tex')).sort((a, b) => a.split('/').length - b.split('/').length || a.localeCompare(b));
    return tex.find(p => /\\documentclass/.test(files.get(p)!.toString('utf8'))) ?? tex[0] ?? null;
}

/**
 * Parses LaTeX: a `.tex` document, or a project zip (such as an Overleaf download or the LaTeX
 * generator's bundle) whose `\input`/`\include` files and images are then read from the archive.
 */
export const parseLatex = async (buffer: Buffer, config: FullOfficeParserConfig): Promise<OfficeParserAST> => {
    checkAbortSignal(config.abortSignal);
    let project: Project | null = null;
    let decoded: { text: string; declared?: string };
    let mainPath: string | undefined;
    if (buffer.length >= 2 && ZIP_MAGIC.every((b, i) => buffer[i] === b)) {
        const entries = await extractFiles(buffer, name => !name.endsWith('/'), config.decompressionLimits, config);
        const files = new Map<string, Buffer>();
        for (const e of entries) {
            const p = normalizeProjectPath('', e.path);
            if (p !== null) files.set(p, e.content);
        }
        const main = findMainFile(files);
        if (!main) {
            logWarning(OfficeWarningType.LATEX_FILE_NOT_FOUND, config, { files: ['main.tex'] });
            return createAST('tex', {}, [], [], config, undefined);
        }
        project = { files, root: '' };
        mainPath = main;
        decoded = decodeTex(files.get(main)!);
    } else {
        decoded = decodeTex(buffer);
    }
    const src = decoded.text.replace(/^﻿/, '').replace(/\r\n?/g, '\n');

    const mainDir = mainPath && mainPath.includes('/') ? mainPath.slice(0, mainPath.lastIndexOf('/')) : '';
    const reader = new LatexReader(config, project, mainDir);
    reader.inputEncoding = decoded.declared;
    const content = reader.parse(src, mainPath);
    const attachments = reader.attachments;

    if (config.ocr && attachments.length) {
        for (const att of attachments) {
            if (!att.mimeType.startsWith('image/')) continue;
            const ocrText = await ocrDuringParse(Buffer.from(att.data, 'base64'), config, att.name);
            if (ocrText !== undefined) att.ocrText = ocrText;
        }
        const attachmentsByName = attachmentLookup(attachments);
        const assign = (nodes: OfficeContentNode[]) => {
            for (const n of nodes) {
                const name = (n.metadata as ImageMetadata | undefined)?.attachmentName;
                if (n.type === 'image' && name) {
                    const ocr = attachmentsByName.get(name)?.ocrText;
                    if (ocr) n.text = ocr;
                }
                if (n.children) assign(n.children);
            }
        };
        assign(content);
    }

    return createAST('tex', reader.metadata, content, attachments, config, reader.auxiliary());
};

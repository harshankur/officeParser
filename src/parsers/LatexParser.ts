import { AdmonitionMetadata, CellMetadata, CodeMetadata, CommentMetadata, FullOfficeParserConfig, HeadingMetadata, ImageMetadata, ListMetadata, NoteMetadata, OfficeAttachment, OfficeAuxiliaryContent, OfficeContentNode, OfficeMetadata, OfficeParserAST, OfficeWarningType, ParagraphMetadata, TextAlignment, TextFormatting, TextMetadata } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { checkAbortSignal, logWarning } from '../utils/errorUtils.js';
import { createAttachment } from '../utils/imageUtils.js';
import { LATEX_SYMBOL_CHARACTERS, LISTINGS_LANGUAGE_NAMES } from '../utils/latexUtils.js';
import { ADMONITION_COLOR } from '../utils/officeGenUtils.js';
import { performOcr } from '../utils/ocrUtils.js';
import { extractFiles } from '../utils/zipUtils.js';

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
const MAX_NESTING_DEPTH = 256;
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
const DRAWING_ENVS = new Set(['tikzpicture', 'pgfpicture', 'picture', 'forest', 'axis', 'circuitikz', 'pspicture', 'tikzcd', 'filecontents', 'filecontents*']);

/**
 * Environments with no meaning of their own for the AST (layout wrappers): their content is read
 * in place. The value is the argument signature consumed after `\begin{name}` (see {@link readArgs}).
 */
const TRANSPARENT_ENVS: Record<string, string> = {
    minipage: 'oomm', columns: 'o', column: 'om', multicols: 'mo', 'multicols*': 'mo', adjustbox: 'm', landscape: '',
    spacing: 'm', singlespace: '', onehalfspace: '', doublespace: '', small: '', footnotesize: '', scriptsize: '',
    large: '', Large: '', sloppypar: '', frontmatter: '', mainmatter: '', appendices: '', subequations: '',
    samepage: '', otherlanguage: 'm', 'otherlanguage*': 'm', english: '', onlyenv: 'O', overprint: 'o', uncoverenv: 'O',
    visibleenv: 'O', actionenv: 'O', titlepage: '', subfigure: 'om', subtable: 'om', framed: '', mdframed: 'o', tcolorbox: 'o',
    theorem: 'o', lemma: 'o', proof: 'o', definition: 'o', corollary: 'o', proposition: 'o', example: 'o', remark: 'o',
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
    newtheorem: 'moo', theoremstyle: 'm', AtBeginDocument: 'm', AtEndDocument: 'm', listfiles: '', hyphenation: 'm',
    enlargethispage: 'sm', newpage: '', color: '', label: '', today: '', column: 'om', textwidth: '', linewidth: '',
    columnwidth: '', paperwidth: '', textheight: '', paperheight: '', baselineskip: '', parindent_: '', tabcolsep: '',
    arraybackslash: '', arrayrulewidth: '', dimexpr: '', fboxsep: '', setbeamersize: 'm', logo: 'm', institute: 'om',
};

/** Commands whose last mandatory argument is ordinary content, with the signature before it. */
const WRAPPER_COMMANDS: Record<string, string> = {
    mbox: '', makebox: 'oo', fbox: '', framebox: 'oo', parbox: 'oom', raisebox: 'moo', scalebox: 'mo', resizebox: 'sm',
    rotatebox: 'om', centerline: '', textnormal: '', textup: '', textmd: '', textsc: '', textulc: '', textbf_: '',
    boxed: '', text: '', only: 'O', uncover: 'O', visible: 'O', invisible: 'O', alt: 'Om', temporal: 'Omm', onslide: 'O',
    structure: 'O', alert: 'O', emph_: '', newblock: '', tcbox: 'o', adjustbox: 'm', shortstack: 'o', underbrace: '',
};

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

    next(atLetter = false): Tok {
        for (;;) {
            const c = this.nextCh();
            if (c === undefined) return { t: 'eof' };
            if (c === '\\') {
                const n = this.peekCh();
                if (n === undefined) { this.afterCs = false; return { t: 'text', v: '\\' }; }
                if (this.isLetter(n, atLetter)) {
                    let name = '';
                    while (this.isLetter(this.peekCh(), atLetter)) name += this.nextCh();
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
                let name = '';
                if (/[A-Za-z]/.test(this.peekCh() ?? '')) { while (/[A-Za-z@]/.test(this.peekCh() ?? '')) name += this.nextCh(); }
                else name = this.nextCh() ?? '';
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
        const snap = this.save();
        this.skipBlanks();
        if (this.peekCh() === '*') { this.nextCh(); return true; }
        this.restore(snap);
        return false;
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
        let out = '';
        for (;;) {
            const c = this.nextCh();
            if (c === undefined) return { text: out, found: false };
            out += c;
            if (out.endsWith(marker)) return this.endRaw({ text: out.slice(0, -marker.length), found: true });
        }
    }

    /** Raw source up to the `\end{name}` matching an already-consumed `\begin{name}`. */
    readRawEnvBody(name: string): string {
        const begin = `\\begin{${name}}`;
        const end = `\\end{${name}}`;
        let depth = 1;
        let out = '';
        for (;;) {
            const c = this.nextCh();
            if (c === undefined) return out;
            out += c;
            if (c === '}' ) {
                if (out.endsWith(end) && !out.endsWith('\\' + end)) {
                    if (--depth === 0) return this.endRaw(out.slice(0, -end.length));
                } else if (out.endsWith(begin)) depth++;
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

interface UserMacro { nargs: number; optDefault: string | null; body: string; }
interface UserEnv { nargs: number; optDefault: string | null; begin: string; end: string; }

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
    return (nodes || []).map(n => (n.type === 'break' ? '\n' : n.text ?? textOf(n.children))).join('');
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
    /** Metadata fields `\hypersetup` set explicitly (`pdftitle`, `pdfauthor`), which `\title`/`\author` do not override. */
    private readonly pdfMetadata = new Set<'title' | 'author'>();
    attachments: OfficeAttachment[] = [];
    private attachmentByPath = new Map<string, string>();
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
        const content = this.finishBlocks(flow.blocks);
        this.resolveRefs();
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
        const cls = /\\documentclass\s*(?:\[([^\]]*)\])?\s*\{([^}]*)\}/.exec(src);
        if (cls) {
            this.native.documentClass = cls[2].trim();
            this.native.classOptions = (cls[1] ?? '').split(',').map(s => s.trim()).filter(Boolean);
            this.beamer = this.native.documentClass === 'beamer';
            const size = this.native.classOptions.map((o: string) => /^(10|11|12)pt$/.exec(o)).find(Boolean);
            if (size) this.classSize = +size[1];
        }
        SECTIONING.forEach((name, rank) => { if (new RegExp(`\\\\${name}\\*?\\s*[[{]`).test(src)) this.sectionRanks.add(rank); });
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
        flow.labelTarget = null;
    }

    private addSpace(flow: Flow): void {
        if (flow.inline.length === 0 && !flow.pending) return;
        if (flow.pending.endsWith(' ')) return;
        this.addText(flow, ' ');
    }

    private flushText(flow: Flow): void {
        if (!flow.pending || !flow.pendingState) { flow.pending = ''; flow.pendingState = null; return; }
        const st = flow.pendingState;
        const text = flow.pendingMono ? flow.pending : ligatures(flow.pending);
        const node: OfficeContentNode = { type: 'text', text };
        const fmt = Object.fromEntries(Object.entries(st.fmt).filter(([, v]) => v !== undefined && v !== false));
        if (Object.keys(fmt).length) node.formatting = fmt as TextFormatting;
        if (st.link && !this.config.ignoreInternalLinks || (st.link && !st.link.internal)) {
            node.metadata = { link: st.link!.url, linkType: st.link!.internal ? 'internal' : 'external' } as TextMetadata;
        }
        flow.inline.push(node);
        flow.pending = '';
        flow.pendingState = null;
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
        if (last && out.indexOf(last) === out.length - 1) last.text = (last.text ?? '').replace(/[ \t]+$/, '');
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
        if (!label || this.config.ignoreInternalLinks) return;
        const pendingText = flow.pending.trim() || this.hasContent(flow.inline);
        if (flow.labelTarget && !pendingText) {
            const meta = (flow.labelTarget.metadata ??= {} as any) as any;
            meta.anchorIds = [...(meta.anchorIds ?? []), label];
            this.labelTargets.set(label, flow.labelTarget);
        } else if (pendingText) {
            flow.anchors.push(label);
            this.labelTargets.set(label, { type: 'paragraph', text: '' } as OfficeContentNode);
        } else {
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
                case 'comment': this.handleComment(flow, tok.v); break;
                case 'cs': {
                    if (stop.items && tok.name === stop.items) { sc.restore(snap); unwind(); return 'item'; }
                    if (tok.name === 'end') {
                        const env = (sc.readRawGroup() ?? '').trim();
                        if (stop.env && env === stop.env) { unwind(); return 'end'; }
                        if (env === 'document') { this.endParagraph(flow); this.finished = true; unwind(); return 'enddoc'; }
                        const user = this.envs.get(env);
                        if (user) this.expand(sc, user.end);
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
        return this.parseFlow(new Scanner(raw), flow, {});
    }

    /** Parses raw source as the body of a separate container (a note, a cell) and returns its blocks. */
    private parseBlocksOf(raw: string, patch?: (f: Flow) => void, freshState = true): OfficeContentNode[] {
        const flow = new Flow();
        patch?.(flow);
        const run = () => { this.parseFlow(new Scanner(raw), flow, {}); this.endParagraph(flow); };
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
        const blocks = this.parseBlocksOf(raw);
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
    private substitute(sc: Scanner, def: { nargs: number; optDefault: string | null; body: string }): string {
        const args: string[] = [];
        for (let k = 0; k < def.nargs; k++) {
            if (k === 0 && def.optDefault !== null) args.push(sc.readRawOptional() ?? def.optDefault);
            else args.push(sc.readRawGroup() ?? '');
        }
        return def.body.replace(/##|#([1-9])/g, (m, n) => (m === '##' ? '#' : (args[+n - 1] ?? '')));
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

    // ── comments ──

    private handleComment(flow: Flow, text: string): void {
        const t = text.replace(/^ /, '');
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
        if (flow.pending.trim() || this.hasContent(flow.inline)) {
            const run = this.anchorRun(flow);
            (run.comments ??= []).push(node);
        } else {
            flow.nextComments.push(node);
        }
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

        if (name in SYMBOLS) {
            if (name === '-' || name === ',') this.addText(flow, SYMBOLS[name]);
            else if (SYMBOLS[name] === '' ) { /* spacing no-op */ }
            else this.addText(flow, SYMBOLS[name]);
            return;
        }
        if (name in ACCENTS) {
            let arg = sc.readRawGroup() ?? '';
            if (arg === '\\i') arg = 'i';
            if (arg === '\\j') arg = 'j';
            const base = arg.startsWith('\\') ? (SYMBOLS[arg.slice(1)] ?? '') : arg;
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
                    this.addInline(flow, { type: 'text', text: key, metadata: { citationKey: key } as TextMetadata });
                });
                return;
            }
            case 'ref': case 'autoref': case 'cref': case 'Cref': case 'eqref': case 'pageref': case 'nameref': case 'vref': case 'Autoref': {
                sc.readStar();
                const label = (sc.readRawGroup() ?? '').trim();
                const node: OfficeContentNode = { type: 'text', text: label };
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
                const pt = /^\s*([\d.]+)\s*(pt)?\s*$/.exec(size ?? '');
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
            case 'centering': this.state.align = 'center'; return;
            case 'raggedleft': this.state.align = 'right'; return;
            case 'raggedright': this.state.align = 'left'; return;
            case 'leftskip': case 'rightskip': case 'hangindent': case 'parindent': case 'hangafter': case 'parskip': {
                const v = this.readDimenAssignment(sc);
                if (name === 'leftskip') this.state.left = v;
                else if (name === 'rightskip') this.state.right = v;
                else if (name === 'hangindent') this.state.hang = v;
                return;
            }
            case 'hspace': {
                const star = sc.readStar();
                const len = this.dimenPt(sc.readRawGroup() ?? '');
                if (star && !flow.pending && !this.hasContent(flow.inline) && len) {
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
            case 'newcommand': case 'renewcommand': case 'providecommand': case 'DeclareRobustCommand': case 'NewDocumentCommand_':
                this.defineMacro(sc, name !== 'providecommand' && name !== 'newcommand'); return;
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
                    else if (!(b.name in SYMBOLS)) this.macros.set(a.name, { nargs: 0, optDefault: null, body: `\\${b.name}` });
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
            case 'author': { sc.readRawOptional(); const raw = sc.readRawGroup(); this.titleParts.author = raw ?? undefined; const a = this.authorText(raw); if (a && !this.pdfMetadata.has('author')) this.metadata.author = a; return; }
            case 'date': { const raw = sc.readRawGroup(); this.titleParts.date = raw ?? undefined; const d = this.plainText(raw); if (d) this.native.date = d; return; }
            case 'subtitle': { sc.readRawOptional(); const raw = sc.readRawGroup(); this.titleParts.subtitle = raw ?? undefined; const t = this.plainText(raw); if (t) this.native.subtitle = t; return; }
            case 'maketitle': this.typesetTitle(flow); return;
            case 'hypersetup': this.hypersetup(sc.readRawGroup() ?? ''); return;
            case 'usepackage': {
                const opts = sc.readRawOptional();
                const pkgs = (sc.readRawGroup() ?? '').split(',').map(p => p.trim()).filter(Boolean);
                (this.native.packages as string[]).push(...pkgs);
                void opts;
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
            case 'iffalse': sc.readRawUntil('\\fi'); return;
            case 'ifPDFTeX': case 'ifXeTeX': case 'ifLuaTeX': case 'iftutex': case 'ifpdf': case 'else': case 'fi': case 'ifdefined':
                return;
        }

        if (name in WRAPPER_COMMANDS) {
            this.readArgs(sc, WRAPPER_COMMANDS[name]);
            this.parseRawInto(flow, sc.readRawGroup());
            return;
        }
        if (name in IGNORED_COMMANDS) {
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
        const m = /^\s*([-+]?[\d.]+)\s*(pt|bp|in|cm|mm|pc|em|ex|px|sp)?\s*$/.exec(raw);
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
        const named = NAMED_COLORS[n];
        return named ? `#${named}` : undefined;
    }

    // ── metadata helpers ──

    /**
     * The block `\maketitle` (or beamer's `\titlepage`) typesets, as content at that spot: a heading
     * styled `Title`, then `Subtitle`, `Author` and `Date` lines, the way a word processor's title
     * page reads. The values also stay in `ast.metadata`. `\thanks` become footnotes on the line that
     * carries them. A date left to `\today` is the day the document is compiled, so it is omitted.
     */
    private typesetTitle(flow: Flow): void {
        const { title, subtitle, author, date } = this.titleParts;
        // LaTeX refuses \maketitle without a \title, and empties the title after typesetting it once.
        if (!title || !this.plainText(this.stripThanks(title))) return;
        if (this.titleTypeset && !this.beamer) return;
        this.titleTypeset = true;
        this.endParagraph(flow);
        const thanks = (raw: string) => raw.replace(/\\thanks\b/g, '\\footnote');
        const heading = this.headingNode(thanks(title), 1);
        Object.assign(heading.metadata as HeadingMetadata, { style: 'Title', alignment: 'center' });
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

    private stripThanks(raw: string | null): string | null {
        return raw === null ? null : raw.replace(/\\thanks\s*\{(?:[^{}]|\{[^{}]*\})*\}/g, '');
    }

    private authorText(raw: string | null): string {
        if (raw === null) return '';
        return this.stripThanks(raw)!.split(/\\and\b|\\AND\b/).map(a => this.plainText(a.replace(/\\\\/g, ' '))).filter(Boolean).join(', ');
    }

    private hypersetup(raw: string): void {
        for (const [key, value] of this.keyValues(raw)) {
            const v = this.plainText(value);
            switch (key) {
                // The PDF metadata a document states outright is its metadata; \title and \author are
                // what the title block prints, and fill in only when it states none.
                case 'pdftitle': if (v) { this.metadata.title = v; this.pdfMetadata.add('title'); } break;
                case 'pdfauthor': if (v) { this.metadata.author = v; this.pdfMetadata.add('author'); } break;
                case 'pdfsubject': this.metadata.subject = v; break;
                case 'pdfkeywords': this.metadata.keywords = v; break;
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
                        else if (k) (this.metadata.customProperties ??= {})[k] = text;
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
        this.withState(s => { s.inline = true; s.align = undefined; }, () => { this.parseFlow(new Scanner(raw), flow, {}); this.flushText(flow); });
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
        let text = this.project!.files.get(resolved)!.toString('utf8').replace(/\r\n?/g, '\n');
        if (subfile) {
            const body = /\\begin\{document\}([\s\S]*?)\\end\{document\}/.exec(text);
            if (body) text = body[1];
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

    private image(sc: Scanner, flow: Flow): void {
        sc.readStar();
        const opts = sc.readRawOptional() ?? '';
        const path = (sc.readRawGroup() ?? '').trim().replace(/^"|"$/g, '');
        const meta: ImageMetadata = { attachmentName: '' };
        const kv = new Map(this.keyValues(opts));
        const width = kv.get('width');
        if (width) {
            const frac = /^\s*([\d.]*)\s*\\(linewidth|textwidth|columnwidth|hsize)\s*$/.exec(width);
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
                name = resolved;
                this.attachments.push(createAttachment(name, this.project!.files.get(resolved)!));
                this.attachmentByPath.set(resolved, name);
            }
            meta.attachmentName = name;
        } else {
            if (!resolved) this.missingFiles.add(path);
            meta.url = path;
            delete (meta as any).attachmentName;
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
            flow.pending = '';
            flow.pendingState = null;
            flow.nextAnchors = [];
            flow.nextComments = [];
            this.bodyStarted = true;
            return;
        }
        if (this.nesting >= MAX_NESTING_DEPTH) { this.tooDeep(flow, sc.readRawEnvBody(env)); return; }

        const user = this.envs.get(env);
        if (user) {
            const code = this.substitute(sc, { nargs: user.nargs, optDefault: user.optDefault, body: user.begin });
            if (!this.expand(sc, code)) sc.readRawEnvBody(env);
            return;
        }
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
        if (env in TABLE_ENVS) { this.table(sc, flow, env); return; }
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
            const text = inner.map(b => b.text ?? '').join('\n').trim();
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
        if (env in TRANSPARENT_ENVS) {
            this.readArgs(sc, TRANSPARENT_ENVS[env]);
            return this.envInto(sc, flow, env);
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
        this.addText(flow, raw.replace(/\\(begin|end)\s*\{[^}]*\}/g, ' ').replace(/\\[A-Za-z@]+\*?|\\.|[{}]/g, ' ').replace(/\s+/g, ' '));
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
            const type = (byColor ?? (byName in ADMONITION_COLOR ? byName : undefined)) as AdmonitionMetadata['admonitionType'] | undefined;
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
        const inner = this.subFlow(sc, env, flow);
        // Labels in a float name its figure or table, wherever in the float they were written.
        const main = inner.find(b => b.type === 'image' || b.type === 'table');
        const kind = env.startsWith('table') || env === 'wraptable' ? 'table' : 'figure';
        if (inner.some(b => (b.metadata as any)?.style === 'Caption')) {
            const n = (this.floatNumbers[kind] = (this.floatNumbers[kind] ?? 0) + 1);
            if (main) (main as any).__number = String(n);
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
            this.nesting++;
            try { r = this.withState(() => { }, () => this.parseFlow(sc, itemFlow, { env, items: 'item' })); } finally { this.nesting--; }
            this.endParagraph(itemFlow);
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
        const aligns = this.columnAligns(spec);
        this.endParagraph(flow);
        const node = this.buildTable(body, aligns, env);
        if (node) this.pushBlock(flow, node.table);
        if (node?.caption.length) for (const c of node.caption) this.pushBlock(flow, c);
    }

    /** Column alignments from a column specification (`l`, `c`, `r`, `p{}` with `>{\centering}`, `*{n}{...}`). */
    private columnAligns(spec: string): (('left' | 'center' | 'right') | undefined)[] {
        let s = spec;
        for (let guard = 0; guard < 8 && /\*\{\s*(\d+)\s*\}\{/.test(s); guard++) {
            s = s.replace(/\*\{\s*(\d+)\s*\}\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}/, (_m, n, inner) => inner.repeat(Math.min(1000, +n)));
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

    private static RULES = /^\s*(?:\\(?:hline|toprule|midrule|bottomrule|endhead|endfirsthead|endfoot|endlastfoot|hhline\s*\{[^}]*\}|cline\s*\{[^}]*\}|cmidrule(?:\s*\([^)]*\))?\s*(?:\[[^\]]*\])?\s*\{[^}]*\}|addlinespace(?:\s*\[[^\]]*\])?|noalign\s*\{[^}]*\}|rowcolor\s*(?:\[[^\]]*\])?\s*\{[^}]*\}|specialrule\s*\{[^}]*\}\{[^}]*\}\{[^}]*\})(?:\[[^\]]*\])?\s*)+/;

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
            const rowColor = /\\rowcolor\s*(?:\[([^\]]*)\])?\s*\{([^}]*)\}/.exec(prefix);
            const cellsNow = [first, ...cells.slice(1)];
            // What follows the last `\\` (usually just a closing rule) or a line holding only a rule
            // is not a row; an empty row written with its `&`s is.
            if (cellsNow.length === 1 && !first.trim() && (rawIndex === raw.length - 1 || prefix)) continue;
            // A caption row (longtable) is the table's caption, not a row.
            const cap = /^\s*\\caption\*?\s*(?:\[[^\]]*\])?\s*\{([\s\S]*)\}\s*$/.exec(first);
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
                const cc = /^\\cellcolor\s*(?:\[([^\]]*)\])?\s*\{([^}]*)\}/.exec(src);
                if (cc) { bg = this.resolveColor(cc[2], cc[1] ?? null); src = src.slice(cc[0].length).trim(); }
                const mr = /^\\multirow\s*(?:\[[^\]]*\])?\s*\{\s*(\d+)\s*\}\s*(?:\[[^\]]*\])?\s*\{[^}]*\}\s*(?:\[[^\]]*\])?\s*\{/.exec(src);
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
        for (const { node, label, kind } of this.refs) {
            const target = this.labelTargets.get(label);
            const number = (target as any)?.__number as string | undefined;
            const title = target?.type === 'heading' ? target.text : undefined;
            let text = kind === 'nameref' ? (title ?? label) : (number ?? title ?? label);
            if (kind === 'eqref') text = `(${text})`;
            node.text = text;
        }
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
    let src: string;
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
        src = files.get(main)!.toString('utf8');
    } else {
        src = buffer.toString('utf8');
    }
    src = src.replace(/^﻿/, '').replace(/\r\n?/g, '\n');

    const mainDir = mainPath && mainPath.includes('/') ? mainPath.slice(0, mainPath.lastIndexOf('/')) : '';
    const reader = new LatexReader(config, project, mainDir);
    const content = reader.parse(src, mainPath);
    const attachments = reader.attachments;

    if (config.ocr && attachments.length) {
        for (const att of attachments) {
            checkAbortSignal(config.abortSignal);
            if (!att.mimeType.startsWith('image/')) continue;
            try { att.ocrText = (await performOcr(Buffer.from(att.data, 'base64'), { ...config.ocrConfig })).trim(); }
            catch (e) { logWarning(OfficeWarningType.OCR_FAILED, config, att.name, e); }
        }
        const assign = (nodes: OfficeContentNode[]) => {
            for (const n of nodes) {
                const name = (n.metadata as ImageMetadata | undefined)?.attachmentName;
                if (n.type === 'image' && name) {
                    const ocr = attachments.find(a => a.name === name)?.ocrText;
                    if (ocr) n.text = ocr;
                }
                if (n.children) assign(n.children);
            }
        };
        assign(content);
    }

    return createAST('tex', reader.metadata, content, attachments, config, reader.auxiliary());
};

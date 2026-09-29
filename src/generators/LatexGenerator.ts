import { zipSync, Zippable } from 'fflate';
import { UniqueNames } from '../utils/uniqueNames.js';
import { trimEndChars } from '../utils/textUtils.js';
import { AdmonitionMetadata, CodeMetadata, CommentMetadata, ConversionResult, GeneratorConfig, HeadingMetadata, ImageMetadata, ListMetadata, NoteMetadata, OfficeContentNode, OfficeParserAST, OfficeWarningType, ParagraphMetadata, TexDocumentClass, TextFormatting, TextMetadata } from '../types.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { ADMONITION_COLOR, decodeBase64, embedUrl, fillSheetRowGaps, hexColor, isHeaderRow, lengthToPt, marginPt, MIME_EXT, paperSizePt, resolveZipInstant, sniffImageSize, toW3CDTF } from '../utils/officeGenUtils.js';
import { LatexMathPlan, LatexScripts, LatexUnicodePlan, LISTINGS_LANGUAGES, mathCommandsOf, planLatexMath, planLatexUnicode, withLatexMathMacros } from '../utils/latexUtils.js';
import { escapeLatex, latexComment, latexSourceComment, sanitizeLatexImagePath, sanitizeLatexMath, sanitizeLatexUrl } from '../utils/sanitize.js';
import { isSourceComment } from '../utils/commentUtils.js';
import { contentHash, imageToTextPdf, newDecodeBudget } from '../utils/textPdf.js';
import { BaseGenerator } from './BaseGenerator.js';
import { lookupTable } from '../utils/lookupUtils.js';
import { appendAll } from '../utils/nodeListUtils.js';

/**
 * A line break inside a paragraph. `\newline` (and `\\`) raise "There's no line here to end" when
 * they come first in a paragraph, which a text run starting with a newline would trigger; `\hfil`
 * starts the line first, so this form is safe anywhere. The trailing `{}` ends the control word.
 */
const LINE_BREAK = '\\hfil\\break{}';

/** Blank line: the paragraph/block separator. */
const BLOCK_SEPARATOR = '\n\n';

/** Name of the `.tex` file inside a bundle (the Overleaf convention). */
const BUNDLE_MAIN_FILE = 'main.tex';

/** Folder images are written to and referenced from, relative to the `.tex`. */
const IMAGE_DIR = 'images';
/** Extensions `\includegraphics` adds to a path given without one (pdfTeX's list, which XeTeX and LuaTeX share). */
const IMAGE_PROBE_EXTENSIONS = ['.pdf', '.png', '.jpg', '.jpeg', '.eps', '.PDF', '.PNG', '.JPG', '.JPEG'];

/**
 * Image types every LaTeX engine can `\includegraphics` directly (pdfTeX, XeTeX, LuaTeX). Anything
 * else (GIF, BMP, TIFF, WebP, SVG, EMF) has to be converted first, so it is bundled but drawn as a
 * framed placeholder.
 */
const INCLUDABLE_IMAGE_EXT: Record<string, string> = lookupTable({ 'image/png': 'png', 'image/jpeg': 'jpg', 'image/jpg': 'jpg', 'application/pdf': 'pdf' });

/** `index` written in letters (0 is `a`, 26 is `ba`), for a macro name, which cannot hold digits. */
function letterIndex(index: number): string {
    let out = '';
    do { out = String.fromCharCode(97 + (index % 26)) + out; index = Math.floor(index / 26); } while (index > 0);
    return out;
}

/** LaTeX's own nesting limit for `itemize`/`enumerate` ("Too deeply nested" past it); beamer allows one level fewer. */
const MAX_LIST_DEPTH = 4;
const MAX_LIST_DEPTH_BEAMER = 3;

/** `enumerate` counter per nesting level of enumerate environments. */
const ENUM_COUNTERS = ['enumi', 'enumii', 'enumiii', 'enumiv'];

/** Guards against a hostile span/column count turning into an enormous grid. */
const MAX_TABLE_COLUMNS = 1000;

/**
 * Widest table written as one `longtable`/`tabular`; a wider one continues below itself in bands
 * of this many columns. Past this the columns are too narrow to read on a page, and a table with
 * hundreds of columns exhausts pdfTeX's and XeTeX's memory building the column templates.
 */
const MAX_TABLE_BAND_COLUMNS = 16;
const MAX_ROW_SPAN = 1000;

/**
 * Wide tables get a smaller font and tighter padding so their columns can hold a word or a number;
 * checked in order, the first step whose column count is reached applies.
 */
const TABLE_COMPACT_STEPS = [
    { minCols: 13, size: '\\scriptsize', tabcolsep: '2pt' },
    { minCols: 9, size: '\\footnotesize', tabcolsep: '3pt' },
    { minCols: 6, size: '\\small', tabcolsep: '4pt' },
];

/** Columns a tab advances to in code blocks (verbatim prints a tab as one space). */
const CODE_TAB_WIDTH = 4;

/** Font size classes accept (`\documentclass[11pt]`). */
const CLASS_FONT_SIZES = [10, 11, 12];
const DEFAULT_CLASS_FONT_SIZE = 11;

/** Leading used with `\fontsize`, as a multiple of the size (LaTeX's own ratio). */
const LEADING_FACTOR = 1.2;

/** Bounds for an explicit run font size, so a hostile `size` cannot request an absurd font. */
const MIN_FONT_SIZE_PT = 4;
const MAX_FONT_SIZE_PT = 100;

/** Pixel density assumed for an image's intrinsic size, matching the HTML/DOCX/ODT generators. */
const IMAGE_DPI = 96;
const PT_PER_INCH = 72;

/** Tallest an image may be drawn, per context, so a tall image cannot run off the page. */
const IMAGE_MAX_HEIGHT_PAGE = '0.9\\textheight';
const IMAGE_MAX_HEIGHT_FRAME = '0.75\\textheight';
const IMAGE_MAX_HEIGHT_HEADER = '1cm';

/** Height of one line of a page header/footer, for `\headheight`. */
const HEADER_LINE_HEIGHT_PT = 14;

/** Extra `\headheight` for a header holding an image ({@link IMAGE_MAX_HEIGHT_HEADER}, 1cm, in points). */
const HEADER_IMAGE_HEIGHT_PT = 28.5;

/** Longest label name written; longer anchor ids are truncated (deterministically). */
const MAX_LABEL_LENGTH = 200;

/** Heading level to sectioning command, per class family. Levels past the list use the last entry. */
const ARTICLE_SECTIONS = ['section', 'subsection', 'subsubsection', 'paragraph', 'subparagraph'];
const CHAPTER_SECTIONS = ['chapter', 'section', 'subsection', 'subsubsection', 'paragraph', 'subparagraph'];

/** Document classes {@link TexGeneratorConfig.documentClass} accepts. */
const DOCUMENT_CLASSES: ReadonlySet<TexDocumentClass> = new Set(['auto', 'article', 'report', 'book', 'beamer']);

/**
 * Font names drawn in the typewriter face. A run's font family is otherwise not carried over:
 * a named font may not be installed where the document is compiled (and cannot be loaded at all
 * under pdfTeX), so only the monospace/typewriter distinction, which is meaningful, is kept.
 */
const MONOSPACE_FONTS = new Set([
    'monospace', 'courier', 'courier new', 'consolas', 'menlo', 'monaco', 'lucida console', 'lucida sans typewriter',
    'source code pro', 'fira code', 'fira mono', 'dejavu sans mono', 'liberation mono', 'ubuntu mono', 'roboto mono',
    'jetbrains mono', 'sf mono', 'cascadia code', 'cascadia mono', 'inconsolata', 'andale mono', 'noto sans mono',
]);


/**
 * Where content is being written, which decides what LaTeX constructs are legal there. The same
 * AST node can land at the top of the page, in a table cell, in a footnote or in a heading, and a
 * construct that is fine in one (a `longtable`, `verbatim`, `\section`) is a fatal error in another.
 */
interface RenderContext {
    /** Sectioning commands are allowed (not in cells, lists, notes, quotes or beamer frames). */
    sections: boolean;
    /** `verbatim`/`lstlisting` are allowed (not inside a macro argument or table cell). */
    verbatim: boolean;
    /** `longtable` is allowed (only at the top level of an article-family body). */
    longtable: boolean;
    /** Display math is allowed (not in a moving argument). */
    display: boolean;
    /** Inside a moving argument (heading, frame title): fragile commands take `\protect`, line breaks become spaces. */
    moving: boolean;
    /**
     * Footnotes: written in place, as marks with deferred text (inside `tabular`), as marks whose text
     * follows the footnote they stand in (in a footnote's own text), in place in a `longtable` cell
     * (which holds a footnote's text back), in parentheses where no mark can be numbered (in an
     * endnote's, a deferred footnote's or a cell footnote's text), or dropped (page headers).
     */
    notes: 'direct' | 'deferred' | 'cell' | 'nested' | 'parenthetical' | 'omit';
    /** `\label`s may be placed. */
    labels: boolean;
    /** Tallest an image may be drawn here. */
    imageMaxHeight: string;
}

/** An image written to the output: its path relative to the `.tex`, and whether LaTeX can include it. */
/** An image as the output refers to it; `bb` is set for one carried inside the .tex, whose size LaTeX then need not measure. */
interface MediaRef { path: string; includable: boolean; mime: string; intrinsic: { w: number; h: number } | null; bb?: string; }

/** One cell position in a laid-out table row. */
type TableSlot =
    | { kind: 'cell'; col: number; cell: OfficeContentNode; colSpan: number; rowSpan: number }
    | { kind: 'covered'; col: number; colSpan: number; origin: OfficeContentNode }
    | { kind: 'spill'; col: number; colSpan: number; origin: OfficeContentNode }
    | { kind: 'empty'; col: number };

/** Packages the body turned out to need, decided while rendering and read when writing the preamble. */
interface PackageUse {
    math: boolean; amssymb: boolean; graphics: boolean; tables: boolean; longtable: boolean; multirow: boolean;
    ulem: boolean; xcolor: boolean; colortbl: boolean; listings: boolean; endnotes: boolean; paragraphFix: boolean; captionof: boolean;
}

/**
 * Largest length written, in points. TeX's own limit is just under 16384pt ("Dimension too large"
 * past it); a page is far smaller, so anything near it is a hostile or corrupt value.
 */
const MAX_DIMENSION_PT = 16000;

/** Largest counter start written (`\setcounter` rejects numbers past 2^31-1 as "Number too big"). */
const MAX_LIST_START = 1000000;

/**
 * Formats a length with at most two decimals, clamped to TeX's range so a hostile value can neither
 * overflow it nor print in exponent notation. Deterministic across platforms.
 */
function fmtPt(n: number): string {
    const clamped = Number.isFinite(n) ? Math.max(-MAX_DIMENSION_PT, Math.min(MAX_DIMENSION_PT, n)) : 0;
    return `${Math.round(clamped * 100) / 100}pt`;
}

/** Reduces an anchor id or link target to a label name TeX and hyperref accept verbatim. */
function labelName(raw: string): string {
    // A character outside ASCII is written as its code point (`-u432-`), so ids in other scripts
    // stay distinct labels: as `-` they all ran together. A label too long to keep whole ends in a
    // hash of all of it, which keeps two that share a start apart.
    const name = String(raw ?? '').replace(/^#/, '').replace(/[^\x00-\x7F]/gu, ch => `-u${ch.codePointAt(0)!.toString(16)}-`).replace(/[^A-Za-z0-9:._-]/g, '-');
    if (name.length <= MAX_LABEL_LENGTH) return name;
    let hash = 0x811c9dc5;
    for (let i = 0; i < name.length; i++) hash = Math.imul(hash ^ name.charCodeAt(i), 0x01000193);
    return `${name.slice(0, MAX_LABEL_LENGTH - 9)}-${(hash >>> 0).toString(16).padStart(8, '0')}`;
}

/**
 * A citation key as `\cite` and `\bibitem` write it: the characters MarkdownParser's own citation
 * recognizer accepts, which also keep the key from closing the argument.
 */
function citationKeyName(key: unknown): string {
    return String(key ?? '').replace(/[^a-zA-Z0-9_:.-]/g, '');
}

/** Expands tabs to spaces at {@link CODE_TAB_WIDTH}-column stops. */
function expandTabs(line: string): string {
    let out = '';
    for (const ch of line) {
        if (ch === '\t') out += ' '.repeat(CODE_TAB_WIDTH - (out.length % CODE_TAB_WIDTH));
        else out += ch;
    }
    return out;
}

/**
 * Generates LaTeX source (`.tex`) from any AST, or with `texConfig.bundle` a zip of the source and
 * its images.
 *
 * The output compiles unmodified with pdfLaTeX, XeLaTeX, LuaLaTeX, upLaTeX, pLaTeX and `latex` (the
 * last three through dvipdfmx): the preamble selects fonts per engine, with TeX Live's fonts for the
 * scripts Latin Modern lacks where they are installed, and loads only the packages the body uses.
 * Every piece of document text is escaped ({@link escapeLatex}), URLs are scheme-checked and made
 * inert ({@link sanitizeLatexUrl}), and math, the one place content is emitted as live LaTeX, must
 * pass {@link sanitizeLatexMath} or it is written as literal text.
 *
 * Mapping, in brief: headings become sectioning commands (unnumbered unless `numberSections`),
 * flat list items are rebuilt into nested `itemize`/`enumerate`, tables become `longtable` (or
 * `tabular` where a `longtable` is not allowed) with `\multicolumn`/`\multirow` merges and ruled
 * grids, footnotes/endnotes become `\footnote`/`\endnote`, comments become LaTeX `%` comments, and
 * a presentation becomes `beamer` frames with speaker notes as `\note`.
 */
/** A title block found in the content (see `LatexGenerator.findTitleBlock`). */
interface TitleBlock {
    title: OfficeContentNode;
    subtitle?: OfficeContentNode;
    institute?: OfficeContentNode;
    author?: OfficeContentNode;
    date?: OfficeContentNode;
    /** The slide holding the block, or null when it is at the top level. */
    container: OfficeContentNode | null;
    nodes: Set<OfficeContentNode>;
}

export class LatexGenerator extends BaseGenerator<'tex'> {
    private readonly docClass: Exclude<TexDocumentClass, 'auto'>;
    private readonly beamer: boolean;
    private classSizePt = DEFAULT_CLASS_FONT_SIZE;
    private dominantSizePt: number | null = null;
    private ctx: RenderContext;

    private readonly uses: PackageUse = {
        math: false, amssymb: false, graphics: false, tables: false, longtable: false, multirow: false,
        ulem: false, xcolor: false, colortbl: false, listings: false, endnotes: false, paragraphFix: false, captionof: false,
    };

    /** Image files the output refers to by path (written into a bundle); `includable` when `\includegraphics` draws one. */
    private readonly media: { path: string; bytes: Uint8Array; includable: boolean }[] = [];
    /** Images carried inside the .tex, in order of first use: each becomes a `filecontents*` block. */
    private readonly carried: { name: string; pdf: string }[] = [];
    /** Carried images by content hash, so an image used twice is carried once (the text compared, not only the hash). */
    private readonly carriedByHash = new Map<string, { pdf: string; ref: MediaRef }[]>();
    /** What decoding images to carry them may cost this document, across all of them. */
    private readonly decodeBudget = newDecodeBudget();
    /** Images the document references by a relative path, with no image data: the caller supplies them. */
    private readonly externalImages = new Set<string>();
    private readonly mediaByAttachment = new Map<string, MediaRef | null>();
    /** Images the source embedded as `data:` URIs, by URI: they carry their bytes, as attachments do. */
    private readonly mediaByDataUri = new Map<string, MediaRef | null>();
    private readonly usedFileNames = new UniqueNames(name => name.toLowerCase());

    private readonly linkTargets = new Set<string>();
    private readonly definedLabels = new Set<string>();
    private readonly emittedLabels = new Set<string>();

    private deferredFootnotes: string[] = [];
    /** The macro holding the number of each note referred to more than once, set at its first reference (see noteMark). */
    private readonly noteNumbers = new Map<OfficeContentNode, string>();
    /** Notes referred to more than once whose first reference wrote them in parentheses. */
    private readonly parentheticalNotes = new Set<OfficeContentNode>();
    /** Texts of the notes a note's own text refers to, written as `\\footnotetext` after it (see noteMark). */
    private nestedFootnotes: string[] = [];
    /** Colors the body uses, as validated `RRGGBB` hex, each defined once by name. */
    private readonly colors = new Set<string>();
    /** The control words the formulas written use, for the packages and definitions they need (see planLatexMath). */
    private readonly mathCommands = new Set<string>();
    /** Every key the AST cites (found before rendering, for finding its bibliography), and those `\cite` wrote and `\bibitem` defined. */
    private readonly citationKeys = new Set<string>();
    private readonly citedKeys = new Set<string>();
    private readonly bibitemKeys = new Set<string>();
    private readonly warnedFeatures = new Set<string>();

    /** Set while rendering a heading whose every run is bold: the heading is bold already. */
    private inImplicitBold = false;
    /** Set while rendering a heading or frame title: the sectioning command decides the size. */
    private suppressSize = false;

    /**
     * The document's title block: a heading styled `Title` (a Word title, or what the LaTeX parser
     * reads from `\maketitle`) and the `Subtitle`/`Author`/`Date` lines right after it. It is
     * typeset with `\maketitle` (a `\titlepage` frame in beamer) where it stands.
     */
    private titleNodes: TitleBlock | null = null;
    /** The title block's `\title`/`\subtitle`/`\author`/`\date` arguments. */
    private titleFields: { title: string; subtitle: string; author: string; institute: string; date: string } | null = null;

    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'tex'>) {
        super('tex', ast, config);
        const requested = DOCUMENT_CLASSES.has(this.config.texConfig.documentClass) ? this.config.texConfig.documentClass : 'auto';
        this.docClass = requested === 'auto'
            ? (this.ast.content.some(n => n.type === 'slide') ? 'beamer' : 'article')
            : requested;
        this.beamer = this.docClass === 'beamer';
        this.ctx = {
            sections: !this.beamer,
            verbatim: true,
            longtable: !this.beamer,
            display: true,
            moving: false,
            notes: 'direct',
            labels: true,
            imageMaxHeight: this.beamer ? IMAGE_MAX_HEIGHT_FRAME : IMAGE_MAX_HEIGHT_PAGE,
        };
    }

    // ── entry point ─────────────────────────────────────────────────────────────────────────────

    async generate(): Promise<ConversionResult<'tex'>> {
        checkAbortSignal(this.config.abortSignal);
        this.prescan();
        this.titleNodes = this.findTitleBlock();
        if (this.titleNodes) this.titleFields = await this.renderTitleFields(this.titleNodes);

        const body = this.beamer ? await this.renderBeamerBody(this.ast.content) : await this.renderFlow(this.ast.content);
        const headerSetup = this.beamer ? this.warnBeamerHeaders() : await this.pageHeaderSetup();
        const endnotes = this.uses.endnotes
            ? (this.beamer ? '\\begin{frame}[allowframebreaks]\n\\theendnotes\n\\end{frame}' : '\\theendnotes')
            : '';

        const standalone = this.config.texConfig.standalone !== false;
        const parts = [body, endnotes].filter(Boolean);
        const meta = this.metadataCommands();
        const unicode = planLatexUnicode([headerSetup, meta.title, ...meta.hypersetup, ...parts].join('\n'));
        if (unicode.packages.has('amssymb')) this.uses.amssymb = true;
        const math = this.mathPlan();
        // A \cite with no \bibitem in the output prints [?] until the user adds a bibliography: said.
        const unresolved = [...this.citedKeys].filter(key => !this.bibitemKeys.has(key));
        if (unresolved.length) this.warn(OfficeWarningType.CITATIONS_NOT_RESOLVED, { keys: unresolved.slice(0, 20), more: Math.max(0, unresolved.length - 20) });
        let tex: string;
        const carried = this.carriedImageBlocks();
        if (standalone) {
            const titleBlock = this.titleBlock();
            tex = `${this.preamble(headerSetup, meta, unicode, math)}${carried ? `\n${carried}\n` : ''}\n\\begin{document}\n\n${[titleBlock, ...parts].filter(Boolean).join(BLOCK_SEPARATOR)}\n\n\\end{document}\n`;
        } else {
            // \definecolor, \providecommand and filecontents* are legal in the body, so a fragment
            // carries its own color definitions, math definitions and images.
            tex = `${this.fragmentHeader(unicode, math)}\n${[carried, this.colorDefinitions().join('\n'), math.definitions.join('\n'), ...parts].filter(Boolean).join(BLOCK_SEPARATOR)}\n`;
        }

        const bundle = this.config.texConfig.bundle === true;
        // The files the .tex reads: an image LaTeX cannot draw is a labelled box, reported on its own.
        const unbundled = bundle ? [] : this.media.filter(m => m.includable).map(m => m.path);
        if (unbundled.length > 0 || this.externalImages.size > 0) {
            const fromDataUris = !bundle && [...this.mediaByDataUri.values()].some(ref => ref?.includable && !ref.bb);
            this.warn(OfficeWarningType.IMAGES_NOT_BUNDLED, { files: unbundled, external: [...this.externalImages], fromDataUris, embedImages: this.config.texConfig.embedImages !== false });
        }
        if (bundle) {
            const { mtime } = resolveZipInstant(this.effectiveMetadata.modified);
            const files: Zippable = { [BUNDLE_MAIN_FILE]: new TextEncoder().encode(tex) };
            for (const m of this.media) files[m.path] = m.bytes;
            return { value: zipSync(files, { mtime }), messages: this.messages };
        }
        return { value: tex, messages: this.messages };
    }

    // ── pre-scan: labels, link targets, dominant font size ───────────────────────────────────────

    /**
     * Walks the whole AST once before rendering. Labels are only written for anchors something links
     * to (office documents carry hundreds of unused `_Toc`/`OLE_LINK` bookmarks), and a link is only
     * written when its target will exist, so the output has neither clutter nor undefined references.
     * The same pass finds the body font size that runs are measured against.
     */
    private prescan(): void {
        const sizeWeights = new Map<number, number>();
        const all: OfficeContentNode[] = [];
        // Each node once: a node the AST shares (a note every reference holds) was visited once per
        // path to it, which doubled per level of notes nested in notes.
        const seen = new Set<OfficeContentNode>();
        const visit = (n: OfficeContentNode | undefined) => {
            if (!n || seen.has(n)) return;
            seen.add(n);
            all.push(n);
            const meta = n.metadata as TextMetadata | undefined;
            // An image that is a link points at a target as a linked run does.
            if (n.type === 'image') {
                const image = n.metadata as ImageMetadata | undefined;
                if (image?.link && (image.linkType === 'internal' || image.link.startsWith('#'))) {
                    for (const candidate of this.labelCandidates(image.link)) this.linkTargets.add(candidate);
                }
            }
            if (n.type === 'text') {
                if (meta?.citationKey) {
                    const key = citationKeyName(meta.citationKey);
                    if (key) this.citationKeys.add(key);
                }
                if (meta?.link && (meta.linkType === 'internal' || meta.link.startsWith('#') || meta.wikilink)) {
                    for (const candidate of this.labelCandidates(meta.link)) this.linkTargets.add(candidate);
                }
                const pt = lengthToPt(n.formatting?.size);
                if (pt && pt > 0 && n.text) {
                    const key = Math.round(pt * 2) / 2;
                    sizeWeights.set(key, (sizeWeights.get(key) || 0) + n.text.length);
                }
            }
            for (const c of n.children || []) visit(c);
            for (const c of n.notes || []) visit(c);
            for (const c of n.comments || []) visit(c);
        };
        for (const n of this.ast.content) visit(n);
        for (const n of this.ast.auxiliary?.headers || []) visit(n);
        for (const n of this.ast.auxiliary?.footers || []) visit(n);

        const docSize = lengthToPt(this.ast.metadata?.formatting?.size);
        if (docSize && docSize > 0) {
            this.dominantSizePt = docSize;
        } else if (sizeWeights.size > 0) {
            this.dominantSizePt = [...sizeWeights.entries()].sort((a, b) => b[1] - a[1] || a[0] - b[0])[0][0];
        }
        if (this.dominantSizePt) {
            const d = this.dominantSizePt;
            this.classSizePt = CLASS_FONT_SIZES.reduce((best, s) => Math.abs(s - d) < Math.abs(best - d) ? s : best, DEFAULT_CLASS_FONT_SIZE);
        }

        if (this.config.ignoreInternalLinks) return;
        for (const n of all) {
            for (const id of ((n.metadata as any)?.anchorIds || []) as string[]) {
                const l = labelName(id);
                if (l && this.linkTargets.has(l)) this.definedLabels.add(l);
            }
            if (n.type === 'heading' && this.config.generateIds) {
                const l = this.headingSlugLabel(n);
                if (l) this.definedLabels.add(l);
            }
        }
    }

    /** Label names a link target may resolve to: the target as written, then its slug. */
    private labelCandidates(target: string): string[] {
        const raw = String(target || '').replace(/^#/, '');
        const out = [labelName(raw)];
        const slug = labelName(this.slugify(raw));
        if (slug && slug !== out[0]) out.push(slug);
        return out.filter(Boolean);
    }

    private headingSlugLabel(node: OfficeContentNode): string {
        return labelName(this.slugify(this.getNodeText(node)));
    }

    /** The label an internal link resolves to, or null when nothing in the output defines it. */
    private resolveLabel(target: string): string | null {
        for (const candidate of this.labelCandidates(target)) {
            if (this.definedLabels.has(candidate)) return candidate;
        }
        return null;
    }

    /**
     * The `\label`s for a node: its linked anchor ids and, for a heading, its generated id. Each label
     * is written once (a duplicate would be "multiply defined"); `phantom` adds the `\phantomsection`
     * a non-heading needs for hyperref to have an anchor at that spot.
     */
    private anchorsFor(node: OfficeContentNode, phantom: boolean): string {
        if (!this.ctx.labels || this.config.ignoreInternalLinks) return '';
        const names: string[] = [];
        for (const id of ((node.metadata as any)?.anchorIds || []) as string[]) names.push(labelName(id));
        if (node.type === 'heading' && this.config.generateIds) names.push(this.headingSlugLabel(node));
        const fresh = [...new Set(names)].filter(l => l && this.definedLabels.has(l) && !this.emittedLabels.has(l));
        if (fresh.length === 0) return '';
        for (const l of fresh) this.emittedLabels.add(l);
        return (phantom ? '\\phantomsection' : '') + fresh.map(l => `\\label{${l}}`).join('');
    }

    // ── context ──────────────────────────────────────────────────────────────────────────────────

    private async withCtx<T>(patch: Partial<RenderContext>, fn: () => Promise<T>): Promise<T> {
        const saved = this.ctx;
        this.ctx = { ...saved, ...patch };
        try { return await fn(); } finally { this.ctx = saved; }
    }

    /** Context for the body of a macro argument (footnote, note, list label, frame note). */
    private argumentCtx(): Partial<RenderContext> {
        return { sections: false, verbatim: false, longtable: false, moving: false, display: true };
    }

    /** A command name, `\protect`ed when inside a moving argument where it would otherwise break. */
    private cmd(name: string): string {
        return `${this.ctx.moving ? '\\protect' : ''}\\${name}`;
    }

    /**
     * The name of a defined color for a validated `RRGGBB` value. Colors are defined once, by name,
     * rather than written inline as `[HTML]{...}`: an optional argument inside a beamer frame title
     * is lost when a frame is continued over several frames, and named colors keep the body short.
     */
    private colorRef(hex: string): string {
        this.uses.xcolor = true;
        this.colors.add(hex);
        return `hex${hex}`;
    }

    private colorDefinitions(): string[] {
        return [...this.colors].sort().map(h => `\\definecolor{hex${h}}{HTML}{${h}}`);
    }

    private warnOnce(key: string, type: OfficeWarningType, info: unknown): void {
        if (this.warnedFeatures.has(key)) return;
        this.warnedFeatures.add(key);
        this.warn(type, info);
    }

    // ── block flow ───────────────────────────────────────────────────────────────────────────────

    /** Inline-level nodes a flow groups into a paragraph. */
    private isInlineNode(node: OfficeContentNode): boolean {
        if (node.type === 'text') return true;
        // A source comment sits in its run (a list item, a cell); one standing alone between blocks
        // forms a run of its own, which renders the same as a block.
        if (isSourceComment(node)) return true;
        if (node.type === 'break') {
            const t = (node.metadata as any)?.breakType;
            return t === undefined || t === 'textWrapping' || t === 'carriageReturn';
        }
        if (node.type === 'code') return (node.metadata as CodeMetadata)?.math === 'inline';
        return false;
    }

    /**
     * Renders a sequence of sibling nodes as blocks separated by blank lines. Runs of inline nodes
     * become one paragraph, runs of list items one nested list, runs of definition terms and
     * descriptions one `description` list; consecutive pages or slides get a page break between them.
     */
    private async renderFlow(nodes: OfficeContentNode[] | undefined): Promise<string> {
        const blocks: string[] = [];
        const items = nodes || [];
        let inline: OfficeContentNode[] = [];
        const flushInline = async () => {
            if (inline.length === 0) return;
            let p = (await this.renderInline(inline)).trim();
            // A run of nothing but source comments stands between blocks, whose blank line already ends
            // the comment line: no `{}` guard needed.
            if (inline.every(n => isSourceComment(n))) p = p.replace(/\n\{\}$/, '');
            if (p) blocks.push(p);
            inline = [];
        };
        let prevPaginated: string | null = null;
        for (let i = 0; i < items.length; i++) {
            checkAbortSignal(this.config.abortSignal);
            const node = items[i];
            if (this.isInlineNode(node)) { inline.push(node); continue; }
            await flushInline();

            if (this.titleNodes?.nodes.has(node)) {
                if (node !== this.titleNodes.title) continue;
                const override = await this.handleOnNode(node);
                if (override === false) continue;
                blocks.push(typeof override === 'string' ? override : `${this.commentsBefore(node)}${this.anchorsFor(node, true)}\\maketitle`);
                prevPaginated = null;
                continue;
            }

            if (node.type === 'list') {
                const listId = String((node.metadata as ListMetadata)?.listId ?? '');
                const run: OfficeContentNode[] = [];
                while (i < items.length && items[i].type === 'list' && String((items[i].metadata as ListMetadata)?.listId ?? '') === listId) run.push(items[i++]);
                i--;
                const list = this.isBibliography(run) ? await this.bibliographyList(run, null) : await this.renderListRun(run);
                if (list) blocks.push(list);
                prevPaginated = null;
                continue;
            }
            if (node.type === 'definitionTerm' || node.type === 'definitionDescription') {
                const run: OfficeContentNode[] = [];
                while (i < items.length && (items[i].type === 'definitionTerm' || items[i].type === 'definitionDescription')) run.push(items[i++]);
                i--;
                const list = await this.renderDescriptionRun(run);
                if (list) blocks.push(list);
                prevPaginated = null;
                continue;
            }

            const override = await this.handleOnNode(node);
            if (override === false) continue;
            if (typeof override === 'string') { blocks.push(override); prevPaginated = null; continue; }
            // A figure's or table's caption: `\captionof` for the table next to it, else the picture.
            if (node.type === 'paragraph' && (node.metadata as ParagraphMetadata | undefined)?.style === 'Caption' && this.ctx.sections) {
                const [prev, next] = [items[i - 1]?.type, items[i + 1]?.type];
                const kind = next === 'table' || next === 'sheet' ? 'table' : prev === 'image' ? 'figure' : prev === 'table' || prev === 'sheet' ? 'table' : 'figure';
                const caption = await this.caption(node, kind);
                if (caption) { blocks.push(caption); prevPaginated = null; continue; }
            }
            // A top-level heading over a bibliography is the heading thebibliography prints (see bibliographyList).
            if (node.type === 'heading' && this.ctx.sections && !this.beamer && ((node.metadata as HeadingMetadata)?.level ?? 1) === 1) {
                let end = i + 1;
                const listId = String((items[end]?.metadata as ListMetadata | undefined)?.listId ?? '');
                while (end < items.length && items[end].type === 'list' && String((items[end].metadata as ListMetadata)?.listId ?? '') === listId) end++;
                const run = items.slice(i + 1, end);
                if (this.isBibliography(run)) {
                    blocks.push(await this.bibliographyList(run, node));
                    i = end - 1;
                    prevPaginated = null;
                    continue;
                }
            }
            if ((node.type === 'page' || node.type === 'slide') && prevPaginated === node.type) blocks.push('\\clearpage');
            const out = (await this.renderBlockNode(node)).trim();
            if (out) blocks.push(out);
            prevPaginated = (node.type === 'page' || node.type === 'slide') ? node.type : null;
        }
        await flushInline();
        return blocks.join(BLOCK_SEPARATOR);
    }

    private async renderBlockNode(node: OfficeContentNode): Promise<string> {
        switch (node.type) {
            case 'paragraph': return this.paragraph(node);
            case 'heading': return this.heading(node);
            case 'list': return this.renderListRun([node]);
            case 'table': return this.table(node);
            case 'image': return this.blockImage(node);
            case 'code': return this.codeBlock(node);
            case 'break': return this.blockBreak(node);
            case 'note': return this.standaloneNote(node);
            case 'comment': return (isSourceComment(node) ? latexSourceComment(node.text || '') : this.commentLines(node)).trimEnd();
            case 'admonition': return this.admonition(node);
            case 'chart': return this.chart(node);
            case 'embed': return this.embed(node);
            case 'definitionList': return this.renderDescriptionRun(node.children || []);
            case 'definitionTerm':
            case 'definitionDescription':
                return this.renderDescriptionRun([node]);
            case 'sheet': return this.sheet(node);
            case 'slide': return this.slideInArticle(node);
            case 'text':
                return this.renderInline([node]);
            case 'page':
            case 'drawing':
            case 'header':
            case 'footer':
            case 'row':
            case 'cell':
                return this.renderFlow(node.children);
            case 'slideMaster':
                // Template furniture (placeholder layouts), not content.
                return '';
        }
    }

    /** Body blocks of a note/comment: its children, or a paragraph synthesized from its text. */
    private bodyBlocks(node: OfficeContentNode): OfficeContentNode[] {
        if (node.children && node.children.length) return node.children;
        return [{ type: 'paragraph', text: node.text || '', children: [{ type: 'text', text: node.text || '' }] } as OfficeContentNode];
    }

    /** Children to render inline for a paragraph-like node, falling back to its own text. */
    private inlineChildren(node: OfficeContentNode): OfficeContentNode[] {
        return node.children && node.children.length ? node.children : [{ type: 'text', text: node.text || '' } as OfficeContentNode];
    }

    // ── paragraphs & headings ────────────────────────────────────────────────────────────────────

    private async paragraph(node: OfficeContentNode): Promise<string> {
        const tag = this.getSemanticMapping(node)?.tag?.toLowerCase();
        const headingTag = tag ? /^h([1-6])$/.exec(tag) : null;
        if (headingTag) return this.heading(node, Number(headingTag[1]));
        if (tag === 'pre' || tag === 'code') return this.codeText(node.text || this.getNodeText(node), undefined, this.commentsBefore(node) + this.anchorsFor(node, true));

        const comments = this.commentsBefore(node);
        const anchors = this.anchorsFor(node, true);
        let body = (await this.renderInline(this.inlineChildren(node))).trim();
        body += await this.notesFor(node);
        if (!body) return (comments + anchors).trimEnd();
        body = this.alignParagraph(body, node.metadata as any);
        if (tag === 'blockquote') body = `\\begin{quote}\n${body}\n\\end{quote}`;
        return `${comments}${anchors}${body}`;
    }

    /**
     * A caption paragraph (a LaTeX float's `\caption`, what the parser reads it as) as `\captionof`
     * (`capt-of`), which numbers it as the figure's or table's caption, so it parses back as one. Not in
     * beamer, and '' for a caption whose text is numbered already (a Word caption, "Figure 1: ..."),
     * which stays a paragraph.
     */
    private async caption(node: OfficeContentNode, kind: 'figure' | 'table'): Promise<string> {
        if (this.beamer) return '';
        const text = this.getNodeText(node).trim();
        if (!text || /^(figure|fig\.?|table|tab\.?|listing|chart|exhibit|abbildung|tabelle|tableau|tabla|figura)\s*[\dIVXivx]/i.test(text)) return '';
        const runs = (await this.withCtx({ moving: true, display: false, verbatim: false, sections: false, longtable: false }, async () => (await this.renderInline(this.inlineChildren(node))) + await this.notesFor(node))).trim();
        if (!runs) return '';
        this.uses.captionof = true;
        return `${this.commentsBefore(node)}${this.anchorsFor(node, true)}\\captionof{${kind}}{${runs}}`;
    }

    /**
     * Applies a paragraph's alignment and indentation. Each is scoped to the paragraph with a group
     * ending in `\par`, so it cannot leak into what follows; left/justified text is LaTeX's default.
     */
    private alignParagraph(body: string, meta: any): string {
        if (this.config.includeFormatting === false || !meta || this.ctx.moving) return body;
        if (meta.alignment === 'center') return `{\\centering ${body}\\par}`;
        if (meta.alignment === 'right') return `{\\raggedleft ${body}\\par}`;
        const ind = meta.paragraphIndentation;
        if (!ind) return body;
        let setup = '';
        const left = Number(ind.left) > 0 ? Number(ind.left) / 20 : 0;
        const right = Number(ind.right) > 0 ? Number(ind.right) / 20 : 0;
        if (left) setup += `\\leftskip=${fmtPt(left)}`;
        if (right) setup += `\\rightskip=${fmtPt(right)}`;
        let lead = '';
        if (Number(ind.firstLine) > 0) lead = `\\hspace*{${fmtPt(Number(ind.firstLine) / 20)}}`;
        else if (Number(ind.hanging) > 0) setup += `\\hangindent=${fmtPt(Number(ind.hanging) / 20)}\\hangafter=1`;
        if (!setup && !lead) return body;
        return `{${setup}${setup ? ' ' : ''}${lead}${body}\\par}`;
    }

    private sectionCommand(level: number): string {
        const table = (this.docClass === 'report' || this.docClass === 'book') ? CHAPTER_SECTIONS : ARTICLE_SECTIONS;
        const cmd = table[Math.min(level, table.length) - 1];
        if (cmd === 'paragraph' || cmd === 'subparagraph') this.uses.paragraphFix = true;
        return cmd;
    }

    /** Renders runs with heading semantics: uniform bold and all sizes belong to the heading itself. */
    private async headingRuns(node: OfficeContentNode, uniformBold: boolean): Promise<string> {
        const savedBold = this.inImplicitBold, savedSize = this.suppressSize;
        this.inImplicitBold = uniformBold;
        this.suppressSize = true;
        try {
            return (await this.renderInline(this.inlineChildren(node))).trim();
        } finally {
            this.inImplicitBold = savedBold;
            this.suppressSize = savedSize;
        }
    }

    private async heading(node: OfficeContentNode, levelOverride?: number): Promise<string> {
        const meta = node.metadata as HeadingMetadata | undefined;
        const level = Math.max(1, Math.min(6, Math.floor(levelOverride ?? meta?.level ?? 1) || 1));
        const comments = this.commentsBefore(node);
        const uniformBold = this.hasUniformFormatting(node, f => f?.bold === true);

        if (this.ctx.sections) {
            const formatted = await this.withCtx({ moving: true, display: false, verbatim: false, sections: false, labels: false, longtable: false },
                async () => (await this.headingRuns(node, uniformBold)) + await this.notesFor(node));
            const labels = this.anchorsFor(node, false);
            if (!formatted) return (comments + labels).trimEnd();
            const plain = escapeLatex(this.getNodeText(node), ' ').trim();
            const short = plain && formatted !== plain ? `[{${plain}}]` : '';
            return `${comments}\\${this.sectionCommand(level)}${short}{${formatted}}${labels}`;
        }

        // Sectioning is not allowed here (a table cell, list item, note, quote or frame): keep the
        // heading's text as a bold line instead.
        const anchors = this.anchorsFor(node, true);
        const formatted = (await this.headingRuns(node, uniformBold)) + await this.notesFor(node);
        if (!formatted) return (comments + anchors).trimEnd();
        return `${comments}${anchors}\\textbf{${formatted}}`;
    }

    // ── inline runs ──────────────────────────────────────────────────────────────────────────────

    /**
     * Renders inline nodes. Consecutive text runs are collected before formatting so that runs with
     * identical formatting share one wrapper (a document split into many same-styled runs would
     * otherwise read `\textbf{a}\textbf{b}`), and consecutive runs with the same link share one
     * `\href`. A run's notes and comments follow it directly.
     */
    private async renderInline(nodes: OfficeContentNode[]): Promise<string> {
        let out = '';
        // The last character written, so a comment can tell whether a space precedes it without
        // reading back `out`, which would copy all of it at each comment.
        let lastWritten = '';
        const write = (text: string) => {
            if (!text) return;
            out += text;
            lastWritten = text[text.length - 1];
        };
        let runs: OfficeContentNode[] = [];
        let group: OfficeContentNode[] = [];
        let groupLink = '';
        const flushRuns = () => {
            if (runs.length === 0) return;
            write(this.formatRuns(runs));
            runs = [];
        };
        const flushGroup = async () => {
            if (group.length === 0) return;
            write(await this.hyperlink(groupLink, group));
            group = [];
            groupLink = '';
        };
        const list = nodes || [];
        // Where `out` ended when a `%` comment line was last written: a run that ends in one gets `{}`
        // after it, since callers trim a run and the `%` would then swallow what they append on the
        // same line (a heading's closing brace, a cell's `&` or `\\`).
        let commentEnd = -1;
        for (let i = 0; i < list.length; i++) {
            const node = list[i];
            checkAbortSignal(this.config.abortSignal);
            const override = await this.handleOnNode(node);
            if (override === false) continue;
            if (typeof override === 'string') { flushRuns(); await flushGroup(); write(override); continue; }
            if (isSourceComment(node)) {
                // A hidden note inside a run: `% <!--...-->` lines. The `%` swallows its line end and TeX
                // skips the next line's leading spaces, so a space that follows the comment is written in
                // front of it instead, or "a<!-- x --> b" would typeset as "ab".
                flushRuns();
                await flushGroup();
                const next = list[i + 1];
                const prev = list[i - 1];
                const spaceAfter = next?.type === 'text' && /^\s/.test(next.text || '');
                const spaceBefore = /\s/.test(lastWritten) || (prev?.type === 'text' && /\s$/.test(prev.text || ''));
                write(`${spaceAfter && !spaceBefore ? ' ' : ''}${latexSourceComment(node.text || '')}`);
                commentEnd = out.length;
                continue;
            }
            const meta = node.metadata as TextMetadata | undefined;
            if (node.type === 'text' && !meta?.citationKey) {
                if (meta?.link) {
                    flushRuns();
                    if (!(group.length && groupLink === meta.link)) { await flushGroup(); groupLink = meta.link; }
                    group.push(node);
                    continue;
                }
                await flushGroup();
                runs.push(node);
                if (node.notes?.length || node.comments?.length) {
                    flushRuns();
                    write(this.inlineComments(node));
                    if (node.comments?.length) commentEnd = out.length;
                    write(await this.notesFor(node));
                }
                continue;
            }
            flushRuns();
            await flushGroup();
            write(await this.inlineNode(node));
        }
        flushRuns();
        await flushGroup();
        if (commentEnd === out.length) out += '{}';
        return out;
    }

    /** Formats text runs, merging neighbours whose formatting is identical into one wrapped span. */
    private formatRuns(nodes: OfficeContentNode[]): string {
        const key = (f: TextFormatting | undefined) => JSON.stringify(Object.entries(f || {}).filter(([, v]) => v !== undefined && v !== false).sort(([a], [b]) => a.localeCompare(b)));
        let out = '';
        let i = 0;
        while (i < nodes.length) {
            const fmt = nodes[i].formatting;
            const k = key(fmt);
            let raw = '';
            while (i < nodes.length && key(nodes[i].formatting) === k) raw += nodes[i++].text || '';
            out += this.formatRun(escapeLatex(raw, this.ctx.moving ? ' ' : LINE_BREAK), raw, fmt);
        }
        return out;
    }

    private async inlineNode(node: OfficeContentNode): Promise<string> {
        switch (node.type) {
            case 'text':
                return this.textRun(node) + this.inlineComments(node) + await this.notesFor(node);
            case 'code': {
                const meta = node.metadata as CodeMetadata | undefined;
                if (meta?.math === 'inline') return this.inlineMath(node.text || '');
                if (meta?.math === 'block') return this.ctx.display ? `\n${this.displayMath(node.text || '')}\n` : this.inlineMath(node.text || '');
                return `${this.cmd('texttt')}{${escapeLatex(node.text || '', this.ctx.moving ? ' ' : LINE_BREAK)}}`;
            }
            case 'image': return this.imageMarkup(node, false);
            case 'break': return this.inlineBreak(node);
            default:
                return node.children && node.children.length
                    ? this.renderInline(node.children)
                    : escapeLatex(node.text || '', this.ctx.moving ? ' ' : LINE_BREAK);
        }
    }

    private inlineBreak(node: OfficeContentNode): string {
        const t = (node.metadata as any)?.breakType;
        if (this.ctx.moving) return ' ';
        if (t === 'page') return '\\newpage{}';
        if (t === 'lastRenderedPage') return '';
        if (t === 'column' || t === 'thematic') return ' ';
        return LINE_BREAK;
    }

    private textRun(node: OfficeContentNode): string {
        const meta = node.metadata as TextMetadata | undefined;
        if (meta?.citationKey) {
            const key = citationKeyName(meta.citationKey);
            if (key) {
                this.citedKeys.add(key);
                return `${this.cmd('cite')}{${key}}`;
            }
        }
        const raw = node.text || '';
        return this.formatRun(escapeLatex(raw, this.ctx.moving ? ' ' : LINE_BREAK), raw, node.formatting);
    }

    /**
     * Wraps escaped text in the commands for its run formatting. Innermost to outermost: highlight,
     * typewriter, size, bold, italic, underline, strikeout, super/subscript, color. Whitespace-only
     * runs carry no visible formatting and are left bare.
     */
    private formatRun(escaped: string, raw: string, fmt: TextFormatting | undefined): string {
        if (!escaped) return '';
        if (!fmt || this.config.includeFormatting === false || !raw.trim()) return escaped;
        let t = escaped;
        const bg = hexColor(fmt.backgroundColor);
        if (bg) {
            // A \colorbox cannot break across lines, so a highlighted sentence would run off the page.
            // Boxing each word keeps the highlight and lets the line break between words. Applied to
            // the bare escaped text, where a space is always a word boundary (none of the escapes
            // contain one), before any wrapper adds braces a split could cut through.
            t = `{${this.cmd('setlength')}{\\fboxsep}{1pt}${t.split(/( +)/).map(w => (w.trim() ? `${this.cmd('colorbox')}{${this.colorRef(bg)}}{${this.cmd('strut')} ${w}}` : w)).join('')}}`;
        }
        if (fmt.font && MONOSPACE_FONTS.has(fmt.font.trim().toLowerCase())) t = `${this.cmd('texttt')}{${t}}`;
        const size = this.suppressSize ? null : this.scaledSize(fmt.size);
        if (size) t = `{${this.cmd('fontsize')}{${fmtPt(size)}}{${fmtPt(size * LEADING_FACTOR)}}${this.cmd('selectfont')} ${t}}`;
        if (fmt.bold && !this.inImplicitBold) t = `${this.cmd('textbf')}{${t}}`;
        if (fmt.italic) t = `${this.cmd('textit')}{${t}}`;
        if (fmt.underline) { this.uses.ulem = true; t = `${this.cmd('uline')}{${t}}`; }
        if (fmt.strikethrough) { this.uses.ulem = true; t = `${this.cmd('sout')}{${t}}`; }
        if (fmt.superscript) t = `${this.cmd('textsuperscript')}{${t}}`;
        else if (fmt.subscript) t = `${this.cmd('textsubscript')}{${t}}`;
        const color = hexColor(fmt.color);
        if (color) t = `${this.cmd('textcolor')}{${this.colorRef(color)}}{${t}}`;
        return t;
    }

    /**
     * A run's size, rescaled from the source's body size to the class's (so a 14pt run in a 12pt
     * document written at 11pt stays proportionally larger), or null when it is the body size.
     */
    private scaledSize(size: string | undefined): number | null {
        const pt = lengthToPt(size);
        if (!pt || pt <= 0) return null;
        const base = this.dominantSizePt || this.classSizePt;
        const scaled = Math.round((pt * this.classSizePt / base) * 2) / 2;
        if (Math.abs(scaled - this.classSizePt) < 0.5) return null;
        return Math.max(MIN_FONT_SIZE_PT, Math.min(MAX_FONT_SIZE_PT, scaled));
    }

    private async hyperlink(link: string, group: OfficeContentNode[]): Promise<string> {
        const inner = this.formatRuns(group);
        let trailing = '';
        for (const n of group) trailing += this.inlineComments(n) + await this.notesFor(n);
        const meta = group[0].metadata as TextMetadata | undefined;
        const internal = meta?.linkType === 'internal' || link.startsWith('#') || meta?.wikilink === true;
        if (!inner) return trailing;
        if (internal) {
            if (this.config.ignoreInternalLinks) return inner + trailing;
            const label = this.resolveLabel(link);
            return (label ? `${this.cmd('hyperref')}[${label}]{${inner}}` : inner) + trailing;
        }
        const url = sanitizeLatexUrl(link);
        return (url ? `${this.cmd('href')}{${url}}{${inner}}` : inner) + trailing;
    }

    // ── math ─────────────────────────────────────────────────────────────────────────────────────

    /**
     * A safe formula as LaTeX: the KaTeX and MathJax macros LaTeX lacks written as what they stand for,
     * and its commands noted, for the packages and definitions the preamble gives them (see mathPlan).
     */
    private mathLatex(latex: string): string {
        const out = withLatexMathMacros(latex);
        mathCommandsOf(out, this.mathCommands);
        return out;
    }

    private inlineMath(source: string): string {
        this.uses.math = true;
        const result = sanitizeLatexMath(source, 'inline');
        if (result.ok) return result.latex ? `$${this.mathLatex(result.latex)}$` : '';
        this.warn(OfficeWarningType.MATH_WRITTEN_AS_TEXT, { commands: result.commands });
        return `\\texttt{${escapeLatex(source, ' ')}}`;
    }

    private displayMath(source: string): string {
        this.uses.math = true;
        const result = sanitizeLatexMath(source, 'block');
        if (result.ok) {
            if (!result.latex) return '';
            const latex = this.mathLatex(result.latex);
            return result.displayEnvironment ? latex : `\\[\n${latex}\n\\]`;
        }
        this.warn(OfficeWarningType.MATH_WRITTEN_AS_TEXT, { commands: result.commands });
        return this.codeText(source, undefined, '');
    }

    // ── notes & comments ─────────────────────────────────────────────────────────────────────────

    private async notesFor(node: OfficeContentNode): Promise<string> {
        if (!node.notes || node.notes.length === 0 || node.type === 'slide') return '';
        let out = '';
        for (const note of node.notes) out += await this.noteMark(note);
        return out;
    }

    /**
     * A note's body. Its runs' sizes are dropped: LaTeX sets footnotes and endnotes in its own
     * smaller size, and a source note's explicit size (Word's 10pt footnote text) says the same.
     */
    private async noteBody(note: OfficeContentNode, inner: RenderContext['notes']): Promise<string> {
        const savedSize = this.suppressSize;
        this.suppressSize = true;
        try {
            return (await this.withCtx({ ...this.argumentCtx(), notes: inner }, () => this.renderFlow(this.bodyBlocks(note)))).trim();
        } finally {
            this.suppressSize = savedSize;
        }
    }

    /**
     * A footnote/endnote at its reference point. Inside `tabular` a `\footnote` would lose its text,
     * so there the mark is written now and the text is queued for `\footnotetext` after the table.
     */
    private async noteMark(note: OfficeContentNode): Promise<string> {
        if (this.ctx.notes === 'omit') return '';
        const meta = note.metadata as NoteMetadata | undefined;
        // A note referred to again: a mark with the number its first reference got (kept in a macro
        // there), its text written once. Written at every reference, one note a small document
        // refers to thousands of times made output of gigabytes. Where its first reference had no
        // number (in parentheses), the text is not repeated.
        const endnote = meta?.noteType === 'endnote';
        const numbered = this.noteNumbers.get(note);
        if (numbered !== undefined) {
            if (this.ctx.notes === 'parenthetical') return '';
            return endnote ? `\\endnotemark[${numbered}]` : `\\footnotemark[${numbered}]`;
        }
        if (this.parentheticalNotes.has(note)) return '';
        // The macro that will hold its number, when it is referred to more than once.
        const shared = this.noteReferences(note) > 1 ? `\\opNote${letterIndex(this.noteNumbers.size)}` : undefined;
        const keep = (counter: string) => shared ? `\\xdef${shared}{\\the\\value{${counter}}}` : '';
        if (shared && this.ctx.notes !== 'parenthetical') this.noteNumbers.set(note, shared);
        // A note a note's text refers to was left out. In a footnote written in place it is a mark,
        // its text after that footnote; where no mark can be numbered with its text (in an endnote,
        // printed at the end, or a footnote whose text follows a table) it is in parentheses.
        if (this.ctx.notes === 'nested') {
            this.nestedFootnotes.push(await this.noteBody(note, 'parenthetical'));
            return `\\footnotemark{}${keep('footnote')}`;
        }
        if (this.ctx.notes === 'parenthetical') {
            if (shared) this.parentheticalNotes.add(note);
            return ` (${await this.noteBody(note, 'parenthetical')})`;
        }
        const direct = meta?.noteType !== 'endnote' && this.ctx.notes === 'direct';
        const outerNested = this.nestedFootnotes;
        this.nestedFootnotes = [];
        let body: string;
        let nested: string[];
        try {
            body = await this.noteBody(note, direct ? 'nested' : 'parenthetical');
        } finally {
            nested = this.nestedFootnotes;
            this.nestedFootnotes = outerNested;
        }
        // The texts of the notes this note's text refers to, numbered to match their marks in it.
        const nestedTexts = nested.length
            ? `\\addtocounter{footnote}{-${nested.length}}${nested.map(b => `\\stepcounter{footnote}\\footnotetext{${b}}`).join('')}`
            : '';
        if (meta?.noteType === 'endnote') {
            this.uses.endnotes = true;
            // (Its number is the counter's after `\endnote`, whose text is printed later.)
            return `${this.cmd('endnote')}{${body}}${keep('endnote')}${nestedTexts}`;
        }
        if (this.ctx.notes === 'deferred') {
            this.deferredFootnotes.push(body);
            return `\\footnotemark{}${keep('footnote')}`;
        }
        // (Its number is kept at the start of its text, before the marks of the notes it refers to.)
        return `${this.cmd('footnote')}{${keep('footnote')}${body}}${nestedTexts}`;
    }

    /** A note with no reference point (an orphan definition): its mark on a line of its own. */
    private async standaloneNote(node: OfficeContentNode): Promise<string> {
        return this.noteMark(node);
    }

    /** Writes the `\footnotetext`s queued inside a `tabular`, numbered to match their marks. */
    private flushDeferredFootnotes(): string {
        if (this.deferredFootnotes.length === 0) return '';
        const n = this.deferredFootnotes.length;
        const texts = this.deferredFootnotes.map(b => `\\stepcounter{footnote}\\footnotetext{${b}}`).join('\n');
        this.deferredFootnotes = [];
        return `\n\\addtocounter{footnote}{-${n}}\n${texts}`;
    }

    /** Plain text of a comment/note body, one line per block. */
    private plainText(node: OfficeContentNode): string {
        if (node.children && node.children.length) {
            const blockish = node.children.some(c => c.type !== 'text' && c.type !== 'break');
            if (blockish) return node.children.map(c => this.plainText(c)).filter(Boolean).join('\n');
            return node.children.map(c => c.type === 'break' ? '\n' : (c.text || this.getNodeText(c))).join('');
        }
        return node.text || '';
    }

    /**
     * A comment as LaTeX `%` lines. Review comments are annotations about the document rather than
     * part of it, and LaTeX has no native annotation; a source comment keeps them for the author
     * without changing the typeset page, and it is safe in every context, even inside an argument.
     */
    private commentLines(comment: OfficeContentNode): string {
        // Written once, at its first reference (see firstWriteOfComment).
        if (!this.firstWriteOfComment(comment)) return '';
        const meta = comment.metadata as CommentMetadata | undefined;
        const who = [meta?.author, meta?.date].filter(Boolean).join(', ');
        return latexComment(`Comment${who ? ` (${who})` : ''}: ${this.plainText(comment)}`);
    }

    private commentsBefore(node: OfficeContentNode): string {
        return (node.comments || []).map(c => this.commentLines(c)).join('');
    }

    /**
     * Comments anchored to a run, written right after it. The `%` swallows its own line end, so the
     * text continues on the next line with no extra space.
     */
    private inlineComments(node: OfficeContentNode): string {
        if (!node.comments || node.comments.length === 0) return '';
        return node.comments.map(c => this.commentLines(c)).join('');
    }

    // ── lists ────────────────────────────────────────────────────────────────────────────────────

    /**
     * Rebuilds nested `itemize`/`enumerate` environments from a run of flat list items. Each item
     * carries its own level and type, so the environment stack is opened and closed as the level
     * changes; a jump of more than one level gets empty `\item[]`s to hang the deeper list on.
     */
    private async renderListRun(items: OfficeContentNode[]): Promise<string> {
        const stack: ('itemize' | 'enumerate')[] = [];
        let out = '';
        const open = (env: 'itemize' | 'enumerate', startIndex: number | undefined) => {
            stack.push(env);
            out += `\\begin{${env}}\n`;
            if (env === 'enumerate' && typeof startIndex === 'number' && startIndex > 0) {
                const depth = stack.filter(e => e === 'enumerate').length;
                out += `\\setcounter{${ENUM_COUNTERS[Math.min(depth, ENUM_COUNTERS.length) - 1]}}{${Math.min(MAX_LIST_START, Math.floor(startIndex))}}\n`;
            }
        };
        const close = () => { out += `\\end{${stack.pop()}}\n`; };

        for (const item of items) {
            checkAbortSignal(this.config.abortSignal);
            const override = await this.handleOnNode(item);
            if (override === false) continue;
            const meta = item.metadata as ListMetadata | undefined;
            const maxDepth = this.beamer ? MAX_LIST_DEPTH_BEAMER : MAX_LIST_DEPTH;
            const level = Math.max(0, Math.min(maxDepth - 1, Math.floor(Number(meta?.indentation) || 0)));
            const env: 'itemize' | 'enumerate' = meta?.listType === 'ordered' ? 'enumerate' : 'itemize';

            while (stack.length > level + 1) close();
            if (stack.length === level + 1 && stack[level] !== env) close();
            while (stack.length < level + 1) {
                const isTarget = stack.length === level;
                open(isTarget ? env : 'itemize', isTarget ? meta?.itemIndex : undefined);
                if (!isTarget) out += '\\item[]\n';
            }

            if (typeof override === 'string') { out += `\\item ${override}\n`; continue; }

            let label = '';
            if (meta?.isTask) {
                this.uses.math = true;
                this.uses.amssymb = true;
                label = meta.checked ? '[$\\boxtimes$]' : '[$\\square$]';
            }
            const comments = this.commentsBefore(item);
            const anchors = this.anchorsFor(item, true);
            const content = (await this.withCtx({ sections: false, longtable: false }, () => this.renderFlow(this.inlineChildren(item)))).trim();
            const notes = await this.notesFor(item);
            out += `${comments}\\item${label} ${anchors}${content}${notes}\n`;
        }
        while (stack.length > 0) close();
        return out.trimEnd();
    }

    /**
     * Whether a run of list items is a bibliography: numbered entries, each anchored by the key a
     * citation names it by (what the LaTeX parser reads `thebibliography` and a `.bib` database as),
     * at least one of them cited.
     */
    private isBibliography(run: OfficeContentNode[]): boolean {
        if (!run.length || !this.citationKeys.size) return false;
        let cited = false;
        for (const item of run) {
            const meta = item.metadata as ListMetadata | undefined;
            const key = citationKeyName(meta?.anchorIds?.[0]);
            if (item.type !== 'list' || meta?.listType !== 'ordered' || Number(meta?.indentation) > 0 || !key) return false;
            if (this.citationKeys.has(key)) cited = true;
        }
        return cited;
    }

    /**
     * A bibliography as `thebibliography`, each entry a `\bibitem` under its key, so the document's
     * `\cite`s resolve to it. The list prints its own heading, `\refname` (`\bibname` in a class with
     * chapters): the heading over it (none when there is none) is set as that name, where it is not
     * the name already.
     */
    private async bibliographyList(run: OfficeContentNode[], heading: OfficeContentNode | null): Promise<string> {
        const lines: string[] = [];
        if (!this.beamer) {
            const chapters = this.docClass === 'report' || this.docClass === 'book';
            const title = heading
                ? await this.withCtx({ moving: true, display: false, verbatim: false, sections: false, labels: false, longtable: false },
                    () => this.headingRuns(heading, this.hasUniformFormatting(heading, f => f?.bold === true)))
                : '';
            const before = heading ? `${this.commentsBefore(heading)}${this.anchorsFor(heading, true)}`.trimEnd() : '';
            if (before) lines.push(before);
            if (title !== (chapters ? 'Bibliography' : 'References')) lines.push(`\\renewcommand{\\${chapters ? 'bibname' : 'refname'}}{${title}}`);
        }
        lines.push(`\\begin{thebibliography}{${'9'.repeat(String(run.length).length)}}`);
        for (const item of run) {
            checkAbortSignal(this.config.abortSignal);
            const override = await this.handleOnNode(item);
            if (override === false) continue;
            const key = citationKeyName((item.metadata as ListMetadata).anchorIds![0]);
            this.bibitemKeys.add(key);
            const content = typeof override === 'string'
                ? override
                : (await this.withCtx({ sections: false, longtable: false }, () => this.renderFlow(this.inlineChildren(item)))).trim() + await this.notesFor(item);
            lines.push(`${this.commentsBefore(item)}\\bibitem{${key}} ${this.anchorsFor(item, true)}${content}`);
        }
        lines.push('\\end{thebibliography}');
        return lines.join('\n');
    }

    /** A run of definition terms and descriptions as one `description` list. */
    private async renderDescriptionRun(nodes: OfficeContentNode[]): Promise<string> {
        let out = '';
        let pendingTerm: string | null = null;
        const flushTerm = () => { if (pendingTerm !== null) { out += `\\item[{${pendingTerm}}]\n`; pendingTerm = null; } };
        for (const node of nodes) {
            checkAbortSignal(this.config.abortSignal);
            const override = await this.handleOnNode(node);
            if (override === false) continue;
            if (typeof override === 'string') { flushTerm(); out += `\\item[] ${override}\n`; continue; }
            if (node.type === 'definitionTerm') {
                flushTerm();
                pendingTerm = (await this.withCtx({ ...this.argumentCtx() }, () => this.renderInline(this.inlineChildren(node)))).trim();
                continue;
            }
            const body = node.type === 'definitionDescription'
                ? await this.withCtx({ sections: false, longtable: false }, () => this.renderFlow(this.inlineChildren(node)))
                : await this.withCtx({ sections: false, longtable: false }, () => this.renderBlockNode(node));
            const label = pendingTerm !== null ? `[{${pendingTerm}}]` : '[]';
            pendingTerm = null;
            out += `\\item${label} ${body.trim()}\n`;
        }
        flushTerm();
        return out ? `\\begin{description}\n${out}\\end{description}` : '';
    }

    // ── tables ───────────────────────────────────────────────────────────────────────────────────

    /**
     * Lays a table's rows out on a grid: each cell at its column, spans reserving the positions they
     * cover, and gaps (a sparse spreadsheet row, a row shorter than the grid) filled with empty slots.
     */
    private layoutTable(rows: OfficeContentNode[]): { cols: number; grid: TableSlot[][] } {
        const cellsOf = (row: OfficeContentNode) => (row.children || []).filter(c => c.type === 'cell');
        let cols = 1;
        for (const row of rows) {
            let width = 0;
            for (const c of cellsOf(row)) {
                const m = c.metadata as any;
                const start = typeof m?.col === 'number' && m.col >= 0 ? Math.max(width, m.col) : width;
                width = start + Math.max(1, Math.min(MAX_TABLE_COLUMNS, Math.floor(Number(m?.colSpan)) || 1));
            }
            cols = Math.max(cols, width);
        }
        cols = Math.min(MAX_TABLE_COLUMNS, cols);

        const active = new Map<number, { remaining: number; colSpan: number; origin: OfficeContentNode }>();
        const grid: TableSlot[][] = [];
        rows.forEach((row, ri) => {
            const cells = cellsOf(row);
            const slots: TableSlot[] = [];
            let col = 0;
            let ci = 0;
            while (col < cols) {
                const span = active.get(col);
                if (span) {
                    slots.push({ kind: 'covered', col, colSpan: span.colSpan, origin: span.origin });
                    span.remaining--;
                    if (span.remaining <= 0) active.delete(col);
                    col += span.colSpan;
                    continue;
                }
                if (ci >= cells.length) {
                    // The rest of a row no span reaches is padded within the document's budget (see
                    // padWithinBudget); a tabular reads a short row as ending in empty cells.
                    let pending = false;
                    for (const start of active.keys()) if (start >= col) { pending = true; break; }
                    if (!pending) {
                        const fill = this.padWithinBudget(cols - col);
                        for (let k = 0; k < fill; k++) slots.push({ kind: 'empty', col: col + k });
                        break;
                    }
                    slots.push({ kind: 'empty', col }); col++; continue;
                }
                const m = cells[ci].metadata as any;
                if (typeof m?.col === 'number' && m.col > col) { slots.push({ kind: 'empty', col }); col++; continue; }
                const cell = cells[ci++];
                let colSpan = Math.max(1, Math.floor(Number(m?.colSpan)) || 1);
                // A span may not run past the grid or into a column a row span above still holds.
                let limit = cols - col;
                for (const start of active.keys()) if (start > col) limit = Math.min(limit, start - col);
                colSpan = Math.min(colSpan, limit);
                const rowSpan = Math.max(1, Math.min(MAX_ROW_SPAN, rows.length - ri, Math.floor(Number(m?.rowSpan)) || 1));
                slots.push({ kind: 'cell', col, cell, colSpan, rowSpan });
                if (rowSpan > 1) active.set(col, { remaining: rowSpan - 1, colSpan, origin: cell });
                col += colSpan;
            }
            // Cells past the widest table LaTeX output lays out are not written: said, not dropped silently.
            if (ci < cells.length) this.warnOnce('table:columns', OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: `table cells past column ${MAX_TABLE_COLUMNS}`, format: 'tex' });
            grid.push(slots);
        });
        return { cols, grid };
    }

    /** Width of a column (or of `span` merged columns) in a ruled `cols`-column table filling the line. */
    private columnWidth(cols: number, span: number): string {
        const avail = `(\\linewidth-${2 * cols}\\tabcolsep-${cols + 1}\\arrayrulewidth)`;
        if (span === 1) return `\\dimexpr${avail}/${cols}\\relax`;
        return `\\dimexpr${avail}*${span}/${cols}+${2 * (span - 1)}\\tabcolsep+${span - 1}\\arrayrulewidth\\relax`;
    }

    private columnSpec(cols: number, span: number, leftRule: boolean, align?: string): string {
        const setup = align === 'center' ? '\\centering' : align === 'right' ? '\\raggedleft' : '\\raggedright';
        return `${leftRule ? '|' : ''}>{${setup}\\arraybackslash}p{${this.columnWidth(cols, span)}}|`;
    }

    /** A cell's own alignment (`CellMetadata.align`), left when it states none. */
    private cellAlign(cell: OfficeContentNode): 'left' | 'center' | 'right' {
        const a = (cell.metadata as any)?.align;
        return a === 'center' || a === 'right' ? a : 'left';
    }

    /**
     * Per-column alignment (`CellMetadata.align`, which a Markdown pipe table's separator row sets, a
     * LaTeX column specification gives and HTML's `align`/`text-align` carry): the alignment most of
     * the column's own cells have, a tie going to the one met first. A merged cell does not vote (a
     * centred header over a left and a right column is no evidence for either), and a cell that differs
     * from its column is written with its own (see renderSlot). Null when every column is left-aligned,
     * so the common case keeps one compact repeated column type.
     */
    private columnAlignments(grid: TableSlot[][], cols: number): ('left' | 'center' | 'right')[] | null {
        if (this.config.includeFormatting === false || grid.length === 0) return null;
        const votes = Array.from({ length: cols }, () => ({ left: 0, center: 0, right: 0, first: [] as ('left' | 'center' | 'right')[] }));
        for (const row of grid) {
            for (const slot of row) {
                if (slot.kind !== 'cell' || slot.colSpan !== 1 || slot.col >= cols) continue;
                const a = this.cellAlign(slot.cell);
                const v = votes[slot.col];
                if (v[a]++ === 0) v.first.push(a);
            }
        }
        const aligns = votes.map(v => v.first.reduce<'left' | 'center' | 'right'>((best, a) => (v[a] > v[best] ? a : best), v.first[0] ?? 'left'));
        return aligns.some(a => a !== 'left') ? aligns : null;
    }

    private async table(node: OfficeContentNode, rowsOverride?: OfficeContentNode[]): Promise<string> {
        const rows = (rowsOverride ?? (node.children || [])).filter(r => r.type === 'row');
        if (rows.length === 0) return '';
        this.uses.tables = true;
        const { cols, grid } = this.layoutTable(rows);
        // The row callback runs once per row, however many column bands the row is split over.
        const rowOverrides: (string | false | void)[] = [];
        for (const row of rows) rowOverrides.push(await this.handleOnNode(row));

        const bands: string[] = [];
        for (let start = 0; start < cols; start += MAX_TABLE_BAND_COLUMNS) {
            const end = Math.min(cols, start + MAX_TABLE_BAND_COLUMNS);
            let bandGrid = grid.map(slots => this.sliceBand(slots, start, end));
            let bandRows = rows, bandOverrides = rowOverrides;
            if (start > 0) {
                // A band after the first holds the rows with something in its columns: holding every
                // row, one wide row over 100,000 narrow ones wrote 100,000 empty rows in each of 63 bands
                // (72 MB from 18 KB). A row a span covers there is kept, so rules and spans still meet.
                const kept = bandGrid.map((slots, ri) => typeof rowOverrides[ri] === 'string' || slots.some(slot => slot.kind !== 'empty'));
                bandRows = rows.filter((_, ri) => kept[ri]);
                bandOverrides = rowOverrides.filter((_, ri) => kept[ri]);
                bandGrid = bandGrid.filter((_, ri) => kept[ri]);
                if (!bandRows.length) continue;
            }
            const band = await this.tableBand(bandRows, bandOverrides, bandGrid, end - start, start === 0);
            // A table wider than one band continues below itself, a band at a time.
            bands.push(start === 0 ? band : `\\textit{(continued: columns ${start + 1}--${end})}\n\n${band}`);
        }
        return `${this.commentsBefore(node)}${this.anchorsFor(node, true)}${bands.join(BLOCK_SEPARATOR)}`;
    }

    /**
     * The part of a laid-out row that falls in columns [start, end), renumbered from 0. A span is
     * clipped to the band; the part of a cell that began in an earlier band becomes a `spill`, drawn
     * as an empty merged cell (it is not a row span, so it must not suppress the rule above it).
     */
    private sliceBand(slots: TableSlot[], start: number, end: number): TableSlot[] {
        const out: TableSlot[] = [];
        for (const slot of slots) {
            const span = slot.kind === 'empty' ? 1 : slot.colSpan;
            const a = Math.max(slot.col, start);
            const b = Math.min(slot.col + span, end);
            if (a >= b) continue;
            if (slot.kind === 'empty') out.push({ kind: 'empty', col: a - start });
            else if (slot.kind === 'cell' && slot.col >= start) out.push({ ...slot, col: a - start, colSpan: b - a });
            else if (slot.kind === 'cell') out.push({ kind: 'spill', col: a - start, colSpan: b - a, origin: slot.cell });
            else out.push({ ...slot, col: a - start, colSpan: b - a });
        }
        return out;
    }

    /** One `longtable`/`tabular` holding a band of at most {@link MAX_TABLE_BAND_COLUMNS} columns. */
    private async tableBand(rows: OfficeContentNode[], rowOverrides: (string | false | void)[], grid: TableSlot[][], cols: number, firstBand: boolean): Promise<string> {
        const longtable = this.ctx.longtable;
        if (longtable) this.uses.longtable = true;
        // Where footnotes are written in place (the body, a longtable's cell), a tabular's notes are
        // marks whose texts it writes after itself; in a tabular within a tabular, the outer one
        // writes them; elsewhere (a note's own text, a page header) a table's notes are what that
        // place makes them. A longtable's cells write theirs in place.
        const inPlace = this.ctx.notes === 'direct' || this.ctx.notes === 'cell';
        const ownsDeferral = !longtable && inPlace;
        const cellNotes: RenderContext['notes'] = longtable
            ? (this.ctx.notes === 'direct' ? 'cell' : this.ctx.notes)
            : (inPlace || this.ctx.notes === 'deferred' ? 'deferred' : this.ctx.notes);

        const env = longtable ? 'longtable' : 'tabular';
        const compact = TABLE_COMPACT_STEPS.find(step => cols >= step.minCols);
        let out = compact ? `{${compact.size}\\setlength{\\tabcolsep}{${compact.tabcolsep}}\n` : '';
        const aligns = this.columnAlignments(grid, cols);
        const spec = aligns ? aligns.map(a => this.columnSpec(cols, 1, false, a)).join('') : `*{${cols}}{${this.columnSpec(cols, 1, false)}}`;
        out += `\\begin{${env}}{|${spec}}\n\\hline\n`;

        let headerDone = !longtable;
        for (let ri = 0; ri < rows.length; ri++) {
            checkAbortSignal(this.config.abortSignal);
            const row = rows[ri];
            const override = rowOverrides[ri];
            if (override === false) continue;
            let line: string;
            if (typeof override === 'string') {
                line = `\\multicolumn{${cols}}{${this.columnSpec(cols, cols, true)}}{${firstBand ? override : ''}}`;
            } else {
                const parts: string[] = [];
                for (const slot of grid[ri]) parts.push(await this.renderSlot(slot, cols, cellNotes, aligns?.[slot.col]));
                line = parts.join(' & ');
            }
            out += `${line} \\\\\n${this.rowRule(grid, ri)}\n`;
            if (!headerDone) {
                const nextIsHeader = ri + 1 < rows.length && isHeaderRow(rows[ri + 1], false);
                const spansCross = grid[ri + 1]?.some(s => s.kind === 'covered');
                if (!isHeaderRow(row, ri === 0)) headerDone = true;
                else if (!nextIsHeader && !spansCross) { out += '\\endhead\n'; headerDone = true; }
            }
        }
        out += `\\end{${env}}`;
        if (compact) out += '\n}';
        if (ownsDeferral) out += this.flushDeferredFootnotes();
        return out;
    }

    /**
     * The rule under row `ri`: a full `\hline`, or `\cline`s that skip the columns a row span carries
     * on into the next row (a rule through them would cut the merged cell).
     */
    private rowRule(grid: TableSlot[][], ri: number): string {
        const next = grid[ri + 1];
        if (!next) return '\\hline';
        const covered = next.filter((s): s is Extract<TableSlot, { kind: 'covered' }> => s.kind === 'covered');
        if (covered.length === 0) return '\\hline';
        const blocked = new Set<number>();
        for (const s of covered) for (let k = 0; k < s.colSpan; k++) blocked.add(s.col + k);
        const width = next.reduce((w, s) => Math.max(w, s.col + (s.kind === 'empty' ? 1 : s.colSpan)), 0);
        const segments: string[] = [];
        let start = -1;
        for (let c = 0; c <= width; c++) {
            const open = c < width && !blocked.has(c);
            if (open && start < 0) start = c;
            if (!open && start >= 0) { segments.push(`\\cline{${start + 1}-${c}}`); start = -1; }
        }
        return segments.join('');
    }

    /**
     * One position of a row. A merged cell is set with its own alignment, and so is a cell whose
     * alignment differs from its column's (`columnAlign`), in a `\multicolumn{1}` of its own.
     */
    private async renderSlot(slot: TableSlot, cols: number, notes: RenderContext['notes'], columnAlign: string | undefined): Promise<string> {
        if (slot.kind === 'empty') return '';
        const firstCol = slot.col === 0;
        const alignOf = (cell: OfficeContentNode) => (this.config.includeFormatting === false ? 'left' : this.cellAlign(cell));
        if (slot.kind === 'covered' || slot.kind === 'spill') {
            const bg = this.cellBackground(slot.origin);
            if (slot.colSpan > 1) return `\\multicolumn{${slot.colSpan}}{${this.columnSpec(cols, slot.colSpan, firstCol, alignOf(slot.origin))}}{${bg}}`;
            return bg;
        }
        const override = await this.handleOnNode(slot.cell);
        if (override === false) return '';
        let content = typeof override === 'string'
            ? override
            : (await this.withCtx({ sections: false, verbatim: false, longtable: false, notes }, () => this.renderFlow(slot.cell.children))).trim();
        if (typeof override !== 'string') {
            content = this.commentsBefore(slot.cell) + this.anchorsFor(slot.cell, true) + content;
        }
        if (slot.rowSpan > 1) {
            this.uses.multirow = true;
            content = `\\multirow{${slot.rowSpan}}{=}{${content}}`;
        }
        content = this.cellBackground(slot.cell) + content;
        const own = alignOf(slot.cell);
        if (slot.colSpan > 1 || own !== (columnAlign ?? 'left')) return `\\multicolumn{${slot.colSpan}}{${this.columnSpec(cols, slot.colSpan, firstCol, own)}}{${content}}`;
        return content;
    }

    private cellBackground(cell: OfficeContentNode): string {
        if (this.config.includeFormatting === false) return '';
        const bg = hexColor((cell.metadata as any)?.backgroundColor);
        if (!bg) return '';
        this.uses.colortbl = true;
        return `\\cellcolor{${this.colorRef(bg)}}`;
    }

    private async sheet(node: OfficeContentNode): Promise<string> {
        const name = (node.metadata as any)?.sheetName;
        const title = name ? (this.ctx.sections ? `\\${this.sectionCommand(1)}{${escapeLatex(String(name), ' ')}}` : `\\textbf{${escapeLatex(String(name), ' ')}}`) : '';
        const children = node.children || [];
        const table = await this.table(node, fillSheetRowGaps(children.filter(c => c.type === 'row'), n => this.takeGridPositions(n)));
        // A sheet's drawings and charts follow its rows as non-row children; render them after the grid.
        const extras = await this.renderFlow(children.filter(c => c.type !== 'row'));
        return [this.anchorsFor(node, true) + title, table, extras].filter(Boolean).join(BLOCK_SEPARATOR);
    }

    private async chart(node: OfficeContentNode): Promise<string> {
        if (this.config.includeCharts === false) return '';
        const meta = node.metadata as any;
        const data = this.getAttachment(meta?.attachmentName)?.chartData;
        if (!data) return escapeLatex(`[Chart: ${meta?.attachmentName || ''}]`);
        const cell = (text: string): OfficeContentNode => ({ type: 'cell', children: [{ type: 'text', text } as OfficeContentNode] } as OfficeContentNode);
        const caption = data.title ? `\\textbf{${escapeLatex(data.title, ' ')}}` : '';
        return [caption, await this.table(this.chartTable(data, cell))].filter(Boolean).join(BLOCK_SEPARATOR);
    }

    // ── images ───────────────────────────────────────────────────────────────────────────────────

    /** An attachment's base name reduced to `[A-Za-z0-9-]`, for the file names the output writes. */
    private fileStem(attachmentName: string): string {
        const base = String(attachmentName).split(/[\\/]/).pop()!.replace(/\.[^.]*$/, '');
        return base.replace(/[^A-Za-z0-9-]+/g, '-').replace(/-+/g, '-').replace(/^-|-$/g, '') || 'image';
    }

    /** A safe, unique file name for an attachment: its base name reduced to `[A-Za-z0-9-]`, with the MIME's extension. */
    private uniqueFileName(attachmentName: string, ext: string): string {
        const stem = this.fileStem(attachmentName);
        return this.usedFileNames.claim(`${stem}.${ext}`, n => `${stem}-${n}.${ext}`);
    }

    /**
     * A PNG or JPEG carried inside the .tex: a `filecontents*` block writes it, as a small PDF, beside
     * the .tex when the document compiles, and `\includegraphics` reads that file. The name holds a
     * hash of the data, so a changed image never meets a stale file, and without `overwrite` LaTeX
     * keeps a file of that name already there, which lets a reader replace the picture. Null for a
     * bundle (which holds the image files themselves), with `embedImages` off, or for an image that
     * cannot be carried (another format, damaged, or past the document's decoding budget).
     */
    private carriedImage(bytes: Uint8Array, mime: string, attachmentName: string): MediaRef | null {
        const { bundle, embedImages } = this.config.texConfig;
        if (bundle === true || embedImages === false || !['png', 'jpg'].includes(INCLUDABLE_IMAGE_EXT[mime])) return null;
        const made = imageToTextPdf(bytes, PT_PER_INCH / IMAGE_DPI, this.decodeBudget);
        if (!made) return null;
        // Hashing the carried PDF, not the image file, gives an image the same name after a round
        // trip through the parser (which drops PNG chunks the picture does not need), and a name
        // already ending in that hash does not gain it twice. The hash is short, so a match is
        // confirmed on the text itself: two pictures sharing a hash are both carried, apart.
        const hash = contentHash(new TextEncoder().encode(made.pdf));
        const sameHash = this.carriedByHash.get(hash) ?? [];
        const known = sameHash.find(entry => entry.pdf === made.pdf);
        if (known) return known.ref;
        const stem = `${this.fileStem(attachmentName).replace(new RegExp(`-${hash}$`), '')}-${hash}`;
        const name = this.usedFileNames.claim(`${stem}.pdf`, n => `${stem}-${n}.pdf`);
        const ref: MediaRef = { path: name, includable: true, mime, intrinsic: sniffImageSize(bytes), bb: made.bbox };
        this.carried.push({ name, pdf: made.pdf });
        this.carriedByHash.set(hash, [...sameHash, { pdf: made.pdf, ref }]);
        return ref;
    }

    /** The `filecontents*` blocks writing the images carried inside the .tex, headed by what they do. */
    private carriedImageBlocks(): string {
        if (!this.carried.length) return '';
        return latexComment('The images, carried in this file: each block writes one, as a small PDF, beside the .tex when it\ncompiles. A file of that name already there is kept, so an image can be replaced.')
            + this.carried.map(c => `\\begin{filecontents*}{${c.name}}\n${c.pdf}\n\\end{filecontents*}`).join('\n');
    }

    private mediaFor(attachmentName: string): MediaRef | null {
        if (this.mediaByAttachment.has(attachmentName)) return this.mediaByAttachment.get(attachmentName)!;
        let ref: MediaRef | null = null;
        const att = this.getAttachment(attachmentName);
        const mime = (att?.mimeType || '').toLowerCase();
        const ext = INCLUDABLE_IMAGE_EXT[mime] ?? MIME_EXT[mime];
        if (!att || !att.data) {
            this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: attachmentName, reason: 'missing attachment' });
        } else if (!ext) {
            this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: attachmentName, reason: 'unsupported mime' });
        } else {
            try {
                const bytes = decodeBase64(att.data);
                ref = this.carriedImage(bytes, mime, attachmentName);
                if (!ref) {
                    const path = `${IMAGE_DIR}/${this.uniqueFileName(attachmentName, ext)}`;
                    this.media.push({ path, bytes, includable: !!INCLUDABLE_IMAGE_EXT[mime] });
                    ref = { path, includable: !!INCLUDABLE_IMAGE_EXT[mime], mime, intrinsic: sniffImageSize(bytes) };
                }
            } catch {
                this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: attachmentName });
            }
        }
        this.mediaByAttachment.set(attachmentName, ref);
        return ref;
    }

    /**
     * An image the source embedded as a `data:` URI (a Markdown or HTML image parsed without
     * `extractAttachments`): decoded and written under `images/` like an attachment, since the URI
     * carries the image itself. Null when the URI cannot be decoded or holds no image type LaTeX or
     * the bundle can store.
     */
    private mediaForDataUri(uri: string): MediaRef | null {
        if (this.mediaByDataUri.has(uri)) return this.mediaByDataUri.get(uri)!;
        let ref: MediaRef | null = null;
        const m = /^data:([^;,]*)((?:;[^;,]*)*?)(;base64)?,(.*)$/is.exec(uri.trim());
        const mime = (m?.[1] || '').trim().toLowerCase();
        const ext = INCLUDABLE_IMAGE_EXT[mime] ?? MIME_EXT[mime];
        if (!m || !ext) {
            this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: 'data: URI image', reason: m ? 'unsupported mime' : 'malformed data: URI' });
        } else {
            try {
                const bytes = m[3] ? decodeBase64(m[4]) : new TextEncoder().encode(decodeURIComponent(m[4]));
                ref = this.carriedImage(bytes, mime, 'image');
                if (!ref) {
                    const path = `${IMAGE_DIR}/${this.uniqueFileName('image', ext)}`;
                    this.media.push({ path, bytes, includable: !!INCLUDABLE_IMAGE_EXT[mime] });
                    ref = { path, includable: !!INCLUDABLE_IMAGE_EXT[mime], mime, intrinsic: sniffImageSize(bytes) };
                }
            } catch {
                this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: 'data: URI image', reason: 'undecodable data' });
            }
        }
        this.mediaByDataUri.set(uri, ref);
        return ref;
    }

    /**
     * `\includegraphics` options drawing an image at its natural size but never wider than the line
     * or taller than the context allows. The bounds are TeX conditionals rather than a computed
     * number because the line width depends on where the image lands (a cell, a list, a frame).
     */
    private imageSize(node: OfficeContentNode, intrinsic: { w: number; h: number } | null, natural = false): string {
        const maxH = this.ctx.imageMaxHeight;
        const meta = node.metadata as ImageMetadata | undefined;
        const pct = meta?.width ? /^\s*([\d.]+)\s*%\s*$/.exec(meta.width) : null;
        if (pct) {
            const frac = Math.max(0.01, Math.min(1, parseFloat(pct[1]) / 100));
            return `width=${Math.round(frac * 1000) / 1000}\\linewidth,height=${maxH},keepaspectratio`;
        }
        let w: number | null = null;
        let h: number | null = null;
        const explicit = meta?.width ? lengthToPt(meta.width) : null;
        if (explicit && explicit > 0) {
            w = explicit;
            h = intrinsic ? explicit * intrinsic.h / intrinsic.w : null;
        } else if (node.bounds && node.bounds.width > 0 && node.bounds.height > 0) {
            w = node.bounds.width;
            h = node.bounds.height;
        } else if (intrinsic) {
            w = intrinsic.w / IMAGE_DPI * PT_PER_INCH;
            h = intrinsic.h / IMAGE_DPI * PT_PER_INCH;
        }
        // An image of unknown size drawn from a path keeps its natural size, as its source did.
        if (!w) return natural ? '' : `width=\\linewidth,height=${maxH},keepaspectratio`;
        const width = `{\\ifdim ${fmtPt(w)}>\\linewidth\\linewidth\\else ${fmtPt(w)}\\fi}`;
        const height = h ? `{\\ifdim ${fmtPt(h)}>${maxH}${maxH}\\else ${fmtPt(h)}\\fi}` : maxH;
        return `width=${width},height=${height},keepaspectratio`;
    }

    /** An image's recognized (OCR) text: a line of text, or a layout-preserving block when multi-line. */
    private ocrMarkup(ocr: string, block: boolean): string {
        if (!ocr) return '';
        if (/\n/.test(ocr)) return block ? this.codeText(ocr, undefined, '') : this.inlineCodeLines(ocr);
        return escapeLatex(ocr, ' ');
    }

    private imageMarkup(node: OfficeContentNode, block: boolean): string {
        const mode = this.imageMode();
        if (mode === 'none') return '';
        const meta = node.metadata as ImageMetadata | undefined;
        const ocr = (node.text || '').trim();
        if (mode === 'ocr-text-only') return this.ocrMarkup(ocr, block);

        const alt = meta?.altText || '';
        // graphicx's `alt` key holds the picture's alternative text wherever it stands (a tagged PDF reads it).
        const altKey = alt ? `alt={${escapeLatex(alt, ' ')}}` : '';
        const options = (...keys: string[]) => keys.filter(Boolean).join(',');
        let img = '';
        const dataUri = !meta?.attachmentName && meta?.url && /^\s*data:/i.test(meta.url) ? meta.url : null;
        if (meta?.attachmentName || dataUri) {
            const ref = dataUri ? this.mediaForDataUri(dataUri) : this.mediaFor(meta!.attachmentName!);
            if (ref?.includable) {
                this.uses.graphics = true;
                // A carried image states its size (bb), so the DVI engines need not run extractbb for it.
                img = `\\includegraphics[${options(ref.bb ? `bb=${ref.bb}` : '', this.imageSize(node, ref.intrinsic), altKey)}]{${ref.path}}`;
            } else if (ref) {
                this.warnOnce(`image:${ref.mime}`, OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: `${ref.mime} image`, format: 'tex' });
                img = `\\fbox{${escapeLatex(`Image: ${alt || ref.path}`, ' ')}}`;
            }
        } else if (meta?.url) {
            const path = sanitizeLatexImagePath(meta.url);
            if (path) {
                // A relative path (from a LaTeX \includegraphics, or an HTML/Markdown image beside its
                // page) stays a reference to that file, as in the source; the warning names it.
                this.uses.graphics = true;
                this.externalImages.add(path);
                const keys = options(this.imageSize(node, null, true), altKey);
                const include = `\\includegraphics${keys ? `[${keys}]` : ''}{${path}}`;
                // A bundle promises a zip that compiles as is, but it cannot contain a file the source
                // only named. There the image is drawn when the file has been added beside the .tex,
                // and a labelled box stands in otherwise.
                img = this.config.texConfig.bundle === true
                    ? this.ifImageExists(path, include, `\\fbox{${escapeLatex(`Image: ${alt || path}`, ' ')}}`)
                    : include;
            } else {
                // LaTeX cannot fetch a web image, and reads no file outside the document's folder.
                const remote = /^(?:https?|ftp):|^\/\//i.test(meta.url.trim());
                this.warnOnce(remote ? 'image:remote' : 'image:path', OfficeWarningType.CONTENT_NOT_REPRESENTABLE,
                    { feature: remote ? 'remote image' : 'image path that is absolute, leaves its folder or uses characters other than letters, digits, spaces and . _ - /', format: 'tex' });
                const url = sanitizeLatexUrl(meta.url);
                const text = escapeLatex(alt || meta.url, ' ');
                // Linked to the image, unless the picture is a link itself: then to that link, below.
                img = url && !meta.link ? `${this.cmd('href')}{${url}}{${text}}` : text;
            }
        }
        if (!img) return alt ? escapeLatex(alt, ' ') : this.ocrMarkup(ocr, block);

        // A picture that is a link, as a linked run is: `\hyperref` to an internal target's label
        // (none under ignoreInternalLinks), `\href` to an external one.
        const link = meta?.link;
        if (link) {
            if (meta.linkType === 'internal' || link.startsWith('#')) {
                const label = this.config.ignoreInternalLinks ? null : this.resolveLabel(link);
                if (label) img = `${this.cmd('hyperref')}[${label}]{${img}}`;
            } else {
                const url = sanitizeLatexUrl(link);
                if (url) img = `${this.cmd('href')}{${url}}{${img}}`;
            }
        }
        if (mode === 'image+ocr-text' && ocr) return `${img}${block ? BLOCK_SEPARATOR : ' '}${this.ocrMarkup(ocr, block)}`;
        return img;
    }

    /**
     * `\IfFileExists` over the names `\includegraphics` would try: the path itself, then, for a path
     * with no extension, the path plus each graphics extension the engines look for.
     */
    private ifImageExists(path: string, include: string, fallback: string): string {
        const names = /\.[A-Za-z0-9]+$/.test(path.split('/').pop() || '') ? [path] : [path, ...IMAGE_PROBE_EXTENSIONS.map(ext => path + ext)];
        return names.reduceRight((otherwise, name) => `\\IfFileExists{${name}}{${include}}{${otherwise}}`, fallback);
    }

    private blockImage(node: OfficeContentNode): string {
        const meta = node.metadata as ImageMetadata | undefined;
        const img = this.imageMarkup(node, true);
        if (!img) return '';
        const anchors = this.anchorsFor(node, true);
        const align = this.config.includeFormatting === false ? undefined : meta?.align;
        const body = align === 'center' ? `{\\centering ${img}\\par}` : align === 'right' ? `{\\raggedleft ${img}\\par}` : img;
        return `${this.commentsBefore(node)}${anchors}${body}`;
    }

    // ── code, breaks, admonitions, embeds ────────────────────────────────────────────────────────

    private codeBlock(node: OfficeContentNode): string {
        const meta = node.metadata as CodeMetadata | undefined;
        const prefix = this.commentsBefore(node) + this.anchorsFor(node, true);
        if (meta?.math === 'block' || meta?.math === 'inline') {
            const math = this.ctx.display && meta.math === 'block' ? this.displayMath(node.text || '') : this.inlineMath(node.text || '');
            return math ? prefix + math : prefix.trimEnd();
        }
        return this.codeText(node.text || this.getNodeText(node), meta?.language, prefix);
    }

    /**
     * A code block. Where environments that read raw input are allowed it is `lstlisting` (for a
     * language `listings` knows, ASCII-only, since `listings` mangles multi-byte UTF-8 under pdfTeX)
     * or `verbatim`; neither is used if the code contains its own end marker, which would close the
     * environment early and turn the rest of the code into live LaTeX. Nor in a `beamer` document if
     * it contains `\end{frame}`: a frame holding verbatim material is `fragile`, which beamer reads
     * raw up to a line `\end{frame}`, so that line in the code would end the frame and run the lines
     * after it as LaTeX (reading files, or running commands under shell escape). Elsewhere (a cell, a
     * note) it is typewriter text with every character escaped.
     */
    private codeText(text: string, language: string | undefined, prefix: string): string {
        const lines = trimEndChars(String(text ?? '').replace(/\r\n?/g, '\n'), '\n').split('\n').map(expandTabs);
        const code = lines.join('\n');
        if (!this.ctx.verbatim || (this.beamer && /\\end\s*\{\s*frame\s*\}/.test(code))) return prefix + this.inlineCodeLines(code);
        const listingsLang = language ? LISTINGS_LANGUAGES[language.trim().toLowerCase()] : undefined;
        if (listingsLang && /^[\x20-\x7E\n]*$/.test(code) && !/\\end\s*\{\s*lstlisting\s*\}/.test(code)) {
            this.uses.listings = true;
            return `${prefix}\\begin{lstlisting}[language={${listingsLang}}]\n${code}\n\\end{lstlisting}`;
        }
        if (!/\\end\s*\{\s*verbatim\s*\}/.test(code) && !/[\x00-\x08\x0B-\x1F\x7F]/.test(code)) {
            return `${prefix}\\begin{verbatim}\n${code}\n\\end{verbatim}`;
        }
        return prefix + this.inlineCodeLines(code);
    }

    /**
     * Typewriter lines with every character escaped, for places a verbatim environment cannot go.
     * Indentation and runs of spaces are kept with `\phantom` boxes (plain spaces would collapse, and
     * a space at the start of a line would be dropped).
     */
    private inlineCodeLines(text: string): string {
        const lines = String(text ?? '').replace(/\r\n?/g, '\n').split('\n').map(expandTabs);
        const rendered = lines.map(line => {
            let out = '';
            let i = 0;
            while (i < line.length) {
                if (line[i] === ' ') {
                    let j = i;
                    while (j < line.length && line[j] === ' ') j++;
                    const run = j - i;
                    out += (i === 0 || run > 1) ? `\\phantom{${'x'.repeat(run)}}` : ' ';
                    i = j;
                    continue;
                }
                let j = i;
                while (j < line.length && line[j] !== ' ') j++;
                out += escapeLatex(line.slice(i, j), ' ');
                i = j;
            }
            return out;
        });
        const sep = this.ctx.moving ? ' ' : LINE_BREAK;
        return `{\\ttfamily ${rendered.join(sep)}}`;
    }

    private blockBreak(node: OfficeContentNode): string {
        const t = (node.metadata as any)?.breakType;
        if (t === 'page') return this.beamer ? '' : '\\clearpage';
        if (t === 'thematic') return '\\begin{center}\\rule{0.5\\linewidth}{0.5pt}\\end{center}';
        return '';
    }

    private async admonition(node: OfficeContentNode): Promise<string> {
        const meta = node.metadata as AdmonitionMetadata | undefined;
        const type = String(meta?.admonitionType || 'note');
        const color = ADMONITION_COLOR[type] || ADMONITION_COLOR.note;
        const title = meta?.title || (type.charAt(0).toUpperCase() + type.slice(1));
        const colorName = this.colorRef(color);
        const body = (await this.withCtx({ sections: false, longtable: false }, () => this.renderFlow(node.children))).trim();
        return `${this.commentsBefore(node)}\\begin{quote}\n\\textbf{\\textcolor{${colorName}}{${escapeLatex(title, ' ')}}}\\par\n${body}\n\\end{quote}`;
    }

    private embed(node: OfficeContentNode): string {
        const meta = node.metadata as any;
        const rawUrl = embedUrl(meta);
        const url = rawUrl ? sanitizeLatexUrl(String(rawUrl)) : '';
        const label = escapeLatex(String(meta?.label || node.text || rawUrl || ''), ' ');
        if (!url) {
            this.warnOnce('embed', OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: 'embed', format: 'tex' });
            return label;
        }
        return `${this.cmd('href')}{${url}}{${label || url}}`;
    }

    // ── slides ───────────────────────────────────────────────────────────────────────────────────

    /**
     * The body of a `beamer` document. Slides (and pages and sheets) become one frame each. Any other
     * top-level content is grouped into frames, starting a new frame at each level-1/2 heading, which
     * becomes the frame title.
     */
    private async renderBeamerBody(nodes: OfficeContentNode[]): Promise<string> {
        const frames: string[] = [];
        let pending: { title: OfficeContentNode | null; body: OfficeContentNode[] } | null = null;
        const isFramed = (n: OfficeContentNode | undefined) => n?.type === 'slide' || n?.type === 'page' || n?.type === 'sheet';
        const isTitleHeading = (n: OfficeContentNode | undefined) => n?.type === 'heading' && ((n.metadata as HeadingMetadata)?.level ?? 1) <= 2;
        const flushPending = async (at: number) => {
            const p = pending as { title: OfficeContentNode | null; body: OfficeContentNode[] } | null;
            pending = null;
            if (!p || (!p.title && !p.body.length)) return;
            // A heading with nothing under it before the next slide (past any headings like it) is a
            // `\section` between frames, as beamer's own documents have (what the parser reads them as),
            // not a frame of its own.
            let next = at;
            while (isTitleHeading(nodes[next])) next++;
            if (p.title && !p.body.length && isFramed(nodes[next])) {
                const section = await this.withCtx({ sections: true }, () => this.heading(p.title!));
                if (section) frames.push(section);
                return;
            }
            frames.push(await this.frame(p.title, p.body, [], null));
        };
        for (let i = 0; i < nodes.length; i++) {
            const node = nodes[i];
            checkAbortSignal(this.config.abortSignal);
            if (this.titleNodes?.nodes.has(node)) {
                if (node !== this.titleNodes.title) continue;
                await flushPending(i);
                const override = await this.handleOnNode(node);
                if (override === false) continue;
                frames.push(typeof override === 'string' ? override : await this.frame(null, [], [], node, '\\titlepage'));
                continue;
            }
            const framed = isFramed(node);
            const titleHeading = isTitleHeading(node);
            if (!framed && !titleHeading) {
                (pending ??= { title: null, body: [] }).body.push(node);
                continue;
            }
            await flushPending(i);
            const override = await this.handleOnNode(node);
            if (override === false) continue;
            if (typeof override === 'string') { frames.push(override); continue; }
            if (titleHeading) { pending = { title: node, body: [] }; continue; }
            if (node.type === 'sheet') { frames.push(await this.frame(null, [], [], node)); continue; }
            if (this.titleNodes && this.titleNodes.container === node) {
                // The title slide: \titlepage typesets the block; anything else on it follows.
                const rest = (node.children || []).filter(c => !this.titleNodes!.nodes.has(c));
                frames.push(await this.frame(null, rest, node.notes || [], node, '\\titlepage'));
                continue;
            }
            const children = node.children || [];
            const titleIdx = children.findIndex(c => c.type === 'heading');
            const title = titleIdx >= 0 ? children[titleIdx] : null;
            const body = titleIdx >= 0 ? [...children.slice(0, titleIdx), ...children.slice(titleIdx + 1)] : children;
            frames.push(await this.frame(title, body, node.notes || [], node));
        }
        await flushPending(nodes.length);
        return frames.join(BLOCK_SEPARATOR);
    }

    /** A frame title's (or subtitle's) text, in a moving argument. */
    private async frameHeading(heading: OfficeContentNode): Promise<string> {
        const override = await this.handleOnNode(heading);
        if (override === false) return '';
        if (typeof override === 'string') return override.trim();
        return (await this.withCtx({ moving: true, display: false, verbatim: false, labels: false },
            async () => (await this.headingRuns(heading, this.hasUniformFormatting(heading, f => f?.bold === true))) + await this.notesFor(heading))).trim();
    }

    private async frame(title: OfficeContentNode | null, body: OfficeContentNode[], notes: OfficeContentNode[], owner: OfficeContentNode | null, lead = ''): Promise<string> {
        const anchors = owner ? this.anchorsFor(owner, true) : '';
        const titleTex = title ? await this.frameHeading(title) : '';
        // A heading of a lower level right under the title is the frame's subtitle (what the parser
        // reads `\framesubtitle` and `\begin{frame}{Title}{Subtitle}` as).
        const level = (n: OfficeContentNode | undefined) => Math.floor(Number((n?.metadata as HeadingMetadata | undefined)?.level)) || 1;
        const first = body[0];
        const subtitle = title && first?.type === 'heading' && level(first) > level(title) ? first : null;
        if (subtitle) body = body.slice(1);
        const subtitleTex = subtitle ? await this.frameHeading(subtitle) : '';
        let content: string;
        if (owner?.type === 'sheet') {
            content = await this.withCtx({ sections: false, longtable: false }, () => this.sheetInFrame(owner));
        } else {
            content = await this.withCtx({ sections: false, longtable: false }, () => this.renderFlow(body));
        }
        let noteTex = '';
        for (const note of notes) {
            const text = (await this.withCtx({ ...this.argumentCtx(), notes: 'omit' }, () => this.renderFlow(this.bodyBlocks(note)))).trim();
            if (text) noteTex += `\n\\note{${text}}`;
        }
        // allowframebreaks continues a slide whose content is taller than the frame on a follow-up
        // frame instead of letting it run off the bottom; office slides often hold more text than a
        // beamer frame fits. fragile is required around verbatim material.
        const options = ['allowframebreaks', ...(/\\begin\{(verbatim|lstlisting)\}/.test(content) ? ['fragile'] : [])];
        // beamer reads a `{` right after \begin{frame} as the frame title, so a body that starts with a
        // group (an aligned paragraph, a sized run) needs something else in front of it.
        const lines = [`\\begin{frame}[${options.join(',')}]`];
        lines.push(titleTex ? `\\frametitle{${titleTex}}` : '\\relax');
        if (subtitleTex) lines.push(`\\framesubtitle{${subtitleTex}}`);
        for (const heading of [title, subtitle]) {
            const labels = heading ? this.anchorsFor(heading, true) : '';
            if (labels) lines.push(labels);
        }
        if (anchors) lines.push(anchors);
        if (lead) lines.push(lead);
        if (content.trim()) lines.push(content.trim());
        lines.push(`\\end{frame}${noteTex}`);
        return lines.join('\n');
    }

    private async sheetInFrame(node: OfficeContentNode): Promise<string> {
        const name = (node.metadata as any)?.sheetName;
        const children = node.children || [];
        const table = await this.table(node, fillSheetRowGaps(children.filter(c => c.type === 'row'), n => this.takeGridPositions(n)));
        const extras = await this.renderFlow(children.filter(c => c.type !== 'row'));
        return [name ? `\\textbf{${escapeLatex(String(name), ' ')}}` : '', table, extras].filter(Boolean).join(BLOCK_SEPARATOR);
    }

    /** A slide in a non-`beamer` document: its content, then its speaker notes as a labelled quote. */
    private async slideInArticle(node: OfficeContentNode): Promise<string> {
        const content = await this.renderFlow(node.children);
        let notes = '';
        for (const note of node.notes || []) {
            const text = (await this.withCtx({ sections: false, longtable: false, notes: 'omit' }, () => this.renderFlow(this.bodyBlocks(note)))).trim();
            if (text) notes += `${notes ? BLOCK_SEPARATOR : ''}${text}`;
        }
        const parts = [this.anchorsFor(node, true) + content];
        if (notes) parts.push(`\\begin{quote}\n\\textbf{Speaker notes}\\par\n${notes}\n\\end{quote}`);
        return parts.filter(p => p.trim()).join(BLOCK_SEPARATOR);
    }

    // ── page headers & footers ───────────────────────────────────────────────────────────────────

    private pickHeaderFooter(nodes: OfficeContentNode[] | undefined): OfficeContentNode | undefined {
        if (!nodes || nodes.length === 0) return undefined;
        return nodes.find(n => (n.metadata as any)?.type === 'default') ?? nodes[0];
    }

    /** One header/footer's content as the lines of a `fancyhdr` field, plus its horizontal position. */
    private async headerFieldContent(node: OfficeContentNode): Promise<{ text: string; lines: number; position: 'L' | 'C' | 'R' }> {
        const lines: string[] = [];
        let position: 'L' | 'C' | 'R' | null = null;
        const collect = async (n: OfficeContentNode): Promise<void> => {
            if (n.type === 'paragraph' || n.type === 'heading' || n.type === 'list') {
                if (!position) {
                    const a = (n.metadata as any)?.alignment;
                    position = a === 'center' ? 'C' : a === 'right' ? 'R' : 'L';
                }
                const line = (await this.renderInline(this.inlineChildren(n))).trim();
                if (line) lines.push(line);
                return;
            }
            if (n.type === 'table') {
                for (const row of n.children || []) {
                    const cells: string[] = [];
                    for (const cell of row.children || []) {
                        const parts: string[] = [];
                        for (const c of cell.children || []) {
                            const t = (await this.renderInline(this.inlineChildren(c))).trim();
                            if (t) parts.push(t);
                        }
                        if (parts.length) cells.push(parts.join(' '));
                    }
                    if (cells.length) lines.push(cells.join('\\quad '));
                }
                return;
            }
            if (n.type === 'image') { const img = this.imageMarkup(n, false); if (img) lines.push(img); return; }
            if (n.type === 'text') { const t = (await this.renderInline([n])).trim(); if (t) lines.push(t); return; }
            for (const c of n.children || []) await collect(c);
        };
        await this.withCtx({ sections: false, verbatim: false, longtable: false, display: false, notes: 'omit', labels: false, imageMaxHeight: IMAGE_MAX_HEIGHT_HEADER },
            async () => { for (const c of node.children || []) await collect(c); });
        return { text: lines.join(' \\\\ '), lines: Math.max(1, lines.length), position: position ?? 'L' };
    }

    /** `fancyhdr` setup reproducing the document's running header and footer, or '' when it has none. */
    private async pageHeaderSetup(): Promise<string> {
        const header = this.pickHeaderFooter(this.ast.auxiliary?.headers);
        const footer = this.pickHeaderFooter(this.ast.auxiliary?.footers);
        if (!header && !footer) return '';
        const head = header ? await this.headerFieldContent(header) : null;
        const foot = footer ? await this.headerFieldContent(footer) : null;
        if (!head?.text && !foot?.text) return '';
        const lines = [
            '\\pagestyle{fancy}',
            '\\fancyhf{}',
            '\\renewcommand{\\headrulewidth}{0pt}',
        ];
        if (head?.text) {
            lines.push(`\\fancyhead[${head.position}]{${head.text}}`);
            lines.push(`\\setlength{\\headheight}{${fmtPt(head.lines * HEADER_LINE_HEIGHT_PT + (/\\includegraphics/.test(head.text) ? HEADER_IMAGE_HEIGHT_PT : 0))}}`);
        }
        lines.push(foot?.text ? `\\fancyfoot[${foot.position}]{${foot.text}}` : '\\fancyfoot[C]{\\thepage}');
        // \maketitle and \chapter switch their page to the `plain` style, which would drop the
        // document's running header there; make `plain` the same as the header/footer set up here.
        lines.push('\\fancypagestyle{plain}[fancy]{}');
        return lines.join('\n');
    }

    private warnBeamerHeaders(): string {
        if ((this.ast.auxiliary?.headers?.length || 0) + (this.ast.auxiliary?.footers?.length || 0) > 0) {
            this.warnOnce('header', OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: 'page header/footer', format: 'tex (beamer)' });
        }
        return '';
    }

    // ── preamble & metadata ──────────────────────────────────────────────────────────────────────

    /** `\title`/`\author`/`\date` plus the hyperref PDF metadata, from the effective metadata. */
    private metadataCommands(): { title: string; hypersetup: string[] } {
        const m = this.effectiveMetadata;
        const esc = (v: unknown) => escapeLatex(String(v ?? ''), ' ').trim();
        const out: string[] = [];
        const f = this.titleFields;
        if (f) {
            // The title block in the content decides what \maketitle prints, including an empty
            // author or date where the block has none (LaTeX would otherwise print today's date).
            out.push(`\\title{${f.title}}`);
            if (f.subtitle) out.push(`\\subtitle{${f.subtitle}}`);
            out.push(`\\author{${f.author}}`);
            if (f.institute) out.push(`\\institute{${f.institute}}`);
            out.push(`\\date{${f.date}}`);
        } else {
            if (m.title) out.push(`\\title{${esc(m.title)}}`);
            if (m.author) out.push(`\\author{${esc(m.author)}}`);
            const dateIso = toW3CDTF(m.modified) ?? toW3CDTF(m.created);
            if (m.title) out.push(`\\date{${dateIso ? dateIso.slice(0, 10) : ''}}`);
        }

        const pdfDate = (v: unknown) => { const iso = toW3CDTF(v); return iso ? `D:${iso.replace(/[-:T]/g, '').replace('Z', '')}Z` : ''; };
        const keys: string[] = [];
        if (m.title) keys.push(`pdftitle={${esc(m.title)}}`);
        if (m.author) keys.push(`pdfauthor={${esc(m.author)}}`);
        if (m.subject) keys.push(`pdfsubject={${esc(m.subject)}}`);
        if (m.keywords) keys.push(`pdfkeywords={${esc(m.keywords)}}`);
        const lang = String(m.language ?? '').replace(/[^A-Za-z0-9-]/g, '');
        if (lang) keys.push(`pdflang={${lang}}`);
        const created = pdfDate(m.created);
        if (created) keys.push(`pdfcreationdate={${created}}`);
        const modified = pdfDate(m.modified);
        if (modified) keys.push(`pdfmoddate={${modified}}`);
        const info: string[] = [];
        if (m.description) info.push(`Description={${esc(m.description)}}`);
        if (m.lastModifiedBy) info.push(`LastModifiedBy={${esc(m.lastModifiedBy)}}`);
        for (const [key, value] of Object.entries(m.customProperties || {})) {
            // A PDF info key is a PDF name: keep it to letters and digits.
            const name = String(key).replace(/[^A-Za-z0-9]/g, '');
            if (!name) continue;
            const v = value instanceof Date ? (toW3CDTF(value) ?? '') : String(value);
            info.push(`${name}={${esc(v)}}`);
        }
        if (info.length) keys.push(`pdfinfo={${info.join(',')}}`);
        return { title: out.join('\n'), hypersetup: keys };
    }

    /**
     * Finds the title block: the first heading styled `Title` at the top level (or, failing that, on
     * a slide), with the `Subtitle` and `Institute` (beamer only: `article` has neither `\subtitle` nor
     * `\institute`), `Author` and `Date` lines that directly follow it, each at most once.
     */
    private findTitleBlock(): TitleBlock | null {
        const scan = (siblings: OfficeContentNode[], container: OfficeContentNode | null): TitleBlock | null => {
            const i = siblings.findIndex(n => n.type === 'heading' && (n.metadata as HeadingMetadata | undefined)?.style === 'Title');
            if (i < 0) return null;
            const block: TitleBlock = { title: siblings[i], container, nodes: new Set([siblings[i]]) };
            for (const n of siblings.slice(i + 1)) {
                const style = n.type === 'paragraph' ? (n.metadata as ParagraphMetadata | undefined)?.style : undefined;
                // An institute line is beamer's `\\institute`; `article` has none, so there it stays a line of its own after the block.
                if (style === 'Institute' && !this.beamer) continue;
                const key = style === 'Subtitle' && this.beamer ? 'subtitle' : style === 'Institute' ? 'institute' : style === 'Author' ? 'author' : style === 'Date' ? 'date' : null;
                if (!key || block[key]) break;
                block[key] = n;
                block.nodes.add(n);
            }
            return block;
        };
        const top = scan(this.ast.content, null);
        if (top) return top;
        for (const n of this.ast.content) {
            if (n.type !== 'slide') continue;
            const onSlide = scan(n.children || [], n);
            if (onSlide) return onSlide;
        }
        return null;
    }

    /** The title block's lines as `\title`/`\subtitle`/`\author`/`\date` arguments; notes become `\thanks`. */
    private async renderTitleFields(block: TitleBlock): Promise<{ title: string; subtitle: string; author: string; institute: string; date: string }> {
        const field = async (node: OfficeContentNode | undefined): Promise<string> => {
            if (!node) return '';
            return this.withCtx({ moving: true, display: false, verbatim: false, sections: false, labels: false, longtable: false, notes: 'omit' }, async () => {
                const runs = await this.headingRuns(node, this.hasUniformFormatting(node, fmt => fmt?.bold === true));
                const notes: OfficeContentNode[] = [];
                const collect = (n: OfficeContentNode) => { if (n.notes) appendAll(notes, n.notes); n.children?.forEach(collect); };
                collect(node);
                let thanks = '';
                for (const note of notes) thanks += `\\thanks{${await this.noteBody(note, 'parenthetical')}}`;
                return runs + thanks;
            });
        };
        return { title: await field(block.title), subtitle: await field(block.subtitle), author: await field(block.author), institute: await field(block.institute), date: await field(block.date) };
    }

    private titleBlock(): string {
        // A title block in the content is typeset where it stands.
        if (this.titleNodes) return '';
        if (!this.config.renderMetadata || !this.effectiveMetadata.title) return '';
        return this.beamer ? '\\begin{frame}\n\\titlepage\n\\end{frame}' : '\\maketitle';
    }

    /**
     * What the formulas written need beyond amsmath and amssymb (see planLatexMath): packages such as
     * `bm` or `mathtools` for the commands a source loaded them for. A command nothing the output
     * loads defines prints its own name, so the document compiles and shows where it is, and is
     * reported.
     */
    private mathPlan(): LatexMathPlan {
        const plan = planLatexMath(this.mathCommands);
        for (const p of plan.packages) {
            if (p.name === 'xcolor') this.uses.xcolor = true;
            if (p.name === 'graphicx') this.uses.graphics = true;
        }
        const unknown = plan.undefinedCommands;
        if (unknown.length) {
            // Named up to a point: a document of thousands of them gets a message, not a megabyte.
            const named = unknown.slice(0, 20).map(c => `\\${c}`).join(', ') + (unknown.length > 20 ? ` and ${unknown.length - 20} more` : '');
            this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: `math commands no package the output loads defines (${named}), each printed as its name`, format: 'tex' });
        }
        return plan;
    }

    /** `\usepackage` lines for what the body uses (hyperref and the font setup are handled separately). */
    private packageLines(unicode: LatexUnicodePlan, math: LatexMathPlan): string[] {
        const u = this.uses;
        const lines: string[] = [];
        if (u.math || u.amssymb) lines.push('\\usepackage{amsmath,amssymb}');
        // The packages of the commands the formulas use; xcolor and graphicx are loaded below, as the body's.
        for (const p of math.packages) if (p.name !== 'xcolor' && p.name !== 'graphicx') lines.push(`\\usepackage${p.options ? `[${p.options}]` : ''}{${p.name}}`);
        if (unicode.packages.has('pifont')) lines.push('\\usepackage{pifont}');
        if (unicode.packages.has('newunicodechar')) lines.push('\\usepackage{newunicodechar}');
        if (u.graphics && !this.beamer) lines.push('\\usepackage{graphicx}');
        if (u.tables) lines.push(`\\usepackage{${['array', u.longtable ? 'longtable' : '', u.multirow ? 'multirow' : ''].filter(Boolean).join(',')}}`);
        if (u.ulem) lines.push('\\usepackage[normalem]{ulem}');
        if (u.xcolor && !this.beamer) lines.push(u.colortbl ? '\\usepackage[table]{xcolor}' : '\\usepackage{xcolor}');
        if (u.listings) lines.push('\\usepackage{listings}');
        if (u.endnotes) lines.push('\\usepackage{endnotes}');
        if (u.captionof) lines.push('\\usepackage{capt-of}');
        return lines;
    }

    private preamble(headerSetup: string, meta: { title: string; hypersetup: string[] }, unicode: LatexUnicodePlan, math: LatexMathPlan): string {
        const classOptions = this.beamer
            ? ['aspectratio=169', `${this.classSizePt}pt`, ...(this.uses.colortbl ? ['xcolor=table'] : [])]
            : [`${this.classSizePt}pt`];
        const lines: string[] = [
            '% Generated by officeParser.',
            '\\RequirePackage{iftex}',
            '% DVI output (latex, uplatex or platex, then dvipdfmx): a global class option gives every package the dvipdfmx driver.',
            '\\def\\officeparserdriver{}\\ifpdf\\else\\ifXeTeX\\else\\def\\officeparserdriver{dvipdfmx,}\\fi\\fi',
            `\\documentclass[\\officeparserdriver ${classOptions.join(',')}]{${this.docClass}}`,
            // XeLaTeX and LuaLaTeX read Unicode and OpenType fonts; pdfLaTeX, and upLaTeX/pLaTeX (which read CJK
            // natively), use 8-bit fonts.
            '\\iftutex',
            '  \\usepackage{fontspec}',
            ...this.scriptFontLines(unicode.scripts).map(l => `  ${l}`),
            '\\else',
            '  \\usepackage[T1]{fontenc}',
            '  \\usepackage[utf8]{inputenc}',
            '  \\usepackage{lmodern}',
            '\\fi',
        ];
        if (!this.beamer) {
            const cfg = this.config.texConfig;
            const paper = paperSizePt(cfg.format);
            const geometry = [
                `paperwidth=${fmtPt(paper.w)}`, `paperheight=${fmtPt(paper.h)}`,
                ...(cfg.landscape ? ['landscape'] : []),
                `top=${fmtPt(marginPt(cfg.margin?.top))}`, `bottom=${fmtPt(marginPt(cfg.margin?.bottom))}`,
                `left=${fmtPt(marginPt(cfg.margin?.left))}`, `right=${fmtPt(marginPt(cfg.margin?.right))}`,
            ];
            lines.push(`\\usepackage[${geometry.join(',')}]{geometry}`);
            lines.push('\\usepackage{parskip}');
        }
        lines.push(...this.packageLines(unicode, math));
        if (headerSetup) lines.push('\\usepackage{fancyhdr}');
        if (!this.beamer) lines.push('\\usepackage[hyperfootnotes=false]{hyperref}');

        lines.push(`\\hypersetup{${['colorlinks=true', 'linkcolor=blue', 'urlcolor=blue', 'citecolor=blue', ...meta.hypersetup].join(',\n  ')}}`);
        if (!this.beamer && !this.config.texConfig.numberSections) lines.push('\\setcounter{secnumdepth}{-\\maxdimen}');
        // fullflexible keeps the typewriter font's own spacing (the default fixed columns spread
        // letters apart); keepspaces preserves indentation.
        if (this.uses.listings) lines.push('\\lstset{basicstyle=\\ttfamily\\small,columns=fullflexible,keepspaces=true,breaklines=true,showstringspaces=false}');
        // A slide continued over several frames repeats its title; without this beamer also numbers
        // every title of such a slide, including slides that fit on one frame.
        if (this.beamer) lines.push('\\setbeamertemplate{frametitle continuation}{}');
        if (this.uses.paragraphFix && !this.beamer) {
            // \paragraph and \subparagraph are run-in headings by default; give them their own line
            // like every other heading level (a positive after-skip makes a display heading).
            lines.push(
                '\\makeatletter',
                '\\renewcommand\\paragraph{\\@startsection{paragraph}{4}{\\z@}{3.25ex \\@plus1ex \\@minus.2ex}{1.5ex \\@plus .2ex}{\\normalfont\\normalsize\\bfseries}}',
                '\\renewcommand\\subparagraph{\\@startsection{subparagraph}{5}{\\z@}{3.25ex \\@plus1ex \\@minus .2ex}{1.5ex \\@plus .2ex}{\\normalfont\\normalsize\\bfseries}}',
                '\\makeatother',
            );
        }
        lines.push(...this.colorDefinitions());
        if (headerSetup) lines.push(headerSetup);
        lines.push(...this.unicodeLines(unicode));
        if (math.definitions.length) {
            // Made at the start of the document, after every package's own definitions (a command one
            // of them defines keeps its meaning); the LaTeX parser reads the formulas as written.
            lines.push(
                '% Math commands no package loaded here defines, each printed as its name until you define it (or load its package).',
                '\\AtBeginDocument{%',
                ...math.definitions.map(d => `  ${d}%`),
                '}',
            );
        }
        if (meta.title) lines.push(meta.title);
        return lines.join('\n') + '\n';
    }

    /**
     * Declarations for the document's characters some engine cannot typeset as is: symbol and space
     * fallbacks under every engine (`\newunicodechar` for XeLaTeX and LuaLaTeX, `inputenc`'s
     * `\DeclareUnicodeCharacter` for the 8-bit engines), and replacements there for characters
     * pdfLaTeX has no glyph for at all.
     */
    private unicodeLines(unicode: LatexUnicodePlan): string[] {
        if (unicode.eightBitEngines.length === 0) return [];
        const markers = unicode.eightBitEngines.length > unicode.unicodeEngines.length;
        return [
            '\\iftutex',
            ...unicode.unicodeEngines.map(l => `  ${l}`),
            '\\else',
            ...(markers ? [
                '  % pdfLaTeX has no glyph for some of these characters (upLaTeX and pLaTeX set the CJK ones themselves);',
                '  % XeLaTeX and LuaLaTeX typeset them all, in a font that has them.',
            ] : []),
            ...unicode.eightBitEngines.map(l => `  ${l}`),
            '\\fi',
        ];
    }

    /**
     * Fonts under XeLaTeX and LuaLaTeX for the scripts Latin Modern lacks: Computer Modern Unicode for
     * Greek and Cyrillic, and TeX Live's Harano Aji (Japanese), Fandol (Chinese) and UnFonts (Korean)
     * through xeCJK or luatexja, which also break lines between CJK characters. The main CJK font is that
     * of the language the characters are mostly in, the others its fallbacks for what it lacks. Each part applies only where its package and fonts are installed (a missing font is a
     * fatal error), so a smaller TeX installation still compiles the document, those characters blank.
     */
    private scriptFontLines(scripts: LatexScripts): string[] {
        const lines: string[] = [];
        if (scripts.greekCyrillic) {
            lines.push(
                '% Greek and Cyrillic: Computer Modern Unicode, which covers them (Latin Modern does not).',
                '\\IfFontExistsTF{cmunrm.otf}{%',
                '  \\setmainfont{cmunrm.otf}[BoldFont=cmunbx.otf,ItalicFont=cmunti.otf,BoldItalicFont=cmunbi.otf]%',
                '  \\setsansfont{cmunss.otf}[BoldFont=cmunsx.otf,ItalicFont=cmunsi.otf,BoldItalicFont=cmunso.otf]%',
                '  \\setmonofont{cmuntt.otf}[BoldFont=cmuntb.otf,ItalicFont=cmunit.otf,BoldItalicFont=cmuntx.otf]}{}',
            );
        }
        if (!scripts.cjkLanguage) return lines;
        const families = {
            ja: { rm: 'HaranoAjiMincho-Regular.otf', rmBold: 'HaranoAjiMincho-Bold.otf', sf: 'HaranoAjiGothic-Regular.otf', sfBold: 'HaranoAjiGothic-Bold.otf' },
            zh: { rm: 'FandolSong-Regular.otf', rmBold: 'FandolSong-Bold.otf', sf: 'FandolHei-Regular.otf', sfBold: 'FandolHei-Bold.otf' },
            ko: { rm: 'UnBatang.ttf', rmBold: 'UnBatangBold.ttf', sf: 'UnDotum.ttf', sfBold: 'UnDotumBold.ttf' },
        };
        const main = scripts.cjkLanguage;
        // Han and kana a main font lacks come from the other Chinese or Japanese font, Hangul from the Korean one.
        const hanFallbacks = (['zh', 'ja'] as const).filter(k => k !== main);
        const hangulFallback = scripts.hangul && main !== 'ko';
        const xeFallbacks = [...hanFallbacks, ...(hangulFallback ? ['ko' as const] : [])];
        const used = [main, ...xeFallbacks].map(k => families[k]);
        const f = families[main];
        // In a mostly Latin text, the punctuation CJK shares with it (curly quotes, dashes, the ellipsis,
        // arrows and symbols) stays in the Latin font rather than taking full-width CJK forms.
        const latinPunctuation = !scripts.cjkDominant;
        const xe = [
            '\\usepackage{xeCJK}%',
            '\\xeCJKsetup{CJKspace=true}%',
            ...(latinPunctuation ? ['\\xeCJKDeclareCharClass{Default}{"2014, "2018, "2019, "201C, "201D, "2026}%'] : []),
            `\\setCJKmainfont{${f.rm}}[BoldFont=${f.rmBold}]%`,
            `\\setCJKsansfont{${f.sf}}[BoldFont=${f.sfBold}]%`,
            `\\setCJKmonofont{${f.sf}}%`,
            `\\setCJKfallbackfamilyfont{\\CJKrmdefault}{${xeFallbacks.map(k => `{${families[k].rm}}`).join(',')}}%`,
            `\\setCJKfallbackfamilyfont{\\CJKsfdefault}{${xeFallbacks.map(k => `{${families[k].sf}}`).join(',')}}%`,
            `\\setCJKfallbackfamilyfont{\\CJKttdefault}{${xeFallbacks.map(k => `{${families[k].sf}}`).join(',')}}%`,
            '\\xeCJKsetup{AutoFallBack=true}%',
        ];
        // luatexja takes Hangul from its own per-range AltFont: through a luaotfload fallback it fails in luatexja's glue.
        const hangulRanges = ['"1100-"11FF', '"3130-"318F', '"A960-"A97F', '"AC00-"D7FF'];
        const lua = (face: 'rm' | 'sf', bold: boolean) => {
            const opts = [
                ...(bold ? [`BoldFont=${face === 'rm' ? f.rmBold : f.sfBold}`] : []),
                `RawFeature={fallback=officeparser${face}}`,
                ...(hangulFallback ? [`AltFont={${hangulRanges.map(r => `{Range=${r}, Font=${families.ko[face]}}`).join(',')}}`] : []),
            ];
            return `[${opts.join(', ')}]`;
        };
        const fallbackList = (face: 'rm' | 'sf') => hanFallbacks.map(k => `"${families[k][face]}:mode=node;"`).join(', ');
        const luaLines = [
            '\\usepackage{luatexja-fontspec}%',
            `\\ltjsetparameter{jacharrange={-2${latinPunctuation ? ', -3, -9' : ''}}}%`,
            `\\directlua{luaotfload.add_fallback("officeparserrm", {${fallbackList('rm')}})}%`,
            `\\directlua{luaotfload.add_fallback("officeparsersf", {${fallbackList('sf')}})}%`,
            `\\setmainjfont{${f.rm}}${lua('rm', true)}%`,
            `\\setsansjfont{${f.sf}}${lua('sf', true)}%`,
            `\\setmonojfont{${f.sf}}${lua('sf', false)}%`,
        ];
        // One regular font per package stands for its whole family.
        const guards = used.map(u => u.rm);
        lines.push(
            '% Chinese, Japanese and Korean: TeX Live\'s Harano Aji, Fandol and UnFonts, through xeCJK or luatexja.',
            `${guards.map(g => `\\IfFontExistsTF{${g}}{`).join('')}%`,
            '\\ifXeTeX',
            '  \\IfFileExists{xeCJK.sty}{%',
            ...xe.map(l => `    ${l}`),
            '  }{}%',
            '\\else',
            '  \\IfFileExists{luatexja-fontspec.sty}{%',
            ...luaLines.map(l => `    ${l}`),
            '  }{}%',
            `\\fi${guards.map(() => '}{}').join('')}`,
        );
        return lines;
    }

    /** The comment heading a body-only fragment: the packages the including document must load. */
    private fragmentHeader(unicode: LatexUnicodePlan, math: LatexMathPlan): string {
        const unicodeLines = this.unicodeLines(unicode);
        const fonts = this.scriptFontLines(unicode.scripts);
        // The per-engine declarations test the engine with iftex's \iftutex.
        const packages = [...(unicodeLines.length || fonts.length ? ['\\usepackage{iftex}'] : []), ...this.packageLines(unicode, math)];
        if (this.uses.graphics && this.beamer) packages.push('\\usepackage{graphicx}');
        packages.push('\\usepackage{hyperref}');
        // Under XeLaTeX and LuaLaTeX (after fontspec), the fonts for the scripts Latin Modern lacks.
        const fontSetup = fonts.length ? ['\\iftutex', ...fonts.map(l => `  ${l}`), '\\fi'] : [];
        return latexComment(['LaTeX fragment generated by officeParser. The including document needs, in its preamble:', ...packages, ...fontSetup, ...unicodeLines].join('\n'));
    }
}

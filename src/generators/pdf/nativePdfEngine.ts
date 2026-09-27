/**
 * A first-party PDF layout engine built directly on `pdf-lib`, an alternative to the default
 * HTML-through-a-headless-browser route (`PdfGenerator`'s Puppeteer path).
 *
 * It walks the AST and paints it onto pages itself - text with inline bold/italic/size/color,
 * headings, lists (bulleted and numbered, nested), grid tables with borders, embedded images, page
 * breaks and horizontal rules - handling word wrapping and pagination as it goes. It needs no browser
 * and runs unchanged in Node and the browser, so it is the way to produce a real PDF client-side (the
 * browser Puppeteer path can only hand back HTML). Its trade-off is fidelity: it uses the Standard-14
 * Standard-14 fonts (the Helvetica and Courier faces; Times is not embedded) rather than the
 * document's own, and its layout is a clean reflow, not a pixel match of the HTML renderer. Choose it
 * with `pdfConfig.engine: 'native'`.
 *
 * `pdf-lib` is an optional peer dependency, imported lazily so it is only required when this engine
 * is selected.
 *
 * @module generators/pdf/nativePdfEngine
 */

import { EmbedMetadata, FullGeneratorConfig, ImageMode, OfficeContentNode, OfficeErrorType, OfficeMetadata, OfficeParserAST, OfficeWarningType, TextFormatting } from '../../types.js';
import { getAbortError, getOfficeError } from '../../utils/errorUtils.js';
import { documentLanguage, embedUrl, isHeaderRow, lengthToPt, paperSizePt, resolveImageMode, sniffImageSize } from '../../utils/officeGenUtils.js';

/**
 * Megapixel ceiling for an embedded image. `embedPng`/`embedJpg` decode the full bitmap, so a
 * decompression-bomb image (a few header bytes declaring enormous dimensions) would force a multi-GB,
 * uncatchable allocation. A real image is far below this; a larger one degrades to its alt text.
 */
const MAX_IMAGE_PIXELS = 40_000_000;

/**
 * Code points WinAnsi (CP1252) encodes beyond Latin-1, which the Standard-14 fonts accept (smart
 * quotes, dashes, bullet, ellipsis, trademark, euro, etc.). Everything outside Latin-1 and this set
 * has no glyph in those fonts and makes pdf-lib throw, so it is replaced rather than drawn.
 */
const CP1252_EXTRA = new Set([0x20AC, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x017D, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x017E, 0x0178]);

/** Maps a string to the WinAnsi-encodable subset the Standard-14 fonts can draw, flagging any loss. */
function toWinAnsi(text: string): { text: string; changed: boolean } {
    let out = '', changed = false;
    for (const ch of text) {
        const cp = ch.codePointAt(0)!;
        if ((cp >= 0x20 && cp <= 0x7E) || (cp >= 0xA0 && cp <= 0xFF) || CP1252_EXTRA.has(cp)) out += ch;
        else if (cp === 0x09) out += '    ';                          // tab -> spaces
        else if (cp === 0x0A || cp === 0x0D) out += ' ';             // stray newline in a single line
        // Zero-width and formatting characters (ZWSP/ZWNJ/ZWJ/word-joiner/BOM) carry no glyph; drop them
        // silently rather than drawing a visible "?" mid-word and flagging a spurious loss.
        else if (cp === 0x200B || cp === 0x200C || cp === 0x200D || cp === 0x2060 || cp === 0xFEFF) { /* drop */ }
        else { out += '?'; changed = true; }
    }
    return { text: out, changed };
}

/** Lazily loads `pdf-lib`, throwing a typed, actionable error when it is not installed. */
async function loadPdfLib(config: OfficeParserAST['config']): Promise<any> {
    try {
        // @ts-ignore - optional peer dependency, resolved at runtime
        return await import('pdf-lib');
    } catch (e) {
        throw getOfficeError(OfficeErrorType.PDF_GENERATION_FAILED, config,
            "The native PDF engine requires the optional peer dependency 'pdf-lib'. Install it with `npm install pdf-lib`, or use the default engine (pdfConfig.engine: 'html').");
    }
}

/** Resolves a numeric length from a pdf.js-style value that may be a number (px) or a unit string. */
function toPoints(v: string | number | undefined, fallback: number): number {
    if (typeof v === 'number' && Number.isFinite(v)) return v * 0.75; // css px -> pt
    if (typeof v === 'string') {
        const m = /^([\d.]+)\s*(pt|px|in|cm|mm)?$/i.exec(v.trim());
        if (m) {
            const n = parseFloat(m[1]);
            switch (m[2]?.toLowerCase()) {
                case 'in': return n * 72;
                case 'cm': return n * 28.3465;
                case 'mm': return n * 2.83465;
                case 'pt': return n;
                default: return n * 0.75; // px
            }
        }
    }
    return fallback;
}

/** A single word ready to place, carrying the font/size/color it should be drawn with. */
interface Piece { text: string; font: any; size: number; color: any; underline: boolean; }

/** The embedded Standard-14 fonts the engine draws with. */
interface Fonts { regular: any; bold: any; italic: any; boldItalic: any; mono: any; }

/**
 * Stateful layout cursor: owns the pdf-lib document, the current page and the y position, and knows
 * how to wrap text, paginate, and paint each node type.
 */
class NativeLayout {
    private page: any;
    /** Distance of the cursor from the top of the page, in points (converted to pdf-lib's y-up on draw). */
    private y = 0;
    /** Resolved image handling, shared with every other generator via {@link resolveImageMode}. */
    private readonly imageMode: ImageMode;
    /** `name -> attachment` index, so image lookups are O(1) rather than a scan per image node. */
    private readonly attachmentsByName = new Map<string, OfficeParserAST['attachments'][number]>();
    /** Footnote/endnote bodies gathered from `node.notes` during the walk, drawn at document end. */
    private readonly collectedNotes: OfficeContentNode[] = [];

    constructor(
        private readonly pdf: any,
        private readonly lib: any,
        private readonly fonts: Fonts,
        private readonly pageW: number,
        private readonly pageH: number,
        private readonly margin: { top: number; right: number; bottom: number; left: number },
        private readonly config: FullGeneratorConfig,
        private readonly ast: OfficeParserAST,
        // Reports a warning through the generator's own channel, so it lands in `result.messages`
        // (and reaches onWarning) rather than only firing onWarning directly.
        private readonly reportWarning: (type: OfficeWarningType, info?: any) => void,
    ) {
        this.imageMode = resolveImageMode(this.config.includeImages);
        for (const a of this.ast.attachments || []) {
            if (a.name && !this.attachmentsByName.has(a.name)) this.attachmentsByName.set(a.name, a);
        }
        this.newPage();
    }

    private get contentWidth(): number { return this.pageW - this.margin.left - this.margin.right; }
    private get bottom(): number { return this.pageH - this.margin.bottom; }

    /** Starts a fresh page and resets the cursor to the top margin. */
    private newPage(): void {
        this.page = this.pdf.addPage([this.pageW, this.pageH]);
        this.y = this.margin.top;
    }

    /**
     * Ensures `h` points of vertical space remain, starting a new page if not. Never starts a page
     * when the cursor is already at the top margin: the block does not fit a whole page, so a blank
     * page in front of it would not help and would just pad the output. Callers that must not lose
     * content taller than a page (tables) split it across pages themselves.
     */
    private ensureSpace(h: number): void {
        if (this.y + h > this.bottom && this.y > this.margin.top) this.newPage();
    }

    /** Picks the font matching a run's formatting (mono for code, bold/italic variants otherwise). */
    private fontFor(fmt: TextFormatting | undefined, mono: boolean): any {
        if (mono) return this.fonts.mono;
        const bold = !!fmt?.bold, italic = !!fmt?.italic;
        if (bold && italic) return this.fonts.boldItalic;
        if (bold) return this.fonts.bold;
        if (italic) return this.fonts.italic;
        return this.fonts.regular;
    }

    /** Parses a `#rrggbb` (or `#rgb`) hex color into a pdf-lib rgb color, or null. */
    private color(hex: string | undefined): any {
        if (!hex || typeof hex !== 'string') return null;
        let h = hex.trim().replace(/^#/, '');
        if (h.length === 3) h = h.split('').map(c => c + c).join('');
        if (!/^[0-9a-f]{6}$/i.test(h)) return null;
        const n = parseInt(h, 16);
        return this.lib.rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255);
    }

    private winAnsiWarned = false;
    /** Whether the embed-as-link warning was given (once per document). */
    private embedWarned = false;
    /** WinAnsi-sanitizes text before it is measured or drawn, warning once if any character is lost. */
    private enc(text: string): string {
        const r = toWinAnsi(text);
        if (r.changed && !this.winAnsiWarned) {
            this.winAnsiWarned = true;
            const info = { format: 'pdf', feature: 'non-WinAnsi characters (native engine)' };
            this.reportWarning(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, info);
        }
        return r.text;
    }

    /**
     * Flattens a node into inline text runs (text + formatting + link flag), recursing into children.
     *
     * The caller's `onNode` hook is applied to every descendant visited here - a text run is a node
     * like any other, and `CommonGeneratorConfig.onNode` promises every node - so `false` drops the run
     * (and its subtree) and a returned string replaces its text. The root `node` itself is NOT hooked:
     * whoever passes it here (render(), or the table/cell walk) has already applied the hook to it, so
     * hooking it again would fire twice for the same node.
     */
    private async collectRuns(node: OfficeContentNode): Promise<{ text: string; fmt: TextFormatting; link: boolean }[]> {
        const out: { text: string; fmt: TextFormatting; link: boolean }[] = [];
        const walk = async (n: OfficeContentNode, inherited: TextFormatting, isRoot: boolean): Promise<void> => {
            let text = n.text;
            let override: string | false | void = undefined;
            if (!isRoot) {
                override = await this.onNodeValue(n);
                if (override === false) return;
                if (typeof override === 'string') text = override;
            }
            const fmt = { ...inherited, ...(n.formatting || {}) };
            const link = !!(n.metadata as any)?.link;
            // A picture in a line the engine draws as text (a list item, heading or note; a paragraph
            // draws its pictures itself) is its alt or recognized text: it was dropped.
            if (n.type === 'image' && !isRoot) {
                const alt = this.inlineImageText(n, typeof override === 'string' ? override : undefined);
                if (alt) out.push({ text: alt, fmt, link });
                return;
            }
            if (n.type === 'text' || (!n.children?.length && text)) {
                if (text) out.push({ text, fmt, link });
                return;
            }
            for (const c of n.children || []) await walk(c, fmt, false);
        };
        await walk(node, {}, true);
        if (!out.length && node.text) out.push({ text: node.text, fmt: node.formatting || {}, link: false });
        return out;
    }

    /**
     * The plain string a table cell (or any node the engine paints as flat text rather than as a block)
     * renders as. Cells built by the markup parsers carry their content only in `children` with no
     * `text` of their own, so reading `cell.text` alone painted every Markdown/HTML table empty.
     * Block children each own a line, which `wrapPlain` keeps as separate lines.
     */
    private async flatText(node: OfficeContentNode): Promise<string> {
        const children = node.children || [];
        // Inline-only children (the common Markdown cell) flatten to a single line; block children
        // (paragraphs, headings, lists, a nested table's rows) are each a line of their own. An image
        // in a cell counts too: the engine cannot embed a picture inside a cell, so it degrades to the
        // image's alt (or OCR) text rather than vanishing silently.
        if (children.some(c => BLOCK_CELL_TYPES.has(c.type) || c.type === 'image')) {
            const lines: string[] = [];
            for (const c of children) {
                // Each block child is a node in its own right, so it gets the hook here (collectRuns
                // only hooks the descendants of whatever root it is handed).
                const override = await this.onNodeValue(c);
                if (override === false) continue;
                let t: string;
                if (typeof override === 'string') t = override;
                else if (c.type === 'image') t = this.cellImageText(c);
                else t = (await this.collectRuns(c)).map(r => r.text).join('');
                if (t !== '') lines.push(t);
            }
            if (lines.length) return lines.join('\n');
        }
        const flat = (await this.collectRuns(node)).map(r => r.text).join('');
        return flat || node.text || '';
    }

    private inlineImageWarned = false;
    /**
     * The text a picture in a line of text degrades to, as the image mode has it (none for 'none', the
     * recognized text for 'ocr-text-only', else the alt text or the recognized text), warning once that
     * the native engine draws such a picture as text. `override` is the onNode hook's replacement.
     */
    private inlineImageText(node: OfficeContentNode, override?: string): string {
        if (override !== undefined) return override;
        if (this.imageMode === 'none') return '';
        const ocr = (node.text || '').trim();
        const text = this.imageMode === 'ocr-text-only' ? ocr : (((node.metadata as any)?.altText || '').trim() || ocr);
        if (text && !this.inlineImageWarned) {
            this.inlineImageWarned = true;
            this.reportWarning(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'pdf', feature: 'image inside a list item, heading or note (native engine renders its alt/OCR text)' });
        }
        return text;
    }

    private cellImageWarned = false;
    /** The text an image in a table cell degrades to (alt text, else OCR text), warning once that the
     * native engine cannot embed a picture inside a cell. */
    private cellImageText(node: OfficeContentNode): string {
        if (!this.cellImageWarned) {
            this.cellImageWarned = true;
            this.reportWarning(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'pdf', feature: 'image inside a table cell (native engine renders its alt/OCR text)' });
        }
        return ((node.metadata as any)?.altText || node.text || '').trim();
    }

    /**
     * Wraps a set of inline runs into lines within `[left, left+width]` and draws them, paginating as
     * needed. Returns nothing; advances the cursor to just below the last line.
     */
    private drawRuns(runs: { text: string; fmt: TextFormatting; link: boolean }[], left: number, width: number, mono = false): void {
        // Expand runs into per-word pieces (a trailing-space flag lets us re-insert inter-word gaps).
        const pieces: { p: Piece; space: boolean }[] = [];
        for (const run of runs) {
            const font = this.fontFor(run.fmt, mono);
            const size = parseFontSize(run.fmt.size) ?? (mono ? 10 : 11);
            const color = run.link ? this.lib.rgb(0.1, 0.32, 0.72) : (this.color(run.fmt.color) ?? this.lib.rgb(0.12, 0.12, 0.12));
            const underline = !!(run.fmt.underline || run.link);
            const words = this.enc(run.text).split(/(\s+)/);
            for (const w of words) {
                if (w === '') continue;
                if (/^\s+$/.test(w)) { if (pieces.length) pieces[pieces.length - 1].space = true; continue; }
                // Split a word wider than the whole line at character boundaries so it wraps instead
                // of running off the page (a long URL, a hash, a base64 token).
                for (const chunk of splitToWidth(w, font, size, width)) {
                    pieces.push({ p: { text: chunk, font, size, color, underline }, space: false });
                }
            }
        }
        if (!pieces.length) return;

        const spaceWidth = (p: Piece) => p.font.widthOfTextAtSize(' ', p.size);
        let line: { p: Piece; space: boolean; w: number }[] = [];
        let lineWidth = 0;

        const flush = () => {
            if (!line.length) return;
            const lh = Math.max(...line.map(it => it.p.size)) * 1.35;
            this.ensureSpace(lh);
            let x = left;
            const baseline = this.pageH - this.y - Math.max(...line.map(it => it.p.size));
            for (const it of line) {
                this.page.drawText(it.p.text, { x, y: baseline, size: it.p.size, font: it.p.font, color: it.p.color });
                if (it.p.underline) {
                    this.page.drawLine({ start: { x, y: baseline - 1.5 }, end: { x: x + it.w, y: baseline - 1.5 }, thickness: 0.5, color: it.p.color });
                }
                x += it.w + (it.space ? spaceWidth(it.p) : 0);
            }
            this.y += lh;
            line = [];
            lineWidth = 0;
        };

        for (const piece of pieces) {
            const w = piece.p.font.widthOfTextAtSize(piece.p.text, piece.p.size);
            // Only the previous piece's OWN trailing space counts toward width: a word split across two
            // runs (or a splitToWidth chunk) has space:false and draws with no gap, so measuring one
            // there wrapped such lines slightly early and disagreed with what flush() draws.
            const prev = line[line.length - 1];
            if (prev && lineWidth + (prev.space ? spaceWidth(prev.p) : 0) + w > width) flush();
            const prevAfter = line[line.length - 1];
            lineWidth += (prevAfter && prevAfter.space ? spaceWidth(prevAfter.p) : 0) + w;
            line.push({ p: piece.p, space: piece.space, w });
        }
        flush();
    }

    /**
     * The caller's `onNode` verdict for one node, normalized to `false` (drop it) / a replacement
     * string / undefined (render normally). Used for the nodes the engine paints as flat text - text
     * runs, table rows and cells - where a returned string is spliced in as text rather than drawn as
     * its own paragraph, so no "not representable" warning is owed.
     */
    private async onNodeValue(node: OfficeContentNode): Promise<false | string | undefined> {
        if (!this.config.onNode) return undefined;
        const override = await this.config.onNode(node);
        if (override === false) return false;
        return typeof override === 'string' ? override : undefined;
    }

    private onNodeWarned = false;
    /**
     * Applies the caller's `onNode` hook, matching {@link BaseGenerator.handleOnNode}: `false` skips the
     * node and its subtree; a returned string replaces the node's output. The native engine paints a
     * layout rather than emitting markup, so it cannot splice a returned string in structurally - it
     * draws it as a plain-text paragraph (warning once) so the override is honored rather than lost.
     * Returns true when the hook handled the node and default rendering should be skipped.
     */
    private async applyOnNode(node: OfficeContentNode): Promise<boolean> {
        if (!this.config.onNode) return false;
        const override = await this.config.onNode(node);
        if (override === false) return true;
        if (typeof override === 'string') {
            if (!this.onNodeWarned) {
                this.onNodeWarned = true;
                const info = { format: 'pdf', feature: 'onNode string override (native engine draws it as plain text)' };
                this.reportWarning(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, info);
            }
            if (override) this.drawRuns([{ text: override, fmt: {}, link: false }], this.margin.left, this.contentWidth);
            this.y += 6;
            return true;
        }
        return false;
    }

    private prevPaginated: string | null = null;
    /** Entry point: renders any node, dispatching by type. `topLevel` drives inter-page pagination. */
    async render(node: OfficeContentNode, topLevel = false): Promise<void> {
        if (this.config.abortSignal?.aborted) throw getAbortError();
        // Preserve source pagination: start a fresh page between consecutive page (or slide) nodes.
        // Done BEFORE the hook so a page/slide the hook drops or replaces with a string still advances
        // the bookkeeping (the next page gets its break, and a replacement string lands on a fresh page)
        // instead of leaving prevPaginated stale.
        if (topLevel && (node.type === 'page' || node.type === 'slide')) {
            if (this.prevPaginated === node.type) this.newPage();
            this.prevPaginated = node.type;
        } else if (topLevel) {
            this.prevPaginated = null;
        }
        // The caller's onNode hook can skip a node/subtree or replace its output, exactly as the
        // text-based generators honor it; applied before any default rendering.
        if (await this.applyOnNode(node)) return;
        switch (node.type) {
            case 'page':
            case 'slide':
            case 'header':
            case 'footer':
                for (const c of node.children || []) await this.render(c);
                this.y += 4;
                return;
            // A spreadsheet is a grid: its rows/cells carry `col`, so lay it out as a table rather than
            // letting each cell fall through to `default` and render as its own stacked paragraph. Its
            // drawing images/charts sit as non-row children after the rows; render them after the grid
            // (as the HTML/DOCX/ODT generators do) rather than dropping them.
            case 'sheet': {
                await this.table(node);
                for (const c of (node.children || [])) if (c.type !== 'row') await this.render(c);
                return;
            }
            case 'heading': return this.heading(node);
            case 'paragraph': return this.paragraph(node);
            case 'list': return this.listItem(node);
            case 'table': return this.table(node);
            case 'image': return this.image(node);
            case 'code': return this.code(node);
            case 'note': return this.note(node);
            case 'break': return this.breakNode(node);
            case 'embed': {
                // A video or framed page cannot play in a PDF: its URL is written, styled as a link (with
                // its label), as the DOCX and ODT generators write it. It was dropped.
                const meta = node.metadata as EmbedMetadata | undefined;
                const url = embedUrl(meta);
                const text = [meta?.label, url].filter(Boolean).join(': ');
                if (!text) return;
                if (!this.embedWarned) {
                    this.embedWarned = true;
                    this.reportWarning(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'pdf', feature: 'embed' });
                }
                this.drawRuns([{ text, fmt: {}, link: !!url }], this.margin.left, this.contentWidth);
                this.y += 6;
                return;
            }
            case 'chart':
                // includeCharts: false omits charts in every generator. The native engine cannot draw a
                // chart, so it otherwise renders the chart's data text (in node.text); skip it when off.
                if (this.config.includeCharts === false) return;
                if (node.text) await this.paragraph(node);
                return;
            default:
                if (node.children?.length) { for (const c of node.children) await this.render(c); }
                else if (node.text) await this.paragraph(node);
                return;
        }
    }

    private async heading(node: OfficeContentNode): Promise<void> {
        const level = Math.min(6, Math.max(1, (node.metadata as any)?.level || 1));
        const size = [24, 20, 16, 14, 12, 11][level - 1];
        // Keep-with-next: a heading should not be stranded as the last thing on a page. Reserve its own
        // line plus a couple of body lines, so a heading near the bottom moves to the next page with the
        // text it titles (the HTML engine expresses the same intent with `page-break-after: avoid`).
        // ensureSpace is a no-op at the top margin, so this never inserts a blank page before a heading.
        this.ensureSpace(size * 1.35 + 11 * 1.35 * 2);
        // Space-before, but not when the heading is already at the very top of a page (including one it
        // was just pushed onto by the reservation above): a gap there would float it below the margin.
        if (this.y > this.margin.top) this.y += size * 0.6;
        const runs = (await this.collectRuns(node)).map(r => ({ ...r, fmt: { ...r.fmt, bold: true, size: `${size}pt` } }));
        this.drawRuns(runs.length ? runs : [{ text: node.text || '', fmt: { bold: true, size: `${size}pt` }, link: false }], this.margin.left, this.contentWidth);
        this.y += size * 0.35;
    }

    private async paragraph(node: OfficeContentNode, indentLeft = 0): Promise<void> {
        // A picture in the paragraph (where Word, ODF and Markdown put one) is drawn where it stands,
        // between the lines of text around it: runs carry text only, and it was dropped.
        if (node.children?.some(child => child.type === 'image')) {
            let run: OfficeContentNode[] = [];
            const flush = async () => {
                const runs = run.length ? await this.collectRuns({ ...node, text: undefined, children: run }) : [];
                if (runs.some(r => r.text.trim())) {
                    this.drawRuns(runs, this.margin.left + indentLeft, this.contentWidth - indentLeft);
                    this.y += 6;
                }
                run = [];
            };
            for (const child of node.children) {
                if (child.type !== 'image') { run.push(child); continue; }
                await flush();
                await this.render(child);
            }
            await flush();
            return;
        }
        const runs = await this.collectRuns(node);
        if (!runs.length) { this.y += 6; return; }
        this.drawRuns(runs, this.margin.left + indentLeft, this.contentWidth - indentLeft);
        this.y += 6;
    }

    /** Draws one already-known plain string as a paragraph (an image's alt text, a note marker). */
    private plainParagraph(text: string): void {
        if (!text) return;
        this.drawRuns([{ text, fmt: {}, link: false }], this.margin.left, this.contentWidth);
        this.y += 6;
    }

    private async listItem(node: OfficeContentNode): Promise<void> {
        const meta = node.metadata as any;
        const level = Math.max(0, meta?.indentation || 0);
        const indent = 18 + level * 18;
        const ordered = meta?.listType === 'ordered';
        const marker = ordered ? `${(meta?.itemIndex ?? 0) + 1}.` : '•';
        // Body: the item's own text runs (its non-list children); nested list children render after.
        const bodyRuns = await this.collectRuns({ ...node, children: (node.children || []).filter(c => c.type !== 'list') });
        // Size the marker to the body text so their baselines line up (a body larger than the old fixed
        // 11pt would otherwise sit below its own marker), and reserve one line's height BEFORE drawing
        // the marker so a body that has to break to the next page carries its marker along instead of
        // orphaning it on this page.
        const bodySize = bodyRuns.length ? Math.max(...bodyRuns.map(r => parseFontSize(r.fmt.size) ?? 11)) : 11;
        const size = bodySize > 0 ? bodySize : 11;
        this.ensureSpace(size * 1.35);
        // Draw the marker, then the item body hanging-indented past it. Space is already reserved, so
        // drawRuns' first line stays on this page and draws at the same baseline as the marker.
        const markerX = this.margin.left + indent - 14;
        const baseline = this.pageH - this.y - size;
        this.page.drawText(this.enc(marker), { x: markerX, y: baseline, size, font: this.fonts.regular, color: this.lib.rgb(0.12, 0.12, 0.12) });
        // Advance for the marker line unless the body has real text to draw: a whitespace-only body
        // makes drawRuns place nothing and never advance `this.y`, so the next item would overprint.
        if (bodyRuns.some(r => r.text && r.text.trim())) this.drawRuns(bodyRuns, this.margin.left + indent, this.contentWidth - indent);
        else this.y += size * 1.35;
        // Route nested list items through render() (not listItem() directly) so the onNode hook and the
        // abort check see them like every other node.
        for (const c of (node.children || []).filter(c => c.type === 'list')) await this.render(c);
    }

    private code(node: OfficeContentNode): void {
        this.y += 2;
        this.drawMonospace(String(node.text || ''));
        this.y += 6;
    }

    /**
     * Draws text line-by-line in a monospace font, preserving leading indentation and column alignment
     * (word wrapping would collapse both). Used for code blocks and for OCR text, whose 2-D page layout
     * a flowing paragraph would destroy. Split BEFORE encoding: `enc` maps every newline to a space
     * (a drawn string cannot contain one), so encoding first would collapse the block onto one line.
     */
    private drawMonospace(text: string): void {
        const size = 10, lh = size * 1.35;
        for (const raw of text.split(/\r\n|\r|\n/)) {
            const expanded = this.enc(raw.replace(/\t/g, '    '));
            const chunks = splitToWidth(expanded || ' ', this.fonts.mono, size, this.contentWidth - 12);
            for (const chunk of (chunks.length ? chunks : [' '])) {
                this.ensureSpace(lh);
                this.page.drawText(chunk, { x: this.margin.left + 6, y: this.pageH - this.y - size, size, font: this.fonts.mono, color: this.lib.rgb(0.12, 0.12, 0.12) });
                this.y += lh;
            }
        }
    }

    /** Draws OCR text: multi-line keeps its 2-D layout in monospace, a single line flows as prose. */
    private async drawOcrText(text: string): Promise<void> {
        if (/[\r\n]/.test(text)) { this.y += 2; this.drawMonospace(text); this.y += 4; }
        else await this.plainParagraph(text);
    }

    private async note(node: OfficeContentNode): Promise<void> {
        const runs = (await this.collectRuns(node)).map(r => ({ ...r, fmt: { ...r.fmt, size: '9pt' } }));
        if (runs.length) this.drawRuns(runs, this.margin.left + 12, this.contentWidth - 12);
        this.y += 4;
    }

    private breakNode(node: OfficeContentNode): void {
        const t = (node.metadata as any)?.breakType;
        if (t === 'page') { this.newPage(); return; }
        if (t === 'thematic') {
            this.ensureSpace(12);
            this.y += 6;
            const yy = this.pageH - this.y;
            this.page.drawLine({ start: { x: this.margin.left, y: yy }, end: { x: this.pageW - this.margin.right, y: yy }, thickness: 0.5, color: this.lib.rgb(0.7, 0.7, 0.7) });
            this.y += 6;
        } else {
            this.y += 8;
        }
    }

    private async image(node: OfficeContentNode): Promise<void> {
        // Honor the same includeImages / imageMode contract as every other generator (HTML, DOCX, ODT):
        // 'none' drops the node entirely; 'ocr-text-only' emits only the recognized text; 'image-only'
        // and 'image+ocr-text' embed the picture, the latter also drawing its OCR text underneath.
        const mode = this.imageMode;
        if (mode === 'none') return;
        const ocr = (node.text || '').trim();
        if (mode === 'ocr-text-only') { if (ocr) await this.drawOcrText(ocr); return; }

        const meta = node.metadata as any;
        const name = meta?.attachmentName;
        // Nothing renderable: keep the alt text (or the OCR text) rather than drawing nothing, and
        // report the failure - the DOCX and ODT generators degrade and warn for the same input, and a
        // picture that silently vanishes gives the caller no way to find out.
        const fallback = () => this.plainParagraph(meta?.altText || ocr);
        const attachment = name ? this.attachmentsByName.get(name) : undefined;
        if (!attachment?.data) {
            if (name) this.reportWarning(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name, reason: attachment ? 'attachment has no data' : 'missing attachment' });
            fallback();
            return;
        }
        try {
            const bytes = base64ToBytes(attachment.data);
            // Reject an image whose declared dimensions are absurd BEFORE decoding it: embedPng/embedJpg
            // allocate the full bitmap, so a decompression bomb would OOM the process uncatchably.
            const dim = sniffImageSize(bytes);
            if (dim && dim.w * dim.h > MAX_IMAGE_PIXELS) {
                this.reportWarning(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name, reason: `image is too large to embed (${dim.w}x${dim.h} pixels)` });
                fallback();
                return;
            }
            const isJpg = /jpe?g/i.test(attachment.extension || '') || attachment.mimeType === 'image/jpeg';
            const img = isJpg ? await this.pdf.embedJpg(bytes) : await this.pdf.embedPng(bytes);
            // Resolve the draw size in points, by priority: an explicit ImageMetadata.width/height
            // (a length like '200px'/'3cm', a number in px, or a '%' of the content width), then PDF
            // page bounds (already points), then the intrinsic pixel size converted px->pt. Aspect ratio
            // is preserved from the intrinsic image when only a width is given.
            const ratio = img.width ? img.height / img.width : 1;
            const meta2 = node.metadata as any;
            const pctOf = (v: any): number | null => (typeof v === 'string' && /^\s*[\d.]+\s*%\s*$/.test(v)) ? parseFloat(v) / 100 : null;
            let w: number, h: number;
            if (meta2?.width != null) {
                const p = pctOf(meta2.width);
                w = p != null ? this.contentWidth * p : toPoints(meta2.width, img.width * 0.75);
                h = (meta2.height != null && pctOf(meta2.height) == null) ? toPoints(meta2.height, w * ratio) : w * ratio;
            } else if (node.bounds?.width) {
                w = node.bounds.width;
                h = node.bounds.height || w * ratio;
            } else {
                w = img.width * 0.75; // intrinsic pixels -> points (96dpi), not points 1:1
                h = img.height * 0.75;
            }
            const scale = Math.min(1, this.contentWidth / w);
            w *= scale; h *= scale;
            if (h > this.bottom - this.margin.top) { const s = (this.bottom - this.margin.top) / h; w *= s; h *= s; }
            this.ensureSpace(h + 6);
            this.page.drawImage(img, { x: this.margin.left, y: this.pageH - this.y - h, width: w, height: h });
            this.y += h + 6;
        } catch {
            // pdf-lib embeds PNG and JPEG only, so a GIF/BMP/WebP/TIFF/SVG attachment lands here.
            this.reportWarning(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name, reason: 'unsupported image format (the native PDF engine embeds PNG and JPEG only)' });
            fallback();
            return;
        }
        // 'image+ocr-text': draw the recognized text just below the successfully embedded image.
        if (mode === 'image+ocr-text' && ocr) await this.drawOcrText(ocr);
    }

    /**
     * Deep-collects footnote/endnote bodies in document order. Parsers attach `.notes` to the TEXT
     * RUN a marker sits in (e.g. WordParser), and a run is never passed to `render()` (its paragraph
     * renders it via `collectRuns`), so a render()-level sweep misses them. Walk the whole subtree.
     */
    collectNotes(node: OfficeContentNode): void {
        if (!node) return;
        if (node.notes?.length) this.collectedNotes.push(...node.notes);
        for (const c of node.children || []) this.collectNotes(c);
    }

    /** Draws the footnote/endnote bodies gathered during the walk, under a short separating rule. */
    async flushNotes(): Promise<void> {
        if (!this.collectedNotes.length) return;
        this.y += 8;
        this.ensureSpace(14);
        const yy = this.pageH - this.y;
        this.page.drawLine({ start: { x: this.margin.left, y: yy }, end: { x: this.margin.left + Math.min(180, this.contentWidth), y: yy }, thickness: 0.5, color: this.lib.rgb(0.7, 0.7, 0.7) });
        this.y += 8;
        for (const n of this.collectedNotes) await this.note(n);
    }

    private async table(node: OfficeContentNode): Promise<void> {
        const rows = (node.children || []).filter(r => r.type === 'row');
        if (!rows.length) return;
        // Total columns = the widest grid extent any cell reaches (col + colSpan), falling back to
        // the largest raw cell count when cells carry no `col`.
        let cols = 0;
        for (const row of rows) {
            for (const cell of (row.children || []).filter(c => c.type === 'cell')) {
                const col = (cell.metadata as CellMeta)?.col;
                const span = (cell.metadata as CellMeta)?.colSpan || 1;
                cols = Math.max(cols, (typeof col === 'number' ? col : 0) + span);
            }
        }
        cols = Math.max(cols, ...rows.map(r => (r.children || []).filter(c => c.type === 'cell').length), 1);
        const colW = this.contentWidth / cols;
        const size = 10, pad = 4;
        const border = this.lib.rgb(0.6, 0.6, 0.6);

        this.y += 2;
        /** Wraps a cell's (un-encoded) text to the width of the columns it spans. */
        const wrap = (text: string, span: number) => wrapPlain(text, this.fonts.regular, size, Math.max(10, colW * span - 2 * pad), t => this.enc(t));
        // Grid occupancy: column -> rows still covered by a rowspan from above, so a spanning cell
        // reserves its column(s) below instead of letting later rows slide left under it.
        const carry = new Map<number, number>();
        for (const row of rows) {
            const rowOverride = await this.onNodeValue(row);
            if (rowOverride === false) continue;
            const cells = (row.children || []).filter(c => c.type === 'cell');
            // Header detection via the shared heuristic every other generator uses (row/cell `style`,
            // `isHeader`, or an all-bold first row), so a bold first row from DOCX/Markdown bolds here
            // too - not only PDF's per-cell `style: 'header'`.
            const rowIsHeader = isHeaderRow(row, row === rows[0]);
            // Place each cell into a grid column: honour explicit `col`, otherwise the next free one.
            let cursor = 0;
            const placed: { col: number; span: number; rowSpan: number; lines: string[]; header: boolean }[] = [];
            if (typeof rowOverride === 'string') {
                // An onNode string for a whole row: paint it across the full grid width, since the
                // engine has no markup to splice a replacement row into.
                placed.push({ col: 0, span: cols, rowSpan: 1, lines: wrap(rowOverride, cols), header: rowIsHeader });
            }
            for (const cell of typeof rowOverride === 'string' ? [] : cells) {
                const override = await this.onNodeValue(cell);
                if (override === false) continue;
                const meta = cell.metadata as CellMeta;
                const span = meta?.colSpan && meta.colSpan > 1 ? meta.colSpan : 1;
                let col = typeof meta?.col === 'number' ? meta.col : -1;
                if (col < 0) { while ((carry.get(cursor) || 0) > 0) cursor++; col = cursor; }
                // Cells built by the markup parsers carry their text only in `children`; reading
                // `cell.text` alone rendered every Markdown/HTML table as an empty grid.
                const text = typeof override === 'string' ? override : await this.flatText(cell);
                placed.push({ col, span, rowSpan: Math.max(1, meta?.rowSpan || 1), lines: wrap(text, span), header: rowIsHeader || meta?.style === 'header' });
                cursor = col + span;
            }
            const lineH = size * 1.35;
            const rowH = Math.max(size * 1.4, ...placed.map(p => p.lines.length * lineH)) + 2 * pad;
            // Draw the row in vertical bands so a row taller than a whole page is split across pages
            // rather than drawn off the bottom (which silently lost the overflowing text). A normal row
            // is a single band. Each band redraws every cell's border and its slice of the wrapped
            // lines, so the cell continues cleanly on the next page.
            const fullBand = Math.max(1, Math.floor((this.bottom - this.margin.top - 2 * pad) / lineH));
            const maxLines = Math.max(1, ...placed.map(p => p.lines.length));
            const newCarry = new Map<number, number>();
            let lineOffset = 0;
            while (lineOffset < maxLines) {
                const remaining = maxLines - lineOffset;
                // Lines that fit below the cursor on the current page. If the whole remaining row fits,
                // draw it in one band; otherwise fill what is left of this page (>=1 line) before moving
                // on, so a tall row starting mid-page does not blank the rest of the page. A band is
                // never taller than a page (fitLines <= fullBand), so no line is ever drawn off-page.
                const fitLines = Math.max(0, Math.floor((this.bottom - this.y - 2 * pad) / lineH));
                let bandLines: number;
                if (fitLines >= remaining) bandLines = remaining;               // the rest fits here
                else if (this.y > this.margin.top && fitLines >= 1) bandLines = fitLines; // fill this page
                else if (this.y > this.margin.top) { this.newPage(); continue; } // nothing fits: fresh page
                else bandLines = Math.min(remaining, fullBand);                  // at page top: a full band
                const whole = lineOffset === 0 && bandLines === maxLines;
                const bandH = whole ? rowH : bandLines * lineH + 2 * pad;
                const topY = this.pageH - this.y;
                for (const p of placed) {
                    const x = this.margin.left + p.col * colW;
                    this.page.drawRectangle({ x, y: topY - bandH, width: colW * p.span, height: bandH, borderColor: border, borderWidth: 0.5, color: undefined });
                    const font = p.header ? this.fonts.bold : this.fonts.regular;
                    let ty = topY - pad - size;
                    for (let li = lineOffset; li < lineOffset + bandLines && li < p.lines.length; li++) {
                        this.page.drawText(p.lines[li], { x: x + pad, y: ty, size, font, color: this.lib.rgb(0.12, 0.12, 0.12) }); ty -= lineH;
                    }
                }
                this.y += bandH;
                lineOffset += bandLines;
            }
            // Record this row's rowspans, then age the carries by one row. The carry reserves the
            // spanned columns below so later rows do not slide under a rowspan cell, and the cell's text
            // is placed correctly in its first row. Its border box, though, covers only that first row:
            // painting the continuation border correctly (a single tall box, or edges without internal
            // lines, and correct across a page break) needs a two-pass layout with pre-measured row
            // heights, which is out of proportion to this opt-in engine's fallback role, so a rowspan
            // cell's lower rows are left unbordered rather than mis-bordered.
            for (const p of placed) {
                if (p.rowSpan > 1) for (let c = p.col; c < p.col + p.span; c++) newCarry.set(c, p.rowSpan - 1);
            }
            for (const [c, v] of [...carry]) { if (v > 1) carry.set(c, v - 1); else carry.delete(c); }
            for (const [c, v] of newCarry) carry.set(c, v);
        }
        this.y += 6;
    }
}

/** Minimal shape of `CellMetadata` the native engine reads. */
interface CellMeta { col?: number; colSpan?: number; rowSpan?: number; style?: string; }

/**
 * Node types that own a line of their own inside a table cell. A cell holding these is flattened one
 * line per child; a cell holding only inline content (the usual Markdown/HTML cell) flattens to one.
 */
const BLOCK_CELL_TYPES = new Set<string>(['paragraph', 'heading', 'list', 'code', 'table', 'row', 'admonition', 'note', 'page', 'slide', 'sheet']);

/** Parses a `TextFormatting.size` ("12pt", "14") into points, or null. */
function parseFontSize(size: string | undefined): number | null {
    if (size == null) return null;
    const s = String(size).trim();
    // Relative units resolve against the engine's 11pt default body size (HTML/Markdown sources store a
    // raw CSS `font-size`, so `1.2em` or `120%` must not be read as 1.2pt / 120pt).
    const rel = /^([\d.]+)\s*(em|rem|%)$/i.exec(s);
    if (rel) {
        const n = parseFloat(rel[1]);
        const pt = (rel[2].toLowerCase() === '%' ? n / 100 : n) * 11;
        return Number.isFinite(pt) && pt > 0 ? pt : null;
    }
    // An explicit absolute unit goes through the shared converter (px->pt at 96dpi, in/cm/mm too); a
    // bare number stays points, matching what the parsers write into `formatting.size`.
    if (/(pt|px|in|cm|mm)\s*$/i.test(s)) {
        const pt = lengthToPt(s);
        return pt != null && pt > 0 ? pt : null;
    }
    const n = parseFloat(s);
    return Number.isFinite(n) && n > 0 ? n : null;
}

/**
 * Splits a single token into chunks each no wider than `width`, at character boundaries. Used so a
 * word longer than the whole line (a long URL, hash or base64 blob) wraps instead of overflowing the
 * page. Returns the token unchanged when it already fits.
 */
function splitToWidth(text: string, font: any, size: number, width: number): string[] {
    text = toWinAnsi(text).text; // defensive: callers enc first, but never measure raw (pdf-lib throws)
    if (width <= 0 || font.widthOfTextAtSize(text, size) <= width) return [text];
    const chunks: string[] = [];
    let cur = '';
    for (const ch of text) {
        if (cur && font.widthOfTextAtSize(cur + ch, size) > width) { chunks.push(cur); cur = ch; }
        else cur += ch;
    }
    if (cur) chunks.push(cur);
    return chunks.length ? chunks : [text];
}

/**
 * Greedy word-wrap of plain, UN-ENCODED text to a pixel width, returning the wrapped lines. The text
 * is split into source lines first and encoded one line at a time: `toWinAnsi` maps every newline to
 * a space (a drawn string cannot contain one), so encoding up front would collapse a multi-line cell
 * onto a single line. `enc` defaults to the bare mapping for callers with no warning channel.
 */
function wrapPlain(text: string, font: any, size: number, width: number, enc: (s: string) => string = t => toWinAnsi(t).text): string[] {
    const out: string[] = [];
    for (const raw of String(text ?? '').split(/\r\n|\r|\n/)) {
        const para = enc(raw);
        const words = para.split(/\s+/).filter(Boolean);
        let line = '';
        for (const w of words) {
            // A single word wider than the whole column (a URL, hash, long identifier) is broken at
            // character boundaries so it wraps instead of overflowing into the next cell or off the page.
            const pieces = width > 0 && font.widthOfTextAtSize(w, size) > width ? splitToWidth(w, font, size, width) : [w];
            for (const piece of pieces) {
                const cand = line ? line + ' ' + piece : piece;
                if (line && font.widthOfTextAtSize(cand, size) > width) { out.push(line); line = piece; }
                else line = cand;
            }
        }
        out.push(line);
    }
    return out.length ? out : [''];
}

/** Decodes a base64 string to bytes in both Node and the browser. */
function base64ToBytes(b64: string): Uint8Array {
    if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(b64, 'base64'));
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
}

/**
 * Writes the document metadata (title, author, dates, keywords) onto the pdf-lib document. `m` is the
 * generator's already-resolved `effectiveMetadata` (overrides merged), so this does not re-merge.
 */
function applyMetadata(pdf: any, m: OfficeMetadata): void {
    const set = (fn: string, v: any) => { try { if (v != null && typeof pdf[fn] === 'function') pdf[fn](v); } catch { /* best effort */ } };
    const asDate = (v: unknown): Date | null => {
        if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
        if (typeof v === 'string' && v) { const d = new Date(v); return isNaN(d.getTime()) ? null : d; }
        return null;
    };
    set('setTitle', m.title);
    set('setAuthor', m.author);
    set('setSubject', m.subject);
    // The document language, as the HTML engine's `<html lang>` gives it (`/Lang`).
    if (m.language || m.nativeProperties?.language) set('setLanguage', documentLanguage(m));
    if (m.keywords) set('setKeywords', String(m.keywords).split(/[,;]\s*/).filter(Boolean));
    set('setCreator', 'officeParser (native engine)');
    set('setProducer', 'officeParser (pdf-lib)');
    const created = asDate(m.created); if (created) set('setCreationDate', created);
    const modified = asDate(m.modified); if (modified) set('setModificationDate', modified);
}

/**
 * Renders an AST to a PDF using the native pdf-lib layout engine.
 *
 * @returns the PDF as bytes.
 */
export async function renderNativePdf(ast: OfficeParserAST, config: FullGeneratorConfig, metadata: OfficeMetadata, reportWarning: (type: OfficeWarningType, info?: any) => void): Promise<Uint8Array> {
    const lib = await loadPdfLib(ast.config);
    // updateMetadata:false stops pdf-lib stamping CreationDate/ModDate with the current time: dates
    // come only from the document's own metadata (via applyMetadata), so a source with no dates
    // produces byte-identical output on every run, matching the reproducible DOCX/ODT/EPUB generators.
    const pdf = await lib.PDFDocument.create({ updateMetadata: false });
    applyMetadata(pdf, metadata);

    const fonts: Fonts = {
        regular: await pdf.embedFont(lib.StandardFonts.Helvetica),
        bold: await pdf.embedFont(lib.StandardFonts.HelveticaBold),
        italic: await pdf.embedFont(lib.StandardFonts.HelveticaOblique),
        boldItalic: await pdf.embedFont(lib.StandardFonts.HelveticaBoldOblique),
        mono: await pdf.embedFont(lib.StandardFonts.Courier),
    };

    const pc = config.pdfConfig;
    const paper = paperSizePt(pc.format);
    // Orient the FORMAT by swapping its dimensions when landscape (the same `landscape ? h : w` the
    // DOCX/ODT generators and Puppeteer use), so a paper that is already landscape - `ledger` (17x11) -
    // rotates to portrait under `landscape: true` instead of the old conditional swap that never
    // rotated a landscape sheet. Explicit width/height then override, and (like Puppeteer) are taken as
    // given rather than re-swapped by the flag.
    let w = pc.landscape ? paper.h : paper.w;
    let h = pc.landscape ? paper.w : paper.h;
    if (pc.width) w = toPoints(pc.width, w);
    if (pc.height) h = toPoints(pc.height, h);

    // Margin sides default to the '' unset sentinel; toPoints maps that (and any unparseable value) to
    // mdef so text is not glued to the sheet edge, while an explicit 0 parses to 0 and is honoured for
    // callers who genuinely want an edge-to-edge native PDF.
    const mdef = 48; // ~0.67in
    const mv = (v: string | number | undefined) => toPoints(v, mdef);
    const margin = { top: mv(pc.margin?.top), right: mv(pc.margin?.right), bottom: mv(pc.margin?.bottom), left: mv(pc.margin?.left) };

    const layout = new NativeLayout(pdf, lib, fonts, w, h, margin, config, ast, reportWarning);
    // The native engine reflows a single stream and has no running page furniture, so master-page
    // headers/footers (ast.auxiliary, present only when parsing kept them) are drawn once - headers
    // before the body, footers after - rather than dropped. Footnote/endnote bodies gathered during the
    // walk are drawn last.
    for (const n of ast.auxiliary?.headers || []) await layout.render(n);
    // Gather footnote/endnote bodies from the whole content tree first (they attach to text runs, deep
    // in the tree, so a render()-level sweep would miss them), then render the body, then draw them.
    for (const node of ast.content) layout.collectNotes(node);
    for (const node of ast.content) await layout.render(node, true);
    for (const n of ast.auxiliary?.footers || []) await layout.render(n);
    await layout.flushNotes();

    return await pdf.save();
}

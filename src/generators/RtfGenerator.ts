import { CodeMetadata, ConversionResult, GeneratorConfig, HeadingMetadata, ListMetadata, OfficeContentNode, OfficeParserAST, OfficeWarningType, TextMetadata } from '../types.js';
import { escapeRtf as escapeRtfShared, sanitizeRtfUrl } from '../utils/sanitize.js';
import { BaseGenerator } from './BaseGenerator.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { isSourceComment } from '../utils/commentUtils.js';

/**
 * Generates high-fidelity RTF (Rich Text Format) from an AST.
 */
export class RtfGenerator extends BaseGenerator<'rtf'> {
    private colorTable: string[] = [];
    /** Each colour's index in {@link colorTable}, so a run finds its colour without scanning the table. */
    private colorIndex = new Map<string, number>();
    private inTable = false;
    /**
     * Set while rendering a heading's children.
     *
     * A heading emits its own `{\\b\\fs44 ...}` wrapper, so a run inside it that also carries bold
     * and a size - which is now the normal case for ODF, where the heading's paragraph style is
     * inherited by its runs - would emit a nested `\\fs28` that *overrides* the outer `\\fs44`.
     * The heading then renders at the body-text size it was styled with rather than at heading
     * size. Suppressing the inherited weight and size inside a heading keeps the heading's own
     * wrapper authoritative; every other property (colour, font) still comes through.
     */
    private inHeading = false;
    /** As `inHeading`, but for the inherited font size - see `hasUniformFormatting`. */
    private headingUniformSize = false;
    /** Whether a code block or equation was written, in the monospace font (`\\f2`) the font table then lists. */
    private usedMonospace = false;
    /** Whether the math-as-source warning was given (once per document). */
    private mathWarned = false;

    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'rtf'>) {
        super('rtf', ast, config);
    }

    async generate(): Promise<ConversionResult<'rtf'>> {
        this.colorTable = [];
        this.colorIndex = new Map();
        this.usedMonospace = false;
        this.mathWarned = false;

        // We first process all nodes to collect colors and analyze structure
        const bodyContent = await this.renderBody(this.ast);

        let output = '{\\rtf1\\ansi\\uc1\\deff0\n';

        // 1. Info Group (Metadata)
        const meta = this.effectiveMetadata;
        if (this.config.renderMetadata && meta) {
            // RTF's \info group is a fixed set of control words with no slot for caller-defined keys.
            this.warnUnrepresentableCustomMetadata('RTF');
            output += '{\\info';
            if (meta.title) output += `{\\title ${this.escapeRtf(meta.title)}}`;
            if (meta.author) output += `{\\author ${this.escapeRtf(meta.author)}}`;
            if (meta.description) output += `{\\comm ${this.escapeRtf(meta.description)}}`;
            // \subject and \keywords are standard \info destinations, so unlike a caller's
            // arbitrary custom keys these DO have a home in RTF.
            if (meta.subject) output += `{\\subject ${this.escapeRtf(meta.subject)}}`;
            if (meta.keywords) output += `{\\keywords ${this.escapeRtf(meta.keywords)}}`;
            output += '}\n';
        }

        // 2. Font Table
        output += `{\\fonttbl{\\f0\\fnil\\fcharset0 Arial;}{\\f1\\fnil\\fcharset0 Times New Roman;}${this.usedMonospace ? '{\\f2\\fmodern\\fcharset0 Courier New;}' : ''}}\n`;

        // 3. Color Table
        if (this.colorTable.length > 0) {
            output += '{\\colortbl;';
            for (const hex of this.colorTable) {
                const r = parseInt(hex.substring(1, 3), 16);
                const g = parseInt(hex.substring(3, 5), 16);
                const b = parseInt(hex.substring(5, 7), 16);
                output += `\\red${r}\\green${g}\\blue${b};`;
            }
            output += '}\n';
        }

        // 4. Body
        output += '\\f0\\fs24\n';
        output += bodyContent;
        output += '}';

        return {
            value: output,
            messages: this.messages
        };
    }

    protected override async processNodeRecursive(node: OfficeContentNode, processor: (node: OfficeContentNode, childrenOutput: string) => Promise<string>): Promise<string> {
        // Mirrors the check in BaseGenerator.processNodeRecursive. This override replaces that
        // method entirely, so without repeating the check here the signal would be silently
        // inert for this generator - which is exactly how it was missed.
        checkAbortSignal(this.config.abortSignal);
        const wasInTable = this.inTable;
        if (node.type === 'table') this.inTable = true;
        const wasInHeading = this.inHeading;
        const wasHeadingSize = this.headingUniformSize;
        if (node.type === 'heading') {
            this.inHeading = this.hasUniformFormatting(node, f => f?.bold === true);
            this.headingUniformSize = this.hasUniformFormatting(node, f => !!f?.size);
        }
        const result = await super.processNodeRecursive(node, processor);
        this.inTable = wasInTable;
        this.inHeading = wasInHeading;
        this.headingUniformSize = wasHeadingSize;
        return result;
    }

    private async renderBody(ast: OfficeParserAST): Promise<string> {
        let body = '';
        this.inTable = false;

        const processor = async (node: OfficeContentNode, childrenOutput: string): Promise<string> => {
            const mapping = this.getSemanticMapping(node);
            if (mapping) {
                if (mapping.tag === 'blockquote') {
                    const pPr = this.inTable ? '\\pard\\intbl' : '\\pard';
                    return `${pPr}\\li720\\ri720\\sa120 ${childrenOutput}\\par\n`;
                }

                const hMatch = mapping.tag.match(/^h([1-6])$/);
                if (hMatch) {
                    const level = parseInt(hMatch[1]);
                    const fontSize = 24 + (6 - level) * 4;
                    const pPr = this.inTable ? '\\pard\\intbl' : '\\pard';
                    return `${pPr}\\s${level}\\sb240\\sa120{\\b\\fs${fontSize} ${childrenOutput}}\\par\n`;
                }
            }

            switch (node.type) {
                case 'text': {
                    let text = this.escapeRtf(node.text || '');
                    const f = node.formatting;
                    const meta = node.metadata as TextMetadata;

                    if (this.config.includeFormatting && f) {
                        let prefix = '';
                        let suffix = '';
                        if (f.bold && !this.inHeading) { prefix += '\\b '; suffix = '\\b0 ' + suffix; }
                        if (f.italic) { prefix += '\\i '; suffix = '\\i0 ' + suffix; }
                        if (f.underline) { prefix += '\\ul '; suffix = '\\ul0 ' + suffix; }
                        if (f.strikethrough) { prefix += '\\strike '; suffix = '\\strike0 ' + suffix; }

                        if (f.color) {
                            // RTF default background is white. Ensure light text without a dark background remains readable.
                            const isTextLight = this.isLightColor(f.color);
                            const isBgLight = !f.backgroundColor || this.isLightColor(f.backgroundColor);
                            
                            if (!(isTextLight && isBgLight)) {
                                const idx = this.getColorIndex(f.color);
                                prefix += `\\cf${idx + 1} `;
                            }
                        }
                        if (f.backgroundColor) {
                            const idx = this.getColorIndex(f.backgroundColor);
                            prefix += `\\highlight${idx + 1} `;
                        }
                        if (f.size && !this.headingUniformSize) {
                            let pt = 12; // default
                            const val = parseFloat(f.size);
                            if (!isNaN(val)) {
                                if (f.size.includes('in')) pt = val * 72;
                                else if (f.size.includes('cm')) pt = val * 28.3465;
                                else if (f.size.includes('mm')) pt = val * 2.83465;
                                else if (f.size.includes('px')) pt = val * 0.75;
                                else pt = val;
                            }
                            prefix += `\\fs${Math.round(pt * 2)} `;
                        }

                        text = `{\\f0 ${prefix}${text}${suffix}}`;
                    }

                    if (meta?.link) {
                        const isInternal = meta.linkType !== 'external';
                        if (!this.config.ignoreInternalLinks || !isInternal) {
                            // Scheme-checked, not merely escaped. escapeRtf neutralizes the field
                            // metacharacters but says nothing about where the link points, so RTF
                            // was the one generator that would emit `javascript:` or a `file://`
                            // /UNC target that HTML and Markdown both reject. On rejection, fall
                            // through to the bare link text - the same degradation as HTML's
                            // href="" and Markdown's [text]().
                            const safeLink = sanitizeRtfUrl(meta.link);
                            if (safeLink) {
                                return `{\\field{\\*\\fldinst{HYPERLINK "${safeLink}"}}{\\fldrslt ${text}}}`;
                            }
                        }
                    }

                    return text;
                }

                case 'heading': {
                    const meta = node.metadata as HeadingMetadata;
                    const level = meta?.level || 1;
                    const fontSize = 24 + (6 - level) * 4;
                    const pPr = this.inTable ? '\\pard\\intbl' : '\\pard';
                    return `${pPr}\\s${level}\\sb240\\sa120{\\b\\fs${fontSize} ${childrenOutput}}\\par\n`;
                }

                case 'paragraph': {
                    let pPr = this.inTable ? '\\pard\\intbl' : '\\pard';
                    pPr += '\\sa120';
                    if (this.config.includeFormatting && node.metadata) {
                        const meta = node.metadata as any;
                        if (meta.alignment) {
                            if (meta.alignment === 'center') pPr += '\\qc';
                            else if (meta.alignment === 'right') pPr += '\\qr';
                            else if (meta.alignment === 'justify') pPr += '\\qj';
                        }
                    }
                    return `${pPr} ${childrenOutput}\\par\n`;
                }

                case 'list': {
                    const meta = node.metadata as ListMetadata;
                    const level = meta?.indentation || 0;
                    const indent = (level + 1) * 360;
                    const isOrdered = meta?.listType === 'ordered';
                    const marker = isOrdered ? `${(meta.itemIndex ?? 0) + 1}. ` : '\\bullet ';
                    const listControl = isOrdered ? '\\pndec' : '\\pnbullet';
                    const pPr = this.inTable ? '\\pard\\intbl' : '\\pard';
                    return `${pPr}\\li${indent}\\fi-360\\ilvl${level}${listControl} ${marker}${childrenOutput}\\par\n`;
                }

                case 'table': {
                    return `\\pard\\sa0\n${childrenOutput}`;
                }

                case 'row': {
                    const cells = node.children || [];
                    const pageWidth = 9000; // Standard twips width
                    const cellWidth = Math.floor(pageWidth / (cells.length || 1));
                    let cellDefs = '';
                    for (let i = 0; i < cells.length; i++) {
                        // Add basic cell borders and calculate width
                        cellDefs += `\\clbrdrt\\brdrs\\brdrw10\\clbrdrl\\brdrs\\brdrw10\\clbrdrb\\brdrs\\brdrw10\\clbrdrr\\brdrs\\brdrw10\\cellx${(i + 1) * cellWidth}`;
                    }
                    return `\\trowd\\trgaph108\\trleft-108${cellDefs}\n${childrenOutput}\\row\n`;
                }

                case 'cell': {
                    return `\\pard\\intbl\\sb60\\sa60 ${childrenOutput}\\cell\n`;
                }

                case 'image': {
                    const mode = this.imageMode();
                    if (mode === 'none') return '';
                    const meta = node.metadata as any;
                    const ocr = (node.text || '').trim();
                    // Recognized text from a scanned image is column-aligned across several lines (the
                    // OCR reader keeps that layout by default). A raw newline is just whitespace to an
                    // RTF reader, so every line break has to become a \line or the whole block
                    // collapses into one run. escapeRtf leaves newlines untouched, so split after it.
                    const ocrRtf = ocr ? `${this.escapeRtf(ocr).split(/\r\n|\r|\n/).join('\\line ')}\\par\n` : '';
                    // ocr-text-only: just the recognized text.
                    if (mode === 'ocr-text-only') return ocrRtf;

                    let pict = '';
                    const attachment = this.getAttachment(meta?.attachmentName);
                    if (attachment && attachment.data) {
                        const type = attachment.extension === 'png' ? 'pngblip' : 'jpegblip';
                        // Convert base64 to hex
                        const binary = atob(attachment.data);
                        let hex = '';
                        for (let i = 0; i < binary.length; i++) {
                            const h = binary.charCodeAt(i).toString(16);
                            hex += h.length === 1 ? '0' + h : h;
                            if (i % 64 === 63) hex += '\n'; // Add newlines for better RTF readability
                        }
                        // Default goals (approx 3 inches wide at 1440 twips per inch)
                        pict = `{\\pict\\${type}\\picwgoal4320\\pichgoal3240\n${hex}\n}\n`;
                        // A picture that is a link is the result of a HYPERLINK field, as linked text is.
                        if (meta?.link && (!this.config.ignoreInternalLinks || meta.linkType === 'external')) {
                            const safeLink = sanitizeRtfUrl(meta.link);
                            if (safeLink) pict = `{\\field{\\*\\fldinst{HYPERLINK "${safeLink}"}}{\\fldrslt ${pict}}}`;
                        }
                    }
                    // image+ocr-text: the image, then its recognized text.
                    return mode === 'image+ocr-text' ? pict + ocrRtf : pict;
                }

                case 'break': {
                    return (node.metadata as any)?.breakType === 'page' ? '\\page\n' : '\\line\n';
                }

                case 'code': {
                    // A code block, or an equation (as its LaTeX source: RTF has no math of its own), in
                    // the monospace font, its lines kept; inline math sits in its line. Both were dropped.
                    const meta = node.metadata as CodeMetadata | undefined;
                    if (meta?.math && !this.mathWarned) {
                        this.mathWarned = true;
                        this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'rtf', feature: 'math' });
                    }
                    this.usedMonospace = true;
                    const lines = this.escapeRtf(node.text || this.getNodeText(node)).split(/\r\n|\r|\n/).join('\\line ');
                    if (meta?.math === 'inline') return `{\\f2 ${lines}}`;
                    const pPr = this.inTable ? '\\pard\\intbl' : '\\pard';
                    return `${pPr}\\sa120{\\f2 ${lines}}\\par\n`;
                }

                case 'definitionTerm':
                case 'definitionDescription': {
                    // A term, bold, and its description, indented, each a paragraph of its own (they ran
                    // into each other and into the next paragraph); one holding paragraphs is them.
                    if (node.children?.some(child => child.type !== 'text' && child.type !== 'break' && child.type !== 'image' && !(child.type === 'code' && (child.metadata as CodeMetadata | undefined)?.math === 'inline'))) return childrenOutput;
                    const pPr = this.inTable ? '\\pard\\intbl' : '\\pard';
                    return node.type === 'definitionTerm'
                        ? `${pPr}\\sa0{\\b ${childrenOutput}}\\par\n`
                        : `${pPr}\\li720\\sa120 ${childrenOutput}\\par\n`;
                }

                case 'comment': {
                    // A comment with only text (a CSV comment line) is a paragraph of it; one with
                    // children (a review comment) is written through them. A source comment (`<!-- -->`,
                    // the author's hidden note) has no place in RTF, and is left out as before.
                    if (isSourceComment(node)) return '';
                    if (node.children?.length || !node.text) return childrenOutput;
                    const pPr = this.inTable ? '\\pard\\intbl' : '\\pard';
                    return `${pPr}\\sa120 ${this.escapeRtf(node.text)}\\par\n`;
                }

                case 'embed': {
                    // RTF has no embed concept - degrade to the URL as plain text rather than
                    // silently dropping the node (it has no children to fall back to).
                    const meta = node.metadata as any;
                    if (!meta?.url) return '';
                    // Rendered as visible text rather than a field, but still a URL a reader may
                    // copy, so it gets the same scheme policy.
                    const safeUrl = sanitizeRtfUrl(meta.url);
                    if (!safeUrl) return '';
                    const pPr = this.inTable ? '\\pard\\intbl' : '\\pard';
                    return `${pPr}\\sa120 ${safeUrl}\\par\n`;
                }

                // Rendered through their children only (RTF has no native form for these). Listed
                // explicitly, with no `default`, so that under `noImplicitReturns` a new
                // OfficeContentNodeType fails to compile until it is classified here.
                case 'chart':
                case 'drawing':
                case 'slide':
                case 'note':
                case 'sheet':
                case 'page':
                case 'header':
                case 'footer':
                case 'slideMaster':
                case 'admonition':
                case 'definitionList':
                    return childrenOutput;
            }
        };

        for (const node of ast.content) {
            body += await this.processNodeRecursive(node, processor);
        }
        if (this.collectedNotes.length > 0) {
            body += '\\pard\\sb120\\sa120\\keepn{\\b Notes:}\\par\n';
            for (const note of this.collectedNotes) {
                body += await this.processNodeRecursive(note, processor);
            }
        }
        return body;
    }

    private getColorIndex(hex: string): number {
        const h = hex.toUpperCase();
        let idx = this.colorIndex.get(h);
        if (idx === undefined) {
            idx = this.colorTable.length;
            this.colorTable.push(h);
            this.colorIndex.set(h, idx);
        }
        return idx;
    }

    private isLightColor(hex: string): boolean {
        if (!hex || hex.length !== 7 || !hex.startsWith('#')) return false;
        const r = parseInt(hex.substring(1, 3), 16);
        const g = parseInt(hex.substring(3, 5), 16);
        const b = parseInt(hex.substring(5, 7), 16);
        if (isNaN(r) || isNaN(g) || isNaN(b)) return false;
        
        // Simple luminance calculation
        const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
        return luminance > 0.8;
    }

    private escapeRtf(text: string): string {
        return escapeRtfShared(text);
    }
}

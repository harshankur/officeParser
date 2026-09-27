/**
 * OpenDocument Format (ODF) Parser
 * 
 * **ODF Overview:**
 * ODF is an open standard for office documents (ISO/IEC 26300).
 * Used by LibreOffice, OpenOffice, and other applications.
 * 
 * **File Structure:**
 * ODF files are ZIP archives containing:
 * - `mimetype` - File type identification
 * - `content.xml` -  Main document content
 * - `styles.xml` - Style definitions
 * - `meta.xml` - Document metadata
 * - `Pictures/*` - Embedded images
 * 
 * **Supported Formats:**
 * - ODT: Text documents (application/vnd.oasis.opendocument.text)
 * - ODP: Presentations (application/vnd.oasis.opendocument.presentation)
 * - ODS: Spreadsheets (application/vnd.oasis.opendocument.spreadsheet)
 * 
 * @module OpenOfficeParser
 */

import { BreakMetadata, CellMetadata, ChartData, ChartMetadata, CodeMetadata, FullOfficeParserConfig, HeadingMetadata, ImageMetadata, ListMetadata, OfficeAttachment, OfficeAuxiliaryContent, OfficeContentNode, OfficeParserAST, OfficeParserConfig, OfficeWarningType, PageMetadata, SheetMetadata, SlideMetadata, SupportedFileType, TextFormatting, TextMetadata } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { extractChartData } from '../utils/chartUtils.js';
import { checkAbortSignal, logWarning } from '../utils/errorUtils.js';
import { mathmlToLatex } from '../utils/mathUtils.js';
import { cellSpan, clampRepeat, MAX_COL_SPAN, MAX_ROW_SPAN } from '../utils/numberUtils.js';
import { TextBuilder } from '../utils/textUtils.js';
import { lookupTable, plainRecord } from '../utils/lookupUtils.js';

/**
 * Tracks how many table cells a single document has been allowed to materialize.
 *
 * ODF encodes runs of identical cells/rows as `table:number-columns-repeated` and
 * `table:number-rows-repeated` rather than repeating markup, so a few hundred bytes of XML can ask
 * for an arbitrary number of nodes - and the two multiply, so a row repeat times a column repeat
 * compounds it. The ZIP limits cannot catch this: the XML is tiny before decompression and the
 * expansion happens afterwards, while building the AST.
 *
 * The budget bounds what gets *materialized*, never the attribute itself. Capping the attribute
 * would break ordinary documents - LibreOffice routinely writes `number-rows-repeated="1048566"`
 * to mean "the rest of the sheet is empty", and those runs are legitimate.
 *
 * Warns once per document rather than per clamp, so a wide sheet doesn't emit thousands of
 * identical warnings.
 */
class CellBudget {
    private remaining: number;
    private warned = false;
    constructor(private limit: number, private config: OfficeParserConfig) {
        this.remaining = limit;
    }
    /** How many of `wanted` may be created; 0 once exhausted. */
    take(wanted: number): number {
        // `!(wanted > 0)` rather than `wanted <= 0` so a NaN is rejected too: `NaN <= 0` is
        // false, so a garbage repeat attribute (`parseInt("abc")`) would otherwise fall through
        // and drain the entire remaining budget, dropping every legitimate cell that followed.
        if (!(wanted > 0)) return 0;
        if (this.remaining <= 0) { this.warn(); return 0; }
        if (wanted <= this.remaining) { this.remaining -= wanted; return wanted; }
        const granted = this.remaining;
        this.remaining = 0;
        this.warn();
        return granted;
    }
    get exhausted(): boolean { return this.remaining <= 0; }
    private warn(): void {
        if (this.warned) return;
        this.warned = true;
        logWarning(OfficeWarningType.TABLE_CELL_LIMIT_EXCEEDED, this.config, this.limit);
    }
}

/** Resolves the configured cell budget, falling back to the documented default. */
const createCellBudget = (config: OfficeParserConfig): CellBudget =>
    new CellBudget(config.decompressionLimits?.maxTableCells ?? 1000000, config);

/**
 * Coerces a `table:number-*-repeated` attribute to a usable repeat count. A missing, zero,
 * negative or non-numeric value becomes 1 (the element renders once), so a garbage attribute
 * can neither drop a cell nor poison arithmetic downstream (`colIndex += NaN`).
 */
/**
 * Paragraph-style properties that outlive the style table: everything here is read once from
 * `style:paragraph-properties` and then applied wherever the style is referenced.
 */
interface ParagraphStyleInfo {
    alignment?: 'left' | 'center' | 'right' | 'justify';
    dropCap?: boolean;
    /** `fo:break-before` - a break emitted before the paragraph that carries the style. */
    breakBefore?: 'page' | 'column';
    /** `fo:break-after` - a break emitted after it. */
    breakAfter?: 'page' | 'column';
}

/**
 * Concatenates a node's descendant text, skipping the subtrees of the given tags. Used for the
 * paragraph text fallback so a comment (`office:annotation`) or note (`text:note`) body - which is
 * attached separately, or dropped under `ignoreComments`/`ignoreNotes` - never leaks into the
 * paragraph's own text the way raw `node.textContent` would.
 */
function textContentSkipping(node: Node, skipTags: Set<string>): string {
    let out = '';
    const walk = (n: Node) => {
        if (isElement(n) && skipTags.has((n as Element).tagName)) return;
        if (n.nodeType === 3) { out += n.nodeValue || ''; return; }
        const kids = n.childNodes;
        if (kids) for (let i = 0; i < kids.length; i++) walk(kids[i]);
    };
    walk(node);
    return out;
}

/** The parseParagraphContent closure's shape, passed to {@link buildAnnotationComment}. */
type ParseParaFn = (
    node: Element,
    paraStyleMap: Record<string, ParagraphStyleInfo>,
    styleMap: Record<string, TextFormatting>,
    config: OfficeParserConfig,
    sourceXml: string,
) => { text: string; children: OfficeContentNode[] };

/**
 * Builds a `comment` node from an `office:annotation` element (its `dc:creator`/`dc:date` and its
 * `text:p` body). Shared by the inline paragraph handler and the cell-/page-level handlers so a comment
 * attached at any of the three ODF placements is built identically. `parsePara` is the caller's
 * `parseParagraphContent` closure.
 */
function buildAnnotationComment(
    element: Element,
    parsePara: ParseParaFn,
    paraStyleMap: Record<string, ParagraphStyleInfo>,
    styleMap: Record<string, TextFormatting>,
    config: OfficeParserConfig,
    sourceXml: string,
): OfficeContentNode {
    const author = getFirstElementByTagName(element, "dc:creator")?.textContent || undefined;
    const date = getFirstElementByTagName(element, "dc:date")?.textContent || undefined;
    const children: OfficeContentNode[] = [];
    let text = '';
    for (const cp of getElementsByTagName(element, "text:p")) {
        const c = parsePara(cp, paraStyleMap, styleMap, config, sourceXml);
        text += (text ? ' ' : '') + c.text;
        children.push({ type: 'paragraph', text: c.text, children: c.children, metadata: {} });
    }
    return { type: 'comment', text, children, metadata: { ...(author ? { author } : {}), ...(date ? { date } : {}) } };
}

/**
 * Merges a style's formatting over what it inherits, dropping any flag the style explicitly turns
 * off rather than carrying a `false` forward.
 *
 * Generators all test these flags for truthiness, so a retained `false` would render the same - but
 * it would not *compare* the same, and `MarkdownGenerator.optimizeNodes` merges adjacent text nodes
 * only when their formatting objects are equal. Leaving `bold: false` on one node and nothing on
 * its neighbour would silently stop that merge and fragment the output. Same reasoning, and same
 * shape, as `WordParser`'s direct-run-property merge.
 */
const mergeFormatting = (inherited: TextFormatting, override?: TextFormatting): TextFormatting => {
    if (!override) return { ...inherited };
    const merged: TextFormatting = { ...inherited };
    for (const key of Object.keys(override) as (keyof TextFormatting)[]) {
        const value = override[key];
        if (value === false) delete merged[key];
        else if (value !== undefined) (merged as Record<string, unknown>)[key] = value;
    }
    return merged;
};

const toRepeatCount = (attr: string | null): number => {
    const n = parseInt(attr || "1");
    return Number.isFinite(n) && n > 0 ? n : 1;
};
import { createAttachment } from '../utils/imageUtils.js';
import { ocrDuringParse } from '../utils/ocrUtils.js';
import { getDirectChildren, getElementsByTagName, getFirstElementByTagName, getRawContent, isElement, parseOfficeMetadata, parseXmlString } from '../utils/xmlUtils.js';
import { extractFiles, findRequiredPart } from '../utils/zipUtils.js';

/**
 * Helper to clean and extract attachment name from xlink:href or paths.
 * Handles trailing slashes, leading "./", and subdirectories.
 */
const cleanAttachmentName = (href: string): string => {
    if (!href) return '';
    const cleaned = href.replace(/^\.\//, '').replace(/\/$/, '');
    return cleaned.split('/').pop() || '';
};

/** The ODF document types this parser handles, used to validate a caller-supplied file type. */
const ODF_FILE_TYPES: SupportedFileType[] = ['odt', 'odp', 'ods', 'odg'];

/**
 * Parses an OpenOffice document (.odt, .odp, .ods) and extracts content.
 *
 * @param buffer - The ODF file as a Buffer
 * @param config - Parser configuration
 * @returns A promise resolving to the parsed AST
 */
export const parseOpenOffice = async (buffer: Buffer, config: FullOfficeParserConfig): Promise<OfficeParserAST> => {
    // Honour cancellation requests immediately, before extracting the ZIP archive.
    // ODF containers (ODT/ODS/ODP) bundle content.xml, styles.xml, and media files;
    // aborting early avoids needlessly inflating and parsing all of those resources.
    checkAbortSignal(config.abortSignal);

    const contentFileRegex = /content\.xml/;
    const objectContentFileRegex = /Object \d+\/content\.xml/;
    const mediaFileRegex = /(Pictures|media)\/.*/;
    const metaFileRegex = /meta\.xml/;
    const stylesFileRegex = /styles\.xml/;
    const mimetypeFileRegex = /mimetype/;

    const files = await extractFiles(
        buffer,
        x =>
            !!x.match(contentFileRegex) ||
            !!x.match(objectContentFileRegex) ||
            !!x.match(metaFileRegex) ||
            !!x.match(stylesFileRegex) ||
            !!x.match(mimetypeFileRegex) ||
            (!!config.extractAttachments && !!x.match(mediaFileRegex)),
        config.decompressionLimits,
        config
    );

    // 1. Determine File Type
    const mimetypeFile = files.find(f => f.path === 'mimetype');
    // The archive's own mimetype entry is authoritative when present. When it is missing,
    // fall back to the type the caller asked for (or that was derived from the extension)
    // rather than assuming text: guessing 'odt' for a spreadsheet sends the parser down the
    // office:text branch, which finds nothing in an office:spreadsheet body and yields an
    // empty document for a perfectly valid file.
    let fileType: SupportedFileType = ODF_FILE_TYPES.includes(config.fileType as SupportedFileType)
        ? config.fileType as SupportedFileType
        : 'odt';
    if (mimetypeFile) {
        const mime = mimetypeFile.content.toString().trim();
        if (mime.includes('spreadsheet')) fileType = 'ods';
        else if (mime.includes('presentation')) fileType = 'odp';
        else if (mime.includes('graphics')) fileType = 'odg'; // covers .odg and the .otg template mimetype
        else if (mime.includes('text')) fileType = 'odt';
    }

    // The document body is the content.xml at the archive root. The fallback stays anchored
    // and excludes embedded objects: an ODF file can carry Object N/content.xml for a chart
    // or formula, and an unanchored match would promote one of those to the document body
    // when the real one is missing, silently parsing a chart as if it were the whole file.
    const mainContentFile = files.find(f => f.path === 'content.xml')
        || findRequiredPart(files,
            path => /(^|\/)content\.xml$/.test(path) && !objectContentFileRegex.test(path),
            config, { fileType, part: 'content.xml' });
    const stylesFile = files.find(f => f.path === 'styles.xml');
    const stylesDom = stylesFile ? parseXmlString(stylesFile.content.toString()) : undefined;
    const content: OfficeContentNode[] = [];
    const notes: OfficeContentNode[] = [];

    // Style Map: styleName -> TextFormatting
    // Inline style parsing (from content.xml automatic styles)
    // Null-prototype, keyed by the document's own style names (see listCounters below).
    const styleMap: { [key: string]: TextFormatting } = Object.create(null);
    const paragraphStyleMap: { [key: string]: ParagraphStyleInfo } = Object.create(null);
    // Null-prototype: `listId` is the raw document `text:style-name`/`xml:id`, so a plain `{}` lets a
    // `listId="__proto__"` write onto `Object.prototype` two levels down (global pollution from one
    // crafted ODF). With no prototype the `__proto__` key is an ordinary own property.
    const listCounters: { [listId: string]: { [level: string]: number } } = Object.create(null); // Track item index per listId/level
    let currentListId: string | null = null;
    let lastListType: 'ordered' | 'unordered' | null = null;
    let lastListStyle: string | null = null;
    let listIdCounter = 0;
    let lastWasList = false;
    // Definite-assignment asserted: `traverse` is assigned unconditionally below before any use
    // (the recursive body is defined once and the whole parse depends on it).
    let traverse!: (node: Element, targetArray: OfficeContentNode[], forceHeading?: boolean, sourceXml?: string, asSheet?: boolean) => void;

    // Helper to parse styles
    const parseStyles = (scope: Document | Element) => {
        const styles = getElementsByTagName(scope, "style:style");
        for (const style of styles) {
            const name = style.getAttribute("style:name");
            if (!name) continue;

            const styleInfo: ParagraphStyleInfo = {};

            // Parse paragraph properties for alignment and drop caps
            const paraProps = getFirstElementByTagName(style, "style:paragraph-properties");
            if (paraProps) {
                const textAlign = paraProps.getAttribute("fo:text-align");
                if (textAlign) {
                    const alignMap: Record<string, 'left' | 'center' | 'right' | 'justify'> = lookupTable({
                        'start': 'left',
                        'left': 'left',
                        'center': 'center',
                        'end': 'right',
                        'right': 'right',
                        'justify': 'justify'
                    });
                    if (alignMap[textAlign]) {
                        styleInfo.alignment = alignMap[textAlign];
                    }
                }

                // Detect Drop Caps
                const dropCap = getFirstElementByTagName(paraProps, "style:drop-cap");
                if (dropCap) {
                    styleInfo.dropCap = true;
                }

                // Page/column breaks. ODF attaches these to the paragraph style rather than
                // writing an inline element the way DOCX's `<w:br w:type="page"/>` does, which is
                // why `includeBreakNodes` produced nothing at all for ODF: there was no inline
                // element to find. Only the two break kinds that map onto a BreakMetadata type
                // are carried; `auto` and `even-page`/`odd-page` have no equivalent.
                const breakBefore = paraProps.getAttribute("fo:break-before");
                if (breakBefore === 'page' || breakBefore === 'column') styleInfo.breakBefore = breakBefore;
                const breakAfter = paraProps.getAttribute("fo:break-after");
                if (breakAfter === 'page' || breakAfter === 'column') styleInfo.breakAfter = breakAfter;
            }

            if (Object.keys(styleInfo).length > 0) {
                paragraphStyleMap[name] = styleInfo;
            }

            // Parse text properties
            const textProps = getFirstElementByTagName(style, "style:text-properties");
            // Parse table cell properties (for ODS background)
            const cellProps = getFirstElementByTagName(style, "style:table-cell-properties");

            const formatting: TextFormatting = {};

            if (cellProps) {
                const bgColor = cellProps.getAttribute("fo:background-color");
                if (bgColor && bgColor !== 'transparent') formatting.backgroundColor = bgColor;
            }

            if (textProps) {
                // Record the *off* states as an explicit `false`, not as an absent key.
                //
                // Now that a paragraph style's text properties are inherited by the runs inside it,
                // a span has to be able to turn one back off: LibreOffice writes
                // `fo:font-weight="normal"` on the span whenever a user un-bolds part of a
                // bold-styled paragraph. With only the `true` side recorded, that span had nothing
                // to override the inherited value with and came out bold - wrong in the opposite
                // direction from the bug the inheritance fixed. `TextFormatting`'s flags are
                // `boolean | undefined` precisely so "explicitly off" is expressible.
                const fontWeight = textProps.getAttribute("fo:font-weight") || textProps.getAttribute("style:font-weight-asian");
                // Numeric weights are the same axis: 600+ is bold, below that is not.
                if (fontWeight) formatting.bold = fontWeight === "bold" || /^[6-9]00$/.test(fontWeight);
                const fontStyle = textProps.getAttribute("fo:font-style") || textProps.getAttribute("style:font-style-asian");
                if (fontStyle) formatting.italic = fontStyle === "italic" || fontStyle === "oblique";
                const underline = textProps.getAttribute("style:text-underline-style");
                if (underline) formatting.underline = underline !== "none";
                const lineThrough = textProps.getAttribute("style:text-line-through-style");
                if (lineThrough) formatting.strikethrough = lineThrough !== "none";
                const size = textProps.getAttribute("fo:font-size") || textProps.getAttribute("style:font-size-asian");
                if (size) formatting.size = size;
                const color = textProps.getAttribute("fo:color");
                if (color) formatting.color = color;

                // Background color (text level) - override cell level if present?
                const bgColor = textProps.getAttribute("fo:background-color");
                if (bgColor && bgColor !== 'transparent') formatting.backgroundColor = bgColor;

                // Font family
                const fontName = textProps.getAttribute("style:font-name") || textProps.getAttribute("fo:font-family");
                if (fontName) formatting.font = fontName;

                // Subscript/Superscript from text-position (e.g., "sub 58%" or "super 58%")
                const textPosition = textProps.getAttribute("style:text-position");
                if (textPosition) {
                    if (textPosition.startsWith("sub")) formatting.subscript = true;
                    if (textPosition.startsWith("super")) formatting.superscript = true;
                }
            }

            if (Object.keys(formatting).length > 0) styleMap[name] = formatting;
        }
    };

    if (stylesDom) {
        parseStyles(stylesDom);
    }

    /**
     * Helper to parse a paragraph node (text:p or text:h) and extract its content.
     * Returns the paragraph content without creating a content node.
     * 
     * @param node - The paragraph element to parse
     */
    /**
     * Helper to parse inline content (text, spans, links, notes, etc.) recursively.
     * 
     * @param node - The element to parse (paragraph, span, or link)
     * @param styleMap - Map of style names to formatting
     * @param config - Parser configuration
     * @param notes - Optional array to collect footnotes/endnotes
     * @param paragraphStyleMap - Map of style names to alignments and props (needed for notes)
     * @param parentFormatting - Formatting inherited from parent (e.g. span inside span)
     * @param linkMetadata - Metadata inherited from parent link
     * @returns Object containing text and children
     */
    const parseInlineContent = (
        node: Node,
        styleMap: Record<string, TextFormatting>,
        config: OfficeParserConfig,
        notes: OfficeContentNode[],
        paragraphStyleMap: Record<string, ParagraphStyleInfo>,
        parentFormatting: TextFormatting = {},
        linkMetadata?: { link?: string; linkType?: 'internal' | 'external' },
        sourceXml: string = ''
    ): { text: string; children: OfficeContentNode[]; anchorIds: string[] } => {
        const children: OfficeContentNode[] = [];
        const anchorIds: string[] = [];
        let fullText = '';

        if (!node.childNodes) return { text: '', children: [], anchorIds: [] };

        for (let i = 0; i < node.childNodes.length; i++) {
            const child = node.childNodes[i];

            if (child.nodeType === 3) { // Text node
                const text = child.textContent || '';
                if (text) {
                    fullText += text;
                    children.push({
                        type: 'text',
                        text: text,
                        formatting: parentFormatting,
                        metadata: linkMetadata ? { ...linkMetadata } : undefined
                    });
                }
            } else if (isElement(child)) {
                const element = child as Element;
                const tagName = element.tagName;

                if (tagName === 'text:bookmark' || tagName === 'text:bookmark-start') {
                    const name = element.getAttribute('text:name');
                    if (name) anchorIds.push(name);
                } else if (tagName === 'text:s') {
                    // Space. `text:c` is attacker-controlled: clamp it so a tiny `<text:s text:c="5e8"/>`
                    // cannot allocate a multi-GB string (a memory DoS the zip cap does not stop, since the
                    // blow-up is in repeat, after inflation).
                    const count = parseInt(element.getAttribute('text:c') || '1');
                    const spaces = ' '.repeat(clampRepeat(count));
                    fullText += spaces;
                    children.push({
                        type: 'text',
                        text: spaces,
                        formatting: parentFormatting,
                        metadata: linkMetadata ? { ...linkMetadata } : undefined
                    });
                } else if (tagName === 'text:tab') {
                    // Tab
                    fullText += '\t';
                    children.push({
                        type: 'text',
                        text: '\t',
                        formatting: parentFormatting,
                        metadata: linkMetadata ? { ...linkMetadata } : undefined
                    });
                } else if (tagName === 'text:soft-page-break') {
                    // The page boundary the editor recorded at its last save. DOCX's equivalent
                    // is `w:lastRenderedPageBreak`, so it maps onto the same break type rather
                    // than onto 'page', which is reserved for a break the author asked for.
                    if (config.includeBreakNodes) {
                        children.push({ type: 'break', metadata: { breakType: 'lastRenderedPage' } as BreakMetadata });
                    }
                } else if (tagName === 'text:line-break') {
                    // Line break
                    fullText += '\n';
                    children.push({
                        type: 'text',
                        text: '\n',
                        formatting: parentFormatting,
                        metadata: { ...(linkMetadata || {}), isLineBreak: true } as any
                    });
                } else if (tagName === 'text:span') {
                    // Formatted text span
                    const styleName = element.getAttribute("text:style-name");
                    const formatting = styleName ? mergeFormatting(parentFormatting, styleMap[styleName]) : parentFormatting;

                    const spanContent = parseInlineContent(element, styleMap, config, notes, paragraphStyleMap, formatting, linkMetadata, sourceXml);
                    fullText += spanContent.text;
                    children.push(...spanContent.children);
                    anchorIds.push(...spanContent.anchorIds);
                } else if (tagName === 'text:a' || tagName === 'draw:a') {
                    // Hyperlink: around text (text:a), or around a frame, a picture that is a link (draw:a)
                    let href = element.getAttribute('xlink:href') || '';
                    const isInternal = href.startsWith('#');
                    const linkType = isInternal ? 'internal' : 'external';

                    if (isInternal) {
                        // ODT internal links can be encoded and might have suffixes like |outline
                        try {
                            href = decodeURIComponent(href).split('|')[0];
                        } catch (e) {
                            href = href.split('|')[0];
                        }
                        // Normalize internal link: if it contains #, keep only from # onwards
                        if (href.includes('#')) {
                            href = '#' + href.split('#').pop();
                        }
                    }

                    let newLinkMetadata: TextMetadata | undefined;
                    if (!isInternal || !config.ignoreInternalLinks) {
                        newLinkMetadata = { link: href, linkType: linkType as 'internal' | 'external' };
                    }

                    const linkContent = parseInlineContent(element, styleMap, config, notes, paragraphStyleMap, parentFormatting, newLinkMetadata, sourceXml);
                    fullText += linkContent.text;
                    children.push(...linkContent.children);
                    anchorIds.push(...linkContent.anchorIds);
                } else if (tagName === 'text:note' && !config.ignoreNotes) {
                    // Footnote or endnote
                    const noteClass = (element.getAttribute('text:note-class') || 'footnote') as 'footnote' | 'endnote';
                    const noteId = element.getAttribute('text:id') || element.getAttribute('xml:id') || undefined;
                    const noteBody = getFirstElementByTagName(element, "text:note-body");

                    if (noteBody) {
                        // Extract note content recursively
                        const notePs = getElementsByTagName(noteBody, "text:p");
                        const noteChildren: OfficeContentNode[] = [];
                        let noteText = '';

                        for (const np of notePs) {
                            const npContent = parseParagraphContent(np, paragraphStyleMap, styleMap, config, sourceXml);
                            noteText += (noteText ? ' ' : '') + npContent.text;

                            const npNode: OfficeContentNode = {
                                type: 'paragraph',
                                text: npContent.text,
                                children: npContent.children,
                                metadata: {
                                    ...(npContent.alignment ? { alignment: npContent.alignment } : {}),
                                    ...(npContent.anchorIds?.length ? { anchorIds: npContent.anchorIds } : {})
                                }
                            };
                            noteChildren.push(npNode);
                        }

                        const noteNode: OfficeContentNode = {
                            type: 'note',
                            text: noteText,
                            children: noteChildren,
                            metadata: {
                                noteType: noteClass,
                                noteId: noteId
                            }
                        };

                        if (children.length > 0 && children[children.length - 1].type === 'text') {
                            const precedingNode = children[children.length - 1];
                            if (!precedingNode.notes) {
                                precedingNode.notes = [];
                            }
                            precedingNode.notes.push(noteNode);
                        } else {
                            const emptyTextNode: OfficeContentNode = { type: 'text', text: '' };
                            emptyTextNode.notes = [noteNode];
                            children.push(emptyTextNode);
                        }
                    }
                } else if (tagName === 'office:annotation' && !config.ignoreComments) {
                    // A point comment. Mirrors the note branch and WordParser's attachment convention:
                    // attach to the preceding text run's `.comments`, else a leading empty text node.
                    const commentNode = buildAnnotationComment(element, parseParagraphContent, paragraphStyleMap, styleMap, config, sourceXml);
                    if (children.length > 0 && children[children.length - 1].type === 'text') {
                        const preceding = children[children.length - 1];
                        if (!preceding.comments) preceding.comments = [];
                        preceding.comments.push(commentNode);
                    } else {
                        const emptyTextNode: OfficeContentNode = { type: 'text', text: '' };
                        emptyTextNode.comments = [commentNode];
                        children.push(emptyTextNode);
                    }
                } else if (tagName === 'draw:frame') {
                    const frame = element;
                    const drawTextBox = getFirstElementByTagName(frame, "draw:text-box");
                    const drawObject = getFirstElementByTagName(frame, "draw:object");

                    if (drawTextBox) {
                        const textBoxChildren: OfficeContentNode[] = [];
                        traverse(drawTextBox, textBoxChildren, false, sourceXml);
                        children.push(...textBoxChildren);
                        const textBoxText = textBoxChildren.map(c => c.text || '').join('\n');
                        fullText += textBoxText;
                    } else if (drawObject) {
                        const href = drawObject.getAttribute("xlink:href");
                        let isFormula = false;
                        let formulaText = '';
                        let attachmentName = '';
                        if (href) {
                            attachmentName = cleanAttachmentName(href);
                            const objectPath = `${attachmentName}/content.xml`;
                            const objectFile = files.find(f => f.path === objectPath || f.path.endsWith(objectPath));
                            if (objectFile) {
                                const objXml = parseXmlString(objectFile.content.toString());
                                const mathNode = getFirstElementByTagName(objXml, "math");
                                if (mathNode) {
                                    isFormula = true;
                                    formulaText = mathmlToLatex(mathNode).trim();
                                }
                            }
                        }

                        if (isFormula) {
                            fullText += formulaText;
                            // A `code` node carrying `math`, not a plain `text` node: the formula
                            // is LaTeX, and marking it as such is what lets generators render it
                            // as maths rather than emit it as prose that happens to contain
                            // backslashes. Same node shape DOCX, PPTX, HTML and Markdown produce.
                            const formulaNode: OfficeContentNode = {
                                type: 'code',
                                text: formulaText,
                                metadata: { math: 'inline', ...(linkMetadata ?? {}) } as CodeMetadata
                            };
                            if (config.includeRawContent) {
                                formulaNode.rawContent = getRawContent(frame, sourceXml, config);
                            }
                            children.push(formulaNode);
                        } else {
                            // Standard inline image extraction fallback if object is not a formula
                            let altText = '';
                            const svgTitle = getFirstElementByTagName(frame, "svg:title");
                            const svgDesc = getFirstElementByTagName(frame, "svg:desc");
                            if (svgTitle && svgTitle.textContent) {
                                altText = svgTitle.textContent;
                            } else if (svgDesc && svgDesc.textContent) {
                                altText = svgDesc.textContent;
                            }

                            let imageHref = '';
                            const drawImages = getElementsByTagName(frame, "draw:image");
                            if (drawImages.length > 0) {
                                imageHref = drawImages[0].getAttribute("xlink:href") || '';
                                if (imageHref) {
                                    imageHref = cleanAttachmentName(imageHref);
                                }
                            }

                            const imageNode: OfficeContentNode = {
                                type: 'image',
                                text: '',
                                children: [],
                                metadata: {
                                    attachmentName: imageHref || attachmentName,
                                    ...(altText ? { altText } : {}),
                                    ...(linkMetadata?.link ? { link: linkMetadata.link, linkType: linkMetadata.linkType } : {})
                                }
                            };
                            if (config.includeRawContent) {
                                imageNode.rawContent = getRawContent(frame, sourceXml, config);
                            }
                            children.push(imageNode);
                        }
                    } else {
                        // Standard inline image extraction fallback
                        let altText = '';
                        const svgTitle = getFirstElementByTagName(frame, "svg:title");
                        const svgDesc = getFirstElementByTagName(frame, "svg:desc");
                        if (svgTitle && svgTitle.textContent) {
                            altText = svgTitle.textContent;
                        } else if (svgDesc && svgDesc.textContent) {
                            altText = svgDesc.textContent;
                        }

                        let imageHref = '';
                        const drawImages = getElementsByTagName(frame, "draw:image");
                        if (drawImages.length > 0) {
                            imageHref = drawImages[0].getAttribute("xlink:href") || '';
                            if (imageHref) {
                                imageHref = cleanAttachmentName(imageHref);
                            }
                        }

                        const imageNode: OfficeContentNode = {
                            type: 'image',
                            text: '',
                            children: [],
                            metadata: {
                                attachmentName: imageHref,
                                ...(altText ? { altText } : {}),
                                // A picture inside a link (draw:a or text:a) is that link.
                                ...(linkMetadata?.link ? { link: linkMetadata.link, linkType: linkMetadata.linkType } : {})
                            }
                        };
                        if (config.includeRawContent) {
                            imageNode.rawContent = getRawContent(frame, sourceXml, config);
                        }
                        children.push(imageNode);
                    }
                }
            }
        }
        return { text: fullText, children, anchorIds };
    };

    /**
     * Helper to parse a paragraph node (text:p or text:h) and extract its content.
     * Returns the paragraph content without creating a content node.
     * 
     * @param node - The paragraph element to parse
     * @param paraStyleMap - Map of style names to alignments/props
     * @param styleMap - Map of style names to formatting
     * @param config - Parser configuration
     * @returns Object containing text, children, alignment, and style info
     */
    const parseParagraphContent = (
        node: Element,
        paraStyleMap: Record<string, ParagraphStyleInfo>,
        styleMap: Record<string, TextFormatting>,
        config: OfficeParserConfig,
        sourceXml: string
    ): { text: string; children: OfficeContentNode[]; alignment?: 'left' | 'center' | 'right' | 'justify'; style?: string; anchorIds?: string[] } => {
        // Get paragraph style for alignment and drop caps
        const paraStyle = node.getAttribute("text:style-name");
        const styleInfo = paraStyle ? paraStyleMap[paraStyle] : undefined;
        const alignment = styleInfo?.alignment;
        const dropCap = styleInfo?.dropCap;
        const formatting = mergeFormatting({}, paraStyle ? styleMap[paraStyle] : undefined);

        // Parse content recursively using the new helper
        const content = parseInlineContent(node, styleMap, config, notes, paraStyleMap, formatting, undefined, sourceXml);

        // Add style name to metadata of children if they don't have one
        if (paraStyle) {
            content.children.forEach(child => {
                if (child.type === 'text') {
                    if (!child.metadata) child.metadata = {};
                    // Only add style if it's a text node and doesn't have one?
                    // Or just add it.
                    // Cast to any to avoid union type issues for now, or check type
                    const meta = child.metadata as TextMetadata;
                    if (!meta.style) meta.style = paraStyle;
                }
            });
        }

        // Fallback: if no children were created but there's text content. Exclude comment/note subtrees
        // so an annotation-only or note-only paragraph does not leak the comment/note body into the
        // paragraph text (e.g. when ignoreComments/ignoreNotes skipped building the child node).
        if (content.children.length === 0) {
            const fullText = textContentSkipping(node, new Set(['office:annotation', 'text:note']));
            if (fullText.trim()) {
                content.text = fullText;
                content.children.push({
                    type: 'text',
                    text: fullText
                });
            }
        }

        // Handle Drop Cap: Apply large font to first letter if configured
        if (dropCap && content.children.length > 0) {
            const firstChild = content.children[0];
            if (firstChild.type === 'text' && firstChild.text) {
                if (firstChild.text.length === 1) {
                    // Already a single letter, just apply formatting
                    firstChild.formatting = { ...firstChild.formatting, size: '58.5pt' };
                } else {
                    // Split text node
                    const firstChar = firstChild.text[0];
                    const restText = firstChild.text.substring(1);

                    const dropCapNode: OfficeContentNode = {
                        type: 'text',
                        text: firstChar,
                        formatting: { ...firstChild.formatting, size: '58.5pt' },
                        metadata: firstChild.metadata
                    };

                    // Update original node
                    firstChild.text = restText;

                    // Insert drop cap node
                    content.children.unshift(dropCapNode);
                }
            }
        }

        return { text: content.text, children: content.children, alignment, style: paraStyle || undefined, anchorIds: content.anchorIds };
    };

    /**
     * Splits paragraph content into multiple segments based on line breaks.
     * Used to handle soft line breaks within list items.
     * 
     * @param pContent - The content of a single paragraph
     * @returns Array of content segments
     */
    const splitParagraphByBreaks = (pContent: { text: string; children: OfficeContentNode[]; alignment?: "left" | "center" | "right" | "justify"; style?: string }): { text: string; children: OfficeContentNode[] }[] => {
        const segments: { text: string; children: OfficeContentNode[] }[] = [];
        let currentText = "";
        let currentChildren: OfficeContentNode[] = [];

        for (const child of pContent.children) {
            if (child.type === "text" && (child.metadata as any)?.isLineBreak) {
                segments.push({ text: currentText, children: currentChildren });
                currentText = "";
                currentChildren = [];
            } else {
                currentText += child.text || "";
                currentChildren.push(child);
            }
        }
        segments.push({ text: currentText, children: currentChildren });
        return segments;
    };


    /**
     * Helper to parse a table node and extract its structure.
     * Properly creates table → row → cell hierarchy with metadata.
     * 
     * @param tableNode - The table:table element
     * @param paraStyleMap - Map of style names to alignments
     * @param styleMap - Map of style names to formatting
     * @param config - Parser configuration
     * @returns Table content node with proper structure
     */
    const parseTable = (
        tableNode: Element,
        paraStyleMap: Record<string, ParagraphStyleInfo>,
        styleMap: Record<string, TextFormatting>,
        config: OfficeParserConfig,
        sourceXml: string,
        cellBudget: CellBudget
    ): OfficeContentNode => {
        const rows: OfficeContentNode[] = [];
        // Rows sit directly under <table:table>, but a repeating header block is wrapped in
        // <table:table-header-rows> - written by LibreOffice and by this library's own ODT generator.
        // Reading only the direct <table:table-row> children dropped every header row on the floor, so
        // walk the direct children in document order and descend into the wrapper, flagging its rows.
        // (Nested tables are still excluded: only direct children are considered at each level.)
        const tableRows: { row: Element; isHeader: boolean }[] = [];
        for (let i = 0; i < (tableNode.childNodes?.length || 0); i++) {
            const child = tableNode.childNodes[i];
            if (!isElement(child)) continue;
            const element = child as Element;
            if (element.tagName === "table:table-row") tableRows.push({ row: element, isHeader: false });
            else if (element.tagName === "table:table-header-rows") {
                for (const headerRow of getDirectChildren(element, "table:table-row")) tableRows.push({ row: headerRow, isHeader: true });
            }
        }
        let rowIndex = 0;

        for (const { row, isHeader } of tableRows) {
            checkAbortSignal(config.abortSignal);
            const cells: OfficeContentNode[] = [];
            // Use getDirectChildren to avoid nested table cells
            const tableCells = getDirectChildren(row, "table:table-cell");
            const rowsRepeated = toRepeatCount(row.getAttribute("table:number-rows-repeated"));

            let colIndex = 0;

            for (const cell of tableCells) {
                const cellChildren: OfficeContentNode[] = [];
                // Built without reading back what it holds, which would copy all of it per paragraph.
                const cellTextParts = new TextBuilder();
                const colsRepeated = toRepeatCount(cell.getAttribute("table:number-columns-repeated"));
                const colSpan = cellSpan(cell.getAttribute("table:number-columns-spanned"), MAX_COL_SPAN);
                const rowSpan = cellSpan(cell.getAttribute("table:number-rows-spanned"), MAX_ROW_SPAN);

                // Helper to recursively process cell children (handles frames, text-boxes, etc. in ODP)
                const processChildren = (node: Element) => {
                    if (!node.childNodes) return;

                    for (let i = 0; i < node.childNodes.length; i++) {
                        const child = node.childNodes[i];
                        if (isElement(child)) { // Element
                            const element = child as Element;

                            if (element.tagName === "text:p" || element.tagName === "text:h") {
                                const pContent = parseParagraphContent(element, paraStyleMap, styleMap, config, sourceXml);
                                let pNode: OfficeContentNode;
                                if (element.tagName === "text:h") {
                                    pNode = {
                                        type: 'heading',
                                        text: pContent.text,
                                        children: pContent.children,
                                        metadata: {
                                            level: parseInt(element.getAttribute("text:outline-level") || "1"),
                                            ...(pContent.alignment ? { alignment: pContent.alignment } : {}),
                                            ...(pContent.style ? { style: pContent.style } : {})
                                        }
                                    };
                                } else {
                                    pNode = {
                                        type: 'paragraph',
                                        text: pContent.text,
                                        children: pContent.children,
                                        metadata: {
                                            ...(pContent.alignment ? { alignment: pContent.alignment } : {}),
                                            ...(pContent.style ? { style: pContent.style } : {})
                                        }
                                    };
                                }

                                // Clean up metadata if empty
                                if (pNode.type === 'paragraph' && Object.keys(pNode.metadata || {}).length === 0) {
                                    delete pNode.metadata;
                                }

                                if (config.includeRawContent) {
                                    pNode.rawContent = getRawContent(element, sourceXml, config);
                                }

                                cellChildren.push(pNode);
                                cellTextParts.append(pContent.text);
                                // Add newline if there are multiple paragraphs/headings
                                if (!cellTextParts.isEmpty() && !cellTextParts.endsWith('\n')) {
                                    cellTextParts.append('\n');
                                }
                            } else if (element.tagName === "table:table") {
                                // Recursive call for nested table
                                const nestedTableNode = parseTable(element, paraStyleMap, styleMap, config, sourceXml, cellBudget);
                                cellChildren.push(nestedTableNode);
                            } else if (element.tagName === "draw:frame" || element.tagName === "draw:text-box") {
                                // Recursively process container content (common in ODP)
                                processChildren(element);
                            }
                        }
                    }
                };

                processChildren(cell);

                let cellText = cellTextParts.toString();
                // Trim trailing newline from cellText
                if (cellText.endsWith('\n')) {
                    cellText = cellText.slice(0, -1);
                }

                // ODS cell notes are `<office:annotation>` children of the cell itself (outside any
                // text:p, so processChildren never sees them); attach them to the cell node's comments.
                const cellComments = config.ignoreComments
                    ? []
                    : getDirectChildren(cell, "office:annotation").map(a =>
                        buildAnnotationComment(a, parseParagraphContent, paraStyleMap, styleMap, config, sourceXml));

                // Add cell(s) for repeated columns
                // Bounded by the document's cell budget, not by the attribute: the repeat count
                // is attacker-influenced and this path materializes a node per iteration.
                const allowedCols = cellBudget.take(colsRepeated);
                for (let k = 0; k < allowedCols; k++) {
                    // Repeat expansion is the one place a small document produces a long loop,
                    // so it is also the one place a caller most needs to be able to cancel.
                    if ((k & 1023) === 0) checkAbortSignal(config.abortSignal);
                    // Apply cell background color if defined in styleMap
                    const cellStyleName = cell.getAttribute("table:style-name");
                    const cellBgColor = cellStyleName && styleMap[cellStyleName]?.backgroundColor;

                    const cellNode: OfficeContentNode = {
                        type: 'cell',
                        text: cellText,
                        // Repeated columns share the (read-only) child nodes via a shallow array copy
                        // rather than deep-copying them: a huge number-columns-repeated on a content-
                        // bearing cell would otherwise duplicate the whole cell body per column (bytes the
                        // cell-count budget does not charge), a multi-GB amplification and CPU hang.
                        children: cellChildren.length > 0 ? (k === 0 ? cellChildren : cellChildren.slice()) : [],
                        metadata: {
                            row: rowIndex,
                            col: colIndex,
                            ...(cellBgColor ? { backgroundColor: cellBgColor } : {})
                        } as CellMetadata
                    };

                    const cellMetadata = cellNode.metadata as CellMetadata;
                    if (colSpan > 1) cellMetadata.colSpan = colSpan;
                    if (rowSpan > 1) cellMetadata.rowSpan = rowSpan;
                    // A row node cannot carry metadata of its own, so the header flag rides on its
                    // cells - the same `style: 'header'` convention the PDF parser uses for TH cells,
                    // which is what `isHeaderRow` (and therefore every generator) reads.
                    if (isHeader) cellMetadata.style = 'header';

                    if (config.includeRawContent) {
                        cellNode.rawContent = getRawContent(cell, sourceXml, config);
                    }

                    // Share the (read-only) comment nodes across repeated columns via a shallow array
                    // copy, rather than deep-copying: a huge number-columns-repeated on a cell carrying a
                    // large annotation would otherwise duplicate the whole comment per column (bytes the
                    // cell budget does not charge), a multi-GB amplification and CPU hang.
                    if (cellComments.length) {
                        cellNode.comments = k === 0 ? cellComments : cellComments.slice();
                    }

                    cells.push(cellNode);
                    colIndex++;
                }
            }

            // Add row(s) for repeated rows. rows x cols cells are materialized (charged against the
            // budget), but a repeated row shallow-clones each cell (own object + own metadata for the
            // row-index fix below) and SHARES the read-only children/comments/text by reference; JSON
            // deep-copying instead serialized every cell's full content per row, so a row of content-
            // bearing cells built a multi-hundred-MB string (a RangeError / OOM the count budget missed).
            const allowedRows = cells.length === 0
                ? (rowsRepeated > 0 ? 1 + cellBudget.take(rowsRepeated - 1) : 0)
                : Math.min(rowsRepeated,
                    1 + Math.floor(cellBudget.take(Math.max(0, (rowsRepeated - 1) * cells.length)) / cells.length));
            for (let k = 0; k < allowedRows; k++) {
                if ((k & 255) === 0) checkAbortSignal(config.abortSignal);
                const rowNode: OfficeContentNode = {
                    type: 'row',
                    children: k === 0 ? cells : cells.map(c => ({ ...c, metadata: c.metadata ? { ...c.metadata } : c.metadata, children: c.children ? c.children.slice() : c.children, comments: c.comments ? c.comments.slice() : c.comments }) as OfficeContentNode)
                };

                // Fix row indices for repeated rows
                if (k > 0) {
                    rowNode.children?.forEach(c => {
                        if (c.metadata && 'row' in c.metadata) {
                            (c.metadata as CellMetadata).row = rowIndex;
                        }
                    });
                }

                if (config.includeRawContent) {
                    rowNode.rawContent = getRawContent(row, sourceXml, config);
                }

                rows.push(rowNode);
                rowIndex++;
            }
        }

        return {
            type: 'table',
            children: rows
        };
    };


    const parseContentXml = (xmlString: string) => {
        const xml = parseXmlString(xmlString, { locator: config.includeRawContent });
        const body = getFirstElementByTagName(xml, "office:body");
        if (!body) return;
        // One budget for the entire document. It has to span every table - spreadsheet sheets,
        // ODT/ODP body tables, and nested tables alike - or a file sidesteps the cap simply by
        // splitting a huge repeat expansion across many small tables. `traverse` and the
        // spreadsheet branch below both close over this; `parseTable` receives it explicitly.
        const cellBudget = createCellBudget(config);
        // Automatic styles are local to content.xml, but their definitions have exactly the
        // shape styles.xml uses, so they go through the same reader rather than a second copy of
        // it - the copy is how `fo:break-before` came to be read in neither place.
        const automaticStyles = getFirstElementByTagName(xml, "office:automatic-styles");
        if (automaticStyles) {
            parseStyles(automaticStyles);
        }

        // Start traversal
        const officeBody = getFirstElementByTagName(xml, "office:body");
        if (officeBody) {
            const bodyContent = getDirectChildren(officeBody, "office:text")[0] ||
                getDirectChildren(officeBody, "office:presentation")[0] ||
                getDirectChildren(officeBody, "office:spreadsheet")[0];
            if (bodyContent) {
                const bodyChildren = getDirectChildren(bodyContent, "*");
                const isSpreadsheet = bodyContent.tagName === "office:spreadsheet";
                for (const child of bodyChildren) {
                    traverse(child, content, false, xmlString, isSpreadsheet);
                }
            }
        }

        /**
         * Recursively traverses a node and its children to extract content.
         * Properly handles paragraphs, headings, tables, lists, and frames.
         * 
         * @param node - The element to traverse
         * @param targetArray - The array to push extracted content nodes to
         * @param forceHeading - If true, treats all paragraphs as headings (used for slide titles)
         * @param sourceXml - The source XML string for raw content extraction
         * @param asSheet - If true, treats tables as sheets (for ODS)
         */
        /**
         * Emits the break a paragraph style asks for, on the given side of that paragraph.
         *
         * ODF has no inline break element for these - `fo:break-before="page"` sits on the style,
         * so the break is a property of the paragraph rather than a run inside it. That makes it a
         * sibling emitted around the paragraph node, not a child of it, which is the one structural
         * difference from how DOCX's `<w:br w:type="page"/>` lands.
         */
        const pushStyleBreak = (styleName: string | null, targetArray: OfficeContentNode[], edge: 'before' | 'after') => {
            if (!config.includeBreakNodes || !styleName) return;
            const info = paragraphStyleMap[styleName];
            const breakType = edge === 'before' ? info?.breakBefore : info?.breakAfter;
            if (!breakType) return;
            targetArray.push({ type: 'break', metadata: { breakType } as BreakMetadata });
        };

        traverse = (node: Element, targetArray: OfficeContentNode[], forceHeading = false, sourceXml = '', asSheet = false) => {
            if (node.tagName === "text:p") {
                pushStyleBreak(node.getAttribute("text:style-name"), targetArray, 'before');
                const pContent = parseParagraphContent(node, paragraphStyleMap, styleMap, config, sourceXml);
                const type = (forceHeading || (node.getAttribute("text:style-name") || '').toLowerCase().includes('title')) ? 'heading' : 'paragraph';

                const metadata: any = {
                    ...(pContent.alignment ? { alignment: pContent.alignment } : {}),
                    ...(pContent.style ? { style: pContent.style } : {}),
                    ...(pContent.anchorIds?.length ? { anchorIds: pContent.anchorIds } : {})
                };

                const nodeId = node.getAttribute("xml:id") || node.getAttribute("text:id");
                if (nodeId) {
                    if (!metadata.anchorIds) metadata.anchorIds = [];
                    metadata.anchorIds.push(nodeId);
                }

                const pNode: OfficeContentNode = {
                    type,
                    text: pContent.text,
                    children: pContent.children,
                    metadata
                };

                if (type === 'heading' && pNode.metadata) {
                    (pNode.metadata as HeadingMetadata).level = (pNode.metadata as HeadingMetadata).level || 1;
                }

                // Clean up metadata if empty
                if (Object.keys(pNode.metadata || {}).length === 0) delete pNode.metadata;

                if (config.includeRawContent) {
                    pNode.rawContent = getRawContent(node, sourceXml, config);
                }
                targetArray.push(pNode);
                pushStyleBreak(node.getAttribute("text:style-name"), targetArray, 'after');
                lastWasList = false;
            } else if (node.tagName === "text:h") {
                pushStyleBreak(node.getAttribute("text:style-name"), targetArray, 'before');
                const level = parseInt(node.getAttribute("text:outline-level") || "1");
                const hContent = parseParagraphContent(node, paragraphStyleMap, styleMap, config, sourceXml);

                const metadata: any = {
                    level,
                    ...(hContent.alignment ? { alignment: hContent.alignment } : {}),
                    ...(hContent.style ? { style: hContent.style } : {}),
                    ...(hContent.anchorIds?.length ? { anchorIds: hContent.anchorIds } : {})
                };

                const nodeId = node.getAttribute("xml:id") || node.getAttribute("text:id");
                if (nodeId) {
                    if (!metadata.anchorIds) metadata.anchorIds = [];
                    metadata.anchorIds.push(nodeId);
                }

                const hNode: OfficeContentNode = {
                    type: 'heading',
                    text: hContent.text,
                    children: hContent.children,
                    metadata
                };

                if (config.includeRawContent) {
                    hNode.rawContent = getRawContent(node, sourceXml, config);
                }
                targetArray.push(hNode);
                pushStyleBreak(node.getAttribute("text:style-name"), targetArray, 'after');
                lastWasList = false;
            } else if (node.tagName === "table:table") {
                // Parse table with proper structure
                const tableNode = parseTable(node, paragraphStyleMap, styleMap, config, sourceXml, cellBudget);

                if (asSheet) {
                    tableNode.type = 'sheet';
                    const sheetName = node.getAttribute("table:name");
                    if (sheetName) {
                        tableNode.metadata = { ...tableNode.metadata, sheetName };
                    }
                }

                const tableId = node.getAttribute("xml:id") || node.getAttribute("table:name");
                if (tableId) {
                    if (!tableNode.metadata) tableNode.metadata = {};
                    (tableNode.metadata as any).anchorIds = (tableNode.metadata as any).anchorIds || [];
                    (tableNode.metadata as any).anchorIds.push(tableId);
                }

                if (config.includeRawContent) {
                    tableNode.rawContent = getRawContent(node, sourceXml, config);
                }
                targetArray.push(tableNode);
                lastWasList = false;
            } else if (node.tagName === "text:list") {
                // Parse list structure with proper listId tracking
                const listItems = getDirectChildren(node, "text:list-item");

                // Determine list type by checking the list style definition
                let listType: 'ordered' | 'unordered' = 'unordered';
                let isVisible = false;
                const listStyleName = node.getAttribute("text:style-name") || node.getAttribute("xml:id");
                let styleNameToCheck = listStyleName;

                // If no style name, check parent list for inherited style
                if (!styleNameToCheck) {
                    let parentNode = node.parentNode;
                    while (parentNode && !styleNameToCheck) {
                        if (parentNode.nodeName === 'text:list') {
                            styleNameToCheck = (parentNode as Element).getAttribute("text:style-name");
                            if (styleNameToCheck) break;
                        }
                        parentNode = parentNode.parentNode;
                    }
                }

                // Try to find list style in automatic styles or styles.xml to determine type and visibility
                if (styleNameToCheck) {
                    if (automaticStyles) {
                        const listStyles = getElementsByTagName(automaticStyles, "text:list-style");
                        for (const listStyle of listStyles) {
                            if (listStyle.getAttribute("style:name") === styleNameToCheck) {
                                // Check if it has bullet or number level styles
                                const bulletLevels = getElementsByTagName(listStyle, "text:list-level-style-bullet");
                                const numberLevels = getElementsByTagName(listStyle, "text:list-level-style-number");
                                const imageLevels = getElementsByTagName(listStyle, "text:list-level-style-image");

                                if (numberLevels.length > 0) {
                                    listType = 'ordered';
                                    isVisible = numberLevels.some(l => !!l.getAttribute("style:num-format"));
                                } else if (bulletLevels.length > 0) {
                                    listType = 'unordered';
                                    isVisible = bulletLevels.some(l => !!l.getAttribute("text:bullet-char"));
                                } else if (imageLevels.length > 0) {
                                    listType = 'unordered';
                                    isVisible = true;
                                }
                                break;
                            }
                        }
                    }

                    if (!isVisible && stylesDom) {
                        const officeStyles = getFirstElementByTagName(stylesDom, "office:styles");
                        if (officeStyles) {
                            const listStyles = getElementsByTagName(officeStyles, "text:list-style");
                            for (const listStyle of listStyles) {
                                if (listStyle.getAttribute("style:name") === styleNameToCheck) {
                                    const bulletLevels = getElementsByTagName(listStyle, "text:list-level-style-bullet");
                                    const numberLevels = getElementsByTagName(listStyle, "text:list-level-style-number");
                                    if (numberLevels.length > 0) {
                                        listType = 'ordered';
                                        isVisible = numberLevels.some(l => !!l.getAttribute("style:num-format"));
                                    } else if (bulletLevels.length > 0) {
                                        listType = 'unordered';
                                        isVisible = bulletLevels.some(l => !!l.getAttribute("text:bullet-char"));
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }

                // If the list is not visible, it's likely a layout list used by Impress.
                // We should traverse its items and treat their content as regular nodes.
                if (!isVisible) {
                    lastWasList = false;
                    for (let i = 0; i < listItems.length; i++) {
                        const item = listItems[i];
                        if (item.childNodes) {
                            for (let j = 0; j < item.childNodes.length; j++) {
                                const child = item.childNodes[j];
                                if (isElement(child)) { // Element
                                    traverse(child, targetArray, forceHeading, sourceXml);
                                }
                            }
                        }
                    }
                    return;
                }

                // List Continuity Logic:
                // If this list follows another list of the same type and style, or we are in ODP and it's sequential,
                // we should reuse the previous listId to maintain numbering.
                const isODP = fileType === 'odp';
                const sameStyle = styleNameToCheck && styleNameToCheck === lastListStyle;
                const sameType = listType === lastListType;

                let listId: string;
                if (lastWasList && (sameStyle || (isODP && sameType))) {
                    listId = currentListId!;
                } else {
                    // New list
                    listId = styleNameToCheck || `list-${++listIdCounter}`;
                    currentListId = listId;
                    lastListType = listType;
                    lastListStyle = styleNameToCheck;
                }
                lastWasList = true;

                // Calculate indentation level by counting parent text:list elements
                let indentation = 0;
                let parent = node.parentNode;
                while (parent) {
                    if (parent.nodeName === 'text:list') {
                        indentation++;
                    }
                    parent = parent.parentNode;
                }

                // Track list counters for this listId (similar to WordParser)
                if (!listCounters[listId]) {
                    listCounters[listId] = Object.create(null);
                }
                const indentKey = indentation.toString();
                if (listCounters[listId][indentKey] === undefined) {
                    listCounters[listId][indentKey] = -1; // Will increment to 0 on first item
                }

                // Process each list item
                for (let i = 0; i < listItems.length; i++) {
                    const item = listItems[i];

                    let hasIndexedThisItem = false;

                    // Iterate over direct children of list item (paragraphs, headings, nested lists)
                    if (item.childNodes) {
                        for (let j = 0; j < item.childNodes.length; j++) {
                            const child = item.childNodes[j];
                            if (isElement(child)) { // Element
                                const element = child;

                                if (element.tagName === "text:p" || element.tagName === "text:h") {
                                    if (!hasIndexedThisItem) {
                                        listCounters[listId][indentKey]++;
                                        hasIndexedThisItem = true;
                                        for (let k = indentation + 1; k < 10; k++) {
                                            if (listCounters[listId][k.toString()] !== undefined) {
                                                listCounters[listId][k.toString()] = -1;
                                            }
                                        }
                                    }
                                    const itemIndex = listCounters[listId][indentKey];
                                    const pContent = parseParagraphContent(element, paragraphStyleMap, styleMap, config, sourceXml);
                                    const segments = splitParagraphByBreaks(pContent);

                                    for (let k = 0; k < segments.length; k++) {
                                        const segment = segments[k];
                                        if (!segment.text.trim() && segment.children.length === 0) continue;

                                        const isFirst = k === 0;
                                        const nodeType = isFirst ? 'list' : 'paragraph';

                                        const node: OfficeContentNode = {
                                            type: nodeType as any,
                                            text: segment.text,
                                            children: segment.children,
                                            metadata: isFirst ? {
                                                listType,
                                                indentation,
                                                itemIndex,
                                                listId,
                                                alignment: pContent.alignment || 'left',
                                                style: pContent.style
                                            } as ListMetadata : {
                                                alignment: pContent.alignment || 'left',
                                                style: pContent.style
                                            }
                                        };

                                        // Special case for headings in lists
                                        if (isFirst && element.tagName === "text:h") {
                                            const level = parseInt(element.getAttribute("text:outline-level") || "1");
                                            (node.metadata as any).level = level;
                                        }

                                        if (config.includeRawContent) node.rawContent = getRawContent(element, sourceXml, config);
                                        targetArray.push(node);
                                    }
                                } else if (element.tagName === "text:list") {
                                    // Recursive call for nested list
                                    traverse(element, targetArray, forceHeading, sourceXml);
                                }
                            }
                        }
                    }
                }
            } else if (node.tagName === "draw:frame") {
                const presClass = node.getAttribute("presentation:class");
                const isHeading = presClass === "title" || presClass === "sub-title";

                // In presentations, frames often contain text-boxes, images, tables, or objects
                const textBox = getFirstElementByTagName(node, "draw:text-box");
                const image = getFirstElementByTagName(node, "draw:image");
                const table = getFirstElementByTagName(node, "table:table");
                const object = getFirstElementByTagName(node, "draw:object");

                if (textBox) {
                    traverse(textBox, targetArray, isHeading || forceHeading, sourceXml);
                } else if (table) {
                    const tableNode = parseTable(table, paragraphStyleMap, styleMap, config, sourceXml, cellBudget);
                    if (config.includeRawContent) tableNode.rawContent = getRawContent(table, sourceXml, config);
                    targetArray.push(tableNode);
                } else if (image) {
                    // Extract alt text from svg:title or svg:desc
                    let altText = '';
                    const svgTitle = getFirstElementByTagName(node, "svg:title");
                    const svgDesc = getFirstElementByTagName(node, "svg:desc");
                    if (svgTitle && svgTitle.textContent) {
                        altText = svgTitle.textContent;
                    } else if (svgDesc && svgDesc.textContent) {
                        altText = svgDesc.textContent;
                    }

                    // Extract image href to link to attachment
                    let imageHref = image.getAttribute("xlink:href") || '';
                    if (imageHref) {
                        imageHref = cleanAttachmentName(imageHref);
                    }

                    const metadata: any = {
                        attachmentName: imageHref,
                        ...(altText ? { altText } : {})
                    };

                    const frameId = node.getAttribute("xml:id") || node.getAttribute("draw:name");
                    if (frameId) {
                        metadata.anchorIds = [frameId];
                    }

                    const imageNode: OfficeContentNode = {
                        type: 'image',
                        text: '',
                        children: [],
                        metadata
                    };
                    if (config.includeRawContent) {
                        imageNode.rawContent = getRawContent(node, sourceXml, config);
                    }
                    targetArray.push(imageNode);
                } else if (object) {
                    // Handle embedded objects like charts or math formulas
                    const href = object.getAttribute("xlink:href");
                    if (href) {
                        const attachmentName = cleanAttachmentName(href);
                        const objectPath = `${attachmentName}/content.xml`;
                        const objectFile = files.find(f => f.path === objectPath || f.path.endsWith(objectPath));

                        if (objectFile) {
                            const objXml = parseXmlString(objectFile.content.toString());
                            const mathNode = getFirstElementByTagName(objXml, "math");

                            if (mathNode) {
                                // Math formula object at block level - a display equation, so the
                                // inner node is `math: 'block'` where the inline site above emits
                                // `math: 'inline'`.
                                const formulaText = mathmlToLatex(mathNode).trim();
                                const formulaNode: OfficeContentNode = {
                                    type: 'paragraph',
                                    text: formulaText,
                                    children: [
                                        {
                                            type: 'code',
                                            text: formulaText,
                                            metadata: { math: 'block' } as CodeMetadata
                                        }
                                    ]
                                };
                                if (config.includeRawContent) {
                                    formulaNode.rawContent = getRawContent(node, sourceXml, config);
                                }
                                targetArray.push(formulaNode);
                            } else {
                                const chartData = extractChartData(objectFile.content);

                                const chartNode: OfficeContentNode = {
                                    type: 'chart',
                                    text: chartData.rawTexts.join(" "),
                                    metadata: {
                                        attachmentName: attachmentName,
                                        chartData
                                    } as ChartMetadata
                                };
                                if (config.includeRawContent) chartNode.rawContent = getRawContent(node, sourceXml, config);
                                targetArray.push(chartNode);
                            }
                        } else {
                            const chartNode: OfficeContentNode = {
                                type: 'chart',
                                text: "",
                                metadata: { attachmentName: attachmentName }
                            };
                            if (config.includeRawContent) chartNode.rawContent = getRawContent(node, sourceXml, config);
                            targetArray.push(chartNode);
                        }
                    }
                }
            } else if (node.tagName === "draw:a") {
                // A frame that is a link (a clickable picture on a slide or drawing): its pictures carry the link.
                const href = node.getAttribute("xlink:href") || '';
                const internal = href.startsWith('#');
                const linked: OfficeContentNode[] = [];
                for (const child of Array.from(node.childNodes)) {
                    if (isElement(child)) traverse(child as Element, linked, forceHeading, sourceXml);
                }
                for (const child of linked) {
                    if (child.type === 'image' && href && (!internal || !config.ignoreInternalLinks)) {
                        child.metadata = { ...child.metadata, link: href, linkType: internal ? 'internal' : 'external' } as ImageMetadata;
                    }
                }
                for (const child of linked) targetArray.push(child);
            } else {
                if (node.childNodes) {
                    for (let i = 0; i < node.childNodes.length; i++) {
                        const child = node.childNodes[i];
                        if (isElement(child)) { // Element
                            traverse(child as Element, targetArray, forceHeading, sourceXml);
                        }
                    }
                }
            }
        };

        // ODS: Spreadsheet
        if (fileType === 'ods') {
            const spreadsheet = getFirstElementByTagName(body, "office:spreadsheet");
            if (spreadsheet) {
                const tables = getElementsByTagName(spreadsheet, "table:table");
                for (let i = 0; i < tables.length; i++) {
                    const table = tables[i];
                    const sheetName = table.getAttribute("table:name") || `Sheet${i + 1}`;
                    const rows: OfficeContentNode[] = [];

                    const tableRows = getElementsByTagName(table, "table:table-row");
                    let rowIndex = 0;

                    for (let r = 0; r < tableRows.length; r++) {
                        checkAbortSignal(config.abortSignal);
                        const row = tableRows[r];
                        const cells: OfficeContentNode[] = [];
                        const tableCells = getElementsByTagName(row, "table:table-cell");

                        let colIndex = 0;
                        const rowsRepeated = toRepeatCount(row.getAttribute("table:number-rows-repeated"));

                        for (let c = 0; c < tableCells.length; c++) {
                            const cell = tableCells[c];
                            const colsRepeated = toRepeatCount(cell.getAttribute("table:number-columns-repeated"));

                            // ODS cell notes are `<office:annotation>` children of the cell; extract them
                            // as comments. Their inner text:p must be kept out of the cell's own text
                            // below (getElementsByTagName is descendant-recursive, so it would otherwise
                            // pull the comment body into the cell value).
                            const cellComments = config.ignoreComments
                                ? []
                                : getDirectChildren(cell, "office:annotation").map(a =>
                                    buildAnnotationComment(a, parseParagraphContent, paragraphStyleMap, styleMap, config, xmlString));
                            const insideAnnotation = (el: Element): boolean => {
                                let p: Node | null = el.parentNode;
                                while (p && p !== cell) {
                                    if (isElement(p) && (p as Element).tagName === 'office:annotation') return true;
                                    p = p.parentNode;
                                }
                                return false;
                            };

                            // Extract text from cell (paragraphs inside cell, excluding a comment's body)
                            let cellText = "";
                            const children: OfficeContentNode[] = [];
                            const ps = getElementsByTagName(cell, "text:p").filter(p => !insideAnnotation(p));
                            for (let p = 0; p < ps.length; p++) {
                                const para = ps[p];

                                // Parse text:span elements for formatted text
                                const spans = getElementsByTagName(para, "text:span");
                                if (spans.length > 0) {
                                    for (const span of spans) {
                                        const styleName = span.getAttribute("text:style-name");
                                        // Through `mergeFormatting` like every other span site, so
                                        // an explicit `false` is dropped rather than written onto
                                        // the node - and so the node gets its own object instead of
                                        // aliasing the shared style-table entry.
                                        const formatting = mergeFormatting({}, styleName ? styleMap[styleName] : undefined);
                                        const text = span.textContent || '';
                                        cellText += text;

                                        const textNode: OfficeContentNode = {
                                            type: 'text',
                                            text: text,
                                            formatting: formatting
                                        };
                                        children.push(textNode);
                                    }
                                } else {
                                    // No spans - just direct text content
                                    const text = para.textContent || '';
                                    cellText += text;
                                    if (text.trim()) {
                                        const textNode: OfficeContentNode = {
                                            type: 'text',
                                            text: text,
                                            formatting: {}
                                        };
                                        children.push(textNode);
                                    }
                                }

                                if (p < ps.length - 1) cellText += "\n";
                            }

                            // Check for embedded draw:frame (images) in cell
                            const drawFrames = getElementsByTagName(cell, "draw:frame");
                            for (const frame of drawFrames) {
                                // Extract alt text from svg:title or svg:desc
                                let altText = '';
                                const svgTitle = getFirstElementByTagName(frame, "svg:title");
                                const svgDesc = getFirstElementByTagName(frame, "svg:desc");
                                if (svgTitle && svgTitle.textContent) {
                                    altText = svgTitle.textContent;
                                } else if (svgDesc && svgDesc.textContent) {
                                    altText = svgDesc.textContent;
                                }

                                // Extract image href
                                let imageHref = '';
                                const drawImages = getElementsByTagName(frame, "draw:image");
                                if (drawImages.length > 0) {
                                    const rawHref = drawImages[0].getAttribute("xlink:href");
                                    if (rawHref) {
                                        imageHref = cleanAttachmentName(rawHref);
                                    }
                                }

                                // Extract chart or math object href
                                let chartHref = '';
                                let isFormula = false;
                                let formulaText = '';
                                const drawObjects = getElementsByTagName(frame, "draw:object");
                                if (drawObjects.length > 0) {
                                    const href = drawObjects[0].getAttribute("xlink:href");
                                    if (href) {
                                        chartHref = cleanAttachmentName(href);
                                        const objectPath = `${chartHref}/content.xml`;
                                        const objectFile = files.find(f => f.path === objectPath || f.path.endsWith(objectPath));
                                        if (objectFile) {
                                            const objXml = parseXmlString(objectFile.content.toString());
                                            const mathNode = getFirstElementByTagName(objXml, "math");
                                            if (mathNode) {
                                                isFormula = true;
                                                formulaText = mathmlToLatex(mathNode).trim();
                                            }
                                        }
                                    }
                                }

                                if (isFormula) {
                                    cellText += formulaText;
                                    const formulaNode: OfficeContentNode = {
                                        type: 'code',
                                        text: formulaText,
                                        metadata: { math: 'inline' } as CodeMetadata
                                    };
                                    if (config.includeRawContent) {
                                        formulaNode.rawContent = getRawContent(frame, xmlString, config);
                                    }
                                    children.push(formulaNode);
                                } else if (drawImages.length > 0) {
                                    // logic for image node
                                    const imageNode: OfficeContentNode = {
                                        type: 'image',
                                        text: '', // Will be populated by assignAttachmentData
                                        children: [],
                                        metadata: {
                                            attachmentName: imageHref || chartHref, // Might be empty, will resolve in assignAttachmentData
                                            ...(altText ? { altText } : {})
                                        } as ImageMetadata
                                    };
                                    if (config.includeRawContent) {
                                        imageNode.rawContent = getRawContent(frame, xmlString, config);
                                    }
                                    children.push(imageNode);
                                } else if (chartHref) {
                                    const chartNode: OfficeContentNode = {
                                        type: 'chart',
                                        text: '', // Will be populated by assignAttachmentData
                                        children: [],
                                        metadata: {
                                            attachmentName: chartHref
                                        } as ChartMetadata
                                    };
                                    if (config.includeRawContent) {
                                        chartNode.rawContent = getRawContent(frame, xmlString, config);
                                    }
                                    children.push(chartNode);
                                }
                            }

                            // Add cell(s). The repeat count is attacker-influenced, so the loop
                            // is bounded by the document's remaining cell budget rather than by
                            // the attribute. An empty ODS cell creates nothing, so it costs no
                            // budget - which is what keeps the huge trailing-empty runs real
                            // files carry (number-columns-repeated="16384") free.
                            // For ODS an empty cell materializes nothing, so a huge
                            // number-columns-repeated on a blank cell (the normal way ODF marks a
                            // trailing empty run) is skipped in O(1) by advancing the column index
                            // rather than spinning the loop colsRepeated times for zero output -
                            // that spin was itself a CPU denial-of-service, unbounded by the cell
                            // budget because it created no cells to charge against.
                            const willMaterialize = (cellText || children.length > 0 || cellComments.length > 0 || fileType !== 'ods');
                            if (!willMaterialize) {
                                colIndex += colsRepeated;
                            } else {
                                const allowedCols = cellBudget.take(colsRepeated);
                                for (let k = 0; k < allowedCols; k++) {
                                    if ((k & 1023) === 0) checkAbortSignal(config.abortSignal);
                                    const cellNode: OfficeContentNode = {
                                        type: 'cell',
                                        text: cellText,
                                        children: children,
                                        metadata: { row: rowIndex, col: colIndex } as CellMetadata
                                    };
                                    if (config.includeRawContent) {
                                        cellNode.rawContent = getRawContent(cell, xmlString, config);
                                    }
                                    // Share the (read-only) comment nodes across repeated columns via a
                                    // shallow array copy, rather than deep-copying: a huge
                                    // number-columns-repeated on a cell carrying a large annotation would
                                    // otherwise duplicate the whole comment per column (bytes the cell
                                    // budget does not charge), a multi-GB amplification and CPU hang.
                                    if (cellComments.length) {
                                        cellNode.comments = k === 0 ? cellComments : cellComments.slice();
                                    }
                                    cells.push(cellNode);
                                    colIndex++;
                                }
                            }
                        }

                        // Add row(s). This is where the two repeats multiply: each repetition
                        // deep-copies the whole cell array, so rows x cols is what actually
                        // exhausts memory. Charge the copies against the same budget.
                        if (cells.length > 0) {
                            const allowedRows = Math.min(
                                rowsRepeated,
                                // The first row reuses `cells` rather than copying, so only the
                                // repeats beyond it cost budget.
                                1 + Math.floor(cellBudget.take(Math.max(0, (rowsRepeated - 1) * cells.length)) / cells.length)
                            );
                            for (let k = 0; k < allowedRows; k++) {
                                if ((k & 255) === 0) checkAbortSignal(config.abortSignal);
                                // First row reuses `cells`; a repeated row shallow-clones each cell (its
                                // own object + own metadata for the row-index fix below) but SHARES the
                                // read-only children/comments/text by reference. JSON deep-copying instead
                                // serialized every cell's full content per row, so a row of many cells each
                                // carrying large text or a comment built a multi-hundred-MB string (a
                                // RangeError / OOM the cell-count budget did not bound).
                                const rowCells = k === 0
                                    ? cells
                                    : cells.map(c => ({ ...c, metadata: c.metadata ? { ...c.metadata } : c.metadata, children: c.children ? c.children.slice() : c.children, comments: c.comments ? c.comments.slice() : c.comments }) as OfficeContentNode);
                                const rowNode: OfficeContentNode = {
                                    type: 'row',
                                    children: rowCells,
                                    metadata: undefined
                                };
                                // Fix row index in metadata for repeated rows
                                if (k > 0) {
                                    rowNode.children?.forEach(c => {
                                        if (c.metadata && 'row' in c.metadata) {
                                            (c.metadata as CellMetadata).row = rowIndex;
                                        }
                                    });
                                }
                                if (config.includeRawContent) {
                                    rowNode.rawContent = getRawContent(row, xmlString, config);
                                }
                                rows.push(rowNode);
                                rowIndex++;
                            }
                        } else {
                            rowIndex += rowsRepeated;
                        }
                    }

                    const sheetNode: OfficeContentNode = {
                        type: 'sheet',
                        children: rows,
                        metadata: { sheetName } as SheetMetadata
                    };
                    if (config.includeRawContent) {
                        sheetNode.rawContent = getRawContent(table, xmlString, config);
                    }
                    content.push(sheetNode);
                }
            }
        }
        // ODP: Presentation
        else if (fileType === 'odp') {
            const presentation = getFirstElementByTagName(body, "office:presentation");
            if (presentation) {
                const pages = getDirectChildren(presentation, "draw:page");
                const odpNotes: OfficeContentNode[] = [];

                for (let i = 0; i < pages.length; i++) {
                    const page = pages[i];
                    const slideNode: OfficeContentNode = {
                        type: 'slide',
                        children: [],
                        metadata: { slideNumber: i + 1 } as SlideMetadata
                    };

                    // Separate page content and notes
                    let noteNode: OfficeContentNode | undefined = undefined;
                    const pageChildren = page.childNodes;
                    if (pageChildren) {
                        for (let j = 0; j < pageChildren.length; j++) {
                            const child = pageChildren[j];
                            if (isElement(child)) { // Element
                                const element = child as Element;
                                if (element.tagName === "presentation:notes") {
                                    if (!config.ignoreNotes) {
                                        noteNode = {
                                            type: 'note',
                                            children: [],
                                            metadata: {
                                                slideNumber: i + 1,
                                                noteId: `slide-note-${i + 1}`
                                            } as SlideMetadata
                                        };
                                        traverse(element, noteNode.children!, false, xmlString);
                                    }
                                    continue;
                                }
                                // A slide-level comment sits as a direct child of draw:page (outside any
                                // text paragraph); attach it to the slide node's comments.
                                if (element.tagName === "office:annotation") {
                                    if (!config.ignoreComments) {
                                        (slideNode.comments ||= []).push(
                                            buildAnnotationComment(element, parseParagraphContent, paragraphStyleMap, styleMap, config, xmlString));
                                    }
                                    continue;
                                }
                                traverse(element, slideNode.children!, false, xmlString);
                            }
                        }
                    }

                    if (config.includeRawContent) {
                        slideNode.rawContent = getRawContent(page, xmlString, config);
                    }
                    content.push(slideNode);

                    if (noteNode && noteNode.children && noteNode.children.length > 0) {
                        if (!slideNode.notes) slideNode.notes = [];
                        slideNode.notes.push(noteNode);
                    }
                }

                if (odpNotes.length > 0) {
                    content.push(...odpNotes);
                }
            }
        }
        // ODG: Drawing / graphics document (LibreOffice Draw). Each draw:page becomes a `page` node
        // (the same primitive PDF uses), and its shapes are walked with the shared, format-agnostic
        // `traverse` that the ODP branch already relies on for frames, text boxes, groups, tables
        // and images.
        else if (fileType === 'odg') {
            const drawing = getFirstElementByTagName(body, "office:drawing");
            if (drawing) {
                const pages = getDirectChildren(drawing, "draw:page");
                for (let i = 0; i < pages.length; i++) {
                    const page = pages[i];
                    const drawName = page.getAttribute("draw:name");
                    const pageNode: OfficeContentNode = {
                        type: 'page',
                        children: [],
                        metadata: {
                            pageNumber: i + 1,
                            ...(drawName ? { pageName: drawName } : {})
                        } as PageMetadata
                    };
                    const pageChildren = page.childNodes;
                    if (pageChildren) {
                        for (let j = 0; j < pageChildren.length; j++) {
                            const child = pageChildren[j];
                            if (isElement(child)) {
                                // A page-level comment is a direct child of draw:page; attach it to the
                                // page node's comments rather than letting traverse drop it.
                                if ((child as Element).tagName === "office:annotation") {
                                    if (!config.ignoreComments) {
                                        (pageNode.comments ||= []).push(
                                            buildAnnotationComment(child as Element, parseParagraphContent, paragraphStyleMap, styleMap, config, xmlString));
                                    }
                                    continue;
                                }
                                traverse(child as Element, pageNode.children!, false, xmlString);
                            }
                        }
                    }
                    if (config.includeRawContent) {
                        pageNode.rawContent = getRawContent(page, xmlString, config);
                    }
                    content.push(pageNode);
                }
            }
        }
        // ODT: Text Document (and generic fallback)
        else {
            const textDoc = getFirstElementByTagName(body, "office:text");
            if (textDoc) {
                traverse(textDoc, content, false, xmlString);
            }
        }
    };

    if (mainContentFile) {
        parseContentXml(mainContentFile.content.toString());
    }

    // Attachments
    const attachments: OfficeAttachment[] = [];
    const mediaFiles = files.filter(f => f.path.match(/(Pictures|media)\/.*/));

    // ODP/ODT Chart Extraction
    if (config.extractAttachments) {
        const objectFiles = files.filter(f => f.path.match(/Object \d+\/content\.xml/));
        for (const objFile of objectFiles) {
            const objXml = parseXmlString(objFile.content.toString());
            const isChart = getElementsByTagName(objXml, "chart:chart").length > 0;

            if (isChart) {
                const objectId = objFile.path.split('/')[0];
                const attachment: OfficeAttachment = {
                    type: 'chart',
                    mimeType: 'application/vnd.oasis.opendocument.chart', // Accurate ODF chart type
                    data: objFile.content.toString('base64'),
                    name: objectId,
                    extension: 'xml'
                };

                // Extract data from chart XML
                const chartData = extractChartData(objFile.content);

                if (chartData.rawTexts.length > 0) {
                    attachment.chartData = chartData;
                }

                attachments.push(attachment);
            }
        }
    }

    if (config.extractAttachments) {
        for (const media of mediaFiles) {
            const attachment = createAttachment(media.path.split('/').pop() || 'image', media.content);
            attachments.push(attachment);

            if (config.ocr && attachment.mimeType.startsWith('image/')) {
                const ocrText = await ocrDuringParse(media.content, config, attachment.name);
                if (ocrText !== undefined) attachment.ocrText = ocrText;
            }
        }
    }

    const metaFile = files.find(f => f.path.match(metaFileRegex));
    const metadata = metaFile ? parseOfficeMetadata(metaFile.content.toString()) : {};

    // Helper: Resolve ODS chart cell references to actual values
    // ODS charts often link to cell ranges (e.g., [Sheet1.$A$1:.$A$5]) instead of embedding values
    const resolveChartReferences = (chartData: ChartData, nodes: OfficeContentNode[]) => {
        const getValuesFromReference = (ref: string): string[] => {
            // Remove brackets: [Sheet.$A$1:.$A$5] -> Sheet.$A$1:.$A$5
            const cleanRef = ref.replace(/^\[|\]$/g, '');
            const [startPart, endPart] = cleanRef.split(':');

            const lastDotIdx = startPart.lastIndexOf('.');
            if (lastDotIdx === -1) return [ref];

            const sheetName = startPart.substring(0, lastDotIdx).replace(/^'|'$/g, '');
            const startCoord = startPart.substring(lastDotIdx + 1).replace(/\$/g, '');

            let endCoord = startCoord;
            if (endPart) {
                if (endPart.startsWith('.')) {
                    endCoord = endPart.substring(1).replace(/\$/g, '');
                } else {
                    const endLastDotIdx = endPart.lastIndexOf('.');
                    endCoord = endPart.substring(endLastDotIdx + 1).replace(/\$/g, '');
                }
            }

            const parseCoord = (coord: string) => {
                const colMatch = coord.match(/[A-Z]+/);
                const rowMatch = coord.match(/\d+/);
                if (!colMatch || !rowMatch) return null;

                const colStr = colMatch[0];
                let colIdx = 0;
                for (let i = 0; i < colStr.length; i++) {
                    colIdx = colIdx * 26 + (colStr.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
                }
                colIdx -= 1;
                const rowIdx = parseInt(rowMatch[0]) - 1;
                return { r: rowIdx, c: colIdx };
            };

            const start = parseCoord(startCoord);
            const end = parseCoord(endCoord);

            if (!start || !end) return [ref];

            const sheet = nodes.find(n => n.type === 'sheet' && (n.metadata as SheetMetadata)?.sheetName === sheetName);
            if (!sheet || !sheet.children) return [ref];

            const values: string[] = [];
            // Collect all matching cells
            for (const row of sheet.children) {
                if (row.children) {
                    for (const cell of row.children) {
                        const meta = cell.metadata as CellMetadata;
                        if (meta && meta.row >= start.r && meta.row <= end.r && meta.col >= start.c && meta.col <= end.c) {
                            values.push(cell.text || '');
                        }
                    }
                }
            }
            return values.length > 0 ? values : [];
        };

        // Resolve DataSets
        for (const ds of chartData.dataSets) {
            const newValues: string[] = [];
            for (const val of ds.values) {
                if (val.startsWith('[')) newValues.push(...getValuesFromReference(val));
                else newValues.push(val);
            }
            ds.values = newValues;
        }

        // Resolve Labels
        const newLabels: string[] = [];
        for (const label of chartData.labels) {
            if (label.startsWith('[')) newLabels.push(...getValuesFromReference(label));
            else newLabels.push(label);
        }
        chartData.labels = newLabels;

        // Rebuild rawTexts
        chartData.rawTexts = [];
        if (chartData.title) chartData.rawTexts.push(chartData.title);
        for (const ds of chartData.dataSets) {
            if (ds.name) chartData.rawTexts.push(ds.name);
            chartData.rawTexts.push(...chartData.labels);
            chartData.rawTexts.push(...ds.values);
        }
    };

    // Apply resolution to all chart attachments
    for (const att of attachments) {
        if (att.type === 'chart' && att.chartData) {
            resolveChartReferences(att.chartData, content);
        }
    }

    // Link OCR and Chart text to content nodes
    // Link OCR and Chart text to content nodes (with heuristic for unlinked images)
    const assignAttachmentData = (nodes: OfficeContentNode[]) => {
        // Step 1: Identify unused image attachments globally
        const usedAttachmentNames = new Set<string>();
        const traverseForNames = (ns: OfficeContentNode[]) => {
            for (const n of ns) {
                if (n.metadata && 'attachmentName' in n.metadata) {
                    const name = (n.metadata as ImageMetadata).attachmentName;
                    if (name) usedAttachmentNames.add(name);
                }
                if (n.children) traverseForNames(n.children);
            }
        };
        traverseForNames(nodes);

        const unusedImages = attachments.filter(a => a.type === 'image' && a.name && !usedAttachmentNames.has(a.name));
        let unusedImageIndex = 0;

        const processNode = (node: OfficeContentNode) => {
            if ((node.type === 'image' || node.type === 'chart') && node.metadata && 'attachmentName' in node.metadata) {
                let attachmentName = (node.metadata as ImageMetadata).attachmentName;

                // Heuristic: If name is empty, try to assign an unused image attachment
                if (!attachmentName && node.type === 'image' && unusedImageIndex < unusedImages.length) {
                    const fallbackAtt = unusedImages[unusedImageIndex++];
                    attachmentName = fallbackAtt.name;
                    (node.metadata as ImageMetadata).attachmentName = attachmentName;
                }

                if (attachmentName) {
                    const attachment = attachments.find(a => a.name === attachmentName);

                    if (attachment) {
                        if (attachment.ocrText) {
                            node.text = attachment.ocrText;
                        }
                        if (attachment.chartData && node.type === 'chart') {
                            node.text = attachment.chartData.rawTexts.join(config.newlineDelimiter);
                        }
                    }
                }
            }
            // Internal recursion
            if (node.children) {
                node.children.forEach(processNode);
            }
        };

        nodes.forEach(processNode);
    };
    assignAttachmentData(content);

    // Create combined styleMap for metadata (matches DOCX format)
    const combinedStyleMap: { [key: string]: { formatting: TextFormatting, alignment?: 'left' | 'center' | 'right' | 'justify' } } = Object.create(null);
    for (const styleName in styleMap) {
        combinedStyleMap[styleName] = {
            formatting: styleMap[styleName],
            alignment: paragraphStyleMap[styleName]?.alignment
        };
    }
    // Also add styles that only have alignment
    for (const styleName in paragraphStyleMap) {
        if (!combinedStyleMap[styleName]) {
            combinedStyleMap[styleName] = {
                formatting: {},
                alignment: paragraphStyleMap[styleName]?.alignment
            };
        }
    }


    // Master-page headers/footers (ODT only). ODF text documents keep authored headers/footers in
    // styles.xml, not content.xml, so they are otherwise invisible to every destination. Restricted
    // to `odt`: a Calc (ODS) or Impress (ODP) master page is print furniture built from auto-fields
    // (`text:sheet-name`, `text:page-number`), whose placeholder text (e.g. LibreOffice's "???")
    // would otherwise leak in as literal header content. Only the first master page's visible
    // header/footer is read, through the same block traversal (the shape WordParser populates).
    let auxiliary: OfficeAuxiliaryContent | undefined;
    if (!config.ignoreHeadersAndFooters && stylesDom && fileType === 'odt') {
        const stylesXmlStr = stylesFile ? stylesFile.content.toString() : undefined;
        const masterStyles = getFirstElementByTagName(stylesDom, "office:master-styles");
        const masterPage = masterStyles ? getFirstElementByTagName(masterStyles, "style:master-page") : undefined;
        if (masterPage) {
            const visible = (el: Element | null | undefined): el is Element => !!el && el.getAttribute("style:display") !== "false";
            const headerEl = getFirstElementByTagName(masterPage, "style:header");
            const footerEl = getFirstElementByTagName(masterPage, "style:footer");
            const headers: OfficeContentNode[] = [];
            const footers: OfficeContentNode[] = [];
            if (visible(headerEl)) traverse(headerEl, headers, false, stylesXmlStr);
            if (visible(footerEl)) traverse(footerEl, footers, false, stylesXmlStr);
            // Resolve any image/chart nodes in the header/footer the same way the body is (OCR text,
            // chart data, attachment linkage); without this a picture in a master-page header/footer
            // would carry no data and render blank in every destination.
            if (headers.length) assignAttachmentData(headers);
            if (footers.length) assignAttachmentData(footers);
            if (headers.length || footers.length) {
                auxiliary = {
                    ...(headers.length ? { headers } : {}),
                    ...(footers.length ? { footers } : {}),
                };
            }
        }
    }

    return createAST(
        fileType,
        {
            ...metadata,
            styleMap: plainRecord(combinedStyleMap)
        },
        content,
        attachments,
        config,
        auxiliary,
    );
};

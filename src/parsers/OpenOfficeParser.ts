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

import { presetOfOdfShape } from '../utils/shapeNames.js';
import { BreakMetadata, CellMetadata, ChartData, ChartMetadata, CodeMetadata, FullOfficeParserConfig, HeaderFooterMetadata, HeadingMetadata, ImageMetadata, IndentationMetadata, ListMetadata, PageSetup, ParagraphSpacing, ShapeMetadata, OfficeAttachment, OfficeAuxiliaryContent, OfficeContentNode, OfficeParserAST, OfficeParserConfig, OfficeWarningType, PageMetadata, SheetMetadata, SlideMetadata, SupportedFileType, TextFormatting, TextMetadata } from '../types.js';
import { MAX_PAGE_PT } from '../utils/pageLimits.js';
import { createAST } from '../utils/astUtils.js';
import { chartRawTexts, extractChartData, MAX_CHART_VALUES } from '../utils/chartUtils.js';
import { checkAbortSignal, logWarning } from '../utils/errorUtils.js';
import { mathmlToLatex } from '../utils/mathUtils.js';
import { cellSpan, clampRepeat, MAX_COL_SPAN, MAX_ROW_SPAN } from '../utils/numberUtils.js';
import { TextBuilder } from '../utils/textUtils.js';
import { lookupTable, plainRecord } from '../utils/lookupUtils.js';
import { lengthToPt } from '../utils/officeGenUtils.js';
import { bulletCharacter, formatListNumber, odfNumberFormat } from '../utils/listMarkers.js';

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
    private readonly weights = new WeakMap<object, number>();
    constructor(private limit: number, private config: OfficeParserConfig) {
        this.remaining = limit;
    }
    /**
     * How many of `repeats` further copies of content weighing `weight` (see `weighCell`) may be
     * made. A repeat shares its content with the first copy, so building it is cheap, but everything
     * after the parse (each pass over the AST, each generator) meets that content once per copy: a
     * 700-byte spreadsheet repeating a 2,000-span cell 100,000 times made 200 MB of text and ran the
     * HTML and CSV generators out of memory. The cell count alone does not bound it, so the content
     * copied is charged to the document's repeated-content budget (see repeatUtils).
     */
    takeRepeats(repeats: number, weight: number): number {
        return takeRepeats(this.config, repeats, weight);
    }
    /**
     * What one copy of a cell writes: its text, and for each node it holds the node's text, formatting
     * and metadata values (a style's 64 KB font is written with every copy) plus 16 for the node itself
     * (a generator spends on a node what it spends on a dozen or two characters of text).
     */
    weighCell(text: string, children: OfficeContentNode[] | undefined, comments: OfficeContentNode[] | undefined): number {
        return text.length + this.weigh(children) + this.weigh(comments);
    }
    /**
     * Counted along every path, as a generator writes a shared node each time it meets it (a nested
     * table's repeated cells share their content), and memoized so a shared subtree is summed once.
     */
    private weigh(nodes: OfficeContentNode[] | undefined): number {
        if (!nodes || nodes.length === 0) return 0;
        let total = this.weights.get(nodes);
        if (total !== undefined) return total;
        total = 0;
        for (const node of nodes) {
            let weight = this.weights.get(node);
            if (weight === undefined) {
                weight = 16 + (node.text?.length ?? 0) + valuesLength(node.formatting) + valuesLength(node.metadata) + this.weigh(node.children) + this.weigh(node.comments) + this.weigh(node.notes);
                this.weights.set(node, weight);
            }
            total += weight;
        }
        this.weights.set(nodes, total);
        return total;
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

/**
 * A repeated row's copy of a cell: its own object and metadata (for the row index), sharing the
 * read-only children, comments and text. Deep copies serialized every cell's content per row, a
 * multi-hundred-MB string from a small row the cell-count budget did not bound. The shared rawContent
 * is charged again, as every node that carries it is.
 */
const repeatCell = (c: OfficeContentNode, config: OfficeParserConfig): OfficeContentNode => ({
    ...c,
    metadata: c.metadata ? { ...c.metadata } : c.metadata,
    children: c.children ? c.children.slice() : c.children,
    comments: c.comments ? c.comments.slice() : c.comments,
    ...(c.rawContent !== undefined ? { rawContent: chargeRawContent(c.rawContent, config) } : {}),
}) as OfficeContentNode;

/** Resolves the configured cell budget (the documented default when not set), plus one cell for each byte of the document (see budgetUtils). */
const createCellBudget = (config: OfficeParserConfig): CellBudget =>
    new CellBudget((config.decompressionLimits?.maxTableCells ?? 1000000) + TABLE_CELLS_PER_BYTE * documentBytesOf(config), config);

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
    /** `fo:margin-left`, `fo:margin-right` and `fo:text-indent`, in twips as Word's are read. */
    indentation?: IndentationMetadata;
    /** `fo:margin-top`, `fo:margin-bottom`, `fo:line-height` and `style:line-height-at-least`. */
    spacing?: ParagraphSpacing;
    /** `style:master-page-name`: a paragraph of this style starts a page of that master page. */
    masterPage?: string;
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
function textContentSkipping(node: Node, skipTags: ReadonlySet<string>): string {
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
/** A note body's or annotation's own paragraphs (see getOutermostElements): a nested note's or annotation's are read by its own reader, and taking them too doubled per level of notes in notes (528 bytes of ODT ended the process out of memory). */
const NESTED_NOTES = new Set(['text:note', 'office:annotation']);
const ownParagraphs = (container: Element): Element[] => getOutermostElements(container, 'text:p', NESTED_NOTES);
/**
 * Text elements in a paragraph that show nothing where they stand: a note (read by its own branch, or
 * ignored), hidden text, a script, the list number the writers number lists with themselves, and blocks
 * that a paragraph cannot hold.
 */
/** What a table cell holds paragraphs in, read through: a section, a list and its items, an index's body. */
const CELL_CONTAINERS: ReadonlySet<string> = new Set(['text:section', 'text:list', 'text:list-item', 'text:list-header', 'text:index-body',
    'text:table-of-content', 'text:illustration-index', 'text:table-index', 'text:object-index', 'text:user-index', 'text:alphabetical-index', 'text:bibliography']);

const INLINE_NOT_SHOWN: ReadonlySet<string> = new Set(['text:note', 'text:hidden-text', 'text:hidden-paragraph', 'text:script', 'text:number', 'text:p', 'text:h', 'text:list', 'text:section', 'text:tracked-changes']);
/**
 * What a paragraph's text, read whole when its content made no runs, leaves out: what shows nothing
 * (a comment and a note are read on their own, hidden text, a script, tracked changes) and a note's
 * citation read as the note. The paragraph read as its whole text showed a hidden-text field's or a
 * script's content.
 */
const FALLBACK_NOT_SHOWN: ReadonlySet<string> = new Set(['office:annotation', 'text:note', 'text:note-ref', 'text:hidden-text', 'text:hidden-paragraph', 'text:script', 'text:tracked-changes']);
/**
 * A frame's own text box, image, table, object, title or description (children of it, as the schema
 * has them): looked up through the frame's whole subtree, frames nested in text boxes took time in
 * the product of their depth and content (4.5 KB of ODT took 30 seconds), and a nested frame's image
 * was taken for the outer one's.
 */
const frameChild = (frame: Element, tag: string): Element | undefined => getDirectChildren(frame, tag)[0];

/**
 * A row's cells in order, those a merge covers (`table:covered-table-cell`) included, so a reader
 * counts their columns: skipped, the cells after a merge took its columns (shifted left).
 */
const rowCells = (row: Element): Element[] => {
    const cells: Element[] = [];
    for (let i = 0; i < (row.childNodes?.length ?? 0); i++) {
        const child = row.childNodes[i];
        if (isElement(child) && (child.tagName === 'table:table-cell' || child.tagName === 'table:covered-table-cell')) cells.push(child);
    }
    return cells;
};
/** A cell's comment is its own node (its body is not the cell's text). */
const CELL_SKIP = new Set(['office:annotation']);
/** A sheet's column definitions come before its rows: a nested table's, in a cell, are not the sheet's. */
const COLUMN_SKIP = new Set(['table:table-row', 'table:table-rows', 'table:table-row-group', 'table:table-header-rows']);

function buildAnnotationComment(
    element: Element,
    parsePara: ParseParaFn,
    paraStyleMap: Record<string, ParagraphStyleInfo>,
    styleMap: Record<string, TextFormatting>,
    config: OfficeParserConfig,
    sourceXml: string,
): OfficeContentNode {
    // The annotation's own author and date (children of it): looked up through its subtree, an
    // annotation in a paragraph of an annotation, and so on, took minutes.
    const author = getDirectChildren(element, "dc:creator")[0]?.textContent || undefined;
    const date = getDirectChildren(element, "dc:date")[0]?.textContent || undefined;
    const children: OfficeContentNode[] = [];
    let text = '';
    for (const cp of ownParagraphs(element)) {
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
import { takeRepeats, valuesLength } from '../utils/repeatUtils.js';
import { documentBytesOf, TABLE_CELLS_PER_BYTE } from '../utils/budgetUtils.js';
import { chargeRawContent, getAllElementsByTagName, getDirectChildren, getOutermostElements, getElementsByTagName, getFirstElementByTagName, getRawContent, isElement, parseOfficeMetadata, parseXmlString } from '../utils/xmlUtils.js';
import { extractFiles, findRequiredPart } from '../utils/zipUtils.js';
import { appendAll, dropTrailingSectionStarts, liftFloatingFrames } from '../utils/nodeListUtils.js';
import { odfPageNumberFormat } from '../utils/pageNumbers.js';
import { inRanges } from '../utils/sheetUtils.js';

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
/** A table-cell style's text looks, which the cell's runs take under their own (its background is the cell's, not theirs). */
function cellTextLooks(style: TextFormatting | undefined): TextFormatting | undefined {
    if (!style) return undefined;
    const { backgroundColor: _background, ...looks } = style;
    return Object.keys(looks).length ? looks : undefined;
}

/** Whether a cell holds something: a value of a type, a paragraph with text, or a comment. */
function holdsValue(cell: Element): boolean {
    if (cell.getAttribute('office:value-type')) return true;
    return getDirectChildren(cell, 'office:annotation').length > 0 || getDirectChildren(cell, 'text:p').some(p => (p.textContent || '').trim() !== '');
}

/**
 * What an ODS cell holds (see CellMetadata.valueType and CellMetadata.value), from its `office:value-type`:
 * a number (a percentage's or a currency amount's too) as written, a date as Excel counts its days (from
 * 1899-12-30, with the day's fraction for its time), a time (a duration) in days, a truth value, an error;
 * text where it is a string or says nothing. Its shown text is the cell's `text`: values were read from it, a
 * number shown as 1.23E+05 written as 123000.
 */
function heldValue(cell: Element): Pick<CellMetadata, 'valueType' | 'value'> {
    // (An error is a string to ODF; LibreOffice says which it is.)
    if (cell.getAttribute('calcext:value-type') === 'error') return { valueType: 'error' };
    const type = cell.getAttribute('office:value-type') || cell.getAttribute('calcext:value-type') || '';
    // (Days to ten places: a second is about a hundred thousandth of one.)
    const days = (n: number) => String(Number(n.toFixed(10)));
    if (type === 'float' || type === 'percentage' || type === 'currency') {
        const value = (cell.getAttribute('office:value') || '').trim();
        return value && Number.isFinite(Number(value)) ? { valueType: 'number', value } : { valueType: 'text' };
    }
    if (type === 'date') {
        const m = /^(-?\d{4,})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/.exec(cell.getAttribute('office:date-value') || '');
        if (!m) return { valueType: 'text' };
        const ms = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] || 0), Number(m[5] || 0)) + Number(m[6] || 0) * 1000;
        return Number.isFinite(ms) ? { valueType: 'date', value: days((ms - Date.UTC(1899, 11, 30)) / 86_400_000) } : { valueType: 'text' };
    }
    if (type === 'time') {
        const m = /^(-)?P(?:(\d+(?:\.\d+)?)D)?(?:T(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?)?$/.exec((cell.getAttribute('office:time-value') || '').trim());
        if (!m || !(m[2] || m[3] || m[4] || m[5])) return { valueType: 'text' };
        const total = Number(m[2] || 0) + (Number(m[3] || 0) + (Number(m[4] || 0) + Number(m[5] || 0) / 60) / 60) / 24;
        return Number.isFinite(total) ? { valueType: 'time', value: days(m[1] ? -total : total) } : { valueType: 'text' };
    }
    if (type === 'boolean') {
        const value = (cell.getAttribute('office:boolean-value') || '').trim().toLowerCase();
        return value === 'true' || value === 'false' ? { valueType: 'boolean', value } : { valueType: 'text' };
    }
    return { valueType: 'text' };
}

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

    // Each embedded object (a formula or chart in `Object N/content.xml`) read once, whatever refers
    // to it: read again for each frame naming it, a 10 KB document naming one object 20,000 times
    // took a minute and a gigabyte. Its text is written at every reference, so references after the
    // first repeat at most REPEATED_OBJECT_TEXT characters of it in all; past that, a reference to a
    // chart is the chart without its text (its data stays in chartData), and to a formula, the
    // formula's first reference is its only copy.
    const REPEATED_OBJECT_TEXT = 16 * 1024 * 1024;
    // Spaces `text:s` elements may make in all (see its reader).
    let spaceBudget = 16 * 1024 * 1024;
    let repeatedObjectText = REPEATED_OBJECT_TEXT;
    const embeddedObjects = new Map<string, { formula?: string; chartData?: ChartData; chartText?: string; uses: number } | null>();
    // The package's files by path, and by each tail of their path one of its last 16 folders starts
    // (`Object 1/content.xml` of `Pictures/Object 1/content.xml`), the first of each: a scan of every
    // file per object took objects x files, and a tail at every folder of a long path read the rest of
    // the path again at each.
    let filesByPath: Map<string, (typeof files)[number]> | undefined;
    const fileAt = (path: string) => {
        if (!filesByPath) {
            filesByPath = new Map();
            for (const f of files) if (!filesByPath.has(f.path)) filesByPath.set(f.path, f);
            for (const f of files) {
                let at = f.path.length;
                for (let depth = 0; depth < 16 && at > 0; depth++) {
                    at = f.path.lastIndexOf('/', at - 1);
                    if (at === -1) break;
                    const tail = f.path.slice(at + 1);
                    if (tail && !filesByPath.has(tail)) filesByPath.set(tail, f);
                }
            }
        }
        return filesByPath.get(path);
    };
    const readEmbeddedObject = (attachmentName: string): { formula?: string; chartData?: ChartData; chartText?: string } | null => {
        let entry = embeddedObjects.get(attachmentName);
        if (entry === undefined) {
            const objectPath = `${attachmentName}/content.xml`;
            const objectFile = fileAt(objectPath);
            entry = null;
            if (objectFile) {
                const mathNode = getFirstElementByTagName(parseXmlString(objectFile.content.toString(), { config }), "math");
                if (mathNode) entry = { formula: mathmlToLatex(mathNode).trim(), uses: 0 };
                else {
                    const chartData = extractChartData(objectFile.content, config);
                    entry = { chartData, chartText: chartData.rawTexts.join(" "), uses: 0 };
                }
            }
            embeddedObjects.set(attachmentName, entry);
        }
        if (!entry) return null;
        if (entry.uses++ === 0) return entry;
        const size = (entry.formula ?? entry.chartText ?? '').length;
        if (size <= repeatedObjectText) { repeatedObjectText -= size; return entry; }
        return { chartData: entry.chartData, chartText: entry.chartData ? '' : undefined, formula: entry.formula === undefined ? undefined : '' };
    };

    // What a list style makes a list (ordered or not, and whether its marks show), from each styles
    // element's list styles indexed by name once: looked up by scanning every list style for each
    // list, a 5 KB document of many lists and list styles took two minutes. (`withImages`: a picture
    // bullet counts, as the automatic styles read it.)
    const listStyleIndex = new Map<Element, Map<string, Element>>();
    const listStyleKinds = new Map<Element, { listType: 'ordered' | 'unordered'; isVisible: boolean } | undefined>();
    let officeStylesElement: Element | null | undefined;
    const listStyleNamed = (styles: Element, name: string): Element | undefined => {
        let index = listStyleIndex.get(styles);
        if (!index) {
            index = new Map();
            for (const listStyle of getElementsByTagName(styles, "text:list-style")) {
                const styleName = listStyle.getAttribute("style:name");
                if (styleName && !index.has(styleName)) index.set(styleName, listStyle);
            }
            listStyleIndex.set(styles, index);
        }
        return index.get(name);
    };
    const listStyleKind = (styles: Element, name: string, withImages: boolean): { listType: 'ordered' | 'unordered'; isVisible: boolean } | undefined => {
        const listStyle = listStyleNamed(styles, name);
        if (!listStyle) return undefined;
        // (A style is read by one of the two lookups only: automatic styles with pictures, office styles without.)
        if (listStyleKinds.has(listStyle)) return listStyleKinds.get(listStyle);
        const bulletLevels = getElementsByTagName(listStyle, "text:list-level-style-bullet");
        const numberLevels = getElementsByTagName(listStyle, "text:list-level-style-number");
        let kind: { listType: 'ordered' | 'unordered'; isVisible: boolean } | undefined;
        if (numberLevels.length > 0) kind = { listType: 'ordered', isVisible: numberLevels.some(l => !!l.getAttribute("style:num-format")) };
        else if (bulletLevels.length > 0) kind = { listType: 'unordered', isVisible: bulletLevels.some(l => !!l.getAttribute("text:bullet-char")) };
        else if (withImages && getElementsByTagName(listStyle, "text:list-level-style-image").length > 0) kind = { listType: 'unordered', isVisible: true };
        listStyleKinds.set(listStyle, kind);
        return kind;
    };

    /** A list style's levels by `text:level` (1-based), from the automatic styles or else the office ones. */
    const odfListLevelsCache = new Map<Element | undefined, Map<string, Map<number, Element> | undefined>>();
    const odfListLevels = (automaticStyles: Element | undefined, name: string): Map<number, Element> | undefined => {
        let cache = odfListLevelsCache.get(automaticStyles);
        if (!cache) odfListLevelsCache.set(automaticStyles, cache = new Map());
        if (cache.has(name)) return cache.get(name);
        if (stylesDom) officeStylesElement ??= getFirstElementByTagName(stylesDom, "office:styles") ?? null;
        const listStyle = (automaticStyles ? listStyleNamed(automaticStyles, name) : undefined)
            ?? (officeStylesElement ? listStyleNamed(officeStylesElement, name) : undefined);
        let levels: Map<number, Element> | undefined;
        if (listStyle) {
            levels = new Map();
            for (const level of Array.from(listStyle.childNodes).filter(isElement)) {
                const n = Number(level.getAttribute("text:level"));
                if (Number.isInteger(n) && n >= 1 && n <= 10 && !levels.has(n)) levels.set(n, level);
            }
        }
        cache.set(name, levels);
        return levels;
    };

    /**
     * A list item's marker as LibreOffice shows it (see ListMetadata.marker): its level's prefix, the
     * numbers of as many levels as it displays (each in its level's format, joined by dots), and suffix;
     * or its bullet's character. Undefined when it is the item's number and a dot, or a round bullet.
     */
    const odfListMarker = (levels: Map<number, Element>, depth: number, counters: Record<string, number>, listType: 'ordered' | 'unordered', itemIndex: number): string | undefined => {
        const level = levels.get(depth + 1);
        if (!level) return undefined;
        let marker: string;
        if (level.tagName === 'text:list-level-style-bullet') {
            const font = getDirectChildren(level, "style:text-properties")[0];
            marker = bulletCharacter(level.getAttribute("text:bullet-char") || '', font?.getAttribute("fo:font-family") || font?.getAttribute("style:font-name") || undefined);
        } else if (level.tagName === 'text:list-level-style-number') {
            const shown = Math.min(Math.max(Number(level.getAttribute("text:display-levels")) || 1, 1), depth + 1);
            const numbers: string[] = [];
            for (let d = depth + 1 - shown; d <= depth; d++) {
                const style = levels.get(d + 1);
                const start = Number(style?.getAttribute("text:start-value")) || 1;
                numbers.push(formatListNumber(start + Math.max(0, counters[String(d)] ?? 0), odfNumberFormat(style?.getAttribute("style:num-format"))));
            }
            marker = (level.getAttribute("style:num-prefix") || '') + numbers.join('.') + (level.getAttribute("style:num-suffix") || '');
        } else return undefined;
        const plain = listType === 'ordered' ? `${itemIndex + 1}.` : '•';
        return marker === plain ? undefined : marker;
    };

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
    const stylesDom = stylesFile ? parseXmlString(stylesFile.content.toString(), { config }) : undefined;

    // The master pages (styles.xml), by name, and the page layouts they name: the document's pages.
    const masterPages = new Map<string, Element>();
    const pageLayouts = new Map<string, Element>();
    if (stylesDom) {
        for (const master of getElementsByTagName(stylesDom, "style:master-page")) {
            const name = master.getAttribute("style:name");
            if (name && !masterPages.has(name)) masterPages.set(name, master);
        }
        for (const layout of getElementsByTagName(stylesDom, "style:page-layout")) {
            const name = layout.getAttribute("style:name");
            if (name && !pageLayouts.has(name)) pageLayouts.set(name, layout);
        }
    }
    /** The master page the document starts on: the first block's style's, else Standard, else the first. */
    let firstMaster: string | undefined = masterPages.has('Standard') ? 'Standard' : masterPages.keys().next().value;
    /** The master page of each section after the first, in order; set while the text body is read. */
    const sectionMasters: string[] = [];
    let readingBody = false, bodyStarted = false;

    /** Whether a master page shows a header or footer of a kind (`style:header`, `-left`, `-first`). */
    const shown = (master: Element | undefined, tag: string): Element | undefined => {
        const el = master ? getDirectChildren(master, tag)[0] : undefined;
        return el && el.getAttribute("style:display") !== 'false' ? el : undefined;
    };

    /**
     * A master page's pages (see PageSetup), from its page layout: how they are numbered and whether the
     * first has a header of its own always; where geometry is wanted and the layout gives a size, the
     * page's size and margins, the body below its header and above its footer (ODF's top margin is the
     * header's; the header's height and gap come between it and the body), and its columns.
     */
    const odfPageSetup = (masterName: string | undefined): PageSetup | undefined => {
        const master = masterName ? masterPages.get(masterName) : undefined;
        const layout = master ? pageLayouts.get(master.getAttribute("style:page-layout-name") || '') : undefined;
        const props = layout ? getDirectChildren(layout, "style:page-layout-properties")[0] : undefined;
        if (!master) return undefined;
        const setup: PageSetup = {};
        // A first page of its own: its own first page's parts, or a master that hands over to another after
        // its first page (a word processor's "First Page" style, and ODT output's section title pages).
        const next = master?.getAttribute("style:next-style-name");
        if (shown(master, "style:header-first") || shown(master, "style:footer-first") || (next && next !== masterName && masterPages.has(next))) setup.differentFirstPage = true;
        // How its pages are numbered (`style:num-format`: i for front matter, 1 where it gives none).
        // (A master of no page layout is on LibreOffice's default one: numbered 1, 2, 3, of no size here.)
        const pageNumbers = odfPageNumberFormat(props?.getAttribute("style:num-format"));
        if (pageNumbers) setup.pageNumberFormat = pageNumbers;
        if (config.ignorePageGeometry || !props) return setup;
        const pt = (el: Element | undefined, attr: string) => { const v = lengthToPt(el?.getAttribute(attr) || undefined); return v != null && v >= 0 && v <= MAX_PAGE_PT ? Math.round(v * 100) / 100 : undefined; };
        const width = pt(props, "fo:page-width"), height = pt(props, "fo:page-height");
        if (!width || !height) return setup;
        const margin = { top: pt(props, "fo:margin-top") ?? 0, right: pt(props, "fo:margin-right") ?? 0, bottom: pt(props, "fo:margin-bottom") ?? 0, left: pt(props, "fo:margin-left") ?? 0 };
        Object.assign(setup, { width, height, margin });
        // The room a header or footer takes: its height and its gap to the body.
        const room = (styleTag: string, gapAttr: string) => {
            const hf = layout ? getDirectChildren(getDirectChildren(layout, styleTag)[0] ?? layout, "style:header-footer-properties")[0] : undefined;
            if (!hf || !getDirectChildren(layout!, styleTag).length) return 0;
            return (pt(hf, "fo:min-height") ?? pt(hf, "svg:height") ?? 0) + (pt(hf, gapAttr) ?? 0);
        };
        if (shown(master, "style:header")) { setup.headerDistance = margin.top; margin.top += room("style:header-style", "fo:margin-bottom"); }
        if (shown(master, "style:footer")) { setup.footerDistance = margin.bottom; margin.bottom += room("style:footer-style", "fo:margin-top"); }
        const columns = odfColumns(getDirectChildren(props, "style:columns")[0], width - margin.left - margin.right);
        if (columns) setup.columns = columns;
        return setup;
    };
    /** A setup where it sets a page (a sheet's or a slide's is its page, nothing without one). */
    const sized = (setup: PageSetup | undefined) => setup?.width ? setup : undefined;
    /** A `style:columns` element's columns (their number, the gap, a line between), where there are two or more. */
    function odfColumns(columns: Element | undefined, available?: number): PageSetup['columns'] | undefined {
        const count = Math.min(16, Math.round(Number(columns?.getAttribute("fo:column-count")) || 0));
        if (!columns || count < 2) return undefined;
        const sep = getDirectChildren(columns, "style:column-sep")[0];
        // Each column's width where they are given (style:column): its relative width holds the column's
        // indents (the halves of the gaps beside it), a share of the width the columns stand in (LibreOffice's
        // shares sum to 65535; a share read as twentieths of a point gave columns of a thousand points).
        const own = getDirectChildren(columns, "style:column");
        const indent = (c: Element, attr: string) => Math.max(0, lengthToPt(c.getAttribute(attr) || undefined) ?? 0);
        const rels = own.map(c => parseFloat(c.getAttribute("style:rel-width") || ''));
        const total = rels.reduce((sum, r) => sum + (Number.isFinite(r) && r > 0 ? r : 0), 0);
        const widths = own.map((c, i) => {
            const rel = rels[i];
            if (!(Number.isFinite(rel) && rel > 0)) return 0;
            const whole = available !== undefined && available > 0 && total > 0 ? rel / total * available : rel / 20;
            return Math.round((whole - indent(c, "fo:start-indent") - indent(c, "fo:end-indent")) * 100) / 100;
        });
        const differ = widths.length === count && widths.every(w => w > 0 && w < 1e5) && widths.some(w => Math.abs(w - widths[0]) > 0.5);
        // The gap: as stated, else the indents between the first two columns.
        const stated = lengthToPt(columns.getAttribute("fo:column-gap") || undefined);
        const gap = stated ?? (own.length >= 2 ? indent(own[0], "fo:end-indent") + indent(own[1], "fo:start-indent") : 0);
        return { count, gap: gap >= 0 && gap < 1584 ? Math.round(gap * 100) / 100 : 0, ...(differ ? { widths } : {}), ...(sep && (sep.getAttribute("style:style") ?? 'solid') !== 'none' ? { separator: true } : {}) };
    }
    /** Each section style's columns (`style:section-properties`'s `style:columns`), by its name; read against the page they stand on (see sectionColumnsOn). */
    const sectionColumns = new Map<string, Element | undefined>();
    /** A section style's columns on a page: their widths a share of the page's text width. */
    const sectionColumnsOn = (styleName: string, page: PageSetup | undefined) => {
        const el = sectionColumns.get(styleName);
        return el ? odfColumns(el, page?.width ? page.width - (page.margin?.left ?? 0) - (page.margin?.right ?? 0) : undefined) : undefined;
    };
    /** The columns of the sections being read, innermost last (a page break inside one is set in them). */
    const openSectionColumns: (string | undefined)[] = [];
    const content: OfficeContentNode[] = [];
    const notes: OfficeContentNode[] = [];

    // Notes by their `text:id`, for a reference to one (`text:note-ref`, a note cited again): the ids
    // the document gives its notes (found first, so a reference before its note is known for one), the
    // notes read so far, the runs waiting for a note read later, and how many note bodies are being read
    // (a reference inside one stays its citation, so no note can come to hold itself).
    const documentNoteIds = new Set<string>();
    const collectNoteIds = (dom: Document) => {
        for (const note of getAllElementsByTagName(dom, 'text:note')) {
            const id = note.getAttribute('text:id');
            if (id) documentNoteIds.add(id);
        }
    };
    if (stylesDom) collectNoteIds(stylesDom);
    const notesById = new Map<string, OfficeContentNode>();
    const runsAwaitingNote = new Map<string, OfficeContentNode[]>();
    let noteBodyDepth = 0;

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

            if (style.getAttribute("style:family") === 'section') {
                const sectionProps = getDirectChildren(style, "style:section-properties")[0];
                if (!sectionColumns.has(name)) {
                    const cols = sectionProps ? getDirectChildren(sectionProps, "style:columns")[0] : undefined;
                    sectionColumns.set(name, cols && Math.round(Number(cols.getAttribute("fo:column-count")) || 0) >= 2 ? cols : undefined);
                }
                continue;
            }
            const styleInfo: ParagraphStyleInfo = {};
            const masterPage = style.getAttribute("style:master-page-name");
            if (masterPage) styleInfo.masterPage = masterPage;

            // Parse paragraph properties for alignment and drop caps
            // A style's own properties (children of it, as the schema has them): looked up through the
            // whole subtree, styles nested in styles took time in the product of their depth and content.
            const paraProps = getDirectChildren(style, "style:paragraph-properties")[0];
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

                // Indents, in twips as the AST has Word's (a negative first-line indent is a hanging one).
                const twips = (attr: string) => { const pt = lengthToPt(paraProps.getAttribute(attr) || undefined); return pt ? Math.round(pt * 20) : 0; };
                const indentation: IndentationMetadata = {};
                const left = twips("fo:margin-left"), right = twips("fo:margin-right"), first = twips("fo:text-indent");
                if (left) indentation.left = left;
                if (right) indentation.right = right;
                if (first > 0) indentation.firstLine = first;
                else if (first < 0) indentation.hanging = -first;
                if (Object.keys(indentation).length) styleInfo.indentation = indentation;

                // Spacing: margins above and below, and the line height as a percentage (proportional)
                // or a length (fixed), or a least height. 100% is what readers assume.
                const spacing: ParagraphSpacing = {};
                // In points to the hundredth (a length in cm is no whole number of them).
                const hundredths = (pt: number) => Math.round(pt * 100) / 100;
                const points = (attr: string) => { const pt = lengthToPt(paraProps.getAttribute(attr) || undefined); return pt == null ? undefined : hundredths(pt); };
                const above = points("fo:margin-top"), below = points("fo:margin-bottom");
                if (above !== undefined && above >= 0) spacing.before = above;
                if (below !== undefined && below >= 0) spacing.after = below;
                const lineHeight = (paraProps.getAttribute("fo:line-height") || '').trim();
                const percent = /^(\d+(?:\.\d+)?)%$/.exec(lineHeight);
                if (percent) { const multiple = Number(percent[1]) / 100; if (multiple > 0 && multiple !== 1) spacing.lineMultiple = multiple; }
                else if (lineHeight && lineHeight !== 'normal') { const pt = lengthToPt(lineHeight); if (pt && pt > 0) spacing.lineHeight = hundredths(pt); }
                const least = points("style:line-height-at-least");
                if (least && least > 0 && spacing.lineHeight === undefined && spacing.lineMultiple === undefined) { spacing.lineHeight = least; spacing.lineAtLeast = true; }
                if (Object.keys(spacing).length) styleInfo.spacing = spacing;

                // Detect Drop Caps
                const dropCap = getDirectChildren(paraProps, "style:drop-cap")[0];
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

            // A table's breaks before and after it (its style's table properties).
            const tableProps = getDirectChildren(style, "style:table-properties")[0];
            if (tableProps && !paraProps) {
                const before = tableProps.getAttribute("fo:break-before"), after = tableProps.getAttribute("fo:break-after");
                if (before === 'page' || before === 'column') styleInfo.breakBefore = before;
                if (after === 'page' || after === 'column') styleInfo.breakAfter = after;
            }

            if (Object.keys(styleInfo).length > 0) {
                paragraphStyleMap[name] = styleInfo;
            }

            // Parse text properties
            const textProps = getDirectChildren(style, "style:text-properties")[0];
            // Parse table cell properties (for ODS background)
            const cellProps = getDirectChildren(style, "style:table-cell-properties")[0];

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
    // ── Drawing styles (graphic, presentation and drawing-page families): slides' shapes, documents' frames ──
    /** The styles of both parts by name (graphic, presentation and drawing-page families among them). */
    let drawingStyles: Map<string, Element> | undefined;
    const drawingStyle = (doc: Document, name: string | null): Element | undefined => {
        if (!name) return undefined;
        if (!drawingStyles) {
            drawingStyles = new Map();
            for (const dom of [doc, stylesDom]) for (const style of dom ? getElementsByTagName(dom, "style:style") : []) {
                const styleName = style.getAttribute("style:name");
                if (styleName && !drawingStyles.has(styleName)) drawingStyles.set(styleName, style);
            }
        }
        return drawingStyles.get(name);
    };
    /** A style's property (on its graphic or drawing-page properties), or its parent's. */
    const styleProperty = (doc: Document, name: string | null, tag: string, attr: string): string | undefined => {
        for (let style = drawingStyle(doc, name), depth = 0; style && depth < 16; style = drawingStyle(doc, style.getAttribute("style:parent-style-name")), depth++) {
            const value = getDirectChildren(style, tag)[0]?.getAttribute(attr);
            if (value) return value;
        }
        return undefined;
    };
    const odpLength = (value: string | null): number | undefined => {
        const pt = value ? lengthToPt(value) : null;
        return pt != null && Number.isFinite(pt) && Math.abs(pt) < 16128 ? pt : undefined;
    };
    /**
     * A document's frame that floats beside the text (anchored to a paragraph, a character or the page,
     * not set in its line as a character): its look and where it stands (ShapeMetadata.float), from its
     * attributes and its graphic style. Undefined for a frame in its line, or one of no width.
     */
    const frameFloat = (frame: Element): ShapeMetadata | undefined => {
        const length = (value: string | null | undefined) => { const pt = odpLength(value ?? null); return pt === undefined ? undefined : Math.round(pt * 100) / 100; };
        const anchor = frame.getAttribute("text:anchor-type") || 'paragraph';
        // (Without geometry, its text is read where it stands, as a frame in its line is. A slide's or a
        // drawing's text box has its place on the slide: a frame in its text is read in it.)
        if (config.ignorePageGeometry || fileType !== 'odt') return undefined;
        if (anchor !== 'paragraph' && anchor !== 'char' && anchor !== 'page') return undefined;
        const doc = frame.ownerDocument, style = frame.getAttribute("draw:style-name");
        const prop = (attr: string) => doc ? styleProperty(doc, style, "style:graphic-properties", attr) : undefined;
        const width = length(frame.getAttribute("svg:width") ?? frame.getAttribute("fo:min-width"));
        if (!width || width <= 0) return undefined;
        const height = length(frame.getAttribute("svg:height") ?? frame.getAttribute("fo:min-height"));
        const hpos = prop("style:horizontal-pos"), hrel = prop("style:horizontal-rel") ?? (anchor === 'page' ? 'page' : 'paragraph');
        const vpos = prop("style:vertical-pos"), vrel = prop("style:vertical-rel") ?? (anchor === 'page' ? 'page' : 'paragraph');
        const align = hpos === 'left' || hpos === 'inside' ? 'left' as const : hpos === 'right' || hpos === 'outside' ? 'right' as const : hpos === 'center' ? 'center' as const : undefined;
        const x = length(frame.getAttribute("svg:x")), y = length(frame.getAttribute("svg:y"));
        // ODF's `none` is no text beside it (above and below only); `run-through` text over or under it.
        const wrapping = prop("style:wrap");
        const wrap = wrapping === 'run-through' ? 'none' as const : !wrapping || wrapping === 'none' ? 'aboveBelow' as const : 'beside' as const;
        const margin = (side: string) => length(prop(`fo:margin-${side}`) ?? prop("fo:margin")) ?? 0;
        const gap = { top: margin('top'), right: margin('right'), bottom: margin('bottom'), left: margin('left') };
        const background = prop("fo:background-color") ?? (prop("draw:fill") === 'solid' ? prop("draw:fill-color") : undefined);
        const fill = background && /^#[0-9a-f]{6}$/i.test(background) ? background.toUpperCase() : undefined;
        // `0.5pt solid #3355aa`: the outline's width and colour.
        const border = /^\s*([\d.]+[a-z]*)\s+[a-z-]+\s+(#[0-9a-f]{6})\s*$/i.exec(prop("fo:border") ?? '');
        const lineWidth = border ? length(border[1]) : undefined;
        const line = border && lineWidth ? { color: border[2].toUpperCase(), width: lineWidth } : undefined;
        const pad = (side: string) => length(prop(`fo:padding-${side}`) ?? prop("fo:padding")) ?? 0;
        return {
            ...(fill || line ? { geometry: 'rect' } : {}),
            ...(fill ? { fill } : {}),
            ...(line ? { line } : {}),
            textInsets: { top: pad('top'), right: pad('right'), bottom: pad('bottom'), left: pad('left') },
            float: {
                horizontalFrom: hrel === 'page' || hrel === 'page-start-margin' || hrel === 'page-end-margin' ? 'page' : hrel === 'page-content' ? 'margin' : 'column',
                verticalFrom: vrel === 'page' ? 'page' : vrel === 'page-content' ? 'margin' : 'paragraph',
                ...(align ? { align } : { x: x ?? 0 }),
                // (Placed in the middle or at the foot of what it is measured down from: at its top.)
                y: vpos === 'from-top' || vpos === undefined ? y ?? 0 : 0,
                width,
                ...(height ? { height } : {}),
                wrap,
                ...(gap.top || gap.right || gap.bottom || gap.left ? { gap } : {}),
            },
        };
    };

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
     * @param withFrames - Whether frames' pictures and objects are read: a spreadsheet cell reads those itself (a text box's text is read either way)
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
        sourceXml: string = '',
        withFrames = true,
        into?: { children: OfficeContentNode[]; anchorIds: string[] }
    ): { text: string; children: OfficeContentNode[]; anchorIds: string[] } => {
        // An element's content goes straight into the paragraph's lists (`into`, for the elements
        // nested in it): returned and copied up, each level of spans, links or fields copied everything
        // inside it again (250 spans around 600,000 runs took four times as long as none).
        const children: OfficeContentNode[] = into?.children ?? [];
        const anchorIds: string[] = into?.anchorIds ?? [];
        const sink = { children, anchorIds };
        let fullText = '';

        if (!node.childNodes) return { text: '', children, anchorIds };

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
                    // The document's runs of spaces are bounded in all too: each clamped run is 10,000 at
                    // most, and 50,000 of them (4 KB of ODT) made 500 million characters. Past the
                    // budget a run is one space (the spacing, never text, is what is lost).
                    const count = Math.min(clampRepeat(parseInt(element.getAttribute('text:c') || '1')), Math.max(1, spaceBudget));
                    spaceBudget -= count;
                    const spaces = ' '.repeat(count);
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
                    // A line break: a `break` node, as a DOCX `w:br` and an HTML `<br>` are, whatever
                    // includeBreakNodes says (a typed line break is content). As a "\n" text run, only
                    // the text, ODT and LaTeX writers showed it: Markdown and HTML wrote a space, and
                    // RTF and DOCX joined the words.
                    fullText += '\n';
                    children.push({
                        type: 'break',
                        formatting: parentFormatting,
                        metadata: { breakType: 'textWrapping', ...(linkMetadata || {}) } as BreakMetadata
                    });
                } else if (tagName === 'text:page-number' || tagName === 'text:page-count') {
                    // A page field: its text is the number when the file was saved, marked so that each
                    // page can show its own (TextMetadata.field).
                    const text = element.textContent || '';
                    // (The format it writes its number in, where it sets one: `style:num-format="i"`.)
                    // (Without one it writes the page's: absent, not decimal.)
                    const fieldFormat = element.getAttribute("style:num-format") ? odfPageNumberFormat(element.getAttribute("style:num-format")) : undefined;
                    fullText += text;
                    children.push({
                        type: 'text',
                        text,
                        formatting: parentFormatting,
                        metadata: { ...(linkMetadata || {}), field: tagName === 'text:page-number' ? 'page' : 'pages', ...(fieldFormat ? { fieldFormat } : {}) } as TextMetadata
                    });
                } else if (tagName === 'text:span') {
                    // Formatted text span
                    const styleName = element.getAttribute("text:style-name");
                    const formatting = styleName ? mergeFormatting(parentFormatting, styleMap[styleName]) : parentFormatting;

                    fullText += parseInlineContent(element, styleMap, config, notes, paragraphStyleMap, formatting, linkMetadata, sourceXml, withFrames, sink).text;
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

                    fullText += parseInlineContent(element, styleMap, config, notes, paragraphStyleMap, parentFormatting, newLinkMetadata, sourceXml, withFrames, sink).text;
                } else if (tagName === 'text:note' && !config.ignoreNotes) {
                    // Footnote or endnote
                    const noteClass = (element.getAttribute('text:note-class') || 'footnote') as 'footnote' | 'endnote';
                    const noteId = element.getAttribute('text:id') || element.getAttribute('xml:id') || undefined;
                    const noteBody = getFirstElementByTagName(element, "text:note-body");

                    if (noteBody) {
                        // Extract note content recursively
                        const notePs = ownParagraphs(noteBody);
                        const noteChildren: OfficeContentNode[] = [];
                        let noteText = '';

                        noteBodyDepth++;
                        for (const np of notePs) {
                            const npContent = parseParagraphContent(np, paragraphStyleMap, styleMap, config, sourceXml);
                            noteText += (noteText ? ' ' : '') + npContent.text;

                            const npNode: OfficeContentNode = {
                                type: 'paragraph',
                                text: npContent.text,
                                children: npContent.children,
                                metadata: {
                                    ...(npContent.alignment ? { alignment: npContent.alignment } : {}),
                                    ...(npContent.indentation ? { paragraphIndentation: npContent.indentation } : {}),
                                    ...(npContent.spacing ? { spacing: npContent.spacing } : {}),
                                    ...(npContent.anchorIds?.length ? { anchorIds: npContent.anchorIds } : {})
                                }
                            };
                            noteChildren.push(npNode);
                        }
                        noteBodyDepth--;

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
                        // The note for references to it: those read before it hold it now.
                        if (noteId && !notesById.has(noteId)) {
                            notesById.set(noteId, noteNode);
                            for (const run of runsAwaitingNote.get(noteId) ?? []) (run.notes ??= []).push(noteNode);
                            runsAwaitingNote.delete(noteId);
                        }
                    }
                } else if (tagName === 'text:note-ref' && documentNoteIds.has(element.getAttribute('text:ref-name') || '')
                    && (element.getAttribute('text:reference-format') || 'text') === 'text' && (config.ignoreNotes || noteBodyDepth === 0)) {
                    // A note cited again (LibreOffice's cross-reference to a footnote, and how the ODT
                    // writer refers to a note it has written): the run before it holds that note, as the
                    // first citation's does (a parser shares one note node among its references). It
                    // was read as a field, its number as literal text. With notes left out it shows
                    // nothing, as a note's own citation does; inside a note it stays its number.
                    if (!config.ignoreNotes) {
                        const name = element.getAttribute('text:ref-name') || '';
                        let run = children[children.length - 1];
                        if (!run || run.type !== 'text') {
                            run = { type: 'text', text: '' };
                            children.push(run);
                        }
                        const note = notesById.get(name);
                        if (note) (run.notes ??= []).push(note);
                        else {
                            const waiting = runsAwaitingNote.get(name);
                            if (waiting) waiting.push(run);
                            else runsAwaitingNote.set(name, [run]);
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
                } else if (tagName === 'text:ruby') {
                    // Ruby: the base text, then its annotation (a reading) in parentheses, as plain text
                    // writes ruby.
                    const base = getDirectChildren(element, 'text:ruby-base')[0];
                    const annotation = getDirectChildren(element, 'text:ruby-text')[0];
                    for (const part of [base, annotation]) {
                        if (!part) continue;
                        // The reading's parenthesis goes before its content: kept only when there is some.
                        const openAt = children.length;
                        if (part === annotation) children.push({ type: 'text', text: '(', formatting: parentFormatting, metadata: linkMetadata ? { ...linkMetadata } : undefined });
                        const partText = parseInlineContent(part, styleMap, config, notes, paragraphStyleMap, parentFormatting, linkMetadata, sourceXml, withFrames, sink).text;
                        if (part === annotation) {
                            // No reading: its parenthesis goes, and a note or picture placed in it stays.
                            if (!partText) { children.splice(openAt, 1); continue; }
                            fullText += '(';
                        }
                        fullText += partText;
                        if (part === annotation) {
                            fullText += ')';
                            children.push({ type: 'text', text: ')', formatting: parentFormatting, metadata: linkMetadata ? { ...linkMetadata } : undefined });
                        }
                    }
                } else if (tagName.startsWith('text:') && !INLINE_NOT_SHOWN.has(tagName)) {
                    // A field (a date, a page or sequence number, a cross-reference, an author's name, a
                    // placeholder), `text:meta`, and any other text element: the text it shows, where it
                    // stands. Each was dropped, so "Figure 3" read as "Figure " and a reference to a
                    // heading as nothing.
                    fullText += parseInlineContent(element, styleMap, config, notes, paragraphStyleMap, parentFormatting, linkMetadata, sourceXml, withFrames, sink).text;
                } else if (tagName === 'draw:frame' && !withFrames) {
                    // A spreadsheet cell reads its frames' pictures and objects itself; here, the text of
                    // a frame's text box only (its own frames, nested, are the cell's too).
                    const drawTextBox = frameChild(element, "draw:text-box");
                    if (drawTextBox) {
                        for (const boxParagraph of ownParagraphs(drawTextBox)) {
                            fullText += parseInlineContent(boxParagraph, styleMap, config, notes, paragraphStyleMap, parentFormatting, linkMetadata, sourceXml, false, sink).text;
                        }
                    }
                } else if (tagName === 'draw:frame') {
                    const frame = element;
                    const drawTextBox = frameChild(frame, "draw:text-box");
                    const drawObject = frameChild(frame, "draw:object");

                    if (drawTextBox) {
                        const textBoxChildren: OfficeContentNode[] = [];
                        traverse(drawTextBox, textBoxChildren, false, sourceXml);
                        const textBoxText = textBoxChildren.map(c => c.text || '').join('\n');
                        // A frame floating beside the text: its blocks in a `drawing` node, which goes before
                        // the paragraph (see liftFloatingFrames); its text is not the paragraph's.
                        const floating = frameFloat(frame);
                        if (floating) {
                            children.push({ type: 'drawing', text: textBoxText, children: textBoxChildren, metadata: floating });
                        } else {
                            appendAll(children, textBoxChildren);
                            fullText += textBoxText;
                        }
                    } else if (drawObject) {
                        const href = drawObject.getAttribute("xlink:href");
                        let isFormula = false;
                        let formulaText = '';
                        let attachmentName = '';
                        if (href) {
                            attachmentName = cleanAttachmentName(href);
                            const object = readEmbeddedObject(attachmentName);
                            if (object?.formula !== undefined) {
                                isFormula = true;
                                formulaText = object.formula;
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
                            const svgTitle = frameChild(frame, "svg:title");
                            const svgDesc = frameChild(frame, "svg:desc");
                            if (svgTitle && svgTitle.textContent) {
                                altText = svgTitle.textContent;
                            } else if (svgDesc && svgDesc.textContent) {
                                altText = svgDesc.textContent;
                            }

                            let imageHref = '';
                            const drawImages = getDirectChildren(frame, "draw:image");
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
                        const svgTitle = frameChild(frame, "svg:title");
                        const svgDesc = frameChild(frame, "svg:desc");
                        if (svgTitle && svgTitle.textContent) {
                            altText = svgTitle.textContent;
                        } else if (svgDesc && svgDesc.textContent) {
                            altText = svgDesc.textContent;
                        }

                        let imageHref = '';
                        const drawImages = getDirectChildren(frame, "draw:image");
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
        sourceXml: string,
        inherited?: TextFormatting
    ): { text: string; children: OfficeContentNode[]; alignment?: 'left' | 'center' | 'right' | 'justify'; indentation?: IndentationMetadata; spacing?: ParagraphSpacing; style?: string; anchorIds?: string[] } => {
        // Get paragraph style for alignment and drop caps
        const paraStyle = node.getAttribute("text:style-name");
        const styleInfo = paraStyle ? paraStyleMap[paraStyle] : undefined;
        const alignment = styleInfo?.alignment;
        const dropCap = styleInfo?.dropCap;
        // The looks of what holds the paragraph (a spreadsheet cell's style), under its own style's.
        const formatting = mergeFormatting(mergeFormatting({}, inherited), paraStyle ? styleMap[paraStyle] : undefined);

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
            const fullText = textContentSkipping(node, FALLBACK_NOT_SHOWN);
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

        return { text: content.text, children: content.children, alignment, indentation: styleInfo?.indentation, spacing: styleInfo?.spacing, style: paraStyle || undefined, anchorIds: content.anchorIds };
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
            if (child.type === "break" && (child.metadata as BreakMetadata | undefined)?.breakType === 'textWrapping') {
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
        // Rows sit directly under <table:table> or in a wrapper: a repeating header block
        // (<table:table-header-rows>, written by LibreOffice and by this library's own ODT generator),
        // <table:table-rows>, or a group of rows (<table:table-row-group>, nested to any depth).
        // Reading only the direct <table:table-row> children dropped every wrapped row, so walk the
        // children in document order and descend into the wrappers, flagging header rows. (Nested
        // tables are still excluded: a cell is never descended into.)
        const tableRows: { row: Element; isHeader: boolean }[] = [];
        const pendingRows: { node: Node; isHeader: boolean }[] = [];
        const pushRowChildren = (element: Element, isHeader: boolean) => {
            for (let i = (element.childNodes?.length || 0) - 1; i >= 0; i--) pendingRows.push({ node: element.childNodes[i], isHeader });
        };
        pushRowChildren(tableNode, false);
        while (pendingRows.length) {
            const { node: child, isHeader } = pendingRows.pop()!;
            if (!isElement(child)) continue;
            const element = child as Element;
            if (element.tagName === "table:table-row") tableRows.push({ row: element, isHeader });
            else if (element.tagName === "table:table-header-rows") pushRowChildren(element, true);
            else if (element.tagName === "table:table-rows" || element.tagName === "table:table-row-group") pushRowChildren(element, isHeader);
        }
        let rowIndex = 0;

        for (const { row, isHeader } of tableRows) {
            checkAbortSignal(config.abortSignal);
            const cells: OfficeContentNode[] = [];
            // The row's own cells (not nested tables'), covered ones included (see rowCells)
            const tableCells = rowCells(row);
            const rowsRepeated = toRepeatCount(row.getAttribute("table:number-rows-repeated"));

            let colIndex = 0;
            // What one copy of this row writes, for its row repeats (see CellBudget.takeRepeats).
            let rowWeight = 0;

            for (const cell of tableCells) {
                const colsRepeated = toRepeatCount(cell.getAttribute("table:number-columns-repeated"));
                // A cell a merge covers holds its column: the next cell's is after it.
                if (cell.tagName === "table:covered-table-cell") {
                    colIndex += colsRepeated;
                    continue;
                }
                const cellChildren: OfficeContentNode[] = [];
                // Built without reading back what it holds, which would copy all of it per paragraph.
                const cellTextParts = new TextBuilder();
                const colSpan = cellSpan(cell.getAttribute("table:number-columns-spanned"), MAX_COL_SPAN);
                const rowSpan = cellSpan(cell.getAttribute("table:number-rows-spanned"), MAX_ROW_SPAN);

                // Helper to recursively process cell children (handles frames, text-boxes, etc. in ODP)
                // The cell style's text looks (bold, italic, size, colour, font), which its runs take under
                // their own: a spreadsheet sets a cell bold through its style. Its background is the cell's.
                const cellLooks = cellTextLooks(styleMap[cell.getAttribute("table:style-name") || '']);
                const processChildren = (node: Element) => {
                    if (!node.childNodes) return;

                    for (let i = 0; i < node.childNodes.length; i++) {
                        const child = node.childNodes[i];
                        if (isElement(child)) { // Element
                            const element = child as Element;

                            if (element.tagName === "text:p" || element.tagName === "text:h") {
                                const pContent = parseParagraphContent(element, paraStyleMap, styleMap, config, sourceXml, cellLooks);
                                let pNode: OfficeContentNode;
                                if (element.tagName === "text:h") {
                                    pNode = {
                                        type: 'heading',
                                        text: pContent.text,
                                        children: pContent.children,
                                        metadata: {
                                            level: parseInt(element.getAttribute("text:outline-level") || "1"),
                                            ...(pContent.alignment ? { alignment: pContent.alignment } : {}),
                                            ...(pContent.indentation ? { paragraphIndentation: pContent.indentation } : {}),
                                            ...(pContent.spacing ? { spacing: pContent.spacing } : {}),
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
                                            ...(pContent.indentation ? { paragraphIndentation: pContent.indentation } : {}),
                                            ...(pContent.spacing ? { spacing: pContent.spacing } : {}),
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
                            } else if (element.tagName === "draw:frame" || element.tagName === "draw:text-box" || CELL_CONTAINERS.has(element.tagName)) {
                                // Recursively process container content: frames (common in ODP), and a section,
                                // a list or an index the cell holds (their paragraphs were dropped).
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
                const cellWeight = cellBudget.weighCell(cellText, cellChildren, cellComments);
                const wantedCols = colsRepeated > 1 ? 1 + cellBudget.takeRepeats(colsRepeated - 1, cellWeight) : colsRepeated;
                const allowedCols = cellBudget.take(wantedCols);
                rowWeight += cellWeight * allowedCols;
                // Serialized once for all the repeats, which share it.
                const cellRaw = config.includeRawContent && allowedCols > 0 ? getRawContent(cell, sourceXml, config) : undefined;
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
                        cellNode.rawContent = k === 0 ? cellRaw : chargeRawContent(cellRaw, config);
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
                // Repeats past the content budget keep their columns, so later cells keep theirs.
                colIndex += colsRepeated - wantedCols;
            }

            // Add row(s) for repeated rows. rows x cols cells are materialized (charged against the
            // budget), but a repeated row shallow-clones each cell (own object + own metadata for the
            // row-index fix below) and SHARES the read-only children/comments/text by reference; JSON
            // deep-copying instead serialized every cell's full content per row, so a row of content-
            // bearing cells built a multi-hundred-MB string (a RangeError / OOM the count budget missed).
            // A row repeat copies every cell's content too (see CellBudget.takeRepeats).
            const wantedRows = rowsRepeated > 1 ? 1 + cellBudget.takeRepeats(rowsRepeated - 1, rowWeight) : rowsRepeated;
            const allowedRows = cells.length === 0
                ? (wantedRows > 0 ? 1 + cellBudget.take(wantedRows - 1) : 0)
                : Math.min(wantedRows,
                    1 + Math.floor(cellBudget.take(Math.max(0, (wantedRows - 1) * cells.length)) / cells.length));
            const rowRaw = config.includeRawContent && allowedRows > 0 ? getRawContent(row, sourceXml, config) : undefined;
            for (let k = 0; k < allowedRows; k++) {
                if ((k & 255) === 0) checkAbortSignal(config.abortSignal);
                const rowNode: OfficeContentNode = {
                    type: 'row',
                    children: k === 0 ? cells : cells.map(c => repeatCell(c, config))
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
                    rowNode.rawContent = k === 0 ? rowRaw : chargeRawContent(rowRaw, config);
                }

                rows.push(rowNode);
                rowIndex++;
            }
            rowIndex += rowsRepeated - wantedRows;
        }

        return {
            type: 'table',
            children: rows
        };
    };


    const parseContentXml = (xmlString: string) => {
        const xml = parseXmlString(xmlString, { config, locator: config.includeRawContent });
        const body = getFirstElementByTagName(xml, "office:body");
        if (!body) return;
        collectNoteIds(xml);
        // One budget for the entire document. It has to span every table - spreadsheet sheets,
        // ODT/ODP body tables, and nested tables alike - or a file sidesteps the cap simply by
        // splitting a huge repeat expansion across many small tables. `traverse` and the
        // spreadsheet branch below both close over this; `parseTable` receives it explicitly.
        const cellBudget = createCellBudget(config);

        // ── Where a slide's shapes are, and what they look like (ODP; see the PPTX reader's) ─────────
        /** A slide's fill: its own drawing-page style's, else its master page's. */
        const odpBackground = (page: Element, doc: Document): string | undefined => {
            const fillOf = (style: string | null) => styleProperty(doc, style, "style:drawing-page-properties", "draw:fill") === 'solid'
                ? styleProperty(doc, style, "style:drawing-page-properties", "draw:fill-color") : undefined;
            const master = masterPages.get(page.getAttribute("draw:master-page-name") || '');
            const color = fillOf(page.getAttribute("draw:style-name")) ?? fillOf(master?.getAttribute("draw:style-name") ?? null);
            return color && /^#[0-9a-f]{6}$/i.test(color) ? color.toUpperCase() : undefined;
        };
        /**
         * Reads a slide's shape with its box (`svg:x`, `svg:y`, `svg:width`, `svg:height`, or a turned
         * shape's `draw:transform`) on each node, and before them a `drawing` node where it is filled,
         * outlined, turned, or sets its text off its top (from its graphic style). A group's shapes each
         * have their own place.
         */
        const placeOdpShape = (element: Element, out: OfficeContentNode[], doc: Document, sourceXml: string): void => {
            if (element.tagName === 'draw:g') {
                for (const child of Array.from(element.childNodes).filter(isElement)) placeOdpShape(child, out, doc, sourceXml);
                return;
            }
            const nodes: OfficeContentNode[] = [];
            traverse(element, nodes, false, sourceXml);
            const width = odpLength(element.getAttribute("svg:width")), height = odpLength(element.getAttribute("svg:height"));
            let x = odpLength(element.getAttribute("svg:x")), y = odpLength(element.getAttribute("svg:y"));
            let rotation = 0;
            // `rotate (a) translate (x y)`: turned by a (radians, anticlockwise) about its corner, which
            // the translation then places; its box is about the same centre, unturned.
            const transform = /rotate\s*\(\s*(-?[\d.eE+-]+)\s*\)\s*translate\s*\(\s*([^\s)]+)\s+([^\s)]+)\s*\)/.exec(element.getAttribute("draw:transform") || '');
            if (transform && width !== undefined && height !== undefined) {
                const angle = Number(transform[1]), tx = odpLength(transform[2]), ty = odpLength(transform[3]);
                if (Number.isFinite(angle) && tx !== undefined && ty !== undefined) {
                    const cos = Math.cos(-angle), sin = Math.sin(-angle);
                    const cx = tx + cos * width / 2 - sin * height / 2, cy = ty + sin * width / 2 + cos * height / 2;
                    x = cx - width / 2; y = cy - height / 2;
                    rotation = ((Math.round(-angle * 180 / Math.PI * 100) / 100) % 360 + 360) % 360;
                }
            }
            if (x === undefined || y === undefined || width === undefined || height === undefined) { appendAll(out, nodes); return; }
            const bounds = { x: Math.round(x * 100) / 100, y: Math.round(y * 100) / 100, width: Math.round(width * 100) / 100, height: Math.round(height * 100) / 100 };
            for (const node of nodes) node.bounds = bounds;
            const style = element.getAttribute("draw:style-name") || element.getAttribute("presentation:style-name");
            const prop = (attr: string) => styleProperty(doc, style, "style:graphic-properties", attr);
            const isShape = element.tagName !== 'draw:frame';
            const geometryType = getFirstElementByTagName(element, "draw:enhanced-geometry")?.getAttribute("draw:type");
            // As DrawingML names it (ShapeMetadata.geometry); a shape of a path of its own, which no preset is, a rectangle.
            const geometry = element.tagName === 'draw:ellipse' || element.tagName === 'draw:circle' ? 'ellipse' : presetOfOdfShape(geometryType) ?? 'rect';
            const fill = isShape && prop("draw:fill") !== 'none' ? prop("draw:fill-color") : prop("draw:fill") === 'solid' ? prop("draw:fill-color") : undefined;
            const stroked = isShape ? prop("draw:stroke") !== 'none' : prop("draw:stroke") === 'solid';
            const strokeColor = stroked ? (prop("svg:stroke-color") ?? (isShape ? '#000000' : undefined)) : undefined;
            const anchor = prop("draw:textarea-vertical-align");
            const textAnchor = anchor === 'middle' ? 'middle' as const : anchor === 'bottom' ? 'bottom' as const : undefined;
            const metadata: ShapeMetadata = {
                ...(fill || strokeColor ? { geometry } : {}),
                ...(fill && /^#[0-9a-f]{6}$/i.test(fill) ? { fill: fill.toUpperCase() } : {}),
                ...(strokeColor && /^#[0-9a-f]{6}$/i.test(strokeColor) ? { line: { color: strokeColor.toUpperCase(), width: odpLength(prop("svg:stroke-width") ?? null) || 0.75 } } : {}),
                ...(rotation ? { rotation } : {}),
                ...(nodes.length && textAnchor ? { textAnchor } : {}),
            };
            if (metadata.fill || metadata.line || (nodes.length && (rotation || textAnchor))) {
                if (nodes.length) metadata.textInsets = { top: odpLength(prop("fo:padding-top") ?? null) ?? 3.6, right: odpLength(prop("fo:padding-right") ?? null) ?? 7.2, bottom: odpLength(prop("fo:padding-bottom") ?? null) ?? 3.6, left: odpLength(prop("fo:padding-left") ?? null) ?? 7.2 };
                out.push({ type: 'drawing', text: '', bounds, metadata });
            }
            appendAll(out, nodes);
        };
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
            const info = styleName ? paragraphStyleMap[styleName] : undefined;
            // A paragraph of a style naming a master page starts a page of it: the first block's sets the
            // document's first pages, a later one's starts a section of its own (with its page).
            if (edge === 'before' && readingBody) {
                const master = info?.masterPage && masterPages.has(info.masterPage) ? info.masterPage : undefined;
                // (The master page in force named again, one that hands over to no other: a new page of it, as a
                // word processor goes on, not a section with a first page of its own.)
                const inForce = sectionMasters[sectionMasters.length - 1] ?? firstMaster;
                const handsOver = (name: string) => { const next = masterPages.get(name)?.getAttribute('style:next-style-name'); return !!next && next !== name; };
                if (!bodyStarted) {
                    bodyStarted = true;
                    if (master) firstMaster = master;
                } else if (master && master === inForce && !handsOver(master)) {
                    if (config.includeBreakNodes) targetArray.push({ type: 'break', metadata: { breakType: 'page' } as BreakMetadata });
                    return;
                } else if (master) {
                    sectionMasters.push(master);
                    // (A section's start is structure, a node whatever includeBreakNodes says.)
                    {
                        const pageSetup = odfPageSetup(master);
                        // A page that starts inside a section of columns (LibreOffice turns to a section's page
                        // on its first paragraph) is set in that section's columns: the section's own break,
                        // just before it with nothing between, is this one.
                        const inSection = openSectionColumns[openSectionColumns.length - 1];
                        if (pageSetup?.width && inSection) {
                            const columns = sectionColumnsOn(inSection, pageSetup);
                            if (columns) pageSetup.columns = columns; else delete pageSetup.columns;
                            const before = targetArray[targetArray.length - 1];
                            if (before?.type === 'break' && (before.metadata as BreakMetadata | undefined)?.breakType === 'section' && (before.metadata as any)?.openedSection) targetArray.pop();
                        }
                        targetArray.push({ type: 'break', metadata: { breakType: 'page', section: sectionMasters.length, ...(pageSetup ? { pageSetup } : {}) } as BreakMetadata });
                    }
                    return;
                }
            }
            if (!config.includeBreakNodes || !styleName) return;
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
                    ...(pContent.indentation ? { paragraphIndentation: pContent.indentation } : {}),
                    ...(pContent.spacing ? { spacing: pContent.spacing } : {}),
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
                    ...(hContent.indentation ? { paragraphIndentation: hContent.indentation } : {}),
                    ...(hContent.spacing ? { spacing: hContent.spacing } : {}),
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
                // (A table's style starts a page, or a page style, as a paragraph's does: LibreOffice's way
                // of turning to a page style at a table.)
                if (!asSheet) pushStyleBreak(node.getAttribute("table:style-name"), targetArray, 'before');
                targetArray.push(tableNode);
                if (!asSheet) pushStyleBreak(node.getAttribute("table:style-name"), targetArray, 'after');
                lastWasList = false;
            } else if (node.tagName === "text:list") {
                // Parse list structure with proper listId tracking. A list header (`text:list-header`)
                // is a list entry without a number: its paragraphs are read as paragraphs, in place.
                const listItems = Array.from(node.childNodes).filter(isElement).filter(e => e.tagName === "text:list-item" || e.tagName === "text:list-header");

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
                    const automatic = automaticStyles ? listStyleKind(automaticStyles, styleNameToCheck, true) : undefined;
                    if (automatic) ({ listType, isVisible } = automatic);

                    if (!isVisible && stylesDom) {
                        officeStylesElement ??= getFirstElementByTagName(stylesDom, "office:styles") ?? null;
                        const office = officeStylesElement ? listStyleKind(officeStylesElement, styleNameToCheck, false) : undefined;
                        if (office) ({ listType, isVisible } = office);
                    }
                }

                // This level's own style (a numbered list's nested level may be bulleted, and the
                // reverse), for its items' kind and marker.
                let depth = 0;
                for (let parent = node.parentNode; parent; parent = parent.parentNode) if (parent.nodeName === 'text:list') depth++;
                const levels = styleNameToCheck ? odfListLevels(automaticStyles ?? undefined, styleNameToCheck) : undefined;
                const ownLevel = levels?.get(depth + 1);
                if (isVisible && ownLevel) listType = ownLevel.tagName === 'text:list-level-style-number' ? 'ordered' : 'unordered';

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
                    if (item.tagName === "text:list-header") {
                        for (const child of Array.from(item.childNodes).filter(isElement)) traverse(child, targetArray, forceHeading, sourceXml);
                        lastWasList = true;
                        continue;
                    }

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
                                    const marker = levels ? odfListMarker(levels, indentation, listCounters[listId], listType, itemIndex) : undefined;
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
                                                ...(marker !== undefined ? { marker } : {}),
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
                const textBox = frameChild(node, "draw:text-box");
                const image = frameChild(node, "draw:image");
                const table = frameChild(node, "table:table");
                const object = frameChild(node, "draw:object");

                if (textBox) {
                    traverse(textBox, targetArray, isHeading || forceHeading, sourceXml);
                } else if (table) {
                    const tableNode = parseTable(table, paragraphStyleMap, styleMap, config, sourceXml, cellBudget);
                    if (config.includeRawContent) tableNode.rawContent = getRawContent(table, sourceXml, config);
                    targetArray.push(tableNode);
                } else if (image) {
                    // Extract alt text from svg:title or svg:desc
                    let altText = '';
                    const svgTitle = frameChild(node, "svg:title");
                    const svgDesc = frameChild(node, "svg:desc");
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
                        const embedded = readEmbeddedObject(attachmentName);

                        if (embedded) {
                            if (embedded.formula !== undefined) {
                                // Math formula object at block level - a display equation, so the
                                // inner node is `math: 'block'` where the inline site above emits
                                // `math: 'inline'`.
                                const formulaText = embedded.formula;
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
                                const chartData = embedded.chartData!;

                                const chartNode: OfficeContentNode = {
                                    type: 'chart',
                                    text: embedded.chartText ?? '',
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
            } else if (node.tagName === "text:section" && readingBody && config.includeBreakNodes && !config.ignorePageGeometry
                && sectionColumns.get(node.getAttribute("text:style-name") || '') && odfPageSetup(sectionMasters[sectionMasters.length - 1] ?? firstMaster)?.width) {
                // A section set in columns of its own, on the page it is on: a section break before it
                // with its columns, and one after it with the page's again (the page it ends on: a page break
                // inside it turns to another).
                const styleName = node.getAttribute("text:style-name") || '';
                const { columns: _pageColumns, ...page } = odfPageSetup(sectionMasters[sectionMasters.length - 1] ?? firstMaster)!;
                const opened: OfficeContentNode = { type: 'break', metadata: { breakType: 'section', pageSetup: { ...page, columns: sectionColumnsOn(styleName, page) } } as BreakMetadata };
                // (Marked while it is the last node, so a page break at the section's start takes its place.)
                (opened.metadata as any).openedSection = true;
                targetArray.push(opened);
                openSectionColumns.push(styleName);
                try {
                    for (let i = 0; i < node.childNodes.length; i++) {
                        const child = node.childNodes[i];
                        if (isElement(child)) traverse(child as Element, targetArray, forceHeading, sourceXml);
                    }
                } finally {
                    openSectionColumns.pop();
                    delete (opened.metadata as any).openedSection;
                }
                // (The page it ends on may be a master of no layout: the page's setup then, columns none.)
                const { columns: endColumns, ...endPage } = odfPageSetup(sectionMasters[sectionMasters.length - 1] ?? firstMaster) ?? {};
                targetArray.push({ type: 'break', metadata: { breakType: 'section', pageSetup: { ...endPage, ...(endColumns ? { columns: endColumns } : {}) } } as BreakMetadata });
            } else if (node.tagName !== "text:tracked-changes") {
                // (Tracked changes hold what a deletion removed, which is not the document's text.)
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
                // The sheets, rows, cells and paragraphs each level holds itself (see getOutermostElements):
                // taken as descendants, a table nested in a cell was read again at each level (586
                // bytes took 7 seconds), and became a sheet of its own.
                const tables = getDirectChildren(spreadsheet, "table:table");
                // The table styles that hide their sheet (`table:display="false"`).
                const hiddenTableStyles = new Set<string>();
                const contentStyles = getFirstElementByTagName(xml, "office:automatic-styles");
                if (contentStyles) {
                    for (const style of getDirectChildren(contentStyles, "style:style")) {
                        if (style.getAttribute("style:family") !== 'table') continue;
                        if (getDirectChildren(style, "style:table-properties")[0]?.getAttribute("table:display") === 'false') hiddenTableStyles.add(style.getAttribute("style:name") || '');
                    }
                }
                const isHiddenVisibility = (value: string | null) => value === 'collapse' || value === 'filter';
                for (let i = 0; i < tables.length; i++) {
                    const table = tables[i];
                    const sheetName = table.getAttribute("table:name") || `Sheet${i + 1}`;
                    const rows: OfficeContentNode[] = [];
                    const sheetHidden = hiddenTableStyles.has(table.getAttribute("table:style-name") || '');
                    // Hidden columns and rows are read, and marked (CellMetadata.hidden): a printout leaves
                    // them out, as the PDF writer does, and every other output keeps them as data. A
                    // column's repeats are ranges; the rows (where nested tables live) are not searched.
                    const hiddenColumns: [number, number][] = [];
                    let columnAt = 0;
                    for (const column of getOutermostElements(table, "table:table-column", COLUMN_SKIP)) {
                        const repeat = toRepeatCount(column.getAttribute("table:number-columns-repeated"));
                        if (isHiddenVisibility(column.getAttribute("table:visibility"))) hiddenColumns.push([columnAt, columnAt + repeat - 1]);
                        columnAt += repeat;
                    }
                    const columnHidden = inRanges(hiddenColumns);

                    const tableRows = getOutermostElements(table, "table:table-row");
                    let rowIndex = 0;

                    for (let r = 0; r < tableRows.length; r++) {
                        checkAbortSignal(config.abortSignal);
                        const row = tableRows[r];
                        const cells: OfficeContentNode[] = [];
                        const tableCells = rowCells(row);

                        let colIndex = 0;
                        const rowHidden = isHiddenVisibility(row.getAttribute("table:visibility"));
                        const rowsRepeated = toRepeatCount(row.getAttribute("table:number-rows-repeated"));
                        // What one copy of this row writes, for its row repeats (see CellBudget.takeRepeats).
                        let rowWeight = 0;

                        for (let c = 0; c < tableCells.length; c++) {
                            const cell = tableCells[c];
                            const colsRepeated = toRepeatCount(cell.getAttribute("table:number-columns-repeated"));
                            // A cell a merge covers holds its column (see rowCells); a sheet's that holds a value (a
                            // spreadsheet keeps one there, hidden, as XLSX's are read) is read at its place, the
                            // merge's: it was left out.
                            if (cell.tagName === "table:covered-table-cell" && !(fileType === 'ods' && holdsValue(cell))) {
                                colIndex += colsRepeated;
                                continue;
                            }
                            const colSpan = cellSpan(cell.getAttribute("table:number-columns-spanned"), MAX_COL_SPAN);
                            const rowSpan = cellSpan(cell.getAttribute("table:number-rows-spanned"), MAX_ROW_SPAN);

                            // ODS cell notes are `<office:annotation>` children of the cell; extract them
                            // as comments. Their inner text:p must be kept out of the cell's own text
                            // below (getElementsByTagName is descendant-recursive, so it would otherwise
                            // pull the comment body into the cell value).
                            const cellComments = config.ignoreComments
                                ? []
                                : getDirectChildren(cell, "office:annotation").map(a =>
                                    buildAnnotationComment(a, parseParagraphContent, paragraphStyleMap, styleMap, config, xmlString));
                            // Extract text from cell (paragraphs inside cell, excluding a comment's body)
                            let cellText = "";
                            const children: OfficeContentNode[] = [];
                            const odsCellLooks = cellTextLooks(styleMap[cell.getAttribute("table:style-name") || '']);
                            const ps = getOutermostElements(cell, "text:p", CELL_SKIP);
                            for (let p = 0; p < ps.length; p++) {
                                const para = ps[p];

                                // The paragraph's inline content as any paragraph's: text around a formatted
                                // span, spaces, tabs, line breaks, links and fields. Taking only the spans'
                                // text when there were any dropped the rest ("Total: " before a bold "5").
                                // (Its frames are read with the cell's, below.)
                                // Its runs take the cell style's looks (a spreadsheet sets a cell bold through its style)
                                // and the paragraph style's under their own; the cell's background is the cell's.
                                const paraStyle = para.getAttribute("text:style-name");
                                const inline = parseInlineContent(para, styleMap, config, notes, paragraphStyleMap, mergeFormatting(mergeFormatting({}, odsCellLooks), paraStyle ? styleMap[paraStyle] : undefined), undefined, xmlString, false);
                                cellText += inline.text;
                                const shown = inline.text.trim() !== '';
                                // A paragraph after another is a line of its own: its runs followed the
                                // last one's, so a cell of "Line one" and "Line two" read "Line oneLine two".
                                const kept = inline.children.filter(child => shown || child.notes || child.comments);
                                if (kept.length > 0 && children.length > 0) children.push({ type: 'break', metadata: { breakType: 'textWrapping' } as BreakMetadata });
                                for (const child of kept) children.push(child);

                                if (p < ps.length - 1) cellText += "\n";
                            }

                            // Check for embedded draw:frame (images) in cell
                            // Each frame on its own (nested ones too), none reading another: all of them.
                            const drawFrames = getAllElementsByTagName(cell, "draw:frame");
                            for (const frame of drawFrames) {
                                // Extract alt text from svg:title or svg:desc
                                let altText = '';
                                const svgTitle = frameChild(frame, "svg:title");
                                const svgDesc = frameChild(frame, "svg:desc");
                                if (svgTitle && svgTitle.textContent) {
                                    altText = svgTitle.textContent;
                                } else if (svgDesc && svgDesc.textContent) {
                                    altText = svgDesc.textContent;
                                }

                                // Extract image href
                                let imageHref = '';
                                const drawImages = getDirectChildren(frame, "draw:image");
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
                                const drawObjects = getDirectChildren(frame, "draw:object");
                                if (drawObjects.length > 0) {
                                    const href = drawObjects[0].getAttribute("xlink:href");
                                    if (href) {
                                        chartHref = cleanAttachmentName(href);
                                        const embedded = readEmbeddedObject(chartHref);
                                        if (embedded?.formula !== undefined) {
                                            isFormula = true;
                                            formulaText = embedded.formula;
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
                            // What a sheet's cell holds, as its file says (`text` is what the sheet shows: 10.46, 1.23E+05).
                            const held = fileType === 'ods' && willMaterialize ? heldValue(cell) : undefined;
                            if (!willMaterialize) {
                                colIndex += colsRepeated;
                            } else {
                                const cellWeight = cellBudget.weighCell(cellText, children, cellComments);
                                const wantedCols = colsRepeated > 1 ? 1 + cellBudget.takeRepeats(colsRepeated - 1, cellWeight) : colsRepeated;
                                const allowedCols = cellBudget.take(wantedCols);
                                rowWeight += cellWeight * allowedCols;
                                // Serialized once for all the repeats, which share it.
                                const cellRaw = config.includeRawContent && allowedCols > 0 ? getRawContent(cell, xmlString, config) : undefined;
                                for (let k = 0; k < allowedCols; k++) {
                                    if ((k & 1023) === 0) checkAbortSignal(config.abortSignal);
                                    const cellNode: OfficeContentNode = {
                                        type: 'cell',
                                        text: cellText,
                                        children: children,
                                        metadata: {
                                            ...held,
                                            row: rowIndex,
                                            col: colIndex,
                                            // Merged cells: the spans were not read, so the merge was lost.
                                            ...(colSpan > 1 ? { colSpan } : {}),
                                            ...(rowSpan > 1 ? { rowSpan } : {}),
                                            ...(rowHidden || columnHidden(colIndex) ? { hidden: true } : {}),
                                        } as CellMetadata
                                    };
                                    if (config.includeRawContent) {
                                        cellNode.rawContent = k === 0 ? cellRaw : chargeRawContent(cellRaw, config);
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
                                // Repeats past the content budget keep their columns, so later cells keep theirs.
                                colIndex += colsRepeated - wantedCols;
                            }
                        }

                        // Add row(s). This is where the two repeats multiply: each repetition
                        // deep-copies the whole cell array, so rows x cols is what actually
                        // exhausts memory. Charge the copies against the same budget.
                        if (cells.length > 0) {
                            // A row repeat copies every cell's content too (see CellBudget.takeRepeats).
                            const wantedRows = rowsRepeated > 1 ? 1 + cellBudget.takeRepeats(rowsRepeated - 1, rowWeight) : rowsRepeated;
                            const allowedRows = Math.min(
                                wantedRows,
                                // The first row reuses `cells` rather than copying, so only the
                                // repeats beyond it cost budget.
                                1 + Math.floor(cellBudget.take(Math.max(0, (wantedRows - 1) * cells.length)) / cells.length)
                            );
                            const rowRaw = config.includeRawContent && allowedRows > 0 ? getRawContent(row, xmlString, config) : undefined;
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
                                    : cells.map(c => repeatCell(c, config));
                                const rowNode: OfficeContentNode = {
                                    type: 'row',
                                    children: rowCells,
                                    metadata: rowHidden ? { hidden: true } : undefined
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
                                    rowNode.rawContent = k === 0 ? rowRaw : chargeRawContent(rowRaw, config);
                                }
                                rows.push(rowNode);
                                rowIndex++;
                            }
                            rowIndex += rowsRepeated - wantedRows;
                        } else {
                            rowIndex += rowsRepeated;
                        }
                    }

                    // How the sheet prints: its table style's master page's pages.
                    // (A sheet whose style names none prints on the Default page style, as Calc prints it.)
                    const sheetMaster = paragraphStyleMap[table.getAttribute("table:style-name") || '']?.masterPage ?? (masterPages.has('Default') ? 'Default' : firstMaster);
                    // (A sheet's print setup is its page: none where it gives no size.)
                    const sheetPage = config.ignorePageGeometry ? undefined : sized(odfPageSetup(sheetMaster));
                    const sheetNode: OfficeContentNode = {
                        type: 'sheet',
                        children: rows,
                        metadata: { sheetName, ...(sheetHidden ? { hidden: true } : {}), ...(sheetPage ? { pageSetup: sheetPage } : {}) } as SheetMetadata
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
                                if (config.ignorePageGeometry) traverse(element, slideNode.children!, false, xmlString);
                                else placeOdpShape(element, slideNode.children!, xml, xmlString);
                            }
                        }
                    }
                    if (!config.ignorePageGeometry) {
                        const background = odpBackground(page, xml);
                        if (background) (slideNode.metadata as SlideMetadata).background = background;
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
                    appendAll(content, odpNotes);
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
                readingBody = true;
                try {
                    traverse(textDoc, content, false, xmlString);
                } finally {
                    readingBody = false;
                }
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
            const objXml = parseXmlString(objFile.content.toString(), { config });
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
                const chartData = extractChartData(objFile.content, config);

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
    const metadata = metaFile ? parseOfficeMetadata(metaFile.content.toString(), config) : {};

    // Helper: Resolve ODS chart cell references to actual values
    // ODS charts often link to cell ranges (e.g., [Sheet1.$A$1:.$A$5]) instead of embedding values
    // Each sheet's cells by row (built once), the references already resolved, and a budget of cells
    // read and values made for all the document's charts: every series scanned the whole sheet and
    // copied what it matched, so a chart of many series over one range (1.8 KB of ODS) took gigabytes.
    // Past the budget, a reference stays its `[range]` text, as one that resolves to nothing does.
    const sheetRows = new Map<OfficeContentNode, { rows: number[]; cells: Map<number, OfficeContentNode[]> }>();
    const resolvedReferences = new Map<string, string[]>();
    let resolutionBudget = 5 * MAX_CHART_VALUES;
    // The sheets by name (the first of a name), found once: a scan of the document per reference took
    // references x sheets (315 KB, three and a half minutes).
    let sheetsByName: Map<string, OfficeContentNode> | undefined;
    const sheetNamed = (nodes: OfficeContentNode[], name: string): OfficeContentNode | undefined => {
        if (!sheetsByName) {
            sheetsByName = new Map();
            for (const n of nodes) {
                const sheetName = n.type === 'sheet' ? (n.metadata as SheetMetadata)?.sheetName : undefined;
                if (typeof sheetName === 'string' && !sheetsByName.has(sheetName)) sheetsByName.set(sheetName, n);
            }
        }
        return sheetsByName.get(name);
    };
    const rowsOf = (sheet: OfficeContentNode) => {
        let index = sheetRows.get(sheet);
        if (!index) {
            const cells = new Map<number, OfficeContentNode[]>();
            for (const row of sheet.children ?? []) for (const cell of row.children ?? []) {
                const r = (cell.metadata as CellMetadata | undefined)?.row;
                if (typeof r !== 'number') continue;
                let list = cells.get(r);
                if (!list) cells.set(r, list = []);
                list.push(cell);
            }
            index = { rows: [...cells.keys()].sort((a, b) => a - b), cells };
            sheetRows.set(sheet, index);
        }
        return index;
    };
    const resolveChartReferences = (chartData: ChartData, nodes: OfficeContentNode[]) => {
        const getValuesFromReference = (ref: string): string[] => {
            const known = resolvedReferences.get(ref);
            if (known) return known;
            const values = resolveReference(ref);
            resolvedReferences.set(ref, values);
            return values;
        };
        const resolveReference = (ref: string): string[] => {
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

            // Every reference is charged, found or not.
            if (--resolutionBudget < 0) return [ref];
            const sheet = sheetNamed(nodes, sheetName);
            if (!sheet || !sheet.children) return [ref];

            const values: string[] = [];
            // The cells of the rows in range, found through the sheet's row index.
            const { rows, cells } = rowsOf(sheet);
            let lo = 0, hi = rows.length;
            while (lo < hi) { const mid = (lo + hi) >> 1; if (rows[mid] < start.r) lo = mid + 1; else hi = mid; }
            for (let i = lo; i < rows.length && rows[i] <= end.r; i++) {
                for (const cell of cells.get(rows[i])!) {
                    if (--resolutionBudget < 0) return [ref];
                    const col = (cell.metadata as CellMetadata).col;
                    if (col >= start.c && col <= end.c) values.push(cell.text || '');
                }
            }
            return values;
        };

        // Resolve DataSets and labels: a resolved range is shared, and each value counts against the budget.
        const resolve = (list: string[]): string[] => {
            const out: string[] = [];
            for (const val of list) {
                const values = val.startsWith('[') ? getValuesFromReference(val) : [val];
                if (values.length > 1 && (resolutionBudget -= values.length) < 0) { out.push(val); continue; }
                for (const v of values) out.push(v);
            }
            return out;
        };
        for (const ds of chartData.dataSets) ds.values = resolve(ds.values);
        chartData.labels = resolve(chartData.labels);
        chartData.rawTexts = chartRawTexts(chartData.title, chartData.dataSets, chartData.labels);
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
        // Both walks visit a node once: repeated cells share their content nodes, so walking every
        // path to them took repeats x content steps (a few hundred KB held minutes of CPU).
        // Step 1: Identify unused image attachments globally
        const usedAttachmentNames = new Set<string>();
        const named = new Set<OfficeContentNode>();
        const traverseForNames = (ns: OfficeContentNode[]) => {
            for (const n of ns) {
                if (named.has(n)) continue;
                named.add(n);
                if (n.metadata && 'attachmentName' in n.metadata) {
                    const name = (n.metadata as ImageMetadata).attachmentName;
                    if (name) usedAttachmentNames.add(name);
                }
                if (n.children) traverseForNames(n.children);
            }
        };
        traverseForNames(nodes);
        const attachmentsByName = new Map<string, OfficeAttachment>();
        for (const a of attachments) if (a.name && !attachmentsByName.has(a.name)) attachmentsByName.set(a.name, a);
        // A chart's text, joined once for every node showing it.
        const chartTexts = new Map<OfficeAttachment, string>();
        const chartTextOf = (attachment: OfficeAttachment): string => {
            let text = chartTexts.get(attachment);
            if (text === undefined) chartTexts.set(attachment, text = attachment.chartData!.rawTexts.join(config.newlineDelimiter));
            return text;
        };

        const unusedImages = attachments.filter(a => a.type === 'image' && a.name && !usedAttachmentNames.has(a.name));
        let unusedImageIndex = 0;

        const processed = new Set<OfficeContentNode>();
        const processNode = (node: OfficeContentNode) => {
            if (processed.has(node)) return;
            processed.add(node);
            if ((node.type === 'image' || node.type === 'chart') && node.metadata && 'attachmentName' in node.metadata) {
                let attachmentName = (node.metadata as ImageMetadata).attachmentName;

                // Heuristic: If name is empty, try to assign an unused image attachment
                if (!attachmentName && node.type === 'image' && unusedImageIndex < unusedImages.length) {
                    const fallbackAtt = unusedImages[unusedImageIndex++];
                    attachmentName = fallbackAtt.name;
                    (node.metadata as ImageMetadata).attachmentName = attachmentName;
                }

                if (attachmentName) {
                    const attachment = attachmentsByName.get(attachmentName);

                    if (attachment) {
                        if (attachment.ocrText) {
                            node.text = attachment.ocrText;
                        }
                        if (attachment.chartData && node.type === 'chart') {
                            node.text = chartTextOf(attachment);
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
        // (A style that only names a master page, as a sheet's table style does, has nothing to show.)
        if (!combinedStyleMap[styleName] && Object.keys(paragraphStyleMap[styleName]).some(key => key !== 'masterPage')) {
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
    if (!config.ignoreHeadersAndFooters && stylesDom && fileType === 'odt' && firstMaster) {
        const stylesXmlStr = stylesFile ? stylesFile.content.toString() : undefined;
        const headers: OfficeContentNode[] = [];
        const footers: OfficeContentNode[] = [];
        // Each section's master page's headers and footers, as header and footer nodes of their kind
        // (default; even for a left page's; first): the first section's, and each later one's whose
        // master page is another (an empty one where it has none, which a section has not carried on).
        // Only the first master page's two were read, as loose paragraphs.
        const masters = [firstMaster, ...sectionMasters];
        const sectioned = new Set(masters).size > 1;
        // Whether any of the document's master pages (or one a first page's hands over to) has a left part of
        // its own: then every section has even pages' parts, as Word's even and odd pages are the document's.
        const facing = [...new Set(masters)].some(name => {
            const master = masterPages.get(name), next = master?.getAttribute('style:next-style-name');
            return [master, next && next !== name ? masterPages.get(next) : undefined].some(m => shown(m, 'style:header-left') || shown(m, 'style:footer-left'));
        });
        // Each part read once and its blocks shared by the sections showing it, a node only where a section's
        // part of a kind and type is other than the one in force (that section takes it on), and each node past
        // a part's first charged to the repeated content budget, going without its blocks once that is spent:
        // sections alternating two master pages read their parts again for each section, gigabytes for a
        // document of kilobytes.
        const readParts = new Map<Element, { children: OfficeContentNode[]; text: string; weight: number }>();
        const inForce = new Map<string, Element | null>();
        masters.forEach((name, section) => {
            if (section > 0 && name === masters[section - 1]) return;
            const master = masterPages.get(name);
            const own = [['style:header', 'header', 'default'], ['style:header-left', 'header', 'even'], ['style:header-first', 'header', 'first'],
                ['style:footer', 'footer', 'default'], ['style:footer-left', 'footer', 'even'], ['style:footer-first', 'footer', 'first']] as const;
            // A section's master page that hands over to another after its first page (a word processor's
            // "First Page" style, `style:next-style-name`) holds the first page's header and footer, and
            // the one it hands over to the rest's: read as every page's, its letterhead was repeated on
            // every page (a later section's too, as ODT output writes one).
            const next = master?.getAttribute('style:next-style-name');
            const rest = next && next !== name ? masterPages.get(next) : undefined;
            const parts: (readonly [Element | undefined, string, 'header' | 'footer', 'default' | 'even' | 'first'])[] = rest
                ? [...(['header', 'footer'] as const).map(kind => [master, shown(master, `style:${kind}-first`) ? `style:${kind}-first` : `style:${kind}`, kind, 'first'] as const),
                    ...own.filter(([, , type]) => type !== 'first').map(([tag, kind, type]) => [rest, tag, kind, type] as const)]
                : own.map(([tag, kind, type]) => [master, tag, kind, type] as const);
            // A left or first part ODF leaves out (or hides) is the right page's, where Word's (which the
            // header and footer nodes follow) shows nothing: where the document has left parts (any master
            // page's), each kind's default part is its left part too where it has none; where this master page
            // has a first page's part of one kind, the other kind's default part is its first page's too. They
            // were lost on those pages, and a later section on this master page showed another's left part.
            // (A first page's master page that hands over shows its own parts alone.)
            const found = parts.map(([from, tag, kind, type]) => [shown(from, tag), kind, type] as const);
            for (const type of ['even', 'first'] as const) {
                if (type === 'even' ? !facing : rest || !found.some(([el, , t]) => el && t === type)) continue;
                for (const kind of ['header', 'footer'] as const) {
                    const index = found.findIndex(([, k, t]) => k === kind && t === type);
                    const fallback = found.find(([el, k, t]) => el && k === kind && t === 'default')?.[0];
                    if (index >= 0 && !found[index][0] && fallback) found[index] = [fallback, kind, type];
                }
            }
            for (const [el, kind, type] of found) {
                if (!el && !(sectioned && section > 0 && type === 'default')) continue;
                const slot = `${kind}|${type}`;
                if (section > 0 && inForce.has(slot) && inForce.get(slot) === (el ?? null)) continue;
                inForce.set(slot, el ?? null);
                let part = el ? readParts.get(el) : undefined;
                const repeat = !!part;
                if (el && !part) {
                    const children: OfficeContentNode[] = [];
                    traverse(el, children, false, stylesXmlStr);
                    // Resolve any image/chart nodes the same way the body's are (OCR text, chart data,
                    // attachment linkage): without it a picture in a header would render blank.
                    if (children.length) assignAttachmentData(children);
                    const text = children.map(c => c.text || '').join('\n');
                    let count = 0;
                    const stack = [...children];
                    while (stack.length) { const n = stack.pop()!; count++; for (const c of n.children ?? []) stack.push(c); }
                    readParts.set(el, part = { children, text, weight: text.length + 16 * count });
                }
                const kept = part && (!repeat || takeRepeats(config, 1, part.weight) > 0);
                (kind === 'header' ? headers : footers).push({ type: kind, text: kept ? part!.text : '', children: kept ? part!.children : [],
                    metadata: { type, ...(sectioned ? { section } : {}) } as HeaderFooterMetadata });
            }
        });
        if (headers.length || footers.length) {
            auxiliary = {
                ...(headers.length ? { headers } : {}),
                ...(footers.length ? { footers } : {}),
            };
        }
    }
    // Frames floating beside a paragraph's text, read where they are anchored, stand before it.
    liftFloatingFrames(content);
    for (const part of [...(auxiliary?.headers ?? []), ...(auxiliary?.footers ?? [])]) if (part.children) liftFloatingFrames(part.children);
    // (A section starting after the last text, or a section of columns ending the document, starts nothing.)
    if (fileType === 'odt') dropTrailingSectionStarts(content);
    // A presentation's page is its slide's size alone: slides are laid out as they are, not within margins.
    // (A text document's numbering and first page are no geometry: kept under ignorePageGeometry.)
    const firstPage = config.ignorePageGeometry && fileType !== 'odt' ? undefined : odfPageSetup(firstMaster);
    const pageSetup = fileType === 'odt' ? firstPage : fileType === 'odp' ? (firstPage?.width ? { width: firstPage.width, height: firstPage.height } : undefined) : sized(firstPage);

    return createAST(
        fileType,
        {
            ...metadata,
            ...(pageSetup ? { pageSetup } : {}),
            styleMap: plainRecord(combinedStyleMap)
        },
        content,
        attachments,
        config,
        auxiliary,
    );
};

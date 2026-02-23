/**
 * Configuration options for the OfficeParser.
 */
export interface OfficeParserConfig {
    /**
     * Flag to show all the logs to console in case of an error irrespective of your own handling.
     * Default is false.
     */
    outputErrorToConsole?: boolean;
    /**
     * The delimiter used for every new line in places that allow multiline text like word.
     * Default is \n.
     */
    newlineDelimiter?: string;
    /**
     * Flag to ignore notes from parsing in files like powerpoint.
     * Default is false. It includes notes in the parsed text by default.
     */
    ignoreNotes?: boolean;
    /**
     * Flag, if set to true, will collectively put all the parsed text from notes at last in files like powerpoint.
     * Default is false. It puts each notes right after its main slide content.
     * If ignoreNotes is set to true, this flag is also ignored.
     * @note This flag currently does not affect RTF files; RTF footnotes/endnotes are always collected and appended at the end of the content.
     */
    putNotesAtLast?: boolean;
    /**
     * Flag to extract attachments like images, charts, etc.
     * Default is false.
     */
    extractAttachments?: boolean;
    /**
     * Flag to include raw content (XML for XML-based formats, RTF for RTF) in the AST.
     * Default is false.
     */
    includeRawContent?: boolean;
    /**
     * Flag to enable OCR for images.
     * Default is false.
     */
    ocr?: boolean;
    /**
     * Language for OCR.
     * Default is 'eng'.
     * 
     * You can provide multiple languages separated by a `+` sign (e.g., 'eng+fra' for English and French).
     * The OCR engine will then attempt to recognize text in any of the specified languages.
     * 
     * See the list of supported languages and their codes here:
     * https://tesseract-ocr.github.io/tessdoc/Data-Files#data-files-for-version-400-november-29-2016
     */
    ocrLanguage?: string;
    /**
     * The URL/path to the PDF.js worker script.
     * 
     * **Mandatory** when using PDF parsing in browser environments to avoid worker configuration errors.
     * If not provided, it defaults to `https://unpkg.com/pdfjs-dist@5.4.530/build/pdf.worker.min.mjs`.
     * You can override this with your own local path or a different CDN link.
     */
    pdfWorkerSrc?: string;
}

/**
 * Supported file types for parsing.
 */
export type SupportedFileType = 'docx' | 'pptx' | 'xlsx' | 'odt' | 'odp' | 'ods' | 'pdf' | 'rtf';

/**
 * Types of content nodes in the AST.
 */
export type OfficeContentNodeType = 'paragraph' | 'heading' | 'table' | 'list' | 'text' | 'image' | 'chart' | 'drawing' | 'slide' | 'note' | 'sheet' | 'row' | 'cell' | 'page';

/**
 * Supported MIME types for attachments.
 */
export type OfficeMimeType =
    | 'image/jpeg'
    | 'image/png'
    | 'image/gif'
    | 'image/bmp'
    | 'image/tiff'
    | 'image/svg+xml'
    | 'application/pdf'
    | 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    | 'application/vnd.oasis.opendocument.chart'
    | 'application/vnd.oasis.opendocument.spreadsheet'
    | 'application/vnd.oasis.opendocument.text'
    | 'application/vnd.oasis.opendocument.presentation';

/**
 * Text formatting options available for text content.
 * Represents common formatting attributes found in office documents (DOCX, RTF, PPTX, etc.).
 * All properties are optional and only present when the formatting is explicitly applied.
 */
export interface TextFormatting {
    /**
     * Whether the text is bold.
     * Corresponds to `<w:b/>` in OOXML, `\b` in RTF.
     * @example true for **bold text**, false or undefined for normal weight
     */
    bold?: boolean;

    /**
     * Whether the text is italic.
     * Corresponds to `<w:i/>` in OOXML, `\i` in RTF.
     * @example true for *italic text*, false or undefined for normal style
     */
    italic?: boolean;

    /**
     * Whether the text is underlined.
     * Corresponds to `<w:u/>` in OOXML, `\ul` in RTF.
     * @example true for underlined text, false or undefined for no underline
     */
    underline?: boolean;

    /**
     * Whether the text has a strikethrough.
     * Corresponds to `<w:strike/>` in OOXML, `\strike` in RTF.
     * @example true for ~~struck through~~ text
     */
    strikethrough?: boolean;

    /**
     * Text color in hex format (#RRGGBB).
     * Extracted from color tables in RTF or XML color attributes in OOXML.
     * @example "#ff0000" for red, "#00ff00" for green, "#0000ff" for blue
     */
    color?: string;

    /**
     * Background/highlight color in hex format (#RRGGBB).
     * Represents the background color or text highlighting.
     * @example "#ffff00" for yellow highlight, "#d3d3d3" for light gray
     */
    backgroundColor?: string;

    /**
     * Font size with units.
     * Most parsers append 'pt' (points), but ODF may use other units like 'in' (inches) or 'cm'.
     * @example "12pt" for 12pt, "14pt" for 14pt, "0.5in" for 0.5 inches
     */
    size?: string;

    /**
     * Font family/typeface name.
     * Extracted from font tables in RTF or font definitions in OOXML.
     * @example "Arial", "Times New Roman", "Calibri", "Ubuntu Mono"
     */
    font?: string;

    /**
     * Whether the text is subscript (e.g., H₂O).
     * Corresponds to `\sub` in RTF, `<w:vertAlign w:val="subscript"/>` in OOXML.
     * Mutually exclusive with superscript.
     * @example true for subscript text like H₂O
     */
    subscript?: boolean;

    /**
     * Whether the text is superscript (e.g., E=mc²).
     * Corresponds to `\super` in RTF, `<w:vertAlign w:val="superscript"/>` in OOXML.
     * Mutually exclusive with subscript.
     * @example true for superscript text like x²
     */
    superscript?: boolean;

    /**
     * The alignment of the text.
     * Common in spreadsheet cells or paragraph styles.
     * @example "center", "right"
     */
    alignment?: 'left' | 'center' | 'right' | 'justify';
}

/**
 * Metadata for a slide in PowerPoint.
 */
export interface SlideMetadata {
    /** The slide number (1-based). */
    slideNumber: number;

    /**
     * The unique ID of the note associated with this slide (if any).
     * @example "slide-note-1"
     */
    noteId?: string;

    /** The style of the slide. */
    style?: string;
}

/**
 * Metadata for a sheet in Excel.
 */
export interface SheetMetadata {
    /** The name of the sheet. */
    sheetName: string;
    /** The style of the sheet. */
    style?: string;
}

/**
 * Metadata for a heading.
 */
export interface HeadingMetadata {
    /** The heading level (e.g., 1 for H1). */
    level: number;
    /** The alignment of the heading. */
    alignment?: 'left' | 'center' | 'right' | 'justify';
    /** The style of the heading. */
    style?: string;
}

/**
 * Metadata for a paragraph.
 */
export interface ParagraphMetadata {
    /** The alignment of the paragraph. */
    alignment?: 'left' | 'center' | 'right' | 'justify';
    /** The style of the paragraph. */
    style?: string;
}

/**
 * Metadata for a list item.
 */
export interface ListMetadata {
    /**
     * The type of list: 'ordered' (numbered) or 'unordered' (bulleted).
     * @example 'ordered' for numbered lists, 'unordered' for bulleted lists
     */
    listType: 'ordered' | 'unordered';

    /**
     * The nesting level (indent level) of the list item, starting from 0.
     * @example 0 for top-level items, 1 for first nested level
     */
    indentation: number;

    /**
     * Text alignment of the list item.
     * @example 'left', 'center', 'right', 'justify'
     */
    alignment: 'left' | 'center' | 'right' | 'justify';

    /**
     * The list ID from the Word document's numbering definition.
     * Used to identify which list definition this item belongs to.
     * @example '1', '2' for different list definitions
     */
    listId: string;

    /**
     * The zero-based index of this item within its list.
     * Continues incrementing even across paragraph interruptions for the same listId.
     * @example 0, 1, 2, 3 for sequential list items
     */
    itemIndex: number;

    /**
     * The style name of the list item.
     * @example "ListParagraph"
     */
    style?: string;
}

/**
 * Metadata for a table cell (primarily used in Excel/spreadsheet parsing).
 * Contains positional information about where the cell appears in the table.
 */
export interface CellMetadata {
    /**
     * The row index of the cell (0-based).
     * @example 0 for the first row, 1 for the second row, etc.
     */
    row: number;
    /**
     * The column index of the cell (0-based).
     * @example 0 for column A, 1 for column B, etc.
     */
    col: number;
    /**
     * The number of rows this cell spans (merges).
     * @example 2 if the cell is merged with the one below it.
     */
    rowSpan?: number;
    /**
     * The number of columns this cell spans (merges).
     * @example 2 if the cell is merged with the one to its right.
     */
    colSpan?: number;
    /** The style of the cell. */
    style?: string;
}

/**
 * Metadata for a chart node in the document.
 * Links the chart node to its corresponding attachment in the attachments array.
 */
export interface ChartMetadata {
    /**
     * The name of the attachment that contains the actual chart data.
     * Use this to look up the full chart data from the attachments array.
     * @example "chart1.xml"
     */
    attachmentName: string;
}

/**
 * Metadata for an image node in the document.
 * Links the image node to its corresponding attachment in the attachments array.
 */
export interface ImageMetadata {
    /**
     * The name of the attachment that contains the actual image data.
     * Use this to look up the full image data from the attachments array.
     * @example "image1.png"
     */
    attachmentName: string;

    /**
     * Alt text (alternative text) describing the image.
     * Extracted from image properties in the document.
     * @example "Company logo"
     */
    altText?: string;
}

/**
 * Metadata for PDF page nodes.
 * Indicates which page of the PDF this content came from.
 */
export interface PageMetadata {
    /**
     * The page number (1-based) from the PDF document.
     * @example 1 for the first page, 2 for the second page, etc.
     */
    pageNumber: number;
}

/**
 * Metadata for text nodes that contain hyperlinks.
 * Used to track hyperlinks in text runs.
 */
export interface TextMetadata {
    /** Style name of the text */
    style?: string;

    /**
     * The hyperlink URL (for external links) or anchor reference (for internal links).
     * @example "https://example.com" or "#_Toc123456"
     */
    link?: string;

    /**
     * Type of hyperlink.
     * - 'internal': Link to a bookmark/anchor within the same document
     * - 'external': Link to an external URL
     */
    linkType?: 'internal' | 'external';
}

/**
 * Metadata for note nodes (footnotes/endnotes).
 * Used in ODT and DOCX files to track notes.
 */
export interface NoteMetadata {
    /**
     * Type of note: 'footnote' or 'endnote'.
     */
    noteType?: 'footnote' | 'endnote';

    /**
     * The unique ID of the note from the source document.
     * @example "1", "2"
     */
    noteId?: string;
}

/**
 * Union type for content metadata.
 */
export type ContentMetadata = SlideMetadata | SheetMetadata | HeadingMetadata | ListMetadata | CellMetadata | ImageMetadata | ChartMetadata | PageMetadata | ParagraphMetadata | TextMetadata | NoteMetadata | undefined;


/**
 * Represents a node in the document content tree.
 * This is the core building block of the parsed document structure.
 * Content nodes can be nested to represent hierarchical document structures
 * (e.g., paragraphs containing text runs, tables containing rows, rows containing cells).
 * 
 * @example
 * // A simple paragraph with formatted text
 * {
 *   type: 'paragraph',
 *   text: 'Hello world',
 *   children: [
 *     { type: 'text', text: 'Hello ', formatting: { bold: true } },
 *     { type: 'text', text: 'world', formatting: { italic: true } }
 *   ]
 * }
 * 
 * @example
 * // A heading with metadata
 * {
 *   type: 'heading',
 *   text: 'Chapter 1',
 *   metadata: { level: 1 },
 *   children: [...]
 * }
 */
export interface OfficeContentNode {
    /**
     * The type of the node.
     * Determines how the node should be interpreted and rendered.
     * Common types: 'paragraph', 'heading', 'table', 'list', 'text', 'image', etc.
     */
    type: OfficeContentNodeType;

    /**
     * The complete text content of the node and all its children combined.
     * For container nodes (paragraph, heading), this is the concatenation of all child text.
     * For leaf nodes (text), this is the actual text content.
     * @example "Hello world" for a paragraph containing "Hello " and "world"
     */
    text?: string;

    /**
     * Child nodes that make up this node's content.
     * Used for hierarchical structures:
     * - Paragraphs contain text runs with different formatting
     * - Tables contain rows
     * - Rows contain cells
     * - Cells contain paragraphs
     * @example [{ type: 'text', text: 'Hello', formatting: { bold: true } }]
     */
    children?: OfficeContentNode[];

    /**
     * Text formatting applied to this node.
     * Only applicable to text-containing nodes.
     * For container nodes like paragraphs, formatting typically appears on child text nodes.
     * @example { bold: true, size: "12", font: "Arial" }
     */
    formatting?: TextFormatting;

    /**
     * Type-specific metadata providing additional context about the node.
     * The metadata structure depends on the node type:
     * - Headings: { level: 1 }
     * - Lists: { listType: 'ordered', indentation: 0 }
     * - Cells: { row: 0, col: 0 }
     * - Slides: { slideNumber: 1 }
     * @example { level: 1 } for a heading
     */
    metadata?: ContentMetadata;

    /**
     * The raw source content for this node.
     * - For XML-based formats (DOCX, XLSX, PPTX): contains the raw XML
     * - For RTF: contains the raw RTF markup
     * - For PDF: typically not available
     * Only populated when `config.includeRawContent` is true.
     * Useful for debugging or when you need access to format-specific features.
     * @example "<w:p><w:r><w:t>Hello</w:t></w:r></w:p>" for DOCX
     */
    rawContent?: string;
}

/**
 * Structured information extracted from a chart.
 */
export interface ChartData {
    /** Chart title (if any) */
    title?: string;

    /** X-axis title (for continuous or categorical axes) */
    xAxisTitle?: string;

    /** Y-axis title (for value or continuous axes) */
    yAxisTitle?: string;

    /** 
     * Collections of data points. 
     * For bar/line charts, each dataset is one 'line' or group of bars.
     * For pie charts, there is typically only one dataset.
     */
    dataSets: {
        /** Name of this data group (e.g., 'Sales 2023') */
        name?: string;
        /** Actual numeric or string values for this group */
        values: string[];
        /** Specific labels for each point in this dataset (if defined per point) */
        pointLabels: string[];
    }[];

    /** 
     * Labels for the chart facets (e.g., 'Jan', 'Feb', 'Mar' on X-axis).
     * These typically correspond to the data points in each dataSet.
     */
    labels: string[];

    /** Every text node discovered in the chart XML (for keyword search/raw extraction) */
    rawTexts: string[];
}

/**
 * Represents an attachment extracted from the document (image, chart, etc.).
 * Attachments are binary resources embedded in the document.
 * Only populated when `config.extractAttachments` is true.
 * 
 * @example
 * ```typescript
 * {
 *   type: 'image',
 *   mimeType: 'image/png',
 *   data: 'iVBORw0KGgoAAAANSUhEUgAA...',  // Base64
 *   name: 'chart1.png',
 *   extension: 'png',
 *   ocrText: 'Sales Chart Q4 2024'  // If OCR was enabled
 * }
 * ```
 */
export interface OfficeAttachment {
    /**
     * The category of the attachment.
     * Helps identify what kind of content this represents.
     * @example 'image' for photos and diagrams, 'chart' for embedded charts
     */
    type: 'image' | 'chart';

    /**
     * The MIME type of the attachment data.
     * Indicates the file format and how the data should be interpreted.
     * @example 'image/png', 'image/jpeg', 'image/svg+xml'
     */
    mimeType: OfficeMimeType;

    /**
     * The attachment content encoded as Base64.
     * This is the actual binary data of the image/chart/etc. encoded for text transmission.
     * Can be used directly in HTML img tags with data URIs or decoded to binary.
     * @example "iVBORw0KGgoAAAANSUhEUgAA..." (truncated)
     */
    data: string;

    /**
     * A unique name for this attachment file.
     * May be derived from the source file or auto-generated.
     * Used to link `ImageMetadata` nodes to their corresponding attachments.
     * @example "image1.png", "chart2.emf", "picture3.jpg"
     */
    name: string;

    /**
     * The file extension (without the dot).
     * Derived from the MIME type or original filename.
     * @example "png", "jpg", "svg"
     */
    extension: string;

    /**
     * Text extracted from the image using Optical Character Recognition (OCR).
     * Only present when:
     * - `config.ocr` is true
     * - `config.extractAttachments` is true
     * - The attachment is an image containing text
     * Uses Tesseract.js with the language specified in `config.ocrLanguage`.
     * @example "Annual Revenue: $1.2M"
     */
    ocrText?: string;

    /**
     * Alt text or description associated with the image in the document.
     * Extracted from the document markup (e.g., wp:docPr descr attribute in DOCX).
     * @example "A chart showing sales growth"
     */
    altText?: string;

    /**
     * Structured data extracted from a chart attachment.
     * Only present if the attachment is a chart and data extraction was successful.
     * Contains series names, values, labels, and titles.
     * @example { title: "Sales Chart", series: [...], categories: [...] }
     */
    chartData?: ChartData;
}

/**
 * Metadata for the parsed file.
 */
export interface OfficeMetadata {
    /** The title of the document. */
    title?: string;
    /** The author of the document. */
    author?: string;
    /** User who last modified the document. */
    lastModifiedBy?: string;
    /** Creation date. */
    created?: Date;
    /** Last modification date. */
    modified?: Date;
    /** Description/Comments. */
    description?: string;
    /** Subject/Topic. */
    subject?: string;
    /** Number of pages (if available). */
    pages?: number;
    /** Document-wide default formatting settings (font, size, color). */
    formatting?: Partial<TextFormatting>;
    /** Style map for styles in the document. */
    styleMap?: Record<string, Partial<TextFormatting>>;
    /**
     * User-defined custom properties embedded in the document.
     * Sources by format:
     * - DOCX/XLSX/PPTX: `docProps/custom.xml` (Office custom document properties)
     * - ODT/ODP/ODS: `meta:user-defined` elements in `meta.xml`
     * - RTF: `\info` sub-destinations such as `\manager`, `\company`, `\keywords`, `\category`
     * - PDF: non-standard entries in the PDF Info dictionary
     * Values are typed as string, number, boolean, or Date where the source format provides type information.
     */
    customProperties?: Record<string, string | number | boolean | Date>;
}

/**
 * The Abstract Syntax Tree (AST) returned by the parser.
 * This is the root data structure representing the entire parsed document.
 * 
 * The AST provides a format-agnostic representation of the document that can be easily
 * processed, transformed, or converted to other formats. It preserves the document's
 * structure, content, formatting, and metadata while abstracting away format-specific details.
 * 
 * @example
 * ```typescript
 * const ast = await OfficeParser.parseOffice('document.docx', {
 *   extractAttachments: true,
 *   includeRawContent: false
 * });
 * 
 * console.log(ast.type); // 'docx'
 * console.log(ast.metadata.author); // 'John Doe'
 * console.log(ast.content.length); // Number of top-level content nodes
 * console.log(ast.toText()); // Plain text representation
 * ```
 */
export interface OfficeParserAST {
    /**
     * The type of the parsed file.
     * Indicates which parser was used and what format the input was in.
     * @example 'docx', 'xlsx', 'pptx', 'rtf', 'pdf', 'odt', 'odp', 'ods'
     */
    type: SupportedFileType;

    /**
     * Document metadata extracted from the file properties.
     * Includes information like author, title, creation date, etc.
     * Availability depends on the file format and whether metadata was present in the source.
     * @example { author: 'John Smith', title: 'Annual Report', created: new Date('2024-01-01') }
     */
    metadata: OfficeMetadata;

    /**
     * The hierarchical content structure of the document.
     * This is an array of top-level content nodes. Each node can have children, creating a tree.
     * For different file types:
     * - DOCX: Array of paragraphs, headings, tables, etc.
     * - XLSX: Array of sheets, each containing rows
     * - PPTX: Array of slides, each containing content nodes
     * - PDF: Array of pages, each containing paragraphs
     * @example [{ type: 'paragraph', text: 'Hello' }, { type: 'heading', text: 'Chapter 1' }]
     */
    content: OfficeContentNode[];

    /**
     * Attachments extracted from the document (images, charts, embedded files).
     * Only populated when `config.extractAttachments` is true.
     * Each attachment includes:
     * - Base64-encoded data
     * - MIME type
     * - Optional OCR text (if `config.ocr` is true)
     * @example [{ type: 'image', mimeType: 'image/png', data: 'base64...', name: 'image1.png' }]
     */
    attachments: OfficeAttachment[];

    /**
     * Converts the entire AST to plain text.
     * This method flattens the document structure and returns just the text content,
     * stripping out all formatting, metadata, and structure.
     * 
     * The text is concatenated using the delimiter specified in `config.newlineDelimiter` (default: '\n').
     * 
     * @returns A plain text representation of the document
     * @example
     * ```typescript
     * const text = ast.toText();
     * console.log(text); // "Hello world\nChapter 1\n..."
     * ```
     */
    toText(): string;
}

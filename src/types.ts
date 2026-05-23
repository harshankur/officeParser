/**
 * Standard error types for OfficeParser.
 * Use these to identify the kind of error being reported.
 */
export enum OfficeErrorType {
    /** Unsupported file extension */
    EXTENSION_UNSUPPORTED = 'EXTENSION_UNSUPPORTED',
    /** File appears to be corrupted or malformed */
    FILE_CORRUPTED = 'FILE_CORRUPTED',
    /** File could not be found at the specified path */
    FILE_DOES_NOT_EXIST = 'FILE_DOES_NOT_EXIST',
    /** Specified location/directory is not reachable or is a directory */
    LOCATION_NOT_FOUND = 'LOCATION_NOT_FOUND',
    /** Arguments passed to the function are missing or invalid */
    IMPROPER_ARGUMENTS = 'IMPROPER_ARGUMENTS',
    /** Error occurred while reading or processing file buffers */
    IMPROPER_BUFFERS = 'IMPROPER_BUFFERS',
    /** Input type is not a supported type (string, Buffer, ArrayBuffer, Uint8Array) */
    INVALID_INPUT = 'INVALID_INPUT',
    /** PDF worker source is missing (required in browser) */
    PDF_WORKER_MISSING = 'PDF_WORKER_MISSING',
    /** Attempted to use Node.js-only features in a browser environment */
    FEATURE_NOT_SUPPORTED_IN_BROWSER = 'FEATURE_NOT_SUPPORTED_IN_BROWSER',
    /** Style mapping string is malformed */
    INVALID_STYLE_MAPPING = 'INVALID_STYLE_MAPPING',
    /** Selector in style mapping is invalid */
    INVALID_SELECTOR = 'INVALID_SELECTOR',
    /** Output mapping in style mapping is invalid */
    INVALID_OUTPUT_MAPPING = 'INVALID_OUTPUT_MAPPING',
    /** Semantic chunking strategy is selected but no embedding function is provided */
    MISSING_EMBEDDING_FUNCTION = 'MISSING_EMBEDDING_FUNCTION'
}

/**
 * Standard warning types for OfficeParser.
 * Use these for reporting non-fatal issues or performance tips.
 */
export enum OfficeWarningType {
    /** Performance advice (e.g., Rosetta translation on Mac) */
    PERFORMANCE_TIP = 'PERFORMANCE_TIP',
    /** OCR processing failed for an attachment */
    OCR_FAILED = 'OCR_FAILED',
    /** Extraction of structured chart data failed */
    CHART_DATA_EXTRACTION_FAILED = 'CHART_DATA_EXTRACTION_FAILED',
    /** Automatic worker path failed, falling back to CDN */
    PDF_WORKER_FALLBACK = 'PDF_WORKER_FALLBACK',
    /** General attachment extraction failure */
    ATTACHMENT_EXTRACTION_FAILED = 'ATTACHMENT_EXTRACTION_FAILED',
    /** Failed to load a specific page in a multi-page document */
    PAGE_LOAD_FAILED = 'PAGE_LOAD_FAILED',
    /** Failed to load a required dynamic dependency */
    DEPENDENCY_LOAD_FAILED = 'DEPENDENCY_LOAD_FAILED',
    /** Failed to extract images from a source */
    IMAGE_EXTRACTION_FAILED = 'IMAGE_EXTRACTION_FAILED',
    /** Failed to extract annotations from a document */
    ANNOTATION_EXTRACTION_FAILED = 'ANNOTATION_EXTRACTION_FAILED',
    /** Failed to process an extracted image bitmap */
    IMAGE_PROCESSING_FAILED = 'IMAGE_PROCESSING_FAILED',
    /** Warning about limitations of browser-based generation */
    BROWSER_GENERATION_LIMITATION = 'BROWSER_GENERATION_LIMITATION',
    /** Specified sheet range in Excel/ODS export was not found */
    SHEET_RANGE_NOT_FOUND = 'SHEET_RANGE_NOT_FOUND',
    /** Buffer content type does not match the provided or expected file extension */
    BUFFER_TYPE_MISMATCH = 'BUFFER_TYPE_MISMATCH',
    /** Failed to detect file type from buffer due to library error or incompatibility */
    FILE_TYPE_DETECTION_FAILED = 'FILE_TYPE_DETECTION_FAILED',
    /** No chunks were generated for the document given the current strategy */
    EMPTY_CHUNK_GENERATED = 'EMPTY_CHUNK_GENERATED',
    /** A node was skipped because it only contained whitespace */
    WHITESPACE_NODE_SKIPPED = 'WHITESPACE_NODE_SKIPPED'
}

/**
 * Configuration options for OCR.
 */
export interface OcrConfig {
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
    language?: string;
    /**
     * Path to the Tesseract worker script.
     * Primarily used for offline/air-gapped environments.
     * Default is ''.
     */
    workerPath?: string;
    /**
     * Path to the Tesseract core script.
     * Primarily used for offline/air-gapped environments.
     * Default is ''.
     */
    corePath?: string;
    /**
     * Path for Tesseract language files (traineddata).
     * Primarily used for offline/air-gapped environments.
     * Default is ''.
     */
    langPath?: string;
    /**
     * Timeout in milliseconds of inactivity before the OCR worker pool is automatically terminated.
     * Set to 0 to disable auto-termination.
     * Default is 10,000 (10 seconds).
     */
    autoTerminateTimeout?: number;
}

/**
 * Configuration options for the OfficeParser.
 */
export interface OfficeParserConfig {
    /**
     * @deprecated Use `onWarning` instead.
     * Flag to show all the logs to console in case of an error irrespective of your own handling.
     * Default is false.
     */
    outputErrorToConsole?: boolean;
    /**
     * Callback for warnings or non-fatal errors encountered during parsing.
     * Allows you to capture issues like OCR failures or attachment extraction errors
     * without stopping the parsing process.
     */
    onWarning?: (issue: OfficeIssue) => void;
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
     * @deprecated Use `ocrConfig.language` instead.
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
     * Shared OCR configuration for worker pooling and offline support.
     * If provided, `ocrLanguage` will be ignored in favor of `ocrConfig.language`.
     */
    ocrConfig?: OcrConfig;

    /**
     * Flag to serialize raw content (XML) as clean, formatted strings.
     * Only relevant when `includeRawContent` is true.
     * Default is true.
     * 
     * If false, the parser will attempt to extract the original raw substring from the 
     * source document instead of re-serializing the DOM node.
     */
    serializeRawContent?: boolean;
    /**
     * Flag to preserve original XML whitespace and line endings when serializing.
     * Only relevant when `includeRawContent` is true and `serializeRawContent` is true.
     * Default is false.
     */
    preserveXmlWhitespace?: boolean;
    /**
     * The URL/path to the PDF.js worker script.
     *
     * **Mandatory** when using PDF parsing in browser environments to avoid worker configuration errors.
     * If not provided, it defaults to `https://cdn.jsdelivr.net/npm/pdfjs-dist@5.6.205/build/pdf.worker.min.mjs`.
     * You can override this with your own local path or a different CDN link.
     */
    pdfWorkerSrc?: string;
    /**
     * Flag to include break nodes in the AST.
     * This is currently only supported for Word documents. (w:br nodes)
     *
     * Default is false
     */
    includeBreakNodes?: boolean;
    /**
     * Flag to ignore all internal (anchor) links during parsing.
     * When true, all bookmarks, cross-references, and internal document jumps are stripped 
     * from the AST. Only external URLs will be preserved.
     * 
     * Use this if you want a "flat" document without any internal interactivity.
     * 
     * Default is false.
     */
    ignoreInternalLinks?: boolean;
    /**
     * Optional hint for the file format.
     * When a Buffer or ArrayBuffer is passed, the parser relies on magic bytes to detect the file type.
     * Text-based formats like 'md', 'html', and 'csv' lack reliable magic bytes.
     * If you are parsing these formats from a Buffer, you must provide this fileType hint.
     * 
     * This is authoritative and is used to determine the file type, so it should be accurate.
     * If provided, this bypasses the magic bytes detection and the file extension-based detection either way.
     * 
     * Default is null.
     */
    fileType?: SupportedFileType | null;
    /**
     * Custom delimiter for CSV files.
     * Defaults to ',' but can be overridden (e.g., ';', '\t').
     */
    csvDelimiter?: string;
}

/**
 * A fully-populated parser configuration containing all options.
 * Used internally for merging and resolution.
 */
export type FullOfficeParserConfig = DeepRequired<OfficeParserConfig>;


/**
 * Represents a single issue (warning, error, or info) generated during document processing.
 */
export interface OfficeIssue {
    /** The severity of the issue. */
    type: 'warning' | 'info' | 'error';
    /** Human-readable message text. */
    message: string;
    /** The specific AST node that triggered this issue, if applicable. */
    node?: OfficeContentNode;
    /** A unique error code for programmatic handling. */
    code: OfficeWarningType | OfficeErrorType;
    /** Optional additional context or original error object. */
    details?: any;
}

/**
 * The result of a document conversion operation.
 */
export interface ConversionResult<D extends string = UniversalGeneratorFormat> {
    /** The actual generated content (HTML, Markdown, Text, OfficeChunk[], etc.). */
    value: D extends 'pdf' ? Uint8Array :
    D extends 'chunks' ? OfficeChunk[] :
    D extends 'csv' ? string | Uint8Array :
    D extends UniversalGeneratorFormat ? string : never;
    /** A collection of issues (warnings/infos) generated during the process. */
    messages: OfficeIssue[];
}

/**
 * Universal formats supported by all source types for generation.
 */
export type UniversalGeneratorFormat = 'text' | 'md' | 'html' | 'pdf' | 'csv' | 'rtf' | 'chunks';

/**
 * Allowed destination formats for a given source type.
 * Currently, all generators are universal across all source formats.
 */
export type SupportedDestination<_T extends SupportedFileType = SupportedFileType> = UniversalGeneratorFormat;

/**
 * Configuration options for the OfficeGenerator.
 */
/**
 * Common configuration options for all generators.
 */
export interface CommonGeneratorConfig {
    /**
     * Callback called for every node during generation.
     * Allows users to modify nodes before processing, completely override rendering, or filter them out.
     * 
     * #### Callback Capabilities:
     * 1. **Filter/Remove Nodes**: Return `false` to skip a node and all its children.
     * 2. **Override Rendering**: Return a `string` to use that exact text as the output, bypassing default logic and recursion.
     * 3. **Mutate Nodes**: Modify the `node` object directly (e.g., changing `node.text`) and return `void` to let the generator proceed with your changes.
     * 4. **Async Support**: The callback can be `async`, allowing you to fetch external data or perform complex logic during generation.
     */
    onNode?: (node: OfficeContentNode) => string | false | Promise<string | false | void> | void;
    /**
     * Callback for warnings, non-fatal errors, or issues encountered during generation.
     * Allows the process to continue while reporting skipping or approximation of content.
     */
    onWarning?: (issue: OfficeIssue) => void;
    /**
     * Map document styles (e.g., 'Heading 1', 'Intense Quote') to specific semantic elements.
     * 
     * DESIGN PHILOSOPHY:
     * This is the primary way to customize how the library interprets the visual 
     * structure of your source documents. 
     * 
     * To disable all semantic translation and use raw AST types only, 
     * set `ignoreDefaultStyleMap: true` and leave `styleMap` empty.
     * 
     * It supports two formats:
     * 
     * 1. LEGACY STRING DSL:
     * Simple "selector => output" syntax. Highly compatible with mammoth.js style maps.
     * @example ["p[style-name='Heading 1'] => h1"]
     * @example ["p[style='Quote'] => blockquote"]
     * 
     * 2. STRUCTURED OBJECTS (Recommended):
     * More powerful and strictly typed. Ideal for complex logic or when you 
     * need to apply specific classes/attributes for the HTML generator.
     * @example 
     * [
     *   { 
     *     selector: { nodeType: 'paragraph', attributes: { style: 'Heading 1' } }, 
     *     output: { tag: 'h1', classes: ['main-title'], attributes: { id: 'top' } } 
     *   }
     * ]
     * 
     * Note: This property works in conjunction with `ignoreDefaultStyleMap`.
     * Defaults to a robust built-in map that covers common standard Office styles.
     */
    styleMap?: string[] | StructuredStyleMapping[];
    /**
     * Whether to include visual formatting like font size, font family, and colors in the output.
     * Set to false for clean, semantic output.
     * Defaults to true.
     */
    includeFormatting?: boolean;
    /**
     * Whether to automatically generate unique slug-based IDs for headings.
     * Useful for table-of-contents and anchor links.
     * Defaults to true.
     */
    generateIds?: boolean;
    /**
     * Whether to render document metadata (title, author, etc.) as visible content 
     * in the generated output (e.g., a header block in HTML or plain text).
     * Structural metadata (HTML <meta> tags, Markdown YAML frontmatter) is always included.
     * Defaults to false.
     */
    renderMetadata?: boolean;
    /**
     * Whether to ignore the built-in default style mappings (e.g. "Heading 1" -> h1).
     * Set to true if you want full control over style mapping.
     * Defaults to false.
     */
    ignoreDefaultStyleMap?: boolean;
    /**
     * Whether to include images in the generated output.
     * Defaults to true.
     */
    includeImages?: boolean;
    /**
     * Whether to include interactive charts in the generated output (HTML only).
     * Defaults to true.
     */
    includeCharts?: boolean;
    /**
     * Whether to ignore all internal (anchor) links and anchor IDs during generation.
     * When true, all bookmarks, cross-references, and internal document jumps are stripped.
     * Specifically for Markdown, this removes the {#id} block from headings.
     * Defaults to false.
     */
    ignoreInternalLinks?: boolean;
}

/**
 * Destination-aware generator configuration.
 * Restricts format-specific configurations to their respective destinations.
 */
/**
 * Mapping of destination formats to their specific configuration interfaces.
 */
export interface GeneratorSubConfigMap {
    html: HtmlGeneratorConfig;
    md: MdGeneratorConfig;
    pdf: PdfGeneratorConfig;
    csv: CsvGeneratorConfig;
    text: TextGeneratorConfig;
    rtf: RtfGeneratorConfig;
    chunks: ChunkingConfig;
}

/**
 * Configuration options for document generators.
 * 
 * This interface is designed to be format-aware. When you specify a destination format 
 * (e.g., `OfficeGenerator.generate(ast, 'html', config)`), the generic parameter `D` 
 * ensures that only the relevant sub-configuration (e.g., `htmlConfig`) is available 
 * for type checking.
 * 
 * @template D The destination format string. Defaults to `string` for a general configuration.
 */
export type GeneratorConfig<D extends string = string> = CommonGeneratorConfig & {
    [K in keyof GeneratorSubConfigMap as `${K & string}Config`]?: string extends D
    ? GeneratorSubConfigMap[K]
    : (D extends K ? GeneratorSubConfigMap[K] : never);
}

/**
 * Configuration options for the OfficeConverter.
 * Combines relevant parser and generator settings for a seamless one-step conversion.
 * 
 * @template D The destination format string.
 */
/**
 * Configuration options for the OfficeConverter.
 * Combines general generator settings with a specific subset of parser settings.
 * 
 * @template D The destination format string.
 * @template T The source file type.
 */
export type OfficeConverterConfig<D extends string = string, T extends SupportedFileType = SupportedFileType> = {
    /** 
     * Specific configuration for the source parsing phase.
     */
    parseConfig?: OfficeParserConfig & { fileType?: T };
    /**
     * Specific configuration for the destination generation phase.
     */
    generatorConfig?: GeneratorConfig<D>;
    /**
     * Callback for warnings or non-fatal errors encountered during the entire conversion process.
     * This is passed to both the parser and the generator.
     * If provided, this takes precedence over callbacks inside parseConfig or generatorConfig.
     */
    onWarning?: (issue: OfficeIssue) => void;
}

/**
 * Deeply required type helper.
 */
export type DeepRequired<T> = T extends Function | Date | Buffer | RegExp
    ? T
    : T extends Array<infer U>
    ? Array<DeepRequired<U>>
    : T extends object
    ? { [P in keyof T]-?: DeepRequired<T[P]> }
    : T;



/**
 * A fully-populated generator configuration containing all sub-configs.
 * Used internally for merging and resolution.
 * `chunksConfig` is typed as `ChunkingConfig` directly (not DeepRequired) because
 * it is a discriminated union whose members cannot be uniformly deep-required.
 */
export type FullGeneratorConfig = DeepRequired<CommonGeneratorConfig & {
    [K in keyof Omit<GeneratorSubConfigMap, 'chunks'> as `${K}Config`]: GeneratorSubConfigMap[K];
}> & {
    chunksConfig: ChunkingConfig;
};



/**
 * Configuration options for HTML generation.
 */
export interface HtmlGeneratorConfig {
    /**
     * Whether to wrap the output in a full HTML document structure (e.g., <html>, <head>, etc.).
     * Defaults to true.
     */
    standalone?: boolean;
    /**
     * URL for the Chart.js library to use when 'includeCharts' is true.
     * Defaults to 'https://cdn.jsdelivr.net/npm/chart.js'.
     */
    chartJsSrc?: string;
}

/**
 * Configuration options for PDF generation.
 * Maps closely to Puppeteer's PDF options.
 */
export interface PdfGeneratorConfig {
    /** Paper format. Defaults to 'A4'. */
    format?: 'letter' | 'legal' | 'tabloid' | 'ledger' | 'a0' | 'a1' | 'a2' | 'a3' | 'a4' | 'a5' | 'a6' | 'Letter' | 'Legal' | 'Tabloid' | 'Ledger' | 'A0' | 'A1' | 'A2' | 'A3' | 'A4' | 'A5' | 'A6';
    /** Paper width, accepts values labeled with units (e.g., '5in', '3cm') or numbers (in pixels). */
    width?: string | number;
    /** Paper height, accepts values labeled with units (e.g., '5in', '3cm') or numbers (in pixels). */
    height?: string | number;
    /** Whether to print in landscape orientation. Defaults to false. */
    landscape?: boolean;
    /** Whether to print background graphics. Defaults to true. */
    printBackground?: boolean;
    /** Scale of the webpage rendering. Defaults to 1. */
    scale?: number;
    /** Paper margins. */
    margin?: {
        top?: string | number;
        right?: string | number;
        bottom?: string | number;
        left?: string | number;
    };
    /** Whether to display header and footer. Defaults to false. */
    displayHeaderFooter?: boolean;
    /** HTML template for the print header. */
    headerTemplate?: string;
    /** HTML template for the print footer. */
    footerTemplate?: string;
    /** 
     * Optional Puppeteer launch options for Node.js environment. 
     * Useful for setting custom executable paths or args in CI/CD.
     */
    launchOptions?: any;
}

/**
 * Structured style mapping definition for the StyleMapper.
 * 
 * DESIGN PHILOSOPHY: "Semantic Translation"
 * -----------------------------------------
 * Office documents (Word, RTF, PPTX) often use custom or localized style names 
 * (e.g., "Heading 1" in English vs "Titre 1" in French, or "MyCompany-Quote").
 * 
 * This interface allows you to create a "semantic bridge" between these arbitrary 
 * source styles and a universal vocabulary of document elements.
 * 
 * WHY USE HTML TAGS FOR NON-HTML OUTPUT?
 * --------------------------------------
 * We use HTML tags (`h1`, `blockquote`, `code`, `pre`) as a "Universal Intermediate 
 * Language". By mapping a custom Word style to `blockquote`, you are defining its 
 * SEMANTIC MEANING rather than its physical appearance.
 * 
 * Each generator then interprets this meaning natively:
 * - HTML Generator: Directly renders the `<blockquote>` tag with your classes.
 * - Markdown Generator: Sees 'blockquote' and renders the standard `> ` prefix.
 * - Text Generator: Sees 'blockquote' and applies appropriate structural indentation.
 */
export interface StructuredStyleMapping {
    /** 
     * The criteria used to identify which AST nodes should be transformed. 
     * Think of this as the "Source Filter".
     */
    selector: {
        /** 
         * The structural type of the node (e.g., 'paragraph', 'heading', 'text'). 
         * Most style mappings target 'paragraph' nodes to convert them into headers or blocks.
         */
        nodeType?: string;
        /** 
         * A dictionary of attributes to match on the node.
         * 
         * The most common use case is matching the 'style' attribute from 
         * Word documents (e.g., { style: 'Intense Quote' }).
         * 
         * Matchers:
         * - Literal: `style: 'Heading 1'` matches exactly.
         * - Operator: `{ value: 'Title', operator: '~=' }` matches if the word 'Title' 
         *   is found within the style name.
         */
        attributes?: Record<string, string | number | boolean | { value: string | number | boolean, operator: '=' | '~=' }>;
    };
    /** 
     * The target representation for the matched node.
     * Think of this as the "Semantic Meaning" you want to assign to the match.
     */
    output: {
        /** 
         * The universal semantic tag (e.g., 'h1', 'h2', 'blockquote', 'code', 'pre', 'u').
         * All generators use this tag to decide their native output syntax.
         */
        tag: string;
        /** 
         * CSS classes to apply to the output. 
         * This is utilized by the HTML generator to allow for downstream CSS styling.
         */
        classes?: string[];
        /** 
         * Key-value pair of HTML attributes (like 'id', 'data-*', or 'style') to apply. 
         * Primarily used by the HTML generator for high-fidelity conversion.
         */
        attributes?: Record<string, string>;
        /** 
         * If true, prevents the generator from collapsing this element into 
         * adjacent elements of the same type. 
         * 
         * For example, multiple paragraphs mapped to 'blockquote' normally merge into 
         * one big blockquote. Setting `fresh: true` forces them to be separate blocks.
         */
        fresh?: boolean;
    };
}

/**
 * Configuration options for RTF generation.
 */
export interface RtfGeneratorConfig {
    // Reserved for future RTF-specific options like page size or font embedding
}

/**
 * Configuration options for CSV generation.
 */
export interface CsvGeneratorConfig {
    /**
     * Range of sheets to export.
     * Supports formats like "1", "1-3", "1,2", "1,3-5,7".
     * 1-based indexing.
     * Default is '' (all sheets).
     */
    sheets?: string;
    /**
     * Whether to merge all selected sheets into a single CSV.
     * If false, returns a ZIP archive containing individual CSV files.
     * Defaults to false.
     */
    mergeSheets?: boolean;
    /**
     * Custom delimiter for CSV files.
     * Defaults to ','.
     */
    columnDelimiter?: string;
}

/**
 * Configuration options for Markdown generation.
 */
export interface MdGeneratorConfig {
    /**
     * Whether to fallback to HTML tags for features not supported by standard Markdown.
     * 
     * Markdown has limited support for complex document structures. This flag controls how 
     * the generator handles features that cannot be represented in pure Markdown:
     * 
     * 1. If a feature is NOT supported natively by Markdown (e.g., nested tables, text alignment,
     *    underline, subscript/superscript):
     *    - If true: The generator will use HTML tags (<u>, <sub>, <div>, <table>, etc.) to 
     *      maintain high fidelity.
     *    - If false: The generator will skip or simplify the feature (e.g., ignoring alignment,
     *      skipping underline, or hoisting nested tables out of their cells).
     * 
     * 2. If a feature IS supported by Markdown but a higher quality version is possible 
     *    via HTML (e.g., tables with merged cells):
     *    - If true: Use HTML for better fidelity.
     *    - If false: Use native Markdown syntax (e.g., a standard GFM table grid).
     * 
     * Defaults to true.
     */
    fallbackToHtml?: boolean;
}

/**
 * Configuration options for plain text generation.
 */
export interface TextGeneratorConfig {
    /**
     * The delimiter used for every new line.
     * Defaults to '\n'.
     */
    newlineDelimiter?: string;
    /**
     * Whether to attempt to preserve the original document layout.
     * If true, tables will be rendered with separators and aligned columns.
     * If false, output will be a flat stream of text nodes.
     * Defaults to false.
     */
    preserveLayout?: boolean;
}


// ─── Chunking Types ───────────────────────────────────────────────────────────

/**
 * The strategy used for chunking a document for RAG pipelines.
 * - 'fixed-size': Traditional character/token count based splitting.
 * - 'document-structure': Leverages the AST to split at natural document boundaries.
 * - 'semantic': Uses embedding similarity to find natural topic breakpoints.
 */
export type ChunkingStrategy = 'fixed-size' | 'document-structure' | 'semantic';

/**
 * Base configuration applicable to all chunking strategies.
 */
export interface BaseChunkingConfig {
    /**
     * The strategy used for chunking.
     * Default is 'document-structure'.
     */
    strategy?: ChunkingStrategy;

    /**
     * A function that measures the size of a text string.
     * Defaults to character count: `(text) => text.length`.
     * Override with a token counter (e.g., `tiktoken`) for strict LLM context window adherence.
     */
    lengthFunction?: (text: string) => number;

    /**
     * Whether to strip leading/trailing whitespace from each chunk.
     * Default is true.
     */
    stripWhitespace?: boolean;

    /**
     * Whether to include rich AST metadata (page number, slide number, heading, etc.)
     * in the generated chunk objects.
     * Default is true.
     */
    includeMetadata?: boolean;

    /**
     * Whether to include the starting character index of each chunk
     * relative to the whole document. Useful for UI text highlighting.
     * Default is false.
     */
    addStartIndex?: boolean;
    /**
     * Optional custom regex (as string or RegExp object) to identify sentence boundaries.
     * Use this for languages or specific document types that require custom splitting logic.
     * If provided, it overrides or augments the default segmenter.
     * @example /[。？！]/
     */
    sentenceBoundaryRegex?: string | RegExp;
    /**
     * Optional list of abbreviations to ignore when splitting text into sentences.
     * These words, if followed by a period, will not be treated as sentence boundaries.
     * Use this to handle language-specific or domain-specific abbreviations.
     * @example ["Inc", "Ltd", "approx"]
     */
    abbreviations?: string[];
}

/**
 * Configuration for Fixed-Size Chunking.
 * Cuts text based on a maximum size limit with an optional overlap.
 * This is equivalent to LangChain's `RecursiveCharacterTextSplitter`.
 */
export interface FixedSizeChunkingConfig extends BaseChunkingConfig {
    strategy: 'fixed-size';

    /**
     * Maximum size of the chunk, measured by `lengthFunction`.
     * Default is 1000 characters.
     */
    chunkSize?: number;

    /**
     * Number of characters/tokens to overlap between consecutive chunks
     * to avoid losing context at boundaries.
     * Rule of thumb: ~10–20% of `chunkSize`.
     * Default is 200.
     */
    chunkOverlap?: number;

    /**
     * Ordered list of separators to try when splitting.
     * The chunker tries each in order; if a split would exceed `chunkSize`,
     * it tries the next separator.
     * Default is ['\n\n', '\n', ' ', ''].
     */
    separators?: string[];
}

/**
 * Configuration for Document-Structure Chunking.
 * Uses the officeParser AST to split at natural document boundaries like
 * headings, paragraphs, slides, or pages. This is the recommended strategy
 * as it preserves semantic context from the document's own structure.
 */
export interface DocumentStructureChunkingConfig extends BaseChunkingConfig {
    strategy: 'document-structure';

    /**
     * The primary structural element at which to force a chunk boundary.
     * - 'paragraph': Never cross a paragraph boundary (finest-grained, most precise).
     * - 'heading': Split at every heading change.
     * - 'page': Chunks never span multiple pages (PDF only).
     * - 'slide': Chunks never span multiple slides (PPTX/ODP only).
     * - 'sheet': Chunks never span multiple sheets (XLSX/ODS only).
     * Default is 'paragraph'.
     */
    splitBy?: 'page' | 'slide' | 'sheet' | 'heading' | 'paragraph';

    /**
     * Maximum size of a chunk (measured by `lengthFunction`).
     * If a single structural unit (e.g., one paragraph) exceeds this limit,
     * it will be further split using a recursive character splitter.
     * Default is 1000 characters.
     */
    maxChunkSize?: number;

    /**
     * How to handle table nodes when splitting.
     * - 'row': Split by rows, REPEATING the header row in every chunk so the LLM
     *   always understands what the columns mean. (Highly recommended for RAG)
     * - 'flatten': Convert the table to plain text and split like a regular block.
     * Default is 'row'.
     */
    tableSplitStrategy?: 'row' | 'flatten';
}

/**
 * Configuration for Semantic Chunking.
 * Uses an embedding model to detect topic shifts and create boundaries
 * where content meaning naturally changes. Computationally expensive but
 * produces the highest quality chunks.
 */
export interface SemanticChunkingConfig extends BaseChunkingConfig {
    strategy: 'semantic';

    /**
     * A user-provided async function to generate vector embeddings for a text string.
     * Required. Example: a wrapper around OpenAI's `text-embedding-3-small`.
     * @example async (text) => await openai.embeddings.create({ input: text, model: 'text-embedding-3-small' }).then(r => r.data[0].embedding)
     */
    embeddingFunction: (text: string) => Promise<number[]>;

    /**
     * The cosine similarity threshold below which a chunk boundary is created.
     * When the similarity between two adjacent sentences drops below this value,
     * a new chunk starts. Higher = more splits, smaller chunks.
     * Default is 0.8.
     */
    similarityThreshold?: number;

    /**
     * Maximum size of a chunk even if semantic similarity remains high.
     * Prevents runaway chunks when an entire document is on one topic.
     * Default is 2000 characters.
     */
    maxChunkSize?: number;

    /**
     * Number of surrounding sentences to include when computing similarity
     * for a sentence. A larger window reduces noise from single odd sentences.
     * Default is 1.
     */
    bufferSize?: number;
    /**
     * Number of sentences to process in a single batch when calling the embedding function.
     * Higher values are faster but may trigger API rate limits.
     * Default is 50.
     */
    embeddingBatchSize?: number;
}

/**
 * Discriminated union of all chunking strategy configurations.
 */
export type ChunkingConfig = FixedSizeChunkingConfig | DocumentStructureChunkingConfig | SemanticChunkingConfig;

/**
 * Represents a single document chunk ready for a RAG (Retrieval-Augmented Generation) pipeline.
 * 
 * Chunks are the result of splitting a document into smaller, semantically coherent 
 * pieces that fit within the context window of an LLM. Each chunk includes the 
 * extracted text and rich AST-derived metadata for citations and filtered retrieval.
 */
export interface OfficeChunk {
    /** The text content of this chunk. This is what gets embedded. */
    text: string;

    /**
     * Rich contextual metadata extracted from the AST.
     * Use this to populate vector DB metadata fields for filtered retrieval
     * and for LLM citations.
     */
    metadata: {
        /** The source file format (e.g., 'docx', 'pptx', 'pdf'). */
        sourceType: SupportedFileType;
        /** Page number (1-based), if available (PDF). */
        pageNumber?: number;
        /** Slide number (1-based), if available (PPTX/ODP). */
        slideNumber?: number;
        /** Sheet name, if available (XLSX/ODS). */
        sheetName?: string;
        /** The text of the nearest heading above this chunk in the document. */
        closestHeading?: string;
        /** True if this chunk is part of a table split. */
        isTableChunk?: boolean;
        /** Extensible for user-defined metadata. */
        [key: string]: any;
    };

    /** The start character index of this chunk in the full document text. Only set when `addStartIndex` is true. */
    startIndex?: number;
    /** The end character index of this chunk in the full document text. Only set when `addStartIndex` is true. */
    endIndex?: number;
}

// ─── End Chunking Types ────────────────────────────────────────────────────────

/**
 * Supported file types for parsing.
 */
export type SupportedFileType = 'docx' | 'pptx' | 'xlsx' | 'odt' | 'odp' | 'ods' | 'pdf' | 'rtf' | 'md' | 'html' | 'csv';

/**
 * Types of content nodes in the AST.
 */
export type OfficeContentNodeType = 'paragraph' | 'heading' | 'table' | 'list' | 'text' | 'image' | 'chart' | 'drawing' | 'slide' | 'note' | 'sheet' | 'row' | 'cell' | 'page' | 'break' | 'code' | 'comment';

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
    | 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    | 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    | 'application/vnd.oasis.opendocument.chart'
    | 'application/vnd.oasis.opendocument.spreadsheet'
    | 'application/vnd.oasis.opendocument.text'
    | 'application/vnd.oasis.opendocument.presentation'
    | 'application/rtf'
    | 'text/csv'
    | 'text/markdown'
    | 'text/html';

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
    /** Unique anchor IDs for internal linking. */
    anchorIds?: string[];
}

/**
 * Metadata for a sheet in Excel.
 */
export interface SheetMetadata {
    /** The name of the sheet. */
    sheetName: string;
    /** The style of the sheet. */
    style?: string;
    /** Unique anchor IDs for internal linking. */
    anchorIds?: string[];
}

/**
 * Detailed indentation information for paragraphs and headings.
 * Values are typically in twentieths of a point (twips) in OOXML.
 */
export interface IndentationMetadata {
    /** Left indentation. */
    left?: number;
    /** Right indentation. */
    right?: number;
    /** First line indentation. */
    firstLine?: number;
    /** Hanging indentation. */
    hanging?: number;
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
    /** Detailed indentation information. */
    paragraphIndentation?: IndentationMetadata;
    /** Unique anchor IDs for internal linking. */
    anchorIds?: string[];
}

/**
 * Metadata for a paragraph.
 */
export interface ParagraphMetadata {
    /** The alignment of the paragraph. */
    alignment?: 'left' | 'center' | 'right' | 'justify';
    /** The style of the paragraph. */
    style?: string;
    /** Detailed indentation information. */
    paragraphIndentation?: IndentationMetadata;
    /** Unique anchor IDs for internal linking. */
    anchorIds?: string[];
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

    /** Detailed indentation information. */
    paragraphIndentation?: IndentationMetadata;

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
    /** Unique anchor IDs for internal linking. */
    anchorIds?: string[];
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
    /** Unique anchor IDs for internal linking. */
    anchorIds?: string[];
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
    /** Unique anchor IDs for internal linking. */
    anchorIds?: string[];
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

    /**
     * URL of the image if it is an external link.
     * Typical for HTML or Markdown images that point to remote servers.
     * @example "https://example.com/image.png"
     */
    url?: string;
    /** Unique anchor IDs for internal linking. */
    anchorIds?: string[];
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
    /** Unique anchor IDs for internal linking. */
    anchorIds?: string[];
}

/**
 * Metadata for break nodes.
 * Used in DOCX files to track line and page breaks.
 */
export interface BreakMetadata {
    /**
     * Type of break. The break type determines the next location where
     * text shall be placed.
     * - 'column': The next text will be placed in the next column.
     * - 'page': The next text will be placed on the next page.
     * - 'lastRenderedPage': The editing application has inserted a soft break on the last save.
     * - 'textWrapping' (default, assumed when not specified): The next text will be placed on the next line.
     * - 'carriageReturn': An explicit carriage return (w:cr) equivalent to a hard line break.
     */
    breakType: 'column' | 'page' | 'lastRenderedPage' | 'textWrapping' | 'carriageReturn';

    /**
     * Specifies the location which shall be used as the next available line when breakType
     * has a value of 'textWrapping'. Should be ignored for other break types.
     * - 'all': text wrapping break shall advance the text to the next line which spans the full width of the line
     * - 'left': text wrapping break shall restart in next text region unblocked on the left
     * - 'none': text wrapping break shall advance the text to the next line regardless of any floating objects
     * - 'right': text wrapping break shall restart in next text region unblocked on the right
     */
    clear?: 'all' | 'left' | 'none' | 'right';
}

/**
 * Metadata for a code block.
 */
export interface CodeMetadata {
    /** The programming language of the code block (e.g., 'typescript', 'python') */
    language?: string;
    /** Unique anchor IDs for internal linking. */
    anchorIds?: string[];
}

/**
 * Union type for content metadata.
 */
export type ContentMetadata = SlideMetadata | SheetMetadata | HeadingMetadata | ListMetadata | CellMetadata | ImageMetadata | ChartMetadata | PageMetadata | ParagraphMetadata | TextMetadata | NoteMetadata | BreakMetadata | CodeMetadata | undefined;


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
     * - PDF: non-standard entries in the PDF Info dictionary
     * RTF does not support custom properties; the `\info` group is not extracted.
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
 * console.log((await ast.to('md')).value); // Markdown representation
 * console.log((await ast.to('html')).value); // HTML representation
 * console.log((await ast.to('rtf')).value); // RTF representation
 * console.log((await ast.to('csv')).value); // CSV representation
 * console.log((await ast.to('chunks')).value); // Chunks representation
 * ```
 */
export interface OfficeParserAST {
    /**
     * The original configuration used to parse this document.
     * This includes options like OCR settings, delimiter choices, and filtering flags.
     */
    config: OfficeParserConfig;

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

    /** Any warnings or non-fatal issues encountered during parsing. */
    warnings: OfficeIssue[];

    /**
     * @deprecated Use `.to('text')` instead.
     * Note: This method is synchronous, while the new `.to()` method is asynchronous.
     * 
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

    /**
     * Converts this AST to the specified destination format.
     * This is the recommended way to convert the AST to different formats.
     * 
     * @param destination The target format (e.g., 'text', 'md', 'html', 'pdf').
     * @param config Optional configuration for the generator.
     * @returns A promise resolving to the generated content (string or Buffer).
     * @example
     * ```typescript
     * const html = await ast.to('html', { includeFormatting: false });
     * const md = await ast.to('md');
     * ```
     */
    to<T extends this, D extends SupportedDestination<T['type']>>(
        this: T,
        destination: D,
        config?: GeneratorConfig<D>
    ): Promise<ConversionResult<D>>;
}


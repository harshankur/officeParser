/**
 * Standard error types for OfficeParser.
 * Use these to identify the kind of error being reported.
 */
export enum OfficeErrorType {
    /** Unsupported file extension */
    EXTENSION_UNSUPPORTED = 'EXTENSION_UNSUPPORTED',
    /** Unsupported output generator format */
    FORMAT_UNSUPPORTED = 'FORMAT_UNSUPPORTED',
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
    MISSING_EMBEDDING_FUNCTION = 'MISSING_EMBEDDING_FUNCTION',
    /** The operation was aborted */
    OPERATION_ABORTED = 'OPERATION_ABORTED',
    /** ZIP entry count exceeds limit */
    ZIP_ENTRY_COUNT_LIMIT_EXCEEDED = 'ZIP_ENTRY_COUNT_LIMIT_EXCEEDED',
    /** ZIP entry missing a valid declared size */
    ZIP_ENTRY_INVALID_SIZE = 'ZIP_ENTRY_INVALID_SIZE',
    /** ZIP uncompressed size limit exceeded */
    ZIP_SIZE_LIMIT_EXCEEDED = 'ZIP_SIZE_LIMIT_EXCEEDED',
    /** Embedding call timed out */
    EMBEDDING_TIMEOUT = 'EMBEDDING_TIMEOUT'
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
    WHITESPACE_NODE_SKIPPED = 'WHITESPACE_NODE_SKIPPED',
    /** The HTML generator containerWidth option is invalid */
    INVALID_CONTAINER_WIDTH = 'INVALID_CONTAINER_WIDTH'
}

/**
 * Consolidated timeout settings for OCR operations.
 * Preferred over the individual flat timeout properties on {@link OcrConfig},
 * which are now deprecated.
 * 
 * If a key is present here, it takes priority over the corresponding deprecated
 * flat property (e.g. `timeout.autoTerminate` wins over `autoTerminateTimeout`).
 * Set any value to `0` to disable that specific timeout.
 */
export interface OcrTimeoutConfig {
    /**
     * Timeout in milliseconds of inactivity before the OCR worker pool is
     * automatically terminated and freed.
     * 
     * The timer resets every time a new OCR job is enqueued.  When the last
     * job completes and this duration passes without a new one, the entire
     * worker pool is torn down so that no background threads keep the Node.js
     * process alive unnecessarily.
     * 
     * Set to `0` to keep workers alive indefinitely (useful when you want to
     * call {@link terminateOcr} manually at shutdown time).
     * Default is 10,000 ms (10 seconds).
     */
    autoTerminate?: number;
    /**
     * Timeout in milliseconds for initializing a Tesseract worker
     * (loading the JS runtime, downloading or loading the `.traineddata`
     * language file) or for re-initializing an existing worker with a
     * different language.
     * 
     * Multi-language combinations (e.g. `'por+eng+spa'`) must download a
     * separate `.traineddata` file for each language and are therefore
     * particularly susceptible to slow networks.  Tune this value upward if
     * your OCR environment has high network latency or if you are loading
     * languages from disk in a large container image.
     * 
     * When the timeout fires, the failed job is rejected with a non-fatal
     * {@link OfficeWarningType.OCR_FAILED} warning and parsing continues
     * without OCR output for that image.  The stalled worker is terminated
     * and removed from the pool to prevent thread leaks.
     * 
     * Set to `0` to wait indefinitely (not recommended for production; a hung
     * network request will block the entire OCR queue for that language).
     * Default is 60,000 ms (60 seconds).
     */
    workerLoad?: number;
    /**
     * Timeout in milliseconds for the actual OCR text-recognition call
     * (`worker.recognize(image)`) on an already-initialized Tesseract worker.
     * 
     * Recognition time scales with image resolution and the number of active
     * languages.  Very high-resolution scans or unusual character sets can
     * exceed the default.  If this timeout fires, the job is rejected with a
     * non-fatal {@link OfficeWarningType.OCR_FAILED} warning; the worker is
     * terminated and evicted from the pool because its internal state after a
     * mid-recognition timeout is undefined.
     * 
     * Set to `0` to wait indefinitely.
     * Default is 30,000 ms (30 seconds).
     */
    recognition?: number;
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
     * Consolidated timeout settings for all OCR operations.
     * 
     * Prefer this over the deprecated flat timeout properties.
     * If `timeout.autoTerminate` is set, it takes priority over the deprecated `autoTerminateTimeout`.
     */
    timeout?: OcrTimeoutConfig;
    /**
     * @deprecated Use `timeout.autoTerminate` instead.
     * 
     * Timeout in milliseconds of inactivity before the OCR worker pool is automatically terminated.
     * Set to 0 to disable auto-termination.
     * Default is 10,000 (10 seconds).
     * 
     * If `timeout.autoTerminate` is also set, that value takes priority over this one.
     */
    autoTerminateTimeout?: number;
    /**
     * An optional AbortSignal propagated from the main parser configuration to abort active OCR jobs.
     * If the signal is aborted:
     * 1. Any pending OCR jobs in the scheduler queue are rejected immediately.
     * 2. Any active OCR job running on a Tesseract worker will reject, the worker will be
     *    terminated, and it will be removed from the pool to avoid hanging worker threads.
     * 
     * Developers should prefer passing this at the top level of `parseOffice` (as `config.abortSignal`),
     * which automatically propagates here.
     */
    abortSignal?: AbortSignal | null;
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
     * Flag to ignore comments from parsing.
     * Default is false.
     */
    ignoreComments?: boolean;
    /**
     * Flag to ignore headers and footers from parsing.
     * Default is false.
     */
    ignoreHeadersAndFooters?: boolean;
    /**
     * Flag to ignore slide masters from parsing in PowerPoint.
     * Default is false.
     */
    ignoreSlideMasters?: boolean;
    /**
     * @deprecated Notes are now structurally attached to the specific nodes they belong to via `node.notes`.
     * This option is now completely ignored by all parsers.
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
     * An optional AbortSignal to cancel the parsing operation.
     * When aborted, the parser immediately rejects with a standard AbortError (DOMException).
     * 
     * ### Format-Specific Abort Behavior:
     * - **PDF**: Checked between page loads and before individual image OCR operations.
     * - **RTF**: Checked before parsing/traversal and before running OCR on image attachments.
     * - **DOCX/XLSX/PPTX/ODF**: Checked during zip decompression before loading and parsing XML files.
     * - **CSV/MD/HTML**: Checked at the start of the parsing phase.
     * 
     * Note: If an OCR operation is currently running on a Tesseract worker when aborted,
     * the worker will be terminated and removed from the worker pool automatically to prevent leaks.
     */
    abortSignal?: AbortSignal | null;

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
     * If not provided, it defaults to `https://cdn.jsdelivr.net/npm/pdfjs-dist@6.1.200/build/pdf.worker.min.mjs`.
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
    /**
     * Limits and checks applied during ZIP extraction to protect against excessive
     * memory and resource usage.
     */
    decompressionLimits?: DecompressionLimits;
}

/**
 * Limits applied to ZIP archive decompression.
 */
export interface DecompressionLimits {
    /**
     * Maximum allowed total uncompressed size (in bytes) of files extracted from a ZIP archive.
     * Applies to OOXML (DOCX, XLSX, PPTX) and ODF (ODT, ODP, ODS) formats.
     * Default is 536870912 (512 MB).
     */
    maxUncompressedBytes?: number;
    /**
     * Maximum allowed number of entries (files and directories) in a ZIP archive.
     * Applies to OOXML (DOCX, XLSX, PPTX) and ODF (ODT, ODP, ODS) formats.
     * Default is 10000.
     */
    maxZipEntries?: number;
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
type ConversionValue<D extends UniversalGeneratorFormat> =
    D extends 'pdf' ? Uint8Array | string :
    D extends 'chunks' ? OfficeChunk[] :
    D extends 'csv' ? string | Uint8Array :
    D extends 'epub' ? Uint8Array :
    string;

export interface ConversionResult<D extends UniversalGeneratorFormat> {
    /** The actual generated content (HTML, Markdown, Text, OfficeChunk[], etc.). */
    value: ConversionValue<D>;
    /** A collection of issues (warnings/infos) generated during the process. */
    messages: OfficeIssue[];
}

/**
 * Universal formats supported by all source types for generation.
 */
export type UniversalGeneratorFormat = 'text' | 'md' | 'html' | 'pdf' | 'csv' | 'rtf' | 'chunks' | 'epub';

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
     * 4. **Async Support**: The callback can be `async`, allowing you to load external data or perform complex logic during generation.
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
    /**
     * An optional AbortSignal to cancel the generation operation.
     * When aborted, the generator immediately rejects with a standard AbortError.
     * Currently supported by PdfGenerator and ChunkingGenerator.
     */
    abortSignal?: AbortSignal | null;
}

/**
 * Destination-aware generator configuration.
 * Restricts format-specific configurations to their respective destinations.
 */
/**
 * Maps a destination format string to its corresponding specific configuration object type.
 */
type GeneratorSpecificConfig<D extends string> = 
    D extends 'html' ? { htmlConfig?: HtmlGeneratorConfig } :
    D extends 'md' ? { mdConfig?: MdGeneratorConfig } :
    D extends 'pdf' ? { pdfConfig?: PdfGeneratorConfig } :
    D extends 'csv' ? { csvConfig?: CsvGeneratorConfig } :
    D extends 'text' ? { textConfig?: TextGeneratorConfig } :
    D extends 'rtf' ? { rtfConfig?: RtfGeneratorConfig } :
    D extends 'chunks' ? { chunksConfig?: ChunkingConfig } :
    Partial<{
        htmlConfig: HtmlGeneratorConfig;
        mdConfig: MdGeneratorConfig;
        pdfConfig: PdfGeneratorConfig;
        csvConfig: CsvGeneratorConfig;
        textConfig: TextGeneratorConfig;
        rtfConfig: RtfGeneratorConfig;
        chunksConfig: ChunkingConfig;
    }>;

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
export type GeneratorConfig<D extends string = string> = CommonGeneratorConfig & GeneratorSpecificConfig<D>;

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
    htmlConfig: HtmlGeneratorConfig;
    mdConfig: MdGeneratorConfig;
    pdfConfig: PdfGeneratorConfig;
    csvConfig: CsvGeneratorConfig;
    textConfig: TextGeneratorConfig;
    rtfConfig: RtfGeneratorConfig;
}> & {
    chunksConfig: ChunkingConfig;
};



/**
 * Configuration options for granular raw HTML injections.
 */
export interface HtmlInjectionConfig {
    /** Raw HTML injected immediately after the opening <head> tag */
    headStart?: string;
    /** Raw HTML injected immediately before the closing </head> tag */
    headEnd?: string;
    /** Raw HTML injected immediately after the opening <body> tag */
    bodyStart?: string;
    /** Raw HTML injected immediately before the closing </body> tag */
    bodyEnd?: string;
}

/**
 * Granular control over which parts of the full HTML "document envelope" are emitted.
 * Shorthand: `standalone: true` == every part on (a complete document); `standalone: false` ==
 * every part off (a bare content fragment). When an object is passed, any field you omit
 * defaults to its "on" (standalone) value.
 */
export interface StandaloneConfig {
    /**
     * Wrap the output in `<!DOCTYPE html><html><head>…</head><body>…</body></html>`.
     * When false, only the inner content fragment is emitted. Defaults to true.
     */
    document?: boolean;
    /**
     * Emit `<title>` and `<meta>` tags (author, description, dates, custom properties) in the head.
     * Only meaningful when `document` is true. Defaults to true.
     */
    metaTags?: boolean;
    /**
     * How the library's built-in CSS is delivered:
     * - `'full'` — the complete premium stylesheet using global selectors (`body`, `h1`, `table`, …).
     *   This is what `standalone: true` has always emitted.
     * - `'scoped'` — the same styling, scoped under the fragment's container via CSS `@scope` so it
     *   cannot leak into a host page's own styles. Requires a modern browser engine (Chrome 118+,
     *   Safari 17.4+, Firefox 128+).
     * - `'none'` — no stylesheet is emitted; the host page (or EPUB reader, or rich-text editor)
     *   supplies its own styling.
     * The boolean shorthand for `standalone` maps `true` → `'full'`, `false` → `'none'`.
     * Defaults to `'full'`.
     */
    styles?: 'full' | 'scoped' | 'none';
    /**
     * Emit injected `<script>` tags: the Chart.js loader (when `includeCharts` is true and charts
     * are present) and the spreadsheet interactivity script. Defaults to true.
     */
    scripts?: boolean;
    /**
     * Apply `injections.headStart` / `injections.headEnd`. Only meaningful when `document` is true
     * (there is no `<head>` to inject into otherwise). Defaults to true.
     */
    headInjections?: boolean;
    /**
     * Apply `injections.bodyStart` / `injections.bodyEnd`. Applies even when generating a bare
     * fragment (`document: false`), since these wrap body *content*, not the document shell.
     * Defaults to true.
     */
    bodyInjections?: boolean;
}

/**
 * Configuration options for HTML generation.
 */
export interface HtmlGeneratorConfig {
    /**
     * Whether to wrap the output in a full HTML document structure (e.g., <html>, <head>, etc.).
     * Pass an object instead of a boolean for granular control over individual parts of the
     * envelope (document shell, meta tags, styles, scripts, injections) - see `StandaloneConfig`.
     * Defaults to true.
     */
    standalone?: boolean | StandaloneConfig;
    /**
     * URL for the Chart.js library to use when 'includeCharts' is true.
     * Defaults to 'https://cdn.jsdelivr.net/npm/chart.js'.
     */
    chartJsSrc?: string;
    /**
     * Custom container width for the generated HTML.
     * Can be a number (pixels) or string (e.g., '900px', '100%').
     * If not specified or set to 'auto', it defaults based on the content type:
     * - Spreadsheet: '100%'
     * - Presentation/Slides: '1100px'
     * - Standard Document (PDF/DOCX/RTF/etc.): '900px'
     */
    containerWidth?: string | number;
    /**
     * Custom CSS to append to the generated HTML document.
     * This CSS will be included in the `<style>` block and can be used to style 
     * custom classes added during AST manipulation or override default styles.
     */
    customCss?: string;
    /**
     * Granular injection points for custom HTML, scripts, and styles.
     */
    injections?: HtmlInjectionConfig;
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
    /**
     * Timeout in milliseconds for PDF generation.
     * Limits the time spent waiting for Puppeteer to launch, load content, and render PDF.
     * Defaults to 30000 ms (30 seconds). Set to 0 to disable.
     */
    timeout?: number;
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
        nodeType?: OfficeContentNodeType;
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
     * Defaults to true.
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
    /**
     * Timeout in milliseconds for individual embedding API calls.
     * Defaults to 10000 ms (10 seconds). Set to 0 to disable.
     */
    timeout?: number;
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
export type SupportedFileType = 'docx' | 'pptx' | 'xlsx' | 'odt' | 'odp' | 'ods' | 'pdf' | 'rtf' | 'md' | 'html' | 'csv' | 'epub';

/**
 * Types of content nodes in the AST.
 */
export type OfficeContentNodeType = 'paragraph' | 'heading' | 'table' | 'list' | 'text' | 'image' | 'chart' | 'drawing' | 'slide' | 'note' | 'sheet' | 'row' | 'cell' | 'page' | 'break' | 'code' | 'comment' | 'header' | 'footer' | 'slideMaster' | 'embed' | 'admonition' | 'definitionList' | 'definitionTerm' | 'definitionDescription';

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
 * Text alignment options.
 * Common in spreadsheet cells, paragraph styles, and text elements.
 */
export type TextAlignment = 'left' | 'center' | 'right' | 'justify';

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
    alignment?: TextAlignment;
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
    alignment?: TextAlignment;
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
    alignment?: TextAlignment;
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
    alignment: TextAlignment;

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

    /** True when this list item is a GFM task-list item (checkbox), regardless of checked state. */
    isTask?: boolean;
    /** Checked state for a task-list item. Only meaningful when isTask is true. */
    checked?: boolean;
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
    /** Background color for this cell in hex format (e.g. #FFFFFF). */
    backgroundColor?: string;
}

/**
 * Metadata for a table.
 */
export interface TableMetadata {
    /** Unique anchor IDs for internal linking. */
    anchorIds?: string[];
    /**
     * Layout alignment of the table on the page (e.g. inscript-editor's `CustomTable`).
     * @example 'center'
     */
    align?: 'left' | 'center' | 'right';
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
    /**
     * Display width of the image (e.g. inscript-editor's `CustomImage`), as a CSS length or percentage.
     * @example "50%"
     */
    width?: string;
    /**
     * Layout alignment of the image (e.g. inscript-editor's `CustomImage`).
     * @example 'center'
     */
    align?: 'left' | 'center' | 'right';
}

/**
 * Metadata for an embedded external media node (e.g. a YouTube video).
 * Markdown has no native syntax for this - see `MarkdownGenerator`'s `embed` case.
 */
export interface EmbedMetadata {
    /** The kind of embed. Only 'youtube' is supported today; the shape is generic for future providers. */
    embedType: 'youtube';
    /** The provider-specific video ID (e.g. the 11-character YouTube video ID). */
    videoId: string;
    /** The original/canonical URL of the embedded media, if known. */
    url?: string;
    /** Display width, as a CSS length or percentage. */
    width?: string;
    /** Layout alignment of the embed. */
    align?: 'left' | 'center' | 'right';
}

/**
 * Metadata for an admonition/alert node (e.g. GitHub's `> [!NOTE]` or GLFM's `:::note`).
 * `MarkdownParser` accepts both syntaxes; `MarkdownGenerator` only ever writes the
 * blockquote form. Children are block content (paragraphs) wrapped by the admonition.
 */
export interface AdmonitionMetadata {
    admonitionType: 'note' | 'tip' | 'important' | 'warning' | 'caution';
    /** Optional custom title; falls back to the type label. */
    title?: string;
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

    /**
     * When set, this text is an abbreviation and this is its full-form expansion,
     * rendered as `<abbr title="...">`. Populated from Markdown Extra's
     * `*[HTML]: Hypertext Markup Language` syntax or an HTML `<abbr>` tag.
     */
    abbreviationTitle?: string;

    /**
     * When set, this text is a Pandoc/MultiMarkdown-style citation reference
     * (`[@citekey]`), and this is the bare citekey (e.g. "smith2024"). Bibliography
     * resolution (author/year display, .bib management) is left to the consuming app.
     */
    citationKey?: string;

    /**
     * True when this is an Obsidian-style wikilink (`[[page]]` / `[[page|alias]]`).
     * `link` holds the bare page name and `linkType` is always 'internal'; the
     * per-workspace enable/disable toggle lives in markdownwriter, not here -
     * officeParser always parses/generates the syntax.
     */
    wikilink?: boolean;
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
    /** The slide number this note is associated with (used in PowerPoint). */
    slideNumber?: number;
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
    /**
     * When set, this node is a LaTeX math expression rather than a code block. `node.text`
     * holds the bare LaTeX (delimiters excluded); 'inline' round-trips as `$...$`,
     * 'block' as `$$...$$`. Matches inscript-editor's math node (Roadmap Step 11.5).
     */
    math?: 'inline' | 'block';
}

/**
 * Metadata for a comment/annotation.
 */
export interface CommentMetadata {
    author?: string;
    initials?: string;
    date?: string;
    commentId?: string;
}

/**
 * Metadata for a header or footer.
 */
export interface HeaderFooterMetadata {
    type: 'default' | 'first' | 'even' | string;
}

/**
 * Union type for content metadata.
 */
export type ContentMetadata = SlideMetadata | SheetMetadata | HeadingMetadata | ListMetadata | CellMetadata | ImageMetadata | ChartMetadata | PageMetadata | ParagraphMetadata | TextMetadata | NoteMetadata | BreakMetadata | CodeMetadata | CommentMetadata | HeaderFooterMetadata | TableMetadata | EmbedMetadata | AdmonitionMetadata | undefined;


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
/**
 * Shared properties available on all document content nodes.
 */
export interface BaseContentNode {
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
     * Comments attached to this specific node.
     * Keeps annotations completely separate from the actual content flow.
     */
    comments?: OfficeContentNode[];

    /**
     * Notes (like footnotes or slide notes) attached to this specific node.
     * Keeps notes separate from the actual structural children.
     */
    notes?: OfficeContentNode[];

    /**
     * Text formatting applied to this node.
     * Only applicable to text-containing nodes.
     * For container nodes like paragraphs, formatting typically appears on child text nodes.
     * @example { bold: true, size: "12", font: "Arial" }
     */
    formatting?: TextFormatting;

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
export type OfficeContentNode = BaseContentNode & (
    | { type: 'slide'; metadata?: SlideMetadata }
    | { type: 'sheet'; metadata?: SheetMetadata }
    | { type: 'heading'; metadata?: HeadingMetadata }
    | { type: 'list'; metadata?: ListMetadata }
    | { type: 'cell'; metadata?: CellMetadata }
    | { type: 'image'; metadata?: ImageMetadata }
    | { type: 'chart'; metadata?: ChartMetadata }
    | { type: 'page'; metadata?: PageMetadata }
    | { type: 'paragraph'; metadata?: ParagraphMetadata }
    | { type: 'text'; metadata?: TextMetadata }
    | { type: 'note'; metadata?: NoteMetadata }
    | { type: 'break'; metadata?: BreakMetadata }
    | { type: 'code'; metadata?: CodeMetadata }
    | { type: 'comment'; metadata?: CommentMetadata }
    | { type: 'header'; metadata?: HeaderFooterMetadata }
    | { type: 'footer'; metadata?: HeaderFooterMetadata }
    | { type: 'table'; metadata?: TableMetadata }
    | { type: 'row'; metadata?: undefined }
    | { type: 'drawing'; metadata?: undefined }
    | { type: 'slideMaster'; metadata?: SlideMetadata }
    | { type: 'embed'; metadata?: EmbedMetadata }
    | { type: 'admonition'; metadata?: AdmonitionMetadata }
    | { type: 'definitionList'; metadata?: undefined }
    | { type: 'definitionTerm'; metadata?: undefined }
    | { type: 'definitionDescription'; metadata?: undefined }
);

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
    /** Keywords associated with the document. */
    keywords?: string;
    /** 
     * Contains all format-specific metadata fields extracted verbatim.
     * Consumers can use this to access properties not mapped to the standard OfficeMetadata fields.
     * Examples: all <meta> tags in HTML, app.xml properties in DOCX, XMP dicts in PDF.
     */
    nativeProperties?: Record<string, any>;
}

/**
 * Contains out-of-band layout elements and templates that are not part of the main document flow.
 */
export interface OfficeAuxiliaryContent {
    /** Headers extracted from the document. */
    headers?: OfficeContentNode[];
    /** Footers extracted from the document. */
    footers?: OfficeContentNode[];
    /** Slide Masters extracted from presentations. */
    slideMasters?: OfficeContentNode[];
}

/**
 * The Root Abstract Syntax Tree (AST) representing a parsed Office Document.
 * This is the ultimate output of `OfficeParser.parseOffice()`.
 * 
 * DESIGN PHILOSOPHY:
 * The AST is designed to be a universal, format-agnostic representation of document content.
 * Whether the input was a PDF, DOCX, XLSX, Markdown, or HTML file, the resulting AST
 * uses the same consistent structure (`OfficeContentNode` trees).
 * 
 * ### Key Top-Level Properties:
 * - `metadata`: Document-level properties (author, title, stats).
 * - `content`: The main sequential flow of the document (paragraphs, tables, slides, sheets).
 * - `attachments`: Extracted binary assets (images, embedded files).
 * - `auxiliary`: Out-of-band layout/template elements (headers, footers, slide masters).
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
     * Out-of-band layout and template elements that are not part of the main text flow.
     * Extracted only if the respective `ignore...` config flags are false.
     * Contains elements like `headers`, `footers`, and `slideMasters`.
     */
    auxiliary?: OfficeAuxiliaryContent;

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

declare global {
    const __SLIM__: boolean | undefined;
}


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
    /** The document is encrypted (PDF, or a password-protected OOXML/ODF) and no password was supplied */
    PASSWORD_REQUIRED = 'PASSWORD_REQUIRED',
    /** The supplied password did not decrypt the document */
    PASSWORD_INCORRECT = 'PASSWORD_INCORRECT',
    /** An encrypted document could not be decrypted for a structural reason (malformed container, unsupported cipher/scheme) */
    DOCUMENT_DECRYPTION_FAILED = 'DOCUMENT_DECRYPTION_FAILED',
    /** A template render was given a document type it cannot template */
    TEMPLATE_UNSUPPORTED_FORMAT = 'TEMPLATE_UNSUPPORTED_FORMAT',
    /** A template placeholder had no matching data field and `onMissing: 'error'` was set */
    TEMPLATE_FIELD_MISSING = 'TEMPLATE_FIELD_MISSING',
    /** PDF generation failed (e.g. the puppeteer engine is unavailable or rendering errored) */
    PDF_GENERATION_FAILED = 'PDF_GENERATION_FAILED',
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
    /** ZIP data yielded no readable entries (corrupt, truncated, or not a ZIP archive) */
    ZIP_NO_ENTRIES_FOUND = 'ZIP_NO_ENTRIES_FOUND',
    /** ZIP data is truncated: the End of Central Directory record is absent */
    ZIP_TRUNCATED = 'ZIP_TRUNCATED',
    /** A readable ZIP archive is missing the part its document format requires */
    REQUIRED_PART_MISSING = 'REQUIRED_PART_MISSING',
    /** Document element/structure nesting exceeded the safe recursion depth */
    MAX_NESTING_DEPTH_EXCEEDED = 'MAX_NESTING_DEPTH_EXCEEDED',
    /** Embedding call timed out */
    EMBEDDING_TIMEOUT = 'EMBEDDING_TIMEOUT',
    /** OCR workers were terminated (`terminateOcr()`) before an image was recognized; the parse reports it as OCR_FAILED */
    OCR_TERMINATED = 'OCR_TERMINATED'
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
    /** An image exceeded `maxInlineImageBytes` and was emitted as a name reference instead of inlined */
    IMAGE_NOT_INLINED = 'IMAGE_NOT_INLINED',
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
    INVALID_CONTAINER_WIDTH = 'INVALID_CONTAINER_WIDTH',
    /** A document's repeated-cell expansion hit the configured cell limit and was truncated */
    TABLE_CELL_LIMIT_EXCEEDED = 'TABLE_CELL_LIMIT_EXCEEDED',
    /** A metadata override could not be represented in the destination format's vocabulary */
    METADATA_NOT_REPRESENTABLE = 'METADATA_NOT_REPRESENTABLE',
    /** A content feature (e.g. math, an embedded object) has no faithful representation in the destination format and was downgraded or dropped */
    CONTENT_NOT_REPRESENTABLE = 'CONTENT_NOT_REPRESENTABLE',
    /** A styleMap output.tag was not an allowed element name and was ignored */
    INVALID_STYLE_MAP_TAG = 'INVALID_STYLE_MAP_TAG',
    /** A workbook archive contains no worksheet parts (chartsheet-only workbooks are legitimate) */
    NO_WORKSHEETS_FOUND = 'NO_WORKSHEETS_FOUND',
    /** A presentation archive contains no slides (a zero-slide presentation is legitimate) */
    NO_SLIDES_FOUND = 'NO_SLIDES_FOUND',
    /** A PDF's tagged-structure tree was absent, incomplete, or flagged unreliable; heuristics were used instead */
    PDF_STRUCT_TREE_UNRELIABLE = 'PDF_STRUCT_TREE_UNRELIABLE',
    /** A PDF page yielded mostly unmappable glyphs (broken/missing ToUnicode); extracted text is likely garbage */
    PDF_TEXT_ENCODING_SUSPECT = 'PDF_TEXT_ENCODING_SUSPECT',
    /** A PDF yielded essentially no text; it is very likely a scanned/image-only document needing OCR */
    PDF_NO_TEXT_EXTRACTED = 'PDF_NO_TEXT_EXTRACTED',
    /** A PDF's bookmark outline was cut short by the depth/size cap or could not be read; `ast.auxiliary.outline` holds only what was recovered */
    PDF_OUTLINE_TRUNCATED = 'PDF_OUTLINE_TRUNCATED',
    /** `ocr: true` was set without `extractAttachments: true`; OCR runs over extracted images (in every format), so no OCR was performed */
    OCR_REQUIRES_ATTACHMENTS = 'OCR_REQUIRES_ATTACHMENTS',
    /** A config option was passed that this version does not recognize (e.g. a key renamed in a major release); it had no effect */
    UNRECOGNIZED_CONFIG_OPTION = 'UNRECOGNIZED_CONFIG_OPTION',
    /**
     * A generator option was given a value it does not accept (a document class, paper format,
     * margin, PDF engine, HTML stylesheet mode, Markdown dialect or chunking choice that is not one of its
     * choices). The option's default was used instead; the message names the option, the value and
     * what it accepts.
     */
    INVALID_CONFIG_VALUE = 'INVALID_CONFIG_VALUE',
    /** A math expression used an unsafe LaTeX command or was malformed, so it was written to LaTeX output as literal text rather than typeset math */
    MATH_WRITTEN_AS_TEXT = 'MATH_WRITTEN_AS_TEXT',
    /**
     * LaTeX output references image files that are not part of it: without `texConfig.bundle`, the
     * images the `.tex` does not carry inside it (with `texConfig.embedImages` off, an image other than
     * a readable PNG or JPEG, or one past the decoding limits; their bytes are in `ast.attachments`),
     * and, in either mode, images the source referred to only by a relative path, with no data to
     * package. The message names each.
     */
    IMAGES_NOT_BUNDLED = 'IMAGES_NOT_BUNDLED',
    /** LaTeX input used commands or environments the parser does not interpret; their text content was kept where it had any */
    LATEX_CONSTRUCT_NOT_INTERPRETED = 'LATEX_CONSTRUCT_NOT_INTERPRETED',
    /** LaTeX input hit a macro-expansion, file-inclusion or nesting-depth limit; the rest of the affected construct was not expanded (content nested past the limit is kept as plain text) */
    LATEX_EXPANSION_LIMIT_REACHED = 'LATEX_EXPANSION_LIMIT_REACHED',
    /** LaTeX input referenced files (`\input`, `\include`, `\includegraphics`) that were not available to the parser */
    LATEX_FILE_NOT_FOUND = 'LATEX_FILE_NOT_FOUND'
}

/**
 * Consolidated timeout settings for OCR operations.
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
     * Reconstruct the recognized text's two-dimensional page layout from Tesseract's per-word
     * bounding boxes, instead of returning the flat, linearized string. Words keep their relative
     * horizontal position (right-hand text stays on the right, columns line up) and lines/blocks keep
     * their vertical order and gaps, so a scanned table, form or multi-column page reads spatially -
     * the same idea as `textConfig.preserveLayout` for born-digital PDF text. The block is
     * left-normalized so there is no large leading indent.
     *
     * Turn it off to get Tesseract's flat reading-order text (a single space between words, one line
     * per line). Simple images (a logo, a caption) read almost identically either way.
     *
     * Default is true.
     */
    preserveLayout?: boolean;
    /**
     * Consolidated timeout settings for all OCR operations.
     */
    timeout?: OcrTimeoutConfig;
    /**
     * An optional AbortSignal propagated from the main parser configuration to abort active OCR jobs.
     * If the signal is aborted:
     * 1. Any pending OCR jobs in the scheduler queue are rejected immediately.
     * 2. Any active OCR job running on a Tesseract worker will reject, the worker will be
     *    terminated, and it will be removed from the pool to avoid hanging worker threads.
     * 3. The parse rejects with an AbortError, as it does for the top-level `abortSignal`.
     * 
     * Developers should prefer passing this at the top level of `parseOffice` (as `config.abortSignal`),
     * which automatically propagates here.
     */
    abortSignal?: AbortSignal | null;
}

/**
 * Configuration options shared across every input format.
 */
export interface CommonOfficeParserConfig {
    /**
     * Callback for warnings or non-fatal errors encountered during parsing.
     * Allows you to capture issues like OCR failures or attachment extraction errors
     * without stopping the parsing process.
     */
    onWarning?: (issue: OfficeIssue) => void;
    /**
     * Password for a password-protected document. Applies to every format that supports encryption:
     * PDF, encrypted OOXML (`.docx`/`.xlsx`/`.pptx`, ECMA-376 agile or standard AES), and encrypted
     * ODF (`.odt`/`.ods`/`.odp`/`.odg`, AES-CBC). Ignored for unencrypted files.
     *
     * When a document is encrypted and neither this nor `onPassword` yields a working password,
     * parsing rejects with `PASSWORD_REQUIRED` (none supplied) or `PASSWORD_INCORRECT` (supplied but
     * wrong).
     *
     * Default is '' (no password).
     */
    password?: string;
    /**
     * Callback invoked when an encrypted document needs a password that `password` did not satisfy,
     * so it can be supplied lazily or interactively (a prompt, a vault lookup) instead of up front.
     * Applies to every encryptable format (PDF/OOXML/ODF); mirrors pdf.js's own `onPassword` hook.
     *
     * Called with `'required'` when the document is encrypted and no password was given, or
     * `'incorrect'` when the last attempt was wrong. Return a password (sync or async) to retry;
     * return `undefined`/`''` to stop, in which case parsing rejects with `PASSWORD_REQUIRED` or
     * `PASSWORD_INCORRECT` as it would with no callback. The callback is asked at most 3 times in total
     * (shared across `'required'`/`'incorrect'`), so one that keeps returning a wrong password cannot
     * loop forever.
     *
     * By default this is unset, so an encrypted document without a valid `password` simply throws:
     * an undecryptable document is unrecoverable for that call, so it is an error rather than a
     * warning. The callback is the escape hatch for handling it gracefully.
     */
    onPassword?: (reason: 'required' | 'incorrect') => string | undefined | Promise<string | undefined>;
    /**
     * The delimiter used for every new line in places that allow multiline text like word.
     * Default is \n.
     */
    newlineDelimiter?: string;
    /**
     * Flag to ignore notes from parsing. Default is false (notes are extracted).
     * Applies to: footnotes/endnotes (DOCX, ODT, RTF, PDF tagged, HTML, Markdown, EPUB) and speaker
     * notes (PPTX, ODP). No effect on formats that carry no notes (XLSX, ODS, ODG, CSV).
     */
    ignoreNotes?: boolean;
    /**
     * Flag to ignore comments from parsing. Default is false (comments are extracted onto `node.comments`).
     * Applies to: DOCX, XLSX, PPTX, every ODF type (ODT/ODS/ODP/ODG) and LaTeX (`% Comment (Author,
     * date): text` lines). Not applicable to PDF, RTF or EPUB (no comments are parsed there). Source-level
     * comments are not governed by this flag: the CSV `#`-row convention produces top-level `comment`
     * nodes, and Markdown/HTML `<!-- ... -->` (and LaTeX `% <!-- ... -->` lines) produce `comment` nodes
     * with `metadata.sourceSyntax: 'html'` (HTML and EPUB only under `HtmlParserConfig.preserveComments`).
     */
    ignoreComments?: boolean;
    /**
     * Flag to ignore headers and footers from parsing. Default is false (they are extracted into
     * `ast.auxiliary.headers`/`.footers`). Extracted for DOCX, PDF (the top/bottom running bands) and
     * ODT (Writer master pages). It is a no-op for ODS/ODP/ODG, XLSX, PPTX and RTF, where running
     * headers/footers are not extracted at all.
     */
    ignoreHeadersAndFooters?: boolean;
    /**
     * Flag to ignore slide masters from parsing. Default is false. PPTX only; ODP master pages are not
     * extracted, so the flag is a no-op for ODP.
     */
    ignoreSlideMasters?: boolean;
    /**
     * Flag to extract attachments like images, charts, etc. into `ast.attachments` (Base64). Default
     * is false. It is also the switch that lets image bytes reach any generated output (HTML/EPUB embed,
     * DOCX/ODT media, native PDF) and that OCR (`ocr: true`) runs over. HTML and Markdown input capture
     * images only from `data:` URIs; a remote-URL image is referenced, not fetched.
     */
    extractAttachments?: boolean;
    /**
     * Flag to include raw content (XML for XML-based formats, RTF for RTF) in the AST.
     * Default is false.
     */
    includeRawContent?: boolean;
    /**
     * Flag to enable OCR for images. Default is false. OCR runs over EXTRACTED images, so it requires
     * `extractAttachments: true` as well - in every format (PDF page images, and embedded images in
     * DOCX/PPTX/XLSX/ODF/RTF/HTML/Markdown/EPUB). Setting `ocr: true` alone performs no OCR and raises
     * an `OCR_REQUIRES_ATTACHMENTS` warning. Uses Tesseract.js (`ocrConfig.language`, default 'eng').
     */
    ocr?: boolean;
    /**
     * Shared OCR configuration for worker pooling and offline support, including the recognition
     * language (`ocrConfig.language`, default `'eng'`).
     */
    ocrConfig?: OcrConfig;
    /**
     * An optional AbortSignal to cancel the parsing operation.
     * When aborted, the parser rejects with a standard AbortError (DOMException). Once the signal has
     * fired, the parse never resolves, including when it fires while OCR is recognizing an image
     * (that is a cancellation, not an `OCR_FAILED` recognition).
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
     * The URL/path to the PDF.js worker script, used only on the PDF path in the browser.
     *
     * Defaults to `https://cdn.jsdelivr.net/npm/pdfjs-dist@6.2.108/build/pdf.worker.min.mjs`, so you only
     * need to set it to point at a self-hosted copy when the CDN is unreachable or a CSP blocks it. The
     * slim browser bundle ships no default, so there it is required for PDF parsing. Ignored in Node.
     */
    pdfWorkerSrc?: string;
    /**
     * Flag to include break nodes in the AST. Default is false.
     * Applies to Word (`w:br`/`w:cr`) and ODF (`fo:break-before`/`fo:break-after`, `text:soft-page-break`),
     * where breaks are otherwise invisible. HTML and Markdown always emit `break` nodes (a `<br>`/hard
     * line break, and `<hr>`/`---` as a `thematic`/`page` break), since a break is content there; this
     * flag does not gate those.
     */
    includeBreakNodes?: boolean;
    /**
     * Flag to ignore all internal (anchor) links during parsing.
     * When true, all bookmarks, cross-references, and internal document jumps are stripped
     * from the AST. Only external URLs will be preserved. Honored at parse time by DOCX, ODF and PDF;
     * HTML/Markdown anchors are handled by the generator-side flag of the same name instead.
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
     * Besides the formats' own names it accepts the {@link FileTypeAlias} names (`latex`, `ltx`, the
     * ODF template extensions, and `zip`).
     * 
     * Default is null.
     */
    fileType?: SupportedFileType | FileTypeAlias | null;
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
    /**
     * Omit the geometric layout data ("where on the page is this?") from the AST.
     *
     * By default a parser that knows the geometry attaches, to every content node, a
     * {@link NodeBounds} box (`node.bounds`: the node's `{ x, y, width, height }` rectangle on its
     * page, in points) and, to each page node, the page dimensions ({@link PageMetadata.pageWidth},
     * `pageHeight`, `rotation`). Today only the PDF parser produces any of this. Set this flag to
     * drop all of it, giving a smaller AST that matches the pre-8.0 output shape.
     *
     * Note this also turns off the layout-faithful `.to('text')` rendering, which needs the boxes to
     * align columns and tables; without them the text generator falls back to plain flowing text.
     *
     * Default is false (bounds are emitted).
     */
    ignorePageGeometry?: boolean;
}

/**
 * Format-specific options for PDF parsing.
 *
 * Mirrors {@link HtmlParserConfig}: everything intrinsically PDF-only lives here, while
 * cross-format flags (e.g. `ignorePageGeometry`, `ignoreInternalLinks`, `ignoreHeadersAndFooters`)
 * stay flat on {@link CommonOfficeParserConfig}.
 */
export interface PdfParserConfig {
    /**
     * Use the PDF's tagged-structure tree (headings, tables, lists, notes) when the document
     * declares one and it passes reliability checks.
     *
     * When false, or when the tree is missing or flagged unreliable, all structure is recovered from
     * geometry instead (column detection, line and paragraph reconstruction). Set false to force
     * the geometric path even on tagged PDFs.
     *
     * This is the whole-tree switch. To keep tagged tables/lists/notes but override only how heading
     * levels are decided, leave this true and use `headingDetection: 'font-size'` (see below) instead
     * of turning tags off entirely.
     *
     * Default is true.
     */
    useTags?: boolean;
    /**
     * Detect columns and floating blocks (via a recursive XY-cut) so multi-column and
     * float-beside-text pages read in the correct order.
     *
     * When false, each page is read strictly top-to-bottom, left-to-right, which is faster but
     * interleaves columns.
     *
     * Only used on the geometric path (untagged PDFs, or when `useTags` is false).
     *
     * Default is true.
     */
    detectColumns?: boolean;
    /**
     * Merge words split across a line break by a trailing hyphen (e.g. "extre-\nmely" -> "extremely").
     * Soft hyphens (U+00AD) are always removed; a literal hyphen is dropped only before a lowercase
     * continuation.
     *
     * Default is true.
     */
    mergeHyphenatedWords?: boolean;
    /**
     * Tolerance, as a fraction of font size, for grouping text fragments onto the same visual line
     * by their baseline. Larger values are more forgiving of baseline jitter but risk merging
     * adjacent lines.
     *
     * Default is 0.35.
     */
    lineToleranceFactor?: number;
    /**
     * Threshold, as a fraction of font size, for the horizontal gap that triggers inserting a space
     * between two adjacent text fragments. Larger values insert fewer spaces.
     *
     * Default is 0.25.
     */
    spaceToleranceFactor?: number;
    /**
     * How heading levels are decided. This is independent of `useTags`: it re-decides only the
     * heading level, while `useTags` still governs whether tables/lists/notes come from the tags.
     * - 'auto': trust the source. On the tagged path, a heading's level comes from its tag (`H1`..
     *   `H6`); on the geometric path (untagged, or `useTags: false`), from a font-size and weight
     *   heuristic.
     * - 'font-size': re-level headings with the size/weight heuristic even on a tagged PDF. Tagged
     *   tables and lists are still honored, and a tagged heading (`H`/`H1`..`H6`) is re-leveled by its
     *   font size (and demoted to a paragraph when it is not visually heading-like); a block the tags
     *   call a plain paragraph is left as one. Use this when a PDF's heading tags are present but wrong
     *   or flat.
     * - 'off': never emit headings; every block is a paragraph.
     *
     * Default is 'auto'.
     */
    headingDetection?: 'auto' | 'font-size' | 'off';
    /**
     * Restrict parsing to a subset of pages, e.g. '1-3,7' or '2,4,6'. Pages are 1-based and the
     * output keeps their original page numbers. An empty string parses all pages; an unparseable or
     * out-of-order value is ignored and all pages are parsed.
     *
     * Default is '' (all pages).
     */
    pageRange?: string;
    /**
     * Unicode-normalize the text pdf.js extracts: ligatures are expanded, combining marks composed,
     * and original whitespace regularized, so the output is clean, searchable text. Turn this off to
     * get the raw source glyphs verbatim (ligatures, combining marks and whitespace preserved) when
     * you need byte-faithful fidelity to what the PDF actually stores.
     *
     * Default is true (text is normalized).
     */
    normalizeText?: boolean;
    /**
     * Extract each text run's fill color into `formatting.color` (hex `#rrggbb`). pdf.js exposes no
     * color on its text content, so this is recovered from the page's operator list. On the default
     * text path the operator list is only fetched for the first page(s) that introduce a font, so
     * turning this on fetches it for every page instead: about 1.6x parse time on a text-heavy PDF,
     * and near-free when `extractAttachments` already fetches the operator list. Color is part of a document's
     * content, like bold or font, so it is extracted by default; set this to `false` to skip it on a
     * throughput-focused text/RAG path that does not need color. Pure black (`#000000`) is treated as
     * the default and left unset, so only genuinely colored text carries a `color`, mirroring how the
     * DOCX/RTF parsers report it. Runs painted with a pattern, shading or transparent fill are left
     * uncolored rather than guessed.
     *
     * Highlight annotations are unaffected by this flag: they populate `formatting.backgroundColor`
     * regardless, since they come from the annotation list that is already read for hyperlinks.
     *
     * Default is true.
     */
    extractTextColor?: boolean;
}

/**
 * Format-specific options for HTML (and XHTML/EPUB, which parse through the same code path).
 *
 * Note there is deliberately no `MdParserConfig`: the Markdown parser populates its
 * dialect-provenance metadata (e.g. `AdmonitionMetadata.sourceSyntax`) unconditionally because
 * doing so costs nothing and changes no existing field's value, so it has nothing to configure.
 * An empty placeholder interface would be worse than useless here - `interface X {}` accepts any
 * non-nullish value in TypeScript, so `mdParserConfig: 5` would type-check.
 */
export interface HtmlParserConfig {
    /**
     * Preserve source HTML attributes that no typed metadata field consumed, on
     * `OfficeContentNode.htmlAttributes`, so they can be replayed on generation.
     *
     * Off by default: with it off nothing is populated, so the AST is byte-identical to previous
     * releases, and the attribute-replay surface stays something a consumer opts into rather than
     * something switched on for every existing caller. Captured values are sanitized on the way in
     * *and* on the way out - see `BaseContentNode.htmlAttributes`.
     *
     * Defaults to false.
     */
    preserveAttributes?: boolean;
    /**
     * Preserve `<iframe>` embeds that aren't recognized as a known provider (YouTube is always
     * recognized). By default every non-YouTube iframe is dropped, which is a deliberate security
     * posture other consumers rely on; set this to opt back in. `true` preserves any iframe; an
     * array is a hostname allowlist (an entry matches the src's host exactly or as a `.`-suffix,
     * so `"vimeo.com"` also matches `player.vimeo.com`). Preserved iframes become `embed` nodes
     * with `embedType: 'iframe'`; on generation the `src` is still scheme-checked (only http/https
     * survive). This also governs a raw `<iframe>` block encountered in Markdown input.
     *
     * Defaults to false.
     */
    preserveIframes?: boolean | string[];
    /**
     * Preserve HTML comments (`<!-- ... -->`) found in HTML and EPUB input as `comment` nodes with
     * `metadata.sourceSyntax: 'html'`, so they survive an HTML -> AST -> HTML/Markdown round trip
     * instead of being dropped. Conditional comments (`<!--[if ...]> ... <![endif]-->`, Office/IE
     * directives rather than authored notes) are always dropped. Off by default: HTML in the wild
     * carries many comments (analytics markers, build stamps), and emitting them into converted
     * Markdown is something a consumer opts into. The `data-html-comment` shape that
     * `HtmlGeneratorConfig.sourceAttributes` emits is always read, independent of this option.
     * Comments in Markdown input are always parsed as comment nodes.
     *
     * Defaults to false.
     */
    preserveComments?: boolean;
    /**
     * Import ambiguous "folk" embed forms in Markdown as embeds: a standalone Obsidian-style image
     * whose URL is a YouTube link (`![](https://youtube.com/watch?v=ID)`), and the clickable
     * thumbnail-link (`[![alt](https://img.youtube.com/vi/ID/…)](watch-url)`). Both become a
     * `embedType: 'youtube'` embed (rendered from the validated id, so it is safe). Off by default:
     * auto-upgrading an image/link to an embed is a heuristic that could mangle a genuinely-intended
     * image link, so a consumer opts in. The unambiguous forms (`<div data-youtube-video>`, a bare
     * YouTube `<iframe>`, the `::youtube` directive) are always recognized, independent of this flag.
     *
     * Defaults to false.
     */
    embedFolkForms?: boolean;
}

/**
 * Format-specific options for LaTeX parsing (a `.tex` file or a LaTeX project zip).
 */
export interface TexParserConfig {
    /**
     * What `\today` prints. LaTeX prints the date the document is compiled; by default the parser
     * does the same with the date of the parse, written the way the document's language writes
     * a date ("September 25, 2026" in English, "25. September 2026" in German). Set a string to
     * print that instead: a fixed date, so that parsing the same file gives the same output on
     * every day, or a placeholder of your own to find and replace later.
     *
     * Default is '' (the date of the parse).
     */
    today?: string;
}

/**
 * Maps an input format string to its corresponding format-specific parser configuration, mirroring
 * `GeneratorSpecificConfig<D>` on the generator side. Unlike the generator side, the input format is
 * usually runtime-detected rather than known statically at the `parseOffice()` call site, so this
 * mainly exists for internal typing/extensibility rather than compile-time narrowing per call.
 */
type ParserSpecificConfig<F extends string> =
    F extends 'html' | 'epub' ? { htmlParserConfig?: HtmlParserConfig } :
    F extends 'pdf' ? { pdfParserConfig?: PdfParserConfig } :
    F extends 'tex' ? { texParserConfig?: TexParserConfig } :
    Partial<{ htmlParserConfig: HtmlParserConfig; pdfParserConfig: PdfParserConfig; texParserConfig: TexParserConfig }>;

/**
 * Configuration options for the OfficeParser.
 */
export type OfficeParserConfig<F extends string = string> = CommonOfficeParserConfig & ParserSpecificConfig<F>;

/**
 * Limits applied to ZIP archive decompression.
 */
export interface DecompressionLimits {
    /**
     * Maximum allowed total uncompressed size (in bytes) of files extracted from a ZIP archive.
     * Applies to every ZIP-backed input: OOXML (DOCX, XLSX, PPTX), ODF (ODT, ODS, ODP, ODG), EPUB and LaTeX project zips.
     * Default is 536870912 (512 MB).
     */
    maxUncompressedBytes?: number;
    /**
     * Maximum allowed number of entries (files and directories) in a ZIP archive.
     * Applies to every ZIP-backed input: OOXML (DOCX, XLSX, PPTX), ODF (ODT, ODS, ODP, ODG), EPUB and LaTeX project zips.
     * Default is 10000.
     */
    maxZipEntries?: number;
    /**
     * Maximum number of table cells materialized from a single document.
     *
     * ODF encodes runs of identical cells and rows with `table:number-columns-repeated` and
     * `table:number-rows-repeated` rather than repeating the markup, so a few hundred bytes of XML
     * can ask the parser to build an arbitrary number of nodes - and because the two multiply, a
     * row repeat times a column repeat compounds it. The ZIP limits above cannot catch this: the
     * XML is tiny before decompression and the expansion happens afterwards, while building the
     * AST.
     *
     * Real documents are nowhere near this. The repeat counts LibreOffice writes are large
     * (`number-rows-repeated="1048566"` is routine) but they sit on *empty* trailing runs, which
     * are skipped for spreadsheets; the bundled fixtures top out around 350 cells.
     *
     * On reaching the limit the parser stops materializing further cells, emits a
     * `TABLE_CELL_LIMIT_EXCEEDED` warning, and returns what it has rather than throwing, so a
     * genuinely enormous sheet still yields usable output. Raise it if you routinely process
     * spreadsheets larger than this; note the memory cost scales with it.
     *
     * Default is 1000000.
     */
    maxTableCells?: number;
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
 * An Error thrown by OfficeParser, carrying the structured issue that produced it.
 *
 * Catching code can branch on `error.officeIssue.code`, the same stable enum used for warnings,
 * instead of matching against message text. Errors that originate outside the library (and
 * `AbortError`, which is deliberately re-thrown untouched so cancellation stays detectable via
 * `error.name`) do not carry this property, hence the optional marker.
 *
 * @example
 * ```typescript
 * try {
 *     await parseOffice(buffer, { fileType: 'docx' });
 * } catch (err) {
 *     if ((err as OfficeError).officeIssue?.code === OfficeErrorType.REQUIRED_PART_MISSING) {
 *         // the archive is readable, but it is not a docx
 *     }
 * }
 * ```
 */
export interface OfficeError extends Error {
    /** The structured issue this error was created from. */
    officeIssue?: OfficeIssue;
}

/**
 * The result of a document conversion operation.
 */
type ConversionValue<D extends UniversalGeneratorFormat> =
    D extends 'pdf' ? Uint8Array | string :
    D extends 'chunks' ? OfficeChunk[] :
    D extends 'csv' ? string | Uint8Array :
    D extends 'epub' ? Uint8Array :
    D extends 'docx' ? Uint8Array :
    D extends 'odt' ? Uint8Array :
    D extends 'tex' ? string | Uint8Array :
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
export type UniversalGeneratorFormat = 'text' | 'md' | 'html' | 'pdf' | 'csv' | 'rtf' | 'chunks' | 'epub' | 'docx' | 'odt' | 'tex';

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
/**
 * Per-field overrides for the metadata written into generated output.
 *
 * Field names mirror `OfficeMetadata` so the same vocabulary describes what was parsed and what
 * gets written. Only the fields generators can actually represent are listed; arbitrary
 * caller-defined entries go in `custom`.
 *
 * **Not every format can represent every field.** HTML (`<meta>`) and Markdown (YAML frontmatter)
 * accept anything; EPUB's OPF is a closed Dublin Core vocabulary and RTF's `\info` group has a
 * fixed set of control words, so a `custom` entry has nowhere to go in those. Rather than
 * silently dropping it, generators report the loss through `onWarning`
 * (`OfficeWarningType.MetadataNotRepresentable`) and continue.
 */
export interface MetadataOverrides {
    /** Document title. */
    title?: string;
    /** Document author. */
    author?: string;
    /** Description/comments. */
    description?: string;
    /** Subject/topic. */
    subject?: string;
    /** Keywords. */
    keywords?: string;
    /** User who last modified the document. */
    lastModifiedBy?: string;
    /** Creation date. */
    created?: Date;
    /**
     * Last modification date.
     *
     * EPUB writes it as the required `dcterms:modified` property and as the mtime on every zip
     * entry. When unset, the source document's own `metadata.modified` is used, falling back to
     * the current time only if the document has none.
     */
    modified?: Date;
    /**
     * Language tag (e.g. `'en'`, `'de-DE'`). Written as HTML `lang`, EPUB, DOCX and ODT
     * `dc:language`, the PDF `/Lang` entry, and LaTeX `pdflang`.
     */
    language?: string;
    /**
     * Arbitrary caller-defined key/value pairs, kept in their own bucket rather than mixed in
     * beside the named fields above: with a bare index signature a typo like `titel` would
     * silently become a custom entry instead of a compile error.
     *
     * Written where the format allows it (HTML `<meta name="custom:KEY">`, Markdown frontmatter);
     * reported via `onWarning` where it does not (EPUB, RTF).
     */
    custom?: Record<string, string | number | boolean | Date>;
}

/**
 * How a generator renders an image node. See {@link CommonGeneratorConfig.includeImages}.
 * "OCR text" is the image node's recognized text (populated for scanned PDF images).
 */
export type ImageMode = 'image-only' | 'image+ocr-text' | 'ocr-text-only' | 'none';

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
     * Set to false for clean, semantic output. Defaults to true.
     * Applies to the formatting-carrying generators: HTML, Markdown, DOCX, ODT and RTF (text/CSV/chunks
     * carry no run formatting, so it is a no-op there).
     */
    includeFormatting?: boolean;
    /**
     * Whether to automatically generate unique slug-based IDs for headings.
     * Useful for table-of-contents and anchor links. Defaults to true.
     * Applies to HTML, Markdown, DOCX and ODT (the formats that carry an anchor/bookmark id).
     */
    generateIds?: boolean;
    /**
     * Whether to render document metadata (title, author, etc.) as visible content
     * in the generated output (e.g., a header block in HTML or plain text).
     * Structural metadata (HTML <meta> tags, Markdown YAML frontmatter) is always included.
     * Defaults to false. Rendered by CSV, DOCX, HTML (and the Puppeteer PDF engine), EPUB, text, ODT
     * and RTF; the native PDF engine and Markdown do not render it as visible content.
     */
    renderMetadata?: boolean;
    /**
     * Overrides for the document metadata written into the generated output, applied on top of
     * `ast.metadata`.
     *
     * Merged **per field**, so setting only `modified` leaves the parsed title, author, and
     * everything else intact. Every field is optional; an omitted field keeps the source
     * document's value.
     *
     * These are output overrides only - `ast.metadata` itself is never mutated, so the same AST
     * can be generated repeatedly with different metadata.
     *
     * @example Set the modification date written into the output
     * ```typescript
     * await ast.to('epub', { metadataOverrides: { modified: new Date('2024-01-01T00:00:00Z') } });
     * ```
     * @example Rebrand the output without touching the parsed document
     * ```typescript
     * await ast.to('html', {
     *     metadataOverrides: { title: 'Q4 Report', author: 'Acme Inc', custom: { department: 'Finance' } },
     * });
     * ```
     */
    metadataOverrides?: MetadataOverrides;
    /**
     * Whether to ignore the built-in default style mappings (e.g. "Heading 1" -> h1).
     * Set to true if you want full control over style mapping.
     * Defaults to false.
     */
    ignoreDefaultStyleMap?: boolean;
    /**
     * How to render an image node in the generated output. Accepts a boolean for backward
     * compatibility (`true` = `'image-only'`, `false` = `'none'`) or one of the {@link ImageMode}
     * strings:
     * - `'image-only'` (default): embed the image (inlined as a `data:` URI when under
     *   `maxInlineImageBytes`, otherwise referenced by name); no OCR/recognized text. ONE exception:
     *   in Markdown and plain-text output, an image OVER `maxInlineImageBytes` that cannot be embedded
     *   falls back to its recognized (OCR) text when it has any, since a bare placeholder would lose a
     *   scanned page's whole content (see `maxInlineImageBytes`). Use `'none'` to guarantee no image text.
     * - `'image+ocr-text'`: embed the image, then its recognized (OCR) text below it.
     * - `'ocr-text-only'`: only the recognized (OCR) text, no image.
     * - `'none'`: omit the image entirely.
     *
     * Plain-text output cannot embed an image, so there `'image-only'`/`'image+ocr-text'` render an
     * `[Image: name]` placeholder (plus the OCR text for `'image+ocr-text'`), `'ocr-text-only'` renders
     * just the OCR text, and `'none'` renders nothing.
     *
     * Defaults to `true` (`'image-only'`).
     */
    includeImages?: boolean | ImageMode;
    /**
     * Maximum size, in bytes of decoded image data, of an image that HTML/Markdown will inline as a
     * `data:` URI.
     *
     * Over the cap under the default `'image-only'` mode, Markdown and plain text render the image's
     * recognized (OCR) text when it has any (multi-line text as a fenced block in Markdown), and
     * otherwise a compact reference to the attachment name; Markdown also emits the
     * `IMAGE_NOT_INLINED` warning so a caller can ship the file alongside the output. Fragment HTML
     * instead keeps `<img src="name">`, a reference an HTML consumer can resolve, and never falls
     * back to text. Standalone HTML always inlines, whatever this value is: a standalone document
     * has nowhere else to resolve the image from. Plain text never inlines at all, so for it the cap
     * only decides whether a large image contributes its recognized text (over the cap) or the
     * `[Image: name]` placeholder (under it).
     *
     * This guards against pathologically large single lines. A scanned PDF page, for instance, is
     * one big image; inlined as a multi-megabyte `data:` URI it can overflow downstream Markdown
     * parsers. Set to `0` to disable inlining entirely (always reference), or `Infinity` to always
     * inline regardless of size.
     *
     * Defaults to 1500000 (about 1.5 MB of image data). Note the `data:` URI itself is roughly a
     * third larger, since base64 encodes 3 bytes as 4 characters.
     */
    maxInlineImageBytes?: number;
    /**
     * Whether to include charts in the generated output. Defaults to true.
     * HTML renders an interactive Chart.js canvas; DOCX and ODT render the chart's data as a table;
     * plain text and the native PDF engine render the chart's data text (its series, which live in
     * `node.text`); Markdown and RTF render nothing for a chart. `false` omits charts in every generator.
     */
    includeCharts?: boolean;
    /**
     * Whether to ignore all internal (anchor) links and anchor IDs during generation.
     * When true, all bookmarks, cross-references, and internal document jumps are stripped.
     * Specifically for Markdown, this removes the {#id} block from headings.
     * Applies to HTML, Markdown, DOCX, ODT and RTF. Defaults to false.
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
    D extends 'docx' ? { docxConfig?: DocxGeneratorConfig } :
    D extends 'odt' ? { odtConfig?: OdtGeneratorConfig } :
    D extends 'tex' ? { texConfig?: TexGeneratorConfig } :
    D extends 'chunks' ? { chunksConfig?: ChunkingConfig } :
    Partial<{
        htmlConfig: HtmlGeneratorConfig;
        mdConfig: MdGeneratorConfig;
        pdfConfig: PdfGeneratorConfig;
        csvConfig: CsvGeneratorConfig;
        textConfig: TextGeneratorConfig;
        rtfConfig: RtfGeneratorConfig;
        docxConfig: DocxGeneratorConfig;
        odtConfig: OdtGeneratorConfig;
        texConfig: TexGeneratorConfig;
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
export type FullGeneratorConfig = DeepRequired<Omit<CommonGeneratorConfig, 'metadataOverrides'> & {
    htmlConfig: HtmlGeneratorConfig;
    mdConfig: MdGeneratorConfig;
    pdfConfig: PdfGeneratorConfig;
    csvConfig: CsvGeneratorConfig;
    textConfig: TextGeneratorConfig;
    rtfConfig: RtfGeneratorConfig;
    docxConfig: DocxGeneratorConfig;
    odtConfig: OdtGeneratorConfig;
    texConfig: TexGeneratorConfig;
}> & {
    chunksConfig: ChunkingConfig;
    // Deliberately not DeepRequired: every field is meant to stay optional, since the whole
    // point is overriding individual fields without having to supply the rest.
    metadataOverrides: MetadataOverrides;
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
     * - `'full'`: the complete premium stylesheet using global selectors (`body`, `h1`, `table`, …).
     *   This is what `standalone: true` has always emitted.
     * - `'scoped'`: the same styling, scoped under the fragment's container via CSS `@scope` so it
     *   cannot leak into a host page's own styles. Requires a modern browser engine (Chrome 118+,
     *   Safari 17.4+, Firefox 128+).
     * - `'none'`: no stylesheet is emitted; the host page (or EPUB reader, or rich-text editor)
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
    /**
     * Carry each rich node's raw source in a `data-*` attribute, with undelimited text content,
     * so attribute-driven structured consumers (rich-text editors, custom viewers) can rehydrate
     * the node from the markup rather than re-parsing the display text. Affects wikilinks
     * (adds `data-wikilink`/`data-target`/`data-alias`), citations (a `<span class="citation">`
     * carrying `data-key` instead of `<cite>`), math (the LaTeX in `data-math`, undelimited),
     * mermaid (a `<div class="mermaid" data-mermaid>` instead of `<pre><code>`) and source comments
     * (an empty `<span data-html-comment="…">` instead of `<!-- … -->`, since an editor's DOM parser
     * discards real comment nodes).
     *
     * Off by default; the default output is byte-identical to previous releases. The widened
     * `HtmlParser` reads every shape this emits, so output stays self-round-trippable.
     */
    sourceAttributes?: boolean;
    /**
     * Emit a generic (non-YouTube) iframe embed as a gated placeholder,
     * `<div data-embed-gated data-embed-src="…" …>`, instead of a live `<iframe>`. The gated shape
     * never auto-loads its src: an editor renders a click-to-load placeholder from it, and
     * `HtmlParser` reads it back to the same `embed` node. The src is scheme-checked (`sanitizeUrl`)
     * on emit. Off by default; the default output (a live `<iframe>`) is unchanged. YouTube embeds
     * are unaffected (they already render from a validated id).
     */
    gatedEmbeds?: boolean;
    /**
     * Omit an inline run `color` that is effectively the document default (near-black or near-white),
     * so imported text inherits the reader's theme instead of being pinned to black (unreadable on a
     * dark background) or white (unreadable on a light one). Only near-black/near-white run colours are
     * dropped; every deliberately-coloured run is emitted unchanged. Word's own "automatic" colour
     * (`w:val="auto"`) already carries no colour, so this targets documents that hard-code an explicit
     * default-ish colour. Off by default; the default output is unchanged.
     */
    omitDefaultTextColor?: boolean;
}

/**
 * Named paper sizes shared by every generator that lays out pages (PDF, DOCX, ODT, LaTeX), so a page size
 * is written the same way whatever the destination. Case-insensitive: `'A4'` and `'a4'` are the same
 * size. The A-series is ISO 216; Letter/Legal/Tabloid/Ledger are the US/ANSI sizes.
 */
export type PaperFormat =
    'letter' | 'legal' | 'tabloid' | 'ledger' | 'a0' | 'a1' | 'a2' | 'a3' | 'a4' | 'a5' | 'a6' |
    'Letter' | 'Legal' | 'Tabloid' | 'Ledger' | 'A0' | 'A1' | 'A2' | 'A3' | 'A4' | 'A5' | 'A6';

/**
 * Configuration options for PDF generation.
 * Maps closely to Puppeteer's PDF options.
 */
export interface PdfGeneratorConfig {
    /**
     * Which engine renders the PDF:
     * - `'html'` (default): render the document to the library's HTML and rasterize/print it through
     *   a headless browser (Puppeteer in Node; the browser's own print in a browser). Highest visual
     *   fidelity, but needs the optional `puppeteer` peer dependency in Node, and in a browser it can
     *   only hand back HTML rather than PDF bytes.
     * - `'native'`: lay the AST out directly with `pdf-lib` (an optional peer dependency) - no
     *   browser required, works identically in Node and the browser (so it can produce real PDF bytes
     *   client-side), and is much lighter. The trade-off is fidelity: it draws in the Standard-14
     *   Helvetica and Courier fonts rather than the document's own and performs a clean reflow rather
     *   than matching the HTML renderer pixel for pixel. `tagged`/`outline` do not apply to it.
     *
     * Defaults to `'html'`.
     */
    engine?: 'html' | 'native';
    /**
     * Emit a tagged (accessible / PDF-UA-friendly) PDF, so screen readers can follow the document
     * structure. Passed through to the rendering engine; ignored by engines that do not support it.
     * Defaults to true.
     */
    tagged?: boolean;
    /**
     * Emit a bookmark outline built from the document's headings (experimental in the rendering
     * engine). Defaults to false.
     */
    outline?: boolean;
    /** Paper format. Defaults to 'A4'. See {@link PaperFormat}. */
    format?: PaperFormat;
    /** Paper width, accepts values labeled with units (e.g., '5in', '3cm') or numbers (in pixels). */
    width?: string | number;
    /** Paper height, accepts values labeled with units (e.g., '5in', '3cm') or numbers (in pixels). */
    height?: string | number;
    /** Whether to print in landscape orientation. Defaults to false. */
    landscape?: boolean;
    /** Whether to print background graphics. Defaults to true. HTML engine only. */
    printBackground?: boolean;
    /** Scale of the webpage rendering. Defaults to 1. HTML engine only. */
    scale?: number;
    /**
     * Paper margins. The default depends on the engine: the `'html'` engine uses 0 on every side (the
     * body carries its own padding), while the `'native'` engine uses a small default (~48pt) so text
     * is not glued to the sheet edge. An explicit value (including `0`) is honored by both engines.
     */
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
 * Configuration options for DOCX (Word) generation.
 */
export interface DocxGeneratorConfig {
    /**
     * Page size for the document section (`w:pgSz`). One of the shared {@link PaperFormat} names,
     * the same set and spelling {@link PdfGeneratorConfig.format} accepts. Defaults to `'A4'`.
     */
    format?: PaperFormat;
    /** Landscape orientation (swaps the page dimensions and sets `w:orient`). Defaults to false. */
    landscape?: boolean;
    /**
     * Page margins, converted to twips for `w:pgMar`. Each side is a number of points (1/72 inch) or
     * a unit-labeled string (`'1in'`, `'2cm'`, `'36pt'`, `'48px'`); a bare number is points. Defaults
     * to 72 (Word's standard one inch) on every side.
     */
    margin?: { top?: number | string; right?: number | string; bottom?: number | string; left?: number | string };
}

/**
 * Configuration options for ODT (OpenDocument Text) generation.
 */
export interface OdtGeneratorConfig {
    /**
     * Page size for the page layout (`style:page-layout` in styles.xml). One of the shared
     * {@link PaperFormat} names, the same set {@link PdfGeneratorConfig.format} accepts. Defaults to
     * `'A4'`.
     */
    format?: PaperFormat;
    /** Landscape orientation (swaps page dimensions and sets `style:print-orientation`). Defaults to false. */
    landscape?: boolean;
    /**
     * Page margins, written as `fo:margin-*` lengths. Each side is a number of points (1/72 inch) or
     * a unit-labeled string (`'1in'`, `'2cm'`, `'36pt'`, `'48px'`); a bare number is points. Defaults
     * to 72 (the standard one inch) on every side.
     */
    margin?: { top?: number | string; right?: number | string; bottom?: number | string; left?: number | string };
}

/**
 * The LaTeX document class {@link TexGeneratorConfig.documentClass} selects. `'auto'` picks `beamer`
 * for a presentation (content made of slides) and `article` for everything else.
 */
export type TexDocumentClass = 'auto' | 'article' | 'report' | 'book' | 'beamer';

/**
 * Configuration options for LaTeX (`.tex`) generation.
 *
 * The output compiles with pdfLaTeX, XeLaTeX, LuaLaTeX, upLaTeX, pLaTeX and `latex` (the last three
 * through `dvipdfmx`), TeX Live 2021 or later: the preamble loads `fontspec` under XeTeX and LuaTeX
 * (with TeX Live's fonts for Greek, Cyrillic and CJK text where the document has it) and
 * `fontenc`/`inputenc` under pdfTeX and the pTeX family, gives every package the `dvipdfmx` driver
 * in DVI mode, and loads only the packages the document actually uses.
 */
export interface TexGeneratorConfig {
    /**
     * The document class. `'auto'` (default) writes a `beamer` presentation when the content is made
     * of slides (PPTX/ODP) and an `article` otherwise. `'report'` and `'book'` map level-1 headings to
     * `\chapter`; `'beamer'` turns each slide into a frame, or, for a non-presentation source, starts
     * a new frame at every level-1/2 heading.
     */
    documentClass?: TexDocumentClass;
    /**
     * Whether to emit a complete document (preamble, `\begin{document}`...`\end{document}`). When
     * false, only the body is emitted, headed by a comment listing the packages it needs, for pasting
     * into an existing document. Defaults to true.
     */
    standalone?: boolean;
    /**
     * When true, the result is a zip (`Uint8Array`) holding `main.tex` plus every image as a file under
     * `images/`, ready to compile or upload to Overleaf. When false (default) the result is the `.tex`
     * source as a string, carrying its PNG and JPEG images inside it (see `embedImages`); any other
     * image is referenced as `images/<name>`, and the `IMAGES_NOT_BUNDLED` warning names the files the
     * caller must place there (their bytes are in `ast.attachments`).
     */
    bundle?: boolean;
    /**
     * Whether a `.tex` (without `bundle`) carries its PNG and JPEG images inside it, so the one file
     * compiles with its pictures. LaTeX reads images only from files, so each image is written into
     * the `.tex` as a `filecontents*` block holding it as a small all-ASCII PDF; compiling writes that
     * file beside the `.tex` (keeping a file of that name already there) and `\includegraphics` reads
     * it. Every engine reads it; with `--output-directory`, pdfLaTeX and LuaLaTeX still find the
     * files, but XeLaTeX and dvipdfmx look beside the `.tex`, so use `bundle` there. The data adds
     * about a quarter to each image's size. A PNG that must be decoded to be carried (transparency,
     * interlacing) is carried up to 16 megapixels, and a document's decoded images up to 256 in all;
     * past that it is referenced as a file. When false, images are referenced as `images/<name>`
     * files and reported with `IMAGES_NOT_BUNDLED`. Defaults to true. Ignored with `bundle`.
     */
    embedImages?: boolean;
    /** Number sections (`1`, `1.1`, ...). Defaults to false, matching office documents, whose headings are unnumbered by default. */
    numberSections?: boolean;
    /**
     * Paper size, written as a `geometry` option. One of the shared {@link PaperFormat} names.
     * Defaults to `'A4'`. Ignored by `beamer`, whose frame size is fixed.
     */
    format?: PaperFormat;
    /** Landscape orientation. Defaults to false. Ignored by `beamer`. */
    landscape?: boolean;
    /**
     * Page margins, written as `geometry` options. Each side is a number of points (1/72 inch) or a
     * unit-labeled string (`'1in'`, `'2cm'`, `'36pt'`, `'48px'`); a bare number is points. Defaults to
     * 72 (one inch) on every side. Ignored by `beamer`.
     */
    margin?: { top?: number | string; right?: number | string; bottom?: number | string; left?: number | string };
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
 * Named Markdown dialect presets for `MarkdownDialectConfig`/`MdGeneratorConfig.dialect`.
 * `'extended'` is officeParser's own kitchen-sink default and reproduces this library's
 * historical output exactly (every feature on, GitHub-style admonitions).
 */
export type MarkdownDialectPreset = 'extended' | 'github' | 'gitlab' | 'obsidian' | 'pandoc' | 'commonmark';

/*
 * Per-capability syntax variants for `MarkdownDialectConfig`. Each is named for the *syntax* it
 * selects, never for a product, so a convention shared by several flavors is a single value and a
 * preset simply points at it (e.g. both the `obsidian` and `extended` presets select `'equals'`
 * highlight). `'none'` is the explicit off switch, mirroring `math`'s existing `'dollar' | 'none'`;
 * `undefined`/omitted means "inherit from the `extends` preset", never "off". Modelling these as
 * unions rather than booleans lets a second syntax be added later without a breaking change.
 */

/** Admonition syntax: `'blockquote'` = GitHub `> [!NOTE]`, `'fence'` = GitLab `:::note`,
 *  `'fence-attribute'` = Pandoc `::: {.note}`, `'none'` = plain bold-labeled blockquote. */
export type AdmonitionSyntax = 'blockquote' | 'fence' | 'fence-attribute' | 'none';
/** `==text==` highlight (`'equals'`), or `'none'` to disable. */
export type HighlightSyntax = 'equals' | 'none';
/** GFM `~~text~~` strikethrough (`'tilde'`), or `'none'`. */
export type StrikethroughSyntax = 'tilde' | 'none';
/** `Term`/`: Description` definition lists (`'colon'`), or `'none'`. */
export type DefinitionListSyntax = 'colon' | 'none';
/** `[^id]` footnotes (`'caret'`), or `'none'`. */
export type FootnoteSyntax = 'caret' | 'none';
/** `[@citekey]` citations (`'at'`), or `'none'`. */
export type CitationSyntax = 'at' | 'none';
/** `[[Page]]` wikilinks (`'double-bracket'`), or `'none'`. */
export type WikilinkSyntax = 'double-bracket' | 'none';
/** `{width=50%}` attribute lists (`'brace'`), or `'none'`. */
export type AttributeListSyntax = 'brace' | 'none';
/**
 * How an `embed` node is written to Markdown:
 * - `'html'` (default): the single-line `<div data-youtube-video="ID">` / `<iframe src=...>` block
 *   this library has always emitted. Round-trips through officeParser, but renders as an invisible
 *   empty box on GitHub.
 * - `'directive'`: a remark-directive leaf, `::youtube[Label]{id=... width=... align=...}` /
 *   `::embed[Label]{src=... width=... height=... align=...}`. Round-trips within an editor that
 *   understands it; renders verbatim (not just the label) on GitHub, so it is an editor format, not
 *   a GitHub-interop one.
 * - `'link'`: a plain `[YouTube](url)` / `[Embed](url)`.
 * - `'thumbnail'`: a YouTube-only clickable thumbnail `[![Label](.../vi/ID/hqdefault.jpg)](watch)`,
 *   the best GitHub degrade; a non-YouTube embed falls back to `'link'`.
 */
export type EmbedSyntax = 'html' | 'directive' | 'link' | 'thumbnail';

/**
 * @deprecated Legacy flavor names for `MarkdownDialectConfig.admonitions`. Use the syntax names
 * instead: `'github'` -> `'blockquote'`, `'gitlab'` -> `'fence'`, `'pandoc'` -> `'fence-attribute'`.
 * These aliases still resolve to the same output and will be removed in the next major version.
 */
export type DeprecatedAdmonitionFlavor = 'github' | 'gitlab' | 'pandoc';

/**
 * @deprecated Boolean toggles for dialect capability fields are deprecated in favor of the
 * syntax-name unions: `true` maps to that field's on-value (e.g. `'tilde'`), `false` maps to
 * `'none'`. Booleans keep working via coercion and will be removed in the next major version.
 */
export type DeprecatedDialectToggle = boolean;

/**
 * Granular control over which native Markdown syntax the generator emits for constructs that
 * differ across real-world dialects (e.g. GitHub's `> [!NOTE]` vs GitLab's `:::note` vs Pandoc's
 * `::: {.note}` admonitions). Shorthand: pass a `MarkdownDialectPreset` string for a named target;
 * pass an object for granular control. Any field you omit from the object form falls back to the
 * preset named by `extends` (default `'extended'`) - **not** to whatever preset may have been
 * ambient before, since config merging replaces the whole field rather than layering on top of it.
 */
export interface MarkdownDialectConfig {
    /** Base preset any omitted field inherits from. Defaults to 'extended'. */
    extends?: MarkdownDialectPreset;
    /**
     * Admonition syntax: `'blockquote'` = GitHub `> [!NOTE]`, `'fence'` = GitLab `:::note`,
     * `'fence-attribute'` = Pandoc `::: {.note}`, `'none'` = a plain bold-labeled blockquote with no
     * special marker. Omit to inherit from the `extends` preset. The legacy flavor names
     * `'github'`/`'gitlab'`/`'pandoc'` are accepted as deprecated aliases (see
     * `DeprecatedAdmonitionFlavor`) and will be removed in the next major version.
     */
    admonitions?: AdmonitionSyntax | DeprecatedAdmonitionFlavor;
    /**
     * Markdown Extra/Pandoc-style `Term`/`: Description` definition lists (`'colon'`), or `'none'`
     * to render terms and descriptions as plain paragraphs. Omit to inherit from `extends`. Passing
     * a boolean is deprecated: `true` = `'colon'`, `false` = `'none'` (removed next major).
     */
    definitionLists?: DefinitionListSyntax | DeprecatedDialectToggle;
    /**
     * `[^id]` footnote references/definitions (`'caret'`), or `'none'` to inline note content as a
     * parenthetical right at the reference point. Omit to inherit from `extends`. Passing a boolean
     * is deprecated: `true` = `'caret'`, `false` = `'none'` (removed next major).
     */
    footnotes?: FootnoteSyntax | DeprecatedDialectToggle;
    /**
     * Pandoc-style `[@citekey]` citations (`'at'`), or `'none'` to emit `[citekey]` (brackets, no
     * `@`). Omit to inherit from `extends`. Passing a boolean is deprecated: `true` = `'at'`,
     * `false` = `'none'` (removed next major).
     */
    citations?: CitationSyntax | DeprecatedDialectToggle;
    /**
     * Obsidian-style `[[Page]]`/`[[Page|Alias]]` wikilinks (`'double-bracket'`), or `'none'` to fall
     * back to a plain `[text](url)` link using the same target. Omit to inherit from `extends`.
     * Passing a boolean is deprecated: `true` = `'double-bracket'`, `false` = `'none'` (removed next major).
     */
    wikilinks?: WikilinkSyntax | DeprecatedDialectToggle;
    /** Inline `$...$`/block `$$...$$` math delimiters (`'dollar'`), or `'none'` for bare LaTeX text. */
    math?: 'dollar' | 'none';
    /**
     * Pandoc-style `{width=50% .centered}` attribute lists after images/tables (`'brace'`), or
     * `'none'`. Omit to inherit from `extends`. Passing a boolean is deprecated: `true` = `'brace'`,
     * `false` = `'none'` (removed next major).
     */
    attributeLists?: AttributeListSyntax | DeprecatedDialectToggle;
    /**
     * GFM `~~text~~` strikethrough (`'tilde'`; not part of base CommonMark), or `'none'`. Omit to
     * inherit from `extends`. Passing a boolean is deprecated: `true` = `'tilde'`, `false` = `'none'`
     * (removed next major).
     */
    strikethrough?: StrikethroughSyntax | DeprecatedDialectToggle;
    /**
     * `==text==` highlight (`'equals'`; Obsidian/extended flavors, NOT GFM or CommonMark where `==`
     * is literal text), or `'none'`. When `'equals'`, a highlighted run round-trips as `==text==`
     * and `==text==` is read back as a highlight; when `'none'`, a highlight falls back to an HTML
     * `<mark>`/`<span>` per `fallbackToHtml.inlineFormatting`, and `==text==` stays literal on parse.
     * Omit to inherit from `extends`.
     */
    highlight?: HighlightSyntax;
    /** Unordered list bullet character. */
    bulletListMarker?: '-' | '*' | '+';
    /** Ordered list marker punctuation. */
    orderedListMarker?: '.' | ')';
    /** Emphasis delimiter style for bold/italic. */
    emphasisMarker?: 'asterisk' | 'underscore';
    /** Table syntax: native GFM pipe tables, or forced HTML `<table>` (required for strict
     *  CommonMark, which has no table syntax of its own). */
    tables?: 'native' | 'html';
    /**
     * How an `embed` node is written to Markdown (`'html'` | `'directive'` | `'link'` |
     * `'thumbnail'`; see `EmbedSyntax`). This is the authority for embed form. When omitted, the
     * deprecated `fallbackToHtml.embeds` boolean is honored (`true`/unset maps to `'html'`, `false`
     * to `'link'`), then the default `'html'`.
     */
    embeds?: EmbedSyntax;
}

/**
 * Granular control over when the Markdown generator falls back to raw HTML tags for features
 * standard Markdown can't express natively. Shorthand: `true`/`false` (via
 * `MdGeneratorConfig.fallbackToHtml`) turns every part on/off at once; pass an object instead to
 * control them independently. Omitted object fields default to on, matching the boolean shorthand.
 */
export interface FallbackToHtmlConfig {
    /** Underline/subscript/superscript via `<u>`/`<sub>`/`<sup>`. */
    textFormatting?: boolean;
    /** Heading/paragraph text alignment via `<div style="text-align:...">`. */
    alignment?: boolean;
    /** Internal-link/heading `<a id>`/`<a name>` anchor tags. */
    anchors?: boolean;
    /** Nested-table and merged-cell (colspan/rowspan) HTML `<table>` fallback. */
    tables?: boolean;
    /**
     * YouTube embed `<div data-youtube-video>` vs. a plain link.
     * @deprecated Use `mdConfig.dialect.embeds` (`EmbedSyntax`) instead, which also selects the
     * `'directive'` and `'thumbnail'` forms. When `dialect.embeds` is unset this boolean is still
     * honored (`true` maps to `'html'`, `false` to `'link'`); it will be removed in the next major.
     */
    embeds?: boolean;
    /** Multi-line table cell content joined with `<br>` instead of a space. */
    cellLineBreaks?: boolean;
    /**
     * Multi-paragraph list-item content (an HTML `<li>` with several `<p>` children) joined with
     * `<br>` instead of a space, so it stays on the item's single Markdown line. Block children of
     * an item (a code fence or table inside `<li>`) degrade under this join, the same way they do
     * inside a table cell under `cellLineBreaks`.
     */
    itemLineBreaks?: boolean;
    /**
     * Inline text color, highlight, and font size via a `<span style="color:...;background-color:...;
     * font-size:...">` run, which the Markdown parser reads back. These have no Markdown syntax and
     * are silently lost otherwise. Unlike the other fields this is **off by default even when
     * `fallbackToHtml` is `true`**, because it changes default output; enable it explicitly with
     * `fallbackToHtml: { inlineFormatting: true }`.
     */
    inlineFormatting?: boolean;
}

/**
 * Configuration options for Markdown generation.
 */
export interface MdGeneratorConfig {
    /**
     * Whether to fallback to HTML tags for features not supported by standard Markdown.
     * Pass an object instead of a boolean for granular control over individual parts (text
     * formatting, alignment, anchors, tables, embeds, cell line breaks) - see
     * `FallbackToHtmlConfig`. Omitted object fields default to on, matching `true`.
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
    fallbackToHtml?: boolean | FallbackToHtmlConfig;

    /**
     * Target Markdown dialect for generation - which native syntax to emit for constructs that
     * differ across real-world targets (GitHub/GitLab/Obsidian/Pandoc/strict CommonMark). See
     * `MarkdownDialectConfig` for the full per-feature field list. Defaults to `'extended'`
     * (officeParser's own historical kitchen-sink behavior, unchanged from prior versions).
     */
    dialect?: MarkdownDialectPreset | MarkdownDialectConfig;
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
     * If true, tables are rendered with separators and aligned columns, and list items get their
     * markers and indentation.
     * If false, output is a flat stream of text nodes (cells are tab-separated).
     * Defaults to **true**.
     */
    preserveLayout?: boolean;
    /**
     * Whether to append the collected footnotes/endnotes as a trailing `--- Notes ---` section.
     * Set false to omit it when you want only the document body; the notes are still parsed and
     * remain available on the AST, they are simply not rendered into the text output.
     *
     * Note this differs from the parser's `ignoreNotes`, which discards notes at parse time so they
     * never reach the AST at all. Use this when you want the AST to keep them but the text output
     * to leave them out.
     *
     * This generation-time suppression is specific to plain-text output. Other generators have no
     * equivalent switch (Markdown can inline them via `dialect.footnotes: 'none'`; HTML/DOCX/ODT/PDF
     * always render collected notes); use the parser's `ignoreNotes`, or `onNode`, to drop them there.
     *
     * Defaults to true.
     */
    renderNotes?: boolean;
    /**
     * String inserted between top-level page nodes (PDF) in the rendered text. Set to '\f' for a
     * form feed between pages (matching `pdftotext`), or a custom banner. Applies in both the
     * layout-faithful and flowing text modes.
     *
     * Defaults to '\n' (a single blank line between pages).
     */
    pageSeparator?: string;
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

/** A single value substituted for a template placeholder. `null`/`undefined` render as empty text. */
export type TemplateValue = string | number | boolean | Date | null | undefined;

/**
 * A flat map of placeholder name to value for {@link OfficeTemplate.render}. A key `name` fills the
 * placeholder `{{name}}` (see `delimiters`). Keys absent from the map are governed by `onMissing`.
 */
export type TemplateData = Record<string, TemplateValue>;

/**
 * Options for {@link OfficeTemplate.render}. `data` is the only required field: a single
 * {@link TemplateData} produces one document, an array of them produces one document per entry
 * (a mail-merge).
 */
export interface TemplateConfig {
    /** The field values. One object -> one rendered document; an array -> one document per entry. */
    data: TemplateData | TemplateData[];
    /**
     * Placeholder delimiters. Defaults to `{{` and `}}`, so `{{name}}` is a placeholder. A placeholder
     * name is matched even when the source splits it across several runs (a common Word quirk).
     */
    delimiters?: { start: string; end: string };
    /**
     * What to do with a placeholder whose name is absent from `data`:
     * - `'keep'` (default): leave the placeholder text untouched.
     * - `'empty'`: replace it with nothing.
     * - `'error'`: reject with `TEMPLATE_FIELD_MISSING`.
     * A key that is present but `null`/`undefined` always renders as empty (it is not "missing").
     */
    onMissing?: 'keep' | 'empty' | 'error';
    /**
     * Password, if the template document is itself encrypted; it is decrypted before rendering. Note
     * the rendered output is the plain (unencrypted) document.
     */
    password?: string;
    /**
     * Callback invoked when the template is encrypted and `password` did not decrypt it, mirroring the
     * parser's `onPassword`. Called with `'required'` when the template is encrypted and no password
     * was given, or `'incorrect'` when the last attempt was wrong. Return a password (sync or async) to
     * retry, or `undefined` to give up (rejecting with `PASSWORD_REQUIRED`/`PASSWORD_INCORRECT`).
     * The callback is asked at most 3 times in total, so one that keeps returning a wrong password
     * cannot loop forever.
     */
    onPassword?: (reason: 'required' | 'incorrect') => string | undefined | Promise<string | undefined>;
    /**
     * Optional hint for the template's format. DOCX is the only supported template format today, and the
     * format is always detected from the bytes, so this hint does not influence detection: it only
     * refines the `TEMPLATE_UNSUPPORTED_FORMAT` message when an unsupported file is passed. It is kept as
     * the extension point for a second template format.
     */
    fileType?: 'docx';
    /**
     * Limits on decompressing the (untrusted) template zip, same shape and defaults as the parser's.
     * Guards against zip-bomb templates. Defaults to 512 MiB / 10000 entries.
     */
    decompressionLimits?: DecompressionLimits;
    /**
     * Optional callback for non-fatal issues. Also suppresses the console fallback that would
     * otherwise print a thrown error, so a caller handling the rejection is not double-notified.
     */
    onWarning?: (issue: OfficeIssue) => void;
}

/**
 * Supported file types for parsing.
 */
export type SupportedFileType = 'docx' | 'pptx' | 'xlsx' | 'odt' | 'odp' | 'ods' | 'odg' | 'pdf' | 'rtf' | 'md' | 'html' | 'csv' | 'epub' | 'tex';

/**
 * Other names the `fileType` option accepts, as the matching file extensions do: `latex` and `ltx`
 * for `tex`; the ODF template packages `ott`, `ots`, `otp` and `otg` for `odt`, `ods`, `odp` and
 * `odg`; and `zip`, which is parsed as whatever the archive holds (a LaTeX project, an office
 * document).
 */
export type FileTypeAlias = 'latex' | 'ltx' | 'ott' | 'ots' | 'otp' | 'otg' | 'zip';

/**
 * A structural stand-in for the web `Blob`/`File` so `parseOffice`/`convert` accept them in the
 * browser without pulling the DOM lib into this package's types. Any object with an
 * `arrayBuffer()` method qualifies. When `name` is present (as on a `File`) it is used only for
 * extension-based type detection, never as a filesystem path.
 */
export interface BlobLike {
    arrayBuffer(): Promise<ArrayBuffer>;
    name?: string;
}

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
     * Text alignment for this cell's column, from the GFM pipe-table separator row
     * (`:---` left, `:---:` center, `---:` right). All cells in a column carry the same value;
     * the Markdown generator reads it from the header row to emit the separator.
     */
    align?: 'left' | 'center' | 'right';
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
     * Layout alignment of the table on the page (e.g. an editor's custom table node).
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
     * Display width of the image (e.g. an editor's custom image node), as a CSS length or percentage.
     * @example "50%"
     */
    width?: string;
    /**
     * Layout alignment of the image (e.g. an editor's custom image node).
     * @example 'center'
     */
    align?: 'left' | 'center' | 'right';
    /** Advisory image title (Markdown `![alt](url "title")`, HTML `<img title>`), if any. */
    title?: string;
}

/**
 * Metadata for an embedded external media node (e.g. a YouTube video).
 * Markdown has no native syntax for this - see `MarkdownGenerator`'s `embed` case.
 */
export interface EmbedMetadata {
    /**
     * The kind of embed. 'youtube' is recognized from a `data-youtube-video` wrapper or a YouTube
     * iframe; 'iframe' is a generic preserved iframe (opt-in via `HtmlParserConfig.preserveIframes`).
     */
    embedType: 'youtube' | 'iframe';
    /** The provider-specific video ID (e.g. the 11-character YouTube video ID). Absent for generic iframes. */
    videoId?: string;
    /** The original/canonical URL of the embedded media, if known. For a generic iframe, its `src`. */
    url?: string;
    /** Display width, as a CSS length or percentage. */
    width?: string;
    /** Display height, as a CSS length or percentage. */
    height?: string;
    /** Layout alignment of the embed. */
    align?: 'left' | 'center' | 'right';
    /** Human-readable label for the embed (e.g. the `[Label]` of a `::youtube[Label]{...}` leaf
     *  directive, or a gated embed's caption). Purely descriptive; never a trust or render input. */
    label?: string;
}

/**
 * Metadata for an admonition/alert node (e.g. GitHub's `> [!NOTE]` or GLFM's `:::note`).
 * `MarkdownParser` accepts both syntaxes (and generates either, plus Pandoc's `::: {.note}`,
 * depending on `MdGeneratorConfig.dialect`). Children are block content (paragraphs) wrapped by
 * the admonition.
 */
export interface AdmonitionMetadata {
    admonitionType: 'note' | 'tip' | 'important' | 'warning' | 'caution';
    /** Optional custom title; falls back to the type label. */
    title?: string;
    /** Which concrete input syntax produced this node. Always populated by the parser. */
    sourceSyntax?: 'github' | 'gitlab';
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
    /**
     * Page width in PDF points (1/72 inch), after applying the page's own rotation. Matches the
     * coordinate space of child {@link NodeBounds}. Absent when `ignorePageGeometry` is set.
     * @example 612 for US Letter portrait
     */
    pageWidth?: number;
    /**
     * Page height in PDF points (1/72 inch), after applying the page's own rotation.
     * Absent when `ignorePageGeometry` is set.
     * @example 792 for US Letter portrait
     */
    pageHeight?: number;
    /**
     * The page's clockwise rotation in degrees (`/Rotate`), one of 0/90/180/270. Emitted only when
     * non-zero. The reported `pageWidth`/`pageHeight` and all child bounds are already in this
     * rotated space.
     */
    rotation?: number;
    /**
     * The printed page label from the PDF's `/PageLabels` tree (e.g. `"iv"`, `"A-1"`), when it
     * exists and differs from the plain 1-based `pageNumber`. Front matter numbered in roman
     * numerals is the common case. Absent when the document has no page labels.
     * @example "iv" for the fourth page of front matter
     */
    pageLabel?: string;
    /**
     * The drawing page's name from ODG's `draw:name` (e.g. "Slide 1", "Flowchart", or a
     * user-assigned title). ODG only; absent for PDF pages. The PDF-specific fields above
     * (`pageWidth`/`pageHeight`/`rotation`/`pageLabel`) are not populated for ODG in v1.
     * @example "Flowchart"
     */
    pageName?: string;
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
    /** Advisory link title (Markdown `[text](url "title")`, HTML `<a title>`), if any. */
    title?: string;
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
    /**
     * True for a footnote/endnote definition that no reference points at (an "orphan").
     * The Markdown parser sets this when it recovers a `[^id]: ...` definition with no matching
     * `[^id]` reference so the definition is preserved rather than dropped. Generators route such
     * notes into their footnotes section without a citation marker, and the HTML generator omits
     * the (otherwise dangling) back-link.
     */
    unreferenced?: boolean;
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
     * - 'thematic': A thematic break (Markdown `---`/`***`/`___`, HTML `<hr>`) - a horizontal
     *   rule separating sections, distinct from a page break. Emitted as `---` in Markdown and
     *   `<hr>` in HTML.
     */
    breakType: 'column' | 'page' | 'lastRenderedPage' | 'textWrapping' | 'carriageReturn' | 'thematic';

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
     * 'block' as `$$...$$`. Matches attribute-driven editors' math nodes.
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
    /**
     * `'html'` marks a SOURCE comment, `<!-- ... -->` in Markdown or HTML: the author's hidden note,
     * not a review annotation. Its node's `text` is the raw text between `<!--` and `-->`, verbatim
     * (whitespace included) and it has no children. The Markdown and HTML generators re-emit it as a
     * comment (the HTML generator as a `data-html-comment` element under
     * `HtmlGeneratorConfig.sourceAttributes`), and the LaTeX generator as `% <!--...-->` lines, which the
     * LaTeX parser reads back; every other output format omits it, since none has a hidden-comment
     * construct and rendering it would reveal a note the author hid. Absent for review comments
     * (DOCX/PPTX/XLSX/ODF, LaTeX `% Comment:` lines), whose rendering is unchanged.
     */
    sourceSyntax?: 'html';
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
 * An axis-aligned bounding box describing where a node sits on its page.
 *
 * Coordinates are in PDF points (1/72 inch), in post-rotation page space with the origin at the
 * page's top-left corner and y growing downward, i.e. exactly what pdf.js renders at scale 1. This
 * matches the `pageWidth`/`pageHeight` on {@link PageMetadata}. Values are rounded to 2 decimals.
 *
 * Populated by the PDF parser unless `ignorePageGeometry` is set. Container nodes (paragraph, table,
 * row, cell) carry the union of their children's boxes.
 */
export interface NodeBounds {
    /** Distance from the page's left edge to the box's left edge. */
    x: number;
    /** Distance from the page's top edge to the box's top edge. */
    y: number;
    /** Box width. */
    width: number;
    /** Box height. */
    height: number;
}

/**
 * Shared properties available on all document content nodes.
 */
export interface BaseContentNode {
    /**
     * Where this node sits on its page, as an axis-aligned box in page coordinates.
     * See {@link NodeBounds} for the coordinate convention. Present only when the parser knows
     * the geometry (currently PDF) and `ignorePageGeometry` is not set.
     */
    bounds?: NodeBounds;

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

    /**
     * Source HTML attributes that no typed metadata field consumed, preserved for round-trip
     * fidelity (e.g. a `data-*` attribute an editor round-trips through officeParser).
     *
     * Only populated by the HTML/XHTML parser, only for elements it recognises, and only when
     * `htmlParserConfig.preserveAttributes` is enabled - so by default this is always absent.
     *
     * Sanitized on both legs, since an AST can also be constructed programmatically rather than
     * parsed: event handlers (`on*`) and `srcdoc` are never carried, URL-bearing attributes go
     * through the same URL sanitizer as typed fields, and every value is escaped on output. A
     * typed field always wins over a same-named entry here.
     *
     * Ignored by the non-HTML generators (Markdown, RTF, CSV, text, chunking) by design - these
     * are HTML attributes and have no meaning in those targets.
     * @example { 'data-tracking-id': 'abc123', 'class': 'lead' }
     */
    htmlAttributes?: Record<string, string>;
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
     * Uses Tesseract.js with the language specified in `config.ocrConfig.language`.
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
     * The document's primary language as a BCP 47 tag, when the source declares one: the PDF
     * `/Lang` entry (or the text-content language hint), EPUB `dc:language`, or a LaTeX document's
     * `pdflang` or babel/polyglossia main language. Generators write it back where their format has
     * a place for it.
     * @example "en", "en-US", "fr"
     */
    language?: string;
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
    /**
     * The document outline (bookmarks / table of contents), as a tree of `list` nodes. Each item's
     * text is the bookmark title and its `metadata.link` points to the destination: the anchor of the
     * nearest heading on the target page when one is close, else the page anchor `#page=N`, else
     * `#internal`. Nested bookmarks are the item's `children`. Populated for PDFs that declare an
     * outline, unless `ignoreInternalLinks` is set. Absent otherwise.
     */
    outline?: OfficeContentNode[];
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
 * console.log((await ast.to('text')).value); // Plain text representation
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


import { ChunkingConfig, CsvGeneratorConfig, DeepRequired, DocumentStructureChunkingConfig, DocxGeneratorConfig, FixedSizeChunkingConfig, FullGeneratorConfig, HtmlGeneratorConfig, HtmlParserConfig, MdGeneratorConfig, OcrConfig, OcrTimeoutConfig, OdtGeneratorConfig, OfficeParserConfig, PdfGeneratorConfig, PdfParserConfig, SemanticChunkingConfig, TexGeneratorConfig, TexParserConfig, TextGeneratorConfig } from './types.js';

const PDFJS_VERSION = '6.2.108';
const DEFAULT_PDF_WORKER_SRC = typeof __SLIM__ !== 'undefined' && __SLIM__ ? '' : `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}/build/pdf.worker.min.mjs`;

/**
 * The default regex used for identifying sentence boundaries.
 * When this default is used, the generator employs a high-fidelity "robust"
 * segmenter that accounts for common abbreviations (Mr., Dr., etc.).
 */
export const DEFAULT_SENTENCE_BOUNDARY_REGEX = /[.!?。！？]/;

/**
 * Common abbreviations that should not trigger a sentence split when followed by a period.
 */
export const DEFAULT_ABBREVIATIONS = ['Mr', 'Dr', 'Ms', 'Inc', 'Ltd', 'Prof', 'Sr', 'Jr', 'vs', 'etc'];

/** Default timeout values for OCR */
const DEFAULT_OCR_TIMEOUT: Required<OcrTimeoutConfig> = {
    autoTerminate: 10000,
    workerLoad: 60000,
    recognition: 30000,
}

/**
 * Default configuration for OCR.
 */
const DEFAULT_OCR_CONFIG: DeepRequired<OcrConfig> = {
    language: 'eng',
    workerPath: '',
    corePath: '',
    langPath: '',
    preserveLayout: true,
    timeout: DEFAULT_OCR_TIMEOUT,
    abortSignal: null,
};

/**
 * Default configuration for HTML/XHTML parsing. `preserveAttributes` is off so that the AST is
 * byte-identical to previous releases unless a caller opts in - see `HtmlParserConfig`.
 */
const DEFAULT_HTML_PARSER_CONFIG: DeepRequired<HtmlParserConfig> = {
    preserveAttributes: false,
    preserveIframes: false,
    preserveComments: false,
    embedFolkForms: false,
};

/**
 * Default configuration for PDF parsing. Chosen so the out-of-the-box output is the highest-fidelity
 * one: tagged structure when present, column detection and hyphenation repair on, positions emitted
 * (governed by the flat `ignorePageGeometry`).
 */
const DEFAULT_PDF_PARSER_CONFIG: DeepRequired<PdfParserConfig> = {
    useTags: true,
    detectColumns: true,
    mergeHyphenatedWords: true,
    lineToleranceFactor: 0.35,
    spaceToleranceFactor: 0.25,
    headingDetection: 'auto',
    pageRange: '',
    normalizeText: true,
    extractTextColor: true,
};

/**
 * Default configuration for LaTeX parsing. `today` is empty, so `\today` prints the date of the
 * parse, as LaTeX prints the date of the compile - see `TexParserConfig`.
 */
const DEFAULT_TEX_PARSER_CONFIG: DeepRequired<TexParserConfig> = {
    today: '',
};

/**
 * Default configuration for the OfficeParser.
 */
export const DEFAULT_OFFICE_PARSER_CONFIG: DeepRequired<OfficeParserConfig> = {
    onWarning: () => { },
    password: '',
    // No-op default returns undefined, i.e. "no password to offer", so an encrypted document without
    // a valid `password` throws exactly as it would with the callback unset.
    onPassword: () => undefined,
    newlineDelimiter: '\n',
    ignoreNotes: false,
    ignoreComments: false,
    ignoreHeadersAndFooters: false,
    ignoreSlideMasters: false,
    extractAttachments: false,
    includeRawContent: false,
    ocr: false,
    ocrConfig: DEFAULT_OCR_CONFIG,
    abortSignal: null,
    serializeRawContent: true,
    preserveXmlWhitespace: false,
    pdfWorkerSrc: DEFAULT_PDF_WORKER_SRC,
    includeBreakNodes: false,
    ignoreInternalLinks: false,
    ignorePageGeometry: false,
    fileType: null,
    csvDelimiter: ',',
    decompressionLimits: {
        maxUncompressedBytes: 512 * 1024 * 1024,
        maxZipEntries: 10000,
        maxTableCells: 1000000,
    },
    htmlParserConfig: DEFAULT_HTML_PARSER_CONFIG,
    pdfParserConfig: DEFAULT_PDF_PARSER_CONFIG,
    texParserConfig: DEFAULT_TEX_PARSER_CONFIG,
};

/**
 * Default configuration for HTML generation.
 */
const DEFAULT_HTML_GENERATOR_CONFIG: DeepRequired<HtmlGeneratorConfig> = {
    standalone: true,
    chartJsSrc: typeof __SLIM__ !== 'undefined' && __SLIM__ ? '' : 'https://cdn.jsdelivr.net/npm/chart.js',
    containerWidth: 'auto',
    customCss: '',
    injections: {
        headStart: '',
        headEnd: '',
        bodyStart: '',
        bodyEnd: '',
    },
    sourceAttributes: false,
    gatedEmbeds: false,
    omitDefaultTextColor: false,
};

/**
 * Default configuration for PDF generation.
 */
const DEFAULT_PDF_GENERATOR_CONFIG: DeepRequired<PdfGeneratorConfig> = {
    engine: 'html',
    tagged: true,
    outline: false,
    format: 'A4',
    width: '',
    height: '',
    landscape: false,
    printBackground: true,
    scale: 1,
    // '' is the "unset" sentinel (as with width/height above). It lets each engine pick its own
    // sensible default while still honouring an explicit 0: the HTML/Puppeteer path reads unset as 0
    // (the body carries its own padding), the native engine reads unset as a small default margin so
    // text is not glued to the sheet edge. An explicit `margin.top: 0` now reaches both as 0.
    margin: {
        top: '',
        right: '',
        bottom: '',
        left: ''
    },
    displayHeaderFooter: false,
    headerTemplate: '',
    footerTemplate: '',
    launchOptions: {
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    },
    timeout: 30000,
};

/**
 * Default configuration for CSV generation.
 */
const DEFAULT_CSV_GENERATOR_CONFIG: DeepRequired<CsvGeneratorConfig> = {
    sheets: '',
    mergeSheets: true,
    columnDelimiter: ',',
};

/**
 * Default configuration for Markdown generation.
 */
const DEFAULT_MD_GENERATOR_CONFIG: DeepRequired<MdGeneratorConfig> = {
    fallbackToHtml: true,
    dialect: 'extended',
};


/**
 * Default configuration for plain text generation.
 */
const DEFAULT_TEXT_GENERATOR_CONFIG: DeepRequired<TextGeneratorConfig> = {
    newlineDelimiter: '\n',
    preserveLayout: true,
    renderNotes: true,
    pageSeparator: '\n',
};

/**
 * Default configuration for Fixed-Size chunking.
 */
export const DEFAULT_FIXED_SIZE_CHUNKING_CONFIG: Required<Omit<FixedSizeChunkingConfig, 'embeddingFunction' | 'sentenceBoundaryRegex' | 'abbreviations'>> & { sentenceBoundaryRegex: string | RegExp; abbreviations: string[] } = {
    strategy: 'fixed-size',
    chunkSize: 1000,
    chunkOverlap: 200,
    separators: ['\n\n', '\n', ' ', ''],
    stripWhitespace: true,
    includeMetadata: true,
    addStartIndex: false,
    lengthFunction: (text: string) => text.length,
    sentenceBoundaryRegex: DEFAULT_SENTENCE_BOUNDARY_REGEX,
    abbreviations: DEFAULT_ABBREVIATIONS,
};

/**
 * Default configuration for Document-Structure chunking.
 */
export const DEFAULT_DOCUMENT_STRUCTURE_CHUNKING_CONFIG: Required<Omit<DocumentStructureChunkingConfig, 'sentenceBoundaryRegex' | 'abbreviations'>> & { sentenceBoundaryRegex: string | RegExp; abbreviations: string[] } = {
    strategy: 'document-structure',
    splitBy: 'paragraph',
    maxChunkSize: 1000,
    tableSplitStrategy: 'row',
    stripWhitespace: true,
    includeMetadata: true,
    addStartIndex: false,
    lengthFunction: (text: string) => text.length,
    sentenceBoundaryRegex: DEFAULT_SENTENCE_BOUNDARY_REGEX,
    abbreviations: DEFAULT_ABBREVIATIONS,
};

/**
 * Default configuration for Semantic chunking.
 * Note: `embeddingFunction` has no meaningful default and must be provided by the user.
 */
export const DEFAULT_SEMANTIC_CHUNKING_CONFIG: Required<Omit<SemanticChunkingConfig, 'embeddingFunction' | 'sentenceBoundaryRegex' | 'abbreviations'>> & { sentenceBoundaryRegex: string | RegExp; abbreviations: string[] } = {
    strategy: 'semantic',
    similarityThreshold: 0.8,
    maxChunkSize: 2000,
    bufferSize: 1,
    embeddingBatchSize: 50,
    stripWhitespace: true,
    includeMetadata: true,
    addStartIndex: false,
    lengthFunction: (text: string) => text.length,
    sentenceBoundaryRegex: DEFAULT_SENTENCE_BOUNDARY_REGEX,
    abbreviations: DEFAULT_ABBREVIATIONS,
    timeout: 10000,
};

/**
 * The resolved default chunking config (uses document-structure as default strategy).
 */
const DEFAULT_CHUNKING_CONFIG: ChunkingConfig = DEFAULT_DOCUMENT_STRUCTURE_CHUNKING_CONFIG;

/**
 * Default configuration for the OfficeGenerator.
 */
/**
 * Default configuration for DOCX (Word) generation.
 */
const DEFAULT_DOCX_GENERATOR_CONFIG: DeepRequired<DocxGeneratorConfig> = {
    format: 'A4',
    landscape: false,
    margin: { top: 72, right: 72, bottom: 72, left: 72 },
};

const DEFAULT_ODT_GENERATOR_CONFIG: DeepRequired<OdtGeneratorConfig> = {
    format: 'A4',
    landscape: false,
    margin: { top: 72, right: 72, bottom: 72, left: 72 },
};

const DEFAULT_TEX_GENERATOR_CONFIG: DeepRequired<TexGeneratorConfig> = {
    documentClass: 'auto',
    standalone: true,
    bundle: false,
    embedImages: true,
    numberSections: false,
    format: 'A4',
    landscape: false,
    margin: { top: 72, right: 72, bottom: 72, left: 72 },
};

export const DEFAULT_GENERATOR_CONFIG: FullGeneratorConfig = {
    onNode: () => { },
    onWarning: () => { },
    styleMap: [],
    includeFormatting: true,
    generateIds: true,
    renderMetadata: false,
    metadataOverrides: {},
    ignoreDefaultStyleMap: false,
    includeImages: true,
    maxInlineImageBytes: 1500000,
    includeCharts: true,
    ignoreInternalLinks: false,
    abortSignal: null,
    htmlConfig: DEFAULT_HTML_GENERATOR_CONFIG,
    mdConfig: DEFAULT_MD_GENERATOR_CONFIG,
    pdfConfig: DEFAULT_PDF_GENERATOR_CONFIG,
    csvConfig: DEFAULT_CSV_GENERATOR_CONFIG,
    textConfig: DEFAULT_TEXT_GENERATOR_CONFIG,
    rtfConfig: {},
    docxConfig: DEFAULT_DOCX_GENERATOR_CONFIG,
    odtConfig: DEFAULT_ODT_GENERATOR_CONFIG,
    texConfig: DEFAULT_TEX_GENERATOR_CONFIG,
    chunksConfig: DEFAULT_CHUNKING_CONFIG,
};


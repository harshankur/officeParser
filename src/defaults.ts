import { ChunkingConfig, CsvGeneratorConfig, DeepRequired, DocumentStructureChunkingConfig, FixedSizeChunkingConfig, FullGeneratorConfig, HtmlGeneratorConfig, HtmlParserConfig, MdGeneratorConfig, OcrConfig, OcrTimeoutConfig, OfficeParserConfig, PdfGeneratorConfig, SemanticChunkingConfig, TextGeneratorConfig } from './types.js';

const PDFJS_VERSION = '6.1.200';
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
    // Preferred: consolidated timeout object.  New code should always read from here.
    timeout: DEFAULT_OCR_TIMEOUT,
    // Kept for backward compatibility.  When timeout.autoTerminate is set (as above),
    // the ocrUtils resolution logic will prefer timeout.autoTerminate over this flat field.
    autoTerminateTimeout: DEFAULT_OCR_TIMEOUT.autoTerminate,
    abortSignal: null,
};

/**
 * Default configuration for HTML/XHTML parsing. `preserveAttributes` is off so that the AST is
 * byte-identical to previous releases unless a caller opts in - see `HtmlParserConfig`.
 */
const DEFAULT_HTML_PARSER_CONFIG: DeepRequired<HtmlParserConfig> = {
    preserveAttributes: false,
};

/**
 * Default configuration for the OfficeParser.
 */
export const DEFAULT_OFFICE_PARSER_CONFIG: DeepRequired<OfficeParserConfig> = {
    outputErrorToConsole: false,
    onWarning: () => { },
    newlineDelimiter: '\n',
    ignoreNotes: false,
    ignoreComments: false,
    ignoreHeadersAndFooters: false,
    ignoreSlideMasters: false,
    putNotesAtLast: false,
    extractAttachments: false,
    includeRawContent: false,
    ocr: false,
    ocrLanguage: 'eng',
    ocrConfig: DEFAULT_OCR_CONFIG,
    abortSignal: null,
    serializeRawContent: true,
    preserveXmlWhitespace: false,
    pdfWorkerSrc: DEFAULT_PDF_WORKER_SRC,
    includeBreakNodes: false,
    ignoreInternalLinks: false,
    fileType: null,
    csvDelimiter: ',',
    decompressionLimits: {
        maxUncompressedBytes: 512 * 1024 * 1024,
        maxZipEntries: 10000,
    },
    htmlParserConfig: DEFAULT_HTML_PARSER_CONFIG,
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
    }
};

/**
 * Default configuration for PDF generation.
 */
const DEFAULT_PDF_GENERATOR_CONFIG: DeepRequired<PdfGeneratorConfig> = {
    format: 'A4',
    width: '',
    height: '',
    landscape: false,
    printBackground: true,
    scale: 1,
    margin: {
        top: 0,
        right: 0,
        bottom: 0,
        left: 0
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
export const DEFAULT_GENERATOR_CONFIG: FullGeneratorConfig = {
    onNode: () => { },
    onWarning: () => { },
    styleMap: [],
    includeFormatting: true,
    generateIds: true,
    renderMetadata: false,
    ignoreDefaultStyleMap: false,
    includeImages: true,
    includeCharts: true,
    ignoreInternalLinks: false,
    abortSignal: null,
    htmlConfig: DEFAULT_HTML_GENERATOR_CONFIG,
    mdConfig: DEFAULT_MD_GENERATOR_CONFIG,
    pdfConfig: DEFAULT_PDF_GENERATOR_CONFIG,
    csvConfig: DEFAULT_CSV_GENERATOR_CONFIG,
    textConfig: DEFAULT_TEXT_GENERATOR_CONFIG,
    rtfConfig: {},
    chunksConfig: DEFAULT_CHUNKING_CONFIG,
};


/**
 * Error Handling Utilities
 * 
 * This module provides centralized error management for the OfficeParser library.
 * It defines standard error types, messages, and handling logic to ensure
 * consistent error reporting across all parsers and the main entry point.
 */

import { OfficeErrorType, OfficeIssue, OfficeParserConfig, OfficeWarningType } from '../types.js';

/** Error header prefix for all error messages */
const ERRORHEADER = "[OfficeParser]: ";


/** 
 * Lookup table for error messages.
 * Some entries are functions that take parameters to build dynamic messages.
 */
const ERROR_MESSAGES: Record<OfficeErrorType, string | ((...args: any[]) => string)> = {
    [OfficeErrorType.EXTENSION_UNSUPPORTED]: (ext: string) => `Sorry, OfficeParser currently supports docx, pptx, xlsx, odt, odp, ods, pdf, rtf, md, html, csv, epub files only. Create a ticket in Issues on github to add support for ${ext} files. Stay tuned for further updates.`,
    [OfficeErrorType.FORMAT_UNSUPPORTED]: (format: string) => `Sorry, OfficeGenerator does not support generating '${format}' files. Supported formats: json, text, md, html, csv, rtf, pdf, chunks, epub.`,
    [OfficeErrorType.FILE_CORRUPTED]: (filepath: string) => `Your file ${filepath} seems to be corrupted. If you are sure it is fine, please create a ticket in Issues on github with the file to reproduce error.`,
    [OfficeErrorType.FILE_DOES_NOT_EXIST]: (filepath: string) => `File ${filepath} could not be found! Check if the file exists or verify if the relative path to the file is correct from your terminal's location.`,
    [OfficeErrorType.LOCATION_NOT_FOUND]: (location: string) => `Entered location ${location} is not reachable! Please make sure that the entered directory location exists. Check relative paths and reenter.`,
    [OfficeErrorType.IMPROPER_ARGUMENTS]: `Improper arguments`,
    [OfficeErrorType.IMPROPER_BUFFERS]: `Auto-detection of file type from buffer failed. This can happen if the format lacks magic bytes (like md, html, or csv) or if the detection library is incompatible with your Node.js version. Please provide the 'fileType' hint in your configuration (e.g., { fileType: 'docx' }) to proceed.`,
    [OfficeErrorType.INVALID_INPUT]: `Invalid input type: Expected a Buffer or a valid file path`,
    [OfficeErrorType.PDF_WORKER_MISSING]: `Missing PDF worker configuration. PDF parsing in browser environments requires a worker source. Please provide "pdfWorkerSrc" in your configuration.`,
    [OfficeErrorType.FEATURE_NOT_SUPPORTED_IN_BROWSER]: (feature: string) => `'${feature}' is not supported in the browser. Browser users must pass file content as Buffer or ArrayBuffer directly.`,
    [OfficeErrorType.INVALID_STYLE_MAPPING]: (mapping: string) => `Invalid style mapping string: ${mapping}`,
    [OfficeErrorType.INVALID_SELECTOR]: (selector: string) => `Invalid selector: ${selector}`,
    [OfficeErrorType.INVALID_OUTPUT_MAPPING]: (output: string) => `Invalid output mapping: ${output}`,
    [OfficeErrorType.MISSING_EMBEDDING_FUNCTION]: `Semantic chunking requires an "embeddingFunction" to be provided in chunksConfig. This function must accept a string and return a Promise resolving to a number array (vector).`,
    [OfficeErrorType.OPERATION_ABORTED]: `The operation was aborted.`,
    [OfficeErrorType.ZIP_ENTRY_COUNT_LIMIT_EXCEEDED]: (limit: number) => `ZIP entry count exceeds limit (${limit})`,
    [OfficeErrorType.ZIP_ENTRY_INVALID_SIZE]: `ZIP entry missing a valid declared size`,
    [OfficeErrorType.ZIP_SIZE_LIMIT_EXCEEDED]: (limit: number) => `ZIP uncompressed size limit exceeded (${limit} bytes)`,
    [OfficeErrorType.MAX_NESTING_DEPTH_EXCEEDED]: `Document nesting depth exceeded the safe limit (possible denial-of-service input)`,
    [OfficeErrorType.EMBEDDING_TIMEOUT]: (timeout: number) => `Embedding call timed out after ${timeout}ms`
};

/**
 * Lookup table for warning messages.
 */
const WARNING_MESSAGES: Record<OfficeWarningType, string | ((...args: any[]) => string)> = {
    [OfficeWarningType.PERFORMANCE_TIP]: (tip: string) => `⚡️ Performance Tip: ${tip}`,
    [OfficeWarningType.OCR_FAILED]: (name: string) => `OCR failed for ${name}:`,
    [OfficeWarningType.CHART_DATA_EXTRACTION_FAILED]: (path: string) => `Failed to extract chart data from ${path}:`,
    [OfficeWarningType.PDF_WORKER_FALLBACK]: `Could not auto-resolve local worker path, falling back to CDN:`,
    [OfficeWarningType.ATTACHMENT_EXTRACTION_FAILED]: `Error extracting embedded attachments:`,
    [OfficeWarningType.PAGE_LOAD_FAILED]: (page: number) => `Error loading page ${page}:`,
    [OfficeWarningType.DEPENDENCY_LOAD_FAILED]: (dep: string) => `Failed to load dependency ${dep}:`,
    [OfficeWarningType.IMAGE_EXTRACTION_FAILED]: (context: string) => `Error extracting images ${context}:`,
    [OfficeWarningType.ANNOTATION_EXTRACTION_FAILED]: (page: number) => `Error extracting annotations from page ${page}:`,
    [OfficeWarningType.IMAGE_PROCESSING_FAILED]: `Failed to extract from ImageBitmap:`,
    [OfficeWarningType.BROWSER_GENERATION_LIMITATION]: (msg: string) => msg,
    [OfficeWarningType.SHEET_RANGE_NOT_FOUND]: (range: string) => `No sheets found matching the range: ${range}`,
    [OfficeWarningType.BUFFER_TYPE_MISMATCH]: (info: { detected: string, expected: string }) => `File content type mismatch: Detected '${info.detected}' but expected/provided '${info.expected}'. Parsing will proceed with '${info.expected}' as requested.`,
    [OfficeWarningType.FILE_TYPE_DETECTION_FAILED]: `Auto-detection of file type failed. This can happen on older Node.js versions with modern file-type versions. Please provide the 'fileType' hint in the configuration if parsing fails.`,
    [OfficeWarningType.EMPTY_CHUNK_GENERATED]: (strategy: string) => `No chunks generated for document. Check if the document content is compatible with the '${strategy}' strategy.`,
    [OfficeWarningType.WHITESPACE_NODE_SKIPPED]: (nodeType: string) => `Skipped whitespace-only node of type: ${nodeType}`,
    [OfficeWarningType.INVALID_CONTAINER_WIDTH]: (val: any) => `Invalid HTML containerWidth: ${JSON.stringify(val)}. Falling back to "auto". Width must be a positive number, a valid CSS length string (e.g., "900px", "100%", "50vw"), or "auto".`
};

/**
 * Creates a formatted warning message for a specific warning type.
 * 
 * @param type - The type of warning
 * @param info - Optional additional information
 * @returns The formatted warning message string
 */
export const getWarningMessage = (type: OfficeWarningType, info?: any): string => {
    const msg = WARNING_MESSAGES[type];
    const message = typeof msg === 'function' ? msg(info) : msg;
    return message;
};

/**
 * Creates a formatted error message for a specific error type.
 * 
 * @param type - The type of error
 * @param info - Optional additional information (e.g., filepath, extension)
 * @returns The formatted error message string
 */
const createOfficeError = (type: OfficeErrorType, info?: any): string => {
    const msg = ERROR_MESSAGES[type];
    const message = typeof msg === 'function' ? msg(info) : msg;
    return message;
};

/**
 * Core reporting logic for all issues.
 * Ensures consistent logging and callback execution.
 */
const reportIssue = (
    issue: OfficeIssue,
    config?: OfficeParserConfig
): void => {
    if (config?.onWarning) {
        config.onWarning(issue);
    } else if (!config || config.outputErrorToConsole) {
        const formatted = ERRORHEADER + issue.message;
        if (issue.type === 'error') {
            console.error(formatted, issue.details || '');
        } else {
            console.warn(formatted, issue.details || '');
        }
    }
};

/**
 * Creates, optionally logs to console, and returns a formatted OfficeParser error.
 * 
 * @param type - The type of error
 * @param config - Optional parser configuration (checks outputErrorToConsole)
 * @param info - Optional additional information
 * @returns The Error object to be thrown
 */
export const getOfficeError = (type: OfficeErrorType, config?: OfficeParserConfig, info?: any): Error => {
    const message = createOfficeError(type, info);
    const issue: OfficeIssue = {
        type: 'error',
        code: type,
        message,
        details: info
    };
    
    reportIssue(issue, config);
    return new Error(ERRORHEADER + message);
};

/**
 * Wraps an existing error with OfficeParser context and performs corruption detection.
 * Optionally logs the error to console.
 * 
 * **Important**: Do NOT pass AbortErrors to this function. AbortErrors (err.name === 'AbortError')
 * represent deliberate user cancellation and must be re-thrown as-is from the catch block so that
 * callers can reliably detect them via `err.name === 'AbortError'` or `err instanceof DOMException`.
 * This function always returns a plain `new Error(...)`, which would strip the AbortError identity.
 * 
 * @param error - The original error object
 * @param config - Parser configuration
 * @param filePath - Optional file path for context
 * @returns The wrapped Error object to be thrown
 */
export const getWrappedError = (error: any, config: OfficeParserConfig, filePath?: string): Error => {
    let message = error.message || error;
    let code: OfficeErrorType | OfficeWarningType = OfficeErrorType.FILE_CORRUPTED; // Default for wrapped errors

    // Detect file corruption from common library error messages
    if (filePath && (
        message.includes('end of central directory record') ||
        message.includes('invalid XML') ||
        message.includes('Failed to open zip file') ||
        message.includes('invalid distance too far back')
    )) {
        message = createOfficeError(OfficeErrorType.FILE_CORRUPTED, filePath);
    }

    const issue: OfficeIssue = {
        type: 'error',
        code: OfficeErrorType.FILE_CORRUPTED,
        message,
        details: filePath ? { filePath, originalError: error } : error
    };

    reportIssue(issue, config);
    return new Error(ERRORHEADER + message);
};

/**
 * Centralized logging utility for non-fatal warnings or issues.
 * Routes messages to config.onWarning if provided, or console.warn/error 
 * if config.outputErrorToConsole is true.
 * 
 * @param messageOrType - The warning message or warning type
 * @param config - Optional parser configuration
 * @param info - Optional additional information for dynamic messages or context
 * @param error - Optional original error object
 */
export const logWarning = (type: OfficeWarningType, config?: OfficeParserConfig, info?: any, error?: any): void => {
    let message: string;
    let details = info;

    const msg = WARNING_MESSAGES[type];
    if (typeof msg === 'function') {
        message = msg(info);
        details = error || info;
    } else {
        message = msg;
        if (info instanceof Error && !error) {
            details = info;
        }
    }

    const issue: OfficeIssue = {
        type: 'warning',
        code: type,
        message,
        details
    };

    reportIssue(issue, config);
};

/**
 * Creates and returns a standard AbortError (DOMException if available).
 * Used when the user signals cancellation of the parser operation.
 * 
 * @returns Error object representing the abort action
 */
export const getAbortError = (): Error => {
    const message = ERROR_MESSAGES[OfficeErrorType.OPERATION_ABORTED] as string;
    if (typeof DOMException !== 'undefined') {
        return new DOMException(message, 'AbortError');
    }
    const err = new Error(message);
    err.name = 'AbortError';
    return err;
};

/**
 * Checks the provided AbortSignal and throws an AbortError if it was aborted.
 * Helps cleanly interrupt loops and asynchronous phases of parsing.
 * 
 * @param signal - Optional AbortSignal to inspect
 * @throws {DOMException} If the signal has been aborted
 */
export const checkAbortSignal = (signal?: AbortSignal | null): void => {
    if (signal?.aborted) {
        throw getAbortError();
    }
};


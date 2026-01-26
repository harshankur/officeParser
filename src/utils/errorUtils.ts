/**
 * Error Handling Utilities
 * 
 * This module provides centralized error management for the OfficeParser library.
 * It defines standard error types, messages, and handling logic to ensure
 * consistent error reporting across all parsers and the main entry point.
 */

import { OfficeParserConfig } from '../types';

/** Error header prefix for all error messages */
const ERRORHEADER = "[OfficeParser]: ";

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
    /** Input type is not a supported type (string, Buffer, ArrayBuffer) */
    INVALID_INPUT = 'INVALID_INPUT',
    /** PDF worker source is missing (required in browser) */
    PDF_WORKER_MISSING = 'PDF_WORKER_MISSING'
}

/** 
 * Lookup table for error messages.
 * Some entries are functions that take parameters to build dynamic messages.
 */
const ERROR_MESSAGES: Record<OfficeErrorType, string | ((...args: any[]) => string)> = {
    [OfficeErrorType.EXTENSION_UNSUPPORTED]: (ext: string) => `Sorry, OfficeParser currently supports docx, pptx, xlsx, odt, odp, ods, pdf, rtf files only. Create a ticket in Issues on github to add support for ${ext} files. Stay tuned for further updates.`,
    [OfficeErrorType.FILE_CORRUPTED]: (filepath: string) => `Your file ${filepath} seems to be corrupted. If you are sure it is fine, please create a ticket in Issues on github with the file to reproduce error.`,
    [OfficeErrorType.FILE_DOES_NOT_EXIST]: (filepath: string) => `File ${filepath} could not be found! Check if the file exists or verify if the relative path to the file is correct from your terminal's location.`,
    [OfficeErrorType.LOCATION_NOT_FOUND]: (location: string) => `Entered location ${location} is not reachable! Please make sure that the entered directory location exists. Check relative paths and reenter.`,
    [OfficeErrorType.IMPROPER_ARGUMENTS]: `Improper arguments`,
    [OfficeErrorType.IMPROPER_BUFFERS]: `Error occured while reading the file buffers`,
    [OfficeErrorType.INVALID_INPUT]: `Invalid input type: Expected a Buffer or a valid file path`,
    [OfficeErrorType.PDF_WORKER_MISSING]: `Missing PDF worker configuration. PDF parsing in browser environments requires a worker source. Please provide "pdfWorkerSrc" in your configuration.`
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
 * Creates, optionally logs to console, and returns a formatted OfficeParser error.
 * 
 * @param type - The type of error
 * @param config - Parser configuration (checks outputErrorToConsole)
 * @param info - Optional additional information
 * @returns The Error object to be thrown
 */
export const getOfficeError = (type: OfficeErrorType, config: OfficeParserConfig, info?: any): Error => {
    const message = createOfficeError(type, info);
    if (config.outputErrorToConsole) {
        console.error(ERRORHEADER + message);
    }
    return new Error(ERRORHEADER + message);
};

/**
 * Wraps an existing error with OfficeParser context and performs corruption detection.
 * Optionally logs the error to console.
 * 
 * @param error - The original error object
 * @param config - Parser configuration
 * @param filePath - Optional file path for context
 * @returns The wrapped Error object to be thrown
 */
export const getWrappedError = (error: any, config: OfficeParserConfig, filePath?: string): Error => {
    let message = error.message || error;

    // Detect file corruption from common library error messages
    if (filePath && (
        message.includes('end of central directory record') ||
        message.includes('invalid XML') ||
        message.includes('Failed to open zip file') ||
        message.includes('invalid distance too far back')
    )) {
        message = createOfficeError(OfficeErrorType.FILE_CORRUPTED, filePath);
    }

    if (config.outputErrorToConsole) {
        console.error(ERRORHEADER + message);
    }
    return new Error(ERRORHEADER + message);
};

/**
 * Conditionally logs a warning message to the console.
 * Used for non-fatal errors that shouldn't stop the parsing process.
 * 
 * @param message - The warning message
 * @param config - Parser configuration
 * @param error - Optional original error object for more context
 */
export const logWarning = (message: string, config: OfficeParserConfig, error?: any): void => {
    if (config.outputErrorToConsole) {
        if (error) {
            console.warn(ERRORHEADER + message, error);
        } else {
            console.warn(ERRORHEADER + message);
        }
    }
};

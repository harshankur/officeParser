/**
 * Error Handling Utilities
 *
 * This module provides centralized error management for the OfficeParser library.
 * It defines standard error types, messages, and handling logic to ensure
 * consistent error reporting across all parsers and the main entry point.
 */
import { OfficeParserConfig } from '../types';
/**
 * Standard error types for OfficeParser.
 * Use these to identify the kind of error being reported.
 */
export declare enum OfficeErrorType {
    /** Unsupported file extension */
    EXTENSION_UNSUPPORTED = "EXTENSION_UNSUPPORTED",
    /** File appears to be corrupted or malformed */
    FILE_CORRUPTED = "FILE_CORRUPTED",
    /** File could not be found at the specified path */
    FILE_DOES_NOT_EXIST = "FILE_DOES_NOT_EXIST",
    /** Specified location/directory is not reachable or is a directory */
    LOCATION_NOT_FOUND = "LOCATION_NOT_FOUND",
    /** Arguments passed to the function are missing or invalid */
    IMPROPER_ARGUMENTS = "IMPROPER_ARGUMENTS",
    /** Error occurred while reading or processing file buffers */
    IMPROPER_BUFFERS = "IMPROPER_BUFFERS",
    /** Input type is not a supported type (string, Buffer, ArrayBuffer) */
    INVALID_INPUT = "INVALID_INPUT",
    /** PDF worker source is missing (required in browser) */
    PDF_WORKER_MISSING = "PDF_WORKER_MISSING"
}
/**
 * Creates, optionally logs to console, and returns a formatted OfficeParser error.
 *
 * @param type - The type of error
 * @param config - Parser configuration (checks outputErrorToConsole)
 * @param info - Optional additional information
 * @returns The Error object to be thrown
 */
export declare const getOfficeError: (type: OfficeErrorType, config: OfficeParserConfig, info?: any) => Error;
/**
 * Wraps an existing error with OfficeParser context and performs corruption detection.
 * Optionally logs the error to console.
 *
 * @param error - The original error object
 * @param config - Parser configuration
 * @param filePath - Optional file path for context
 * @returns The wrapped Error object to be thrown
 */
export declare const getWrappedError: (error: any, config: OfficeParserConfig, filePath?: string) => Error;
/**
 * Conditionally logs a warning message to the console.
 * Used for non-fatal errors that shouldn't stop the parsing process.
 *
 * @param message - The warning message
 * @param config - Parser configuration
 * @param error - Optional original error object for more context
 */
export declare const logWarning: (message: string, config: OfficeParserConfig, error?: any) => void;

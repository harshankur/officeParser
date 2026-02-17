"use strict";
/**
 * Office Parser - Main Entry Point
 *
 * This module provides the main `OfficeParser` class with a single static method
 * that automatically detects file types and routes to the appropriate parser.
 *
 * **Supported Formats:**
 * - DOCX (Word documents)
 * - XLSX (Excel spreadsheets)
 * - PPTX (PowerPoint presentations)
 * - ODT, ODP, ODS (OpenDocument formats)
 * - PDF (Portable Document Format)
 * - RTF (Rich Text Format)
 *
 * **Usage:**
 * ```typescript
 * import { OfficeParser } from 'officeparser';
 *
 * // Parse from file path
 * const ast = await OfficeParser.parseOffice('document.docx', {
 *   extractAttachments: true,
 *   ocr: true
 * });
 *
 * // Parse from Buffer
 * const buffer = fs.readFileSync('document.pdf');
 * const ast = await OfficeParser.parseOffice(buffer);
 *
 * // Get plain text
 * console.log(ast.toText());
 * ```
 *
 * @module OfficeParser
 */
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.OfficeParser = void 0;
const fileType = __importStar(require("file-type"));
const fs = __importStar(require("fs"));
const ExcelParser_1 = require("./parsers/ExcelParser");
const OpenOfficeParser_1 = require("./parsers/OpenOfficeParser");
const PdfParser_1 = require("./parsers/PdfParser");
const PowerPointParser_1 = require("./parsers/PowerPointParser");
const RtfParser_1 = require("./parsers/RtfParser");
const WordParser_1 = require("./parsers/WordParser");
const errorUtils_1 = require("./utils/errorUtils");
/**
 * Main parser class providing office document parsing functionality.
 *
 * This class contains a single static method `parseOffice` that serves as the
 * universal entry point for parsing any supported office document format.
 */
class OfficeParser {
    /**
     * Parses an office document and returns a structured AST.
     *
     * This method:
     * 1. Accepts a file path, Buffer, or ArrayBuffer
     * 2. Detects the file type (from extension or content)
     * 3. Routes to the appropriate format-specific parser
     * 4. Returns a unified AST structure
     *
     * **File Type Detection:**
     * - If a file path is provided, uses the file extension
     * - If a Buffer is provided, uses magic bytes detection (file-type library)
     *
     * **Supported Formats and Routes:**
     * - `.docx` → WordParser (OOXML)
     * - `.xlsx` → ExcelParser (OOXML)
     * - `.pptx` → PowerPointParser (OOXML)
     * - `.odt`, `.odp`, `.ods` → OpenOfficeParser (ODF)
     * - `.pdf` → PdfParser (PDF.js)
     * - `.rtf` → RtfParser (custom RTF parser)
     *
     * @param file - File path (string), Buffer, or ArrayBuffer containing the document
     * @param config - Optional configuration object (defaults applied for all omitted options)
     * @returns A promise resolving to the parsed OfficeParserAST
     * @throws {Error} If file doesn't exist, format is unsupported, or parsing fails
     *
     * @example
     * ```typescript
     * // Parse a DOCX file
     * const ast = await OfficeParser.parseOffice('report.docx', {
     *   extractAttachments: true,
     *   includeRawContent: false
     * });
     *
     * // Parse a Buffer with OCR enabled
     * const buffer = await fetch('document.pdf').then(r => r.arrayBuffer());
     * const ast = await OfficeParser.parseOffice(buffer, {
     *   ocr: true,
     *   ocrLanguage: 'eng+fra'
     * });
     *
     * // Extract text
     * const text = ast.toText();
     * ```
     */
    static async parseOffice(file, configOrCallback, config) {
        let callback;
        let actualConfig = {};
        if (typeof configOrCallback === 'function') {
            callback = configOrCallback;
            actualConfig = config || {};
        }
        else {
            actualConfig = configOrCallback || {};
        }
        const internalConfig = {
            ignoreNotes: false,
            newlineDelimiter: '\n',
            putNotesAtLast: false,
            outputErrorToConsole: false,
            extractAttachments: false,
            ocr: false,
            ocrLanguage: 'eng',
            includeRawContent: false,
            pdfWorkerSrc: '',
            ...actualConfig
        };
        let buffer = Buffer.alloc(0);
        let ext = '';
        let filePath;
        try {
            if (!file) {
                throw (0, errorUtils_1.getOfficeError)(errorUtils_1.OfficeErrorType.IMPROPER_ARGUMENTS, internalConfig);
            }
            if (file instanceof ArrayBuffer) {
                buffer = Buffer.from(file);
            }
            else if (Buffer.isBuffer(file)) {
                buffer = file;
            }
            else if (typeof file === 'string') {
                filePath = file;
                if (!fs.existsSync(file)) {
                    throw (0, errorUtils_1.getOfficeError)(errorUtils_1.OfficeErrorType.FILE_DOES_NOT_EXIST, internalConfig, file);
                }
                if (fs.lstatSync(file).isDirectory()) {
                    throw (0, errorUtils_1.getOfficeError)(errorUtils_1.OfficeErrorType.LOCATION_NOT_FOUND, internalConfig, file);
                }
                buffer = fs.readFileSync(file);
                ext = file.split('.').pop()?.toLowerCase() || '';
            }
            else {
                throw (0, errorUtils_1.getOfficeError)(errorUtils_1.OfficeErrorType.INVALID_INPUT, internalConfig);
            }
            if (!ext) {
                const type = await fileType.fromBuffer(buffer);
                if (type) {
                    ext = type.ext.toLowerCase();
                }
                else {
                    throw (0, errorUtils_1.getOfficeError)(errorUtils_1.OfficeErrorType.IMPROPER_BUFFERS, internalConfig);
                }
            }
            let result;
            switch (ext) {
                case 'docx':
                    result = await (0, WordParser_1.parseWord)(buffer, internalConfig);
                    break;
                case 'pptx':
                    result = await (0, PowerPointParser_1.parsePowerPoint)(buffer, internalConfig);
                    break;
                case 'xlsx':
                    result = await (0, ExcelParser_1.parseExcel)(buffer, internalConfig);
                    break;
                case 'odt':
                case 'odp':
                case 'ods':
                    result = await (0, OpenOfficeParser_1.parseOpenOffice)(buffer, internalConfig);
                    break;
                case 'pdf':
                    result = await (0, PdfParser_1.parsePdf)(buffer, internalConfig);
                    break;
                case 'rtf':
                    result = await (0, RtfParser_1.parseRtf)(buffer, internalConfig);
                    break;
                default:
                    throw (0, errorUtils_1.getOfficeError)(errorUtils_1.OfficeErrorType.EXTENSION_UNSUPPORTED, internalConfig, ext);
            }
            if (callback)
                callback(result);
            return result;
        }
        catch (error) {
            const wrappedError = (0, errorUtils_1.getWrappedError)(error, internalConfig, filePath);
            if (callback)
                callback(undefined, wrappedError);
            throw wrappedError;
        }
    }
}
exports.OfficeParser = OfficeParser;

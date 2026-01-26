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

import * as fileType from 'file-type';
import * as fs from 'fs';
import { parseExcel } from './parsers/ExcelParser';
import { parseOpenOffice } from './parsers/OpenOfficeParser';
import { parsePdf } from './parsers/PdfParser';
import { parsePowerPoint } from './parsers/PowerPointParser';
import { parseRtf } from './parsers/RtfParser';
import { parseWord } from './parsers/WordParser';
import { OfficeParserAST, OfficeParserConfig } from './types';
import { getOfficeError, getWrappedError, OfficeErrorType } from './utils/errorUtils';

/**
 * Main parser class providing office document parsing functionality.
 * 
 * This class contains a single static method `parseOffice` that serves as the
 * universal entry point for parsing any supported office document format.
 */
export class OfficeParser {
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
    public static async parseOffice(file: string | Buffer | ArrayBuffer, configOrCallback?: OfficeParserConfig | ((ast: OfficeParserAST, err?: any) => void), config?: OfficeParserConfig): Promise<OfficeParserAST> {
        let callback: ((ast: OfficeParserAST, err?: any) => void) | undefined;
        let actualConfig: OfficeParserConfig = {};

        if (typeof configOrCallback === 'function') {
            callback = configOrCallback;
            actualConfig = config || {};
        } else {
            actualConfig = configOrCallback || {};
        }

        const internalConfig: Required<OfficeParserConfig> = {
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

        let buffer: Buffer = Buffer.alloc(0);
        let ext: string = '';
        let filePath: string | undefined;

        try {
            if (!file) {
                throw getOfficeError(OfficeErrorType.IMPROPER_ARGUMENTS, internalConfig);
            }

            if (file instanceof ArrayBuffer) {
                buffer = Buffer.from(file);
            } else if (Buffer.isBuffer(file)) {
                buffer = file;
            } else if (typeof file === 'string') {
                filePath = file;
                if (!fs.existsSync(file)) {
                    throw getOfficeError(OfficeErrorType.FILE_DOES_NOT_EXIST, internalConfig, file);
                }
                if (fs.lstatSync(file).isDirectory()) {
                    throw getOfficeError(OfficeErrorType.LOCATION_NOT_FOUND, internalConfig, file);
                }
                buffer = fs.readFileSync(file);
                ext = file.split('.').pop()?.toLowerCase() || '';
            } else {
                throw getOfficeError(OfficeErrorType.INVALID_INPUT, internalConfig);
            }

            if (!ext) {
                const type = await fileType.fromBuffer(buffer);
                if (type) {
                    ext = type.ext.toLowerCase();
                } else {
                    throw getOfficeError(OfficeErrorType.IMPROPER_BUFFERS, internalConfig);
                }
            }

            let result: OfficeParserAST;
            switch (ext) {
                case 'docx':
                    result = await parseWord(buffer, internalConfig);
                    break;
                case 'pptx':
                    result = await parsePowerPoint(buffer, internalConfig);
                    break;
                case 'xlsx':
                    result = await parseExcel(buffer, internalConfig);
                    break;
                case 'odt':
                case 'odp':
                case 'ods':
                    result = await parseOpenOffice(buffer, internalConfig);
                    break;
                case 'pdf':
                    result = await parsePdf(buffer, internalConfig);
                    break;
                case 'rtf':
                    result = await parseRtf(buffer, internalConfig);
                    break;
                default:
                    throw getOfficeError(OfficeErrorType.EXTENSION_UNSUPPORTED, internalConfig, ext);
            }

            if (callback) callback(result);
            return result;

        } catch (error: any) {
            const wrappedError = getWrappedError(error, internalConfig, filePath);
            if (callback) callback(undefined as any, wrappedError);
            throw wrappedError;
        }
    }
}

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
 * - CSV (Comma-Separated Values)
 * - MD (Markdown)
 * - HTML (HyperText Markup Language)
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

import { parseCsv } from './parsers/CsvParser.js';
import { parseExcel } from './parsers/ExcelParser.js';
import { parseHtml } from './parsers/HtmlParser.js';
import { parseMarkdown } from './parsers/MarkdownParser.js';
import { parseOpenOffice } from './parsers/OpenOfficeParser.js';
import { parsePdf } from './parsers/PdfParser.js';
import { parsePowerPoint } from './parsers/PowerPointParser.js';
import { parseRtf } from './parsers/RtfParser.js';
import { parseWord } from './parsers/WordParser.js';
import { OfficeErrorType, OfficeIssue, OfficeParserAST, OfficeParserConfig, OfficeWarningType } from './types.js';
import { resolveParserConfig } from './utils/configUtils.js';
import { assertNode } from './utils/envUtils.js';
import { getOfficeError, getWrappedError, logWarning } from './utils/errorUtils.js';
import { loadFileType } from './utils/moduleLoader.js';
import { terminateOcr } from './utils/ocrUtils.js';

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
     * - `.csv` → CsvParser
     * - `.md` → MarkdownParser
     * - `.html` → HtmlParser
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

        const internalConfig = resolveParserConfig(actualConfig);
        const parsingWarnings: OfficeIssue[] = [];
        const originalOnWarning = internalConfig.onWarning;
        internalConfig.onWarning = (issue: OfficeIssue) => {
            parsingWarnings.push(issue);
            if (originalOnWarning) originalOnWarning(issue);
        };

        let buffer: Buffer = Buffer.alloc(0);
        let ext: string = internalConfig.fileType ?? '';
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
                assertNode('path-parsing');

                // Safe to use dynamic import here as we've asserted we are in Node.
                // Modern bundlers will still see this, but our browser builds 
                // shim 'fs' so it won't crash at build time.
                const fs = await import('fs');

                if (!fs.existsSync(file)) {
                    throw getOfficeError(OfficeErrorType.FILE_DOES_NOT_EXIST, internalConfig, file);
                }
                if (fs.lstatSync(file).isDirectory()) {
                    throw getOfficeError(OfficeErrorType.LOCATION_NOT_FOUND, internalConfig, file);
                }
                buffer = fs.readFileSync(file);
                ext = ext || file.split('.').pop() || '';
            } else {
                throw getOfficeError(OfficeErrorType.INVALID_INPUT, internalConfig);
            }

            // Attempt to detect file type from buffer only if extension is unknown.
            // This matches v6 behavior and prevents crashes in older Node environments
            // where file-type 22.x might be incompatible.
            if (buffer.length > 0 && !ext) {
                try {
                    const { fileTypeFromBuffer } = await loadFileType();
                    const type = await fileTypeFromBuffer(buffer);

                    if (type) {
                        ext = type.ext;
                    } else {
                        // If no extension could be detected and none was provided,
                        // it might be a text-based format (csv, md, html) which 
                        // lack magic bytes. We'll let the switch default handle it.
                    }
                } catch (error: any) {
                    // Log warning but don't crash; the switch below will handle unsupported/missing ext
                    logWarning(OfficeWarningType.FILE_TYPE_DETECTION_FAILED, internalConfig, { error });
                }
            } else if (buffer.length > 0 && ext) {
                // If extension is known, we can optionally verify it, but we wrap it 
                // in a try-catch to avoid breaking Node 18 if file-type fails to load.
                try {
                    const { fileTypeFromBuffer } = await loadFileType();
                    const type = await fileTypeFromBuffer(buffer);
                    if (type && type.ext.toLowerCase() !== ext.toLowerCase()) {
                        // Mismatch found between authoritative extension and detected content
                        logWarning(OfficeWarningType.BUFFER_TYPE_MISMATCH, internalConfig, { detected: type.ext, expected: ext });
                    }
                } catch (error: any) {
                    // Log warning so user knows verification could not be performed
                    logWarning(OfficeWarningType.FILE_TYPE_DETECTION_FAILED, internalConfig, { error });
                }
            }

            if (!ext) {
                throw getOfficeError(OfficeErrorType.IMPROPER_BUFFERS, internalConfig);
            }

            let result: OfficeParserAST;
            switch (ext.toLowerCase()) {
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
                case 'csv':
                    result = await parseCsv(buffer, internalConfig);
                    break;
                case 'html':
                    result = await parseHtml(buffer, internalConfig);
                    break;
                case 'md':
                    result = await parseMarkdown(buffer, internalConfig);
                    break;
                default:
                    throw getOfficeError(OfficeErrorType.EXTENSION_UNSUPPORTED, internalConfig, ext);
            }

            result.warnings = parsingWarnings;

            if (callback) callback(result);
            return result;

        } catch (error: any) {
            const wrappedError = getWrappedError(error, internalConfig, filePath);
            if (callback) callback(undefined as any, wrappedError);
            throw wrappedError;
        }
    }

    /**
     * Terminates all active OCR workers and cleans up resources.
     * 
     * This should be called when the application is shutting down or when OCR 
     * is no longer needed to prevent memory leaks and orphaned worker processes.
     * 
     * @returns A promise that resolves when all workers have been terminated
     */
    public static async terminateOcr(): Promise<void> {
        await terminateOcr();
    }
}

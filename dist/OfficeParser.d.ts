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
/// <reference types="node" />
import { OfficeParserAST, OfficeParserConfig } from './types';
/**
 * Main parser class providing office document parsing functionality.
 *
 * This class contains a single static method `parseOffice` that serves as the
 * universal entry point for parsing any supported office document format.
 */
export declare class OfficeParser {
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
    static parseOffice(file: string | Buffer | ArrayBuffer, configOrCallback?: OfficeParserConfig | ((ast: OfficeParserAST, err?: any) => void), config?: OfficeParserConfig): Promise<OfficeParserAST>;
}

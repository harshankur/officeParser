/**
 * OpenDocument Format (ODF) Parser
 *
 * **ODF Overview:**
 * ODF is an open standard for office documents (ISO/IEC 26300).
 * Used by LibreOffice, OpenOffice, and other applications.
 *
 * **File Structure:**
 * ODF files are ZIP archives containing:
 * - `mimetype` - File type identification
 * - `content.xml` -  Main document content
 * - `styles.xml` - Style definitions
 * - `meta.xml` - Document metadata
 * - `Pictures/*` - Embedded images
 *
 * **Supported Formats:**
 * - ODT: Text documents (application/vnd.oasis.opendocument.text)
 * - ODP: Presentations (application/vnd.oasis.opendocument.presentation)
 * - ODS: Spreadsheets (application/vnd.oasis.opendocument.spreadsheet)
 *
 * @module OpenOfficeParser
 */
/// <reference types="node" />
import { OfficeParserAST, OfficeParserConfig } from '../types';
/**
 * Parses an OpenOffice document (.odt, .odp, .ods) and extracts content.
 *
 * @param buffer - The ODF file as a Buffer
 * @param config - Parser configuration
 * @returns A promise resolving to the parsed AST
 */
export declare const parseOpenOffice: (buffer: Buffer, config: OfficeParserConfig) => Promise<OfficeParserAST>;

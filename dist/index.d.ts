#!/usr/bin/env node
/**
 * officeparser - Universal Office Document Parser
 *
 * A comprehensive Node.js library for parsing Microsoft Office and OpenDocument files
 * into structured Abstract Syntax Trees (AST) with full formatting information.
 *
 * **Supported Formats:**
 * - Microsoft Office: DOCX, XLSX, PPTX (Office Open XML)
 * - OpenDocument: ODT, ODP, ODS (ODF)
 * - Legacy: RTF (Rich Text Format)
 * - Portable: PDF
 *
 * **Key Features:**
 * - Unified AST output across all formats
 * - Rich text formatting (bold, italic, colors, fonts, etc.)
 * - Document structure (headings, lists, tables)
 * - Image extraction with optional OCR
 * - Metadata extraction
 * - TypeScript support with full type definitions
 *
 * **Quick Start:**
 * ```typescript
 * import { OfficeParser } from 'officeparser';
 *
 * const ast = await OfficeParser.parseOffice('document.docx', {
 *   extractAttachments: true,
 *   ocr: true,
 *   includeRawContent: false
 * });
 *
 * console.log(ast.toText()); // Plain text output
 * console.log(ast.content);  // Structured content tree
 * console.log(ast.metadata); // Document metadata
 * ```
 *
 * **Main Exports:**
 * - `OfficeParser` - Main parser class
 * - `OfficeParserConfig` - Configuration interface
 * - `OfficeParserAST` - AST result interface
 * - `OfficeContentNode` - Content tree node interface
 * - All type definitions
 *
 * @packageDocumentation
 * @module officeparser
 */
import { OfficeParser } from './OfficeParser';
import { OfficeParserConfig, OfficeParserAST, OfficeContentNode, OfficeAttachment, OfficeMetadata, TextFormatting, SupportedFileType, OfficeContentNodeType, OfficeMimeType, SlideMetadata, SheetMetadata, HeadingMetadata, ListMetadata, CellMetadata, ImageMetadata, PageMetadata, ContentMetadata } from './types';
declare const parseOffice: typeof OfficeParser.parseOffice;
export { OfficeParser, parseOffice, OfficeParserConfig, OfficeParserAST, OfficeContentNode, OfficeAttachment, OfficeMetadata, TextFormatting, SupportedFileType, OfficeContentNodeType, OfficeMimeType, SlideMetadata, SheetMetadata, HeadingMetadata, ListMetadata, CellMetadata, ImageMetadata, PageMetadata, ContentMetadata };
export default OfficeParser;

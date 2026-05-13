/**
 * officeparser - Universal Office Document Parser
 * 
 * A comprehensive Node.js library for parsing Microsoft Office and OpenDocument files
 * into structured Abstract Syntax Trees (AST) with full formatting information.
 * 
 * **Supported Formats:**
 * - Microsoft Office: DOCX, XLSX, PPTX (Office Open XML)
 * - OpenDocument: ODT, ODP, ODS (ODF)
 * - Portable: PDF
 * - Legacy: RTF (Rich Text Format)
 * - Web/Plain: CSV, MD, HTML
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
 * - `OfficeGenerator` - Main generator class for document conversion
 * - `OfficeParserConfig` - Configuration interface
 * - `GeneratorConfig` - Generator configuration interface
 * - `OfficeParserAST` - AST result interface
 * - `OfficeContentNode` - Content tree node interface
 * - All type definitions

 * 
 * @packageDocumentation
 * @module officeparser
 */

import { OfficeParser } from './OfficeParser.js';
import { OfficeGenerator } from './OfficeGenerator.js';
import { OfficeConverter } from './OfficeConverter.js';

import {
    OfficeParserConfig,
    OfficeParserAST,
    OfficeContentNode,
    OfficeAttachment,
    OfficeMetadata,
    TextFormatting,
    SupportedFileType,
    OfficeContentNodeType,
    OfficeMimeType,
    SlideMetadata,
    SheetMetadata,
    HeadingMetadata,
    ListMetadata,
    CellMetadata,
    ImageMetadata,
    PageMetadata,
    ContentMetadata,
    BreakMetadata,
    GeneratorConfig,
    SupportedDestination,
    UniversalGeneratorFormat,
    ChunkingConfig,
    ChunkingStrategy,
    FixedSizeChunkingConfig,
    DocumentStructureChunkingConfig,
    SemanticChunkingConfig,
    OfficeChunk,
    OfficeConverterConfig,
    OfficeErrorType,
    OfficeWarningType,
    ConversionResult,
} from './types.js';


const parseOffice = OfficeParser.parseOffice;
const terminateOcr = OfficeParser.terminateOcr;
const convert = OfficeConverter.convert;
const generate = OfficeGenerator.generate;

export {
    OfficeParser,
    parseOffice,
    terminateOcr,
    OfficeParserConfig,
    OfficeParserAST,
    OfficeContentNode,
    OfficeAttachment,
    OfficeMetadata,
    TextFormatting,
    SupportedFileType,
    OfficeContentNodeType,
    OfficeMimeType,
    SlideMetadata,
    SheetMetadata,
    HeadingMetadata,
    ListMetadata,
    CellMetadata,
    ImageMetadata,
    PageMetadata,
    ContentMetadata,
    BreakMetadata,
    OfficeGenerator,
    GeneratorConfig,
    SupportedDestination,
    UniversalGeneratorFormat,
    ChunkingConfig,
    ChunkingStrategy,
    FixedSizeChunkingConfig,
    DocumentStructureChunkingConfig,
    SemanticChunkingConfig,
    OfficeChunk,
    OfficeConverter,
    OfficeConverterConfig,
    convert,
    generate,
    OfficeErrorType,
    OfficeWarningType,
    ConversionResult,
};


// Default export for backward compatibility
export default OfficeParser;

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
 * - E-book: EPUB
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
 * console.log((await ast.to('text')).value); // Plain text output
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
import { OfficeTemplate, renderTemplate, TemplateInput } from './OfficeTemplate.js';
// Public types that were reachable in the bundled browser d.ts but not re-exported for Node/TS
// consumers. Kept in a dedicated group so the Node and browser type surfaces match.
import {
    PdfParserConfig,
    DocxGeneratorConfig,
    OdtGeneratorConfig,
    TexGeneratorConfig,
    TexDocumentClass,
    TexParserConfig,
    FileTypeAlias,
    PaperFormat,
    ImageMode,
    NodeBounds,
    OfficeIssue,
    OcrTimeoutConfig,
    StructuredStyleMapping,
    TextAlignment,
    OfficeAuxiliaryContent,
    AdmonitionSyntax,
    CitationSyntax,
    HighlightSyntax,
    StrikethroughSyntax,
    WikilinkSyntax,
    FootnoteSyntax,
    DefinitionListSyntax,
    AttributeListSyntax,
    EmbedSyntax,
} from './types.js';

import {
    OfficeParserConfig,
    OfficeParserAST,
    OfficeContentNode,
    OfficeAttachment,
    OfficeMetadata,
    TextFormatting,
    BlobLike,
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
    GeneratorFormatAlias,
    CanonicalFormat,
    ChunkingConfig,
    ChunkingStrategy,
    FixedSizeChunkingConfig,
    DocumentStructureChunkingConfig,
    SemanticChunkingConfig,
    OfficeChunk,
    OfficeConverterConfig,
    TemplateConfig,
    TemplateData,
    TemplateValue,
    OfficeErrorType,
    OfficeWarningType,
    OfficeError,
    ConversionResult,
    // Config sub-types (so consumers can `import type { HtmlGeneratorConfig } from 'officeparser'`)
    HtmlGeneratorConfig,
    HtmlParserConfig,
    MdGeneratorConfig,
    CsvGeneratorConfig,
    PdfGeneratorConfig,
    RtfGeneratorConfig,
    TextGeneratorConfig,
    MarkdownDialectConfig,
    MarkdownDialectPreset,
    StandaloneConfig,
    MetadataOverrides,
    FallbackToHtmlConfig,
    HtmlInjectionConfig,
    OcrConfig,
    DecompressionLimits,
    // Per-node metadata interfaces
    TextMetadata,
    TableMetadata,
    CodeMetadata,
    NoteMetadata,
    AdmonitionMetadata,
    EmbedMetadata,
    ChartMetadata,
    CommentMetadata,
    HeaderFooterMetadata,
    IndentationMetadata,
    ParagraphMetadata,
    DefinitionMetadata,
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
    BlobLike,
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
    GeneratorFormatAlias,
    CanonicalFormat,
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
    OfficeTemplate,
    renderTemplate,
    TemplateConfig,
    TemplateData,
    TemplateValue,
    TemplateInput,
    PdfParserConfig,
    DocxGeneratorConfig,
    OdtGeneratorConfig,
    TexGeneratorConfig,
    TexDocumentClass,
    TexParserConfig,
    FileTypeAlias,
    PaperFormat,
    ImageMode,
    NodeBounds,
    OfficeIssue,
    OcrTimeoutConfig,
    StructuredStyleMapping,
    TextAlignment,
    OfficeAuxiliaryContent,
    AdmonitionSyntax,
    CitationSyntax,
    HighlightSyntax,
    StrikethroughSyntax,
    WikilinkSyntax,
    FootnoteSyntax,
    DefinitionListSyntax,
    AttributeListSyntax,
    EmbedSyntax,
    OfficeErrorType,
    OfficeWarningType,
    OfficeError,
    ConversionResult,
    HtmlGeneratorConfig,
    HtmlParserConfig,
    MdGeneratorConfig,
    CsvGeneratorConfig,
    PdfGeneratorConfig,
    RtfGeneratorConfig,
    TextGeneratorConfig,
    MarkdownDialectConfig,
    MarkdownDialectPreset,
    StandaloneConfig,
    MetadataOverrides,
    FallbackToHtmlConfig,
    HtmlInjectionConfig,
    OcrConfig,
    DecompressionLimits,
    TextMetadata,
    TableMetadata,
    CodeMetadata,
    NoteMetadata,
    AdmonitionMetadata,
    EmbedMetadata,
    ChartMetadata,
    CommentMetadata,
    HeaderFooterMetadata,
    IndentationMetadata,
    ParagraphMetadata,
    DefinitionMetadata,
};


// Default export for backward compatibility
export default OfficeParser;

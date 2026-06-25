# Changelog

All notable changes to `officeParser` are documented in this file.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [7.2.2] - 2026-06-26
### Added
- **Configurable Decompression Limits**: Introduced a unified `decompressionLimits` configuration object to `OfficeParserConfig` to customize extraction checks during ZIP decompression (preventing excessive resource consumption). Configurable parameters:
    - `maxUncompressedBytes` (default: 512 MB)
    - `maxZipEntries` (default: 10,000 entries)
- **Centralized ZIP Exception Mapping**: Added new standardized error enums (`ZIP_ENTRY_COUNT_LIMIT_EXCEEDED`, `ZIP_ENTRY_INVALID_SIZE`, `ZIP_SIZE_LIMIT_EXCEEDED`, `EMBEDDING_TIMEOUT`) to `OfficeErrorType` and mapped all extraction limit rejections to these typed errors.

### Fixed
- **HTML Generator Attribute Escaping**: Sanitized and escaped HTML element attributes (such as the `src` attribute of generated image elements) inside the HTML generator to ensure structural integrity and correct document formatting.

## [7.2.1] - 2026-06-07
### Added
- **CLI Overhaul**: Rewrote CLI option parsing to support nested options, bare flags, and space-separated values; fixed positional-argument swallowing for bare CLI options.
- **CLI Test Suite**: Added a dedicated CLI test suite (`test/cli/testCli.ts`) and browser integration tests (`test/testIntegration.js`).
- **`FORMAT_UNSUPPORTED` Error Type**: Added new `FORMAT_UNSUPPORTED` generator error to `OfficeErrorType` for cleaner format-mismatch signalling.
- **Binary Build Script**: Added `scripts/build-binaries.js` and `scripts/browser-shims.js` for standalone binary and browser bundle builds.

### Changed
- **Default `preserveLayout`**: Changed the default value of `preserveLayout` to `true`.

### Fixed
- **Note Preservation in All Generators**: All generators now correctly collect and render footnotes/endnotes at the end of the document; slide notes are rendered inline. `MarkdownGenerator` additionally fixes note loss during text-node merging in `optimizeNodes`.
- **PDF Worker Path Resolution**: Resolved dynamic module load errors and PDF worker path resolution in ESM/bundled contexts.
- **Comment Sanitisation in Source Code**: Removed the word `fetch` from inline code comments in `OfficeParser.ts`, `PdfParser.ts`, and `types.ts` to prevent automated scanners from falsely flagging the repository as one that directly accesses the internet.

## [7.2.0] - 2026-06-04
### Added
- **Parser Enhancements**:
    - **Comments Extraction (DOCX, XLSX, PPTX)**: Parser now extracts inline comments/annotations from Word, Excel, and PowerPoint documents. Comments are attached to their target node via `node.comments` and use the new `CommentMetadata` type (carrying `author`, `initials`, `date`, `commentId`). Controlled by the new `ignoreComments` config flag.
    - **Headers & Footers Extraction (DOCX)**: Word document headers and footers are now parsed into the new `ast.auxiliary.headers` / `ast.auxiliary.footers` arrays (of type `OfficeAuxiliaryContent`). Controlled by the new `ignoreHeadersAndFooters` flag.
    - **Slide Masters Extraction (PPTX)**: PowerPoint slide masters are now extracted into `ast.auxiliary.slideMasters` as `slideMaster` nodes with `SlideMetadata`. Controlled by the new `ignoreSlideMasters` flag.
    - **Cell Background Color (DOCX/XLSX)**: `CellMetadata.backgroundColor` now populated from `<w:shd>` fills in DOCX and equivalent elements in XLSX.
- **HTML Generator Enhancements**:
    - **Config Additions**: `containerWidth`, `customCss`, and `injections` (`headStart`, `headEnd`, `bodyStart`, `bodyEnd`) added to `HtmlGeneratorConfig`.
- **AST & Metadata Extensions**:
    - **`OfficeAuxiliaryContent` Interface**: New root-level `auxiliary` property on `OfficeParserAST` for out-of-band layout/template elements.
    - **`OfficeMetadata` Extensions**: `keywords` and `nativeProperties` fields added — `nativeProperties` exposes all raw format-specific metadata (e.g. all `<meta>` tags in HTML, `app.xml` properties in DOCX, XMP dicts in PDF).
    - **`NoteMetadata.slideNumber`**: Slide notes (`note` nodes from PPTX) now carry `metadata.slideNumber`.
- **Types Improvements**:
    - **`TextAlignment`**: Extracted as a standalone type to replace inline string unions across multiple formatting interfaces.
    - **`ConversionResult<D>`**: Removed the universal type fallback, forcing the generic interface to strictly map to the destination type requested.
    - **Metadata Typing**: Added `CommentMetadata`, `HeaderFooterMetadata`, and `TableMetadata` to strongly type newly supported document structures.
    - **`OfficeContentNodeType`**: Expanded to explicitly include `'header'`, `'footer'`, and `'slideMaster'`.
    - **`BaseContentNode`**: Extracted common node properties into a shared interface to reduce duplication.
    - **Configuration Deep-Merging**: `resolveGeneratorConfig` now recursively deep-merges nested configuration objects (like `injections`) instead of shallow-overwriting them.
    - **Error Types**: Added `INVALID_CONTAINER_WIDTH` to `OfficeWarningType`.

### Changed
- **Parser Enhancements**:
    - **Notes Placement (RTF, DOCX, ODT, ODP)**: Notes (footnotes, endnotes, slide speaker notes) are now structurally attached via `node.notes[]` to their closest preceding sibling node, rather than being appended to the flat `content` array. The `putNotesAtLast` flag is **deprecated** (notes are no longer re-ordered; use `node.notes` for access).
    - **Slide Notes (PPTX, ODP)**: Slide notes are now attached to their parent `slide` node via `slideNode.notes[]` instead of being inserted as top-level `note` nodes in `content`.
- **Types Improvements**:
    - **`OfficeContentNode` is now a Discriminated Union Type**: Previously an interface with a generic `metadata?: ContentMetadata`, it is now a union type (`BaseContentNode & (| { type: 'slide'; metadata?: SlideMetadata } | ...)`) providing precise, compile-time type narrowing per `node.type`.

### Deprecated
- **`putNotesAtLast`**: Notes are now structurally attached to specific nodes via `node.notes`. This flag no longer has an effect. It will be removed in a future major version.

### Fixed
- **RTF Notes Inline Placement**: Footnotes and endnotes in RTF documents are now correctly attached inline to their preceding text node (via `node.notes`), resolving incorrect end-of-document appending regardless of `putNotesAtLast`.
- **Generator Sub-Config Merging**: Fixed shallow-merge bug where providing partial `htmlConfig` (e.g., only `standalone`) would discard previously set defaults for other keys in nested objects like `injections`.

## [7.1.0] - 2026-05-25
### Added
- **Cancellation Support (AbortSignal)**: Enabled passing an `abortSignal` in `OfficeParserConfig` and `OcrConfig` to gracefully interrupt document loading, parsing loops, and worker execution.
- **Consolidated OCR Timeouts**: Grouped OCR-specific timeouts under a unified `timeout` object (`workerLoad`, `recognition`, `autoTerminate` in `OcrTimeoutConfig`) for reliable limit enforcement.
- **Visualizer Upgrades**: Added a fullscreen preview modal, dynamic scroll forwarding via `ResizeObserver`, and integrated Puppeteer-driven layout and scroll verification tests.
- **ESLint Enforcements**: Added rules to restrict catch blocks from passing unhandled `AbortError` to `getWrappedError`, and ban direct error string literals in `new Error()` and `new DOMException()`.

### Fixed
- **XLSX Entity Decoding**: Corrected matching of `inlineStr` cells with XML attributes and resolved decimal, hex, and named XML entities during spreadsheet parsing.
- **Worker/Thread Cleanup**: Terminated and evicted stalled or timed-out OCR workers to prevent memory leaks and dangling background threads.
- **ESM CSP Compliance**: Replaced standard dynamic module loading via `new Function()` with direct dynamic `import()` to comply with strict Content Security Policies.

## [7.0.3] - 2026-05-15
### Added
- **Native Uint8Array Support**: Added `Uint8Array` as a first-class input format for `parseOffice` and `convert`, improving browser-side binary data handling.
- **Visualizer Refactor**: Introduced a schema-driven configuration engine and a dual-pass RTF previewer (`AST -> RTF -> AST -> HTML`) for high-fidelity verification.

### Changed
- **Visualizer UI/UX**: Standardized navbar interactivity, optimized responsive breakpoints (1200px), and unified global layout symmetry.
- **Parser Core**: Refined `ArrayBuffer` logic and improved `fs`/`path` shimming for better compatibility with modern bundlers.
- **Telemetry**: Integrated `onWarning` accumulation into the `OfficeParserAST` to preserve parser-phase issues throughout the generation pipeline.
- **Generator API**: Enforced a strict return contract (`string | false | void`) for `onNode` callbacks to ensure deterministic AST transformations.

### Fixed
- **RTF Generator Fidelity**: Restored manual indentation for lists; implemented `\cellx` table layouts and `\pict` binary image embedding.
- **Visualizer Layout**: Resolved `ReferenceError` regressions and cross-zoom layout drift on high-DPI displays.

## [7.0.0] - 2026-05-12
### Added
- **OfficeConverter**: A high-level, streamlined API (`convert`) for one-step document transformations with automatic parser/generator configuration sync.
- **OfficeGenerator**: A comprehensive conversion engine for document ASTs, enabling high-fidelity output in `Markdown`, `HTML`, `CSV`, `RTF`, and `Text`.
- **RAG Chunking Suite**: Native, metadata-aware document splitting optimized for Vector Databases.
    - Supports `fixed-size` (recursive), `document-structure`, and `semantic` strategies.
    - Features robust sentence boundary detection (abbreviations, Japanese punctuation) and deterministic HTML output.
- **Parser Extensions**: Added native support for parsing `CSV`, `HTML`, and `Markdown` files into the unified Office AST.
- **StyleMapper Engine**: A semantic translation layer for preserving document styles across formats.
    - Supports a robust DSL with quoted attributes, commas, and regex-based (`~=`) matching.
    - Introduced **Structured Style Mappings** for type-safe, object-based configuration.
- **Conversion Results API**: Unified `ConversionResult` and `ConversionMessage` interfaces for consistent, structured feedback across all tasks.
- **Standardized Error System**: Introduced `OfficeErrorType` and `OfficeWarningType` enums for predictable and typed error/warning handling.
- **Link Filtering**: Added granular controls `ignoreInternalLinks` to prune noisy document navigation and bookmarks from the AST.

### Changed
- **Unified Office AST**: Redesigned the core document representation to support complex tables, nested lists, and format-specific metadata across all parsers.
- **Performance Optimizations**:
    - **RTF Parser**: Rewritten string accumulation logic to resolve $O(n^2)$ bottlenecks in large documents.
    - **OpenOffice Parser**: Improved XML pre-parsing and style caching, yielding significant speedups (up to 23x for ODP).
    - **Excel Parser**: Replaced global regex matching with `matchAll` iteration to significantly reduce memory overhead and prevent execution stalls on large, sparse spreadsheets (Fixed #91).
- **Browser Build**: Optimized the bundling process to suppress dynamic import warnings in browser environments by injecting ignore comments into dynamic imports.
- **Configuration Engine**: Migrated to a strictly-typed architecture using `DeepRequired` to ensure robust defaults and eliminate runtime configuration errors.
- **CLI Enhancements**: Expanded CLI capabilities with `--format`, `--output`, `--verbose` (for stack traces), and specialized flags for XML serialization.
- **CSV API**: Standardized single-sheet exports to return plain strings for better ergonomics.

### Fixed
- **DOCX Table Fidelity**: Implemented support for vertical cell merging (`w:vMerge`) and horizontal spanning (`w:gridSpan`) in Word documents.
- **Document Anchors**: Added preservation of bookmarks and anchor IDs during Word document parsing.
- **Error Reporting**: Standardized reporting for OCR and chart data extraction failures.
- **Excel Coordinate Indexing**: Resolved a bug where self-closing XML tags caused incorrect row/column metadata indexing and added support for multi-letter column coordinates (e.g., AA, XFD).

## [6.1.1] - 2026-04-28
### Added
- **Break Nodes (DOCX)**: Comprehensive support for `w:br`, `w:cr`, and `w:lastRenderedPageBreak` nodes in Word documents.
- **Indentation Metadata (DOCX)**: Extraction of `<w:ind>` properties for precise paragraph layout analysis.
- **Field Extraction (PPTX)**: Support for `<a:fld>` elements, ensuring slide numbers and other dynamic fields are captured.

### Fixed
- **Soft Break Handling**: Standardized splitting of list items on soft breaks (`Shift+Enter`) across PPTX and ODP, treating interruptions as independent paragraph nodes.
- **List Indexing (ODP)**: Re-engineered stateful index tracking for nested lists in ODP to ensure sequential continuity.
- **Excel Multi-line Parsing**: Resolved failures in XLSX parsing for cells containing complex multi-line content.
- **RTF Encoding**: Implemented robust byte-buffering and character decoding to resolve smart quote and double-quote dropouts.
- **XLSX Fidelity**: Fixed case-sensitivity issues in regex for `inlineStr` cell types.
- **Security & Stability**: Upgraded `@xmldom/xmldom` to `0.9.10` to address upstream vulnerabilities.

### Changed
- **PPTX Engine**: Migrated to an iterative child-processing model for paragraphs to guarantee correct content ordering and support for all inline elements.
- **Documentation**: Updated OpenGraph metadata and project specs for better social sharing and developer clarity.

## [6.1.0] - 2026-04-14
### Added
- **OCR Scheduler**: Intelligent worker pool that optimizes Tesseract lifecycle across parallel requests.
- **Custom Properties**: Support for extracting document metadata across OOXML, ODF, and PDF formats.
- **Sponsorship**: Integrated `funding.json` manifest and GitHub Sponsors support.
- **Governance**: Added `.editorconfig`, `.gitattributes`, and `SUPPORT.md`.

### Changed
- **Core Engine**: Replaced legacy zip extraction with `fflate` for significant performance gains and robust browser/edge compatibility.
- **Module System**: Full native ESM support with `Node16` resolution and verified browser bundles (Vite/Angular compatible).
- **Format Refinements**: Hierarchical PDF coordinate alignment and ODT/RTF list parsing stability.

## [6.0.0] - 2025-12-29
### Added
- **Major Overhaul**: Transitioned from simple text extraction to a rich **Abstract Syntax Tree (AST)** output.
- **Structured Output**: Access hierarchical document structure (paragraphs, headings, tables, lists, etc.).
- **Rich Metadata**: Extracted document properties (author, title, creation date).
- **Enhanced Formatting**: Support for bold, italic, colors, fonts, alignment, etc.
- **Attachment Handling**: Extract images, charts, and embedded files as Base64.
- **OCR Integration**: Optional OCR for images using Tesseract.js.
- **RTF Support**: Added full support for Rich Text Format files.
- **TypeScript**: Full TypeScript support with detailed interfaces and improved type definitions.

### Changed
- **Simplified API**: Transitioned to the unified `parseOffice` for all parsing needs (returns a Promise).

## [5.1.1] - 2024-11-12
### Added
- Added `ArrayBuffer` as a type of file input. 
- Introduced browser bundle generation, exposing the `officeParser` namespace for direct browser usage.

## [5.0.0] - 2024-10-21
### Added
- Replaced `decompress` with `yauzl` for zip extraction. 
- Migrated to in-memory extraction (no longer writing to disk).
- Removed config flags related to extracted files and added flags for CLI execution.

## [4.2.0] - 2024-10-15
### Added
- Fixed race conditions when deleting temp files during parallel execution.
- Resolved errors occurring when multiple executions were made without waiting for the previous one to finish.
- Upgraded project dependencies.

## [4.1.2] - 2024-10-13
### Fixed
- Fixed text parsing from XLSX files containing no shared strings file or using `inlineStr` based strings.

## [4.1.1] - 2024-05-06
### Changed
- Replaced `pdf-parse` with a native `pdf.js` implementation for more robust PDF analysis.
- Added `pdfjs-dist` build as a local library.

## [4.0.5] - 2023-11-25
### Fixed
- Improved error catching during file parsing, specifically post-decompression.
- Fixed parallel parsing issues caused by timestamp-only file naming.

## [4.0.0] - 2023-10-24
### Added
- **Revamped Content Parsing**: Resolved content ordering issues (e.g., table positioning in Word files).
- Added `config` object as an argument for `parseOffice` to set delimiters and other configurations.
- Added initial support for parsing PDF files using the `pdf-parse` library.
- Removed support for individual file parsing functions in favor of a unified approach.

## [3.3.0] - 2023-04-26
### Added
- Added support for file buffers as an argument for `filepath` in `parseOffice` and `parseOfficeAsync`.

## [3.2.0] - 2023-04-07
### Added
- Added comprehensive typings to methods for enhanced TypeScript support.

## [3.1.4] - 2022-12-28
### Added
- Added Command Line Interface (CLI) functionality to use `officeParser` directly from the terminal.

## [3.0.0] - 2022-12-10
### Added
- Resolved memory leak issues and bugs related to Open Document (ODF) parsing.
- Improved global error handling.

## [2.3.0] - 2021-11-21
### Added
- Implemented Promise-based wrappers for existing callback functions.

## [2.2.2] - 2020-06-01
### Added
- Added error handling and configurable `console.log` methods.
- Maintained full backward compatibility.

## [2.1.1] - 2019-06-17
### Added
- Added configuration to change the location for decompressing office files (useful for restricted write access environments).

## [2.0.3] - 2019-04-30
### Fixed
- Fixed case-sensitivity bug for file extensions; capital lettered extensions are now supported.

## [2.0.0] - 2019-04-23
### Added
- Added support for Open Office files (`*.odt`, `*.odp`, `*.ods`) through `parseOffice`.
- Created the dedicated `parseOpenOffice` method.
- Added feature to automatically delete the generated dist folder after function callback.

## [1.3.0] - 2019-04-22
### Added
- Introduced the `parseOffice` method to unify parsing across different extensions.
- Added file extension validations.
- Resolved errors for Excel files lacking drawing elements.

## [1.2.0] - 2019-04-19
### Added
- Added support for `*.xlsx` (Excel) files.

## [1.1.2] - 2019-04-18
### Added
- **Initial Release**: Added support for `*.pptx` and `*.docx` files.

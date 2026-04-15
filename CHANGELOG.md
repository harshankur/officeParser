# Changelog

All notable changes to `officeParser` are documented in this file.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

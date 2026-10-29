# officeParser: Universal Office Document Parser & Generator

A robust, strictly-typed **Node.js and Browser** library for parsing office files into a rich **Abstract Syntax Tree (AST)** and generating high-fidelity output in multiple formats.

**Parses:** [`docx`](https://en.wikipedia.org/wiki/Office_Open_XML) · [`pptx`](https://en.wikipedia.org/wiki/Office_Open_XML) · [`xlsx`](https://en.wikipedia.org/wiki/Office_Open_XML) · [`odt`](https://en.wikipedia.org/wiki/OpenDocument) · [`odp`](https://en.wikipedia.org/wiki/OpenDocument) · [`ods`](https://en.wikipedia.org/wiki/OpenDocument) · [`odg`](https://en.wikipedia.org/wiki/OpenDocument) · [`pdf`](https://en.wikipedia.org/wiki/PDF) · [`rtf`](https://en.wikipedia.org/wiki/Rich_Text_Format) · [`csv`](https://en.wikipedia.org/wiki/Comma-separated_values) · [`md`](https://en.wikipedia.org/wiki/Markdown) · [`html`](https://en.wikipedia.org/wiki/HTML) · [`epub`](https://en.wikipedia.org/wiki/EPUB) · [`tex`](https://en.wikipedia.org/wiki/LaTeX) (LaTeX, including Overleaf project zips)

**Generates:** `DOCX` · `ODT` · `XLSX` · `ODS` · `PPTX` · `ODP` · `LaTeX` · `Markdown` · `HTML` · `CSV` · `RTF` · `PDF` · `EPUB` · `Plain Text` · `RAG Chunks`

[![npm version](https://badge.fury.io/js/officeparser.svg)](https://badge.fury.io/js/officeparser)
[![Total Downloads](https://img.shields.io/npm/dt/officeparser.svg)](https://www.npmjs.com/package/officeparser)
[![Weekly Downloads](https://img.shields.io/npm/dw/officeparser.svg)](https://www.npmjs.com/package/officeparser)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

### 🌟 [Live Interactive AST Visualizer & Documentation](https://harshankur.github.io/officeParser/) 🌟
*Upload any office file in your browser: inspect the AST, tweak config, and preview generated output in real-time.*

- **AST Visualizer**: Inspect the hierarchical node tree, metadata, and raw content
- **Config Configurator**: Tweak options (`ignoreNotes`, `ocr`, `newlineDelimiter`) and see results instantly
- **Debugging**: Identify exactly how nodes are interpreted
- **Format Specs**: Read detailed specs for the AST structure and all config options

---

### 📝 [Changelog](CHANGELOG.md)

---

## What's New in v8

- **Rebuilt PDF text extraction.** PDF is no longer treated as a page of flat lines. Tagged PDFs now yield real `heading` (with correct levels), `table`/`row`/`cell`, `list` and footnote/endnote `note` nodes, and one `paragraph` per paragraph. Untagged PDFs recover the same structure geometrically. Multi-column and float-beside-text pages are read in the correct order (recursive XY-cut), broken and glued words are fixed from inter-fragment spacing, super/subscripts and hyphenated line-breaks are rejoined, rotated text is recovered, and internal links resolve to the target section. `.to('text')` is **layout-faithful by default**, rendering each page as a spatial grid so columns and tables line up like the source. Per-run color/highlight extraction (on by default; set `pdfParserConfig.extractTextColor: false` to skip it) and merged-cell (`colSpan`/`rowSpan`) recovery round it out, and every node carries page geometry (`bounds`).
- **Password-protected documents.** Encrypted PDF, OOXML (`docx`/`xlsx`/`pptx`) and ODF (`odt`/`ods`/`odp`/`odg`) open through one unified `password` / `onPassword` option, across parsing, conversion and templating.
- **Native DOCX & ODT generation**, plus a **native PDF engine** (`pdfConfig.engine: 'native'`, built on `pdf-lib`) that produces real PDF bytes with no headless browser, in Node and the browser alike.
- **Templates / mail-merge** via `OfficeTemplate.render` (fill a DOCX template's `{{placeholders}}`, `{{ if }}` conditions and `{{ for }}` loops, down to one table row per item, values formatted with Word-style pictures, single or batch), and **ODG parsing** (LibreOffice Draw).
- **LaTeX in both directions** (8.1): `.tex` files and Overleaf project zips parse into the same AST as every other format (sections, lists, tables with merged cells, figures, math, footnotes, citations, cross-references, user macros, `beamer` slides), so LaTeX converts to DOCX, ODT, HTML, Markdown and the rest. And `to('tex')` turns any parsed document into LaTeX source that compiles unmodified with pdfLaTeX, XeLaTeX, LuaLaTeX, upLaTeX, pLaTeX and `latex` (the last three through dvipdfmx), presentations included (as `beamer` frames), carrying its images inside the one `.tex` file (or, in bundle mode, as files in a zip). See [LaTeX Support](#latex-support).
- **XLSX, ODS, PPTX and ODP output** (8.3): `to('xlsx')` and `to('ods')` write a workbook (a spreadsheet sheet for sheet, its values typed and shown in their number formats; any other document as its tables, each a sheet, and its text on a sheet of its own), and `to('pptx')` and `to('odp')` write a deck (a presentation slide for slide; a PDF's pages as slides with each paragraph where the page set it; any other document flowed onto slides, a slide for each heading). See [XlsxGeneratorConfig](#xlsxgeneratorconfig) and [PptxGeneratorConfig](#pptxgeneratorconfig).

See the [full changelog](CHANGELOG.md) for the complete list, including breaking changes.

---

## Table of Contents
- [What's New in v8](#whats-new-in-v8)
- [Install](#install-via-npm)
- [Command Line Usage](#command-line-usage)
- [Quick Decision Guide](#quick-decision-guide)
- [Library Usage: Parsing](#library-usage-parsing)
  - [Async/Await](#asyncawait)
  - [Callback (Backward Compat)](#callback-backward-compat)
  - [File Buffers, ArrayBuffers & Blobs](#file-buffers-arraybuffers--blobs)
  - [`ast.to()`: Generate from AST](#astto-generate-from-ast)
  - [`.to('text')`: Plain Text Extraction](#totext-plain-text-extraction)
- [OfficeGenerator](#officegenerator)
- [OfficeConverter: One-Step API](#officeconverter-one-step-api)
- [OfficeTemplate: Mail-Merge / Document Generation](#officetemplate-mail-merge--document-generation)
- [Native RAG Chunking](#native-rag-chunking)
- [The AST Structure](#the-ast-structure)
- [Deep Dive: Document Components](#deep-dive-document-components)
- [Markdown Dialect Support](#markdown-dialect-support)
- [EPUB Support](#epub-support)
- [LaTeX Support](#latex-support)
- [Performance Highlights](#performance-highlights)
- [Advanced AST Usage](#advanced-ast-usage)
- [Configuration Reference](#configuration-reference)
  - [OfficeParserConfig](#officeparserconfig)
  - [GeneratorConfig (Common)](#generatorconfig-common)
  - [onNode Callback](#onnode-callback-advanced-node-manipulation)
  - [styleMap: Semantic Style Mapping](#stylemap-semantic-style-mapping)
  - [HtmlGeneratorConfig](#htmlgeneratorconfig)
  - [MdGeneratorConfig](#mdgeneratorconfig)
  - [PdfGeneratorConfig](#pdfgeneratorconfig)
  - [DocxGeneratorConfig](#docxgeneratorconfig)
  - [OdtGeneratorConfig](#odtgeneratorconfig)
  - [XlsxGeneratorConfig](#xlsxgeneratorconfig)
  - [OdsGeneratorConfig](#odsgeneratorconfig)
  - [PptxGeneratorConfig](#pptxgeneratorconfig)
  - [OdpGeneratorConfig](#odpgeneratorconfig)
  - [TexGeneratorConfig](#texgeneratorconfig)
  - [CsvGeneratorConfig](#csvgeneratorconfig)
  - [TextGeneratorConfig](#textgeneratorconfig)
  - [metadataOverrides](#metadataoverrides)
  - [OfficeConverterConfig](#officeconverterconfig)
  - [ChunkingConfig](#chunkingconfig)
- [OCR Scheduler & Resource Management](#ocr-scheduler--resource-management)
- [Browser Usage](#browser-usage)
- [Troubleshooting & Common Issues](#troubleshooting--common-issues)
- [Known Limitations](#known-limitations)
- [Security & Trust Boundary](#security--trust-boundary)
- [Contributing](#contributing)

---

## Install via npm

```bash
npm i officeparser
```

> [!NOTE]
> Requires Node.js >= 22.13.

---

## Command Line Usage

```bash
# Full AST as JSON (default)
npx officeparser /path/to/file.docx

# Plain text output
npx officeparser /path/to/file.docx --to=text

# Convert DOCX to Markdown and save
npx officeparser report.docx --to=md --output=report.md

# Convert PPTX to HTML with OCR (OCR runs over extracted images, so --extractAttachments is required)
npx officeparser presentation.pptx --to=html --output=preview.html --ocr --extractAttachments

# Convert XLSX to CSV with a custom delimiter
npx officeparser data.xlsx --to=csv --csvDelimiter=";"

# Generate RAG chunks
npx officeparser document.pdf --to=chunks

# Convert DOCX to EPUB (--extractAttachments is required to embed images)
npx officeparser book.docx --extractAttachments --to=epub --output=book.epub

# Convert Markdown (or any source) to a Word document
npx officeparser notes.md --extractAttachments --to=docx --output=notes.docx

# Convert a Word document (or any source) to OpenDocument Text
npx officeparser report.docx --extractAttachments --to=odt --output=report.odt

# Convert a document's tables to a workbook (each table a sheet, its text on a sheet named Text)
npx officeparser report.docx --to=xlsx --output=report.xlsx
npx officeparser report.pdf --to=ods --odsConfig.sheetPerTable=false --output=report.ods

# Convert a PDF or any other source to a presentation (--extractAttachments is required for pictures and charts)
npx officeparser slides.pdf --extractAttachments --to=pptx --output=slides.pptx
npx officeparser slides.pdf --extractAttachments --to=pptx --pptxConfig.backgrounds=page1.png,page2.png --output=slides.pptx

# Convert any source to LaTeX: a .tex file, or a zip of main.tex plus its images
npx officeparser paper.docx --extractAttachments --to=tex --output=paper.tex
npx officeparser paper.docx --extractAttachments --to=tex --texConfig.bundle --output=paper.zip

# Convert LaTeX to Word: a single .tex, or an Overleaf project zip (which brings its \input files and images)
npx officeparser paper.tex --to=docx --output=paper.docx
npx officeparser overleaf-project.zip --extractAttachments --to=docx --output=paper.docx

# Overriding file extension mapping
npx officeparser my_document --fileType=docx --to=json
```

### CLI Syntax
- **Values:** Flags can be passed as `--flag=value` or `--flag value`.
- **Booleans:** Bare flags imply `true` (e.g. `--ocr` is equivalent to `--ocr=true`). Negation flags start with `no-` (e.g. `--no-ocr` is equivalent to `--ocr=false`).
- **Nested Objects:** You can pass nested properties directly using JSON dot-notation (e.g. `--ocrConfig.language=fra` or `--htmlConfig.containerWidth=900px`).
- **Images:** the CLI parses directly, so add `--extractAttachments` for images to reach *any* output (HTML/EPUB embed them, DOCX/ODT/PPTX/ODP/Markdown/native-PDF include them, and LaTeX carries PNG and JPEG images inside the `.tex`, or with `--texConfig.bundle` packages them beside it). Without it, an image node has no bytes and HTML/Markdown emit a name-only `<img src="image1.png">` reference. (The `OfficeConverter`/`convert()` API auto-enables this; the CLI does not.)

### CLI Options

| Flag | Values | Default | Description |
|------|--------|---------|-------------|
| `--to` | `json\|text\|md\|html\|csv\|rtf\|pdf\|docx\|odt\|xlsx\|ods\|pptx\|odp\|tex\|epub\|chunks` | `json` | Output format (`latex` is accepted as an alias of `tex`) |
| `--output` | path | (none) | Write output to a file |
| `--fileType` | `docx\|xlsx\|pptx\|odt\|odp\|ods\|odg\|pdf\|rtf\|csv\|md\|html\|epub\|tex` | (none) | Explicitly override input file type detection. Also accepts `latex`/`ltx`, the ODF template names `ott`/`ots`/`otp`/`otg`, and `zip` (parsed as whatever the archive holds) |
| `--ocr` | boolean | `false` | Enable OCR for images (also requires `--extractAttachments`; OCR runs over extracted images) |
| `--ocrConfig.language` | string | `eng` | Tesseract language(s), e.g. `deu` or `eng+fra` |
| `--ocrConfig.preserveLayout` | boolean | `true` | Keep the line layout of recognized text |
| `--password` | string | (none) | Password for an encrypted document (PDF, OOXML, or ODF) |
| `--allowRemoteAssets` | boolean | `true` | `false` fails with `REMOTE_ASSET_BLOCKED` where the pdf.js worker or Tesseract's files would come from the library's default CDN address, and fetches no pdf.js CMap or image decoder from it (`PDF_ASSET_UNAVAILABLE`) |
| `--pdfCMapUrl` | string | pdfjs-dist's `cmaps/` | Folder of pdf.js's CMaps, for the text of fonts a PDF does not embed (CJK) |
| `--pdfDecoderUrl` | string | pdfjs-dist's `wasm/` | Folder of pdf.js's image decoders (JBIG2, CCITT fax, JPEG 2000) |
| `--ocrConfig.cacheMethod` | `write\|readOnly\|refresh\|none` | `write` | Where OCR keeps the language data it loads (`none` stores nothing) |
| `--extractAttachments` | boolean | `false` | Extract images/charts as Base64 |
| `--ignoreNotes` | boolean | `false` | Ignore footnotes/endnotes/speaker notes |
| `--ignoreComments` | boolean | `false` | Ignore inline comments |
| `--ignoreHeadersAndFooters` | boolean | `false` | Ignore headers and footers |
| `--ignoreSlideMasters` | boolean | `false` | Ignore slide masters |
| `--ignoreInternalLinks` | boolean | `false` | Ignore internal links |
| `--newlineDelimiter` | string | `\n` | Delimiter between lines/blocks in plaintext outputs |
| `--csvDelimiter` | string | `auto` | Delimiter of CSV files; `auto` reads it from the file (comma, semicolon, tab or bar) |
| `--includeRawContent` | boolean | `false` | Include raw XML/RTF in nodes |
| `--serializeRawContent` | boolean | `true` | Include stringified XML in metadata |
| `--preserveXmlWhitespace` | boolean | `false` | Keep raw formatting space |
| `--includeBreakNodes` | boolean | `false` | Include layout break nodes (DOCX, ODF and LaTeX page and column breaks; a typed line break is always kept) |
| `--ignorePageGeometry` | boolean | `false` | Omit per-node bounding boxes and page dimensions |
| `--pdfParserConfig.useTags` | boolean | `true` | Use the PDF tag tree; `false` forces geometry-only structure |
| `--pdfParserConfig.detectColumns` | boolean | `true` | Multi-column reading-order detection |
| `--pdfParserConfig.pageRange` | string | all | Parse only the given pages, e.g. `1-3,7` |
| `--htmlParserConfig.preserveComments` | boolean | `false` | Keep HTML/EPUB `<!-- -->` comments as `comment` nodes |
| `--texParserConfig.today` | string | the date of the parse | What `\today` prints in LaTeX input, e.g. `--texParserConfig.today="May 1, 2024"` |
| `--pdfParserConfig.headingDetection` | `auto\|font-size\|tags\|off` | `auto` | How PDF headings are decided: from tags and a font-size heuristic, the heuristic alone, tags alone, or never |
| `--pdfParserConfig.mergeHyphenatedWords` | boolean | `true` | Rejoin words hyphenated across line breaks |
| `--pdfParserConfig.normalizeText` | boolean | `true` | Unicode/ligature normalization of extracted text |
| `--pdfParserConfig.extractTextColor` | boolean | `true` | Record each run's fill colour in `formatting.color`, and the coloured rectangle it sits on in `formatting.backgroundColor` (set `false` to skip for speed) |
| `--pdfParserConfig.keepHeadersAndFootersInPages` | boolean | `false` | Keep a PDF's running headers and footers (a tagged PDF's artifacts, or the lines an untagged one repeats at the top and bottom of its pages) in each page's content, in place, instead of `auxiliary` |
| `--pdfParserConfig.maxTextItems` | number | `20000` | Base of the text items a PDF may yield (plus one per byte of the file) |
| `--pdfParserConfig.maxOperators` | number | `250000` | Base of the drawing operators kept from a PDF (plus four per byte of the file) |
| `--pdfParserConfig.maxAnnotations` | number | `10000` | Base of the annotations read from a PDF (plus one per 32 bytes of the file) |
| `--pdfParserConfig.maxTimeMs` | number | `5000` | Base of the CPU time the separate pdf.js process may spend on a PDF (plus 20 ms per KB of the file) |
| `--pdfParserConfig.maxImagePixels` | number | `40000000` | Most pixels a PDF picture may have to be extracted |
| `--pdfParserConfig.separateProcess` | boolean | `true` | Run pdf.js in a separate process under a memory limit (Node) |
| `--pdfParserConfig.processMemoryMb` | number | `1024` | Heap the separate pdf.js process may use |
| `--verbose` | boolean | `false` | Show full error stack traces and warning logs |
| `--includeFormatting` | boolean | `true` | Include formatting style map matching |
| `--renderMetadata` | boolean | `false` | Render metadata as visible content in the generated output |
| `--includeImages` | `image-only\|image+ocr-text\|ocr-text-only\|none` | `image-only` | How image nodes render. Works as `--includeImages=<mode>` or `--includeImages <mode>`; a bare `--includeImages` means `image-only` |
| `--maxInlineImageBytes` | number | `1500000` | Largest image HTML/Markdown inlines as a `data:` URI (`0` never inlines) |
| `--htmlConfig.containerWidth` | string \| number | `auto` | HTML output container width (e.g. `900px`, `100%`) |
| `--textConfig.pageSeparator` | string | `\n` | Separator written between pages in text output |
| `--pdfConfig.engine` | `html\|native` | `html` | PDF engine: Puppeteer (`html`) or pdf-lib (`native`, no browser) |
| `--pdfConfig.fonts.regular` | file path | (none) | Font file (`.ttf`/`.otf`/`.woff`/`.woff2`) the native engine draws in; also `.bold`, `.italic`, `.boldItalic`, `.mono` |
| `--pdfConfig.fallbackFonts` | file paths, comma-separated | (none) | Font files tried in order for the characters the fonts before them lack (native engine) |
| `--pdfConfig.subsetFonts` | boolean | `true` | `false` embeds whole font files, not only the glyphs used |
| `--pdfConfig.outline` | boolean | `false` | Write a bookmark outline from the headings |
| `--pdfConfig.format` | `auto\|A4\|Letter\|...` | `auto` | Paper format; `auto` is the document's own pages (each section's, each sheet's, a slide's size), else A4 |
| `--pdfConfig.printHidden` | boolean | `false` | Print hidden sheets, rows and columns |
| `--pdfConfig.printSpeakerNotes` | boolean | `false` | Print each slide's speaker notes after it |
| `--xlsxConfig.sheetPerTable` | boolean | `true` | XLSX: each of a document's tables on a sheet of its own; `false` puts them all on one sheet. `--odsConfig.sheetPerTable` for ODS |
| `--xlsxConfig.includeText` | boolean | `true` | XLSX: a document's text outside its tables on a sheet named `Text`; `false` leaves it out. `--odsConfig.includeText` for ODS |
| `--pptxConfig.backgrounds` | picture file paths, comma-separated | (none) | PPTX: a PNG or JPEG file for each page of a PDF (or slide of a presentation), in order, drawn as its slide's background under its text; an empty entry for a page without one (`a.png,,c.png`). A file that cannot be read ends the run. `--odpConfig.backgrounds` for ODP |
| `--texConfig.bundle` | boolean | `false` | LaTeX: write a zip of `main.tex` plus its `images/` instead of the `.tex` alone |
| `--texConfig.embedImages` | boolean | `true` | LaTeX: `false` refers to `images/` files instead of carrying PNG and JPEG images inside the `.tex` |
| `--texConfig.documentClass` | `auto\|article\|report\|book\|beamer` | `auto` | LaTeX document class (`auto` = `beamer` for presentations, `article` otherwise) |
| `--texConfig.standalone` | boolean | `true` | LaTeX: `false` writes the body only, for pasting into an existing document |
| ~~`--format`~~ | `json\|text\|md\|html\|csv\|rtf\|pdf\|docx\|odt\|xlsx\|ods\|pptx\|odp\|tex\|epub\|chunks` | `json` | **Deprecated.** Use `--to` |
| ~~`--toText`~~ | | | **Removed in v8.** Use `--to=text`. |
| ~~`--ocrLanguage`~~ | | | **Removed in v8.** Use `--ocrConfig.language`. |
| ~~`--putNotesAtLast`~~ | | | **Removed in v8.** Notes are attached structurally via `node.notes`. |
| ~~`--outputErrorToConsole`~~ | | | **Removed in v8.** Use `--verbose`. |

Every removed flag above exits with status 1 and prints its replacement, rather than being accepted
and ignored. An unrecognized or renamed **config** key (say `--ocrConfig.autoTerminateTimeout`) is not fatal, but the
CLI always prints the warning naming its replacement, with or without `--verbose`.

---

## Quick Decision Guide

| Goal | API to use |
|------|-----------|
| Extract text / AST from a file | `OfficeParser.parseOffice(file)` |
| Convert directly to another format | `OfficeConverter.convert(file, 'md')` |
| Parse first, then generate | `parseOffice()` → `OfficeGenerator.generate(ast, 'html')` |
| Convert on the AST itself (shorthand) | `ast.to('md')` |
| RAG pipeline chunking | `OfficeConverter.convert(file, 'chunks', {...})` |

---

## Library Usage: Parsing

### Async/Await

```js
const officeParser = require('officeparser');

const ast = await officeParser.parseOffice('/path/to/file.docx');

console.log(ast.type);       // 'docx'
console.log(ast.metadata);   // { author, title, created, ... }
console.log(ast.content);    // Array of hierarchical nodes
console.log(ast.attachments);// Images/charts (if extractAttachments: true)
console.log(ast.warnings);   // Non-fatal issues from parsing phase
```

**TypeScript (named import):**
```ts
import { OfficeParser } from 'officeparser';

const ast = await OfficeParser.parseOffice('report.docx', {
    extractAttachments: true,
    ocr: true,
});
```

### Callback (Backward Compat)

```js
officeParser.parseOffice('/path/to/file.docx', async function(ast, err) {
    if (err) { console.error(err); return; }
    console.log((await ast.to('text')).value);
});
```

### File Buffers, ArrayBuffers & Blobs

Pass a `Buffer`, `ArrayBuffer`, `Uint8Array`, or a web `Blob`/`File` instead of a file path:

```js
const fs = require('fs');
const buffer = fs.readFileSync('/path/to/file.pdf');
const ast = await officeParser.parseOffice(buffer);
```

In the browser you can hand a `File`/`Blob` straight from an `<input type="file">`, with no need to
read it into a buffer first. A `File`'s name drives type detection, so no `fileType` hint is
needed when the name has a recognizable extension:

```js
// input.files[0] is a File (e.g. "report.docx")
const ast = await officeParser.parseOffice(input.files[0]);
```

> [!IMPORTANT]
> **Text-based formats from buffers need a `fileType` hint.**
> Formats like `md`, `html`, `csv` and `tex` have no magic bytes, so the parser cannot
> auto-detect them from a buffer. You **must** provide `fileType` in that case:
> ```js
> const ast = await officeParser.parseOffice(markdownBuffer, { fileType: 'md' });
> ```

> [!NOTE]
> **ZIP-backed formats are identified from inside the archive.** DOCX, XLSX, PPTX, ODT, ODS, ODP
> and EPUB are all ZIP files, and telling them apart from the first bytes alone is unreliable for
> archives written by streaming producers or holding very many parts. When the byte signature is
> inconclusive, the archive is opened and the format is read from its own declaration
> (`[Content_Types].xml`, or the `mimetype` entry), so these parse from a buffer without a hint.
> A LaTeX project zip is recognized the same way, by a `.tex` file with a `\documentclass` near the
> archive root, and so is a file named `.zip`: that extension names no format, so the archive's
> contents decide which parser runs.
> Supplying `fileType` remains the fastest and most certain route: it decides which parser runs,
> and for these formats no archive inspection is done at all.

### Cancellation with AbortSignal

You can pass a standard `AbortSignal` (e.g. from an `AbortController`) to cancel an active parse operation. This is especially useful for setting request-level timeouts or canceling long-running parses (like large PDFs with OCR).

```js
const controller = new AbortController();

// Cancel parsing if it takes longer than 5 seconds
setTimeout(() => controller.abort(), 5000);

try {
    const ast = await officeParser.parseOffice('large_scanned_file.pdf', {
        abortSignal: controller.signal,
        ocr: true,
        extractAttachments: true // page-image OCR needs this; ocr alone does nothing
    });
} catch (err) {
    if (err.name === 'AbortError') {
        console.log('Parsing was cancelled.');
    } else {
        console.error('Parsing failed:', err);
    }
}
```

The signal is checked between steps, and it stops work that waits: OCR, PDF pages, reading an
archive. Reading a document's own markup (its XML, HTML, Markdown or LaTeX) is synchronous work, and
a timer cannot fire while it runs, so an abort requested during it takes effect only once it ends.
To bound the time an untrusted document can take, parse it in a worker thread or child process and
end that when your time limit passes; the signal alone cannot. PDFs are the exception in Node: pdf.js
runs in a separate process (`pdfParserConfig.separateProcess`), which the signal ends mid-stream, and
pdf.js in the host stops at its next time slice.

> [!IMPORTANT]
> **AbortError Propagation**
> When parsing is cancelled via `AbortSignal`, the parser rejects with a standard `AbortError` (a `DOMException` or an Error with `name: 'AbortError'`).
> This error is *not* wrapped in standard OfficeParser error types so that you can reliably detect cancellation using `error.name === 'AbortError'`.

> [!NOTE]
> **Cancellation is cooperative**
> The signal is checked between steps: an already-aborted signal rejects before any work, and an abort
> is seen at the next check (between archive reads, pages, OCR jobs, or batches of parsed tokens). A
> step that runs synchronously, such as parsing a `.tex` file or one large XML part, finishes before a
> timer's abort can run; those steps are bounded in size instead.

> [!NOTE]
> **Worker Cleanup on Abort**
> If an OCR job is actively running in the background when the signal is aborted, `officeParser` automatically terminates the Tesseract worker process immediately and removes it from the pool to prevent thread/memory leaks.

### Custom OCR Timeouts

To prevent the parser from hanging indefinitely due to slow network connections (when downloading Tesseract language datasets) or complex image processing, you can configure granular timeouts under `ocrConfig.timeout`.

```js
const ast = await officeParser.parseOffice('scanned_document.pdf', {
    ocr: true,
    extractAttachments: true, // required: OCR of a PDF's page images runs through the attachment path
    ocrConfig: {
        timeout: {
            workerLoad: 30000,    // 30s max to load worker & download language training files
            recognition: 15000,   // 15s max per image text recognition
            autoTerminate: 10000  // 10s of inactivity before terminating idle workers
        }
    }
});
```

> [!TIP]
> **Non-Fatal Timeout Recovery**
> If `workerLoad` or `recognition` timeouts are exceeded, the parser will log a warning in `ast.warnings` and **continue parsing the rest of the document**. The overall promise resolves successfully with the text extracted from the document layers (rather than failing the entire parse).

### OCR Layout Reconstruction

By default (`ocrConfig.preserveLayout: true`) the recognized text keeps its two-dimensional page layout, rebuilt from Tesseract's per-word bounding boxes: a scanned table, form or multi-column page keeps its columns (right-hand text stays on the right, labels and values line up) instead of collapsing to a flat reading-order string. It is the OCR analogue of `textConfig.preserveLayout` for born-digital PDFs. Set it `false` for the plain, linearized text.

```js
const ast = await officeParser.parseOffice('scanned_invoice.pdf', {
    ocr: true,
    extractAttachments: true, // required: OCR of a PDF's page images runs through the attachment path
    ocrConfig: { preserveLayout: true } // default; false = flat reading-order text
});
```

### `ast.to()`: Generate from AST

The preferred way to convert a parsed AST to another format. Returns a `ConversionResult`.

```ts
// ConversionResult shape:
// { value: string | Uint8Array | OfficeChunk[], messages: OfficeIssue[] }

const { value: markdown, messages } = await ast.to('md');
const { value: html }               = await ast.to('html', { includeFormatting: false });
const { value: chunks }             = await ast.to('chunks', { chunksConfig: { strategy: 'fixed-size', chunkSize: 800 } });
const { value: pdfBytes }           = await ast.to('pdf'); // Uint8Array
const { value: deck }               = await ast.to('pptx'); // Uint8Array (also docx, odt, xlsx, ods, odp, epub)
```

### `.to('text')`: Plain Text Extraction

Plain text comes from `.to('text')`, which is asynchronous and configurable. Its defaults render
tables as aligned grids, lists with markers and indentation, and include notes and image
placeholders:

```js
// Default: aligned tables, list markers, notes, image placeholders, layout-faithful PDF pages
const { value } = await ast.to('text');

// Flat stream of text, no grid alignment or markers
const { value } = await ast.to('text', {
    includeImages: false,
    textConfig: { preserveLayout: false, renderNotes: false },
});
```

| Feature | default | flat (`preserveLayout: false`) | governed by |
|---|---|---|---|
| Tables | aligned grid | one cell per line, tab-separated | `textConfig.preserveLayout` (default `true`) |
| Lists | markers + indentation | plain text | `textConfig.preserveLayout` (default `true`) |
| PDF pages | spatial monospace grid (columns/tables aligned like the page) | flowing text | `textConfig.preserveLayout` + geometry |
| Footnotes/endnotes | emitted | emitted | `textConfig.renderNotes` (default `true`) |
| Image placeholders | emitted | emitted | `includeImages` (default `true`) |

For PDFs with page geometry (the default, unless `ignorePageGeometry` is set), `preserveLayout` renders
each page as a spatial monospace grid so multi-column text and tables line up much like the original
page, similar to `pdftotext -layout`. Use `textConfig.pageSeparator` (default `'\n'`, or `'\f'` for a
form feed) to control what goes between pages.

Spreadsheets (CSV/ODS/XLSX) are unaffected by `preserveLayout`: it governs `table`/`list` nodes,
while spreadsheet content is `sheet`/`row`/`cell`. There the default aligned grid is the most
faithful rendering.

> [!NOTE]
> The synchronous `ast.toText()` method was **removed in v8**. Use `(await ast.to('text')).value`,
> which produces the same content at its defaults and adds the configuration above.

---

## OfficeGenerator

Use `OfficeGenerator.generate(ast, format, config?)` when you need to produce output from an already-parsed AST:

```ts
import { OfficeParser, OfficeGenerator } from 'officeparser';

const ast = await OfficeParser.parseOffice('report.docx');

// Convert to Markdown
const { value: md } = await OfficeGenerator.generate(ast, 'md');

// Convert to HTML with style mapping
const { value: html } = await OfficeGenerator.generate(ast, 'html', {
    includeFormatting: true,
    styleMap: [
        {
            selector: { nodeType: 'paragraph', attributes: { style: 'Heading 1' } },
            output: { tag: 'h1', classes: ['main-title'] }
        }
    ]
});

// Convert to CSV (spreadsheets)
const { value: csv } = await OfficeGenerator.generate(ast, 'csv');
```

**Supported destinations:** `'text'` · `'md'` · `'html'` · `'csv'` · `'rtf'` · `'pdf'` · `'docx'` · `'odt'` · `'xlsx'` · `'ods'` · `'pptx'` · `'odp'` · `'tex'` (alias `'latex'`) · `'epub'` · `'chunks'`

> [!NOTE]
> **PDF generation** uses a headless browser by default (`pdfConfig.engine: 'html'`), which needs the
> optional `puppeteer` peer dependency:
> ```bash
> npm install puppeteer
> ```
> Or choose `pdfConfig.engine: 'native'` to lay the document out directly with `pdf-lib`
> (`npm install pdf-lib`): no browser, and the only engine that produces a real PDF in the browser
> (import from `officeparser/browser-native-pdf` for the client-side path). To draw text the
> Standard-14 fonts cannot (Greek, Cyrillic, Hindi, Arabic, Chinese, ...), give it fonts and
> `npm install @pdf-lib/fontkit`.
> See [PdfGeneratorConfig](#pdfgeneratorconfig).
>
> **EPUB generation with images** requires `extractAttachments: true` on the parse step that
> produced the AST. See [EPUB Support](#epub-support). So do pictures and charts in **PPTX and ODP**
> output: a chart is written as a table of its data, which is in its attachment.

---

## OfficeConverter: One-Step API

`OfficeConverter.convert()` combines parsing and generation in a single call. It automatically syncs parser options from the generator config: unless you set `parseConfig.extractAttachments` explicitly, it is enabled when the output will render images or charts, or when you enable `parseConfig.ocr`. An explicit `parseConfig.extractAttachments` (including `false`) always wins, and `parseConfig.ocr` is honored (so `{ parseConfig: { ocr: true } }` produces OCR text through the converter, given an image-or-OCR output mode).

```ts
import { OfficeConverter } from 'officeparser';

// Minimal usage
const { value: markdown } = await OfficeConverter.convert('report.docx', 'md');

// With config
const { value: html, messages } = await OfficeConverter.convert('data.xlsx', 'html', {
    parseConfig: {
        ignoreNotes: true,
        newlineDelimiter: '\n\n',
    },
    generatorConfig: {
        includeFormatting: true,
        styleMap: [
            {
                selector: { attributes: { style: { value: 'Header', operator: '~=' } } },
                output: { tag: 'h2', classes: ['data-header'] }
            }
        ]
    },
    onWarning: (issue) => console.warn(`[${issue.code}] ${issue.message}`)
});
```

> [!IMPORTANT]
> The `OfficeConverterConfig` shape uses **nested** `parseConfig` and `generatorConfig` sub-objects.
> Do **not** put parser or generator options at the top level; only `onWarning` lives there. An option
> placed there has no effect, and is reported as `UNRECOGNIZED_CONFIG_OPTION` naming where it belongs
> (`texConfig` under `generatorConfig`, `ocr` under `parseConfig`).

---

## OfficeTemplate: Mail-Merge / Document Generation

`OfficeTemplate.render()` (alias `renderTemplate`) fills a **DOCX template**'s tags from your data and returns a new `.docx`: `{{placeholder}}` fields, `{{ if }}` conditions that keep or drop content, and `{{ for }}` loops that repeat it, down to one table row per item. It is not parsing or conversion: the template is copied and only its tags change, so **all of the template's formatting, layout and structure are preserved**. Give it one data object for one document, or an array for a batch (one document per entry, a classic mail-merge). Think of it as a zero-dependency take on Adobe's Document Generation API.

```ts
import { OfficeTemplate } from 'officeparser';
import { writeFileSync } from 'fs';

// One document.
const bytes = await OfficeTemplate.render('invoice-template.docx', {
    data: {
        name: 'Acme Corp', amount: '$1,250.00', due: '2026-10-01', paid: false,
        lines: [
            { description: 'Widgets', qty: 2, price: '$1,000.00' },
            { description: 'Support', qty: 1, price: '$250.00' },
        ],
    },
});
writeFileSync('invoice-acme.docx', bytes); // Uint8Array

// A batch: one .docx per row.
const docs = await OfficeTemplate.render('invoice-template.docx', {
    data: [
        { name: 'Acme Corp', amount: '$1,250.00' },
        { name: 'Globex',    amount: '$980.00'   },
    ],
});
docs.forEach((d, i) => writeFileSync(`invoice-${i}.docx`, d));
```

### Conditions and loops

Type the tags in Word, as ordinary text:

```text
{{ if paid }}
Paid in full. Thank you!
{{ else if total > 1000 }}
Please pay {{ amount }} by {{ due }}. Large invoices may be paid in instalments.
{{ else }}
Please pay {{ amount }} by {{ due }}.
{{ endif }}

{{ for line in lines }}
{{ line.description }} x {{ line.qty }}: {{ line.price }}
{{ endfor }}
```

- **Where the tags sit decides what they cover.** Both tags in one paragraph: the text between them. Tags in separate paragraphs: the paragraphs between them; when every tag of the block has a paragraph of its own, those paragraphs are removed with the tags, so no blank line is left (a tag beside text cuts its paragraph where it stands, and the block's other tag paragraphs then stay, empty). Tags in different cells of one table: whole rows. Put `{{ for line in lines }}` in a row's first cell and `{{ endfor }}` in its last, and the row repeats once per item; close it in a later row and that group of rows repeats. `{{ if }}` keeps or drops rows the same way, each branch starting a row of its own.
- **Conditions** take a field (`{{ if paid }}`) or a comparison: `==`, `!=`, `<`, `<=`, `>`, `>=`, joined with `and`, `or`, `not` and parentheses, against numbers, quoted strings, `true`, `false` and `null`. `false`, `null`, `""`, `0` and an empty list count as false. A comparison takes two values of the same kind (two numbers, two strings, two dates) or an empty one: comparing the text `"150"` with the number `100` rejects with `TEMPLATE_DATA_INVALID` rather than silently picking a branch. Keywords are case-insensitive, and quotes may be straight or curly, as Word's AutoCorrect writes them.
- **Loops** name their item: `{{ line.price }}` in a list of objects, `{{ t }}` in a list of plain values. Inside a loop, `loop.index` (from 0), `loop.number` (from 1), `loop.first`, `loop.last` and `loop.length` describe the current pass, so `{{ for t in tags }}{{ t }}{{ if not loop.last }}, {{ endif }}{{ endfor }}` writes `red, blue`. Loops and conditions nest.
- **Data can be nested.** `{{ customer.address.city }}` follows a path and `{{ items.0.name }}` reads an item of a list. A top-level key that itself holds a dot (`"invoice.total"`) is read as typed before it is read as a path. Only the data's own keys are read, never `__proto__` or `constructor`.
- **Mistakes are typed errors.** A block that is never closed, a closing tag with no block, or tags that cannot enclose content together (one in a table and the other outside it, one inside a tracked change, one in a Word field and the other outside it) reject with `TEMPLATE_SYNTAX`, naming the tag and the part. Text that starts like a block tag but is not one (`{{ If you have questions, call us }}`) stays text, as prose between the delimiters always has; when the blocks around it then fail to pair, the error names it and why it could not be read (`{{ if a = 1 }}`: use `==` to compare). A value that does not fit its tag (a loop over text, a list placed as a field) rejects with `TEMPLATE_DATA_INVALID`.
- **Output Word opens.** Copies of pictures, bookmarks (their names too), paragraphs and content controls get fresh ids, a footnote, endnote or comment in a copy gets a copy of its own, filled in for that pass (its fields read the pass's item), a table left with no row is removed, and a cell a branch emptied keeps an empty paragraph.

### Formatting values

A field can be written through filters, after `|`. `format` takes a picture, as Word's own mail merge does (its `\#` and `\@` switches; Excel's number formats read number pictures the same way, apart from its thousands scaling, a comma after the last digit, which a picture rejects):

```text
{{ total | format "$#,##0.00" }}              $1,250.50
{{ total | format "#,##0.00 '€'" }}           1,250.50 €
{{ loss | format "$#,##0.00;($#,##0.00)" }}   ($320.00)
{{ rate | format "0.0%" }}                    12.5%
{{ due | format "MMMM d, yyyy" }}             November 1, 2026
{{ due | format "dd/MM/yyyy HH:mm" }}         01/11/2026 09:05
{{ name | upper }}                            ADA LOVELACE
```

- **Number pictures** use `0` (a digit, always written), `#` (a digit, written when there is one), `,` (groups of three), `.` (the decimal point) and `%` (a percentage); up to three sections, `positive;negative;zero`; text in `'quotes'` or after `\` is written as is. Values are rounded half away from zero, as Word and Excel round, and text that is a number (`"1250.5"`, as form data carries it) is taken too.
- **Date pictures** use `d` `dd` `ddd` `dddd` (day), `M` `MM` `MMM` `MMMM` (month), `yy` `yyyy`, `h` `hh` (12-hour), `H` `HH` (24-hour), `m` `mm` (minutes), `s` `ss`, and `AM/PM` (`am/pm`, `A/P`). As in Word, `M` is the month and `m` the minute: Excel's `dd/mm/yyyy` is `dd/MM/yyyy` here. A value is a `Date` or ISO 8601 text (`"2026-11-01"`, `"2026-11-01T09:05:00Z"`; without an offset it is UTC).
- **The same on every machine.** The formatting is officeParser's own, not the runtime's `Intl`, so a template renders the same in every Node version and browser. What it writes is yours to set: `numberSeparators` (`{ decimal: ',', group: '.' }` writes `1.250,50`), `dateNames` (month and day names, AM/PM), and `utcOffsetMinutes` (dates are written in UTC unless you move them).
- **`upper` and `lower`** change case; filters chain left to right (`{{ code | lower | upper }}`).
- **Your own filters:** `formatters: { initials: (v) => ..., round: (v, places) => ... }`, used as `{{ name | initials }}` or `{{ x | round 2 }}`. Each gets the value as the data holds it (a list or an object too) and the strings and numbers written after its name.
- A picture given the wrong kind of value (a date picture on a number, a number picture on text that is no number) rejects with `TEMPLATE_DATA_INVALID`; a picture that cannot be read rejects with `TEMPLATE_SYNTAX`. A word after `|` that is no filter leaves the whole tag as text, like other prose between the delimiters. An empty value stays empty, and a missing one follows `onMissing`.

### Fields

- **Run-aware.** Word often splits a typed `{{name}}` across several runs (`{{`, `na`, `me}}`); it is filled anyway, and a value takes the **formatting of the run its placeholder sat in** (a bold `{{amount}}` renders bold). Block tags are found the same way.
- **Everywhere text lives.** Tags in the body, headers, footers, footnotes/endnotes and comments are all filled. Values may contain `\n` (rendered as line breaks).
- **Field names** are Unicode letters and digits plus `_`, `.`, `-` (e.g. `{{invoice.total}}`, `{{customer-name}}`). `if`, `else`, `endif`, `for` and `endfor` are the block tags' own words. Other text between the delimiters, such as prose with spaces, is left as literal text, so it is never mistaken for a field.
- **Deterministic** output (pinned zip timestamps): the same template + data always renders byte-identical bytes.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `data` | `TemplateData \| TemplateData[]` | (required) | Field values, which may be nested objects and lists. One object is one document; an array is one document per entry |
| `delimiters` | `{ start: string; end: string }` | `{{ }}` | Tag delimiters, for fields and block tags alike |
| `numberSeparators` | `{ decimal?, group? }` | `{ decimal: '.', group: ',' }` | What a number picture writes for its point and groups |
| `dateNames` | `TemplateDateNames` | English | Month and day names (full and short) and AM/PM markers a date picture writes |
| `utcOffsetMinutes` | `number` | `0` | Minutes added to UTC before a date picture writes a date (330 for India); a fixed offset, the same on every machine |
| `formatters` | `Record<string, TemplateFormatter>` | (none) | Filters of your own, used after `\|` in a tag |
| `onMissing` | `'keep' \| 'empty' \| 'error'` | `'keep'` | A field the data does not have: leave it as typed (a block whose condition or list uses it is left as typed, tags included), treat it as empty (blank text, a false condition, a loop of nothing), or reject with `TEMPLATE_FIELD_MISSING`. A field present but `null`/`undefined` always renders empty |
| `password` | `string` | (none) | Decrypt the template first, if it is itself password-protected |
| `onPassword` | `(reason) => string \| undefined \| Promise<...>` | (none) | Supplies a password when `password` did not decrypt the template, as the parser's `onPassword` does |
| `maxPasswordAttempts` | `number` | `3` | How many times `onPassword` is asked before the render rejects |
| `decompressionLimits` | `DecompressionLimits` | parser's | Bounds unzipping an untrusted template; `maxXmlElements` (2,000,000) bounds its elements, its tags and the tokens they hold together (`XML_ELEMENT_LIMIT_EXCEEDED`), and `maxRepeatedContent` (16 MiB) what loops and repeated values add, past which the render rejects with `OUTPUT_TOO_LARGE` |

Only DOCX is supported today (other OOXML/ODF formats will follow); a non-DOCX template rejects with `TEMPLATE_UNSUPPORTED_FORMAT`.

---

## Native RAG Chunking

`officeParser` provides native document chunking for Retrieval-Augmented Generation (RAG) pipelines with three strategies:

### Strategy 1: Document Structure (Default)
Splits at natural AST boundaries (paragraphs, headings, pages, slides, sheets). Preserves logical flow.

```ts
const { value: chunks } = await OfficeConverter.convert('report.docx', 'chunks', {
    generatorConfig: {
        chunksConfig: {
            strategy: 'document-structure',
            splitBy: 'heading',    // 'paragraph' | 'heading' | 'page' | 'slide' | 'sheet'
            maxChunkSize: 1500,
            tableSplitStrategy: 'row', // repeats header row in every chunk, ideal for RAG
        }
    }
});
```

### Strategy 2: Fixed-Size (Recursive)
Splits by character count with overlap. Equivalent to LangChain's `RecursiveCharacterTextSplitter`.

```ts
const { value: chunks } = await OfficeConverter.convert('report.docx', 'chunks', {
    generatorConfig: {
        chunksConfig: {
            strategy: 'fixed-size',
            chunkSize: 1000,
            chunkOverlap: 200,
        }
    }
});
console.log(`Generated ${chunks.length} chunks`);
```

### Strategy 3: Semantic
Uses cosine similarity between sentence embeddings to find topic boundaries. Requires you to provide an `embeddingFunction`.

```ts
import OpenAI from 'openai';
const openai = new OpenAI();

const { value: chunks } = await OfficeConverter.convert('report.docx', 'chunks', {
    generatorConfig: {
        chunksConfig: {
            strategy: 'semantic',
            embeddingFunction: async (text) => {
                const res = await openai.embeddings.create({
                    input: text, model: 'text-embedding-3-small'
                });
                return res.data[0].embedding;
            },
            similarityThreshold: 0.8,
            maxChunkSize: 2000,
        }
    }
});
```

### The `OfficeChunk` Object

`generate(ast, 'chunks')` (and `ast.to('chunks')`) resolves to a real `OfficeChunk[]` **array**, not a JSON string - serialize it to JSON/JSONL yourself if your pipeline needs that.

Every chunk contains text and rich metadata for citations and filtered retrieval:

```ts
interface OfficeChunk {
    text: string;
    /** Rich metadata for filtered retrieval */
    metadata: {
        sourceType: string;       // e.g., 'docx', 'pdf'
        pageNumber?: number;      // (PDF only)
        slideNumber?: number;     // (PPTX only)
        sheetName?: string;       // (XLSX only)
        closestHeading?: string;  // Nearest heading above this chunk
        isTableChunk?: boolean;   // True if part of a split table
    };
    startIndex?: number;          // Character offset (if addStartIndex: true)
    endIndex?: number;            // End character offset (if addStartIndex: true)
}
```

---

## The AST Structure

`OfficeParserAST` is a format-agnostic document representation:

```text
OfficeParserAST
├── type: 'docx' | 'pdf' | 'xlsx' | 'csv' | 'md' | 'epub' | 'tex' | ...  (14 formats)
├── metadata: { author, title, created, modified, keywords, customProperties, nativeProperties, styleMap, pageSetup, ... }
├── content: [ OfficeContentNode ]
│   ├── type: 'paragraph' | 'heading' | 'table' | 'list' | 'image' | 'chart' | 'drawing' | 'comment' | 'admonition' | 'embed' | 'definitionList' | ...
│   ├── text: string  (concatenated text of node + all descendants)
│   ├── children: [ OfficeContentNode ]  (recursive structural children)
│   ├── notes: [ OfficeContentNode ]     (footnotes/endnotes/slide notes attached to this node; see below)
│   ├── comments: [ OfficeContentNode ] (inline comments attached to this node)
│   ├── formatting: { bold, italic, underline, color, size, font, alignment, ... }
│   ├── bounds?: { x, y, width, height, baseline? }  (place on its page or slide: PDF, PPTX, ODP; baseline: a PDF text run's)
│   └── metadata: { level, listId, marker, row, col, rowSpan, colSpan, backgroundColor, spacing, style, ... }
├── auxiliary?: OfficeAuxiliaryContent   (out-of-band layout elements)
│   ├── headers?: OfficeContentNode[]   ('header' nodes: DOCX parts, PDF top band (tagged artifacts, or lines repeated across an untagged PDF's pages), ODT master pages, LaTeX fancyhdr; metadata { type, section, hidden })
│   ├── footers?: OfficeContentNode[]   ('footer' nodes, likewise)
│   ├── slideMasters?: OfficeContentNode[] (PPTX slide masters)
│   └── outline?: OfficeContentNode[]   (PDF bookmark outline)
├── attachments: [ OfficeAttachment ]  (populated when extractAttachments: true)
│   ├── type: 'image' | 'chart'
│   ├── name: string
│   ├── mimeType: string
│   ├── data: string  (Base64)
│   ├── ocrText?: string  (if ocr: true AND extractAttachments: true)
│   └── chartData?: { title, dataSets, labels }
├── warnings: OfficeIssue[]  (non-fatal issues from the parsing phase)
├── config: OfficeParserConfig  (the resolved parse config; `.to()` inherits newlineDelimiter/onWarning from it)
└── to(format, config?)  (format: 'html'|'md'|'text'|'csv'|'rtf'|'pdf'|'docx'|'odt'|'xlsx'|'ods'|'pptx'|'odp'|'tex'|'epub'|'chunks', returns { value, messages })
```

A note or comment referred to from several places is one node that each reference's `notes` (or `comments`) array holds. Test for a note you have already seen by identity before handling it again, and do not mutate it expecting only one reference to change. (`JSON.stringify` writes it at every reference.)

### `OfficeIssue`: Warning / Error Object

All warnings and errors (from both parsing and generation) use this shape:

```ts
interface OfficeIssue {
    type: 'warning' | 'info' | 'error';
    code: OfficeWarningType | OfficeErrorType;  // typed enum, e.g. 'OCR_FAILED'
    message: string;
    node?: OfficeContentNode;  // the node that triggered the issue, if any
    details?: any;             // original error or extra context
    page?: number;             // the PDF page it is on (1-based), where the reader knows it
}
```

Thrown errors carry the same object on `error.officeIssue`, so a failed parse is identified by
the same stable `code` you would branch on for a warning, rather than by matching message text:

```js
try {
    const ast = await officeParser.parseOffice(buffer, { fileType: 'docx' });
} catch (err) {
    switch (err.officeIssue?.code) {
        case 'ZIP_NO_ENTRIES_FOUND':  // not a ZIP archive at all
        case 'ZIP_TRUNCATED':         // cut off in transfer, entries incomplete
        case 'REQUIRED_PART_MISSING': // readable ZIP, but not the format it claims
            console.error('Unusable file:', err.officeIssue.message);
            break;
        default:
            throw err;
    }
}
```

> [!IMPORTANT]
> **A corrupt file throws; it does not parse as an empty document.** If an archive is not
> readable, is truncated, or is missing the part its format requires (`word/document.xml`,
> `xl/workbook.xml`, `ppt/presentation.xml`, ODF `content.xml`, the EPUB OPF), parsing rejects
> with one of the codes above. An empty result therefore means the document really is empty.
> Files that are legitimately empty still parse, and say so through `onWarning` /
> `ast.warnings` (`NO_WORKSHEETS_FOUND` for a chartsheet-only workbook, `NO_SLIDES_FOUND` for a
> presentation with no slides).

#### Warning codes (`type: 'warning' | 'info'`, delivered to `onWarning` and collected in `ast.warnings`)

These never throw; they report a degraded-but-successful outcome you may branch on by `code`.

| Code | Phase | Meaning / what to do |
|---|---|---|
| `OCR_REQUIRES_ATTACHMENTS` | parse | `ocr: true` without `extractAttachments: true`; no OCR ran. Set both. |
| `PDF_NO_TEXT_EXTRACTED` | parse | A PDF yielded ~no text (likely scanned). Set `ocr: true` + `extractAttachments: true`. When a CMap its fonts need could not be loaded (`PDF_ASSET_UNAVAILABLE`), the message says so instead. |
| `PDF_ASSET_UNAVAILABLE` | parse | A file pdf.js needed could not be had: a CMap (text in the fonts using it is missing) or an image decoder (pictures needing it are not decoded). Once per file per document, with the first page that needed it (`page`); the message names the file, why (missing, failed, or blocked by `allowRemoteAssets: false`) and the option to set (`pdfCMapUrl` or `pdfDecoderUrl`). |
| `PDF_STRUCT_TREE_UNRELIABLE` | parse | PDF tag tree absent/incomplete; structure recovered geometrically. |
| `PDF_TEXT_ENCODING_SUSPECT` | parse | A fifth or more of a PDF's characters (of at least 50) are unmappable glyphs (broken ToUnicode); text may be garbage. Consider OCR. |
| `PDF_OUTLINE_TRUNCATED` | parse | Bookmark outline hit the depth/size cap; `ast.auxiliary.outline` is partial. |
| `PDF_WORKER_MISSING` / `PDF_WORKER_FALLBACK` | parse | The pdf.js worker could not be loaded / a fallback was used (set `pdfWorkerSrc`). |
| `REMOTE_ASSET_BLOCKED` | parse (error) | `allowRemoteAssets` is `false` and the parse would have fetched the pdf.js worker or Tesseract's files from the library's default CDN address. The message names the options to set (`pdfWorkerSrc`, `ocrConfig.workerPath`/`corePath`/`langPath`). |
| `NO_WORKSHEETS_FOUND` / `NO_SLIDES_FOUND` | parse | A legitimately empty workbook/presentation. |
| `TABLE_CELL_LIMIT_EXCEEDED` | parse | The cells an ODF or XLSX document yields passed `decompressionLimits.maxTableCells` plus one per byte of the document; the rest were not read. |
| `REPEATED_CONTENT_LIMIT_EXCEEDED` | parse | The document repeated `decompressionLimits.maxRepeatedContent` (plus 16 characters per byte of the document) of content by reference (ODF repeated cells and chart values, XLSX shared strings, style values, link targets, chart text per frame, LaTeX titles per reference); later repeats were not made, shortened or went without the value. |
| `PDF_SEPARATE_PROCESS_UNAVAILABLE` | parse | pdf.js could not start in a separate process (Node), so it runs in the host; a hostile PDF can then exhaust its memory, and pdf.js's time is not bounded (bound it with `abortSignal`). |
| `PDF_CONTENT_LIMIT_EXCEEDED` | parse | A PDF passed one of its content limits (plus an allowance per byte), which the message names: past `pdfParserConfig.maxTextItems` or `maxTimeMs` the rest of it was not read; past `maxOperators` its images, text colours and font styles from that page on were not (its text was). |
| `ALT_CHUNK_NOT_READ` | parse | A DOCX alternative-format chunk (`w:altChunk`) was not read: its part is missing, is of a format other than HTML, MHT, RTF, plain text or DOCX, is a DOCX inside a DOCX chunk, or could not be read (not a ZIP, no document part, nested too deep). The rest of the document is read; a chunk past the document's limits (`maxXmlElements`, `maxUncompressedBytes`) still fails the parse. Saving the document again in Word merges chunks into it. |
| `CONTENT_PART_NOT_READ` | parse | A part of the document was not read. A chapter an EPUB's spine lists: its file is missing from the archive (or has an extension other than `.xhtml`, `.html`, `.htm`, `.xht` or `.xml`), or it is encrypted (DRM, listed in `META-INF/encryption.xml`); a book whose chapters are all encrypted throws `DOCUMENT_DECRYPTION_FAILED` instead. In a PPTX, a part that is not XML and holds none of the slides' text: the presentation's slide list or relationships (the slides are then read in the order of their file numbers), or the relationships of a notes page or a slide master (its links and pictures are then not resolved). In a PDF, an XObject whose text pdf.js could not read (with its `page`). |
| `RAW_CONTENT_LIMIT_EXCEEDED` | parse | With `includeRawContent`, the document's nodes reached `decompressionLimits.maxRawContentLength` of raw content; the remaining nodes carry none. |
| `IMAGE_EXTRACTION_FAILED` / `IMAGE_PROCESSING_FAILED` / `ATTACHMENT_EXTRACTION_FAILED` | parse | An image/attachment could not be extracted or decoded; it was skipped or degraded. A PDF's has the page the picture is on (`page`), and covers the pictures pdf.js leaves out without drawing them: one larger than `pdfParserConfig.maxImagePixels` (the message names it), a small inline picture it cannot decode, one without dimensions, an XObject it cannot read. In the browser these are reported where the library can start pdf.js's worker itself (a policy forbidding `blob:` workers, or a `GlobalWorkerOptions.workerPort` you set, leaves them unreported). |
| `ANNOTATION_EXTRACTION_FAILED` / `CHART_DATA_EXTRACTION_FAILED` | parse | A PDF annotation / a chart's data could not be read. |
| `OCR_FAILED` | parse | OCR ran but failed for an image (see `details`). |
| `LATEX_CONSTRUCT_NOT_INTERPRETED` | parse | The LaTeX input used commands or environments the parser does not interpret (the message names them). Text inside them was kept; drawings such as TikZ pictures were omitted. |
| `LATEX_EXPANSION_LIMIT_REACHED` | parse | A LaTeX document hit a bound on macro expansion, file inclusion or nesting depth (the guard against expansion bombs, include cycles and runaway nesting); macros or files past it were not expanded, and content nested past it is kept as plain text. |
| `LATEX_FILE_NOT_FOUND` | parse | The LaTeX input includes files or images the parser could not read (a `.tex` holds only the files it carries in `filecontents` blocks). Parse the project as a `.zip` to include them; images are kept as path references. |
| `FILE_TYPE_DETECTION_FAILED` / `BUFFER_TYPE_MISMATCH` | parse | Type could not be sniffed / disagreed with the `fileType` hint. |
| `PASSWORD_REQUIRED` / `PASSWORD_INCORRECT` | parse | Encrypted input; supply `password`/`onPassword` (these also throw when parsing cannot continue). |
| `UNRECOGNIZED_CONFIG_OPTION` | config | A config key this version does not know (often a typo or a removed/renamed option); it had no effect. Raised for parser and generator configs, and for a `convert()` option placed at the top level instead of under `parseConfig`/`generatorConfig` (the message says where it belongs). |
| `INVALID_CONFIG_VALUE` | config | A generator option got a value it does not accept (a `documentClass`, paper `format` or margin, `pdfConfig.engine`, `htmlConfig.standalone.styles`, a Markdown dialect preset, a chunking strategy/`splitBy`/`tableSplitStrategy`) or `texParserConfig.today` is not a string; the option's default was used, and the message names the option, the value and what it accepts. |
| `CONTENT_NOT_REPRESENTABLE` | generate | A node type has no faithful form in the target format and was downgraded or omitted (e.g. math/embeds in DOCX/ODT, a table-less document to CSV, pictures and charts in XLSX/ODS, which a cell cannot hold, or a chart without its data in PPTX/ODP). From the native PDF engine, `details.reason` says which, as a value to branch on: `'unencodable-text'` (characters no font given could draw, with `details.count` and a `details.sample` of the first few), `'image-in-table-cell'`, `'image-in-text-line'`, `'chart'`, `'embed'`, `'onnode-string-override'`, each raised once per conversion, and `'svg'`, raised for each SVG picture with parts the engine does not draw (named in `details.feature`). HTML raises it with `details.reason: 'chart'` for a chart written as its data table because there is no chart script to draw it with (`htmlConfig.chartJsSrc: ''`). |
| `METADATA_NOT_REPRESENTABLE` | generate | A metadata field could not be represented in the target format. |
| `TABLE_GRID_LIMIT_EXCEEDED` | generate | The document's tables needed more empty grid positions than one output fills (a million, plus 16 per byte of the document): a table's cells were laid out closer, short rows were not padded to the table's width, or a sparse sheet's empty rows were not written; or plain text stopped lining columns up (16 million spaces, plus 16 per byte). |
| `IMAGE_NOT_INLINED` | generate | An image was referenced by name instead of inlined: it is over `maxInlineImageBytes` (Markdown / fragment HTML), or the pictures inlined in the document reached 128 MB in all (Markdown, HTML of either kind, RTF; a picture is inlined at every place that shows it). |
| `IMAGES_NOT_BUNDLED` | generate | LaTeX output references image files the `.tex` does not carry (with `texConfig.embedImages: false`, an image other than a readable PNG or JPEG, or one past the decoding limits); the message names them. Ship them alongside, or set `texConfig.bundle: true`. It also names images the source referred to only by a relative path, with no image data (a `.tex` without its project, an HTML page's `<img src="pics/a.png">`): supply those at that path yourself, since not even a bundle can contain them. |
| `MATH_WRITTEN_AS_TEXT` | generate | A math expression used an unsafe LaTeX command (file access, shell, redefinition) or was malformed, so LaTeX output shows it as literal text instead of typesetting it. |
| `CITATIONS_NOT_RESOLVED` | generate | LaTeX output cites keys (`\cite{key}`) that have no entry in a bibliography it holds, so LaTeX prints `[?]` for them until one is added (a `\bibliography{file}` with a `.bib` file, or a `thebibliography` list); the message names the keys. |
| `PDF_GENERATION_FAILED` | generate | PDF generation failed (e.g. Puppeteer missing for `engine: 'html'`). |
| `INVALID_STYLE_MAPPING` / `INVALID_STYLE_MAP_TAG` | generate | A `styleMap` entry/tag was invalid and ignored. |
| `TEMPLATE_UNSUPPORTED_FORMAT` / `TEMPLATE_FIELD_MISSING` | template | The template format is unsupported / a `{{field}}` had no value under `onMissing: 'error'`. |
| `TEMPLATE_SYNTAX` | template | Blocks do not pair (a tag that could not be read is named, where that is why), or a block's tags cannot enclose content together; `details` holds the `reason` (`'unreadable'`, `'unclosed'`, `'unopened'`, `'mismatched'`, `'else-after-else'`, `'crossing'`, `'row-branch'`, `'overlap'`, `'malformed-xml'`; see `TemplateSyntaxReason`), the `tag` as typed and the `part`. |
| `TEMPLATE_DATA_INVALID` | template | A value does not fit its tag: a loop over something that is not a list, a list, object, function or symbol placed as a field, a comparison of two kinds of value, or a picture given the wrong kind of value; `details` holds the `reason` (`'not-a-list'`, `'not-text'`, `'incomparable'`, `'unformattable'`; see `TemplateDataReason`), the `tag`, the `part` and the value's `kind`. |
| `PAGE_LOAD_FAILED` | parse | A PDF page could not be processed and was skipped (partial content). |
| `SHEET_RANGE_NOT_FOUND` | generate | A `csvConfig.sheets` range matched no sheet, so CSV output is empty. |
| `EMPTY_CHUNK_GENERATED` / `WHITESPACE_NODE_SKIPPED` / `BROWSER_GENERATION_LIMITATION` / `PERFORMANCE_TIP` / `DEPENDENCY_LOAD_FAILED` | generate | Diagnostic/informational notes from the chunking and PDF generators. |

Generating throws `OUTPUT_TOO_LARGE` when the output would grow past what a string or array holds, or when an AST built in code shares its nodes, records, lists or long strings along more paths than a writer follows (a shared value is written at every use, so a few KB of AST could ask for gigabytes). Parsed documents stay inside the bound: a record many nodes share (a spreadsheet's style, given to every cell) weighs at each later holder only what it weighs past what a node's own record may, and the bound grows with what the AST holds and with the document it was parsed from (its own `config.decompressionLimits.maxRepeatedContent`, up to 1 GiB, plus 16 characters per byte of that document). An AST nested past what the stack holds throws `MAX_NESTING_DEPTH_EXCEEDED`.

The full enum lives in `OfficeWarningType` / `OfficeErrorType` (`src/types.ts`); the error codes used in the `catch` above are the `OfficeErrorType` members.

---

## Per-Format Capability Matrix

What each parser extracts differs by format, because the source formats themselves differ. This table
is the authoritative reference; the option docs point back to it. `Y` = extracted by default (subject to
the relevant `ignore*`/`extractAttachments` flag), `–` = the format has no such construct or it is not
extracted (the matching `ignore*` flag is then a no-op).

| Input | Comments (`node.comments`) | Notes (`node.notes`) | Headers/footers (`ast.auxiliary`) | Slide masters | Images (needs `extractAttachments`) | Charts | Tables (colSpan/rowSpan) |
|---|---|---|---|---|---|---|---|
| DOCX | Y | footnotes/endnotes | Y | – | Y | Y | Y |
| XLSX | Y | – | – | – | Y | Y | grid |
| PPTX | Y | speaker notes | – | Y | Y | Y | Y |
| ODT  | Y (in text) | footnotes/endnotes | Y (master pages) | – | Y | Y | Y |
| ODS  | Y (cell notes) | – | – | – | Y | Y | grid |
| ODP  | Y (page) | speaker notes | – | – (ODP masters not extracted) | Y | Y | Y |
| ODG  | Y (page) | – | – | – | Y | – | Y |
| PDF  | – | footnotes/endnotes (tagged) | Y (top/bottom bands) | – | Y | – | Y (spans: tagged, drawn with ruling lines, or a cell set between the rows it covers) |
| RTF  | Y (annotations) | footnotes/endnotes | Y | – | Y | – | Y |
| HTML | `<!-- -->` become `comment` nodes* (opt-in: `preserveComments`) | footnotes/endnotes | – | – | Y (`data:` only) | – | Y |
| MD   | `<!-- -->` become `comment` nodes* | footnotes/endnotes | – | – | Y (`data:` only) | – | Y (HTML-table fallback) |
| CSV  | `#`-rows become `comment` nodes* | – | – | – | – | – | rows |
| EPUB | `<!-- -->` become `comment` nodes* (opt-in: `preserveComments`) | footnotes/endnotes | – | – | Y | – | Y |
| TEX  | Y (`% Comment (Author, date):` lines); `% <!-- -->` lines become `comment` nodes* | footnotes/endnotes | Y (`fancyhdr`) | – | Y (from a project zip) | – | Y |

Notes: comments land on `node.comments[]` (with `author`/`date`, and a PowerPoint reply's `parentId`) except source-level comments, which
are *not* governed by `ignoreComments`: CSV's leading-`#` rows become top-level `comment` nodes, and
Markdown/HTML `<!-- ... -->` (and LaTeX `% <!-- ... -->` lines) become `comment` nodes (block or
inline) marked `metadata.sourceSyntax: 'html'` whose `text` is the raw comment body. The Markdown,
HTML and LaTeX generators keep those as comments; every other output format omits them (a hidden note
stays hidden). `ignoreNotes` /
`ignoreComments` / `ignoreHeadersAndFooters` / `ignoreSlideMasters` each remove the corresponding
column and are a no-op wherever it shows `–`. OCR (`ocr: true`) recognizes text from any extracted
image and therefore also needs `extractAttachments: true`.

Content kept in parts of its own is read where it stands: SmartArt text (as a nested bulleted list, in
DOCX after the paragraph drawing it), DOCX content controls, text boxes (their paragraphs, lists and
tables, after the paragraph drawing them) and alternative-format chunks (`w:altChunk`: HTML, MHT, RTF,
plain text or a DOCX, in the body, headers, footers, notes and comments), and RTF shape text boxes (as
blocks right after the paragraph the shape is anchored in). Text a tracked
change deleted or moved away (DOCX, ODT) is not read; insertions are.

A PPTX's slides are read in the order its slide list (`p:sldIdLst`) shows them, not the order of their
file names, and `slideNumber` is a slide's place in that order. Each slide has the notes page its own
relationships name. A slide part the list no longer names (a deleted slide an editor left in the
package) is not read. A slide list or presentation relationships that are not XML leave the slides in
the order of their file numbers, and relationships of a notes page or a slide master that are not XML
leave its links and pictures unresolved; each is reported with a `CONTENT_PART_NOT_READ` warning.

---

## Deep Dive: Document Components

### 1. Lists

```text
List Node
├── type: 'list'
├── metadata: {
│       listId: '1',          // items with the same listId belong to one logical list
│       listType: 'ordered' | 'unordered',
│       indentation: 0,       // nesting level (0-based)
│       itemIndex: 0,         // sequential position within the list level
│       paragraphIndentation: { left, hanging, right, firstLine }
│   }
└── children: [ Text content ]
```

> [!TIP]
> Even if a list is interrupted by a regular paragraph, `itemIndex` keeps incrementing for the same `listId`, so numbering stays correct.

### 2. Tables

Tables follow a strict `table → row → cell` hierarchy:

```text
Table Node (type: 'table')
└── children: Row Nodes (type: 'row')
    └── children: Cell Nodes (type: 'cell')
        ├── metadata: { row, col, rowSpan?, colSpan?, backgroundColor? }
        └── children: [ Paragraph | List | Table | ... ]
```

- `row` / `col`: zero-based grid position (a cell without them takes the next place in reading order). Generators fill
  the grid between cells, so the grids of a document's tables and sheets may hold 1,000,000 empty positions in all, plus
  16 per byte of the document parsed (spans count); one holding more than remain (a cell far from the rest) is laid out
  closer when written, with a `TABLE_GRID_LIMIT_EXCEEDED` warning: the rows and columns no cell starts or ends in are left
  out, and if that is not enough, each row's cells follow one another
- `rowSpan` / `colSpan`: merged cells (DOCX, PPTX, XLSX merged ranges, ODF, HTML, Markdown HTML-tables, and PDF tables that are tagged or drawn with ruling lines, or whose merged cell is set between the rows it covers)
- `backgroundColor`: the cell's shading, where the source sets one (a PDF table's cell takes the colour of the filled rectangle drawn last under its middle, tagged or not, and none where that rectangle is white)
- Cells can contain nested tables

> [!NOTE]
> **Header rows.** A header row is flagged on its cells' metadata (`style: 'header'`, or `isHeader`),
> set by the parsers that mark one (DOCX `w:tblHeader`, ODF `table:table-header-rows`, RTF `\trhdr`,
> HTML `<th>`/`<thead>`, the row above a Markdown table's `| --- |` line, tagged-PDF `TH`); generators read it through one shared heuristic (a marked row, or an
> all-bold first row). One format-imposed asymmetry: a **Markdown** table always renders a header row
> (the GFM `| --- |` separator is mandatory syntax), whereas **HTML** emits `<thead>` only for a
> detected header. So a table with no real header prints a header in Markdown output but not in HTML.

### 3. Images & OCR

```text
Image Node (type: 'image')
├── metadata: { attachmentName: 'img1.png', altText: '...' }
└── → Attachment: { data: 'base64...', ocrText: '...' }
```

- Set `extractAttachments: true` to populate `attachment.data`
- Set `ocr: true` (requires `extractAttachments: true`) to populate `ocrText`

### 4. Charts

```text
Chart Node (type: 'chart')
├── metadata: { attachmentName: 'chart1.xml' }
└── → Attachment: { chartData: { title, dataSets, labels } }
```

### 5. Text Formatting

```ts
formatting: {
    bold?: boolean
    italic?: boolean
    underline?: boolean
    strikethrough?: boolean
    color?: string          // '#RRGGBB'
    backgroundColor?: string
    size?: string           // e.g. '12pt'
    font?: string
    subscript?: boolean
    superscript?: boolean
    alignment?: 'left' | 'center' | 'right' | 'justify'
    invisible?: boolean     // drawn invisibly: a PDF's text in rendering mode 3 or 7 (a scan's recognised text)
}
```

> [!NOTE]
> On a **content node**, an absent flag and `false` mean the same thing: the flag is simply not
> applied. On **`ast.metadata.styleMap`**, they differ: an absent flag means the style says nothing
> about that property (so it inherits), while `false` means the style explicitly turns it off
> (ODF's `fo:font-weight="normal"`, DOCX's `<w:b w:val="0"/>`). Code resolving inheritance itself
> must test `=== undefined`, not truthiness, or it will treat "explicitly off" as "unspecified".

### 6. Break Nodes (DOCX, ODF and LaTeX)

When `includeBreakNodes: true`, break elements appear as nodes:

```text
Break Node (type: 'break')
└── metadata: {
        breakType: 'textWrapping' | 'page' | 'column' | 'lastRenderedPage' | 'carriageReturn' | 'thematic' | 'section',
        clear?: 'all' | 'left' | 'none' | 'right',
        section?: number,       // on a break that starts a section of the document: its number (1 for the second)
        pageSetup?: PageSetup   // the pages of the section the break starts, with their columns
    }
```

> [!NOTE]
> **Sections and columns.** A `'page'` break that starts a section (a Word section, a paragraph of an
> ODT style naming another master page, an RTF `\sect`, a PDF page of another size or turned, before its
> `page` node), and a `'section'` break, which starts a section
> on the same page (a Word continuous section, an ODF section set in columns of its own, an RTF section
> of `\sbknone`), carry the new section's
> `pageSetup` (under `ignorePageGeometry`, only its page numbering and first-page headers). Its `columns` are
> `{ count, gap, widths?, separator? }`: the number of columns, the gap between them in points, their
> widths in points where they differ (left to right), and whether a line is drawn between them, read
> from Word's `w:cols`, ODF's `style:columns` (a column's `rel-width` its share of the text width) and
> RTF's `\cols` and `\colw`. `ast.metadata.pageSetup.columns` is the first section's. An ODF section
> in columns has a `'section'` break before it and one after it, which returns to the page's own
> columns. The native PDF engine lays text out in the columns, balancing a section that ends on its
> page; DOCX output writes each section with its page, columns, header and footer, ODT a master page
> for a section on a page of its own (with its header and footer) and a section of the columns, RTF
> each section's page words, columns, header and footer, and LaTeX `multicols` (a long table between
> the columns, across the page); other outputs write nothing for a `'section'` break. A PDF section's
> `pageSetup` is its page's width and height alone, as `ast.metadata.pageSetup` is the first page's: a
> PDF has no margins, so the writers' own are used. HTML writes no break for it (the page before ends
> there already), and under `ignorePageGeometry` a PDF has no sections.

> [!NOTE]
> Break nodes have no `text` property, but `ast.to('text')` automatically converts them to the configured newline delimiter.

> [!NOTE]
> `includeBreakNodes` gates layout breaks in DOCX, ODF and LaTeX only (page and column breaks, otherwise invisible); a
> section's start in a Word, RTF or ODT document, or at a PDF's page of another size, is always a `break` node with `metadata.section`. A
> line break the author typed (DOCX `w:br`/`w:cr`, ODF `text:line-break`, LaTeX `\\`) is content, and is a `break` node whatever the flag says. **HTML and
> Markdown always emit break nodes regardless of the flag**, because a break is explicit content there:
> a `<br>`/hard line break becomes a `carriageReturn` break, and `<hr>`/`---` a `thematic` break (a
> Markdown `\f`-style page break maps to `page`).

> [!NOTE]
> DOCX writes breaks inline (`w:br`/`w:cr`), so they land as children of the paragraph. ODF instead
> carries page and column breaks on the paragraph *style* (`fo:break-before`/`fo:break-after`), so those
> are emitted as siblings around the paragraph rather than inside it. `<text:soft-page-break/>` maps onto
> `lastRenderedPage`, the same type as DOCX's `w:lastRenderedPageBreak`.

### 6b. Equations

Equations are extracted from every format that can carry them and normalized to **LaTeX**, so a
formula means the same thing whichever format it arrived in:

| Source format | Markup in the file |
|---|---|
| DOCX, PPTX | OOXML `<m:oMath>` / `<m:oMathPara>` |
| ODT, ODP, ODS | MathML inside the embedded formula object |
| HTML, EPUB | native MathML `<math>` |
| Markdown | `$inline$` / `$$block$$` |

They all land as the same node:

```text
Code Node (type: 'code')
├── text: '\\frac{1}{2}'          // LaTeX, whatever the source markup was
└── metadata: { math: 'inline' | 'block' }
```

Fractions, sub/superscripts, radicals, delimiters, n-ary operators (sums, integrals), named
functions, accents, bars, matrices and math alphabets (`ℝ`, `𝒜`, …) are all preserved. When a
document supplies its own TeX source in an `<annotation encoding="application/x-tex">`, that is
used verbatim in preference to anything reconstructed from the presentation markup.

> [!NOTE]
> Equation text is *structure*, not prose: a fraction whose numerator and denominator are simply
> concatenated reads as a different number rather than as obviously-missing content. Consumers that
> index document text should treat `code` nodes carrying `math` as opaque LaTeX rather than
> splitting them as words.

**On generation**, an equation's fate depends on the target: LaTeX output typesets it as real math
(`$…$`, `\[…\]`, or a bare `align`-style environment), after a safety check that refuses any command
reaching outside the formula (see [TexGeneratorConfig](#texgeneratorconfig)); HTML and Markdown keep it
as LaTeX (a `$…$`/`$$…$$` delimited block or a `data-math` attribute); DOCX and ODT downgrade it to its
LaTeX text and emit a `CONTENT_NOT_REPRESENTABLE` warning (no native OMML/ODF-math is written); plain
text, RTF and the PDF engines render the LaTeX string as-is without a warning. So the LaTeX always
survives, and LaTeX, HTML and Markdown keep it as math.

### 7. Document Metadata

```ts
ast.metadata = {
    author?: string
    title?: string
    created?: Date
    modified?: Date
    description?: string
    keywords?: string                            // NEW: Keywords from document properties
    customProperties?: Record<string, any>       // User-defined metadata from the document
    nativeProperties?: Record<string, any>       // NEW: All format-specific raw metadata
    pageSetup?: PageSetup                        // The document's page: width, height, margin, headerDistance, footerDistance (points), differentFirstPage, columns, pageNumberFormat
    styleMap?: Record<string, TextFormatting>    // Named styles → formatting definitions
    formatting?: TextFormatting                  // Document-wide defaults
}
```

**Accessing native properties (format-specific metadata):**
```js
const ast = await officeParser.parseOffice('contract.docx');
console.log(ast.metadata.nativeProperties);
// DOCX: { Pages: 5, Application: 'Microsoft Word' }
// HTML: { description: 'My page', 'og:title': 'Title' }
// PDF:  { Title: 'Report', XMP: { ... } }
```

### 8. Admonitions, Embeds & Definition Lists

```text
Admonition Node (type: 'admonition')
├── metadata: { admonitionType: 'note' | 'tip' | 'important' | 'warning' | 'caution', title?: string }
└── children: [ Paragraph | List | ... ]   (block content)

Embed Node (type: 'embed')
└── metadata: { embedType: 'youtube' | 'iframe', videoId?: string, url?: string, width?: string, height?: string, align?: string }

Definition List Node (type: 'definitionList')
└── children:
    ├── Definition Term (type: 'definitionTerm')
    └── Definition Description (type: 'definitionDescription')
```

- `admonition` round-trips through both Markdown (`> [!NOTE]` / `:::note ... :::` / Pandoc's `::: {.note} ... :::`) and HTML (`<div class="admonition admonition-note" data-type="note">`)
- `embed` models YouTube videos and generic iframes. Markdown form is selected by `mdConfig.dialect.embeds`: `'html'` (default; the `<div data-youtube-video>` / `<iframe>` block), `'directive'` (a `::youtube[…]{…}` / `::embed[…]{…}` leaf directive), `'link'`, or `'thumbnail'` (YouTube-only clickable preview). A generic iframe is captured only under `htmlParserConfig.preserveIframes` (the trust input) and can be emitted as an inert click-to-load placeholder via `htmlConfig.gatedEmbeds`. The `'directive'` form is an editor round-trip format, not GitHub-rendered
- Abbreviations (`*[HTML]: Hypertext Markup Language`) are stored as `TextMetadata.abbreviationTitle` on the abbreviated text node rather than as a separate node type

---

## Markdown Dialect Support

Beyond CommonMark/GFM basics, `MarkdownParser`/`MarkdownGenerator` support an extended dialect aimed at
full-fidelity round-tripping with rich Markdown editors. Every construct below parses to a first-class
AST node/metadata field and regenerates back to the canonical syntax shown, so `.md → AST → .md` is
idempotent and `.md → AST → HTML → AST → .md` survives unchanged.

> Markdown-input parsing options that are not dialect toggles live on `htmlParserConfig` (Markdown
> shares the HTML parser for embeds): `preserveIframes` and `embedFolkForms` govern raw `<iframe>`
> blocks and folk embed forms encountered in `.md`. There is no separate `mdParserConfig`.
> (`preserveComments` governs HTML and EPUB input only: comments in Markdown are always kept, see the table below.)

| Feature | Markdown syntax | AST representation |
|---|---|---|
| Task lists (GFM) | `- [x] Done` / `- [ ] Todo` | `ListMetadata.isTask` / `.checked` |
| Admonitions | `> [!NOTE]` (also accepts GLFM `:::note ... :::` and Pandoc `::: {.note} ... :::` on import) | `type: 'admonition'`, `AdmonitionMetadata` |
| Footnotes | `Text[^1]` + `[^1]: Definition` | `type: 'note'`, keyed by footnote id |
| Definition lists | `Term\n: Definition` | `type: 'definitionList'` / `'definitionTerm'` / `'definitionDescription'` |
| Abbreviations | `*[HTML]: Hypertext Markup Language` | `TextMetadata.abbreviationTitle` |
| HTML comments | `<!-- note -->` on its own lines (may span lines, blank ones included) or inline in a run | `type: 'comment'`, `CommentMetadata.sourceSyntax: 'html'`, raw body in `text`; re-emitted byte-for-byte by the Markdown and HTML generators, kept by the LaTeX generator as `% <!-- ... -->` lines (which the LaTeX parser reads back), omitted by every other generator. `<!-->` and `<!--->` are empty comments; write `\<!--` for literal text |
| Attribute lists | `![alt](img.png){width=50% .centered}` | `ImageMetadata.width` / `.align`, `TableMetadata.align` |
| Citations | `[@smith2024]` | `TextMetadata.citationKey` |
| Wikilinks | `[[Page]]` / `[[Page\|Alias]]` | `TextMetadata.wikilink`, `.link`, `.linkType` |
| Highlight | `==text==` | `TextMetadata.backgroundColor` |
| Link/image titles | `[text](url "Title")` / `![alt](img.png "Title")` | `TextMetadata.title` / `ImageMetadata.title` |
| Linked images (badges) | `[![alt](img.png "Title")](url "Link title")` | `ImageMetadata.link` / `.linkType` / `.linkTitle` |
| Inline/block math | `$E=mc^2$` / `` $$...$$ `` | `type: 'code'`, `CodeMetadata.math` (`'inline' \| 'block'`). `$$...$$` is display math wherever it is written: inside a paragraph, the paragraph is split around it; in a heading, list item, table cell, quote or note, where a block cannot go, it is inline math |
| Embeds | `::youtube[Label]{id=… width=… align=…}` / `::embed[Label]{src=… …}` (leaf directive; see `mdConfig.dialect.embeds`) | `type: 'embed'`, `EmbedMetadata` |
| Frontmatter arrays | `tags: [a, b]` or `tags: ["a","b"]` | Real array in `metadata.customProperties`/`nativeProperties` |
| MDX components (import-only) | `<Component prop="x">...</Component>` | Stripped; inner Markdown is kept. Never generated back. |

Text is written so that it reads back as itself, in CommonMark renderers and in the parser alike: a
character that would be markup gets a backslash (`\*`, `` \` ``, `\[`, `\$`, and `\_` except inside a
word), a block marker gets one only where it starts a line (`\# `, `\- `, `1\. `), and the `&` of a
character reference is written `&amp;`. The parser reads emphasis, code spans, links and fences as
CommonMark does: `***text***` is bold and italic, an underscore inside a word is text, link text may
hold brackets and a title parentheses, a hard break is two trailing spaces or a backslash, and a fence
may be indented under a list item.

> [!NOTE]
> MDX/JSX stripping is one-directional (parse-only): officeParser never authors JSX back into Markdown,
> and a component inside code is code, left as written.
> Wikilink enable/disable and citekey→bibliography resolution are application-level concerns; officeParser
> always parses/generates the syntax itself.

The same round-trip fidelity extends to HTML, so content saved from a rich-text editor survives a
save→reload cycle:

| HTML attribute | AST field | Notes |
|---|---|---|
| `data-width` / `data-align` / inline `style="width:…"` on `<img>` | `ImageMetadata.width` / `.align` | |
| `data-align` on `<table>` | `TableMetadata.align` | Where the table stands on the page, not how its columns are aligned. In Markdown it is an attribute list under the table (`{align=right}`), in a dialect that writes one there (`attributeLists: 'brace'`, as `extended` has; the `pandoc` preset writes none, since Pandoc shows that line as text) |
| `text-align` on `<th>`/`<td>` | `CellMetadata.align` | A column's alignment, written as the GFM markers (`:---`, `:---:`, `---:`) of the delimiter row |
| `<th>`, and any cell of a `<thead>` | `CellMetadata.style: 'header'` | A header row stays one, whether or not its text is bold |
| `title` on `<a>` / `<img>` | `TextMetadata.title` / `ImageMetadata.title` | Survives both directions (`[text](url "Title")` in Markdown) |
| `<a href>` around `<img>` | `ImageMetadata.link` / `.linkType` / `.linkTitle` | A linked picture (a badge); written back the same way, and as a linked picture in DOCX, ODT, RTF and LaTeX |
| `colspan` / `rowspan` on `<td>`/`<th>` | `CellMetadata.colSpan` / `.rowSpan` | Previously dropped on HTML import; merged cells now survive a save→reload cycle |
| `<div data-youtube-video="ID">` / `<iframe src="...youtube.com...">` | `type: 'embed'` | |
| `<ul data-type="taskList">` / `<li data-checked data-type="taskItem">` | `ListMetadata.isTask` / `.checked` | The shape Tiptap's task list reads and writes |

---

## EPUB Support

EPUB files are ZIP archives of XHTML content plus an OPF manifest. `EpubParser` unzips the archive,
resolves the spine's reading order from `content.opf`, and parses each XHTML document through the
existing `HtmlParser`, so EPUB content shares the same AST shape (and the same Markdown-dialect
fidelity above) as every other format. Dublin Core metadata (`dc:title`, `dc:creator`, `dc:description`,
`dc:subject`, `dc:date`, `dc:publisher`, `dc:language`, `dc:identifier`) maps into `ast.metadata` /
`ast.metadata.nativeProperties`, and cover art is exposed via `metadata.customProperties.coverImageName`.
A chapter the spine lists that cannot be read (missing from the archive, or encrypted by DRM, as
`META-INF/encryption.xml` lists it) is reported with a `CONTENT_PART_NOT_READ` warning; a book whose
chapters are all encrypted throws `DOCUMENT_DECRYPTION_FAILED`. Font obfuscation is not encryption
of the content, and is passed over. EPUB 3 footnotes (`<a epub:type="noteref">` citing an
`<aside epub:type="footnote">`) are read as notes.

`EpubGenerator` renders the AST through `HtmlGenerator` and packages the result as a minimal, valid
EPUB 3 (`mimetype`, `META-INF/container.xml`, an OPF manifest, a nav document, and one XHTML chapter).

> [!IMPORTANT]
> **Pass `extractAttachments: true` when converting to or from EPUB if the document has images.**
> Without it, the parser never pulls embedded image bytes out of the source document, so there is
> nothing for the EPUB generator to package, so images silently disappear even though everything else
> converts correctly. Images are packaged as real zip entries (`OEBPS/images/...`) declared in the OPF
> manifest, not `data:` URIs, because most EPUB reading systems do not render `data:` URIs in image `src`.
>
> This only matters for the two-step `OfficeParser.parseOffice()` → `OfficeGenerator.generate()` API
> and the CLI. [`OfficeConverter.convert()`](#officeconverter-one-step-api) enables `extractAttachments`
> automatically unless you explicitly set `generatorConfig.includeImages: false`.
>
> ```bash
> npx officeparser book.docx --extractAttachments --to=epub --output=book.epub
> ```

## LaTeX Support

LaTeX is supported in both directions, zero extra dependencies, in Node and the browser.

**Parsing.** A `.tex` file (or `.latex`/`.ltx`) parses into the same AST as every other format, so a
paper converts to DOCX, ODT, HTML, Markdown, EPUB or chunks like any Word document would. So does a
**project zip**, such as Overleaf's "Download source": the main file (a `.tex` no other file includes
and not a `subfiles`/`standalone` part: the shallowest `main.tex`, else the shallowest with a
`\documentclass`) is read, and its `\input`/`\include`/`\subfile` files and
`\includegraphics` images are taken from the archive (respecting `\graphicspath`). Nothing is ever
read from the filesystem, and no path can leave the project.

```js
const ast = await OfficeParser.parseOffice('paper.tex');
const { value: docx } = await ast.to('docx');

// A whole project, images included
const project = await OfficeParser.parseOffice('overleaf-project.zip', { extractAttachments: true });
```

| LaTeX | AST |
|---|---|
| `\section` ... `\subparagraph`, `\chapter`, `\part` | `heading` nodes, `\label`s as heading ids |
| `\textbf`, `\emph`, `\underline`, `\sout`, `\texttt`, `\textcolor`, `\hl`, `\large`, `{\bfseries ...}`, `\MakeUppercase`, accents, ligatures | formatted text runs (small caps are not a run format: `\textsc` keeps its text) |
| `itemize`, `enumerate` (nested, `\setcounter`), `description`, `\item[$\square$]` | lists, definition lists, task items |
| `tabular`, `tabularx`, `longtable`, `\multicolumn`, `\multirow`, `\cellcolor`, booktabs | `table` with `colSpan`/`rowSpan`, alignment and cell colours |
| `figure`/`table` floats, `\caption`, `\includegraphics` | images (bytes from a project zip or a `filecontents` block; a PDF that is only a picture, as the generator carries images, becomes that JPEG or PNG) and captions |
| `filecontents`, `filecontents*` | the file it writes, as compiling does (a file already there is kept unless `overwrite`), for `\input` and `\includegraphics` to read |
| `$...$`, `\[...\]`, `equation`, `align`, `gather` | math, as LaTeX, with your macros (and `\DeclareMathOperator`, `\DeclarePairedDelimiter`) expanded and comments dropped |
| `\footnote`, `\endnote` | `note` nodes |
| `\href`, `\url`, `\ref`, `\eqref`, `\nameref`, `\autoref`, `\cref`, `\hyperref` | links; references resolve to section, table, figure and equation numbers (`\autoref` and `\cref` with their names: "section 1") |
| `\cite`, `thebibliography`, `\bibliography`/`\printbibliography` | citations and the bibliography list: from a project's `.bib` (its cited entries), or its `.bbl` when it ships one; without the file, the cited keys, and the file is reported |
| `verbatim`, `lstlisting`, `minted`, `\verb` | `code` nodes with their language |
| `quote`, `quotation`, `verse` | quotes |
| `% <!-- ... -->` lines | a source `comment` node (the form the generator writes a Markdown/HTML `<!-- -->` in); other `%` comments are dropped |
| `fancyhdr` headers and footers | `ast.auxiliary` |
| `\maketitle`, beamer `\titlepage` | a title block where it stands: a `heading` styled `Title`, then `Author`, `Institute` and `Date` lines (beamer adds `Subtitle`), `\thanks` as footnotes |
| `\title`, `\author`, `\date`, `\hypersetup{pdf...}` | `ast.metadata` (`pdftitle`/`pdfauthor` win over `\title`/`\author`) |
| `beamer` frames, `\framesubtitle`, `\note`, overlays | `slide` nodes with speaker notes |
| `\newcommand`, `\renewcommand`, `\def`, `\newenvironment`, `\NewDocumentCommand`/`\NewDocumentEnvironment` (`m o O s t d D r R v b` arguments, `\IfBooleanTF`, `\IfNoValueTF`) | expanded; `expl3` code blocks are skipped |
| `\ifXeTeX`/`\ifLuaTeX`/`\ifPDFTeX` (iftex), `\newif` switches, `\ifdefined`, `\@ifpackageloaded`, `ifthen` booleans, etoolbox toggles | decided as pdfLaTeX would compile the document: only the branch taken is read (a test that cannot be decided, such as `\ifnum` on a counter, reads both and is reported) |
| `\newtheorem` (shared counters, numbering within sections, `\theoremstyle`), thmtools `\declaretheorem`, `proof` | "**Theorem 2.1** (Note)**.**" before the body (italic in the `plain` style), proofs ending in □; `\ref` reads theorem numbers. Theorem environments a class provides (Springer's llncs, svjour) need no `\newtheorem`, and one defined in a package the parser cannot see is headed but unnumbered; beamer's are titled blocks |
| babel/polyglossia: `\foreignlanguage`, `\text<language>`, `otherlanguage`, language environments, `\babeltags` | their text; the main language (babel's `main=` or last language, polyglossia's `\setdefaultlanguage`) becomes `metadata.language` unless `pdflang` states it |
| `\keywords`, `keywords`/`IEEEkeywords` environments; amsart `\address`, `\email`, `\urladdr`, `\subjclass`; IEEEtran author blocks, `\IEEEPARstart`; KOMA-Script `\minisec`, `\dictum`; `\epigraph` | `metadata.keywords` (and the printed "Keywords:" line), `nativeProperties`, authors' names in `metadata.author`, text |
| `\bf`, `\it`, `\tt`, `\sl`, `\sf`, `\rm`, `\documentstyle` (LaTeX 2.09) | formatted text runs, the document class |
| `inputenc` encodings (`latin1`, `latin9`, `cp1252`, `koi8-r`, ...), a `% !TEX encoding` line | the file decoded in that encoding (an undeclared 8-bit file reads as Windows-1252), included files too |

The parser reads LaTeX (and LaTeX 2.09). Plain TeX's `\bye`, `\beginsection` and glue are understood, and a ConTeXt
document (`\starttext`) has its sections, lists, code and text read, with a `LATEX_CONSTRUCT_NOT_INTERPRETED`
warning that it is not LaTeX.

LaTeX is a programming language, so the parser is a bounded interpreter: macro expansion, expanded
text (included files count toward it), include depth and nesting all have hard limits (past one,
expansion stops with `LATEX_EXPANSION_LIMIT_REACHED`, and include cycles are cut), parsing time grows
linearly with the document, nothing is executed, and zip reading obeys `decompressionLimits`. The
parse itself runs synchronously, so an `abortSignal` is honoured before it starts, while the
project zip is read and while OCR runs, but a timer cannot interrupt a `.tex` being parsed (see
[Cancellation](#cancellation-with-abortsignal)); it is bounded instead. Constructs it does not interpret keep their text and are named once in
`LATEX_CONSTRUCT_NOT_INTERPRETED`; drawings (TikZ) are omitted. Review comments survive a round trip:
a `% Comment (Author, date): text` line, which is how the generator writes a comment, becomes a
comment on the node that follows.

**Generation.** `to('tex')` writes LaTeX from any parsed document; see
[TexGeneratorConfig](#texgeneratorconfig). The two directions agree: LaTeX that officeParser writes
parses back to the same structure, and a generate, parse, generate cycle reaches a fixed point after
one round (the first regeneration may normalize spacing, and image sizes the source never stated).
In bundle mode, an image the source only named by path is written as
`\IfFileExists{path}{\includegraphics{path}}{\fbox{...}}`, so the zip compiles whether or not you add the file.

**Engines.** The generated source compiles unmodified with pdfLaTeX, XeLaTeX, LuaLaTeX, upLaTeX, pLaTeX
and `latex`, TeX Live 2021 and later (pLaTeX before TeX Live 2023 reads only JIS X 0208 characters,
so it stops at Korean text). The preamble tells the engines apart with `iftex`: XeLaTeX and LuaLaTeX load `fontspec`; pdfLaTeX,
upLaTeX and pLaTeX load 8-bit fonts (`fontenc`, `inputenc`, `lmodern`); and in DVI mode (`latex`,
`uplatex` or `platex`, then `dvipdfmx`) every package gets the `dvipdfmx` driver through a class option,
so images, links and colours work there too.

| Characters | pdfLaTeX | XeLaTeX, LuaLaTeX | upLaTeX, pLaTeX |
|---|---|---|---|
| Latin, symbols, typographic spaces | yes (a fallback for each symbol the fonts lack) | yes | yes |
| Greek | as math letters | yes, in Computer Modern Unicode | in the Japanese font |
| Cyrillic | a visible `[U+XXXX]` marker | yes, in Computer Modern Unicode | in the Japanese font |
| Chinese, Japanese, Korean | a visible `[U+XXXX]` marker | yes: Fandol, Harano Aji or UnFonts (whichever language the text is mostly in, the others as fallbacks) through `xeCJK` or `luatexja`, which also break lines between CJK characters | Japanese, and the Chinese its Japanese fonts cover; Hangul needs other fonts, and pLaTeX before TeX Live 2023 cannot read it at all |
| Emoji | a visible `[U+XXXX]` marker | blank unless you add a font that has them | a missing-glyph box |

The fonts are TeX Live's own. The preamble uses each only where it is installed, so a smaller
installation (BasicTeX, TinyTeX) still compiles the document, the characters it has no font for left
blank.

---

## Performance Highlights

Key internal optimizations shipped in recent versions:

- **OpenOffice (ODP)**: Up to **23× faster** parsing via optimized XML pre-parsing and style caching
- **Excel Memory**: Resolved O(n) memory overhead on large sparse spreadsheets using iterative stream-based parsing
- **RTF Parser**: Rewrote string accumulation loop to eliminate O(n²) bottleneck in large files
- **Table Fidelity (DOCX)**: Native support for vertical cell merging (`vMerge`) and horizontal spanning (`gridSpan`)

---

## Advanced AST Usage

### Extract all headings
```js
const headings = ast.content.filter(n => n.type === 'heading' && n.metadata?.level === 1);
console.log(headings.map(h => h.text));
```

### Extract comments
```ts
// Comments can be attached to any nested node, so we must traverse recursively
const printComments = (nodes: OfficeContentNode[]) => {
    nodes.forEach(node => {
        if (node.comments) {
            node.comments.forEach(c => {
                console.log(`Comment by ${c.metadata?.author}: ${c.text}`);
            });
        }
        if (node.children) {
            printComments(node.children);
        }
    });
};

printComments(ast.content);
```

Set `ignoreComments: true` to skip extraction.

### Extract footnotes, endnotes & slide notes
```ts
// Slide speaker notes (PPTX) live on the slide node itself
const slide = ast.content.find(n => n.type === 'slide');
console.log(slide?.notes?.map(n => n.text));

// Footnotes and endnotes (DOCX, ODT, RTF, PDF, HTML, Markdown, EPUB, LaTeX) can be deeply nested, so we traverse recursively:
const printNotes = (nodes: OfficeContentNode[]) => {
    nodes.forEach(node => {
        if (node.notes) {
            node.notes.forEach(note => console.log(note.text));
        }
        if (node.children) {
            printNotes(node.children);
        }
    });
};

printNotes(ast.content);
```

> [!IMPORTANT]
> `putNotesAtLast` was **removed in v8**. Notes are always attached via `node.notes`.

### Access headers, footers & slide masters
```ts
// These are NOT in ast.content; use ast.auxiliary
console.log(ast.auxiliary?.headers?.map(h => h.text));   // DOCX headers
console.log(ast.auxiliary?.footers?.map(f => f.text));   // DOCX footers
console.log(ast.auxiliary?.slideMasters?.length);         // PPTX slide masters
```

Set `ignoreHeadersAndFooters: true` or `ignoreSlideMasters: true` to skip extraction.

### Extract images with OCR text
```js
const ast = await officeParser.parseOffice('report.docx', { extractAttachments: true, ocr: true });
ast.attachments.filter(a => a.mimeType?.startsWith('image/')).forEach(img => {
    console.log(`${img.name}: ${img.ocrText ?? 'no OCR'}`);
});
```

### Extract tables to CSV manually
```js
ast.content.filter(n => n.type === 'table').forEach((table, i) => {
    const csv = table.children
        .filter(r => r.type === 'row')
        .map(r => r.children.filter(c => c.type === 'cell')
            .map(c => `"${c.text.replace(/"/g, '""')}"`)
            .join(','))
        .join('\n');
    console.log(`Table ${i + 1}:\n${csv}`);
});
```

### Find all bold text runs
```js
function findBold(nodes) {
    return nodes.flatMap(n => [
        ...(n.type === 'text' && n.formatting?.bold ? [n.text] : []),
        ...(n.children ? findBold(n.children) : [])
    ]);
}
console.log(findBold(ast.content));
```

### Extract footnotes / endnotes
```js
function extractNotes(nodes) {
    return nodes.flatMap(n => [
        ...(n.type === 'note' ? [{ id: n.metadata.noteId, text: n.text, type: n.metadata.noteType }] : []),
        ...(n.children ? extractNotes(n.children) : [])
    ]);
}
console.log(extractNotes(ast.content));
```

### Search for a term (TypeScript)
```ts
import { OfficeParser } from 'officeparser';

async function contains(filePath: string, term: string): Promise<boolean> {
    const ast = await OfficeParser.parseOffice(filePath);
    return (await ast.to('text')).value.includes(term);
}
```

---

## Configuration Reference

### OfficeParserConfig

Pass as the second argument to `parseOffice(file, config)`.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `newlineDelimiter` | `string` | `'\n'` | Joins multi-line text inside the AST's pre-flattened `.text` (RTF table cells, chart text, PDF page text); also the default for `textConfig.newlineDelimiter` in `.to('text')` when that is not set explicitly. Not read by the Word parser |
| `password` | `string` | `''` | Password for a password-protected document. Applies to every encryptable format: PDF, encrypted OOXML (`.docx`/`.xlsx`/`.pptx`, ECMA-376 agile or standard AES), and encrypted ODF (`.odt`/`.ods`/`.odp`/`.odg`, AES-CBC with PBKDF2). A missing password rejects with `PASSWORD_REQUIRED`, a wrong one with `PASSWORD_INCORRECT`. Ignored for unencrypted files. *ODF note:* LibreOffice 24.8+ defaults to AES-256-GCM with Argon2id key derivation ("wholesome encryption"), which is not supported and rejects with `DOCUMENT_DECRYPTION_FAILED`; re-save with the classic AES-CBC/PBKDF2 scheme (or an earlier LibreOffice) to parse it |
| `onPassword` | `(reason: 'required' \| 'incorrect') => string \| undefined \| Promise<...>` | (none) | Called when an encrypted document needs a password `password` did not satisfy, so it can be supplied lazily or interactively (prompt, vault). Return a password to retry (asked at most `maxPasswordAttempts` times), or `undefined` to reject as above. Works for every encryptable format (PDF/OOXML/ODF); mirrors pdf.js's `onPassword` |
| `maxPasswordAttempts` | `number` | `3` | How many times `onPassword` is asked for one document before the parse rejects. A whole number of 0 or more |
| `allowRemoteAssets` | `boolean` | `true` | Whether the library may fall back to its built-in CDN addresses for the pdf.js worker (`pdfWorkerSrc`), pdf.js's CMaps and image decoders (`pdfCMapUrl`, `pdfDecoderUrl`) and Tesseract's worker, core and language data (`ocrConfig.workerPath`/`corePath`/`langPath`). `false` makes a parse that would use the worker's or Tesseract's default reject with `REMOTE_ASSET_BLOCKED`, naming the option to set, for an application that must make no request to another site; a CMap or decoder a document needs is then not fetched from its default, and the parse goes on without it with `PDF_ASSET_UNAVAILABLE`. An address you pass yourself is used as given. Output written from such a parse names no chart script address unless `htmlConfig.chartJsSrc` is set |
| `ignoreNotes` | `boolean` | `false` | Ignore footnotes/endnotes (DOCX, ODT, RTF, PDF, HTML, Markdown, EPUB, LaTeX) and speaker notes (PPTX/ODP, LaTeX `beamer` `\note`). See the [capability matrix](#per-format-capability-matrix) |
| `ignoreComments` | `boolean` | `false` | Ignore comments/annotations, attached by default via `node.comments[]`. Applies to DOCX, XLSX, PPTX, every ODF type (ODT/ODS/ODP/ODG), RTF (`\annotation`) and LaTeX (`% Comment (Author, date): text` lines). Source-level comments (CSV `#` rows, Markdown/HTML `<!-- -->`, LaTeX `% <!-- -->` lines) are not governed by it. See the [capability matrix](#per-format-capability-matrix) |
| `ignoreHeadersAndFooters` | `boolean` | `false` | Skip headers & footers (populated in `ast.auxiliary.headers/footers` by default). Extracted for DOCX, PDF, ODT, RTF and LaTeX (`fancyhdr`) only; a no-op for ODS/ODP/ODG, XLSX and PPTX. See the [capability matrix](#per-format-capability-matrix) |
| `ignoreSlideMasters` | `boolean` | `false` | Skip PPTX slide masters (populated in `ast.auxiliary.slideMasters` by default). PPTX only; ODP masters are not extracted |
| `extractAttachments` | `boolean` | `false` | Populate `ast.attachments` with Base64 images/charts |
| `ocr` | `boolean` | `false` | Run Tesseract OCR on images (requires `extractAttachments: true`) |
| `ocrConfig` | `OcrConfig` | see below | OCR settings (populated defaults: `language: 'eng'`, `preserveLayout: true`, worker/timeout defaults). See the [OCR section](#ocr-scheduler--resource-management) |
| `includeRawContent` | `boolean` | `false` | Attach raw XML/RTF source to each node |
| `serializeRawContent` | `boolean` | `true` | Re-serialize XML to clean strings (only if `includeRawContent: true`) |
| `preserveXmlWhitespace` | `boolean` | `false` | Preserve original XML whitespace during serialization |
| `includeBreakNodes` | `boolean` | `false` | Include layout break nodes: DOCX page, column and last-rendered-page breaks, ODF `fo:break-before`/`fo:break-after` and `text:soft-page-break`, LaTeX `\newpage`/`\clearpage`. The break that starts a Word, RTF or ODT section, or a PDF's page of another size (with `metadata.section` and its `pageSetup`), and the `'section'` breaks around an ODF section in columns of its own, are the document's structure and always nodes. A line break the author typed (DOCX `w:br`/`w:cr`, ODF `text:line-break`, a LaTeX `\\` or `\newline`, HTML `<br>`) is always a `break` node inside its paragraph, and a DOCX tab is always a `\t` in the text |
| `ignoreInternalLinks` | `boolean` | `false` | Strip bookmarks and internal cross-references from AST (now honored for PDF too) |
| `ignorePageGeometry` | `boolean` | `false` | Omit the geometric layout data: per-node bounding boxes (`node.bounds`: PDF pages, PPTX and ODP slides), page dimensions and page setups (`metadata.pageSetup`, a section's or sheet's `pageSetup`; a text document's keeps how each section numbers its pages and whether its first page has a header of its own), slides' `drawing` nodes and backgrounds, and a floating frame's place (`ShapeMetadata.float`: its text is read where it stands) |
| `fileType` | `SupportedFileType \| FileTypeAlias \| null` | `null` | **Required for text-based binary data** (`'md'`, `'html'`, `'csv'`, `'tex'`) as these lack magic bytes. Also accepts the names the matching extensions route by (`FileTypeAlias`): `'latex'`/`'ltx'` for `tex`, the ODF template names `'ott'`/`'ots'`/`'otp'`/`'otg'`, and `'zip'` (parsed as whatever the archive holds). |
| `csvDelimiter` | `string` | `'auto'` | Delimiter of CSV input: `'auto'` reads it from the file (comma, semicolon, tab or bar, counted outside quotes); any other value is used as given. CSV output writes the delimiter read |
| `decompressionLimits` | `DecompressionLimits` | `{ maxUncompressedBytes: 512MB, maxZipEntries: 10000, maxTableCells: 1000000, maxXmlElements: 2000000, maxRepeatedContent: 16777216, maxRawContentLength: 67108864 }` | Limits applied during ZIP extraction (and to the cells an ODF or XLSX document yields, the content a document repeats by reference, the `rawContent` nodes carry, and the XML elements a document's parts may hold; past `maxXmlElements` the parse rejects with `XML_ELEMENT_LIMIT_EXCEEDED`) to protect against excessive memory and resource usage. `maxTableCells` and `maxRepeatedContent` grow by one cell and 16 characters per byte of the document, so a large document is not cut short for its size while a small crafted one stays bounded |
| `htmlParserConfig` | `HtmlParserConfig` | `{}` | HTML/XHTML/EPUB parsing options **(and Markdown input: `preserveIframes`/`embedFolkForms` govern raw `<iframe>` blocks and folk embeds in `.md` too)**. `preserveAttributes` (`boolean`, default `false`): keep generic source attributes no typed field consumed on `node.htmlAttributes`. `preserveIframes` (`boolean \| string[]`, default `false`): preserve non-YouTube `<iframe>` embeds (otherwise dropped) as `embed` nodes: `true` for any, or a hostname allowlist; the src is scheme-checked on generation. `embedFolkForms` (`boolean`, default `false`): opt in to importing ambiguous folk embed forms (Obsidian `![](youtube-url)`, thumbnail-link) as YouTube embeds. `preserveComments` (`boolean`, default `false`): keep `<!-- ... -->` comments in HTML and EPUB input as `comment` nodes (`metadata.sourceSyntax: 'html'`) instead of dropping them; conditional comments (`<!--[if …]>`) are always dropped. The `data-html-comment` shape `sourceAttributes` emits is always read. |
| `pdfWorkerSrc` | `string` | CDN (jsDelivr) | Path/URL to `pdf.worker.min.mjs` (required in browser) |
| `pdfCMapUrl` | `string` | pdfjs-dist's `cmaps/` in Node; CDN (jsDelivr) in the browser; `''` in the slim bundles | Folder (ending in `/`) of pdf.js's packed CMaps. pdf.js reads the text of a font a PDF does not embed through one when the font names a predefined CJK encoding (`UniJIS-UTF16-H`, `90ms-RKSJ-H`...); without it that text is lost. Only the files a document needs are read, each at most 4 MB and 15 s, and a folder that keeps a document waiting 15 s in all is asked nothing more for it. One that cannot be had raises `PDF_ASSET_UNAVAILABLE` and the rest of the document is read. In Node a path or `file:`/`http(s):` address you set replaces the installed folder |
| `pdfDecoderUrl` | `string` | pdfjs-dist's `wasm/` in Node; CDN (jsDelivr) in the browser; `''` in the slim bundles | Folder (ending in `/`) of pdf.js's image decoders: `jbig2.wasm` (JBIG2 and CCITT fax pictures, common in scans) and `openjpeg.wasm` (JPEG 2000), with the JavaScript fallbacks beside them (`*_nowasm_fallback.js`). Without them such pictures cannot be decoded. Read as `pdfCMapUrl` is; one that cannot be had raises `PDF_ASSET_UNAVAILABLE`, and each picture then not decoded `IMAGE_EXTRACTION_FAILED` |
| `pdfParserConfig` | `PdfParserConfig` | see below | PDF-specific options ([table below](#pdfparserconfig)) |
| `texParserConfig` | `TexParserConfig` | `{ today: '' }` | LaTeX options. `today` (`string`): what `\today` prints. `''` (default) prints the date of the parse, as LaTeX prints the date of the compile, in the document's language ("September 25, 2026" in English); set a fixed date, so that the same file parses the same way every day, or a placeholder of your own to find and replace later |
| `onWarning` | `(issue: OfficeIssue) => void` | (none) | Callback for non-fatal parsing issues |
| `abortSignal` | `AbortSignal \| null` | `null` | Optional signal to cancel parsing (rejects with AbortError). Once it fires the parse never resolves, even when it fires while OCR is recognizing an image. It cannot interrupt the synchronous reading of a document's markup; see [Cancellation](#cancellation-with-abortsignal) |

---

### PdfParserConfig

PDF-specific options, passed as `pdfParserConfig` on the parser config.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `useTags` | `boolean` | `true` | Use the tagged-structure tree (headings, tables, lists, notes) when present and reliable; set `false` to force geometry-only extraction |
| `detectColumns` | `boolean` | `true` | Recover reading order for multi-column and float-beside-text pages (recursive XY-cut) |
| `mergeHyphenatedWords` | `boolean` | `true` | Join words split across a line break by a trailing hyphen |
| `lineToleranceFactor` | `number` | `0.35` | Baseline tolerance (fraction of font size) for grouping fragments onto one line |
| `spaceToleranceFactor` | `number` | `0.25` | Gap threshold (fraction of font size) for inserting a space between fragments |
| `headingDetection` | `'auto' \| 'font-size' \| 'tags' \| 'off'` | `'auto'` | How heading levels are decided. `'auto'`: from tags when tagged, else a size/weight heuristic. `'font-size'`: re-level headings by the heuristic even on a tagged PDF (tables/lists stay tagged; a tagged heading is re-leveled by size and may be demoted). `'tags'`: only a heading the document tags (`H`, `H1`..`H6`) is a heading; text assembled from geometry (an untagged PDF, a page its tags cover too little of, text outside the tags, or `useTags: false`) never is, and the text is the same as under `'auto'`. `'off'`: never emit headings. A PDF heading says what decided it in `metadata.levelSource` (`'tags'` or `'font-size'`) |
| `pageRange` | `string` | `''` (all) | Restrict to given pages, e.g. `'1-3,7'`. Output keeps original page numbers |
| `normalizeText` | `boolean` | `true` | Unicode-normalize extracted text (expand ligatures, compose combining marks, regularize whitespace). Set `false` to preserve the raw source glyphs verbatim |
| `extractTextColor` | `boolean` | `true` | Extract each run's fill color into `formatting.color`, and the color of the filled rectangle drawn last under the run's whole box (white text on a colored bar; opaque, within its clip, none where it is white) into `formatting.backgroundColor`. Recovered from the operator list; on by default (color is content like bold/font). Costs about 1.6x parse time on a text-heavy PDF, near-free when `extractAttachments`/`ocr` already fetch the operator list; set `false` to skip it. Pure black is left unset. Highlight annotations set `formatting.backgroundColor` regardless of this flag |
| `keepHeadersAndFootersInPages` | `boolean` | `false` | Keep each page's running header and footer in that page's content, where they stand (the header before the page's text, the footer after it), instead of `auxiliary.headers`/`footers`. A tagged PDF marks them as artifacts; an untagged one's are the lines repeated in one place at the top or bottom of its pages (on two in five of them, no larger than the text, printed alike, or numbering the pages with a number set apart: at the line's end, at its start before a separator ("3 | Report"), framed by ornaments ("| 3 |", "~ 3 ~"), after a page's word, or as the whole line, a chapter's "2-3" included; a date's, a ticket's, a section's, a label's or a count's number does not count, so "4 May 2024" at the top of each page stays in the text), and its page numbers. Page numbers are page fields where the pages read run from page 1, and the page count only where all are read. Either goes to `auxiliary` as a word processor's header does where it is the same on every page read (the first of three or more may have none) once its page numbers are fields, a line larger than the text a paragraph of it: DOCX, ODT, RTF, LaTeX and the native PDF engine write it on every page, each numbered its own; text, Markdown, HTML (and the default HTML PDF engine), EPUB and slides leave it out. One whose text changes from page to page (a chapter's head, pages numbered from 147), or that only some pages have (a sheet's name), stays in its pages; a cover's own, where only the first page read differs and every other page (two at least) has the same one, stays on the cover. Turn it on for a page-faithful result (a slide per page), so every page keeps its own. `ignoreHeadersAndFooters` still drops them |
| `maxTextItems` | `number` | `20000` | Base of the text items one PDF may yield; the limit is this plus one per byte of the file, an item counting once more per 64 characters. Nested form XObjects can multiply a few KB into millions of items; past the limit the rest of the document is not read (`PDF_CONTENT_LIMIT_EXCEEDED`). |
| `maxOperators` | `number` | `250000` | Base of the drawing operators kept from one PDF (for images, text color and font styles); the limit is this plus four per byte of the file. Paths do not count (a page keeps up to 40,000 of them apart, for bullets drawn as shapes and tables drawn with ruling lines); a shown string counts once more per 64 glyphs. Past it no more operators are read: from that page on, images, text colors and font styles are missing, and the text is still read (`PDF_CONTENT_LIMIT_EXCEEDED`). |
| `maxAnnotations` | `number` | `10000` | Base of the annotations (links, highlights) one PDF may have read; the limit is this plus one per 32 bytes of the file. Past it the rest of the document's links and highlights are not read (`ANNOTATION_EXTRACTION_FAILED`). |
| `maxTimeMs` | `number` | `5000` | Base of the CPU time the separate pdf.js process may spend reading one PDF; the limit is this plus 20 ms per KB of the file. CPU time, not time passed, so a busy host, many parses at once or a loaded machine never cut a document short. Past it the process is ended and the rest of the document is not read (`PDF_CONTENT_LIMIT_EXCEEDED`). pdf.js in the host (`separateProcess: false`, or the fallback) has no time limit, since no measure of time there is the document's own; bound it with `abortSignal`. |
| `maxImagePixels` | `number` | `40000000` | The most pixels (width times height) a picture may have to be extracted with `extractAttachments`. A few bytes of a compressed picture can declare dimensions that decode to gigabytes, so pdf.js leaves out a larger one without decoding it, reported with `IMAGE_EXTRACTION_FAILED` (with its page) naming this option. Raise it for high-resolution scans (an A3 page at 600 dpi is about 42 million pixels); `Infinity` is no limit. |
| `separateProcess` | `boolean` | `true` | Run pdf.js in a separate process (Node) under `processMemoryMb`, so a PDF that inflates into gigabytes inside pdf.js fails with `PDF_PROCESS_FAILED` instead of ending the host; `abortSignal` and `maxTimeMs` stop it mid-request. Processes are pooled, at most one per CPU, each reading one document at a time (further parses wait), and end with the host. Falls back to the host with `PDF_SEPARATE_PROCESS_UNAVAILABLE`. |
| `processMemoryMb` | `number` | `1024` | Heap, in MB, of the separate pdf.js process (64 or more). |

---

### GeneratorConfig (Common)

Options shared by all generator formats. Pass to `OfficeGenerator.generate(ast, format, config)` or `ast.to(format, config)`.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `includeFormatting` | `boolean` | `true` | Include bold/italic/colors/sizes in output (HTML, Markdown, DOCX, ODT, LaTeX, RTF, PPTX, ODP, and XLSX and ODS cells' bold; a no-op for text/CSV/chunks, which carry no run formatting) |
| `generateIds` | `boolean` | `true` | Slug-based heading anchors: `id` attributes on HTML headings, and a `{#slug}` suffix on Markdown headings (`# Title {#title}`, kramdown/Pandoc). Set `false` to omit both, useful when the Markdown is rendered by GFM/CommonMark, which show `{#slug}` as literal text. A top-level option (not under `mdConfig`/`htmlConfig`); it affects HTML, Markdown, DOCX, ODT and LaTeX (the formats that carry a heading anchor/bookmark id; in LaTeX a `\label`). |
| `renderMetadata` | `boolean` | `false` | Render title/author as a visible header block. Rendered by CSV, DOCX, HTML (and the Puppeteer PDF engine), EPUB, text, ODT, LaTeX (`\maketitle`, or a beamer title frame, unless the content has its own title block) and RTF; the native PDF engine, Markdown, XLSX, ODS, PPTX and ODP do not |
| `metadataOverrides` | `MetadataOverrides` | `{}` | Override the metadata embedded in the output, merged per field over `ast.metadata` |
| `includeImages` | `boolean \| 'image-only' \| 'image+ocr-text' \| 'ocr-text-only' \| 'none'` | `true` | How to render an image node. `true`=`'image-only'` (embed the image, no OCR text); `'image+ocr-text'` (image then its recognized/OCR text); `'ocr-text-only'` (OCR text, no image); `false`=`'none'` (omit). In plain-text output an image becomes an `[Image: name]` placeholder (plus OCR text for `'image+ocr-text'`), or just the OCR text for `'ocr-text-only'`. In chunks a picture is `[Image: alt]` for its alt text, with its OCR text, except under `'none'` and `'ocr-text-only'` (OCR text only). XLSX and ODS, whose cells hold no pictures, write a picture's OCR text in its place under `'ocr-text-only'` and `'image+ocr-text'`; PPTX and ODP write it in the picture's place under `'ocr-text-only'`, and under `'image+ocr-text'` after a flowing document's picture slide or in a placed slide's notes |
| `maxInlineImageBytes` | `number` | `1500000` | Max decoded image size, in bytes, that is inlined as a `data:` URI (HTML/Markdown); the base64 URI itself is ~1/3 larger, so a scanned page cannot emit a multi-megabyte line that breaks downstream parsers. Under the default `image-only` mode an image over the cap renders its recognized/OCR text when it has any (multi-line OCR as a fenced block in Markdown), otherwise a compact name reference; Markdown still emits the `IMAGE_NOT_INLINED` warning. Plain text follows the same rule. **Standalone HTML inlines whatever the cap**: a self-contained document has nowhere else to resolve the image from. In every mode (standalone HTML and `Infinity` included), the pictures one document writes inline stop at 128 MB in all (a picture is inlined at every place showing it); the rest are name references, with `IMAGE_NOT_INLINED`. `0` never inlines, `Infinity` inlines every picture within those 128 MB |
| `includeCharts` | `boolean` | `true` | Include charts: HTML renders an interactive Chart.js canvas (or, with no `chartJsSrc`, the chart's data as a table), DOCX/ODT/LaTeX render the chart's data as a table, PPTX/ODP a table of its data in the chart's place, plain text and the native PDF engine render the chart's data text; Markdown and RTF render nothing for a chart, and XLSX/ODS, whose cells cannot hold one, report it (`CONTENT_NOT_REPRESENTABLE`). `false` omits charts in every generator |
| `ignoreInternalLinks` | `boolean` | `false` | Strip bookmarks and internal anchors from output (HTML, Markdown, DOCX, ODT, LaTeX, RTF, and the native PDF engine's links into the document) |
| `ignoreDefaultStyleMap` | `boolean` | `false` | Disable built-in style mappings (e.g., "Heading 1" → h1) |
| `styleMap` | `string[] \| StructuredStyleMapping[]` | `[]` | Custom semantic style mappings |
| `onNode` | `(node) => string \| false \| void` | (none) | Per-node callback for filtering, overriding, or mutating |
| `onWarning` | `(issue: OfficeIssue) => void` | (none) | Callback for non-fatal generation issues |
| `abortSignal` | `AbortSignal \| null` | `null` | Optional signal to cancel the generation operation (rejects with AbortError). Checked between steps by every generator (both PDF engines: the native one checks before each node it draws; XLSX and ODS before each cell and block of text; PPTX and ODP before each node a slide takes) |

---

### `onNode` Callback: Advanced Node Manipulation

Called for **every node** in the AST during generation. Can be `async`.

| Return value | Effect |
|---|---|
| `false` | Skip this node and all its children |
| `string` | Use this string as the output for this node, skip default logic |
| `void` | Proceed with default rendering (mutations to `node` are applied) |

```ts
const { value: md } = await ast.to('md', {
    onNode: async (node) => {
        // Skip all images
        if (node.type === 'image') return false;

        // Redact secrets (mutate then proceed)
        if (node.text?.includes('SECRET_KEY')) {
            node.text = node.text.replace(/SECRET_KEY: \w+/, 'SECRET_KEY: [REDACTED]');
        }

        // Custom rendering for a specific style
        if (node.metadata?.style === 'Callout') {
            return `> [!INFO]\n> ${node.text}`;
        }
    }
});
```

---

### `styleMap`: Semantic Style Mapping

Maps document style names to semantic output elements. Two formats supported:

#### Structured Objects (Recommended)

```ts
styleMap: [
    {
        selector: { nodeType: 'paragraph', attributes: { style: 'Heading 1' } },
        output: { tag: 'h1', classes: ['main-title'], attributes: { id: 'top' } }
    },
    {
        // '~=' operator matches if the word 'Quote' appears anywhere in the style name
        selector: { attributes: { style: { value: 'Quote', operator: '~=' } } },
        output: { tag: 'blockquote', fresh: true }
    }
]
```

`fresh: true` prevents the generator from merging adjacent nodes of the same tag into one block.

#### Legacy String DSL

Compatible with `mammoth.js` style maps:

```js
styleMap: [
    "p[style-name='Heading 1'] => h1",
    "p[style~='Title'] => h2",
    "p[style-name='Quote'][lang='en'] => blockquote"
]
```

---

### HtmlGeneratorConfig

Pass as `htmlConfig` inside `GeneratorConfig`.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `standalone` | `boolean \| StandaloneConfig` | `true` | Controls the HTML "document envelope" (see below) |
| `chartJsSrc` | `string` | jsDelivr CDN | URL for the Chart.js library, loaded only by a page that draws a chart. `''` loads none: a chart is then written as a table of its values, with a `CONTENT_NOT_REPRESENTABLE` warning. Under `allowRemoteAssets: false` it is `''` unless set |
| `containerWidth` | `string \| number` | `'auto'` | Max width of the content container. Positive number (px), CSS length string (`'900px'`, `'100%'`, `'60vw'`), or `'auto'`. Invalid values fall back to `'auto'` with an `INVALID_CONTAINER_WIDTH` warning |
| `customCss` | `string` | `''` | Raw CSS injected into the `<style>` block; use this to override built-in styles |
| `injections.headStart` | `string` | `''` | Raw HTML injected after `<head>` |
| `injections.headEnd` | `string` | `''` | Raw HTML injected before `</head>` |
| `injections.bodyStart` | `string` | `''` | Raw HTML injected after `<body>` |
| `injections.bodyEnd` | `string` | `''` | Raw HTML injected before `</body>` |
| `sourceAttributes` | `boolean` | `false` | Carry each rich node's raw source in a `data-*` attribute (undelimited text), so attribute-driven consumers can rehydrate it: `data-wikilink`/`data-target`/`data-alias` on wikilinks, a `<span class="citation" data-key>` for citations, the LaTeX in `data-math`, a `<div class="mermaid" data-mermaid>` for mermaid, and an empty `<span data-html-comment="…">` for a source comment (editors' DOM parsers discard real `<!-- -->`). Off = byte-identical to before; the parser reads every shape it emits. Forced off for PDF/EPUB |
| `omitDefaultTextColor` | `boolean` | `false` | Omit an inline run `color` equal to the document default (near-black or near-white), so imported text adapts to the reader's light/dark theme instead of being pinned to black or white. Only near-black/near-white run colours are dropped; deliberately-coloured runs are emitted unchanged (Word's `w:val="auto"` already carries no colour). Off = byte-identical to before |

#### `standalone`: granular envelope control

`standalone` conflates several independent decisions: whether to emit the `<!doctype>/<html>/<head>/
<body>` shell, how CSS is delivered, and whether to inject scripts/meta tags/injections. The boolean
shorthand still works: **`true`/omitted turns every part on** (a complete document); **`false` turns
every part off** (a bare content fragment, safe to drop into a page you don't control). Pass an
object instead for granular control; any field you omit defaults to its "on" (standalone) value:

| `StandaloneConfig` field | Type | Default | Description |
|--------|------|---------|-------------|
| `document` | `boolean` | `true` | Wrap in `<!DOCTYPE html><html><head>…</head><body>…</body></html>` |
| `metaTags` | `boolean` | `true` | Emit `<title>`/`<meta>` tags. Only meaningful when `document` is true |
| `styles` | `'full' \| 'scoped' \| 'none'` | `'full'` | See below |
| `scripts` | `boolean` | `true` | Emit the Chart.js CDN loader and spreadsheet-interactivity `<script>` tags |
| `headInjections` | `boolean` | `true` | Apply `injections.headStart`/`headEnd`. Only meaningful when `document` is true |
| `bodyInjections` | `boolean` | `true` | Apply `injections.bodyStart`/`bodyEnd`, even to a bare fragment |

`styles` controls how the built-in stylesheet is delivered:
- **`'full'`**: the complete stylesheet using global selectors (`body`, `h1`, `table`, …). This is
  what `standalone: true` has always emitted.
- **`'scoped'`**: the same styling, scoped under the fragment's own wrapper via CSS `@scope` so it
  cannot leak onto a host page's elements. Requires a modern engine (Chrome 118+, Safari 17.4+,
  Firefox 128+); for universal support use `'none'` (bring your own CSS) or `'full'`.
- **`'none'`**: no stylesheet at all; the host page (or rich-text editor, or EPUB reader) supplies
  its own styling.

```js
// A styled fragment to embed in your own page, without a document shell:
await ast.to('html', { htmlConfig: { standalone: { document: false } } });

// The same, but with styles scoped so they can't leak onto your page's own elements:
await ast.to('html', { htmlConfig: { standalone: { document: false, styles: 'scoped' } } });

// A completely bare fragment (no shell, no styles, no scripts), e.g. for a rich-text editor:
await ast.to('html', { htmlConfig: { standalone: false } });
```

> [!NOTE]
> **Behavior change from `standalone: false`:** previously this emitted a fragment with a *global,
> unscoped* `<style>` block. It now emits a genuinely bare fragment (no `<style>` at all), matching
> "every part off." If you relied on the old styled-fragment behavior, pass
> `{ document: false }` (or `{ document: false, styles: 'full' }`) instead.

### MdGeneratorConfig

Pass as `mdConfig` inside `GeneratorConfig`.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `fallbackToHtml` | `boolean \| FallbackToHtmlConfig` | `true` | Use HTML tags for features Markdown cannot represent (underlines, merged table cells, embeds, etc.). Pass an object for per-feature control. `cellLineBreaks`/`itemLineBreaks` (default on) join multi-line table-cell / multi-paragraph list-item content (and a heading's or definition's line breaks) with `<br>` instead of a space. `inlineFormatting` (default `false`, opt-in even when the boolean is `true`) additionally round-trips inline color/highlight/font-size as `<span style="...">` runs. |
| `dialect` | `MarkdownDialectPreset \| MarkdownDialectConfig` | `'extended'` | Which native syntax to emit for constructs that differ across targets (GitHub/GitLab/Obsidian/Pandoc/CommonMark). Each capability is typed by the syntax it selects (e.g. `strikethrough: 'tilde'`, `highlight: 'equals'`, `admonitions: 'blockquote'`), with `'none'` to turn it off. See [Markdown Dialect Support](#markdown-dialect-support). The old `boolean` toggles and admonition flavour names (`'github'`/`'gitlab'`/`'pandoc'`) still work but are deprecated. |

### PdfGeneratorConfig

Pass as `pdfConfig` inside `GeneratorConfig`. The default `'html'` engine requires the optional `puppeteer` peer dependency; the `'native'` engine requires the optional `pdf-lib` peer dependency instead, and `@pdf-lib/fontkit` as well when it is given fonts. In the **prebuilt browser bundle** `pdf-lib` and `@pdf-lib/fontkit` are stubbed so the bundle stays self-contained, so client-side native PDF export means importing from the dedicated entry **`officeparser/browser-native-pdf`** and installing `pdf-lib` (a consumer that bundles officeParser from source resolves both normally).

**Fonts in the native engine.** Without fonts the native engine draws in the Standard-14 Helvetica and Courier fonts, which have Western European characters only: anything else is drawn as `?` and reported once (`CONTENT_NOT_REPRESENTABLE`, `details.reason: 'unencodable-text'`, with how many characters and a sample). Give it font files to draw the rest:

```js
const noto = new Uint8Array(await (await fetch('/fonts/NotoSans-Regular.ttf')).arrayBuffer());
const notoBold = new Uint8Array(await (await fetch('/fonts/NotoSans-Bold.ttf')).arrayBuffer());
const devanagari = new Uint8Array(await (await fetch('/fonts/NotoSansDevanagari-Regular.ttf')).arrayBuffer());

const { value: pdf } = await ast.to('pdf', {
    pdfConfig: {
        engine: 'native',
        fonts: { regular: noto, bold: notoBold },   // also italic, boldItalic, mono
        fallbackFonts: [devanagari],                // tried in order for what the fonts above lack
    },
});
```

Each font is the bytes of a TrueType, OpenType, WOFF or WOFF2 file (a web font, such as Fontsource's, is embedded as the TrueType or OpenType font it holds). Each character is drawn in the first font that has it, so one document can mix scripts. Words are shaped by the font (ligatures, Arabic joining, Indic conjuncts), and a right-to-left line starts at the right edge with its words in visual order; text of the Indic and Southeast Asian scripts, and letters with combining marks, is drawn with each glyph where the font places it (through its positioning table, `GPOS`) and marked with what it says (`/ActualText`), so it copies in the order it is written. Fonts are subset to the glyphs used, but for OpenType fonts of CFF outlines, and fonts with color layers (`COLR`), which are embedded whole. A font that cannot be read (a damaged one, a collection such as a `.ttc`), or whose glyphs to be drawn cannot be read or written into the PDF, fails the conversion (`PDF_GENERATION_FAILED`) and names the option it was given in.

**Which font draws what.** A character is drawn by the first font whose glyph for it draws something. A font whose glyph for it is empty is passed over, and a character no font draws is drawn as `?` and reported with the characters no font has: an empty glyph is a font's defect, or, in a color font, the stand-in for a color glyph. A character blank by design (a space, a control or format character, the Braille blank cell U+2800) is drawn by the first font that has it, and so is a combining mark that a font substituting glyphs as it shapes (`GSUB`, `morx`) has with the letter before it (an invisible virama that stacks the next letter). The Standard-14 fonts have no rupee sign `₹` (U+20B9), ruble sign `₽` (U+20BD) or emoji, and fonts differ: of Fontsource's files, Noto Sans has `₹` and `₽` in its `latin-ext` file and neither in its `latin` one, Noto Sans Devanagari has `₹` but not `₽`, and Liberation Sans has neither. Emoji are drawn by a monochrome font of outlines such as Noto Emoji. A color emoji font (Noto Color Emoji, Apple Color Emoji) draws its emoji from color images or layers that the native engine does not draw: they are counted with the characters no font has, and a color font none of whose characters has an outline is reported (`INVALID_CONFIG_VALUE`) and not used.

**The maintained pdf-lib fork.** `@cantoo/pdf-lib` 2.x (the same API, with encryption) is supported in place of `pdf-lib`, installed under its name, with the fork's own fontkit under `@pdf-lib/fontkit` when fonts are used (each pdf-lib subsets fonts only with its own fontkit):

```json
"dependencies": {
    "pdf-lib": "npm:@cantoo/pdf-lib@^2.11.1",
    "@pdf-lib/fontkit": "npm:@cantoo/fontkit@^2.0.12"
}
```

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `engine` | `'html' \| 'native'` | `'html'` | `'html'`: render through a headless browser (Puppeteer in Node; native print in the browser) for maximum fidelity. `'native'`: lay the AST out directly with `pdf-lib` - no browser, runs in Node and the browser (real PDF bytes client-side), much lighter, but draws in the fonts it is given (the Standard-14 fonts when given none) and reflows rather than pixel-matching |
| `fonts` | `{ regular?, bold?, italic?, boldItalic?, mono? }` of `Uint8Array \| ArrayBuffer` | none | Native engine: the font file (`.ttf`, `.otf`, `.woff` or `.woff2` bytes) each style is drawn in. A style given none is drawn in `regular`; code stays in Courier unless `mono` is given. Needs `@pdf-lib/fontkit`. Bytes that are not a font fail the conversion (`PDF_GENERATION_FAILED`) |
| `fallbackFonts` | `(Uint8Array \| ArrayBuffer)[]` | `[]` | Native engine: more font files (TrueType, OpenType, WOFF, WOFF2), tried in order for each character the fonts before them lack. Given without `fonts`, text stays in the Standard-14 fonts and these draw only what those lack. A color font none of whose characters has an outline is reported (`INVALID_CONFIG_VALUE`) and not used |
| `subsetFonts` | `boolean` | `true` | Native engine: embed only the glyphs used of each font given. `false` embeds whole fonts (a WOFF or WOFF2 file as the font it holds). An OpenType font of CFF outlines, or a font with color layers (`COLR`), is embedded whole either way |
| `tagged` | `boolean` | `true` | Emit a tagged, PDF/UA-accessible PDF (structure tree). HTML engine only; the native engine ignores it |
| `outline` | `boolean` | `false` | Emit a bookmark outline from the document's headings, nested by heading level. Both engines |
| `format` | `string` | `'auto'` | Paper format (`'A4'`, `'Letter'`, `'Legal'`, etc.). `'auto'`: the document's own pages (`metadata.pageSetup`; the native engine gives each section and sheet its own and takes unset margins from the document), else A4 |
| `width` | `string \| number` | `''` | Paper width (e.g., `'5in'`, `'3cm'`) or pixels |
| `height` | `string \| number` | `''` | Paper height (e.g., `'5in'`, `'3cm'`) or pixels |
| `landscape` | `boolean` | `false` | Landscape page orientation (both engines) |
| `printBackground` | `boolean` | `true` | Print background graphics. **HTML engine only** |
| `margin` | `object` | see note | Page margins (`top`, `right`, `bottom`, `left`). Default differs by engine: the HTML engine uses `0` (the body carries its own padding); the native engine uses a small default (~48pt) so text is not glued to the sheet edge. An explicit value (including `0`) is honored by both |
| `displayHeaderFooter` | `boolean` | `false` | Show print header/footer. **HTML engine only** |
| `headerTemplate` | `string` | `''` | HTML template for the print header. **HTML engine only** |
| `footerTemplate` | `string` | `''` | HTML template for the print footer. **HTML engine only** |
| `scale` | `number` | `1` | Rendering scale factor. **HTML engine only** |
| `launchOptions` | `object` | headless defaults | Puppeteer launch options (e.g., `executablePath`). **HTML engine only** |
| `timeout` | `number` | `30000` | PDF rendering timeout in milliseconds. Set to `0` to disable. **HTML engine only** |
| `printHidden` | `boolean` | `false` | Print hidden sheets, rows and columns, which a printout leaves out. Both engines |
| `printSpeakerNotes` | `boolean` | `false` | Print each slide's speaker notes after it; a printout leaves them out. Both engines |

> **Native engine (`engine: 'native'`) differences beyond fidelity:** it ignores the Puppeteer-only
> options above (and `tagged`), uses a non-zero default `margin` (the document's own, where it sets
> one), draws a chart as a bar chart of its data, and does not apply `renderMetadata`. It draws the
> document's own headers and footers on every page, footnotes at the foot of their page, clickable
> links, slides as they are laid out, and text in a section's columns. It draws an SVG picture as vector
> graphics in its place (paths, shapes, text, embedded PNG and JPEG pictures, fills, strokes, dashes
> and opacity, styled by attributes, `style` and the picture's own stylesheet); a gradient is drawn in
> one colour, and filters, masks, clip paths and patterns are left out, both reported. Choose it to
> avoid the headless browser; choose `'html'` for pixel-matching fidelity, other chart kinds and PDF/UA
> tags.

### DocxGeneratorConfig

Pass as `docxConfig` inside `GeneratorConfig`. The DOCX generator writes a real WordprocessingML package (`.docx`) with zero extra dependencies, so it runs identically in Node and the browser and returns a `Uint8Array`. It reproduces headings, styled runs, tables (including merged cells), lists, images, hyperlinks, footnotes/endnotes, comments, headers/footers and metadata from any parsed source.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `format` | `PaperFormat` | `'auto'` | Page size for the document section (`w:pgSz`). `'auto'` is the document's own page (`metadata.pageSetup`), else A4; or the names `pdfConfig.format` takes (`A4`, `Letter`, `Legal`, `A3`, `Tabloid`, …; case-insensitive) |
| `landscape` | `boolean` | `false` | Landscape orientation (swaps page dimensions and sets `w:orient`) |
| `margin` | `object` | the document's, else `72` | Page margins (`top`, `right`, `bottom`, `left`). Each is a number of points (1/72 inch) or a unit string (`'1in'`, `'2cm'`, `'36pt'`). A side left unset is the document's own margin under `format: 'auto'`, else 72 (one inch) |

```typescript
import { OfficeConverter } from 'officeparser';
import { writeFileSync } from 'fs';

// Any supported source → Word. convert() extracts the images itself; the CLI needs --extractAttachments.
const { value } = await OfficeConverter.convert('report.md', 'docx', {
    generatorConfig: { docxConfig: { format: 'Letter', margin: { top: 36, right: 36, bottom: 36, left: 36 } } }
});
writeFileSync('report.docx', value); // value is a Uint8Array
```

### OdtGeneratorConfig

Pass as `odtConfig` inside `GeneratorConfig`. The ODT generator writes a real OpenDocument Text package (`.odt`) with zero extra dependencies, so it runs identically in Node and the browser and returns a `Uint8Array`. It is the round-trip partner of the ODF parser: headings, styled runs, tables (including merged cells), nested lists, images, hyperlinks, footnotes, comments and headers/footers all re-parse. Formatting is carried in ODF automatic styles, and the package is reproducible (pinned zip mtimes, deterministic style/id names).

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `format` | `PaperFormat` | `'auto'` | Page size for the page layout (`style:page-layout`). `'auto'` is the document's own page (`metadata.pageSetup`), else A4; or the names `pdfConfig.format` takes (`A4`, `Letter`, `Legal`, `A3`, `Tabloid`, …; case-insensitive) |
| `landscape` | `boolean` | `false` | Landscape orientation (swaps page dimensions and sets `style:print-orientation`) |
| `margin` | `object` | the document's, else `72` | Page margins (`top`, `right`, `bottom`, `left`). Each is a number of points (1/72 inch) or a unit string (`'1in'`, `'2cm'`, `'36pt'`). A side left unset is the document's own margin under `format: 'auto'`, else 72 (one inch) |

```typescript
import { OfficeConverter } from 'officeparser';
import { writeFileSync } from 'fs';

// Any supported source → OpenDocument Text. convert() extracts the images itself; the CLI needs --extractAttachments.
const { value } = await OfficeConverter.convert('report.docx', 'odt');
writeFileSync('report.odt', value); // value is a Uint8Array
```

### XlsxGeneratorConfig

Pass as `xlsxConfig` inside `GeneratorConfig`. The XLSX generator writes an Excel workbook (`.xlsx`) with zero extra dependencies, so it runs identically in Node and the browser and returns a `Uint8Array`. A spreadsheet (XLSX, ODS, CSV) is written sheet for sheet: each cell's value typed (numbers, percentages, currency amounts, dates and times) and shown in its number format (read from `CellMetadata.shownText`), header cells bold, merged cells merged, and hidden sheets, rows and columns hidden. A sheet's stored numbers stay numbers, exponents and all (`1E-10`, `0.30000000000000004`). Any other document is written as its tables, each on a sheet named after the heading before it (else `Table 1`, `Table 2`, ...), with its cells typed from their text: `1,250.5` is a number, `19%` a percentage, `€1,250.00` a currency amount (`($80.00)` a negative one), `2026-10-01` a date, `14:05` a time. A code with leading zeros (`007`), a number with a leading plus (`+4915112345678`) or a number of more than fifteen digits (a card number, which a spreadsheet would round) stays text, and a whole number of twelve or more digits is shown in full. The document's paragraphs, headings and list items outside its tables go on a sheet named `Text`, one a row, headings bold. Footnotes, endnotes, comments and speaker notes go on a sheet named `Notes` after the others (`Footnote 1`, `Endnote 1`, `Comment (author)`, `Slide 2`, each beside its text), and the text cites each note by its label there (`[Footnote 1]`). A chart is a sheet of its data (its labels down, its series across), a list item keeps its marker, and an embedded video or page is its address. XLSX holds at most 32,767 characters in a cell: a longer one is cut there and reported. Pictures, which a cell cannot hold, are reported with `CONTENT_NOT_REPRESENTABLE`, their alternative text written in their place (their recognized text under `includeImages: 'ocr-text-only'`, both under `'image+ocr-text'`); a chart without its data is reported too (not where `includeCharts` is `false`); a document with neither a table nor text is written as one empty sheet, and reported.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `sheetPerTable` | `boolean` | `true` | Each of a document's tables on a sheet of its own. `false` puts them all on one sheet named `Tables`, an empty row between them. No effect on a spreadsheet, whose sheets are written as they are |
| `includeText` | `boolean` | `true` | A document's text outside its tables on a sheet named `Text`. `false` leaves it out. No effect on a spreadsheet |

```typescript
import { OfficeConverter } from 'officeparser';
import { writeFileSync } from 'fs';

// A report's tables on one sheet, without its text
const { value } = await OfficeConverter.convert('report.docx', 'xlsx', {
    generatorConfig: { xlsxConfig: { sheetPerTable: false, includeText: false } }
});
writeFileSync('report-tables.xlsx', value); // value is a Uint8Array
```

### OdsGeneratorConfig

Pass as `odsConfig` inside `GeneratorConfig`. The ODS generator writes an OpenDocument spreadsheet (`.ods`) from the same workbook as XLSX output, with the same options: a spreadsheet sheet for sheet, any other document as its tables and its text, each cell typed (`office:value-type` `float`, `percentage`, `currency`, `date` or `time`, with the text the cell shows beside its value). It returns a `Uint8Array` and needs no extra dependencies.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `sheetPerTable` | `boolean` | `true` | Each of a document's tables on a sheet of its own; `false` puts them all on one sheet named `Tables` |
| `includeText` | `boolean` | `true` | A document's text outside its tables on a sheet named `Text`; `false` leaves it out |

```typescript
import { OfficeConverter } from 'officeparser';
import { writeFileSync } from 'fs';

const { value } = await OfficeConverter.convert('data.xlsx', 'ods');
writeFileSync('data.ods', value); // value is a Uint8Array
```

### PptxGeneratorConfig

Pass as `pptxConfig` inside `GeneratorConfig`. The PPTX generator writes a PowerPoint presentation (`.pptx`) with zero extra dependencies, so it runs identically in Node and the browser and returns a `Uint8Array`. What the slides hold depends on the source:

- **A presentation** (PPTX, ODP) is written slide for slide at its slide size: each text box, shape (its outline, fill, line, turn and mirroring, with the text in it), picture and table in its place (`bounds`), the slide's background colour, and its speaker notes on the slide's notes page. A chart is a table of its data in the chart's place.
- **A PDF** becomes a slide for each page, at the pages' most common size (a page of another size is scaled into it, centred). Each paragraph is a text box where the page set it, its lines broken and spaced as on the page (each line on the baseline the page drew it on, where the runs' `bounds.baseline` give it), in its runs' fonts, sizes and colours; a paragraph set beside a floating table or picture is a box for each group of its lines. Tables keep their measured row heights and their cells' shading, pictures their places, and footnotes and a page's notes are at the foot of their page. A page's invisible text (a scan's recognised words laid over its picture, `TextFormatting.invisible`) is written fully transparent: not seen, still there to select and search. A table whose text is all invisible is drawn without rules (its cells' text in their places where the page gave them).
- **Any other document** flows onto 16:9 slides: a slide for each heading of the first two levels, titled by it, with the text under it as bullets (lower headings bold), carried on to more slides of the same title where it does not fit. Text before the first heading is under the document's title, from its metadata. Each table is on slides of its own, ten rows a slide with its header rows repeated, and each sheet of a workbook is a table under the sheet's name. Pictures (those in a table's cells too) are on slides of their own, charts are tables of their data, a presentation's speaker notes go with the slides its content made, and footnotes and endnotes, numbered where they are cited, are listed on slides at the end.

A numbered list item carries its own number as the document shows it (`2.`, `1.1.`, `iv.`), and a picture its alternative text. A code block is set in a monospaced font, line for line; an embedded video or page is its link; a picture known only by its address is its alternative text, and reported; comments go into the speaker notes of the slide they are on. A PDF page shown turned (`/Rotate`) has its text and pictures turned with it. A presentation or PDF parsed with `ignorePageGeometry` has no places on its pages, so it flows as any other document. Pictures and charts need `extractAttachments: true` on the parse (`OfficeConverter.convert` sets it); PNG, JPEG, GIF and BMP pictures are written, and others are reported with `IMAGE_PROCESSING_FAILED`. Links to web, mail and phone addresses stay clickable. What a slide cannot hold (a chart without its data, a picture without a place on a page) is reported with `CONTENT_NOT_REPRESENTABLE`.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `backgrounds` | `(Uint8Array \| null \| undefined)[]` | `[]` | A picture (PNG or JPEG bytes) for each page of a PDF or slide of a presentation, in order, drawn as its slide's background under its text; `null` or `undefined` for one without. Give a PDF's pages drawn without their text, and the slides look as the pages did (shapes, gradients and charts included) with the text over them to edit; a page's tables are then their cells' text in place over the picture, which holds their rules. The page's pictures are placed over the background as objects of their own too, so a picture the background also holds is there twice and moving one uncovers its copy beneath. A flowing document's slides take them in order |

```typescript
import { OfficeConverter } from 'officeparser';
import { readFileSync, writeFileSync } from 'fs';

// A PDF as slides, each over a picture of its page drawn without its text (rendered beforehand)
const backgrounds = ['page1.png', 'page2.png'].map(file => new Uint8Array(readFileSync(file)));
const { value } = await OfficeConverter.convert('slides.pdf', 'pptx', {
    generatorConfig: { pptxConfig: { backgrounds } }
});
writeFileSync('slides.pptx', value); // value is a Uint8Array
```

### OdpGeneratorConfig

Pass as `odpConfig` inside `GeneratorConfig`. The ODP generator writes an OpenDocument presentation (`.odp`) from the same deck as PPTX output, with the same option: a presentation slide for slide, a PDF's pages as slides with their text in place, any other document flowed onto slides. A shape is drawn in ODF's own outline of its kind where it has one (rectangles, rounded rectangles, ellipses, triangles, diamonds, polygons, stars, arrows, lines), else as a rectangle, which is reported. It returns a `Uint8Array` and needs no extra dependencies.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `backgrounds` | `(Uint8Array \| null \| undefined)[]` | `[]` | A picture (PNG or JPEG bytes) for each page of a PDF or slide of a presentation (a flowing document's slides, in order), under its slide's text; `null` or `undefined` for one without. As PPTX's: the page's pictures are placed over it as objects of their own too |

```typescript
import { OfficeConverter } from 'officeparser';
import { writeFileSync } from 'fs';

const { value } = await OfficeConverter.convert('deck.pptx', 'odp');
writeFileSync('deck.odp', value); // value is a Uint8Array
```

### TexGeneratorConfig

Pass as `texConfig` inside `GeneratorConfig`. The LaTeX generator turns any parsed document into LaTeX source that compiles unmodified with **pdfLaTeX, XeLaTeX, LuaLaTeX, upLaTeX, pLaTeX and `latex`** (the last three through `dvipdfmx`), TeX Live 2021 or later: the preamble selects fonts per engine (`fontspec` under XeTeX and LuaTeX, with fonts for Greek, Cyrillic and CJK text where the document has it; `fontenc`/`inputenc` otherwise; see [Engines](#latex-support)) and loads only the packages the document actually uses. The value is a `string` (the `.tex` source), or a `Uint8Array` zip when `bundle` is set. Zero extra dependencies; runs in Node and the browser.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `documentClass` | `'auto' \| 'article' \| 'report' \| 'book' \| 'beamer'` | `'auto'` | `auto` writes a `beamer` presentation when the content is made of slides (PPTX/ODP) and an `article` otherwise. `report`/`book` map level-1 headings to `\chapter`. `beamer` for a non-presentation source starts a new frame at each level-1/2 heading |
| `standalone` | `boolean` | `true` | `false` emits only the body, headed by a comment listing the packages (and, for Greek, Cyrillic or CJK text, the font setup) the including document needs |
| `bundle` | `boolean` | `false` | `true` returns a zip holding `main.tex` and every image as a file under `images/`, ready to compile or upload to Overleaf. With `false` the result is the `.tex` alone, carrying its PNG and JPEG images inside it (see `embedImages`) |
| `embedImages` | `boolean` | `true` | Carry each PNG and JPEG image inside the `.tex`, so the one file compiles with its pictures (see **Images inside the `.tex`** below). `false` refers to `images/<name>` files instead, and an `IMAGES_NOT_BUNDLED` warning names the files to place there (their bytes are in `ast.attachments`). Ignored with `bundle` |
| `numberSections` | `boolean` | `false` | Number sections (`1`, `1.1`, ...). Off by default, matching office documents, whose headings are unnumbered |
| `format` | `PaperFormat` | `'auto'` | Paper size, written as a `geometry` option. `'auto'` is the document's own page, else A4; or the names `pdfConfig.format` takes. Ignored by `beamer` |
| `landscape` | `boolean` | `false` | Landscape orientation. Ignored by `beamer` |
| `margin` | `object` | the document's, else `72` | Page margins (`top`, `right`, `bottom`, `left`), each a number of points or a unit string (`'1in'`, `'2cm'`, `'36pt'`). A side left unset is the document's own margin under `format: 'auto'`, else 72 (one inch). Ignored by `beamer` |

```typescript
import { OfficeConverter } from 'officeparser';
import { writeFileSync } from 'fs';

// One .tex file that carries its images
const { value: tex } = await OfficeConverter.convert('paper.docx', 'tex');
writeFileSync('paper.tex', tex as string);

// The images as separate files: main.tex + images/ in one zip
const { value: zip } = await OfficeConverter.convert('paper.docx', 'tex', { generatorConfig: { texConfig: { bundle: true } } });
writeFileSync('paper.zip', zip as Uint8Array);
```

**How the document maps to LaTeX**

| Content | LaTeX |
|---|---|
| Headings | `\section` ... `\subparagraph` (`\chapter` first in `report`/`book`), unnumbered unless `numberSections`; heading ids and linked bookmarks become `\label`s |
| Run formatting | `\textbf`, `\textit`, `\uline`/`\sout` (`ulem`), `\textsuperscript`/`\textsubscript`, `\textcolor`, a word-wrapping highlight, `\texttt` for monospace fonts, `\fontsize` for sizes that differ from the body size |
| Lists | nested `itemize`/`enumerate` rebuilt from the flat list items (continued numbering kept), task items as check boxes, definition lists as `description` |
| Tables | `longtable` (page-breaking, header row repeated) or `tabular` where a `longtable` cannot go; ruled grid, `\multicolumn`/`\multirow` merges, column alignment (each column's most common, a cell that differs in a `\multicolumn{1}` of its own) and cell colours. A table wider than 16 columns continues below itself in bands of 16 |
| Links, citations | `\href` (scheme-checked), `\hyperref` for internal links whose target exists, `\cite{key}`; a bibliography (a numbered list anchored by the cited keys, as the parser reads one) as `thebibliography`, and citations with no entry in it reported (`CITATIONS_NOT_RESOLVED`) |
| Notes, comments | `\footnote` (deferred to `\footnotetext` inside a `tabular`), `\endnote` (`endnotes` package), review comments as LaTeX `%` comments, and a hidden `<!-- -->` note as `% <!-- ... -->` lines |
| Code, math | `lstlisting` for a language `listings` knows, `verbatim` otherwise; math as live LaTeX after a safety check, loading the packages its commands come from (`bm`, `mathtools`, `cancel`, `mathrsfs`, `siunitx`, `mhchem`, ...), KaTeX's and MathJax's own macros (`\R`) written as LaTeX, and a command nothing loaded defines printing its own name (reported as `CONTENT_NOT_REPRESENTABLE`) |
| Images, charts, embeds | `\includegraphics` at natural size, bounded to the line and page, of a PNG or JPEG carried inside the `.tex` (below) or, in a bundle, a file under `images/`; an image the source referred to only by a plain relative path keeps that path (`\includegraphics{figures/diagram}`; spaces are fine, and a URL's `%20`-style escapes are decoded to the file name), while a web image, or a path that is absolute or leaves the document's folder, becomes a link, since TeX cannot fetch the one and must not read the other; alternative text as the `alt` key; a caption as `\captionof`; charts as a data table; embeds as a link |
| Slides | `beamer` frames (the slide's first heading is the frame title, a lower-level heading right under it `\framesubtitle`, speaker notes become `\note`, long slides continue on another frame); a heading between slides is a `\section` |
| Page header/footer | `fancyhdr` |
| Title block | a heading styled `Title` (a Word title, or a parsed `\maketitle`) and the `Author`/`Date` lines right after it (and `Subtitle`/`Institute` in beamer): `\maketitle` where it stands (a `\titlepage` frame in beamer), printing only those lines |
| Metadata | the PDF metadata via `\hypersetup` (custom properties included); `\title`/`\author`/`\date` from the title block when there is one, else from the metadata |

**Safety.** LaTeX is a programming language, so every piece of document text is escaped, URLs are scheme-checked (the same allowlist as the DOCX/ODT generators) and percent-encoded, image paths are reduced to a safe file name inside `images/`, and a code block that contains its own end marker is not put in a verbatim environment. Math is the one place document content is emitted as live LaTeX; an expression that uses a command able to read or write files, run programs, or redefine commands (`\input`, `\write18`, `\openin`, `\catcode`, `\def`, ...), or that is structurally unbalanced, is written as literal text instead, with a `MATH_WRITTEN_AS_TEXT` warning. Compile untrusted output without `--shell-escape`, as you would any LaTeX you did not write.

**Images inside the `.tex`.** LaTeX reads images only from files, and a `.tex` can write a text file while it compiles (the kernel's `filecontents*` environment), so that is how a lone `.tex` carries its pictures: each PNG or JPEG becomes a block holding the image as a small PDF whose every byte is printable ASCII (the image keeps its own compression behind an ASCII85 layer, about a quarter larger), named after a hash of its content. Compiling writes that file beside the `.tex`, keeping a file of that name already there, which lets you replace a picture, and a plain `\includegraphics` reads it with every engine, with no shell escape or Ghostscript needed. Compile in the `.tex`'s folder, as every editor and Overleaf do: with `--output-directory`, pdfLaTeX and LuaLaTeX still find the files, but XeLaTeX and dvipdfmx look beside the `.tex`, so use `bundle` there. A JPEG, or a PNG a PDF can take as it is, is carried whatever its size; a PNG that must be decoded to be carried (transparency, a transparent palette colour, interlacing) is carried up to 16 megapixels, and a document's decoded images up to 256 megapixels in all, past which it is referenced as a file (and named by `IMAGES_NOT_BUNDLED`). Other image formats (GIF, BMP, TIFF, WebP, SVG, EMF) cannot be carried or drawn; they appear as a labelled box. The LaTeX parser reads these blocks back, so the images return as the original JPEG, or a PNG with the same pixels.

**Characters.** A character the default fonts lack (check marks, arrows, many math symbols, dingbats, unusual spaces) gets a fallback under every engine. A script pdfLaTeX cannot typeset at all (CJK, Cyrillic, emoji, ...) is shown there as a `[U+XXXX]` marker (Greek as math letters). XeLaTeX and LuaLaTeX set Greek, Cyrillic and CJK text in fonts TeX Live ships, where they are installed; for emoji or another script, add a `\setmainfont` that covers it. The table under **Engines** in [LaTeX Support](#latex-support) lists what each engine shows.

### CsvGeneratorConfig

Pass as `csvConfig` inside `GeneratorConfig`.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `sheets` | `string` | `''` | Sheet range to export: `'1'`, `'1-3'`, `'1,3'` (1-based). Empty = all sheets |
| `mergeSheets` | `boolean` | `true` | Merge all sheets into one CSV. When more than one sheet is merged, each is preceded by a `# Sheet: <name>` marker row and followed by a blank line (officeParser's own CSV parser reads these back as `comment` nodes; Excel/pandas read them as data rows). If `false`, returns a ZIP archive with one CSV per sheet and no markers |
| `columnDelimiter` | `string` | `','` | Output column delimiter |

### TextGeneratorConfig

Pass as `textConfig` inside `GeneratorConfig`.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `newlineDelimiter` | `string` | `'\n'` | String inserted between structural blocks |
| `preserveLayout` | `boolean` | `true` | Render tables with aligned columns using whitespace, and PDF pages as a spatial monospace grid |
| `renderNotes` | `boolean` | `true` | Append the collected footnote/endnote section |
| `pageSeparator` | `string` | `'\n'` | String inserted between PDF pages (set `'\f'` for a form feed) |

### metadataOverrides

Part of the common `GeneratorConfig` (not format-specific). Overrides the metadata embedded in
generated output, applied **per field** on top of `ast.metadata`, so setting one field leaves the
rest of the parsed metadata intact. `ast.metadata` itself is never mutated, so the same AST can be
generated repeatedly with different metadata.

| Field | Type | Written as |
|-------|------|-----------|
| `title` | `string` | HTML `<title>`/`<meta>`, EPUB `dc:title`, Markdown frontmatter, RTF `\title`, LaTeX `\title` and `pdftitle` |
| `author` | `string` | HTML `<meta name="author">`, EPUB `dc:creator`, frontmatter, RTF `\author`, LaTeX `\author` and `pdfauthor` |
| `description` | `string` | HTML `<meta name="description">`, EPUB `dc:description`, frontmatter, LaTeX PDF info `Description` |
| `subject` / `keywords` / `lastModifiedBy` | `string` | Where the destination format has a slot (LaTeX: `pdfsubject`, `pdfkeywords`, PDF info `LastModifiedBy`) |
| `created` / `modified` | `Date` | HTML `dcterms.*`, EPUB `dcterms:modified`, frontmatter, LaTeX `pdfcreationdate` / `\date` and `pdfmoddate` |
| `language` | `string` | HTML `lang`, EPUB/DOCX/ODT `dc:language`, PDF `/Lang`, LaTeX `pdflang` |
| `custom` | `Record<string, string \| number \| boolean \| Date>` | HTML `<meta name="custom:KEY">`, Markdown frontmatter, LaTeX PDF info entries |

```js
// Rebrand the output without touching the parsed document
const { value } = await ast.to('html', {
    metadataOverrides: { title: 'Q4 Report', author: 'Acme Inc', custom: { department: 'Finance' } },
});
```

#### Dates

`created` and `modified` take a `Date`. When `modified` is unset, officeParser uses the source
document's own `ast.metadata.modified`, falling back to the current time only if the document has
none. Dates outside the 1980-2099 range representable in a ZIP timestamp are clamped where EPUB
writes them onto zip entries.

```js
const { value } = await ast.to('epub', {
    metadataOverrides: { modified: new Date('2024-01-01T00:00:00Z') },
});
```

#### Not every format can represent every field

HTML `<meta>` tags and Markdown frontmatter are open vocabularies and accept anything. EPUB's OPF
is a closed Dublin Core vocabulary and RTF's `\info` group has a fixed set of control words, so a
`custom` entry has nowhere to go in either. Rather than dropping it silently, those generators
report it through `onWarning` (`OfficeWarningType.METADATA_NOT_REPRESENTABLE`) and continue; the
named fields still apply.

---

### OfficeConverterConfig

Configuration for `OfficeConverter.convert(file, format, config)`.

| Option | Type | Description |
|--------|------|-------------|
| `parseConfig` | `OfficeParserConfig` | Settings for the parsing phase |
| `generatorConfig` | `GeneratorConfig` | Settings for the generation phase |
| `onWarning` | `(issue: OfficeIssue) => void` | Global warning callback (overrides phase-specific ones) |

---

### ChunkingConfig

`ChunkingConfig` is a **discriminated union**: the available options depend on the `strategy` field.

#### Common Options (all strategies)

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `strategy` | `string` | `'document-structure'` | Chunking strategy |
| `stripWhitespace` | `boolean` | `true` | Trim leading/trailing whitespace from each chunk |
| `includeMetadata` | `boolean` | `true` | Include page/slide/heading metadata in each chunk |
| `addStartIndex` | `boolean` | `false` | Add `startIndex` character offset to chunk metadata |
| `lengthFunction` | `(text) => number` | `text.length` | Custom size measurer (e.g., token counter) |
| `sentenceBoundaryRegex` | `string \| RegExp` | `/[.!?。！？]/` | Custom regex for sentence boundary detection |
| `abbreviations` | `string[]` | common list | Abbreviations to skip when splitting on `.` |

#### `strategy: 'fixed-size'`

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `chunkSize` | `number` | `1000` | Maximum characters per chunk |
| `chunkOverlap` | `number` | `200` | Character overlap between consecutive chunks |
| `separators` | `string[]` | `['\n\n','\n',' ','']` | Ordered list of separators to try |

#### `strategy: 'document-structure'`

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `splitBy` | `string` | `'paragraph'` | `'paragraph'` · `'heading'` · `'page'` · `'slide'` · `'sheet'` |
| `maxChunkSize` | `number` | `1000` | Max characters per chunk (oversized units are split recursively) |
| `tableSplitStrategy` | `string` | `'row'` | `'row'` (repeats header in each chunk) or `'flatten'` |

#### `strategy: 'semantic'`

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `embeddingFunction` | `(text) => Promise<number[]>` | **required** | Async embedding function |
| `similarityThreshold` | `number` | `0.8` | Cosine similarity threshold; lower = fewer boundaries |
| `maxChunkSize` | `number` | `2000` | Max characters even if similarity stays high |
| `bufferSize` | `number` | `1` | Surrounding sentences used when computing similarity |
| `embeddingBatchSize` | `number` | `50` | Sentences per embedding API batch |
| `timeout` | `number` | `10000` | Timeout in milliseconds for individual embedding API calls. Set to `0` to disable. |

---

## OCR Scheduler & Resource Management

When `ocr: true` is set, `officeParser` maintains an intelligent **Smart Worker Pool** backed by Tesseract.js:

- **Dynamic Affinity**: Workers persist with their last-used language, avoiding re-initialization overhead.
- **LRU Re-allocation**: When a new language is requested and the pool is full, the Least Recently Used idle worker is re-initialized.
- **Auto-Termination**: Workers shut down after 10 seconds of inactivity (configurable via `ocrConfig.timeout.autoTerminate`).

### OCR Config (`ocrConfig`)

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `language` | `string` | `'eng'` | Tesseract language code(s), e.g. `'eng+fra'` |
| `workerPath` | `string` | `''` | Custom path to Tesseract worker script |
| `corePath` | `string` | `''` | Custom path to Tesseract core script |
| `langPath` | `string` | `''` | Custom path for language data files |
| `cacheMethod` | `'write' \| 'readOnly' \| 'refresh' \| 'none'` | `'write'` | Where Tesseract keeps the language data it loads (IndexedDB in the browser, a file in the working directory in Node). `'none'` neither reads nor stores it, for an application that must write nothing to the device |
| `preserveLayout` | `boolean` | `true` | Reconstruct the recognized text's 2-D page layout from Tesseract's per-word boxes (columns line up, right-hand text stays right), instead of flat reading-order text. Set `false` for the flat string |
| `timeout` | `OcrTimeoutConfig` | `{}` | Consolidated timeouts: `autoTerminate`, `workerLoad`, `recognition` |

See all language codes at [tesseract-ocr.github.io](https://tesseract-ocr.github.io/tessdoc/Data-Files).

### `OfficeParser.terminateOcr()`

In **short-lived scripts** (CLI tools, one-off automation), call `terminateOcr()` after processing to bypass the idle timer and exit immediately:

```js
const officeParser = require('officeparser');

const ast = await officeParser.parseOffice('file.pdf', { ocr: true, extractAttachments: true });
// ... process results ...
await officeParser.terminateOcr(); // immediate exit
```

> [!TIP]
> The built-in CLI (`npx officeparser ...`) handles this automatically.
> Only call it manually in your own scripts.

A parse still recognizing an image when `terminateOcr()` is called keeps its other content: that image's
recognition fails with the `OCR_TERMINATED` error code, which the parse reports as an `OCR_FAILED` warning.

---

## Browser Usage

Five bundles are available in the `dist/` directory:

| Bundle | Type | Description |
|--------|------|-------------|
| `officeparser.browser.mjs` | ESM | Standard ESM bundle for modern bundlers (Vite, Webpack, Next.js). |
| `officeparser.browser.iife.js` | IIFE | Standard UMD bundle for direct `<script>` inclusion (exposes global `officeParser`). |
| `officeparser.browser.slim.mjs` | ESM | Slim ESM bundle with Tesseract.js (OCR) stubbed out and remote CDN URLs removed. |
| `officeparser.browser.slim.iife.js` | IIFE | Slim UMD bundle with Tesseract.js (OCR) stubbed out and remote CDN URLs removed. |
| `officeparser.browser.native-pdf.mjs` | ESM | ESM bundle (the `officeparser/browser-native-pdf` subpath export) that leaves `pdf-lib` and `@pdf-lib/fontkit` external, for client-side native PDF generation. |

### Manifest V3 & Extension Compliance (Slim Bundles)
For strict browser environments like **Chrome/Edge Manifest V3 extensions**, remotely hosted code is forbidden. Use the **slim** bundles (`officeparser.browser.slim.mjs` or `officeparser.browser.slim.iife.js`) as they do not include default remote CDN urls or the Tesseract OCR engine.

### Content Security Policy
Every browser bundle loads under a policy with no `'unsafe-eval'` (`script-src 'self'`): none evaluates code from a string as it loads. What a policy still has to allow is what you choose to load: the pdf.js worker for PDF input (`pdfWorkerSrc`, a CDN by default, so host it yourself under `'self'`), and for OCR, Tesseract's worker and WebAssembly.

For PDF input, pdf.js also loads files when a document needs them: its CMaps (`pdfCMapUrl`) for text in fonts the PDF does not embed, chiefly Chinese, Japanese and Korean, and its image decoders (`pdfDecoderUrl`) for JBIG2, CCITT fax and JPEG 2000 pictures. To host them yourself, or to work offline, copy the `cmaps/` and `wasm/` folders of `node_modules/pdfjs-dist` (the version the library uses, `6.2.108`) beside your copy of the worker and point both options at them:

```js
const ast = await officeParser.parseOffice(pdfBytes, {
    pdfWorkerSrc: '/pdfjs/pdf.worker.min.mjs',
    pdfCMapUrl: '/pdfjs/cmaps/',
    pdfDecoderUrl: '/pdfjs/wasm/',
});
```

The library reads them with `fetch()`, so a policy allows their address in `connect-src`. The decoders are WebAssembly, which needs `'wasm-unsafe-eval'` in `script-src`; without it pdf.js's worker imports the slower JavaScript decoder beside each (`wasm/*_nowasm_fallback.js`), which `script-src` (or `worker-src`) must then allow. The library starts pdf.js's worker itself, from a small `blob:` module that passes pdf.js's warnings back (about pictures pdf.js leaves out) and then imports `pdfWorkerSrc`, so a policy allows `blob:` in `worker-src`, as pdf.js itself needs for a worker on another origin; where it cannot (or a `GlobalWorkerOptions.workerPort` is set), pdf.js starts its own worker as before and those pictures go unreported.

### ESM (Vite / Webpack / Next.js)

```js
import { OfficeParser } from 'officeparser';

const handleFile = async (event) => {
    const file = event.target.files[0];
    const buffer = await file.arrayBuffer();
    const ast = await OfficeParser.parseOffice(new Uint8Array(buffer));
    console.log((await ast.to('text')).value);
};
```

### Script Tag

```html
<script src="dist/officeparser.browser.iife.js"></script>
<script>
    async function handleFile(event) {
        const file = event.target.files[0];
        const buffer = await file.arrayBuffer();
        const ast = await officeParser.parseOffice(new Uint8Array(buffer));
        console.log((await ast.to('text')).value);
    }
</script>
```

> [!NOTE]
> **File paths don't work in the browser.** Always pass a `Buffer`, `ArrayBuffer`, or `Uint8Array`.
> Passing a path string will throw a descriptive `FEATURE_NOT_SUPPORTED_IN_BROWSER` error.

### PDF Worker Configuration

When parsing PDFs in the browser, a Web Worker is required. If `pdfWorkerSrc` is omitted, a jsDelivr CDN link is used automatically:

```js
// Uses default CDN worker:
const ast = await officeParser.parseOffice(pdfArrayBuffer);

// Or specify your own:
const ast = await officeParser.parseOffice(pdfArrayBuffer, {
    pdfWorkerSrc: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@6.2.108/build/pdf.worker.min.mjs'
});
```

> [!NOTE]
> The `pdfjs-dist` worker version must match the version bundled with `officeparser` (currently `pdfjs-dist@6.2.108`).

---

## Troubleshooting & Common Issues

| Symptom | Fix |
|---------|-----|
| Node.js process stays alive after finishing | Call `await officeParser.terminateOcr()` at end of script when OCR was used |
| `"Worker not found"` in browser for PDF | Verify `pdfWorkerSrc` points to `pdf.worker.min.mjs` matching `pdfjs-dist@6.2.108` |
| Japanese, Chinese or Korean PDF text missing (`PDF_ASSET_UNAVAILABLE` for a CMap) | Set `pdfCMapUrl` to a copy of `pdfjs-dist/cmaps/` you host (or leave `allowRemoteAssets` on) |
| Low OCR accuracy | Verify `ocrConfig.language` matches the document language; quality depends on image resolution |
| Out of memory on large Excel files | Call `await ast.to('text')` early and discard the AST object to allow garbage collection |
| `md`/`html`/`csv`/`tex` buffer not detected | Add `fileType: 'md'` (or `'html'`, `'csv'`, `'tex'`) to config (these formats have no magic bytes) |
| LaTeX images or `\input` chapters missing (`LATEX_FILE_NOT_FOUND`) | A `.tex` file does not contain the files it includes. Parse the whole project as a `.zip` (e.g. Overleaf's "Download source"), with `extractAttachments: true` for the images |
| `IMPROPER_BUFFERS` error | Usually means no file extension and no `fileType` hint was provided for a buffer input |
| PDF generation fails | The default engine needs the optional peer dependency (`npm install puppeteer`); or switch to the dependency-free native engine with `pdfConfig: { engine: 'native' }` (bundled `pdf-lib`, lighter output, no Chromium) |

For a full debugging guide, visit the [Live Documentation](https://harshankur.github.io/officeParser/#spec/debugging).

---

## Known Limitations

1. **ODT/ODS Charts**: May show inaccurate data when the chart references external cell ranges or uses complex layout-based data.
2. **PDF Images**: A picture the PDF stores as one plain JPEG (8-bit gray or RGB, no mask, `/Decode`, EXIF rotation or colour transform, and no markers pdf.js reads its colours by otherwise than a browser) keeps its own bytes (`pdf_image_p<page>_<n>.jpg`, `image/jpeg`); every other picture is re-encoded as PNG (`pdf_image_p<page>_<n>.png`, `image/png`) on both Node and the browser, since a PDF stores image data in formats no viewer opens directly. v7 emitted BMP; code that filters attachments by `.bmp` must be updated.
3. **PDF structure without tags**: Tables, lists and headings come from the PDF's tag tree when present. For untagged PDFs they are recovered geometrically, which is best-effort: complex float-beside-text layouts and tables without a tag tree may not separate perfectly. Column reading order, paragraphs and word spacing are handled on both paths. A filled form's values are read as text at their fields (a check box or radio button as ☒ or ☐; a field no viewer shows, of no size or off the page, is not read), and a date written one character per box as one value; an XFA-only form has no fields on its pages, so its values are not read.
4. **PDF text decoration and spans**: text colour, and the coloured rectangle a run sits on, are extracted by default (`pdfParserConfig.extractTextColor`); set it `false` to skip the extra operator-list pass on a throughput-focused text path. Underline and strikethrough are still not extracted: they are drawn as separate graphics operators rather than carried as text properties. Vertical (top-to-bottom) writing is read but not laid out spatially. Table cell `colSpan`/`rowSpan` are recovered best-effort on the tagged path, from the geometry of the empty placeholder cells the tag tree pads a merge with; untagged PDFs expose no spans.
5. **LaTeX input** is interpreted, not compiled: officeParser reads the document the way a converter does (structure, text, formatting, macros you define) rather than running TeX, so drawings (TikZ, pgfplots, `picture`) are omitted, packages are not loaded (their commands keep their text when unknown), and `\cite` keys stay as keys in the text (a bibliography lists each cited `.bib` entry as authors, title, where it appeared and year, not in the style the document chose).
6. **LaTeX output** is a faithful conversion, not a typesetting clone of the source: named font families are not carried over (only monospace), images in formats LaTeX cannot include (GIF, BMP, TIFF, WebP, SVG, EMF) are packaged but drawn as placeholders, `\cite` keys with no bibliography entry in the document are emitted as they are (add your own `.bib`; `CITATIONS_NOT_RESOLVED` names them), and very wide spreadsheets continue in 16-column bands.

---

## Security & Trust Boundary

`officeParser` is a **parsing, generation, and conversion** library. Like any parser, its whole job
is to open and interpret files it is handed, and those files may come from an untrusted source (a
user upload, an email attachment, a scraped document). A parser that accepts arbitrary documents
has a large and inherently open attack surface.

I do sanitize output and apply hardening where I can: injection escaping across the
HTML/CSS/URL/script/CSV/RTF/Markdown/LaTeX sinks, bounded macro expansion when reading LaTeX (which
never executes anything or touches the filesystem), Markdown parsing whose time grows linearly with the
text, decompression limits, decoding limits for the images LaTeX carries and reads, some resource and
recursion bounds, and SSRF precautions during PDF rendering. I fix issues as I learn of them (see
[CHANGELOG.md](CHANGELOG.md)). But this is **best-effort, not a guarantee.** A document parser of
this size will have attack vectors I have not found or have not yet addressed, and no amount of
internal hardening makes it safe to feed fully untrusted input without your own precautions.

**Treat this as garbage in, garbage out.** The library does its best with what you give it, but
responsibility for what you feed it, and for the effect a malicious file has on your system, rests
with you. If you process files from untrusted sources, sanitize and validate them at your own
boundary, and run the parsing in isolation appropriate to your threat model: sandboxing or
containerization, memory and time limits (a worker you can end: `abortSignal` cannot interrupt a
parse already running, except a PDF's in Node, whose pdf.js it stops mid-stream), a low-privilege process, and the `abortSignal` and `decompressionLimits`
options this library exposes. Do not rely on any single library's hardening
as a complete defense.

I am the sole maintainer, with no security team behind me. I take legitimate reports seriously and
will fix what I reasonably can, but I cannot commit to a response or resolution timeline. The
software is provided "AS IS" without warranty of any kind (see [LICENSE](LICENSE)). To report an
issue privately, see [SECURITY.md](SECURITY.md).

---

**npm**: [https://npmjs.com/package/officeparser](https://npmjs.com/package/officeparser)

**github**: [https://github.com/harshankur/officeParser](https://github.com/harshankur/officeParser)

## Support the Project

If `officeParser` has helped you save time, consider supporting its continued development. Your sponsorship helps maintain the project, add new features, and keep it robust for everyone.

<a href="https://github.com/sponsors/harshankur">
  <img src="https://img.shields.io/badge/Sponsor-GitHub-ea4aaa?style=for-the-badge&logo=github-sponsors" height="36">
</a>
<a href="https://www.buymeacoffee.com/harshankur">
  <img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" height="36" alt="Buy Me A Coffee">
</a>

## Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for details.

## License

This project is licensed under the MIT License; see the [LICENSE](LICENSE) file for details.

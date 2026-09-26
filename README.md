# officeParser: Universal Office Document Parser & Generator

A robust, strictly-typed **Node.js and Browser** library for parsing office files into a rich **Abstract Syntax Tree (AST)** and generating high-fidelity output in multiple formats.

**Parses:** [`docx`](https://en.wikipedia.org/wiki/Office_Open_XML) · [`pptx`](https://en.wikipedia.org/wiki/Office_Open_XML) · [`xlsx`](https://en.wikipedia.org/wiki/Office_Open_XML) · [`odt`](https://en.wikipedia.org/wiki/OpenDocument) · [`odp`](https://en.wikipedia.org/wiki/OpenDocument) · [`ods`](https://en.wikipedia.org/wiki/OpenDocument) · [`odg`](https://en.wikipedia.org/wiki/OpenDocument) · [`pdf`](https://en.wikipedia.org/wiki/PDF) · [`rtf`](https://en.wikipedia.org/wiki/Rich_Text_Format) · [`csv`](https://en.wikipedia.org/wiki/Comma-separated_values) · [`md`](https://en.wikipedia.org/wiki/Markdown) · [`html`](https://en.wikipedia.org/wiki/HTML) · [`epub`](https://en.wikipedia.org/wiki/EPUB) · [`tex`](https://en.wikipedia.org/wiki/LaTeX) (LaTeX, including Overleaf project zips)

**Generates:** `DOCX` · `ODT` · `LaTeX` · `Markdown` · `HTML` · `CSV` · `RTF` · `PDF` · `EPUB` · `Plain Text` · `RAG Chunks`

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
- **Templates / mail-merge** via `OfficeTemplate.render` (fill a DOCX template's `{{placeholders}}`, single or batch), and **ODG parsing** (LibreOffice Draw).
- **LaTeX in both directions** (8.1): `.tex` files and Overleaf project zips parse into the same AST as every other format (sections, lists, tables with merged cells, figures, math, footnotes, citations, cross-references, user macros, `beamer` slides), so LaTeX converts to DOCX, ODT, HTML, Markdown and the rest. And `to('tex')` turns any parsed document into LaTeX source that compiles unmodified with pdfLaTeX, XeLaTeX, LuaLaTeX, upLaTeX, pLaTeX and `latex` (the last three through dvipdfmx), presentations included (as `beamer` frames), carrying its images inside the one `.tex` file (or, in bundle mode, as files in a zip). See [LaTeX Support](#latex-support).

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
- **Images:** the CLI parses directly, so add `--extractAttachments` for images to reach *any* output (HTML/EPUB embed them, DOCX/ODT/Markdown/native-PDF include them, and LaTeX carries PNG and JPEG images inside the `.tex`, or with `--texConfig.bundle` packages them beside it). Without it, an image node has no bytes and HTML/Markdown emit a name-only `<img src="image1.png">` reference. (The `OfficeConverter`/`convert()` API auto-enables this; the CLI does not.)

### CLI Options

| Flag | Values | Default | Description |
|------|--------|---------|-------------|
| `--to` | `json\|text\|md\|html\|csv\|rtf\|pdf\|docx\|odt\|tex\|epub\|chunks` | `json` | Output format (`latex` is accepted as an alias of `tex`) |
| `--output` | path | (none) | Write output to a file |
| `--fileType` | `docx\|xlsx\|pptx\|odt\|odp\|ods\|odg\|pdf\|rtf\|csv\|md\|html\|epub\|tex` | (none) | Explicitly override input file type detection. Also accepts `latex`/`ltx`, the ODF template names `ott`/`ots`/`otp`/`otg`, and `zip` (parsed as whatever the archive holds) |
| `--ocr` | boolean | `false` | Enable OCR for images (also requires `--extractAttachments`; OCR runs over extracted images) |
| `--ocrConfig.language` | string | `eng` | Tesseract language(s), e.g. `deu` or `eng+fra` |
| `--ocrConfig.preserveLayout` | boolean | `true` | Keep the line layout of recognized text |
| `--password` | string | (none) | Password for an encrypted document (PDF, OOXML, or ODF) |
| `--extractAttachments` | boolean | `false` | Extract images/charts as Base64 |
| `--ignoreNotes` | boolean | `false` | Ignore footnotes/endnotes/speaker notes |
| `--ignoreComments` | boolean | `false` | Ignore inline comments |
| `--ignoreHeadersAndFooters` | boolean | `false` | Ignore headers and footers |
| `--ignoreSlideMasters` | boolean | `false` | Ignore slide masters |
| `--ignoreInternalLinks` | boolean | `false` | Ignore internal links |
| `--newlineDelimiter` | string | `\n` | Delimiter between lines/blocks in plaintext outputs |
| `--csvDelimiter` | string | `,` | Custom delimiter for CSV files |
| `--includeRawContent` | boolean | `false` | Include raw XML/RTF in nodes |
| `--serializeRawContent` | boolean | `true` | Include stringified XML in metadata |
| `--preserveXmlWhitespace` | boolean | `false` | Keep raw formatting space |
| `--includeBreakNodes` | boolean | `false` | Include break nodes (DOCX, ODF and LaTeX page breaks) |
| `--ignorePageGeometry` | boolean | `false` | Omit per-node bounding boxes and page dimensions |
| `--pdfParserConfig.useTags` | boolean | `true` | Use the PDF tag tree; `false` forces geometry-only structure |
| `--pdfParserConfig.detectColumns` | boolean | `true` | Multi-column reading-order detection |
| `--pdfParserConfig.pageRange` | string | all | Parse only the given pages, e.g. `1-3,7` |
| `--htmlParserConfig.preserveComments` | boolean | `false` | Keep HTML/EPUB `<!-- -->` comments as `comment` nodes |
| `--texParserConfig.today` | string | the date of the parse | What `\today` prints in LaTeX input, e.g. `--texParserConfig.today="May 1, 2024"` |
| `--pdfParserConfig.headingDetection` | `auto\|font-size\|off` | `auto` | How headings are inferred on the geometry path |
| `--pdfParserConfig.mergeHyphenatedWords` | boolean | `true` | Rejoin words hyphenated across line breaks |
| `--pdfParserConfig.normalizeText` | boolean | `true` | Unicode/ligature normalization of extracted text |
| `--pdfParserConfig.extractTextColor` | boolean | `true` | Record each run's fill colour in `formatting.color` (set `false` to skip for speed) |
| `--verbose` | boolean | `false` | Show full error stack traces and warning logs |
| `--includeFormatting` | boolean | `true` | Include formatting style map matching |
| `--renderMetadata` | boolean | `false` | Render metadata as visible content in the generated output |
| `--includeImages` | `image-only\|image+ocr-text\|ocr-text-only\|none` | `image-only` | How image nodes render. Works as `--includeImages=<mode>` or `--includeImages <mode>`; a bare `--includeImages` means `image-only` |
| `--maxInlineImageBytes` | number | `1500000` | Largest image HTML/Markdown inlines as a `data:` URI (`0` never inlines) |
| `--htmlConfig.containerWidth` | string \| number | `auto` | HTML output container width (e.g. `900px`, `100%`) |
| `--textConfig.pageSeparator` | string | `\n` | Separator written between pages in text output |
| `--pdfConfig.engine` | `html\|native` | `html` | PDF engine: Puppeteer (`html`) or pdf-lib (`native`, no browser) |
| `--texConfig.bundle` | boolean | `false` | LaTeX: write a zip of `main.tex` plus its `images/` instead of the `.tex` alone |
| `--texConfig.embedImages` | boolean | `true` | LaTeX: `false` refers to `images/` files instead of carrying PNG and JPEG images inside the `.tex` |
| `--texConfig.documentClass` | `auto\|article\|report\|book\|beamer` | `auto` | LaTeX document class (`auto` = `beamer` for presentations, `article` otherwise) |
| `--texConfig.standalone` | boolean | `true` | LaTeX: `false` writes the body only, for pasting into an existing document |
| ~~`--format`~~ | `json\|text\|md\|html\|csv\|rtf\|pdf\|docx\|odt\|tex\|epub\|chunks` | `json` | **Deprecated.** Use `--to` |
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

**Supported destinations:** `'text'` · `'md'` · `'html'` · `'csv'` · `'rtf'` · `'pdf'` · `'docx'` · `'odt'` · `'tex'` (alias `'latex'`) · `'epub'` · `'chunks'`

> [!NOTE]
> **PDF generation** uses a headless browser by default (`pdfConfig.engine: 'html'`), which needs the
> optional `puppeteer` peer dependency:
> ```bash
> npm install puppeteer
> ```
> Or choose `pdfConfig.engine: 'native'` to lay the document out directly with `pdf-lib`
> (`npm install pdf-lib`): no browser, and the only engine that produces a real PDF in the browser
> (import from `officeparser/browser-native-pdf` for the client-side path).
> See [PdfGeneratorConfig](#pdfgeneratorconfig).
>
> **EPUB generation with images** requires `extractAttachments: true` on the parse step that
> produced the AST. See [EPUB Support](#epub-support).

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

`OfficeTemplate.render()` (alias `renderTemplate`) fills a **DOCX template**'s `{{placeholder}}` tags from your data and returns a new `.docx`. It is not parsing or conversion: the template is copied and only the placeholders are substituted, so **all of the template's formatting, layout and structure are preserved**. Give it one data object for one document, or an array for a batch (one document per entry, a classic mail-merge). Think of it as a zero-dependency take on Adobe's Document Generation API.

```ts
import { OfficeTemplate } from 'officeparser';
import { writeFileSync } from 'fs';

// One document.
const bytes = await OfficeTemplate.render('invoice-template.docx', {
    data: { name: 'Acme Corp', amount: '$1,250.00', due: '2026-10-01' },
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

- **Run-aware.** Word often splits a typed `{{name}}` across several runs (`{{`, `na`, `me}}`); it is filled anyway, and a value takes the **formatting of the run its placeholder sat in** (a bold `{{amount}}` renders bold).
- **Everywhere text lives.** Placeholders in the body, headers, footers, footnotes/endnotes and comments are all filled. Values may contain `\n` (rendered as line breaks).
- **Placeholder names** are Unicode letters and digits plus `_`, `.`, `-` (e.g. `{{invoice.total}}`, `{{customer-name}}`). A name containing a space or other punctuation is not recognized and is left as literal text, so surrounding prose between the delimiters is never mistaken for a field.
- **Deterministic** output (pinned zip timestamps): the same template + data always renders byte-identical bytes.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `data` | `TemplateData \| TemplateData[]` | (required) | Field values. One object is one document; an array is one document per entry |
| `delimiters` | `{ start: string; end: string }` | `{{ }}` | Placeholder delimiters |
| `onMissing` | `'keep' \| 'empty' \| 'error'` | `'keep'` | A placeholder with no matching field: leave it, blank it, or reject with `TEMPLATE_FIELD_MISSING`. A field present but `null`/`undefined` always renders empty |
| `password` | `string` | (none) | Decrypt the template first, if it is itself password-protected |

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
├── metadata: { author, title, created, modified, keywords, customProperties, nativeProperties, styleMap, ... }
├── content: [ OfficeContentNode ]
│   ├── type: 'paragraph' | 'heading' | 'table' | 'list' | 'image' | 'chart' | 'comment' | 'admonition' | 'embed' | 'definitionList' | ...
│   ├── text: string  (concatenated text of node + all descendants)
│   ├── children: [ OfficeContentNode ]  (recursive structural children)
│   ├── notes: [ OfficeContentNode ]     (footnotes/endnotes/slide notes attached to this node)
│   ├── comments: [ OfficeContentNode ] (inline comments attached to this node)
│   ├── formatting: { bold, italic, underline, color, size, font, alignment, ... }
│   └── metadata: { level, listId, row, col, rowSpan, colSpan, backgroundColor, style, ... }
├── auxiliary?: OfficeAuxiliaryContent   (out-of-band layout elements)
│   ├── headers?: OfficeContentNode[]   (DOCX, PDF top band, ODT master pages, LaTeX fancyhdr)
│   ├── footers?: OfficeContentNode[]   (DOCX, PDF bottom band, ODT master pages, LaTeX fancyhdr)
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
└── to(format, config?)  (format: 'html'|'md'|'text'|'csv'|'rtf'|'pdf'|'docx'|'odt'|'tex'|'epub'|'chunks', returns { value, messages })
```

### `OfficeIssue`: Warning / Error Object

All warnings and errors (from both parsing and generation) use this shape:

```ts
interface OfficeIssue {
    type: 'warning' | 'info' | 'error';
    code: OfficeWarningType | OfficeErrorType;  // typed enum, e.g. 'OCR_FAILED'
    message: string;
    node?: OfficeContentNode;  // the node that triggered the issue, if any
    details?: any;             // original error or extra context
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
| `PDF_NO_TEXT_EXTRACTED` | parse | A PDF yielded ~no text (likely scanned). Set `ocr: true` + `extractAttachments: true`. |
| `PDF_STRUCT_TREE_UNRELIABLE` | parse | PDF tag tree absent/incomplete; structure recovered geometrically. |
| `PDF_TEXT_ENCODING_SUSPECT` | parse | PDF glyphs mostly unmappable (broken ToUnicode); text may be garbage. Consider OCR. |
| `PDF_OUTLINE_TRUNCATED` | parse | Bookmark outline hit the depth/size cap; `ast.auxiliary.outline` is partial. |
| `PDF_WORKER_MISSING` / `PDF_WORKER_FALLBACK` | parse | The pdf.js worker could not be loaded / a fallback was used (set `pdfWorkerSrc`). |
| `NO_WORKSHEETS_FOUND` / `NO_SLIDES_FOUND` | parse | A legitimately empty workbook/presentation. |
| `TABLE_CELL_LIMIT_EXCEEDED` | parse | A table exceeded `decompressionLimits.maxTableCells`; it was clamped. |
| `IMAGE_EXTRACTION_FAILED` / `IMAGE_PROCESSING_FAILED` / `ATTACHMENT_EXTRACTION_FAILED` | parse | An image/attachment could not be extracted or decoded; it was skipped or degraded. |
| `ANNOTATION_EXTRACTION_FAILED` / `CHART_DATA_EXTRACTION_FAILED` | parse | A PDF annotation / a chart's data could not be read. |
| `OCR_FAILED` | parse | OCR ran but failed for an image (see `details`). |
| `LATEX_CONSTRUCT_NOT_INTERPRETED` | parse | The LaTeX input used commands or environments the parser does not interpret (the message names them). Text inside them was kept; drawings such as TikZ pictures were omitted. |
| `LATEX_EXPANSION_LIMIT_REACHED` | parse | A LaTeX document hit a bound on macro expansion, file inclusion or nesting depth (the guard against expansion bombs, include cycles and runaway nesting); macros or files past it were not expanded, and content nested past it is kept as plain text. |
| `LATEX_FILE_NOT_FOUND` | parse | The LaTeX input includes files or images the parser could not read (a `.tex` holds only the files it carries in `filecontents` blocks). Parse the project as a `.zip` to include them; images are kept as path references. |
| `FILE_TYPE_DETECTION_FAILED` / `BUFFER_TYPE_MISMATCH` | parse | Type could not be sniffed / disagreed with the `fileType` hint. |
| `PASSWORD_REQUIRED` / `PASSWORD_INCORRECT` | parse | Encrypted input; supply `password`/`onPassword` (these also throw when parsing cannot continue). |
| `UNRECOGNIZED_CONFIG_OPTION` | config | A config key this version does not know (often a typo or a removed/renamed option); it had no effect. Raised for parser and generator configs, and for a `convert()` option placed at the top level instead of under `parseConfig`/`generatorConfig` (the message says where it belongs). |
| `INVALID_CONFIG_VALUE` | config | A generator option got a value it does not accept (a `documentClass`, paper `format` or margin, `pdfConfig.engine`, `htmlConfig.standalone.styles`, a Markdown dialect preset, a chunking strategy/`splitBy`/`tableSplitStrategy`) or `texParserConfig.today` is not a string; the option's default was used, and the message names the option, the value and what it accepts. |
| `CONTENT_NOT_REPRESENTABLE` | generate | A node type has no faithful form in the target format and was downgraded or omitted (e.g. math/embeds in DOCX/ODT, a table-less document to CSV). |
| `METADATA_NOT_REPRESENTABLE` | generate | A metadata field could not be represented in the target format. |
| `IMAGE_NOT_INLINED` | generate | An image over `maxInlineImageBytes` was referenced by name instead of inlined (Markdown / fragment HTML). |
| `IMAGES_NOT_BUNDLED` | generate | LaTeX output references image files the `.tex` does not carry (with `texConfig.embedImages: false`, an image other than a readable PNG or JPEG, or one past the decoding limits); the message names them. Ship them alongside, or set `texConfig.bundle: true`. It also names images the source referred to only by a relative path, with no image data (a `.tex` without its project, an HTML page's `<img src="pics/a.png">`): supply those at that path yourself, since not even a bundle can contain them. |
| `MATH_WRITTEN_AS_TEXT` | generate | A math expression used an unsafe LaTeX command (file access, shell, redefinition) or was malformed, so LaTeX output shows it as literal text instead of typesetting it. |
| `PDF_GENERATION_FAILED` | generate | PDF generation failed (e.g. Puppeteer missing for `engine: 'html'`). |
| `INVALID_STYLE_MAPPING` / `INVALID_STYLE_MAP_TAG` | generate | A `styleMap` entry/tag was invalid and ignored. |
| `TEMPLATE_UNSUPPORTED_FORMAT` / `TEMPLATE_FIELD_MISSING` | template | The template format is unsupported / a `{{field}}` had no value under `onMissing: 'error'`. |
| `PAGE_LOAD_FAILED` | parse | A PDF page could not be processed and was skipped (partial content). |
| `SHEET_RANGE_NOT_FOUND` | generate | A `csvConfig.sheets` range matched no sheet, so CSV output is empty. |
| `EMPTY_CHUNK_GENERATED` / `WHITESPACE_NODE_SKIPPED` / `BROWSER_GENERATION_LIMITATION` / `PERFORMANCE_TIP` / `DEPENDENCY_LOAD_FAILED` | generate | Diagnostic/informational notes from the chunking and PDF generators. |

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
| PDF  | – | footnotes/endnotes (tagged) | Y (top/bottom bands) | – | Y | – | Y (spans: tagged only) |
| RTF  | – | footnotes/endnotes | – (dropped) | – | Y | – | Y |
| HTML | `<!-- -->` become `comment` nodes* (opt-in: `preserveComments`) | footnotes/endnotes | – | – | Y (`data:` only) | – | Y |
| MD   | `<!-- -->` become `comment` nodes* | footnotes/endnotes | – | – | Y (`data:` only) | – | Y (HTML-table fallback) |
| CSV  | `#`-rows become `comment` nodes* | – | – | – | – | – | rows |
| EPUB | `<!-- -->` become `comment` nodes* (opt-in: `preserveComments`) | footnotes/endnotes | – | – | Y | – | Y |
| TEX  | Y (`% Comment (Author, date):` lines); `% <!-- -->` lines become `comment` nodes* | footnotes/endnotes | Y (`fancyhdr`) | – | Y (from a project zip) | – | Y |

Notes: comments land on `node.comments[]` (with `author`/`date`) except source-level comments, which
are *not* governed by `ignoreComments`: CSV's leading-`#` rows become top-level `comment` nodes, and
Markdown/HTML `<!-- ... -->` (and LaTeX `% <!-- ... -->` lines) become `comment` nodes (block or
inline) marked `metadata.sourceSyntax: 'html'` whose `text` is the raw comment body. The Markdown,
HTML and LaTeX generators keep those as comments; every other output format omits them (a hidden note
stays hidden). `ignoreNotes` /
`ignoreComments` / `ignoreHeadersAndFooters` / `ignoreSlideMasters` each remove the corresponding
column and are a no-op wherever it shows `–`. OCR (`ocr: true`) recognizes text from any extracted
image and therefore also needs `extractAttachments: true`.

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
        ├── metadata: { row, col, rowSpan?, colSpan? }
        └── children: [ Paragraph | List | Table | ... ]
```

- `row` / `col`: zero-based grid position
- `rowSpan` / `colSpan`: merged cells (DOCX, ODF, HTML, Markdown HTML-tables, and tagged PDF)
- Cells can contain nested tables

> [!NOTE]
> **Header rows.** A header row is flagged on its cells' metadata (`style: 'header'`, or `isHeader`),
> set by the parsers that mark one (DOCX `w:tblHeader`, ODF `table:table-header-rows`, HTML `<th>`/
> `<thead>`, tagged-PDF `TH`); generators read it through one shared heuristic (a marked row, or an
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
}
```

> [!NOTE]
> On a **content node**, an absent flag and `false` mean the same thing: the flag is simply not
> applied. On **`ast.metadata.styleMap`**, they differ: an absent flag means the style says nothing
> about that property (so it inherits), while `false` means the style explicitly turns it off
> (ODF's `fo:font-weight="normal"`, DOCX's `<w:b w:val="0"/>`). Code resolving inheritance itself
> must test `=== undefined`, not truthiness, or it will treat "explicitly off" as "unspecified".

### 6. Break Nodes (DOCX and ODF)

When `includeBreakNodes: true`, break elements appear as nodes:

```text
Break Node (type: 'break')
└── metadata: {
        breakType: 'textWrapping' | 'page' | 'column' | 'lastRenderedPage' | 'carriageReturn' | 'thematic',
        clear?: 'all' | 'left' | 'none' | 'right'
    }
```

> [!NOTE]
> Break nodes have no `text` property, but `ast.to('text')` automatically converts them to the configured newline delimiter.

> [!NOTE]
> `includeBreakNodes` gates DOCX/ODF only (where a break is otherwise invisible layout). **HTML and
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

- `admonition` round-trips through both Markdown (`> [!NOTE]` / `:::note ... :::`) and HTML (`<div class="admonition admonition-note" data-type="note">`)
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
| Admonitions | `> [!NOTE]` (also accepts GLFM `:::note ... :::` on import) | `type: 'admonition'`, `AdmonitionMetadata` |
| Footnotes | `Text[^1]` + `[^1]: Definition` | `type: 'note'`, keyed by footnote id |
| Definition lists | `Term\n: Definition` | `type: 'definitionList'` / `'definitionTerm'` / `'definitionDescription'` |
| Abbreviations | `*[HTML]: Hypertext Markup Language` | `TextMetadata.abbreviationTitle` |
| HTML comments | `<!-- note -->` on its own lines (may span lines, blank ones included) or inline in a run | `type: 'comment'`, `CommentMetadata.sourceSyntax: 'html'`, raw body in `text`; re-emitted byte-for-byte by the Markdown and HTML generators, kept by the LaTeX generator as `% <!-- ... -->` lines (which the LaTeX parser reads back), omitted by every other generator. `<!-->` and `<!--->` are empty comments; write `\<!--` for literal text |
| Attribute lists | `![alt](img.png){width=50% .centered}` | `ImageMetadata.width` / `.align`, `TableMetadata.align` |
| Citations | `[@smith2024]` | `TextMetadata.citationKey` |
| Wikilinks | `[[Page]]` / `[[Page\|Alias]]` | `TextMetadata.wikilink`, `.link`, `.linkType` |
| Highlight | `==text==` | `TextMetadata.backgroundColor` |
| Link/image titles | `[text](url "Title")` / `![alt](img.png "Title")` | `TextMetadata.title` / `ImageMetadata.title` |
| Inline/block math | `$E=mc^2$` / `` $$...$$ `` | `type: 'code'`, `CodeMetadata.math` (`'inline' \| 'block'`). `$$...$$` is display math wherever it is written: inside a paragraph, the paragraph is split around it; in a heading, list item, table cell, quote or note, where a block cannot go, it is inline math |
| Embeds | `::youtube[Label]{id=… width=… align=…}` / `::embed[Label]{src=… …}` (leaf directive; see `mdConfig.dialect.embeds`) | `type: 'embed'`, `EmbedMetadata` |
| Frontmatter arrays | `tags: [a, b]` or `tags: ["a","b"]` | Real array in `metadata.customProperties`/`nativeProperties` |
| MDX components (import-only) | `<Component prop="x">...</Component>` | Stripped; inner Markdown is kept. Never generated back. |

> [!NOTE]
> MDX/JSX stripping is one-directional (parse-only): officeParser never authors JSX back into Markdown.
> Wikilink enable/disable and citekey→bibliography resolution are application-level concerns; officeParser
> always parses/generates the syntax itself.

The same round-trip fidelity extends to HTML, so content saved from a rich-text editor survives a
save→reload cycle:

| HTML attribute | AST field | Notes |
|---|---|---|
| `data-width` / `data-align` / inline `style="width:…"` on `<img>` | `ImageMetadata.width` / `.align` | |
| `data-align` on `<table>` | `TableMetadata.align` | Emitted/parsed as per-column GFM markers (`:---`, `:---:`, `---:`); alignment rides `CellMetadata.align` |
| `title` on `<a>` / `<img>` | `TextMetadata.title` / `ImageMetadata.title` | Survives both directions (`[text](url "Title")` in Markdown) |
| `colspan` / `rowspan` on `<td>`/`<th>` | `CellMetadata.colSpan` / `.rowSpan` | Previously dropped on HTML import; merged cells now survive a save→reload cycle |
| `<div data-youtube-video="ID">` / `<iframe src="...youtube.com...">` | `type: 'embed'` | |
| `<ul data-type="taskList">` / `<li data-checked>` | `ListMetadata.isTask` / `.checked` | |

---

## EPUB Support

EPUB files are ZIP archives of XHTML content plus an OPF manifest. `EpubParser` unzips the archive,
resolves the spine's reading order from `content.opf`, and parses each XHTML document through the
existing `HtmlParser`, so EPUB content shares the same AST shape (and the same Markdown-dialect
fidelity above) as every other format. Dublin Core metadata (`dc:title`, `dc:creator`, `dc:description`,
`dc:subject`, `dc:date`, `dc:publisher`, `dc:language`, `dc:identifier`) maps into `ast.metadata` /
`ast.metadata.nativeProperties`, and cover art is exposed via `metadata.customProperties.coverImageName`.

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
**project zip**, such as Overleaf's "Download source": the main file (`main.tex`, else the top-level
`.tex` with a `\documentclass`) is read, and its `\input`/`\include`/`\subfile` files and
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
| `\textbf`, `\emph`, `\underline`, `\sout`, `\texttt`, `\textsc`, `\textcolor`, `\hl`, `\large`, `{\bfseries ...}`, accents, ligatures | formatted text runs |
| `itemize`, `enumerate` (nested, `\setcounter`), `description`, `\item[$\square$]` | lists, definition lists, task items |
| `tabular`, `tabularx`, `longtable`, `\multicolumn`, `\multirow`, `\cellcolor`, booktabs | `table` with `colSpan`/`rowSpan`, alignment and cell colours |
| `figure`/`table` floats, `\caption`, `\includegraphics` | images (bytes from a project zip or a `filecontents` block; a PDF that is only a picture, as the generator carries images, becomes that JPEG or PNG) and captions |
| `filecontents`, `filecontents*` | the file it writes, as compiling does (a file already there is kept unless `overwrite`), for `\input` and `\includegraphics` to read |
| `$...$`, `\[...\]`, `equation`, `align`, `gather` | math, as LaTeX, with your macros expanded |
| `\footnote`, `\endnote` | `note` nodes |
| `\href`, `\url`, `\ref`, `\eqref`, `\nameref`, `\hyperref` | links; references resolve to section, table, figure and equation numbers |
| `\cite` and `thebibliography` | citations and the bibliography list |
| `verbatim`, `lstlisting`, `minted`, `\verb` | `code` nodes with their language |
| `quote`, `quotation`, `verse` | quotes |
| `% <!-- ... -->` lines | a source `comment` node (the form the generator writes a Markdown/HTML `<!-- -->` in); other `%` comments are dropped |
| `fancyhdr` headers and footers | `ast.auxiliary` |
| `\maketitle`, beamer `\titlepage` | a title block where it stands: a `heading` styled `Title`, then `Author` and `Date` lines (beamer adds `Subtitle`), `\thanks` as footnotes |
| `\title`, `\author`, `\date`, `\hypersetup{pdf...}` | `ast.metadata` (`pdftitle`/`pdfauthor` win over `\title`/`\author`) |
| `beamer` frames, `\framesubtitle`, `\note`, overlays | `slide` nodes with speaker notes |
| `\newcommand`, `\renewcommand`, `\def`, `\newenvironment`, `\NewDocumentCommand`/`\NewDocumentEnvironment` (`m o O s t d D r R v b` arguments, `\IfBooleanTF`, `\IfNoValueTF`) | expanded; `expl3` code blocks are skipped |
| `\ifXeTeX`/`\ifLuaTeX`/`\ifPDFTeX` (iftex), `\newif` switches, `\ifdefined`, `\@ifpackageloaded`, `ifthen` booleans, etoolbox toggles | decided as pdfLaTeX would compile the document: only the branch taken is read (a test that cannot be decided, such as `\ifnum` on a counter, reads both and is reported) |
| `\newtheorem` (shared counters, numbering within sections, `\theoremstyle`), thmtools `\declaretheorem`, `proof` | "**Theorem 2.1** (Note)**.**" before the body (italic in the `plain` style), proofs ending in □; `\ref` reads theorem numbers. Theorem environments a class provides (Springer's llncs, svjour) need no `\newtheorem`, and one defined in a package the parser cannot see is headed but unnumbered; beamer's are titled blocks |
| babel/polyglossia: `\foreignlanguage`, `\text<language>`, `otherlanguage`, language environments, `\babeltags` | their text; the main language (babel's `main=` or last language, polyglossia's `\setdefaultlanguage`) becomes `metadata.language` unless `pdflang` states it |
| `\keywords`, `keywords`/`IEEEkeywords` environments; amsart `\address`, `\email`, `\urladdr`, `\subjclass`; IEEEtran author blocks, `\IEEEPARstart`; KOMA-Script `\minisec`, `\dictum`; `\epigraph` | `metadata.keywords` (and the printed "Keywords:" line), `nativeProperties`, authors' names in `metadata.author`, text |
| `\bf`, `\it`, `\tt`, `\sl`, `\sf`, `\rm`, `\sc`, `\documentstyle` (LaTeX 2.09) | formatted text runs, the document class |
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
| `onPassword` | `(reason: 'required' \| 'incorrect') => string \| undefined \| Promise<...>` | (none) | Called when an encrypted document needs a password `password` did not satisfy, so it can be supplied lazily or interactively (prompt, vault). Return a password to retry (capped), or `undefined` to reject as above. Works for every encryptable format (PDF/OOXML/ODF); mirrors pdf.js's `onPassword` |
| `ignoreNotes` | `boolean` | `false` | Ignore footnotes/endnotes (DOCX, ODT, RTF, PDF, HTML, Markdown, EPUB, LaTeX) and speaker notes (PPTX/ODP, LaTeX `beamer` `\note`). See the [capability matrix](#per-format-capability-matrix) |
| `ignoreComments` | `boolean` | `false` | Ignore comments/annotations, attached by default via `node.comments[]`. Applies to DOCX, XLSX, PPTX, every ODF type (ODT/ODS/ODP/ODG) and LaTeX (`% Comment (Author, date): text` lines). Source-level comments (CSV `#` rows, Markdown/HTML `<!-- -->`, LaTeX `% <!-- -->` lines) are not governed by it. See the [capability matrix](#per-format-capability-matrix) |
| `ignoreHeadersAndFooters` | `boolean` | `false` | Skip headers & footers (populated in `ast.auxiliary.headers/footers` by default). Extracted for DOCX, PDF, ODT and LaTeX (`fancyhdr`) only; a no-op for ODS/ODP/ODG, XLSX, PPTX and RTF. See the [capability matrix](#per-format-capability-matrix) |
| `ignoreSlideMasters` | `boolean` | `false` | Skip PPTX slide masters (populated in `ast.auxiliary.slideMasters` by default). PPTX only; ODP masters are not extracted |
| `extractAttachments` | `boolean` | `false` | Populate `ast.attachments` with Base64 images/charts |
| `ocr` | `boolean` | `false` | Run Tesseract OCR on images (requires `extractAttachments: true`) |
| `ocrConfig` | `OcrConfig` | see below | OCR settings (populated defaults: `language: 'eng'`, `preserveLayout: true`, worker/timeout defaults). See the [OCR section](#ocr-scheduler--resource-management) |
| `includeRawContent` | `boolean` | `false` | Attach raw XML/RTF source to each node |
| `serializeRawContent` | `boolean` | `true` | Re-serialize XML to clean strings (only if `includeRawContent: true`) |
| `preserveXmlWhitespace` | `boolean` | `false` | Preserve original XML whitespace during serialization |
| `includeBreakNodes` | `boolean` | `false` | Include typed break nodes: DOCX `w:br`/`w:cr`, ODF `fo:break-before`/`fo:break-after` and `text:soft-page-break`, LaTeX `\newpage`/`\clearpage` (a LaTeX line break, `\\` or `\newline`, is always kept inside its paragraph) |
| `ignoreInternalLinks` | `boolean` | `false` | Strip bookmarks and internal cross-references from AST (now honored for PDF too) |
| `ignorePageGeometry` | `boolean` | `false` | Omit the geometric layout data: per-node bounding boxes (`node.bounds`) and page dimensions. Currently produced by the PDF parser |
| `fileType` | `SupportedFileType \| FileTypeAlias \| null` | `null` | **Required for text-based binary data** (`'md'`, `'html'`, `'csv'`, `'tex'`) as these lack magic bytes. Also accepts the names the matching extensions route by (`FileTypeAlias`): `'latex'`/`'ltx'` for `tex`, the ODF template names `'ott'`/`'ots'`/`'otp'`/`'otg'`, and `'zip'` (parsed as whatever the archive holds). |
| `csvDelimiter` | `string` | `','` | Input delimiter when parsing CSV files |
| `decompressionLimits` | `DecompressionLimits` | `{ maxUncompressedBytes: 512MB, maxZipEntries: 10000, maxTableCells: 1000000 }` | **New**: Limits applied during ZIP extraction (and ODF cell expansion) to protect against excessive memory and resource usage |
| `htmlParserConfig` | `HtmlParserConfig` | `{}` | HTML/XHTML/EPUB parsing options **(and Markdown input: `preserveIframes`/`embedFolkForms` govern raw `<iframe>` blocks and folk embeds in `.md` too)**. `preserveAttributes` (`boolean`, default `false`): keep generic source attributes no typed field consumed on `node.htmlAttributes`. `preserveIframes` (`boolean \| string[]`, default `false`): preserve non-YouTube `<iframe>` embeds (otherwise dropped) as `embed` nodes: `true` for any, or a hostname allowlist; the src is scheme-checked on generation. `embedFolkForms` (`boolean`, default `false`): opt in to importing ambiguous folk embed forms (Obsidian `![](youtube-url)`, thumbnail-link) as YouTube embeds. `preserveComments` (`boolean`, default `false`): keep `<!-- ... -->` comments in HTML and EPUB input as `comment` nodes (`metadata.sourceSyntax: 'html'`) instead of dropping them; conditional comments (`<!--[if …]>`) are always dropped. The `data-html-comment` shape `sourceAttributes` emits is always read. |
| `pdfWorkerSrc` | `string` | CDN (jsDelivr) | Path/URL to `pdf.worker.min.mjs` (required in browser) |
| `pdfParserConfig` | `PdfParserConfig` | see below | PDF-specific options ([table below](#pdfparserconfig)) |
| `texParserConfig` | `TexParserConfig` | `{ today: '' }` | LaTeX options. `today` (`string`): what `\today` prints. `''` (default) prints the date of the parse, as LaTeX prints the date of the compile, in the document's language ("September 25, 2026" in English); set a fixed date, so that the same file parses the same way every day, or a placeholder of your own to find and replace later |
| `onWarning` | `(issue: OfficeIssue) => void` | (none) | Callback for non-fatal parsing issues |
| `abortSignal` | `AbortSignal \| null` | `null` | Optional signal to cancel parsing (rejects with AbortError). Once it fires the parse never resolves, even when it fires while OCR is recognizing an image |

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
| `headingDetection` | `'auto' \| 'font-size' \| 'off'` | `'auto'` | How heading levels are decided. `'auto'`: from tags when tagged, else a size/weight heuristic. `'font-size'`: re-level headings by the heuristic even on a tagged PDF (tables/lists stay tagged; a tagged heading is re-leveled by size and may be demoted). `'off'`: never emit headings |
| `pageRange` | `string` | `''` (all) | Restrict to given pages, e.g. `'1-3,7'`. Output keeps original page numbers |
| `normalizeText` | `boolean` | `true` | Unicode-normalize extracted text (expand ligatures, compose combining marks, regularize whitespace). Set `false` to preserve the raw source glyphs verbatim |
| `extractTextColor` | `boolean` | `true` | Extract each run's fill color into `formatting.color`. Recovered from the operator list; on by default (color is content like bold/font). Costs about 1.6x parse time on a text-heavy PDF, near-free when `extractAttachments`/`ocr` already fetch the operator list; set `false` to skip it. Pure black is left unset. Highlight annotations set `formatting.backgroundColor` regardless of this flag |

---

### GeneratorConfig (Common)

Options shared by all generator formats. Pass to `OfficeGenerator.generate(ast, format, config)` or `ast.to(format, config)`.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `includeFormatting` | `boolean` | `true` | Include bold/italic/colors/sizes in output (HTML, Markdown, DOCX, ODT, LaTeX, RTF; a no-op for text/CSV/chunks, which carry no run formatting) |
| `generateIds` | `boolean` | `true` | Slug-based heading anchors: `id` attributes on HTML headings, and a `{#slug}` suffix on Markdown headings (`# Title {#title}`, kramdown/Pandoc). Set `false` to omit both, useful when the Markdown is rendered by GFM/CommonMark, which show `{#slug}` as literal text. A top-level option (not under `mdConfig`/`htmlConfig`); it affects HTML, Markdown, DOCX, ODT and LaTeX (the formats that carry a heading anchor/bookmark id; in LaTeX a `\label`). |
| `renderMetadata` | `boolean` | `false` | Render title/author as a visible header block. Rendered by CSV, DOCX, HTML (and the Puppeteer PDF engine), EPUB, text, ODT, LaTeX (`\maketitle`, or a beamer title frame, unless the content has its own title block) and RTF; the native PDF engine and Markdown do not |
| `metadataOverrides` | `MetadataOverrides` | `{}` | Override the metadata embedded in the output, merged per field over `ast.metadata` |
| `includeImages` | `boolean \| 'image-only' \| 'image+ocr-text' \| 'ocr-text-only' \| 'none'` | `true` | How to render an image node. `true`=`'image-only'` (embed the image, no OCR text); `'image+ocr-text'` (image then its recognized/OCR text); `'ocr-text-only'` (OCR text, no image); `false`=`'none'` (omit). In plain-text output an image becomes an `[Image: name]` placeholder (plus OCR text for `'image+ocr-text'`), or just the OCR text for `'ocr-text-only'` |
| `maxInlineImageBytes` | `number` | `1500000` | Max decoded image size, in bytes, that is inlined as a `data:` URI (HTML/Markdown); the base64 URI itself is ~1/3 larger, so a scanned page cannot emit a multi-megabyte line that breaks downstream parsers. Under the default `image-only` mode an image over the cap renders its recognized/OCR text when it has any (multi-line OCR as a fenced block in Markdown), otherwise a compact name reference; Markdown still emits the `IMAGE_NOT_INLINED` warning. Plain text follows the same rule. **Standalone HTML always inlines**, whatever the cap: a self-contained document has nowhere else to resolve the image from. `0` never inlines, `Infinity` always inlines |
| `includeCharts` | `boolean` | `true` | Include charts: HTML renders an interactive Chart.js canvas, DOCX/ODT/LaTeX render the chart's data as a table, plain text and the native PDF engine render the chart's data text; Markdown and RTF render nothing for a chart. `false` omits charts in every generator |
| `ignoreInternalLinks` | `boolean` | `false` | Strip bookmarks and internal anchors from output (HTML, Markdown, DOCX, ODT, LaTeX, RTF) |
| `ignoreDefaultStyleMap` | `boolean` | `false` | Disable built-in style mappings (e.g., "Heading 1" → h1) |
| `styleMap` | `string[] \| StructuredStyleMapping[]` | `[]` | Custom semantic style mappings |
| `onNode` | `(node) => string \| false \| void` | (none) | Per-node callback for filtering, overriding, or mutating |
| `onWarning` | `(issue: OfficeIssue) => void` | (none) | Callback for non-fatal generation issues |
| `abortSignal` | `AbortSignal \| null` | `null` | Optional signal to cancel the generation operation (rejects with AbortError). Checked between steps by the HTML, Markdown, RTF, DOCX, ODT, EPUB, LaTeX, PDF and chunking generators; text and CSV generation run to completion |

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
| `chartJsSrc` | `string` | jsDelivr CDN | URL for the Chart.js library |
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
| `fallbackToHtml` | `boolean \| FallbackToHtmlConfig` | `true` | Use HTML tags for features Markdown cannot represent (underlines, merged table cells, embeds, etc.). Pass an object for per-feature control. `cellLineBreaks`/`itemLineBreaks` (default on) join multi-line table-cell / multi-paragraph list-item content with `<br>` instead of a space. `inlineFormatting` (default `false`, opt-in even when the boolean is `true`) additionally round-trips inline color/highlight/font-size as `<span style="...">` runs. |
| `dialect` | `MarkdownDialectPreset \| MarkdownDialectConfig` | `'extended'` | Which native syntax to emit for constructs that differ across targets (GitHub/GitLab/Obsidian/Pandoc/CommonMark). Each capability is typed by the syntax it selects (e.g. `strikethrough: 'tilde'`, `highlight: 'equals'`, `admonitions: 'blockquote'`), with `'none'` to turn it off. See [Markdown Dialect Support](#markdown-dialect-support). The old `boolean` toggles and admonition flavour names (`'github'`/`'gitlab'`/`'pandoc'`) still work but are deprecated. |

### PdfGeneratorConfig

Pass as `pdfConfig` inside `GeneratorConfig`. The default `'html'` engine requires the optional `puppeteer` peer dependency; the `'native'` engine requires the optional `pdf-lib` peer dependency instead. In the **prebuilt browser bundle** `pdf-lib` is stubbed so the bundle stays self-contained, so client-side native PDF export means importing from the dedicated entry **`officeparser/browser-native-pdf`** and installing `pdf-lib` (a consumer that bundles officeParser from source resolves `pdf-lib` normally).

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `engine` | `'html' \| 'native'` | `'html'` | `'html'`: render through a headless browser (Puppeteer in Node; native print in the browser) for maximum fidelity. `'native'`: lay the AST out directly with `pdf-lib` - no browser, runs in Node and the browser (real PDF bytes client-side), much lighter, but uses Standard-14 fonts and reflows rather than pixel-matching |
| `tagged` | `boolean` | `true` | Emit a tagged, PDF/UA-accessible PDF (structure tree). HTML engine only; the native engine ignores it |
| `outline` | `boolean` | `false` | Emit a heading-based bookmark outline. HTML engine only; the native engine ignores it |
| `format` | `string` | `'A4'` | Paper format (`'A4'`, `'Letter'`, `'Legal'`, etc.) |
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

> **Native engine (`engine: 'native'`) differences beyond fidelity:** it ignores the Puppeteer-only
> options above (and `tagged`/`outline`), uses a non-zero default `margin`, does not render charts,
> and does not apply `renderMetadata`. Choose it to avoid the headless browser; choose `'html'` for
> pixel-matching fidelity, charts, print headers/footers and PDF/UA tags.

### DocxGeneratorConfig

Pass as `docxConfig` inside `GeneratorConfig`. The DOCX generator writes a real WordprocessingML package (`.docx`) with zero extra dependencies, so it runs identically in Node and the browser and returns a `Uint8Array`. It reproduces headings, styled runs, tables (including merged cells), lists, images, hyperlinks, footnotes/endnotes, comments, headers/footers and metadata from any parsed source.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `format` | `PaperFormat` | `'A4'` | Page size for the document section (`w:pgSz`). Same names as `pdfConfig.format` (`A4`, `Letter`, `Legal`, `A3`, `Tabloid`, …; case-insensitive) |
| `landscape` | `boolean` | `false` | Landscape orientation (swaps page dimensions and sets `w:orient`) |
| `margin` | `object` | `{72,72,72,72}` | Page margins (`top`, `right`, `bottom`, `left`). Each is a number of points (1/72 inch) or a unit string (`'1in'`, `'2cm'`, `'36pt'`); 72 = Word's standard one inch |

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
| `format` | `PaperFormat` | `'A4'` | Page size for the page layout (`style:page-layout`). Same names as `pdfConfig.format` (`A4`, `Letter`, `Legal`, `A3`, `Tabloid`, …; case-insensitive) |
| `landscape` | `boolean` | `false` | Landscape orientation (swaps page dimensions and sets `style:print-orientation`) |
| `margin` | `object` | `{72,72,72,72}` | Page margins (`top`, `right`, `bottom`, `left`). Each is a number of points (1/72 inch) or a unit string (`'1in'`, `'2cm'`, `'36pt'`) |

```typescript
import { OfficeConverter } from 'officeparser';
import { writeFileSync } from 'fs';

// Any supported source → OpenDocument Text. convert() extracts the images itself; the CLI needs --extractAttachments.
const { value } = await OfficeConverter.convert('report.docx', 'odt');
writeFileSync('report.odt', value); // value is a Uint8Array
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
| `format` | `PaperFormat` | `'A4'` | Paper size, written as a `geometry` option. Same names as `pdfConfig.format`. Ignored by `beamer` |
| `landscape` | `boolean` | `false` | Landscape orientation. Ignored by `beamer` |
| `margin` | `object` | `{72,72,72,72}` | Page margins (`top`, `right`, `bottom`, `left`), each a number of points or a unit string (`'1in'`, `'2cm'`, `'36pt'`). Ignored by `beamer` |

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
| Tables | `longtable` (page-breaking, header row repeated) or `tabular` where a `longtable` cannot go; ruled grid, `\multicolumn`/`\multirow` merges, column alignment and cell colours. A table wider than 16 columns continues below itself in bands of 16 |
| Links, citations | `\href` (scheme-checked), `\hyperref` for internal links whose target exists, `\cite{key}` |
| Notes, comments | `\footnote` (deferred to `\footnotetext` inside a `tabular`), `\endnote` (`endnotes` package), review comments as LaTeX `%` comments, and a hidden `<!-- -->` note as `% <!-- ... -->` lines |
| Code, math | `lstlisting` for a language `listings` knows, `verbatim` otherwise; math as live LaTeX after a safety check |
| Images, charts, embeds | `\includegraphics` at natural size, bounded to the line and page, of a PNG or JPEG carried inside the `.tex` (below) or, in a bundle, a file under `images/`; an image the source referred to only by a plain relative path keeps that path (`\includegraphics{figures/diagram}`; spaces are fine, and a URL's `%20`-style escapes are decoded to the file name), while a web image, or a path that is absolute or leaves the document's folder, becomes a link, since TeX cannot fetch the one and must not read the other; charts as a data table; embeds as a link |
| Slides | `beamer` frames (the slide's first heading is the frame title, speaker notes become `\note`, long slides continue on another frame) |
| Page header/footer | `fancyhdr` |
| Title block | a heading styled `Title` (a Word title, or a parsed `\maketitle`) and the `Author`/`Date` lines right after it: `\maketitle` where it stands (a `\titlepage` frame in beamer), printing only those lines |
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

---

## Browser Usage

Five bundles are available in the `dist/` directory:

| Bundle | Type | Description |
|--------|------|-------------|
| `officeparser.browser.mjs` | ESM | Standard ESM bundle for modern bundlers (Vite, Webpack, Next.js). |
| `officeparser.browser.iife.js` | IIFE | Standard UMD bundle for direct `<script>` inclusion (exposes global `officeParser`). |
| `officeparser.browser.slim.mjs` | ESM | Slim ESM bundle with Tesseract.js (OCR) stubbed out and remote CDN URLs removed. |
| `officeparser.browser.slim.iife.js` | IIFE | Slim UMD bundle with Tesseract.js (OCR) stubbed out and remote CDN URLs removed. |
| `officeparser.browser.native-pdf.mjs` | ESM | ESM bundle (the `officeparser/browser-native-pdf` subpath export) that leaves `pdf-lib` external, for client-side native PDF generation. |

### Manifest V3 & Extension Compliance (Slim Bundles)
For strict browser environments like **Chrome/Edge Manifest V3 extensions**, remotely hosted code is forbidden. Use the **slim** bundles (`officeparser.browser.slim.mjs` or `officeparser.browser.slim.iife.js`) as they do not include default remote CDN urls or the Tesseract OCR engine.

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
2. **PDF Images**: Extracted and re-encoded as PNG (`pdf_image_p<page>_<n>.png`, `image/png`) on both Node and the browser, since a PDF stores image data in formats no viewer opens directly. v7 emitted BMP; code that filters attachments by `.bmp` must be updated.
3. **PDF structure without tags**: Tables, lists and headings come from the PDF's tag tree when present. For untagged PDFs they are recovered geometrically, which is best-effort: complex float-beside-text layouts and tables without a tag tree may not separate perfectly. Column reading order, paragraphs and word spacing are handled on both paths.
4. **PDF text decoration and spans**: text colour is extracted by default (`pdfParserConfig.extractTextColor`); set it `false` to skip the extra operator-list pass on a throughput-focused text path. Underline and strikethrough are still not extracted: they are drawn as separate graphics operators rather than carried as text properties. Vertical (top-to-bottom) writing is read but not laid out spatially. Table cell `colSpan`/`rowSpan` are recovered best-effort on the tagged path, from the geometry of the empty placeholder cells the tag tree pads a merge with; untagged PDFs expose no spans.
5. **LaTeX input** is interpreted, not compiled: officeParser reads the document the way a converter does (structure, text, formatting, macros you define) rather than running TeX, so drawings (TikZ, pgfplots, `picture`) are omitted, packages are not loaded (their commands keep their text when unknown), and `\cite` keys stay as keys unless the document has a `thebibliography` (`.bib` files are not processed).
6. **LaTeX output** is a faithful conversion, not a typesetting clone of the source: named font families are not carried over (only monospace), images in formats LaTeX cannot include (GIF, BMP, TIFF, WebP, SVG, EMF) are packaged but drawn as placeholders, `\cite` keys are emitted without a bibliography (add your own `.bib`), and very wide spreadsheets continue in 16-column bands.

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
containerization, memory and time limits, a low-privilege process, and the `abortSignal` and
`decompressionLimits` options this library exposes. Do not rely on any single library's hardening
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

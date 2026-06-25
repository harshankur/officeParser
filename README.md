# officeParser: Universal Office Document Parser & Generator

A robust, strictly-typed **Node.js and Browser** library for parsing office files into a rich **Abstract Syntax Tree (AST)** and generating high-fidelity output in multiple formats.

**Parses:** [`docx`](https://en.wikipedia.org/wiki/Office_Open_XML) · [`pptx`](https://en.wikipedia.org/wiki/Office_Open_XML) · [`xlsx`](https://en.wikipedia.org/wiki/Office_Open_XML) · [`odt`](https://en.wikipedia.org/wiki/OpenDocument) · [`odp`](https://en.wikipedia.org/wiki/OpenDocument) · [`ods`](https://en.wikipedia.org/wiki/OpenDocument) · [`pdf`](https://en.wikipedia.org/wiki/PDF) · [`rtf`](https://en.wikipedia.org/wiki/Rich_Text_Format) · [`csv`](https://en.wikipedia.org/wiki/Comma-separated_values) · [`md`](https://en.wikipedia.org/wiki/Markdown) · [`html`](https://en.wikipedia.org/wiki/HTML)

**Generates:** `Markdown` · `HTML` · `CSV` · `RTF` · `PDF` · `Plain Text` · `RAG Chunks`

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

## Table of Contents
- [Install](#install-via-npm)
- [Command Line Usage](#command-line-usage)
- [Quick Decision Guide](#quick-decision-guide)
- [Library Usage: Parsing](#library-usage-parsing)
  - [Async/Await](#asyncawait)
  - [Callback (Backward Compat)](#callback-backward-compat)
  - [File Buffers & ArrayBuffers](#file-buffers--arraybuffers)
  - [`ast.to()`: Generate from AST](#astto-generate-from-ast)
  - [`ast.toText()`: Quick Text Extraction](#asttotext-quick-text-extraction)
- [OfficeGenerator](#officegenerator)
- [OfficeConverter: One-Step API](#officeconverter-one-step-api)
- [Native RAG Chunking](#native-rag-chunking)
- [The AST Structure](#the-ast-structure)
- [Deep Dive: Document Components](#deep-dive-document-components)
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
  - [CsvGeneratorConfig](#csvgeneratorconfig)
  - [TextGeneratorConfig](#textgeneratorconfig)
  - [OfficeConverterConfig](#officeconverterconfig)
  - [ChunkingConfig](#chunkingconfig)
- [OCR Scheduler & Resource Management](#ocr-scheduler--resource-management)
- [Browser Usage](#browser-usage)
- [Troubleshooting & Common Issues](#troubleshooting--common-issues)
- [Known Limitations](#known-limitations)
- [Contributing](#contributing)

---

## Install via npm

```bash
npm i officeparser
```

---

## Command Line Usage

```bash
# Full AST as JSON (default)
npx officeparser /path/to/file.docx

# Plain text output
npx officeparser /path/to/file.docx --to=text

# Convert DOCX to Markdown and save
npx officeparser report.docx --to=md --output=report.md

# Convert PPTX to HTML (using a bare flag for ocr)
npx officeparser presentation.pptx --to=html --output=preview.html --ocr

# Convert XLSX to CSV with a custom delimiter
npx officeparser data.xlsx --to=csv --csvDelimiter=";"

# Generate RAG chunks
npx officeparser document.pdf --to=chunks

# Overriding file extension mapping
npx officeparser my_document --fileType=docx --to=json
```

### CLI Syntax
- **Values:** Flags can be passed as `--flag=value` or `--flag value`.
- **Booleans:** Bare flags imply `true` (e.g. `--ocr` is equivalent to `--ocr=true`). Negation flags start with `no-` (e.g. `--no-ocr` is equivalent to `--ocr=false`).
- **Nested Objects:** You can pass nested properties directly using JSON dot-notation (e.g. `--ocrConfig.language=fra` or `--htmlConfig.containerWidth=900px`).

### CLI Options

| Flag | Values | Default | Description |
|------|--------|---------|-------------|
| `--to` | `json\|text\|md\|html\|csv\|rtf\|pdf\|chunks` | `json` | Output format |
| `--output` | path | — | Write output to a file |
| `--fileType` | `docx\|xlsx\|pptx\|odt\|odp\|ods\|pdf\|rtf\|csv\|md\|html` | — | Explicitly override input file type detection |
| `--ocr` | boolean | `false` | Enable OCR for images |
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
| `--includeBreakNodes` | boolean | `false` | Include break nodes (DOCX only) |
| `--verbose` | boolean | `false` | Show full error stack traces and warning logs |
| `--includeFormatting` | boolean | `true` | Include formatting style map matching |
| `--renderMetadata` | boolean | `false` | Render metadata as visible content in the generated output |
| `--htmlConfig.containerWidth` | string \| number | `auto` | HTML output container width (e.g. `900px`, `100%`) |
| ~~`--format`~~ | `json\|text\|md\|html\|csv\|rtf\|pdf\|chunks` | `json` | **Deprecated.** Use `--to` |
| ~~`--toText`~~ | `true\|false` | `false` | **Deprecated.** Use `--to=text` |
| ~~`--ocrLanguage`~~ | string | `eng` | **Deprecated.** Use `--ocrConfig.language` |
| ~~`--putNotesAtLast`~~ | `true\|false` | `false` | **Deprecated and ignored.** Notes are attached structurally to their nodes. |
| ~~`--outputErrorToConsole`~~ | `true\|false` | `false` | **Deprecated.** Use `--verbose` |

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
officeParser.parseOffice('/path/to/file.docx', function(ast, err) {
    if (err) { console.error(err); return; }
    console.log(ast.toText());
});
```

### File Buffers & ArrayBuffers

Pass a `Buffer`, `ArrayBuffer`, or `Uint8Array` instead of a file path:

```js
const fs = require('fs');
const buffer = fs.readFileSync('/path/to/file.pdf');
const ast = await officeParser.parseOffice(buffer);
```

> [!IMPORTANT]
> **Text-based formats from buffers need a `fileType` hint.**
> Formats like `md`, `html`, and `csv` have no magic bytes, so the parser cannot
> auto-detect them from a buffer. You **must** provide `fileType` in that case:
> ```js
> const ast = await officeParser.parseOffice(markdownBuffer, { fileType: 'md' });
> ```

### Cancellation with AbortSignal

You can pass a standard `AbortSignal` (e.g. from an `AbortController`) to cancel an active parse operation. This is especially useful for setting request-level timeouts or canceling long-running parses (like large PDFs with OCR).

```js
const controller = new AbortController();

// Cancel parsing if it takes longer than 5 seconds
setTimeout(() => controller.abort(), 5000);

try {
    const ast = await officeParser.parseOffice('large_scanned_file.pdf', {
        abortSignal: controller.signal,
        ocr: true
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
> **Worker Cleanup on Abort**
> If an OCR job is actively running in the background when the signal is aborted, `officeParser` automatically terminates the Tesseract worker process immediately and removes it from the pool to prevent thread/memory leaks.

### Custom OCR Timeouts

To prevent the parser from hanging indefinitely due to slow network connections (when downloading Tesseract language datasets) or complex image processing, you can configure granular timeouts under `ocrConfig.timeout`.

```js
const ast = await officeParser.parseOffice('scanned_document.pdf', {
    ocr: true,
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

### `ast.to()`: Generate from AST

The preferred way to convert a parsed AST to another format. Returns a `ConversionResult`.

```ts
// ConversionResult shape:
// { value: string | Uint8Array | OfficeChunk[], messages: OfficeIssue[] }

const { value: markdown, messages } = await ast.to('md');
const { value: html }               = await ast.to('html', { includeFormatting: false });
const { value: chunks }             = await ast.to('chunks', { strategy: 'fixed-size', chunkSize: 800 });
const { value: pdfBytes }           = await ast.to('pdf'); // Uint8Array
```

### `ast.toText()`: Quick Text Extraction

> [!NOTE]
> `toText()` is **synchronous** and deprecated in favour of the async `ast.to('text')`.
> It remains available for backward compatibility.

```js
const text = ast.toText(); // synchronous, returns plain string
```

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

**Supported destinations:** `'text'` · `'md'` · `'html'` · `'csv'` · `'rtf'` · `'pdf'` · `'chunks'`

> [!NOTE]
> **PDF generation** requires the optional `puppeteer` peer dependency:
> ```bash
> npm install puppeteer
> ```

---

## OfficeConverter: One-Step API

`OfficeConverter.convert()` combines parsing and generation in a single call. It automatically syncs parser options from generator config (e.g., enables `extractAttachments` when images are requested).

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
> Do **not** put parser or generator options at the top level; only `onWarning` lives there.

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
├── type: 'docx' | 'pdf' | 'xlsx' | 'csv' | 'md' | ...  (11 formats)
├── metadata: { author, title, created, modified, keywords, customProperties, nativeProperties, styleMap, ... }
├── content: [ OfficeContentNode ]
│   ├── type: 'paragraph' | 'heading' | 'table' | 'list' | 'image' | 'chart' | 'comment' | ...
│   ├── text: string  (concatenated text of node + all descendants)
│   ├── children: [ OfficeContentNode ]  (recursive structural children)
│   ├── notes: [ OfficeContentNode ]     (footnotes/endnotes/slide notes attached to this node)
│   ├── comments: [ OfficeContentNode ] (inline comments attached to this node)
│   ├── formatting: { bold, italic, underline, color, size, font, alignment, ... }
│   └── metadata: { level, listId, row, col, rowSpan, colSpan, backgroundColor, style, ... }
├── auxiliary?: OfficeAuxiliaryContent   (out-of-band layout elements)
│   ├── headers?: OfficeContentNode[]   (DOCX headers)
│   ├── footers?: OfficeContentNode[]   (DOCX footers)
│   └── slideMasters?: OfficeContentNode[] (PPTX slide masters)
├── attachments: [ OfficeAttachment ]  (populated when extractAttachments: true)
│   ├── type: 'image' | 'chart'
│   ├── name: string
│   ├── mimeType: string
│   ├── data: string  (Base64)
│   ├── ocrText?: string  (if ocr: true)
│   └── chartData?: { title, dataSets, labels }
├── warnings: OfficeIssue[]  (non-fatal issues from the parsing phase)
├── to(format, config?)  (format: 'html'|'md'|'text'|'csv'|'rtf'|'pdf'|'chunks', returns { value, messages })
└── ~~toText()~~             (Deprecated: use .to('text') instead)
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
- `rowSpan` / `colSpan`: merged cells (primarily ODF formats)
- Cells can contain nested tables

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

### 6. Break Nodes (DOCX only)

When `includeBreakNodes: true`, break elements appear as nodes:

```text
Break Node (type: 'break')
└── metadata: {
        breakType: 'textWrapping' | 'page' | 'column' | 'lastRenderedPage' | 'carriageReturn',
        clear?: 'all' | 'left' | 'none' | 'right'
    }
```

> [!NOTE]
> Break nodes have no `text` property, but `ast.toText()` and `ast.to('text')` automatically convert them to the configured newline delimiter.

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

// Footnotes and endnotes (DOCX/RTF) can be deeply nested, so we traverse recursively:
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
> `putNotesAtLast` is **deprecated**. Notes are always attached via `node.notes`; this flag has no effect and will be removed in a future major version.

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
| `newlineDelimiter` | `string` | `'\n'` | Delimiter inserted between lines in text output |
| `ignoreNotes` | `boolean` | `false` | Ignore footnotes/endnotes (DOCX, RTF) and speaker notes (PPTX/ODP) |
| `ignoreComments` | `boolean` | `false` | **New**: Ignore inline comments/annotations (DOCX, XLSX, PPTX), attached by default via `node.comments[]` |
| `ignoreHeadersAndFooters` | `boolean` | `false` | **New**: Skip DOCX headers & footers (populated in `ast.auxiliary.headers/footers` by default) |
| `ignoreSlideMasters` | `boolean` | `false` | **New**: Skip PPTX slide masters (populated in `ast.auxiliary.slideMasters` by default) |
| ~~`putNotesAtLast`~~ | `boolean` | `false` | **Deprecated**: Notes are now attached via `node.notes[]`. This flag has no effect |
| `extractAttachments` | `boolean` | `false` | Populate `ast.attachments` with Base64 images/charts |
| `ocr` | `boolean` | `false` | Run Tesseract OCR on images (requires `extractAttachments: true`) |
| `ocrConfig` | `OcrConfig` | `{}` | OCR worker pool settings (see [OCR section](#ocr-scheduler--resource-management)) |
| `includeRawContent` | `boolean` | `false` | Attach raw XML/RTF source to each node |
| `serializeRawContent` | `boolean` | `true` | Re-serialize XML to clean strings (only if `includeRawContent: true`) |
| `preserveXmlWhitespace` | `boolean` | `false` | Preserve original XML whitespace during serialization |
| `includeBreakNodes` | `boolean` | `false` | Include `w:br` / `w:cr` as typed break nodes (DOCX only) |
| `ignoreInternalLinks` | `boolean` | `false` | Strip bookmarks and internal cross-references from AST |
| `fileType` | `SupportedFileType \| null` | `null` | **Required for text-based binary data** (`'md'`, `'html'`, `'csv'`) as these lack magic bytes. |
| `csvDelimiter` | `string` | `','` | Input delimiter when parsing CSV files |
| `decompressionLimits` | `DecompressionLimits` | `{ maxUncompressedBytes: 512MB, maxZipEntries: 10000 }` | **New**: Limits applied during ZIP extraction to protect against excessive memory and resource usage |
| `pdfWorkerSrc` | `string` | CDN (jsDelivr) | Path/URL to `pdf.worker.min.mjs` (required in browser) |
| `onWarning` | `(issue: OfficeIssue) => void` | — | Callback for non-fatal parsing issues |
| `abortSignal` | `AbortSignal \| null` | `null` | Optional signal to cancel parsing (rejects with AbortError) |
| ~~`outputErrorToConsole`~~ | `boolean` | `false` | **Deprecated.** Use `onWarning` instead |

---

### GeneratorConfig (Common)

Options shared by all generator formats. Pass to `OfficeGenerator.generate(ast, format, config)` or `ast.to(format, config)`.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `includeFormatting` | `boolean` | `true` | Include bold/italic/colors/sizes in output |
| `generateIds` | `boolean` | `true` | Add slug-based `id` attributes to headings |
| `renderMetadata` | `boolean` | `false` | Render title/author as visible header block |
| `includeImages` | `boolean` | `true` | Include image nodes in output |
| `includeCharts` | `boolean` | `true` | Include interactive charts (HTML only) |
| `ignoreInternalLinks` | `boolean` | `false` | Strip bookmarks and internal anchors from output |
| `ignoreDefaultStyleMap` | `boolean` | `false` | Disable built-in style mappings (e.g., "Heading 1" → h1) |
| `styleMap` | `string[] \| StructuredStyleMapping[]` | `[]` | Custom semantic style mappings |
| `onNode` | `(node) => string \| false \| void` | — | Per-node callback for filtering, overriding, or mutating |
| `onWarning` | `(issue: OfficeIssue) => void` | — | Callback for non-fatal generation issues |
| `abortSignal` | `AbortSignal \| null` | `null` | Optional signal to cancel the generation operation (rejects with AbortError) |

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
| `standalone` | `boolean` | `true` | Wrap output in a full `<html>` document with CSS |
| `chartJsSrc` | `string` | jsDelivr CDN | URL for the Chart.js library |
| `containerWidth` | `string \| number` | `'auto'` | Max width of the content container. Positive number (px), CSS length string (`'900px'`, `'100%'`, `'60vw'`), or `'auto'`. Invalid values fall back to `'auto'` with an `INVALID_CONTAINER_WIDTH` warning |
| `customCss` | `string` | `''` | Raw CSS injected into the `<style>` block; use this to override built-in styles |
| `injections.headStart` | `string` | `''` | Raw HTML injected after `<head>` |
| `injections.headEnd` | `string` | `''` | Raw HTML injected before `</head>` |
| `injections.bodyStart` | `string` | `''` | Raw HTML injected after `<body>` |
| `injections.bodyEnd` | `string` | `''` | Raw HTML injected before `</body>` |

### MdGeneratorConfig

Pass as `mdConfig` inside `GeneratorConfig`.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `fallbackToHtml` | `boolean` | `true` | Use HTML tags for features Markdown cannot represent (underlines, merged table cells, etc.) |

### PdfGeneratorConfig

Pass as `pdfConfig` inside `GeneratorConfig`. Requires the optional `puppeteer` peer dependency.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `format` | `string` | `'A4'` | Paper format (`'A4'`, `'Letter'`, `'Legal'`, etc.) |
| `width` | `string \| number` | `''` | Paper width (e.g., `'5in'`, `'3cm'`) or pixels |
| `height` | `string \| number` | `''` | Paper height (e.g., `'5in'`, `'3cm'`) or pixels |
| `landscape` | `boolean` | `false` | Landscape page orientation |
| `printBackground` | `boolean` | `true` | Print background graphics |
| `margin` | `object` | `{0,0,0,0}` | Page margins (`top`, `right`, `bottom`, `left`) |
| `displayHeaderFooter` | `boolean` | `false` | Show print header/footer |
| `headerTemplate` | `string` | `''` | HTML template for the print header |
| `footerTemplate` | `string` | `''` | HTML template for the print footer |
| `scale` | `number` | `1` | Rendering scale factor |
| `launchOptions` | `object` | headless defaults | Puppeteer launch options (e.g., `executablePath`) |
| `timeout` | `number` | `30000` | PDF rendering timeout in milliseconds. Set to `0` to disable. |

### CsvGeneratorConfig

Pass as `csvConfig` inside `GeneratorConfig`.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `sheets` | `string` | `''` | Sheet range to export: `'1'`, `'1-3'`, `'1,3'` (1-based). Empty = all sheets |
| `mergeSheets` | `boolean` | `true` | Merge all sheets into one CSV. If `false`, returns a ZIP archive |
| `columnDelimiter` | `string` | `','` | Output column delimiter |

### TextGeneratorConfig

Pass as `textConfig` inside `GeneratorConfig`.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `newlineDelimiter` | `string` | `'\n'` | String inserted between structural blocks |
| `preserveLayout` | `boolean` | `true` | Render tables with aligned columns using whitespace |

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
- **Auto-Termination**: Workers shut down after 10 seconds of inactivity (configurable via `ocrConfig.autoTerminateTimeout`).

### OCR Config (`ocrConfig`)

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `language` | `string` | `'eng'` | Tesseract language code(s), e.g. `'eng+fra'` |
| `workerPath` | `string` | `''` | Custom path to Tesseract worker script |
| `corePath` | `string` | `''` | Custom path to Tesseract core script |
| `langPath` | `string` | `''` | Custom path for language data files |
| `timeout` | `OcrTimeoutConfig` | `{}` | Consolidated timeouts: `autoTerminate`, `workerLoad`, `recognition` |
| ~~`autoTerminateTimeout`~~ | `number` | `10000` | **Deprecated.** Use `timeout.autoTerminate` instead |

See all language codes at [tesseract-ocr.github.io](https://tesseract-ocr.github.io/tessdoc/Data-Files).

### `OfficeParser.terminateOcr()`

In **short-lived scripts** (CLI tools, one-off automation), call `terminateOcr()` after processing to bypass the idle timer and exit immediately:

```js
const officeParser = require('officeparser');

const ast = await officeParser.parseOffice('file.pdf', { ocr: true });
// ... process results ...
await officeParser.terminateOcr(); // immediate exit
```

> [!TIP]
> The built-in CLI (`npx officeparser ...`) handles this automatically.
> Only call it manually in your own scripts.

---

## Browser Usage

Two bundles are available in the `dist/` directory:

| Bundle | Usage |
|--------|-------|
| `officeparser.browser.mjs` | ESM, use with `import` statements or modern bundlers (Vite, Webpack, Next.js) |
| `officeparser.browser.iife.js` | IIFE, use with a `<script>` tag; exposes the global `officeParser` object |

### ESM (Vite / Webpack / Next.js)

```js
import { OfficeParser } from 'officeparser';

const handleFile = async (event) => {
    const file = event.target.files[0];
    const buffer = await file.arrayBuffer();
    const ast = await OfficeParser.parseOffice(new Uint8Array(buffer));
    console.log(ast.toText());
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
        console.log(ast.toText());
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
    pdfWorkerSrc: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@5.6.205/build/pdf.worker.min.mjs'
});
```

> [!NOTE]
> The `pdfjs-dist` worker version must match the version bundled with `officeparser` (currently **5.6.205**).

---

## Troubleshooting & Common Issues

| Symptom | Fix |
|---------|-----|
| Node.js process stays alive after finishing | Call `await officeParser.terminateOcr()` at end of script when OCR was used |
| `"Worker not found"` in browser for PDF | Verify `pdfWorkerSrc` points to `pdf.worker.min.mjs` matching version `5.6.205` |
| Low OCR accuracy | Verify `ocrConfig.language` matches the document language; quality depends on image resolution |
| Out of memory on large Excel files | Call `ast.toText()` early and discard the AST object to allow garbage collection |
| `md`/`html`/`csv` buffer not detected | Add `fileType: 'md'` (or `'html'`, `'csv'`) to config (these formats have no magic bytes) |
| `IMPROPER_BUFFERS` error | Usually means no file extension and no `fileType` hint was provided for a buffer input |
| PDF generation fails | Install the optional peer dependency: `npm install puppeteer` |

For a full debugging guide, visit the [Live Documentation](https://harshankur.github.io/officeParser/#spec/debugging).

---

## Known Limitations

1. **ODT/ODS Charts**: May show inaccurate data when the chart references external cell ranges or uses complex layout-based data.
2. **PDF Images (Browser)**: Extracted as BMP files for cross-platform compatibility. Conversion is automatic.

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

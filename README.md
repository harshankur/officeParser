# officeParser 📄🚀 - The Most Versatile Office Parser & Generator

A robust, strictly-typed Node.js and Browser library for parsing and generating office files. It not only extracts content from [`docx`](https://en.wikipedia.org/wiki/Office_Open_XML), [`pptx`](https://en.wikipedia.org/wiki/Office_Open_XML), [`xlsx`](https://en.wikipedia.org/wiki/Office_Open_XML), [`odt`](https://en.wikipedia.org/wiki/OpenDocument), [`odp`](https://en.wikipedia.org/wiki/OpenDocument), [`ods`](https://en.wikipedia.org/wiki/OpenDocument), [`pdf`](https://en.wikipedia.org/wiki/PDF), [`rtf`](https://en.wikipedia.org/wiki/Rich_Text_Format), [`csv`](https://en.wikipedia.org/wiki/Comma-separated_values), [`md`](https://en.wikipedia.org/wiki/Markdown), and [`html`](https://en.wikipedia.org/wiki/HTML) into a rich Abstract Syntax Tree (AST), but also provides a powerful generation engine to convert that AST into formats like **Markdown**, **HTML**, **CSV**, **RTF**, **Text**, **PDF**, and **JSON**, including native **RAG-focused chunking** support.

[![npm version](https://badge.fury.io/js/officeparser.svg)](https://badge.fury.io/js/officeparser)
[![Total Downloads](https://img.shields.io/npm/dt/officeparser.svg)](https://www.npmjs.com/package/officeparser)
[![Weekly Downloads](https://img.shields.io/npm/dw/officeparser.svg)](https://www.npmjs.com/package/officeparser)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

### 🌟 [Live Interactive AST Visualizer & Documentation](https://harshankur.github.io/officeParser/) 🌟
*Test any office file in your browser and see the extracted AST, text, and preview in real-time which is rebuilt from the AST!*

**What you can do there:**
- **AST Visualizer**: Upload any office file and inspect the hierarchical AST structure, metadata, and raw content.
- **Config Configurator**: Tweak parsing options (like `ignoreNotes`, `ocr`, `newlineDelimiter`) and see the results instantly.
- **Debugging**: Use the visualizer to debug parsing issues by inspecting exactly how nodes are interpreted.
- **Format Specs**: Read detailed specifications for the AST structure and configuration options.

---


---

### 📝 [Changelog](CHANGELOG.md)
*Detailed release notes and the full history of updates are available in the project changelog.*

---

## Install via npm

```bash
npm i officeparser
```

## Command Line usage
You can use `officeparser` directly from the terminal to extract content as JSON AST, plain text, or generate new formats like Markdown and HTML.

```bash
# Get full AST as JSON (default)
npx officeparser /path/to/officeFile.docx

# Get plain text only
npx officeparser /path/to/officeFile.docx --toText=true

# Generate Markdown file
npx officeparser report.docx --format=md --output=report.md

# Generate HTML file with specific output
npx officeparser presentation.pptx --format=html --output=preview.html

# Convert spreadsheet to CSV
npx officeparser data.xlsx --format=csv
```

### Config Options:
- `--format=[json|text|md|html|csv|rtf|pdf|chunks]` The output format. Default is `json`.
- `--output=[path]`                    Optional file path to write the output to.
- `--toText=[true|false]`              Legacy flag to output only plain text. Use `--format=text` instead.
- `--ignoreNotes=[true|false]`          Flag to ignore notes from files like PowerPoint. Default is false.
- `--newlineDelimiter=[delimiter]`      The delimiter to use for new lines. Default is `\n`.
- `--putNotesAtLast=[true|false]`       Flag to collect notes at the end of files like PowerPoint. Default is false.
- `--outputErrorToConsole=[true|false]` **(Deprecated)** Flag to output errors to the console. Use `onWarning` callback in library usage.
- `--extractAttachments=[true|false]`   Flag to extract images/charts as Base64. Default is false.
- `--ocr=[true|false]`                  Flag to enable OCR for extracted images. Default is false.
- `--includeRawContent=[true|false]`    Flag to include raw XML/RTF content in nodes. Default is false.
- `--includeBreakNodes=[true|false]`    Flag to include break nodes. Currently only available for DOCX documents.
- `--verbose=[true|false]`              Show full error stack traces.


## Library Usage
In **v7.0.0**, the library has evolved into a dual-purpose **Parser** and **Generator**. You can first parse any office file into a structured AST and then use the `OfficeGenerator` to transform that AST into various formats or chunks.

### Getting Started (Async/Await)
```js
const officeParser = require('officeparser');

async function parseMyFile() {
    try {
        // parseOffice returns an OfficeParserAST object
        const ast = await officeParser.parseOffice("/path/to/officeFile.docx");
        
        // Use the built-in helper to get plain text (similar to old behavior)
        const text = ast.toText();
        console.log(text);

        // Access structured content
        console.log(ast.content);  // Array of hierarchical nodes (paragraphs, tables, etc.)
        console.log(ast.metadata); // Document properties (author, title, etc.)
    } catch (err) {
        console.error(err);
    }
}
```

### Helper Function for Text Extraction (Modern simple way)
If you only need the text and want to maintain a simple one-liner, you can use this pattern:
```js
// Simple helper to get text directly
const getText = async (file, config) => (await officeParser.parseOffice(file, config)).toText();

// usage
const text = await getText("/path/to/officeFile.docx");
console.log(text);
```

## Using the OfficeGenerator
The `OfficeGenerator` is a powerful tool to convert your AST into human-readable formats or structured data.

```typescript
import { OfficeParser, OfficeGenerator } from 'officeparser';

const ast = await OfficeParser.parseOffice('report.docx');

// 1. Convert to Markdown
const md = await OfficeGenerator.generate(ast, 'md');
console.log(md.value);

// 2. Convert to HTML with structured style mapping (Recommended)
const html = await OfficeGenerator.generate(ast, 'html', {
    includeFormatting: true,
    styleMap: [
        { 
            selector: { nodeType: 'paragraph', attributes: { style: 'Heading 1' } }, 
            output: { tag: 'h1', classes: ['main-title'] } 
        }
    ]
});
console.log(html.value);

// 3. Convert to CSV (for spreadsheets)
const csv = await OfficeGenerator.generate(ast, 'csv');
console.log(csv.value);
```

## The New "One-Step" API: `OfficeConverter`
In **v7.0.0**, we introduced the `OfficeConverter.convert` method. This is the new high-level API designed for one-step transformations where you don't need to manually interact with the AST. It automatically handles parser and generator configuration synchronization.

```typescript
import { OfficeConverter } from 'officeparser';

// One-step conversion from DOCX to Markdown
const result = await OfficeConverter.convert('report.docx', 'md');
console.log(result.value); // The generated Markdown string
console.log(result.messages); // Array of warnings/info (e.g., "Skipped unsupported drawing")

// Complex conversion with nested configuration
const htmlResult = await OfficeConverter.convert('data.xlsx', 'html', {
    parseConfig: { 
        ignoreNotes: true 
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
    onWarning: (msg) => console.warn("Conversion Warning:", msg)
});
```

## Native RAG Chunking
`officeParser` provides native support for document chunking, specifically designed for Retrieval-Augmented Generation (RAG) workflows. It offers three distinct strategies to split your documents while maintaining context and metadata.

### 1. Fixed-Size Strategy (Recursive)
Splits text into chunks based on character count with a specified overlap. It uses smart boundary detection to avoid cutting in the middle of sentences or paragraphs.

### 2. Document Structure Strategy
Splits the document at natural structural boundaries like pages (PDF/Word), slides (PPTX), or high-level headings. This preserves the logical flow of the document.

### 3. Semantic Strategy
Uses cosine similarity between sentence embeddings to identify coherent topic boundaries. This ensures that each chunk contains semantically related content (requires an embedding function).

### The `OfficeChunk` Interface
Every chunk produced contains not just text, but rich metadata to help your RAG pipeline:
```typescript
{
  text: string;           // The chunk content
  metadata: {
    sourceType: string;   // e.g., "docx", "pdf"
    pageNumber?: number;  // Current page
    slideNumber?: number; // Current slide
    closestHeading?: string; // The heading this chunk belongs to
    chunkIndex: number;   // Sequential index
  }
}
```

#### Example: Generating Chunks
```typescript
const chunks = await OfficeGenerator.generate(ast, 'chunks', {
    strategy: 'fixed-size',
    maxChunkSize: 1000,
    chunkOverlap: 200
});
console.log(`Generated ${chunks.value.length} chunks`);
```

### Using Callbacks (Backward Compatibility Support)
Callbacks are still supported for those preferred, but the data returned is now the AST object.
```js
const officeParser = require('officeparser');

officeParser.parseOffice("/path/to/officeFile.docx", function(ast, err) {
    if (err) {
        console.error(err);
        return;
    }
    // Get text from AST
    console.log(ast.toText());
});
```

### Using File Buffers or ArrayBuffers
You can pass a file path string, a Node.js `Buffer`, or an `ArrayBuffer`.
```js
const fs = require('fs');
const officeParser = require('officeparser');
const buffer = fs.readFileSync("/path/to/officeFile.pdf");

officeParser.parseOffice(buffer)
    .then(ast => console.log(ast.toText()))
    .catch(console.error);
```

## The AST Structure
The `OfficeParserAST` provides a format-agnostic representation of your document, allowing you to traverse and manipulate content as a tree.

### Visualizing the AST
The `OfficeParserAST` provides a format-agnostic representation of your document. Below is a simplified visualization of how the tree is structured:

```text
OfficeParserAST
├── type: "docx" | "pdf" | "xlsx" | "csv" | "md" | ... (11 formats supported)
├── metadata: { author, title, created, modified, ..., customProperties }
├── content: [ OfficeContentNode ]
│   ├── type: "paragraph" | "heading" | "table" | "list" | ...
│   ├── text: "Concatenated text of this node and all children"
│   ├── children: [ OfficeContentNode ] (recursive)
│   ├── formatting: { bold, italic, color, size, font, ... }
│   ├── metadata: { level, listId, paragraphIndentation, row, col, ... }
│   └── rawContent: "<xml>...</xml>" (if enabled)
├── attachments: [ OfficeAttachment ]
│   ├── type: "image" | "chart"
│   ├── name: "image1.png"
│   ├── data: "base64..."
│   ├── ocrText: "Text extracted via OCR"
│   └── chartData: { title, dataSets, labels, ... }
└── toText(): Function -> returns full plain text
```

#### Representative JSON Snippet
```json
{
  "type": "docx",
  "metadata": { "author": "John Doe", "title": "Annual Report", "customProperties": { "Department": "Finance" } },
  "content": [
    {
      "type": "heading",
      "text": "Introduction",
      "metadata": { "level": 1 },
      "children": [
        { "type": "text", "text": "Introduction", "formatting": { "bold": true } }
      ]
    },
    {
      "type": "paragraph",
      "text": "This is a report with an image.",
      "children": [
        { "type": "text", "text": "This is a report with an " },
        { "type": "image", "metadata": { "attachmentName": "img1.png" } }
      ]
    }
  ],
  "attachments": [
    { "name": "img1.png", "type": "image", "data": "iVBOR...", "ocrText": "Extracted Text" }
  ]
}
```

## Deep Dive: Document Components

### 1. Working with Lists
Lists are represented as sequential `list` nodes. To reconstruct or track a list, use the `metadata` fields:

```text
List Node
├── type: "list"
├── metadata: { 
    listId: "1", 
    listType: "ordered", 
    indentation: 0, 
    paragraphIndentation: { left: 720, hanging: 360 },
    itemIndex: 0 
}
└── children: [ Text Content... ]
```

- **`listId`**: A unique identifier for the list definition. Multiple items with the same `listId` belong to the same logical list.
- **`indentation`**: The structural nesting level (0-based).
- **`paragraphIndentation`**: The physical indentation formatting in twentieths of a point (twips) (e.g., `left`, `right`, `firstLine`, `hanging`).
- **`itemIndex`**: The sequential position within that list level.
- **`listType`**: Either `ordered` (numbered) or `unordered` (bulleted).

> [!TIP]
> Even if a list is interrupted by a regular paragraph, the `itemIndex` will continue to increment for the same `listId`, allowing you to maintain correct numbering.

### 2. Navigating Tables
Tables follow a strict hierarchy: `table` -> `row` -> `cell`.

```text
Table Node
├── type: "table"
└── children: [ Row Node ]
    ├── type: "row"
    └── children: [ Cell Node ]
        ├── type: "cell"
        ├── metadata: { row, col, rowSpan, colSpan }
        └── children: [ Paragraph/List/etc. ]
```

- **`row` / `col`**: Zero-based indices for grid positioning.
- **`rowSpan` / `colSpan`** (Optional): Integer values indicating merged cells (primarily in ODF formats). If absent, the cell is not merged.
- **Recursive Content**: Cells contain their own `children` array, which can include paragraphs, lists, or even other nested tables.

### 3. Charts & Data
When a chart is discovered, it's added as a `chart` node in the content and a corresponding `OfficeAttachment`.

```text
Chart Node
├── type: "chart"
├── metadata: { attachmentName: "chart1.xml" }
└── Attachment (Linked)
    └── chartData: { title, dataSets: [...], labels: [...] }
```

- **`attachmentName`**: Links the content node to the `attachments` array.
- **`chartData`**: A structured object containing titles, axis labels, and category/series data.

### 4. Images, OCR & Alt Text
Images are linked via `attachmentName` and can contain valuable metadata:

```text
Image Node
├── type: "image"
├── metadata: { attachmentName: "img1.png", altText: "..." }
└── Attachment (Linked)
    ├── data: "base64..."
    └── ocrText: "Extracted via OCR"
```

- **OCR Text**: If `ocr: true` is set in config, `ocrText` will contain the text found within the image.
- **Alt Text**: Extracted from the document's internal image descriptions.
- **Formatting**: `OfficeContentNode` images may also have parent alignment metadata.

### 5. Text Formatting
Each `OfficeContentNode` can have a `formatting` object that defines how the text should be styled.

```text
Text Node
└── formatting: {
    bold: boolean,
    italic: boolean,
    underline: boolean,
    strikethrough: boolean,
    color: "#hex",
    backgroundColor: "#hex",
    size: "12pt",
    font: "Arial",
    subscript: boolean,
    superscript: boolean,
    alignment: "left" | "center" | "right" | "justify"
}
```

Formatting can be found at two levels:
1.  **Node Level**: Applied directly to a text run or paragraph.
2.  **Document Level**: Found in `ast.metadata.formatting` (defaults) or `ast.metadata.styleMap` (named styles).

### 6. Breaks
Breaks are currently only supported when parsing DOCX-documents. Breaks are added as a node of type `break` and carry metadata of the type `BreakMetadata`. When `includeRawContent` is enabled, they also include the `rawContent` string from the original XML.

```text
Break Node
├── type: "break"
└── metadata: {
        breakType: "textWrapping" | "page" | "column" | "lastRenderedPage" | "carriageReturn",
        clear?: "all" | "left" | "none" | "right"
    }
```

- `breakType`: Type of break. `textWrapping` (default) is a standard line break, `page` is a page break, `column` is a break to the next column, `lastRenderedPage` is a soft break inserted by Word, and `carriageReturn` is an explicit carriage return (`w:cr`).
- `clear`: Relevant for `textWrapping`. Indicates if text should wrap around floating objects.

> [!NOTE]
> Even though break nodes don't have a `text` property, the `ast.toText()` method will automatically convert them to newlines (`\n`) or the configured delimiter in the final string output.

### 7. Advanced Metadata
The `ast.metadata` object provides document-wide context:
- **`styleMap`**: A dictionary of style names to their `TextFormatting` definitions found in the document.
- **`formatting`**: Document-wide default settings (e.g., default font or font size).
- **`customProperties`**: A dictionary of user-defined metadata embedded in the document (OOXML `custom.xml`, ODF `meta:user-defined`, or PDF Info dictionary).

### 8. Custom Properties
You can access custom user-defined metadata that might be embedded in the document:

```javascript
const ast = await officeParser.parseOffice("contract.docx");
console.log("Custom Metadata:", ast.metadata.customProperties);
// Output: { "ProjectID": "ABC-123", "InternalReview": true }
```

## Performance & Fidelity Highlights (v7.0.0)
The v7.0.0 release brings significant internal optimizations and fidelity improvements:
- **OpenOffice Speedups**: Up to **23x faster** parsing for ODP presentations thanks to optimized XML caching.
- **Excel Memory Efficiency**: Resolved $O(n)$ memory overhead issues for large spreadsheets (#91) by switching to iterative stream-based parsing.
- **RTF Performance**: Rewritten core loop to resolve $O(n^2)$ bottlenecks during string accumulation.
- **Advanced Table Fidelity**: Native support for **vertical cell merging** (`vMerge`) and **horizontal spanning** (`gridSpan`) in DOCX, ensuring complex tables look exactly as they do in Word.
- **Parser Extensions**: You can now parse `CSV`, `Markdown`, and `HTML` files *into* the unified Office AST, allowing you to use the `OfficeGenerator` on them just like any other format.

### Advanced AST Usage
Beyond using `ast.toText()`, you can interact with the structural data directly:

#### 1. Extract all images and their OCR text
```javascript
const ast = await officeParser.parseOffice("report.docx", { ocr: true });
const images = ast.attachments.filter(a => a.mimeType.startsWith('image/'));
images.forEach(img => {
    console.log(`Image: ${img.name} (OCR: ${img.ocrText || 'N/A'})`);
});
```

#### 2. Find specific headings
```javascript
const headings = ast.content.filter(node => node.type === 'heading' && node.metadata?.level === 1);
console.log("Main Chapters:", headings.map(h => h.text));
```

#### 3. Custom output (e.g., Simple Markdown conversion)
```javascript
const toMarkdown = (nodes) => {
    return nodes.map(node => {
        if (node.type === 'heading') return `${'#'.repeat(node.metadata?.level || 1)} ${node.text}`;
        if (node.type === 'list') return `- ${node.text}`;
        if (node.type === 'table') return "[Table Data]"; // expand children for actual table
        return node.text;
    }).join('\n\n');
};
console.log(toMarkdown(ast.content));
```

#### 4. Extracting Tables to CSV
Iterate through table nodes and their children (rows -> cells) to build a CSV string.
```javascript
const tables = ast.content.filter(node => node.type === 'table');
tables.forEach((table, index) => {
    const csv = table.children
        .filter(row => row.type === 'row')
        .map(row => 
            row.children
                .filter(cell => cell.type === 'cell')
                .map(cell => `"${cell.text.replace(/"/g, '""')}"`) // Escape quotes
                .join(',')
        )
        .join('\n');
    console.log(`Table ${index + 1} CSV:\n${csv}`);
});
```

#### 5. Filtering by Formatting (e.g., Bold Text)
Find all text nodes that have specific formatting applied.
```javascript
function findBoldText(nodes) {
    let results = [];
    nodes.forEach(node => {
        if (node.type === 'text' && node.formatting?.bold) {
            results.push(node.text);
        }
        if (node.children) {
            results = results.concat(findBoldText(node.children));
        }
    });
    return results;
}

const boldStrings = findBoldText(ast.content);
console.log("Bold Text Found:", boldStrings);
```

#### 6. Processing Footnotes/Endnotes
If you kept notes inline (default behavior), you can extract them into a separate list for processing.
```javascript
function extractNotes(nodes) {
    let notes = [];
    nodes.forEach(node => {
        if (node.type === 'note') {
            notes.push({ id: node.metadata.noteId, text: node.text, type: node.metadata.noteType });
        }
        if (node.children) {
            notes = notes.concat(extractNotes(node.children));
        }
    });
    return notes;
}

const allNotes = extractNotes(ast.content);
console.log("Document Notes:", allNotes);
```

## Configuration Object: OfficeParserConfig
Pass an optional config object as the second argument to `parseOffice`.

| Flag | DataType | Default | Explanation |
|------|----------|---------|-------------|
| `outputErrorToConsole` | boolean | `false` | **Deprecated**: Use `onWarning` instead. Show logs to console in case of an error. |
| `newlineDelimiter` | string | `\n` | Delimiter for new lines in text output. |
| `ignoreNotes` | boolean | `false` | Ignore notes in files like PowerPoint/ODP. |
| `putNotesAtLast` | boolean | `false` | Put notes text at the end of the document. |
| `extractAttachments` | boolean | `false` | Extract images and charts as Base64. |
| `includeRawContent` | boolean | `false` | Include raw XML/RTF markup in the nodes. |
| `serializeRawContent` | boolean | `true` | Re-serializes raw XML to clean strings. |
| `preserveXmlWhitespace` | boolean | `false` | Preserves original XML whitespace. |
| `ocr` | boolean | `false` | Enable OCR for images (requires `extractAttachments: true`). |
| `pdfWorkerSrc` | string | `(see below)` | Path to PDF.js worker. |
| `ocrConfig` | object | `{}` | OCR Scheduler configuration. |
| `includeBreakNodes` | boolean | `false` | Include `w:br`, `w:cr` nodes (DOCX only).|
| `ignoreInternalLinks` | boolean | `false` | Remove all bookmarks and internal jumps. |
| `csvDelimiter` | string | `,` | Custom delimiter for parsing CSV files. |
| `fileType` | string | `null` | Manual format override (authoritative). |

## Generator Configuration: GeneratorConfig
Configuration options for `OfficeGenerator.generate`.

| Flag | DataType | Default | Explanation |
|------|----------|---------|-------------|
| `includeFormatting` | boolean | `false` | Whether to include semantic styles (bold, italic) in output. |
| `styleMap` | string[] \| array | `[]` | Array of style mappings (DSL strings or structured objects). |
| `ignoreDefaultStyleMap`| boolean | `false` | Ignore the library's default style mappings. |
| `includeMetadata` | boolean | `false` | Include document metadata in the output (e.g., as frontmatter). |
| `onNode` | function | `undefined` | Callback to intercept/modify any node during generation. |

### 🛠️ Advanced Node Manipulation (Pro Users)
The `onNode` callback is a powerful tool that gives you complete control over the generation process. It is called for **every single node** in the AST before it is rendered.

#### Callback Capabilities:
1.  **Filter/Remove Nodes**: Return `false` to skip a node and all its children.
2.  **Override Rendering**: Return a `string` to use that exact text as the output, bypassing default logic and recursion.
3.  **Mutate Nodes**: Modify the `node` object directly (e.g., changing `node.text`) and return `void` to let the generator proceed with your changes.
4.  **Async Support**: The callback can be `async`, allowing you to fetch external data or perform complex logic during generation.

#### Pro Example:
```typescript
const result = await ast.to('md', {
    onNode: async (node) => {
        // 1. Skip all images
        if (node.type === 'image') return false;

        // 2. Redact sensitive info by mutating the node
        if (node.text?.includes('SECRET_KEY')) {
            node.text = node.text.replace(/SECRET_KEY: \w+/, 'SECRET_KEY: [REDACTED]');
        }

        // 3. Custom rendering for specific styles
        if (node.metadata?.style === 'Callout') {
            return `> [!INFO]\n> ${node.text}`;
        }

        // 4. Proceed with default rendering (implicitly returns void)
    }
});
```

### Advanced Style Mapping (Semantic Translation)
The `styleMap` configuration is the primary way to define the "semantic meaning" of document styles. We recommend using **Structured Style Mappings** for full type safety and power.

#### 1. Structured Style Mappings (Recommended)
Use structured objects to match nodes based on type and attributes, and specify detailed output properties like classes and custom attributes.

```typescript
styleMap: [
    { 
        selector: { 
            nodeType: 'paragraph', 
            attributes: { style: 'Heading 1' } 
        }, 
        output: { 
            tag: 'h1', 
            classes: ['main-title'], 
            attributes: { id: 'top' } 
        } 
    },
    {
        // Use operators like '~=' for partial matches
        selector: { attributes: { style: { value: 'Quote', operator: '~=' } } },
        output: { tag: 'blockquote' }
    }
]
```

#### 2. Legacy String DSL
The library also maintains support for a simple string-based DSL, highly compatible with `mammoth.js`.

- **Literal Matching**: `"p[style-name='Heading 1'] => h1"`
- **Regex-like Matching**: `"p[style~='Title'] => h2"`
- **Attribute Filters**: `"p[style-name='Quote'][lang='en'] => blockquote"`

## Chunking Configuration: ChunkingConfig
Specific options when using `format: 'chunks'`.

| Flag | DataType | Default | Explanation |
|------|----------|---------|-------------|
| `strategy` | string | `'fixed-size'`| The chunking strategy (`fixed-size`, `document-structure`, `semantic`). |
| `maxChunkSize` | number | `1000` | Maximum characters per chunk. |
| `chunkOverlap` | number | `200` | Overlap between consecutive chunks. |
| `similarityThreshold`| number | `0.5` | Threshold for semantic splitting (0.0 to 1.0). |
| `embedBatchSize` | number | `50` | Batch size for embedding requests. |

### OCR Scheduler & Resource Management
If your application uses OCR, `officeParser` utilizes an intelligent **Smart Worker Pool** to maintain a background worker pool and optimize repeated parse requests.

- **Dynamic Affinity**: Workers in the pool persist with their last used language affinity. 
- **LRU Re-allocation**: If a new language is requested and the pool is full, the manager identifies the **Least Recently Used (LRU)** idle worker and re-initializes it for the new language. This avoids the overhead of destroying and recreating workers.
- **Auto-Termination**: Workers are automatically cleaned up after 10 seconds of inactivity (configurable via `ocrConfig.autoTerminateTimeout`).

#### `OfficeParser.terminateOcr()`
If you have used OCR (`{ ocr: true }`) in a short-lived script (like CLI tools or one-off automation), we recommend explicitly calling `terminateOcr()` after your processing is finished. This bypasses the 10-second idle timer and allows the process to return to the terminal prompt immediately.

> [!NOTE]
> If OCR was not used, this function is a no-op and does not need to be called.

```js
const officeParser = require('officeparser');

async function runCleaner() {
    await officeParser.parseOffice("file.pdf", { ocr: true });
    // ... process results ...

    // Manually kill OCR workers for an immediate exit
    await officeParser.terminateOcr();
}
```

> [!TIP]
> This is handled automatically in the built-in CLI (`npx officeparser ...`). You only need to call this manually if you are using the library in your own custom script and want a snappy exit.

```js
const config = {
    newlineDelimiter: "\n\n",
    extractAttachments: true,
    ocr: true,
    ocrLanguage: 'eng+fra+esp' // Supports English, French, and Spanish simultaneously
};

const ast = await officeParser.parseOffice("report.docx", config);
console.log(`Extracted ${ast.attachments.length} images`);
```

## Examples

**Search for a term in a document (TypeScript)**
```ts
import { OfficeParser } from 'officeparser';

async function hasSearchTerm(filePath: string, term: string): Promise<boolean> {
    const ast = await OfficeParser.parseOffice(filePath);
    return ast.toText().includes(term);
}
```

**Extracting Images and their OCR text**
```js
const officeParser = require('officeparser');

const config = { extractAttachments: true, ocr: true };
officeParser.parseOffice("presentation.pptx", config).then(ast => {
    ast.attachments.forEach(attachment => {
        if (attachment.type === 'image') {
            console.log(`Image: ${attachment.name}`);
            console.log(`OCR Text: ${attachment.ocrText}`);
            fs.writeFileSync(attachment.name, Buffer.from(attachment.data, 'base64'));
        }
    });
});
```

## Browser Usage
The library provides two types of browser bundles in the `dist/` directory:
1. **`officeparser.browser.iife.js`**: Standard IIFE bundle for direct `<script>` tag usage. Exposes the global `officeParser` namespace.
2. **`officeparser.browser.mjs`**: Modern ESM bundle for use with `import` statements or modern bundlers.

### Usage (ESM)
If you are using a modern bundler like **Vite**, **Webpack**, or **Next.js**:

```javascript
import { OfficeParser } from 'officeparser';

const handleFile = async (event) => {
    const file = event.target.files[0];
    const buffer = await file.arrayBuffer();
    
    try {
        // Pass the Buffer or Uint8Array directly
        const ast = await OfficeParser.parseOffice(new Uint8Array(buffer));
        console.log(ast.toText());
    } catch (err) {
        console.error(err);
    }
};
```

> [!NOTE] 
> **Why `fs` fails in the browser**: Browsers do not have a built-in file system. If you try to pass a file path string in the browser, `officeParser` will throw a descriptive "Fail-Fast" error instead of crashing mysteriously: 
> `[officeparser] Node.js 'fs' module is not available in the browser. Please pass a Buffer or Uint8Array instead.`

### Usage (Script Tag)
Include the IIFE bundle available in the release assets or your `dist/` folder. This exposes the global `officeParser` object.

```html
<script src="dist/officeparser.browser.iife.js"></script>
<script>
    async function handleFile(event) {
        const file = event.target.files[0];
        const buffer = await file.arrayBuffer();
        
        try {
            // Reconstruct as Uint8Array for the parser
            const ast = await officeParser.parseOffice(new Uint8Array(buffer));
            console.log(ast.toText());
        } catch (error) {
            console.error("Parsing failed:", error);
        }
    }
</script>
```

### PDF Worker Configuration in Browser
When using `officeparser` in a browser environment to parse PDF files, you may provide the `pdfWorkerSrc` configuration option. If not provided, it defaults to a CDN link for `pdfjs-dist@5.6.205`.

```javascript
const file = ...; // File object or ArrayBuffer

// It will use the default CDN worker if pdfWorkerSrc is omitted
const ast = await officeParser.parseOffice(file);

// Or override it with your own path or a different version:
const ast2 = await officeParser.parseOffice(file, {
    pdfWorkerSrc: "https://cdn.jsdelivr.net/npm/pdfjs-dist@5.6.205/build/pdf.worker.min.mjs"
});
```

> **Note:** The version of `pdfjs-dist` in the worker source should match the version used by `officeparser` (currently `5.6.205`).

## Troubleshooting & Common Issues

- **Node.js process stays alive after finishing**: If using OCR, the worker pool stays warm for 10s by default. Use `await terminateOcr()` at the end of your script for a snappy exit.
- **"Worker not found" in Browser**: Ensure `pdfWorkerSrc` is correctly pointed to the `pdf.worker.min.mjs` file matching version `5.6.205`.
- **OCR accuracy is low**: Verify your `ocrConfig.language` matches the document content. Note that OCR quality depends on image resolution.
- **Out of memory on large files**: For massive spreadsheets, consider using `ast.toText()` early and allowing the full AST object to be garbage-collected.

For a comprehensive guide, visit our [Debugging & Troubleshooting Documentation](https://harshankur.github.io/officeParser/#spec/debugging).


## Known Limitations
1. **ODT/ODS Charts**: Extraction may occasionally show inaccurate data when referencing external cell ranges or complex layout-based data.
2. **PDF Images**: PDF images are extracted as BMP files in the browser for compatibility. This conversion happens automatically.
3. **RTF Footnotes**: The `putNotesAtLast` configuration is currently not supported for RTF files; footnotes and endnotes are always collected and appended to the end of the content.

----------

**npm**: [https://npmjs.com/package/officeparser](https://npmjs.com/package/officeparser)

**github**: [https://github.com/harshankur/officeParser](https://github.com/harshankur/officeParser)

## Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for details on how to get started.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

# Docstream

A universal Node.js & Browser library to parse any office document — legacy or modern — into structured text, AST or Markdown. Supports doc, xls, ppt, docx, xlsx, pptx, odt, ods, odp, pdf, rtf and more.

[![Node.js Version](https://img.shields.io/badge/node-%3E%3D18-brightgreen)](https://nodejs.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## Supported Formats

| Format | Extension | Type |
|--------|-----------|------|
| Word (OOXML) | `.docx` | Modern |
| Excel (OOXML) | `.xlsx` | Modern |
| PowerPoint (OOXML) | `.pptx` | Modern |
| OpenDocument Text | `.odt` | Modern |
| OpenDocument Spreadsheet | `.ods` | Modern |
| OpenDocument Presentation | `.odp` | Modern |
| PDF | `.pdf` | Modern |
| Rich Text Format | `.rtf` | Legacy |
| Word 97-2003 | `.doc` | Legacy (planned) |
| Excel 97-2003 | `.xls` | Legacy (planned) |
| PowerPoint 97-2003 | `.ppt` | Legacy (planned) |

## Install

```bash
npm i docstream
```

## Command Line Usage

Parse any office file directly from the terminal. Returns the full AST as JSON by default, or plain text with `--toText`.

```bash
# Get full AST as JSON (default)
npx docstream /path/to/officeFile.docx

# Get plain text only
npx docstream /path/to/officeFile.docx --toText=true

# Use configuration options
npx docstream /path/to/officeFile.docx --ignoreNotes=true --newlineDelimiter=" "
```

### CLI Config Options

| Option | Description |
|--------|-------------|
| `--toText=[true\|false]` | Output plain text instead of JSON AST |
| `--ignoreNotes=[true\|false]` | Ignore notes (e.g. PowerPoint speaker notes). Default: `false` |
| `--newlineDelimiter=[delimiter]` | Delimiter for new lines. Default: `\n` |
| `--putNotesAtLast=[true\|false]` | Collect notes at end of document. Default: `false` |
| `--outputErrorToConsole=[true\|false]` | Log errors to console. Default: `false` |
| `--extractAttachments=[true\|false]` | Extract images/charts as Base64. Default: `false` |
| `--ocr=[true\|false]` | Enable OCR for extracted images. Default: `false` |
| `--includeRawContent=[true\|false]` | Include raw XML/RTF content in nodes. Default: `false` |

## Library Usage

### Getting Started (Async/Await)

```js
const docstream = require('docstream');

async function parseMyFile() {
    try {
        // parseOffice returns an OfficeParserAST object
        const ast = await docstream.parseOffice("/path/to/officeFile.docx");

        // Get plain text
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

### Quick Text Extraction

```js
const getText = async (file, config) => (await docstream.parseOffice(file, config)).toText();

const text = await getText("/path/to/officeFile.docx");
console.log(text);
```

### Using Callbacks

Callbacks are supported for backward compatibility. The data returned is the AST object.

```js
const docstream = require('docstream');

docstream.parseOffice("/path/to/officeFile.docx", function(ast, err) {
    if (err) {
        console.error(err);
        return;
    }
    console.log(ast.toText());
});
```

### Using File Buffers or ArrayBuffers

You can pass a file path string, a Node.js `Buffer`, or an `ArrayBuffer`.

```js
const fs = require('fs');
const docstream = require('docstream');
const buffer = fs.readFileSync("/path/to/officeFile.pdf");

docstream.parseOffice(buffer)
    .then(ast => console.log(ast.toText()))
    .catch(console.error);
```

## The AST Structure

`OfficeParserAST` provides a format-agnostic representation of any document, allowing you to traverse and manipulate content as a tree.

```text
OfficeParserAST
├── type: "docx" | "pptx" | "xlsx" | ...
├── metadata: { author, title, created, modified, ... }
├── content: [ OfficeContentNode ]
│   ├── type: "paragraph" | "heading" | "table" | "list" | ...
│   ├── text: "Concatenated text of this node and all children"
│   ├── children: [ OfficeContentNode ] (recursive)
│   ├── formatting: { bold, italic, color, size, font, ... }
│   ├── metadata: { level, listId, row, col, ... }
│   └── rawContent: "<xml>...</xml>" (if enabled)
├── attachments: [ OfficeAttachment ]
│   ├── type: "image" | "chart"
│   ├── name: "image1.png"
│   ├── data: "base64..."
│   ├── ocrText: "Text extracted via OCR"
│   └── chartData: { title, dataSets, labels, ... }
├── toText(): returns full plain text
└── toMarkdown(): returns Markdown representation
```

### Representative JSON

```json
{
  "type": "docx",
  "metadata": { "author": "John Doe", "title": "Annual Report" },
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

### Lists

Lists are represented as sequential `list` nodes. To reconstruct or track a list, use the `metadata` fields:

```text
List Node
├── type: "list"
├── metadata: {
│   listId: "1",
│   listType: "ordered",
│   indentation: 0,
│   itemIndex: 0
│ }
└── children: [ Text Content... ]
```

- **`listId`**: Unique identifier for the list definition. Items with the same `listId` belong to the same logical list.
- **`indentation`**: Nesting level (0-based).
- **`itemIndex`**: Sequential position within that list level.
- **`listType`**: Either `ordered` (numbered) or `unordered` (bulleted).

> [!TIP]
> Even if a list is interrupted by a regular paragraph, the `itemIndex` will continue to increment for the same `listId`, allowing you to maintain correct numbering.

### Tables

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
- **`rowSpan` / `colSpan`** (Optional): Integer values indicating merged cells. If absent, the cell is not merged.
- Cells contain their own `children` array, which can include paragraphs, lists, or nested tables.

### Charts & Data

When a chart is discovered, it's added as a `chart` node in the content and a corresponding `OfficeAttachment`.

```text
Chart Node
├── type: "chart"
├── metadata: { attachmentName: "chart1.xml" }
└── Attachment (Linked)
    └── chartData: { title, dataSets: [...], labels: [...] }
```

### Images, OCR & Alt Text

```text
Image Node
├── type: "image"
├── metadata: { attachmentName: "img1.png", altText: "..." }
└── Attachment (Linked)
    ├── data: "base64..."
    └── ocrText: "Extracted via OCR"
```

- **OCR Text**: Set `ocr: true` in config to extract text from images via Tesseract.js.
- **Alt Text**: Extracted from the document's internal image descriptions.

### Text Formatting

Each `OfficeContentNode` can have a `formatting` object:

```text
Text Node
└── formatting: {
    bold, italic, underline, strikethrough,
    color: "#hex", backgroundColor: "#hex",
    size: "12pt", font: "Arial",
    subscript, superscript,
    alignment: "left" | "center" | "right" | "justify"
}
```

Formatting appears at two levels:
1. **Node Level**: Applied directly to a text run or paragraph.
2. **Document Level**: Found in `ast.metadata.formatting` (defaults) or `ast.metadata.styleMap` (named styles).

## Configuration: `OfficeParserConfig`

Pass an optional config object as the second argument to `parseOffice`.

| Flag | Type | Default | Description |
|------|------|---------|-------------|
| `outputErrorToConsole` | boolean | `false` | Log errors to console |
| `newlineDelimiter` | string | `\n` | Delimiter for new lines in text output |
| `ignoreNotes` | boolean | `false` | Ignore notes in PowerPoint/ODP files |
| `putNotesAtLast` | boolean | `false` | Append notes at end of document (not supported for RTF) |
| `extractAttachments` | boolean | `false` | Extract images and charts as Base64 |
| `ocr` | boolean | `false` | Enable OCR for images (requires `extractAttachments: true`) |
| `ocrLanguage` | string | `eng` | OCR language(s), e.g. `'eng+fra+esp'`. See [language codes](https://tesseract-ocr.github.io/tessdoc/Data-Files#data-files-for-version-400-november-29-2016) |
| `includeRawContent` | boolean | `false` | Include raw XML/RTF markup in nodes |
| `pdfWorkerSrc` | string | CDN | Path to PDF.js worker. Defaults to CDN for `pdfjs-dist@5.4.530` |

```js
const config = {
    newlineDelimiter: "\n\n",
    extractAttachments: true,
    ocr: true,
    ocrLanguage: 'eng+fra+esp'
};

const ast = await docstream.parseOffice("report.docx", config);
console.log(`Extracted ${ast.attachments.length} images`);
```

## Markdown Output

The `toMarkdown()` method on the AST result converts the parsed document into clean Markdown, preserving headings, lists, tables, bold/italic formatting, images (as references), and footnotes.

```js
const docstream = require('docstream');

const ast = await docstream.parseOffice("report.docx");
const markdown = ast.toMarkdown();
console.log(markdown);
```

Output:

```markdown
# Introduction

This is the **first paragraph** with *italic text* and a [link](https://example.com).

## Section 1

- Item one
- Item two
- Item three

| Name  | Score |
|-------|-------|
| Alice | 95    |
| Bob   | 87    |
```

`toMarkdown()` works with all supported formats. Combined with `extractAttachments: true`, image nodes are rendered as `![alt](attachment-name)` references.

## Legacy Format Support

docstream adds native support for `.doc`, `.xls`, and `.ppt` (Office 97-2003 binary formats) without requiring LibreOffice or any external dependency. These parsers read the OLE2 Compound Binary File (CFB) container directly and extract content from the underlying binary streams (Word Binary, BIFF8, and PowerPoint Binary respectively).

This means you can parse legacy Office files in the same way as modern formats — no system-level dependencies, no subprocess spawning, and full cross-platform compatibility including the browser.

```js
// Works exactly the same as modern formats
const ast = await docstream.parseOffice("legacy-report.doc");
console.log(ast.toText());

const ast2 = await docstream.parseOffice("budget.xls");
console.log(ast2.content); // Tables with rows and cells
```

## Examples

**Search for a term (TypeScript)**

```ts
import { OfficeParser } from 'docstream';

async function hasSearchTerm(filePath: string, term: string): Promise<boolean> {
    const ast = await OfficeParser.parseOffice(filePath);
    return ast.toText().includes(term);
}
```

**Extract images with OCR**

```js
const docstream = require('docstream');

const config = { extractAttachments: true, ocr: true };
docstream.parseOffice("presentation.pptx", config).then(ast => {
    ast.attachments.forEach(attachment => {
        if (attachment.type === 'image') {
            console.log(`Image: ${attachment.name}`);
            console.log(`OCR Text: ${attachment.ocrText}`);
            fs.writeFileSync(attachment.name, Buffer.from(attachment.data, 'base64'));
        }
    });
});
```

**Find specific headings**

```js
const ast = await docstream.parseOffice("document.docx");
const headings = ast.content.filter(node => node.type === 'heading' && node.metadata?.level === 1);
console.log("Main Chapters:", headings.map(h => h.text));
```

**Extract tables to CSV**

```js
const tables = ast.content.filter(node => node.type === 'table');
tables.forEach((table, index) => {
    const csv = table.children
        .filter(row => row.type === 'row')
        .map(row =>
            row.children
                .filter(cell => cell.type === 'cell')
                .map(cell => `"${cell.text.replace(/"/g, '""')}"`)
                .join(',')
        )
        .join('\n');
    console.log(`Table ${index + 1} CSV:\n${csv}`);
});
```

**Find bold text**

```js
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

**Extract footnotes/endnotes**

```js
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

## Browser Usage

The browser bundle exposes the `officeParser` namespace. Include the bundle file from the release assets.

```html
<script src="dist/officeparser.browser.js"></script>
<script>
    async function handleFile(file) {
        // file: a File object from <input> or an ArrayBuffer
        try {
            const ast = await officeParser.parseOffice(file, { ocr: true });
            console.log(ast.toText());
            console.log("Metadata:", ast.metadata);
        } catch (error) {
            console.error(error);
        }
    }
</script>
```

### PDF Worker Configuration in Browser

When parsing PDFs in the browser, you can provide `pdfWorkerSrc` in the config. If omitted, it defaults to a CDN link for `pdfjs-dist@5.4.530`.

```js
// Uses default CDN worker
const ast = await officeParser.parseOffice(file);

// Override with your own path
const ast2 = await officeParser.parseOffice(file, {
    pdfWorkerSrc: "https://unpkg.com/pdfjs-dist@5.4.530/build/pdf.worker.min.mjs"
});
```

> **Note:** The `pdfjs-dist` version in the worker source should match the version used by docstream (currently `5.4.530`).

## Known Limitations

1. **ODT/ODS Charts**: Extraction may show inaccurate data when referencing external cell ranges or complex layouts.
2. **PDF Images**: PDF images are extracted as BMP in the browser for compatibility.
3. **RTF Footnotes**: `putNotesAtLast` is not supported for RTF; notes are always appended at the end.

## Roadmap

- [x] DOCX, XLSX, PPTX, ODF, PDF, RTF parsing (from officeParser)
- [x] AST output with metadata, formatting and attachments
- [ ] Markdown output (`toMarkdown()`)
- [ ] Legacy `.doc` support (Word 97-2003 Binary)
- [ ] Legacy `.xls` support (Excel BIFF8)
- [ ] Legacy `.ppt` support (PowerPoint 97-2003 Binary)
- [ ] Fix: `process is not defined` in browser environments ([issue #67](https://github.com/harshankur/officeParser/issues/67))
- [ ] Fix: background process leak on top-level require ([issue #59](https://github.com/harshankur/officeParser/issues/59))
- [ ] Page numbers in Word documents ([issue #71](https://github.com/harshankur/officeParser/issues/71))

## Credits & References

This project builds on the work of several open-source projects and specifications:

- **[officeParser](https://github.com/harshankur/officeParser)** by harshankur — The original parser this project is forked from. Provides DOCX, XLSX, PPTX, ODF, PDF, and RTF parsing with AST output. (MIT)
- **[mammoth.js](https://github.com/mwilliamson/mammoth.js)** by mwilliamson — DOCX to semantic HTML conversion, referenced for the Markdown output pipeline. (MIT)
- **[turndown](https://github.com/mixmark-io/turndown)** by mixmark-io — HTML to Markdown conversion engine. (MIT)
- **[markitdown](https://github.com/microsoft/markitdown)** by Microsoft — Modular document converter architecture, used as design inspiration. (MIT)
- **[olefile](https://github.com/decalage2/olefile)** by decalage2 — OLE2/Compound Binary File format parsing algorithms, ported to TypeScript for legacy format support. (BSD-2-Clause)
- **[Apache POI](https://github.com/apache/poi)** — Java reference implementation for BIFF8 (XLS), Word Binary (DOC), and PPT binary format structures. (Apache 2.0)
- **[pdf.js](https://github.com/mozilla/pdf.js)** by Mozilla — PDF text and image extraction engine. (Apache 2.0)
- **Microsoft Open Specifications** — Official binary format documentation:
  [MS-DOC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc),
  [MS-XLS](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls),
  [MS-PPT](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt)

## Contributing

Contributions are welcome. See [CONTRIBUTING.md](CONTRIBUTING.md) for details.

## License

MIT — see [LICENSE](LICENSE).

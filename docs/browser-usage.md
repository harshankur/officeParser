# Browser Usage Guide

`officeparser` now provides a robust ESM and IIFE browser bundle.

## Features
- **Universal Support**: Works in all modern browsers and environments (Webpack, Vite, Parcel).
- **Fail-Fast `fs` Stub**: Tries to use Node.js `fs` in the browser will now throw a descriptive error instead of failing silently or crashing during build.
- **Buffer Support**: Seamlessly handles `Buffer` (from `buffer` polyfill) or standard `Uint8Array`.

## Installation

```bash
npm install officeparser
```

## Usage (ESM)

If you are using a modern bundler like Vite, Webpack, or Next.js:

```javascript
import { OfficeParser } from 'officeparser';

const handleFile = async (event) => {
    const file = event.target.files[0];
    const buffer = await file.arrayBuffer();
    
    try {
        const ast = await OfficeParser.parseOffice(new Uint8Array(buffer), {
            includeRawContent: true
        });
        console.log("Parsed AST:", ast);
        console.log("Text:", ast.toText());
    } catch (err) {
        console.error("Parsing failed:", err);
    }
};
```

## Usage (Script Tag)

You can also use the IIFE bundle directly in a script tag:

```html
<script src="dist/officeparser.browser.iife.js"></script>
<script>
    const handleFile = async (file) => {
        const buffer = await file.arrayBuffer();
        const ast = await officeParser.parseOffice(new Uint8Array(buffer));
        alert(ast.toText());
    };
</script>
```

## Important: Why `fs` fails in the browser
Browsers do not have a built-in file system API like Node.js (`fs`). 
Legacy versions of `officeparser` might have tried to polyfill `fs` with an empty object, leading to confusing `undefined is not a function` errors.

Our new **Fail-Fast Stub** will instead throw a clear error:
`[officeparser] Node.js 'fs' module is not available in the browser. Please pass a Buffer or Uint8Array to parseOffice instead of a file path.`

This ensures you know exactly why the code isn't working as expected.

## XML Serialization Options

When using `includeRawContent: true`, you can now control how XML is returned:

```javascript
const ast = await OfficeParser.parseOffice(buffer, {
    includeRawContent: true,
    serializeRawContent: false, // Return original XML substring from source (default: true)
    preserveXmlWhitespace: false // Minify serialized XML (default: false)
});
```

- **serializeRawContent: true** (Default): Re-serializes the XML from the DOM. This ensures valid XML but may change whitespace or attribute order.
- **serializeRawContent: false**: Extracts the exact substring from the original file based on DOM locators. This is faster and preserves the byte-for-byte original content.

## OCR Scheduler & Resource Management

If you enable OCR in the browser (`{ ocr: true }`), `officeParser` will initialize a pool of Tesseract.js workers managed by an intelligent **Smart Worker Pool**. 

- **Efficient Switching**: These workers persist with their language affinity. If you switch between parsing of English and French documents, the pool will automatically re-allocate workers using an **LRU (Least Recently Used)** strategy—re-initializing the oldest idle worker to the new language only when necessary.
- **Resource Cleanup**: Workers are automatically cleaned up after an inactivity timeout of 10 seconds.

### Manual OCR Cleanup
If you have performed a parse with OCR enabled and want to free up memory and terminate workers immediately:

```javascript
await officeParser.parseOffice(file, { ocr: true });
// ... process results ...

// Kill all background workers immediately
await officeParser.terminateOcr();
```

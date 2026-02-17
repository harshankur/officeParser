/**
 * PDF Parser
 *
 * Extracts text, metadata, images, links, and attachments from PDF files using PDF.js (pdfjs-dist).
 *
 * **Features:**
 * - Text extraction with formatting (bold, italic, font, size)
 * - Comprehensive metadata extraction (title, author, subject, creator, producer, creation/modification dates)
 * - Hyperlink extraction from PDF annotations
 * - Heading detection via font size heuristics
 * - Image extraction as attachments with optional OCR (using Tesseract.js)
 * - Embedded file attachment extraction
 * - Layout preservation (respects order of text and images)
 *
 * **PDF Format Limitations (compared to DOCX/ODT):**
 *
 * PDF was designed as a "page description language" for visual fidelity, not semantic structure.
 * The following features **cannot be reliably extracted** from PDFs:
 *
 * - **Tables**: PDF has no table structure. Tables are just text positioned to look tabular.
 *   Extracting tables would require complex spatial analysis with many false positives.
 *   See: https://stackoverflow.com/questions/36978446/why-is-it-difficult-to-extract-data-from-pdfs
 *
 * - **Lists**: PDF has no list structure. Bullets/numbers are just text characters.
 *   No hierarchy or list type information is stored. Would require heuristic detection
 *   that would have many edge cases and errors.
 *
 * - **Styles**: PDF has no style definitions like "Heading1" or "Normal". Only visual
 *   properties (font, size) exist. We use font size heuristics to detect headings.
 *
 * - **Notes (Footnotes/Endnotes)**: PDF has no concept of footnotes/endnotes as structured
 *   elements. They're just smaller text at the bottom of pages.
 *
 * - **Text Color**: While PDF stores color, pdfjs-dist doesn't expose text color in the
 *   textContent API. Would require parsing the operator stream which is complex.
 *
 * - **Background Color**: Same limitation as text color.
 *
 * - **Underline/Strikethrough**: These are drawn as separate line elements in PDF,
 *   not properties of text. Association would require spatial analysis.
 *
 * **Parsing Approach:**
 * 1. Load PDF document using pdfjs-dist.
 * 2. Extract global metadata from document info dictionary.
 * 3. Extract embedded file attachments.
 * 4. Iterate through each page:
 *    a. Collect text items with position and formatting.
 *    b. Collect link annotations with associated text.
 *    c. Collect images from the operator list.
 *    d. Sort all items by vertical position (top-to-bottom reading order).
 *    e. Group text into paragraphs/headings based on line breaks and font sizes.
 *    f. Process images as attachments with optional OCR.
 * 5. Apply heading detection based on font size heuristics.
 *
 * @module PdfParser
 * @see https://mozilla.github.io/pdf.js/ PDF.js documentation
 * @see https://www.adobe.com/devnet/pdf/pdf_reference.html PDF Reference
 */
/// <reference types="node" />
import { OfficeParserAST, OfficeParserConfig } from '../types';
/**
 * Parses a PDF file and extracts content.
 *
 * @param buffer - The PDF file buffer
 * @param config - Parser configuration
 * @returns Promise resolving to the parsed AST
 */
export declare const parsePdf: (buffer: Buffer, config: OfficeParserConfig) => Promise<OfficeParserAST>;

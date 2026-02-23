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

import { ImageMetadata, OfficeAttachment, OfficeContentNode, OfficeMetadata, OfficeParserAST, OfficeParserConfig, TextFormatting, TextMetadata } from '../types';
import { getOfficeError, logWarning, OfficeErrorType } from '../utils/errorUtils';
import { createAttachment } from '../utils/imageUtils';
import { performOcr } from '../utils/ocrUtils';

/** Represents a text item with position and formatting */
interface TextPageItem {
    type: 'text';
    x: number;
    y: number;
    width: number;
    height: number;
    text: string;
    fontName?: string;
    formatting: TextFormatting;
}

/** Represents an image with position */
interface ImagePageItem {
    type: 'image';
    x: number;
    y: number;
    name: string;
    data: Uint8Array | Uint8ClampedArray;
    width: number;
    height: number;
    kind?: number;
}

/** Represents a link annotation */
interface LinkAnnotation {
    rect: number[];  // [x1, y1, x2, y2]
    url?: string;
    dest?: any;  // Internal destination
}

type PageItem = TextPageItem | ImagePageItem;

/**
 * Parses PDF creation/modification date strings.
 * PDF dates are in format: D:YYYYMMDDHHmmSSOHH'mm'
 * Where O is timezone offset direction (+/-), HH is hours, mm is minutes.
 * 
 * @param dateString - The PDF date string
 * @returns Parsed Date object or undefined if parsing fails
 */
function parsePdfDate(dateString: string | undefined): Date | undefined {
    if (!dateString) return undefined;

    try {
        // Remove "D:" prefix if present
        let str = dateString.startsWith('D:') ? dateString.slice(2) : dateString;

        // Extract components: YYYYMMDDHHmmSS
        const year = parseInt(str.slice(0, 4), 10);
        const month = parseInt(str.slice(4, 6), 10) - 1; // 0-indexed
        const day = parseInt(str.slice(6, 8), 10) || 1;
        const hour = parseInt(str.slice(8, 10), 10) || 0;
        const minute = parseInt(str.slice(10, 12), 10) || 0;
        const second = parseInt(str.slice(12, 14), 10) || 0;

        // Handle timezone if present
        const tzMatch = str.slice(14).match(/([+-Z])(\d{2})'?(\d{2})?'?/);
        if (tzMatch) {
            const tzSign = tzMatch[1] === '-' ? -1 : 1;
            const tzHours = parseInt(tzMatch[2], 10) || 0;
            const tzMinutes = parseInt(tzMatch[3], 10) || 0;
            const offset = tzSign * (tzHours * 60 + tzMinutes);

            // Create date in UTC and adjust for timezone
            const utc = Date.UTC(year, month, day, hour, minute, second);
            return new Date(utc - offset * 60000);
        }

        return new Date(year, month, day, hour, minute, second);
    } catch {
        // Fallback: try native Date parsing
        try {
            return new Date(dateString);
        } catch {
            return undefined;
        }
    }
}

/**
 * Calculates statistics about font sizes in the document.
 * Used for heading detection heuristics.
 */
function calculateFontStats(pageItems: PageItem[][]): { median: number; max: number } {
    const sizes: number[] = [];

    for (const page of pageItems) {
        for (const item of page) {
            if (item.type === 'text' && item.height > 0) {
                sizes.push(item.height);
            }
        }
    }

    if (sizes.length === 0) return { median: 12, max: 12 };

    sizes.sort((a, b) => a - b);
    const median = sizes[Math.floor(sizes.length / 2)];
    const max = sizes[sizes.length - 1];

    return { median, max };
}

/**
 * Determines if a text item should be considered a heading based on its font size.
 * 
 * Heuristic: Text that is at least 20% larger than the median body text size
 * is considered a heading. The level (1-6) is determined by relative size.
 * 
 * @param fontSize - The font size of the text
 * @param fontStats - Statistics about fonts in the document
 * @returns Heading level (1-6) or 0 if not a heading
 */
function detectHeadingLevel(fontSize: number, fontStats: { median: number; max: number }): number {
    // If font is less than 20% larger than median, it's not a heading
    if (fontSize <= fontStats.median * 1.2) return 0;

    // Calculate heading level based on how much larger than median
    const ratio = fontSize / fontStats.median;

    if (ratio >= 2.0) return 1;      // 2x or more = H1
    if (ratio >= 1.7) return 2;      // 1.7x-2x = H2
    if (ratio >= 1.5) return 3;      // 1.5x-1.7x = H3
    if (ratio >= 1.35) return 4;     // 1.35x-1.5x = H4
    if (ratio >= 1.2) return 5;      // 1.2x-1.35x = H5

    return 0;  // Below threshold
}

/**
 * Checks if a text position falls within a link annotation rectangle.
 */
/**
 * Checks if a text item falls within a link annotation rectangle.
 * Uses the center point of the text item to avoid false positives for text
 * that starts immediately after a link (which would share the same x coordinate boundary).
 */
function findLinkForText(item: TextPageItem, annotations: LinkAnnotation[]): LinkAnnotation | undefined {
    const centerX = item.x + (item.width / 2);
    const centerY = item.y + (item.height / 2);

    for (const annot of annotations) {
        // PDF annotation rects are [x1, y1, x2, y2] in page coordinates
        const [x1, y1, x2, y2] = annot.rect;
        const minX = Math.min(x1, x2);
        const maxX = Math.max(x1, x2);
        const minY = Math.min(y1, y2);
        const maxY = Math.max(y1, y2);

        // Check if center point is within the rect (strict, no tolerance)
        if (centerX >= minX && centerX <= maxX &&
            centerY >= minY && centerY <= maxY) {
            return annot;
        }
    }
    return undefined;
}

/**
 * Encodes raw RGBA data into a 24-bit BMP buffer with a white background.
 * Transparency (alpha channel) is flattened against white.
 * 
 * @param width - Image width
 * @param height - Image height
 * @param data - RGBA pixel data
 * @returns BMP Buffer
 */
function encodeBmp(width: number, height: number, data: Uint8Array | Uint8ClampedArray): Buffer {
    // BMP row size must be a multiple of 4 bytes
    const rowSize = Math.floor((24 * width + 31) / 32) * 4;
    const padding = rowSize - (width * 3);
    const headerSize = 54; // 14 (File Header) + 40 (DIB Header)
    const imageSize = rowSize * height;
    const fileSize = headerSize + imageSize;
    const buffer = Buffer.alloc(fileSize);

    // --- File Header (14 bytes) ---
    buffer.write('BM', 0);             // Signature
    buffer.writeUInt32LE(fileSize, 2); // File Size
    buffer.writeUInt32LE(0, 6);        // Reserved
    buffer.writeUInt32LE(headerSize, 10); // Offset to pixel data

    // --- DIB Header (BITMAPINFOHEADER - 40 bytes) ---
    buffer.writeUInt32LE(40, 14);      // Header Size
    buffer.writeInt32LE(width, 18);    // Width
    buffer.writeInt32LE(-height, 22);  // Height (negative for top-down)
    buffer.writeUInt16LE(1, 26);       // Planes
    buffer.writeUInt16LE(24, 28);      // Bit Count (24-bit RGB)
    buffer.writeUInt32LE(0, 30);       // Compression (BI_RGB)
    buffer.writeUInt32LE(imageSize, 34); // Image Size
    buffer.writeInt32LE(2835, 38);     // X PixelsPerMeter (72 DPI)
    buffer.writeInt32LE(2835, 42);     // Y PixelsPerMeter (72 DPI)
    buffer.writeUInt32LE(0, 46);       // Colors Used
    buffer.writeUInt32LE(0, 50);       // Colors Important

    // --- Pixel Data ---
    let offset = headerSize;
    for (let y = 0; y < height; y++) {
        for (let x = 0; x < width; x++) {
            const i = (y * width + x) * 4;
            // RGBA input
            const r = data[i + 0];
            const g = data[i + 1];
            const b = data[i + 2];
            const a = data[i + 3];

            // Flatten alpha against white background
            // out = alpha * pixel + (1 - alpha) * white
            // white = 255
            const alpha = a / 255;
            const outR = Math.round(r * alpha + 255 * (1 - alpha));
            const outG = Math.round(g * alpha + 255 * (1 - alpha));
            const outB = Math.round(b * alpha + 255 * (1 - alpha));

            // Write as BGR (BMP standard)
            buffer[offset + 0] = outB;
            buffer[offset + 1] = outG;
            buffer[offset + 2] = outR;
            offset += 3;
        }
        // Write padding
        for (let p = 0; p < padding; p++) {
            buffer[offset] = 0;
            offset++;
        }
    }

    return buffer;
}

/**
 * Converts raw PDF image data to a buffer for attachment extraction.
 * 
 * **Important Limitation:**
 * PDF images are stored as raw pixel data (RGB, RGBA, or grayscale), not as encoded
 * image files like PNG or JPEG. This function converts the raw data to a normalized
 * RGBA buffer, but this is NOT a valid image file format.
 * 
 * For display, the raw RGBA data would need to be encoded to PNG/JPEG, which requires
 * an additional library like `sharp` or `pngjs`. Currently, this is stored as raw bytes.
 * 
 * OCR is NOT supported for PDF images because Tesseract.js requires encoded image files
 * (PNG, JPEG, etc.), not raw pixel data. To enable OCR, a PNG encoder would need to be added.
 * 
 * @param data - Raw pixel data from PDF.js
 * @param width - Image width in pixels
 * @param height - Image height in pixels
 * @param kind - PDF.js image kind (1=Grayscale, 2=RGB, 3=RGBA)
 * @returns Buffer containing RGBA pixel data (NOT an encoded image file)
 */
function convertToRgbaBuffer(data: Uint8Array | Uint8ClampedArray, width: number, height: number, kind?: number): Buffer {
    // PDF.js image kind values:
    // 1 = GRAYSCALE
    // 2 = RGB  
    // 3 = RGBA

    let rgbaData: Uint8ClampedArray;

    if (kind === 1) {
        // Grayscale - expand to RGBA
        rgbaData = new Uint8ClampedArray(width * height * 4);
        for (let i = 0; i < width * height; i++) {
            const gray = data[i];
            rgbaData[i * 4] = gray;
            rgbaData[i * 4 + 1] = gray;
            rgbaData[i * 4 + 2] = gray;
            rgbaData[i * 4 + 3] = 255;
        }
    } else if (kind === 2 || data.length === width * height * 3) {
        // RGB - add alpha channel
        rgbaData = new Uint8ClampedArray(width * height * 4);
        for (let i = 0; i < width * height; i++) {
            rgbaData[i * 4] = data[i * 3];
            rgbaData[i * 4 + 1] = data[i * 3 + 1];
            rgbaData[i * 4 + 2] = data[i * 3 + 2];
            rgbaData[i * 4 + 3] = 255;
        }
    } else {
        // Assume RGBA
        rgbaData = data instanceof Uint8ClampedArray ? data : new Uint8ClampedArray(data);
    }

    return Buffer.from(rgbaData.buffer, rgbaData.byteOffset, rgbaData.byteLength);
}

/**
 * Parses a PDF file and extracts content.
 * 
 * @param buffer - The PDF file buffer
 * @param config - Parser configuration
 * @returns Promise resolving to the parsed AST
 */
export const parsePdf = async (buffer: Buffer, config: OfficeParserConfig): Promise<OfficeParserAST> => {
    let pdfjs: any;

    // Check if we are in a Node.js environment
    // @ts-ignore
    if (typeof window === 'undefined') {
        // Helper to bypass TS/Webpack/Other transpilers converting import() to require() when compiling to CJS
        // Defined here to avoid CSP 'unsafe-eval' issues in browser environments
        const dynamicImport = new Function('specifier', 'return import(specifier)');

        try {
            // Use legacy build for Node.js (required for pdfjs-dist v5+)
            pdfjs = await dynamicImport('pdfjs-dist/legacy/build/pdf.mjs');
        } catch (e) {
            // Fallback to standard if legacy not found (e.g. older versions)
            pdfjs = await dynamicImport('pdfjs-dist');
        }
    } else {
        pdfjs = await import('pdfjs-dist');
    }

    // Configure worker

    if (config.pdfWorkerSrc) {
        pdfjs.GlobalWorkerOptions.workerSrc = config.pdfWorkerSrc;
    } else {
        // Fallbacks when no workerSrc is provided
        // @ts-ignore
        if (typeof window !== 'undefined') {
            // Browser: Default to CDN
            pdfjs.GlobalWorkerOptions.workerSrc = 'https://unpkg.com/pdfjs-dist@5.4.530/build/pdf.worker.min.mjs';
        } else {
            // Node.js: Try to auto-resolve local worker path to avoid remote fetch errors
            try {
                // We use require.resolve to find the exact path of the installed package.
                // @ts-ignore - 'require' is available in Node.js/CommonJS environment
                const workerPath = require.resolve('pdfjs-dist/legacy/build/pdf.worker.mjs');
                pdfjs.GlobalWorkerOptions.workerSrc = workerPath;
            } catch (e) {
                if (config.outputErrorToConsole) console.warn("[PdfParser] Could not auto-resolve local worker path:", e);
            }
        }
    }

    const uint8Array = new Uint8Array(buffer);
    const loadingTask = pdfjs.getDocument({
        data: uint8Array,
        verbosity: 0 // ERRORS only, suppresses warnings
    });

    // Handle loading errors, specifically missing worker in browser
    let pdfDocument;
    try {
        pdfDocument = await loadingTask.promise;
    } catch (e: any) {
        if (e.message?.includes('workerSrc') || e.message?.includes('No "GlobalWorkerOptions.workerSrc" specified')) {
            throw getOfficeError(OfficeErrorType.PDF_WORKER_MISSING, config);
        }
        throw e;
    }

    const content: OfficeContentNode[] = [];
    const attachments: OfficeAttachment[] = [];
    const numPages = pdfDocument.numPages;

    // Collect all page items for font statistics before processing
    const allPageItems: PageItem[][] = [];

    // --- Metadata Extraction ---
    // Extract all available metadata from the PDF info dictionary.
    // Note: Some metadata fields depend on how the PDF was created.
    // - Producer: Software that created the PDF
    // - Creator: Application that made the original document
    // - Keywords, Description are rarely present
    const meta = await pdfDocument.getMetadata().catch(() => ({ info: {} }));
    const info = meta.info as Record<string, any>;

    const metadata: OfficeMetadata = {
        pages: numPages,
        title: info?.Title || undefined,
        author: info?.Author || undefined,
        subject: info?.Subject || undefined,
        description: info?.Keywords || undefined,  // Map Keywords to description as closest match
        created: parsePdfDate(info?.CreationDate),
        modified: parsePdfDate(info?.ModDate),
        // Note: lastModifiedBy is not available in PDF format - there's no concept of "last modifier"
        // The Author field only tracks original author.
    };

    // Extract non-standard entries from the PDF Info dictionary as custom properties.
    // The standard keys are defined by the PDF spec; anything else is user/tool-defined.
    const standardPdfInfoKeys = new Set([
        'Title', 'Author', 'Subject', 'Keywords', 'Creator', 'Producer',
        'CreationDate', 'ModDate', 'Trapped', 'IsAcroFormPresent', 'IsXFAPresent',
        'IsCollectionPresent', 'IsSignaturesPresent', 'PDFFormatVersion'
    ]);
    if (info) {
        const customProperties: Record<string, string | number | boolean | Date> = {};
        for (const key of Object.keys(info)) {
            if (standardPdfInfoKeys.has(key)) continue;
            const val = info[key];
            if (val === null || val === undefined) continue;
            // pdf.js groups document-level custom metadata under a 'Custom' object.
            // Flatten its entries directly into customProperties.
            if (key === 'Custom' && typeof val === 'object' && !Array.isArray(val) && !(val instanceof Date)) {
                for (const [customKey, customVal] of Object.entries(val)) {
                    if (customVal === null || customVal === undefined) continue;
                    if (typeof customVal === 'string' || typeof customVal === 'number' || typeof customVal === 'boolean' || customVal instanceof Date) {
                        customProperties[customKey] = customVal;
                    }
                }
                continue;
            }
            if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean' || val instanceof Date) {
                customProperties[key] = val;
            }
        }
        if (Object.keys(customProperties).length > 0) {
            metadata.customProperties = customProperties;
        }
    }

    // --- Embedded File Attachment Extraction ---
    /**
     * PDF can contain embedded files (not images in content, but attached files).
     * These are separate from images in the page content stream.
     */
    try {
        const embeddedFiles = await pdfDocument.getAttachments();
        if (embeddedFiles && config.extractAttachments) {
            for (const name in embeddedFiles) {
                const file = embeddedFiles[name];
                const fileBuffer = Buffer.from(file.content);
                const attachment = createAttachment(file.filename, fileBuffer);
                attachments.push(attachment);
            }
        }
    } catch (e) {
        if (config.outputErrorToConsole) console.error("Error extracting embedded attachments:", e);
    }

    // --- First Pass: Collect all items for font statistics ---
    for (let i = 1; i <= numPages; i++) {
        let page: any;
        let textContent;
        const pageItems: PageItem[] = [];

        try {
            page = await pdfDocument.getPage(i);
            // Extract text content
            textContent = await page.getTextContent();
        } catch (e: any) {
            if (config.outputErrorToConsole) console.warn(`[PdfParser] Error loading page ${i}:`, e);
            // Push empty items to maintain index alignment for second pass
            allPageItems.push(pageItems);
            continue;
        }

        const commonObjs = page.commonObjs;

        for (const item of textContent.items) {
            const transform = item.transform;
            const x = transform[4];
            const y = transform[5];
            const width = item.width || 0;
            const height = item.height || Math.abs(transform[3]) || 12;

            // Extract formatting from font
            const formatting: TextFormatting = {};
            let fontName: string | undefined;

            if (item.fontName && commonObjs) {
                try {
                    if (commonObjs.has(item.fontName)) {
                        // Use callback-based get to ensure safe resolution
                        const fontData = await new Promise<any>((resolve) => {
                            // @ts-ignore
                            commonObjs.get(item.fontName, (data: any) => resolve(data));
                        });
                        if (fontData?.name) {
                            // Remove PDF subset prefix (6 uppercase letters + '+')
                            fontName = fontData.name.replace(/^[A-Z]{6}\+/, '');
                            formatting.font = fontName;

                            // Detect bold/italic from font name
                            const lowerName = fontData.name.toLowerCase();
                            if (lowerName.includes('bold')) formatting.bold = true;
                            if (lowerName.includes('italic') || lowerName.includes('oblique')) formatting.italic = true;
                        }
                    }
                } catch {
                    // Font lookup failed, continue without font info
                }
            }

            if (height > 0) {
                formatting.size = Math.round(height).toString();
            }

            pageItems.push({
                type: 'text',
                x,
                y,
                width,
                height,
                text: item.str,
                fontName,
                formatting
            });
        }

        // Extract images if enabled
        if (config.extractAttachments || config.ocr) {
            try {
                const ops = await page.getOperatorList();
                const fnArray = ops.fnArray;
                const argsArray = ops.argsArray;

                for (let j = 0; j < fnArray.length; j++) {
                    const fn = fnArray[j];

                    if (fn === pdfjs.OPS.dependency) {
                        const deps = argsArray[j];
                        for (const dep of deps) {
                            // In pdfjs-dist v3+, get() throws if not resolved unless a callback is provided.
                            // We must use the callback pattern to wait for resolution.
                            try {
                                if (page.objs.has(dep)) continue;

                                await new Promise<void>((resolve) => {
                                    const timeout = setTimeout(() => {
                                        resolve();
                                    }, 500);

                                    page.objs.get(dep, (data: any) => {
                                        clearTimeout(timeout);
                                        resolve();
                                    });
                                });
                            } catch (e: any) {
                                if (config.outputErrorToConsole) {
                                    console.error(`[PdfParser] Failed to load dependency ${dep}:`, e);
                                }
                            }
                        }
                    }

                    if (fn === pdfjs.OPS.paintImageXObject || fn === pdfjs.OPS.paintXObject) {
                        const imgName = argsArray[j][0];

                        try {
                            let hasObj = page.objs.has(imgName);
                            let targetObjs = page.objs;

                            if (!hasObj && page.commonObjs.has(imgName)) {
                                hasObj = true;
                                targetObjs = page.commonObjs;
                            }

                            if (hasObj) {
                                // Use callback-based get to ensure safe resolution
                                const rawObj = await new Promise<any>((resolve) => {
                                    // @ts-ignore
                                    targetObjs.get(imgName, (data: any) => resolve(data));
                                });
                                const imgObj = rawObj as any;

                                // Browser-specific: Handle ImageBitmap if data is missing
                                if (typeof window !== 'undefined' && !imgObj.data && imgObj.bitmap) {
                                    try {
                                        const canvas = document.createElement('canvas');
                                        canvas.width = imgObj.width;
                                        canvas.height = imgObj.height;
                                        const ctx = canvas.getContext('2d');
                                        if (ctx) {
                                            ctx.drawImage(imgObj.bitmap, 0, 0);
                                            imgObj.data = ctx.getImageData(0, 0, imgObj.width, imgObj.height).data;
                                            imgObj.kind = 3; // RGBA
                                        }
                                    } catch (e) {
                                        if (config.outputErrorToConsole) console.error(`[PdfParser] Failed to extract from ImageBitmap:`, e);
                                    }
                                }

                                if (imgObj?.data && imgObj.width > 0 && imgObj.height > 0) {
                                    // Find position from transform matrix
                                    let imgX = 0, imgY = 0;
                                    for (let k = j - 1; k >= 0; k--) {
                                        if (fnArray[k] === pdfjs.OPS.transform) {
                                            imgX = argsArray[k][4];
                                            imgY = argsArray[k][5];
                                            break;
                                        }
                                    }

                                    pageItems.push({
                                        type: 'image',
                                        x: imgX,
                                        y: imgY,
                                        name: imgName,
                                        data: imgObj.data,
                                        width: imgObj.width,
                                        height: imgObj.height,
                                        kind: imgObj.kind
                                    });
                                }
                            }
                        } catch {
                            // Image access failed, continue
                        }
                    }
                }
            } catch (e) {
                if (config.outputErrorToConsole) console.error(`Error extracting images from page ${i}:`, e);
            }
        }

        allPageItems.push(pageItems);
    }

    // Calculate font statistics for heading detection
    const fontStats = calculateFontStats(allPageItems);

    // --- Second Pass: Process pages with font statistics ---
    for (let i = 0; i < allPageItems.length; i++) {
        const pageNum = i + 1;
        let page: any;
        try {
            page = await pdfDocument.getPage(pageNum);
        } catch (e: any) {
            if (config.outputErrorToConsole) console.warn(`[PdfParser] Error loading page ${pageNum} in second pass:`, e);
            continue;
        }

        const pageItems = allPageItems[i];
        const pageContent: OfficeContentNode[] = [];

        // Extract link annotations for this page
        const annotations: LinkAnnotation[] = [];
        try {
            const annots = await page.getAnnotations();
            for (const annot of annots) {
                if (annot.subtype === 'Link' && annot.rect) {
                    annotations.push({
                        rect: annot.rect,
                        url: annot.url,
                        dest: annot.dest
                    });
                }
            }
        } catch (e) {
            if (config.outputErrorToConsole) console.error(`Error extracting annotations from page ${pageNum}:`, e);
        }

        // Sort items: Y descending (top to bottom), then X ascending (left to right)
        pageItems.sort((a, b) => {
            if (Math.abs(b.y - a.y) > 5) return b.y - a.y;
            return a.x - b.x;
        });

        // Process sorted items into content nodes
        let currentNode: OfficeContentNode | null = null;
        let currentNodeFontSize = 0;
        let lastY = -1;
        let imageCounter = 0;

        for (const item of pageItems) {
            if (item.type === 'text') {
                const text = item.text;
                if (!text) continue;

                // Check for new line
                const isNewLine = lastY !== -1 && Math.abs(item.y - lastY) > 5;

                if (isNewLine && currentNode) {
                    // Finalize and push current node
                    if ((currentNode.text || '').trim().length > 0) {
                        pageContent.push(currentNode);
                    }
                    currentNode = null;
                }

                // Skip pure whitespace at start of lines
                if (!currentNode && text.trim().length === 0) {
                    lastY = item.y;
                    continue;
                }

                // Determine if this should be a heading
                const headingLevel = detectHeadingLevel(item.height, fontStats);

                if (!currentNode) {
                    // Start new node
                    if (headingLevel > 0) {
                        currentNode = {
                            type: 'heading',
                            text: '',
                            children: [],
                            metadata: { level: headingLevel }
                        };
                    } else {
                        currentNode = {
                            type: 'paragraph',
                            text: '',
                            children: []
                        };
                    }
                    currentNodeFontSize = item.height;
                }

                // Handle whitespace
                if (text.trim().length === 0) {
                    if (currentNode.children && currentNode.children.length > 0) {
                        const lastChild = currentNode.children[currentNode.children.length - 1];
                        if (lastChild.type === 'text' && lastChild.text) {
                            lastChild.text += text;
                            currentNode.text += text;
                        }
                    }
                    lastY = item.y;
                    continue;
                }

                // Add space between words if needed
                if (currentNode.text && currentNode.text.length > 0 && !currentNode.text.endsWith(' ')) {
                    currentNode.text += ' ';
                    if (currentNode.children && currentNode.children.length > 0) {
                        const lastChild = currentNode.children[currentNode.children.length - 1];
                        if (lastChild.type === 'text' && lastChild.text) {
                            lastChild.text += ' ';
                        }
                    }
                }

                currentNode.text += text;

                // Check for link
                const link = findLinkForText(item, annotations);
                let textMetadata: TextMetadata | undefined;
                if (link) {
                    if (link.url) {
                        textMetadata = {
                            link: link.url,
                            linkType: link.url.startsWith('#') ? 'internal' : 'external'
                        };
                    } else if (link.dest) {
                        // Internal destination
                        textMetadata = {
                            link: typeof link.dest === 'string' ? `#${link.dest}` : '#internal',
                            linkType: 'internal'
                        };
                    }
                }

                // Try to merge with last child if same formatting and no link change
                let merged = false;
                if (currentNode.children && currentNode.children.length > 0 && !textMetadata) {
                    const lastChild = currentNode.children[currentNode.children.length - 1];
                    if (lastChild.type === 'text' &&
                        isSameFormatting(lastChild.formatting, item.formatting) &&
                        !lastChild.metadata) {
                        lastChild.text = (lastChild.text || '') + text;
                        merged = true;
                    }
                }

                if (!merged) {
                    const textNode: OfficeContentNode = {
                        type: 'text',
                        text: text,
                        formatting: Object.keys(item.formatting).length > 0 ? item.formatting : undefined
                    };
                    if (textMetadata) {
                        textNode.metadata = textMetadata;
                    }
                    currentNode.children?.push(textNode);
                }

                lastY = item.y;

            } else if (item.type === 'image') {
                // Flush current node
                if (currentNode) {
                    if ((currentNode.text || '').trim().length > 0) {
                        pageContent.push(currentNode);
                    }
                    currentNode = null;
                }

                imageCounter++;
                // Note: Using .bmp extension since we encode to BMP for broad compatibility
                const attachmentName = `pdf_image_p${pageNum}_${imageCounter}.bmp`;

                /**
                 * Image extraction for PDF files.
                 * 
                 * PDF stores images as raw pixel data. We convert to BMP for compatibility.
                 */
                if (config.extractAttachments) {
                    try {
                        const imageBuffer = convertToRgbaBuffer(
                            item.data,
                            item.width,
                            item.height,
                            item.kind
                        );

                        // Encode as BMP
                        const bmpBuffer = encodeBmp(item.width, item.height, new Uint8Array(imageBuffer));
                        const attachment = createAttachment(attachmentName, bmpBuffer);
                        attachment.mimeType = 'image/bmp';

                        // Perform OCR if enabled
                        if (config.ocr) {
                            try {
                                // Skip OCR for very small images/artifacts (e.g. < 10px) to avoid Tesseract warnings
                                if (item.width >= 10 && item.height >= 10) {
                                    attachment.ocrText = (await performOcr(bmpBuffer, config.ocrLanguage)).trim();
                                }
                            } catch (e) {
                                if (config.outputErrorToConsole) console.error(`OCR failed for ${attachmentName}:`, e);
                            }
                        }

                        attachments.push(attachment);

                        // Create image content node
                        const imageMetadata: ImageMetadata = {
                            attachmentName,
                        };

                        pageContent.push({
                            type: 'image',
                            text: attachment.ocrText || '',
                            metadata: { ...imageMetadata }
                        });

                    } catch (e) {
                        logWarning(`Failed to process image ${attachmentName}:`, config, e);
                    }
                }
            }
        }

        // Flush last node
        if (currentNode && (currentNode.text || '').trim().length > 0) {
            pageContent.push(currentNode);
        }

        // Add page node to content
        content.push({
            type: 'page',
            children: pageContent,
            text: pageContent.map(node => node.text).join(config.newlineDelimiter ?? '\n\n'),
            metadata: { pageNumber: pageNum }
        });
    }

    return {
        type: 'pdf',
        metadata: metadata,
        content: content,
        attachments: attachments,
        toText: () => content.map(c => c.text).join(config.newlineDelimiter ?? '\n\n')
    };
};

/**
 * Helper to compare two text formatting objects.
 * Returns true if both have the same properties with the same values.
 */
function isSameFormatting(a: TextFormatting | undefined, b: TextFormatting | undefined): boolean {
    if (!a && !b) return true;
    if (!a || !b) return false;

    const keysA = Object.keys(a).sort();
    const keysB = Object.keys(b).sort();

    if (keysA.length !== keysB.length) return false;

    for (let i = 0; i < keysA.length; i++) {
        const key = keysA[i] as keyof TextFormatting;
        if (keysA[i] !== keysB[i]) return false;
        if (a[key] !== b[key]) return false;
    }

    return true;
}

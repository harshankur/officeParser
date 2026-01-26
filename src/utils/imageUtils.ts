/**
 * Image Processing Utilities
 * 
 * Provides helper functions for working with image attachments extracted from office documents.
 * Handles MIME type conversions, file extension mapping, and attachment object creation.
 * 
 * @module imageUtils
 */

import { OfficeAttachment, OfficeMimeType } from '../types';

/**
 * Converts a file extension to its corresponding MIME type.
 * 
 * Used when creating attachments to determine the MIME type from a filename.
 * The extension check is case-insensitive.
 * 
 * @param ext - The file extension (with or without a dot, e.g., 'png', '.png')
 * @returns The corresponding MIME type string
 * @example
 * ```typescript
 * getMimeFromExtension('png'); // Returns 'image/png'
 * getMimeFromExtension('JPG'); // Returns 'image/jpeg' (case-insensitive)
 * getMimeFromExtension('unknown'); // Returns 'application/octet-stream'
 * ```
 */
export const getMimeFromExtension = (ext: string): string => {
    switch (ext.toLowerCase()) {
        case 'jpg':
        case 'jpeg': return 'image/jpeg';
        case 'png': return 'image/png';
        case 'gif': return 'image/gif';
        case 'bmp': return 'image/bmp';
        case 'tiff': return 'image/tiff';
        case 'webp': return 'image/webp';
        default: return 'application/octet-stream'; // Generic binary MIME type
    }
};

/**
 * Detects the MIME type from file magic bytes (file signature).
 * 
 * This is useful for files with incorrect or missing extensions (like .tmp files).
 * Inspects the first few bytes of the file to determine the actual format.
 * 
 * @param buffer - The file content as a Buffer
 * @returns The detected MIME type, or undefined if not recognized
 * @example
 * ```typescript
 * const pngBuffer = fs.readFileSync('image.tmp');
 * getMimeFromBytes(pngBuffer); // Returns 'image/png' if it's a PNG file
 * ```
 */
export const getMimeFromBytes = (buffer: Buffer): string | undefined => {
    if (buffer.length < 4) return undefined;

    // PNG: 89 50 4E 47 (0x89 "PNG")
    if (buffer[0] === 0x89 && buffer[1] === 0x50 && buffer[2] === 0x4E && buffer[3] === 0x47) {
        return 'image/png';
    }

    // JPEG: FF D8 FF
    if (buffer[0] === 0xFF && buffer[1] === 0xD8 && buffer[2] === 0xFF) {
        return 'image/jpeg';
    }

    // GIF: 47 49 46 38 ("GIF8")
    if (buffer[0] === 0x47 && buffer[1] === 0x49 && buffer[2] === 0x46 && buffer[3] === 0x38) {
        return 'image/gif';
    }

    // BMP: 42 4D ("BM")
    if (buffer[0] === 0x42 && buffer[1] === 0x4D) {
        return 'image/bmp';
    }

    // TIFF: 49 49 2A 00 (little endian) or 4D 4D 00 2A (big endian)
    if ((buffer[0] === 0x49 && buffer[1] === 0x49 && buffer[2] === 0x2A && buffer[3] === 0x00) ||
        (buffer[0] === 0x4D && buffer[1] === 0x4D && buffer[2] === 0x00 && buffer[3] === 0x2A)) {
        return 'image/tiff';
    }

    // WebP: 52 49 46 46 ... 57 45 42 50 ("RIFF" ... "WEBP")
    if (buffer.length >= 12 &&
        buffer[0] === 0x52 && buffer[1] === 0x49 && buffer[2] === 0x46 && buffer[3] === 0x46 &&
        buffer[8] === 0x57 && buffer[9] === 0x45 && buffer[10] === 0x42 && buffer[11] === 0x50) {
        return 'image/webp';
    }

    return undefined;
};

/**
 * Creates an OfficeAttachment object from image data.
 * 
 * This is a convenience function that:
 * 1. Extracts the file extension from the filename
 * 2. Determines the MIME type from the extension
 * 3. Encodes the image buffer as Base64
 * 4. Constructs a properly formatted OfficeAttachment object
 * 
 * @param name - The filename of the image (e.g., 'image1.png', 'chart.jpg')
 * @param content - The image data as a Node.js Buffer
 * @returns An OfficeAttachment object ready to be added to the attachments array
 * 
 * @example
 * ```typescript
 * const imageBuffer = fs.readFileSync('photo.png');
 * const attachment = createAttachment('photo.png', imageBuffer);
 * 
 * console.log(attachment.type); // 'image'
 * console.log(attachment.mimeType); // 'image/png'
 * console.log(attachment.name); // 'photo.png'
 * console.log(attachment.extension); // 'png'
 * console.log(attachment.data); // 'iVBORw0KGgoAAAANSUhEUgAA...' (Base64)
 * ```
 */
export const createAttachment = (name: string, content: Buffer): OfficeAttachment => {
    // Step 1: Extract file extension from the filename
    // Example: 'image1.png' -> 'png'
    const ext = name.split('.').pop() || '';

    // Step 2: Try to detect MIME type from magic bytes first (more reliable)
    // This handles cases like .tmp files or mismatched extensions
    let mime = getMimeFromBytes(content);

    // Step 3: Fallback to extension if magic bytes couldn't detect it
    // Useful for SVG or formats not covered by getMimeFromBytes
    if (!mime) {
        mime = getMimeFromExtension(ext);
    }

    // Step 4: Create and return the attachment object
    return {
        type: 'image', // All attachments created here are images
        mimeType: mime as OfficeMimeType, // Cast to our supported MIME types
        data: content.toString('base64'), // Convert buffer to Base64 string
        name: name, // Original filename
        extension: ext // File extension for reference
    };
};

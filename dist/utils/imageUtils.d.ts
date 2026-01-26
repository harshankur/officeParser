/**
 * Image Processing Utilities
 *
 * Provides helper functions for working with image attachments extracted from office documents.
 * Handles MIME type conversions, file extension mapping, and attachment object creation.
 *
 * @module imageUtils
 */
/// <reference types="node" />
import { OfficeAttachment } from '../types';
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
export declare const getMimeFromExtension: (ext: string) => string;
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
export declare const getMimeFromBytes: (buffer: Buffer) => string | undefined;
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
export declare const createAttachment: (name: string, content: Buffer) => OfficeAttachment;

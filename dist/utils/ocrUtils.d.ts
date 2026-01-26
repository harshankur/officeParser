/**
 * OCR (Optical Character Recognition) Utilities
 *
 * This module provides functions for extracting text from images using Tesseract.js.
 * Used when `config.ocr` is enabled to extract text from embedded images in documents.
 *
 * @module ocrUtils
 */
/// <reference types="node" />
/**
 * Performs Optical Character Recognition (OCR) on an image to extract text.
 *
 * Uses Tesseract.js to recognize text in the provided image buffer.
 * This is useful for extracting text from screenshots, scanned documents,
 * charts with labels, or any image containing text.
 *
 * The function creates a new Tesseract worker, processes the image,
 * and properly terminates the worker to free resources.
 *
 * @param imageBuffer - The image data as a Node.js Buffer (PNG, JPEG, etc.)
 * @param language - The language code for OCR (default: 'eng' for English).
 *                   Supports ISO 639-2/T three-letter codes: 'eng', 'spa', 'fra', 'deu', etc.
 *                   Multiple languages can be combined with '+': 'eng+fra'
 * @returns A promise that resolves to the recognized text as a string
 * @throws {Error} If the image cannot be processed or Tesseract initialization fails
 *
 * @example
 * ```typescript
 * // Extract text from an English image
 * const text = await performOcr(imageBuffer, 'eng');
 * console.log(text); // "Annual Revenue: $1.2M"
 *
 * // Extract text from a multilingual image
 * const text = await performOcr(imageBuffer, 'eng+spa');
 * ```
 *
 * @see https://github.com/naptha/tesseract.js for supported languages and options
 */
export declare const performOcr: (image: Buffer | string, language?: string) => Promise<string>;

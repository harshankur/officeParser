/**
 * OCR (Optical Character Recognition) Utilities
 * 
 * This module provides functions for extracting text from images using Tesseract.js.
 * Used when `config.ocr` is enabled to extract text from embedded images in documents.
 * 
 * @module ocrUtils
 */

import Tesseract from "tesseract.js";

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
 * @param worker - The Tesseract worker to use for OCR.
 * @returns A promise that resolves to the recognized text as a string
 * @throws {Error} If the image cannot be processed or Tesseract initialization fails
 * 
 * @example
 * ```typescript
 * // Extract text using a Tesseract worker
 * const text = await performOcr(imageBuffer, worker);
 * console.log(text); // "Annual Revenue: $1.2M"
 * 
 * ```
 * 
 * @see https://github.com/naptha/tesseract.js for supported languages and options
 */
export const performOcr = async (image: Buffer | string, worker: Tesseract.Worker): Promise<string> => {

  // Step 1: Prepare image data
  let inputImage: any = image;

  // In browser environment, convert Buffer to Blob for better compatibility
  // @ts-ignore
  if (typeof window !== 'undefined' && typeof Blob !== 'undefined' && Buffer.isBuffer(image)) {
    inputImage = new Blob([image as any], { type: 'image/bmp' });
  }

  // Step 2: Perform OCR
  const ret = await worker.recognize(inputImage);

  return ret.data.text;
};

/**
 * ZIP Archive Extraction Utilities
 * 
 * Provides functions for extracting files from ZIP archives.
 * Essential for parsing OOXML (DOCX, XLSX, PPTX) and ODF (ODT, ODP, ODS) files,
 * which are all ZIP archives containing XML and media files.
 * 
 * Office File Structure:
 * - DOCX: ZIP containing word/document.xml, word/styles.xml, word/media/*, etc.
 * - XLSX: ZIP containing xl/workbook.xml, xl/worksheets/sheet1.xml, etc.
 * - PPTX: ZIP containing ppt/slides/slide1.xml, ppt/media/*, etc.
 * - ODF: Similar structure with content.xml, styles.xml, etc.
 * 
 * @module zipUtils
 */

import { unzip } from 'fflate';
import { DecompressionLimits, OfficeErrorType } from '../types.js';
import { getOfficeError } from './errorUtils.js';

/**
 * Represents a file extracted from a ZIP archive.
 * Contains the file's path within the archive and its content as a Buffer.
 */
interface ZipFileContent {
    /**
     * The relative path of the file within the ZIP archive.
     * @example "word/document.xml", "xl/worksheets/sheet1.xml", "ppt/slides/slide1.xml"
     */
    path: string;

    /**
     * The file content as a Node.js Buffer.
     * Can be converted to string for XML files or used directly for binary files (images, etc.).
     * @example Buffer containing XML text or binary image data
     */
    content: Buffer;
}

/**
 * Extracts files from a ZIP archive with optional filtering.
 * 
 * This function:
 * 1. Opens the ZIP archive from a Buffer
 * 2. Iterates through all entries in the archive
 * 3. Applies a filter function to determine which files to extract
 * 4. Extracts matching files and returns them as an array
 * 
 * Uses lazy entry reading for better memory efficiency with large archives.
 * Files are extracted asynchronously and collected into an array.
 * 
 * @param zipInput - The ZIP file as a Node.js Buffer
 * @param filterFn - A predicate function to determine which files to extract.
 *                   Receives the filename and returns true to extract, false to skip.
 * @returns A promise resolving to an array of extracted files
 * @throws {Error} If the ZIP file cannot be opened or an entry cannot be read
 * 
 * @example
 * ```typescript
 * // Extract only XML files from a DOCX
 * const files = await extractFiles(docxBuffer, (fileName) => fileName.endsWith('.xml'));
 * 
 * // Extract document.xml specifically
 * const files = await extractFiles(docxBuffer, (fileName) => 
 *   fileName === 'word/document.xml'
 * );
 * 
 * // Extract all files
 * const allFiles = await extractFiles(zipBuffer, () => true);
 * 
 * // Extract everything except media files
 * const files = await extractFiles(zipBuffer, (fileName) => 
 *   !fileName.startsWith('word/media/')
 * );
 * ```
 * 
 * @see https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT ZIP file format specification
 */
export const extractFiles = (
    zipInput: Buffer,
    filterFn: (fileName: string) => boolean,
    limits: DecompressionLimits
): Promise<ZipFileContent[]> => {
    const maxUncompressedBytes = limits?.maxUncompressedBytes !== undefined && Number.isFinite(limits.maxUncompressedBytes) && limits.maxUncompressedBytes >= 0
        ? limits.maxUncompressedBytes
        : 512 * 1024 * 1024;

    const maxZipEntries = limits?.maxZipEntries !== undefined && Number.isFinite(limits.maxZipEntries) && limits.maxZipEntries >= 0
        ? limits.maxZipEntries
        : 10000;

    return new Promise((resolve, reject) => {
        let totalEntryCount = 0;
        let entryCount = 0;
        let declaredTotal = 0;

        unzip(
            new Uint8Array(zipInput.buffer, zipInput.byteOffset, zipInput.byteLength),
            {
                filter: (file) => {
                    totalEntryCount++;
                    if (totalEntryCount > maxZipEntries) {
                        reject(getOfficeError(OfficeErrorType.ZIP_ENTRY_COUNT_LIMIT_EXCEEDED, undefined, maxZipEntries));
                        return false;
                    }

                    if (!filterFn(file.name)) return false;

                    if (typeof file.originalSize !== 'number' || !Number.isFinite(file.originalSize) || file.originalSize < 0) {
                        reject(getOfficeError(OfficeErrorType.ZIP_ENTRY_INVALID_SIZE));
                        return false;
                    }

                    entryCount++;
                    declaredTotal += file.originalSize;

                    if (declaredTotal > maxUncompressedBytes) {
                        reject(getOfficeError(OfficeErrorType.ZIP_SIZE_LIMIT_EXCEEDED, undefined, maxUncompressedBytes));
                        return false;
                    }
                    return true;
                }
            },
            (err, decompressed) => {
                if (err) return reject(err);
                resolve(Object.entries(decompressed).map(([path, data]) => ({
                    path,
                    content: Buffer.from(data)
                })));
            }
        );
    });
};

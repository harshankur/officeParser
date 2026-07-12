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

import { Unzip, UnzipInflate } from 'fflate';
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
        // Decompress as a stream and cap on the ACTUAL inflated byte count rather than
        // the size declared in the ZIP header. The declared size is attacker-controlled,
        // so a "zip bomb" can understate it and still inflate to gigabytes; counting real
        // output bytes and aborting once the limit is crossed is the only reliable guard.
        const results: ZipFileContent[] = [];
        let totalEntryCount = 0;
        let actualTotalBytes = 0;
        let pendingFiles = 0;
        let pushComplete = false;
        let settled = false;

        const fail = (err: unknown) => {
            if (settled) return;
            settled = true;
            reject(err);
        };
        const maybeResolve = () => {
            if (!settled && pushComplete && pendingFiles === 0) {
                settled = true;
                resolve(results);
            }
        };

        const unzipper = new Unzip((file) => {
            if (settled) return;
            totalEntryCount++;
            if (totalEntryCount > maxZipEntries) {
                fail(getOfficeError(OfficeErrorType.ZIP_ENTRY_COUNT_LIMIT_EXCEEDED, undefined, maxZipEntries));
                return;
            }
            if (!filterFn(file.name)) return;

            const name = file.name;
            const chunks: Buffer[] = [];
            pendingFiles++;
            file.ondata = (err, chunk, final) => {
                if (settled) return;
                if (err) { fail(err); return; }
                if (chunk && chunk.length) {
                    actualTotalBytes += chunk.length;
                    if (actualTotalBytes > maxUncompressedBytes) {
                        fail(getOfficeError(OfficeErrorType.ZIP_SIZE_LIMIT_EXCEEDED, undefined, maxUncompressedBytes));
                        return;
                    }
                    chunks.push(Buffer.from(chunk));
                }
                if (final) {
                    results.push({ path: name, content: Buffer.concat(chunks) });
                    pendingFiles--;
                    maybeResolve();
                }
            };
            try {
                file.start();
            } catch (e) {
                fail(e);
            }
        });
        unzipper.register(UnzipInflate);

        const src = new Uint8Array(zipInput.buffer, zipInput.byteOffset, zipInput.byteLength);
        // Feed the compressed input in bounded chunks so the actual-size check can abort a
        // bomb early; the worst-case overshoot is one chunk's worth of inflation.
        const PUSH_CHUNK = 1 << 16; // 64 KiB
        try {
            let offset = 0;
            while (offset < src.length && !settled) {
                const end = Math.min(offset + PUSH_CHUNK, src.length);
                unzipper.push(src.subarray(offset, end), end >= src.length);
                offset = end;
            }
        } catch (e) {
            fail(e);
            return;
        }
        pushComplete = true;
        maybeResolve();
    });
};

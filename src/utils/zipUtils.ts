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

import yauzl from 'yauzl';
import concat from 'concat-stream';

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
export const extractFiles = (zipInput: Buffer, filterFn: (fileName: string) => boolean): Promise<ZipFileContent[]> => {
    return new Promise((resolve, reject) => {
        // Step 1: Open the ZIP archive from the buffer
        // lazyEntries: true means we manually control when to read each entry (better memory usage)
        yauzl.fromBuffer(zipInput, { lazyEntries: true }, (err, zipfile) => {
            if (err) return reject(err);
            if (!zipfile) return reject(new Error("Failed to open zip file"));

            // Array to collect all extracted files
            const extractedFiles: ZipFileContent[] = [];

            // Step 2: Start reading the first entry
            // This triggers the 'entry' event
            zipfile.readEntry();

            // Step 3: Handle each entry (file or directory) in the ZIP
            zipfile.on('entry', (entry: yauzl.Entry) => {
                // Step 3a: Check if this file should be extracted using the filter function
                if (filterFn(entry.fileName)) {
                    // Step 3b: Open a read stream for this entry
                    zipfile.openReadStream(entry, (err, readStream) => {
                        if (err) return reject(err);
                        if (!readStream) return reject(new Error("Failed to open read stream"));

                        // Step 3c: Pipe the stream through concat to collect all data into a single Buffer
                        // This is necessary because streams deliver data in chunks
                        readStream.pipe(concat((data: Buffer) => {
                            // Step 3d: Add the extracted file to our results
                            extractedFiles.push({
                                path: entry.fileName,
                                content: data
                            });

                            // Step 3e: Continue to the next entry
                            zipfile.readEntry();
                        }));
                    });
                } else {
                    // Step 3f: Skip this entry and move to the next one
                    zipfile.readEntry();
                }
            });

            // Step 4: All entries have been processed
            zipfile.on('end', () => resolve(extractedFiles));

            // Step 5: Handle any errors during extraction
            zipfile.on('error', reject);
        });
    });
};

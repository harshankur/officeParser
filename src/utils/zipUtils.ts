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
import { DecompressionLimits, OfficeErrorType, OfficeParserConfig } from '../types.js';
import { getOfficeError } from './errorUtils.js';

/**
 * Signature of the End Of Central Directory record ("PK\x05\x06"), the trailer every ZIP
 * archive ends with. Its presence is what distinguishes a complete archive from one that
 * was cut off in transfer.
 */
const EOCD_SIGNATURE = Buffer.from([0x50, 0x4b, 0x05, 0x06]);

/** Size of the fixed portion of an End Of Central Directory record, in bytes. */
const EOCD_RECORD_MIN_BYTES = 22;

/** Maximum size of the optional archive comment that may trail the EOCD record, in bytes. */
const ZIP_MAX_COMMENT_BYTES = 65535;

/**
 * How far back from the end of the input the EOCD record may start. The record is last in
 * the file apart from its own variable-length comment, so searching this window is
 * sufficient and bounded regardless of archive size.
 */
const EOCD_SEARCH_WINDOW_BYTES = EOCD_RECORD_MIN_BYTES + ZIP_MAX_COMMENT_BYTES;

/**
 * Represents a file extracted from a ZIP archive.
 * Contains the file's path within the archive and its content as a Buffer.
 */
export interface ZipFileContent {
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
 * @param limits - Decompression limits guarding against zip bombs
 * @param config - Parser configuration, so extraction failures honour `onWarning` /
 *                 `outputErrorToConsole` like every other reported issue
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
    limits: DecompressionLimits,
    config?: OfficeParserConfig
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
                fail(getOfficeError(OfficeErrorType.ZIP_ENTRY_COUNT_LIMIT_EXCEEDED, config, maxZipEntries));
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
                        fail(getOfficeError(OfficeErrorType.ZIP_SIZE_LIMIT_EXCEEDED, config, maxUncompressedBytes));
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
        // fflate's streaming Unzip emits nothing, and no error, when the input is not a
        // ZIP archive, unlike the central-directory-based unzip() this replaced (which
        // rejected with "invalid zip data"). Zero entries can never be a valid document
        // here, since every ZIP-backed format requires at least one part, so treat it as
        // corrupt input rather than resolving into an empty, successfully-parsed document.
        //
        // This is checked before the truncation check below because it produces the better
        // message for input that is not an archive at all, and because garbage that happens
        // to contain the EOCD signature would otherwise slip past.
        if (totalEntryCount === 0) {
            fail(getOfficeError(OfficeErrorType.ZIP_NO_ENTRIES_FOUND, config));
            return;
        }
        // The streaming reader recovers entries from local file headers alone, so an archive
        // cut short still yields whatever entries preceded the cut - silently, and possibly
        // missing parts that came after it. The central-directory-based reader used before
        // 7.3.0 rejected such input outright. Requiring the trailer that terminates every
        // complete archive restores that: absent it, the data is truncated and the entries
        // recovered cannot be trusted to be the whole document.
        //
        // Deliberately not gated on pendingFiles: when a cut lands inside an entry's
        // compressed data that entry's final callback never fires, so this is also what
        // settles the promise instead of leaving the caller waiting forever.
        const tail = zipInput.subarray(Math.max(0, zipInput.length - EOCD_SEARCH_WINDOW_BYTES));
        const eocdIndex = tail.lastIndexOf(EOCD_SIGNATURE);
        if (eocdIndex === -1 || eocdIndex + EOCD_RECORD_MIN_BYTES > tail.length) {
            fail(getOfficeError(OfficeErrorType.ZIP_TRUNCATED, config));
            return;
        }
        maybeResolve();
    });
};

/**
 * Finds the archive part that every document of a given format must contain, and fails loudly
 * when it is absent.
 *
 * A readable ZIP archive is not by itself a document: an archive can decompress perfectly and
 * still be a renamed photo bundle, a partial upload, or a file mislabeled with the wrong
 * extension. Without this check a parser finds no content to walk and returns an empty AST,
 * which a caller cannot distinguish from a document that genuinely has nothing in it. Every
 * ZIP-backed format has one part it cannot be valid without, so its absence is a hard error.
 *
 * Pass the parser's own regex/predicate for the part rather than a fresh copy of the path, so
 * this check and the code that later reads the part cannot drift apart.
 *
 * @param files - The entries extracted from the archive
 * @param matcher - Predicate identifying the required part by its path within the archive
 * @param config - Parser configuration, so the error is reported through the caller's handlers
 * @param info - The document format and the human-readable part name, used in the message
 * @returns The matching entry
 * @throws {Error} A typed REQUIRED_PART_MISSING error when no entry matches
 *
 * @example
 * ```typescript
 * const document = findRequiredPart(files, p => !!p.match(documentFileRegex), config,
 *     { fileType: 'docx', part: 'word/document.xml' });
 * ```
 */
export const findRequiredPart = (
    files: ZipFileContent[],
    matcher: (path: string) => boolean,
    config: OfficeParserConfig,
    info: { fileType: string, part: string }
): ZipFileContent => {
    const found = files.find(file => matcher(file.path));
    if (!found) throw getOfficeError(OfficeErrorType.REQUIRED_PART_MISSING, config, info);
    return found;
};


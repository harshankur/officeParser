/**
 * Image Processing Utilities
 * 
 * Provides helper functions for working with image attachments extracted from office documents.
 * Handles MIME type conversions, file extension mapping, and attachment object creation.
 * 
 * @module imageUtils
 */

import { OfficeAttachment, OfficeContentNode, OfficeMimeType } from '../types';

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
        // Windows metafiles, by the names RFC 7903 gives them
        case 'emf': return 'image/emf';
        case 'wmf': return 'image/wmf';
        case 'pdf': return 'application/pdf';
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

    // PDF: 25 50 44 46 ("%PDF"), as a LaTeX project's figures often are
    if (buffer[0] === 0x25 && buffer[1] === 0x50 && buffer[2] === 0x44 && buffer[3] === 0x46) {
        return 'application/pdf';
    }

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

    // EMF: an EMR_HEADER record (type 1) whose signature at byte 40 is " EMF"
    if (buffer.length >= 44 && buffer[0] === 1 && buffer[1] === 0 && buffer[2] === 0 && buffer[3] === 0 &&
        buffer[40] === 0x20 && buffer[41] === 0x45 && buffer[42] === 0x4D && buffer[43] === 0x46) {
        return 'image/emf';
    }

    // WMF: the placeable header's key D7 CD C6 9A, or a METAHEADER with no placeable header (type 1 or 2,
    // a header of 9 words, version 0x0100 or 0x0300). Six bytes, since this also types arbitrary
    // embedded files: four bytes alone are too common to mean a metafile.
    if (buffer[0] === 0xD7 && buffer[1] === 0xCD && buffer[2] === 0xC6 && buffer[3] === 0x9A) {
        return 'image/wmf';
    }
    if (buffer.length >= 6 && (buffer[0] === 1 || buffer[0] === 2) && buffer[1] === 0 && buffer[2] === 9 && buffer[3] === 0 &&
        buffer[4] === 0 && (buffer[5] === 1 || buffer[5] === 3)) {
        return 'image/wmf';
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

/** Precomputed CRC-32 table (polynomial 0xEDB88320) for PNG chunk checksums. */
const CRC32_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
        let c = n;
        for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
        table[n] = c >>> 0;
    }
    return table;
})();

/** CRC-32 over a byte range, as PNG requires over each chunk's type+data. */
export function crc32(bytes: Uint8Array): number {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = CRC32_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
}

/**
 * Points each node in `nodes` (their children, notes and comments too) that names a renamed attachment
 * (its `attachmentName`, or an equation's `fallbackAttachmentName`) at its new name: content read from one part (an EPUB chapter, a DOCX chunk) into a document whose
 * attachments already hold some of its names. A node shared along several paths is renamed once.
 */
export const renameAttachments = (nodes: OfficeContentNode[], renamed: ReadonlyMap<string, string>): void => {
    if (renamed.size === 0) return;
    const seen = new Set<OfficeContentNode>();
    const pending: OfficeContentNode[] = [];
    for (let i = nodes.length - 1; i >= 0; i--) pending.push(nodes[i]);
    while (pending.length) {
        const node = pending.pop()!;
        if (!node || typeof node !== 'object' || seen.has(node)) continue;
        seen.add(node);
        const meta = node.metadata as { attachmentName?: unknown; fallbackAttachmentName?: unknown } | undefined;
        const name = meta && typeof meta.attachmentName === 'string' ? renamed.get(meta.attachmentName) : undefined;
        // (An equation's picture of itself, CodeMetadata.fallbackAttachmentName, too.)
        const fallback = meta && typeof meta.fallbackAttachmentName === 'string' ? renamed.get(meta.fallbackAttachmentName) : undefined;
        if (name !== undefined || fallback !== undefined) (node as { metadata?: unknown }).metadata = { ...meta, ...(name !== undefined ? { attachmentName: name } : {}), ...(fallback !== undefined ? { fallbackAttachmentName: fallback } : {}) };
        for (const list of [node.children, node.notes, node.comments]) {
            if (Array.isArray(list)) for (let i = list.length - 1; i >= 0; i--) pending.push(list[i]);
        }
    }
};

/** A picture format a file's first bytes name (see {@link pictureSignature}). */
export type PictureSignature = 'PNG' | 'JPEG' | 'GIF' | 'BMP' | 'TIFF' | 'WebP' | 'PDF' | 'EMF' | 'WMF';

/**
 * A picture's format by its signature alone (PNG, JPEG, GIF, BMP, TIFF, WebP, PDF, EMF, WMF), or undefined when
 * its first bytes are none of them (an SVG, which is text, has none). A writer that decides how to hold a picture
 * by this, not by its name or media type, never writes one format's bytes labelled as another's. Reads at most
 * the first 44 bytes.
 */
export function pictureSignature(bytes: Uint8Array): PictureSignature | undefined {
    const b = bytes, n = bytes.length;
    if (n < 4) return undefined;
    if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47) return 'PNG';
    if (b[0] === 0xFF && b[1] === 0xD8 && b[2] === 0xFF) return 'JPEG';
    if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38) return 'GIF';
    if ((b[0] === 0x49 && b[1] === 0x49 && b[2] === 0x2A && b[3] === 0x00) || (b[0] === 0x4D && b[1] === 0x4D && b[2] === 0x00 && b[3] === 0x2A)) return 'TIFF';
    if (n >= 12 && b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) return 'WebP';
    if (b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46) return 'PDF';
    // EMF: an EMR_HEADER record (type 1) whose signature at byte 40 is " EMF".
    if (n >= 44 && b[0] === 1 && b[1] === 0 && b[2] === 0 && b[3] === 0 && b[40] === 0x20 && b[41] === 0x45 && b[42] === 0x4D && b[43] === 0x46) return 'EMF';
    // WMF: the placeable header's key, or a METAHEADER (type 1 or 2, 9 words, version 0x0100 or 0x0300).
    if (b[0] === 0xD7 && b[1] === 0xCD && b[2] === 0xC6 && b[3] === 0x9A) return 'WMF';
    if (n >= 6 && (b[0] === 1 || b[0] === 2) && b[1] === 0 && b[2] === 9 && b[3] === 0 && b[4] === 0 && (b[5] === 1 || b[5] === 3)) return 'WMF';
    // BMP: "BM" and a DIB header of a size one of its versions has (the two letters alone are common text).
    if (n >= 18 && b[0] === 0x42 && b[1] === 0x4D) {
        const header = (b[14] | b[15] << 8 | b[16] << 16 | b[17] << 24) >>> 0;
        if (header === 12 || header === 16 || header === 40 || header === 52 || header === 56 || header === 64 || header === 108 || header === 124) return 'BMP';
    }
    return undefined;
}

/**
 * What a picture's bytes are, named for a message: its format by its signature (PNG, JPEG, GIF, BMP, TIFF,
 * WebP, EMF, WMF, SVG, PDF), XML that is no picture, else the format its media type or extension names
 * (`unknown` when none does). Only the first 512 bytes are read.
 */
export function pictureKindOf(bytes: Uint8Array, mimeType?: string, extension?: string): string {
    const ascii = (from: number, to: number) => String.fromCharCode(...bytes.subarray(from, to));
    const signature = pictureSignature(bytes);
    if (signature) return signature;
    // (Read by index alone, so a Uint8Array serves where there is no Buffer. A BMP named by its two letters.)
    const mime = getMimeFromBytes(bytes as Buffer);
    if (mime) return mime === 'application/pdf' ? 'PDF' : mime === 'image/webp' ? 'WebP' : mime.slice(6).toUpperCase();
    // (A byte order mark read byte by byte is three characters.)
    const head = ascii(0, Math.min(bytes.length, 512)).replace(/^\xEF\xBB\xBF/, '').trimStart();
    if (/^(<\?xml[^>]*>\s*)?(<!--[\s\S]*?-->\s*)*(<!DOCTYPE[^>]*>\s*)?<svg[\s>]/i.test(head)) return 'SVG';
    if (head.startsWith('<')) return 'XML';
    // (A name of a format is a short word: a media type or an extension of any other length or characters is
    // none, and named no further into a message.)
    const named = (mimeType?.match(/^image\/(?:x-)?([\w.+-]+)$/i)?.[1] ?? extension ?? '').replace(/^svg\+xml$/i, 'svg');
    const upper = /^[\w.+-]{1,16}$/.test(named) ? named.toUpperCase() : '';
    return upper === 'JPG' ? 'JPEG' : upper === 'TIF' ? 'TIFF' : upper === 'WEBP' ? 'WebP' : upper || 'unknown';
}

/**
 * {@link pictureKindOf} for a picture held as base64 (an attachment's `data`), decoding only the start of
 * it, so a picture shown many times is not decoded whole at each. Never throws: data that is not base64
 * is named by its media type or extension alone.
 */
export function pictureKindOfBase64(data: string | undefined, mimeType?: string, extension?: string): string {
    let bytes = new Uint8Array(0);
    try {
        const head = (data ?? '').slice(0, 1024).replace(/[^A-Za-z0-9+/]/g, '');
        const bin = atob(head.slice(0, head.length - head.length % 4));
        bytes = Uint8Array.from(bin, ch => ch.charCodeAt(0));
    } catch { /* named by its media type or extension */ }
    return pictureKindOf(bytes, mimeType, extension);
}

/** Pictures of a kind {@link pictureKindOf} named, for a message: "EMF pictures", or "pictures of its kind" for `unknown`. */
export function picturesOfKind(kind: string): string {
    return kind === 'unknown' ? 'pictures of its kind' : `${kind} pictures`;
}

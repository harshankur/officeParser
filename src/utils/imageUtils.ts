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
 * at its new name: content read from one part (an EPUB chapter, a DOCX chunk) into a document whose
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
        const meta = node.metadata as { attachmentName?: unknown } | undefined;
        const name = meta && typeof meta.attachmentName === 'string' ? renamed.get(meta.attachmentName) : undefined;
        if (name !== undefined) (node as { metadata?: unknown }).metadata = { ...meta, attachmentName: name };
        for (const list of [node.children, node.notes, node.comments]) {
            if (Array.isArray(list)) for (let i = list.length - 1; i >= 0; i--) pending.push(list[i]);
        }
    }
};

/**
 * What a picture's bytes are, named for a message: its format by its signature (PNG, JPEG, GIF, BMP, TIFF,
 * WebP, EMF, WMF, SVG, PDF), XML that is no picture, else the format its media type or extension names
 * (`unknown` when none does). Only the first 512 bytes are read.
 */
export function pictureKindOf(bytes: Uint8Array, mimeType?: string, extension?: string): string {
    const ascii = (from: number, to: number) => String.fromCharCode(...bytes.subarray(from, to));
    const u32 = (at: number) => bytes.length >= at + 4 ? (bytes[at] | bytes[at + 1] << 8 | bytes[at + 2] << 16 | bytes[at + 3] << 24) >>> 0 : -1;
    // (Read by index alone, so a Uint8Array serves where there is no Buffer.)
    const mime = getMimeFromBytes(bytes as Buffer);
    if (mime) return mime === 'application/pdf' ? 'PDF' : mime === 'image/webp' ? 'WebP' : mime.slice(6).toUpperCase();
    if (u32(0) === 1 && ascii(40, 44) === ' EMF') return 'EMF';
    if (u32(0) === 0x9AC6CDD7 || ((bytes[0] === 1 || bytes[0] === 2) && bytes[1] === 0 && bytes[2] === 9 && bytes[3] === 0)) return 'WMF';
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

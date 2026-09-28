/**
 * MHT (MIME HTML, RFC 2557): a web page and the pictures it shows in one MIME `multipart/related`
 * message. A DOCX can hold one as an alternative-format chunk (`w:altChunk`); html-docx-js, for one,
 * writes a document's whole body that way.
 *
 * Read in one pass over the message: each boundary is found once, from where the last one ended.
 */

/** One part of an MHT message, its body decoded. */
export interface MhtPart {
    /** The part's media type, lower case (`text/html`, `image/png`). */
    contentType: string;
    /** Its `charset` parameter, if any. */
    charset?: string;
    /** Its `Content-Location`: the URL the page refers to it by. */
    location?: string;
    body: Buffer;
}

/** How many lines `text` has. */
const lineCount = (text: string): number => {
    let count = 1;
    for (let i = text.indexOf('\n'); i !== -1; i = text.indexOf('\n', i + 1)) count++;
    return count;
};

/**
 * A header block's fields by lower-case name, continuation lines unfolded. Each line is charged to
 * `charge` before any is split out; a field's continuation lines are joined once, at the end.
 */
const readHeaders = (block: string, charge: (count: number) => void): Map<string, string> => {
    charge(lineCount(block));
    const folded = new Map<string, string[]>();
    let name: string | undefined;
    for (const line of block.split(/\r?\n/)) {
        if ((line.startsWith(' ') || line.startsWith('\t')) && name) {
            folded.get(name)!.push(line.trim());
            continue;
        }
        const colon = line.indexOf(':');
        if (colon <= 0) { name = undefined; continue; }
        name = line.slice(0, colon).trim().toLowerCase();
        folded.set(name, [line.slice(colon + 1).trim()]);
    }
    const headers = new Map<string, string>();
    for (const [key, values] of folded) headers.set(key, values.join(' '));
    return headers;
};

/** A parameter of a header value (`boundary="x"` of `multipart/related; boundary="x"`). */
const parameter = (value: string | undefined, key: string): string | undefined => {
    if (!value) return undefined;
    for (const piece of value.split(';').slice(1)) {
        const eq = piece.indexOf('=');
        if (eq === -1 || piece.slice(0, eq).trim().toLowerCase() !== key) continue;
        const raw = piece.slice(eq + 1).trim();
        return raw.length >= 2 && raw.startsWith('"') && raw.endsWith('"') ? raw.slice(1, -1) : raw;
    }
    return undefined;
};

/** Quoted-printable text (RFC 2045) as the bytes it encodes. */
const decodeQuotedPrintable = (text: string): Buffer => {
    const bytes = Buffer.allocUnsafe(text.length);
    let length = 0;
    for (let i = 0; i < text.length; i++) {
        const c = text.charCodeAt(i);
        if (c !== 61 /* = */) { bytes[length++] = c & 0xff; continue; }
        // A soft line break (`=` at a line's end) joins the lines.
        if (text[i + 1] === '\r' && text[i + 2] === '\n') { i += 2; continue; }
        if (text[i + 1] === '\n') { i += 1; continue; }
        const hex = text.slice(i + 1, i + 3);
        if (/^[0-9A-Fa-f]{2}$/.test(hex)) { bytes[length++] = parseInt(hex, 16); i += 2; continue; }
        bytes[length++] = c;
    }
    return bytes.subarray(0, length);
};

/** A part: its headers, then a blank line, then its body. */
const readPart = (text: string, charge: (count: number) => void): MhtPart => {
    const blank = /\r?\n\r?\n/.exec(text);
    const headerBlock = blank ? text.slice(0, blank.index) : '';
    const rawBody = blank ? text.slice(blank.index + blank[0].length) : text;
    const headers = readHeaders(headerBlock, charge);
    const typeHeader = headers.get('content-type');
    const encoding = (headers.get('content-transfer-encoding') || '').trim().toLowerCase();
    const body = encoding === 'base64' ? Buffer.from(rawBody.replace(/[^A-Za-z0-9+/=]/g, ''), 'base64')
        : encoding === 'quoted-printable' ? decodeQuotedPrintable(rawBody)
            : Buffer.from(rawBody, 'latin1');
    return {
        contentType: (typeHeader?.split(';')[0] || 'text/plain').trim().toLowerCase(),
        charset: parameter(typeHeader, 'charset'),
        location: headers.get('content-location')?.trim() || undefined,
        body,
    };
};

/**
 * The parts of an MHT message, in order. A message that is not multipart is one part: its own headers
 * and body. The bytes are read as Latin-1, so each byte is one character and a body keeps its bytes.
 * `charge` is given each part and header line before it is read (a node budget, which throws past it):
 * 63 KB of DOCX held an MHT chunk of 16 million empty parts.
 */
export const readMht = (buffer: Buffer, charge: (count: number) => void = () => {}): MhtPart[] => {
    const text = buffer.toString('latin1');
    const blank = /\r?\n\r?\n/.exec(text);
    const top = readHeaders(blank ? text.slice(0, blank.index) : '', charge);
    const boundary = parameter(top.get('content-type'), 'boundary');
    if (!boundary) return [readPart(text, charge)];
    const delimiter = '--' + boundary;
    const parts: MhtPart[] = [];
    let at = text.indexOf(delimiter, blank ? blank.index : 0);
    while (at !== -1) {
        const start = at + delimiter.length;
        // The closing delimiter (`--boundary--`) ends the message.
        if (text.startsWith('--', start)) break;
        const next = text.indexOf(delimiter, start);
        const end = next === -1 ? text.length : next;
        // The line break before a delimiter belongs to the delimiter.
        const part = text.slice(start, end).replace(/^[ \t]*\r?\n/, '').replace(/\r?\n$/, '');
        charge(1);
        parts.push(readPart(part, charge));
        at = next;
    }
    return parts;
};

/** A part's text in its charset: UTF-8 unless it names another this runtime can decode. */
export const mhtPartText = (part: MhtPart): string => {
    try {
        return new TextDecoder(part.charset || 'utf-8').decode(part.body);
    } catch {
        return new TextDecoder('utf-8').decode(part.body);
    }
};

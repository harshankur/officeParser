/**
 * Decoders for the legacy encodings a document can be in (an HTML page's declared one, an RTF code
 * page, a `.tex` file's `inputenc`), by their WHATWG label.
 */

/** What a reader uses of a `TextDecoder`: the encoding it decodes, and the decoding of bytes in it. */
export interface ByteDecoder {
    readonly encoding: string;
    decode(bytes: Uint8Array): string;
}

/**
 * The characters of Windows-1252's bytes 0x80 to 0x9F, where it differs from Latin-1: punctuation and
 * letters (0x92 is a right single quote, 0x80 the euro sign). The five it leaves undefined (0x81, 0x8D,
 * 0x8F, 0x90, 0x9D) are the control characters of the same number, as the Encoding Standard has them.
 */
const WINDOWS_1252_HIGH = '\u20AC\u0081\u201A\u0192\u201E\u2026\u2020\u2021\u02C6\u2030\u0160\u2039\u0152\u008D\u017D\u008F'
    + '\u0090\u2018\u2019\u201C\u201D\u2022\u2013\u2014\u02DC\u2122\u0161\u203A\u0153\u009D\u017E\u0178';

/** The bytes decoded at a time: what is held beside the text is twice this, however long the document. */
const WINDOWS_1252_SLICE = 1 << 20;

/**
 * Bytes as Windows-1252 text, in time linear in their number. Each byte's character is written as a
 * UTF-16 code unit, low byte first, and the units are read by the runtime's UTF-16 decoder, which is
 * many times faster than building the string a character at a time.
 */
export function decodeWindows1252(bytes: Uint8Array): string {
    const utf16 = new TextDecoder('utf-16le');
    const units = new Uint8Array(2 * Math.min(bytes.length, WINDOWS_1252_SLICE));
    let text = '';
    for (let start = 0; start < bytes.length; start += WINDOWS_1252_SLICE) {
        const end = Math.min(bytes.length, start + WINDOWS_1252_SLICE);
        for (let i = start, at = 0; i < end; i++, at += 2) {
            const byte = bytes[i];
            const code = byte >= 0x80 && byte <= 0x9F ? WINDOWS_1252_HIGH.charCodeAt(byte - 0x80) : byte;
            units[at] = code & 0xFF;
            units[at + 1] = code >> 8;
        }
        text += utf16.decode(units.subarray(0, 2 * (end - start)));
    }
    return text;
}

/**
 * A decoder for the encoding `label` names (a WHATWG label, as `TextDecoder` takes). It throws, as
 * `new TextDecoder(label)` does, for a label this runtime does not decode.
 *
 * Windows-1252 (which `latin1`, `iso-8859-1` and `ascii` are labels of too) is decoded here and not by
 * the runtime: the `TextDecoder` of Node 22.13.0 to 22.22.0 reads its bytes 0x80 to 0x9F as Latin-1
 * (nodejs/node#56542, corrected in 22.22.1), so the curly quotes, dashes and euro signs of a
 * Windows-1252 page, an RTF file or a `.tex` file came out as control characters there.
 */
export function textDecoder(label: string): ByteDecoder {
    const native = new TextDecoder(label);
    return native.encoding === 'windows-1252' ? { encoding: 'windows-1252', decode: decodeWindows1252 } : native;
}

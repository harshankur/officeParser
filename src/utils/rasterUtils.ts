/**
 * Pictures in the formats a PDF cannot hold, turned into PNG for the native PDF engine (pdf-lib embeds
 * PNG and JPEG only): a GIF's first frame and an uncompressed BMP. Each is decoded within a pixel budget
 * the caller gives, checked against the size the file declares before anything is allocated.
 */
import { zlibSync } from 'fflate';
import { crc32 } from './imageUtils.js';

/** A picture as RGBA pixels, row by row, four bytes each. */
export interface Raster { width: number; height: number; rgba: Uint8Array }

/** The picture as a PNG (8-bit RGBA, no filtering), or null when it has no pixels. */
export function encodePng(raster: Raster): Uint8Array | null {
    const { width, height, rgba } = raster;
    if (!(width > 0 && height > 0) || rgba.length < width * height * 4) return null;
    const stride = width * 4;
    const filtered = new Uint8Array((stride + 1) * height);
    for (let y = 0; y < height; y++) filtered.set(rgba.subarray(y * stride, (y + 1) * stride), y * (stride + 1) + 1);
    const chunk = (type: string, data: Uint8Array): Uint8Array => {
        const out = new Uint8Array(12 + data.length);
        const view = new DataView(out.buffer);
        view.setUint32(0, data.length);
        for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
        out.set(data, 8);
        view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
        return out;
    };
    const header = new Uint8Array(13);
    const hv = new DataView(header.buffer);
    hv.setUint32(0, width);
    hv.setUint32(4, height);
    header[8] = 8; // bit depth
    header[9] = 6; // RGBA
    const parts = [new Uint8Array([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]), chunk('IHDR', header), chunk('IDAT', zlibSync(filtered)), chunk('IEND', new Uint8Array(0))];
    const png = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
    let at = 0;
    for (const p of parts) { png.set(p, at); at += p.length; }
    return png;
}

/**
 * A GIF's first frame, drawn on its logical screen (transparent where the frame does not cover it or its
 * transparent colour is), or null when the file cannot be read or its screen holds more than `maxPixels`.
 */
export function decodeGif(bytes: Uint8Array, maxPixels: number): Raster | null {
    if (bytes.length < 13 || bytes[0] !== 0x47 || bytes[1] !== 0x49 || bytes[2] !== 0x46) return null;
    const u16 = (at: number) => bytes[at] | (bytes[at + 1] << 8);
    const width = u16(6), height = u16(8);
    if (!(width > 0 && height > 0) || width * height > maxPixels) return null;
    let at = 13;
    const readTable = (flags: number): Uint8Array | null => {
        if (!(flags & 0x80)) return null;
        const size = 3 * (1 << ((flags & 0x07) + 1));
        if (at + size > bytes.length) return null;
        const table = bytes.subarray(at, at + size);
        at += size;
        return table;
    };
    const globalTable = readTable(bytes[10]);
    let transparent = -1;
    while (at < bytes.length) {
        const block = bytes[at++];
        if (block === 0x3B) return null; // the trailer: no image
        if (block === 0x21) {
            const label = bytes[at++];
            if (label === 0xF9 && bytes[at] >= 4) {
                // Graphic control: its packed flags say whether a colour index is transparent.
                if (bytes[at + 1] & 0x01) transparent = bytes[at + 4];
            }
            // Skip the extension's sub-blocks.
            while (at < bytes.length && bytes[at] !== 0) at += bytes[at] + 1;
            at++;
            continue;
        }
        if (block !== 0x2C || at + 9 > bytes.length) return null;
        const left = u16(at), top = u16(at + 2), frameW = u16(at + 4), frameH = u16(at + 6), flags = bytes[at + 8];
        at += 9;
        const table = readTable(flags) ?? globalTable;
        if (!table || !(frameW > 0 && frameH > 0) || frameW * frameH > maxPixels) return null;
        const minCodeSize = bytes[at++];
        if (!(minCodeSize >= 2 && minCodeSize <= 11)) return null;
        // The image data's sub-blocks, joined.
        const chunks: Uint8Array[] = [];
        let total = 0;
        while (at < bytes.length && bytes[at] !== 0) {
            const n = bytes[at];
            chunks.push(bytes.subarray(at + 1, at + 1 + n));
            total += n;
            at += n + 1;
        }
        const data = new Uint8Array(total);
        let off = 0;
        for (const c of chunks) { data.set(c, off); off += c.length; }
        const indices = lzwDecode(data, minCodeSize, frameW * frameH);
        const rgba = new Uint8Array(width * height * 4);
        // An interlaced frame's rows come in four passes.
        const rowOrder: number[] = [];
        if (flags & 0x40) {
            for (const [start, step] of [[0, 8], [4, 8], [2, 4], [1, 2]]) for (let y = start; y < frameH; y += step) rowOrder.push(y);
        } else for (let y = 0; y < frameH; y++) rowOrder.push(y);
        for (let r = 0; r < frameH; r++) {
            const y = top + rowOrder[r];
            if (y >= height) continue;
            for (let x = 0; x < frameW; x++) {
                const px = left + x;
                if (px >= width) continue;
                const index = indices[r * frameW + x];
                if (index === transparent || index * 3 + 2 >= table.length) continue;
                const o = (y * width + px) * 4;
                rgba[o] = table[index * 3];
                rgba[o + 1] = table[index * 3 + 1];
                rgba[o + 2] = table[index * 3 + 2];
                rgba[o + 3] = 255;
            }
        }
        return { width, height, rgba };
    }
    return null;
}

/** A GIF frame's colour indices from its LZW-compressed data, at most `count` of them (the rest 0). */
function lzwDecode(data: Uint8Array, minCodeSize: number, count: number): Uint8Array {
    const out = new Uint8Array(count);
    const clear = 1 << minCodeSize, end = clear + 1;
    const prefix = new Int16Array(4096), suffix = new Uint8Array(4096), firstOf = new Uint8Array(4096);
    for (let i = 0; i < clear; i++) { prefix[i] = -1; suffix[i] = i; firstOf[i] = i; }
    const stack = new Uint8Array(4096);
    let codeSize = minCodeSize + 1, next = end + 1, previous = -1;
    let bits = 0, buffer = 0, written = 0;
    for (let i = 0; i < data.length && written < count; i++) {
        buffer |= data[i] << bits;
        bits += 8;
        while (bits >= codeSize && written < count) {
            const code = buffer & ((1 << codeSize) - 1);
            buffer >>>= codeSize;
            bits -= codeSize;
            if (code === clear) { codeSize = minCodeSize + 1; next = end + 1; previous = -1; continue; }
            if (code === end) return out;
            let current = code;
            let depth = 0;
            if (code >= next) {
                // The code being defined now: the previous string and its own first character.
                if (previous < 0) return out;
                stack[depth++] = firstOf[previous];
                current = previous;
            }
            while (current >= clear && depth < 4096) { stack[depth++] = suffix[current]; current = prefix[current]; }
            if (current < 0 || depth >= 4096) return out;
            stack[depth++] = current;
            const first = current;
            while (depth > 0 && written < count) out[written++] = stack[--depth];
            if (previous >= 0 && next < 4096) {
                prefix[next] = previous;
                suffix[next] = first;
                firstOf[next] = firstOf[previous];
                next++;
                if (next === 1 << codeSize && codeSize < 12) codeSize++;
            }
            previous = code;
        }
    }
    return out;
}

/**
 * An uncompressed BMP (1, 4, 8, 24 or 32 bits a pixel; 32-bit with or without bit fields), or null when
 * it is another kind or holds more than `maxPixels`. Rows are stored bottom up unless the height is
 * negative.
 */
export function decodeBmp(bytes: Uint8Array, maxPixels: number): Raster | null {
    if (bytes.length < 54 || bytes[0] !== 0x42 || bytes[1] !== 0x4D) return null;
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const dataAt = view.getUint32(10, true), headerSize = view.getUint32(14, true);
    if (headerSize < 40) return null;
    const width = view.getInt32(18, true), rawHeight = view.getInt32(22, true);
    const bpp = view.getUint16(28, true), compression = view.getUint32(30, true);
    const height = Math.abs(rawHeight), topDown = rawHeight < 0;
    if (!(width > 0 && height > 0) || width * height > maxPixels) return null;
    if (!(compression === 0 || (compression === 3 && bpp === 32))) return null;
    if (![1, 4, 8, 24, 32].includes(bpp)) return null;
    const stride = Math.ceil((width * bpp) / 32) * 4;
    if (dataAt + stride * height > bytes.length) return null;
    let palette: Uint8Array | null = null;
    if (bpp <= 8) {
        const colors = view.getUint32(46, true) || (1 << bpp);
        const paletteAt = 14 + headerSize;
        if (paletteAt + colors * 4 > dataAt) return null;
        palette = bytes.subarray(paletteAt, paletteAt + colors * 4);
    }
    const rgba = new Uint8Array(width * height * 4);
    for (let y = 0; y < height; y++) {
        const row = dataAt + (topDown ? y : height - 1 - y) * stride;
        for (let x = 0; x < width; x++) {
            const o = (y * width + x) * 4;
            if (bpp === 24 || bpp === 32) {
                const p = row + x * (bpp / 8);
                rgba[o] = bytes[p + 2]; rgba[o + 1] = bytes[p + 1]; rgba[o + 2] = bytes[p]; rgba[o + 3] = 255;
            } else {
                const bit = x * bpp;
                const index = (bytes[row + (bit >> 3)] >> (8 - bpp - (bit & 7))) & ((1 << bpp) - 1);
                const p = index * 4;
                if (!palette || p + 2 >= palette.length) continue;
                rgba[o] = palette[p + 2]; rgba[o + 1] = palette[p + 1]; rgba[o + 2] = palette[p]; rgba[o + 3] = 255;
            }
        }
    }
    return { width, height, rgba };
}

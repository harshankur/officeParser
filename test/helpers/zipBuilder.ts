/**
 * A minimal ZIP writer for tests, able to produce layouts the usual libraries will not.
 *
 * `fflate`'s `zipSync` always writes entry sizes into the local file header, which means it
 * cannot express an archive written by a streaming producer: those set general-purpose flag
 * bit 3 and defer the sizes to a data descriptor that trails the entry data. That layout is
 * what broke magic-byte type detection in issue #82, so reproducing the bug at all requires
 * writing the archive by hand.
 *
 * Entries are stored uncompressed. Tests here care about archive structure, not compression.
 */

/** Local file header signature ("PK\x03\x04"). */
const LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50;

/** Central directory file header signature ("PK\x01\x02"). */
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;

/** End of central directory signature ("PK\x05\x06"). */
const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50;

/** Data descriptor signature ("PK\x07\x08"). */
const DATA_DESCRIPTOR_SIGNATURE = 0x08074b50;

/** Byte lengths of the fixed portions of each record. */
const LOCAL_HEADER_BYTES = 30;
const CENTRAL_HEADER_BYTES = 46;
const END_OF_CENTRAL_DIRECTORY_BYTES = 22;
const DATA_DESCRIPTOR_BYTES = 16;

/** Marks the filename as UTF-8. */
const FLAG_UTF8_FILENAME = 0x800;

/** Marks sizes and CRC as deferred to a trailing data descriptor. */
const FLAG_DATA_DESCRIPTOR = 0x008;

/** Stored, i.e. no compression. */
const COMPRESSION_STORED = 0;

/** ZIP version needed to extract, 2.0. */
const VERSION_NEEDED = 20;

const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
        let c = n;
        for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
        table[n] = c >>> 0;
    }
    return table;
})();

/** CRC-32 as ZIP defines it. */
export const crc32 = (bytes: Uint8Array): number => {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
};

/** One entry to write into the archive. */
export interface ZipBuilderEntry {
    /** Path within the archive. */
    name: string;
    /** Raw contents, stored uncompressed. */
    data: Uint8Array;
    /**
     * Write this entry the way a streaming producer would: general-purpose flag bit 3 set,
     * zeroed sizes in the local header, and a data descriptor after the data.
     */
    dataDescriptor?: boolean;
}

/**
 * Builds a ZIP archive from entries, in the order given.
 *
 * Order matters for these tests: type detection looks for the part that names the format, and
 * the bug under test is about how much data sits in front of it.
 *
 * @param entries - Entries to write, in archive order
 * @returns The complete archive
 *
 * @example
 * ```typescript
 * // A presentation whose declaration sits behind a large streamed entry
 * buildZip([
 *     { name: 'ppt/theme/theme1.xml', data: big, dataDescriptor: true },
 *     { name: '[Content_Types].xml', data: contentTypes, dataDescriptor: true },
 * ]);
 * ```
 */
export function buildZip(entries: ZipBuilderEntry[]): Buffer {
    const encoder = new TextEncoder();
    const pieces: Buffer[] = [];
    const centralRecords: Buffer[] = [];
    let offset = 0;

    for (const entry of entries) {
        const name = encoder.encode(entry.name);
        const checksum = crc32(entry.data);
        const streamed = !!entry.dataDescriptor;
        const flags = FLAG_UTF8_FILENAME | (streamed ? FLAG_DATA_DESCRIPTOR : 0);

        const local = Buffer.alloc(LOCAL_HEADER_BYTES + name.length);
        local.writeUInt32LE(LOCAL_FILE_HEADER_SIGNATURE, 0);
        local.writeUInt16LE(VERSION_NEEDED, 4);
        local.writeUInt16LE(flags, 6);
        local.writeUInt16LE(COMPRESSION_STORED, 8);
        local.writeUInt16LE(0, 10);      // modification time
        local.writeUInt16LE(0x21, 12);   // modification date, a valid 1980-01-01
        // A streamed entry cannot know these yet, which is the whole point of bit 3.
        local.writeUInt32LE(streamed ? 0 : checksum, 14);
        local.writeUInt32LE(streamed ? 0 : entry.data.length, 18);
        local.writeUInt32LE(streamed ? 0 : entry.data.length, 22);
        local.writeUInt16LE(name.length, 26);
        local.writeUInt16LE(0, 28);      // extra field length
        Buffer.from(name).copy(local, LOCAL_HEADER_BYTES);

        const localOffset = offset;
        pieces.push(local, Buffer.from(entry.data));
        offset += local.length + entry.data.length;

        if (streamed) {
            const descriptor = Buffer.alloc(DATA_DESCRIPTOR_BYTES);
            descriptor.writeUInt32LE(DATA_DESCRIPTOR_SIGNATURE, 0);
            descriptor.writeUInt32LE(checksum, 4);
            descriptor.writeUInt32LE(entry.data.length, 8);
            descriptor.writeUInt32LE(entry.data.length, 12);
            pieces.push(descriptor);
            offset += descriptor.length;
        }

        // The central directory always carries real values, even for a streamed entry.
        const central = Buffer.alloc(CENTRAL_HEADER_BYTES + name.length);
        central.writeUInt32LE(CENTRAL_DIRECTORY_SIGNATURE, 0);
        central.writeUInt16LE(VERSION_NEEDED, 4);
        central.writeUInt16LE(VERSION_NEEDED, 6);
        central.writeUInt16LE(flags, 8);
        central.writeUInt16LE(COMPRESSION_STORED, 10);
        central.writeUInt16LE(0, 12);
        central.writeUInt16LE(0x21, 14);
        central.writeUInt32LE(checksum, 16);
        central.writeUInt32LE(entry.data.length, 20);
        central.writeUInt32LE(entry.data.length, 24);
        central.writeUInt16LE(name.length, 28);
        central.writeUInt16LE(0, 30);    // extra field length
        central.writeUInt16LE(0, 32);    // comment length
        central.writeUInt16LE(0, 34);    // disk number start
        central.writeUInt16LE(0, 36);    // internal attributes
        central.writeUInt32LE(0, 38);    // external attributes
        central.writeUInt32LE(localOffset, 42);
        Buffer.from(name).copy(central, CENTRAL_HEADER_BYTES);
        centralRecords.push(central);
    }

    const centralDirectory = Buffer.concat(centralRecords);
    const end = Buffer.alloc(END_OF_CENTRAL_DIRECTORY_BYTES);
    end.writeUInt32LE(END_OF_CENTRAL_DIRECTORY_SIGNATURE, 0);
    end.writeUInt16LE(0, 4);             // this disk
    end.writeUInt16LE(0, 6);             // disk with central directory
    end.writeUInt16LE(entries.length, 8);
    end.writeUInt16LE(entries.length, 10);
    end.writeUInt32LE(centralDirectory.length, 12);
    end.writeUInt32LE(offset, 16);
    end.writeUInt16LE(0, 20);            // comment length

    return Buffer.concat([...pieces, centralDirectory, end]);
}

/**
 * Deterministic filler that does not compress away, for padding an archive to a target size.
 *
 * @param size - Number of bytes to produce
 */
export const incompressibleBytes = (size: number): Uint8Array => {
    const bytes = new Uint8Array(size);
    let seed = 7;
    for (let i = 0; i < size; i++) {
        seed = (seed * 1103515245 + 12345) & 0x7fffffff;
        bytes[i] = seed & 0xff;
    }
    return bytes;
};

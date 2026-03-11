/**
 * OLE2/CFB (Compound File Binary) Container Reader
 *
 * Pure TypeScript implementation of the OLE2 (also known as Compound File Binary or
 * Structured Storage) container format used by legacy Microsoft Office files (.doc, .xls, .ppt).
 *
 * Ported from the algorithms in:
 * - ref/olefile/olefile/olefile.py (BSD-2 license, Python OLE2 parser)
 * - Microsoft [MS-CFB] specification
 *
 * The OLE2 format is essentially a filesystem-within-a-file. It contains:
 * - A 512-byte header with FAT sector locations
 * - A FAT (File Allocation Table) mapping sector chains
 * - A directory of named streams (like files in a filesystem)
 * - A mini-FAT for streams smaller than 4096 bytes
 * - The actual stream data stored across linked sectors
 *
 * @module ole2
 */

// ============================================================================
// Constants
// ============================================================================

/** OLE2 magic bytes: D0 CF 11 E0 A1 B1 1A E1 */
const OLE2_MAGIC = Buffer.from([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);

/** Special sector markers */
const ENDOFCHAIN = 0xFFFFFFFE;  // End of a sector chain
const FREESECT   = 0xFFFFFFFF;  // Free/unallocated sector
const FATSECT    = 0xFFFFFFFD;  // Sector used by the FAT itself
const DIFSECT    = 0xFFFFFFFC;  // Sector used by the DIFAT
const NOSTREAM   = 0xFFFFFFFF;  // No child/sibling in directory red-black tree

/** Directory entry types */
const STGTY_EMPTY   = 0;
const STGTY_STORAGE = 1;
const STGTY_STREAM  = 2;
const STGTY_ROOT    = 5;

/** Mini-stream threshold: streams smaller than this use the mini-FAT */
const MINI_STREAM_CUTOFF = 4096;

/** Mini-sector size is always 64 bytes */
const MINI_SECTOR_SIZE = 64;

/** Minimum valid OLE2 file size (header + at least one sector) */
const MIN_FILE_SIZE = 512 + 512;

/** Maximum number of FAT sector indices in the header (at offset 0x4C) */
const HEADER_DIFAT_ENTRIES = 109;

// ============================================================================
// Types
// ============================================================================

/** Represents a directory entry in the OLE2 container */
export interface OLE2DirectoryEntry {
    /** Entry name (decoded from UTF-16LE) */
    name: string;
    /** Entry type: 0=empty, 1=storage, 2=stream, 5=root */
    entryType: number;
    /** Index of left sibling in red-black tree */
    leftSibling: number;
    /** Index of right sibling in red-black tree */
    rightSibling: number;
    /** Index of child node (for storages/root) */
    child: number;
    /** First sector of this stream's data */
    startSector: number;
    /** Stream size in bytes */
    size: number;
    /** Stream ID (index in directory array) */
    sid: number;
}

/** Parsed OLE2 container providing stream access */
export interface OLE2Container {
    /** All directory entries */
    entries: OLE2DirectoryEntry[];
    /** Get a named stream's data. Throws if not found. */
    getStream(name: string): Buffer;
    /** Check if a named stream exists */
    hasStream(name: string): boolean;
    /** List all stream names (type=2) in the container */
    listStreams(): string[];
    /** Get the sector size (512 or 4096) */
    sectorSize: number;
}

// ============================================================================
// Implementation
// ============================================================================

/**
 * Parse an OLE2/CFB container from a buffer.
 *
 * @param data - Buffer containing the OLE2 file
 * @returns An OLE2Container providing access to named streams
 * @throws Error if the file is not a valid OLE2 container
 */
export function parseOLE2(data: Buffer): OLE2Container {
    // ------------------------------------------------------------------
    // 1. Validate minimum size and magic bytes
    // ------------------------------------------------------------------
    if (data.length < MIN_FILE_SIZE) {
        throw new Error('OLE2: File too small to be a valid OLE2 container');
    }

    if (!data.subarray(0, 8).equals(OLE2_MAGIC)) {
        throw new Error('OLE2: Invalid magic bytes - not an OLE2/CFB file');
    }

    // ------------------------------------------------------------------
    // 2. Parse header fields
    // ------------------------------------------------------------------
    const byteOrder = data.readUInt16LE(0x1C);
    if (byteOrder !== 0xFFFE) {
        throw new Error(`OLE2: Unsupported byte order 0x${byteOrder.toString(16)} (expected 0xFFFE little-endian)`);
    }

    const sectorShift = data.readUInt16LE(0x1E);
    if (sectorShift !== 9 && sectorShift !== 12) {
        throw new Error(`OLE2: Unsupported sector shift ${sectorShift} (expected 9 or 12)`);
    }
    const sectorSize = 1 << sectorShift;  // 512 or 4096

    const miniSectorShift = data.readUInt16LE(0x20);
    if (miniSectorShift !== 6) {
        throw new Error(`OLE2: Unsupported mini-sector shift ${miniSectorShift} (expected 6)`);
    }

    const numFatSectors = data.readUInt32LE(0x2C);
    const firstDirSector = data.readUInt32LE(0x30);
    const firstMiniFatSector = data.readUInt32LE(0x3C);
    const numMiniFatSectors = data.readUInt32LE(0x40);
    const firstDifatSector = data.readUInt32LE(0x44);
    const numDifatSectors = data.readUInt32LE(0x48);

    // Calculate total sectors in file (sector 0 starts after the 512-byte header)
    // For 4096-byte sectors, the header occupies the first sector
    const totalSectors = Math.floor((data.length - sectorSize) / sectorSize);

    // ------------------------------------------------------------------
    // 3. Helper: read a sector by index
    // ------------------------------------------------------------------
    function readSector(sectorIndex: number): Buffer {
        // File offset = (sectorIndex + 1) * sectorSize
        // The +1 accounts for the header which occupies the space before sector 0
        const offset = (sectorIndex + 1) * sectorSize;
        if (offset + sectorSize > data.length) {
            throw new Error(`OLE2: Sector ${sectorIndex} is out of bounds (offset ${offset}, file size ${data.length})`);
        }
        return data.subarray(offset, offset + sectorSize);
    }

    // ------------------------------------------------------------------
    // 4. Build the FAT (File Allocation Table)
    // ------------------------------------------------------------------
    // The FAT is an array of 32-bit sector indices. Each entry at index N
    // tells you the next sector in the chain starting from sector N.
    // Special values: ENDOFCHAIN, FREESECT, FATSECT, DIFSECT

    // Step 4a: Collect FAT sector indices from header DIFAT array (up to 109)
    const fatSectorIndices: number[] = [];

    // Read the 109 DIFAT entries in the header (at offset 0x4C)
    for (let i = 0; i < HEADER_DIFAT_ENTRIES && fatSectorIndices.length < numFatSectors; i++) {
        const idx = data.readUInt32LE(0x4C + i * 4);
        if (idx === FREESECT || idx === ENDOFCHAIN) break;
        fatSectorIndices.push(idx);
    }

    // Step 4b: If there are DIFAT sectors, follow the DIFAT chain for more FAT sector indices
    if (numDifatSectors > 0 && firstDifatSector < ENDOFCHAIN) {
        let difatSector = firstDifatSector;
        const entriesPerDifat = (sectorSize / 4) - 1;  // Last entry is link to next DIFAT sector

        for (let d = 0; d < numDifatSectors; d++) {
            if (difatSector >= ENDOFCHAIN) break;
            const difatData = readSector(difatSector);

            // Read FAT sector indices from this DIFAT sector
            for (let i = 0; i < entriesPerDifat && fatSectorIndices.length < numFatSectors; i++) {
                const idx = difatData.readUInt32LE(i * 4);
                if (idx === FREESECT || idx === ENDOFCHAIN) break;
                fatSectorIndices.push(idx);
            }

            // Follow link to next DIFAT sector (last 4 bytes)
            difatSector = difatData.readUInt32LE(entriesPerDifat * 4);
        }
    }

    // Step 4c: Read all FAT sectors and assemble the FAT array
    const fat: number[] = [];
    const entriesPerSector = sectorSize / 4;

    for (const fatSectorIdx of fatSectorIndices) {
        const sectorData = readSector(fatSectorIdx);
        for (let i = 0; i < entriesPerSector; i++) {
            fat.push(sectorData.readUInt32LE(i * 4));
        }
    }

    // Trim FAT to actual number of sectors
    fat.length = Math.min(fat.length, totalSectors);

    // ------------------------------------------------------------------
    // 5. Helper: read a chain of sectors following the FAT
    // ------------------------------------------------------------------
    function readChain(startSector: number, streamSize: number): Buffer {
        if (startSector === ENDOFCHAIN || streamSize === 0) {
            return Buffer.alloc(0);
        }

        const expectedSectors = Math.ceil(streamSize / sectorSize);
        const chunks: Buffer[] = [];
        let currentSector = startSector;
        let readCount = 0;

        while (currentSector < ENDOFCHAIN && readCount < expectedSectors) {
            if (currentSector >= fat.length) {
                throw new Error(`OLE2: Sector ${currentSector} out of FAT range (FAT size: ${fat.length})`);
            }
            chunks.push(readSector(currentSector));
            currentSector = fat[currentSector];
            readCount++;

            // Safety: prevent infinite loops
            if (readCount > fat.length) {
                throw new Error('OLE2: Circular FAT chain detected');
            }
        }

        const result = Buffer.concat(chunks);
        return result.subarray(0, streamSize);
    }

    // ------------------------------------------------------------------
    // 6. Parse directory entries
    // ------------------------------------------------------------------
    // The directory is a stream starting at firstDirSector, read via FAT chain.
    // We don't know the exact size, so read until chain ends.
    const dirChunks: Buffer[] = [];
    let dirSector = firstDirSector;
    let dirReadCount = 0;

    while (dirSector < ENDOFCHAIN && dirReadCount < fat.length + 1) {
        dirChunks.push(readSector(dirSector));
        if (dirSector >= fat.length) break;
        dirSector = fat[dirSector];
        dirReadCount++;
    }

    const dirData = Buffer.concat(dirChunks);
    const numDirEntries = Math.floor(dirData.length / 128);
    const entries: OLE2DirectoryEntry[] = [];

    for (let i = 0; i < numDirEntries; i++) {
        const offset = i * 128;
        const entryType = dirData.readUInt8(offset + 0x42);

        if (entryType === STGTY_EMPTY) {
            // Skip empty entries but keep SID alignment
            entries.push({
                name: '',
                entryType: STGTY_EMPTY,
                leftSibling: NOSTREAM,
                rightSibling: NOSTREAM,
                child: NOSTREAM,
                startSector: 0,
                size: 0,
                sid: i,
            });
            continue;
        }

        // Name: UTF-16LE, up to 64 bytes. nameLen includes null terminator.
        const nameLen = dirData.readUInt16LE(offset + 0x40);
        const nameBytes = nameLen > 2 ? nameLen - 2 : 0;  // Exclude null terminator
        const name = dirData.subarray(offset, offset + nameBytes).toString('utf16le');

        entries.push({
            name,
            entryType,
            leftSibling: dirData.readUInt32LE(offset + 0x44),
            rightSibling: dirData.readUInt32LE(offset + 0x48),
            child: dirData.readUInt32LE(offset + 0x4C),
            startSector: dirData.readUInt32LE(offset + 0x74),
            size: dirData.readUInt32LE(offset + 0x78),
            sid: i,
        });
    }

    // ------------------------------------------------------------------
    // 7. Build mini-FAT and mini-stream (for streams < 4096 bytes)
    // ------------------------------------------------------------------
    let miniFat: number[] = [];
    let miniStream: Buffer = Buffer.alloc(0);

    // Root entry (SID 0) holds the mini-stream container
    const rootEntry = entries[0];

    if (numMiniFatSectors > 0 && firstMiniFatSector < ENDOFCHAIN) {
        // Read mini-FAT sectors via FAT chain
        const miniFatChunks: Buffer[] = [];
        let mfSector = firstMiniFatSector;
        let mfCount = 0;

        while (mfSector < ENDOFCHAIN && mfCount < numMiniFatSectors) {
            miniFatChunks.push(readSector(mfSector));
            if (mfSector >= fat.length) break;
            mfSector = fat[mfSector];
            mfCount++;
        }

        const miniFatData = Buffer.concat(miniFatChunks);
        for (let i = 0; i < miniFatData.length / 4; i++) {
            miniFat.push(miniFatData.readUInt32LE(i * 4));
        }
    }

    // Read the mini-stream container from the root entry's FAT chain
    if (rootEntry && rootEntry.size > 0 && rootEntry.startSector < ENDOFCHAIN) {
        miniStream = readChain(rootEntry.startSector, rootEntry.size);
    }

    // ------------------------------------------------------------------
    // 8. Helper: read a mini-stream chain
    // ------------------------------------------------------------------
    function readMiniChain(startSector: number, streamSize: number): Buffer {
        if (startSector === ENDOFCHAIN || streamSize === 0) {
            return Buffer.alloc(0);
        }

        const expectedMiniSectors = Math.ceil(streamSize / MINI_SECTOR_SIZE);
        const chunks: Buffer[] = [];
        let currentSector = startSector;
        let readCount = 0;

        while (currentSector < ENDOFCHAIN && readCount < expectedMiniSectors) {
            if (currentSector >= miniFat.length) {
                throw new Error(`OLE2: Mini-sector ${currentSector} out of mini-FAT range`);
            }
            const miniOffset = currentSector * MINI_SECTOR_SIZE;
            if (miniOffset + MINI_SECTOR_SIZE > miniStream.length) {
                throw new Error(`OLE2: Mini-sector ${currentSector} exceeds mini-stream size`);
            }
            chunks.push(miniStream.subarray(miniOffset, miniOffset + MINI_SECTOR_SIZE));
            currentSector = miniFat[currentSector];
            readCount++;

            if (readCount > miniFat.length) {
                throw new Error('OLE2: Circular mini-FAT chain detected');
            }
        }

        const result = Buffer.concat(chunks);
        return result.subarray(0, streamSize);
    }

    // ------------------------------------------------------------------
    // 9. Build stream access interface
    // ------------------------------------------------------------------

    /** Find a directory entry by name (case-insensitive) */
    function findEntry(name: string): OLE2DirectoryEntry | undefined {
        const lowerName = name.toLowerCase();
        return entries.find(e =>
            e.entryType !== STGTY_EMPTY &&
            e.name.toLowerCase() === lowerName
        );
    }

    /** Read a stream by its directory entry */
    function readStreamEntry(entry: OLE2DirectoryEntry): Buffer {
        if (entry.size === 0) {
            return Buffer.alloc(0);
        }

        // Streams < MINI_STREAM_CUTOFF use the mini-FAT (except for the root entry)
        if (entry.size < MINI_STREAM_CUTOFF && entry.entryType === STGTY_STREAM) {
            return readMiniChain(entry.startSector, entry.size);
        }

        // Larger streams use the regular FAT
        return readChain(entry.startSector, entry.size);
    }

    return {
        entries,
        sectorSize,

        getStream(name: string): Buffer {
            const entry = findEntry(name);
            if (!entry || entry.entryType !== STGTY_STREAM) {
                throw new Error(`OLE2: Stream "${name}" not found in container`);
            }
            return readStreamEntry(entry);
        },

        hasStream(name: string): boolean {
            const entry = findEntry(name);
            return !!entry && entry.entryType === STGTY_STREAM;
        },

        listStreams(): string[] {
            return entries
                .filter(e => e.entryType === STGTY_STREAM)
                .map(e => e.name);
        },
    };
}

/**
 * Check if a buffer starts with OLE2 magic bytes.
 * Useful for quick format detection before full parsing.
 */
export function isOLE2(data: Buffer | Uint8Array): boolean {
    if (data.length < 8) return false;
    return data[0] === 0xD0 && data[1] === 0xCF && data[2] === 0x11 && data[3] === 0xE0 &&
           data[4] === 0xA1 && data[5] === 0xB1 && data[6] === 0x1A && data[7] === 0xE1;
}

/**
 * Detect the specific Office format inside an OLE2 container by examining
 * its directory entries. Returns 'doc', 'xls', or 'ppt' (or undefined if unknown).
 */
export function detectOLE2Format(container: OLE2Container): 'doc' | 'xls' | 'ppt' | undefined {
    const streams = container.listStreams();
    const streamNames = new Set(streams.map(s => s.toLowerCase()));

    // Excel: contains "Workbook" or "Book" stream
    if (streamNames.has('workbook') || streamNames.has('book')) {
        return 'xls';
    }

    // PowerPoint: contains "PowerPoint Document" stream
    if (streams.some(s => s.toLowerCase() === 'powerpoint document')) {
        return 'ppt';
    }

    // Word: contains "WordDocument" stream
    if (streamNames.has('worddocument')) {
        return 'doc';
    }

    return undefined;
}

/**
 * Excel 97-2003 Binary (.xls) Parser — BIFF8 Format
 *
 * Pure TypeScript implementation that reads .xls files by:
 * 1. Using the OLE2 reader to extract the "Workbook" stream
 * 2. Parsing BIFF8 records (4-byte header: type + length, followed by data)
 * 3. Building a Shared String Table (SST) for text lookups
 * 4. Extracting sheet names, rows, and cell values
 * 5. Returning a valid OfficeParserAST
 *
 * Ported from the algorithms in:
 * - ref/poi/poi/src/main/java/org/apache/poi/hssf/ (Apache 2.0, Java HSSF)
 *   Specifically: HSSFWorkbook.java, RecordFactory.java, SSTRecord.java,
 *   BoundSheetRecord.java, LabelSSTRecord.java, NumberRecord.java
 *
 * @module xls
 */

import {
    CellMetadata, OfficeContentNode, OfficeParserAST,
    OfficeParserConfig, SheetMetadata, TextFormatting
} from '../../types';
import { parseOLE2 } from './ole2';

// ============================================================================
// BIFF8 Record Type Constants
// ============================================================================

const RT_BOF        = 0x0809;
const RT_EOF        = 0x000A;
const RT_BOUNDSHEET = 0x0085;  // Sheet information
const RT_SST        = 0x00FC;  // Shared String Table
const RT_CONTINUE   = 0x003C;  // Continuation record
const RT_LABELSST   = 0x00FD;  // Cell referencing SST
const RT_LABEL      = 0x0204;  // Cell with inline string (legacy)
const RT_NUMBER     = 0x0203;  // Cell with IEEE double
const RT_RK         = 0x027E;  // Cell with compact number
const RT_MULRK      = 0x00BD;  // Multiple RK cells in one record
const RT_MULBLANK   = 0x00BE;  // Multiple blank cells
const RT_FORMULA    = 0x0006;  // Cell with formula
const RT_STRING     = 0x0207;  // String result of a formula
const RT_BLANK      = 0x0201;  // Blank cell with formatting
const RT_BOOLERR    = 0x0205;  // Boolean or error cell
const RT_DIMENSION  = 0x0200;  // Sheet dimensions
const RT_ROW        = 0x0208;  // Row information
const RT_FONT       = 0x0031;  // Font record
const RT_FORMAT     = 0x041E;  // Number format string
const RT_XF         = 0x00E0;  // Extended format (cell style)

/** Maximum BIFF8 record data size (before CONTINUE) */
const MAX_RECORD_DATA = 8224;

// ============================================================================
// Types
// ============================================================================

interface BoundSheet {
    /** Absolute byte offset of the sheet's BOF record within the workbook stream */
    bofOffset: number;
    /** 0=visible, 1=hidden, 2=very hidden */
    visibility: number;
    /** Sheet type: 0=worksheet, 2=chart, 6=VBA module */
    sheetType: number;
    /** Sheet name */
    name: string;
}

interface CellValue {
    row: number;
    col: number;
    value: string;
    xfIndex: number;
}

interface FontRecord {
    bold: boolean;
    italic: boolean;
    underline: boolean;
    strikethrough: boolean;
    size: number;     // In twips (1/20 of a point)
    color: number;    // Color index
    name: string;
}

interface XFRecord {
    fontIndex: number;
    formatIndex: number;
    alignment: number;  // 0=general, 1=left, 2=center, 3=right, 4=fill, 5=justify
}

// ============================================================================
// Helpers
// ============================================================================

/**
 * Read a BIFF8 Unicode string from a buffer.
 *
 * BIFF8 strings have the format:
 * - 2 bytes: character count
 * - 1 byte: option flags (bit 0 = 1 for UTF-16LE, 0 for compressed Latin-1;
 *            bit 2 = rich text; bit 3 = Asian phonetic)
 * - [2 bytes: rich text run count] (if bit 2 set)
 * - [4 bytes: phonetic size] (if bit 3 set)
 * - String data
 * - [Rich text run data] (4 bytes per run)
 * - [Phonetic data]
 *
 * @returns [decoded string, total bytes consumed]
 */
function readBiffString(buf: Buffer, offset: number): [string, number] {
    if (offset + 3 > buf.length) return ['', 3];

    const charCount = buf.readUInt16LE(offset);
    const flags = buf.readUInt8(offset + 2);
    let pos = offset + 3;

    const isUnicode = (flags & 0x01) !== 0;
    const hasRichText = (flags & 0x08) !== 0;
    const hasPhonetic = (flags & 0x04) !== 0;

    let richTextRunCount = 0;
    let phoneticSize = 0;

    if (hasRichText) {
        richTextRunCount = buf.readUInt16LE(pos);
        pos += 2;
    }
    if (hasPhonetic) {
        phoneticSize = buf.readUInt32LE(pos);
        pos += 4;
    }

    let text: string;
    if (isUnicode) {
        const byteLen = charCount * 2;
        if (pos + byteLen > buf.length) {
            text = buf.subarray(pos, buf.length).toString('utf16le');
            pos = buf.length;
        } else {
            text = buf.subarray(pos, pos + byteLen).toString('utf16le');
            pos += byteLen;
        }
    } else {
        // Compressed: Latin-1, one byte per character
        const byteLen = charCount;
        if (pos + byteLen > buf.length) {
            text = buf.subarray(pos, buf.length).toString('latin1');
            pos = buf.length;
        } else {
            text = buf.subarray(pos, pos + byteLen).toString('latin1');
            pos += byteLen;
        }
    }

    // Skip rich text runs (4 bytes each) and phonetic data
    pos += richTextRunCount * 4;
    pos += phoneticSize;

    return [text, pos - offset];
}

/**
 * Read a short BIFF string (1-byte length prefix, used in BOUNDSHEET).
 * @returns [decoded string, bytes consumed]
 */
function readShortBiffString(buf: Buffer, offset: number): [string, number] {
    if (offset + 2 > buf.length) return ['', 2];

    const charCount = buf.readUInt8(offset);
    const flags = buf.readUInt8(offset + 1);
    let pos = offset + 2;

    const isUnicode = (flags & 0x01) !== 0;

    let text: string;
    if (isUnicode) {
        const byteLen = charCount * 2;
        text = buf.subarray(pos, pos + byteLen).toString('utf16le');
        pos += byteLen;
    } else {
        text = buf.subarray(pos, pos + charCount).toString('latin1');
        pos += charCount;
    }

    return [text, pos - offset];
}

/**
 * Decode an RK number to a double.
 * Bits 0: 0=IEEE 64-bit, 1=integer in bits 2-31
 * Bit  1: if set, divide result by 100
 */
function decodeRK(rk: number): number {
    const isInteger = (rk & 0x02) !== 0;
    const div100 = (rk & 0x01) !== 0;

    let value: number;
    if (isInteger) {
        // Signed integer in bits 2-31
        value = rk >> 2;
    } else {
        // IEEE 64-bit double with bits 0-31 being the low 32 bits set to 0
        // and rk (with bits 0-1 cleared) being the high 32 bits
        const buf = Buffer.alloc(8);
        buf.writeUInt32LE(0, 0);          // Low 32 bits = 0
        buf.writeInt32LE(rk & 0xFFFFFFFC, 4);  // High 32 bits = rk with low 2 bits cleared
        value = buf.readDoubleLE(0);
    }

    if (div100) {
        value /= 100;
    }

    return value;
}

/**
 * Format a numeric value as a string, avoiding excessive decimal places.
 */
function formatNumber(value: number): string {
    if (Number.isInteger(value)) return value.toString();
    // Use toPrecision to avoid floating point artifacts, then parseFloat to clean
    const s = parseFloat(value.toPrecision(15)).toString();
    return s;
}

// ============================================================================
// SST Parsing with CONTINUE record support
// ============================================================================

/**
 * Parse the Shared String Table (SST) from a potentially multi-record buffer.
 * The SST record may be followed by CONTINUE records if it exceeds 8224 bytes.
 *
 * SST structure:
 * - 4 bytes: total string references in workbook
 * - 4 bytes: number of unique strings
 * - Followed by unique strings in BIFF8 Unicode format
 */
function parseSST(sstData: Buffer, continueBuffers: Buffer[]): string[] {
    const strings: string[] = [];

    if (sstData.length < 8) return strings;

    const uniqueCount = sstData.readUInt32LE(4);
    let pos = 8;  // Skip total refs (4) + unique count (4)

    // Combine SST data with any CONTINUE records into one logical buffer
    // This simplifies parsing, though not perfectly handling mid-string CONTINUE splits
    const allBuffers = [sstData.subarray(pos)];
    for (const cont of continueBuffers) {
        allBuffers.push(cont);
    }
    const combined = Buffer.concat(allBuffers);
    pos = 0;

    for (let i = 0; i < uniqueCount && pos < combined.length; i++) {
        if (pos + 3 > combined.length) break;

        const charCount = combined.readUInt16LE(pos);
        const flags = combined.readUInt8(pos + 2);
        let strPos = pos + 3;

        const isUnicode = (flags & 0x01) !== 0;
        const hasRichText = (flags & 0x08) !== 0;
        const hasPhonetic = (flags & 0x04) !== 0;

        let richTextRunCount = 0;
        let phoneticSize = 0;

        if (hasRichText) {
            if (strPos + 2 > combined.length) break;
            richTextRunCount = combined.readUInt16LE(strPos);
            strPos += 2;
        }
        if (hasPhonetic) {
            if (strPos + 4 > combined.length) break;
            phoneticSize = combined.readUInt32LE(strPos);
            strPos += 4;
        }

        let text: string;
        if (isUnicode) {
            const byteLen = charCount * 2;
            if (strPos + byteLen > combined.length) {
                text = combined.subarray(strPos, combined.length).toString('utf16le');
                strPos = combined.length;
            } else {
                text = combined.subarray(strPos, strPos + byteLen).toString('utf16le');
                strPos += byteLen;
            }
        } else {
            if (strPos + charCount > combined.length) {
                text = combined.subarray(strPos, combined.length).toString('latin1');
                strPos = combined.length;
            } else {
                text = combined.subarray(strPos, strPos + charCount).toString('latin1');
                strPos += charCount;
            }
        }

        // Skip rich text run data (4 bytes per run)
        strPos += richTextRunCount * 4;
        // Skip phonetic data
        strPos += phoneticSize;

        strings.push(text);
        pos = strPos;
    }

    return strings;
}

// ============================================================================
// Main Parser
// ============================================================================

/**
 * Parse an Excel 97-2003 (.xls) file and return an OfficeParserAST.
 *
 * @param fileBuffer - Buffer containing the .xls file
 * @param config - OfficeParser configuration
 * @returns Parsed AST with sheets, rows, and cells
 */
export async function parseXls(fileBuffer: Buffer, config: Required<OfficeParserConfig>): Promise<OfficeParserAST> {
    // 1. Open OLE2 container and extract Workbook stream
    const ole2 = parseOLE2(fileBuffer);
    let workbookStream: Buffer;

    if (ole2.hasStream('Workbook')) {
        workbookStream = ole2.getStream('Workbook');
    } else if (ole2.hasStream('Book')) {
        // Older BIFF5 name
        workbookStream = ole2.getStream('Book');
    } else {
        throw new Error('XLS: No "Workbook" or "Book" stream found in OLE2 container');
    }

    // 2. Parse BIFF8 records from the workbook stream
    const boundSheets: BoundSheet[] = [];
    const sst: string[] = [];
    const fonts: FontRecord[] = [];
    const xfRecords: XFRecord[] = [];
    let pos = 0;
    let globalBofSeen = false;
    let bofNesting = 0;

    // Phase 1: Read global workbook records (everything before the first sheet BOF/EOF pair)
    while (pos + 4 <= workbookStream.length) {
        const recordType = workbookStream.readUInt16LE(pos);
        const recordLen = workbookStream.readUInt16LE(pos + 2);
        const recordData = workbookStream.subarray(pos + 4, pos + 4 + recordLen);
        pos += 4 + recordLen;

        if (recordType === RT_BOF) {
            bofNesting++;
            if (!globalBofSeen) {
                globalBofSeen = true;
                continue;
            }
            // Skip sub-streams (sheet data in first pass - we read them in Phase 2)
            // Jump to matching EOF
            let depth = 1;
            while (depth > 0 && pos + 4 <= workbookStream.length) {
                const rt = workbookStream.readUInt16LE(pos);
                const rl = workbookStream.readUInt16LE(pos + 2);
                pos += 4 + rl;
                if (rt === RT_BOF) depth++;
                else if (rt === RT_EOF) depth--;
            }
            continue;
        }

        if (recordType === RT_EOF) {
            // End of global workbook records
            break;
        }

        switch (recordType) {
            case RT_BOUNDSHEET: {
                if (recordData.length >= 8) {
                    const bofOffset = recordData.readUInt32LE(0);
                    const visibility = recordData.readUInt8(4);
                    const sheetType = recordData.readUInt8(5);
                    const [name] = readShortBiffString(recordData, 6);
                    boundSheets.push({ bofOffset, visibility, sheetType, name });
                }
                break;
            }

            case RT_SST: {
                // Collect any CONTINUE records that follow
                const continueBuffers: Buffer[] = [];
                while (pos + 4 <= workbookStream.length) {
                    const nextType = workbookStream.readUInt16LE(pos);
                    if (nextType !== RT_CONTINUE) break;
                    const nextLen = workbookStream.readUInt16LE(pos + 2);
                    continueBuffers.push(workbookStream.subarray(pos + 4, pos + 4 + nextLen));
                    pos += 4 + nextLen;
                }
                sst.push(...parseSST(recordData, continueBuffers));
                break;
            }

            case RT_FONT: {
                if (recordData.length >= 14) {
                    const height = recordData.readUInt16LE(0);  // In twips (1/20 pt)
                    const fontFlags = recordData.readUInt16LE(2);
                    const colorIdx = recordData.readUInt16LE(4);
                    const boldWeight = recordData.readUInt16LE(6);
                    const underlineType = recordData.readUInt8(10);
                    const [fontName] = readShortBiffString(recordData, 14);
                    fonts.push({
                        bold: boldWeight >= 700,
                        italic: (fontFlags & 0x02) !== 0,
                        underline: underlineType !== 0,
                        strikethrough: (fontFlags & 0x08) !== 0,
                        size: height,
                        color: colorIdx,
                        name: fontName,
                    });
                }
                break;
            }

            case RT_XF: {
                if (recordData.length >= 8) {
                    const fontIndex = recordData.readUInt16LE(0);
                    const formatIndex = recordData.readUInt16LE(2);
                    const alignByte = recordData.readUInt8(6);
                    const alignment = alignByte & 0x07;
                    xfRecords.push({ fontIndex, formatIndex, alignment });
                }
                break;
            }
        }
    }

    // 3. Phase 2: Read each sheet's cell data
    const content: OfficeContentNode[] = [];

    for (let sheetIdx = 0; sheetIdx < boundSheets.length; sheetIdx++) {
        const sheet = boundSheets[sheetIdx];

        // Skip non-worksheet sheets (charts, VBA modules) and hidden sheets
        if (sheet.sheetType !== 0) continue;

        // Navigate to the sheet's BOF record
        pos = sheet.bofOffset;
        if (pos + 4 > workbookStream.length) continue;

        // Verify BOF
        const bofType = workbookStream.readUInt16LE(pos);
        if (bofType !== RT_BOF) continue;
        const bofLen = workbookStream.readUInt16LE(pos + 2);
        pos += 4 + bofLen;

        // Read cell records for this sheet
        const cells: CellValue[] = [];
        let lastFormulaRow = -1;
        let lastFormulaCol = -1;
        let lastFormulaXf = -1;

        while (pos + 4 <= workbookStream.length) {
            const recordType = workbookStream.readUInt16LE(pos);
            const recordLen = workbookStream.readUInt16LE(pos + 2);
            const recordData = workbookStream.subarray(pos + 4, pos + 4 + recordLen);
            pos += 4 + recordLen;

            if (recordType === RT_EOF) break;

            // Skip nested BOF/EOF (e.g., embedded chart sub-streams)
            if (recordType === RT_BOF) {
                let depth = 1;
                while (depth > 0 && pos + 4 <= workbookStream.length) {
                    const rt = workbookStream.readUInt16LE(pos);
                    const rl = workbookStream.readUInt16LE(pos + 2);
                    pos += 4 + rl;
                    if (rt === RT_BOF) depth++;
                    else if (rt === RT_EOF) depth--;
                }
                continue;
            }

            switch (recordType) {
                case RT_LABELSST: {
                    // Cell referencing SST: row(2), col(2), xf(2), sstIndex(4)
                    if (recordData.length >= 10) {
                        const row = recordData.readUInt16LE(0);
                        const col = recordData.readUInt16LE(2);
                        const xf = recordData.readUInt16LE(4);
                        const sstIdx = recordData.readUInt32LE(6);
                        const value = sstIdx < sst.length ? sst[sstIdx] : '';
                        cells.push({ row, col, value, xfIndex: xf });
                    }
                    break;
                }

                case RT_LABEL: {
                    // Inline string cell: row(2), col(2), xf(2), string
                    if (recordData.length >= 8) {
                        const row = recordData.readUInt16LE(0);
                        const col = recordData.readUInt16LE(2);
                        const xf = recordData.readUInt16LE(4);
                        const [value] = readBiffString(recordData, 6);
                        cells.push({ row, col, value, xfIndex: xf });
                    }
                    break;
                }

                case RT_NUMBER: {
                    // Numeric cell: row(2), col(2), xf(2), double(8)
                    if (recordData.length >= 14) {
                        const row = recordData.readUInt16LE(0);
                        const col = recordData.readUInt16LE(2);
                        const xf = recordData.readUInt16LE(4);
                        const value = recordData.readDoubleLE(6);
                        cells.push({ row, col, value: formatNumber(value), xfIndex: xf });
                    }
                    break;
                }

                case RT_RK: {
                    // Compact number: row(2), col(2), xf(2), rk(4)
                    if (recordData.length >= 10) {
                        const row = recordData.readUInt16LE(0);
                        const col = recordData.readUInt16LE(2);
                        const xf = recordData.readUInt16LE(4);
                        const rkVal = recordData.readInt32LE(6);
                        const value = decodeRK(rkVal);
                        cells.push({ row, col, value: formatNumber(value), xfIndex: xf });
                    }
                    break;
                }

                case RT_MULRK: {
                    // Multiple RK cells: row(2), firstCol(2), [xf(2)+rk(4)]..., lastCol(2)
                    if (recordData.length >= 6) {
                        const row = recordData.readUInt16LE(0);
                        const firstCol = recordData.readUInt16LE(2);
                        const lastCol = recordData.readUInt16LE(recordData.length - 2);
                        let rkPos = 4;
                        for (let col = firstCol; col <= lastCol && rkPos + 6 <= recordData.length - 2; col++) {
                            const xf = recordData.readUInt16LE(rkPos);
                            const rkVal = recordData.readInt32LE(rkPos + 2);
                            const value = decodeRK(rkVal);
                            cells.push({ row, col, value: formatNumber(value), xfIndex: xf });
                            rkPos += 6;
                        }
                    }
                    break;
                }

                case RT_FORMULA: {
                    // Formula cell: row(2), col(2), xf(2), result(8), flags(2), reserved(4), tokens(...)
                    if (recordData.length >= 20) {
                        const row = recordData.readUInt16LE(0);
                        const col = recordData.readUInt16LE(2);
                        const xf = recordData.readUInt16LE(4);

                        // Check if formula result is a string (indicated by 0xFF in byte 6 and 0x00 in byte 12)
                        const byte6 = recordData.readUInt8(6);
                        const byte12 = recordData.readUInt8(12);

                        if (byte6 === 0x00 && byte12 === 0xFF) {
                            // String result - the STRING record follows
                            lastFormulaRow = row;
                            lastFormulaCol = col;
                            lastFormulaXf = xf;
                        } else if (byte6 === 0x01 && byte12 === 0xFF) {
                            // Boolean result
                            const boolVal = recordData.readUInt8(8);
                            cells.push({ row, col, value: boolVal ? 'TRUE' : 'FALSE', xfIndex: xf });
                        } else if (byte6 === 0x02 && byte12 === 0xFF) {
                            // Error result
                            const errCode = recordData.readUInt8(8);
                            const errMap: Record<number, string> = {
                                0x00: '#NULL!', 0x07: '#DIV/0!', 0x0F: '#VALUE!',
                                0x17: '#REF!', 0x1D: '#NAME?', 0x24: '#NUM!', 0x2A: '#N/A'
                            };
                            cells.push({ row, col, value: errMap[errCode] || '#ERR!', xfIndex: xf });
                        } else if (byte6 === 0x03 && byte12 === 0xFF) {
                            // Empty string result
                            cells.push({ row, col, value: '', xfIndex: xf });
                        } else {
                            // Numeric result
                            const value = recordData.readDoubleLE(6);
                            cells.push({ row, col, value: formatNumber(value), xfIndex: xf });
                        }
                    }
                    break;
                }

                case RT_STRING: {
                    // String result of a preceding FORMULA record
                    if (lastFormulaRow >= 0 && recordData.length >= 3) {
                        const [value] = readBiffString(recordData, 0);
                        cells.push({
                            row: lastFormulaRow,
                            col: lastFormulaCol,
                            value,
                            xfIndex: lastFormulaXf
                        });
                        lastFormulaRow = -1;
                    }
                    break;
                }

                case RT_BOOLERR: {
                    // Boolean/Error cell: row(2), col(2), xf(2), value(1), isError(1)
                    if (recordData.length >= 8) {
                        const row = recordData.readUInt16LE(0);
                        const col = recordData.readUInt16LE(2);
                        const xf = recordData.readUInt16LE(4);
                        const val = recordData.readUInt8(6);
                        const isError = recordData.readUInt8(7);

                        if (isError) {
                            const errMap: Record<number, string> = {
                                0x00: '#NULL!', 0x07: '#DIV/0!', 0x0F: '#VALUE!',
                                0x17: '#REF!', 0x1D: '#NAME?', 0x24: '#NUM!', 0x2A: '#N/A'
                            };
                            cells.push({ row, col, value: errMap[val] || '#ERR!', xfIndex: xf });
                        } else {
                            cells.push({ row, col, value: val ? 'TRUE' : 'FALSE', xfIndex: xf });
                        }
                    }
                    break;
                }
            }
        }

        // 4. Build AST nodes for this sheet
        // Sort cells by row, then column
        cells.sort((a, b) => a.row !== b.row ? a.row - b.row : a.col - b.col);

        // Group cells into rows
        const rowMap = new Map<number, CellValue[]>();
        for (const cell of cells) {
            let rowCells = rowMap.get(cell.row);
            if (!rowCells) {
                rowCells = [];
                rowMap.set(cell.row, rowCells);
            }
            rowCells.push(cell);
        }

        const rowNodes: OfficeContentNode[] = [];
        const sortedRows = Array.from(rowMap.keys()).sort((a, b) => a - b);

        for (const rowIdx of sortedRows) {
            const rowCells = rowMap.get(rowIdx)!;
            const cellNodes: OfficeContentNode[] = [];

            for (const cell of rowCells) {
                // Build formatting from XF/Font records
                let formatting: TextFormatting | undefined;
                if (cell.xfIndex < xfRecords.length) {
                    const xf = xfRecords[cell.xfIndex];
                    if (xf.fontIndex < fonts.length) {
                        const font = fonts[xf.fontIndex >= 4 ? xf.fontIndex - 1 : xf.fontIndex]; // Font index 4 is skipped in BIFF8
                        if (font) {
                            formatting = {};
                            if (font.bold) formatting.bold = true;
                            if (font.italic) formatting.italic = true;
                            if (font.underline) formatting.underline = true;
                            if (font.strikethrough) formatting.strikethrough = true;
                            if (font.size) formatting.size = `${font.size / 20}pt`;
                            if (font.name) formatting.font = font.name;
                            // Only include formatting if it has properties
                            if (Object.keys(formatting).length === 0) formatting = undefined;
                        }
                    }
                }

                const textNode: OfficeContentNode = {
                    type: 'text',
                    text: cell.value,
                    ...(formatting ? { formatting } : {}),
                };

                const cellNode: OfficeContentNode = {
                    type: 'cell',
                    text: cell.value,
                    children: [textNode],
                    metadata: {
                        row: cell.row,
                        col: cell.col,
                    } as CellMetadata,
                };

                cellNodes.push(cellNode);
            }

            const rowText = rowCells.map(c => c.value).filter(v => v).join('\t');
            const rowNode: OfficeContentNode = {
                type: 'row',
                text: rowText,
                children: cellNodes,
            };

            rowNodes.push(rowNode);
        }

        const sheetText = rowNodes.map(r => r.text).filter(t => t).join(config.newlineDelimiter ?? '\n');
        const sheetNode: OfficeContentNode = {
            type: 'sheet',
            text: sheetText,
            children: rowNodes,
            metadata: {
                sheetName: sheet.name,
            } as SheetMetadata,
        };

        content.push(sheetNode);
    }

    // 5. Build and return the AST
    const delimiter = config.newlineDelimiter ?? '\n';

    const ast: OfficeParserAST = {
        type: 'xls' as any,
        metadata: {},
        content,
        attachments: [],
        toText: () => {
            return content
                .map(sheet => {
                    if (!sheet.children) return sheet.text || '';
                    return sheet.children
                        .map(row => {
                            if (!row.children) return row.text || '';
                            return row.children
                                .map(cell => cell.text || '')
                                .join('\t');
                        })
                        .filter(t => t !== '')
                        .join(delimiter);
                })
                .filter(t => t !== '')
                .join(delimiter);
        },
    };

    return ast;
}

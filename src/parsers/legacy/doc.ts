/**
 * Word 97-2003 Binary (.doc) Parser
 *
 * Pure TypeScript implementation that reads .doc files by:
 * 1. Using the OLE2 reader to extract WordDocument + Table streams
 * 2. Parsing the FIB (File Information Block) to locate text and formatting
 * 3. Reading the piece table (CLX) to reconstruct document text
 * 4. Parsing paragraph properties (PAPX) to detect paragraphs, headings, and tables
 * 5. Reading the stylesheet (STSH) to identify heading styles
 * 6. Returning a valid OfficeParserAST
 *
 * Ported from the algorithms in:
 * - ref/poi/poi-scratchpad/src/main/java/org/apache/poi/hwpf/ (Apache 2.0)
 *   Specifically: HWPFDocument.java, FileInformationBlock.java,
 *   TextPieceTable.java, PAPBinTable.java, StyleSheet.java,
 *   Paragraph.java, CharacterRun.java, NotesTables.java
 *
 * @module doc
 */

import {
    HeadingMetadata, ListMetadata, NoteMetadata, OfficeContentNode,
    OfficeParserAST, OfficeParserConfig, TextFormatting
} from '../../types';
import { parseOLE2 } from './ole2';

// ============================================================================
// Constants
// ============================================================================

/** Word binary format identifier (at FIB offset 0) */
const WORD_MAGIC = 0xA5EC;

/** Special characters in Word binary documents */
const PARA_MARK = 0x0D;       // Paragraph end
const CELL_MARK = 0x07;       // Table cell end
const SECTION_MARK = 0x0C;    // Section break
const FOOTNOTE_MARK = 0x02;   // Footnote/endnote reference

// ============================================================================
// FIB Parsing
// ============================================================================

interface FIB {
    /** Which table stream: false=0Table, true=1Table */
    whichTableStream: boolean;
    /** Is document encrypted? */
    encrypted: boolean;
    /** Is document complex (piece table needed)? */
    complex: boolean;

    // Text stream lengths for subdocuments (character counts)
    ccpText: number;        // Main document text
    ccpFtn: number;         // Footnotes
    ccpHdd: number;         // Headers/footers
    ccpAtn: number;         // Annotations
    ccpEdn: number;         // Endnotes

    // Table stream offsets and sizes
    fcClx: number;          // Complex file table (piece table)
    lcbClx: number;
    fcStshf: number;        // Stylesheet
    lcbStshf: number;
    fcPlcfbtePapx: number;  // Paragraph properties bin table
    lcbPlcfbtePapx: number;
    fcPlcfbteChpx: number;  // Character properties bin table
    lcbPlcfbteChpx: number;
    fcPlcffndRef: number;   // Footnote references
    lcbPlcffndRef: number;
    fcPlcffndTxt: number;   // Footnote text positions
    lcbPlcffndTxt: number;
    fcPlcfendRef: number;   // Endnote references
    lcbPlcfendRef: number;
    fcPlcfendTxt: number;   // Endnote text positions
    lcbPlcfendTxt: number;
}

/**
 * Parse the File Information Block (FIB) from the WordDocument stream.
 * The FIB is the master index to all document structures.
 */
function parseFIB(mainStream: Buffer): FIB {
    if (mainStream.length < 68) {
        throw new Error('DOC: WordDocument stream too small for FIB');
    }

    const wIdent = mainStream.readUInt16LE(0);
    if (wIdent !== WORD_MAGIC) {
        throw new Error(`DOC: Invalid FIB magic 0x${wIdent.toString(16)} (expected 0xA5EC)`);
    }

    const flags = mainStream.readUInt16LE(0x0A);
    const whichTableStream = (flags & 0x0200) !== 0;
    const encrypted = (flags & 0x0100) !== 0;
    const complex = (flags & 0x0004) !== 0;

    if (encrypted) {
        throw new Error('DOC: Encrypted documents are not supported');
    }

    // FibRgLw97: starts at offset 0x22 (after FibBase + FibRgW97)
    // But the exact offset depends on the FIB version. For Word 97-2003:
    // FibBase = 32 bytes (0x00-0x1F)
    // Then csw (2 bytes) at 0x20, followed by FibRgW97 array
    // Then cslw (2 bytes), followed by FibRgLw97 array

    // Read csw (count of short words in FibRgW97)
    const csw = mainStream.readUInt16LE(0x20);
    const fibRgW97Offset = 0x22;
    const afterFibRgW97 = fibRgW97Offset + csw * 2;

    // Read cslw (count of long words in FibRgLw97)
    const cslw = mainStream.readUInt16LE(afterFibRgW97);
    const fibRgLw97Offset = afterFibRgW97 + 2;

    // FibRgLw97 fields (each 4 bytes):
    // Index 0: cbMac
    // Index 3: ccpText (offset +12)
    // Index 4: ccpFtn  (offset +16)
    // Index 5: ccpHdd  (offset +20)
    // Index 7: ccpAtn  (offset +28)
    // Index 8: ccpEdn  (offset +32)
    const ccpText = cslw > 3 ? mainStream.readInt32LE(fibRgLw97Offset + 12) : 0;
    const ccpFtn = cslw > 4 ? mainStream.readInt32LE(fibRgLw97Offset + 16) : 0;
    const ccpHdd = cslw > 5 ? mainStream.readInt32LE(fibRgLw97Offset + 20) : 0;
    // Index 6: ccpMcr (macro, not used)
    const ccpAtn = cslw > 7 ? mainStream.readInt32LE(fibRgLw97Offset + 28) : 0;
    const ccpEdn = cslw > 8 ? mainStream.readInt32LE(fibRgLw97Offset + 32) : 0;

    // After FibRgLw97 comes cbRgFcLcb (2 bytes) + FibRgFcLcb array
    const afterFibRgLw97 = fibRgLw97Offset + cslw * 4;
    const cbRgFcLcb = mainStream.readUInt16LE(afterFibRgLw97);
    const fcLcbOffset = afterFibRgLw97 + 2;

    // Helper to read fc/lcb pair from FibRgFcLcb97 (each pair is 8 bytes: fc(4) + lcb(4))
    function readFcLcb(index: number): [number, number] {
        if (index >= cbRgFcLcb) return [0, 0];
        const off = fcLcbOffset + index * 8;
        if (off + 8 > mainStream.length) return [0, 0];
        return [mainStream.readUInt32LE(off), mainStream.readUInt32LE(off + 4)];
    }

    // FibRgFcLcb97 field indices (from Apache POI FIBFieldHandler.java):
    const [fcStshf, lcbStshf] = readFcLcb(1);        // STSHF: Stylesheet
    const [fcPlcffndRef, lcbPlcffndRef] = readFcLcb(2);   // PLCFFNDREF: Footnote refs
    const [fcPlcffndTxt, lcbPlcffndTxt] = readFcLcb(3);   // PLCFFNDTXT: Footnote text
    const [fcPlcfbteChpx, lcbPlcfbteChpx] = readFcLcb(12); // PLCFBTECHPX: CHPX bin table
    const [fcPlcfbtePapx, lcbPlcfbtePapx] = readFcLcb(13); // PLCFBTEPAPX: PAPX bin table
    const [fcClx, lcbClx] = readFcLcb(33);            // CLX: Complex File Table (piece table)
    const [fcPlcfendRef, lcbPlcfendRef] = readFcLcb(46);  // PLCFENDREF: Endnote refs
    const [fcPlcfendTxt, lcbPlcfendTxt] = readFcLcb(47);  // PLCFENDTXT: Endnote text

    return {
        whichTableStream,
        encrypted,
        complex,
        ccpText, ccpFtn, ccpHdd, ccpAtn, ccpEdn,
        fcClx, lcbClx,
        fcStshf, lcbStshf,
        fcPlcfbtePapx, lcbPlcfbtePapx,
        fcPlcfbteChpx, lcbPlcfbteChpx,
        fcPlcffndRef, lcbPlcffndRef,
        fcPlcffndTxt, lcbPlcffndTxt,
        fcPlcfendRef, lcbPlcfendRef,
        fcPlcfendTxt, lcbPlcfendTxt,
    };
}

// ============================================================================
// Text Piece Table (CLX) Parsing
// ============================================================================

interface TextPiece {
    /** Character position start (in the unified text) */
    cpStart: number;
    /** Character position end */
    cpEnd: number;
    /** File offset in the WordDocument stream */
    fileOffset: number;
    /** True if text is UTF-16LE, false if ANSI (CP-1252) */
    unicode: boolean;
}

/**
 * Parse the Complex File Table (CLX) from the table stream.
 * The CLX contains the piece table that maps character positions to
 * byte positions in the WordDocument stream.
 *
 * @returns Array of text pieces and the reconstructed full document text
 */
function parseTextPieces(tableStream: Buffer, fib: FIB, mainStream: Buffer): string {
    if (fib.lcbClx === 0) {
        // No CLX = simple document, text starts right after FIB
        // Read from beginning of text in main stream
        // For simple files, text runs from fcMin to fcMac (deprecated fields)
        // But typically all text is right after the FIB
        // Use ccpText to determine length
        const totalCcp = fib.ccpText + fib.ccpFtn + fib.ccpHdd + fib.ccpAtn + fib.ccpEdn;
        if (totalCcp === 0) return '';

        // In simple mode, try reading as ANSI from after FIB
        // The FIB is at least 68 bytes, but the actual text offset varies
        // For BIFF8/Word97, text typically starts at 0x800 (2048) in the main stream
        // But this is unreliable. Without CLX, we fall back to scanning
        return '';
    }

    const clxStart = fib.fcClx;
    const clxEnd = clxStart + fib.lcbClx;

    if (clxEnd > tableStream.length) {
        throw new Error('DOC: CLX extends beyond table stream');
    }

    // The CLX can contain Grpprls (type 0x01) followed by a Pcdt (type 0x02)
    let pos = clxStart;

    // Skip any Grpprl entries (type 0x01)
    while (pos < clxEnd && tableStream.readUInt8(pos) === 0x01) {
        const cbGrpprl = tableStream.readUInt16LE(pos + 1);
        pos += 3 + cbGrpprl;
    }

    // Now expect Pcdt (type 0x02)
    if (pos >= clxEnd || tableStream.readUInt8(pos) !== 0x02) {
        throw new Error('DOC: Expected Pcdt (0x02) in CLX, not found');
    }
    pos += 1;

    const pcdtSize = tableStream.readUInt32LE(pos);
    pos += 4;

    // The PlcPcd is a PLCF: array of (n+1) CPs followed by n PieceDescriptors (8 bytes each)
    // Total size = (n+1)*4 + n*8 = 4*n + 4 + 8*n = 12*n + 4
    // So n = (pcdtSize - 4) / 12
    const numPieces = Math.floor((pcdtSize - 4) / 12);

    if (numPieces <= 0) return '';

    // Read CP array (n+1 entries)
    const cps: number[] = [];
    for (let i = 0; i <= numPieces; i++) {
        cps.push(tableStream.readInt32LE(pos + i * 4));
    }
    pos += (numPieces + 1) * 4;

    // Read PieceDescriptors (8 bytes each)
    const pieces: TextPiece[] = [];
    for (let i = 0; i < numPieces; i++) {
        const pdOffset = pos + i * 8;
        // PieceDescriptor: flags(2) + fc(4) + prm(2)
        const fc = tableStream.readUInt32LE(pdOffset + 2);

        // Bit 30 (0x40000000) clear = Unicode, set = ANSI
        const unicode = (fc & 0x40000000) === 0;
        let fileOffset: number;

        if (unicode) {
            fileOffset = fc;
        } else {
            // Clear bit 30 and divide by 2 to get actual ANSI offset
            fileOffset = (fc & ~0x40000000) >>> 1;
        }

        pieces.push({
            cpStart: cps[i],
            cpEnd: cps[i + 1],
            fileOffset,
            unicode,
        });
    }

    // Reconstruct the full document text from pieces
    let text = '';
    for (const piece of pieces) {
        const charCount = piece.cpEnd - piece.cpStart;
        if (charCount <= 0) continue;

        if (piece.unicode) {
            const byteLen = charCount * 2;
            if (piece.fileOffset + byteLen <= mainStream.length) {
                text += mainStream.subarray(piece.fileOffset, piece.fileOffset + byteLen).toString('utf16le');
            }
        } else {
            if (piece.fileOffset + charCount <= mainStream.length) {
                text += mainStream.subarray(piece.fileOffset, piece.fileOffset + charCount).toString('latin1');
            }
        }
    }

    return text;
}

// ============================================================================
// Stylesheet Parsing
// ============================================================================

interface StyleEntry {
    name: string;
    styleType: number;  // 1=paragraph, 2=character
    baseStyle: number;  // Parent style index
}

/**
 * Parse the StyleSheet (STSH) from the table stream.
 * Returns style definitions used to identify headings and list styles.
 */
function parseStylesheet(tableStream: Buffer, fib: FIB): StyleEntry[] {
    const styles: StyleEntry[] = [];

    if (fib.lcbStshf === 0 || fib.fcStshf + fib.lcbStshf > tableStream.length) {
        return styles;
    }

    let pos = fib.fcStshf;

    // STSHI header: cbStshi(2) + stshif data
    const cbStshi = tableStream.readUInt16LE(pos);
    pos += 2;

    if (cbStshi < 4) return styles;

    // stshif contains: cstd(2) - count of styles
    const cstd = tableStream.readUInt16LE(pos);
    const cbSTDBaseInFile = tableStream.readUInt16LE(pos + 2);

    pos += cbStshi;  // Skip rest of STSHI

    // Read each StyleDescription (STD)
    for (let i = 0; i < cstd && pos < fib.fcStshf + fib.lcbStshf; i++) {
        const cbStd = tableStream.readUInt16LE(pos);
        pos += 2;

        if (cbStd === 0) {
            styles.push({ name: '', styleType: 0, baseStyle: 0xFFF });
            continue;
        }

        const stdStart = pos;

        // StdfBase (variable size, at least 10 bytes)
        if (cbStd >= 10 && pos + 10 <= tableStream.length) {
            const sti = tableStream.readUInt16LE(pos) & 0x0FFF;      // Style identifier
            const flags = tableStream.readUInt16LE(pos);
            const stk = (flags >> 12) & 0x0F;                         // Style type
            const istdBase = tableStream.readUInt16LE(pos + 2) & 0x0FFF; // Base style

            // Skip past StdfBase + StdfPost2000
            let nameOffset = stdStart + cbSTDBaseInFile;

            // Read style name (Unicode string with 2-byte length prefix)
            let name = '';
            if (nameOffset + 2 <= stdStart + cbStd && nameOffset + 2 <= tableStream.length) {
                const nameLen = tableStream.readUInt16LE(nameOffset);
                nameOffset += 2;
                if (nameLen > 0 && nameOffset + nameLen * 2 <= tableStream.length) {
                    name = tableStream.subarray(nameOffset, nameOffset + nameLen * 2).toString('utf16le');
                }
            }

            styles.push({
                name,
                styleType: stk,
                baseStyle: istdBase,
            });
        } else {
            styles.push({ name: '', styleType: 0, baseStyle: 0xFFF });
        }

        pos = stdStart + cbStd;
    }

    return styles;
}

// ============================================================================
// Paragraph Properties (PAPX) Parsing
// ============================================================================

interface ParagraphInfo {
    /** Start character position */
    cpStart: number;
    /** End character position */
    cpEnd: number;
    /** Style index */
    istd: number;
    /** Is this paragraph in a table? */
    inTable: boolean;
    /** Is this the last paragraph of a table row? */
    tableRowEnd: boolean;
    /** Table nesting level */
    tableLevel: number;
    /** Justification: 0=left, 1=center, 2=right, 3=both */
    justification: number;
    /** List level (for outline/list numbering) */
    outlineLevel: number;
    /** List ID (ilfo) for numbered/bulleted lists */
    listId: number;
    /** List level for numbered/bulleted lists */
    listLevel: number;
}

/**
 * Parse the Paragraph Properties Bin Table (PlcBtePapx) to get
 * paragraph boundaries and basic properties.
 *
 * This is a simplified approach: we primarily use style indices and
 * paragraph markers in the text to detect paragraphs, headings, and tables.
 */
function parseParagraphProperties(
    tableStream: Buffer,
    mainStream: Buffer,
    fib: FIB,
    textLength: number
): ParagraphInfo[] {
    const paragraphs: ParagraphInfo[] = [];

    if (fib.lcbPlcfbtePapx === 0) return paragraphs;
    if (fib.fcPlcfbtePapx + fib.lcbPlcfbtePapx > tableStream.length) return paragraphs;

    // PlcBtePapx is a PLCF: (n+1) CPs (4 bytes each) + n page numbers (4 bytes each)
    // Total size = (n+1)*4 + n*4 = 8n + 4
    // n = (size - 4) / 8
    const plcSize = fib.lcbPlcfbtePapx;
    const numEntries = Math.floor((plcSize - 4) / 8);
    if (numEntries <= 0) return paragraphs;

    const plcStart = fib.fcPlcfbtePapx;

    // Read CP array
    const cps: number[] = [];
    for (let i = 0; i <= numEntries; i++) {
        cps.push(tableStream.readUInt32LE(plcStart + i * 4));
    }

    // Read page numbers (FDPs = Formatted Disk Pages, 512 bytes each)
    const pageNumOffset = plcStart + (numEntries + 1) * 4;

    for (let i = 0; i < numEntries; i++) {
        const pageNum = tableStream.readUInt32LE(pageNumOffset + i * 4);
        const pageOffset = pageNum * 512;

        if (pageOffset + 512 > mainStream.length) continue;

        // Parse the PapxFkp (Formatted Disk Page for paragraph properties)
        // Last byte of the FKP = number of entries (crun)
        const crun = mainStream.readUInt8(pageOffset + 511);

        // FKP layout:
        // - (crun+1) FCs (4 bytes each) at start
        // - crun BX entries (13 bytes each for Papx): each BX has offset(1 byte) to PAPX in page
        for (let j = 0; j < crun; j++) {
            const fcFirst = mainStream.readUInt32LE(pageOffset + j * 4);
            const fcLim = mainStream.readUInt32LE(pageOffset + (j + 1) * 4);

            // BX offset is at (crun+1)*4 + j*13
            const bxPos = pageOffset + (crun + 1) * 4 + j * 13;
            if (bxPos + 1 > mainStream.length) continue;

            const papxOffsetInPage = mainStream.readUInt8(bxPos) * 2;
            if (papxOffsetInPage === 0) {
                // No PAPX data, use default style (istd=0)
                paragraphs.push({
                    cpStart: fcFirst,
                    cpEnd: fcLim,
                    istd: 0,
                    inTable: false,
                    tableRowEnd: false,
                    tableLevel: 0,
                    justification: 0,
                    outlineLevel: 9,
                    listId: 0,
                    listLevel: 0,
                });
                continue;
            }

            const papxPos = pageOffset + papxOffsetInPage;
            if (papxPos + 2 > mainStream.length) continue;

            // PAPX: first byte is cb (count of bytes / 2), if 0 then next byte is cb
            let cb = mainStream.readUInt8(papxPos);
            let sprmStart = papxPos + 1;

            if (cb === 0) {
                cb = mainStream.readUInt8(papxPos + 1);
                sprmStart = papxPos + 2;
            }

            // First 2 bytes of grpprl are istd (style index)
            let istd = 0;
            if (sprmStart + 2 <= mainStream.length) {
                istd = mainStream.readUInt16LE(sprmStart);
            }

            // Parse SPRMs for table and list info
            let inTable = false;
            let tableRowEnd = false;
            let tableLevel = 0;
            let justification = 0;
            let outlineLevel = 9;
            let listId = 0;
            let listLevel = 0;

            const sprmEnd = Math.min(papxPos + cb * 2 + 1, mainStream.length);
            let sprmPos = sprmStart + 2;  // Skip istd

            while (sprmPos + 2 <= sprmEnd) {
                const sprm = mainStream.readUInt16LE(sprmPos);
                const sprmOp = sprm & 0x01FF;
                const sprmType = (sprm >> 13) & 0x07;

                // Determine SPRM operand size
                let operandSize: number;
                switch (sprmType) {
                    case 0: operandSize = 1; break;  // Toggle
                    case 1: operandSize = 1; break;  // Byte
                    case 2: operandSize = 2; break;  // Word
                    case 3: operandSize = 4; break;  // Dword
                    case 4: operandSize = 2; break;  // Word
                    case 5: operandSize = 2; break;  // Word
                    case 6: // Variable
                        if (sprmPos + 3 <= sprmEnd) {
                            operandSize = mainStream.readUInt8(sprmPos + 2) + 1;
                        } else {
                            operandSize = 1;
                        }
                        break;
                    case 7: operandSize = 3; break;  // 3 bytes
                    default: operandSize = 1;
                }

                sprmPos += 2;  // Past the sprm id

                // Check specific SPRM IDs
                // sprmPFInTable (0x2416): paragraph is in table
                if (sprm === 0x2416 && sprmPos < sprmEnd) {
                    inTable = mainStream.readUInt8(sprmPos) !== 0;
                }
                // sprmPFTtp (0x2417): paragraph is table row terminator
                if (sprm === 0x2417 && sprmPos < sprmEnd) {
                    tableRowEnd = mainStream.readUInt8(sprmPos) !== 0;
                }
                // sprmPItap (0x6649): table nesting level
                if (sprm === 0x6649 && sprmPos + 4 <= sprmEnd) {
                    tableLevel = mainStream.readInt32LE(sprmPos);
                }
                // sprmPJc (0x2461): justification
                if ((sprm === 0x2461 || sprm === 0x2403) && sprmPos < sprmEnd) {
                    justification = mainStream.readUInt8(sprmPos);
                }
                // sprmPOutLvl (0x2640): outline level
                if (sprm === 0x2640 && sprmPos < sprmEnd) {
                    outlineLevel = mainStream.readUInt8(sprmPos);
                }
                // sprmPIlfo (0x460B): list ID
                if (sprm === 0x460B && sprmPos + 2 <= sprmEnd) {
                    listId = mainStream.readUInt16LE(sprmPos);
                }
                // sprmPIlvl (0x260A): list level
                if (sprm === 0x260A && sprmPos < sprmEnd) {
                    listLevel = mainStream.readUInt8(sprmPos);
                }

                sprmPos += operandSize;
            }

            paragraphs.push({
                cpStart: fcFirst,
                cpEnd: fcLim,
                istd, inTable, tableRowEnd, tableLevel,
                justification, outlineLevel, listId, listLevel,
            });
        }
    }

    // Sort by start position
    paragraphs.sort((a, b) => a.cpStart - b.cpStart);
    return paragraphs;
}

// ============================================================================
// Main Parser
// ============================================================================

/**
 * Parse a Word 97-2003 (.doc) file and return an OfficeParserAST.
 *
 * @param fileBuffer - Buffer containing the .doc file
 * @param config - OfficeParser configuration
 * @returns Parsed AST with paragraphs, headings, tables, and lists
 */
export async function parseDoc(fileBuffer: Buffer, config: Required<OfficeParserConfig>): Promise<OfficeParserAST> {
    // 1. Open OLE2 container
    const ole2 = parseOLE2(fileBuffer);

    if (!ole2.hasStream('WordDocument')) {
        throw new Error('DOC: No "WordDocument" stream found in OLE2 container');
    }

    const mainStream = ole2.getStream('WordDocument');

    // 2. Parse FIB
    const fib = parseFIB(mainStream);

    // 3. Get table stream
    const tableStreamName = fib.whichTableStream ? '1Table' : '0Table';
    if (!ole2.hasStream(tableStreamName)) {
        throw new Error(`DOC: Table stream "${tableStreamName}" not found`);
    }
    const tableStream = ole2.getStream(tableStreamName);

    // 4. Extract document text via piece table
    const fullText = parseTextPieces(tableStream, fib, mainStream);

    if (fullText.length === 0) {
        return {
            type: 'doc' as any,
            metadata: {},
            content: [],
            attachments: [],
            toText: () => '',
        };
    }

    // 5. Parse stylesheet for heading detection
    const styles = parseStylesheet(tableStream, fib);

    // Helper to check if a style is a heading and get its level
    function getHeadingLevel(istd: number): number | undefined {
        if (istd >= styles.length) return undefined;
        const style = styles[istd];
        if (!style || !style.name) return undefined;

        const name = style.name.toLowerCase().replace(/\0/g, '');

        // Check for "heading N" pattern (English)
        const match = name.match(/^heading\s*(\d+)$/);
        if (match) {
            const level = parseInt(match[1], 10);
            if (level >= 1 && level <= 9) return level;
        }

        // Check for "title" style
        if (name === 'title') return 1;
        if (name === 'subtitle') return 2;

        // Check parent style chain
        if (style.baseStyle !== 0xFFF && style.baseStyle !== istd && style.baseStyle < styles.length) {
            return getHeadingLevel(style.baseStyle);
        }

        return undefined;
    }

    // Helper to check if a style is a list style
    function isListStyle(istd: number): 'ordered' | 'unordered' | undefined {
        if (istd >= styles.length) return undefined;
        const style = styles[istd];
        if (!style || !style.name) return undefined;

        const name = style.name.toLowerCase().replace(/\0/g, '');
        if (name.includes('list bullet') || name.includes('listbullet')) return 'unordered';
        if (name.includes('list number') || name.includes('listnumber')) return 'ordered';
        if (name.includes('list paragraph') || name.includes('listparagraph')) return undefined;  // Could be either

        return undefined;
    }

    // 6. Split text into paragraphs by scanning for paragraph marks
    //    The main document text is the first ccpText characters
    const mainText = fullText.substring(0, fib.ccpText);

    // Also extract footnote and endnote text
    const footnoteStart = fib.ccpText + 1;  // +1 for separator
    const footnoteText = fib.ccpFtn > 0
        ? fullText.substring(footnoteStart, footnoteStart + fib.ccpFtn - 1)
        : '';

    const endnoteStart = footnoteStart + fib.ccpFtn + fib.ccpHdd + fib.ccpAtn;
    const endnoteText = fib.ccpEdn > 0
        ? fullText.substring(endnoteStart, endnoteStart + fib.ccpEdn - 1)
        : '';

    // 7. Parse paragraph properties for PAPX info
    const papxInfo = parseParagraphProperties(tableStream, mainStream, fib, mainText.length);

    // Build a map from character position → PAPX info
    // For simplicity, we'll match paragraphs by scanning for \r (0x0D) in the text
    const content: OfficeContentNode[] = [];
    const notes: OfficeContentNode[] = [];

    // State for table detection
    let currentTableRows: OfficeContentNode[][] = [];
    let currentTableRowCells: OfficeContentNode[] = [];
    let inTable = false;

    // State for list detection
    let currentListId = '';
    let currentListIndex = 0;

    // Split main text into paragraphs at paragraph marks
    const paragraphTexts: string[] = [];
    let paraStart = 0;
    for (let i = 0; i < mainText.length; i++) {
        const ch = mainText.charCodeAt(i);
        if (ch === PARA_MARK || ch === CELL_MARK) {
            paragraphTexts.push(mainText.substring(paraStart, i));
            paraStart = i + 1;
        }
    }
    // Last segment
    if (paraStart < mainText.length) {
        paragraphTexts.push(mainText.substring(paraStart));
    }

    // Match paragraphs with PAPX info (simplified: use style index at each paragraph position)
    let cpPosition = 0;
    let papxIdx = 0;

    for (const paraText of paragraphTexts) {
        const textContent = paraText
            .replace(/[\x00-\x06\x08-\x0C\x0E-\x1F]/g, '') // Remove control characters
            .replace(/\x07/g, '');  // Remove cell marks

        cpPosition += paraText.length + 1;  // +1 for the paragraph/cell mark

        // Find the PAPX entry covering this paragraph's position
        let papx: ParagraphInfo | undefined;
        while (papxIdx < papxInfo.length && papxInfo[papxIdx].cpEnd <= cpPosition - paraText.length - 1) {
            papxIdx++;
        }
        if (papxIdx < papxInfo.length) {
            papx = papxInfo[papxIdx];
        }

        const istd = papx?.istd ?? 0;

        // Check if this paragraph is in a table
        const isInTable = papx?.inTable ?? false;
        const isRowEnd = papx?.tableRowEnd ?? false;

        if (isInTable && !isRowEnd) {
            // This is a table cell content
            const cellNode: OfficeContentNode = {
                type: 'cell',
                text: textContent,
                children: [{
                    type: 'text',
                    text: textContent,
                }],
                metadata: {
                    row: currentTableRows.length,
                    col: currentTableRowCells.length,
                },
            };
            currentTableRowCells.push(cellNode);
            inTable = true;
            continue;
        }

        if (isRowEnd) {
            // End of a table row
            if (currentTableRowCells.length > 0) {
                currentTableRows.push(currentTableRowCells);
                currentTableRowCells = [];
            }
            inTable = true;
            continue;
        }

        // If we were in a table and now we're not, emit the table
        if (inTable && !isInTable) {
            if (currentTableRowCells.length > 0) {
                currentTableRows.push(currentTableRowCells);
                currentTableRowCells = [];
            }
            if (currentTableRows.length > 0) {
                const tableNode = buildTableNode(currentTableRows, config);
                content.push(tableNode);
                currentTableRows = [];
            }
            inTable = false;
        }

        // Skip empty paragraphs (but still process table endings above)
        if (!textContent.trim()) continue;

        // Determine alignment
        const jc = papx?.justification ?? 0;
        const alignment = jc === 1 ? 'center' as const :
            jc === 2 ? 'right' as const :
                jc === 3 ? 'justify' as const : 'left' as const;

        // Check for heading
        const headingLevel = getHeadingLevel(istd);
        if (headingLevel !== undefined) {
            content.push({
                type: 'heading',
                text: textContent,
                children: [{ type: 'text', text: textContent }],
                metadata: {
                    level: Math.min(headingLevel, 6),
                    alignment,
                } as HeadingMetadata,
            });
            currentListId = '';
            continue;
        }

        // Check for list item
        const listType = isListStyle(istd);
        const hasListId = papx?.listId && papx.listId > 0;

        if (listType || hasListId) {
            const lType = listType || 'unordered';
            const lid = String(papx?.listId || istd);
            const indentation = papx?.listLevel ?? 0;

            if (lid !== currentListId) {
                currentListId = lid;
                currentListIndex = 0;
            }

            content.push({
                type: 'list',
                text: textContent,
                children: [{ type: 'text', text: textContent }],
                metadata: {
                    listType: lType,
                    indentation,
                    listId: lid,
                    itemIndex: currentListIndex,
                    alignment,
                } as ListMetadata,
            });
            currentListIndex++;
            continue;
        }

        // Regular paragraph
        currentListId = '';
        content.push({
            type: 'paragraph',
            text: textContent,
            children: [{ type: 'text', text: textContent }],
            metadata: { alignment },
        });
    }

    // Flush any remaining table
    if (currentTableRowCells.length > 0) {
        currentTableRows.push(currentTableRowCells);
    }
    if (currentTableRows.length > 0) {
        content.push(buildTableNode(currentTableRows, config));
    }

    // 8. Process footnotes and endnotes
    if (!config.ignoreNotes) {
        if (footnoteText.trim()) {
            const footnoteParagraphs = footnoteText.split('\r').filter(t => t.trim());
            for (let i = 0; i < footnoteParagraphs.length; i++) {
                const noteNode: OfficeContentNode = {
                    type: 'note',
                    text: footnoteParagraphs[i].replace(/[\x00-\x06\x08-\x0C\x0E-\x1F]/g, ''),
                    children: [{
                        type: 'text',
                        text: footnoteParagraphs[i].replace(/[\x00-\x06\x08-\x0C\x0E-\x1F]/g, ''),
                    }],
                    metadata: {
                        noteType: 'footnote',
                        noteId: String(i + 1),
                    } as NoteMetadata,
                };

                if (config.putNotesAtLast) {
                    notes.push(noteNode);
                } else {
                    content.push(noteNode);
                }
            }
        }

        if (endnoteText.trim()) {
            const endnoteParagraphs = endnoteText.split('\r').filter(t => t.trim());
            for (let i = 0; i < endnoteParagraphs.length; i++) {
                const noteNode: OfficeContentNode = {
                    type: 'note',
                    text: endnoteParagraphs[i].replace(/[\x00-\x06\x08-\x0C\x0E-\x1F]/g, ''),
                    children: [{
                        type: 'text',
                        text: endnoteParagraphs[i].replace(/[\x00-\x06\x08-\x0C\x0E-\x1F]/g, ''),
                    }],
                    metadata: {
                        noteType: 'endnote',
                        noteId: String(i + 1),
                    } as NoteMetadata,
                };

                if (config.putNotesAtLast) {
                    notes.push(noteNode);
                } else {
                    content.push(noteNode);
                }
            }
        }
    }

    // Append notes at the end if configured
    if (config.putNotesAtLast && notes.length > 0) {
        content.push(...notes);
    }

    // 9. Build and return AST
    const delimiter = config.newlineDelimiter ?? '\n';

    return {
        type: 'doc' as any,
        metadata: {},
        content,
        attachments: [],
        toText: () => {
            return content
                .map(node => node.text || '')
                .filter(t => t !== '')
                .join(delimiter);
        },
    };
}

// ============================================================================
// Table Building Helper
// ============================================================================

function buildTableNode(
    rows: OfficeContentNode[][],
    config: Required<OfficeParserConfig>
): OfficeContentNode {
    const delimiter = config.newlineDelimiter ?? '\n';
    const rowNodes: OfficeContentNode[] = [];

    for (let r = 0; r < rows.length; r++) {
        const cells = rows[r];
        // Update cell metadata with correct row indices
        for (let c = 0; c < cells.length; c++) {
            if (cells[c].metadata) {
                (cells[c].metadata as any).row = r;
                (cells[c].metadata as any).col = c;
            }
        }

        const rowText = cells.map(c => c.text || '').join('\t');
        rowNodes.push({
            type: 'row',
            text: rowText,
            children: cells,
        });
    }

    const tableText = rowNodes.map(r => r.text || '').join(delimiter);
    return {
        type: 'table',
        text: tableText,
        children: rowNodes,
    };
}

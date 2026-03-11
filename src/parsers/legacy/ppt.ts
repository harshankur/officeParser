/**
 * PowerPoint 97-2003 Binary (.ppt) Parser
 *
 * Pure TypeScript implementation that reads .ppt files by:
 * 1. Using the OLE2 reader to extract the "PowerPoint Document" stream
 * 2. Following the UserEditAtom chain from "Current User" to find all records
 * 3. Locating SlideListWithText containers to extract slide text
 * 4. Parsing TextHeaderAtom, TextCharsAtom, and TextBytesAtom records
 * 5. Returning a valid OfficeParserAST with slides in presentation order
 *
 * Ported from the algorithms in:
 * - ref/poi/poi-scratchpad/src/main/java/org/apache/poi/hslf/ (Apache 2.0)
 *   Specifically: HSLFSlideShow.java, HSLFSlideShowImpl.java,
 *   HSLFTextParagraph.java, Record.java, RecordTypes.java,
 *   TextHeaderAtom.java, TextCharsAtom.java, TextBytesAtom.java
 *
 * @module ppt
 */

import {
    OfficeContentNode, OfficeParserAST,
    OfficeParserConfig, SlideMetadata
} from '../../types';
import { parseOLE2 } from './ole2';

// ============================================================================
// PPT Record Type Constants
// ============================================================================

// Container types (recVer = 0xF)
const RT_DOCUMENT           = 1000;
const RT_SLIDE              = 1006;
const RT_NOTES              = 1008;
const RT_SLIDE_LIST_WITH_TEXT = 4080;
const RT_ENVIRONMENT        = 1010;

// Atom types (recVer = 0x0)
const RT_DOCUMENT_ATOM      = 1001;
const RT_SLIDE_ATOM         = 1007;
const RT_SLIDE_PERSIST_ATOM = 1011;
const RT_TEXT_HEADER_ATOM   = 3999;
const RT_TEXT_CHARS_ATOM    = 4000;
const RT_TEXT_BYTES_ATOM    = 4008;
const RT_STYLE_TEXT_PROP    = 4001;
const RT_USER_EDIT_ATOM     = 4085;
const RT_CURRENT_USER_ATOM  = 4086;
const RT_PERSIST_PTR_HOLDER = 6002;
const RT_END_DOCUMENT       = 1002;
const RT_CSTRING            = 4026;

/** Text types from TextHeaderAtom */
const TEXT_TYPE_TITLE        = 0;
const TEXT_TYPE_BODY         = 1;
const TEXT_TYPE_NOTES        = 4;
const TEXT_TYPE_CENTER_BODY  = 5;
const TEXT_TYPE_CENTER_TITLE = 6;
const TEXT_TYPE_HALF_BODY    = 7;
const TEXT_TYPE_QUARTER_BODY = 8;

// ============================================================================
// Types
// ============================================================================

interface PPTRecord {
    /** Record version (bits 0-3 of first 2 bytes) */
    recVer: number;
    /** Record instance (bits 4-15 of first 2 bytes) */
    recInstance: number;
    /** Record type */
    recType: number;
    /** Record data length (excluding 8-byte header) */
    recLen: number;
    /** Offset of the record data in the stream (after 8-byte header) */
    dataOffset: number;
    /** Offset of the record header (start of the 8-byte header) */
    headerOffset: number;
}

interface SlideText {
    /** The text type (title, body, notes, etc.) */
    textType: number;
    /** The extracted text */
    text: string;
}

interface SlideData {
    /** Slide index in presentation order (0-based) */
    index: number;
    /** Persist ID reference */
    refId: number;
    /** Slide identifier */
    slideId: number;
    /** Text blocks for this slide */
    textBlocks: SlideText[];
}

// ============================================================================
// Record Parsing
// ============================================================================

/**
 * Parse a single PPT record header at the given offset.
 * PPT records have an 8-byte header:
 * - Bytes 0-1: recVer (4 bits) + recInstance (12 bits)
 * - Bytes 2-3: recType (16 bits)
 * - Bytes 4-7: recLen (32 bits, length of data after header)
 */
function readRecordHeader(stream: Buffer, offset: number): PPTRecord | null {
    if (offset + 8 > stream.length) return null;

    const verInst = stream.readUInt16LE(offset);
    const recVer = verInst & 0x0F;
    const recInstance = (verInst >> 4) & 0x0FFF;
    const recType = stream.readUInt16LE(offset + 2);
    const recLen = stream.readUInt32LE(offset + 4);

    return {
        recVer,
        recInstance,
        recType,
        recLen,
        dataOffset: offset + 8,
        headerOffset: offset,
    };
}

/**
 * Check if a record is a container (has child records).
 * Containers have recVer = 0xF.
 */
function isContainer(rec: PPTRecord): boolean {
    return rec.recVer === 0x0F;
}

/**
 * Iterate over child records within a container.
 */
function* iterateChildren(stream: Buffer, parent: PPTRecord): Generator<PPTRecord> {
    let pos = parent.dataOffset;
    const end = Math.min(parent.dataOffset + parent.recLen, stream.length);

    while (pos + 8 <= end) {
        const rec = readRecordHeader(stream, pos);
        if (!rec) break;

        // Safety: don't go past the parent's boundary
        if (rec.dataOffset + rec.recLen > end) {
            // Truncated record, skip it
            break;
        }

        yield rec;
        pos = rec.dataOffset + rec.recLen;
    }
}

/**
 * Find all records of a given type within a container (non-recursive).
 */
function findRecordsInContainer(stream: Buffer, parent: PPTRecord, type: number): PPTRecord[] {
    const results: PPTRecord[] = [];
    for (const child of iterateChildren(stream, parent)) {
        if (child.recType === type) {
            results.push(child);
        }
    }
    return results;
}

/**
 * Recursively find all records of a given type within a container and its children.
 */
function findRecordsRecursive(stream: Buffer, parent: PPTRecord, type: number): PPTRecord[] {
    const results: PPTRecord[] = [];
    for (const child of iterateChildren(stream, parent)) {
        if (child.recType === type) {
            results.push(child);
        }
        if (isContainer(child)) {
            results.push(...findRecordsRecursive(stream, child, type));
        }
    }
    return results;
}

// ============================================================================
// Text Extraction
// ============================================================================

/**
 * Extract text from a TextCharsAtom (Unicode, UTF-16LE).
 */
function readTextCharsAtom(stream: Buffer, rec: PPTRecord): string {
    const len = Math.min(rec.recLen, stream.length - rec.dataOffset);
    if (len <= 0) return '';
    return stream.subarray(rec.dataOffset, rec.dataOffset + len).toString('utf16le');
}

/**
 * Extract text from a TextBytesAtom (ANSI/Latin-1).
 */
function readTextBytesAtom(stream: Buffer, rec: PPTRecord): string {
    const len = Math.min(rec.recLen, stream.length - rec.dataOffset);
    if (len <= 0) return '';
    return stream.subarray(rec.dataOffset, rec.dataOffset + len).toString('latin1');
}

/**
 * Extract text blocks from a SlideListWithText container.
 * Each SlidePersistAtom marks the start of a new slide's text.
 * TextHeaderAtom + TextCharsAtom/TextBytesAtom pairs follow.
 */
function extractSlideTexts(stream: Buffer, slwt: PPTRecord): SlideData[] {
    const slides: SlideData[] = [];
    let currentSlide: SlideData | null = null;
    let currentTextType = TEXT_TYPE_BODY;
    let slideIndex = 0;

    for (const child of iterateChildren(stream, slwt)) {
        switch (child.recType) {
            case RT_SLIDE_PERSIST_ATOM: {
                // SlidePersistAtom marks the start of a new slide's text
                // Structure: refID(4), flags(4), numTexts(4), slideId(4), reserved(4)
                if (child.recLen >= 20) {
                    const refId = stream.readUInt32LE(child.dataOffset);
                    const slideId = stream.readUInt32LE(child.dataOffset + 12);
                    currentSlide = {
                        index: slideIndex++,
                        refId,
                        slideId,
                        textBlocks: [],
                    };
                    slides.push(currentSlide);
                }
                break;
            }

            case RT_TEXT_HEADER_ATOM: {
                // TextHeaderAtom: 4 bytes indicating text type
                if (child.recLen >= 4) {
                    currentTextType = stream.readInt32LE(child.dataOffset);
                }
                break;
            }

            case RT_TEXT_CHARS_ATOM: {
                // Unicode text
                const text = readTextCharsAtom(stream, child);
                if (currentSlide && text) {
                    currentSlide.textBlocks.push({
                        textType: currentTextType,
                        text,
                    });
                }
                break;
            }

            case RT_TEXT_BYTES_ATOM: {
                // ANSI text
                const text = readTextBytesAtom(stream, child);
                if (currentSlide && text) {
                    currentSlide.textBlocks.push({
                        textType: currentTextType,
                        text,
                    });
                }
                break;
            }
        }
    }

    return slides;
}

// ============================================================================
// Slide Container Text Extraction
// ============================================================================

/**
 * Extract text from individual Slide containers (type 1006) found in the stream.
 * This is used as a fallback when SlideListWithText doesn't contain text atoms
 * (only SlidePersistAtom records), which happens in many PPT files where
 * the text is stored inside the Slide containers themselves.
 */
function extractTextFromSlideContainers(
    stream: Buffer,
    documentRecord: PPTRecord,
    existingSlides: SlideData[]
): SlideData[] {
    // Scan the entire stream (after Document container) for Slide containers
    const slides: SlideData[] = [];
    let slideIndex = 0;
    let pos = 0;

    while (pos + 8 <= stream.length) {
        const rec = readRecordHeader(stream, pos);
        if (!rec) break;

        if (rec.recLen > stream.length - rec.dataOffset) {
            pos += 8;
            continue;
        }

        if (rec.recType === RT_SLIDE && isContainer(rec)) {
            const slideData: SlideData = {
                index: slideIndex,
                refId: existingSlides[slideIndex]?.refId ?? slideIndex + 1,
                slideId: existingSlides[slideIndex]?.slideId ?? slideIndex,
                textBlocks: [],
            };

            // Extract text from this Slide container recursively
            extractTextFromContainer(stream, rec, slideData);

            // Filter out master slide template text
            slideData.textBlocks = slideData.textBlocks.filter(tb => {
                const t = tb.text.trim();
                return t && !t.startsWith('Click to edit') && t !== '*';
            });

            if (slideData.textBlocks.length > 0) {
                slides.push(slideData);
                slideIndex++;
            }
        }

        // Skip to next top-level record
        if (isContainer(rec)) {
            pos = rec.dataOffset + rec.recLen;
        } else {
            pos = rec.dataOffset + rec.recLen;
        }
    }

    return slides;
}

/**
 * Extract text from Notes containers (type 1008) found in the stream.
 */
function extractTextFromNotesContainers(
    stream: Buffer,
    documentRecord: PPTRecord,
    existingNotes: SlideData[]
): SlideData[] {
    const notes: SlideData[] = [];
    let noteIndex = 0;
    let pos = 0;

    while (pos + 8 <= stream.length) {
        const rec = readRecordHeader(stream, pos);
        if (!rec) break;

        if (rec.recLen > stream.length - rec.dataOffset) {
            pos += 8;
            continue;
        }

        if (rec.recType === RT_NOTES && isContainer(rec)) {
            const noteData: SlideData = {
                index: noteIndex,
                refId: existingNotes[noteIndex]?.refId ?? noteIndex + 1,
                slideId: existingNotes[noteIndex]?.slideId ?? noteIndex,
                textBlocks: [],
            };

            extractTextFromContainer(stream, rec, noteData);

            // Filter out template text and slide number placeholders
            noteData.textBlocks = noteData.textBlocks.filter(tb => {
                const t = tb.text.trim();
                return t && !t.startsWith('Click to edit') && t !== '*' && !/^\d+$/.test(t);
            });

            if (noteData.textBlocks.length > 0) {
                notes.push(noteData);
                noteIndex++;
            }
        }

        if (isContainer(rec)) {
            pos = rec.dataOffset + rec.recLen;
        } else {
            pos = rec.dataOffset + rec.recLen;
        }
    }

    return notes;
}

/**
 * Recursively extract TextHeaderAtom + TextCharsAtom/TextBytesAtom from a container.
 */
function extractTextFromContainer(stream: Buffer, parent: PPTRecord, slideData: SlideData): void {
    let currentTextType = TEXT_TYPE_BODY;

    for (const child of iterateChildren(stream, parent)) {
        switch (child.recType) {
            case RT_TEXT_HEADER_ATOM:
                if (child.recLen >= 4) {
                    currentTextType = stream.readInt32LE(child.dataOffset);
                }
                break;

            case RT_TEXT_CHARS_ATOM: {
                const text = readTextCharsAtom(stream, child);
                if (text) {
                    slideData.textBlocks.push({ textType: currentTextType, text });
                }
                break;
            }

            case RT_TEXT_BYTES_ATOM: {
                const text = readTextBytesAtom(stream, child);
                if (text) {
                    slideData.textBlocks.push({ textType: currentTextType, text });
                }
                break;
            }

            default:
                if (isContainer(child)) {
                    extractTextFromContainer(stream, child, slideData);
                }
                break;
        }
    }
}

// ============================================================================
// Main Parser
// ============================================================================

/**
 * Parse a PowerPoint 97-2003 (.ppt) file and return an OfficeParserAST.
 *
 * @param fileBuffer - Buffer containing the .ppt file
 * @param config - OfficeParser configuration
 * @returns Parsed AST with slides containing paragraphs
 */
export async function parsePpt(fileBuffer: Buffer, config: Required<OfficeParserConfig>): Promise<OfficeParserAST> {
    // 1. Open OLE2 container
    const ole2 = parseOLE2(fileBuffer);

    // Find the PowerPoint Document stream (case-insensitive)
    let pptStreamName: string | undefined;
    for (const name of ole2.listStreams()) {
        if (name.toLowerCase() === 'powerpoint document') {
            pptStreamName = name;
            break;
        }
    }

    if (!pptStreamName) {
        throw new Error('PPT: No "PowerPoint Document" stream found in OLE2 container');
    }

    const pptStream = ole2.getStream(pptStreamName);

    if (pptStream.length < 8) {
        throw new Error('PPT: PowerPoint Document stream is too small');
    }

    // 2. Find the Document container record
    // Strategy: Scan top-level records to find the Document container (type 1000)
    // Top-level records in the PowerPoint Document stream include:
    // - CurrentUser info (not always at top level)
    // - UserEditAtom chain
    // - PersistDirectoryEntry
    // - Document container
    //
    // We scan sequentially for the Document record

    let documentRecord: PPTRecord | null = null;
    let pos = 0;

    // First pass: Find the Document container by scanning all top-level records
    while (pos + 8 <= pptStream.length) {
        const rec = readRecordHeader(pptStream, pos);
        if (!rec) break;

        // Safety: check for unreasonable record lengths
        if (rec.recLen > pptStream.length - rec.dataOffset) {
            // Truncated record, adjust length
            rec.recLen = pptStream.length - rec.dataOffset;
        }

        if (rec.recType === RT_DOCUMENT && isContainer(rec)) {
            documentRecord = rec;
            break;
        }

        pos = rec.dataOffset + rec.recLen;
    }

    if (!documentRecord) {
        // Fallback: try to extract text directly by scanning for text atoms
        return extractTextByScanning(pptStream, config);
    }

    // 3. Find SlideListWithText containers inside the Document
    //    Instance 0 = slide text, instance 1 = master slide text, instance 2 = notes text
    const slwtRecords = findRecordsInContainer(pptStream, documentRecord, RT_SLIDE_LIST_WITH_TEXT);

    let slideSLWT: PPTRecord | null = null;
    let notesSLWT: PPTRecord | null = null;

    for (const slwt of slwtRecords) {
        if (slwt.recInstance === 0) slideSLWT = slwt;
        else if (slwt.recInstance === 2) notesSLWT = slwt;
    }

    // 4. Extract slide text from the SlideListWithText
    let slideDataList: SlideData[] = slideSLWT
        ? extractSlideTexts(pptStream, slideSLWT)
        : [];

    // If SlideListWithText didn't contain text (only SlidePersistAtoms),
    // extract text directly from Slide containers (type 1006) in the stream
    const hasText = slideDataList.some(s => s.textBlocks.length > 0);
    if (!hasText) {
        slideDataList = extractTextFromSlideContainers(pptStream, documentRecord, slideDataList);
    }

    // Extract notes text
    let notesDataList: SlideData[] = notesSLWT && !config.ignoreNotes
        ? extractSlideTexts(pptStream, notesSLWT)
        : [];

    // If notes SlideListWithText didn't contain text, extract from Notes containers
    if (!config.ignoreNotes && notesSLWT && !notesDataList.some(s => s.textBlocks.length > 0)) {
        notesDataList = extractTextFromNotesContainers(pptStream, documentRecord, notesDataList);
    }

    // Map notes to slides by index (notes are in the same order as slides)
    const notesMap = new Map<number, SlideData>();
    for (let i = 0; i < notesDataList.length; i++) {
        notesMap.set(i, notesDataList[i]);
    }

    // 5. Build AST content nodes
    const content: OfficeContentNode[] = [];
    const notes: OfficeContentNode[] = [];

    for (const slideData of slideDataList) {
        const slideNumber = slideData.index + 1;
        const slideChildren: OfficeContentNode[] = [];

        for (const block of slideData.textBlocks) {
            // Split text by paragraph marks (\r in PPT)
            const paragraphs = block.text.split('\r');

            for (const para of paragraphs) {
                const cleanText = para.replace(/[\x00-\x06\x08-\x0C\x0E-\x1F]/g, '').trim();
                if (!cleanText) continue;

                // Title text types become headings
                const isTitle = block.textType === TEXT_TYPE_TITLE ||
                    block.textType === TEXT_TYPE_CENTER_TITLE;

                if (isTitle) {
                    slideChildren.push({
                        type: 'heading',
                        text: cleanText,
                        children: [{ type: 'text', text: cleanText }],
                        metadata: { level: 1 },
                    });
                } else {
                    slideChildren.push({
                        type: 'paragraph',
                        text: cleanText,
                        children: [{ type: 'text', text: cleanText }],
                    });
                }
            }
        }

        const slideText = slideChildren.map(c => c.text || '').filter(t => t).join(config.newlineDelimiter ?? '\n');

        const slideNode: OfficeContentNode = {
            type: 'slide',
            text: slideText,
            children: slideChildren,
            metadata: {
                slideNumber,
                noteId: notesMap.has(slideData.index) ? `slide-note-${slideNumber}` : undefined,
            } as SlideMetadata,
        };

        content.push(slideNode);

        // Process notes for this slide
        if (!config.ignoreNotes) {
            const noteData = notesMap.get(slideData.index);
            if (noteData) {
                const noteChildren: OfficeContentNode[] = [];
                for (const block of noteData.textBlocks) {
                    const paragraphs = block.text.split('\r');
                    for (const para of paragraphs) {
                        const cleanText = para.replace(/[\x00-\x06\x08-\x0C\x0E-\x1F]/g, '').trim();
                        if (!cleanText) continue;
                        noteChildren.push({
                            type: 'paragraph',
                            text: cleanText,
                            children: [{ type: 'text', text: cleanText }],
                        });
                    }
                }

                if (noteChildren.length > 0) {
                    const noteText = noteChildren.map(c => c.text || '').filter(t => t).join(config.newlineDelimiter ?? '\n');
                    const noteNode: OfficeContentNode = {
                        type: 'note',
                        text: noteText,
                        children: noteChildren,
                        metadata: {
                            noteId: `slide-note-${slideNumber}`,
                        },
                    };

                    if (config.putNotesAtLast) {
                        notes.push(noteNode);
                    } else {
                        content.push(noteNode);
                    }
                }
            }
        }
    }

    // Append notes at the end if configured
    if (config.putNotesAtLast && notes.length > 0) {
        content.push(...notes);
    }

    // 6. Build and return AST
    const delimiter = config.newlineDelimiter ?? '\n';

    return {
        type: 'ppt' as any,
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
// Fallback: Direct Text Scanning
// ============================================================================

/**
 * Fallback text extraction by scanning the entire stream for TextCharsAtom
 * and TextBytesAtom records. Used when the Document container can't be found.
 */
function extractTextByScanning(
    pptStream: Buffer,
    config: Required<OfficeParserConfig>
): OfficeParserAST {
    const content: OfficeContentNode[] = [];
    let pos = 0;
    let slideNumber = 1;
    let currentChildren: OfficeContentNode[] = [];
    let lastTextType = TEXT_TYPE_BODY;

    while (pos + 8 <= pptStream.length) {
        const rec = readRecordHeader(pptStream, pos);
        if (!rec) break;

        // Safety
        const nextPos = rec.dataOffset + rec.recLen;
        if (nextPos <= pos || nextPos > pptStream.length + 8) {
            pos += 8;
            continue;
        }

        switch (rec.recType) {
            case RT_SLIDE_PERSIST_ATOM: {
                // New slide boundary - flush current
                if (currentChildren.length > 0) {
                    const slideText = currentChildren.map(c => c.text || '').filter(t => t).join(config.newlineDelimiter ?? '\n');
                    content.push({
                        type: 'slide',
                        text: slideText,
                        children: currentChildren,
                        metadata: { slideNumber } as SlideMetadata,
                    });
                    slideNumber++;
                    currentChildren = [];
                }
                break;
            }

            case RT_TEXT_HEADER_ATOM: {
                if (rec.recLen >= 4 && rec.dataOffset + 4 <= pptStream.length) {
                    lastTextType = pptStream.readInt32LE(rec.dataOffset);
                }
                break;
            }

            case RT_TEXT_CHARS_ATOM: {
                const text = readTextCharsAtom(pptStream, rec);
                if (text) {
                    addTextParagraphs(currentChildren, text, lastTextType, config);
                }
                break;
            }

            case RT_TEXT_BYTES_ATOM: {
                const text = readTextBytesAtom(pptStream, rec);
                if (text) {
                    addTextParagraphs(currentChildren, text, lastTextType, config);
                }
                break;
            }
        }

        // For containers, step into them; for atoms, skip past
        if (isContainer(rec)) {
            pos = rec.dataOffset;  // Step into container
        } else {
            pos = nextPos;
        }
    }

    // Flush last slide
    if (currentChildren.length > 0) {
        const slideText = currentChildren.map(c => c.text || '').filter(t => t).join(config.newlineDelimiter ?? '\n');
        content.push({
            type: 'slide',
            text: slideText,
            children: currentChildren,
            metadata: { slideNumber } as SlideMetadata,
        });
    }

    // If no slides were created, put everything under one slide
    if (content.length === 0 && currentChildren.length > 0) {
        content.push({
            type: 'slide',
            text: currentChildren.map(c => c.text || '').join(config.newlineDelimiter ?? '\n'),
            children: currentChildren,
            metadata: { slideNumber: 1 } as SlideMetadata,
        });
    }

    const delimiter = config.newlineDelimiter ?? '\n';

    return {
        type: 'ppt' as any,
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

/**
 * Helper to add text paragraphs from a raw text block to a children array.
 */
function addTextParagraphs(
    children: OfficeContentNode[],
    text: string,
    textType: number,
    config: Required<OfficeParserConfig>
): void {
    const paragraphs = text.split('\r');
    const isTitle = textType === TEXT_TYPE_TITLE || textType === TEXT_TYPE_CENTER_TITLE;

    for (const para of paragraphs) {
        const cleanText = para.replace(/[\x00-\x06\x08-\x0C\x0E-\x1F]/g, '').trim();
        if (!cleanText) continue;

        if (isTitle) {
            children.push({
                type: 'heading',
                text: cleanText,
                children: [{ type: 'text', text: cleanText }],
                metadata: { level: 1 },
            });
        } else {
            children.push({
                type: 'paragraph',
                text: cleanText,
                children: [{ type: 'text', text: cleanText }],
            });
        }
    }
}

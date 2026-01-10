/**
 * RTF (Rich Text Format) Parser
 * 
 * **RTF Format Overview:**
 * RTF is a proprietary document format developed by Microsoft in 1987.
 * Unlike OOXML formats (DOCX), RTF is a plain text format using control words and groups.
 * 
 * **Basic RTF Structure:**
 * ```rtf
 * {\rtf1\ansi
 *   {\fonttbl{\f0 Arial;}{\f1 Times;}}
 *   {\colortbl;\red255\green0\blue0;\red0\green0\blue255;}
 *   \f0\fs24 This is \b bold\b0  and \i italic\i0  text.\par
 * }
 * ```
 * 
 * **RTF Elements:**
 * 1. **Control Words**: Backslash followed by letters and optional parameter (e.g., `\fs24`, `\b`, `\par`)
 * 2. **Control Symbols**: Backslash followed by single special char (e.g., `\'xx` for hex char, `\{`, `\}`, `\\`)
 * 3. **Groups**: Content enclosed in `{...}` - creates formatting scope
 * 4. **Text**: Plain text characters
 * 
 * **Key Control Words:**
 * - `\rtf1` - RTF version 1
 * - `\fs24` - Font size in half-points (24 = 12pt)
 * - `\b`,`\i`,`\ul`,`\strike` - Bold, italic, underline, strikethrough
 * - `\par` - Paragraph break
 * - `\f0` - Switch to font 0 from font table
 * - `\cf1` - Text color from color table
 * - `\cb2` - Background color from color table
 * - `\sub`,`\super` - Subscript, superscript
 * 
 * **Parser Architecture:**
 * This module uses a two-phase approach:
 * 1. **Lexical Phase** (`SimpleRtfParser`): Tokenizes RTF into groups, control words, and text
 * 2. **Semantic Phase** (`parseRtf`): Traverses the token tree to build the AST
 * 
 * @module RtfParser
 * @see https://www.biblioscape.com/rtf15_spec.htm RTF 1.5 Specification
 * @see https://latex2rtf.sourceforge.net/RTF-Spec-1.2.pdf RTF 1.2 Specification
 */

import { ImageMetadata, ListMetadata, NoteMetadata, OfficeAttachment, OfficeContentNode, OfficeMimeType, OfficeParserAST, OfficeParserConfig, TextFormatting } from '../types';
import { logWarning } from '../utils/errorUtils';
import { performOcr } from '../utils/ocrUtils';

/**
 * Represents an RTF group (content enclosed in braces).
 * Groups create formatting scopes and can contain other groups, control words, or text.
 * 
 * @example
 * // RTF: {\b bold text}
 * // Parsed as group containing control word 'b' and text 'bold text'
 */
export interface RtfGroup {
    /** Node type identifier */
    type: 'group';

    /** Contents of this group (can be nested groups, control words, or text) */
    content: (RtfGroup | RtfText | RtfControl)[];

    /**
     * The destination control word for this group (if any).
     * Destinations specify what the group contains (e.g., 'fonttbl' for font table).
     * Common destinations: 'fonttbl', 'colortbl', 'footer', 'header', 'pict', 'footnote'
     * @example "fonttbl", "colortbl", "footnote"
     */
    destination?: string;
}

/**
 * Represents plain text content in RTF.
 * Text nodes contain the actual displayable characters.
 */
export interface RtfText {
    /** Node type identifier */
    type: 'text';

    /** The text content (may be a single character or multiple characters) */
    value: string;
}

/**
 * Represents an RTF control word or control symbol.
 * Control words modify formatting, specify special characters, or provide document structure.
 * 
 * @example
 * // \fs24 - Control word "fs" with parameter 24
 * { type: 'control', value: 'fs', param: 24 }
 * 
 * // \b - Control word "b" with no parameter (defaults to "on")
 * { type: 'control', value: 'b', param: undefined }
 * 
 * // \u1234 - Unicode character U+1234
 * { type: 'control', value: 'u', param: 1234 }
 */
export interface RtfControl {
    /** Node type identifier */
    type: 'control';

    /**
     * The control word name (without backslash).
     * @example "fs" for font size, "b" for bold, "par" for paragraph
     */
    value: string;

    /**
     * Optional numeric parameter.
     * - For `\fs24`: param = 24 (12pt font)
     * - For `\b`: param = undefined (defaults to "on")
     * - For `\b0`: param = 0 (explicitly "off")
     * - For `\u1234`: param = 1234 (Unicode code point)
     */
    param?: number;
}

/**
 * Union type representing any RTF parse tree node.
 */
export type RtfNode = RtfGroup | RtfText | RtfControl;
/**
 * Represents every RTF-supported image format that this parser handles.
 * This ensures compile-time safety. Adding a new format forces all maps to be updated.
 */
type RtfImageFormat =
    | 'png'
    | 'jpeg'
    | 'gif'
    | 'tiff'
    | 'bmp';

/**
 * Lookup table mapping RTF control words to internal formats.
 * Fully typed: if a key maps to an unsupported format, TypeScript throws an error.
 */
const RTF_BLIP_MAP: Record<string, RtfImageFormat> = {
    // Raster formats
    pngblip: 'png',
    jpegblip: 'jpeg',
    gifblip: 'gif',
    tiffblip: 'tiff',
    dibitmap: 'bmp',
    wbitmap: 'bmp',
};

/**
 * Lookup table mapping internal formats to MIME types.
 * Again fully typed: if a format is missing from this map, TS errors.
 */
const IMAGE_MIME_MAP: Record<RtfImageFormat, OfficeMimeType> = {
    png: 'image/png',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    tiff: 'image/tiff',
    bmp: 'image/bmp',
};

/**
 * Low-level RTF tokenizer that parses RTF syntax into a tree structure.
 * 
 * This class performs lexical analysis on RTF content, breaking it down into:
 * - Groups (enclosed in braces)
 * - Control words (e.g., `\fs24`, `\b`)  
 * - Control symbols (e.g., `\'xx`, `\{`)
 * - Plain text
 * 
 * The parser uses a byte-level approach and maintains a stack to track nested groups.
 * 
 * @example
 * ```typescript
 * const buffer = Buffer.from('{\\rtf1 Hello \\b world\\b0}');
 * const parser = new SimpleRtfParser(buffer);
 * const tree = parser.parse();
 * // tree.content contains parsed RTF nodes
 * ```
 */
export class SimpleRtfParser {
    /** Current position in the buffer */
    private index: number = 0;

    /** The RTF content as a Buffer */
    private buffer: Buffer;

    /** Total length of the buffer */
    private length: number;

    /**
     * Creates a new RTF parser.
     * @param buffer - The RTF file content as a Buffer
     */
    constructor(buffer: Buffer) {
        this.buffer = buffer;
        this.length = buffer.length;
    }

    public parse(): RtfGroup {
        const root: RtfGroup = { type: 'group', content: [] };
        const stack: RtfGroup[] = [root];

        while (this.index < this.length) {
            const char = this.buffer[this.index];
            const currentGroup = stack[stack.length - 1];

            if (char === 0x7B) { // '{'
                this.index++;
                const newGroup: RtfGroup = { type: 'group', content: [] };
                currentGroup.content.push(newGroup);
                stack.push(newGroup);
            } else if (char === 0x7D) { // '}'
                this.index++;
                if (stack.length > 1) {
                    stack.pop();
                }
                // If stack is 1 (root), we ignore extra closing braces or just stop?
                // RTF should be balanced, but let's be robust.
            } else if (char === 0x5C) { // '\'
                this.index++;
                this.parseControl(currentGroup);
            } else if (char === 0x0D || char === 0x0A) { // CR or LF
                this.index++; // Ignore newlines in RTF source
            } else {
                this.parseText(currentGroup);
            }
        }

        return root;
    }

    private parseControl(group: RtfGroup) {
        if (this.index >= this.length) return;

        const char = this.buffer[this.index];

        // Special control symbols
        if (char === 0x7B || char === 0x7D || char === 0x5C) { // \{ \} \\
            group.content.push({ type: 'text', value: String.fromCharCode(char) });
            this.index++;
            return;
        }

        if (char === 0x27) { // \'xx (hex)
            this.index++;
            if (this.index + 1 < this.length) {
                const hex = this.buffer.toString('utf8', this.index, this.index + 2);
                const code = parseInt(hex, 16);
                if (!isNaN(code)) {
                    // RTF hex escapes represent bytes in the document's code page (usually Windows-1252)
                    // Characters 0x80-0x9F in Windows-1252 don't map directly to Unicode
                    // We need to convert them properly
                    const windows1252ToUnicode: { [key: number]: number } = {
                        0x80: 0x20AC, // €
                        0x82: 0x201A, // ‚
                        0x83: 0x0192, // ƒ
                        0x84: 0x201E, // „
                        0x85: 0x2026, // …
                        0x86: 0x2020, // †
                        0x87: 0x2021, // ‡
                        0x88: 0x02C6, // ˆ
                        0x89: 0x2030, // ‰
                        0x8A: 0x0160, // Š
                        0x8B: 0x2039, // ‹
                        0x8C: 0x0152, // Œ
                        0x8E: 0x017D, // Ž
                        0x91: 0x2018, // '
                        0x92: 0x2019, // '
                        0x93: 0x201C, // "
                        0x94: 0x201D, // "
                        0x95: 0x2022, // •
                        0x96: 0x2013, // –
                        0x97: 0x2014, // —
                        0x98: 0x02DC, // ˜
                        0x99: 0x2122, // ™
                        0x9A: 0x0161, // š
                        0x9B: 0x203A, // ›
                        0x9C: 0x0153, // œ
                        0x9E: 0x017E, // ž
                        0x9F: 0x0178  // Ÿ
                    };

                    const unicodeCode = windows1252ToUnicode[code] || code;
                    group.content.push({ type: 'text', value: String.fromCharCode(unicodeCode) });
                }
                this.index += 2;
            }
            return;
        }

        if (char === 0x2A) { // \* (ignorable destination)
            // We treat this as a control word named '*'
            group.content.push({ type: 'control', value: '*' });
            this.index++;
            return;
        }

        // Control word
        let name = '';
        while (this.index < this.length) {
            const c = this.buffer[this.index];
            if ((c >= 0x61 && c <= 0x7A) || (c >= 0x41 && c <= 0x5A)) { // a-z or A-Z
                name += String.fromCharCode(c);
                this.index++;
            } else {
                break;
            }
        }

        let param: number | undefined = undefined;
        let hasParam = false;
        let paramStr = '';

        // Check for parameter (digits, potentially negative)
        if (this.index < this.length && this.buffer[this.index] === 0x2D) { // -
            paramStr += '-';
            this.index++;
        }
        while (this.index < this.length) {
            const c = this.buffer[this.index];
            if (c >= 0x30 && c <= 0x39) { // 0-9
                paramStr += String.fromCharCode(c);
                this.index++;
                hasParam = true;
            } else {
                break;
            }
        }
        if (hasParam) {
            param = parseInt(paramStr, 10);
        }

        // Space after control word is consumed
        if (this.index < this.length && this.buffer[this.index] === 0x20) {
            this.index++;
        }

        // Handle \binN
        if (name === 'bin' && param !== undefined && param > 0) {
            // Skip N bytes of binary data
            this.index += param;
            // \binN is not added to content as we want to ignore it
            return;
        }

        group.content.push({ type: 'control', value: name, param });

        // If this is the first control word in the group, it might be the destination
        if (group.content.length === 1 && group.type === 'group') {
            group.destination = name;
        } else if (group.content.length === 2 && group.content[0].type === 'control' && group.content[0].value === '*') {
            // If first was *, second is destination
            group.destination = name;
        }
    }

    private parseText(group: RtfGroup) {
        let text = '';
        while (this.index < this.length) {
            const char = this.buffer[this.index];
            if (char === 0x7B || char === 0x7D || char === 0x5C || char === 0x0D || char === 0x0A) {
                break;
            }
            // Basic ASCII text.
            text += String.fromCharCode(char);
            this.index++;
        }
        if (text.length > 0) {
            group.content.push({ type: 'text', value: text });
        }
    }
}

/**
 * Parses an RTF file and returns the AST.
 * 
 * **RTF Format Limitations:**
 * The following features are NOT supported due to RTF format constraints:
 * 
 * 1. **Images/Attachments**: RTF `\pict` contains device-dependent metafile data
 *    (WMF/EMF/DIB). Extracting as portable images requires complex metafile parsing.
 *    
 * 2. **StyleMap**: RTF uses inline formatting rather than named style definitions.
 *    There is no direct equivalent to DOCX's style.xml.
 *    
 * **Note Support:**
 * RTF supports both footnotes and endnotes using the `\footnote` group.
 * The `\fet` control word distinguishes between them:
 * - `\fet0` = footnotes only (default)
 * - `\fet1` = endnotes only
 * - `\fet2` = both footnotes and endnotes
 * 
 * @param buffer The file buffer.
 * @param config The parser configuration.
 * @returns The parsed AST.
 */
export const parseRtf = async (buffer: Buffer, config: OfficeParserConfig): Promise<OfficeParserAST> => {
    try {
        const parser = new SimpleRtfParser(buffer);
        const doc = parser.parse();

        // Extract font and color tables
        const fontTable = extractFontTable(doc);
        const colorTable = extractColorTable(doc);

        const content: OfficeContentNode[] = [];
        const notes: OfficeContentNode[] = [];
        const attachments: OfficeAttachment[] = [];

        // State for paragraph construction
        let currentParagraphText = '';
        let currentParagraphChildren: OfficeContentNode[] = [];
        let currentParagraphRaw = '';

        // State for text run construction
        let currentRunText = '';
        let currentFormatting: TextFormatting = {};

        // Target for content (main body or notes)
        let currentTarget = content;

        // Paragraph-level state
        let paragraphIndent = 0;
        let paragraphAlignment: 'left' | 'center' | 'right' | 'justify' = 'left';
        let isListItem = false;
        let listType: 'ordered' | 'unordered' | undefined;
        let headingLevel: number | undefined;
        let currentListId: string | undefined;

        // Persistent list state for listtext/pntext detection
        // These persist across paragraphs to allow list items without explicit \ls
        let lastKnownListId: string | undefined;
        let lastKnownListType: 'ordered' | 'unordered' | undefined;

        // ═══════════════════════════════════════════════════════════════════
        // Table state tracking (Stack-based for nesting)
        // ═══════════════════════════════════════════════════════════════════
        interface TableContext {
            rows: OfficeContentNode[];
            currentCells: OfficeContentNode[]; // Cells in the current row
            currentCellContent: OfficeContentNode[]; // Content of the currently open cell
            rowIndex: number;
        }

        let tableStack: TableContext[] = [];
        let currentFootnoteId = 0;

        // ═══════════════════════════════════════════════════════════════════
        // Note type tracking (footnotes vs endnotes)
        // ═══════════════════════════════════════════════════════════════════
        // RTF uses \fet to distinguish note types:
        // \fet0 = footnotes only (default)
        // \fet1 = endnotes only
        // \fet2 = both footnotes and endnotes
        let fetValue = 0; // Default to footnotes only

        // Helper to get current table context
        const getCurrentTable = (): TableContext | undefined => tableStack.length > 0 ? tableStack[tableStack.length - 1] : undefined;

        // Helper to ensure a table context exists (for top-level tables)
        const ensureTableContext = () => {
            if (tableStack.length === 0) {
                tableStack.push({
                    rows: [],
                    currentCells: [],
                    currentCellContent: [],
                    rowIndex: 0
                });
            }
        };

        let inTable = false;
        let paragraphInTable = false;
        let tableId = 0;

        // Table cell properties tracking
        interface CellProps {
            isMergedContinuation: boolean;
        }
        let rowCellProps: CellProps[] = [];
        let currentCellDefinitionProps: CellProps = { isMergedContinuation: false };
        let cellContentIndex = 0;

        // ═══════════════════════════════════════════════════════════════════
        // List state tracking (Word 97+ uses \ls for list style ID)
        // ═══════════════════════════════════════════════════════════════════
        let listIdCounter = 0;
        const listStyleIdMap: { [key: number]: string } = {};

        // List definition state for parsing \listtable
        let parsingListTable = false;
        let parsingListDefinition = false;
        let currentDefinedListId: number | undefined;
        let currentDefinedListType: 'ordered' | 'unordered' | undefined;
        const listTypeMap: { [key: number]: 'ordered' | 'unordered' } = {};

        // List override state for parsing \listoverridetable
        let parsingListOverrideTable = false;
        let currentListOverrideListId: number | undefined;
        let currentListOverrideLs: number | undefined;
        const listOverrideMap: { [lsId: number]: number } = {}; // Maps \ls ID to \listid

        // List counters for itemIndex tracking
        // Map: listId -> indentation level -> count
        const listCounters: { [key: string]: { [key: number]: number } } = {};

        // ═══════════════════════════════════════════════════════════════════
        // Hyperlink state (RTF uses \field{\*\fldinst HYPERLINK "url"})
        // ═══════════════════════════════════════════════════════════════════
        let currentLinkUrl: string | undefined;

        // Helper to check if formatting changed
        const formattingChanged = (a: TextFormatting, b: TextFormatting): boolean => {
            return a.bold !== b.bold ||
                a.italic !== b.italic ||
                a.underline !== b.underline ||
                a.strikethrough !== b.strikethrough ||
                a.size !== b.size ||
                a.font !== b.font ||
                a.color !== b.color ||
                a.backgroundColor !== b.backgroundColor ||
                a.subscript !== b.subscript ||
                a.superscript !== b.superscript;
        };

        // Helper to flush current run to paragraph children
        const flushRun = () => {
            if (currentRunText) {
                const node: OfficeContentNode = {
                    type: 'text',
                    text: currentRunText,
                    formatting: { ...currentFormatting }
                };
                // Use TextMetadata.link for hyperlinks
                if (currentLinkUrl) {
                    node.metadata = {
                        link: currentLinkUrl,
                        linkType: classifyLinkType(currentLinkUrl)
                    };
                }
                currentParagraphChildren.push(node);
                currentParagraphText += currentRunText;
                currentRunText = '';
            }
        };

        let isFlushingTable = false;

        // Helper to flush current paragraph
        const flushParagraph = () => {
            flushRun(); // Ensure last run is added

            // Check if we need to end the table
            // If we were in a table, but this paragraph is NOT marked as in-table, 
            // and we have content, then the table has ended.
            const hasContent = currentParagraphText || currentParagraphChildren.length > 0;
            if (inTable && !paragraphInTable && !isFlushingTable && hasContent) {
                // CRITICAL: Save current paragraph content before flushing table
                // because flushTable() -> flushRow() -> flushCell() -> flushParagraph()
                // would otherwise process this content during the table flush
                const savedParagraphText = currentParagraphText;
                const savedParagraphChildren = [...currentParagraphChildren];
                const savedParagraphRaw = currentParagraphRaw;

                // Clear buffers so nested flushParagraph() doesn't process them
                currentParagraphText = '';
                currentParagraphChildren = [];
                currentParagraphRaw = '';

                flushTable();

                // Restore the saved content for processing after the table
                currentParagraphText = savedParagraphText;
                currentParagraphChildren = savedParagraphChildren;
                currentParagraphRaw = savedParagraphRaw;
            }

            if (hasContent) {
                let nodeType: string = 'paragraph';
                let metadata: any = undefined;

                if (headingLevel !== undefined && headingLevel > 0) {
                    nodeType = 'heading';
                    metadata = { level: headingLevel };
                    // Reset list context when we encounter a heading
                    lastKnownListId = undefined;
                    lastKnownListType = undefined;
                } else if (isListItem) {
                    nodeType = 'list';

                    // Use lastKnownListId if currentListId is not set
                    // (happens when list item is detected via listtext/pntext)
                    const effectiveListId = currentListId || lastKnownListId;
                    const effectiveListType = listType || lastKnownListType || 'unordered';

                    // Calculate itemIndex
                    let itemIndex = 0;
                    if (effectiveListId) {
                        if (!listCounters[effectiveListId]) {
                            listCounters[effectiveListId] = {};
                        }
                        if (listCounters[effectiveListId][paragraphIndent] === undefined) {
                            listCounters[effectiveListId][paragraphIndent] = 0;
                        } else {
                            listCounters[effectiveListId][paragraphIndent]++;
                        }
                        itemIndex = listCounters[effectiveListId][paragraphIndent];
                    }

                    metadata = {
                        listType: effectiveListType,
                        indentation: paragraphIndent,
                        listId: effectiveListId || '',
                        itemIndex: itemIndex,
                        alignment: paragraphAlignment,
                    } as ListMetadata;

                    // Save for next listtext detection
                    if (effectiveListId) {
                        lastKnownListId = effectiveListId;
                    }
                    if (effectiveListType) {
                        lastKnownListType = effectiveListType;
                    }
                }

                const node: OfficeContentNode = {
                    type: nodeType as any,
                    text: currentParagraphText,
                    children: currentParagraphChildren,
                    formatting: undefined,
                    metadata: metadata
                };

                if (config.includeRawContent && currentParagraphRaw) {
                    node.rawContent = currentParagraphRaw;
                }

                // If we're building a table, add to current cell 
                // but ONLY if this paragraph was actually marked as in-table
                if (inTable && paragraphInTable) {
                    ensureTableContext();
                    getCurrentTable()!.currentCellContent.push(node);
                } else {
                    currentTarget.push(node);
                }

                currentParagraphText = '';
                currentParagraphChildren = [];
                currentParagraphRaw = '';

                // Reset paragraph-level state
                paragraphIndent = 0;
                isListItem = false;
                listType = undefined;
                headingLevel = undefined;
                currentListId = undefined;
            }
        };

        // Helper to flush current cell
        const flushCell = (tableCtx?: TableContext) => {
            flushParagraph();
            const ctx = tableCtx || getCurrentTable();
            if (!ctx) return undefined;

            // Always return a cell node, even if empty, to preserve table structure (grid)
            const cellNode: OfficeContentNode = {
                type: 'cell',
                text: ctx.currentCellContent.map(c => c.text).join('\n'),
                children: [...ctx.currentCellContent],
                metadata: {
                    row: ctx.rowIndex,
                    col: ctx.currentCells.length
                }
            };
            ctx.currentCellContent = [];
            return cellNode;
        };

        // Helper to flush current row - creates a row node from collected cells
        const flushRow = (tableCtx?: TableContext) => {
            const ctx = tableCtx || getCurrentTable();
            if (!ctx) return;

            const cell = flushCell(ctx);
            // Only add cell if it has content (prevents phantom empty cells during cleanup)
            if (cell && (cell.children && cell.children.length > 0 || cell.text)) {
                ctx.currentCells.push(cell);
            }

            if (ctx.currentCells.length > 0) {
                const rowNode: OfficeContentNode = {
                    type: 'row',
                    text: ctx.currentCells.map(c => c.text).filter(t => t !== '').join(config.newlineDelimiter ?? '\n'),
                    children: [...ctx.currentCells]
                };
                ctx.rows.push(rowNode);
                ctx.currentCells = [];
                ctx.rowIndex++;
            }
        };

        // Helper to flush table
        const flushTable = () => {
            if (isFlushingTable) return;
            isFlushingTable = true;

            const ctx = getCurrentTable();
            if (!ctx) {
                isFlushingTable = false;
                return;
            }

            flushRow(ctx);

            if (ctx.rows.length > 0) {
                tableId++;
                const tableNode: OfficeContentNode = {
                    type: 'table',
                    text: ctx.rows.map(r => r.text).join('\n'), // Aggregate text from rows
                    children: [...ctx.rows]
                };

                // If we have a parent table, add this table to the parent's current cell
                if (tableStack.length > 1) {
                    const parentCtx = tableStack[tableStack.length - 2];
                    parentCtx.currentCellContent.push(tableNode);
                } else {
                    currentTarget.push(tableNode);
                }
            }

            // Pop the table from stack
            tableStack.pop();

            // If stack is empty, we are out of table mode
            if (tableStack.length === 0) {
                inTable = false;
                paragraphInTable = false;
            }

            isFlushingTable = false;
        };

        // Extract hyperlink URL from field instruction group
        // Recursively searches for HYPERLINK "url" pattern in nested groups
        const extractHyperlinkUrl = (group: RtfGroup): string | undefined => {
            let url: string | undefined;

            // Helper to recursively find hyperlink URL
            const findUrl = (node: RtfGroup | RtfText | RtfControl): string | undefined => {
                if (node.type === 'text') {
                    // Check for HYPERLINK "url" pattern
                    const text = node.value;
                    const match = text.match(/HYPERLINK\s+"([^"]+)"/i);
                    if (match) {
                        return match[1];
                    }
                    // Also check for URL after HYPERLINK on same or separate text node
                    const urlOnlyMatch = text.match(/"(https?:\/\/[^"]+|mailto:[^"]+|#[^"]+)"/);
                    if (urlOnlyMatch) {
                        return urlOnlyMatch[1];
                    }
                } else if (node.type === 'group') {
                    // Check if this is a fldinst group (contains field instruction)
                    let foundHyperlink = false;
                    let foundUrl: string | undefined;

                    for (const child of node.content) {
                        if (child.type === 'text') {
                            const text = child.value.toUpperCase();
                            if (text.includes('HYPERLINK')) {
                                foundHyperlink = true;
                            }
                            // Look for quoted URL
                            const urlMatch = child.value.match(/"([^"]+)"/);
                            if (urlMatch && foundHyperlink) {
                                foundUrl = urlMatch[1];
                            }
                        } else if (child.type === 'group') {
                            const nestedUrl = findUrl(child);
                            if (nestedUrl) {
                                foundUrl = nestedUrl;
                            }
                        }
                    }
                    if (foundUrl) return foundUrl;
                }
                return undefined;
            };

            // Search the field group for fldinst
            for (const child of group.content) {
                if (child.type === 'group') {
                    const foundUrl = findUrl(child);
                    if (foundUrl) return foundUrl;
                }
            }

            return url;
        };

        /**
         * Determines whether a hyperlink URL is internal or external.
         * Internal = bookmark references (no scheme or starts with "#").
         * External = any scheme like http, https, mailto, ftp, file, etc.
         */
        const classifyLinkType = (url: string): 'internal' | 'external' => {
            // Trim whitespace
            const clean = url.trim();

            // Internal pattern 1: starts with "#"
            if (clean.startsWith('#')) {
                return 'internal';
            }

            // Internal pattern 2: no scheme at all (pure bookmark)
            // Detect schemes by checking "something:" prefix
            if (!/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(clean)) {
                return 'internal';
            }

            // Everything else is external
            return 'external';
        };

        // Helper to extract text content from a group (for list marker detection)
        // Recursively collects all text content from a group
        const extractTextFromGroup = (group: RtfGroup): string => {
            let text = '';

            const collectText = (node: RtfNode): void => {
                if (node.type === 'text') {
                    text += node.value;
                } else if (node.type === 'group') {
                    for (const child of node.content) {
                        collectText(child);
                    }
                }
            };

            for (const child of group.content) {
                collectText(child);
            }

            return text;
        };

        /**
         * Extracts an image attachment from an RTF \pict group.
         * Uses lookup tables for clean and safe handling of all formats.
         *
         * @param pictGroup The RtfGroup node that represents a \pict group.
         * @returns An OfficeAttachment or undefined when unsupported or invalid.
         */
        const extractPictAttachment = (pictGroup: RtfGroup): OfficeAttachment | undefined => {
            /** Internal format detected from the RTF pict group */
            let imageFormat: RtfImageFormat | undefined;

            /** Hexadecimal string extracted from the pict binary section */
            let hexData = '';

            // -------------------------------------------------------------
            // Walk through pict group content to detect the blip type and gather hex data
            // -------------------------------------------------------------
            for (const child of pictGroup.content) {
                // If the node is a control word, we try to resolve it from lookup map
                if (child.type === 'control') {
                    // Lookup directly instead of if/else
                    const mapped = RTF_BLIP_MAP[child.value];
                    if (mapped) {
                        imageFormat = mapped;
                    }
                }
                else if (child.type === 'text') {
                    // Append only valid hex characters
                    hexData += child.value.replace(/[^0-9a-fA-F]/g, '');
                }
                else if (child.type === 'group') {
                    // Skip nested structures like picprop or blipuid
                }
            }

            // Missing or unknown format → stop
            if (!imageFormat || hexData.length === 0) {
                return undefined;
            }

            // Map internal format to MIME
            const mimeType = IMAGE_MIME_MAP[imageFormat];
            if (!mimeType) {
                return undefined;
            }

            // -------------------------------------------------------------
            // Convert hex → binary and construct the attachment object
            // -------------------------------------------------------------
            try {
                // Convert hex into a raw buffer
                const buffer = Buffer.from(hexData, 'hex');

                // Derive file extension directly from format
                const extension = imageFormat;

                // Generate a stable incremental filename
                const name = `image_${attachments.length + 1}.${extension}`;

                // Build and return final attachment
                return {
                    type: 'image',
                    mimeType: mimeType,
                    data: buffer.toString('base64'),
                    name: name,
                    extension: extension
                };
            }
            catch {
                // If conversion fails, ignore this image
                return undefined;
            }
        };

        // Helper to serialize RTF control word
        const serializeRtfControl = (node: RtfControl): string => {
            // Symbol control words (non-alpha)
            if (!/^[a-zA-Z]/.test(node.value)) {
                return `\\${node.value}`;
            }
            // Alpha control words
            let res = `\\${node.value}`;
            if (node.param !== undefined) {
                res += node.param;
            }
            // Add space delimiter for safety
            res += ' ';
            return res;
        };

        // Helper to serialize RTF text
        const serializeRtfText = (node: RtfText): string => {
            // Escape special characters: \, {, }
            return node.value.replace(/([\\{}])/g, '\\$1');
        };

        // Recursive function to traverse the RTF tree
        const traverse = (node: RtfNode, formatting: TextFormatting, depth: number = 0) => {
            if (node.type === 'group') {
                const ignoreList = [
                    'fonttbl', 'colortbl', 'stylesheet', 'info', 'macpict',
                    'pmmetafile', 'wmetafile', 'dibitmap', 'bitmap', 'object',
                    'nextGenerator', 'header', 'footer', 'nonshppict', 'xml', 'private',
                    'upnp', 'ud', 'filetbl', 'operator', 'author', 'creatim', 'revtim', 'printim', 'comment',
                    'fldinst', 'listtext', 'pntext' // Ignore list marker text (handled separately)
                ];

                let isIgnored = false;
                let isFootnote = false;
                let isHyperlinkField = false;
                let isPict = false;

                // Add group start to raw content
                // Note: We don't add ignored groups to rawContent to keep it clean
                // But we might want to if we want full fidelity. 
                // For now, let's include everything in rawContent except truly skipped stuff?
                // The user asked for "raw content which is probably the rtf group".
                // If we skip 'fonttbl', it's fine as it's not part of the content.
                // But 'listtext' IS part of the content structure even if we parse it separately.
                // Let's stick to the plan: if ignored, we might skip it in rawContent too, 
                // OR we include it. 
                // If I include it, `currentParagraphRaw` might get huge with font tables if they were inside the paragraph (unlikely).
                // Usually font tables are at document root.
                // `traverse` is called on `doc`.
                // `currentParagraphRaw` is reset on `flushParagraph`.
                // So if we are at root level, `currentParagraphRaw` accumulates everything until the first paragraph ends.
                // This might include the header/fonttbl if they are before the first \par.
                // That seems correct for "raw content" of the first node?
                // Actually, `fonttbl` is usually before any text.
                // If we include it, the first paragraph node will contain the entire font table in its rawContent.
                // That might be annoying.
                // Let's ONLY add to `currentParagraphRaw` if NOT ignored.

                if (node.destination) {
                    if (node.destination === 'footnote') {
                        isFootnote = true;
                    } else if (node.destination === 'field') {
                        // Check if this is a hyperlink field
                        const url = extractHyperlinkUrl(node);
                        if (url) {
                            // Flush any pending text before starting the link context
                            // This prevents previous text from inheriting the link
                            flushRun();
                            isHyperlinkField = true;
                            currentLinkUrl = url;
                        }
                    } else if (node.destination === 'listtable') {
                        // We want to parse list definitions
                        parsingListTable = true;
                    } else if (parsingListTable && node.destination === 'list') {
                        parsingListDefinition = true;
                        currentDefinedListId = undefined;
                        currentDefinedListType = undefined;
                    } else if (node.destination === 'listoverridetable') {
                        parsingListOverrideTable = true;
                    } else if (parsingListOverrideTable && node.destination === 'listoverride') {
                        // Reset per override group
                        currentListOverrideListId = undefined;
                        currentListOverrideLs = undefined;
                    } else if (node.destination === 'pict') {
                        // Handle picture extraction
                        if (config.extractAttachments) {
                            isPict = true;
                        } else {
                            isIgnored = true;
                        }
                    } else if (ignoreList.includes(node.destination)) {
                        isIgnored = true;
                    } else if (node.content.length > 0 && node.content[0].type === 'control' && node.content[0].value === '*') {
                        // Ignorable destination, but allow certain ones for:
                        // - fldinst: hyperlinks
                        // - nesttableprops: nested tables
                        // - shppict: shape pictures (contain pict groups)
                        const allowedIgnorable = ['fldinst', 'nesttableprops'];
                        if (config.extractAttachments) {
                            allowedIgnorable.push('shppict', 'listpicture');
                        }
                        if (!allowedIgnorable.includes(node.destination || '')) {
                            isIgnored = true;
                        }
                    }
                }

                // ═══════════════════════════════════════════════════════════
                // Handle listtext and pntext: These indicate the current paragraph
                // is a list item. We extract list type info before ignoring content.
                // ═══════════════════════════════════════════════════════════
                if (node.destination === 'listtext' || node.destination === 'pntext') {
                    // If this is the first list item, reset the indent
                    if (!isListItem)
                        paragraphIndent = 0;
                    isListItem = true;

                    // Try to determine list type from the marker content
                    // Bullets (unordered): '·', '•', 'o', '§', etc.
                    // Numbers (ordered): '1.', '2.', 'i.', 'ii.', 'a.', 'A.', etc.
                    const markerText = extractTextFromGroup(node);
                    if (markerText) {
                        const trimmed = markerText.trim();
                        // Check for common bullet characters
                        const bulletChars = ['·', '•', 'o', '§', '■', '□', '●', '○', '◆', '◇', '►', '▸', '\u00b7', '\u2022', '\u25cf', '\u25cb'];
                        const isBullet = bulletChars.some(b => trimmed.includes(b)) ||
                            // Font symbol bullets often use characters from Symbol font
                            (trimmed.length === 1 && !/[0-9a-zA-Z]/.test(trimmed));

                        if (isBullet) {
                            listType = 'unordered';
                        } else if (/^[0-9ivxlcdm]+[\.\)]/i.test(trimmed) || /^[a-z][\.\)]/i.test(trimmed)) {
                            // Matches: 1., 2), i., ii., a., A), etc.
                            listType = 'ordered';
                        }
                        // If we can't determine, leave listType as is (might be set by \ls/\levelnfc)
                    }

                    // Still mark as ignored to skip the marker text content
                    isIgnored = true;
                }

                if (isIgnored) return;

                // Append group start to raw content
                currentParagraphRaw += '{';

                // Handle pict group: extract image and add to content tree
                if (isPict) {
                    const attachment = extractPictAttachment(node);
                    if (attachment) {
                        attachments.push(attachment);
                        // Only add image node to content if this is NOT a list definition picture
                        // List pictures (bullets) should not appear in content, only as attachments
                        if (!parsingListTable && !parsingListDefinition) {
                            // Also add an image node to the content tree (like DOCX)
                            flushParagraph();
                            currentTarget.push({
                                type: 'image',
                                text: '',
                                metadata: {
                                    attachmentName: attachment.name || `image_${attachments.length}`
                                }
                            });
                        }
                    }
                    // We still traverse pict content to reconstruct raw RTF?
                    // No, extractPictAttachment consumes it.
                    // But we want it in rawContent?
                    // If we return here, we miss the closing '}'.
                    // And we miss the content in rawContent.
                    // Let's traverse it purely for rawContent if needed, but `extractPictAttachment` doesn't modify the tree.
                    // But `extractPictAttachment` does not return the raw string.
                    // So we should probably continue traversal but suppress text extraction?
                    // The original code returned here: `return; // Don't traverse pict content as text`
                    // So we should do the same, but we need to append the content to `currentParagraphRaw`.
                    // We can manually serialize the group content here.
                    for (const child of node.content) {
                        if (child.type === 'control') currentParagraphRaw += serializeRtfControl(child);
                        else if (child.type === 'text') currentParagraphRaw += serializeRtfText(child);
                        else if (child.type === 'group') {
                            // Recursive serialization for nested groups in pict (e.g. blipuid)
                            // We can't easily recurse `traverse` because it has side effects (text extraction).
                            // We need a pure serializer or just let `traverse` run but with a flag?
                            // Or just ignore the raw content of the image binary data?
                            // Image binary data can be huge.
                            // Maybe we shouldn't include the full hex dump in `rawContent`?
                            // The user said "raw content which is probably the rtf group".
                            // Including 5MB of hex data in the JSON AST might be bad.
                            // But for consistency, it is the raw content.
                            // Let's include it for now.
                            // To do this without side effects, we need a separate serialize function?
                            // Or just call traverse and ensure `isPict` logic prevents text extraction.
                            // Wait, `isPict` is true for this node.
                            // If we recurse, `isPict` will be false for children (unless they are also pict).
                            // But we want to suppress text extraction for children of pict.
                            // The original code did `return`.
                            // So we should manually serialize children here.
                            // Let's define a simple recursive serializer.
                            const serializeGroupContent = (g: RtfGroup) => {
                                for (const c of g.content) {
                                    if (c.type === 'control') currentParagraphRaw += serializeRtfControl(c);
                                    else if (c.type === 'text') currentParagraphRaw += serializeRtfText(c);
                                    else if (c.type === 'group') {
                                        currentParagraphRaw += '{';
                                        serializeGroupContent(c);
                                        currentParagraphRaw += '}';
                                    }
                                }
                            };
                            serializeGroupContent(node);
                        }
                    }
                    currentParagraphRaw += '}';
                    return;
                }

                // Handle footnote: switch target to notes
                const previousTarget = currentTarget;
                if (isFootnote) {
                    if (config.ignoreNotes) {
                        return; // Skip footnote content entirely
                    }
                    flushParagraph();
                    currentFootnoteId++;

                    // Determine note type based on \fet value
                    let noteType: 'footnote' | 'endnote' = 'footnote';
                    if (fetValue === 1) {
                        // \fet1 means all notes are endnotes
                        noteType = 'endnote';
                    } else if (fetValue === 2) {
                        // \fet2 means both types exist
                        // Check for \ftnalt marker to distinguish endnotes from footnotes
                        // \footnote\ftnalt indicates an endnote
                        const hasFtnalt = node.content.some(child =>
                            child.type === 'control' && child.value === 'ftnalt'
                        );
                        noteType = hasFtnalt ? 'endnote' : 'footnote';
                    }
                    // fetValue === 0 (default) means footnotes only

                    const noteNode: OfficeContentNode = {
                        type: 'note',
                        children: [],
                        metadata: {
                            noteId: currentFootnoteId.toString(),
                            noteType: noteType
                        } as NoteMetadata
                    };
                    notes.push(noteNode);
                    currentTarget = noteNode.children!;
                }

                // Create a new formatting context for the group
                const groupFormatting = { ...formatting };

                for (const child of node.content) {
                    // Skip fldinst groups (we already extracted the URL)
                    if (child.type === 'group' && child.destination === 'fldinst') {
                        // We still want it in rawContent!
                        // So we should traverse it but suppress text extraction?
                        // Or just serialize it?
                        // `fldinst` contains the URL.
                        // If we skip it in `traverse`, we miss it in `rawContent`.
                        // Let's traverse it but maybe the `fldinst` logic inside `traverse` handles it?
                        // The original code:
                        // if (child.type === 'group' && child.destination === 'fldinst') { continue; }
                        // This skips the child entirely.
                        // So we need to manually serialize it if we want it in rawContent.
                        currentParagraphRaw += '{';
                        // We need to serialize the content of fldinst
                        const serializeGroupContent = (g: RtfGroup) => {
                            for (const c of g.content) {
                                if (c.type === 'control') currentParagraphRaw += serializeRtfControl(c);
                                else if (c.type === 'text') currentParagraphRaw += serializeRtfText(c);
                                else if (c.type === 'group') {
                                    currentParagraphRaw += '{';
                                    serializeGroupContent(c);
                                    currentParagraphRaw += '}';
                                }
                            }
                        };
                        serializeGroupContent(child);
                        currentParagraphRaw += '}';
                        continue;
                    }
                    traverse(child, groupFormatting, depth + 1);
                }

                if (node.destination === 'listtable') {
                    parsingListTable = false;
                } else if (node.destination === 'list') {
                    parsingListDefinition = false;
                    if (currentDefinedListId !== undefined && currentDefinedListType !== undefined) {
                        listTypeMap[currentDefinedListId] = currentDefinedListType;
                    }
                } else if (node.destination === 'listoverridetable') {
                    parsingListOverrideTable = false;
                } else if (parsingListOverrideTable && node.destination === 'listoverride') {
                    // End of listoverride group - populate map
                    if (currentListOverrideLs !== undefined && currentListOverrideListId !== undefined) {
                        listOverrideMap[currentListOverrideLs] = currentListOverrideListId;
                    }
                }

                if (isFootnote) {
                    flushParagraph();
                    currentTarget = previousTarget;
                }

                // Clear link URL after processing the field group
                if (isHyperlinkField) {
                    flushRun();
                    currentLinkUrl = undefined;
                }

                // Append group end to raw content
                currentParagraphRaw += '}';

            } else if (node.type === 'text') {
                if (parsingListTable || parsingListOverrideTable) {
                    // Even if we don't extract text, we might want it in rawContent?
                    // Yes, rawContent should reflect the source.
                    currentParagraphRaw += serializeRtfText(node);
                    return;
                }

                if (formattingChanged(currentFormatting, formatting)) {
                    flushRun();
                    currentFormatting = { ...formatting };
                }
                currentRunText += node.value;
                currentParagraphRaw += serializeRtfText(node);
            } else if (node.type === 'control') {
                // Append control to raw content
                currentParagraphRaw += serializeRtfControl(node);

                // Handle list definition control words
                if (parsingListTable) {
                    if (node.value === 'listid') {
                        currentDefinedListId = node.param;
                    } else if (node.value === 'levelnfc' || node.value === 'levelnfcn') {
                        // 0 = Arabic, 1 = Upper Roman, 2 = Lower Roman, 3 = Upper Alpha, 4 = Lower Alpha -> Ordered
                        // 23 = Bullet, 255 = None -> Unordered
                        const isOrdered = node.param !== undefined && (node.param === 0 || (node.param >= 0 && node.param <= 4));
                        // Only set if not already set (or prioritize ordered if mixed?)
                        // We'll assume if any level is ordered, it's ordered.
                        // Or if we haven't set it yet.
                        if (!currentDefinedListType || (currentDefinedListType === 'unordered' && isOrdered)) {
                            currentDefinedListType = isOrdered ? 'ordered' : 'unordered';
                        }
                    }
                    return;
                }

                // Handle list override control words
                if (parsingListOverrideTable) {
                    if (node.value === 'listid') {
                        currentListOverrideListId = node.param;
                    } else if (node.value === 'ls') {
                        currentListOverrideLs = node.param;
                    }
                    return;
                }
                // Paragraph control words
                if (node.value === 'par') {
                    flushParagraph();
                    currentFormatting = { ...formatting };
                }
                // Table control words
                else if (node.value === 'trowd') {
                    // Table row definition - start of a new row
                    // Check if we are starting a nested table
                    // If we are already in a table, and we have content in the current cell, 
                    // then this trowd implies a nested table start.
                    const ctx = getCurrentTable();
                    if (inTable && ctx && ctx.currentCellContent.length > 0) {
                        // Start nested table
                        ensureTableContext(); // Should already exist if inTable is true
                        // Push new table context
                        tableStack.push({
                            rows: [],
                            currentCells: [],
                            currentCellContent: [],
                            rowIndex: 0
                        });
                    } else {
                        if (!inTable) {
                            inTable = true;
                            ensureTableContext();
                        }
                    }

                    // After \trowd we are inside a table row, so content should go to table cells.
                    // Many RTF files don't use \intbl, relying solely on \trowd...\cell...\row structure.
                    paragraphInTable = true;

                    // Reset cell properties for the new row definition
                    rowCellProps = [];
                    currentCellDefinitionProps = { isMergedContinuation: false };
                    cellContentIndex = 0;
                } else if (node.value === 'clvmrg') {
                    // Vertical merge continuation
                    currentCellDefinitionProps.isMergedContinuation = true;
                } else if (node.value === 'clmgf') {
                    // Vertical merge first cell (reset continuation flag if set, though usually mutually exclusive)
                    currentCellDefinitionProps.isMergedContinuation = false;
                } else if (node.value === 'cellx') {
                    // End of cell definition
                    rowCellProps.push({ ...currentCellDefinitionProps });
                    // Reset for next cell
                    currentCellDefinitionProps = { isMergedContinuation: false };
                } else if (node.value === 'cell') {
                    // End of cell - add it to current row
                    // Force paragraphInTable = true because \cell implies we are in a table cell
                    paragraphInTable = true;

                    // Check if this cell is a merged continuation
                    let isMergedContinuation = false;
                    if (cellContentIndex < rowCellProps.length) {
                        isMergedContinuation = rowCellProps[cellContentIndex].isMergedContinuation;
                    }
                    cellContentIndex++;

                    const cell = flushCell();
                    // Only add if not a merged continuation
                    if (cell) {
                        if (!isMergedContinuation) {
                            const ctx = getCurrentTable();
                            if (ctx) ctx.currentCells.push(cell);
                        }
                    }
                    currentFormatting = { ...formatting };
                } else if (node.value === 'nestcell') {
                    // End of cell in outer table (nested context)
                    // If we are in an inner table, we need to close it and return to outer

                    // First, flush the current cell of the inner table (if any pending)
                    // Actually, nestcell ends the OUTER cell.
                    // So the inner table should have been finished by now?
                    // Usually inner table ends with \row.

                    // If we are in a nested table (stack > 1), we should pop until we are at the outer table?
                    // Or maybe just pop one level?
                    if (tableStack.length > 1) {
                        // Flush the inner table if it has pending rows
                        const innerCtx = getCurrentTable();
                        if (innerCtx && (innerCtx.rows.length > 0 || innerCtx.currentCells.length > 0)) {
                            flushTable(); // This pops the stack
                        }
                    }

                    // Now we are (hopefully) at the outer table level
                    // Treat as a regular cell end for the outer table
                    paragraphInTable = true;
                    const cell = flushCell();
                    if (cell) {
                        const ctx = getCurrentTable();
                        if (ctx) ctx.currentCells.push(cell);
                    }
                    currentFormatting = { ...formatting };

                } else if (node.value === 'row') {
                    // End of row
                    flushRow();
                    currentFormatting = { ...formatting };
                    // Reset content index for safety (though trowd usually does it)
                    cellContentIndex = 0;
                    // Critical: Reset paragraphInTable after row ends.
                    // Subsequent paragraphs must explicitly use \intbl to be part of the table.
                    // Without this, content after the last \row gets incorrectly merged.
                    paragraphInTable = false;
                } else if (node.value === 'nestrow') {
                    // End of row in outer table
                    // If we are still in inner table context, flush it
                    if (tableStack.length > 1) {
                        flushTable();
                    }

                    flushRow();
                    currentFormatting = { ...formatting };
                    cellContentIndex = 0;
                } else if (node.value === 'intbl') {
                    // Paragraph is in a table
                    inTable = true;
                    paragraphInTable = true;
                    ensureTableContext();
                } else if (node.value === 'pard') {
                    // Reset paragraph properties
                    paragraphInTable = false;
                    // Reset other props...
                    paragraphIndent = 0;
                    paragraphAlignment = 'left';
                    isListItem = false;
                    listType = undefined;
                    headingLevel = undefined;
                    currentListId = undefined;
                    // Reset paragraph-level background (cbpat) to prevent leaking to next paragraph
                    formatting.backgroundColor = undefined;
                }
                // Text flow control
                else if (node.value === 'tab') {
                    if (formattingChanged(currentFormatting, formatting)) {
                        flushRun();
                        currentFormatting = { ...formatting };
                    }
                    currentRunText += '\t';
                } else if (node.value === 'line') {
                    if (formattingChanged(currentFormatting, formatting)) {
                        flushRun();
                        currentFormatting = { ...formatting };
                    }
                    currentRunText += '\n';
                }
                // Quote characters
                else if (node.value === 'lquote') {
                    // Left single quotation mark (U+2018)
                    if (formattingChanged(currentFormatting, formatting)) {
                        flushRun();
                        currentFormatting = { ...formatting };
                    }
                    currentRunText += '\u2018';
                } else if (node.value === 'rquote') {
                    // Right single quotation mark (U+2019)
                    if (formattingChanged(currentFormatting, formatting)) {
                        flushRun();
                        currentFormatting = { ...formatting };
                    }
                    currentRunText += '\u2019';
                } else if (node.value === 'ldblquote') {
                    // Left double quotation mark (U+201C)
                    if (formattingChanged(currentFormatting, formatting)) {
                        flushRun();
                        currentFormatting = { ...formatting };
                    }
                    currentRunText += '\u201C';
                } else if (node.value === 'rdblquote') {
                    // Right double quotation mark (U+201D)
                    if (formattingChanged(currentFormatting, formatting)) {
                        flushRun();
                        currentFormatting = { ...formatting };
                    }
                    currentRunText += '\u201D';
                }
                // Unicode character
                else if (node.value === 'u') {
                    if (node.param !== undefined) {
                        let code = node.param;
                        if (code < 0) code += 65536;
                        if (formattingChanged(currentFormatting, formatting)) {
                            flushRun();
                            currentFormatting = { ...formatting };
                        }
                        currentRunText += String.fromCharCode(code);
                    }
                }
                // Character formatting
                else if (node.value === 'b') {
                    formatting.bold = (node.param !== 0);
                } else if (node.value === 'i') {
                    formatting.italic = (node.param !== 0);
                } else if (node.value === 'ul') {
                    formatting.underline = (node.param !== 0);
                } else if (node.value === 'ulnone') {
                    formatting.underline = false;
                } else if (node.value === 'strike') {
                    formatting.strikethrough = (node.param !== 0);
                } else if (node.value === 'plain') {
                    // Reset all character formatting
                    formatting.bold = false;
                    formatting.italic = false;
                    formatting.underline = false;
                    formatting.strikethrough = false;
                    formatting.subscript = false;
                    formatting.superscript = false;
                    formatting.size = undefined;
                    formatting.font = undefined;
                    formatting.color = undefined;
                    formatting.backgroundColor = undefined;
                }
                // Font size (\fs - in half-points)
                else if (node.value === 'fs') {
                    if (node.param !== undefined) {
                        formatting.size = (node.param / 2).toString() + 'pt';
                    }
                }
                // Font family (\f)
                else if (node.value === 'f') {
                    if (node.param !== undefined && fontTable[node.param]) {
                        formatting.font = fontTable[node.param];
                    }
                }
                // Text color (\cf)
                else if (node.value === 'cf') {
                    if (node.param !== undefined && colorTable[node.param]) {
                        formatting.color = colorTable[node.param];
                    }
                }
                // Note type (\fet)
                else if (node.value === 'fet') {
                    // \fet0 = footnotes only (default)
                    // \fet1 = endnotes only
                    // \fet2 = both footnotes and endnotes
                    if (node.param !== undefined) {
                        fetValue = node.param;
                    }
                }
                // Background/highlight color (\cb, \highlight, \chcbpat, \cbpat)
                // \chcbpat = character background pattern color (used for shading)
                // \cbpat = paragraph background pattern color
                else if (node.value === 'cb' || node.value === 'highlight' || node.value === 'chcbpat' || node.value === 'cbpat') {
                    if (node.param !== undefined && colorTable[node.param]) {
                        formatting.backgroundColor = colorTable[node.param];
                    }
                }
                // Subscript
                else if (node.value === 'sub') {
                    formatting.subscript = true;
                    formatting.superscript = false;
                }
                // Superscript
                else if (node.value === 'super') {
                    formatting.superscript = true;
                    formatting.subscript = false;
                }
                // No subscript/superscript
                else if (node.value === 'nosupersub') {
                    formatting.subscript = false;
                    formatting.superscript = false;
                }
                // ═══════════════════════════════════════════════════════════
                // List control words
                // ═══════════════════════════════════════════════════════════
                // Paragraph indentation (\li - left indent in twips)
                else if (node.value === 'li') {
                    if (node.param !== undefined) {
                        // Convert twips to a simpler unit (720 twips = 1 inch, ~0.5 inch per level)
                        paragraphIndent = Math.floor(node.param / 360);
                    }
                }
                // List style ID (Word 97+)
                else if (node.value === 'ls') {
                    if (node.param !== undefined) {
                        // If this is the first list item, reset the indent
                        if (!isListItem)
                            paragraphIndent = 0;
                        isListItem = true;
                        // Generate or retrieve list ID
                        if (!listStyleIdMap[node.param]) {
                            listIdCounter++;
                            listStyleIdMap[node.param] = `rtf-list-${listIdCounter}`;
                        }
                        currentListId = listStyleIdMap[node.param];

                        // Look up type from list definition
                        // First check override map to get real list ID
                        const realListId = listOverrideMap[node.param] !== undefined ? listOverrideMap[node.param] : node.param;

                        if (listTypeMap[realListId]) {
                            listType = listTypeMap[realListId];
                        }
                    }
                }
                // List indent level (Word 97+)
                else if (node.value === 'ilvl') {
                    if (node.param !== undefined) {
                        isListItem = true;
                        paragraphIndent = node.param;
                    }
                }
                // List numbering level (\pnlvl)
                else if (node.value === 'pnlvl') {
                    isListItem = true;
                    if (node.param !== undefined) {
                        paragraphIndent = node.param;
                    }
                }
                // List numbering format
                else if (node.value === 'levelnfc' || node.value === 'pnf') {
                    // 0 = Arabic (1, 2, 3), 1 = Roman upper, 2 = Roman lower, 
                    // 3 = Letter upper, 4 = Letter lower, 23 = Bullet
                    if (node.param !== undefined) {
                        isListItem = true;
                        if (node.param === 23) {
                            listType = 'unordered';
                        } else {
                            listType = 'ordered';
                        }
                    }
                }
                // Ordered list indicator
                else if (node.value === 'pndec' || node.value === 'pnord' || node.value === 'pnlcltr' || node.value === 'pnucltr') {
                    // If this is the first list item, reset the indent
                    if (!isListItem)
                        paragraphIndent = 0; isListItem = true;
                    listType = 'ordered';
                }
                // Unordered list indicator
                else if (node.value === 'pnbullet' || node.value === 'pncard') {
                    // If this is the first list item, reset the indent
                    if (!isListItem)
                        paragraphIndent = 0;
                    isListItem = true;
                    listType = 'unordered';
                }
                // Style-based heading detection (\s)
                else if (node.value === 's') {
                    if (node.param !== undefined) {
                        // Common heading styles: s1-s9 (though this varies by document)
                        if (node.param >= 1 && node.param <= 9) {
                            headingLevel = node.param;
                        }
                    }
                }
                // Paragraph alignment
                else if (node.value === 'ql') {
                    paragraphAlignment = 'left';
                } else if (node.value === 'qc') {
                    paragraphAlignment = 'center';
                } else if (node.value === 'qr') {
                    paragraphAlignment = 'right';
                } else if (node.value === 'qj') {
                    paragraphAlignment = 'justify';
                }
            }
        };

        traverse(doc, {});

        // Flush any remaining table
        const finalCtx = getCurrentTable();
        if (inTable || (finalCtx && (finalCtx.rows.length > 0 || finalCtx.currentCells.length > 0))) {
            flushTable();
        }

        flushParagraph();

        // Notes handling:
        // - If putNotesAtLast is false, notes should be added inline during traversal
        //   (currently they go to 'notes' array, then we append them here - this is wrong)
        // - If putNotesAtLast is true, notes are appended at the very end (see below)
        // 
        // For now, when putNotesAtLast is false, we append notes immediately after content
        // This isn't truly "inline" but it's better than at the end
        // TODO: Implement true inline placement during traversal
        if (!config.putNotesAtLast && notes.length > 0) {
            content.push(...notes);
            notes.length = 0; // Clear so they don't get appended again
        }

        // Perform OCR if enabled
        if (config.ocr && config.extractAttachments) {
            for (const attachment of attachments) {
                if (attachment.mimeType.startsWith('image/')) {
                    try {
                        // Convert base64 data back to Buffer for Tesseract.js
                        // Passing base64 string directly would be interpreted as a file path,
                        // causing ENAMETOOLONG error for large images.
                        const imageBuffer = Buffer.from(attachment.data, 'base64');
                        attachment.ocrText = (await performOcr(imageBuffer, config.ocrLanguage)).trim();
                    } catch (e) {
                        logWarning(`OCR failed for ${attachment.name}:`, config, e);
                    }
                }
            }

            // Link OCR text and altText to image nodes in content
            const assignOcr = (nodes: OfficeContentNode[]) => {
                for (const node of nodes) {
                    if (node.type === 'image' && node.metadata && 'attachmentName' in node.metadata) {
                        const meta = node.metadata as ImageMetadata;
                        const attachment = attachments.find(a => a.name === meta.attachmentName);
                        if (attachment) {
                            // Propagate OCR text to image node
                            if (attachment.ocrText) {
                                node.text = attachment.ocrText;
                            }
                            // Propagate altText if available
                            if (attachment.altText) {
                                meta.altText = attachment.altText;
                            }
                        }
                    }
                    if (node.children) {
                        assignOcr(node.children);
                    }
                }
            };
            assignOcr(content);
        }

        // Final pass to ensure all 'note' nodes have their 'text' property populated
        // (This supports the simple toText implementation)
        const populateNoteText = (nodes: OfficeContentNode[]) => {
            for (const node of nodes) {
                if (node.type === 'note' && node.children) {
                    const getText = (n: OfficeContentNode): string => {
                        if (n.children && n.children.length > 0) return n.children.map(getText).join('');
                        return n.text || '';
                    };
                    node.text = node.children.map(getText).join('').trim();
                }
                if (node.children) {
                    populateNoteText(node.children);
                }
            }
        };
        populateNoteText(content);
        populateNoteText(notes);

        const result: OfficeParserAST = {
            type: 'rtf',
            metadata: {
                // RTF Limitation: No style map available (RTF uses inline styles)
            },
            content: content,
            attachments: attachments, // PNG and JPEG images extracted from \\pict groups
            toText: () => {
                let text = content.map(c => c.text).join(config.newlineDelimiter ?? '\n');
                if (config.putNotesAtLast && notes.length > 0) {
                    text += (config.newlineDelimiter ?? '\n') + notes.map(c => c.text).join(config.newlineDelimiter ?? '\n');
                }
                return text;
            }
        };

        // If putNotesAtLast is true, append notes to the end of the content array
        if (config.putNotesAtLast && notes.length > 0) {
            content.push(...notes);
        }

        return result;
    } catch (err) {
        throw err;
    }
};

// Helper function to extract font table from RTF document
function extractFontTable(doc: RtfGroup): { [key: number]: string } {
    const fontTable: { [key: number]: string } = {};

    // Recursive helper to find the font table group at any depth
    const findAndParseFontTable = (group: RtfGroup): boolean => {
        for (const node of group.content) {
            if (node.type === 'group') {
                if (node.destination === 'fonttbl') {
                    // Iterate through font definitions
                    for (const fontNode of node.content) {
                        if (fontNode.type === 'group') {
                            let fontIndex: number | undefined;
                            let fontName = '';

                            for (const item of fontNode.content) {
                                if (item.type === 'control' && item.value === 'f') {
                                    fontIndex = item.param;
                                } else if (item.type === 'text') {
                                    // Font name (may have trailing semicolon)
                                    fontName += item.value;
                                }
                            }

                            if (fontIndex !== undefined && fontName) {
                                // Remove trailing semicolon and whitespace
                                fontName = fontName.replace(/;$/, '').trim();
                                fontTable[fontIndex] = fontName;
                            }
                        }
                    }
                    return true; // Found and parsed
                }
                // Recurse into child groups
                if (findAndParseFontTable(node)) {
                    return true;
                }
            }
        }
        return false;
    };

    findAndParseFontTable(doc);
    return fontTable;
}

// Helper function to extract color table from RTF document
function extractColorTable(doc: RtfGroup): { [key: number]: string } {
    const colorTable: { [key: number]: string } = {};

    // Recursive helper to find the color table group at any depth
    const findAndParseColorTable = (group: RtfGroup): boolean => {
        for (const node of group.content) {
            if (node.type === 'group') {
                if (node.destination === 'colortbl') {
                    let colorIndex = 0;
                    let red = 0, green = 0, blue = 0;

                    for (const item of node.content) {
                        if (item.type === 'control') {
                            if (item.value === 'red' && item.param !== undefined) {
                                red = item.param;
                            } else if (item.value === 'green' && item.param !== undefined) {
                                green = item.param;
                            } else if (item.value === 'blue' && item.param !== undefined) {
                                blue = item.param;
                            }
                        } else if (item.type === 'text' && item.value === ';') {
                            // Semicolon marks end of color definition
                            const hex = `#${red.toString(16).padStart(2, '0')}${green.toString(16).padStart(2, '0')}${blue.toString(16).padStart(2, '0')}`;
                            colorTable[colorIndex] = hex;
                            colorIndex++;
                            red = 0;
                            green = 0;
                            blue = 0;
                        }
                    }
                    return true; // Found and parsed
                }
                // Recurse into child groups
                if (findAndParseColorTable(node)) {
                    return true;
                }
            }
        }
        return false;
    };

    findAndParseColorTable(doc);
    return colorTable;
}


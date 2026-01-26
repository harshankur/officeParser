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
/// <reference types="node" />
import { OfficeParserAST, OfficeParserConfig } from '../types';
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
export declare class SimpleRtfParser {
    /** Current position in the buffer */
    private index;
    /** The RTF content as a Buffer */
    private buffer;
    /** Total length of the buffer */
    private length;
    /**
     * Creates a new RTF parser.
     * @param buffer - The RTF file content as a Buffer
     */
    constructor(buffer: Buffer);
    parse(): RtfGroup;
    private parseControl;
    private parseText;
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
export declare const parseRtf: (buffer: Buffer, config: OfficeParserConfig) => Promise<OfficeParserAST>;

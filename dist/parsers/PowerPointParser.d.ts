/**
 * PowerPoint Presentation (PPTX) Parser
 *
 * **PPTX Format Overview:**
 * PPTX is the default format for Microsoft PowerPoint since Office 2007, based on OOXML.
 *
 * **File Structure:**
 * - `ppt/presentation.xml` - Presentation structure and slide list
 * - `ppt/slides/slide1.xml` - Individual slide content
 * - `ppt/notesSlides/notesSlide1.xml` - Speaker notes
 * - `ppt/slideLayouts/*` - Slide layout definitions
 * - `ppt/media/*` - Embedded images and media
 *
 * **Key Elements:**
 * - `<p:sld>` - Slide
 * - `<p:txBody>` - Text body containing paragraphs
 * - `<a:p>` - Paragraph
 * - `<a:r>` - Text run with formatting
 * - `<a:t>` - Text content
 *
 * @module PowerPointParser
 * @see https://www.ecma-international.org/publications-and-standards/standards/ecma-376/
 */
/// <reference types="node" />
import { OfficeParserAST, OfficeParserConfig } from '../types';
/**
 * Parses a PowerPoint presentation (.pptx) and extracts slides and notes.
 *
 * @param buffer - The PPTX file as a Buffer
 * @param config - Parser configuration
 * @returns A promise resolving to the parsed AST
 */
export declare const parsePowerPoint: (buffer: Buffer, config: OfficeParserConfig) => Promise<OfficeParserAST>;

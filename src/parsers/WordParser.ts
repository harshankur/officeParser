/**
 * Word Document (DOCX) Parser
 * 
 * **DOCX Format Overview:**
 * DOCX is the default format for Microsoft Word documents since Office 2007.
 * It's based on the Office Open XML (OOXML) standard (ECMA-376, ISO/IEC 29500).
 * 
 * **File Structure:**
 * DOCX files are ZIP archives containing:
 * - `word/document.xml` - Main document content
 * - `word/styles.xml` - Style definitions
 * - `word/numbering.xml` - List numbering definitions
 * - `word/footnotes.xml` - Footnotes content
 * - `word/media/*` - Embedded images and media
 * - `docProps/core.xml` - Document metadata
 * - `[Content_Types].xml` - MIME type mappings
 * 
 * **XML Structure (word/document.xml):**
 * ```xml
 * <w:document>
 *   <w:body>
 *     <w:p>                    <!-- Paragraph -->
 *       <w:pPr>                <!-- Paragraph properties -->
 *         <w:pStyle w:val="Heading1"/>
 *       </w:pPr>
 *       <w:r>                  <!-- Run (text with same formatting) -->
 *         <w:rPr>              <!-- Run properties -->
 *           <w:b/>             <!-- Bold -->
 *           <w:sz w:val="24"/> <!-- Font size (half-points) -->
 *         </w:rPr>
 *         <w:t>Hello</w:t>     <!-- Text -->
 *       </w:r>
 *     </w:p>
 *   </w:body>
 * </w:document>
 * ```
 *
 * **Key OOXML Elements:**
 * - `<w:p>` - Paragraph
 * - `<w:r>` - Run (contiguous text with same formatting)
 * - `<w:t>` - Text content
 * - `<w:br>` - Line or page break
 * - `<w:b>`, `<w:i>`, `<w:u>` - Bold, italic, underline
 * - `<w:pStyle>` - Paragraph style (for headings)
 * - `<w:numPr>` - List numbering properties
 * - `<w:tbl>` - Table
 * - `<w:drawing>` - Drawing/image
 *
 * **Parsing Approach:**
 * 1. Extract ZIP contents
 * 2. Parse word/document.xml for structure and text
 * 3. Extract formatting from run properties (rPr)
 * 4. Identify headings via paragraph styles
 * 5. Extract footnotes from word/footnotes.xml
 * 6. Process embedded images from word/media/*
 * 7. Parse metadata from docProps/core.xml
 *
 * @module WordParser
 * @see https://www.ecma-international.org/publications-and-standards/standards/ecma-376/ OOXML Standard
 * @see https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/ [MS-DOCX] Specification
 */

import { BreakMetadata, CellMetadata, FullOfficeParserConfig, ImageMetadata, IndentationMetadata, ListMetadata, OfficeAttachment, OfficeContentNode, OfficeParserAST, TextFormatting, TextMetadata, OfficeWarningType } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { checkAbortSignal, logWarning } from '../utils/errorUtils.js';
import { createAttachment } from '../utils/imageUtils.js';
import { performOcr } from '../utils/ocrUtils.js';
import { getDirectChildren, getElementsByTagName, getFirstElementByTagName, getRawContent, isElement, parseOfficeMetadata, parseOOXMLCustomProperties, parseXmlString, serializeXml } from '../utils/xmlUtils.js';
import { extractFiles } from '../utils/zipUtils.js';

/**
 * Parses a Word document (.docx) and extracts content, formatting, and metadata.
 * 
 * The parsing process:
 * 1. Unzip the DOCX file
 * 2. Parse word/document.xml to extract paragraphs and runs
 * 3. Extract text formatting from run properties
 * 4. Identify headings from paragraph styles
 * 5. Process lists from numbering properties
 * 6. Extract images and optionally perform OCR
 * 7. Parse document metadata
 * 
 * @param buffer - The DOCX file as a Buffer
 * @param config - Parser configuration options
 * @returns A promise resolving to the parsed AST
 */
export const parseWord = async (buffer: Buffer, config: FullOfficeParserConfig): Promise<OfficeParserAST> => {
    // Honour cancellation requests immediately — before opening the ZIP archive, loading XML
    // files, or kicking off any OCR work.  DOCX files can be large and the inflate + XML-parse
    // steps are synchronous-heavy, so failing fast here avoids wasted CPU time.
    checkAbortSignal(config.abortSignal);

    const documentFileRegex = /word\/document[\d+]?.xml/;
    const footnotesFileRegex = /word\/footnotes[\d+]?.xml/;
    const endnotesFileRegex = /word\/endnotes[\d+]?.xml/;
    const numberingFileRegex = /word\/numbering[\d+]?.xml/;
    const mediaFileRegex = /(word\/)?media\/.*/;
    const corePropsFileRegex = /docProps\/core[\d+]?.xml/;
    const customPropsFileRegex = /docProps\/custom\.xml/;
    const relsFileRegex = /word\/_rels\/document[\d+]?.xml\.rels/;
    const stylesFileRegex = /word\/styles[\d+]?.xml/;


    // Helper to extract formatting from run properties XML string
    const extractFormattingFromXml = (rPr: Element): TextFormatting => {
        const formatting: TextFormatting = {};

        // Helper to check boolean properties (e.g., <w:b />, <w:i w:val="0" />)
        const getBoolVal = (parent: Element, tagName: string): boolean | null => {
            const el = getFirstElementByTagName(parent, tagName);
            if (el) {
                const val = el.getAttribute('w:val');
                // In OOXML, if the element is present without w:val, it's true.
                // If w:val is present, it can be '1', 'true', 'on' for true.
                if (val === null) return true;
                return val === '1' || val === 'true' || val === 'on';
            }
            return null;
        };

        const bold = getBoolVal(rPr, 'w:b');
        if (bold !== null) formatting.bold = bold;

        const italic = getBoolVal(rPr, 'w:i');
        if (italic !== null) formatting.italic = italic;

        const u = getFirstElementByTagName(rPr, 'w:u');
        if (u) {
            const val = u.getAttribute('w:val');
            // If val is missing, it's a default underline (true). 
            // If val is present, it's true unless explicit 'none'.
            if (!val || val !== 'none') {
                formatting.underline = true;
            }
        }

        const strike = getBoolVal(rPr, 'w:strike');
        const dstrike = getBoolVal(rPr, 'w:dstrike');
        if (strike !== null) formatting.strikethrough = strike;
        else if (dstrike !== null) formatting.strikethrough = dstrike;

        // Font size (w:sz) - stored in half-points
        const sz = getFirstElementByTagName(rPr, 'w:sz');
        if (sz) {
            const val = sz.getAttribute('w:val');
            if (val) {
                formatting.size = (parseInt(val, 10) / 2).toString() + 'pt';
            }
        }

        // Color (w:color)
        const color = getFirstElementByTagName(rPr, 'w:color');
        if (color) {
            const val = color.getAttribute('w:val');
            if (val && val !== 'auto') {
                formatting.color = '#' + val;
            }
        }

        // Background color (w:shd) - shading
        const shd = getFirstElementByTagName(rPr, 'w:shd');
        if (shd) {
            const val = shd.getAttribute('w:fill');
            if (val && val !== 'auto') {
                formatting.backgroundColor = '#' + val;
            }
        }

        // Highlight (w:highlight) - maps to background color in our AST
        const highlight = getFirstElementByTagName(rPr, 'w:highlight');
        if (highlight) {
            const val = highlight.getAttribute('w:val');
            if (val && val !== 'none') {
                const colorMap: { [key: string]: string } = {
                    'yellow': '#FFFF00', 'green': '#00FF00', 'cyan': '#00FFFF', 'magenta': '#FF00FF',
                    'blue': '#0000FF', 'red': '#FF0000', 'darkBlue': '#00008B', 'darkCyan': '#008B8B',
                    'darkGreen': '#006400', 'darkMagenta': '#8B008B', 'darkRed': '#8B0000',
                    'darkYellow': '#808000', 'darkGray': '#A9A9A9', 'lightGray': '#D3D3D3', 'black': '#000000'
                };
                formatting.backgroundColor = colorMap[val] || val;
            }
        }

        // Font family (w:rFonts)
        const rFonts = getFirstElementByTagName(rPr, 'w:rFonts');
        if (rFonts) {
            // Priority: ascii (Western) > hAnsi (High ANSI)
            const font = rFonts.getAttribute('w:ascii') || rFonts.getAttribute('w:hAnsi');
            if (font) {
                formatting.font = font;
            }
        }

        // Subscript/Superscript (w:vertAlign)
        const vertAlign = getFirstElementByTagName(rPr, 'w:vertAlign');
        if (vertAlign) {
            const val = vertAlign.getAttribute('w:val');
            if (val === 'subscript') formatting.subscript = true;
            else if (val === 'superscript') formatting.superscript = true;
        }

        return formatting;
    };

    // Helper to extract indentation from paragraph properties XML string
    const extractIndentationFromXml = (pPr: Element): IndentationMetadata | undefined => {
        const ind = getFirstElementByTagName(pPr, "w:ind");
        if (ind) {
            const indentation: IndentationMetadata = {};
            const left = ind.getAttribute("w:left") || ind.getAttribute("w:start");
            const right = ind.getAttribute("w:right") || ind.getAttribute("w:end");
            const firstLine = ind.getAttribute("w:firstLine");
            const hanging = ind.getAttribute("w:hanging");

            if (left) indentation.left = parseInt(left, 10);
            if (right) indentation.right = parseInt(right, 10);
            if (firstLine) indentation.firstLine = parseInt(firstLine, 10);
            if (hanging) indentation.hanging = parseInt(hanging, 10);

            return Object.keys(indentation).length > 0 ? indentation : undefined;
        }
        return undefined;
    };

    /**
     * Resolves mc:AlternateContent by preferring mc:Fallback if choice namespace is not recognized,
     * or simply the first available valid child.
     */
    const resolveAlternateContent = (element: Element): Node[] => {
        const choice = getFirstElementByTagName(element, "mc:Choice");
        // In most cases, mc:Choice contains the modern version, but mc:Fallback is safer for legacy compatibility
        // Mammoth often skips Choice if it's not handled. We'll try Choice first.
        if (choice) return Array.from(choice.childNodes);

        const fallback = getFirstElementByTagName(element, "mc:Fallback");
        if (fallback) return Array.from(fallback.childNodes);

        return Array.from(element.childNodes);
    };

    const files = await extractFiles(buffer, x =>
        !!x.match(documentFileRegex) ||
        !!x.match(footnotesFileRegex) ||
        !!x.match(endnotesFileRegex) ||
        !!x.match(numberingFileRegex) ||
        !!x.match(corePropsFileRegex) ||
        !!x.match(customPropsFileRegex) ||
        !!x.match(relsFileRegex) ||
        !!x.match(stylesFileRegex) ||
        (!!config.extractAttachments && !!x.match(mediaFileRegex))
    );

    // Extract metadata
    const corePropsFile = files.find(f => f.path.match(corePropsFileRegex));
    const metadata = corePropsFile ? parseOfficeMetadata(corePropsFile.content.toString()) : {};
    const customPropsFile = files.find(f => f.path.match(customPropsFileRegex));
    if (customPropsFile) {
        const customProperties = parseOOXMLCustomProperties(customPropsFile.content.toString());
        if (Object.keys(customProperties).length > 0) metadata.customProperties = customProperties;
    }

    const footnoteMap = new Map<string, OfficeContentNode[]>();
    const endnoteMap = new Map<string, OfficeContentNode[]>();
    const collectedNotes: OfficeContentNode[] = [];
    const attachments: OfficeAttachment[] = [];
    const mediaFiles = files.filter(f => f.path.match(mediaFileRegex));

    // Extract relationships
    const relsFile = files.find(f => f.path.match(relsFileRegex));
    const relsMap: { [key: string]: string } = {};
    if (relsFile) {
        const relsXml = parseXmlString(relsFile.content.toString());
        const relationships = getElementsByTagName(relsXml, "Relationship");
        for (const relationship of relationships) {
            const id = relationship.getAttribute("Id");
            const target = relationship.getAttribute("Target");
            if (id && target) {
                relsMap[id] = target;
            }
        }
    }

    const numberingFile = files.find(f => f.path.match(numberingFileRegex));
    const numberingMap: { [key: string]: { [key: string]: { numFmt: string, lvlText: string, start: number } } } = {};

    if (numberingFile) {
        const numberingXml = parseXmlString(numberingFile.content.toString());
        const nums = getElementsByTagName(numberingXml, "w:num");
        const abstractNums = getElementsByTagName(numberingXml, "w:abstractNum");

        const abstractNumMap: { [key: string]: Element } = {};
        for (const abstractNum of abstractNums) {
            const abstractNumId = abstractNum.getAttribute("w:abstractNumId");
            if (abstractNumId) {
                abstractNumMap[abstractNumId] = abstractNum;
            }
        }

        for (const num of nums) {
            const numId = num.getAttribute("w:numId");
            const abstractNumIdNode = getFirstElementByTagName(num, "w:abstractNumId");
            const abstractNumId = abstractNumIdNode?.getAttribute("w:val");

            if (numId && abstractNumId && abstractNumMap[abstractNumId]) {
                numberingMap[numId] = {};

                // Inherit from abstractNum
                const lvls = getElementsByTagName(abstractNumMap[abstractNumId], "w:lvl");
                for (const lvl of lvls) {
                    const ilvl = lvl.getAttribute("w:ilvl");
                    const numFmtNode = getFirstElementByTagName(lvl, "w:numFmt");
                    const lvlTextNode = getFirstElementByTagName(lvl, "w:lvlText");
                    const startNode = getFirstElementByTagName(lvl, "w:start");
                    if (ilvl) {
                        numberingMap[numId][ilvl] = {
                            numFmt: numFmtNode?.getAttribute("w:val") || 'decimal',
                            lvlText: lvlTextNode?.getAttribute("w:val") || '',
                            start: parseInt(startNode?.getAttribute("w:val") || '1', 10)
                        };
                    }
                }

                // Apply instance overrides (w:lvlOverride)
                const overrides = getElementsByTagName(num, "w:lvlOverride");
                for (const override of overrides) {
                    const ilvl = override.getAttribute("w:ilvl");
                    if (ilvl && numberingMap[numId][ilvl]) {
                        const startOverride = getFirstElementByTagName(override, "w:startOverride");
                        if (startOverride) {
                            numberingMap[numId][ilvl].start = parseInt(startOverride.getAttribute("w:val") || '1', 10);
                        }
                    }
                }
            }
        }
    }

    // Parse Styles
    const stylesFile = files.find(f => f.path.match(stylesFileRegex));
    const styleMap: { [key: string]: { formatting: TextFormatting, alignment?: 'left' | 'center' | 'right' | 'justify', backgroundColor?: string, paragraphIndentation?: IndentationMetadata } } = {};

    if (stylesFile) {
        const stylesXml = parseXmlString(stylesFile.content.toString());
        const styles = getElementsByTagName(stylesXml, "w:style");

        for (const style of styles) {
            const styleId = style.getAttribute("w:styleId");
            if (styleId) {
                const rPr = getFirstElementByTagName(style, "w:rPr");
                const pPr = getFirstElementByTagName(style, "w:pPr");

                const formatting = rPr ? extractFormattingFromXml(rPr) : {};
                let alignment: 'left' | 'center' | 'right' | 'justify' | undefined = undefined;
                let backgroundColor: string | undefined = undefined;
                let paragraphIndentation: IndentationMetadata | undefined = undefined;

                if (pPr) {
                    const jc = getFirstElementByTagName(pPr, "w:jc");
                    if (jc) {
                        const val = jc.getAttribute("w:val");
                        if (val === 'left' || val === 'center' || val === 'right' || val === 'justify') {
                            alignment = val;
                        }
                    }
                    const shd = getFirstElementByTagName(pPr, "w:shd");
                    if (shd) {
                        const fill = shd.getAttribute("w:fill");
                        if (fill && fill !== 'auto') backgroundColor = '#' + fill;
                    }

                    const ind = extractIndentationFromXml(pPr);
                    if (ind) paragraphIndentation = ind;
                }

                styleMap[styleId] = { formatting, alignment, backgroundColor, paragraphIndentation };
            }
        }
    }

    // Extract document defaults
    let docDefaults: Partial<TextFormatting> = {};

    if (stylesFile) {
        const stylesXml = parseXmlString(stylesFile.content.toString());
        const docDefaultsNode = getFirstElementByTagName(stylesXml, "w:docDefaults");
        if (docDefaultsNode) {
            const rPrDefaultNode = getFirstElementByTagName(docDefaultsNode, "w:rPrDefault");
            if (rPrDefaultNode) {
                const rPr = getFirstElementByTagName(rPrDefaultNode, "w:rPr");
                if (rPr) {
                    docDefaults = extractFormattingFromXml(rPr);
                }
            }
        }
    }

    // Detect the default paragraph style (for international compatibility)
    let defaultParaStyleId: string | undefined = undefined;
    if (stylesFile) {
        const stylesXml = parseXmlString(stylesFile.content.toString());
        const styles = getElementsByTagName(stylesXml, "w:style");

        // Look for a style with w:type="paragraph" and w:default="1"
        for (const style of styles) {
            const styleType = style.getAttribute("w:type");
            const isDefault = style.getAttribute("w:default");
            const styleId = style.getAttribute("w:styleId");

            if (styleType === "paragraph" && isDefault === "1" && styleId) {
                defaultParaStyleId = styleId;
                break;
            }
        }

        // Fallback: if no default found, try "Normal"
        if (!defaultParaStyleId && styleMap["Normal"]) {
            defaultParaStyleId = "Normal";
        }
    }



    const content: OfficeContentNode[] = [];
    const rawContents: string[] = [];
    const numberingState: { [key: string]: { [key: string]: number } } = {};
    const listCounters: { [key: string]: { [key: string]: number } } = {}; // Track item index per listId/level

    // Helper to parse a paragraph node
    const parseParagraph = (pNode: Element, documentContent: string, pendingAnchorIds: string[] = []): OfficeContentNode => {
        // Check if it's a list item
        const numPr = getFirstElementByTagName(pNode, "w:numPr");
        const isList = !!numPr;

        // Check if it's a heading
        const pPr = getFirstElementByTagName(pNode, "w:pPr");
        const pStyle = pPr ? getFirstElementByTagName(pPr, "w:pStyle") : null;
        const pStyleVal = pStyle?.getAttribute("w:val");
        const isHeading = pStyleVal ? (pStyleVal.startsWith("Heading") || pStyleVal === "Title") : false;

        // Extract Paragraph Style Properties
        const styleProps = pStyleVal && styleMap[pStyleVal] ? styleMap[pStyleVal] : { formatting: {} };

        // Extract Alignment
        let alignment = styleProps.alignment;
        if (pPr) {
            const jc = getFirstElementByTagName(pPr, "w:jc");
            if (jc) {
                const val = jc.getAttribute("w:val");
                if (val === 'left' || val === 'center' || val === 'right' || val === 'justify') {
                    alignment = val;
                }
            }
        }

        // Extract Indentation
        let paraIndentation = styleProps.paragraphIndentation;
        if (pPr) {
            const ind = extractIndentationFromXml(pPr);
            if (ind) {
                paraIndentation = { ...paraIndentation, ...ind };
            }
        }

        // Extract Paragraph Background
        let paraBackgroundColor = styleProps.backgroundColor;
        if (pPr) {
            const shd = getFirstElementByTagName(pPr, "w:shd");
            if (shd) {
                const fill = shd.getAttribute("w:fill");
                if (fill && fill !== 'auto') {
                    paraBackgroundColor = '#' + fill;
                }
            }
        }

        // Extract paragraph-level run properties
        let paragraphRunFormatting: TextFormatting = { ...styleProps.formatting };
        if (pPr) {
            const pPrRPr = getFirstElementByTagName(pPr, "w:rPr");
            if (pPrRPr) {
                const pPrFormatting = extractFormattingFromXml(pPrRPr);
                for (const key in pPrFormatting) {
                    const value = pPrFormatting[key as keyof TextFormatting];
                    if (value === false) {
                        delete paragraphRunFormatting[key as keyof TextFormatting];
                    } else if (value !== undefined) {
                        (paragraphRunFormatting as any)[key] = value;
                    }
                }
            }
        }

        // Extract text and children
        let text = '';
        const children: OfficeContentNode[] = [];

        // Traverse children of paragraph (runs, hyperlinks, etc.)
        const processChildNode = (node: Node) => {
            if (isElement(node) && node.nodeName === 'w:r') {
                const runNode = node;
                const rPr = getFirstElementByTagName(runNode, "w:rPr");

                // Formatting
                let formatting: TextFormatting = {};
                // Apply paragraph-level formatting
                for (const key in paragraphRunFormatting) {
                    (formatting as any)[key] = (paragraphRunFormatting as any)[key];
                }

                // Check for run style
                const rStyle = rPr ? getFirstElementByTagName(rPr, "w:rStyle") : null;
                const rStyleVal = rStyle ? rStyle.getAttribute("w:val") : pStyleVal;
                if (rStyleVal && styleMap[rStyleVal]) {
                    for (const key in styleMap[rStyleVal].formatting) {
                        (formatting as any)[key] = (styleMap[rStyleVal].formatting as any)[key];
                    }
                }

                // Apply direct run properties
                if (rPr) {
                    const directFormatting = extractFormattingFromXml(rPr);
                    for (const key in directFormatting) {
                        const value = directFormatting[key as keyof TextFormatting];
                        if (value === false) {
                            delete formatting[key as keyof TextFormatting];
                        } else if (value !== undefined) {
                            formatting[key as keyof TextFormatting] = value as any;
                        }
                    }
                }

                // Inherit paragraph background
                if (!formatting.backgroundColor && paraBackgroundColor) {
                    formatting.backgroundColor = paraBackgroundColor;
                }

                for (const child of runNode.childNodes) {
                    if (!isElement(child)) continue;

                    // also handle unprefixed version (mirroring the behaviour of getElementsByTagName)

                    // Text content
                    if (child.tagName === "w:t" || child.tagName === "t") {
                        const tNode = child;

                        const tContent = tNode.textContent || '';
                        text += tContent;
                        const textNode: OfficeContentNode = {
                            type: 'text',
                            text: tContent,
                            formatting: formatting
                        };
                        if (config.includeRawContent) {
                            textNode.rawContent = getRawContent(tNode, documentContent, config);
                        }
                        // Always set a style: run style > paragraph style > detected default
                        // Use detected default style for international compatibility
                        const nodeStyle = rStyleVal || pStyleVal || defaultParaStyleId;
                        if (nodeStyle) {
                            textNode.metadata = { style: nodeStyle };
                        }
                        children.push(textNode);
                    }
                    // Break nodes
                    else if (config.includeBreakNodes &&
                        (child.tagName === "w:br"
                            || child.tagName === "br"
                            || child.tagName === "w:cr"
                            || child.tagName === "cr")
                    ) {
                        const brNode = child;

                        let breakType: BreakMetadata['breakType'] = 'textWrapping';
                        if (child.tagName === "w:cr" || child.tagName === "cr") {
                            breakType = 'carriageReturn';
                        } else {
                            const nodeBreakType = brNode.getAttribute("w:type") || brNode.getAttribute("type");
                            if (nodeBreakType !== null) {
                                breakType = nodeBreakType as BreakMetadata['breakType'];
                            }
                        }

                        let breakClear: BreakMetadata["clear"] = undefined;
                        if (breakType === 'textWrapping' && brNode.getAttribute("w:clear") !== null) {
                            breakClear = brNode.getAttribute("w:clear") as BreakMetadata["clear"];
                        }

                        const breakNode: OfficeContentNode = {
                            type: 'break',
                            metadata: { breakType, clear: breakClear }
                        };

                        if (config.includeRawContent) {
                            breakNode.rawContent = getRawContent(brNode, documentContent, config);
                        }

                        children.push(breakNode);
                    } else if (config.includeBreakNodes && (child.tagName === "w:lastRenderedPageBreak" || child.tagName === "lastRenderedPageBreak")) {
                        const breakNode: OfficeContentNode = {
                            type: 'break',
                            metadata: { breakType: 'lastRenderedPage' }
                        };

                        if (config.includeRawContent) {
                            breakNode.rawContent = getRawContent(child, documentContent, config);
                        }

                        children.push(breakNode);
                    }
                }

                // Images/Drawings
                if (config.extractAttachments) {
                    const drawings = getElementsByTagName(runNode, "w:drawing");
                    const picts = getElementsByTagName(runNode, "w:pict");
                    const allImages = [...drawings, ...picts];

                    for (const imgNode of allImages) {
                        const imgXml = serializeXml(imgNode);

                        // Extract Alt Text
                        let altText = '';
                        const docPr = getFirstElementByTagName(imgNode, "wp:docPr");
                        if (docPr) {
                            altText = docPr.getAttribute("descr") || docPr.getAttribute("title") || '';
                        }

                        // Extract Relationship ID
                        let rId = '';
                        const blip = getFirstElementByTagName(imgNode, "a:blip");
                        if (blip) {
                            rId = blip.getAttribute("r:embed") || '';
                        } else {
                            const imagedata = getFirstElementByTagName(imgNode, "v:imagedata");
                            if (imagedata) {
                                rId = imagedata.getAttribute("r:id") || '';
                            }
                        }

                        if (rId && relsMap[rId]) {
                            const target = relsMap[rId];
                            const filename = target.split('/').pop();
                            if (filename) {
                                const imageNode: OfficeContentNode = {
                                    type: 'image',
                                    text: '',
                                    metadata: { attachmentName: filename, altText: altText }
                                };
                                if (config.includeRawContent) {
                                    imageNode.rawContent = getRawContent(imgNode, documentContent, config);
                                }
                                children.push(imageNode);
                            }
                        } else {
                            const imageNode: OfficeContentNode = {
                                type: 'image',
                                text: '',
                            };
                            if (config.includeRawContent) {
                                imageNode.rawContent = getRawContent(imgNode, documentContent, config);
                            }
                            children.push(imageNode);
                        }
                    }
                }

                // Footnotes/Endnotes inside runs
                if (!config.ignoreNotes) {
                    const footnoteRef = getFirstElementByTagName(runNode, "w:footnoteReference");
                    if (footnoteRef) {
                        const id = footnoteRef.getAttribute("w:id");
                        if (id && footnoteMap.has(id)) {
                            const noteNodes = footnoteMap.get(id)!;
                            const noteNode: OfficeContentNode = {
                                type: 'note',
                                text: noteNodes.map((n: OfficeContentNode) => n.text).join(' '),
                                children: noteNodes,
                                metadata: { noteType: 'footnote', noteId: id }
                            } as any;

                            if (config.putNotesAtLast) {
                                collectedNotes.push(noteNode);
                            } else {
                                children.push(noteNode);
                            }
                        }
                    }

                    const endnoteRef = getFirstElementByTagName(runNode, "w:endnoteReference");
                    if (endnoteRef) {
                        const id = endnoteRef.getAttribute("w:id");
                        if (id && endnoteMap.has(id)) {
                            const noteNodes = endnoteMap.get(id)!;
                            const noteNode: OfficeContentNode = {
                                type: 'note',
                                text: noteNodes.map((n: OfficeContentNode) => n.text).join(' '),
                                children: noteNodes,
                                metadata: { noteType: 'endnote', noteId: id }
                            } as any;

                            if (config.putNotesAtLast) {
                                collectedNotes.push(noteNode);
                            } else {
                                children.push(noteNode);
                            }
                        }
                    }
                }
            } else if (isElement(node) && node.nodeName === 'w:hyperlink') {
                const hlNode = node;
                const rId = hlNode.getAttribute("r:id");
                const anchor = hlNode.getAttribute("w:anchor");

                let linkMetadata: TextMetadata | undefined;
                if (anchor && !config.ignoreInternalLinks) {
                    linkMetadata = { link: '#' + anchor, linkType: 'internal' };
                } else if (rId && relsMap[rId]) {
                    linkMetadata = { link: relsMap[rId], linkType: 'external' };
                }

                // Process children of hyperlink (usually runs)
                const hlChildren = Array.from(hlNode.childNodes);
                for (const child of hlChildren) {
                    // Capture the current length of children to apply metadata to new nodes
                    const startIndex = children.length;
                    processChildNode(child);
                    // Apply link metadata to the newly added text nodes
                    if (linkMetadata) {
                        for (let i = startIndex; i < children.length; i++) {
                            if (children[i].type === 'text') {
                                children[i].metadata = { ...(children[i].metadata ?? {}), ...linkMetadata };
                            }
                        }
                    }
                }
            } else if (isElement(node) && node.nodeName === 'w:bookmarkStart') {
                const bookmarkName = node.getAttribute("w:name");
                if (bookmarkName && !bookmarkName.startsWith('_GoBack') && !config.ignoreInternalLinks) {
                    anchorIds.push(bookmarkName);
                }
            } else if (isElement(node) && (node.nodeName === 'mc:AlternateContent' || node.nodeName === 'AlternateContent')) {
                const resolved = resolveAlternateContent(node);
                for (const rNode of resolved) processChildNode(rNode);
            } else if (isElement(node) && (node.nodeName === 'w:pict' || node.nodeName === 'pict' || node.nodeName === 'w:drawing' || node.nodeName === 'drawing')) {
                // Extract text boxes from legacy shapes or modern drawings
                const textBoxes = getElementsByTagName(node, "w:txbxContent");
                for (const txbx of textBoxes) {
                    const txbxChildren = Array.from(txbx.childNodes);
                    for (const txbxChild of txbxChildren) {
                        if (isElement(txbxChild) && txbxChild.nodeName === 'w:p') {
                            const nestedP = parseParagraph(txbxChild, documentContent);
                            children.push(...(nestedP.children || []));
                            text += nestedP.text;
                        }
                    }
                }
            } else if (node.childNodes.length > 0) {
                // Generic fallback for unknown elements that might contain content
                for (const child of Array.from(node.childNodes)) processChildNode(child);
            }
        };

        const anchorIds: string[] = [...pendingAnchorIds];
        const childNodes = Array.from(pNode.childNodes);
        for (const child of childNodes) {
            processChildNode(child);
        }

        const commonMetadata = anchorIds.length > 0 ? { anchorIds } : {};

        if (isList) {
            const numIdNode = getFirstElementByTagName(numPr, "w:numId");
            const ilvlNode = getFirstElementByTagName(numPr, "w:ilvl");
            const numId = numIdNode ? numIdNode.getAttribute("w:val") || '0' : '0';
            const ilvl = ilvlNode ? parseInt(ilvlNode.getAttribute("w:val") || '0', 10) : 0;

            let listType: 'ordered' | 'unordered' = 'ordered';
            let itemIndex = 0;
            if (numId && numberingMap[numId]) {
                const ilvlStr = ilvl.toString();
                if (!numberingState[numId]) numberingState[numId] = {};
                if (!numberingState[numId][ilvlStr]) numberingState[numId][ilvlStr] = 0;
                numberingState[numId][ilvlStr]++;
                for (let k = ilvl + 1; k < 10; k++) {
                    if (numberingState[numId][k.toString()]) numberingState[numId][k.toString()] = 0;
                }
                const numFmt = numberingMap[numId][ilvlStr]?.numFmt || 'decimal';
                listType = numFmt === 'bullet' ? 'unordered' : 'ordered';

                // Track itemIndex (starts at override or default, continues across interruptions for same listId)
                if (!listCounters[numId]) listCounters[numId] = {};
                if (listCounters[numId][ilvlStr] === undefined) {
                    listCounters[numId][ilvlStr] = (numberingMap[numId][ilvlStr]?.start ?? 1) - 1;
                } else {
                    listCounters[numId][ilvlStr]++;
                }
                itemIndex = listCounters[numId][ilvlStr];
            }

            const listNode: OfficeContentNode = {
                type: 'list',
                text: text,
                children: children,
                metadata: {
                    listType,
                    indentation: ilvl,
                    paragraphIndentation: paraIndentation,
                    alignment: (alignment || 'left') as 'left' | 'center' | 'right' | 'justify',
                    listId: numId,
                    itemIndex: itemIndex,
                    style: pStyleVal,
                    ...commonMetadata
                } as ListMetadata
            };

            if (config.includeRawContent) listNode.rawContent = getRawContent(pNode, documentContent, config);
            return listNode;

        } else if (isHeading) {
            const level = pStyleVal ? parseInt(pStyleVal.replace("Heading", ""), 10) || 1 : 1;
            const headingNode: OfficeContentNode = {
                type: 'heading',
                text: text,
                children: children,
                metadata: { level, alignment, paragraphIndentation: paraIndentation, style: pStyleVal ?? undefined, ...commonMetadata }
            };
            if (config.includeRawContent) headingNode.rawContent = getRawContent(pNode, documentContent, config);
            return headingNode;
        } else {
            const paraNode: OfficeContentNode = {
                type: 'paragraph',
                text: text,
                children: children,
                metadata: { alignment, paragraphIndentation: paraIndentation, style: pStyleVal ?? undefined, ...commonMetadata }
            };
            if (config.includeRawContent) paraNode.rawContent = getRawContent(pNode, documentContent, config);
            return paraNode;
        }
    };

    // Helper to parse a table node
    const parseTable = (tblNode: Element, documentContent: string, pendingAnchorIds: string[] = []): OfficeContentNode => {
        const rows: OfficeContentNode[] = [];
        const trNodes = getDirectChildren(tblNode, "w:tr");
        // Track vertical merges: colIndex -> { startCellNode, rowSpan }
        const vMergeMap = new Map<number, { node: OfficeContentNode, span: number }>();

        for (let rIndex = 0; rIndex < trNodes.length; rIndex++) {
            const trNode = trNodes[rIndex];
            const cells: OfficeContentNode[] = [];
            // Only get direct child cells, not nested table cells
            const tcNodes = getDirectChildren(trNode, "w:tc");

            let visualCol = 0;
            for (let tcIndex = 0; tcIndex < tcNodes.length; tcIndex++) {
                const tcNode = tcNodes[tcIndex];
                const tcPr = getFirstElementByTagName(tcNode, "w:tcPr");

                // Horizontal merge (colspan)
                let colSpan = 1;
                if (tcPr) {
                    const gridSpan = getFirstElementByTagName(tcPr, "w:gridSpan");
                    if (gridSpan) {
                        colSpan = parseInt(gridSpan.getAttribute("w:val") || "1", 10);
                    }
                }

                let vMergeRestart = false;
                let isVMerge = false;
                if (tcPr) {
                    const vMerge = getFirstElementByTagName(tcPr, "w:vMerge");
                    if (vMerge) {
                        isVMerge = true;
                        const val = vMerge.getAttribute("w:val");
                        // If it's explicit restart, or if we don't have an active merge for this column, treat as restart
                        if (val === "restart" || !vMergeMap.has(visualCol)) {
                            vMergeRestart = true;
                        }
                    }
                }

                const cellChildren: OfficeContentNode[] = [];
                let cellText = '';

                // Cells contain paragraphs (and other block-level elements)
                const cellContentNodes = Array.from(tcNode.childNodes);
                for (const child of cellContentNodes) {
                    if (isElement(child) && child.nodeName === 'w:p') {
                        const pNode = parseParagraph(child, documentContent);
                        cellChildren.push(pNode);
                        cellText += pNode.text;
                    } else if (isElement(child) && child.nodeName === 'w:tbl') {
                        const nestedTable = parseTable(child, documentContent);
                        cellChildren.push(nestedTable);
                    }
                }

                const cellNode: OfficeContentNode = {
                    type: 'cell',
                    text: cellText,
                    children: cellChildren,
                    metadata: { row: rIndex, col: visualCol } as CellMetadata
                };

                if (colSpan > 1) (cellNode.metadata as CellMetadata).colSpan = colSpan;

                if (isVMerge) {
                    if (vMergeRestart) {
                        vMergeMap.set(visualCol, { node: cellNode, span: 1 });
                        cells.push(cellNode);
                    } else {
                        const mergeInfo = vMergeMap.get(visualCol);
                        if (mergeInfo) {
                            mergeInfo.span++;
                            (mergeInfo.node.metadata as CellMetadata).rowSpan = mergeInfo.span;

                            if (cellChildren.length > 0) {
                                if (!mergeInfo.node.children) mergeInfo.node.children = [];
                                mergeInfo.node.children.push(...cellChildren);
                                mergeInfo.node.text += " " + cellText;
                            }
                        } else {
                            // Fallback: if we found a continue but no restart, treat as normal cell
                            cells.push(cellNode);
                        }
                    }
                } else {
                    vMergeMap.delete(visualCol);
                    cells.push(cellNode);
                }

                visualCol += colSpan;
            }

            const rowNode: OfficeContentNode = {
                type: 'row',
                children: cells,
            };
            rows.push(rowNode);
        }

        return {
            type: 'table',
            children: rows
        };
    };

    // Pre-process footnotes and endnotes to be inserted inline later
    if (!config.ignoreNotes) {
        const footnotesFile = files.find(f => f.path.match(footnotesFileRegex));
        if (footnotesFile) {
            const footnotesDoc = parseXmlString(footnotesFile.content.toString());
            const footnoteXml = footnotesFile.content.toString();
            const footnoteNodes = getElementsByTagName(footnotesDoc, "w:footnote");
            for (const node of footnoteNodes) {
                const id = node.getAttribute("w:id");
                if (!id || id === "-1" || id === "0") continue;
                const pNodes = getElementsByTagName(node, "w:p");
                footnoteMap.set(id, pNodes.map(p => parseParagraph(p, footnoteXml)));
            }
        }

        const endnotesFile = files.find(f => f.path.match(endnotesFileRegex));
        if (endnotesFile) {
            const endnotesDoc = parseXmlString(endnotesFile.content.toString());
            const endnoteXml = endnotesFile.content.toString();
            const endnoteNodes = getElementsByTagName(endnotesDoc, "w:endnote");
            for (const node of endnoteNodes) {
                const id = node.getAttribute("w:id");
                if (!id || id === "-1" || id === "0") continue;
                const pNodes = getElementsByTagName(node, "w:p");
                endnoteMap.set(id, pNodes.map(p => parseParagraph(p, endnoteXml)));
            }
        }
    }

    for (const file of files) {
        if (file.path.match(mediaFileRegex)) continue;
        if (file.path.match(numberingFileRegex)) continue;
        if (file.path.match(relsFileRegex)) continue;
        if (file.path.match(stylesFileRegex)) continue;
        if (file.path.match(footnotesFileRegex)) continue;
        if (file.path.match(endnotesFileRegex)) continue;

        const documentContent = file.content.toString();
        if (config.includeRawContent) {
            rawContents.push(documentContent);
        }

        const doc = parseXmlString(documentContent, { locator: config.includeRawContent });
        const body = getFirstElementByTagName(doc, "w:body");
        if (body) {
            const bodyChildren = Array.from(body.childNodes);
            let pendingAnchorIds: string[] = [];

            for (const child of bodyChildren) {
                if (isElement(child)) {
                    if (child.nodeName === 'w:p') {
                        content.push(parseParagraph(child, documentContent, pendingAnchorIds));
                        pendingAnchorIds = [];
                    } else if (child.nodeName === 'w:tbl') {
                        content.push(parseTable(child, documentContent, pendingAnchorIds));
                        pendingAnchorIds = [];
                    } else if (child.nodeName === 'w:bookmarkStart') {
                        const bookmarkName = child.getAttribute("w:name");
                        if (bookmarkName && !bookmarkName.startsWith('_GoBack') && !config.ignoreInternalLinks) {
                            pendingAnchorIds.push(bookmarkName);
                        }
                    }
                }
            }
        }
    }


    // Extract attachments
    if (config.extractAttachments) {
        for (const media of mediaFiles) {
            const attachment = createAttachment(media.path.split('/').pop() || 'image', media.content);
            attachments.push(attachment);

            if (config.ocr) {
                if (attachment.mimeType.startsWith('image/')) {
                    try {
                        attachment.ocrText = (await performOcr(media.content, { ...config.ocrConfig })).trim();
                    } catch (e) {
                        logWarning(OfficeWarningType.OCR_FAILED, config, attachment.name, e);
                    }
                }
            }
        }

        // Assign OCR text to image nodes
        if (config.ocr) {
            const assignOcr = (nodes: OfficeContentNode[]) => {
                for (const node of nodes) {
                    if (node.type === 'image' && 'attachmentName' in (node.metadata || {})) {
                        const meta = node.metadata as ImageMetadata;
                        const attachment = attachments.find(a => a.name === meta.attachmentName);
                        if (attachment && attachment.ocrText) {
                            node.text = attachment.ocrText;
                            attachment.altText = meta.altText;
                        }
                    }
                    if (node.children) {
                        assignOcr(node.children);
                    }
                }
            };
            assignOcr(content);
        }
    }

    if (config.putNotesAtLast && collectedNotes.length > 0) {
        content.push(...collectedNotes);
    }

    const toTextSync = () => content.map(c => {
        // Recursive text extraction
        const getText = (node: OfficeContentNode): string => {
            let t = '';
            if (node.children) {
                t += node.children.map(getText).filter(t => t != '').join(!node.children[0]?.children ? '' : config.newlineDelimiter);
            }
            else if (node.type === 'break') {
                t += config.newlineDelimiter;
            }
            else
                t += node.text || '';
            return t;
        };
        return getText(c);
    }).filter(t => t != '').join(config.newlineDelimiter);

    return createAST(
        'docx',
        { ...metadata, formatting: docDefaults, styleMap: styleMap },
        content,
        attachments,
        config,
        toTextSync
    );
};



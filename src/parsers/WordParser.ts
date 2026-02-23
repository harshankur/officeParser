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

import { XMLSerializer } from '@xmldom/xmldom';
import { ImageMetadata, ListMetadata, OfficeAttachment, OfficeContentNode, OfficeParserAST, OfficeParserConfig, TextFormatting, TextMetadata } from '../types';
import { logWarning } from '../utils/errorUtils';
import { createAttachment } from '../utils/imageUtils';
import { performOcr } from '../utils/ocrUtils';
import { getDirectChildren, getElementsByTagName, parseOfficeMetadata, parseOOXMLCustomProperties, parseXmlString } from '../utils/xmlUtils';
import { extractFiles } from '../utils/zipUtils';

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
export const parseWord = async (buffer: Buffer, config: OfficeParserConfig): Promise<OfficeParserAST> => {
    const documentFileRegex = /word\/document[\d+]?.xml/;
    const footnotesFileRegex = /word\/footnotes[\d+]?.xml/;
    const endnotesFileRegex = /word\/endnotes[\d+]?.xml/;
    const numberingFileRegex = /word\/numbering[\d+]?.xml/;
    const mediaFileRegex = /(word\/)?media\/.*/;
    const corePropsFileRegex = /docProps\/core[\d+]?.xml/;
    const customPropsFileRegex = /docProps\/custom\.xml/;
    const relsFileRegex = /word\/_rels\/document[\d+]?.xml\.rels/;
    const stylesFileRegex = /word\/styles[\d+]?.xml/;

    const xmlSerializer = new XMLSerializer();

    // Helper to extract formatting from run properties XML string
    const extractFormattingFromXml = (rPr: Element): TextFormatting => {
        const formatting: TextFormatting = {};
        const rPrString = xmlSerializer.serializeToString(rPr);

        // Helper to check boolean properties
        const getBoolVal = (xmlSnippet: string, tagName: string): boolean | null => {
            const regex = new RegExp(`<${tagName}(?:\\s+w:val="([^"]+)")?\\s*\\/?>`);
            const match = xmlSnippet.match(regex);
            if (match) {
                const val = match[1];
                if (val === undefined) return true;
                return val === '1' || val === 'true' || val === 'on';
            }
            return null;
        };

        const bold = getBoolVal(rPrString, 'w:b');
        if (bold !== null) formatting.bold = bold;

        const italic = getBoolVal(rPrString, 'w:i');
        if (italic !== null) formatting.italic = italic;

        const underlineMatch = rPrString.match(/<w:u(?: w:val="([^"]+)")?\/?>/);
        if (underlineMatch) {
            const val = underlineMatch[1];
            // If val is missing, it's a default underline (true). 
            // If val is present, it's true unless explicit 'none'.
            if (!val || val !== 'none') {
                formatting.underline = true;
            }
        }

        const strike = getBoolVal(rPrString, 'w:strike');
        const dstrike = getBoolVal(rPrString, 'w:dstrike');
        if (strike !== null) formatting.strikethrough = strike;
        else if (dstrike !== null) formatting.strikethrough = dstrike;

        // Font size
        const szMatch = rPrString.match(/<w:sz w:val="(\d+)"/);
        if (szMatch) formatting.size = (parseInt(szMatch[1]) / 2).toString() + 'pt';

        // Color
        const colorMatch = rPrString.match(/<w:color w:val="([^"]+)"/);
        if (colorMatch && colorMatch[1] !== 'auto') formatting.color = '#' + colorMatch[1];

        // Background color (shading)
        const shdMatch = rPrString.match(/<w:shd[^>]*w:fill="([^"]+)"/);
        if (shdMatch && shdMatch[1] !== 'auto') formatting.backgroundColor = '#' + shdMatch[1];

        // Highlight (map to backgroundColor)
        const highlightMatch = rPrString.match(/<w:highlight w:val="([^"]+)"/);
        if (highlightMatch && highlightMatch[1] !== 'none') {
            const colorMap: { [key: string]: string } = {
                'yellow': '#FFFF00', 'green': '#00FF00', 'cyan': '#00FFFF', 'magenta': '#FF00FF',
                'blue': '#0000FF', 'red': '#FF0000', 'darkBlue': '#00008B', 'darkCyan': '#008B8B',
                'darkGreen': '#006400', 'darkMagenta': '#8B008B', 'darkRed': '#8B0000',
                'darkYellow': '#808000', 'darkGray': '#A9A9A9', 'lightGray': '#D3D3D3', 'black': '#000000'
            };
            formatting.backgroundColor = colorMap[highlightMatch[1]] || highlightMatch[1];
        }

        // Font family
        const rFontsMatch = rPrString.match(/<w:rFonts[^>]*w:ascii="([^"]+)"/);
        if (rFontsMatch) {
            formatting.font = rFontsMatch[1];
        } else {
            const hAnsiMatch = rPrString.match(/<w:rFonts[^>]*w:hAnsi="([^"]+)"/);
            if (hAnsiMatch) formatting.font = hAnsiMatch[1];
        }

        // Subscript/Superscript
        const vertAlignMatch = rPrString.match(/<w:vertAlign w:val="([^"]+)"/);
        if (vertAlignMatch) {
            if (vertAlignMatch[1] === 'subscript') formatting.subscript = true;
            if (vertAlignMatch[1] === 'superscript') formatting.superscript = true;
        }

        return formatting;
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
        for (let i = 0; i < relationships.length; i++) {
            const id = relationships[i].getAttribute("Id");
            const target = relationships[i].getAttribute("Target");
            if (id && target) {
                relsMap[id] = target;
            }
        }
    }

    const numberingFile = files.find(f => f.path.match(numberingFileRegex));
    const numberingMap: { [key: string]: { [key: string]: { numFmt: string, lvlText: string } } } = {};

    if (numberingFile) {
        const numberingXml = parseXmlString(numberingFile.content.toString());
        const nums = getElementsByTagName(numberingXml, "w:num");
        const abstractNums = getElementsByTagName(numberingXml, "w:abstractNum");

        const abstractNumMap: { [key: string]: any } = {};
        for (let i = 0; i < abstractNums.length; i++) {
            const abstractNumId = abstractNums[i].getAttribute("w:abstractNumId");
            if (abstractNumId) {
                abstractNumMap[abstractNumId] = abstractNums[i];
            }
        }

        for (let i = 0; i < nums.length; i++) {
            const numId = nums[i].getAttribute("w:numId");
            const abstractNumIdNode = getElementsByTagName(nums[i], "w:abstractNumId")[0];
            const abstractNumId = abstractNumIdNode?.getAttribute("w:val");

            if (numId && abstractNumId && abstractNumMap[abstractNumId]) {
                numberingMap[numId] = {};
                const lvls = getElementsByTagName(abstractNumMap[abstractNumId], "w:lvl");
                for (let j = 0; j < lvls.length; j++) {
                    const ilvl = lvls[j].getAttribute("w:ilvl");
                    const numFmtNode = getElementsByTagName(lvls[j], "w:numFmt")[0];
                    const lvlTextNode = getElementsByTagName(lvls[j], "w:lvlText")[0];
                    if (ilvl) {
                        numberingMap[numId][ilvl] = {
                            numFmt: numFmtNode?.getAttribute("w:val") || 'decimal',
                            lvlText: lvlTextNode?.getAttribute("w:val") || ''
                        };
                    }
                }
            }
        }
    }

    // Parse Styles
    const stylesFile = files.find(f => f.path.match(stylesFileRegex));
    const styleMap: { [key: string]: { formatting: TextFormatting, alignment?: 'left' | 'center' | 'right' | 'justify', backgroundColor?: string } } = {};

    if (stylesFile) {
        const stylesXml = parseXmlString(stylesFile.content.toString());
        const styles = getElementsByTagName(stylesXml, "w:style");

        for (let i = 0; i < styles.length; i++) {
            const styleId = styles[i].getAttribute("w:styleId");
            if (styleId) {
                const rPr = getElementsByTagName(styles[i], "w:rPr")[0];
                const pPr = getElementsByTagName(styles[i], "w:pPr")[0];

                const formatting = rPr ? extractFormattingFromXml(rPr) : {};
                let alignment: 'left' | 'center' | 'right' | 'justify' | undefined = undefined;
                let backgroundColor: string | undefined = undefined;

                if (pPr) {
                    const jc = getElementsByTagName(pPr, "w:jc")[0];
                    if (jc) {
                        const val = jc.getAttribute("w:val");
                        if (val === 'left' || val === 'center' || val === 'right' || val === 'justify') {
                            alignment = val;
                        }
                    }
                    const shd = getElementsByTagName(pPr, "w:shd")[0];
                    if (shd) {
                        const fill = shd.getAttribute("w:fill");
                        if (fill && fill !== 'auto') backgroundColor = '#' + fill;
                    }
                }

                styleMap[styleId] = { formatting, alignment, backgroundColor };
            }
        }
    }

    // Extract document defaults
    let docDefaults: Partial<TextFormatting> = {};

    if (stylesFile) {
        const stylesXml = parseXmlString(stylesFile.content.toString());
        const docDefaultsNode = getElementsByTagName(stylesXml, "w:docDefaults")[0];
        if (docDefaultsNode) {
            const rPrDefaultNode = getElementsByTagName(docDefaultsNode, "w:rPrDefault")[0];
            if (rPrDefaultNode) {
                const rPr = getElementsByTagName(rPrDefaultNode, "w:rPr")[0];
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
        for (let i = 0; i < styles.length; i++) {
            const styleType = styles[i].getAttribute("w:type");
            const isDefault = styles[i].getAttribute("w:default");
            const styleId = styles[i].getAttribute("w:styleId");

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
    const parseParagraph = (pNode: Element): OfficeContentNode => {
        const pXml = xmlSerializer.serializeToString(pNode);

        // Check if it's a list item
        const numPr = getElementsByTagName(pNode, "w:numPr")[0];
        const isList = !!numPr;

        // Check if it's a heading
        const pPr = getElementsByTagName(pNode, "w:pPr")[0];
        const pStyle = pPr ? getElementsByTagName(pPr, "w:pStyle")[0] : null;
        const pStyleVal = pStyle ? pStyle.getAttribute("w:val") : null;
        const isHeading = pStyleVal ? (pStyleVal.startsWith("Heading") || pStyleVal === "Title") : false;

        // Extract Paragraph Style Properties
        const styleProps = pStyleVal && styleMap[pStyleVal] ? styleMap[pStyleVal] : { formatting: {} };

        // Extract Alignment
        let alignment = styleProps.alignment;
        if (pPr) {
            const jc = getElementsByTagName(pPr, "w:jc")[0];
            if (jc) {
                const val = jc.getAttribute("w:val");
                if (val === 'left' || val === 'center' || val === 'right' || val === 'justify') {
                    alignment = val;
                }
            }
        }

        // Extract Paragraph Background
        let paraBackgroundColor = styleProps.backgroundColor;
        if (pPr) {
            const shd = getElementsByTagName(pPr, "w:shd")[0];
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
            const pPrRPr = getElementsByTagName(pPr, "w:rPr")[0];
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
            if (node.nodeName === 'w:r') {
                const runNode = node as Element;
                const rPr = getElementsByTagName(runNode, "w:rPr")[0];

                // Formatting
                let formatting: TextFormatting = {};
                // Apply paragraph-level formatting
                for (const key in paragraphRunFormatting) {
                    (formatting as any)[key] = (paragraphRunFormatting as any)[key];
                }

                // Check for run style
                const rStyle = rPr ? getElementsByTagName(rPr, "w:rStyle")[0] : null;
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

                // Text content
                const tNodes = getElementsByTagName(runNode, "w:t");
                for (const tNode of tNodes) {
                    const tContent = tNode.textContent || '';
                    text += tContent;
                    const textNode: OfficeContentNode = {
                        type: 'text',
                        text: tContent,
                        formatting: formatting
                    };
                    if (config.includeRawContent) {
                        textNode.rawContent = xmlSerializer.serializeToString(tNode);
                    }
                    // Always set a style: run style > paragraph style > detected default
                    // Use detected default style for international compatibility
                    const nodeStyle = rStyleVal || pStyleVal || defaultParaStyleId;
                    if (nodeStyle) {
                        textNode.metadata = { style: nodeStyle };
                    }
                    children.push(textNode);
                }

                // Images/Drawings
                if (config.extractAttachments) {
                    const drawings = getElementsByTagName(runNode, "w:drawing");
                    const picts = getElementsByTagName(runNode, "w:pict");
                    const allImages = [...drawings, ...picts];

                    for (const imgNode of allImages) {
                        const imgXml = xmlSerializer.serializeToString(imgNode);

                        // Extract Alt Text
                        let altText = '';
                        const docPr = getElementsByTagName(imgNode, "wp:docPr")[0];
                        if (docPr) {
                            altText = docPr.getAttribute("descr") || docPr.getAttribute("title") || '';
                        }

                        // Extract Relationship ID
                        let rId = '';
                        const blip = getElementsByTagName(imgNode, "a:blip")[0];
                        if (blip) {
                            rId = blip.getAttribute("r:embed") || '';
                        } else {
                            const imagedata = getElementsByTagName(imgNode, "v:imagedata")[0];
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
                                    imageNode.rawContent = imgXml;
                                }
                                children.push(imageNode);
                            }
                        } else {
                            const imageNode: OfficeContentNode = {
                                type: 'image',
                                text: '',
                            };
                            if (config.includeRawContent) {
                                imageNode.rawContent = imgXml;
                            }
                            children.push(imageNode);
                        }
                    }
                }

                // Footnotes/Endnotes inside runs
                if (!config.ignoreNotes) {
                    const footnoteRef = getElementsByTagName(runNode, "w:footnoteReference")[0];
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

                    const endnoteRef = getElementsByTagName(runNode, "w:endnoteReference")[0];
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
            } else if (node.nodeName === 'w:hyperlink') {
                const hlNode = node as Element;
                const rId = hlNode.getAttribute("r:id");
                const anchor = hlNode.getAttribute("w:anchor");

                let linkMetadata: TextMetadata | undefined;
                if (anchor) {
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
            }
        };

        const childNodes = Array.from(pNode.childNodes);
        for (const child of childNodes) {
            processChildNode(child);
        }

        if (isList) {
            const numIdNode = getElementsByTagName(numPr, "w:numId")[0];
            const ilvlNode = getElementsByTagName(numPr, "w:ilvl")[0];
            const numId = numIdNode ? numIdNode.getAttribute("w:val") || '0' : '0';
            const ilvl = ilvlNode ? parseInt(ilvlNode.getAttribute("w:val") || '0') : 0;

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

                // Track itemIndex (starts at 0, continues across interruptions for same listId)
                if (!listCounters[numId]) listCounters[numId] = {};
                if (listCounters[numId][ilvlStr] === undefined) {
                    listCounters[numId][ilvlStr] = 0;
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
                    alignment: (alignment || 'left') as 'left' | 'center' | 'right' | 'justify',
                    listId: numId,
                    itemIndex: itemIndex,
                    style: pStyleVal
                } as ListMetadata
            };
            if (config.includeRawContent) listNode.rawContent = pXml;
            return listNode;

        } else if (isHeading) {
            const level = pStyleVal ? parseInt(pStyleVal.replace("Heading", "")) || 1 : 1;
            const headingNode: OfficeContentNode = {
                type: 'heading',
                text: text,
                children: children,
                metadata: { level, alignment, style: pStyleVal ?? undefined }
            };
            if (config.includeRawContent) headingNode.rawContent = pXml;
            return headingNode;
        } else {
            const paraNode: OfficeContentNode = {
                type: 'paragraph',
                text: text,
                children: children,
                metadata: { alignment, style: pStyleVal ?? undefined }
            };
            if (config.includeRawContent) paraNode.rawContent = pXml;
            return paraNode;
        }
    };

    // Helper to parse a table node
    const parseTable = (tblNode: Element): OfficeContentNode => {
        const rows: OfficeContentNode[] = [];
        // Only get direct child rows, not nested table rows
        const trNodes = getDirectChildren(tblNode, "w:tr");

        for (let rIndex = 0; rIndex < trNodes.length; rIndex++) {
            const trNode = trNodes[rIndex];
            const cells: OfficeContentNode[] = [];
            // Only get direct child cells, not nested table cells
            const tcNodes = getDirectChildren(trNode, "w:tc");

            for (let cIndex = 0; cIndex < tcNodes.length; cIndex++) {
                const tcNode = tcNodes[cIndex];
                const cellChildren: OfficeContentNode[] = [];
                let cellText = '';


                // Cells contain paragraphs (and other block-level elements)
                const cellContentNodes = Array.from(tcNode.childNodes);
                for (const child of cellContentNodes) {
                    if (child.nodeName === 'w:p') {
                        const pNode = parseParagraph(child as Element);
                        cellChildren.push(pNode);
                        cellText += pNode.text;
                    } else if (child.nodeName === 'w:tbl') {
                        // Nested table
                        const nestedTable = parseTable(child as Element);
                        cellChildren.push(nestedTable);
                        // Don't add nested table text to cell text - it will be handled recursively
                    }
                }

                const cellNode: OfficeContentNode = {
                    type: 'cell',
                    text: cellText,
                    children: cellChildren,
                    metadata: { row: rIndex, col: cIndex }
                };
                cells.push(cellNode);
            }

            const rowNode: OfficeContentNode = {
                type: 'row',
                children: cells
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
            const footnoteNodes = getElementsByTagName(footnotesDoc, "w:footnote");
            for (const node of footnoteNodes) {
                const id = node.getAttribute("w:id");
                if (!id || id === "-1" || id === "0") continue;
                const pNodes = getElementsByTagName(node, "w:p");
                footnoteMap.set(id, pNodes.map(p => parseParagraph(p)));
            }
        }

        const endnotesFile = files.find(f => f.path.match(endnotesFileRegex));
        if (endnotesFile) {
            const endnotesDoc = parseXmlString(endnotesFile.content.toString());
            const endnoteNodes = getElementsByTagName(endnotesDoc, "w:endnote");
            for (const node of endnoteNodes) {
                const id = node.getAttribute("w:id");
                if (!id || id === "-1" || id === "0") continue;
                const pNodes = getElementsByTagName(node, "w:p");
                endnoteMap.set(id, pNodes.map(p => parseParagraph(p)));
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

        const doc = parseXmlString(documentContent);
        const body = getElementsByTagName(doc, "w:body")[0];
        if (body) {
            const bodyChildren = Array.from(body.childNodes);
            for (const child of bodyChildren) {
                if (child.nodeName === 'w:p') {
                    content.push(parseParagraph(child as Element));
                } else if (child.nodeName === 'w:tbl') {
                    content.push(parseTable(child as Element));
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
                        attachment.ocrText = (await performOcr(media.content, config.ocrLanguage)).trim();
                    } catch (e) {
                        logWarning(`OCR failed for ${attachment.name}:`, config, e);
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

    return {
        type: 'docx',
        metadata: { ...metadata, formatting: docDefaults, styleMap: styleMap },
        content: content,
        attachments: attachments,
        toText: () => content.map(c => {
            // Recursive text extraction
            const getText = (node: OfficeContentNode): string => {
                let t = '';
                if (node.children) {
                    t += node.children.map(getText).filter(t => t != '').join(!node.children[0]?.children ? '' : config.newlineDelimiter ?? '\n');
                }
                else
                    t += node.text || '';
                return t;
            };
            return getText(c);
        }).filter(t => t != '').join(config.newlineDelimiter ?? '\n')
    };
};


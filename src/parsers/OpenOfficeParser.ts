/**
 * OpenDocument Format (ODF) Parser
 * 
 * **ODF Overview:**
 * ODF is an open standard for office documents (ISO/IEC 26300).
 * Used by LibreOffice, OpenOffice, and other applications.
 * 
 * **File Structure:**
 * ODF files are ZIP archives containing:
 * - `mimetype` - File type identification
 * - `content.xml` -  Main document content
 * - `styles.xml` - Style definitions
 * - `meta.xml` - Document metadata
 * - `Pictures/*` - Embedded images
 * 
 * **Supported Formats:**
 * - ODT: Text documents (application/vnd.oasis.opendocument.text)
 * - ODP: Presentations (application/vnd.oasis.opendocument.presentation)
 * - ODS: Spreadsheets (application/vnd.oasis.opendocument.spreadsheet)
 * 
 * @module OpenOfficeParser
 */

import { CellMetadata, ChartData, ChartMetadata, HeadingMetadata, ImageMetadata, ListMetadata, NoteMetadata, OfficeAttachment, OfficeContentNode, OfficeParserAST, OfficeParserConfig, SheetMetadata, SlideMetadata, SupportedFileType, TextFormatting, TextMetadata } from '../types';
import { extractChartData } from '../utils/chartUtils';
import { logWarning } from '../utils/errorUtils';
import { createAttachment } from '../utils/imageUtils';
import { performOcr } from '../utils/ocrUtils';
import { getDirectChildren, getElementsByTagName, parseOfficeMetadata, parseXmlString } from '../utils/xmlUtils';
import { extractFiles } from '../utils/zipUtils';

/**
 * Parses an OpenOffice document (.odt, .odp, .ods) and extracts content.
 * 
 * @param buffer - The ODF file as a Buffer
 * @param config - Parser configuration
 * @returns A promise resolving to the parsed AST
 */
export const parseOpenOffice = async (buffer: Buffer, config: OfficeParserConfig): Promise<OfficeParserAST> => {
    const contentFileRegex = /content\.xml/;
    const objectContentFileRegex = /Object \d+\/content\.xml/;
    const mediaFileRegex = /(Pictures|media)\/.*/;
    const metaFileRegex = /meta\.xml/;
    const stylesFileRegex = /styles\.xml/;
    const mimetypeFileRegex = /mimetype/;

    const files = await extractFiles(buffer, x =>
        !!x.match(contentFileRegex) ||
        !!x.match(objectContentFileRegex) ||
        !!x.match(metaFileRegex) ||
        !!x.match(stylesFileRegex) ||
        !!x.match(mimetypeFileRegex) ||
        (!!config.extractAttachments && !!x.match(mediaFileRegex))
    );

    // 1. Determine File Type
    const mimetypeFile = files.find(f => f.path === 'mimetype');
    let fileType: SupportedFileType = 'odt'; // Default
    if (mimetypeFile) {
        const mime = mimetypeFile.content.toString().trim();
        if (mime.includes('spreadsheet')) fileType = 'ods';
        else if (mime.includes('presentation')) fileType = 'odp';
        else if (mime.includes('text')) fileType = 'odt';
    }

    const mainContentFile = files.find(f => f.path === 'content.xml') || files.find(f => f.path.match(contentFileRegex));
    const stylesFile = files.find(f => f.path === 'styles.xml');
    const content: OfficeContentNode[] = [];
    const notes: OfficeContentNode[] = [];

    // Style Map: styleName -> TextFormatting
    // Inline style parsing (from content.xml automatic styles)
    const styleMap: { [key: string]: TextFormatting } = {};
    const paragraphStyleMap: { [key: string]: { alignment?: 'left' | 'center' | 'right' | 'justify', dropCap?: boolean } } = {};
    const listCounters: { [listId: string]: { [level: string]: number } } = {}; // Track item index per listId/level

    // Helper to parse styles
    const parseStyles = (xmlString: string) => {
        const xml = parseXmlString(xmlString);
        const styles = getElementsByTagName(xml, "style:style");
        for (const style of styles) {
            const name = style.getAttribute("style:name");
            if (!name) continue;

            const styleInfo: { alignment?: 'left' | 'center' | 'right' | 'justify', dropCap?: boolean } = {};

            // Parse paragraph properties for alignment and drop caps
            const paraProps = getElementsByTagName(style, "style:paragraph-properties")[0];
            if (paraProps) {
                const textAlign = paraProps.getAttribute("fo:text-align");
                if (textAlign) {
                    const alignMap: Record<string, 'left' | 'center' | 'right' | 'justify'> = {
                        'start': 'left',
                        'left': 'left',
                        'center': 'center',
                        'end': 'right',
                        'right': 'right',
                        'justify': 'justify'
                    };
                    if (alignMap[textAlign]) {
                        styleInfo.alignment = alignMap[textAlign];
                    }
                }

                // Detect Drop Caps
                const dropCap = getElementsByTagName(paraProps, "style:drop-cap")[0];
                if (dropCap) {
                    styleInfo.dropCap = true;
                }
            }

            if (Object.keys(styleInfo).length > 0) {
                paragraphStyleMap[name] = styleInfo;
            }

            // Parse text properties
            const textProps = getElementsByTagName(style, "style:text-properties")[0];
            // Parse table cell properties (for ODS background)
            const cellProps = getElementsByTagName(style, "style:table-cell-properties")[0];

            const formatting: TextFormatting = {};

            if (cellProps) {
                const bgColor = cellProps.getAttribute("fo:background-color");
                if (bgColor && bgColor !== 'transparent') formatting.backgroundColor = bgColor;
            }

            if (textProps) {
                if (textProps.getAttribute("fo:font-weight") === "bold" || textProps.getAttribute("style:font-weight-asian") === "bold") formatting.bold = true;
                if (textProps.getAttribute("fo:font-style") === "italic" || textProps.getAttribute("style:font-style-asian") === "italic") formatting.italic = true;
                if (textProps.getAttribute("style:text-underline-style") === "solid") formatting.underline = true;
                if (textProps.getAttribute("style:text-line-through-style") === "solid") formatting.strikethrough = true;
                const size = textProps.getAttribute("fo:font-size") || textProps.getAttribute("style:font-size-asian");
                if (size) formatting.size = size;
                const color = textProps.getAttribute("fo:color");
                if (color) formatting.color = color;

                // Background color (text level) - override cell level if present?
                const bgColor = textProps.getAttribute("fo:background-color");
                if (bgColor && bgColor !== 'transparent') formatting.backgroundColor = bgColor;

                // Font family
                const fontName = textProps.getAttribute("style:font-name") || textProps.getAttribute("fo:font-family");
                if (fontName) formatting.font = fontName;

                // Subscript/Superscript from text-position (e.g., "sub 58%" or "super 58%")
                const textPosition = textProps.getAttribute("style:text-position");
                if (textPosition) {
                    if (textPosition.startsWith("sub")) formatting.subscript = true;
                    if (textPosition.startsWith("super")) formatting.superscript = true;
                }

                if (Object.keys(formatting).length > 0) styleMap[name] = formatting;
            }
        }
    };

    if (stylesFile) {
        parseStyles(stylesFile.content.toString());
    }

    /**
     * Helper to parse a paragraph node (text:p or text:h) and extract its content.
     * Returns the paragraph content without creating a content node.
     * 
     * @param node - The paragraph element to parse
     * Helper to parse inline content (text, spans, links, notes, etc.) recursively.
     * 
     * @param node - The element to parse (paragraph, span, or link)
     * @param styleMap - Map of style names to formatting
     * @param config - Parser configuration
     * @param notes - Optional array to collect footnotes/endnotes
     * @param paragraphStyleMap - Map of style names to alignments and props (needed for notes)
     * @param parentFormatting - Formatting inherited from parent (e.g. span inside span)
     * @param linkMetadata - Metadata inherited from parent link
     * @returns Object containing text and children
     */
    const parseInlineContent = (
        node: Node,
        styleMap: Record<string, TextFormatting>,
        config: OfficeParserConfig,
        notes: OfficeContentNode[],
        paragraphStyleMap: Record<string, { alignment?: 'left' | 'center' | 'right' | 'justify', dropCap?: boolean }>,
        parentFormatting: TextFormatting = {},
        linkMetadata?: { link?: string; linkType?: 'internal' | 'external' }
    ): { text: string; children: OfficeContentNode[] } => {
        const children: OfficeContentNode[] = [];
        let fullText = '';

        if (!node.childNodes) return { text: '', children: [] };

        for (let i = 0; i < node.childNodes.length; i++) {
            const child = node.childNodes[i];

            if (child.nodeType === 3) { // Text node
                const text = child.textContent || '';
                if (text) {
                    fullText += text;
                    children.push({
                        type: 'text',
                        text: text,
                        formatting: parentFormatting,
                        metadata: linkMetadata ? { ...linkMetadata } : undefined
                    });
                }
            } else if (child.nodeType === 1) {
                const element = child as Element;
                const tagName = element.tagName;

                if (tagName === 'text:s') {
                    // Space
                    const count = parseInt(element.getAttribute('text:c') || '1');
                    const spaces = ' '.repeat(count);
                    fullText += spaces;
                    children.push({
                        type: 'text',
                        text: spaces,
                        formatting: parentFormatting,
                        metadata: linkMetadata ? { ...linkMetadata } : undefined
                    });
                } else if (tagName === 'text:tab') {
                    // Tab
                    fullText += '\t';
                    children.push({
                        type: 'text',
                        text: '\t',
                        formatting: parentFormatting,
                        metadata: linkMetadata ? { ...linkMetadata } : undefined
                    });
                } else if (tagName === 'text:line-break') {
                    // Line break
                    fullText += '\n';
                    children.push({
                        type: 'text',
                        text: '\n',
                        formatting: parentFormatting,
                        metadata: linkMetadata ? { ...linkMetadata } : undefined
                    });
                } else if (tagName === 'text:span') {
                    // Formatted text span
                    const styleName = element.getAttribute("text:style-name");
                    const formatting = styleName ? { ...parentFormatting, ...styleMap[styleName] } : parentFormatting;

                    const spanContent = parseInlineContent(element, styleMap, config, notes, paragraphStyleMap, formatting, linkMetadata);
                    fullText += spanContent.text;
                    children.push(...spanContent.children);
                } else if (tagName === 'text:a') {
                    // Hyperlink
                    const href = element.getAttribute('xlink:href') || '';
                    const linkType = href.startsWith('#') ? 'internal' : 'external';
                    const newLinkMetadata = { link: href, linkType: linkType as 'internal' | 'external' };

                    const linkContent = parseInlineContent(element, styleMap, config, notes, paragraphStyleMap, parentFormatting, newLinkMetadata);
                    fullText += linkContent.text;
                    children.push(...linkContent.children);
                } else if (tagName === 'text:note' && !config.ignoreNotes) {
                    // Footnote or endnote
                    const noteClass = (element.getAttribute('text:note-class') || 'footnote') as 'footnote' | 'endnote';
                    const noteId = element.getAttribute('text:id') || element.getAttribute('xml:id') || undefined;
                    const noteBody = getElementsByTagName(element, "text:note-body")[0];

                    if (noteBody) {
                        // Extract note content recursively
                        const notePs = getElementsByTagName(noteBody, "text:p");
                        const noteChildren: OfficeContentNode[] = [];
                        let noteText = '';

                        for (const np of notePs) {
                            const npContent = parseParagraphContent(np, paragraphStyleMap, styleMap, config);
                            noteText += (noteText ? ' ' : '') + npContent.text;

                            const npNode: OfficeContentNode = {
                                type: 'paragraph',
                                text: npContent.text,
                                children: npContent.children,
                                metadata: npContent.alignment ? { alignment: npContent.alignment } : undefined
                            };
                            noteChildren.push(npNode);
                        }

                        const noteNode: OfficeContentNode = {
                            type: 'note',
                            text: noteText,
                            children: noteChildren,
                            metadata: {
                                noteType: noteClass,
                                noteId: noteId
                            }
                        };

                        if (config.putNotesAtLast) {
                            notes.push(noteNode);
                        } else {
                            children.push(noteNode);
                        }
                    }
                } else if (tagName === 'draw:frame') {
                    // Inline image
                    const frame = element;

                    // Extract alt text
                    let altText = '';
                    const svgTitle = getElementsByTagName(frame, "svg:title")[0];
                    const svgDesc = getElementsByTagName(frame, "svg:desc")[0];
                    if (svgTitle && svgTitle.textContent) {
                        altText = svgTitle.textContent;
                    } else if (svgDesc && svgDesc.textContent) {
                        altText = svgDesc.textContent;
                    }

                    // Extract image href
                    let imageHref = '';
                    const drawImages = getElementsByTagName(frame, "draw:image");
                    if (drawImages.length > 0) {
                        imageHref = drawImages[0].getAttribute("xlink:href") || '';
                        if (imageHref) {
                            const parts = imageHref.split('/');
                            imageHref = parts[parts.length - 1];
                        }
                    }

                    const imageNode: OfficeContentNode = {
                        type: 'image',
                        text: '',
                        children: [],
                        metadata: {
                            attachmentName: imageHref,
                            ...(altText ? { altText } : {})
                        }
                    };
                    if (config.includeRawContent) {
                        imageNode.rawContent = frame.toString();
                    }
                    children.push(imageNode);
                }
            }
        }
        return { text: fullText, children };
    };

    /**
     * Helper to parse a paragraph node (text:p or text:h) and extract its content.
     * Returns the paragraph content without creating a content node.
     * 
     * @param node - The paragraph element to parse
     * @param paraStyleMap - Map of style names to alignments/props
     * @param styleMap - Map of style names to formatting
     * @param config - Parser configuration
     * @returns Object containing text, children, alignment, and style info
     */
    const parseParagraphContent = (
        node: Element,
        paraStyleMap: Record<string, { alignment?: 'left' | 'center' | 'right' | 'justify', dropCap?: boolean }>,
        styleMap: Record<string, TextFormatting>,
        config: OfficeParserConfig
    ): { text: string; children: OfficeContentNode[]; alignment?: 'left' | 'center' | 'right' | 'justify'; style?: string } => {
        // Get paragraph style for alignment and drop caps
        const paraStyle = node.getAttribute("text:style-name");
        const styleInfo = paraStyle ? paraStyleMap[paraStyle] : undefined;
        const alignment = styleInfo?.alignment;
        const dropCap = styleInfo?.dropCap;

        // Parse content recursively using the new helper
        const content = parseInlineContent(node, styleMap, config, notes, paraStyleMap);

        // Add style name to metadata of children if they don't have one
        if (paraStyle) {
            content.children.forEach(child => {
                if (child.type === 'text') {
                    if (!child.metadata) child.metadata = {};
                    // Only add style if it's a text node and doesn't have one?
                    // Or just add it.
                    // Cast to any to avoid union type issues for now, or check type
                    const meta = child.metadata as TextMetadata;
                    if (!meta.style) meta.style = paraStyle;
                }
            });
        }

        // Fallback: if no children were created but there's text content
        if (content.children.length === 0 && node.textContent) {
            const fullText = node.textContent;
            if (fullText.trim()) {
                content.text = fullText;
                content.children.push({
                    type: 'text',
                    text: fullText
                });
            }
        }

        // Handle Drop Cap: Apply large font to first letter if configured
        if (dropCap && content.children.length > 0) {
            const firstChild = content.children[0];
            if (firstChild.type === 'text' && firstChild.text) {
                if (firstChild.text.length === 1) {
                    // Already a single letter, just apply formatting
                    firstChild.formatting = { ...firstChild.formatting, size: '58.5pt' };
                } else {
                    // Split text node
                    const firstChar = firstChild.text[0];
                    const restText = firstChild.text.substring(1);

                    const dropCapNode: OfficeContentNode = {
                        type: 'text',
                        text: firstChar,
                        formatting: { ...firstChild.formatting, size: '58.5pt' },
                        metadata: firstChild.metadata
                    };

                    // Update original node
                    firstChild.text = restText;

                    // Insert drop cap node
                    content.children.unshift(dropCapNode);
                }
            }
        }

        return { text: content.text, children: content.children, alignment, style: paraStyle || undefined };
    };

    /**
     * Helper to parse a table node and extract its structure.
     * Properly creates table → row → cell hierarchy with metadata.
     * 
     * @param tableNode - The table:table element
     * @param paraStyleMap - Map of style names to alignments
     * @param styleMap - Map of style names to formatting
     * @param config - Parser configuration
     * @returns Table content node with proper structure
     */
    const parseTable = (
        tableNode: Element,
        paraStyleMap: Record<string, { alignment?: 'left' | 'center' | 'right' | 'justify', dropCap?: boolean }>,
        styleMap: Record<string, TextFormatting>,
        config: OfficeParserConfig
    ): OfficeContentNode => {
        const rows: OfficeContentNode[] = [];
        // Use getDirectChildren to avoid nested table rows
        const tableRows = getDirectChildren(tableNode, "table:table-row");
        let rowIndex = 0;

        for (const row of tableRows) {
            const cells: OfficeContentNode[] = [];
            // Use getDirectChildren to avoid nested table cells
            const tableCells = getDirectChildren(row, "table:table-cell");
            const rowsRepeated = parseInt(row.getAttribute("table:number-rows-repeated") || "1");

            let colIndex = 0;

            for (const cell of tableCells) {
                const cellChildren: OfficeContentNode[] = [];
                let cellTextRef = { value: '' };
                const colsRepeated = parseInt(cell.getAttribute("table:number-columns-repeated") || "1");
                const colSpan = parseInt(cell.getAttribute("table:number-columns-spanned") || "1");
                const rowSpan = parseInt(cell.getAttribute("table:number-rows-spanned") || "1");

                // Helper to recursively process cell children (handles frames, text-boxes, etc. in ODP)
                const processChildren = (node: Element) => {
                    if (!node.childNodes) return;

                    for (let i = 0; i < node.childNodes.length; i++) {
                        const child = node.childNodes[i];
                        if (child.nodeType === 1) { // Element
                            const element = child as Element;

                            if (element.tagName === "text:p" || element.tagName === "text:h") {
                                const pContent = parseParagraphContent(element, paraStyleMap, styleMap, config);
                                const pNode: OfficeContentNode = {
                                    type: element.tagName === "text:h" ? 'heading' : 'paragraph',
                                    text: pContent.text,
                                    children: pContent.children,
                                    metadata: {
                                        ...(pContent.alignment ? { alignment: pContent.alignment } : {}),
                                        ...(pContent.style ? { style: pContent.style } : {})
                                    }
                                };

                                // Clean up metadata if empty
                                if (Object.keys(pNode.metadata || {}).length === 0) delete pNode.metadata;

                                if (element.tagName === "text:h") {
                                    if (!pNode.metadata) pNode.metadata = {};
                                    (pNode.metadata as HeadingMetadata).level = parseInt(element.getAttribute("text:outline-level") || "1");
                                }

                                if (config.includeRawContent) {
                                    pNode.rawContent = element.toString();
                                }

                                cellChildren.push(pNode);
                                cellTextRef.value += pContent.text;
                                // Add newline if there are multiple paragraphs/headings
                                if (cellTextRef.value && !cellTextRef.value.endsWith('\n')) {
                                    cellTextRef.value += '\n';
                                }
                            } else if (element.tagName === "table:table") {
                                // Recursive call for nested table
                                const nestedTableNode = parseTable(element, paraStyleMap, styleMap, config);
                                cellChildren.push(nestedTableNode);
                            } else if (element.tagName === "draw:frame" || element.tagName === "draw:text-box") {
                                // Recursively process container content (common in ODP)
                                processChildren(element);
                            }
                        }
                    }
                };

                processChildren(cell);

                let cellText = cellTextRef.value;
                // Trim trailing newline from cellText
                if (cellText.endsWith('\n')) {
                    cellText = cellText.slice(0, -1);
                }

                // Add cell(s) for repeated columns
                for (let k = 0; k < colsRepeated; k++) {
                    const cellNode: OfficeContentNode = {
                        type: 'cell',
                        text: cellText,
                        children: cellChildren.length > 0 ? (k === 0 ? cellChildren : JSON.parse(JSON.stringify(cellChildren))) : [],
                        metadata: { row: rowIndex, col: colIndex } as CellMetadata
                    };

                    const cellMetadata = cellNode.metadata as CellMetadata;
                    if (colSpan > 1) cellMetadata.colSpan = colSpan;
                    if (rowSpan > 1) cellMetadata.rowSpan = rowSpan;

                    if (config.includeRawContent) {
                        cellNode.rawContent = cell.toString();
                    }

                    cells.push(cellNode);
                    colIndex++;
                }
            }

            // Add row(s) for repeated rows
            for (let k = 0; k < rowsRepeated; k++) {
                const rowNode: OfficeContentNode = {
                    type: 'row',
                    children: k === 0 ? cells : JSON.parse(JSON.stringify(cells))
                };

                // Fix row indices for repeated rows
                if (k > 0) {
                    rowNode.children?.forEach(c => {
                        if (c.metadata && 'row' in c.metadata) {
                            (c.metadata as CellMetadata).row = rowIndex;
                        }
                    });
                }

                if (config.includeRawContent) {
                    rowNode.rawContent = row.toString();
                }

                rows.push(rowNode);
                rowIndex++;
            }
        }

        return {
            type: 'table',
            children: rows
        };
    };


    const parseContentXml = (xmlString: string) => {
        const xml = parseXmlString(xmlString);
        const body = getElementsByTagName(xml, "office:body")[0];
        // Parse automatic styles (local to content.xml)
        const automaticStyles = getElementsByTagName(xml, "office:automatic-styles")[0];
        if (automaticStyles) {
            const styles = getElementsByTagName(automaticStyles, "style:style");
            for (const style of styles) {
                const name = style.getAttribute("style:name");
                if (!name) continue;

                // Parse paragraph properties for alignment
                const paraProps = getElementsByTagName(style, "style:paragraph-properties")[0];
                const styleInfo: { alignment?: 'left' | 'center' | 'right' | 'justify', dropCap?: boolean } = {};

                if (paraProps) {
                    const textAlign = paraProps.getAttribute("fo:text-align");
                    if (textAlign) {
                        const alignMap: Record<string, 'left' | 'center' | 'right' | 'justify'> = {
                            'start': 'left',
                            'left': 'left',
                            'center': 'center',
                            'end': 'right',
                            'right': 'right',
                            'justify': 'justify'
                        };
                        if (alignMap[textAlign]) {
                            styleInfo.alignment = alignMap[textAlign];
                        }
                    }

                    const dropCap = getElementsByTagName(paraProps, "style:drop-cap")[0];
                    if (dropCap) styleInfo.dropCap = true;
                }

                if (Object.keys(styleInfo).length > 0) {
                    paragraphStyleMap[name] = styleInfo;
                }

                const textProps = getElementsByTagName(style, "style:text-properties")[0];
                if (textProps) {
                    const formatting: TextFormatting = {};
                    if (textProps.getAttribute("fo:font-weight") === "bold" || textProps.getAttribute("style:font-weight-asian") === "bold") formatting.bold = true;
                    if (textProps.getAttribute("fo:font-style") === "italic" || textProps.getAttribute("style:font-style-asian") === "italic") formatting.italic = true;
                    if (textProps.getAttribute("style:text-underline-style") === "solid") formatting.underline = true;
                    if (textProps.getAttribute("style:text-line-through-style") === "solid") formatting.strikethrough = true;
                    const size = textProps.getAttribute("fo:font-size") || textProps.getAttribute("style:font-size-asian");
                    if (size) formatting.size = size;
                    const color = textProps.getAttribute("fo:color");
                    if (color) formatting.color = color;

                    // Background color
                    const bgColor = textProps.getAttribute("fo:background-color");
                    if (bgColor && bgColor !== 'transparent') formatting.backgroundColor = bgColor;

                    // Font family
                    const fontName = textProps.getAttribute("style:font-name") || textProps.getAttribute("fo:font-family");
                    if (fontName) formatting.font = fontName;

                    // Subscript/Superscript from text-position (e.g., "sub 58%" or "super 58%")
                    const textPosition = textProps.getAttribute("style:text-position");
                    if (textPosition) {
                        if (textPosition.startsWith("sub")) formatting.subscript = true;
                        if (textPosition.startsWith("super")) formatting.superscript = true;
                    }

                    if (Object.keys(formatting).length > 0) styleMap[name] = formatting;
                }
            }
        }

        /**
         * Recursively traverses a node and its children to extract content.
         * Properly handles paragraphs, headings, tables, lists, and frames.
         * 
         * @param node - The element to traverse
         * @param targetArray - The array to push extracted content nodes to
         * @param forceHeading - If true, treats all paragraphs as headings (used for slide titles)
         */
        const traverse = (node: Element, targetArray: OfficeContentNode[], forceHeading = false) => {
            if (node.tagName === "text:p") {
                const pContent = parseParagraphContent(node, paragraphStyleMap, styleMap, config);
                const type = (forceHeading || (node.getAttribute("text:style-name") || '').toLowerCase().includes('title')) ? 'heading' : 'paragraph';

                const pNode: OfficeContentNode = {
                    type,
                    text: pContent.text,
                    children: pContent.children,
                    metadata: {
                        ...(pContent.alignment ? { alignment: pContent.alignment } : {}),
                        ...(pContent.style ? { style: pContent.style } : {})
                    }
                };
                if (type === 'heading' && pNode.metadata) {
                    (pNode.metadata as HeadingMetadata).level = (pNode.metadata as HeadingMetadata).level || 1;
                }

                // Clean up metadata if empty
                if (Object.keys(pNode.metadata || {}).length === 0) delete pNode.metadata;

                if (config.includeRawContent) {
                    pNode.rawContent = node.toString();
                }
                targetArray.push(pNode);
            } else if (node.tagName === "text:h") {
                const level = parseInt(node.getAttribute("text:outline-level") || "1");
                const hContent = parseParagraphContent(node, paragraphStyleMap, styleMap, config);

                const hNode: OfficeContentNode = {
                    type: 'heading',
                    text: hContent.text,
                    children: hContent.children,
                    metadata: {
                        level,
                        ...(hContent.alignment ? { alignment: hContent.alignment } : {}),
                        ...(hContent.style ? { style: hContent.style } : {})
                    }
                };
                if (config.includeRawContent) {
                    hNode.rawContent = node.toString();
                }
                targetArray.push(hNode);
            } else if (node.tagName === "table:table") {
                // Parse table with proper structure
                const tableNode = parseTable(node, paragraphStyleMap, styleMap, config);
                if (config.includeRawContent) {
                    tableNode.rawContent = node.toString();
                }
                targetArray.push(tableNode);
            } else if (node.tagName === "text:list") {
                // Parse list structure with proper listId tracking
                const listItems = getDirectChildren(node, "text:list-item");

                // Get list style name to use as listId (or generate one)
                const listStyleName = node.getAttribute("text:style-name") || node.getAttribute("xml:id");
                const listId = listStyleName || `list-${targetArray.length}`;

                // Determine list type by checking the list style definition
                let listType: 'ordered' | 'unordered' = 'unordered';
                let isVisible = false;
                let styleNameToCheck = listStyleName;

                // If no style name, check parent list for inherited style
                if (!styleNameToCheck) {
                    let parentNode = node.parentNode;
                    while (parentNode && !styleNameToCheck) {
                        if (parentNode.nodeName === 'text:list') {
                            styleNameToCheck = (parentNode as Element).getAttribute("text:style-name");
                            if (styleNameToCheck) break;
                        }
                        parentNode = parentNode.parentNode;
                    }
                }

                // Try to find list style in automatic styles to determine type and visibility
                if (styleNameToCheck) {
                    const automaticStyles = getElementsByTagName(parseXmlString(mainContentFile?.content.toString() || ''), "office:automatic-styles")[0];
                    if (automaticStyles) {
                        const listStyles = getElementsByTagName(automaticStyles, "text:list-style");
                        for (const listStyle of listStyles) {
                            if (listStyle.getAttribute("style:name") === styleNameToCheck) {
                                // Check if it has bullet or number level styles
                                const bulletLevels = getElementsByTagName(listStyle, "text:list-level-style-bullet");
                                const numberLevels = getElementsByTagName(listStyle, "text:list-level-style-number");
                                const imageLevels = getElementsByTagName(listStyle, "text:list-level-style-image");

                                if (numberLevels.length > 0) {
                                    listType = 'ordered';
                                    isVisible = numberLevels.some(l => !!l.getAttribute("style:num-format"));
                                } else if (bulletLevels.length > 0) {
                                    listType = 'unordered';
                                    isVisible = bulletLevels.some(l => !!l.getAttribute("text:bullet-char"));
                                }

                                if (imageLevels.length > 0) isVisible = true;
                                break;
                            }
                        }
                    }

                    // Also check in styles.xml if still unordered and hidden
                    if (stylesFile && !isVisible) {
                        const stylesXml = parseXmlString(stylesFile.content.toString());
                        const listStyles = getElementsByTagName(stylesXml, "text:list-style");
                        for (const listStyle of listStyles) {
                            if (listStyle.getAttribute("style:name") === styleNameToCheck) {
                                const bulletLevels = getElementsByTagName(listStyle, "text:list-level-style-bullet");
                                const numberLevels = getElementsByTagName(listStyle, "text:list-level-style-number");
                                const imageLevels = getElementsByTagName(listStyle, "text:list-level-style-image");

                                if (numberLevels.length > 0) {
                                    listType = 'ordered';
                                    isVisible = numberLevels.some(l => !!l.getAttribute("style:num-format"));
                                } else if (bulletLevels.length > 0) {
                                    listType = 'unordered';
                                    isVisible = bulletLevels.some(l => !!l.getAttribute("text:bullet-char"));
                                }

                                if (imageLevels.length > 0) isVisible = true;
                                break;
                            }
                        }
                    }
                }

                // If the list is not visible, it's likely a layout list used by Impress.
                // We should traverse its items and treat their content as regular nodes.
                if (!isVisible) {
                    for (let i = 0; i < listItems.length; i++) {
                        const item = listItems[i];
                        if (item.childNodes) {
                            for (let j = 0; j < item.childNodes.length; j++) {
                                const child = item.childNodes[j];
                                if (child.nodeType === 1) { // Element
                                    traverse(child as Element, targetArray, forceHeading);
                                }
                            }
                        }
                    }
                    return;
                }

                // Calculate indentation level by counting parent text:list elements
                let indentation = 0;
                let parent = node.parentNode;
                while (parent) {
                    if (parent.nodeName === 'text:list') {
                        indentation++;
                    }
                    parent = parent.parentNode;
                }

                // Track list counters for this listId (similar to WordParser)
                if (!listCounters[listId]) {
                    listCounters[listId] = {};
                }
                const indentKey = indentation.toString();
                if (listCounters[listId][indentKey] === undefined) {
                    listCounters[listId][indentKey] = -1; // Will increment to 0 on first item
                }

                // Process each list item
                for (let i = 0; i < listItems.length; i++) {
                    const item = listItems[i];

                    // Increment item index for this list/level
                    listCounters[listId][indentKey]++;
                    const itemIndex = listCounters[listId][indentKey];

                    // Reset deeper levels when we encounter an item at this level
                    for (let k = indentation + 1; k < 10; k++) {
                        if (listCounters[listId][k.toString()] !== undefined) {
                            listCounters[listId][k.toString()] = -1;
                        }
                    }

                    // Iterate over direct children of list item (paragraphs, headings, nested lists)
                    if (item.childNodes) {
                        for (let j = 0; j < item.childNodes.length; j++) {
                            const child = item.childNodes[j];
                            if (child.nodeType === 1) { // Element
                                const element = child as Element;

                                if (element.tagName === "text:p") {
                                    const pContent = parseParagraphContent(element, paragraphStyleMap, styleMap, config);
                                    const listNode: OfficeContentNode = {
                                        type: 'list',
                                        text: pContent.text,
                                        children: pContent.children,
                                        metadata: {
                                            listType,
                                            indentation,
                                            itemIndex,
                                            listId,
                                            alignment: pContent.alignment || 'left',
                                            style: pContent.style
                                        } as ListMetadata
                                    };
                                    if (config.includeRawContent) listNode.rawContent = element.toString();
                                    targetArray.push(listNode);
                                } else if (element.tagName === "text:h") {
                                    const level = parseInt(element.getAttribute("text:outline-level") || "1");
                                    const hContent = parseParagraphContent(element, paragraphStyleMap, styleMap, config);
                                    const listNode: OfficeContentNode = {
                                        type: 'list',
                                        text: hContent.text,
                                        children: hContent.children,
                                        metadata: {
                                            listType,
                                            indentation,
                                            itemIndex,
                                            listId,
                                            ...(hContent.alignment ? { alignment: hContent.alignment } : {}),
                                            style: hContent.style
                                        }
                                    };
                                    if (config.includeRawContent) listNode.rawContent = element.toString();
                                    targetArray.push(listNode);
                                } else if (element.tagName === "text:list") {
                                    // Recursive call for nested list
                                    traverse(element, targetArray, forceHeading);
                                }
                            }
                        }
                    }
                }
            } else if (node.tagName === "draw:frame") {
                const presClass = node.getAttribute("presentation:class");
                const isHeading = presClass === "title" || presClass === "sub-title";

                // In presentations, frames often contain text-boxes, images, tables, or objects
                const textBox = getElementsByTagName(node, "draw:text-box")[0];
                const image = getElementsByTagName(node, "draw:image")[0];
                const table = getElementsByTagName(node, "table:table")[0];
                const object = getElementsByTagName(node, "draw:object")[0];

                if (textBox) {
                    traverse(textBox, targetArray, isHeading || forceHeading);
                } else if (table) {
                    const tableNode = parseTable(table, paragraphStyleMap, styleMap, config);
                    if (config.includeRawContent) tableNode.rawContent = table.toString();
                    targetArray.push(tableNode);
                } else if (image) {
                    // Extract alt text from svg:title or svg:desc
                    let altText = '';
                    const svgTitle = getElementsByTagName(node, "svg:title")[0];
                    const svgDesc = getElementsByTagName(node, "svg:desc")[0];
                    if (svgTitle && svgTitle.textContent) {
                        altText = svgTitle.textContent;
                    } else if (svgDesc && svgDesc.textContent) {
                        altText = svgDesc.textContent;
                    }

                    // Extract image href to link to attachment
                    let imageHref = image.getAttribute("xlink:href") || '';
                    if (imageHref) {
                        const parts = imageHref.split('/');
                        imageHref = parts[parts.length - 1];
                    }

                    const imageNode: OfficeContentNode = {
                        type: 'image',
                        text: '',
                        children: [],
                        metadata: {
                            attachmentName: imageHref,
                            ...(altText ? { altText } : {})
                        }
                    };
                    if (config.includeRawContent) {
                        imageNode.rawContent = node.toString();
                    }
                    targetArray.push(imageNode);
                } else if (object) {
                    // Handle embedded objects like charts
                    const href = object.getAttribute("xlink:href");
                    if (href) {
                        const attachmentName = href.split('/')[0];
                        const objectPath = `${attachmentName}/content.xml`;
                        const objectFile = files.find(f => f.path === objectPath || f.path.endsWith(objectPath));

                        if (objectFile) {
                            const chartData = extractChartData(objectFile.content);

                            const chartNode: OfficeContentNode = {
                                type: 'chart',
                                text: chartData.rawTexts.join(" "),
                                metadata: {
                                    attachmentName: attachmentName,
                                    chartData
                                } as ChartMetadata
                            };
                            if (config.includeRawContent) chartNode.rawContent = node.toString();
                            targetArray.push(chartNode);
                        } else {
                            const chartNode: OfficeContentNode = {
                                type: 'chart',
                                text: "",
                                metadata: { attachmentName: attachmentName }
                            };
                            if (config.includeRawContent) chartNode.rawContent = node.toString();
                            targetArray.push(chartNode);
                        }
                    }
                }
            } else {
                if (node.childNodes) {
                    for (let i = 0; i < node.childNodes.length; i++) {
                        const child = node.childNodes[i];
                        if (child.nodeType === 1) { // Element
                            traverse(child as Element, targetArray, forceHeading);
                        }
                    }
                }
            }
        };

        // ODS: Spreadsheet
        if (fileType === 'ods') {
            const spreadsheet = getElementsByTagName(body, "office:spreadsheet")[0];
            if (spreadsheet) {
                const tables = getElementsByTagName(spreadsheet, "table:table");
                for (let i = 0; i < tables.length; i++) {
                    const table = tables[i];
                    const sheetName = table.getAttribute("table:name") || `Sheet${i + 1}`;
                    const rows: OfficeContentNode[] = [];

                    const tableRows = getElementsByTagName(table, "table:table-row");
                    let rowIndex = 0;

                    for (let r = 0; r < tableRows.length; r++) {
                        const row = tableRows[r];
                        const cells: OfficeContentNode[] = [];
                        const tableCells = getElementsByTagName(row, "table:table-cell");

                        let colIndex = 0;
                        const rowsRepeated = parseInt(row.getAttribute("table:number-rows-repeated") || "1");

                        for (let c = 0; c < tableCells.length; c++) {
                            const cell = tableCells[c];
                            const colsRepeated = parseInt(cell.getAttribute("table:number-columns-repeated") || "1");

                            // Extract text from cell (paragraphs inside cell)
                            let cellText = "";
                            const children: OfficeContentNode[] = [];
                            const ps = getElementsByTagName(cell, "text:p");
                            for (let p = 0; p < ps.length; p++) {
                                const para = ps[p];

                                // Parse text:span elements for formatted text
                                const spans = getElementsByTagName(para, "text:span");
                                if (spans.length > 0) {
                                    for (const span of spans) {
                                        const styleName = span.getAttribute("text:style-name");
                                        const formatting = styleName ? styleMap[styleName] : {};
                                        const text = span.textContent || '';
                                        cellText += text;

                                        const textNode: OfficeContentNode = {
                                            type: 'text',
                                            text: text,
                                            formatting: formatting
                                        };
                                        children.push(textNode);
                                    }
                                } else {
                                    // No spans - just direct text content
                                    const text = para.textContent || '';
                                    cellText += text;
                                    if (text.trim()) {
                                        const textNode: OfficeContentNode = {
                                            type: 'text',
                                            text: text,
                                            formatting: {}
                                        };
                                        children.push(textNode);
                                    }
                                }

                                if (p < ps.length - 1) cellText += "\n";
                            }

                            // Check for embedded draw:frame (images) in cell
                            const drawFrames = getElementsByTagName(cell, "draw:frame");
                            for (const frame of drawFrames) {
                                // Extract alt text from svg:title or svg:desc
                                let altText = '';
                                const svgTitle = getElementsByTagName(frame, "svg:title")[0];
                                const svgDesc = getElementsByTagName(frame, "svg:desc")[0];
                                if (svgTitle && svgTitle.textContent) {
                                    altText = svgTitle.textContent;
                                } else if (svgDesc && svgDesc.textContent) {
                                    altText = svgDesc.textContent;
                                }

                                // Extract image href
                                let imageHref = '';
                                const drawImages = getElementsByTagName(frame, "draw:image");
                                if (drawImages.length > 0) {
                                    const rawHref = drawImages[0].getAttribute("xlink:href");
                                    if (rawHref) {
                                        const parts = rawHref.split('/');
                                        imageHref = parts[parts.length - 1];
                                    }
                                }

                                // Extract chart object href
                                let chartHref = '';
                                const drawObjects = getElementsByTagName(frame, "draw:object");
                                if (drawObjects.length > 0) {
                                    const href = drawObjects[0].getAttribute("xlink:href");
                                    if (href) {
                                        // Object href is usually "./Object 1"
                                        chartHref = href.split('/')[0];
                                    }
                                }

                                if (drawImages.length > 0) {
                                    // logic for image node
                                    const imageNode: OfficeContentNode = {
                                        type: 'image',
                                        text: '', // Will be populated by assignAttachmentData
                                        children: [],
                                        metadata: {
                                            attachmentName: imageHref, // Might be empty, will resolve in assignAttachmentData
                                            ...(altText ? { altText } : {})
                                        } as ImageMetadata
                                    };
                                    if (config.includeRawContent) {
                                        imageNode.rawContent = frame.toString();
                                    }
                                    children.push(imageNode);
                                } else if (chartHref) {
                                    const chartNode: OfficeContentNode = {
                                        type: 'chart',
                                        text: '', // Will be populated by assignAttachmentData
                                        children: [],
                                        metadata: {
                                            attachmentName: chartHref
                                        } as ChartMetadata
                                    };
                                    children.push(chartNode);
                                }
                            }

                            // Add cell(s)
                            for (let k = 0; k < colsRepeated; k++) {
                                // For ODS (spreadsheets), we skip empty cells to avoid massive ASTs (millions of cells)
                                // but for ODP/ODT (presentation/text), cells are part of a defined table grid
                                // Also include cells that have children (e.g., image nodes) even if no text
                                if (cellText || children.length > 0 || fileType !== 'ods') {
                                    const cellNode: OfficeContentNode = {
                                        type: 'cell',
                                        text: cellText,
                                        children: children,
                                        metadata: { row: rowIndex, col: colIndex } as CellMetadata
                                    };
                                    if (config.includeRawContent) {
                                        cellNode.rawContent = cell.toString();
                                    }
                                    cells.push(cellNode);
                                }
                                colIndex++;
                            }
                        }

                        // Add row(s)
                        if (cells.length > 0) {
                            for (let k = 0; k < rowsRepeated; k++) {
                                const rowNode: OfficeContentNode = {
                                    type: 'row',
                                    children: JSON.parse(JSON.stringify(cells)), // Deep copy for repeated rows
                                    metadata: undefined
                                };
                                // Fix row index in metadata for repeated rows
                                if (k > 0) {
                                    rowNode.children?.forEach(c => {
                                        if (c.metadata && 'row' in c.metadata) {
                                            (c.metadata as CellMetadata).row = rowIndex;
                                        }
                                    });
                                }
                                if (config.includeRawContent) {
                                    rowNode.rawContent = row.toString();
                                }
                                rows.push(rowNode);
                                rowIndex++;
                            }
                        } else {
                            rowIndex += rowsRepeated;
                        }
                    }

                    const sheetNode: OfficeContentNode = {
                        type: 'sheet',
                        children: rows,
                        metadata: { sheetName } as SheetMetadata
                    };
                    if (config.includeRawContent) {
                        sheetNode.rawContent = table.toString();
                    }
                    content.push(sheetNode);
                }
            }
        }
        // ODP: Presentation
        else if (fileType === 'odp') {
            const presentation = getElementsByTagName(body, "office:presentation")[0];
            if (presentation) {
                const pages = getDirectChildren(presentation, "draw:page");
                const odpNotes: OfficeContentNode[] = [];

                for (let i = 0; i < pages.length; i++) {
                    const page = pages[i];
                    const slideNode: OfficeContentNode = {
                        type: 'slide',
                        children: [],
                        metadata: { slideNumber: i + 1 } as SlideMetadata
                    };

                    // Separate page content and notes
                    let noteNode: OfficeContentNode | undefined = undefined;
                    const pageChildren = page.childNodes;
                    if (pageChildren) {
                        for (let j = 0; j < pageChildren.length; j++) {
                            const child = pageChildren[j];
                            if (child.nodeType === 1) { // Element
                                const element = child as Element;
                                if (element.tagName === "presentation:notes") {
                                    if (!config.ignoreNotes) {
                                        noteNode = {
                                            type: 'note',
                                            children: [],
                                            metadata: {
                                                slideNumber: i + 1,
                                                noteId: `slide-note-${i + 1}`
                                            } as SlideMetadata
                                        };
                                        traverse(element, noteNode.children!);
                                    }
                                    continue;
                                }
                                traverse(element, slideNode.children!);
                            }
                        }
                    }

                    if (config.includeRawContent) {
                        slideNode.rawContent = page.toString();
                    }
                    content.push(slideNode);

                    if (noteNode && noteNode.children && noteNode.children.length > 0) {
                        if (config.putNotesAtLast) {
                            odpNotes.push(noteNode);
                        } else {
                            content.push(noteNode);
                        }
                    }
                }

                if (odpNotes.length > 0) {
                    content.push(...odpNotes);
                }
            }
        }
        // ODT: Text Document (and generic fallback)
        else {
            const textDoc = getElementsByTagName(body, "office:text")[0];
            if (textDoc) {
                traverse(textDoc, content);
            }
        }
    };

    if (mainContentFile) {
        parseContentXml(mainContentFile.content.toString());
    }

    // Attachments
    const attachments: OfficeAttachment[] = [];
    const mediaFiles = files.filter(f => f.path.match(/(Pictures|media)\/.*/));

    // ODP/ODT Chart Extraction
    if (config.extractAttachments) {
        const objectFiles = files.filter(f => f.path.match(/Object \d+\/content\.xml/));
        for (const objFile of objectFiles) {
            const objXml = parseXmlString(objFile.content.toString());
            const isChart = getElementsByTagName(objXml, "chart:chart").length > 0;

            if (isChart) {
                const objectId = objFile.path.split('/')[0];
                const attachment: OfficeAttachment = {
                    type: 'chart',
                    mimeType: 'application/vnd.oasis.opendocument.chart', // Accurate ODF chart type
                    data: objFile.content.toString('base64'),
                    name: objectId,
                    extension: 'xml'
                };

                // Extract data from chart XML
                const chartData = extractChartData(objFile.content);

                if (chartData.rawTexts.length > 0) {
                    attachment.chartData = chartData;
                }

                attachments.push(attachment);
            }
        }
    }

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
    }

    const metaFile = files.find(f => f.path.match(metaFileRegex));
    const metadata = metaFile ? parseOfficeMetadata(metaFile.content.toString()) : {};

    // Helper: Resolve ODS chart cell references to actual values
    // ODS charts often link to cell ranges (e.g., [Sheet1.$A$1:.$A$5]) instead of embedding values
    const resolveChartReferences = (chartData: ChartData, nodes: OfficeContentNode[]) => {
        const getValuesFromReference = (ref: string): string[] => {
            // Remove brackets: [Sheet.$A$1:.$A$5] -> Sheet.$A$1:.$A$5
            const cleanRef = ref.replace(/^\[|\]$/g, '');
            const [startPart, endPart] = cleanRef.split(':');

            const lastDotIdx = startPart.lastIndexOf('.');
            if (lastDotIdx === -1) return [ref];

            const sheetName = startPart.substring(0, lastDotIdx).replace(/^'|'$/g, '');
            const startCoord = startPart.substring(lastDotIdx + 1).replace(/\$/g, '');

            let endCoord = startCoord;
            if (endPart) {
                if (endPart.startsWith('.')) {
                    endCoord = endPart.substring(1).replace(/\$/g, '');
                } else {
                    const endLastDotIdx = endPart.lastIndexOf('.');
                    endCoord = endPart.substring(endLastDotIdx + 1).replace(/\$/g, '');
                }
            }

            const parseCoord = (coord: string) => {
                const colMatch = coord.match(/[A-Z]+/);
                const rowMatch = coord.match(/\d+/);
                if (!colMatch || !rowMatch) return null;

                const colStr = colMatch[0];
                let colIdx = 0;
                for (let i = 0; i < colStr.length; i++) {
                    colIdx = colIdx * 26 + (colStr.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
                }
                colIdx -= 1;
                const rowIdx = parseInt(rowMatch[0]) - 1;
                return { r: rowIdx, c: colIdx };
            };

            const start = parseCoord(startCoord);
            const end = parseCoord(endCoord);

            if (!start || !end) return [ref];

            const sheet = nodes.find(n => n.type === 'sheet' && (n.metadata as SheetMetadata)?.sheetName === sheetName);
            if (!sheet || !sheet.children) return [ref];

            const values: string[] = [];
            // Collect all matching cells
            for (const row of sheet.children) {
                if (row.children) {
                    for (const cell of row.children) {
                        const meta = cell.metadata as CellMetadata;
                        if (meta && meta.row >= start.r && meta.row <= end.r && meta.col >= start.c && meta.col <= end.c) {
                            values.push(cell.text || '');
                        }
                    }
                }
            }
            return values.length > 0 ? values : [];
        };

        // Resolve DataSets
        for (const ds of chartData.dataSets) {
            const newValues: string[] = [];
            for (const val of ds.values) {
                if (val.startsWith('[')) newValues.push(...getValuesFromReference(val));
                else newValues.push(val);
            }
            ds.values = newValues;
        }

        // Resolve Labels
        const newLabels: string[] = [];
        for (const label of chartData.labels) {
            if (label.startsWith('[')) newLabels.push(...getValuesFromReference(label));
            else newLabels.push(label);
        }
        chartData.labels = newLabels;

        // Rebuild rawTexts
        chartData.rawTexts = [];
        if (chartData.title) chartData.rawTexts.push(chartData.title);
        for (const ds of chartData.dataSets) {
            if (ds.name) chartData.rawTexts.push(ds.name);
            chartData.rawTexts.push(...chartData.labels);
            chartData.rawTexts.push(...ds.values);
        }
    };

    // Apply resolution to all chart attachments
    for (const att of attachments) {
        if (att.type === 'chart' && att.chartData) {
            resolveChartReferences(att.chartData, content);
        }
    }

    // Link OCR and Chart text to content nodes
    // Link OCR and Chart text to content nodes (with heuristic for unlinked images)
    const assignAttachmentData = (nodes: OfficeContentNode[]) => {
        // Step 1: Identify unused image attachments globally
        const usedAttachmentNames = new Set<string>();
        const traverseForNames = (ns: OfficeContentNode[]) => {
            for (const n of ns) {
                if (n.metadata && 'attachmentName' in n.metadata) {
                    const name = (n.metadata as ImageMetadata).attachmentName;
                    if (name) usedAttachmentNames.add(name);
                }
                if (n.children) traverseForNames(n.children);
            }
        };
        traverseForNames(nodes);

        const unusedImages = attachments.filter(a => a.type === 'image' && a.name && !usedAttachmentNames.has(a.name));
        let unusedImageIndex = 0;

        const processNode = (node: OfficeContentNode) => {
            if ((node.type === 'image' || node.type === 'chart') && node.metadata && 'attachmentName' in node.metadata) {
                let attachmentName = (node.metadata as ImageMetadata).attachmentName;

                // Heuristic: If name is empty, try to assign an unused image attachment
                if (!attachmentName && node.type === 'image' && unusedImageIndex < unusedImages.length) {
                    const fallbackAtt = unusedImages[unusedImageIndex++];
                    attachmentName = fallbackAtt.name;
                    (node.metadata as ImageMetadata).attachmentName = attachmentName;
                }

                if (attachmentName) {
                    const attachment = attachments.find(a => a.name === attachmentName);

                    if (attachment) {
                        if (attachment.ocrText) {
                            node.text = attachment.ocrText;
                        }
                        if (attachment.chartData && node.type === 'chart') {
                            node.text = attachment.chartData.rawTexts.join(config.newlineDelimiter ?? '\n');
                        }
                    }
                }
            }
            // Internal recursion
            if (node.children) {
                node.children.forEach(processNode);
            }
        };

        nodes.forEach(processNode);
    };
    assignAttachmentData(content);

    // Create combined styleMap for metadata (matches DOCX format)
    const combinedStyleMap: { [key: string]: { formatting: TextFormatting, alignment?: 'left' | 'center' | 'right' | 'justify' } } = {};
    for (const styleName in styleMap) {
        combinedStyleMap[styleName] = {
            formatting: styleMap[styleName],
            alignment: paragraphStyleMap[styleName]?.alignment
        };
    }
    // Also add styles that only have alignment
    for (const styleName in paragraphStyleMap) {
        if (!combinedStyleMap[styleName]) {
            combinedStyleMap[styleName] = {
                formatting: {},
                alignment: paragraphStyleMap[styleName]?.alignment
            };
        }
    }

    // Append notes to content if configured
    if (config.putNotesAtLast && notes.length > 0) {
        content.push(...notes);
    }

    return {
        type: fileType,
        metadata: {
            ...metadata,
            styleMap: combinedStyleMap
        },
        content: content,
        attachments: attachments,
        toText: () => content.map(c => {
            const getText = (node: OfficeContentNode): string => {
                let t = '';
                if (node.children && node.children.length > 0) {
                    // Check if children have their own children (container vs leaf)
                    // If children are leaf nodes (text/image), join with empty string
                    // If children are container nodes (paragraphs/rows), join with newline
                    const hasGrandChildren = node.children.some(child => child.children && child.children.length > 0);
                    const separator = hasGrandChildren ? (config.newlineDelimiter ?? '\n') : '';

                    t += node.children.map(getText).filter(t => t != '').join(separator);
                } else {
                    t += node.text || '';
                }
                return t;
            };
            return getText(c);
        }).filter(t => t != '').join(config.newlineDelimiter ?? '\n')
    };
};

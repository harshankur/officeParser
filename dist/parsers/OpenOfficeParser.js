"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.parseOpenOffice = void 0;
const chartUtils_1 = require("../utils/chartUtils");
const errorUtils_1 = require("../utils/errorUtils");
const imageUtils_1 = require("../utils/imageUtils");
const ocrUtils_1 = require("../utils/ocrUtils");
const xmlUtils_1 = require("../utils/xmlUtils");
const zipUtils_1 = require("../utils/zipUtils");
/**
 * Parses an OpenOffice document (.odt, .odp, .ods) and extracts content.
 *
 * @param buffer - The ODF file as a Buffer
 * @param config - Parser configuration
 * @returns A promise resolving to the parsed AST
 */
const parseOpenOffice = async (buffer, config) => {
    const contentFileRegex = /content\.xml/;
    const objectContentFileRegex = /Object \d+\/content\.xml/;
    const mediaFileRegex = /(Pictures|media)\/.*/;
    const metaFileRegex = /meta\.xml/;
    const stylesFileRegex = /styles\.xml/;
    const mimetypeFileRegex = /mimetype/;
    const files = await (0, zipUtils_1.extractFiles)(buffer, x => !!x.match(contentFileRegex) ||
        !!x.match(objectContentFileRegex) ||
        !!x.match(metaFileRegex) ||
        !!x.match(stylesFileRegex) ||
        !!x.match(mimetypeFileRegex) ||
        (!!config.extractAttachments && !!x.match(mediaFileRegex)));
    const mimetypeFile = files.find(f => f.path === 'mimetype');
    let fileType = 'odt'; // Default
    if (mimetypeFile) {
        const mime = mimetypeFile.content.toString().trim();
        if (mime.includes('spreadsheet'))
            fileType = 'ods';
        else if (mime.includes('presentation'))
            fileType = 'odp';
        else if (mime.includes('text'))
            fileType = 'odt';
    }
    const mainContentFile = files.find(f => f.path === 'content.xml') || files.find(f => f.path.match(contentFileRegex));
    const stylesFile = files.find(f => f.path === 'styles.xml');
    const content = [];
    const notes = [];
    // Style Map: styleName -> TextFormatting
    // Inline style parsing (from content.xml automatic styles)
    const styleMap = {};
    const paragraphStyleMap = {};
    const listCounters = {}; // Track item index per listId/level
    // Helper to parse styles
    const parseStyles = (xmlString) => {
        const xml = (0, xmlUtils_1.parseXmlString)(xmlString);
        const styles = (0, xmlUtils_1.getElementsByTagName)(xml, "style:style");
        for (const style of styles) {
            const name = style.getAttribute("style:name");
            if (!name)
                continue;
            const styleInfo = {};
            // Parse paragraph properties for alignment and drop caps
            const paraProps = (0, xmlUtils_1.getElementsByTagName)(style, "style:paragraph-properties")[0];
            if (paraProps) {
                const textAlign = paraProps.getAttribute("fo:text-align");
                if (textAlign) {
                    const alignMap = {
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
                const dropCap = (0, xmlUtils_1.getElementsByTagName)(paraProps, "style:drop-cap")[0];
                if (dropCap) {
                    styleInfo.dropCap = true;
                }
            }
            if (Object.keys(styleInfo).length > 0) {
                paragraphStyleMap[name] = styleInfo;
            }
            // Parse text properties
            const textProps = (0, xmlUtils_1.getElementsByTagName)(style, "style:text-properties")[0];
            // Parse table cell properties (for ODS background)
            const cellProps = (0, xmlUtils_1.getElementsByTagName)(style, "style:table-cell-properties")[0];
            const formatting = {};
            if (cellProps) {
                const bgColor = cellProps.getAttribute("fo:background-color");
                if (bgColor && bgColor !== 'transparent')
                    formatting.backgroundColor = bgColor;
            }
            if (textProps) {
                if (textProps.getAttribute("fo:font-weight") === "bold" || textProps.getAttribute("style:font-weight-asian") === "bold")
                    formatting.bold = true;
                if (textProps.getAttribute("fo:font-style") === "italic" || textProps.getAttribute("style:font-style-asian") === "italic")
                    formatting.italic = true;
                if (textProps.getAttribute("style:text-underline-style") === "solid")
                    formatting.underline = true;
                if (textProps.getAttribute("style:text-line-through-style") === "solid")
                    formatting.strikethrough = true;
                const size = textProps.getAttribute("fo:font-size") || textProps.getAttribute("style:font-size-asian");
                if (size)
                    formatting.size = size;
                const color = textProps.getAttribute("fo:color");
                if (color)
                    formatting.color = color;
                // Background color (text level) - override cell level if present?
                const bgColor = textProps.getAttribute("fo:background-color");
                if (bgColor && bgColor !== 'transparent')
                    formatting.backgroundColor = bgColor;
                // Font family
                const fontName = textProps.getAttribute("style:font-name") || textProps.getAttribute("fo:font-family");
                if (fontName)
                    formatting.font = fontName;
                // Subscript/Superscript from text-position (e.g., "sub 58%" or "super 58%")
                const textPosition = textProps.getAttribute("style:text-position");
                if (textPosition) {
                    if (textPosition.startsWith("sub"))
                        formatting.subscript = true;
                    if (textPosition.startsWith("super"))
                        formatting.superscript = true;
                }
                if (Object.keys(formatting).length > 0)
                    styleMap[name] = formatting;
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
    const parseInlineContent = (node, styleMap, config, notes, paragraphStyleMap, parentFormatting = {}, linkMetadata) => {
        const children = [];
        let fullText = '';
        if (!node.childNodes)
            return { text: '', children: [] };
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
            }
            else if (child.nodeType === 1) {
                const element = child;
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
                }
                else if (tagName === 'text:tab') {
                    // Tab
                    fullText += '\t';
                    children.push({
                        type: 'text',
                        text: '\t',
                        formatting: parentFormatting,
                        metadata: linkMetadata ? { ...linkMetadata } : undefined
                    });
                }
                else if (tagName === 'text:line-break') {
                    // Line break
                    fullText += '\n';
                    children.push({
                        type: 'text',
                        text: '\n',
                        formatting: parentFormatting,
                        metadata: linkMetadata ? { ...linkMetadata } : undefined
                    });
                }
                else if (tagName === 'text:span') {
                    // Formatted text span
                    const styleName = element.getAttribute("text:style-name");
                    const formatting = styleName ? { ...parentFormatting, ...styleMap[styleName] } : parentFormatting;
                    const spanContent = parseInlineContent(element, styleMap, config, notes, paragraphStyleMap, formatting, linkMetadata);
                    fullText += spanContent.text;
                    children.push(...spanContent.children);
                }
                else if (tagName === 'text:a') {
                    // Hyperlink
                    const href = element.getAttribute('xlink:href') || '';
                    const linkType = href.startsWith('#') ? 'internal' : 'external';
                    const newLinkMetadata = { link: href, linkType: linkType };
                    const linkContent = parseInlineContent(element, styleMap, config, notes, paragraphStyleMap, parentFormatting, newLinkMetadata);
                    fullText += linkContent.text;
                    children.push(...linkContent.children);
                }
                else if (tagName === 'text:note' && !config.ignoreNotes) {
                    // Footnote or endnote
                    const noteClass = (element.getAttribute('text:note-class') || 'footnote');
                    const noteId = element.getAttribute('text:id') || element.getAttribute('xml:id') || undefined;
                    const noteBody = (0, xmlUtils_1.getElementsByTagName)(element, "text:note-body")[0];
                    if (noteBody) {
                        // Extract note content recursively
                        const notePs = (0, xmlUtils_1.getElementsByTagName)(noteBody, "text:p");
                        const noteChildren = [];
                        let noteText = '';
                        for (const np of notePs) {
                            const npContent = parseParagraphContent(np, paragraphStyleMap, styleMap, config);
                            noteText += (noteText ? ' ' : '') + npContent.text;
                            const npNode = {
                                type: 'paragraph',
                                text: npContent.text,
                                children: npContent.children,
                                metadata: npContent.alignment ? { alignment: npContent.alignment } : undefined
                            };
                            noteChildren.push(npNode);
                        }
                        const noteNode = {
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
                        }
                        else {
                            children.push(noteNode);
                        }
                    }
                }
                else if (tagName === 'draw:frame') {
                    // Inline image
                    const frame = element;
                    // Extract alt text
                    let altText = '';
                    const svgTitle = (0, xmlUtils_1.getElementsByTagName)(frame, "svg:title")[0];
                    const svgDesc = (0, xmlUtils_1.getElementsByTagName)(frame, "svg:desc")[0];
                    if (svgTitle && svgTitle.textContent) {
                        altText = svgTitle.textContent;
                    }
                    else if (svgDesc && svgDesc.textContent) {
                        altText = svgDesc.textContent;
                    }
                    // Extract image href
                    let imageHref = '';
                    const drawImages = (0, xmlUtils_1.getElementsByTagName)(frame, "draw:image");
                    if (drawImages.length > 0) {
                        imageHref = drawImages[0].getAttribute("xlink:href") || '';
                        if (imageHref) {
                            const parts = imageHref.split('/');
                            imageHref = parts[parts.length - 1];
                        }
                    }
                    const imageNode = {
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
    const parseParagraphContent = (node, paraStyleMap, styleMap, config) => {
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
                    if (!child.metadata)
                        child.metadata = {};
                    // Only add style if it's a text node and doesn't have one?
                    // Or just add it.
                    // Cast to any to avoid union type issues for now, or check type
                    const meta = child.metadata;
                    if (!meta.style)
                        meta.style = paraStyle;
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
                }
                else {
                    // Split text node
                    const firstChar = firstChild.text[0];
                    const restText = firstChild.text.substring(1);
                    const dropCapNode = {
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
    const parseTable = (tableNode, paraStyleMap, styleMap, config) => {
        const rows = [];
        // Use getDirectChildren to avoid nested table rows
        const tableRows = (0, xmlUtils_1.getDirectChildren)(tableNode, "table:table-row");
        let rowIndex = 0;
        for (const row of tableRows) {
            const cells = [];
            // Use getDirectChildren to avoid nested table cells
            const tableCells = (0, xmlUtils_1.getDirectChildren)(row, "table:table-cell");
            const rowsRepeated = parseInt(row.getAttribute("table:number-rows-repeated") || "1");
            let colIndex = 0;
            for (const cell of tableCells) {
                const cellChildren = [];
                let cellTextRef = { value: '' };
                const colsRepeated = parseInt(cell.getAttribute("table:number-columns-repeated") || "1");
                const colSpan = parseInt(cell.getAttribute("table:number-columns-spanned") || "1");
                const rowSpan = parseInt(cell.getAttribute("table:number-rows-spanned") || "1");
                // Helper to recursively process cell children (handles frames, text-boxes, etc. in ODP)
                const processChildren = (node) => {
                    if (!node.childNodes)
                        return;
                    for (let i = 0; i < node.childNodes.length; i++) {
                        const child = node.childNodes[i];
                        if (child.nodeType === 1) { // Element
                            const element = child;
                            if (element.tagName === "text:p" || element.tagName === "text:h") {
                                const pContent = parseParagraphContent(element, paraStyleMap, styleMap, config);
                                const pNode = {
                                    type: element.tagName === "text:h" ? 'heading' : 'paragraph',
                                    text: pContent.text,
                                    children: pContent.children,
                                    metadata: {
                                        ...(pContent.alignment ? { alignment: pContent.alignment } : {}),
                                        ...(pContent.style ? { style: pContent.style } : {})
                                    }
                                };
                                // Clean up metadata if empty
                                if (Object.keys(pNode.metadata || {}).length === 0)
                                    delete pNode.metadata;
                                if (element.tagName === "text:h") {
                                    if (!pNode.metadata)
                                        pNode.metadata = {};
                                    pNode.metadata.level = parseInt(element.getAttribute("text:outline-level") || "1");
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
                            }
                            else if (element.tagName === "table:table") {
                                // Recursive call for nested table
                                const nestedTableNode = parseTable(element, paraStyleMap, styleMap, config);
                                cellChildren.push(nestedTableNode);
                            }
                            else if (element.tagName === "draw:frame" || element.tagName === "draw:text-box") {
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
                    const cellNode = {
                        type: 'cell',
                        text: cellText,
                        children: cellChildren.length > 0 ? (k === 0 ? cellChildren : JSON.parse(JSON.stringify(cellChildren))) : [],
                        metadata: { row: rowIndex, col: colIndex }
                    };
                    const cellMetadata = cellNode.metadata;
                    if (colSpan > 1)
                        cellMetadata.colSpan = colSpan;
                    if (rowSpan > 1)
                        cellMetadata.rowSpan = rowSpan;
                    if (config.includeRawContent) {
                        cellNode.rawContent = cell.toString();
                    }
                    cells.push(cellNode);
                    colIndex++;
                }
            }
            // Add row(s) for repeated rows
            for (let k = 0; k < rowsRepeated; k++) {
                const rowNode = {
                    type: 'row',
                    children: k === 0 ? cells : JSON.parse(JSON.stringify(cells))
                };
                // Fix row indices for repeated rows
                if (k > 0) {
                    rowNode.children?.forEach(c => {
                        if (c.metadata && 'row' in c.metadata) {
                            c.metadata.row = rowIndex;
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
    const parseContentXml = (xmlString) => {
        const xml = (0, xmlUtils_1.parseXmlString)(xmlString);
        const body = (0, xmlUtils_1.getElementsByTagName)(xml, "office:body")[0];
        // Parse automatic styles (local to content.xml)
        const automaticStyles = (0, xmlUtils_1.getElementsByTagName)(xml, "office:automatic-styles")[0];
        if (automaticStyles) {
            const styles = (0, xmlUtils_1.getElementsByTagName)(automaticStyles, "style:style");
            for (const style of styles) {
                const name = style.getAttribute("style:name");
                if (!name)
                    continue;
                // Parse paragraph properties for alignment
                const paraProps = (0, xmlUtils_1.getElementsByTagName)(style, "style:paragraph-properties")[0];
                const styleInfo = {};
                if (paraProps) {
                    const textAlign = paraProps.getAttribute("fo:text-align");
                    if (textAlign) {
                        const alignMap = {
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
                    const dropCap = (0, xmlUtils_1.getElementsByTagName)(paraProps, "style:drop-cap")[0];
                    if (dropCap)
                        styleInfo.dropCap = true;
                }
                if (Object.keys(styleInfo).length > 0) {
                    paragraphStyleMap[name] = styleInfo;
                }
                const textProps = (0, xmlUtils_1.getElementsByTagName)(style, "style:text-properties")[0];
                if (textProps) {
                    const formatting = {};
                    if (textProps.getAttribute("fo:font-weight") === "bold" || textProps.getAttribute("style:font-weight-asian") === "bold")
                        formatting.bold = true;
                    if (textProps.getAttribute("fo:font-style") === "italic" || textProps.getAttribute("style:font-style-asian") === "italic")
                        formatting.italic = true;
                    if (textProps.getAttribute("style:text-underline-style") === "solid")
                        formatting.underline = true;
                    if (textProps.getAttribute("style:text-line-through-style") === "solid")
                        formatting.strikethrough = true;
                    const size = textProps.getAttribute("fo:font-size") || textProps.getAttribute("style:font-size-asian");
                    if (size)
                        formatting.size = size;
                    const color = textProps.getAttribute("fo:color");
                    if (color)
                        formatting.color = color;
                    // Background color
                    const bgColor = textProps.getAttribute("fo:background-color");
                    if (bgColor && bgColor !== 'transparent')
                        formatting.backgroundColor = bgColor;
                    // Font family
                    const fontName = textProps.getAttribute("style:font-name") || textProps.getAttribute("fo:font-family");
                    if (fontName)
                        formatting.font = fontName;
                    // Subscript/Superscript from text-position (e.g., "sub 58%" or "super 58%")
                    const textPosition = textProps.getAttribute("style:text-position");
                    if (textPosition) {
                        if (textPosition.startsWith("sub"))
                            formatting.subscript = true;
                        if (textPosition.startsWith("super"))
                            formatting.superscript = true;
                    }
                    if (Object.keys(formatting).length > 0)
                        styleMap[name] = formatting;
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
        const traverse = (node, targetArray, forceHeading = false) => {
            if (node.tagName === "text:p") {
                const pContent = parseParagraphContent(node, paragraphStyleMap, styleMap, config);
                const type = (forceHeading || (node.getAttribute("text:style-name") || '').toLowerCase().includes('title')) ? 'heading' : 'paragraph';
                const pNode = {
                    type,
                    text: pContent.text,
                    children: pContent.children,
                    metadata: {
                        ...(pContent.alignment ? { alignment: pContent.alignment } : {}),
                        ...(pContent.style ? { style: pContent.style } : {})
                    }
                };
                if (type === 'heading' && pNode.metadata) {
                    pNode.metadata.level = pNode.metadata.level || 1;
                }
                // Clean up metadata if empty
                if (Object.keys(pNode.metadata || {}).length === 0)
                    delete pNode.metadata;
                if (config.includeRawContent) {
                    pNode.rawContent = node.toString();
                }
                targetArray.push(pNode);
            }
            else if (node.tagName === "text:h") {
                const level = parseInt(node.getAttribute("text:outline-level") || "1");
                const hContent = parseParagraphContent(node, paragraphStyleMap, styleMap, config);
                const hNode = {
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
            }
            else if (node.tagName === "table:table") {
                // Parse table with proper structure
                const tableNode = parseTable(node, paragraphStyleMap, styleMap, config);
                if (config.includeRawContent) {
                    tableNode.rawContent = node.toString();
                }
                targetArray.push(tableNode);
            }
            else if (node.tagName === "text:list") {
                // Parse list structure with proper listId tracking
                const listItems = (0, xmlUtils_1.getDirectChildren)(node, "text:list-item");
                // Get list style name to use as listId (or generate one)
                const listStyleName = node.getAttribute("text:style-name") || node.getAttribute("xml:id");
                const listId = listStyleName || `list-${targetArray.length}`;
                // Determine list type by checking the list style definition
                let listType = 'unordered';
                let isVisible = false;
                let styleNameToCheck = listStyleName;
                // If no style name, check parent list for inherited style
                if (!styleNameToCheck) {
                    let parentNode = node.parentNode;
                    while (parentNode && !styleNameToCheck) {
                        if (parentNode.nodeName === 'text:list') {
                            styleNameToCheck = parentNode.getAttribute("text:style-name");
                            if (styleNameToCheck)
                                break;
                        }
                        parentNode = parentNode.parentNode;
                    }
                }
                // Try to find list style in automatic styles to determine type and visibility
                if (styleNameToCheck) {
                    const automaticStyles = (0, xmlUtils_1.getElementsByTagName)((0, xmlUtils_1.parseXmlString)(mainContentFile?.content.toString() || ''), "office:automatic-styles")[0];
                    if (automaticStyles) {
                        const listStyles = (0, xmlUtils_1.getElementsByTagName)(automaticStyles, "text:list-style");
                        for (const listStyle of listStyles) {
                            if (listStyle.getAttribute("style:name") === styleNameToCheck) {
                                // Check if it has bullet or number level styles
                                const bulletLevels = (0, xmlUtils_1.getElementsByTagName)(listStyle, "text:list-level-style-bullet");
                                const numberLevels = (0, xmlUtils_1.getElementsByTagName)(listStyle, "text:list-level-style-number");
                                const imageLevels = (0, xmlUtils_1.getElementsByTagName)(listStyle, "text:list-level-style-image");
                                if (numberLevels.length > 0) {
                                    listType = 'ordered';
                                    isVisible = numberLevels.some(l => !!l.getAttribute("style:num-format"));
                                }
                                else if (bulletLevels.length > 0) {
                                    listType = 'unordered';
                                    isVisible = bulletLevels.some(l => !!l.getAttribute("text:bullet-char"));
                                }
                                if (imageLevels.length > 0)
                                    isVisible = true;
                                break;
                            }
                        }
                    }
                    // Also check in styles.xml if still unordered and hidden
                    if (stylesFile && !isVisible) {
                        const stylesXml = (0, xmlUtils_1.parseXmlString)(stylesFile.content.toString());
                        const listStyles = (0, xmlUtils_1.getElementsByTagName)(stylesXml, "text:list-style");
                        for (const listStyle of listStyles) {
                            if (listStyle.getAttribute("style:name") === styleNameToCheck) {
                                const bulletLevels = (0, xmlUtils_1.getElementsByTagName)(listStyle, "text:list-level-style-bullet");
                                const numberLevels = (0, xmlUtils_1.getElementsByTagName)(listStyle, "text:list-level-style-number");
                                const imageLevels = (0, xmlUtils_1.getElementsByTagName)(listStyle, "text:list-level-style-image");
                                if (numberLevels.length > 0) {
                                    listType = 'ordered';
                                    isVisible = numberLevels.some(l => !!l.getAttribute("style:num-format"));
                                }
                                else if (bulletLevels.length > 0) {
                                    listType = 'unordered';
                                    isVisible = bulletLevels.some(l => !!l.getAttribute("text:bullet-char"));
                                }
                                if (imageLevels.length > 0)
                                    isVisible = true;
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
                                    traverse(child, targetArray, forceHeading);
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
                                const element = child;
                                if (element.tagName === "text:p") {
                                    const pContent = parseParagraphContent(element, paragraphStyleMap, styleMap, config);
                                    const listNode = {
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
                                        }
                                    };
                                    if (config.includeRawContent)
                                        listNode.rawContent = element.toString();
                                    targetArray.push(listNode);
                                }
                                else if (element.tagName === "text:h") {
                                    const level = parseInt(element.getAttribute("text:outline-level") || "1");
                                    const hContent = parseParagraphContent(element, paragraphStyleMap, styleMap, config);
                                    const listNode = {
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
                                    if (config.includeRawContent)
                                        listNode.rawContent = element.toString();
                                    targetArray.push(listNode);
                                }
                                else if (element.tagName === "text:list") {
                                    // Recursive call for nested list
                                    traverse(element, targetArray, forceHeading);
                                }
                            }
                        }
                    }
                }
            }
            else if (node.tagName === "draw:frame") {
                const presClass = node.getAttribute("presentation:class");
                const isHeading = presClass === "title" || presClass === "sub-title";
                // In presentations, frames often contain text-boxes, images, tables, or objects
                const textBox = (0, xmlUtils_1.getElementsByTagName)(node, "draw:text-box")[0];
                const image = (0, xmlUtils_1.getElementsByTagName)(node, "draw:image")[0];
                const table = (0, xmlUtils_1.getElementsByTagName)(node, "table:table")[0];
                const object = (0, xmlUtils_1.getElementsByTagName)(node, "draw:object")[0];
                if (textBox) {
                    traverse(textBox, targetArray, isHeading || forceHeading);
                }
                else if (table) {
                    const tableNode = parseTable(table, paragraphStyleMap, styleMap, config);
                    if (config.includeRawContent)
                        tableNode.rawContent = table.toString();
                    targetArray.push(tableNode);
                }
                else if (image) {
                    // Extract alt text from svg:title or svg:desc
                    let altText = '';
                    const svgTitle = (0, xmlUtils_1.getElementsByTagName)(node, "svg:title")[0];
                    const svgDesc = (0, xmlUtils_1.getElementsByTagName)(node, "svg:desc")[0];
                    if (svgTitle && svgTitle.textContent) {
                        altText = svgTitle.textContent;
                    }
                    else if (svgDesc && svgDesc.textContent) {
                        altText = svgDesc.textContent;
                    }
                    // Extract image href to link to attachment
                    let imageHref = image.getAttribute("xlink:href") || '';
                    if (imageHref) {
                        const parts = imageHref.split('/');
                        imageHref = parts[parts.length - 1];
                    }
                    const imageNode = {
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
                }
                else if (object) {
                    // Handle embedded objects like charts
                    const href = object.getAttribute("xlink:href");
                    if (href) {
                        const attachmentName = href.split('/')[0];
                        const objectPath = `${attachmentName}/content.xml`;
                        const objectFile = files.find(f => f.path === objectPath || f.path.endsWith(objectPath));
                        if (objectFile) {
                            const chartData = (0, chartUtils_1.extractChartData)(objectFile.content);
                            const chartNode = {
                                type: 'chart',
                                text: chartData.rawTexts.join(" "),
                                metadata: {
                                    attachmentName: attachmentName,
                                    chartData
                                }
                            };
                            if (config.includeRawContent)
                                chartNode.rawContent = node.toString();
                            targetArray.push(chartNode);
                        }
                        else {
                            const chartNode = {
                                type: 'chart',
                                text: "",
                                metadata: { attachmentName: attachmentName }
                            };
                            if (config.includeRawContent)
                                chartNode.rawContent = node.toString();
                            targetArray.push(chartNode);
                        }
                    }
                }
            }
            else {
                if (node.childNodes) {
                    for (let i = 0; i < node.childNodes.length; i++) {
                        const child = node.childNodes[i];
                        if (child.nodeType === 1) { // Element
                            traverse(child, targetArray, forceHeading);
                        }
                    }
                }
            }
        };
        // ODS: Spreadsheet
        if (fileType === 'ods') {
            const spreadsheet = (0, xmlUtils_1.getElementsByTagName)(body, "office:spreadsheet")[0];
            if (spreadsheet) {
                const tables = (0, xmlUtils_1.getElementsByTagName)(spreadsheet, "table:table");
                for (let i = 0; i < tables.length; i++) {
                    const table = tables[i];
                    const sheetName = table.getAttribute("table:name") || `Sheet${i + 1}`;
                    const rows = [];
                    const tableRows = (0, xmlUtils_1.getElementsByTagName)(table, "table:table-row");
                    let rowIndex = 0;
                    for (let r = 0; r < tableRows.length; r++) {
                        const row = tableRows[r];
                        const cells = [];
                        const tableCells = (0, xmlUtils_1.getElementsByTagName)(row, "table:table-cell");
                        let colIndex = 0;
                        const rowsRepeated = parseInt(row.getAttribute("table:number-rows-repeated") || "1");
                        for (let c = 0; c < tableCells.length; c++) {
                            const cell = tableCells[c];
                            const colsRepeated = parseInt(cell.getAttribute("table:number-columns-repeated") || "1");
                            // Extract text from cell (paragraphs inside cell)
                            let cellText = "";
                            const children = [];
                            const ps = (0, xmlUtils_1.getElementsByTagName)(cell, "text:p");
                            for (let p = 0; p < ps.length; p++) {
                                const para = ps[p];
                                // Parse text:span elements for formatted text
                                const spans = (0, xmlUtils_1.getElementsByTagName)(para, "text:span");
                                if (spans.length > 0) {
                                    for (const span of spans) {
                                        const styleName = span.getAttribute("text:style-name");
                                        const formatting = styleName ? styleMap[styleName] : {};
                                        const text = span.textContent || '';
                                        cellText += text;
                                        const textNode = {
                                            type: 'text',
                                            text: text,
                                            formatting: formatting
                                        };
                                        children.push(textNode);
                                    }
                                }
                                else {
                                    // No spans - just direct text content
                                    const text = para.textContent || '';
                                    cellText += text;
                                    if (text.trim()) {
                                        const textNode = {
                                            type: 'text',
                                            text: text,
                                            formatting: {}
                                        };
                                        children.push(textNode);
                                    }
                                }
                                if (p < ps.length - 1)
                                    cellText += "\n";
                            }
                            // Check for embedded draw:frame (images) in cell
                            const drawFrames = (0, xmlUtils_1.getElementsByTagName)(cell, "draw:frame");
                            for (const frame of drawFrames) {
                                // Extract alt text from svg:title or svg:desc
                                let altText = '';
                                const svgTitle = (0, xmlUtils_1.getElementsByTagName)(frame, "svg:title")[0];
                                const svgDesc = (0, xmlUtils_1.getElementsByTagName)(frame, "svg:desc")[0];
                                if (svgTitle && svgTitle.textContent) {
                                    altText = svgTitle.textContent;
                                }
                                else if (svgDesc && svgDesc.textContent) {
                                    altText = svgDesc.textContent;
                                }
                                // Extract image href
                                let imageHref = '';
                                const drawImages = (0, xmlUtils_1.getElementsByTagName)(frame, "draw:image");
                                if (drawImages.length > 0) {
                                    const rawHref = drawImages[0].getAttribute("xlink:href");
                                    if (rawHref) {
                                        const parts = rawHref.split('/');
                                        imageHref = parts[parts.length - 1];
                                    }
                                }
                                // Extract chart object href
                                let chartHref = '';
                                const drawObjects = (0, xmlUtils_1.getElementsByTagName)(frame, "draw:object");
                                if (drawObjects.length > 0) {
                                    const href = drawObjects[0].getAttribute("xlink:href");
                                    if (href) {
                                        // Object href is usually "./Object 1"
                                        chartHref = href.split('/')[0];
                                    }
                                }
                                if (drawImages.length > 0) {
                                    // logic for image node
                                    const imageNode = {
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
                                else if (chartHref) {
                                    const chartNode = {
                                        type: 'chart',
                                        text: '',
                                        children: [],
                                        metadata: {
                                            attachmentName: chartHref
                                        }
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
                                    const cellNode = {
                                        type: 'cell',
                                        text: cellText,
                                        children: children,
                                        metadata: { row: rowIndex, col: colIndex }
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
                                const rowNode = {
                                    type: 'row',
                                    children: JSON.parse(JSON.stringify(cells)),
                                    metadata: undefined
                                };
                                // Fix row index in metadata for repeated rows
                                if (k > 0) {
                                    rowNode.children?.forEach(c => {
                                        if (c.metadata && 'row' in c.metadata) {
                                            c.metadata.row = rowIndex;
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
                        else {
                            rowIndex += rowsRepeated;
                        }
                    }
                    const sheetNode = {
                        type: 'sheet',
                        children: rows,
                        metadata: { sheetName }
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
            const presentation = (0, xmlUtils_1.getElementsByTagName)(body, "office:presentation")[0];
            if (presentation) {
                const pages = (0, xmlUtils_1.getDirectChildren)(presentation, "draw:page");
                const odpNotes = [];
                for (let i = 0; i < pages.length; i++) {
                    const page = pages[i];
                    const slideNode = {
                        type: 'slide',
                        children: [],
                        metadata: { slideNumber: i + 1 }
                    };
                    // Separate page content and notes
                    let noteNode = undefined;
                    const pageChildren = page.childNodes;
                    if (pageChildren) {
                        for (let j = 0; j < pageChildren.length; j++) {
                            const child = pageChildren[j];
                            if (child.nodeType === 1) { // Element
                                const element = child;
                                if (element.tagName === "presentation:notes") {
                                    if (!config.ignoreNotes) {
                                        noteNode = {
                                            type: 'note',
                                            children: [],
                                            metadata: {
                                                slideNumber: i + 1,
                                                noteId: `slide-note-${i + 1}`
                                            }
                                        };
                                        traverse(element, noteNode.children);
                                    }
                                    continue;
                                }
                                traverse(element, slideNode.children);
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
                        }
                        else {
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
            const textDoc = (0, xmlUtils_1.getElementsByTagName)(body, "office:text")[0];
            if (textDoc) {
                traverse(textDoc, content);
            }
        }
    };
    if (mainContentFile) {
        parseContentXml(mainContentFile.content.toString());
    }
    // Attachments
    const attachments = [];
    const mediaFiles = files.filter(f => f.path.match(/(Pictures|media)\/.*/));
    // ODP/ODT Chart Extraction
    if (config.extractAttachments) {
        const objectFiles = files.filter(f => f.path.match(/Object \d+\/content\.xml/));
        for (const objFile of objectFiles) {
            const objXml = (0, xmlUtils_1.parseXmlString)(objFile.content.toString());
            const isChart = (0, xmlUtils_1.getElementsByTagName)(objXml, "chart:chart").length > 0;
            if (isChart) {
                const objectId = objFile.path.split('/')[0];
                const attachment = {
                    type: 'chart',
                    mimeType: 'application/vnd.oasis.opendocument.chart',
                    data: objFile.content.toString('base64'),
                    name: objectId,
                    extension: 'xml'
                };
                // Extract data from chart XML
                const chartData = (0, chartUtils_1.extractChartData)(objFile.content);
                if (chartData.rawTexts.length > 0) {
                    attachment.chartData = chartData;
                }
                attachments.push(attachment);
            }
        }
    }
    if (config.extractAttachments) {
        for (const media of mediaFiles) {
            const attachment = (0, imageUtils_1.createAttachment)(media.path.split('/').pop() || 'image', media.content);
            attachments.push(attachment);
            if (config.ocr) {
                if (attachment.mimeType.startsWith('image/')) {
                    try {
                        attachment.ocrText = (await (0, ocrUtils_1.performOcr)(media.content, config.ocrLanguage)).trim();
                    }
                    catch (e) {
                        (0, errorUtils_1.logWarning)(`OCR failed for ${attachment.name}:`, config, e);
                    }
                }
            }
        }
    }
    const metaFile = files.find(f => f.path.match(metaFileRegex));
    const metadata = metaFile ? (0, xmlUtils_1.parseOfficeMetadata)(metaFile.content.toString()) : {};
    // Helper: Resolve ODS chart cell references to actual values
    // ODS charts often link to cell ranges (e.g., [Sheet1.$A$1:.$A$5]) instead of embedding values
    const resolveChartReferences = (chartData, nodes) => {
        const getValuesFromReference = (ref) => {
            // Remove brackets: [Sheet.$A$1:.$A$5] -> Sheet.$A$1:.$A$5
            const cleanRef = ref.replace(/^\[|\]$/g, '');
            const [startPart, endPart] = cleanRef.split(':');
            const lastDotIdx = startPart.lastIndexOf('.');
            if (lastDotIdx === -1)
                return [ref];
            const sheetName = startPart.substring(0, lastDotIdx).replace(/^'|'$/g, '');
            const startCoord = startPart.substring(lastDotIdx + 1).replace(/\$/g, '');
            let endCoord = startCoord;
            if (endPart) {
                if (endPart.startsWith('.')) {
                    endCoord = endPart.substring(1).replace(/\$/g, '');
                }
                else {
                    const endLastDotIdx = endPart.lastIndexOf('.');
                    endCoord = endPart.substring(endLastDotIdx + 1).replace(/\$/g, '');
                }
            }
            const parseCoord = (coord) => {
                const colMatch = coord.match(/[A-Z]+/);
                const rowMatch = coord.match(/\d+/);
                if (!colMatch || !rowMatch)
                    return null;
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
            if (!start || !end)
                return [ref];
            const sheet = nodes.find(n => n.type === 'sheet' && n.metadata?.sheetName === sheetName);
            if (!sheet || !sheet.children)
                return [ref];
            const values = [];
            // Collect all matching cells
            for (const row of sheet.children) {
                if (row.children) {
                    for (const cell of row.children) {
                        const meta = cell.metadata;
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
            const newValues = [];
            for (const val of ds.values) {
                if (val.startsWith('['))
                    newValues.push(...getValuesFromReference(val));
                else
                    newValues.push(val);
            }
            ds.values = newValues;
        }
        // Resolve Labels
        const newLabels = [];
        for (const label of chartData.labels) {
            if (label.startsWith('['))
                newLabels.push(...getValuesFromReference(label));
            else
                newLabels.push(label);
        }
        chartData.labels = newLabels;
        // Rebuild rawTexts
        chartData.rawTexts = [];
        if (chartData.title)
            chartData.rawTexts.push(chartData.title);
        for (const ds of chartData.dataSets) {
            if (ds.name)
                chartData.rawTexts.push(ds.name);
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
    const assignAttachmentData = (nodes) => {
        // Step 1: Identify unused image attachments globally
        const usedAttachmentNames = new Set();
        const traverseForNames = (ns) => {
            for (const n of ns) {
                if (n.metadata && 'attachmentName' in n.metadata) {
                    const name = n.metadata.attachmentName;
                    if (name)
                        usedAttachmentNames.add(name);
                }
                if (n.children)
                    traverseForNames(n.children);
            }
        };
        traverseForNames(nodes);
        const unusedImages = attachments.filter(a => a.type === 'image' && a.name && !usedAttachmentNames.has(a.name));
        let unusedImageIndex = 0;
        const processNode = (node) => {
            if ((node.type === 'image' || node.type === 'chart') && node.metadata && 'attachmentName' in node.metadata) {
                let attachmentName = node.metadata.attachmentName;
                // Heuristic: If name is empty, try to assign an unused image attachment
                if (!attachmentName && node.type === 'image' && unusedImageIndex < unusedImages.length) {
                    const fallbackAtt = unusedImages[unusedImageIndex++];
                    attachmentName = fallbackAtt.name;
                    node.metadata.attachmentName = attachmentName;
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
    const combinedStyleMap = {};
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
    /**
     * Converts a table node to a TableBlock.
     */
    const convertTableToBlock = (tableNode) => {
        const rows = [];
        if (tableNode.children) {
            for (const rowNode of tableNode.children) {
                if (rowNode.type === 'row' && rowNode.children) {
                    const cols = [];
                    for (const cellNode of rowNode.children) {
                        if (cellNode.type === 'cell') {
                            // Extract text from cell (including nested content)
                            const getCellText = (node) => {
                                let text = node.text || '';
                                if (node.children && node.children.length > 0) {
                                    const childTexts = node.children.map(getCellText).filter(t => t !== '');
                                    if (childTexts.length > 0) {
                                        text += (text ? ' ' : '') + childTexts.join(' ');
                                    }
                                }
                                return text;
                            };
                            const cellText = getCellText(cellNode);
                            cols.push({ value: cellText });
                        }
                    }
                    if (cols.length > 0) {
                        rows.push({ cols });
                    }
                }
            }
        }
        return {
            type: 'table',
            rows
        };
    };
    /**
     * Converts a chart node to a ChartBlock.
     */
    const convertChartToBlock = (chartNode, attachments) => {
        if (chartNode.type !== 'chart')
            return null;
        const chartMetadata = chartNode.metadata;
        // Try to get chartData from metadata first (if it was stored there)
        const metadataWithChartData = chartMetadata;
        if (metadataWithChartData?.chartData) {
            return {
                type: 'chart',
                chartData: metadataWithChartData.chartData,
                chartType: metadataWithChartData.chartData.chartType
            };
        }
        // Otherwise, try to find it from attachments
        if (chartMetadata?.attachmentName) {
            const attachment = attachments.find(a => a.name === chartMetadata.attachmentName && a.type === 'chart');
            if (attachment?.chartData) {
                return {
                    type: 'chart',
                    chartData: attachment.chartData,
                    chartType: attachment.chartData.chartType
                };
            }
        }
        return null;
    };
    /**
     * Extracts blocks from content nodes in document order.
     */
    const extractBlocksFromContent = (nodes, attachments) => {
        const blocks = [];
        const traverse = (node) => {
            // Process node based on type
            if (node.type === 'table') {
                blocks.push(convertTableToBlock(node));
            }
            else if (node.type === 'chart') {
                const chartBlock = convertChartToBlock(node, attachments);
                if (chartBlock) {
                    blocks.push(chartBlock);
                }
            }
            else if (node.type === 'image') {
                const imageMetadata = node.metadata;
                const attachmentName = imageMetadata?.attachmentName;
                if (attachmentName) {
                    const attachment = attachments.find(a => a.name === attachmentName && a.type === 'image');
                    if (attachment) {
                        const buffer = Buffer.from(attachment.data, 'base64');
                        blocks.push({
                            type: 'image',
                            buffer,
                            mimeType: attachment.mimeType,
                            filename: attachment.name
                        });
                    }
                }
            }
            else if (node.text && node.text.trim()) {
                // Create text block for nodes with text content
                // Only create text blocks for leaf nodes or nodes with meaningful text
                if (!node.children || node.children.length === 0 || node.text.trim()) {
                    blocks.push({
                        type: 'text',
                        content: node.text
                    });
                }
            }
            // Recursively process children
            if (node.children) {
                for (const child of node.children) {
                    traverse(child);
                }
            }
        };
        for (const node of nodes) {
            traverse(node);
        }
        return blocks;
    };
    /**
     * Extracts images list from attachments.
     */
    const extractImagesList = (attachments) => {
        return attachments
            .filter(att => att.type === 'image')
            .map(att => ({
            buffer: Buffer.from(att.data, 'base64'),
            mimeType: att.mimeType,
            filename: att.name
        }));
    };
    // Generate fullText
    const fullText = content.map(c => {
        const getText = (node) => {
            let t = '';
            if (node.children && node.children.length > 0) {
                // Check if children have their own children (container vs leaf)
                // If children are leaf nodes (text/image), join with empty string
                // If children are container nodes (paragraphs/rows), join with newline
                const hasGrandChildren = node.children.some(child => child.children && child.children.length > 0);
                const separator = hasGrandChildren ? (config.newlineDelimiter ?? '\n') : '';
                t += node.children.map(getText).filter(t => t != '').join(separator);
            }
            else {
                t += node.text || '';
            }
            return t;
        };
        return getText(c);
    }).filter(t => t != '').join(config.newlineDelimiter ?? '\n');
    // Extract blocks
    const blocks = extractBlocksFromContent(content, attachments);
    // Extract images
    const images = extractImagesList(attachments);
    return {
        type: fileType,
        metadata: {
            ...metadata,
            styleMap: combinedStyleMap
        },
        content: content,
        attachments: attachments,
        fullText,
        blocks,
        images,
        toText: () => fullText
    };
};
exports.parseOpenOffice = parseOpenOffice;

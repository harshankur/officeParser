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

import { attachmentLookup } from '../utils/repeatUtils.js';
import { BreakMetadata, CellMetadata, CodeMetadata, CommentMetadata, FullOfficeParserConfig, ImageMetadata, IndentationMetadata, ListMetadata, OfficeAttachment, OfficeContentNode, OfficeParserAST, TextFormatting, TextMetadata } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { createAttachment } from '../utils/imageUtils.js';
import { isEmptyMath, ommlToLatex } from '../utils/mathUtils.js';
import { ocrDuringParse } from '../utils/ocrUtils.js';
import { getChildElements, getDirectChildren, getElementsByTagName, getOutermostElements, getFirstElementByTagName, getRawContent, isElement, parseOfficeMetadata, parseOOXMLAppProperties, parseOOXMLCustomProperties, parseXmlString } from '../utils/xmlUtils.js';
import { extractFiles, findRequiredPart } from '../utils/zipUtils.js';
import { lookupTable, plainRecord, setOwn } from '../utils/lookupUtils.js';
import { cellSpan, MAX_COL_SPAN } from '../utils/numberUtils.js';
import { appendAll } from '../utils/nodeListUtils.js';

/** Where a text box's paragraph is read into: the paragraph drawing the text box (see parseParagraph). */
interface ParagraphSink {
    children: OfficeContentNode[];
    notes: OfficeContentNode[];
    comments: OfficeContentNode[];
    anchorIds: string[];
    insideLink: boolean;
}

/** A text box's content, which the text box's own paragraphs are parsed from (see ownFirst). */
const TEXT_BOX: ReadonlySet<string> = new Set(['w:txbxContent', 'txbxContent']);

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
    // Honour cancellation requests immediately, before opening the ZIP archive, loading XML
    // files, or kicking off any OCR work.  DOCX files can be large and the inflate + XML-parse
    // steps are synchronous-heavy, so failing fast here avoids wasted CPU time.
    checkAbortSignal(config.abortSignal);

    const documentFileRegex = /word\/document[\d+]?.xml/;
    const footnotesFileRegex = /word\/footnotes[\d+]?.xml/;
    const endnotesFileRegex = /word\/endnotes[\d+]?.xml/;
    const commentsFileRegex = /word\/comments[\d+]?.xml/;
    // Headers and footers are the only parts a document can have many of: Word writes up to
    // three per section (default, first page, even pages), so a handful of sections is enough
    // to reach header10.xml. The single-character form the other parts use stops matching at
    // nine, which would drop those later files as silently as not extracting them at all.
    const headerFileRegex = /word\/header\d*\.xml/;
    const footerFileRegex = /word\/footer\d*\.xml/;
    const numberingFileRegex = /word\/numbering[\d+]?.xml/;
    const mediaFileRegex = /(word\/)?media\/.*/;
    const corePropsFileRegex = /docProps\/core[\d+]?.xml/;
    const customPropsFileRegex = /docProps\/custom\.xml/;
    const appPropsFileRegex = /docProps\/app[\d+]?.xml/;
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
                const colorMap: { [key: string]: string } = lookupTable({
                    'yellow': '#FFFF00', 'green': '#00FF00', 'cyan': '#00FFFF', 'magenta': '#FF00FF',
                    'blue': '#0000FF', 'red': '#FF0000', 'darkBlue': '#00008B', 'darkCyan': '#008B8B',
                    'darkGreen': '#006400', 'darkMagenta': '#8B008B', 'darkRed': '#8B0000',
                    'darkYellow': '#808000', 'darkGray': '#A9A9A9', 'lightGray': '#D3D3D3', 'black': '#000000'
                });
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
     * The first `tag` in `element` outside the text boxes it draws: parsing a text box's paragraphs reads
     * what those hold, and a run finding a text box's picture or note reference as its own showed it twice.
     */
    const ownFirst = (element: Element, tag: string): Element | undefined => getOutermostElements(element, tag, TEXT_BOX)[0];

    /**
     * The child elements of `container` as Word lays them out: a content control (`w:sdt`) or custom XML
     * element (`w:customXml`) stands for what it wraps, at any depth. A table of contents, a cover page or
     * a form is a content control around paragraphs, rows or cells, and reading direct children alone
     * dropped all of it.
     */
    const layoutChildren = (container: Element): Element[] => {
        const out: Element[] = [];
        const pending: Node[] = [];
        const pushChildren = (element: Element) => { for (let i = element.childNodes.length - 1; i >= 0; i--) pending.push(element.childNodes[i]); };
        pushChildren(container);
        while (pending.length) {
            const node = pending.pop()!;
            if (!isElement(node)) continue;
            if (node.nodeName === 'w:sdt') {
                const inner = getDirectChildren(node, "w:sdtContent")[0];
                if (inner) pushChildren(inner);
            } else if (node.nodeName === 'w:customXml') {
                pushChildren(node);
            } else {
                out.push(node);
            }
        }
        return out;
    };

    /**
     * Resolves mc:AlternateContent by preferring mc:Fallback if choice namespace is not recognized,
     * or simply the first available valid child.
     */
    const resolveAlternateContent = (element: Element): Node[] => {
        // Its own Choice and Fallback (children, as the schema has them): looked up through the whole
        // subtree, AlternateContent nested in AlternateContent with neither took time in the product of
        // the depth and the content (4.9 KB, a minute and a half).
        const choice = getDirectChildren(element, "mc:Choice")[0];
        // In most cases, mc:Choice contains the modern version, but mc:Fallback is safer for legacy compatibility
        // Mammoth often skips Choice if it's not handled. We'll try Choice first.
        if (choice) return Array.from(choice.childNodes);

        const fallback = getDirectChildren(element, "mc:Fallback")[0];
        if (fallback) return Array.from(fallback.childNodes);

        return Array.from(element.childNodes);
    };

    const files = await extractFiles(
        buffer,
        x =>
            !!x.match(documentFileRegex) ||
            !!x.match(footnotesFileRegex) ||
            !!x.match(endnotesFileRegex) ||
            !!x.match(numberingFileRegex) ||
            !!x.match(corePropsFileRegex) ||
            !!x.match(customPropsFileRegex) ||
            !!x.match(appPropsFileRegex) ||
            !!x.match(relsFileRegex) ||
            !!x.match(stylesFileRegex) ||
            (!config.ignoreComments && !!x.match(commentsFileRegex)) ||
            (!config.ignoreHeadersAndFooters && (!!x.match(headerFileRegex) || !!x.match(footerFileRegex))) ||
            (!!config.extractAttachments && !!x.match(mediaFileRegex)),
        config.decompressionLimits,
        config
    );

    // A DOCX without its main document part is not a DOCX. Checked with the same regex the
    // parse loop below uses to recognize it, so the two cannot fall out of step.
    findRequiredPart(files, path => !!path.match(documentFileRegex), config,
        { fileType: 'docx', part: 'word/document.xml' });

    // Extract metadata
    const corePropsFile = files.find(f => f.path.match(corePropsFileRegex));
    const metadata = corePropsFile ? parseOfficeMetadata(corePropsFile.content.toString(), config) : {};
    const customPropsFile = files.find(f => f.path.match(customPropsFileRegex));
    if (customPropsFile) {
        const customProperties = parseOOXMLCustomProperties(customPropsFile.content.toString(), config);
        if (Object.keys(customProperties).length > 0) metadata.customProperties = customProperties;
    }
    const appPropsFile = files.find(f => f.path.match(appPropsFileRegex));
    if (appPropsFile) {
        const appProperties = parseOOXMLAppProperties(appPropsFile.content.toString(), config);
        if (Object.keys(appProperties).length > 0) {
            metadata.nativeProperties = appProperties;
            if (appProperties['Pages'] && typeof appProperties['Pages'] === 'number') {
                metadata.pages = appProperties['Pages'];
            }
        }
    }

    const footnoteMap = new Map<string, OfficeContentNode[]>();
    // The note or comment node of each id, built at its first reference and shared by the rest: built
    // per reference (its text joined each time), a small file referring to one large note thousands
    // of times filled the heap.
    const referencedNodes = new Map<string, OfficeContentNode>();
    const referencedNode = (kind: 'footnote' | 'endnote' | 'comment', id: string, build: () => OfficeContentNode): OfficeContentNode => {
        let node = referencedNodes.get(`${kind}:${id}`);
        if (!node) referencedNodes.set(`${kind}:${id}`, node = build());
        return node;
    };
    const endnoteMap = new Map<string, OfficeContentNode[]>();
    const commentMap = new Map<string, OfficeContentNode[]>();
    const commentMetadataMap = new Map<string, CommentMetadata>();
    const attachments: OfficeAttachment[] = [];
    const attachmentsByName = attachmentLookup(attachments);
    const mediaFiles = files.filter(f => f.path.match(mediaFileRegex));

    // Extract relationships
    const relsFile = files.find(f => f.path.match(relsFileRegex));
    // Null-prototype, as every map keyed by the document's own ids and names is here (see numberingMap).
    const relsMap: { [key: string]: string } = Object.create(null);
    if (relsFile) {
        const relsXml = parseXmlString(relsFile.content.toString(), { config });
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
    // Null-prototype: `numId` is the raw document `w:numId/@w:val`, so a plain `{}` here lets a
    // `numId="__proto__"` read `Object.prototype` (truthy) and, two levels down, write onto it -
    // global prototype pollution from one crafted .docx. With no prototype, `map["__proto__"]` is a
    // normal absent key and the numbering guard simply skips the undefined definition.
    const numberingMap: { [key: string]: { [key: string]: { numFmt: string, lvlText: string, start: number } } } = Object.create(null);

    if (numberingFile) {
        const numberingXml = parseXmlString(numberingFile.content.toString(), { config });
        const nums = getElementsByTagName(numberingXml, "w:num");
        const abstractNums = getElementsByTagName(numberingXml, "w:abstractNum");

        // Null-prototype: a `w:abstractNumId` of `toString` found the prototype's function, and the
        // parse failed on it.
        const abstractNumMap: { [key: string]: Element } = Object.create(null);
        for (const abstractNum of abstractNums) {
            const abstractNumId = abstractNum.getAttribute("w:abstractNumId");
            if (abstractNumId) {
                abstractNumMap[abstractNumId] = abstractNum;
            }
        }

        // Each abstract numbering's levels, read once.
        const abstractLevelTables = new Map<string, { [key: string]: { numFmt: string, lvlText: string, start: number } }>();
        const abstractLevels = (abstractNumId: string) => {
            let table = abstractLevelTables.get(abstractNumId);
            if (!table) {
                table = Object.create(null) as { [key: string]: { numFmt: string, lvlText: string, start: number } };
                for (const lvl of getChildElements(abstractNumMap[abstractNumId], "w:lvl")) {
                    const ilvl = lvl.getAttribute("w:ilvl");
                    if (!ilvl) continue;
                    setOwn(table, ilvl, {
                        numFmt: getChildElements(lvl, "w:numFmt")[0]?.getAttribute("w:val") || 'decimal',
                        lvlText: getChildElements(lvl, "w:lvlText")[0]?.getAttribute("w:val") || '',
                        start: parseInt(getChildElements(lvl, "w:start")[0]?.getAttribute("w:val") || '1', 10)
                    });
                }
                abstractLevelTables.set(abstractNumId, table);
            }
            return table;
        };

        for (const num of nums) {
            const numId = num.getAttribute("w:numId");
            const abstractNumIdNode = getFirstElementByTagName(num, "w:abstractNumId");
            const abstractNumId = abstractNumIdNode?.getAttribute("w:val");

            if (numId && abstractNumId && abstractNumMap[abstractNumId]) {
                // Its levels are the abstract numbering's, read once and shared through the prototype
                // (the root of which has none, so `w:ilvl="__proto__"` finds nothing inherited): read for
                // each num naming it, 20,000 nums over 20,000 levels took minutes. An override is a copy
                // of its level, own to this num.
                numberingMap[numId] = Object.create(abstractLevels(abstractNumId));

                // Apply instance overrides (w:lvlOverride)
                const overrides = getChildElements(num, "w:lvlOverride");
                for (const override of overrides) {
                    const ilvl = override.getAttribute("w:ilvl");
                    if (ilvl && numberingMap[numId][ilvl]) {
                        const startOverride = getChildElements(override, "w:startOverride")[0];
                        if (startOverride) {
                            setOwn(numberingMap[numId], ilvl, { ...numberingMap[numId][ilvl], start: parseInt(startOverride.getAttribute("w:val") || '1', 10) });
                        }
                    }
                }
            }
        }
    }

    // Parse Styles
    const stylesFile = files.find(f => f.path.match(stylesFileRegex));
    const styleMap: { [key: string]: { formatting: TextFormatting, alignment?: 'left' | 'center' | 'right' | 'justify', backgroundColor?: string, paragraphIndentation?: IndentationMetadata } } = Object.create(null);

    if (stylesFile) {
        const stylesXml = parseXmlString(stylesFile.content.toString(), { config });
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
        const stylesXml = parseXmlString(stylesFile.content.toString(), { config });
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
        const stylesXml = parseXmlString(stylesFile.content.toString(), { config });
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
    // Null-prototype for the same reason as numberingMap: both are keyed by the raw document `numId`
    // and written two levels deep, which is the prototype-pollution vector.
    const numberingState: { [key: string]: { [key: string]: number } } = Object.create(null);
    const listCounters: { [key: string]: { [key: string]: number } } = Object.create(null); // Track item index per listId/level

    // Helper to parse a paragraph node
    /**
     * @param into - The paragraph a text box's paragraph is read into (its children, notes, comments,
     * bookmarks, and whether a link encloses the text box): pushed there directly, where copying each
     * text box paragraph's content up into the one around it copied it again at every level of nesting.
     */
    const parseParagraph = (pNode: Element, documentContent: string, pendingAnchorIds: string[] = [], into?: ParagraphSink): OfficeContentNode => {
        // Check if it's a list item
        // The paragraph's own properties (children of it, as the schema has them): looked up through
        // its whole subtree, a text box's numbering made the paragraph around it a list item, and a
        // lookup at each level of nested text boxes took time in the square of their content.
        const pPr = getDirectChildren(pNode, "w:pPr")[0];
        const numPr = pPr ? getDirectChildren(pPr, "w:numPr")[0] : undefined;
        const isList = !!numPr;

        // Check if it's a heading
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

        // Runs inherit their base formatting from the style chain: the paragraph style (seeded
        // here, and re-applied via the run-style path below), then any character style, then the
        // run's own properties. The paragraph-mark run properties (`<w:pPr><w:rPr>`) format only
        // the paragraph mark glyph itself per OOXML ISO 29500 §17.3.1.29, so they are deliberately
        // NOT folded into the run base - doing so bled the paragraph mark's bold/italic/color/etc.
        // onto every run in the paragraph (issue #109).
        const paragraphRunFormatting: TextFormatting = { ...styleProps.formatting };

        // Extract text and children
        let text = '';
        const children: OfficeContentNode[] = into?.children ?? [];
        const notes: OfficeContentNode[] = into?.notes ?? [];
        const comments: OfficeContentNode[] = into?.comments ?? [];

        // Traverse children of paragraph (runs, hyperlinks, etc.)
        // Whether a hyperlink enclosing the node being read gives its runs their link (see w:hyperlink).
        let insideLink = into?.insideLink ?? false;
        const processChildNode = (node: Node) => {
            if (isElement(node) && (node.nodeName === 'w:r' || node.nodeName === 'm:r')) {
                const runNode = node;
                const rPr = getDirectChildren(runNode, "w:rPr")[0];

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
                    if (child.tagName === "w:t" || child.tagName === "t" || child.tagName === "m:t") {
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

                // The run's drawings: for mc:AlternateContent, the branch resolveAlternateContent picks
                // (reading both made one picture two), and none inside a text box of theirs.
                const runDrawings: Element[] = [];
                const pendingDrawing: Node[] = [];
                const pushNodes = (nodes: ArrayLike<Node>) => { for (let i = nodes.length - 1; i >= 0; i--) pendingDrawing.push(nodes[i]); };
                pushNodes(runNode.childNodes);
                while (pendingDrawing.length) {
                    const item = pendingDrawing.pop()!;
                    if (!isElement(item) || TEXT_BOX.has(item.nodeName)) continue;
                    if (item.nodeName === 'w:drawing' || item.nodeName === 'drawing' || item.nodeName === 'w:pict' || item.nodeName === 'pict') runDrawings.push(item);
                    else if (item.nodeName === 'mc:AlternateContent' || item.nodeName === 'AlternateContent') pushNodes(resolveAlternateContent(item));
                    else pushNodes(item.childNodes);
                }

                // Images/Drawings
                if (config.extractAttachments) {
                    for (const imgNode of runDrawings) {
                        // Extract Alt Text
                        let altText = '';
                        const docPr = ownFirst(imgNode, "wp:docPr");
                        if (docPr) {
                            altText = docPr.getAttribute("descr") || docPr.getAttribute("title") || '';
                        }
                        // A picture that is a link itself (Insert > Link on a picture writes it on docPr).
                        const pictureLinkRid = docPr ? getFirstElementByTagName(docPr, "a:hlinkClick")?.getAttribute("r:id") : null;
                        const pictureLink = pictureLinkRid && relsMap[pictureLinkRid] ? { link: relsMap[pictureLinkRid], linkType: 'external' as const } : {};

                        // Extract Relationship ID
                        let rId = '';
                        const blip = ownFirst(imgNode, "a:blip");
                        if (blip) {
                            rId = blip.getAttribute("r:embed") || '';
                        } else {
                            const imagedata = ownFirst(imgNode, "v:imagedata");
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
                                    metadata: { attachmentName: filename, altText: altText, ...pictureLink } as ImageMetadata
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

                // A text box the run draws: its paragraphs' content joins this one (as a text box drawn
                // straight in the paragraph does), where it was dropped.
                for (const drawing of runDrawings) {
                    for (const txbx of getOutermostElements(drawing, "w:txbxContent")) {
                        for (const txbxParagraph of getOutermostElements(txbx, "w:p")) {
                            text += parseParagraph(txbxParagraph, documentContent, [], { children, notes, comments, anchorIds, insideLink }).text;
                        }
                    }
                }

                // Footnotes/Endnotes inside runs
                if (!config.ignoreNotes) {
                    const footnoteRef = ownFirst(runNode, "w:footnoteReference");
                    if (footnoteRef) {
                        const id = footnoteRef.getAttribute("w:id");
                        if (id && footnoteMap.has(id)) {
                            const noteNodes = footnoteMap.get(id)!;
                            const noteNode = referencedNode('footnote', id, () => ({
                                type: 'note',
                                text: noteNodes.map((n: OfficeContentNode) => n.text).join(' '),
                                children: noteNodes,
                                metadata: { noteType: 'footnote', noteId: id }
                            } as OfficeContentNode));
                            if (children.length > 0) {
                                const target = children[children.length - 1];
                                if (!target.notes) target.notes = [];
                                target.notes.push(noteNode);
                            } else {
                                notes.push(noteNode);
                            }
                        }
                    }

                    const endnoteRef = ownFirst(runNode, "w:endnoteReference");
                    if (endnoteRef) {
                        const id = endnoteRef.getAttribute("w:id");
                        if (id && endnoteMap.has(id)) {
                            const noteNodes = endnoteMap.get(id)!;
                            const noteNode = referencedNode('endnote', id, () => ({
                                type: 'note',
                                text: noteNodes.map((n: OfficeContentNode) => n.text).join(' '),
                                children: noteNodes,
                                metadata: { noteType: 'endnote', noteId: id }
                            } as OfficeContentNode));
                            if (children.length > 0) {
                                const target = children[children.length - 1];
                                if (!target.notes) target.notes = [];
                                target.notes.push(noteNode);
                            } else {
                                notes.push(noteNode);
                            }
                        }
                    }
                }

                // Comments inside runs
                if (!config.ignoreComments) {
                    const commentRef = ownFirst(runNode, "w:commentReference");
                    if (commentRef) {
                        const id = commentRef.getAttribute("w:id");
                        if (id && commentMap.has(id)) {
                            const commentNodes = commentMap.get(id)!;
                            const commentInfo = commentMetadataMap.get(id);
                            const commentNode = referencedNode('comment', id, () => ({
                                type: 'comment',
                                text: commentNodes.map((n: OfficeContentNode) => n.text).join(' '),
                                children: commentNodes,
                                metadata: commentInfo || { commentId: id }
                            } as OfficeContentNode));
                            if (children.length > 0) {
                                const target = children[children.length - 1];
                                if (!target.comments) target.comments = [];
                                target.comments.push(commentNode);
                            } else {
                                comments.push(commentNode);
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

                // Process children of hyperlink (usually runs). The outermost hyperlink with a link gives
                // its runs (and a picture in it) that link, once, after they are read: every level of
                // nested hyperlinks gave it again to all the runs inside it, so 1,000 levels around
                // 400,000 runs (23 KB) took 17 seconds.
                const applies = !!linkMetadata && !insideLink;
                const wasInside = insideLink;
                if (applies) insideLink = true;
                const startIndex = children.length;
                for (const child of Array.from(hlNode.childNodes)) processChildNode(child);
                insideLink = wasInside;
                if (applies) {
                    for (let i = startIndex; i < children.length; i++) {
                        if (children[i].type === 'text' || children[i].type === 'image') {
                            children[i].metadata = { ...(children[i].metadata ?? {}), ...linkMetadata } as any;
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
                // The picture's own text boxes, not those in their paragraphs: parsing a text box's
                // paragraphs reaches those, and taking them here too read nested text boxes again at
                // each level (1.3 KB took half a minute).
                const textBoxes = getOutermostElements(node, "w:txbxContent");
                for (const txbx of textBoxes) {
                    // Its paragraphs wherever they sit: in a table or a content control of the text box too.
                    for (const txbxParagraph of getOutermostElements(txbx, "w:p")) {
                        text += parseParagraph(txbxParagraph, documentContent, [], { children, notes, comments, anchorIds, insideLink }).text;
                    }
                }
            } else if (isElement(node) && (node.nodeName === 'm:oMath' || node.nodeName === 'oMath'
                || node.nodeName === 'm:oMathPara' || node.nodeName === 'oMathPara')) {
                // Equations. Without this branch they reach the generic fallback below, which
                // recurses into every child and concatenates the `m:t` runs with no separators -
                // so `<m:num>1</m:num><m:den>2</m:den>` came out as "12". That is worse than
                // dropping the formula: the result still reads as a number, so nothing downstream
                // can tell it is wrong.
                //
                // `m:oMathPara` is a display equation on its own line; a bare `m:oMath` is inline.
                const isBlock = node.nodeName === 'm:oMathPara' || node.nodeName === 'oMathPara';
                const latex = ommlToLatex(node);
                if (!isEmptyMath(latex)) {
                    text += latex;
                    children.push({
                        type: 'code',
                        text: latex,
                        metadata: { math: isBlock ? 'block' : 'inline' } as CodeMetadata
                    });
                }
            } else if (node.childNodes.length > 0) {
                // Generic fallback for unknown elements that might contain content
                for (const child of Array.from(node.childNodes)) processChildNode(child);
            }
        };

        const anchorIds: string[] = into?.anchorIds ?? [...pendingAnchorIds];
        const childNodes = Array.from(pNode.childNodes);
        for (const child of childNodes) {
            processChildNode(child);
        }

        const commonMetadata = anchorIds.length > 0 ? { anchorIds } : {};

        if (isList) {
            const numIdNode = getFirstElementByTagName(numPr, "w:numId");
            const ilvlNode = getFirstElementByTagName(numPr, "w:ilvl");
            const numId = numIdNode ? numIdNode.getAttribute("w:val") || '0' : '0';
            // Word's levels are 0 to 8: a level of -20,000,000 ran the loop clearing deeper levels below
            // for 20 million steps per paragraph, and one past 2^53 never ended.
            const levelValue = ilvlNode ? parseInt(ilvlNode.getAttribute("w:val") || '0', 10) : 0;
            const ilvl = Number.isFinite(levelValue) ? Math.max(0, Math.min(8, levelValue)) : 0;

            let listType: 'ordered' | 'unordered' = 'ordered';
            let itemIndex = 0;
            if (numId && numberingMap[numId]) {
                const ilvlStr = ilvl.toString();
                if (!numberingState[numId]) numberingState[numId] = Object.create(null);
                if (!numberingState[numId][ilvlStr]) numberingState[numId][ilvlStr] = 0;
                numberingState[numId][ilvlStr]++;
                for (let k = ilvl + 1; k < 10; k++) {
                    if (numberingState[numId][k.toString()]) numberingState[numId][k.toString()] = 0;
                }
                const numFmt = numberingMap[numId][ilvlStr]?.numFmt || 'decimal';
                listType = numFmt === 'bullet' ? 'unordered' : 'ordered';

                // Track itemIndex (starts at override or default, continues across interruptions for same listId)
                if (!listCounters[numId]) listCounters[numId] = Object.create(null);
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
                ...(notes.length > 0 ? { notes } : {}),
                ...(comments.length > 0 ? { comments } : {}),
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

            if (config.includeRawContent && !into) listNode.rawContent = getRawContent(pNode, documentContent, config);
            return listNode;

        } else if (isHeading) {
            const level = pStyleVal ? parseInt(pStyleVal.replace("Heading", ""), 10) || 1 : 1;
            const headingNode: OfficeContentNode = {
                type: 'heading',
                text: text,
                children: children,
                ...(notes.length > 0 ? { notes } : {}),
                ...(comments.length > 0 ? { comments } : {}),
                metadata: { level, alignment, paragraphIndentation: paraIndentation, style: pStyleVal ?? undefined, ...commonMetadata }
            };
            if (config.includeRawContent && !into) headingNode.rawContent = getRawContent(pNode, documentContent, config);
            return headingNode;
        } else {
            const paraNode: OfficeContentNode = {
                type: 'paragraph',
                text: text,
                children: children,
                ...(notes.length > 0 ? { notes } : {}),
                ...(comments.length > 0 ? { comments } : {}),
                metadata: { alignment, paragraphIndentation: paraIndentation, style: pStyleVal ?? undefined, ...commonMetadata }
            };
            if (config.includeRawContent && !into) paraNode.rawContent = getRawContent(pNode, documentContent, config);
            return paraNode;
        }
    };

    // Helper to parse a table node
    const parseTable = (tblNode: Element, documentContent: string, pendingAnchorIds: string[] = []): OfficeContentNode => {
        const rows: OfficeContentNode[] = [];
        const trNodes = layoutChildren(tblNode).filter(n => n.nodeName === 'w:tr');
        // Track vertical merges: colIndex -> { startCellNode, rowSpan }
        const vMergeMap = new Map<number, { node: OfficeContentNode, span: number }>();

        for (let rIndex = 0; rIndex < trNodes.length; rIndex++) {
            const trNode = trNodes[rIndex];
            const cells: OfficeContentNode[] = [];
            // The row's own cells, not nested table cells (a content control around a cell stands for it)
            const tcNodes = layoutChildren(trNode).filter(n => n.nodeName === 'w:tc');
            // <w:trPr><w:tblHeader/> marks a row that repeats as the table's header on every page -
            // Word's own header-row flag, and what this library's DOCX generator writes. Without
            // reading it, a header row survived only when it happened to be all-bold.
            const trPr = getDirectChildren(trNode, "w:trPr")[0];
            const tblHeader = trPr ? getFirstElementByTagName(trPr, "w:tblHeader") : null;
            // ST_OnOff turns the toggle off with "false"/"0"/"off"; a bare <w:tblHeader/> or any other
            // value (true/1/on) is on.
            const tblHeaderVal = tblHeader?.getAttribute("w:val");
            const isHeaderRowNode = !!tblHeader && tblHeaderVal !== "false" && tblHeaderVal !== "0" && tblHeaderVal !== "off";

            let visualCol = 0;
            for (let tcIndex = 0; tcIndex < tcNodes.length; tcIndex++) {
                const tcNode = tcNodes[tcIndex];
                const tcPr = getDirectChildren(tcNode, "w:tcPr")[0];

                // Horizontal merge (colspan)
                let colSpan = 1;
                if (tcPr) {
                    const gridSpan = getFirstElementByTagName(tcPr, "w:gridSpan");
                    if (gridSpan) {
                        // Held to 1 through MAX_COL_SPAN: a span of billions set the next cell's column there.
                        colSpan = cellSpan(gridSpan.getAttribute("w:val"), MAX_COL_SPAN);
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
                const cellContentNodes = layoutChildren(tcNode);
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
                // A row node cannot carry metadata of its own, so the header flag rides on its cells -
                // the same `style: 'header'` convention the PDF parser uses for TH cells, which is what
                // `isHeaderRow` (and therefore every generator) reads.
                if (isHeaderRowNode) (cellNode.metadata as CellMetadata).style = 'header';

                if (tcPr) {
                    const shd = getFirstElementByTagName(tcPr, "w:shd");
                    if (shd) {
                        const fill = shd.getAttribute("w:fill");
                        if (fill && fill !== "auto") {
                            (cellNode.metadata as CellMetadata).backgroundColor = "#" + fill;
                        }
                    }
                }

                if (isVMerge) {
                    if (vMergeRestart) {
                        vMergeMap.set(visualCol, { node: cellNode, span: 1 });
                        cells.push(cellNode);
                    } else {
                        const mergeInfo = vMergeMap.get(visualCol);
                        if (mergeInfo) {
                            mergeInfo.span++;
                            (mergeInfo.node.metadata as CellMetadata).rowSpan = mergeInfo.span;

                            // A continuation cell must still contain a block child, so a generated
                            // document carries an empty <w:p/> here; only real content is worth
                            // folding into the merged cell, otherwise the round-trip gains an empty
                            // paragraph and a stray space. `cellText` is built from `w:p` runs only, so
                            // guarding on it alone drops a continuation whose sole content is an
                            // image-only paragraph or a nested table (both have empty text) - guard on
                            // "has a renderable child" instead, and only extend the text when there is
                            // some.
                            const hasFoldableContent = cellChildren.some(c =>
                                c.type !== 'paragraph'
                                || (c.text?.trim().length ?? 0) > 0
                                || (c.children?.length ?? 0) > 0
                            );
                            if (hasFoldableContent) {
                                if (!mergeInfo.node.children) mergeInfo.node.children = [];
                                appendAll(mergeInfo.node.children, cellChildren);
                                if (cellText.trim()) mergeInfo.node.text += " " + cellText;
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
            const footnotesDoc = parseXmlString(footnotesFile.content.toString(), { config });
            const footnoteXml = footnotesFile.content.toString();
            const footnoteNodes = getElementsByTagName(footnotesDoc, "w:footnote");
            for (const node of footnoteNodes) {
                const id = node.getAttribute("w:id");
                if (!id || id === "-1" || id === "0") continue;
                const pNodes = getOutermostElements(node, "w:p");
                footnoteMap.set(id, pNodes.map(p => parseParagraph(p, footnoteXml)));
            }
        }

        const endnotesFile = files.find(f => f.path.match(endnotesFileRegex));
        if (endnotesFile) {
            const endnotesDoc = parseXmlString(endnotesFile.content.toString(), { config });
            const endnoteXml = endnotesFile.content.toString();
            const endnoteNodes = getElementsByTagName(endnotesDoc, "w:endnote");
            for (const node of endnoteNodes) {
                const id = node.getAttribute("w:id");
                if (!id || id === "-1" || id === "0") continue;
                const pNodes = getOutermostElements(node, "w:p");
                endnoteMap.set(id, pNodes.map(p => parseParagraph(p, endnoteXml)));
            }
        }
    }

    // Pre-process comments
    if (!config.ignoreComments) {
        const commentsFile = files.find(f => f.path.match(commentsFileRegex));
        if (commentsFile) {
            const commentsDoc = parseXmlString(commentsFile.content.toString(), { config });
            const commentsXml = commentsFile.content.toString();
            const commentNodes = getElementsByTagName(commentsDoc, "w:comment");
            for (const node of commentNodes) {
                const id = node.getAttribute("w:id");
                if (!id) continue;
                const author = node.getAttribute("w:author") || undefined;
                const date = node.getAttribute("w:date") || undefined;
                const initials = node.getAttribute("w:initials") || undefined;

                commentMetadataMap.set(id, { commentId: id, author, date, initials });
                const pNodes = getOutermostElements(node, "w:p");
                commentMap.set(id, pNodes.map(p => parseParagraph(p, commentsXml)));
            }
        }
    }

    // Pre-process headers and footers
    const headers: OfficeContentNode[] = [];
    const footers: OfficeContentNode[] = [];
    if (!config.ignoreHeadersAndFooters) {
        const headerFiles = files.filter(f => f.path.match(headerFileRegex));
        for (const hFile of headerFiles) {
            const hDoc = parseXmlString(hFile.content.toString(), { config });
            const hXml = hFile.content.toString();
            for (const child of layoutChildren(hDoc.documentElement)) {
                if (child.nodeName === 'w:p') headers.push(parseParagraph(child, hXml));
                else if (child.nodeName === 'w:tbl') headers.push(parseTable(child, hXml));
            }
        }

        const footerFiles = files.filter(f => f.path.match(footerFileRegex));
        for (const fFile of footerFiles) {
            const fDoc = parseXmlString(fFile.content.toString(), { config });
            const fXml = fFile.content.toString();
            for (const child of layoutChildren(fDoc.documentElement)) {
                if (child.nodeName === 'w:p') footers.push(parseParagraph(child, fXml));
                else if (child.nodeName === 'w:tbl') footers.push(parseTable(child, fXml));
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
        if (file.path.match(commentsFileRegex)) continue;
        if (file.path.match(headerFileRegex)) continue;
        if (file.path.match(footerFileRegex)) continue;

        const documentContent = file.content.toString();
        const doc = parseXmlString(documentContent, { config, locator: config.includeRawContent });
        const body = getFirstElementByTagName(doc, "w:body");
        if (body) {
            const bodyChildren = layoutChildren(body);
            let pendingAnchorIds: string[] = [];

            for (const child of bodyChildren) {
                checkAbortSignal(config.abortSignal);
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

            if (config.ocr && attachment.mimeType.startsWith('image/')) {
                const ocrText = await ocrDuringParse(media.content, config, attachment.name);
                if (ocrText !== undefined) attachment.ocrText = ocrText;
            }
        }

        // Assign OCR text to image nodes
        if (config.ocr) {
            const assignOcr = (nodes: OfficeContentNode[]) => {
                for (const node of nodes) {
                    if (node.type === 'image' && 'attachmentName' in (node.metadata || {})) {
                        const meta = node.metadata as ImageMetadata;
                        const attachment = attachmentsByName.get(meta.attachmentName);
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

    const auxiliaryContent = (headers.length > 0 || footers.length > 0) ? {
        ...(headers.length > 0 ? { headers } : {}),
        ...(footers.length > 0 ? { footers } : {})
    } : undefined;

    return createAST(
        'docx',
        { ...metadata, formatting: docDefaults, styleMap: plainRecord(styleMap) },
        content,
        attachments,
        config,
        auxiliaryContent,
    );
};

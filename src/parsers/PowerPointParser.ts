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

import { XMLSerializer } from '@xmldom/xmldom';
import { ChartMetadata, ImageMetadata, ListMetadata, OfficeAttachment, OfficeContentNode, OfficeParserAST, OfficeParserConfig, SlideMetadata, TextFormatting } from '../types';
import { extractChartData } from '../utils/chartUtils';
import { logWarning } from '../utils/errorUtils';
import { createAttachment } from '../utils/imageUtils';
import { performOcr } from '../utils/ocrUtils';
import { getElementsByTagName, parseOfficeMetadata, parseOOXMLCustomProperties, parseXmlString } from '../utils/xmlUtils';
import { extractFiles } from '../utils/zipUtils';

/**
 * Parses a PowerPoint presentation (.pptx) and extracts slides and notes.
 * 
 * @param buffer - The PPTX file as a Buffer
 * @param config - Parser configuration
 * @returns A promise resolving to the parsed AST
 */
export const parsePowerPoint = async (buffer: Buffer, config: OfficeParserConfig): Promise<OfficeParserAST> => {
    const allFilesRegex = /ppt\/(notesSlides|slides)\/(notesSlide|slide)\d+.xml/g;
    const slidesRegex = /ppt\/slides\/slide\d+.xml/g;
    const slideRelsRegex = /ppt\/slides\/_rels\/slide\d+\.xml\.rels/;
    const slideNumberRegex = /lide(\d+)\.xml/;
    const mediaFileRegex = /ppt\/media\/.*/;
    const chartFileRegex = /ppt\/charts\/chart\d+\.xml/;
    const corePropsFileRegex = /docProps\/core\.xml/;
    const customPropsFileRegex = /docProps\/custom\.xml/;
    const xmlSerializer = new XMLSerializer();

    const files = await extractFiles(buffer, x =>
        !!x.match(config.ignoreNotes ? slidesRegex : allFilesRegex) ||
        !!x.match(corePropsFileRegex) ||
        !!x.match(customPropsFileRegex) ||
        !!x.match(slideRelsRegex) ||
        (!!config.extractAttachments && (!!x.match(mediaFileRegex) || !!x.match(chartFileRegex)))
    );

    // Extract metadata
    const corePropsFile = files.find(f => f.path.match(corePropsFileRegex));
    const metadata = corePropsFile ? parseOfficeMetadata(corePropsFile.content.toString()) : {};
    const customPropsFile = files.find(f => f.path.match(customPropsFileRegex));
    if (customPropsFile) {
        const customProperties = parseOOXMLCustomProperties(customPropsFile.content.toString());
        if (Object.keys(customProperties).length > 0) metadata.customProperties = customProperties;
    }

    // Sort files
    files.sort((a, b) => {
        const aMatch = a.path.match(slideNumberRegex);
        const bMatch = b.path.match(slideNumberRegex);
        const aNum = aMatch ? parseInt(aMatch[1]) : 0;
        const bNum = bMatch ? parseInt(bMatch[1]) : 0;
        return aNum - bNum;
    });

    const content: OfficeContentNode[] = [];
    const rawContents: string[] = [];
    const slideRelsMap: Record<number, Record<string, { type: string, target: string }>> = {};

    let currentListId = 0;
    let runningListIndex = 0;

    let lastWasList = false;
    let lastListType: 'ordered' | 'unordered' | null = null;
    let lastListIndent = 0;

    // per indent counters (for nested lists)
    const levelCounters: { [level: number]: number } = {};

    // Helper to parse a table node
    const parseTable = (tblNode: Element): OfficeContentNode => {
        const rows: OfficeContentNode[] = [];
        const trNodes = getElementsByTagName(tblNode, "a:tr");

        for (let rIndex = 0; rIndex < trNodes.length; rIndex++) {
            const trNode = trNodes[rIndex];
            const cells: OfficeContentNode[] = [];
            const tcNodes = getElementsByTagName(trNode, "a:tc");

            for (let cIndex = 0; cIndex < tcNodes.length; cIndex++) {
                const tcNode = tcNodes[cIndex];
                const cellChildren: OfficeContentNode[] = [];
                let cellText = '';

                // Cells contain text bodies (txBody) which contain paragraphs
                const txBody = getElementsByTagName(tcNode, "a:txBody")[0];
                if (txBody) {
                    const paragraphs = getElementsByTagName(txBody, "a:p");
                    for (const p of paragraphs) {
                        // Reuse paragraph parsing logic if possible, or duplicate for now
                        // For simplicity, duplicating basic logic here as the main loop one is tied to shapes
                        const pNode: OfficeContentNode = {
                            type: 'paragraph',
                            text: '',
                            children: [],
                            metadata: {}
                        };

                        if (config.includeRawContent) {
                            pNode.rawContent = p.toString();
                        }

                        const runs = getElementsByTagName(p, "a:r");
                        for (const r of runs) {
                            const t = getElementsByTagName(r, "a:t")[0];
                            if (t && t.childNodes[0]) {
                                const textContent = t.childNodes[0].nodeValue || '';
                                pNode.text += textContent;

                                const rPr = getElementsByTagName(r, "a:rPr")[0];
                                const formatting: TextFormatting = {};
                                if (rPr) {
                                    if (rPr.getAttribute("b") === "1") formatting.bold = true;
                                    if (rPr.getAttribute("i") === "1") formatting.italic = true;
                                    if (rPr.getAttribute("u") === "sng") formatting.underline = true;
                                    if (rPr.getAttribute("strike") === "sngStrike") formatting.strikethrough = true;
                                    const sz = rPr.getAttribute("sz");
                                    if (sz) formatting.size = (parseInt(sz) / 100).toString() + 'pt';

                                    const solidFill = getElementsByTagName(rPr, "a:solidFill")[0];
                                    if (solidFill) {
                                        const srgbClr = getElementsByTagName(solidFill, "a:srgbClr")[0];
                                        if (srgbClr) {
                                            const val = srgbClr.getAttribute("val");
                                            if (val) formatting.color = '#' + val;
                                        }
                                    }

                                    const latin = getElementsByTagName(rPr, "a:latin")[0];
                                    if (latin) {
                                        const typeface = latin.getAttribute("typeface");
                                        if (typeface) formatting.font = typeface;
                                    }
                                }

                                pNode.children?.push({
                                    type: 'text',
                                    text: textContent,
                                    formatting: formatting
                                });
                            }
                        }

                        cellChildren.push(pNode);
                        cellText += pNode.text;
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

    /** Extract an AST node for p:pic */
    const extractImageNode = (imageNode: Element, slideNumber: number): OfficeContentNode | null => {
        const blip = getElementsByTagName(imageNode, "a:blip")[0];
        if (!blip) return null;

        const rId = blip.getAttribute("r:embed");
        if (!rId) return null;

        const rel = slideRelsMap[slideNumber]?.[rId];
        if (!rel || rel.type !== "image") return null;

        const attachmentName = rel.target;

        const nvPicPr = getElementsByTagName(imageNode, "p:nvPicPr")[0];
        const cNvPr = nvPicPr ? getElementsByTagName(nvPicPr, "p:cNvPr")[0] : null;

        const altText = cNvPr?.getAttribute("descr") || undefined;

        return {
            type: "image",
            text: '',
            metadata:
            {
                attachmentName,
                altText,
            }
        };
    }

    /**
     * Extract an AST node for p:graphicFrame that contains a chart.
     * This mirrors extractImageNode, but for charts instead of images.
     *
     * Steps:
     * 1. Find <a:graphicData> inside the graphicFrame.
     * 2. Check if uri is the chart namespace.
     * 3. Extract <c:chart> and read r:id.
     * 4. Resolve the relationship using slideRelsMap.
     * 5. Produce a chart node with attachmentName (like images).
     * 6. Chart text/data will be injected later in the pipeline.
     *
     * @param frameNode p:graphicFrame element
     * @param slideNumber Current slide number for relationship resolution
     * @returns A chart OfficeContentNode or null if not a chart.
     */
    const extractChartNode = (frameNode: Element, slideNumber: number): OfficeContentNode | null => {
        // Step 1: Find <a:graphicData>
        const graphicData = getElementsByTagName(frameNode, "a:graphicData")[0];
        if (!graphicData) {
            return null;
        }

        // Step 2: Verify chart namespace
        // Must be: http://schemas.openxmlformats.org/drawingml/2006/chart
        const uri = graphicData.getAttribute("uri");
        const isChartGraphic = uri === "http://schemas.openxmlformats.org/drawingml/2006/chart";
        if (!isChartGraphic) {
            return null;
        }

        // Step 3: Find <c:chart>
        const cChart = getElementsByTagName(graphicData, "c:chart")[0];
        if (!cChart) {
            return null;
        }

        // Step 4: Extract r:id (relationship id)
        const rId = cChart.getAttribute("r:id");
        if (!rId) {
            return null;
        }

        // Step 5: Resolve relationship target from slideRelsMap
        const rel = slideRelsMap[slideNumber]?.[rId];
        if (!rel || rel.type !== "chart") {
            return null;
        }

        // rel.target will be something like "chart1.xml"
        const attachmentName = rel.target;

        // Step 6: Build AST node
        const chartNode: OfficeContentNode =
        {
            type: "chart",
            text: "",              // chart text gets filled later
            metadata:
            {
                attachmentName     // name used to link to attachments & chartData
            }
        };

        // Optional: include raw XML of the whole frame
        if (config.includeRawContent) {
            chartNode.rawContent = xmlSerializer.serializeToString(frameNode);
        }

        return chartNode;
    };

    /** Extract an AST node for p:graphicFrame */
    const extractGraphicFrameNode = (frameNode: Element, slideNumber: number): OfficeContentNode | null => {
        const tbl = getElementsByTagName(frameNode, "a:tbl")[0];
        if (tbl) {
            const tableNode = parseTable(tbl);
            if (config.includeRawContent) {
                tableNode.rawContent = xmlSerializer.serializeToString(frameNode);
            }
            if (tableNode.children && tableNode.children.length > 0) {
                return tableNode;
            }
        }
        if (frameNode.getElementsByTagName("c:chart").length > 0) {
            const chartNode = extractChartNode(frameNode, slideNumber);
            if (chartNode) {
                return chartNode;
            }
        }
        return null;
    }

    /** Extract text and hyperlinks from a p:sp shape. */
    const extractShapeNodes = (spNode: Element, slideNumber: number): OfficeContentNode[] => {
        const nodes: OfficeContentNode[] = [];
        // Check for placeholder type (title, body, etc.)
        const nvSpPr = getElementsByTagName(spNode, "p:nvSpPr")[0];
        const nvPr = nvSpPr ? getElementsByTagName(nvSpPr, "p:nvPr")[0] : null;
        const ph = nvPr ? getElementsByTagName(nvPr, "p:ph")[0] : null;
        const type = ph ? ph.getAttribute("type") : "body";

        const isTitle = type === "title" || type === "ctrTitle";

        const txBody = getElementsByTagName(spNode, "p:txBody")[0];
        if (txBody) {
            const paragraphs = getElementsByTagName(txBody, "a:p");

            for (let i = 0; i < paragraphs.length; i++) {
                const p = paragraphs[i];
                const pNode: OfficeContentNode = {
                    type: isTitle ? 'heading' : 'paragraph',
                    text: '',
                    children: [],
                    metadata: isTitle ? { level: 1 } : {}
                };

                // Paragraph Alignment and List Detection
                const pPr = getElementsByTagName(p, "a:pPr")[0];
                let isList = false;
                let listType: 'ordered' | 'unordered' = 'unordered';
                let lvl = 0;

                if (pPr) {
                    const lvlAttr = pPr.getAttribute("lvl");
                    if (lvlAttr) lvl = parseInt(lvlAttr);

                    const buAutoNum = getElementsByTagName(pPr, "a:buAutoNum")[0];
                    const buChar = getElementsByTagName(pPr, "a:buChar")[0];
                    const buBlip = getElementsByTagName(pPr, "a:buBlip")[0];
                    const buNode = getElementsByTagName(pPr, "a:bu")[0];

                    if (buAutoNum) {
                        isList = true;
                        listType = 'ordered';
                    } else if (buChar || buBlip) {
                        isList = true;
                        listType = 'unordered';
                    } else if (buNode) {
                        // inherited bullet from a list style
                        isList = true;
                        listType = 'unordered';
                    }

                    const algn = pPr.getAttribute("algn");
                    if (algn) {
                        const alignMap: Record<string, 'left' | 'center' | 'right' | 'justify'> = {
                            'l': 'left',
                            'ctr': 'center',
                            'r': 'right',
                            'just': 'justify'
                        };
                        if (alignMap[algn]) {
                            (pNode.metadata as any).alignment = alignMap[algn];
                        }
                    }
                }

                if (isList) {
                    pNode.type = 'list';

                    const ilvl = lvl;

                    // detect a new list when bullet type changes or previous was not a list
                    const newList =
                        !lastWasList ||
                        listType !== lastListType;

                    if (newList) {
                        // new list → new ID
                        currentListId++;

                        // clear counters for nested levels
                        for (const k in levelCounters) {
                            delete levelCounters[k];
                        }

                        // start item index at 1
                        runningListIndex = 0;
                        levelCounters[ilvl] = 0;
                    }
                    else {
                        // same listId, but indentation may change

                        // if going deeper → start at 1 for that level
                        if (ilvl > lastListIndent) {
                            runningListIndex = 0;
                            levelCounters[ilvl] = 0;
                        }
                        // if going shallower → restore previous level counter + 1
                        else if (ilvl < lastListIndent) {
                            // remove deeper counters
                            for (const lvlKey in levelCounters) {
                                const lv = parseInt(lvlKey);
                                if (lv > ilvl) delete levelCounters[lv];
                            }

                            // continue counter at this level
                            const prev = levelCounters[ilvl] || 0;
                            runningListIndex = prev + 1;
                            levelCounters[ilvl] = runningListIndex;
                        }
                        // same level → increment
                        else {
                            const prev = levelCounters[ilvl] || 0;
                            runningListIndex = prev + 1;
                            levelCounters[ilvl] = runningListIndex;
                        }
                    }

                    // update tracking state
                    lastWasList = true;
                    lastListType = listType;
                    lastListIndent = ilvl;

                    // metadata output
                    pNode.metadata = {
                        ...pNode.metadata,
                        listType,
                        indentation: ilvl,
                        listId: currentListId.toString(),
                        itemIndex: runningListIndex,
                        alignment: (pNode.metadata as any)?.alignment || 'left',
                    } as ListMetadata;
                }
                else {
                    lastWasList = false;
                    lastListType = null;
                    lastListIndent = 0;
                }

                if (isTitle) {
                    pNode.metadata = { ...pNode.metadata, level: 1 };
                }

                if (config.includeRawContent) {
                    pNode.rawContent = p.toString();
                }

                const runs = getElementsByTagName(p, "a:r");

                for (let j = 0; j < runs.length; j++) {
                    const r = runs[j];
                    const t = getElementsByTagName(r, "a:t")[0];
                    if (t && t.childNodes[0]) {
                        const textContent = t.childNodes[0].nodeValue || '';
                        pNode.text += textContent;

                        const rPr = getElementsByTagName(r, "a:rPr")[0];
                        const formatting: TextFormatting = {};
                        if (rPr) {
                            if (rPr.getAttribute("b") === "1") formatting.bold = true;
                            if (rPr.getAttribute("i") === "1") formatting.italic = true;
                            if (rPr.getAttribute("u") === "sng") formatting.underline = true;
                            if (rPr.getAttribute("strike") === "sngStrike") formatting.strikethrough = true;

                            const sz = rPr.getAttribute("sz");
                            if (sz) formatting.size = (parseInt(sz) / 100).toString() + 'pt';

                            // Color extraction
                            const solidFill = getElementsByTagName(rPr, "a:solidFill")[0];
                            if (solidFill) {
                                const srgbClr = getElementsByTagName(solidFill, "a:srgbClr")[0];
                                if (srgbClr) {
                                    const val = srgbClr.getAttribute("val");
                                    if (val) formatting.color = '#' + val;
                                }
                            }

                            // Highlight extraction
                            const highlight = getElementsByTagName(rPr, "a:highlight")[0];
                            if (highlight) {
                                const srgbClr = getElementsByTagName(highlight, "a:srgbClr")[0];
                                if (srgbClr) {
                                    const val = srgbClr.getAttribute("val");
                                    if (val) formatting.backgroundColor = '#' + val;
                                }
                            }

                            // Font family
                            const latin = getElementsByTagName(rPr, "a:latin")[0];
                            if (latin) {
                                const typeface = latin.getAttribute("typeface");
                                if (typeface) formatting.font = typeface;
                            }

                            // Subscript/Superscript
                            const baseline = rPr.getAttribute("baseline");
                            if (baseline) {
                                const baselineVal = parseInt(baseline);
                                if (baselineVal < 0) formatting.subscript = true;
                                if (baselineVal > 0) formatting.superscript = true;
                            }
                        }

                        const textNode: OfficeContentNode = {
                            type: 'text',
                            text: textContent,
                            formatting: formatting
                        };

                        // Check for Hyperlinks
                        const hlinkClick = getElementsByTagName(r, "a:hlinkClick")[0];
                        // Check if this run has a hyperlink click action
                        if (hlinkClick) {
                            // Relationship ID for the link
                            const rId = hlinkClick.getAttribute("r:id");
                            // Optional action attribute, often for internal jumps
                            const action = hlinkClick.getAttribute("action");
                            // Result placeholders
                            let link: string | undefined;
                            let linkType: "internal" | "external" | undefined;
                            // Case 1: Relationship exists in slideRelsMap and is a real hyperlink (external URL)
                            if (rId
                                && slideRelsMap[slideNumber]
                                && slideRelsMap[slideNumber][rId]
                                && slideRelsMap[slideNumber][rId].type === "hyperlink") {
                                // External URL
                                link = slideRelsMap[slideNumber][rId].target;
                                linkType = "external";
                            }
                            // Case 2: Relationship exists and is an internal slide reference
                            else if (rId
                                && slideRelsMap[slideNumber]
                                && slideRelsMap[slideNumber][rId]
                                && slideRelsMap[slideNumber][rId].type === "slide") {
                                // Example target: ppt/slides/slide3.xml
                                link = slideRelsMap[slideNumber][rId].target;
                                linkType = "internal";
                            }
                            // Case 3: action attribute like ppaction://hlinksldjump
                            else if (action) {
                                link = action;
                                linkType = "internal";
                            }
                            // Assign metadata only if a link was actually discovered
                            if (link) {
                                textNode.metadata = { link, linkType };
                            }
                        }

                        pNode.children?.push(textNode);
                    }
                }

                if (pNode.text) {
                    nodes.push(pNode);
                }
            }
        }
        return nodes;
    }

    /**
     * Recursively traverses a PowerPoint shape tree (p:spTree),
     * including grouped shapes (p:grpSp), and dispatches each element
     * to the appropriate handler (shape, image, chart, etc.).
     *
     * This function preserves the visual order because it processes
     * children in the order they appear in the XML. It also accumulates
     * transforms so nested groups inherit positional transforms.
     *
     * @param treeNode The XML node representing <p:spTree>
     * @param parentTransform Transform inherited from parent groups (defaults to identity)
     */
    function traverseSpTree(treeNode: Element, slideNumber: number): OfficeContentNode[] {
        const nodes: OfficeContentNode[] = [];
        // Process children in XML order (this preserves Z-order)
        for (const child of Array.from(treeNode?.childNodes || [])) {
            if (child.nodeType !== 1) {
                continue;
            }

            const element = child as Element;
            const tag = element.tagName;

            // Case 1: Normal shape
            if (tag === "p:sp") {
                nodes.push(...extractShapeNodes(element, slideNumber));
            }
            // Case 2: Inline picture
            else if (tag === "p:pic") {
                const imageNode = extractImageNode(element, slideNumber);
                if (imageNode) {
                    nodes.push(imageNode);
                }
            }
            // Case 3: Chart or other graphic frame
            else if (tag === "p:graphicFrame") {
                const tableNode = extractGraphicFrameNode(element, slideNumber);
                if (tableNode) {
                    nodes.push(tableNode);
                }
            }
            // Case 4: Grouped shape (recursive!)
            else if (tag === "p:grpSp") {
                // Extract the nested <p:spTree> inside the group
                const nestedTree = element.getElementsByTagName("p:spTree")[0];
                // Recurse into the nested tree
                nodes.push(...traverseSpTree(nestedTree, slideNumber));
            }
        }
        return nodes;
    }

    // First pass: Process relationships
    for (const file of files) {
        // Check whether this file is a slideX.xml.rels file
        if (file.path.match(slideRelsRegex)) {
            /**
             * Builds a map of slide number to a map of relationship IDs containing:
             * - type: The relationship category (image, hyperlink, chart, etc)
             * - target: The fully normalized target path inside the PPTX zip
             *
             * Example structure:
             * {
             *     1: {
             *         "rId2": { type: "image", target: "ppt/media/image3.png" },
             *         "rId5": { type: "hyperlink", target: "https://example.com" }
             *     }
             * }
             *
             * @param files All extracted PPTX ZIP files.
             * @param slideRelsMap A map of slide number to relationship info.
             */
            // Extract slide number from path
            const match = file.path.match(/slide(\d+)\.xml\.rels/);
            if (match) {
                // Convert matched number to integer
                const slideNum = parseInt(match[1]);
                // Prepare map for this slide
                slideRelsMap[slideNum] = {};
                // Parse the rels XML
                const relsXml = parseXmlString(file.content.toString());
                // Get all Relationship nodes
                const relationships = getElementsByTagName(relsXml, "Relationship");
                // Loop through each relationship node
                for (let i = 0; i < relationships.length; i++) {
                    // Relationship ID, Example: "rId2"
                    const id = relationships[i].getAttribute("Id");
                    // Relationship Type, Example: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                    const typeAttr = relationships[i].getAttribute("Type");
                    // Raw Target, may be relative or absolute
                    const targetRaw = relationships[i].getAttribute("Target");
                    // Only proceed if ID and Type exist
                    if (id && typeAttr && targetRaw) {
                        // Simplify Type to a short keyword (image, hyperlink, chart, etc)
                        // This is optional but very useful.
                        let simplifiedType = "other";
                        // Check image
                        if (typeAttr.includes("relationships/image")) {
                            simplifiedType = "image";
                        }
                        // Check hyperlink
                        else if (typeAttr.includes("relationships/hyperlink")) {
                            simplifiedType = "hyperlink";
                        }
                        // Check chart
                        else if (typeAttr.includes("relationships/chart")) {
                            simplifiedType = "chart";
                        }
                        // Check slide references
                        else if (typeAttr.includes("relationships/slide")) {
                            simplifiedType = "slide";
                        }
                        // Check notes
                        else if (typeAttr.includes("relationships/notesSlide")) {
                            simplifiedType = "notes";
                        }
                        // Now normalize the target only if it is a local file path.
                        // Hyperlinks are external and should not be normalized.
                        let normalizedTarget = targetRaw;
                        // Local paths never contain "http" or "https"
                        const isExternal = targetRaw.startsWith("http://") || targetRaw.startsWith("https://");

                        // If not external, normalize the target which is just the name of the item.
                        if (!isExternal) {
                            normalizedTarget = normalizedTarget.split('/').pop() || '';
                        }
                        // Finally store full relationship info
                        slideRelsMap[slideNum][id] = {
                            type: simplifiedType,
                            target: normalizedTarget
                        };
                    }
                }
            }
        }
    }

    // Now for processing all the other files - slides and notes.
    for (const file of files) {
        if (file.path.match(mediaFileRegex)) continue;
        if (file.path.match(chartFileRegex)) continue;
        if (file.path.match(slideRelsRegex)) continue;
        if (file.path.match(corePropsFileRegex)) continue;

        const xmlContentString = file.content.toString();
        const xml = parseXmlString(xmlContentString);
        if (config.includeRawContent) {
            rawContents.push(xmlContentString);
        }

        const slideMatch = file.path.match(slideNumberRegex);
        const slideNumber = slideMatch ? parseInt(slideMatch[1]) : 0;
        const isNote = file.path.includes("notesSlide");

        const slideNode: OfficeContentNode = {
            type: isNote ? 'note' : 'slide',
            children: [],
            metadata: {
                slideNumber: slideNumber,
                ...(isNote ? { noteId: `slide-note-${slideNumber}` } : {})
            } as SlideMetadata
        };

        if (config.includeRawContent) {
            slideNode.rawContent = file.content.toString();
        }

        /**
         * Extract slide contents in correct document order by scanning p:spTree children.
         * This ensures p:pic, p:sp, p:graphicFrame appear in AST in the exact sequence.
         */
        const spTree = getElementsByTagName(xml, "p:spTree")[0];
        if (spTree) {
            slideNode.children?.push(...traverseSpTree(spTree, slideNumber));
        }

        if (slideNode.children && slideNode.children.length > 0) {
            content.push(slideNode);
        }
    }

    const attachments: OfficeAttachment[] = [];
    const mediaFiles = files.filter(f => f.path.match(/ppt\/media\/.*/));
    const chartFiles = files.filter(f => f.path.match(/ppt\/charts\/chart\d+\.xml/));

    // First run to extract attachments and to assign ocr to image files.
    if (config.extractAttachments) {
        // Extract media files as attachments
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

        // Extract chart files as attachments
        for (const chart of chartFiles) {
            const attachment: OfficeAttachment = {
                type: 'chart',
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // Generic XML type for now
                data: chart.content.toString('base64'),
                name: chart.path.split('/').pop() || '',
                extension: 'xml'
            };
            attachments.push(attachment);

            // Extract text from chart XML
            try {
                const chartData = await extractChartData(chart.content);
                // Assign chartData to attachment
                attachment.chartData = chartData;
            }
            catch (e) {
                logWarning(`Failed to extract text from chart ${chart.path}:`, config, e);
            }
        }

        // Loop through nodes to find images and charts and link their text and chartData
        const assignAttachmentData = (nodes: OfficeContentNode[]) => {
            for (const node of nodes) {
                if ('attachmentName' in (node.metadata || {})) {
                    const meta = node.metadata as ImageMetadata | ChartMetadata;
                    const attachment = attachments.find(a => a.name === meta.attachmentName);
                    if (attachment) {
                        if (node.type === 'image') {
                            attachment.altText = (meta as ImageMetadata).altText;
                            if (attachment.ocrText)
                                node.text = attachment.ocrText;
                        }
                        if (node.type === 'chart') {
                            node.text = attachment.chartData?.rawTexts.join(config.newlineDelimiter || '\n');
                        }
                    }
                }
                if (node.children) {
                    assignAttachmentData(node.children);
                }
            }
        };
        assignAttachmentData(content);
    }

    // Finally, if the notes are required to be at the end of the document, move them there.
    if (!config.ignoreNotes && config.putNotesAtLast) {
        content.sort((a, b) => {
            const aIsNote = a.type === 'note' ? 1 : 0;
            const bIsNote = b.type === 'note' ? 1 : 0;
            return aIsNote - bIsNote;
        });
    }

    return {
        type: 'pptx',
        metadata: metadata,
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

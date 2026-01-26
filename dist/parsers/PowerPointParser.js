"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.parsePowerPoint = void 0;
const xmldom_1 = require("@xmldom/xmldom");
const chartUtils_1 = require("../utils/chartUtils");
const errorUtils_1 = require("../utils/errorUtils");
const imageUtils_1 = require("../utils/imageUtils");
const ocrUtils_1 = require("../utils/ocrUtils");
const xmlUtils_1 = require("../utils/xmlUtils");
const zipUtils_1 = require("../utils/zipUtils");
/**
 * Parses a PowerPoint presentation (.pptx) and extracts slides and notes.
 *
 * @param buffer - The PPTX file as a Buffer
 * @param config - Parser configuration
 * @returns A promise resolving to the parsed AST
 */
const parsePowerPoint = async (buffer, config) => {
    const allFilesRegex = /ppt\/(notesSlides|slides)\/(notesSlide|slide)\d+.xml/g;
    const slidesRegex = /ppt\/slides\/slide\d+.xml/g;
    const slideRelsRegex = /ppt\/slides\/_rels\/slide\d+\.xml\.rels/;
    const slideNumberRegex = /lide(\d+)\.xml/;
    const mediaFileRegex = /ppt\/media\/.*/;
    const chartFileRegex = /ppt\/charts\/chart\d+\.xml/;
    const corePropsFileRegex = /docProps\/core\.xml/;
    const xmlSerializer = new xmldom_1.XMLSerializer();
    const files = await (0, zipUtils_1.extractFiles)(buffer, x => !!x.match(config.ignoreNotes ? slidesRegex : allFilesRegex) ||
        !!x.match(corePropsFileRegex) ||
        !!x.match(slideRelsRegex) ||
        (!!config.extractAttachments && (!!x.match(mediaFileRegex) || !!x.match(chartFileRegex))));
    // Extract metadata
    const corePropsFile = files.find(f => f.path.match(corePropsFileRegex));
    const metadata = corePropsFile ? (0, xmlUtils_1.parseOfficeMetadata)(corePropsFile.content.toString()) : {};
    // Sort files
    files.sort((a, b) => {
        const aMatch = a.path.match(slideNumberRegex);
        const bMatch = b.path.match(slideNumberRegex);
        const aNum = aMatch ? parseInt(aMatch[1]) : 0;
        const bNum = bMatch ? parseInt(bMatch[1]) : 0;
        return aNum - bNum;
    });
    const content = [];
    const rawContents = [];
    const slideRelsMap = {};
    // Map to track OfficeContentNode -> XML Element for position extraction
    const elementMap = new Map();
    let currentListId = 0;
    let runningListIndex = 0;
    let lastWasList = false;
    let lastListType = null;
    let lastListIndent = 0;
    // per indent counters (for nested lists)
    const levelCounters = {};
    /**
     * Extracts coordinate data from PPTX XML transform elements.
     * Looks for a:xfrm or p:xfrm containing a:off (x, y) and a:ext (width, height).
     * Values are in EMU (English Metric Units): 1 inch = 914400 EMU.
     */
    const extractCoordinates = (element) => {
        // Try p:xfrm first (PowerPoint namespace), then a:xfrm (DrawingML namespace)
        let xfrm = (0, xmlUtils_1.getElementsByTagName)(element, "p:xfrm")[0];
        if (!xfrm) {
            xfrm = (0, xmlUtils_1.getElementsByTagName)(element, "a:xfrm")[0];
        }
        if (!xfrm) {
            return undefined;
        }
        // Extract offset (x, y)
        const off = (0, xmlUtils_1.getElementsByTagName)(xfrm, "a:off")[0];
        const x = off ? parseInt(off.getAttribute("x") || "0", 10) : 0;
        const y = off ? parseInt(off.getAttribute("y") || "0", 10) : 0;
        // Extract extent (width, height)
        const ext = (0, xmlUtils_1.getElementsByTagName)(xfrm, "a:ext")[0];
        const width = ext ? parseInt(ext.getAttribute("cx") || "0", 10) : 0;
        const height = ext ? parseInt(ext.getAttribute("cy") || "0", 10) : 0;
        // Extract rotation (in 60000ths of a degree)
        const rot = xfrm.getAttribute("rot");
        const rotation = rot ? parseInt(rot, 10) / 60000 : undefined;
        // Only return if we have valid dimensions
        if (width > 0 && height > 0) {
            return {
                x,
                y,
                width,
                height,
                ...(rotation !== undefined ? { rotation } : {})
            };
        }
        return undefined;
    };
    // Helper to parse a table node
    const parseTable = (tblNode) => {
        const rows = [];
        const trNodes = (0, xmlUtils_1.getElementsByTagName)(tblNode, "a:tr");
        for (let rIndex = 0; rIndex < trNodes.length; rIndex++) {
            const trNode = trNodes[rIndex];
            const cells = [];
            const tcNodes = (0, xmlUtils_1.getElementsByTagName)(trNode, "a:tc");
            for (let cIndex = 0; cIndex < tcNodes.length; cIndex++) {
                const tcNode = tcNodes[cIndex];
                const cellChildren = [];
                let cellText = '';
                // Cells contain text bodies (txBody) which contain paragraphs
                const txBody = (0, xmlUtils_1.getElementsByTagName)(tcNode, "a:txBody")[0];
                if (txBody) {
                    const paragraphs = (0, xmlUtils_1.getElementsByTagName)(txBody, "a:p");
                    for (const p of paragraphs) {
                        // Reuse paragraph parsing logic if possible, or duplicate for now
                        // For simplicity, duplicating basic logic here as the main loop one is tied to shapes
                        const pNode = {
                            type: 'paragraph',
                            text: '',
                            children: [],
                            metadata: {}
                        };
                        if (config.includeRawContent) {
                            pNode.rawContent = p.toString();
                        }
                        const runs = (0, xmlUtils_1.getElementsByTagName)(p, "a:r");
                        for (const r of runs) {
                            const t = (0, xmlUtils_1.getElementsByTagName)(r, "a:t")[0];
                            if (t && t.childNodes[0]) {
                                const textContent = t.childNodes[0].nodeValue || '';
                                pNode.text += textContent;
                                const rPr = (0, xmlUtils_1.getElementsByTagName)(r, "a:rPr")[0];
                                const formatting = {};
                                if (rPr) {
                                    if (rPr.getAttribute("b") === "1")
                                        formatting.bold = true;
                                    if (rPr.getAttribute("i") === "1")
                                        formatting.italic = true;
                                    if (rPr.getAttribute("u") === "sng")
                                        formatting.underline = true;
                                    if (rPr.getAttribute("strike") === "sngStrike")
                                        formatting.strikethrough = true;
                                    const sz = rPr.getAttribute("sz");
                                    if (sz)
                                        formatting.size = (parseInt(sz) / 100).toString() + 'pt';
                                    const solidFill = (0, xmlUtils_1.getElementsByTagName)(rPr, "a:solidFill")[0];
                                    if (solidFill) {
                                        const srgbClr = (0, xmlUtils_1.getElementsByTagName)(solidFill, "a:srgbClr")[0];
                                        if (srgbClr) {
                                            const val = srgbClr.getAttribute("val");
                                            if (val)
                                                formatting.color = '#' + val;
                                        }
                                    }
                                    const latin = (0, xmlUtils_1.getElementsByTagName)(rPr, "a:latin")[0];
                                    if (latin) {
                                        const typeface = latin.getAttribute("typeface");
                                        if (typeface)
                                            formatting.font = typeface;
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
                const cellNode = {
                    type: 'cell',
                    text: cellText,
                    children: cellChildren,
                    metadata: { row: rIndex, col: cIndex }
                };
                cells.push(cellNode);
            }
            const rowNode = {
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
    const extractImageNode = (imageNode, slideNumber) => {
        const blip = (0, xmlUtils_1.getElementsByTagName)(imageNode, "a:blip")[0];
        if (!blip)
            return null;
        const rId = blip.getAttribute("r:embed");
        if (!rId)
            return null;
        const rel = slideRelsMap[slideNumber]?.[rId];
        if (!rel || rel.type !== "image")
            return null;
        const attachmentName = rel.target;
        const nvPicPr = (0, xmlUtils_1.getElementsByTagName)(imageNode, "p:nvPicPr")[0];
        const cNvPr = nvPicPr ? (0, xmlUtils_1.getElementsByTagName)(nvPicPr, "p:cNvPr")[0] : null;
        const altText = cNvPr?.getAttribute("descr") || undefined;
        const node = {
            type: "image",
            text: '',
            metadata: {
                attachmentName,
                altText,
            }
        };
        // Track element for position extraction
        elementMap.set(node, imageNode);
        return node;
    };
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
    const extractChartNode = (frameNode, slideNumber) => {
        // Step 1: Find <a:graphicData>
        const graphicData = (0, xmlUtils_1.getElementsByTagName)(frameNode, "a:graphicData")[0];
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
        const cChart = (0, xmlUtils_1.getElementsByTagName)(graphicData, "c:chart")[0];
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
        const chartNode = {
            type: "chart",
            text: "",
            metadata: {
                attachmentName // name used to link to attachments & chartData
            }
        };
        // Optional: include raw XML of the whole frame
        if (config.includeRawContent) {
            chartNode.rawContent = xmlSerializer.serializeToString(frameNode);
        }
        // Track element for position extraction
        elementMap.set(chartNode, frameNode);
        return chartNode;
    };
    /** Extract an AST node for p:graphicFrame */
    const extractGraphicFrameNode = (frameNode, slideNumber) => {
        const tbl = (0, xmlUtils_1.getElementsByTagName)(frameNode, "a:tbl")[0];
        if (tbl) {
            const tableNode = parseTable(tbl);
            // Track element for position extraction
            elementMap.set(tableNode, frameNode);
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
    };
    /** Extract text and hyperlinks from a p:sp shape. */
    const extractShapeNodes = (spNode, slideNumber) => {
        const nodes = [];
        // Check for placeholder type (title, body, etc.)
        const nvSpPr = (0, xmlUtils_1.getElementsByTagName)(spNode, "p:nvSpPr")[0];
        const nvPr = nvSpPr ? (0, xmlUtils_1.getElementsByTagName)(nvSpPr, "p:nvPr")[0] : null;
        const ph = nvPr ? (0, xmlUtils_1.getElementsByTagName)(nvPr, "p:ph")[0] : null;
        const type = ph ? ph.getAttribute("type") : "body";
        const isTitle = type === "title" || type === "ctrTitle";
        const txBody = (0, xmlUtils_1.getElementsByTagName)(spNode, "p:txBody")[0];
        if (txBody) {
            const paragraphs = (0, xmlUtils_1.getElementsByTagName)(txBody, "a:p");
            for (let i = 0; i < paragraphs.length; i++) {
                const p = paragraphs[i];
                const pNode = {
                    type: isTitle ? 'heading' : 'paragraph',
                    text: '',
                    children: [],
                    metadata: isTitle ? { level: 1 } : {}
                };
                // Track element for position extraction (use spNode for shape position)
                // All paragraphs in the same shape share the shape's position
                elementMap.set(pNode, spNode);
                // Paragraph Alignment and List Detection
                const pPr = (0, xmlUtils_1.getElementsByTagName)(p, "a:pPr")[0];
                let isList = false;
                let listType = 'unordered';
                let lvl = 0;
                if (pPr) {
                    const lvlAttr = pPr.getAttribute("lvl");
                    if (lvlAttr)
                        lvl = parseInt(lvlAttr);
                    const buAutoNum = (0, xmlUtils_1.getElementsByTagName)(pPr, "a:buAutoNum")[0];
                    const buChar = (0, xmlUtils_1.getElementsByTagName)(pPr, "a:buChar")[0];
                    const buBlip = (0, xmlUtils_1.getElementsByTagName)(pPr, "a:buBlip")[0];
                    const buNode = (0, xmlUtils_1.getElementsByTagName)(pPr, "a:bu")[0];
                    if (buAutoNum) {
                        isList = true;
                        listType = 'ordered';
                    }
                    else if (buChar || buBlip) {
                        isList = true;
                        listType = 'unordered';
                    }
                    else if (buNode) {
                        // inherited bullet from a list style
                        isList = true;
                        listType = 'unordered';
                    }
                    const algn = pPr.getAttribute("algn");
                    if (algn) {
                        const alignMap = {
                            'l': 'left',
                            'ctr': 'center',
                            'r': 'right',
                            'just': 'justify'
                        };
                        if (alignMap[algn]) {
                            pNode.metadata.alignment = alignMap[algn];
                        }
                    }
                }
                if (isList) {
                    pNode.type = 'list';
                    const ilvl = lvl;
                    // detect a new list when bullet type changes or previous was not a list
                    const newList = !lastWasList ||
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
                                if (lv > ilvl)
                                    delete levelCounters[lv];
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
                        alignment: pNode.metadata?.alignment || 'left',
                    };
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
                const runs = (0, xmlUtils_1.getElementsByTagName)(p, "a:r");
                for (let j = 0; j < runs.length; j++) {
                    const r = runs[j];
                    const t = (0, xmlUtils_1.getElementsByTagName)(r, "a:t")[0];
                    if (t && t.childNodes[0]) {
                        const textContent = t.childNodes[0].nodeValue || '';
                        pNode.text += textContent;
                        const rPr = (0, xmlUtils_1.getElementsByTagName)(r, "a:rPr")[0];
                        const formatting = {};
                        if (rPr) {
                            if (rPr.getAttribute("b") === "1")
                                formatting.bold = true;
                            if (rPr.getAttribute("i") === "1")
                                formatting.italic = true;
                            if (rPr.getAttribute("u") === "sng")
                                formatting.underline = true;
                            if (rPr.getAttribute("strike") === "sngStrike")
                                formatting.strikethrough = true;
                            const sz = rPr.getAttribute("sz");
                            if (sz)
                                formatting.size = (parseInt(sz) / 100).toString() + 'pt';
                            // Color extraction
                            const solidFill = (0, xmlUtils_1.getElementsByTagName)(rPr, "a:solidFill")[0];
                            if (solidFill) {
                                const srgbClr = (0, xmlUtils_1.getElementsByTagName)(solidFill, "a:srgbClr")[0];
                                if (srgbClr) {
                                    const val = srgbClr.getAttribute("val");
                                    if (val)
                                        formatting.color = '#' + val;
                                }
                            }
                            // Highlight extraction
                            const highlight = (0, xmlUtils_1.getElementsByTagName)(rPr, "a:highlight")[0];
                            if (highlight) {
                                const srgbClr = (0, xmlUtils_1.getElementsByTagName)(highlight, "a:srgbClr")[0];
                                if (srgbClr) {
                                    const val = srgbClr.getAttribute("val");
                                    if (val)
                                        formatting.backgroundColor = '#' + val;
                                }
                            }
                            // Font family
                            const latin = (0, xmlUtils_1.getElementsByTagName)(rPr, "a:latin")[0];
                            if (latin) {
                                const typeface = latin.getAttribute("typeface");
                                if (typeface)
                                    formatting.font = typeface;
                            }
                            // Subscript/Superscript
                            const baseline = rPr.getAttribute("baseline");
                            if (baseline) {
                                const baselineVal = parseInt(baseline);
                                if (baselineVal < 0)
                                    formatting.subscript = true;
                                if (baselineVal > 0)
                                    formatting.superscript = true;
                            }
                        }
                        const textNode = {
                            type: 'text',
                            text: textContent,
                            formatting: formatting
                        };
                        // Check for Hyperlinks
                        const hlinkClick = (0, xmlUtils_1.getElementsByTagName)(r, "a:hlinkClick")[0];
                        // Check if this run has a hyperlink click action
                        if (hlinkClick) {
                            // Relationship ID for the link
                            const rId = hlinkClick.getAttribute("r:id");
                            // Optional action attribute, often for internal jumps
                            const action = hlinkClick.getAttribute("action");
                            // Result placeholders
                            let link;
                            let linkType;
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
    };
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
    function traverseSpTree(treeNode, slideNumber) {
        const nodes = [];
        // Process children in XML order (this preserves Z-order)
        for (const child of Array.from(treeNode?.childNodes || [])) {
            if (child.nodeType !== 1) {
                continue;
            }
            const element = child;
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
                const relsXml = (0, xmlUtils_1.parseXmlString)(file.content.toString());
                // Get all Relationship nodes
                const relationships = (0, xmlUtils_1.getElementsByTagName)(relsXml, "Relationship");
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
        if (file.path.match(mediaFileRegex))
            continue;
        if (file.path.match(chartFileRegex))
            continue;
        if (file.path.match(slideRelsRegex))
            continue;
        if (file.path.match(corePropsFileRegex))
            continue;
        const xmlContentString = file.content.toString();
        const xml = (0, xmlUtils_1.parseXmlString)(xmlContentString);
        if (config.includeRawContent) {
            rawContents.push(xmlContentString);
        }
        const slideMatch = file.path.match(slideNumberRegex);
        const slideNumber = slideMatch ? parseInt(slideMatch[1]) : 0;
        const isNote = file.path.includes("notesSlide");
        const slideNode = {
            type: isNote ? 'note' : 'slide',
            children: [],
            metadata: {
                slideNumber: slideNumber,
                ...(isNote ? { noteId: `slide-note-${slideNumber}` } : {})
            }
        };
        if (config.includeRawContent) {
            slideNode.rawContent = file.content.toString();
        }
        /**
         * Extract slide contents in correct document order by scanning p:spTree children.
         * This ensures p:pic, p:sp, p:graphicFrame appear in AST in the exact sequence.
         */
        const spTree = (0, xmlUtils_1.getElementsByTagName)(xml, "p:spTree")[0];
        if (spTree) {
            slideNode.children?.push(...traverseSpTree(spTree, slideNumber));
        }
        if (slideNode.children && slideNode.children.length > 0) {
            content.push(slideNode);
        }
    }
    const attachments = [];
    const mediaFiles = files.filter(f => f.path.match(/ppt\/media\/.*/));
    const chartFiles = files.filter(f => f.path.match(/ppt\/charts\/chart\d+\.xml/));
    // First run to extract attachments and to assign ocr to image files.
    if (config.extractAttachments) {
        // Extract media files as attachments
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
        // Extract chart files as attachments
        for (const chart of chartFiles) {
            const attachment = {
                type: 'chart',
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                data: chart.content.toString('base64'),
                name: chart.path.split('/').pop() || '',
                extension: 'xml'
            };
            attachments.push(attachment);
            // Extract text from chart XML
            try {
                const chartData = (0, chartUtils_1.extractChartData)(chart.content);
                // Assign chartData to attachment
                attachment.chartData = chartData;
            }
            catch (e) {
                (0, errorUtils_1.logWarning)(`Failed to extract text from chart ${chart.path}:`, config, e);
            }
        }
        // Loop through nodes to find images and charts and link their text and chartData
        const assignAttachmentData = (nodes) => {
            for (const node of nodes) {
                if ('attachmentName' in (node.metadata || {})) {
                    const meta = node.metadata;
                    const attachment = attachments.find(a => a.name === meta.attachmentName);
                    if (attachment) {
                        if (node.type === 'image') {
                            attachment.altText = meta.altText;
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
    /**
     * Converts a table node to a TableBlock.
     */
    const convertTableToBlock = (tableNode, position) => {
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
            rows,
            ...(position ? { position } : {})
        };
    };
    /**
     * Converts a chart node to a ChartBlock.
     */
    const convertChartToBlock = (chartNode, attachments, position) => {
        if (chartNode.type !== 'chart')
            return null;
        const chartMetadata = chartNode.metadata;
        // Try to get chartData from metadata first (if it was stored there)
        const metadataWithChartData = chartMetadata;
        if (metadataWithChartData?.chartData) {
            return {
                type: 'chart',
                chartData: metadataWithChartData.chartData,
                chartType: metadataWithChartData.chartData.chartType,
                ...(position ? { position } : {})
            };
        }
        // Otherwise, try to find it from attachments
        if (chartMetadata?.attachmentName) {
            const attachment = attachments.find(a => a.name === chartMetadata.attachmentName && a.type === 'chart');
            if (attachment?.chartData) {
                return {
                    type: 'chart',
                    chartData: attachment.chartData,
                    chartType: attachment.chartData.chartType,
                    ...(position ? { position } : {})
                };
            }
        }
        return null;
    };
    /**
     * Extracts blocks from content nodes in document order.
     * Ensures all content types (tables, charts, images, text) are captured as blocks.
     */
    const extractBlocksFromContent = (nodes, attachments, elementMap) => {
        const blocks = [];
        const traverse = (node) => {
            // Get position from element map
            const element = elementMap.get(node);
            const position = element ? extractCoordinates(element) : undefined;
            // Process node based on type - prioritize specific block types
            if (node.type === 'table') {
                blocks.push(convertTableToBlock(node, position));
                // Don't traverse children of tables (already processed in convertTableToBlock)
                return;
            }
            else if (node.type === 'chart') {
                const chartBlock = convertChartToBlock(node, attachments, position);
                if (chartBlock) {
                    blocks.push(chartBlock);
                }
                // Don't traverse children of charts
                return;
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
                            filename: attachment.name,
                            ...(position ? { position } : {})
                        });
                    }
                }
                // Don't traverse children of images
                return;
            }
            // For text content: create text blocks for paragraphs, headings, lists, and text nodes
            // Skip slides/notes containers and other structural nodes that are just containers
            if (node.text && node.text.trim()) {
                // Create text blocks for content nodes (not containers like slide/note/row/cell)
                if (node.type === 'text' || node.type === 'paragraph' || node.type === 'heading' || node.type === 'list') {
                    // Create text block - include all text content
                    blocks.push({
                        type: 'text',
                        content: node.text.trim(),
                        ...(position ? { position } : {})
                    });
                }
            }
            // Recursively process children (for slides, notes, and other container nodes)
            // This ensures we traverse the entire tree and don't miss any content
            if (node.children) {
                for (const child of node.children) {
                    traverse(child);
                }
            }
        };
        // Traverse all top-level nodes (slides, notes)
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
        // Recursive text extraction
        const getText = (node) => {
            let t = '';
            if (node.children) {
                t += node.children.map(getText).filter(t => t != '').join(!node.children[0]?.children ? '' : config.newlineDelimiter ?? '\n');
            }
            else
                t += node.text || '';
            return t;
        };
        return getText(c);
    }).filter(t => t != '').join(config.newlineDelimiter ?? '\n');
    // Extract blocks
    const blocks = extractBlocksFromContent(content, attachments, elementMap);
    // Extract images
    const images = extractImagesList(attachments);
    console.log("Document blocks", blocks.filter(x => x.type === "chart"));
    return {
        type: 'pptx',
        metadata: metadata,
        content: content,
        attachments: attachments,
        fullText,
        blocks,
        images,
        toText: () => fullText
    };
};
exports.parsePowerPoint = parsePowerPoint;

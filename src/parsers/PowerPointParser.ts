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

import { attachmentLookup } from '../utils/repeatUtils.js';
import { BreakMetadata, ChartMetadata, CodeMetadata, CommentMetadata, FullOfficeParserConfig, HeadingMetadata, ImageMetadata, ListMetadata, OfficeAttachment, OfficeContentNode, OfficeParserAST, OfficeWarningType, SlideMetadata, TextFormatting } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { extractChartData } from '../utils/chartUtils.js';
import { checkAbortSignal, logWarning } from '../utils/errorUtils.js';
import { createAttachment } from '../utils/imageUtils.js';
import { isEmptyMath, ommlToLatex } from '../utils/mathUtils.js';
import { ocrDuringParse } from '../utils/ocrUtils.js';
import { getChildElements, getDirectChildren, getElementsByTagName, getFirstElementByTagName, getRawContent, isElement, parseOfficeMetadata, parseOOXMLAppProperties, parseOOXMLCustomProperties, parseXmlString } from '../utils/xmlUtils.js';
import { extractFiles, findRequiredPart, resolvePartPath } from '../utils/zipUtils.js';
import { lookupTable } from '../utils/lookupUtils.js';
import { appendAll } from '../utils/nodeListUtils.js';
import { diagramList, readDiagram } from '../utils/diagramUtils.js';
import { alternateContentBranch, PRESENTATION_NAMESPACES } from '../utils/markupCompatibility.js';

/** An element's own text (its text and CDATA children), not its descendants': a classic comment's `p:text` is a plain string. */
const ownText = (element: Element): string => {
    const parts: string[] = [];
    for (let i = 0; i < element.childNodes.length; i++) {
        const child = element.childNodes[i];
        if (child.nodeType === 3 || child.nodeType === 4) parts.push(child.nodeValue || '');
    }
    return parts.join('');
};

/** A comment's paragraph of `children`, its text theirs (a line break a new line). */
const commentParagraph = (children: OfficeContentNode[]): OfficeContentNode => ({
    type: 'paragraph',
    text: children.map(child => child.type === 'break' ? '\n' : child.text ?? '').join(''),
    children,
});

/** A modern comment's DrawingML paragraph (`a:p`): its runs' and fields' text, and its line breaks. */
const drawingParagraph = (paragraph: Element): OfficeContentNode => {
    const children: OfficeContentNode[] = [];
    for (let i = 0; i < paragraph.childNodes.length; i++) {
        const child = paragraph.childNodes[i];
        if (!isElement(child)) continue;
        if (child.tagName === 'a:r' || child.tagName === 'a:fld') {
            const text = getChildElements(child, 'a:t').map(t => t.textContent || '').join('');
            if (text) children.push({ type: 'text', text });
        } else if (child.tagName === 'a:br') {
            children.push({ type: 'break', metadata: { breakType: 'textWrapping' } as BreakMetadata });
        }
    }
    return commentParagraph(children);
};

/**
 * Parses a PowerPoint presentation (.pptx) and extracts slides and notes.
 * 
 * @param buffer - The PPTX file as a Buffer
 * @param config - Parser configuration
 * @returns A promise resolving to the parsed AST
 */
export const parsePowerPoint = async (buffer: Buffer, config: FullOfficeParserConfig): Promise<OfficeParserAST> => {
    // Honour cancellation requests immediately, before extracting the ZIP archive.
    // PPTX presentations can have many slides with media/charts and optional OCR per image,
    // so an early abort prevents decompressing and traversing data that will be discarded.
    checkAbortSignal(config.abortSignal);

    const allFilesRegex = /ppt\/(notesSlides|slides)\/(notesSlide|slide)\d+.xml/g;
    const slidesRegex = /ppt\/slides\/slide\d+.xml/g;
    const slideRelsRegex = /ppt\/slides\/_rels\/slide\d+\.xml\.rels/;
    // The relationships of the other parts read by them: a notes page's and a slide master's own
    // (their `r:id`s are not a slide's), and the presentation's, which name its slides in order.
    const notesRelsRegex = /ppt\/notesSlides\/_rels\/notesSlide\d+\.xml\.rels/;
    const slideMasterRelsRegex = /ppt\/slideMasters\/_rels\/slideMaster\d+\.xml\.rels/;
    const presentationRelsRegex = /ppt\/_rels\/presentation\.xml\.rels/;
    /** A relationships part: the folder of the part it belongs to, and that part's name. */
    const relsPartRegex = /^(.*\/)?_rels\/([^/]+)\.rels$/;
    const slideNumberRegex = /lide(\d+)\.xml/;
    const mediaFileRegex = /ppt\/media\/.*/;
    const chartFileRegex = /ppt\/charts\/chart\d+\.xml/;
    const corePropsFileRegex = /docProps\/core\.xml/;
    const customPropsFileRegex = /docProps\/custom\.xml/;
    const appPropsFileRegex = /docProps\/app\.xml/;

    // Comments as PowerPoint wrote them until 2019 (`comment1.xml`, authors in `commentAuthors.xml`) and
    // since (`modernComment_*.xml`, `p188:cm`, authors in `authors.xml`), which newer versions write alone.
    const commentsFileRegex = /ppt\/comments\/(?:comment\d+|modernComment_[^/]+)\.xml/;
    const commentAuthorsRegex = /ppt\/(?:commentAuthors|authors)\.xml/;
    const slideMastersRegex = /ppt\/slideMasters\/slideMaster\d+\.xml/;
    const presentationFileRegex = /ppt\/presentation\.xml/;
    const diagramDataRegex = /ppt\/diagrams\/data\d+\.xml/;

    const files = await extractFiles(
        buffer,
        x =>
            !!x.match(config.ignoreNotes ? slidesRegex : allFilesRegex) ||
            !!x.match(corePropsFileRegex) ||
            !!x.match(customPropsFileRegex) ||
            !!x.match(appPropsFileRegex) ||
            !!x.match(slideRelsRegex) ||
            (!config.ignoreNotes && !!x.match(notesRelsRegex)) ||
            (!config.ignoreComments && (!!x.match(commentsFileRegex) || !!x.match(commentAuthorsRegex))) ||
            (!config.ignoreSlideMasters && (!!x.match(slideMastersRegex) || !!x.match(slideMasterRelsRegex))) ||
            !!x.match(presentationFileRegex) ||
            !!x.match(presentationRelsRegex) ||
            !!x.match(diagramDataRegex) ||
            (!!config.extractAttachments && (!!x.match(mediaFileRegex) || !!x.match(chartFileRegex))),
        config.decompressionLimits,
        config
    );

    // ppt/presentation.xml is the part that makes an archive a presentation, and unlike the
    // slides it is always present: PowerPoint can save a deck with no slides at all, so an
    // empty ppt/slides/ is a warning rather than a failure.
    findRequiredPart(files, path => !!path.match(presentationFileRegex), config,
        { fileType: 'pptx', part: 'ppt/presentation.xml' });

    if (!files.some(file => !!file.path.match(slidesRegex)))
        logWarning(OfficeWarningType.NO_SLIDES_FOUND, config);

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
        if (Object.keys(appProperties).length > 0) metadata.nativeProperties = appProperties;
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
    /**
     * A part's relationships by id. `type` is a short keyword (image, hyperlink, chart, slide, notes,
     * comments, or other). `target` is an external target's address, and of an internal one the file
     * name alone, which is what attachments and the parts looked up by name go by; `path` is an
     * internal target's whole path in the package.
     */
    type PartRelationships = Record<string, { type: string, target: string, path?: string }>;
    // Each part's relationships, by the part's path. A part's ids are its own: a notes page's and a slide
    // master's were looked up among those of the slide whose file had the same number, so a link or a
    // picture in one was another part's (or none).
    const relsByPart = new Map<string, PartRelationships>();

    // Null-prototype, as every map keyed by the document's own ids is here.
    const authorMap: Record<string, { author?: string, initials?: string }> = Object.create(null);
    if (!config.ignoreComments) {
        const authorsFile = files.find(f => f.path === 'ppt/commentAuthors.xml');
        if (authorsFile) {
            const authorsXml = parseXmlString(authorsFile.content.toString(), { config });
            const authorNodes = getElementsByTagName(authorsXml, "p:cmAuthor");
            for (const aNode of authorNodes) {
                const id = aNode.getAttribute("id");
                if (id !== null) {
                    authorMap[id] = {
                        author: aNode.getAttribute("name") || undefined,
                        initials: aNode.getAttribute("initials") || undefined
                    };
                }
            }
        }
        const modernAuthorsFile = files.find(f => f.path === 'ppt/authors.xml');
        if (modernAuthorsFile) {
            for (const aNode of getElementsByTagName(parseXmlString(modernAuthorsFile.content.toString(), { config }), "p188:author")) {
                const id = aNode.getAttribute("id");
                if (id !== null && !(id in authorMap)) {
                    authorMap[id] = {
                        author: aNode.getAttribute("name") || undefined,
                        initials: aNode.getAttribute("initials") || undefined
                    };
                }
            }
        }
    }

    // The comments of each comments part, read once and shared by every relationship naming it: read
    // again for each, a slide naming one large comments part thousands of times (19 KB of PPTX) ran the
    // process out of memory. The writers write a shared comment once.
    const commentsByPart = new Map<string, OfficeContentNode[]>();
    const attachedCommentParts = new Set<string>();
    // The package's parts by file name (what a relationship's `target` keeps, see PartRelationships), the
    // first of each: a scan of every part for each comments target took targets x parts.
    const fileByName = new Map<string, (typeof files)[number]>();
    for (const f of files) { const name = f.path.slice(f.path.lastIndexOf('/') + 1); if (!fileByName.has(name)) fileByName.set(name, f); }
    const commentsOfPart = (target: string): OfficeContentNode[] => {
        let comments = commentsByPart.get(target);
        if (comments) return comments;
        comments = [];
        const cFile = fileByName.get(target);
        /** A comment of `paragraphs`, with its author (and initials) by `authorId`, its date, and the comment it replies to. */
        const commentOf = (paragraphs: OfficeContentNode[], authorId: string | null, info: { commentId?: string; date?: string; parentId?: string }): void => {
            if (!paragraphs.length) return;
            const authorData = authorId !== null ? authorMap[authorId] : undefined;
            const metadata: CommentMetadata = {};
            if (info.commentId) metadata.commentId = info.commentId;
            if (authorData?.author) metadata.author = authorData.author;
            if (authorData?.initials) metadata.initials = authorData.initials;
            if (info.date) metadata.date = info.date;
            if (info.parentId) metadata.parentId = info.parentId;
            comments!.push({ type: 'comment', text: paragraphs.map(p => p.text).join(' '), children: paragraphs, metadata });
        };
        if (cFile) {
            const cXml = parseXmlString(cFile.content.toString(), { config });
            // A comment as PowerPoint wrote it until 2019: its text is a plain string (`p:text`), a
            // paragraph a line, which was never read (only DrawingML text, which it holds none of, was
            // looked for). A reply (PowerPoint 2013 and later) names its comment in `p15:threadingInfo`.
            const classicId = (authorId: string | null, idx: string | null) => authorId !== null && idx !== null ? `${authorId}-${idx}` : undefined;
            for (const cNode of getElementsByTagName(cXml, "p:cm")) {
                const textNode = getChildElements(cNode, "p:text")[0];
                const paragraphs = textNode ? ownText(textNode).split(/\r\n?|\n/).filter(line => line.trim()).map(line => commentParagraph([{ type: 'text', text: line }])) : [];
                const extensions = getChildElements(cNode, "p:extLst")[0];
                let parent: Element | undefined;
                for (const extension of extensions ? getChildElements(extensions, "p:ext") : []) {
                    const threading = getChildElements(extension, "p15:threadingInfo")[0];
                    parent = threading ? getChildElements(threading, "p15:parentCm")[0] : undefined;
                    if (parent) break;
                }
                commentOf(paragraphs, cNode.getAttribute("authorId"), {
                    commentId: classicId(cNode.getAttribute("authorId"), cNode.getAttribute("idx")),
                    date: cNode.getAttribute("dt") || undefined,
                    parentId: parent ? classicId(parent.getAttribute("authorId"), parent.getAttribute("idx")) : undefined,
                });
            }
            // A modern comment, then each of its replies: a comment each, a paragraph each of their
            // text's paragraphs (joined, they ran together).
            const modernParagraphs = (holder: Element): OfficeContentNode[] => {
                const body = getChildElements(holder, "p188:txBody")[0];
                return body ? getChildElements(body, "a:p").map(drawingParagraph).filter(p => p.text?.trim()) : [];
            };
            for (const cNode of getElementsByTagName(cXml, "p188:cm")) {
                const commentId = cNode.getAttribute("id") || undefined;
                commentOf(modernParagraphs(cNode), cNode.getAttribute("authorId"), { commentId, date: cNode.getAttribute("created") || undefined });
                const replies = getChildElements(cNode, "p188:replyLst")[0];
                for (const reply of replies ? getChildElements(replies, "p188:reply") : []) {
                    commentOf(modernParagraphs(reply), reply.getAttribute("authorId"), { commentId: reply.getAttribute("id") || undefined, date: reply.getAttribute("created") || undefined, parentId: commentId });
                }
            }
        }
        commentsByPart.set(target, comments);
        return comments;
    };

    let currentListId = 0;
    let runningListIndex = 0;

    let lastWasList = false;
    let lastListType: 'ordered' | 'unordered' | null = null;
    let lastListIndent = 0;

    // per indent counters (for nested lists)
    const levelCounters: { [level: number]: number } = {};

    // Helper to parse a table node
    const parseTable = (tblNode: Element, xmlContentString: string): OfficeContentNode => {
        const rows: OfficeContentNode[] = [];
        // Each level read as children (see getChildElements): rows, cells and paragraphs nested in
        // their own kind, read as descendants, were read again at each level (2.9 KB took 18 seconds).
        const trNodes = getChildElements(tblNode, "a:tr");

        for (let rIndex = 0; rIndex < trNodes.length; rIndex++) {
            const trNode = trNodes[rIndex];
            const cells: OfficeContentNode[] = [];
            const tcNodes = getChildElements(trNode, "a:tc");

            for (let cIndex = 0; cIndex < tcNodes.length; cIndex++) {
                const tcNode = tcNodes[cIndex];
                const cellChildren: OfficeContentNode[] = [];
                let cellText = '';

                // Cells contain text bodies (txBody) which contain paragraphs
                const txBody = getChildElements(tcNode, "a:txBody")[0];
                if (txBody) {
                    const paragraphs = getChildElements(txBody, "a:p");
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
                            pNode.rawContent = getRawContent(p, xmlContentString, config);
                        }

                        const runs = getElementsByTagName(p, "a:r");
                        for (const r of runs) {
                            const t = getFirstElementByTagName(r, "a:t");
                            if (t && t.childNodes[0]) {
                                const textContent = t.childNodes[0].nodeValue || '';
                                pNode.text += textContent;

                                const rPr = getFirstElementByTagName(r, "a:rPr");
                                const formatting: TextFormatting = {};
                                if (rPr) {
                                    if (rPr.getAttribute("b") === "1") formatting.bold = true;
                                    if (rPr.getAttribute("i") === "1") formatting.italic = true;
                                    if (rPr.getAttribute("u") === "sng") formatting.underline = true;
                                    if (rPr.getAttribute("strike") === "sngStrike") formatting.strikethrough = true;
                                    const sz = rPr.getAttribute("sz");
                                    if (sz) formatting.size = (parseInt(sz) / 100).toString() + 'pt';

                                    const solidFill = getFirstElementByTagName(rPr, "a:solidFill");
                                    if (solidFill) {
                                        const srgbClr = getFirstElementByTagName(solidFill, "a:srgbClr");
                                        if (srgbClr) {
                                            const val = srgbClr.getAttribute("val");
                                            if (val) formatting.color = '#' + val;
                                        }
                                    }

                                    const latin = getFirstElementByTagName(rPr, "a:latin");
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

                let backgroundColor: string | undefined;
                const tcPr = getFirstElementByTagName(tcNode, "a:tcPr");
                if (tcPr) {
                    for (const child of Array.from(tcPr.childNodes)) {
                        if (isElement(child) && child.nodeName === "a:solidFill") {
                            const srgbClr = getFirstElementByTagName(child, "a:srgbClr");
                            if (srgbClr) {
                                const val = srgbClr.getAttribute("val");
                                if (val) backgroundColor = "#" + val;
                            }
                            break;
                        }
                    }
                }

                const cellNode: OfficeContentNode = {
                    type: 'cell',
                    text: cellText,
                    children: cellChildren,
                    metadata: { row: rIndex, col: cIndex, ...(backgroundColor ? { backgroundColor } : {}) }
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

    /**
     * Where an `a:hlinkClick` leads, for a run or a picture in a part with the given relationships: a
     * hyperlink relationship is external, a slide relationship or a bare action (such as the next slide)
     * internal.
     */
    const hlinkClickTarget = (hlinkClick: Element | null | undefined, rels: PartRelationships | undefined): { link: string; linkType: 'internal' | 'external' } | undefined => {
        if (!hlinkClick) return undefined;
        const rId = hlinkClick.getAttribute("r:id");
        const action = hlinkClick.getAttribute("action");
        const rel = rId ? rels?.[rId] : undefined;
        if (rel?.type === "hyperlink") return { link: rel.target, linkType: "external" };
        if (rel?.type === "slide") return { link: rel.target, linkType: "internal" };
        if (action) return { link: action, linkType: "internal" };
        return undefined;
    };

    /** Extract an AST node for p:pic */
    const extractImageNode = (imageNode: Element, rels: PartRelationships | undefined, xmlContentString: string): OfficeContentNode | null => {
        const blip = getFirstElementByTagName(imageNode, "a:blip");
        if (!blip) return null;

        const rId = blip.getAttribute("r:embed");
        if (!rId) return null;

        const rel = rels?.[rId];
        if (!rel || rel.type !== "image") return null;

        const attachmentName = rel.target;

        const nvPicPr = getFirstElementByTagName(imageNode, "p:nvPicPr");
        const cNvPr = nvPicPr ? getFirstElementByTagName(nvPicPr, "p:cNvPr") : null;

        const altText = cNvPr?.getAttribute("descr") || undefined;
        // A picture that is a link (its click action is a hyperlink or a jump to a slide).
        const link = hlinkClickTarget(cNvPr ? getFirstElementByTagName(cNvPr, "a:hlinkClick") : null, rels);

        return {
            type: "image",
            text: '',
            metadata:
            {
                attachmentName,
                altText,
                ...link,
            }
        };
    }

    /**
     * Extract an AST node for p:graphicFrame that contains a chart.
     * ... (comments omitted for brevity) ...
     */
    const extractChartNode = (frameNode: Element, rels: PartRelationships | undefined, xmlContentString: string): OfficeContentNode | null => {
        // Step 1: Find <a:graphicData>
        const graphicData = getFirstElementByTagName(frameNode, "a:graphicData");
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
        const cChart = getFirstElementByTagName(graphicData, "c:chart");
        if (!cChart) {
            return null;
        }

        // Step 4: Extract r:id (relationship id)
        const rId = cChart.getAttribute("r:id");
        if (!rId) {
            return null;
        }

        // Step 5: Resolve relationship target from the part's relationships
        const rel = rels?.[rId];
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
            chartNode.rawContent = getRawContent(frameNode, xmlContentString, config);
        }

        return chartNode;
    };

    /** Extract an AST node for p:graphicFrame */
    const extractGraphicFrameNode = (frameNode: Element, rels: PartRelationships | undefined, xmlContentString: string): OfficeContentNode | null => {
        const tbl = getFirstElementByTagName(frameNode, "a:tbl");
        if (tbl) {
            const tableNode = parseTable(tbl, xmlContentString);
            if (config.includeRawContent) {
                tableNode.rawContent = getRawContent(frameNode, xmlContentString, config);
            }
            if (tableNode.children && tableNode.children.length > 0) {
                return tableNode;
            }
        }
        if (frameNode.getElementsByTagName("c:chart").length > 0) {
            const chartNode = extractChartNode(frameNode, rels, xmlContentString);
            if (chartNode) {
                return chartNode;
            }
        }
        return null;
    }

    // SmartArt data parts read so far: each is shown once, on the first frame showing it (a part shown by
    // many frames is one diagram, and its text was written again at each).
    const readDiagrams = new Set<string>();
    let diagramCount = 0;
    /** A SmartArt frame's text as a list (see readDiagram); undefined for a frame that is not SmartArt. */
    const extractDiagramNodes = (frameNode: Element, rels: PartRelationships | undefined): OfficeContentNode[] | undefined => {
        const relIds = getFirstElementByTagName(frameNode, "dgm:relIds");
        if (!relIds) return undefined;
        const rId = relIds.getAttribute("r:dm");
        const target = rId ? rels?.[rId]?.target : undefined;
        const file = target ? fileByName.get(target) : undefined;
        if (!file || !diagramDataRegex.test(file.path) || readDiagrams.has(file.path)) return [];
        readDiagrams.add(file.path);
        return diagramList(readDiagram(file.content.toString(), config), `smartart-${++diagramCount}`);
    };

    /** Extract text and hyperlinks from a p:sp shape. */
    const extractShapeNodes = (spNode: Element, rels: PartRelationships | undefined, xmlContentString: string): OfficeContentNode[] => {
        const nodes: OfficeContentNode[] = [];
        // Check for placeholder type (title, body, etc.)
        const nvSpPr = getFirstElementByTagName(spNode, "p:nvSpPr");
        const nvPr = nvSpPr ? getFirstElementByTagName(nvSpPr, "p:nvPr") : null;
        const ph = nvPr ? getFirstElementByTagName(nvPr, "p:ph") : null;
        const type = ph ? ph.getAttribute("type") : "body";

        const isTitle = type === "title" || type === "ctrTitle";

        const txBody = getChildElements(spNode, "p:txBody")[0];
        if (txBody) {
            const paragraphs = getChildElements(txBody, "a:p");

            for (let i = 0; i < paragraphs.length; i++) {
                const p = paragraphs[i];
                let pNode: OfficeContentNode;
                if (isTitle) {
                    pNode = {
                        type: 'heading',
                        text: '',
                        children: [],
                        metadata: { level: 1 }
                    };
                } else {
                    pNode = {
                        type: 'paragraph',
                        text: '',
                        children: [],
                        metadata: {}
                    };
                }

                // Paragraph Alignment and List Detection
                const pPr = getChildElements(p, "a:pPr")[0];
                let isList = false;
                let listType: 'ordered' | 'unordered' = 'unordered';
                let lvl = 0;

                if (pPr) {
                    const lvlAttr = pPr.getAttribute("lvl");
                    // DrawingML's levels are 0 to 8.
                    if (lvlAttr) { const value = parseInt(lvlAttr, 10); lvl = Number.isFinite(value) ? Math.max(0, Math.min(8, value)) : 0; }

                    const buAutoNum = getChildElements(pPr, "a:buAutoNum")[0];
                    const buChar = getChildElements(pPr, "a:buChar")[0];
                    const buBlip = getChildElements(pPr, "a:buBlip")[0];
                    const buNode = getFirstElementByTagName(pPr, "a:bu");

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
                        const alignMap: Record<string, 'left' | 'center' | 'right' | 'justify'> = lookupTable({
                            'l': 'left',
                            'ctr': 'center',
                            'r': 'right',
                            'just': 'justify'
                        });
                        if (alignMap[algn]) {
                            (pNode.metadata as any).alignment = alignMap[algn];
                        }
                    }
                }

                if (isList) {
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
                    pNode = {
                        type: 'list',
                        text: pNode.text,
                        children: pNode.children,
                        metadata: {
                            listType,
                            indentation: ilvl,
                            listId: currentListId.toString(),
                            itemIndex: runningListIndex,
                            alignment: (pNode.metadata as any)?.alignment || 'left',
                        }
                    };
                }
                else {
                    lastWasList = false;
                    lastListType = null;
                    lastListIndent = 0;
                }

                if (isTitle && pNode.type === 'heading') {
                    pNode.metadata = { ...pNode.metadata, level: 1 } as HeadingMetadata;
                }

                if (config.includeRawContent) {
                    pNode.rawContent = getRawContent(p, xmlContentString, config);
                }

                // Process all children of <a:p> in order (runs, breaks, fields)
                const children = Array.from(p.childNodes);
                let activeNode = pNode;
                nodes.push(activeNode);

                for (const childNode of children) {
                    if (!isElement(childNode)) continue;
                    const element = childNode;
                    const tag = element.tagName;

                    if (tag === "a:r" || tag === "a:fld") {
                        const t = getFirstElementByTagName(element, "a:t");
                        if (t && t.childNodes[0]) {
                            const textContent = t.childNodes[0].nodeValue || "";
                            activeNode.text += textContent;

                            const rPr = getFirstElementByTagName(element, "a:rPr");
                            const formatting: TextFormatting = {};
                            if (rPr) {
                                if (rPr.getAttribute("b") === "1") formatting.bold = true;
                                if (rPr.getAttribute("i") === "1") formatting.italic = true;
                                if (rPr.getAttribute("u") === "sng") formatting.underline = true;
                                if (rPr.getAttribute("strike") === "sngStrike") formatting.strikethrough = true;

                                const sz = rPr.getAttribute("sz");
                                if (sz) formatting.size = (parseInt(sz) / 100).toString() + "pt";

                                // Color extraction
                                const solidFill = getFirstElementByTagName(rPr, "a:solidFill");
                                if (solidFill) {
                                    const srgbClr = getFirstElementByTagName(solidFill, "a:srgbClr");
                                    if (srgbClr) {
                                        const val = srgbClr.getAttribute("val");
                                        if (val) formatting.color = "#" + val;
                                    }
                                }

                                // Highlight extraction
                                const highlight = getFirstElementByTagName(rPr, "a:highlight");
                                if (highlight) {
                                    const srgbClr = getFirstElementByTagName(highlight, "a:srgbClr");
                                    if (srgbClr) {
                                        const val = srgbClr.getAttribute("val");
                                        if (val) formatting.backgroundColor = "#" + val;
                                    }
                                }

                                // Font family
                                const latin = getFirstElementByTagName(rPr, "a:latin");
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
                            const link = hlinkClickTarget(getFirstElementByTagName(element, "a:hlinkClick"), rels);
                            if (link) {
                                textNode.metadata = link;
                            }

                            activeNode.children?.push(textNode);
                        }
                    } else if (tag === "a:br") {
                        if (isList) {
                            // Split the list item on soft break into a paragraph node
                            activeNode = {
                                type: 'paragraph',
                                text: '',
                                children: [],
                                metadata: {
                                    paragraphIndentation: { left: lvl },
                                    alignment: (pNode.metadata as any)?.alignment || 'left'
                                }
                            };
                            if (config.includeRawContent) {
                                activeNode.rawContent = getRawContent(p, xmlContentString, config);
                            }
                            nodes.push(activeNode);
                        } else {
                            // In a normal paragraph, just add a newline
                            activeNode.text += "\n";
                            activeNode.children?.push({ type: 'text', text: "\n" });
                        }
                    } else {
                        // Equations. This loop dispatches on `a:r`/`a:fld`, so an `m:oMath` -
                        // which is a sibling of the runs, not one of them - was never visited at
                        // all and the formula vanished from the slide without a warning.
                        //
                        // PowerPoint writes the equation either directly in the paragraph or
                        // wrapped in `mc:AlternateContent`/`a14:m` for pre-2010 readers, so take
                        // the element itself when it is the equation and search inside it
                        // otherwise. `getElementsByTagName` returns document order, which is the
                        // order the equations are read in.
                        const isMath = tag === "m:oMath" || tag === "m:oMathPara";
                        const equations = isMath ? [element] : getElementsByTagName(element, "m:oMath");
                        for (const equation of equations) {
                            const latex = ommlToLatex(equation);
                            if (isEmptyMath(latex)) continue;
                            activeNode.text += latex;
                            activeNode.children?.push({
                                type: 'code',
                                text: latex,
                                metadata: { math: tag === "m:oMathPara" ? 'block' : 'inline' } as CodeMetadata
                            });
                        }
                    }
                }
            }
        }
        return nodes.filter(n => n.text?.trim() || (n.children && n.children.length > 0));
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
     * @param rels The relationships of the part the tree is in, which its `r:id`s are looked up in
     * @param xmlContentString The source XML string for raw content extraction
     */
    function traverseSpTree(treeNode: Element, rels: PartRelationships | undefined, xmlContentString: string): OfficeContentNode[] {
        const nodes: OfficeContentNode[] = [];
        // Process children in XML order (this preserves Z-order), a branch of mc:AlternateContent in its place
        const pending: Node[] = [];
        const pushChildren = (children: ArrayLike<Node>) => { for (let i = children.length - 1; i >= 0; i--) pending.push(children[i]); };
        pushChildren(treeNode?.childNodes || []);
        while (pending.length) {
            const child = pending.pop()!;
            if (!isElement(child)) {
                continue;
            }

            const element = child;
            const tag = element.tagName;

            // Case 1: Normal shape
            if (tag === "p:sp") {
                appendAll(nodes, extractShapeNodes(element, rels, xmlContentString));
            }
            // Case 2: Inline picture
            else if (tag === "p:pic") {
                const imageNode = extractImageNode(element, rels, xmlContentString);
                if (imageNode) {
                    nodes.push(imageNode);
                }
            }
            // Case 3: Chart or other graphic frame
            else if (tag === "p:graphicFrame") {
                const diagram = extractDiagramNodes(element, rels);
                if (diagram) {
                    appendAll(nodes, diagram);
                } else {
                    const tableNode = extractGraphicFrameNode(element, rels, xmlContentString);
                    if (tableNode) {
                        nodes.push(tableNode);
                    }
                }
            }
            // Case 4: Grouped shape (recursive!)
            else if (tag === "p:grpSp") {
                // Recurse into the group element itself which holds the child shapes
                appendAll(nodes, traverseSpTree(element, rels, xmlContentString));
            }
            // Case 5: Shapes written for newer readers (an equation, a 3D model, a chart of a newer kind)
            // beside a fallback for older ones: the branch this reader understands (see
            // alternateContentBranch), read in place (the whole element was dropped).
            else if (tag === "mc:AlternateContent") {
                const branch = alternateContentBranch(element, PRESENTATION_NAMESPACES);
                if (branch) pushChildren(branch.childNodes);
            }
        }
        return nodes;
    }

    type PackageFile = (typeof files)[number];

    /**
     * The XML of a part a deck is still read without: the presentation's slide list and its
     * relationships (without either, the slides are read in the order of their file numbers), and the
     * relationships of a notes page and of a slide master (without them, the links and pictures the
     * page names are not resolved). One that is not XML is reported (CONTENT_PART_NOT_READ) and
     * undefined: read as every other part is, where bad XML ends the parse, a deck whose slides are
     * all readable was refused for a part that holds none of its text. A limit the document is held to
     * as a whole ends the parse, as it does in any other part.
     */
    const readDispensablePart = (file: PackageFile, without: string): ReturnType<typeof parseXmlString> | undefined => {
        try {
            return parseXmlString(file.content.toString(), { config });
        } catch (error: any) {
            if (error?.officeIssue || error?.name === 'AbortError') throw error;
            const detail = String(error?.message ?? error).slice(0, 200);
            logWarning(OfficeWarningType.CONTENT_PART_NOT_READ, config, { part: file.path, reason: `it is not XML that can be read (${detail}), so ${without}` });
            return undefined;
        }
    };

    // First pass: each part's relationships, by the part they belong to.
    for (const file of files) {
        const relsPart = file.path.match(relsPartRegex);
        if (!relsPart) continue;
        // The folder of the part the relationships belong to, which their targets are relative to.
        const folder = relsPart[1] ?? '';
        // Null-prototype, as every map keyed by the document's own ids is here.
        const partRels: PartRelationships = Object.create(null);
        relsByPart.set(folder + relsPart[2], partRels);
        // Parse the rels XML. A slide's own relationships are how its notes, comments and pictures are
        // found: not readable, they end the parse as the slide itself would.
        const relsXml = slideRelsRegex.test(file.path)
            ? parseXmlString(file.content.toString(), { config })
            : readDispensablePart(file, presentationRelsRegex.test(file.path)
                ? 'the slides are read in the order of their file numbers'
                : 'the links and pictures its part names are not resolved');
        if (!relsXml) continue;
        // Loop through each relationship node
        for (const relationship of getElementsByTagName(relsXml, "Relationship")) {
            // Relationship ID, Example: "rId2"
            const id = relationship.getAttribute("Id");
            // Relationship Type, Example: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            const typeAttr = relationship.getAttribute("Type");
            // Raw Target, may be relative or absolute
            const targetRaw = relationship.getAttribute("Target");
            // Only proceed if ID and Type exist
            if (!id || !typeAttr || !targetRaw) continue;
            // Simplify Type to a short keyword (image, hyperlink, chart, etc)
            let simplifiedType = "other";
            if (typeAttr.includes("relationships/image")) simplifiedType = "image";
            else if (typeAttr.includes("relationships/hyperlink")) simplifiedType = "hyperlink";
            else if (typeAttr.includes("relationships/chart")) simplifiedType = "chart";
            else if (typeAttr.includes("relationships/slide")) simplifiedType = "slide";
            else if (typeAttr.includes("relationships/notesSlide")) simplifiedType = "notes";
            else if (typeAttr.includes("relationships/comments")) simplifiedType = "comments";
            // A hyperlink's address is external and kept as it is. A local target is kept as its file
            // name and as its whole path in the package (see PartRelationships).
            const isExternal = targetRaw.startsWith("http://") || targetRaw.startsWith("https://");
            partRels[id] = isExternal
                ? { type: simplifiedType, target: targetRaw }
                : { type: simplifiedType, target: targetRaw.split('/').pop() || '', path: resolvePartPath(folder, targetRaw) };
        }
    }

    const fileByPath = new Map<string, PackageFile>();
    for (const f of files) if (!fileByPath.has(f.path)) fileByPath.set(f.path, f);

    // The slides in the order the presentation shows them: its slide list (`p:sldIdLst`), whose entries
    // each name a slide part by relationship. The number in a part's name is not its place: an editor
    // that moves a slide by rewriting the list leaves the parts' names as they were, and one that
    // deletes a slide by dropping its entry can leave its part in the package. Read in the order of
    // their file numbers, such a deck's slides came out in the wrong order, the deleted ones among them.
    // A presentation whose list names no slide part there is has every slide part, in the order of
    // their numbers (which `files` is in).
    const slideParts = files.filter(f => !!f.path.match(slidesRegex));
    const isSlidePart = new Set(slideParts);
    const listedSlides = new Set<PackageFile>();
    const presentationFile = files.find(f => !!f.path.match(presentationFileRegex));
    if (presentationFile) {
        // A presentation part that cannot be read names no slides.
        const presentationXml = readDispensablePart(presentationFile, 'the slides are read in the order of their file numbers');
        const slideList = presentationXml ? getFirstElementByTagName(presentationXml, "p:sldIdLst") : undefined;
        const presentationRels = relsByPart.get(presentationFile.path);
        for (const entry of slideList ? getChildElements(slideList, "p:sldId") : []) {
            const rId = entry.getAttribute("r:id");
            const path = rId ? presentationRels?.[rId]?.path : undefined;
            const slidePart = path ? fileByPath.get(path) : undefined;
            if (slidePart && isSlidePart.has(slidePart)) listedSlides.add(slidePart);
        }
    }
    const deck = listedSlides.size > 0 ? [...listedSlides] : slideParts;

    /**
     * Fills the node of a slide, a notes page or a slide master from its part: what the part's shape
     * tree holds, read with the part's own relationships.
     */
    const readPart = (file: PackageFile, node: OfficeContentNode): OfficeContentNode => {
        const xmlContentString = file.content.toString();
        const xml = parseXmlString(xmlContentString, { config, locator: config.includeRawContent });
        if (config.includeRawContent) {
            node.rawContent = getRawContent(xml, xmlContentString, config);
        }
        const spTree = getFirstElementByTagName(xml, "p:spTree");
        if (spTree) appendAll(node.children!, traverseSpTree(spTree, relsByPart.get(file.path), xmlContentString));
        return node;
    };

    const slideMasters: OfficeContentNode[] = [];
    for (const file of files) {
        if (!file.path.match(slideMastersRegex)) continue;
        checkAbortSignal(config.abortSignal);
        const masterMatch = file.path.match(/slideMaster(\d+)\.xml/);
        const master = readPart(file, { type: 'slideMaster', children: [], metadata: { slideNumber: masterMatch ? parseInt(masterMatch[1]) : 0 } });
        if (master.children!.length > 0) slideMasters.push(master);
    }

    // Notes pages read so far: each is one slide's, read once, for the first slide naming it (read for
    // every relationship naming it, one page named thousands of times is read thousands of times).
    const readNotesPages = new Set<PackageFile>();
    for (let index = 0; index < deck.length; index++) {
        checkAbortSignal(config.abortSignal);
        const file = deck[index];
        // A slide's number is its place in the presentation, as PowerPoint numbers it.
        const slideNumber = index + 1;
        const slideNode = readPart(file, { type: 'slide', children: [], metadata: { slideNumber } });
        const slideRels = relsByPart.get(file.path);
        const relationships = slideRels ? Object.values(slideRels) : [];

        // Process comments
        if (!config.ignoreComments) {
            for (const rel of relationships) {
                if (rel.type !== "comments") continue;
                // Each comments part once, on the first slide naming it (a part is one slide's):
                // attached for every relationship, or every slide, naming it, one large part
                // was repeated thousands of times (see commentsOfPart).
                if (attachedCommentParts.has(rel.target)) continue;
                attachedCommentParts.add(rel.target);
                const comments = commentsOfPart(rel.target);
                if (!comments.length) continue;
                const slideComments = slideNode.comments ??= [];
                for (const comment of comments) slideComments.push(comment);
            }
        }

        // The slide's notes: the notes page its relationships name. A notes page's file number is not
        // its slide's (PowerPoint numbers notes pages in the order they are made, so a deck whose third
        // slide alone has notes holds `notesSlide1.xml`): matched by number, notes were given to the
        // slide before or after their own.
        for (const rel of relationships) {
            const notesFile = rel.type === "notes" && rel.path ? fileByPath.get(rel.path) : undefined;
            if (!notesFile || readNotesPages.has(notesFile)) continue;
            readNotesPages.add(notesFile);
            const note = readPart(notesFile, { type: 'note', children: [], metadata: { slideNumber, noteId: `slide-note-${slideNumber}` } });
            if (note.children!.length > 0) (slideNode.notes ??= []).push(note);
        }

        // A slide holding nothing (no content, notes or comments) is not a node.
        if (slideNode.children!.length > 0 || slideNode.notes || slideNode.comments) content.push(slideNode);
    }

    const attachments: OfficeAttachment[] = [];
    const attachmentsByName = attachmentLookup(attachments);
    const mediaFiles = files.filter(f => f.path.match(/ppt\/media\/.*/));
    const chartFiles = files.filter(f => f.path.match(/ppt\/charts\/chart\d+\.xml/));

    // First run to extract attachments and to assign ocr to image files.
    if (config.extractAttachments) {
        // Extract media files as attachments
        for (const media of mediaFiles) {
            const attachment = createAttachment(media.path.split('/').pop() || 'image', media.content);
            attachments.push(attachment);

            if (config.ocr && attachment.mimeType.startsWith('image/')) {
                const ocrText = await ocrDuringParse(media.content, config, attachment.name);
                if (ocrText !== undefined) attachment.ocrText = ocrText;
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
                const chartData = await extractChartData(chart.content, config);
                // Assign chartData to attachment
                attachment.chartData = chartData;
            }
            catch (e) {
                logWarning(OfficeWarningType.CHART_DATA_EXTRACTION_FAILED, config, chart.path, e);
            }
        }

        // Loop through nodes to find images and charts and link their text and chartData
        const assignAttachmentData = (nodes: OfficeContentNode[]) => {
            for (const node of nodes) {
                if ('attachmentName' in (node.metadata || {})) {
                    const meta = node.metadata as ImageMetadata | ChartMetadata;
                    const attachment = attachmentsByName.get(meta.attachmentName);
                    if (attachment) {
                        if (node.type === 'image') {
                            attachment.altText = (meta as ImageMetadata).altText;
                            if (attachment.ocrText)
                                node.text = attachment.ocrText;
                        }
                        if (node.type === 'chart') {
                            node.text = attachmentsByName.chartText(attachment, config.newlineDelimiter);
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

    // Notes are structurally attached to their respective slides (via node.notes), not collected at the end.

    const auxiliaryContent = slideMasters.length > 0 ? {
        slideMasters
    } : undefined;

    return createAST(
        'pptx',
        metadata,
        content,
        attachments,
        config,
        auxiliaryContent,
    );
};

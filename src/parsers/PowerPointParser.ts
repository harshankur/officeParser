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
import { BreakMetadata, ChartMetadata, CodeMetadata, CommentMetadata, FullOfficeParserConfig, HeadingMetadata, ImageMetadata, ListMetadata, OfficeAttachment, OfficeContentNode, OfficeParserAST, OfficeWarningType, ShapeMetadata, SlideMetadata, TextFormatting, TextMetadata } from '../types.js';
import { groupTransform, IDENTITY, rgbOf, roundBox, shapeLook, textFrameOf, themeColors, Transform, xfrmOf } from '../utils/drawingmlUtils.js';
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
    // A slide's layout and master (and their relationships), and the theme: where a placeholder that
    // gives no place of its own is, the background a slide takes on, and the theme's colours.
    const slideLayoutRegex = /ppt\/slideLayouts\/(?:_rels\/)?slideLayout\d+\.xml(?:\.rels)?$/;
    const themeRegex = /ppt\/theme\/theme\d+\.xml$/;
    // A notes master and its relationships: its colour map and theme give the notes' colours.
    const notesMasterRegex = /ppt\/notesMasters\/(?:_rels\/)?notesMaster\d+\.xml(?:\.rels)?$/;
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
            // (The masters, layouts and themes give the slides' places and their text's colours.)
            !!x.match(slideMastersRegex) || !!x.match(slideMasterRelsRegex) || slideLayoutRegex.test(x) || themeRegex.test(x) ||
            (!config.ignoreNotes && notesMasterRegex.test(x)) ||
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
    // The slides' size (`p:sldSz`, in EMU) is the presentation's page: its slides print one to a page of it.
    if (!config.ignorePageGeometry) {
        const presentation = files.find(f => f.path.match(presentationFileRegex))?.content.toString() ?? '';
        const size = /<p:sldSz\b[^>]*>/.exec(presentation)?.[0] ?? '';
        const emu = (name: string) => Number(new RegExp(`\\b${name}="(\\d{1,12})"`).exec(size)?.[1]);
        const width = emu('cx') / 12700, height = emu('cy') / 12700;
        if (width > 0 && height > 0 && width <= 4032 && height <= 4032) metadata.pageSetup = { width: Math.round(width * 100) / 100, height: Math.round(height * 100) / 100 };
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

                                    const color = runColor(rPr);
                                    if (color) formatting.color = color;

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

    // ── Where a slide's shapes are, and what they look like ─────────────────────────────────────────
    // (Unless ignorePageGeometry.) Each shape's paragraphs, pictures, tables and charts get its box on the
    // slide (`bounds`), a shape with a fill, a line, a turn or its text set off its top a `drawing` node
    // (ShapeMetadata) before them, and the slide its background: a deck was laid out as flowing text.

    /** A slide layout's or master's placeholders (by `idx` and by type), background, text styles and part. */
    interface TemplatePart { byIdx: Map<string, Element>; byType: Map<string, Element>; background?: Element; txStyles?: Element; colorMap?: Element; path: string }
    const templateParts = new Map<string, TemplatePart | null>();
    const templatePart = (path: string | undefined): TemplatePart | undefined => {
        if (!path) return undefined;
        if (templateParts.has(path)) return templateParts.get(path) ?? undefined;
        let part: TemplatePart | null = null;
        const file = fileByPath.get(path);
        const doc = file ? readDispensablePart(file, 'the places and look its slides take from it are not known') : undefined;
        if (doc) {
            part = { byIdx: new Map(), byType: new Map(), path, txStyles: getFirstElementByTagName(doc, "p:txStyles") ?? undefined };
            const spTree = getFirstElementByTagName(doc, "p:spTree");
            for (const sp of spTree ? getElementsByTagName(spTree, "p:sp") : []) {
                const ph = getFirstElementByTagName(getDirectChildren(sp, "p:nvSpPr")[0] ?? sp, "p:ph");
                if (!ph) continue;
                const idx = ph.getAttribute("idx"), type = ph.getAttribute("type") || 'body';
                if (idx && !part.byIdx.has(idx)) part.byIdx.set(idx, sp);
                if (!part.byType.has(type)) part.byType.set(type, sp);
            }
            part.background = getFirstElementByTagName(doc, "p:bg") ?? undefined;
            part.colorMap = colorMapOf(doc);
        }
        templateParts.set(path, part);
        return part ?? undefined;
    };
    /** The part a slide's or layout's relationships name as its layout or master. */
    const templateOf = (path: string, kind: 'layout' | 'master' | 'theme' | 'notesMaster'): string | undefined =>
        Object.values(relsByPart.get(path) ?? {}).find(rel => rel.type === kind)?.path;
    /** Each master's theme's colours, read once. */
    const themes = new Map<string, Map<string, string>>();
    const themeOf = (masterPath: string | undefined): Map<string, string> => {
        if (!masterPath) return new Map();
        let colors = themes.get(masterPath);
        if (!colors) {
            const themePath = templateOf(masterPath, 'theme');
            const file = themePath ? fileByPath.get(themePath) : undefined;
            const doc = file ? readDispensablePart(file, 'its colours are not known') : undefined;
            themes.set(masterPath, colors = themeColors(doc));
        }
        return colors;
    };
    /**
     * A master's (`p:clrMap`), or a layout's or slide's overriding (`p:clrMapOvr`), colour map: which of
     * the theme's colours a shape's `bg1`, `tx1`, `bg2` and `tx2` are. A dark master maps the text to the
     * light colour, which read as the dark one drew its text in the background's colour.
     */
    function colorMapOf(doc: Document): Element | undefined {
        const root = doc.documentElement;
        const own = root ? getDirectChildren(root, "p:clrMap")[0] : undefined;
        if (own) return own;
        const override = root ? getDirectChildren(root, "p:clrMapOvr")[0] : undefined;
        return override ? getDirectChildren(override, "a:overrideClrMapping")[0] : undefined;
    }
    /** The theme's colours under the colour maps given (the last given that has one is used). */
    const mappedThemes = new Map<string, Map<string, string>>();
    function mappedTheme(theme: Map<string, string>, themeKey: string, maps: (Element | undefined)[]): Map<string, string> {
        const map = [...maps].reverse().find(Boolean);
        if (!map) return theme;
        const pairs = Array.from(map.attributes).map(a => [a.name, a.value] as const).filter(([name]) => /^[a-zA-Z0-9]+$/.test(name));
        const key = `${themeKey}\u0000${pairs.map(([n, v]) => `${n}=${v}`).join(';')}`;
        let mapped = mappedThemes.get(key);
        if (!mapped) {
            mapped = new Map(theme);
            for (const [name, value] of pairs) { const color = theme.get(value); if (color) mapped.set(name, color); }
            mappedThemes.set(key, mapped);
        }
        return mapped;
    }
    /** A slide's, layout's or master's background colour (its `p:bg`), in the theme's colours. */
    function backgroundOf(bg: Element | undefined, theme: Map<string, string>): string | undefined {
        if (!bg) return undefined;
        const bgPr = getDirectChildren(bg, "p:bgPr")[0];
        const solid = bgPr ? getDirectChildren(bgPr, "a:solidFill")[0] : undefined;
        const gradientStop = bgPr ? getFirstElementByTagName(bgPr, "a:gs") : undefined;
        const ref = getDirectChildren(bg, "p:bgRef")[0];
        const colorHolder = solid ?? gradientStop ?? ref;
        const hex = colorHolder ? rgbOf(Array.from(colorHolder.childNodes).filter(isElement)[0], theme) : undefined;
        return hex ? `#${hex}` : undefined;
    }

    /** The slide being read: its layout, its master and their theme (see placeShape); undefined elsewhere. */
    let placing: { layout?: TemplatePart; master?: TemplatePart; theme: Map<string, string> } | undefined;
    /** The theme's colours, under its colour maps, of the slide or master being read (empty elsewhere). */
    let partTheme = new Map<string, string>();
    /**
     * A run's own colour (`a:solidFill`): a colour of its own, or one of its theme's (`a:schemeClr`, as
     * its colour map maps it), with its shades and tints. A theme's colour was not read.
     */
    function runColor(rPr: Element): string | undefined {
        const fill = getDirectChildren(rPr, "a:solidFill")[0];
        const hex = fill ? rgbOf(Array.from(fill.childNodes).filter(isElement)[0], placing?.theme ?? partTheme) : undefined;
        return hex ? `#${hex}` : undefined;
    }

    /** The placeholder of the layout and of the master a slide's placeholder takes what it does not say from. */
    const placeholderChain = (sp: Element): Element[] => {
        const ph = getFirstElementByTagName(getDirectChildren(sp, "p:nvSpPr")[0] ?? sp, "p:ph");
        if (!ph || !placing) return [];
        const idx = ph.getAttribute("idx"), type = ph.getAttribute("type") || 'body';
        const masterType = type === 'ctrTitle' ? 'title' : type === 'subTitle' || type === 'obj' ? 'body' : type;
        const fromLayout = placing.layout ? ((idx ? placing.layout.byIdx.get(idx) : undefined) ?? placing.layout.byType.get(type)) : undefined;
        const fromMaster = placing.master ? (placing.master.byType.get(masterType) ?? placing.master.byType.get(type)) : undefined;
        return [fromLayout, fromMaster].filter((el): el is Element => !!el);
    };

    /**
     * The size and colour a shape's text is at each level where its runs give none: its own list style's,
     * its placeholder's in the layout and master, the master's text styles for its kind (title, body,
     * other), then the presentation's default. Titles printed at body size, and a master's white text on
     * a dark slide in the reader's default colour.
     */
    /** The presentation's default text style (`p:defaultTextStyle`), read with its slide list. */
    let presentationTextStyle: Element | null | undefined;
    const textDefaults = (sp: Element): ((level: number) => { size?: string; color?: string }) => {
        const ph = getFirstElementByTagName(getDirectChildren(sp, "p:nvSpPr")[0] ?? sp, "p:ph");
        const type = ph ? ph.getAttribute("type") || 'body' : undefined;
        const kind = !ph ? 'p:otherStyle' : type === 'title' || type === 'ctrTitle' ? 'p:titleStyle' : 'p:bodyStyle';
        const chain = [sp, ...placeholderChain(sp)];
        const styles = placing?.master?.txStyles ? getDirectChildren(placing.master.txStyles, kind)[0] : undefined;
        const theme = placing?.theme ?? new Map<string, string>();
        const runDefaults = (list: Element | null | undefined, level: number) => {
            const lvl = list ? getDirectChildren(list, `a:lvl${level + 1}pPr`)[0] : undefined;
            return lvl ? getDirectChildren(lvl, "a:defRPr")[0] : undefined;
        };
        const sizeOf = (defRPr: Element | undefined) => { const sz = Number(defRPr?.getAttribute("sz")); return Number.isFinite(sz) && sz >= 100 && sz <= 400000 ? `${sz / 100}pt` : undefined; };
        const colorOf = (defRPr: Element | undefined) => { const fill = defRPr ? getDirectChildren(defRPr, "a:solidFill")[0] : undefined; const hex = fill ? rgbOf(Array.from(fill.childNodes).filter(isElement)[0], theme) : undefined; return hex ? `#${hex}` : undefined; };
        const lists = (level: number) => [
            ...chain.map(shape => { const txBody = getDirectChildren(shape, "p:txBody")[0]; return runDefaults(txBody ? getDirectChildren(txBody, "a:lstStyle")[0] : undefined, level); }),
            runDefaults(styles, level),
            ...(!ph ? [runDefaults(presentationTextStyle, level)] : []),
        ];
        return level => {
            const found = lists(level);
            return { size: found.map(sizeOf).find(Boolean), color: found.map(colorOf).find(Boolean) };
        };
    };

    /**
     * Places a shape's nodes on the slide: its box (its own, else its placeholder's in the layout or
     * master, through the groups it is in) on each, and before them a `drawing` node where it has a fill,
     * a line or a turn, or sets its text off its top (see ShapeMetadata).
     */
    const placeShape = (sp: Element, nodes: OfficeContentNode[], transform: Transform): OfficeContentNode[] => {
        if (!placing) return nodes;
        const chain = placeholderChain(sp);
        const xfrm = [sp, ...chain].map(el => xfrmOf(getDirectChildren(el, "p:spPr")[0])).find(Boolean);
        if (!xfrm) return nodes;
        const bounds = roundBox(transform(xfrm));
        for (const node of nodes) node.bounds = bounds;
        const look = shapeLook(sp, placing.theme);
        const bodyPrs = [sp, ...chain].map(el => getDirectChildren(getDirectChildren(el, "p:txBody")[0] ?? el, "a:bodyPr")[0]).filter(Boolean);
        const frame = textFrameOf(bodyPrs.find(b => b.getAttribute("anchor")) ?? bodyPrs[0]);
        // Something to draw (a fill, a line), or text that is turned or set off the top of its box.
        const turned = !!(xfrm.rotation || xfrm.flipH || xfrm.flipV);
        if (!look.fill && !look.line && !(nodes.length && (turned || (frame.anchor && frame.anchor !== 'top')))) return nodes;
        const metadata: ShapeMetadata = {
            ...look,
            ...(xfrm.rotation ? { rotation: Math.round(xfrm.rotation * 100) / 100 } : {}),
            ...(xfrm.flipH ? { flipH: true } : {}),
            ...(xfrm.flipV ? { flipV: true } : {}),
            ...(nodes.length && frame.anchor && frame.anchor !== 'top' ? { textAnchor: frame.anchor } : {}),
            ...(nodes.length ? { textInsets: frame.insets } : {}),
        };
        return [{ type: 'drawing', text: '', bounds, metadata }, ...nodes];
    };

    /** Places a picture's, table's or chart's nodes at their frame's box (`p:spPr` or `p:xfrm`). */
    const placeFrame = (holder: Element | undefined, nodes: OfficeContentNode[], transform: Transform): OfficeContentNode[] => {
        if (!placing) return nodes;
        const xfrm = xfrmOf(holder);
        if (xfrm) { const bounds = roundBox(transform(xfrm)); for (const node of nodes) node.bounds = bounds; }
        return nodes;
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
        // The sizes its text takes where its runs give none (on a slide, where its layout is known).
        const defaults = placing ? textDefaults(spNode) : undefined;

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

                                const color = runColor(rPr);
                                if (color) formatting.color = color;

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

                            if (defaults && (!formatting.size || !formatting.color)) {
                                const inherited = defaults(lvl);
                                if (!formatting.size && inherited.size) formatting.size = inherited.size;
                                if (!formatting.color && inherited.color) formatting.color = inherited.color;
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
                            // A slide number: the number of the page it is printed on (TextMetadata.field).
                            if (tag === "a:fld" && element.getAttribute("type") === "slidenum") textNode.metadata = { ...(textNode.metadata as object), field: 'page' } as TextMetadata;

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
    function traverseSpTree(treeNode: Element, rels: PartRelationships | undefined, xmlContentString: string, transform: Transform = IDENTITY): OfficeContentNode[] {
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
                appendAll(nodes, placeShape(element, extractShapeNodes(element, rels, xmlContentString), transform));
            }
            // Case 2: Inline picture
            else if (tag === "p:pic") {
                const imageNode = extractImageNode(element, rels, xmlContentString);
                if (imageNode) {
                    placeFrame(getDirectChildren(element, "p:spPr")[0], [imageNode], transform);
                    nodes.push(imageNode);
                }
            }
            // Case 3: Chart or other graphic frame
            else if (tag === "p:graphicFrame") {
                const diagram = extractDiagramNodes(element, rels);
                if (diagram) {
                    appendAll(nodes, placeFrame(element, diagram, transform));
                } else {
                    const tableNode = extractGraphicFrameNode(element, rels, xmlContentString);
                    if (tableNode) {
                        placeFrame(element, [tableNode], transform);
                        nodes.push(tableNode);
                    }
                }
            }
            // Case 4: Grouped shape (recursive!)
            else if (tag === "p:grpSp") {
                // Recurse into the group element itself which holds the child shapes, placed by its transform
                appendAll(nodes, traverseSpTree(element, rels, xmlContentString, placing ? groupTransform(getDirectChildren(element, "p:grpSpPr")[0], transform) : transform));
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
            else if (typeAttr.endsWith("relationships/slideLayout")) simplifiedType = "layout";
            else if (typeAttr.endsWith("relationships/slideMaster")) simplifiedType = "master";
            else if (typeAttr.endsWith("relationships/theme")) simplifiedType = "theme";
            else if (typeAttr.endsWith("relationships/notesMaster")) simplifiedType = "notesMaster";
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
        presentationTextStyle = presentationXml ? getFirstElementByTagName(presentationXml, "p:defaultTextStyle") : null;
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
    const readPart = (file: PackageFile, node: OfficeContentNode, withXml?: (xml: Document) => void): OfficeContentNode => {
        const xmlContentString = file.content.toString();
        const xml = parseXmlString(xmlContentString, { config, locator: config.includeRawContent });
        withXml?.(xml);
        if (config.includeRawContent) {
            node.rawContent = getRawContent(xml, xmlContentString, config);
        }
        const spTree = getFirstElementByTagName(xml, "p:spTree");
        if (spTree) appendAll(node.children!, traverseSpTree(spTree, relsByPart.get(file.path), xmlContentString));
        return node;
    };

    const slideMasters: OfficeContentNode[] = [];
    for (const file of config.ignoreSlideMasters ? [] : files) {
        if (!file.path.match(slideMastersRegex)) continue;
        checkAbortSignal(config.abortSignal);
        const masterMatch = file.path.match(/slideMaster(\d+)\.xml/);
        partTheme = mappedTheme(themeOf(file.path), file.path, [templatePart(file.path)?.colorMap]);
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
        {
            const layout = templatePart(templateOf(file.path, 'layout'));
            const master = layout ? templatePart(templateOf(layout.path, 'master')) : undefined;
            const masterPath = master?.path;
            partTheme = mappedTheme(themeOf(masterPath), masterPath ?? '', [master?.colorMap, layout?.colorMap]);
            if (!config.ignorePageGeometry) placing = { layout, master, theme: partTheme };
        }
        let slideNode: OfficeContentNode;
        try {
            const metadata: SlideMetadata = { slideNumber };
            slideNode = readPart(file, { type: 'slide', children: [], metadata }, xml => {
                // A slide's own colour map, over its layout's and master's.
                const slideMap = colorMapOf(xml);
                if (slideMap) {
                    const layout = templatePart(templateOf(file.path, 'layout'));
                    const master = layout ? templatePart(templateOf(layout.path, 'master')) : undefined;
                    partTheme = mappedTheme(themeOf(master?.path), master?.path ?? '', [master?.colorMap, layout?.colorMap, slideMap]);
                    if (placing) placing.theme = partTheme;
                }
                if (!placing) return;
                const background = backgroundOf(getFirstElementByTagName(xml, "p:bg") ?? undefined, placing.theme)
                    ?? backgroundOf(placing.layout?.background, placing.theme) ?? backgroundOf(placing.master?.background, placing.theme);
                if (background) metadata.background = background;
            });
        } finally {
            placing = undefined;
        }
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
            // In its notes master's colours, not the slide's: a dark slide master's light text is no
            // colour for notes on a white page.
            const notesMaster = templatePart(templateOf(notesFile.path, 'notesMaster'));
            const notesTheme = (own?: Element) => mappedTheme(themeOf(notesMaster?.path), notesMaster?.path ?? '', [notesMaster?.colorMap, own]);
            partTheme = notesTheme();
            const note = readPart(notesFile, { type: 'note', children: [], metadata: { slideNumber, noteId: `slide-note-${slideNumber}` } }, xml => {
                const own = colorMapOf(xml);
                if (own) partTheme = notesTheme(own);
            });
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

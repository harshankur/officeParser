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
import { BreakMetadata, CellMetadata, CodeMetadata, CommentMetadata, FullOfficeParserConfig, ImageMetadata, IndentationMetadata, ListMetadata, OfficeAttachment, OfficeContentNode, OfficeErrorType, OfficeIssue, OfficeParserAST, OfficeWarningType, TableMetadata, TextFormatting, TextMetadata } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { checkAbortSignal, getOfficeError, logWarning } from '../utils/errorUtils.js';
import { createAttachment, renameAttachments } from '../utils/imageUtils.js';
import { isEmptyMath, ommlToLatex } from '../utils/mathUtils.js';
import { ocrDuringParse } from '../utils/ocrUtils.js';
import { getAllElementsByTagName, getChildElements, getDirectChildren, getElementsByTagName, getOutermostElements, getFirstElementByTagName, getRawContent, isElement, parseOfficeMetadata, parseOOXMLAppProperties, parseOOXMLCustomProperties, parseXmlString, takeNodes, takeXmlElements, countByte } from '../utils/xmlUtils.js';
import { extractFiles, findRequiredPart, maxUncompressedBytesOf } from '../utils/zipUtils.js';
import { lookupTable, plainRecord, setOwn } from '../utils/lookupUtils.js';
import { cellSpan, MAX_COL_SPAN } from '../utils/numberUtils.js';
import { appendAll } from '../utils/nodeListUtils.js';
import { diagramList, readDiagram } from '../utils/diagramUtils.js';
import { UniqueNames } from '../utils/uniqueNames.js';
import { decodeMarkup, mhtPartText, MhtPart, readMht } from '../utils/mhtUtils.js';
import { alternateContentBranch, isAlternateContent, WORD_NAMESPACES } from '../utils/markupCompatibility.js';
import { parseHtml } from './HtmlParser.js';
import { parseRtf } from './RtfParser.js';

/**
 * A DOCX read as another DOCX's alternative-format chunk (see readChunk): the bytes it may inflate, what
 * the enclosing document left of `decompressionLimits.maxUncompressedBytes`, and the bytes it did.
 */
interface WordChunk {
    bytesLeft: number;
    bytesUsed: number;
}

/** SmartArt data parts (see readDiagram). */
const diagramDataRegex = /^word\/diagrams\/data\d*\.xml$/;

/** The kinds of part an alternative-format chunk (`w:altChunk`) is read from, by extension: HTML, MHT, RTF, plain text or a DOCX. */
const altChunkExtensionRegex = /\.(?:html?|xht(?:ml)?|mht(?:ml)?|rtf|txt|docx)$/i;

/** The relationship parts of the document's parts (the document, notes, comments, headers and footers). */
const partRelsRegex = /^word\/_rels\/([^/]+)\.rels$/;

/**
 * What a chunk's reader may throw that ends the whole parse rather than the chunk: the budgets the
 * document is held to as a whole, which a chunk's content counts against as the document's own does
 * (skipping the chunk that crossed one would read the rest past it). Anything else a chunk's reader
 * throws (not a ZIP, no document part, nested too deep to read) is that chunk's: it is not read, and
 * the document is.
 */
const DOCUMENT_BUDGETS: ReadonlySet<string> = new Set([
    OfficeErrorType.XML_ELEMENT_LIMIT_EXCEEDED,
    OfficeErrorType.ZIP_SIZE_LIMIT_EXCEEDED,
    OfficeErrorType.ZIP_ENTRY_COUNT_LIMIT_EXCEEDED,
    OfficeErrorType.OPERATION_ABORTED,
    OfficeErrorType.OCR_TERMINATED,
]);

/** The instruction of a complex field, kept up to this length: a NOTEREF's is a few words. */
const MAX_FIELD_INSTRUCTION = 256;

/** Maximum nesting depth for paragraph children (hyperlinks, fields, fallback wrappers) before DoS rejection. */
const MAX_WORD_CHILD_DEPTH = 512;

/**
 * A NOTEREF field (a note's number, shown again where its bookmark names): its bookmark, then its
 * switches. Formatted as the note's reference mark (`\f`), it is how Word writes a note referred to again.
 */
const noteRefFieldRegex = /^\s*NOTEREF\s+"?([^\s"\\]+)"?([^]*)$/i;

/** A relationship's target as a package path, resolved against the folder of the part naming it. */
const resolvePartPath = (folder: string, target: string): string => {
    const segments: string[] = [];
    for (const segment of (target.startsWith('/') ? target.slice(1) : folder + target).split('/')) {
        if (segment === '..') segments.pop();
        else if (segment && segment !== '.') segments.push(segment);
    }
    return segments.join('/');
};

/** A text box's content, which the text box's own blocks are parsed from (see ownFirst). */
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
export const parseWord = async (buffer: Buffer, config: FullOfficeParserConfig, chunk?: WordChunk): Promise<OfficeParserAST> => {
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
     * element (`w:customXml`) stands for what it wraps, at any depth, and alternate content for the branch
     * this reader reads. A table of contents, a cover page or a form is a content control around
     * paragraphs, rows or cells, and reading direct children alone dropped all of it.
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
            } else if (isAlternateContent(node)) {
                const branch = alternateContentBranch(node, WORD_NAMESPACES);
                if (branch) pushChildren(branch);
            } else {
                out.push(node);
            }
        }
        return out;
    };

    /**
     * The content of mc:AlternateContent this reader reads (see alternateContentBranch): the first Choice
     * whose namespaces it understands, else the Fallback.
     */
    const resolveAlternateContent = (element: Element): Node[] => {
        const branch = alternateContentBranch(element, WORD_NAMESPACES);
        return branch ? Array.from(branch.childNodes) : [];
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
            partRelsRegex.test(x) ||
            !!x.match(stylesFileRegex) ||
            (!config.ignoreComments && !!x.match(commentsFileRegex)) ||
            (!config.ignoreHeadersAndFooters && (!!x.match(headerFileRegex) || !!x.match(footerFileRegex))) ||
            (!!config.extractAttachments && !!x.match(mediaFileRegex)) ||
            diagramDataRegex.test(x),
        // A DOCX read as a chunk inflates within what the enclosing document left.
        chunk ? { ...config.decompressionLimits, maxUncompressedBytes: chunk.bytesLeft } : config.decompressionLimits,
        config
    );
    let inflatedBytes = files.reduce((total, f) => total + f.content.length, 0);
    if (chunk) chunk.bytesUsed = inflatedBytes;

    // A DOCX without its main document part is not a DOCX. Checked with the same regex the
    // parse loop below uses to recognize it, so the two cannot fall out of step.
    findRequiredPart(files, path => !!path.match(documentFileRegex), config,
        { fileType: 'docx', part: 'word/document.xml' });

    // Each part's relationships, by the part's path: an id in a note, comment, header or footer names a
    // relationship of that part's own, and looking it up among the document's found another target, or
    // none. Null-prototype, as every map keyed by the document's own ids and names is here (see
    // numberingMap). The targets of alternative-format chunk relationships are gathered on the way.
    const relsByPart = new Map<string, { [key: string]: string }>();
    const chunkTargets = new Set<string>();
    for (const f of files) {
        const owner = partRelsRegex.exec(f.path);
        if (!owner) continue;
        const rels: { [key: string]: string } = Object.create(null);
        for (const relationship of getElementsByTagName(parseXmlString(f.content.toString(), { config }), "Relationship")) {
            const id = relationship.getAttribute("Id");
            const target = relationship.getAttribute("Target");
            if (!id || !target) continue;
            rels[id] = target;
            if ((relationship.getAttribute("Type") || '').endsWith('/aFChunk') && relationship.getAttribute("TargetMode") !== 'External') {
                chunkTargets.add(resolvePartPath('word/', target));
            }
        }
        relsByPart.set(`word/${owner[1]}`, rels);
    }
    const noRels: { [key: string]: string } = Object.create(null);
    const relsOf = (partPath: string): { [key: string]: string } => relsByPart.get(partPath) ?? noRels;
    // The relationships of the part being read (see withPartRels): a paragraph's pictures, links and
    // SmartArt resolve against them.
    let partRels = noRels;
    const withPartRels = <T>(partPath: string, read: () => T): T => {
        const previous = partRels;
        partRels = relsOf(partPath);
        try {
            return read();
        } finally {
            partRels = previous;
        }
    };

    // The parts alternative-format chunks are read from, inflated second and only those an aFChunk
    // relationship names: every part of a chunk's kind was inflated, an embedded object
    // (`word/embeddings/*.docx`) too, against the document's limit.
    const byteLimit = chunk ? chunk.bytesLeft : maxUncompressedBytesOf(config.decompressionLimits);
    const chunkFiles = chunkTargets.size === 0 ? [] : await extractFiles(
        buffer,
        x => chunkTargets.has(x) && altChunkExtensionRegex.test(x),
        { ...config.decompressionLimits, maxUncompressedBytes: Math.max(0, byteLimit - inflatedBytes) },
        config
    );
    for (const f of chunkFiles) inflatedBytes += f.content.length;
    if (chunk) chunk.bytesUsed = inflatedBytes;

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

    // A note referred to again: Word writes a later reference to a note as a NOTEREF field formatted as
    // its mark, naming a bookmark around the first reference. Each bookmark open where a note reference
    // is read (by its id) names that note, and a NOTEREF naming it is read as a reference to the note.
    // An open bookmark names the first note reference in it, and is then no longer looked at.
    const openBookmarks = new Map<string, string>();
    const noteByBookmark = new Map<string, OfficeContentNode>();
    // The bookmarks a NOTEREF was read from, and the anchors internal links go to: a bookmark that marked
    // a note's first reference for a NOTEREF, and nothing a link goes to, is not the document's anchor.
    const noteRefBookmarks = new Set<string>();
    const linkedAnchors = new Set<string>();
    const openBookmark = (bookmark: Element): void => {
        const id = bookmark.getAttribute("w:id");
        const name = bookmark.getAttribute("w:name");
        if (id !== null && name) openBookmarks.set(id, name);
    };
    const closeBookmark = (bookmark: Element): void => {
        const id = bookmark.getAttribute("w:id");
        if (id !== null) openBookmarks.delete(id);
    };
    const noteReferenced = (note: OfficeContentNode): void => {
        for (const name of openBookmarks.values()) if (!noteByBookmark.has(name)) noteByBookmark.set(name, note);
        openBookmarks.clear();
    };

    // Alternative-format chunks (`w:altChunk`): content in another format (HTML, MHT, RTF, plain text or
    // a DOCX) that Word merges into the document where the chunk stands; html-docx-js writes a document's
    // whole body that way, and it was dropped. Each chunk's content by its element, read before the part
    // holding it (whose tables are read synchronously); each part once, at the first chunk naming it.
    const chunkParts = new Map<string, (typeof files)[number]>();
    for (const f of chunkFiles) if (!chunkParts.has(f.path)) chunkParts.set(f.path, f);
    const chunkNodes = new Map<Element, OfficeContentNode[]>();
    const placedChunks = new Set<Element>();
    // Once per kind of chunk not read: a document of a million chunks naming nothing is one warning.
    const unreadChunkKinds = new Set<string>();
    const chunkNotRead = (kind: string, reason: string): void => {
        if (unreadChunkKinds.has(kind)) return;
        unreadChunkKinds.add(kind);
        logWarning(OfficeWarningType.ALT_CHUNK_NOT_READ, config, reason);
    };
    const readChunkParts = new Set<string>();
    // A chunk's pictures take names none of the document's own media, or other chunks' pictures, take.
    const attachmentNames = new UniqueNames();
    for (const media of mediaFiles) attachmentNames.add(media.path.split('/').pop() || '');
    const claimAttachmentName = (name: string): string => {
        const dot = name.lastIndexOf('.');
        const stem = dot > 0 ? name.slice(0, dot) : name;
        const extension = dot > 0 ? name.slice(dot) : '';
        return attachmentNames.claim(name, n => `${stem}-${n}${extension}`);
    };
    // What a DOCX chunk may inflate: what this document left of the limit, shared by all its chunks.
    let chunkBytesLeft = byteLimit - inflatedBytes;
    const adoptChunk = (ast: OfficeParserAST): OfficeContentNode[] => {
        const renamed = new Map<string, string>();
        for (const attachment of ast.attachments) {
            const name = claimAttachmentName(attachment.name);
            if (name !== attachment.name) renamed.set(attachment.name, name);
            attachments.push({ ...attachment, name });
        }
        renameAttachments(ast.content, renamed);
        return ast.content;
    };
    /** A chunk part's content, read by its kind (`extension`). */
    const readChunkPart = async (content: Buffer, path: string, extension: string): Promise<OfficeContentNode[]> => {
        if (extension === 'docx') {
            const nested: WordChunk = { bytesLeft: Math.max(0, chunkBytesLeft), bytesUsed: 0 };
            try {
                return adoptChunk(await parseWord(content, config, nested));
            } finally {
                chunkBytesLeft -= nested.bytesUsed;
            }
        }
        // A chunk's content counts against the document's element budget before it is read, as its XML
        // does (see takeNodes): an RTF chunk by its groups and control words, plain text by its lines.
        if (extension === 'rtf') {
            takeNodes(countByte(content, 0x5C) + countByte(content, 0x7B), config);
            return adoptChunk(await parseRtf(content, config));
        }
        if (extension === 'txt') {
            takeNodes(countByte(content, 0x0A) + countByte(content, 0x0D) + 1, config);
            const lines = decodeMarkup(content, false).split(/\r\n?|\n/);
            return lines.filter(line => line.trim()).map(line => ({ type: 'paragraph', text: line, children: [{ type: 'text', text: line }] }));
        }
        let html: string;
        let imageAttachment: ((src: string) => string | undefined) | undefined;
        if (extension === 'mht' || extension === 'mhtml') {
            const messageParts = readMht(content, count => takeNodes(count, config));
            const page = messageParts.find(p => p.contentType === 'text/html') ?? messageParts[0];
            html = page ? mhtPartText(page) : '';
            // A picture the page shows from the message: found by its Content-Location, else its file
            // name; an attachment once however many pictures show it.
            const byLocation = new Map<string, MhtPart>();
            const byName = new Map<string, MhtPart>();
            for (const p of messageParts) {
                if (p === page || !p.location || !p.contentType.startsWith('image/')) continue;
                if (!byLocation.has(p.location)) byLocation.set(p.location, p);
                const name = p.location.slice(p.location.lastIndexOf('/') + 1);
                if (name && !byName.has(name)) byName.set(name, p);
            }
            const shown = new Map<MhtPart, string>();
            imageAttachment = src => {
                if (!config.extractAttachments) return undefined;
                const found = byLocation.get(src) ?? byName.get(src.slice(src.lastIndexOf('/') + 1));
                if (!found) return undefined;
                let name = shown.get(found);
                if (name === undefined) {
                    const attachment = createAttachment(claimAttachmentName(found.location!.slice(found.location!.lastIndexOf('/') + 1) || 'image'), found.body);
                    attachments.push(attachment);
                    shown.set(found, name = attachment.name);
                }
                return name;
            };
        } else {
            // In the encoding its byte order mark or its <meta> names (see decodeMarkup).
            html = decodeMarkup(content, true);
        }
        // Its elements count against the document's budget, as an EPUB chapter's do (see maxXmlElements).
        takeXmlElements(html, config);
        return adoptChunk(await parseHtml(Buffer.from(html, 'utf8'), config, { imageAttachment }));
    };
    /**
     * The content of the chunk `element` stands for, its relationship looked up in `rels` (those of the
     * part holding it). A chunk that cannot be read (not a ZIP, no document part in it, nested too deep)
     * is not read, with a warning, and the document is: one such chunk failed the whole document. What
     * ends the parse still does (see DOCUMENT_BUDGETS).
     */
    const readChunk = async (element: Element, rels: { [key: string]: string }): Promise<OfficeContentNode[]> => {
        const rId = element.getAttribute("r:id");
        const target = rId ? rels[rId] : undefined;
        if (!target) {
            chunkNotRead('relationship', `its relationship (${rId ?? 'none'}) names no part`);
            return [];
        }
        const path = resolvePartPath('word/', target);
        if (readChunkParts.has(path)) return [];
        readChunkParts.add(path);
        const part = chunkParts.get(path);
        if (!part) {
            chunkNotRead('part', `its part (${path}) is missing, or not HTML, MHT, RTF, plain text or DOCX`);
            return [];
        }
        const extension = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
        // One level: a DOCX chunk's own DOCX chunks are not read.
        if (extension === 'docx' && chunk) {
            chunkNotRead('nested', `it is a DOCX (${path}) inside a DOCX chunk`);
            return [];
        }
        // An error the chunk's reader reports as it throws is held back until it is known to end the
        // parse: a chunk that is not a DOCX reported the document as corrupt, and was then not read.
        const reported: OfficeIssue[] = [];
        const onWarning = config.onWarning;
        config.onWarning = issue => { if (issue.type === 'error') reported.push(issue); else onWarning?.(issue); };
        try {
            return await readChunkPart(part.content, path, extension);
        } catch (error: any) {
            if (error?.name === 'AbortError' || DOCUMENT_BUDGETS.has(error?.officeIssue?.code)) {
                for (const issue of reported) onWarning?.(issue);
                throw error;
            }
            const detail = String(error?.officeIssue?.code ?? error?.message ?? error).slice(0, 200);
            chunkNotRead('unreadable', `its part (${path}) could not be read (${detail})`);
            return [];
        } finally {
            config.onWarning = onWarning;
        }
    };
    /** Reads, before `root`'s part is parsed, the chunks in it, their relationships among `rels`. */
    const readChunksIn = async (root: Element, rels: { [key: string]: string }): Promise<void> => {
        for (const element of getAllElementsByTagName(root, "w:altChunk")) chunkNodes.set(element, await readChunk(element, rels));
    };
    // SmartArt data parts by path, and those read so far: each is shown once, at the first drawing
    // showing it.
    const diagramParts = new Map<string, (typeof files)[number]>();
    for (const f of files) if (diagramDataRegex.test(f.path) && !diagramParts.has(f.path)) diagramParts.set(f.path, f);
    const readDiagrams = new Set<string>();
    let diagramCount = 0;
    /**
     * The SmartArt a drawing shows, as a list (see readDiagram, diagramList), or none: not SmartArt, or
     * shown before.
     */
    const diagramNodes = (drawing: Element): OfficeContentNode[] => {
        const relIds = ownFirst(drawing, "dgm:relIds");
        const rId = relIds?.getAttribute("r:dm");
        const target = rId ? partRels[rId] : undefined;
        if (!target) return [];
        const path = resolvePartPath('word/', target);
        const part = readDiagrams.has(path) ? undefined : diagramParts.get(path);
        if (!part) return [];
        readDiagrams.add(path);
        return diagramList(readDiagram(part.content.toString(), config), `smartart-${++diagramCount}`);
    };

    /** A chunk's content where it stands, read by readChunk before the part holding it. */
    const placeChunk = (element: Element): OfficeContentNode[] => {
        placedChunks.add(element);
        return chunkNodes.get(element) ?? [];
    };

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

    /**
     * The note a formatted NOTEREF field's instruction names (see noteRefFieldRegex): `note` is the note
     * whose first reference its bookmark encloses, if one was read before it. Undefined for any other
     * field, a NOTEREF showing the number as text (no `\f`) included.
     */
    const noteRefField = (instruction: string): { note?: OfficeContentNode } | undefined => {
        const match = noteRefFieldRegex.exec(instruction.slice(0, MAX_FIELD_INSTRUCTION));
        if (!match || !/\\f(?![a-z])/i.test(match[2])) return undefined;
        const note = noteByBookmark.get(match[1]);
        if (note) noteRefBookmarks.add(match[1]);
        return { note };
    };

    // Helper to parse a paragraph node
    /**
     * @param lifted - The block list the paragraph is read into (see readBlocks), its place in it taken:
     * what the paragraph draws that a paragraph cannot hold (a text box's paragraphs, lists and tables, a
     * SmartArt list) is read into it, after the paragraph. Read into the paragraph's own text, a text
     * box's paragraphs ran together ("Intro text.Box titleBox line two"), and its lists and tables were
     * flattened.
     */
    const parseParagraph = (pNode: Element, documentContent: string, pendingAnchorIds: string[] = [], lifted: OfficeContentNode[] = []): OfficeContentNode => {
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
        const children: OfficeContentNode[] = [];
        const notes: OfficeContentNode[] = [];
        const comments: OfficeContentNode[] = [];
        const anchorIds: string[] = [...pendingAnchorIds];

        /** A note or comment referred to where the paragraph stands: on the node before it, else the paragraph. */
        const attach = (node: OfficeContentNode, list: 'notes' | 'comments'): void => {
            if (children.length > 0) {
                const target = children[children.length - 1];
                if (!target[list]) target[list] = [];
                target[list]!.push(node);
            } else {
                (list === 'notes' ? notes : comments).push(node);
            }
        };

        // Complex fields (w:fldChar) open in the paragraph, innermost last. A formatted NOTEREF's result
        // (the note's number) is read as a reference to its note, not as text.
        const fields: { instruction: string; resolved: boolean; hidesResult: boolean }[] = [];
        let hiddenResults = 0;
        const resolveField = (field: { instruction: string; resolved: boolean; hidesResult: boolean }): void => {
            if (field.resolved) return;
            field.resolved = true;
            const reference = noteRefField(field.instruction);
            if (!reference) return;
            if (reference.note) attach(reference.note, 'notes');
            // Unresolved (its note not read before it), its number stays as text; with notes ignored, it goes as their marks do.
            if (reference.note || config.ignoreNotes) {
                field.hidesResult = true;
                hiddenResults++;
            }
        };

        // Traverse children of paragraph (runs, hyperlinks, etc.)
        // Whether a hyperlink enclosing the node being read gives its runs their link (see w:hyperlink).
        let insideLink = false;
        const processChildNode = (node: Node, depth = 0) => {
            if (depth > MAX_WORD_CHILD_DEPTH) throw getOfficeError(OfficeErrorType.MAX_NESTING_DEPTH_EXCEEDED, config);
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

                // The run's content in order, alternate content as the branch this reader reads: an emoji
                // Word writes as `w16se:symEx`, its character in the Fallback's `w:t`, was dropped.
                const runItems: Node[] = [];
                const pushRunItems = (nodes: ArrayLike<Node>) => { for (let i = nodes.length - 1; i >= 0; i--) runItems.push(nodes[i]); };
                pushRunItems(runNode.childNodes);
                while (runItems.length) {
                    const child = runItems.pop()!;
                    if (!isElement(child)) continue;
                    if (isAlternateContent(child)) {
                        pushRunItems(resolveAlternateContent(child));
                        continue;
                    }

                    // also handle unprefixed version (mirroring the behaviour of getElementsByTagName)

                    // Text content (none of a field result read as a note reference)
                    if (child.tagName === "w:t" || child.tagName === "t" || child.tagName === "m:t") {
                        if (hiddenResults > 0) continue;
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
                    // Complex fields: where one begins, its instruction, where its result begins, where it ends.
                    else if (child.tagName === "w:fldChar") {
                        const kind = child.getAttribute("w:fldCharType");
                        if (kind === 'begin') fields.push({ instruction: '', resolved: false, hidesResult: false });
                        else if (kind === 'separate' && fields.length) resolveField(fields[fields.length - 1]);
                        else if (kind === 'end' && fields.length) {
                            const field = fields[fields.length - 1];
                            resolveField(field);
                            fields.pop();
                            if (field.hidesResult) hiddenResults--;
                        }
                    } else if (child.tagName === "w:instrText") {
                        const field = fields[fields.length - 1];
                        if (field && !field.resolved && field.instruction.length < MAX_FIELD_INSTRUCTION) {
                            field.instruction += (child.textContent || '').slice(0, MAX_FIELD_INSTRUCTION - field.instruction.length);
                        }
                    }
                    // A tab is a character of the text, whatever column its tab stop sets: left out, "A<tab>B"
                    // read "AB".
                    else if (child.tagName === "w:tab" || child.tagName === "tab") {
                        if (hiddenResults > 0) continue;
                        text += '\t';
                        children.push({ type: 'text', text: '\t', formatting } as OfficeContentNode);
                    }
                    // Break nodes
                    else if (child.tagName === "w:br"
                            || child.tagName === "br"
                            || child.tagName === "w:cr"
                            || child.tagName === "cr"
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
                        // A line break the author typed is content, kept whatever includeBreakNodes says, as
                        // HTML's <br> is: left out, "Line one<br/>Line two" read "Line oneLine two". A page or
                        // column break is layout, kept with includeBreakNodes.
                        const lineBreak = breakType === 'textWrapping' || breakType === 'carriageReturn';
                        if (!lineBreak && !config.includeBreakNodes) continue;
                        if (lineBreak && hiddenResults > 0) continue;
                        if (lineBreak) text += '\n';

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
                    else if (isAlternateContent(item)) pushNodes(resolveAlternateContent(item));
                    else pushNodes(item.childNodes);
                }

                // Pictures: a drawing showing one (a DrawingML blip, or a VML image). A text box, SmartArt,
                // chart or plain shape shows none, and gave an image node of nothing (`![image]()`).
                if (config.extractAttachments) {
                    for (const imgNode of runDrawings) {
                        const blip = ownFirst(imgNode, "a:blip");
                        const imagedata = blip ? undefined : ownFirst(imgNode, "v:imagedata");
                        if (!blip && !imagedata) continue;
                        // Extract Alt Text
                        let altText = '';
                        const docPr = ownFirst(imgNode, "wp:docPr");
                        if (docPr) {
                            altText = docPr.getAttribute("descr") || docPr.getAttribute("title") || '';
                        }
                        // A picture that is a link itself (Insert > Link on a picture writes it on docPr).
                        const pictureLinkRid = docPr ? getFirstElementByTagName(docPr, "a:hlinkClick")?.getAttribute("r:id") : null;
                        const pictureLink = pictureLinkRid && partRels[pictureLinkRid] ? { link: partRels[pictureLinkRid], linkType: 'external' as const } : {};

                        // Extract Relationship ID
                        const rId = (blip ? blip.getAttribute("r:embed") : imagedata!.getAttribute("r:id")) || '';

                        if (rId && partRels[rId]) {
                            const target = partRels[rId];
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
                            // A picture whose data the package does not hold: its alt text, where it has one.
                            const imageNode: OfficeContentNode = {
                                type: 'image',
                                text: '',
                                ...(altText ? { metadata: { altText } as ImageMetadata } : {}),
                            };
                            if (config.includeRawContent) {
                                imageNode.rawContent = getRawContent(imgNode, documentContent, config);
                            }
                            children.push(imageNode);
                        }
                    }
                }

                // What the run's drawings hold that a paragraph cannot: SmartArt (a list, as a slide's
                // is), and text boxes' blocks, read after the paragraph.
                for (const drawing of runDrawings) {
                    appendAll(lifted, diagramNodes(drawing));
                    for (const txbx of getOutermostElements(drawing, "w:txbxContent")) readBlocks(txbx, documentContent, lifted);
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
                            attach(noteNode, 'notes');
                            noteReferenced(noteNode);
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
                            attach(noteNode, 'notes');
                            noteReferenced(noteNode);
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
                            attach(commentNode, 'comments');
                        }
                    }
                }
            } else if (isElement(node) && node.nodeName === 'w:hyperlink') {
                const hlNode = node;
                const rId = hlNode.getAttribute("r:id");
                const anchor = hlNode.getAttribute("w:anchor");
                if (anchor) linkedAnchors.add(anchor);

                let linkMetadata: TextMetadata | undefined;
                if (anchor && !config.ignoreInternalLinks) {
                    linkMetadata = { link: '#' + anchor, linkType: 'internal' };
                } else if (rId && partRels[rId]) {
                    linkMetadata = { link: partRels[rId], linkType: 'external' };
                }

                // Process children of hyperlink (usually runs). The outermost hyperlink with a link gives
                // its runs (and a picture in it) that link, once, after they are read: every level of
                // nested hyperlinks gave it again to all the runs inside it, so 1,000 levels around
                // 400,000 runs (23 KB) took 17 seconds.
                const applies = !!linkMetadata && !insideLink;
                const wasInside = insideLink;
                if (applies) insideLink = true;
                const startIndex = children.length;
                for (const child of Array.from(hlNode.childNodes)) processChildNode(child, depth + 1);
                insideLink = wasInside;
                if (applies) {
                    for (let i = startIndex; i < children.length; i++) {
                        if (children[i].type === 'text' || children[i].type === 'image') {
                            children[i].metadata = { ...(children[i].metadata ?? {}), ...linkMetadata } as any;
                        }
                    }
                }
            } else if (isElement(node) && node.nodeName === 'w:bookmarkStart') {
                openBookmark(node);
                const bookmarkName = node.getAttribute("w:name");
                if (bookmarkName && !bookmarkName.startsWith('_GoBack') && !config.ignoreInternalLinks) {
                    anchorIds.push(bookmarkName);
                }
            } else if (isElement(node) && node.nodeName === 'w:bookmarkEnd') {
                closeBookmark(node);
            } else if (isElement(node) && node.nodeName === 'w:fldSimple') {
                // A simple field: a formatted NOTEREF is a reference to its note (its result, the number,
                // not text); any other field is its result.
                const reference = noteRefField(node.getAttribute("w:instr") || '');
                if (reference?.note) attach(reference.note, 'notes');
                if (!reference || (!reference.note && !config.ignoreNotes)) {
                    for (const child of Array.from(node.childNodes)) processChildNode(child, depth + 1);
                }
            } else if (isElement(node) && isAlternateContent(node)) {
                const resolved = resolveAlternateContent(node);
                for (const rNode of resolved) processChildNode(rNode, depth + 1);
            } else if (isElement(node) && (node.nodeName === 'w:pict' || node.nodeName === 'pict' || node.nodeName === 'w:drawing' || node.nodeName === 'drawing')) {
                // A legacy shape or a drawing standing in the paragraph itself: its SmartArt and its own
                // text boxes' blocks, read after the paragraph. The picture's own text boxes, not those in
                // their paragraphs: reading a text box's paragraphs reaches those, and taking them here too
                // read nested text boxes again at each level (1.3 KB took half a minute).
                appendAll(lifted, diagramNodes(node));
                for (const txbx of getOutermostElements(node, "w:txbxContent")) readBlocks(txbx, documentContent, lifted);
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
            } else if (isElement(node) && (node.nodeName === 'w:moveFrom' || node.nodeName === 'w:del')) {
                // Content a tracked change removed, or moved to where a `w:moveTo` holds it: not the
                // document's text (read, moved text came out twice). Insertions are read as text.
            } else if (node.childNodes.length > 0) {
                // Generic fallback for unknown elements that might contain content
                for (const child of Array.from(node.childNodes)) processChildNode(child, depth + 1);
            }
        };

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

            if (config.includeRawContent) listNode.rawContent = getRawContent(pNode, documentContent, config);
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
            if (config.includeRawContent) headingNode.rawContent = getRawContent(pNode, documentContent, config);
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
            if (config.includeRawContent) paraNode.rawContent = getRawContent(pNode, documentContent, config);
            return paraNode;
        }
    };

    /**
     * Reads the blocks of `container` (a body, cell, note, comment, header, footer or text box) into
     * `out`: its paragraphs, tables and chunks, a bookmark standing between blocks given to the block
     * after it. A paragraph's place in `out` is taken before it is read, so what it draws that a
     * paragraph cannot hold follows it in the same list (see parseParagraph): each text box is read once
     * however deep text boxes nest, where copying each level's content up into the one around it copied
     * it again per level.
     */
    const readBlocks = (container: Element, xml: string, out: OfficeContentNode[]): void => {
        let pendingAnchorIds: string[] = [];
        for (const child of layoutChildren(container)) {
            checkAbortSignal(config.abortSignal);
            if (child.nodeName === 'w:p') {
                const at = out.length;
                out.push(undefined as unknown as OfficeContentNode);
                out[at] = parseParagraph(child, xml, pendingAnchorIds, out);
                pendingAnchorIds = [];
            } else if (child.nodeName === 'w:tbl') {
                out.push(parseTable(child, xml, pendingAnchorIds));
                pendingAnchorIds = [];
            } else if (child.nodeName === 'w:bookmarkStart') {
                openBookmark(child);
                const bookmarkName = child.getAttribute("w:name");
                if (bookmarkName && !bookmarkName.startsWith('_GoBack') && !config.ignoreInternalLinks) {
                    pendingAnchorIds.push(bookmarkName);
                }
            } else if (child.nodeName === 'w:bookmarkEnd') {
                closeBookmark(child);
            } else if (child.nodeName === 'w:altChunk') {
                appendAll(out, placeChunk(child));
            }
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

                // Cells contain paragraphs (and other block-level elements): tables, chunks, and what
                // their paragraphs draw (text boxes, SmartArt).
                const cellChildren: OfficeContentNode[] = [];
                readBlocks(tcNode, documentContent, cellChildren);
                let cellText = '';
                for (const child of cellChildren) if (child.type !== 'table') cellText += child.text ?? '';

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

        // A bookmark standing before the table (a link's target, a caption's label) is the table's.
        return {
            type: 'table',
            children: rows,
            ...(pendingAnchorIds.length > 0 ? { metadata: { anchorIds: [...pendingAnchorIds] } as TableMetadata } : {}),
        };
    };

    /**
     * Reads a part (a note's, comment's, header's or footer's, or the document's own): its chunks first,
     * then `read`, synchronously, with the part's own relationships. A bookmark does not reach past its part.
     */
    const readPart = async (path: string, root: Element, read: () => void): Promise<void> => {
        await readChunksIn(root, relsOf(path));
        openBookmarks.clear();
        withPartRels(path, read);
    };

    // Pre-process footnotes and endnotes to be inserted inline later
    if (!config.ignoreNotes) {
        const footnotesFile = files.find(f => f.path.match(footnotesFileRegex));
        if (footnotesFile) {
            const footnotesDoc = parseXmlString(footnotesFile.content.toString(), { config });
            const footnoteXml = footnotesFile.content.toString();
            await readPart(footnotesFile.path, footnotesDoc.documentElement, () => {
                for (const node of getElementsByTagName(footnotesDoc, "w:footnote")) {
                    const id = node.getAttribute("w:id");
                    if (!id || id === "-1" || id === "0") continue;
                    const blocks: OfficeContentNode[] = [];
                    readBlocks(node, footnoteXml, blocks);
                    footnoteMap.set(id, blocks);
                }
            });
        }

        const endnotesFile = files.find(f => f.path.match(endnotesFileRegex));
        if (endnotesFile) {
            const endnotesDoc = parseXmlString(endnotesFile.content.toString(), { config });
            const endnoteXml = endnotesFile.content.toString();
            await readPart(endnotesFile.path, endnotesDoc.documentElement, () => {
                for (const node of getElementsByTagName(endnotesDoc, "w:endnote")) {
                    const id = node.getAttribute("w:id");
                    if (!id || id === "-1" || id === "0") continue;
                    const blocks: OfficeContentNode[] = [];
                    readBlocks(node, endnoteXml, blocks);
                    endnoteMap.set(id, blocks);
                }
            });
        }
    }

    // Pre-process comments
    if (!config.ignoreComments) {
        const commentsFile = files.find(f => f.path.match(commentsFileRegex));
        if (commentsFile) {
            const commentsDoc = parseXmlString(commentsFile.content.toString(), { config });
            const commentsXml = commentsFile.content.toString();
            await readPart(commentsFile.path, commentsDoc.documentElement, () => {
                for (const node of getElementsByTagName(commentsDoc, "w:comment")) {
                    const id = node.getAttribute("w:id");
                    if (!id) continue;
                    const author = node.getAttribute("w:author") || undefined;
                    const date = node.getAttribute("w:date") || undefined;
                    const initials = node.getAttribute("w:initials") || undefined;

                    commentMetadataMap.set(id, { commentId: id, author, date, initials });
                    const blocks: OfficeContentNode[] = [];
                    readBlocks(node, commentsXml, blocks);
                    commentMap.set(id, blocks);
                }
            });
        }
    }

    // Pre-process headers and footers
    const headers: OfficeContentNode[] = [];
    const footers: OfficeContentNode[] = [];
    if (!config.ignoreHeadersAndFooters) {
        for (const [regex, out] of [[headerFileRegex, headers], [footerFileRegex, footers]] as const) {
            for (const partFile of files.filter(f => f.path.match(regex))) {
                const partXml = partFile.content.toString();
                const partDoc = parseXmlString(partXml, { config });
                if (partDoc.documentElement) await readPart(partFile.path, partDoc.documentElement, () => readBlocks(partDoc.documentElement, partXml, out));
            }
        }
    }

    for (const file of files) {
        if (!file.path.match(documentFileRegex)) continue;
        const documentContent = file.content.toString();
        const doc = parseXmlString(documentContent, { config, locator: config.includeRawContent });
        const body = getFirstElementByTagName(doc, "w:body");
        if (body) await readPart(file.path, body, () => readBlocks(body, documentContent, content));
    }
    // A chunk standing where no content is read from would be lost unseen.
    for (const [element, nodes] of chunkNodes) {
        if (!placedChunks.has(element) && nodes.length > 0) chunkNotRead('unplaced', 'it stands where the document\'s blocks are not read');
    }

    // Bookmarks that only marked a note's first reference for the NOTEREF read as a reference to it: not
    // the document's anchors (written back as such, each save added one), unless a link goes to them.
    const noteMarks = new Set<string>();
    for (const name of noteRefBookmarks) if (!linkedAnchors.has(name)) noteMarks.add(name);
    if (noteMarks.size > 0) {
        const seen = new Set<OfficeContentNode>();
        const pending: OfficeContentNode[] = [];
        for (const list of [content, headers, footers]) appendAll(pending, list);
        while (pending.length) {
            const node = pending.pop()!;
            if (seen.has(node)) continue;
            seen.add(node);
            const meta = node.metadata as { anchorIds?: string[] } | undefined;
            if (meta?.anchorIds?.some(id => noteMarks.has(id))) {
                const kept = meta.anchorIds.filter(id => !noteMarks.has(id));
                if (kept.length > 0) meta.anchorIds = kept;
                else delete meta.anchorIds;
            }
            for (const list of [node.children, node.notes, node.comments]) if (list) appendAll(pending, list);
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

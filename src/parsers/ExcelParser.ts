/**
 * Excel Spreadsheet (XLSX) Parser
 * 
 * **XLSX Format Overview:**
 * XLSX is the default format for Microsoft Excel since Office 2007, based on OOXML.
 * 
 * **File Structure:**
 * - `xl/workbook.xml` - Workbook structure and sheet list
 * - `xl/worksheets/sheet1.xml` - Individual sheet data
 * - `xl/sharedStrings.xml` - Shared string table (cell text)
 * - `xl/styles.xml` - Cell styling information
 * - `xl/drawings/*` - Charts and drawings
 * - `xl/media/*` - Embedded images
 * 
 * **Key Elements:**
 * - `<row>` - Table row with row index
 * - `<c r="A1">` - Cell with reference (A1, B2, etc.)
 * - `<v>` - Cell value (number or shared string index)
 * - `<t="s">` - Cell type (s=string, n=number, b=boolean)
 * 
 * @module ExcelParser
 * @see https://www.ecma-international.org/publications-and-standards/standards/ecma-376/
 */

import { ChartMetadata, FullOfficeParserConfig, ImageMetadata, OfficeAttachment, OfficeContentNode, OfficeParserAST, OfficeWarningType, TextFormatting } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { extractChartData } from '../utils/chartUtils.js';
import { checkAbortSignal, logWarning } from '../utils/errorUtils.js';
import { createAttachment } from '../utils/imageUtils.js';
import { performOcr } from '../utils/ocrUtils.js';
import { decodeXmlEntities, getElementsByTagName, parseOfficeMetadata, parseOOXMLAppProperties, parseOOXMLCustomProperties, parseXmlString } from '../utils/xmlUtils.js';
import { extractFiles, findRequiredPart } from '../utils/zipUtils.js';

/** How a number format shows a cell value that it treats as a date or a time. */
type DateFormatKind = 'date' | 'time' | 'datetime' | 'elapsed';

/** Built-in number formats that show a date or a time (ECMA-376 Part 1, 18.8.30). */
const BUILTIN_DATE_FORMATS = new Map<string, DateFormatKind>([
    ['14', 'date'], ['15', 'date'], ['16', 'date'], ['17', 'date'], ['18', 'time'], ['19', 'time'],
    ['20', 'time'], ['21', 'time'], ['22', 'datetime'], ['45', 'time'], ['46', 'elapsed'], ['47', 'time'],
]);

/** Serials past 9999-12-31, the last date a workbook can hold, are shown as plain numbers. */
const MAX_DATE_SERIAL = 2958466;

/**
 * Classifies a number format code by the date and time parts it shows. Quoted text, escaped
 * characters and `[...]` sections (colours, locales, conditions) are not tokens, except `[h]`,
 * `[m]` and `[s]`, which show elapsed time.
 */
const getDateFormatKind = (code: string): DateFormatKind | undefined => {
    if (/\[(?:h+|m+|s+)\]/i.test(code)) return 'elapsed';
    const tokens = code.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, '').toLowerCase();
    const hasDate = tokens.includes('y') || tokens.includes('d');
    const hasTime = tokens.includes('h') || tokens.includes('s');
    if (hasDate) return hasTime ? 'datetime' : 'date';
    return hasTime ? 'time' : undefined;
};

/**
 * Writes a date/time serial as ISO text in the form its number format shows: `2024-09-30`,
 * `14:05:00`, `2024-09-30 14:05:00`, or `36:00:00` for an elapsed duration. Returns `undefined`
 * for a value that is not a serial a workbook can hold, which then stays as stored.
 */
const formatDateSerial = (value: string, kind: DateFormatKind, date1904: boolean): string | undefined => {
    const serial = Number(value);
    if (value === '' || !Number.isFinite(serial) || serial < 0 || serial >= MAX_DATE_SERIAL) return undefined;
    const pad = (n: number) => String(n).padStart(2, '0');
    // Rounded to the second, as Excel shows it: a NOW() stamp of 14:04:59.9 reads 14:05:00
    const seconds = Math.round(serial * 86400);
    const clock = (total: number) => `${pad(Math.floor(total / 60) % 60)}:${pad(total % 60)}`;
    if (kind === 'elapsed') return `${Math.floor(seconds / 3600)}:${clock(seconds)}`;
    const time = `${pad(Math.floor(seconds / 3600) % 24)}:${clock(seconds)}`;
    if (kind === 'time') return time;
    // The 1904 system counts from 1904-01-01. The 1900 system counts from 1899-12-31 and includes a
    // 1900-02-29 that never existed, as serial 60, so serials from 60 on count from 1899-12-30.
    const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, serial < 60 ? 31 : 30);
    const date = new Date(epoch + Math.floor(seconds / 86400) * 86400000).toISOString().slice(0, 10);
    return kind === 'date' ? date : `${date} ${time}`;
};

/**
 * Parses an Excel spreadsheet (.xlsx) and extracts sheets, rows, and cells.
 * 
 * @param buffer - The XLSX file as a Buffer
 * @param config - Parser configuration
 * @returns A promise resolving to the parsed AST
 */
export const parseExcel = async (buffer: Buffer, config: FullOfficeParserConfig): Promise<OfficeParserAST> => {
    // Honour cancellation requests immediately — before extracting the ZIP archive.
    // XLSX parsing involves decompressing multiple XML sheets and potentially running OCR
    // on embedded chart images, so short-circuiting here saves significant work.
    checkAbortSignal(config.abortSignal);

    const sheetsRegex = /xl\/worksheets\/sheet\d+.xml/g;
    const drawingsRegex = /xl\/drawings\/drawing\d+.xml/g;
    const chartsRegex = /xl\/charts\/chart\d+.xml/g;
    const stringsFilePath = 'xl/sharedStrings.xml';
    const mediaFileRegex = /xl\/media\/.*/;
    const corePropsFileRegex = /docProps\/core\.xml/;
    const customPropsFileRegex = /docProps\/custom\.xml/;
    const appPropsFileRegex = /docProps\/app\.xml/;

    const relsRegex = /xl\/worksheets\/_rels\/sheet\d+\.xml\.rels/g;
    const drawingRelsRegex = /xl\/drawings\/_rels\/drawing\d+\.xml\.rels/g;
    const commentsRegex = /xl\/comments\d+\.xml/g;

    const files = await extractFiles(
        buffer,
        (x: string) =>
            !!x.match(sheetsRegex) ||
            !!x.match(drawingsRegex) ||
            !!x.match(chartsRegex) ||
            (!config.ignoreComments && !!x.match(commentsRegex)) ||
            x === stringsFilePath ||
            x === 'xl/styles.xml' ||
            x === 'xl/workbook.xml' ||
            x === 'xl/_rels/workbook.xml.rels' ||
            !!x.match(corePropsFileRegex) ||
            !!x.match(customPropsFileRegex) ||
            !!x.match(appPropsFileRegex) ||
            (!!config.extractAttachments && (!!x.match(mediaFileRegex) || !!x.match(drawingRelsRegex))) ||
            ((!!config.extractAttachments || !config.ignoreComments) && !!x.match(relsRegex)),
        config.decompressionLimits,
        config
    );

    // Every workbook has xl/workbook.xml; without it the archive is not a spreadsheet.
    // Resolved up front so a file that cannot be a workbook fails before any of the parsing
    // work below, and read again further down for the sheet-name map.
    const workbookFile = findRequiredPart(files, path => path === 'xl/workbook.xml', config,
        { fileType: 'xlsx', part: 'xl/workbook.xml' });
    // Date serials count from 1900 unless the workbook uses the 1904 date system
    const date1904 = /<(?:\w+:)?workbookPr\b[^>]*\bdate1904="(?:1|true)"/.test(workbookFile.content.toString());

    // Worksheets, by contrast, are not guaranteed: a workbook holding only chartsheets is
    // valid and simply has no cell text to extract. Warn rather than fail, so the caller can
    // tell "nothing to read here" from "we read nothing".
    if (!files.some(file => !!file.path.match(sheetsRegex)))
        logWarning(OfficeWarningType.NO_WORKSHEETS_FOUND, config);

    const sharedStringsFile = files.find(f => f.path === stringsFilePath);
    // Updated to store structured content (rich text runs) or simple string
    const sharedStrings: (string | OfficeContentNode[])[] = [];

    if (sharedStringsFile) {
        const xml = parseXmlString(sharedStringsFile.content.toString());
        const siNodes = getElementsByTagName(xml, "si");
        for (const si of siNodes) {
            const runNodes = getElementsByTagName(si, "r");
            if (runNodes.length > 0) {
                // Rich text with runs
                const runs: OfficeContentNode[] = [];
                for (const run of runNodes) {
                    const tNode = getElementsByTagName(run, "t")[0];
                    if (tNode) {
                        const text = tNode.textContent || '';
                        // Extract run formatting 
                        const rPr = getElementsByTagName(run, "rPr")[0];
                        const formatting: TextFormatting = {};
                        if (rPr) {
                            if (getElementsByTagName(rPr, "b").length > 0) formatting.bold = true;
                            if (getElementsByTagName(rPr, "i").length > 0) formatting.italic = true;
                            if (getElementsByTagName(rPr, "u").length > 0) formatting.underline = true;
                            if (getElementsByTagName(rPr, "strike").length > 0) formatting.strikethrough = true;

                            const sz = getElementsByTagName(rPr, "sz")[0];
                            if (sz) formatting.size = sz.getAttribute("val") + 'pt';

                            const color = getElementsByTagName(rPr, "color")[0];
                            if (color) {
                                const rgb = color.getAttribute("rgb");
                                if (rgb) formatting.color = '#' + rgb.substring(2);
                            }

                            const rFont = getElementsByTagName(rPr, "rFont")[0];
                            if (rFont) formatting.font = rFont.getAttribute("val") || undefined;

                            const vertAlign = getElementsByTagName(rPr, "vertAlign")[0];
                            if (vertAlign) {
                                const val = vertAlign.getAttribute("val");
                                if (val === "subscript") formatting.subscript = true;
                                if (val === "superscript") formatting.superscript = true;
                            }
                        }
                        runs.push({
                            type: 'text',
                            text: text,
                            formatting: Object.keys(formatting).length > 0 ? formatting : undefined
                        });
                    }
                }
                sharedStrings.push(runs);
            } else {
                // Simple text case
                const tNodes = getElementsByTagName(si, "t");
                let text = '';
                for (const t of tNodes) {
                    text += t.textContent || '';
                }
                sharedStrings.push(text);
            }
        }
    }

    // Parse styles to build formatting map
    const stylesFile = files.find(f => f.path === 'xl/styles.xml');
    const cellFormatMap: Record<number, TextFormatting> = {};
    // A date or a time is stored as a day count and shown through its style's number format
    const cellDateFormatMap: Record<number, DateFormatKind> = {};

    if (stylesFile) {
        const xml = parseXmlString(stylesFile.content.toString());

        // Parse custom number formats (numFmtId -> format code)
        const numFmtCodes: Record<string, string> = Object.create(null);
        const numFmtsNode = getElementsByTagName(xml, "numFmts")[0];
        if (numFmtsNode) {
            for (const numFmt of getElementsByTagName(numFmtsNode, "numFmt")) {
                const id = numFmt.getAttribute("numFmtId");
                const code = numFmt.getAttribute("formatCode");
                if (id && code !== null) numFmtCodes[id] = code;
            }
        }

        // Parse fonts
        const fontsNode = getElementsByTagName(xml, "fonts")[0];
        const fonts: TextFormatting[] = [];
        if (fontsNode) {
            const fontNodes = getElementsByTagName(fontsNode, "font");
            for (const font of fontNodes) {
                const formatting: TextFormatting = {};
                if (getElementsByTagName(font, "b").length > 0) formatting.bold = true;
                if (getElementsByTagName(font, "i").length > 0) formatting.italic = true;
                if (getElementsByTagName(font, "u").length > 0) formatting.underline = true;
                if (getElementsByTagName(font, "strike").length > 0) formatting.strikethrough = true;

                const szNode = getElementsByTagName(font, "sz")[0];
                if (szNode) {
                    const val = szNode.getAttribute("val");
                    if (val) formatting.size = val + 'pt';
                }

                const colorNode = getElementsByTagName(font, "color")[0];
                if (colorNode) {
                    const rgb = colorNode.getAttribute("rgb");
                    if (rgb) formatting.color = '#' + rgb.substring(2); // Remove alpha channel
                }

                const nameNode = getElementsByTagName(font, "name")[0];
                if (nameNode) {
                    const val = nameNode.getAttribute("val");
                    if (val) formatting.font = val;
                }

                const vertAlignNode = getElementsByTagName(font, "vertAlign")[0];
                if (vertAlignNode) {
                    const val = vertAlignNode.getAttribute("val");
                    if (val === "subscript") formatting.subscript = true;
                    if (val === "superscript") formatting.superscript = true;
                }

                fonts.push(formatting);
            }
        }

        // Parse fills (for background color)
        const fillsNode = getElementsByTagName(xml, "fills")[0];
        const fills: TextFormatting[] = [];
        if (fillsNode) {
            const fillNodes = getElementsByTagName(fillsNode, "fill");
            for (const fill of fillNodes) {
                const formatting: TextFormatting = {};
                const patternFill = getElementsByTagName(fill, "patternFill")[0];
                if (patternFill) {
                    const fgColor = getElementsByTagName(patternFill, "fgColor")[0];
                    if (fgColor) {
                        const rgb = fgColor.getAttribute("rgb");
                        const theme = fgColor.getAttribute("theme");

                        if (rgb && rgb !== "00000000") { // Not default/auto
                            formatting.backgroundColor = '#' + rgb.substring(2);
                        } else if (theme) {
                            // Basic mapping for standard Office themes (Dark 1, Light 1, Dark 2, Light 2)
                            // 0: Light 1 (White), 1: Dark 1 (Black), 2: Light 2 (Tan/Gray), 3: Dark 2 (Blue/Grey)
                            const themeIdx = parseInt(theme);
                            if (themeIdx === 0) formatting.backgroundColor = '#FFFFFF';
                            else if (themeIdx === 1) formatting.backgroundColor = '#000000';
                            else if (themeIdx === 2) formatting.backgroundColor = '#EEECE1'; // Standard Light 2
                            else if (themeIdx === 3) formatting.backgroundColor = '#1F497D'; // Standard Dark 2
                        }
                    }
                }
                fills.push(formatting);
            }
        }

        // Parse cellXfs (cell format definitions)
        const cellXfsNode = getElementsByTagName(xml, "cellXfs")[0];
        if (cellXfsNode) {
            const xfNodes = getElementsByTagName(cellXfsNode, "xf");
            for (let i = 0; i < xfNodes.length; i++) {
                const xf = xfNodes[i];
                const formatting: TextFormatting = {};

                const numFmtId = xf.getAttribute("numFmtId") || '0';
                const dateFormatKind = numFmtId in numFmtCodes
                    ? getDateFormatKind(numFmtCodes[numFmtId])
                    : BUILTIN_DATE_FORMATS.get(numFmtId);
                if (dateFormatKind) cellDateFormatMap[i] = dateFormatKind;

                const fontId = xf.getAttribute("fontId");
                if (fontId) {
                    const fontIdx = parseInt(fontId);
                    if (fonts[fontIdx]) {
                        Object.assign(formatting, fonts[fontIdx]);
                    }
                }

                const fillId = xf.getAttribute("fillId");
                if (fillId) {
                    const fillIdx = parseInt(fillId);
                    if (fills[fillIdx] && fills[fillIdx].backgroundColor) {
                        formatting.backgroundColor = fills[fillIdx].backgroundColor;
                    }
                }

                const alignmentNode = getElementsByTagName(xf, "alignment")[0];
                if (alignmentNode) {
                    const horizontal = alignmentNode.getAttribute("horizontal");
                    if (horizontal === 'center' || horizontal === 'right' || horizontal === 'justify' || horizontal === 'left') {
                        formatting.alignment = horizontal;
                    }
                }

                cellFormatMap[i] = formatting;
            }
        }
    }

    const attachments: OfficeAttachment[] = [];
    const mediaFiles = files.filter(f => f.path.match(/xl\/media\/.*/));
    const chartFiles = files.filter(f => f.path.match(chartsRegex));

    // Map to store image details by drawing file path and relationship ID
    // Null-prototype: keyed by document-derived drawing paths / rel ids, so a plain object would let a
    // crafted `__proto__` key throw (or hit a shared prototype) instead of creating an own entry.
    const drawingImageMap: Record<string, Record<string, { path: string, altText?: string }>> = Object.create(null);

    if (config.extractAttachments) {
        // 1. Parse Drawing Rels to map rIds to media paths
        const drawingRelsFiles = files.filter(f => f.path.match(drawingRelsRegex));
        for (const relFile of drawingRelsFiles) {
            const drawingFilename = relFile.path.split('/').pop()?.replace('.rels', '') || '';
            const drawingPath = `xl/drawings/${drawingFilename}`;

            const relsXml = parseXmlString(relFile.content.toString());
            const relationships = getElementsByTagName(relsXml, "Relationship");

            if (!drawingImageMap[drawingPath]) {
                drawingImageMap[drawingPath] = Object.create(null);
            }

            for (const rel of relationships) {
                const id = rel.getAttribute("Id");
                const target = rel.getAttribute("Target");
                if (id && target && target.includes('media/')) {
                    // Target is usually like "../media/image1.png"
                    const mediaPath = 'xl/' + target.replace('../', '');
                    drawingImageMap[drawingPath][id] = { path: mediaPath };
                }
            }
        }

        // 2. Parse Drawings to get Alt Text and link to Rels
        const drawingFiles = files.filter(f => f.path.match(drawingsRegex));
        for (const drawingFile of drawingFiles) {
            const xml = parseXmlString(drawingFile.content.toString());
            const pics = getElementsByTagName(xml, "xdr:pic"); // SpreadsheetML drawing

            const rels = drawingImageMap[drawingFile.path] || {};

            for (const pic of pics) {
                const blipFill = getElementsByTagName(pic, "xdr:blipFill")[0];
                const blip = blipFill ? getElementsByTagName(blipFill, "a:blip")[0] : null;
                const embedId = blip ? blip.getAttribute("r:embed") : null;

                const nvPicPr = getElementsByTagName(pic, "xdr:nvPicPr")[0];
                const cNvPr = nvPicPr ? getElementsByTagName(nvPicPr, "xdr:cNvPr")[0] : null;
                const altText = cNvPr ? (cNvPr.getAttribute("descr") || cNvPr.getAttribute("name")) : undefined;

                if (embedId && rels[embedId]) {
                    rels[embedId].altText = altText || '';
                }
            }
        }

        // 3. Process Media Files
        for (const media of mediaFiles) {
            const attachment = createAttachment(media.path.split('/').pop() || 'image', media.content);

            // Try to find alt text for this media
            let altText = '';
            for (const drawingPath in drawingImageMap) {
                for (const rId in drawingImageMap[drawingPath]) {
                    if (drawingImageMap[drawingPath][rId].path === media.path) {
                        altText = drawingImageMap[drawingPath][rId].altText || '';
                        break;
                    }
                }
                if (altText) break;
            }
            if (altText) attachment.altText = altText;

            attachments.push(attachment);

            if (config.ocr) {
                if (attachment.mimeType.startsWith('image/')) {
                    try {
                        const ocrText = (await performOcr(media.content, { ...config.ocrConfig })).trim();
                        if (ocrText) {
                            attachment.ocrText = ocrText;
                        }
                    } catch (e) {
                        logWarning(OfficeWarningType.OCR_FAILED, config, attachment.name, e);
                    }
                }
            }
        }

        for (const chart of chartFiles) {
            const attachment: OfficeAttachment = {
                type: 'chart',
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                data: chart.content.toString('base64'),
                name: chart.path.split('/').pop() || '',
                extension: 'xml'
            };

            // Extract structured chart data
            try {
                const chartData = extractChartData(chart.content);
                attachment.chartData = chartData;
            } catch (e) {
                logWarning(OfficeWarningType.CHART_DATA_EXTRACTION_FAILED, config, chart.path, e);
            }

            attachments.push(attachment);
        }
    }

    // Build map of drawing rId -> chart attachment name for linking
    // Null-prototype for the same reason as drawingImageMap (document-derived keys).
    const drawingChartMap: Record<string, Record<string, string>> = Object.create(null);
    if (config.extractAttachments) {
        const drawingRelsFiles = files.filter(f => f.path.match(drawingRelsRegex));
        for (const relFile of drawingRelsFiles) {
            const drawingFilename = relFile.path.split('/').pop()?.replace('.rels', '') || '';
            const drawingPath = `xl/drawings/${drawingFilename}`;

            const relsXml = parseXmlString(relFile.content.toString());
            const relationships = getElementsByTagName(relsXml, "Relationship");

            if (!drawingChartMap[drawingPath]) {
                drawingChartMap[drawingPath] = Object.create(null);
            }

            for (const rel of relationships) {
                const id = rel.getAttribute("Id");
                const target = rel.getAttribute("Target");
                const type = rel.getAttribute("Type");
                if (id && target && type && type.includes('chart')) {
                    // Target is like "../charts/chart1.xml"
                    const chartName = target.split('/').pop() || '';
                    drawingChartMap[drawingPath][id] = chartName;
                }
            }
        }
    }

    // Parse workbook.xml to get sheet names and map them to sheet files
    const sheetNameMap: Record<string, string> = {};
    const workbookRelsFile = files.find(f => f.path === 'xl/_rels/workbook.xml.rels');

    if (workbookRelsFile) {
        // Parse rels to get rId -> file mapping
        const relsXml = parseXmlString(workbookRelsFile.content.toString());
        const relationships = getElementsByTagName(relsXml, "Relationship");
        const rIdToFile: Record<string, string> = {};

        for (const rel of relationships) {
            const rId = rel.getAttribute("Id");
            const target = rel.getAttribute("Target");
            if (rId && target) {
                // Target is like "worksheets/sheet1.xml"
                const filename = target.split('/').pop() || '';
                rIdToFile[rId] = filename;
            }
        }

        // Parse workbook.xml to get sheet name -> rId mapping
        const workbookXml = parseXmlString(workbookFile.content.toString());
        const sheets = getElementsByTagName(workbookXml, "sheet");

        for (const sheet of sheets) {
            checkAbortSignal(config.abortSignal);
            const name = sheet.getAttribute("name");
            const rId = sheet.getAttribute("r:id");
            if (name && rId && rIdToFile[rId]) {
                sheetNameMap[rIdToFile[rId]] = name;
            }
        }
    }

    const content: OfficeContentNode[] = [];

    for (const file of files) {
        if (file.path.match(mediaFileRegex)) continue;
        if (file.path === stringsFilePath) continue;
        if (file.path === 'xl/styles.xml') continue;
        if (file.path.match(drawingsRegex)) continue;
        if (file.path.match(chartsRegex)) continue;
        if (file.path.match(relsRegex)) continue;
        if (file.path.match(drawingRelsRegex)) continue;

        if (file.path.match(sheetsRegex)) {
            const sheetFilename = file.path.split('/').pop() || '';
            const relsFilename = `xl/worksheets/_rels/${sheetFilename}.rels`;
            const relsFile = files.find(f => f.path === relsFilename);

            const drawingMap: Record<string, string> = {}; // rId -> drawingPath
            // Null-prototype: keyed by the document-derived cell ref, so a crafted ref of `__proto__`
            // creates an own entry instead of throwing on `Object.prototype.push`.
            const sheetCommentsMap: Record<string, OfficeContentNode[]> = Object.create(null);

            if (relsFile) {
                const relsXml = parseXmlString(relsFile.content.toString());
                const relationships = getElementsByTagName(relsXml, "Relationship");
                for (const rel of relationships) {
                    const id = rel.getAttribute("Id");
                    const target = rel.getAttribute("Target");
                    const type = rel.getAttribute("Type");
                    
                    if (id && target && type) {
                        if (config.extractAttachments && type.includes('drawing')) {
                            drawingMap[id] = 'xl/drawings/' + target.replace('../drawings/', '');
                        } else if (!config.ignoreComments && type.includes('comments')) {
                            const commentsPath = 'xl/' + target.replace('../', '');
                            const cFile = files.find(f => f.path === commentsPath);
                            if (cFile) {
                                const cXml = parseXmlString(cFile.content.toString());
                                const commentNodes = getElementsByTagName(cXml, "comment");
                                const authorsList = getElementsByTagName(cXml, "author");
                                const authors = authorsList.map(a => a.textContent || '');
                                
                                for (const cNode of commentNodes) {
                                    const ref = cNode.getAttribute("ref");
                                    const authorId = cNode.getAttribute("authorId");
                                    const author = authorId !== null ? authors[parseInt(authorId)] : undefined;
                                    
                                    const tNodes = getElementsByTagName(cNode, "t");
                                    const text = tNodes.map(t => t.textContent || '').join('');
                                    
                                    if (ref && text) {
                                        if (!sheetCommentsMap[ref]) sheetCommentsMap[ref] = [];
                                        sheetCommentsMap[ref].push({
                                            type: 'comment',
                                            text: text,
                                            children: [{ type: 'text', text: text, formatting: {} }],
                                            metadata: author ? { author } : undefined
                                        });
                                    }
                                }
                            }
                        }
                    }
                }
            }

            const rows: OfficeContentNode[] = [];
            const sheetXml = file.content.toString();
            // regex to match <row> elements, capturing:
            // 1. attributes (e.g., r="1")
            // 2. whether it's self-closing (/>)
            // 3. inner content (for non-self-closing rows)
            const rowRegex = /<row\b([^>]*?)(?:(\/>)|(>([\s\S]*?)<\/row>))/g;
            // matchAll provides an iterator over all matches, which is much more efficient than 
            // iterating over a massive sparse row range declared in spreadsheet dimensions.
            const rowMatches = sheetXml.matchAll(rowRegex);

            /** Helper to convert Excel column string (A, B, AA, etc.) to 0-based index */
            const colToNumber = (col: string): number => {
                let num = 0;
                for (let i = 0; i < col.length; i++) {
                    num = num * 26 + (col.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
                }
                return num - 1;
            };

            let lastRowIndex = -1;

            for (const rowMatch of rowMatches) {
                checkAbortSignal(config.abortSignal);
                const rowXml = rowMatch[0];
                const rowAttrs = rowMatch[1];
                const isSelfClosing = !!rowMatch[2];
                const rowContent = rowMatch[4] || "";

                if (!isSelfClosing && !rowContent.includes('<c')) continue;

                const cells: OfficeContentNode[] = [];
                // regex to match <c> (cell) elements within a row, capturing:
                // 1. cell attributes (e.g., r="A1", t="s")
                // 2. whether it's self-closing (/>)
                // 3. inner content (e.g., <v> value)
                const cRegex = /<c\b([^>]*?)(?:(\/>)|(>([\s\S]*?)<\/c>))/g;
                const cMatches = rowContent.matchAll(cRegex);

                const rMatch = rowAttrs.match(/r="(\d+)"/);
                const rowIndex = rMatch ? parseInt(rMatch[1]) - 1 : lastRowIndex + 1;
                lastRowIndex = rowIndex;

                let lastColIndex = -1;

                for (const cMatch of cMatches) {
                    const cXml = cMatch[0];
                    const cAttrs = cMatch[1];
                    const cContent = cMatch[4] || "";

                    // Extract cell value
                    const typeMatch = cAttrs.match(/t="([a-zA-Z]+)"/);
                    const type = typeMatch ? typeMatch[1] : 'n'; // n = number (default)

                    // Extract cell style index
                    const styleMatch = cAttrs.match(/s="(\d+)"/);
                    const styleIdx = styleMatch ? parseInt(styleMatch[1]) : undefined;

                    const vMatch = cContent.match(/<v>([\s\S]*?)<\/v>/);
                    const tMatch = cContent.match(/<t\b[^>]*>([\s\S]*?)<\/t>/);

                    let text = '';
                    let cellNodes: OfficeContentNode[] = [];

                    if (type === 's' && vMatch) {
                        const idx = parseInt(vMatch[1]);
                        const content = sharedStrings[idx];
                        if (Array.isArray(content)) {
                            // Rich text runs. Share the (read-only) run nodes across every cell that
                            // references this shared string via a shallow array copy, rather than
                            // deep-copying them per cell: XLSX has no cell budget, so a large rich-text
                            // shared string referenced by many cells would otherwise amplify to N x its
                            // size in the AST. The run nodes are never mutated in place downstream.
                            cellNodes = content.slice();
                            text = cellNodes.map(n => n.text).join('');
                        } else {
                            text = content || '';
                        }
                    } else if (type === 'inlineStr' && tMatch) {
                        text = decodeXmlEntities(tMatch[1].trim());
                    } else if (type === 'b' && vMatch) {
                        // A boolean is stored as 1 or 0 and shown as TRUE or FALSE
                        text = vMatch[1].trim() === '1' ? 'TRUE' : 'FALSE';
                    } else if (vMatch) {
                        text = vMatch[1].trim();
                        // A number in a date or time style is a day count: show the date or time instead
                        const dateFormatKind = type === 'n' && styleIdx !== undefined ? cellDateFormatMap[styleIdx] : undefined;
                        if (dateFormatKind) text = formatDateSerial(text, dateFormatKind, date1904) ?? text;
                    }

                    // Parse cell coordinate
                    const coordMatch = cAttrs.match(/r="([A-Z]+)(\d+)"/);
                    let colIndex: number;
                    let ref: string | undefined;
                    if (coordMatch) {
                        ref = coordMatch[1] + coordMatch[2];
                        colIndex = colToNumber(coordMatch[1]);
                        // If row index is missing in cell coord (unlikely but possible), use rowIndex
                    } else {
                        colIndex = lastColIndex + 1;
                    }
                    lastColIndex = colIndex;

                    if (text || cellNodes.length > 0) {
                        const cellFormatting = (styleIdx !== undefined && cellFormatMap[styleIdx]) ? cellFormatMap[styleIdx] : {};

                        if (cellNodes.length > 0) {
                            // If we have specific runs, merge cell styles into them if run style is missing
                            // But usually run style overrides cell style (except maybe background)
                            for (const node of cellNodes) {
                                if (!node.formatting) node.formatting = {};
                                // Cell background always applies
                                if (cellFormatting.backgroundColor) node.formatting.backgroundColor = cellFormatting.backgroundColor;
                                // Cell alignment always applies
                                if (cellFormatting.alignment) node.formatting.alignment = cellFormatting.alignment;

                                // Font defaults from cell style if not in run
                                if (!node.formatting.font && cellFormatting.font) node.formatting.font = cellFormatting.font;
                                if (!node.formatting.size && cellFormatting.size) node.formatting.size = cellFormatting.size;
                            }
                        } else {
                            // Simple text node
                            cellNodes.push({
                                type: 'text',
                                text: text,
                                formatting: cellFormatting
                            });
                        }

                        const commentsNodeList = (ref && sheetCommentsMap[ref]) ? sheetCommentsMap[ref] : undefined;

                        const cellNode: OfficeContentNode = {
                            type: 'cell',
                            text: text,
                            children: cellNodes,
                            comments: commentsNodeList,
                            metadata: { row: rowIndex, col: colIndex }
                        };
                        if (config.includeRawContent) {
                            cellNode.rawContent = cXml;
                        }
                        cells.push(cellNode);
                    }
                }

                if (cells.length > 0) {
                    const rowNode: OfficeContentNode = {
                        type: 'row',
                        children: cells,
                        metadata: undefined
                    };
                    if (config.includeRawContent) {
                        rowNode.rawContent = rowXml;
                    }
                    rows.push(rowNode);
                }
            }

            // Handle Drawings in Sheet (images and charts)
            if (config.extractAttachments) {
                const drawingMatches = file.content.toString().match(/<drawing r:id="(.*?)"/g);
                if (drawingMatches) {
                    for (const match of drawingMatches) {
                        const rIdMatch = match.match(/r:id="(.*?)"/);
                        const rId = rIdMatch ? rIdMatch[1] : null;

                        if (rId && drawingMap[rId]) {
                            const drawingPath = drawingMap[rId];

                            // Find all images in this drawing
                            const images = drawingImageMap[drawingPath];
                            if (images) {
                                for (const imgId in images) {
                                    const imgInfo = images[imgId];
                                    const attachment = attachments.find(a => a.name === imgInfo.path.split('/').pop());
                                    if (attachment) {
                                        const imageNode: OfficeContentNode = {
                                            type: 'image',
                                            text: '', // Will be populated by assignAttachmentData
                                            children: [],
                                            metadata: {
                                                attachmentName: attachment.name || 'unknown',
                                                altText: imgInfo.altText || undefined
                                            } as ImageMetadata
                                        };
                                        rows.push(imageNode);
                                    }
                                }
                            }

                            // Find all charts in this drawing
                            const charts = drawingChartMap[drawingPath];
                            if (charts) {
                                for (const chartRId in charts) {
                                    const chartName = charts[chartRId];
                                    const attachment = attachments.find(a => a.name === chartName);
                                    if (attachment) {
                                        const chartNode: OfficeContentNode = {
                                            type: 'chart',
                                            text: '', // Will be populated by assignAttachmentData
                                            children: [],
                                            metadata: {
                                                attachmentName: chartName
                                            } as ChartMetadata
                                        };
                                        rows.push(chartNode);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Get proper sheet name from workbook.xml mapping, fallback to filename
            const sheetFileName = file.path.split('/').pop() || 'Sheet';
            const sheetName = sheetNameMap[sheetFileName] || sheetFileName;

            content.push({
                type: 'sheet',
                children: rows,
                metadata: { sheetName },
                rawContent: config.includeRawContent ? file.content.toString() : undefined
            });
        }
    }

    const corePropsFile = files.find(f => f.path.match(corePropsFileRegex));
    const metadata = corePropsFile ? parseOfficeMetadata(corePropsFile.content.toString()) : {};
    const customPropsFile = files.find(f => f.path.match(customPropsFileRegex));
    if (customPropsFile) {
        const customProperties = parseOOXMLCustomProperties(customPropsFile.content.toString());
        if (Object.keys(customProperties).length > 0) metadata.customProperties = customProperties;
    }
    const appPropsFile = files.find(f => f.path.match(appPropsFileRegex));
    if (appPropsFile) {
        const appProperties = parseOOXMLAppProperties(appPropsFile.content.toString());
        if (Object.keys(appProperties).length > 0) metadata.nativeProperties = appProperties;
    }

    // Link OCR text and chart data to content nodes (like PPTX parser)
    const assignAttachmentData = (nodes: OfficeContentNode[]) => {
        for (const node of nodes) {
            if ('attachmentName' in (node.metadata || {})) {
                const meta = node.metadata as ImageMetadata | ChartMetadata;
                const attachment = attachments.find(a => a.name === meta.attachmentName);
                if (attachment) {
                    if (node.type === 'image') {
                        // Link OCR text to image node
                        if (attachment.ocrText) {
                            node.text = attachment.ocrText;
                        }
                        // Copy altText to attachment
                        if ((meta as ImageMetadata).altText) {
                            attachment.altText = (meta as ImageMetadata).altText;
                        }
                    }
                    if (node.type === 'chart') {
                        // Link chart data text to chart node
                        if (attachment.chartData) {
                            node.text = attachment.chartData.rawTexts.join(config.newlineDelimiter);
                        }
                    }
                }
            }
            if (node.children) {
                assignAttachmentData(node.children);
            }
        }
    };
    assignAttachmentData(content);

    return createAST(
        'xlsx',
        metadata,
        content,
        attachments,
        config,
        undefined,
    );
};

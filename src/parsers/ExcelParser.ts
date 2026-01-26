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

import { Block, ChartBlock, ChartMetadata, CoordinateData, ImageBlock, ImageMetadata, OfficeAttachment, OfficeContentNode, OfficeParserAST, OfficeParserConfig, TableBlock, TextBlock, TextFormatting } from '../types';
import { extractChartData } from '../utils/chartUtils';
import { logWarning } from '../utils/errorUtils';
import { createAttachment } from '../utils/imageUtils';
import { performOcr } from '../utils/ocrUtils';
import { getElementsByTagName, parseOfficeMetadata, parseXmlString } from '../utils/xmlUtils';
import { extractFiles } from '../utils/zipUtils';

/**
 * Parses an Excel spreadsheet (.xlsx) and extracts sheets, rows, and cells.
 * 
 * @param buffer - The XLSX file as a Buffer
 * @param config - Parser configuration
 * @returns A promise resolving to the parsed AST
 */
export const parseExcel = async (buffer: Buffer, config: OfficeParserConfig): Promise<OfficeParserAST> => {
    const sheetsRegex = /xl\/worksheets\/sheet\d+.xml/g;
    const drawingsRegex = /xl\/drawings\/drawing\d+.xml/g;
    const chartsRegex = /xl\/charts\/chart\d+.xml/g;
    const stringsFilePath = 'xl/sharedStrings.xml';
    const mediaFileRegex = /xl\/media\/.*/;
    const corePropsFileRegex = /docProps\/core\.xml/;

    const relsRegex = /xl\/worksheets\/_rels\/sheet\d+\.xml\.rels/g;
    const drawingRelsRegex = /xl\/drawings\/_rels\/drawing\d+\.xml\.rels/g;

    const files = await extractFiles(buffer, (x: string) =>
        !!x.match(sheetsRegex) ||
        !!x.match(drawingsRegex) ||
        !!x.match(chartsRegex) ||
        x === stringsFilePath ||
        x === 'xl/styles.xml' ||
        x === 'xl/workbook.xml' ||
        x === 'xl/_rels/workbook.xml.rels' ||
        !!x.match(corePropsFileRegex) ||
        (!!config.extractAttachments && (!!x.match(mediaFileRegex) || !!x.match(relsRegex) || !!x.match(drawingRelsRegex)))
    );

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

    if (stylesFile) {
        const xml = parseXmlString(stylesFile.content.toString());

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
    const drawingImageMap: Record<string, Record<string, { path: string, altText?: string }>> = {};

    if (config.extractAttachments) {
        // 1. Parse Drawing Rels to map rIds to media paths
        const drawingRelsFiles = files.filter(f => f.path.match(drawingRelsRegex));
        for (const relFile of drawingRelsFiles) {
            const drawingFilename = relFile.path.split('/').pop()?.replace('.rels', '') || '';
            const drawingPath = `xl/drawings/${drawingFilename}`;

            const relsXml = parseXmlString(relFile.content.toString());
            const relationships = getElementsByTagName(relsXml, "Relationship");

            if (!drawingImageMap[drawingPath]) {
                drawingImageMap[drawingPath] = {};
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
                        const ocrText = await performOcr(media.content, config.ocrLanguage);
                        if (ocrText.trim()) {
                            attachment.ocrText = ocrText.trim();
                        }
                    } catch (e) {
                        logWarning(`OCR failed for ${attachment.name}:`, config, e);
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
                logWarning(`Failed to extract chart data from ${chart.path}:`, config, e);
            }

            attachments.push(attachment);
        }
    }

    // Build map of drawing rId -> chart attachment name for linking
    const drawingChartMap: Record<string, Record<string, string>> = {};
    if (config.extractAttachments) {
        const drawingRelsFiles = files.filter(f => f.path.match(drawingRelsRegex));
        for (const relFile of drawingRelsFiles) {
            const drawingFilename = relFile.path.split('/').pop()?.replace('.rels', '') || '';
            const drawingPath = `xl/drawings/${drawingFilename}`;

            const relsXml = parseXmlString(relFile.content.toString());
            const relationships = getElementsByTagName(relsXml, "Relationship");

            if (!drawingChartMap[drawingPath]) {
                drawingChartMap[drawingPath] = {};
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
    const workbookFile = files.find(f => f.path === 'xl/workbook.xml');
    const workbookRelsFile = files.find(f => f.path === 'xl/_rels/workbook.xml.rels');

    if (workbookFile && workbookRelsFile) {
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
            const name = sheet.getAttribute("name");
            const rId = sheet.getAttribute("r:id");
            if (name && rId && rIdToFile[rId]) {
                sheetNameMap[rIdToFile[rId]] = name;
            }
        }
    }

    const content: OfficeContentNode[] = [];
    const rawContents: string[] = [];
    // Map to track OfficeContentNode -> XML Element for position extraction
    const elementMap = new Map<OfficeContentNode, Element>();

    for (const file of files) {
        if (file.path.match(mediaFileRegex)) continue;
        if (file.path === stringsFilePath) continue;
        if (file.path === 'xl/styles.xml') continue;
        if (file.path.match(drawingsRegex)) continue;
        if (file.path.match(chartsRegex)) continue;
        if (file.path.match(relsRegex)) continue;
        if (file.path.match(drawingRelsRegex)) continue;

        if (file.path.match(sheetsRegex)) {
            if (config.includeRawContent) {
                rawContents.push(file.content.toString());
            }

            const rows: OfficeContentNode[] = [];
            const rowRegex = /<row.*?>[\s\S]*?<\/row>/g;
            const rowMatches = file.content.toString().match(rowRegex);

            if (rowMatches) {
                for (const rowXml of rowMatches) {
                    const cells: OfficeContentNode[] = [];
                    const cRegex = /<c.*?>[\s\S]*?<\/c>/g;
                    const cMatches = rowXml.match(cRegex);

                    const rMatch = rowXml.match(/r="(\d+)"/);
                    const rowIndex = rMatch ? parseInt(rMatch[1]) - 1 : 0;

                    if (cMatches) {
                        for (const cXml of cMatches) {
                            // Extract cell value
                            const typeMatch = cXml.match(/t="([a-z]+)"/);
                            const type = typeMatch ? typeMatch[1] : 'n'; // n = number (default)

                            const vMatch = cXml.match(/<v>(.*?)<\/v>/);
                            const tMatch = cXml.match(/<t>(.*?)<\/t>/);

                            let text = '';
                            let cellNodes: OfficeContentNode[] = [];

                            if (type === 's' && vMatch) {
                                const idx = parseInt(vMatch[1]);
                                const content = sharedStrings[idx];
                                if (Array.isArray(content)) {
                                    // Rich text runs
                                    // Deep copy runs to avoid reference issues if reused
                                    cellNodes = JSON.parse(JSON.stringify(content));
                                    text = cellNodes.map(n => n.text).join('');
                                } else {
                                    text = content || '';
                                }
                            } else if (type === 'inlineStr' && tMatch) {
                                text = tMatch[1];
                            } else if (vMatch) {
                                text = vMatch[1];
                            }

                            // Parse cell coordinate
                            const coordMatch = cXml.match(/r="([A-Z]+)(\d+)"/);
                            const colStr = coordMatch ? coordMatch[1] : '';
                            const colIndex = colStr.charCodeAt(0) - 'A'.charCodeAt(0);

                            if (text || cellNodes.length > 0) {
                                // Extract cell style index
                                const styleMatch = cXml.match(/s="(\d+)"/);
                                const styleIdx = styleMatch ? parseInt(styleMatch[1]) : undefined;
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

                                const cellNode: OfficeContentNode = {
                                    type: 'cell',
                                    text: text,
                                    children: cellNodes,
                                    metadata: { row: rowIndex, col: colIndex }
                                };
                                if (config.includeRawContent) {
                                    cellNode.rawContent = cXml;
                                }
                                cells.push(cellNode);
                            }
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
            }

            // Handle Drawings in Sheet (images and charts)
            if (config.extractAttachments) {
                // Parse Sheet Rels to map drawing rIds
                const sheetFilename = file.path.split('/').pop() || '';
                const relsFilename = `xl/worksheets/_rels/${sheetFilename}.rels`;
                const relsFile = files.find(f => f.path === relsFilename);

                const drawingMap: Record<string, string> = {}; // rId -> drawingPath

                if (relsFile) {
                    const relsXml = parseXmlString(relsFile.content.toString());
                    const relationships = getElementsByTagName(relsXml, "Relationship");
                    for (const rel of relationships) {
                        const id = rel.getAttribute("Id");
                        const target = rel.getAttribute("Target");
                        const type = rel.getAttribute("Type");
                        if (id && target && type && type.includes('drawing')) {
                            drawingMap[id] = 'xl/drawings/' + target.replace('../drawings/', '');
                        }
                    }
                }

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
                            node.text = attachment.chartData.rawTexts.join(config.newlineDelimiter || '\n');
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

    /**
     * Extracts coordinate data from Excel drawing XML anchor elements.
     * Handles three anchor types:
     * - xdr:twoCellAnchor - anchored between two cells (has xdr:from and xdr:to)
     * - xdr:oneCellAnchor - anchored to single cell (has xdr:from and xdr:ext)
     * - xdr:absoluteAnchor - absolute positioning (has xdr:pos and xdr:ext)
     * Values are in EMU (English Metric Units): 1 inch = 914400 EMU.
     */
    const extractCoordinates = (anchorElement: Element): CoordinateData | undefined => {
        // Try twoCellAnchor first
        let from = getElementsByTagName(anchorElement, "xdr:from")[0];
        let to = getElementsByTagName(anchorElement, "xdr:to")[0];
        let ext = getElementsByTagName(anchorElement, "xdr:ext")[0];
        
        if (from && to && ext) {
            // twoCellAnchor: calculate position from cell anchors
            // For simplicity, use the extent (size) and calculate position from from cell
            const fromCol = getElementsByTagName(from, "xdr:col")[0]?.textContent || "0";
            const fromRow = getElementsByTagName(from, "xdr:row")[0]?.textContent || "0";
            const fromColOff = getElementsByTagName(from, "xdr:colOff")[0]?.textContent || "0";
            const fromRowOff = getElementsByTagName(from, "xdr:rowOff")[0]?.textContent || "0";
            
            const width = ext ? parseInt(ext.getAttribute("cx") || "0", 10) : 0;
            const height = ext ? parseInt(ext.getAttribute("cy") || "0", 10) : 0;
            
            // Calculate approximate position (Excel uses cell-based positioning)
            // For now, use the offsets as position
            const x = parseInt(fromColOff, 10);
            const y = parseInt(fromRowOff, 10);
            
            if (width > 0 && height > 0) {
                return { x, y, width, height };
            }
        }
        
        // Try oneCellAnchor
        if (from && ext && !to) {
            const fromColOff = getElementsByTagName(from, "xdr:colOff")[0]?.textContent || "0";
            const fromRowOff = getElementsByTagName(from, "xdr:rowOff")[0]?.textContent || "0";
            
            const width = ext ? parseInt(ext.getAttribute("cx") || "0", 10) : 0;
            const height = ext ? parseInt(ext.getAttribute("cy") || "0", 10) : 0;
            
            const x = parseInt(fromColOff, 10);
            const y = parseInt(fromRowOff, 10);
            
            if (width > 0 && height > 0) {
                return { x, y, width, height };
            }
        }
        
        // Try absoluteAnchor
        const pos = getElementsByTagName(anchorElement, "xdr:pos")[0];
        if (pos && ext) {
            const x = parseInt(pos.getAttribute("x") || "0", 10);
            const y = parseInt(pos.getAttribute("y") || "0", 10);
            const width = parseInt(ext.getAttribute("cx") || "0", 10);
            const height = parseInt(ext.getAttribute("cy") || "0", 10);
            
            if (width > 0 && height > 0) {
                return { x, y, width, height };
            }
        }
        
        return undefined;
    };

    // Parse drawing files to map chart/image nodes to their anchor elements for position extraction
    if (config.extractAttachments) {
        const drawingFiles = files.filter(f => f.path.match(drawingsRegex));
        for (const drawingFile of drawingFiles) {
            const xml = parseXmlString(drawingFile.content.toString());
            
            // Find all anchors (twoCellAnchor, oneCellAnchor, absoluteAnchor)
            const twoCellAnchors = getElementsByTagName(xml, "xdr:twoCellAnchor");
            const oneCellAnchors = getElementsByTagName(xml, "xdr:oneCellAnchor");
            const absoluteAnchors = getElementsByTagName(xml, "xdr:absoluteAnchor");
            
            const allAnchors = [...twoCellAnchors, ...oneCellAnchors, ...absoluteAnchors];
            
            for (const anchor of allAnchors) {
                // Find chart or picture in this anchor
                const chart = getElementsByTagName(anchor, "xdr:graphicFrame")[0];
                const pic = getElementsByTagName(anchor, "xdr:pic")[0];
                
                if (chart) {
                    // Find the chart node in content that matches this chart
                    const cChart = getElementsByTagName(chart, "c:chart")[0];
                    const rId = cChart?.getAttribute("r:id");
                    if (rId) {
                        // Find chart node by matching attachment name
                        const drawingPath = drawingFile.path;
                        const charts = drawingChartMap[drawingPath];
                        if (charts && charts[rId]) {
                            const chartName = charts[rId];
                            // Find the chart node in content
                            const findChartNode = (nodes: OfficeContentNode[]): OfficeContentNode | null => {
                                for (const node of nodes) {
                                    if (node.type === 'chart') {
                                        const meta = node.metadata as ChartMetadata;
                                        if (meta?.attachmentName === chartName) {
                                            return node;
                                        }
                                    }
                                    if (node.children) {
                                        const found = findChartNode(node.children);
                                        if (found) return found;
                                    }
                                }
                                return null;
                            };
                            const chartNode = findChartNode(content);
                            if (chartNode) {
                                elementMap.set(chartNode, anchor);
                            }
                        }
                    }
                }
                
                if (pic) {
                    // Find the image node in content that matches this picture
                    const blipFill = getElementsByTagName(pic, "xdr:blipFill")[0];
                    const blip = blipFill ? getElementsByTagName(blipFill, "a:blip")[0] : null;
                    const embedId = blip ? blip.getAttribute("r:embed") : null;
                    
                    if (embedId) {
                        const drawingPath = drawingFile.path;
                        const images = drawingImageMap[drawingPath];
                        if (images && images[embedId]) {
                            const imgPath = images[embedId].path;
                            const imgName = imgPath.split('/').pop();
                            // Find the image node in content
                            const findImageNode = (nodes: OfficeContentNode[]): OfficeContentNode | null => {
                                for (const node of nodes) {
                                    if (node.type === 'image') {
                                        const meta = node.metadata as ImageMetadata;
                                        if (meta?.attachmentName === imgName) {
                                            return node;
                                        }
                                    }
                                    if (node.children) {
                                        const found = findImageNode(node.children);
                                        if (found) return found;
                                    }
                                }
                                return null;
                            };
                            const imageNode = findImageNode(content);
                            if (imageNode) {
                                elementMap.set(imageNode, anchor);
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * Converts a sheet node to a table block.
     */
    const convertSheetToTableBlock = (sheetNode: OfficeContentNode, position?: CoordinateData): TableBlock => {
        const rows: Array<{ cols: Array<{ value: string }> }> = [];
        
        if (sheetNode.children) {
            for (const rowNode of sheetNode.children) {
                if (rowNode.type === 'row' && rowNode.children) {
                    const cols: Array<{ value: string }> = [];
                    
                    for (const cellNode of rowNode.children) {
                        if (cellNode.type === 'cell') {
                            // Extract text from cell (including nested content)
                            const getCellText = (node: OfficeContentNode): string => {
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
    const convertChartToBlock = (chartNode: OfficeContentNode, attachments: OfficeAttachment[], position?: CoordinateData): ChartBlock | null => {
        if (chartNode.type !== 'chart') return null;
        
        const chartMetadata = chartNode.metadata as ChartMetadata | undefined;
        
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
    const extractBlocksFromContent = (nodes: OfficeContentNode[], attachments: OfficeAttachment[], elementMap: Map<OfficeContentNode, Element>): Block[] => {
        const blocks: Block[] = [];
        
        const traverse = (node: OfficeContentNode) => {
            // Get position from element map
            const element = elementMap.get(node);
            const position = element ? extractCoordinates(element) : undefined;
            
            // Process node based on type - prioritize specific block types
            if (node.type === 'chart') {
                const chartBlock = convertChartToBlock(node, attachments, position);
                if (chartBlock) {
                    blocks.push(chartBlock);
                }
                // Don't traverse children of charts
                return;
            } else if (node.type === 'image') {
                const imageMetadata = node.metadata as ImageMetadata | undefined;
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
            } else if (node.type === 'sheet') {
                // For sheets, extract table block from rows/cells, but also traverse to find charts/images
                // First, collect rows that contain only cells (for table block)
                const tableRows: OfficeContentNode[] = [];
                const otherRows: OfficeContentNode[] = [];
                
                if (node.children) {
                    for (const rowNode of node.children) {
                        if (rowNode.type === 'row') {
                            // Check if row contains only cells (no charts/images)
                            const hasOnlyCells = !rowNode.children || rowNode.children.every(child => child.type === 'cell');
                            if (hasOnlyCells) {
                                tableRows.push(rowNode);
                            } else {
                                otherRows.push(rowNode);
                            }
                        } else {
                            // Charts/images are added directly to rows array
                            otherRows.push(rowNode);
                        }
                    }
                }
                
                // Create table block from rows with cells
                if (tableRows.length > 0) {
                    const sheetWithTableRows: OfficeContentNode = {
                        ...node,
                        children: tableRows
                    };
                    const tableBlock = convertSheetToTableBlock(sheetWithTableRows, position);
                    if (tableBlock.rows.length > 0) {
                        blocks.push(tableBlock);
                    }
                }
                
                // Traverse other rows (which may contain charts/images) and cells for text blocks
                for (const rowNode of otherRows) {
                    traverse(rowNode);
                }
                
                // Also traverse table rows for text blocks from cells
                for (const rowNode of tableRows) {
                    if (rowNode.children) {
                        for (const cellNode of rowNode.children) {
                            if (cellNode.type === 'cell' && cellNode.text && cellNode.text.trim()) {
                                const cellElement = elementMap.get(cellNode);
                                const cellPosition = cellElement ? extractCoordinates(cellElement) : undefined;
                                blocks.push({
                                    type: 'text',
                                    content: cellNode.text.trim(),
                                    ...(cellPosition ? { position: cellPosition } : {})
                                });
                            }
                        }
                    }
                }
                
                return;
            } else if (node.type === 'cell' && node.text && node.text.trim()) {
                // Create text block for cells with text content
                blocks.push({
                    type: 'text',
                    content: node.text.trim(),
                    ...(position ? { position } : {})
                });
            }
            
            // Recursively process children
            if (node.children) {
                for (const child of node.children) {
                    traverse(child);
                }
            }
        };
        
        // Traverse all top-level nodes (sheets)
        for (const node of nodes) {
            traverse(node);
        }
        
        return blocks;
    };

    /**
     * Extracts images list from attachments.
     */
    const extractImagesList = (attachments: OfficeAttachment[]): Array<{ buffer: Buffer; mimeType: string; filename?: string }> => {
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
    }).filter(t => t != '').join(config.newlineDelimiter ?? '\n');

    // Extract blocks
    const blocks = extractBlocksFromContent(content, attachments, elementMap);

    // Extract images
    const images = extractImagesList(attachments);


    return {
        type: 'xlsx',
        metadata: metadata,
        content: content,
        attachments: attachments,
        fullText,
        blocks,
        images,
        toText: () => fullText
    };
};

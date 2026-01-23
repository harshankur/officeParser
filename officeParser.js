#!/usr/bin/env node

// @ts-check

const concat = require('concat-stream');
const { DOMParser } = require('@xmldom/xmldom');
const fileType = require('file-type');
const fs = require('fs');
const yauzl = require('yauzl');

/** Header for error messages */
const ERRORHEADER = '[OfficeParser]: ';
/** Error messages */
const ERRORMSG = {
    extensionUnsupported: (ext) => `Sorry, OfficeParser currently support docx, pptx, xlsx, odt, odp, ods, pdf files only. Create a ticket in Issues on github to add support for ${ext} files. Stay tuned for further updates.`,
    fileCorrupted: (filepath) => `Your file ${filepath} seems to be corrupted. If you are sure it is fine, please create a ticket in Issues on github with the file to reproduce error.`,
    fileDoesNotExist: (filepath) => `File ${filepath} could not be found! Check if the file exists or verify if the relative path to the file is correct from your terminal's location.`,
    locationNotFound: (location) => `Entered location ${location} is not reachable! Please make sure that the entered directory location exists. Check relative paths and reenter.`,
    improperArguments: `Improper arguments`,
    improperBuffers: `Error occured while reading the file buffers`,
    invalidInput: `Invalid input type: Expected a Buffer or a valid file path`
};

/** Returns parsed xml document for a given xml text.
 * @param {string} xml The xml string from the doc file
 * @returns {XMLDocument}
 */
const parseString = (xml) => {
    const parser = new DOMParser();
    return parser.parseFromString(xml, 'text/xml');
};

/** MIME type mapping for common image extensions
 * @type {Object.<string, string>}
 */
const MIME_TYPE_MAP = {
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'png': 'image/png',
    'gif': 'image/gif',
    'bmp': 'image/bmp',
    'tif': 'image/tiff',
    'tiff': 'image/tiff',
    'webp': 'image/webp',
    'svg': 'image/svg+xml',
    'ico': 'image/x-icon'
};

/** Get proper MIME type from filename extension
 * @param {string} filename The filename with extension
 * @returns {string} The MIME type
 */
function getMimeTypeFromFilename(filename) {
    const extension = filename.split('.').pop()?.toLowerCase();
    return MIME_TYPE_MAP[extension] || 'application/octet-stream';
}

/** Extracts text from XML paragraphs following Office XML patterns
 * @param {XMLDocument} xmlDoc - Parsed XML document
 * @param {string} paragraphTag - Tag name for paragraphs (e.g., "w:p", "a:p")
 * @param {string} textTag - Tag name for text nodes (e.g., "w:t", "a:t")
 * @param {string} delimiter - Delimiter for joining paragraphs
 * @returns {string} Extracted text
 */
function extractTextFromXmlParagraphs(xmlDoc, paragraphTag, textTag, delimiter) {
    const paragraphNodes = xmlDoc.getElementsByTagName(paragraphTag);
    return Array.from(paragraphNodes)
        .filter(paragraphNode => paragraphNode.getElementsByTagName(textTag).length !== 0)
        .map(paragraphNode => {
            const textNodes = paragraphNode.getElementsByTagName(textTag);
            return Array.from(textNodes)
                .filter(textNode => textNode.childNodes[0]?.nodeValue)
                .map(textNode => textNode.childNodes[0].nodeValue)
                .join('');
        })
        .join(delimiter);
}

/** Retrieves an image from PDF.js object stores with timeout fallback
 * @param {Object} page - PDF page object
 * @param {string} imageName - Image resource identifier
 * @param {number} [timeout=500] - Timeout in milliseconds (default: 500ms to prevent hanging on missing resources)
 * @returns {Promise<Object|null>} Image object or null if not found
 */
async function getPdfImageResource(page, imageName, timeout = 500) {
    const getFromStore = (store) => new Promise(resolve => {
        store.get(imageName, obj => resolve(obj || null));
        setTimeout(() => resolve(null), timeout);
    });

    return await getFromStore(page.commonObjs) || await getFromStore(page.objs);
}

/** @typedef {Object} OfficeParserConfig
 * @property {boolean} [outputErrorToConsole] Flag to show all the logs to console in case of an error irrespective of your own handling. Default is false.
 * @property {string}  [newlineDelimiter]     The delimiter used for every new line in places that allow multiline text like word. Default is \n.
 * @property {boolean} [ignoreNotes]          Flag to ignore notes from parsing in files like powerpoint. Default is false. It includes notes in the parsed text by default.
 * @property {boolean} [putNotesAtLast]       Flag, if set to true, will collectively put all the parsed text from notes at last in files like powerpoint. Default is false. It puts each notes right after its main slide content. If ignoreNotes is set to true, this flag is also ignored.
 * @property {boolean} [extractImages]        Flag to extract images from files like docx and pdf. Default is false. If set to true, the return object will contain an 'images' array.
 * @property {boolean} [extractCharts]        Flag to extract charts from files like docx. Default is false. If set to true, the return object will contain an 'charts' array.
 */

/**
 * @typedef {Object} TextBlock
 * @property {'text'} type
 * @property {string} content
 */

/**
 * @typedef {Object} ImageBlock
 * @property {'image'} type
 * @property {Buffer} buffer
 * @property {string} mimeType
 * @property {string} [filename]
 */

/**
 * @typedef {Object} TableBlock
 * @property {'table'} type
 * @property {string} name
 * @property {Array<{ cols: Array<{ value: string }> }>} rows
 */

/**
 * @typedef {Object} HierarchicalCategory
 * @property {Array<string>} levels Array of hierarchical level labels
 * @property {string} value The leaf value for this category
 */

/**
 * @typedef {Object} ChartBlock
 * @property {'chart'} type
 * @property {string} chartType
 * @property {Array<{ categories: Array<string | HierarchicalCategory>, values: Array<number> }>} series
 */

/**
 * @typedef {TextBlock | ImageBlock | TableBlock | ChartBlock} Block
 */

/**
 * @typedef {Object} ParseOfficeResult
 * @property {string} text The full extracted text content, preserved for backwards compatibility.
 * @property {Block[]} blocks An ordered array of content blocks (e.g., text, images, tables, charts) preserving the document structure.
 * @property {Array<{ name: string, rows: Array<{ cols: Array<{ value: string }> }> }>} [tables] Array of all extracted tables (for convenience).
 * @property {Array<{ chartType: string, series: Array<{ categories: Array<string | HierarchicalCategory>, values: Array<number> }> }>} [charts] Array of all extracted charts (for convenience).
 */

/** Creates a text block
 * @param {string} content The text content
 * @returns {TextBlock}
 */
function createTextBlock(content) {
    return { type: 'text', content };
}

/** Creates an image block
 * @param {Buffer} buffer The image data
 * @param {string} mimeType The MIME type of the image
 * @param {string} [filename] Optional filename
 * @returns {ImageBlock}
 */
function createImageBlock(buffer, mimeType, filename) {
    return { type: 'image', buffer, mimeType, filename };
}

/** Creates a table block
 * @param {string} name The table name
 * @param {Array<{ cols: Array<{ value: string }> }>} rows The table rows
 * @returns {TableBlock}
 */
function createTableBlock(name, rows) {
    return { type: 'table', name, rows };
}

/** Creates a chart block
 * @param {string} chartType The type of chart
 * @param {Array<{ categories: Array<string | HierarchicalCategory>, values: Array<number> }>} series The chart series with categories and values
 * @returns {ChartBlock}
 */
function createChartBlock(chartType, series) {
    return { type: 'chart', chartType, series };
}

/** Extract text from a cell node (w:tc)
 * @param {Element} cellNode The cell node
 * @returns {string} Extracted text from the cell
 */
function extractCellText(cellNode) {
    const paragraphs = cellNode.getElementsByTagName('w:p');
    const cellTexts = [];
    for (let i = 0; i < paragraphs.length; i++) {
        const textNodes = paragraphs[i].getElementsByTagName('w:t');
        const paraText = Array.from(textNodes)
            .map(t => t.childNodes[0]?.nodeValue || '')
            .join('');
        if (paraText) {
            cellTexts.push(paraText);
        }
    }
    return cellTexts.join(' ');
}

/** Extract text from a paragraph node (w:p)
 * @param {Element} paraNode The paragraph node
 * @returns {string} Extracted text from the paragraph
 */
function extractParagraphText(paraNode) {
    const textNodes = paraNode.getElementsByTagName('w:t');
    return Array.from(textNodes)
        .map(t => t.childNodes[0]?.nodeValue || '')
        .join('');
}

/** Extract table rows from a table node
 * @param {Element} tableNode The table node (w:tbl)
 * @returns {Array<{ cols: Array<{ value: string }> }>} Array of rows with columns
 */
function extractTableRows(tableNode) {
    const rows = [];
    const rowNodes = tableNode.getElementsByTagName('w:tr');
    for (let i = 0; i < rowNodes.length; i++) {
        const rowNode = rowNodes[i];
        const cellNodes = rowNode.getElementsByTagName('w:tc');
        const cols = [];
        for (let j = 0; j < cellNodes.length; j++) {
            const cellNode = cellNodes[j];
            // Check for merged cells (gridSpan)
            const gridSpan = cellNode.getElementsByTagName('w:gridSpan')[0];
            const span = gridSpan ? parseInt(gridSpan.getAttribute('w:val') || '1', 10) : 1;
            const cellText = extractCellText(cellNode);
            cols.push({ value: cellText });
            // If cell is merged, add empty cells for the span
            for (let k = 1; k < span; k++) {
                cols.push({ value: '' });
            }
        }
        rows.push({ cols });
    }
    return rows;
}

/** Extract table name from caption or generate sequential name
 * @param {XMLDocument} xmlDoc The XML document
 * @param {Element} tableNode The table node
 * @param {number} tableIndex The index of the table (0-based)
 * @returns {string} Table name
 */
function extractTableName(xmlDoc, tableNode, tableIndex) {
    // Try to find caption in following paragraphs
    const body = xmlDoc.getElementsByTagName('w:body')[0];
    if (!body) return `Table ${tableIndex + 1}`;

    const bodyChildren = Array.from(body.childNodes);
    const tableIndexInBody = bodyChildren.indexOf(tableNode);

    // Check following paragraphs for caption (usually 1-2 paragraphs after table)
    for (let i = 1; i <= 3 && tableIndexInBody + i < bodyChildren.length; i++) {
        const nextNode = bodyChildren[tableIndexInBody + i];
        if (nextNode.nodeType === 1 && nextNode.nodeName === 'w:p') {
            const paraText = extractParagraphText(/** @type {Element} */(nextNode));
            // Look for "Table" followed by a number
            const tableMatch = paraText.match(/Table\s+(\d+)/i);
            if (tableMatch) {
                // Extract the full caption text
                return paraText.trim();
            }
            // Check for SEQ field codes
            const instrTexts = /** @type {Element} */ (nextNode).getElementsByTagName('w:instrText');
            for (let j = 0; j < instrTexts.length; j++) {
                const instr = instrTexts[j].childNodes[0]?.nodeValue || '';
                if (instr.includes('SEQ Table')) {
                    // Found SEQ field, get the full paragraph text as caption
                    return paraText.trim() || `Table ${tableIndex + 1}`;
                }
            }
        }
    }

    // Check preceding paragraphs
    for (let i = 1; i <= 2 && tableIndexInBody - i >= 0; i++) {
        const prevNode = bodyChildren[tableIndexInBody - i];
        if (prevNode.nodeType === 1 && prevNode.nodeName === 'w:p') {
            const paraText = extractParagraphText(/** @type {Element} */(prevNode));
            const tableMatch = paraText.match(/Table\s+(\d+)/i);
            if (tableMatch) {
                return paraText.trim();
            }
        }
    }

    // Generate sequential name if no caption found
    return `Table ${tableIndex + 1}`;
}

/** Parse legacy c: namespace chart data
 * @param {XMLDocument} chartDoc The parsed chart XML document
 * @param {OfficeParserConfig} [config] Config object for error handling
 * @returns {{ chartType: string, series: Array<{ categories: Array<string>, values: Array<number> }> } | null} Chart data or null if parsing fails
 */
function parseLegacyChart(chartDoc, config) {
    try {
        // Extract chart type from plotArea
        const plotArea = chartDoc.getElementsByTagName('c:plotArea')[0];
        if (!plotArea) return null;

        let chartType = 'unknown';
        const plotAreaChildren = plotArea.childNodes;
        const child = /** @type {Element} */ (plotAreaChildren[1]);

        const tagName = /** @type {Element} */ (child).tagName;
        if (tagName === 'c:barChart') chartType = 'bar';
        else if (tagName === 'c:lineChart') chartType = 'line';
        else if (tagName === 'c:pieChart') chartType = 'pie';
        else if (tagName === 'c:areaChart') chartType = 'area';
        else if (tagName === 'c:scatterChart') chartType = 'scatter';
        else if (tagName === 'c:bubbleChart') chartType = 'bubble';
        else if (tagName === 'c:doughnutChart') chartType = 'doughnut';
        else if (tagName === 'c:radarChart') chartType = 'radar';
        else if (tagName === 'c:surfaceChart') chartType = 'surface';
        else if (tagName === 'c:stockChart') chartType = 'stock';
        
        // Extract categories and values
        const series = [];
        const serNodes = child.getElementsByTagName('c:ser');
        for (let i = 0; i < serNodes.length; i++) {
            const serNode = serNodes[i];
            const catNode = serNode.getElementsByTagName('c:cat')[0];
            const valNode = serNode.getElementsByTagName('c:val')[0];
            const categories = [];
            const values = [];
            if (catNode) {
                const pts = catNode.getElementsByTagName('c:pt');
                for (let j = 0; j < pts.length; j++) {
                    const pt = pts[j];
                    const v = pt.getElementsByTagName('c:v')[0];
                    if (v && v.childNodes[0]) {
                        categories.push(v.childNodes[0].nodeValue || '');
                    }
                }
            }
            if (valNode) {
                const pts = valNode.getElementsByTagName('c:pt');
                for (let j = 0; j < pts.length; j++) {
                    const pt = pts[j];
                    const v = pt.getElementsByTagName('c:v')[0];
                    if (v && v.childNodes[0]) {
                        values.push(parseFloat(v.childNodes[0].nodeValue || '0'));
                    }
                }
            }
            series.push({ categories, values });
        }

        if (series.length === 0) {
            return null;
        }

        return { chartType, series };
    } catch (e) {
        if (config && config.outputErrorToConsole) {
            console.warn(`${ERRORHEADER}Error parsing legacy chart data:`, e.message);
        }
        return null;
    }
}

/** Parse chartex (cx:) namespace chart data with hierarchical structure
 * @param {XMLDocument} chartDoc The parsed chart XML document
 * @param {OfficeParserConfig} [config] Config object for error handling
 * @returns {{ chartType: string, series: Array<{ categories: Array<string | HierarchicalCategory>, values: Array<number> }> } | null} Chart data or null if parsing fails
 */
function parseChartexChart(chartDoc, config) {
    try {
        // Extract chart type from cx:series layoutId attribute
        const plotArea = chartDoc.getElementsByTagName('cx:plotArea')[0];
        if (!plotArea) return null;

        const plotAreaRegion = plotArea.getElementsByTagName('cx:plotAreaRegion')[0];
        if (!plotAreaRegion) return null;

        const seriesNodes = plotAreaRegion.getElementsByTagName('cx:series');
        if (seriesNodes.length === 0) return null;

        // Get chart type from first series layoutId
        const firstSeries = seriesNodes[0];
        const chartType = firstSeries.getAttribute('layoutId') || 'unknown';

        // Extract data from cx:chartData
        const chartData = chartDoc.getElementsByTagName('cx:chartData')[0];
        if (!chartData) return null;

        // Build a map of data by id
        const dataMap = {};
        const dataElements = chartData.getElementsByTagName('cx:data');
        for (let i = 0; i < dataElements.length; i++) {
            const dataEl = dataElements[i];
            const dataId = dataEl.getAttribute('id') || '0';
            dataMap[dataId] = dataEl;
        }

        const series = [];

        // Process each series
        for (let i = 0; i < seriesNodes.length; i++) {
            const serNode = seriesNodes[i];
            const dataIdNode = serNode.getElementsByTagName('cx:dataId')[0];
            const dataId = dataIdNode ? dataIdNode.getAttribute('val') || '0' : '0';
            
            const dataEl = dataMap[dataId];
            if (!dataEl) {
                if (config && config.outputErrorToConsole) {
                    console.warn(`${ERRORHEADER}Chart series dataId ${dataId} not found`);
                }
                continue;
            }

            // Parse categories
            const categories = [];
            
            // Check for string categories (cx:strDim type="cat")
            const strDim = dataEl.getElementsByTagName('cx:strDim')[0];
            if (strDim && strDim.getAttribute('type') === 'cat') {
                const lvlNodes = strDim.getElementsByTagName('cx:lvl');
                const levelCount = lvlNodes.length;
                
                if (levelCount > 0) {
                    // Get ptCount from first level
                    const firstLvl = lvlNodes[0];
                    const ptCount = parseInt(firstLvl.getAttribute('ptCount') || '0', 10);
                    
                    // Build hierarchical structure
                    // Note: In XML, first level is innermost (e.g., Leaf), last level is outermost (e.g., Branch)
                    // We want levels array to go from outermost to innermost, with value being the innermost
                    for (let idx = 0; idx < ptCount; idx++) {
                        const levelValues = [];
                        
                        // Extract from each level (innermost to outermost)
                        for (let lvlIdx = 0; lvlIdx < levelCount; lvlIdx++) {
                            const lvl = lvlNodes[lvlIdx];
                            const pt = Array.from(lvl.getElementsByTagName('cx:pt')).find(
                                p => parseInt(p.getAttribute('idx') || '-1', 10) === idx
                            );
                            
                            if (pt) {
                                const ptText = pt.textContent || '';
                                levelValues.push(ptText);
                            }
                        }
                        
                        if (levelValues.length > 0) {
                            // The first element is the innermost (value), rest are levels (outermost to innermost)
                            const value = levelValues[0];
                            const levels = levelValues.slice(1).reverse(); // Reverse to get outermost to innermost
                            categories.push({ levels, value });
                        }
                    }
                }
            }
            
            // Check for numeric categories (cx:numDim type="catVal")
            const numDimCat = Array.from(dataEl.getElementsByTagName('cx:numDim')).find(
                dim => dim.getAttribute('type') === 'catVal'
            );
            if (numDimCat && categories.length === 0) {
                const lvl = numDimCat.getElementsByTagName('cx:lvl')[0];
                if (lvl) {
                    const pts = lvl.getElementsByTagName('cx:pt');
                    for (let j = 0; j < pts.length; j++) {
                        const pt = pts[j];
                        const ptText = pt.textContent || '';
                        categories.push({ levels: [], value: ptText });
                    }
                }
            }

            // Parse values (cx:numDim type="size" or type="val")
            const values = [];
            const numDimVal = Array.from(dataEl.getElementsByTagName('cx:numDim')).find(
                dim => dim.getAttribute('type') === 'size' || dim.getAttribute('type') === 'val'
            );
            if (numDimVal) {
                const lvl = numDimVal.getElementsByTagName('cx:lvl')[0];
                if (lvl) {
                    const pts = lvl.getElementsByTagName('cx:pt');
                    for (let j = 0; j < pts.length; j++) {
                        const pt = pts[j];
                        const ptText = pt.textContent || '';
                        const numValue = parseFloat(ptText) || 0;
                        values.push(numValue);
                    }
                }
            }

            if (categories.length > 0 || values.length > 0) {
                series.push({ categories, values });
            }
        }

        if (series.length === 0) {
            return null;
        }

        return { chartType, series };
    } catch (e) {
        if (config && config.outputErrorToConsole) {
            console.warn(`${ERRORHEADER}Error parsing chartex chart data:`, e.message);
        }
        return null;
    }
}

/** Parse chart data from chart XML content
 * @param {string} chartPartContent The chart XML content
 * @param {OfficeParserConfig} [config] Config object for error handling
 * @returns {{ chartType: string, series: Array<{ categories: Array<string | HierarchicalCategory>, values: Array<number> }> } | null} Chart data or null if parsing fails
 */
function parseChartData(chartPartContent, config) {
    try {
        const chartDoc = parseString(chartPartContent);
        
        // Detect namespace by checking root element
        const rootElement = chartDoc.documentElement;
        const rootTagName = rootElement ? rootElement.tagName : '';
        
        // Check if it's chartex (cx:) namespace
        if (rootTagName === 'cx:chartSpace' || rootTagName.includes('chartSpace')) {
            // Check if it has cx: namespace by looking for cx:chartData
            const chartData = chartDoc.getElementsByTagName('cx:chartData')[0];
            if (chartData) {
                return parseChartexChart(chartDoc, config);
            }
        }
        
        // Default to legacy c: namespace
        return parseLegacyChart(chartDoc, config);
    } catch (e) {
        if (config && config.outputErrorToConsole) {
            console.warn(`${ERRORHEADER}Error parsing chart data:`, e.message);
        }
        return null;
    }
}

/** Extract charts from an embedded Excel file
 * @param {Buffer} excelBuffer The Excel file buffer (ZIP archive)
 * @param {OfficeParserConfig} config Config object for error handling
 * @returns {Promise<Array<{ chartType: string, series: Array<{ categories: Array<string | HierarchicalCategory>, values: Array<number> }> }>>} Array of chart data
 */
function extractChartsFromExcelEmbedding(excelBuffer, config) {
    const excelChartFileRegex = /xl\/charts\/chart\d+\.xml/g;

    return extractFiles(excelBuffer, x => !!x.match(excelChartFileRegex), false)
        .then(chartFiles => {
            const charts = [];

            for (const chartFile of chartFiles) {
                const chartContent = chartFile.content.toString();
                const chartData = parseChartData(chartContent, config);
                
                if (chartData) {
                    charts.push(chartData);
                }
            }

            return charts;
        })
        .catch(error => {
            if (config && config.outputErrorToConsole) {
                console.warn(`${ERRORHEADER}Error extracting charts from Excel embedding:`, error.message);
            }
            return [];
        });
}

/** Parse relationships from Word document.xml.rels file
 * @param {Buffer|string} relsFileContent Content of the rels file
 * @returns {Object.<string, string>} Map of relationship IDs to targets
 */
function parseWordRelationships(relsFileContent) {
    const rels = /** @type {Object.<string, string>} */ ({});
    const relsDoc = parseString(relsFileContent.toString());
    const relationships = relsDoc.getElementsByTagName('Relationship');
    for (let i = 0; i < relationships.length; i++) {
        const rel = relationships[i];
        const id = rel.getAttribute('Id');
        if (id) {
            rels[id] = rel.getAttribute('Target') || '';
        }
    }
    return rels;
}


/** Main async function for parsing text from word files
 * @param {string | Buffer}    file     File path or Buffers
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {void}
 */
function parseWord(file, callback, config) {
    /** The target content xml file for the docx file. */
    const mainContentFileRegex = /word\/document[\d+]?.xml/g;
    const footnotesFileRegex = /word\/footnotes[\d+]?.xml/g;
    const endnotesFileRegex = /word\/endnotes[\d+]?.xml/g;
    const relsFileRegex = /word\/_rels\/document.xml.rels/g;
    const mediaFileRegex = /word\/media\//g;
    const chartFileRegex = /word\/charts\/chart(Ex)?\d+\.xml/g;
    const chartRelsFileRegex = /word\/charts\/_rels\/chart(Ex)?\d+\.xml\.rels/g;
    const embeddingsFileRegex = /word\/embeddings\/.*\.xlsx/g;

    const filesToExtract = [mainContentFileRegex, footnotesFileRegex, endnotesFileRegex];
    // Always extract chart-related files
    filesToExtract.push(chartFileRegex, chartRelsFileRegex, embeddingsFileRegex);
    if (config.extractImages) {
        filesToExtract.push(relsFileRegex, mediaFileRegex);
    }
    if (config.extractCharts) {
        filesToExtract.push(relsFileRegex, chartRelsFileRegex);
    }

    // Extract text/XML files as strings
    const textFilesPromise = extractFiles(file, x => filesToExtract.some(fileRegex => x.match(fileRegex)), false);

    // Extract media files as buffers if extractImages is true
    const mediaFilesPromise = config.extractImages
        ? extractFiles(file, x => !!x.match(mediaFileRegex), true)
        : Promise.resolve([]);

    // Extract embedded Excel files as buffers (they are ZIP archives)
    const excelFilesPromise = extractFiles(file, x => !!x.match(embeddingsFileRegex), true);

    Promise.all([textFilesPromise, mediaFilesPromise, excelFilesPromise])
        .then(([textFiles, mediaFilesBuffers, excelFilesBuffers]) => {
            // Create a map to merge files, with buffer versions taking precedence
            const filesMap = new Map();
            textFiles.forEach(f => filesMap.set(f.path, f));
            mediaFilesBuffers.forEach(f => filesMap.set(f.path, f));
            excelFilesBuffers.forEach(f => filesMap.set(f.path, f));
            const files = Array.from(filesMap.values());

            return files;
        })
        .then(files => {
            
            // Verify if atleast the document xml file exists in the extracted files list.
            if (!files.some(file => file.path.match(mainContentFileRegex)))
                throw ERRORMSG.fileCorrupted(file);

            // Extract content files once for both image and text extraction
            const contentFiles = files
                .filter(file => file.path.match(mainContentFileRegex) || file.path.match(footnotesFileRegex) || file.path.match(endnotesFileRegex))
                .map(file => file.content.toString());

            const relsFile = files.find(file => file.path.match(relsFileRegex));
            const mediaFiles = files.filter(file => file.path.match(mediaFileRegex));
            const chartFiles = files.filter(file => file.path.match(chartFileRegex));
            const chartRelsFiles = files.filter(file => file.path.match(chartRelsFileRegex));
            const embeddingsFiles = files.filter(file => file.path.match(embeddingsFileRegex));
            const relationships = relsFile ? parseWordRelationships(relsFile.content) : {};

            
            // Parse chart relationships
            const chartRelationships = /** @type {Object.<string, Object.<string, string>>} */ ({});
            chartRelsFiles.forEach(chartRelsFile => {
                const chartPath = chartRelsFile.path.replace('/_rels', '').replace('.rels', '');
                chartRelationships[chartPath] = parseWordRelationships(chartRelsFile.content);
            });

            const delimiter = config.newlineDelimiter ?? '\n';
            const blocks = [];
            const textParts = [];
            const tables = [];
            const charts = [];
            const excelChartPromises = []; // Collect promises for async Excel chart extraction
            let tableIndex = 0;

            // ******************************** word xml files explanation ************************************
            // Structure of xmlContent of a word file has paragraphs in w:p tags.
            // Text content is in w:t tags, images are in w:drawing tags containing a:blip references.
            // Tables are in w:tbl tags, charts are embedded via w:object or chart parts.
            // We traverse w:body child nodes sequentially to maintain document order.
            // ************************************************************************************************
            contentFiles.forEach(xmlContent => {
                const xmlDoc = parseString(xmlContent);
                const body = xmlDoc.getElementsByTagName('w:body')[0];

                if (!body) return;

                // Traverse body children sequentially to maintain document order
                const bodyChildren = Array.from(body.childNodes);

                for (let i = 0; i < bodyChildren.length; i++) {
                    const node = bodyChildren[i];

                    // Skip non-element nodes
                    if (node.nodeType !== 1) continue;

                    const tagName = node.nodeName;

                    // Handle paragraphs (w:p)
                    if (tagName === 'w:p') {
                        const paraNode = /** @type {Element} */ (node);
                        // Extract text from paragraph
                        const textNodes = paraNode.getElementsByTagName('w:t');
                        const paragraphText = Array.from(textNodes)
                            .map(t => t.childNodes[0]?.nodeValue || '')
                            .join('');

                        if (paragraphText) {
                            blocks.push(createTextBlock(paragraphText));
                            textParts.push(paragraphText);
                        }

                        // Extract images from paragraph if requested
                        if (config.extractImages) {
                            const drawingNodes = paraNode.getElementsByTagName('w:drawing');
                            for (let j = 0; j < drawingNodes.length; j++) {
                                const blip = drawingNodes[j].getElementsByTagName('a:blip')[0];
                                if (blip) {
                                    const embedId = blip.getAttribute('r:embed');
                                    if (embedId && relationships[embedId]) {
                                        const imagePath = 'word/' + relationships[embedId].replace('../', '');
                                        const imageFile = mediaFiles.find(mf => mf.path === imagePath);
                                        if (imageFile) {
                                            const filename = imagePath.split('/').pop();
                                            const imageBuffer = Buffer.isBuffer(imageFile.content) ? imageFile.content : Buffer.from(imageFile.content);
                                            blocks.push(createImageBlock(imageBuffer, getMimeTypeFromFilename(filename), filename));
                                            textParts.push(`<image ${filename}/>`);
                                        } else if (config.outputErrorToConsole) {
                                            console.warn(`${ERRORHEADER}Image referenced but not found: ${imagePath}`);
                                        }
                                    }
                                }
                            }
                        }

                        if (config.extractCharts) {
                            const chartNodes = paraNode.getElementsByTagName('c:chart');
                            for (let k = 0; k < chartNodes.length; k++) {
                                const chartNode = chartNodes[k];
                                const chartId = chartNode.getAttribute('r:id');

                                if (chartId && relationships[chartId]) {
                                    const target = relationships[chartId];

                                    // Check if it's a chart part
                                    if (target.includes('charts/chart')) {
                                        const chartPath = 'word/' + target.replace('../', '');
                                        const chartFile = chartFiles.find(cf => cf.path === chartPath);

                                        if (chartFile) {
                                            const chartData = parseChartData(chartFile.content.toString(), config);
                                            
                                            if (chartData) {
                                                charts.push(chartData);
                                                blocks.push(createChartBlock(chartData.chartType, chartData.series));
                                                textParts.push(`<chart ${chartData.chartType}/>`);

                                            }
                                        } else if (config.outputErrorToConsole) {
                                            console.warn(`${ERRORHEADER}Chart referenced but not found: ${chartPath}`);
                                        }
                                    }
                                    // Check if it's an embedded Excel file (which may contain charts)
                                    else if (target.includes('embeddings/') && target.endsWith('.xlsx')) {
                                        const embedPath = 'word/' + target.replace('../', '');
                                        const embedFile = embeddingsFiles.find(ef => ef.path === embedPath);

                                        if (embedFile) {
                                            // Ensure we have a Buffer (not string)
                                            const excelBuffer = Buffer.isBuffer(embedFile.content)
                                                ? embedFile.content
                                                : Buffer.from(embedFile.content);

                                            // Extract charts from Excel file asynchronously
                                            const chartPromise = extractChartsFromExcelEmbedding(excelBuffer, config)
                                                .then(excelCharts => {
                                                    // Add each chart found
                                                    excelCharts.forEach(chartData => {
                                                        charts.push(chartData);
                                                        blocks.push(createChartBlock(chartData.chartType, chartData.series));
                                                        textParts.push(`<chart ${chartData.chartType}/>`);
                                                    });
                                                })
                                                .catch(error => {
                                                    if (config.outputErrorToConsole) {
                                                        console.warn(`${ERRORHEADER}Error extracting charts from embedded Excel:`, error.message);
                                                    }
                                                });

                                            excelChartPromises.push(chartPromise);
                                        }
                                    }
                                }
                            }

                            const chartxNodes = paraNode.getElementsByTagName('cx:chart');
                            for (let k = 0; k < chartxNodes.length; k++) {
                                const chartNode = chartxNodes[k];
                                const chartId = chartNode.getAttribute('r:id');

                                if (chartId && relationships[chartId]) {
                                    const target = relationships[chartId];

                                    // Check if it's a chart part
                                    if (target.includes('charts/chart')) {
                                        const chartPath = 'word/' + target.replace('../', '');
                                        const chartFile = chartFiles.find(cf => cf.path === chartPath);

                                        if (chartFile) {
                                            const chartData = parseChartData(chartFile.content.toString(), config);
                                            
                                            if (chartData) {
                                                charts.push(chartData);
                                                blocks.push(createChartBlock(chartData.chartType, chartData.series));
                                                textParts.push(`<chart ${chartData.chartType}/>`);

                                            }
                                        } else if (config.outputErrorToConsole) {
                                            console.warn(`${ERRORHEADER}Chart referenced but not found: ${chartPath}`);
                                        }
                                    }
                                    // Check if it's an embedded Excel file (which may contain charts)
                                    else if (target.includes('embeddings/') && target.endsWith('.xlsx')) {
                                        const embedPath = 'word/' + target.replace('../', '');
                                        const embedFile = embeddingsFiles.find(ef => ef.path === embedPath);

                                        if (embedFile) {
                                            // Ensure we have a Buffer (not string)
                                            const excelBuffer = Buffer.isBuffer(embedFile.content)
                                                ? embedFile.content
                                                : Buffer.from(embedFile.content);

                                            // Extract charts from Excel file asynchronously
                                            const chartPromise = extractChartsFromExcelEmbedding(excelBuffer, config)
                                                .then(excelCharts => {
                                                    // Add each chart found
                                                    excelCharts.forEach(chartData => {
                                                        charts.push(chartData);
                                                        blocks.push(createChartBlock(chartData.chartType, chartData.series));
                                                        textParts.push(`<chart ${chartData.chartType}/>`);
                                                    });
                                                })
                                                .catch(error => {
                                                    if (config.outputErrorToConsole) {
                                                        console.warn(`${ERRORHEADER}Error extracting charts from embedded Excel:`, error.message);
                                                    }
                                                });

                                            excelChartPromises.push(chartPromise);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    // Handle tables (w:tbl)
                    else if (tagName === 'w:tbl') {
                        const tableNode = /** @type {Element} */ (node);
                        const rows = extractTableRows(tableNode);
                        const name = extractTableName(xmlDoc, tableNode, tableIndex);
                        const tableData = { name, rows };

                        tables.push(tableData);
                        blocks.push(createTableBlock(name, rows));
                        textParts.push(`<table ${name}/>`);
                        tableIndex++;
                    }

                }
            });

            // Wait for all Excel chart extraction promises to complete
            return Promise.all(excelChartPromises).then(() => {
                console.log("Chart data:", JSON.stringify(charts))
                const text = textParts.join(delimiter);
                callback({ text, blocks, tables, charts }, undefined);
            });
        })
        .catch(e => callback(undefined, e));
}

/** Main function for parsing text from PowerPoint files
 * @param {string | Buffer}    file     File path or Buffers
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {void}
 */
function parsePowerPoint(file, callback, config) {
    // Files regex that hold our content of interest
    const allFilesRegex = /ppt\/(notesSlides|slides)\/(notesSlide|slide)\d+.xml/g;
    const slidesRegex = /ppt\/slides\/slide\d+.xml/g;
    const slideNumberRegex = /lide(\d+)\.xml/;
    const relsFileRegex = /ppt\/slides\/_rels\/slide\d+.xml.rels/g;
    const mediaFileRegex = /ppt\/media\//g;

    const filesToExtract = [config.ignoreNotes ? slidesRegex : allFilesRegex];
    if (config.extractImages) {
        filesToExtract.push(relsFileRegex, mediaFileRegex);
    }

    extractFiles(file, x => filesToExtract.some(fileRegex => x.match(fileRegex)), config.extractImages)
        .then(files => {
            // Sort files by slide number and their notes (if any).
            files.sort((a, b) => {
                const matchedANumber = parseInt(a.path.match(slideNumberRegex)?.at(1), 10);
                const matchedBNumber = parseInt(b.path.match(slideNumberRegex)?.at(1), 10);

                const aNumber = isNaN(matchedANumber) ? Infinity : matchedANumber;
                const bNumber = isNaN(matchedBNumber) ? Infinity : matchedBNumber;

                return aNumber - bNumber || Number(a.path.includes('notes')) - Number(b.path.includes('notes'));
            });

            // Verify if atleast the slides xml files exist in the extracted files list.
            if (files.length == 0 || !files.map(file => file.path).some(filename => filename.match(slidesRegex)))
                throw ERRORMSG.fileCorrupted(file);

            // Check if any sorting is required.
            if (!config.ignoreNotes && config.putNotesAtLast)
                // Sort files according to previous order of taking text out of ppt/slides followed by ppt/notesSlides
                // For this we are looking at the index of notes which results in -1 in the main slide file and exists at a certain index in notes file names.
                files.sort((a, b) => a.path.indexOf('notes') - b.path.indexOf('notes'));

            return files;
        })
        // ******************************** powerpoint xml files explanation ************************************
        // Structure of xmlContent of a powerpoint file is simple.
        // There are multiple xml files for each slide and correspondingly their notesSlide files.
        // All text nodes are within a:t tags and each of the text nodes that belong in one paragraph are clubbed together within a a:p tag.
        // Images are referenced via a:blip tags within p:pic elements.
        // For blocks, we process each slide and extract text paragraphs and images in document order.
        // ******************************************************************************************************
        .then(files => {
            const contentFiles = files.filter(file => file.path.match(allFilesRegex));
            const relsFiles = files.filter(file => file.path.match(relsFileRegex));
            const mediaFiles = files.filter(file => file.path.match(mediaFileRegex));

            // Parse relationship files and map them to their corresponding slide
            // Each slide has its own rels file with relationship IDs scoped to that slide
            const slideRelationships = {};
            relsFiles.forEach(relsFile => {
                const slidePath = relsFile.path.replace('/_rels', '').replace('.rels', '');
                slideRelationships[slidePath] = parseWordRelationships(relsFile.content);
            });

            const delimiter = config.newlineDelimiter ?? '\n';
            const blocks = [];
            const textParts = [];

            contentFiles.forEach(file => {
                const xmlDoc = parseString(file.content.toString());
                const slideRels = slideRelationships[file.path] || {};

                // Extract text from all paragraphs in this slide
                const slideText = extractTextFromXmlParagraphs(xmlDoc, 'a:p', 'a:t', delimiter);
                // Always add to textParts to preserve slide boundaries (empty slides produce newlines)
                textParts.push(slideText);
                if (slideText) {
                    blocks.push(createTextBlock(slideText));
                }

                // Extract images from this slide if requested
                // Note: Images appear after text per slide. For true document-order interleaving
                // within a slide, a deeper XML tree traversal would be needed.
                if (config.extractImages) {
                    const blipNodes = xmlDoc.getElementsByTagName('a:blip');
                    for (let i = 0; i < blipNodes.length; i++) {
                        const blip = blipNodes[i];
                        const embedId = blip.getAttribute('r:embed');
                        if (embedId && slideRels[embedId]) {
                            const imagePath = 'ppt/' + slideRels[embedId].replace('../', '');
                            const imageFile = mediaFiles.find(mf => mf.path === imagePath);
                            if (imageFile) {
                                const filename = imagePath.split('/').pop();
                                blocks.push(createImageBlock(imageFile.content, getMimeTypeFromFilename(filename), filename));
                                // Add image placeholder to text
                                textParts.push(`<image ${filename}/>`);
                            } else if (config.outputErrorToConsole) {
                                console.warn(`${ERRORHEADER}Image referenced but not found: ${imagePath}`);
                            }
                        }
                    }
                }
            });

            const text = textParts.join(delimiter);

            callback({ text, blocks }, undefined);
        })
        .catch(e => callback(undefined, e));
}

/** Main function for parsing text from Excel files
 * @param {string | Buffer}    file     File path or Buffers
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {void}
 */
function parseExcel(file, callback, config) {
    // Files regex that hold our content of interest
    const sheetsRegex = /xl\/worksheets\/sheet\d+.xml/g;
    const drawingsRegex = /xl\/drawings\/drawing\d+.xml/g;
    const chartsRegex = /xl\/charts\/chart\d+.xml/g;
    const stringsFilePath = 'xl/sharedStrings.xml';

    extractFiles(file, x => [sheetsRegex, drawingsRegex, chartsRegex].some(fileRegex => x.match(fileRegex)) || x == stringsFilePath)
        .then(files => {
            // Verify if atleast the slides xml files exist in the extracted files list.
            if (files.length == 0 || !files.map(file => file.path).some(filename => filename.match(sheetsRegex)))
                throw ERRORMSG.fileCorrupted(file);

            return {
                sheetFiles: files.filter(file => file.path.match(sheetsRegex)).map(file => file.content),
                drawingFiles: files.filter(file => file.path.match(drawingsRegex)).map(file => file.content),
                chartFiles: files.filter(file => file.path.match(chartsRegex)).map(file => file.content),
                sharedStringsFile: files.filter(file => file.path == stringsFilePath).map(file => file.content)[0],
            };
        })
        // ********************************** excel xml files explanation ***************************************
        // Structure of xmlContent of an excel file is a bit complex.
        // We usually have a sharedStrings.xml file which has strings inside t tags
        // However, this file is not necessary to be present. It is sometimes absent if the file has no shared strings indices represented in v nodes.
        // Each sheet has an individual sheet xml file which has numbers in v tags (probably value) inside c tags (probably cell)
        // Each value of v tag is to be used as it is if the "t" attribute (probably type) of c tag is not "s" (probably shared string)
        // If the "t" attribute of c tag is "s", then we use the value to select value from sharedStrings array with the value as its index.
        // However, if the "t" attribute of c tag is "inlineStr", strings can be inline inside "is"(probably inside String) > "t".
        // We extract either the inline strings or use the value to get numbers of text from shared strings.
        // Drawing files contain all text for each drawing and have text nodes in a:t and paragraph nodes in a:p.
        // ******************************************************************************************************
        .then(xmlContentFilesObject => {
            /** Store all the text content to respond */
            const responseText = [];

            /** Function to check if the given c node is a valid inline string node. */
            function isValidInlineStringCNode(cNode) {
                // Initial check to see if the passed node is a cNode
                if (cNode.tagName.toLowerCase() != 'c')
                    return false;
                if (cNode.getAttribute('t') != 'inlineStr')
                    return false;
                const childNodesNamedIs = cNode.getElementsByTagName('is');
                if (childNodesNamedIs.length != 1)
                    return false;
                const childNodesNamedT = childNodesNamedIs[0].getElementsByTagName('t');
                if (childNodesNamedT.length != 1)
                    return false;
                return childNodesNamedT[0].childNodes[0] && childNodesNamedT[0].childNodes[0].nodeValue != '';
            }

            /** Function to check if the given c node has a valid v node */
            function hasValidVNodeInCNode(cNode) {
                return cNode.getElementsByTagName('v')[0]
                    && cNode.getElementsByTagName('v')[0].childNodes[0]
                    && cNode.getElementsByTagName('v')[0].childNodes[0].nodeValue != '';
            }

            /** Find text nodes with t tags in sharedStrings.xml file. If the sharedStringsFile is not present, we return an empty array. */
            const sharedStringsXmlSiNodesList = xmlContentFilesObject.sharedStringsFile != undefined
                ? parseString(xmlContentFilesObject.sharedStringsFile).getElementsByTagName('si')
                : [];

            /** Create shared string array. This will be used as a map to get strings from within sheet files. */
            const sharedStrings = Array.from(sharedStringsXmlSiNodesList)
                .map(siNode => {
                    // Concatenate all <t> nodes within the <si> node
                    return Array.from(siNode.getElementsByTagName('t'))
                        .map(tNode => tNode.childNodes[0]?.nodeValue ?? '') // Extract text content from each <t> node
                        .join(''); // Combine all <t> node text into a single string
                });

            // Parse Sheet files
            xmlContentFilesObject.sheetFiles.forEach(sheetXmlContent => {
                /** Find text nodes with c tags in sharedStrings xml file */
                const sheetsXmlCNodesList = parseString(sheetXmlContent).getElementsByTagName('c');
                // Traverse through the nodes list and fill responseText with either the number value in its v node or find a mapped string from sharedStrings or an inline string.
                responseText.push(
                    Array.from(sheetsXmlCNodesList)
                        // Filter out invalid c nodes
                        .filter(cNode => isValidInlineStringCNode(cNode) || hasValidVNodeInCNode(cNode))
                        .map(cNode => {
                            // Processing if this is a valid inline string c node.
                            if (isValidInlineStringCNode(cNode))
                                return cNode.getElementsByTagName('is')[0].getElementsByTagName('t')[0].childNodes[0].nodeValue;

                            // Processing if this c node has a valid v node.
                            if (hasValidVNodeInCNode(cNode)) {
                                /** Flag whether this node's value represents an index in the shared string array */
                                const isIndexInSharedStrings = cNode.getAttribute('t') == 's';
                                /** Find value nodes represented by v tags */
                                const value = cNode.getElementsByTagName('v')[0].childNodes[0].nodeValue;
                                const valueAsIndex = Number(value);
                                // Validate text
                                if (isIndexInSharedStrings && (valueAsIndex != parseInt(value, 10) || valueAsIndex >= sharedStrings.length))
                                    throw ERRORMSG.fileCorrupted(file);

                                return isIndexInSharedStrings
                                    ? sharedStrings[valueAsIndex]
                                    : value;
                            }
                            handleError(`Invalid c node found in sheet xml content: ${cNode}`, callback, config.outputErrorToConsole);
                            return '';
                        })
                        .join(config.newlineDelimiter ?? '\n')
                );
            });

            // Parse Drawing files
            xmlContentFilesObject.drawingFiles.forEach(drawingXmlContent => {
                const xmlDoc = parseString(drawingXmlContent);
                const text = extractTextFromXmlParagraphs(xmlDoc, 'a:p', 'a:t', config.newlineDelimiter ?? '\n');
                responseText.push(text);
            });

            // Parse Chart files
            xmlContentFilesObject.chartFiles.forEach(chartXmlContent => {
                /** Find text nodes with c:v tags */
                const chartsXmlCVNodesList = parseString(chartXmlContent).getElementsByTagName('c:v');
                /** Store all the text content to respond */
                responseText.push(
                    Array.from(chartsXmlCVNodesList)
                        .filter(cVNode => cVNode.childNodes[0] && cVNode.childNodes[0].nodeValue)
                        .map(cVNode => cVNode.childNodes[0].nodeValue)
                        .join(config.newlineDelimiter ?? '\n')
                );
            });

            const text = responseText.join(config.newlineDelimiter ?? '\n');
            const blocks = text ? [createTextBlock(text)] : [];

            callback({ text, blocks }, undefined);
        })
        .catch(e => callback(undefined, e));
}

/** Main function for parsing text from open office files
 * @param {string | Buffer}    file     File path or Buffers
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {void}
 */
function parseOpenOffice(file, callback, config) {
    /** The target content xml file for the openoffice file. */
    const mainContentFilePath = 'content.xml';
    const objectContentFilesRegex = /Object \d+\/content.xml/g;
    const mediaFileRegex = /Pictures\//g;
    const objectReplacementsRegex = /ObjectReplacements\/Object \d+/g;

    const filesToExtract = [mainContentFilePath, objectContentFilesRegex];
    if (config.extractImages) {
        filesToExtract.push(mediaFileRegex, objectReplacementsRegex);
    }

    extractFiles(file, x => filesToExtract.some(filePattern => {
        if (typeof filePattern === 'string') return x == filePattern;
        return !!x.match(filePattern);
    }), config.extractImages)
        .then(files => {
            // Verify if atleast the content xml file exists in the extracted files list.
            if (!files.map(file => file.path).includes(mainContentFilePath))
                throw ERRORMSG.fileCorrupted(file);

            const result = {
                mainContentFile: files.filter(file => file.path == mainContentFilePath).map(file => file.content.toString())[0],
                objectContentFiles: files.filter(file => file.path.match(objectContentFilesRegex)).map(file => file.content.toString()),
                mediaFiles: config.extractImages ? files.filter(file => file.path.match(mediaFileRegex) || file.path.match(objectReplacementsRegex)) : []
            };

            return result;
        })
        // ********************************** openoffice xml files explanation **********************************
        // Structure of xmlContent of openoffice files is simple.
        // All text nodes are within text:h and text:p tags with all kinds of formatting within nested tags.
        // All text in these tags are separated by new line delimiters.
        // Objects like charts in ods files are in Object d+/content.xml with the same way as above.
        // Images are referenced via draw:image tags with xlink:href attributes.
        // ******************************************************************************************************
        .then(xmlContentFilesObject => {
            /** Store all the notes text content to respond */
            const notesText = [];
            /** Store all the text content to respond */
            let responseText = [];
            /** Blocks array for ordered content */
            const blocks = [];

            /** List of allowed text tags */
            const allowedTextTags = ['text:p', 'text:h'];
            /** List of notes tags */
            const notesTag = 'presentation:notes';

            /** Main dfs traversal function that goes from one node to its children and returns the value out. */
            function extractAllTextsFromNode(root) {
                const xmlTextArray = [];
                for (let i = 0; i < root.childNodes.length; i++)
                    traversal(root.childNodes[i], xmlTextArray, true);
                return xmlTextArray.join('');
            }

            /** Traversal function that gets recursive calling. */
            function traversal(node, xmlTextArray, isFirstRecursion) {
                if (!node.childNodes || node.childNodes.length == 0) {
                    if (node.parentNode.tagName.indexOf('text') == 0 && node.nodeValue) {
                        // If the corresponding value is of type float, we take the value from office:value attribute.
                        // However, it is not on the parentNode but rather grandparentNode.
                        const value = node.parentNode.parentNode?.getAttribute('office:value-type') == 'float'
                            ? Number(node.parentNode.parentNode.getAttribute('office:value'))
                            : node.nodeValue;

                        if (isNotesNode(node.parentNode) && (config.putNotesAtLast || config.ignoreNotes)) {
                            notesText.push(value);
                            if (allowedTextTags.includes(node.parentNode.tagName) && !isFirstRecursion)
                                notesText.push(config.newlineDelimiter ?? '\n');
                        } else {
                            xmlTextArray.push(value);
                            if (allowedTextTags.includes(node.parentNode.tagName) && !isFirstRecursion)
                                xmlTextArray.push(config.newlineDelimiter ?? '\n');
                        }
                    }
                    return;
                }

                for (let i = 0; i < node.childNodes.length; i++)
                    traversal(node.childNodes[i], xmlTextArray, false);
            }

            /** Checks if the given node has an ancestor which is a notes tag. We use this information to put the notes in the response text and its position. */
            function isNotesNode(node) {
                if (node.tagName == notesTag)
                    return true;
                if (node.parentNode)
                    return isNotesNode(node.parentNode);
                return false;
            }

            /** Checks if the given node has an ancestor which is also an allowed text tag. In that case, we ignore the child text tag. */
            function isInvalidTextNode(node) {
                if (allowedTextTags.includes(node.tagName))
                    return true;
                if (node.parentNode)
                    return isInvalidTextNode(node.parentNode);
                return false;
            }

            /** The xml string parsed as xml array */
            const xmlContentArray = [xmlContentFilesObject.mainContentFile, ...xmlContentFilesObject.objectContentFiles]
                .filter(content => content)  // Filter out undefined content
                .map(xmlContent => {
                    try {
                        return parseString(xmlContent);
                    } catch (e) {
                        if (config.outputErrorToConsole) {
                            console.error(`${ERRORHEADER}Error parsing XML content:`, e.message);
                        }
                        return null;
                    }
                })
                .filter(doc => doc);  // Filter out any failed parses

            // Iterate over each xmlContent and extract text and images in document order
            xmlContentArray.forEach(xmlContent => {
                /** Find all nodes to process in document order */
                const allNodes = Array.from(xmlContent.getElementsByTagName('*'));

                allNodes.forEach(node => {
                    // Handle text paragraphs
                    if (allowedTextTags.includes(node.tagName) && !isInvalidTextNode(node.parentNode)) {
                        const textContent = extractAllTextsFromNode(node);
                        if (textContent) {
                            blocks.push(createTextBlock(textContent));
                            if (!isNotesNode(node) || (!config.ignoreNotes && !config.putNotesAtLast)) {
                                responseText.push(textContent);
                            }
                        }
                    }

                    // Handle images
                    if (node.tagName === 'draw:image' && config.extractImages) {
                        const href = node.getAttribute('xlink:href');
                        if (href) {
                            const normalizedHref = href.startsWith('./') ? href.substring(2) : href;
                            const imageFile = xmlContentFilesObject.mediaFiles.find(mf => mf.path === normalizedHref);
                            if (imageFile) {
                                const filename = href.split('/').pop();
                                blocks.push(createImageBlock(imageFile.content, getMimeTypeFromFilename(filename), filename));
                                // Add image placeholder to text
                                responseText.push(`<image ${filename}/>`);
                            } else if (config.outputErrorToConsole) {
                                console.warn(`${ERRORHEADER}Image referenced but not found: ${href}`);
                            }
                        }
                    }
                });
            });

            // Add notes text at the end if the user config says so.
            if (!config.ignoreNotes && config.putNotesAtLast)
                responseText = [...responseText, ...notesText];

            const text = responseText.join(config.newlineDelimiter ?? '\n');

            // Respond by calling the Callback function.
            callback({ text, blocks }, undefined);
        })
        .catch(e => callback(undefined, e));
}


/** Main function for parsing text from pdf files
 * @param {string | Buffer}    file     File path or Buffers
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {Promise<void>}
 */
async function parsePdf(file, callback, config) {
    const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
    const delimiter = config.newlineDelimiter ?? '\n';

    pdfjs.getDocument(file instanceof Buffer ? new Uint8Array(file) : file).promise
        .then(async document => {
            const blocks = [];
            const pagePromises = Array.from({ length: document.numPages }, (_, index) => document.getPage(index + 1));
            const pages = await Promise.all(pagePromises);

            // First, get all text content for the full text output (backwards compatibility)
            // This uses the original logic that flattens all items across pages
            const textContentArray = [];
            for (const page of pages) {
                const textContent = await page.getTextContent();
                textContentArray.push(textContent);
            }

            // Build full text using original algorithm (line breaks based on transform[5])
            const text = textContentArray
                .map(textContent => textContent.items)
                .flat()
                .reduce((a, v) => (
                    'str' in v && v.str != ''
                        ? {
                            text: a.text + (v.transform[5] != a.transform5 ? delimiter : '') + v.str,
                            transform5: v.transform[5]
                        } : {
                            text: a.text,
                            transform5: a.transform5
                        }
                ), { text: '', transform5: undefined }).text;

            // Build blocks per page for document order
            const imagePlaceholders = [];
            let imageIndex = 0;

            for (let i = 0; i < pages.length; i++) {
                const page = pages[i];
                const textContent = textContentArray[i];

                // Extract text block for this page
                const pageText = textContent.items
                    .reduce((a, v) => (
                        'str' in v && v.str != ''
                            ? {
                                text: a.text + (v.transform[5] != a.transform5 ? delimiter : '') + v.str,
                                transform5: v.transform[5]
                            } : {
                                text: a.text,
                                transform5: a.transform5
                            }
                    ), { text: '', transform5: undefined }).text;

                if (pageText) {
                    blocks.push(createTextBlock(pageText));
                }

                // Extract images from this page if requested
                if (config.extractImages) {
                    const operatorList = await page.getOperatorList(true);
                    const imageOps = operatorList.fnArray
                        .map((fn, j) => fn === pdfjs.OPS.paintImageXObject ? operatorList.argsArray[j][0] : null)
                        .filter(op => op);

                    for (const op of imageOps) {
                        const image = await getPdfImageResource(page, op);
                        if (image) {
                            const mimeType = image.kind === pdfjs.ImageKind.JPEG ? 'image/jpeg' : 'image/png';
                            const extension = mimeType === 'image/jpeg' ? 'jpg' : 'png';
                            const filename = `image_${++imageIndex}.${extension}`;
                            blocks.push(createImageBlock(Buffer.from(image.data), mimeType, filename));
                            // Collect image placeholder
                            imagePlaceholders.push(`<image ${filename}/>`);
                        } else if (config.outputErrorToConsole) {
                            console.warn(`${ERRORHEADER}PDF image resource not found: ${op}`);
                        }
                    }
                }
            }

            // Append image placeholders at the end of text (PDF images don't have exact positions)
            const finalText = imagePlaceholders.length > 0
                ? text + delimiter + imagePlaceholders.join(delimiter)
                : text;
            callback({ text: finalText, blocks }, undefined);
        })
        .catch(e => callback(undefined, e));
}

/** Main async function with callback to execute parseOffice for supported files
 * @param {string | Buffer | ArrayBuffer} srcFile      File path or file buffers or Javascript ArrayBuffer
 * @param {function}                      callback     Callback function that returns value or error
 * @param {OfficeParserConfig}            [config={}]  [OPTIONAL]: Config Object for officeParser
 * @returns {void}
 */
function parseOffice(srcFile, callback, config = {}) {
    // Make a clone of the config with default values such that none of the config flags are undefined.
    /** @type {OfficeParserConfig} */
    const internalConfig = {
        ignoreNotes: false,
        newlineDelimiter: '\n',
        putNotesAtLast: false,
        outputErrorToConsole: false,
        extractImages: false,
        ...config
    };

    // Our internal code can process regular node Buffers or file path.
    // So, if the src file was presented as ArrayBuffers, we create Buffers from them.
    const file = srcFile instanceof ArrayBuffer ? Buffer.from(srcFile)
        : srcFile;

    /**
     * Prepare file for processing
     * @type {Promise<{ file:string | Buffer, ext: string}>}
     */
    const filePreparedPromise = new Promise((res, rej) => {
        // Check if buffer
        if (Buffer.isBuffer(file))
            // Guess file type from buffer
            return fileType.fromBuffer(file)
                .then(data => res({ file: file, ext: data.ext.toLowerCase() }))
                .catch(() => rej(ERRORMSG.improperBuffers));
        else if (typeof file === 'string') {
            // Not buffers but real file path.
            // Check if file exists
            if (!fs.existsSync(file))
                throw ERRORMSG.fileDoesNotExist(file);

            // resolve promise
            res({ file: file, ext: file.split('.').pop() });
        } else
            rej(ERRORMSG.invalidInput);
    });

    // Process filePreparedPromise resolution.
    filePreparedPromise
        .then(({ file, ext }) => {
            // Switch between parsing functions depending on extension.
            switch (ext) {
                case 'docx':
                    parseWord(file, internalCallback, internalConfig);
                    break;
                case 'pptx':
                    parsePowerPoint(file, internalCallback, internalConfig);
                    break;
                case 'xlsx':
                    parseExcel(file, internalCallback, internalConfig);
                    break;
                case 'odt':
                case 'odp':
                case 'ods':
                    parseOpenOffice(file, internalCallback, internalConfig);
                    break;
                case 'pdf':
                    parsePdf(file, internalCallback, internalConfig);
                    break;

                default:
                    internalCallback(undefined, ERRORMSG.extensionUnsupported(ext));  // Call the internalCallback function which removes the temp files if required.
            }

            /** Internal callback function that calls the user's callback function passed in argument and removes the temp files if required */
            function internalCallback(data, err) {
                // Check if there is an error. Throw if there is an error.
                if (err)
                    return handleError(err, callback, internalConfig.outputErrorToConsole);

                if (typeof data === 'object' && data !== null) {
                    if (!data.blocks) {
                        data.blocks = data.text ? [{ type: 'text', content: data.text }] : [];
                    }
                    if (!data.text) {
                        data.text = data.blocks.filter(b => b.type === 'text').map(b => b.content).join(internalConfig.newlineDelimiter ?? '\n');
                    }
                    // Ensure tables and charts are always present for consistent API (Word returns these; other formats get [])
                    if (data.tables === undefined) data.tables = [];
                    if (data.charts === undefined) data.charts = [];
                    callback(data, undefined);
                } else {
                    callback({ text: data, blocks: data ? [{ type: 'text', content: data }] : [], tables: [], charts: [] }, undefined);
                }
            }
        })
        .catch(error => handleError(error, callback, internalConfig.outputErrorToConsole));
}

/** Main async function that can be used with await to execute parseOffice. Or it can be used with promises.
 * @param {string | Buffer | ArrayBuffer} srcFile     File path or file buffers or Javascript ArrayBuffer
 * @param {OfficeParserConfig}            [config={}] [OPTIONAL]: Config Object for officeParser
 * @returns {Promise<string>}
 */
function parseOfficeAsync(srcFile, config = {}) {
    return new Promise((res, rej) => {
        parseOffice(srcFile, function (data, err) {
            if (err)
                return rej(err);
            return res(data);
        }, config);
    });
}

/** Extract specific files from either a ZIP file buffer or file path based on a filter function.
 * @param {Buffer|string}          zipInput ZIP file input, either a Buffer or a file path (string).
 * @param {(x: string) => boolean} filterFn A function that receives the entry object and returns true if the file should be extracted.
 * @param {boolean}                [asBuffer=false] Whether to extract the file as a buffer or string.
 * @returns {Promise<{ path: string, content: string|Buffer }[]>} Resolves to an array of objects
 */
function extractFiles(zipInput, filterFn, asBuffer = false) {
    return new Promise((res, rej) => {
        /** Processes zip file and resolves with the path of file and their content.
         * @param {yauzl.ZipFile} zipfile
         */
        const processZipfile = (zipfile) => {
            /** @type {{ path: string, content: string|Buffer }[]} */
            const extractedFiles = [];
            zipfile.readEntry();

            /** @param {yauzl.Entry} entry  */
            function processEntry(entry) {
                // Use the filter function to determine if the file should be extracted
                if (filterFn(entry.fileName)) {
                    zipfile.openReadStream(entry, (err, readStream) => {
                        if (err)
                            return rej(err);

                        // Use concat-stream to collect the data into a single Buffer
                        readStream.pipe(concat(data => {
                            extractedFiles.push({
                                path: entry.fileName,
                                content: asBuffer ? data : data.toString()
                            });
                            zipfile.readEntry(); // Continue reading entries
                        }));
                    });
                } else
                    zipfile.readEntry(); // Skip entries that don't match the filter
            }

            zipfile.on('entry', processEntry);
            zipfile.on('end', () => res(extractedFiles));
            zipfile.on('error', rej);
        };

        // Determine whether the input is a buffer or file path
        if (Buffer.isBuffer(zipInput)) {
            // Process ZIP from Buffer
            yauzl.fromBuffer(zipInput, { lazyEntries: true }, (err, zipfile) => {
                if (err) return rej(err);
                processZipfile(zipfile);
            });
        } else if (typeof zipInput === 'string') {
            // Process ZIP from File Path
            yauzl.open(zipInput, { lazyEntries: true }, (err, zipfile) => {
                if (err) return rej(err);
                processZipfile(zipfile);
            });
        } else
            rej(ERRORMSG.invalidInput);
    });
}

/** Handle error by logging it to console if permitted by the config.
 * And after that, trigger the callback function with the error value.
 * @param {string}   error                Error text
 * @param {function} callback             Callback function provided by the caller
 * @param {boolean}  outputErrorToConsole Flag to log error to console.
 * @returns {void}
 */
function handleError(error, callback, outputErrorToConsole) {
    if (error && outputErrorToConsole)
        console.error(ERRORHEADER + error);

    callback(undefined, new Error(ERRORHEADER + error));
}


// Export functions
module.exports.parseOffice = parseOffice;
module.exports.parseOfficeAsync = parseOfficeAsync;


// Run this library on CLI
if ((typeof process.argv[0] === 'string' && (process.argv[0].split('/').pop() == 'node' || process.argv[0].split('/').pop() == 'npx')) &&
    (typeof process.argv[1] === 'string' && (process.argv[1].split('/').pop() == 'officeParser.js' || process.argv[1].split('/').pop().toLowerCase() == 'officeparser'))) {

    // Extract arguments after the script is called
    /** Stores the list of arguments for this CLI call
     * @type {string[]}
     */
    const args = process.argv.slice(2);
    /** Stores the file argument for this CLI call
     * @type {string | Buffer | undefined}
     */
    let fileArg = undefined;
    /** Stores the config arguments for this CLI call
     * @type {string[]}
     */
    const configArgs = [];

    /** Function to identify if an argument is a config option (i.e., --key=value)
     * @param {string} arg Argument passed in the CLI call.
     */
    function isConfigOption(arg) {
        return arg.startsWith('--') && arg.includes('=');
    }

    // Loop through arguments to separate file path and config options
    args.forEach(arg => {
        if (isConfigOption(arg))
            // It's a config option
            configArgs.push(arg);
        else if (!fileArg)
            // First non-config argument is assumed to be the file path
            fileArg = arg;
    });

    // Check if we have a valid file argument
    // If not, we return error and we write the instructions on how to use the library on the terminal.
    if (fileArg != undefined) {
        /** Helper function to parse config arguments from CLI
         * @param {string[]} args List of string arguments that we need to parse to understand the config flag they represent.
         */
        function parseCLIConfigArgs(args) {
            /** @type {OfficeParserConfig} */
            const config = {};
            args.forEach(arg => {
                // Split the argument by '=' to differentiate between the key and value
                const [key, value] = arg.split('=');

                // We only care about the keys that are important to us. We ignore any other key.
                switch (key) {
                    case '--ignoreNotes':
                        config.ignoreNotes = value.toLowerCase() === 'true';
                        break;
                    case '--newlineDelimiter':
                        config.newlineDelimiter = value;
                        break;
                    case '--putNotesAtLast':
                        config.putNotesAtLast = value.toLowerCase() === 'true';
                        break;
                    case '--outputErrorToConsole':
                        config.outputErrorToConsole = value.toLowerCase() === 'true';
                        break;
                    case '--extractImages':
                        config.extractImages = value.toLowerCase() === 'true';
                        break;
                    case '--extractCharts':
                        config.extractCharts = value.toLowerCase() === 'true';
                        break;
                }
            });

            return config;
        }

        // Parse CLI config arguments
        const config = parseCLIConfigArgs(configArgs);

        // Execute parseOfficeAsync with file and config
        parseOfficeAsync(fileArg, config)
            .then(result => {
                if (result.blocks && result.blocks.length > 0) {
                    const imageBlocks = result.blocks.filter(b => b.type === 'image');
                    console.log(`\n[Extracted ${result.blocks.length} block(s), ${imageBlocks.length} image(s)]`);

                    // Save images to current directory when extractImages is enabled
                    if (config.extractImages && imageBlocks.length > 0) {
                        imageBlocks.forEach((image, index) => {
                            const extension = image.mimeType.split('/')[1] || 'bin';
                            const filename = image.filename || `image_${index + 1}.${extension}`;
                            fs.writeFileSync(filename, image.buffer);
                            console.log(`  Saved: ${filename}`);
                        });
                    }
                }
            })
            .catch(error => console.error(ERRORHEADER + error));
    } else {
        console.error(ERRORMSG.improperArguments);

        const CLI_INSTRUCTIONS =
            `
=== How to Use officeParser CLI ===

Usage:
    node officeparser [--configOption=value] [FILE_PATH]

Example:
    node officeparser --ignoreNotes=true --putNotesAtLast=true ./example.docx
    node officeparser --extractImages=true ./document.docx

Config Options:
    --ignoreNotes=[true|false]          Flag to ignore notes from files like PowerPoint. Default is false.
    --newlineDelimiter=[delimiter]      The delimiter to use for new lines. Default is '\\n'.
    --putNotesAtLast=[true|false]       Flag to collect notes at the end of files like PowerPoint. Default is false.
    --outputErrorToConsole=[true|false] Flag to output errors to the console. Default is false.
    --extractImages=[true|false]        Flag to extract images from files. Default is false. Images are saved to current directory.

Note:
    The order of file path and config options doesn't matter.
`;
        // Usage instructions for the user
        console.log(CLI_INSTRUCTIONS);
    }
}

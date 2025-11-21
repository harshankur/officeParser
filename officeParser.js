#!/usr/bin/env node

// @ts-check

const concat        = require('concat-stream');
const { DOMParser } = require('@xmldom/xmldom');
const fileType      = require('file-type');
const fs            = require('fs');
const yauzl         = require('yauzl');

/** Header for error messages */
const ERRORHEADER = "[OfficeParser]: ";
/** Error messages */
const ERRORMSG = {
    extensionUnsupported: (ext) =>      `Sorry, OfficeParser currently support docx, pptx, xlsx, odt, odp, ods, pdf files only. Create a ticket in Issues on github to add support for ${ext} files. Stay tuned for further updates.`,
    fileCorrupted:        (filepath) => `Your file ${filepath} seems to be corrupted. If you are sure it is fine, please create a ticket in Issues on github with the file to reproduce error.`,
    fileDoesNotExist:     (filepath) => `File ${filepath} could not be found! Check if the file exists or verify if the relative path to the file is correct from your terminal's location.`,
    locationNotFound:     (location) => `Entered location ${location} is not reachable! Please make sure that the entered directory location exists. Check relative paths and reenter.`,
    improperArguments:                  `Improper arguments`,
    improperBuffers:                    `Error occured while reading the file buffers`,
    invalidInput:                       `Invalid input type: Expected a Buffer or a valid file path`
}

/** Returns parsed xml document for a given xml text.
 * @param {string} xml The xml string from the doc file
 * @returns {XMLDocument}
 */
const parseString = (xml) => {
    let parser = new DOMParser();
    return parser.parseFromString(xml, "text/xml");
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
 */

/**
 * @typedef {Object} Image
 * @property {Buffer} buffer The image data.
 * @property {string} type The MIME type of the image (e.g., 'image/jpeg', 'image/png').
 * @property {string} [filename] The original filename of the image, if available.
 */

/**
 * @typedef {Object} ParseOfficeResult
 * @property {string} text The extracted text content.
 * @property {Image[]} images An array of extracted images. This will be empty if `extractImages` is false or no images are found.
 */


/** Parse relationships from Word document.xml.rels file
 * @param {Buffer|string} relsFileContent Content of the rels file
 * @returns {Object.<string, string>} Map of relationship IDs to targets
 */
function parseWordRelationships(relsFileContent) {
    const rels = {};
    const relsDoc = parseString(relsFileContent.toString());
    const relationships = relsDoc.getElementsByTagName('Relationship');
    for (let i = 0; i < relationships.length; i++) {
        const rel = relationships[i];
        const id = rel.getAttribute('Id');
        if (id) {
            rels[id] = rel.getAttribute('Target');
        }
    }
    return rels;
}

/** Extract images from Word document content files
 * @param {Array<string>} contentFiles Array of XML content strings
 * @param {Object.<string, string>} relationships Map of relationship IDs to targets
 * @param {Array<{path: string, content: Buffer}>} mediaFiles Array of media file objects
 * @param {OfficeParserConfig} config Config object
 * @returns {Array<{buffer: Buffer, type: string, filename: string}>} Array of extracted images
 */
function extractImagesFromDocx(contentFiles, relationships, mediaFiles, config) {
    const images = [];

    contentFiles.forEach(xmlContent => {
        const xmlDoc = parseString(xmlContent);
        const drawingNodes = xmlDoc.getElementsByTagName('w:drawing');

        for (let i = 0; i < drawingNodes.length; i++) {
            const blip = drawingNodes[i].getElementsByTagName('a:blip')[0];
            if (blip) {
                const embedId = blip.getAttribute('r:embed');
                if (embedId && relationships[embedId]) {
                    const imagePath = 'word/' + relationships[embedId].replace('../', '');
                    const imageFile = mediaFiles.find(mf => mf.path === imagePath);
                    if (imageFile) {
                        const filename = imagePath.split('/').pop();
                        images.push({
                            buffer: imageFile.content,
                            type: getMimeTypeFromFilename(filename),
                            filename: filename
                        });
                    } else if (config.outputErrorToConsole) {
                        console.warn(`${ERRORHEADER}Image referenced but not found: ${imagePath}`);
                    }
                }
            }
        }
    });

    return images;
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
    const footnotesFileRegex   = /word\/footnotes[\d+]?.xml/g;
    const endnotesFileRegex    = /word\/endnotes[\d+]?.xml/g;
    const relsFileRegex        = /word\/_rels\/document.xml.rels/g;
    const mediaFileRegex       = /word\/media\//g;

    const filesToExtract = [mainContentFileRegex, footnotesFileRegex, endnotesFileRegex];
    if (config.extractImages) {
        filesToExtract.push(relsFileRegex, mediaFileRegex);
    }

    extractFiles(file, x => filesToExtract.some(fileRegex => x.match(fileRegex)), config.extractImages)
        .then(files => {
            // Verify if atleast the document xml file exists in the extracted files list.
            if (!files.some(file => file.path.match(mainContentFileRegex)))
                throw ERRORMSG.fileCorrupted(file);

            // Extract content files once for both image and text extraction
            const contentFiles = files
                .filter(file => file.path.match(mainContentFileRegex) || file.path.match(footnotesFileRegex) || file.path.match(endnotesFileRegex))
                .map(file => file.content.toString());

            // Extract images if requested
            let images = [];
            if (config.extractImages) {
                const relsFile = files.find(file => file.path.match(relsFileRegex));
                const mediaFiles = files.filter(file => file.path.match(mediaFileRegex));
                const relationships = relsFile ? parseWordRelationships(relsFile.content) : {};
                images = extractImagesFromDocx(contentFiles, relationships, mediaFiles, config);
            }

            const delimiter = config.newlineDelimiter ?? '\n';
            const responseText = contentFiles.map(xmlContent => {
                const xmlDoc = parseString(xmlContent);
                return extractTextFromXmlParagraphs(xmlDoc, 'w:p', 'w:t', delimiter);
            });

            const text = responseText.join(delimiter);

            callback({ text: text, images: images }, undefined);
        })
        .catch(e => callback(undefined, e));
}

/** Extract images from PowerPoint slide content files
 * @param {Array<{path: string, content: Buffer}>} contentFiles Array of slide file objects
 * @param {Object.<string, Object.<string, string>>} slideRelationships Map of slide path to relationship map
 * @param {Array<{path: string, content: Buffer}>} mediaFiles Array of media file objects
 * @param {OfficeParserConfig} config Config object
 * @returns {Array<{buffer: Buffer, type: string, filename: string}>} Array of extracted images
 */
function extractImagesFromPptx(contentFiles, slideRelationships, mediaFiles, config) {
    const images = [];

    contentFiles.forEach(file => {
        const xmlDoc = parseString(file.content.toString());
        const blipNodes = xmlDoc.getElementsByTagName('a:blip');

        // Get relationships for this specific slide
        const relationships = slideRelationships[file.path] || {};

        for (let i = 0; i < blipNodes.length; i++) {
            const blip = blipNodes[i];
            const embedId = blip.getAttribute('r:embed');
            if (embedId && relationships[embedId]) {
                const imagePath = 'ppt/' + relationships[embedId].replace('../', '');
                const imageFile = mediaFiles.find(mf => mf.path === imagePath);
                if (imageFile) {
                    const filename = imagePath.split('/').pop();
                    images.push({
                        buffer: imageFile.content,
                        type: getMimeTypeFromFilename(filename),
                        filename: filename
                    });
                } else if (config.outputErrorToConsole) {
                    console.warn(`${ERRORHEADER}Image referenced but not found: ${imagePath}`);
                }
            }
        }
    });

    return images;
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
    const slidesRegex   = /ppt\/slides\/slide\d+.xml/g;
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
        // So, we will filter out all the empty a:p tags and then combine all the a:t tag text inside for creating our response text.
        // ******************************************************************************************************
        .then(files => {
            // Extract content files for text processing
            const contentFiles = files
                .filter(file => file.path.match(allFilesRegex));

            // Extract images if requested
            let images = [];
            if (config.extractImages) {
                const relsFiles = files.filter(file => file.path.match(relsFileRegex));
                const mediaFiles = files.filter(file => file.path.match(mediaFileRegex));

                // Parse relationship files and map them to their corresponding slide
                // Each slide has its own rels file with relationship IDs scoped to that slide
                const slideRelationships = {};
                relsFiles.forEach(relsFile => {
                    // Convert ppt/slides/_rels/slide1.xml.rels -> ppt/slides/slide1.xml
                    const slidePath = relsFile.path.replace('/_rels', '').replace('.rels', '');
                    slideRelationships[slidePath] = parseWordRelationships(relsFile.content);
                });

                images = extractImagesFromPptx(contentFiles, slideRelationships, mediaFiles, config);
            }

            const delimiter = config.newlineDelimiter ?? '\n';
            const responseText = contentFiles.map(file => {
                const xmlDoc = parseString(file.content.toString());
                return extractTextFromXmlParagraphs(xmlDoc, 'a:p', 'a:t', delimiter);
            });

            // Respond by calling the Callback function.
            callback({ text: responseText.join(delimiter), images: images }, undefined);
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
    const sheetsRegex     = /xl\/worksheets\/sheet\d+.xml/g;
    const drawingsRegex   = /xl\/drawings\/drawing\d+.xml/g;
    const chartsRegex     = /xl\/charts\/chart\d+.xml/g;
    const stringsFilePath = 'xl/sharedStrings.xml';

    extractFiles(file, x => [sheetsRegex, drawingsRegex, chartsRegex].some(fileRegex => x.match(fileRegex)) || x == stringsFilePath)
        .then(files => {
            // Verify if atleast the slides xml files exist in the extracted files list.
            if (files.length == 0 || !files.map(file => file.path).some(filename => filename.match(sheetsRegex)))
                throw ERRORMSG.fileCorrupted(file);

            return {
                sheetFiles:        files.filter(file => file.path.match(sheetsRegex)).map(file => file.content),
                drawingFiles:      files.filter(file => file.path.match(drawingsRegex)).map(file => file.content),
                chartFiles:        files.filter(file => file.path.match(chartsRegex)).map(file => file.content),
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
            let responseText = [];

            /** Function to check if the given c node is a valid inline string node. */
            function isValidInlineStringCNode(cNode) {
                // Initial check to see if the passed node is a cNode
                if (cNode.tagName.toLowerCase() != 'c')
                    return false;
                if (cNode.getAttribute("t") != 'inlineStr')
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
                return cNode.getElementsByTagName("v")[0]
                    && cNode.getElementsByTagName("v")[0].childNodes[0]
                    && cNode.getElementsByTagName("v")[0].childNodes[0].nodeValue != ''
            }

            /** Find text nodes with t tags in sharedStrings.xml file. If the sharedStringsFile is not present, we return an empty array. */
            const sharedStringsXmlSiNodesList = xmlContentFilesObject.sharedStringsFile != undefined
                ? parseString(xmlContentFilesObject.sharedStringsFile).getElementsByTagName("si")
                : [];

            /** Create shared string array. This will be used as a map to get strings from within sheet files. */
            const sharedStrings = Array.from(sharedStringsXmlSiNodesList)
                .map(siNode => {
                    // Concatenate all <t> nodes within the <si> node
                    return Array.from(siNode.getElementsByTagName("t"))
                        .map(tNode => tNode.childNodes[0]?.nodeValue ?? '') // Extract text content from each <t> node
                        .join(''); // Combine all <t> node text into a single string
                });

            // Parse Sheet files
            xmlContentFilesObject.sheetFiles.forEach(sheetXmlContent => {
                /** Find text nodes with c tags in sharedStrings xml file */
                const sheetsXmlCNodesList = parseString(sheetXmlContent).getElementsByTagName("c");
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
                                const isIndexInSharedStrings = cNode.getAttribute("t") == "s";
                                /** Find value nodes represented by v tags */
                                const value = cNode.getElementsByTagName("v")[0].childNodes[0].nodeValue;
                                const valueAsIndex = Number(value);
                                // Validate text
                                if (isIndexInSharedStrings && (valueAsIndex != parseInt(value, 10) || valueAsIndex >= sharedStrings.length))
                                    throw ERRORMSG.fileCorrupted(file);

                                return isIndexInSharedStrings
                                        ? sharedStrings[valueAsIndex]
                                        : value;
                            }
                            // Should not reach here. If we do, it means we are not filtering out items that we are not ready to process.
                            // Not the case now but it could happen if we change the filtering logic without updating the processing logic.
                            // So, it is better to error out here.
                            handleError(`Invalid c node found in sheet xml content: ${cNode}`, callback, config.outputErrorToConsole);
                            return '';
                        })
                        // Join each cell text within a sheet with a space.
                        .join(config.newlineDelimiter ?? "\n")
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
                const chartsXmlCVNodesList = parseString(chartXmlContent).getElementsByTagName("c:v");
                /** Store all the text content to respond */
                responseText.push(
                    Array.from(chartsXmlCVNodesList)
                        .filter(cVNode => cVNode.childNodes[0] && cVNode.childNodes[0].nodeValue)
                        .map(cVNode => cVNode.childNodes[0].nodeValue)
                        .join(config.newlineDelimiter ?? "\n")
                );
            });

            // Respond by calling the Callback function.
            callback(responseText.join(config.newlineDelimiter ?? "\n"), undefined);
        })
        .catch(e => callback(undefined, e));
}


/** Extract images from OpenOffice content files
 * @param {Array<Document>} xmlDocuments Array of parsed XML documents
 * @param {Array<{path: string, content: Buffer}>} mediaFiles Array of media file objects
 * @param {OfficeParserConfig} config Config object
 * @returns {Array<{buffer: Buffer, type: string, filename: string}>} Array of extracted images
 */
function extractImagesFromOpenOffice(xmlDocuments, mediaFiles, config) {
    const images = [];
    const addedImages = new Set(); // Track to avoid duplicates

    xmlDocuments.forEach(xmlDoc => {
        if (!xmlDoc) {
            if (config.outputErrorToConsole) {
                console.warn(`${ERRORHEADER}Skipping undefined XML document in image extraction`);
            }
            return;
        }
        const imageNodes = xmlDoc.getElementsByTagName('draw:image');

        for (let i = 0; i < imageNodes.length; i++) {
            const imageNode = imageNodes[i];
            const href = imageNode.getAttribute('xlink:href');
            if (href && !addedImages.has(href)) {
                const imageFile = mediaFiles.find(mf => mf.path === href);
                if (imageFile) {
                    const filename = href.split('/').pop();
                    images.push({
                        buffer: imageFile.content,
                        type: getMimeTypeFromFilename(filename),
                        filename: filename
                    });
                    addedImages.add(href);
                } else if (config.outputErrorToConsole) {
                    console.warn(`${ERRORHEADER}Image referenced but not found: ${href}`);
                }
            }
        }
    });

    return images;
}

/** Main function for parsing text from open office files
 * @param {string | Buffer}    file     File path or Buffers
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {void}
 */
function parseOpenOffice(file, callback, config) {
    /** The target content xml file for the openoffice file. */
    const mainContentFilePath     = 'content.xml';
    const objectContentFilesRegex = /Object \d+\/content.xml/g;
    const mediaFileRegex          = /Pictures\//g;

    const filesToExtract = [mainContentFilePath, objectContentFilesRegex];
    if (config.extractImages) {
        filesToExtract.push(mediaFileRegex);
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
                mainContentFile:    files.filter(file => file.path == mainContentFilePath).map(file => file.content.toString())[0],
                objectContentFiles: files.filter(file => file.path.match(objectContentFilesRegex)).map(file => file.content.toString()),
            };

            if (config.extractImages) {
                result.mediaFiles = files.filter(file => file.path.match(mediaFileRegex));
            }

            return result;
        })
        // ********************************** openoffice xml files explanation **********************************
        // Structure of xmlContent of openoffice files is simple.
        // All text nodes are within text:h and text:p tags with all kinds of formatting within nested tags.
        // All text in these tags are separated by new line delimiters.
        // Objects like charts in ods files are in Object d+/content.xml with the same way as above.
        // ******************************************************************************************************
        .then(xmlContentFilesObject => {
            /** Store all the notes text content to respond */
            let notesText = [];
            /** Store all the text content to respond */
            let responseText = [];

            /** List of allowed text tags */
            const allowedTextTags = ["text:p", "text:h"];
            /** List of notes tags */
            const notesTag = "presentation:notes";

            /** Main dfs traversal function that goes from one node to its children and returns the value out. */
            function extractAllTextsFromNode(root) {
                let xmlTextArray = []
                for (let i = 0; i < root.childNodes.length; i++)
                    traversal(root.childNodes[i], xmlTextArray, true);
                return xmlTextArray.join("");
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
                                notesText.push(config.newlineDelimiter ?? "\n");
                        }
                        else {
                            xmlTextArray.push(value);
                            if (allowedTextTags.includes(node.parentNode.tagName) && !isFirstRecursion)
                                xmlTextArray.push(config.newlineDelimiter ?? "\n");
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
            // Iterate over each xmlContent and extract text from them.
            xmlContentArray.forEach(xmlContent => {
                /** Find text nodes with text:h and text:p tags in xmlContent */
                const xmlTextNodesList = [...Array.from(xmlContent
                                                .getElementsByTagName("*"))
                                                .filter(node => allowedTextTags.includes(node.tagName)
                                                    && !isInvalidTextNode(node.parentNode))
                                            ];
                /** Store all the text content to respond */
                responseText.push(
                    xmlTextNodesList
                        // Add every text information from within this textNode and combine them together.
                        .map(textNode => extractAllTextsFromNode(textNode))
                        .filter(text => text != "")
                        .join(config.newlineDelimiter ?? "\n")
                );
            });

            // Add notes text at the end if the user config says so.
            // Note that we already have pushed the text content to notesText array while extracting all texts from the nodes.
            if (!config.ignoreNotes && config.putNotesAtLast)
                responseText = [...responseText, ...notesText];

            // Extract images if requested
            let images = [];
            if (config.extractImages && xmlContentFilesObject.mediaFiles) {
                try {
                    images = extractImagesFromOpenOffice(xmlContentArray, xmlContentFilesObject.mediaFiles, config);
                } catch (error) {
                    if (config.outputErrorToConsole) {
                        console.error(`${ERRORHEADER}Error extracting images from OpenOffice file:`, error.message);
                    }
                    // Continue with empty images array
                }
            }

            // Respond by calling the Callback function.
            callback({ text: responseText.join(config.newlineDelimiter ?? "\n"), images: images }, undefined);
        })
        .catch(e => callback(undefined, e));
}

/** Extract text from all PDF pages
 * @param {Array<Object>} pages Array of PDF page objects
 * @param {OfficeParserConfig} config Config object
 * @returns {Promise<string>} Extracted text from all pages
 */
async function extractTextFromPdfPages(pages, config) {
    const delimiter = config.newlineDelimiter ?? '\n';
    const textContentArray = [];

    for (const page of pages) {
        const textContent = await page.getTextContent();
        textContentArray.push(textContent);
    }

    // str already contains any space that was in the text.
    // So, we only care about when to add the new line.
    // That we determine using transform[5] value which is the y-coordinate of the item object.
    // So, if there is a mismatch in the transform[5] value between the current item and the previous item, we put a line break.
    const responseText = textContentArray
        .map(textContent => textContent.items)      // Get all the items
        .flat()                                     // Flatten all the items object
        .reduce((a, v) =>  (
            // the items could be TextItem or a TextMarkedContent.
            // We are only interested in the TextItem which has a str property.
            'str' in v && v.str != ''
                ? {
                    text: a.text + (v.transform[5] != a.transform5 ? delimiter : '') + v.str,
                    transform5: v.transform[5]
                } : {
                    text: a.text,
                    transform5: a.transform5
                }
        ),
        {
            text: '',
            transform5: undefined
        }).text;

    return responseText;
}

/** Extract images from all PDF pages
 * @param {Array<Object>} pages Array of PDF page objects
 * @param {Object} pdfjs PDF.js module
 * @param {OfficeParserConfig} config Config object
 * @returns {Promise<Array<{buffer: Buffer, type: string}>>} Array of extracted images
 */
async function extractImagesFromPdfPages(pages, pdfjs, config) {
    const images = [];

    for (const page of pages) {
        const operatorList = await page.getOperatorList(true);
        const imageOps = operatorList.fnArray
            .map((fn, j) => fn === pdfjs.OPS.paintImageXObject ? operatorList.argsArray[j][0] : null)
            .filter(op => op);

        for (const op of imageOps) {
            const image = await getPdfImageResource(page, op);
            if (image) {
                images.push({
                    buffer: Buffer.from(image.data),
                    type: image.kind === pdfjs.ImageKind.JPEG ? 'image/jpeg' : 'image/png'
                });
            } else if (config.outputErrorToConsole) {
                console.warn(`${ERRORHEADER}PDF image resource not found: ${op}`);
            }
        }
    }

    return images;
}

/** Main function for parsing text from pdf files
 * @param {string | Buffer}    file     File path or Buffers
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {Promise<void>}
 */
async function parsePdf(file, callback, config) {
    // Wait for pdfjs module to be loaded once
    // Lazy import pdfjs to avoid Node startup issues for environments that don't use PDF parsing
    const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');

    // Get the pdfjs document for the filepath or Uint8Array buffers.
    // pdfjs does not accept Buffers directly, so we convert them to Uint8Array.
    pdfjs.getDocument(file instanceof Buffer ? new Uint8Array(file) : file).promise
        .then(async document => {
            // Load all pages
            const pagePromises = Array.from({ length: document.numPages }, (_, index) => document.getPage(index + 1));
            const pages = await Promise.all(pagePromises);

            // Extract text from all pages
            const text = await extractTextFromPdfPages(pages, config);

            // Extract images if requested
            const images = config.extractImages
                ? await extractImagesFromPdfPages(pages, pdfjs, config)
                : [];

            callback({ text, images }, undefined);
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
    let file = srcFile instanceof ArrayBuffer ? Buffer.from(srcFile)
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
            res({ file: file, ext: file.split(".").pop() });
        }
        else
            rej(ERRORMSG.invalidInput);
    });

    // Process filePreparedPromise resolution.
    filePreparedPromise
        .then(({ file, ext }) => {
            // Switch between parsing functions depending on extension.
            switch (ext) {
                case "docx":
                    parseWord(file, internalCallback, internalConfig);
                    break;
                case "pptx":
                    parsePowerPoint(file, internalCallback, internalConfig);
                    break;
                case "xlsx":
                    parseExcel(file, internalCallback, internalConfig);
                    break;
                case "odt":
                case "odp":
                case "ods":
                    parseOpenOffice(file, internalCallback, internalConfig);
                    break;
                case "pdf":
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
                    callback(data, undefined);
                } else {
                    callback({ text: data, images: [] }, undefined);
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
                }
                else
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
module.exports.parseOffice      = parseOffice;
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
                }
            });

            return config;
        }

        // Parse CLI config arguments
        const config = parseCLIConfigArgs(configArgs);

        // Execute parseOfficeAsync with file and config
        parseOfficeAsync(fileArg, config)
            .then(result => {
                console.log(result.text);
                if (result.images && result.images.length > 0) {
                    console.log(`\n[Extracted ${result.images.length} image(s)]`);
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

Config Options:
    --ignoreNotes=[true|false]          Flag to ignore notes from files like PowerPoint. Default is false.
    --newlineDelimiter=[delimiter]      The delimiter to use for new lines. Default is '\\n'.
    --putNotesAtLast=[true|false]       Flag to collect notes at the end of files like PowerPoint. Default is false.
    --outputErrorToConsole=[true|false] Flag to output errors to the console. Default is false.
    --extractImages=[true|false]        Flag to extract images from docx and pdf files. Default is false.

Note:
    The order of file path and config options doesn't matter.
`;
        // Usage instructions for the user
        console.log(CLI_INSTRUCTIONS);
    }
}

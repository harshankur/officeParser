#!/usr/bin/env node

const decompress    = require('decompress');
const fs            = require('fs');
const rimraf        = require('rimraf');
const fileType      = require('file-type');
const pdfParse      = require('pdf-parse');
const { DOMParser } = require('@xmldom/xmldom');

/** Header for error messages */
const ERRORHEADER = "[OfficeParser]: ";
/** Error messages */
const ERRORMSG = {
    extensionUnsupported: (ext) =>      `Sorry, OfficeParser currently support docx, pptx, xlsx, odt, odp, ods, pdf files only. Create a ticket in Issues on github to add support for ${ext} files. Stay tuned for further updates.`,
    fileCorrupted:        (filepath) => `Your file ${filepath} seems to be corrupted. If you are sure it is fine, please create a ticket in Issues on github with the file to reproduce error.`,
    fileDoesNotExist:     (filepath) => `File ${filepath} could not be found! Check if the file exists or verify if the relative path to the file is correct from your terminal's location.`,
    locationNotFound:     (location) => `Entered location ${location} is not reachable! Please make sure that the entered directory location exists. Check relative paths and reenter.`,
    improperArguments:                  `Improper arguments`,
    improperBuffers:                    `Error occured while reading the file buffers`
}
/** Default sublocation for decompressing files under the current directory. */
const DEFAULTDECOMPRESSSUBLOCATION = "officeParserTemp";

/** Console error if allowed
 * @param {string} errorMessage         Error message to show on the console
 * @param {string} outputErrorToConsole Flag to show log on console. Ignore if not true.
 * @returns {void}
 */
function consoleError(errorMessage, outputErrorToConsole) {
    if (!errorMessage || !outputErrorToConsole)
        return;
    console.error(ERRORHEADER + errorMessage);
}

/** Returns parsed xml document for a given xml text.
 * @param {string} xml The xml string from the doc file
 * @returns {XMLDocument}
 */
const parseString = (xml) => {
    let parser = new DOMParser();
    return parser.parseFromString(xml, "text/xml");
};

/** @typedef {Object} OfficeParserConfig
 * @property {string}  [tempFilesLocation]    The directory where officeparser stores the temp files . The final decompressed data will be put inside officeParserTemp folder within your directory. Please ensure that this directory actually exists. Default is officeParsertemp.
 * @property {boolean} [preserveTempFiles]    Flag to not delete the internal content files and the duplicate temp files that it uses after unzipping office files. Default is false. It deletes all of those files.
 * @property {boolean} [outputErrorToConsole] Flag to show all the logs to console in case of an error irrespective of your own handling.
 * @property {string}  [newlineDelimiter]     The delimiter used for every new line in places that allow multiline text like word. Default is \n.
 * @property {boolean} [ignoreNotes]          Flag to ignore notes from parsing in files like powerpoint. Default is false. It includes notes in the parsed text by default.
 * @property {boolean} [putNotesAtLast]       Flag, if set to true, will collectively put all the parsed text from notes at last in files like powerpoint. Default is false. It puts each notes right after its main slide content. If ignoreNotes is set to true, this flag is also ignored.
 */


/** Main function for parsing text from word files
 * @param {string}             filepath File path
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {void}
 */
function parseWord(filepath, callback, config) {
    /** The target content xml file for the docx file. */
    const mainContentFile = 'word/document.xml';
    const footnotesFile   = 'word/footnotes.xml';
    const endnotesFile    = 'word/endnotes.xml';
    /** The decompress location which contains the filename in it */
    const decompressLocation = `${config.tempFilesLocation}/${filepath.split("/").pop()}`;
    decompress(filepath,
        decompressLocation,
        { filter: x => [mainContentFile, footnotesFile, endnotesFile].includes(x.path) }
    )
    .then(files => {
        if (files.length == 0)
            throw ERRORMSG.fileCorrupted(filepath);

        return [...files.filter(file => file.path == mainContentFile),
                    ...files.filter(file => file.path == footnotesFile),
                    ...files.filter(file => file.path == endnotesFile)
                    ]
                    .map(file => fs.readFileSync(`${decompressLocation}/${file.path}`, 'utf8'));
    })
    // ************************************* word xml files explanation *************************************
    // Structure of xmlContent of a word file is simple.
    // All text nodes are within w:t tags and each of the text nodes that belong in one paragraph are clubbed together within a w:p tag.
    // So, we will filter out all the empty w:p tags and then combine all the w:t tag text inside for creating our response text.
    // ******************************************************************************************************
    .then(xmlContentArray => {
        /** Store all the text content to respond */
        let responseText = [];

        xmlContentArray.forEach(xmlContent => {
            /** Find text nodes with w:p tags */
            const xmlParagraphNodesList = parseString(xmlContent).getElementsByTagName("w:p");
            /** Store all the text content to respond */
            responseText.push(
                Array.from(xmlParagraphNodesList)
                    // Filter paragraph nodes than do not have any text nodes which are identifiable by w:t tag
                    .filter(paragraphNode => paragraphNode.getElementsByTagName("w:t").length != 0)
                    .map(paragraphNode => {
                        // Find text nodes with w:t tags
                        const xmlTextNodeList = paragraphNode.getElementsByTagName("w:t");
                        // Join the texts within this paragraph node without any spaces or delimiters.
                        return Array.from(xmlTextNodeList).map(textNode => textNode.childNodes[0].nodeValue).join("");
                    })
                    // Join each paragraph text with a new line delimiter.
                    .join(config.newlineDelimiter ?? "\n")
            );
        });

        // Join all responseText array
        responseText = responseText.join(config.newlineDelimiter ?? "\n");
        // Respond by calling the Callback function.
        callback(responseText, undefined);
    })
    .catch(e => callback(undefined, e));
}

/** Main function for parsing text from PowerPoint files
 * @param {string}             filepath File path
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {void}
 */
function parsePowerPoint(filepath, callback, config) {
    // Files regex that hold our content of interest
    const allFilesRegex = /ppt\/(notesSlides|slides)\/(notesSlide|slide)\d+.xml/g;
    const slidesRegex   = /ppt\/slides\/slide\d+.xml/g;

    /** The decompress location which contains the filename in it */
    const decompressLocation = `${config.tempFilesLocation}/${filepath.split("/").pop()}`;
    decompress(filepath,
        decompressLocation,
        { filter: x => x.path.match(config.ignoreNotes ? slidesRegex : allFilesRegex) }
    )
    .then(files => {
        // Check if files is corrupted
        if (files.length == 0)
            throw ERRORMSG.fileCorrupted(filepath);

        // Check if any sorting is required.
        if (!config.ignoreNotes && config.putNotesAtLast)
            // Sort files according to previous order of taking text out of ppt/slides followed by ppt/notesSlides
            // For this we are looking at the index of notes which results in -1 in the main slide file and exists at a certain index in notes file names.
            files.sort((a,b) => a.path.indexOf("notes") - b.path.indexOf("notes"));

        // Returning an array of all the xml contents read using fs.readFileSync
        return files.map(file => fs.readFileSync(`${decompressLocation}/${file.path}`, 'utf8'));
    })
    // ******************************** powerpoint xml files explanation ************************************
    // Structure of xmlContent of a powerpoint file is simple.
    // There are multiple xml files for each slide and correspondingly their notesSlide files.
    // All text nodes are within a:t tags and each of the text nodes that belong in one paragraph are clubbed together within a a:p tag.
    // So, we will filter out all the empty a:p tags and then combine all the a:t tag text inside for creating our response text.
    // ******************************************************************************************************
    .then(xmlContentArray => {
        /** Store all the text content to respond */
        let responseText = [];

        xmlContentArray.forEach(xmlContent => {
            /** Find text nodes with a:p tags */
            const xmlParagraphNodesList = parseString(xmlContent).getElementsByTagName("a:p");
            /** Store all the text content to respond */
            responseText.push(
                Array.from(xmlParagraphNodesList)
                    // Filter paragraph nodes than do not have any text nodes which are identifiable by a:t tag
                    .filter(paragraphNode => paragraphNode.getElementsByTagName("a:t").length != 0)
                    .map(paragraphNode => {
                        /** Find text nodes with a:t tags */
                        const xmlTextNodeList = paragraphNode.getElementsByTagName("a:t");
                        return Array.from(xmlTextNodeList).map(textNode => textNode.childNodes[0].nodeValue).join("");
                    })
                    .join(config.newlineDelimiter ?? "\n")
            );
        });

        // Join all responseText array
        responseText = responseText.join(config.newlineDelimiter ?? "\n");
        // Respond by calling the Callback function.
        callback(responseText, undefined);
    })
    .catch(e => callback(undefined, e));
}

/** Main function for parsing text from Excel files
 * @param {string}             filepath File path
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {void}
 */
function parseExcel(filepath, callback, config) {
    // Files regex that hold our content of interest
    const sheetsRegex     = /xl\/worksheets\/sheet\d+.xml/g;
    const drawingsRegex   = /xl\/drawings\/drawing\d+.xml/g;
    const chartsRegex     = /xl\/charts\/chart\d+.xml/g;
    const stringsFilePath = 'xl/sharedStrings.xml';

    /** The decompress location which contains the filename in it */
    const decompressLocation = `${config.tempFilesLocation}/${filepath.split("/").pop()}`;
    decompress(filepath,
        decompressLocation,
        { filter: x => ([sheetsRegex, drawingsRegex, chartsRegex].findIndex(fileRegex => x.path.match(fileRegex)) > -1) || (x.path == stringsFilePath )}
    )
    .then(files => {
        if (files.length == 0)
            throw ERRORMSG.fileCorrupted(filepath);

        return {
            sheetFiles:        files.filter(file => file.path.match(sheetsRegex)).map(file => fs.readFileSync(`${decompressLocation}/${file.path}`, 'utf8')),
            drawingFiles:      files.filter(file => file.path.match(drawingsRegex)).map(file => fs.readFileSync(`${decompressLocation}/${file.path}`, 'utf8')),
            chartFiles:        files.filter(file => file.path.match(chartsRegex)).map(file => fs.readFileSync(`${decompressLocation}/${file.path}`, 'utf8')),
            sharedStringsFile: files.filter(file => file.path == stringsFilePath).map(file => fs.readFileSync(`${decompressLocation}/${file.path}`, 'utf8'))[0],
        };
    })
    // ********************************** excel xml files explanation ***************************************
    // Structure of xmlContent of an excel file is a bit complex.
    // We have a sharedStrings.xml file which has strings inside t tags
    // Each sheet has an individual sheet xml file which has numbers in v tags (probably value) inside c tags (probably cell)
    // Each value of v tag is to be used as it is if the "t" attribute (probably type) of c tag is not "s" (probably shared string)
    // If the "t" attribute of c tag is "s", then we use the value to select value from sharedStrings array with the value as its index.
    // Drawing files contain all text for each drawing and have text nodes in a:t and paragraph nodes in a:p.
    // ******************************************************************************************************
    .then(xmlContentFilesObject => {
        /** Store all the text content to respond */
        let responseText = [];

        /** Find text nodes with t tags in sharedStrings xml file */
        const sharedStringsXmlTNodesList = parseString(xmlContentFilesObject.sharedStringsFile).getElementsByTagName("t");
        /** Create shared string array. This will be used as a map to get strings from within sheet files. */
        const sharedStrings = Array.from(sharedStringsXmlTNodesList).map(tNode => tNode.childNodes[0].nodeValue);

        // Parse Sheet files
        xmlContentFilesObject.sheetFiles.forEach(sheetXmlContent => {
            /** Find text nodes with c tags in sharedStrings xml file */
            const sheetsXmlCNodesList = parseString(sheetXmlContent).getElementsByTagName("c");
            // Traverse through the nodes list and fill responseText with either the number value in its v node or find a mapped string from sharedStrings.
            responseText.push(
                Array.from(sheetsXmlCNodesList)
                    // Filter c nodes than do not have any v nodes
                    .filter(cNode => cNode.getElementsByTagName("v").length != 0)
                    .map(cNode => {
                        /** Flag whether this node's value represents a string index */
                        const isString = cNode.getAttribute("t") == "s";
                        /** Find value nodes represented by v tags */
                        const value = cNode.getElementsByTagName("v")[0].childNodes[0].nodeValue;
                        // Validate text
                        if (isString && value >= sharedStrings.length)
                            throw ERRORMSG.fileCorrupted(filepath);

                        return isString
                                ? sharedStrings[value]
                                : value;
                    })
                    // Join each cell text within a sheet with a space.
                    .join(config.newlineDelimiter ?? "\n")
            );
        });

        // Parse Drawing files
        xmlContentFilesObject.drawingFiles.forEach(drawingXmlContent => {
            /** Find text nodes with a:p tags */
            const drawingsXmlParagraphNodesList = parseString(drawingXmlContent).getElementsByTagName("a:p");
            /** Store all the text content to respond */
            responseText.push(
                Array.from(drawingsXmlParagraphNodesList)
                    // Filter paragraph nodes than do not have any text nodes which are identifiable by a:t tag
                    .filter(paragraphNode => paragraphNode.getElementsByTagName("a:t").length != 0)
                    .map(paragraphNode => {
                        /** Find text nodes with a:t tags */
                        const xmlTextNodeList = paragraphNode.getElementsByTagName("a:t");
                        return Array.from(xmlTextNodeList).map(textNode => textNode.childNodes[0].nodeValue).join("");
                    })
                    .join(config.newlineDelimiter ?? "\n")
            );
        });

        // Parse Chart files
        xmlContentFilesObject.chartFiles.forEach(chartXmlContent => {
            /** Find text nodes with c:v tags */
            const chartsXmlCVNodesList = parseString(chartXmlContent).getElementsByTagName("c:v");
            /** Store all the text content to respond */
            responseText.push(
                Array.from(chartsXmlCVNodesList)
                    .map(cVNode => cVNode.childNodes[0].nodeValue)
                    .join(config.newlineDelimiter ?? "\n")
            );
        });

        // Join all responseText array
        responseText = responseText.join(config.newlineDelimiter ?? "\n");
        // Respond by calling the Callback function.
        callback(responseText, undefined);
    })
    .catch(e => callback(undefined, e));
}


/** Main function for parsing text from open office files
 * @param {string}             filepath File path
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {void}
 */
function parseOpenOffice(filepath, callback, config) {
    /** The target content xml file for the openoffice file. */
    const mainContentFilePath     = 'content.xml';
    const objectContentFilesRegex = /Object \d+\/content.xml/g;

    /** The decompress location which contains the filename in it */
    const decompressLocation = `${config.tempFilesLocation}/${filepath.split("/").pop()}`;
    decompress(filepath,
        decompressLocation,
        { filter: x => x.path == mainContentFilePath || x.path.match(objectContentFilesRegex) }
    )
    .then(files => {
        if (files.length == 0)
            throw ERRORMSG.fileCorrupted(filepath);

        return {
            mainContentFile:    files.filter(file => file.path == mainContentFilePath).map(file => fs.readFileSync(`${decompressLocation}/${file.path}`, 'utf8'))[0],
            objectContentFiles: files.filter(file => file.path.match(objectContentFilesRegex)).map(file => fs.readFileSync(`${decompressLocation}/${file.path}`, 'utf8')),
        }
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
            if(!node.childNodes || node.childNodes.length == 0)
            {
                if (node.parentNode.tagName.indexOf('text') == 0 && node.nodeValue) {
                    if (isNotesNode(node.parentNode) && (config.putNotesAtLast || config.ignoreNotes)) {
                        notesText.push(node.nodeValue);
                        if (allowedTextTags.includes(node.parentNode.tagName) && !isFirstRecursion)
                            notesText.push(config.newlineDelimiter ?? "\n");
                    }
                    else {
                        xmlTextArray.push(node.nodeValue);
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
        const xmlContentArray = [xmlContentFilesObject.mainContentFile, ...xmlContentFilesObject.objectContentFiles].map(xmlContent => parseString(xmlContent));
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

        // Join all responseText array
        responseText = responseText.join(config.newlineDelimiter ?? "\n");
        // Respond by calling the Callback function.
        callback(responseText, undefined);
    })
    .catch(e => callback(undefined, e));
}

/** Header for error messages */
const PDFPARSEERRORHEADER = "[pdf-parse]: ";

/** Main function for parsing text from pdf files
 * @param {string}             filepath File path
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {void}
 */
function parsePdf(filepath, callback, config) {
    // Get the data buffer for the given file path.
    const dataBuffer = fs.readFileSync(filepath);

    pdfParse(dataBuffer)
        .then(data => {
            let text = data.text;
            if (!!config.newlineDelimiter && config.newlineDelimiter != "\n")
                text = text.replaceAll("\n", config.newlineDelimiter)
            callback(text, undefined);
        })
        .catch(e => callback(undefined, PDFPARSEERRORHEADER + e));
}

/** Main async function with callback to execute parseOffice for supported files
 * @param {string | Buffer}    file        File path or file buffers
 * @param {function}           callback    Callback function that returns value or error
 * @param {OfficeParserConfig} [config={}] [OPTIONAL]: Config Object for officeParser
 * @returns {void}
 */
function parseOffice(file, callback, config = {}) {
    // Make a clone of the config.
    const internalConfig = { ...config };
    // Prepare file for processing
    const filePreparedPromise = new Promise((res, rej) => {
        // Check if decompress location in the config is present.
        // If it is valid, we set the final decompression location in the config.
        // If it is not valid, we reject the promise with appropriate error message.
        if (!internalConfig.tempFilesLocation)
            internalConfig.tempFilesLocation = DEFAULTDECOMPRESSSUBLOCATION;
        else {
            if (!fs.existsSync(internalConfig.tempFilesLocation))
            {
                rej(ERRORMSG.locationNotFound(internalConfig.tempFilesLocation));
                return;
            }
            internalConfig.tempFilesLocation = `${internalConfig.tempFilesLocation}${internalConfig.tempFilesLocation.endsWith('/') ? '' : '/'}${DEFAULTDECOMPRESSSUBLOCATION}`;
        }

        // create temp file subdirectory if it does not exist
        fs.mkdirSync(`${internalConfig.tempFilesLocation}/tempfiles`, { recursive: true });

        // Check if buffer
        if (Buffer.isBuffer(file)) {
            // Guess file type from buffer
            fileType.fromBuffer(file)
                .then(data =>
                {
                    // temp file name
                    const newfilepath = getNewFileName(internalConfig.tempFilesLocation, data.ext.toLowerCase());
                    // write new file
                    fs.writeFileSync(newfilepath, file);
                    // resolve promise
                    res(newfilepath);
                })
                .catch(() => rej(ERRORMSG.improperBuffers));
            return;
        }

        // Not buffers but real file path.

        // Check if file exists
        if (!fs.existsSync(file))
            throw ERRORMSG.fileDoesNotExist(file);

        // temp file name
        const newfilepath = getNewFileName(internalConfig.tempFilesLocation, file.split(".").pop().toLowerCase());
        // Copy the file into a temp location with the temp name
        fs.copyFileSync(file, newfilepath)
        // resolve promise
        res(newfilepath);
    });

    // Process filePreparedPromise resolution.
    filePreparedPromise
        .then(filepath => {
            // File extension. Already in lowercase when we prepared the temp file above.
            const extension = filepath.split(".").pop();

            // Switch between parsing functions depending on extension.
            switch(extension) {
                case "docx":
                    parseWord(filepath, internalCallback, internalConfig);
                    break;
                case "pptx":
                    parsePowerPoint(filepath, internalCallback, internalConfig);
                    break;
                case "xlsx":
                    parseExcel(filepath, internalCallback, internalConfig);
                    break;
                case "odt":
                case "odp":
                case "ods":
                    parseOpenOffice(filepath, internalCallback, internalConfig);
                    break;
                case "pdf":
                    parsePdf(filepath, internalCallback, internalConfig);
                    break;

                default:
                    throw ERRORMSG.extensionUnsupported(extension);
            }

            /** Internal callback function that calls the user's callback function passed in argument and removes the temp files if required */
            function internalCallback(data, err) {
                // Check if we need to preserve unzipped content files or delete them.
                if (!internalConfig.preserveTempFiles)
                    // Delete decompress sublocation.
                    rimraf(internalConfig.tempFilesLocation, rimrafErr => consoleError(rimrafErr, internalConfig.outputErrorToConsole));

                // Check if there is an error. Throw if there is an error.
                if (err)
                    return handleError(err, callback, internalConfig.outputErrorToConsole);

                // Call the original callback
                callback(data, undefined);
            }
        })
        .catch(error => handleError(error, callback, internalConfig.outputErrorToConsole));
}

/**
 * Main async function that can be used with await to execute parseOffice. Or it can be used with promises.
 * @param {string | Buffer}    file        File path or file buffers
 * @param {OfficeParserConfig} [config={}] [OPTIONAL]: Config Object for officeParser
 * @returns {Promise<string>}
 */
function parseOfficeAsync(file, config = {}) {
    return new Promise((res, rej) => {
        parseOffice(file, function (data, err) {
            if (err)
                return rej(err);
            return res(data);
        }, config);
    });
}

/** Global file name iterator. */
let globalFileNameIterator = 0;
/**
 * File Name generator that takes the extension as an input and returns a file name that comprises a timestamp and an incrementing number
 * to allow the files to be sorted in chronological order
 * @param {string} tempFilesLocation Directory whether this new file needs to be stored
 * @param {string} ext               File extension for this new generated file name
 * @returns {string} 
 */
function getNewFileName(tempFilesLocation, ext) {
    // Get the iterator part of the file name
    let iteratorPart = (globalFileNameIterator++).toString().padStart(5, '0');
    // We want the iterator part of the file name to be of 5 digits.
    // Therefore, when the iterator crosses into 6 digits, we reset it to 0.
    if (globalFileNameIterator > 99999)
        globalFileNameIterator = 0;

    // Return the file name
    return `${tempFilesLocation}/tempfiles/${new Date().getTime().toString() + iteratorPart}.${ext}`;
}

/**
 * Handle error by logging it to console if permitted by the config.
 * And after that, trigger the callback function with the error value.
 * @param {string}   error                Error text
 * @param {function} callback             Callback function provided by the caller
 * @param {boolean}  outputErrorToConsole Flag to log error to console.
 * @returns {void}
 */
function handleError(error, callback, outputErrorToConsole) {
    consoleError(error, outputErrorToConsole);
    callback(undefined, ERRORHEADER + error);
}


// Export functions
module.exports.parseOffice      = parseOffice;
module.exports.parseOfficeAsync = parseOfficeAsync;


// Run this library on CLI
if ((process.argv[0].split('/').pop() == "node" || process.argv[0].split('/').pop() == "npx") && (process.argv[1].split('/').pop() == "officeParser.js" || process.argv[1].split('/').pop() == "officeparser")) {
    if (process.argv.length == 2) {
        // continue
    }
    else if (process.argv.length == 3)
        parseOfficeAsync(process.argv[2])
            .then(text => console.log(text))
            .catch(error => console.error(ERRORHEADER + error))
    else
        console.error(ERRORMSG.improperArguments)
}
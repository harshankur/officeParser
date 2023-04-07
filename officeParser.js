#!/usr/bin/env node

const decompress = require('decompress');
const xml2js = require('xml2js')
const fs = require('fs')
const rimraf = require('rimraf');

/** Header for error messages */
const ERRORHEADER = "[OfficeParser]: ";
/** Error messages */
const ERRORMSG = {
    extensionUnsupported: (ext) =>      `${ERRORHEADER}Sorry, OfficeParser currently support docx, pptx, xlsx, odt, odp, ods files only. Create a ticket in Issues on github to add support for ${ext} files. Stay tuned for further updates.`,
    fileCorrupted:        (filename) => `${ERRORHEADER}Your file ${filename} seems to be corrupted. If you are sure it is fine, please create a ticket in Issues on github with the file to reproduce error.`,
    fileDoesNotExist:     (filename) => `${ERRORHEADER}File ${filename} could not be found! Check if the file exists or verify if the relative path to the file is correct from your terminal's location.`,
    locationNotFound:     (location) => `${ERRORHEADER}Entered location ${location} is not valid! Check relative paths and reenter. OfficeParser will use root directory as decompress location.`,
    improperArguments:                  `${ERRORHEADER}Improper arguments`
}
/** Default sublocation for decompressing files under the current directory. */
const DEFAULTDECOMPRESSSUBLOCATION = "officeDist";
/** Location for decompressing files. Default is "officeDist" */
let decompressSubLocation = DEFAULTDECOMPRESSSUBLOCATION;
/** Flag to output errors to console other than normal error handling. Default is false as we anyway push the message for error handling. */
let outputErrorToConsole = false;

/** Console error if allowed
 * @param {string} errorMessage Error message to show on the console
 * @returns {void}
 */
function consoleError(errorMessage) {
    if (outputErrorToConsole)
        console.error(errorMessage);
}

/** Custom parseString promise as the native has bugs
 * @param {string} xml The xml string from the doc file
 * @param {boolean} [ignoreAttrs=true] Optional: Ignore attributes part of xml and focus only on the content.
 * @returns {Promise<string>}
 */
const parseStringPromise = (xml, ignoreAttrs = true) => new Promise((resolve, reject) => {
    xml2js.parseString(xml, { "ignoreAttrs": ignoreAttrs }, (err, result) => {
        if (err)
            reject(err);
        resolve(result);
    });
});


/** Main function for parsing text from word files
 * @param {string} filename File path
 * @param {function} callback Callback function that returns value or error
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {void}
 */
function parseWord(filename, callback, deleteOfficeDist = true) {
    if (!fs.existsSync(filename)) {
        consoleError(ERRORMSG.fileDoesNotExist(filename));
        return callback(undefined, ERRORMSG.fileDoesNotExist(filename));
    }
    const ext = filename.split(".").pop().toLowerCase();
    if (ext != 'docx') {
        consoleError(ERRORMSG.extensionUnsupported(extension));
        return callback(undefined, ERRORMSG.extensionUnsupported(ext));
    }

    /** Store all the text content to respond */
    let responseText = [];

    /** Extracting text from Word files xml objects converted to js */
    function extractTextFromWordXmlObjects(xmlObjects) {
        // specifically for Arrays
        if (Array.isArray(xmlObjects)) {
            xmlObjects.forEach(item =>
                (typeof item == "string") && (item != "")
                    ? responseText.push(item)
                    : extractTextFromWordXmlObjects(item))
        }
        // for other JS Object
        else if (typeof xmlObjects == "object") {
            for (const [key, value] of Object.entries(xmlObjects)) {
                (typeof value == "string") || (typeof value[0] == "string")
                    ? (key == "w:t" || key == "_") && value != ""
                        ? responseText.push(value)
                        : undefined
                    : extractTextFromWordXmlObjects(value);
            }
        }
    }

    const contentFile = 'word/document.xml';
    decompress(filename,
        decompressSubLocation,
        { filter: x => x.path == contentFile }
    )
    .then(files => {
        if (files.length != 1) {
            consoleError(ERRORMSG.fileCorrupted(filename));
            return callback(undefined, ERRORMSG.fileCorrupted(filename));
        }

        return fs.readFileSync(`${decompressSubLocation}/${contentFile}`, 'utf8');
    })
    .then(xmlContent => parseStringPromise(xmlContent))
    .then(xmlObjects => {
        extractTextFromWordXmlObjects(xmlObjects);
        const returnCallbackPromise = new Promise((res, rej) => {
            if (deleteOfficeDist)
                rimraf(decompressSubLocation, err => {
                    if (err)
                        consoleError(err);
                    res();
                });
            else
                res();
        });

        returnCallbackPromise
        .then(() => callback(responseText.join(" "), undefined));

    })
    .catch(error => {
        consoleError(error)
        return callback(undefined, error);
    });
}

/** Main function for parsing text from PowerPoint files
 * @param {string} filename File path
 * @param {function} callback Callback function that returns value or error
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {void}
 */
function parsePowerPoint(filename, callback, deleteOfficeDist = true) {
    if (!fs.existsSync(filename)) {
        consoleError(ERRORMSG.fileDoesNotExist(filename));
        return callback(undefined, ERRORMSG.fileDoesNotExist(filename));
    }
    const ext = filename.split(".").pop().toLowerCase();
    if (ext != 'pptx') {
        consoleError(ERRORMSG.extensionUnsupported(extension));
        return callback(undefined, ERRORMSG.extensionUnsupported(ext));
    }

    /** Store all the text content to respond */
    let responseText = [];

    /** Extracting text from powerpoint files xml objects converted to js */
    function extractTextFromPowerPointXmlObjects(xmlObjects) {
        // specifically for Arrays
        if (Array.isArray(xmlObjects)) {
            xmlObjects.forEach(item =>
                (typeof item == "string") && (item != "")
                    ? responseText.push(item)
                    : extractTextFromPowerPointXmlObjects(item))
        }
        // for other JS Object
        else if (typeof xmlObjects == "object") {
            for (const [key, value] of Object.entries(xmlObjects)) {
                (typeof value == "string") || (typeof value[0] == "string")
                    ? (key == "a:t" || key == "_") && value != ""
                        ? responseText.push(value)
                        : undefined
                    : extractTextFromPowerPointXmlObjects(value);
            }
        }
    }

    // Files regex that hold our content of interest
    const contentFiles = [
        /ppt\/slides\/slide\d+.xml/g,
        /ppt\/notesSlides\/notesSlide\d+.xml/g
    ]

    decompress(filename,
        decompressSubLocation,
        { filter: x => contentFiles.findIndex(fileRegex => x.path.match(fileRegex)) > -1 }
    )
    .then(files => {
        // Sort files according to previous order of taking text out of ppt/slides followed by ppt/notesSlides
        files.sort((a,b) => contentFiles.findIndex(fileRegex => a.path.match(fileRegex)) -  contentFiles.findIndex(fileRegex => b.path.match(fileRegex)))

        if (files.length == 0) {
            consoleError(ERRORMSG.fileCorrupted(filename));
            return callback(undefined, ERRORMSG.fileCorrupted(filename));
        }

        // Returning an array of all the xml contents read using fs.readFileSync
        return files.map(file => fs.readFileSync(`${decompressSubLocation}/${file.path}`, 'utf8'))
    })
    .then(xmlContentArray => Promise.all(xmlContentArray.map(xmlContent => parseStringPromise(xmlContent))))    // Returning an array of all parseStringPromise responses
    .then(xmlObjectsArray => {
        xmlObjectsArray.forEach(xmlObjects => extractTextFromPowerPointXmlObjects(xmlObjects)); // Extracting text from all xml js objects with our conditions

        const returnCallbackPromise = new Promise((res, rej) => {
            if (deleteOfficeDist)
                rimraf(decompressSubLocation, err => {
                    if (err)
                        consoleError(err);
                    res();
                });
            else
                res();
        });

        returnCallbackPromise
        .then(() => callback(responseText.join(" "), undefined));

    })
    .catch(error => {
        consoleError(error)
        return callback(undefined, error);
    });
}

/** Main function for parsing text from Excel files
 * @param {string} filename File path
 * @param {function} callback Callback function that returns value or error
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {void}
 */
function parseExcel(filename, callback, deleteOfficeDist = true) {
    if (!fs.existsSync(filename)) {
        consoleError(ERRORMSG.fileDoesNotExist(filename));
        return callback(undefined, ERRORMSG.fileDoesNotExist(filename));
    }
    const ext = filename.split(".").pop().toLowerCase();
    if (ext != 'xlsx') {
        consoleError(ERRORMSG.extensionUnsupported(extension));
        return callback(undefined, ERRORMSG.extensionUnsupported(ext));
    }

    /** Store all the text content to respond */
    let responseText = [];

    function extractTextFromExcelXmlObjects2dArray(xmlObjects2dArray) {
        xmlObjects2dArray[0].map(xmlObjects => extractTextFromExcelXmlObjects(xmlObjects, 0));
        xmlObjects2dArray[1].map(xmlObjects => extractTextFromExcelXmlObjects(xmlObjects, 1));
        xmlObjects2dArray[2].map(xmlObjects => extractTextFromExcelXmlObjects(xmlObjects, 2));
    }

    /** Extracting text from Excel files xml objects converted to js */
    function extractTextFromExcelXmlObjects(xmlObjects, contentFilesIndex) {
        switch(contentFilesIndex) {
            case 0: {       // worksheet
                // specifically for Arrays
                if (Array.isArray(xmlObjects)) {
                    xmlObjects.forEach(item =>
                        item["v"]
                            ? ((item["$"]["t"] != "s"))
                                ? responseText.push(item["v"][0])
                                : undefined
                            : extractTextFromExcelXmlObjects(item, contentFilesIndex))
                }
                // for other JS Object
                else if (typeof xmlObjects == "object") {
                    for (const [key, value] of Object.entries(xmlObjects)) {
                        value["v"]
                            ? ((value["$"]["t"] == "s"))
                                ? responseText.push(value["v"][0])
                                : undefined
                            : extractTextFromExcelXmlObjects(value, contentFilesIndex);
                    }
                }
                break;
            }
            case 1: {       // sharedStrings
                // specifically for Arrays
                if (Array.isArray(xmlObjects)) {
                    xmlObjects.forEach(item =>
                        (typeof item == "string") && (item != "")
                            ? responseText.push(item)
                            : extractTextFromExcelXmlObjects(item, contentFilesIndex))
                }
                // for other JS Object
                else if (typeof xmlObjects == "object") {
                    for (const [key, value] of Object.entries(xmlObjects)) {
                        (typeof value == "string") || (typeof value[0] == "string")
                            ? (key == "t" || key == "_") && (value != "")
                                ? responseText.push(value)
                                : undefined
                            : extractTextFromExcelXmlObjects(value, contentFilesIndex);
                    }
                }
                break;
            }
            case 2: {       // drawings
                // specifically for Arrays
                if (Array.isArray(xmlObjects)) {
                    xmlObjects.forEach(item =>
                        (typeof item == "string") && (item != "")
                            ? responseText.push(item)
                            : extractTextFromExcelXmlObjects(item, contentFilesIndex))
                }
                // for other JS Object
                else if (typeof xmlObjects == "object") {
                    for (const [key, value] of Object.entries(xmlObjects)) {
                        (typeof value == "string") || (typeof value[0] == "string")
                            ? (key == "a:t" || key == "_") && (value != "")
                                ? responseText.push(value)
                                : undefined
                            : extractTextFromExcelXmlObjects(value, contentFilesIndex);
                    }
                }
                break;
            }
        }
    }

    // Files regex that hold our content of interest
    const contentFiles = [
        /xl\/worksheets\/sheet\d+.xml/g,
        /xl\/sharedStrings.xml/g,
        /xl\/drawings\/drawing\d+.xml/g,
    ]

    decompress(filename,
        decompressSubLocation,
        { filter: x => contentFiles.findIndex(fileRegex => x.path.match(fileRegex)) > -1 }
    )
    .then(files => {
        // arrange files into 2d array of files organized in contentFiles order, separated by array elements
        const files2dArray = [];
        contentFiles.forEach(fileRegex => files2dArray.push(files.filter(file => file.path.match(fileRegex))))

        if (files.length == 0) {
            consoleError(ERRORMSG.fileCorrupted(filename));
            return callback(undefined, ERRORMSG.fileCorrupted(filename));
        }

        // Returning a 2dArray of all the xml contents read using fs.readFileSync and separated by array elements
        return files2dArray.map(files => files.map(file => fs.readFileSync(`${decompressSubLocation}/${file.path}`, 'utf8')))
    })
    .then(xmlContent2dArray => Promise.all(xmlContent2dArray.map(xmlContentArray => Promise.all(xmlContentArray.map(xmlContent => parseStringPromise(xmlContent, false))))))    // Returning a 2dArray of all parseStringPromise responses
    .then(xmlObjects2dArray => {
        extractTextFromExcelXmlObjects2dArray(xmlObjects2dArray); // Extracting text from all xml js objects with our conditions

        const returnCallbackPromise = new Promise((res, rej) => {
            if (deleteOfficeDist)
                rimraf(decompressSubLocation, err => {
                    if (err)
                        consoleError(err);
                    res();
                });
            else
                res();
        });

        returnCallbackPromise
        .then(() => callback(responseText.join(" "), undefined));

    })
    .catch(error => {
        consoleError(error)
        return callback(undefined, error);
    });
}


/** Main function for parsing text from open office files
 * @param {string} filename File path
 * @param {function} callback Callback function that returns value or error
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {void}
 */
function parseOpenOffice(filename, callback, deleteOfficeDist = true) {
    if (!fs.existsSync(filename)) {
        consoleError(ERRORMSG.fileDoesNotExist(filename));
        return callback(undefined, ERRORMSG.fileDoesNotExist(filename));
    }
    const ext = filename.split(".").pop().toLowerCase();
    if (!["odt", "odp", "ods"].includes(ext)) {
        consoleError(ERRORMSG.extensionUnsupported(extension));
        return callback(undefined, ERRORMSG.extensionUnsupported(ext));
    }

    /** Store all the text content to respond */
    let responseText = [];
    /** Extracting text from Open Office files xml objects converted to js */
    function extractTextFromOpenOfficeXmlObjects(xmlObjects) {
        // specifically for Arrays
        if (Array.isArray(xmlObjects)) {
            xmlObjects.forEach(item =>
                (typeof item == "string") && (item != "")
                    ? responseText.push(item)
                    : extractTextFromOpenOfficeXmlObjects(item))
        }
        // for other JS Object
        else if (typeof xmlObjects == "object") {
            for (const [key, value] of Object.entries(xmlObjects)) {
                typeof value == "string"
                    ? value != ""
                        ? responseText.push(value)
                        : undefined
                    : extractTextFromOpenOfficeXmlObjects(value);
            }
        }
    }

    const contentFile = 'content.xml';
    decompress(filename,
        decompressSubLocation,
        { filter: x => x.path == contentFile }
    )
    .then(files => {
        if (files.length != 1) {
            consoleError(ERRORMSG.fileCorrupted(filename));
            return callback(undefined, ERRORMSG.fileCorrupted(filename));
        }

        return fs.readFileSync(`${decompressSubLocation}/${contentFile}`, 'utf8');
    })
    .then(xmlContent => parseStringPromise(xmlContent))
    .then(xmlObjects => {
        extractTextFromOpenOfficeXmlObjects(xmlObjects);
        const returnCallbackPromise = new Promise((res, rej) => {
            if (deleteOfficeDist)
                rimraf(decompressSubLocation, err => {
                    if (err)
                        consoleError(err);
                    res();
                });
            else
                res();
        });

        returnCallbackPromise
        .then(() => callback(responseText.join(" "), undefined));

    })
    .catch(error => {
        consoleError(error)
        return callback(undefined, error);
    });
}


/** Main async function with callback to execute parseOffice for supported files
 * @param {string} filename File path
 * @param {function} callback Callback function that returns value or error
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {void}
 */
function parseOffice(filename, callback, deleteOfficeDist = true) {
    if (!fs.existsSync(filename)) {
        consoleError(ERRORMSG.fileDoesNotExist(filename));
        return callback(undefined, ERRORMSG.fileDoesNotExist(filename));
    }
    var extension = filename.split(".").pop().toLowerCase();

    switch(extension)
    {
        case "docx":
            parseWord(filename, (data, err) => callback(data, err), deleteOfficeDist);
            return;
        case "pptx":
            parsePowerPoint(filename, (data, err) => callback(data, err), deleteOfficeDist);
            return;
        case "xlsx":
            parseExcel(filename, (data, err) => callback(data, err), deleteOfficeDist);
            return;
        case "odt":
        case "odp":
        case "ods":
            parseOpenOffice(filename, (data, err) => callback(data, err), deleteOfficeDist);
            return;

        default:
            consoleError(ERRORMSG.extensionUnsupported(extension));
            callback(undefined, ERRORMSG.extensionUnsupported(extension));
    }
}

/**
 * Set decompression directory. The final decompressed data will be put inside officeDist folder within your directory
 * @param {string} newLocation Relative path to the directory that will contain officeDist folder with decompressed data
 * @returns {void}
 */
function setDecompressionLocation(newLocation) {
    if (newLocation != undefined) {
        newLocation = `${newLocation}${newLocation.endsWith('/') ? '' : '/'}${DEFAULTDECOMPRESSSUBLOCATION}`

        if (fs.existsSync(newLocation))
            decompressSubLocation = newLocation;
        return;
    }
    consoleError(ERRORMSG.locationNotFound(newLocation));
    decompressSubLocation = DEFAULTDECOMPRESSSUBLOCATION;
}

/** Enable console output
 * @returns {void}
 */
function enableConsoleOutput() {
    outputErrorToConsole = true;
}

/** Disabled console output
 * @returns {void}
 */
function disableConsoleOutput() {
    outputErrorToConsole = false;
}


// #region Promise versions of above functions

/** Async function that can be used with await to execute parseWord. Or it can be used with promises.
 * @param {string} filename File path
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {void}
 */
var parseWordAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parseWord(filename, function (data, error) {
                if (error)
                    return reject(error);
                return resolve(data);
            }, deleteOfficeDist);
        }
        catch (error) {
            return reject(error);
        }
    })
}

/** Async function that can be used with await to execute parsePowerPoint. Or it can be used with promises.
 * @param {string} filename File path
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {void}
 */
var parsePowerPointAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parsePowerPoint(filename, function (data, err) {
                if (err)
                    return reject(err);
                return resolve(data);
            }, deleteOfficeDist);
        }
        catch (error) {
            return reject(error);
        }
    })
}

/** Async function that can be used with await to execute parseExcel. Or it can be used with promises.
 * @param {string} filename File path
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {void}
 */
var parseExcelAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parseExcel(filename, function (data, err) {
                if (err)
                    return reject(err);
                return resolve(data);
            }, deleteOfficeDist);
        }
        catch (error) {
            return reject(error);
        }
    })
}

/** Async function that can be used with await to execute parseOpenOffice. Or it can be used with promises.
 * @param {string} filename File path
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {void}
 */
var parseOpenOfficeAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parseOpenOffice(filename, function (data, err) {
                if (err)
                    return reject(err);
                return resolve(data);
            }, deleteOfficeDist);
        }
        catch (error) {
            return reject(error);
        }
    })
}

/**
 * Main async function that can be used with await to execute parseOffice. Or it can be used with promises.
 * @param {string} filename File path
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {void}
 */
var parseOfficeAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parseOffice(filename, function (data, err) {
                if (err)
                    return reject(err);
                return resolve(data);
            }, deleteOfficeDist);
        }
        catch (error) {
            return reject(error);
        }
    })
}
// #endregion Async Versions

module.exports.parseWord = parseWord;
module.exports.parsePowerPoint = parsePowerPoint;
module.exports.parseExcel = parseExcel;
module.exports.parseOpenOffice = parseOpenOffice;
module.exports.parseOffice = parseOffice;
module.exports.parseWordAsync = parseWordAsync;
module.exports.parsePowerPointAsync = parsePowerPointAsync;
module.exports.parseExcelAsync = parseExcelAsync;
module.exports.parseOpenOfficeAsync = parseOpenOfficeAsync;
module.exports.parseOfficeAsync = parseOfficeAsync;
module.exports.setDecompressionLocation = setDecompressionLocation;
module.exports.enableConsoleOutput = enableConsoleOutput;
module.exports.disableConsoleOutput = disableConsoleOutput;


// Run this library on CLI
if ((process.argv[0].split('/').pop() == "node" || process.argv[0].split('/').pop() == "npx") && (process.argv[1].split('/').pop() == "officeParser.js" || process.argv[1].split('/').pop() == "officeparser")) {
    if (process.argv.length == 2) {
        // continue
    }
    else if (process.argv.length == 3)
        parseOfficeAsync(process.argv[2])
        .then(text => console.log(text))
        .catch(error => console.error(error))
    else
        console.error(ERRORMSG.improperArguments)
}
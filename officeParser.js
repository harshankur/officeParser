const decompress = require('decompress');
const xml2js = require('xml2js')
const fs = require('fs')
const util = require('util');
const rimraf = require('rimraf');

/** Header for error messages */
const ERRORHEADER = "[OfficeParser]: ";
/** Error messages */
const ERRORMSG = {
    extensionUnsupported: (ext) => `${ERRORHEADER}Sorry, we currently support docx, pptx, xlsx, odt, odp, ods files only. Create a ticket in Issues on github to add support for ${ext} files. Stay tuned for further updates.`,
    fileCorrupted: (filename) => `${ERRORHEADER}Your file ${filename} seems to be corrupted. If you are sure it is fine, please create a ticket in Issues on github with the file to reproduce error.`,
    fileDoesNotExist: (filename) => `${ERRORHEADER}File ${filename} could not be found! Check if the file exists or verify if the relative path to the file is correct from your terminal's location.`,
}
/** Location for decompressing files. Default is "officeDist" */
let decompressLocation = "officeDist";
/** Flag to output errors to console other than normal error handling. Default is false as we anyway push the message for error handling. */
let outputErrorToConsole = false;

/** Console error if allowed */
function consoleError(error) {
    if (outputErrorToConsole)
        console.error(error);
}

/** Custom parseString promise as the native has bugs */
const parseStringPromise = xml => new Promise((resolve, reject) => {
    xml2js.parseString(xml, { "ignoreAttrs": true }, (err, result) => {
        if (err)
            reject(err);
        resolve(result);
    });
});




/** Main function for parsing text from word files */
function parseWord(filename, callback, deleteOfficeDist = true) {
    if (!fs.existsSync(filename))
        return callback(undefined, ERRORMSG.fileDoesNotExist(filename))
    const ext = filename.split(".").pop().toLowerCase();
    if (ext != 'docx')
        return callback(undefined, ERRORMSG.extensionUnsupported(ext));

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
                typeof value == "string"
                    ? (key == "w:t" || key == "_") && value != ""
                        ? responseText.push(value)
                        : undefined
                    : extractTextFromWordXmlObjects(value);
            }
        }
    }

    const contentFile = 'word/document.xml';
    decompress(filename,
        decompressLocation,
        { filter: x => x.path == contentFile }
    )
    .then(files => {
        if (files.length != 1)
            return callback(undefined, ERRORMSG.fileCorrupted(filename));

        return fs.readFileSync(`${decompressLocation}/${contentFile}`, 'utf8');
    })
    .then(xmlContent => parseStringPromise(xmlContent))
    .then(xmlObjects => {
        extractTextFromWordXmlObjects(xmlObjects);
        const returnCallbackPromise = new Promise((res, rej) =>
        {
            if (deleteOfficeDist)
                rimraf(decompressLocation, err => res(consoleError(err)));
            else
                res();
        })

        returnCallbackPromise
        .then(() => callback(responseText.join(" "), undefined));

    })
    .catch(error => {
        consoleError(error)
        return callback(undefined, error);
    });
}

/** Main function for parsing text from PowerPoint files */
function parsePowerPoint(filename, callback, deleteOfficeDist = true) {
    if (!fs.existsSync(filename))
        return callback(undefined, ERRORMSG.fileDoesNotExist(filename))
    const ext = filename.split(".").pop().toLowerCase();
    if (ext != 'pptx')
        return callback(undefined, ERRORMSG.extensionUnsupported(ext));

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
                typeof value == "string"
                    ? (key == "a:t" || key == "_") && value != ""
                        ? responseText.push(value)
                        : undefined
                    : extractTextFromPowerPointXmlObjects(value);
            }
        }
    }

    // Files that hold our content of interest
    const contentFiles = [
        {
            folder: "ppt/slides",
            fileExtension: "xml"
        },
        {
            folder: "ppt/notesSlides",
            fileExtension: "xml"
        }
    ];

    decompress(filename,
        decompressLocation,
        // We are looking for files that have same folder (starting substring) as our content files and have same fileExtensions
        { filter: x => contentFiles.findIndex(file => x.path.indexOf(file.folder) == 0 && x.path.split('.').pop() == file.fileExtension) > -1 }
    )
    .then(files => {
        // Sort files according to previous order of taking text out of ppt/slides followed by ppt/notesSlides
        files.sort((a,b) => contentFiles.findIndex(file => a.path.indexOf(file.folder) == 0) -  contentFiles.findIndex(file => b.path.indexOf(file.folder) == 0))

        if (files.length == 0)
            return callback(undefined, ERRORMSG.fileCorrupted(filename));

        // Returning a promise that resolves after all the xml contents have been read using fs.readFileSync
        return allTextReadPromise = new Promise((resolve, reject) =>
        {
            let xmlContentArray = [];
            for (let file of files)
            {
                xmlContentArray.push(fs.readFileSync(`${decompressLocation}/${file.path}`, 'utf8'))
            }
            resolve(xmlContentArray);
        })
    })
    .then(xmlContentArray => Promise.all(xmlContentArray.map(xmlContent => parseStringPromise(xmlContent))))    // Returning an array of all parseStringPromise responses
    .then(xmlObjectsArray => {
        xmlObjectsArray.forEach(xmlObjects => extractTextFromPowerPointXmlObjects(xmlObjects)); // Extracting text from all xml js objects with our conditions

        const returnCallbackPromise = new Promise((res, rej) =>
        {
            if (deleteOfficeDist)
                rimraf(decompressLocation, err => res(consoleError(err)));
            else
                res();
        })

        returnCallbackPromise
        .then(() => callback(responseText.join(" "), undefined));

    })
    .catch(error => {
        consoleError(error)
        return callback(undefined, error);
    });
}


// #region textFetchFromExcel

var myTextExcel = [];

async function scanForTextExcelDrawing(result) {
    if (Array.isArray(result)) {
        for (var i = 0; i < result.length; i++) {
            if (typeof(result[i]) == "string" && result[i] != "") {
                await myTextExcel.push(result[i]);
            }
            else {
                await scanForTextExcelDrawing(result[i]);
            }
        }
        return;
    }
    else if (typeof(result) == "object") {
        for (var property in result) {
            if (result.hasOwnProperty(property)) {
                if (typeof(result[property]) == "string") {
                    if ((property == "a:t" || property == "_") && result[property] != "") {
                        await myTextExcel.push(result[property]);
                    }
                }
                else if (typeof(result[property][0]) == "string") {
                    if ((property == "a:t" || property == "_") && result[property] != "") {
                        await myTextExcel.push(result[property]);
                    }
                }
                else {
                    await scanForTextExcelDrawing(result[property]);
                }
            }
        }
        return;
    }
}


async function scanForTextExcelSharedStrings(result) {
    if (Array.isArray(result)) {
        for (var i = 0; i < result.length; i++) {
            if (typeof(result[i]) == "string" && result[i] != "") {
                await myTextExcel.push(result[i]);
            }
            else {
                await scanForTextExcelSharedStrings(result[i]);
            }
        }
        return;
    }
    else if(typeof(result) == "object") {
        for (var property in result) {
            if (result.hasOwnProperty(property)) {
                if (typeof(result[property]) == "string") {
                    if ((property == "t" || property == "_") && result[property] != "") {
                        await myTextExcel.push(result[property]);
                    }
                }
                else if (typeof(result[property][0]) == "string") {
                    if ((property == "t" || property == "_") && result[property] != "") {
                        await myTextExcel.push(result[property]);
                    }
                }
                else {
                    await scanForTextExcelSharedStrings(result[property]);
                }
            }
        }
        return;
    }
}

async function scanForTextExcelWorkSheet(result) {
    if (Array.isArray(result)) {
        for (var i = 0; i < result.length; i++) {
            if (result[i]["v"]) {
                if ((result[i]["$"]["t"] != "s")) {
                    await myTextExcel.push(result[i]["v"][0]);
                }
            }
            else {
                await scanForTextExcelWorkSheet(result[i]);
            }
        }
        return;
    }
    else if(typeof(result) == "object") {
        for (var property in result) {
            if (result.hasOwnProperty(property)) {
                if (result[property]["v"]) {
                    if ((result[property]["$"]["t"] == "s")) {
                        await myTextExcel.push(result[property]["v"][0]);
                    }
                }
                else {
                    await scanForTextExcelWorkSheet(result[property]);
                }
            }
        }
        return;
    }
}

var parseExcel = function (filename, callback, deleteOfficeDist = true) {
    if (validateFileExtension(filename, ["xlsx"])) {
        try {
            decompress(filename, decompressLocation).then(async files => {
                myTextExcel = [];
    
    
                if (fs.existsSync(decompressLocation + '/xl/worksheets')) {
                    var workSheetsNum = await fs.readdirSync(decompressLocation + '/xl/worksheets').length;
                    for (var i = 0; i < workSheetsNum - 1; i++) {
                        var parser = new xml2js.Parser();
    
                        var myData = await util.promisify(fs.readFile)(`${decompressLocation}/xl/worksheets/sheet${i + 1}.xml`, 'utf8');
                        var result = await util.promisify(parser.parseString.bind(parser))(myData);
                        await scanForTextExcelWorkSheet(result);
                    }
    
                    if (fs.existsSync(decompressLocation + '/xl/sharedStrings.xml')) {
                        var parser = new xml2js.Parser();
    
                        var myData = await util.promisify(fs.readFile)(`${decompressLocation}/xl/sharedStrings.xml`, 'utf8');
                        var result = await util.promisify(parser.parseString.bind(parser))(myData);
                        await scanForTextExcelSharedStrings(result);
                    }
    
                    if (fs.existsSync(decompressLocation + '/xl/drawings')) {
                        var drawingsNum = await fs.readdirSync(decompressLocation + '/xl/drawings').length;
    
                        for (var i = 0; i < drawingsNum; i++) {
                            var parser = new xml2js.Parser();
    
                            var myData = await util.promisify(fs.readFile)(`${decompressLocation}/xl/drawings/drawing${i + 1}.xml`, 'utf8');
                            var result = await util.promisify(parser.parseString.bind(parser))(myData);
                            await scanForTextExcelDrawing(result);
                        }
                    }
    
                    callback(myTextExcel.join(" "), undefined);
                    
                    if (deleteOfficeDist == true) {
                        rimraf(decompressLocation, function () {});
                    }
                }
                else {
                    if (deleteOfficeDist == true) {
                        rimraf(decompressLocation, function () {});
                    }
                }
            })
            .catch(function (err) {
                if (outputErrorToConsole) console.log(err);
                return callback(undefined, err);
            });
        }
        catch (err) {
            if (outputErrorToConsole) console.log(err);
            return callback(undefined, err);
        }
    }
} 

// #endregion textFetchFromExcel



/** Main function for parsing text from open office files */
function parseOpenOffice(filename, callback, deleteOfficeDist = true) {
    if (!fs.existsSync(filename))
        return callback(undefined, ERRORMSG.fileDoesNotExist(filename))
    const ext = filename.split(".").pop().toLowerCase();
    if (!["odt", "odp", "ods"].includes(ext))
        return callback(undefined, ERRORMSG.extensionUnsupported(ext));

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
        decompressLocation,
        { filter: x => x.path == contentFile }
    )
    .then(files => {
        if (files.length != 1)
            return callback(undefined, ERRORMSG.fileCorrupted(filename));

        return fs.readFileSync(`${decompressLocation}/${contentFile}`, 'utf8');
    })
    .then(xmlContent => parseStringPromise(xmlContent))
    .then(xmlObjects => {
        extractTextFromOpenOfficeXmlObjects(xmlObjects);
        const returnCallbackPromise = new Promise((res, rej) =>
        {
            if (deleteOfficeDist)
                rimraf(decompressLocation, err => res(consoleError(err)));
            else
                res();
        })

        returnCallbackPromise
        .then(() => callback(responseText.join(" "), undefined));

    })
    .catch(error => {
        consoleError(error)
        return callback(undefined, error);
    });
}



// #region validate file names

/**
 * Legacy method to check whether file functions called are indeed applicable with files provided
 * @param {*} filename file path
 * @param {*} extension extensions of supported files
 */
function validateFileExtension(filename, extension) {
    for (var extensionIterator = 0; extensionIterator < extension.length; extensionIterator++) {
        if (extension[extensionIterator] == filename.split(".").pop().toLowerCase()) {
            return true;
        }
    }
    console.log(`[OfficeParser]: Sorry, we currently support docx, pptx, xlsx, odt, odp, ods files only. Make sure you pass appropriate file with the parsing functions.`);
    return false;
}

// #endregion validate file names

// #region parse office

function parseOffice(filename, callback, deleteOfficeDist = true) {
    var extension = filename.split(".").pop().toLowerCase();

    if (extension == "docx") {
        parseWord(filename, function (data) {
            callback(data);
        }, deleteOfficeDist);
    }
    else if (extension == "pptx") {
        parsePowerPoint(filename, function (data) {
            callback(data);
        }, deleteOfficeDist);
    }
    else if (extension == "xlsx") {
        parseExcel(filename, function (data) {
            callback(data);
        }, deleteOfficeDist);
    }
    else if (extension == "odt" || extension == "odp" || extension == "ods") {
        parseOpenOffice(filename, function (data) {
            callback(data);
        }, deleteOfficeDist);
    }
    else {
        var errMessage = `[OfficeParser]: Sorry, we currently support docx, pptx, xlsx, odt, odp, ods files only. Create a ticket in Issues on github to add support for ${extension} files. Stay tuned for further updates.`;
        if (outputErrorToConsole) console.log(errMessage);
        return callback(undefined, errMessage);
    }
    
}

// #endregion parse office


// #region setDecompressionLocation
function setDecompressionLocation(newLocation) {
    if (newLocation != undefined) {
        decompressLocation = newLocation + "/officeDist";
    }
    else {
        decompressLocation = "officeDist";
    }
}

// #endregion setDecompressionLocation

// #region setConsoleOutput
function enableConsoleOutput() {
    outputErrorToConsole = true;
}

function disableConsoleOutput() {
    outputErrorToConsole = false;
}

// #endregion setConsoleOutput

// #region Async Versions
var parseWordAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        parseWord(filename, function (data, error) {
            if (error)
                return reject(error);
            return resolve(data);
        }, deleteOfficeDist);
    })
}

var parsePowerPointAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parsePowerPoint(filename, function (data, err) {
                if (err) return reject(err);
                return resolve(data);
            },deleteOfficeDist);
        } catch (error) {
            return reject(error);
        }
    })
}

var parseExcelAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parseExcel(filename, function (data, err) {
                if (err) return reject(err);
                return resolve(data);
            },deleteOfficeDist);
        } catch (error) {
            return reject(error);
        }
    })
}

var parseOpenOfficeAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parseOpenOffice(filename, function (data, err) {
                if (err) return reject(err);
                return resolve(data);
            },deleteOfficeDist);
        } catch (error) {
            return reject(error);
        }
    })
}

var parseOfficeAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parseOffice(filename, function (data, err) {
                if (err) return reject(err);
                return resolve(data);
            },deleteOfficeDist);
        } catch (error) {
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

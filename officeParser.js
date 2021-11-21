const decompress = require('decompress');
const xml2js = require('xml2js')
const fs = require('fs')
const util = require('util');
const rimraf = require('rimraf');

// #region textFetchFromWord

var myTextWord = [];
var decompressLocation = "officeDist";
var outputToConsoleWhenRequired = true;     // Default is set to true to let the user identify source of an error in node terminal

async function scanForTextWord(result) {
    if (Array.isArray(result)) {
        for (var i = 0; i < result.length; i++) {
            if (typeof(result[i]) == "string" && result[i] != "") {
                await myTextWord.push(result[i]);
            }
            else {
                await scanForTextWord(result[i]);
            }
        }
        return;
    }
    else if (typeof(result) == "object") {
        for (var property in result) {
            if (result.hasOwnProperty(property)) {
                if (typeof(result[property]) == "string") {
                    if ((property == "w:t" || property == "_") && result[property] != "") {
                        await myTextWord.push(result[property]);
                    }
                }
                else {
                    await scanForTextWord(result[property]);
                }
            }
        }
        return;
    }
}

var parseWord = async function (filename, callback, deleteOfficeDist = true) {
    if (validateFileExtension(filename, ["docx"])) {
        try {
            decompress(filename, decompressLocation).then(files => {
                myTextWord = [];
                if (fs.existsSync(decompressLocation + "/word/document.xml")) {
                    fs.readFile(decompressLocation + '/word/document.xml', 'utf8', function (err,data) {
                        if (err) {
                            if (outputToConsoleWhenRequired) console.log(err);
                            return callback(undefined, err);
                        }
                        xml2js.parseString(data, async function (err, result) {
                            await scanForTextWord(result);
    
                            callback(myTextWord.join(" "), undefined);
                            if (deleteOfficeDist == true) {
                                rimraf(decompressLocation, function () {});
                            }
                        });
                    });
                }
                else {
                    if (deleteOfficeDist == true) {
                        rimraf(decompressLocation, function () {});
                    }
                }
            })
            .catch(function (err) {
                if (outputToConsoleWhenRequired) console.log(err);
                return callback(undefined, err);
            });
        }
        catch(err) {
            if (outputToConsoleWhenRequired) console.log(err);
            return callback(undefined, err);
        }
    }
}

// #endregion textFetchFromWord



// #region textFetchFromPowerPoint

var myTextPowerPoint = [];

async function scanForTextPowerPoint(result) {
    if (Array.isArray(result)) {
        for (var i = 0; i < result.length; i++) {
            if (typeof(result[i]) == "string" && result[i] != "") {
                await myTextPowerPoint.push(result[i]);
            }
            else {
                await scanForTextPowerPoint(result[i]);
            }
        }
        return;
    }
    else if (typeof(result) == "object") {
        for (var property in result) {
            if (result.hasOwnProperty(property)) {
                if (typeof(result[property]) == "string") {
                    if ((property == "a:t" || property == "_") && result[property] != "") {
                        await myTextPowerPoint.push(result[property]);
                    }
                }
                else if (typeof(result[property][0]) == "string") {
                    if ((property == "a:t" || property == "_") && result[property] != "") {
                        await myTextPowerPoint.push(result[property]);
                    }
                }
                else {
                    await scanForTextPowerPoint(result[property]);
                }
            }
        }
        return;
    }
}

var parsePowerPoint = function (filename, callback, deleteOfficeDist = true) {
    if (validateFileExtension(filename, ["pptx"])) {
        try {
            decompress(filename, decompressLocation).then(async files => {
                myTextPowerPoint = [];
    
                if (fs.existsSync(decompressLocation + '/ppt/slides')) {
                    var slidesNum = await fs.readdirSync(decompressLocation + '/ppt/slides').length;
    
                    for (var i = 0; i < slidesNum - 1; i++) {
                        var parser = new xml2js.Parser();
    
                        var myData = await util.promisify(fs.readFile)(`${decompressLocation}/ppt/slides/slide${i + 1}.xml`, 'utf8');
                        var result = await util.promisify(parser.parseString.bind(parser))(myData);
                        await scanForTextPowerPoint(result);
                    }
    
                    if (fs.existsSync(decompressLocation + '/ppt/notesSlides')) {
                        var notesSlidesNum = await fs.readdirSync(decompressLocation + '/ppt/notesSlides').length;
                        for (var i = 0; i < notesSlidesNum - 1; i++) {
                            var parser = new xml2js.Parser();
        
                            var myData = await util.promisify(fs.readFile)(`${decompressLocation}/ppt/notesSlides/notesSlide${i + 1}.xml`, 'utf8');
                            var result = await util.promisify(parser.parseString.bind(parser))(myData);
                            await scanForTextPowerPoint(result);
                        }
                    }
                    
                    callback(myTextPowerPoint.join(" "), undefined);
    
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
                if (outputToConsoleWhenRequired) console.log(err);
                return callback(undefined, err);
            });
        }
        catch (err) {
            if (outputToConsoleWhenRequired) console.log(err);
            return callback(undefined, err);
        }
    }
} 

// #endregion textFetchFromPowerPoint


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
                if (outputToConsoleWhenRequired) console.log(err);
                return callback(undefined, err);
            });
        }
        catch (err) {
            if (outputToConsoleWhenRequired) console.log(err);
            return callback(undefined, err);
        }
    }
} 

// #endregion textFetchFromExcel



// #region textFetchFromOpenOffice

var myTextOpenOffice = [];

async function scanForTextOpenOffice(result) {
    if (Array.isArray(result)) {
        for (var i = 0; i < result.length; i++) {
            if (typeof(result[i]) == "string" && result[i] != "") {
                await myTextOpenOffice.push(result[i]);
            }
            else {
                await scanForTextOpenOffice(result[i]);
            }
        }
        return;
    }
    else if (typeof(result) == "object") {
        for (var property in result) {
            if (result.hasOwnProperty(property)) {
                if (typeof(result[property]) == "string") {
                    if (result[property] != "") {
                        await myTextOpenOffice.push(result[property]);
                    }
                }
                else {
                    await scanForTextOpenOffice(result[property]);
                }
                
            }
        }
        return;
    }
}

var parseOpenOffice = async function (filename, callback, deleteOfficeDist = true) {
    if (validateFileExtension(filename, ["odt", "odp", "ods"])) {
        try {
            decompress(filename, decompressLocation).then(files => {
                myTextOpenOffice = [];
                if (fs.existsSync(decompressLocation + "/content.xml")) {
                    fs.readFile(decompressLocation + '/content.xml', 'utf8', function (err,data) {
                        if (err) {
                            if (outputToConsoleWhenRequired) console.log(err);
                            return callback(undefined, err);
                        }
                         
                        xml2js.parseString(data, {"ignoreAttrs": true}, async function (err, result) {
                            await scanForTextOpenOffice(result);
    
                            callback(myTextOpenOffice.join(" "), undefined);
                            if (deleteOfficeDist == true) {
                                rimraf(decompressLocation, function () {});
                            }
                        });
                    });
                }
                else {
                    if (deleteOfficeDist == true) {
                        rimraf(decompressLocation, function () {});
                    }
                }
            })
            .catch(function (err) {
                if (outputToConsoleWhenRequired) console.log(err);
                return callback(undefined, err);
            });
        }
        catch (err) {
            if (outputToConsoleWhenRequired) console.log(err);
            return callback(undefined, err);
        }
    }
}

// #endregion textFetchFromOpenOffice




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
        if (outputToConsoleWhenRequired) console.log(errMessage);
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
    outputToConsoleWhenRequired = true;
}

function disableConsoleOutput() {
    outputToConsoleWhenRequired = false;
}

// #endregion setConsoleOutput

// #region Async Versions
var parseWordAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parseWord(filename, function (data, err) {
                if (err) return reject(err);
                return resolve(data);
            },deleteOfficeDist);
        } catch (error) {
            return reject(error);
        }
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

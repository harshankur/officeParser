const decompress = require('decompress');
const xml2js = require('xml2js')
const fs = require('fs')
const util = require('util');
const rimraf = require('rimraf');

// #region textFetchFromWord

var myTextWord = [];
var decompressLocation = "officeDist";

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
    else if(typeof(result) == "object") {
        for (var property in result) {
            if (result.hasOwnProperty(property)) {
                if (typeof(result[property]) == "string"){
                    if ((property == "w:t" || property == "_" ) && result[property] != "") {
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

var parseWord = async function(filename, callback, deleteOfficeDist = true) {
    if (validateFileExtension(filename, ["docx"])) {
        decompress(filename, decompressLocation).then(files => {
            myTextWord = [];
            if (fs.existsSync(decompressLocation + "/word/document.xml")) {
                fs.readFile(decompressLocation + '/word/document.xml', 'utf8', function (err,data) {
                    if (err) {
                       return console.log(err);
                     }
                     
                     xml2js.parseString(data, async function (err, result) {
                        await scanForTextWord(result);
        
                        callback(myTextWord.join(" "));
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
        });
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
    else if(typeof(result) == "object") {
        for (var property in result) {
            if (result.hasOwnProperty(property)) {
                if (typeof(result[property]) == "string"){
                    if ((property == "a:t" || property == "_") && result[property] != "") {
                        await myTextPowerPoint.push(result[property]);
                    }
                }
                else if (typeof(result[property][0]) == "string"){
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


var parsePowerPoint = function(filename, callback, deleteOfficeDist = true) {
    if (validateFileExtension(filename, ["pptx"])) {
        decompress(filename, decompressLocation).then(async files => {
            myTextPowerPoint = [];

            if (fs.existsSync(decompressLocation + '/ppt/slides')) {
                var slidesNum = await fs.readdirSync(decompressLocation + '/ppt/slides').length;
                

                for (var i=0; i<slidesNum-1; i++) {
                    var parser = new xml2js.Parser();

                    var myData = await util.promisify(fs.readFile)(`${decompressLocation}/ppt/slides/slide${i+1}.xml`, 'utf8');
                    var result = await util.promisify(parser.parseString.bind(parser))(myData);
                    await scanForTextPowerPoint(result);
                }

                if (fs.existsSync(decompressLocation + '/ppt/notesSlides')) {
                    var notesSlidesNum = await fs.readdirSync(decompressLocation + '/ppt/notesSlides').length;
                    for (var i=0; i<notesSlidesNum-1; i++) {
                        var parser = new xml2js.Parser();
    
                        var myData = await util.promisify(fs.readFile)(`${decompressLocation}/ppt/notesSlides/notesSlide${i+1}.xml`, 'utf8');
                        var result = await util.promisify(parser.parseString.bind(parser))(myData);
                        await scanForTextPowerPoint(result);
                    }
                }
                
                callback(myTextPowerPoint.join(" "));
                    
                if (deleteOfficeDist == true) {
                    rimraf(decompressLocation, function () {});
                }
            }
            else {
                if (deleteOfficeDist == true) {
                    rimraf(decompressLocation, function () {});
                }
            }
        });
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
    else if(typeof(result) == "object") {
        for (var property in result) {
            if (result.hasOwnProperty(property)) {
                if (typeof(result[property]) == "string"){
                    if ((property == "a:t" || property == "_") && result[property] != "") {
                        await myTextExcel.push(result[property]);
                    }
                }
                else if (typeof(result[property][0]) == "string"){
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
                if (typeof(result[property]) == "string"){
                    if ((property == "t" || property == "_") && result[property] != "") {
                        await myTextExcel.push(result[property]);
                    }
                }
                else if (typeof(result[property][0]) == "string"){
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

var parseExcel = function(filename, callback, deleteOfficeDist = true) {
    if (validateFileExtension(filename, ["xlsx"])) {
        decompress(filename, decompressLocation).then(async files => {
            myTextExcel = [];


            if (fs.existsSync(decompressLocation + '/xl/worksheets')) {
                var workSheetsNum = await fs.readdirSync(decompressLocation + '/xl/worksheets').length;
                for (var i=0; i<workSheetsNum-1; i++) {
                    var parser = new xml2js.Parser();

                    var myData = await util.promisify(fs.readFile)(`${decompressLocation}/xl/worksheets/sheet${i+1}.xml`, 'utf8');
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

                    for (var i=0; i<drawingsNum; i++) {
                        var parser = new xml2js.Parser();

                        var myData = await util.promisify(fs.readFile)(`${decompressLocation}/xl/drawings/drawing${i+1}.xml`, 'utf8');
                        var result = await util.promisify(parser.parseString.bind(parser))(myData);
                        await scanForTextExcelDrawing(result);
                    }
                }

                callback(myTextExcel.join(" "));
                
                if (deleteOfficeDist == true) {
                    rimraf(decompressLocation, function () {});
                }
            }
            else {
                if (deleteOfficeDist == true) {
                    rimraf(decompressLocation, function () {});
                }
            }
        });
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
    else if(typeof(result) == "object") {
        for (var property in result) {
            if (result.hasOwnProperty(property)) {
                if (typeof(result[property]) == "string"){
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

var parseOpenOffice = async function(filename, callback, deleteOfficeDist = true) {
    if (validateFileExtension(filename, ["odt", "odp", "ods"])) {
        decompress(filename, decompressLocation).then(files => {
            myTextOpenOffice = [];
            if (fs.existsSync(decompressLocation + "/content.xml")) {
                fs.readFile(decompressLocation + '/content.xml', 'utf8', function (err,data) {
                    if (err) {
                       return console.log(err);
                     }
                     
                     xml2js.parseString(data, {"ignoreAttrs": true}, async function (err, result) {
                        await scanForTextOpenOffice(result);
        
                        callback(myTextOpenOffice.join(" "));
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
        });
    }
}

// #endregion textFetchFromOpenOffice




// #region validate file names

function validateFileExtension(filename, extension) {
    for (var extensionIterator = 0; extensionIterator < extension.length; extensionIterator++) {
        if (extension[extensionIterator] == filename.split(".").pop().toLowerCase()) {
            return true;
        }
    }
    console.log(`Sorry, we currently support docx, pptx, xlsx, odt, odp, ods files only. Make sure you pass appropriate file with the parsing functions.`);
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
        console.log(`Sorry, we currently support docx, pptx, xlsx, odt, odp, ods files only. Create a ticket in Issues on github to add support for ${extension} files. Stay tuned for further updates.`);
    }
    
}

// #endregion parse office


// #region setDecompressionLocation
function setDecompressionLocation(newLocation) {
    decompressLocation = newLocation + "/officeDist";
}

// #endregion setDecompressionLocation



module.exports.parseWord = parseWord;
module.exports.parsePowerPoint = parsePowerPoint;
module.exports.parseExcel = parseExcel;
module.exports.parseOpenOffice = parseOpenOffice;
module.exports.parseOffice = parseOffice;
module.exports.setDecompressionLocation = setDecompressionLocation;

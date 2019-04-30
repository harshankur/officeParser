const decompress = require('decompress');
const xml2js = require('xml2js')
const fs = require('fs')
const util = require('util');
const rimraf = require('rimraf');

// #region textFetchFromWord

var myTextWord = [];

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
        decompress(filename, 'officeDist').then(files => {
            myTextWord = [];
            if (fs.existsSync("officeDist/word/document.xml")) {
                fs.readFile('officeDist/word/document.xml', 'utf8', function (err,data) {
                    if (err) {
                       return console.log(err);
                     }
                     
                     xml2js.parseString(data, async function (err, result) {
                        await scanForTextWord(result);
        
                        callback(myTextWord.join(" "));
                        if (deleteOfficeDist == true) {
                            rimraf("officeDist", function () {});
                        }
                     });
                 });
            }
            else {
                if (deleteOfficeDist == true) {
                    rimraf("officeDist", function () {});
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
        decompress(filename, 'officeDist').then(async files => {
            myTextPowerPoint = [];

            if (fs.existsSync('officeDist/ppt/slides')) {
                var slidesNum = await fs.readdirSync('officeDist/ppt/slides').length;
                

                for (var i=0; i<slidesNum-1; i++) {
                    var parser = new xml2js.Parser();

                    var myData = await util.promisify(fs.readFile)(`officeDist/ppt/slides/slide${i+1}.xml`, 'utf8');
                    var result = await util.promisify(parser.parseString.bind(parser))(myData);
                    await scanForTextPowerPoint(result);
                }

                if (fs.existsSync('officeDist/ppt/notesSlides')) {
                    var notesSlidesNum = await fs.readdirSync('officeDist/ppt/notesSlides').length;
                    for (var i=0; i<notesSlidesNum-1; i++) {
                        var parser = new xml2js.Parser();
    
                        var myData = await util.promisify(fs.readFile)(`officeDist/ppt/notesSlides/notesSlide${i+1}.xml`, 'utf8');
                        var result = await util.promisify(parser.parseString.bind(parser))(myData);
                        await scanForTextPowerPoint(result);
                    }
                }
                
                callback(myTextPowerPoint.join(" "));
                    
                if (deleteOfficeDist == true) {
                    rimraf("officeDist", function () {});
                }
            }
            else {
                if (deleteOfficeDist == true) {
                    rimraf("officeDist", function () {});
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
        decompress(filename, 'officeDist').then(async files => {
            myTextExcel = [];


            if (fs.existsSync('officeDist/xl/worksheets')) {
                var workSheetsNum = await fs.readdirSync('officeDist/xl/worksheets').length;
                for (var i=0; i<workSheetsNum-1; i++) {
                    var parser = new xml2js.Parser();

                    var myData = await util.promisify(fs.readFile)(`officeDist/xl/worksheets/sheet${i+1}.xml`, 'utf8');
                    var result = await util.promisify(parser.parseString.bind(parser))(myData);
                    await scanForTextExcelWorkSheet(result);
                }

                if (fs.existsSync('officeDist/xl/sharedStrings.xml')) {
                    var parser = new xml2js.Parser();

                    var myData = await util.promisify(fs.readFile)(`officeDist/xl/sharedStrings.xml`, 'utf8');
                    var result = await util.promisify(parser.parseString.bind(parser))(myData);
                    await scanForTextExcelSharedStrings(result);
                }
            
                if (fs.existsSync('officeDist/xl/drawings')) {
                    var drawingsNum = await fs.readdirSync('officeDist/xl/drawings').length;

                    for (var i=0; i<drawingsNum; i++) {
                        var parser = new xml2js.Parser();

                        var myData = await util.promisify(fs.readFile)(`officeDist/xl/drawings/drawing${i+1}.xml`, 'utf8');
                        var result = await util.promisify(parser.parseString.bind(parser))(myData);
                        await scanForTextExcelDrawing(result);
                    }
                }

                callback(myTextExcel.join(" "));
                
                if (deleteOfficeDist == true) {
                    rimraf("officeDist", function () {});
                }
            }
            else {
                if (deleteOfficeDist == true) {
                    rimraf("officeDist", function () {});
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
        decompress(filename, 'officeDist').then(files => {
            myTextOpenOffice = [];
            if (fs.existsSync("officeDist/content.xml")) {
                fs.readFile('officeDist/content.xml', 'utf8', function (err,data) {
                    if (err) {
                       return console.log(err);
                     }
                     
                     xml2js.parseString(data, {"ignoreAttrs": true}, async function (err, result) {
                        await scanForTextOpenOffice(result);
        
                        callback(myTextOpenOffice.join(" "));
                        if (deleteOfficeDist == true) {
                            rimraf("officeDist", function () {});
                        }
                     });
                 });
            }
            else {
                if (deleteOfficeDist == true) {
                    rimraf("officeDist", function () {});
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



module.exports.parseWord = parseWord;
module.exports.parsePowerPoint = parsePowerPoint;
module.exports.parseExcel = parseExcel;
module.exports.parseOpenOffice = parseOpenOffice;
module.exports.parseOffice = parseOffice;

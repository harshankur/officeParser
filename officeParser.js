const decompress = require('decompress');
const xml2js = require('xml2js')
const fs = require('fs')
const util = require('util');

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
                    if ((property == "w:t" || property == "_" ) && result[i] != "") {
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

var parseWord = async function(filename, callback) {
    if (validateFileExtension(filename, "docx")) {
        decompress(filename, 'dist').then(files => {
            myTextWord = [];
            if (fs.existsSync("dist/word/document.xml")) {
                fs.readFile('dist/word/document.xml', 'utf8', function (err,data) {
                    if (err) {
                       return console.log(err);
                     }
                     
                     xml2js.parseString(data, async function (err, result) {
                         await scanForTextWord(result);
         
                         callback(myTextWord.join(" "));
                     });
                 });
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


var parsePowerPoint = function(filename, callback) {
    if (validateFileExtension(filename, "pptx")) {
        decompress(filename, 'dist').then(async files => {
            myTextPowerPoint = [];

            if (fs.existsSync('dist/ppt/slides')) {
                var slidesNum = await fs.readdirSync('dist/ppt/slides').length;
                

                for (var i=0; i<slidesNum-1; i++) {
                    var parser = new xml2js.Parser();

                    var myData = await util.promisify(fs.readFile)(`dist/ppt/slides/slide${i+1}.xml`, 'utf8');
                    var result = await util.promisify(parser.parseString.bind(parser))(myData);
                    await scanForTextPowerPoint(result);
                }

                if (fs.existsSync('dist/ppt/notesSlides')) {
                    var notesSlidesNum = await fs.readdirSync('dist/ppt/notesSlides').length;
                    for (var i=0; i<notesSlidesNum-1; i++) {
                        var parser = new xml2js.Parser();
    
                        var myData = await util.promisify(fs.readFile)(`dist/ppt/notesSlides/notesSlide${i+1}.xml`, 'utf8');
                        var result = await util.promisify(parser.parseString.bind(parser))(myData);
                        await scanForTextPowerPoint(result);
                    }
                }
                
                callback(myTextPowerPoint.join(" "));
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

var parseExcel = function(filename, callback) {
    if (validateFileExtension(filename, "xlsx")) {
        decompress(filename, 'dist').then(async files => {
            myTextExcel = [];


            if (fs.existsSync('dist/xl/worksheets')) {
                var workSheetsNum = await fs.readdirSync('dist/xl/worksheets').length;
                for (var i=0; i<workSheetsNum-1; i++) {
                    var parser = new xml2js.Parser();

                    var myData = await util.promisify(fs.readFile)(`dist/xl/worksheets/sheet${i+1}.xml`, 'utf8');
                    var result = await util.promisify(parser.parseString.bind(parser))(myData);
                    await scanForTextExcelWorkSheet(result);
                }

                if (fs.existsSync('dist/xl/sharedStrings.xml')) {
                    var parser = new xml2js.Parser();

                    var myData = await util.promisify(fs.readFile)(`dist/xl/sharedStrings.xml`, 'utf8');
                    var result = await util.promisify(parser.parseString.bind(parser))(myData);
                    await scanForTextExcelSharedStrings(result);
                }
            
                if (fs.existsSync('dist/xl/drawings')) {
                    var drawingsNum = await fs.readdirSync('dist/xl/drawings').length;

                    for (var i=0; i<drawingsNum; i++) {
                        var parser = new xml2js.Parser();

                        var myData = await util.promisify(fs.readFile)(`dist/xl/drawings/drawing${i+1}.xml`, 'utf8');
                        var result = await util.promisify(parser.parseString.bind(parser))(myData);
                        await scanForTextExcelDrawing(result);
                    }
                }

                callback(myTextExcel.join(" "));
            }
        });
    }
} 

// #endregion textFetchFromExcel


// #region validate file names

function validateFileExtension(filename, extension) {
    if (extension == filename.split(".").pop()) {
        return true;
    }
    else {
        console.log(`Sorry, we support docx, pptx and xlsx files only. Stay tuned for further support.`);
    }
    return false;
}

// #endregion validate file names




module.exports.parseWord = parseWord;
module.exports.parsePowerPoint = parsePowerPoint;
module.exports.parseExcel = parseExcel;

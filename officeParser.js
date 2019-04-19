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

var parseWord = function(filename, callback) {
    decompress(filename, 'dist').then(files => {
        myTextWord = [];
        fs.readFile('dist/word/document.xml', 'utf8', function (err,data) {
           if (err) {
              return console.log(err);
            }
            
            xml2js.parseString(data, async function (err, result) {
                await scanForTextWord(result);

                callback(myTextWord.join(" "));
            });
        });
    });
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
    decompress(filename, 'dist').then(async files => {
        myTextPowerPoint = [];

        var slidesNum = await fs.readdirSync('dist/ppt/slides').length;
        var notesSlidesNum = await fs.readdirSync('dist/ppt/notesSlides').length;
        console.log(slidesNum);

        for (var i=0; i<slidesNum-1; i++) {
            var parser = new xml2js.Parser();

            var myData = await util.promisify(fs.readFile)(`dist/ppt/slides/slide${i+1}.xml`, 'utf8');
            var result = await util.promisify(parser.parseString.bind(parser))(myData);
            await scanForTextPowerPoint(result);
        }

        for (var i=0; i<notesSlidesNum-1; i++) {
            var parser = new xml2js.Parser();

            var myData = await util.promisify(fs.readFile)(`dist/ppt/notesSlides/notesSlide${i+1}.xml`, 'utf8');
            var result = await util.promisify(parser.parseString.bind(parser))(myData);
            await scanForTextPowerPoint(result);
        }

        callback(myTextPowerPoint.join(" "));
    });
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
    decompress(filename, 'dist').then(async files => {
        myTextExcel = [];

        var workSheetsNum = await fs.readdirSync('dist/xl/worksheets').length;
        var drawingsNum = await fs.readdirSync('dist/xl/drawings').length;
        console.log(workSheetsNum);
        console.log(drawingsNum);

        for (var i=0; i<workSheetsNum-1; i++) {
            var parser = new xml2js.Parser();

            var myData = await util.promisify(fs.readFile)(`dist/xl/worksheets/sheet${i+1}.xml`, 'utf8');
            var result = await util.promisify(parser.parseString.bind(parser))(myData);
            await scanForTextExcelWorkSheet(result);
        }
        console.log(result["worksheet"]["sheetData"][0]["row"][4]["c"][0]["v"][0]);


        var parser = new xml2js.Parser();

        var myData = await util.promisify(fs.readFile)(`dist/xl/sharedStrings.xml`, 'utf8');
        var result = await util.promisify(parser.parseString.bind(parser))(myData);
        await scanForTextExcelSharedStrings(result);


        for (var i=0; i<drawingsNum; i++) {
            var parser = new xml2js.Parser();

            var myData = await util.promisify(fs.readFile)(`dist/xl/drawings/drawing${i+1}.xml`, 'utf8');
            var result = await util.promisify(parser.parseString.bind(parser))(myData);
            await scanForTextExcelDrawing(result);
        }

        callback(myTextExcel.join(" "));
    });
} 

// #endregion textFetchFromExcel

module.exports.parseWord = parseWord;
module.exports.parsePowerPoint = parsePowerPoint;
module.exports.parseExcel = parseExcel;

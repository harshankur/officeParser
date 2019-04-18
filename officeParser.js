const decompress = require('decompress');
const xml2js = require('xml2js')
const fs = require('fs')

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


module.exports.parseWord = parseWord;
module.exports.parsePowerPoint = parsePowerPoint;

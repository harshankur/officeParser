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
                // fs.writeFile("files/test.txt", myTextWord.join(" "), function(err) {
                //     if(err) {
                //         return console.log(err);
                //     }
                
                //     console.log("The file was saved!");
                // }); 
                callback(myTextWord.join(" "));
            });
        });
    });
}

// #endregion textFetchFromWord


module.exports.parseWord = parseWord;

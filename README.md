# officeParser
A Node.js library to parse text out of any office file. 

~~Currently supports docx. Support for pptx and xlsx coming soon~~
~~Currently supports docx and pptx. xlsx support coming soon.~~

##### Update
*Supports docx, pptx and xlsx files. Investigating possibilities of fetching text from doc, ppt, pps, xls files.*



## Install via npm


```
npm i officeparser
```

----------

**Usage**
```
const officeParser = require('officeparser');

officeParser.parseWord("/path/to/word.docx", function(data){
        // "data" string in the callback here is the text parsed from the word file passed in the first argument above
        console.log(data)
})

officeParser.parsePowerPoint("/path/to/powerpoint.pptx", function(data){
        // "data" string in the callback here is the text parsed from the powerpoint file passed in the first argument above
        console.log(data)
})

officeParser.parseExcel("/path/to/excel.xlsx", function(data){
        // "data" string in the callback here is the text parsed from the excel file passed in the first argument above
        console.log(data)
})
```

**Example**
```
const officeParser = require('officeparser');

officeParser.parseWord("C:\\files\\myText.docx", function(data){
        var newText = data + "look, I can parse a word file"
        callSomeOtherFunction(newText);
})

officeParser.parsePowerPoint("/Users/harsh/Desktop/files/mySlides.pptx", function(data){
        var newText = data + "look, I can parse a powerpoint file"
        callSomeOtherFunction(newText);
})

// Using relative path for file is also fine
officeParser.parseExcel("files/myWorkSheet.xlsx", function(data){
        var newText = data + "look, I can parse an excel file"
        callSomeOtherFunction(newText);
})
```

----------

**github**
https://github.com/harshankur/officeParser

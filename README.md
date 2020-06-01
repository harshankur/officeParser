# officeParser
A Node.js library to parse text out of any office file. 

### Supported File Types

- [`docx`](https://en.wikipedia.org/wiki/Office_Open_XML)
- [`pptx`](https://en.wikipedia.org/wiki/Office_Open_XML)
- [`xlsx`](https://en.wikipedia.org/wiki/Office_Open_XML)
- [`odt`](https://en.wikipedia.org/wiki/OpenDocument)
- [`odp`](https://en.wikipedia.org/wiki/OpenDocument)
- [`ods`](https://en.wikipedia.org/wiki/OpenDocument)


#### Update
* 2020/06/01 - Added error handling and console.log enable/disable methods. Default is set at enabled. Everything backward compatible.
* 2019/06/17 - Added method to change location for decompressing office files in places with restricted write access.
* 2019/04/30 - Removed case sensitive file extension bug. File names with capital lettered extensions now supported.
* 2019/04/23 - Added support for open office files *.odt, *.odp, *.ods through parseOffice function. Created a new method parseOpenOffice for those who prefer targetted functions. 
* 2019/04/23 - Added feature to delete the generated dist folder after function callback
* 2019/04/22 - Added parseOffice method to avoid confusion between type of file and their extension
* 2019/04/22 - Added file extension validations. Removed errors for excel files with no drawing elements.
* 2019/04/19 - Support added for *.xlsx files.
* 2019/04/18 - Support added for *.pptx files.



## Install via npm


```
npm i officeparser
```

----------

**Usage**
```
const officeParser = require('officeparser');

officeParser.parseOffice("/path/to/officeFile", function(data, err){
        // "data" string in the callback here is the text parsed from the office file passed in the first argument above
        if (err) return console.log(err);
        console.log(data)
})

```

**Please take note: I have breached convention in placing err as second argument in my callback but please understand that I had to do it to not break other people's existing modules.**

*Optionally change decompression location for office Files at personalised locations for environments with restricted write access*

```
const officeParser = require('officeparser');

// Default decompress location for office Files is "officeDist" in the directory where Node is started. 
// Put this file before parseOffice method to take effect.
officeParser.setDecompressionLocation("/tmp");  // New decompression location would be "/tmp/officeDist"

// P.S.: Setting location on a Windows environment with '\' heirarchy requires to be entered twice '\\'
officeParser.setDecompressionLocation("C:\\tmp");  // New decompression location would be "C:\tmp\officeDist"


officeParser.parseOffice("/path/to/officeFile", function(data, err){
        // "data" string in the callback here is the text parsed from the office file passed in the first argument above
        if (err) return console.log(err);
        console.log(data)
})
```

*Optionally add false as 3rd variable to parseOffice to not delete the generated officeDist folder*

```
officeParser.parseOffice("/path/to/officeFile", function(data, err){
        // "data" string in the callback here is the text parsed from the office file passed in the first argument above
        if (err) return console.log(err);
        console.log(data)
}, false)
```

**Example**
```
const officeParser = require('officeparser');

officeParser.parseOffice("C:\\files\\myText.docx", function(data, err){
        if (err) return console.log(err);
        var newText = data + "look, I can parse a word file"
        callSomeOtherFunction(newText);
})

officeParser.parseOffice("/Users/harsh/Desktop/files/mySlides.pptx", function(data, err){
        if (err) return console.log(err);
        var newText = data + "look, I can parse a powerpoint file"
        callSomeOtherFunction(newText);
})

// Using relative path for file is also fine
officeParser.parseOffice("files/myWorkSheet.ods", function(data, err){
        if (err) return console.log(err);
        var newText = data + "look, I can parse an excel file"
        callSomeOtherFunction(newText);
})
```


----------

### Old but functional way of extracting text from word, powerpoint and excel files
*These were the initial methods of parsing text till parseOffice method came into existence. These still exist and form the skeleton to this module as parseOffice redirects the below functions anyway. These functions will forever remain available to guarantee long-term usage of this module. I will ensure backward-compatibility with all previous versions.*

**Usage**
```
const officeParser = require('officeparser');

officeParser.parseWord("/path/to/word.docx", function(data, err){
        // "data" string in the callback here is the text parsed from the word file passed in the first argument above
        if (err) return console.log(err);
        console.log(data)
})

officeParser.parsePowerPoint("/path/to/powerpoint.pptx", function(data, err){
        // "data" string in the callback here is the text parsed from the powerpoint file passed in the first argument above
        if (err) return console.log(err);
        console.log(data)
})

officeParser.parseExcel("/path/to/excel.xlsx", function(data, err){
        // "data" string in the callback here is the text parsed from the excel file passed in the first argument above
        if (err) return console.log(err);
        console.log(data)
})

officeParser.parseOpenOffice("/path/to/writer.odt", function(data, err){
        // "data" string in the callback here is the text parsed from the writer file passed in the first argument above
        if (err) return console.log(err);
        console.log(data)
})
```

**Example**
```
const officeParser = require('officeparser');

officeParser.parseWord("C:\\files\\myText.docx", function(data, err){
        if (err) return console.log(err);
        var newText = data + "look, I can parse a word file"
        callSomeOtherFunction(newText);
})

officeParser.parsePowerPoint("/Users/harsh/Desktop/files/mySlides.pptx", function(data, err){
        if (err) return console.log(err);
        var newText = data + "look, I can parse a powerpoint file"
        callSomeOtherFunction(newText);
})

// Using relative path for file is also fine
officeParser.parseExcel("files/myWorkSheet.xlsx", function(data, err){
        if (err) return console.log(err);
        var newText = data + "look, I can parse an excel file"
        callSomeOtherFunction(newText);
})

officeParser.parseOpenOffice("files/myDocument.odt", function(data, err){
        if (err) return console.log(err);
        var newText = data + "look, I can parse an OpenOffice file"
        callSomeOtherFunction(newText);
})
```

----------

**github**
https://github.com/harshankur/officeParser

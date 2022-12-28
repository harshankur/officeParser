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
* 2022/12/28 - Added command line method to use officeParser with or without installing it and instantly get parsed content on the console.
* 2022/12/10 - Fixed memory leak issues, bugs related to parsing open document files and improved error handling.
* 2021/11/21 - Added promise way to existing callback functions.
* 2020/06/01 - Added error handling and console.log enable/disable methods. Default is set at enabled. Everything backward compatible.
* 2019/06/17 - Added method to change location for decompressing office files in places with restricted write access.
* 2019/04/30 - Removed case sensitive file extension bug. File names with capital lettered extensions now supported.
* 2019/04/23 - Added support for open office files *.odt, *.odp, *.ods through parseOffice function. Created a new method parseOpenOffice for those who prefer targetted functions. 
* 2019/04/23 - Added feature to delete the generated dist folder after function callback.
* 2019/04/22 - Added parseOffice method to avoid confusion between type of file and their extension.
* 2019/04/22 - Added file extension validations. Removed errors for excel files with no drawing elements.
* 2019/04/19 - Support added for *.xlsx files.
* 2019/04/18 - Support added for *.pptx files.



## Install via npm


```
npm i officeparser
```

## Command Line usage
If you want to call the installed officeParser.js file, use below command
```
node </path/to/officeParser.js> <fileName>
```

Otherwise, you can simply use npx to instantly extract parsed data.
```
npx officeparser <fileName>
```

----------

**Library Usage**
```js
const officeParser = require('officeparser');

// callback
officeParser.parseOffice("/path/to/officeFile", function(data, err){
    // "data" string in the callback here is the text parsed from the office file passed in the first argument above
    if (err) return console.log(err);
    console.log(data)
})

// promise
officeParser.parseOfficeAsync("/path/to/officeFile");
// "data" string in the promise here is the text parsed from the office file passed in the argument above
.then((data) => {
    console.log(data)
})
.catch(err) => {
    console.log(err)
}

// async/await
try {
    // "data" string returned from promise here is the text parsed from the office file passed in the argument
    const data = await officeParser.parseOfficeAsync("/path/to/officeFile");
    console.log(data);
} catch (err) {
    // resolve error
    console.log(err);
}
```

**Please take note: I have breached convention in placing err as second argument in my callback but please understand that I had to do it to not break other people's existing modules.**

*Optionally change decompression location for office Files at personalised locations for environments with restricted write access*

```js
const officeParser = require('officeparser');

// Default decompress location for office Files is "officeDist" in the directory where Node is started. 
// Put this file before parseOffice method to take effect.
officeParser.setDecompressionLocation("/tmp");  // New decompression location would be "/tmp/officeDist"

// P.S.: Setting location on a Windows environment with '\' hierarchy requires to be entered twice '\\'
officeParser.setDecompressionLocation("C:\\tmp");  // New decompression location would be "C:\tmp\officeDist"


officeParser.parseOffice("/path/to/officeFile", function(data, err){
    // "data" string in the callback here is the text parsed from the office file passed in the first argument above
    if (err) return console.log(err);
    console.log(data)
})
```

*Optionally add false as 3rd variable to parseOffice to not delete the generated officeDist folder*

```js
// callback
officeParser.parseOffice("/path/to/officeFile", function(data, err){
    if (err) return console.log(err);
    console.log(data)
}, false)

// promise
officeParser.parseOfficeAsync("/path/to/officeFile", false);
.then((data) => {
    console.log(data)
})
.catch(err) => {
    console.log(err)
}

// async/await
try {
    const data = await officeParser.parseOfficeAsync("/path/to/officeFile", false);
    console.log(data);
} catch (err) {
    // resolve error
    console.log(err);
}
```

**Example**
```js
const officeParser = require('officeparser');

// callback
officeParser.parseOffice("C:\\files\\myText.docx", function(data, err){
    if (err) return console.log(err);
    var newText = data + "look, I can parse a word file"
    callSomeOtherFunction(newText);
})


// promise
officeParser.parseOfficeAsync("/Users/harsh/Desktop/files/mySlides.pptx");
.then((data) => {
    var newText = data + "look, I can parse a powerpoint file"
    callSomeOtherFunction(newText);
})
.catch(err) => {
    console.log(err)
}

// Using relative path for file is also fine
officeParser.parseOffice("files/myWorkSheet.ods", function(data, err){
    if (err) return console.log(err);
    var newText = data + "look, I can parse an excel file"
    callSomeOtherFunction(newText);
})

// async/await
try {
    const data = await officeParser.parseOfficeAsync("/Users/harsh/Desktop/files/mySlides.pptx");
    let newText = data + "look, I can parse a powerpoint file";
    await callSomeOtherFunction(newText);
} catch (err) {
    // resolve error
    console.log(err);
}
```


----------

### Old but functional way of extracting text from word, powerpoint and excel files
*These were the initial methods of parsing text till parseOffice method came into existence. These still exist and form the skeleton to this module as parseOffice redirects the below functions anyway. These functions will forever remain available to guarantee long-term usage of this module. I will ensure backward-compatibility with all previous versions.*

**Usage**
```js
const officeParser = require('officeparser');

// callback
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

// promise
officeParser.parseWordAsync("/path/to/word.docx");
.then((data) => {
    // data is the parsed text
})
officeParser.parsePowerPointAsync("/path/to/powerpoint.pptx");
.then((data) => {
    // data is the parsed text
})
officeParser.parseExcelAsync("/path/to/excel.xlsx");
.then((data) => {
    // data is the parsed text
})
officeParser.parseOpenOfficeAsync("/path/to/writer.odt");
.then((data) => {
    // data is the parsed text
})

// async/await
try {
    // "data" string returned from promise here is the text parsed from the office file passed in the first argument
    const data1 = await officeParser.parseWordAsync("/path/to/word.docx");

    const data2 = await officeParser.parsePowerPointAsync("/path/to/powerpoint.pptx");

    const data3 = await officeParser.parseExcelAsync("/path/to/excel.xlsx");

    const data3 = await officeParser.parseOpenOfficeAsync("/path/to/writer.odt");
} catch (err) {
    // resolve error
    console.log(err);
}
```

**Example**
```js
const officeParser = require('officeparser');

// callback
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

// promise
officeParser.parseWordAsync("C:\\files\\myText.docx");
.then((data) => {
    let newText1 = data1 + "look, I can parse a word file";
    callSomeOtherFunction(newText1);
})
.catch(err) => {
    console.log(err)
}

officeParser.parsePowerPointAsync("/Users/harsh/Desktop/files/mySlides.pptx");
.then((data) => {
    let newText2 = data2 + "look, I can parse a powerpoint file";
    callSomeOtherFunction(newText2);
})
.catch(err) => {
    console.log(err)
}

officeParser.parseExcelAsync("files/myWorkSheet.xlsx");
.then((data) => {
    let newText3 = data3 + "look, I can parse an excel file";
    callSomeOtherFunction(newText3);
})
.catch(err) => {
    console.log(err)
}

officeParser.parseOpenOfficeAsync("files/myDocument.odt");
.then((data) => {
    let newText4 = data4 + "look, I can parse an OpenOffice file";
    callSomeOtherFunction(newText4);
})
.catch(err) => {
    console.log(err)
}


// async/await
try {
    const data1 = await officeParser.parseWordAsync("C:\\files\\myText.docx");
    let newText1 = data1 + "look, I can parse a word file";
    await callSomeOtherFunction(newText1);

    const data2 = await officeParser.parsePowerPointAsync("/Users/harsh/Desktop/files/mySlides.pptx");
    let newText2 = data2 + "look, I can parse a powerpoint file";
    await callSomeOtherFunction(newText2);

    // Using relative path for file is also fine
    const data3 = await officeParser.parseExcelAsync("files/myWorkSheet.xlsx");
    let newText3 = data3 + "look, I can parse an excel file";
    await callSomeOtherFunction(newText3);

    const data4 = await officeParser.parseOpenOfficeAsync("files/myDocument.odt");
    let newText4 = data4 + "look, I can parse an OpenOffice file";
    await callSomeOtherFunction(newText4);
} catch (err) {
    // resolve error
    console.log(err);
}
```

----------

**github**
https://github.com/harshankur/officeParser

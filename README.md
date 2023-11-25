# officeParser
A Node.js library to parse text out of any office file. 

### Supported File Types

- [`docx`](https://en.wikipedia.org/wiki/Office_Open_XML)
- [`pptx`](https://en.wikipedia.org/wiki/Office_Open_XML)
- [`xlsx`](https://en.wikipedia.org/wiki/Office_Open_XML)
- [`odt`](https://en.wikipedia.org/wiki/OpenDocument)
- [`odp`](https://en.wikipedia.org/wiki/OpenDocument)
- [`ods`](https://en.wikipedia.org/wiki/OpenDocument)
- [`pdf`](https://en.wikipedia.org/wiki/PDF)


#### Update
* 2023/11/25 - Fixed error catching when an error occurs within the parsing of a file, especially after decompressing it. Also fixed the problem with parallel parsing of files as we were using only timestamp in file names.
* 2023/10/24 - Revamped content parsing code. Fixed order of content in files, especially in word files where table information would always land up at the end of the text. Added config object as argument for parseOffice which can be used to set new line delimiter and multiple other configurations. Added support for parsing pdf files using the popular npm library pdf-parse. Removed support for individual file parsing functions.
* 2023/04/26 - Added support for file buffers as argument for filepath for parseOffice and parseOfficeAsync
* 2023/04/07 - Added typings to methods to help with Typescript projects.
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


## Library Usage
```js
const officeParser = require('officeparser');

// callback
officeParser.parseOffice("/path/to/officeFile", function(data, err) {
    // "data" string in the callback here is the text parsed from the office file passed in the first argument above
    if (err) {
        console.log(err);
        return;
    }
    console.log(data);
})

// promise
officeParser.parseOfficeAsync("/path/to/officeFile");
// "data" string in the promise here is the text parsed from the office file passed in the argument above
    .then(data => console.log(data))
    .catch(err => console.error(err))

// async/await
try {
    // "data" string returned from promise here is the text parsed from the office file passed in the argument
    const data = await officeParser.parseOfficeAsync("/path/to/officeFile");
    console.log(data);
} catch (err) {
    // resolve error
    console.log(err);
}

// USING FILE BUFFERS
// instead of file path, you can also pass file buffers of one of the supported files
// on parseOffice or parseOfficeAsync functions.

// get file buffers
const fileBuffers = fs.readFileSync("/path/to/officeFile");
// get parsed text from officeParser
// NOTE: Only works with parseOffice. Old functions are not supported.
officeParser.parseOfficeAsync(fileBuffers);
    .then(data => console.log(data))
    .catch(err => console.error(err))
```

### Configuration Object: OfficeParserConfig
*Optionally add a config object as 3rd variable to parseOffice for the following configurations*
| Flag                 | DataType | Default          | Explanation                                                                                                                                                                                                                                     |
|----------------------|----------|------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| tempFilesLocation    | string   | officeParserTemp | The directory where officeparser stores the temp files . The final decompressed data will be put inside officeParserTemp folder within your directory. **Please ensure that this directory actually exists.** Default is officeParserTemp.      |
| preserveTempFiles    | boolean  | false            | Flag to not delete the internal content files and the possible duplicate temp files that it uses after unzipping office files. Default is false. It always deletes all of those files.                                                          |
| outputErrorToConsole | boolean  | false            | Flag to show all the logs to console in case of an error. Default is false.                                                                                                                                                                     |
| newlineDelimiter     | string   | \n               | The delimiter used for every new line in places that allow multiline text like word. Default is \n.                                                                                                                                             |
| ignoreNotes          | boolean  | false            | Flag to ignore notes from parsing in files like powerpoint. Default is false. It includes notes in the parsed text by default.                                                                                                                  |
| putNotesAtLast       | boolean  | false            | Flag, if set to true, will collectively put all the parsed text from notes at last in files like powerpoint. Default is false. It puts each notes right after its main slide content. If ignoreNotes is set to true, this flag is also ignored. |
<br>

```js
const config = {
    newlineDelimiter: " ",  // Separate new lines with a space instead of the default \n.
    ignoreNotes: true       // Ignore notes while parsing presentation files like pptx or odp.
}

// callback
officeParser.parseOffice("/path/to/officeFile", function(data, err){
    if (err) {
        console.log(err);
        return;
    }
    console.log(data);
}, config)

// promise
officeParser.parseOfficeAsync("/path/to/officeFile", config);
    .then(data => console.log(data))
    .catch(err => console.error(err))
```

**Example - JavaScript**
```js
const officeParser = require('officeparser');

const config = {
    newlineDelimiter: " ",  // Separate new lines with a space instead of the default \n.
    ignoreNotes: true       // Ignore notes while parsing presentation files like pptx or odp.
}

// relative path is also fine => eg: files/myWorkSheet.ods
officeParser.parseOfficeAsync("/Users/harsh/Desktop/files/mySlides.pptx", config);
    .then(data => {
        const newText = data + " look, I can parse a powerpoint file";
        callSomeOtherFunction(newText);
    })
    .catch(err => console.error(err));

// Search for a term in the parsed text.
function searchForTermInOfficeFile(searchterm, filepath) {
    return officeParser.parseOfficeAsync(filepath)
        .then(data => data.indexOf(searchterm) != -1)
}
```


**Example - TypeScript**
```ts
const officeParser = require('officeparser');

const config: OfficeParserConfig = {
    newlineDelimiter: " ",  // Separate new lines with a space instead of the default \n.
    ignoreNotes: true       // Ignore notes while parsing presentation files like pptx or odp.
}

// relative path is also fine => eg: files/myWorkSheet.ods
officeParser.parseOfficeAsync("/Users/harsh/Desktop/files/mySlides.pptx", config);
    .then(data => {
        const newText = data + " look, I can parse a powerpoint file";
        callSomeOtherFunction(newText);
    })
    .catch(err => console.error(err));

// Search for a term in the parsed text.
function searchForTermInOfficeFile(searchterm: string, filepath: string): Promise<boolean> {
    return officeParser.parseOfficeAsync(filepath)
        .then(data => data.indexOf(searchterm) != -1)
}
```
\
**Please take note: I have breached convention in placing err as second argument in my callback but please understand that I had to do it to not break other people's existing modules.**


## Known Bugs
1. Inconsistency and incorrectness in the positioning of footnotes and endnotes in .docx files where the footnotes and endnotes would end up at the end of the parsed text whereas it would be positioned exactly after the referenced word in .odt files.
2. The charts and objects information of .odt files are not accurate and may end up showing a few NaN in some cases.
----------

**npm**
https://npmjs.com/package/officeparser

**github**
https://github.com/harshankur/officeParser

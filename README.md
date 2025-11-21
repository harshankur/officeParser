# officeParser
A Node.js library to parse text and images out of any office file.

### Supported File Types

- [`docx`](https://en.wikipedia.org/wiki/Office_Open_XML)
- [`pptx`](https://en.wikipedia.org/wiki/Office_Open_XML)
- [`xlsx`](https://en.wikipedia.org/wiki/Office_Open_XML)
- [`odt`](https://en.wikipedia.org/wiki/OpenDocument)
- [`odp`](https://en.wikipedia.org/wiki/OpenDocument)
- [`ods`](https://en.wikipedia.org/wiki/OpenDocument)
- [`pdf`](https://en.wikipedia.org/wiki/PDF)


#### Update
* 2025/01 - **Breaking Change**: Replaced `images` array with `blocks` array that preserves document order. Blocks contain both text and images in the order they appear in the document. Use `data.blocks.filter(b => b.type === 'image')` to get images.
* 2024/11/20 - Added image extraction support for docx and pdf files.
* 2024/11/12 - Added ArrayBuffer as a type of file input. Generating bundle files now which exposes namespace officeParser to be able to access parseOffice and parseOfficeAsync directly on the browser. Extracting text out of pdf files does not work currently in browser bundles.
* 2024/10/21 - Replaced extracting zip files from decompress to yauzl. This means that we now extract files in memory and we no longer need to write them to disk. Removed config flags related to extracted files. Added flags for CLI execution.
* 2024/10/15 - Fixed erroring out while deleting temp files when multiple worker threads make parallel executions resulting in same file name for multiple files. Fixed erroring out when multiple executions are made without waiting for the previous execution to finish which resulted in deleting the file from other execution. Upgraded dependencies.
* 2024/10/13 - Fixed parsing text from xlsx files which contain no shared strings file and files which have inlineStr based strings.
* 2024/05/06 - Replaced pdf parsing support from pdf-parse library to natively building it using pdf.js library from Mozilla by analyzing its output. Added pdfjs-dist build as a local library.
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
node <path/to/officeParser.js> [--configOption=value] [FILE_PATH]
node officeparser [--configOption=value] [FILE_PATH]
```

Otherwise, you can simply use npx without installing the node module to instantly extract parsed data.
```
npx officeparser [--configOption=value] [FILE_PATH]
```

### Config Options:
- `--ignoreNotes=[true|false]`          Flag to ignore notes from files like PowerPoint. Default is false.
- `--newlineDelimiter=[delimiter]`      The delimiter to use for new lines. Default is `\n`.
- `--putNotesAtLast=[true|false]`       Flag to collect notes at the end of files like PowerPoint. Default is false.
- `--outputErrorToConsole=[true|false]` Flag to output errors to the console. Default is false.
- `--extractImages=[true|false]`        Flag to extract images from files. Default is false. **Images are saved to current directory.**

## Library Usage
```js
const officeParser = require('officeparser');

// callback
officeParser.parseOffice("/path/to/officeFile", function(data, err) {
    // "data" is an object with "text" and "blocks" properties
    if (err) {
        console.log(err);
        return;
    }
    console.log(data.text);
    console.log(data.blocks); // Array of content blocks (text and images in document order)
})

// promise
officeParser.parseOfficeAsync("/path/to/officeFile");
// "data" is an object with "text" and "blocks" properties
    .then(data => {
        console.log(data.text);
        console.log(data.blocks);
    })
    .catch(err => console.error(err))

// async/await
try {
    // "data" is an object with "text" and "blocks" properties
    const data = await officeParser.parseOfficeAsync("/path/to/officeFile");
    console.log(data.text);
    console.log(data.blocks);
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
    .then(data => {
        console.log(data.text);
        console.log(data.blocks);
    })
    .catch(err => console.error(err))
```

### Block-Based Output Structure
`officeParser` returns content in an ordered `blocks` array that preserves the document structure. Each block is either a text block or an image block, appearing in the order they occur in the document.

**Block Types:**

**TextBlock:**
- `type`: `'text'`
- `content`: The text content string

**ImageBlock** (when `extractImages: true`):
- `type`: `'image'`
- `buffer`: A `Buffer` containing the image data
- `mimeType`: The MIME type of the image (e.g., `image/jpeg`, `image/png`)
- `filename`: The filename of the image (when available)

**Example:**
```js
const officeParser = require('officeparser');
const fs = require('fs');

const config = { extractImages: true };

officeParser.parseOfficeAsync("/path/to/document.docx", config)
    .then(data => {
        console.log(data.text); // Full text (backwards compatible)

        // Get images from blocks
        const imageBlocks = data.blocks.filter(b => b.type === 'image');
        console.log(`Found ${imageBlocks.length} images`);

        // Process blocks in document order
        data.blocks.forEach((block, index) => {
            if (block.type === 'text') {
                console.log(`Text: ${block.content.slice(0, 50)}...`);
            } else if (block.type === 'image') {
                const extension = block.mimeType.split('/')[1];
                const filename = block.filename || `image_${index}.${extension}`;
                fs.writeFileSync(filename, block.buffer);
            }
        });
    })
    .catch(err => console.error(err));
```

### Configuration Object: OfficeParserConfig
*Optionally add a config object as 3rd variable to parseOffice for the following configurations*
| Flag                 | DataType | Default          | Explanation                                                                                                                                                                                                                                     |
|----------------------|----------|------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| outputErrorToConsole | boolean  | false            | Flag to show all the logs to console in case of an error. Default is false.                                                                                                                                                                     |
| newlineDelimiter     | string   | \n               | The delimiter used for every new line in places that allow multiline text like word. Default is \n.                                                                                                                                             |
| ignoreNotes          | boolean  | false            | Flag to ignore notes from parsing in files like powerpoint. Default is false. It includes notes in the parsed text by default.                                                                                                                  |
| putNotesAtLast       | boolean  | false            | Flag, if set to true, will collectively put all the parsed text from notes at last in files like powerpoint. Default is false. It puts each notes right after its main slide content. If ignoreNotes is set to true, this flag is also ignored. |
| extractImages        | boolean  | false            | Flag to extract images from files. Default is false. If set to true, the `blocks` array will contain image blocks alongside text blocks.                                                                                                        |
<br>

```js
const config = {
    newlineDelimiter: " ",  // Separate new lines with a space instead of the default \n.
    ignoreNotes: true,      // Ignore notes while parsing presentation files like pptx or odp.
    extractImages: true     // Extract images from docx and pdf files.
}

// callback
officeParser.parseOffice("/path/to/officeFile", function(data, err){
    if (err) {
        console.log(err);
        return;
    }
    console.log(data.text);
}, config)

// promise
officeParser.parseOfficeAsync("/path/to/officeFile", config);
    .then(data => console.log(data.text))
    .catch(err => console.error(err))
```

**Example - JavaScript**
```js
const officeParser = require('officeparser');
const fs = require('fs');

const config = {
    newlineDelimiter: " ",  // Separate new lines with a space instead of the default \n.
    ignoreNotes: true,      // Ignore notes while parsing presentation files like pptx or odp.
    extractImages: true     // Extract images from files.
}

// relative path is also fine => eg: files/myWorkSheet.ods
officeParser.parseOfficeAsync("/Users/harsh/Desktop/files/mySlides.pptx", config);
    .then(data => {
        const newText = data.text + " look, I can parse a powerpoint file";
        callSomeOtherFunction(newText);
        // Save images from blocks
        data.blocks
            .filter(block => block.type === 'image')
            .forEach((image, index) => {
                fs.writeFileSync(`image_${index}.${image.mimeType.split('/')[1]}`, image.buffer);
            });
    })
    .catch(err => console.error(err));

// Search for a term in the parsed text.
function searchForTermInOfficeFile(searchterm, filepath) {
    return officeParser.parseOfficeAsync(filepath)
        .then(data => data.text.indexOf(searchterm) != -1)
}
```


**Example - TypeScript**
```ts
import { OfficeParserConfig, parseOfficeAsync, ParseOfficeResult, Block } from 'officeparser';
import * as fs from 'fs';

const config: OfficeParserConfig = {
    newlineDelimiter: " ",  // Separate new lines with a space instead of the default \n.
    ignoreNotes: true,      // Ignore notes while parsing presentation files like pptx or odp.
    extractImages: true     // Extract images from files.
}

// relative path is also fine => eg: files/myWorkSheet.ods
parseOfficeAsync("/Users/harsh/Desktop/files/mySlides.pptx", config)
    .then((data: ParseOfficeResult) => {
        const newText = data.text + " look, I can parse a powerpoint file";
        callSomeOtherFunction(newText);
        // Save images from blocks
        data.blocks
            .filter((block): block is Extract<Block, { type: 'image' }> => block.type === 'image')
            .forEach((image, index) => {
                fs.writeFileSync(`image_${index}.${image.mimeType.split('/')[1]}`, image.buffer);
            });
    })
    .catch(err => console.error(err));

// Search for a term in the parsed text.
function searchForTermInOfficeFile(searchterm: string, filepath: string): Promise<boolean> {
    return parseOfficeAsync(filepath)
        .then((data: ParseOfficeResult) => data.text.indexOf(searchterm) != -1)
}
```
\n**Please take note: I have breached convention in placing err as second argument in my callback but please understand that I had to do it to not break other people's existing modules.**

## Browser Usage
Download the bundle file available as part of the release asset.
Include this bundle file in your browser html file and access `parseOffice` and `parseOfficeAsync` under the **`officeParser`** namespace.

**Example**
```html
<head>
    ...
    <!-- Include bundle file in the script tag. -->
    <script src="officeParserBundle@5.1.0.js"></script>
</head>
<body>
    ...
    <input type="file" id="fileInput" />
    ...
    <script>
        document.getElementById('fileInput').addEventListener('change', async function(event) {
            const outputDiv = document.getElementById('output');
            const file = event.target.files[0];
            try {
                // Your configuration options for officeParser
                const config = {
                    outputErrorToConsole: false,
                    newlineDelimiter: '\n',
                    ignoreNotes: false,
                    putNotesAtLast: false,
                    extractImages: true
                };

                const arrayBuffer = await file.arrayBuffer();
                const result = await officeParser.parseOfficeAsync(arrayBuffer, config);
                // result contains the extracted text and blocks (text and images in document order).
                console.log(result.text);
                console.log(result.blocks);
            }
            catch (error) {
                // Handle error
            }
        });
    </script>
</body>
```


## Known Bugs
1. Inconsistency and incorrectness in the positioning of footnotes and endnotes in .docx files where the footnotes and endnotes would end up at the end of the parsed text whereas it would be positioned exactly after the referenced word in .odt files.
2. The charts and objects information of .odt files are not accurate and may end up showing a few NaN in some cases.
3. Extracting texts in browser bundles does not work for pdf files.
----------

**npm**
https://npmjs.com/package/officeparser

**github**
https://github.com/harshankur/officeParser
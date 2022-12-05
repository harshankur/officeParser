const officeParser = require("../officeParser");
const fs = require("fs");

// File names of test files and their text output content
// test file name style => test.<ext>
// test content output => test.<ext>.txt

/** Get filename for an extension */
function getFilename(ext, isContentFile = false) {
    return `test/test.${ext}` + (isContentFile ? `.txt` : '');
}

/** Create content file the test file with passed extension */
function createContentFile(ext) {
    return officeParser.parseOfficeAsync(getFilename(ext))
    .then(text => fs.writeFileSync(getFilename(ext, true), text))
}


process.argv.length == 3
    ? createContentFile(process.argv[2])
        .then(() => console.log(`Created text content file for ${process.argv[2]} => ${getFilename(process.argv[2], true)}`))
        .catch((error) => console.error(error))
    : console.error("Arguments missing")
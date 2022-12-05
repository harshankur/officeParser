const officeParser = require("../officeParser");
const fs = require("fs");

// File names of test files and their text output content
// test file name style => test.<ext>
// test content output => test.<ext>.txt

/** List of all supported extensions with office Parser */
const supportedExtensionTests = [
    {
        ext: "docx",
        testAvailable: true,
    },
    {
        ext: "xlsx",
        testAvailable: false,
    },
    {
        ext: "pptx",
        testAvailable: false,
    },
    {
        ext: "odt",
        testAvailable: true,
    },
    {
        ext: "odp",
        testAvailable: false,
    },
    {
        ext: "ods",
        testAvailable: false,
    },
]

/** Get filename for an extension */
function getFilename(ext, isContentFile = false) {
    return `test/test.${ext}` + (isContentFile ? `.txt` : '');
}

/** Run test for a passed extension */
function runTest(ext) {
    return officeParser.parseOfficeAsync(getFilename(ext))
    .then(text =>
        fs.readFileSync(getFilename(ext, true)) == text
            ? console.log(`[${ext}]=> Passed`)
            : console.log(`[${ext}]=> Failed`)
    )
}

// Run all test files with test content
supportedExtensionTests.forEach(test => test.testAvailable ? runTest(test.ext) : console.log(`[${test.ext}]=> Skipped`));

// @ts-check

const officeParser = require("../officeParser");
const fs = require("fs");
const supportedExtensions = require("../supportedExtensions");

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
        testAvailable: true,
    },
    {
        ext: "pptx",
        testAvailable: true,
    },
    {
        ext: "odt",
        testAvailable: true,
    },
    {
        ext: "odp",
        testAvailable: true,
    },
    {
        ext: "ods",
        testAvailable: true,
    },
    {
        ext: "pdf",
        testAvailable: true
    }
];

/** List of files with images for testing image extraction */
const imageTestFiles = [
    { filename: "docwithimages.docx", expectedImageCount: { exact: 3 } },
    { filename: "docwithimages.pdf", expectedImageCount: { exact: 3 } },
    { filename: "docwithimages.odt", expectedImageCount: { exact: 3 } },
    { filename: "presentationwithimages.pptx", expectedImageCount: { exact: 2 } },
    { filename: "presentationwithimages.odp", expectedImageCount: { exact: 2 } }
];

/** Config file for performing tests */
const config = {
    preserveTempFiles: true,
    outputErrorToConsole: true
}

/** Local list of supported extensions in test file */
const localSupportedExtensionsList = supportedExtensionTests.map(test => test.ext);

/** Get filename for an extension */
function getFilename(ext, isContentFile = false) {
    return `test/files/test.${ext}` + (isContentFile ? `.txt` : '');
}

/** Run test for a passed extension */
function runTest(ext, buffer, extractImages) {
    const testConfig = { ...config, extractImages: extractImages };
    return officeParser.parseOfficeAsync(buffer ? fs.readFileSync(getFilename(ext)) : getFilename(ext), testConfig)
        .then(result => {
            const expectedText = fs.readFileSync(getFilename(ext, true), 'utf8').trim();
            // Strip image placeholders from actual text for comparison (they're only added when extractImages=true)
            const actualText = result.text.replace(/<image [^>]+\/>\n?/g, '').trim();

            // Validate text content
            const textMatch = expectedText === actualText;

            // Validate image extraction from blocks
            const imageBlocks = result.blocks ? result.blocks.filter(b => b.type === 'image') : [];
            const imageCheck = !extractImages || (
                Array.isArray(result.blocks) &&
                ((['docx', 'pdf', 'pptx', 'odt', 'odp', 'ods'].includes(ext) && imageBlocks.length >= 0) ||
                 (!['docx', 'pdf', 'pptx', 'odt', 'odp', 'ods'].includes(ext) && imageBlocks.length === 0))
            );

            if (textMatch && imageCheck) {
                console.log(`[${ext.padEnd(4)}: ${buffer ? 'buffer' : 'file  '} | extractImages: ${extractImages}] => Passed`);
            } else {
                console.log(`[${ext.padEnd(4)}: ${buffer ? 'buffer' : 'file  '} | extractImages: ${extractImages}] => Failed (text: ${textMatch}, images: ${imageCheck})`);
            }
        })
        .catch(error => console.log("ERROR: " + error));
}

/** Check if all images are unique by comparing their buffer contents */
function areImagesUnique(imageBlocks) {
    const bufferHashes = new Set();
    for (const image of imageBlocks) {
        // Use buffer content as a simple hash
        const hash = image.buffer.toString('base64');
        if (bufferHashes.has(hash)) {
            return false; // Duplicate found
        }
        bufferHashes.add(hash);
    }
    return true;
}

/** Test image extraction for files with images */
async function runImageExtractionTest(testFile) {
    const testConfig = { ...config, extractImages: true };
    return officeParser.parseOfficeAsync(`test/files/${testFile.filename}`, testConfig)
        .then(result => {
            const imageBlocks = result.blocks.filter(b => b.type === 'image');
            // Validate image count
            const imageCount = imageBlocks.length;
            let imagesPassed = false;

            if (testFile.expectedImageCount.exact !== undefined) {
                imagesPassed = imageCount === testFile.expectedImageCount.exact;
            } else if (testFile.expectedImageCount.min !== undefined) {
                imagesPassed = imageCount >= testFile.expectedImageCount.min;
            }

            // Validate images are unique (not duplicates)
            const imagesUnique = areImagesUnique(imageBlocks);
            imagesPassed = imagesPassed && imagesUnique;

            // Validate text content (strip image placeholders for comparison)
            const expectedText = fs.readFileSync(`test/files/${testFile.filename}.txt`, 'utf8').trim();
            const actualText = result.text.replace(/<image [^>]+\/>\n?/g, '').trim();
            const textPassed = expectedText === actualText;

            // Overall pass/fail
            const passed = imagesPassed && textPassed;
            const status = passed ? 'Passed' : 'Failed';

            const imageDetails = testFile.expectedImageCount.exact !== undefined
                ? `expected: ${testFile.expectedImageCount.exact}, got: ${imageCount}`
                : `expected: >=${testFile.expectedImageCount.min}, got: ${imageCount}`;

            const uniqueInfo = imagesUnique ? '' : ' [DUPLICATES DETECTED]';
            console.log(`[${testFile.filename.padEnd(30)}] => ${status} (text: ${textPassed}, images: ${imagesPassed} - ${imageDetails}${uniqueInfo})`);
        })
        .catch(error => console.log(`[${testFile.filename.padEnd(30)}] => Error: ${error.message}`));
}

async function runAllTests() {
    console.log("\n=== Running standard format tests ===");
    for (let i = 0; i < supportedExtensionTests.length; i++)
    {
        const test = supportedExtensionTests[i];
        if (test.testAvailable) {
            await runTest(test.ext, false, false);
            await runTest(test.ext, true, false);
            await runTest(test.ext, false, true);
            await runTest(test.ext, true, true);
        }
        else
            console.log(`[${test.ext}]=> Skipped`);
    }

    console.log("\n=== Running image extraction tests ===");
    for (let i = 0; i < imageTestFiles.length; i++) {
        await runImageExtractionTest(imageTestFiles[i]);
    }
}

// Run all test files with test content if no argument passed.
if (process.argv.length == 2)
{
    // Test to check all items in local extension list are present in supportedExtensions.js file
    localSupportedExtensionsList
    .every(ext => supportedExtensions.includes(ext))
        ? console.log("All extensions in test files found in primary supportedExtensions.js file")
        : console.warn("Extension in test files missing from primary supportedExtensions.js file");

    // Test to check all items in supportedExtensions.js file are present in local extension list
    supportedExtensions
    .every(ext => localSupportedExtensionsList.includes(ext))
        ? console.log("All extensions in primary supportedExtensions.js file found in test file")
        : console.warn("Extension in primary supportedExtensions.js file missing from test file");

    runAllTests();
}
else if (process.argv.length == 3)
{
    if (localSupportedExtensionsList.includes(process.argv[2]))
        officeParser.parseOfficeAsync(getFilename(process.argv[2]), config)
            .then(result => console.log(result.text))
            .catch(error => console.log("ERROR: " + error))
    else
        console.error("The requested extension test is not currently available.");
}
else
    console.error("Invalid arguments");

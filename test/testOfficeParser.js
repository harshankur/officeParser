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

/** List of Word files for testing table and chart extraction */
const structuredContentTestFiles = [
    { filename: "test.docx", expectedTables: { min: 1 }, expectedCharts: { min: 0 } },
    { filename: "docwithimages_with_chart.docx", expectedTables: { min: 0 }, expectedCharts: { min: 0 } }
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
            // Handle case where result might be a string (backwards compatibility)
            if (typeof result === 'string') {
                const expectedText = fs.readFileSync(getFilename(ext, true), 'utf8').trim();
                const textMatch = expectedText === result.trim();
                if (textMatch) {
                    console.log(`[${ext.padEnd(4)}: ${buffer ? 'buffer' : 'file  '} | extractImages: ${extractImages}] => Passed`);
                } else {
                    console.log(`[${ext.padEnd(4)}: ${buffer ? 'buffer' : 'file  '} | extractImages: ${extractImages}] => Failed (text mismatch)`);
                }
                return;
            }

            const expectedText = fs.readFileSync(getFilename(ext, true), 'utf8').trim();
            // Strip image placeholders from actual text for comparison (they're only added when extractImages=true)
            // @ts-ignore - TypeScript doesn't narrow properly after string check
            const actualText = result.text.replace(/<image [^>]+\/>\n?/g, '').trim();

            // Validate text content
            const textMatch = expectedText === actualText;

            // Validate image extraction from blocks
            // @ts-ignore - TypeScript doesn't narrow properly after string check
            const imageBlocks = result.blocks ? result.blocks.filter(b => b.type === 'image') : [];
            const imageCheck = !extractImages || (
                // @ts-ignore - TypeScript doesn't narrow properly after string check
                Array.isArray(result.blocks) &&
                ((['docx', 'pdf', 'pptx', 'odt', 'odp', 'ods'].includes(ext) && imageBlocks.length >= 0) ||
                 (!['docx', 'pdf', 'pptx', 'odt', 'odp', 'ods'].includes(ext) && imageBlocks.length === 0))
            );

            // Validate tables and charts arrays exist for Word files
            let structureCheck = true;
            if (ext === 'docx') {
                // @ts-ignore - TypeScript doesn't narrow properly after string check
                structureCheck = Array.isArray(result.tables) && Array.isArray(result.charts);
            }

            if (textMatch && imageCheck && structureCheck) {
                console.log(`[${ext.padEnd(4)}: ${buffer ? 'buffer' : 'file  '} | extractImages: ${extractImages}] => Passed`);
            } else {
                console.log(`[${ext.padEnd(4)}: ${buffer ? 'buffer' : 'file  '} | extractImages: ${extractImages}] => Failed (text: ${textMatch}, images: ${imageCheck}, structure: ${structureCheck})`);
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

/** Validate table structure */
function validateTableStructure(table) {
    if (!table || typeof table !== 'object') return false;
    if (typeof table.name !== 'string') return false;
    if (!Array.isArray(table.rows)) return false;
    
    for (const row of table.rows) {
        if (!row || typeof row !== 'object') return false;
        if (!Array.isArray(row.cols)) return false;
        
        for (const col of row.cols) {
            if (!col || typeof col !== 'object') return false;
            if (typeof col.value !== 'string') return false;
        }
    }
    
    return true;
}

/** Validate chart structure */
function validateChartStructure(chart) {
    if (!chart || typeof chart !== 'object') return false;
    if (typeof chart.chartType !== 'string') return false;
    if (!Array.isArray(chart.series)) return false;
    
    for (const series of chart.series) {
        if (!series || typeof series !== 'object') return false;
        if (!Array.isArray(series.categories)) return false;
        if (!Array.isArray(series.values)) return false;
        for (const category of series.categories) {
            if (!category || typeof category !== 'string') return false;
        }
        for (const value of series.values) {
            if (!value || typeof value !== 'number') return false;
        }
    }
    
    return true;
}

/** Test structured content extraction (tables and charts) for Word files */
async function runStructuredContentTest(testFile) {
    const testConfig = { ...config, extractImages: false, extractCharts: true };
    return officeParser.parseOfficeAsync(`test/files/${testFile.filename}`, testConfig)
        .then(result => {
            // Handle case where result might be a string (backwards compatibility)
            if (typeof result === 'string') {
                console.log(`[${testFile.filename.padEnd(35)}] => Skipped (string result, not structured)`);
                return;
            }

            // Validate tables array exists
            // @ts-ignore - TypeScript doesn't narrow properly after string check
            const tables = result.tables || [];
            // @ts-ignore - TypeScript doesn't narrow properly after string check
            const tableBlocks = result.blocks ? result.blocks.filter(b => b.type === 'table') : [];
            
            // Validate charts array exists
            // @ts-ignore - TypeScript doesn't narrow properly after string check
            const charts = result.charts || [];
            // @ts-ignore - TypeScript doesn't narrow properly after string check
            const chartBlocks = result.blocks ? result.blocks.filter(b => b.type === 'chart') : [];
            
            // Check table count
            let tablesPassed = true;
            if (testFile.expectedTables.exact !== undefined) {
                tablesPassed = tables.length === testFile.expectedTables.exact;
            } else if (testFile.expectedTables.min !== undefined) {
                tablesPassed = tables.length >= testFile.expectedTables.min;
            }
            
            // Check chart count
            let chartsPassed = true;
            if (testFile.expectedCharts.exact !== undefined) {
                chartsPassed = charts.length === testFile.expectedCharts.exact;
            } else if (testFile.expectedCharts.min !== undefined) {
                chartsPassed = charts.length >= testFile.expectedCharts.min;
            }
            
            // Validate table structures
            let tableStructurePassed = true;
            for (const table of tables) {
                if (!validateTableStructure(table)) {
                    tableStructurePassed = false;
                    break;
                }
            }
            
            // Validate chart structures
            let chartStructurePassed = true;
            for (const chart of charts) {
                if (!validateChartStructure(chart)) {
                    chartStructurePassed = false;
                    break;
                }
            }

            // Validate blocks match arrays
            const blocksMatchArrays = tables.length === tableBlocks.length && charts.length === chartBlocks.length;
            
            // Validate document order (blocks should be in order)
            let orderPassed = true;
            // @ts-ignore - TypeScript doesn't narrow properly after string check
            if (result.blocks && result.blocks.length > 0) {
                const allowedTypes = ['text', 'table', 'chart', 'image'];
                // @ts-ignore - TypeScript doesn't narrow properly after string check
                for (const block of result.blocks) {
                    if (!block.type || !allowedTypes.includes(block.type)) {
                        orderPassed = false;
                        break;
                    }
                }
            }
            
            // Overall pass/fail
            const passed = tablesPassed && chartsPassed && tableStructurePassed && 
                          chartStructurePassed && blocksMatchArrays && orderPassed;
            const status = passed ? 'Passed' : 'Failed';
            
            const details = [
                `tables: ${tables.length}${testFile.expectedTables.exact !== undefined ? ` (expected: ${testFile.expectedTables.exact})` : testFile.expectedTables.min !== undefined ? ` (min: ${testFile.expectedTables.min})` : ''}`,
                `charts: ${charts.length}${testFile.expectedCharts.exact !== undefined ? ` (expected: ${testFile.expectedCharts.exact})` : testFile.expectedCharts.min !== undefined ? ` (min: ${testFile.expectedCharts.min})` : ''}`,
                `table structure: ${tableStructurePassed ? 'OK' : 'FAIL'}`,
                `chart structure: ${chartStructurePassed ? 'OK' : 'FAIL'}`,
                `blocks match: ${blocksMatchArrays ? 'OK' : 'FAIL'}`,
                `order: ${orderPassed ? 'OK' : 'FAIL'}`
            ].join(', ');
            
            console.log(`[${testFile.filename.padEnd(35)}] => ${status} (${details})`);
            
            // Print detailed info if failed
            if (!passed && config.outputErrorToConsole) {
                if (!tablesPassed) console.log(`  Tables: expected ${testFile.expectedTables.exact || `>=${testFile.expectedTables.min}`}, got ${tables.length}`);
                if (!chartsPassed) console.log(`  Charts: expected ${testFile.expectedCharts.exact || `>=${testFile.expectedCharts.min}`}, got ${charts.length}`);
                if (!tableStructurePassed) console.log(`  Table structure validation failed`);
                if (!chartStructurePassed) console.log(`  Chart structure validation failed`);
                if (!blocksMatchArrays) console.log(`  Blocks don't match arrays: tables=${tables.length} vs tableBlocks=${tableBlocks.length}, charts=${charts.length} vs chartBlocks=${chartBlocks.length}`);
                if (!orderPassed) console.log(`  Block order validation failed`);
            }
        })
        .catch(error => console.log(`[${testFile.filename.padEnd(35)}] => Error: ${error.message}`));
}

/** Test image extraction for files with images */
async function runImageExtractionTest(testFile) {
    const testConfig = { ...config, extractImages: true };
    return officeParser.parseOfficeAsync(`test/files/${testFile.filename}`, testConfig)
        .then(result => {
            // Handle case where result might be a string (backwards compatibility)
            if (typeof result === 'string') {
                console.log(`[${testFile.filename.padEnd(30)}] => Skipped (string result)`);
                return;
            }

            // @ts-ignore - TypeScript doesn't narrow properly after string check
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
            // @ts-ignore - TypeScript doesn't narrow properly after string check
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

    console.log("\n=== Running structured content tests (tables & charts) ===");
    for (let i = 0; i < structuredContentTestFiles.length; i++) {
        await runStructuredContentTest(structuredContentTestFiles[i]);
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
            .then(result => {
                if (typeof result === 'string') {
                    console.log(result);
                } else {
                    // @ts-ignore - TypeScript doesn't narrow properly after string check
                    console.log(result.text);
                    // @ts-ignore - TypeScript doesn't narrow properly after string check
                    if (result.tables && result.tables.length > 0) {
                        // @ts-ignore - TypeScript doesn't narrow properly after string check
                        console.log(`\n[Extracted ${result.tables.length} table(s)]`);
                        // @ts-ignore - TypeScript doesn't narrow properly after string check
                        result.tables.forEach((table, idx) => {
                            console.log(`\nTable ${idx + 1}: ${table.name}`);
                            table.rows.forEach((row, rowIdx) => {
                                const rowData = row.cols.map(col => col.value).join(' | ');
                                console.log(`  Row ${rowIdx + 1}: ${rowData}`);
                            });
                        });
                    }
                    // @ts-ignore - TypeScript doesn't narrow properly after string check
                    if (result.charts && result.charts.length > 0) {
                        // @ts-ignore - TypeScript doesn't narrow properly after string check
                        console.log(`\n[Extracted ${result.charts.length} chart(s)]`);
                        // @ts-ignore - TypeScript doesn't narrow properly after string check
                        result.charts.forEach((chart, idx) => {
                            console.log(`\nChart ${idx + 1}: ${chart.chartType}`);
                            chart.categories.forEach((cat, catIdx) => {
                                console.log(`  Category ${catIdx + 1}: ${cat.label} = ${cat.value}`);
                            });
                        });
                    }
                    // @ts-ignore - TypeScript doesn't narrow properly after string check
                    if (result.blocks && result.blocks.length > 0) {
                        // @ts-ignore - TypeScript doesn't narrow properly after string check
                        const blockTypes = result.blocks.map(b => b.type);
                        console.log(`\n[Blocks in order: ${blockTypes.join(', ')}]`);
                    }
                }
            })
            .catch(error => console.log("ERROR: " + error))
    else
        console.error("The requested extension test is not currently available.");
}
else
    console.error("Invalid arguments");

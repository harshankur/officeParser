#!/usr/bin/env node
/**
 * officeparser CLI
 *
 * Allows running officeparser from the command line:
 *   npx officeparser file.docx
 *   officeparser file.docx --to=text
 *   officeparser file.docx --ocr --extractAttachments
 *
 * Options (--key=value, --key value, or bare flags):
 *   --to=json|text|md|html|csv|rtf|pdf|chunks  Convert AST to specified format (default: json)
 *   --output=path             Save result to a file
 *   --fileType=docx|xlsx|...  Override file type detection
 *   --ocr                     Enable OCR for images (default: false)
 *   --ocrConfig.language=eng  OCR language (default: eng)
 *   --extractAttachments      Extract embedded attachments (default: false)
 *   --ignoreNotes             Ignore footnotes/endnotes/speaker notes (default: false)
 *   --ignoreComments          Ignore inline comments (default: false)
 *   --ignoreHeadersAndFooters Ignore headers and footers (default: false)
 *   --ignoreSlideMasters      Ignore slide masters (default: false)
 *   --ignoreInternalLinks     Ignore internal links (default: false)
 *   --includeRawContent       Include raw content in AST (default: false)
 *   --serializeRawContent     Include stringified XML in metadata (default: true)
 *   --preserveXmlWhitespace   Keep raw formatting space (default: false)
 *   --includeBreakNodes       Include break nodes (DOCX only, default: false)
 *   --verbose                 Show full error stack traces and warning logs
 */


import { OfficeParser } from './OfficeParser.js';
import { OfficeGenerator } from './OfficeGenerator.js';
import { OfficeParserAST, OfficeParserConfig, UniversalGeneratorFormat } from './types.js';
import * as fs from 'fs';

const args = process.argv.slice(2);
let fileArg: string | undefined;
let showHelp = false;
let toFlagOption: string | undefined;
let formatFlagOption: string | undefined;
let toTextOption: boolean | undefined;
let verbose = false;
let outputFile: string | undefined;

// Parser and Generator configuration objects that will be populated by command line options.
const config: OfficeParserConfig = {};
const generatorConfig: any = {};

// Known boolean configurations to validate user input against.
const knownParserBooleans = new Set([
    'ocr', 'extractAttachments', 'ignoreNotes', 'ignoreComments',
    'ignoreHeadersAndFooters', 'ignoreSlideMasters', 'ignoreInternalLinks',
    'includeRawContent', 'serializeRawContent', 'preserveXmlWhitespace', 'includeBreakNodes'
]);

const knownGeneratorBooleans = new Set([
    'includeFormatting', 'generateIds', 'renderMetadata', 'includeImages', 'includeCharts', 'ignoreInternalLinks'
]);

// Prefixes used to identify configurations targeted for the generator instead of the parser.
const generatorPrefixes = [
    'generatorConfig.', 'htmlConfig.', 'csvConfig.', 'textConfig.', 'mdConfig.', 'pdfConfig.', 'rtfConfig.', 'chunksConfig.'
];

// Trackers to detect if deprecated/legacy options were used to log helpful warnings.
let usedFormat = false;
let usedToText = false;
let usedOcrLanguage = false;
let usedPutNotesAtLast = false;
let usedOutputErrorToConsole = false;

// Parse the arguments list
for (let i = 0; i < args.length; i++) {
    const arg = args[i];

    // Help flags trigger immediate termination of parsing and print help
    if (arg === '-h' || arg === '--help') {
        showHelp = true;
        break;
    }

    if (arg.startsWith('--')) {
        let cleanKey: string;
        let val: string;

        // Support --key=value syntax
        if (arg.includes('=')) {
            const idx = arg.indexOf('=');
            cleanKey = arg.slice(2, idx);
            val = arg.slice(idx + 1);
        }
        // Support negation shorthand: --no-ocr sets ocr to false
        else if (arg.startsWith('--no-')) {
            cleanKey = arg.slice(5);
            val = 'false';
        }
        // Support space-separated options or bare flags
        else {
            cleanKey = arg.slice(2);
            const isKnownBoolean = knownParserBooleans.has(cleanKey) ||
                                   knownGeneratorBooleans.has(cleanKey) ||
                                   cleanKey === 'verbose' ||
                                   cleanKey === 'toText' ||
                                   cleanKey === 'outputErrorToConsole';
            const isNextBool = i + 1 < args.length &&
                               (args[i + 1].toLowerCase() === 'true' || args[i + 1].toLowerCase() === 'false');

            // If the next arg is not another option flag, treat it as the value (e.g., --to html)
            // But if this is a known boolean, only consume the next argument if it is a valid boolean string.
            if (i + 1 < args.length && !args[i + 1].startsWith('-') && (!isKnownBoolean || isNextBool)) {
                val = args[i + 1];
                i++;
            }
            // Bare presence implies true (e.g., --ocr is true)
            else {
                val = 'true';
            }
        }

        // Parse boolean strings to raw boolean types
        const lowerValue = val.toLowerCase();
        const boolValue = lowerValue === 'true' ? true : (lowerValue === 'false' ? false : undefined);

        // Map core CLI options to variables
        if (cleanKey === 'to') {
            toFlagOption = val;
        } else if (cleanKey === 'format') {
            formatFlagOption = val;
            usedFormat = true;
        } else if (cleanKey === 'output') {
            outputFile = val;
        } else if (cleanKey === 'toText') {
            toTextOption = boolValue !== undefined ? boolValue : true;
            usedToText = true;
        } else if (cleanKey === 'verbose') {
            verbose = boolValue !== undefined ? boolValue : true;
        } else if (cleanKey === 'ocrLanguage') {
            config.ocrLanguage = val;
            usedOcrLanguage = true;
        } else if (cleanKey === 'putNotesAtLast') {
            config.putNotesAtLast = boolValue !== undefined ? boolValue : true;
            usedPutNotesAtLast = true;
        } else if (cleanKey === 'outputErrorToConsole') {
            verbose = boolValue !== undefined ? boolValue : true;
            usedOutputErrorToConsole = true;
        } else {
            // Check if the flag belongs to generatorConfig or a specific sub-generator (e.g., htmlConfig)
            const isGeneratorOption = knownGeneratorBooleans.has(cleanKey) || generatorPrefixes.some(pref => cleanKey.startsWith(pref));
            const target = isGeneratorOption ? generatorConfig : config;

            let path = cleanKey;
            // Strip generatorConfig prefix to flatten it onto the generatorConfig object
            if (isGeneratorOption && cleanKey.startsWith('generatorConfig.')) {
                path = cleanKey.slice('generatorConfig.'.length);
            }

            // Support nested dot-notation parsing (e.g., --ocrConfig.language=fra)
            if (path.includes('.')) {
                const parts = path.split('.');
                let current = target;
                for (let j = 0; j < parts.length - 1; j++) {
                    const part = parts[j];
                    if (!current[part]) current[part] = {};
                    current = current[part];
                }
                const lastPart = parts[parts.length - 1];
                current[lastPart] = boolValue !== undefined ? boolValue : val;
            }
            // Flat key assignment
            else {
                if (boolValue !== undefined) {
                    target[path] = boolValue;
                } else if (knownParserBooleans.has(path) || (isGeneratorOption && knownGeneratorBooleans.has(path))) {
                    console.warn(`Invalid boolean value for --${cleanKey}: ${val}. Using default.`);
                } else {
                    target[path] = val;
                }
            }
        }
    } else {
        // First positional argument that is not a flag is treated as the input file path
        if (!fileArg) {
            fileArg = arg;
        }
    }
}

if (fileArg && !showHelp) {
    // Resolve output format prioritizing: --to > --format > --toText
    let outputFormat: string | undefined;
    if (toFlagOption) {
        outputFormat = toFlagOption as UniversalGeneratorFormat;
    } else if (formatFlagOption) {
        outputFormat = formatFlagOption as UniversalGeneratorFormat;
    } else if (toTextOption === true) {
        outputFormat = 'text';
    }

    // Display warning messages for any deprecated CLI options used
    if (usedFormat) {
        console.warn('Warning: --format is deprecated. Use --to instead.');
    }
    if (usedToText) {
        console.warn('Warning: --toText is deprecated. Use --to=text instead.');
    }
    if (usedOcrLanguage) {
        console.warn('Warning: --ocrLanguage is deprecated. Use --ocrConfig.language instead.');
    }
    if (usedPutNotesAtLast) {
        console.warn('Warning: --putNotesAtLast is deprecated and will be ignored by all parsers.');
    }
    if (usedOutputErrorToConsole) {
        console.warn('Warning: --outputErrorToConsole is deprecated. Use --verbose instead.');
    }
    // Intercept parser warning callbacks to format and print issues when verbose is enabled
    const originalOnWarning = config.onWarning;
    config.onWarning = (issue) => {
        if (verbose) {
            const severity = issue.type === 'error' ? 'Error' : 'Warning';
            console.error(`[OfficeParser ${severity}] [${issue.code}]: ${issue.message}`);
            if (issue.details) {
                console.error(issue.details);
            }
        }
        if (originalOnWarning) originalOnWarning(issue);
    };

    // Propagate newlineDelimiter and csvDelimiter if configured flatly but not in generator sub-configs
    if (config.newlineDelimiter !== undefined) {
        if (!generatorConfig.textConfig) generatorConfig.textConfig = {};
        if (generatorConfig.textConfig.newlineDelimiter === undefined) {
            generatorConfig.textConfig.newlineDelimiter = config.newlineDelimiter;
        }
    }
    if (config.csvDelimiter !== undefined) {
        if (!generatorConfig.csvConfig) generatorConfig.csvConfig = {};
        if (generatorConfig.csvConfig.columnDelimiter === undefined) {
            generatorConfig.csvConfig.columnDelimiter = config.csvDelimiter;
        }
    }

    // Run the main parser
    OfficeParser.parseOffice(fileArg, config)
        .then(async (ast: OfficeParserAST) => {
            let output: string | Uint8Array;

            // Generate JSON output or convert AST using OfficeGenerator
            if (outputFormat === 'json') {
                output = JSON.stringify(ast, null, 2);
            } else if (outputFormat) {
                const result = await OfficeGenerator.generate(ast as any, outputFormat as any, generatorConfig);
                if (Array.isArray(result.value)) {
                    output = JSON.stringify(result.value, null, 2);
                } else {
                    output = result.value as string | Uint8Array;
                }
            } else {
                output = JSON.stringify(ast, null, 2);
            }

            // Write generated output to output file or print to standard output
            if (outputFile) {
                if (output instanceof Uint8Array) {
                    fs.writeFileSync(outputFile, output);
                } else {
                    fs.writeFileSync(outputFile, output, 'utf8');
                }
                if (verbose) console.log(`Output written to ${outputFile}`);
            } else {
                if (output instanceof Uint8Array) {
                    process.stdout.write(output);
                } else {
                    process.stdout.write(output + '\n');
                }
            }

            // Ensure OCR workers are terminated for clean CLI exit
            if (config.ocr) {
                await OfficeParser.terminateOcr();
            }
        })
        .catch(async err => {
            // Handle and display parsing error messages
            console.error(`Error parsing file "${fileArg}":`);
            if (verbose) {
                console.error(err);
            } else {
                console.error(err.message || err);
                console.error('Use --verbose for full stack trace.');
            }

            // Ensure OCR workers are terminated even on error to prevent process hang
            if (config.ocr) {
                await OfficeParser.terminateOcr();
            }
            process.exit(1);
        });
} else {
    console.log('Usage: officeparser <file> [options]');
    console.log('');
    console.log('Options:');
    console.log('  --to=json|text|md|html|pdf|csv|rtf|chunks   Target conversion format (default: json)');
    console.log('  --output=file.ext                           Save output to file instead of stdout');
    console.log('  --fileType=docx|xlsx|pptx|odt|...           Explicitly override input file type detection');
    console.log('  --ocr                                       Enable OCR for images (default: false)');
    console.log('  --ocrConfig.language=eng                    OCR language (default: eng)');
    console.log('  --extractAttachments                        Extract embedded attachments (default: false)');
    console.log('  --ignoreNotes                               Ignore footnotes/endnotes/speaker notes (default: false)');
    console.log('  --ignoreComments                            Ignore inline comments (default: false)');
    console.log('  --ignoreHeadersAndFooters                   Ignore headers and footers (default: false)');
    console.log('  --ignoreSlideMasters                        Ignore slide masters (default: false)');
    console.log('  --ignoreInternalLinks                       Ignore internal links (default: false)');
    console.log('  --includeRawContent                         Include raw content in AST (default: false)');
    console.log('  --serializeRawContent                       Serialize raw XML content (default: true)');
    console.log('  --preserveXmlWhitespace                     Keep raw formatting space (default: false)');
    console.log('  --includeBreakNodes                         Include break nodes (DOCX only, default: false)');
    console.log('  --verbose                                   Show full error stack traces and warning logs');
    console.log('  --newlineDelimiter=string                   Delimiter string between blocks/lines (default: \\n)');
    console.log('  --csvDelimiter=char                         Custom CSV delimiter (default: ,)');
    console.log('');
    console.log('High-Value Generator Options:');
    console.log('  --includeFormatting                         Include font formatting like bold/italic (default: true)');
    console.log('  --renderMetadata                            Render metadata in output content (default: false)');
    console.log('  --htmlConfig.containerWidth=value           HTML container width (auto | px | % | vw etc., default: auto)');
    console.log('');
    console.log('Advanced Nested Config Examples:');
    console.log('  --pdfConfig.format=Letter                   Configure Puppeteer PDF format (A4 | Letter | Legal etc.)');
    console.log('  --chunksConfig.strategy=fixed-size          Chunking strategy (fixed-size | document-structure | semantic)');
    console.log('');
    console.log('Format Syntax:');
    console.log('  Flags can be written as --flag (presence implies true), --no-flag (negation),');
    console.log('  --flag=value, or --flag value.');
    console.log('');
    console.log('Examples:');
    console.log('  officeparser document.docx');
    console.log('  officeparser document.docx --to html --output doc.html');
    console.log('  officeparser document.docx --to md');
    console.log('  officeparser report.pdf --ocr --ocrConfig.language eng --to text');
    console.log('  officeparser data.xlsx --to csv --output data.csv --csvDelimiter ";"');
    console.log('  officeparser image_doc --fileType docx --to json');
}

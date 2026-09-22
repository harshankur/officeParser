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
 *   --to=json|text|md|html|csv|rtf|pdf|docx|odt|tex|epub|chunks  Convert AST to specified format (default: json)
 *   --output=path             Save result to a file
 *   --fileType=docx|xlsx|...  Override file type detection
 *   --ocr                     Enable OCR for images (default: false); also requires --extractAttachments
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
 *   --includeBreakNodes       Include break nodes (DOCX & ODF, default: false)
 *   --ignorePageGeometry      Omit per-node bounding boxes and page dimensions (default: false)
 *   --password=secret         Password for an encrypted document (PDF, OOXML, or ODF); or set
 *                             OFFICEPARSER_PASSWORD to keep the secret out of the process list
 *                             and shell history
 *   --includeImages=<mode>    image-only | image+ocr-text | ocr-text-only | none (default: image-only)
 *   --maxInlineImageBytes=N   Largest image inlined as a data: URI by HTML/Markdown (default: 1500000)
 *   --pdfParserConfig.useTags=false     Geometry-only PDF structure (default: true)
 *   --pdfConfig.engine=native Render PDF with pdf-lib instead of Puppeteer (default: html)
 *   --verbose                 Show full error stack traces and warning logs
 *
 * Removed in v8 (these exit with a migration hint): --toText, --ocrLanguage,
 * --putNotesAtLast, --outputErrorToConsole. Run with --help for the full option list.
 */


import { OfficeParser } from './OfficeParser.js';
import { OfficeGenerator } from './OfficeGenerator.js';
import { OfficeParserAST, OfficeParserConfig, OfficeWarningType, UniversalGeneratorFormat } from './types.js';
import * as fs from 'fs';

const args = process.argv.slice(2);
let fileArg: string | undefined;
let showHelp = false;
let toFlagOption: string | undefined;
let formatFlagOption: string | undefined;
let verbose = false;
let outputFile: string | undefined;

// Parser and Generator configuration objects that will be populated by command line options.
const config: OfficeParserConfig = {};
const generatorConfig: any = {};

// Known boolean configurations to validate user input against.
const knownParserBooleans = new Set([
    'ocr', 'extractAttachments', 'ignoreNotes', 'ignoreComments',
    'ignoreHeadersAndFooters', 'ignoreSlideMasters', 'ignoreInternalLinks',
    'includeRawContent', 'serializeRawContent', 'preserveXmlWhitespace', 'includeBreakNodes',
    'ignorePageGeometry',
    // Dotted boolean keys are listed so a bare `--group.flag` does not swallow the following file
    // argument as its value (the isKnownBoolean check keys off the full dotted name).
    'htmlParserConfig.preserveAttributes', 'htmlParserConfig.preserveIframes', 'htmlParserConfig.embedFolkForms',
    'pdfParserConfig.useTags', 'pdfParserConfig.detectColumns',
    'pdfParserConfig.mergeHyphenatedWords', 'pdfParserConfig.normalizeText', 'pdfParserConfig.extractTextColor',
    'ocrConfig.preserveLayout',
]);

const knownGeneratorBooleans = new Set([
    'includeFormatting', 'generateIds', 'renderMetadata', 'includeImages', 'includeCharts', 'ignoreInternalLinks',
    // Dotted generator booleans, same rationale as the parser set above.
    'pdfConfig.tagged', 'pdfConfig.outline', 'pdfConfig.landscape', 'pdfConfig.printBackground', 'pdfConfig.displayHeaderFooter',
    'docxConfig.landscape', 'odtConfig.landscape',
    'texConfig.standalone', 'texConfig.bundle', 'texConfig.numberSections', 'texConfig.landscape',
    'textConfig.preserveLayout', 'textConfig.renderNotes',
    'htmlConfig.standalone', 'htmlConfig.sourceAttributes', 'htmlConfig.gatedEmbeds',
    'mdConfig.fallbackToHtml', 'mdConfig.fallbackToHtml.inlineFormatting',
]);

/**
 * Keys that are listed as booleans above but also accept a fixed set of string values, mapped to
 * every value the space-separated form may consume. Without this, `--includeImages image+ocr-text`
 * reads as a bare boolean and leaves the mode behind as a stray positional, which the CLI then
 * takes for the input file.
 */
const knownEnumValues: Record<string, Set<string>> = {
    includeImages: new Set(['image-only', 'image+ocr-text', 'ocr-text-only', 'none']),
};

/**
 * Options removed in v8, mapped to the message shown before exiting non-zero. Accepting them
 * silently is worse than rejecting them: `--ocrLanguage=deu` would run English OCR and
 * `--putNotesAtLast` would produce a document whose notes are somewhere else entirely, with no
 * indication either flag did nothing.
 */
const REMOVED_CLI_FLAGS: Record<string, string> = {
    toText: '--toText was removed in v8. Use --to=text instead.',
    ocrLanguage: '--ocrLanguage was removed in v8. Use --ocrConfig.language instead.',
    putNotesAtLast: '--putNotesAtLast was removed in v8. Notes are attached to the node they belong to (node.notes) and rendered in place.',
    outputErrorToConsole: '--outputErrorToConsole was removed in v8. Use --verbose to print warnings and errors.',
};

// Prefixes used to identify configurations targeted for the generator instead of the parser.
const generatorPrefixes = [
    'generatorConfig.', 'htmlConfig.', 'csvConfig.', 'textConfig.', 'mdConfig.', 'pdfConfig.', 'rtfConfig.', 'docxConfig.', 'odtConfig.', 'texConfig.', 'chunksConfig.'
];

// Trackers to detect if deprecated/legacy options were used to log helpful warnings.
let usedFormat = false;

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
                                   cleanKey === 'verbose';
            const nextValue = i + 1 < args.length ? args[i + 1].toLowerCase() : '';
            const isNextBool = i + 1 < args.length &&
                               (nextValue === 'true' || nextValue === 'false' ||
                                knownEnumValues[cleanKey]?.has(nextValue) === true);

            // If the next arg is not another option flag, treat it as the value (e.g., --to html)
            // But if this is a known boolean, only consume the next argument if it is a valid boolean
            // string, or one of the enum values that key also accepts (e.g. --includeImages none).
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
        } else if (REMOVED_CLI_FLAGS[cleanKey]) {
            console.error(`Error: ${REMOVED_CLI_FLAGS[cleanKey]}`);
            process.exit(1);
        } else if (cleanKey === 'verbose') {
            verbose = boolValue !== undefined ? boolValue : true;
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
                } else if (path === 'includeImages') {
                    // includeImages also accepts an ImageMode string (e.g. --includeImages=image+ocr-text);
                    // it stays in knownGeneratorBooleans so a bare flag is true and never eats a positional.
                    target[path] = val;
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
    // Resolve output format prioritizing: --to > --format
    let outputFormat: string | undefined;
    if (toFlagOption) {
        outputFormat = toFlagOption as UniversalGeneratorFormat;
    } else if (formatFlagOption) {
        outputFormat = formatFlagOption as UniversalGeneratorFormat;
    }

    // Display warning messages for any deprecated CLI options used
    if (usedFormat) {
        console.warn('Warning: --format is deprecated. Use --to instead.');
    }
    // Intercept parser warning callbacks to format and print issues when verbose is enabled.
    // An unrecognized option is the exception: it reports a mistake in the command that was just
    // typed, not a detail of the document being parsed, so it always prints. Hiding it behind
    // --verbose is how a misspelled or renamed flag ends up silently doing nothing.
    const originalOnWarning = config.onWarning;
    config.onWarning = (issue) => {
        if (verbose || issue.code === OfficeWarningType.UNRECOGNIZED_CONFIG_OPTION) {
            const severity = issue.type === 'error' ? 'Error' : 'Warning';
            console.error(`[OfficeParser ${severity}] [${issue.code}]: ${issue.message}`);
            // The message already names every offending key, so the raw details object is only
            // noise unless the caller asked for the full log.
            if (verbose && issue.details) {
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

    // Password from the environment, so a script need not put a secret on the command line (where it
    // is visible in the process list / shell history). An explicit --password still wins.
    if (config.password === undefined && process.env.OFFICEPARSER_PASSWORD) {
        config.password = process.env.OFFICEPARSER_PASSWORD;
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
    console.log('  --to=json|text|md|html|pdf|csv|rtf|docx|odt|tex|epub|chunks  Target conversion format (default: json)');
    console.log('  --output=file.ext                           Save output to file instead of stdout');
    console.log('  --fileType=docx|xlsx|pptx|odt|...           Explicitly override input file type detection');
    console.log('  --ocr                                       Enable OCR for images (default: false; also requires --extractAttachments)');
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
    console.log('  --includeBreakNodes                         Include break nodes (DOCX & ODF, default: false)');
    console.log('  --ignorePageGeometry                        Omit per-node bounding boxes and page dimensions (default: false)');
    console.log('  --verbose                                   Show full error stack traces and warning logs');
    console.log('  --newlineDelimiter=string                   Delimiter string between blocks/lines (default: \\n)');
    console.log('  --csvDelimiter=char                         Custom CSV delimiter (default: ,)');
    console.log('  --password=secret                           Password for an encrypted document (PDF, OOXML, or ODF)');
    console.log('                                              (or set OFFICEPARSER_PASSWORD to keep it out of the process list)');
    console.log('  --htmlParserConfig.preserveIframes          Keep non-YouTube <iframe> embeds (dropped by default)');
    console.log('  --ocrConfig.preserveLayout=false            Flatten OCR text instead of keeping its line layout (default: true)');
    console.log('');
    console.log('PDF Parser Options (pdfParserConfig.*):');
    console.log('  --pdfParserConfig.useTags=false             Ignore the tagged-structure tree, use geometry only (default: true)');
    console.log('  --pdfParserConfig.detectColumns=false       Disable multi-column reading-order detection (default: true)');
    console.log('  --pdfParserConfig.pageRange=1-3,7           Parse only the given pages (default: all)');
    console.log('  --pdfParserConfig.headingDetection=off      Heading detection: auto | font-size | off (default: auto)');
    console.log('  --pdfParserConfig.mergeHyphenatedWords=false Keep words hyphenated across line breaks (default: true)');
    console.log('  --pdfParserConfig.normalizeText=false       Skip Unicode/ligature normalization of PDF text (default: true)');
    console.log('  --pdfParserConfig.extractTextColor          Record each PDF run\'s fill colour in formatting.color (default: true; set false to skip)');
    console.log('');
    console.log('High-Value Generator Options:');
    console.log('  --includeFormatting                         Include font formatting like bold/italic (default: true)');
    console.log('  --renderMetadata                            Render metadata in output content (default: false)');
    console.log('  --includeImages=<mode>                      Image rendering: image-only | image+ocr-text | ocr-text-only | none (default: image-only)');
    console.log('  --maxInlineImageBytes=1500000               Largest image HTML/Markdown inlines as a data: URI (0 = never, default: 1500000)');
    console.log('  --htmlConfig.containerWidth=value           HTML container width (auto | px | % | vw etc., default: auto)');
    console.log('  --htmlConfig.sourceAttributes               Carry each rich node\'s source in a data-* attribute (default: false)');
    console.log('  --textConfig.pageSeparator=string           Separator written between pages in text output (default: \\n)');
    console.log('');
    console.log('Advanced Nested Config Examples:');
    console.log('  --pdfConfig.engine=native                   PDF engine: html (Puppeteer) | native (pdf-lib, no browser) (default: html)');
    console.log('  --pdfConfig.format=Letter                   PDF paper format (A4 | Letter | Legal | A3 ...), both engines');
    console.log('  --chunksConfig.strategy=fixed-size          Chunking strategy (fixed-size | document-structure | semantic)');
    console.log('  --mdConfig.dialect=github                   Markdown dialect (extended | github | gitlab | obsidian | pandoc | commonmark)');
    console.log('  --mdConfig.fallbackToHtml=false             Disable HTML fallback for unsupported Markdown features (default: true)');
    console.log('  --mdConfig.fallbackToHtml.inlineFormatting  Round-trip inline color/highlight/font-size as <span style> (opt-in, default: false)');
    console.log('  --texConfig.bundle                          LaTeX: write a zip of main.tex plus its images/ (default: false, .tex only)');
    console.log('  --texConfig.documentClass=report            LaTeX class: auto | article | report | book | beamer (default: auto)');
    console.log('  --texConfig.standalone=false                LaTeX: emit the body only, without the preamble (default: true)');
    console.log('');
    console.log('Format Syntax:');
    console.log('  Flags can be written as --flag (presence implies true), --no-flag (negation),');
    console.log('  --flag=value, or --flag value.');
    console.log('  --includeImages accepts both forms: --includeImages=image+ocr-text or');
    console.log('  --includeImages image+ocr-text (a bare --includeImages still means "on").');
    console.log('');
    console.log('Examples:');
    console.log('  officeparser document.docx');
    console.log('  officeparser document.docx --to html --output doc.html');
    console.log('  officeparser document.docx --to md');
    console.log('  officeparser report.pdf --ocr --extractAttachments --ocrConfig.language eng --to text');
    console.log('  officeparser data.xlsx --to csv --output data.csv --csvDelimiter ";"');
    console.log('  officeparser document.docx --extractAttachments --to epub --output document.epub');
    console.log('  officeparser notes.md --extractAttachments --to odt --output notes.odt');
    console.log('  officeparser report.docx --extractAttachments --to tex --texConfig.bundle --output report.zip');
    console.log('  officeparser paper.tex --to docx --output paper.docx         (or an Overleaf project .zip)');
    console.log('  officeparser image_doc --fileType docx --to json');
}

#!/usr/bin/env node
/**
 * officeparser CLI
 *
 * Allows running officeparser from the command line:
 *   npx officeparser file.docx
 *   officeparser file.docx --toText=true
 *   officeparser file.docx --ocr=true --extractAttachments=true
 *
 * Options (--key=value):
 *   --format=json|text|md|html|csv|rtf|pdf|chunks  Convert AST to specified format
 *   --output=path             Save result to a file
 *   --toText=true             Legacy flag for plain text output
 *   --ocr=true                Enable OCR for images
 *   --ocrLanguage=eng         OCR language (default: eng)
 *   --extractAttachments=true Extract embedded attachments
 *   --ignoreNotes=true        Ignore footnotes/endnotes
 *   --putNotesAtLast=true     Move notes to end of document
 *   --includeRawContent=true  Include raw content in AST
 *   --outputErrorToConsole=true  Log errors to console
 */

import { OfficeParser } from './OfficeParser.js';
import { OfficeGenerator } from './OfficeGenerator.js';
import { OfficeParserAST, OfficeParserConfig, UniversalGeneratorFormat } from './types.js';
import * as fs from 'fs';

const args = process.argv.slice(2);
let fileArg: string | undefined;
let toText = false;
let verbose = false;
let outputFormat: UniversalGeneratorFormat | undefined;
let outputFile: string | undefined;
const configArgs: string[] = [];

function isConfigOption(arg: string) {
    return arg.startsWith('--') && arg.includes('=');
}

args.forEach(arg => {
    if (isConfigOption(arg)) {
        configArgs.push(arg);
    } else if (!fileArg) {
        fileArg = arg;
    }
});

if (fileArg) {
    const config: OfficeParserConfig = {};
    configArgs.forEach(arg => {
        const [key, value] = arg.split('=');
        const cleanKey = key.replace('--', '');

        const lowerValue = value.toLowerCase();
        const boolValue = lowerValue === 'true' ? true : (lowerValue === 'false' ? false : undefined);

        const knownBooleans = new Set([
            'toText', 'ocr', 'extractAttachments', 'ignoreNotes', 'putNotesAtLast', 
            'includeRawContent', 'outputErrorToConsole', 'serializeRawContent', 
            'preserveXmlWhitespace', 'includeBreakNodes', 'verbose'
        ]);

        if (cleanKey === 'format') {
            outputFormat = value as UniversalGeneratorFormat;
        } else if (cleanKey === 'output') {
            outputFile = value;
        } else {
            if (boolValue !== undefined) {
                if (cleanKey === 'toText') toText = boolValue;
                else if (cleanKey === 'verbose') {
                    verbose = boolValue;
                    if (verbose) config.outputErrorToConsole = true;
                } else {
                    // @ts-ignore
                    config[cleanKey] = boolValue;
                }
            } else if (knownBooleans.has(cleanKey)) {
                console.warn(`Invalid boolean value for --${cleanKey}: ${value}. Using default.`);
            } else {
                // @ts-ignore
                config[cleanKey] = value;
            }
        }
    });

    OfficeParser.parseOffice(fileArg, config)
        .then(async (ast: OfficeParserAST) => {
            let output: string | Uint8Array;

            if (outputFormat) {
                const result = await OfficeGenerator.generate(ast as any, outputFormat as any);
                if (Array.isArray(result.value)) {
                    output = JSON.stringify(result.value, null, 2);
                } else {
                    output = result.value as string | Uint8Array;
                }
            } else if (toText) {
                output = ast.toText();
            } else {
                output = JSON.stringify(ast, null, 2);
            }

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
            console.error(`Error parsing file "${fileArg}":`);
            if (verbose) {
                console.error(err);
            } else {
                console.error(err.message || err);
                console.error('Use --verbose=true for full stack trace.');
            }

            // Ensure OCR workers are terminated even on error
            if (config.ocr) {
                await OfficeParser.terminateOcr();
            }
            process.exit(1);
        });
} else {
    console.log('Usage: officeparser <file> [--option=value]');
    console.log('');
    console.log('Options:');
    console.log('  --format=json|md|html|rtf|csv|text|pdf|chunks  Convert to specified format');
    console.log('  --output=file.ext            Save output to file instead of stdout');
    console.log('  --toText=true                Output plain text instead of JSON AST (legacy)');
    console.log('  --ocr=true                   Enable OCR for images');
    console.log('  --ocrLanguage=eng            OCR language (default: eng)');
    console.log('  --extractAttachments=true    Extract embedded attachments');
    console.log('  --ignoreNotes=true           Ignore footnotes/endnotes');
    console.log('  --putNotesAtLast=true        Move notes to end of document');
    console.log('  --includeRawContent=true     Include raw content in AST');
    console.log('  --serializeRawContent=true   Serialize raw XML content (default: true)');
    console.log('  --preserveXmlWhitespace=true Preserve whitespace in serialized XML (default: false)');
    console.log('  --includeBreakNodes=false    Include break nodes (DOCX only, default: false)');
    console.log('  --verbose=true               Show full error stack traces');
    console.log('');
    console.log('Examples:');
    console.log('  officeparser document.docx');
    console.log('  officeparser document.docx --format=html --output=doc.html');
    console.log('  officeparser document.docx --format=md');
    console.log('  officeparser report.pdf --ocr=true --format=text');
    console.log('  officeparser data.xlsx --format=csv --output=data.csv');
    console.log('  officeparser complex.docx --serializeRawContent=false --includeRawContent=true');
}

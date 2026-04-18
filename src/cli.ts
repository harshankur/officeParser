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
 *   --toText=true             Output plain text instead of JSON AST
 *   --ocr=true                Enable OCR for images
 *   --ocrLanguage=eng         OCR language (default: eng)
 *   --extractAttachments=true Extract embedded attachments
 *   --ignoreNotes=true        Ignore footnotes/endnotes
 *   --putNotesAtLast=true     Move notes to end of document
 *   --includeRawContent=true  Include raw content in AST
 *   --outputErrorToConsole=true  Log errors to console
 */

import { OfficeParser } from './OfficeParser.js';
import { OfficeParserAST, OfficeParserConfig } from './types.js';

const args = process.argv.slice(2);
let fileArg: string | undefined;
let toText = false;
let verbose = false;
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

        if (cleanKey === 'toText') {
            if (boolValue !== undefined) toText = boolValue;
            else console.warn(`Invalid value for toText: ${value}`);
        } else if (cleanKey === 'verbose') {
            if (boolValue !== undefined) verbose = boolValue;
            else console.warn(`Invalid value for verbose: ${value}`);
        } else {
            // @ts-ignore
            if (boolValue !== undefined) config[cleanKey] = boolValue;
            // @ts-ignore
            else config[cleanKey] = value;
        }
    });

    OfficeParser.parseOffice(fileArg, config)
        .then(async (ast: OfficeParserAST) => {
            if (toText) {
                process.stdout.write(ast.toText() + '\n');
            } else {
                process.stdout.write(JSON.stringify(ast, null, 2) + '\n');
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
    console.log('  --toText=true                Output plain text instead of JSON AST');
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
    console.log('  officeparser document.docx --toText=true');
    console.log('  officeparser report.pdf --ocr=true --extractAttachments=true');
    console.log('  officeparser complex.docx --serializeRawContent=false --includeRawContent=true');
}

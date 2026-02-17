#!/usr/bin/env node
"use strict";
/**
 * officeparser - Universal Office Document Parser
 *
 * A comprehensive Node.js library for parsing Microsoft Office and OpenDocument files
 * into structured Abstract Syntax Trees (AST) with full formatting information.
 *
 * **Supported Formats:**
 * - Microsoft Office: DOCX, XLSX, PPTX (Office Open XML)
 * - OpenDocument: ODT, ODP, ODS (ODF)
 * - Legacy: RTF (Rich Text Format)
 * - Portable: PDF
 *
 * **Key Features:**
 * - Unified AST output across all formats
 * - Rich text formatting (bold, italic, colors, fonts, etc.)
 * - Document structure (headings, lists, tables)
 * - Image extraction with optional OCR
 * - Metadata extraction
 * - TypeScript support with full type definitions
 *
 * **Quick Start:**
 * ```typescript
 * import { OfficeParser } from 'officeparser';
 *
 * const ast = await OfficeParser.parseOffice('document.docx', {
 *   extractAttachments: true,
 *   ocr: true,
 *   includeRawContent: false
 * });
 *
 * console.log(ast.toText()); // Plain text output
 * console.log(ast.content);  // Structured content tree
 * console.log(ast.metadata); // Document metadata
 * ```
 *
 * **Main Exports:**
 * - `OfficeParser` - Main parser class
 * - `OfficeParserConfig` - Configuration interface
 * - `OfficeParserAST` - AST result interface
 * - `OfficeContentNode` - Content tree node interface
 * - All type definitions
 *
 * @packageDocumentation
 * @module officeparser
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.parseOffice = exports.OfficeParser = void 0;
const OfficeParser_1 = require("./OfficeParser");
Object.defineProperty(exports, "OfficeParser", { enumerable: true, get: function () { return OfficeParser_1.OfficeParser; } });
const parseOffice = OfficeParser_1.OfficeParser.parseOffice;
exports.parseOffice = parseOffice;
// Default export for backward compatibility
exports.default = OfficeParser_1.OfficeParser;
// CLI handling - allows running as: node index.js file.docx
if (typeof require !== 'undefined' && typeof module !== 'undefined' && require.main === module) {
    const args = process.argv.slice(2);
    let fileArg;
    let toText = false;
    const configArgs = [];
    function isConfigOption(arg) {
        return arg.startsWith('--') && arg.includes('=');
    }
    args.forEach(arg => {
        if (isConfigOption(arg)) {
            configArgs.push(arg);
        }
        else if (!fileArg) {
            fileArg = arg;
        }
    });
    if (fileArg) {
        const config = {};
        configArgs.forEach(arg => {
            const [key, value] = arg.split('=');
            const cleanKey = key.replace('--', '');
            if (cleanKey === 'toText') {
                if (value.toLowerCase() === 'true')
                    toText = true;
                else if (value.toLowerCase() === 'false')
                    toText = false;
                else
                    console.log(`Invalid value for toText: ${value}`);
            }
            // @ts-ignore
            else if (value.toLowerCase() === 'true')
                config[cleanKey] = true;
            // @ts-ignore
            else if (value.toLowerCase() === 'false')
                config[cleanKey] = false;
            // @ts-ignore
            else
                config[cleanKey] = value;
        });
        OfficeParser_1.OfficeParser.parseOffice(fileArg, config)
            .then((ast) => {
            if (toText)
                console.log(ast.toText());
            else
                console.log(JSON.stringify(ast, null, 2));
        })
            .catch(console.error);
    }
    else {
        console.log("Usage: node officeparser [file] [--option=value]");
    }
}

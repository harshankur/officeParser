/**
 * OfficeTemplate: fill a document template's `{{placeholder}}` tags with data and get back new
 * document(s) - a lightweight mail-merge / document-generation entry point (no extra dependencies; built on the bundled fflate).
 *
 * This is deliberately separate from parsing and generation: it does not build an AST or convert
 * anything, it copies the template and substitutes only the placeholders, so every bit of the
 * original formatting, layout and structure is preserved. Give it one data object to get one
 * document back, or an array to get one document per entry (a batch mail-merge).
 *
 * Today it templates DOCX (`.docx`); the run-aware substitution generalizes to the other OOXML/ODF
 * text formats later. See {@link module:template/docxTemplate} for the substitution details.
 *
 * @example
 * ```typescript
 * import { OfficeTemplate } from 'officeparser';
 * import { writeFileSync } from 'fs';
 *
 * // One document.
 * const bytes = await OfficeTemplate.render('invoice-template.docx', {
 *   data: { name: 'Acme Corp', amount: '$1,250.00', due: '2026-10-01' },
 * });
 * writeFileSync('invoice-acme.docx', bytes);
 *
 * // A batch: one document per row.
 * const docs = await OfficeTemplate.render('invoice-template.docx', {
 *   data: [ { name: 'Acme Corp', amount: '$1,250.00' }, { name: 'Globex', amount: '$980.00' } ],
 * });
 * docs.forEach((d, i) => writeFileSync(`invoice-${i}.docx`, d));
 * ```
 *
 * @module OfficeTemplate
 */

import { decryptIfNeeded } from './crypto/decryptContainer.js';
import { DEFAULT_OFFICE_PARSER_CONFIG, DEFAULT_TEMPLATE_CONFIG } from './defaults.js';
import { compileDocxTemplate, DocxTemplateOptions, openDocx, renderDocxTemplate } from './template/docxTemplate.js';
import { BlobLike, OfficeErrorType, OfficeParserConfig, TemplateConfig, TemplateData, TemplateDataDetails, TemplateFormatter, TemplateSyntaxDetails } from './types.js';
import { PictureSettings } from './template/templateFormat.js';
import { assertNode } from './utils/envUtils.js';
import { getOfficeError } from './utils/errorUtils.js';

/** Accepted template inputs, matching `parseOffice`'s first argument. */
export type TemplateInput = string | Buffer | ArrayBuffer | Uint8Array | BlobLike;

/** Reads the template input into a Node Buffer (path via fs in Node; bytes/blob anywhere). */
async function readInput(input: TemplateInput, errCfg: OfficeParserConfig): Promise<Buffer> {
    if (typeof input === 'string') {
        assertNode('path-parsing', errCfg);
        const fs = await import('fs');
        try { return fs.readFileSync(input); }
        catch { throw getOfficeError(OfficeErrorType.FILE_DOES_NOT_EXIST, errCfg, input); }
    }
    if (input instanceof ArrayBuffer) return Buffer.from(input);
    // A Node Buffer is a Uint8Array, so this branch handles both.
    if (input instanceof Uint8Array) return Buffer.from(input.buffer, input.byteOffset, input.byteLength);
    if (input && typeof (input as BlobLike).arrayBuffer === 'function') return Buffer.from(await (input as BlobLike).arrayBuffer());
    throw getOfficeError(OfficeErrorType.INVALID_INPUT, errCfg);
}

/** The names a filter of the caller's own may have: a word, not one of the built-in filters. */
const FORMATTER_NAME = /^[\p{L}_][\p{L}\p{N}_]*$/u;
const BUILT_IN_FILTER_NAMES = new Set(['format', 'upper', 'lower']);

/** What a picture writes, from the config and the defaults; a setting of the wrong shape is refused, typed. */
function pictureSettings(config: TemplateConfig, errCfg: OfficeParserConfig): PictureSettings {
    const separators = { ...DEFAULT_TEMPLATE_CONFIG.numberSeparators, ...config.numberSeparators };
    const names = { ...DEFAULT_TEMPLATE_CONFIG.dateNames, ...config.dateNames };
    const offset = config.utcOffsetMinutes ?? DEFAULT_TEMPLATE_CONFIG.utcOffsetMinutes;
    const strings = (list: unknown, length: number): boolean => Array.isArray(list) && list.length === length && list.every(x => typeof x === 'string');
    const refuse = (detail: string): never => { throw getOfficeError(OfficeErrorType.IMPROPER_ARGUMENTS, errCfg, detail); };
    if (typeof separators.decimal !== 'string' || typeof separators.group !== 'string') refuse('numberSeparators.decimal and .group are text');
    if (!strings(names.months, 12) || !strings(names.monthsShort, 12)) refuse('dateNames.months and .monthsShort are twelve names each');
    if (!strings(names.days, 7) || !strings(names.daysShort, 7)) refuse('dateNames.days and .daysShort are seven names each, Sunday first');
    if (typeof names.am !== 'string' || typeof names.pm !== 'string') refuse('dateNames.am and .pm are text');
    if (!Number.isInteger(offset) || Math.abs(offset) > MAX_UTC_OFFSET_MINUTES) refuse(`utcOffsetMinutes is a whole number of minutes, at most ${MAX_UTC_OFFSET_MINUTES} either way`);
    return {
        decimal: separators.decimal, group: separators.group, months: names.months, monthsShort: names.monthsShort,
        days: names.days, daysShort: names.daysShort, am: names.am, pm: names.pm, utcOffsetMinutes: offset,
    };
}

/** The most a UTC offset can be: a day, either way. */
const MAX_UTC_OFFSET_MINUTES = 24 * 60;

/** The caller's own filters: functions, named as words, none named as a built-in filter. */
function templateFormatters(config: TemplateConfig, errCfg: OfficeParserConfig): Record<string, TemplateFormatter> {
    const formatters = config.formatters ?? {};
    if (typeof formatters !== 'object' || formatters === null) throw getOfficeError(OfficeErrorType.IMPROPER_ARGUMENTS, errCfg, 'formatters is an object of functions by name');
    const own: Record<string, TemplateFormatter> = Object.create(null);
    for (const [name, run] of Object.entries(formatters)) {
        if (typeof run !== 'function') throw getOfficeError(OfficeErrorType.IMPROPER_ARGUMENTS, errCfg, `formatters.${name} is not a function`);
        if (!FORMATTER_NAME.test(name) || BUILT_IN_FILTER_NAMES.has(name)) throw getOfficeError(OfficeErrorType.IMPROPER_ARGUMENTS, errCfg, `formatters.${name} is not a word, or is the name of a built-in filter (format, upper, lower)`);
        own[name] = run;
    }
    return own;
}

/** A single data entry must be a plain object (not null/array/primitive). */
function isDataObject(d: unknown): d is TemplateData {
    return typeof d === 'object' && d !== null && !Array.isArray(d);
}

/**
 * Template renderer. See {@link module:OfficeTemplate} for an overview and {@link TemplateConfig} for
 * the options.
 */
export class OfficeTemplate {
    /** Renders a single document from a template and one data map. */
    static render(template: TemplateInput, config: TemplateConfig & { data: TemplateData }): Promise<Uint8Array>;
    /** Renders one document per entry from a template and an array of data maps (a batch mail-merge). */
    static render(template: TemplateInput, config: TemplateConfig & { data: TemplateData[] }): Promise<Uint8Array[]>;
    static render(template: TemplateInput, config: TemplateConfig): Promise<Uint8Array | Uint8Array[]>;
    static async render(template: TemplateInput, config: TemplateConfig): Promise<Uint8Array | Uint8Array[]> {
        // A config object (even without onWarning) suppresses errorUtils' console fallback, so a thrown
        // rejection is not also printed to stderr; a caller-supplied onWarning is still honored.
        const errCfg: OfficeParserConfig = { onWarning: config?.onWarning };
        if (!config || config.data == null) throw getOfficeError(OfficeErrorType.IMPROPER_ARGUMENTS, errCfg, 'data is required: an object, or an array of objects');
        // Validate the data up front: one object, or an array of objects. Rejects a null/primitive
        // batch entry with a typed error instead of a raw TypeError mid-render.
        const rows = Array.isArray(config.data) ? config.data : [config.data];
        if (!rows.every(isDataObject)) throw getOfficeError(OfficeErrorType.IMPROPER_ARGUMENTS, errCfg, 'data is an object, or an array of objects');

        let bytes = await readInput(template, errCfg);

        // Decrypt an encrypted template up front (shared with the parser), so rendering sees plaintext.
        // Threads through `password` and `onPassword` with the same capped-retry semantics as parsing.
        bytes = await decryptIfNeeded(bytes, {
            onWarning: config.onWarning,
            password: config.password,
            onPassword: config.onPassword,
            maxPasswordAttempts: config.maxPasswordAttempts,
            decompressionLimits: config.decompressionLimits,
        });

        const entries = await openDocx(bytes, config.decompressionLimits, errCfg);
        if (!entries) {
            const hint = config.fileType && config.fileType !== 'docx' ? `'${config.fileType}' is not supported` : 'the input is not a .docx';
            throw getOfficeError(OfficeErrorType.TEMPLATE_UNSUPPORTED_FORMAT, errCfg, hint);
        }

        const options: DocxTemplateOptions = {
            start: config.delimiters?.start || DEFAULT_TEMPLATE_CONFIG.delimiters.start,
            end: config.delimiters?.end || DEFAULT_TEMPLATE_CONFIG.delimiters.end,
            onMissing: config.onMissing || DEFAULT_TEMPLATE_CONFIG.onMissing,
            // A fixed local-time instant: fflate stamps zip entries from local-time fields, so building the
            // date from local fields makes the DOS timestamp identical in every timezone (deterministic
            // output cross-machine) and safely inside fflate's 1980-2099 range everywhere.
            mtime: new Date(2001, 0, 1, 0, 0, 0),
            onFieldMissing: (key: string): never => { throw getOfficeError(OfficeErrorType.TEMPLATE_FIELD_MISSING, errCfg, key); },
            maxRepeatedContent: config.decompressionLimits?.maxRepeatedContent ?? DEFAULT_TEMPLATE_CONFIG.maxRepeatedContent,
            onRepeatLimit: (): never => { throw getOfficeError(OfficeErrorType.OUTPUT_TOO_LARGE, errCfg, { repeatedLimit: options.maxRepeatedContent }); },
            onSyntax: (details: TemplateSyntaxDetails): never => { throw getOfficeError(OfficeErrorType.TEMPLATE_SYNTAX, errCfg, details); },
            onDataInvalid: (details: TemplateDataDetails): never => { throw getOfficeError(OfficeErrorType.TEMPLATE_DATA_INVALID, errCfg, details); },
            onNestingLimit: (): never => { throw getOfficeError(OfficeErrorType.MAX_NESTING_DEPTH_EXCEEDED, errCfg); },
            onOutputLimit: (): never => { throw getOfficeError(OfficeErrorType.OUTPUT_TOO_LARGE, errCfg); },
            pictures: pictureSettings(config, errCfg),
            formatters: templateFormatters(config, errCfg),
            maxXmlElements: config.decompressionLimits?.maxXmlElements ?? DEFAULT_OFFICE_PARSER_CONFIG.decompressionLimits.maxXmlElements,
            onElementLimit: (): never => { throw getOfficeError(OfficeErrorType.XML_ELEMENT_LIMIT_EXCEEDED, errCfg, options.maxXmlElements); },
        };

        // Compiled once: the tags are read, their blocks paired and scoped, and every tag error raised
        // before any document is rendered; then each data entry renders from the same compiled template.
        const compiled = compileDocxTemplate(entries, options);
        const renderOne = (data: TemplateData): Uint8Array => renderDocxTemplate(compiled, data, options);

        return Array.isArray(config.data) ? rows.map(renderOne) : renderOne(config.data);
    }
}

/** Standalone alias for {@link OfficeTemplate.render}. */
export const renderTemplate = OfficeTemplate.render;

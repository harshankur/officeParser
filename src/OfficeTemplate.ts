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
import { BlobLike, OfficeErrorType, OfficeParserConfig, OfficeWarningType, TemplateConfig, TemplateData, TemplateDataDetails, TemplateFormatter, TemplateSyntaxDetails } from './types.js';
import { PictureSettings } from './template/templateFormat.js';
import { assertNode } from './utils/envUtils.js';
import { getOfficeError, logWarning } from './utils/errorUtils.js';
import { decompressionLimitsOf, maxPasswordAttemptsOf, PROTOTYPE_POLLUTION_KEYS } from './utils/configUtils.js';

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

/** Every option render takes (the compiler holds this to TemplateConfig's own keys). */
const TEMPLATE_OPTIONS: Record<keyof TemplateConfig, true> = {
    data: true, delimiters: true, onMissing: true, numberSeparators: true, dateNames: true, utcOffsetMinutes: true, formatters: true,
    password: true, onPassword: true, maxPasswordAttempts: true, fileType: true, decompressionLimits: true, onWarning: true,
};

/**
 * The options of a render's config that it does not take, as dotted paths: at the top, and inside
 * `delimiters` and `decompressionLimits` (`numberSeparators` and `dateNames` are checked with their
 * values). A misspelt one left its default in place without a word; each is now reported.
 */
function unknownTemplateOptions(config: TemplateConfig): string[] {
    const unknown: string[] = [];
    const check = (container: unknown, path: string, known: object): void => {
        if (!container || typeof container !== 'object' || Array.isArray(container)) return;
        for (const key of Object.keys(container)) {
            if (!PROTOTYPE_POLLUTION_KEYS.has(key) && !Object.prototype.hasOwnProperty.call(known, key)) unknown.push(path ? `${path}.${key}` : key);
        }
    };
    check(config, '', TEMPLATE_OPTIONS);
    check(config.delimiters, 'delimiters', DEFAULT_TEMPLATE_CONFIG.delimiters);
    check(config.decompressionLimits, 'decompressionLimits', DEFAULT_OFFICE_PARSER_CONFIG.decompressionLimits);
    return unknown;
}

/**
 * What a picture writes, from the config and the defaults; a setting of the wrong shape is refused, typed.
 * A name `numberSeparators` or `dateNames` does not have is added to `unknown`, to be reported.
 */
function pictureSettings(config: TemplateConfig, errCfg: OfficeParserConfig, unknown: string[]): PictureSettings {
    const refuse = (detail: string): never => { throw getOfficeError(OfficeErrorType.IMPROPER_ARGUMENTS, errCfg, detail); };
    const own = <T extends object>(option: 'numberSeparators' | 'dateNames', defaults: T, shape: string): Partial<T> => {
        const given: unknown = config[option];
        if (given === undefined || given === null) return {};
        if (typeof given !== 'object' || Array.isArray(given)) refuse(`${option} is an object: ${shape}`);
        const taken: Record<string, unknown> = {};
        for (const key of Object.keys(given as object)) {
            if (Object.prototype.hasOwnProperty.call(defaults, key)) taken[key] = (given as Record<string, unknown>)[key];
            else unknown.push(`${option}.${key}`);
        }
        return taken as Partial<T>;
    };
    const separators = { ...DEFAULT_TEMPLATE_CONFIG.numberSeparators, ...own('numberSeparators', DEFAULT_TEMPLATE_CONFIG.numberSeparators, 'decimal and group') };
    const names = { ...DEFAULT_TEMPLATE_CONFIG.dateNames, ...own('dateNames', DEFAULT_TEMPLATE_CONFIG.dateNames, 'months, monthsShort, days, daysShort, am and pm') };
    const offset = config.utcOffsetMinutes ?? DEFAULT_TEMPLATE_CONFIG.utcOffsetMinutes;
    const strings = (list: unknown, length: number): boolean => Array.isArray(list) && list.length === length && list.every(x => typeof x === 'string');
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

/**
 * `onMissing` as given when it is one the renderer knows; another value is reported (INVALID_CONFIG_VALUE)
 * and the default used. One such as `'skip'` emptied fields but kept the blocks they decide as typed.
 */
function onMissingOf(config: TemplateConfig, errCfg: OfficeParserConfig): DocxTemplateOptions['onMissing'] {
    const value: unknown = config.onMissing;
    if (value === 'keep' || value === 'empty' || value === 'error') return value;
    if (value !== undefined && value !== null && value !== '') {
        logWarning(OfficeWarningType.INVALID_CONFIG_VALUE, errCfg, { option: 'onMissing', value, expected: "'keep', 'empty' or 'error'", fallback: DEFAULT_TEMPLATE_CONFIG.onMissing });
    }
    return DEFAULT_TEMPLATE_CONFIG.onMissing;
}

/**
 * The tag delimiters, each as given or its default when left out or empty. Delimiters that are not an
 * object, or one that is not text, are refused, typed: a number crashed the render untyped, and text
 * or a list in place of the object was passed over without a word.
 */
function delimitersOf(config: TemplateConfig, errCfg: OfficeParserConfig): { start: string; end: string } {
    const given: unknown = config.delimiters;
    const defaults = DEFAULT_TEMPLATE_CONFIG.delimiters;
    if (given === undefined || given === null) return { ...defaults };
    const refuse = (): never => { throw getOfficeError(OfficeErrorType.IMPROPER_ARGUMENTS, errCfg, 'delimiters is an object: start and end, each text'); };
    if (typeof given !== 'object' || Array.isArray(given)) refuse();
    const { start, end } = given as { start?: unknown; end?: unknown };
    if ((start !== undefined && start !== null && typeof start !== 'string') || (end !== undefined && end !== null && typeof end !== 'string')) refuse();
    return { start: (start as string | undefined) || defaults.start, end: (end as string | undefined) || defaults.end };
}

/** The most a UTC offset can be: a day, either way. */
const MAX_UTC_OFFSET_MINUTES = 24 * 60;

/** The caller's own filters: functions, named as words, none named as a built-in filter. */
function templateFormatters(config: TemplateConfig, errCfg: OfficeParserConfig): Record<string, TemplateFormatter> {
    const formatters = config.formatters ?? {};
    if (typeof formatters !== 'object' || formatters === null || Array.isArray(formatters)) throw getOfficeError(OfficeErrorType.IMPROPER_ARGUMENTS, errCfg, 'formatters is an object of functions by name');
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

        // The options themselves, before the template is read: a setting of the wrong shape is refused,
        // and one render does not take is reported, as a parse reports its own.
        const unknownOptions = unknownTemplateOptions(config);
        const pictures = pictureSettings(config, errCfg, unknownOptions);
        const delimiters = delimitersOf(config, errCfg);
        if (unknownOptions.length) logWarning(OfficeWarningType.UNRECOGNIZED_CONFIG_OPTION, errCfg, { keys: unknownOptions });

        // The limits a render is held to, each a number of 0 or more (another value is reported, and the default used).
        const limits = decompressionLimitsOf(config.decompressionLimits, errCfg);
        let bytes = await readInput(template, errCfg);

        // Decrypt an encrypted template up front (shared with the parser), so rendering sees plaintext.
        // Threads through `password` and `onPassword` with the same capped-retry semantics as parsing.
        bytes = await decryptIfNeeded(bytes, {
            onWarning: config.onWarning,
            password: config.password,
            onPassword: config.onPassword,
            maxPasswordAttempts: maxPasswordAttemptsOf(config),
            decompressionLimits: limits,
        });

        const entries = await openDocx(bytes, limits, errCfg);
        if (!entries) {
            const hint = config.fileType && config.fileType !== 'docx' ? `'${config.fileType}' is not supported` : 'the input is not a .docx';
            throw getOfficeError(OfficeErrorType.TEMPLATE_UNSUPPORTED_FORMAT, errCfg, hint);
        }

        const options: DocxTemplateOptions = {
            start: delimiters.start,
            end: delimiters.end,
            onMissing: onMissingOf(config, errCfg),
            // A fixed local-time instant: fflate stamps zip entries from local-time fields, so building the
            // date from local fields makes the DOS timestamp identical in every timezone (deterministic
            // output cross-machine) and safely inside fflate's 1980-2099 range everywhere.
            mtime: new Date(2001, 0, 1, 0, 0, 0),
            onFieldMissing: (key: string): never => { throw getOfficeError(OfficeErrorType.TEMPLATE_FIELD_MISSING, errCfg, key); },
            maxRepeatedContent: limits.maxRepeatedContent,
            onRepeatLimit: (): never => { throw getOfficeError(OfficeErrorType.OUTPUT_TOO_LARGE, errCfg, { repeatedLimit: options.maxRepeatedContent }); },
            onSyntax: (details: TemplateSyntaxDetails): never => { throw getOfficeError(OfficeErrorType.TEMPLATE_SYNTAX, errCfg, details); },
            onDataInvalid: (details: TemplateDataDetails): never => { throw getOfficeError(OfficeErrorType.TEMPLATE_DATA_INVALID, errCfg, details); },
            onNestingLimit: (): never => { throw getOfficeError(OfficeErrorType.MAX_NESTING_DEPTH_EXCEEDED, errCfg); },
            onOutputLimit: (): never => { throw getOfficeError(OfficeErrorType.OUTPUT_TOO_LARGE, errCfg); },
            pictures,
            formatters: templateFormatters(config, errCfg),
            maxXmlElements: limits.maxXmlElements,
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

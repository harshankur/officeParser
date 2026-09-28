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
import { openDocx, renderDocxTemplate } from './template/docxTemplate.js';
import { BlobLike, OfficeErrorType, OfficeParserConfig, TemplateConfig, TemplateData } from './types.js';
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
        if (!config || config.data == null) throw getOfficeError(OfficeErrorType.IMPROPER_ARGUMENTS, errCfg);
        // Validate the data up front: one object, or an array of objects. Rejects a null/primitive
        // batch entry with a typed error instead of a raw TypeError mid-render.
        const rows = Array.isArray(config.data) ? config.data : [config.data];
        if (!rows.every(isDataObject)) throw getOfficeError(OfficeErrorType.IMPROPER_ARGUMENTS, errCfg);

        let bytes = await readInput(template, errCfg);

        // Decrypt an encrypted template up front (shared with the parser), so rendering sees plaintext.
        // Threads through `password` and `onPassword` with the same capped-retry semantics as parsing.
        bytes = await decryptIfNeeded(bytes, {
            onWarning: config.onWarning,
            password: config.password,
            onPassword: config.onPassword,
            decompressionLimits: config.decompressionLimits,
        });

        const entries = await openDocx(bytes, config.decompressionLimits, errCfg);
        if (!entries) {
            const hint = config.fileType && config.fileType !== 'docx' ? `'${config.fileType}' is not supported` : 'the input is not a .docx';
            throw getOfficeError(OfficeErrorType.TEMPLATE_UNSUPPORTED_FORMAT, errCfg, hint);
        }

        const start = config.delimiters?.start || '{{';
        const end = config.delimiters?.end || '}}';
        const onMissing = config.onMissing || 'keep';
        // A fixed local-time instant: fflate stamps zip entries from local-time fields, so building the
        // date from local fields makes the DOS timestamp identical in every timezone (deterministic
        // output cross-machine) and safely inside fflate's 1980-2099 range everywhere.
        const mtime = new Date(2001, 0, 1, 0, 0, 0);
        const onFieldMissing = (key: string): never => { throw getOfficeError(OfficeErrorType.TEMPLATE_FIELD_MISSING, errCfg, key); };
        const maxRepeatedContent = config.decompressionLimits?.maxRepeatedContent ?? 16 * 1024 * 1024;
        const onRepeatLimit = (): never => { throw getOfficeError(OfficeErrorType.OUTPUT_TOO_LARGE, errCfg, { repeatedLimit: maxRepeatedContent }); };

        const renderOne = (data: TemplateData): Uint8Array =>
            renderDocxTemplate(entries, data, { start, end, onMissing, mtime, onFieldMissing, maxRepeatedContent, onRepeatLimit });

        return Array.isArray(config.data) ? rows.map(renderOne) : renderOne(config.data);
    }
}

/** Standalone alias for {@link OfficeTemplate.render}. */
export const renderTemplate = OfficeTemplate.render;

import { DEFAULT_DOCUMENT_STRUCTURE_CHUNKING_CONFIG, DEFAULT_FIXED_SIZE_CHUNKING_CONFIG, DEFAULT_GENERATOR_CONFIG, DEFAULT_OFFICE_PARSER_CONFIG, DEFAULT_SEMANTIC_CHUNKING_CONFIG } from '../defaults.js';
import { FullGeneratorConfig, FullOfficeParserConfig, GeneratorConfig, OfficeIssue, OfficeParserConfig, OfficeWarningType } from '../types.js';
import { logWarning } from './errorUtils.js';
import { PAPER_FORMAT_NAMES } from './officeGenUtils.js';
import { lookupTable } from './lookupUtils.js';

/**
 * Keys that must never be copied from a caller-supplied config onto one of our objects.
 *
 * A config that arrived via `JSON.parse` can carry `__proto__` as a genuine **own enumerable**
 * property (an object *literal* cannot - there `__proto__` invokes the setter at parse time),
 * which is exactly the shape of a host application accepting a JSON config blob. Copying that
 * key reaches `Object.prototype` and corrupts every object in the process.
 *
 * `constructor` and `prototype` are included because they are the other two names that reach a
 * prototype through an ordinary property write.
 */
export const PROTOTYPE_POLLUTION_KEYS = new Set(['__proto__', 'constructor', 'prototype']);

/**
 * Every recognized top-level parser-config key. `DEFAULT_OFFICE_PARSER_CONFIG` is `DeepRequired`, so
 * its own keys are exactly the full option surface; a caller key absent from this set is genuinely
 * unrecognized (a typo, or a key that no longer exists), never a valid-but-omitted option.
 */
export const RECOGNIZED_PARSER_KEYS = new Set(Object.keys(DEFAULT_OFFICE_PARSER_CONFIG));

/**
 * Options renamed or removed in a major release, mapped old key -> what to use instead, so the
 * "unrecognized option" warning can point an upgrading caller straight at the replacement instead
 * of only saying the key did nothing.
 *
 * A value that starts with `(` is prose rather than a config path: the key has no one-to-one
 * replacement, so the warning renders the sentence as-is (see `UNRECOGNIZED_CONFIG_OPTION`).
 * Nested keys are spelled with their full dotted path, matching how the check below reports them.
 */
const RENAMED_PARSER_KEYS: Record<string, string> = lookupTable({
    ocrLanguage: 'ocrConfig.language',
    outputErrorToConsole: 'onWarning',
    putNotesAtLast: '(removed; notes are attached to the node they belong to, as node.notes)',
    'ocrConfig.autoTerminateTimeout': 'ocrConfig.timeout.autoTerminate',
});

/**
 * Every recognized key of each nested parser-config container, keyed by the container's dotted path
 * and derived from the defaults exactly as `RECOGNIZED_PARSER_KEYS` is.
 *
 * Nested keys need a check of their own because the merge below only ever spreads a caller's
 * sub-object over the matching defaults. A key that no longer exists there (a v7
 * `ocrConfig.autoTerminateTimeout`, say) is copied in and then simply never read, so the caller's
 * setting is silently replaced by the default it meant to override.
 */
const RECOGNIZED_NESTED_PARSER_KEYS: Record<string, Set<string>> = lookupTable({
    ocrConfig: new Set(Object.keys(DEFAULT_OFFICE_PARSER_CONFIG.ocrConfig)),
    'ocrConfig.timeout': new Set(Object.keys(DEFAULT_OFFICE_PARSER_CONFIG.ocrConfig.timeout)),
    pdfParserConfig: new Set(Object.keys(DEFAULT_OFFICE_PARSER_CONFIG.pdfParserConfig)),
    htmlParserConfig: new Set(Object.keys(DEFAULT_OFFICE_PARSER_CONFIG.htmlParserConfig)),
    texParserConfig: new Set(Object.keys(DEFAULT_OFFICE_PARSER_CONFIG.texParserConfig)),
    decompressionLimits: new Set(Object.keys(DEFAULT_OFFICE_PARSER_CONFIG.decompressionLimits)),
});

/**
 * Appends every key of `container` that the container at `path` does not recognize to `out`, as a
 * dotted path. Recurses into a sub-container that has recognized keys of its own (`ocrConfig.timeout`).
 */
function collectUnknownNestedKeys(container: unknown, path: string, out: string[]): void {
    if (!container || typeof container !== 'object') return;
    const recognized = RECOGNIZED_NESTED_PARSER_KEYS[path];
    if (!recognized) return;
    for (const key of Object.keys(container)) {
        if (PROTOTYPE_POLLUTION_KEYS.has(key)) continue;
        const fullPath = `${path}.${key}`;
        if (RECOGNIZED_NESTED_PARSER_KEYS[fullPath]) {
            collectUnknownNestedKeys((container as any)[key], fullPath, out);
            continue;
        }
        if (!recognized.has(key)) out.push(fullPath);
    }
}

/**
 * Returns a copy of `source` with prototype-reaching keys removed.
 *
 * Needed before `Object.assign`, which does **not** pollute `Object.prototype` (it writes via
 * `[[Set]]`, so `__proto__` invokes the inherited setter rather than creating an own property) -
 * but that setter is not inert: it **replaces the target's prototype**, so the returned config
 * silently inherits attacker-chosen properties for every field the defaults don't set as an own
 * property. Narrower than global pollution, still wrong. Do not "simplify" this away on the
 * grounds that `Object.assign` is safe; it is safe only against the *global* variant.
 */
function withoutPrototypeKeys<T extends object>(source: T): T {
    const safe: any = {};
    for (const key of Object.keys(source)) {
        if (PROTOTYPE_POLLUTION_KEYS.has(key)) continue;
        safe[key] = (source as any)[key];
    }
    return safe;
}

/**
 * Deep clones an object, specifically handling arrays and plain objects.
 */
function deepClone<T>(obj: T): T {
    if (obj === null || typeof obj !== 'object') {
        return obj;
    }

    if (Array.isArray(obj)) {
        return (obj as any).map((item: any) => deepClone(item));
    }

    const cloned: any = {};
    for (const key in obj) {
        if (Object.prototype.hasOwnProperty.call(obj, key)) {
            cloned[key] = deepClone((obj as any)[key]);
        }
    }
    return cloned;
}

/**
 * Checks if a configuration object is a FullGeneratorConfig.
 */
export function isFullGeneratorConfig(config: any): config is FullGeneratorConfig {
    return !!(config &&
        typeof config === 'object' &&
        'textConfig' in config &&
        'htmlConfig' in config &&
        'pdfConfig' in config &&
        'csvConfig' in config &&
        'docxConfig' in config &&
        'odtConfig' in config &&
        'texConfig' in config &&
        'onNode' in config);
}

/**
 * Checks if a configuration object is a FullOfficeParserConfig.
 */
export function isFullParserConfig(config: any): config is FullOfficeParserConfig {
    return !!(config &&
        typeof config === 'object' &&
        'ocrConfig' in config &&
        typeof config.ocrConfig === 'object' &&
        'language' in config.ocrConfig &&
        'workerPath' in config.ocrConfig);
}

/**
 * Resolves a full parser configuration by merging defaults and user-provided overrides.
 *
 * The returned object always belongs solely to the caller of this function. That matters
 * because a parse installs per-call state on the config it is handed, such as the collector
 * that gathers warnings for one document's `ast.warnings`. Returning the caller's own object
 * would attach that state to an object they may reuse, so a second parse would append its
 * warnings to the first document's already-returned AST, and each parse would retain the
 * previous one's state for as long as the config lived.
 *
 * Only the configuration containers are copied. Callbacks and `abortSignal` keep their
 * identity, since a copy of an `AbortSignal` would no longer be tied to its controller.
 *
 * @param userConfig - Optional configuration provided by the user
 * @returns A fully populated configuration object, owned by the caller
 */
export function resolveParserConfig(
    userConfig?: OfficeParserConfig | FullOfficeParserConfig
): FullOfficeParserConfig {
    if (isFullParserConfig(userConfig)) {
        const resolved: FullOfficeParserConfig = { ...userConfig };
        resolved.ocrConfig = { ...userConfig.ocrConfig };
        if (userConfig.ocrConfig?.timeout) {
            resolved.ocrConfig.timeout = { ...userConfig.ocrConfig.timeout };
        }
        resolved.decompressionLimits = {
            ...(userConfig.decompressionLimits ?? DEFAULT_OFFICE_PARSER_CONFIG.decompressionLimits)
        };
        if (userConfig.htmlParserConfig) {
            resolved.htmlParserConfig = { ...userConfig.htmlParserConfig };
        }
        if (userConfig.pdfParserConfig) {
            resolved.pdfParserConfig = { ...userConfig.pdfParserConfig };
        }
        if (userConfig.texParserConfig) {
            resolved.texParserConfig = { ...userConfig.texParserConfig };
        }
        return resolved;
    }

    // 1. Start with full defaults (deep cloned)
    const config: FullOfficeParserConfig = deepClone(DEFAULT_OFFICE_PARSER_CONFIG);

    if (!userConfig) {
        return config;
    }

    // 2. Merge user config
    // We handle ocrConfig, decompressionLimits and the per-format parser configs specially to avoid
    // shallow-overwriting the whole nested objects
    const { ocrConfig, decompressionLimits, htmlParserConfig, pdfParserConfig, texParserConfig, ...rest } = userConfig;
    Object.assign(config, withoutPrototypeKeys(rest));

    // Flag any option the caller passed that this version does not recognize: a typo, or a key renamed
    // in a major release. Copying it onto `config` above is harmless, but leaving it silently unused
    // means the caller's intent never takes effect with no signal at all. One warning per resolve names
    // exactly which keys did nothing (and, for a known rename, the replacement to use instead).
    //
    // The nested containers are checked too, since a stale key inside one of them is the more
    // dangerous case: the merges below quietly drop it and the caller keeps this version's default.
    const unknownKeys = Object.keys(rest).filter(k => !RECOGNIZED_PARSER_KEYS.has(k) && !PROTOTYPE_POLLUTION_KEYS.has(k));
    collectUnknownNestedKeys(ocrConfig, 'ocrConfig', unknownKeys);
    collectUnknownNestedKeys(pdfParserConfig, 'pdfParserConfig', unknownKeys);
    collectUnknownNestedKeys(htmlParserConfig, 'htmlParserConfig', unknownKeys);
    collectUnknownNestedKeys(texParserConfig, 'texParserConfig', unknownKeys);
    collectUnknownNestedKeys(decompressionLimits, 'decompressionLimits', unknownKeys);
    if (unknownKeys.length) {
        logWarning(OfficeWarningType.UNRECOGNIZED_CONFIG_OPTION, config, { keys: unknownKeys, renames: RENAMED_PARSER_KEYS });
    }

    if (decompressionLimits) {
        config.decompressionLimits = {
            ...config.decompressionLimits,
            ...decompressionLimits,
        };
    }

    if (htmlParserConfig) {
        config.htmlParserConfig = {
            ...config.htmlParserConfig,
            ...htmlParserConfig,
        };
    }

    if (pdfParserConfig) {
        config.pdfParserConfig = {
            ...config.pdfParserConfig,
            ...pdfParserConfig,
        };
    }

    if (texParserConfig) {
        config.texParserConfig = {
            ...config.texParserConfig,
            ...texParserConfig,
        };
        const today = config.texParserConfig.today;
        if (typeof today !== 'string') {
            logWarning(OfficeWarningType.INVALID_CONFIG_VALUE, config, { option: 'texParserConfig.today', value: today, expected: 'a string', fallback: '' });
            config.texParserConfig.today = '';
        }
    }

    if (ocrConfig) {
        const { timeout, ...ocrRest } = ocrConfig;
        config.ocrConfig = {
            ...config.ocrConfig,
            ...ocrRest,
            timeout: {
                autoTerminate: timeout?.autoTerminate !== undefined ? timeout.autoTerminate : config.ocrConfig.timeout.autoTerminate,
                workerLoad: timeout?.workerLoad !== undefined ? timeout.workerLoad : config.ocrConfig.timeout.workerLoad,
                recognition: timeout?.recognition !== undefined ? timeout.recognition : config.ocrConfig.timeout.recognition,
            }
        };
    }

    // 3. Propagate the top-level abortSignal to ocrConfig so the OCR subsystem is aware of it
    if (config.abortSignal) {
        config.ocrConfig.abortSignal = config.abortSignal;
    }

    return config;
}

/** The per-destination and metadata sub-objects a generator config groups its settings into. */
const GENERATOR_CONFIG_CONTAINERS = [
    'metadataOverrides', 'htmlConfig', 'mdConfig', 'pdfConfig',
    'csvConfig', 'textConfig', 'rtfConfig', 'docxConfig', 'odtConfig', 'texConfig', 'chunksConfig',
] as const;

/**
 * Copies a generator config's containers so writes during generation cannot reach the caller.
 *
 * One level is enough: the containers are what generation writes to. Everything else is copied
 * by reference on purpose, since callbacks, `styleMap` and `abortSignal` are values whose
 * identity matters, and a duplicated `AbortSignal` would no longer be tied to its controller.
 *
 * @param source - The caller's configuration
 * @returns An equivalent configuration owned by us
 */
function copyGeneratorConfigContainers(source: FullGeneratorConfig): FullGeneratorConfig {
    const copy: any = { ...source };
    for (const key of GENERATOR_CONFIG_CONTAINERS) {
        const container = (source as any)[key];
        if (container && typeof container === 'object') copy[key] = { ...container };
    }
    return copy;
}

/**
 * Every recognized top-level generator-config key and, per per-format container, every key that
 * container recognizes: the generator-side counterpart of RECOGNIZED_PARSER_KEYS, derived from the
 * defaults the same way. Chunking takes the keys of all three strategies (plus `embeddingFunction`,
 * which has no default); metadata overrides list their named fields, since their default is empty.
 */
export const RECOGNIZED_GENERATOR_KEYS = new Set(Object.keys(DEFAULT_GENERATOR_CONFIG));
const RECOGNIZED_NESTED_GENERATOR_KEYS: Record<string, Set<string>> = lookupTable({
    ...Object.fromEntries(GENERATOR_CONFIG_CONTAINERS.map(key => [key, new Set(Object.keys((DEFAULT_GENERATOR_CONFIG as any)[key] ?? {}))])),
    chunksConfig: new Set([
        ...Object.keys(DEFAULT_FIXED_SIZE_CHUNKING_CONFIG), ...Object.keys(DEFAULT_DOCUMENT_STRUCTURE_CHUNKING_CONFIG),
        ...Object.keys(DEFAULT_SEMANTIC_CHUNKING_CONFIG), 'embeddingFunction',
    ]),
    metadataOverrides: new Set(['title', 'author', 'description', 'subject', 'keywords', 'lastModifiedBy', 'created', 'modified', 'language', 'custom']),
});

/**
 * The top-level keys and per-format container keys of a caller's generator config that this
 * version does not recognize, as dotted paths (as the parser side reports its own).
 */
function unknownGeneratorKeys(userConfig: Record<string, any>): string[] {
    const unknown: string[] = [];
    for (const key of Object.keys(userConfig)) {
        if (PROTOTYPE_POLLUTION_KEYS.has(key)) continue;
        if (!RECOGNIZED_GENERATOR_KEYS.has(key)) { unknown.push(key); continue; }
        const recognized = RECOGNIZED_NESTED_GENERATOR_KEYS[key];
        const container = userConfig[key];
        if (!recognized || !container || typeof container !== 'object' || Array.isArray(container)) continue;
        for (const inner of Object.keys(container)) {
            if (!PROTOTYPE_POLLUTION_KEYS.has(inner) && !recognized.has(inner)) unknown.push(`${key}.${inner}`);
        }
    }
    return unknown;
}

/**
 * Resolves a full, destination-specific configuration by merging defaults,
 * AST-level settings, and user-provided overrides.
 *
 * As with {@link resolveParserConfig}, the returned object belongs solely to the caller of this
 * function, so that per-run normalization cannot edit a config the caller still holds.
 *
 * Problems with the configuration itself (an unrecognized key, a value an option does not accept)
 * are reported once the config is resolved: to its `onWarning`, and to `onIssue` when given, which
 * a generator uses to list them among its result's messages.
 *
 * @param destination - The target format
 * @param userConfig - Optional configuration provided by the user
 * @param astConfig - Optional configuration from the source AST (for inheritance)
 * @param onIssue - Optional collector for the configuration problems found while resolving
 * @returns A fully populated configuration object, owned by the caller
 */
export function resolveGeneratorConfig<D extends string>(
    destination: D,
    astConfig?: OfficeParserConfig,
    userConfig?: GeneratorConfig<D> | FullGeneratorConfig,
    onIssue?: (issue: OfficeIssue) => void
): FullGeneratorConfig {
    // Configuration problems go to the resolved config's own handler (set once the merge below is
    // done) and to the optional collector.
    let resolvedConfig: FullGeneratorConfig | undefined;
    const reporter = { onWarning: (issue: OfficeIssue) => { onIssue?.(issue); resolvedConfig?.onWarning?.(issue); } } as any;
    // Already complete, so nothing to merge. Still copied rather than handed straight back, for
    // the same reason as resolveParserConfig: generation writes to the config it is given. The
    // width check below normalizes an invalid `containerWidth` to 'auto', and doing that to the
    // caller's own object both edits a value they still hold and silences the warning on every
    // later run, so the same config would report a problem once and then appear clean.
    if (isFullGeneratorConfig(userConfig) && !astConfig) {
        const resolved = copyGeneratorConfigContainers(userConfig);
        resolvedConfig = resolved;
        validateHtmlConfigWidth(resolved.htmlConfig, reporter);
        validateGeneratorChoices(resolved, reporter);
        return resolved;
    }

    // 1. Start with full defaults (deep cloned to avoid reference sharing)
    const config: FullGeneratorConfig = deepClone(DEFAULT_GENERATOR_CONFIG);
    let unknownKeys: string[] = [];

    // 2. Merge common properties and sub-configs
    if (userConfig) {
        // As on the parser side: an option this version does not recognize (a typo such as
        // `texConfig.bundel`, or a parser option passed to a generator) would otherwise do nothing
        // with no signal at all.
        unknownKeys = unknownGeneratorKeys(userConfig as Record<string, any>);
        // Extract sub-configs to avoid shallow-overwriting the whole sub-config objects
        const { htmlConfig, mdConfig, pdfConfig, csvConfig, textConfig, rtfConfig, docxConfig, odtConfig, texConfig, chunksConfig, ...commonProps } = userConfig as any;
        Object.assign(config, withoutPrototypeKeys(commonProps));

        // Merge sub-configs individually, ignoring undefined properties to preserve defaults
        const mergeSubConfig = (target: any, source: any) => {
            if (!source) return;
            for (const key in source) {
                // Both guards are load-bearing and neither subsumes the other. The own-property
                // check (matching deepClone above) stops inherited enumerable properties, which
                // matters once anything else in the process has already polluted a prototype. It
                // does NOT stop this attack on its own: `JSON.parse('{"__proto__":{...}}')` yields
                // `__proto__` as an own enumerable key, so it passes hasOwnProperty and would be
                // written straight through to Object.prototype by the recursion below.
                if (!Object.prototype.hasOwnProperty.call(source, key)) continue;
                if (PROTOTYPE_POLLUTION_KEYS.has(key)) continue;
                if (source[key] !== undefined) {
                    // Deep merge plain objects (like injections or margin)
                    if (
                        typeof source[key] === 'object' && 
                        source[key] !== null && 
                        !Array.isArray(source[key]) &&
                        !(source[key] instanceof Function) &&
                        !(source[key] instanceof Date) &&
                        !(source[key] instanceof RegExp) &&
                        !(source[key] instanceof Buffer)
                    ) {
                        if (!target[key] || typeof target[key] !== 'object') {
                            target[key] = {};
                        }
                        mergeSubConfig(target[key], source[key]);
                    } else {
                        target[key] = source[key];
                    }
                }
            }
        };

        if (htmlConfig) mergeSubConfig(config.htmlConfig, htmlConfig);
        if (mdConfig) mergeSubConfig(config.mdConfig, mdConfig);
        if (pdfConfig) mergeSubConfig(config.pdfConfig, pdfConfig);
        if (csvConfig) mergeSubConfig(config.csvConfig, csvConfig);
        if (textConfig) mergeSubConfig(config.textConfig, textConfig);
        if (rtfConfig) mergeSubConfig(config.rtfConfig, rtfConfig);
        if (docxConfig) mergeSubConfig(config.docxConfig, docxConfig);
        if (odtConfig) mergeSubConfig(config.odtConfig, odtConfig);
        if (texConfig) mergeSubConfig(config.texConfig, texConfig);
        if (chunksConfig) mergeSubConfig(config.chunksConfig, chunksConfig);
    }


    // 3. Inherit from AST config if not explicitly provided
    if (astConfig) {
        if (userConfig?.onWarning === undefined) {
            config.onWarning = astConfig.onWarning || config.onWarning;
        }

        // Inherit newlineDelimiter for text-based generators. Only a short one: an AST (from JSON, say)
        // carries its own config, and a delimiter is written at every line, so a long one multiplied the
        // output by the document's line count.
        const shortDelimiter = (value: unknown): string | undefined => typeof value === 'string' && value.length <= MAX_INHERITED_DELIMITER ? value : undefined;
        const astNewline = shortDelimiter(astConfig.newlineDelimiter);
        if (astNewline && ['text', 'md', 'rtf'].includes(destination)) {
            // If user didn't specify a newline delimiter in their specific config, use AST's
            if (destination === 'text' && (userConfig as any)?.textConfig?.newlineDelimiter === undefined) {
                config.textConfig.newlineDelimiter = astNewline;
            }
            // For MD and RTF, they use common newline settings or internal defaults.
            // We ensure the resolved config reflects this if possible, or generators can check astConfig directly.
            // Since FullGeneratorConfig doesn't have an 'mdConfig', we rely on the generator implementation.
        }

        // Inherit the parse-side `csvDelimiter` into CSV output when the user did not set the
        // generator's `csvConfig.columnDelimiter`, so `parseOffice(f, { csvDelimiter: ';' }).to('csv')`
        // matches the CLI's `--csvDelimiter=';'` (which wires the same propagation). Precedence:
        // csvConfig.columnDelimiter > csvDelimiter > ','.
        // One character from the AST's own config, and none that ends a row or starts a formula: the AST
        // (built in code, or JSON) is not the caller, and a delimiter of `\n=` wrote a formula line that
        // no cell of the document held.
        const astCsvDelim = shortDelimiter(astConfig.csvDelimiter);
        const safeAstCsvDelim = astCsvDelim && astCsvDelim.length === 1 && !'\r\n"=+-@'.includes(astCsvDelim) ? astCsvDelim : undefined;
        if (safeAstCsvDelim && destination === 'csv' && (userConfig as any)?.csvConfig?.columnDelimiter === undefined) {
            config.csvConfig.columnDelimiter = safeAstCsvDelim;
        }
    }

    resolvedConfig = config;
    if (unknownKeys.length) logWarning(OfficeWarningType.UNRECOGNIZED_CONFIG_OPTION, reporter, { keys: unknownKeys });
    validateHtmlConfigWidth(config.htmlConfig, reporter);
    validateGeneratorChoices(config, reporter);
    return config;
}

/** The longest line or column delimiter a generator takes from the AST's own config (see resolveGeneratorConfig). */
const MAX_INHERITED_DELIMITER = 16;

/**
 * Generator options that take one of a fixed set of named values, with the values each accepts.
 * Kept here, next to the check that uses them; each list is the runtime form of its type in
 * types.ts (a unit test keeps the Markdown presets in step with the generator's table).
 */
const GENERATOR_CHOICES: { container: string; key: string; choices: readonly string[]; caseInsensitive?: boolean }[] = [
    { container: 'texConfig', key: 'documentClass', choices: ['auto', 'article', 'report', 'book', 'beamer'] },
    { container: 'pdfConfig', key: 'engine', choices: ['html', 'native'] },
    { container: 'chunksConfig', key: 'strategy', choices: ['fixed-size', 'document-structure', 'semantic'] },
    { container: 'chunksConfig', key: 'splitBy', choices: ['page', 'slide', 'sheet', 'heading', 'paragraph'] },
    { container: 'chunksConfig', key: 'tableSplitStrategy', choices: ['row', 'flatten'] },
    ...['pdfConfig', 'docxConfig', 'odtConfig', 'texConfig'].map(container => ({ container, key: 'format', choices: PAPER_FORMAT_NAMES, caseInsensitive: true })),
];

/** The Markdown dialect presets `mdConfig.dialect` (or its `extends`) may name. */
export const MARKDOWN_DIALECT_PRESET_NAMES: readonly string[] = ['extended', 'github', 'gitlab', 'obsidian', 'pandoc', 'commonmark'];

/** The generator configs with page margins, whose sides take points or a length. */
const MARGIN_CONTAINERS = ['pdfConfig', 'docxConfig', 'odtConfig', 'texConfig'];

/** Whether a margin side is usable: a finite number of points, a length (`'1in'`, `'2cm'`), or `''` (unset). */
function isValidMargin(value: unknown): boolean {
    if (typeof value === 'number') return Number.isFinite(value);
    if (typeof value !== 'string') return false;
    return value.trim() === '' || /^\s*-?(\d+(?:\.\d*)?|\.\d+)\s*(?:(pt|in|cm|mm|px)\s*)?$/i.test(value);
}

/**
 * Checks the generator options that take named values (see GENERATOR_CHOICES), the Markdown
 * dialect preset and the page margins. An unusable value is reported with INVALID_CONFIG_VALUE
 * and replaced by the option's default, so every generator sees the same, valid choice rather
 * than each falling back its own way.
 */
function validateGeneratorChoices(config: FullGeneratorConfig, reporter: any): void {
    const anyConfig = config as any;
    const invalid = (option: string, value: unknown, expected: string, fallback: unknown) =>
        logWarning(OfficeWarningType.INVALID_CONFIG_VALUE, reporter, { option, value, expected, fallback });
    for (const { container, key, choices, caseInsensitive } of GENERATOR_CHOICES) {
        const target = anyConfig[container];
        const value = target?.[key];
        if (value === undefined) continue;
        const ok = typeof value === 'string' && choices.includes(caseInsensitive ? value.toLowerCase() : value);
        if (ok) continue;
        const fallback = (DEFAULT_GENERATOR_CONFIG as any)[container]?.[key];
        invalid(`${container}.${key}`, value, `one of ${choices.join(', ')}`, fallback);
        target[key] = fallback;
    }
    // Nested objects (a dialect object, a margin) may still be the caller's own: a fix replaces them
    // with a corrected copy rather than writing into them.
    const md = anyConfig.mdConfig;
    const presetOk = (v: unknown) => typeof v === 'string' && MARKDOWN_DIALECT_PRESET_NAMES.includes(v);
    const expectedPreset = `one of ${MARKDOWN_DIALECT_PRESET_NAMES.join(', ')}`;
    if (typeof md?.dialect === 'string' || (md?.dialect !== undefined && (typeof md.dialect !== 'object' || md.dialect === null))) {
        if (!presetOk(md.dialect)) {
            invalid('mdConfig.dialect', md.dialect, expectedPreset, 'extended');
            md.dialect = 'extended';
        }
    } else if (md?.dialect && md.dialect.extends !== undefined && !presetOk(md.dialect.extends)) {
        invalid('mdConfig.dialect.extends', md.dialect.extends, expectedPreset, 'extended');
        md.dialect = { ...md.dialect, extends: 'extended' };
    }
    // `htmlConfig.standalone` is a boolean or an object whose `styles` names a stylesheet mode.
    const standalone = anyConfig.htmlConfig?.standalone;
    const STYLE_MODES = ['full', 'scoped', 'none'];
    if (standalone && typeof standalone === 'object' && standalone.styles !== undefined && !STYLE_MODES.includes(standalone.styles)) {
        invalid('htmlConfig.standalone.styles', standalone.styles, `one of ${STYLE_MODES.join(', ')}`, 'full');
        anyConfig.htmlConfig.standalone = { ...standalone, styles: 'full' };
    }
    for (const container of MARGIN_CONTAINERS) {
        const margin = anyConfig[container]?.margin;
        if (!margin || typeof margin !== 'object') continue;
        for (const side of ['top', 'right', 'bottom', 'left']) {
            const value = anyConfig[container].margin[side];
            if (value === undefined || isValidMargin(value)) continue;
            const fallback = (DEFAULT_GENERATOR_CONFIG as any)[container]?.margin?.[side];
            invalid(`${container}.margin.${side}`, value, "a number of points or a length such as '1in', '2cm', '36pt'", fallback);
            anyConfig[container].margin = { ...anyConfig[container].margin, [side]: fallback };
        }
    }
}

/**
 * Validates the containerWidth option for HTML generation.
 * Can be 'auto', a positive number, or a positive CSS length/percentage string.
 */
export function isValidContainerWidth(width: any): boolean {
    if (width === 'auto') return true;
    if (typeof width === 'number') {
        return Number.isFinite(width) && width > 0;
    }
    if (typeof width === 'string') {
        const val = width.trim().toLowerCase();
        if (val === 'auto') return true;
        const match = val.match(/^((?:\d*\.)?\d+)(px|%|em|rem|vw|vh|vmin|vmax|ch|in|cm|mm|pt|pc)?$/);
        if (!match) return false;
        const numericValue = parseFloat(match[1]);
        return numericValue > 0;
    }
    return false;
}

/**
 * Emits a warning and falls back to 'auto' if the HTML containerWidth is invalid.
 */
function validateHtmlConfigWidth(htmlConfig: any, reporter: any): void {
    if (htmlConfig?.containerWidth !== undefined) {
        const width = htmlConfig.containerWidth;
        if (!isValidContainerWidth(width)) {
            logWarning(OfficeWarningType.INVALID_CONTAINER_WIDTH, reporter, width);
            htmlConfig.containerWidth = 'auto';
        }
    }
}

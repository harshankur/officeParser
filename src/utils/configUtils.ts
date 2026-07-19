import { DEFAULT_GENERATOR_CONFIG, DEFAULT_OFFICE_PARSER_CONFIG } from '../defaults.js';
import { FullGeneratorConfig, FullOfficeParserConfig, GeneratorConfig, OfficeParserConfig, OfficeWarningType } from '../types.js';
import { logWarning } from './errorUtils.js';

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
const PROTOTYPE_POLLUTION_KEYS = new Set(['__proto__', 'constructor', 'prototype']);

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
 * @param userConfig - Optional configuration provided by the user
 * @returns A fully populated configuration object
 */
export function resolveParserConfig(
    userConfig?: OfficeParserConfig | FullOfficeParserConfig
): FullOfficeParserConfig {
    if (isFullParserConfig(userConfig)) {
        if (!userConfig.decompressionLimits) {
            userConfig.decompressionLimits = {
                maxUncompressedBytes: 512 * 1024 * 1024,
                maxZipEntries: 10000,
            };
        }
        return userConfig;
    }

    // 1. Start with full defaults (deep cloned)
    const config: FullOfficeParserConfig = deepClone(DEFAULT_OFFICE_PARSER_CONFIG);

    if (!userConfig) {
        return config;
    }

    // 2. Merge user config
    // We handle ocrConfig, decompressionLimits, and htmlParserConfig specially to avoid
    // shallow-overwriting the whole nested objects
    const { ocrConfig, decompressionLimits, htmlParserConfig, ...rest } = userConfig;
    Object.assign(config, withoutPrototypeKeys(rest));

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

    // 3. Handle legacy ocrLanguage mapping if not explicitly set in ocrConfig
    if (userConfig.ocrLanguage && !userConfig.ocrConfig?.language) {
        config.ocrConfig.language = userConfig.ocrLanguage;
    }

    // 4. Propagate the top-level abortSignal to ocrConfig so the OCR subsystem is aware of it
    if (config.abortSignal) {
        config.ocrConfig.abortSignal = config.abortSignal;
    }

    return config;
}

/**
 * Resolves a full, destination-specific configuration by merging defaults, 
 * AST-level settings, and user-provided overrides.
 * 
 * @param destination - The target format
 * @param userConfig - Optional configuration provided by the user
 * @param astConfig - Optional configuration from the source AST (for inheritance)
 * @returns A fully populated configuration object
 */
export function resolveGeneratorConfig<D extends string>(
    destination: D,
    astConfig?: OfficeParserConfig,
    userConfig?: GeneratorConfig<D> | FullGeneratorConfig
): FullGeneratorConfig {
    // If it's already a full config and we don't need to merge AST config, return it as is.
    // We assume FullGeneratorConfig is already "safe" (references resolved).
    if (isFullGeneratorConfig(userConfig) && !astConfig) {
        validateHtmlConfigWidth(userConfig.htmlConfig, userConfig);
        return userConfig;
    }

    // 1. Start with full defaults (deep cloned to avoid reference sharing)
    const config: FullGeneratorConfig = deepClone(DEFAULT_GENERATOR_CONFIG);

    // 2. Merge common properties and sub-configs
    if (userConfig) {
        // Extract sub-configs to avoid shallow-overwriting the whole sub-config objects
        const { htmlConfig, mdConfig, pdfConfig, csvConfig, textConfig, chunksConfig, ...commonProps } = userConfig as any;
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
        if (chunksConfig) mergeSubConfig(config.chunksConfig, chunksConfig);
    }


    // 3. Inherit from AST config if not explicitly provided
    if (astConfig) {
        if (userConfig?.onWarning === undefined) {
            config.onWarning = astConfig.onWarning || config.onWarning;
        }

        // Inherit newlineDelimiter for text-based generators
        const astNewline = astConfig.newlineDelimiter;
        if (astNewline && ['text', 'md', 'rtf'].includes(destination)) {
            // If user didn't specify a newline delimiter in their specific config, use AST's
            if (destination === 'text' && (userConfig as any)?.textConfig?.newlineDelimiter === undefined) {
                config.textConfig.newlineDelimiter = astNewline;
            }
            // For MD and RTF, they use common newline settings or internal defaults.
            // We ensure the resolved config reflects this if possible, or generators can check astConfig directly.
            // Since FullGeneratorConfig doesn't have an 'mdConfig', we rely on the generator implementation.
        }
    }

    validateHtmlConfigWidth(config.htmlConfig, config);
    return config;
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
function validateHtmlConfigWidth(htmlConfig: any, config: any): void {
    if (htmlConfig?.containerWidth !== undefined) {
        const width = htmlConfig.containerWidth;
        if (!isValidContainerWidth(width)) {
            logWarning(OfficeWarningType.INVALID_CONTAINER_WIDTH, config as any, width);
            htmlConfig.containerWidth = 'auto';
        }
    }
}

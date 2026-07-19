import { DEFAULT_GENERATOR_CONFIG, DEFAULT_OFFICE_PARSER_CONFIG } from '../defaults.js';
import { FullGeneratorConfig, FullOfficeParserConfig, GeneratorConfig, OfficeParserConfig, OfficeWarningType } from '../types.js';
import { logWarning } from './errorUtils.js';

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
    Object.assign(config, rest);

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
        Object.assign(config, commonProps);

        // Merge sub-configs individually, ignoring undefined properties to preserve defaults
        const mergeSubConfig = (target: any, source: any) => {
            if (!source) return;
            for (const key in source) {
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

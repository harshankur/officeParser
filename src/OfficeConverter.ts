import { OfficeGenerator } from './OfficeGenerator.js';
import { OfficeParser } from './OfficeParser.js';
import { BlobLike, ConversionResult, GeneratorConfig, OfficeConverterConfig, OfficeIssue, OfficeParserConfig, OfficeWarningType, SupportedDestination, SupportedFileType } from './types.js';
import { PROTOTYPE_POLLUTION_KEYS, RECOGNIZED_GENERATOR_KEYS, RECOGNIZED_PARSER_KEYS } from './utils/configUtils.js';
import { logWarning } from './utils/errorUtils.js';
import { resolveImageMode } from './utils/officeGenUtils.js';

/** The keys `convert()` itself reads; parser and generator options go inside the first two. */
const CONVERTER_KEYS = new Set(['parseConfig', 'generatorConfig', 'onWarning']);

/**
 * A `convert()` config's top-level keys it does not read, each with where it belongs when it is a
 * parser or generator option (`{ texConfig }` belongs in `generatorConfig`), as the
 * UNRECOGNIZED_CONFIG_OPTION warning reports them.
 */
function misplacedConverterKeys(config: object | undefined): { keys: string[]; renames: Record<string, string> } {
    const keys = Object.keys(config ?? {}).filter(k => !CONVERTER_KEYS.has(k) && !PROTOTYPE_POLLUTION_KEYS.has(k));
    const renames: Record<string, string> = {};
    for (const key of keys) {
        const parser = RECOGNIZED_PARSER_KEYS.has(key), generator = RECOGNIZED_GENERATOR_KEYS.has(key);
        if (parser && generator) renames[key] = '(it goes under parseConfig or generatorConfig)';
        else if (parser) renames[key] = `parseConfig.${key}`;
        else if (generator) renames[key] = `generatorConfig.${key}`;
    }
    return { keys, renames };
}

/**
 * Utility type to infer the file type from a file path string literal.
 */
type InferFileTypeFromPath<T> = T extends `${string}.${infer E}`
    ? (Lowercase<E> extends SupportedFileType ? Lowercase<E> : SupportedFileType)
    : SupportedFileType;

/**
 * Main converter class providing a streamlined one-step API for document conversion.
 * 
 * This class coordinates the `OfficeParser` and `OfficeGenerator` to transform 
 * documents from one format to another (e.g., DOCX to Markdown, PDF to HTML).
 */
export class OfficeConverter {
    /**
     * Converts an office document from its source format to a specified destination format.
     * 
     * This method:
     * 1. Detects the source file type and parses it into a unified AST using `OfficeParser`.
     * 2. Automatically configures the parser based on the generator requirements (e.g., enabling
     *    attachment extraction if images are requested in the output).
     * 3. Generates the destination document from the AST using `OfficeGenerator`.
     * 
     * @template F The inferred type of the input file (path string or buffer).
     * @template T The authoritative source file type (inferred from path or config).
     * 
     * @param file - File path (string), Buffer, or ArrayBuffer containing the source document.
     * @param destination - The target format (e.g., 'md', 'html', 'pdf', 'text', 'chunks').
     * @param config - Optional unified configuration for both the parser and generator phases.
     * 
     * @returns A promise resolving to the ConversionResult containing the value and messages.
     * @throws {Error} If the source format is unsupported or parsing/generation fails.
     * 
     * @example
     * ```typescript
     * // Convert Word to Markdown with a single call
     * const { value: markdown } = await OfficeConverter.convert('report.docx', 'md');
     * 
     * // Convert PDF to HTML (for OCR, pass parseConfig: { ocr: true } - attachments auto-extract)
     * const { value: html } = await OfficeConverter.convert(buffer, 'html', {
     *   generatorConfig: {
     *     includeImages: true
     *   }
     * });
     * ```
     */
    public static async convert<
        F extends string | Buffer | ArrayBuffer | Uint8Array | BlobLike,
        T extends SupportedFileType = InferFileTypeFromPath<F>,
        D extends SupportedDestination<T> = SupportedDestination<T>
    >(
        file: F,
        destination: D,
        config?: OfficeConverterConfig<D, T>
    ): Promise<ConversionResult<D>> {
        // 1. Prepare Parser Configuration
        // We prioritize the top-level onWarning if provided.
        const parserConfig: OfficeParserConfig = {
            ...config?.parseConfig,
            onWarning: config?.onWarning || config?.parseConfig?.onWarning,
        };

        // Whether the caller pinned `extractAttachments` explicitly (true OR false); an explicit value
        // always wins over the auto-sync below. Captured before the undefined-key prune.
        const callerSetAttachments = config?.parseConfig?.extractAttachments !== undefined;

        // Remove undefined keys to prevent overwriting defaults in resolveParserConfig
        (Object.keys(parserConfig) as (keyof OfficeParserConfig)[]).forEach(
            (key) => parserConfig[key] === undefined && delete parserConfig[key]
        );

        /**
         * AUTOMATIC CONFIGURATION SYNC
         * `parseConfig.ocr` is honored (it is no longer forced off). We sync `extractAttachments` from
         * the generator configuration unless the caller set it explicitly.
         */
        if (!callerSetAttachments) {
            // Extract attachments when the generator will render an image or its OCR text (any
            // includeImages mode except false/'none'), when charts are included, or when the caller
            // enabled OCR (which needs the extracted images to run over).
            const im = config?.generatorConfig?.includeImages;
            // Resolve through the shared mapper so a CLI-style `'false'` string (and `'none'`) is honored,
            // not just the boolean/'none' literals - otherwise `--includeImages=false` still extracts.
            const wantsImageOrText = resolveImageMode(im) !== 'none';
            parserConfig.extractAttachments = wantsImageOrText || (config?.generatorConfig?.includeCharts !== false) || !!parserConfig.ocr;
        }

        // A parser or generator option at the top level is not read; say where it belongs.
        const misplaced = misplacedConverterKeys(config);
        const configIssues: OfficeIssue[] = [];
        if (misplaced.keys.length) {
            const onWarning = config?.onWarning;
            logWarning(OfficeWarningType.UNRECOGNIZED_CONFIG_OPTION, { onWarning: (issue: OfficeIssue) => { configIssues.push(issue); onWarning?.(issue); } }, misplaced);
        }

        // 2. Parse the source document into the universal AST
        const ast = await OfficeParser.parseOffice(file, parserConfig);
        // The parse's own warnings, taken now: the AST goes on collecting the generator's too, which
        // the generation result already lists.
        const parseWarnings = [...(ast.warnings || [])];

        // 3. Generate the destination document from the AST
        const generatorConfig = {
            ...config?.generatorConfig,
            onWarning: config?.onWarning || config?.generatorConfig?.onWarning,
        };

        // The spread widens the object to an inferred literal; it is a `GeneratorConfig<D>` by
        // construction (config.generatorConfig is already typed for D, and onWarning is common to
        // every destination), so the cast restates what the types otherwise lose.
        const result = await OfficeGenerator.generate(ast, destination, generatorConfig as GeneratorConfig<D>);
        result.messages = [...configIssues, ...parseWarnings, ...result.messages];
        return result;
    }
}

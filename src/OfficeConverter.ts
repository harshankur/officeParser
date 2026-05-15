import { OfficeGenerator } from './OfficeGenerator.js';
import { OfficeParser } from './OfficeParser.js';
import { ConversionResult, OfficeConverterConfig, OfficeParserConfig, SupportedDestination, SupportedFileType } from './types.js';

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
     * // Convert PDF to HTML (Note: OCR is disabled in this one-step API)
     * const { value: html } = await OfficeConverter.convert(buffer, 'html', {
     *   generatorConfig: {
     *     includeImages: true
     *   }
     * });
     * ```
     */
    public static async convert<
        F extends string | Buffer | ArrayBuffer | Uint8Array,
        T extends SupportedFileType = InferFileTypeFromPath<F>
    >(
        file: F,
        destination: SupportedDestination<T>,
        config?: OfficeConverterConfig<SupportedDestination<T>, T>
    ): Promise<ConversionResult<SupportedDestination<T>>> {
        // 1. Prepare Parser Configuration
        // We prioritize the top-level onWarning if provided.
        const parserConfig: OfficeParserConfig = {
            ...config?.parseConfig,
            onWarning: config?.onWarning || config?.parseConfig?.onWarning,
        };

        // Remove OCR settings for the streamlined converter as requested
        parserConfig.ocr = false;

        // Remove undefined keys to prevent overwriting defaults in resolveParserConfig
        (Object.keys(parserConfig) as (keyof OfficeParserConfig)[]).forEach(
            (key) => parserConfig[key] === undefined && delete parserConfig[key]
        );

        /**
         * AUTOMATIC CONFIGURATION SYNC
         * We sync extractAttachments from the generator configuration.
         */
        parserConfig.extractAttachments = (config?.generatorConfig?.includeImages !== false) || (config?.generatorConfig?.includeCharts !== false);

        // 2. Parse the source document into the universal AST
        const ast = await OfficeParser.parseOffice(file, parserConfig);

        // 3. Generate the destination document from the AST
        const generatorConfig = {
            ...config?.generatorConfig,
            onWarning: config?.onWarning || config?.generatorConfig?.onWarning,
        };

        const result = await OfficeGenerator.generate(ast, destination, generatorConfig);
        result.messages = [...(ast.warnings || []), ...result.messages];
        return result;
    }
}

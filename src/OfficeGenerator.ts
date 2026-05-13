import { ChunkingGenerator } from './generators/ChunkingGenerator.js';
import { CsvGenerator } from './generators/CsvGenerator.js';
import { HtmlGenerator } from './generators/HtmlGenerator.js';
import { MarkdownGenerator } from './generators/MarkdownGenerator.js';
import { PdfGenerator } from './generators/PdfGenerator.js';
import { RtfGenerator } from './generators/RtfGenerator.js';
import { TextGenerator } from './generators/TextGenerator.js';
import { ConversionResult, GeneratorConfig, OfficeErrorType, OfficeParserAST, SupportedDestination, SupportedFileType } from './types.js';
import { getOfficeError } from './utils/errorUtils.js';

/**
 * Main generator class providing document conversion functionality.
 */
export class OfficeGenerator {
    /**
     * Generates a file of the specified type from an AST.
     * This is the single source of truth for generation logic.
     * 
     * @param ast - The OfficeParserAST to generate from
     * @param destination - The target format (e.g., 'text', 'md', 'html', 'pdf')
     * @param config - Optional configuration for the generator
     * @returns A promise resolving to the ConversionResult containing the value and messages
     * @throws {Error} If the destination format is unsupported
     */
    public static async generate<T extends SupportedFileType, D extends SupportedDestination<T>>(
        ast: OfficeParserAST & { type: T },
        destination: D,
        config?: GeneratorConfig<D>
    ): Promise<ConversionResult> {
        switch (destination.toLowerCase() as SupportedDestination<T>) {
            case 'text':
                return new TextGenerator(ast, config as GeneratorConfig<'text'>).generate();
            case 'md':
                return new MarkdownGenerator(ast, config as GeneratorConfig<'md'>).generate();
            case 'html':
                return new HtmlGenerator(ast, config as GeneratorConfig<'html'>).generate();
            case 'pdf':
                return new PdfGenerator(ast, config as GeneratorConfig<'pdf'>).generate();
            case 'csv':
                return new CsvGenerator(ast, config as GeneratorConfig<'csv'>).generate();
            case 'rtf':
                return new RtfGenerator(ast, config as GeneratorConfig<'rtf'>).generate();
            case 'chunks':
                return new ChunkingGenerator(ast, config as GeneratorConfig<'chunks'>).generate();

            default:
                throw getOfficeError(OfficeErrorType.EXTENSION_UNSUPPORTED, undefined, destination);
        }
    }
}

import { BaseGenerator } from './generators/BaseGenerator.js';
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
    ): Promise<ConversionResult<D>> {
        let generator: BaseGenerator<any>;

        switch (destination.toLowerCase() as SupportedDestination<T>) {
            case 'text':
                generator = new TextGenerator(ast, config as GeneratorConfig<'text'>);
                break;
            case 'md':
                generator = new MarkdownGenerator(ast, config as GeneratorConfig<'md'>);
                break;
            case 'html':
                generator = new HtmlGenerator(ast, config as GeneratorConfig<'html'>);
                break;
            case 'pdf':
                generator = new PdfGenerator(ast, config as GeneratorConfig<'pdf'>);
                break;
            case 'csv':
                generator = new CsvGenerator(ast, config as GeneratorConfig<'csv'>);
                break;
            case 'rtf':
                generator = new RtfGenerator(ast, config as GeneratorConfig<'rtf'>);
                break;
            case 'chunks':
                generator = new ChunkingGenerator(ast, config as GeneratorConfig<'chunks'>);
                break;
            default:
                throw getOfficeError(OfficeErrorType.EXTENSION_UNSUPPORTED, undefined, destination);
        }

        return generator.generate() as Promise<ConversionResult<D>>;
    }
}

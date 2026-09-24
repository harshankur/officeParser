import { BaseGenerator } from './generators/BaseGenerator.js';
import { ChunkingGenerator } from './generators/ChunkingGenerator.js';
import { CsvGenerator } from './generators/CsvGenerator.js';
import { DocxGenerator } from './generators/DocxGenerator.js';
import { EpubGenerator } from './generators/EpubGenerator.js';
import { HtmlGenerator } from './generators/HtmlGenerator.js';
import { MarkdownGenerator } from './generators/MarkdownGenerator.js';
import { OdtGenerator } from './generators/OdtGenerator.js';
import { LatexGenerator } from './generators/LatexGenerator.js';
import { PdfGenerator } from './generators/PdfGenerator.js';
import { RtfGenerator } from './generators/RtfGenerator.js';
import { TextGenerator } from './generators/TextGenerator.js';
import { ConversionResult, GeneratorConfig, OfficeErrorType, OfficeParserAST, SupportedDestination, SupportedFileType, UniversalGeneratorFormat } from './types.js';
import { withoutSourceComments } from './utils/commentUtils.js';
import { getOfficeError } from './utils/errorUtils.js';

/**
 * Main generator class providing document conversion functionality.
 */
export class OfficeGenerator {
    /**
     * Normalizes format aliases (e.g., 'txt' to 'text', 'markdown' to 'md', 'latex' to 'tex') to standard internal formats.
     */
    public static normalizeDestination(dest: string): UniversalGeneratorFormat {
        const d = dest?.toLowerCase();
        if (d === 'txt') return 'text';
        if (d === 'markdown') return 'md';
        if (d === 'latex') return 'tex';
        return d as UniversalGeneratorFormat;
    }

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
        const normalizedDestination = OfficeGenerator.normalizeDestination(destination);
        // A source comment (`<!-- ... -->`, CommentMetadata.sourceSyntax 'html') is the author's hidden
        // note. Only Markdown, HTML and LaTeX (as `%` lines) can carry it as a comment; every other format
        // has no hidden-comment construct (EPUB is XHTML, where `--` inside a comment is illegal), so it is
        // removed here, once, rather than each generator having to remember to skip it.
        const keepsComments = normalizedDestination === 'md' || normalizedDestination === 'html' || normalizedDestination === 'tex';
        const input = keepsComments ? ast : withoutSourceComments(ast);

        switch (normalizedDestination) {
            case 'text':
                generator = new TextGenerator(input, config as GeneratorConfig<'text'>);
                break;
            case 'md':
                generator = new MarkdownGenerator(input, config as GeneratorConfig<'md'>);
                break;
            case 'html':
                generator = new HtmlGenerator(input, config as GeneratorConfig<'html'>);
                break;
            case 'pdf':
                generator = new PdfGenerator(input, config as GeneratorConfig<'pdf'>);
                break;
            case 'csv':
                generator = new CsvGenerator(input, config as GeneratorConfig<'csv'>);
                break;
            case 'rtf':
                generator = new RtfGenerator(input, config as GeneratorConfig<'rtf'>);
                break;
            case 'chunks':
                generator = new ChunkingGenerator(input, config as GeneratorConfig<'chunks'>);
                break;
            case 'epub':
                generator = new EpubGenerator(input, config as GeneratorConfig<'epub'>);
                break;
            case 'docx':
                generator = new DocxGenerator(input, config as GeneratorConfig<'docx'>);
                break;
            case 'odt':
                generator = new OdtGenerator(input, config as GeneratorConfig<'odt'>);
                break;
            case 'tex':
                generator = new LatexGenerator(input, config as GeneratorConfig<'tex'>);
                break;
            default:
                throw getOfficeError(OfficeErrorType.FORMAT_UNSUPPORTED, undefined, destination);
        }

        return generator.generate() as Promise<ConversionResult<D>>;
    }
}

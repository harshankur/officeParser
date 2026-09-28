import { BaseGenerator } from './generators/BaseGenerator.js';
import { sharedNodeVisits } from './utils/nodeListUtils.js';
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
import { withKnownNodeTypes } from './utils/nodeTypeUtils.js';
import { withBoundedSheetGrids } from './utils/sheetGridUtils.js';
import { withWellTypedValues } from './utils/valueTypeUtils.js';
import { getOfficeError } from './utils/errorUtils.js';

/** Nodes a writer may meet more than once along the AST's paths, beyond what repeated content allows (see sharedNodeVisits). */
const MAX_SHARED_NODE_VISITS = 4_000_000;

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
        // A node of a type the AST does not define is written as its content, which no generator
        // then has to know how to handle; a value of another type than the AST defines (an array
        // for a string, text for a number) is coerced or removed, so none reaches a writer that
        // escapes one kind and writes another raw. A table or sheet whose cells span a grid too
        // large to write (a few cells far apart) is laid out closer, before any writer fills it.
        // An AST nested past what the stack holds (a sheet in a cell in a sheet, 800 deep), or output past
        // what a string holds, is rejected with an OfficeParser error, not the engine's own RangeError.
        const asNestingError = (error: unknown): unknown => {
            if (!(error instanceof RangeError)) return error;
            const reporter = config?.onWarning ? config : ast.config ?? config;
            if (/call stack/i.test(error.message)) return getOfficeError(OfficeErrorType.MAX_NESTING_DEPTH_EXCEEDED, reporter);
            if (/invalid (string|array) length|allocation failed/i.test(error.message)) return getOfficeError(OfficeErrorType.OUTPUT_TOO_LARGE, reporter);
            return error;
        };
        let input: OfficeParserAST;
        try {
            // Nodes shared along many paths (see sharedNodeVisits) past what the parsers' own sharing
            // reaches are refused before anything reads them along each path: the passes below write a
            // node of a type the AST does not define as its content once per path to it, and an AST of
            // such nodes each holding the next twice, 30 deep, ended the process out of memory before the
            // check after them ran.
            // The parse's repeat budget widens the allowance, a number up to 1 GiB of it: an AST built in
            // code carries a config of its own, and a limit of Infinity (or text) in it turned the check off.
            const configured: unknown = ast.config?.decompressionLimits?.maxRepeatedContent;
            const repeatedContent = typeof configured === 'number' && configured >= 0 ? Math.min(configured, 1024 * 1024 * 1024) : 16 * 1024 * 1024;
            const tooShared = (candidate: OfficeParserAST) => sharedNodeVisits(candidate) > MAX_SHARED_NODE_VISITS + repeatedContent / 16;
            if (tooShared(ast)) throw getOfficeError(OfficeErrorType.OUTPUT_TOO_LARGE, config?.onWarning ? config : ast.config ?? config);
            const known = withBoundedSheetGrids(withKnownNodeTypes(withWellTypedValues(ast)));
            input = keepsComments ? known : withoutSourceComments(known);
            if (tooShared(input)) throw getOfficeError(OfficeErrorType.OUTPUT_TOO_LARGE, config?.onWarning ? config : ast.config ?? config);
        } catch (error) {
            throw asNestingError(error);
        }

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
                // Reported where a generator would report it: the caller's handler, else the AST's.
                throw getOfficeError(OfficeErrorType.FORMAT_UNSUPPORTED, config?.onWarning ? config : ast.config ?? config, destination);
        }

        try {
            return await generator.generate() as ConversionResult<D>;
        } catch (error) {
            throw asNestingError(error);
        }
    }
}

import { OfficeGenerator } from '../OfficeGenerator.js';
import { ConversionResult, GeneratorConfig, OfficeAttachment, OfficeAuxiliaryContent, OfficeContentNode, OfficeMetadata, OfficeParserAST, OfficeParserConfig, SupportedDestination, SupportedFileType } from '../types.js';

/**
 * Creates a fully-featured OfficeParserAST object with conversion methods.
 * 
 * This helper ensures that all ASTs returned by officeParser have the latest
 * conversion methods (.to()) and maintain backward compatibility (.toText()).
 * 
 * @param type - The detected file type
 * @param metadata - Document metadata
 * @param content - Parsed content nodes
 * @param attachments - Extracted attachments
 * @param config - Original parser configuration
 * @param toTextSync - Synchronous text extraction logic (for backward compatibility)
 * @returns An object conforming to OfficeParserAST
 */
export function createAST(
    type: SupportedFileType,
    metadata: OfficeMetadata,
    content: OfficeContentNode[],
    attachments: OfficeAttachment[],
    config: OfficeParserConfig,
    auxiliary: OfficeAuxiliaryContent | undefined,
    toTextSync: () => string
): OfficeParserAST {
    return {
        config,
        type,
        metadata,
        content,
        attachments,
        auxiliary,
        warnings: [],
        toText: toTextSync,
        async to<T extends OfficeParserAST, D extends SupportedDestination<T['type']>>(
            this: T,
            destination: D,
            genConfig?: GeneratorConfig<D>
        ): Promise<ConversionResult<D>> {
            return OfficeGenerator.generate(this as any, destination, genConfig) as Promise<ConversionResult<D>>;
        }
    };
}

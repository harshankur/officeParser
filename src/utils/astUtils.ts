import { boundRepeatedValues } from './repeatUtils.js';
import { OfficeGenerator } from '../OfficeGenerator.js';
import { CanonicalFormat, ConversionResult, GeneratorConfig, OfficeAttachment, OfficeAuxiliaryContent, OfficeContentNode, OfficeMetadata, OfficeParserAST, OfficeParserConfig, SupportedDestination, SupportedFileType } from '../types.js';

/**
 * Creates a fully-featured OfficeParserAST object with conversion methods.
 *
 * This helper ensures that all ASTs returned by officeParser expose the `.to()` conversion method.
 *
 * @param type - The detected file type
 * @param metadata - Document metadata
 * @param content - Parsed content nodes
 * @param attachments - Extracted attachments
 * @param config - Original parser configuration
 * @param auxiliary - Out-of-band content (headers, footers, slide masters)
 * @returns An object conforming to OfficeParserAST
 */
export function createAST(
    type: SupportedFileType,
    metadata: OfficeMetadata,
    content: OfficeContentNode[],
    attachments: OfficeAttachment[],
    config: OfficeParserConfig,
    auxiliary: OfficeAuxiliaryContent | undefined,
): OfficeParserAST {
    // Last, over everything the parser built: values a definition gave many nodes are bounded in all.
    boundRepeatedValues([content, auxiliary?.headers, auxiliary?.footers, auxiliary?.slideMasters, auxiliary?.outline], attachments, config);
    return {
        config,
        type,
        metadata,
        content,
        attachments,
        auxiliary,
        warnings: [],
        async to<T extends OfficeParserAST, D extends SupportedDestination<T['type']>>(
            this: T,
            destination: D,
            genConfig?: GeneratorConfig<D>
        ): Promise<ConversionResult<CanonicalFormat<D>>> {
            return OfficeGenerator.generate(this as any, destination, genConfig) as Promise<ConversionResult<CanonicalFormat<D>>>;
        }
    };
}

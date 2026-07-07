import { OfficeIssue, ConversionResult, FullGeneratorConfig, GeneratorConfig, OfficeContentNode, OfficeParserAST, OfficeWarningType, StructuredStyleMapping, UniversalGeneratorFormat } from '../types.js';
import { resolveGeneratorConfig } from '../utils/configUtils.js';
import { getWarningMessage } from '../utils/errorUtils.js';
import { StyleMapper } from '../utils/styleMapper.js';

/**
 * Base class for all document generators.
 * Provides common traversal logic and configuration handling.
 */
export abstract class BaseGenerator<D extends UniversalGeneratorFormat = UniversalGeneratorFormat> {
    protected config: FullGeneratorConfig;
    protected ast: OfficeParserAST;
    protected messages: OfficeIssue[] = [];
    protected styleMapper: StyleMapper;
    protected collectedNotes: OfficeContentNode[] = [];

    constructor(protected destination: D, ast: OfficeParserAST, config?: GeneratorConfig<D> | FullGeneratorConfig) {
        this.config = resolveGeneratorConfig(destination, ast.config, config);
        this.ast = ast;
        this.styleMapper = new StyleMapper(this.config.styleMap, this.config.ignoreDefaultStyleMap);
    }

    /**
     * Retrieves the semantic mapping for a node, respecting the includeFormatting flag.
     * Per design requirements: Style mapping is bypassed if formatting is disabled.
     */
    protected getSemanticMapping(node: OfficeContentNode) {
        if (this.config.includeFormatting === false) {
            return undefined;
        }
        return this.styleMapper.getMapping(node);
    }

    /**
     * Entry point for generation.
     */
    abstract generate(): Promise<ConversionResult<D>>;

    /**
     * Centralized logic for handling the onNode callback.
     * Evaluates the callback and returns a result that tells the generator how to proceed.
     * 
     * @returns 
     * - `string`: Use this as the node's output, skip default processing.
     * - `false`: Skip this node and its subtree.
     * - `void`: Proceed with default processing.
     */
    protected async handleOnNode(node: OfficeContentNode): Promise<string | false | void> {
        const result = await this.config.onNode(node);

        if (result === false) return false;
        if (typeof result === 'string') return result;
    }

    /**
     * Recursively processes nodes and builds output.
     * 
     * @param node - The current node being processed
     * @param processor - A function that takes a node and its children's output and returns the node's output string.
     * @returns The generated string for this node and its subtree.
     */
    protected async processNodeRecursive(
        node: OfficeContentNode,
        processor: (node: OfficeContentNode, childrenOutput: string) => string | Promise<string>
    ): Promise<string> {
        const override = await this.handleOnNode(node);

        if (override === false) return '';
        if (typeof override === 'string') return override;

        let childrenOutput = '';
        if (node.children) {
            for (const child of node.children) {
                childrenOutput += await this.processNodeRecursive(child, processor);
            }
        }

        if (node.notes && node.notes.length > 0) {
            if (node.type !== 'slide') {
                this.collectedNotes.push(...node.notes);
            }
        }

        let result = await processor(node, childrenOutput);

        if (node.type === 'slide' && node.notes && node.notes.length > 0) {
            for (const note of node.notes) {
                result += await this.processNodeRecursive(note, processor);
            }
        }

        return result;
    }

    /**
     * Helper to generate a unique ID (slug) from text.
     */
    protected slugify(text: string): string {
        return text
            .toLowerCase()
            .replace(/[^\w\s-]/g, '')
            .replace(/[\s_-]+/g, '-')
            .replace(/^-+|-+$/g, '');
    }

    private noteFootnoteKeys = new Map<OfficeContentNode, string>();
    private usedFootnoteKeys = new Set<string>();
    private footnoteKeyCounter = 0;

    /**
     * Assigns a stable, unique reference key to a footnote/endnote node, reused for both
     * its inline reference marker and its collected definition. Source ids aren't reliably
     * unique across a document - DOCX/ODT number footnotes and endnotes in separate
     * sequences, so both can carry noteId "1" - so a source id is only reused when it
     * hasn't already been claimed; otherwise a sequential counter guarantees uniqueness.
     */
    protected getFootnoteKey(note: OfficeContentNode): string {
        const cached = this.noteFootnoteKeys.get(note);
        if (cached) return cached;

        const preferred = (note.metadata as any)?.noteId;
        let key: string;
        if (preferred && !this.usedFootnoteKeys.has(preferred)) {
            key = preferred;
        } else {
            do {
                this.footnoteKeyCounter++;
                key = String(this.footnoteKeyCounter);
            } while (this.usedFootnoteKeys.has(key));
        }

        this.usedFootnoteKeys.add(key);
        this.noteFootnoteKeys.set(note, key);
        return key;
    }

    /**
     * Recursively extracts plain text from a node and its children.
     */
    protected getNodeText(node: OfficeContentNode): string {
        if (node.text) return node.text;
        if (node.children) {
            return node.children.map(c => this.getNodeText(c)).join('');
        }
        return '';
    }

    /**
     * Reports a warning to the user and collects it for the final result.
     */
    protected warn(type: OfficeWarningType, info?: any, node?: OfficeContentNode): void {
        const message = getWarningMessage(type, info);
        const issue: OfficeIssue = {
            type: 'warning',
            code: type,
            message,
            node,
            details: info
        };
        this.messages.push(issue);
        this.config.onWarning(issue);
    }
}

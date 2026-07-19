import { ConversionResult, GeneratorConfig, OfficeContentNode, OfficeParserAST } from '../types.js';
import { BaseGenerator } from './BaseGenerator.js';

const escapeRegExpChars = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * Generates plain text from an AST.
 */
export class TextGenerator extends BaseGenerator<'text'> {
    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'text'>) {
        super('text', ast, config);
    }


    /**
     * Generates plain text by concatenating text content from nodes.
     */
    async generate(): Promise<ConversionResult<'text'>> {
        let output = '';
        const newline = this.config.textConfig.newlineDelimiter;

        // Add Metadata Header
        if (this.config.renderMetadata && this.ast.metadata) {
            if (this.ast.metadata.title) output += `Title: ${this.ast.metadata.title}${newline}`;
            if (this.ast.metadata.author) output += `Author: ${this.ast.metadata.author}${newline}`;
            if (this.ast.metadata.created) output += `Created: ${new Date(this.ast.metadata.created).toLocaleString()}${newline}`;
            output += `-------------------${newline}${newline}`;
        }

        const processor = async (node: OfficeContentNode, childrenOutput: string): Promise<string> => {
            // Return raw text for text nodes
            if (node.type === 'text' || node.type === 'code') {
                return node.text || '';
            }

            // Handle explicit breaks
            if (node.type === 'break') {
                return newline;
            }

            if (node.type === 'image') {
                if (!this.config.includeImages) return '';
                const meta = node.metadata as any;
                return `[Image: ${meta?.altText || meta?.attachmentName || 'Untitled'}]${newline}`;
            }

            if (node.type === 'embed') {
                const meta = node.metadata as any;
                return meta?.url ? `[${meta.embedType === 'youtube' ? 'YouTube' : 'Embed'}: ${meta.url}]${newline}` : '';
            }

            if (node.type === 'admonition') {
                const meta = node.metadata as any;
                if (childrenOutput.trim() === '') return '';
                const label = (meta?.admonitionType || 'note').toUpperCase();
                return `[${label}] ${childrenOutput.trim()}${newline}`;
            }

            if (node.type === 'table' && this.config.textConfig.preserveLayout) {
                return await this.renderTable(node, processor, newline);
            }

            if (node.type === 'list' && this.config.textConfig.preserveLayout) {
                const meta = node.metadata as any;
                const indentSpaces = ' '.repeat(4);
                const indent = indentSpaces.repeat(meta?.indentation || 0);
                const marker = meta?.listType === 'ordered' ? `${(meta.itemIndex ?? 0) + 1}. ` : '- ';
                return `${indent}${marker}${childrenOutput.trimStart()}` + (childrenOutput.endsWith(newline) ? '' : newline);
            }

            // Append newline for block-level elements to maintain structure
            const blockTypes = ['paragraph', 'heading', 'row', 'sheet', 'slide', 'note', 'list', 'table', 'code'];
            if (blockTypes.includes(node.type)) {
                if (childrenOutput.trim() === '') return '';
                return childrenOutput + (childrenOutput.endsWith(newline) ? '' : newline);
            }

            return childrenOutput;
        };

        for (const node of this.ast.content) {
            output += await this.processNodeRecursive(node, processor);
        }

        if (this.collectedNotes.length > 0) {
            output += `${newline}${newline}--- Notes ---${newline}`;
            for (const note of this.collectedNotes) {
                output += await this.processNodeRecursive(note, processor);
            }
        }

        // Every block-level node above unconditionally appends its own trailing `newline` as a
        // separator from whatever sibling follows - including, unavoidably, the very last one,
        // which has no sibling to separate from - and renderTable below unconditionally *prepends*
        // one too, as a separator from whatever precedes it (also unavoidably applied when a table
        // is the very first/only node). Both are pure generator artifacts, never part of the
        // document's actual content, so a run of exactly this delimiter at either end is the only
        // thing safe to strip. Nothing else is: not leading/trailing spaces or tabs (e.g. an
        // intentionally-indented opening line, or trailing spaces on the last line - both real
        // content), and not any whitespace that isn't composed of this exact repeated delimiter. A
        // blanket trim()/trimEnd() would silently destroy all of those.
        const d = escapeRegExpChars(newline);
        const leadingOrTrailingArtifact = new RegExp(`^(?:${d})+|(?:${d})+$`, 'g');
        return {
            value: output.replace(leadingOrTrailingArtifact, ''),
            messages: this.messages
        };
    }

    private async renderTable(node: OfficeContentNode, processor: any, newline: string): Promise<string> {
        if (!node.children || node.children.length === 0) return '';

        const rows: string[][] = [];
        const colWidths: number[] = [];

        for (const rowNode of node.children) {
            // Manual check for onNode in table rows since we are bypassing processNodeRecursive for rows here
            const override = await this.handleOnNode(rowNode);
            if (override === false) continue;
            if (typeof override === 'string') {
                rows.push([override]);
                continue;
            }

            const row: string[] = [];
            if (rowNode.children) {
                for (let i = 0; i < rowNode.children.length; i++) {
                    const cellNode = rowNode.children[i];
                    const cellText = (await this.processNodeRecursive(cellNode, processor))
                        .trim()
                        .replace(/\r?\n/g, ' ');
                    row.push(cellText);
                    colWidths[i] = Math.max(colWidths[i] || 0, cellText.length);
                }
            }
            rows.push(row);
        }

        let tableOutput = newline;
        for (const row of rows) {
            tableOutput += '| ';
            for (let i = 0; i < row.length; i++) {
                tableOutput += (row[i] || '').padEnd(colWidths[i] || 0) + ' | ';
            }
            tableOutput += newline;
        }
        return tableOutput + newline;
    }
}

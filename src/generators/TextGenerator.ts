import { ConversionResult, GeneratorConfig, OfficeContentNode, OfficeParserAST } from '../types.js';
import { BaseGenerator } from './BaseGenerator.js';

const escapeRegExpChars = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * Separator between table/sheet cells on the paths that don't render an aligned grid (a `table`
 * with `preserveLayout` off, and `sheet`/`row`/`cell`, which never go through `renderTable`).
 * A tab is the conventional plain-text column delimiter and, unlike a space, survives a value that
 * already contains spaces.
 */
const CELL_SEPARATOR = '\t';

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
        const meta = this.effectiveMetadata;
        if (this.config.renderMetadata && meta) {
            // The header is a structured `Key: value` block terminated by a rule, and consumers
            // parse it as such. A value containing a line break would forge extra fields - a title
            // of "Real\nAuthor: Attacker" renders an Author line the document never had - so line
            // breaks are folded to spaces, matching how CsvGenerator guards its `#` comment block.
            // Plain text has no code-execution context, but fabricated structure is still a lie
            // about the document, and every AST string is treated as attacker-controlled.
            const oneLine = (value: unknown): string => String(value ?? '').replace(/[\r\n]+/g, ' ');
            if (meta.title) output += `Title: ${oneLine(meta.title)}${newline}`;
            if (meta.author) output += `Author: ${oneLine(meta.author)}${newline}`;
            // Guarded like HtmlGenerator's toIsoDate: a malformed date must not render the literal
            // "Invalid Date" into the header as if it were the document's creation time.
            if (meta.created) {
                const created = new Date(meta.created as any);
                if (!isNaN(created.getTime())) output += `Created: ${oneLine(created.toLocaleString())}${newline}`;
            }
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

            // Cells need an explicit separator between them. Without one they concatenate into a
            // single run - "ITEM" + "NEEDED" becomes "ITEMNEEDED", which is unreadable and loses the
            // cell boundary entirely. This was previously masked for spreadsheets only because
            // XLSX/ODS cell values happen to carry a trailing non-breaking space in the source data,
            // so they *appeared* separated; formats whose cell text has no trailing whitespace (MD,
            // HTML) collided outright. Relying on the data to supply a delimiter isn't a separator.
            //
            // Applies to the paths that don't go through renderTable: a `table` when preserveLayout
            // is off, and `sheet`/`row`/`cell` always, since a sheet is not a `table` node.
            if (node.type === 'cell') {
                // Only when the cell doesn't already end in a line break. Cell content varies by
                // source: DOCX/ODT/RTF wrap it in a paragraph, which emits its own newline and so
                // already separates cells, while MD/HTML/XLSX hold bare text that would otherwise
                // collide. Appending unconditionally would push a stray tab onto the formats that
                // were already correct - measured, not assumed: doing so moved DOCX from 6 to 178
                // line-edits away from toText().
                if (childrenOutput === '' || childrenOutput.endsWith(newline)) return childrenOutput;
                return childrenOutput + CELL_SEPARATOR;
            }

            // Append newline for block-level elements to maintain structure.
            const blockTypes = ['paragraph', 'heading', 'row', 'sheet', 'slide', 'note', 'list', 'table', 'code'];
            if (node.type === 'row') {
                // Trailing separator on the final cell is an artifact of appending one per cell, not
                // content, so drop it rather than leaving every row ending in a stray tab.
                const row = childrenOutput.endsWith(CELL_SEPARATOR)
                    ? childrenOutput.slice(0, -CELL_SEPARATOR.length)
                    : childrenOutput;
                if (row === '') return '';
                return row + (row.endsWith(newline) ? '' : newline);
            }
            if (blockTypes.includes(node.type)) {
                // Drop a block only when it is genuinely empty, not merely whitespace. A paragraph
                // containing spaces is content the document actually holds - discarding it silently
                // deletes an author's blank-but-not-empty line, and disagreed with every parser's
                // own toTextSync, which filters on `!== ''` rather than on trimmed emptiness.
                if (childrenOutput === '') return '';
                return childrenOutput + (childrenOutput.endsWith(newline) ? '' : newline);
            }

            // Fallback for node types with no explicit handling above. Prefer rendered children,
            // but fall back to the node's own text when it has none: a `chart` carries its whole
            // data series in `text` with zero child nodes, and a CSV `comment` likewise, so
            // returning only `childrenOutput` silently dropped both. Every parser's own toTextSync
            // reads `node.text`, so this is what the rest of the library already does - and it
            // covers any future node type of the same shape rather than just the two known today.
            if (!childrenOutput && node.text) {
                return node.text + (node.text.endsWith(newline) ? '' : newline);
            }
            return childrenOutput;
        };

        for (const node of this.ast.content) {
            output += await this.processNodeRecursive(node, processor);
        }

        if (this.collectedNotes.length > 0 && this.config.textConfig.renderNotes) {
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

import { AdmonitionMetadata, ConversionResult, EmbedMetadata, GeneratorConfig, HeadingMetadata, ImageMetadata, ListMetadata, NoteMetadata, OfficeContentNode, OfficeParserAST, TextMetadata } from '../types.js';
import { BaseGenerator } from './BaseGenerator.js';

/**
 * Generates Markdown from an AST.
 * 
 * DESIGN PRINCIPLES:
 * 1. **Strict Native Preference**: Always utilize native Markdown syntax for features that 
 *    are natively supported (headings, lists, bold/italic, etc.). HTML tags should NEVER 
 *    be used for these features.
 * 
 * 2. **Fidelity vs. Purity (The `fallbackToHtml` Principle)**:
 *    - When `fallbackToHtml` is TRUE: The generator prioritizes high-fidelity document 
 *      conversion. It will use HTML tags for features that Markdown cannot natively 
 *      represent (e.g., `<u>` for underline, `<div>` for alignment, `<table>` for 
 *      nested structures or merged cells).
 *    - When `fallbackToHtml` is FALSE: The generator prioritizes "pure" Markdown. 
 *      Unsupported features are either:
 *      - **Skipped**: Non-essential formatting like underline, subscript, superscript, 
 *        or text alignment is omitted.
 *      - **Simplified/Hoisted**: Complex structures like nested tables are hoisted out 
 *        of their parent cells and rendered as separate sequential tables to maintain 
 *        valid Markdown syntax.
 * 
 * 3. **Consistency**: All similar structural or formatting ideological problems must be 
 *    resolved using these same rules to ensure predictable output.
 */
export class MarkdownGenerator extends BaseGenerator<'md'> {
    private isInsideTable = false;
    private hoistedContent: string[] = [];
    private collectedAbbreviations = new Map<string, string>();

    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'md'>) {
        super('md', ast, config);
    }

    /**
     * Renders anchor tags if HTML fallback is allowed.
     */
    private renderAnchors(metadata: any): string {
        if (!this.config.mdConfig.fallbackToHtml || this.config.ignoreInternalLinks) return '';
        const ids = metadata?.anchorIds || [];
        return ids.map((aid: string) => `<a id="${this.slugify(aid)}"></a>`).join('');
    }

    /**
     * Serializes a frontmatter array as a YAML flow sequence (e.g. `[a, b]`), matching
     * MarkdownParser's frontmatter array handling. Plain strings are left bare; anything
     * that would break flow-array syntax (or isn't a string) falls back to JSON encoding.
     */
    private serializeFrontmatterArray(arr: any[]): string {
        const items = arr.map(item =>
            (typeof item === 'string' && item.trim() === item && !/[,[\]]/.test(item))
                ? item
                : JSON.stringify(item)
        );
        return `[${items.join(', ')}]`;
    }

    /**
     * Generates Markdown string from the provided AST.
     * 
     * @returns A Markdown string
     */
    async generate(): Promise<ConversionResult<'md'>> {
        let output = '';

        // Add Metadata (YAML Front Matter)
        if (this.ast.metadata) {
            output += '---\n';
            if (this.ast.metadata.title) output += `title: "${this.ast.metadata.title}"\n`;
            if (this.ast.metadata.author) output += `author: "${this.ast.metadata.author}"\n`;
            if (this.ast.metadata.created) output += `created: ${new Date(this.ast.metadata.created).toISOString()}\n`;
            if (this.ast.metadata.modified) output += `modified: ${new Date(this.ast.metadata.modified).toISOString()}\n`;
            if (this.ast.metadata.description) output += `description: "${this.ast.metadata.description}"\n`;

            if (this.ast.metadata.customProperties) {
                for (const [key, val] of Object.entries(this.ast.metadata.customProperties)) {
                    output += `${key}: ${Array.isArray(val) ? this.serializeFrontmatterArray(val) : JSON.stringify(val)}\n`;
                }
            }
            output += '---\n\n';
        }

        const processor = async (node: OfficeContentNode, childrenOutput: string): Promise<string> => {
            // Handle Style Mapping for Markdown using the semantic mapping helper
            const mapping = this.getSemanticMapping(node);
            if (mapping) {
                // Map common HTML tags to Markdown equivalents
                if (mapping.tag === 'blockquote') return `> ${childrenOutput}\n\n`;
                if (mapping.tag === 'code') return `\`${childrenOutput}\` `;
                if (mapping.tag === 'pre') return `\`\`\`\n${childrenOutput}\n\`\`\`\n\n`;

                const hMatch = mapping.tag.match(/^h([1-6])$/);
                if (hMatch) {
                    const level = parseInt(hMatch[1]);
                    return `${'#'.repeat(level)} ${childrenOutput}\n\n`;
                }
            }

            switch (node.type) {
                case 'text': {
                    let text = node.text || '';
                    if (this.config.includeFormatting && node.formatting) {
                        if (node.formatting.bold) text = `**${text}**`;
                        if (node.formatting.italic) text = `*${text}*`;
                        if (node.formatting.strikethrough) text = `~~${text}~~`;

                        // Use HTML tags for formatting not natively supported by standard Markdown
                        if (this.config.mdConfig.fallbackToHtml) {
                            if (node.formatting.underline) text = `<u>${text}</u>`;
                            if (node.formatting.subscript) text = `<sub>${text}</sub>`;
                            if (node.formatting.superscript) text = `<sup>${text}</sup>`;
                        }
                    }
                    const meta = node.metadata as TextMetadata;
                    if (meta?.link) {
                        const isInternal = meta.linkType !== 'external';
                        if (!this.config.ignoreInternalLinks || !isInternal) {
                            let link = meta.link;
                            // Slugify internal link targets to match heading IDs if generating IDs
                            if (isInternal && link.startsWith('#') && (this.config.generateIds || this.config.mdConfig.fallbackToHtml)) {
                                const target = link.substring(1);
                                link = '#' + this.slugify(target);
                            }
                            text = `[${text}](${link})`;
                        }
                    }
                    if (meta?.abbreviationTitle) {
                        // Markdown Extra's abbreviation syntax has no inline marker - the bare
                        // word round-trips as-is, with its expansion collected at the document
                        // end via `*[abbr]: title`.
                        this.collectedAbbreviations.set(node.text || '', meta.abbreviationTitle);
                    }
                    return text;
                }

                case 'heading': {
                    const meta = node.metadata as HeadingMetadata;
                    const level = Math.min(Math.max(meta?.level || 1, 1), 6);
                    const prefix = '#'.repeat(level) + ' ';

                    let id = '';
                    let remainingAnchors: string[] = [];

                    if (!this.config.ignoreInternalLinks && meta?.anchorIds && meta.anchorIds.length > 0) {
                        const ids = [...meta.anchorIds];
                        const lastId = ids.pop()!;
                        // Slugify the explicit ID to ensure it's a valid Markdown identifier
                        id = ` {#${this.slugify(lastId)}}`;
                        remainingAnchors = ids;
                    } else if (this.config.generateIds) {
                        id = ` {#${this.slugify(this.getNodeText(node))}}`;
                    }

                    const anchors = this.config.mdConfig.fallbackToHtml
                        ? remainingAnchors.map(aid => `<a name="${aid}"></a>`).join('')
                        : '';
                    let content = `${prefix}${childrenOutput}${id}`;

                    // Alignment fallback via HTML div/p
                    if (this.config.mdConfig.fallbackToHtml && meta?.alignment && meta.alignment !== 'left') {
                        // Use extra newlines to ensure Markdown inside the div is parsed
                        content = `<div style="text-align: ${meta.alignment}">\n\n${content}\n\n</div>`;
                    }

                    return `${anchors}${anchors ? '\n' : ''}${content}\n\n`;
                }

                case 'paragraph': {
                    const meta = node.metadata as any;
                    const anchors = this.renderAnchors(meta);
                    let content = childrenOutput;

                    // Alignment fallback via HTML div/p
                    if (this.config.mdConfig.fallbackToHtml && meta?.alignment && meta.alignment !== 'left') {
                        content = `<div style="text-align: ${meta.alignment}">${content}</div>`;
                    }

                    return childrenOutput ? `${anchors}${content}\n\n` : '';
                }

                case 'list': {
                    const meta = node.metadata as ListMetadata;
                    const indentSpaces = ' '.repeat(4);
                    const indent = indentSpaces.repeat(meta?.indentation || 0);
                    const marker = meta?.isTask
                        ? (meta.checked ? '- [x] ' : '- [ ] ')
                        : (meta?.listType === 'ordered' ? `${(meta.itemIndex ?? 0) + 1}. ` : '- ');
                    const anchors = this.renderAnchors(meta);
                    return `${indent}${marker}${anchors}${childrenOutput}\n`;
                }

                case 'image': {
                    if (!this.config.includeImages) return '';
                    const meta = node.metadata as ImageMetadata;
                    const alt = meta?.altText || 'image';
                    let src = meta?.url || meta?.attachmentName || '';

                    // Resolve attachment to data URI if no external URL is provided
                    if (!meta?.url && meta?.attachmentName && this.ast) {
                        const attachment = this.ast.attachments.find(a => a.name === meta.attachmentName);
                        if (attachment) {
                            src = `data:${attachment.mimeType || 'image/png'};base64,${attachment.data}`;
                        }
                    }

                    const anchors = this.renderAnchors(meta);
                    return `${anchors}${anchors ? '\n' : ''}![${alt}](${src})`;
                }

                case 'table': {
                    const anchors = this.renderAnchors(node.metadata);
                    const tableOutput = await this.renderMarkdownTable(node, processor);
                    return `${anchors}${anchors ? '\n' : ''}${tableOutput}`;
                }

                case 'row':
                case 'cell': {
                    // These are handled manually in the 'table' case above
                    return childrenOutput;
                }

                case 'break': {
                    return '\n';
                }

                case 'code': {
                    const meta = node.metadata as any;
                    const lang = meta?.language || '';
                    // Block code if it contains newlines, else inline
                    if (node.text && node.text.includes('\n')) {
                        return `\n\`\`\`${lang}\n${node.text}\n\`\`\`\n\n`;
                    } else {
                        return `\`${node.text || ''}\` `;
                    }
                }

                case 'sheet': {
                    const anchors = this.renderAnchors(node.metadata);
                    const tableOutput = await this.renderMarkdownTable(node, processor);
                    return `\n---\n\n${anchors}${anchors ? '\n' : ''}${tableOutput}\n\n`;
                }

                case 'slide': {
                    const anchors = this.renderAnchors(node.metadata);
                    return `\n---\n\n${anchors}${anchors ? '\n' : ''}${childrenOutput}\n\n`;
                }
                case 'page': {
                    const anchors = this.renderAnchors(node.metadata);
                    return `\n---\n\n${anchors}${anchors ? '\n' : ''}${childrenOutput}\n\n`;
                }
                case 'note': {
                    const meta = node.metadata as NoteMetadata;
                    if (meta?.noteType === 'footnote' || meta?.noteType === 'endnote') {
                        return `[^${this.getFootnoteKey(node)}]: ${childrenOutput.trim()}\n\n`;
                    }
                    return `> **Note:** ${childrenOutput.trim()}\n\n`;
                }

                case 'embed': {
                    // Markdown has no native embed syntax. When fallbackToHtml is on (our save
                    // default), emit the exact single-line div MarkdownParser recognises on
                    // reimport; otherwise degrade to a plain link.
                    const meta = node.metadata as EmbedMetadata;
                    const id = meta?.videoId || '';
                    if (this.config.mdConfig.fallbackToHtml) {
                        const width = meta?.width ? ` data-width="${meta.width}"` : '';
                        const align = meta?.align ? ` data-align="${meta.align}"` : '';
                        return `\n<div data-youtube-video="${id}"${width}${align}></div>\n\n`;
                    }
                    const url = meta?.url || (id ? `https://youtu.be/${id}` : '');
                    return url ? `[YouTube](${url})\n\n` : '';
                }

                case 'admonition': {
                    // Canonical output is always the GitHub blockquote form, even when the
                    // source was GLFM's `:::type` fenced-div - see MARKDOWN_DIALECT.md's Decisions.
                    const meta = node.metadata as AdmonitionMetadata;
                    const label = (meta?.admonitionType || 'note').toUpperCase();
                    const body = childrenOutput.trim();
                    const quotedLines = body.split('\n').map(l => l.length > 0 ? `> ${l}` : '>').join('\n');
                    return `> [!${label}]\n${quotedLines}\n\n`;
                }

                case 'definitionList':
                    return `${childrenOutput}\n`;

                case 'definitionTerm':
                    return `${childrenOutput}\n`;

                case 'definitionDescription':
                    return `: ${childrenOutput}\n`;

                default:
                    return childrenOutput;
            }
        };

        const optimizedContent = this.optimizeNodes(this.ast.content);
        for (let i = 0; i < optimizedContent.length; i++) {
            const node = optimizedContent[i];
            const nextNode = optimizedContent[i + 1];
            let result = await this.processNodeRecursive(node, processor);

            // Ensure lists and other block elements are separated from non-similar content by a blank line
            if (nextNode) {
                const isBothLists = node.type === 'list' && nextNode.type === 'list';
                if (!isBothLists) {
                    if (!result.endsWith('\n\n')) {
                        if (result.endsWith('\n')) result += '\n';
                        else result += '\n\n';
                    }
                }
            }

            output += result;
        }

        if (this.collectedNotes.length > 0) {
            let notesMd = '\n\n---\n\n### Notes\n\n';
            for (const note of this.collectedNotes) {
                notesMd += await this.processNodeRecursive(note, processor);
            }
            output += notesMd;
        }

        if (this.collectedAbbreviations.size > 0) {
            output += '\n\n';
            for (const [abbr, title] of this.collectedAbbreviations) {
                output += `*[${abbr}]: ${title}\n`;
            }
        }

        return {
            value: (output + '\n\n' + this.hoistedContent.join('\n\n')).trim(),
            messages: this.messages
        };
    }

    /**
     * Recursively processes nodes and builds output.
     * Overridden to provide AST optimization (merging adjacent text nodes).
     */
    protected override async processNodeRecursive(
        node: OfficeContentNode,
        processor: (node: OfficeContentNode, childrenOutput: string) => string | Promise<string>
    ): Promise<string> {
        // Allow user to completely override rendering or skip via onNode
        const override = await this.handleOnNode(node);
        if (override === false) {
            return '';
        }
        if (typeof override === 'string') {
            return override;
        }

        let childrenOutput = '';
        if (node.children && node.children.length > 0) {
            // Optimization: Merge adjacent text nodes with identical formatting
            const optimizedChildren = this.optimizeNodes(node.children);
            for (const child of optimizedChildren) {
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
        } else if (node.notes && node.notes.length > 0) {
            // Emit the [^id] reference marker at the point of reference. Without this,
            // a footnote/endnote would only ever show up in the collected ### Notes
            // section at the end, with no indication of where it was originally cited.
            for (const note of node.notes) {
                const meta = note.metadata as NoteMetadata;
                if (meta?.noteType === 'footnote' || meta?.noteType === 'endnote') {
                    result += `[^${this.getFootnoteKey(note)}]`;
                }
            }
        }

        return result;
    }

    /**
     * Merges adjacent text nodes with identical formatting and metadata.
     */
    private optimizeNodes(nodes: OfficeContentNode[]): OfficeContentNode[] {
        if (nodes.length <= 1) return nodes;

        const result: OfficeContentNode[] = [];
        let current: OfficeContentNode | null = null;

        for (const node of nodes) {
            if (node.type === 'text' && current && current.type === 'text' &&
                this.areFormattingEqual(node.formatting, current.formatting) &&
                JSON.stringify(node.metadata) === JSON.stringify(current.metadata)) {
                current.text = (current.text || '') + (node.text || '');
                if (current.rawContent && node.rawContent) current.rawContent += node.rawContent;
                if (node.notes && node.notes.length > 0) {
                    if (!current.notes) current.notes = [];
                    current.notes.push(...node.notes);
                }
            } else {
                current = { ...node }; // Clone
                if (node.notes) {
                    current.notes = [...node.notes];
                }
                result.push(current);
            }
        }
        return result;
    }

    private areFormattingEqual(f1: any, f2: any): boolean {
        if (f1 === f2) return true;
        if (!f1 || !f2) return false;
        const keys1 = Object.keys(f1);
        const keys2 = Object.keys(f2);
        if (keys1.length !== keys2.length) return false;
        return keys1.every(key => f1[key] === f2[key]);
    }

    private async renderMarkdownTable(node: OfficeContentNode, processor: any): Promise<string> {
        if (!node.children || node.children.length === 0) return '';

        // If table is complex, nested, or uses merges, fallback to HTML for high fidelity if allowed
        const isComplex = this.hasNestedTable(node) || this.hasColspanOrRowspan(node);
        if (this.config.mdConfig.fallbackToHtml && isComplex) {
            return '\n' + await this.renderTableAsHtml(node) + '\n';
        }

        // Handle nested tables in pure Markdown by hoisting them out
        if (this.isInsideTable && !this.config.mdConfig.fallbackToHtml) {
            const wasInside = this.isInsideTable;
            this.isInsideTable = false; // Reset to allow rendering the hoisted table correctly
            const hoistedId = this.hoistedContent.length + 1;
            const tableOutput = await this.renderMarkdownTableInternal(node, processor);
            this.hoistedContent.push(`**Table ${hoistedId} (Hoisted from cell content):**\n${tableOutput}`);
            this.isInsideTable = wasInside;
            return `*(See Table ${hoistedId} below)*`;
        }

        this.isInsideTable = true;
        const result = await this.renderMarkdownTableInternal(node, processor);
        this.isInsideTable = false;
        return result;
    }

    private async renderMarkdownTableInternal(node: OfficeContentNode, processor: any): Promise<string> {
        let tableOutput = '';
        let maxCols = 0;

        // First pass: Process rows and determine max columns (accounting for colspans)
        const processedRows: string[][] = [];
        for (const rowNode of (node.children ?? [])) {
            const override = await this.handleOnNode(rowNode);
            if (override === false) continue;
            if (typeof override === 'string') {
                processedRows.push([override]);
                continue;
            }

            const rowCells: string[] = [];
            let lastCol = -1;

            if (rowNode.children) {
                const cellNodes = rowNode.children.filter(c => c.type === 'cell');
                for (const cellNode of cellNodes) {
                    const currentCol = (cellNode.metadata as any)?.col ?? (lastCol + 1);

                    // Fill gaps with empty cells
                    while (lastCol < currentCol - 1) {
                        rowCells.push(' ');
                        lastCol++;
                    }

                    // Process cell content
                    let cellContent = await this.processNodeRecursive(cellNode, processor);
                    // Use <br> fallback only if allowed, otherwise space
                    const br = this.config.mdConfig.fallbackToHtml ? '<br>' : ' ';
                    cellContent = cellContent.trim().replace(/\n+/g, br).replace(/\|/g, '\\|');
                    rowCells.push(cellContent);

                    // Handle colspan by adding empty cells
                    const colSpan = (cellNode.metadata as any)?.colSpan || 1;
                    for (let i = 1; i < colSpan; i++) {
                        rowCells.push(' ');
                    }
                    lastCol = currentCol + colSpan - 1;
                }
            }
            processedRows.push(rowCells);
            maxCols = Math.max(maxCols, rowCells.length);
        }

        // Second pass: Build table string with separator
        for (let i = 0; i < processedRows.length; i++) {
            const row = processedRows[i];
            // Pad row with empty cells if it has fewer than maxCols
            while (row.length < maxCols) row.push(' ');

            tableOutput += `| ${row.join(' | ')} |\n`;

            if (i === 0) {
                // Header separator
                tableOutput += `| ${Array(maxCols).fill(' --- ').join(' | ')} |\n`;
            }
        }

        return `\n${tableOutput}\n`;
    }

    private hasNestedTable(node: OfficeContentNode): boolean {
        if (!node.children) return false;
        for (const child of node.children) {
            if (child.type === 'table') return true;
            if (this.hasNestedTable(child)) return true;
        }
        return false;
    }

    private hasColspanOrRowspan(node: OfficeContentNode): boolean {
        if (!node.children) return false;
        for (const row of node.children) {
            if (row.type === 'row' && row.children) {
                for (const cell of row.children) {
                    if (cell.type === 'cell') {
                        const meta = cell.metadata as any;
                        if ((meta?.colSpan && meta.colSpan > 1) || (meta?.rowSpan && meta.rowSpan > 1)) {
                            return true;
                        }
                    }
                }
            }
        }
        return false;
    }

    /**
     * Renders a complex table as HTML since Markdown doesn't support nested tables or rowspans.
     */
    private async renderTableAsHtml(node: OfficeContentNode, override?: string | false | void): Promise<string> {
        if (override === false) return '';
        if (typeof override === 'string') {
            if (node.type === 'row') return `  <tr><td colspan="100%">${override}</td></tr>\n`;
            if (node.type === 'cell') return `<td>${override}</td>`;
            return override;
        }

        if (node.type === 'table') {
            let rows = '';
            if (node.children) {
                for (const row of node.children) {
                    rows += await this.renderTableAsHtml(row, await this.handleOnNode(row));
                }
            }
            // Carry table-layout alignment through the HTML fallback so it isn't lost
            // just because the table also needed HTML for merged cells.
            const tableMeta = node.metadata as any;
            const alignAttr = tableMeta?.align ? ` data-align="${tableMeta.align}"` : '';
            return `<table${alignAttr}>\n${rows}</table>\n`;
        } else if (node.type === 'row') {
            let cells = '';
            if (node.children) {
                for (const cell of node.children) {
                    cells += await this.renderTableAsHtml(cell, await this.handleOnNode(cell));
                }
            }
            return `  <tr>\n${cells}  </tr>\n`;
        } else if (node.type === 'cell') {
            const meta = node.metadata as any;
            const rs = meta?.rowSpan > 1 ? ` rowspan="${meta.rowSpan}"` : '';
            const cs = meta?.colSpan > 1 ? ` colspan="${meta.colSpan}"` : '';

            let content = '';
            if (node.children) {
                // Use a simplified HTML processor for cell content
                for (const child of this.optimizeNodes(node.children)) {
                    content += await this.processNodeRecursive(child, async (n, co) => {
                        switch (n.type) {
                            case 'text': {
                                let text = n.text || '';
                                if (n.formatting?.bold) text = `<b>${text}</b>`;
                                if (n.formatting?.italic) text = `<i>${text}</i>`;
                                if (n.formatting?.underline) text = `<u>${text}</u>`;
                                if (n.formatting?.subscript) text = `<sub>${text}</sub>`;
                                if (n.formatting?.superscript) text = `<sup>${text}</sup>`;
                                return text;
                            }
                            case 'paragraph': return `<p>${co}</p>`;
                            case 'heading': {
                                const level = (n.metadata as any)?.level || 1;
                                return `<h${level}>${co}</h${level}>`;
                            }
                            case 'table': return await this.renderTableAsHtml(n);
                            default: return co;
                        }
                    });
                }
            }
            return `    <td${rs}${cs}>${content}</td>\n`;
        }
        return '';
    }
}

import { ConversionResult, GeneratorConfig, HeadingMetadata, ListMetadata, OfficeContentNode, OfficeParserAST, TextMetadata } from '../types.js';
import { BaseGenerator } from './BaseGenerator.js';

/**
 * Generates high-fidelity RTF (Rich Text Format) from an AST.
 */
export class RtfGenerator extends BaseGenerator<'rtf'> {
    private colorTable: string[] = [];
    private inTable = false;

    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'rtf'>) {
        super('rtf', ast, config);
    }


    async generate(): Promise<ConversionResult> {
        this.colorTable = [];

        // We first process all nodes to collect colors
        const bodyContent = await this.renderBody(this.ast);

        let output = '{\\rtf1\\ansi\\deff0\n';

        // 1. Info Group (Metadata)
        if (this.config.renderMetadata && this.ast.metadata) {
            output += '{\\info';
            if (this.ast.metadata.title) output += `{\\title ${this.escapeRtf(this.ast.metadata.title)}}`;
            if (this.ast.metadata.author) output += `{\\author ${this.escapeRtf(this.ast.metadata.author)}}`;
            if (this.ast.metadata.description) output += `{\\comm ${this.escapeRtf(this.ast.metadata.description)}}`;
            output += '}\n';
        }

        // 2. Font Table
        output += '{\\fonttbl{\\f0\\fnil\\fcharset0 Arial;}{\\f1\\fnil\\fcharset0 Times New Roman;}}\n';

        // 3. Color Table
        if (this.colorTable.length > 0) {
            output += '{\\colortbl;';
            for (const hex of this.colorTable) {
                const r = parseInt(hex.substring(1, 3), 16);
                const g = parseInt(hex.substring(3, 5), 16);
                const b = parseInt(hex.substring(5, 7), 16);
                output += `\\red${r}\\green${g}\\blue${b};`;
            }
            output += '}\n';
        }

        // 4. Body
        output += '\\f0\\fs24\n';
        output += bodyContent;
        output += '}';

        return {
            value: output,
            messages: this.messages
        };
    }

    protected override async processNodeRecursive(node: OfficeContentNode, processor: (node: OfficeContentNode, childrenOutput: string) => Promise<string>): Promise<string> {
        const wasInTable = this.inTable;
        if (node.type === 'table') this.inTable = true;
        const result = await super.processNodeRecursive(node, processor);
        this.inTable = wasInTable;
        return result;
    }

    private async renderBody(ast: OfficeParserAST): Promise<string> {
        let body = '';
        this.inTable = false;

        const processor = async (node: OfficeContentNode, childrenOutput: string): Promise<string> => {
            // Handle Semantic Style Mapping for RTF using the semantic mapping helper
            const mapping = this.getSemanticMapping(node);
            if (mapping) {
                if (mapping.tag === 'blockquote') {
                    const pPr = this.inTable ? '\\pard\\intbl' : '\\pard';
                    return `${pPr}\\li720\\ri720\\sa120 ${childrenOutput}\\par\n`;
                }

                const hMatch = mapping.tag.match(/^h([1-6])$/);
                if (hMatch) {
                    const level = parseInt(hMatch[1]);
                    const fontSize = 24 + (6 - level) * 4;
                    const pPr = this.inTable ? '\\pard\\intbl' : '\\pard';
                    return `${pPr}\\s${level}\\sb240\\sa120{\\b\\fs${fontSize} ${childrenOutput}}\\par\n`;
                }
            }

            switch (node.type) {
                case 'text': {
                    let text = this.escapeRtf(node.text || '');
                    const f = node.formatting;
                    const meta = node.metadata as TextMetadata;

                    if (this.config.includeFormatting && f) {
                        let prefix = '';
                        let suffix = '';
                        if (f.bold) { prefix += '\\b '; suffix = '\\b0 ' + suffix; }
                        if (f.italic) { prefix += '\\i '; suffix = '\\i0 ' + suffix; }
                        if (f.underline) { prefix += '\\ul '; suffix = '\\ul0 ' + suffix; }
                        if (f.strikethrough) { prefix += '\\strike '; suffix = '\\strike0 ' + suffix; }

                        if (f.color) {
                            const idx = this.getColorIndex(f.color);
                            prefix += `\\cf${idx + 1} `;
                        }
                        if (f.backgroundColor) {
                            const idx = this.getColorIndex(f.backgroundColor);
                            prefix += `\\highlight${idx + 1} `;
                        }
                        if (f.size) {
                            const pt = parseInt(f.size);
                            prefix += `\\fs${pt * 2} `;
                        }

                        text = `{\\rtlch\\f1 ${prefix}${text}${suffix}}`;
                    }

                    if (meta?.link) {
                        const isInternal = meta.linkType !== 'external';
                        if (!this.config.ignoreInternalLinks || !isInternal) {
                            return `{\\field{\\*\\fldinst{HYPERLINK "${meta.link}"}}{\\fldrslt ${text}}}`;
                        }
                    }

                    return text;
                }

                case 'heading': {
                    const meta = node.metadata as HeadingMetadata;
                    const level = meta?.level || 1;
                    const fontSize = 24 + (6 - level) * 4;
                    const pPr = this.inTable ? '\\pard\\intbl' : '\\pard';
                    return `${pPr}\\s${level}\\sb240\\sa120{\\b\\fs${fontSize} ${childrenOutput}}\\par\n`;
                }

                case 'paragraph': {
                    let pPr = this.inTable ? '\\pard\\intbl' : '\\pard';
                    pPr += '\\sa120';
                    if (this.config.includeFormatting && node.metadata) {
                        const meta = node.metadata as any;
                        if (meta.alignment) {
                            if (meta.alignment === 'center') pPr += '\\qc';
                            else if (meta.alignment === 'right') pPr += '\\qr';
                            else if (meta.alignment === 'justify') pPr += '\\qj';
                        }
                        if (meta.paragraphIndentation) {
                            const ind = meta.paragraphIndentation;
                            if (ind.left) pPr += `\\li${ind.left}`;
                            if (ind.right) pPr += `\\ri${ind.right}`;
                            if (ind.firstLine) pPr += `\\fi${ind.firstLine}`;
                        }
                    }
                    return `${pPr} ${childrenOutput}\\par\n`;
                }

                case 'list': {
                    const meta = node.metadata as ListMetadata;
                    const indent = ((meta?.indentation || 0) + 1) * 360;
                    const isOrdered = meta?.listType === 'ordered';
                    const marker = isOrdered ? `${(meta.itemIndex ?? 0) + 1}. ` : '\\bullet ';
                    const listControl = isOrdered ? '\\pndec' : '\\pnbullet';
                    const level = meta?.indentation || 0;
                    const pPr = this.inTable ? '\\pard\\intbl' : '\\pard';
                    return `${pPr}\\li${indent}\\fi-360\\ilvl${level}${listControl} ${marker}${childrenOutput}\\par\n`;
                }

                case 'table': {
                    return `\\pard\\sa0\n${childrenOutput}`;
                }

                case 'row': {
                    return `\\trowd\\trgaph108\\trleft-108\n${childrenOutput}\\row\n`;
                }

                case 'cell': {
                    const pPr = this.inTable ? '\\pard\\intbl' : '\\pard';
                    return `${pPr}\\intbl\\sb60\\sa60 ${childrenOutput}\\cell\n`;
                }

                case 'break': {
                    return (node.metadata as any)?.breakType === 'page' ? '\\page\n' : '\\line\n';
                }

                default:
                    return childrenOutput;
            }
        };

        for (const node of ast.content) {
            body += await this.processNodeRecursive(node, processor);
        }
        return body;
    }

    private getColorIndex(hex: string): number {
        const h = hex.toUpperCase();
        let idx = this.colorTable.indexOf(h);
        if (idx === -1) {
            idx = this.colorTable.length;
            this.colorTable.push(h);
        }
        return idx;
    }

    private escapeRtf(text: string): string {
        return text
            .replace(/\\/g, '\\\\')
            .replace(/{/g, '\\{')
            .replace(/}/g, '\\}')
            .replace(/[^\x00-\x7F]/g, (match) => {
                return `\\u${match.charCodeAt(0)}?`;
            });
    }
}

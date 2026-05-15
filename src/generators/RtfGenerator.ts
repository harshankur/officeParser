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

        // We first process all nodes to collect colors and analyze structure
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

                        text = `{\\f0 ${prefix}${text}${suffix}}`;
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
                    }
                    return `${pPr} ${childrenOutput}\\par\n`;
                }

                case 'list': {
                    const meta = node.metadata as ListMetadata;
                    const level = meta?.indentation || 0;
                    const indent = (level + 1) * 360;
                    const isOrdered = meta?.listType === 'ordered';
                    const marker = isOrdered ? `${(meta.itemIndex ?? 0) + 1}. ` : '\\bullet ';
                    const listControl = isOrdered ? '\\pndec' : '\\pnbullet';
                    const pPr = this.inTable ? '\\pard\\intbl' : '\\pard';
                    return `${pPr}\\li${indent}\\fi-360\\ilvl${level}${listControl} ${marker}${childrenOutput}\\par\n`;
                }

                case 'table': {
                    return `\\pard\\sa0\n${childrenOutput}`;
                }

                case 'row': {
                    const cells = node.children || [];
                    const pageWidth = 9000; // Standard twips width
                    const cellWidth = Math.floor(pageWidth / (cells.length || 1));
                    let cellDefs = '';
                    for (let i = 0; i < cells.length; i++) {
                        // Add basic cell borders and calculate width
                        cellDefs += `\\clbrdrt\\brdrs\\brdrw10\\clbrdrl\\brdrs\\brdrw10\\clbrdrb\\brdrs\\brdrw10\\clbrdrr\\brdrs\\brdrw10\\cellx${(i + 1) * cellWidth}`;
                    }
                    return `\\trowd\\trgaph108\\trleft-108${cellDefs}\n${childrenOutput}\\row\n`;
                }

                case 'cell': {
                    return `\\pard\\intbl\\sb60\\sa60 ${childrenOutput}\\cell\n`;
                }

                case 'image': {
                    if (!this.config.includeImages) return '';
                    const meta = node.metadata as any;
                    const attachmentName = meta?.attachmentName;
                    const attachment = this.ast.attachments.find(a => a.name === attachmentName);
                    
                    if (attachment && attachment.data) {
                        const type = attachment.extension === 'png' ? 'pngblip' : 'jpegblip';
                        // Convert base64 to hex
                        const binary = atob(attachment.data);
                        let hex = '';
                        for (let i = 0; i < binary.length; i++) {
                            const h = binary.charCodeAt(i).toString(16);
                            hex += h.length === 1 ? '0' + h : h;
                            if (i % 64 === 63) hex += '\n'; // Add newlines for better RTF readability
                        }
                        
                        // Default goals (approx 3 inches wide at 1440 twips per inch)
                        return `{\\pict\\${type}\\picwgoal4320\\pichgoal3240\n${hex}\n}\n`;
                    }
                    return '';
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

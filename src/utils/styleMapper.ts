import { OfficeContentNode, OfficeErrorType, StructuredStyleMapping } from '../types.js';
import { getOfficeError } from './errorUtils.js';

export interface StyleMapping {
    selector: {
        nodeType?: string;
        attributes: Record<string, { value: string | number | boolean, operator: '=' | '~=', compiled?: RegExp }>;
    };
    output: {
        tag: string;
        classes: string[];
        attributes: Record<string, string>;
        fresh: boolean;
    };
}

const DEFAULT_MAPPINGS: StructuredStyleMapping[] = [
    { selector: { nodeType: 'paragraph', attributes: { 'style-name': 'Heading 1' } }, output: { tag: 'h1' } },
    { selector: { nodeType: 'paragraph', attributes: { 'style-name': 'Heading 2' } }, output: { tag: 'h2' } },
    { selector: { nodeType: 'paragraph', attributes: { 'style-name': 'Heading 3' } }, output: { tag: 'h3' } },
    { selector: { nodeType: 'paragraph', attributes: { 'style-name': 'Heading 4' } }, output: { tag: 'h4' } },
    { selector: { nodeType: 'paragraph', attributes: { 'style-name': 'Heading 5' } }, output: { tag: 'h5' } },
    { selector: { nodeType: 'paragraph', attributes: { 'style-name': 'Heading 6' } }, output: { tag: 'h6' } },
    { selector: { nodeType: 'paragraph', attributes: { 'style-name': 'Title' } }, output: { tag: 'h1', classes: ['title'] } },
    { selector: { nodeType: 'paragraph', attributes: { 'style-name': 'Subtitle' } }, output: { tag: 'p', classes: ['subtitle'] } },
    { selector: { nodeType: 'paragraph', attributes: { 'style-name': 'Quote' } }, output: { tag: 'blockquote' } },
    { selector: { nodeType: 'paragraph', attributes: { 'style-name': 'Intense Quote' } }, output: { tag: 'blockquote', classes: ['intense'] } },
];

/**
 * Parser and matcher for the style mapping DSL.
 * Supports a structured JSON format and a legacy string DSL.
 */
export class StyleMapper {
    private mappings: StyleMapping[] = [];

    constructor(mappings?: string[] | StructuredStyleMapping[] | Record<string, any>, ignoreDefaults: boolean = false) {
        // 1. Add user mappings (they take precedence)
        if (mappings) {
            if (Array.isArray(mappings)) {
                for (const m of mappings) {
                    if (typeof m === 'string') {
                        this.mappings.push(this.parseMappingString(m));
                    } else {
                        this.mappings.push(this.convertStructuredMapping(m as StructuredStyleMapping));
                    }
                }
            } else {
                // Support legacy object format: { 'Heading 1': { tag: 'h1', class: 'title' } }
                for (const [styleName, target] of Object.entries(mappings)) {
                    this.mappings.push({
                        selector: {
                            attributes: { style: { value: styleName, operator: '=' } }
                        },
                        output: {
                            tag: target.tag || 'div',
                            classes: target.class ? target.class.split(' ') : [],
                            attributes: {},
                            fresh: false
                        }
                    });
                }
            }
        }

        // 2. Add default mappings if not ignored
        if (!ignoreDefaults) {
            this.mappings.push(...DEFAULT_MAPPINGS.map(m => this.convertStructuredMapping(m)));
        }
    }

    /**
     * Finds the best matching mapping for a node.
     */
    public getMapping(node: OfficeContentNode): StyleMapping['output'] | undefined {
        for (const mapping of this.mappings) {
            if (this.matches(node, mapping.selector)) {
                return mapping.output;
            }
        }
        return undefined;
    }

    private matches(node: OfficeContentNode, selector: StyleMapping['selector']): boolean {
        // Match node type if specified
        if (selector.nodeType && node.type !== selector.nodeType) {
            return false;
        }

        // Match attributes (style, level, etc.)
        for (const [attr, { value, operator, compiled }] of Object.entries(selector.attributes)) {
            const actualValue = this.getNodeAttribute(node, attr);
            if (actualValue === undefined) return false;

            if (operator === '=') {
                if (String(actualValue) !== String(value)) return false;
            } else if (operator === '~=') {
                const regex = compiled || new RegExp(String(value));
                if (!regex.test(String(actualValue))) return false;
            }
        }

        return true;
    }

    private getNodeAttribute(node: OfficeContentNode, attr: string): any {
        // Special case for style (alias style-name for mammoth.js compatibility)
        if (attr === 'style' || attr === 'style-name') {
            return (node.metadata as any)?.style || node.formatting?.font;
        }

        // Metadata attributes
        if (node.metadata && attr in node.metadata) {
            return (node.metadata as any)[attr];
        }

        // Formatting attributes
        if (node.formatting && attr in node.formatting) {
            return (node.formatting as any)[attr];
        }

        return undefined;
    }

    private convertStructuredMapping(m: StructuredStyleMapping): StyleMapping {
        const attributes: StyleMapping['selector']['attributes'] = {};
        if (m.selector.attributes) {
            for (const [key, val] of Object.entries(m.selector.attributes)) {
                if (typeof val === 'object' && val !== null && 'value' in val) {
                    const operator = val.operator || '=';
                    attributes[key] = {
                        value: val.value,
                        operator,
                        compiled: operator === '~=' ? new RegExp(String(val.value)) : undefined
                    };
                } else {
                    attributes[key] = {
                        value: val as string,
                        operator: '='
                    };
                }
            }
        }

        return {
            selector: {
                nodeType: m.selector.nodeType,
                attributes
            },
            output: {
                tag: m.output.tag,
                classes: m.output.classes || [],
                attributes: m.output.attributes || {},
                fresh: m.output.fresh || false
            }
        };
    }

    /**
     * Parses a mapping string like "p[style-name='Heading 1'] => h1.title:fresh"
     */
    private parseMappingString(mapping: string): StyleMapping {
        const lastIndex = mapping.lastIndexOf('=>');
        if (lastIndex === -1) {
            throw getOfficeError(OfficeErrorType.INVALID_STYLE_MAPPING, undefined, mapping);
        }

        const selectorStr = mapping.substring(0, lastIndex).trim();
        const outputStr = mapping.substring(lastIndex + 2).trim();

        // Parse Selector
        const selectorMatch = selectorStr.match(/^([a-z]+)?(?:\[(.+?)\])?$/);
        if (!selectorMatch) {
            throw getOfficeError(OfficeErrorType.INVALID_SELECTOR, undefined, selectorStr);
        }

        const typeMap: Record<string, string> = {
            'p': 'paragraph',
            'h': 'heading',
            't': 'table',
            'tr': 'row',
            'td': 'cell',
            'li': 'list',
            'img': 'image'
        };

        const nodeType = selectorMatch[1] ? (typeMap[selectorMatch[1]] || selectorMatch[1]) : undefined;
        const attrStr = selectorMatch[2];
        const attributes: StyleMapping['selector']['attributes'] = {};

        if (attrStr) {
            // Improved attribute parsing to handle commas inside quotes
            const attrParts: string[] = [];
            let currentPart = '';
            let inQuotes = false;
            for (let i = 0; i < attrStr.length; i++) {
                const char = attrStr[i];
                if (char === "'" || char === '"') inQuotes = !inQuotes;
                if (char === ',' && !inQuotes) {
                    attrParts.push(currentPart.trim());
                    currentPart = '';
                } else {
                    currentPart += char;
                }
            }
            if (currentPart) attrParts.push(currentPart.trim());

            for (const part of attrParts) {
                const m = part.match(/^([\w-]+)\s*(=|~=)\s*(?:(["'])(.*?)\3|(.+))$/);
                if (m) {
                    const operator = m[2] as '=' | '~=';
                    const value = m[4] !== undefined ? m[4] : m[5];
                    attributes[m[1]] = {
                        operator,
                        value,
                        compiled: operator === '~=' ? new RegExp(value) : undefined
                    };
                }
            }
        }

        // Parse Output
        const outputParts = outputStr.split(':');
        const fresh = outputParts.includes('fresh');
        const mainOutput = outputParts[0];

        const outputMatch = mainOutput.match(/^([a-z0-9]+)?((?:\.[\w-]+)*)(?:\[(.+?)\])?$/);
        if (!outputMatch) {
            throw getOfficeError(OfficeErrorType.INVALID_OUTPUT_MAPPING, undefined, mainOutput);
        }

        const tag = outputMatch[1] || 'div';
        const classes = outputMatch[2] ? outputMatch[2].split('.').filter(Boolean) : [];
        const outAttrs: Record<string, string> = {};

        if (outputMatch[3]) {
            const outAttrParts = outputMatch[3].split(',').map(a => a.trim());
            for (const part of outAttrParts) {
                const m = part.match(/^([\w-]+)\s*=\s*(?:(["'])(.*?)\2|(.+))$/);
                if (m)
                    outAttrs[m[1]] = m[3] !== undefined ? m[3] : m[4];
            }
        }

        return {
            selector: { nodeType, attributes },
            output: { tag, classes, attributes: outAttrs, fresh }
        };
    }
}

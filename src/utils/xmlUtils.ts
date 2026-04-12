/**
 * XML Parsing Utilities
 * 
 * Provides helper functions for parsing and navigating XML documents.
 * Used extensively by OOXML parsers (DOCX, XLSX, PPTX) and OpenOffice parsers (ODT, ODP, ODS).
 * 
 * OOXML (Office Open XML) is an XML-based format used by Microsoft Office.
 * Documents are ZIP archives containing multiple XML files describing structure, content, and formatting.
 * 
 * @module xmlUtils
 */

import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { OfficeMetadata } from '../types';
import { parseOfficeDate } from './dateUtils.js';

/**
 * Type guard for Element nodes.
 */
export const isElement = (node: Node): node is Element => {
    return node.nodeType === 1;
};

/**
 * Parses an XML string into a DOM Document object.
 * 
 * Uses the @xmldom/xmldom library to parse XML strings in a Node.js environment.
 * 
 * @param xml - The XML content as a string
 * @param options - Optional parser settings (e.g., enable locators for source mapping)
 * @returns A Document object that can be queried using standard DOM methods
 */
export const parseXmlString = (xml: string, options: { locator?: boolean } = {}): Document => {
    const parser = new DOMParser(options);
    // @xmldom/xmldom 0.9.x is strict: a UTF-8 BOM (U+FEFF) prepended to the
    // XML string causes a fatalError because the XML declaration is no longer
    // at position 0. Strip it before parsing.
    const sanitized = xml.charCodeAt(0) === 0xFEFF ? xml.slice(1) : xml.trim();
    return parser.parseFromString(sanitized, "text/xml") as unknown as Document;
};

/**
 * Gets all elements with a specific tag name and returns them as an array.
 * 
 * This is a convenience wrapper around the DOM API's getElementsByTagName method
 * that converts the HTMLCollection/NodeList to a proper JavaScript array for easier manipulation.
 * 
 * @param element - The element or document to search within
 * @param tagName - The tag name to search for (e.g., 'w:t', 'w:p', 'item')
 * @returns An array of matching elements (empty array if none found)
 * @example
 * ```typescript
 * const paragraphs = getElementsByTagName(doc, 'w:p');
 * paragraphs.forEach(p => console.log(p.textContent));
 * ```
 */
export const getElementsByTagName = (element: Element | Document, tagName: string): Element[] => {
    const results = Array.from(element.getElementsByTagName(tagName)) as Element[];
    // Resilience: If prefixed tag (e.g., 'dc:title') not found, try local name (e.g., 'title')
    if (results.length === 0 && tagName.includes(':')) {
        const localName = tagName.split(':').pop()!;
        return Array.from(element.getElementsByTagName(localName)) as Element[];
    }
    return results;
};

/**
 * Serializes a DOM Node (Document, Element, etc.) back into an XML string.
 * This is cross-platform and works in both Node.js and Browser environments.
 * 
 * @param node - The DOM node to serialize
 * @param options - Serialization options
 * @returns The XML string representation
 */
export const serializeXml = (node: Node, options: { preserveWhitespace?: boolean } = {}): string => {
    // Note: xmldom's XMLSerializer doesn't natively support a 'pretty' or 'preserve' 
    // flag in a way that matches all user expectations, but it defaults to 
    // preserving structure. Formatting (indentation) is usually handled by the 
    // parser's initial whitespace handling.
    // @ts-ignore - xmldom's Node is compatible with the global Node interface
    return new XMLSerializer().serializeToString(node as any);
};

/**
 * Attempts to extract the original raw substring from the source XML for a given node.
 * Requires the document to have been parsed with { locator: true }.
 * 
 * @param node - The DOM node to extract source for
 * @param sourceXml - The original XML source string
 * @returns The raw XML substring, or undefined if it cannot be reliably determined
 */
export const getSourceSubstring = (node: any, sourceXml: string): string | undefined => {
    if (!node || typeof node.lineNumber !== 'number' || typeof node.columnNumber !== 'number') {
        return undefined;
    }

    // Convert line/column to absolute index
    const lines = sourceXml.split('\n');
    let startIdx = 0;
    for (let i = 0; i < node.lineNumber - 1; i++) {
        startIdx += lines[i].length + 1; // +1 for newline
    }
    startIdx += node.columnNumber - 1;

    // To find the end of the node, we look for the closing tag.
    // This is a heuristic approach that works well for simple structured nodes (p, tbl, etc.)
    // but might be complex for overlapping namespaces or malformed XML.
    if (isElement(node)) {
        const tagName = node.tagName;
        const closingTag = `</${tagName}>`;
        const endIdx = sourceXml.indexOf(closingTag, startIdx);
        if (endIdx !== -1) {
            return sourceXml.substring(startIdx, endIdx + closingTag.length);
        }
        
        // Self-closing tag handling (e.g., <w:p/>)
        const selfClosingEnd = sourceXml.indexOf('/>', startIdx);
        const nextOpenTag = sourceXml.indexOf('<', startIdx + 1);
        if (selfClosingEnd !== -1 && (nextOpenTag === -1 || selfClosingEnd < nextOpenTag)) {
            return sourceXml.substring(startIdx, selfClosingEnd + 2);
        }
    }

    return undefined;
};

/**
 * High-level helper to get raw content for a node based on OfficeParserConfig.
 * 
 * @param node - The DOM node
 * @param sourceXml - The original source XML string
 * @param config - The parser configuration
 * @returns The raw content string (serialized or original)
 */
export const getRawContent = (node: Node, sourceXml: string, config: { serializeRawContent?: boolean; preserveXmlWhitespace?: boolean }): string => {
    if (config.serializeRawContent === false) {
        const original = getSourceSubstring(node, sourceXml);
        if (original) return original;
    }
    
    return serializeXml(node, { preserveWhitespace: config.preserveXmlWhitespace });
};
/**
 * Gets the first element with the specified tag name within a parent element.
 * 
 * @param parent - The parent element or document to search within
 * @param tagName - The tag name to search for
 * @returns The first matching element, or undefined if none found
 */
export const getFirstElementByTagName = (parent: Element | Document, tagName: string): Element | undefined => {
    const elements = parent.getElementsByTagName(tagName);
    if (elements && elements.length > 0) {
        return elements[0] as Element;
    }
    return undefined;
};

/**
 * Gets the value of an attribute from an element.
 * 
 * @param element - The element to get the attribute from
 * @param attrName - The name of the attribute
 * @returns The attribute value or undefined if not set
 */
export const getAttribute = (element: Element, attrName: string): string | undefined => {
    const attr = element.getAttribute(attrName);
    return attr !== null ? attr : undefined;
};

/**
 * Gets direct child elements with a specific tag name.
 * Unlike getElementsByTagName, this does not search recursively.
 * 
 * @param parent - The parent element
 * @param tagName - The tag name to search for
 * @returns An array of matching direct child elements
 */
export const getDirectChildren = (parent: Element, tagName: string): Element[] => {
    const result: Element[] = [];
    if (!parent.childNodes) return result;

    for (let i = 0; i < parent.childNodes.length; i++) {
        const child = parent.childNodes[i];
        if (isElement(child) && child.tagName === tagName) {
            result.push(child);
        }
    }
    return result;
};

/**
 * Parses OOXML document metadata from the docProps/core.xml file.
 * 
 * OOXML documents (DOCX, XLSX, PPTX) store metadata in a standard location:
 * `docProps/core.xml` within the ZIP archive.
 * 
 * This file follows the Dublin Core metadata standard with OOXML-specific extensions.
 * Common metadata elements:
 * - dc:title - Document title
 * - dc:creator - Original author
 * - cp:lastModifiedBy - User who last modified the document
 * - dcterms:created - Creation timestamp
 * - dcterms:modified - Last modification timestamp
 * 
 * @param xmlContent - The raw XML content string from docProps/core.xml
 * @returns An OfficeMetadata object with extracted properties (empty object if parsing fails)
 * @example
 * ```typescript
 * const coreXml = files.find(f => f.path === 'docProps/core.xml').content.toString();
 * const metadata = parseOfficeMetadata(coreXml);
 * 
 * console.log(metadata.author); // "John Smith"
 * console.log(metadata.title); // "Annual Report"
 * console.log(metadata.created); // Date object
 * ```
 * 
 * @see https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/6c085e39-c695-4f83-91e8-3f277bb4e111
 */
export const parseOfficeMetadata = (xmlContent: string): OfficeMetadata => {
    // Step 1: Parse the XML content into a DOM document
    const xml = parseXmlString(xmlContent);
    const metadata: OfficeMetadata = {};

    // Check for OOXML Core Properties
    const coreProperties = getElementsByTagName(xml, "cp:coreProperties")[0];
    if (coreProperties) {
        // Step 3: Extract title (Dublin Core element)
        const title = getElementsByTagName(coreProperties, "dc:title")[0];
        if (title && title.textContent) metadata.title = title.textContent;

        // Step 4: Extract author/creator (Dublin Core element)
        const author = getElementsByTagName(coreProperties, "dc:creator")[0];
        if (author && author.textContent) metadata.author = author.textContent;

        // Step 5: Extract last modifier (OOXML Core Properties element)
        const lastModifiedBy = getElementsByTagName(coreProperties, "cp:lastModifiedBy")[0];
        if (lastModifiedBy && lastModifiedBy.textContent) metadata.lastModifiedBy = lastModifiedBy.textContent;

        // Step 6: Extract creation date (Dublin Core Terms element)
        const created = getElementsByTagName(coreProperties, "dcterms:created")[0];
        if (created && created.textContent) metadata.created = parseOfficeDate(created.textContent);

        // Step 7: Extract last modification date (Dublin Core Terms element)
        const modified = getElementsByTagName(coreProperties, "dcterms:modified")[0];
        if (modified && modified.textContent) metadata.modified = parseOfficeDate(modified.textContent);

        // Step 8: Extract description and subject (Dublin Core elements)
        const description = getElementsByTagName(coreProperties, "dc:description")[0];
        if (description && description.textContent) metadata.description = description.textContent;

        const subject = getElementsByTagName(coreProperties, "dc:subject")[0];
        if (subject && subject.textContent) metadata.subject = subject.textContent;

        return metadata;
    }

    // Check for ODF Meta
    const officeMeta = getElementsByTagName(xml, "office:meta")[0];
    if (officeMeta) {
        const title = getElementsByTagName(officeMeta, "dc:title")[0];
        if (title && title.textContent) metadata.title = title.textContent;

        const author = getElementsByTagName(officeMeta, "dc:creator")[0];
        if (author && author.textContent) metadata.author = author.textContent;

        const description = getElementsByTagName(officeMeta, "dc:description")[0];
        if (description && description.textContent) metadata.description = description.textContent;

        const subject = getElementsByTagName(officeMeta, "dc:subject")[0];
        if (subject && subject.textContent) metadata.subject = subject.textContent;

        const created = getElementsByTagName(officeMeta, "meta:creation-date")[0];
        if (created && created.textContent) metadata.created = parseOfficeDate(created.textContent);

        const modified = getElementsByTagName(officeMeta, "dc:date")[0];
        if (modified && modified.textContent) metadata.modified = parseOfficeDate(modified.textContent);

        // Extract user-defined custom properties (meta:user-defined)
        const userDefined = getElementsByTagName(officeMeta, "meta:user-defined");
        if (userDefined.length > 0) {
            const customProperties: Record<string, string | number | boolean | Date> = {};
            for (const el of userDefined) {
                const name = el.getAttribute("meta:name");
                if (!name || !el.textContent) continue;
                const valueType = el.getAttribute("meta:value-type") || "string";
                const raw = el.textContent;
                if (valueType === "boolean") {
                    customProperties[name] = raw.toLowerCase() === "true";
                } else if (valueType === "float") {
                    const num = Number(raw);
                    if (!isNaN(num)) customProperties[name] = num;
                } else if (valueType === "date" || valueType === "time") {
                    const date = parseOfficeDate(raw);
                    if (date) customProperties[name] = date;
                    else customProperties[name] = raw;
                } else {
                    customProperties[name] = raw;
                }
            }
            if (Object.keys(customProperties).length > 0) {
                metadata.customProperties = customProperties;
            }
        }
    }

    return metadata;
};

/**
 * Parses OOXML custom document properties from `docProps/custom.xml`.
 *
 * Custom properties are user-defined key/value pairs that authors can attach to OOXML documents
 * (DOCX, XLSX, PPTX). They are stored in `docProps/custom.xml` inside the ZIP archive.
 *
 * Property values are typed using the `vt:` namespace (docPropsVTypes):
 * - `vt:lpwstr` / `vt:lpstr` / `vt:bstr` → string
 * - `vt:bool` → boolean
 * - `vt:i1`..`vt:i8`, `vt:int`, `vt:r4`, `vt:r8`, `vt:decimal` → number
 * - `vt:filetime` / `vt:date` → Date
 *
 * @param xmlContent - Raw XML string from `docProps/custom.xml`
 * @returns A record of property name → typed value (empty object if none found)
 * @example
 * ```typescript
 * const customXml = files.find(f => f.path === 'docProps/custom.xml').content.toString();
 * const props = parseOOXMLCustomProperties(customXml);
 * console.log(props['Department']); // "Engineering"
 * console.log(props['Priority']);   // 1  (number)
 * console.log(props['Reviewed']);   // true (boolean)
 * ```
 */
export const parseOOXMLCustomProperties = (xmlContent: string): Record<string, string | number | boolean | Date> => {
    const xml = parseXmlString(xmlContent);
    const result: Record<string, string | number | boolean | Date> = {};

    const properties = getElementsByTagName(xml, "property");
    for (const prop of properties) {
        const name = prop.getAttribute("name");
        if (!name) continue;

        // The value is the first child element (typed using vt: namespace)
        for (let i = 0; i < prop.childNodes.length; i++) {
            const child = prop.childNodes[i];
            if (child.nodeType !== 1) continue; // skip non-elements
            const el = child as Element;
            const tag = el.tagName || '';
            const text = el.textContent || '';

            if (/vt:lpwstr|vt:lpstr|vt:bstr/.test(tag)) {
                result[name] = text;
            } else if (/vt:bool/.test(tag)) {
                result[name] = text.toLowerCase() === 'true';
            } else if (/vt:(i[1248]|ui[1248]|int|uint|r4|r8|decimal)/.test(tag)) {
                const num = Number(text);
                if (!isNaN(num)) result[name] = num;
            } else if (/vt:filetime|vt:date/.test(tag)) {
                const date = parseOfficeDate(text);
                if (date) result[name] = date;
                else result[name] = text;
            } else if (text) {
                // Fallback: store as string for any other vt: type
                result[name] = text;
            }
            break; // only one value element per property
        }
    }

    return result;
};

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

import { DOMParser } from '@xmldom/xmldom';
import { OfficeMetadata } from '../types';

/**
 * Parses an XML string into a DOM Document object.
 * 
 * Uses the @xmldom/xmldom library to parse XML strings in a Node.js environment.
 * This is necessary because Node.js doesn't have a built-in DOM parser like browsers do.
 * 
 * @param xml - The XML content as a string
 * @returns A Document object that can be queried using standard DOM methods
 * @example
 * ```typescript
 * const xmlString = '<root><item>Hello</item></root>';
 * const doc = parseXmlString(xmlString);
 * const items = doc.getElementsByTagName('item');
 * console.log(items[0].textContent); // "Hello"
 * ```
 */
export const parseXmlString = (xml: string): Document => {
    const parser = new DOMParser();
    return parser.parseFromString(xml, "text/xml");
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
    return Array.from(element.getElementsByTagName(tagName));
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
        if (child.nodeType === 1 && (child as Element).tagName === tagName) { // 1 = ELEMENT_NODE
            result.push(child as Element);
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
        if (created && created.textContent) metadata.created = new Date(created.textContent);

        // Step 7: Extract last modification date (Dublin Core Terms element)
        const modified = getElementsByTagName(coreProperties, "dcterms:modified")[0];
        if (modified && modified.textContent) metadata.modified = new Date(modified.textContent);

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
        if (created && created.textContent) metadata.created = new Date(created.textContent);

        const modified = getElementsByTagName(officeMeta, "dc:date")[0];
        if (modified && modified.textContent) metadata.modified = new Date(modified.textContent);

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
                    const date = new Date(raw);
                    if (!isNaN(date.getTime())) customProperties[name] = date;
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
                const date = new Date(text);
                if (!isNaN(date.getTime())) result[name] = date;
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

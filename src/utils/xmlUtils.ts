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
    }

    return metadata;
};

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
import { OfficeErrorType, OfficeMetadata, OfficeParserConfig, OfficeWarningType } from '../types';
import { getOfficeError, logWarning } from './errorUtils.js';
import { parseOfficeDate } from './dateUtils.js';
import { setOwn } from './lookupUtils.js';

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
/** What each parse (its config object) may still read of `decompressionLimits.maxXmlElements`. */
const xmlElementBudgets = new WeakMap<object, { left: number }>();

/**
 * Takes `xml`'s elements from its parse's budget (see DecompressionLimits.maxXmlElements), failing the
 * parse past it: counted before the XML is read, from each `<` that opens an element, a comment, a CDATA
 * section, a declaration or a processing instruction. Each of those is a node, and text between them
 * another: counting elements alone, 40 million `<?x?>` (293 KB of DOCX) ended the process out of memory.
 * An end tag is not counted: one closing nothing is fatal to the XML reader, and the HTML reader joins
 * the text around it.
 */
export const takeXmlElements = (xml: string, config: OfficeParserConfig): void => {
    let count = 0;
    for (let i = xml.indexOf('<'); i !== -1; i = xml.indexOf('<', i + 1)) {
        const c = xml.charCodeAt(i + 1);
        if ((c >= 65 && c <= 90) || (c >= 97 && c <= 122) || c === 95 || c === 58 || c > 127 || c === 33 /* ! */ || c === 63 /* ? */) count++;
    }
    takeNodes(count, config);
};

/**
 * Takes `count` from the parse's element budget (see takeXmlElements): for content inside a package that
 * is not XML (a DOCX chunk of plain text, RTF or MHT), whose lines, control words or parts become nodes
 * as elements do. Counted before the content is read, since it inflates out of a zip: 105 KB of DOCX
 * held an RTF chunk of 12 million paragraphs.
 */
export const takeNodes = (count: number, config: OfficeParserConfig): void => {
    let budget = xmlElementBudgets.get(config);
    if (!budget) xmlElementBudgets.set(config, budget = { left: config.decompressionLimits?.maxXmlElements ?? DEFAULT_MAX_XML_ELEMENTS });
    budget.left -= count;
    if (budget.left < 0) throw getOfficeError(OfficeErrorType.XML_ELEMENT_LIMIT_EXCEEDED, config, config.decompressionLimits?.maxXmlElements ?? DEFAULT_MAX_XML_ELEMENTS);
};

/** How many times `byte` occurs in `buffer`. */
export const countByte = (buffer: Uint8Array, byte: number): number => {
    let count = 0;
    for (let i = buffer.indexOf(byte); i !== -1; i = buffer.indexOf(byte, i + 1)) count++;
    return count;
};
const DEFAULT_MAX_XML_ELEMENTS = 2000000;

export const parseXmlString = (xml: string, options: { locator?: boolean; config?: OfficeParserConfig } = {}): Document => {
    if (options.config) takeXmlElements(xml, options.config);
    // Recoverable problems (an entity it cannot resolve, which it keeps as written) are not printed:
    // xmldom writes each to the console by default, so a document could fill a host's logs at will.
    // A fatal one still throws after this returns.
    const parser = new DOMParser({ locator: options.locator, onError: () => {} });
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
    // The outermost matches only (see getOutermostElements): a match nested in another is reached by
    // reading that one, and returning it too read nested content again at every level (notes, comments
    // or equations nested in their own kind grew with the square of the depth, or ran out of memory).
    // getAllElementsByTagName keeps the nested ones, for a reader that handles each flatly.
    const results = getOutermostElements(element, tagName);
    // Resilience: If prefixed tag (e.g., 'dc:title') not found, try local name (e.g., 'title')
    if (results.length === 0 && tagName.includes(':')) return getOutermostElements(element, tagName.split(':').pop()!);
    return results;
};

/** Every descendant named `tagName`, nested ones included, for a reader that handles each on its own (reading none of the others again). */
export const getAllElementsByTagName = (element: Element | Document, tagName: string): Element[] => {
    const results = Array.from(element.getElementsByTagName(tagName)) as Element[];
    if (results.length === 0 && tagName.includes(':')) return Array.from(element.getElementsByTagName(tagName.split(':').pop()!)) as Element[];
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
const serializer = new XMLSerializer();

export const serializeXml = (node: Node, options: { preserveWhitespace?: boolean } = {}): string => {
    // Note: xmldom's XMLSerializer doesn't natively support a 'pretty' or 'preserve' 
    // flag in a way that matches all user expectations, but it defaults to 
    // preserving structure. Formatting (indentation) is usually handled by the 
    // parser's initial whitespace handling.
    // @ts-ignore - xmldom's Node is compatible with the global Node interface
    return serializer.serializeToString(node as any);
};

/** What getSourceSubstring returns for an element longer than it may read. */
const SOURCE_TOO_LONG = Symbol('source too long');

/**
 * Attempts to extract the original raw substring from the source XML for a given node.
 * Requires the document to have been parsed with { locator: true }.
 * 
 * @param node - The DOM node to extract source for
 * @param sourceXml - The original XML source string
 * @param maxLength - The longest substring to read: past it, the scan stops (SOURCE_TOO_LONG)
 * @returns The raw XML substring, or undefined if it cannot be reliably determined
 */
const getSourceSubstring = (node: any, sourceXml: string, maxLength = Infinity): string | undefined | typeof SOURCE_TOO_LONG => {
    if (!node || typeof node.lineNumber !== 'number' || typeof node.columnNumber !== 'number') {
        return undefined;
    }
    const starts = lineStartsOf(sourceXml);
    if (node.lineNumber < 1 || node.lineNumber > starts.length) return undefined;
    const startIdx = starts[node.lineNumber - 1] + node.columnNumber - 1;
    if (!isElement(node) || !sourceXml.startsWith('<' + node.tagName, startIdx)) return undefined;
    const tagName: string = node.tagName;

    // The start tag's end, past quoted attribute values (which may hold `>`).
    const startTagEnd = tagEndAt(sourceXml, startIdx + 1 + tagName.length);
    if (startTagEnd === -1) return undefined;
    if (sourceXml.charCodeAt(startTagEnd - 1) === 47 /* / */) return sourceXml.substring(startIdx, startTagEnd + 1);

    // Its matching end tag: elements of the same name inside it open and close their own. Found by a
    // forward scan over the element alone (what it returns), where searching for the first `</name>`
    // cut a nested element short, and searched the rest of the part for each element that has none.
    // Searched only up to where the next node after it starts: an element the reader closed for a
    // malformed part has no end tag, and each such one searched the rest of the part.
    // Read no further than `maxLength` past its start either: what is left of the rawContent budget.
    const following = followingStart(node, starts, sourceXml.length);
    const truncated = following - startIdx > maxLength;
    const within = sourceXml.substring(0, truncated ? startIdx + maxLength : following);
    const open = '<' + tagName;
    const close = '</' + tagName + '>';
    let depth = 1;
    let at = startTagEnd + 1;
    // The next start tag of the name, kept until the scan passes it: searched again after each end
    // tag, a run of end tags with no start tag left searched to the end every time (160,000 nested
    // `w:t`, 3 KB of DOCX, took two minutes).
    let nextOpen = within.indexOf(open, at);
    while (true) {
        const nextClose = within.indexOf(close, at);
        if (nextClose === -1) return truncated ? SOURCE_TOO_LONG : undefined;
        while (nextOpen !== -1 && nextOpen < nextClose) {
            const after = sourceXml.charCodeAt(nextOpen + open.length);
            const end = tagEndAt(sourceXml, nextOpen + open.length);
            if (end === -1) return undefined;
            // The same name (not a longer one it starts), and not self-closing.
            if ((after === 62 || after === 47 || after === 32 || after === 9 || after === 10 || after === 13) && sourceXml.charCodeAt(end - 1) !== 47) depth++;
            nextOpen = within.indexOf(open, end + 1);
        }
        at = nextClose + close.length;
        if (--depth === 0) return sourceXml.substring(startIdx, at);
    }
};

/** Where the first node after `node` in the document (not inside it) starts, or `fallback`. */
const followingStart = (node: any, starts: number[], fallback: number): number => {
    for (let n = node; n; n = n.parentNode) {
        let sibling = n.nextSibling;
        while (sibling && typeof sibling.lineNumber !== 'number') sibling = sibling.nextSibling;
        if (sibling && sibling.lineNumber >= 1 && sibling.lineNumber <= starts.length) return starts[sibling.lineNumber - 1] + sibling.columnNumber - 1;
    }
    return fallback;
};

/** Where the tag whose name ends before `from` ends (its `>`), past quoted attribute values; -1 if it does not. */
const tagEndAt = (xml: string, from: number): number => {
    let quote = 0;
    for (let i = from; i < xml.length; i++) {
        const c = xml.charCodeAt(i);
        if (quote) { if (c === quote) quote = 0; }
        else if (c === 34 || c === 39) quote = c;
        else if (c === 62) return i;
        else if (c === 60) return -1;
    }
    return -1;
};

/**
 * Where each line of `source` starts, for the parse's locator positions: computed once per source,
 * where splitting the source into lines for every node took time in the square of its size.
 */
let lineStartsSource: string | undefined;
let lineStarts: number[] = [];
const lineStartsOf = (source: string): number[] => {
    if (source !== lineStartsSource) {
        lineStarts = [0];
        for (let i = source.indexOf('\n'); i !== -1; i = source.indexOf('\n', i + 1)) lineStarts.push(i + 1);
        lineStartsSource = source;
    }
    return lineStarts;
};

/**
 * High-level helper to get raw content for a node based on OfficeParserConfig.
 * 
 * @param node - The DOM node
 * @param sourceXml - The original source XML string
 * @param config - The parser configuration
 * @returns The raw content string (serialized or original)
 */
export const getRawContent = (node: Node, sourceXml: string, config: OfficeParserConfig): string | undefined => {
    // A spent budget is checked before serializing: nested tables re-serialize every level below
    // them, so skipping the work, not only the result, is what keeps it linear.
    const budget = rawContentBudget(config);
    if (budget.spent) return undefined;
    let raw: string | undefined;
    if (config.serializeRawContent === false) {
        const found = getSourceSubstring(node, sourceXml, budget.left);
        // Longer than the budget has left: it is spent, without reading or serializing the rest.
        if (found === SOURCE_TOO_LONG) { spendRawContent(config); return undefined; }
        raw = found;
    }
    if (!raw) raw = serializeXml(node, { preserveWhitespace: config.preserveXmlWhitespace });
    return chargeRawContent(raw, config);
};

/** What each parse (its config object) may still attach as rawContent (see DecompressionLimits.maxRawContentLength). */
const rawContentBudgets = new WeakMap<object, { left: number; spent: boolean }>();
const DEFAULT_MAX_RAW_CONTENT_LENGTH = 64 * 1024 * 1024;
const rawContentBudget = (config: OfficeParserConfig): { left: number; spent: boolean } => {
    let budget = rawContentBudgets.get(config);
    if (!budget) rawContentBudgets.set(config, budget = { left: config.decompressionLimits?.maxRawContentLength ?? DEFAULT_MAX_RAW_CONTENT_LENGTH, spent: false });
    return budget;
};

/**
 * Charges `raw` to its parse's rawContent budget and returns it, or undefined once the budget is spent
 * (warning once): a node's rawContent holds the markup of everything nested in it, so nested tables and
 * repeated ODF cells carry the same bytes many times over. A repeat that shares a string is charged
 * again, since each node that carries it costs its length to anyone who serializes the AST.
 */
export const chargeRawContent = (raw: string | undefined, config: OfficeParserConfig): string | undefined => {
    if (raw === undefined) return undefined;
    const budget = rawContentBudget(config);
    if (!budget.spent && raw.length <= budget.left) {
        budget.left -= raw.length;
        return raw;
    }
    spendRawContent(config);
    return undefined;
};

/** Marks the parse's rawContent budget spent, warning the first time. */
const spendRawContent = (config: OfficeParserConfig): void => {
    const budget = rawContentBudget(config);
    budget.left = 0;
    if (!budget.spent) {
        budget.spent = true;
        logWarning(OfficeWarningType.RAW_CONTENT_LIMIT_EXCEEDED, config, config.decompressionLimits?.maxRawContentLength ?? DEFAULT_MAX_RAW_CONTENT_LENGTH);
    }
};
/**
 * Gets the first element with the specified tag name within a parent element.
 * 
 * @param parent - The parent element or document to search within
 * @param tagName - The tag name to search for
 * @returns The first matching element, or undefined if none found
 */
export const getFirstElementByTagName = (parent: Element | Document, tagName: string): Element | undefined => {
    // Walks the descendants in document order and stops at the first match: xmldom's
    // getElementsByTagName builds the whole list first, so each call read the whole subtree, and a
    // lookup at each level of nested content (tables in tables, frames in frames) took time in the
    // product of the depth and the subtree.
    let node: Node | null = parent.firstChild;
    while (node) {
        if (node.nodeType === 1 && (node as Element).tagName === tagName) return node as Element;
        if (node.firstChild) { node = node.firstChild; continue; }
        while (node && node !== parent && !node.nextSibling) node = node.parentNode;
        if (!node || node === parent) return undefined;
        node = node.nextSibling;
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
 * The `tag` elements under `root` not inside another `tag` element under it, in document order, and
 * not inside any element named in `skip`: a note's own paragraphs, not a nested text box's; a sheet's
 * rows, not those of a table in one of its cells. Descendant lookups (getElementsByTagName) return the
 * nested ones as well, and a reader that also reaches them through the outer ones read them once per
 * level of nesting, which grew with the square of the depth, or doubled per level.
 */
export const getOutermostElements = (root: Element | Document, tag: string, skip?: ReadonlySet<string>): Element[] => {
    const out: Element[] = [];
    const stack: Element[] = [];
    const pushChildren = (node: Node) => {
        for (let i = node.childNodes.length - 1; i >= 0; i--) {
            const child = node.childNodes[i];
            if (isElement(child)) stack.push(child);
        }
    };
    pushChildren(root);
    while (stack.length) {
        const node = stack.pop()!;
        if (node.nodeName === tag) out.push(node);
        else if (!skip?.has(node.nodeName)) pushChildren(node);
    }
    return out;
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
 * The child elements of `parent` named `tagName`, falling back to its local name when none carries the
 * prefix (as getElementsByTagName does). For content that nests (a table's rows and cells, a text
 * body's paragraphs, a shared string's runs), reading children rather than descendants reads each
 * level once: descendant lookups read the nested levels again at each level.
 */
export const getChildElements = (parent: Element | Document, tagName: string): Element[] => {
    const node = (parent as Document).documentElement && parent.nodeType === 9 ? (parent as Document).documentElement : parent as Element;
    const own = getDirectChildren(node, tagName);
    if (own.length > 0 || !tagName.includes(':')) return own;
    return getDirectChildren(node, tagName.split(':').pop()!);
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
export const parseOfficeMetadata = (xmlContent: string, config?: OfficeParserConfig): OfficeMetadata => {
    // Step 1: Parse the XML content into a DOM document
    const xml = parseXmlString(xmlContent, { config });
    const metadata: OfficeMetadata = {};

    // Check for OOXML Core Properties
    const coreProperties = getElementsByTagName(xml, "cp:coreProperties")[0];
    if (coreProperties) {
        metadata.nativeProperties = {};
        for (let i = 0; i < coreProperties.childNodes.length; i++) {
            const child = coreProperties.childNodes[i];
            if (isElement(child)) {
                setOwn(metadata.nativeProperties, child.tagName, child.textContent);
            }
        }

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

        const keywords = getElementsByTagName(coreProperties, "cp:keywords")[0];
        if (keywords && keywords.textContent) metadata.keywords = keywords.textContent;

        return metadata;
    }

    // Check for ODF Meta
    const officeMeta = getElementsByTagName(xml, "office:meta")[0];
    if (officeMeta) {
        metadata.nativeProperties = {};
        for (let i = 0; i < officeMeta.childNodes.length; i++) {
            const child = officeMeta.childNodes[i];
            if (isElement(child)) {
                setOwn(metadata.nativeProperties, child.tagName, child.textContent);
            }
        }

        const title = getElementsByTagName(officeMeta, "dc:title")[0];
        if (title && title.textContent) metadata.title = title.textContent;

        const author = getElementsByTagName(officeMeta, "dc:creator")[0];
        if (author && author.textContent) metadata.author = author.textContent;

        const description = getElementsByTagName(officeMeta, "dc:description")[0];
        if (description && description.textContent) metadata.description = description.textContent;

        const subject = getElementsByTagName(officeMeta, "dc:subject")[0];
        if (subject && subject.textContent) metadata.subject = subject.textContent;

        const keywordElements = getElementsByTagName(officeMeta, "meta:keyword");
        if (keywordElements.length > 0) {
            metadata.keywords = keywordElements.map(k => k.textContent).filter(Boolean).join(', ');
        }

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
                    setOwn(customProperties, name, raw.toLowerCase() === "true");
                } else if (valueType === "float") {
                    const num = Number(raw);
                    if (!isNaN(num)) setOwn(customProperties, name, num);
                } else if (valueType === "date" || valueType === "time") {
                    const date = parseOfficeDate(raw);
                    if (date) setOwn(customProperties, name, date);
                    else setOwn(customProperties, name, raw);
                } else {
                    setOwn(customProperties, name, raw);
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
export const parseOOXMLCustomProperties = (xmlContent: string, config?: OfficeParserConfig): Record<string, string | number | boolean | Date> => {
    const xml = parseXmlString(xmlContent, { config });
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
                setOwn(result, name, text);
            } else if (/vt:bool/.test(tag)) {
                setOwn(result, name, text.toLowerCase() === 'true');
            } else if (/vt:(i[1248]|ui[1248]|int|uint|r4|r8|decimal)/.test(tag)) {
                const num = Number(text);
                if (!isNaN(num)) setOwn(result, name, num);
            } else if (/vt:filetime|vt:date/.test(tag)) {
                const date = parseOfficeDate(text);
                if (date) setOwn(result, name, date);
                else setOwn(result, name, text);
            } else if (text) {
                // Fallback: store as string for any other vt: type
                setOwn(result, name, text);
            }
            break; // only one value element per property
        }
    }

    return result;
};

/**
 * Parses OOXML application properties from `docProps/app.xml`.
 * 
 * Application properties contain document statistics and application settings.
 * 
 * @param xmlContent - Raw XML string from `docProps/app.xml`
 * @returns A record of property name -> typed value
 */
export const parseOOXMLAppProperties = (xmlContent: string, config?: OfficeParserConfig): Record<string, string | number | boolean> => {
    const xml = parseXmlString(xmlContent, { config });
    const result: Record<string, string | number | boolean> = {};

    const appProperties = getElementsByTagName(xml, "Properties")[0];
    if (appProperties) {
        for (let i = 0; i < appProperties.childNodes.length; i++) {
            const child = appProperties.childNodes[i];
            if (isElement(child)) {
                const text = child.textContent || '';
                if (text.toLowerCase() === 'true') result[child.tagName] = true;
                else if (text.toLowerCase() === 'false') result[child.tagName] = false;
                else if (!isNaN(Number(text)) && text.trim() !== '') result[child.tagName] = Number(text);
                else result[child.tagName] = text;
            }
        }
    }
    return result;
};

/**
 * Decodes XML entities (standard named entities, decimal, and hexadecimal entities) in a string.
 * Useful when parsing content with regular expressions instead of a full DOM parser.
 * 
 * @param text - The XML-encoded string
 * @returns The decoded string
 */
export const decodeXmlEntities = (text: string): string => {
    // A reference holds no `&`: each `&` without a `;` read on to the end of the text.
    return text.replace(/&([^;&]+);/g, (match, entity) => {
        if (entity.startsWith('#')) {
            if (entity[1] === 'x' || entity[1] === 'X') {
                const hex = entity.slice(2);
                const code = parseInt(hex, 16);
                return (!isNaN(code) && code >= 0 && code <= 0x10FFFF) ? String.fromCodePoint(code) : match;
            } else {
                const dec = entity.slice(1);
                const code = parseInt(dec, 10);
                return (!isNaN(code) && code >= 0 && code <= 0x10FFFF) ? String.fromCodePoint(code) : match;
            }
        }
        switch (entity) {
            case 'amp': return '&';
            case 'lt': return '<';
            case 'gt': return '>';
            case 'quot': return '"';
            case 'apos': return "'";
            default: return match;
        }
    });
};


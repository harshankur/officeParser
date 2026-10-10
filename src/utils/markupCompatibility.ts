/**
 * Markup Compatibility (ECMA-376 Part 3): `mc:AlternateContent` holds content for newer readers in
 * `mc:Choice` branches, each naming (in `Requires`) the namespaces a reader must understand to read it,
 * and a `mc:Fallback` for a reader that understands none of them. A reader takes the first Choice whose
 * namespaces it understands, else the Fallback. Taking the first Choice whatever it required read content
 * in a form the reader does not know: an emoji Word writes as `w16se:symEx`, its character only in the
 * Fallback's `w:t`, was dropped, and a chart of a newer kind was read as an empty drawing where its
 * Fallback holds a picture of it.
 */

import { isElement } from './xmlUtils.js';

/** The namespaces a reader understands: by URI, and by the prefix producers bind them to. */
export interface UnderstoodNamespaces {
    uris: ReadonlySet<string>;
    prefixes: ReadonlySet<string>;
}

const understood = (namespaces: Record<string, string>): UnderstoodNamespaces => ({
    uris: new Set(Object.values(namespaces)),
    prefixes: new Set(Object.keys(namespaces)),
});

/** DrawingML and VML, which both the Word and PowerPoint readers read pictures, shapes, math and SmartArt of. */
const DRAWING_NAMESPACES: Record<string, string> = {
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    a14: 'http://schemas.microsoft.com/office/drawing/2010/main',
    pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    dgm: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
    c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    v: 'urn:schemas-microsoft-com:vml',
    o: 'urn:schemas-microsoft-com:office:office',
};

/**
 * What the Word reader understands: WordprocessingML itself, and the drawing containers whose text boxes
 * and pictures it reads (Word 2010 shapes, groups and canvases). Word's later additions (`w14`, `w15`,
 * `w16se` and the like) are not: their Fallback is what the reader reads.
 */
export const WORD_NAMESPACES = understood({
    ...DRAWING_NAMESPACES,
    w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
    wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    w10: 'urn:schemas-microsoft-com:office:word',
});

/** What the PowerPoint reader understands: PresentationML and the drawing namespaces. */
export const PRESENTATION_NAMESPACES = understood({
    ...DRAWING_NAMESPACES,
    p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
});

/** Each document's namespace declarations on its root element, read once. */
const rootDeclarations = new WeakMap<Document, Map<string, string>>();

/** The namespace a document's root element binds `prefix` to. */
const rootNamespace = (element: Element, prefix: string): string | undefined => {
    const document = element.ownerDocument;
    if (!document?.documentElement) return undefined;
    let declarations = rootDeclarations.get(document);
    if (!declarations) {
        declarations = new Map();
        const attributes = document.documentElement.attributes;
        for (let i = 0; i < attributes.length; i++) {
            const name = attributes[i].name;
            if (name.startsWith('xmlns:')) declarations.set(name.slice(6), attributes[i].value);
        }
        rootDeclarations.set(document, declarations);
    }
    return declarations.get(prefix);
};

/**
 * Whether `prefix`, as a Choice requires it, names a namespace the reader understands. The prefix is
 * looked up where producers declare it (on the Choice, on its AlternateContent, or on the part's root)
 * rather than through every ancestor: AlternateContent nested thousands deep made each lookup a walk of
 * the depth. A prefix declared nowhere there is taken by its name.
 */
const understands = (choice: Element, prefix: string, namespaces: UnderstoodNamespaces): boolean => {
    const declared = choice.getAttribute(`xmlns:${prefix}`)
        || (choice.parentNode && isElement(choice.parentNode) ? choice.parentNode.getAttribute(`xmlns:${prefix}`) : null)
        || rootNamespace(choice, prefix);
    return declared ? namespaces.uris.has(declared) : namespaces.prefixes.has(prefix);
};

const isBranch = (node: Element, name: 'Choice' | 'Fallback'): boolean => node.nodeName === `mc:${name}` || node.nodeName === name;

/**
 * The branch of `alternateContent` a reader understanding `namespaces` reads: its first Choice whose
 * required namespaces are all understood, else its Fallback, else none. Its own children only, as the
 * schema has them: looked up through its subtree, AlternateContent nested in AlternateContent took time
 * in the product of the depth and the content.
 */
export const alternateContentBranch = (alternateContent: Element, namespaces: UnderstoodNamespaces): Element | undefined => {
    const children = alternateContent.childNodes;
    for (let i = 0; i < children.length; i++) {
        const child = children[i];
        if (!isElement(child)) continue;
        if (isBranch(child, 'Choice')) {
            const required = new Set((child.getAttribute('Requires') || '').split(/\s+/).filter(Boolean));
            // More prefixes than the reader understands cannot all be understood (and are not looked up).
            if (required.size <= namespaces.prefixes.size && [...required].every(prefix => understands(child, prefix, namespaces))) return child;
        } else if (isBranch(child, 'Fallback')) {
            return child;
        }
    }
    return undefined;
};

/** Whether `element` is an `mc:AlternateContent`. */
export const isAlternateContent = (element: Element): boolean => element.nodeName === 'mc:AlternateContent' || element.nodeName === 'AlternateContent';

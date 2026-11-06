/**
 * Where a DOCTYPE with an XML internal subset ends. HTML ends every DOCTYPE at its first `>` (a `[`
 * in it is an error, not a subset); XML (an XHTML page) lets a DOCTYPE carry an internal subset in
 * brackets, whose entity values may hold `>` and `]`. Dependency free, so the type sniffer can use it.
 */

/** XML white space (S): space, tab, carriage return, line feed. */
const isXmlSpace = (c: string | undefined): boolean => c === ' ' || c === '\t' || c === '\n' || c === '\r';

/** A parameter-entity reference (`%name;`), read where it starts. */
const PARAMETER_ENTITY = /%[A-Za-z_:][\w.:-]*;/y;

/**
 * The index just past the `>` of the DOCTYPE at `start` (`<!DOCTYPE`, any case), when it has an
 * internal subset: a `[` before its first `>` outside quoted literals, then a well-formed subset
 * (markup declarations, whose quoted values are passed over whole, comments, processing instructions,
 * parameter-entity references and white space) closed by `]`, white space and `>`. Otherwise -1: the
 * declaration ends at its first `>`, as HTML reads every DOCTYPE (`<!DOCTYPE html [>` and a `[` in a
 * quoted public id included). One forward pass from `start`, each search resuming where the last ended,
 * so a caller reading one DOCTYPE per document stays linear.
 */
export const doctypeSubsetEnd = (text: string, start: number): number => Math.max(-1, doctypeSubsetScan(text, start));

/**
 * doctypeSubsetEnd's reading of a text that may be the start of a document only (the type sniffer's):
 * where the DOCTYPE's subset ends, -1 where it is not well formed, and -2 where the text ends before
 * the DOCTYPE does (more of the document decides).
 */
export const doctypeSubsetScan = (text: string, start: number): number => {
    const n = text.length;
    let i = start + 9;
    // The name and external id, to the subset's `[`: a quoted literal may hold `[` or `>`.
    for (; i < n && text[i] !== '[' && text[i] !== '>'; i++) {
        if (text[i] === '"' || text[i] === '\'') {
            i = text.indexOf(text[i], i + 1);
            if (i === -1) return -2;
        }
    }
    if (i >= n) return -2;
    if (text[i] !== '[') return -1;
    i++;
    while (i < n) {
        const c = text[i];
        if (isXmlSpace(c)) { i++; continue; }
        if (c === ']') {
            i++;
            while (isXmlSpace(text[i])) i++;
            return i >= n ? -2 : text[i] === '>' ? i + 1 : -1;
        }
        if (c === '%') {
            PARAMETER_ENTITY.lastIndex = i;
            if (!PARAMETER_ENTITY.test(text)) return /^%[A-Za-z_:]?[\w.:-]*$/.test(text.slice(i)) ? -2 : -1;
            i = PARAMETER_ENTITY.lastIndex;
            continue;
        }
        if (c !== '<') return -1;
        if (i + 3 >= n) return -2;
        if (text.startsWith('<!--', i)) {
            const end = text.indexOf('-->', i + 4);
            if (end === -1) return -2;
            i = end + 3;
        } else if (text[i + 1] === '?') {
            const end = text.indexOf('?>', i + 2);
            if (end === -1) return -2;
            i = end + 2;
        } else if (text[i + 1] === '!' && /[A-Za-z]/.test(text[i + 2] ?? '')) {
            // `<!ENTITY`, `<!ELEMENT`, `<!ATTLIST`, `<!NOTATION`: to its `>`, quoted values passed over whole.
            let j = i + 2;
            for (; j < n && text[j] !== '>'; j++) {
                if (text[j] === '"' || text[j] === '\'') {
                    j = text.indexOf(text[j], j + 1);
                    if (j === -1) return -2;
                }
            }
            if (j >= n) return -2;
            i = j + 1;
        } else {
            return -1;
        }
    }
    return -2;
};

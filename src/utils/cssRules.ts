/**
 * A document's own stylesheet (its `<style>` elements), read so that what its rules set (a colour, a
 * background, bold, alignment) reaches the elements they select, as a browser shows them: an HTML page
 * styled by classes lost all of it. Selectors are compounds of a type, classes and an id (`.note`,
 * `td.total`, `#title`, `.box td`), joined by descendant and child combinators, and name a class or an
 * id somewhere: a rule for every element of a type is the page's theme. A rule with any other selector
 * (attributes, pseudo-classes, siblings) is passed over, as are at-rules. Only the properties reading
 * looks at are kept, each once a rule, and an `!important` one wins over an element's own `style`.
 * Everything is bounded: a stylesheet's size, its rules, a selector's parts, a value's length and the
 * rules any one element takes.
 */

/** At most this much stylesheet text is read. */
const MAX_CSS_LENGTH = 256 * 1024;
/** At most this many selectors (a rule's comma-separated list counting each) are kept. */
const MAX_SELECTORS = 1024;
/** At most this many compounds in one selector. */
const MAX_COMPOUNDS = 4;
/** At most this many rules are tried for one element (an id's, then its classes', its type's, any). */
const MAX_CANDIDATES = 32;
/** At most this many ancestors are looked through for a descendant combinator's compound. */
const MAX_ANCESTORS = 16;
/** At most this long a value is kept (a `url()`'s address is left out of it first). */
const MAX_VALUE_LENGTH = 256;

/** The properties reading looks at (and their `-webkit-` and other vendors' names). */
const READ_PROPERTIES = new Set([
    'color', 'background', 'background-color', 'font-family', 'font-size', 'font-style', 'font-weight',
    'text-align', 'text-decoration', 'text-decoration-line', 'margin', 'margin-top', 'margin-bottom',
    'margin-left', 'margin-right', 'line-height', 'width', 'list-style', 'list-style-type',
]);

/** A rule's declarations as kept: the properties read, the last of each, in two by `!important`. */
export interface Declarations { normal: Map<string, string>; important: Map<string, string> }

/** Splits declarations at `;` outside quotes and parentheses. */
function splitDeclarations(text: string): string[] {
    const out: string[] = [];
    let depth = 0, quote = '', start = 0;
    for (let i = 0; i < text.length; i++) {
        const c = text[i];
        if (quote) { if (c === quote) quote = ''; }
        else if (c === '"' || c === "'") quote = c;
        else if (c === '(') depth++;
        else if (c === ')') { if (depth) depth--; }
        else if (c === ';' && !depth) { out.push(text.slice(start, i)); start = i + 1; }
    }
    out.push(text.slice(start));
    return out;
}

/** A value with each `url(...)`'s address left out (`url()`): a picture's data is no colour or size. */
function withoutAddresses(value: string): string {
    const lower = value.toLowerCase();
    let out = '', at = 0;
    for (let open = lower.indexOf('url(', at); open >= 0; open = lower.indexOf('url(', at)) {
        // A quoted address ends at its quote, whatever `)` it holds (`url("img (1).png")`).
        let from = open + 4;
        while (from < value.length && /\s/.test(value[from])) from++;
        const quote = value[from] === '"' || value[from] === "'" ? value[from] : '';
        const end = quote ? value.indexOf(quote, from + 1) : from - 1;
        const close = end < 0 ? -1 : value.indexOf(')', end + 1);
        out += value.slice(at, open) + 'url()';
        at = close < 0 ? value.length : close + 1;
    }
    return out + value.slice(at);
}

/** The declarations of a rule's body (or a `style`) reading looks at, the last of each, by `!important`. */
export function readDeclarations(text: string): Declarations {
    const kept: Declarations = { normal: new Map(), important: new Map() };
    for (const declaration of splitDeclarations(text)) {
        const colon = declaration.indexOf(':');
        if (colon < 0) continue;
        const property = declaration.slice(0, colon).trim().toLowerCase();
        if (!READ_PROPERTIES.has(property.replace(/^-(?:webkit|moz|ms|o)-/, ''))) continue;
        let value = withoutAddresses(declaration.slice(colon + 1)).trim();
        // A trailing `!important`, found from the end.
        let important = false;
        if (value.toLowerCase().endsWith('important')) {
            let bang = value.length - 'important'.length;
            while (bang > 0 && /\s/.test(value[bang - 1])) bang--;
            if (value[bang - 1] === '!') { value = value.slice(0, bang - 1).trim(); important = true; }
        }
        if (!value || value.length > MAX_VALUE_LENGTH) continue;
        // (A later normal one does not undo an `!important` one.)
        if (important) { kept.important.set(property, value); kept.normal.delete(property); }
        else if (!kept.important.has(property)) kept.normal.set(property, value);
    }
    return kept;
}

/** Declarations written as a `style`. */
export function writeDeclarations(declarations: Map<string, string>): string {
    return [...declarations].map(([property, value]) => `${property}: ${value}`).join('; ');
}

/** One compound selector: a type (or any), classes and an id, all of which an element must have. */
interface Compound { tag?: string; classes: string[]; id?: string }

/** A selector, its compounds right to left with the combinator before each, and its declarations. */
interface Selector {
    /** Rightmost first; `combinator[i]` joins `compounds[i]` to `compounds[i + 1]`. */
    compounds: Compound[];
    combinators: ('descendant' | 'child')[];
    specificity: number;
    order: number;
    declarations: Declarations;
}

/** A stylesheet's selectors, indexed by what their rightmost compound needs most. */
export interface Stylesheet {
    byId: Map<string, Selector[]>;
    byClass: Map<string, Selector[]>;
    byTag: Map<string, Selector[]>;
    any: Selector[];
}

/** The element shape matched: a tag, its attributes and its parent. */
export interface StyledElement {
    tagName?: string;
    attributes?: Record<string, string>;
    parent?: StyledElement;
}

const COMPOUND = /^([a-z][a-z0-9-]*|\*)?((?:[.#][a-z_][a-z0-9_-]*)*)$/i;

function readCompound(text: string): Compound | undefined {
    const m = COMPOUND.exec(text);
    if (!m || (!m[1] && !m[2])) return undefined;
    const compound: Compound = { classes: [] };
    if (m[1] && m[1] !== '*') compound.tag = m[1].toLowerCase();
    for (const part of m[2].match(/[.#][^.#]+/g) ?? []) {
        if (part[0] === '.') compound.classes.push(part.slice(1));
        else if (compound.id === undefined) compound.id = part.slice(1);
        else return undefined;
    }
    return compound;
}

function readSelector(text: string, declarations: Declarations, order: number): Selector | undefined {
    // (Whitespace collapsed first: `\s*>\s*` over a long run of spaces is tried at each of them.)
    const tokens = text.replace(/\s+/g, ' ').replace(/ ?> ?/g, ' > ').split(' ').filter(Boolean);
    const compounds: Compound[] = [];
    const combinators: ('descendant' | 'child')[] = [];
    let pending: 'descendant' | 'child' | undefined;
    for (const token of tokens) {
        if (token === '>') { if (!compounds.length || pending === 'child') return undefined; pending = 'child'; continue; }
        const compound = readCompound(token);
        if (!compound) return undefined;
        if (compounds.length) combinators.push(pending ?? 'descendant');
        pending = undefined;
        compounds.push(compound);
        if (compounds.length > MAX_COMPOUNDS) return undefined;
    }
    if (!compounds.length || pending) return undefined;
    // A rule for every element of a type (`p`, `td a`) is the page's theme, as a reader's own stylesheet
    // is, and no part of the document; a class or an id names content the author styled.
    if (!compounds.some(c => c.classes.length || c.id !== undefined)) return undefined;
    compounds.reverse();
    combinators.reverse();
    const ids = compounds.filter(c => c.id !== undefined).length;
    const classes = compounds.reduce((n, c) => n + c.classes.length, 0);
    const tags = compounds.filter(c => c.tag).length;
    return { compounds, combinators, specificity: ids * 10_000 + classes * 100 + tags, order, declarations };
}

/** Reads the rules of the stylesheets given (their texts, in document order). */
export function readStylesheet(texts: string[]): Stylesheet {
    const sheet: Stylesheet = { byId: new Map(), byClass: new Map(), byTag: new Map(), any: [] };
    let css = '';
    for (const text of texts) { if (css.length >= MAX_CSS_LENGTH) break; css += text.slice(0, MAX_CSS_LENGTH - css.length) + '\n'; }
    css = css.replace(/\/\*[\s\S]*?(?:\*\/|$)/g, ' ');
    let count = 0, at = 0;
    while (at < css.length && count < MAX_SELECTORS) {
        const open = css.indexOf('{', at);
        if (open < 0) break;
        const prelude = css.slice(at, open).trim();
        // A block's end, past nested blocks (an at-rule's): its rules are passed over with it.
        let depth = 1, end = open + 1;
        while (end < css.length && depth > 0) { const c = css[end]; if (c === '{') depth++; else if (c === '}') depth--; end++; }
        const body = css.slice(open + 1, end - 1);
        at = end;
        if (!prelude || prelude.startsWith('@') || body.includes('{')) continue;
        const declarations = readDeclarations(body);
        if (!declarations.normal.size && !declarations.important.size) continue;
        for (const text of prelude.split(',')) {
            const selector = readSelector(text, declarations, count);
            if (!selector) continue;
            count++;
            const right = selector.compounds[0];
            const index = right.id !== undefined ? [sheet.byId, right.id] as const
                : right.classes.length ? [sheet.byClass, right.classes[0]] as const
                : right.tag ? [sheet.byTag, right.tag] as const : undefined;
            if (index) { const list = index[0].get(index[1]); if (list) list.push(selector); else index[0].set(index[1], [selector]); }
            else sheet.any.push(selector);
            if (count >= MAX_SELECTORS) break;
        }
    }
    return sheet;
}

/** Whether a stylesheet has any rule. */
export function isEmptyStylesheet(sheet: Stylesheet): boolean {
    return !sheet.any.length && !sheet.byId.size && !sheet.byClass.size && !sheet.byTag.size;
}

/** An element's classes, split once. */
const CLASSES = new WeakMap<StyledElement, string[]>();
function classesOf(element: StyledElement): string[] {
    let classes = CLASSES.get(element);
    if (!classes) {
        const value = element.attributes?.class;
        classes = value ? value.split(/\s+/).filter(Boolean).slice(0, 32) : [];
        CLASSES.set(element, classes);
    }
    return classes;
}

function matchesCompound(element: StyledElement, compound: Compound, classes: string[]): boolean {
    if (compound.tag && element.tagName?.toLowerCase() !== compound.tag) return false;
    if (compound.id !== undefined && element.attributes?.id !== compound.id) return false;
    return compound.classes.every(c => classes.includes(c));
}

/** Whether `element` is selected: its own compound, then its ancestors' (a bounded number of them). */
function matches(element: StyledElement, selector: Selector): boolean {
    if (!matchesCompound(element, selector.compounds[0], classesOf(element))) return false;
    let current: StyledElement | undefined = element;
    for (let i = 1; i < selector.compounds.length; i++) {
        const compound = selector.compounds[i];
        if (selector.combinators[i - 1] === 'child') {
            current = current?.parent;
            if (!current || !matchesCompound(current, compound, classesOf(current))) return false;
            continue;
        }
        let found = false;
        for (let up: StyledElement | undefined = current?.parent, steps = 0; up && steps < MAX_ANCESTORS; up = up.parent, steps++) {
            if (matchesCompound(up, compound, classesOf(up))) { current = up; found = true; break; }
        }
        if (!found) return false;
    }
    return true;
}

/**
 * The declarations the stylesheet gives an element, as a browser applies them (least specific first,
 * then in document order, the last of each property kept): the normal ones, to stand before its own
 * `style`, which comes after and so wins, and the `!important` ones, to stand after it.
 */
export function declarationsFor(sheet: Stylesheet, element: StyledElement): { normal: string; important: string } {
    const candidates: Selector[] = [];
    const add = (list: Selector[] | undefined) => { if (list) for (const s of list) { if (candidates.length >= MAX_CANDIDATES) return; candidates.push(s); } };
    const id = element.attributes?.id;
    if (id) add(sheet.byId.get(id));
    for (const c of classesOf(element)) add(sheet.byClass.get(c));
    if (element.tagName) add(sheet.byTag.get(element.tagName.toLowerCase()));
    add(sheet.any);
    const matched = [...new Set(candidates)].filter(s => matches(element, s));
    if (!matched.length) return { normal: '', important: '' };
    matched.sort((a, b) => a.specificity - b.specificity || a.order - b.order);
    const normal = new Map<string, string>(), important = new Map<string, string>();
    for (const s of matched) {
        for (const [property, value] of s.declarations.normal) { normal.delete(property); normal.set(property, value); }
        for (const [property, value] of s.declarations.important) { important.delete(property); important.set(property, value); }
    }
    return { normal: writeDeclarations(normal), important: writeDeclarations(important) };
}

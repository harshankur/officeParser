/**
 * A document's own stylesheet (its `<style>` elements), read so that what its rules set (a colour, a
 * background, bold, alignment) reaches the elements they select, as a browser shows them: an HTML page
 * styled by classes lost all of it. Selectors are compounds of a type, classes and an id (`.note`,
 * `td.total`, `#title`, `.box td`), joined by descendant and child combinators, and name a class or an
 * id somewhere: a rule for every element of a type is the page's theme. A rule with any other selector
 * (attributes, pseudo-classes, siblings) is passed over, as are at-rules. Everything is bounded: a
 * stylesheet's size, its rules, a selector's parts and the rules any one element takes.
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

/** One compound selector: a type (or any), classes and an id, all of which an element must have. */
interface Compound { tag?: string; classes: string[]; id?: string }

/** A selector, its compounds right to left with the combinator before each, and its declarations. */
interface Selector {
    /** Rightmost first; `combinator[i]` joins `compounds[i]` to `compounds[i + 1]`. */
    compounds: Compound[];
    combinators: ('descendant' | 'child')[];
    specificity: number;
    order: number;
    declarations: string;
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

function readSelector(text: string, declarations: string, order: number): Selector | undefined {
    const tokens = text.trim().replace(/\s*>\s*/g, ' > ').split(/\s+/).filter(Boolean);
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
        for (const text of prelude.split(',')) {
            const selector = readSelector(text, body, count);
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
 * The declarations the stylesheet gives an element, in the order a browser applies them (least specific
 * first, then in document order), to stand before its own `style`, which comes after and so wins.
 */
export function declarationsFor(sheet: Stylesheet, element: StyledElement): string {
    const candidates: Selector[] = [];
    const add = (list: Selector[] | undefined) => { if (list) for (const s of list) { if (candidates.length >= MAX_CANDIDATES) return; candidates.push(s); } };
    const id = element.attributes?.id;
    if (id) add(sheet.byId.get(id));
    for (const c of classesOf(element)) add(sheet.byClass.get(c));
    if (element.tagName) add(sheet.byTag.get(element.tagName.toLowerCase()));
    add(sheet.any);
    const matched = [...new Set(candidates)].filter(s => matches(element, s));
    if (!matched.length) return '';
    matched.sort((a, b) => a.specificity - b.specificity || a.order - b.order);
    return matched.map(s => s.declarations.trim()).filter(Boolean).join('; ');
}

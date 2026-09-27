import { OfficeContentNode, OfficeParserAST } from '../types.js';

/**
 * Metadata and formatting fields whose type is not a string: numbers (levels, indices, positions,
 * spans, page measures), booleans, and the one list (`anchorIds`). Every other known field is a
 * string, or a union of string literals.
 */
const NUMBER_FIELDS = new Set(['level', 'indentation', 'itemIndex', 'row', 'col', 'rowSpan', 'colSpan', 'slideNumber', 'pageNumber', 'pageWidth', 'pageHeight', 'rotation', 'pages']);
const BOOLEAN_FIELDS = new Set(['isTask', 'checked', 'unreferenced', 'bold', 'italic', 'underline', 'strikethrough', 'subscript', 'superscript']);
const STRING_FIELDS = new Set([
    'style', 'alignment', 'listType', 'listId', 'noteId', 'sheetName', 'align', 'backgroundColor', 'color', 'size', 'font',
    'attachmentName', 'altText', 'url', 'width', 'height', 'title', 'link', 'linkType', 'linkTitle', 'embedType', 'videoId', 'label',
    'admonitionType', 'sourceSyntax', 'pageLabel', 'pageName', 'abbreviationTitle', 'citationKey', 'wikilink', 'noteType', 'breakType',
    'clear', 'language', 'math', 'author', 'initials', 'date', 'commentId', 'type', 'chartType',
]);
/** `paragraphIndentation`'s fields, all numbers (points). */
const INDENTATION_FIELDS = ['left', 'right', 'firstLine', 'hanging'];

const isPrimitive = (value: unknown): value is string | number | boolean => typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean';

/** `value` as a string: itself, a number or boolean written out, a list's plain values joined; else undefined. */
function asString(value: unknown): string | undefined {
    if (typeof value === 'string') return value;
    if (value instanceof Date) return Number.isNaN(value.getTime()) ? undefined : value.toISOString();
    if (typeof value === 'number' || typeof value === 'boolean') return String(value);
    if (Array.isArray(value)) return value.flat(8).filter(isPrimitive).map(String).join(' ');
    return undefined;
}

/** `value` as a finite number: itself, or a string holding one; else undefined. */
function asNumber(value: unknown): number | undefined {
    if (typeof value === 'number') return Number.isFinite(value) ? value : undefined;
    if (typeof value === 'string' && /^\s*-?\d+(?:\.\d+)?\s*$/.test(value)) return Number(value);
    return undefined;
}

/** `value` as a boolean: itself, or `'true'`/`'false'`; else undefined. */
function asBoolean(value: unknown): boolean | undefined {
    if (typeof value === 'boolean') return value;
    if (value === 'true' || value === 'false') return value === 'true';
    return undefined;
}

/**
 * `record` (a node's metadata or formatting) with each known field of the wrong type coerced to its
 * type, or removed when it cannot be; `record` itself when every field is well typed. Unknown fields
 * are kept as they are: no writer reads them.
 */
function normalizeRecord(record: Record<string, unknown>): Record<string, unknown> {
    let out: Record<string, unknown> | undefined;
    const set = (key: string, value: unknown) => {
        out ??= { ...record };
        if (value === undefined) delete out[key];
        else out[key] = value;
    };
    for (const key of Object.keys(record)) {
        const value = record[key];
        if (value === undefined || value === null) continue;
        if (key === 'anchorIds') {
            if (Array.isArray(value) && value.every(id => typeof id === 'string')) continue;
            const ids = (Array.isArray(value) ? value.flat(8) : [value]).filter(isPrimitive).map(String);
            set(key, ids.length ? ids : undefined);
        } else if (key === 'paragraphIndentation') {
            if (typeof value !== 'object' || Array.isArray(value)) { set(key, undefined); continue; }
            const indentation = value as Record<string, unknown>;
            if (INDENTATION_FIELDS.every(field => indentation[field] === undefined || typeof indentation[field] === 'number')) continue;
            const fixed: Record<string, number> = {};
            for (const field of INDENTATION_FIELDS) {
                const n = asNumber(indentation[field]);
                if (n !== undefined) fixed[field] = n;
            }
            set(key, fixed);
        } else if (NUMBER_FIELDS.has(key)) {
            if (typeof value !== 'number' || !Number.isFinite(value)) set(key, asNumber(value));
        } else if (BOOLEAN_FIELDS.has(key)) {
            if (typeof value !== 'boolean') set(key, asBoolean(value));
        } else if (STRING_FIELDS.has(key)) {
            if (typeof value !== 'string') set(key, asString(value));
        }
    }
    return out ?? record;
}

/** `node` with well-typed text, metadata, formatting and attributes, and its descendants the same; `node` itself when nothing changes. */
function normalizeNode(node: OfficeContentNode): OfficeContentNode {
    let next: any = node;
    const change = (key: string, value: unknown) => {
        if (next === node) next = { ...node };
        if (value === undefined) delete next[key];
        else next[key] = value;
    };
    if (node.text !== undefined && node.text !== null && typeof node.text !== 'string') change('text', asString(node.text));
    for (const key of ['metadata', 'formatting'] as const) {
        const record = (node as any)[key];
        if (record === undefined || record === null) continue;
        if (typeof record !== 'object' || Array.isArray(record)) { change(key, undefined); continue; }
        const fixed = normalizeRecord(record);
        if (fixed !== record) change(key, fixed);
    }
    const attributes = (node as any).htmlAttributes;
    if (attributes !== undefined && attributes !== null) {
        if (typeof attributes !== 'object' || Array.isArray(attributes)) change('htmlAttributes', undefined);
        else if (Object.values(attributes).some(v => typeof v !== 'string')) {
            change('htmlAttributes', Object.fromEntries(Object.entries(attributes).map(([k, v]) => [k, asString(v)]).filter(([, v]) => v !== undefined)));
        }
    }
    for (const key of ['children', 'notes', 'comments'] as const) {
        const list = (node as any)[key];
        if (list === undefined || list === null) continue;
        if (!Array.isArray(list)) { change(key, undefined); continue; }
        const fixed = normalizeNodes(list);
        if (fixed !== list) change(key, fixed);
    }
    return next;
}

function normalizeNodes(nodes: OfficeContentNode[]): OfficeContentNode[] {
    let changed = false;
    const out = nodes.filter(node => node && typeof node === 'object' && !Array.isArray(node)).map(node => {
        const fixed = normalizeNode(node);
        if (fixed !== node) changed = true;
        return fixed;
    });
    return changed || out.length !== nodes.length ? out : nodes;
}

/**
 * The AST with every node's text, metadata, formatting and preserved attributes of the type the AST
 * defines: a number where a number goes (a heading level, a list index, a cell position), a string
 * where a string goes, a list of strings for `anchorIds`. A value of another type (from a hand-built
 * AST, or JSON) is coerced where it can be and removed where it cannot, so no writer escapes one kind
 * of value and writes another raw (an array's text went into an attribute unescaped) or fails on it.
 * Returns `ast` itself (same object, `.to()` intact) when everything is well typed; the input is never
 * mutated.
 */
export function withWellTypedValues<T extends OfficeParserAST>(ast: T): T {
    const content = Array.isArray(ast.content) ? normalizeNodes(ast.content) : [];
    const metadata = normalizeDocumentMetadata(ast.metadata);
    return content === ast.content && metadata === ast.metadata ? ast : { ...ast, content, metadata };
}

/** The document's own text fields (written into `<title>`, `<meta>`, core properties) as strings, and its custom properties as plain values or lists of them. */
const DOCUMENT_STRING_FIELDS = ['title', 'author', 'lastModifiedBy', 'description', 'subject', 'keywords', 'language'];

function normalizeDocumentMetadata(metadata: OfficeParserAST['metadata']): OfficeParserAST['metadata'] {
    if (!metadata || typeof metadata !== 'object') return {} as OfficeParserAST['metadata'];
    let out: any;
    const set = (key: string, value: unknown) => {
        out ??= { ...metadata };
        if (value === undefined) delete out[key];
        else out[key] = value;
    };
    for (const key of DOCUMENT_STRING_FIELDS) {
        const value = (metadata as any)[key];
        if (value !== undefined && value !== null && typeof value !== 'string') set(key, asString(value));
    }
    const pages = (metadata as any).pages;
    if (pages !== undefined && pages !== null && typeof pages !== 'number') set('pages', asNumber(pages));
    const formatting = (metadata as any).formatting;
    if (formatting && typeof formatting === 'object' && !Array.isArray(formatting)) {
        const fixed = normalizeRecord(formatting);
        if (fixed !== formatting) set('formatting', fixed);
    }
    // A custom property is a plain value, a date, or a list of plain values (front matter's `tags: [a, b]`).
    const plainValue = (v: unknown) => isPrimitive(v) || v instanceof Date || (Array.isArray(v) && v.every(isPrimitive));
    const custom = (metadata as any).customProperties;
    if (custom && typeof custom === 'object' && !Array.isArray(custom) && !Object.values(custom).every(plainValue)) {
        set('customProperties', Object.fromEntries(Object.entries(custom)
            .map(([k, v]) => [k, plainValue(v) ? v : asString(v)] as const).filter(([, v]) => v !== undefined)));
    }
    return out ?? metadata;
}

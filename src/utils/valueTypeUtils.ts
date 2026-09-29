import { mapNodeLists, NodeMemo, nodeMemo } from './nodeListUtils.js';
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

/** The formatting fields the AST defines (TextFormatting); no writer reads another. */
const FORMATTING_FIELDS = new Set(['bold', 'italic', 'underline', 'strikethrough', 'color', 'backgroundColor', 'size', 'font', 'subscript', 'superscript', 'alignment']);
/** Metadata fields kept, besides the typed ones above, when a record holds more fields than any node's does (MAX_METADATA_FIELDS). */
const OTHER_METADATA_FIELDS = new Set(['anchorIds', 'paragraphIndentation', 'bounds', 'isLineBreak', 'hidden', 'rawText', 'formula', 'value', 'valueType', 'target', 'targetMode']);
/** More fields than any node's metadata holds: past it only the fields the AST defines are kept. */
const MAX_METADATA_FIELDS = 64;
/** The longest string a list is joined into where a string goes (a list there is a mistake, not content). */
const MAX_JOINED_LENGTH = 65536;

/**
 * What one run of withWellTypedValues has read, so that a list or record many nodes share is read once
 * (set for the run's length, which is synchronous: a later run reads the AST afresh, as its caller may
 * have changed it).
 */
interface ReadOnce {
    joinedLists: WeakMap<unknown[], string>;
    stringLists: WeakMap<unknown[], string[]>;
    idLists: WeakMap<unknown[], string[] | null>;
    records: { formatting: WeakMap<object, Record<string, unknown>>; metadata: WeakMap<object, Record<string, unknown>> };
    attributes: WeakMap<object, Record<string, string>>;
}
let readOnce: ReadOnce | undefined;
const newReadOnce = (): ReadOnce => ({
    joinedLists: new WeakMap(), stringLists: new WeakMap(), idLists: new WeakMap(),
    records: { formatting: new WeakMap(), metadata: new WeakMap() }, attributes: new WeakMap(),
});

/**
 * `value` as a string: itself, a number or boolean written out, a list's plain values joined (one level,
 * at most MAX_JOINED_LENGTH characters); else undefined. Flattened eight levels deep, lists holding the
 * same list ten times each made a hundred million values from a handful of arrays.
 */
function asString(value: unknown): string | undefined {
    if (typeof value === 'string') return value;
    if (value instanceof Date) return Number.isNaN(value.getTime()) ? undefined : value.toISOString();
    if (typeof value === 'number' || typeof value === 'boolean') return String(value);
    if (Array.isArray(value)) {
        let joined = readOnce?.joinedLists.get(value);
        if (joined === undefined) {
            joined = '';
            for (let i = 0; i < value.length && joined.length < MAX_JOINED_LENGTH; i++) {
                const item = value[i];
                if (isPrimitive(item)) joined += (joined ? ' ' : '') + String(item);
            }
            joined = joined.slice(0, MAX_JOINED_LENGTH);
            readOnce?.joinedLists.set(value, joined);
        }
        return joined;
    }
    return undefined;
}

/** A list of ids as one list of distinct strings, read once however many records share it. */
function asIds(value: unknown): string[] | undefined {
    if (!Array.isArray(value)) return isPrimitive(value) ? [String(value)] : undefined;
    let ids = readOnce?.idLists.get(value);
    if (ids === undefined) {
        // Each id once: one node naming the same bookmark 60,000 times (a 7 KB DOCX) made a writer
        // mint 60,000 bookmarks for it. Plain values only, one level.
        const seen = new Set<string>();
        let clean = true;
        for (const id of value) {
            if (typeof id !== 'string') clean = false;
            if (!isPrimitive(id)) continue;
            const text = String(id);
            if (seen.has(text)) clean = false;
            else seen.add(text);
        }
        ids = clean ? null : [...seen];
        readOnce?.idLists.set(value, ids);
    }
    return ids === null ? value as string[] : ids.length ? ids : undefined;
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
 * type, or removed when it cannot be; `record` itself when every field is well typed. Formatting keeps
 * only the fields the AST defines (no writer reads another, and a wrapper's formatting is given to every
 * run in it: 16,000 keys on one wrapper of 16,000 runs, 597 KB of JSON, filled the heap). Metadata keeps
 * unknown fields, unless it holds more than MAX_METADATA_FIELDS, which no node's does. Read once
 * however many nodes share it: one record of 20,000 keys given to 20,000 nodes took a minute.
 */
function normalizeRecord(record: Record<string, unknown>, kind: 'formatting' | 'metadata'): Record<string, unknown> {
    const known = readOnce?.records[kind].get(record);
    if (known) return known;
    const result = normalizeRecordOnce(record, kind);
    readOnce?.records[kind].set(record, result);
    return result;
}

function normalizeRecordOnce(record: Record<string, unknown>, kind: 'formatting' | 'metadata'): Record<string, unknown> {
    let out: Record<string, unknown> | undefined;
    const set = (key: string, value: unknown) => {
        out ??= { ...record };
        if (value === undefined) delete out[key];
        else out[key] = value;
    };
    const keys = Object.keys(record);
    const onlyKnown = kind === 'formatting' || keys.length > MAX_METADATA_FIELDS;
    if (onlyKnown && keys.some(key => !(kind === 'formatting' ? FORMATTING_FIELDS.has(key) : STRING_FIELDS.has(key) || NUMBER_FIELDS.has(key) || BOOLEAN_FIELDS.has(key) || OTHER_METADATA_FIELDS.has(key)))) {
        out = {};
        for (const key of keys) {
            if (kind === 'formatting' ? FORMATTING_FIELDS.has(key) : STRING_FIELDS.has(key) || NUMBER_FIELDS.has(key) || BOOLEAN_FIELDS.has(key) || OTHER_METADATA_FIELDS.has(key)) out[key] = record[key];
        }
    }
    const source = out ?? record;
    for (const key of Object.keys(source)) {
        const value = source[key];
        if (value === undefined || value === null) continue;
        if (key === 'anchorIds') {
            const ids = asIds(value);
            if (ids !== value) set(key, ids);
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

/**
 * `node` with well-typed text, metadata, formatting and attributes, and its descendants the same; `node`
 * itself when nothing changes. Each node once, its result shared (`done`): a node the AST shares (one
 * note every reference holds) was read once per path to it, which doubled per level of notes nested in
 * notes, and came out as a copy per reference.
 */
function normalizeNode(node: OfficeContentNode, done: NodeMemo<OfficeContentNode>): OfficeContentNode {
    const known = done.get(node);
    if (known) return known;
    done.set(node, node);
    const result = normalizeNodeOnce(node, done);
    done.set(node, result);
    return result;
}

function normalizeNodeOnce(node: OfficeContentNode, done: NodeMemo<OfficeContentNode>): OfficeContentNode {
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
        const fixed = normalizeRecord(record, key);
        if (fixed !== record) change(key, fixed);
    }
    const attributes = (node as any).htmlAttributes;
    if (attributes !== undefined && attributes !== null) {
        if (typeof attributes !== 'object' || Array.isArray(attributes)) change('htmlAttributes', undefined);
        else {
            const fixed = normalizeAttributes(attributes);
            if (fixed !== attributes) change('htmlAttributes', fixed);
        }
    }
    for (const key of ['children', 'notes', 'comments'] as const) {
        const list = (node as any)[key];
        if (list === undefined || list === null) continue;
        if (!Array.isArray(list)) { change(key, undefined); continue; }
        const fixed = normalizeNodes(list, done);
        if (fixed !== list) change(key, fixed);
    }
    return next;
}

/** Attributes as strings, read once however many nodes share them. */
function normalizeAttributes(attributes: Record<string, unknown>): Record<string, string> {
    let result = readOnce?.attributes.get(attributes);
    if (!result) {
        result = Object.values(attributes).every(v => typeof v === 'string')
            ? attributes as Record<string, string>
            : Object.fromEntries(Object.entries(attributes).map(([k, v]) => [k, asString(v)]).filter(([, v]) => v !== undefined)) as Record<string, string>;
        readOnce?.attributes.set(attributes, result);
    }
    return result;
}

function normalizeNodes(nodes: OfficeContentNode[], done: NodeMemo<OfficeContentNode>): OfficeContentNode[] {
    // Copied at its first change only: a list most passes leave as it is.
    let out: OfficeContentNode[] | undefined;
    for (let i = 0; i < nodes.length; i++) {
        const node = nodes[i];
        if (!node || typeof node !== 'object' || Array.isArray(node)) { out ??= nodes.slice(0, i); continue; }
        const fixed = normalizeNode(node, done);
        if (fixed !== node) out ??= nodes.slice(0, i);
        out?.push(fixed);
    }
    return out ?? nodes;
}

/**
 * The AST with every node's text, metadata, formatting and preserved attributes of the type the AST
 * defines: a number where a number goes (a heading level, a list index, a cell position), a string
 * where a string goes, a list of strings for `anchorIds`. A value of another type (from a hand-built
 * AST, or JSON) is coerced where it can be and removed where it cannot, so no writer escapes one kind
 * of value and writes another raw (an array's text went into an attribute unescaped) or fails on it.
 * Returns `ast` itself (same object, `.to()` intact) when everything is well typed; the input is never
 * mutated. A `tree` (see sharedNodeVisits) is read without a memo.
 */
export function withWellTypedValues<T extends OfficeParserAST>(ast: T, options: { tree?: boolean } = {}): T {
    const seen = nodeMemo<OfficeContentNode>(options.tree);
    readOnce = newReadOnce();
    try {
        const withNodes = mapNodeLists(ast, nodes => normalizeNodes(nodes, seen));
        const metadata = normalizeDocumentMetadata(ast.metadata);
        const attachments = normalizeAttachments(ast.attachments);
        if (metadata === ast.metadata && attachments === ast.attachments) return withNodes;
        return { ...withNodes, metadata, attachments };
    } finally {
        readOnce = undefined;
    }
}

/** The most entries a chart's list of labels, series or values is read with (a chart table writes far fewer). */
const MAX_CHART_ENTRIES = 1_000_000;
/**
 * `value` as a list of strings, one level, at most MAX_CHART_ENTRIES: an entry that is not a plain value
 * is empty. A list many series share is read once (an AST built in code gave 1,000 series one list of
 * 100,000 values, and reading it per series filled the heap); sharedNodeVisits weighs it per series.
 */
const asStrings = (value: unknown): string[] => {
    if (!Array.isArray(value)) return [];
    let strings = readOnce?.stringLists.get(value);
    if (!strings) {
        strings = [];
        for (let i = 0; i < value.length && i < MAX_CHART_ENTRIES; i++) strings.push(isPrimitive(value[i]) ? String(value[i]) : '');
        readOnce?.stringLists.set(value, strings);
    }
    return strings;
};

/**
 * Each attachment's text fields as strings and its chart data as lists of plain strings. A chart's
 * labels and values are written into a table at every chart showing it, and a label that was a list
 * holding a list ten times each, eight levels down, was written out as a hundred million values
 * (200 MB of LaTeX from a handful of arrays).
 */
function normalizeAttachments(attachments: OfficeParserAST['attachments']): OfficeParserAST['attachments'] {
    if (!Array.isArray(attachments)) return [];
    let changed = false;
    const out = attachments.map(attachment => {
        if (!attachment || typeof attachment !== 'object') { changed = true; return undefined; }
        let next: any = attachment;
        const set = (key: string, value: unknown) => {
            if (next === attachment) next = { ...attachment };
            if (value === undefined) delete next[key];
            else next[key] = value;
        };
        for (const key of ['name', 'mimeType', 'type', 'extension', 'altText', 'ocrText'] as const) {
            const value = (attachment as any)[key];
            if (value !== undefined && value !== null && typeof value !== 'string') set(key, asString(value));
        }
        const chart = (attachment as any).chartData;
        if (chart !== undefined && chart !== null) {
            if (typeof chart !== 'object' || Array.isArray(chart)) set('chartData', undefined);
            else {
                const dataSets = Array.isArray(chart.dataSets) ? chart.dataSets.slice(0, MAX_CHART_ENTRIES) : [];
                set('chartData', {
                    ...(chart.title !== undefined && { title: asString(chart.title) }),
                    ...(chart.xAxisTitle !== undefined && { xAxisTitle: asString(chart.xAxisTitle) }),
                    ...(chart.yAxisTitle !== undefined && { yAxisTitle: asString(chart.yAxisTitle) }),
                    dataSets: dataSets.map((series: any) => series && typeof series === 'object'
                        ? { ...(series.name !== undefined && { name: asString(series.name) }), values: asStrings(series.values), pointLabels: asStrings(series.pointLabels) }
                        : { values: [], pointLabels: [] }),
                    labels: asStrings(chart.labels),
                    rawTexts: asStrings(chart.rawTexts),
                });
            }
        }
        if (next !== attachment) changed = true;
        return next;
    }).filter(attachment => attachment !== undefined);
    return changed ? out : attachments;
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
        const fixed = normalizeRecord(formatting, 'formatting');
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

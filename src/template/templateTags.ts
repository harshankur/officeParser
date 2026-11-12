/**
 * The tags a template is written with, apart from where they sit in a document: what a tag's text
 * means (a field, or one of the block tags `if`, `else if`, `else`, `endif`, `for`, `endfor`), the
 * conditions an `if` takes, and how a condition is worked out against the data.
 *
 * A condition is read by a small hand-written parser (no `eval`, no `Function`): comparisons
 * (`== != < <= > >=`, and `<>` for `!=`, as spreadsheets write it), `and`, `or`, `not`, parentheses,
 * data paths, numbers, quoted strings, `true`, `false` and `null`. A value may go through filters before
 * it is compared, as a field's does (`answer | lower == "yes"`). Keywords are case-insensitive, so a tag Word's AutoCorrect gave a
 * capital (`{{ If paid }}`) still reads, and a string may be quoted with straight or curly quotes,
 * since AutoCorrect curls quotes typed in Word.
 *
 * @module template/templateTags
 */

import { TemplateValue } from '../types.js';

/** A data path: `customer.address.city` is `['customer', 'address', 'city']`. */
export type TemplatePath = string[];

/** The comparison operators a condition takes. */
export type CompareOperator = '==' | '!=' | '<' | '<=' | '>' | '>=';

/** A parsed condition. */
export type TemplateExpression =
    | { kind: 'path'; path: TemplatePath; text: string }
    | { kind: 'literal'; value: string | number | boolean | null }
    | { kind: 'not'; operand: TemplateExpression }
    /** `a and b and c` holds its operands side by side, so a long chain is worked out in a loop, not by recursion. */
    | { kind: 'and' | 'or'; operands: TemplateExpression[] }
    | { kind: 'compare'; operator: CompareOperator; left: TemplateExpression; right: TemplateExpression }
    /**
     * A value through filters, `answer | lower`: the value (a path or a literal), its filters as written, and
     * the chain's place among the tag's filtered values in the order written, from 0, by which the engine
     * finds the chain it compiled once.
     */
    | { kind: 'filtered'; operand: Extract<TemplateExpression, { kind: 'path' | 'literal' }>; filters: TemplateFilter[]; chain: number };

/** What a tag's text means. */
export type TemplateTag =
    | { kind: 'field'; path: TemplatePath; text: string; filters?: TemplateFilter[] }
    | { kind: 'if'; condition: TemplateExpression }
    | { kind: 'elseif'; condition: TemplateExpression }
    | { kind: 'else' }
    | { kind: 'endif' }
    | { kind: 'for'; name: string; list: TemplatePath; listText: string }
    | { kind: 'endfor' };

/** A filter a field's value, or a value in a condition, goes through, written after `|`: its name, and the values written after it. */
export interface TemplateFilter { name: string; args: (string | number)[] }

/** The block tags, as a tag's kind. */
export type TemplateBlockTagKind = Exclude<TemplateTag['kind'], 'field'>;

/** A tag read from its text: what it is, or why it cannot be read. `null` text is not a tag (prose). */
export type TagReading = { ok: true; tag: TemplateTag } | { ok: false; reason: string; misspelled?: true; respelt?: string } | null;

/** Block keywords as other template languages spell them, which mean a block tag here, misspelt. */
const MISSPELLED = new Map([['elseif', 'else if'], ['elsif', 'else if'], ['elif', 'else if']]); // a Map: `constructor` is no misspelling
const MISSPELLED_END = new Set(['if', 'for']);

/** The name a loop gives its own details (`loop.index`, `loop.first`, ...). */
export const LOOP_DETAILS_NAME = 'loop';

/** How deeply parentheses and `not` may nest in one condition. */
export const MAX_CONDITION_DEPTH = 32;

/** Keywords of the tag language, matched case-insensitively. */
const KEYWORD = {
    if: 'if', else: 'else', endif: 'endif', for: 'for', in: 'in', endfor: 'endfor',
    and: 'and', or: 'or', not: 'not', true: 'true', false: 'false', null: 'null',
} as const;

/** The block tags with nothing after their word, shared: a template may hold a great many. */
const ELSE: TemplateTag = Object.freeze({ kind: 'else' });
const ENDIF: TemplateTag = Object.freeze({ kind: 'endif' });
const ENDFOR: TemplateTag = Object.freeze({ kind: 'endfor' });

/** Words that cannot name a field, a loop item or the first step of a path. */
const RESERVED = new Set<string>(Object.values(KEYWORD));

/** A field's text as before block tags existed: Unicode letters and digits plus `_ . -`. */
const FIELD_TEXT = /^[\p{L}\p{N}_.\-]+$/u;

/** One character of a name or a number: a path is a run of these, split at its dots. */
const NAME_CHAR = /[\p{L}\p{N}_.\-]/u;

/** A run of name characters that is a number. */
const NUMBER_TEXT = /^-?\d+(?:\.\d+)?$/;

/** One step of a path (no dots). */
const STEP_TEXT = /^[\p{L}\p{N}_\-]+$/u;

/**
 * Quote characters, by family: a string opens with any of a family and closes at the next of the
 * same family, so it does not matter which way AutoCorrect curled either quote (Word curls a quote
 * typed after `=` as a closing one). Double: straight, curly, low-9, reversed-9, guillemets.
 * Single: straight, curly, low-9, reversed-9, single guillemets.
 */
export const DOUBLE_QUOTES = '"\u201C\u201D\u201E\u201F\u00AB\u00BB';
export const SINGLE_QUOTES = '\'\u2018\u2019\u201A\u201B\u2039\u203A';
/** The guillemets, inside which French AutoCorrect adds a no-break space that is not the string's own. */
const GUILLEMETS = '\u00AB\u00BB\u2039\u203A';
/** The no-break spaces French AutoCorrect puts inside guillemets. */
const GUILLEMET_SPACES = '\u00A0\u202F';

/**
 * A string without the no-break spaces at its ends, trimmed by index: a pattern anchored at the end
 * (`[...]+$`) retries from every position of a long run of them, in time quadratic in its length.
 */
function trimGuillemetSpaces(value: string): string {
    let start = 0;
    let end = value.length;
    while (start < end && GUILLEMET_SPACES.includes(value[start])) start++;
    while (end > start && GUILLEMET_SPACES.includes(value[end - 1])) end--;
    return value.slice(start, end);
}

/** Comparison operators as typed, longest first, with the symbols Word's math AutoCorrect can write; `<>` is `!=`, as spreadsheets write it. */
const OPERATORS: [string, CompareOperator][] = [
    ['==', '=='], ['!=', '!='], ['<=', '<='], ['>=', '>='], ['<>', '!='], ['<', '<'], ['>', '>'],
    ['\u2264', '<='], ['\u2265', '>='], ['\u2260', '!='],
];

/**
 * Operators of other languages a condition does not take, as typed, each with what to write instead: checked
 * before {@link OPERATORS}, which would read their start (`==` of `===`) and blame what follows. `||` is
 * among them, so it is named before a single `|` is read as a filter's.
 */
const MISTYPED: [string, string][] = [
    ['===', 'use == to compare'], ['!==', 'use != to compare'],
    ['=>', 'use >= to compare'], ['=<', 'use <= to compare'], ['=!', 'use != to compare'],
    ['&&', 'write and to join conditions'], ['||', 'write or to join conditions'],
];

type Token =
    | { t: 'name'; text: string }
    | { t: 'number'; value: number }
    | { t: 'string'; value: string }
    | { t: 'op'; op: CompareOperator }
    | { t: '(' } | { t: ')' } | { t: '|' };

type Tokens = { ok: true; tokens: Token[] } | { ok: false; reason: string };

/** Splits a condition's text into tokens, in one forward pass. */
function tokenize(text: string): Tokens {
    const tokens: Token[] = [];
    for (let at = 0; at < text.length;) {
        const ch = text[at];
        if (/\s/u.test(ch)) { at++; continue; }
        if (ch === '(' || ch === ')') { tokens.push({ t: ch }); at++; continue; }
        const mistyped = MISTYPED.find(([typed]) => text.startsWith(typed, at));
        if (mistyped) return { ok: false, reason: `uses ${mistyped[0]}; ${mistyped[1]}` };
        const operator = OPERATORS.find(([typed]) => text.startsWith(typed, at));
        if (operator) { tokens.push({ t: 'op', op: operator[1] }); at += operator[0].length; continue; }
        if (ch === '=') return { ok: false, reason: 'compares with a single =; use == to compare' };
        if (ch === '!') return { ok: false, reason: 'uses !; write not before a condition, or != to compare' };
        if (ch === '&') return { ok: false, reason: 'uses &; write and to join conditions' };
        if (ch === '|') { tokens.push({ t: '|' }); at++; continue; }
        const family = DOUBLE_QUOTES.includes(ch) ? DOUBLE_QUOTES : SINGLE_QUOTES.includes(ch) ? SINGLE_QUOTES : null;
        if (family) {
            let close = at + 1;
            while (close < text.length && !family.includes(text[close])) close++;
            if (close >= text.length) return { ok: false, reason: 'has a string with no closing quote' };
            let value = text.slice(at + 1, close);
            if (GUILLEMETS.includes(ch)) value = trimGuillemetSpaces(value);
            tokens.push({ t: 'string', value });
            at = close + 1;
            continue;
        }
        if (NAME_CHAR.test(String.fromCodePoint(text.codePointAt(at)!))) {
            // By code point, so a letter outside the BMP (two UTF-16 units) is one name character, as in a field.
            let end = at;
            while (end < text.length) {
                const point = String.fromCodePoint(text.codePointAt(end)!);
                if (!NAME_CHAR.test(point)) break;
                end += point.length;
            }
            const run = text.slice(at, end);
            tokens.push(NUMBER_TEXT.test(run) ? { t: 'number', value: Number(run) } : { t: 'name', text: run });
            at = end;
            continue;
        }
        return { ok: false, reason: `has a character a condition does not use: ${ch}` };
    }
    return { ok: true, tokens };
}

/** A path's steps from its text, or null when a step is empty (`a..b`, `.a`) or not a name. */
function readPath(text: string): TemplatePath | null {
    const steps = text.split('.');
    return steps.every(step => STEP_TEXT.test(step)) ? steps : null;
}

/** Parses a condition's tokens (recursive descent, nesting bounded by {@link MAX_CONDITION_DEPTH}). */
function parseCondition(tokens: Token[]): { ok: true; expression: TemplateExpression } | { ok: false; reason: string } {
    let at = 0;
    // How many filtered values have been read: each one's place, by which the engine finds its compiled chain.
    let chains = 0;
    let failure: string | null = null;
    const fail = (reason: string): null => { failure ??= reason; return null; };
    const keyword = (token: Token | undefined, word: string): boolean => token?.t === 'name' && token.text.toLowerCase() === word;

    const primary = (depth: number): TemplateExpression | null => {
        const token = tokens[at];
        if (!token) return fail('ends before its condition does');
        if (token.t === '(') {
            if (depth >= MAX_CONDITION_DEPTH) return fail(`nests parentheses or not more than ${MAX_CONDITION_DEPTH} deep`);
            at++;
            const inner = or(depth + 1);
            if (!inner) return null;
            if (tokens[at]?.t !== ')') return fail('has a ( with no closing )');
            at++;
            return inner;
        }
        if (token.t === 'number') { at++; return { kind: 'literal', value: token.value }; }
        if (token.t === 'string') { at++; return { kind: 'literal', value: token.value }; }
        if (token.t === 'name') {
            const word = token.text.toLowerCase();
            if (word === KEYWORD.true || word === KEYWORD.false) { at++; return { kind: 'literal', value: word === KEYWORD.true }; }
            if (word === KEYWORD.null) { at++; return { kind: 'literal', value: null }; }
            if (RESERVED.has(word)) return fail(`uses ${token.text} where a value belongs`);
            const path = readPath(token.text);
            if (!path) return fail(`has a path that is not one: ${token.text}`);
            if (RESERVED.has(path[0].toLowerCase())) return fail(`starts a path with the word ${path[0]}`);
            at++;
            return { kind: 'path', path, text: token.text };
        }
        if (token.t === ')') return fail('has a ) with no opening (');
        if (token.t === '|') return fail('has | with nothing before it to filter');
        return fail('has a comparison with nothing before it');
    };
    /**
     * A value and the filters written after it (`answer | lower`, `due | format "yyyy"`): each filter a name
     * and the strings or numbers after it. Filters follow a path or a literal, never parentheses. Each token
     * is looked at once; filters do not nest.
     */
    const filtered = (depth: number): TemplateExpression | null => {
        const opening = tokens[at];
        const value = primary(depth);
        if (!value || tokens[at]?.t !== '|') return value;
        if (opening?.t === '(' || (value.kind !== 'path' && value.kind !== 'literal')) return fail('puts a filter after parentheses; write it after the value inside them, as (answer | lower)');
        const filters: TemplateFilter[] = [];
        while (tokens[at]?.t === '|') {
            at++;
            const name = tokens[at];
            if (name?.t !== 'name' || !STEP_TEXT.test(name.text) || RESERVED.has(name.text.toLowerCase())) return fail('has | with no filter\'s name after it');
            at++;
            const args: (string | number)[] = [];
            for (let token = tokens[at]; token?.t === 'string' || token?.t === 'number'; token = tokens[++at]) args.push(token.value);
            filters.push({ name: name.text, args });
        }
        return { kind: 'filtered', operand: value, filters, chain: chains++ };
    };
    const compare = (depth: number): TemplateExpression | null => {
        const left = filtered(depth);
        if (!left) return null;
        const token = tokens[at];
        if (token?.t !== 'op') return left;
        at++;
        const right = filtered(depth);
        if (!right) return null;
        if (tokens[at]?.t === 'op') return fail('chains comparisons; join them with and');
        return { kind: 'compare', operator: token.op, left, right };
    };
    const not = (depth: number): TemplateExpression | null => {
        if (!keyword(tokens[at], KEYWORD.not)) return compare(depth);
        if (depth >= MAX_CONDITION_DEPTH) return fail(`nests parentheses or not more than ${MAX_CONDITION_DEPTH} deep`);
        at++;
        const operand = not(depth + 1);
        return operand && { kind: 'not', operand };
    };
    const and = (depth: number): TemplateExpression | null => {
        const first = not(depth);
        if (!first || !keyword(tokens[at], KEYWORD.and)) return first;
        const operands = [first];
        while (keyword(tokens[at], KEYWORD.and)) {
            at++;
            const next = not(depth);
            if (!next) return null;
            operands.push(next);
        }
        return { kind: 'and', operands };
    };
    function or(depth: number): TemplateExpression | null {
        const first = and(depth);
        if (!first || !keyword(tokens[at], KEYWORD.or)) return first;
        const operands = [first];
        while (keyword(tokens[at], KEYWORD.or)) {
            at++;
            const next = and(depth);
            if (!next) return null;
            operands.push(next);
        }
        return { kind: 'or', operands };
    }

    if (!tokens.length) return { ok: false, reason: 'has no condition' };
    const expression = or(0);
    if (!expression) return { ok: false, reason: failure ?? 'cannot be read' };
    if (at < tokens.length) return { ok: false, reason: 'has more after its condition than a condition takes' };
    return { ok: true, expression };
}

/** The first word of a tag's text and what follows it. */
function splitFirstWord(text: string): [string, string] {
    const match = /^(\p{L}+)([\s\S]*)$/u.exec(text);
    return match ? [match[1].toLowerCase(), match[2].trim()] : ['', text];
}

/** Splits text at each `|` outside a quoted string; null when a quote is never closed. */
function splitAtPipes(text: string): string[] | null {
    const parts: string[] = [];
    let start = 0;
    for (let at = 0; at < text.length; at++) {
        const ch = text[at];
        const family = DOUBLE_QUOTES.includes(ch) ? DOUBLE_QUOTES : SINGLE_QUOTES.includes(ch) ? SINGLE_QUOTES : null;
        if (family) {
            let close = at + 1;
            while (close < text.length && !family.includes(text[close])) close++;
            if (close >= text.length) return null;
            at = close;
        } else if (ch === '|') {
            parts.push(text.slice(start, at));
            start = at + 1;
        }
    }
    parts.push(text.slice(start));
    return parts;
}

/**
 * A field with filters, `path | name "argument" | name`: each filter a name and the strings or numbers
 * after it. Text not of that shape is prose and stays literal text (null), as text between the
 * delimiters always has; whether each name is a filter is decided where the template is compiled.
 */
function readFilteredField(text: string): TagReading {
    const parts = splitAtPipes(text);
    if (!parts) return null;
    const head = parts[0].trim();
    if (!FIELD_TEXT.test(head) || RESERVED.has(head.toLowerCase())) return null;
    const filters: TemplateFilter[] = [];
    for (const part of parts.slice(1)) {
        const tokens = tokenize(part);
        if (!tokens.ok || !tokens.tokens.length) return null;
        const [name, ...rest] = tokens.tokens;
        if (name.t !== 'name' || !STEP_TEXT.test(name.text)) return null;
        const args: (string | number)[] = [];
        for (const token of rest) {
            if (token.t === 'string' || token.t === 'number') args.push(token.value);
            else return null;
        }
        filters.push({ name: name.text, args });
    }
    return { ok: true, tag: { kind: 'field', path: readPath(head) ?? [head], text: head, filters } };
}

/**
 * Reads a tag's text: the characters between the delimiters, XML entities already decoded. Returns
 * null for text that is not a tag (prose between braces stays literal text, as it always has), and a
 * failure for a tag that starts like a block tag but cannot be read.
 */
export function readTag(rawText: string): TagReading {
    const text = rawText.trim();
    if (!text) return null;
    if (FIELD_TEXT.test(text)) {
        const word = text.toLowerCase();
        if (word === KEYWORD.else) return { ok: true, tag: ELSE };
        if (word === KEYWORD.endif) return { ok: true, tag: ENDIF };
        if (word === KEYWORD.endfor) return { ok: true, tag: ENDFOR };
        if (word === KEYWORD.if || word === KEYWORD.for) return { ok: false, reason: word === KEYWORD.if ? 'has no condition' : 'does not say what to loop over' };
        // As before block tags existed, a field is any such text; a path that is not one (`a..b`)
        // still reads as a field, found only by its exact key.
        return { ok: true, tag: { kind: 'field', path: readPath(text) ?? [text], text } };
    }
    const [word, rest] = splitFirstWord(text);
    if (text.includes('|') && word !== KEYWORD.if && word !== KEYWORD.else && word !== KEYWORD.for) return readFilteredField(text);
    if (word === KEYWORD.if || word === KEYWORD.else) {
        let conditionText = rest;
        if (word === KEYWORD.else) {
            const [next, afterIf] = splitFirstWord(rest);
            if (!rest) return { ok: true, tag: ELSE };
            if (next !== KEYWORD.if) return { ok: false, reason: 'has more after else than else if takes' };
            conditionText = afterIf;
        }
        const tokens = tokenize(conditionText);
        if (!tokens.ok) return tokens;
        const parsed = parseCondition(tokens.tokens);
        if (!parsed.ok) return parsed;
        return { ok: true, tag: word === KEYWORD.if ? { kind: 'if', condition: parsed.expression } : { kind: 'elseif', condition: parsed.expression } };
    }
    if (word === KEYWORD.for) {
        const match = /^(\S+)\s+(\S+)\s+(\S+)$/u.exec(rest);
        if (!match || match[2].toLowerCase() !== KEYWORD.in) return { ok: false, reason: 'is not written as for item in list' };
        const [, name, , listText] = match;
        if (!STEP_TEXT.test(name)) return { ok: false, reason: `names its item ${name}, which is not a name` };
        if (RESERVED.has(name.toLowerCase())) return { ok: false, reason: `names its item with the word ${name}` };
        if (name === LOOP_DETAILS_NAME) return { ok: false, reason: `names its item ${LOOP_DETAILS_NAME}, the name of the loop's own details` };
        const list = readPath(listText);
        if (!list || RESERVED.has(list[0].toLowerCase())) return { ok: false, reason: `loops over ${listText}, which is not a path` };
        return { ok: true, tag: { kind: 'for', name, list, listText } };
    }
    if (word === KEYWORD.endif || word === KEYWORD.endfor) return { ok: false, reason: `has more after ${word} than ${word} takes` };
    // `elseif`, `elif`, `end if`: a block tag as Jinja, Handlebars or Liquid spell it. Read as prose, the
    // branch it opens vanished without a word; it is a mistake to name at once.
    // `respelt` is the tag's text as it should be typed: the keyword respelt, the rest as typed.
    const spelling = MISSPELLED.get(word);
    if (spelling) return { ok: false, misspelled: true, reason: `is written ${word}; a branch is written ${spelling}`, respelt: rest ? `${spelling} ${rest}` : spelling };
    if (word === 'end') {
        const [next] = splitFirstWord(rest);
        if (MISSPELLED_END.has(next) && rest.trim().toLowerCase() === next) return { ok: false, misspelled: true, reason: `is written end ${next}; a block is closed with end${next}`, respelt: `end${next}` };
    }
    return null;
}

/** Stands for a field the data does not have (distinct from a field that is null). */
export const MISSING: unique symbol = Symbol('missing');

/** A value as a condition sees it: the data's value, or {@link MISSING}. */
export type ConditionValue = TemplateValue | typeof MISSING;

/** Why two values cannot be compared, for the caller's typed error. */
export type Incomparable = { left: string; right: string };

/** A value's kind, as an error message names it. */
export function kindOfValue(value: unknown): string {
    if (value === null || value === undefined) return 'empty';
    if (Array.isArray(value)) return 'list';
    if (value instanceof Date) return 'date';
    if (typeof value === 'string') return 'text';
    if (typeof value === 'boolean') return 'true/false';
    if (typeof value === 'number' || typeof value === 'bigint') return 'number';
    if (typeof value === 'function') return 'function';
    if (typeof value === 'symbol') return 'symbol';
    return 'object';
}

/** Whether a value counts as true: false, null, empty text, 0, NaN, an empty list and an invalid date do not. */
export function isTruthy(value: TemplateValue): boolean {
    if (value === null || value === undefined || value === false || value === '' || value === 0) return false;
    if (typeof value === 'number') return !Number.isNaN(value);
    if (Array.isArray(value)) return value.length > 0;
    if (value instanceof Date) return !Number.isNaN(value.getTime());
    return true;
}

/** Whether two values of one kind are equal; a bigint and a number by their value (1n and 1 are the same number). */
function same(a: unknown, b: unknown): boolean {
    if (typeof a === 'bigint' && typeof b === 'number') return Number.isInteger(b) && a === BigInt(b);
    if (typeof a === 'number' && typeof b === 'bigint') return Number.isInteger(a) && BigInt(a) === b;
    return a === b;
}

/** What working out a condition needs from the data and the engine. */
export interface ConditionReader {
    /** A path's value: the data's, or {@link MISSING} (or what the caller's `onMissing` gives it). */
    lookup(path: TemplatePath, text: string): ConditionValue;
    /** Raises the caller's typed error for two values that cannot be compared. */
    incomparable(why: Incomparable): never;
    /** Told how much of two strings a comparison reads, so it can be counted. */
    compared?(length: number): void;
    /** Text compared with a number, read as one; null when it is none. */
    readNumber?(text: string): number | null;
    /** A value through the tag's filter chain `chain` (see the `filtered` expression). */
    filter(chain: number, value: TemplateValue): TemplateValue;
    /** Two texts compared: the answer when the reader worked it out (as numbers, or as text, once, when it read them to report them), undefined to compare them as text. */
    compareTexts?(operator: CompareOperator, left: string, right: string): boolean | undefined;
}

/**
 * Compares two values. Null (or undefined) equals only null and is never ordered. Otherwise both
 * must be the same kind (text, number, true/false, date), save that text compared with a number is
 * read as a number (`readNumber`), when it is one; anything else is reported through `incomparable`,
 * since a silent false would pick the wrong branch without a word. Two texts compare character by
 * character, unless `compareTexts` answers (two texts that both read as numbers, compared as numbers).
 */
function compareValues(operator: CompareOperator, left: TemplateValue, right: TemplateValue, reader: ConditionReader): boolean {
    const leftEmpty = left === null || left === undefined;
    const rightEmpty = right === null || right === undefined;
    if (leftEmpty || rightEmpty) {
        if (operator === '==') return leftEmpty && rightEmpty;
        if (operator === '!=') return !(leftEmpty && rightEmpty);
        return false;
    }
    let leftKind = kindOfValue(left);
    let rightKind = kindOfValue(right);
    // Text that is a number, beside a number, is that number: data from a form or a table is all text.
    if (reader.readNumber && leftKind === 'text' && rightKind === 'number') {
        const n = reader.readNumber(left as string);
        if (n !== null) { left = n; leftKind = 'number'; }
    } else if (reader.readNumber && leftKind === 'number' && rightKind === 'text') {
        const n = reader.readNumber(right as string);
        if (n !== null) { right = n; rightKind = 'number'; }
    }
    if (leftKind !== rightKind || leftKind === 'list' || leftKind === 'object' || leftKind === 'function' || leftKind === 'symbol') reader.incomparable({ left: leftKind, right: rightKind });
    if (typeof left === 'string' && typeof right === 'string') {
        const decided = reader.compareTexts?.(operator, left, right);
        if (decided !== undefined) return decided;
        reader.compared?.(Math.min(left.length, right.length));
    }
    const a = left instanceof Date ? left.getTime() : left as string | number | boolean;
    const b = right instanceof Date ? right.getTime() : right as string | number | boolean;
    switch (operator) {
        case '==': return same(a, b);
        case '!=': return !same(a, b);
        case '<': return a < b;
        case '<=': return a <= b;
        case '>': return a > b;
        case '>=': return a >= b;
    }
}

/** A value of a condition: a literal, a path, or one of them through filters. */
type ConditionOperand = Extract<TemplateExpression, { kind: 'path' | 'literal' | 'filtered' }>;

/** Whether a condition is a value (not a comparison, nor one joined or negated). */
function isOperand(expression: TemplateExpression): expression is ConditionOperand {
    return expression.kind === 'path' || expression.kind === 'literal' || expression.kind === 'filtered';
}

/**
 * A value as a condition reads it, beside an operator or alone: a literal's, a path's (MISSING when the
 * data does not have it), or a value through its filters (a missing value goes through none).
 */
function valueOf(expression: ConditionOperand, reader: ConditionReader): ConditionValue {
    if (expression.kind === 'literal') return expression.value;
    if (expression.kind === 'path') return reader.lookup(expression.path, expression.text);
    const value = valueOf(expression.operand, reader);
    return value === MISSING ? MISSING : reader.filter(expression.chain, value);
}

/**
 * Works out a condition. `reader.lookup` reads a path from the data (returning {@link MISSING} for a
 * field the data does not have, or the value the caller's `onMissing` gives it). A missing value makes
 * the whole condition missing, so the block it heads can be kept as typed. `and` and `or` stop as soon
 * as their answer is known, as they do in JavaScript. A filtered value goes through `reader.filter`.
 */
export function evaluateCondition(expression: TemplateExpression, reader: ConditionReader): boolean | typeof MISSING {
    switch (expression.kind) {
        case 'literal':
        case 'path':
        case 'filtered': {
            const value = valueOf(expression, reader);
            return value === MISSING ? MISSING : isTruthy(value);
        }
        case 'not': {
            const value = evaluateCondition(expression.operand, reader);
            return value === MISSING ? MISSING : !value;
        }
        case 'and':
        case 'or': {
            // `and` stops at the first false operand, `or` at the first true one.
            const stopAt = expression.kind === 'or';
            for (const operand of expression.operands) {
                const value = evaluateCondition(operand, reader);
                if (value === MISSING) return MISSING;
                if (value === stopAt) return stopAt;
            }
            return !stopAt;
        }
        case 'compare': {
            // The parser only puts values beside an operator; a parenthesised condition is true/false.
            const operand = (side: TemplateExpression): ConditionValue => isOperand(side) ? valueOf(side, reader) : evaluateCondition(side, reader) as ConditionValue;
            const left = operand(expression.left);
            if (left === MISSING) return MISSING;
            const right = operand(expression.right);
            if (right === MISSING) return MISSING;
            return compareValues(expression.operator, left, right, reader);
        }
    }
}

/** How much work a condition is to work out, counted once: one per value, operator and path step. */
export function conditionSize(expression: TemplateExpression): number {
    switch (expression.kind) {
        case 'literal': return 1;
        // Reading a path compares its text with names and keys: as long as its text.
        case 'path': return 1 + expression.text.length;
        case 'not': return 1 + conditionSize(expression.operand);
        case 'and':
        case 'or': return expression.operands.reduce((total, operand) => total + conditionSize(operand), 1);
        case 'compare': return 1 + conditionSize(expression.left) + conditionSize(expression.right);
        // Each filter is one more step, and its name and arguments are read: as long as their text.
        case 'filtered': {
            let size = 1 + conditionSize(expression.operand);
            for (const filter of expression.filters) {
                size += 1 + filter.name.length;
                for (const arg of filter.args) size += String(arg).length;
            }
            return size;
        }
    }
}

/** A condition's filter chains, by their place (see the `filtered` expression): each chain's filters as written. Iterative. */
export function conditionFilters(expression: TemplateExpression): TemplateFilter[][] {
    const chains: TemplateFilter[][] = [];
    const todo: TemplateExpression[] = [expression];
    while (todo.length) {
        const e = todo.pop()!;
        switch (e.kind) {
            case 'path': case 'literal': break;
            case 'filtered': chains[e.chain] = e.filters; break;
            case 'not': todo.push(e.operand); break;
            case 'and': case 'or': for (const operand of e.operands) todo.push(operand); break;
            case 'compare': todo.push(e.left, e.right); break;
        }
    }
    return chains;
}

/** The filters a condition's values go through, by name, each once, in the order first written. */
export function conditionFilterNames(expression: TemplateExpression): string[] {
    const names = new Set<string>();
    for (const chain of conditionFilters(expression)) for (const filter of chain) names.add(filter.name);
    return [...names];
}

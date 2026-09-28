/**
 * The content one document repeats by reference, bounded in all (`decompressionLimits.maxRepeatedContent`).
 *
 * A document can state something once and use it many times: an ODF cell repeated with
 * `table:number-columns-repeated`, an XLSX shared string shown in many cells, a style's font or a
 * relationship's link target given to every run naming it, a comment author named by id. The AST
 * shares the value, so it stays small, but every generator writes it once per use: a 1.7 KB DOCX
 * giving one 64 KB font to 2,000 runs made 131 MB of HTML, and 2.7 KB of XLSX showing one 1 MB
 * string in 400 cells made 400 MB of CSV. Each use after the first is charged here.
 */

import { OfficeContentNode, OfficeParserConfig, OfficeWarningType } from '../types.js';
import { logWarning } from './errorUtils.js';

/**
 * What a copy may weigh without charge: short values repeat in every real document (a category in
 * a thousand cells, a font on every run), and the cell and element limits bound how often.
 */
export const REPEAT_ALLOWANCE = 64;
const DEFAULT_MAX_REPEATED_CONTENT = 16 * 1024 * 1024;

/** What each parse (its config object) may still repeat. */
const budgets = new WeakMap<object, { left: number; warned: boolean }>();
const budgetOf = (config: OfficeParserConfig): { left: number; warned: boolean } => {
    let budget = budgets.get(config);
    if (!budget) budgets.set(config, budget = { left: config.decompressionLimits?.maxRepeatedContent ?? DEFAULT_MAX_REPEATED_CONTENT, warned: false });
    return budget;
};

/**
 * How many of `copies` further copies of content weighing `weight` (characters, plus 16 for each node
 * a copy holds) the parse may still make: each costs what it weighs past REPEAT_ALLOWANCE. Warns once
 * when it refuses any.
 */
export const takeRepeats = (config: OfficeParserConfig, copies: number, weight: number): number => {
    if (!(copies > 0)) return 0;
    const cost = weight - REPEAT_ALLOWANCE;
    if (!(cost > 0)) return copies;
    const budget = budgetOf(config);
    const granted = Math.max(0, Math.min(copies, Math.floor(budget.left / cost)));
    budget.left -= granted * cost;
    if (granted < copies && !budget.warned) {
        budget.warned = true;
        logWarning(OfficeWarningType.REPEATED_CONTENT_LIMIT_EXCEEDED, config, config.decompressionLimits?.maxRepeatedContent ?? DEFAULT_MAX_REPEATED_CONTENT);
    }
    return granted;
};

/** The start of `text` a copy past the budget keeps: REPEAT_ALLOWANCE characters, not splitting a surrogate pair. */
export const repeatPreview = (text: string): string => {
    if (text.length <= REPEAT_ALLOWANCE) return text;
    const code = text.charCodeAt(REPEAT_ALLOWANCE - 1);
    return text.slice(0, code >= 0xD800 && code <= 0xDBFF ? REPEAT_ALLOWANCE - 1 : REPEAT_ALLOWANCE) + '…';
};

/** Formatting values a style gives every run it applies to. */
const FORMATTING_VALUES = ['font', 'color', 'backgroundColor', 'size'] as const;
/**
 * Metadata values a definition gives every node naming it (a relationship's link target, a style name,
 * a comment author). Identities (`noteId`, `commentId`, `attachmentName`, `url`, anchors) are not
 * values: without them a reference breaks.
 */
const METADATA_VALUES = ['link', 'linkTitle', 'title', 'style', 'altText', 'author', 'initials', 'date', 'abbreviationTitle', 'backgroundColor', 'language', 'citationKey', 'label'] as const;

/**
 * Charges each formatting or metadata value a node shares with an earlier one, past REPEAT_ALLOWANCE
 * characters; once the budget is spent, a node repeating a long value goes without it (its own copy of
 * the holder, so nodes sharing the holder keep theirs). A node reached twice (shared) counts once.
 */
export const boundRepeatedValues = (roots: (OfficeContentNode[] | undefined)[], config: OfficeParserConfig): void => {
    const seenNodes = new Set<OfficeContentNode>();
    const seenValues = new Set<string>();
    const bound = (node: OfficeContentNode, field: 'formatting' | 'metadata', keys: readonly string[]): void => {
        const holder = node[field] as Record<string, unknown> | undefined;
        if (!holder || typeof holder !== 'object') return;
        let own: Record<string, unknown> | undefined;
        for (const key of keys) {
            const value = holder[key];
            if (typeof value !== 'string' || value.length <= REPEAT_ALLOWANCE) continue;
            if (!seenValues.has(value)) { seenValues.add(value); continue; }
            if (takeRepeats(config, 1, value.length) > 0) continue;
            own ??= { ...holder };
            delete own[key];
        }
        if (own) (node as any)[field] = own;
    };
    const stack: OfficeContentNode[] = [];
    for (const root of roots) if (root) for (let i = root.length - 1; i >= 0; i--) stack.push(root[i]);
    while (stack.length) {
        const node = stack.pop()!;
        if (!node || typeof node !== 'object' || seenNodes.has(node)) continue;
        seenNodes.add(node);
        bound(node, 'formatting', FORMATTING_VALUES);
        bound(node, 'metadata', METADATA_VALUES);
        for (const list of [node.comments, node.notes, node.children]) {
            if (Array.isArray(list)) for (let i = list.length - 1; i >= 0; i--) stack.push(list[i]);
        }
    }
};

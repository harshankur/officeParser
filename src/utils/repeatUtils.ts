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

import { OfficeAttachment, OfficeContentNode, OfficeParserConfig, OfficeWarningType } from '../types.js';
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

/** The characters of the string values an object holds directly (a node's formatting or metadata). */
export const valuesLength = (values: object | undefined): number => {
    let length = 0;
    if (values) for (const key in values) {
        const value = (values as Record<string, unknown>)[key];
        if (typeof value === 'string') length += value.length;
    }
    return length;
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

/** The longest attachment name kept as it is (a longer one is shortened; see boundRepeatedValues). */
const MAX_ATTACHMENT_NAME = 128;

/**
 * Bounds what a document repeats by reference, over everything a parser built:
 *
 * - Each formatting or metadata value a node shares with an earlier one is charged past
 *   REPEAT_ALLOWANCE characters; once the budget is spent, a node repeating a long value goes without
 *   it (on its own copy of the holder, so nodes sharing the holder keep theirs).
 * - The text a node takes from its attachment (a chart's data, a picture's recognized text) is whole
 *   on the first node showing that attachment and charged on the others, past the budget their start:
 *   one 100 KB chart framed 2,000 times made 200 MB of text.
 * - An attachment name longer than MAX_ATTACHMENT_NAME is shortened, on the attachment and on every
 *   node naming it alike: it is written at every picture showing it, and one relationship target of
 *   100 KB shown 2,000 times made 200 MB.
 *
 * A node reached twice (shared) counts once.
 */
export const boundRepeatedValues = (roots: (OfficeContentNode[] | undefined)[], attachments: OfficeAttachment[], config: OfficeParserConfig): void => {
    const seenNodes = new Set<OfficeContentNode>();
    const seenValues = new Set<string>();
    const shownAttachments = new Set<string>();
    const shortNames = new Map<string, string>();
    const shortName = (name: string): string => {
        let short = shortNames.get(name);
        if (short === undefined) {
            const extension = /\.[A-Za-z0-9]{1,10}$/.exec(name)?.[0] ?? '';
            short = `${name.slice(0, 64)}~${shortNames.size + 1}${extension}`;
            shortNames.set(name, short);
        }
        return short;
    };
    for (const attachment of attachments) {
        if (typeof attachment?.name === 'string' && attachment.name.length > MAX_ATTACHMENT_NAME) attachment.name = shortName(attachment.name);
    }
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
        const metadata = node.metadata as Record<string, unknown> | undefined;
        let attachmentName = metadata && typeof metadata === 'object' && typeof metadata.attachmentName === 'string' ? metadata.attachmentName : undefined;
        if (attachmentName !== undefined) {
            if (attachmentName.length > MAX_ATTACHMENT_NAME) metadata!.attachmentName = attachmentName = shortName(attachmentName);
            if (typeof node.text === 'string' && node.text.length > REPEAT_ALLOWANCE) {
                if (!shownAttachments.has(attachmentName)) shownAttachments.add(attachmentName);
                else if (takeRepeats(config, 1, node.text.length) === 0) node.text = repeatPreview(node.text);
            }
        }
        for (const list of [node.comments, node.notes, node.children]) {
            if (Array.isArray(list)) for (let i = list.length - 1; i >= 0; i--) stack.push(list[i]);
        }
    }
};

/**
 * Finds a parse's attachments by name through an index (a scan per node took nodes x attachments), and
 * joins a chart's text once for every node showing it (joined per node, one chart framed 2,000 times
 * built its text 2,000 times). Attachments added after a lookup are indexed at the next one.
 */
export const attachmentLookup = (attachments: OfficeAttachment[]) => {
    const index = new Map<string, OfficeAttachment>();
    let indexed = 0;
    const chartTexts = new Map<OfficeAttachment, string>();
    return {
        get(name: string | undefined): OfficeAttachment | undefined {
            for (; indexed < attachments.length; indexed++) {
                const attachment = attachments[indexed];
                if (attachment?.name && !index.has(attachment.name)) index.set(attachment.name, attachment);
            }
            return name === undefined ? undefined : index.get(name);
        },
        chartText(attachment: OfficeAttachment, delimiter: string): string | undefined {
            if (!attachment.chartData) return undefined;
            let text = chartTexts.get(attachment);
            if (text === undefined) chartTexts.set(attachment, text = attachment.chartData.rawTexts.join(delimiter));
            return text;
        },
    };
};

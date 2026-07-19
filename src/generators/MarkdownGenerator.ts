import { AdmonitionMetadata, BreakMetadata, CodeMetadata, ConversionResult, EmbedMetadata, FallbackToHtmlConfig, GeneratorConfig, HeadingMetadata, ImageMetadata, ListMetadata, MarkdownDialectConfig, MarkdownDialectPreset, NoteMetadata, OfficeContentNode, OfficeParserAST, TableMetadata, TextMetadata } from '../types.js';
import { escapeHtml, markdownEscapeText, sanitizeCssValue, sanitizeMarkdownUrl } from '../utils/sanitize.js';
import { BaseGenerator } from './BaseGenerator.js';

type ResolvedMarkdownDialect = Required<Omit<MarkdownDialectConfig, 'extends'>>;
type ResolvedFallbackToHtml = Required<FallbackToHtmlConfig>;

/**
 * Named Markdown dialect presets. `extended` reproduces this library's historical output
 * exactly (every feature on, GitHub-style admonitions) - the backward-compatibility anchor.
 */
const MARKDOWN_DIALECT_PRESETS: Record<MarkdownDialectPreset, ResolvedMarkdownDialect> = {
    extended: { admonitions: 'github', definitionLists: true, footnotes: true, citations: true, wikilinks: true, math: 'dollar', attributeLists: true, strikethrough: true, bulletListMarker: '-', orderedListMarker: '.', emphasisMarker: 'asterisk', tables: 'native' },
    github: { admonitions: 'github', definitionLists: false, footnotes: true, citations: false, wikilinks: false, math: 'dollar', attributeLists: false, strikethrough: true, bulletListMarker: '-', orderedListMarker: '.', emphasisMarker: 'asterisk', tables: 'native' },
    gitlab: { admonitions: 'gitlab', definitionLists: false, footnotes: true, citations: false, wikilinks: false, math: 'dollar', attributeLists: false, strikethrough: true, bulletListMarker: '-', orderedListMarker: '.', emphasisMarker: 'asterisk', tables: 'native' },
    obsidian: { admonitions: 'github', definitionLists: false, footnotes: true, citations: false, wikilinks: true, math: 'dollar', attributeLists: false, strikethrough: true, bulletListMarker: '-', orderedListMarker: '.', emphasisMarker: 'asterisk', tables: 'native' },
    pandoc: { admonitions: 'pandoc', definitionLists: true, footnotes: true, citations: true, wikilinks: false, math: 'dollar', attributeLists: true, strikethrough: true, bulletListMarker: '-', orderedListMarker: '.', emphasisMarker: 'asterisk', tables: 'native' },
    commonmark: { admonitions: 'none', definitionLists: false, footnotes: false, citations: false, wikilinks: false, math: 'none', attributeLists: false, strikethrough: false, bulletListMarker: '-', orderedListMarker: '.', emphasisMarker: 'asterisk', tables: 'html' },
};

/**
 * Normalizes `MdGeneratorConfig.dialect` into a fully-resolved preset. A string names a preset
 * directly; an object's `extends` field (default `'extended'`) names the base preset that any
 * omitted field falls back to - NOT "whatever preset was ambient before", since config merging
 * replaces the whole `dialect` field rather than layering an object on top of a prior string.
 */
function resolveDialect(dialect: MarkdownDialectPreset | MarkdownDialectConfig | undefined): ResolvedMarkdownDialect {
    if (dialect === undefined) return MARKDOWN_DIALECT_PRESETS.extended;
    if (typeof dialect === 'string') return MARKDOWN_DIALECT_PRESETS[dialect] ?? MARKDOWN_DIALECT_PRESETS.extended;

    const base = MARKDOWN_DIALECT_PRESETS[dialect.extends ?? 'extended'] ?? MARKDOWN_DIALECT_PRESETS.extended;
    return {
        admonitions: dialect.admonitions ?? base.admonitions,
        definitionLists: dialect.definitionLists ?? base.definitionLists,
        footnotes: dialect.footnotes ?? base.footnotes,
        citations: dialect.citations ?? base.citations,
        wikilinks: dialect.wikilinks ?? base.wikilinks,
        math: dialect.math ?? base.math,
        attributeLists: dialect.attributeLists ?? base.attributeLists,
        strikethrough: dialect.strikethrough ?? base.strikethrough,
        bulletListMarker: dialect.bulletListMarker ?? base.bulletListMarker,
        orderedListMarker: dialect.orderedListMarker ?? base.orderedListMarker,
        emphasisMarker: dialect.emphasisMarker ?? base.emphasisMarker,
        tables: dialect.tables ?? base.tables,
    };
}

/**
 * Normalizes `MdGeneratorConfig.fallbackToHtml` into a fully resolved object, mirroring
 * `HtmlGenerator`'s `resolveStandalone()` pattern: `true`/undefined turns every part on; `false`
 * turns every part off; an object's omitted fields default to on.
 */
function resolveFallbackToHtml(fallbackToHtml: boolean | FallbackToHtmlConfig | undefined): ResolvedFallbackToHtml {
    const uniform = (on: boolean): ResolvedFallbackToHtml => ({
        textFormatting: on, alignment: on, anchors: on, tables: on, embeds: on, cellLineBreaks: on,
    });
    if (fallbackToHtml === undefined || typeof fallbackToHtml === 'boolean') return uniform(fallbackToHtml ?? true);
    const on = uniform(true);
    return {
        textFormatting: fallbackToHtml.textFormatting ?? on.textFormatting,
        alignment: fallbackToHtml.alignment ?? on.alignment,
        anchors: fallbackToHtml.anchors ?? on.anchors,
        tables: fallbackToHtml.tables ?? on.tables,
        embeds: fallbackToHtml.embeds ?? on.embeds,
        cellLineBreaks: fallbackToHtml.cellLineBreaks ?? on.cellLineBreaks,
    };
}

/**
 * Generates Markdown from an AST.
 * 
 * DESIGN PRINCIPLES:
 * 1. **Strict Native Preference**: Always utilize native Markdown syntax for features that 
 *    are natively supported (headings, lists, bold/italic, etc.). HTML tags should NEVER 
 *    be used for these features.
 * 
 * 2. **Fidelity vs. Purity (The `fallbackToHtml` Principle)**:
 *    - When a given `fallbackToHtml` part is TRUE: The generator prioritizes high-fidelity
 *      document conversion for that part. It will use HTML tags for features that Markdown
 *      cannot natively represent (e.g., `<u>` for underline, `<div>` for alignment, `<table>`
 *      for nested structures or merged cells).
 *    - When FALSE: The generator prioritizes "pure" Markdown for that part.
 *      Unsupported features are either:
 *      - **Skipped**: Non-essential formatting like underline, subscript, superscript,
 *        or text alignment is omitted.
 *      - **Simplified/Hoisted**: Complex structures like nested tables are hoisted out
 *        of their parent cells and rendered as separate sequential tables to maintain
 *        valid Markdown syntax.
 *
 * 3. **Consistency**: All similar structural or formatting ideological problems must be
 *    resolved using these same rules to ensure predictable output.
 *
 * 4. **Dialect (`MdGeneratorConfig.dialect`)**: A second, independent axis from `fallbackToHtml` -
 *    which *native* Markdown syntax to emit for constructs with more than one real-world
 *    convention (admonitions, definition lists, footnotes, citations, wikilinks, math, list/
 *    emphasis markers, tables). See `resolveDialect()` and `MARKDOWN_DIALECT_PRESETS` above.
 */
export class MarkdownGenerator extends BaseGenerator<'md'> {
    private isInsideTable = false;
    private hoistedContent: string[] = [];
    private collectedAbbreviations = new Map<string, string>();
    private resolvedDialect: ResolvedMarkdownDialect;
    private resolvedFallbackToHtml: ResolvedFallbackToHtml;

    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'md'>) {
        super('md', ast, config);
        this.resolvedDialect = resolveDialect(this.config.mdConfig.dialect);
        this.resolvedFallbackToHtml = resolveFallbackToHtml(this.config.mdConfig.fallbackToHtml);
    }

    /**
     * Renders anchor tags if HTML fallback is allowed.
     */
    private renderAnchors(metadata: any): string {
        if (!this.resolvedFallbackToHtml.anchors || this.config.ignoreInternalLinks) return '';
        const ids = metadata?.anchorIds || [];
        return ids.map((aid: string) => `<a id="${this.slugify(aid)}"></a>`).join('');
    }

    /**
     * Serializes a frontmatter array as a YAML flow sequence (e.g. `[a, b]`), matching
     * MarkdownParser's frontmatter array handling. Plain strings are left bare; anything
     * that would break flow-array syntax (or isn't a string) falls back to JSON encoding.
     */
    private serializeFrontmatterArray(arr: any[]): string {
        const items = arr.map(item =>
            (typeof item === 'string' && item.trim() === item && !/[,[\]]/.test(item))
                ? item
                : JSON.stringify(item)
        );
        return `[${items.join(', ')}]`;
    }

    /**
     * Renders a Pandoc-style attribute list (e.g. `{width=50% align=left}`) from
     * ImageMetadata/TableMetadata's width/align fields - the canonical form is always
     * `key=value`, matching MarkdownParser's own vocabulary (MARKDOWN_DIALECT.md §15).
     */
    private renderAttributeList(meta: { width?: string; align?: string } | undefined): string {
        if (!this.resolvedDialect.attributeLists) return '';
        if (!meta?.width && !meta?.align) return '';
        const parts: string[] = [];
        // Strip characters that would break out of the `{...}` attribute list.
        if (meta.width) parts.push(`width=${String(meta.width).replace(/[{}\s]+/g, '')}`);
        if (meta.align) parts.push(`align=${String(meta.align).replace(/[{}\s]+/g, '')}`);
        return `{${parts.join(' ')}}`;
    }

    /** Converts a document-supplied date to an ISO string, or '' if invalid
     *  (a malformed date would otherwise throw a RangeError and abort generation). */
    private toIsoDate(value: unknown): string {
        if (value === undefined || value === null || value === '') return '';
        const d = new Date(value as any);
        return isNaN(d.getTime()) ? '' : d.toISOString();
    }

    /**
     * Generates Markdown string from the provided AST.
     * 
     * @returns A Markdown string
     */
    async generate(): Promise<ConversionResult<'md'>> {
        let output = '';

        // Add Metadata (YAML Front Matter)
        const meta = this.effectiveMetadata;
        if (meta) {
            output += '---\n';
            // JSON-encode scalar values so a title/author/description containing a
            // quote or newline can't break out of the YAML string and inject
            // arbitrary front-matter keys. (JSON.stringify of a benign value yields
            // the same `"..."` form as before, so normal output is unchanged.)
            if (meta.title) output += `title: ${JSON.stringify(meta.title)}\n`;
            if (meta.author) output += `author: ${JSON.stringify(meta.author)}\n`;
            const createdIso = this.toIsoDate(meta.created);
            if (createdIso) output += `created: ${createdIso}\n`;
            const modifiedIso = this.toIsoDate(meta.modified);
            if (modifiedIso) output += `modified: ${modifiedIso}\n`;
            if (meta.description) output += `description: ${JSON.stringify(meta.description)}\n`;

            if (meta.customProperties) {
                for (const [key, val] of Object.entries(meta.customProperties)) {
                    // Strip newlines/colons from the key so it can't inject a new mapping.
                    const safeKey = String(key).replace(/[\r\n:]+/g, ' ').trim();
                    output += `${safeKey}: ${Array.isArray(val) ? this.serializeFrontmatterArray(val) : JSON.stringify(val)}\n`;
                }
            }
            output += '---\n\n';
        }

        const processor = async (node: OfficeContentNode, childrenOutput: string): Promise<string> => {
            // Handle Style Mapping for Markdown using the semantic mapping helper
            const mapping = this.getSemanticMapping(node);
            if (mapping) {
                // Map common HTML tags to Markdown equivalents
                if (mapping.tag === 'blockquote') return `> ${childrenOutput}\n\n`;
                if (mapping.tag === 'code') return `\`${childrenOutput}\` `;
                if (mapping.tag === 'pre') return `\`\`\`\n${childrenOutput}\n\`\`\`\n\n`;

                const hMatch = mapping.tag.match(/^h([1-6])$/);
                if (hMatch) {
                    const level = parseInt(hMatch[1]);
                    return `${'#'.repeat(level)} ${childrenOutput}\n\n`;
                }
            }

            switch (node.type) {
                case 'text': {
                    // Entity-encode angle brackets so document text can't inject a raw
                    // HTML tag (e.g. <script>) when the Markdown is rendered to HTML.
                    let text = markdownEscapeText(node.text || '');
                    if (this.config.includeFormatting && node.formatting) {
                        const emphasisAsterisk = this.resolvedDialect.emphasisMarker === 'asterisk';
                        if (node.formatting.bold) text = emphasisAsterisk ? `**${text}**` : `__${text}__`;
                        if (node.formatting.italic) text = emphasisAsterisk ? `*${text}*` : `_${text}_`;
                        if (node.formatting.strikethrough && this.resolvedDialect.strikethrough) text = `~~${text}~~`;

                        // Use HTML tags for formatting not natively supported by standard Markdown
                        if (this.resolvedFallbackToHtml.textFormatting) {
                            if (node.formatting.underline) text = `<u>${text}</u>`;
                            if (node.formatting.subscript) text = `<sub>${text}</sub>`;
                            if (node.formatting.superscript) text = `<sup>${text}</sup>`;
                        }
                    }
                    const meta = node.metadata as TextMetadata;
                    if (meta?.wikilink && this.resolvedDialect.wikilinks) {
                        // Obsidian syntax: bare page name, or page|alias when the display
                        // text differs from the page name. Strip the `[]|`/newline chars
                        // that would break out of the `[[...]]` wrapper.
                        const page = (meta.link || '').replace(/[[\]|\r\n]+/g, '');
                        const alias = (node.text || '').replace(/[[\]|\r\n]+/g, '');
                        text = (node.text && node.text !== (meta.link || '')) ? `[[${page}|${alias}]]` : `[[${page}]]`;
                    } else if (meta?.link) {
                        const isInternal = meta.linkType !== 'external';
                        if (!this.config.ignoreInternalLinks || !isInternal) {
                            let link = meta.link;
                            // Slugify internal link targets to match heading IDs if generating IDs
                            if (isInternal && link.startsWith('#') && (this.config.generateIds || this.resolvedFallbackToHtml.anchors)) {
                                const target = link.substring(1);
                                link = '#' + this.slugify(target);
                            }
                            // Reject javascript:/data: schemes and encode `()`/whitespace so the
                            // URL can't break out of `](...)` or inject a script link.
                            text = `[${text}](${sanitizeMarkdownUrl(link)})`;
                        }
                    }
                    if (meta?.abbreviationTitle) {
                        // Markdown Extra's abbreviation syntax has no inline marker - the bare
                        // word round-trips as-is, with its expansion collected at the document
                        // end via `*[abbr]: title`.
                        this.collectedAbbreviations.set(node.text || '', meta.abbreviationTitle);
                    }
                    if (meta?.citationKey) {
                        const key = String(meta.citationKey).replace(/[[\]\r\n]+/g, '');
                        text = this.resolvedDialect.citations ? `[@${key}]` : `[${key}]`;
                    }
                    return text;
                }

                case 'heading': {
                    const meta = node.metadata as HeadingMetadata;
                    const level = Math.min(Math.max(meta?.level || 1, 1), 6);
                    const prefix = '#'.repeat(level) + ' ';

                    let id = '';
                    let remainingAnchors: string[] = [];

                    if (!this.config.ignoreInternalLinks && meta?.anchorIds && meta.anchorIds.length > 0) {
                        const ids = [...meta.anchorIds];
                        const lastId = ids.pop()!;
                        // Slugify the explicit ID to ensure it's a valid Markdown identifier
                        id = ` {#${this.slugify(lastId)}}`;
                        remainingAnchors = ids;
                    } else if (this.config.generateIds) {
                        id = ` {#${this.slugify(this.getNodeText(node))}}`;
                    }

                    const anchors = this.resolvedFallbackToHtml.anchors
                        ? remainingAnchors.map(aid => `<a name="${this.slugify(aid)}"></a>`).join('')
                        : '';
                    let content = `${prefix}${childrenOutput}${id}`;

                    // Alignment fallback via HTML div/p
                    if (this.resolvedFallbackToHtml.alignment && meta?.alignment && meta.alignment !== 'left') {
                        // Use extra newlines to ensure Markdown inside the div is parsed
                        content = `<div style="text-align: ${sanitizeCssValue(meta.alignment)}">\n\n${content}\n\n</div>`;
                    }

                    return `${anchors}${anchors ? '\n' : ''}${content}\n\n`;
                }

                case 'paragraph': {
                    const meta = node.metadata as any;
                    const anchors = this.renderAnchors(meta);
                    let content = childrenOutput;

                    // Alignment fallback via HTML div/p
                    if (this.resolvedFallbackToHtml.alignment && meta?.alignment && meta.alignment !== 'left') {
                        content = `<div style="text-align: ${sanitizeCssValue(meta.alignment)}">${content}</div>`;
                    }

                    return childrenOutput ? `${anchors}${content}\n\n` : '';
                }

                case 'list': {
                    const meta = node.metadata as ListMetadata;
                    const indentSpaces = ' '.repeat(4);
                    const indent = indentSpaces.repeat(meta?.indentation || 0);
                    const bullet = `${this.resolvedDialect.bulletListMarker} `;
                    const marker = meta?.isTask
                        ? (meta.checked ? `${bullet}[x] ` : `${bullet}[ ] `)
                        : (meta?.listType === 'ordered' ? `${(meta.itemIndex ?? 0) + 1}${this.resolvedDialect.orderedListMarker} ` : bullet);
                    const anchors = this.renderAnchors(meta);
                    return `${indent}${marker}${anchors}${childrenOutput}\n`;
                }

                case 'image': {
                    if (!this.config.includeImages) return '';
                    const meta = node.metadata as ImageMetadata;
                    const alt = meta?.altText || 'image';
                    let src = meta?.url || meta?.attachmentName || '';

                    // Resolve attachment to data URI if no external URL is provided
                    if (!meta?.url && meta?.attachmentName && this.ast) {
                        const attachment = this.ast.attachments.find(a => a.name === meta.attachmentName);
                        if (attachment) {
                            src = `data:${attachment.mimeType || 'image/png'};base64,${attachment.data}`;
                        }
                    }

                    const anchors = this.renderAnchors(meta);
                    // Strip `[]` from alt (would close the `![...]`) and neutralize the URL scheme.
                    const safeAlt = markdownEscapeText(alt).replace(/[[\]]/g, '');
                    const safeSrc = sanitizeMarkdownUrl(src, { allowDataImage: true });
                    return `${anchors}${anchors ? '\n' : ''}![${safeAlt}](${safeSrc})${this.renderAttributeList(meta)}`;
                }

                case 'table': {
                    const anchors = this.renderAnchors(node.metadata);
                    const tableOutput = await this.renderMarkdownTable(node, processor);
                    // The HTML-fallback path (merged cells/nested tables, or a dialect that forces
                    // HTML tables outright) already carries data-align on the <table> tag directly -
                    // only the plain pipe-table form needs the attribute-list syntax for alignment.
                    const usedHtmlFallback = this.resolvedDialect.tables === 'html' ||
                        (this.resolvedFallbackToHtml.tables && (this.hasNestedTable(node) || this.hasColspanOrRowspan(node)));
                    const attrList = usedHtmlFallback ? '' : this.renderAttributeList(node.metadata as TableMetadata);
                    if (attrList) {
                        // Must glue directly below the last row with no blank line, or
                        // MarkdownParser's block splitter won't see it as part of the same block.
                        return `${anchors}${anchors ? '\n' : ''}${tableOutput.replace(/\n+$/, '\n')}${attrList}\n`;
                    }
                    return `${anchors}${anchors ? '\n' : ''}${tableOutput}`;
                }

                case 'row':
                case 'cell': {
                    // These are handled manually in the 'table' case above
                    return childrenOutput;
                }

                case 'break': {
                    // A hard line break (CommonMark: two trailing spaces before the
                    // newline) round-trips back to a distinct 'break' node on reparse;
                    // every other breakType (including 'page', used for a thematic-break
                    // HR) keeps emitting a bare newline, unchanged.
                    const meta = node.metadata as BreakMetadata | undefined;
                    if (meta?.breakType === 'carriageReturn') return '  \n';
                    return '\n';
                }

                case 'code': {
                    const meta = node.metadata as CodeMetadata;
                    if (meta?.math === 'block') {
                        return this.resolvedDialect.math === 'dollar' ? `\n$$\n${node.text || ''}\n$$\n\n` : `\n${node.text || ''}\n\n`;
                    }
                    if (meta?.math === 'inline') {
                        return this.resolvedDialect.math === 'dollar' ? `$${node.text || ''}$` : (node.text || '');
                    }
                    const lang = (meta?.language || '').replace(/[\r\n`]+/g, '');
                    // Block code if it contains newlines, else inline
                    if (node.text && node.text.includes('\n')) {
                        // Fence with one more backtick than the longest run inside the content
                        // so an embedded ``` can't close the block early and inject markup.
                        const longestRun = Math.max(0, ...(node.text.match(/`+/g) || []).map(s => s.length));
                        const fence = '`'.repeat(Math.max(3, longestRun + 1));
                        return `\n${fence}${lang}\n${node.text}\n${fence}\n\n`;
                    } else {
                        const t = node.text || '';
                        const longestRun = Math.max(0, ...(t.match(/`+/g) || []).map(s => s.length));
                        const fence = '`'.repeat(Math.max(1, longestRun + 1));
                        const pad = (t.startsWith('`') || t.endsWith('`')) ? ' ' : '';
                        return `${fence}${pad}${t}${pad}${fence} `;
                    }
                }

                case 'sheet': {
                    const anchors = this.renderAnchors(node.metadata);
                    const tableOutput = await this.renderMarkdownTable(node, processor);
                    return `\n---\n\n${anchors}${anchors ? '\n' : ''}${tableOutput}\n\n`;
                }

                case 'slide': {
                    const anchors = this.renderAnchors(node.metadata);
                    return `\n---\n\n${anchors}${anchors ? '\n' : ''}${childrenOutput}\n\n`;
                }
                case 'page': {
                    const anchors = this.renderAnchors(node.metadata);
                    return `\n---\n\n${anchors}${anchors ? '\n' : ''}${childrenOutput}\n\n`;
                }
                case 'note': {
                    const meta = node.metadata as NoteMetadata;
                    if (meta?.noteType === 'footnote' || meta?.noteType === 'endnote') {
                        if (!this.resolvedDialect.footnotes) {
                            // Dialect has no footnote syntax - the caller inlines this bare body
                            // as a parenthetical at the reference point instead of collecting it
                            // into an end-of-document "### Notes" section under a [^id] marker.
                            return childrenOutput.trim();
                        }
                        return `[^${this.getFootnoteKey(node)}]: ${childrenOutput.trim()}\n\n`;
                    }
                    return `> **Note:** ${childrenOutput.trim()}\n\n`;
                }

                case 'embed': {
                    // Markdown has no native embed syntax. When fallbackToHtml.embeds is on (our
                    // save default), emit the exact single-line div MarkdownParser recognises on
                    // reimport; otherwise degrade to a plain link.
                    const meta = node.metadata as EmbedMetadata;
                    const id = meta?.videoId || '';
                    if (this.resolvedFallbackToHtml.embeds) {
                        const width = meta?.width ? ` data-width="${escapeHtml(meta.width)}"` : '';
                        const align = meta?.align ? ` data-align="${escapeHtml(meta.align)}"` : '';
                        return `\n<div data-youtube-video="${escapeHtml(id)}"${width}${align}></div>\n\n`;
                    }
                    const url = meta?.url || (id ? `https://youtu.be/${id}` : '');
                    return url ? `[YouTube](${sanitizeMarkdownUrl(url)})\n\n` : '';
                }

                case 'admonition': {
                    const meta = node.metadata as AdmonitionMetadata;
                    const type = meta?.admonitionType || 'note';
                    const label = type.toUpperCase();
                    const body = childrenOutput.trim();

                    switch (this.resolvedDialect.admonitions) {
                        case 'gitlab':
                            // GLFM fenced-div: no dedicated title syntax, so a custom title (if
                            // any) is folded into the body as a bold first line.
                            return `:::${type}\n${meta?.title ? `**${meta.title}**\n\n` : ''}${body}\n:::\n\n`;
                        case 'pandoc':
                            // Pandoc's own fenced-div-with-class syntax; same title handling as gitlab.
                            return `::: {.${type}}\n${meta?.title ? `**${meta.title}**\n\n` : ''}${body}\n:::\n\n`;
                        case 'none': {
                            // Degrade to a plain bold-labeled blockquote, no special marker.
                            const quotedLines = body.split('\n').map(l => l.length > 0 ? `> ${l}` : '>').join('\n');
                            const heading = meta?.title || label.charAt(0) + label.slice(1).toLowerCase();
                            return `> **${heading}:**\n${quotedLines}\n\n`;
                        }
                        case 'github':
                        default: {
                            // Canonical GitHub blockquote form. No dedicated title syntax either
                            // (matches this library's historical output).
                            const quotedLines = body.split('\n').map(l => l.length > 0 ? `> ${l}` : '>').join('\n');
                            return `> [!${label}]\n${quotedLines}\n\n`;
                        }
                    }
                }

                case 'definitionList':
                    if (!this.resolvedDialect.definitionLists) return `${childrenOutput}\n`;
                    return `${childrenOutput}\n`;

                case 'definitionTerm':
                    if (!this.resolvedDialect.definitionLists) return `**${childrenOutput}**\n\n`;
                    return `${childrenOutput}\n`;

                case 'definitionDescription':
                    if (!this.resolvedDialect.definitionLists) return `${childrenOutput}\n\n`;
                    return `: ${childrenOutput}\n`;

                default:
                    return childrenOutput;
            }
        };

        const optimizedContent = this.optimizeNodes(this.ast.content);
        for (let i = 0; i < optimizedContent.length; i++) {
            const node = optimizedContent[i];
            const nextNode = optimizedContent[i + 1];
            let result = await this.processNodeRecursive(node, processor);

            // Ensure lists and other block elements are separated from non-similar content by a blank line
            if (nextNode) {
                const isBothLists = node.type === 'list' && nextNode.type === 'list';
                if (!isBothLists) {
                    if (!result.endsWith('\n\n')) {
                        if (result.endsWith('\n')) result += '\n';
                        else result += '\n\n';
                    }
                }
            }

            output += result;
        }

        if (this.collectedNotes.length > 0) {
            let notesMd = '\n\n---\n\n### Notes\n\n';
            for (const note of this.collectedNotes) {
                notesMd += await this.processNodeRecursive(note, processor);
            }
            output += notesMd;
        }

        if (this.collectedAbbreviations.size > 0) {
            output += '\n\n';
            for (const [abbr, title] of this.collectedAbbreviations) {
                output += `*[${abbr}]: ${title}\n`;
            }
        }

        // Only a run of literal "\n" at either end is ever a generator artifact here: block
        // separators, the notes/abbreviations sections, the unconditional '\n\n' before
        // hoistedContent (added even when hoistedContent is empty), and renderMarkdownTable's
        // HTML-fallback branches, which unconditionally wrap in a leading+trailing '\n' as
        // separators from whatever precedes/follows (in practice this rarely surfaces at the very
        // start of `output` today since frontmatter's own "---" almost always precedes real
        // content first - see the type doc on `ast.metadata` - but the strip is correct regardless
        // of what precedes it). Nothing else at either end is a generator artifact: not leading
        // whitespace, and not any other kind of trailing whitespace, both of which would be real
        // document content. See the identical reasoning in TextGenerator.generate().
        return {
            value: (output + '\n\n' + this.hoistedContent.join('\n\n')).replace(/^\n+|\n+$/g, ''),
            messages: this.messages
        };
    }

    /**
     * Recursively processes nodes and builds output.
     * Overridden to provide AST optimization (merging adjacent text nodes).
     */
    protected override async processNodeRecursive(
        node: OfficeContentNode,
        processor: (node: OfficeContentNode, childrenOutput: string) => string | Promise<string>
    ): Promise<string> {
        // Allow user to completely override rendering or skip via onNode
        const override = await this.handleOnNode(node);
        if (override === false) {
            return '';
        }
        if (typeof override === 'string') {
            return override;
        }

        let childrenOutput = '';
        if (node.children && node.children.length > 0) {
            // Optimization: Merge adjacent text nodes with identical formatting
            const optimizedChildren = this.optimizeNodes(node.children);
            for (const child of optimizedChildren) {
                childrenOutput += await this.processNodeRecursive(child, processor);
            }
        }

        // When the dialect has no footnote syntax, a footnote/endnote is inlined right at its
        // reference point instead (see below) - so it must not also be collected into the
        // end-of-document "### Notes" section, or its content would be duplicated.
        const isInlinedFootnote = (note: OfficeContentNode): boolean => {
            const meta = note.metadata as NoteMetadata;
            return (meta?.noteType === 'footnote' || meta?.noteType === 'endnote') && !this.resolvedDialect.footnotes;
        };

        if (node.notes && node.notes.length > 0) {
            if (node.type !== 'slide') {
                this.collectedNotes.push(...node.notes.filter(note => !isInlinedFootnote(note)));
            }
        }

        let result = await processor(node, childrenOutput);

        if (node.type === 'slide' && node.notes && node.notes.length > 0) {
            for (const note of node.notes) {
                result += await this.processNodeRecursive(note, processor);
            }
        } else if (node.notes && node.notes.length > 0) {
            for (const note of node.notes) {
                const meta = note.metadata as NoteMetadata;
                if (meta?.noteType !== 'footnote' && meta?.noteType !== 'endnote') continue;
                if (isInlinedFootnote(note)) {
                    // Markdown-specific degrade (not RTF/plain-text's "drop the marker, just
                    // append at the end" convention): inline the note's rendered body as a
                    // parenthetical right where it's referenced, since Markdown readers benefit
                    // from an inline association those simpler formats don't need in the same way.
                    const body = await this.processNodeRecursive(note, processor);
                    result += ` (Note: ${body})`;
                } else {
                    // Emit the [^id] reference marker at the point of reference. Without this,
                    // a footnote/endnote would only ever show up in the collected ### Notes
                    // section at the end, with no indication of where it was originally cited.
                    result += `[^${this.getFootnoteKey(note)}]`;
                }
            }
        }

        return result;
    }

    /**
     * Merges adjacent text nodes with identical formatting and metadata.
     */
    private optimizeNodes(nodes: OfficeContentNode[]): OfficeContentNode[] {
        if (nodes.length <= 1) return nodes;

        const result: OfficeContentNode[] = [];
        let current: OfficeContentNode | null = null;

        for (const node of nodes) {
            if (node.type === 'text' && current && current.type === 'text' &&
                this.areFormattingEqual(node.formatting, current.formatting) &&
                JSON.stringify(node.metadata) === JSON.stringify(current.metadata)) {
                current.text = (current.text || '') + (node.text || '');
                if (current.rawContent && node.rawContent) current.rawContent += node.rawContent;
                if (node.notes && node.notes.length > 0) {
                    if (!current.notes) current.notes = [];
                    current.notes.push(...node.notes);
                }
            } else {
                current = { ...node }; // Clone
                if (node.notes) {
                    current.notes = [...node.notes];
                }
                result.push(current);
            }
        }
        return result;
    }

    private areFormattingEqual(f1: any, f2: any): boolean {
        if (f1 === f2) return true;
        if (!f1 || !f2) return false;
        const keys1 = Object.keys(f1);
        const keys2 = Object.keys(f2);
        if (keys1.length !== keys2.length) return false;
        return keys1.every(key => f1[key] === f2[key]);
    }

    private async renderMarkdownTable(node: OfficeContentNode, processor: any): Promise<string> {
        if (!node.children || node.children.length === 0) return '';

        // A dialect that has no native table syntax at all (e.g. strict CommonMark) always
        // renders as HTML, regardless of complexity - this is a separate axis from the
        // nested/merged-cell HTML fallback below, which only applies to otherwise-native tables.
        if (this.resolvedDialect.tables === 'html') {
            return '\n' + await this.renderTableAsHtml(node) + '\n';
        }

        // If table is complex, nested, or uses merges, fallback to HTML for high fidelity if allowed
        const isComplex = this.hasNestedTable(node) || this.hasColspanOrRowspan(node);
        if (this.resolvedFallbackToHtml.tables && isComplex) {
            return '\n' + await this.renderTableAsHtml(node) + '\n';
        }

        // Handle nested tables in pure Markdown by hoisting them out
        if (this.isInsideTable && !this.resolvedFallbackToHtml.tables) {
            const wasInside = this.isInsideTable;
            this.isInsideTable = false; // Reset to allow rendering the hoisted table correctly
            const hoistedId = this.hoistedContent.length + 1;
            const tableOutput = await this.renderMarkdownTableInternal(node, processor);
            this.hoistedContent.push(`**Table ${hoistedId} (Hoisted from cell content):**\n${tableOutput}`);
            this.isInsideTable = wasInside;
            return `*(See Table ${hoistedId} below)*`;
        }

        this.isInsideTable = true;
        const result = await this.renderMarkdownTableInternal(node, processor);
        this.isInsideTable = false;
        return result;
    }

    private async renderMarkdownTableInternal(node: OfficeContentNode, processor: any): Promise<string> {
        let tableOutput = '';
        let maxCols = 0;

        // First pass: Process rows and determine max columns (accounting for colspans)
        const processedRows: string[][] = [];
        for (const rowNode of (node.children ?? [])) {
            const override = await this.handleOnNode(rowNode);
            if (override === false) continue;
            if (typeof override === 'string') {
                processedRows.push([override]);
                continue;
            }

            const rowCells: string[] = [];
            let lastCol = -1;

            if (rowNode.children) {
                const cellNodes = rowNode.children.filter(c => c.type === 'cell');
                for (const cellNode of cellNodes) {
                    const currentCol = (cellNode.metadata as any)?.col ?? (lastCol + 1);

                    // Fill gaps with empty cells
                    while (lastCol < currentCol - 1) {
                        rowCells.push(' ');
                        lastCol++;
                    }

                    // Process cell content
                    let cellContent = await this.processNodeRecursive(cellNode, processor);
                    // Use <br> fallback only if allowed, otherwise space
                    const br = this.resolvedFallbackToHtml.cellLineBreaks ? '<br>' : ' ';
                    cellContent = cellContent.trim().replace(/\n+/g, br).replace(/\|/g, '\\|');
                    rowCells.push(cellContent);

                    // Handle colspan by adding empty cells
                    const colSpan = (cellNode.metadata as any)?.colSpan || 1;
                    for (let i = 1; i < colSpan; i++) {
                        rowCells.push(' ');
                    }
                    lastCol = currentCol + colSpan - 1;
                }
            }
            processedRows.push(rowCells);
            maxCols = Math.max(maxCols, rowCells.length);
        }

        // Second pass: Build table string with separator
        for (let i = 0; i < processedRows.length; i++) {
            const row = processedRows[i];
            // Pad row with empty cells if it has fewer than maxCols
            while (row.length < maxCols) row.push(' ');

            tableOutput += `| ${row.join(' | ')} |\n`;

            if (i === 0) {
                // Header separator
                tableOutput += `| ${Array(maxCols).fill(' --- ').join(' | ')} |\n`;
            }
        }

        return `\n${tableOutput}\n`;
    }

    private hasNestedTable(node: OfficeContentNode): boolean {
        if (!node.children) return false;
        for (const child of node.children) {
            if (child.type === 'table') return true;
            if (this.hasNestedTable(child)) return true;
        }
        return false;
    }

    private hasColspanOrRowspan(node: OfficeContentNode): boolean {
        if (!node.children) return false;
        for (const row of node.children) {
            if (row.type === 'row' && row.children) {
                for (const cell of row.children) {
                    if (cell.type === 'cell') {
                        const meta = cell.metadata as any;
                        if ((meta?.colSpan && meta.colSpan > 1) || (meta?.rowSpan && meta.rowSpan > 1)) {
                            return true;
                        }
                    }
                }
            }
        }
        return false;
    }

    /**
     * Renders a complex table as HTML since Markdown doesn't support nested tables or rowspans.
     */
    private async renderTableAsHtml(node: OfficeContentNode, override?: string | false | void): Promise<string> {
        if (override === false) return '';
        if (typeof override === 'string') {
            if (node.type === 'row') return `  <tr><td colspan="100%">${override}</td></tr>\n`;
            if (node.type === 'cell') return `<td>${override}</td>`;
            return override;
        }

        if (node.type === 'table') {
            let rows = '';
            if (node.children) {
                for (const row of node.children) {
                    rows += await this.renderTableAsHtml(row, await this.handleOnNode(row));
                }
            }
            // Carry table-layout alignment through the HTML fallback so it isn't lost
            // just because the table also needed HTML for merged cells.
            const tableMeta = node.metadata as any;
            const alignAttr = tableMeta?.align ? ` data-align="${escapeHtml(tableMeta.align)}"` : '';
            return `<table${alignAttr}>\n${rows}</table>\n`;
        } else if (node.type === 'row') {
            let cells = '';
            if (node.children) {
                for (const cell of node.children) {
                    cells += await this.renderTableAsHtml(cell, await this.handleOnNode(cell));
                }
            }
            return `  <tr>\n${cells}  </tr>\n`;
        } else if (node.type === 'cell') {
            const meta = node.metadata as any;
            const rs = meta?.rowSpan > 1 ? ` rowspan="${meta.rowSpan}"` : '';
            const cs = meta?.colSpan > 1 ? ` colspan="${meta.colSpan}"` : '';

            let content = '';
            if (node.children) {
                // Use a simplified HTML processor for cell content
                for (const child of this.optimizeNodes(node.children)) {
                    content += await this.processNodeRecursive(child, async (n, co) => {
                        switch (n.type) {
                            case 'text': {
                                // Inside HTML table cells, entity-encode angle brackets so cell
                                // text can't inject a raw tag (e.g. </td><script>).
                                let text = markdownEscapeText(n.text || '');
                                if (n.formatting?.bold) text = `<b>${text}</b>`;
                                if (n.formatting?.italic) text = `<i>${text}</i>`;
                                if (n.formatting?.underline) text = `<u>${text}</u>`;
                                if (n.formatting?.subscript) text = `<sub>${text}</sub>`;
                                if (n.formatting?.superscript) text = `<sup>${text}</sup>`;
                                return text;
                            }
                            case 'paragraph': return `<p>${co}</p>`;
                            case 'heading': {
                                const level = Math.min(Math.max(Number((n.metadata as any)?.level) || 1, 1), 6);
                                return `<h${level}>${co}</h${level}>`;
                            }
                            case 'table': return await this.renderTableAsHtml(n);
                            default: return co;
                        }
                    });
                }
            }
            return `    <td${rs}${cs}>${content}</td>\n`;
        }
        return '';
    }
}

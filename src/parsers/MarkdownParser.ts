import { AdmonitionMetadata, BreakMetadata, CodeMetadata, EmbedMetadata, FullOfficeParserConfig, HeadingMetadata, ImageMetadata, ListMetadata, OfficeAttachment, OfficeContentNode, OfficeMetadata, OfficeParserAST, TextFormatting, TextMetadata } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { checkAbortSignal } from '../utils/errorUtils.js';

// Sentinel node type for a standalone bookmark-anchor block (e.g. `<a id="x"></a>` on its
// own line). A post-parse pass folds these into the following node's anchorIds so they
// round-trip as real anchors rather than being escaped to visible text on regeneration.
const ANCHOR_PLACEHOLDER = '__anchorPlaceholder__';

/**
 * Splits the inner content of a YAML flow array (`a, "b, c", d`) on top-level commas,
 * ignoring commas inside single- or double-quoted items.
 */
const splitFlowArrayItems = (inner: string): string[] => {
    const items: string[] = [];
    let current = '';
    let quote: '"' | '\'' | null = null;
    for (const ch of inner) {
        if (quote) {
            current += ch;
            if (ch === quote) quote = null;
        } else if (ch === '"' || ch === '\'') {
            quote = ch;
            current += ch;
        } else if (ch === ',') {
            items.push(current.trim());
            current = '';
        } else {
            current += ch;
        }
    }
    if (current.trim() !== '') items.push(current.trim());
    return items;
};

/**
 * Maps every accepted-on-import admonition type spelling (GitHub's five plus GLFM's
 * `danger`) to the canonical AdmonitionMetadata type. Per MARKDOWN_DIALECT.md's
 * Decisions, `danger` folds into `caution` - there is no separate danger type.
 */
const ADMONITION_TYPE_MAP: Record<string, AdmonitionMetadata['admonitionType']> = {
    note: 'note',
    tip: 'tip',
    important: 'important',
    warning: 'warning',
    caution: 'caution',
    danger: 'caution'
};

export const parseMarkdown = async (buffer: Buffer, config: FullOfficeParserConfig): Promise<OfficeParserAST> => {
    // Honour cancellation requests before the line-by-line Markdown scanning loop begins.
    // Markdown parsing is entirely synchronous and CPU-bound, so failing fast avoids
    // processing content whose result will be discarded anyway.
    checkAbortSignal(config.abortSignal);

    let textStr = buffer.toString('utf-8');
    textStr = textStr.replace(/\r\n/g, '\n');

    const content: OfficeContentNode[] = [];
    const metadata: OfficeMetadata = {};
    const attachments: OfficeAttachment[] = [];

    // Parse YAML Front Matter
    if (textStr.startsWith('---\n')) {
        const endIdx = textStr.indexOf('\n---\n', 4);
        if (endIdx !== -1) {
            const frontMatter = textStr.substring(4, endIdx);
            textStr = textStr.substring(endIdx + 5);

            const lines = frontMatter.split('\n');
            const customProps: Record<string, any> = {};
            const nativeProps: Record<string, any> = {};

            for (const line of lines) {
                const match = line.match(/^([^:]+):\s*(.*)$/);
                if (match) {
                    const key = match[1].trim();
                    const rawVal = match[2].trim();
                    const val = rawVal.replace(/^"(.*)"$/, '$1');

                    let parsedVal: any = val;
                    if (rawVal.startsWith('[') && rawVal.endsWith(']')) {
                        // Flow-array (`tags: [a, b]`) or JSON-array (`tags: ["a","b"]`) value -
                        // parse into a real array instead of storing the literal bracket string,
                        // so it round-trips symmetrically with MarkdownGenerator's frontmatter output.
                        try {
                            const jsonParsed = JSON.parse(rawVal);
                            parsedVal = Array.isArray(jsonParsed) ? jsonParsed : val;
                        } catch {
                            const inner = rawVal.slice(1, -1).trim();
                            parsedVal = inner === '' ? [] : splitFlowArrayItems(inner).map(item => item.replace(/^['"](.*)['"]$/, '$1'));
                        }
                    } else if (val === 'true') parsedVal = true;
                    else if (val === 'false') parsedVal = false;
                    else if (!isNaN(Number(val)) && val !== '') parsedVal = Number(val);

                    nativeProps[key] = parsedVal;

                    if (key === 'title') metadata.title = val;
                    else if (key === 'author') metadata.author = val;
                    else if (key === 'created') metadata.created = new Date(val);
                    else if (key === 'modified') metadata.modified = new Date(val);
                    else if (key === 'description') metadata.description = val;
                    else {
                        customProps[key] = parsedVal;
                    }
                }
            }
            if (Object.keys(customProps).length > 0) metadata.customProperties = customProps;
            if (Object.keys(nativeProps).length > 0) metadata.nativeProperties = nativeProps;
        }
    }

    // Strip MDX/JSX component tags (parse-only - we never author MDX). Components are
    // distinguished from plain HTML by an uppercase-leading tag name, matching React/MDX
    // convention. Self-closing components are removed entirely; paired components keep
    // their inner Markdown content. Iterate to a fixed point so nested components (of
    // different names) are all unwrapped, not just the outermost one.
    // Cap the passes: each iteration unwraps one nesting level, so a pathologically
    // deep `<A><A>...</A></A>` input would otherwise be O(depth * n). Real documents
    // nest only a handful of levels; anything past the cap is left as-is.
    let previousTextStr;
    let mdxPasses = 0;
    const MAX_MDX_PASSES = 100;
    do {
        previousTextStr = textStr;
        textStr = textStr.replace(/<[A-Z][A-Za-z0-9]*(?:\s+[^>]*?)?\/>/g, '');
        textStr = textStr.replace(/<([A-Z][A-Za-z0-9]*)(?:\s+[^>]*?)?>([\s\S]*?)<\/\1>/g, (_m, _name, inner) => inner);
    } while (textStr !== previousTextStr && ++mdxPasses < MAX_MDX_PASSES);

    // Extract code blocks first to protect their contents. Accepts both backtick and
    // tilde fences (CommonMark's two fence characters); the backreference on the fence
    // run means a `~~~`-fenced block isn't closed early by a stray ``` inside it, and
    // vice versa.
    const codeBlocks: string[] = [];
    textStr = textStr.replace(/^(`{3,}|~{3,})(\w*)\n([\s\S]*?)\n\1$/gm, (match, _fence, lang, code) => {
        const id = `__CODE_BLOCK_${codeBlocks.length}__`;
        codeBlocks.push(JSON.stringify({ lang, code }));
        return `\n\n${id}\n\n`;
    });

    // Extract block math ($$\n...\n$$) before block splitting, mirroring the code-block
    // pre-pass above - its body may contain blank lines that would otherwise fragment it.
    // Inline math ($...$) is handled directly in parseInline below.
    const mathBlocks: string[] = [];
    textStr = textStr.replace(/^\$\$\n([\s\S]*?)\n\$\$$/gm, (_match, latex) => {
        const id = `__MATH_BLOCK_${mathBlocks.length}__`;
        mathBlocks.push(latex);
        return `\n\n${id}\n\n`;
    });

    // Extract GLFM-style fenced-div admonitions (`:::note ... :::`) before block splitting,
    // since their body may itself contain blank lines that would otherwise fragment them.
    // The `> [!NOTE]` GitHub form doesn't need this - it's detected inline in the blockquote
    // branch below, since a `>`-prefixed block never contains a real blank line.
    const admonitionBlocks: string[] = [];
    textStr = textStr.replace(/^:::(\w+)[ \t]*\n([\s\S]*?)\n:::[ \t]*$/gm, (match, type, body) => {
        const admonitionType = ADMONITION_TYPE_MAP[type.toLowerCase()];
        if (!admonitionType) return match; // Unrecognised type - leave as literal text.
        const id = `__ADMONITION_${admonitionBlocks.length}__`;
        admonitionBlocks.push(JSON.stringify({ admonitionType, body }));
        return `\n\n${id}\n\n`;
    });

    // Extract footnote definitions (`[^id]: text`) before block splitting, since
    // definitions conventionally live at the end of the document, after every place
    // they're referenced - inline parsing below needs the full map upfront. v1 only
    // supports single-line definitions (MultiMarkdown/Pandoc/GLFM's common baseline).
    const footnoteDefinitions = new Map<string, string>();
    textStr = textStr.replace(/^\[\^([^\]]+)\]:[ \t]*(.*)$/gm, (_match, id, definition) => {
        footnoteDefinitions.set(id, definition.trim());
        return '';
    });

    // Extract Markdown Extra abbreviation definitions (`*[HTML]: Hypertext Markup Language`)
    // before block splitting, for the same reason as footnotes: they conventionally live
    // at the end of the document.
    const abbreviationDefinitions = new Map<string, string>();
    textStr = textStr.replace(/^\*\[([^\]]+)\]:[ \t]*(.*)$/gm, (_match, abbr, definition) => {
        abbreviationDefinitions.set(abbr, definition.trim());
        return '';
    });

    // Extract link/image reference definitions (`[ref]: /url "title"`) before block
    // splitting, for the same reason as footnotes/abbreviations: they conventionally
    // live at the end of the document, after every place they're referenced. Keyed by
    // trimmed/lowercased label, matching CommonMark's case-insensitive reference matching.
    const linkDefinitions = new Map<string, { url: string; title?: string }>();
    textStr = textStr.replace(/^\[([^\]]+)\]:[ \t]*(\S+)(?:[ \t]+"([^"]*)")?[ \t]*$/gm, (_match, label, url, title) => {
        linkDefinitions.set(label.trim().toLowerCase(), { url, title });
        return '';
    });

    // Parses a Pandoc-style attribute list body (the part inside `{...}`), e.g.
    // `width=50% .centered` or `align=right`. Per MARKDOWN_DIALECT.md §15's Decisions,
    // the vocabulary matches ImageMetadata/TableMetadata's own width/align fields;
    // several class-name spellings are accepted on import for compatibility with
    // hand-written content, but the generator only ever emits canonical `align=value`.
    const parseAttributeList = (attrStr: string): { width?: string; align?: 'left' | 'center' | 'right' } => {
        const result: { width?: string; align?: 'left' | 'center' | 'right' } = {};
        for (const token of attrStr.trim().split(/\s+/).filter(Boolean)) {
            const kv = token.match(/^([a-zA-Z-]+)=(.+)$/);
            if (kv) {
                if (kv[1] === 'width') result.width = kv[2];
                else if (kv[1] === 'align' && ['left', 'center', 'right'].includes(kv[2])) result.align = kv[2] as any;
            } else if (token.startsWith('.')) {
                const cls = token.slice(1).toLowerCase();
                if (cls === 'left' || cls === 'align-left') result.align = 'left';
                else if (cls === 'center' || cls === 'centered' || cls === 'align-center') result.align = 'center';
                else if (cls === 'right' || cls === 'align-right') result.align = 'right';
            }
        }
        return result;
    };

    const parseInline = (text: string, currentFormatting: TextFormatting = {}): OfficeContentNode[] => {
        const nodes: OfficeContentNode[] = [];
        const plainText = (t: string): OfficeContentNode => ({ type: 'text', text: t, formatting: Object.keys(currentFormatting).length > 0 ? { ...currentFormatting } : undefined });

        // Builds the same image/link node shape regardless of whether the URL came from
        // an inline `(url)` or a resolved reference definition - shared by the inline
        // image/link branch and the two reference-style branches below.
        const buildLinkOrImageNodes = (isImage: boolean, altText: string, url: string, attrsStr?: string): OfficeContentNode[] => {
            if (isImage) {
                // Pandoc-style attribute list immediately after an image, e.g. {width=50% .centered}
                const attrs = attrsStr !== undefined ? parseAttributeList(attrsStr) : undefined;
                if (url.startsWith('data:')) {
                    const dataMatch = url.match(/^data:([^;]+);base64,(.*)$/);
                    if (dataMatch && config.extractAttachments) {
                        const mimeType = dataMatch[1] as any;
                        const data = dataMatch[2];
                        const name = `image_${attachments.length + 1}.${mimeType.split('/')[1]}`;
                        attachments.push({
                            type: 'image',
                            mimeType,
                            data,
                            name,
                            extension: mimeType.split('/')[1]
                        });
                        return [{ type: 'image', metadata: { attachmentName: name, altText, ...attrs } as ImageMetadata }];
                    }
                }
                return [{ type: 'image', metadata: { url, altText, ...attrs } as ImageMetadata }];
            }
            const linkNodes = parseInline(altText, currentFormatting);
            linkNodes.forEach(n => {
                if (n.type === 'text') {
                    n.metadata = { link: url, linkType: 'external' } as TextMetadata;
                }
            });
            return linkNodes;
        };

        // Regex matches (named groups): esc=escaped punctuation char | imgBang/imgAlt/imgUrl/imgAttrs=inline
        // image or link | boldStar/boldUnderscore=bold | italicStar/italicUnderscore=italic | strike=strikethrough |
        // codeFence/codeContent=inline code (backreferenced fence run, so a shorter embedded backtick run
        // doesn't close the span early) | underline/subscript/superscript=HTML tag formatting |
        // footnoteId | citationKey | wikiPage/wikiAlias | refBang/refText/refId=explicit or collapsed
        // reference link/image `[text][ref]`/`[text][]` | shortBang/shortText=shortcut reference `[text]`
        // (deliberately the most generic bracket pattern, so it must stay last among `[`-starting
        // alternatives) | autolinkUrl=`<url>` autolink | mathInline.
        //
        // Named groups (rather than positional match[N] indices) mean adding a new alternative never
        // requires renumbering every existing dispatch arm.
        //
        // Escape must be listed first since only a literal backslash can start that alternative, so it
        // never shadows another branch; but a code span's match consumes its whole span atomically (the
        // exec loop's lastIndex jumps past the entire matched span), so a backslash *inside* a code span
        // is never independently offered to the escape branch regardless of listing order - CommonMark's
        // "backslashes are not special inside code spans" rule holds by construction, not extra logic.
        //
        // Underscore emphasis has no CommonMark flanking-delimiter-run detection, so an intraword
        // underscore (e.g. "foo_bar_baz") will incorrectly italicize - an accepted, documented
        // simplification, not something this pass attempts to fix.
        //
        // Inline math requires no whitespace right after the opening $ or right before the
        // closing $, the common heuristic (matching Pandoc/KaTeX) for avoiding false
        // positives on currency like "$5 and $10".
        const regex = /\\(?<esc>[!-\/:-@\[-`{-~])|(?<imgBang>!?)\[(?<imgAlt>.*?)\]\((?<imgUrl>.*?)\)(?:\{(?<imgAttrs>[^}]*)\})?|\*\*(?<boldStar>.+?)\*\*|__(?<boldUnderscore>.+?)__|\*(?<italicStar>.+?)\*|_(?<italicUnderscore>.+?)_|~~(?<strike>.+?)~~|(?<codeFence>`+)(?<codeContent>(?:(?!\k<codeFence>)[\s\S])+?)\k<codeFence>(?!`)|<u>(?<underline>.+?)<\/u>|<sub>(?<subscript>.+?)<\/sub>|<sup>(?<superscript>.+?)<\/sup>|\[\^(?<footnoteId>[^\]]+)\]|\[@(?<citationKey>[a-zA-Z0-9_:.-]+)\]|\[\[(?<wikiPage>[^\]|]+)(?:\|(?<wikiAlias>[^\]]+))?\]\]|(?<refBang>!?)\[(?<refText>[^\]]*)\]\[(?<refId>[^\]]*)\]|(?<shortBang>!?)\[(?<shortText>[^\]]+)\]|<(?<autolinkUrl>(?:https?|mailto):[^\s<>]+)>|\$(?!\s)(?<mathInline>[^$\n]+?)(?<!\s)\$/g;
        let lastIndex = 0;
        let match;

        while ((match = regex.exec(text)) !== null) {
            if (match.index > lastIndex) {
                nodes.push(plainText(text.substring(lastIndex, match.index)));
            }
            const g = match.groups!;

            if (g.esc !== undefined) { // Backslash-escaped punctuation
                nodes.push(plainText(g.esc));
            } else if (g.imgAlt !== undefined) { // Image or Link
                nodes.push(...buildLinkOrImageNodes(g.imgBang === '!', g.imgAlt, g.imgUrl, g.imgAttrs));
            } else if (g.boldStar !== undefined) { // Bold (**)
                nodes.push(...parseInline(g.boldStar, { ...currentFormatting, bold: true }));
            } else if (g.boldUnderscore !== undefined) { // Bold (__)
                nodes.push(...parseInline(g.boldUnderscore, { ...currentFormatting, bold: true }));
            } else if (g.italicStar !== undefined) { // Italic (*)
                nodes.push(...parseInline(g.italicStar, { ...currentFormatting, italic: true }));
            } else if (g.italicUnderscore !== undefined) { // Italic (_)
                nodes.push(...parseInline(g.italicUnderscore, { ...currentFormatting, italic: true }));
            } else if (g.strike !== undefined) { // Strikethrough
                nodes.push(...parseInline(g.strike, { ...currentFormatting, strikethrough: true }));
            } else if (g.codeContent !== undefined) { // Inline code (any matching backtick-run length)
                nodes.push({ type: 'text', text: g.codeContent, formatting: { ...currentFormatting, font: 'monospace' } });
            } else if (g.underline !== undefined) { // Underline
                nodes.push(...parseInline(g.underline, { ...currentFormatting, underline: true }));
            } else if (g.subscript !== undefined) { // Subscript
                nodes.push(...parseInline(g.subscript, { ...currentFormatting, subscript: true }));
            } else if (g.superscript !== undefined) { // Superscript
                nodes.push(...parseInline(g.superscript, { ...currentFormatting, superscript: true }));
            } else if (g.footnoteId !== undefined) { // Footnote reference
                const noteId = g.footnoteId;
                const definition = footnoteDefinitions.get(noteId);
                const noteChildren = definition !== undefined ? parseInline(definition) : [];
                const noteNode: OfficeContentNode = {
                    type: 'note',
                    text: noteChildren.map(c => c.text || '').join(''),
                    children: noteChildren,
                    metadata: { noteType: 'footnote', noteId }
                };
                // Notes attach to the preceding text node (matches WordParser's convention);
                // fall back to an empty text node if the reference opens the inline run.
                if (nodes.length > 0) {
                    const target = nodes[nodes.length - 1];
                    if (!target.notes) target.notes = [];
                    target.notes.push(noteNode);
                } else {
                    nodes.push({ type: 'text', text: '', notes: [noteNode] });
                }
            } else if (g.citationKey !== undefined) { // Citation reference
                nodes.push({ type: 'text', text: g.citationKey, metadata: { citationKey: g.citationKey } as TextMetadata });
            } else if (g.wikiPage !== undefined) { // Wikilink
                const page = g.wikiPage.trim();
                const alias = g.wikiAlias?.trim();
                nodes.push({ type: 'text', text: alias || page, metadata: { link: page, linkType: 'internal', wikilink: true } as TextMetadata });
            } else if (g.refText !== undefined) { // Explicit/collapsed reference link or image: [text][ref] / [text][]
                const isImage = g.refBang === '!';
                const label = g.refText;
                const refId = (g.refId || label).trim().toLowerCase();
                const def = linkDefinitions.get(refId);
                if (def) {
                    nodes.push(...buildLinkOrImageNodes(isImage, label, def.url));
                } else {
                    // Not a known reference - preserve the literal bracketed text unchanged.
                    nodes.push(plainText(text.substring(match.index, match.index + match[0].length)));
                }
            } else if (g.shortText !== undefined) { // Shortcut reference: [text]
                const isImage = g.shortBang === '!';
                const label = g.shortText;
                const def = linkDefinitions.get(label.trim().toLowerCase());
                if (def) {
                    nodes.push(...buildLinkOrImageNodes(isImage, label, def.url));
                } else {
                    // Not a known reference - ordinary bracketed prose, preserve unchanged.
                    nodes.push(plainText(`${g.shortBang}[${label}]`));
                }
            } else if (g.autolinkUrl !== undefined) { // <url> autolink
                const url = g.autolinkUrl;
                nodes.push({ type: 'text', text: url, formatting: Object.keys(currentFormatting).length > 0 ? { ...currentFormatting } : undefined, metadata: { link: url, linkType: 'external' } as TextMetadata });
            } else if (g.mathInline !== undefined) { // Inline math
                nodes.push({ type: 'code', text: g.mathInline, metadata: { math: 'inline' } as CodeMetadata });
            }

            lastIndex = regex.lastIndex;
        }

        if (lastIndex < text.length) {
            nodes.push(plainText(text.substring(lastIndex)));
        }

        return applyAbbreviations(decodeHtmlEntities(nodes));
    };

    const escapeRegExpChars = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

    // A deliberately small, common-entity lookup (not the full HTML5 named-character-
    // reference table) - keeps this a plain object rather than needing a dependency.
    const NAMED_HTML_ENTITIES: Record<string, string> = {
        amp: '&', lt: '<', gt: '>', quot: '"', apos: '\'', nbsp: ' ',
        copy: '©', reg: '®', mdash: '—', ndash: '–', hellip: '…'
    };

    // Decodes HTML named entities and numeric/hex character references (&#NN;/&#xHH;)
    // in plain text nodes, skipping monospace (inline code) nodes since CommonMark does
    // not decode entities inside code spans. The regex only ever matches syntactically
    // well-formed &name;/&#NN;/&#xHH; tokens to begin with, so ordinary text containing
    // a bare "&" (e.g. "Q&A", "Fish & Chips") never matches at all; an unrecognized-but-
    // well-formed token (e.g. "&foo;") is left untouched on a lookup miss - no risk of
    // double-decoding or corrupting text that merely resembles an entity.
    const decodeHtmlEntities = (nodes: OfficeContentNode[]): OfficeContentNode[] => {
        return nodes.map(node => {
            if (node.type !== 'text' || !node.text || node.formatting?.font === 'monospace') return node;
            const text = node.text.replace(/&(#\d+|#[xX][0-9a-fA-F]+|[a-zA-Z][a-zA-Z0-9]*);/g, (full: string, ref: string) => {
                if (ref[0] === '#') {
                    const codePoint = ref[1].toLowerCase() === 'x' ? parseInt(ref.slice(2), 16) : parseInt(ref.slice(1), 10);
                    return (isNaN(codePoint) || codePoint < 0 || codePoint > 0x10FFFF) ? full : String.fromCodePoint(codePoint);
                }
                return NAMED_HTML_ENTITIES[ref] ?? full;
            });
            return text === node.text ? node : { ...node, text };
        });
    };

    // Splits abbreviation occurrences out of plain text nodes so they carry
    // TextMetadata.abbreviationTitle, rendered as <abbr title> in HTML/editor output.
    const applyAbbreviations = (nodes: OfficeContentNode[]): OfficeContentNode[] => {
        if (abbreviationDefinitions.size === 0) return nodes;
        const pattern = new RegExp(`\\b(${[...abbreviationDefinitions.keys()].map(escapeRegExpChars).join('|')})\\b`, 'g');

        const result: OfficeContentNode[] = [];
        for (const node of nodes) {
            if (node.type !== 'text' || !node.text || node.metadata) {
                result.push(node);
                continue;
            }

            let lastIndex = 0;
            let match: RegExpExecArray | null;
            let matched = false;
            pattern.lastIndex = 0;
            while ((match = pattern.exec(node.text)) !== null) {
                matched = true;
                if (match.index > lastIndex) {
                    result.push({ type: 'text', text: node.text.substring(lastIndex, match.index), formatting: node.formatting });
                }
                result.push({
                    type: 'text',
                    text: match[0],
                    formatting: node.formatting,
                    metadata: { abbreviationTitle: abbreviationDefinitions.get(match[0]) } as TextMetadata
                });
                lastIndex = pattern.lastIndex;
            }

            if (!matched) {
                result.push(node);
                continue;
            }
            if (lastIndex < node.text.length) {
                result.push({ type: 'text', text: node.text.substring(lastIndex), formatting: node.formatting });
            }
        }
        return result;
    };

    // Splits a paragraph-shaped block's internal lines into inline-parsed content,
    // inserting a real 'break' node for a hard line break (a line ending in 2+ trailing
    // spaces or a trailing backslash) instead of collapsing it to a space. A plain single
    // newline with no such marker is still a soft break and collapses to a space,
    // unchanged from before - CommonMark itself renders a soft break as a space/newline.
    const splitParagraphLines = (block: string): OfficeContentNode[] => {
        const lines = block.split('\n');
        const children: OfficeContentNode[] = [];
        lines.forEach((line, i) => {
            const hardBreak = /(?: {2,}|\\)$/.test(line);
            children.push(...parseInline(line.replace(/(?: {2,}|\\)$/, '')));
            if (i < lines.length - 1) {
                if (hardBreak) {
                    children.push({ type: 'break', metadata: { breakType: 'carriageReturn' } as BreakMetadata });
                } else {
                    children.push({ type: 'text', text: ' ' });
                }
            }
        });
        return children;
    };

    // Builds an admonition node from its raw body text, splitting on blank lines into
    // paragraph children. v1 only supports inline content inside admonitions (no nested
    // lists/headings/code) - acceptable per the roadmap's first cut.
    const buildAdmonitionNode = (admonitionType: AdmonitionMetadata['admonitionType'], body: string, sourceSyntax: 'github' | 'gitlab'): OfficeContentNode => {
        const paragraphs = body.split(/\n\n+/).map(p => p.trim()).filter(Boolean);
        const children: OfficeContentNode[] = paragraphs.map(p => ({
            type: 'paragraph',
            children: splitParagraphLines(p)
        }));
        return {
            type: 'admonition',
            metadata: { admonitionType, sourceSyntax } as AdmonitionMetadata,
            children
        };
    };

    const rawBlocks = textStr.split(/\n\n+/);
    const blocks: string[] = [];

    // Sub-split blocks that contain headings or lists without double newlines
    for (const rawBlock of rawBlocks) {
        if (!rawBlock.trim()) continue;

        // Match headings or lists that might be joined with other text via single newline
        const lines = rawBlock.split('\n');
        let currentSubBlock: string[] = [];
        // Tracks whether we're currently "inside" a list (a list-item line, or an
        // indented continuation line right after one) so a continuation line doesn't
        // itself get treated as the boundary that splits the list into a new block -
        // see the "Lists" block dispatch below, which merges such a line into the
        // previous item's content instead of dropping it.
        let inList: boolean = false;
        for (const line of lines) {
            const isHeading = !!line.match(/^(?:<a[^>]*><\/a>)*\s*#{1,6}\s+/);
            const isList = !!line.match(/^(\s*)([-*+]|\d+[.)])\s+/);
            const isHtmlTag = !!line.match(/^<\/?div[^>]*>$/i);
            // A non-list, non-blank, indented (>=2 columns or a tab) line encountered
            // while already inside a list is a continuation of the current item, not a
            // new construct. Scoped to a single such line at a time (no nested
            // code/blockquote/sub-list/multi-paragraph items - those require un-splitting
            // already-separated raw blocks, out of scope here).
            const isContinuation: boolean = !isList && inList && /^(?: {2,}|\t)/.test(line) && line.trim().length > 0;
            const staysInListMode: boolean = isList || isContinuation;

            // Split if:
            // 1. Current line is a heading
            // 2. Current line enters or leaves "list mode" relative to the previous line
            // 3. Current line is an HTML tag (div)
            if ((isHeading || isHtmlTag || (staysInListMode !== inList)) && currentSubBlock.length > 0) {
                blocks.push(currentSubBlock.join('\n'));
                currentSubBlock = [];
            }

            currentSubBlock.push(line);
            inList = staysInListMode;

            // Headings and HTML tags are single-line blocks for our state machine
            if (isHeading || isHtmlTag) {
                blocks.push(currentSubBlock.join('\n'));
                currentSubBlock = [];
                inList = false;
            }
        }
        if (currentSubBlock.length > 0) {
            blocks.push(currentSubBlock.join('\n'));
        }
    }

    let listIdCounter = 1;
    let currentAlignment: 'left' | 'center' | 'right' | 'justify' | undefined = undefined;

    for (let block of blocks) {
        // Preserved before the generic trim() below, which strips the first line's
        // leading indentation - the indented-code-block check further down needs every
        // line's original indentation, including the first.
        const untrimmedBlock = block;
        block = block.trim();
        if (!block) continue;

        // Standalone anchor-only block: one or more empty `<a name|id="…"></a>` tags on their
        // own line (bookmark targets the MarkdownGenerator emits just before a heading/paragraph).
        // Capture them as a placeholder so the post-loop pass can re-attach them to the following
        // node's anchorIds — otherwise the tag-opening `<` is escaped and they render as visible text.
        if (/^(?:\s*<a\s[^>]*>\s*<\/a>\s*)+$/i.test(block)) {
            const anchorIds: string[] = [];
            for (const m of block.matchAll(/<a\s[^>]*\b(?:name|id)="([^"]*)"/gi)) {
                if (m[1]) anchorIds.push(m[1]);
            }
            if (anchorIds.length > 0) {
                content.push({ type: ANCHOR_PLACEHOLDER as any, metadata: { anchorIds } as any, children: [] });
                continue;
            }
        }

        // Check for alignment wrapper start/end
        const alignStartMatch = block.match(/^<div\s+(?:style="text-align:\s*(left|center|right|justify);?"|align="(left|center|right|justify)")>$/i);
        if (alignStartMatch) {
            currentAlignment = (alignStartMatch[1] || alignStartMatch[2]).toLowerCase() as any;
            continue;
        }
        if (block.match(/^<\/div>$/i)) {
            currentAlignment = undefined;
            continue;
        }

        let alignment = currentAlignment;
        // Check for single-line alignment wrapper (for compatibility)
        const alignMatch = block.match(/^<div\s+(?:style="text-align:\s*(left|center|right|justify);?"|align="(left|center|right|justify)")>\s*([\s\S]*?)\s*<\/div>$/i);
        if (alignMatch) {
            alignment = (alignMatch[1] || alignMatch[2]).toLowerCase() as any;
            block = alignMatch[3];
        }

        // YouTube embed fallback: MarkdownGenerator's 'embed' case emits a single-line
        // <div data-youtube-video="ID" data-width="…" data-align="…"></div> when fallbackToHtml
        // is on; recognise it here so a saved-then-reopened .md keeps the video.
        const youtubeMatch = block.match(/^<div\s+data-youtube-video="([^"]*)"([^>]*)>\s*<\/div>$/i);
        if (youtubeMatch) {
            const videoId = youtubeMatch[1];
            const attrsStr = youtubeMatch[2];
            const widthMatch = attrsStr.match(/data-width="([^"]*)"/i);
            const youtubeAlignMatch = attrsStr.match(/data-align="([^"]*)"/i);
            const embedAlign = youtubeAlignMatch && (['left', 'center', 'right'] as const).includes(youtubeAlignMatch[1] as any) ? youtubeAlignMatch[1] as 'left' | 'center' | 'right' : undefined;
            const embedUrl = videoId ? `https://www.youtube.com/watch?v=${videoId}` : undefined;
            content.push({
                type: 'embed',
                // Childless nodes need .text so generic AST consumers (toText, chunking)
                // don't silently drop them.
                text: embedUrl,
                metadata: {
                    embedType: 'youtube',
                    videoId,
                    url: embedUrl,
                    width: widthMatch?.[1],
                    align: embedAlign
                } as EmbedMetadata
            });
            continue;
        }

        // Code Block
        const codeMatch = block.match(/^__CODE_BLOCK_(\d+)__$/);
        if (codeMatch) {
            const data = JSON.parse(codeBlocks[parseInt(codeMatch[1])]);
            content.push({
                type: 'code',
                text: data.code,
                metadata: { language: data.lang } as CodeMetadata
            });
            continue;
        }

        // GLFM-style fenced-div admonition, extracted to a placeholder above
        const admonitionBlockMatch = block.match(/^__ADMONITION_(\d+)__$/);
        if (admonitionBlockMatch) {
            const data = JSON.parse(admonitionBlocks[parseInt(admonitionBlockMatch[1])]);
            content.push(buildAdmonitionNode(data.admonitionType, data.body, 'gitlab'));
            continue;
        }

        // Block math ($$...$$), extracted to a placeholder above
        const mathBlockMatch = block.match(/^__MATH_BLOCK_(\d+)__$/);
        if (mathBlockMatch) {
            content.push({
                type: 'code',
                text: mathBlocks[parseInt(mathBlockMatch[1])],
                metadata: { math: 'block' } as CodeMetadata
            });
            continue;
        }

        // Heading (allowing for leading HTML anchors and trailing {#anchor})
        const headingMatch = block.match(/^((?:<a[^>]*><\/a>)*)\s*(#{1,6})\s+(.*?)(?:\s+\{#([^}]+)\})?\s*$/s);
        if (headingMatch) {
            const leadingAnchorsRaw = headingMatch[1];
            const rawText = headingMatch[3];
            const explicitAnchor = headingMatch[4];

            const anchorIds: string[] = [];
            if (leadingAnchorsRaw) {
                const idMatches = leadingAnchorsRaw.matchAll(/<a\s[^>]*\b(?:name|id)="([^"]+)"/gi);
                for (const m of idMatches) anchorIds.push(m[1]);
            }
            if (explicitAnchor) anchorIds.push(explicitAnchor);

            const children = parseInline(rawText);
            content.push({
                type: 'heading',
                text: children.map(c => c.text || '').join(''),
                metadata: {
                    level: headingMatch[2].length,
                    alignment,
                    anchorIds: anchorIds.length > 0 ? anchorIds : undefined
                } as HeadingMetadata,
                children
            });
            continue;
        }

        // Setext heading (Text\n===  or  Text\n---): a line of text immediately followed
        // by a lone `=`/`-` underline with no blank line between them. By the time a
        // block reaches this point, the sub-splitter above has already separated out any
        // genuinely blank-line-preceded thematic break into its own isolated block (which
        // has no preceding text line to combine with here), so this only fires for the
        // ambiguous "text directly above a dash/equals-only line" shape setext needs.
        // Scoped to a single line immediately above the underline becoming the heading
        // text; multi-line setext text (CommonMark's "Foo\nbar\n===" merging into one
        // heading) is an explicitly out-of-scope simplification - any earlier lines in
        // the block are pushed as a separate paragraph first.
        const setextMatch = block.match(/^([\s\S]*)\n([=]+|-+)[ \t]*$/);
        if (setextMatch) {
            const lines = setextMatch[1].split('\n');
            const headingLine = lines[lines.length - 1];
            const earlierLines = lines.slice(0, -1).join('\n').trim();
            if (earlierLines) {
                content.push({
                    type: 'paragraph',
                    metadata: { alignment } as any,
                    children: splitParagraphLines(earlierLines)
                });
            }
            const children = parseInline(headingLine);
            content.push({
                type: 'heading',
                text: children.map(c => c.text || '').join(''),
                metadata: { level: setextMatch[2][0] === '=' ? 1 : 2, alignment } as HeadingMetadata,
                children
            });
            continue;
        }

        // Blockquote
        const quoteMatch = block.match(/^>\s+(.*)$/s);
        if (quoteMatch) {
            // [ \t]? (not \s+) so a bare ">" paragraph-separator line (used between
            // multi-paragraph admonition bodies) also dequotes to an empty line. Repeat
            // until no line still starts with ">" so arbitrarily-nested blockquotes
            // (`> > quoted`, `> > > quoted`, ...) are fully unwrapped rather than only
            // stripping one level.
            let dequoted = quoteMatch[1];
            while (/^>/m.test(dequoted)) {
                dequoted = dequoted.replace(/^>[ \t]?/gm, '');
            }

            // GitHub-style admonition: `> [!NOTE]` on the first quoted line.
            const admonitionHeaderMatch = dequoted.match(/^\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*\n?([\s\S]*)$/i);
            if (admonitionHeaderMatch) {
                const admonitionType = admonitionHeaderMatch[1].toLowerCase() as AdmonitionMetadata['admonitionType'];
                content.push(buildAdmonitionNode(admonitionType, admonitionHeaderMatch[2], 'github'));
                continue;
            }

            content.push({
                type: 'paragraph',
                metadata: { style: 'Quote' } as any,
                children: parseInline(dequoted)
            });
            continue;
        }

        // Definition list (Markdown Extra / Pandoc / Kramdown): a term line followed by
        // one or more ": definition" lines, e.g.:
        //   Term
        //   : Definition of the term.
        const definitionListMatch = block.match(/^([^\n:][^\n]*)\n((?::[ \t]+.+(?:\n:[ \t]+.+)*))$/);
        if (definitionListMatch) {
            const term = definitionListMatch[1];
            const definitions = definitionListMatch[2].split('\n').map(line => line.replace(/^:[ \t]+/, ''));
            content.push({
                type: 'definitionList',
                children: [
                    { type: 'definitionTerm', children: parseInline(term) },
                    ...definitions.map(def => ({ type: 'definitionDescription' as const, children: parseInline(def) }))
                ]
            });
            continue;
        }

        // Lists
        if (block.match(/^(\s*)([-*+]|\d+[.)])\s+/)) {
            const lines = block.split('\n');
            const listId = `md-list-${listIdCounter++}`;
            const listCounters = new Map<number, number>();
            // Relative indent stack (not a fixed-width divisor) so nesting level is
            // computed from what indentation actually appeared in this block, rather
            // than assuming a specific indent width. This makes the parser agnostic to
            // 2-space (hand-written), 4-space (this generator's own output), or
            // tab-indented (normalized to a 4-column stop) nested lists.
            const indentStack: number[] = [];
            // The most recently pushed list-item node, so a following indented
            // continuation line (see the sub-splitter above) can be merged into it
            // instead of being silently dropped.
            let lastListNode: OfficeContentNode | undefined;

            for (const line of lines) {
                const match = line.match(/^(\s*)([-*+]|\d+[.)])\s+(.*)$/);
                if (match) {
                    const rawIndent = match[1].replace(/\t/g, '    ').length;
                    while (indentStack.length > 0 && rawIndent <= indentStack[indentStack.length - 1]) {
                        indentStack.pop();
                    }
                    const level = indentStack.length;
                    indentStack.push(rawIndent);

                    // Purge any deeper levels' counters now that we're back at this
                    // level - otherwise a nested sub-list under a later sibling item
                    // would incorrectly continue a previous sibling's child numbering
                    // instead of restarting at 0.
                    for (const key of [...listCounters.keys()]) {
                        if (key > level) listCounters.delete(key);
                    }

                    const marker = match[2];
                    const isOrdered = !!marker.match(/\d+[.)]/);
                    const listType: 'ordered' | 'unordered' = isOrdered ? 'ordered' : 'unordered';

                    if (listCounters.get(level) === undefined) {
                        if (isOrdered) {
                            const startNum = parseInt(marker, 10);
                            listCounters.set(level, isNaN(startNum) ? 0 : startNum - 1);
                        } else {
                            listCounters.set(level, 0);
                        }
                    } else {
                        listCounters.set(level, listCounters.get(level)! + 1);
                    }

                    let itemText = match[3];
                    let isTask: boolean | undefined;
                    let checked: boolean | undefined;
                    const taskMatch = itemText.match(/^\[([ xX])\]\s+(.*)$/);
                    if (taskMatch) {
                        isTask = true;
                        checked = taskMatch[1].toLowerCase() === 'x';
                        itemText = taskMatch[2];
                    }

                    const children = parseInline(itemText);
                    const listNode: OfficeContentNode = {
                        type: 'list',
                        text: children.map(c => c.text || '').join(''),
                        metadata: {
                            listType,
                            indentation: level,
                            alignment: alignment || 'left',
                            listId,
                            itemIndex: listCounters.get(level),
                            isTask,
                            checked
                        } as ListMetadata,
                        children
                    };
                    content.push(listNode);
                    lastListNode = listNode;
                } else if (lastListNode && line.trim().length > 0 && /^(?: {2,}|\t)/.test(line)) {
                    // Indented continuation line: merge its inline content into the
                    // previous item rather than dropping it. Scoped to a single such
                    // line (no nested code/blockquote/sub-list/multi-paragraph items).
                    const continuationChildren = parseInline(line.trim());
                    lastListNode.children = [...(lastListNode.children || []), { type: 'text', text: ' ' }, ...continuationChildren];
                    lastListNode.text = (lastListNode.children || []).map(c => c.text || '').join('');
                }
            }
            continue;
        }

        // Table (Simple Pipe or HTML)
        if ((block.includes('|') && block.match(/\n\s*\|?[-:| ]+\|?\s*\n/)) || block.includes('<table')) {
            // Pandoc-style trailing attribute list (`{align=right}`) immediately after the
            // table, or Kramdown's `{: align=right}` on its own following line - both land
            // in this same raw block since there's no blank line separating them.
            let tableAlign: 'left' | 'center' | 'right' | undefined;
            const tableAttrLineMatch = block.match(/\n\{:?\s*([^}]*)\}\s*$/);
            if (tableAttrLineMatch) {
                tableAlign = parseAttributeList(tableAttrLineMatch[1]).align;
                block = block.slice(0, tableAttrLineMatch.index);
            }

            if (block.includes('<table')) {
                // Basic HTML table recognition (extracting rows/cells)
                const tableTagMatch = block.match(/<table([^>]*)>/i);
                const tableAlignMatch = tableTagMatch?.[1]?.match(/data-align=["']?(left|center|right)["']?/i);

                const rows: OfficeContentNode[] = [];
                const trRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
                let trMatch;
                while ((trMatch = trRegex.exec(block)) !== null) {
                    const tdRegex = /<(?:td|th)([^>]*)>([\s\S]*?)<\/(?:td|th)>/gi;
                    let tdMatch;
                    const cells: OfficeContentNode[] = [];
                    while ((tdMatch = tdRegex.exec(trMatch[1])) !== null) {
                        const attrs = tdMatch[1];
                        const contentStr = tdMatch[2].trim();
                        const colSpanMatch = attrs.match(/colspan=["']?(\d+)["']?/i);
                        const rowSpanMatch = attrs.match(/rowspan=["']?(\d+)["']?/i);

                        cells.push({
                            type: 'cell',
                            metadata: {
                                colSpan: colSpanMatch ? parseInt(colSpanMatch[1]) : undefined,
                                rowSpan: rowSpanMatch ? parseInt(rowSpanMatch[1]) : undefined
                            } as any,
                            children: parseInline(contentStr.replace(/<[^>]*>/g, ''))
                        });
                    }
                    if (cells.length > 0) rows.push({ type: 'row', children: cells });
                }
                const resolvedAlign = tableAlign || (tableAlignMatch ? tableAlignMatch[1].toLowerCase() as 'left' | 'center' | 'right' : undefined);
                if (rows.length > 0) {
                    content.push({
                        type: 'table',
                        metadata: resolvedAlign ? { align: resolvedAlign } : undefined,
                        children: rows
                    });
                    continue;
                }
            } else {
                const lines = block.trim().split('\n');
                const rows: OfficeContentNode[] = [];
                for (let i = 0; i < lines.length; i++) {
                    if (lines[i].match(/^\|?\s*:?-+:?\s*(?:\|\s*:?-+:?\s*)*\|?$/)) continue; // Separator row (per-cell `:?-+:?`, GFM-style; accepts short cells like `|-|-|`)

                    const cellsStr = lines[i].replace(/^\||\|$/g, '').split('|');
                    const cells: OfficeContentNode[] = cellsStr.map(c => {
                        // Recognize the MarkdownGenerator's own cell-alignment fallback,
                        // `<div style="text-align: X">…</div>`, and lift it into an aligned
                        // paragraph so it round-trips as alignment instead of being escaped to
                        // visible text on regeneration. Unwrap wherever it sits (e.g. inside **…**).
                        let cellText = c.trim();
                        let cellAlign: 'left' | 'center' | 'right' | 'justify' | undefined;
                        cellText = cellText.replace(
                            /<div\s+style="text-align:\s*(left|center|right|justify);?"\s*>([\s\S]*?)<\/div>/gi,
                            (_m, a: string, inner: string) => { cellAlign = a.toLowerCase() as any; return inner; }
                        );
                        const inline = parseInline(cellText, i === 0 ? { bold: true } : {});
                        if (cellAlign && cellAlign !== 'left') {
                            return {
                                type: 'cell',
                                children: [{ type: 'paragraph', metadata: { alignment: cellAlign } as any, children: inline }]
                            } as OfficeContentNode;
                        }
                        return { type: 'cell', children: inline } as OfficeContentNode;
                    });
                    rows.push({ type: 'row', children: cells });
                }
                content.push({ type: 'table', metadata: tableAlign ? { align: tableAlign } : undefined, children: rows });
                continue;
            }
        }

        // Indented code block (4-space or tab indent on every non-blank line). Only
        // reaches this point once heading/blockquote/definition-list/list/table have
        // already failed to claim the block; since list continuation lines are now
        // handled inside the "Lists" branch above and the sub-splitter already isolates
        // list/heading content into their own blocks, a block that's uniformly indented
        // here is not a list by construction. A partially-indented block (some lines
        // indented, some not) falls through to Paragraph unchanged.
        {
            const codeLines = untrimmedBlock.split('\n');
            const nonBlankLines = codeLines.filter(l => l.trim().length > 0);
            if (nonBlankLines.length > 0 && nonBlankLines.every(l => /^(?: {4}|\t)/.test(l))) {
                const stripped = codeLines.map(l => l.replace(/^(?: {4}|\t)/, '')).join('\n');
                content.push({ type: 'code', text: stripped });
                continue;
            }
        }

        // Hr
        if (block.match(/^---+$|^\*\*\*+$|^___+$/)) {
            content.push({ type: 'break', metadata: { breakType: 'page' } });
            continue;
        }

        // Paragraph
        content.push({
            type: 'paragraph',
            metadata: { alignment } as any,
            children: splitParagraphLines(block)
        });
    }

    // Fold standalone anchor placeholders into the following content node's anchorIds so a
    // bookmark target emitted on its own line round-trips as a real anchor. A trailing placeholder
    // with no following node attaches to the previous node instead; if the document is nothing but
    // anchors, they are dropped (there is no node to host them).
    if (content.some(n => (n.type as any) === ANCHOR_PLACEHOLDER)) {
        const merged: OfficeContentNode[] = [];
        let carried: string[] = [];
        for (const node of content) {
            if ((node.type as any) === ANCHOR_PLACEHOLDER) {
                carried.push(...(((node.metadata as any)?.anchorIds as string[]) || []));
                continue;
            }
            if (carried.length > 0) {
                const meta: any = node.metadata || (node.metadata = {} as any);
                meta.anchorIds = [...carried, ...((meta.anchorIds as string[]) || [])];
                carried = [];
            }
            merged.push(node);
        }
        if (carried.length > 0 && merged.length > 0) {
            const last: any = merged[merged.length - 1].metadata || (merged[merged.length - 1].metadata = {} as any);
            last.anchorIds = [...((last.anchorIds as string[]) || []), ...carried];
        }
        content.length = 0;
        content.push(...merged);
    }

    const toTextSync = () => content.map(n => {
        const getText = (node: OfficeContentNode): string => {
            if (node.type === 'text' || node.type === 'code') return node.text || '';
            if (node.type === 'break') return '\n';
            // Childless nodes still carry meaningful text - fall back to it instead of
            // silently vanishing from plain-text/RAG-chunk output.
            if (node.type === 'embed') return (node.metadata as EmbedMetadata)?.url || '';
            if (node.children) {
                const isBlock = ['table', 'row', 'list', 'sheet', 'slide', 'admonition', 'definitionList'].includes(node.type);
                return node.children.map(getText).join(isBlock ? config.newlineDelimiter : '');
            }
            return '';
        };
        return getText(n);
    }).join(config.newlineDelimiter)
        .replace(/\n{3,}/g, '\n\n'); // Normalize excessive whitespace

    return createAST('md', metadata, content, attachments, config, undefined, toTextSync);
};

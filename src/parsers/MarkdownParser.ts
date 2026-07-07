import { AdmonitionMetadata, CodeMetadata, EmbedMetadata, FullOfficeParserConfig, HeadingMetadata, ImageMetadata, ListMetadata, OfficeAttachment, OfficeContentNode, OfficeMetadata, OfficeParserAST, TextFormatting, TextMetadata } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { checkAbortSignal } from '../utils/errorUtils.js';

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

    // Extract code blocks first to protect their contents
    const codeBlocks: string[] = [];
    textStr = textStr.replace(/^```(\w*)\n([\s\S]*?)\n```/gm, (match, lang, code) => {
        const id = `__CODE_BLOCK_${codeBlocks.length}__`;
        codeBlocks.push(JSON.stringify({ lang, code }));
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

        // Regex matches: 1=!, 2=alt, 3=url, 4=attrs | 5=bold | 6=italic | 7=strike | 8=code | 9=underline | 10=subscript | 11=superscript | 12=footnote id
        const regex = /(!?)\[(.*?)\]\((.*?)\)(?:\{([^}]*)\})?|\*\*(.+?)\*\*|\*(.+?)\*|~~(.+?)~~|`(.+?)`|<u>(.+?)<\/u>|<sub>(.+?)<\/sub>|<sup>(.+?)<\/sup>|\[\^([^\]]+)\]/g;
        let lastIndex = 0;
        let match;

        while ((match = regex.exec(text)) !== null) {
            if (match.index > lastIndex) {
                nodes.push({ type: 'text', text: text.substring(lastIndex, match.index), formatting: Object.keys(currentFormatting).length > 0 ? { ...currentFormatting } : undefined });
            }

            if (match[2] !== undefined) { // Image or Link
                const isImage = match[1] === '!';
                const altText = match[2];
                const url = match[3];
                // Pandoc-style attribute list immediately after an image, e.g. {width=50% .centered}
                const attrs = isImage && match[4] !== undefined ? parseAttributeList(match[4]) : undefined;
                if (isImage) {
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
                            nodes.push({ type: 'image', metadata: { attachmentName: name, altText, ...attrs } as ImageMetadata });
                        } else {
                            nodes.push({ type: 'image', metadata: { url, altText, ...attrs } as ImageMetadata });
                        }
                    } else {
                        nodes.push({ type: 'image', metadata: { url, altText, ...attrs } as ImageMetadata });
                    }
                } else {
                    const linkNodes = parseInline(altText, currentFormatting);
                    linkNodes.forEach(n => {
                        if (n.type === 'text') {
                            n.metadata = { link: url, linkType: 'external' } as TextMetadata;
                        }
                    });
                    nodes.push(...linkNodes);
                }
            } else if (match[5]) { // Bold
                nodes.push(...parseInline(match[5], { ...currentFormatting, bold: true }));
            } else if (match[6]) { // Italic
                nodes.push(...parseInline(match[6], { ...currentFormatting, italic: true }));
            } else if (match[7]) { // Strikethrough
                nodes.push(...parseInline(match[7], { ...currentFormatting, strikethrough: true }));
            } else if (match[8]) { // Inline Code
                nodes.push({ type: 'text', text: match[8], formatting: { ...currentFormatting, font: 'monospace' } });
            } else if (match[9]) { // Underline
                nodes.push(...parseInline(match[9], { ...currentFormatting, underline: true }));
            } else if (match[10]) { // Subscript
                nodes.push(...parseInline(match[10], { ...currentFormatting, subscript: true }));
            } else if (match[11]) { // Superscript
                nodes.push(...parseInline(match[11], { ...currentFormatting, superscript: true }));
            } else if (match[12]) { // Footnote reference
                const noteId = match[12];
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
            }

            lastIndex = regex.lastIndex;
        }

        if (lastIndex < text.length) {
            nodes.push({ type: 'text', text: text.substring(lastIndex), formatting: Object.keys(currentFormatting).length > 0 ? { ...currentFormatting } : undefined });
        }

        return applyAbbreviations(nodes);
    };

    const escapeRegExpChars = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

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

    // Builds an admonition node from its raw body text, splitting on blank lines into
    // paragraph children. v1 only supports inline content inside admonitions (no nested
    // lists/headings/code) - acceptable per the roadmap's first cut.
    const buildAdmonitionNode = (admonitionType: AdmonitionMetadata['admonitionType'], body: string): OfficeContentNode => {
        const paragraphs = body.split(/\n\n+/).map(p => p.trim()).filter(Boolean);
        const children: OfficeContentNode[] = paragraphs.map(p => ({
            type: 'paragraph',
            children: parseInline(p.replace(/\n/g, ' '))
        }));
        return {
            type: 'admonition',
            metadata: { admonitionType } as AdmonitionMetadata,
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
        for (const line of lines) {
            const isHeading = !!line.match(/^(?:<a[^>]*><\/a>)*\s*#{1,6}\s+/);
            const isList = !!line.match(/^(\s*)([-*+]|\d+\.)\s+/);
            const isHtmlTag = !!line.match(/^<\/?div[^>]*>$/i);
            const prevWasList = currentSubBlock.length > 0 && !!currentSubBlock[currentSubBlock.length - 1].match(/^(\s*)([-*+]|\d+\.)\s+/);

            // Split if:
            // 1. Current line is a heading
            // 2. Current line is a list item but previous was NOT
            // 3. Current line is NOT a list item but previous WAS
            // 4. Current line is an HTML tag (div)
            if ((isHeading || isHtmlTag || (isList !== prevWasList)) && currentSubBlock.length > 0) {
                blocks.push(currentSubBlock.join('\n'));
                currentSubBlock = [];
            }

            currentSubBlock.push(line);

            // Headings and HTML tags are single-line blocks for our state machine
            if (isHeading || isHtmlTag) {
                blocks.push(currentSubBlock.join('\n'));
                currentSubBlock = [];
            }
        }
        if (currentSubBlock.length > 0) {
            blocks.push(currentSubBlock.join('\n'));
        }
    }

    let listIdCounter = 1;
    let currentAlignment: 'left' | 'center' | 'right' | 'justify' | undefined = undefined;

    for (let block of blocks) {
        block = block.trim();
        if (!block) continue;

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
            content.push(buildAdmonitionNode(data.admonitionType, data.body));
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
                const idMatches = leadingAnchorsRaw.matchAll(/<a\s+name="([^"]+)"/gi);
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

        // Blockquote
        const quoteMatch = block.match(/^>\s+(.*)$/s);
        if (quoteMatch) {
            // [ \t]? (not \s+) so a bare ">" paragraph-separator line (used between
            // multi-paragraph admonition bodies) also dequotes to an empty line.
            const dequoted = quoteMatch[1].replace(/^>[ \t]?/gm, '');

            // GitHub-style admonition: `> [!NOTE]` on the first quoted line.
            const admonitionHeaderMatch = dequoted.match(/^\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*\n?([\s\S]*)$/i);
            if (admonitionHeaderMatch) {
                const admonitionType = admonitionHeaderMatch[1].toLowerCase() as AdmonitionMetadata['admonitionType'];
                content.push(buildAdmonitionNode(admonitionType, admonitionHeaderMatch[2]));
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
        if (block.match(/^(\s*)([-*+]|\d+\.)\s+/)) {
            const lines = block.split('\n');
            const listId = `md-list-${listIdCounter++}`;
            const listCounters: Record<number, number> = {};

            for (const line of lines) {
                const match = line.match(/^(\s*)([-*+]|\d+\.)\s+(.*)$/);
                if (match) {
                    const indent = match[1].length / 2;
                    const level = Math.floor(indent);
                    const marker = match[2];
                    const isOrdered = !!marker.match(/\d+\./);
                    const listType: 'ordered' | 'unordered' = isOrdered ? 'ordered' : 'unordered';

                    if (listCounters[level] === undefined) {
                        if (isOrdered) {
                            const startNum = parseInt(marker, 10);
                            listCounters[level] = isNaN(startNum) ? 0 : startNum - 1;
                        } else {
                            listCounters[level] = 0;
                        }
                    } else {
                        listCounters[level]++;
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
                    content.push({
                        type: 'list',
                        text: children.map(c => c.text || '').join(''),
                        metadata: {
                            listType,
                            indentation: level,
                            alignment: alignment || 'left',
                            listId,
                            itemIndex: listCounters[level],
                            isTask,
                            checked
                        } as ListMetadata,
                        children
                    });
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
                    if (lines[i].match(/^\|?[-:| ]*---[-:| ]*\|?$/)) continue; // Separator row (requires at least one triple-hyphen)

                    const cellsStr = lines[i].replace(/^\||\|$/g, '').split('|');
                    const cells: OfficeContentNode[] = cellsStr.map(c => ({
                        type: 'cell',
                        children: parseInline(c.trim(), i === 0 ? { bold: true } : {})
                    }));
                    rows.push({ type: 'row', children: cells });
                }
                content.push({ type: 'table', metadata: tableAlign ? { align: tableAlign } : undefined, children: rows });
                continue;
            }
        }

        // Hr
        if (block.match(/^---+|^\*\*\*+|___+$/)) {
            content.push({ type: 'break', metadata: { breakType: 'page' } });
            continue;
        }

        // Paragraph
        content.push({
            type: 'paragraph',
            metadata: { alignment } as any,
            children: parseInline(block.replace(/\n/g, ' '))
        });
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

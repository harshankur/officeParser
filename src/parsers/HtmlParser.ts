import { AdmonitionMetadata, CellMetadata, CodeMetadata, EmbedMetadata, FullOfficeParserConfig, HeadingMetadata, ImageMetadata, ListMetadata, OfficeAttachment, OfficeContentNode, OfficeErrorType, OfficeMetadata, OfficeParserAST, ParagraphMetadata, TableMetadata, TextFormatting, TextMetadata } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { checkAbortSignal, getOfficeError } from '../utils/errorUtils.js';

interface HtmlNode {
    type: 'element' | 'text';
    tagName?: string;
    attributes?: Record<string, string>;
    text?: string;
    children: HtmlNode[];
    parent?: HtmlNode;
}

const parseAttributes = (attrString: string): Record<string, string> => {
    const attrs: Record<string, string> = {};
    const regex = /([a-zA-Z0-9\-:]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+)))?/g;
    let match;
    while ((match = regex.exec(attrString)) !== null) {
        const name = match[1].toLowerCase();
        const value = match[2] !== undefined ? match[2] : (match[3] !== undefined ? match[3] : (match[4] || ''));
        attrs[name] = value;
    }
    return attrs;
};

const parseHtmlTree = (html: string): HtmlNode => {
    const root: HtmlNode = { type: 'element', tagName: 'root', children: [], attributes: {} };
    let current = root;
    let cursor = 0;

    while (cursor < html.length) {
        const tagStart = html.indexOf('<', cursor);

        if (tagStart === -1) {
            const text = html.substring(cursor);
            if (text) current.children.push({ type: 'text', text, children: [], parent: current });
            break;
        }

        if (tagStart > cursor) {
            const text = html.substring(cursor, tagStart);
            if (text) current.children.push({ type: 'text', text, children: [], parent: current });
        }

        if (html.startsWith('<!--', tagStart)) {
            const commentEnd = html.indexOf('-->', tagStart + 4);
            cursor = commentEnd !== -1 ? commentEnd + 3 : html.length;
            continue;
        }

        // indexOf (not substring().match) so scanning for the tag end is O(1) in
        // allocation — a document with many "<" chars would otherwise be O(n^2).
        const tagEndIdx = html.indexOf('>', tagStart);
        if (tagEndIdx === -1) {
            const text = html.substring(tagStart);
            current.children.push({ type: 'text', text, children: [], parent: current });
            break;
        }

        const tagContent = html.substring(tagStart + 1, tagEndIdx);
        cursor = tagEndIdx + 1;

        const isClosing = tagContent.startsWith('/');
        const isSelfClosing = tagContent.endsWith('/');
        const tagCore = tagContent.replace(/^\/|\/$/g, '').trim();

        const firstSpace = tagCore.search(/\s/);
        const tagName = (firstSpace === -1 ? tagCore : tagCore.substring(0, firstSpace)).toLowerCase();
        const attrString = firstSpace === -1 ? '' : tagCore.substring(firstSpace);

        if (!tagName || !tagName.match(/^[a-z0-9\-]+$/)) {
            // Probably not a real tag, e.g., < 5
            current.children.push({ type: 'text', text: `<${tagContent}>`, children: [], parent: current });
            continue;
        }

        if (isClosing) {
            let p: HtmlNode | undefined = current;
            while (p && p.tagName !== tagName) {
                p = p.parent;
            }
            if (p && p.parent) {
                current = p.parent;
            }
        } else {
            const node: HtmlNode = {
                type: 'element',
                tagName,
                attributes: parseAttributes(attrString),
                children: [],
                parent: current
            };
            current.children.push(node);

            const voidElements = new Set(['area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'param', 'source', 'track', 'wbr', '!doctype']);
            if (!isSelfClosing && !voidElements.has(tagName)) {
                current = node;

                if (tagName === 'script' || tagName === 'style') {
                    // Case-insensitive search from `cursor` via a sticky-ish regex, instead of
                    // lower-casing the whole document on every <script>/<style> (was O(n^2)).
                    // tagName is validated to /^[a-z0-9-]+$/ above, so it's safe to interpolate.
                    const closeRe = new RegExp(`</${tagName}>`, 'gi');
                    closeRe.lastIndex = cursor;
                    const closeMatch = closeRe.exec(html);
                    if (closeMatch) {
                        node.children.push({
                            type: 'text',
                            text: html.substring(cursor, closeMatch.index),
                            children: [],
                            parent: node
                        });
                        cursor = closeMatch.index + closeMatch[0].length;
                        current = node.parent!;
                    }
                }
            }
        }
    }

    return root;
};

export const parseHtml = async (buffer: Buffer, config: FullOfficeParserConfig): Promise<OfficeParserAST> => {
    // Honour cancellation requests before the HTML tree is built and traversed.
    // The custom recursive HTML parser can be expensive for large documents;
    // rejecting early here prevents both the parsing and the subsequent AST construction.
    checkAbortSignal(config.abortSignal);

    const textStr = buffer.toString('utf-8');
    const root = parseHtmlTree(textStr);

    // Find head and body
    let head: HtmlNode | undefined;
    let body: HtmlNode = root;

    const findNode = (node: HtmlNode, tag: string): HtmlNode | undefined => {
        if (node.tagName === tag) return node;
        for (const child of node.children) {
            const found = findNode(child, tag);
            if (found) return found;
        }
        return undefined;
    };

    const htmlNode = findNode(root, 'html');
    if (htmlNode) {
        head = findNode(htmlNode, 'head');
        body = findNode(htmlNode, 'body') || htmlNode;
    }

    const metadata: OfficeMetadata = {};
    const attachments: OfficeAttachment[] = [];

    if (head) {
        const titleNode = findNode(head, 'title');
        if (titleNode && titleNode.children.length > 0 && titleNode.children[0].text) {
            metadata.title = titleNode.children[0].text;
        }

        metadata.nativeProperties = {};
        for (const child of head.children) {
            if (child.tagName === 'meta') {
                const name = child.attributes?.name || child.attributes?.property || child.attributes?.['http-equiv'];
                if (name) {
                    metadata.nativeProperties[name] = child.attributes?.content || '';
                }
            }
        }

        const extractMeta = (name: string): string | undefined => {
            for (const child of head!.children) {
                if (child.tagName === 'meta' && (child.attributes?.name === name || child.attributes?.property === name)) {
                    return child.attributes?.content;
                }
            }
            return undefined;
        };

        const author = extractMeta('author');
        if (author) metadata.author = author;
        const desc = extractMeta('description');
        if (desc) metadata.description = desc;

        const created = extractMeta('dcterms.created');
        if (created) metadata.created = new Date(created);
        const modified = extractMeta('dcterms.modified');
        if (modified) metadata.modified = new Date(modified);
        const lastMod = extractMeta('lastModifiedBy');
        if (lastMod) metadata.lastModifiedBy = lastMod;

        // Custom properties
        const customProps: Record<string, string | number | boolean | Date> = {};
        for (const child of head.children) {
            if (child.tagName === 'meta' && child.attributes?.name?.startsWith('custom:')) {
                const key = child.attributes.name.substring(7);
                const val = child.attributes.content || '';
                // Try to infer type
                if (val === 'true') customProps[key] = true;
                else if (val === 'false') customProps[key] = false;
                else if (!isNaN(Number(val)) && val.trim() !== '') customProps[key] = Number(val);
                else if (!isNaN(Date.parse(val)) && val.includes(':')) customProps[key] = new Date(val);
                else customProps[key] = val;
            }
        }
        if (Object.keys(customProps).length > 0) metadata.customProperties = customProps;
    }

    const content: OfficeContentNode[] = [];
    let htmlListIdCounter = 1;

    interface ListContext {
        listId: string;
        type: 'ordered' | 'unordered';
        level: number;
        counters: Record<number, number>;
        isTask?: boolean;
    }

    // Finds the checked state from a nested <input type="checkbox"> (GFM task-list items
    // nest it inside a <label>, so it isn't a direct child of the <li>).
    const findNestedCheckboxChecked = (n: HtmlNode): boolean | undefined => {
        if (n.tagName === 'input' && (n.attributes?.type || '').toLowerCase() === 'checkbox') {
            return 'checked' in (n.attributes || {});
        }
        for (const child of n.children) {
            const found = findNestedCheckboxChecked(child);
            if (found !== undefined) return found;
        }
        return undefined;
    };

    // Populated from a <section data-footnotes> block (found and parsed before the main
    // body loop, since references can appear anywhere earlier in the document) and
    // consulted by parseChildren's <sup data-footnote-ref> handling below.
    const footnoteDefinitions = new Map<string, OfficeContentNode[]>();

    const parseNode = (node: HtmlNode, currentFormatting: TextFormatting = {}, listContext?: ListContext, depth: number = 0): OfficeContentNode | OfficeContentNode[] | null => {
        // Guard against a maliciously deep element tree (e.g. tens of thousands of
        // nested <div>) recursing until the call stack overflows. Real documents
        // nest only a few dozen levels; this trips well before a RangeError.
        if (depth > 1000) {
            throw getOfficeError(OfficeErrorType.MAX_NESTING_DEPTH_EXCEEDED);
        }
        if (node.type === 'text') {
            let decodedText = (node.text || '')
                .replace(/&nbsp;/g, ' ')
                .replace(/&lt;/g, '<')
                .replace(/&gt;/g, '>')
                .replace(/&amp;/g, '&')
                .replace(/&quot;/g, '"')
                .replace(/&#39;/g, "'");

            if (!config.preserveXmlWhitespace) {
                decodedText = decodedText.replace(/\s+/g, ' ');
            }
            if (!decodedText.trim() && !config.preserveXmlWhitespace) return null;

            const textNode: OfficeContentNode = {
                type: 'text',
                text: decodedText,
                formatting: Object.keys(currentFormatting).length > 0 ? { ...currentFormatting } : undefined
            };

            if (config.includeRawContent && node.text) {
                // For text nodes in this manual parser, we just use the decoded text as raw content
                // as we don't have accurate locators for the original source slice
                textNode.rawContent = node.text;
            }

            return textNode;
        }

        if (node.type === 'element' && node.tagName) {
            const tagName = node.tagName;
            const newFormatting = { ...currentFormatting };

            if (tagName === 'b' || tagName === 'strong') newFormatting.bold = true;
            if (tagName === 'i' || tagName === 'em') newFormatting.italic = true;
            if (tagName === 'u') newFormatting.underline = true;
            if (tagName === 'strike' || tagName === 's' || tagName === 'del') newFormatting.strikethrough = true;
            if (tagName === 'sub') newFormatting.subscript = true;
            if (tagName === 'sup') newFormatting.superscript = true;
            if (tagName === 'code') newFormatting.font = 'monospace';

            const styleAttr = node.attributes?.style || '';
            const alignAttr = node.attributes?.align || '';
            if (styleAttr || alignAttr) {
                if (styleAttr.includes('font-weight: bold')) newFormatting.bold = true;
                if (styleAttr.includes('font-style: italic')) newFormatting.italic = true;
                if (styleAttr.includes('text-decoration: underline')) newFormatting.underline = true;
                if (styleAttr.includes('text-decoration: line-through')) newFormatting.strikethrough = true;

                const colorMatch = styleAttr.match(/color:\s*([^;]+)/);
                if (colorMatch) newFormatting.color = colorMatch[1].trim();

                const bgMatch = styleAttr.match(/background-color:\s*([^;]+)/);
                if (bgMatch) newFormatting.backgroundColor = bgMatch[1].trim();

                const sizeMatch = styleAttr.match(/font-size:\s*([^;]+)/);
                if (sizeMatch) newFormatting.size = sizeMatch[1].trim();

                const fontMatch = styleAttr.match(/font-family:\s*([^;]+)/);
                if (fontMatch) newFormatting.font = fontMatch[1].trim().split(',')[0].replace(/['"]/g, '');

                const alignmentMatch = styleAttr.match(/text-align:\s*(left|center|right|justify)/);
                if (alignmentMatch) {
                    newFormatting.alignment = alignmentMatch[1].toLowerCase() as any;
                } else if (alignAttr) {
                    const align = alignAttr.toLowerCase();
                    if (['left', 'center', 'right', 'justify'].includes(align)) {
                        newFormatting.alignment = align as any;
                    }
                }
            }

            const anchorIds = node.attributes?.id ? [node.attributes.id] : [];

            const parseChildren = (n: HtmlNode, fmt: TextFormatting, lCtx?: any): OfficeContentNode[] => {
                const kids: OfficeContentNode[] = [];
                for (const child of n.children) {
                    // Footnote/endnote reference: attach as .notes on the preceding node
                    // instead of inserting a visible node, matching WordParser's convention.
                    if (child.type === 'element' && child.tagName === 'sup' && child.attributes?.['data-footnote-ref'] !== undefined) {
                        const key = child.attributes['data-footnote-ref'];
                        const definition = footnoteDefinitions.get(key);
                        const noteNode: OfficeContentNode = {
                            type: 'note',
                            text: (definition || []).map(d => d.text || '').join(''),
                            children: definition || [],
                            metadata: { noteType: 'footnote', noteId: key }
                        };
                        if (kids.length > 0) {
                            const target = kids[kids.length - 1];
                            if (!target.notes) target.notes = [];
                            target.notes.push(noteNode);
                        } else {
                            kids.push({ type: 'text', text: '', notes: [noteNode] });
                        }
                        continue;
                    }

                    const parsed = parseNode(child, fmt, lCtx, depth + 1);
                    if (parsed) {
                        if (Array.isArray(parsed)) kids.push(...parsed);
                        else kids.push(parsed);
                    }
                }
                return kids;
            };

            // YouTube embeds: inscript-editor's Youtube node renders
            // <div data-youtube-video="ID" data-width="…" data-align="…">…<iframe…></div>.
            // Recognise both the wrapper div and a bare iframe so externally-authored HTML
            // (and a saved-then-reopened .md that fell back to raw HTML) both round-trip.
            if (tagName === 'div' && node.attributes?.['data-youtube-video'] !== undefined) {
                const videoId = node.attributes['data-youtube-video'] || '';
                const width = node.attributes?.['data-width'];
                const embedAlignAttr = node.attributes?.['data-align'];
                const embedAlign = (['left', 'center', 'right'] as const).includes(embedAlignAttr as any) ? embedAlignAttr as 'left' | 'center' | 'right' : undefined;
                const embedUrl = videoId ? `https://www.youtube.com/watch?v=${videoId}` : undefined;
                const embedNode: OfficeContentNode = {
                    type: 'embed',
                    // Childless nodes need .text so generic AST consumers (toText, chunking)
                    // don't silently drop them.
                    text: embedUrl,
                    metadata: {
                        embedType: 'youtube',
                        videoId,
                        url: embedUrl,
                        width,
                        align: embedAlign
                    } as EmbedMetadata
                };
                if (config.includeRawContent) embedNode.rawContent = '<div data-youtube-video>...</div>';
                return embedNode;
            }
            if (tagName === 'iframe') {
                const src = node.attributes?.src || '';
                const ytMatch = /youtube(?:-nocookie)?\.com/.test(src) ? src.match(/(?:embed\/|v=)([^&?/\s]+)/) : null;
                if (ytMatch) {
                    const embedUrl = `https://www.youtube.com/watch?v=${ytMatch[1]}`;
                    const embedNode: OfficeContentNode = {
                        type: 'embed',
                        text: embedUrl,
                        metadata: { embedType: 'youtube', videoId: ytMatch[1], url: embedUrl } as EmbedMetadata
                    };
                    if (config.includeRawContent) embedNode.rawContent = '<iframe>...</iframe>';
                    return embedNode;
                }
                return null;
            }

            // Footnotes section: its definitions were already extracted up front (see
            // footnoteDefinitions below), so skip it here wherever it appears in the tree -
            // it isn't necessarily a direct child of <body> (e.g. it may be nested inside
            // a non-standalone HtmlGenerator output's wrapping <div>).
            if (tagName === 'section' && node.attributes?.['data-footnotes'] !== undefined) {
                return null;
            }

            // Math: proposed contract (no editor node built yet) - HtmlGenerator emits
            // <span/div class="math math-inline|math-block" data-math="inline|block">
            // with the $-delimited LaTeX as the visible (escaped) text content.
            if ((tagName === 'div' || tagName === 'span') && node.attributes?.['data-math'] !== undefined) {
                const mathMode: 'inline' | 'block' = node.attributes['data-math'] === 'block' ? 'block' : 'inline';
                const rawText = node.children.map(c => c.text || '').join('')
                    .replace(/&nbsp;/g, ' ')
                    .replace(/&lt;/g, '<')
                    .replace(/&gt;/g, '>')
                    .replace(/&amp;/g, '&')
                    .replace(/&quot;/g, '"')
                    .replace(/&#39;/g, '\'');
                const delimiter = mathMode === 'block' ? '$$' : '$';
                const latex = rawText.startsWith(delimiter) && rawText.endsWith(delimiter)
                    ? rawText.slice(delimiter.length, -delimiter.length)
                    : rawText;
                return {
                    type: 'code',
                    text: latex,
                    metadata: { math: mathMode } as CodeMetadata
                };
            }

            // Admonition: inscript-editor's Admonition node renders
            // <div class="admonition admonition-note" data-type="note">…children…</div>.
            if (tagName === 'div' && (node.attributes?.class || '').split(/\s+/).includes('admonition')) {
                const admonitionTypeAttr = node.attributes?.['data-type'];
                const admonitionType = (['note', 'tip', 'important', 'warning', 'caution'] as const).includes(admonitionTypeAttr as any)
                    ? admonitionTypeAttr as AdmonitionMetadata['admonitionType']
                    : 'note';
                const admonitionNode: OfficeContentNode = {
                    type: 'admonition',
                    metadata: { admonitionType } as AdmonitionMetadata,
                    children: parseChildren(node, newFormatting, listContext)
                };
                if (config.includeRawContent) admonitionNode.rawContent = '<div class="admonition">...</div>';
                return admonitionNode;
            }

            // Skip structural containers produced by HtmlGenerator to avoid deep AST nesting
            if (tagName === 'div' && (
                node.attributes?.class === 'container' ||
                node.attributes?.class === 'spreadsheet-container' ||
                node.attributes?.class === 'presentation-container' ||
                node.attributes?.class === 'pdf-container' ||
                node.attributes?.class === 'metadata-summary' ||
                node.attributes?.class === 'image-container' ||
                node.attributes?.class === 'chart-container' ||
                node.attributes?.class === 'table-container' ||
                node.attributes?.class === 'caption' ||
                node.attributes?.class === 'sheet' ||
                node.attributes?.class === 'page' ||
                node.attributes?.class === 'slide' ||
                node.attributes?.class === 'note-content'
            )) {
                return parseChildren(node, newFormatting, listContext);
            }
            if (tagName === 'article') {
                return parseChildren(node, newFormatting, listContext);
            }

            if (tagName === 'p' || tagName === 'div') {
                const children = parseChildren(node, newFormatting, listContext);

                // If it's a div and contains block elements, return children directly
                const hasBlockElements = children.some(c => ['paragraph', 'table', 'heading', 'list', 'image', 'chart', 'code', 'embed', 'admonition', 'definitionList'].includes(c.type));
                if (tagName === 'div' && hasBlockElements) {
                    return children;
                }

                // Flatten nested paragraphs to avoid deep AST nesting (e.g. from notes)
                const flattenedChildren: OfficeContentNode[] = [];
                for (const child of children) {
                    if (child.type === 'paragraph' && child.children) {
                        flattenedChildren.push(...child.children);
                    } else {
                        flattenedChildren.push(child);
                    }
                }

                const pNode: OfficeContentNode = {
                    type: 'paragraph',
                    metadata: { alignment: newFormatting.alignment, anchorIds: anchorIds.length > 0 ? anchorIds : undefined } as ParagraphMetadata,
                    children: flattenedChildren
                };

                if (config.includeRawContent) {
                    // Note: Since this is a manual parser without locators, we can't easily get the original source slice.
                    // We'll skip rawContent for structural nodes here unless we want to implement index tracking in parseHtmlTree.
                }

                return pNode;
            }
            if (tagName.match(/^h[1-6]$/)) {
                const level = parseInt(tagName.substring(1));
                const hNode: OfficeContentNode = {
                    type: 'heading',
                    metadata: { level, alignment: newFormatting.alignment, anchorIds: anchorIds.length > 0 ? anchorIds : undefined } as HeadingMetadata,
                    children: parseChildren(node, newFormatting, listContext)
                };
                return hNode;
            }
            if (tagName === 'dl') {
                return {
                    type: 'definitionList',
                    children: parseChildren(node, newFormatting, listContext)
                };
            }
            if (tagName === 'dt') {
                return {
                    type: 'definitionTerm',
                    children: parseChildren(node, newFormatting, listContext)
                };
            }
            if (tagName === 'dd') {
                return {
                    type: 'definitionDescription',
                    children: parseChildren(node, newFormatting, listContext)
                };
            }
            if (tagName === 'abbr') {
                const title = node.attributes?.title;
                const children = parseChildren(node, newFormatting, listContext);
                if (title) {
                    children.forEach(c => {
                        if (c.type === 'text') {
                            c.metadata = { ...c.metadata, abbreviationTitle: title } as TextMetadata;
                        }
                    });
                }
                return children;
            }
            if (tagName === 'cite' && node.attributes?.['data-citation-key'] !== undefined) {
                const citationKey = node.attributes['data-citation-key'];
                return {
                    type: 'text',
                    text: citationKey,
                    formatting: Object.keys(newFormatting).length > 0 ? { ...newFormatting } : undefined,
                    metadata: { citationKey } as TextMetadata
                };
            }
            if (tagName === 'ul' || tagName === 'ol') {
                const isNewTopLevel = !listContext;
                const newListContext: ListContext = {
                    listId: isNewTopLevel ? `html-list-${htmlListIdCounter++}` : listContext!.listId,
                    type: tagName === 'ol' ? 'ordered' : 'unordered',
                    level: isNewTopLevel ? 0 : listContext!.level + 1,
                    counters: isNewTopLevel ? {} : { ...listContext!.counters }, // Clone to avoid side effects on parent levels
                    isTask: node.attributes?.['data-type'] === 'taskList'
                };

                // Initialize counter for this level
                if (tagName === 'ol' && node.attributes?.start) {
                    const start = parseInt(node.attributes.start, 10);
                    newListContext.counters[newListContext.level] = isNaN(start) ? 0 : start - 1;
                } else {
                    newListContext.counters[newListContext.level] = 0;
                }

                return parseChildren(node, currentFormatting, newListContext);
            }
            if (tagName === 'li') {
                if (listContext) {
                    if (node.attributes?.value) {
                        const val = parseInt(node.attributes.value, 10);
                        if (!isNaN(val)) listContext.counters[listContext.level] = val;
                    } else {
                        listContext.counters[listContext.level]++;
                    }
                }

                const children = parseChildren(node, newFormatting, listContext);
                const nestedLists = children.filter(c => c.type === 'list');
                const selfChildren = children.filter(c => c.type !== 'list');

                let isTask: boolean | undefined;
                let checked: boolean | undefined;
                if (listContext?.isTask) {
                    isTask = true;
                    const dataChecked = node.attributes?.['data-checked'];
                    checked = dataChecked !== undefined ? dataChecked === 'true' : (findNestedCheckboxChecked(node) ?? false);
                }

                const selfNode: OfficeContentNode = {
                    type: 'list',
                    text: selfChildren.map(c => c.text || '').join(''),
                    metadata: {
                        listType: listContext?.type || 'unordered',
                        indentation: listContext?.level || 0,
                        alignment: newFormatting.alignment || 'left',
                        listId: listContext?.listId || 'html-list-none',
                        itemIndex: (listContext?.counters[listContext.level] ?? 1) - 1,
                        anchorIds: anchorIds.length > 0 ? anchorIds : undefined,
                        isTask,
                        checked
                    } as ListMetadata,
                    children: selfChildren
                };

                return [selfNode, ...nestedLists];
            }
            if (tagName === 'table') {
                // CustomTable (inscript-editor) renders data-align on the <table> itself.
                const tableAlignAttr = node.attributes?.['data-align'];
                const tableAlign = (['left', 'center', 'right'] as const).includes(tableAlignAttr as any) ? tableAlignAttr as 'left' | 'center' | 'right' : undefined;

                const tableNode: OfficeContentNode = {
                    type: 'table',
                    metadata: { anchorIds: anchorIds.length > 0 ? anchorIds : undefined, align: tableAlign } as TableMetadata,
                    children: parseChildren(node, newFormatting, listContext)
                };
                if (config.includeRawContent) {
                    tableNode.rawContent = '<table>...</table>';
                }
                return tableNode;
            }
            if (tagName === 'tr') {
                const rowNode: OfficeContentNode = {
                    type: 'row',
                    children: parseChildren(node, newFormatting, listContext)
                };
                if (config.includeRawContent) {
                    rowNode.rawContent = '<tr>...</tr>';
                }
                return rowNode;
            }
            if (tagName === 'td' || tagName === 'th') {
                // Merged cells: mirrors the colspan/rowspan reading already done in
                // MarkdownParser's inline HTML-table handler.
                const colSpanAttr = node.attributes?.colspan;
                const rowSpanAttr = node.attributes?.rowspan;
                const colSpan = colSpanAttr ? parseInt(colSpanAttr, 10) : undefined;
                const rowSpan = rowSpanAttr ? parseInt(rowSpanAttr, 10) : undefined;

                const cellNode: OfficeContentNode = {
                    type: 'cell',
                    metadata: {
                        colSpan: colSpan && !isNaN(colSpan) ? colSpan : undefined,
                        rowSpan: rowSpan && !isNaN(rowSpan) ? rowSpan : undefined
                    } as CellMetadata,
                    children: parseChildren(node, newFormatting, listContext)
                };
                if (config.includeRawContent) {
                    cellNode.rawContent = '<td>...</td>';
                }
                return cellNode;
            }
            if (tagName === 'img') {
                const src = node.attributes?.src;
                const alt = node.attributes?.alt;

                // CustomImage (inscript-editor) renders data-width/data-align, falling back to
                // parsing the inline style for consumers that only emit the CSS.
                const imgStyle = node.attributes?.style || '';
                const width = node.attributes?.['data-width'] || imgStyle.match(/width:\s*([^;]+)/)?.[1]?.trim();
                const alignAttr = node.attributes?.['data-align']
                    || (imgStyle.includes('margin-left: 0') && !imgStyle.includes('margin-right: 0') ? 'left'
                        : (imgStyle.includes('margin-right: 0') && !imgStyle.includes('margin-left: 0') ? 'right' : undefined));
                const align = (['left', 'center', 'right'] as const).includes(alignAttr as any) ? alignAttr as 'left' | 'center' | 'right' : undefined;

                let imageNode: OfficeContentNode;
                if (src?.startsWith('data:')) {
                    const match = src.match(/^data:([^;]+);base64,(.*)$/);
                    if (match && config.extractAttachments) {
                        const mimeType = match[1] as any;
                        const data = match[2];
                        const name = `image_${attachments.length + 1}.${mimeType.split('/')[1]}`;
                        attachments.push({
                            type: 'image',
                            mimeType,
                            data,
                            name,
                            extension: mimeType.split('/')[1]
                        });
                        imageNode = {
                            type: 'image',
                            metadata: {
                                attachmentName: name,
                                altText: alt,
                                width,
                                align
                            } as ImageMetadata
                        };
                    } else {
                        imageNode = {
                            type: 'image',
                            metadata: {
                                url: src,
                                altText: alt,
                                width,
                                align
                            } as ImageMetadata
                        };
                    }
                } else {
                    imageNode = {
                        type: 'image',
                        metadata: {
                            url: src,
                            altText: alt,
                            anchorIds: anchorIds.length > 0 ? anchorIds : undefined,
                            width,
                            align
                        } as ImageMetadata
                    };
                }

                if (config.includeRawContent) {
                    imageNode.rawContent = '<img>';
                }
                return imageNode;
            }
            if (tagName === 'a') {
                const href = node.attributes?.href;
                const wikilinkPage = node.attributes?.['data-wikilink-page'];
                const children = parseChildren(node, newFormatting, listContext);
                if (wikilinkPage !== undefined) {
                    children.forEach(c => {
                        if (c.type === 'text') {
                            c.metadata = { ...c.metadata, link: wikilinkPage, linkType: 'internal', wikilink: true } as TextMetadata;
                        }
                    });
                } else if (href) {
                    const linkType = href.startsWith('#') ? 'internal' : 'external';
                    children.forEach(c => {
                        if (c.type === 'text') {
                            c.metadata = { ...c.metadata, link: href, linkType } as TextMetadata;
                        }
                    });
                }
                return children;
            }
            if (tagName === 'br') {
                const brNode: OfficeContentNode = { type: 'break', metadata: { breakType: 'textWrapping' } };
                if (config.includeRawContent) {
                    brNode.rawContent = '<br/>';
                }
                return brNode;
            }
            if (tagName === 'pre') {
                const codeNode = node.children.find(c => c.tagName === 'code');
                let language;
                let codeText = '';
                if (codeNode) {
                    const classAttr = codeNode.attributes?.class || '';
                    const langMatch = classAttr.split(' ').find((c: string) => c.startsWith('language-'));
                    if (langMatch) language = langMatch.replace('language-', '');
                    codeText = codeNode.children.map(c => c.text || '').join('');
                } else {
                    codeText = node.children.map(c => c.text || '').join('');
                }

                const preNode: OfficeContentNode = {
                    type: 'code',
                    text: codeText,
                    metadata: { language, anchorIds: anchorIds.length > 0 ? anchorIds : undefined } as CodeMetadata
                };
                if (config.includeRawContent) {
                    preNode.rawContent = '<pre>...</pre>';
                }
                return preNode;
            }

            if (tagName === 'script' || tagName === 'style' || tagName === '!doctype') {
                return null;
            }

            return parseChildren(node, newFormatting, listContext);
        }

        return null;
    };

    // Extract <section data-footnotes> up front so its definitions are available to
    // <sup data-footnote-ref> references encountered anywhere earlier in the body.
    const findFootnotesSection = (n: HtmlNode): HtmlNode | undefined => {
        if (n.tagName === 'section' && n.attributes?.['data-footnotes'] !== undefined) return n;
        for (const child of n.children) {
            const found = findFootnotesSection(child);
            if (found) return found;
        }
        return undefined;
    };
    const footnotesSectionNode = findFootnotesSection(body);
    if (footnotesSectionNode) {
        for (const item of footnotesSectionNode.children) {
            if (item.type !== 'element') continue;
            const key = item.attributes?.['data-footnote-id'];
            if (!key) continue;

            // Strip the generated back-reference link ("↩") - it's round-trip plumbing,
            // not part of the footnote's actual content.
            const filteredChildren = item.children.filter(c =>
                !(c.tagName === 'a' && (c.attributes?.href || '').startsWith('#footnote-ref-'))
            );
            const contentNodes: OfficeContentNode[] = [];
            for (const child of filteredChildren) {
                const parsed = parseNode(child);
                if (parsed) {
                    if (Array.isArray(parsed)) contentNodes.push(...parsed);
                    else contentNodes.push(parsed);
                }
            }
            footnoteDefinitions.set(key, contentNodes);
        }
    }
    for (const child of body.children) {
        const parsed = parseNode(child);
        if (parsed) {
            if (Array.isArray(parsed)) {
                parsed.forEach(p => {
                    if (p.type === 'text') {
                        // Wrap direct body text in paragraphs
                        content.push({ type: 'paragraph', children: [p] });
                    } else {
                        content.push(p);
                    }
                });
            } else {
                if (parsed.type === 'text') {
                    content.push({ type: 'paragraph', children: [parsed] });
                } else {
                    content.push(parsed);
                }
            }
        }
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

    return createAST('html', metadata, content, attachments, config, undefined, toTextSync);
};

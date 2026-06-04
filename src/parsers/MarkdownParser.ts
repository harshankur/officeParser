import { CodeMetadata, FullOfficeParserConfig, HeadingMetadata, ImageMetadata, ListMetadata, OfficeAttachment, OfficeContentNode, OfficeMetadata, OfficeParserAST, TextFormatting, TextMetadata } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { checkAbortSignal } from '../utils/errorUtils.js';

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
                    let val = match[2].trim().replace(/^"(.*)"$/, '$1');

                    let parsedVal: any = val;
                    if (val === 'true') parsedVal = true;
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

    const parseInline = (text: string, currentFormatting: TextFormatting = {}): OfficeContentNode[] => {
        const nodes: OfficeContentNode[] = [];

        // Regex matches: 1=!, 2=alt, 3=url | 4=bold | 5=italic | 6=strike | 7=code | 8=underline | 9=subscript | 10=superscript
        const regex = /(!?)\[(.*?)\]\((.*?)\)|\*\*(.+?)\*\*|\*(.+?)\*|~~(.+?)~~|`(.+?)`|<u>(.+?)<\/u>|<sub>(.+?)<\/sub>|<sup>(.+?)<\/sup>/g;
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
                            nodes.push({ type: 'image', metadata: { attachmentName: name, altText } as ImageMetadata });
                        } else {
                            nodes.push({ type: 'image', metadata: { url, altText } as ImageMetadata });
                        }
                    } else {
                        nodes.push({ type: 'image', metadata: { url, altText } as ImageMetadata });
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
            } else if (match[4]) { // Bold
                nodes.push(...parseInline(match[4], { ...currentFormatting, bold: true }));
            } else if (match[5]) { // Italic
                nodes.push(...parseInline(match[5], { ...currentFormatting, italic: true }));
            } else if (match[6]) { // Strikethrough
                nodes.push(...parseInline(match[6], { ...currentFormatting, strikethrough: true }));
            } else if (match[7]) { // Inline Code
                nodes.push({ type: 'text', text: match[7], formatting: { ...currentFormatting, font: 'monospace' } });
            } else if (match[8]) { // Underline
                nodes.push(...parseInline(match[8], { ...currentFormatting, underline: true }));
            } else if (match[9]) { // Subscript
                nodes.push(...parseInline(match[9], { ...currentFormatting, subscript: true }));
            } else if (match[10]) { // Superscript
                nodes.push(...parseInline(match[10], { ...currentFormatting, superscript: true }));
            }

            lastIndex = regex.lastIndex;
        }

        if (lastIndex < text.length) {
            nodes.push({ type: 'text', text: text.substring(lastIndex), formatting: Object.keys(currentFormatting).length > 0 ? { ...currentFormatting } : undefined });
        }

        return nodes;
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
            content.push({
                type: 'paragraph',
                metadata: { style: 'Quote' } as any,
                children: parseInline(quoteMatch[1].replace(/^>\s+/gm, ''))
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

                    const children = parseInline(match[3]);
                    content.push({
                        type: 'list',
                        text: children.map(c => c.text || '').join(''),
                        metadata: {
                            listType,
                            indentation: level,
                            alignment: alignment || 'left',
                            listId,
                            itemIndex: listCounters[level]
                        } as ListMetadata,
                        children
                    });
                }
            }
            continue;
        }

        // Table (Simple Pipe or HTML)
        if ((block.includes('|') && block.match(/\n\s*\|?[-:| ]+\|?\s*\n/)) || block.includes('<table')) {
            if (block.includes('<table')) {
                // Basic HTML table recognition (extracting rows/cells)
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
                if (rows.length > 0) {
                    content.push({ type: 'table', children: rows });
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
                content.push({ type: 'table', children: rows });
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
            if (node.children) {
                const isBlock = ['table', 'row', 'list', 'sheet', 'slide'].includes(node.type);
                return node.children.map(getText).join(isBlock ? config.newlineDelimiter : '');
            }
            return '';
        };
        return getText(n);
    }).join(config.newlineDelimiter)
        .replace(/\n{3,}/g, '\n\n'); // Normalize excessive whitespace

    return createAST('md', metadata, content, attachments, config, undefined, toTextSync);
};

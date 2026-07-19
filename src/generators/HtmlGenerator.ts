import { AdmonitionMetadata, CellMetadata, CodeMetadata, ConversionResult, EmbedMetadata, GeneratorConfig, HeadingMetadata, ImageMetadata, ListMetadata, NoteMetadata, OfficeContentNode, OfficeParserAST, PageMetadata, SlideMetadata, StandaloneConfig, TableMetadata, TextMetadata } from '../types.js';
import { BaseGenerator } from './BaseGenerator.js';
import { escapeHtml, sanitizeCssValue, sanitizeUrl, sanitizeImageUrl, serializeForInlineScript } from '../utils/sanitize.js';

type ResolvedStandalone = Required<StandaloneConfig>;

/**
 * Attributes that carry a URL and therefore must go through `sanitizeUrl` rather than plain
 * escaping - an escaped `javascript:` payload is still a `javascript:` payload.
 */
const URL_BEARING_ATTRS = new Set([
    'href', 'src', 'srcset', 'action', 'formaction', 'poster', 'cite', 'data', 'background', 'ping',
]);

/**
 * Renders `node.htmlAttributes` (see `BaseContentNode.htmlAttributes`) as an attribute string.
 *
 * This re-applies the parser's filtering rather than trusting it, because an AST can be built
 * programmatically and handed straight to the generator - the parse-side pass is defence in depth,
 * not the only gate. `class` is returned separately so the caller can merge it into the class
 * attribute it already composes: emitting a second `class=` would be invalid HTML and, worse, a
 * *fatal* XML well-formedness error once EpubGenerator converts the output to XHTML.
 */
function renderHtmlAttributeBag(
    node: OfficeContentNode,
    alreadyEmitted: Iterable<string> = []
): { attrs: string; className?: string } {
    const bag = node.htmlAttributes;
    if (!bag) return { attrs: '' };

    const taken = new Set([...alreadyEmitted].map(k => k.toLowerCase()));
    let attrs = '';
    let className: string | undefined;

    for (const [rawKey, rawValue] of Object.entries(bag)) {
        const key = rawKey.toLowerCase();
        // Same policy as the parser, restated here because this path is independently reachable.
        if (/^on/i.test(key)) continue;
        if (key === 'srcdoc' || key === 'style' || key === 'id') continue;
        if (!/^[a-zA-Z][a-zA-Z0-9-]*$/.test(key)) continue;
        if (taken.has(key)) continue;

        if (key === 'class') {
            className = String(rawValue);
            continue;
        }
        if (URL_BEARING_ATTRS.has(key)) {
            const safe = sanitizeUrl(String(rawValue));
            if (!safe) continue;
            attrs += ` ${key}="${safe}"`;
            continue;
        }
        attrs += ` ${key}="${escapeHtml(String(rawValue))}"`;
    }
    return { attrs, className };
}

/**
 * Normalizes `HtmlGeneratorConfig.standalone` (`boolean | StandaloneConfig`) into a fully
 * resolved object. `true`/undefined turns every part on (a complete standalone document);
 * `false` turns every part off (a bare content fragment). When an object is passed, any field
 * left unspecified defaults to its "on" value, matching the boolean-shorthand semantics.
 */
function resolveStandalone(standalone: boolean | StandaloneConfig | undefined): ResolvedStandalone {
    const uniform = (on: boolean): ResolvedStandalone => ({
        document: on,
        metaTags: on,
        styles: on ? 'full' : 'none',
        scripts: on,
        headInjections: on,
        bodyInjections: on,
    });

    if (standalone === undefined || typeof standalone === 'boolean') {
        return uniform(standalone ?? true);
    }

    const on = uniform(true);
    return {
        document: standalone.document ?? on.document,
        metaTags: standalone.metaTags ?? on.metaTags,
        styles: standalone.styles ?? on.styles,
        scripts: standalone.scripts ?? on.scripts,
        headInjections: standalone.headInjections ?? on.headInjections,
        bodyInjections: standalone.bodyInjections ?? on.bodyInjections,
    };
}

/**
 * Generates semantic, high-fidelity HTML from an AST.
 */
export class HtmlGenerator extends BaseGenerator<'html'> {
    private chartCounter = 0;
    private isSpreadsheetMode = false;

    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'html'>) {
        super('html', ast, config);
    }

    /**
     * Generates HTML string from the provided AST.
     * 
     * @returns An HTML string
     */
    async generate(): Promise<ConversionResult<'html'>> {
        this.isSpreadsheetMode = this.ast.content.some(n => n.type === 'sheet');
        const isPresentation = this.ast.content.some(n => n.type === 'slide');
        const isPdf = this.ast.content.some(n => n.type === 'page');

        let containerClass = 'container';
        if (this.isSpreadsheetMode) containerClass = 'spreadsheet-container';
        else if (isPresentation) containerClass = 'presentation-container';
        else if (isPdf) containerClass = 'pdf-container';

        let bodyContent = await this.processNodeArray(this.ast.content);

        if (this.collectedNotes.length > 0) {
            // Footnotes/endnotes get their own <section data-footnotes> (the agreed
            // contract with inscript-editor's footnote node); other note types (e.g.
            // slide speaker notes) keep the existing generic notes wrapper.
            const footnotes = this.collectedNotes.filter(n => {
                const t = (n.metadata as any)?.noteType;
                return t === 'footnote' || t === 'endnote';
            });
            const otherNotes = this.collectedNotes.filter(n => !footnotes.includes(n));

            if (footnotes.length > 0) {
                let footnotesHtml = '';
                for (const note of footnotes) {
                    footnotesHtml += await this.processNodeRecursive(note, this.nodeProcessor.bind(this));
                }
                // data-footnotes carries an explicit empty value (not a bare attribute) so
                // the markup is valid XHTML too - EpubGenerator embeds this verbatim, and
                // XML rejects valueless attributes. HtmlParser only checks for presence.
                bodyContent += `\n<section data-footnotes="">\n${footnotesHtml}\n</section>\n`;
            }

            if (otherNotes.length > 0) {
                let notesHtml = '';
                for (const note of otherNotes) {
                    notesHtml += await this.processNodeRecursive(note, this.nodeProcessor.bind(this));
                }
                bodyContent += `\n<div class="document-notes-section">\n<hr class="page-break">\n${notesHtml}\n</div>\n`;
            }
        }

        const metadataBlock = this.config.renderMetadata ? this.renderMetadataSummary() : '';

        let title = 'Document';
        let metaTags = '';
        let spreadsheetTabs = '';
        let spreadsheetScript = '';

        if (this.isSpreadsheetMode) {
            const sheets: OfficeContentNode[] = [];
            for (const node of this.ast.content) {
                if (node.type === 'sheet') {
                    const override = await this.handleOnNode(node);
                    if (override !== false) {
                        sheets.push(node);
                    }
                }
            }
            const tabs = sheets.map((n, i) => {
                const sheetName = (n.metadata as any)?.sheetName || `Sheet ${i + 1}`;
                return `<a href="#sheet-${i}" class="spreadsheet-tab">${this.escape(sheetName)}</a>`;
            }).join('');
            spreadsheetTabs = `<div class="spreadsheet-tabs">${tabs}</div>`;
            spreadsheetScript = `
<script>
    function initSpreadsheetResizing() {
        document.querySelectorAll('.excel-grid').forEach(table => {
            if (table.dataset.resizingInitialized) return;
            if (table.offsetWidth === 0) return; // skip hidden tables
            table.dataset.resizingInitialized = 'true';
            
            const sheetId = table.parentElement.id || 'sheet';
            const docId = window.location.pathname;

            // Freeze initial auto-layout widths of columns and set table layout to fixed
            const colHeaders = table.querySelectorAll('.excel-col-header');
            colHeaders.forEach((header, index) => {
                let currentWidth = header.offsetWidth;
                const savedWidth = localStorage.getItem(docId + '_' + sheetId + '_col_' + index);
                if (savedWidth) {
                    currentWidth = parseInt(savedWidth, 10);
                }
                header.style.width = currentWidth + 'px';
                header.style.minWidth = currentWidth + 'px';
            });
            table.style.width = table.offsetWidth + 'px';
            table.style.tableLayout = 'fixed';

            // Freeze initial auto-layout heights of rows
            table.querySelectorAll('tr').forEach((row, index) => {
                let currentHeight = row.offsetHeight;
                const savedHeight = localStorage.getItem(docId + '_' + sheetId + '_row_' + index);
                if (savedHeight) {
                    currentHeight = parseInt(savedHeight, 10);
                }
                row.style.height = currentHeight + 'px';
            });

            colHeaders.forEach((header, index) => {
                if (header.querySelector('.col-resizer')) return;
                const resizer = document.createElement('div');
                resizer.className = 'col-resizer';
                header.appendChild(resizer);

                let startX = 0;
                let startWidth = 0;
                let startTableWidth = 0;

                const onMouseMove = (e) => {
                    const width = startWidth + (e.clientX - startX);
                    if (width > 40) {
                        header.style.width = width + 'px';
                        header.style.minWidth = width + 'px';
                        table.style.width = (startTableWidth + (width - startWidth)) + 'px';
                    }
                };

                const onMouseUp = (e) => {
                    resizer.classList.remove('resizing');
                    document.removeEventListener('mousemove', onMouseMove);
                    document.removeEventListener('mouseup', onMouseUp);
                    const finalWidth = startWidth + (e.clientX - startX);
                    if (finalWidth > 40) {
                        localStorage.setItem(docId + '_' + sheetId + '_col_' + index, finalWidth);
                    }
                };

                resizer.addEventListener('mousedown', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    startX = e.clientX;
                    startWidth = header.offsetWidth;
                    startTableWidth = table.offsetWidth;
                    resizer.classList.add('resizing');
                    document.addEventListener('mousemove', onMouseMove);
                    document.addEventListener('mouseup', onMouseUp);
                });
            });

            table.querySelectorAll('.excel-row-num').forEach((rowHeader, index) => {
                if (rowHeader.querySelector('.row-resizer')) return;
                const resizer = document.createElement('div');
                resizer.className = 'row-resizer';
                rowHeader.appendChild(resizer);

                const row = rowHeader.parentElement;
                let startY = 0;
                let startHeight = 0;

                const onMouseMove = (e) => {
                    const height = startHeight + (e.clientY - startY);
                    if (height > 20) {
                        row.style.height = height + 'px';
                    }
                };

                const onMouseUp = (e) => {
                    resizer.classList.remove('resizing');
                    document.removeEventListener('mousemove', onMouseMove);
                    document.removeEventListener('mouseup', onMouseUp);
                    const finalHeight = startHeight + (e.clientY - startY);
                    if (finalHeight > 20) {
                        localStorage.setItem(docId + '_' + sheetId + '_row_' + index, finalHeight);
                    }
                };

                resizer.addEventListener('mousedown', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    startY = e.clientY;
                    startHeight = row.offsetHeight;
                    resizer.classList.add('resizing');
                    document.addEventListener('mousemove', onMouseMove);
                    document.addEventListener('mouseup', onMouseUp);
                });
            });
        });
    }

    function switchSheet() {
        try {
            let hash = window.location.hash;
            if (!hash || hash === '#' || !hash.startsWith('#sheet-')) hash = '#sheet-0';
            
            const sheets = document.querySelectorAll('.spreadsheet-sheet');
            const tabs = document.querySelectorAll('.spreadsheet-tab');
            
            if (sheets.length === 0) return;

            sheets.forEach(s => s.classList.remove('active'));
            tabs.forEach(t => t.classList.remove('active'));
            
            const activeSheet = document.querySelector(hash) || sheets[0];
            activeSheet.classList.add('active');
            
            const activeTab = document.querySelector('a[href="' + hash + '"]') || tabs[0];
            if (activeTab) activeTab.classList.add('active');

            // Trigger chart re-render/resize when sheet becomes visible
            window.dispatchEvent(new Event('resize'));
            if (window.Chart) {
                Object.values(window.Chart.instances || {}).forEach(chart => {
                    if (activeSheet.contains(chart.canvas)) {
                        chart.resize();
                        chart.update();
                    }
                });
            }

            initSpreadsheetResizing();
        } catch (e) {
            console.error('Sheet switch failed:', e);
            const firstSheet = document.querySelector('.spreadsheet-sheet');
            if (firstSheet) firstSheet.classList.add('active');
        }
    }
    
    window.addEventListener('hashchange', switchSheet);
    if (document.readyState === 'complete') switchSheet();
    else window.addEventListener('load', switchSheet);
</script>`;
        }

        const sa = resolveStandalone(this.config.htmlConfig.standalone);

        if (sa.document && sa.metaTags) {
            title = this.effectiveMetadata.title || 'Document';
            metaTags = this.renderMetaTags();
        }

        const styleBlock = sa.styles === 'none' ? '' : `<style>${
            sa.styles === 'scoped'
                ? this.getScopedPremiumStyles(this.isSpreadsheetMode, isPresentation, isPdf)
                : this.getPremiumStyles(this.isSpreadsheetMode, isPresentation, isPdf)
        }</style>`;

        // 'scoped' styles are anchored to this wrapper via CSS @scope, so custom properties and
        // base body-level styling attach here instead of leaking onto a host page's real <html>/
        // <body> when the output is embedded as a fragment.
        const scopeOpen = sa.styles === 'scoped' ? '<div class="op-html-scope">' : '';
        const scopeClose = sa.styles === 'scoped' ? '</div>' : '';

        const inj = this.config.htmlConfig.injections;
        const headInjectionsOn = sa.document && sa.headInjections;
        const chartScriptTag = (sa.scripts && this.config.includeCharts) ? `<script src="${this.config.htmlConfig.chartJsSrc}"></script>` : '';
        const spreadsheetScriptOut = sa.scripts ? spreadsheetScript : '';

        const value = sa.document ? `<!DOCTYPE html>
<html lang="en">
<head>
    ${headInjectionsOn ? inj.headStart : ''}
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${this.escape(title)}</title>
    ${metaTags}
    ${chartScriptTag}
    ${styleBlock}
    ${headInjectionsOn ? inj.headEnd : ''}
</head>
<body>
    ${sa.bodyInjections ? inj.bodyStart : ''}
    ${scopeOpen}
    <div class="${containerClass}">
        <article>
            ${metadataBlock}
            ${bodyContent}
        </article>
        ${spreadsheetTabs}
    </div>
    ${scopeClose}
    ${spreadsheetScriptOut}
    ${sa.bodyInjections ? inj.bodyEnd : ''}
</body>
</html>` : `${styleBlock}${sa.bodyInjections ? inj.bodyStart : ''}${scopeOpen}<div class="${containerClass}">${metadataBlock}${bodyContent}${spreadsheetTabs}</div>${scopeClose}${spreadsheetScriptOut}${sa.bodyInjections ? inj.bodyEnd : ''}`;

        return {
            value,
            messages: this.messages
        };
    }

    private renderMetaTags(): string {
        if (!this.ast?.metadata) return '';
        const m = this.effectiveMetadata;
        const tags: string[] = [];
        if (m.author) tags.push(`<meta name="author" content="${this.escape(m.author)}">`);
        if (m.description) tags.push(`<meta name="description" content="${this.escape(m.description)}">`);
        const created = this.toIsoDate(m.created);
        const modified = this.toIsoDate(m.modified);
        if (created) tags.push(`<meta name="dcterms.created" content="${created}">`);
        if (modified) tags.push(`<meta name="dcterms.modified" content="${modified}">`);
        if (m.lastModifiedBy) tags.push(`<meta name="lastModifiedBy" content="${this.escape(m.lastModifiedBy)}">`);

        if (m.customProperties) {
            for (const [key, val] of Object.entries(m.customProperties)) {
                tags.push(`<meta name="custom:${this.escape(key)}" content="${this.escape(String(val))}">`);
            }
        }
        return tags.join('\n    ');
    }

    private renderMetadataSummary(): string {
        if (!this.ast?.metadata) return '';
        const m = this.effectiveMetadata;

        let customPropsHtml = '';
        if (m.customProperties && Object.keys(m.customProperties).length > 0) {
            const items = Object.entries(m.customProperties)
                .map(([k, v]) => `<div class="meta-tag"><strong>${this.escape(k)}:</strong> ${this.escape(String(v))}</div>`)
                .join('');
            customPropsHtml = `<div class="meta-custom-section">
                <div class="meta-section-title">🏷️ Custom Properties</div>
                <div class="meta-tags-grid">${items}</div>
            </div>`;
        }

        const addField = (label: string, val?: string | Date) => {
            if (!val) return '';
            const display = val instanceof Date ? val.toLocaleString() : val;
            return `<div class="meta-item">
                <div class="meta-label">${label}</div>
                <div class="meta-value">${this.escape(display)}</div>
            </div>`;
        };

        return `<div class="metadata-summary">
            <div class="meta-grid">
                ${addField('Title', m.title)}
                ${addField('Author', m.author)}
                ${addField('Created', m.created)}
                ${addField('Modified', m.modified)}
            </div>
            ${customPropsHtml}
        </div>`;
    }

    /**
     * Processes an array of nodes, handling list grouping and nesting.
     */
    private async processNodeArray(nodes: OfficeContentNode[]): Promise<string> {
        let html = '';
        // Stack to track active lists: { indentation, type, isTask }
        const listStack: { indentation: number, type: 'ordered' | 'unordered', isTask: boolean }[] = [];

        const openListTag = (type: 'ordered' | 'unordered', isTask: boolean) => {
            if (isTask) return '<ul data-type="taskList">';
            return type === 'ordered' ? '<ol>' : '<ul>';
        };
        const closeListTag = (type: 'ordered' | 'unordered') => type === 'ordered' ? '</ol>' : '</ul>';

        const closeListsToLevel = (level: number) => {
            while (listStack.length > 0 && listStack[listStack.length - 1].indentation > level) {
                const list = listStack.pop();
                html += closeListTag(list!.type) + '\n\n';
            }
        };

        for (const node of nodes) {
            // Check if node should be filtered out or overridden
            const override = await this.handleOnNode(node);
            if (override === false) {
                continue;
            }

            if (node.type === 'list') {
                const meta = node.metadata as ListMetadata;
                const type = meta?.listType === 'ordered' ? 'ordered' : 'unordered';
                const isTask = !!meta?.isTask;
                const indentation = meta?.indentation || 0;

                // Close deeper lists
                closeListsToLevel(indentation);

                // Handle current level
                if (listStack.length > 0 && listStack[listStack.length - 1].indentation === indentation) {
                    if (listStack[listStack.length - 1].type !== type || listStack[listStack.length - 1].isTask !== isTask) {
                        // Type changed at same level
                        const last = listStack.pop();
                        html += closeListTag(last!.type) + '\n';
                        html += openListTag(type, isTask) + '\n';
                        listStack.push({ indentation, type, isTask });
                    }
                } else {
                    // Start a new nested list
                    html += openListTag(type, isTask) + '\n';
                    listStack.push({ indentation, type, isTask });
                }

                html += await this.processNodeRecursive(node, this.nodeProcessor.bind(this), override);
            } else {
                // Non-list node closes all active lists
                closeListsToLevel(-1);
                let result = await this.processNodeRecursive(node, this.nodeProcessor.bind(this), override);

                // Add newlines for readability of the HTML source
                if (!result.endsWith('\n\n')) {
                    if (result.endsWith('\n')) result += '\n';
                    else result += '\n\n';
                }
                html += result;
            }
        }

        closeListsToLevel(-1);
        return html;
    }

    /**
     * Overridden to handle children using processNodeArray for list grouping.
     */
    private tableNestingLevel = 0;

    protected override async processNodeRecursive(
        node: OfficeContentNode,
        processor: (node: OfficeContentNode, childrenOutput: string) => string | Promise<string>,
        override?: string | boolean | void
    ): Promise<string> {
        // Use pre-evaluated override if provided, otherwise call handleOnNode
        const actualOverride = override !== undefined ? override : await this.handleOnNode(node);

        // Returning false skips the node and its children
        if (actualOverride === false) {
            return '';
        }

        // Returning a string overrides default rendering and recursion
        if (typeof actualOverride === 'string') {
            return actualOverride;
        }

        const isTable = node.type === 'table' || node.type === 'sheet';
        if (isTable) this.tableNestingLevel++;

        let childrenOutput = '';
        if (node.children && node.children.length > 0) {
            childrenOutput = await this.processNodeArray(node.children);
        } else if (node.text && node.type !== 'text') {
            // Fallback for nodes that have text property but no children (e.g. simple paragraphs)
            childrenOutput = this.escape(node.text);
        }

        if (node.notes && node.notes.length > 0) {
            if (node.type !== 'slide') {
                this.collectedNotes.push(...node.notes);
            }
        }

        let result = await processor(node, childrenOutput);

        if (isTable) this.tableNestingLevel--;

        if (node.type === 'slide' && node.notes && node.notes.length > 0) {
            for (const note of node.notes) {
                result += await this.processNodeRecursive(note, processor);
            }
        } else if (node.notes && node.notes.length > 0) {
            // Emit the reference marker at the point of citation. Without this, a
            // footnote/endnote would only ever show up in the collected footnotes
            // section, with no indication of where it was originally cited.
            for (const note of node.notes) {
                const meta = note.metadata as any;
                if (meta?.noteType === 'footnote' || meta?.noteType === 'endnote') {
                    const key = this.escape(this.getFootnoteKey(note));
                    result += `<sup data-footnote-ref="${key}" id="footnote-ref-${key}"><a href="#footnote-${key}">${key}</a></sup>`;
                }
            }
        }

        return result;
    }

    /**
     * Internal processor for individual nodes.
     */
    private async nodeProcessor(node: OfficeContentNode, childrenOutput: string): Promise<string> {
        // Handle Style Mapping using the semantic mapping helper
        const mapping = this.getSemanticMapping(node);

        let tag = mapping?.tag || this.getDefaultTag(node);

        // Handle Attributes from mapping
        let mappedAttrs = '';
        if (mapping?.attributes) {
            for (const [key, val] of Object.entries(mapping.attributes)) {
                mappedAttrs += ` ${key}="${this.escape(val)}"`;
            }
        }

        // Preserved source attributes (opt-in; absent unless htmlParserConfig.preserveAttributes).
        // Dedupe against whatever the mapping already emitted so a typed field always wins, and
        // take the bag's `class` back as a value to merge below rather than a second attribute.
        const bag = renderHtmlAttributeBag(node, Object.keys(mapping?.attributes || {}));
        // Fold into mappedAttrs rather than threading a separate fragment through all ~20 emission
        // sites: every site already interpolates mappedAttrs, so this cannot miss one (a miss would
        // silently drop preserved attributes), and an empty bag contributes nothing, so output for
        // nodes without one stays byte-identical.
        mappedAttrs += bag.attrs;

        // Combine classes from mapping, defaults, and any preserved source class. Merging here is
        // what keeps `<p class="lead">` from either losing "lead" or emitting a duplicate `class`.
        const classes = mapping?.classes ? [...mapping.classes] : [];
        if (bag.className) {
            for (const c of bag.className.split(/\s+/).filter(Boolean)) {
                if (!classes.includes(c)) classes.push(c);
            }
        }
        const className = classes.length > 0 ? ` class="${this.escape(classes.join(' '))}"` : '';

        // Handle ID and Anchors
        let idAttr = '';
        let extraAnchors = '';

        const anchorIds = this.config.ignoreInternalLinks ? [] : [...((node.metadata as any)?.anchorIds || [])];

        if (this.config.generateIds) {
            if (node.type === 'heading') {
                const slug = this.slugify(node.text || '');
                if (!anchorIds.includes(slug)) anchorIds.push(slug);
            } else if (node.type === 'sheet') {
                const sheetIndex = this.ast?.content.filter(n => n.type === 'sheet').indexOf(node) ?? 0;
                const sheetId = `sheet-${sheetIndex}`;
                if (!anchorIds.includes(sheetId)) anchorIds.push(sheetId);
            }
        }

        if (anchorIds.length > 0) {
            idAttr = ` id="${this.escape(anchorIds[0])}"`;
            if (anchorIds.length > 1) {
                extraAnchors = anchorIds.slice(1).map(aid => `<a id="${this.escape(aid)}" name="${this.escape(aid)}"></a>`).join('');
            }
        }

        // Inline Styles for structural nodes
        let styleAttr = '';
        if (this.config.includeFormatting && node.type !== 'text') {
            const styles = this.getInlineStyles(node);
            if (styles) styleAttr = ` style="${styles}"`;
        }

        switch (node.type) {
            case 'text':
                return this.formatText(node, node.text || '');

            case 'image': {
                if (!this.config.includeImages) return '';
                const meta = node.metadata as ImageMetadata;
                const attachmentName = meta?.attachmentName;
                let src = meta?.url || attachmentName || '';

                if (!meta?.url && attachmentName && this.ast) {
                    const attachment = this.ast.attachments.find(a => a.name === attachmentName);
                    if (attachment) {
                        src = `data:${attachment.mimeType || 'image/png'};base64,${attachment.data}`;
                    }
                }
                // Match CustomImage's exact data-width/data-align + style contract so a loaded
                // image re-hydrates the editor node without losing size/alignment.
                let imgDataAttrs = '';
                const imgStyleParts: string[] = [];
                const baseImgStyle = this.getInlineStyles(node);
                if (baseImgStyle) imgStyleParts.push(baseImgStyle);
                if (meta?.width) {
                    imgDataAttrs += ` data-width="${this.escape(meta.width)}"`;
                    // Sanitize before it enters the style="" attribute: an unescaped width
                    // (e.g. `1px" onerror="alert(1)`) would otherwise break out and inject an
                    // event handler, and a CSS `url(...)` would fetch a remote resource.
                    const safeWidth = sanitizeCssValue(meta.width);
                    if (safeWidth) imgStyleParts.push(`width: ${safeWidth}`);
                }
                if (meta?.align) {
                    imgDataAttrs += ` data-align="${this.escape(meta.align)}"`;
                    const ml = meta.align === 'left' ? '0' : 'auto';
                    const mr = meta.align === 'right' ? '0' : 'auto';
                    imgStyleParts.push('display: block', `margin-left: ${ml}`, `margin-right: ${mr}`);
                }
                const imgStyleAttr = imgStyleParts.length > 0 ? ` style="${imgStyleParts.join('; ')}"` : '';

                const img = `<img src="${sanitizeImageUrl(src)}" alt="${this.escape(node.text || meta?.altText || '')}"${className}${mappedAttrs}${imgDataAttrs}${imgStyleAttr}>`;
                const content = this.config.includeFormatting ? `<div class="image-container">${img}<div class="caption">${this.escape(attachmentName || '')}</div></div>` : img;
                return `${extraAnchors}<div${idAttr}>${content}</div>`;
            }

            case 'chart': {
                if (!this.config.includeCharts) return '';
                const meta = node.metadata as any; // ChartMetadata
                this.chartCounter++;
                const chartId = `chart-${this.chartCounter}`;
                const chartAttName = meta?.attachmentName;
                const chartAttachment = this.ast?.attachments.find(a => a.name === chartAttName);

                if (chartAttachment && (chartAttachment as any).chartData) {
                    const chartData = (chartAttachment as any).chartData;
                    const canvas = `<div class="chart-container"><canvas id="${chartId}"></canvas></div>`;
                    const script = `
<script>
    (function() {
        const initChart = () => {
            const ctx = document.getElementById('${chartId}').getContext('2d');
            const chartData = ${serializeForInlineScript(chartData)};
            const getRandomColor = (index, alpha) => {
                const colors = [
                    'rgba(255, 99, 132, ' + alpha + ')',
                    'rgba(54, 162, 235, ' + alpha + ')',
                    'rgba(255, 206, 86, ' + alpha + ')',
                    'rgba(75, 192, 192, ' + alpha + ')',
                    'rgba(153, 102, 255, ' + alpha + ')',
                    'rgba(255, 159, 64, ' + alpha + ')'
                ];
                return colors[index % colors.length];
            };
            if (typeof Chart === 'undefined') return;
            try {
                const canvas = document.getElementById('${chartId}');
                if (!canvas) return;
                const datasets = chartData.dataSets.map((ds, index) => ({
                    label: ds.name || 'Series ' + (index + 1),
                    data: ds.values.map(Number),
                    backgroundColor: getRandomColor(index, 0.5),
                    borderColor: getRandomColor(index, 1),
                    borderWidth: 1
                }));
                new Chart(canvas, {
                    type: 'bar',
                    data: {
                        labels: chartData.labels,
                        datasets: datasets
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            title: {
                                display: !!chartData.title,
                                text: chartData.title
                            }
                        }
                    }
                });
            } catch (e) {
                console.error('Failed to initialize chart ${chartId}:', e);
            }
        };

        let retries = 0;
        const tryInit = () => {
            if (typeof Chart !== 'undefined') {
                initChart();
            } else if (retries < 10) {
                retries++;
                setTimeout(tryInit, 500);
            }
        };

        if (document.readyState === 'complete') tryInit();
        else window.addEventListener('load', tryInit);
    })();
</script>`;
                    return `${extraAnchors}${canvas}${script}`;
                }
                return '';
            }

            case 'break':
                return (node.metadata as any)?.breakType === 'page' ? '<hr class="page-break">' : '<br>';

            case 'code': {
                const meta = node.metadata as CodeMetadata;
                if (meta?.math) {
                    // No pinned editor contract yet (inscript-editor's math node is v1.2,
                    // not built) - this is the proposed shape: data-math signals the
                    // display mode, and the visible text keeps its $ delimiters so the
                    // raw LaTeX degrades gracefully without a KaTeX renderer.
                    const delimited = meta.math === 'block' ? `$$${node.text || ''}$$` : `$${node.text || ''}$`;
                    const tag = meta.math === 'block' ? 'div' : 'span';
                    return `${extraAnchors}<${tag} class="math math-${this.escape(meta.math)}" data-math="${this.escape(meta.math)}"${idAttr}${mappedAttrs}${styleAttr}>${this.escape(delimited)}</${tag}>`;
                }
                const lang = meta?.language ? ` class="language-${this.escape(meta.language)}"` : '';
                const codeHtml = `<code${lang}>${this.escape(node.text || '')}</code>`;
                if (node.text && node.text.includes('\n')) {
                    return `${extraAnchors}<pre${idAttr}${className}${mappedAttrs}${styleAttr}>${codeHtml}</pre>`;
                } else {
                    return `${extraAnchors}<span${idAttr}${className}${mappedAttrs}${styleAttr}>${codeHtml}</span>`;
                }
            }

            case 'list': {
                const meta = node.metadata as ListMetadata;
                if (meta?.isTask) {
                    const checkedAttr = ` data-checked="${meta.checked ? 'true' : 'false'}"`;
                    const checkedBool = meta.checked ? ' checked' : '';
                    return `${extraAnchors}<li${checkedAttr}${idAttr}${className}${mappedAttrs}${styleAttr}><label><input type="checkbox"${checkedBool}><span></span></label><div>${childrenOutput}</div></li>`;
                }
                const value = (meta?.listType === 'ordered' && typeof meta.itemIndex === 'number')
                    ? ` value="${meta.itemIndex + 1}"`
                    : '';
                return `${extraAnchors}<li${value}${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</li>`;
            }

            case 'table': {
                // Smart Table Header Detection
                let finalChildren = childrenOutput;
                const rows = node.children || [];
                if (rows.length > 0 && rows[0].type === 'row') {
                    const firstRow = rows[0];
                    const firstRowCells = firstRow.children || [];

                    // Heuristic: Is the first row a header?
                    // 1. Explicitly marked via style containing "Header"
                    // 2. All cells have bold formatting
                    // 3. All cells have a background color different from the second row (if exists)
                    const isHeaderStyle = (firstRow.metadata as any)?.style?.toLowerCase().includes('header');
                    const allBold = firstRowCells.length > 0 && firstRowCells.every(c =>
                        c.children?.every(child => child.formatting?.bold === true)
                    );

                    if (isHeaderStyle || allBold) {
                        // Re-process children with <thead> wrap for the first row
                        const headOutput = await this.processNodeRecursive(firstRow, async (n, children) => {
                            return children.replace(/<td/g, '<th').replace(/<\/td>/g, '</th>');
                        });
                        const bodyRows = rows.slice(1);
                        const bodyOutput = await this.processNodeArray(bodyRows);
                        finalChildren = `<thead>${headOutput}</thead><tbody>${bodyOutput}</tbody>`;
                    }
                }
                // Match CustomTable's exact data-align + margin style contract so a loaded
                // table re-hydrates the editor node without losing its layout alignment.
                const tableMeta = node.metadata as TableMetadata;
                let tableDataAttrs = '';
                let tableStyleAttr = styleAttr;
                if (tableMeta?.align) {
                    tableDataAttrs = ` data-align="${this.escape(tableMeta.align)}"`;
                    const ml = tableMeta.align === 'left' ? '0' : 'auto';
                    const mr = tableMeta.align === 'right' ? '0' : 'auto';
                    const marginStyle = `margin-left: ${ml}; margin-right: ${mr}`;
                    tableStyleAttr = styleAttr
                        ? ` style="${styleAttr.replace(/^ style="|"$/g, '')}; ${marginStyle}"`
                        : ` style="${marginStyle}"`;
                }
                const tableHtml = `<table${idAttr}${className}${mappedAttrs}${tableDataAttrs}${tableStyleAttr}>${finalChildren}</table>`;
                const result = this.tableNestingLevel > 1 ? tableHtml : `<div class="table-container">${tableHtml}</div>`;
                return `${extraAnchors}${result}`;
            }

            case 'row': {
                if (node.children) {
                    let sparseChildren = '';
                    let lastCol = -1;
                    const cellNodes = node.children.filter(c => c.type === 'cell');
                    if (cellNodes.length > 0 && cellNodes.some(c => (c.metadata as any)?.col !== undefined)) {
                        for (const cell of cellNodes) {
                            const currentCol = (cell.metadata as any)?.col ?? (lastCol + 1);

                            // Fill gaps with empty cells
                            while (lastCol < currentCol - 1) {
                                sparseChildren += '<td></td>';
                                lastCol++;
                            }

                            sparseChildren += await this.processNodeRecursive(cell, this.nodeProcessor.bind(this));

                            const colSpan = (cell.metadata as any)?.colSpan || 1;
                            lastCol = currentCol + colSpan - 1;
                        }
                        return `<tr${idAttr}${className}${mappedAttrs}${styleAttr}>${sparseChildren}</tr>`;
                    }
                }
                return `<tr${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</tr>`;
            }

            case 'cell': {
                const meta = node.metadata as CellMetadata;
                const rowSpan = (meta?.rowSpan && meta.rowSpan > 1) ? ` rowspan="${meta.rowSpan}"` : '';
                const colSpan = (meta?.colSpan && meta.colSpan > 1) ? ` colspan="${meta.colSpan}"` : '';
                return `<td${rowSpan}${colSpan}${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</td>`;
            }

            case 'sheet': {
                const rows = node.children || [];

                // Find grid bounds
                let maxRow = -1;
                let maxCol = -1;

                for (const child of rows) {
                    if (child.type === 'row') {
                        for (const cell of child.children || []) {
                            if (cell.type === 'cell') {
                                const meta = cell.metadata as CellMetadata;
                                if (meta) {
                                    const r = meta.row;
                                    const c = meta.col;
                                    const rSpan = meta.rowSpan || 1;
                                    const cSpan = meta.colSpan || 1;
                                    if (r + rSpan - 1 > maxRow) maxRow = r + rSpan - 1;
                                    if (c + cSpan - 1 > maxCol) maxCol = c + cSpan - 1;
                                }
                            }
                        }
                    }
                }

                let tableHtml = '';

                if (maxRow >= 0 && maxCol >= 0) {
                    // Populate cell grid and track merged cells
                    const grid: (OfficeContentNode | null)[][] = Array.from(
                        { length: maxRow + 1 },
                        () => Array(maxCol + 1).fill(null)
                    );
                    const mergedCovered: boolean[][] = Array.from(
                        { length: maxRow + 1 },
                        () => Array(maxCol + 1).fill(false)
                    );
                    const rowNodeMap = new Map<number, OfficeContentNode>();

                    for (const child of rows) {
                        if (child.type === 'row') {
                            const cellsInRow = child.children?.filter(c => c.type === 'cell') || [];
                            if (cellsInRow.length > 0) {
                                const r = (cellsInRow[0].metadata as CellMetadata).row;
                                rowNodeMap.set(r, child);

                                for (const cell of cellsInRow) {
                                    const meta = cell.metadata as CellMetadata;
                                    if (meta) {
                                        const c = meta.col;
                                        if (r >= 0 && r <= maxRow && c >= 0 && c <= maxCol) {
                                            grid[r][c] = cell;

                                            const rSpan = meta.rowSpan || 1;
                                            const cSpan = meta.colSpan || 1;
                                            if (rSpan > 1 || cSpan > 1) {
                                                for (let rOffset = 0; rOffset < rSpan; rOffset++) {
                                                    for (let cOffset = 0; cOffset < cSpan; cOffset++) {
                                                        if (rOffset === 0 && cOffset === 0) continue;
                                                        const targetR = r + rOffset;
                                                        const targetC = c + cOffset;
                                                        if (targetR <= maxRow && targetC <= maxCol) {
                                                            mergedCovered[targetR][targetC] = true;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Build column headers (A, B, C...)
                    let ths = '<th class="excel-row-num-header"></th>';
                    for (let c = 0; c <= maxCol; c++) {
                        ths += `<th class="excel-col-header">${this.getColumnLetter(c)}</th>`;
                    }
                    const thead = `<thead><tr>${ths}</tr></thead>`;

                    // Build rows
                    let tbodyRows = '';
                    for (let r = 0; r <= maxRow; r++) {
                        const rowNode = rowNodeMap.get(r);
                        let trAttrs = '';
                        if (rowNode) {
                            const mapping = this.getSemanticMapping(rowNode);
                            const rClasses = ['excel-row'];
                            if (mapping?.classes) rClasses.push(...mapping.classes);
                            trAttrs += ` class="${rClasses.join(' ')}"`;
                            if (mapping?.attributes) {
                                for (const [key, val] of Object.entries(mapping.attributes)) {
                                    trAttrs += ` ${key}="${this.escape(val)}"`;
                                }
                            }
                            const rAnchorIds = this.config.ignoreInternalLinks ? [] : [...((rowNode.metadata as any)?.anchorIds || [])];
                            if (rAnchorIds.length > 0) {
                                trAttrs += ` id="${this.escape(rAnchorIds[0])}"`;
                            }
                            if (this.config.includeFormatting) {
                                const styles = this.getInlineStyles(rowNode);
                                if (styles) trAttrs += ` style="${styles}"`;
                            }
                        } else {
                            trAttrs = ' class="excel-row"';
                        }

                        let rowCellsHtml = `<td class="excel-row-num">${r + 1}</td>`;
                        for (let c = 0; c <= maxCol; c++) {
                            if (mergedCovered[r][c]) {
                                continue;
                            }
                            const cell = grid[r][c];
                            if (cell) {
                                const cellHtml = await this.processNodeRecursive(cell, this.nodeProcessor.bind(this));
                                rowCellsHtml += cellHtml;
                            } else {
                                rowCellsHtml += '<td class="excel-cell-empty"></td>';
                            }
                        }
                        tbodyRows += `<tr${trAttrs}>${rowCellsHtml}</tr>\n`;
                    }
                    const tbody = `<tbody>${tbodyRows}</tbody>`;
                    tableHtml = `<table class="spreadsheet-table excel-grid">${thead}${tbody}</table>`;
                }

                // Process non-row elements (images, charts, etc.)
                const nonRowNodes = rows.filter(c => c.type !== 'row');
                let nonRowHtml = '';
                if (nonRowNodes.length > 0) {
                    nonRowHtml = await this.processNodeArray(nonRowNodes);
                }

                const isFirstSheet = this.ast.content.filter(n => n.type === 'sheet')[0] === node;
                const isActive = isFirstSheet;
                const sheetIndex = this.ast.content.filter(n => n.type === 'sheet').indexOf(node);
                const sheetId = `sheet-${sheetIndex}`;

                // Merge classes correctly to avoid duplicate class attributes
                const mergedClasses = ['spreadsheet-sheet'];
                if (isActive) mergedClasses.push('active');
                if (classes.length > 0) mergedClasses.push(...classes);
                const classAttr = ` class="${mergedClasses.join(' ')}"`;

                // Ensure we don't have duplicate IDs
                const finalIdAttr = ` id="${sheetId}"`;

                return `${extraAnchors}<div${finalIdAttr}${classAttr}${mappedAttrs}${styleAttr}>${tableHtml}${nonRowHtml}</div>`;
            }

            case 'paragraph':
            case 'heading': {
                const tag = node.type === 'heading' ? `h${(node.metadata as HeadingMetadata)?.level || 1}` : 'p';

                // Normalize empty paragraphs so DOCX and PPTX empty cells render with consistent height
                // Strip tags to check if it's purely empty or just contains non-breaking spaces (like PPTX)
                const textOnly = childrenOutput.replace(/<[^>]+>/g, '').trim();
                if (!textOnly && !node.children?.some(c => c.type === 'image' || c.type === 'chart')) {
                    const extraClass = className ? ` class="${className.replace('class="', '').replace('"', '')} empty-paragraph"` : ' class="empty-paragraph"';
                    return `${extraAnchors}<${tag}${idAttr}${extraClass}${mappedAttrs}${styleAttr}><br></${tag}>`;
                }

                return `${extraAnchors}<${tag}${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</${tag}>`;
            }
            case 'slide': {
                const meta = node.metadata as SlideMetadata;
                const slideNum = this.escape(String(meta?.slideNumber || ''));
                return `${extraAnchors}<section class="slide" data-slide-num="${slideNum}"${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</section>`;
            }
            case 'page': {
                const meta = node.metadata as PageMetadata;
                const pageNum = this.escape(String(meta?.pageNumber || ''));
                return `${extraAnchors}<section class="page" data-page-num="${pageNum}"${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</section>`;
            }
            case 'note': {
                const meta = node.metadata as NoteMetadata;
                if (meta?.noteType === 'footnote' || meta?.noteType === 'endnote') {
                    const key = this.escape(this.getFootnoteKey(node));
                    return `<p id="footnote-${key}" data-footnote-id="${key}">${childrenOutput} <a href="#footnote-ref-${key}">↩</a></p>`;
                }
                const noteClass = meta?.noteType ? ` note-${this.escape(meta.noteType)}` : '';
                return `${extraAnchors}<div class="slide-note${noteClass}"${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</div>`;
            }

            case 'embed': {
                // Match the Youtube extension's exact wrapper shape so a loaded embed
                // re-hydrates the editor's Youtube node.
                const meta = node.metadata as EmbedMetadata;
                const id = meta?.videoId || '';
                const width = meta?.width || '100%';
                const align = meta?.align || 'center';
                const ml = align === 'left' ? '0' : 'auto';
                const mr = align === 'right' ? '0' : 'auto';
                const iframe = id
                    ? `<iframe src="https://www.youtube.com/embed/${this.escape(id)}" title="YouTube video player" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen></iframe>`
                    : '';
                return `${extraAnchors}<div data-youtube-video="${this.escape(id)}" data-width="${this.escape(width)}" data-align="${this.escape(align)}" class="youtube-embed"${idAttr}${mappedAttrs} style="width: ${sanitizeCssValue(width)}; margin-left: ${ml}; margin-right: ${mr};">${iframe}</div>`;
            }

            case 'admonition': {
                // Match inscript-editor's Admonition node wrapper so a loaded admonition
                // reaches the editor as that node instead of a plain blockquote.
                const meta = node.metadata as AdmonitionMetadata;
                const admonitionType = this.escape(meta?.admonitionType || 'note');
                return `${extraAnchors}<div class="admonition admonition-${admonitionType}" data-type="${admonitionType}"${idAttr}${mappedAttrs}${styleAttr}>${childrenOutput}</div>`;
            }

            case 'definitionList':
                return `${extraAnchors}<dl${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</dl>`;

            case 'definitionTerm':
                return `${extraAnchors}<dt${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</dt>`;

            case 'definitionDescription':
                return `${extraAnchors}<dd${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</dd>`;

            default:
                return childrenOutput;
        }
    }


    private getDefaultTag(node: OfficeContentNode): string {
        switch (node.type) {
            case 'paragraph': return 'p';
            case 'heading': {
                const level = (node.metadata as HeadingMetadata)?.level || 1;
                return `h${Math.min(Math.max(level, 1), 6)}`;
            }
            case 'list': return 'li';
            default: return 'div';
        }
    }

    private formatText(node: OfficeContentNode, text: string): string {
        let result = this.escape(text);
        const f = node.formatting;

        if (this.config.includeFormatting && f) {
            if (f.bold) result = `<b>${result}</b>`;
            if (f.italic) result = `<i>${result}</i>`;
            if (f.underline) result = `<u>${result}</u>`;
            if (f.strikethrough) result = `<strike>${result}</strike>`;
            if (f.subscript) result = `<sub>${result}</sub>`;
            if (f.superscript) result = `<sup>${result}</sup>`;

            const styles = this.getInlineStyles(node);
            if (styles) {
                result = `<span style="${styles}">${result}</span>`;
            }
        }

        const meta = node.metadata as TextMetadata;
        if (meta?.wikilink) {
            // No pinned editor contract yet for wikilinks - data-wikilink-page preserves the
            // exact page name separately from the display text/alias and from href (which
            // the host app's resolver may rewrite to a real URL).
            if (!this.config.ignoreInternalLinks) {
                result = `<a href="#${this.escape(this.slugify(meta.link || ''))}" data-wikilink-page="${this.escape(meta.link || '')}">${result}</a>`;
            }
        } else if (meta?.link) {
            const isInternal = meta.linkType !== 'external';
            if (!this.config.ignoreInternalLinks || !isInternal) {
                result = `<a href="${sanitizeUrl(meta.link)}"${meta.linkType === 'external' ? ' target="_blank"' : ''}>${result}</a>`;
            }
        }
        if (meta?.abbreviationTitle) {
            result = `<abbr title="${this.escape(meta.abbreviationTitle)}">${result}</abbr>`;
        }
        if (meta?.citationKey) {
            // No pinned editor contract yet for citations (unlike footnotes/admonitions) -
            // this is the proposed shape: a <cite> carrying the bare key, matching Pandoc's
            // [@citekey] on the Markdown side. Keep in sync if inscript-editor's citation
            // node lands with a different attribute name.
            result = `<cite data-citation-key="${this.escape(meta.citationKey)}">[@${this.escape(meta.citationKey)}]</cite>`;
        }

        return result;
    }

    private getInlineStyles(node: OfficeContentNode): string {
        const styles: string[] = [];

        // Colors/sizes/fonts/alignments are free strings from an untrusted document;
        // run each through sanitizeCssValue so it can't break out of the style="" attribute
        // or inject a resource-fetching CSS construct. Drop the declaration if nothing survives.
        const pushSafe = (prop: string, value: string) => {
            const safe = sanitizeCssValue(value);
            if (safe) styles.push(`${prop}: ${safe}`);
        };

        if (node.metadata) {
            const meta = node.metadata as any;
            if (meta.alignment) pushSafe('text-align', meta.alignment);
            if (meta.backgroundColor) pushSafe('background-color', meta.backgroundColor);
            if (meta.verticalAlign) pushSafe('vertical-align', meta.verticalAlign);
            if (meta.paragraphIndentation) {
                const ind = meta.paragraphIndentation;
                if (ind.left) styles.push(`margin-left: ${ind.left / 20}pt`);
                if (ind.right) styles.push(`margin-right: ${ind.right / 20}pt`);
                if (ind.firstLine) styles.push(`text-indent: ${ind.firstLine / 20}pt`);
            }
        }

        if (node.formatting) {
            const f = node.formatting;
            if (f.color) pushSafe('color', f.color);
            if (f.backgroundColor) pushSafe('background-color', f.backgroundColor);
            if (f.size) pushSafe('font-size', f.size);
            if (f.font) {
                const safeFont = sanitizeCssValue(f.font);
                if (safeFont) styles.push(`font-family: ${safeFont}, sans-serif`);
            }
        }

        return styles.join('; ');
    }

    private getPremiumStyles(isSpreadsheet: boolean = false, isPresentation: boolean = false, isPdf: boolean = false): string {
        let resolvedWidth = this.config.htmlConfig.containerWidth;
        if (!resolvedWidth || resolvedWidth === 'auto') {
            if (isSpreadsheet) {
                resolvedWidth = '100%';
            } else if (isPresentation) {
                resolvedWidth = '297mm';
            } else {
                resolvedWidth = '900px';
            }
        } else if (typeof resolvedWidth === 'number') {
            resolvedWidth = `${resolvedWidth}px`;
        }

        return `
            :root {
                --primary-color: #2c3e50;
                --text-color: #333;
                --bg-color: #f3f4f6;
                --container-bg: #ffffff;
                --border-color: #e9ecef;
                --accent-color: #3498db;
                --shadow: 0 10px 25px rgba(0,0,0,0.05);
                --container-width: ${resolvedWidth};
            }
            * {
                box-sizing: border-box;
            }
            body {
                font-family: 'Inter', -apple-system, sans-serif;
                background-color: var(--bg-color);
                color: var(--text-color);
                line-height: 1.6;
                margin: 0;
                padding: ${isSpreadsheet ? '0' : '50px 20px'};
            }
            .container {
                max-width: var(--container-width);
                margin: 0 auto;
                background: var(--container-bg);
                padding: 60px 80px;
                border-radius: 12px;
                box-shadow: var(--shadow);
            }
            .presentation-container, .pdf-container {
                max-width: var(--container-width);
                margin: 0 auto;
            }
            .presentation-container .metadata-summary, .pdf-container .metadata-summary {
                background: white;
                margin-bottom: 40px;
                padding: 30px;
                border-radius: 12px;
                box-shadow: var(--shadow);
            }
            .chart-container {
                margin: 30px auto;
                max-width: 900px;
                padding: 25px;
                background: white;
                border: 1px solid var(--border-color);
                border-radius: 16px;
                box-shadow: var(--shadow);
                height: 450px;
            }
            .chart-container canvas {
                width: 100% !important;
                height: 100% !important;
            }
            .spreadsheet-container {
                width: 100%;
                height: 100vh;
                display: flex;
                flex-direction: column;
                background: white;
            }
            .spreadsheet-container article {
                flex: 1;
                overflow: hidden;
                display: flex;
                flex-direction: column;
            }
            h1, h2, h3, h4, h5, h6 {
                color: var(--primary-color);
                margin-top: 1.6em;
                margin-bottom: 0.8em;
                font-weight: 700;
            }
            h1 { font-size: 2.4em; border-bottom: 2px solid var(--border-color); padding-bottom: 15px; }
            h2 { font-size: 1.9em; }
            h3 { font-size: 1.5em; }
            
            p { margin-bottom: 1.3em; }
            p.empty-paragraph { margin: 0; min-height: 1em; }
            
            ul, ol { margin: 1em 0; padding-left: 2em; }
            li { margin-bottom: 0.25em; }
            li > p { margin-bottom: 0.25em; }
            
            .table-container {
                width: fit-content;
                max-width: 100%;
                overflow-x: auto;
                -webkit-overflow-scrolling: touch;
                margin: 25px auto;
                border: 1px solid var(--border-color);
                border-radius: 8px;
                background: white;
            }
            table {
                width: auto;
                max-width: 100%;
                min-width: ${isSpreadsheet ? '100%' : '300px'};
                border-collapse: separate;
                border-spacing: 0;
                margin: 0;
                border: none;
            }
            th, td {
                padding: 8px 12px;
                border-bottom: 1px solid var(--border-color);
                border-right: 1px solid var(--border-color);
                text-align: left;
                vertical-align: top;
                overflow-wrap: break-word;
            }
            th:last-child, td:last-child {
                border-right: none;
            }
            tr:last-child th, tr:last-child td {
                border-bottom: none;
            }
            th, td > p {
                margin: 0;
            }
            td > *:last-child, th > *:last-child {
                margin-bottom: 0;
            }
            th {
                background-color: #f8f9fa;
                color: var(--primary-color);
                font-weight: 600;
                text-transform: uppercase;
                font-size: 0.85em;
                letter-spacing: 0.05em;
                position: ${isSpreadsheet ? 'sticky' : 'static'};
                top: 0;
                z-index: 10;
            }
            tr:nth-child(even) { background-color: #fdfdfd; }
            tr:hover { background-color: #f1f4f9; }
            
            /* Spreadsheet Specific */
            .spreadsheet-sheet {
                display: none;
                flex: 1;
                overflow: auto;
                position: relative;
            }
            .spreadsheet-sheet.active {
                display: block;
            }
            .spreadsheet-sheet td {
                padding: 8px 12px;
                font-size: 13px;
                white-space: nowrap;
            }
            /* Spreadsheet Grid styling */
            .excel-grid {
                border-collapse: collapse;
                border-spacing: 0;
                background: var(--container-bg);
                border: 1px solid var(--border-color);
                width: max-content;
                max-width: none;
                min-width: 0;
            }
            .excel-grid th, .excel-grid td {
                border: 1px solid var(--border-color);
                padding: 4px 8px;
                font-size: 13px;
                line-height: 1.2;
                overflow: hidden;
                text-overflow: ellipsis;
                white-space: nowrap;
            }
            .excel-col-header {
                background: #f8f9fa;
                color: #5f6368;
                font-weight: 500;
                text-align: center;
                user-select: none;
                border-bottom: 2px solid var(--border-color);
                position: sticky;
                top: 0;
                z-index: 10;
                min-width: 100px;
            }
            .col-resizer {
                position: absolute;
                top: 0;
                right: 0;
                width: 4px;
                height: 100%;
                cursor: col-resize;
                user-select: none;
                z-index: 20;
            }
            .col-resizer:hover, .col-resizer.resizing {
                background: var(--accent-color, #3498db);
            }
            .excel-row-num {
                background: #f8f9fa;
                color: #5f6368;
                text-align: center;
                font-weight: 500;
                width: 45px;
                min-width: 45px !important;
                max-width: 45px;
                user-select: none;
                border-right: 2px solid var(--border-color);
                position: sticky;
                left: 0;
                z-index: 5;
            }
            .row-resizer {
                position: absolute;
                bottom: 0;
                left: 0;
                width: 100%;
                height: 4px;
                cursor: row-resize;
                user-select: none;
                z-index: 20;
            }
            .row-resizer:hover, .row-resizer.resizing {
                background: var(--accent-color, #3498db);
            }
            .excel-row-num-header {
                background: #f1f3f4;
                width: 45px;
                min-width: 45px !important;
                max-width: 45px;
                position: sticky;
                top: 0;
                left: 0;
                z-index: 15;
                border-right: 2px solid var(--border-color);
                border-bottom: 2px solid var(--border-color);
            }
            .excel-cell-empty {
                background: var(--container-bg);
            }
            .excel-grid tr:hover td {
                background-color: #f1f4f9;
            }
            .excel-grid tr:hover td.excel-row-num {
                background-color: #f8f9fa; /* Keep header color */
            }

            /* Tab Bar */
            .spreadsheet-tabs {
                background: #f1f3f4;
                border-top: 1px solid var(--border-color);
                display: flex;
                padding: 0 20px;
                z-index: 9999;
                height: 35px;
                align-items: center;
                overflow-x: auto;
                box-shadow: 0 -2px 10px rgba(0,0,0,0.05);
            }
            .spreadsheet-tab {
                padding: 0 20px;
                height: 100%;
                display: flex;
                align-items: center;
                text-decoration: none;
                color: #5f6368;
                font-size: 13px;
                border-right: 1px solid var(--border-color);
                background: #f1f3f4;
                white-space: nowrap;
                transition: all 0.2s;
                cursor: pointer;
            }
            .spreadsheet-tab:hover { background: #e8eaed; }
            .spreadsheet-tab.active { background: white; color: var(--accent-color); font-weight: 600; border-bottom: 2px solid var(--accent-color); }

            /* Slide & Page Separation */
            .slide {
                background: white;
                aspect-ratio: 297 / 210;
                margin: 0 auto 40px auto;
                padding: 4% 6%;
                border-radius: 12px;
                box-shadow: 0 10px 30px rgba(0,0,0,0.1);
                position: relative;
                display: flex;
                flex-direction: column;
                justify-content: flex-start;
                border: 1px solid var(--border-color);
                box-sizing: border-box;
                width: 100%;
                overflow: hidden;
                overflow-y: auto;
                overflow-wrap: break-word;
                page-break-after: always;
            }
            .slide:has(+ .slide-note) {
                margin-bottom: 0;
                border-bottom-left-radius: 0;
                border-bottom-right-radius: 0;
                border-bottom: none;
            }
            .slide table {
                width: auto;
                max-width: 100%;
                font-size: 0.85em;
                margin: 10px 0;
            }
            .slide th, .slide td {
                padding: 6px 10px;
            }
            .slide > * {
                flex-shrink: 0;
            }
            .slide > .table-container {
                flex-shrink: 1;
                min-height: 0;
            }
            .slide .table-container {
                margin: 10px auto;
                overflow-y: auto;
            }
            .slide .chart-container {
                margin: 15px auto;
                max-width: 100%;
                padding: 10px;
                height: 240px;
                box-shadow: none;
                border-radius: 8px;
            }
            .slide img {
                max-height: 200px;
                width: auto;
                margin: 10px auto;
                box-shadow: none;
            }
            .slide .image-container {
                margin: 15px 0;
            }
            .slide h1 { font-size: 1.8em; margin-top: 0.4em; margin-bottom: 0.3em; padding-bottom: 5px; }
            .slide h2 { font-size: 1.4em; margin-top: 0.4em; }
            .slide h3 { font-size: 1.25em; }
            .slide p { margin-bottom: 0.6em; font-size: 0.95em; }
            .slide li > p { margin-bottom: 0.25em; }
            .slide ul, .slide ol { margin: 0.6em 0; }
            .slide li { margin-bottom: 0.25em; }
            .slide-note {
                background: #fdfdfd;
                border: 1px solid var(--border-color);
                border-top: 1px dashed var(--border-color);
                border-radius: 0 0 12px 12px;
                padding: 30px 50px;
                margin: 0 auto 40px auto;
                width: 100%;
                box-sizing: border-box;
                font-size: 0.95em;
                color: var(--text-color);
                position: relative;
                box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            }
            .note-footnote, .note-endnote {
                margin: 30px auto;
                max-width: 90%;
            }
            .slide-note::before {
                content: "SLIDE NOTES";
                font-size: 0.7rem;
                font-weight: 800;
                color: #b2bec3;
                display: block;
                margin-bottom: 12px;
                letter-spacing: 1.5px;
            }
            .note-footnote::before { content: "FOOTNOTE" !important; }
            .note-endnote::before { content: "ENDNOTE" !important; }
            .slide-note p { margin-bottom: 0.8em; }

            .slide::after {
                content: "Slide " attr(data-slide-num);
                position: absolute;
                bottom: 20px;
                right: 30px;
                font-size: 0.8em;
                color: #999;
                font-weight: 500;
            }
            .page {
                background: white;
                aspect-ratio: 1 / 1.4142;
                margin: 0 auto 40px auto;
                padding: 8% 10%;
                box-shadow: 0 5px 20px rgba(0,0,0,0.08);
                position: relative;
                border: 1px solid var(--border-color);
                box-sizing: border-box;
                width: 100%;
                page-break-after: always;
            }
            .page::after {
                content: "Page " attr(data-page-num);
                position: absolute;
                bottom: 20px;
                right: 30px;
                font-size: 0.8em;
                color: #999;
                font-weight: 500;
            }
            
            /* High Fidelity Presentation Mode */
            @media screen and (max-width: 800px) {
                .slide { min-height: auto; aspect-ratio: auto; padding: 40px 30px; }
                .page { aspect-ratio: auto; padding: 40px; }
            }
            
            /* Nested Table Styles */
            td table {
                margin: 15px 0;
                border-radius: 6px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.03);
                background-color: #ffffff;
                border: 1px solid var(--border-color);
                overflow: hidden;
            }
            td td {
                padding: 10px 14px;
                font-size: 0.95em;
            }
            
            img {
                max-width: 100%;
                height: auto;
                border-radius: 10px;
                display: block;
                margin: 40px auto;
                box-shadow: 0 15px 35px rgba(0,0,0,0.1);
            }

            .image-container { text-align: center; margin: 30px 0; }
            .caption { font-size: 0.8em; color: #636e72; margin-top: 8px; font-style: italic; }
            
            body { margin: 0; padding: 0; }
            
            .page-break { border: none; border-top: 2px dashed var(--border-color); margin: 40px 0; position: relative; }
            .page-break::after { content: 'PAGE BREAK'; position: absolute; top: -10px; left: 50%; transform: translateX(-50%); background: white; padding: 0 15px; font-size: 10px; color: #b2bec3; font-weight: bold; letter-spacing: 1px; }

            /* Metadata Styles */
            .metadata-summary { 
                background: #f8f9fa; 
                border: 1px solid #e9ecef; 
                border-radius: 12px; 
                padding: 25px; 
                margin: ${isSpreadsheet ? '20px' : '0 0 40px 0'};
            }
            .meta-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 15px; }
            .meta-item { border-bottom: 1px solid #eee; padding-bottom: 8px; }
            .meta-label { font-size: 0.65rem; color: #adb5bd; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px; }
            .meta-value { font-size: 0.9rem; color: #495057; font-weight: 600; }
            .meta-custom-section { margin-top: 20px; border-top: 2px solid #eee; padding-top: 15px; }
            .meta-section-title { font-size: 0.75rem; color: var(--accent-color); font-weight: 700; margin-bottom: 10px; }
            .meta-tags-grid { display: flex; flex-wrap: wrap; gap: 8px; }
            .meta-tag { background: rgba(52, 152, 219, 0.05); border: 1px solid rgba(52, 152, 219, 0.1); padding: 4px 10px; border-radius: 6px; font-size: 0.8rem; color: #495057; }
            
            span[style*="color"] { font-weight: inherit; }

            /* --- Print Optimization --- */
            @media print {
                @page {
                    margin: 1.5cm;
                }

                body {
                    background: white !important;
                    color: black !important;
                    margin: 0 !important;
                    padding: 0 !important;
                }
                
                /* Hide web-only interactive elements */
                .spreadsheet-tabs, .sync-btn, button, .reparse-btn {
                    display: none !important;
                }

                /* Flatten Spreadsheets: Show all sheets in PDF */
                .spreadsheet-sheet {
                    display: block !important;
                    opacity: 1 !important;
                    visibility: visible !important;
                    page-break-after: always !important;
                    margin-bottom: 3rem !important;
                    height: auto !important;
                    min-height: auto !important;
                    overflow: visible !important;
                }

                .page, .slide, .metadata-summary, .container, .pdf-container, .presentation-container, article {
                    box-shadow: none !important;
                    border: none !important;
                    page-break-inside: avoid !important;
                    break-inside: avoid !important;
                    margin-bottom: 2rem !important;
                    max-width: none !important;
                    width: 100% !important;
                    height: auto !important;
                    min-height: auto !important;
                    overflow: visible !important;
                    display: block !important;
                }

                h1, h2, h3, h4, h5, h6 {
                    page-break-after: avoid !important;
                    break-after: avoid !important;
                }

                table, tr, img, .chart-container, li, .image-container {
                    page-break-inside: avoid !important;
                    break-inside: avoid !important;
                }

                a {
                    text-decoration: none !important;
                    color: black !important;
                }

                /* Avoid orphans/widows */
                p, li {
                    orphans: 3;
                    widows: 3;
                }

                /* Ensure full width and reset web-specific heights */
                .page, .slide, .spreadsheet-sheet {
                    width: 100% !important;
                    max-width: none !important;
                    min-height: auto !important;
                    padding: 0 !important;
                    margin: 0 0 2rem 0 !important;
                    border: none !important;
                    display: block !important;
                    height: auto !important;
                }
            }

            /* --- Custom User CSS --- */
            ${this.config.htmlConfig.customCss}
        `;
    }

    /**
     * Same as `getPremiumStyles()`, but wrapped in a CSS `@scope` block anchored to the
     * `.op-html-scope` wrapper so the rules only apply within the generated fragment - they
     * cannot leak onto a host page's own elements. `:root` and `body` selectors specifically
     * target the real page root/body, so they're remapped to `:scope` (the scope root, i.e. the
     * `.op-html-scope` wrapper) first; every other selector is naturally confined by `@scope`
     * without needing per-selector rewriting. `customCss` is included in this scoping too.
     */
    private getScopedPremiumStyles(isSpreadsheet: boolean = false, isPresentation: boolean = false, isPdf: boolean = false): string {
        const css = this.getPremiumStyles(isSpreadsheet, isPresentation, isPdf)
            .replace(/:root(\s*\{)/g, ':scope$1')
            .replace(/(^|\n)(\s*)body(\s*\{)/g, '$1$2:scope$3');
        return `@scope (.op-html-scope) {\n${css}\n}`;
    }

    protected override slugify(text: string): string {
        return text.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
    }

    private getColumnLetter(colIndex: number): string {
        let temp = colIndex;
        let letter = '';
        while (temp >= 0) {
            letter = String.fromCharCode((temp % 26) + 65) + letter;
            temp = Math.floor(temp / 26) - 1;
        }
        return letter;
    }

    // Attribute/text escaping, URL sanitizing, and inline-script serialization all
    // live in ../utils/sanitize.js so every generator shares one implementation.
    // escape() stays as a thin wrapper because it has many call sites here.
    private escape(text: string): string {
        return escapeHtml(text);
    }

    /** Converts a document-supplied date to an ISO string, or '' if it is invalid
     *  (a malformed date would otherwise throw a RangeError and abort generation). */
    private toIsoDate(value: unknown): string {
        if (value === undefined || value === null || value === '') return '';
        const d = new Date(value as any);
        return isNaN(d.getTime()) ? '' : d.toISOString();
    }
}

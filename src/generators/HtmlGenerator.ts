import { AdmonitionMetadata, CellMetadata, CodeMetadata, ConversionResult, EmbedMetadata, GeneratorConfig, HeadingMetadata, ImageMetadata, ListMetadata, NoteMetadata, OfficeContentNode, OfficeParserAST, OfficeWarningType, PageMetadata, SlideMetadata, StandaloneConfig, TableMetadata, TextMetadata } from '../types.js';
import { BaseGenerator } from './BaseGenerator.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { base64ByteLength, documentLanguage, isHeaderRow, resolveEmbed } from '../utils/officeGenUtils.js';
import { escapeHtml, isSafeHtmlAttributeName, isSafeStyleMapTag, sanitizeCommentText, sanitizeCssValue, sanitizeUrl, sanitizeImageUrl, serializeForInlineScript } from '../utils/sanitize.js';
import { isSourceComment } from '../utils/commentUtils.js';

/** A child that is a block of its own (a code block or display equation), which a <p> cannot hold. */
const isBlockInParagraph = (child: OfficeContentNode): boolean =>
    (child.type === 'code' && (child.metadata as CodeMetadata | undefined)?.math !== 'inline')
    // A rule or page break is an <hr>, which cannot sit in a <p> either (a Word page break is a break in its paragraph).
    || (child.type === 'break' && ['thematic', 'page'].includes((child.metadata as { breakType?: string } | undefined)?.breakType ?? ''));

type ResolvedStandalone = Required<StandaloneConfig>;

/**
 * A boxed `onNode` verdict, passed from the caller that already asked the hook down to the recursion
 * that would otherwise ask it again. Boxing matters because "render normally" is itself `undefined`,
 * so only the presence of the box distinguishes an answered hook from an unasked one - `onNode` is
 * allowed to have side effects, and must fire exactly once per node.
 */
type OnNodeVerdict = { value: string | false | void };

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
        if (!isSafeHtmlAttributeName(key)) continue;
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
 * Whether a colour is effectively the document default: near-black or near-white. Consulted only when
 * `htmlConfig.omitDefaultTextColor` is on, to drop a run colour that would otherwise pin imported text
 * to black or white regardless of the reader's theme. Parses `#rgb`/`#rrggbb`; anything else (named
 * colours, `rgb(...)`) is treated as a deliberate colour and kept.
 */
function isNearDefaultColor(color: string): boolean {
    let h = color.trim().replace(/^#/, '');
    if (h.length === 3) h = h.split('').map(c => c + c).join('');
    if (!/^[0-9a-fA-F]{6}$/.test(h)) return false;
    const r = parseInt(h.slice(0, 2), 16), g = parseInt(h.slice(2, 4), 16), b = parseInt(h.slice(4, 6), 16);
    return (r <= 24 && g <= 24 && b <= 24) || (r >= 231 && g >= 231 && b >= 231);
}

/**
 * Generates semantic, high-fidelity HTML from an AST.
 */
export class HtmlGenerator extends BaseGenerator<'html'> {
    private chartCounter = 0;
    private isSpreadsheetMode = false;
    /**
     * Set while rendering a heading's children, so `formatText` can drop the run-level bold and
     * font-size the `<hN>` already establishes. See the note there, and the identical flag in
     * `RtfGenerator`, where the same inherited size actively shrinks the heading.
     */
    private inHeading = false;
    /** As `inHeading`, but for the inherited font size - see `hasUniformFormatting`. */
    private headingUniformSize = false;
    /** Memoized: does this run emit a full standalone HTML document (vs a fragment)? */
    private _standaloneDoc?: boolean;
    private get emitsStandaloneDocument(): boolean {
        if (this._standaloneDoc === undefined) this._standaloneDoc = resolveStandalone(this.config.htmlConfig.standalone).document;
        return this._standaloneDoc;
    }

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
        // Key off the source format, not "has a page node": ODG (Draw) also emits `page` nodes but is
        // not a PDF, so it should not get the PDF container class or premium PDF styles.
        const isPdf = this.ast.type === 'pdf';

        let containerClass = 'container';
        if (this.isSpreadsheetMode) containerClass = 'spreadsheet-container';
        else if (isPresentation) containerClass = 'presentation-container';
        else if (isPdf) containerClass = 'pdf-container';

        let bodyContent = await this.processNodeArray(this.ast.content);

        if (this.collectedNotes.length > 0) {
            // De-duplicate by node identity first. A table row with sparse column metadata
            // re-processes its cells in `case 'row'` after they were already processed for
            // `childrenOutput`, so a footnote referenced inside a cell gets pushed here twice -
            // the same object reference both times, which a Set collapses back to one. Genuinely
            // distinct notes (even two references to the same id) are different objects and stay.
            const collectedNotes = [...new Set(this.collectedNotes)];
            // Footnotes/endnotes get their own <section data-footnotes> (the agreed
            // contract with attribute-driven editors' footnote nodes); other note types (e.g.
            // slide speaker notes) keep the existing generic notes wrapper.
            const footnotes = collectedNotes.filter(n => {
                const t = (n.metadata as any)?.noteType;
                return t === 'footnote' || t === 'endnote';
            });
            const otherNotes = collectedNotes.filter(n => !footnotes.includes(n));

            if (footnotes.length > 0) {
                let footnotesHtml = '';
                for (const note of footnotes) {
                    footnotesHtml += await this.processNodeRecursive(note, this.boundNodeProcessor, this.collectedNoteOverrides.get(note));
                }
                // data-footnotes carries an explicit empty value (not a bare attribute) so
                // the markup is valid XHTML too - EpubGenerator embeds this verbatim, and
                // XML rejects valueless attributes. HtmlParser only checks for presence.
                bodyContent += `\n<section data-footnotes="">\n${footnotesHtml}\n</section>\n`;
            }

            if (otherNotes.length > 0) {
                let notesHtml = '';
                for (const note of otherNotes) {
                    notesHtml += await this.processNodeRecursive(note, this.boundNodeProcessor, this.collectedNoteOverrides.get(note));
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
            // The tab bar lists the sheets the body actually rendered. That decision was already made
            // (and recorded) during the body walk above, so read it back rather than asking `onNode` a
            // second time about every sheet - the hook is allowed to have side effects.
            const sheets = this.ast.content.filter(node => node.type === 'sheet' && !this.onNodeSkipped.has(node));
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
<html lang="${documentLanguage(this.effectiveMetadata)}">
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
        // Both are parsed from every OOXML/ODF core-properties block but had no output sink at
        // all, so they were silently dropped on every conversion. `keywords` is the standard HTML
        // meta name; `subject` has no standard equivalent, so it uses its Dublin Core name.
        if (m.keywords) tags.push(`<meta name="keywords" content="${this.escape(m.keywords)}">`);
        if (m.subject) tags.push(`<meta name="DC.subject" content="${this.escape(m.subject)}">`);

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
    /**
     * Renders a table's rows with an HTML grid-occupancy model, so a cell's `rowSpan` reserves its
     * column in the rows below rather than letting the plain per-row gap-filler shift the cells that
     * sit under it. Returns one `<tr>…</tr>` string per row (row 0's `<td>`s are promoted to `<th>`
     * when `headerFirst`). Only invoked for tables that actually contain a rowSpan.
     */
    private async renderRowsWithRowspans(rows: OfficeContentNode[], headerFirst: boolean): Promise<{ trs: string[]; headerRowEmitted: boolean }> {
        const out: string[] = [];
        // Whether the first (header) row actually survived to `out[0]`: if `onNode` drops it, `out[0]`
        // is a body row and the caller must not wrap it in <thead>.
        let headerRowEmitted = false;
        // grid column -> number of rows it stays occupied by a rowspan started above this row.
        const carry = new Map<number, number>();
        for (let r = 0; r < rows.length; r++) {
            // This path builds each <tr> itself rather than recursing into the row node, so the row
            // has to be offered to `onNode` here or it would be the one node type the hook never saw
            // in a rowspan table: `false` drops the row, a string replaces its markup outright.
            const rowOverride = await this.handleOnNode(rows[r]);
            if (rowOverride === false) continue;
            if (typeof rowOverride === 'string') { out.push(rowOverride); if (r === 0) headerRowEmitted = true; continue; }
            const cells = (rows[r].children || []).filter(c => c.type === 'cell');
            const newCarry = new Map<number, number>();
            let col = 0;
            let ci = 0;
            let tr = '';
            while (ci < cells.length) {
                while ((carry.get(col) || 0) > 0) col++;               // skip columns held by a rowspan
                const cell = cells[ci];
                const meta = cell.metadata as CellMetadata;
                const target = typeof meta?.col === 'number' ? meta.col : col;
                while (col < target) {                                  // genuine empty gaps
                    if ((carry.get(col) || 0) > 0) { col++; continue; }
                    tr += '<td></td>';
                    col++;
                }
                let cellHtml = await this.processNodeRecursive(cell, this.boundNodeProcessor);
                // Promote the cell's own tag to <th> for the header row. Match the first <td> (not the
                // string start) so a leading extraAnchors prefix does not defeat the promotion, and the
                // last </td> (anchored at end) so a nested table's inner cells are left as <td>.
                if (headerFirst && r === 0) cellHtml = cellHtml.replace(/<td/, '<th').replace(/<\/td>$/, '</th>');
                tr += cellHtml;
                const cSpan = (meta?.colSpan && meta.colSpan > 1) ? meta.colSpan : 1;
                const rSpan = (meta?.rowSpan && meta.rowSpan > 1) ? meta.rowSpan : 1;
                if (rSpan > 1) for (let cc = col; cc < col + cSpan; cc++) newCarry.set(cc, rSpan - 1);
                col += cSpan;
                ci++;
            }
            // Age existing carries by one row, then fold in rowspans this row started.
            for (const [c, v] of [...carry.entries()]) { if (v > 1) carry.set(c, v - 1); else carry.delete(c); }
            for (const [c, v] of newCarry) carry.set(c, v);
            out.push(`<tr>${tr}</tr>`);
            if (r === 0) headerRowEmitted = true;
        }
        return { trs: out, headerRowEmitted };
    }

    private async processNodeArray(nodes: OfficeContentNode[]): Promise<string> {
        // Whether these siblings form a run of text (a paragraph's children) rather than a list of blocks.
        const runHasText = nodes.some(n => n.type === 'text');
        let html = '';
        // Stack to track active lists. `liClose` is the currently-open item's deferred closing
        // suffix (`</li>`, or `</div></li>` for a task item): a list item is rendered WITHOUT its
        // close so a deeper list can land inside it (spec-valid `<li>a<ul>...</ul></li>` rather
        // than the invalid `<li>a</li><ul>...</ul>` sibling shape). The close is emitted when a
        // same-level sibling arrives, when the level is popped, or at the end.
        const listStack: { indentation: number, type: 'ordered' | 'unordered', isTask: boolean, liClose: string }[] = [];

        const openListTag = (type: 'ordered' | 'unordered', isTask: boolean) => {
            if (isTask) return '<ul data-type="taskList">';
            return type === 'ordered' ? '<ol>' : '<ul>';
        };
        const closeListTag = (type: 'ordered' | 'unordered') => type === 'ordered' ? '</ol>' : '</ul>';

        const closeListsToLevel = (level: number) => {
            while (listStack.length > 0 && listStack[listStack.length - 1].indentation > level) {
                const list = listStack.pop();
                html += list!.liClose + closeListTag(list!.type) + '\n\n';
            }
        };

        for (const node of nodes) {
            // Check if node should be filtered out or overridden
            const override = await this.handleOnNode(node);
            if (override === false) {
                // Remembered so a later pass (the spreadsheet tab bar) can honour the same verdict
                // without asking the hook again.
                this.onNodeSkipped.add(node);
                continue;
            }

            // A top-level footnote/endnote note is an orphan definition (the MarkdownParser
            // recovers unreferenced `[^id]: ...` defs as trailing note nodes). Route it into the
            // collected footnotes so it renders inside `<section data-footnotes>` - where HtmlParser
            // reads it back on import - instead of inline outside the section with a dead back-link.
            const orphanMeta = node.metadata as NoteMetadata;
            const orphanNoteType = orphanMeta?.noteType;
            if (node.type === 'note' && orphanMeta?.unreferenced && (orphanNoteType === 'footnote' || orphanNoteType === 'endnote')) {
                // Only an unreferenced (orphan) definition is hoisted into <section data-footnotes>.
                // The `unreferenced` guard also keeps the generators in agreement at depth: this
                // routing runs in every processNodeArray call, so without it a `note` sitting as a
                // CHILD of a container (a consumer-built AST; no shipped parser emits this) would be
                // hoisted here and then given an `<a href="#footnote-ref-N">` back-link with no
                // anchor - the exact dangling link the orphan handling removes. MarkdownGenerator's
                // equivalent routing is top-level only, so gating on the flag matches it.
                this.collectedNotes.push(node);
                // Remember the hook's verdict (already asked above) so the footnotes-section render
                // reuses it instead of firing onNode a second time for this note.
                this.collectedNoteOverrides.set(node, { value: override });
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
                    const top = listStack[listStack.length - 1];
                    if (top.type !== type || top.isTask !== isTask) {
                        // Kind changed at the same level: close the open item and the old list,
                        // then open the replacement list.
                        const last = listStack.pop();
                        html += last!.liClose + closeListTag(last!.type) + '\n';
                        html += openListTag(type, isTask) + '\n';
                        listStack.push({ indentation, type, isTask, liClose: '' });
                    } else {
                        // Sibling at the same level: close the previous item before this one opens.
                        html += top.liClose;
                    }
                } else {
                    // Deeper level (or the first list): open a nested list INSIDE the currently
                    // open item, leaving the parent <li>'s close pending on its stack frame.
                    html += openListTag(type, isTask) + '\n';
                    listStack.push({ indentation, type, isTask, liClose: '' });
                }

                html += await this.processNodeRecursive(node, this.boundNodeProcessor, { value: override });
                // Defer this item's close so a nested list can land inside it. A string override is
                // a complete replacement item that already carries its own close, so add none.
                listStack[listStack.length - 1].liClose = (typeof override === 'string')
                    ? ''
                    : (isTask ? '</div></li>' : '</li>');
            } else {
                // Non-list node closes all active lists
                closeListsToLevel(-1);
                let result = await this.processNodeRecursive(node, this.boundNodeProcessor, { value: override });

                // Add a blank line after BLOCK nodes for readable HTML source. Inline nodes (a
                // paragraph's text/link runs) must concatenate with no separator: adding `\n\n`
                // around an inline <a> put a blank line inside the <p>, which reparsed as a stray
                // space before the following punctuation (`[video](url) .`). A source comment inside a
                // run is inline too: a blank line after it would part the words around it, and so is
                // anything in a paragraph's or heading's line (a picture, inline math) but a line
                // break, after which the line starts afresh.
                const inlineNode = node.type === 'text' || (isSourceComment(node) && runHasText) || (this.inlineDepth > 0 && node.type !== 'break');
                if (!inlineNode && !result.endsWith('\n\n')) {
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
     * The first-row-is-a-header heuristic: an explicit "header" row style, or an all-bold first row.
     * Read both when deciding how to render the table and (via {@link rendersOwnChildren}) before the
     * generic children pass, so the two cannot disagree about which shape the table takes.
     */
    private firstRowIsHeader(node: OfficeContentNode): boolean {
        const rows = node.children || [];
        if (!rows.length || rows[0].type !== 'row') return false;
        // Use the heuristic every other generator shares, so a header row survives into <thead> the
        // same way it does in DOCX/ODT/native-PDF output. It adds the cell-level `style: 'header'` /
        // `isHeader` convention that the Word, ODF and tagged-PDF parsers write (a repeating header row
        // that is not bold), on top of the row-level style and all-bold-first-row cases this had.
        return isHeaderRow(rows[0], true);
    }

    /** True when any cell of the table spans more than one row. */
    private tableHasRowSpan(node: OfficeContentNode): boolean {
        return (node.children || []).some(r => r.type === 'row' &&
            (r.children || []).some(c => c.type === 'cell' && ((c.metadata as CellMetadata)?.rowSpan || 1) > 1));
    }

    /**
     * True when this node's own branch in {@link nodeProcessor} lays its children out itself and so
     * discards the generic children pass entirely: a `table` that re-renders its rows (a rowspan grid,
     * or a first row promoted into `<thead>`), a `sheet` (which rebuilds the whole grid from the cell
     * indices), or a sparse `row` (whose cells carry explicit column indices). Checked before that
     * pass runs - rendering the subtree twice was pure waste, and it fired the caller's `onNode` hook
     * a second time for every node inside it.
     *
     * Each condition mirrors the branch it guards; they must stay in step, so both read the same
     * helpers rather than restating the heuristic.
     */
    private rendersOwnChildren(node: OfficeContentNode): boolean {
        if (node.type === 'sheet') return true;
        if (node.type === 'table') return this.tableHasRowSpan(node) || this.firstRowIsHeader(node);
        if (node.type === 'row') {
            const cells = (node.children || []).filter(c => c.type === 'cell');
            return cells.length > 0 && cells.some(c => (c.metadata as CellMetadata)?.col !== undefined);
        }
        return false;
    }

    /** Nodes whose `onNode` verdict was `false`, so a later pass can honour it without re-asking. */
    private readonly onNodeSkipped = new WeakSet<OfficeContentNode>();
    /** Verdicts for orphan notes hoisted into the footnotes section, so onNode is not asked again when
     *  they are rendered there (it was already asked in processNodeArray, firing the hook twice). */
    private readonly collectedNoteOverrides = new Map<OfficeContentNode, OnNodeVerdict>();

    /**
     * The default node processor, bound once so a recursion can tell it apart from a caller-supplied
     * one by identity (see {@link rendersOwnChildren}) - and so the binding is not re-allocated at
     * every call site.
     */
    private readonly boundNodeProcessor = this.nodeProcessor.bind(this);

    /**
     * Overridden to handle children using processNodeArray for list grouping.
     */
    private tableNestingLevel = 0;
    /** How many paragraphs or headings are being written: a picture in one is written inline. */
    private inlineDepth = 0;

    protected override async processNodeRecursive(
        node: OfficeContentNode,
        processor: (node: OfficeContentNode, childrenOutput: string) => string | Promise<string>,
        override?: OnNodeVerdict
    ): Promise<string> {
        // Mirrors the check in BaseGenerator.processNodeRecursive. This override replaces that
        // method entirely, so without repeating the check here the signal would be silently
        // inert for this generator - which is exactly how it was missed.
        checkAbortSignal(this.config.abortSignal);
        const wasInHeading = this.inHeading;
        const wasHeadingSize = this.headingUniformSize;
        if (node.type === 'heading') {
            this.inHeading = this.hasUniformFormatting(node, f => f?.bold === true);
            this.headingUniformSize = this.hasUniformFormatting(node, f => !!f?.size);
        }
        try {
            return await this.processNodeRecursiveInner(node, processor, override);
        } finally {
            this.inHeading = wasInHeading;
            this.headingUniformSize = wasHeadingSize;
        }
    }

    private async processNodeRecursiveInner(
        node: OfficeContentNode,
        processor: (node: OfficeContentNode, childrenOutput: string) => string | Promise<string>,
        override?: OnNodeVerdict
    ): Promise<string> {
        // Use the pre-evaluated verdict when the caller already asked the hook, otherwise ask it here.
        // The verdict is boxed because "no override" IS `undefined`: comparing the bare value against
        // undefined could not tell an answered hook from an unasked one, so every node processed
        // through processNodeArray fired `onNode` a second time here.
        const actualOverride = override ? override.value : await this.handleOnNode(node);

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

        // A paragraph holding a block (a code block's <pre>, a display equation's <div>) is written as
        // its parts, since HTML cannot nest a block in a <p>: a DOM parser closes the paragraph at it,
        // and XHTML (EPUB's) rejects it. Each run of inline content is a paragraph, each block stands
        // on its own, and the paragraph's id and anchors go on its first part.
        const splitsAroundBlocks = node.type === 'paragraph' && processor === this.boundNodeProcessor && !!node.children?.some(isBlockInParagraph);
        // A paragraph or heading holds phrasing content only: a picture in one is written inline.
        const holdsInline = node.type === 'paragraph' || node.type === 'heading';
        let splitResult: string | undefined;
        let childrenOutput = '';
        if (holdsInline) this.inlineDepth++;
        try {
        if (splitsAroundBlocks) {
            splitResult = '';
            let run: OfficeContentNode[] = [];
            let part: OfficeContentNode = node;
            const flush = async () => {
                if (run.some(child => child.type !== 'text' || (child.text ?? '').trim() || child.notes?.length)) {
                    splitResult += await processor({ ...part, children: run }, await this.processNodeArray(run));
                    part = { ...node, metadata: { ...(node.metadata as object), anchorIds: undefined } as any };
                }
                run = [];
            };
            for (const child of node.children!) {
                if (isBlockInParagraph(child)) {
                    await flush();
                    // The block stands on its own, outside the paragraph's line of text.
                    this.inlineDepth--;
                    try {
                        splitResult += await this.processNodeArray([child]);
                    } finally {
                        this.inlineDepth++;
                    }
                } else {
                    run.push(child);
                }
            }
            await flush();
        }

        if (!splitsAroundBlocks && node.children && node.children.length > 0) {
            // A node whose own branch lays its children out throws this output away, so skip producing
            // it: the subtree would otherwise be walked twice, firing `onNode` twice for every node in
            // it. Only when the DEFAULT processor is rendering this node, though - a caller that passes
            // its own processor (the header-row promotion) consumes the children output itself.
            if (processor !== this.boundNodeProcessor || !this.rendersOwnChildren(node)) {
                childrenOutput = await this.processNodeArray(node.children);
            }
        } else if (node.text && node.type !== 'text') {
            // Fallback for nodes that have text property but no children (e.g. simple paragraphs)
            childrenOutput = this.escape(node.text);
        }
        } finally {
            if (holdsInline) this.inlineDepth--;
        }

        if (node.notes && node.notes.length > 0) {
            if (node.type !== 'slide') {
                this.collectedNotes.push(...node.notes);
            }
        }

        let result = splitResult ?? await processor(node, childrenOutput);

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

        // A styleMap's `output.tag` was previously written here and then shadowed by a `const tag`
        // in every switch branch, so HtmlGenerator silently ignored it - while MarkdownGenerator
        // and RtfGenerator both honoured it and the README documented it as working. Honoured now,
        // but only through the allowlist: the tag is interpolated into both `<TAG>` and `</TAG>`,
        // and the shadowing bug was the sole reason a hostile value could not inject. A rejected
        // tag falls back to the default rather than being emitted.
        const mappedTag = isSafeStyleMapTag(mapping?.tag) ? mapping!.tag.toLowerCase() : undefined;
        if (mapping?.tag && !mappedTag) {
            this.warn(OfficeWarningType.INVALID_STYLE_MAP_TAG, mapping.tag, node);
        }

        // Handle Attributes from mapping
        let mappedAttrs = '';
        if (mapping?.attributes) {
            for (const [key, val] of Object.entries(mapping.attributes)) {
                // The value is escaped, but the NAME needs validating too: a key containing a
                // quote closes the attribute and opens another, so `x" onmouseover="alert(1)" z`
                // yields a live event handler however carefully the value is escaped. The
                // attribute bag and the parser already apply this policy; this path did not.
                if (!isSafeHtmlAttributeName(key)) continue;
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

        // An empty id is none (`id=""` names nothing a link can reach).
        const anchorIds: string[] = this.config.ignoreInternalLinks ? [] : ((node.metadata as any)?.anchorIds || []).filter((id: string) => !!id);

        if (this.config.generateIds) {
            if (node.type === 'heading') {
                // From the heading's text, whether the node carries it or only its runs do. A heading
                // whose text slugifies to nothing (no Latin letters or digits) gets no generated id,
                // rather than an empty one.
                const slug = this.slugify(node.text || this.getNodeText(node));
                if (slug && !anchorIds.includes(slug)) anchorIds.push(slug);
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
                const mode = this.imageMode();
                if (mode === 'none') return '';
                const meta = node.metadata as ImageMetadata;
                const attachmentName = meta?.attachmentName;
                const ocr = (node.text || '').trim();

                // ocr-text-only: emit the recognized text as a visible block, no <img>. A <pre>
                // preserves the 2-D column layout the OCR reconstruction encodes with spaces. Carry the
                // node's mapped classes (merged with the intrinsic `ocr-text` class), preserved
                // attributes and inline style through too, so this mode is not the one image emission
                // site that silently drops them.
                if (mode === 'ocr-text-only') {
                    if (!ocr) return '';
                    const preClass = ` class="${this.escape(['ocr-text', ...classes].join(' '))}"`;
                    return `${extraAnchors}<pre${preClass}${mappedAttrs}${idAttr}${styleAttr}>${this.escape(ocr)}</pre>`;
                }

                let src = meta?.url || attachmentName || '';
                if (!meta?.url && attachmentName && this.ast) {
                    const attachment = this.getAttachment(attachmentName);
                    if (attachment) {
                        const bytes = base64ByteLength(attachment.data);
                        // A self-contained (standalone) HTML document must embed its images: a name
                        // reference there is a broken image with no packager to resolve it. So the size
                        // cap (which exists to keep a huge base64 line out of Markdown, or out of a
                        // fragment a consumer post-processes) is lifted for a standalone document.
                        const cap = this.emitsStandaloneDocument ? Infinity : this.config.maxInlineImageBytes;
                        if (bytes <= cap) {
                            src = `data:${attachment.mimeType || 'image/png'};base64,${attachment.data}`;
                        } else {
                            // Fragment over the cap: keep the name reference the consumer resolves, but
                            // surface it so a large image degrading to a bare src is never silent.
                            this.warn(OfficeWarningType.IMAGE_NOT_INLINED, { name: attachmentName, bytes, limit: this.config.maxInlineImageBytes });
                        }
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
                } else if (node.bounds && node.bounds.width > 0) {
                    // PDF images carry their on-page size in points; use it so a high-DPI scanned page
                    // does not render at its intrinsic pixel size. Only PDF-sourced nodes have bounds,
                    // so images from other formats are unaffected.
                    imgStyleParts.push(`width: ${Math.round(node.bounds.width)}pt`, 'max-width: 100%');
                }
                if (meta?.align) {
                    imgDataAttrs += ` data-align="${this.escape(meta.align)}"`;
                    const ml = meta.align === 'left' ? '0' : 'auto';
                    const mr = meta.align === 'right' ? '0' : 'auto';
                    imgStyleParts.push('display: block', `margin-left: ${ml}`, `margin-right: ${mr}`);
                }
                const imgStyleAttr = imgStyleParts.length > 0 ? ` style="${imgStyleParts.join('; ')}"` : '';

                const imgTitle = meta?.title ? ` title="${this.escape(meta.title)}"` : '';
                // alt is the descriptive alt text, not the OCR text: OCR text is surfaced visibly under
                // 'image+ocr-text' rather than hidden in alt (where a broken/referenced image would leak
                // it into the rendered page).
                const imgWithId = (id: string): string => {
                    const tag = `<img src="${sanitizeImageUrl(src)}" alt="${this.escape(meta?.altText || '')}"${id}${imgTitle}${className}${mappedAttrs}${imgDataAttrs}${imgStyleAttr}>`;
                    // A picture that is a link (a badge) is wrapped in it, as a linked run is.
                    if (!meta?.link || (this.config.ignoreInternalLinks && meta.linkType !== 'external')) return tag;
                    const linkTitle = meta.linkTitle ? ` title="${this.escape(meta.linkTitle)}"` : '';
                    return `<a href="${sanitizeUrl(meta.link)}"${linkTitle}${meta.linkType === 'external' ? ' target="_blank"' : ''}>${tag}</a>`;
                };
                // In a paragraph or heading (phrasing content only) the picture is written inline, its
                // id on the <img>: the block markup below inside a <p> is split apart by every browser
                // and parser.
                if (this.inlineDepth > 0) {
                    const ocrText = mode === 'image+ocr-text' && ocr ? `<br><span class="ocr-text">${this.escape(ocr)}</span>` : '';
                    return `${extraAnchors}${imgWithId(idAttr)}${ocrText}`;
                }
                // Among the blocks too its id is on the <img>, after the anchors of its other ids, as they
                // are read back: on the wrapper it was read after them, and the order flipped each save.
                const img = imgWithId(idAttr);
                let content = this.config.includeFormatting ? `<div class="image-container">${img}<div class="caption">${this.escape(attachmentName || '')}</div></div>` : img;
                // image+ocr-text: the image, then its recognized text (a <pre> keeps the 2-D layout).
                if (mode === 'image+ocr-text' && ocr) content += `<pre class="ocr-text">${this.escape(ocr)}</pre>`;
                return `${extraAnchors}<div>${content}</div>`;
            }

            case 'chart': {
                if (!this.config.includeCharts) return '';
                const meta = node.metadata as any; // ChartMetadata
                this.chartCounter++;
                const chartId = `chart-${this.chartCounter}`;
                const chartAttName = meta?.attachmentName;
                const chartAttachment = this.getAttachment(chartAttName);

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

            case 'break': {
                const breakType = (node.metadata as any)?.breakType;
                if (breakType === 'page') return `${extraAnchors}<hr class="page-break"${idAttr}>`;
                // A thematic break is a plain rule; the parser reads a bare <hr> back as one.
                if (breakType === 'thematic') return `${extraAnchors}<hr${idAttr}>`;
                return '<br>';
            }

            case 'code': {
                const meta = node.metadata as CodeMetadata;
                if (meta?.math) {
                    const tag = meta.math === 'block' ? 'div' : 'span';
                    if (this.config.htmlConfig.sourceAttributes) {
                        // Attribute-driven emission: the raw (undelimited) LaTeX lives in both
                        // data-math and the text content, and the class token carries the mode.
                        // The widened HtmlParser reads this back (a data-math value other than
                        // inline/block is treated as LaTeX).
                        const latex = node.text || '';
                        return `${extraAnchors}<${tag} class="math math-${this.escape(meta.math)}" data-math="${this.escape(latex)}"${idAttr}${mappedAttrs}${styleAttr}>${this.escape(latex)}</${tag}>`;
                    }
                    // Default emission: data-math names the display mode, and the visible text
                    // keeps its $ delimiters so the raw LaTeX degrades gracefully without a
                    // KaTeX renderer.
                    const delimited = meta.math === 'block' ? `$$${node.text || ''}$$` : `$${node.text || ''}$`;
                    return `${extraAnchors}<${tag} class="math math-${this.escape(meta.math)}" data-math="${this.escape(meta.math)}"${idAttr}${mappedAttrs}${styleAttr}>${this.escape(delimited)}</${tag}>`;
                }
                if (this.config.htmlConfig.sourceAttributes && meta?.language === 'mermaid') {
                    // Attribute-driven emission: a <div class="mermaid" data-mermaid> the widened
                    // parser maps back to a mermaid code node. escape() encodes '>' so diagram
                    // arrows (-->), plus newlines and quotes, stay inside the tag and attribute.
                    const code = node.text || '';
                    return `${extraAnchors}<div class="mermaid" data-mermaid="${this.escape(code)}"${idAttr}${mappedAttrs}${styleAttr}>${this.escape(code)}</div>`;
                }
                const lang = meta?.language ? ` class="language-${this.escape(meta.language)}"` : '';
                const codeHtml = `<code${lang}>${this.escape(node.text || '')}</code>`;
                // A `code` node is always block-level (inline code is a monospace text run, emitted
                // as <code> by formatText), so it is a <pre>, whatever its length and whether or not
                // it names a language. A one-line block without one used to be a <span><code>, which
                // re-imports as inline code and which strict CodeBlock parsers (only <pre><code>) miss.
                return `${extraAnchors}<pre${idAttr}${className}${mappedAttrs}${styleAttr}>${codeHtml}</pre>`;
            }

            case 'list': {
                // The closing suffix (`</div></li>` for a task item, `</li>` otherwise) is emitted
                // by processNodeArray's list stack, not here, so a nested list can be placed inside
                // this item before it closes. See `listStack`/`liClose` there.
                const meta = node.metadata as ListMetadata;
                if (meta?.isTask) {
                    const checkedAttr = ` data-checked="${meta.checked ? 'true' : 'false'}"`;
                    const checkedBool = meta.checked ? ' checked' : '';
                    return `${extraAnchors}<li${checkedAttr}${idAttr}${className}${mappedAttrs}${styleAttr}><label><input type="checkbox"${checkedBool}><span></span></label><div>${childrenOutput}`;
                }
                const value = (meta?.listType === 'ordered' && typeof meta.itemIndex === 'number')
                    ? ` value="${meta.itemIndex + 1}"`
                    : '';
                return `${extraAnchors}<li${value}${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}`;
            }

            case 'table': {
                // Smart Table Header Detection
                let finalChildren = childrenOutput;
                const rows = node.children || [];
                const firstRowIsHeader = this.firstRowIsHeader(node);
                // When any cell spans rows, the plain per-row path (which only fills horizontal gaps)
                // would shift the cells sitting under a rowspan. Render with an HTML grid-occupancy
                // model instead, so a rowspan reserves its column in the rows below. Scoped to tables
                // that actually contain a rowspan, so every other table renders exactly as before.
                const hasRowSpan = this.tableHasRowSpan(node);
                if (hasRowSpan) {
                    const rowNodes = rows.filter(r => r.type === 'row');
                    const { trs, headerRowEmitted } = await this.renderRowsWithRowspans(rowNodes, firstRowIsHeader);
                    // A rowspan started in the header row cannot cross an HTML <thead>/<tbody>
                    // boundary, so only split off a <thead> when the header row has no downward span.
                    const firstRowSpansDown = (rowNodes[0]?.children || []).some(c => c.type === 'cell' && ((c.metadata as CellMetadata)?.rowSpan || 1) > 1);
                    // Only wrap trs[0] in <thead> when the header row actually survived onNode; if the
                    // hook dropped it, trs[0] is a body row and belongs in <tbody>, not <thead>.
                    finalChildren = (firstRowIsHeader && headerRowEmitted && !firstRowSpansDown && trs.length)
                        ? `<thead>${trs[0]}</thead><tbody>${trs.slice(1).join('')}</tbody>`
                        : trs.join('');
                } else if (rows.length > 0 && rows[0].type === 'row') {
                    const firstRow = rows[0];
                    if (firstRowIsHeader) {
                        // Promote the first row to header cells wrapped in a <tr>. The <tr> is required:
                        // header cells directly under <thead> (`<thead><th>…`) are invalid HTML that
                        // HtmlParser does not read back as a row, so a md -> HTML -> md round trip lost
                        // the header. onNode is asked once for the row and once per cell (matching the
                        // default path); a dropped/overridden row is honoured.
                        const rowVerdict = await this.handleOnNode(firstRow);
                        if (rowVerdict === false) {
                            // Header row dropped by the hook: render the remaining rows as a plain body,
                            // no <thead>. Use rows.slice(1) (not a filter that re-includes firstRow), so
                            // the already-answered header row is not walked - and onNode not asked - twice.
                            finalChildren = await this.processNodeArray(rows.slice(1));
                        } else {
                            let headInner: string;
                            if (typeof rowVerdict === 'string') {
                                headInner = rowVerdict;
                            } else {
                                // Promote each TOP-LEVEL cell's own <td> to <th> with the anchored replace
                                // the rowspan path uses (first `<td`, last `</td>`), so a table nested
                                // inside a header cell keeps its own <td> body cells as <td>.
                                let tr = '';
                                for (const cell of (firstRow.children || []).filter(c => c.type === 'cell')) {
                                    let cellHtml = await this.processNodeRecursive(cell, this.boundNodeProcessor);
                                    cellHtml = cellHtml.replace(/<td/, '<th').replace(/<\/td>$/, '</th>');
                                    tr += cellHtml;
                                }
                                headInner = `<tr>${tr}</tr>`;
                            }
                            const bodyOutput = await this.processNodeArray(rows.slice(1));
                            finalChildren = `<thead>${headInner}</thead><tbody>${bodyOutput}</tbody>`;
                        }
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

                            sparseChildren += await this.processNodeRecursive(cell, this.boundNodeProcessor);

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
                // Emit extraAnchors (the 2nd+ anchor ids of a cell with several bookmarks) like every
                // other node, so they are not silently dropped. They precede the <td>, so the header
                // promotion below matches the first <td> rather than the string start.
                return `${extraAnchors}<td${rowSpan}${colSpan}${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</td>`;
            }

            case 'sheet': {
                const rows = node.children || [];

                // Each cell's place in the grid: the row and column a spreadsheet parser gives it, else
                // (a sheet built by hand) the next free place in reading order, past the cells a merge
                // above covers. A cell with no place was left out of the grid, and with it the sheet.
                const places = new Map<OfficeContentNode, { r: number; c: number }>();
                const coveredThrough: number[] = [];
                let nextRow = 0;
                for (const child of rows) {
                    if (child.type !== 'row') continue;
                    const cellsInRow = (child.children || []).filter(c => c.type === 'cell');
                    const firstRow = (cellsInRow[0]?.metadata as CellMetadata | undefined)?.row;
                    const rowIndex = typeof firstRow === 'number' ? firstRow : nextRow;
                    let col = 0;
                    for (const cell of cellsInRow) {
                        const meta = cell.metadata as CellMetadata | undefined;
                        const r = typeof meta?.row === 'number' ? meta.row : rowIndex;
                        let c = typeof meta?.col === 'number' ? meta.col : col;
                        if (typeof meta?.col !== 'number') while ((coveredThrough[c] ?? -1) >= r) c++;
                        places.set(cell, { r, c });
                        const cSpan = meta?.colSpan || 1;
                        if ((meta?.rowSpan || 1) > 1) for (let k = c; k < c + cSpan; k++) coveredThrough[k] = Math.max(coveredThrough[k] ?? -1, r + (meta?.rowSpan || 1) - 1);
                        col = c + cSpan;
                    }
                    nextRow = rowIndex + 1;
                }

                // Find grid bounds
                let maxRow = -1;
                let maxCol = -1;
                for (const [cell, { r, c }] of places) {
                    const meta = cell.metadata as CellMetadata | undefined;
                    const rSpan = meta?.rowSpan || 1;
                    const cSpan = meta?.colSpan || 1;
                    if (r + rSpan - 1 > maxRow) maxRow = r + rSpan - 1;
                    if (c + cSpan - 1 > maxCol) maxCol = c + cSpan - 1;
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
                                rowNodeMap.set(places.get(cellsInRow[0])!.r, child);

                                for (const cell of cellsInRow) {
                                    const meta = cell.metadata as CellMetadata | undefined;
                                    const { r, c } = places.get(cell)!;
                                    if (r >= 0 && r <= maxRow && c >= 0 && c <= maxCol) {
                                        grid[r][c] = cell;

                                        const rSpan = meta?.rowSpan || 1;
                                        const cSpan = meta?.colSpan || 1;
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
                            // This branch assembles each <tr> from the grid rather than recursing into
                            // the row node, so the row is offered to `onNode` here - otherwise it would
                            // be the one node type the hook never saw in a spreadsheet.
                            const rowOverride = await this.handleOnNode(rowNode);
                            if (rowOverride === false) continue;
                            if (typeof rowOverride === 'string') { tbodyRows += rowOverride; continue; }
                            const mapping = this.getSemanticMapping(rowNode);
                            const rClasses = ['excel-row'];
                            if (mapping?.classes) rClasses.push(...mapping.classes);
                            // Escaped like the `className` built for every other node type. This
                            // path rebuilds the class attribute from the raw mapping array rather
                            // than reusing that value, and was the only place it went out unescaped.
                            trAttrs += ` class="${this.escape(rClasses.join(' '))}"`;
                            if (mapping?.attributes) {
                                for (const [key, val] of Object.entries(mapping.attributes)) {
                                    if (!isSafeHtmlAttributeName(key)) continue;
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
                                const cellHtml = await this.processNodeRecursive(cell, this.boundNodeProcessor);
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
                // Escaped for the same reason as the row above; `classes` here also carries the
                // attribute bag's raw className, so escaping at the join covers both sources.
                const classAttr = ` class="${this.escape(mergedClasses.join(' '))}"`;

                // The sheet's element takes the id its tab links to; its own ids are anchors before it
                // (the first was dropped for that id, which the extra anchors then repeated).
                const finalIdAttr = ` id="${sheetId}"`;
                const sheetAnchors = anchorIds.filter(aid => aid !== sheetId).map(aid => `<a id="${this.escape(aid)}" name="${this.escape(aid)}"></a>`).join('');

                return `${sheetAnchors}<div${finalIdAttr}${classAttr}${mappedAttrs}${styleAttr}>${tableHtml}${nonRowHtml}</div>`;
            }

            case 'paragraph':
            case 'heading': {
                // The styleMap tag wins over the structural default; that is the whole point of
                // mapping "Heading 1"/"Intense Quote" onto a semantic element.
                const tag = mappedTag ?? (node.type === 'heading' ? `h${(node.metadata as HeadingMetadata)?.level || 1}` : 'p');

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
                // Emit an id of `page=N` so internal links from parsed PDFs (`href="#page=N"`) resolve
                // in the generated HTML and printed PDF. If the section already has an id, add a
                // separate leading anchor instead of overwriting it. Only when there is a real page
                // number: an empty `page=` id on every unnumbered page would be a duplicate (invalid HTML).
                const pageAnchor = (pageNum && idAttr) ? `<a id="page=${pageNum}"></a>` : '';
                const pageIdAttr = idAttr || (pageNum ? ` id="page=${pageNum}"` : '');
                return `${extraAnchors}${pageAnchor}<section class="page" data-page-num="${pageNum}"${pageIdAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</section>`;
            }
            case 'note': {
                const meta = node.metadata as NoteMetadata;
                if (meta?.noteType === 'footnote' || meta?.noteType === 'endnote') {
                    const key = this.escape(this.getFootnoteKey(node));
                    // A <div> wrapper, not <p>: childrenOutput is block content, so a <p> wrapper
                    // nests <p> inside <p> (every DOM parser splits it, leaving the wrapper empty),
                    // and attribute-driven editors match `div[data-footnote-id]`. This changes the
                    // default footnote-definition markup, which was broken-by-construction before.
                    // An unreferenced (orphan) note has no citation anchor, so the back-link would
                    // dangle - omit it.
                    const backLink = meta?.unreferenced ? '' : ` <a href="#footnote-ref-${key}">↩</a>`;
                    // The note's own ids (a bookmark in it) start it, where the parser gives them back.
                    const noteAnchors = anchorIds.map(aid => `<a id="${this.escape(aid)}" name="${this.escape(aid)}"></a>`).join('');
                    return `<div id="footnote-${key}" data-footnote-id="${key}">${noteAnchors}${childrenOutput}${backLink}</div>`;
                }
                const noteClass = meta?.noteType ? ` note-${this.escape(meta.noteType)}` : '';
                return `${extraAnchors}<div class="slide-note${noteClass}"${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</div>`;
            }

            case 'embed': {
                const meta = node.metadata as EmbedMetadata;
                // An embed naming no type (built by hand), or a YouTube one by its URL alone, is what
                // it carries: its URL was lost in an empty YouTube wrapper.
                const embed = resolveEmbed(meta);
                if (embed?.kind === 'iframe') {
                    // Generic preserved iframe. sanitizeUrl scheme-checks the src (only http/https
                    // and the other non-executing schemes survive), so a javascript:/data: src is
                    // dropped even with preservation on. The node only exists via opt-in parsing or
                    // a programmatic AST, so this guard is unconditional.
                    const src = sanitizeUrl(embed.url);
                    if (!src) return '';
                    const w = meta?.width ? ` width="${this.escape(meta.width)}"` : '';
                    const h = meta?.height ? ` height="${this.escape(meta.height)}"` : '';
                    if (this.config.htmlConfig.gatedEmbeds) {
                        // Inert, never-auto-loading placeholder: the editor renders click-to-load from
                        // it, and HtmlParser reads it back to the same embed node. src already sanitized.
                        const a = meta?.align ? ` data-embed-align="${this.escape(meta.align)}"` : '';
                        const l = meta?.label ? ` data-embed-label="${this.escape(meta.label)}"` : '';
                        const dw = meta?.width ? ` data-embed-width="${this.escape(meta.width)}"` : '';
                        const dh = meta?.height ? ` data-embed-height="${this.escape(meta.height)}"` : '';
                        return `${extraAnchors}<div data-embed-gated data-embed-src="${src}"${dw}${dh}${a}${l}${idAttr}${mappedAttrs}${styleAttr}></div>`;
                    }
                    return `${extraAnchors}<iframe src="${src}"${w}${h}${idAttr}${mappedAttrs}${styleAttr}></iframe>`;
                }
                // Match the attribute-driven Youtube wrapper shape so a loaded embed re-hydrates
                // an editor's Youtube node.
                const id = embed?.kind === 'youtube' ? embed.videoId : '';
                const width = meta?.width || '100%';
                const align = meta?.align || 'center';
                const ml = align === 'left' ? '0' : 'auto';
                const mr = align === 'right' ? '0' : 'auto';
                const iframe = id
                    ? `<iframe src="https://www.youtube.com/embed/${this.escape(id)}" title="YouTube video player" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen></iframe>`
                    : '';
                // Carry the human label so it survives AST -> editor-HTML -> AST, at parity with the
                // generic gated path's data-embed-label. Only emitted when a label exists (from a
                // `::youtube[Label]` directive), so unlabeled youtube embeds are byte-identical.
                const ytLabel = meta?.label ? ` data-embed-label="${this.escape(meta.label)}"` : '';
                return `${extraAnchors}<div data-youtube-video="${this.escape(id)}" data-width="${this.escape(width)}" data-align="${this.escape(align)}"${ytLabel} class="youtube-embed"${idAttr}${mappedAttrs} style="width: ${sanitizeCssValue(width)}; margin-left: ${ml}; margin-right: ${mr};">${iframe}</div>`;
            }

            case 'admonition': {
                // Match the attribute-driven admonition wrapper so a loaded admonition
                // reaches an editor as that node instead of a plain blockquote.
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

            // Rendered through their children only. Listed explicitly, with no `default`, so that
            // under `noImplicitReturns` a new OfficeContentNodeType fails to compile until it is
            // classified here rather than silently degrading to its children.
            // A source comment (`<!-- ... -->`) stays a hidden note: a real comment, or - under
            // `sourceAttributes`, for an editor whose DOM parser discards comment nodes - an empty span
            // carrying the raw text in an escaped attribute, so it is data and can never become markup.
            // A review comment keeps its existing rendering (its children).
            case 'comment':
                if (isSourceComment(node)) {
                    return this.config.htmlConfig.sourceAttributes
                        ? `<span data-html-comment="${escapeHtml(node.text || '')}"></span>`
                        : `<!--${sanitizeCommentText(node.text || '')}-->`;
                }
                return childrenOutput;

            case 'drawing':
            case 'header':
            case 'footer':
            case 'slideMaster':
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
            // Every other type is a generic block. No `default`, so a new node type must be placed.
            case 'text':
            case 'table':
            case 'image':
            case 'chart':
            case 'drawing':
            case 'slide':
            case 'note':
            case 'sheet':
            case 'row':
            case 'cell':
            case 'page':
            case 'break':
            case 'code':
            case 'comment':
            case 'header':
            case 'footer':
            case 'slideMaster':
            case 'embed':
            case 'admonition':
            case 'definitionList':
            case 'definitionTerm':
            case 'definitionDescription':
                return 'div';
        }
    }

    private formatText(node: OfficeContentNode, text: string): string {
        let result = this.escape(text);
        const f = node.formatting;

        if (this.config.includeFormatting && f) {
            // Inline code: a monospace run becomes `<code>`, not a `font-family: monospace` span, so
            // an editor keying on <code> sees it and it re-imports as inline code (HtmlParser maps
            // <code> back to a monospace run). Innermost, so bold/italic wrap it (`<b><code>…`).
            if (f.font === 'monospace') result = `<code>${result}</code>`;
            // Inside an `<hN>`, the heading's own styling is authoritative. A run that also carries
            // bold and a font size - the normal case for ODF, where a heading's paragraph style is
            // inherited by its runs - would wrap the text in `<b>` the heading already implies and,
            // worse, in a `<span style="font-size: 14pt">` that *shrinks* the heading to the size
            // its paragraph style happened to name. See the same suppression in RtfGenerator.
            if (f.bold && !this.inHeading) result = `<b>${result}</b>`;
            if (f.italic) result = `<i>${result}</i>`;
            if (f.underline) result = `<u>${result}</u>`;
            if (f.strikethrough) result = `<strike>${result}</strike>`;
            if (f.subscript) result = `<sub>${result}</sub>`;
            if (f.superscript) result = `<sup>${result}</sup>`;

            const styles = this.headingUniformSize
                ? this.getInlineStyles(node, { skipFontSize: true, skipBackgroundColor: true })
                : this.getInlineStyles(node, { skipBackgroundColor: true });
            if (styles) {
                result = `<span style="${styles}">${result}</span>`;
            }
            // Highlight -> <mark>, not a <span style="background-color">. Tiptap's Highlight
            // extension parseHTML matches exactly `mark`, so an editor round trip only rehydrates
            // the highlight from a <mark>; a bare span comes back as unhighlighted text. Kept
            // outside the colour/size span above so a run carrying both still rehydrates both. The
            // widened HtmlParser reads this shape back (style wins over data-color). Behaviour
            // change (was a span), noted in the changelog.
            if (f.backgroundColor) {
                const safeBg = sanitizeCssValue(f.backgroundColor);
                if (safeBg) {
                    result = `<mark data-color="${this.escape(safeBg)}" style="background-color: ${safeBg}">${result}</mark>`;
                }
            }
        }

        const meta = node.metadata as TextMetadata;
        if (meta?.wikilink) {
            // data-wikilink-page preserves the exact page name separately from the display
            // text/alias and from href (which a host resolver may rewrite to a real URL).
            if (!this.config.ignoreInternalLinks) {
                // Attribute-driven emission adds data-wikilink/data-target (and data-alias when the
                // display text differs from the page) alongside the default attributes. It is a
                // superset - the widened parser still resolves via data-wikilink-page first.
                const extra = this.config.htmlConfig.sourceAttributes
                    ? ` data-wikilink="true" data-target="${this.escape(meta.link || '')}"`
                        + (node.text && node.text !== meta.link ? ` data-alias="${this.escape(node.text)}"` : '')
                    : '';
                result = `<a href="#${this.escape(this.slugify(meta.link || ''))}" data-wikilink-page="${this.escape(meta.link || '')}"${extra}>${result}</a>`;
            }
        } else if (meta?.link) {
            const isInternal = meta.linkType !== 'external';
            if (!this.config.ignoreInternalLinks || !isInternal) {
                const linkTitle = meta.title ? ` title="${this.escape(meta.title)}"` : '';
                result = `<a href="${sanitizeUrl(meta.link)}"${linkTitle}${meta.linkType === 'external' ? ' target="_blank"' : ''}>${result}</a>`;
            }
        }
        if (meta?.abbreviationTitle) {
            result = `<abbr title="${this.escape(meta.abbreviationTitle)}">${result}</abbr>`;
        }
        if (meta?.citationKey) {
            result = this.config.htmlConfig.sourceAttributes
                // Attribute-driven emission: a <span class="citation"> carrying data-key, which the
                // widened parser reads back (the default <cite> shape is not attribute-keyed).
                ? `<span class="citation" data-key="${this.escape(meta.citationKey)}">[@${this.escape(meta.citationKey)}]</span>`
                // Default emission: a <cite> carrying the bare key, matching Pandoc's [@citekey].
                : `<cite data-citation-key="${this.escape(meta.citationKey)}">[@${this.escape(meta.citationKey)}]</cite>`;
        }

        return result;
    }

    private getInlineStyles(node: OfficeContentNode, options: { skipFontSize?: boolean, skipBackgroundColor?: boolean } = {}): string {
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
            // A table cell's column alignment (GFM `:---`/`:---:`/`---:`) lives on
            // `CellMetadata.align`, not `alignment`. Emit it as `text-align` on the `<th>`/`<td>`
            // so `HtmlParser` reads it back and the pipe-table markers survive AST -> HTML -> AST.
            // An unaligned cell (no `align`) adds nothing, keeping its HTML byte-identical.
            if (node.type === 'cell' && meta.align) pushSafe('text-align', meta.align);
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
            // Skip a run colour equal to the document default (near-black/near-white) when the caller
            // opts in, so imported text adapts to the reader's theme instead of being pinned.
            if (f.color && !(this.config.htmlConfig.omitDefaultTextColor && isNearDefaultColor(f.color))) pushSafe('color', f.color);
            // Highlights are emitted as <mark> by formatText (see there); when that path owns the
            // background it passes skipBackgroundColor so the colour is not also duplicated here.
            if (f.backgroundColor && !options.skipBackgroundColor) pushSafe('background-color', f.backgroundColor);
            if (f.size && !options.skipFontSize) pushSafe('font-size', f.size);
            // A monospace run is emitted as <code> by formatText, so it must not also become a
            // font-family style here (that was the old, non-semantic inline-code shape).
            if (f.font && f.font !== 'monospace') {
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
            .ocr-text { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; font-size: 0.85em; line-height: 1.4; white-space: pre; overflow-x: auto; margin: 8px 0; color: #2d3436; }
            
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

                table, tr, img, .chart-container, li, .image-container, figure, blockquote, pre {
                    page-break-inside: avoid !important;
                    break-inside: avoid !important;
                }

                /* Repeat a table's header (and footer) rows on every page it spans. */
                thead { display: table-header-group !important; }
                tfoot { display: table-footer-group !important; }

                /* A page-break node (<hr class="page-break">) is a real page break here, not the
                   decorative dashed rule it is on screen: force the break and hide the marker. */
                .page-break {
                    border: none !important;
                    margin: 0 !important;
                    height: 0 !important;
                    page-break-after: always !important;
                    break-after: page !important;
                }
                .page-break::after { content: none !important; display: none !important; }

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
        return text.toLowerCase().replace(/[^\p{L}\p{M}\p{N}]+/gu, '-').replace(/(^-|-$)/g, '');
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

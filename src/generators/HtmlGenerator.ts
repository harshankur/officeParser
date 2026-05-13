import { CellMetadata, CodeMetadata, ConversionResult, GeneratorConfig, HeadingMetadata, ImageMetadata, ListMetadata, NoteMetadata, OfficeContentNode, OfficeParserAST, PageMetadata, SlideMetadata, TextMetadata } from '../types.js';
import { BaseGenerator } from './BaseGenerator.js';

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
    async generate(): Promise<ConversionResult> {
        this.isSpreadsheetMode = this.ast.content.some(n => n.type === 'sheet');
        const isPresentation = this.ast.content.some(n => n.type === 'slide');
        const isPdf = this.ast.content.some(n => n.type === 'page');

        let containerClass = 'container';
        if (this.isSpreadsheetMode) containerClass = 'spreadsheet-container';
        else if (isPresentation) containerClass = 'presentation-container';
        else if (isPdf) containerClass = 'pdf-container';

        const bodyContent = await this.processNodeArray(this.ast.content);

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

        if (this.config.htmlConfig.standalone) {
            title = this.ast.metadata?.title || 'Document';
            metaTags = this.renderMetaTags();
        }

        const styles = this.config.htmlConfig.standalone ? '' : `<style>${this.getPremiumStyles(this.isSpreadsheetMode, isPresentation, isPdf)}</style>`;

        const value = this.config.htmlConfig.standalone ? `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${this.escape(title)}</title>
    ${metaTags}
    ${this.config.includeCharts ? `<script src="${this.config.htmlConfig.chartJsSrc}"></script>` : ''}
    <style>
        ${this.getPremiumStyles(this.isSpreadsheetMode, isPresentation, isPdf)}
    </style>
</head>
<body>
    <div class="${containerClass}">
        <article>
            ${metadataBlock}
            ${bodyContent}
        </article>
        ${spreadsheetTabs}
    </div>
    ${spreadsheetScript}
</body>
</html>` : `${styles}<div class="${containerClass}">${metadataBlock}${bodyContent}${spreadsheetTabs}</div>${spreadsheetScript}`;

        return {
            value,
            messages: this.messages
        };
    }

    private renderMetaTags(): string {
        if (!this.ast?.metadata) return '';
        const m = this.ast.metadata;
        const tags: string[] = [];
        if (m.author) tags.push(`<meta name="author" content="${this.escape(m.author)}">`);
        if (m.description) tags.push(`<meta name="description" content="${this.escape(m.description)}">`);
        if (m.created) tags.push(`<meta name="dcterms.created" content="${new Date(m.created).toISOString()}">`);
        if (m.modified) tags.push(`<meta name="dcterms.modified" content="${new Date(m.modified).toISOString()}">`);
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
        const m = this.ast.metadata;

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
        // Stack to track active lists: { indentation, type }
        const listStack: { indentation: number, type: 'ordered' | 'unordered' }[] = [];

        const closeListsToLevel = (level: number) => {
            while (listStack.length > 0 && listStack[listStack.length - 1].indentation > level) {
                const list = listStack.pop();
                html += (list?.type === 'ordered' ? '</ol>' : '</ul>') + '\n\n';
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
                const indentation = meta?.indentation || 0;

                // Close deeper lists
                closeListsToLevel(indentation);

                // Handle current level
                if (listStack.length > 0 && listStack[listStack.length - 1].indentation === indentation) {
                    if (listStack[listStack.length - 1].type !== type) {
                        // Type changed at same level
                        const last = listStack.pop();
                        html += (last?.type === 'ordered' ? '</ol>' : '</ul>') + '\n';
                        html += (type === 'ordered' ? '<ol>' : '<ul>') + '\n';
                        listStack.push({ indentation, type });
                    }
                } else {
                    // Start a new nested list
                    html += (type === 'ordered' ? '<ol>' : '<ul>') + '\n';
                    listStack.push({ indentation, type });
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

        const result = await processor(node, childrenOutput);

        if (isTable) this.tableNestingLevel--;

        return result;
    }

    /**
     * Internal processor for individual nodes.
     */
    private async nodeProcessor(node: OfficeContentNode, childrenOutput: string): Promise<string> {
        // Handle Style Mapping using the semantic mapping helper
        const mapping = this.getSemanticMapping(node);

        let tag = mapping?.tag || this.getDefaultTag(node);

        // Combine classes from mapping and defaults
        const classes = mapping?.classes ? [...mapping.classes] : [];
        const className = classes.length > 0 ? ` class="${classes.join(' ')}"` : '';

        // Handle Attributes from mapping
        let mappedAttrs = '';
        if (mapping?.attributes) {
            for (const [key, val] of Object.entries(mapping.attributes)) {
                mappedAttrs += ` ${key}="${this.escape(val)}"`;
            }
        }

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
            idAttr = ` id="${anchorIds[0]}"`;
            if (anchorIds.length > 1) {
                extraAnchors = anchorIds.slice(1).map(aid => `<a id="${aid}" name="${aid}"></a>`).join('');
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
                const img = `<img src="${src}" alt="${this.escape(node.text || meta?.altText || '')}"${className}${mappedAttrs}${styleAttr}>`;
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
            const chartData = ${JSON.stringify(chartData)};
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
                const tableHtml = `<table${idAttr}${className}${mappedAttrs}${styleAttr}>${finalChildren}</table>`;
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
                // For sheet-based files, wrap rows in a table
                let finalChildren = childrenOutput;
                const rows = node.children || [];

                // Apply same header detection as 'table' nodes
                if (rows.length > 0 && rows[0].type === 'row') {
                    const firstRow = rows[0];
                    const firstRowCells = firstRow.children || [];
                    const isHeaderStyle = (firstRow.metadata as any)?.style?.toLowerCase().includes('header');
                    const allBold = firstRowCells.length > 0 && firstRowCells.every(c =>
                        c.children?.every(child => child.formatting?.bold === true)
                    );

                    if (isHeaderStyle || allBold) {
                        const headOutput = await this.processNodeRecursive(firstRow, async (n, children) => {
                            return children.replace(/<td/g, '<th').replace(/<\/td>/g, '</th>');
                        });
                        const bodyRows = rows.slice(1);
                        const bodyOutput = await this.processNodeArray(bodyRows);
                        finalChildren = `<thead>${headOutput}</thead><tbody>${bodyOutput}</tbody>`;
                    }
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

                return `${extraAnchors}<div${finalIdAttr}${classAttr}${mappedAttrs}${styleAttr}><table class="spreadsheet-table">${finalChildren}</table></div>`;
            }

            case 'paragraph':
            case 'heading': {
                const tag = node.type === 'heading' ? `h${(node.metadata as HeadingMetadata)?.level || 1}` : 'p';
                return `${extraAnchors}<${tag}${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</${tag}>`;
            }
            case 'slide': {
                const meta = node.metadata as SlideMetadata;
                const slideNum = meta?.slideNumber || '';
                return `${extraAnchors}<section class="slide" data-slide-num="${slideNum}"${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</section>`;
            }
            case 'page': {
                const meta = node.metadata as PageMetadata;
                const pageNum = meta?.pageNumber || '';
                return `${extraAnchors}<section class="page" data-page-num="${pageNum}"${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</section>`;
            }
            case 'note': {
                const meta = node.metadata as NoteMetadata;
                const noteClass = meta?.noteType ? ` note-${meta.noteType}` : '';
                return `${extraAnchors}<div class="slide-note${noteClass}"${idAttr}${className}${mappedAttrs}${styleAttr}>${childrenOutput}</div>`;
            }

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
        if (meta?.link) {
            const isInternal = meta.linkType !== 'external';
            if (!this.config.ignoreInternalLinks || !isInternal) {
                result = `<a href="${meta.link}"${meta.linkType === 'external' ? ' target="_blank"' : ''}>${result}</a>`;
            }
        }

        return result;
    }

    private getInlineStyles(node: OfficeContentNode): string {
        const styles: string[] = [];

        if (node.metadata) {
            const meta = node.metadata as any;
            if (meta.alignment) styles.push(`text-align: ${meta.alignment}`);
            if (meta.backgroundColor) styles.push(`background-color: ${meta.backgroundColor}`);
            if (meta.verticalAlign) styles.push(`vertical-align: ${meta.verticalAlign}`);
            if (meta.paragraphIndentation) {
                const ind = meta.paragraphIndentation;
                if (ind.left) styles.push(`margin-left: ${ind.left / 20}pt`);
                if (ind.right) styles.push(`margin-right: ${ind.right / 20}pt`);
                if (ind.firstLine) styles.push(`text-indent: ${ind.firstLine / 20}pt`);
            }
        }

        if (node.formatting) {
            const f = node.formatting;
            if (f.color) styles.push(`color: ${f.color}`);
            if (f.backgroundColor) styles.push(`background-color: ${f.backgroundColor}`);
            if (f.size) styles.push(`font-size: ${f.size}`);
            if (f.font) styles.push(`font-family: ${f.font}, sans-serif`);
        }

        return styles.join('; ');
    }

    private getPremiumStyles(isSpreadsheet: boolean = false, isPresentation: boolean = false, isPdf: boolean = false): string {
        return `
            :root {
                --primary-color: #2c3e50;
                --text-color: #333;
                --bg-color: #f3f4f6;
                --container-bg: #ffffff;
                --border-color: #e9ecef;
                --accent-color: #3498db;
                --shadow: 0 10px 25px rgba(0,0,0,0.05);
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
                max-width: 850px;
                margin: 0 auto;
                background: var(--container-bg);
                padding: 60px 80px;
                border-radius: 12px;
                box-shadow: var(--shadow);
            }
            .presentation-container, .pdf-container {
                max-width: 1200px;
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
            
            ul, ol { margin: 1.5em 0; padding-left: 2em; }
            li { margin-bottom: 0.6em; }
            
            .table-container {
                width: 100%;
                overflow-x: auto;
                -webkit-overflow-scrolling: touch;
                margin: 25px 0;
            }
            table {
                width: auto;
                max-width: 100%;
                min-width: ${isSpreadsheet ? '100%' : '300px'};
                border-collapse: collapse;
                margin: 0 auto;
                border: 1px solid var(--border-color);
                border-radius: 8px;
                overflow: hidden;
            }
            th, td {
                padding: 12px 16px;
                border-bottom: 1px solid var(--border-color);
                border-right: 1px solid var(--border-color);
                text-align: left;
                vertical-align: top;
                overflow-wrap: break-word;
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
            tr:last-child td { border-bottom: none; }
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
                min-width: 100px;
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
                min-height: 600px;
                margin: 0 auto 40px auto;
                padding: 60px 80px;
                border-radius: 12px;
                box-shadow: 0 10px 30px rgba(0,0,0,0.1);
                position: relative;
                display: flex;
                flex-direction: column;
                justify-content: flex-start;
                border: 1px solid var(--border-color);
                box-sizing: border-box;
                width: 100%;
                overflow-wrap: break-word;
                page-break-after: always;
            }
            .slide-note {
                background: #fdfdfd;
                border: 1px solid var(--border-color);
                border-left: 4px solid var(--accent-color);
                border-radius: 12px;
                padding: 30px 50px;
                margin: 20px auto 60px auto;
                max-width: 1000px;
                font-size: 0.95em;
                color: var(--text-color);
                position: relative;
                box-shadow: 0 4px 15px rgba(0,0,0,0.05);
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
                min-height: 1000px;
                margin: 0 auto 40px auto;
                padding: 70px 90px;
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
                .slide { min-height: auto; padding: 40px 30px; }
                .page { padding: 40px; }
            }
            
            /* Nested Table Styles */
            td table {
                margin: 15px 0;
                border-radius: 6px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.03);
                background-color: #ffffff;
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
        `;
    }

    protected override slugify(text: string): string {
        return text.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
    }

    private escape(text: string): string {
        if (typeof text !== 'string') return text;
        return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }
}

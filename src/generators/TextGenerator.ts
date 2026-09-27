import { ConversionResult, GeneratorConfig, OfficeContentNode, OfficeContentNodeType, OfficeParserAST } from '../types.js';
import { BaseGenerator } from './BaseGenerator.js';
import { clampRepeat, median } from '../utils/numberUtils.js';
import { base64ByteLength } from '../utils/officeGenUtils.js';
import { trimRepeated } from '../utils/textUtils.js';

/**
 * Separator between table/sheet cells on the paths that don't render an aligned grid (a `table`
 * with `preserveLayout` off, and `sheet`/`row`/`cell`, which never go through `renderTable`).
 * A tab is the conventional plain-text column delimiter and, unlike a space, survives a value that
 * already contains spaces.
 */
const CELL_SEPARATOR = '\t';

/**
 * How plain text treats each node type once the explicit branches in the processor have had their
 * say: a `block` is its rendered children followed by a newline (a paragraph-like unit), an
 * `inline` contributes its children as-is. A full `Record` rather than an allowlist so the compiler
 * rejects a new OfficeContentNodeType until it is classified here.
 */
const TEXT_NODE_CLASS: Readonly<Record<OfficeContentNodeType, 'block' | 'inline'>> = {
    paragraph: 'block', heading: 'block', row: 'block', sheet: 'block', slide: 'block', note: 'block',
    list: 'block', table: 'block', code: 'block',
    // A definition's term and description are lines of their own, and an admonition's text ends its
    // line (they ran into each other and into the paragraph after them).
    admonition: 'block', definitionList: 'block', definitionTerm: 'block', definitionDescription: 'block',
    text: 'inline', image: 'inline', chart: 'inline', drawing: 'inline', cell: 'inline', page: 'inline',
    break: 'inline', comment: 'inline', header: 'inline', footer: 'inline', slideMaster: 'inline',
    embed: 'inline',
};

/**
 * Generates plain text from an AST.
 */
export class TextGenerator extends BaseGenerator<'text'> {
    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'text'>) {
        super('text', ast, config);
    }


    /**
     * Generates plain text by concatenating text content from nodes.
     */
    async generate(): Promise<ConversionResult<'text'>> {
        let output = '';
        const newline = this.config.textConfig.newlineDelimiter;

        // Add Metadata Header
        const meta = this.effectiveMetadata;
        if (this.config.renderMetadata && meta) {
            // The header is a structured `Key: value` block terminated by a rule, and consumers
            // parse it as such. A value containing a line break would forge extra fields - a title
            // of "Real\nAuthor: Attacker" renders an Author line the document never had - so line
            // breaks are folded to spaces, matching how CsvGenerator guards its `#` comment block.
            // Plain text has no code-execution context, but fabricated structure is still a lie
            // about the document, and every AST string is treated as attacker-controlled.
            const oneLine = (value: unknown): string => String(value ?? '').replace(/[\r\n]+/g, ' ');
            if (meta.title) output += `Title: ${oneLine(meta.title)}${newline}`;
            if (meta.author) output += `Author: ${oneLine(meta.author)}${newline}`;
            // Guarded like HtmlGenerator's toIsoDate: a malformed date must not render the literal
            // "Invalid Date" into the header as if it were the document's creation time.
            if (meta.created) {
                const created = new Date(meta.created as any);
                if (!isNaN(created.getTime())) output += `Created: ${oneLine(created.toLocaleString())}${newline}`;
            }
            output += `-------------------${newline}${newline}`;
        }

        const processor = async (node: OfficeContentNode, childrenOutput: string): Promise<string> => {
            // includeCharts: false omits charts in every generator. Plain text otherwise renders a
            // chart's data series (it lives in `node.text`); this drops it when charts are turned off.
            if (node.type === 'chart' && this.config.includeCharts === false) return '';

            // Return raw text for text nodes
            if (node.type === 'text' || node.type === 'code') {
                return node.text || '';
            }

            // Handle explicit breaks
            if (node.type === 'break') {
                return newline;
            }

            if (node.type === 'image') {
                const mode = this.imageMode();
                if (mode === 'none') return '';
                const meta = node.metadata as any;
                const ocr = (node.text || '').trim();
                // Plain text cannot embed the image: 'ocr-text-only' is just the recognized text,
                // 'image+ocr-text' is the placeholder plus the text, and 'image-only' is the placeholder.
                if (mode === 'ocr-text-only') return ocr ? `${ocr}${newline}` : '';
                const label = `[Image: ${meta?.altText || meta?.attachmentName || 'Untitled'}]`;
                if (mode === 'image+ocr-text' && ocr) return `${label}${newline}${ocr}${newline}`;
                // An image too large to inline anywhere renders its recognized text instead of a bare
                // placeholder, matching what maxInlineImageBytes promises and what Markdown now does:
                // for a scanned page that text is the whole content, and a `[Image: ...]` line loses it.
                if (mode === 'image-only' && ocr && this.overInlineCap(meta?.attachmentName)) return `${ocr}${newline}`;
                return `${label}${newline}`;
            }

            if (node.type === 'embed') {
                const meta = node.metadata as any;
                return meta?.url ? `[${meta.embedType === 'youtube' ? 'YouTube' : 'Embed'}: ${meta.url}]${newline}` : '';
            }

            if (node.type === 'admonition') {
                const meta = node.metadata as any;
                if (childrenOutput.trim() === '') return '';
                const label = (meta?.admonitionType || 'note').toUpperCase();
                return `[${label}] ${childrenOutput.trim()}${newline}`;
            }

            if (node.type === 'table' && this.config.textConfig.preserveLayout) {
                return await this.renderTable(node, processor, newline);
            }

            if (node.type === 'list' && this.config.textConfig.preserveLayout) {
                const meta = node.metadata as any;
                const indentSpaces = ' '.repeat(4);
                // Clamp the nesting depth: `indentation` derives from an uncapped document `ilvl`, so a
                // hostile value would otherwise repeat the indent into a multi-GB string.
                const indent = indentSpaces.repeat(clampRepeat(meta?.indentation || 0, 64));
                const marker = meta?.listType === 'ordered' ? `${(meta.itemIndex ?? 0) + 1}. ` : '- ';
                return `${indent}${marker}${childrenOutput.trimStart()}` + (childrenOutput.endsWith(newline) ? '' : newline);
            }

            // Cells need an explicit separator between them. Without one they concatenate into a
            // single run - "ITEM" + "NEEDED" becomes "ITEMNEEDED", which is unreadable and loses the
            // cell boundary entirely. This was previously masked for spreadsheets only because
            // XLSX/ODS cell values happen to carry a trailing non-breaking space in the source data,
            // so they *appeared* separated; formats whose cell text has no trailing whitespace (MD,
            // HTML) collided outright. Relying on the data to supply a delimiter isn't a separator.
            //
            // Applies to the paths that don't go through renderTable: a `table` when preserveLayout
            // is off, and `sheet`/`row`/`cell` always, since a sheet is not a `table` node.
            if (node.type === 'cell') {
                // Only when the cell doesn't already end in a line break. Cell content varies by
                // source: DOCX/ODT/RTF wrap it in a paragraph, which emits its own newline and so
                // already separates cells, while MD/HTML/XLSX hold bare text that would otherwise
                // collide. Appending unconditionally would push a stray tab onto the formats that
                // were already correct - measured, not assumed: doing so moved DOCX from 6 to 178
                // line-edits away from the reference plain-text output.
                if (childrenOutput === '' || childrenOutput.endsWith(newline)) return childrenOutput;
                return childrenOutput + CELL_SEPARATOR;
            }

            // Append newline for block-level elements to maintain structure.
            if (node.type === 'row') {
                // Trailing separator on the final cell is an artifact of appending one per cell, not
                // content, so drop it rather than leaving every row ending in a stray tab.
                const row = childrenOutput.endsWith(CELL_SEPARATOR)
                    ? childrenOutput.slice(0, -CELL_SEPARATOR.length)
                    : childrenOutput;
                if (row === '') return '';
                return row + (row.endsWith(newline) ? '' : newline);
            }
            if (TEXT_NODE_CLASS[node.type] === 'block') {
                // Drop a block only when it is genuinely empty, not merely whitespace. A paragraph
                // containing spaces is content the document actually holds - discarding it silently
                // deletes an author's blank-but-not-empty line, so filter on `!== ''` rather than on
                // trimmed emptiness.
                if (childrenOutput === '') return '';
                return childrenOutput + (childrenOutput.endsWith(newline) ? '' : newline);
            }

            // Fallback for node types with no explicit handling above. Prefer rendered children,
            // but fall back to the node's own text when it has none: a `chart` carries its whole
            // data series in `text` with zero child nodes, and a CSV `comment` likewise, so
            // returning only `childrenOutput` silently dropped both. Reading `node.text` here covers
            // any future node type of the same shape rather than just the two known today.
            if (!childrenOutput && node.text && !node.children?.length) {
                return node.text + (node.text.endsWith(newline) ? '' : newline);
            }
            return childrenOutput;
        };

        const pageSeparator = this.config.textConfig.pageSeparator;
        let firstTopLevel = true;
        for (const node of this.ast.content) {
            // PDF pages with geometry render as a spatial grid (preserveLayout); everything else flows.
            const useLayout = node.type === 'page' && this.config.textConfig.preserveLayout && this.pageHasLayout(node);
            const piece = useLayout
                ? await this.renderPageLayout(node, processor, newline)
                : await this.processNodeRecursive(node, processor);
            if (!firstTopLevel && node.type === 'page' && piece) output += pageSeparator;
            output += piece;
            firstTopLevel = false;
        }

        if (this.collectedNotes.length > 0 && this.config.textConfig.renderNotes) {
            output += `${newline}${newline}--- Notes ---${newline}`;
            for (const note of this.collectedNotes) {
                output += await this.processNodeRecursive(note, processor);
            }
        }

        // Every block-level node above unconditionally appends its own trailing `newline` as a
        // separator from whatever sibling follows - including, unavoidably, the very last one,
        // which has no sibling to separate from - and renderTable below unconditionally *prepends*
        // one too, as a separator from whatever precedes it (also unavoidably applied when a table
        // is the very first/only node). Both are pure generator artifacts, never part of the
        // document's actual content, so a run of exactly this delimiter at either end is the only
        // thing safe to strip. Nothing else is: not leading/trailing spaces or tabs (e.g. an
        // intentionally-indented opening line, or trailing spaces on the last line - both real
        // content), and not any whitespace that isn't composed of this exact repeated delimiter. A
        // blanket trim()/trimEnd() would silently destroy all of those. (Scanned from each end: a
        // pattern anchored to the end retries every newline of a long run inside the text.)
        return {
            value: trimRepeated(output, newline),
            messages: this.messages
        };
    }

    /**
     * A block starts a line of its own: a list item's text followed by a nested definition list or
     * table ran into the block's first line (`- itemT`).
     */
    protected override childSeparator(previous: string, child: OfficeContentNode): string {
        const newline = this.config.textConfig.newlineDelimiter;
        return TEXT_NODE_CLASS[child.type] === 'block' && previous && !previous.endsWith(newline) ? newline : '';
    }

    /**
     * True when the named attachment is larger than `maxInlineImageBytes`, the size past which an
     * image can no longer travel inside the output. Plain text never embeds a picture at all, so the
     * cap is what decides whether the image is still resolvable alongside the text (small: the
     * `[Image: name]` reference points at something a packager can supply) or effectively lost (large:
     * the recognized text is all that is left of it).
     */
    private overInlineCap(attachmentName: string | undefined): boolean {
        if (!attachmentName) return false;
        const attachment = this.getAttachment(attachmentName);
        return !!attachment && base64ByteLength(attachment.data) > this.config.maxInlineImageBytes;
    }

    /** True when a page carries the geometry needed for spatial layout rendering. */
    private pageHasLayout(page: OfficeContentNode): boolean {
        if (!(page.metadata as any)?.pageWidth) return false;
        let found = false;
        const check = (n: OfficeContentNode) => {
            if (found) return;
            if (n.type === 'text' && n.bounds && (n.text || '').trim()) { found = true; return; }
            for (const c of n.children || []) check(c);
        };
        check(page);
        return found;
    }

    /**
     * Renders one PDF page as a spatial monospace grid: each text run is placed at the column its
     * page x maps to, so columns and tables line up much like the original page (pdftotext -layout).
     * Falls back to flowing the page's children when the geometry is too degenerate to grid.
     */
    private async renderPageLayout(page: OfficeContentNode, processor: any, newline: string): Promise<string> {
        // Memoize onNode verdicts for this page: `collect` visits every descendant, and a flow fallback
        // re-walks the same nodes, so without this the hook would fire twice per node on a fallback page.
        this.onNodeMemo = new WeakMap();
        try {
            return await this.renderPageLayoutInner(page, processor, newline);
        } finally {
            this.onNodeMemo = null;
        }
    }

    private async renderPageLayoutInner(page: OfficeContentNode, processor: any, newline: string): Promise<string> {
        const override = await this.handleOnNode(page);
        if (override === false) return '';
        if (typeof override === 'string') return override + (override.endsWith(newline) ? '' : newline);

        interface Atom { text: string; x: number; y: number; w: number; h: number; }
        const atoms: Atom[] = [];
        const collect = async (n: OfficeContentNode): Promise<void> => {
            const ov = await this.handleOnNode(n);
            if (ov === false) return;
            if (typeof ov === 'string') {
                if (n.bounds && ov) atoms.push({ text: ov, x: n.bounds.x, y: n.bounds.y, w: n.bounds.width, h: n.bounds.height });
                return;
            }
            if (n.type === 'text' && n.bounds && (n.text || '').length) {
                atoms.push({ text: n.text!, x: n.bounds.x, y: n.bounds.y, w: n.bounds.width, h: n.bounds.height });
            } else if (n.type === 'image') {
                const mode = this.imageMode();
                if (mode !== 'none' && n.bounds) {
                    const m = n.metadata as any;
                    const ocr = (n.text || '').trim();
                    // Same over-cap rule as the flow path: an image too large to travel with the
                    // output renders its recognized text rather than a placeholder that loses it.
                    const useOcr = !!ocr && (mode === 'ocr-text-only' || (mode === 'image-only' && this.overInlineCap(m?.attachmentName)));
                    const text = useOcr ? ocr : (mode === 'ocr-text-only' ? '' : `[Image: ${m?.altText || m?.attachmentName || 'Untitled'}]`);
                    if (text) atoms.push({ text, x: n.bounds.x, y: n.bounds.y, w: n.bounds.width, h: n.bounds.height });
                }
                return;
            } else if (n.type === 'list') {
                // The parser strips the item's marker into metadata; re-synthesize it (as flow mode does)
                // and fold it into the item's first placed atom so bullets/numbers survive layout mode.
                const before = atoms.length;
                for (const c of n.children || []) await collect(c);
                if (atoms.length > before) {
                    const meta = n.metadata as any;
                    const marker = meta?.listType === 'ordered' ? `${(meta.itemIndex ?? 0) + 1}. ` : '- ';
                    let firstIdx = before;
                    for (let k = before + 1; k < atoms.length; k++) {
                        const f = atoms[firstIdx];
                        if (atoms[k].y < f.y - 0.5 || (Math.abs(atoms[k].y - f.y) < 1 && atoms[k].x < f.x)) firstIdx = k;
                    }
                    const f = atoms[firstIdx];
                    const perChar = f.text.length ? f.w / f.text.length : 6;
                    f.text = marker + f.text;
                    f.x = Math.max(0, f.x - marker.length * perChar);
                    f.w = f.w + marker.length * perChar;
                }
                return;
            }
            for (const c of n.children || []) await collect(c);
        };
        for (const c of page.children || []) await collect(c);

        // Fall back to flow when there is nothing to place or the character width is degenerate.
        const flowFallback = async (): Promise<string> => {
            let s = '';
            for (const c of page.children || []) s += await this.processNodeRecursive(c, processor);
            return s;
        };
        if (!atoms.length) return flowFallback();

        const charWidths = atoms.filter(a => a.text.trim().length >= 3).map(a => a.w / a.text.length);
        const charW = median(charWidths);
        if (!(charW >= 2 && charW <= 20)) return flowFallback();

        const sortedX = atoms.map(a => a.x).sort((p, q) => p - q);
        const marginX = sortedX[Math.floor(0.02 * sortedX.length)] ?? sortedX[0];
        const pageWidth = (page.metadata as any).pageWidth as number;
        const maxCols = Math.ceil(pageWidth / charW) * 2;

        // Cluster atoms into rows by vertical band overlap.
        atoms.sort((a, b) => a.y - b.y || a.x - b.x);
        const rows: Atom[][] = [];
        for (const a of atoms) {
            const row = rows[rows.length - 1];
            if (row) {
                const ref = row[0];
                const centerA = a.y + a.h / 2;
                const centerRef = ref.y + ref.h / 2;
                if (Math.abs(centerA - centerRef) < 0.6 * Math.min(a.h, ref.h)) { row.push(a); continue; }
            }
            rows.push([a]);
        }

        const rowStrings = rows.map(row => {
            row.sort((a, b) => a.x - b.x);
            let line = '';
            let prevRight: number | null = null; // right edge (x) of the last placed atom
            for (const a of row) {
                let col = Math.round((a.x - marginX) / charW);
                if (col < 0) col = 0;
                if (col > maxCols) col = maxCols;
                if (col > line.length) {
                    line += ' '.repeat(col - line.length);
                } else if (line.length > 0) {
                    // The column is at or before the current line end. Only insert a space when this
                    // atom is separated from the previous one by a real horizontal gap; adjacent runs
                    // (a formatting/link boundary inside a word) must join with no space.
                    const gap = prevRight != null ? a.x - prevRight : 0;
                    if (gap > 0.25 * charW) line += ' ';
                }
                line += a.text;
                prevRight = a.x + a.w;
            }
            return line.trimEnd();
        });

        const centers = rows.map(r => r[0].y + r[0].h / 2);
        const deltas: number[] = [];
        for (let i = 1; i < centers.length; i++) deltas.push(centers[i] - centers[i - 1]);
        const pitch = median(deltas) || 12;

        let out = '';
        for (let i = 0; i < rowStrings.length; i++) {
            if (i > 0) {
                out += newline;
                if (centers[i] - centers[i - 1] > 1.8 * pitch) out += newline;
            }
            out += rowStrings[i];
        }
        return out + newline;
    }


    private async renderTable(node: OfficeContentNode, processor: any, newline: string): Promise<string> {
        if (!node.children || node.children.length === 0) return '';

        const rows: string[][] = [];
        const colWidths: number[] = [];

        for (const rowNode of node.children) {
            // Manual check for onNode in table rows since we are bypassing processNodeRecursive for rows here
            const override = await this.handleOnNode(rowNode);
            if (override === false) continue;
            if (typeof override === 'string') {
                rows.push([override]);
                continue;
            }

            const row: string[] = [];
            if (rowNode.children) {
                for (let i = 0; i < rowNode.children.length; i++) {
                    const cellNode = rowNode.children[i];
                    const cellText = (await this.processNodeRecursive(cellNode, processor))
                        .trim()
                        .replace(/\r?\n/g, ' ');
                    row.push(cellText);
                    colWidths[i] = Math.max(colWidths[i] || 0, cellText.length);
                }
            }
            rows.push(row);
        }

        let tableOutput = newline;
        for (const row of rows) {
            tableOutput += '| ';
            for (let i = 0; i < row.length; i++) {
                tableOutput += (row[i] || '').padEnd(colWidths[i] || 0) + ' | ';
            }
            tableOutput += newline;
        }
        return tableOutput + newline;
    }
}

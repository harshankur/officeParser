import { zipSync } from 'fflate';
import { ConversionResult, GeneratorConfig, OfficeContentNode, OfficeParserAST, OfficeWarningType } from '../types.js';
import { parseRangeString } from '../utils/sheetUtils.js';
import { csvSafeCell } from '../utils/sanitize.js';
import { checkAbortSignal, shownTo80 } from '../utils/errorUtils.js';
import { BaseGenerator } from './BaseGenerator.js';

/** The most cells any of `rows` has (read in a loop: spread into Math.max, a large sheet overflowed the stack). */
const widestRow = (rows: string[][]): number => {
    let widest = 0;
    for (const row of rows) if (row.length > widest) widest = row.length;
    return widest;
};

/** Whether `node` is a chart or holds one (read with a stack: a deep tree is no deep recursion). */
const holdsChart = (node: OfficeContentNode): boolean => {
    const stack = [node];
    while (stack.length) {
        const next = stack.pop()!;
        if (next.type === 'chart') return true;
        if (next.children) for (const child of next.children) stack.push(child);
    }
    return false;
};

/**
 * Generates CSV files from an AST.
 */
export class CsvGenerator extends BaseGenerator<'csv'> {
    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'csv'>) {
        super('csv', ast, config);
    }

    /** The charts this output left out (CSV writes sheets and tables), reported together once (see chartsLeftOutNamed). */
    private chartsLeftOut = 0;
    /**
     * The first chart left out in document order, named in the report, and where it stands: `2k` before the k-th sheet
     * or table (0-based) and after the one before it, `2k + 1` in the k-th.
     */
    private firstChartLeftOut: { node: OfficeContentNode; place: number } | undefined;

    /** Counts `node`, a chart left out at `place` (see firstChartLeftOut), when charts are wanted (includeCharts). */
    private chartLeftOut(node: OfficeContentNode, place: number): void {
        if (!this.config.includeCharts) return;
        this.chartsLeftOut++;
        if (!this.firstChartLeftOut || place < this.firstChartLeftOut.place) this.firstChartLeftOut = { node, place };
    }

    /**
     * The charts left out, as a report names them: how many, and the first's title (else its attachment's name), cut to
     * 80 characters: `a chart ("Sales")`, `3 charts (the first "Sales")`.
     */
    private chartsLeftOutNamed(): { text: string; count: number; name?: string } {
        const count = this.chartsLeftOut;
        const node = this.firstChartLeftOut?.node;
        const attachmentName = (node?.metadata as { attachmentName?: string } | undefined)?.attachmentName;
        const title = this.getAttachment(attachmentName)?.chartData?.title;
        const named = (typeof title === 'string' ? title.trim() : '') || (typeof attachmentName === 'string' ? attachmentName : '');
        const name = named ? shownTo80(named.replace(/\s+/g, ' ')) : undefined;
        const text = count === 1 ? (name ? `a chart ("${name}")` : 'a chart') : (name ? `${count} charts (the first "${name}")` : `${count} charts`);
        return { text, count, ...(name ? { name } : {}) };
    }

    /**
     * A document CSV has no sheet or table to write from: '', reported (CSV output is tabular content only), naming
     * the charts it held, which CSV does not write either.
     */
    private noTables(): ConversionResult<'csv'> {
        if (this.chartsLeftOut) {
            const { text, count, name } = this.chartsLeftOutNamed();
            this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: `a document with no spreadsheet or table, only ${text}`, format: 'csv', reason: 'chart', count, ...(name ? { name } : {}) });
        } else {
            this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: 'a document with no spreadsheet or table', format: 'csv' });
        }
        return { value: '', messages: this.messages };
    }

    /** Reports, once, the charts this output left out beside the sheets and tables it wrote. */
    private warnChartsLeftOut(): void {
        if (!this.chartsLeftOut) return;
        const { text, count, name } = this.chartsLeftOutNamed();
        this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: `${text}, which CSV does not write (it writes sheets and tables)`, format: 'csv', reason: 'chart', count, ...(name ? { name } : {}) });
    }

    /**
     * Counts the charts in `node` (a sheet's or table's child it does not write: not a row or comment; a row's child that
     * is not a cell), the `index`-th sheet or table's, left out (see chartLeftOut). onNode is asked of each node here once,
     * as collectSheetLikeNodes asks of what stands outside the tables: a node it drops (`false`) is left out with what it
     * holds, unreported; a string, which CSV writes only in a cell or for a row, is passed over as there, so a chart given
     * one is left out and counted. Asked only where a chart is: a picture beside the rows is not asked about, as before.
     */
    private async chartsHeldBy(node: OfficeContentNode, index: number): Promise<void> {
        if (!holdsChart(node)) return;
        // In document order: the first chart is the one named.
        const stack = [node];
        while (stack.length) {
            checkAbortSignal(this.config.abortSignal);
            const held = stack.pop()!;
            if (await this.handleOnNode(held) === false) continue;
            if (held.type === 'chart') this.chartLeftOut(held, 2 * index + 1);
            else if (held.children) for (let i = held.children.length - 1; i >= 0; i--) stack.push(held.children[i]);
        }
    }


    /**
     * Generates CSV content from the provided AST.
     * 
     * @returns A CSV string or a ZIP archive containing multiple CSVs
     */
    async generate(): Promise<ConversionResult<'csv'>> {
        const csvConfig = this.config.csvConfig;
        const delimiter = csvConfig.columnDelimiter;
        const mergeSheets = csvConfig.mergeSheets;

        // 1. Collect all "sheet-like" nodes (sheets and tables), and the charts standing outside them
        const sheetNodes: OfficeContentNode[] = [];
        const chartsOutside: { node: OfficeContentNode; place: number }[] = [];
        await this.collectSheetLikeNodes(this.ast.content, sheetNodes, chartsOutside);

        if (sheetNodes.length === 0) {
            // "An empty result is never silent": the document carried no sheet or table for CSV to
            // export (CSV output is tabular-content only), so surface it rather than returning ''.
            for (const chart of chartsOutside) this.chartLeftOut(chart.node, chart.place);
            return this.noTables();
        }

        const metadataHeader = this.config.renderMetadata ? this.renderMetadata(this.ast, delimiter) : '';

        // 2. Filter sheets based on range (each with its index among them all, for where its charts stand)
        let selected = sheetNodes.map((node, index) => ({ node, index }));
        if (csvConfig.sheets) {
            const indices = parseRangeString(csvConfig.sheets);
            selected = indices
                .filter(i => i > 0 && i <= sheetNodes.length)
                .map(i => ({ node: sheetNodes[i - 1], index: i - 1 }));
        }

        if (selected.length === 0) {
            this.warn(OfficeWarningType.SHEET_RANGE_NOT_FOUND, csvConfig.sheets);
            return { value: '', messages: this.messages };
        }
        // Writing every sheet and table, the charts outside them are left out of the output; a range leaves them out
        // with what else it does not select.
        if (!csvConfig.sheets) for (const chart of chartsOutside) this.chartLeftOut(chart.node, chart.place);

        // 3. Generate CSV content for each selected node
        const sheetData: { name: string; rows: string[][] }[] = [];
        let globalMaxCols = 0;

        for (let i = 0; i < selected.length; i++) {
            const { node, index } = selected[i];
            const name = (node.metadata as any)?.sheetName || `Sheet${i + 1}`;
            const rows = await this.renderNodeToRows(node, index);

            const maxCols = widestRow(rows);
            if (mergeSheets) {
                globalMaxCols = Math.max(globalMaxCols, maxCols);
            }

            sheetData.push({ name, rows });
        }
        // The charts in them, and beside them, which CSV does not write: reported once.
        this.warnChartsLeftOut();

        // 4. Handle merging or separate files
        if (mergeSheets) {
            const mergedLines: string[] = [];
            if (metadataHeader) mergedLines.push(metadataHeader.trim());

            for (const sheet of sheetData) {
                if (sheetData.length > 1) mergedLines.push(`# Sheet: ${this.sanitizeComment(sheet.name, delimiter)}`);
                for (const row of sheet.rows) {
                    const paddedRow = [...row];
                    // Don't pad comments
                    if (row.length > 1 || (row.length === 1 && !row[0].startsWith('#'))) {
                        for (let k = this.padWithinBudget(globalMaxCols - paddedRow.length); k > 0; k--) paddedRow.push('');
                    }
                    mergedLines.push(paddedRow.map(v => this.escapeCsvValue(v, delimiter)).join(delimiter));
                }
                mergedLines.push(''); // Blank line between sheets
            }

            return {
                value: mergedLines.join('\n'),
                messages: this.messages
            };
        } else if (sheetData.length > 1) {
            // Create a ZIP archive for multiple sheets
            const zipFiles: Record<string, Uint8Array> = {};
            // Each its own file: two names made alike (`A/B` and `A_B`; `Sales` and `sales`, one file where case is
            // not told apart) were one file, the first overwritten. A name an earlier sheet has is numbered, `A_B-2.csv`,
            // past the names given and every sheet's own name (a sheet named `A_B-2` keeps it), from the number its name
            // reached (each base name's count only rises: linear in the sheets).
            const fileBase = (name: string) => name.replace(/[^\w\s-]/g, '_');
            const own = new Set(sheetData.map(sheet => `${fileBase(sheet.name)}.csv`.toLowerCase()));
            const taken = new Set<string>();
            const nextNumber = new Map<string, number>();
            for (const sheet of sheetData) {
                const sheetMaxCols = widestRow(sheet.rows);
                const csvLines: string[] = [];
                if (metadataHeader) csvLines.push(metadataHeader.trim());

                for (const row of sheet.rows) {
                    const paddedRow = [...row];
                    // Don't pad comments
                    if (row.length > 1 || (row.length === 1 && !row[0].startsWith('#'))) {
                        for (let k = this.padWithinBudget(sheetMaxCols - paddedRow.length); k > 0; k--) paddedRow.push('');
                    }
                    csvLines.push(paddedRow.map(v => this.escapeCsvValue(v, delimiter)).join(delimiter));
                }

                const base = fileBase(sheet.name);
                let fileName = `${base}.csv`;
                if (taken.has(fileName.toLowerCase())) {
                    const key = base.toLowerCase();
                    let n = nextNumber.get(key) ?? 2;
                    while (taken.has(`${key}-${n}.csv`) || own.has(`${key}-${n}.csv`)) n++;
                    fileName = `${base}-${n}.csv`;
                    nextNumber.set(key, n + 1);
                }
                taken.add(fileName.toLowerCase());
                zipFiles[fileName] = new TextEncoder().encode(csvLines.join('\n'));
            }

            const zipBuffer = zipSync(zipFiles);
            return {
                value: zipBuffer,
                messages: this.messages
            };
        } else {
            // Single sheet: return as plain string
            const sheet = sheetData[0];
            const sheetMaxCols = widestRow(sheet.rows);
            const csvLines: string[] = [];
            if (metadataHeader) csvLines.push(metadataHeader.trim());

            for (const row of sheet.rows) {
                const paddedRow = [...row];
                // Don't pad comments
                if (row.length > 1 || (row.length === 1 && !row[0].startsWith('#'))) {
                    for (let k = this.padWithinBudget(sheetMaxCols - paddedRow.length); k > 0; k--) paddedRow.push('');
                }
                csvLines.push(paddedRow.map(v => this.escapeCsvValue(v, delimiter)).join(delimiter));
            }

            return {
                value: csvLines.join('\n'),
                messages: this.messages
            };
        }
    }

    /**
     * Recursively finds all nodes that can be treated as sheets (sheet or table), into `result` in document order, and
     * the charts standing outside them, into `charts` with where each stands (see firstChartLeftOut).
     */
    private async collectSheetLikeNodes(nodes: OfficeContentNode[], result: OfficeContentNode[], charts: { node: OfficeContentNode; place: number }[]): Promise<void> {
        for (const node of nodes) {
            const override = await this.handleOnNode(node);
            if (override === false) {
                continue;
            }

            if (node.type === 'sheet' || node.type === 'table') {
                result.push(node);
            } else if (node.type === 'chart') {
                charts.push({ node, place: 2 * result.length });
            } else if (node.children) {
                await this.collectSheetLikeNodes(node.children, result, charts);
            }
        }
    }

    /**
     * Renders a sheet or table node, the `index`-th of them all, to raw row data, counting the charts it holds, which
     * CSV does not write (see chartLeftOut).
     */
    private async renderNodeToRows(node: OfficeContentNode, index: number): Promise<string[][]> {
        if (!node.children) return [];

        const rows: string[][] = [];
        // A chart beside the rows (an XLSX sheet's), or in a row beside its cells, counted in its place among them.
        const countsCharts = this.config.includeCharts;

        // Text processor for cell content (a chart in a cell writes nothing: counted)
        const cellProcessor = (n: OfficeContentNode, co: string) => {
            if (n.type === 'text' || n.type === 'code') return n.text || '';
            if (n.type === 'break') return '\n';
            if (n.type === 'chart') this.chartLeftOut(n, 2 * index + 1);
            return co;
        };

        for (const rowNode of node.children) {
            if (rowNode.type !== 'row' && rowNode.type !== 'comment') {
                if (countsCharts) await this.chartsHeldBy(rowNode, index);
                continue;
            }
            const override = await this.handleOnNode(rowNode);
            if (override === false) {
                continue;
            }
            if (typeof override === 'string') {
                rows.push([override]);
                continue;
            }

            if (rowNode.type === 'comment') {
                rows.push([rowNode.text || '']);
                continue;
            }

            if (!rowNode.children) {
                rows.push([]);
                continue;
            }

            const rowValues: string[] = [];
            let lastCol = -1;

            for (const cell of rowNode.children) {
                if (cell.type !== 'cell') {
                    if (countsCharts) await this.chartsHeldBy(cell, index);
                    continue;
                }
                const currentCol = (cell.metadata as any)?.col ?? (lastCol + 1);
                
                // Fill gaps with empty strings
                while (lastCol < currentCol - 1) {
                    rowValues.push('');
                    lastCol++;
                }

                const cellText = await this.processNodeRecursive(cell, cellProcessor);
                rowValues.push(cellText);
                
                // Handle colSpan: move lastCol forward
                const colSpan = (cell.metadata as any)?.colSpan || 1;
                lastCol = currentCol + colSpan - 1;
            }
            rows.push(rowValues);
        }

        return rows;
    }

    /**
     * Escapes a value for CSV formatting: RFC 4180 quoting plus a spreadsheet
     * formula-injection guard (see csvSafeCell in ../utils/sanitize.js).
     */
    private escapeCsvValue(val: string, delimiter: string): string {
        return csvSafeCell(val, delimiter);
    }

    /** Sanitizes a value for a `#` comment line (document metadata or sheet name).
     *  Comment lines are free text prefixed with `#`, not RFC-4180 cells, so they
     *  can't be quoted; instead we (a) collapse newlines so the value can't break out
     *  and inject a new row, and (b) replace the column delimiter with a space so a
     *  value like `x,=1+1` can't split into a second cell that a spreadsheet would
     *  evaluate as a formula (CSV formula/DDE injection). */
    private sanitizeComment(val: string, delimiter: string): string {
        let s = String(val ?? '').replace(/[\r\n]+/g, ' ');
        if (delimiter) s = s.split(delimiter).join(' ');
        return s;
    }

    /**
     * Renders metadata as comments.
     */
    private renderMetadata(_ast: OfficeParserAST, delimiter: string): string {
        // Via effectiveMetadata rather than the passed AST so `metadataOverrides` applies here
        // as it does in every other generator.
        const m = this.effectiveMetadata;
        if (!m || Object.keys(m).length === 0) return '';
        let output = '';
        if (m.title) output += `# Title: ${this.sanitizeComment(m.title, delimiter)}\n`;
        if (m.author) output += `# Author: ${this.sanitizeComment(m.author, delimiter)}\n`;
        if (m.created) output += `# Created: ${this.sanitizeComment(new Date(m.created).toLocaleString(), delimiter)}\n`;
        if (m.modified) output += `# Modified: ${this.sanitizeComment(new Date(m.modified).toLocaleString(), delimiter)}\n`;
        if (m.customProperties) {
            for (const [k, v] of Object.entries(m.customProperties)) {
                output += `# ${this.sanitizeComment(k, delimiter)}: ${this.sanitizeComment(String(v), delimiter)}\n`;
            }
        }
        return output ? output + '\n' : '';
    }
}

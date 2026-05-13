import { zipSync } from 'fflate';
import { ConversionResult, GeneratorConfig, OfficeContentNode, OfficeParserAST, OfficeWarningType } from '../types.js';
import { parseRangeString } from '../utils/sheetUtils.js';
import { BaseGenerator } from './BaseGenerator.js';

/**
 * Generates CSV files from an AST.
 */
export class CsvGenerator extends BaseGenerator<'csv'> {
    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'csv'>) {
        super('csv', ast, config);
    }


    /**
     * Generates CSV content from the provided AST.
     * 
     * @returns A CSV string or a ZIP archive containing multiple CSVs
     */
    async generate(): Promise<ConversionResult> {
        const csvConfig = this.config.csvConfig;
        const delimiter = csvConfig.columnDelimiter;
        const mergeSheets = csvConfig.mergeSheets;

        // 1. Collect all "sheet-like" nodes (sheets and tables)
        const sheetNodes = await this.collectSheetLikeNodes(this.ast.content);

        if (sheetNodes.length === 0) {
            return { value: '', messages: this.messages };
        }

        const metadataHeader = this.config.renderMetadata ? this.renderMetadata(this.ast) : '';

        // 2. Filter sheets based on range
        let selectedNodes = sheetNodes;
        if (csvConfig.sheets) {
            const indices = parseRangeString(csvConfig.sheets);
            selectedNodes = indices
                .filter(i => i > 0 && i <= sheetNodes.length)
                .map(i => sheetNodes[i - 1]);
        }

        if (selectedNodes.length === 0) {
            this.warn(OfficeWarningType.SHEET_RANGE_NOT_FOUND, csvConfig.sheets);
            return { value: '', messages: this.messages };
        }

        // 3. Generate CSV content for each selected node
        const sheetData: { name: string; rows: string[][] }[] = [];
        let globalMaxCols = 0;

        for (let i = 0; i < selectedNodes.length; i++) {
            const node = selectedNodes[i];
            const name = (node.metadata as any)?.sheetName || `Sheet${i + 1}`;
            const rows = await this.renderNodeToRows(node);

            const maxCols = Math.max(...rows.map(r => r.length), 0);
            if (mergeSheets) {
                globalMaxCols = Math.max(globalMaxCols, maxCols);
            }

            sheetData.push({ name, rows });
        }

        // 4. Handle merging or separate files
        if (mergeSheets) {
            const mergedLines: string[] = [];
            if (metadataHeader) mergedLines.push(metadataHeader.trim());

            for (const sheet of sheetData) {
                if (sheetData.length > 1) mergedLines.push(`# Sheet: ${sheet.name}`);
                for (const row of sheet.rows) {
                    const paddedRow = [...row];
                    // Don't pad comments
                    if (row.length > 1 || (row.length === 1 && !row[0].startsWith('#'))) {
                        while (paddedRow.length < globalMaxCols) paddedRow.push('');
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
            for (const sheet of sheetData) {
                const sheetMaxCols = Math.max(...sheet.rows.map(r => r.length), 0);
                const csvLines: string[] = [];
                if (metadataHeader) csvLines.push(metadataHeader.trim());

                for (const row of sheet.rows) {
                    const paddedRow = [...row];
                    // Don't pad comments
                    if (row.length > 1 || (row.length === 1 && !row[0].startsWith('#'))) {
                        while (paddedRow.length < sheetMaxCols) paddedRow.push('');
                    }
                    csvLines.push(paddedRow.map(v => this.escapeCsvValue(v, delimiter)).join(delimiter));
                }

                const fileName = `${sheet.name.replace(/[^\w\s-]/g, '_')}.csv`;
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
            const sheetMaxCols = Math.max(...sheet.rows.map(r => r.length), 0);
            const csvLines: string[] = [];
            if (metadataHeader) csvLines.push(metadataHeader.trim());

            for (const row of sheet.rows) {
                const paddedRow = [...row];
                // Don't pad comments
                if (row.length > 1 || (row.length === 1 && !row[0].startsWith('#'))) {
                    while (paddedRow.length < sheetMaxCols) paddedRow.push('');
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
     * Recursively finds all nodes that can be treated as sheets (sheet or table).
     */
    private async collectSheetLikeNodes(nodes: OfficeContentNode[]): Promise<OfficeContentNode[]> {
        const result: OfficeContentNode[] = [];
        for (const node of nodes) {
            const override = await this.handleOnNode(node);
            if (override === false) {
                continue;
            }

            if (node.type === 'sheet' || node.type === 'table') {
                result.push(node);
            } else if (node.children) {
                result.push(...(await this.collectSheetLikeNodes(node.children)));
            }
        }
        return result;
    }

    /**
     * Renders a sheet or table node to raw row data.
     */
    private async renderNodeToRows(node: OfficeContentNode): Promise<string[][]> {
        if (!node.children) return [];

        const rows: string[][] = [];
        const rowNodes = node.children.filter(c => c.type === 'row' || c.type === 'comment');

        // Text processor for cell content
        const cellProcessor = (n: OfficeContentNode, co: string) => {
            if (n.type === 'text' || n.type === 'code') return n.text || '';
            if (n.type === 'break') return '\n';
            return co;
        };

        for (const rowNode of rowNodes) {
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

            const cellNodes = rowNode.children.filter(c => c.type === 'cell');
            const rowValues: string[] = [];
            let lastCol = -1;

            for (const cell of cellNodes) {
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
     * Escapes a value for CSV formatting.
     */
    private escapeCsvValue(val: string, delimiter: string): string {
        const needsQuotes = val.includes(delimiter) || val.includes('"') || val.includes('\n') || val.includes('\r');
        if (!needsQuotes) return val;

        // Double up existing quotes and wrap in quotes
        return `"${val.replace(/"/g, '""')}"`;
    }

    /**
     * Renders metadata as comments.
     */
    private renderMetadata(ast: OfficeParserAST): string {
        if (!ast.metadata) return '';
        const m = ast.metadata;
        let output = '';
        if (m.title) output += `# Title: ${m.title}\n`;
        if (m.author) output += `# Author: ${m.author}\n`;
        if (m.created) output += `# Created: ${new Date(m.created).toLocaleString()}\n`;
        if (m.modified) output += `# Modified: ${new Date(m.modified).toLocaleString()}\n`;
        if (m.customProperties) {
            for (const [k, v] of Object.entries(m.customProperties)) {
                output += `# ${k}: ${v}\n`;
            }
        }
        return output ? output + '\n' : '';
    }
}

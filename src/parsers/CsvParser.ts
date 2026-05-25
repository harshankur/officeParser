import { CellMetadata, FullOfficeParserConfig, OfficeContentNode, OfficeParserAST, SheetMetadata } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { checkAbortSignal } from '../utils/errorUtils.js';

/**
 * Parses a CSV file and extracts a single sheet with rows and cells.
 * 
 * @param buffer - The CSV file as a Buffer
 * @param config - Parser configuration
 * @returns A promise resolving to the parsed AST
 */
export const parseCsv = async (buffer: Buffer, config: FullOfficeParserConfig): Promise<OfficeParserAST> => {
    // Honour cancellation requests before the character-by-character parsing loop starts.
    // CSV has no OCR or async I/O, but very large files can still occupy the thread for a
    // noticeable duration, so short-circuiting on an aborted signal is still worthwhile.
    checkAbortSignal(config.abortSignal);

    const textStr = buffer.toString('utf-8');
    const delimiter = config.csvDelimiter;

    const records: string[][] = [];
    let currentRow: string[] = [];
    let currentCell = '';
    let inQuotes = false;

    for (let i = 0; i < textStr.length; i++) {
        const char = textStr[i];
        const nextChar = textStr[i + 1];

        if (inQuotes) {
            if (char === '"') {
                if (nextChar === '"') {
                    currentCell += '"';
                    i++; // Skip the escaped quote
                } else {
                    inQuotes = false;
                }
            } else {
                currentCell += char;
            }
        } else {
            if (char === '"') {
                inQuotes = true;
            } else if (textStr.substring(i, i + delimiter.length) === delimiter) {
                currentRow.push(currentCell);
                currentCell = '';
                i += delimiter.length - 1; // Skip the rest of the delimiter
            } else if (char === '\n') {
                currentRow.push(currentCell);
                records.push(currentRow);
                currentRow = [];
                currentCell = '';
            } else if (char === '\r') {
                // Ignore carriage return outside of quotes
            } else {
                currentCell += char;
            }
        }
    }

    if (currentCell !== '' || currentRow.length > 0) {
        currentRow.push(currentCell);
        records.push(currentRow);
    }

    // Filter out trailing empty row if the file ended with a newline
    if (records.length > 0 && records[records.length - 1].length === 1 && records[records.length - 1][0] === '') {
        records.pop();
    }

    const rows: OfficeContentNode[] = [];

    records.forEach((record: string[], rowIndex: number) => {
        // Handle comment rows
        if (record.length === 1 && record[0].startsWith('#')) {
            rows.push({
                type: 'comment',
                text: record[0]
            });
            return;
        }

        const cells: OfficeContentNode[] = [];

        record.forEach((val, colIndex) => {
            if (val && val.trim() !== '') {
                const cellMeta: CellMetadata = { row: rowIndex, col: colIndex };
                cells.push({
                    type: 'cell',
                    text: val,
                    metadata: cellMeta,
                    children: [{ type: 'text', text: val }]
                });
            }
        });

        if (cells.length > 0) {
            rows.push({
                type: 'row',
                children: cells
            });
        }
    });

    const sheetMeta: SheetMetadata = { sheetName: 'Sheet1' };
    const sheetNode: OfficeContentNode = {
        type: 'sheet',
        metadata: sheetMeta,
        children: rows,
        rawContent: config.includeRawContent ? textStr : undefined
    };

    const toTextSync = () => {
        return records.map((record: string[]) => record.filter(cell => cell.trim() !== '').join(config.newlineDelimiter))
            .join(config.newlineDelimiter)
            .replace(/\n{3,}/g, '\n\n');
    };

    return createAST('csv', { title: 'Sheet1' }, [sheetNode], [], config, toTextSync);
};

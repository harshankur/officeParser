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

    // A UTF-8 byte order mark (Excel writes one) is no text of the first cell.
    const textStr = buffer.toString('utf-8').replace(/^\uFEFF/, '');
    // 'auto' (the default) reads the delimiter from the file, and the parse's config says which it was,
    // so CSV output written from it uses the same one.
    const delimiter = config.csvDelimiter === 'auto' || !config.csvDelimiter ? sniffDelimiter(textStr) : config.csvDelimiter;
    config.csvDelimiter = delimiter;

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

    return createAST('csv', { title: 'Sheet1' }, [sheetNode], [], config, undefined);
};

/** The delimiters `csvDelimiter: 'auto'` chooses among, in the order a tie goes to. */
const CANDIDATE_DELIMITERS = [',', ';', '\t', '|'] as const;

/** At most this many records, and characters, are read to choose a delimiter. */
const SNIFF_RECORDS = 50;
const SNIFF_CHARACTERS = 64 * 1024;

/**
 * The delimiter a CSV file is written with, as LibreOffice and Python's csv.Sniffer find it: of `,`, `;`,
 * tab and `|`, counted outside quotes in each of the first records, the one that splits the most records
 * into the same number of fields (more than one). German, French and other European spreadsheets write
 * `;` with decimal commas: read with `,`, `Müller;1.234,56` split the number. A tie between `,` and `;`
 * goes to `;` where every comma is between two digits (`a;1,5`), a decimal comma. A comma when none does.
 */
export function sniffDelimiter(text: string): string {
    const counts = CANDIDATE_DELIMITERS.map(() => [] as number[]);
    let current = CANDIDATE_DELIMITERS.map(() => 0);
    let inQuotes = false;
    let records = 0;
    // Whether every comma outside quotes is between two digits, as a decimal comma is.
    let decimalCommas = true;
    const end = Math.min(text.length, SNIFF_CHARACTERS);
    const finishRecord = () => {
        if (current.some(n => n > 0)) current.forEach((n, i) => counts[i].push(n));
        current = CANDIDATE_DELIMITERS.map(() => 0);
        records++;
    };
    for (let i = 0; i < end && records < SNIFF_RECORDS; i++) {
        const ch = text[i];
        if (ch === '"') { inQuotes = !inQuotes; continue; }
        if (inQuotes) continue;
        if (ch === '\n') { finishRecord(); continue; }
        const at = CANDIDATE_DELIMITERS.indexOf(ch as typeof CANDIDATE_DELIMITERS[number]);
        if (at >= 0) current[at]++;
        if (ch === ',' && !(/\d/.test(text[i - 1] ?? '') && /\d/.test(text[i + 1] ?? ''))) decimalCommas = false;
    }
    if (records < SNIFF_RECORDS && !inQuotes) finishRecord();
    let best = 0, bestScore = 0;
    const scores = counts.map(perRecord => {
        // The records that split into the most common number of fields, when that number is above one.
        const tally = new Map<number, number>();
        for (const n of perRecord) if (n > 0) tally.set(n, (tally.get(n) ?? 0) + 1);
        return Math.max(0, ...tally.values());
    });
    scores.forEach((score, i) => { if (score > bestScore) { best = i; bestScore = score; } });
    if (best === 0 && bestScore > 0 && scores[1] === bestScore && decimalCommas) return ';';
    return CANDIDATE_DELIMITERS[best];
}

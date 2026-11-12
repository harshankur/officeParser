/**
 * Parses a range string (e.g., "1", "1-3", "1,2", "1,3-5, 7") into an array of numbers.
 * 
 * @param rangeStr - The range string to parse
 * @returns An array of unique, sorted numbers (1-based indices)
 */
export function parseRangeString(rangeStr: string): number[] {
    const result = new Set<number>();
    const segments = rangeStr.split(',');

    for (const segment of segments) {
        const trimmed = segment.trim();
        if (trimmed.includes('-')) {
            const [startStr, endStr] = trimmed.split('-');
            const start = parseInt(startStr, 10);
            const end = parseInt(endStr, 10);
            if (!isNaN(start) && !isNaN(end)) {
                const actualStart = Math.min(start, end);
                const actualEnd = Math.max(start, end);
                for (let i = actualStart; i <= actualEnd; i++) {
                    result.add(i);
                }
            }
        } else {
            const val = parseInt(trimmed, 10);
            if (!isNaN(val)) {
                result.add(val);
            }
        }
    }

    return Array.from(result).sort((a, b) => a - b);
}

/** A sheet's last column (XFD) and last row (1,048,576), 0-based. */
const LAST_COLUMN = 16383;
const LAST_ROW = 1048575;

/** A run of the characters a name, a number or a reference is made of (`A1`, `$B$5`, `_xlfn.XLOOKUP`, `1.5E`). */
const FORMULA_TOKEN = /[\p{L}\p{N}_.$\\?]+/uy;
/** An error value (`#REF!`, `#N/A`, `#DIV/0!`), which holds what reads as a reference. */
const FORMULA_ERROR = /#[A-Za-z][A-Za-z0-9_/]*[!?]?/y;
const REFERENCE_PATTERNS = {
    cell: /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/,
    column: /^(\$?)([A-Za-z]{1,3})$/,
    row: /^(\$?)(\d{1,7})$/,
} as const;

const columnIndex = (letters: string): number => {
    let index = 0;
    for (const ch of letters.toUpperCase()) index = index * 26 + ch.charCodeAt(0) - 64;
    return index - 1;
};

const columnLetters = (index: number): string => {
    let letters = '';
    for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + (n - 1) % 26) + letters;
    return letters;
};

/** A column part (`$` and letters) moved by `by` unless anchored; null past the sheet's edge, undefined for no column. */
function movedColumn(anchor: string, letters: string, by: number): string | null | undefined {
    const index = columnIndex(letters);
    if (index > LAST_COLUMN) return undefined;
    if (anchor || !by) return anchor + letters;
    const moved = index + by;
    return moved < 0 || moved > LAST_COLUMN ? null : columnLetters(moved);
}

/** A row part (`$` and digits) moved by `by` unless anchored; null past the sheet's edge, undefined for no row. */
function movedRow(anchor: string, digits: string, by: number): string | null | undefined {
    const index = Number(digits) - 1;
    if (index < 0 || index > LAST_ROW) return undefined;
    if (anchor || !by) return anchor + digits;
    const moved = index + by;
    return moved < 0 || moved > LAST_ROW ? null : String(moved + 1);
}

/** A token as a reference of `kind` moved by rows and columns; null past the sheet's edge, undefined when it is none. */
function movedReference(token: string, kind: keyof typeof REFERENCE_PATTERNS, rows: number, cols: number): string | null | undefined {
    const m = REFERENCE_PATTERNS[kind].exec(token);
    if (!m) return undefined;
    if (kind === 'row') return movedRow(m[1], m[2], rows);
    const column = movedColumn(m[1], m[2], cols);
    if (kind === 'column' || column === undefined) return column;
    const row = movedRow(m[3], m[4], rows);
    if (row === undefined) return undefined;
    return column === null || row === null ? null : column + row;
}

/** Where a quoted part (`"text"`, `'Sheet name'`, its quote doubled inside) that starts at `from` ends. */
function quotedEnd(formula: string, from: number): number {
    const quote = formula[from];
    for (let at = from + 1; at < formula.length; at++) {
        if (formula[at] !== quote) continue;
        if (formula[at + 1] !== quote) return at + 1;
        at++;
    }
    return formula.length;
}

/** Where a bracketed part (`[Col]`, `[[#This Row],[Col]]`, `[1]`, `'` escaping a character in it) that starts at `from` ends. */
function bracketEnd(formula: string, from: number): number {
    let depth = 0;
    for (let at = from; at < formula.length; at++) {
        const ch = formula[at];
        if (ch === "'") at++;
        else if (ch === '[') depth++;
        else if (ch === ']' && --depth === 0) return at + 1;
    }
    return formula.length;
}

/**
 * A formula (as XLSX holds it, without `=`) moved by `rows` and `cols`, as Excel moves a formula filled
 * to another cell: each relative part of a reference moves, `$`-anchored parts stay, and a reference
 * moved past the sheet's edge is `#REF!` (a range with either end past it, the whole range). Strings,
 * quoted and unquoted sheet names, bracketed parts (table columns, a linked workbook's index), error
 * values, numbers and names (`LOG10(`, `Tax_2024`) are kept as written. One pass over the formula.
 */
export function shiftFormula(formula: string, rows: number, cols: number): string {
    let out = '';
    let at = 0;
    const tokenAt = (from: number): string | undefined => {
        FORMULA_TOKEN.lastIndex = from;
        return FORMULA_TOKEN.exec(formula)?.[0];
    };
    // A function's name, a sheet's, a table's: what a reference never is.
    const namesSomething = (end: number): boolean => formula[end] === '(' || formula[end] === '!' || formula[end] === '[';
    while (at < formula.length) {
        const ch = formula[at];
        if (ch === '"' || ch === "'" || ch === '[') {
            const end = ch === '[' ? bracketEnd(formula, at) : quotedEnd(formula, at);
            out += formula.slice(at, end);
            at = end;
            continue;
        }
        if (ch === '#') {
            FORMULA_ERROR.lastIndex = at;
            const error = FORMULA_ERROR.exec(formula)?.[0];
            if (error) {
                out += error;
                at += error.length;
                continue;
            }
        }
        const token = tokenAt(at);
        if (!token) {
            out += ch;
            at++;
            continue;
        }
        const end = at + token.length;
        if (namesSomething(end)) {
            out += token;
            at = end;
            continue;
        }
        const second = formula[end] === ':' ? tokenAt(end + 1) : undefined;
        if (second !== undefined && !namesSomething(end + 1 + second.length)) {
            const range = (['cell', 'column', 'row'] as const).map(kind => [movedReference(token, kind, rows, cols), movedReference(second, kind, rows, cols)])
                .find(([from, to]) => from !== undefined && to !== undefined);
            if (range) {
                out += range[0] === null || range[1] === null ? '#REF!' : `${range[0]}:${range[1]}`;
                at = end + 1 + second.length;
                continue;
            }
        }
        const cell = movedReference(token, 'cell', rows, cols);
        out += cell === undefined ? token : cell ?? '#REF!';
        at = end;
    }
    return out;
}

/**
 * Whether a column is in one of the ranges given (`[first, last]`, inclusive), found by halving: a sheet
 * with as many hidden ranges as cells was looked through whole for every cell.
 */
export function inRanges(ranges: [number, number][]): (index: number) => boolean {
    if (!ranges.length) return () => false;
    const sorted = [...ranges].sort((a, b) => a[0] - b[0]);
    // Overlapping and touching ranges made one, so the starts and ends both rise.
    const merged: [number, number][] = [];
    for (const [first, last] of sorted) {
        const top = merged[merged.length - 1];
        if (top && first <= top[1] + 1) top[1] = Math.max(top[1], last); else merged.push([first, last]);
    }
    return (index: number) => {
        let low = 0, high = merged.length - 1;
        while (low <= high) {
            const mid = (low + high) >> 1;
            if (index < merged[mid][0]) high = mid - 1;
            else if (index > merged[mid][1]) low = mid + 1;
            else return true;
        }
        return false;
    };
}

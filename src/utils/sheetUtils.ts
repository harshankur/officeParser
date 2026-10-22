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

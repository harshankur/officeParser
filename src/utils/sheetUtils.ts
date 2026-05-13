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

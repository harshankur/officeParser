/**
 * Date Parsing Utilities
 * 
 * Provides robust functions for parsing date strings from various office formats.
 * Handles standard ISO dates, PDF-specific date formats, and malformed strings.
 * 
 * @module dateUtils
 */

/**
 * Parses a date string into a Date object.
 * Handles standard ISO formats and falls back to native parsing.
 * Returns undefined instead of "Invalid Date" if parsing fails.
 * 
 * @param dateString - The date string to parse
 * @returns Parsed Date object or undefined if parsing fails
 */
export function parseOfficeDate(dateString: string | undefined): Date | undefined {
    if (!dateString) return undefined;

    try {
        // PDF-specific format detection: D:YYYYMMDDHHmmSSOHH'mm'
        if (dateString.startsWith('D:')) {
            return parsePdfDate(dateString);
        }

        const date = new Date(dateString);
        return isNaN(date.getTime()) ? undefined : date;
    } catch {
        return undefined;
    }
}

/**
 * Internal helper for PDF-specific date format: D:YYYYMMDDHHmmSSOHH'mm'
 * @param dateString - The PDF date string
 */
function parsePdfDate(dateString: string): Date | undefined {
    try {
        // Remove "D:" prefix
        let str = dateString.slice(2);

        // Extract components: YYYYMMDDHHmmSS
        const year = parseInt(str.slice(0, 4), 10);
        const month = parseInt(str.slice(4, 6), 10) - 1; // 0-indexed
        const day = parseInt(str.slice(6, 8), 10) || 1;
        const hour = parseInt(str.slice(8, 10), 10) || 0;
        const minute = parseInt(str.slice(10, 12), 10) || 0;
        const second = parseInt(str.slice(12, 14), 10) || 0;

        // Handle timezone if present
        const tzMatch = str.slice(14).match(/([+-Z])(\d{2})'?(\d{2})?'?/);
        if (tzMatch) {
            if (tzMatch[1] === 'Z') {
                return new Date(Date.UTC(year, month, day, hour, minute, second));
            }
            
            const tzSign = tzMatch[1] === '-' ? -1 : 1;
            const tzHours = parseInt(tzMatch[2], 10) || 0;
            const tzMinutes = parseInt(tzMatch[3], 10) || 0;
            const offset = tzSign * (tzHours * 60 + tzMinutes);

            // Create date in UTC and adjust for timezone
            const utc = Date.UTC(year, month, day, hour, minute, second);
            return new Date(utc - offset * 60000);
        }

        return new Date(year, month, day, hour, minute, second);
    } catch {
        return undefined;
    }
}

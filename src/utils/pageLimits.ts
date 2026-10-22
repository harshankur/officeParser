/**
 * The pages documents are read and written on: a side of at most two hundred inches (OOXML's and PDF's
 * limit), so a poster's A2, A1 or A0 page and an architect's E sheet keep their size; a larger number is
 * no page, and the document's own page (or A4) stands for it.
 *
 * @module utils/pageLimits
 */

/** The longest side a page can have, in points (200 inches). */
export const MAX_PAGE_PT = 14_400;

/** Whether a width and height in points make a page: at least an inch each way, at most MAX_PAGE_PT. */
export const isPageSize = (width: number, height: number): boolean =>
    Number.isFinite(width) && Number.isFinite(height) && width >= 72 && height >= 72 && width <= MAX_PAGE_PT && height <= MAX_PAGE_PT;

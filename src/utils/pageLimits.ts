/**
 * The pages documents are read and written on: a side of at most two hundred inches (OOXML's and PDF's
 * limit), so a poster's A2, A1 or A0 page and an architect's E sheet keep their size; a larger number is
 * no page, and the document's own page (or A4) stands for it.
 *
 * @module utils/pageLimits
 */
import type { PageSetup } from '../types.js';

/** The longest side a page can have, in points (200 inches). */
export const MAX_PAGE_PT = 14_400;

/** Whether a width and height in points make a page: at least an inch each way, at most MAX_PAGE_PT. */
export const isPageSize = (width: number | undefined, height: number | undefined): boolean =>
    Number.isFinite(width) && Number.isFinite(height) && width! >= 72 && height! >= 72 && width! <= MAX_PAGE_PT && height! <= MAX_PAGE_PT;

/** A PageSetup that sets its page: a width and height that make one (see isPageSize). */
export type SizedPageSetup = PageSetup & { width: number; height: number };

/** Whether a section's setup sets its page (see isPageSize): one that only numbers its pages sets none. */
export const hasPageSize = (setup: PageSetup | undefined): setup is SizedPageSetup => !!setup && isPageSize(setup.width, setup.height);

/**
 * A section's setup on the page it is on: its own where it sets a page, else the document's page (its
 * size, margins and header and footer distances) with the section's own columns, numbering and title
 * page. (A section that only numbers its pages its own way is on the document's pages, not on A4.)
 */
export function onDocumentPage(setup: PageSetup | undefined, documentPage: PageSetup | undefined): PageSetup | undefined {
    if (!setup || hasPageSize(setup)) return setup;
    const { width: _w, height: _h, margin: _m, headerDistance: _hd, footerDistance: _fd, ...own } = setup;
    if (!hasPageSize(documentPage)) return own;
    const { width, height, margin, headerDistance, footerDistance } = documentPage;
    return { width, height, ...(margin ? { margin } : {}), ...(headerDistance !== undefined ? { headerDistance } : {}), ...(footerDistance !== undefined ? { footerDistance } : {}), ...own };
}

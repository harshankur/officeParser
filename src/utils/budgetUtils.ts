/**
 * The size of the document behind each parse, for the budgets that grow with it.
 *
 * A budget (`decompressionLimits.maxTableCells`, `maxRepeatedContent`, the grid positions writers fill)
 * bounds what a small file can amplify into: a few hundred bytes asking for a million cells, or one
 * string shown in a million places. A flat bound also cuts short a large document that is simply
 * large (a 4 MB workbook of 1.2 million cells lost its last 200,000), so each of these budgets grows
 * with the bytes of the document, as `pdfParserConfig`'s do: an allowance per byte that real
 * documents stay well inside, and that a small crafted file cannot use to amplify.
 *
 * The size is kept by the parse's config object (what every parser and budget of that parse holds)
 * and by the AST the parse gives, for the budgets its writers apply.
 */

/** The bytes of the document each parse (its config object) or AST came from. */
const documentBytes = new WeakMap<object, number>();

/**
 * Records that `holder` (a parse's config object, or an AST) came from a document of `bytes`. The
 * first record stands: a document read inside another (a DOCX chunk) shares its parse's config, and
 * its budgets are the outer document's.
 */
export const noteDocumentBytes = (holder: object, bytes: number): void => {
    if (!documentBytes.has(holder) && Number.isFinite(bytes) && bytes > 0) documentBytes.set(holder, bytes);
};

/** The bytes of the document `holder` came from; 0 when not known (an AST built in code, or copied). */
export const documentBytesOf = (holder: object | undefined): number => (holder && documentBytes.get(holder)) || 0;

/**
 * Cells a document may yield for each of its bytes, beyond `decompressionLimits.maxTableCells`. Real
 * workbooks hold under half a cell per byte of their zip (a cell's XML compresses to two bytes or more).
 */
export const TABLE_CELLS_PER_BYTE = 1;

/**
 * Characters a document may repeat by reference for each of its bytes, beyond
 * `decompressionLimits.maxRepeatedContent`. A workbook showing one 200-character note in 150,000 cells
 * (1.4 MB) repeats 20 million characters.
 */
export const REPEATED_CHARACTERS_PER_BYTE = 16;

/**
 * Empty grid positions a document's tables may fill for each of its bytes, beyond the million every
 * document may (see sheetGridUtils). A sparse sheet of 3,000 rows with values in its first and 400th
 * columns (19 KB) fills 1.2 million.
 */
export const GRID_POSITIONS_PER_BYTE = 16;

/**
 * Spaces plain text may add to line up tables and page layouts for each byte of the document, beyond
 * the 16 million every document may (see TextGenerator): a workbook of a million cells lines its
 * columns up with about ten million.
 */
export const LAYOUT_SPACES_PER_BYTE = 16;

/**
 * Excel Spreadsheet (XLSX) Parser
 * 
 * **XLSX Format Overview:**
 * XLSX is the default format for Microsoft Excel since Office 2007, based on OOXML.
 * 
 * **File Structure:**
 * - `xl/workbook.xml` - Workbook structure and sheet list
 * - `xl/worksheets/sheet1.xml` - Individual sheet data
 * - `xl/sharedStrings.xml` - Shared string table (cell text)
 * - `xl/styles.xml` - Cell styling information
 * - `xl/drawings/*` - Charts and drawings
 * - `xl/media/*` - Embedded images
 * 
 * **Key Elements:**
 * - `<row>` - Table row with row index
 * - `<c r="A1">` - Cell with reference (A1, B2, etc.)
 * - `<v>` - Cell value (number or shared string index)
 * - `<t="s">` - Cell type (s=string, n=number, b=boolean)
 * 
 * @module ExcelParser
 * @see https://www.ecma-international.org/publications-and-standards/standards/ecma-376/
 */

import { ChartMetadata, FullOfficeParserConfig, ImageMetadata, OfficeAttachment, OfficeContentNode, OfficeParserAST, OfficeWarningType, PageSetup, TextFormatting, TextMetadata } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { extractChartData } from '../utils/chartUtils.js';
import { checkAbortSignal, logWarning } from '../utils/errorUtils.js';
import { createAttachment } from '../utils/imageUtils.js';
import { ocrDuringParse } from '../utils/ocrUtils.js';
import { attachmentLookup, repeatPreview, takeRepeats, valuesLength } from '../utils/repeatUtils.js';
import { documentBytesOf, TABLE_CELLS_PER_BYTE } from '../utils/budgetUtils.js';
import { chargeRawContent, decodeXmlEntities, getChildElements, getElementsByTagName, parseOfficeMetadata, parseOOXMLAppProperties, parseOOXMLCustomProperties, parseXmlString, takeNodes } from '../utils/xmlUtils.js';
import { MAX_COL_SPAN, MAX_ROW_SPAN } from '../utils/numberUtils.js';
import { BUILTIN_NUMBER_FORMATS, excelShownText } from '../utils/excelFormat.js';
import { extractFiles, findRequiredPart } from '../utils/zipUtils.js';
import { inRanges } from '../utils/sheetUtils.js';

/** Whether the character at `at` of `xml` can continue a tag name (so `<c` there opens `<col`, not `<c`). */
const continuesName = (xml: string, at: number): boolean => /[\w:.-]/.test(xml[at] ?? '');

/** Where the next `<tag` element starts in `xml`, from `from` on; -1 when none does. */
function nextElement(xml: string, tag: string, from: number): number {
    for (let at = xml.indexOf(`<${tag}`, from); at !== -1; at = xml.indexOf(`<${tag}`, at + tag.length + 1)) {
        if (!continuesName(xml, at + tag.length + 1)) return at;
    }
    return -1;
}

/**
 * Each `<tag>` element of `xml`, in order: its attributes, whether it closes itself, its content, and
 * the whole of it. The content runs to its closing tag or, left unclosed, to where the next element of
 * its kind starts (else the end), so an unclosed row keeps its cells. Linear in the length of `xml`:
 * the closing tag last found is kept while it lies ahead, where a pattern looking for one from each
 * start (`<row>([\s\S]*?)</row>`) read the rest of the sheet again for every unclosed row, which took
 * a minute for a few megabytes of them.
 */
function* xmlElements(xml: string, tag: string): Generator<{ attrs: string; selfClosing: boolean; content: string; whole: string }> {
    const close = `</${tag}>`;
    let closeAt = -2;
    let at = nextElement(xml, tag, 0);
    while (at !== -1) {
        const gt = xml.indexOf('>', at);
        if (gt === -1) return;
        const selfClosing = xml[gt - 1] === '/';
        const next = nextElement(xml, tag, gt + 1);
        if (selfClosing) {
            yield { attrs: xml.slice(at + tag.length + 1, gt - 1), selfClosing, content: '', whole: xml.slice(at, gt + 1) };
        } else {
            if (closeAt !== -1 && closeAt <= gt) closeAt = xml.indexOf(close, gt + 1);
            const closed = closeAt !== -1 && (next === -1 || closeAt < next);
            const end = closed ? closeAt : next === -1 ? xml.length : next;
            yield { attrs: xml.slice(at + tag.length + 1, gt), selfClosing, content: xml.slice(gt + 1, end), whole: xml.slice(at, closed ? end + close.length : end) };
        }
        at = next;
    }
}

/** The content of the first `<tag>` element of `xml` (see xmlElements), or undefined when there is none. */
function firstElementContent(xml: string, tag: string): string | undefined {
    for (const element of xmlElements(xml, tag)) return element.content;
    return undefined;
}

/** An attribute's value in the attributes of an element (see xmlElements), its references decoded. */
const attributeOf = (attrs: string, name: 'val' | 'rgb'): string | undefined => {
    const match = (name === 'val' ? /(?:^|\s)val\s*=\s*(?:"([^"]*)"|'([^']*)')/ : /(?:^|\s)rgb\s*=\s*(?:"([^"]*)"|'([^']*)')/).exec(attrs);
    return match ? decodeXmlEntities(match[1] ?? match[2]) : undefined;
};

/** An attribute's value by its (plain, unprefixed) name in the attributes of an element, or undefined. */
const attributeNamed = (attrs: string, name: string): string | undefined => {
    const match = new RegExp(`(?:^|\\s)${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(attrs);
    return match ? (match[1] ?? match[2]) : undefined;
};

/** Whether a boolean attribute's value is true (`1` or `true`, as the schema's xsd:boolean has it). */
const isTrue = (value: string | undefined): boolean => value === '1' || value === 'true';

/** A cell reference's 0-based column and row (`B3` is 1, 2), or null. */
function cellAt(ref: string): { col: number; row: number } | null {
    const match = /^\$?([A-Z]{1,3})\$?(\d{1,7})$/.exec(ref.trim());
    if (!match) return null;
    let col = 0;
    for (const ch of match[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
    return { col: col - 1, row: Number(match[2]) - 1 };
}

/**
 * A sheet's merged ranges (`<mergeCell ref="A1:F1"/>`) as the spans of their top-left cells, by `row:col`.
 * The cells a merge covers are left as read (a writer sets an empty one under the span).
 */
function mergeSpans(sheetXml: string): Map<string, { colSpan: number; rowSpan: number }> {
    const spans = new Map<string, { colSpan: number; rowSpan: number }>();
    const merges = firstElementContent(sheetXml, 'mergeCells');
    if (!merges) return spans;
    for (const merge of xmlElements(merges, 'mergeCell')) {
        const ref = /(?:^|\s)ref\s*=\s*["']([^"']*)["']/.exec(merge.attrs)?.[1];
        const [from, to] = (ref ?? '').split(':').map(cellAt);
        if (!from || !to || to.col < from.col || to.row < from.row) continue;
        spans.set(`${from.row}:${from.col}`, {
            colSpan: Math.min(to.col - from.col + 1, MAX_COL_SPAN),
            rowSpan: Math.min(to.row - from.row + 1, MAX_ROW_SPAN),
        });
    }
    return spans;
}

/**
 * A sheet's hyperlinks (`<hyperlink ref="A1" r:id="rId1"/>`, or `location="Sheet2!A1"` for a place in
 * the workbook) as the link of each cell their range covers, by `row:col`; at most MAX_LINKED_CELLS cells.
 */
function cellLinks(sheetXml: string, targets: Map<string, string>): Map<string, { link: string; linkType: 'internal' | 'external' }> {
    const links = new Map<string, { link: string; linkType: 'internal' | 'external' }>();
    const block = firstElementContent(sheetXml, 'hyperlinks');
    if (!block) return links;
    const attr = (attrs: string, name: string) => new RegExp(`(?:^|\\s)${name}\\s*=\\s*["']([^"']*)["']`).exec(attrs)?.[1];
    let visited = 0;
    for (const h of xmlElements(block, 'hyperlink')) {
        const id = attr(h.attrs, 'r:id'), location = attr(h.attrs, 'location');
        const target = id ? targets.get(id) : undefined;
        // (A relationship's target is read decoded already; an attribute of the sheet's text is not.)
        const value = target ? { link: target, linkType: 'external' as const } : location ? { link: `#${decodeXmlEntities(location)}`, linkType: 'internal' as const } : undefined;
        const [from, to] = (attr(h.attrs, 'ref') ?? '').split(':').map(cellAt);
        if (!value || !from) continue;
        const last = to ?? from;
        // Bounded by the positions visited, not the cells marked: a range repeated a thousand times marks
        // nothing new and visited its hundred thousand cells each time.
        for (let r = from.row; r <= last.row && visited < MAX_LINKED_CELLS; r++) {
            for (let c = from.col; c <= last.col && visited < MAX_LINKED_CELLS; c++, visited++) if (!links.has(`${r}:${c}`)) links.set(`${r}:${c}`, value);
        }
        if (visited >= MAX_LINKED_CELLS) break;
    }
    return links;
}

/** The most cell positions a sheet's hyperlinks visit (a range of a million cells is one link, not a million). */
const MAX_LINKED_CELLS = 100_000;

/** Excel's paper sizes (`pageSetup paperSize`) in points, upright: the common ones. */
const PAPER_SIZES: Record<string, [number, number]> = {
    '1': [612, 792], '3': [792, 1224], '4': [1224, 792], '5': [612, 1008], '8': [841.89, 1190.55], '9': [595.28, 841.89], '11': [419.53, 595.28], '70': [297.64, 419.53],
};

/**
 * How a sheet prints (see SheetMetadata.pageSetup): its `pageSetup`'s paper (Letter when it names none,
 * as Excel's default is) turned by its orientation, and its `pageMargins` (inches). Undefined for a
 * sheet that sets neither.
 */
function sheetPageSetup(sheetXml: string): PageSetup | undefined {
    // A custom view's own print setup (`customSheetViews`) is no part of the sheet's.
    const views = sheetXml.indexOf('<customSheetViews');
    if (views >= 0) {
        const close = sheetXml.indexOf('</customSheetViews>', views);
        sheetXml = sheetXml.slice(0, views) + (close >= 0 ? sheetXml.slice(close + '</customSheetViews>'.length) : '');
    }
    const setup = /<pageSetup\b[^>]*>/.exec(sheetXml)?.[0];
    const margins = /<pageMargins\b[^>]*>/.exec(sheetXml)?.[0];
    if (!setup && !margins) return undefined;
    const paper = PAPER_SIZES[setup ? attributeNamed(setup.slice(10), 'paperSize') ?? '1' : '1'] ?? PAPER_SIZES['1'];
    const landscape = setup ? attributeNamed(setup.slice(10), 'orientation') === 'landscape' : false;
    const [w, h] = landscape ? [Math.max(...paper), Math.min(...paper)] : [Math.min(...paper), Math.max(...paper)];
    const result: PageSetup = { width: w, height: h };
    if (margins) {
        const inch = (name: string) => { const v = Number(attributeNamed(margins.slice(12), name)); return Number.isFinite(v) && v >= 0 && v < 22 ? Math.round(v * 7200) / 100 : undefined; };
        const [top, right, bottom, left] = ['top', 'right', 'bottom', 'left'].map(inch);
        if (top !== undefined && right !== undefined && bottom !== undefined && left !== undefined) result.margin = { top, right, bottom, left };
        const header = inch('header'), footer = inch('footer');
        if (header !== undefined) result.headerDistance = header;
        if (footer !== undefined) result.footerDistance = footer;
    }
    return result;
}

/**
 * The 0-based columns a sheet hides, as ranges [first, last]: its `<cols>`'s `<col>` elements with
 * `hidden`, each covering `min` to `max` (1-based).
 */
function hiddenColumnRanges(sheetXml: string): [number, number][] {
    const cols = firstElementContent(sheetXml, 'cols');
    if (!cols) return [];
    const ranges: [number, number][] = [];
    for (const col of xmlElements(cols, 'col')) {
        if (!isTrue(attributeNamed(col.attrs, 'hidden'))) continue;
        const min = Number(attributeNamed(col.attrs, 'min')), max = Number(attributeNamed(col.attrs, 'max'));
        if (Number.isInteger(min) && min >= 1) ranges.push([min - 1, (Number.isInteger(max) && max >= min ? max : min) - 1]);
    }
    return ranges;
}

/**
 * The text an element's XML content holds, as an XML reader gives it: references decoded, line breaks
 * normalized, CDATA sections as written, and markup inside it left out.
 */
const xmlText = (content: string): string => {
    const decode = (raw: string) => decodeXmlEntities(raw.includes('\r') ? raw.replace(/\r\n?/g, '\n') : raw);
    if (!content.includes('<')) return decode(content);
    const parts: string[] = [];
    let at = 0;
    while (at < content.length) {
        const lt = content.indexOf('<', at);
        if (lt === -1) { parts.push(decode(content.slice(at))); break; }
        parts.push(decode(content.slice(at, lt)));
        if (content.startsWith('<![CDATA[', lt)) {
            const end = content.indexOf(']]>', lt + 9);
            parts.push(content.slice(lt + 9, end === -1 ? content.length : end));
            at = end === -1 ? content.length : end + 3;
        } else {
            const gt = content.indexOf('>', lt + 1);
            at = gt === -1 ? content.length : gt + 1;
        }
    }
    return parts.join('');
};

/**
 * A rich text run's formatting from its `rPr` content. A toggle (`b`, `i`, `strike`) is on unless its
 * `val` turns it off (`<b val="0"/>`), and `u` unless it is `none`.
 */
function runFormatting(rPr: string): TextFormatting {
    const formatting: TextFormatting = {};
    const first = (tag: string) => { for (const element of xmlElements(rPr, tag)) return element; return undefined; };
    const on = (tag: string, off: RegExp): boolean => {
        const element = first(tag);
        if (!element) return false;
        const val = attributeOf(element.attrs, 'val');
        return val === undefined || !off.test(val);
    };
    if (on('b', /^(0|false|off)$/i)) formatting.bold = true;
    if (on('i', /^(0|false|off)$/i)) formatting.italic = true;
    if (on('u', /^none$/i)) formatting.underline = true;
    if (on('strike', /^(0|false|off)$/i)) formatting.strikethrough = true;
    const size = first('sz') && attributeOf(first('sz')!.attrs, 'val');
    if (size) formatting.size = size + 'pt';
    const color = first('color');
    const rgb = color && attributeOf(color.attrs, 'rgb');
    // ARGB (as Excel writes it), or RGB.
    if (rgb && /^[0-9A-Fa-f]{6}(?:[0-9A-Fa-f]{2})?$/.test(rgb)) formatting.color = '#' + (rgb.length === 8 ? rgb.slice(2) : rgb);
    const font = first('rFont');
    const fontName = font && attributeOf(font.attrs, 'val');
    if (fontName) formatting.font = fontName;
    const vertAlign = first('vertAlign');
    const position = vertAlign && attributeOf(vertAlign.attrs, 'val');
    if (position === 'subscript') formatting.subscript = true;
    if (position === 'superscript') formatting.superscript = true;
    return formatting;
}

/** How a number format shows a cell value that it treats as a date or a time. */
type DateFormatKind = 'date' | 'time' | 'datetime' | 'elapsed-h' | 'elapsed-m' | 'elapsed-s';

/** Built-in number formats that show a date or a time (ECMA-376 Part 1, 18.8.30). */
const BUILTIN_DATE_FORMATS = new Map<string, DateFormatKind>([
    ['14', 'date'], ['15', 'date'], ['16', 'date'], ['17', 'date'], ['18', 'time'], ['19', 'time'],
    ['20', 'time'], ['21', 'time'], ['22', 'datetime'], ['45', 'time'], ['46', 'elapsed-h'], ['47', 'time'],
]);

/** 9999-12-31, the last day a workbook can hold, as a day count in the 1900 and 1904 date systems. */
const LAST_DAY_1900 = 2958465;
const LAST_DAY_1904 = 2957003;

/**
 * Classifies a number format code by the date and time parts it shows. Only the first section
 * counts, the one a positive number uses. Quoted text, escaped characters, padding (`_x`) and fill
 * (`*x`) characters and `[...]` sections (colours, locales, conditions) are not tokens, except
 * `[h]`, `[m]` and `[s]`, which show elapsed time in total hours, minutes or seconds. `AM/PM` and
 * `A/P` show a time. One forward pass, so a crafted code cannot slow it down.
 */
const getDateFormatKind = (code: string): DateFormatKind | undefined => {
    const lower = code.toLowerCase();
    let date = false, time = false, month = false;
    const elapsed = { h: false, m: false, s: false };
    for (let i = 0; i < lower.length; i++) {
        const c = lower[i];
        if (c === ';') break;
        if (c === '\\' || c === '_' || c === '*') {
            i++;
        } else if (c === '"' || c === '[') {
            const end = lower.indexOf(c === '"' ? '"' : ']', i + 1);
            if (end < 0) break;
            const unit = lower[i + 1];
            if (c === '[' && end > i + 1 && (unit === 'h' || unit === 'm' || unit === 's')) {
                let same = true;
                for (let j = i + 2; j < end && same; j++) same = lower[j] === unit;
                if (same) elapsed[unit] = true;
            }
            i = end;
        } else if (lower.startsWith('am/pm', i)) {
            time = true;
            i += 4;
        } else if (lower.startsWith('a/p', i)) {
            time = true;
            i += 2;
        } else if (c === 'y' || c === 'd') {
            date = true;
        } else if (c === 'h' || c === 's') {
            time = true;
        } else if (c === 'm') {
            month = true;
        }
    }
    if (elapsed.h) return 'elapsed-h';
    if (elapsed.m) return 'elapsed-m';
    if (elapsed.s) return 'elapsed-s';
    // An `m` is a minute next to an hour or a second, and a month otherwise (`mmm yyyy`, `mmmm`)
    if (date || (month && !time)) return time ? 'datetime' : 'date';
    return time ? 'time' : undefined;
};

/**
 * Writes a date/time serial as ISO text in the form its number format shows: `2024-09-30`,
 * `14:05:00`, `2024-09-30 14:05:00`, or an elapsed duration counted in its unit (`36:00:00`,
 * `2160:00`, `129600`). Returns `undefined` for a value that is not a serial a workbook can hold,
 * which then stays as stored.
 */
const formatDateSerial = (value: string, kind: DateFormatKind, date1904: boolean): string | undefined => {
    const serial = Number(value);
    if (value === '' || !Number.isFinite(serial) || serial < 0) return undefined;
    const pad = (n: number) => String(n).padStart(2, '0');
    // Rounded to the second, as Excel shows it: a NOW() stamp of 14:04:59.9 reads 14:05:00
    const seconds = Math.round(serial * 86400);
    const days = Math.floor(seconds / 86400);
    // Past 9999-12-31 a value shows as a plain number, judged as shown, after the rounding
    if (days > (date1904 ? LAST_DAY_1904 : LAST_DAY_1900)) return undefined;
    if (kind === 'elapsed-s') return String(seconds);
    if (kind === 'elapsed-m') return `${Math.floor(seconds / 60)}:${pad(seconds % 60)}`;
    const clock = (total: number) => `${pad(Math.floor(total / 60) % 60)}:${pad(total % 60)}`;
    if (kind === 'elapsed-h') return `${Math.floor(seconds / 3600)}:${clock(seconds)}`;
    const time = `${pad(Math.floor(seconds / 3600) % 24)}:${clock(seconds)}`;
    if (kind === 'time') return time;
    // The 1904 system counts from 1904-01-01. The 1900 system counts from 1899-12-31 and includes a
    // 1900-02-29 that never existed, as day 60, so days after it count from 1899-12-30.
    const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, days < 60 ? 31 : 30);
    const date = !date1904 && days === 60
        ? '1900-02-29'
        : new Date(epoch + days * 86400000).toISOString().slice(0, 10);
    return kind === 'date' ? date : `${date} ${time}`;
};

/**
 * Parses an Excel spreadsheet (.xlsx) and extracts sheets, rows, and cells.
 * 
 * @param buffer - The XLSX file as a Buffer
 * @param config - Parser configuration
 * @returns A promise resolving to the parsed AST
 */
export const parseExcel = async (buffer: Buffer, config: FullOfficeParserConfig): Promise<OfficeParserAST> => {
    // Honour cancellation requests immediately, before extracting the ZIP archive.
    // XLSX parsing involves decompressing multiple XML sheets and potentially running OCR
    // on embedded chart images, so short-circuiting here saves significant work.
    checkAbortSignal(config.abortSignal);

    const sheetsRegex = /xl\/worksheets\/sheet\d+.xml/g;
    const drawingsRegex = /xl\/drawings\/drawing\d+.xml/g;
    const chartsRegex = /xl\/charts\/chart\d+.xml/g;
    const stringsFilePath = 'xl/sharedStrings.xml';
    const mediaFileRegex = /xl\/media\/.*/;
    const corePropsFileRegex = /docProps\/core\.xml/;
    const customPropsFileRegex = /docProps\/custom\.xml/;
    const appPropsFileRegex = /docProps\/app\.xml/;

    const relsRegex = /xl\/worksheets\/_rels\/sheet\d+\.xml\.rels/g;
    const drawingRelsRegex = /xl\/drawings\/_rels\/drawing\d+\.xml\.rels/g;
    const commentsRegex = /xl\/comments\d+\.xml/g;

    const files = await extractFiles(
        buffer,
        (x: string) =>
            !!x.match(sheetsRegex) ||
            !!x.match(drawingsRegex) ||
            !!x.match(chartsRegex) ||
            (!config.ignoreComments && !!x.match(commentsRegex)) ||
            x === stringsFilePath ||
            x === 'xl/styles.xml' ||
            x === 'xl/workbook.xml' ||
            x === 'xl/_rels/workbook.xml.rels' ||
            !!x.match(corePropsFileRegex) ||
            !!x.match(customPropsFileRegex) ||
            !!x.match(appPropsFileRegex) ||
            (!!config.extractAttachments && (!!x.match(mediaFileRegex) || !!x.match(drawingRelsRegex))) ||
            !!x.match(relsRegex),
        config.decompressionLimits,
        config
    );

    // Every workbook has xl/workbook.xml; without it the archive is not a spreadsheet.
    // Resolved up front so a file that cannot be a workbook fails before any of the parsing
    // work below, and read again further down for the sheet-name map.
    const workbookFile = findRequiredPart(files, path => path === 'xl/workbook.xml', config,
        { fileType: 'xlsx', part: 'xl/workbook.xml' });
    // Date serials count from 1900 unless the workbook uses the 1904 date system
    const date1904 = /<(?:\w+:)?workbookPr\b[^>]*\bdate1904=["'](?:1|true)["']/.test(workbookFile.content.toString());

    // Worksheets, by contrast, are not guaranteed: a workbook holding only chartsheets is
    // valid and simply has no cell text to extract. Warn rather than fail, so the caller can
    // tell "nothing to read here" from "we read nothing".
    if (!files.some(file => !!file.path.match(sheetsRegex)))
        logWarning(OfficeWarningType.NO_WORKSHEETS_FOUND, config);

    const sharedStringsFile = files.find(f => f.path === stringsFilePath);
    // Updated to store structured content (rich text runs) or simple string
    const sharedStrings: (string | OfficeContentNode[])[] = [];
    // Cells the document's sheets may still yield (see the sheet loop).
    // One cell more for each byte of the document (see budgetUtils), so a large workbook is not cut short.
    const cellLimit = (config.decompressionLimits?.maxTableCells ?? 1000000) + TABLE_CELLS_PER_BYTE * documentBytesOf(config);
    let cellsLeft = cellLimit;
    let cellLimitWarned = false;
    // The text of each rich shared string, joined once: joined per cell, a large string many cells
    // show was copied for each, and a small file filled the heap.
    const richStringText = new Map<number, string>();
    // How many cells have shown each shared string so far.
    const sharedStringUses = new Map<number, number>();
    // What each rich shared string's runs weigh beyond their text: 16 a run, and its formatting values.
    // Cells showing the string share its runs, and every writer writes a run's font or colour again in
    // each cell: counted by run alone, one run of a 64 KB font shown in 7,500 cells (1.8 KB of XLSX)
    // made 492 MB of HTML.
    const richStringRunWeight = new Map<number, number>();

    if (sharedStringsFile) {
        // Read as the sheets are, without building XML elements: a workbook of a million distinct
        // strings (5.5 MB) made two million elements, a gigabyte of heap, and was refused by the element
        // budget, where each string is only its text. A rich string's runs become nodes, so they are
        // charged to the element budget as elements are; a plain string costs an array entry, and eight
        // of them are charged as one element (an empty `<si/>` is five bytes of a zip's XML).
        let plainStrings = 0;
        for (const si of xmlElements(sharedStringsFile.content.toString(), 'si')) {
            // A string's text and runs come before its phonetic readings (`rPh`), which are not its text.
            const phonetic = nextElement(si.content, 'rPh', 0);
            const body = phonetic === -1 ? si.content : si.content.slice(0, phonetic);
            const runs: OfficeContentNode[] = [];
            let runCount = 0;
            for (const run of xmlElements(body, 'r')) {
                runCount++;
                const t = firstElementContent(run.content, 't');
                if (t === undefined) continue;
                const rPr = firstElementContent(run.content, 'rPr');
                const formatting = rPr === undefined ? {} : runFormatting(rPr);
                runs.push({ type: 'text', text: xmlText(t), formatting: Object.keys(formatting).length > 0 ? formatting : undefined });
            }
            if (runCount > 0) {
                takeNodes(runCount, config);
                sharedStrings.push(runs);
            } else {
                if (++plainStrings % 8 === 0) takeNodes(1, config);
                let text = '';
                for (const t of xmlElements(body, 't')) text += xmlText(t.content);
                sharedStrings.push(text);
            }
        }
    }

    // Parse styles to build formatting map
    const stylesFile = files.find(f => f.path === 'xl/styles.xml');
    const cellFormatMap: Record<number, TextFormatting> = {};
    // A date or a time is stored as a day count and shown through its style's number format
    const cellDateFormatMap: Record<number, DateFormatKind> = {};
    // Each cell style's number format code, which writes the text the sheet shows (see excelShownText).
    const cellNumberFormatMap: Record<number, string> = {};

    if (stylesFile) {
        const xml = parseXmlString(stylesFile.content.toString(), { config });

        // Parse custom number formats (numFmtId -> format code)
        // Classified once per number format, however many cell formats share it.
        const numFmtKinds = new Map<string, DateFormatKind | undefined>();
        const numFmtCodes = new Map<string, string>(BUILTIN_NUMBER_FORMATS);
        const numFmtsNode = getElementsByTagName(xml, "numFmts")[0];
        if (numFmtsNode) {
            for (const numFmt of getChildElements(numFmtsNode, "numFmt")) {
                const id = numFmt.getAttribute("numFmtId");
                const code = numFmt.getAttribute("formatCode");
                if (id && code !== null) { numFmtKinds.set(id, getDateFormatKind(code)); numFmtCodes.set(id, code); }
            }
        }

        // Parse fonts
        const fontsNode = getElementsByTagName(xml, "fonts")[0];
        const fonts: TextFormatting[] = [];
        if (fontsNode) {
            // Each font, fill and format read from its own children (see getChildElements).
            const fontNodes = getChildElements(fontsNode, "font");
            for (const font of fontNodes) {
                const formatting: TextFormatting = {};
                if (getChildElements(font, "b").length > 0) formatting.bold = true;
                if (getChildElements(font, "i").length > 0) formatting.italic = true;
                if (getChildElements(font, "u").length > 0) formatting.underline = true;
                if (getChildElements(font, "strike").length > 0) formatting.strikethrough = true;

                const szNode = getChildElements(font, "sz")[0];
                if (szNode) {
                    const val = szNode.getAttribute("val");
                    if (val) formatting.size = val + 'pt';
                }

                const colorNode = getChildElements(font, "color")[0];
                if (colorNode) {
                    const rgb = colorNode.getAttribute("rgb");
                    if (rgb) formatting.color = '#' + rgb.substring(2); // Remove alpha channel
                }

                const nameNode = getChildElements(font, "name")[0];
                if (nameNode) {
                    const val = nameNode.getAttribute("val");
                    if (val) formatting.font = val;
                }

                const vertAlignNode = getChildElements(font, "vertAlign")[0];
                if (vertAlignNode) {
                    const val = vertAlignNode.getAttribute("val");
                    if (val === "subscript") formatting.subscript = true;
                    if (val === "superscript") formatting.superscript = true;
                }

                fonts.push(formatting);
            }
        }

        // Parse fills (for background color)
        const fillsNode = getElementsByTagName(xml, "fills")[0];
        const fills: TextFormatting[] = [];
        if (fillsNode) {
            const fillNodes = getChildElements(fillsNode, "fill");
            for (const fill of fillNodes) {
                const formatting: TextFormatting = {};
                const patternFill = getChildElements(fill, "patternFill")[0];
                if (patternFill) {
                    const fgColor = getChildElements(patternFill, "fgColor")[0];
                    if (fgColor) {
                        const rgb = fgColor.getAttribute("rgb");
                        const theme = fgColor.getAttribute("theme");

                        if (rgb && rgb !== "00000000") { // Not default/auto
                            formatting.backgroundColor = '#' + rgb.substring(2);
                        } else if (theme) {
                            // Basic mapping for standard Office themes (Dark 1, Light 1, Dark 2, Light 2)
                            // 0: Light 1 (White), 1: Dark 1 (Black), 2: Light 2 (Tan/Gray), 3: Dark 2 (Blue/Grey)
                            const themeIdx = parseInt(theme);
                            if (themeIdx === 0) formatting.backgroundColor = '#FFFFFF';
                            else if (themeIdx === 1) formatting.backgroundColor = '#000000';
                            else if (themeIdx === 2) formatting.backgroundColor = '#EEECE1'; // Standard Light 2
                            else if (themeIdx === 3) formatting.backgroundColor = '#1F497D'; // Standard Dark 2
                        }
                    }
                }
                fills.push(formatting);
            }
        }

        // Parse cellXfs (cell format definitions)
        const cellXfsNode = getElementsByTagName(xml, "cellXfs")[0];
        if (cellXfsNode) {
            const xfNodes = getChildElements(cellXfsNode, "xf");
            for (let i = 0; i < xfNodes.length; i++) {
                const xf = xfNodes[i];
                const formatting: TextFormatting = {};

                const numFmtId = xf.getAttribute("numFmtId") || '0';
                const dateFormatKind = numFmtKinds.has(numFmtId)
                    ? numFmtKinds.get(numFmtId)
                    : BUILTIN_DATE_FORMATS.get(numFmtId);
                if (dateFormatKind) cellDateFormatMap[i] = dateFormatKind;
                const formatCode = numFmtCodes.get(numFmtId);
                if (formatCode !== undefined) cellNumberFormatMap[i] = formatCode;

                const fontId = xf.getAttribute("fontId");
                if (fontId) {
                    const fontIdx = parseInt(fontId);
                    if (fonts[fontIdx]) {
                        Object.assign(formatting, fonts[fontIdx]);
                    }
                }

                const fillId = xf.getAttribute("fillId");
                if (fillId) {
                    const fillIdx = parseInt(fillId);
                    if (fills[fillIdx] && fills[fillIdx].backgroundColor) {
                        formatting.backgroundColor = fills[fillIdx].backgroundColor;
                    }
                }

                const alignmentNode = getChildElements(xf, "alignment")[0];
                if (alignmentNode) {
                    const horizontal = alignmentNode.getAttribute("horizontal");
                    if (horizontal === 'center' || horizontal === 'right' || horizontal === 'justify' || horizontal === 'left') {
                        formatting.alignment = horizontal;
                    }
                }

                cellFormatMap[i] = formatting;
            }
        }
    }

    const attachments: OfficeAttachment[] = [];
    const attachmentsByName = attachmentLookup(attachments);
    // Drawings already placed on a sheet (see the sheet loop).
    const drawingsPlaced = new Set<string>();
    const mediaFiles = files.filter(f => f.path.match(/xl\/media\/.*/));
    const chartFiles = files.filter(f => f.path.match(chartsRegex));

    // Map to store image details by drawing file path and relationship ID
    // Null-prototype: keyed by document-derived drawing paths / rel ids, so a plain object would let a
    // crafted `__proto__` key throw (or hit a shared prototype) instead of creating an own entry.
    const drawingImageMap: Record<string, Record<string, { path: string, altText?: string }>> = Object.create(null);

    if (config.extractAttachments) {
        // 1. Parse Drawing Rels to map rIds to media paths
        const drawingRelsFiles = files.filter(f => f.path.match(drawingRelsRegex));
        for (const relFile of drawingRelsFiles) {
            const drawingFilename = relFile.path.split('/').pop()?.replace('.rels', '') || '';
            const drawingPath = `xl/drawings/${drawingFilename}`;

            const relsXml = parseXmlString(relFile.content.toString(), { config });
            const relationships = getElementsByTagName(relsXml, "Relationship");

            if (!drawingImageMap[drawingPath]) {
                drawingImageMap[drawingPath] = Object.create(null);
            }

            for (const rel of relationships) {
                const id = rel.getAttribute("Id");
                const target = rel.getAttribute("Target");
                if (id && target && target.includes('media/')) {
                    // Target is usually like "../media/image1.png"
                    const mediaPath = 'xl/' + target.replace('../', '');
                    drawingImageMap[drawingPath][id] = { path: mediaPath };
                }
            }
        }

        // 2. Parse Drawings to get Alt Text and link to Rels
        const drawingFiles = files.filter(f => f.path.match(drawingsRegex));
        for (const drawingFile of drawingFiles) {
            const xml = parseXmlString(drawingFile.content.toString(), { config });
            const pics = getElementsByTagName(xml, "xdr:pic"); // SpreadsheetML drawing

            // Null-prototype, as the map is: a plain `{}` answered an `r:embed` of `__proto__` with
            // Object.prototype, and the alt text was written onto every object in the process.
            const rels: Record<string, { path: string, altText?: string }> = drawingImageMap[drawingFile.path] || Object.create(null);

            for (const pic of pics) {
                const blipFill = getElementsByTagName(pic, "xdr:blipFill")[0];
                const blip = blipFill ? getElementsByTagName(blipFill, "a:blip")[0] : null;
                const embedId = blip ? blip.getAttribute("r:embed") : null;

                const nvPicPr = getElementsByTagName(pic, "xdr:nvPicPr")[0];
                const cNvPr = nvPicPr ? getElementsByTagName(nvPicPr, "xdr:cNvPr")[0] : null;
                const altText = cNvPr ? (cNvPr.getAttribute("descr") || cNvPr.getAttribute("name")) : undefined;

                if (embedId && rels[embedId]) {
                    rels[embedId].altText = altText || '';
                }
            }
        }

        // 3. Process Media Files
        // Each picture's alt text, by its media path: the first drawing showing it with one (each
        // drawing by its first relationship to it). Looked up per media file through every drawing's
        // relationships, 2,000 media files beside 200,000 relationships took 90 seconds.
        const altTextByMedia = new Map<string, string>();
        for (const drawingPath in drawingImageMap) {
            const seenInDrawing = new Set<string>();
            for (const rId in drawingImageMap[drawingPath]) {
                const image = drawingImageMap[drawingPath][rId];
                if (seenInDrawing.has(image.path)) continue;
                seenInDrawing.add(image.path);
                if (image.altText && !altTextByMedia.has(image.path)) altTextByMedia.set(image.path, image.altText);
            }
        }
        for (const media of mediaFiles) {
            const attachment = createAttachment(media.path.split('/').pop() || 'image', media.content);

            const altText = altTextByMedia.get(media.path);
            if (altText) attachment.altText = altText;

            attachments.push(attachment);

            if (config.ocr && attachment.mimeType.startsWith('image/')) {
                const ocrText = await ocrDuringParse(media.content, config, attachment.name);
                if (ocrText) attachment.ocrText = ocrText;
            }
        }

        for (const chart of chartFiles) {
            const attachment: OfficeAttachment = {
                type: 'chart',
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                data: chart.content.toString('base64'),
                name: chart.path.split('/').pop() || '',
                extension: 'xml'
            };

            // Extract structured chart data
            try {
                const chartData = extractChartData(chart.content, config);
                attachment.chartData = chartData;
            } catch (e) {
                logWarning(OfficeWarningType.CHART_DATA_EXTRACTION_FAILED, config, chart.path, e);
            }

            attachments.push(attachment);
        }
    }

    // Build map of drawing rId -> chart attachment name for linking
    // Null-prototype for the same reason as drawingImageMap (document-derived keys).
    const drawingChartMap: Record<string, Record<string, string>> = Object.create(null);
    if (config.extractAttachments) {
        const drawingRelsFiles = files.filter(f => f.path.match(drawingRelsRegex));
        for (const relFile of drawingRelsFiles) {
            const drawingFilename = relFile.path.split('/').pop()?.replace('.rels', '') || '';
            const drawingPath = `xl/drawings/${drawingFilename}`;

            const relsXml = parseXmlString(relFile.content.toString(), { config });
            const relationships = getElementsByTagName(relsXml, "Relationship");

            if (!drawingChartMap[drawingPath]) {
                drawingChartMap[drawingPath] = Object.create(null);
            }

            for (const rel of relationships) {
                const id = rel.getAttribute("Id");
                const target = rel.getAttribute("Target");
                const type = rel.getAttribute("Type");
                if (id && target && type && type.includes('chart')) {
                    // Target is like "../charts/chart1.xml"
                    const chartName = target.split('/').pop() || '';
                    drawingChartMap[drawingPath][id] = chartName;
                }
            }
        }
    }

    // Parse workbook.xml to get sheet names and map them to sheet files
    // Null-prototype, as every map keyed by the document's own names and ids is here.
    const sheetNameMap: Record<string, string> = Object.create(null);
    // The sheets the workbook hides (`state="hidden"` or `"veryHidden"`), by file.
    const hiddenSheets = new Set<string>();
    const workbookRelsFile = files.find(f => f.path === 'xl/_rels/workbook.xml.rels');

    if (workbookRelsFile) {
        // Parse rels to get rId -> file mapping
        const relsXml = parseXmlString(workbookRelsFile.content.toString(), { config });
        const relationships = getElementsByTagName(relsXml, "Relationship");
        const rIdToFile: Record<string, string> = Object.create(null);

        for (const rel of relationships) {
            const rId = rel.getAttribute("Id");
            const target = rel.getAttribute("Target");
            if (rId && target) {
                // Target is like "worksheets/sheet1.xml"
                const filename = target.split('/').pop() || '';
                rIdToFile[rId] = filename;
            }
        }

        // Parse workbook.xml to get sheet name -> rId mapping
        const workbookXml = parseXmlString(workbookFile.content.toString(), { config });
        const sheets = getElementsByTagName(workbookXml, "sheet");

        for (const sheet of sheets) {
            checkAbortSignal(config.abortSignal);
            const name = sheet.getAttribute("name");
            const rId = sheet.getAttribute("r:id");
            if (name && rId && rIdToFile[rId]) {
                sheetNameMap[rIdToFile[rId]] = name;
                const state = sheet.getAttribute("state");
                if (state === 'hidden' || state === 'veryHidden') hiddenSheets.add(rIdToFile[rId]);
            }
        }
    }

    const content: OfficeContentNode[] = [];
    // The package's parts by path (a scan of all of them per sheet and per relationship took sheets x
    // parts), and the comments parts already attached: each is read, and its comments attached, once
    // (to the first sheet naming it). Read again for each relationship naming it, one comments part of
    // 1 MB named 2,000 times (3.7 KB of XLSX) ended the process out of memory.
    const fileByPath = new Map<string, (typeof files)[number]>();
    for (const f of files) if (!fileByPath.has(f.path)) fileByPath.set(f.path, f);
    const attachedCommentParts = new Set<string>();

    for (const file of files) {
        if (file.path.match(mediaFileRegex)) continue;
        if (file.path === stringsFilePath) continue;
        if (file.path === 'xl/styles.xml') continue;
        if (file.path.match(drawingsRegex)) continue;
        if (file.path.match(chartsRegex)) continue;
        if (file.path.match(relsRegex)) continue;
        if (file.path.match(drawingRelsRegex)) continue;

        if (file.path.match(sheetsRegex)) {
            const sheetFilename = file.path.split('/').pop() || '';
            const relsFilename = `xl/worksheets/_rels/${sheetFilename}.rels`;
            const relsFile = fileByPath.get(relsFilename);

            const drawingMap: Record<string, string> = Object.create(null); // rId -> drawingPath
            // Null-prototype: keyed by the document-derived cell ref, so a crafted ref of `__proto__`
            // creates an own entry instead of throwing on `Object.prototype.push`.
            const sheetCommentsMap: Record<string, OfficeContentNode[]> = Object.create(null);
            // The sheet's hyperlinks' targets by relationship id (see cellLinks).
            const hyperlinkTargets = new Map<string, string>();

            if (relsFile) {
                const relsXml = parseXmlString(relsFile.content.toString(), { config });
                const relationships = getElementsByTagName(relsXml, "Relationship");
                for (const rel of relationships) {
                    const id = rel.getAttribute("Id");
                    const target = rel.getAttribute("Target");
                    const type = rel.getAttribute("Type");
                    
                    if (id && target && type) {
                        if (type.endsWith('/hyperlink')) hyperlinkTargets.set(id, target);
                        else if (config.extractAttachments && type.includes('drawing')) {
                            drawingMap[id] = 'xl/drawings/' + target.replace('../drawings/', '');
                        } else if (!config.ignoreComments && type.includes('comments')) {
                            const commentsPath = 'xl/' + target.replace('../', '');
                            const cFile = attachedCommentParts.has(commentsPath) ? undefined : fileByPath.get(commentsPath);
                            if (cFile) {
                                attachedCommentParts.add(commentsPath);
                                const cXml = parseXmlString(cFile.content.toString(), { config });
                                const commentNodes = getElementsByTagName(cXml, "comment");
                                const authorsList = getElementsByTagName(cXml, "author");
                                const authors = authorsList.map(a => a.textContent || '');
                                
                                for (const cNode of commentNodes) {
                                    const ref = cNode.getAttribute("ref");
                                    const authorId = cNode.getAttribute("authorId");
                                    const author = authorId !== null ? authors[parseInt(authorId)] : undefined;
                                    
                                    const tNodes = getElementsByTagName(cNode, "t");
                                    const text = tNodes.map(t => t.textContent || '').join('');
                                    
                                    if (ref && text) {
                                        if (!sheetCommentsMap[ref]) sheetCommentsMap[ref] = [];
                                        sheetCommentsMap[ref].push({
                                            type: 'comment',
                                            text: text,
                                            children: [{ type: 'text', text: text, formatting: {} }],
                                            metadata: author ? { author } : undefined
                                        });
                                    }
                                }
                            }
                        }
                    }
                }
            }

            const rows: OfficeContentNode[] = [];
            const sheetXml = file.content.toString();
            // The <row> elements as they stand, rather than a walk over the (possibly vast, sparse)
            // row range the sheet's dimensions declare.
            const rowElements = xmlElements(sheetXml, 'row');
            // Hidden columns and rows are read, and marked (CellMetadata.hidden): a printout leaves them
            // out, as the PDF writer does, and every other output keeps them as data.
            const hiddenColumns = hiddenColumnRanges(sheetXml);
            // Merged ranges, as the spans of their first cells: a merged title was one narrow cell.
            const merged = mergeSpans(sheetXml);
            // Its cells' hyperlinks (a web or mail address, or a place in the workbook), by `row:col`.
            const links = cellLinks(sheetXml, hyperlinkTargets);
            const columnHidden = inRanges(hiddenColumns);

            /** Helper to convert Excel column string (A, B, AA, etc.) to 0-based index */
            const colToNumber = (col: string): number => {
                let num = 0;
                for (let i = 0; i < col.length; i++) {
                    num = num * 26 + (col.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
                }
                return num - 1;
            };

            let lastRowIndex = -1;

            for (const rowElement of rowElements) {
                checkAbortSignal(config.abortSignal);
                if (cellsLeft <= 0) {
                    // Spent exactly at the row before: this row's cells are the ones not read.
                    if (!cellLimitWarned && /<(?:\w+:)?c[\s/>]/.test(rowElement.whole)) { cellLimitWarned = true; logWarning(OfficeWarningType.TABLE_CELL_LIMIT_EXCEEDED, config, cellLimit); }
                    break;
                }
                const rowXml = rowElement.whole;
                const rowAttrs = rowElement.attrs;
                const isSelfClosing = rowElement.selfClosing;
                const rowContent = rowElement.content;

                if (!isSelfClosing && !rowContent.includes('<c')) continue;

                const cells: OfficeContentNode[] = [];
                // The row's <c> (cell) elements: attributes (r="A1", t="s") and content (<v>, <is>).
                const cellElements = xmlElements(rowContent, 'c');

                const rowHidden = isTrue(attributeNamed(rowAttrs, 'hidden'));
                const rMatch = rowAttrs.match(/r="(\d+)"/);
                const rowIndex = rMatch ? parseInt(rMatch[1]) - 1 : lastRowIndex + 1;
                lastRowIndex = rowIndex;

                let lastColIndex = -1;

                for (const cellElement of cellElements) {
                    // The document's cells within maxTableCells, as ODF's are: a sheet is read without
                    // building XML elements, so the element budget does not bound it, and 16 million
                    // cells (854 KB of zip) ran the process out of memory.
                    if (--cellsLeft < 0) {
                        if (!cellLimitWarned) { cellLimitWarned = true; logWarning(OfficeWarningType.TABLE_CELL_LIMIT_EXCEEDED, config, cellLimit); }
                        break;
                    }
                    const cXml = cellElement.whole;
                    const cAttrs = cellElement.attrs;
                    const cContent = cellElement.content;

                    // Extract cell value
                    const typeMatch = cAttrs.match(/t="([a-zA-Z]+)"/);
                    const type = typeMatch ? typeMatch[1] : 'n'; // n = number (default)

                    // Extract cell style index
                    const styleMatch = cAttrs.match(/s="(\d+)"/);
                    const styleIdx = styleMatch ? parseInt(styleMatch[1]) : undefined;

                    const vContent = firstElementContent(cContent, 'v');
                    const isContent = type === 'inlineStr' ? firstElementContent(cContent, 'is') : undefined;

                    let text = '';
                    // What the sheet shows for a number, through its style's number format, where that
                    // differs from the value read (see CellMetadata.shownText).
                    let shownText: string | undefined;
                    let cellNodes: OfficeContentNode[] = [];

                    // Shared rich-text runs, merged with this cell's style below (the runs themselves are shared).
                    let sharedRuns: OfficeContentNode[] | undefined;
                    if (type === 's' && vContent !== undefined) {
                        const idx = parseInt(vContent);
                        const content = sharedStrings[idx];
                        let joined = typeof content === 'string' ? content : richStringText.get(idx);
                        if (joined === undefined && Array.isArray(content)) richStringText.set(idx, joined = content.map(n => n.text).join(''));
                        // Each cell after the first showing a string writes it again, so those are charged
                        // to the document's repeated-content budget (see repeatUtils): 2.7 KB showing one
                        // 1 MB string in 400 cells made 400 MB of CSV. Past it a cell shows the string's start.
                        const shown = sharedStringUses.get(idx) ?? 0;
                        sharedStringUses.set(idx, shown + 1);
                        let runWeight = Array.isArray(content) ? richStringRunWeight.get(idx) : 0;
                        if (runWeight === undefined && Array.isArray(content)) {
                            runWeight = 0;
                            for (const run of content) runWeight += 16 + valuesLength(run.formatting) + valuesLength(run.metadata);
                            richStringRunWeight.set(idx, runWeight);
                        }
                        const weight = (joined?.length ?? 0) + (runWeight ?? 0);
                        if (shown > 0 && takeRepeats(config, 1, weight) === 0) {
                            text = repeatPreview(joined ?? '');
                        } else if (Array.isArray(content)) {
                            sharedRuns = content;
                            text = joined ?? '';
                        } else {
                            text = content || '';
                        }
                    } else if (isContent !== undefined) {
                        // Its text and each rich run's, in order (the first alone dropped every later
                        // run), up to its phonetic readings (`rPh`, which follow the runs).
                        const phonetic = nextElement(isContent, 'rPh', 0);
                        const body = phonetic === -1 ? isContent : isContent.slice(0, phonetic);
                        let joined = '';
                        for (const t of xmlElements(body, 't')) joined += t.content;
                        text = decodeXmlEntities(joined.trim());
                    } else if (type === 'b' && vContent !== undefined) {
                        // A boolean is stored as 1 or 0 and shown as TRUE or FALSE.
                        text = vContent.trim() === '1' ? 'TRUE' : vContent.trim() === '0' ? 'FALSE' : vContent.trim();
                    } else if (vContent !== undefined) {
                        text = vContent.trim();
                        // A number in a date or time style is a day count: show the date or time instead
                        const dateFormatKind = type === 'n' && styleIdx !== undefined ? cellDateFormatMap[styleIdx] : undefined;
                        if (dateFormatKind) text = formatDateSerial(text, dateFormatKind, date1904) ?? text;
                        const formatCode = type === 'n' && styleIdx !== undefined ? cellNumberFormatMap[styleIdx] : undefined;
                        if (formatCode !== undefined) {
                            const shown = excelShownText(vContent.trim(), formatCode, date1904);
                            if (shown !== undefined && shown !== text) shownText = shown;
                        }
                    }

                    // Parse cell coordinate
                    const coordMatch = cAttrs.match(/r="([A-Z]+)(\d+)"/);
                    let colIndex: number;
                    let ref: string | undefined;
                    if (coordMatch) {
                        ref = coordMatch[1] + coordMatch[2];
                        colIndex = colToNumber(coordMatch[1]);
                        // If row index is missing in cell coord (unlikely but possible), use rowIndex
                    } else {
                        colIndex = lastColIndex + 1;
                    }
                    lastColIndex = colIndex;

                    if (text || sharedRuns?.length) {
                        const cellFormatting = (styleIdx !== undefined && cellFormatMap[styleIdx]) ? cellFormatMap[styleIdx] : {};

                        if (sharedRuns?.length) {
                            // The cell's style under each run's own: its background and alignment always
                            // apply, its font and size where a run has none. On copies of the runs, which
                            // every cell showing the string shares: merged into them in place, the last
                            // cell's style was every cell's.
                            const { backgroundColor, alignment, font, size } = cellFormatting;
                            cellNodes = backgroundColor || alignment || font || size
                                ? sharedRuns.map(run => ({
                                    ...run,
                                    formatting: {
                                        ...run.formatting,
                                        ...(backgroundColor ? { backgroundColor } : {}),
                                        ...(alignment ? { alignment } : {}),
                                        ...(font && !run.formatting?.font ? { font } : {}),
                                        ...(size && !run.formatting?.size ? { size } : {}),
                                    },
                                }))
                                : sharedRuns.slice();
                        } else {
                            // Simple text node
                            cellNodes.push({
                                type: 'text',
                                text: text,
                                formatting: cellFormatting
                            });
                        }

                        // The cell's hyperlink on its runs, as a document's links are on theirs.
                        const link = links.get(`${rowIndex}:${colIndex}`);
                        if (link) cellNodes = cellNodes.map(run => run.type === 'text' ? { ...run, metadata: { ...(run.metadata as TextMetadata | undefined), ...link } as TextMetadata } : run);

                        const commentsNodeList = (ref && sheetCommentsMap[ref]) ? sheetCommentsMap[ref] : undefined;

                        const span = merged.get(`${rowIndex}:${colIndex}`);
                        const cellNode: OfficeContentNode = {
                            type: 'cell',
                            text: text,
                            children: cellNodes,
                            comments: commentsNodeList,
                            metadata: {
                                row: rowIndex, col: colIndex,
                                ...(span && span.colSpan > 1 ? { colSpan: span.colSpan } : {}),
                                ...(span && span.rowSpan > 1 ? { rowSpan: span.rowSpan } : {}),
                                ...(shownText !== undefined ? { shownText } : {}),
                                ...(rowHidden || columnHidden(colIndex) ? { hidden: true } : {}),
                            }
                        };
                        if (config.includeRawContent) {
                            cellNode.rawContent = chargeRawContent(cXml, config);
                        }
                        cells.push(cellNode);
                    }
                }

                if (cells.length > 0) {
                    const rowNode: OfficeContentNode = {
                        type: 'row',
                        children: cells,
                        metadata: rowHidden ? { hidden: true } : undefined
                    };
                    if (config.includeRawContent) {
                        rowNode.rawContent = chargeRawContent(rowXml, config);
                    }
                    rows.push(rowNode);
                }
            }

            // Handle Drawings in Sheet (images and charts)
            if (config.extractAttachments) {
                // Each <drawing r:id="..."> the sheet holds (scanned as its rows are: a pattern looking
                // for the closing quote from each start read the rest of the line again for each).
                {
                    for (const drawing of xmlElements(sheetXml, 'drawing')) {
                        const rId = /(?:^|\s)r:id="([^"]*)"/.exec(drawing.attrs)?.[1] ?? null;

                        // Each drawing's pictures and charts once, on the first sheet naming it: placed at
                        // every reference, 4.5 KB naming one drawing of 1,000 pictures 1,000 times made a
                        // million picture nodes, and 20 KB ran the process out of memory.
                        if (rId && drawingMap[rId] && !drawingsPlaced.has(drawingMap[rId])) {
                            const drawingPath = drawingMap[rId];
                            drawingsPlaced.add(drawingPath);

                            // Find all images in this drawing
                            const images = drawingImageMap[drawingPath];
                            if (images) {
                                for (const imgId in images) {
                                    const imgInfo = images[imgId];
                                    const attachment = attachmentsByName.get(imgInfo.path.split('/').pop());
                                    if (attachment) {
                                        const imageNode: OfficeContentNode = {
                                            type: 'image',
                                            text: '', // Will be populated by assignAttachmentData
                                            children: [],
                                            metadata: {
                                                attachmentName: attachment.name || 'unknown',
                                                altText: imgInfo.altText || undefined
                                            } as ImageMetadata
                                        };
                                        rows.push(imageNode);
                                    }
                                }
                            }

                            // Find all charts in this drawing
                            const charts = drawingChartMap[drawingPath];
                            if (charts) {
                                for (const chartRId in charts) {
                                    const chartName = charts[chartRId];
                                    const attachment = attachmentsByName.get(chartName);
                                    if (attachment) {
                                        const chartNode: OfficeContentNode = {
                                            type: 'chart',
                                            text: '', // Will be populated by assignAttachmentData
                                            children: [],
                                            metadata: {
                                                attachmentName: chartName
                                            } as ChartMetadata
                                        };
                                        rows.push(chartNode);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Get proper sheet name from workbook.xml mapping, fallback to filename
            const sheetFileName = file.path.split('/').pop() || 'Sheet';
            const sheetName = sheetNameMap[sheetFileName] || sheetFileName;

            const pageSetup = config.ignorePageGeometry ? undefined : sheetPageSetup(sheetXml);
            content.push({
                type: 'sheet',
                children: rows,
                metadata: { sheetName, ...(hiddenSheets.has(sheetFileName) ? { hidden: true } : {}), ...(pageSetup ? { pageSetup } : {}) },
                rawContent: config.includeRawContent ? chargeRawContent(file.content.toString(), config) : undefined
            });
        }
    }

    const corePropsFile = files.find(f => f.path.match(corePropsFileRegex));
    const metadata = corePropsFile ? parseOfficeMetadata(corePropsFile.content.toString(), config) : {};
    const customPropsFile = files.find(f => f.path.match(customPropsFileRegex));
    if (customPropsFile) {
        const customProperties = parseOOXMLCustomProperties(customPropsFile.content.toString(), config);
        if (Object.keys(customProperties).length > 0) metadata.customProperties = customProperties;
    }
    const appPropsFile = files.find(f => f.path.match(appPropsFileRegex));
    if (appPropsFile) {
        const appProperties = parseOOXMLAppProperties(appPropsFile.content.toString(), config);
        if (Object.keys(appProperties).length > 0) metadata.nativeProperties = appProperties;
    }

    // Link OCR text and chart data to content nodes (like PPTX parser)
    const assignAttachmentData = (nodes: OfficeContentNode[]) => {
        for (const node of nodes) {
            if ('attachmentName' in (node.metadata || {})) {
                const meta = node.metadata as ImageMetadata | ChartMetadata;
                const attachment = attachmentsByName.get(meta.attachmentName);
                if (attachment) {
                    if (node.type === 'image') {
                        // Link OCR text to image node
                        if (attachment.ocrText) {
                            node.text = attachment.ocrText;
                        }
                        // Copy altText to attachment
                        if ((meta as ImageMetadata).altText) {
                            attachment.altText = (meta as ImageMetadata).altText;
                        }
                    }
                    if (node.type === 'chart') {
                        // Link chart data text to chart node
                        if (attachment.chartData) {
                            node.text = attachmentsByName.chartText(attachment, config.newlineDelimiter);
                        }
                    }
                }
            }
            if (node.children) {
                assignAttachmentData(node.children);
            }
        }
    };
    assignAttachmentData(content);

    return createAST(
        'xlsx',
        metadata,
        content,
        attachments,
        config,
        undefined,
    );
};

import { zipSync, Zippable } from 'fflate';
import { ConversionResult, GeneratorConfig, OfficeContentNode, OfficeParserAST, OfficeWarningType } from '../types.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { escapeXml, stripInvalidXmlChars } from '../utils/sanitize.js';
import { encUrl, odfUserDefinedXml, resolveZipInstant, toW3CDTF } from '../utils/officeGenUtils.js';
import { BaseGenerator } from './BaseGenerator.js';
import { buildWorkbook, cellTextProcessor, noteLabels, rowAt, sheetIndex, sheetRef, writableLink, CellValue, hiddenColumns, wholeInFull, WorkSheet } from './sheets/workbook.js';

/** Escapes text content for an XML part after stripping XML-illegal characters. */
function xmlText(s: string | undefined): string {
    return escapeXml(stripInvalidXmlChars(s ?? ''));
}

/** The longest sheet name ODS output writes (ODF sets none; Excel's 31 is XLSX's alone). */
const MAX_SHEET_NAME = 255;

const ODF_NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
    + 'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
    + 'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" '
    + 'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:xlink="http://www.w3.org/1999/xlink"';

/**
 * Generates ODS spreadsheets from an AST: a spreadsheet's sheets as they are (values typed and shown in
 * their formats, merged cells, hidden sheets, rows and columns), or a document's tables as sheets and its
 * text on a sheet of its own (see buildWorkbook). Pictures and charts, which a cell cannot hold, are
 * reported.
 */
export class OdsGenerator extends BaseGenerator<'ods'> {
    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'ods'>) {
        super('ods', ast, config);
    }

    /** The automatic styles the content uses: data styles, cell styles, column widths, table styles. */
    private readonly styles = new Map<string, { name: string; xml: string }>();
    private styleNamed(key: string, prefix: string, xml: (name: string) => string): string {
        let style = this.styles.get(key);
        if (!style) {
            const name = `${prefix}${this.styles.size + 1}`;
            style = { name, xml: xml(name) };
            this.styles.set(key, style);
        }
        return style.name;
    }

    /** A number's data style elements (`number:number`), its decimals and grouping. */
    private static numberXml(v: CellValue): string {
        return `<number:number number:decimal-places="${Math.min(v.decimals ?? 0, 15)}" number:min-decimal-places="${Math.min(v.decimals ?? 0, 15)}" number:min-integer-digits="1"${v.grouping ? ' number:grouping="true"' : ''}/>`;
    }

    /** The data style a value is shown in, or none (a plain number or text). */
    private dataStyle(v: CellValue): string | undefined {
        switch (v.kind) {
            case 'number': {
                // (A spreadsheet's own format where ODF can say it: its scientific notation, its fraction, its words
                // around the number, `0.0 "kg"`. Else the format read from what it showed.)
                const own = v.code ? this.ownStyle(v.code) : undefined;
                if (own) return own;
                // (A whole number of twelve to fifteen digits, an id, is shown in full, not in scientific notation.)
                if (!v.grouping && !v.decimals && !wholeInFull(v)) return undefined;
                return this.styleNamed(`n|${v.decimals}|${v.grouping}`, 'N', name => `<number:number-style style:name="${name}">${OdsGenerator.numberXml(v)}</number:number-style>`);
            }
            case 'percent':
                return this.styleNamed(`p|${v.decimals}|${v.grouping}`, 'N', name => `<number:percentage-style style:name="${name}">${OdsGenerator.numberXml(v)}<number:text>%</number:text></number:percentage-style>`);
            case 'currency': {
                // (A spreadsheet's own format where ODF can say it, as a number's: `"€"#,##0.0,,` showing millions.)
                const own = v.code ? this.ownStyle(v.code) : undefined;
                if (own) return own;
                const symbol = `<number:currency-symbol>${xmlText(v.symbol)}</number:currency-symbol>`;
                return this.styleNamed(`c|${v.decimals}|${v.grouping}|${v.symbol}|${v.symbolAfter}`, 'N', name => `<number:currency-style style:name="${name}">`
                    + (v.symbolAfter ? `${OdsGenerator.numberXml(v)}<number:text> </number:text>${symbol}` : `${symbol}${OdsGenerator.numberXml(v)}`) + `</number:currency-style>`);
            }
            case 'date':
                return this.styleNamed('d', 'N', name => `<number:date-style style:name="${name}"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style>`);
            case 'datetime':
                // (One shown as its time of day: a time style on the date it holds.)
                if (v.clock) return this.timeOfDayStyle();
                return this.styleNamed('dt', 'N', name => `<number:date-style style:name="${name}"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/><number:text> </number:text><number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/></number:date-style>`);
            case 'time':
                // (A duration of a day or more in its hours: `number:truncate-on-overflow` off.)
                if ((v.number ?? 0) >= 1) return this.styleNamed('th', 'N', name => `<number:time-style style:name="${name}" number:truncate-on-overflow="false"><number:hours/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/></number:time-style>`);
                return this.timeOfDayStyle();
            case 'boolean':
                return this.styleNamed('b', 'N', name => `<number:boolean-style style:name="${name}"><number:boolean/></number:boolean-style>`);
            default: return undefined;
        }
    }

    /**
     * A spreadsheet's own number format as an ODF number style, where ODF can say it (see odfNumberParts): one
     * section, or more, the second the negative numbers' (`#,##0.0;(#,##0.0)`) and the third zero's, as a style
     * of the negative ones mapping the others to theirs, as LibreOffice writes one; a fourth, text's, is left.
     * Undefined for any other code.
     */
    private ownStyle(code: string): string | undefined {
        // (Its sections, split at semicolons outside quotes and not escaped.)
        const sections: string[] = [];
        let current = '', quoted = false;
        for (let i = 0; i < code.length; i++) {
            const c = code[i];
            if (c === '\\' && !quoted) { current += c + (code[i + 1] ?? ''); i++; continue; }
            if (c === '"') quoted = !quoted;
            if (c === ';' && !quoted) { sections.push(current); current = ''; continue; }
            current += c;
        }
        sections.push(current);
        // (Positive numbers', negative ones', zero's, and text's, which a number never shows: it is left.)
        if (sections.length > 4) return undefined;
        const numeric = sections.slice(0, 3);
        // (A section of nothing shows nothing: `0;-0;;@` hides zero, `;;;` every number.)
        const parts = numeric.map(section => section === '' ? '<number:text></number:text>' : odfNumberParts(section));
        if (parts.some(p => p === undefined)) return undefined;
        const [positive, negative, zero] = parts as string[];
        const styleOf = (inner: string) => this.styleNamed(`nc|${inner}`, 'N', name => `<number:number-style style:name="${name}">${inner}</number:number-style>`);
        if (negative === undefined) return styleOf(positive);
        // The negative numbers' style, mapping the others to theirs, as LibreOffice writes one (the sign is the code's).
        const maps = zero === undefined ? `<style:map style:condition="value()&gt;=0" style:apply-style-name="${styleOf(positive)}"/>`
            : `<style:map style:condition="value()&gt;0" style:apply-style-name="${styleOf(positive)}"/><style:map style:condition="value()=0" style:apply-style-name="${styleOf(zero)}"/>`;
        return this.styleNamed(`nc|${positive}|${negative}|${zero ?? ''}`, 'N', name => `<number:number-style style:name="${name}">${negative}${maps}</number:number-style>`);
    }

    /** A time of day's data style (`14:05:00`). */
    private timeOfDayStyle(): string {
        return this.styleNamed('t', 'N', name => `<number:time-style style:name="${name}"><number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/></number:time-style>`);
    }

    /** A cell's style: its data style, bold or not, wrapped or not; none for a plain cell. */
    private cellStyle(v: CellValue, bold: boolean, wrap: boolean): string | undefined {
        const data = this.dataStyle(v);
        if (!data && !bold && !wrap) return undefined;
        return this.styleNamed(`ce|${data}|${bold}|${wrap}`, 'ce', name => `<style:style style:name="${name}" style:family="table-cell" style:parent-style-name="Default"${data ? ` style:data-style-name="${data}"` : ''}>`
            + (wrap ? '<style:table-cell-properties fo:wrap-option="wrap" style:vertical-align="top"/>' : '')
            + (bold ? '<style:text-properties fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"/>' : '') + `</style:style>`);
    }

    async generate(): Promise<ConversionResult<'ods'>> {
        // onNode is asked once a node: a cell's verdict is read when it is kept and again for its text.
        this.onNodeMemo = new WeakMap();
        try {
            return await this.write();
        } finally {
            this.onNodeMemo = null;
        }
    }

    private async write(): Promise<ConversionResult<'ods'>> {
        const cfg = this.config.odsConfig;
        // Every node the workbook takes is asked for here, so the signal is heard node by node.
        const keep = async (node: OfficeContentNode) => { checkAbortSignal(this.config.abortSignal); const v = await this.handleOnNode(node); return v === false ? false : typeof v === 'string' ? v : true; };
        const labels = await noteLabels(this.ast.content, keep);
        const textProcessor = cellTextProcessor(this.imageMode(), labels);
        const chartData = (node: OfficeContentNode) => this.ast.attachments?.find(a => a?.name === (node.metadata as { attachmentName?: string } | undefined)?.attachmentName)?.chartData;
        const { sheets, unrepresentable } = await buildWorkbook(this.ast.content, {
            chartTable: node => {
                const data = chartData(node);
                return data ? this.chartTable(data, text => ({ type: 'cell', text, children: [{ type: 'text', text }] } as OfficeContentNode), true) : undefined;
            },
            chartTitle: node => chartData(node)?.title,
            textOf: async node => (await this.processNodeRecursive(node, textProcessor)).replace(/\n+$/, ''),
            keep,
        }, { sheetPerTable: cfg.sheetPerTable !== false, includeText: cfg.includeText !== false, includeCharts: this.config.includeCharts !== false, imageMode: this.imageMode(), includeFormatting: this.config.includeFormatting !== false, noteLabels: labels, sheetNameLength: MAX_SHEET_NAME });
        for (const feature of unrepresentable) this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature, format: 'ods' });
        if (!sheets.length) {
            this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: 'a document with no table or text', format: 'ods' });
            sheets.push({ name: 'Sheet1', rows: [], widths: [] });
        }
        if (sheets.every(s => s.hidden)) delete sheets[0].hidden;
        const sheetNames = sheetIndex(sheets);
        const tables = sheets.map(sheet => this.tableXml(sheet, sheetNames)).join('');
        for (const feature of this.linkLosses) this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature, format: 'ods' });
        const { iso, mtime } = resolveZipInstant(this.effectiveMetadata.modified);
        const enc = (s: string) => new TextEncoder().encode(s);
        const m = this.effectiveMetadata;
        const el = (tag: string, val: string | undefined) => val ? `<${tag}>${xmlText(val)}</${tag}>` : '';
        const files: Zippable = {
            // The mimetype first and stored, as ODF readers find it.
            mimetype: [enc('application/vnd.oasis.opendocument.spreadsheet'), { level: 0 }],
            'content.xml': enc(`<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${ODF_NS} office:version="1.2">`
                + `<office:automatic-styles>${[...this.styles.values()].map(s => s.xml).join('')}</office:automatic-styles>`
                + `<office:body><office:spreadsheet>${tables}</office:spreadsheet></office:body></office:document-content>`),
            'styles.xml': enc(`<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ${ODF_NS} office:version="1.2"><office:styles>`
                + `<style:style style:name="Default" style:family="table-cell"/></office:styles></office:document-styles>`),
            'meta.xml': enc(`<?xml version="1.0" encoding="UTF-8"?>\n<office:document-meta ${ODF_NS} office:version="1.2"><office:meta>`
                + el('dc:title', m.title) + el('meta:initial-creator', m.author) + el('dc:creator', m.lastModifiedBy || m.author) + el('dc:subject', m.subject) + el('dc:description', m.description)
                + (m.keywords ? String(m.keywords).split(',').map(k => k.trim()).filter(Boolean).map(k => `<meta:keyword>${xmlText(k)}</meta:keyword>`).join('') : '')
                + odfUserDefinedXml(m.customProperties)
                + `<meta:creation-date>${toW3CDTF(m.created) ?? iso}</meta:creation-date><dc:date>${iso}</dc:date><meta:generator>officeParser</meta:generator></office:meta></office:document-meta>`),
            'META-INF/manifest.xml': enc(`<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">`
                + `<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>`
                + ['content.xml', 'styles.xml', 'meta.xml'].map(p => `<manifest:file-entry manifest:full-path="${p}" manifest:media-type="text/xml"/>`).join('')
                + `</manifest:manifest>`),
        };
        return { value: zipSync(files, { mtime }), messages: this.messages };
    }

    /** Whether a link to a place no written sheet holds was left out (reported once). */
    private readonly linkLosses = new Set<string>();

    /**
     * A sheet as a table: its columns' widths (hidden ones collapsed), its rows and cells, merges covering
     * cells. Written where its cells stand, empty rows and columns between them repeated, not one by one:
     * a sheet of two cells a million rows (or sixteen thousand columns) apart wrote every position.
     */
    private tableXml(sheet: WorkSheet, sheetNames: ReadonlyMap<string, string>): string {
        const hiddenCols = hiddenColumns(sheet);
        const tableStyle = sheet.hidden ? this.styleNamed('ta|hidden', 'ta', name => `<style:style style:name="${name}" style:family="table"><style:table-properties table:display="false"/></style:style>`) : undefined;
        // The columns with widths of their own; those between them, one default column repeated.
        const columnXml: string[] = [];
        let nextColumn = 0;
        sheet.widths.forEach((w, i) => {
            if (i > nextColumn) columnXml.push(`<table:table-column${i - nextColumn > 1 ? ` table:number-columns-repeated="${i - nextColumn}"` : ''}/>`);
            const style = this.styleNamed(`co|${w}`, 'co', name => `<style:style style:name="${name}" style:family="table-column"><style:table-column-properties style:column-width="${(w * 0.19).toFixed(3)}cm"/></style:style>`);
            columnXml.push(`<table:table-column table:style-name="${style}"${hiddenCols.has(i) ? ' table:visibility="collapse"' : ''}/>`);
            nextColumn = i + 1;
        });
        const columns = columnXml.join('');
        // Positions a merged cell covers, (row, col), and each row's covered columns (as many as the
        // workbook's merges cover, see withinRows).
        const covered = new Set<string>();
        const coveredCols = new Map<number, number[]>();
        for (const [i, row] of sheet.rows.entries()) {
            const r = rowAt(row, i);
            for (const c of row.cells) {
                for (let dr = 0; dr < c.rowSpan; dr++) for (let dc = 0; dc < c.colSpan; dc++) {
                    if (!dr && !dc) continue;
                    covered.add(`${r + dr},${c.col + dc}`);
                    let list = coveredCols.get(r + dr);
                    if (!list) coveredCols.set(r + dr, list = []);
                    list.push(c.col + dc);
                }
            }
        }
        for (const list of coveredCols.values()) list.sort((a, b) => a - b);
        // The rows to write, in order: the sheet's, and those only a merge reaches.
        const listed = sheet.rows.map((row, i) => ({ r: rowAt(row, i), row }));
        const listedRows = new Set(listed.map(l => l.r));
        const onlyCovered = [...coveredCols.keys()].filter(r => !listedRows.has(r)).sort((a, b) => a - b);
        const order: { r: number; row: WorkSheet['rows'][number] }[] = [];
        for (let a = 0, b = 0; a < listed.length || b < onlyCovered.length;) {
            if (b >= onlyCovered.length || (a < listed.length && listed[a].r <= onlyCovered[b])) order.push(listed[a++]);
            else order.push({ r: onlyCovered[b++], row: { cells: [] } });
        }
        // Rows in order, a run of empty ones (nothing in them, none covered) one row repeated.
        const rowXml: string[] = [];
        let emptyRows = 0, emptyHidden = false, next = 0;
        const flushRows = () => {
            if (!emptyRows) return;
            rowXml.push(`<table:table-row${emptyRows > 1 ? ` table:number-rows-repeated="${emptyRows}"` : ''}${emptyHidden ? ' table:visibility="collapse"' : ''}><table:table-cell/></table:table-row>`);
            emptyRows = 0;
        };
        const empty = (count: number, hidden: boolean) => {
            if (emptyRows && emptyHidden !== hidden) flushRows();
            emptyHidden = hidden;
            emptyRows += count;
        };
        for (const { r, row } of order) {
            if (r < next) continue;
            if (r > next) empty(r - next, false);
            next = r + 1;
            const cols = coveredCols.get(r) ?? [];
            if (!row.cells.length && !cols.length) { empty(1, !!row.hidden); continue; }
            flushRows();
            rowXml.push(this.rowXml(row, r, cols, covered, sheetNames));
        }
        flushRows();
        const rows = rowXml.join('');
        return `<table:table table:name="${xmlText(sheet.name)}"${tableStyle ? ` table:style-name="${tableStyle}"` : ''}>${columns || '<table:table-column/>'}${rows || '<table:table-row><table:table-cell/></table:table-row>'}</table:table>`;
    }

    /** A row of a sheet's table: its cells, the empty positions between them repeated, and those merges cover. */
    private rowXml(row: WorkSheet['rows'][number], r: number, coveredCols: readonly number[], covered: Set<string>, sheetNames: ReadonlyMap<string, string>): string {
        const out: string[] = [];
        const cells = row.cells;
        let next = 0, ci = 0, ki = 0;
        while (ci < cells.length || ki < coveredCols.length) {
            const cellCol = ci < cells.length ? cells[ci].col : Infinity, coveredCol = ki < coveredCols.length ? coveredCols[ki] : Infinity;
            const at = Math.min(cellCol, coveredCol);
            if (at < next) { if (cellCol === at) ci++; else ki++; continue; }
            if (at > next) out.push(`<table:table-cell${at - next > 1 ? ` table:number-columns-repeated="${at - next}"` : ''}/>`);
            if (cellCol !== at) {
                // A run of covered positions, one covered cell repeated.
                let run = 1;
                ki++;
                while (ki < coveredCols.length && coveredCols[ki] === at + run && (ci >= cells.length || cells[ci].col !== at + run)) { run++; ki++; }
                out.push(`<table:covered-table-cell${run > 1 ? ` table:number-columns-repeated="${run}"` : ''}/>`);
                next = at + run;
                continue;
            }
            const cell = cells[ci++];
            if (coveredCol === at) ki++;
            next = at + 1;
            // A value under another cell's merge (a spreadsheet keeps one there, hidden) is a covered cell
            // holding it, as ODF has it: a cell of its own there overlapped the merge.
            const tag = covered.has(`${r},${at}`) ? 'table:covered-table-cell' : 'table:table-cell';
            const style = this.cellStyle(cell.value, cell.bold, cell.text.includes('\n'));
            const attrs = (style ? ` table:style-name="${style}"` : '')
                + (tag === 'table:table-cell' && cell.colSpan > 1 ? ` table:number-columns-spanned="${cell.colSpan}"` : '') + (tag === 'table:table-cell' && cell.rowSpan > 1 ? ` table:number-rows-spanned="${cell.rowSpan}"` : '')
                + valueAttrs(cell.value);
            // (What the cell shows: a spreadsheet's number in its format, as ODF holds it beside the value.)
            // Its link on its text, line by line.
            // (A place in the workbook as ODF names a cell, `#Sheet.A1`, where the sheet is written; else reported.)
            // (The first of its links a workbook holds, see writableLink; the rest reported.)
            const chosen = writableLink(cell.links, sheetNames);
            for (const why of chosen.lost) this.linkLosses.add(why);
            // (A range as ODF writes one: each end on its sheet.)
            const href = chosen.place ? `#${sheetRef(chosen.place.sheet)}.${chosen.place.cell}${chosen.place.to ? `:${sheetRef(chosen.place.sheet)}.${chosen.place.to}` : ''}` : chosen.link ? encUrl(chosen.link) : undefined;
            const linked = (line: string) => href ? `<text:a xlink:type="simple" xlink:href="${xmlText(href)}">${xmlText(line)}</text:a>` : xmlText(line);
            const paragraphs = cell.shown ? cell.shown.split('\n').map(line => `<text:p>${linked(line)}</text:p>`).join('') : '';
            // Its comments as the cell's annotation (ODF holds one a cell: the first comment's author its own,
            // the others' named in their text).
            const note = cell.comments?.length ? `<office:annotation>${cell.comments[0].author ? `<dc:creator>${xmlText(cell.comments[0].author)}</dc:creator>` : ''}`
                + cell.comments.map((c, i) => ((i && c.author ? `${c.author}: ` : '') + c.text).split('\n').map(line => `<text:p>${xmlText(line)}</text:p>`).join('')).join('') + `</office:annotation>` : '';
            out.push(`<${tag}${attrs}>${note}${paragraphs}</${tag}>`);
        }
        if (!out.length) out.push('<table:table-cell/>');
        return `<table:table-row${row.hidden ? ' table:visibility="collapse"' : ''}>${out.join('')}</table:table-row>`;
    }
}

/**
 * An Excel number format code as an ODF number style's parts, where it is one ODF can say: one section (no
 * colour, condition, percent or text placeholder), words before and after (quoted, escaped, or plain
 * spaces and signs), and a number (`#,##0.00`, scaled by a thousand for each comma ending its integer part),
 * a number in scientific notation (`0.00E+00`) or a fraction (`# ?/?`) between them. Undefined for any other code.
 */
export function odfNumberParts(source: string): string | undefined {
    if (source.length > 255) return undefined;
    // Its colour (`[Red]`), as the text's; its padding the width of a character (`_)`) a space, and its fill (`* `)
    // none: a code of them (Excel's own accounting and comma styles) was refused, and the value shown bare.
    const colour = /^\[(black|blue|cyan|green|magenta|red|white|yellow)\]/i.exec(source);
    let code = '';
    for (let i = colour ? colour[0].length : 0, quoted = false; i < source.length; i++) {
        const c = source[i];
        if (c === '"') quoted = !quoted;
        else if (!quoted && c === '\\') { code += c + (source[i + 1] ?? ''); i++; continue; }
        else if (!quoted && c === '_') { code += ' '; i++; continue; }
        else if (!quoted && c === '*') { i++; continue; }
        code += c;
    }
    if (/[;[%@*_]/.test(code.replace(/"[^"]*"|\\./g, ''))) return undefined;
    const words: string[] = [];
    let rest = code, number = '';
    const literal = (text: string) => text ? `<number:text>${xmlText(text)}</number:text>` : '';
    if (colour) words.push(`<style:text-properties fo:color="${COLOURS.get(colour[1].toLowerCase())}"/>`);
    // The words before the number.
    let before = '';
    for (let m: RegExpExecArray | null; (m = /^(?:"([^"]*)"|\\(.)|([ \-+$€£¥():/]))/.exec(rest));) { before += m[1] ?? m[2] ?? m[3]; rest = rest.slice(m[0].length); }
    // (Words alone, a section showing them for its numbers: `"-"` for zero, `"zero"`.)
    if (!rest) return before ? [...words, literal(before)].join('') : undefined;
    const body = /^([#0?,]*(?:\.[#0?]+,*)?)(?:([Ee][+-]0+)|\s+([#0?]+)\/([#0?]+|\d+))?/.exec(rest);
    if (!body || !/[#0?]/.test(body[1] + (body[3] ?? ''))) return undefined;
    rest = rest.slice(body[0].length);
    let after = '';
    for (let m: RegExpExecArray | null; (m = /^(?:"([^"]*)"|\\(.)|([ \-+$€£¥():/]))/.exec(rest));) { after += m[1] ?? m[2] ?? m[3]; rest = rest.slice(m[0].length); }
    if (rest) return undefined;
    const [written, writtenFraction = ''] = body[1].split('.');
    // (Commas ending the integer part, or the decimals, scale the number by a thousand each, `#,##0,,"M"` and
    // `0.0,"K"` showing 1234567 as 1M and 1234.6K: they were read as grouping, and it showed 1,234,567M.)
    const ending = (text: string) => /,+$/.exec(text)?.[0].length ?? 0;
    const scaling = ending(written) + ending(writtenFraction), whole = written.slice(0, written.length - ending(written));
    const fraction = writtenFraction.slice(0, writtenFraction.length - ending(writtenFraction));
    if (scaling && (body[2] || body[3])) return undefined;
    const decimals = fraction.length, minDecimals = (fraction.match(/0/g) || []).length;
    const minInteger = (whole.match(/0/g) || []).length;
    const grouping = whole.includes(',') ? ' number:grouping="true"' : '';
    const factor = scaling ? ` number:display-factor="${1000 ** Math.min(scaling, 5)}"` : '';
    if (body[2]) number = `<number:scientific-number number:decimal-places="${decimals}" number:min-decimal-places="${minDecimals}" number:min-integer-digits="${Math.max(1, minInteger)}" number:min-exponent-digits="${body[2].length - 2}"${grouping}/>`;
    else if (body[3]) number = `<number:fraction number:min-integer-digits="${minInteger}" number:min-numerator-digits="${body[3].length}" ${/^\d+$/.test(body[4]) ? `number:denominator-value="${body[4]}"` : `number:min-denominator-digits="${body[4].length}"`}${grouping}/>`;
    else number = `<number:number number:decimal-places="${decimals}" number:min-decimal-places="${minDecimals}" number:min-integer-digits="${minInteger}"${grouping}${factor}/>`;
    words.push(literal(before), number, literal(after));
    return words.join('');
}

/** Excel's colours a format names, as ODF writes them. */
const COLOURS = new Map([['black', '#000000'], ['blue', '#0000FF'], ['cyan', '#00FFFF'], ['green', '#00FF00'], ['magenta', '#FF00FF'], ['red', '#FF0000'], ['white', '#FFFFFF'], ['yellow', '#FFFF00']]);

/** A value's ODF attributes: its type and the value held (text cells hold theirs in their paragraphs). */
function valueAttrs(v: CellValue): string {
    switch (v.kind) {
        case 'number': return ` office:value-type="float" office:value="${v.number}"`;
        case 'percent': return ` office:value-type="percentage" office:value="${v.number}"`;
        case 'currency': return ` office:value-type="currency" office:value="${v.number}"`;
        case 'date': case 'datetime': return ` office:value-type="date" office:date-value="${v.iso}"`;
        case 'time': return ` office:value-type="time" office:time-value="${v.iso}"`;
        case 'boolean': return ` office:value-type="boolean" office:boolean-value="${v.number ? 'true' : 'false'}"`;
        default: return ' office:value-type="string"';
    }
}

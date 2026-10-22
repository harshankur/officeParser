import { zipSync, Zippable } from 'fflate';
import { ConversionResult, GeneratorConfig, OfficeContentNode, OfficeParserAST, OfficeWarningType } from '../types.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { escapeXml, stripInvalidXmlChars } from '../utils/sanitize.js';
import { encUrl, ooxmlCustomPropertiesXml, resolveZipInstant, toW3CDTF } from '../utils/officeGenUtils.js';
import { BaseGenerator } from './BaseGenerator.js';
import { buildWorkbook, cellTextProcessor, noteLabels, sheetIndex, sheetRef, writableLink, MAX_LINK_LENGTH, MAX_SHEET_LINKS, CellValue, hiddenColumns, wholeInFull, WorkSheet } from './sheets/workbook.js';

/** Escapes text content for an XML part after stripping XML-illegal characters. */
function xmlText(s: string | undefined): string {
    return escapeXml(stripInvalidXmlChars(s ?? ''));
}

/** A column's letters (0 is A, 26 AA). */
function columnName(index: number): string {
    let name = '';
    for (let k = index + 1; k > 0; k = Math.floor((k - 1) / 26)) name = String.fromCharCode(65 + ((k - 1) % 26)) + name;
    return name;
}

/** Excel's built-in number formats a cell is shown in, by their codes. */
const BUILTIN_FORMATS = new Map<string, number>([['General', 0], ['0', 1], ['0.00', 2], ['#,##0', 3], ['#,##0.00', 4], ['0%', 9], ['0.00%', 10]]);

/** The most characters a cell holds in Excel. */
const MAX_CELL_TEXT = 32_767;

/** `text` cut to at most `max` UTF-16 units, at a whole character. */
function cutTo(text: string, max: number): string {
    if (text.length <= max) return text;
    let end = max;
    if (/[\ud800-\udbff]/.test(text[end - 1])) end--;
    return text.slice(0, end);
}

/** The number format code a value is shown in. */
function formatCode(v: CellValue): string {
    // (The spreadsheet's own, as Excel wrote it.)
    if (v.code) return v.code;
    const digits = (v.grouping ? '#,##0' : '0') + (v.decimals ? '.' + '0'.repeat(Math.min(v.decimals, 15)) : '');
    switch (v.kind) {
        // (A whole number of twelve to fifteen digits, an id, is shown in full, where General shows 1.23457E+11.)
        case 'number': return v.grouping || v.decimals ? digits : wholeInFull(v) ? '0' : 'General';
        case 'percent': return `${digits}%`;
        case 'currency': {
            const symbol = `"${(v.symbol ?? '').replace(/"/g, '')}"`;
            return v.symbolAfter ? `${digits}\\ ${symbol}` : `${symbol}${digits}`;
        }
        case 'date': return 'yyyy\\-mm\\-dd';
        case 'datetime': return 'yyyy\\-mm\\-dd\\ hh:mm:ss';
        // (A duration of a day or more in its hours, as a clock would turn over.)
        case 'time': return (v.number ?? 0) >= 1 ? '[h]:mm:ss' : 'hh:mm:ss';
        // (Excel shows a truth value as one whatever its format; a reader that shows the 1 it holds is told to.)
        case 'boolean': return '"TRUE";"TRUE";"FALSE"';
        default: return 'General';
    }
}

/**
 * Generates XLSX workbooks from an AST: a spreadsheet's sheets as they are (values typed and shown in
 * their formats, merged cells, hidden sheets, rows and columns), or a document's tables as sheets and
 * its text on a sheet of its own (see buildWorkbook). Pictures and charts, which a cell cannot hold, are
 * reported.
 */
export class XlsxGenerator extends BaseGenerator<'xlsx'> {
    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'xlsx'>) {
        super('xlsx', ast, config);
    }

    async generate(): Promise<ConversionResult<'xlsx'>> {
        // onNode is asked once a node: a cell's verdict is read when it is kept and again for its text.
        this.onNodeMemo = new WeakMap();
        try {
            return await this.write();
        } finally {
            this.onNodeMemo = null;
        }
    }

    private async write(): Promise<ConversionResult<'xlsx'>> {
        const cfg = this.config.xlsxConfig;
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
        }, { sheetPerTable: cfg.sheetPerTable !== false, includeText: cfg.includeText !== false, includeCharts: this.config.includeCharts !== false, imageMode: this.imageMode(), includeFormatting: this.config.includeFormatting !== false, noteLabels: labels });
        for (const feature of unrepresentable) this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature, format: 'xlsx' });
        // A workbook has a sheet at least: an empty one where the document held nothing a sheet takes.
        if (!sheets.length) {
            this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: 'a document with no table or text', format: 'xlsx' });
            sheets.push({ name: 'Sheet1', rows: [], widths: [] });
        }
        // At least one sheet shown: a workbook of hidden sheets alone does not open.
        if (sheets.every(s => s.hidden)) delete sheets[0].hidden;

        const styles = new StyleTable();
        const sheetNames = sheetIndex(sheets);
        const sheetParts = sheets.map(sheet => this.sheetXml(sheet, styles, sheetNames));
        for (const feature of this.linkLosses) this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature, format: 'xlsx' });
        const { iso, mtime } = resolveZipInstant(this.effectiveMetadata.modified);
        const enc = (s: string) => new TextEncoder().encode(s);
        const custom = ooxmlCustomPropertiesXml(this.effectiveMetadata.customProperties);
        const files: Zippable = {
            '[Content_Types].xml': enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
                + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`
                + `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`
                + `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`
                + sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')
                + `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`
                + `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`
                + (custom ? `<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>` : '') + `</Types>`),
            '_rels/.rels': enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
                + `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>`
                + `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`
                + `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>`
                + (custom ? `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>` : '') + `</Relationships>`),
            'xl/workbook.xml': enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
                + `<bookViews><workbookView activeTab="${Math.max(0, sheets.findIndex(s => !s.hidden))}"/></bookViews><sheets>`
                + sheets.map((s, i) => `<sheet name="${xmlText(s.name)}" sheetId="${i + 1}"${s.hidden ? ' state="hidden"' : ''} r:id="rId${i + 1}"/>`).join('')
                + `</sheets></workbook>`),
            'xl/_rels/workbook.xml.rels': enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
                + sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')
                + `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`),
            'xl/styles.xml': enc(styles.xml()),
            'docProps/core.xml': enc(this.coreProps(iso)),
            'docProps/app.xml': enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>officeParser</Application></Properties>`),
        };
        if (custom) files['docProps/custom.xml'] = enc(custom);
        sheetParts.forEach((part, i) => {
            files[`xl/worksheets/sheet${i + 1}.xml`] = enc(part.xml);
            if (part.rels) files[`xl/worksheets/_rels/sheet${i + 1}.xml.rels`] = enc(part.rels);
        });
        return { value: zipSync(files, { mtime }), messages: this.messages };
    }

    /** Whether a cell cut to Excel's 32,767 characters was reported (once an output). */
    private cutCellReported = false;

    /** A sheet's part: its columns' widths (hidden ones hidden), its rows and cells, its merged ranges, its cells' links (and their relationships). */
    /** The links left out (each kind reported once). */
    private readonly linkLosses = new Set<string>();

    private sheetXml(sheet: WorkSheet, styles: StyleTable, sheetNames: ReadonlyMap<string, string>): { xml: string; rels?: string } {
        const hiddenCols = hiddenColumns(sheet);
        const cols = sheet.widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"${hiddenCols.has(i) ? ' hidden="1"' : ''}/>`).join('');
        const merges: string[] = [], links: string[] = [], rels: string[] = [];
        const rows = sheet.rows.map((row, r) => {
            if (!row.cells.length && !row.hidden) return '';
            const cells = row.cells.map(cell => {
                const ref = `${columnName(cell.col)}${r + 1}`;
                // The first of its links the workbook holds (a place on a sheet written, or an address Excel
                // opens), the rest reported; at most MAX_SHEET_LINKS a sheet, as Excel opens.
                const chosen = writableLink(cell.links, sheetNames, MAX_LINK_LENGTH);
                for (const why of chosen.lost) this.linkLosses.add(why);
                if (chosen.link && links.length >= MAX_SHEET_LINKS) this.linkLosses.add('links past a sheet\'s 65,530');
                else if (chosen.place) links.push(`<hyperlink ref="${ref}" location="${xmlText(`${sheetRef(chosen.place.sheet)}!${chosen.place.cell}${chosen.place.to ? `:${chosen.place.to}` : ''}`)}"/>`);
                else if (chosen.link) {
                    const id = `rId${rels.length + 1}`;
                    rels.push(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xmlText(encUrl(chosen.link))}" TargetMode="External"/>`);
                    links.push(`<hyperlink ref="${ref}" r:id="${id}"/>`);
                }
                if (cell.colSpan > 1 || cell.rowSpan > 1) merges.push(`${ref}:${columnName(cell.col + cell.colSpan - 1)}${r + cell.rowSpan}`);
                const wrap = cell.text.includes('\n');
                const style = styles.index(formatCode(cell.value), cell.bold, wrap);
                const s = style ? ` s="${style}"` : '';
                // A truth value and an error as a spreadsheet holds them (`t="b"`, `t="e"`).
                if (cell.value.kind === 'boolean') return `<c r="${ref}"${s} t="b"><v>${cell.value.number ? 1 : 0}</v></c>`;
                if (cell.value.kind === 'error') return `<c r="${ref}"${s} t="e"><v>${xmlText(cell.text.trim().toUpperCase())}</v></c>`;
                if (cell.value.kind !== 'text' && typeof cell.value.number === 'number') return `<c r="${ref}"${s}><v>${cell.value.number}</v></c>`;
                if (!cell.text) return style ? `<c r="${ref}"${s}/>` : '';
                // Excel holds 32,767 characters a cell: the rest is reported, not written.
                if (cell.text.length > MAX_CELL_TEXT && !this.cutCellReported && (this.cutCellReported = true)) this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: 'cell text past 32,767 characters', format: 'xlsx' });
                return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${xmlText(cutTo(cell.text, MAX_CELL_TEXT))}</t></is></c>`;
            }).join('');
            return `<row r="${r + 1}"${row.hidden ? ' hidden="1"' : ''}>${cells}</row>`;
        }).join('');
        const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
            + (cols ? `<cols>${cols}</cols>` : '') + `<sheetData>${rows}</sheetData>`
            + (merges.length ? `<mergeCells count="${merges.length}">${merges.map(ref => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>` : '')
            + (links.length ? `<hyperlinks>${links.join('')}</hyperlinks>` : '')
            + `</worksheet>`;
        return { xml, ...(rels.length ? { rels: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>` } : {}) };
    }

    private coreProps(iso: string): string {
        const m = this.effectiveMetadata;
        const created = toW3CDTF(m.created) ?? iso;
        const el = (tag: string, val: string | undefined) => val ? `<${tag}>${xmlText(val)}</${tag}>` : '';
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
            + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" `
            + `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
            + el('dc:title', m.title) + el('dc:creator', m.author) + el('dc:subject', m.subject)
            + el('dc:description', m.description) + el('cp:keywords', m.keywords) + el('cp:lastModifiedBy', m.lastModifiedBy)
            + `<dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created>`
            + `<dcterms:modified xsi:type="dcterms:W3CDTF">${iso}</dcterms:modified></cp:coreProperties>`;
    }
}

/** The cell formats a workbook uses (`cellXfs`): a number format, bold or not, wrapped or not. */
class StyleTable {
    private readonly formats = new Map<string, number>();
    private readonly xfs = new Map<string, number>([['0|0|0', 0]]);
    private readonly xfXml: string[] = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];

    /** The index of a cell format, added where new; 0 is the plain one. */
    index(code: string, bold: boolean, wrap: boolean): number {
        let id = BUILTIN_FORMATS.get(code);
        if (id === undefined) {
            id = this.formats.get(code) ?? 164 + this.formats.size;
            this.formats.set(code, id);
        }
        const key = `${id}|${bold ? 1 : 0}|${wrap ? 1 : 0}`;
        let at = this.xfs.get(key);
        if (at === undefined) {
            at = this.xfXml.length;
            this.xfs.set(key, at);
            this.xfXml.push(`<xf numFmtId="${id}" fontId="${bold ? 1 : 0}" fillId="0" borderId="0" xfId="0"${id ? ' applyNumberFormat="1"' : ''}${bold ? ' applyFont="1"' : ''}${wrap ? ' applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>' : '/>'}`);
        }
        return at;
    }

    xml(): string {
        const numFmts = [...this.formats].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${xmlText(code)}"/>`).join('');
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
            + (numFmts ? `<numFmts count="${this.formats.size}">${numFmts}</numFmts>` : '')
            + `<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>`
            + `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>`
            + `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>`
            + `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>`
            + `<cellXfs count="${this.xfXml.length}">${this.xfXml.join('')}</cellXfs>`
            + `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
    }
}

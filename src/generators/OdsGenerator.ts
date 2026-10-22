import { zipSync, Zippable } from 'fflate';
import { ConversionResult, GeneratorConfig, OfficeContentNode, OfficeParserAST, OfficeWarningType } from '../types.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { escapeXml, stripInvalidXmlChars } from '../utils/sanitize.js';
import { odfUserDefinedXml, resolveZipInstant, toW3CDTF } from '../utils/officeGenUtils.js';
import { BaseGenerator } from './BaseGenerator.js';
import { buildWorkbook, cellTextProcessor, noteLabels, CellValue, hiddenColumns, wholeInFull, WorkSheet } from './sheets/workbook.js';

/** Escapes text content for an XML part after stripping XML-illegal characters. */
function xmlText(s: string | undefined): string {
    return escapeXml(stripInvalidXmlChars(s ?? ''));
}

const ODF_NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
    + 'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
    + 'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" '
    + 'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/"';

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
            case 'number':
                // (A whole number of twelve to fifteen digits, an id, is shown in full, not in scientific notation.)
                if (!v.grouping && !v.decimals && !wholeInFull(v)) return undefined;
                return this.styleNamed(`n|${v.decimals}|${v.grouping}`, 'N', name => `<number:number-style style:name="${name}">${OdsGenerator.numberXml(v)}</number:number-style>`);
            case 'percent':
                return this.styleNamed(`p|${v.decimals}|${v.grouping}`, 'N', name => `<number:percentage-style style:name="${name}">${OdsGenerator.numberXml(v)}<number:text>%</number:text></number:percentage-style>`);
            case 'currency': {
                const symbol = `<number:currency-symbol>${xmlText(v.symbol)}</number:currency-symbol>`;
                return this.styleNamed(`c|${v.decimals}|${v.grouping}|${v.symbol}|${v.symbolAfter}`, 'N', name => `<number:currency-style style:name="${name}">`
                    + (v.symbolAfter ? `${OdsGenerator.numberXml(v)}<number:text> </number:text>${symbol}` : `${symbol}${OdsGenerator.numberXml(v)}`) + `</number:currency-style>`);
            }
            case 'date':
                return this.styleNamed('d', 'N', name => `<number:date-style style:name="${name}"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style>`);
            case 'datetime':
                return this.styleNamed('dt', 'N', name => `<number:date-style style:name="${name}"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/><number:text> </number:text><number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/></number:date-style>`);
            case 'time':
                return this.styleNamed('t', 'N', name => `<number:time-style style:name="${name}"><number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/></number:time-style>`);
            default: return undefined;
        }
    }

    /** A cell's style: its data style, bold or not, wrapped or not; none for a plain cell. */
    private cellStyle(v: CellValue, bold: boolean, wrap: boolean): string | undefined {
        const data = this.dataStyle(v);
        if (!data && !bold && !wrap) return undefined;
        return this.styleNamed(`ce|${data}|${bold}|${wrap}`, 'ce', name => `<style:style style:name="${name}" style:family="table-cell" style:parent-style-name="Default"${data ? ` style:data-style-name="${data}"` : ''}>`
            + (wrap ? '<style:table-cell-properties fo:wrap-option="wrap" style:vertical-align="top"/>' : '')
            + (bold ? '<style:text-properties fo:font-weight="bold"/>' : '') + `</style:style>`);
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
        const labels = noteLabels(this.ast.content);
        const textProcessor = cellTextProcessor(this.imageMode(), labels);
        const { sheets, unrepresentable } = await buildWorkbook(this.ast.content, {
            chartTable: node => {
                const data = this.ast.attachments?.find(a => a?.name === (node.metadata as { attachmentName?: string } | undefined)?.attachmentName)?.chartData;
                return data ? this.chartTable(data, text => ({ type: 'cell', text, children: [{ type: 'text', text }] } as OfficeContentNode)) : undefined;
            },
            textOf: async node => (await this.processNodeRecursive(node, textProcessor)).replace(/\n+$/, ''),
            // Every node the workbook takes is asked for here, so the signal is heard node by node.
            keep: async node => { checkAbortSignal(this.config.abortSignal); const v = await this.handleOnNode(node); return v === false ? false : typeof v === 'string' ? v : true; },
        }, { sheetPerTable: cfg.sheetPerTable !== false, includeText: cfg.includeText !== false, includeCharts: this.config.includeCharts !== false, imageMode: this.imageMode(), includeFormatting: this.config.includeFormatting !== false, noteLabels: labels });
        for (const feature of unrepresentable) this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature, format: 'ods' });
        if (!sheets.length) {
            this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: 'a document with no table or text', format: 'ods' });
            sheets.push({ name: 'Sheet1', rows: [], widths: [] });
        }
        if (sheets.every(s => s.hidden)) delete sheets[0].hidden;
        const tables = sheets.map(sheet => this.tableXml(sheet)).join('');
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

    /** A sheet as a table: its columns' widths (hidden ones collapsed), its rows and cells, merges covering cells. */
    private tableXml(sheet: WorkSheet): string {
        const hiddenCols = hiddenColumns(sheet);
        const tableStyle = sheet.hidden ? this.styleNamed('ta|hidden', 'ta', name => `<style:style style:name="${name}" style:family="table"><style:table-properties table:display="false"/></style:style>`) : undefined;
        const columns = sheet.widths.map((w, i) => {
            const style = this.styleNamed(`co|${w}`, 'co', name => `<style:style style:name="${name}" style:family="table-column"><style:table-column-properties style:column-width="${(w * 0.19).toFixed(3)}cm"/></style:style>`);
            return `<table:table-column table:style-name="${style}"${hiddenCols.has(i) ? ' table:visibility="collapse"' : ''}/>`;
        }).join('');
        // Positions a merged cell covers: (row, col), and each row's last covered column.
        const covered = new Set<string>();
        const lastCovered = new Map<number, number>();
        for (const [r, row] of sheet.rows.entries()) for (const c of row.cells) {
            for (let dr = 0; dr < c.rowSpan; dr++) for (let dc = 0; dc < c.colSpan; dc++) {
                if (!dr && !dc) continue;
                covered.add(`${r + dr},${c.col + dc}`);
                lastCovered.set(r + dr, Math.max(lastCovered.get(r + dr) ?? -1, c.col + dc));
            }
        }
        const rows = sheet.rows.map((row, r) => {
            const out: string[] = [];
            let col = 0;
            const byCol = new Map(row.cells.map(c => [c.col, c]));
            // Empty positions as one repeated cell; covered ones as covered cells.
            let empty = 0;
            const flush = () => { if (empty) { out.push(`<table:table-cell${empty > 1 ? ` table:number-columns-repeated="${empty}"` : ''}/>`); empty = 0; } };
            let last = lastCovered.get(r) ?? -1;
            for (const c of row.cells) last = Math.max(last, c.col + c.colSpan - 1);
            for (; col <= last; col++) {
                const cell = byCol.get(col);
                if (covered.has(`${r},${col}`) && !cell) { flush(); out.push('<table:covered-table-cell/>'); continue; }
                if (!cell) { empty++; continue; }
                flush();
                const style = this.cellStyle(cell.value, cell.bold, cell.text.includes('\n'));
                const attrs = (style ? ` table:style-name="${style}"` : '')
                    + (cell.colSpan > 1 ? ` table:number-columns-spanned="${cell.colSpan}"` : '') + (cell.rowSpan > 1 ? ` table:number-rows-spanned="${cell.rowSpan}"` : '')
                    + valueAttrs(cell.value);
                // (What the cell shows: a spreadsheet's number in its format, as ODF holds it beside the value.)
                const paragraphs = cell.shown ? cell.shown.split('\n').map(line => `<text:p>${xmlText(line)}</text:p>`).join('') : '';
                out.push(`<table:table-cell${attrs}>${paragraphs}</table:table-cell>`);
            }
            if (!out.length && !empty) out.push('<table:table-cell/>');
            flush();
            return `<table:table-row${row.hidden ? ' table:visibility="collapse"' : ''}>${out.join('')}</table:table-row>`;
        }).join('');
        return `<table:table table:name="${xmlText(sheet.name)}"${tableStyle ? ` table:style-name="${tableStyle}"` : ''}>${columns || '<table:table-column/>'}${rows || '<table:table-row><table:table-cell/></table:table-row>'}</table:table>`;
    }
}

/** A value's ODF attributes: its type and the value held (text cells hold theirs in their paragraphs). */
function valueAttrs(v: CellValue): string {
    switch (v.kind) {
        case 'number': return ` office:value-type="float" office:value="${v.number}"`;
        case 'percent': return ` office:value-type="percentage" office:value="${v.number}"`;
        case 'currency': return ` office:value-type="currency" office:value="${v.number}"`;
        case 'date': case 'datetime': return ` office:value-type="date" office:date-value="${v.iso}"`;
        case 'time': return ` office:value-type="time" office:time-value="${v.iso}"`;
        default: return ' office:value-type="string"';
    }
}

import { ChartData, OfficeAttachment, OfficeIssue, ConversionResult, FullGeneratorConfig, GeneratorConfig, ImageMode, OfficeContentNode, OfficeMetadata, OfficeParserAST, OfficeWarningType, StructuredStyleMapping, UniversalGeneratorFormat } from '../types.js';
import { isFullGeneratorConfig, resolveGeneratorConfig } from '../utils/configUtils.js';
import { checkAbortSignal, getWarningMessage } from '../utils/errorUtils.js';
import { forgetFurniture, isPlaceholderNode, mathPictureOf, resolveImageMode } from '../utils/officeGenUtils.js';
import { StyleMapper } from '../utils/styleMapper.js';
import { isSourceComment } from '../utils/commentUtils.js';
import { gridPositionsFor } from '../utils/sheetGridUtils.js';
import { MAX_LAYOUT_COLUMNS } from '../utils/tableLayout.js';
import { atPosition, chartPositions, chartReferenceText, isXyChart, outerLevelRows, shownOr, xyPositions } from '../utils/chartUtils.js';
import { NoteMarks } from '../utils/noteMarks.js';
import { REPEAT_ALLOWANCE, repeatedContentOfAst, repeatPreview, utf8Length } from '../utils/repeatUtils.js';

/** The most cells one chart's data table holds in DOCX, ODT and LaTeX output (see BaseGenerator.chartTable). */
const MAX_CHART_TABLE_CELLS = 100_000;

/**
 * A chart's table as an output wrote it first, the cells it holds and what one copy of it weighs (see
 * BaseGenerator.chartTable, BaseGenerator.chartTableWritten).
 */
type ChartTableRecord = { table: OfficeContentNode; cells: number; weight: number };

/**
 * What a writer wrote for a chart (see BaseGenerator.chartTableWritten): `text` where the chart stands; `weight`, UTF-8
 * bytes of all it wrote (a table it hoisted elsewhere too); `undo`, taking back what it wrote elsewhere, called when
 * the copy is not kept.
 */
export type ChartWritten = { text: string; weight: number; undo?: () => void };

/** The most picture bytes one document writes inline in all (see BaseGenerator.inlineWithinBudget). */
export const MAX_INLINED_IMAGE_BYTES = 128 * 1024 * 1024;

/**
 * Base class for all document generators.
 * Provides common traversal logic and configuration handling.
 */
export abstract class BaseGenerator<D extends UniversalGeneratorFormat = UniversalGeneratorFormat> {
    protected config: FullGeneratorConfig;
    protected ast: OfficeParserAST;
    protected messages: OfficeIssue[] = [];
    protected styleMapper: StyleMapper;
    protected collectedNotes: OfficeContentNode[] = [];
    private readonly collectedNoteSet = new Set<OfficeContentNode>();
    private readonly writtenComments = new Set<OfficeContentNode>();
    private inlinedImageBytes = 0;
    private paddingCellsLeft: number;
    private paddingLimit: number;
    private gridLimitWarned = false;
    /** Point indexes of each chart's data, found once however many chart nodes show it. */
    private readonly chartPositionsOf = new Map<ChartData, number[] | undefined>();
    /** What each chart's title weighs as this writer writes it, recorded where it is first written (see chartTitle); its tables: see chartTables. */
    private readonly chartTitleWeights = new Map<ChartData, number>();
    /**
     * What chart text this output may still write again (see chartCopy): the parse's repeat budget,
     * `decompressionLimits.maxRepeatedContent` plus 16 characters per byte of the document (repeatedContentOfAst).
     */
    private chartRepeatsLeft: number;
    /** The chart-copy budget this output started with, for the warning's message. */
    private readonly chartRepeatLimit: number;
    private chartRepeatsWarned = false;

    /**
     * `text` as a copy writes it, of chart text this output wrote before (a title at another chart node showing the
     * same data, an axis title over another series' columns): whole when it weighs REPEAT_ALLOWANCE or less, or while
     * the budget lasts, each copy costing its `weight` past REPEAT_ALLOWANCE (what the writer writes of it, in UTF-8
     * bytes: see chartTextWeight); past it, its start (repeatPreview), reported once (REPEATED_CONTENT_LIMIT_EXCEEDED).
     * The original is whole where it was first written. A small chart part whose axis title stood over each of 499
     * series' columns made 400 MB of output.
     */
    private chartCopy(text: string, weight = this.chartTextWeight(text)): string {
        const cost = weight - REPEAT_ALLOWANCE;
        if (cost <= 0) return text;
        if (cost <= this.chartRepeatsLeft) { this.chartRepeatsLeft -= cost; return text; }
        this.warnChartRepeats();
        return repeatPreview(text);
    }

    /** Reports, once per output, that chart text written again passed the chart-copy budget (see chartCopy). */
    private warnChartRepeats(): void {
        if (this.chartRepeatsWarned) return;
        this.chartRepeatsWarned = true;
        this.warn(OfficeWarningType.REPEATED_CONTENT_LIMIT_EXCEEDED, this.chartRepeatLimit);
    }

    /** Each chart's table as this output wrote it first (per ChartData and caption), and what one copy of it weighs. */
    private readonly chartTables = new Map<ChartData, { caption?: ChartTableRecord; plain?: ChartTableRecord }>();

    /**
     * What one cell of a chart's table weighs beyond its text: a cell's markup for each column it spans, as a deck or
     * workbook writer writes it (a PPTX `a:tc` with its text body, an XLSX `c` with its inline string; see chartTable).
     * Measured per writer (testChartTableCopies checks it holds).
     */
    protected chartCellWeight = 16;

    /**
     * What `text` weighs as this writer writes it in a chart's caption or table cell, in UTF-8 bytes: its escaped form,
     * the heaviest the writer has (a line break, a tab and an escaped character as the writer writes each). Charged for
     * a title or an axis title written again (see chartTitle, chartCopy) and, by deck and workbook writers, for each
     * cell of a table copy (see chartTable). The default: the text as it is.
     */
    protected chartTextWeight(text: string): number {
        return utf8Length(text);
    }

    /**
     * What a writer writes for a copy of a chart's table past the budgets (see chartTable): `[Chart: <title>]`, the
     * title cut to its start (repeatPreview), else `[Chart: <attachment name>]`, the form a chart with no data takes
     * (see chartReferenceText).
     */
    protected chartReference(data: ChartData, name: string | undefined): string {
        return chartReferenceText(data, name);
    }

    /**
     * A chart's table (see chartTable) for a writer that places a table where a chart stands (a deck's slide, a
     * workbook's sheet): past the budgets, a table of one cell holding the chart's reference (see chartReference).
     */
    protected chartTableOrReference(data: ChartData, name: string | undefined, cellOf: (text: string) => OfficeContentNode, caption = false): OfficeContentNode {
        return this.chartTable(data, cellOf, caption)
            ?? { type: 'table', children: [{ type: 'row', children: [cellOf(this.chartReference(data, name))] }] } as OfficeContentNode;
    }

    /**
     * What a chart's title weighs as this writer writes it beside or in its table (see chartTitle), in UTF-8 bytes: by
     * default as a cell's text (chartTextWeight); a writer whose caption writes it otherwise says so.
     */
    protected chartTitleWeight(title: string): number {
        return this.chartTextWeight(title);
    }

    /**
     * A chart's title as the writer writes it beside or in its table: '' for none, whole at the first chart node showing
     * the data, a copy (see chartCopy) at each later one, charged what the first weighed (measured once a chart).
     */
    protected chartTitle(data: ChartData): string {
        if (typeof data.title !== 'string' || !data.title) return '';
        const weight = this.chartTitleWeights.get(data);
        if (weight === undefined) { this.chartTitleWeights.set(data, this.chartTitleWeight(data.title)); return data.title; }
        return this.chartCopy(data.title, weight);
    }

    /** What each chart's text weighs as written in a line (see chartLineWritten), the most a copy of it wrote; present once the first is written. */
    private readonly chartLineWeights = new Map<ChartData | string, number>();

    /**
     * A chart's text as written in place (a chart standing in a line in DOCX, ODT and LaTeX; a chart without its data in every
     * flow writer): `write` of its node's text (its data's lines), else its reference (see chartReferenceText) for a chart
     * of no text where charts are written, at the first such chart node of its chart (its data, else its attachment's
     * name); at a later one the same, while
     * the chart-copy budget (see chartCopy) holds what it writes (UTF-8 bytes, decided on the most a copy has weighed,
     * charged what this one wrote); past it `write` of the chart's reference (see chartReference), reported once
     * (REPEATED_CONTENT_LIMIT_EXCEEDED). A 64 KB DOCX of 1,000 paragraphs each holding a chart of 20,000 values made
     * 358 MB of DOCX XML: the parse charged the text's characters, the writer wrote 30 bytes a line.
     */
    protected chartLineWritten(node: OfficeContentNode, write: (text: string) => string): string {
        const name = (node.metadata as { attachmentName?: string } | undefined)?.attachmentName;
        const data = this.getAttachment(name)?.chartData;
        const own = this.getNodeText(node);
        // A chart of no text (a DOCX or PPTX chart parsed without extractAttachments): its reference, as it is written past
        // the budget. In a line it wrote nothing. (Not where charts are left out: includeCharts false keeps what it wrote.)
        const text = (own.trim() || this.config.includeCharts === false) ? own : chartReferenceText(data, name);
        const key = data ?? name;
        if (!key) return write(text);
        const recorded = this.chartLineWeights.get(key);
        if (recorded === undefined) {
            const first = write(text);
            this.chartLineWeights.set(key, utf8Length(first));
            return first;
        }
        if (recorded <= this.chartRepeatsLeft) {
            const copy = write(text);
            const weight = utf8Length(copy);
            if (weight > recorded) this.chartLineWeights.set(key, weight);
            if (weight <= this.chartRepeatsLeft) {
                this.chartRepeatsLeft -= weight;
                return copy;
            }
        }
        this.warnChartRepeats();
        return write(chartReferenceText(data, name));
    }

    /** `text` as written in place, weighed (see ChartWritten). */
    protected chartWritten(text: string): ChartWritten {
        return { text, weight: utf8Length(text) };
    }

    /**
     * A chart's table as a flow writer writes it (DOCX, ODT, LaTeX, HTML, Markdown, RTF), through `write`: at the first chart
     * node showing `data`, built (as chartTable builds one, without a title row) and written whole, and what it wrote
     * recorded; at a later one, the same table written again while the chart-copy budget (see chartCopy) and the grid
     * budget hold what the copy weighs: decided before writing on the most a copy of it has weighed, then charged what this
     * copy wrote (UTF-8 bytes, markup and escapes included). A copy that does not fit is taken back (`undo`), and the
     * recorded weight raised to it, so it is written at most once per chart. Past either budget: undefined, reported once
     * (REPEATED_CONTENT_LIMIT_EXCEEDED), and the writer writes the chart's reference (see chartReference). An 89 KB PPTX
     * whose chart's labels were `^` wrote 18 times its charge in LaTeX and ran a 2 GB heap out.
     */
    protected async chartTableWritten(data: ChartData, cellOf: (text: string) => OfficeContentNode, write: (table: OfficeContentNode) => Promise<ChartWritten>): Promise<string | undefined> {
        let records = this.chartTables.get(data);
        const recorded = records?.plain;
        if (!recorded) {
            let cells = 0;
            const counted = (text: string) => { cells++; return cellOf(text); };
            const table = isXyChart(data) ? this.chartXyTable(data, counted, false) : this.chartGridTable(data, counted, false);
            const first = await write(table);
            if (!records) this.chartTables.set(data, records = {});
            records.plain = { table, cells, weight: first.weight };
            return first.text;
        }
        if (recorded.cells > this.paddingCellsLeft || recorded.weight > this.chartRepeatsLeft) {
            this.warnChartRepeats();
            return undefined;
        }
        const copy = await write(recorded.table);
        recorded.weight = Math.max(recorded.weight, copy.weight);
        if (copy.weight > this.chartRepeatsLeft) {
            copy.undo?.();
            this.warnChartRepeats();
            return undefined;
        }
        this.paddingCellsLeft -= recorded.cells;
        this.chartRepeatsLeft -= copy.weight;
        return copy.text;
    }

    /**
     * A chart's data as DOCX, ODT and LaTeX write it: a header of series names, then a row per label, or a
     * row per point index when the chart's lists give them (see chartPositions). Each label and value is
     * written as the chart shows it (`shownLabels`, `shownValues`) where the chart keeps that text. A chart
     * of several category levels (`labelLevels`) has a column for each outer level, outermost first, before
     * its labels: an outer label in the row of the point it starts at, blank below until the next.
     * The table is labels x series, which a 9.6 KB PPTX made nine million cells (the process ran out of
     * memory), and it is written at every chart showing the data; so it is built within the grid budget
     * (see padWithinBudget) and at most MAX_CHART_TABLE_CELLS cells: at most MAX_LAYOUT_COLUMNS - 1 series,
     * and the rows those hold. What is left out is reported (CONTENT_NOT_REPRESENTABLE). A scatter or bubble
     * chart's table is its series side by side instead (see chartXyTable).
     *
     * A chart node after the first showing the same data (and `caption`) gets the first table's node again, its copy
     * charged to the chart-copy budget (see chartCopy) and its cells to the grid budget: 1,000 frames of one 20,000-value
     * chart in 65 KB of DOCX made 348 MB of RTF. A copy weighs an estimate from the writer's escape and markup, measured
     * per writer: chartCellWeight for each column a cell spans and chartTextWeight of its text (a deck or workbook
     * writer writes the table long after this; a flow writer measures what it writes, see chartTableWritten). Past
     * either budget it is undefined, reported once (REPEATED_CONTENT_LIMIT_EXCEEDED), and the writer writes the chart's
     * reference (see chartReference); its data is whole where it was first written.
     */
    protected chartTable(data: ChartData, cellOf: (text: string) => OfficeContentNode, caption = false): OfficeContentNode | undefined {
        const key = caption ? 'caption' : 'plain';
        let records = this.chartTables.get(data);
        const recorded = records?.[key];
        if (recorded) {
            if (recorded.cells <= this.paddingCellsLeft && recorded.weight <= this.chartRepeatsLeft) {
                this.paddingCellsLeft -= recorded.cells;
                this.chartRepeatsLeft -= recorded.weight;
                return recorded.table;
            }
            this.warnChartRepeats();
            return undefined;
        }
        const made: OfficeContentNode[] = [];
        let textWeight = 0;
        const counted = (text: string) => { const cell = cellOf(text); made.push(cell); textWeight += this.chartTextWeight(text); return cell; };
        const table = isXyChart(data) ? this.chartXyTable(data, counted, caption) : this.chartGridTable(data, counted, caption);
        // A cell spanning columns (the title's row, an XY series' name) is written as a cell a column (PPTX hMerge, XLSX merges).
        let positions = 0;
        for (const cell of made) positions += Math.max(1, Math.floor(Number((cell.metadata as { colSpan?: unknown } | undefined)?.colSpan)) || 1);
        if (!records) this.chartTables.set(data, records = {});
        records[key] = { table, cells: made.length, weight: positions * this.chartCellWeight + textWeight };
        return table;
    }

    /** A chart's data of labels and series as a table, the first time it is written (see chartTable). */
    private chartGridTable(data: ChartData, cellOf: (text: string) => OfficeContentNode, caption: boolean): OfficeContentNode {
        const dataSets = Array.isArray(data.dataSets) ? data.dataSets : [];
        const labels = Array.isArray(data.labels) ? data.labels : [];
        const levelColumns = Array.isArray(data.labelLevels) && data.labelLevels.length > 1 ? data.labelLevels.length - 1 : 0;
        const series = dataSets.slice(0, MAX_LAYOUT_COLUMNS - 1 - levelColumns);
        const width = levelColumns + series.length + 1;
        if (!this.chartPositionsOf.has(data)) this.chartPositionsOf.set(data, chartPositions(data));
        const positions = this.chartPositionsOf.get(data);
        const rowCount = positions ? positions.length : labels.length;
        const shownLabels = shownOr(data.shownLabels, labels);
        const labelAt = (i: number) => positions ? atPosition(shownLabels, data.labelIndexes, positions[i]) : shownLabels[i];
        const valueAt = (d: ChartData['dataSets'][number] | undefined, i: number) => {
            const values = shownOr(d?.shownValues, d?.values);
            return positions ? atPosition(values, d?.valueIndexes, positions[i]) : values[i];
        };
        const allowed = Math.min(rowCount + 1, Math.floor(Math.min(this.paddingCellsLeft, MAX_CHART_TABLE_CELLS) / width));
        this.paddingCellsLeft -= allowed * width;
        // (An outer label at a point no row has, or past the rows, has no row: it is reported, as the table limit's
        // report where the limit cut the table, else as a category label with no point.)
        const outer = levelColumns ? outerLevelRows(data, Math.max(0, allowed - 1), positions) : { levels: [], whole: true };
        if (series.length < dataSets.length || allowed < rowCount + 1) {
            this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: 'chart data past the table limit', format: this.destination });
        } else if (!outer.whole) {
            this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: 'chart category label with no point', format: this.destination });
        }
        const rows: OfficeContentNode[] = [];
        // Its title, where asked (a writer that writes no title of its own beside the table): a row across it.
        const title = caption ? this.chartTitle(data).trim() : '';
        if (title) {
            const cell = cellOf(title);
            cell.metadata = { ...(cell.metadata ?? {}), colSpan: width, style: 'header' } as OfficeContentNode['metadata'];
            rows.push({ type: 'row', children: [cell] } as OfficeContentNode);
        }
        if (allowed > 0) rows.push({ type: 'row', children: [...outer.levels.map(() => cellOf('')), cellOf(''), ...series.map(d => cellOf(d?.name || ''))] } as OfficeContentNode);
        for (let i = 0; i < allowed - 1; i++) {
            rows.push({ type: 'row', children: [...outer.levels.map(level => cellOf(level[i])), cellOf(String(labelAt(i) ?? '')), ...series.map(d => cellOf(String(valueAt(d, i) ?? '')))] } as OfficeContentNode);
        }
        return { type: 'table', children: rows } as OfficeContentNode;
    }

    /** Point indexes of each scatter or bubble chart's data, found once however many chart nodes show it. */
    private readonly xyPositionsOf = new Map<ChartData, number[] | undefined>();

    /**
     * A scatter or bubble chart's data as a table (its series hold X values of their own, `xValues`): no label
     * column, but for each series its X column, its value column and, for a bubble series, its size column; a
     * first header row of series names across their columns and a second of what each column holds (the
     * series' axis titles where the chart has them, else X and Y, and Size); a row per point index (see
     * xyPositions), else per place. Values as the chart shows them (`shownValues`). Within the same budgets as
     * chartTable: the series that fit MAX_LAYOUT_COLUMNS columns, and what is left out reported.
     */
    private chartXyTable(data: ChartData, cellOf: (text: string) => OfficeContentNode, caption: boolean): OfficeContentNode {
        const dataSets = Array.isArray(data.dataSets) ? data.dataSets : [];
        const axisById = new Map((Array.isArray(data.axes) ? data.axes : []).map(axis => [axis?.id, axis] as const));
        const titleOf = (set: ChartData['dataSets'][number], which: 0 | 1, chartTitle: string | undefined, fallback: string) =>
            (Array.isArray(set?.axisIds) ? axisById.get(set.axisIds[which])?.title : undefined) || chartTitle || fallback;
        type Column = { list: unknown[]; indexes: unknown; heading: string };
        const groups: { name: string; columns: Column[] }[] = [];
        let width = 0;
        for (const set of dataSets) {
            const columns: Column[] = [
                { list: Array.isArray(set?.xValues) ? set.xValues : [], indexes: set?.xValueIndexes, heading: titleOf(set, 0, data.xAxisTitle, 'X') },
                { list: shownOr(set?.shownValues, set?.values), indexes: set?.valueIndexes, heading: titleOf(set, 1, data.yAxisTitle, 'Y') },
                ...(Array.isArray(set?.bubbleSizes) && set.bubbleSizes.length ? [{ list: set.bubbleSizes, indexes: set.bubbleSizeIndexes, heading: 'Size' }] : []),
            ];
            if (width + columns.length > MAX_LAYOUT_COLUMNS) break;
            width += columns.length;
            groups.push({ name: set?.name || '', columns });
        }
        if (!this.xyPositionsOf.has(data)) this.xyPositionsOf.set(data, xyPositions(data));
        const positions = this.xyPositionsOf.get(data);
        const rowCount = positions ? positions.length : groups.reduce((most, group) => group.columns.reduce((m, column) => Math.max(m, column.list.length), most), 0);
        const allowed = width ? Math.min(rowCount + 2, Math.floor(Math.min(this.paddingCellsLeft, MAX_CHART_TABLE_CELLS) / width)) : 0;
        this.paddingCellsLeft -= allowed * width;
        if (groups.length < dataSets.length || allowed < rowCount + 2) {
            this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: 'chart data past the table limit', format: this.destination });
        }
        const rows: OfficeContentNode[] = [];
        const spanning = (text: string, colSpan: number, style?: string) => {
            const cell = cellOf(text);
            if (colSpan > 1 || style) cell.metadata = { ...(cell.metadata ?? {}), ...(colSpan > 1 && { colSpan }), ...(style && { style }) } as OfficeContentNode['metadata'];
            return cell;
        };
        const title = caption && width ? this.chartTitle(data).trim() : '';
        if (title) rows.push({ type: 'row', children: [spanning(title, width, 'header')] } as OfficeContentNode);
        if (allowed > 0) rows.push({ type: 'row', children: groups.map(group => spanning(group.name, group.columns.length)) } as OfficeContentNode);
        // Each column's heading is the title of the axis it is drawn on: the series of a group share their axes, so a
        // title stands over two columns a series. The first of each is whole, the others copies (see chartCopy), each
        // charged what the first weighed (measured once a heading; at most MAX_LAYOUT_COLUMNS of them).
        const headings = new Map<string, number>();
        const headingOf = (heading: string) => {
            const weight = headings.get(heading);
            if (weight !== undefined) return this.chartCopy(heading, weight);
            headings.set(heading, this.chartTextWeight(heading));
            return heading;
        };
        if (allowed > 1) rows.push({ type: 'row', children: groups.flatMap(group => group.columns.map(column => cellOf(headingOf(column.heading)))) } as OfficeContentNode);
        for (let i = 0; i < allowed - 2; i++) {
            rows.push({
                type: 'row',
                children: groups.flatMap(group => group.columns.map(column => cellOf(String((positions ? atPosition(column.list, column.indexes, positions[i]) : column.list[i]) ?? '')))),
            } as OfficeContentNode);
        }
        return { type: 'table', children: rows } as OfficeContentNode;
    }

    /** Whether `positions` more grid positions (rows a sparse sheet fills back in) fit the same budget, all or none. */
    protected takeGridPositions(positions: number): boolean {
        if (!(positions >= 0)) return false;
        if (positions > this.paddingCellsLeft) { this.warnGridLimit({ rowsNotFilled: true }); return false; }
        this.paddingCellsLeft -= positions;
        return true;
    }

    /**
     * How many of `missing` empty cells a row may be padded with to reach its table's width, taken from
     * the document's budget (the grid budget's size, see gridPositionsFor). Rows are padded to the widest:
     * one row of 10,000 cells over 10,000 rows of one made 100 million positions from 4 KB, in a table
     * the grid budget had already laid out as tightly as it could. Past the budget a row keeps the cells
     * it has, which is reported once (TABLE_GRID_LIMIT_EXCEEDED); Markdown, CSV and LaTeX read a short
     * row as ending in empty cells.
     */
    protected padWithinBudget(missing: number): number {
        if (!(missing > 0)) return 0;
        const allowed = Math.min(missing, this.paddingCellsLeft);
        this.paddingCellsLeft -= allowed;
        if (allowed < missing) this.warnGridLimit({ unpadded: true });
        return allowed;
    }

    /**
     * Whether a picture of `bytes` may still be written inline (a `data:` URI, RTF picture data), taking
     * it from the document's budget of MAX_INLINED_IMAGE_BYTES; when not, an IMAGE_NOT_INLINED warning.
     * A format that cannot point one picture at another's data writes it at every place showing it, so
     * a small document showing one large picture many times made output past what a string can hold.
     */
    protected inlineWithinBudget(bytes: number, name: string | undefined): boolean {
        if (this.inlinedImageBytes + bytes > MAX_INLINED_IMAGE_BYTES) {
            this.warn(OfficeWarningType.IMAGE_NOT_INLINED, { name, bytes, limit: MAX_INLINED_IMAGE_BYTES, inDocument: true });
            return false;
        }
        this.inlinedImageBytes += bytes;
        return true;
    }

    /**
     * Whether `comment` is being written for the first time (and marks it written). A comment the AST
     * shares among several references is written once, at the first: written at each, one comment a
     * small document referred to thousands of times (or comments referring to each other twice each)
     * made output of gigabytes. Marked before its body is written, so a comment inside itself cannot
     * recurse.
     */
    protected firstWriteOfComment(comment: OfficeContentNode): boolean {
        if (this.writtenComments.has(comment)) return false;
        this.writtenComments.add(comment);
        return true;
    }
    private noteReferenceCounts: Map<OfficeContentNode, number> | undefined;

    /**
     * How many references the AST makes to `note` (a parser shares one note node among all the
     * references to it). A writer that writes a note at its reference writes it in full once and
     * refers back to it from the others, so a note referred to thousands of times is not written
     * thousands of times.
     */
    protected noteReferences(note: OfficeContentNode): number {
        if (!this.noteReferenceCounts) {
            const counts = new Map<OfficeContentNode, number>();
            const seen = new Set<OfficeContentNode>();
            const stack: OfficeContentNode[] = [...(this.ast?.content ?? [])];
            while (stack.length) {
                const node = stack.pop()!;
                if (!node || typeof node !== 'object') continue;
                for (const n of node.notes ?? []) counts.set(n, (counts.get(n) ?? 0) + 1);
                if (seen.has(node)) continue;
                seen.add(node);
                for (const list of [node.children, node.notes, node.comments]) if (list) for (const child of list) stack.push(child);
            }
            this.noteReferenceCounts = counts;
        }
        return this.noteReferenceCounts.get(note) ?? 0;
    }
    /** Lazily-built `name -> attachment` index for {@link getAttachment}. */
    private attachmentIndex?: Map<string, OfficeAttachment>;

    /**
     * @param destination - The format this writer writes.
     * @param ast - The document written.
     * @param config - The generator configuration, resolved here with the AST's (`ast.config`).
     * @param configResolved - `config` is one a writer already resolved (EPUB and the browser-printed PDF write their HTML
     *   with theirs): it holds the AST's config, and its problems were reported. Taken as it is, it is not checked again.
     */
    constructor(protected destination: D, ast: OfficeParserAST, config?: GeneratorConfig<D> | FullGeneratorConfig, configResolved = false) {
        // Problems with the configuration itself are among the result's messages too. A configuration a writer already
        // resolved holds the AST's config and was checked: resolved again with it, an option it did not recognize was
        // reported a second time.
        this.config = resolveGeneratorConfig(destination, configResolved && isFullGeneratorConfig(config) ? undefined : ast.config, config, issue => this.messages.push(issue));
        this.ast = ast;
        // (A header list edited in place since the last generation is read afresh.)
        forgetFurniture(ast.auxiliary?.headers, ast.auxiliary?.footers);
        this.styleMapper = new StyleMapper(this.config.styleMap, this.config.ignoreDefaultStyleMap, this.config);
        this.paddingCellsLeft = this.paddingLimit = gridPositionsFor(ast);
        this.chartRepeatsLeft = this.chartRepeatLimit = repeatedContentOfAst(ast);
    }

    /**
     * Reports that `tables` of the document's tables were laid out closer before writing, to fit the grid
     * budget (see withBoundedSheetGrids): their cells no longer stand where the document put them.
     */
    public reportTablesLaidOut(tables: number): void {
        this.warn(OfficeWarningType.TABLE_GRID_LIMIT_EXCEEDED, { laidOut: tables, limit: this.paddingLimit });
    }

    /** Reports, once per output, that the grid budget left rows unpadded or a sparse sheet's empty rows unwritten. */
    private warnGridLimit(info: { unpadded?: boolean; rowsNotFilled?: boolean }): void {
        if (this.gridLimitWarned) return;
        this.gridLimitWarned = true;
        this.warn(OfficeWarningType.TABLE_GRID_LIMIT_EXCEEDED, { ...info, limit: this.paddingLimit });
    }

    /**
     * Resolves an attachment by name via a lazily-built index, replacing a linear
     * `ast.attachments.find(...)` per image/chart node. First-match-wins keeps the exact semantics of
     * `find`. Every generator resolves image/chart bytes this way, so a media-heavy document (for
     * example a fully inlined self-contained export) stays O(nodes + attachments) rather than
     * O(nodes x attachments).
     */
    protected getAttachment(name: string | undefined): OfficeAttachment | undefined {
        if (!name) return undefined;
        if (!this.attachmentIndex) {
            this.attachmentIndex = new Map();
            for (const a of this.ast.attachments || []) {
                if (a.name && !this.attachmentIndex.has(a.name)) this.attachmentIndex.set(a.name, a);
            }
        }
        return this.attachmentIndex.get(name);
    }

    /** The picture `node`, an equation, is written as by a writer that cannot typeset math (see mathPictureOf), or undefined. */
    protected mathPicture(node: OfficeContentNode): OfficeContentNode | undefined {
        const name = (node.metadata as { fallbackAttachmentName?: unknown } | undefined)?.fallbackAttachmentName;
        return mathPictureOf(node, typeof name === 'string' ? this.getAttachment(name) : undefined, this.imageMode());
    }

    /**
     * Resolves `config.includeImages` (a boolean, a CLI-provided `'true'`/`'false'` string, or an
     * {@link ImageMode}) to a single mode. Every generator's image handling must route through this
     * rather than a truthy check, since a mode string like `'none'` is truthy.
     */
    protected imageMode(): ImageMode {
        return resolveImageMode(this.config.includeImages);
    }

    /**
     * The document metadata a generator should write out: `ast.metadata` with
     * `config.metadataOverrides` applied on top, per field.
     *
     * Every generator must read metadata through here rather than touching `this.ast.metadata`
     * directly, so an override reaches all of them uniformly instead of one format at a time.
     *
     * Merged rather than replaced, so overriding one field doesn't blank the rest, and computed
     * fresh rather than cached on the AST: overrides are an output concern, and mutating
     * `ast.metadata` would leak one generation's settings into the next use of the same AST.
     * `custom` merges into `customProperties` so callers see one bucket regardless of origin.
     */
    protected get effectiveMetadata(): OfficeMetadata {
        const base = this.ast.metadata || {};
        const overrides = this.config.metadataOverrides;
        if (!overrides) return base;

        const { custom, language, ...named } = overrides;
        const merged: OfficeMetadata = { ...base };
        // Assign only the fields actually supplied - spreading `named` wholesale would write
        // `undefined` over parsed values for every field the caller left out.
        for (const [key, value] of Object.entries(named)) {
            if (value !== undefined) (merged as any)[key] = value;
        }
        // `language` lives in two places generators read from: the top-level `OfficeMetadata.language`
        // (what PdfParser sets and the DOCX/ODT generators read) and `nativeProperties.language` (what
        // EpubGenerator reads). Write the override to both so it reaches every generator.
        if (language !== undefined) {
            merged.language = language;
            merged.nativeProperties = { ...(base.nativeProperties || {}), language };
        }
        if (custom && Object.keys(custom).length > 0) {
            merged.customProperties = { ...(base.customProperties || {}), ...custom };
        }
        return merged;
    }

    /**
     * Reports caller-supplied `metadataOverrides.custom` entries that the destination format has
     * no way to represent (EPUB's OPF and RTF's `\info` both have fixed vocabularies).
     *
     * Warning rather than dropping silently: a caller who sets metadata and never sees it in the
     * output otherwise has no way to find out. Only `custom` keys are reported - the named fields
     * map onto something in every format that carries metadata at all.
     */
    protected warnUnrepresentableCustomMetadata(format: string): void {
        const custom = this.config.metadataOverrides?.custom;
        if (!custom) return;
        const keys = Object.keys(custom);
        if (keys.length === 0) return;
        this.warn(OfficeWarningType.METADATA_NOT_REPRESENTABLE, { keys, format });
    }

    /**
     * Retrieves the semantic mapping for a node, respecting the includeFormatting flag.
     * Per design requirements: Style mapping is bypassed if formatting is disabled.
     */
    protected getSemanticMapping(node: OfficeContentNode) {
        if (this.config.includeFormatting === false) {
            return undefined;
        }
        return this.styleMapper.getMapping(node);
    }

    /**
     * Entry point for generation.
     */
    abstract generate(): Promise<ConversionResult<D>>;

    /**
     * Centralized logic for handling the onNode callback.
     * Evaluates the callback and returns a result that tells the generator how to proceed.
     * 
     * @returns 
     * - `string`: Use this as the node's output, skip default processing.
     * - `false`: Skip this node and its subtree.
     * - `void`: Proceed with default processing.
     */
    /**
     * When set, memoizes each node's onNode verdict, so a generator that visits a node more than once
     * in a single pass (e.g. TextGenerator's spatial-layout collect followed by a flow fallback) still
     * asks the hook exactly once per node, per the documented contract. Left null by default.
     */
    protected onNodeMemo: WeakMap<OfficeContentNode, string | false | undefined> | null = null;

    protected async handleOnNode(node: OfficeContentNode): Promise<string | false | void> {
        if (this.onNodeMemo?.has(node)) return this.onNodeMemo.get(node);
        const result = await this.config.onNode(node);
        let verdict: string | false | undefined = result === false ? false : (typeof result === 'string' ? result : undefined);
        // A content control's placeholder prompt is not the document's text (TextMetadata.placeholder).
        if (verdict === undefined && isPlaceholderNode(node) && !this.config.includePlaceholderText) verdict = false;
        this.onNodeMemo?.set(node, verdict);
        return verdict;
    }

    /**
     * Recursively processes nodes and builds output.
     * 
     * @param node - The current node being processed
     * @param processor - A function that takes a node and its children's output and returns the node's output string.
     * @returns The generated string for this node and its subtree.
     */
    protected async processNodeRecursive(
        node: OfficeContentNode,
        processor: (node: OfficeContentNode, childrenOutput: string) => string | Promise<string>
    ): Promise<string> {
        // Every text-based generator funnels its whole traversal through here, so one check
        // makes `abortSignal` effective for all of them. Previously the signal was read once
        // before generation began and never again, which meant it could decline to start work
        // but could not stop work already underway - not much use against a document large
        // enough to be worth aborting. The check is a property read on an optional signal, so
        // the per-node cost is negligible.
        checkAbortSignal(this.config.abortSignal);
        const override = await this.handleOnNode(node);

        if (override === false) return '';
        if (typeof override === 'string') return override;
        // A shape drawn on a slide (ShapeMetadata) writes nothing: its text is in the nodes after it. (One
        // holding nodes, as a drawing could before, writes them.)
        if (node.type === 'drawing' && !node.children?.length) return '';

        let childrenOutput = '';
        if (node.children && !this.walkedByProcessor(node)) {
            // The output of the child before, for what goes between it and the next (see childSeparator).
            let previous = '';
            for (const child of node.children) {
                const piece = await this.processNodeRecursive(child, processor);
                if (!piece) continue;
                childrenOutput += this.childSeparator(previous, child, piece) + piece;
                previous = piece;
            }
        }

        if (node.notes && node.notes.length > 0) {
            if (node.type !== 'slide') {
                // Each note once, however many references share it: listed per reference, one note a
                // small document refers to thousands of times was written out that many times.
                for (const note of node.notes) {
                    if (this.collectedNoteSet.has(note)) continue;
                    this.collectedNoteSet.add(note);
                    this.collectedNotes.push(note);
                }
            }
        }

        let result = await processor(node, childrenOutput);

        if (node.type === 'slide' && node.notes && node.notes.length > 0) {
            for (const note of node.notes) {
                result += await this.processNodeRecursive(note, processor);
            }
        }

        return result;
    }

    /**
     * Whether the processor renders `node`'s children itself, so they are not rendered first for it
     * (it is given '' as their output). Rendered twice, a table nested in a table nested in a table
     * doubles the work at each level: 22 levels took seconds and 30 would take half an hour.
     */
    protected walkedByProcessor(_node: OfficeContentNode): boolean {
        return false;
    }

    /**
     * What goes between `previous`, the output of a node's child (empty for none), and `piece`, the
     * output of the child after it, `child`. Given that output alone, not all the node's output so
     * far: reading the end of a string built up piece by piece copies all of it each time.
     */
    protected childSeparator(_previous: string, _child: OfficeContentNode, _piece: string): string {
        return '';
    }

    /**
     * Helper to generate a unique ID (slug) from text. Letters and digits of every script are kept, as
     * GitHub's heading ids keep them (`#überblick`, `#введение`): only punctuation and symbols go.
     */
    protected slugify(text: string): string {
        return text
            .toLowerCase()
            .replace(/[^\p{L}\p{M}\p{N}\s_-]/gu, '')
            .replace(/[\s_-]+/g, '-')
            .replace(/^-+|-+$/g, '');
    }

    private readonly noteMarks = new NoteMarks();

    /**
     * The mark a note's references show (see NoteMarks): the one its source records, or its ordinal
     * within its kind. Never its key, which names it in ids and links only.
     */
    protected getNoteMark(note: OfficeContentNode): string {
        return this.noteMarks.markOf(note);
    }

    private noteFootnoteKeys = new Map<OfficeContentNode, string>();
    private usedFootnoteKeys = new Set<string>();
    private footnoteKeyCounter = 0;

    /**
     * Assigns a stable, unique reference key to a footnote/endnote node, reused for both
     * its inline reference marker and its collected definition. Source ids aren't reliably
     * unique across a document - DOCX/ODT number footnotes and endnotes in separate
     * sequences, so both can carry noteId "1" - so a source id is only reused when it
     * hasn't already been claimed; otherwise a sequential counter guarantees uniqueness.
     */
    protected getFootnoteKey(note: OfficeContentNode): string {
        const cached = this.noteFootnoteKeys.get(note);
        if (cached) return cached;

        const preferred = (note.metadata as any)?.noteId;
        // The id is document-supplied and lands inside `[^...]` in Markdown, where the parser's
        // own recognizer accepts everything but `]` - so a note id carrying markup round-tripped
        // into the output verbatim. Accept it only when it looks like a label; otherwise fall
        // through to the sequential counter, which is always safe. Real ids are numeric (DOCX),
        // `ftn1`-shaped (ODT) or short slugs (`fn1`), and a Markdown label like `[^my note]`
        // still qualifies, so the fallback fires only for genuinely hostile input.
        const isLabelLike = typeof preferred === 'string' && /^[A-Za-z0-9_.:-][A-Za-z0-9 _.:-]*$/.test(preferred);
        let key: string;
        if (preferred && isLabelLike && !this.usedFootnoteKeys.has(preferred)) {
            key = preferred;
        } else {
            do {
                this.footnoteKeyCounter++;
                key = String(this.footnoteKeyCounter);
            } while (this.usedFootnoteKeys.has(key));
        }

        this.usedFootnoteKeys.add(key);
        this.noteFootnoteKeys.set(note, key);
        return key;
    }

    /**
     * Takes `key` from the keys getFootnoteKey gives notes, for an id a writer derives from a note's key
     * (a second reference's `KEY-2`), so no note is given it later. False when a note or an earlier claim
     * holds it.
     */
    protected claimFootnoteKey(key: string): boolean {
        if (this.usedFootnoteKeys.has(key)) return false;
        this.usedFootnoteKeys.add(key);
        return true;
    }

    /**
     * True when every content-bearing text descendant satisfies `test` - i.e. the property is
     * uniform across the whole node and therefore says nothing the node type does not already say.
     *
     * Used to decide whether a heading's or header row's inherited formatting can be dropped. The
     * distinction matters: an ODF heading whose paragraph style is bold and 14pt yields a heading
     * where *every* run is bold and 14pt, and re-emitting that gives `# **Heading**` in Markdown
     * and, worse in RTF/HTML, an inner font-size that overrides the heading's own and visibly
     * shrinks it. But `# Normal **Bold** Normal` is an author contrasting one word against the
     * rest, and dropping that would discard real meaning. Only the uniform case is safe.
     *
     * Returns false when there is no text to judge, so an empty or image-only node never triggers
     * suppression.
     */
    protected hasUniformFormatting(
        node: OfficeContentNode,
        test: (formatting: OfficeContentNode['formatting']) => boolean
    ): boolean {
        let sawText = false;
        const walk = (n: OfficeContentNode): boolean => {
            if (n.type === 'text') {
                // Whitespace-only runs carry no visible formatting either way, so they neither
                // count as evidence nor veto - otherwise a stray unformatted space between two
                // bold runs would defeat the check on almost every real heading.
                if (!(n.text || '').trim()) return true;
                sawText = true;
                return test(n.formatting);
            }
            return (n.children ?? []).every(walk);
        };
        const uniform = (node.children ?? []).every(walk);
        return sawText && uniform;
    }

    /**
     * Recursively extracts plain text from a node and its children. A source comment (the author's
     * hidden note, `<!-- ... -->`) is not text, so it contributes nothing.
     */
    protected getNodeText(node: OfficeContentNode): string {
        if (isSourceComment(node)) return '';
        // A content control's placeholder prompt is no text of the node (TextMetadata.placeholder): a
        // heading holding one was given an id made from it. A paragraph's `.text` leaves prompts out,
        // so to write them it is read from its runs.
        const prompts = this.config.includePlaceholderText;
        if (isPlaceholderNode(node) && !prompts) return '';
        if (node.text && !(prompts && node.children?.some(isPlaceholderNode))) return node.text;
        if (node.children) {
            return node.children.map(c => this.getNodeText(c)).join('');
        }
        return '';
    }

    /**
     * Reports a warning to the user and collects it for the final result.
     */
    protected warn(type: OfficeWarningType, info?: any, node?: OfficeContentNode): void {
        const message = getWarningMessage(type, info);
        const issue: OfficeIssue = {
            type: 'warning',
            code: type,
            message,
            node,
            details: info
        };
        this.messages.push(issue);
        this.config.onWarning(issue);
    }
}

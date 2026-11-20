import { ChartAxis, ChartData, ChartType, OfficeAttachment, OfficeParserConfig } from "../types";
import { parseXmlString, getElementsByTagName, getDirectChildren, getFirstElementByTagName, getChildElements } from "./xmlUtils";
import { repeatPreview, takeRepeats } from "./repeatUtils";
import { excelShownText } from "./excelFormat";
import { lookupTable } from "./lookupUtils";

/**
 * Extracts a single text element located at:
 * c:title -> c:tx -> c:rich -> a:p -> a:r -> a:t
 * OR c:tx -> c:strRef -> c:strCache -> c:pt -> c:v
 * @param el Chart element
 * @param tagName The tag to search for (e.g., "c:title", "c:tx")
 */
const extractOpenXmlRichText = (el: Element, tagName: string): string | undefined => {
    const target = (el.localName === tagName || el.tagName === tagName) ? el : getFirstElementByTagName(el, tagName);
    if (!target) return undefined;

    // 1. Try c:rich or a:p (standard rich text)
    // (Through the helpers, which read nested content once: see getElementsByTagName.)
    const richNodes = getElementsByTagName(target, "c:rich");
    const pNodes = getElementsByTagName(target, "a:p");
    const textContainers = richNodes.length > 0 ? Array.from(richNodes) : Array.from(pNodes);

    if (textContainers.length > 0) {
        let acc = "";
        for (const container of textContainers) {
            const tNodes = getElementsByTagName(container, "a:t");
            for (let i = 0; i < tNodes.length; i++) {
                acc += (tNodes[i].textContent || "") + " ";
            }
        }
        if (acc.trim()) return acc.trim();
    }

    // 2. Try c:v (cached values/strings)
    const vNode = getFirstElementByTagName(target, "c:v");
    if (vNode && vNode.textContent) {
        return vNode.textContent.trim() || undefined;
    }

    return undefined;
};

/**
 * The most series one chart's table makes, and values it holds in all. A chart's table can repeat a
 * cell any number of times (`table:number-columns-repeated`), and a chart of 1 KB asking for a hundred
 * million ran the process out of memory.
 */
const MAX_CHART_SERIES = 10_000;
export const MAX_CHART_VALUES = 1_000_000;

/**
 * A chart's text: its title (when given), its category labels once, then each series' name and values, and a
 * scatter or bubble series' own X values (unless they are the labels, the very list: the first series') and
 * bubble sizes after them. The labels were written again for every series, so a small chart of many series
 * and many labels made text of their product (gigabytes from a few KB).
 */
export function chartRawTexts(title: string | undefined, dataSets: ChartData['dataSets'], labels: string[]): string[] {
    const raw: string[] = [];
    if (title) raw.push(title);
    for (const label of labels) raw.push(label);
    for (const ds of dataSets) {
        if (ds.name) raw.push(ds.name);
        for (const value of ds.values) raw.push(value);
        if (ds.xValues && ds.xValues !== labels) for (const value of ds.xValues) raw.push(value);
        if (ds.bubbleSizes) for (const value of ds.bubbleSizes) raw.push(value);
    }
    return raw;
}

/** A point's index as a chart states it (`c:pt idx`, `c:dLbl/c:idx`): a whole number from 0, else undefined. */
const pointIndex = (value: string | null | undefined): number | undefined => {
    if (!value || !/^\d+$/.test(value)) return undefined;
    const index = Number(value);
    return Number.isSafeInteger(index) ? index : undefined;
};

/** Whether `indexes` gives each of `length` entries a point index, in ascending order (see ChartData.labelIndexes). */
export const isPointIndexList = (indexes: unknown, length: number): indexes is number[] => {
    if (!Array.isArray(indexes) || indexes.length !== length) return false;
    let last = -1;
    for (const index of indexes) {
        if (!Number.isSafeInteger(index) || index <= last) return false;
        last = index;
    }
    return true;
};

/** Collects the point indexes of a list's entries. Gives none when an entry has no index, the indexes do not ascend, or the list is empty. */
const pointIndexes = () => {
    const indexes: number[] = [];
    let positioned = true;
    return {
        add(index: number | undefined) {
            if (!positioned) return;
            if (index === undefined || index <= (indexes[indexes.length - 1] ?? -1)) positioned = false;
            else indexes.push(index);
        },
        get: (): number[] | undefined => positioned && indexes.length > 0 ? indexes : undefined,
    };
};

/**
 * The point indexes that hold a label at any category level, or a value, ascending: writers write one row or
 * group per index and pair each list's entries with it (see atPosition), so an outer label covering only
 * blank points has a row of its own. Undefined when the labels or a series' values have entries and no valid
 * point indexes; writers then pair the lists by place. An outer level of no valid point indexes adds none (it
 * is written one label a row, see outerLevelRows). The result holds no more indexes than the lists hold
 * entries, whatever `pointCount` says.
 */
export function chartPositions(data: ChartData): number[] | undefined {
    const lists: [unknown, unknown][] = [[data.labels, data.labelIndexes]];
    if (Array.isArray(data.dataSets)) {
        for (const series of data.dataSets) if (series && typeof series === "object") lists.push([series.values, series.valueIndexes]);
    }
    if (Array.isArray(data.labelLevels)) {
        for (const level of data.labelLevels.slice(1)) {
            const labels = Array.isArray(level?.labels) ? level.labels : [];
            if (labels.length > 0 && isPointIndexList(level.labelIndexes, labels.length)) lists.push([labels, level.labelIndexes]);
        }
    }
    const positions: number[] = [];
    let positioned = false;
    for (const [list, indexes] of lists) {
        const length = Array.isArray(list) ? list.length : 0;
        if (indexes === undefined && length === 0) continue;
        if (!isPointIndexList(indexes, length)) return undefined;
        positioned = true;
        for (const index of indexes) positions.push(index);
    }
    if (!positioned) return undefined;
    positions.sort((a, b) => a - b);
    return positions.filter((position, i) => i === 0 || position !== positions[i - 1]);
}

/**
 * The point indexes a scatter or bubble chart's table writes a row for: those of every series' X values,
 * values and bubble sizes, ascending, each once (see chartPositions). Undefined when a list with entries has
 * no valid point indexes: rows then pair the lists by place.
 */
export function xyPositions(data: ChartData): number[] | undefined {
    const positions: number[] = [];
    for (const series of Array.isArray(data.dataSets) ? data.dataSets : []) {
        if (!series || typeof series !== "object") continue;
        for (const [list, indexes] of [[series.xValues, series.xValueIndexes], [series.values, series.valueIndexes], [series.bubbleSizes, series.bubbleSizeIndexes]] as [unknown, unknown][]) {
            const length = Array.isArray(list) ? list.length : 0;
            if (indexes === undefined && length === 0) continue;
            if (!isPointIndexList(indexes, length)) return undefined;
            for (const index of indexes) positions.push(index);
        }
    }
    positions.sort((a, b) => a - b);
    return positions.filter((position, i) => i === 0 || position !== positions[i - 1]);
}

/**
 * Whether a chart is a scatter or bubble chart as its data says: a series holds X values or bubble sizes of
 * its own. (Not by its chart type alone: an ODF scatter chart's X values are its labels.)
 */
export const isXyChart = (data: ChartData): boolean =>
    Array.isArray(data.dataSets) && data.dataSets.some(series => (Array.isArray(series?.xValues) && series.xValues.length > 0) || (Array.isArray(series?.bubbleSizes) && series.bubbleSizes.length > 0));

/**
 * What a writer writes for a chart it does not write again (past the chart-copy budget): `[Chart: <title>]`, the
 * title cut to its start (repeatPreview), else `[Chart: <attachment name>]`, the form a chart with no data takes.
 */
export const chartReferenceText = (data: ChartData | undefined, name: string | undefined): string => {
    const title = typeof data?.title === 'string' ? data.title.trim() : '';
    return `[Chart: ${title ? repeatPreview(title) : name ?? ''}]`;
};

/**
 * The row each outer category label stands on, the one rule every writer places them by (outerLevelRows,
 * the HTML chart script): for each level past the innermost (`labelLevels` after its first), innermost
 * first, the row of each of its labels among `count` rows, row i holding point `positions[i]` (point i where
 * `positions` is undefined); -1 for a label no row has (one at a point no row has, or past the rows). A label
 * stands on the row of the point it starts at; a level of no point indexes holds its labels in order, one a
 * row. `whole` says whether every outer label has a row. Linear in the rows and the labels.
 */
export function outerLabelRows(data: ChartData, count: number, positions: readonly number[] | undefined): { rows: number[][]; whole: boolean } {
    const levels = Array.isArray(data.labelLevels) ? data.labelLevels.slice(1) : [];
    let rowOf: Map<number, number> | undefined;
    let whole = true;
    const rows = levels.map(level => {
        const labels = Array.isArray(level?.labels) ? level.labels : [];
        const indexes = isPointIndexList(level?.labelIndexes, labels.length) ? level.labelIndexes : undefined;
        const placed: number[] = [];
        for (let k = 0; k < labels.length; k++) {
            let row = indexes ? indexes[k] : k;
            if (indexes && positions) {
                if (!rowOf) {
                    rowOf = new Map();
                    for (let r = 0; r < count && r < positions.length; r++) rowOf.set(positions[r], r);
                }
                row = rowOf.get(indexes[k]) ?? -1;
            }
            if (row >= count) row = -1;
            if (row < 0) whole = false;
            placed.push(row);
        }
        return placed;
    });
    return { rows, whole };
}

/**
 * A chart's outer category levels as writers show them: for each level past the innermost (`labelLevels`
 * after its first), outermost first, its label at each of `count` rows, at the row of the point it starts
 * at and blank below until the next (as the sheet behind the chart holds them), placed by outerLabelRows
 * (row i holds point `positions[i]`, or point i). `whole` says whether every outer label has a row. Linear
 * in the rows and the labels. No levels for a chart of one level.
 */
export function outerLevelRows(data: ChartData, count: number, positions: readonly number[] | undefined): { levels: string[][]; whole: boolean } {
    const levels = Array.isArray(data.labelLevels) ? data.labelLevels.slice(1) : [];
    const { rows, whole } = outerLabelRows(data, count, positions);
    const written = rows.map((placed, i) => {
        const level = levels[i];
        const labels = Array.isArray(level?.labels) ? level.labels : [];
        const cells: string[] = new Array(count).fill('');
        placed.forEach((row, k) => { if (row >= 0) cells[row] = String(labels[k] ?? ''); });
        return cells;
    });
    return { levels: written.reverse(), whole };
}

/**
 * What a writer shows for a chart's list: its shown texts (`shownValues`, `shownLabels`, the text the
 * chart shows for each entry) where they line up with the list, else the list itself.
 */
export function shownOr(shown: unknown, list: unknown): unknown[] {
    const raw = Array.isArray(list) ? list : [];
    return Array.isArray(shown) && shown.length === raw.length ? shown : raw;
}

/** The entry of `list` whose point index is `position`, found by binary search of `indexes`; undefined when none has it. */
export function atPosition(list: unknown, indexes: unknown, position: number): unknown {
    if (!Array.isArray(list) || !Array.isArray(indexes)) return undefined;
    let low = 0, high = indexes.length - 1;
    while (low <= high) {
        const mid = (low + high) >> 1;
        if (indexes[mid] === position) return list[mid];
        if (indexes[mid] < position) low = mid + 1;
        else high = mid - 1;
    }
    return undefined;
}

/** A repeat count as a document states it: at least 1, NaN as 1. */
const repeats = (value: string | null): number => {
    const n = parseInt(value || '1', 10);
    return Number.isFinite(n) && n > 1 ? n : 1;
};

/**
 * A series' values or categories: their texts, the point index of each (`indexes`, see readCache), and,
 * with a number format, that format and each text as the chart shows it (`shown`, lining up with `texts`).
 */
interface OpenXmlCache {
    texts: string[];
    indexes?: number[];
    formatCode?: string;
    shown?: string[];
}

/**
 * Reads the texts of every `c:v` below a series' `c:val`, `c:yVal`, `c:cat` or `c:xVal` (at most `limit`,
 * blank ones left out), the point index of each (its parent's `c:pt idx`), and, where the cache has a number
 * format (`c:formatCode`, not General), each as the chart shows it (see ChartData's `shownValues`; a
 * point's own format wins). Gives no indexes for a `c:v` outside a `c:pt`, or for a cache of several
 * levels, whose indexes start again at 0 in each level. `shown`, where given, shares `indexes`; like
 * `indexes`, it is given only with entries.
 */
const readCache = (node: Element, limit: number, date1904: boolean): OpenXmlCache => {
    const cache = getElementsByTagName(node, "c:numCache")[0] ?? getElementsByTagName(node, "c:numLit")[0];
    const code = cache ? getChildElements(cache, "c:formatCode")[0]?.textContent?.trim() : undefined;
    const formatCode = code && code.length <= 255 && !/^general$/i.test(code) ? code : undefined;
    const vNodes = getElementsByTagName(node, "c:v");
    const texts: string[] = [];
    const shown: string[] = [];
    const indexes = pointIndexes();
    for (let j = 0; j < vNodes.length && texts.length < limit; j++) {
        const text = vNodes[j].textContent?.trim();
        if (!text) continue;
        texts.push(text);
        const pt = vNodes[j].parentNode as Element | null;
        indexes.add(pt?.nodeType === 1 && pt.tagName === "c:pt" ? pointIndex(pt.getAttribute("idx")) : undefined);
        if (!formatCode) continue;
        const pointCode = pt?.localName === "pt" ? pt.getAttribute("formatCode") : null;
        shown.push(excelShownText(text, pointCode || formatCode, date1904) ?? text);
    }
    const positions = indexes.get();
    return { texts, ...(positions && { indexes: positions }), ...(formatCode && { formatCode, ...(shown.length > 0 && { shown }) }) };
};

/** The most levels a chart's categories are read with (OpenXML's `c:multiLvlStrCache` levels, `c:lvl`). */
export const MAX_CHART_LEVELS = 16;

/**
 * A series' categories: their labels (the innermost level, as for a chart of one level) and, for a cache of
 * more than one level (quarters under years, `c:multiLvlStrCache`), each level, innermost first, read with
 * its own point indexes. At most MAX_CHART_LEVELS levels, whose labels share one MAX_CHART_VALUES budget.
 * A multi-level cache's `c:v` were read as one list, every level's labels run together (#139).
 */
const readCategories = (catNode: Element, date1904: boolean): { labels: OpenXmlCache; levels?: OpenXmlCache[] } => {
    const multiLevel = getElementsByTagName(catNode, "c:multiLvlStrCache")[0];
    const lvls = multiLevel ? getDirectChildren(multiLevel, "c:lvl").slice(0, MAX_CHART_LEVELS) : [];
    if (lvls.length < 2) return { labels: readCache(lvls[0] ?? catNode, MAX_CHART_VALUES, date1904) };
    let left = MAX_CHART_VALUES;
    const levels = lvls.map(lvl => {
        const level = readCache(lvl, left, date1904);
        left -= level.texts.length;
        return level;
    });
    return { labels: levels[0], levels };
};

/** The most axes a chart is read with (OpenXML's `c:catAx`, `c:dateAx`, `c:valAx`, `c:serAx`; ODF's `chart:axis`). */
export const MAX_CHART_AXES = 64;
/** The most axes a series is drawn on (its `axisIds`). */
export const MAX_SERIES_AXES = 4;

/** The kinds of axis a chart has (ChartAxis.kind). */
export const CHART_AXIS_KINDS: ReadonlySet<string> = new Set(['category', 'date', 'value', 'series']);
/** The sides an axis is drawn on (ChartAxis.position). */
export const CHART_AXIS_POSITIONS: ReadonlySet<string> = new Set(['bottom', 'left', 'right', 'top']);
/** The kinds of plot group a series is drawn as (ChartType). */
export const CHART_TYPES: ReadonlySet<string> = new Set(['area', 'bar', 'bubble', 'doughnut', 'line', 'ofPie', 'pie', 'radar', 'scatter', 'stock', 'surface']);
/** The ways a chart draws a blank point (ChartData.displayBlanksAs). */
export const CHART_BLANKS: ReadonlySet<string> = new Set(['gap', 'span', 'zero']);

/** Each OpenXML `c:dispBlanksAs`'s way of drawing a blank point. */
const OPENXML_BLANKS: Record<string, NonNullable<ChartData['displayBlanksAs']>> = lookupTable({ gap: 'gap', span: 'span', zero: 'zero' });

/** Each OpenXML axis element's kind. */
const OPENXML_AXIS_KINDS: Record<string, ChartAxis['kind']> = lookupTable({ catAx: 'category', dateAx: 'date', valAx: 'value', serAx: 'series' });
/** Each OpenXML `c:axPos`'s side. */
const OPENXML_AXIS_POSITIONS: Record<string, NonNullable<ChartAxis['position']>> = lookupTable({ b: 'bottom', l: 'left', r: 'right', t: 'top' });
/** Each OpenXML plot group's chart type, 3D and flat alike. */
const OPENXML_CHART_TYPES: Record<string, ChartType> = lookupTable({
    barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', pieChart: 'pie', pie3DChart: 'pie',
    doughnutChart: 'doughnut', areaChart: 'area', area3DChart: 'area', scatterChart: 'scatter', bubbleChart: 'bubble',
    radarChart: 'radar', stockChart: 'stock', surfaceChart: 'surface', surface3DChart: 'surface', ofPieChart: 'ofPie',
});

/** An OOXML boolean (CT_Boolean): true unless its `val` is 0 or false; no element, undefined. */
const openXmlFlag = (element: Element | undefined): boolean | undefined => {
    if (!element) return undefined;
    const value = element.getAttribute("val");
    return value === null || !/^(?:0|false)$/i.test(value.trim());
};

/**
 * A chart's plot area as OpenXML holds it: every axis (at most MAX_CHART_AXES), in the order it lists them,
 * and each plot group's chart type and axis ids (at most MAX_SERIES_AXES), in document order.
 */
const readOpenXmlPlotArea = (plotArea: Element | undefined): { axes: ChartAxis[]; groups: Map<Node, { chartType: ChartType; axisIds: string[] }> } => {
    const axes: ChartAxis[] = [];
    const groups = new Map<Node, { chartType: ChartType; axisIds: string[] }>();
    const children = plotArea?.childNodes;
    for (let i = 0; children && i < children.length; i++) {
        const child = children[i];
        if (child.nodeType !== 1) continue;
        const element = child as Element;
        const local = element.localName || element.tagName.replace(/^.*:/, '');
        const kind = OPENXML_AXIS_KINDS[local];
        if (kind) {
            if (axes.length >= MAX_CHART_AXES) continue;
            const id = getDirectChildren(element, "c:axId")[0]?.getAttribute("val") ?? '';
            const position = OPENXML_AXIS_POSITIONS[getDirectChildren(element, "c:axPos")[0]?.getAttribute("val") ?? ''];
            const titleEl = getDirectChildren(element, "c:title")[0];
            const title = titleEl ? extractOpenXmlRichText(titleEl, "c:title") : undefined;
            const deleted = openXmlFlag(getDirectChildren(element, "c:delete")[0]);
            axes.push({ id, kind, ...(position && { position }), ...(title !== undefined && { title }), ...(deleted !== undefined && { deleted }) });
            continue;
        }
        const chartType = OPENXML_CHART_TYPES[local];
        if (!chartType) continue;
        const axisIds = getDirectChildren(element, "c:axId").slice(0, MAX_SERIES_AXES).map(axId => axId.getAttribute("val") ?? '').filter(Boolean);
        groups.set(element, { chartType, axisIds });
    }
    return { axes, groups };
};

/**
 * Extracts fully structured chart data from OpenXML (PPTX, XLSX) chart XML.
 * @param xmlBuffer Chart XML buffer
 */
const extractOpenXmlChartData = (xmlBuffer: Buffer, config?: OfficeParserConfig): ChartData => {
    const xml = xmlBuffer.toString("utf8");
    const dom = parseXmlString(xml, { config });
    const root = dom.documentElement as Element;
    if (!root) return { title: undefined, xAxisTitle: undefined, yAxisTitle: undefined, dataSets: [], labels: [], rawTexts: [] };

    // The chart's own title (`c:chart`'s), not the first anywhere: a chart of none took an axis's.
    const chartEl = getDirectChildren(root, "c:chart")[0] ?? root;
    const titleEl = getDirectChildren(chartEl, "c:title")[0];
    let title = titleEl ? extractOpenXmlRichText(titleEl, "c:title") : undefined;

    // Every axis, and each plot group's chart type and axes (#143).
    const { axes, groups } = readOpenXmlPlotArea(getDirectChildren(chartEl, "c:plotArea")[0]);
    const displayBlanksAs = OPENXML_BLANKS[getDirectChildren(chartEl, "c:dispBlanksAs")[0]?.getAttribute("val") ?? ''];
    const axisById = new Map<string, ChartAxis>();
    for (const axis of axes) if (axis.id && !axisById.has(axis.id)) axisById.set(axis.id, axis);

    // Axis titles: those of the primary axes, the axes the chart lists first (as Office and LibreOffice take
    // them; LibreOffice writes a secondary group first), read through the plot group drawn on them, its X axis
    // then its value axis (a date axis is an X axis, and a scatter chart's horizontal value axis its X axis,
    // not its Y). That group is the one whose axes the chart lists first: each group's axes, each once,
    // ranked by where the chart lists them (an axis it does not list, or none, after every one it does),
    // compared from the first listed, the first such group in document order on a tie. So a group sharing
    // the first axis with another is told from it by its other axis. Where the group names no axis the chart
    // has, the first category axis's and the first value axis's.
    let xAxisTitle: string | undefined = undefined;
    let yAxisTitle: string | undefined = undefined;
    const listedAt = new Map<string, number>();
    axes.forEach((axis, i) => { if (axis.id && !listedAt.has(axis.id)) listedAt.set(axis.id, i); });
    const ranksOf = (ids: string[]) => [...new Set(ids)].map(id => listedAt.get(id) ?? axes.length).sort((a, b) => a - b);
    const listedBefore = (a: number[], b: number[]) => {
        for (let i = 0; i < MAX_SERIES_AXES; i++) {
            const x = a[i] ?? axes.length, y = b[i] ?? axes.length;
            if (x !== y) return x < y;
        }
        return false;
    };
    let primaryAxes: string[] | undefined;
    let primaryRanks: number[] | undefined;
    for (const group of groups.values()) {
        const ranks = ranksOf(group.axisIds);
        if (!primaryRanks || listedBefore(ranks, primaryRanks)) { primaryAxes = group.axisIds; primaryRanks = ranks; }
    }
    const [firstX, firstY] = (primaryAxes ?? []).map(id => axisById.get(id));
    if (firstX || firstY) {
        xAxisTitle = firstX?.title;
        yAxisTitle = firstY?.title;
    } else {
        const catAxes = getElementsByTagName(root, "c:catAx");
        if (catAxes.length > 0) xAxisTitle = extractOpenXmlRichText(catAxes[0] as Element, "c:title");
        const valAxes = getElementsByTagName(root, "c:valAx");
        if (valAxes.length > 0) yAxisTitle = extractOpenXmlRichText(valAxes[0] as Element, "c:title");
    }

    // Extract Series (dataSets)
    // The outermost series, at most MAX_CHART_SERIES of them, holding at most MAX_CHART_VALUES values
    // in all: series nested in series each read every value below them (22 KB, out of memory). X values
    // and bubble sizes are read within budgets of their own, as values and categories are.
    const seriesNodes = getElementsByTagName(root, "c:ser").slice(0, MAX_CHART_SERIES);
    let valuesLeft = MAX_CHART_VALUES, xValuesLeft = MAX_CHART_VALUES, sizesLeft = MAX_CHART_VALUES;
    const dataSets: ChartData['dataSets'] = [];
    let sharedLabels: OpenXmlCache = { texts: [] };
    let labelLevels: OpenXmlCache[] | undefined;
    // Whether a series' categories were kept: the first series holding labels at any level.
    let categoriesRead = false;
    // Whether the labels are a series' X values or the innermost level: lists the chart gives elsewhere.
    let labelsShared = false;
    let pointCount: number | undefined;
    // `c:date1904` with no `val` means true (OOXML's CT_Boolean default).
    const date1904Val = getChildElements(root, "c:date1904")[0]?.getAttribute("val");
    const date1904 = date1904Val !== undefined && !/^(?:0|false)$/i.test((date1904Val ?? "").trim());

    for (let i = 0; i < seriesNodes.length; i++) {
        const ser = seriesNodes[i] as Element;

        // DataSet Name (c:tx)
        const name = extractOpenXmlRichText(ser, "c:tx");

        // Values (c:val)
        const valNode = getChildElements(ser, "c:val")[0] || getChildElements(ser, "c:yVal")[0];
        const valCache: OpenXmlCache = valNode ? readCache(valNode, valuesLeft, date1904) : { texts: [] };
        valuesLeft -= valCache.texts.length;

        // A scatter or bubble series' own X values (c:xVal) and bubble sizes (c:bubbleSize), each list read
        // within a budget of its own, as values and categories are (#144): every series' X values were taken
        // to be the first's, and its sizes were not read.
        const xNode = getChildElements(ser, "c:xVal")[0];
        const xCache = xNode ? readCache(xNode, xValuesLeft, date1904) : undefined;
        xValuesLeft -= xCache?.texts.length ?? 0;
        const sizeNode = getChildElements(ser, "c:bubbleSize")[0];
        const sizeCache = sizeNode ? readCache(sizeNode, sizesLeft, date1904) : undefined;
        sizesLeft -= sizeCache?.texts.length ?? 0;

        // Point Labels (data labels)
        const pointLabels: string[] = [];
        const pointLabelPoints = pointIndexes();
        const dLbls = getElementsByTagName(ser, "c:dLbl");
        for (let j = 0; j < dLbls.length; j++) {
            const lbl = extractOpenXmlRichText(dLbls[j], "c:tx");
            if (!lbl) continue;
            pointLabels.push(lbl);
            pointLabelPoints.add(pointIndex(getDirectChildren(dLbls[j], "c:idx")[0]?.getAttribute("val")));
        }

        // Categories (labels) - c:cat or c:xVal
        const catNode = getChildElements(ser, "c:cat")[0] || xNode;
        if (catNode && !categoriesRead) {
            // (A series' X values are its labels' very list, so the chart's text holds them once.)
            const { labels, levels } = catNode === xNode && xCache ? { labels: xCache, levels: undefined } : readCategories(catNode, date1904);
            // Kept where any level has labels: an empty innermost level under outer ones lost the outer ones.
            if (labels.texts.length > 0 || levels?.some(level => level.texts.length > 0)) {
                categoriesRead = true;
                labelsShared = catNode === xNode || levels !== undefined;
                sharedLabels = labels;
                labelLevels = levels;
                pointCount = pointIndex(getElementsByTagName(catNode, "c:ptCount")[0]?.getAttribute("val"));
            }
        }

        const pointLabelIndexes = pointLabelPoints.get();
        // The plot group the series is in: its chart type, and the axes it is drawn on.
        const group = ser.parentNode ? groups.get(ser.parentNode) : undefined;
        dataSets.push({
            ...(group && { chartType: group.chartType }),
            ...(group?.axisIds.length && { axisIds: [...group.axisIds] }),
            name,
            values: valCache.texts,
            ...(valCache.indexes && { valueIndexes: valCache.indexes }),
            ...(valCache.formatCode && { formatCode: valCache.formatCode, ...(valCache.shown && { shownValues: valCache.shown }) }),
            pointLabels,
            ...(pointLabelIndexes && { pointLabelIndexes }),
            // (An optional list is given only with entries, as the AST's lists are normalized.)
            ...(xCache && xCache.texts.length > 0 && { xValues: xCache.texts, ...(xCache.indexes && { xValueIndexes: xCache.indexes }) }),
            ...(sizeCache && sizeCache.texts.length > 0 && { bubbleSizes: sizeCache.texts, ...(sizeCache.indexes && { bubbleSizeIndexes: sizeCache.indexes }) }),
        });
    }

    // A title of no text of its own is the one Office draws for it: a chart of one series its name, else "Chart
    // Title" (as PowerPoint, Excel and LibreOffice show it; it was left out).
    if (titleEl && title === undefined && !getDirectChildren(titleEl, "c:tx").length) title = dataSets.length === 1 && dataSets[0].name ? dataSets[0].name : "Chart Title";

    // (Every level's labels, innermost first. Read before the labels are copied: a series' X values that are
    // the labels, the very list, are written once.)
    const rawTexts = chartRawTexts(undefined, dataSets, labelLevels ? labelLevels.flatMap(level => level.texts) : sharedLabels.texts);
    // Labels that are a series' X values or the innermost level are a copy: a change to one list changed the other.
    const labelIndexes = sharedLabels.indexes && labelsShared ? [...sharedLabels.indexes] : sharedLabels.indexes;

    return {
        title,
        xAxisTitle,
        yAxisTitle,
        dataSets,
        labels: labelsShared ? [...sharedLabels.texts] : sharedLabels.texts,
        ...(labelIndexes && { labelIndexes }),
        ...(pointCount !== undefined && { pointCount }),
        ...(displayBlanksAs && { displayBlanksAs }),
        ...(sharedLabels.formatCode && { labelFormatCode: sharedLabels.formatCode, ...(sharedLabels.shown && { shownLabels: sharedLabels.shown }) }),
        ...(labelLevels && { labelLevels: labelLevels.map(level => ({ labels: level.texts, ...(level.indexes && { labelIndexes: level.indexes }) })) }),
        ...(axes.length > 0 && { axes }),
        rawTexts
    };
};

/**
 * Extracts structured chart data from ODF (ODP, ODS) chart content.xml.
 * @param xmlBuffer Chart XML buffer
 */
/**
 * How many of `copies` of a repeated cell's `text` are whole: every series a repeated cell fills
 * carries its text, and the chart's text is written with each, so 1.3 KB repeating one 100 KB value
 * 2,500 times made 250 MB. Copies after the first are charged to the document's repeated-content
 * budget (see repeatUtils); the rest hold the text's start.
 */
const copiesWhole = (config: OfficeParserConfig | undefined, copies: number, text: string | undefined): number =>
    copies <= 1 || !text ? copies : 1 + takeRepeats(config ?? {}, copies - 1, text.length);

/** What holds a table's rows besides the table itself. */
const ROW_GROUPS = new Set(["table:table-header-rows", "table:table-rows", "table:table-row-group"]);

/** A table's rows in document order, those in row groups (nested ones too) included, each read once. */
const tableRows = (table: Element): Element[] => {
    const rows: Element[] = [];
    // Children still to read, last first: a group's are put in its place.
    const pending: Node[] = [];
    for (let i = table.childNodes.length - 1; i >= 0; i--) pending.push(table.childNodes[i]);
    while (pending.length > 0) {
        const node = pending.pop()!;
        if (node.nodeType !== 1) continue;
        const element = node as Element;
        if (element.tagName === "table:table-row") rows.push(element);
        else if (ROW_GROUPS.has(element.tagName)) {
            for (let i = element.childNodes.length - 1; i >= 0; i--) pending.push(element.childNodes[i]);
        }
    }
    return rows;
};

/** Each ODF `chart:class`'s chart type. */
const ODF_CHART_TYPES: Record<string, ChartType> = lookupTable({
    bar: 'bar', line: 'line', area: 'area', circle: 'pie', ring: 'doughnut', scatter: 'scatter', bubble: 'bubble',
    radar: 'radar', 'filled-radar': 'radar', stock: 'stock', surface: 'surface',
});
/** Each ODF `chart:treat-empty-cells`'s way of drawing a blank point (`ignore` continues the plot past it). */
const ODF_BLANKS: Record<string, NonNullable<ChartData['displayBlanksAs']>> = lookupTable({ 'leave-gap': 'gap', ignore: 'span', 'use-zero': 'zero' });

const extractOdfChartData = (xmlBuffer: Buffer, config?: OfficeParserConfig): ChartData => {
    const xml = xmlBuffer.toString("utf8");
    const dom = parseXmlString(xml, { config });

    const chart = getElementsByTagName(dom, "chart:chart")[0] || dom.documentElement;
    // (The chart's own title, not the first anywhere: a chart of none took an axis's.)
    const titleNode = getDirectChildren(chart as Element, "chart:title")[0];
    const title = titleNode ? getElementsByTagName(titleNode, "text:p")[0]?.textContent || undefined : undefined;

    const table = getElementsByTagName(chart, "table:table")[0];
    const dataSets: ChartData['dataSets'] = [];
    const labels: string[] = [];
    // Each data row is a point (its position the row's, from 0): a row with a blank category keeps no label and a
    // blank value no value, and each label and value its point's index, as OpenXML charts keep them (#136). A
    // blank category dropped its label but not its values, which then paired with the next row's category, and a
    // blank value was read as the text "NaN" LibreOffice writes for it.
    const labelIndexes: number[] = [];
    let pointCount: number | undefined;

    if (table) {
        // Chart with embedded data table (common in ODP presentations). Its rows in document order,
        // including those in row groups: LibreOffice writes the data rows in `table:table-rows`, and
        // read from the table's own rows alone, its charts had series names and no values.
        const rows = tableRows(table);

        if (rows.length > 0) {
            // Header row for series names
            const headerCells = getDirectChildren(rows[0], "table:table-cell");
            for (let j = 1; j < headerCells.length && dataSets.length < MAX_CHART_SERIES; j++) {
                // A repeated cell is as many series as it repeats, up to the chart's limit.
                const colsRepeated = Math.min(repeats(headerCells[j].getAttribute("table:number-columns-repeated")), MAX_CHART_SERIES - dataSets.length);
                const name = getDirectChildren(headerCells[j], "text:p")[0]?.textContent || undefined;
                const whole = copiesWhole(config, colsRepeated, name);
                for (let k = 0; k < colsRepeated; k++) {
                    dataSets.push({ name: k < whole || name === undefined ? name : repeatPreview(name), values: [], pointLabels: [] });
                }
            }

            // Data rows: a repeated cell fills as many series as remain in its row, and the chart
            // holds at most MAX_CHART_VALUES values.
            let valuesLeft = MAX_CHART_VALUES;
            const valueIndexes: number[][] = dataSets.map(() => []);
            pointCount = 0;
            for (let i = 1; i < rows.length && valuesLeft > 0; i++) {
                const dataCells = getDirectChildren(rows[i], "table:table-cell");
                if (dataCells.length > 0) {
                    const position = pointCount++;
                    const label = getDirectChildren(dataCells[0], "text:p")[0]?.textContent || undefined;
                    if (label) { labels.push(label); labelIndexes.push(position); }

                    let dsIdx = 0;
                    for (let j = 1; j < dataCells.length && dsIdx < dataSets.length; j++) {
                        const colsRepeated = Math.min(repeats(dataCells[j].getAttribute("table:number-columns-repeated")), dataSets.length - dsIdx, valuesLeft);
                        const val = dataCells[j].getAttribute("office:value") || getDirectChildren(dataCells[j], "text:p")[0]?.textContent || "";
                        // (A blank value is a gap at its point: its series has no value there.)
                        if (isBlankOdfValue(val)) { dsIdx += colsRepeated; continue; }
                        const whole = copiesWhole(config, colsRepeated, val);
                        for (let k = 0; k < colsRepeated; k++) {
                            valueIndexes[dsIdx].push(position);
                            dataSets[dsIdx++].values.push(k < whole ? val : repeatPreview(val));
                        }
                        valuesLeft -= colsRepeated;
                    }
                }
            }
            // (Index lists only with entries: a series of no values has none.)
            dataSets.forEach((dataSet, i) => { if (valueIndexes[i].length > 0) dataSet.valueIndexes = valueIndexes[i]; });
        }
    } else {
        // Chart with cell references (common in ODS spreadsheets)
        // Extract series info from chart:series elements
        const seriesNodes = getElementsByTagName(chart, "chart:series").slice(0, MAX_CHART_SERIES);
        for (const series of seriesNodes) {
            // Series label is in chart:label-cell-address attribute
            const labelAddr = series.getAttribute("chart:label-cell-address");
            const valuesAddr = series.getAttribute("chart:values-cell-range-address");

            // Try to get series name from any text:p in a title
            let name: string | undefined = undefined;
            const seriesLabels = getElementsByTagName(series, "text:p");
            if (seriesLabels.length > 0) {
                name = seriesLabels[0].textContent || undefined;
            }

            // If no embedded name, use cell address as identifier
            if (!name && labelAddr) {
                // Extract just the cell reference part, e.g., "Sheet1.$D$2" -> "D2"
                const cellRef = labelAddr.split('.').pop()?.replace(/\$/g, '') || labelAddr;
                name = `Series ${cellRef}`;
            }

            // Values cell range can give us some info
            const dataSet: ChartData['dataSets'][0] = {
                name,
                values: [],
                pointLabels: []
            };

            // Add values address as a hint in rawTexts
            if (valuesAddr) {
                dataSet.values.push(`[${valuesAddr}]`);
            }

            dataSets.push(dataSet);
        }

        // Extract category labels from chart:categories
        const categories = getElementsByTagName(chart, "chart:categories")[0];
        if (categories) {
            const catRange = categories.getAttribute("table:cell-range-address");
            if (catRange) {
                labels.push(`[${catRange}]`);
            }
        }
    }

    // Every axis (#143): its id its `chart:name`, its kind by its dimension (an X axis of a scatter or bubble
    // chart measures values, and LibreOffice marks a date axis); ODF states no side.
    const chartClass = ODF_CHART_TYPES[((chart as Element).getAttribute?.("chart:class") || '').replace(/^chart:/, '')];
    const axes: ChartAxis[] = [];
    const dimensions: string[] = [];
    for (const axis of getElementsByTagName(dom, "chart:axis").slice(0, MAX_CHART_AXES)) {
        const dimension = axis.getAttribute("chart:dimension") || '';
        const kind: ChartAxis['kind'] | undefined = dimension === 'x' ? (chartClass === 'scatter' || chartClass === 'bubble' ? 'value' : axis.getAttribute("chartooo:axis-type") === 'date' ? 'date' : 'category')
            : dimension === 'y' ? 'value' : dimension === 'z' ? 'series' : undefined;
        if (!kind) continue;
        const axisTitleNode = getElementsByTagName(axis, "chart:title")[0];
        const axisTitle = axisTitleNode ? getElementsByTagName(axisTitleNode, "text:p")[0]?.textContent || undefined : undefined;
        axes.push({ id: axis.getAttribute("chart:name") || `${dimension}-${axes.length + 1}`, kind, ...(axisTitle !== undefined && { title: axisTitle }) });
        dimensions.push(dimension);
    }
    // The primary axis of each dimension (named so), else the first: its title is the chart's axis title (the
    // last axis of a dimension won, even one with no title).
    const primary = (dimension: string) => axes.find(a => a.id === `primary-${dimension}`) ?? axes.find((_, i) => dimensions[i] === dimension);
    const xAxis = primary('x'), yAxis = primary('y');
    const xAxisTitle = xAxis?.title;
    const yAxisTitle = yAxis?.title;

    // Each series' chart type (its own `chart:class`, else the chart's) and the axes it is drawn on: the
    // primary X axis and the axis it is attached to (else the primary Y axis). Series are the chart's
    // `chart:series` in order, as its table's columns are.
    const seriesElements = getElementsByTagName(chart, "chart:series");
    dataSets.forEach((dataSet, i) => {
        const element = seriesElements[i];
        const chartType = ODF_CHART_TYPES[(element?.getAttribute("chart:class") || '').replace(/^chart:/, '')] ?? chartClass;
        if (chartType) dataSet.chartType = chartType;
        if (axes.length) {
            const axisIds = [xAxis?.id, element?.getAttribute("chart:attached-axis") || yAxis?.id].filter((id): id is string => !!id);
            if (axisIds.length) dataSet.axisIds = axisIds;
        }
    });

    // From the plot area's style only, as ODF and LibreOffice read it. Microsoft Office writes it on the chart's
    // own style, and LibreOffice ignores it there.
    const plotStyle = getElementsByTagName(chart, "chart:plot-area")[0]?.getAttribute("chart:style-name");
    const plotStyleEl = plotStyle ? getElementsByTagName(dom, "style:style").find(style => style.getAttribute("style:name") === plotStyle) : undefined;
    const treatEmpty = plotStyleEl ? getDirectChildren(plotStyleEl, "style:chart-properties")[0]?.getAttribute("chart:treat-empty-cells") : undefined;
    const displayBlanksAs = ODF_BLANKS[treatEmpty ?? ''];

    const rawTexts = chartRawTexts(title, dataSets, labels);

    return {
        title,
        xAxisTitle,
        yAxisTitle,
        dataSets,
        labels,
        ...(pointCount !== undefined && { ...(labelIndexes.length > 0 && { labelIndexes }), pointCount }),
        ...(displayBlanksAs && { displayBlanksAs }),
        ...(axes.length > 0 && { axes }),
        rawTexts
    };
};

/** Whether a value an ODF chart's table holds is blank: empty, or the NaN LibreOffice writes for an empty cell. */
const isBlankOdfValue = (value: string): boolean => {
    const v = value.trim();
    return v === '' || /^[-+]?(?:1\.#)?nan$/i.test(v);
};

/**
 * Universal chart data extractor that selects logic based on XML content.
 * @param xmlBuffer Chart XML buffer
 */
export const extractChartData = (xmlBuffer: Buffer, config?: OfficeParserConfig): ChartData => {
    // An ODF part declares the OpenDocument namespaces on its root, which LibreOffice writes with some
    // forty declarations: looked for in the first 500 bytes, its chart namespace (past byte 750) was
    // missed and every LibreOffice chart was read as OpenXML, empty.
    const head = xmlBuffer.toString("utf8", 0, 64 * 1024);
    if (head.includes("urn:oasis:names:tc:opendocument:xmlns:")) {
        return extractOdfChartData(xmlBuffer, config);
    } else {
        return extractOpenXmlChartData(xmlBuffer, config);
    }
};

/**
 * Each attachment found by a name exactly as `attachments.find(a => a?.name === name)` finds it: the first of the name,
 * and for a chart node with no name, or an empty one, the first attachment with none, or an empty one (as XLSX and ODS
 * output matches a chart node). From a table built in one pass, so a workbook of many charts takes time
 * growing with their number, not its square.
 */
export const attachmentsByExactName = (attachments: readonly (OfficeAttachment | null | undefined)[] | undefined): ((name: unknown) => OfficeAttachment | null | undefined) => {
    const byName = new Map<unknown, OfficeAttachment | null | undefined>();
    // (NaN is no attachment's name to `===`: left out, so a lookup by it finds none, as find does.)
    for (const attachment of attachments ?? []) {
        const name = attachment?.name;
        if (name === name && !byName.has(name)) byName.set(name, attachment);
    }
    return name => byName.get(name);
};

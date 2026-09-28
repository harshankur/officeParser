import { ChartData } from "../types";
import { parseXmlString, getElementsByTagName, getDirectChildren } from "./xmlUtils";

/**
 * Extracts a single text element located at:
 * c:title -> c:tx -> c:rich -> a:p -> a:r -> a:t
 * OR c:tx -> c:strRef -> c:strCache -> c:pt -> c:v
 * @param el Chart element
 * @param tagName The tag to search for (e.g., "c:title", "c:tx")
 */
const extractOpenXmlRichText = (el: Element, tagName: string): string | undefined => {
    const target = (el.localName === tagName || el.tagName === tagName) ? el : el.getElementsByTagName(tagName)[0];
    if (!target) return undefined;

    // 1. Try c:rich or a:p (standard rich text)
    const richNodes = target.getElementsByTagName("c:rich");
    const pNodes = target.getElementsByTagName("a:p");
    const textContainers = richNodes.length > 0 ? Array.from(richNodes) : Array.from(pNodes);

    if (textContainers.length > 0) {
        let acc = "";
        for (const container of textContainers) {
            const tNodes = container.getElementsByTagName("a:t");
            for (let i = 0; i < tNodes.length; i++) {
                acc += (tNodes[i].textContent || "") + " ";
            }
        }
        if (acc.trim()) return acc.trim();
    }

    // 2. Try c:v (cached values/strings)
    const vNode = target.getElementsByTagName("c:v")[0];
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
 * A chart's text: its title (when given), its category labels once, then each series' name and values.
 * The labels were written again for every series, so a small chart of many series and many labels
 * made text of their product (gigabytes from a few KB).
 */
export function chartRawTexts(title: string | undefined, dataSets: ChartData['dataSets'], labels: string[]): string[] {
    const raw: string[] = [];
    if (title) raw.push(title);
    for (const label of labels) raw.push(label);
    for (const ds of dataSets) {
        if (ds.name) raw.push(ds.name);
        for (const value of ds.values) raw.push(value);
    }
    return raw;
}

/** A repeat count as a document states it: at least 1, NaN as 1. */
const repeats = (value: string | null): number => {
    const n = parseInt(value || '1', 10);
    return Number.isFinite(n) && n > 1 ? n : 1;
};

/**
 * Extracts fully structured chart data from OpenXML (PPTX, XLSX) chart XML.
 * @param xmlBuffer Chart XML buffer
 */
const extractOpenXmlChartData = (xmlBuffer: Buffer): ChartData => {
    const xml = xmlBuffer.toString("utf8");
    const dom = parseXmlString(xml);
    const root = dom.documentElement as Element;
    if (!root) return { title: undefined, xAxisTitle: undefined, yAxisTitle: undefined, dataSets: [], labels: [], rawTexts: [] };

    const title = extractOpenXmlRichText(root, "c:title");

    // Axis titles
    let xAxisTitle: string | undefined = undefined;
    let yAxisTitle: string | undefined = undefined;

    const catAxes = root.getElementsByTagName("c:catAx");
    if (catAxes.length > 0) xAxisTitle = extractOpenXmlRichText(catAxes[0] as Element, "c:title");

    const valAxes = root.getElementsByTagName("c:valAx");
    if (valAxes.length > 0) yAxisTitle = extractOpenXmlRichText(valAxes[0] as Element, "c:title");

    // Extract Series (dataSets)
    const seriesNodes = root.getElementsByTagName("c:ser");
    const dataSets: ChartData['dataSets'] = [];
    const sharedLabels: string[] = [];

    for (let i = 0; i < seriesNodes.length; i++) {
        const ser = seriesNodes[i] as Element;

        // DataSet Name (c:tx)
        const name = extractOpenXmlRichText(ser, "c:tx");

        // Values (c:val)
        const values: string[] = [];
        const valNode = ser.getElementsByTagName("c:val")[0] || ser.getElementsByTagName("c:yVal")[0];
        if (valNode) {
            const vNodes = valNode.getElementsByTagName("c:v");
            for (let j = 0; j < vNodes.length; j++) {
                const v = vNodes[j].textContent?.trim();
                if (v) values.push(v);
            }
        }

        // Point Labels (data labels)
        const pointLabels: string[] = [];
        const dLbls = ser.getElementsByTagName("c:dLbl");
        for (let j = 0; j < dLbls.length; j++) {
            const lbl = extractOpenXmlRichText(dLbls[j], "c:tx");
            if (lbl) pointLabels.push(lbl);
        }

        // Categories (labels) - c:cat or c:xVal
        const catNode = ser.getElementsByTagName("c:cat")[0] || ser.getElementsByTagName("c:xVal")[0];
        if (catNode) {
            const vNodes = catNode.getElementsByTagName("c:v");
            const localLabels: string[] = [];
            for (let j = 0; j < vNodes.length; j++) {
                const v = vNodes[j].textContent?.trim();
                if (v) localLabels.push(v);
            }
            if (localLabels.length > 0 && sharedLabels.length === 0) {
                for (const label of localLabels) sharedLabels.push(label);
            }
        }

        dataSets.push({ name, values, pointLabels });
    }

    const rawTexts = chartRawTexts(undefined, dataSets, sharedLabels);

    return {
        title,
        xAxisTitle,
        yAxisTitle,
        dataSets,
        labels: sharedLabels,
        rawTexts
    };
};

/**
 * Extracts structured chart data from ODF (ODP, ODS) chart content.xml.
 * @param xmlBuffer Chart XML buffer
 */
const extractOdfChartData = (xmlBuffer: Buffer): ChartData => {
    const xml = xmlBuffer.toString("utf8");
    const dom = parseXmlString(xml);

    const chart = getElementsByTagName(dom, "chart:chart")[0] || dom.documentElement;
    const titleNode = getElementsByTagName(chart, "chart:title")[0];
    const title = titleNode ? getElementsByTagName(titleNode, "text:p")[0]?.textContent || undefined : undefined;

    const table = getElementsByTagName(chart, "table:table")[0];
    const dataSets: ChartData['dataSets'] = [];
    const labels: string[] = [];

    if (table) {
        // Chart with embedded data table (common in ODP presentations)
        let rows: Element[] = [];
        const headerRowsNode = getDirectChildren(table, "table:table-header-rows")[0];
        if (headerRowsNode) {
            rows.push(...getDirectChildren(headerRowsNode, "table:table-row"));
        }
        rows.push(...getDirectChildren(table, "table:table-row"));

        if (rows.length > 0) {
            // Header row for series names
            const headerCells = getDirectChildren(rows[0], "table:table-cell");
            for (let j = 1; j < headerCells.length && dataSets.length < MAX_CHART_SERIES; j++) {
                // A repeated cell is as many series as it repeats, up to the chart's limit.
                const colsRepeated = Math.min(repeats(headerCells[j].getAttribute("table:number-columns-repeated")), MAX_CHART_SERIES - dataSets.length);
                const name = getDirectChildren(headerCells[j], "text:p")[0]?.textContent || undefined;
                for (let k = 0; k < colsRepeated; k++) {
                    dataSets.push({ name, values: [], pointLabels: [] });
                }
            }

            // Data rows: a repeated cell fills as many series as remain in its row, and the chart
            // holds at most MAX_CHART_VALUES values.
            let valuesLeft = MAX_CHART_VALUES;
            for (let i = 1; i < rows.length && valuesLeft > 0; i++) {
                const dataCells = getDirectChildren(rows[i], "table:table-cell");
                if (dataCells.length > 0) {
                    const label = getDirectChildren(dataCells[0], "text:p")[0]?.textContent || undefined;
                    if (label) labels.push(label);

                    let dsIdx = 0;
                    for (let j = 1; j < dataCells.length && dsIdx < dataSets.length; j++) {
                        const colsRepeated = Math.min(repeats(dataCells[j].getAttribute("table:number-columns-repeated")), dataSets.length - dsIdx, valuesLeft);
                        const val = dataCells[j].getAttribute("office:value") || getDirectChildren(dataCells[j], "text:p")[0]?.textContent || "";
                        for (let k = 0; k < colsRepeated; k++) dataSets[dsIdx++].values.push(val);
                        valuesLeft -= colsRepeated;
                    }
                }
            }
        }
    } else {
        // Chart with cell references (common in ODS spreadsheets)
        // Extract series info from chart:series elements
        const seriesNodes = getElementsByTagName(chart, "chart:series");
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

    // Axis titles
    const axes = getElementsByTagName(dom, "chart:axis");
    let xAxisTitle: string | undefined = undefined;
    let yAxisTitle: string | undefined = undefined;
    for (const axis of axes) {
        const dimension = axis.getAttribute("chart:dimension");
        const axisTitleNode = getElementsByTagName(axis, "chart:title")[0];
        const axisTitle = axisTitleNode ? getElementsByTagName(axisTitleNode, "text:p")[0]?.textContent || undefined : undefined;
        if (dimension === 'x') xAxisTitle = axisTitle;
        else if (dimension === 'y') yAxisTitle = axisTitle;
    }

    const rawTexts = chartRawTexts(title, dataSets, labels);

    return {
        title,
        xAxisTitle,
        yAxisTitle,
        dataSets,
        labels,
        rawTexts
    };
};

/**
 * Universal chart data extractor that selects logic based on XML content.
 * @param xmlBuffer Chart XML buffer
 */
export const extractChartData = (xmlBuffer: Buffer): ChartData => {
    const head = xmlBuffer.toString("utf8", 0, 500);
    if (head.includes("urn:oasis:names:tc:opendocument:xmlns:chart:1.0")) {
        return extractOdfChartData(xmlBuffer);
    } else {
        return extractOpenXmlChartData(xmlBuffer);
    }
};
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
 * Extracts fully structured chart data from OpenXML (PPTX, XLSX) chart XML.
 * @param xmlBuffer Chart XML buffer
 */
const extractOpenXmlChartData = (xmlBuffer: Buffer): ChartData => {
    const xml = xmlBuffer.toString("utf8");
    const dom = parseXmlString(xml);
    const root = dom.documentElement;

    const title = extractOpenXmlRichText(root, "c:title");

    // Extract chart type from plotArea
    let chartType: string | undefined = undefined;
    const plotArea = root.getElementsByTagName("c:plotArea")[0];
    if (plotArea) {
        // Check for common chart types in plotArea
        const chartTypeElements = [
            'c:barChart', 'c:lineChart', 'c:pieChart', 'c:columnChart',
            'c:areaChart', 'c:scatterChart', 'c:bubbleChart', 'c:doughnutChart',
            'c:radarChart', 'c:surfaceChart', 'c:ofPieChart', 'c:stockChart'
        ];
        for (const tagName of chartTypeElements) {
            if (plotArea.getElementsByTagName(tagName).length > 0) {
                // Extract base name (e.g., 'barChart' -> 'bar')
                chartType = tagName.replace('c:', '').replace('Chart', '').toLowerCase();
                break;
            }
        }
    }

    // Axis titles
    let xAxisTitle: string | undefined = undefined;
    let yAxisTitle: string | undefined = undefined;

    const catAxes = root.getElementsByTagName("c:catAx");
    if (catAxes.length > 0) xAxisTitle = extractOpenXmlRichText(catAxes[0], "c:title");

    const valAxes = root.getElementsByTagName("c:valAx");
    if (valAxes.length > 0) yAxisTitle = extractOpenXmlRichText(valAxes[0], "c:title");

    // Extract Series (dataSets)
    const seriesNodes = root.getElementsByTagName("c:ser");
    const dataSets: ChartData['dataSets'] = [];
    const sharedLabels: string[] = [];

    for (let i = 0; i < seriesNodes.length; i++) {
        const ser = seriesNodes[i];

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
                sharedLabels.push(...localLabels);
            }
        }

        dataSets.push({ name, values, pointLabels });
    }

    // Structured rawTexts: for each dataset: Name -> Labels -> Values
    const rawTexts: string[] = [];
    for (const ds of dataSets) {
        if (ds.name) rawTexts.push(ds.name);
        rawTexts.push(...sharedLabels);
        rawTexts.push(...ds.values);
    }

    return {
        title,
        chartType,
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
    const rawTexts: string[] = [];

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
            for (let j = 1; j < headerCells.length; j++) {
                const colsRepeated = parseInt(headerCells[j].getAttribute("table:number-columns-repeated") || "1");
                const name = getDirectChildren(headerCells[j], "text:p")[0]?.textContent || undefined;
                for (let k = 0; k < colsRepeated; k++) {
                    dataSets.push({ name, values: [], pointLabels: [] });
                }
            }

            // Data rows
            for (let i = 1; i < rows.length; i++) {
                const dataCells = getDirectChildren(rows[i], "table:table-cell");
                if (dataCells.length > 0) {
                    const label = getDirectChildren(dataCells[0], "text:p")[0]?.textContent || undefined;
                    if (label) labels.push(label);

                    let dsIdx = 0;
                    for (let j = 1; j < dataCells.length; j++) {
                        const colsRepeated = parseInt(dataCells[j].getAttribute("table:number-columns-repeated") || "1");
                        const val = dataCells[j].getAttribute("office:value") || getDirectChildren(dataCells[j], "text:p")[0]?.textContent || "";
                        for (let k = 0; k < colsRepeated; k++) {
                            if (dataSets[dsIdx]) {
                                dataSets[dsIdx].values.push(val);
                            }
                            dsIdx++;
                        }
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

    // Extract chart type from chart:class attribute or chart type elements
    let chartType: string | undefined = undefined;
    const chartClass = chart.getAttribute("chart:class");
    if (chartClass) {
        // chart:class values like "chart:bar", "chart:line", "chart:pie", etc.
        chartType = chartClass.replace('chart:', '').toLowerCase();
    }

    // Structured rawTexts: title + series info
    if (title) rawTexts.push(title);
    for (const ds of dataSets) {
        if (ds.name) rawTexts.push(ds.name);
        rawTexts.push(...labels);
        rawTexts.push(...ds.values);
    }

    return {
        title,
        chartType,
        xAxisTitle,
        yAxisTitle,
        dataSets,
        labels,
        rawTexts
    };
};

/**
 * Extracts chart data from chartex (cx:) namespace charts with hierarchical categories.
 * @param xmlBuffer Chart XML buffer
 */
const extractChartexChartData = (xmlBuffer: Buffer): ChartData => {
    const xml = xmlBuffer.toString("utf8");
    const dom = parseXmlString(xml);

    // Find plotArea
    const plotArea = getElementsByTagName(dom, "cx:plotArea")[0];
    if (!plotArea) {
        // Fallback: try to extract basic info
        return {
            title: undefined,
            xAxisTitle: undefined,
            yAxisTitle: undefined,
            dataSets: [],
            labels: [],
            rawTexts: []
        };
    }

    const plotAreaRegion = getElementsByTagName(plotArea, "cx:plotAreaRegion")[0];
    if (!plotAreaRegion) {
        return {
            title: undefined,
            xAxisTitle: undefined,
            yAxisTitle: undefined,
            dataSets: [],
            labels: [],
            rawTexts: []
        };
    }

    // Extract chart type from first series layoutId
    const seriesNodes = getElementsByTagName(plotAreaRegion, "cx:series");
    let chartType: string | undefined = undefined;
    if (seriesNodes.length === 0) {
        return {
            title: undefined,
            chartType: undefined,
            xAxisTitle: undefined,
            yAxisTitle: undefined,
            dataSets: [],
            labels: [],
            rawTexts: []
        };
    }
    
    // Extract chart type from first series layoutId
    if (seriesNodes.length > 0) {
        const firstSeries = seriesNodes[0];
        const layoutId = firstSeries.getAttribute("layoutId");
        if (layoutId) {
            // layoutId values map to chart types (e.g., "100" = column, "101" = bar, etc.)
            // Common mappings: 100=column, 101=bar, 102=line, 103=pie, 104=area, etc.
            const layoutIdMap: Record<string, string> = {
                '100': 'column', '101': 'bar', '102': 'line', '103': 'pie',
                '104': 'area', '105': 'scatter', '106': 'bubble', '107': 'doughnut',
                '108': 'radar', '109': 'surface'
            };
            chartType = layoutIdMap[layoutId] || undefined;
        }
    }

    // Extract chartData
    const chartData = getElementsByTagName(dom, "cx:chartData")[0];
    if (!chartData) {
        return {
            title: undefined,
            chartType: undefined,
            xAxisTitle: undefined,
            yAxisTitle: undefined,
            dataSets: [],
            labels: [],
            rawTexts: []
        };
    }

    // Build a map of data by id
    const dataMap: { [key: string]: Element } = {};
    const dataElements = getElementsByTagName(chartData, "cx:data");
    for (const dataEl of dataElements) {
        const dataId = dataEl.getAttribute("id") || '0';
        dataMap[dataId] = dataEl;
    }

    const dataSets: ChartData['dataSets'] = [];
    const labels: string[] = [];
    const rawTexts: string[] = [];

    // Process each series
    for (const serNode of seriesNodes) {
        const dataIdNode = getElementsByTagName(serNode, "cx:dataId")[0];
        const dataId = dataIdNode ? (dataIdNode.getAttribute("val") || '0') : '0';

        const dataEl = dataMap[dataId];
        if (!dataEl) {
            continue;
        }

        // Parse categories
        const categories: string[] = [];

        // Check for string categories (cx:strDim type="cat")
        const strDim = getElementsByTagName(dataEl, "cx:strDim")[0];
        if (strDim && strDim.getAttribute("type") === 'cat') {
            const lvlNodes = getElementsByTagName(strDim, "cx:lvl");
            const levelCount = lvlNodes.length;

            if (levelCount > 0) {
                // Get ptCount from first level
                const firstLvl = lvlNodes[0];
                const ptCount = parseInt(firstLvl.getAttribute("ptCount") || '0', 10);

                // Build hierarchical structure
                // In XML, first level is innermost (e.g., Leaf), last level is outermost (e.g., Branch)
                // We want to extract the innermost value for each category
                for (let idx = 0; idx < ptCount; idx++) {
                    const levelValues: string[] = [];

                    // Extract from each level (innermost to outermost)
                    for (const lvl of lvlNodes) {
                        const pts = getElementsByTagName(lvl, "cx:pt");
                        const pt = Array.from(pts).find(
                            p => parseInt(p.getAttribute("idx") || '-1', 10) === idx
                        );

                        if (pt) {
                            const ptText = pt.textContent || '';
                            levelValues.push(ptText);
                        }
                    }

                    if (levelValues.length > 0) {
                        // The first element is the innermost (value)
                        const value = levelValues[0];
                        categories.push(value);
                    }
                }
            }
        }

        // Check for numeric categories (cx:numDim type="catVal")
        if (categories.length === 0) {
            const numDimCat = Array.from(getElementsByTagName(dataEl, "cx:numDim")).find(
                dim => dim.getAttribute("type") === 'catVal'
            );
            if (numDimCat) {
                const lvl = getElementsByTagName(numDimCat, "cx:lvl")[0];
                if (lvl) {
                    const pts = getElementsByTagName(lvl, "cx:pt");
                    for (const pt of pts) {
                        const ptText = pt.textContent || '';
                        if (ptText) categories.push(ptText);
                    }
                }
            }
        }

        // Parse values (cx:numDim type="size" or type="val")
        const values: string[] = [];
        const numDimVal = Array.from(getElementsByTagName(dataEl, "cx:numDim")).find(
            dim => dim.getAttribute("type") === 'size' || dim.getAttribute("type") === 'val'
        );
        if (numDimVal) {
            const lvl = getElementsByTagName(numDimVal, "cx:lvl")[0];
            if (lvl) {
                const pts = getElementsByTagName(lvl, "cx:pt");
                for (const pt of pts) {
                    const ptText = pt.textContent || '';
                    if (ptText) values.push(ptText);
                }
            }
        }

        // Extract series name if available
        let name: string | undefined = undefined;
        const txNode = getElementsByTagName(serNode, "c:tx")[0];
        if (txNode) {
            name = extractOpenXmlRichText(txNode, "c:tx");
        }

        if (categories.length > 0 || values.length > 0) {
            dataSets.push({ name, values, pointLabels: [] });
            if (categories.length > 0 && labels.length === 0) {
                labels.push(...categories);
            }
        }
    }

    // Extract title
    const title = extractOpenXmlRichText(dom.documentElement, "c:title");

    // Extract axis titles
    let xAxisTitle: string | undefined = undefined;
    let yAxisTitle: string | undefined = undefined;
    const catAxes = getElementsByTagName(dom.documentElement, "c:catAx");
    if (catAxes.length > 0) xAxisTitle = extractOpenXmlRichText(catAxes[0], "c:title");
    const valAxes = getElementsByTagName(dom.documentElement, "c:valAx");
    if (valAxes.length > 0) yAxisTitle = extractOpenXmlRichText(valAxes[0], "c:title");

    // Build rawTexts
    if (title) rawTexts.push(title);
    for (const ds of dataSets) {
        if (ds.name) rawTexts.push(ds.name);
        rawTexts.push(...labels);
        rawTexts.push(...ds.values);
    }

    return {
        title,
        chartType,
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
    
    // Check for chartex/chartx namespace (cx:)
    if (head.includes("cx:") || head.includes("cx:plotArea") || head.includes("cx:chartData")) {
        return extractChartexChartData(xmlBuffer);
    }
    
    // Check for ODF namespace
    if (head.includes("urn:oasis:names:tc:opendocument:xmlns:chart:1.0")) {
        return extractOdfChartData(xmlBuffer);
    }
    
    // Default to OpenXML
    return extractOpenXmlChartData(xmlBuffer);
};
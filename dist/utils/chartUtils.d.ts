/// <reference types="node" />
import { ChartData } from "../types";
/**
 * Universal chart data extractor that selects logic based on XML content.
 * @param xmlBuffer Chart XML buffer
 */
export declare const extractChartData: (xmlBuffer: Buffer) => ChartData;

#!/usr/bin/env node

export type TextBlock = {
    type: 'text';
    content: string;
}

export type ImageBlock = {
    type: 'image';
    buffer: Buffer;
    mimeType: string;
    filename?: string;
}

export type HierarchicalCategory = {
    levels: string[];
    value: string;
}

export type TableBlock = {
    type: 'table';
    name: string;
    rows: Array<{ cols: Array<{ value: string }> }>;
}

export type ChartBlock = {
    type: 'chart';
    chartType: string;
    series: Array<{ categories: Array<string | HierarchicalCategory>, values: Array<number> }>;
}

export type Block = TextBlock | ImageBlock | TableBlock | ChartBlock;

type DocImage = { buffer: Buffer; type: string; filename?: string }

export type CoordinateData = {
    x: number;
    y: number;
    width: number;
    height: number;
    rotation?: number;
    zIndex?: number;
}

export type PowerPointTextElement = {
    type: 'text';
    content: string;
    coordinates: CoordinateData;
    slideNumber?: string;
}

export type PowerPointImageElement = {
    type: 'image';
    buffer: Buffer;
    mimeType: string;
    filename?: string;
    coordinates: CoordinateData;
    slideNumber?: string;
}

export type PowerPointShapeElement = {
    type: 'shape';
    text?: string;
    shapeType: string;
    coordinates: CoordinateData;
    slideNumber?: string;
}

export type PowerPointElement = PowerPointTextElement | PowerPointImageElement | PowerPointShapeElement;

export type Table = {
    name: string;
    rows: Array<{ cols: Array<{ value: string }> }>;
}

export type Chart = {
    chartType: string;
    series: Array<{ categories: Array<string | HierarchicalCategory>, values: Array<number> }>;
}

export type Image = {
    buffer: Buffer;
    mimeType: string;
    filename?: string;
}

export type OfficeParserConfig = {
    /**
     * Flag to show all the logs to console in case of an error irrespective of your own handling. Default is false.
     */
    outputErrorToConsole?: boolean;
    /**
     * The delimiter used for every new line in places that allow multiline text like word. Default is \n.
     */
    newlineDelimiter?: string;
    /**
     * Flag to ignore notes from parsing in files like powerpoint. Default is false. It includes notes in the parsed text by default.
     */
    ignoreNotes?: boolean;
    /**
     * Flag, if set to true, will collectively put all the parsed text from notes at last in files like powerpoint. Default is false. It puts each notes right after its main slide content. If ignoreNotes is set to true, this flag is also ignored.
     */
    putNotesAtLast?: boolean;
    /**
     * Flag to extract images from files. Default is false. If set to true, the blocks array will contain image blocks alongside text blocks.
     */
    extractImages?: boolean;
    /**
     * Flag to extract charts from files. Default is false. If set to true, the return object will contain a 'charts' array.
     */
    extractCharts?: boolean;
}

export type ParseOfficeResult = {
    text: string;
    blocks: Block[];
    elements?: PowerPointElement[];
    slides?: { [key: string]: PowerPointElement[] };
    tables?: Table[];
    charts?: Chart[];
    images?: Image[];
}

/** Main async function with callback to execute parseOffice for supported files
 * @param {string | Buffer | ArrayBuffer} srcFile      File path or file buffers or Javascript ArrayBuffer
 * @param {(data: ParseOfficeResult | undefined, error: Error | undefined) => void} callback     Callback function that returns value or error
 * @param {OfficeParserConfig}            [config={}]  [OPTIONAL]: Config Object for officeParser
 * @returns {void}
 */
export function parseOffice(srcFile: string | Buffer | ArrayBuffer, callback: (data: ParseOfficeResult | undefined, error: Error | undefined) => void, config?: OfficeParserConfig): void;

/** Main async function that can be used with await to execute parseOffice. Or it can be used with promises.
 * @param {string | Buffer | ArrayBuffer} srcFile     File path or file buffers or Javascript ArrayBuffer
 * @param {OfficeParserConfig}            [config={}] [OPTIONAL]: Config Object for officeParser
 * @returns {Promise<ParseOfficeResult>}
 */
export function parseOfficeAsync(srcFile: string | Buffer | ArrayBuffer, config?: OfficeParserConfig): Promise<ParseOfficeResult>;

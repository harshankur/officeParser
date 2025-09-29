#!/usr/bin/env node
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
};
/** Main async function with callback to execute parseOffice for supported files
 * @param {string | Buffer | ArrayBuffer} srcFile      File path or file buffers or Javascript ArrayBuffer
 * @param {function}                      callback     Callback function that returns value or error
 * @param {OfficeParserConfig}            [config={}]  [OPTIONAL]: Config Object for officeParser
 * @returns {void}
 */
export function parseOffice(srcFile: string | Buffer | ArrayBuffer, callback: Function, config?: OfficeParserConfig): void;
/** Main async function that can be used with await to execute parseOffice. Or it can be used with promises.
 * @param {string | Buffer | ArrayBuffer} srcFile     File path or file buffers or Javascript ArrayBuffer
 * @param {OfficeParserConfig}            [config={}] [OPTIONAL]: Config Object for officeParser
 * @returns {Promise<string>}
 */
export function parseOfficeAsync(srcFile: string | Buffer | ArrayBuffer, config?: OfficeParserConfig): Promise<string>;
//# sourceMappingURL=officeParser.d.ts.map
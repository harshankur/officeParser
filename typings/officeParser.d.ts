#!/usr/bin/env node
/** Main function for parsing text from word files
 * @param {string} filename File path
 * @param {function} callback Callback function that returns value or error
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {void}
 */
export function parseWord(filename: string, callback: Function, deleteOfficeDist?: boolean): void;
/** Main function for parsing text from PowerPoint files
 * @param {string} filename File path
 * @param {function} callback Callback function that returns value or error
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {void}
 */
export function parsePowerPoint(filename: string, callback: Function, deleteOfficeDist?: boolean): void;
/** Main function for parsing text from Excel files
 * @param {string} filename File path
 * @param {function} callback Callback function that returns value or error
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {void}
 */
export function parseExcel(filename: string, callback: Function, deleteOfficeDist?: boolean): void;
/** Main function for parsing text from open office files
 * @param {string} filename File path
 * @param {function} callback Callback function that returns value or error
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {void}
 */
export function parseOpenOffice(filename: string, callback: Function, deleteOfficeDist?: boolean): void;
/** Main async function with callback to execute parseOffice for supported files
 * @param {string} filename File path
 * @param {function} callback Callback function that returns value or error
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {void}
 */
export function parseOffice(filename: string, callback: Function, deleteOfficeDist?: boolean): void;
/** Async function that can be used with await to execute parseWord. Or it can be used with promises.
 * @param {string} filename File path
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {Promise<string>}
 */
export function parseWordAsync(filename: string, deleteOfficeDist?: boolean): Promise<string>;
/** Async function that can be used with await to execute parsePowerPoint. Or it can be used with promises.
 * @param {string} filename File path
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {Promise<string>}
 */
export function parsePowerPointAsync(filename: string, deleteOfficeDist?: boolean): Promise<string>;
/** Async function that can be used with await to execute parseExcel. Or it can be used with promises.
 * @param {string} filename File path
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {Promise<string>}
 */
export function parseExcelAsync(filename: string, deleteOfficeDist?: boolean): Promise<string>;
/** Async function that can be used with await to execute parseOpenOffice. Or it can be used with promises.
 * @param {string} filename File path
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {Promise<string>}
 */
export function parseOpenOfficeAsync(filename: string, deleteOfficeDist?: boolean): Promise<string>;
/**
 * Main async function that can be used with await to execute parseOffice. Or it can be used with promises.
 * @param {string} filename File path
 * @param {boolean} [deleteOfficeDist=true] Optional: Delete the officeDist directory created while unarchiving the doc file to get its content underneath. By default, we delete those files after we are done reading them.
 * @returns {Promise<string>}
 */
export function parseOfficeAsync(filename: string, deleteOfficeDist?: boolean): Promise<string>;
/**
 * Set decompression directory. The final decompressed data will be put inside officeDist folder within your directory
 * @param {string} newLocation Relative path to the directory that will contain officeDist folder with decompressed data
 * @returns {void}
 */
export function setDecompressionLocation(newLocation: string): void;
/** Enable console output
 * @returns {void}
 */
export function enableConsoleOutput(): void;
/** Disabled console output
 * @returns {void}
 */
export function disableConsoleOutput(): void;
//# sourceMappingURL=officeParser.d.ts.map
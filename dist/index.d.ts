/**
 * Generates an XLSX file by replacing data in a template file.
 * @throws {Error} Throws an error if either the templateURL or data is not provided, or if unable to fetch the template file.
 */
export declare function generateXlsx(templateURL: string, data: Object): Promise<void | Blob>;
/**
 * Downloads an XLSX file by replacing data in a template file and initiates the download in the browser.
 */
export declare function downloadXlsx(templateURL: string, data: Object, fileName: string): Promise<void>;

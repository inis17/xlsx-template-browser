
declare module 'xlsx-template-browser' {
  /**
 * Generates an XLSX file by replacing data in a template file.
 * @param {string} templateURL - The URL of the template XLSX file.
 * @param {Object} data - The data object containing values to replace in the template.
 * @returns {Promise<Blob>} A promise that resolves with the Blob of the generated XLSX file.
 * @throws {Error} Throws an error if either the templateURL or data is not provided, or if unable to fetch the template file.
 */
  export async function generateXlsx(templateURL: string, data: object): Promise<Blob>
  /**
 * Downloads an XLSX file by replacing data in a template file and initiates the download in the browser.
 * @param {string} templateURL - The URL of the template XLSX file.
 * @param {Object} data - The data object containing values to replace in the template.
 * @param {string} [fileName] - The name of the file to be downloaded. If not provided, a default name based on the current date will be used.
 * @returns {Promise<void>} A promise that resolves once the download is initiated.
 */
  export async function downloadXlsx(templateURL: string, data: object, fileName: string): Promise<void>

}



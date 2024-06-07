import { downloadBlob } from './helpers'
import replaceInTemplate from './main'

/**
 * Generates an XLSX file by replacing data in a template file.
 * @throws {Error} Throws an error if either the templateURL or data is not provided, or if unable to fetch the template file.
 */
export async function generateXlsx(templateURL: string, data: Object): Promise<void | Blob> {
  if (!templateURL || !data) throw new Error('No templateURL or data provided')
  return await fetch(templateURL)
    .then(response => {
      if (response.ok) return response.arrayBuffer()
      throw new Error(`Unable to fetch the template at: ${templateURL}`)
    })
    .then(template => replaceInTemplate(template, data))
    .catch(e => console.error(e))
}

/**
 * Downloads an XLSX file by replacing data in a template file and initiates the download in the browser.
 */
export async function downloadXlsx(templateURL: string, data: Object, fileName: string): Promise<void> {
  fileName = fileName || `${new Date().toISOString().substring(0, 10)} - Report.xlsx`
  const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  const xlsxBlob = await generateXlsx(templateURL, data)
  if (xlsxBlob) downloadBlob(xlsxBlob, fileName, mimeType)
}

export function valueToString(value: string | string[] | object): string {
  if (!value) return ''
  if (Array.isArray(value)) return value.map(valueToString).toString()
  if (typeof value === 'object') value = JSON.stringify(value)
  return value.toString()
}

function dateToExcel(value: Date): number {
  let date = new Date(value)
  return 25569 + ((date.getTime() - (date.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24))
}

/**
 * Guesses the Excel data type for the given JavaScript primitive value.
 */
export function guessDataType(value: Object): { cellType: string, value: string | number } {
  if (typeof value === 'undefined') return { cellType: 's', value: '' }
  if (typeof value === 'number') {
    if (isFinite(value)) return { cellType: 'n', value }
    return undefined
  }
  if (value instanceof Date) {
    if (isNaN(value.getDate())) return undefined
    return { cellType: 'n', value: dateToExcel(value) }
  }
  if (typeof value === 'boolean') return { cellType: 'b', value: value ? 1 : 0 }
  return { cellType: 's', value: valueToString(value) }
}

/**
 * Creates a new shared string for an Excel file.
 */
export function createSharedString(text: string, template?: Element): Element {
  const si = template ? template.cloneNode() as Element : document.createElement('si')
  let t = document.createElement('t')
  t.textContent = text
  si.appendChild(t)
  return si
}

type Cell = {
  value: Element | null | undefined
  row: number
  column: number
  template: Element
  cellType: string | null
}
/**
 * Creates a cell with the specified value.
 */
export function createCell(c: Cell): Element {
  // If the cell is not a string, then skip
  if (!c.cellType && !c.value) {
    return c.template
  }
  // If the cell is formula, then skip
  if (c.cellType === 'f') {
    // Check precalculated value
    const v = c.value.querySelector('v')
    if (v) v.remove()
    return c.value
  }
  // Clone initial node with the style or create a new one
  const cell = c.template ? c.template.cloneNode() as Element : document.createElement('c')
  // Create a cell reference
  if (c.row && c.column) {
    const column = typeof c.column === 'number' ? getExcelColumnName(c.column) : c.column
    cell.setAttribute('r', column + c.row)
  }
  // If no value is provided, return cell
  if (typeof c.value === 'undefined') return cell
  const valueTag = document.createElement('v')
  cell.setAttribute('t', c.cellType)
  valueTag.textContent = c.value.textContent
  cell.appendChild(valueTag)
  return cell
}

/**
 * Converts a numeric index to an Excel column name.
 */
function getExcelColumnName(index: number): string {
  let result = ''
  while (index > 0) {
    const remainder = (index - 1) % 26
    result = String.fromCharCode(65 + remainder) + result
    index = Math.floor((index - 1) / 26)
  }
  return result
}

/**
 * Converts an Excel column name to its numeric index.
 */
export function getExcelColumnIndex(columnName: string): number {
  columnName = columnName.replace(/\d+/g, '')
  let index = 0
  for (let i = 0; i < columnName.length; i++) {
    const charCode = columnName.charCodeAt(i) - 64
    index = index * 26 + charCode
  }
  return index
}

// /**
//  * Escapes unsafe XML characters in a given string to ensure its safe inclusion
//  * in XML documents, preventing parsing issues.
//  */
// export function escapeXml(unsafe: string): string {
//   const unsafeChars = {
//     '<': '&lt;',
//     '>': '&gt;',
//     '&': '&amp;',
//     '\'': '&apos;',
//     '"': '&quot;'
//   }
//   return unsafe.replace(/[<>&'"]/g, (c) => { return unsafeChars[c] })
// }

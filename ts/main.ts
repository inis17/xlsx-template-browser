// Refer to Excel OpenXML format:
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cell?view=openxml-2.8.1
// Import JSZip library for working with ZIP files
import JSZip from 'jszip'
import {
  createCell,
  createSharedString,
  getExcelColumnIndex,
  guessDataType,
  valueToString
} from './excelHelpers'

import { getByNotation } from './helpers'

async function replaceInTemplate(template: ArrayBuffer, data: object) {
  if (!data) throw new Error(`xlsx-template-browser - replaceInTemplate - The data passed is empty`)
  if (!template) throw new Error(`xlsx-template-browser - replaceInTemplate - The template passed is empty`)

  // Regular expression for parsing a property accessor enclosed in ${}
  const accessorRegex = /(?<=^\$\{)(?<table>table:)?(?<accessor>[^}]+)(?=}$)/
  // Global regular expression for extracting all property accessors enclosed in ${}
  const globalAccessorRegex = /\$\{[^}]*}/g

  /* INITIALIZATION */
  const parser = new DOMParser()
  const serializer = new XMLSerializer()

  const new_zip = new JSZip()
  const template_zip = await new_zip.loadAsync(template)

  /* PROCESSING OF THE SHARED STRINGS */
  const sharedStringsFile = template_zip.file('xl/sharedStrings.xml')
  if (!sharedStringsFile) throw new Error(`xlsx-template-browser - replaceInTemplate - Unable to get file 'xl/sharedStrings.xml' from the provided template`)
  const xmlText = await sharedStringsFile.async('string')

  const xml = parser.parseFromString(xmlText, 'application/xml')
  const sharedStringTable = xml.querySelector('sst')
  if (!sharedStringTable) throw new Error(`xlsx-template-browser - replaceInTemplate - sharedStrings.xml do not contain a <sst> delimiter`)

  const valuesToReplace: any[] = []
  let newSharedStrings: Element[] = []

  // We need to replace only text values with the new ones, therefore we process shared strings first
  // Get all string items
  xml.querySelectorAll('si').forEach((stringItem, i) => {
    // Get all rich format tags
    const r = stringItem.querySelector('r')
    if (r) {
      const xmlString = stringItem.innerHTML
      const newString = xmlString.replace(globalAccessorRegex, s => valueToString(getByNotation(data, s.replace(/\$\{|}/g, ''))))
      stringItem.innerHTML = newString
      newSharedStrings.push(stringItem)
      valuesToReplace[i] = { value: newSharedStrings.length - 1, cellType: 's' }
      return
    }

    // Get all text tags
    const t = stringItem.querySelector('t')
    if (!t) return
    const textValue = t.textContent ? t.textContent : ""
    let match = textValue.match(accessorRegex)
    // If the cell does not contain only one placeholder, check for multiple ones
    if (!match) {
      // Check if the text contains any accessors at all
      let newText = textValue.replace(globalAccessorRegex, s => valueToString(getByNotation(data, s.replace(/\$\{|}/g, ''))))
      newSharedStrings.push(createSharedString(newText, stringItem))
      valuesToReplace[i] = { value: newSharedStrings.length - 1, cellType: 's' }
      return
    }
    // Process a value with an accessor only
    let { accessor, table } = match.groups as { accessor: string, table: string }
    // If we have table values, we need to process them in a separate way
    const isTable = typeof table === 'string'
    let value = getByNotation(data, accessor)
    // Save tables for further processing
    if (isTable) {
      if (isTable && Array.isArray(value))
        valuesToReplace[i] = { isTable, value: value.map(guessDataType) }
      return
    }
    if (Array.isArray(value)) {
      valuesToReplace[i] = value.map(guessDataType)
      return false
    }
    valuesToReplace[i] = guessDataType(value)
  })

  /* PROCESSING OF WORKSHEETS */

  // To prevent string duplication, store unique strings (they will be added to shared strings)
  const newStrings: string[] = []

  // Adds a string to the collection, avoiding duplication, and returns its index.
  function addString(value: string): number {
    let stringIndex = newStrings.indexOf(value)
    if (stringIndex === -1) {
      newStrings.push(value)
      stringIndex = newStrings.length - 1
    }
    return stringIndex + newSharedStrings.length
  }

  // Get all worksheets
  const worksheets = Object.keys(template_zip.files).filter(el => /xl\/worksheets\/[^/]+$/.test(el))
  for (let worksheet of worksheets) {
    const file = template_zip.file(worksheet)
    if (!file) throw new Error(`xlsx-template-browser - replaceInTemplate - Unable to open the 'xl/worksheet/' file`)
    let xmlText = await file.async('string')
    let xml = parser.parseFromString(xmlText, 'application/xml')
    let rows = xml.querySelectorAll('sheetData row')
    const newRows: Element[] = []
    // Process cells in rows
    let rowOffset = 0
    for (let row of rows) {
      // Since some rows can be skipped, we want to get an actual row number from the template
      const rowAttribute = row.getAttribute('r')
      if (rowAttribute === null) throw new Error(`xlsx-template-browser - replaceInTemplate - The template sheets have invalid xml: rows do not posses 'r' attribute`)
      let currentRow = parseInt(rowAttribute)
      // List of new cells
      let newCells: any[] = []
      // If we have an array of values to extend the list of cells, we should keep an offset to move static cells
      let cellOffset = 0
      // Process each cell
      row.querySelectorAll('c').forEach((c) => {
        // Current cell index (note that Excel has 1-based index)
        let index = getExcelColumnIndex(c.getAttribute('r'))
        let newIndex = index + cellOffset
        // Get the cell value tag
        let v = c.querySelector('v') || {}
        // Check if the cell contains formula, and skip
        let isFormula = c.querySelector('f')
        if (isFormula) {
          const value = c
          const cellType = 'f'
          newCells[newIndex] = Object.assign({}, { value, cellType }, { template: c })
          return
        }
        // Check if the cell contains string
        let isString = c.getAttribute('t')
        if (!isString || isString !== 's') {
          newCells[newIndex] = Object.assign({}, { template: c })
          return
        }
        // Get the new value to replace
        let newValue = v instanceof Element && v.textContent !== null ? valuesToReplace[v.textContent] : {}
        if (!Array.isArray(newValue)) {
          let { value, cellType, isTable } = newValue
          // If the new value is an array from the table, then return an array of values
          if (isTable) {
            newCells[newIndex] = value.map(el => {
              if (!el) return { template: c }
              if (el.cellType === 's' && typeof el.value === 'string') el.value = addString(el.value)
              return { value: el.value, cellType: el.cellType, template: c }
            })
            return false
          }
          // Return a new value
          if (cellType === 's' && typeof value === 'string') value = addString(value)
          newCells[newIndex] = Object.assign({}, { value, cellType }, { template: c })
          return false
        }
        // If the value is an array (not from table), then extend the existing list of cells
        for (let i = 0; i < newValue.length; i++) {
          let { value, cellType } = newValue[i] || {}
          if (cellType === 's' && typeof value === 'string') value = addString(value)
          newCells[newIndex + i] = Object.assign({}, { value, cellType }, { template: c })
          if (i) cellOffset++
        }
      })

      /*  GENERATION OF THE NEW ROWS */

      // Check if the row contains arrays and the values should be duplicated
      let length = Math.max(...newCells.filter(el => Array.isArray(el)).map(el => el.length))

      // Create a row
      if (length <= 0) {
        let rowIndex = rowOffset + currentRow
        const rowValues = newCells.map((el, i) => {
          if (!el) return undefined
          return createCell(Object.assign({}, el, { row: rowIndex, column: i }))
        }).filter(Boolean)
        const newRow = row.cloneNode() as Element
        newRow.setAttribute('r', rowIndex.toString())
        rowValues.forEach(value => newRow.append(value))
        newRows.push(newRow)
        continue
      }

      // Create table rows
      for (let i = 0; i < length; i++) {
        let rowIndex = rowOffset + currentRow
        rowOffset++
        const rowValues = newCells.map((el, index) => {
          if (!el) return undefined
          if (Array.isArray(el)) {
            if (typeof el[i] === 'undefined') return undefined
            return createCell(Object.assign({}, el[i], { row: rowIndex, column: index }))
          }
          return createCell(Object.assign({}, el, { row: rowIndex, column: index }))
        })
        const newRow = row.cloneNode() as Element
        newRow.setAttribute('r', rowIndex.toString())
        rowValues.forEach(value => newRow.append(value))
        newRows.push(newRow)
      }
    }

    // Generate new sheet data
    const SheetData = xml.querySelector('sheetData')
    if (!SheetData) throw new Error(`xlsx-template-browser - replaceInTemplate - Unable the find the selector: sheetData`)
    const newSheetData = SheetData.cloneNode() as Element
    newRows.forEach(row => newSheetData.append(row))
    SheetData.replaceWith(newSheetData)
    const newData = serializer.serializeToString(xml).replace(/ ?xmlns="http:\/\/www\.w3\.org\/1999\/xhtml"/g, '')
    // Save the data of the worksheet
    new_zip.file(worksheet, newData)
  }

  /* FINALIZE NEW SHARED STRINGS */
  newSharedStrings = [...newSharedStrings, ...newStrings.map(el => createSharedString(el))]
  const newSst = sharedStringTable.cloneNode() as Element
  newSharedStrings.forEach(si => newSst.append(si))
  sharedStringTable.replaceWith(newSst)
  const newData = serializer.serializeToString(xml).replace(/ ?xmlns="http:\/\/www\.w3\.org\/1999\/xhtml"/g, '')
  // Save new shared strings
  new_zip.file('xl/sharedStrings.xml', newData)
  return new_zip.generateAsync({ type: 'blob' })
}
export default replaceInTemplate

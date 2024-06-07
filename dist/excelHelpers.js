"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.escapeXml = exports.getExcelColumnIndex = exports.createCell = exports.createSharedString = exports.guessDataType = exports.valueToString = void 0;
function valueToString(value) {
    if (!value)
        return '';
    if (Array.isArray(value))
        return value.map(valueToString).toString();
    if (typeof value === 'object')
        value = JSON.stringify(value);
    return value.toString();
}
exports.valueToString = valueToString;
function dateToExcel(value) {
    let date = new Date(value);
    return 25569 + ((date.getTime() - (date.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
}
/**
 * Guesses the Excel data type for the given JavaScript primitive value.
                        or undefined if the data type is not recognized.
 */
function guessDataType(value) {
    if (typeof value === 'undefined')
        return { cellType: 's', value: '' };
    if (typeof value === 'number') {
        if (isFinite(value))
            return { cellType: 'n', value };
        return undefined;
    }
    if (value instanceof Date) {
        if (isNaN(value.getDate()))
            return undefined;
        return { cellType: 'n', value: dateToExcel(value) };
    }
    if (typeof value === 'boolean')
        return { cellType: 'b', value: value ? 1 : 0 };
    return { cellType: 's', value: valueToString(value) };
}
exports.guessDataType = guessDataType;
/**
 * Creates a new shared string for an Excel file.
 * @param {string} text - The text content of the shared string.
 * @param {Element} template - (Optional) An existing node with cell to clone for the shared string.
 * @returns {Element} - The newly created or cloned shared string element.
 */
function createSharedString(text, template) {
    const si = template ? template.cloneNode() : document.createElement('si');
    let t = document.createElement('t');
    t.textContent = text;
    si.appendChild(t);
    return si;
}
exports.createSharedString = createSharedString;
/**
 * Creates a cell with the specified value.
 * @param {Object} options - An object containing parameters for cell creation.
 * @param {*} options.value - The value to be set in the cell.
 * @param {string|number} options.row - The row reference for the cell.
 * @param {string|number} options.column - The column reference (either string or numeric) for the cell.
 * @param {Element} options.template - (Optional) An existing template to clone for the cell.
 * @param {string} options.cellType - The type of the cell (e.g., 's', 'n', 'b', 'f') for Excel formatting.
 * @returns {Element} - The created cell element.
 */
const createCell = ({ value, row, column, template, cellType }) => {
    // If the cell is not a string, then skip
    if (!cellType && !value) {
        return template;
    }
    // If the cell is formula, then skip
    if (cellType === 'f') {
        // Check precalculated value
        const v = value.querySelector('v');
        if (v)
            v.remove();
        return value;
    }
    // Clone initial node with the style or create a new one
    const cell = template ? template.cloneNode() : document.createElement('c');
    // Create a cell reference
    if (row && column) {
        column = typeof column === 'number' ? getExcelColumnName(column) : column;
        cell.setAttribute('r', column + row);
    }
    // If no value is provided, return cell
    if (typeof value === 'undefined')
        return cell;
    const valueTag = document.createElement('v');
    cell.setAttribute('t', cellType);
    valueTag.textContent = value;
    cell.appendChild(valueTag);
    return cell;
};
exports.createCell = createCell;
/**
 * Converts a numeric index to an Excel column name.
 */
const getExcelColumnName = (index) => {
    let result = '';
    while (index > 0) {
        const remainder = (index - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        index = Math.floor((index - 1) / 26);
    }
    return result;
};
/**
 * Converts an Excel column name to its numeric index.
 * @param {string} columnName - The Excel column name to be converted.
 * @returns {number} - The numeric index corresponding to the Excel column name.
 */
const getExcelColumnIndex = (columnName) => {
    columnName = columnName.replace(/\d+/g, '');
    let index = 0;
    for (let i = 0; i < columnName.length; i++) {
        const charCode = columnName.charCodeAt(i) - 64;
        index = index * 26 + charCode;
    }
    return index;
};
exports.getExcelColumnIndex = getExcelColumnIndex;
/**
 * Escapes unsafe XML characters in a given string to ensure its safe inclusion
 * in XML documents, preventing parsing issues.
 *
 * @param {string} unsafe - The input string containing XML-unsafe characters.
 * @returns {string} - A new string with XML-unsafe characters replaced by their
 * corresponding HTML entities.
 */
const unsafeChars = {
    '<': '&lt;',
    '>': '&gt;',
    '&': '&amp;',
    '\'': '&apos;',
    '"': '&quot;'
};
const escapeXml = (unsafe) => {
    return unsafe.replace(/[<>&'"]/g, (c) => { return unsafeChars[c]; });
};
exports.escapeXml = escapeXml;

export declare function valueToString(value: string | string[] | object): string;
/**
 * Guesses the Excel data type for the given JavaScript primitive value.
                        or undefined if the data type is not recognized.
 */
export declare function guessDataType(value: Object): {
    cellType: string;
    value: string;
} | {
    cellType: string;
    value: number;
};
/**
 * Creates a new shared string for an Excel file.
 * @param {string} text - The text content of the shared string.
 * @param {Element} template - (Optional) An existing node with cell to clone for the shared string.
 * @returns {Element} - The newly created or cloned shared string element.
 */
export declare function createSharedString(text: string, template?: Element): Element;
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
export declare const createCell: ({ value, row, column, template, cellType }: {
    value: any;
    row: any;
    column: any;
    template: any;
    cellType: any;
}) => any;
/**
 * Converts an Excel column name to its numeric index.
 * @param {string} columnName - The Excel column name to be converted.
 * @returns {number} - The numeric index corresponding to the Excel column name.
 */
export declare const getExcelColumnIndex: (columnName: any) => number;
export declare const escapeXml: (unsafe: string) => string;

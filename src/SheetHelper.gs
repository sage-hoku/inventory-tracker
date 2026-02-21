
/**
 * Helper class for working with Google Sheets.
 * Provides convenient access to named ranges, columns, and common sheet operations.
 */
class SheetHelper {
    /**
     * Creates a new SheetHelper instance.
     * @param {string} name - The name of the sheet to wrap
     */
    constructor(name) {
        this.sheet = importSheet(name);
        this.init();
    }

    /**
     * Initializes cached named ranges and columns.
     * Called automatically by constructor.
     */
    init() {
        this.namedRangesFull = Object.fromEntries(
            this.sheet.getNamedRanges().map(range => [range.getName(), range.getRange()])
        );
        
        this.namedRanges = Object.fromEntries(
            this.sheet.getNamedRanges().map(range => [range.getName(), stripRange(range.getRange())])
        );

        this.namedColumns = Object.fromEntries(
            this.sheet.getNamedRanges().map(range => [
                range.getName(), 
                flatten(stripRange(range.getRange(), {values: true}))
            ])
        );
    }

    /**
     * Gets a named range or range by column number.
     * @param {string|number} target - Named range name or column number
     * @returns {Range} The requested range
     * @throws {Error} If target is empty or invalid type
     */
    range(target) {
        if (!target && target !== 0) throw new Error('.range given empty target');
        
        if (typeof target === 'string') {
            return this.namedRanges[target];
        }
        
        if (typeof target === 'number') {
            return this.getRange(1, target, this.getLastRow());
        }
        
        throw new Error(`Invalid target type: ${typeof target}`);
    }

    /**
     * Gets column values from a named range or column number.
     * @param {string|number} target - Named range name or column number
     * @returns {Array} Flattened array of column values
     * @throws {Error} If target is empty or invalid type
     */
    col(target) {
        if (!target && target !== 0) throw new Error('.col given empty target');
        
        if (typeof target === 'string') {
            return this.namedColumns[target];
        }
        
        if (typeof target === 'number') {
            return flatten(stripRange(this.getRange(1, target, this.getLastRow()), {values: true}));
        }
        
        throw new Error(`Invalid target type: ${typeof target}`);
    }

    /**
     * Gets the last row with data in the sheet.
     * @returns {number} Last row number
     */
    getLastRow() {
        return this.sheet.getLastRow();
    }

    /**
     * Gets the last column with data in the sheet.
     * @returns {number} Last column number
     */
    getLastColumn() {
        return this.sheet.getLastColumn();
    }

    /**
     * Proxy method for sheet.getRange().
     * @param {...*} args - Arguments to pass to getRange
     * @returns {Range} The requested range
     */
    getRange(...args) {
        return this.sheet.getRange(...args);
    }
  
    /**
     * Gets the name of the sheet.
     * @returns {string} Sheet name
     */
    getName() {
        return this.sheet.getName();
    }

    /**
     * Gets the number of frozen header rows.
     * @returns {number} Number of frozen rows
     */
    getHeaderRows() {
        return this.sheet.getFrozenRows();
    }

    /**
     * Gets all data in the sheet (excluding headers).
     * @returns {Range} Range containing all data
     */
    getData() {
        return this.getRange(this.getHeaderRows() + 1, 1, this.getLastRow() - this.getHeaderRows(), this.getLastColumn());
    }

    /**
     * Auto-resizes columns to fit content.
     * @param {...*} args - Arguments to pass to autoResizeColumns
     * @returns {Sheet} The sheet for chaining
     */
    autoResizeColumns(...args) {
        return this.sheet.autoResizeColumns(...args);
    }

    /**
     * Clears a rectangular region of cells.
     * @param {number} x - Starting row
     * @param {number} y - Starting column
     * @param {number} length - Number of rows to clear
     * @param {number} [width=1] - Number of columns to clear
     * @returns {Range} The cleared range
     */
    clearCells(x, y, length, width = 1) {
        return this.sheet.getRange(x, y, length, width).clear();
    }

    /**
     * Measures the last row with non-empty text across a set of contiguous columns.
     * Columns are checked left-to-right.
     * 
     * @param {number} startCol - The first column to check
     * @param {number} [endCol] - The final column to check (inclusive). Defaults to startCol.
     * @returns {number} The lowest populated row in the range
     */
    getLastRowByColumn(startCol, endCol = undefined) {
        if (endCol === undefined) endCol = startCol;
        
        const distance = (endCol - startCol + 1) ?? 0;

        return Math.max(
            ...Array.from({ length: distance }, (_, i) => getLastRow(this.sheet, i + startCol))
        );
    }
    
    /**
     * Gets a named range with advanced stripping options.
     * @param {string} name - Named range name
     * @param {Object} [options={}] - Stripping options (see stripRange)
     * @returns {Range|Array} Modified range or its values
     */
    getNamedRangeAdvancedOptions(name, options = {}) {
        const range = this.namedRanges[name];
        return stripRange(range, options);
    }

    /**
     * Deletes rows where a column matches a specific value.
     * @param {number} column - Column number to check
     * @param {*} value - Value to match for deletion
     * @returns {boolean} True if successful
     */
    static deleteRowIf(column, value) {
        const lastRow = getLastRowByColumn(column);
        const verticalOffset = sheet.getFrozenRows() + 1;
        
        const rangeToCheck = sheet.getRange(verticalOffset, column, lastRow);
        const columnValues = flatten(rangeToCheck.getValues());
        
        // Delete rows matching value, bottom to top
        for (let r = lastRow - verticalOffset; r >= 0; r--) {
            if (columnValues[r] === value) {
                sheet.deleteRow(r + verticalOffset);
            }
        }

        return true;
    }
}

/**
 * Gets a sheet by name from the active spreadsheet.
 * @param {string} name - Sheet name
 * @returns {Sheet} The requested sheet
 */
function importSheet(name) {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

/**
 * Flattens a 2D array into a 1D array.
 * @param {Array[]} array - 2D array to flatten
 * @returns {Array} Flattened 1D array
 */
function flatten(array) {
    return [].concat(...array);
}

/**
 * Finds the last non-empty row in a specific column.
 * @param {Sheet} sheet - The sheet to check
 * @param {number} columnIndex - Column number (1-indexed)
 * @returns {number} The last non-empty row number
 */
function getLastRow(sheet, columnIndex) {
    const lastRow = sheet.getLastRow();
    return lastRow - sheet.getRange(1, columnIndex, lastRow)
        .getValues()
        .reverse()
        .findIndex(c => c[0] !== '');
}

/**
 * Returns a range with headers and/or trailing empty rows removed.
 * 
 * @param {Range} range - The range to modify
 * @param {Object} [options={}] - Configuration options
 * @param {boolean} [options.stripHeaders=true] - Remove frozen rows from top
 * @param {boolean} [options.stripEmpty=true] - Remove empty rows from bottom
 * @param {boolean} [options.values=false] - Return values instead of range
 * @returns {Range|Array} Modified range or its values
 */
function stripRange(range, options = {}) {
    const {
        stripHeaders = true,
        stripEmpty = true,
        values = false
    } = options;

    // Return early if no modifications needed
    if (!stripHeaders && !stripEmpty) {
        return values ? range.getValues() : range;
    }

    const sheet = range.getSheet();
    const startCol = range.getColumn();
    const numColumns = range.getNumColumns();

    // Calculate start row (skip headers if requested)
    const startRow = stripHeaders 
        ? Math.max(sheet.getFrozenRows() + 1, range.getRow())
        : range.getRow();

    // Calculate end row (trim empty rows if requested)
    const endRow = stripEmpty
        ? getLastRow(sheet, startCol)
        : range.getLastRow();

    const numRows = Math.max(1, endRow - startRow + 1);

    const result = sheet.getRange(startRow, startCol, numRows, numColumns);
    return values ? result.getValues() : result;
}

/**
 * Displays a notification card (for Google Workspace Add-ons).
 * @param {string} text - Text to display
 * @returns {Card} A card with the notification
 */
function notify(text) {
    Logger.log(text);
    return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle(text))
        .build();
}

class SheetHelper 
{
    constructor (name) 
    {
        this.sheet = importSheet(name);

        this.init(); 
    }

    init() 
    {
        this.namedRangesFull = Object.fromEntries(
            this.sheet.getNamedRanges().map(range => [range.getName(), range.getRange()])
        );
        
        this.namedRanges = Object.fromEntries(
            this.sheet.getNamedRanges().map(range => [range.getName(), stripRange(range.getRange())])
        )

        this.namedColumns = Object.fromEntries(
            this.sheet.getNamedRanges().map(range => [
                range.getName(), 
                flatten(stripRange(range.getRange(), {values: true}))
            ]) 
        );
    } 

    range(target) {
        if (!target) throw new Error('.range given empty target');
        
        if (typeof target === 'string') {
            return this.namedRanges[target];
        }
        
        if (typeof target === 'number') {
            return this.getRange(1, target, this.getLastRow());
        }
        
        throw new Error(`Invalid target type: ${typeof target}`);
    }

    col(target) {
        if (!target) throw new Error('.col given empty target');
        
        if (typeof target === 'string') {
            return this.namedColumns[target];
        }
        
        if (typeof target === 'number') {
            return stripRange(this.getRange(1, target, this.getLastRow()), {values: true});
        }
        
        throw new Error(`Invalid target type: ${typeof target}`);
    }

    getLastRow() {
        return this.sheet.getLastRow();
    }

    getLastColumn() {
        return this.sheet.getLastColumn();
    }

    // Proxy all original Sheet methods
    getRange(...args) {
        return this.sheet.getRange(...args);
    }
  
    getName() {
        return this.sheet.getName();
    }

    getHeaderRows() {
        return this.sheet.getFrozenRows();
    }

    getData() {
        return this.getRange(this.getHeaderRows(),1,this.getLastRow(),this.getLastColumn()); 
    }

    autoResizeColumns(...args) {
        return this.sheet.autoResizeColumns(...args);
    }

    clearCells(x, y, length, width = 1) 
    {
        return this.sheet.getRange(x,y, length, width).clear(); 
    }

    /**
     * Measures the last row with non-empty text of a set of contiguous columns. 
     * Columns are checked left-to-right.
     *  
     *  !Tip: for a full Range, use range.getLastRow()
     *
     *  @param   {Number} startCol = the first column to check. 
     *  @param   {Number} endCol   = the final column to check, inclusive.
     *
     *  @returns {Number}        ==> a number representing the lowest populated row in the range. 
     */
    getLastRowByColumn(startCol, endCol = undefined) 
    {
        if (endCol === undefined) endCol = startCol;
        
        const distance = (endCol - startCol + 1) ?? 0;

        return Math.max(
            ...Array.from({ length: distance }, (_, i) => getLastRow(this.sheet, i + startCol))
        );
    }
    
    getNamedRangeAdvancedOptions(name, options = {}) {
        const range = this.namedRanges[name];
        return stripRange(range, options);
    }

    static deleteRowIf(column, value) {
        // Get last column and header rows
        const lastRow = getLastRowByColumn(column);
        const verticalOffset = sheet.getFrozenRows() + 1;
        
        // Build table of column values
        const rangeToCheck = sheet.getRange(verticalOffset, column, lastRow);
        const columnValues = [].concat(...rangeToCheck.getValues());
        
        // Delete rows matching value, bottom to top.
        for (var r = lastRow - verticalOffset; r >= 0; r--) {
            if (columnValues[r] === value) {
            sheet.deleteRow(r + verticalOffset);
            } 
        }

        return true;
    }
}

function importSheet(name) 
{
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
} 

function flatten(array) 
{
    return [].concat(...array);
}

function getLastRow(sheet, columnIndex) 
{
    const lastRow = sheet.getLastRow();
    return lastRow - sheet.getRange(1, columnIndex, lastRow)
      .getValues()              // 2-D array
      .reverse()                // Search from bottom-up
      .findIndex(c=>c[0]!='')   // Take only the first element
    ;  
}

/**
 * Returns a range with headers and/or trailing empty rows removed.
 *
 * @param   {Range}   range                 = The range to modify
 * @param   {Object}  options               = Configuration options
 * @param   {Boolean} options.stripHeaders .. Remove frozen rows from top (default: true)
 * @param   {Boolean} options.stripEmpty   .. Remove empty rows from bottom (default: true)
 * @param   {Boolean} options.values       .. Return values instead of range (default: false)
 * @returns {Range|Array}                 ==> Modified range or its values
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

function getNamedRange(name, options = undefined) 
    {
        options = 
        {
            'showHeader': false, 
            'showLastRowsIncludingEmpty': false,
            ...options
        };
        
        const range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(name);
        const sheet = range.getSheet();

        const startCol   = range.getColumn();
        const endCol     = range.getLastColumn();
        const numColumns = range.getNumColumns();

        const startRow = (options && options.showHeader === false) ?
            Math.max(
            sheet.getFrozenRows() + 1, 
            range.getRow()
            ) : 
            range.getRow()
        ;

        const endRow = (options && options.showLastRowsIncludingEmpty === false) ?
            getLastRowByRange(sheet, startCol, endCol) :
            range.getLastRow()
        ;

        const numRows = (endRow - startRow) + 1;

        return sheet.getRange(startRow, startCol, numRows, numColumns); 
}

function notify(text) 
{
    Logger.log(text);
    return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle(text))
        .build()
    ;
}
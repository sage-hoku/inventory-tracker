class SheetHelper 
{
    constructor (name) 
    {
        this.sheet = importSheet(name);

        this.init(); 
    }

    init() 
    {
        this.namedRanges = Object.fromEntries(
            this.sheet.getNamedRanges().map(range => [range.getName(), range.getRange()])
        );

        this.namedColumns = Object.fromEntries(
            this.sheet.getNamedRanges().map(range => [
                range.getName(), 
                getNamedRangeAdvancedOptions(range.getName())
            ]) 
        );
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
    
    /**
     * Returns a modified version of ae range depending on the options specified.
     *
     * @param   {String}  name         = the name of the NamedRange to search for.
     * @param   {Object}  options.     = {
     * @param   {Boolean} values           Default: False, see @returns below. 
     * @param   {Boolean} hideHeader       Default: True, removes header rows from range.
     * @param   {Boolean} hideEmpty        Default: True, removes empty rows after populated range.
     *                                   } 
     * @returns {Any}                ==> Either returns a Range or Array of values,
     *                                   depending on whether options.values flag is true.                                             
     */
    getNamedRangeAdvancedOptions(name, options = undefined) 
    {   
        // Use default options if none provided.
        options = 
        {   
            'values':      false,
            'hideHeader':  true, 
            'hideEmpty':   true,
            ...options  // user inputs will override defaults.
        };

        // Match by name of range
        const range = this.namedRanges.find((NR) => NR.getName() === name ) 
            .getRange();

        // Return early if all options are false
        if (Object.values(options).every(v => !v)) return range;

        // === Build a getRange() query (row, col, numRows, numCols) === //
        
        // Cols are easy
        const startCol   = range.getColumn();
        const endCol     = range.getLastColumn();
        const numColumns = range.getNumColumns();

        // Rows depend on 'options' settings
        // If hideHeader, exclude frozen rows.
        const startRow = (options.hideHeader === true) ?
            Math.max(
                sheet.getFrozenRows() + 1, 
                range.getRow()
            ) 
            : range.getRow();

        // If !hideEmpty, we end the range at the last non-empty row. 
        const endRow = (options.hideEmpty === true) ?
            getLastRowInArea(startCol, endCol) : range.getLastRow();``

        const numRows = (endRow - startRow) + 1;

        const result = sheet.getRange(startRow, startCol, numRows, numColumns);     
        return (options.values === true) ? result : result.getValues();
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
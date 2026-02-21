// #region "Setup"
/**
 * Initializes SheetHelper instances for each sheet.
 *
 * @returns {Object} Object containing SheetHelper instances for each sheet
 */
function initializeSheets() {
    return {
        Database:     new SheetHelper('Persistent Database'),
        Groceries:    new SheetHelper('Grocery List'),
        Import:       new SheetHelper('Costco Import'),
        Inventory:    new SheetHelper('Apartment Inventory'),
        Meta:         new SheetHelper('Metadata'),
        Recipes:      new SheetHelper('Recipes'),
        PriceTracker: new SheetHelper('Price Tracker')
    };
}

// Constants
const IMPORT_HEADER_ROWS = importSheet('Costco Import').getFrozenRows();
const DATABASE_HEADER_ROWS = importSheet('Persistent Database').getFrozenRows();
const TODAY = new Date().toJSON().slice(0, 10);

/**
 * Copies specified values from an object to a range, optionally also to another object.
 * 
 * @param   {Range}    range    = Target range to write values
 * @param   {string[]} copyKeys = Array of object keys to copy
 * @param   {Object}   fromItem = Source object
 * @param   {Object}   [toItem] ~ Optional destination object to update
 * @returns {Object}          ==> Result object with copied status and error message
 */
function copyValuesToRange(range, copyKeys, fromItem, toItem = undefined) {
    const result = {copied: false, errorMessage: ''};
    try {
        const values = [];

        copyKeys.forEach(key => {
            if (toItem) toItem[key] = fromItem[key];
            values.push(fromItem[key]);
        });

        range.setValues([values]);
        result.copied = true;
        return result;
    } catch (e) {
        result.copied = false;
        result.errorMessage = e.message || e.toString();
        return result;
    }
}

/**
 * Creates formulas to check if items exist in the database.
 *
 * @returns {Object} Result object containing:
 *   - values: Array of validation values
 *   - errorMessage: Error message if operation failed
 */
function activateValidationColumn() {
    const result = { values: undefined, errorMessage: '' };
    try {
        const { Import } = initializeSheets();
        // Validation formula for checking IDs against persistent database
        const VAL_STR = 
`=IF(
    AND(
        INDEX(IMPORT_Item_ID, ROW() - ${IMPORT_HEADER_ROWS})<>"",
        INDEX(IMPORT_Item_Label, ROW() - ${IMPORT_HEADER_ROWS})<>"",
        INDEX(IMPORT_Price, ROW() - ${IMPORT_HEADER_ROWS})<>""
    ), 
    IFERROR(
        MATCH(
            INDEX(IMPORT_Item_ID, ROW() - ${IMPORT_HEADER_ROWS}), 
            DB_Item_ID, 
            0
        ),
        FALSE
    ),
    ""
)`;
        const importIDs = Import.col('IMPORT_Item_ID');
        const validationColumn = Import.range('IMPORT_Valid');

        // Set validation formulas
        Import.getRange(
            validationColumn.getRow(),
            validationColumn.getColumn(),
            importIDs.length
        ).setFormula(VAL_STR);
        result.values = flatten(validationColumn.getValues());
        return result;
    }
    catch (e) {
        result.errorMessage = e.message;
        return result;
    } 
}

// #endregion

// #region "Formatting"

// Take newly pasted receipts and format according to the purchase location

function runFormat() {
    const formatResult = formatPastedRows();
    if (formatResult.errorMessage) {
        return SpreadsheetApp.getActive().toast(
            `Unable to format data: ${formatResult.errorMessage}`, 
            "Format Failed"
        );
    } else {
        return SpreadsheetApp.getActive().toast(
            'Formatted ' + formatResult.processedData.length + ' pasted items.',
        );
    }
}

/**
 * Formats pasted receipt data based on the purchase location.
 * Currently supports Costco receipts.
 * 
 * @returns {Object} ==> Result object containing:
 *   - processedData: Array of processed items
 *   - errorMessage: Error message if operation failed
 */
function formatPastedRows() {
    const { Import } = initializeSheets();
    const result = {processedData: [], errorMessage: ''};
    
    try {
        const purchasedLoc = Import.col('IMPORT_Source')[0];

        switch (purchasedLoc.toString().toUpperCase()) {
            case 'COSTCO': 
                Object.assign(result, formatCostcoReceiptData());
                break;

            default:
                result.errorMessage = `Invalid purchase location: "${purchasedLoc}"`;
                break;
        }

        return result;
    } catch (e) {
        result.processedData = [];
        result.errorMessage = e.message || e.toString();
        return result;
    }
}

// #endregion

// #region "Validation"

// Take all items from the import region (newly pasted receipts) and match them against the persistent database
// If new, IMPORT_Last_Price should be empty. 

function runValidate() {
    const validationResult = checkValidItems();
    if (validationResult.errorMessage) {
        return SpreadsheetApp.getActive().toast(
            `Unable to validate data: ${validationResult.errorMessage}`, 
            "Validation Failed"
        );
    } else {
        return SpreadsheetApp.getActive().toast(
            'Validated ' + validationResult.newItems.length + ' new items.'
            + '\n' + validationResult.validItems.length + ' items matched.'
            + '\n' + 'Please review the results in the "User Input" region.',
        );
    }
}

/**
 * Finds and validates imported items against the persistent database.
 * Creates formulas to check if items exist in the database.
 *
 * @returns {Object} Result object containing:
 *   - newItems: Array of items not found in database
 *   - validItems: Array of items matched in database
 *   - errorMessage: Error message if operation failed
 */
function checkValidItems() {
    const { Database, Import } = initializeSheets();
    
    const result = {
        newItems: [], 
        validItems: [], 
        errorMessage: ''
    };

    try {
        const valid = activateValidationColumn().values;

        // Copy all validated rows to the input area, 1 row at a time
        for (let i = 0; i < valid.length; i++)                      
        {
            // Define rows based on number of sheet header rows
            const activeRow = (IMPORT_HEADER_ROWS + 1) + i;
            const foundRow = valid[i] - DATABASE_HEADER_ROWS;
            
            // Clear EVERY row
            Import.clearCells(activeRow, 1, 1, 5);

            // Skip empty rows
            if (valid[i] === '') continue;
            
            // Create a new item if no valid match  
            if (!valid[i]) {
                result.newItems.push({
                    itemID: Import.col('IMPORT_Item_ID')[i],
                    row: activeRow
                });
                continue;
            }

            // Build item object from import sheet
            const item = {
                name:      Import.col('IMPORT_Item_Name')[i],
                category:  Import.col('IMPORT_Category')[i],
                quantity:  Import.col('IMPORT_Quantity')[i],
                unit:      Import.col('IMPORT_Unit')[i],
                lastPrice: Import.col('IMPORT_Last_Price')[i],
                valid:     Import.col('IMPORT_Valid')[i],
                id:        Import.col('IMPORT_Item_ID')[i],
                label:     Import.col('IMPORT_Item_Label')[i],
                purchday:  Import.col('IMPORT_Date_Purchased')[i],
                purchloc:  Import.col('IMPORT_Source')[i],
                row:       activeRow,
                match:     undefined
            };

            // Build database item object
            const databaseItem = {
                row:       foundRow,
                match:     item,
                category:  Database.col('DB_Category')[foundRow],
                quantity:  Database.col('DB_Quantity')[foundRow],
                unit:      Database.col('DB_Unit')[foundRow],
                name:      Database.col('DB_Item_Name')[foundRow],
                lastPrice: Database.col('DB_Price')[foundRow],
                id:        Database.col('DB_Item_ID')[foundRow],
                label:     Database.col('DB_Item_Abbrv')[foundRow],
                purchday:  Database.col('DB_Last_Purchased')[foundRow],
                purchloc:  Database.col('DB_Location_Purchased')[foundRow]
            };

            item.match = databaseItem;
            
            // Copy database item into working area
            const copyKeys = ['name', 'category', 'quantity', 'unit', 'lastPrice'];
            const targetRange = Import.getRange(item.row, 1, 1, 5);

            const copyOperation = copyValuesToRange(targetRange, copyKeys, databaseItem, item);
            
            if (copyOperation.errorMessage) { throw new Error(copyOperation.errorMessage); }

            if (copyOperation.copied) {
                result.validItems.push(item);
            }
        }
        return result;
    } catch (e) {
        result.errorMessage = e.message || e.toString();
        return result;
    }
}

// #endregion

// #region "Data Entry"

// Append all validated items to the persistent database, overwriting old values.

function runEntry() {
    const entryResult = appendItemsToDB();
    if (entryResult.errorMessage) {
        return SpreadsheetApp.getActive().toast(
            `Unable to update Database: ${entryResult.errorMessage}`, 
            "Update Failed"
        );
    } else {
        return SpreadsheetApp.getActive().toast(
            'Added ' + entryResult.newItems.length + ' new items.'
            + '\n' + entryResult.validItems.length + ' items updated.'
            + '\n' + entryResult.priceChanges + ' price changes tracked.'
        );
    }
}

/**
 * Appends new and updated items to the persistent database.
 * Updates existing items if they have empty fields.
 * + TODO: Tracks price changes in the Price Tracker sheet.
 * 
 * @returns {Object} Result object containing:
 *   - newIrems: Array of new items added
 *   - validItems: Array of existing items updated
 *  [- priceChanges: Number of price changes tracked] <- not implemented
 *   - errorMessage: Error message if operation failed
 */
function appendItemsToDB() {
    const { Database, Import }
    const result = {
        newItems: [], 
        validItems: [], 
        errorMessage: ''
    };

    try {
        const valid = Import.col('IMPORT_Valid');
        
        for (let i = 0; i < Import.getLastRow() i++) {
            // Skip empty
            if (valid[i] === '') { continue; }

            // Get data from import sheet
            const itemData = {
                row:       IMPORT_HEADER_ROWS + i;
                id:        Import.col('IMPORT_Item_ID')[i],
                label:     Import.col('IMPORT_Item_Label')[i],
                name:      Import.col('IMPORT_Item_Name')[i],
                category:  Import.col('IMPORT_Category')[i],
                quantity:  Import.col('IMPORT_Quantity')[i],
                unit:      Import.col('IMPORT_Unit')[i],
                price:     Import.col('IMPORT_Price')[i],
                purchday:  Import.col('IMPORT_Date_Purchased')[0] ?? TODAY,
                purchloc:  Import.col('IMPORT_Source')[0] ?? 'Unknown'
            };

            // If valid, append to existing row. Otherwise, append to end of database.
            const databaseRow = (valid[i]) ? valid[i] - DATABASE_HEADER_ROWS : Database.getLastRow() + 1;
            const databaseRange = Database.getRange(databaseRow + 1, 1, 1, 9);
            const copyOperation = copyValuesToRange(
                databaseRange,
                ['id', 'label', 'name', 'category', 'quantity', 'unit', 'price', 'purchday', 'purchloc'],
                itemData
            );
            
            if (valid[i]) { result.validItems.push(itemData); }
            else { result.newItems.push(itemData); }
            
            // // Add initial price to Price Tracker
            // const ptLastRow = PriceTracker.getLastRow() + 1;
            // PriceTracker.getRange(ptLastRow, 1, 1, 3).setValues([[
            //     itemData.id,
            //     itemData.price,
            //     TODAY
            // ]]);
        }

    } catch (e) {
        result.copied = false;
        result.errorMessage = e.message || e.toString();
        return result;
    }
    appendToDB();
}

// #endregion

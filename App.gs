function initializeSheets() {
    return {
      Database:  new SheetHelper('Persistent Database'),
      Groceries: new SheetHelper('Grocery List'),
      Import:    new SheetHelper('Costco Import'),
      Inventory: new SheetHelper('Apartment Inventory'),
      Meta:      new SheetHelper('Metadata'),
      Recipes:   new SheetHelper('Recipes'),
      };
}

const IMPORT_HEADER_ROWS = importSheet('Costco Import').getFrozenRows();
const DATABASE_HEADER_ROWS = importSheet('Persistent Database').getFrozenRows();

// Validation string used for checking IDs against persistent database
// ID if valid, False if not found, 
const VAL_STR = 
`=IF(
    AND(
        INDEX(
          IMPORT_Item_ID,   ROW() - ${IMPORT_HEADER_ROWS}
        )<>"",
        INDEX(
          IMPORT_Item_Label, ROW() - ${IMPORT_HEADER_ROWS}
        )<>"",
        INDEX(
          IMPORT_Price,     ROW() - ${IMPORT_HEADER_ROWS}
        )<>"",
    ), 
    IFERROR(
        MATCH(
            INDEX(
              IMPORT_Item_ID, ROW() - ${IMPORT_HEADER_ROWS}
            ), 
            DB_Item_ID, 0
        ),
        FALSE
    )
)`;

const TODAY = new Date().toJSON().slice(0, 10);

function findValidItems() {
    // create SheetHelpers for Database and Import sheets
    const { Database, Import } = initializeSheets();
    
    // return value 
    const result = {'newItems': [], 'validItems': [], 'errorMessage': ''};

    try 
    { 
        const validationColumn = Import.range('IMPORT_Valid');
        const importIDs        = Import.col('IMPORT_Item_ID');

        Import.getRange(
            validationColumn.getRow(),
            validationColumn.getColumn(),
            importIDs.length
        ).setFormula(VAL_STR);                                   
         
        const valid = flatten(validationColumn.getValues());

        // Copy all validated rows to the input area, 1 row at a time
        for (var i = 0; i < importIDs.length; i++)                      
        {
            const activeRow = (IMPORT_HEADER_ROWS + 1) + i;
            Import.clearCells( activeRow, 1,  1, 5 );   // clear input.

            // Skip empty rows
            if (valid[i] === '') { continue; } 
            
            // Create a new item if no valid match  
            if (!valid[i]) 
            { 
                result.newItems.push({
                    'itemID': Import.col('IMPORT_Item_ID')[i],
                    'row': activeRow
                });
                continue; 
            }

            const item = { 
                'name':      Import.col('IMPORT_Item_Name')[i],
                'category':  Import.col('IMPORT_Category')[i],
                'quantity':  Import.col('IMPORT_Quantity')[i],
                'unit':      Import.col('IMPORT_Unit')[i],
                'lastPrice': Import.col('IMPORT_Last_Price')[i],
                'valid':     Import.col('IMPORT_Valid')[i],
                'id':        Import.col('IMPORT_Item_ID')[i],
                'label':     Import.col('IMPORT_Item_Label')[i],
                'purchday':  Import.col('IMPORT_Date_Purchased')[i],
                'purchloc':  Import.col('IMPORT_Source')[i],
                'row':       activeRow,
                'match':     undefined
            };
            
            const foundRow = valid[i] - DATABASE_HEADER_ROWS;

            const databaseItem = {
                'row':       foundRow,
                'match':     item,
                'category':  Database.col('DB_Category')[foundRow],
                'quantity':  Database.col('DB_Quantity')[foundRow],
                'unit':      Database.col('DB_Unit')[foundRow],
                'name':      Database.col('DB_Item_Name')[foundRow],
                'lastPrice': Database.col('DB_Price')[foundRow],
                'id':        Database.col('DB_Item_ID')[foundRow],
                'label':     Database.col('DB_Item_Abbrv')[foundRow],
                'purchday':  Database.col('DB_Last_Purchased')[foundRow],
                'purchloc':  Database.col('DB_Location_Purchased')[foundRow]
            }

            item.match = databaseItem;
            
            // copy database item into working area
            const copyKeys = ['name', 'category', 'quantity', 'unit', 'lastPrice'];
            const targetRange = Import.getRange(item.row, 1, 1, 5);

            const copied = copyValuesToRange(targetRange, copyKeys, databaseItem, item);
            
            if (copied) { 
                result.validItems.push(item); 
            }
        }
        return result;
    }
    catch (e) 
    {
        result.errorMessage = e.errorMessage;
        return result;
    }
}

// copy the specified values from an object to a range, optionally also to another object.
function copyValuesToRange(range, copyKeys, fromItem, toItem = undefined) 
{
    const result = {'copied': false, 'errorMessage': ''};
    try 
    {
        const values = [];

        copyKeys.forEach(key => {
            if (toItem) toItem[key] = fromItem[key];
            values.push(fromItem[key]);
        });

        range.setValues([values]);
        result.copied = true;
        return result;
    } 
    catch (e) 
    {
        result.copied       = false; 
        result.errorMessage = e.errorMessage;
        return result;
    }
}

function blah () {
      //   if (valid[i - 3] === "ADD") {
    //     const r = appendRow_DB + result.addedRows - 3;                             // next empty row in DB
        
    //     const row_Item_IMPORT  = sheet_IMPORT.getRange(i,1,  1,4);                 // first 4 cols of a row: Name, Category, Quantity, Unit 
    //     const row_Item_DB      = sheet_DB.getRange(r,2,  1,4);                     // same cells, but in Persistent DB
        
    //     const col_Price_IMPORT = sheet_IMPORT.getRange("E" + i);                   // Price column
    //     const col_Price_DB     = sheet_DB.getRange("F" + r);    

    //     const col_ID_IMPORT    = sheet_IMPORT.getRange("H" + i);                   // ID column
    //     const col_ID_DB        = sheet_DB.getRange("A" + r);

    //     const col_Date_DB      = sheet_DB.getRange("G" + r);                       // Purchased Date column

    //     row_Item_IMPORT.copyTo(row_Item_DB);

    //     col_Price_IMPORT.copyTo(col_Price_DB, {contentsOnly:true});

    //     col_ID_IMPORT.copyTo(col_ID_DB,{contentsOnly:true});

    //     col_Date_DB.setValue(TODAY);                                               // Add date

    //     result.addedRows++;                                                        // Only increment on success       
}

function formatPastedRows() 
{
    const { Import } = initializeSheets();
    const result = {'processedData': [], 'errorMessage': ''};
    try {

        const purchasedLoc  = Import.col('IMPORT_Source')[0];

        switch (purchasedLoc.toString().toUpperCase()) {
            case 'COSTCO': 
                Object.assign(result, formatCostcoReceiptData());
                break;

            default:
                result.errorMessage = "Invalid purchase location! \"" + purchasedLoc + "\"";
                break;
        }

        return result;
    }
    catch (e) 
    {
        result.processedData = [];
        result.errorMessage  = e.errorMessage;

        return result;
    }
}



function lookupIDsInPersistentDatabase() {
    const result = {'successfulMatches': 0, 'errorMessage': ''};

    const lastRow_IMPORT = getLastRow(sheet_IMPORT, ID_COLUMN_IMPORT);             // Find the last row of ID column
    
    const col_Validation_IMPORT = sheet_IMPORT.getRange('F3:F' + lastRow_IMPORT);  // Check if item ID already exists in persistent database...
    col_Validation_IMPORT.setFormula(VALFXNSTR);

    const valid = [].concat(...col_Validation_IMPORT.getValues());                 // Create new 1D array

    for (let i = 0; i < valid.length; i++) {
        const inputRange = sheet_IMPORT.getRange((i+1),1, 1,4)
        
        if (valid[i] !== 'ADD') { inputRange.clear(); continue; }

        
    }
    
    return result;
}

function main() {
    const result = appendMissingImportedPurchases();

    if (result.addedRows > 0) {
        return SpreadsheetApp.getActive().toast("Added " + result.addedRows + " new item IDs.", "Append to Database");
    }

    return SpreadsheetApp.getActive().toast('Unable to update Database:' + result.errorMessage, "Append Failed");
}
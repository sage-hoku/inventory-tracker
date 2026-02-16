const {
    Database  = {'name': 'Persistent Database', 'headerRows': 1},
    Groceries = {'name': 'Grocery List',        'headerRows': 0},
    Import    = {'name': 'Costco Import',       'headerRows': 2},
    Inventory = {'name': 'Apartment Inventory', 'headerRows': 2},
    Meta      = {'name': 'Metadata',            'headerRows': 1},
    Recipes   = {'name': 'Recipes',             'headerRows': 0},
} = SHEETS_BY_NAME;

// Validation string used for checking IDs against persistent database
// ID if valid, False if not found, 
const VAL_STR = 
`=IF(
    AND(
        INDEX(
          IMPORT_Item_ID,   ROW() - ${Import.headerRows}
        )<>"",
        INDEX(
          IMPORT_Item_Label, ROW() - ${Import.headerRows}
        )<>"",
        INDEX(
          IMPORT_Price,     ROW() - ${Import.headerRows}
        )<>"",
    ), 
    IFERROR(
        MATCH(
            INDEX(
              IMPORT_Item_ID, ROW() - ${Import.headerRows}
            ), 
            DB_Item_ID, 0
        ),
        FALSE
    )
)`;

// named ranges
const PURCHLOC = 'IMPORT_Source';

const TODAY = new Date().toJSON().slice(0, 10);

function appendMissingImportedPurchases() 
{
    const result = {'newItems': [], 'validItems': [], 'errorMessage': ''};

    try 
    { 
        const sheets = {
        'Import':   new SheetHelper('Costco Import'),
        'Database': new SheetHelper('Persistent Database')
        };
        
        const validationColumn = sheets['Import'].namedRanges['IMPORT_Valid'];                  // Check if item ID already exists in persistent database...
        validationColumn.setFormula(VAL_STR);                                   // ...using the formula above to ID-match.
        
        const OHI = sheets['Import'].getHeaderRows();
        const OHDR = sheets['Database'].getHeaderRows();

        const newData = sheets['Import'].getData().getValues();
        const oldData = sheets['Database'].getData().getValues();
        
        const databaseIDs = sheets['Database'].namedRanges['DB_Item_ID'];
        const valid = flatten(validationColumn.getRange());

        for (var i = 0; i < newData.length; i++)                      // Copy all validated rows, 1 row at a time
        {
            sheets['Import'].clearCells( i + (OHI + 1), 1,  1, 5 );   // clear input.

            // valid = match found in DB; !valid = no match found (new item)  
            if (!valid[i]) 
            { 
                result.newItems.push({
                    'itemID': sheets['Import'].namedRanges['IMPORT_Item_ID'][i],
                    'row': i + (OHI + 1) 
                });
                continue; 
            }

            const item = { 
                'name':      newData[i][sheets['Import'].namedRanges['IMPORT_Item_Name'].getColumn()],
                'category':  newData[i][sheets['Import'].namedRanges['IMPORT_Category'].getColumn()],
                'quantity':  newData[i][sheets['Import'].namedRanges['IMPORT_Quantity'].getColumn()],
                'unit':      newData[i][sheets['Import'].namedRanges['IMPORT_Unit'].getColumn()],
                'lastPrice': newData[i][sheets['Import'].namedRanges['IMPORT_Last_Price'].getColumn()],
                'valid':     newData[i][sheets['Import'].namedRanges['IMPORT_Valid'].getColumn()],
                'id':        newData[i][sheets['Import'].namedRanges['IMPORT_Item_ID'].getColumn()],
                'label':     newData[i][sheets['Import'].namedRanges['IMPORT_Item_Label'].getColumn()],
                'purchday':  newData[i][sheets['Import'].namedRanges['IMPORT_Date_Purchased'].getColumn()],
                'purchloc':  newData[i][sheets['Import'].namedRanges['IMPORT_Source'].getColumn()],
                'row':       i + (OHI + 1),
                'match':     undefined
            };
            
            const row = databaseIDs.createTextFinder(item.id).findNext().getRow();

            const databaseItem = {
                'category':  oldData[row][sheets['Database'].namedRanges['DB_Category'].getColumn()],
                'quantity':  oldData[row][sheets['Database'].namedRanges['DB_Quantity'].getColumn()],
                'unit':      oldData[row][sheets['Database'].namedRanges['DB_Unit'].getColumn()],
                'name':      oldData[row][sheets['Database'].namedRanges['DB_Item_Name'].getColumn()],
                'lastPrice': oldData[row][sheets['Database'].namedRanges['DB_Price'].getColumn()],
                'id':        oldData[row][sheets['Database'].namedRanges['DB_Item_ID'].getColumn()],
                'label':     oldData[row][sheets['Database'].namedRanges['DB_Item_Abbrv'].getColumn()],
                'purchday':  oldData[row][sheets['Database'].namedRanges['DB_Last_Purchased'].getColumn()],
                'purchloc':  oldData[row][sheets['Database'].namedRanges['DB_Location_Purchased'].getColumn()],
                'row':       row,
                'match':     item
            }

            item.match = databaseItem;
            
            // copy database item into working area
            const copyKeys = ['name', 'category', 'quantity', 'unit', 'lastPrice'];
            const targetRange = sheets['Import'].getRange(item.row, 1, 1, 5);

            const copied = copyValuesToRange(targetRange, copyKeys, databaseItem, item);
            
            result.validItems.push(item);
        }
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
    const result = {'processedData': [], 'errorMessage': ''};
    try {
        const sheet = new SheetHelper('Costco Import');

        const purchasedLoc  = sheet.namedRanges[PURCHLOC].getValues()[0][0];

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
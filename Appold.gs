// Validation string used for checking IDs against persistent database
const VALFXNSTR = 
`=IF(
  AND( 
    G3="E",H3<>"",
    NOT(XLOOKUP(
      H3,
      DB_Item_ID,
      DB_Item_ID,
      FALSE
    ))
  ),
  "ADD",
  "IGNORE"
)`;

const SHEET_IMPORT = importSheet('Costco Import');                             // Open sheets for importing new data and storing
const SHEET_DB =     importSheet('Persistent Database');

const ID_COLUMN_IMPORT = 8;
const ID_COLUMN_DB = 1;

const HEADER_ROWS_IMPORT = 2;
const HEADER_ROWS_DB = 2;

const TODAY = new Date().toJSON().slice(0, 10);


function appendMissingImportedPurchases() 
{
  const result = {'addedRows': 0, 'errorMessage': ''};

  try 
  {  
    const sheet_IMPORT = importSheet('Costco Import');                             // Open sheets for importing new data and storing
    const sheet_DB =     importSheet('Persistent Database');

    const lastRow_IMPORT = getLastRow(sheet_IMPORT, ID_COLUMN_IMPORT);             // Find the last row of ID column
   
    const col_Validation_IMPORT = sheet_IMPORT.getRange('F3:F' + lastRow_IMPORT);  // Check if item ID already exists in persistent database...
    col_Validation_IMPORT.setFormula(VALFXNSTR);                                   // ...using the formula above to ID-match.
    
    const valid = [].concat(...col_Validation_IMPORT.getValues());                 // Create new 1D array
  
    var appendRow_DB = sheet_DB.getLastRow() + 1;                                  // Find the last row for the first 7 cols, use as pointer
    // Logger.log("Last row: " + appendRow_DB);                                    // debug

    
    for (i = HEADER_ROWS_IMPORT + 1; i < lastRow_IMPORT; i++)                      // Copy all validated rows, 1 row at a time
    {                                         
      if (valid[i - 3] === "ADD") {
        const r = appendRow_DB + result.addedRows - 3;                             // next empty row in DB
        
        const row_Item_IMPORT  = sheet_IMPORT.getRange(i,1,  1,4);                 // first 4 cols of a row: Name, Category, Quantity, Unit 
        const row_Item_DB      = sheet_DB.getRange(r,2,  1,4);                     // same cells, but in Persistent DB
        
        const col_Price_IMPORT = sheet_IMPORT.getRange("E" + i);                   // Price column
        const col_Price_DB     = sheet_DB.getRange("F" + r);    

        const col_ID_IMPORT    = sheet_IMPORT.getRange("H" + i);                   // ID column
        const col_ID_DB        = sheet_DB.getRange("A" + r);

        const col_Date_DB      = sheet_DB.getRange("G" + r);                       // Purchased Date column

        row_Item_IMPORT.copyTo(row_Item_DB);

        col_Price_IMPORT.copyTo(col_Price_DB, {contentsOnly:true});

        col_ID_IMPORT.copyTo(col_ID_DB,{contentsOnly:true});

        col_Date_DB.setValue(TODAY);                                               // Add date

        result.addedRows++;                                                        // Only increment on success       
      }
    }
  }
  
  catch (e) {
    return {'addedRows': 0, 'errorMessage': e.errorMessage}
  }

  return result;

}

const SHEET_COSTCO = 'Costco Import';

function formatPastedRows() {
  const result = {'successfulDeletions': 0, 'itemIDs': 0, 'errorMessage': '' };
  const sheet = new SheetHelper('Costco Import');

  const 
  // try {
    
  // }
  
  // catch (e) {
  //   result.successfulDeletions = [];
  //   result.itemIDs = [];
  //   result.errorMessage = e.errorMessage;
  // }

  // finally {
  //   for (var r = SHEET_IMPORT.getLastRow(); r > HEADER_ROWS_IMPORT; r++) {

  //   }

  //   return result;
  // }
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

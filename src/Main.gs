// #region "Setup"
const _cache = {};

const SHEET_IMPORT = 'Receipt Import';

function mapRegionBoundaries(regions) {
    const map = {};

    for (let i = 0; i < regions.length; i++) {
        map[regions[i]] = [i, regions[i+1] ?? regions[i]];
    }
    return map;
}

function getSheetDataCached(sheetName) {
  if (!_cache[sheetName]) _cache[sheetName] = getSheetData(sheetName);
  return _cache[sheetName];
}

function getSheetData(sheetName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const frozen = sheet.getFrozenRows();
    const data = sheet.getDataRange().getValues();

    const headers = data[frozen - 1];

    const columns = {}
    headers.forEach((name, i) => columns[name] = i);

    const rows = data.slice(frozen).map((row, i) => ({
        _rowIndex: (i + 1) + frozen, // actual sheet row number (1-based + header offset)
        _raw: row,
        ...Object.fromEntries(headers.map((h, i) => [h, row[i]]))
    }));

    return { sheet, columns, rows, headers, regions: mapRegionBoundaries(data.slice(0, frozen)), numHeaderRows: frozen };
}
// #endregion

// #region "Search"
function findRows(sheetName, columnName, value) {
  const { rows } = getSheetData(sheetName);
  return rows.filter(row => row[columnName] === value);
}
// #endregion

// #region "Delete"
function deleteRows(sheetName, columnName, value) {
  const { sheet, rows } = getSheetData(sheetName);
  
  // Collect row indices to delete (iterate in reverse to avoid index shifting)
  const toDelete = rows
    .filter(row => row[columnName] === value)
    .map(row => row._rowIndex)
    .sort((a, b) => b - a); // reverse order

  toDelete.forEach(rowIndex => sheet.deleteRow(rowIndex));
}
// #endregion

// #region "Copy" 
// Usage: copyRows('Import', 'Database', 'Item ID', '123456789')

function copyRows(sourceSheetName, destSheetName, columnName, value) {
  const { rows, headers } = getSheetData(sourceSheetName);
  const destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(destSheetName);
  
  const matchingRaws = rows
    .filter(row => row[columnName] === value)
    .map(row => row._raw);

  if (matchingRaws.length > 0) {
    destSheet.getRange(
      destSheet.getLastRow() + 1, 1,
      matchingRaws.length, headers.length
    ).setValues(matchingRaws);
  }
}

// #endregion

// #region Notify
/**
 * Displays a notification card (for Google Workspace Add-ons).
 * @param {string} text - Text to display
 * @returns {Card} A card with the notification
 */
function notify(text, title, duration = 5) {
    Logger.log(text);
    return SpreadsheetApp.getActiveSpreadsheet().toast(text, title, duration);
}

function alert(text) {
    return SpreadsheetApp.getUi().alert(text);
}
// #endregion

// #region "Main"

// #region "Formatting"
function formatPastedReceipt() {
    const result = { message: '', method: '', matches: 0, invalidLines: 0};
    try {
        const Import = getSheetDataCached(SHEET_IMPORT);
        if (!Import.rows[0]['READY']) {
            result.message = 'Validation function has already run.';
            alert('Please input data to the left.');
            throw new Error(result.message);
        }
        result.method = Import.rows[0]['FROM'].toString().toUpperCase();
        if (!result.method) {
            result.message = 'Missing purchase location under column "FROM"';
            alert('Please enter the location of purchase in cell M3.')
            throw new Error(result.message);
        }

        switch (result.method) {
            case 'COSTCO': 
                Object.assign(result, formatCostcoReceiptData(Import));
                break;

            default:
                result.message = `Invalid purchase location: "${result.method}"`;
                break;
        }
        notify(`Matches: ${result.matches}, Ignored: ${result.invalidLines}`,'Formatting Complete');
        return result;
    }
    catch (e) {
      result.message = e.message + `\nformatPastedReceipt: error parsing using method ${ (result.method === '') ? "<none>" : result.method }`;
      alert(result.message);
      throw new Error(result.message);
    }
}

// #endregion

function main() {
  formatPastedReceipt();

}

// #endregion
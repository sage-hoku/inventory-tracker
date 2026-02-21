const OUTPUT_HEADERS = ['ID', 'NAME', 'PRICE', 'PURCHASED', 'FROM'];

const REGEX_EY        = /^E\s+(\d+)\s+(.*?)\s+([\d.]+)\s+Y?\s*$/;
const REGEX_Y         = /^(\d+)\s+(.*?)\s+([\d.]+)\s+Y?\s*$/;
const REGEX_PRICE     = /^(.+?)\s+([\d.]+)\s+Y?\s*$/;
const REGEX_PARTIAL   = /^E\s+(\d+)\s+(.+)$|^(\d+)\s+(.+)$/;
const REGEX_ID        = /^\d+\s+|^E\s+\d+\s+/;

const TODAY = new Date().toISOString().slice(0,10);

// const REGEX_DISCOUNT  = /^\d+\s*\/\s*\d+\s+[\d.]+-?\s*$/;
// const REGEX_FEE       = /^\d+[\s\w]*\/\s*\d+\s+[\d.]+\s*$/;
// const REGEX_LEADING_E = /^E\s+/;


// #region "Ranges"
function toRanges(arr) {
  if (!arr.length) return [];

  const result = [];
  let start = arr[0];
  let len = 1;

  for (let i = 1; i < arr.length; i++) {
    if (arr[i] === arr[i - 1] + 1) {
      len++;
    } else {
      result.push([start, len]);
      start = arr[i];
      len = 1;
    }
  }
  result.push([start, len]);

  return result;
}
// #endregion

// collect continuations
function parseMultiLine(sheet, i, emptyThreshold) {
    const result = { name: '', price: '', lines: 0 };
    try {
        let consecutiveEmpty = 0;
        let j = i + 1;
        while (j < sheet.rows.length && consecutiveEmpty < emptyThreshold) {
            const line = sheet.rows[j]['ID'].toString().trim();
            if (!line) { 
                consecutiveEmpty++; j++;
                continue;
            }
            const partialMatch = parseContinuationLine(line);
            if (!partialMatch) {
                result.lines = j - i;
                return result;
            }
            if (partialMatch.name && partialMatch.price) {
                result.name += partialMatch.name;
                result.price = partialMatch.price;
            }
            j++; consecutiveEmpty = 0;
        }
        return result;
    }
    catch (e) {
        result.message = e.message + "\n" + "Failure on line ${j - sheet.numHeaderRows}";
        return result;
    }
}

function formatCostcoReceiptData(sheet = getSheetDataCached(SHEET_IMPORT)) {
    const result = {message: '', matches: 0, invalidLines: 0};
    try {
        const Import = sheet;
        if (!Import.rows[0]['READY']) {
            result.message = 'Please input data to the left.'
            notify(result.message);
            return result;
        }
        const items        = []; // checkbox values
        const matches      = []; // actual item values
        const invalidLines = [];
        const purchDate = Import.rows[0]['PURCHASED'] ?? TODAY;
        const source    = Import.rows[0]['FROM']      ?? "Unknown";

        let i = 0;
        const emptyThreshold = 2; // number of empty rows before loop exits
        let consecutiveEmpty = 0;
        while (i < Import.rows.length && consecutiveEmpty < emptyThreshold) {
            let line = Import.rows[i]['ID'].toString().trim();
            
            if (!line) { // skip empty lines
                consecutiveEmpty++; i++;
                continue; 
            } 

            // try to match E NAME PRICE Y
            const match = parseLine(line) 
            if (match) { 
                if (match.price) {
                    items.push([true]);
                    matches.push([match.id, match.name, match.price, purchDate, source]);
                } 
                
                else { // Price missing, check for multiple lines
                    const ml = parseMultiLine(Import, i, emptyThreshold);
                    if (ml.message) {
                        throw new Error(ml.message);
                    }
                    if (ml.price) {
                        match.name += " " + ml.name;
                        match.price = ml.price;

                        items.push([true]);
                        matches.push([match.id, match.name, match.price, purchDate, source]);
                    } else {
                      invalidLines.push([line,'','','','']);
                    }
                    i += ml.lines;
                }
            } else {
                invalidLines.push([line,'','','','']);
            }
            consecutiveEmpty = 0; i++;
        }

        for (let b = 0; b < invalidLines.length; b++) {
            items.push([false]); 
            matches.push(invalidLines[b]);
        }

        while (i > items.length) {
          matches.push(['','','','','']);
          i--;
        }

        const checkboxRange = Import.sheet.getRange(
            Import.numHeaderRows + 1, 
            Import.columns['ITEM START'] + 1, 
            items.length
        )
        checkboxRange.insertCheckboxes();
        checkboxRange.uncheck();
        checkboxRange.setValues(items);

        const itemRange = Import.sheet.getRange(
            Import.numHeaderRows + 1, 
            Import.columns['ID'] + 1, 
            matches.length,
            OUTPUT_HEADERS.length
        );
        itemRange.setValues(matches);

        const ready = Import.sheet.getRange(
            Import.numHeaderRows + 1,
            Import.columns['READY'] + 1,
        );

        ready.uncheck();

        result.matches = matches.length;
        result.invalidLines = invalidLines.length;
        return result;
    }
    catch (e) {
        result.message = e.message + "\nformatCostcoReceiptData: error on line ${i}";
        return result;
    }
}

function parseLine(line) {
    // Pattern: ID NAME PRICE Y
    // ID is one or more digits at the start
    const match = line.match(REGEX_EY);
    if (match && match.length === 4) {
        return {
            id:    match[1],
            name:  match[2].trim(),
            price: match[3]
        };
    }

    // Pattern: ID NAME PRICE Y
    const matchNoE = line.match(REGEX_Y);
    if (matchNoE && matchNoE.length === 4) {
        return {
            id:    matchNoE[1],
            name:  matchNoE[2].trim(),
            price: matchNoE[3]
        };
    }

    // If no price found, might be partial line
    const partialMatch = line.match(REGEX_PARTIAL);
    if (partialMatch && partialMatch.length > 1) {
        return {
            id:    partialMatch[1] ?? partialMatch[3],
            name:  (partialMatch[2] ?? partialMatch[4]).trim(),
            price: null
        };
    }

    return null;
}

function parseContinuationLine(line) {
    // End if next item starts.
    const newItemMatch = line.match(REGEX_ID); 
    if (newItemMatch) {
        return null;
    }

    // Continuation line might be: NAME or NAME PRICE Y
    const hasPrice = line.match(REGEX_PRICE);

    if (hasPrice) {
        return {
            name:  hasPrice[1].trim(),
            price: hasPrice[2]
        };
    }

    // Just a name continuation
    return {
        name:  line.trim(),
        price: null
    };
}
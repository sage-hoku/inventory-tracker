const SHEET_COSTCO = 'Costco Import';
const COL_PASTE    = 'IMPORT_Paste'; 

const OUTPUT_HEADERS = ['ID', 'NAME', 'PRICE'];

function formatCostcoReceiptData() 
{
    const result = {'processedData': [], 'errorMessage': ''};
    try 
    {
        const sheet = new SheetHelper(SHEET_COSTCO);

        const pastedData = flatten(sheet.namedRanges[COL_PASTE].getValues());
        const pasteColumnIndex = sheet.namedRanges[COL_PASTE].getColumn();
        
        // Main processing block
        result.processedData = processReceiptLines(pastedData, sheet.getHeaderRows() + 1);

        const CI  = pasteColumnIndex;             // Number of first data column in paste region
        const OH  = OUTPUT_HEADERS.length;        // 3: ID, Name, Price

        const len = result.processedData.length;  // Total item count
        const HR  = sheet.getHeaderRows();        // Lowest header row
        const LR  = sheet.getLastRow();      

        const startRow       =  1 + (HR + len);            
        const numRowsToClear = LR - (HR + len);
        // === Nothing gets changed until here === //

        // Set headers
        sheet.getRange(HR, CI, 1, OH)
            .setValues([OUTPUT_HEADERS]);

        // Write processed data
        if (len <= 0) return false;
        sheet.getRange(HR + 1, CI, len, OH)
            .setValues(result.processedData);

        // Clear empty cells below output
        sheet.getRange(startRow, CI, numRowsToClear, OH).clear();

        // Auto-resize columns
        sheet.autoResizeColumns(CI, OH);

        return result;
    } 
    catch (e) 
    {
        result.processedData = [];
        result.errorMessage  = e.errorMessage;
        return result;
    } 
}
                        
const REGEX_DISCOUNT  = /^\d+\s*\/\s*\d+\s+[\d.]+-?\s*$/;
const REGEX_LEADING_E = /^E\s+/;
const REGEX_ID        = /^\d+\s+/;

// Iterator = i, Lookahead = j. 
// i only advances when j comes across a line with a new ID that is not a discount.
function processReceiptLines(pastedData, offset) 
{
    const results = [];
    let i = offset;

    while (i < pastedData.length) 
    {
        let line = pastedData[i].toString().trim();
        
        // Skip empty lines
        if (!line) { i++; continue; }
        
        // Skip discount/adjustment lines (contain 'ID /ID price-')
        if (line.match(REGEX_DISCOUNT)) { i++; continue; }
        
        // Remove leading 'E ' if present
        line = line.replace(REGEX_LEADING_E, '');
        
        // Parse the line into | ID | NAME | PRICE |
        // If price is not found, then name continues on next line. In this case,
        // fullName is appended to until a price is matched.
        const parsed = parseLine(line);

        // Handle weird cases
        if (!parsed || !parsed.id) { i++; continue; } 

        // Check if item name continues on next line(s)
        let fullName = parsed.name;
        let price    = parsed.price;
        let j = i + 1;
        
        // Look ahead for continuation lines (lines without ID that aren't discount lines)
        while (j < pastedData.length) {
            let nextLine = pastedData[j].toString().trim();
            
            // Skip empty lines
            if (!nextLine) { j++; continue; }
            
            // Remove leading 'E ' if present
            nextLine = nextLine.replace(REGEX_LEADING_E, '');
            
            // Check if it's a discount line
            if (nextLine.match(REGEX_DISCOUNT)) { j++; continue; }
            
            // Check if this line has an ID (starts with digits)
            // If true, nextLine is a new item, stop here
            if (nextLine.match(REGEX_ID)) break;
            
            // This is a continuation line - parse it
            const contParsed = parseContinuationLine(nextLine);
            if (contParsed.name) {
                fullName += ' ' + contParsed.name;
            }
            if (contParsed.price) {
                price = contParsed.price;
            }
            
            j++;
        }
        
        // Add the complete item
        results.push([parsed.id, fullName, price]);
        
        // Skip to the line after all continuation lines
        i = j;
    }

    return results;
}

function parseLine(line) {
    // Pattern: ID NAME PRICE Y
    // ID is one or more digits at the start
    const match = line.match(/^(\d+)\s+(.*?)\s+([\d.]+)\s+Y?\s*$/);

    if (match) {
        return {
            id:    match[1],
            name:  match[2].trim(),
            price: match[3]
        };
    }

    // If no price found, might be partial line
    const partialMatch = line.match(/^(\d+)\s+(.+)$/);
    if (partialMatch) {
        return {
            id:    partialMatch[1],
            name:  partialMatch[2].trim(),
            price: null
        };
    }

    return null;
}

function parseContinuationLine(line) {
    // Continuation line might be: NAME or NAME PRICE Y
    const hasPrice = line.match(/^(.+?)\s+([\d.]+)\s+Y?\s*$/);

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
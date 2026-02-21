/**
 * Test suite for Google Apps Script functions.
 */

/**
 * Tests the findValidItems function and displays results.
 * @returns {Card} Notification card with test results
 */
function test_findValidItems() {
    try {
        const result = findValidItems();
        const { newItems, validItems, errorMessage } = result;
        
        let message = '';
        
        if (errorMessage) {
            message = `Error: ${errorMessage}`;
        } else if (newItems.length > 0 || validItems.length > 0) {
            message = `✓ ${validItems.length} valid items found\n✓ ${newItems.length} new items to add`;
        } else {
            message = 'No items found or added.';
        }
        
        Logger.log(message);
        Logger.log('New Items:', newItems);
        Logger.log('Valid Items:', validItems);
        
        return notify(message);
    } catch (e) {
        const errorMsg = `Test failed: ${e.message || e.toString()}`;
        Logger.log(errorMsg);
        return notify(errorMsg);
    }
}

/**
 * Tests the formatPastedRows function.
 * @returns {Card} Notification card with test results
 */
function test_formatPastedRows() {
    try {
        const result = formatPastedRows();
        const { processedData, errorMessage } = result;
        
        let message = '';
        
        if (errorMessage) {
            message = `Error: ${errorMessage}`;
        } else {
            message = `✓ Processed ${processedData.length} items`;
        }
        
        Logger.log(message);
        Logger.log('Processed Data:', processedData);
        
        return notify(message);
    } catch (e) {
        const errorMsg = `Test failed: ${e.message || e.toString()}`;
        Logger.log(errorMsg);
        return notify(errorMsg);
    }
}

/**
 * Tests the appendToDB function.
 * @returns {Card} Notification card with test results
 */
function test_appendToDB() {
    try {
        const result = appendToDB();
        const { addedRows, updatedRows, priceChanges, errorMessage } = result;
        
        let message = '';
        
        if (errorMessage) {
            message = `Error: ${errorMessage}`;
        } else {
            message = `✓ Added: ${addedRows.length}\n✓ Updated: ${updatedRows.length}\n✓ Price changes: ${priceChanges.length}`;
        }
        
        Logger.log(message);
        Logger.log('Full Result:', result);
        
        return notify(message);
    } catch (e) {
        const errorMsg = `Test failed: ${e.message || e.toString()}`;
        Logger.log(errorMsg);
        return notify(errorMsg);
    }
}

/**
 * Tests the full workflow from format to database append.
 * @returns {Card} Notification card with test results
 */
function test_fullWorkflow() {
    try {
        // Step 1: Format pasted data
        Logger.log('Step 1: Formatting pasted data...');
        const formatResult = formatPastedRows();
        
        if (formatResult.errorMessage) {
            return notify(`Format failed: ${formatResult.errorMessage}`);
        }
        
        Logger.log(`Formatted ${formatResult.processedData.length} items`);
        
        // Step 2: Find valid items
        Logger.log('Step 2: Finding valid items...');
        const validationResult = findValidItems();
        
        if (validationResult.errorMessage) {
            return notify(`Validation failed: ${validationResult.errorMessage}`);
        }
        
        Logger.log(`Found ${validationResult.validItems.length} valid items, ${validationResult.newItems.length} new items`);
        
        // Step 3: Append to database
        Logger.log('Step 3: Appending to database...');
        const appendResult = appendToDB();
        
        if (appendResult.errorMessage) {
            return notify(`Append failed: ${appendResult.errorMessage}`);
        }
        
        const message = `✓ Workflow complete!\nAdded: ${appendResult.addedRows}\nUpdated: ${appendResult.updatedRows}\nPrice changes: ${appendResult.priceChanges}`;
        
        Logger.log(message);
        return notify(message);
    } catch (e) {
        const errorMsg = `Test failed: ${e.message || e.toString()}\n${e.stack || ''}`;
        Logger.log(errorMsg);
        return notify(errorMsg);
    }
}

/**
 * Tests SheetHelper initialization.
 * @returns {Card} Notification card with test results
 */
function test_sheetHelperInit() {
    try {
        const sheets = initializeSheets();
        
        let message = '✓ SheetHelper instances created:\n';
        for (const [name, helper] of Object.entries(sheets)) {
            message += `  - ${name}: ${helper.getName()}\n`;
        }
        
        Logger.log(message);
        return notify(message);
    } catch (e) {
        const errorMsg = `Test failed: ${e.message || e.toString()}`;
        Logger.log(errorMsg);
        return notify(errorMsg);
    }
}

/**
 * Runs all tests sequentially.
 * @returns {void} Logs all test results
 */
function runAllTests() {
    Logger.log('=== Running All Tests ===\n');
    
    Logger.log('--- Test 1: SheetHelper Initialization ---');
    test_sheetHelperInit();
    
    Logger.log('\n--- Test 2: Format Pasted Rows ---');
    test_formatPastedRows();
    
    Logger.log('\n--- Test 3: Find Valid Items ---');
    test_findValidItems();
    
    Logger.log('\n--- Test 4: Append to Database ---');
    test_appendToDB();
    
    Logger.log('\n=== All Tests Complete ===');
}

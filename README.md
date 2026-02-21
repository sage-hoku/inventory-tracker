# Google Sheets Price Tracker - Cleaned Up Scripts

## Changes Made

### 1. **Added Comprehensive Docblocks**
   - Every function now has JSDoc comments explaining:
     - Purpose
     - Parameters with types
     - Return values
     - Throws/errors where applicable

### 2. **Fixed Bugs**
   - Fixed `col()` method variable name error (`col` → `target`)
   - Fixed `instanceof` checks to use `typeof` for primitives
   - Improved error handling with proper error messages
   - Fixed array flattening in `deleteRowIf`

### 3. **New Function: `appendToDB()`**
   This function:
   - Adds new items to the Persistent Database
   - Updates existing items (fills in empty fields)
   - Updates "Last Purchased" date and location
   - Tracks price changes in the Price Tracker sheet
   - Returns detailed results (addedRows, updatedRows, priceChanges)

### 4. **Price Tracker Sheet**
   New sheet to track price history over time.

### 5. **Working Test Suite**
   - `test_findValidItems()` - Tests item validation
   - `test_formatPastedRows()` - Tests Costco receipt parsing
   - `test_appendToDB()` - Tests database append
   - `test_fullWorkflow()` - Tests entire workflow
   - `test_sheetHelperInit()` - Tests sheet initialization
   - `runAllTests()` - Runs all tests sequentially

---

## Price Tracker Sheet Setup

### Sheet Name
`Price Tracker`

### Column Structure

| Column | Header | Description |
|--------|--------|-------------|
| A | Item ID | The unique item identifier |
| B | Initial Price | The first recorded price |
| C | Initial Date | Date of first price |
| D | Price 2 | Second recorded price (when changed) |
| E | Date 2 | Date of second price |
| F | Price 3 | Third recorded price (when changed) |
| G | Date 3 | Date of third price |
| ... | ... | Continue pattern as needed |

### Example Data

```
Item ID | Initial Price | Initial Date | Price 2 | Date 2     | Price 3 | Date 3
--------|---------------|--------------|---------|------------|---------|------------
1234567 | 12.99        | 2024-01-15   | 13.49   | 2024-02-01 | 12.99   | 2024-03-01
2345678 | 8.99         | 2024-01-20   |         |            |         |
3456789 | 25.00        | 2024-01-22   | 24.50   | 2024-02-10 |         |
```

### Named Ranges (Optional but Recommended)

Create these named ranges in the Price Tracker sheet:
- `PT_Item_ID` → Column A (starting after headers)
- `PT_Initial_Price` → Column B
- `PT_Initial_Date` → Column C

### Freeze Header Row
1. Select row 1
2. View → Freeze → 1 row

---

## How It Works

### Price Change Detection

1. When `appendToDB()` runs, it compares:
   - **Old price**: From the Persistent Database
   - **New price**: From the Import sheet

2. If prices differ:
   - Updates the database with the new price
   - Finds the item in Price Tracker
   - Appends new price and date in the next available columns

3. First time seeing an item:
   - Adds to Price Tracker with initial price and date

### Example Flow

**Scenario: Milk price increases**

1. Import sheet has: Milk (ID: 1234567) at $4.29
2. Database shows: Milk last cost $3.99
3. Price Tracker before:
   ```
   1234567 | 3.99 | 2024-01-15 |      |      |
   ```
4. After `appendToDB()`:
   - Database updated to $4.29
   - Price Tracker updated:
   ```
   1234567 | 3.99 | 2024-01-15 | 4.29 | 2024-02-16 |
   ```

---

## Usage Instructions

### 1. Initial Setup
```javascript
// Run once to verify everything works
test_sheetHelperInit()
```

### 2. Import Costco Receipt
1. Paste receipt data into Import sheet (IMPORT_Paste column)
2. Run:
```javascript
formatPastedRows()
```

### 3. Process and Add to Database
```javascript
// Run the full workflow
main()
```

### 4. Testing Individual Components
```javascript
// Test receipt formatting
test_formatPastedRows()

// Test item validation
test_findValidItems()

// Test database append
test_appendToDB()

// Test everything
test_fullWorkflow()
```

---

## File Structure

```
├── SheetHelper.gs      - SheetHelper class and utility functions
├── Main.gs            - Main workflow functions
├── CostcoParser.gs    - Costco receipt parsing logic
└── Tests.gs           - Test suite
```

---

## Common Issues & Solutions

### Issue: Reference Error when calling functions
**Solution**: Make sure all `.gs` files are added to your Google Apps Script project. They share the same global scope.

### Issue: Named ranges not found
**Solution**: Check that all named ranges are properly defined in your sheets:
- IMPORT_Item_ID, IMPORT_Item_Label, etc. in Import sheet
- DB_Item_ID, DB_Category, etc. in Database sheet
- PT_Item_ID in Price Tracker sheet (optional)

### Issue: Price changes not being tracked
**Solution**: 
1. Verify Price Tracker sheet exists with correct name
2. Check that prices are numbers, not text
3. Run `test_appendToDB()` to see detailed error messages

### Issue: Test functions not working
**Solution**: The old test had initialization issues. Use the new test functions which properly handle result objects.

---

## Next Steps

### Recommended Enhancements

1. **Add Data Validation**
   - Ensure prices are positive numbers
   - Validate date formats

2. **Add Price Change Alerts**
   ```javascript
   function checkSignificantPriceChanges() {
     // Alert if price increases > 10%
   }
   ```

3. **Create Price History Chart**
   - Use Google Sheets chart feature
   - Plot price over time for items

4. **Export Price Tracker**
   ```javascript
   function exportPriceHistory(itemId) {
     // Export CSV of price changes for specific item
   }
   ```

---

## Questions?

Check the logs using `Ctrl+Enter` in the script editor, or view the Logger with `View → Logs`.

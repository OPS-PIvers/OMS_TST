function fixAllDates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Fix TST Approvals (New) - Column E (5)
  fixColumn(ss, 'TST Approvals (New)', 5);
  
  // 2. Fix TST Usage (New) - Column C (3)
  fixColumn(ss, 'TST Usage (New)', 3);
}

function fixColumn(ss, sheetName, colIndex) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    console.log(`Sheet '${sheetName}' not found. Skipping.`);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // No data
  
  // Get all dates in the column
  const range = sheet.getRange(2, colIndex, lastRow - 1, 1);
  const values = range.getValues();
  
  const newValues = values.map(row => {
    const val = row[0];
    
    // If it's a Date object, we format it to GMT to "undo" the timezone shift
    // e.g. 2023-12-09 18:00:00 CST -> 2023-12-10 00:00:00 UTC
    if (val instanceof Date) {
      return [Utilities.formatDate(val, "GMT", "yyyy-MM-dd")];
    }
    
    // If it's already a string or empty, leave it alone
    return [val];
  });
  
  // Write back the fixed strings
  // setNumberFormat('@') forces the cell to be treated as Plain Text
  range.setNumberFormat('@').setValues(newValues);
  
  console.log(`Fixed ${sheetName} column ${colIndex}.`);
}

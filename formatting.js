/**
 * Apply formatting to all sheets
 */
function applyFormatting(mainSheet, dashboardSheet) {
  // Format currency cells in main sheet
  var currencyRanges = [
    "B4:B9", // Property Information
    "B12", "B15", // Subject-To
    "B17", "B20", // Seller Financing
    "B23:B26", "B28", // Income and Expenses
    "B31", "B33", "B36", "B40:B42", "B44", // Calculated Metrics
    "B48:C52", // Interest Rate Sensitivity
    "B56:C59", // Appreciation Rate Sensitivity
    "B64:D68", // Scenario Comparison
    "C31", "C33", "C36", "C40:C42", "C44" // Added Column C calculated metrics
  ];
  
  for (var i = 0; i < currencyRanges.length; i++) {
    mainSheet.getRange(currencyRanges[i]).setNumberFormat("$#,##0.00");
  }
  
  // Format percentage cells
  var percentageRanges = [
    "B13", "B18", "B27", "B35", "B38", // Input Percentages
    "B37", "B43", "D48:D52", // Calculated Percentages
    "A48:A52", "A56:A59", // Rate Sensitivity
    "C37", "C43" // Added Column C percentage cells
  ];
  
  for (var i = 0; i < percentageRanges.length; i++) {
    mainSheet.getRange(percentageRanges[i]).setNumberFormat("0.00%");
  }
  
  // Format months with 1 decimal place
  mainSheet.getRange("B32").setNumberFormat("0.0");
  mainSheet.getRange("B66:D66").setNumberFormat("0.0");
  mainSheet.getRange("C32").setNumberFormat("0.0"); // Added Column C
  
  // Format DSCR with 2 decimal places
  mainSheet.getRange("B34").setNumberFormat("0.00");
  mainSheet.getRange("B68:D68").setNumberFormat("0.00");
  mainSheet.getRange("C34").setNumberFormat("0.00"); // Added Column C
  
  // Format the column C (same formatting as column B for calculated metrics)
  var columnCRange = mainSheet.getRange("C31:C44");
  columnCRange.setBackground("#e6f2ff");  // Light blue background (same as formula cells)
  
  // Add conditional formatting for cash flow
  addConditionalFormatting(mainSheet);
  
  // Add alternating row colors - avoid overriding user input highlighting
  for (var i = 1; i <= 70; i++) {
    if (i % 2 == 0 && !isHeaderRow(i) && !isUserInputRow(i)) {
      mainSheet.getRange("A" + i + ":C" + i).setBackground("#f3f3f3");
    }
  }
  
  // Highlight formula cells with light blue background
  var formulaRanges = [
    "B15", "B20", // PMT formulas
    "B31:B34", "B36:B37", "B40:B44", // Calculated metrics
    "B48:D52", // Interest rate sensitivity
    "B56:C59", // Appreciation sensitivity
    "B64:D68",  // Scenario comparison
    "C31:C34", "C36:C37", "C40:C44" // Added Column C metrics
  ];
  
  for (var i = 0; i < formulaRanges.length; i++) {
    mainSheet.getRange(formulaRanges[i]).setBackground("#e6f2ff");
  }
}

/**
 * Helper function to check if a row contains user input fields
 */
function isUserInputRow(rowNum) {
  var userInputRows = [4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 17, 18, 19, 21, 23, 24, 25, 26, 27, 28, 29, 35, 38, 39];
  return userInputRows.indexOf(rowNum) !== -1;
}

/**
 * Add conditional formatting to the sheet
 */
function addConditionalFormatting(sheet) {
  var cashFlowRules = sheet.getConditionalFormatRules();
  var cashFlowRanges = ["B31", "B41", "B44", "C48:C52", "B65:D65", "C31", "C41", "C44"];
  
  for (var i = 0; i < cashFlowRanges.length; i++) {
    var range = sheet.getRange(cashFlowRanges[i]);
    var positiveRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .setBackground("#b7e1cd")  // Light green
        .setRanges([range])
        .build();
    var negativeRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setBackground("#f4c7c3")  // Light red
        .setRanges([range])
        .build();
    
    cashFlowRules.push(positiveRule);
    cashFlowRules.push(negativeRule);
  }
  
  // Add conditional formatting for DSCR
  var dscrRanges = ["B34", "B68:D68", "C34"];
  
  for (var i = 0; i < dscrRanges.length; i++) {
    var range = sheet.getRange(dscrRanges[i]);
    var goodRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(1.25)
        .setBackground("#b7e1cd")  // Light green
        .setRanges([range])
        .build();
    var mediumRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(1, 1.25)
        .setBackground("#fce8b2")  // Light yellow
        .setRanges([range])
        .build();
    var badRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(1)
        .setBackground("#f4c7c3")  // Light red
        .setRanges([range])
        .build();
    
    cashFlowRules.push(goodRule);
    cashFlowRules.push(mediumRule);
    cashFlowRules.push(badRule);
  }
  
  sheet.setConditionalFormatRules(cashFlowRules);
}

/**
 * Helper function to check if a row is a header row
 */
function isHeaderRow(rowNum) {
  var headerRows = [1, 2, 3, 11, 22, 30, 46, 54, 62, 63, 70];
  return headerRows.indexOf(rowNum) !== -1;
}
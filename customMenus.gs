function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Advanced Menu')
      .addItem('Clear Master Sheet', 'clearContentsOnly')
      .addToUi();
}

function clearContentsOnly() {
 var range = SpreadsheetApp
               .getActive()
               .getSheetByName("Master")
               .getRange(2, 8, 35, 7); //Set range to row 2, column 8, 35 rows down, and 7 columns across.
 range.clearContent();
 var checkbox = SpreadsheetApp
               .getActive()
               .getSheetByName("Master")
               .getRange(2, 15, 35, 1); //Set range to row 2, column 15, 35 rows down, and 1 column across.
 checkbox.setValue(false);
}

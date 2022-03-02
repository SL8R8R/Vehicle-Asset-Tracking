try {
  ////////////////////////////////
  // [START modifiable parameters]
  var sheetsToWatch =       ['Master'];
  var columnsToWatch =      ['EXPORT'];
  var valuesToWatch =       [/^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i, /^(true)$/i];
  var targetSheets =        ['46','69','179','213','215','217','220','223','234','243','267','285','295','307','308','309','312','321','534','536','552','553','554','656','675','801','852','867','877','888','936','973','984','985','54-D','67']; 
  var targetSpreadheets =   [];
  var targetIdColumn =      ['Vehicle Unit #'];
  var targetValues =        [46,69,179,213,215,217,220,223,234,243,267,285,295,307,308,309,312,321,534,536,552,553,554,656,675,801,852,867,877,888,936,973,984,985,"54-D",67];
  var transferTypes =       ['PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', 'PASTE_NORMAL', ]; //If adding or removing vehicles, be sure to add/remove the same number of entries from this variable
  var allowArrayFormulas =  [true];
  var copyInsteadOfMove =   [true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, ];
  var numColumnsToMove =    [];
  var changeColumnOrderTo = [];
  var sheetsToSort =        [];
  var columnToSortBy =      [];
  var sortAscending =       [];
  // [END modifiable parameters]
  ////////////////////////////////
} catch (error) {
  showAndThrow_(error);
}

/*
 Documentation for the parameters above
   The script moves a row to one of targetSheets when one of valuesToWatch
   appears in any of columnsToWatch.
   The first value in valuesToWatch moves the row to the first sheet listed
   in targetSheets of the first spreadsheet listed in targetSpreadheets.
   The second value in valuesToWatch moves the row to the second sheet
   listed in targetSheets of the second spreadsheet listed in
   targetSpreadheets, etc.
   If targetValues and targetIdColumn are defined, they are used instead
   of valuesToWatch to decide which one of targetSheets to move to row to.
   
 columnsToWatch
   List the names of columns you want the script to monitor. Use exact
   letter case. There can be any number of columnsToWatch. These columns
   will be monitored only on the sheets listed in sheetsToWatch.
   The column names must appear on the last frozen row in the sheet, or if
   there are no frozen rows, on row 1. Note that vertically merged cells do
   not work well on column name rows. Use text wrap instead of cell merges.
   Enclose column names in single quotes and separate them with commas,
   as in ['Approve', 'Reject'].
   
 sheetsToWatch
   List the sheets with the data you want to move when a magic value is
   entered on their row. The script will monitor the columns named in
   columnsToWatch on these sheets, and it will act when any of the
   valuesToWatch appears in any of the columnsToWatch.
   Not all of the columnsToWatch need to appear on all the sheetsToWatch.
   Enclose sheet names in single quotes and separate them with commas,
   as in ['Data entry', 'Approvals', 'Rejects'].
   
 valuesToWatch
   These values trigger the row move. Use regular expressions.
   See these sites for more information on regular expressions:
     http://en.wikipedia.org/wiki/Regular_expression
     https://github.com/google/re2/wiki/Syntax
     https://regexr.com/
     https://regex101.com/
   Enclose regular expressions in forward slashes and use a
   trailing "i" to ignore letter case.
   To match "accepted" and "AcceptED", use:
     /^(accepted)$/i
   To match "accepted" and "completed", use:
     /^(accepted|completed)$/i
   To match "accepted", "AcceptED" and "accepted-URGENT", use:
     /^(accepted)/i
   To match any value that ends with "done", use:
     /(done)$/i
   To match a ticked checkbox, use:
     /^(true)$/i
   To match a blank checkbox, use:
     /^(false)$/i
   To match any date in the mm/dd/yyyy format, use:
     /^(\d\d?\/\d\d?\/\d\d\d\d)$/
   To match any date in 2020 in the yyyy-mm-dd format, use:
     /^(2020-\d\d?-\d\d?)$/
   Each value can only appear once in the valuesToWatch list.
   
 targetSheets
   List the sheets where you want to move rows to. The number of
   targetSheets must be the same as the number of valuesToWatch.
   Each individual value in valuesToWatch will cause the row to be
   moved to the target sheet that is at the same index position in
   its list as the value that triggered the row move. You can repeat
   a sheet name in the list if necessary. Enclose sheet names in
   single quotes and separate them with commas, as in
   ['Approvals', 'Approvals with prejudice', 'Rejects', 'Rejects'].
   
 targetSpreadheets
   Use spreadsheet IDs as shown in the browser address bar between
   /d/ and /edit in exact letter case.
   Leave the targetSpreadheets list empty to move rows within the
   same spreadsheet file. The number of targetSpreadheets must be
   the same as the number of valuesToWatch, or zero.
   
 targetIdColumn
   Use exact letter case. The name of the column with the values
   that tell which target sheet to move the row to.
   The column identified by targetIdColumn would typically contain
   formulas that give a value that tells which sheet to move the
   row to, depending on other values on the same row.
   There can only be one column name in the targetIdColumn list.
   If a targetIdColumn is defined, a column with that name must
   appear on all sheetsToWatch.
   Leave the targetIdColumn list empty to use valuesToWatch as
   guide to select the target sheet instead.
   
 targetValues
   The regular expressions to match with the value in targetIdColumn
   to determine which target sheet to move the row to.
   Each individual value in targetValues will cause the row to be
   moved to the target sheet that is at the same index position
   in its list as the value that triggered the row move.
   Leave the targetValues list empty to use valuesToWatch as guide
   to select the target sheet instead.
 
 transferTypes
   Determines what to do with rows where the edited value in
   columnsToWatch matches valuesToWatch. The number of items in
   transferTypes must be the same as the number of valuesToWatch,
   or zero.
   Each item in transferTypes must be one of these values:
     'PASTE_VALUES': Paste values only, without formats or formulas.
     'PASTE_FORMAT': Paste values and formats without formulas.
     'PASTE_FORMULA': Paste values and formulas without formats.
     'PASTE_NORMAL': Paste values, formulas and formats.
     'PASTE_NO_BORDERS': Like PASTE_NORMAL but without cell borders.
   If the rows contain formulas, and you are using a transferType
   other than 'PASTE_VALUES', the formulas in each row must use
   absolute references like $A$1 unless the referenced cells are
   in the same row as the formula. Relative references to other
   rows will get mangled during the row move when not using
   the transferType 'PASTE_VALUES'.
   Leave the transferTypes list empty to use the default
   transfer type of 'PASTE_VALUES' for all row moves.
 
 allowArrayFormulas
   Use true to clear in the moved row the columns where the column
   header in the target sheet contains an array formula. This
   makes room for the array formula to fill its result throughout
   the column.
 
 copyInsteadOfMove
   Set true to copy the values and retain the original.
   Set false, or leave empty, to move the values and delete the
   original.
   The number of true/false values in copyInsteadOfMove must be
   the same as the number of valuesToWatch, or zero.
   Leave the copyInsteadOfMove list empty to use the default
   behavior of moving the row instead of copying it.
   
 numColumnsToMove
   The number of columns to move or copy. Leave the list empty
   to copy or move all columns.
   This parameter works only when transferType is PASTE_VALUES.
   The number of values in numColumnsToMove must be the same
   as number of valuesToWatch, or zero.
   
 changeColumnOrderTo
   Defines a custom order of values in the target row.
   Leave empty if column order should not be changed.
   This parameter works only when transferType is PASTE_VALUES.
   The list contains the indices of columns in the source sheet.
   The first column in the target row will contain the
   source row column indicated by the first index number, the
   second column will contain the value in the column indicated
   by the second index number, and so on.
   Use -1 to leave a column empty in the target row.
   If an index number slot is left blank, the target row column
   will contain the value in the same column location in the
   source row.
   The index numbers are zero-based, so column A = 0, B = 1,
   C = 2, D = 3, N = 13, etc.
   Here is an example that copies the first five values from
   a source row to a target row, inserting a gap of one
   column between source row columns 2 and 3, and a gap of two
   columns between source row columns 4 and 5:
   var numColumnsToMove = 5;
   var changeColumnOrderTo = [0, 1, -1, 2, 3, -1, -1, 4];
   
 sheetsToSort
   These sheets get sorted when a row is moved from or to that sheet.
   Enclose sheet names in single quotes and separate them with
   commas, as in ['Data entry', 'Approvals', 'Rejects'].
   Leave the list empty if no sheets should be sorted.
   If you are copy-pasting multiple rows, and several rows get moved,
   the only target sheet that gets sorted is the one that is referred
   to by the last pasted row. The source sheet is always sorted if it
   matches sheetsToSort.
   
 columnToSortBy
   The column number of the column to sort a sheet by.
   Each column number corresponds to a sheet in sheetsToSort, one
   column per sheet, as in [1, 3, 1].
   The index numbers are 1-based, so column A = 1, B = 2, C = 3, etc.
   The number of items in columnToSortBy must be the same as the
   number of sheetsToSort.
   
 sortAscending
   Use true to sort in ascending order and false to sort in
   descending order.
   Each value corresponds to a sheet in sheetsToSort, one value per
   sheet, as in [true, false, true].
   The number of items in sortAscending must be the same as the
   number of sheetsToSort.
   
*/

/**
* Moves rows from sheet to sheet, and sorts the source and target sheets after the move.
*
* @param {Object} e The onEdit event object.
*/
function moveRowsAndSortSheet_(e) {
  // version 1.1, written by --Hyde, 27 June 2020
  try {
    var lock = LockService.getDocumentLock();
    lock.waitLock(30 * 1000);
    var modifiedSheets = moveRowsFromSpreadsheetToSpreadsheet_(e);
    if (modifiedSheets) {
      [modifiedSheets.sourceSheet].concat(modifiedSheets.targetSheets).forEach(function (sheet) {
        sortSheet_(sheet);
      });
    }
  } catch (error) {
    showAndThrow_(error);
  } finally {
    lock.releaseLock();
    //Set current date in Date of Last Export column
      if ([8].indexOf(e.range.columnStart) != -1) {
      e.range.offset(0, 6).setValue(new Date());
    }
  }
}

/**
* Moves a row from a spreadsheet to another spreadsheet file when a magic value is entered in a column.
*
* The name of the sheet to move the row to is derived from the position of the magic value on the valuesToWatch list.
* The targetSpreadheets list uses spreadsheet IDs that can be obtained from the address bar of the browser.
* Use a spreadsheet ID of '' to indicate that the row is to be moved to another tab in the same spreadsheet file.
*
* Globals: see the Global variables section.
* Displays pop-up messages through Spreadsheet.toast().
* Throws errors.
*
* @param {Object} e The 'on edit', 'on form submit' or 'on change' event object.
* @return {Object} An object that lists the sheets that were modified with this structure, or a falsy value if no rows were moved:
*                  {Sheet} sourceSheet The sheet from where a row was moved.
*                  {Sheet[]} targetSheets The sheets to where rows were moved.
*                  {Number} numRowsMoved The number of rows that were moved to another sheet.
*/
function moveRowsFromSpreadsheetToSpreadsheet_(e) {
  // version 2.9.1, written by --Hyde, 18 January 2022
  //  - replace |Tools > Script editor| with |Extensions > Apps Script|
  // version 1.0, written by --Hyde, 14 July 2015
  //  - initial version
  if (e.value === '') { // optimization for single-cell edits
    return;
  }
  if (e.value !== undefined) { // optimization for single-cell edits
    var valuesToWatchIndex = -1;
    for (var i = 0, numRegexes = valuesToWatch.length; i < numRegexes; i++) {
      if (e.value.match(valuesToWatch[i])) {
        valuesToWatchIndex = i;
        break;
      }
    }
    if (valuesToWatchIndex === -1) {
      return null;
    }
  }
  var event = getEventObject_(e);
  if (!event || sheetsToWatch.indexOf(event.sheetName) === -1) {
    return;
  }
  if (targetIdColumn.length) {
    var targetIdColumnNumber = event.columnLabels.indexOf(targetIdColumn[0]) + 1;
    if (!targetIdColumnNumber || !targetValues.length) {
      throw new Error('Could not find target values in target column "' + String(targetIdColumn[0]) + '".');
    }
    var valuesInTargetIdColumn = event.sheet.getRange(event.rowStart, targetIdColumnNumber, event.numRows, 1).getDisplayValues();
  }
  var numRowsMoved = 0;
  var messageOnDisplay = false;
  var sourceSheetNames = [];
  var targetSheetNames = [];
  var targets = [];
  for (var row = event.numRows - 1; row >= 0; row--) {
    for (var column = 0; column < event.numColumns; column++) {
      if (event.rowStart + row <= event.columnLabelRow
        || columnsToWatch.indexOf(event.columnLabels[event.columnStart - 1 + column]) === -1) {
        continue;
      }
      valuesToWatchIndex = -1;
      for (var i = 0, numRegexes = valuesToWatch.length; i < numRegexes; i++) {
        if (event.displayValues[row][column].match(valuesToWatch[i])) {
          valuesToWatchIndex = i;
          break;
        }
      }
      if (valuesToWatchIndex === -1) {
        continue;
      }
      var targetIndex = -1;
      if (targetIdColumn.length) {
        for (var i = 0, numRegexes = targetValues.length; i < numRegexes; i++) {
          if (valuesInTargetIdColumn[row][0].match(targetValues[i])) {
            targetIndex = i;
            break;
          }
        }
      } else {
        targetIndex = valuesToWatchIndex;
      }
      if (targetIndex === -1) {
        continue;
      }
      if (!messageOnDisplay) {
        showMessage_('Moving rows...', 30);
        messageOnDisplay = true;
      }
      var targetSheet = getTargetSheet_(event, targetIndex);
      if (!targetSheet) {
        continue; // skip moving the row if it would end up on the same sheet
      }
      var sourceRange = event.sheet.getRange(event.rowStart + row, 1, 1, event.numSheetColumns);
      var transferType = transferTypes[targetIndex];
      var firstFreeTargetRow = targetSheet.getLastRow() + 1;
      if (firstFreeTargetRow > targetSheet.getMaxRows() && transferType && transferType !== 'PASTE_VALUES') {
        targetSheet.insertRowAfter(targetSheet.getLastRow());
      }
      var targetRange = targetSheet.getRange(firstFreeTargetRow, 1);
      switch (transferType) {
        case 'PASTE_VALUES':
        case undefined:
          targetSheet.appendRow(rearrangeRowValues_(sourceRange, targetIndex));
          break;
        case 'PASTE_FORMAT':
          // @see https://developers.google.com/apps-script/reference/spreadsheet/copy-paste-type
          sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType[transferType], false);
          sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
          break;
        case 'PASTE_FORMULA':
        case 'PASTE_NORMAL':
        case 'PASTE_NO_BORDERS':
          sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType[transferType], false);
          break;
        default:
          throw new Error('Unknown transferType "' + transferType + '".');
      }
      // clear cells in targetRange where column label row contains an array formula
      if (allowArrayFormulas[targetIndex]) {
        var numColumns = sourceRange.getWidth();
        var formulas = targetSheet.getRange(targetSheet.getFrozenRows() || 1, 1, 1, numColumns).getFormulas()[0];
        formulas.forEach(function (formula, index) {
          if (formula.match(/^=.*arrayformula/i)) {
            targetRange.offset(0, index, 1, 1).clearContent();
          }
        });
      }
      numRowsMoved += 1;
      if (!copyInsteadOfMove[targetIndex]) {
        if (event.sheet.getMaxRows() <= event.columnLabelRow + 1) { // avoid deleting the last unfrozen row
          event.sheet.appendRow([null]);
        }
        event.sheet.deleteRow(event.rowStart + row);
      }
      sourceSheetNames = sourceSheetNames.concat(event.sheetName).filter(filterUniques_);
      targetSheetNames = targetSheetNames.concat(targetSheet.getName()).filter(filterUniques_);
      targets = targets.concat(targetSheet).filter(filterUniques_);
    } // column
  } // row
  if (messageOnDisplay) {
    var message = 'Moved ' + numRowsMoved + (numRowsMoved === 1 ? ' row ' : ' rows ') + "from '" + sourceSheetNames.join(', ') + "' to '" + targetSheetNames.join(', ') + "'.";
    showMessage_('Moving rows... done. ' + message);
  }
  return numRowsMoved ? { sourceSheet: event.sheet, targetSheets: targets, numRowsMoved: numRowsMoved } : null;
}

/**
* Sorts a sheet.
*
* Globals: sheetsToSort, columnToSortBy, sortAscending.
*
* @param {Sheet} sheet The sheet to sort.
*/
function sortSheet_(sheet) {
  // version 1.0, written by --Hyde, 19 March 2020
  var sheetIndex = sheetsToSort.indexOf(sheet.getName());
  if (sheetIndex === -1) {
    return;
  }
  sheet.sort(columnToSortBy[sheetIndex], sortAscending[sheetIndex]);
}

/**
* Determines the type of a spreadsheet event and populates an event object.
*
* @param {Object} e The original event object.
* @return {Object} An event object with the following fields, or null if the event type is unknown.
*                  {Spreadsheet} spreadsheet The spreadsheet that was edited.
*                  {Sheet} sheet The sheet that was edited in spreadsheet.
*                  {Range} range The cell or range that was edited in sheet.
*                  {String} sheetName The name of the sheet that was edited.
*                  {Number} rowStart The ordinal number of the first row in range.
*                  {Number} rowEnd The ordinal number of the last row in range.
*                  {Number} columnStart The ordinal number of the first column in range.
*                  {Number} columnEnd The ordinal number of the last column in range.
*                  {Number} numRows The number of rows in range.
*                  {Number} numColumns The number of columns in range.
*                  {Number} numSheetColumns The number of columns in sheet.
*                  {Number} columnLabelRow The 1-based row number where column labels are found.
*                  {String[]} columnLabels The values in row event.columnLabelRow as shown in the spreadsheet as text strings.
*                  {String[][]} displayValues The values in the edited range as shown in the spreadsheet as text strings.
*                  {String} eventType One of ON_EDIT, ON_CHANGE or ON_FORM_SUBMIT.
*                  {String} changeType Always EDIT, and never INSERT_ROW, INSERT_COLUMN, REMOVE_ROW, REMOVE_COLUMN, INSERT_GRID, REMOVE_GRID, FORMAT, or OTHER.
*                  {String} authMode One of ScriptApp.AuthMode.NONE, .LIMITED, .FULL or .CUSTOM_FUNCTION.
*/
function getEventObject_(e) {
  // version 1.4, written by --Hyde, 20 January 2021
  if (!e) {
    return null;
  }
  var event = {};
  if (e.range && JSON.stringify(e.range) !== '{}') { // triggered by ScriptApp.EventType.ON_EDIT or .ON_FORM_SUBMIT
    event.range = e.range;
    event.rowStart = Number(e.range.rowStart);
    event.rowEnd = Number(e.range.rowEnd);
    event.columnStart = Number(e.range.columnStart);
    event.columnEnd = Number(e.range.columnEnd);
    event.changeType = 'EDIT';
    event.eventType = e.namedValues ? 'ON_FORM_SUBMIT' : 'ON_EDIT';
  } else if (e.changeType === 'EDIT') { // triggered by ScriptApp.EventType.ON_CHANGE
    // @see https://developers.google.com/apps-script/guides/triggers/events#change
    // @see https://community.glideapps.com/t/new-row-in-spreadsheet-for-every-user/6475/55
    var ss = SpreadsheetApp.getActive();
    event.range = ss.getActiveRange();
    event.rowStart = event.range.getRow();
    event.rowEnd = event.range.getLastRow();
    event.columnStart = event.range.getColumn();
    event.columnEnd = event.range.getLastColumn();
    event.changeType = e.changeType;
    event.eventType = 'ON_CHANGE';
  } else { // triggered by some other change type
    return null;
  }
  event.authMode = e.authMode; // @see https://developers.google.com/apps-script/reference/script/auth-mode
  event.sheet = event.range.getSheet();
  event.sheetName = event.sheet.getName();
  event.spreadsheet = event.sheet.getParent();
  event.numRows = event.rowEnd - event.rowStart + 1;
  event.numColumns = event.columnEnd - event.columnStart + 1;
  event.columnLabelRow = event.sheet.getFrozenRows() || 1;
  event.numSheetColumns = event.sheet.getLastColumn();
  event.columnLabels = event.sheet.getRange(event.columnLabelRow, 1, 1, event.numSheetColumns).getDisplayValues()[0];
  event.displayValues = event.range.getDisplayValues();
  return event;
}

/**
* Callback function to Array.filter() to return an array where there is just one copy of each individual value.
*
* Usage:
*   var array = [3, 1, 1, 1, '1', 2, '1', 'test', 'test2', 'test'];
*   var unique = array.filter(filterUniques_); // returns [3, 1, "1", 2, "test", "test2"]
*
* @see https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/filter
* @see https://stackoverflow.com/questions/9229645/remove-duplicate-values-from-js-array
*
* @param {Object} element The current element being processed in the array.
* @param {Number} index The index of the current element being processed in the array.
* @param {Object[]} array The array filter was called upon.
* @return {Boolean} True if element is unique in array, false if there is another identical value in array.
*/
function filterUniques_(element, index, array) {
  // version 1.0, written by --Hyde, 30 May 2019
  return array.indexOf(element) === index;
}

/**
* Gets the values in range as a 1D array, adjusts the length of the array to
* numColumnsToMove[targetIndex], and rearranges the array by the indices in changeColumnOrderTo.
* 
* Globals: numColumnsToMove, changeColumnOrderTo.
*
* @param {Range} range The range where to get row values.
* @param {Number} targetIndex The index to use when getting the pertinent value from numColumnsToMove.
* @return {Object[]} The values in row, with length adjusted to numColumnsToMove[targetIndex] and values rearranged by the indices in changeColumnOrderTo.
*/
function rearrangeRowValues_(range, targetIndex) {
  // version 1.0, written by --Hyde, 27 June 2020
  var rowValuesInOriginalOrder = range.getValues()[0];
  if (numColumnsToMove[targetIndex] !== undefined) {
    var rowValues = rowValuesInOriginalOrder.slice(0, numColumnsToMove[targetIndex]);
  } else {
    rowValues = rowValuesInOriginalOrder.slice();
  }
  for (var changeIndex = 0; changeIndex < changeColumnOrderTo.length; changeIndex++) {
    if (changeColumnOrderTo[changeIndex] !== undefined) {
      rowValues[changeIndex] = rowValuesInOriginalOrder[changeColumnOrderTo[changeIndex]];
    } else if ((rowValues[changeIndex] == undefined) || rowValues[changeIndex] === '') {
      rowValues[changeIndex] = null;
    }
  }
  return rowValues;
}

/**
* Gets the target sheet where to move a row.
* 
* Globals: targetSheets, targetSpreadheets.
* Throws errors.
*
* @param {Object} event The event object from getEventObject_().
* @param {Number} targetIndex The index to use when getting the pertinent value from targetSheets and targetSpreadheets.
* @return {Sheet} The target sheet, or null if the target sheet is the same as the source sheet.
*/
function getTargetSheet_(event, targetIndex) {
  // version 1.0, written by --Hyde, 27 June 2020
  var targetSheetName = targetSheets[targetIndex];
  var targetSpreadsheetId = String(targetSpreadheets[targetIndex] || '');
  if (targetSpreadsheetId) {
    try {
      var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
    } catch (error) {
      var ssIdShortened
        = targetSpreadsheetId.length > 13
          ? targetSpreadsheetId.slice(0, 5) + '...' + targetSpreadsheetId.slice(-5)
          : targetSpreadsheetId;
      throw new Error('Could not open a target spreadsheet with ID "' + ssIdShortened + '".');
    }
  } else {
    targetSpreadsheet = event.spreadsheet;
    if (targetSheetName === event.sheetName) {
      return null; // skip moving the row if it would end up in the same sheet
    }
  }
  var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);
  if (!targetSheet) {
    throw new Error("Could not find the target sheet '" + targetSheetName + "'"
      + targetSpreadsheet === event.spreadsheet
      ? '.'
      : ' in spreadsheet "' + targetSpreadsheet.getName() + '".');
  }
  return targetSheet;
}

/**
* Installs a trigger that runs each time the user hand edits the spreadsheet.
* Deletes any previous instances of ON_EDIT and ON_CHANGE triggers.
*
* To permanently install the trigger, choose Run > Run function > installOnEditTrigger.
* You only need to install the trigger once per spreadsheet.
* To review the installed triggers, choose Edit > Current project's triggers.
*/
function installOnEditTrigger() {
  // version 1.0, written by --Hyde, 7 May 2020
  deleteTriggers_(ScriptApp.EventType.ON_EDIT);
  deleteTriggers_(ScriptApp.EventType.ON_CHANGE);
  ScriptApp.newTrigger('moveRowsAndSortSheet_')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

/**
* Installs a trigger that runs each time a form is submitted.
* Deletes any previous instances of ON_FORM_SUBMIT triggers.
*
* To permanently install the trigger, choose Run > Run function > installOnFormSubmitTrigger.
* You only need to install the trigger once per spreadsheet.
* To review the installed triggers, choose Edit > Current project's triggers.
*/
function installOnFormSubmitTrigger() {
  // version 1.0, written by --Hyde, 7 May 2020
  deleteTriggers_(ScriptApp.EventType.ON_FORM_SUBMIT);
  ScriptApp.newTrigger('moveRowsAndSortSheet_')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create();
}

/**
* Installs a trigger that runs each time a change is done by Glide, IFTTT, Zapier or other such tools.
* Deletes any previous instances of ON_EDIT and ON_CHANGE triggers.
* [NOTE] This trigger will also move rows when you hand edit the spreadsheet.
* [NOTE] Tested with Glide, but untested with IFTTT, Zapier.
*
* To permanently install the trigger, choose Run > Run function > installOnChangeTrigger.
* You only need to install the trigger once per spreadsheet.
* To review the installed triggers, choose Edit > Current project's triggers.
*/
function installOnChangeTrigger() {
  // version 1.0, written by --Hyde, 7 May 2020
  deleteTriggers_(ScriptApp.EventType.ON_EDIT);
  deleteTriggers_(ScriptApp.EventType.ON_CHANGE);
  ScriptApp.newTrigger('moveRowsAndSortSheet_')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();
}

/**
* Deletes all installable triggers of the type triggerType associated with the current
* script project that are owned by the current user in the current spreadsheet.
*
* @param {EventType} triggerType One of ScriptApp.EventType.ON_EDIT, .ON_FORM_SUBMIT, .ON_OPEN, .ON_CHANGE, .CLOCK (time-driven triggers) or .ON_EVENT_UPDATED (Calendar events).
*/
function deleteTriggers_(triggerType) {
  // version 1.1, written by --Hyde, 27 June 2020
  var triggers = ScriptApp.getUserTriggers(SpreadsheetApp.getActive());
  for (var i = 0, numTriggers = triggers.length; i < numTriggers; i++) {
    if (triggers[i].getEventType() === triggerType) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
* Shows error.message in a pop-up and throws the error.
*
* @param {Error} error The error to show and throw.
*/
function showAndThrow_(error) {
  // version 1.0, written by --Hyde, 16 April 2020
  var stackCodeLines = String(error.stack).match(/\d+:/);
  if (stackCodeLines) {
    var codeLine = stackCodeLines.join(', ').slice(0, -1);
  } else {
    codeLine = error.stack;
  }
  showMessage_(error.message + ' Code line: ' + codeLine, 30);
  throw error;
}

/**
* Shows a message in a pop-up.
*
* @param {String} message The message to show.
* @param {Number} timeoutSeconds Optional. The number of seconds before the message goes away. Defaults to 5.
*/
function showMessage_(message, timeoutSeconds) {
  // version 1.0, written by --Hyde, 16 April 2020
  SpreadsheetApp.getActive().toast(message, 'Auto Move Rows', timeoutSeconds || 5);
}

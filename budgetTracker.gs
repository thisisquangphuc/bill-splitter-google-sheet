function newBill() {
  writeValueAndKeepFormat("D11:K11",0);
  writeValueAndKeepFormat("D16:F16",0);
  writeValueAndKeepFormat("D17:F17",0);
  writeValueAndKeepFormat("D18:F18",0);

  writeValueAndKeepFormat("D12:K12",0);
  var currentDate = new Date();
  writeValueAndKeepFormat("B11:B12",currentDate);
}

function clearTable() {
  clearRowFromAndKeepBackground(15);
  writeValueAndKeepFormat("D9:K9",0);
}

function copyRangeAndFormatWithInsertion() {
  var sourceSheetName = "Input"; // Replace with the name of the source sheet
  var destinationSheetName = "Report"; // Replace with the name of the destination sheet

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(sourceSheetName);
  var destinationSheet = ss.getSheetByName(destinationSheetName);

  var sourceRange = sourceSheet.getRange("B11:L12"); // Replace with the desired source range
  var destinationRange = destinationSheet.getRange("B15"); // Replace with the desired starting cell in the destination sheet

  // Insert two rows after row 14
  destinationSheet.insertRowsAfter(14, 2);

  // Copy the source range to the newly inserted area
  sourceRange.copyTo(destinationRange);
}


function calculateBill(rangeCell) {
  var person_invole = validateCellValueInRange("Input", "D11:K11");
  var person_count = person_invole.length;

  if (person_count == 0) {
    Browser.msgBox("Nothing to do");
    return;
  }
  Logger.log("person list:" +  person_invole);

  var subTotal = getCellValue("L11");
  var applicationFee = getCellValue("D16:F16");
  var shipPromo = getCellValue("D17:F17");
  var finalBill = getCellValue("D18:F18");
  var finalShipping = applicationFee - shipPromo;
  var total = finalBill - finalShipping;

  var kfactor;
  if (subTotal != 0) {
    kfactor = (total/subTotal)
  } else {
    Browser.msgBox("Devided by 0");
    return;
  }
  var actualPay;

  for (var i = 0; i < person_count; i++) {
    var itemCell = person_invole[i];
    var itemPrice = getCellValue(itemCell);
    
    actualPay = kfactor*itemPrice + finalShipping/person_count;
    writeValueAndKeepFormat(incrementCellReference(itemCell, 1), actualPay)
  }

}

function onEdit(e) {
  var sheet = e.source.getActiveSheet();

  if (sheet.getName() == "Input") {
    var editedCell = e.range;
    // Define the ranges to check
    var rangeD11toK11 = sheet.getRange("D11:K11");
    var mergedRanges = ["D16:F16", "D17:F17", "D18:F18"];
    var rangeMerged = sheet.getRangeList(mergedRanges);

    // Check if the edited cell is within the range D11 to K11 or any of the merged cells D16:F16, D17:F17, and D18:F18
    if (isWithinRange(editedCell, rangeD11toK11) || isWithinMergedRanges(editedCell, rangeMerged)) {
      var convertedValue = editedCell.getValue() * 1000;
      editedCell.setValue(convertedValue);
    }
  } else if (sheet.getName() == "Report") {
    var editedCell = e.range;
    // Define the ranges to check
    var rangeReport = sheet.getRange("D9:K9");
    var rangeCheckbox = sheet.getRange("D10:K10");

    // Check if the edited cell is within the range D11 to K11 or any of the merged cells D16:F16, D17:F17, and D18:F18
    if (isWithinRange(editedCell, rangeReport)){
      var convertedValue = editedCell.getValue() * 1000;
      editedCell.setValue(convertedValue);
    } 

    if (isWithinRange(editedCell, rangeCheckbox) && editedCell.getValue() == true){
      // Browser.msgBox(editedCell.getValue())
      var columnLetter = editedCell.getA1Notation().charAt(0); // Get the column letter from the edited cell

      // Clear content of the entire column starting from row 15
      var startRow = 15;
      var lastRow = sheet.getLastRow();
      var targetRange = sheet.getRange(columnLetter + startRow + ":" + columnLetter + lastRow);
      targetRange.clearContent();
      editedCell.setValue(false);
    }

  }
}

function getColumnLetter(sheetName, cellRange) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName); // Replace "Sheet1" with the name of your sheet

  var cellD10 = sheet.getRange(cellRange);
  var columnLetter = cellD10.getA1Notation().charAt(0); // Extract the first character (column letter)
  return columnLetter;
  Logger.log("Column letter from D10: " + columnLetter);
}


function clearColumnContent(columnLetter, startRow) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  for (var row = startRow; row <= sheet.getMaxRow(); row++) {
    sheet.getRange(row, columnLetter).setValue("");
  }
}

function clearRowFromAndKeepBackground(startRow) {
  var sheetName = "Report"; // Replace "Sheet1" with the name of your sheet

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  var numRows = sheet.getLastRow() - startRow + 1;
  var numCols = sheet.getLastColumn();

  // Get the range from row 14 to the last row and all columns
  var rangeToClear = sheet.getRange(startRow, 2, numRows, 11);

  // Get the background color of the range
  var backgrounds = rangeToClear.getBackgrounds();

  // Clear content and format of the range
  rangeToClear.clear({ contentsOnly: true, formatOnly: true, skipFilteredRows: false });

  // Restore the background color
  rangeToClear.setBackgrounds(backgrounds);

  // Unmerge cells in the range
  // rangeToClear.unmerge();
}


function isWithinRange(cell, range) {
  var cellRow = cell.getRow();
  var cellColumn = cell.getColumn();
  var rangeRowStart = range.getRow();
  var rangeRowEnd = range.getLastRow();
  var rangeColumnStart = range.getColumn();
  var rangeColumnEnd = range.getLastColumn();

  return cellRow >= rangeRowStart && cellRow <= rangeRowEnd && cellColumn >= rangeColumnStart && cellColumn <= rangeColumnEnd;
}

function isWithinMergedRanges(cell, rangeList) {
  var ranges = rangeList.getRanges();
  for (var i = 0; i < ranges.length; i++) {
    if (isWithinRange(cell, ranges[i])) {
      return true;
    }
  }
  return false;
}

function findFormulaCells(sheet, formulaToSearch) {
  var formulaCells = [];
  var dataRange = sheet.getDataRange();
  var formulas = dataRange.getFormulas();

  for (var row = 0; row < formulas.length; row++) {
    for (var col = 0; col < formulas[0].length; col++) {
      var formula = formulas[row][col];
      if (formula.indexOf(formulaToSearch) !== -1) {
        var cell = sheet.getRange(row + 1, col + 1);
        formulaCells.push(cell);
      }
    }
  }

  return formulaCells;
}

function validateCellValueInRange(sheetName, rangeString) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  var range = sheet.getRange(rangeString);
  var values = range.getValues();
  var numRows = values.length;
  var numColumns = values[0].length;

  var validCellRanges = []; // Initialize an empty array to store valid cell ranges

  for (var row = 0; row < numRows; row++) {
    for (var column = 0; column < numColumns; column++) {
      var cellValue = values[row][column];

      // Add your validation logic here
      if (cellValue !== null && cellValue !== "" && !isNaN(cellValue) && cellValue > 0) {
        var cell = sheet.getRange(range.getRow() + row, range.getColumn() + column);
        validCellRanges.push(cell.getA1Notation());
        Logger.log(cell.getA1Notation() + " - Cell value is valid.");
      } else {
        var cell = sheet.getRange(range.getRow() + row, range.getColumn() + column);
        Logger.log(cell.getA1Notation() + " - Cell value is not valid.");
      }
    }
  }

  return validCellRanges;
}

function writeValueAndKeepFormula(cellAddress, value) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getRange(cellAddress);
  
  // Get the existing formula in the cell
  var formula = cell.getFormula();

  // Clear the formula in the cell temporarily
  // cell.setFormula(null);

  cell.clearContent();

  // Re-enter the formula in the cell to preserve it
  cell.setFormula(formula);

}


function writeValueAndKeepFormat(cellRange, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRange(cellRange);
  range.setValue(value);
}

function writeValueAndFormatColor(cellRange, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRange(cellRange);
  range.setValue(value);
}

function countCellsGreaterThanZero(cellRange) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRange(cellRange);
  var values = range.getValues();
  var count = 0;
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] > 0) {
        count++;
      }
    }
  }
  return count;
}

function incrementCellReference(cellReference, step) {
  var column = cellReference.match(/[A-Z]+/)[0];
  var row = Number(cellReference.match(/\d+/)[0]);
  var newRow = row + step;
  return column + newRow;
}

function getCellValue(cellRange) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRange(cellRange);
  var value = range.getValue();
  return value;
}

function sumCells(cellRange) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var values = ss.getRange(cellRange).getValues();
  var sum = 0;
  for (var i = 0; i < values.length; i++) {
    sum += values[i];
  }
  return sum;
}

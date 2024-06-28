
const ss = SpreadsheetApp.getActiveSpreadsheet();
const inventory_sheet = ss.getSheetByName("Inventory");
const permissions_sheet = ss.getSheetByName("Permissions");
const detech_code_sheet = ss.getSheetByName("Detech_Code");


function onEdit(e) {
  const active_sheet = e.source.getActiveSheet();
  if (active_sheet.getName() !== "Permissions") return;

  const activeCell = e.range;
  const activeCellValue = activeCell.getValue();

  if (activeCell.getA1Notation() === "B1") {
    setupAllHandsOnDeckMode(activeCellValue);
  }
}

function setupAllHandsOnDeckMode(activeCellValue) {
  //let fullPermissionsSheetValues = permissions_sheet.



  // Set the background of the text to indicate it has been set or unset
  // DO NOT do this with conditional formatting. If this function does not work because of user input
  // which can happen when the user clicks the check box directly after editing another cell
  // then it will give the wrong indication. So do the background setting here instead
  if (activeCellValue) {
    permissions_sheet.getRange("A1").setBackground("#ea9999");
  } else {
    permissions_sheet.getRange("A1").setBackground("#b6d7a8");
  }
}






/*

=====================================================================================================================================================
        Webpage Code
=====================================================================================================================================================
*/

// Setup the webpage
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Search_Feature_Form');
}

function getCachedInventoryDataGSFunction() {
  const totalInventoryItems = detech_code_sheet.getRange("A6").getValue();
  const totalColumns = detech_code_sheet.getRange("B6").getValue();
  // Get the data range in the "Inventory" sheet starting from A2
  const startRow = 2;
  const startColumn = 1; // Column A
  const numRows = totalInventoryItems; // Number of rows to include in the range
  const numColumns = totalColumns; // Number of columns to include in the range
  // Fetch the data as a 2D array
  return inventory_sheet.getRange(startRow, startColumn, numRows, numColumns).getValues();
}

// Get the current users email to track which permissions correspond to them
function getActiveUserGSFunction() {
  return Session.getActiveUser().getEmail();
}

function getPermissionsListGSFunction(rowStart, colStart=1) {
  // Get the row and column range from cells A2 and B2
  const rowAndColRange = detech_code_sheet.getRange("A2:B2").getValues();
  // Calculate the end row and column numbers
  const rowRangeEnd = rowAndColRange[0][0] + rowStart - 1;
  const colRangeEnd = rowAndColRange[0][1] + colStart - 1;
  // Define the number of rows and columns
  const numRows = rowRangeEnd - rowStart + 1;
  const numCols = colRangeEnd - colStart + 1;
  // Return the data ranges values from the "Permissions" sheet
  return permissions_sheet.getRange(rowStart, colStart, numRows, numCols).getValues();
}

function updateInventoryFromWebpage(row, col, value) {
  try {
    inventory_sheet.getRange(row+1, col).setValue(value); 
    return false; // Update successful
  } catch (e) {
    return e; // Update failed
  }
}


/*
=====================================================================================================================================================
      Simplifier Code
=====================================================================================================================================================
*/

function intToLetter(num) {
    return String.fromCharCode(num+64);
}

function alert(msg) {
  SpreadsheetApp.getUi().alert(msg);
}






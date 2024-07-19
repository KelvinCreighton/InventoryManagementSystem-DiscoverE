
const ss = SpreadsheetApp.getActiveSpreadsheet();
const inventory_sheet = ss.getSheetByName("Inventory");
const permissions_sheet = ss.getSheetByName("Permissions");
const detech_code_sheet = ss.getSheetByName("Detech_Code");
const deleted_archive_sheet = ss.getSheetByName("Deleted_Archive");


function onEdit(e) {
  const active_sheet = e.source.getActiveSheet();
  if (active_sheet.getName() !== "Permissions") return;

  const activeCell = e.range;
  const activeCellValue = activeCell.getValue();

  // Check if the active cell is a dropdown menu
  if (activeCell.getDataValidation().getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    let userEmail = active_sheet.getRange(activeCell.getRow(), 1).getValue();
    updateInventorySheetPermissions(activeCellValue, activeCell.getColumn()-2, userEmail);
  }

  // Finish this later if its still a good idea
  /*if (activeCell.getA1Notation() === "B1") {
    setupAllHandsOnDeckMode(activeCellValue);
  } else if (activeCell.getA1Notation() === "B2" && activeCellValue) {
    addNewUserWithPermissions();
    activeCell.setValue(false);
  }*/
}

function updateInventorySheetPermissions(activeCellValue, activeColumn, userEmail) {
  // Ignore this function for these specific users. They will NOT have any editing permissions
  if (userEmail === "Other Users")
    return;
  
  // Ignore this function for these specific users. They will have FULL editing permissions
  if (userEmail === "detech@ualberta.ca" || userEmail === "degem@ualberta.ca" || userEmail === "desi1@ualberta.ca")
    return;
  
  const protections = inventory_sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (let i = 0; i < protections.length; i++) {
    let protectionColumn = protections[i].getDescription();
    if (protectionColumn !== activeColumn.toString())
      continue;

    if (activeCellValue === "Edit") {
      protections[i].addEditor(userEmail);
      alert("Editing Permissions Added to: " + userEmail);    
    } else {
      protections[i].removeEditor(userEmail);
      alert("Editing Permissions Removed from: " + userEmail);
    }
  }
}

/*function addNewUserWithPermissions() {
  let permissionsStartRow = detech_code_sheet.getRange("C2").getValue();
  let permissionsRowTotal = detech_code_sheet.getRange("A2").getValue();
  let otherUsersRow = permissionsStartRow + permissionsRowTotal - 1;
  permissions_sheet.insertRowBefore(otherUsersRow);
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
*/





/*

=====================================================================================================================================================
        Webpage Code
=====================================================================================================================================================
*/

// Setup the webpage
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Search_Feature_Form");
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

function addInventoryItem(item) {
  try {
    // Add new row with the proper setup
    inventory_sheet.insertRowBefore(2);
    inventory_sheet.getRange("A2").setFormula('=IF(B2 <> "", Row()-1, "")');

    for (let i = 1; i < item.length; i++) {  // Skip the first element which is just the index of the item in the inventory dataset
      let cellRange = intToLetter(i+1) + "2";
      inventory_sheet.getRange(cellRange).setValue(item[i]);
    }
    return false;
  } catch (e) {
    return e;
  }
}

function deleteInventoryRow(rowIndex) {
  try {
    // Archive the deleted row
    let row = inventory_sheet.getRange(rowIndex+1, 2, 1, inventory_sheet.getLastColumn()).getValues()[0];
    deleted_archive_sheet.appendRow(row);
  } catch (e) {
    return e;
  }

  try {
    // Delete the row in the inventory database
    inventory_sheet.deleteRow(rowIndex+1);
    return false;
  } catch (e) {
    return e;
  }

  // I probably dont need to split these try statements but I thought it might be extra safe
}

function updateInventoryFromWebpage(row, col, value, itemName) {
  try {
    let currentItemName = inventory_sheet.getRange(row+1, 2).getValue();
    if (currentItemName !== itemName)
      return "The item may have been deleted or moved in the inventory. The dataset has been reloaded please try again";

    inventory_sheet.getRange(row+1, col).setValue(value);
    return false; // Update successful
  } catch (e) {
    return e;     // Update failed
  }
}


/*
=====================================================================================================================================================
      Simplifier Code
=====================================================================================================================================================
*/

function intToLetter(number) {
  return String.fromCharCode(number + 64);
}

function letterToInt(letter) {
  return letter.charCodeAt(0) - 64;
}

function alert(msg) {
  SpreadsheetApp.getUi().alert(msg);
}



/*


POTENTIAL ISSUES:

There is this huge issue where if a user manually adds or deletes an item in the inventory. since the cache is only updated every 1 minute this could have serious issues if a user were to try modifiying items in a wrong row.

      Syncing Issues:

1: when a user deletes an item and then another user tries to modify that same item without updating the cache
2: when a user deletes an item and then another user tries to delete the same item without updating the cache


SOLUTIONS:


      Syncing Solutions:

1: Refresh the cache every 30 seconds or less
2: When users are modifying anything with the inventory have the cache update. However do not delete the code that modifies the cache directly while it is being updated. This makes things feel fast even if its not.
3: When a user modifies an item have it check if the item still exists. if it no longer exists then alert the user it has been deleted
4: when a user deletes an item have it also check if the item still exists.

*/







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
    updateInventorySheetPermissions(activeCellValue, activeCell.getColumn()-2, activeCell.getRow(), userEmail);
  }

  // Finish this later if its still a good idea
  /*if (activeCell.getA1Notation() === "B1") {
    setupAllHandsOnDeckMode(activeCellValue);
  } else if (activeCell.getA1Notation() === "B2" && activeCellValue) {
    addNewUserWithPermissions();
    activeCell.setValue(false);
  }*/
}

function updateInventorySheetPermissions(activeCellValue, activeColumn, activeRow, userEmail, alertUser=true) {
  // Skip for the index column, this column should never be editable
  if (activeColumn === 1)
    return;
  
  // Ignore this function for these specific users. They will NOT have any editing permissions. This does not need to be changed
  if (userEmail === "Other Users")
    return;
  
  // Ignore this function for these specific users. They will have FULL editing permissions. This does not need to be changed
  if (userEmail === "detech@ualberta.ca" || userEmail === "degem@ualberta.ca" || userEmail === "desi1@ualberta.ca")
    return;

  
  let tempPermissions = permissions_sheet.getRange(activeRow, 4, activeRow, 1000).getValues().flat();
  Logger.log(tempPermissions);
  
  // If a user is allowed to edit any value of an item then force the program to allow an edit permissions in the items "Date" column
  if (activeColumn !== 3) {
    if (activeCellValue === "Edit") { // If the current permissions setting is to allow for an edit then set the date column to edititable
      updateInventorySheetPermissions("Edit", 3, activeRow, userEmail, false);
    } else {  // Otherwise check if there are no other edits then remove the editiable permission of the date column
      let noEdits = true;
      for (let j = 0; j < tempPermissions.length; j++) {
        if (tempPermissions[j] === "Edit") { // An edit was found
          noEdits = false;
          break;
        } else if (tempPermissions[j] === "") {  // No edits were found an the end of the permissions has been reached
          break;
        }
      }
      if (noEdits) {
        updateInventorySheetPermissions("Hidden", 3, activeRow, userEmail, false);
      }
    }
  }
    
  const protections = inventory_sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (let i = 0; i < protections.length; i++) {
    let protectionColumn = protections[i].getDescription();


    // If a user is allowed to delete an item then force the program to allow an edit permissions in the items "Name" column
    if (activeColumn === 0) {
      activeColumn = 2;
      if (activeCellValue === "Add & Delete") { // If it allows delete then force the names to "Edit"
        activeCellValue = "Edit"
      } else {  // If it does not allow delete then find the default value of the names permissions
        activeCellValue = permissions_sheet.getRange(activeRow, 4).getValue();
      }
    }

    // And ignore if the name is changing but the delete is allowed
    if (activeColumn === 2 && activeCellValue !== "Edit" && permissions_sheet.getRange(activeRow, 2).getValue() === "Add & Delete")
      return;

    if (protectionColumn !== activeColumn.toString())
      continue;

    if (activeCellValue === "Edit") {
      protections[i].addEditor(userEmail);
      if (alertUser)
        alert("Editing Permissions Added to: " + userEmail);    
    } else {
      protections[i].removeEditor(userEmail);
      if (alertUser)
        alert("Editing Permissions Removed from: " + userEmail);
    }
  }
}


function timeDrivenTriggerRemoveEmptyItemRows() {
  // Search all rows with empty item names which mean they are flagged for deletion
  const lastItemRow = detech_code_sheet.getRange("A6").getValue();
  let values = inventory_sheet.getRange(2, 2, lastItemRow - 1).getValues();
  let changeHappened = false;
   // Loop from the last row to the first row to avoid indexing issues after row deletion
  for (let y = values.length - 1; y >= 0; y--) {
    if (values[y][0] === "") { // Check if the item name is empty
      inventory_sheet.deleteRow(y + 2);
      changeHappened = true;
    }
  }

  // Clear the remaining rows below the last item row
  const lastRow = inventory_sheet.getLastRow();
  const lastCol = inventory_sheet.getLastColumn();
  inventory_sheet.getRange("A" + (lastItemRow+1) + ":" + intToLetter(lastCol) + lastRow).clear();

  //if (changeHappened)
    //timeDrivenTriggerInventorySheetProtectionsUpdateFunction(true);
}


// Updates the range of the current permissions in each column of the items
// When users add new items they require full edit permissions so the current permissions do not extend past the last item
// Since this is the case and since non Admin users cannot update permissions, a trigger must be run every so often to update the current permissions
// This function checks if there is a new item/items and updates the range of the current permissions accordingly
function timeDrivenTriggerInventorySheetProtectionsUpdateFunction(skipCheck=false) {
  let totalInventoryItems = detech_code_sheet.getRange("A6").getValue();
  let lastRecordedItemIndexRange = detech_code_sheet.getRange("A10");
  // Check for new items and skip this trigger if nothing has changed
  if (!skipCheck && totalInventoryItems === lastRecordedItemIndexRange.getValue())
    return;
  
  // Update each projections range to include the new row
  const protections = inventory_sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  const protectionsColumns = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "A", "OTHER"];
  for (let i = 0; i < protections.length; i++) {
    let protectionDescription = protections[i].getDescription();
    // Check if the current protection is in the desired protection columns array and skip any that is not
    let protectionInColumn = false;
    for (let j = 0; j < protectionsColumns.length; j++) {
      if (protectionDescription === protectionsColumns[j]) {
        protectionInColumn = true;
        break;
      }
    }
    if (!protectionInColumn)
      continue;

    let protectionsRange = protections[i].getRange().getA1Notation();
    let protectionsRangeFirstPart = protectionsRange.slice(0, 4);       // Example: From "A2:A1327" to "A2:A" without the "1327"
    let newProtectionRange = protectionsRangeFirstPart + (totalInventoryItems).toString();
    protections[i].setRange(inventory_sheet.getRange(newProtectionRange));
  }
  
  // Update the checker cell
  lastRecordedItemIndexRange.setValue(totalInventoryItems);
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
  const numRows = totalInventoryItems-1; // Number of rows to include in the range
  const numColumns = totalColumns; // Number of columns to include in the range
  // Fetch the data as a 2D array
  return inventory_sheet.getRange(startRow, startColumn, numRows, numColumns).getValues();
}

// Get the current users email to track which permissions correspond to them
function getActiveUserGSFunction() {
  return Session.getActiveUser().getEmail();
}

function getPermissionsListGSFunction(colStart=1) {
  const rowStart = detech_code_sheet.getRange("C2").getValue();
  // Get the row and column range from cells A2 and B2
  const rowAndColRange = detech_code_sheet.getRange("A2:B2").getValues();
  // Calculate the end row and column numbers
  const rowRangeEnd = rowAndColRange[0][0] + rowStart - 1;
  const colRangeEnd = rowAndColRange[0][1] + colStart - 1;
  // Define the number of rows and columns
  const numRows = rowRangeEnd - rowStart + 1;
  const numCols = colRangeEnd - colStart + 1;
  // Return the data ranges values from the "Permissions" sheet
  return permissionsList = permissions_sheet.getRange(rowStart, colStart, numRows, numCols).getValues();
}

function addInventoryItem(item) {
  try {
    // Set the items date to the current date
    item[2] = getCurrentSemester();

    // Get the last row
    let lastRow = detech_code_sheet.getRange("A6").getValue();
    // Append a new row at the end of the inventory 
    inventory_sheet.insertRowAfter(lastRow);

    // Add the items data from the webpage to the columns in the "Inventory" sheet
    for (let i = 1; i < item.length; i++) {  // Skip the first element which is just the index of the item in the inventory dataset
      let cellRange = intToLetter(i+1) + (lastRow+1).toString();
      inventory_sheet.getRange(cellRange).setValue(item[i]);
    }
    return false;
  } catch (e) {
    return e.message;
  }
}

function deleteInventoryRow(rowIndex, itemName) {
  try {
    let currentItemName = inventory_sheet.getRange(rowIndex, 2).getValue();
    if (currentItemName !== itemName)
      return "The item may have been deleted or moved in the inventory. The dataset has been reloaded please try again";
      
    // Archive the deleted row
    let row = inventory_sheet.getRange(rowIndex, 2, 1, inventory_sheet.getLastColumn()).getValues()[0];
    deleted_archive_sheet.insertRowBefore(2);
    deleted_archive_sheet.getRange(2, 1, 1, row.length).setValues([row]);
  } catch (e) {
    return e.message;
  }

  try {
    // Delete the name of the item to flag the next trigger to delete the row
    inventory_sheet.getRange(rowIndex, 2).setValue("");
    return false;
  } catch (e) {
    return e.message;
  }

  // I probably dont need to split these try statements but I thought it might be extra safe
}

// This function is used when a user adds changes to an items attribute in the website
function updateInventoryFromWebpage(row, col, value, itemName) {
  try {

    let currentItemName = inventory_sheet.getRange(row, 2).getValue();
    if (currentItemName !== itemName)
      return "The item may have been deleted or moved in the inventory. The dataset has been reloaded please try again";

    // Update the items cell
    inventory_sheet.getRange(row, col).setValue(value);
    // Set the date that it was changed
    inventory_sheet.getRange(row, 3).setValue(getCurrentSemester());
    return false; // Update successful
  } catch (e) {
    return e.message;     // Update failed
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

// Check if a string is a whole number
function isWholeNumber(str) {
  return !isNaN(str) && str.trim() !== "";
}

function getCurrentSemester() {
  const date = new Date();
  const day = date.getDate();
  const month = date.getMonth()+1;
  const year = date.getFullYear();

  let semester = "";
  if (month >= 1 && month <= 4) {
    semester = "Winter";
  } else if (month >= 5 && month <= 8) {
    semester = "Summer";
  } else if (month >= 9 && month <= 12) {
    semester = "Fall";
  }
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return semester + " " + year + ", " + monthNames[month-1] + " " + day;
}


/*


POTENTIAL ISSUES:

There is this huge issue where if a user manually adds or deletes an item in the inventory. since the cache is only updated every 60 seconds this could have serious issues if a user were to try modifiying items in an incorrect or misplaced row.

      Syncing Issues:

1: when a user deletes an item and then another user tries to modify that same item without updating the cache
2: when a user deletes an item and then another user tries to delete the same item without updating the cache


SOLUTIONS:


      Syncing Solutions:

1: Refresh the cache every 60 seconds or less
2: When users are modifying anything with the inventory have the cache update. However do not delete the code that modifies the cache directly while it is being updated. This makes things feel fast even if its not.
3: When a user modifies an item have it check if the item still exists. if it no longer exists then alert the user it has been deleted
4: when a user deletes an item have it also check if the item still exists.

*/






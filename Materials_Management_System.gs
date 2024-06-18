
const ss = SpreadsheetApp.getActiveSpreadsheet();
const search_sheet = ss.getSheetByName("Search");
const search_query_sheet = ss.getSheetByName("Search_Query");
const inventory_sheet = ss.getSheetByName("Inventory");
const permissions_sheet = ss.getSheetByName("Permissions");
const detech_code_sheet = ss.getSheetByName("Detech_Code");


function onEdit(e) {
  cachedInventoryData = "test";
  const active_sheet = e.source.getActiveSheet();
  if (active_sheet.getName() !== "Search") return;

  const activeCell = e.range;

  let startRow = 9;
  let startCol = 4;
  let numberOfRows = search_query_sheet.getRange("A8").getValue();
  let numberOfCols = search_query_sheet.getRange("B8").getValue()+1;


  if (activeCell.getA1Notation() === "G1") {
    activeCell.setValue(false);
    resetFormulas(startRow, startCol, 500, numberOfCols);
  } else {
    
    // Check if the range of the active cell is in the search dataset range
    let activeRow = activeCell.getRow();
    let activeCol = activeCell.getColumn();
    if (activeRow >= startRow && activeRow <= startRow+numberOfRows-1 && activeCol >= startCol && activeCol <= startCol+numberOfCols-1) {
      let searchQueryIndexCol = startCol+numberOfCols;
      updateInventoryCell(activeRow, activeCol, searchQueryIndexCol, activeCell);
    }
  }
}


function resetFormulas(startRow, startCol, numberOfRows, numberOfCols) {
  let searchData = search_sheet.getRange(startRow, startCol, numberOfRows, numberOfCols);
  let searchDataValues = searchData.getValues();
  let searchDataFormulas = searchData.getFormulas();
    
  for (let y = 0; y < searchDataValues.length; y++) {
    for (let x = 0; x < searchDataValues[y].length; x++) {
      if (searchDataFormulas[y][x] !== "")  continue;   // Skip any unedited cells
      search_sheet.getRange(startRow+y, startCol+x).setFormula("=Search_Query!" + (intToLetter(startCol+x)) + "" + (startRow+y));
    }
  }
}

function updateInventoryCell(row, col, indexCol, activeCell) {
  let inventoryListRow = search_query_sheet.getRange(row, indexCol).getValue()+1;
  // Error check for index not found
  if (inventoryListRow == 1) {
    alert("Out of Bounds Index in Search_Query Sheet");
    return;
  }
  // Map the active column to the inventory column
  let inventoryListCol = col == 4 ? 2 : col  // col = B if col == D otherwise col = col

  // Map the active column to the inventory column (use this if the mapping changes more)
  /*switch (col) {
    case 4:   // Name
      col = 2;
      break;
    case 5:   // Location
      col = 5;
      break;
    case 6:   // Shelf
      col = 6;
      break;
    case 7:   // Bin
      col = 7;
      break;
    case 8:   // Amount
      col = 8;
      break;
    case 9:   // Default Unit
      col = 9;
      break;
    case 10:  // Conversion Unit
      col = 10;
      break;
    case 11:  // Conversion Quantity
      col = 11;
      break;
  }*/



  inventory_sheet.getRange(inventoryListRow, inventoryListCol).setValue(activeCell.getValue());
  activeCell.setValue("=Search_Query!" + (intToLetter(col)) + "" + (row))
}







/*

===========================================================================================================================================================================

*/


function doGet() {
  return HtmlService.createHtmlOutputFromFile('Search_Feature_Form');
}

function getCachedInventoryData() {
  let totalInventoryItems = search_query_sheet.getRange("B11").getValue();
  let data = inventory_sheet.getRange("A2:I" + (totalInventoryItems+1)).getValues();
  return data;
}

function getActiveUser() {
  return Session.getActiveUser().getEmail();
}

function getPermissionsList() {
  let rowAndColRange = detech_code_sheet.getRange("A2:B2").getValues();
  let row = rowAndColRange[0][0]+3;
  let col = intToLetter(rowAndColRange[0][1]);
  let data = permissions_sheet.getRange("A4:"+col+""+row).getValues();
  return data;
}
/*

===========================================================================================================================================================================

*/

function intToLetter(num) {
    return String.fromCharCode(num+64);
}

function alert(msg) {
  SpreadsheetApp.getUi().alert(msg);
}






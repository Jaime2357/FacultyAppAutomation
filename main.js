// Simple Instructions for Usage:

//  Select the function you'd like to use from the dropdown menu above (next to Debug):
//      - appendSingleRow: Appends the selected row to the last row of the new sheet 
//        based on section and request count. (
//        Notes:
//           * Change variable "SpecificRow" to correct row number -> You'll find this in the "Variables" section below)
//           * Make sure to click Save (Button to the left of "Run") before clicking Run
//      - appendNewRows: Takes all rows in the sheet where the "duplicated" column (usually last column) isn't filled with an
//        'X' and appends them to the new sheet based on section and request count.
//      - appendAllRows: Takes all rows in the sheet and appends them to the new sheet 
//        based on section and request count.
//  Then Click run
//      - A text box will pop up, this will inform you when the process completes. 
//        If an error occurs, it will show up in this text box. - Contact Jaime for big errors



// Conversion Utility function -------------------------------------------------------------------------------------------------------------------------------------

function colNum(column) {
  let depth = column.length - 1;
  return (depth * 26) + letter.charCodeAt(depth) - 64;
}



// Variables -------------------------------------------------------------------------------------------------------------------------------------------------------

const SpecificRow = 1; // For appendSingleRow() -> specify row to append

const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Applications"); //Source Sheet
const destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Section Duplicates"); //Target Sheet

const requestFieldCount = 7; // Number of fields in each specific course request (should be uniform)
const lastColumn = sourceSheet.getLastColumn();

const sectionColumnOffset = 1; //The section is usually the column next to the course request, but just in case <-
const maxRequests = 4; //Will likely not change, but just in case

const reqColumns = [colNum('G'), colNum('N'), colNum('U'), colNum('AB')]; //List of column numbers with Course Request Names
const dCheckColumn = colNum('BM'); // Column with X's to check if a row has already been duplicated



// Utility Functions -----------------------------------------------------------------------------------------------------------------------------------------------function filledCount(){

function filledCount() {
  let values = sourceSheet.getRange(1, 1, sourceSheet.getLastRow(), 1).getValues();
  let count = 0;
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] !== "") {
      count++;
    }
  }
  return count;
}

function setRequests(row) {
  let requests = 4; // Max number of requests
  let cell = sourceSheet.getRange(row, reqColumns[requests]);
  while (cell.isBlank()) { //Iterate until non-null request found
    requests--; // Decrement max number of requests
    cell = sourceSheet.getRange(SpecificRow, reqColumns[requests]);
  };
  return requests;
}

function setSections(row, reqCol) {
  let secCountCol = reqCol + sectionColumnOffset;
  return sourceSheet.getRange(row, secCountCol).getValue();
}

function getRanges(i) {
  let a = reqColumns[0] - 1;
  let b = reqColumns[i - 1];
  let c = reqColumns[i] - 1;
  let d = reqColumns[reqCount] + requestFieldCount;
  return includedColumnsRanges = [
    { start: 1, end: a },
    { start: b, end: c },
    { start: d, end: lastColumn }
  ];
}

function buildRow(range) {
  let values = []
  // Loop through each column and check if it falls within the included ranges
  for (let col = 1; col <= lastColumn; col++) {
    if (range.includes(col)) {
      let cellValue = sourceSheet.getRange(SpecificRow, col).getValue();
      values.push(cellValue);
    }
  }
  return values;
}

function appendTable(builtRow, sectionCount) {
  // Append the values to the next empty rows in the destination sheet
  for (let i = 0; i < sectionCount; i++) {
    let lastRow = destinationSheet.getLastRow();
    destinationSheet.getRange(lastRow + 1, 1, 1, values.length).setValues([builtRow]);
  }
}

function appendRow(row) { // Main driver function
  const reqCount = setRequests(row); // Get the number of requests

  for (let i = 1; i <= reqCount; i++) {
    let sectionCount = setSections(row, reqColumns[i - 1]);
    let range = getRanges(i);
    var builtRow = buildRow(range);
    appendTable(builtRow, sectionCount);
  }
}



// Main Functions --------------------------------------------------------------------------------------------------------------------------------------------------
function appendSingleRow() {
  sourceSheet.getRange(SpecificRow, dCheckColumn).check(); // Checks if this row has already been duplicated
  appendRow(SpecificRow);
}

function appendNewRows() {

  const count = filledCount();
  for (var row = 2; row <= count; row++) {// Repeat process for each row in the original sheet
    if (sourceSheet.getRange(row, dCheckColumn).isBlank()) {// Only duplicate NEW rows (Rows that do not have the duplicate checkbox marked)
      appendRow(row);
    }
    sourceSheet.getRange(row, dCheckColumn).setValue('X');
  }
}

function appendAllRows() {
  const count = filledCount();
  for (var row = 2; row <= count; row++) {
    appendRow(row);
  }
  sourceSheet.getRange(row, dCheckColumn).setValue('X');
}
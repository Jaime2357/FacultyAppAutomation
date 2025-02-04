// Instructions for Usage:

// Select the function you'd like to use from the dropdown menu above (next to Debug):
//     - appendRow: Appends the selected row to the last row of the new sheet 
//       based on section and request count. (Change row Variable to correct row number - specified by dashes)
//     - appendAllRows: Takes all rows in the sheet and appends them to the new sheet 
//       based on section and request count.
// Then Click run
//     - A text box will pop up, this will inform you when the process completes. 
//       If an error occurs, it will show up in this text box.

// New Text

function appendRow() {

  row = 59; //Change based on newest row -----------------------------------------------------

  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Applications");
  var destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Section Duplicates");

  sourceSheet.getRange(row, 61).check();

  var requests = 4;
  var column = 25;
  var cell = sourceSheet.getRange(row, column);
  while(cell.isBlank()){
    requests--;
    cell = sourceSheet.getRange(row, (column -= 6));
  };

  var sectionColumn = 8;
    
  for(var j = 1; j <= requests; j++){ 
    var sectionCount = sourceSheet.getRange(row, sectionColumn).getValue();

    // Get the number of columns in the source sheet
    var lastColumn = sourceSheet.getLastColumn();

    switch(j){
      case 1:
        var includedColumnsRanges = [
        { start: 1, end: 12 },
        { start: 31, end: lastColumn }
        ];
        break;
      case 2:
        var includedColumnsRanges = [
        { start: 1, end: 6 },
        { start: 13, end: 18 },
        { start: 31, end: lastColumn }
        ];
        break;
      case 3:
        var includedColumnsRanges = [
        { start: 1, end: 6 },
        { start: 19, end: 24 },
        { start: 31, end: lastColumn }
        ];
        break;
      case 4:
        var includedColumnsRanges = [
        { start: 1, end: 6 },
        { start: 25, end: lastColumn }
        ];
        break
    }

    // Create an array to store the values of the desired rows
    var values = [];

    // Loop through each column and check if it falls within the included ranges
    for (var col = 1; col <= lastColumn; col++) {
      var isIncluded = includedColumnsRanges.some(function(range) {
      return col >= range.start && col <= range.end;
      });
  
      if (isIncluded) {
        var cellValue = sourceSheet.getRange(row, col).getValue();
        values.push(cellValue);
      }
    }

    // Append the values to the next empty rows in the destination sheet
    for (var k = 0; k < sectionCount; k++) {
      var lastRow = destinationSheet.getLastRow();
      destinationSheet.getRange(lastRow+1, 1, 1, values.length).setValues([values]);
    }

    sectionColumn += 6;
  }
}

function appendNewRows(){

  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Applications");
  var destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Section Duplicates");

  var columnNumber = 1;

  var dataRange = sourceSheet.getRange(1, columnNumber, sourceSheet.getLastRow(), 1);
  var values = dataRange.getValues();

  var count = 0;
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] !== "") {
      count++;
    }
  }

  
  for(var row = 2; row <= count; row++){
    
    if(sourceSheet.getRange(row, 61).isBlank()){
  
      var requests = 4;
      var column = 25;
      var cell = sourceSheet.getRange(row, column);
      while(cell.isBlank()){
        requests--;
        cell = sourceSheet.getRange(row, (column -= 6));
      };

      var sectionColumn = 8;
    
      for(var j = 1; j <= requests; j++){
      
        var sectionCount = sourceSheet.getRange(row, sectionColumn).getValue();

        // Get the number of columns in the source sheet
        var lastColumn = sourceSheet.getLastColumn();

        switch(j){
          case 1:
            var includedColumnsRanges = [
            { start: 1, end: 12 },
            { start: 31, end: lastColumn }
            ];
            break;
          case 2:
            var includedColumnsRanges = [
            { start: 1, end: 6 },
            { start: 13, end: 18 },
            { start: 31, end: lastColumn }
            ];
            break
          case 3:
            var includedColumnsRanges = [
            { start: 1, end: 6 },
            { start: 19, end: 24 },
            { start: 31, end: lastColumn }
            ];
            break
          case 4:
            var includedColumnsRanges = [
            { start: 1, end: 6 },
            { start: 25, end: lastColumn }
            ];
            break
        }

        // Create an array to store the values of the desired rows
        var values = [];

        // Loop through each column and check if it falls within the included ranges
        for (var col = 1; col <= lastColumn; col++) {
          var isIncluded = includedColumnsRanges.some(function(range) {
          return col >= range.start && col <= range.end;
        });
  
        if (isIncluded) {
          var cellValue = sourceSheet.getRange(row, col).getValue();
          values.push(cellValue);
        }
      }

      // Append the values to the next empty rows in the destination sheet
      for (var k = 0; k < sectionCount; k++) {
        var lastRow = destinationSheet.getLastRow();
        destinationSheet.getRange(lastRow+1, 1, 1, values.length).setValues([values]);
      }

      sectionColumn += 6;
      }
    }
    sourceSheet.getRange(row, 61).setValue('X');
  }
}

function appendAllRows() {

  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Applications");
  var destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Section Duplicates");

  var columnNumber = 1;

  var dataRange = sourceSheet.getRange(1, columnNumber, sourceSheet.getLastRow(), 1);
  var values = dataRange.getValues();

  var count = 0;
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] !== "") {
      count++;
    }
  }

  for(var row = 2; row <= count; row++){

    var requests = 4;
    var column = 25;
    var cell = sourceSheet.getRange(row, column);
    while(cell.isBlank()){
      requests--;
      cell = sourceSheet.getRange(row, (column -= 6));
    };

    var sectionColumn = 8;
    
    for(var j = 1; j <= requests; j++){
      
      var sectionCount = sourceSheet.getRange(row, sectionColumn).getValue();

      // Get the number of columns in the source sheet
      var lastColumn = sourceSheet.getLastColumn();

      switch(j){
        case 1:
          var includedColumnsRanges = [
          { start: 1, end: 12 },
          { start: 31, end: lastColumn }
          ];
          break;
        case 2:
          var includedColumnsRanges = [
          { start: 1, end: 6 },
          { start: 13, end: 18 },
          { start: 31, end: lastColumn }
          ];
          break
        case 3:
          var includedColumnsRanges = [
          { start: 1, end: 6 },
          { start: 19, end: 24 },
          { start: 31, end: lastColumn }
          ];
          break
        case 4:
          var includedColumnsRanges = [
          { start: 1, end: 6 },
          { start: 25, end: lastColumn }
          ];
          break
      }

      // Create an array to store the values of the desired rows
      var values = [];

      // Loop through each column and check if it falls within the included ranges
      for (var col = 1; col <= lastColumn; col++) {
        var isIncluded = includedColumnsRanges.some(function(range) {
        return col >= range.start && col <= range.end;
        });
  
        if (isIncluded) {
          var cellValue = sourceSheet.getRange(row, col).getValue();
          values.push(cellValue);
        }
      }

      // Append the values to the next empty rows in the destination sheet
      for (var k = 0; k < sectionCount; k++) {
        var lastRow = destinationSheet.getLastRow();
        destinationSheet.getRange(lastRow+1, 1, 1, values.length).setValues([values]);
      }

      sectionColumn += 6;
    }
    sourceSheet.getRange(row, 61).setValue('X');
  }
}


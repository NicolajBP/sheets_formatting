/**
 * Formats colors in the sheet "2024" so that each cell for 7-day period is colored green if 3 or more cells are non-empty and yellow if 1 or 2 cells are non-empty..
 * Created with assistance from ChatGPT 
 * 02/01/2024
*/

function onChange(e) {
    Logger.log("Script triggered");
  
    var sheet;
  
    // Check if 'e' (event object) is available
    if (e && e.source) {
      sheet = e.source.getSheetByName('2024'); // Replace 'YourSheetName' with the actual name of your sheet
    } else {
      Logger.log('e is not available');
      // If 'e' is not available, use an alternative method to access the active sheet
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      sheet = ss.getSheetByName('2024');
      // Continue with the rest of your code...
    }
  
    // Rest of your code...
    if (sheet) {
      var startRow = 3; // Starting row for your data in column E
      var column = 5; // Column E
      var numRows = sheet.getLastRow() - startRow + 1;
      var dataRange = sheet.getRange(startRow, column, numRows, 1);
      var values = dataRange.getValues();
  
      for (var i = 0; i < numRows; i += 7) {
        var intervalValues = values.slice(i, i + 7).flat();
        var nonEmptyCount = intervalValues.filter(value => value !== "").length;
  
        // Apply conditional formatting based on the count of non-empty cells
        if (nonEmptyCount >= 3) {
          sheet.getRange(startRow + i, column, 7, 1).setBackground('#a8e4b8'); // Green color
          Logger.log('Painted green for rows ' + (startRow + i) + ' to ' + (startRow + i + 6));
        } else if (nonEmptyCount >= 1 && nonEmptyCount <= 2) {
          sheet.getRange(startRow + i, column, 7, 1).setBackground('#f7f798'); // Yellow color
          Logger.log('Painted yellow for rows ' + (startRow + i) + ' to ' + (startRow + i + 6));
        } else {
          sheet.getRange(startRow + i, column, 7, 1).setBackground(null); // Remove background color
          Logger.log('Removed background color for rows ' + (startRow + i) + ' to ' + (startRow + i + 6));
        }
      }
    }
  }
  
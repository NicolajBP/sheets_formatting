/**
 * Formats colors in the sheet "2024" so that each cell in a month is colored green if 2 or more cells are non-empty and yellow if 1 cell is non-empty..
 * Created with assistance from ChatGPT 
 * 03/01/2024
*/


function monthFormatter(e) {
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
      var startRow = 3; // Starting row for your data in column F
      var column = 6; // Column F
      var numRows = sheet.getLastRow() - startRow + 1;
  
      // Fetch all values in the specified range
      var dataRange = sheet.getRange(startRow, column, numRows, 1);
      var values = dataRange.getValues();
  
      // Check if the data range is not empty
      if (values.length > 0) {
        var currentMonthData = [];
        var currentMonth = null;
        var lastRow = startRow;
  
        for (var i = 0; i < numRows; i++) {
          var cellValue = values[i][0];
  
          // Extract day, month, and year from the date
          var date = sheet.getRange(startRow + i, 2).getValue(); // Date is in column B
  
          // Ensure the date is a valid JavaScript Date object
          if (!(date instanceof Date) || isNaN(date.getTime())) {
            Logger.log("Skipping invalid date at row " + (startRow + i));
            continue;
          }
  
          // Format the date to DD/MM/YYYY
          var formattedDate = Utilities.formatDate(date, sheet.getParent().getSpreadsheetTimeZone(), 'dd/MM/yyyy');
          var dateComponents = formattedDate.split('/');
          var newDay = parseInt(dateComponents[0], 10);
          var newMonth = parseInt(dateComponents[1], 10);
          var currentYear = parseInt(dateComponents[2], 10);
  
          // Check if a new month is encountered
          if (currentMonth !== null && newMonth !== currentMonth) {
            setMonthBackground(sheet, currentMonthData, column, lastRow, startRow + i);
            currentMonthData = [];
            lastRow = startRow + i + 1;
          }
  
          currentMonth = newMonth;
  
          // Collect data for the current month
          currentMonthData.push({
            row: startRow + i,
            value: cellValue
          });
        }
  
        // Set background for the last month
        if (currentMonth !== null) {
          setMonthBackground(sheet, currentMonthData, column, lastRow, startRow + numRows);
        }
      } else {
        Logger.log("No data found in the specified range.");
      }
    }
  }
  
  // Helper function to set background for the entire month
  function setMonthBackground(sheet, monthData, column, startRow, endRow) {
    if (monthData.length > 0) {
      var nonEmptyCount = monthData.filter(item => item.value !== "").length;
  
      if (nonEmptyCount >= 2) {
        sheet.getRange(startRow, column, endRow - startRow, 1).setBackground('#a8e4b8'); // Green color for the entire month
        Logger.log('Painted green for the entire month');
      } else if (nonEmptyCount === 1) {
        sheet.getRange(startRow, column, endRow - startRow, 1).setBackground('#f7f798'); // Yellow color for the entire month
        Logger.log('Painted yellow for the entire month');
      } else {
        sheet.getRange(startRow, column, endRow - startRow, 1).setBackground(null); // Remove background color
        Logger.log('Removed background color for the entire month');
      }
    }
  }
  
function Restore(e) {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = e.source.getActiveSheet();
    let range = e.range;
    let newValue = e.value;
    let oldValue = e.oldValue;
  
    let mainSheetName = sheet.getName().replace(/^Archive_/, "");
    let archiveSheetName = "Archive_" + mainSheetName;
    
    if (sheet.getName() === archiveSheetName && range.getColumn() === 9 && oldValue === "Done" && newValue === "Restore") {
      let row = range.getRow();
      let rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  
      // Remove the row ID column
      rowData.shift();
  
      let mainSheet = ss.getSheetByName(mainSheetName);
      let lastRow = mainSheet.getLastRow();
  
      // Create new Task ID based on the last row index
      let newTaskId = lastRow + 1; // Adjust the new Task ID
  
      // Add the new Task ID to the row data
      rowData.unshift(newTaskId);
  
      // Define drop-down options
      let optionsF = ["Low", "Medium", "High"];
      let optionsG = ["Urgent", "Not Urgent"];
      let optionsI = ["Not Started", "In Progress", "Done",'On Hold'];
  
      // Create data validation rules
      let ruleF = SpreadsheetApp.newDataValidation().requireValueInList(optionsF).build();
      let ruleG = SpreadsheetApp.newDataValidation().requireValueInList(optionsG).build();
      let ruleI = SpreadsheetApp.newDataValidation().requireValueInList(optionsI).build();
  
      // Get the new row index
      let newRowIdx = lastRow + 1;
  
      // Apply the data validation rules to the new row in columns F, G, and I
      mainSheet.getRange(newRowIdx, 6).setDataValidation(ruleF); // Column F
      mainSheet.getRange(newRowIdx, 7).setDataValidation(ruleG); // Column G
      mainSheet.getRange(newRowIdx, 9).setDataValidation(ruleI); // Column I
  
      // Append the row data (including the new Task ID) to the main sheet
      mainSheet.appendRow(rowData);
  
      // Delete the row from the archive sheet
      sheet.deleteRow(row);
      let archiveRows = sheet.getDataRange().getValues();
      for (let i = 1; i < archiveRows.length; i++) {
        archiveRows[i][0] = i; // Assuming the Task ID is in the first column
      }
      sheet.getDataRange().setValues(archiveRows);
    }
  }
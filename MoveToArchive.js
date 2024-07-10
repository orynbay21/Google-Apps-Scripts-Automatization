function MoveToArchive() {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheets = ss.getSheets();
    let now = new Date();
  
    for (let sheet of sheets) {
      let sheetName = sheet.getName();
  
      if (sheetName.startsWith('Archive_')) {
        continue;
      }
  
      let archiveSheetName = 'Archive_' + sheetName;
      let archiveSheet = ss.getSheetByName(archiveSheetName);
  
      if (!archiveSheet) {
        archiveSheet = ss.insertSheet(archiveSheetName);
  
        // Set Row ID header and apply bold font
        let rowIdCell = archiveSheet.getRange(1, 1);
        rowIdCell.setValue('Row ID');
        rowIdCell.setFontWeight("bold");
  
        // Get headers from the original sheet and apply bold font
        let headersRange = sheet.getRange(1, 2, 1, sheet.getLastColumn() - 1); // Exclude the first column
        headersRange.setFontWeight("bold");
  
        // Copy headers to the archive sheet and apply filter
        let archiveHeadersRange = archiveSheet.getRange(1, 2, 1, headersRange.getNumColumns());
        headersRange.copyTo(archiveHeadersRange);
        archiveHeadersRange.setFontWeight("bold");
  
        // Freeze the top row to treat it as a header
        archiveSheet.setFrozenRows(1);
      }
  
      let data = sheet.getRange(2, 2, sheet.getLastRow() - 1, sheet.getLastColumn() - 1).getValues(); // Exclude the first column
      let statusColumn = sheet.getRange('I2:I').getValues().filter(r => r[0] != "").flat();
      let completionColumn = sheet.getRange('K2:K').getValues().flat(); // Assuming column K has the completion dates
  
      let rowsToDelete = [];
      let lastRowId = 0;
  
      for (let i = 2; i <= archiveSheet.getLastRow(); i++) {
        let value = archiveSheet.getRange(i, 1).getValue();
        if (!isNaN(value) && value !== "") {
          lastRowId = Math.max(lastRowId, parseInt(value));
        }
      }
  
      for (let i = 0; i < data.length; i++) {
        let status = statusColumn[i];
        let completionDate = completionColumn[i];
  
        if (status && status == "Done" && completionDate) {
          let completedDate = new Date(completionDate);
          let diffDays = (now - completedDate) / (1000 * 60 * 60 * 24);
  
          if (diffDays >= 7) {
            lastRowId++;
            let rowToArchive = [lastRowId].concat(data[i]);
            archiveSheet.appendRow(rowToArchive);
            
            // Set the status cell with a dropdown list in column 'I'
            let appendedRowIndex = archiveSheet.getLastRow();
            let statusCell = archiveSheet.getRange(appendedRowIndex, 9); // Column 'I' is the 9th column
            let rule = SpreadsheetApp.newDataValidation().requireValueInList(['Done', 'Restore']).setAllowInvalid(false).build();
            statusCell.setDataValidation(rule);
            statusCell.setValue('Done'); // Set default value as 'Done'
            // Set the dropdown list for column 'F'
            let columnFCell = archiveSheet.getRange(appendedRowIndex, 6); // Column 'F' is the 6th column
            let columnFRule = SpreadsheetApp.newDataValidation().requireValueInList(['Low', 'Medium','High']).setAllowInvalid(false).build();
            columnFCell.setDataValidation(columnFRule);
            
            // Set the dropdown list for column 'G'
            let columnGCell = archiveSheet.getRange(appendedRowIndex, 7); // Column 'G' is the 7th column
            let columnGRule = SpreadsheetApp.newDataValidation().requireValueInList(['Not Urgent', 'Urgent']).setAllowInvalid(false).build();
            columnGCell.setDataValidation(columnGRule);
  
            rowsToDelete.push(i + 2);
          }
        }
      }
  
      rowsToDelete.reverse().forEach(row => {
        sheet.deleteRow(row);
      });
      //resettig the IDs
      let rows = sheet.getDataRange().getValues();
      for (let i = 0; i < rows.length; i++) {
        rows[i][0] = i; // Assuming the Task ID is in the first column
      }
      sheet.getDataRange().setValues(rows);
  
      
    }
  }
  
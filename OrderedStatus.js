function OrderedStatus(e) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = spreadsheet.getSheets();
    var editedCell = e.range;
    var statusOrder = {
    "Not Started": 1,
    "In Progress": 2,
    "On Hold": 3,
    "Done": 4
    };
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i]; 
      var range = sheet.getRange("B2:N");
      var data = range.getValues();
      if (editedCell.getSheet().getName() == sheet.getName()
      && editedCell.getColumn() == 9) {
        data.sort(function(a, b) {
        var statusA = a[7];
        var statusB = b[7];
        return (statusOrder[statusA] || 5) - (statusOrder[statusB] || 5); });
        range.setValues(data);
        Logger.log('Rows sorted by status in sheet: ' + sheet.getName());
      }
    }
  }
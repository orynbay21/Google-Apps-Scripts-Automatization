
//version 4 with the coloring
function RowAdded(e) {
    if (e.oldValue === "false") {
      var ui = SpreadsheetApp.getUi();
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
      // Get the edited row number
      var editedRow = e.range.getRow();
  
      // Get the email address from the same row (assuming email is in column N)
      var emailAddress = sheet.getRange(editedRow, 14).getValue(); // Column N
      var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
      if (emailAddress) { // Check if email address is not empty
        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        var rowValues = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  
        // Assign each value to a separate variable based on column index
        var RowID = rowValues[0]; // Column A
        var client = rowValues[1];
        var project = rowValues[2];
        var service = rowValues[3];
        var task = rowValues[4];
        var priority = rowValues[5];
        var urgency = rowValues[6];
        var dueDate = new Date(rowValues[7]); // Assuming due date is in column H (index 7)
        var formattedDueDate = Utilities.formatDate(dueDate, timeZone, 'dd/MM/yyyy');
        rowValues[7] = formattedDueDate; // Update the due date in the row values array
  
        var status = rowValues[8]; // Column I
  
        var assignedDate = new Date(rowValues[9]);
        var formattedAssignedDate = Utilities.formatDate(assignedDate, timeZone, 'dd/MM/yyyy');
        rowValues[9] = formattedAssignedDate;
  
        var completionDate = new Date(rowValues[10]);
        var formattedCompletionDate = Utilities.formatDate(completionDate, timeZone, 'dd/MM/yyyy');
        rowValues[10] = formattedCompletionDate; // Update the completion date in the row values array
  
        var comments = rowValues[11];
  
        // Add more variables as needed
  
        // Create the HTML body
        var htmlBody = headers.map(function(header, index) {
          var value = rowValues[index];
          var color = "";
          if (header === "Due Date") {
            var diffTime = dueDate - new Date();
            var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
            if (diffDays <= 3) {
              color = "red";
            } else if (diffDays <= 5) {
              color = "orange";
            } else if (diffDays <= 7) {
              color = "green";
            }
          }
          var valueSpan = color ? "<span style='color:" + color + ";'>" + value + "</span>" : value;
          return "<p><b>" + header + ":</b> " + valueSpan + "</p>";
        }).join("");
  
        var result = ui.alert('Send an email?', 'Yes or no', ui.ButtonSet.YES_NO);
  
        if (result == ui.Button.YES) {
          MailApp.sendEmail({
            to: emailAddress,
            subject: "To-Do List Update",
            htmlBody: "<p>The following task was updated:</p>" + htmlBody
          });
        }
      }
    }
  }
  



  //version 1 when everybody gets an email

// function RowAdded(e) {
//   if (e.oldValue === "false") {
//     var ui = SpreadsheetApp.getUi();
//     var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

//     // Get the edited row number
//     var editedRow = e.range.getRow();
//     // Get all headers (first row)
//     var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

//     // Get all values in the edited row
//     var rowValues = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];

//     // Create a formatted string with headers and corresponding values
//     var rowValuesString = headers.map(function(header, index) {
//       return header + ": " + rowValues[index];
//     }).join("\n");

//     var result = ui.alert('Send an email?', 'Yes or no', ui.ButtonSet.YES_NO);

//     if (result == ui.Button.YES) {
//       // Get all email addresses from the mailing list column (N)
//       var lastRow = sheet.getLastRow();
//       var emailAddresses = sheet.getRange(2, 14, lastRow - 1).getValues().flat(); // Assuming column N starts from row 2
      
//       // Send email to each address in the mailing list
//       emailAddresses.forEach(function(recipient) {
//         if (recipient) { // Check if recipient is not empty
//           MailApp.sendEmail({
//             to: recipient,
//             subject: "To-Do List Update",
//             body: "The following task was updated:\n\n" + rowValuesString
//           });
//         }
//       });
//     }
//   }
// }

//version 2 when only the responsible person gets an email

// function RowAdded(e) {
//   if (e.oldValue === "false") {
//     var ui = SpreadsheetApp.getUi();
//     var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

//     // Get the edited row number
//     var editedRow = e.range.getRow();

//     // Get the email address from the same row (assuming email is in column N)
//     var emailAddress = sheet.getRange(editedRow, 14).getValue(); // Assuming email is in column N

//     if (emailAddress) { // Check if email address is not empty
//       var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
//       var rowValues = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];

//       var rowValuesString = headers.map(function(header, index) {
//         return header + ": " + rowValues[index];
//       }).join("\n");

//       var result = ui.alert('Send an email?', 'Yes or no', ui.ButtonSet.YES_NO);

//       if (result == ui.Button.YES) {
//         MailApp.sendEmail({
//           to: emailAddress,
//           subject: "To-Do List Update",
//           body: "The following task was updated:\n\n" + rowValuesString
//         });
//       }
//     }
//   }
// }

// //version 3
// function RowAdded(e) {
//   if (e.oldValue === "false") {
//     var ui = SpreadsheetApp.getUi();
//     var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

//     // Get the edited row number
//     var editedRow = e.range.getRow();

//     // Get the email address from the same row (assuming email is in column N)
//     var emailAddress = sheet.getRange(editedRow, 14).getValue(); // Column N
//     var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
//     if (emailAddress) { // Check if email address is not empty
//       var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
//       var rowValues = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];

//       // Assign each value to a separate variable based on column index
//       var RowID= rowValues[0]; // Column A
//       var client=rowValues[1];
//       var project=rowValues[2];
//       var service=rowValues[3];
//       var task=rowValues[4];
//       var priority=rowValues[5];
//       var urgency=rowValues[6];
//       var dueDate = new Date(rowValues[7]); // Assuming due date is in column H (index 7)
//       var formattedDueDate = Utilities.formatDate(dueDate, timeZone, 'dd/MM/yyyy');
//       rowValues[7] = formattedDueDate; // Update the due date in the row values array

//       var status = rowValues[8]; // Column C

//       var assignedDate=new Date(rowValues[9]);
//       var formattedAssignedDate= Utilities.formatDate(assignedDate, timeZone, 'dd/MM/yyyy');
//       rowValues[9]=formattedAssignedDate;

//       var completionDate = new Date(rowValues[10]);
//       var formattedCompletionDate = Utilities.formatDate(completionDate, timeZone, 'dd/MM/yyyy');
//       rowValues[10] = formattedCompletionDate; // Update the due date in the row values array

//       var comments=rowValues[11];
//       // Add more variables as needed

//       var rowValuesString = headers.map(function(header, index) {
//         return header + ": " + rowValues[index];
//       }).join("\n");

//       var result = ui.alert('Send an email?', 'Yes or no', ui.ButtonSet.YES_NO);

//       if (result == ui.Button.YES) {
//         MailApp.sendEmail({
//           to: emailAddress,
//           subject: "To-Do List Update",
//           body: "The following task was updated:\n\n" + rowValuesString
//         });
//       }
//     }
//   }
// }
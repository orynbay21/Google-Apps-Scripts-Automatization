function TaskReminder() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = spreadsheet.getSheets(); 
    var timeZone = spreadsheet.getSpreadsheetTimeZone();
    var currentDate = new Date();
    var formattedCurrentDate = Utilities.formatDate(currentDate, "GMT+0", 'E MMM dd yyyy');
    var dayOfWeek = currentDate.getDay(); 
  
    var isBusinessDay = (dayOfWeek >= 1 && dayOfWeek <= 5);
    if (isBusinessDay) {
      sheets.forEach(function(sheet) {
        var dueColumnRange = sheet.getRange('H2:H');
        var dueColumn = dueColumnRange.getValues().flat(); 
        var clientColumn = sheet.getRange('B2:B').getValues().flat();
        var projectColumn = sheet.getRange('C2:C').getValues().flat();
        var serviceColumn = sheet.getRange('D2:D').getValues().flat();
        var statusColumn = sheet.getRange('I2:I').getValues().flat();
        var emails = sheet.getRange('N2:N').getValues().flat();
  
        var numRows = sheet.getLastRow() - 1;
        var tasks = [];
    
        for (var i = 0; i < numRows; i++) {
          var dueDate = new Date(dueColumn[i]);
          var diffTime = dueDate - currentDate;
          var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
          var needsReminder = (statusColumn[i] == 'In Progress' || statusColumn[i] == 'Not Started');
          if (diffDays <= 7 && needsReminder) {
            tasks.push({
              client: clientColumn[i],
              project: projectColumn[i],
              service: serviceColumn[i],
              status: statusColumn[i],
              diffDays: diffDays,
              dueDate: dueDate
            });
          }
        }
    
        var body = tasks.map((task, index) => {
          let color;
          if (task.diffDays <= 3) {
            color = "red";
          } else if (task.diffDays <= 5) {
            color = "orange";
          } else if (task.diffDays <= 7) {
            color = "green";
          }
          let indexText = `<p><b>${index + 1})</b></p>`;
          let clientText = `<p><b>Client:</b> ${task.client}</p>`;
          let projectText = `<p><b>Project:</b> ${task.project}</p>`;
          let serviceText = `<p><b>Service:</b> ${task.service}</p>`;
          let dueText = `<p><b>Due in:</b> ${task.diffDays} days</p>`;
          let statusText = `<p><b>Status:</b> ${task.status}</p>`;
          let dueDateText = `<p><b>Due Date:</b> <span style="color:${color};">${Utilities.formatDate(task.dueDate, timeZone, 'dd/MM/yyyy')}</span></p>`;
    
          let taskDescription = `${indexText}${clientText}${projectText}${serviceText}${dueText}${statusText}${dueDateText}`;
    
          return taskDescription;
        }).join('<br>');
    
        if (body.length > 0) {
          MailApp.sendEmail({
            to: '',
            bcc: emails.join(','),
            subject: 'Task Reminder for ' + formattedCurrentDate + " - " + sheet.getName(),
            htmlBody: body // Use htmlBody to send HTML formatted email
          });
        }
  
        // Apply conditional formatting to the Due Date column
        var rules = sheet.getConditionalFormatRules();
        var newRules = rules.filter(function(rule) {
          var ranges = rule.getRanges();
          return ranges.length > 0 && ranges[0].getA1Notation() !== 'H2:H';
        });
  
        newRules.push(SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=AND(H2<>"" , H2<=TODAY()+3)')
          .setBackground("#FF0000")
          .setRanges([dueColumnRange])
          .build());
  
        newRules.push(SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=AND(H2<>"" , H2<=TODAY()+5)')
          .setBackground("#FFA500")
          .setRanges([dueColumnRange])
          .build());
  
        newRules.push(SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=AND(H2<>"" , H2<=TODAY()+7)')
          .setBackground("#FFFF00")
          .setRanges([dueColumnRange])
          .build());
  
        sheet.setConditionalFormatRules(newRules);
      });
    }
  }
  
  
  
  // function TaskReminder() {
  //     var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //     var sheets = spreadsheet.getSheets(); 
  //     var timeZone = spreadsheet.getSpreadsheetTimeZone();
  //     var currentDate = new Date();
  //     var formattedCurrentDate = Utilities.formatDate(currentDate,
  //     "GMT+0", 'E MMM dd yyyy');
  //     var dayOfWeek = currentDate.getDay(); 
  
  //     var isBusinessDay = (dayOfWeek >= 1 && dayOfWeek <= 5);
  //     if (isBusinessDay){
  //       sheets.forEach(function(sheet) {
  //       var dueColumn = sheet.getRange('H2:H').getValues().flat(); 
  //       var clientColumn = sheet.getRange('B2:B').getValues().flat();
  //       var projectColumn = sheet.getRange('C2:C').getValues().flat();
  //       var serviceColumn = sheet.getRange('D2:D').getValues().flat();
  //       var statusColumn = sheet.getRange('I2:I').getValues().flat();
  //       var emails = sheet.getRange('N2:N').getValues().flat();
  
  //       var numRows = sheet.getLastRow() - 1;
  //       var tasks = [];
    
  //       for (var i = 0; i < numRows; i++) {
  //         var dueDate = new Date(dueColumn[i]);
    
  //         var diffTime = dueDate - currentDate;
  //         var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
  //         var needsReminder = (statusColumn[i] == 'In Progress' || 
  //         statusColumn[i] == 'Not Started');
  //         if (diffDays <= 7 && needsReminder) {
  //           tasks.push({
  //             client: clientColumn[i],
  //             project: projectColumn[i],
  //             service: serviceColumn[i],
  //             status: statusColumn[i],
  //             diffDays: diffDays,
  //             dueDate: dueDate
  //           });
  //         }
  //       }
    
  //       var body = tasks.map((task, index) => {
  //         let indexText = `${index + 1})\n`;
  //         let clientText = `Client: ${task.client}\n`;
  //         let projectText = `Project: ${task.project}\n`;
  //         let serviceText = `Service: ${task.service}\n`;
  //         let dueText = `Due in: ${task.diffDays} days\n`;
  //         let statusText = `Status: ${task.status}\n`;
  //         let dueDateText = `Due Date: ${Utilities.formatDate(task.dueDate, timeZone, 'dd/MM/yyyy')}\n`;
    
  //         let taskDescription = `${indexText}${clientText}${projectText}${serviceText}${dueText}${statusText}${dueDateText}`;
    
  //         return taskDescription;
  //       }).join('\n');
    
    
  //       if (body.length > 0) {
  //         MailApp.sendEmail({
  //           to: '',
  //           bcc: emails.join(','),
  //           subject: 'Task Reminder for ' + formattedCurrentDate + " - " + sheet.getName(),
  //           body: body
  //         });
  //       }
  //     });
  
  //     }
  //   }
  
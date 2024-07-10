function onOpen() {
    let ui = SpreadsheetApp.getUi();
  
    ui.createMenu('Script')
      .addItem('Move to Archive', 'MoveToArchive')
      .addItem('Task Reminder', 'TaskReminder')
      .addItem('Order Statuses', 'OrderedStatus')
      .addItem('Row Added or Updated', 'RowAdded')
      .addItem('Restore', 'Restore')
      .addToUi();
  }
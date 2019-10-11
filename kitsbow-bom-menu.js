function onOpen() {
  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();

  if(ssId == spreadsheetIdShoppingMaster) {
    SpreadsheetApp.getUi()
    .createMenu('Kitsbow')
    .addItem('Clone Shopping List', 'createShoppingList')
    .addItem('Open Google Drive', 'openGoogleDriveFolder')
    .addToUi();
  
  }
  else {
    SpreadsheetApp.getUi()
    .createMenu('Kitsbow')
    .addItem('Update Shopping List', 'createShoppingList')
    .addSeparator()
    .addItem('Open Google Drive', 'openGoogleDriveFolder')
    .addSeparator()
    .addItem('Fetch Materials Reference', 'cloneColors')
    .addItem('Fetch BOM Master', 'cloneColors')
    .addToUi();
  }
};

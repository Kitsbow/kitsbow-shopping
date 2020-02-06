function onOpen() {
  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();

  if(ssId == spreadsheetIdShoppingMaster) {
    SpreadsheetApp.getUi()
    .createMenu('Kitsbow')
    .addItem('Clone Shopping List', 'cloneShoppingList')
    .addItem('Clone CMT List', 'cloneCmtList')
    .addSeparator()
    .addItem('Open Google Drive', 'openGoogleDriveFolder')
    .addToUi();
  
  }
  else {
    SpreadsheetApp.getUi()
    .createMenu('Kitsbow')
    .addItem('Update from SKUs Order', 'determineTypeUpdateListTable')
    .addSeparator()
    .addItem('Open Google Drive', 'openGoogleDriveFolder')
    .addSeparator()
    .addItem('Restore Materials Reference', 'fetchMaterialsReference')
    .addItem('Restore SKU Master', 'fetchSkuMaster')
    //.addItem('Restore BOM Master', 'fetchBomMaster')
    .addToUi();
  }
};

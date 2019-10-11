var sheetNameBomData = '+BOM Data';
var masterIdBom = '1xFYUc59__4-WzmqWKVzdiHBmZawH-VNMyHFWyRPbVJw';
var masterSheetBom = 'BOM Rows';
var colBomSku = 0;
var colBomKpn = 1;
var colBomUsage = 4;

var sheetNameShoppingList = 'SKUs Order';
var colShoppingListSku = 0;
var colShoppingListQty = 1;

var sheetNameOutput = 'Output';
var colOutputKpn = 0;
var colOutputQty = 1;

var sheetNameMaterials = '+Materials Reference';
var masterIdMaterials = '1CFJTacS-FKNwYV31PoX4nH7nlRiUfPAAW2vKfEgdtw4';
var masterSheetMaterials = 'MASTER Materials Reference';

var sheetNameSkuMaster = '+SKU Master v2';
var masterIdSku = '1M-gjTZSBcIwOVnVtUF6C0rjKsTWEV3rM15gewKhCaOg';
var masterSheetSku = 'SKUs';

function updateShoppingList(spreadsheet) {
  if (typeof(spreadsheet) === 'undefined') {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  var sheetShoppingList = spreadsheet.getSheetByName(sheetNameShoppingList);
  var shoppingListData = sheetShoppingList.getDataRange().getValues();
  
  var input = { };
  var output = { };

  // step through all lines in the shopping list to create the shoppingList object
  for(var i = 1; i < shoppingListData.length; i++ ) {
    if(shoppingListData[i][colShoppingListSku].length) {
      input[shoppingListData[i][colShoppingListSku]] = shoppingListData[i][colShoppingListQty];
    }
  }
  
  var sheetBomData = spreadsheet.getSheetByName(sheetNameBomData);
  var bomData = sheetBomData.getDataRange().getValues();
  
  // step through all lines in the bom database
  for(var i = 1; i < bomData.length; i++ ) {
    if(bomData[i][colBomSku].length && input.hasOwnProperty(bomData[i][colBomSku])) {
      if(!output.hasOwnProperty(bomData[i][colBomKpn])){
        output[bomData[i][colBomKpn]] = 0;
      }
      
      output[bomData[i][colBomKpn]] += input[bomData[i][colBomSku]]* bomData[i][colBomUsage];   
    }
  }
  
  var sheetOutput = spreadsheet.getSheetByName(sheetNameOutput);
  // clear the output range
  sheetOutput.getRange(2, colOutputKpn + 1, sheetOutput.getMaxRows() - 1, colOutputQty + 1).clearContent();
  var outputValues = [ ];
  
  // render the object to an array
  Object.keys(output).forEach(function(key) { outputValues.push([ key, output[key] ]);});
  
  // write the values to the output sheet
  outputValues.sort();
  sheetOutput.getRange(2, 1, outputValues.length, outputValues[0].length).setValues(outputValues);
}

function cloneShoppingList() {
  var sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var outputFolder = DriveApp.getFolderById(driveIdShoppingFolder);
    
  // copy the entire spreadsheet
  var destinationSpreadsheet = SpreadsheetApp.open(
    DriveApp.getFileById(sourceSpreadsheet.getId()).makeCopy(createFilename(), outputFolder))  

  // overwrite IMPORTRANGE()-driven sheets with values
  overwriteWithValues(sourceSpreadsheet, destinationSpreadsheet, sheetNameBomData);
  overwriteWithValues(sourceSpreadsheet, destinationSpreadsheet, sheetNameMaterials);
  overwriteWithValues(sourceSpreadsheet, destinationSpreadsheet, sheetNameSkuMaster);

  // delete a few sheets...
  var sheetsToDelete = ['Documentation'];
  sheetsToDelete.forEach( function (sheetName) {
    var sheetToDelete = destinationSpreadsheet.getSheetByName(sheetName);

    if(sheetToDelete) {
      destinationSpreadsheet.deleteSheet(sheetToDelete);
    }
  });

  // make sure the output sheet isn't hidden
  destinationSpreadsheet.getSheetByName(sheetNameOutput).showSheet();
  // build shopping list in destination sheet
  updateShoppingList(destinationSpreadsheet);
  
  // open the new spreadsheet
  openUrl('https://docs.google.com/spreadsheets/d/'+destinationSpreadsheet.getId(), 
    'Opening \''+createFilename()+'\'');
}

function getSkusFromShoppingList() {
  var skus = [ ];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetNameShoppingList);
  var listData = sheet.getDataRange().getValues();

  for(var i=1; i < listData.length; i++) {
    var sku = listData[i][colShoppingListSku];
    if(validateSku(sku)){
      if(skus.indexOf(styleFromSku(sku)) == -1) {
        skus.push(styleFromSku(sku));
      }
    }
  }

  return skus;
}

function createFilename() {
  return getSkusFromShoppingList().join('_')+'_Order';
}

function fetchMaterialsReference() {
  updateSheetFromReference(masterIdMaterials, masterSheetMaterials, sheetNameMaterials);
}

function fetchSkuMaster() {
  updateSheetFromReference(masterIdSku, masterSheetSku, sheetNameSkuMaster);
}

function fetchBomMaster() {
  updateSheetFromReference(masterIdBom, masterSheetBom, sheetNameBomData);
}

function updateSheetFromReference(referenceId, referenceSheetName, localSheetName) {
  var ui = SpreadsheetApp.getUi();
  
  var master = SpreadsheetApp.openById(referenceId);

  if(!master) {
    ui.prompt('Update Failed for \''+localSheetName+'\'',
      'Could not load master spreadsheet.', ui.ButtonSet.OK);
    return;
  }

  var masterSheet = master.getSheetByName(referenceSheetName);

  if(!masterSheet) {
    ui.prompt('Update Failed for \''+localSheetName+'\'',
      'Could not load master sheet.', ui.ButtonSet.OK);
    return;
  }

  var localSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var localSheet = localSpreadsheet.getSheetByName(localSheetName);
  var data = masterSheet.getDataRange().getValues();

  // overwrite the local sheet
  localSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}


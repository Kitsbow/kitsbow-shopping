var sheetNameShoppingList = 'SKUs Order';
var colShoppingListSku = 0;
var colShoppingListQty = 1;

var sheetNameMasterPrefix = 'Master - ';
var sheetNameShopping = 'Shopping List';
var sheetNameCmt = 'CMT Kitting List';
var colOutputKpn = 0;
var colOutputQty = 1;

var typeShopping = 0;
var typeCmt = 1;

var cmtConfig = { 
  template: sheetNameMasterPrefix + sheetNameCmt, 
  sheetName: sheetNameCmt,
  type: typeCmt, 
  sheetsToDelete: ['Documentation', sheetNameMasterPrefix + sheetNameShopping],
  colPartType: 4,
  colVendorName: 5
};

var shopConfig = { 
  template: sheetNameMasterPrefix + sheetNameShopping, 
  sheetName: sheetNameShopping,
  type: typeShopping, 
  sheetsToDelete: ['Documentation', sheetNameMasterPrefix + sheetNameCmt],
  colPartType: 6,
  colVendorName: 7
};

var sheetNameMaterials = '+Materials Reference';
var masterIdMaterials = '1CFJTacS-FKNwYV31PoX4nH7nlRiUfPAAW2vKfEgdtw4';
var masterSheetMaterials = 'MASTER Materials Reference';

var sheetNameSkuMaster = '+SKU Master v2';
var masterIdSku = '1M-gjTZSBcIwOVnVtUF6C0rjKsTWEV3rM15gewKhCaOg';
var masterSheetSku = 'SKUs';

function determineTypeUpdateListTable() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if(spreadsheet.getSheetByName(sheetNameShopping)){
    updateListTable(spreadsheet, shopConfig);
  }
  else if(spreadsheet.getSheetByName(sheetNameCmt)){
    updateListTable(spreadsheet, cmtConfig);
  }
  else {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Could not determine spreadshseet type.', ui.ButtonSet.OK);
  }
}

function updateListTable(spreadsheet, config) {
  if (typeof(spreadsheet) === 'undefined') {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  var sheetList = spreadsheet.getSheetByName(sheetNameShoppingList);
  var listData = sheetList.getDataRange().getValues();
  
  var input = { };
  var output = { };
  var listSkuArr = [];
  var qaReport = "";

  // step through all lines in the shopping list to create the shoppingList object
  for(var i = 1; i < listData.length; i++ ) {
    if(listData[i][colShoppingListSku].length) {
      if (input[listData[i][colShoppingListSku]]) {
        qaReport += '\nWarning: The SKU ' + listData[i][colShoppingListSku] + 
        ' appears more than once in this order. Parts will only be tallied for the last count.';
      }
      input[listData[i][colShoppingListSku]] = listData[i][colShoppingListQty];
      listSkuArr.push(listData[i][colShoppingListSku]);

    }
  }
  
  var bomData = getBomDataForSkus(listSkuArr);
  const KEY_FOUND_POSTFIX = '_found';
  var bomRecordsFound = 0;
  // step through all lines in the bom database
  for(var i = 0; i < bomData.length; i++ ) {
    var nSKU = bomData[i]['sku'];
    var nKPN = bomData[i]['kpn'];
    var nUsage = bomData[i]['part_usage'];
    if(nSKU && 0 < input[nSKU]) {
      if( undefined === output[nKPN] || null === output[nKPN]){
        output[nKPN] = 0;
      }
      output[nKPN] += input[nSKU]* nUsage;   
      input[nSKU + KEY_FOUND_POSTFIX] = 1;
      bomRecordsFound++;
    } else {
      qaReport += '\nNote: missing/invalid count for SKU ' + nSKU;
    }
  }
  if (bomData.length > bomRecordsFound) {
    qaReport += '\nWarning: mismatched count in totals for these SKUs';
  }

  var sheetOutput = spreadsheet.getSheetByName(config.sheetName);
  // clear the output range
  sheetOutput.getRange(2, colOutputKpn + 1, sheetOutput.getMaxRows() - 1, colOutputQty + 1).clearContent();
  var outputValues = [ ];
  
  // render the object to an array
  Object.keys(output).forEach(function(key) { outputValues.push([ key, output[key] ]);});

  if (0 == outputValues.length) {
    outputValues.push(["No Skus were matched in BOM"]);
  }
  // write sorted values to the output sheet in order of part type then vendor name
  outputValues.sort();
  sheetOutput.getRange(2, 1, outputValues.length, outputValues[0].length).setValues(outputValues);
  sheetOutput.getRange(2, 1, outputValues.length, sheetOutput.getLastColumn())
    .sort([{column: config.colPartType, ascending: true}, {column: config.colVendorName, ascending: true}]);

    listSkuArr.forEach(function(skuKey) {
      if ( input[skuKey] && ! input[skuKey + KEY_FOUND_POSTFIX] ) {
        qaReport += '\nWarning: BOM table is missing SKU: ' + skuKey;
      }
    });

    if (qaReport) {
      var ui = SpreadsheetApp.getUi();
      ui.alert( "QA Check for " + spreadsheet.getName() + "\n" + qaReport);
    }
}

function cloneCmtList() {
  aggregateOrderList(cmtConfig);
}

function cloneShoppingList() {
  aggregateOrderList(shopConfig);
}

function aggregateOrderList(config) {
  var sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var outputFolder = DriveApp.getFolderById(driveIdShoppingFolder);
    
  // copy the entire spreadsheet
  var destinationSpreadsheet = SpreadsheetApp.open(
    DriveApp.getFileById(sourceSpreadsheet.getId()).makeCopy(createFilename(config.type), outputFolder))  

  // overwrite IMPORTRANGE()-driven sheets with values
  overwriteWithValues(sourceSpreadsheet, destinationSpreadsheet, sheetNameMaterials);
  overwriteWithValues(sourceSpreadsheet, destinationSpreadsheet, sheetNameSkuMaster);

  // delete a few sheets...
  config.sheetsToDelete.forEach( function (sheetName) {
    var sheetToDelete = destinationSpreadsheet.getSheetByName(sheetName);

    if(sheetToDelete) {
      destinationSpreadsheet.deleteSheet(sheetToDelete);
    }
  });

  // make sure the output sheet isn't hidden
  destinationSpreadsheet.getSheetByName(config.template).showSheet();
  // rename the output sheet
  destinationSpreadsheet.getSheetByName(config.template).setName(config.sheetName);
  // build shopping list in destination sheet
  updateListTable(destinationSpreadsheet, config);
  
  // open the new spreadsheet
  openUrl('https://docs.google.com/spreadsheets/d/'+destinationSpreadsheet.getId(), 
    'Opening \''+destinationSpreadsheet.getName()+'\'');
}

function getStylesFromShoppingList() {
  var styles = [ ];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetNameShoppingList);
  var listData = sheet.getDataRange().getValues();

  for(var i = 1; i < listData.length; i++) {
    var sku = listData[i][colShoppingListSku];
    if(validateSku(sku)){
      if(styles.indexOf(styleFromSku(sku)) == -1) {
        styles.push(styleFromSku(sku));
      }
    }
  }

  return styles;
}

function createFilename(outputType) {
  var skus = getStylesFromShoppingList();
  var today = new Date();
  return 'Style'+ (skus.length==1? ' ':'s ')+skus.join('_')+' '+
    (outputType == typeShopping? 'Shopping v2':'CMT Kitting v2')+' List '+
    today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
}

function fetchMaterialsReference() {
  updateSheetFromReference(masterIdMaterials, masterSheetMaterials, sheetNameMaterials);
}

function fetchSkuMaster() {
  updateSheetFromReference(masterIdSku, masterSheetSku, sheetNameSkuMaster);
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

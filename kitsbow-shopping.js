var sheetNameBomData = '+BOM Data';
var colBomSku = 0;
var colBomKpn = 1;
var colBomUsage = 4;

var sheetNameShoppingList = 'SKUs Order';
var colShoppingListSku = 0;
var colShoppingListQty = 1;

var sheetNameOutput = 'Output';
var colOutputKpn = 0;
var colOutputQty = 1;

function createShoppingList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheetShoppingList = ss.getSheetByName(sheetNameShoppingList);
  var shoppingListData = sheetShoppingList.getDataRange().getValues();
  
  var input = { };
  var output = { };

  // step through all lines in the shopping list to create the shoppingList object
  for(var i = 1; i < shoppingListData.length; i++ ) {
    if(shoppingListData[i][colShoppingListSku].length) {
      input[shoppingListData[i][colShoppingListSku]] = shoppingListData[i][colShoppingListQty];
    }
  }
  
  var sheetBomData = ss.getSheetByName(sheetNameBomData);
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
  
  var sheetOutput = ss.getSheetByName(sheetNameOutput);
  // clear the output range
  sheetOutput.getRange(2, colOutputKpn + 1, sheetOutput.getMaxRows() - 1, colOutputQty + 1).clearContent();
  var outputValues = [ ];
  
  // render the object to an array
  Object.keys(output).forEach(function(key) { outputValues.push([ key, output[key] ]);});
  
  // write the values to the output sheet
  outputValues.sort();
  sheetOutput.getRange(2, 1, outputValues.length, outputValues[0].length).setValues(outputValues);
}

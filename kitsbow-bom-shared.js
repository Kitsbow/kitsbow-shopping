var spreadsheetIdShoppingMaster = '1CMJJKaB8QJISD0uLiQo2qqnV8OlbntyGJeIYI1xNK0k';
var driveIdShoppingFolder = '1LTbaJD6bnZSpO3pGJ_Y3aLhORzoI97Yg';

function openGoogleDriveFolder() {
    openUrl('https://drive.google.com/drive/folders/'+driveIdShoppingFolder,
        'Opening Google Drive');
}

function openUrl(url, title) {
    var html = HtmlService.createHtmlOutput(
        '<html><script>window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};\
        var a = document.createElement(\'a\'); a.href=\''+url+'\'; a.target=\'_blank\';\
        if(document.createEvent){\
        var event=document.createEvent(\'MouseEvents\');\
        if(navigator.userAgent.toLowerCase().indexOf(\'firefox\')>-1){window.document.body.append(a)}\
        event.initEvent(\'click\',true,true); a.dispatchEvent(event);\
        }else{ a.click() }\
        close();\
        </script>\
        <body style=\'word-break:break-word;font-family:sans-serif;\'>Failed to open automatically. <a href=\''+url+'\' target=\'_blank\' onclick=\'window.close()\'>Click here to proceed</a>.</body>\
        <script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>\
        </html>')
    .setWidth( 150 ).setHeight( 10 );
    SpreadsheetApp.getUi().showModalDialog( html, title );
}

function overwriteWithValues(sourceSpreadsheet, destinationSpreadsheet, sheetName) {
    var sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
    var destinationSheet = destinationSpreadsheet.getSheetByName(sheetName);
    var sourceValues = sourceSheet.getRange(1,1,sourceSheet.getMaxRows(),sourceSheet.getMaxColumns()).getValues();
    destinationSheet.getRange(1, 1, destinationSheet.getMaxRows(), destinationSheet.getMaxColumns()).setValues(sourceValues);
  }

function validateSku(skuText) {
    // check if the sku is 12 characters long
    if(skuText.length != 12) {
        return false;
    }

    // check if hyphen separators are in the right place
    if(skuText[4]!='-' || skuText[8]!='-') {
        return false;
    }

    // check if the sku contains 10 digit characters
    if(skuText.match( /\d+/g ).join('').length != 10) {
        return false;
    }
    
    return true;
}

function styleFromSku(skuText) {
    return skuText.substring(0,4);
}

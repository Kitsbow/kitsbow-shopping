var spreadsheetIdShoppingMaster = '1CMJJKaB8QJISD0uLiQo2qqnV8OlbntyGJeIYI1xNK0k';
var driveIdShoppingFolder = '1LTbaJD6bnZSpO3pGJ_Y3aLhORzoI97Yg';
//BigQuery connection info
var bomDataSetId = "production_data";
var bomTableName = "bom";
var bomProjectId = "kitsbow-bom-database";

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

/**
 * Return data from the BOM table as an object array [{colname: value, ...},
 * Current columns are:
 * SKU	STRING	REQUIRED	
 * KPN	STRING	REQUIRED	
 * Notes	STRING	NULLABLE	
 * Placement	STRING	NULLABLE	
 * Usage	FLOAT	NULLABLE	
 * Section	STRING	NULLABLE	
 * Uploaded	DATE	REQUIRED	
 * LastUpdate	TIMESTAMP	NULLABLE	
 * 
 * In this version, the returned values are all of type String.
 * @param {Array} skus String sku list as "NNNN-NNN-NNN"
 */
function getBomDataForSkus(skus) {
    var querySql = "SELECT * FROM `" + bomDataSetId + "." + bomTableName + "` WHERE sku IN (";
    for (var i = 0; i < skus.length; i++) {
      querySql += " '" + skus[i] + "'";
      if (i < skus.length-1) {
        querySql += ",";
      }
    }
    querySql += ")";
  
    var resultsJson = sendSQLQuery(querySql, {type: "BigQuery", projectId: bomProjectId}, "json");
    return resultsJson;
  }
  
/**
 * send SQL query to a database described in 'connection' and return results array
 * @param {String} query SQL Query string 
 * @param {Object} connection contains connection info, including "type" which currently must be "BigQuery"
 * @param {String} returnType "json" for an array of objects
 */
function sendSQLQuery(query, connection, returnType) {
    if (connection.type == "BigQuery") {
        var queryRequest = {
            kind: "json",
            query: query,
            useLegacySql: false
        }
        var queryResults = BigQuery.Jobs.query(queryRequest, connection.projectId);
        var jobId = queryResults.jobReference.jobId;
      
        // repeat check for completed job
        var sleepTimeMs = 500;
        while (!queryResults.jobComplete) {
            Utilities.sleep(sleepTimeMs);
            sleepTimeMs *= 2;
            queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
            //Logger.log("query loop " + queryResults.jobComplete);
        }
      
        // Get all the rows of results.
        var rows = queryResults.rows;
        while (queryResults.pageToken) {
                queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {
                pageToken: queryResults.pageToken
            });
            rows = rows.concat(queryResults.rows);
        }
          //response.rows is an array of objects, with each representing a query result row:
          // f:[ {v: value}, ...]
          // values come from JSON as strings, so have to be coerced to JS types
          // types can be gotten from the 'schema' property of the queryResults
        if (rows && 0 < rows.length) {
            if (returnType == "json") {
                var result = [];
                var fields = queryResults.schema.fields;
                rows.forEach( function(row, i) {
                    var retRowObj = {};
                    row.f.forEach( function (col, j) {
                        retRowObj[fields[j].name] = col.v;
                    });
                    result.push(retRowObj);
                });
                return result;
            } else {
                throw new Error("unrecognized return type: " + returnType);
            }
        } else {
            return [];
        }
    } else {
        throw new Error("unrecognized connection type: " + connection.type);
    }
}

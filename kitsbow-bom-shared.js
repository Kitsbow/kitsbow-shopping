const spreadsheetIdShoppingMaster = '1CMJJKaB8QJISD0uLiQo2qqnV8OlbntyGJeIYI1xNK0k';
const driveIdShoppingFolder = '1LTbaJD6bnZSpO3pGJ_Y3aLhORzoI97Yg';
//BigQuery connection info
const bomDataSetId = "production_data";
const bomTableName = "bom";
const bomProjectId = "kitsbow-bom-database";
//settings worksheet for cloud-sql connection info
const workflowDatabaseConfigId = '16hIeTu9DJOkwPEMhG58LRDwzDl5QdfAeSbYhKwRwV00';
const workflowConfigSheetName = 'Workflow Configuration';

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
   * @param {String} bomProjectId from BigQuery, if needed
   * @param {String} bomDataSetId, if needed
   * @param {String} tableName
   */
  function getBomDataForSkus(skus, bomProjectId, bomDataSetId, tableName) {
    let connectionObj = createCloudJdbcConnection();
    if (!tableName) { 
        tableName = bomTableName; 
    }
    let querySql = "SELECT * FROM " + connectionObj.config['DB_Table_Qualifier'] + tableName + " WHERE sku IN (";
    for (let i = 0; i < skus.length; i++) {
      querySql += " '" + skus[i] + "'";
      if (i < skus.length-1) {
        querySql += ",";
      }
    }
    querySql += ")";
    connectionObj.type = "CloudSQL";
    let resultsJson = sendSQLQuery(querySql, connectionObj, "json");
    connectionObj.connection.close();
    return resultsJson;
  }


  /**
  * send SQL query to a database described in 'connection' and return results array
  * @param {String} query SQL Query string 
  * @param {Object} config contains connection info, including "type" which currently must be either "BigQuery" or "CloudSQL"
  * @param {String} returnType "json" for an array of objects
  */
  function sendSQLQuery(query, config, returnType) {
    if (config.type == "BigQuery") {
        var queryRequest = {
            kind: "json",
            query: query,
            useLegacySql: false
        }
        var queryResults = BigQuery.Jobs.query(queryRequest, config.projectId);
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
    } else if (config.type == 'CloudSQL') {
      if (! config.connection) {
        throw new Error("cloudSQL db connection must be set on config object");
      }
      try {
        var conn = config.connection;
        var resultSet = conn.createStatement().executeQuery(query),
          metaData = resultSet.getMetaData(),
          colCount = metaData.getColumnCount(),
          columnNames = [],
          results = [];
        for (var index = 1; index <= colCount; index++) {
          columnNames.push(metaData.getColumnName(index));
        }

        while (resultSet.next()) {
          var rowObj = {};
          columnNames.forEach( 
            function(column) {
              var value = resultSet.getObject(column);
              rowObj[column] = value;
            }
          )
          results.push(rowObj);
        }
        resultSet.close();
        return results;
      } catch (err) {
        Logger.log(err);
        throw new Error("CloudSQL query " + query + " failed with an error: " + err.stack);
      }
    } else {
        throw new Error("unrecognized connection type: " + config.type);
    }
  }
  
  /**
   * Do string, and other, escaping. This should work for both BigQuery and CloudSQL/MySQL database types.
   * @param {any} value 
   */
  function escapeSQL(value) {
    if (typeof value === 'undefined') {
      return "null";
    }
    if (typeof value === 'string') {
      var escapedValue = value.replace(/'/g, "\\\'");
      escapedValue = escapedValue.replace(/`/g, "\\`");
      escapedValue = escapedValue.replace(/\n/g, "\\n");
      escapedValue = escapedValue.replace(/\r/g, "\\r");
      return "'"+escapedValue+"'";
    }
    return value;
  }


  function getMySQLCloudJdbcConnection(cloudSQLidentifier, databaseName, userName, password) {
    let instanceUrl = 'jdbc:google:mysql://' + cloudSQLidentifier + '/' + databaseName;
    let conn = Jdbc.getCloudSqlConnection(instanceUrl, userName, password);
    return conn;
  }

  /**
   * Using JDBC database connection info from a reference sheet with the properties set,
   * create a JDBC database connection to the Google Cloud SQL database.
   */
  function createCloudJdbcConnection() {
    let config = readExternalConfig();
    let cloudProjectId = config.Cloud_SQL_Project || 'CloudProjectNotSet',
      databaseName = config.Cloud_SQL_Db || 'missing db name',
      userName = config.Cloud_SQL_User,
      password = config.Cloud_SQL_Pass;
    let connection = getMySQLCloudJdbcConnection(cloudProjectId, databaseName, userName, password);
    if (! connection || connection.isClosed()) {
      throw new Error('Unable to get database connection from project ' + cloudProjectId);
    }
    return {
      connection,
      config
    };
  }

  
/**
 * Read properties (key-value pairs in first 2 columns) from external worksheet.
 * @param {String} worksheetId Google worksheet ID, which must contain a sheet named "Workflow Configuration"
 * @returns a configuration object set from the external worksheet
 */
function readExternalConfig(worksheetId = workflowDatabaseConfigId) {
  let configData = {};
  var configWorksheet = SpreadsheetApp.openById(worksheetId);
  var configSheet = configWorksheet.getSheetByName(workflowConfigSheetName);
  if (!configSheet) {
    throw new Error("Invalid configuration worksheet in readExternalConfig");
  }
  var configVals = configSheet.getRange(1,1,39,2).getValues();
  configVals.forEach( 
    function(elem) {
      if (elem[0])
        configData[elem[0]] = elem[1];
    }
  )
  return configData;
}


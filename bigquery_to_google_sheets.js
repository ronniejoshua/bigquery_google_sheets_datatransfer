/*
Google Sheet Format
+--------------------+---------------------+-------------------+
|      Tab Name      |        Query        |    Description    |
+--------------------+---------------------+-------------------+
| my_gsheet_tab_name | starndard_sql_query | query_description |
+--------------------+---------------------+-------------------+
 */ 

/* BigQuery Project ID */
PROJECT_ID = 'gcp-bigquery-project-id';

function onOpen(e) {
    var menu = SpreadsheetApp.getUi().createMenu("Read BQ");
    menu.addSeparator()
    menu.addItem("Run Query", 'run_query');

    menu.addToUi();
}

/**
 * Query data from BigQuery and save it into the spreadsheet
 * @param {string} query The query 
 * @param {string} tab_name The name of the spreadsheet tab where the data will be saved.
 */
function save_into_spreadsheet(query, tab_name) {
    // Query the data from BigQuery
    var request = {
        query: query
    };
    var queryResults = BigQuery.Jobs.query(request, PROJECT_ID);

    var jobId = queryResults.jobReference.jobId;

    // Check on status of the Query Job.
    var sleepTimeMs = 500;
    while (!queryResults.jobComplete) {
        Utilities.sleep(sleepTimeMs);
        sleepTimeMs *= 2;
        queryResults = BigQuery.Jobs.getQueryResults(PROJECT_ID, jobId);
    }

    // Get all the rows of results.
    var rows = queryResults.rows;
    while (queryResults.pageToken) {
        queryResults = BigQuery.Jobs.getQueryResults(PROJECT_ID, jobId, {
            pageToken: queryResults.pageToken
        });
        rows = rows.concat(queryResults.rows);
    }

    var values = []; 
    var tRow = [];
    var fields = queryResults.schema.fields;

    // Build Header
    for (var i in fields) {
        tRow.push(fields[i].name)
    }
    values.push(tRow)

    // Data
    for (var i in rows) {
        var row = rows[i].f, tRow = [];
        for (var f in row) {
            tRow.push(row[f].v);
        }
        values.push(tRow)
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tab_name);
    if (!sheet) { // Tab doesn't exist. Creates a new one
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(tab_name)
    }

    // Crear the tab and past the new data
    sheet.getDataRange().clearContent()
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
}

/**
 * Gets the information about the query to run and fire the query
 */
function run_query() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Settings")
    var num_rows = sheet.getLastRow() - 1;
    Logger.log(num_rows)
    var query_data = sheet.getRange(2, 1, num_rows, 2).getValues();
  for (var i in query_data){
    Logger.log(i)
    var query_name = query_data[i][0];
    var query = query_data[i][1];
    Logger.log(query_name)
    Logger.log(query)
    save_into_spreadsheet(query, query_name)
  }
}

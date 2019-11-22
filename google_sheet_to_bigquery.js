function onOpen() {
    "use strict";
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Write BQ")
        .addItem("Upload", "menuFunction")
        .addToUi();
}

function menuFunction() {
    uploadToBigQuery("my-google-spreadsheet-id");
}


function uploadToBigQuery(my_spreadsheet_id) {
    var upload_time = Utilities.formatDate(new Date(), Session.getTimeZone(), "yyyy-MM-dd HH:mm:ss");
    console.log(upload_time);
    console.log(my_spreadsheet_id);
    var ss = SpreadsheetApp.openById(my_spreadsheet_id);
    var sheets = ss.getSheets();
    //var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    if (sheets.length > 1) {
        console.log(sheets[1].getName());
        console.log(sheets.length);
    }

    for (var i = 0; i < sheets.length; i++) {
        console.log("Processing Sheet: " + sheets[i].getName());
        # Filtering Sheets
        if (sheets[i].getName() != "Settings" && sheets[i].getName() != "Configuration") {
            upload_data(my_spreadsheet_id, sheets[i].getName(), upload_time);
        }
    }
}

function upload_data(my_spreadsheet_id, reqsheet, upload_time) {
    // Define the configurations
    console.log("Start")
    var projectId = 'my-gcp-project-id';
    var datasetId = 'my-bq-dataset';
    var tableId = "my-bq-table-id";
    var table = {
        tableReference: {
            projectId: projectId,
            datasetId: datasetId,
            tableId: tableId
        },
        schema: {
            fields: [
                { name: '_insert_time', type: 'TIMESTAMP' },
                { name: 'date', type: 'DATE' }
            ]
        },
        timePartitioning: {
            type: 'DAY',
            field: '_insert_time'
        }
    };


    if (tableExists(tableId, projectId, datasetId)) {
        console.log("Table Exists")
        var rowsCSV = readData(my_spreadsheet_id, reqsheet, upload_time);
    } else {
        Logger.log("Table Doesn't Exists")
        // BigQuery.Tables.insert also inserts/creates a table
        table = BigQuery.Tables.insert(table, projectId, datasetId);
        console.log('Table created: %s', table.id);
        var rowsCSV = readData(my_spreadsheet_id, reqsheet, upload_time);
    }

    // General - Writing to Big Query
    console.log("Writing to Big Query")
    // Creating a Blob
    var data = Utilities.newBlob(rowsCSV, 'application/octet-stream');
    console.log(data);

    // Create the data upload job.
    var job = {
        configuration: {
            load: {
                destinationTable: {
                    projectId: projectId,
                    datasetId: datasetId,
                    tableId: tableId
                },
                skipLeadingRows: 0
            }
        }
    }

    job = BigQuery.Jobs.insert(job, projectId, data);
    console.log('Load job started. Check on the status of it here: ' + 'https://bigquery.cloud.google.com/jobs/%s', projectId);
}

function tableExists(tableId, projectId, datasetId) {
    // Get a list of all tables in the dataset.
    var tables = BigQuery.Tables.list(projectId, datasetId);
    var tableExists = false;
    // Iterate through each table and check for an id match.
    if (tables.tables != null) {
        for (var i = 0; i < tables.tables.length; i++) {
            var table = tables.tables[i];
            if (table.tableReference.tableId == tableId) {
                tableExists = true;
                break;
            }
        }
    }
    return tableExists;
}

function readData(my_spreadsheet_id, reqsheet, upload_time) {
    var ss = SpreadsheetApp.openById(my_spreadsheet_id)
    SpreadsheetApp.setActiveSpreadsheet(ss)
    // Select Specific Sheet from the Active Spreadsheet
    var sheet = ss.setActiveSheet(ss.getSheetByName(reqsheet));

    // Retrive Values
    var rows = sheet.getDataRange().getValues();
    var dataLength = rows.length

    // Remove the Header Row
    rows.shift();
    console.log(rows.length);

    if (dataLength == 1) {
        console.log("No data")
        return [];
    }
    else {
        console.log(sheet.getName());
        console.log(rows[0]);
        // Iterate throught the rows and modify the array
        for (var j = 0; j < rows.length; j++) {
            rows[j][0] = Utilities.formatDate(rows[j][0], Session.getTimeZone(), "yyyy-MM-dd")
            rows[j][1] = rows[j][7].replace(",", ";")
            rows[j][2] = (rows[j][9].length === 0) ? undefined : Utilities.formatDate(rows[j][9], Session.getTimeZone(), "yyyy-MM-dd")
            rows[j][3] = (rows[j][10].length === 0) ? undefined : Utilities.formatDate(rows[j][10], Session.getTimeZone(), "yyyy-MM-dd HH:mm:ss")
            rows[j].unshift(upload_time)
        }

        console.log("Size: " + rows.length + "," + rows[0].length);

        // create a CSV from the Array Objects to Upload to big query
        var Rows_CSV = rows.join('\n');
        return Rows_CSV;
    }
}

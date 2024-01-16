Office.onReady((info) => {
    // Check that we loaded into Excel
    if (info.host === Office.HostType.Excel) {
        document.getElementById("createTable").onclick = createTable;
        document.getElementById("getResults").onclick = getResults;
    }
});

// Display a status
/**
 * @param {unknown} message
 * @param {boolean} isError
 */
function showStatus(message, isError) {
    $('.status').empty();
    $('<div/>', {
        class: `status-card ms-depth-4 ${isError ? 'error-msg' : 'success-msg'}`
    }).append($('<p/>', {
        class: 'ms-fontSize-24 ms-fontWeight-bold',
        text: isError ? 'An error occurred' : 'Success'
    })).append($('<p/>', {
        class: 'ms-fontSize-16 ms-fontWeight-regular',
        text: message
    })).appendTo('.status');
}
async function createTable() {
    await Excel.run(async (context) => {
        try {
            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            console.log('Creating table');
        
            velvetTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
            velvetTable.name = "VelvetTable";
            

            velvetTable.getHeaderRowRange().values =
                [["ID", "Query", "Expected Summary", "Actual Summary"]];

            /* velvetTable.rows.add(null, [
                ["1", "`You are expert financial analyst. Be terse. Answer the question with minimal facts. What is Google's revenue for year ending 2022?`", "Revenue was $282.8 billion in 2022", ""],
                ["2", "You are expert financial analyst. Be terse. Answer the question with minimal facts. What is Google's revenue for Q1 2023 in billions?", "Revenue was $69.8 billion", ""],
                ["3", "You are expert financial analyst. Be terse. Answer the question with minimal facts. How much did Google invest in research and development (R&D) in 2022?", "Google's parent company Alphabet spent $39.5 billion on research and development (R&D) in 2022", ""],
                ["4", "You are expert financial analyst. Be terse. Answer the question with minimal facts. How much did Google invest in research and development (R&D) in 2022?", "Google's parent company Alphabet spent $39.5 billion on research and development (R&D) in 2022", ""]
            ]); */
            velvetTable.resize('A1:S100');
            // Learn more about the Excel number format syntax in this article:
            // https://support.microsoft.com/office/5026bbd6-04bc-48cd-bf33-80f18b4eae68
            //velvetTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];

            velvetTable.getRange().format.autofitColumns();
            velvetTable.getRange().format.autofitRows();

            await context.sync();
        } catch (error) {
            console.log('Error createTable: ' + error);
            showStatus(`Exception when creating sample data: ${JSON.stringify(error)}`, true);
        }
    });
    
}

function getColumn(table, columnName) {
    try {
        const column = table.columns.getItemOrNullObject(columnName);
        column.load();
        return column;
    } catch (error) {
        console.log('Error getColumn: ' + error);
        showStatus(`Exception when getting column: ${JSON.stringify(error)}`, true);
    }
}

async function getResults() {
    
     await Excel.run(async (context) => {
         try {
                const table = context.workbook.tables.getItem("VelvetTable");
                const queryColumn = getColumn(table, "Query");
                const idColumn = getColumn(table, "ID");
                const actualSummaryColumn = getColumn(table, "Actual Summary");
                table.rows.load('count');
                await context.sync();
                if (!queryColumn.isNullObject && !idColumn.isNullObject) {
                    let rowNum = 1;
                    let id = idColumn.values;
                    let query = queryColumn.values;
                    let returnedSummary = actualSummaryColumn.values;
                    //console.log('Number of rows in table:' + table.rows.count);
                    while (id[rowNum][0] !== undefined && id[rowNum][0] !== null && id[rowNum][0] !== "") {
                        console.log('ID:' + id[rowNum][0]);
                        console.log('Query: ' + query[rowNum][0]);

                        // add to function array
                         callVertexAISearch(rowNum, query[rowNum][0]).then(async function (result) {
                            
                             if (result.response.hasOwnProperty('summary')) {
                                 
                                 const cell = actualSummaryColumn.getRange().getCell(result.rowNum, 0);
                                 cell.values = [[result.response.summary.summaryText]];
                                 await context.sync();  
                                 
                             } else if (result.response.hasOwnProperty('error')) {
                                 throw Error(result.response.error.message);
                             } 
                        
                        }).catch(async function (error) {
                            console.log('Error callVertexAISearch: ' + error);
                            showStatus(`Exception when getting results: ${JSON.stringify(error.message)}`, true);   
                        });  
                        rowNum++;
                    }
                }
                table.getRange().format.autofitColumns();
                table.getRange().format.autofitRows();
                await context.sync();
                showStatus("Calling Vertex AI Search", false);
             
        } catch (error) {
            console.log('Error in getResults: ' + error);
            showStatus(`Exception when getting results: ${JSON.stringify(error)}`, true);
        }
    });
    
}

async function callVertexAISearch(rowNum, query) {

    try {
       
       // Get the value of the preamble field.
        const preamble = $('#preamble').val();
        const summaryResultCount = $('#summaryResultCount').val();
        const maxExtractiveAnswerCount = $('#maxExtractiveAnswerCount').val();
        const ignoreAdversarialQuery = $('#ignoreAdversarialQuery').val();
        const ignoreNonSummarySeekingQuery = $('#ignoreNonSummarySeekingQuery').val();
        const token = $('#access-token').val();
        const projectNumber = $('#project-number').val();
        const datastoreName = $('#datastore-name').val();
        

        var data = {
            query: query,
            page_size: 5,
            offset: 0,
            contentSearchSpec: {
                extractiveContentSpec: { maxExtractiveAnswerCount: `${maxExtractiveAnswerCount}` },
                summarySpec: {
                    summaryResultCount: `${summaryResultCount}`,
                    ignoreAdversarialQuery: `${ignoreAdversarialQuery}`,
                    ignoreNonSummarySeekingQuery: `${ignoreNonSummarySeekingQuery}`,
                    modelPromptSpec: {
                        preamble: `${preamble}`
                    }
                },
            }
        };
       
        const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${projectNumber}/locations/global/collections/default_collection/dataStores/${datastoreName}/servingConfigs/default_search:search`;
       
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(data),
        });
    
        const json = await response.json();
        return {rowNum: rowNum, response: json}; 
       
    } catch (error) {
        console.log('Error calling callVertexAISearch: ' + error);
        throw error;

    }
}

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
            currentWorksheet.load("name");
            await context.sync();
            const worksheetName = currentWorksheet.name;

            range = currentWorksheet.getRange('A1');
            range.values = [["Vertex AI Search Parameters"]];
            range.format.font.bold = true;
            range.format.fill.color = 'yellow';
            range.format.font.size = 16;

            configTable = currentWorksheet.tables.add("A2:B2", true /*hasHeaders*/);
            configTable.name = `${worksheetName}.ConfigTable`;

            configTable.getHeaderRowRange().values =
                [["Config", "Value"]];

            configTable.rows.add(null, [
                ["Vertex AI Search Project Number", "384473000457"],
                ["Vertex AI Search DataStore Name", "alphabet-pdfs_1695783402380"],
                ["Vertex AI Project ID", "argolis-arau"],
                ["Vertex AI Location", "us-central1"],
                ["maxExtractiveAnswerCount (1-5)", "1"],
                ["Preamble", "Put your preamble here"],
                ["summaryResultCount (1-5)", "1"],
                ["ignoreAdversarialQuery (True or False)", "True"],
                ["ignoreNonSummarySeekingQuery (True or False)", "True"]
            ]);

            currentWorksheet.getUsedRange().format.autofitColumns();
            currentWorksheet.getUsedRange().format.autofitRows();
            await context.sync();

            velvetTable = currentWorksheet.tables.add("A15:E15", true /*hasHeaders*/);
            velvetTable.name = `${worksheetName}.TestCasesTable`;

            velvetTable.getHeaderRowRange().values =
                [["ID", "Query", "Expected Summary", "Actual Summary","Summary Score"]];

            /*velvetTable.rows.add(null, [
               ["1", "`You are expert financial analyst. Be terse. Answer the question with minimal facts. What is Google's revenue for year ending 2022?`", "Revenue was $282.8 billion in 2022", ""],
               ["2", "You are expert financial analyst. Be terse. Answer the question with minimal facts. What is Google's revenue for Q1 2023 in billions?", "Revenue was $69.8 billion", ""],
               ["3", "You are expert financial analyst. Be terse. Answer the question with minimal facts. How much did Google invest in research and development (R&D) in 2022?", "Google's parent company Alphabet spent $39.5 billion on research and development (R&D) in 2022", ""],
               ["4", "You are expert financial analyst. Be terse. Answer the question with minimal facts. How much did Google invest in research and development (R&D) in 2022?", "Google's parent company Alphabet spent $39.5 billion on research and development (R&D) in 2022", ""]
           ]); */

            velvetTable.resize('A15:E150');
            currentWorksheet.getUsedRange().format.autofitColumns();
            currentWorksheet.getUsedRange().format.autofitRows();

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
            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            currentWorksheet.load("name");
            await context.sync();
            const worksheetName = currentWorksheet.name;

            const testCasesTable = currentWorksheet.tables.getItem(`${worksheetName}.TestCasesTable`);
            const queryColumn = getColumn(testCasesTable, "Query");
            const idColumn = getColumn(testCasesTable, "ID");
            const actualSummaryColumn = getColumn(testCasesTable, "Actual Summary");
            const expectedSummaryColumn = getColumn(testCasesTable, "Expected Summary");
            const summaryScoreColumn = getColumn(testCasesTable, "Summary Score");

            testCasesTable.rows.load('count');
            await context.sync();

            const configTable = currentWorksheet.tables.getItem(`${worksheetName}.ConfigTable`);
            const configColumn = getColumn(configTable, "Config");
            const valueColumn = getColumn(configTable, "Value");
            valueColumn.load();
            await context.sync();
            const config = {
                vertexAISearchProjectNumber: valueColumn.values[1][0],
                vertexAISearchDataStoreName: valueColumn.values[2][0],
                vertexAIProjectID: valueColumn.values[3][0],
                vertexAILocation: valueColumn.values[4][0],
                maxExtractiveAnswerCount: valueColumn.values[5][0],
                preamble: valueColumn.values[6][0],
                summaryResultCount: valueColumn.values[7][0],
                ignoreAdversarialQuery: valueColumn.values[8][0],
                ignoreNonSummarySeekingQuery: valueColumn.values[9][0],
            };

            if (!queryColumn.isNullObject && !idColumn.isNullObject) {
                let rowNum = 1;
                let id = idColumn.values;
                let query = queryColumn.values;
                let expectedSummary = expectedSummaryColumn.values;
                

                //console.log('Number of rows in table:' + table.rows.count);
                while (rowNum <= testCasesTable.rows.count && id[rowNum][0] !== null && id[rowNum][0] !== "") {
                    //console.log('ID:' + id[rowNum][0]);
                    console.log('Query: ' + query[rowNum][0]);

                    // add to function array
                    callVertexAISearch(rowNum, query[rowNum][0], config).then(async function (result) {

                        if (result.response.hasOwnProperty('summary')) {

                            const cell = actualSummaryColumn.getRange().getCell(result.rowNum, 0);
                            cell.values = [[result.response.summary.summaryText]];
                            await context.sync();

                            const score = await calculateSimilarity(result.response.summary.summaryText, expectedSummary[result.rowNum][0], config);
                            const score_cell = summaryScoreColumn.getRange().getCell(result.rowNum, 0);
                            score_cell.values = [[score]];
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


            currentWorksheet.getUsedRange().format.autofitColumns();
            currentWorksheet.getUsedRange().format.autofitRows();
            await context.sync();
            showStatus("Calling Vertex AI Search", false);

        } catch (error) {
            console.log('Error in getResults: ' + error);
            showStatus(`Exception when getting results: ${JSON.stringify(error)}`, true);
        }
    });

}

async function callVertexAISearch(rowNum, query, config) {

    try {

        const token = $('#access-token').val();
        // Get the value of the preamble field.
        /* const preamble = $('#preamble').val();
        const summaryResultCount = $('#summaryResultCount').val();
        const maxExtractiveAnswerCount = $('#maxExtractiveAnswerCount').val();
        const ignoreAdversarialQuery = $('#ignoreAdversarialQuery').val();
        const ignoreNonSummarySeekingQuery = $('#ignoreNonSummarySeekingQuery').val();
        const projectNumber = $('#project-number').val();
        const datastoreName = $('#datastore-name').val(); */

        const preamble = config.preamble;
        const summaryResultCount = config.summaryResultCount;
        const maxExtractiveAnswerCount = config.maxExtractiveAnswerCount; 
        const ignoreAdversarialQuery = config.ignoreAdversarialQuery;
        const ignoreNonSummarySeekingQuery = config.ignoreNonSummarySeekingQuery;
        const projectNumber = config.vertexAISearchProjectNumber;
        const datastoreName = config.vertexAISearchDataStoreName;


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

        if (!response.ok) {
            throw new Error(`callVertexAISearch: Request failed with status ${response.status}`);
        }
        const json = await response.json();
        return { rowNum: rowNum, response: json };

    } catch (error) {
        console.error('Error calling callVertexAISearch: ', error);
        throw error;

    }
}

async function calculateSimilarity(sentence1, sentence2, config) {
    try {

        const projectId = config.projectId;
        const location = config.location;
        const token = $('#access-token').val();

        // Get embedding for each sentence
        const embedding1 = await getEmbedding(sentence1, token, projectId, location);
        const embedding2 = await getEmbedding(sentence2, token, projectId, location);

        // Calculate cosine similarity (assuming embeddings are NumPy arrays)
        let dotProduct = 0;
        let magnitude1 = 0;
        let magnitude2 = 0;

        for (let i = 0; i < embedding1.length; i++) {
            dotProduct += embedding1[i] * embedding2[i];
            magnitude1 += embedding1[i] * embedding1[i];
            magnitude2 += embedding2[i] * embedding2[i];
        }

        magnitude1 = Math.sqrt(magnitude1);
        magnitude2 = Math.sqrt(magnitude2);

        const cosineSimilarity = dotProduct / (magnitude1 * magnitude2);

        return cosineSimilarity;

    } catch (error) {
        console.error("calculateSimilarity: Error calculating similarity:", error);
        // Handle errors appropriately, e.g., return error message to client
        throw error;
    }
}
async function getEmbedding(text, token, projectId, location) {

    const endpointURL = `https://${location}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${location}/publishers/google/models/textembedding-gecko:predict`

    const instance = {
        instances: [{ // Assuming your model expects a list of instances
            task_type: `SEMANTIC_SIMILARITY`,
            content: text // For text input, use "content" (adjust for images etc.)
        }]
    };

    try {
        const response = await fetch(endpointURL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json; charset=utf-8',
                'Authorization': `Bearer ${token}` // Replace with your authentication
            },
            body: JSON.stringify(instance),
        });

        if (!response.ok) {
            throw new Error(`getEmbedding: Request failed with status ${response.status}`);
        }

        const responseData = await response.json();
        const embedding = responseData.predictions[0].embeddings.values; // Extract embedding (adjust structure if needed)

        return embedding;

    } catch (error) {
        console.error('Error getting embedding:', error);
        throw error;
    }
}

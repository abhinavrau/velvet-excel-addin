

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

                            const score = await calculateSimilarityUsingVertexAI(result.response.summary.summaryText, expectedSummary[result.rowNum][0], config);
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

async function calculateSimilarityUsingVertexAI(sentence1, sentence2, config) {

    try {

        const token = $('#access-token').val();
        const projectId = config.vertexAIProjectID;
        const location = config.vertexAILocation;
        
        var prompt = "You will get two answers to a question, you should determine if they are semantically similar or not. " +
            "examples - answer_1: I was created by X. answer_2: X created me. output: same" +
            "answer_1:There are 52 days in a year. answer_2: A year is fairly long. output: different"

        var full_prompt = `${prompt} answer_1: ${sentence1} answer_2: ${sentence2} output:`;

        // "$prompt". Now answer answer_1:
        var data = {
            instances: [
                { prompt: `${full_prompt}` }
            ],
            parameters: {
                temperature: 0.2,
                maxOutputTokens: 256,
                topK: 40,
                topP: 0.95,
                logprobs: 2
            }
        }

        const url = `https://${location}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${location}/publishers/google/models/text-bison:predict`;

        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(data),
        });

        if (!response.ok) {
            throw new Error(`calculateSimilarityUsingVertexAI: Request failed with status ${response.status}`);
        }
        const json = await response.json();
        return json.predictions[0].content;

    } catch (error) {
        console.error('Error calling calculateSimilarityUsingVertexAI: ', error);
        throw error;

    }
}

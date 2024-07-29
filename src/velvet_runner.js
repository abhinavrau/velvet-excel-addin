
import { calculateSimilarityUsingVertexAI, callVertexAISearch } from "./vertex_ai.js";

import { showStatus } from "./ui.js";


class NotAuthenticatedError extends Error {
    constructor(message = 'User is not authenticated') {
        super(message);
        this.name = 'NotAuthenticatedError';
        this.statusCode = 401; // Optional: HTTP status code for API errors
    }
}

class QuotaError extends Error {
    constructor(message = 'Quota Exceeded') {
        super(message);
        this.name = 'QuotaError';
        this.statusCode = 429; // Optional: HTTP status code for API errors
    }
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

export async function getConfig() {
    var config;
    await Excel.run(async (context) => {

        try {
            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            currentWorksheet.load("name");
            await context.sync();
            const worksheetName = currentWorksheet.name;
            const configTable = currentWorksheet.tables.getItem(`${worksheetName}.ConfigTable`);
            const valueColumn = getColumn(configTable, "Value");
            await context.sync();

            config = {
                vertexAISearchProjectNumber: valueColumn.values[1][0],
                vertexAISearchDataStoreName: valueColumn.values[2][0],
                vertexAIProjectID: valueColumn.values[3][0],
                vertexAILocation: valueColumn.values[4][0],
                extractiveContentSpec: {
                    maxExtractiveAnswerCount: valueColumn.values[5][0] === 0 ? null : valueColumn.values[5][0],
                    maxExtractiveSegmentCount: valueColumn.values[6][0] === 0 ? null : valueColumn.values[6][0],
                },
                maxSnippetCount: valueColumn.values[7][0] === 0 ? null : valueColumn.values[7][0],
                preamble: valueColumn.values[8][0],
                model: valueColumn.values[9][0],
                summaryResultCount: valueColumn.values[10][0],
                useSemanticChunks: valueColumn.values[11][0],
                ignoreAdversarialQuery: valueColumn.values[12][0],
                ignoreNonSummarySeekingQuery: valueColumn.values[13][0],
                summaryMatchingAdditionalPrompt: valueColumn.values[14][0],
                batchSize: valueColumn.values[15][0],
                timeBetweenCallsInSec: valueColumn.values[16][0],
                accessToken: $('#access-token').val(),
            };

        } catch (error) {
            console.error(`Caught Exception in createConfig: ${error} `);
            showStatus(`Caught Exception in createConfig: ${error}`, true);
            throw error;
        }

    });
    return config;
}
export async function executeTests() {

    const config = await getConfig();
    await runTests(config);
}




export async function runTests(config) {

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
            const summaryScoreColumn = getColumn(testCasesTable, "Summary Match");

            const link_1_Column = getColumn(testCasesTable, "Actual Link 1");
            const link_2_Column = getColumn(testCasesTable, "Actual Link 2");
            const link_3_Column = getColumn(testCasesTable, "Actual Link 3");
            const expected_link_1_Column = getColumn(testCasesTable, "Expected Link 1");
            const expected_link_2_Column = getColumn(testCasesTable, "Expected Link 2");
            const expected_link_3_Column = getColumn(testCasesTable, "Expected Link 3");
            const link_p0Column = getColumn(testCasesTable, "First Link Match");
            const link_top2Column = getColumn(testCasesTable, "Link in Top 2");

            testCasesTable.rows.load('count');
            await context.sync();

            if (config.accessToken === "") {
                showStatus(`Error: executeTests: Access token is empty`, true);
                throw error("Access token is empty");
            }


            // Validate config
            const isValid = (config.extractiveContentSpec.maxExtractiveAnswerCount !== null)
                ^ (config.extractiveContentSpec.maxExtractiveSegmentCount !== null)
                ^ (config.maxSnippetCount !== null);

            if (!isValid) {
                // None, multiple, or all variables are non-null    
                showStatus(`Error: executeTests: Only one of the maxExtractiveAnswerCount, maxExtractiveSegmentCount, or maxSnippetCount should be set to a non-zero value`, true);
                return;
            }

            if (queryColumn.isNullObject || idColumn.isNullObject) {
                showStatus(`Error: executeTests: No Query or ID column found in Test Cases Table. Make sure there is an ID and Query column in the Test Cases Table.`, true);
                return;
            }


            let rowNum = 1;
            let id = idColumn.values;
            let query = queryColumn.values;
            let expectedSummary = expectedSummaryColumn.values;
            let expectedLink1 = expected_link_1_Column.values;
            let expectedLink2 = expected_link_2_Column.values;

            let numfails = 0;
            let errorMessages = [];

            // map of promises
            const promiseMap = new Map();
            let stopProcessing = false;
            //console.log('Number of rows in table:' + table.rows.count);
            // Loop through the test cases table ans run the tests
            while (rowNum <= testCasesTable.rows.count && id[rowNum][0] !== null && id[rowNum][0] !== "") {
                console.log('ID:' + id[rowNum][0]);
                console.log('Query: ' + query[rowNum][0]);

                // Batch the calls to Vertex AI since there are throuput checks in place.\
                if (rowNum % config.batchSize === 0) {
                    // delay calls with apropriate time
                    await new Promise(r => setTimeout(r, config.timeBetweenCallsInSec * 1000));
                }

                // Call Vertex AI Search and add the promise to promiseMap
                promiseMap.set(id[rowNum][0], callVertexAISearch(rowNum, query[rowNum][0], config)
                    .then(result => {
                        let response = result.output;
                        let testCaseNum = result.testCaseNum;

                        console.log(`result.output: ${JSON.stringify(result.output)} `);
                        // Check the summary first
                        if (response.hasOwnProperty('summary')) {
                            console.log("Got Summary");
                            processSummary(testCaseNum, response, actualSummaryColumn, expectedSummary, config, summaryScoreColumn, context);
                        }
                        // Check the documents references
                        if (response.hasOwnProperty('results')) {
                            console.log("Got links");
                            checkDocumentLinks(testCaseNum, response, link_1_Column, link_2_Column, link_3_Column, link_p0Column, link_top2Column, expectedLink1, expectedLink2, context);
                        }
                        // check for error json property
                        if (response.hasOwnProperty('error') || result.status_code !== 200) {
                            console.error(`executeTests: VAI returned error for row: ${testCaseNum} errorcode: ${result.status_code} error: ${JSON.stringify(response)}`);
                            errorMessages += `VAI returned error for row:: row: ${testCaseNum} error: ${JSON.stringify(response)}`;
                            if (result.status_code === 401) {
                                throw new NotAuthenticatedError();
                            }
                            if (result.status_code === 429) {
                                throw new QuotaError();
                            }

                        }
                        showStatus(`Processed ${rowNum} test cases.  numFails:${numfails} \n\n ${errorMessages}`, numfails > 0);
                    })
                    .catch(error => {
                        numfails++;
                        if (error instanceof NotAuthenticatedError) {
                            stopProcessing = true;
                            showStatus(`User not Authenticated`, true);
                        }
                        if (error instanceof QuotaError) {
                            stopProcessing = true;
                            showStatus(`API Quota Exceeded`, true);
                        }
                        else {
                            // ouput stacktrace for error 
                            console.error(`executeTests: Error in row: ${testCaseNum} numFails:${numfails} error: ${error} with stack: ${error.stack}`);
                            errorMessages += `executeTests: Error for row: ${testCaseNum}  error: ${error}`;
                        }

                    }));

                if (stopProcessing) {
                    break;
                }
                rowNum++;
            } // end while

            // wait for all the calls to finish
            await Promise.allSettled(promiseMap.values());

            currentWorksheet.getUsedRange().format.autofitColumns();
            currentWorksheet.getUsedRange().format.autofitRows();
            await context.sync();


        } catch (error) {
            console.error(`Caught Exception in executeTests: ${error} `);
            showStatus(`Caught Exception in executeTests: ${JSON.stringify(error)}`, true);
            throw error;
        }

    });

}

function checkDocumentLinks(rowNum, result, link_1_Column, link_2_Column, link_3_Column, link_p0Column, link_top2Column, expectedLink1, expectedLink2, context) {
    var p0_result = null;
    var p2_result = null;
    // Check for document info and linksin the metadata if it exists
    if (result.results[0].document.hasOwnProperty('structData')) {
        link_1_Column.getRange().getCell(rowNum, 0).values = [[result.results[0].document.structData.sharepoint_ref]];
        p0_result = result.results[0].document.structData.title;
    } else if (result.results[0].document.hasOwnProperty('derivedStructData')) {
        link_1_Column.getRange().getCell(rowNum, 0).values = [[result.results[0].document.derivedStructData.link]];
        p0_result = result.results[0].document.derivedStructData.link;
    }
    if (result.results[1].document.hasOwnProperty('structData')) {
        link_2_Column.getRange().getCell(rowNum, 0).values = [[result.results[1].document.structData.sharepoint_ref]];
    } else if (result.results[1].document.hasOwnProperty('derivedStructData')) {
        link_2_Column.getRange().getCell(rowNum, 0).values = [[result.results[1].document.derivedStructData.link]];
        p2_result = result.results[1].document.derivedStructData.link;
    }
    if (result.results[2].document.hasOwnProperty('structData')) {
        link_3_Column.getRange().getCell(rowNum, 0).values = [[result.results[2].document.structData.sharepoint_ref]];
    } else if (result.results[2].document.hasOwnProperty('derivedStructData')) {
        link_3_Column.getRange().getCell(rowNum, 0).values = [[result.results[2].document.derivedStructData.link]];
    }

    // clear the formatting in the cells 
    const link_p0_cell = link_p0Column.getRange().getCell(rowNum, 0);
    link_p0_cell.clear(Excel.ClearApplyTo.formats);
    const link1_cell = link_1_Column.getRange().getCell(rowNum, 0);
    link1_cell.clear(Excel.ClearApplyTo.formats);
    const top2_cell = link_top2Column.getRange().getCell(rowNum, 0);
    top2_cell.clear(Excel.ClearApplyTo.formats);

    // match first link with expected link
    if (p0_result !== null && p0_result === expectedLink1[rowNum][0]) {
        link_p0_cell.values = [["TRUE"]];
    } else {
        link_p0_cell.values = [["FALSE"]];
        link_p0_cell.format.fill.color = '#FFCCCB';
        link1_cell.format.fill.color = '#FFCCCB';

    }

    // match if the top 2 links returned are in the top 2 expected links
    if (p2_result !== null && p2_result === expectedLink2[rowNum][0]
        || p2_result === expectedLink1[rowNum][0]
        || p0_result === expectedLink1[rowNum][0]
        || p0_result === expectedLink2[rowNum][0]) {
        top2_cell.values = [["TRUE"]];
    } else {
        top2_cell.values = [["FALSE"]];
        top2_cell.format.fill.color = '#FFCCCB';
    }
    context.sync();
}

function processSummary(rowNum, result, actualSummaryColumn, expectedSummary, config, summaryScoreColumn, context) {
    console.log('Summary: ' + result.summary.summaryText);
    const cell = actualSummaryColumn.getRange().getCell(rowNum, 0);
    cell.clear(Excel.ClearApplyTo.formats);
    cell.values = [[result.summary.summaryText]];

    // match summaries only if they are not null or empty
    if (expectedSummary[rowNum][0] !== null && expectedSummary[rowNum][0] !== "") {

        const response = calculateSimilarityUsingVertexAI(rowNum, result.summary.summaryText, expectedSummary[rowNum][0], config);
        const score = response.output;

        const score_cell = summaryScoreColumn.getRange().getCell(rowNum, 0);
        score_cell.clear(Excel.ClearApplyTo.formats);
        //console.log('result.rowNum ' + result.rowNum + ' score: ' + score);
        if (score.trim() === 'same') {
            score_cell.values = [["TRUE"]];

        } else {
            score_cell.values = [["FALSE"]];
            score_cell.format.fill.color = '#FFCCCB';
            const actualSummarycell = actualSummaryColumn.getRange().getCell(rowNum, 0);
            actualSummarycell.format.fill.color = '#FFCCCB';

        }
    }
    context.sync();
}


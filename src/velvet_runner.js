
import { NotAuthenticatedError, QuotaError } from "./common.js";
import { calculateSimilarityUsingVertexAI, callVertexAISearch } from "./vertex_ai.js";

import { appendError, appendLog, showStatus } from "./ui.js";



function getColumn(table, columnName) {
    try {
        const column = table.columns.getItemOrNullObject(columnName);
        column.load();
        return column;
    } catch (error) {
        appendError('Error getColumn:',error);
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
            return null;
        }

    });
    return config;
}

var stopProcessing;

export async function executeTests(config) {
    
    if (config == null) {
        return;
    }
   
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

            if (config.accessToken === null || config.accessToken === "") {
                showStatus(`Access token is empty`, true);
                appendError(`Error in runTests: Access token is empty`, null);
                return;
            }


            // Validate config
            const isValid = (config.extractiveContentSpec.maxExtractiveAnswerCount !== null)
                ^ (config.extractiveContentSpec.maxExtractiveSegmentCount !== null)
                ^ (config.maxSnippetCount !== null);

            if (!isValid) {
                // None, multiple, or all variables are non-null    
                showStatus(`Error in runTests: Only one of the maxExtractiveAnswerCount, maxExtractiveSegmentCount, or maxSnippetCount should be set to a non-zero value`, true);
                return;
            }

            if (queryColumn.isNullObject || idColumn.isNullObject) {
                showStatus(`Error in runTests: No Query or ID column found in Test Cases Table. Make sure there is an ID and Query column in the Test Cases Table.`, true);
                return;
            }


            let processedCount = 1;
            let id = idColumn.values;
            let query = queryColumn.values;
            let expectedSummary = expectedSummaryColumn.values;
            let expectedLink1 = expected_link_1_Column.values;
            let expectedLink2 = expected_link_2_Column.values;

            let numfails = 0;
            const countRows = testCasesTable.rows.count;

            // map of promises
            const promiseMap = new Map();
           
            stopProcessing = false;
            // Loop through the test cases table ans run the tests
            while (processedCount <= countRows && id[processedCount][0] !== null && id[processedCount][0] !== "") {
                

                // Batch the calls to Vertex AI since there are throuput checks in place.\
                if (processedCount % config.batchSize === 0) {
                    // delay calls with apropriate time
                    await new Promise(r => setTimeout(r, config.timeBetweenCallsInSec * 1000));
                }
                // Stop processing if there errors
                if (stopProcessing) {
                    appendLog("Stopping execution.", null);
                    break;
                }
                appendLog(`testCaseID: ${id[processedCount][0]} Start Processing.`);
                showStatus(`Processing testCaseID: ${id[processedCount][0]}`, false);
                // Call Vertex AI Search asynchronously and add the promise to promiseMap
                promiseMap.set(processedCount, callVertexAISearch(processedCount, query[processedCount][0], config)
                    .then(async result => {
                        let response = result.output;
                        let testCaseRowNum = result.testCaseRowNum;
                        
                        // Check the summary first
                        if (response.hasOwnProperty('summary')) {
                            await processSummary(testCaseRowNum, response, actualSummaryColumn, expectedSummary, config, summaryScoreColumn, context);
                            appendLog(`testCaseID: ${testCaseRowNum} Processed Summary.`);
                        }
                        // Check the documents references
                        if (response.hasOwnProperty('results')) {
                            checkDocumentLinks(testCaseRowNum, response, link_1_Column, link_2_Column, link_3_Column, link_p0Column, link_top2Column, expectedLink1, expectedLink2, context);
                            appendLog(`testCaseID: ${testCaseRowNum} Processed Doc Links.`);
                        }
                       
                    })
                    .catch(error => {
                        var errorMessage = "";
                        if (error instanceof NotAuthenticatedError) {
                            errorMessage = `User not Authenticated. Re-authenticate and try again.`;
                            stopProcessing = true;
                        }
                        else if (error instanceof QuotaError) {
                            errorMessage = `API Quota Exceeded for testCaseID: ${testCaseNum}. Reduce test cases or increase timeouts.`;
                            stopProcessing = true;
                        }
                        else {
                            errorMessage = `Error for testCaseID: ${testCaseNum}  error: ${error} with stack: ${error.stack}`;
                            
                        }
                        numfails++;
                        appendError(errorMessage, error);
                        
                        
                    }));
                
                processedCount++;
            } // end while

            // wait for all the calls to finish
            await Promise.allSettled(promiseMap.values());
            var stoppedReason = "";
            if (numfails > 0) {
                stoppedReason = `Failed: ${numfails}. See logs for details.`;
            }
            if (processedCount <= countRows && ( id[processedCount][0] === null || id[processedCount][0] === "")) {
                stoppedReason += ` Empty ID encountered after ${processedCount-1} test cases.`;
            }
            var summary = `Finished! Successful: ${(processedCount - numfails) - 1}. ${stoppedReason}`;
            appendLog(summary);

            showStatus(summary, numfails > 0);
            
            // autofit the content
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

export async function stopTests() { 
    stopProcessing = true; // Set the stop signal flag
    appendLog("Cancel Tests Clicked. Stopping test execution...");
}


function checkDocumentLinks(rowNum, result, link_1_Column, link_2_Column, link_3_Column, link_p0Column, link_top2Column, expectedLink1, expectedLink2, context) {
    var p0_result = null;
    var p2_result = null;
    const link_1_cell = link_1_Column.getRange().getCell(rowNum, 0);
    const link_2_cell = link_2_Column.getRange().getCell(rowNum, 0);
    const link_3_cell = link_3_Column.getRange().getCell(rowNum, 0);

    // Check for document info and linksin the metadata if it exists
    if (result.results[0].document.hasOwnProperty('structData')) {
        link_1_cell.values = [[result.results[0].document.structData.sharepoint_ref]];
        p0_result = result.results[0].document.structData.title;
    } else if (result.results[0].document.hasOwnProperty('derivedStructData')) {
        link_1_cell.values = [[result.results[0].document.derivedStructData.link]];
        p0_result = result.results[0].document.derivedStructData.link;
    }
    if (result.results[1].document.hasOwnProperty('structData')) {
        link_2_cell.values = [[result.results[1].document.structData.sharepoint_ref]];
    } else if (result.results[1].document.hasOwnProperty('derivedStructData')) {
        link_2_cell.values = [[result.results[1].document.derivedStructData.link]];
        p2_result = result.results[1].document.derivedStructData.link;
    }
    if (result.results[2].document.hasOwnProperty('structData')) {
        link_3_cell.values = [[result.results[2].document.structData.sharepoint_ref]];
    } else if (result.results[2].document.hasOwnProperty('derivedStructData')) {
        link_3_cell.values = [[result.results[2].document.derivedStructData.link]];
    }

    // clear the formatting in the cells 
    const link_p0_cell = link_p0Column.getRange().getCell(rowNum, 0);
    link_p0_cell.clear(Excel.ClearApplyTo.formats);
    link_1_cell.clear(Excel.ClearApplyTo.formats);
    const top2_cell = link_top2Column.getRange().getCell(rowNum, 0);
    top2_cell.clear(Excel.ClearApplyTo.formats);

    // match first link with expected link
    if (p0_result !== null && p0_result === expectedLink1[rowNum][0]) {
        link_p0_cell.values = [["TRUE"]];
    } else {
        link_p0_cell.values = [["FALSE"]];
        link_p0_cell.format.fill.color = '#FFCCCB';
        link_1_cell.format.fill.color = '#FFCCCB';

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
 async function processSummary(rowNum, result, actualSummaryColumn, expectedSummary, config, summaryScoreColumn, context) {
    
     // Set the actual summary
    const cell = actualSummaryColumn.getRange().getCell(rowNum, 0);
    cell.clear(Excel.ClearApplyTo.formats);
    cell.values = [[result.summary.summaryText]];
    context.sync();
    // match summaries only if they are not null or empty
    if (expectedSummary[rowNum][0] !== null && expectedSummary[rowNum][0] !== "") {
        const score_cell = summaryScoreColumn.getRange().getCell(rowNum, 0);
        score_cell.clear(Excel.ClearApplyTo.formats);

        // Catch any errors here and report it in the cell. We don't want failures here to stop processing.
        try {
           
            const response = await calculateSimilarityUsingVertexAI(rowNum, result.summary.summaryText, expectedSummary[rowNum][0], config);
            const score = response.output;

            if (score.trim() === 'same') {
                score_cell.values = [["TRUE"]];

            } else  {
                score_cell.values = [["FALSE"]];
                score_cell.format.fill.color = '#FFCCCB';
                const actualSummarycell = actualSummaryColumn.getRange().getCell(rowNum, 0);
                actualSummarycell.format.fill.color = '#FFCCCB';

            } 
        } catch(err) {
            // put the error in the cell.
            score_cell.values = [[err.message]];
            score_cell.format.fill.color = '#FFCCCB';
            const actualSummarycell = actualSummaryColumn.getRange().getCell(rowNum, 0);
            actualSummarycell.format.fill.color = '#FFCCCB';
        }
        finally {
            context.sync();
        }
       
    }
   
}


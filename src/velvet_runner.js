
import { calculateSimilarityUsingVertexAI, callVertexAISearch } from "./vertex_ai.js";

import { showStatus } from "./ui.js";


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

export async function executeTests() {

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
                accessToken: $('#access-token').val()
            };

            if (!queryColumn.isNullObject && !idColumn.isNullObject) {
                let rowNum = 1;
                let id = idColumn.values;
                let query = queryColumn.values;
                let expectedSummary = expectedSummaryColumn.values;
                let expectedLink1 = expected_link_1_Column.values;
                let expectedLink2 = expected_link_2_Column.values;

                let numfails = 0;
                //console.log('Number of rows in table:' + table.rows.count);
                while (rowNum <= testCasesTable.rows.count && id[rowNum][0] !== null && id[rowNum][0] !== "") {
                    console.log('ID:' + id[rowNum][0]);
                    console.log('Query: ' + query[rowNum][0]);


                    // check the modulus of 10 for the rownum so be batch 10 calls to vertex AI
                    if (rowNum % 10 === 0) {
                        showStatus(`Processed ${rowNum} test cases`, false);
                        console.log('Sleeping: ' + rowNum);
                        // sleep for 2 seconds
                        await new Promise(r => setTimeout(r, 3000));
                    }
                  

                    // add to function array
                    callVertexAISearch(rowNum, query[rowNum][0], config).then(async function (result) {

                        if (result.response.hasOwnProperty('summary')) {


                            console.log('Summary: ' + result.response.summary.summaryText);
                            const cell = actualSummaryColumn.getRange().getCell(result.rowNum, 0);
                            cell.clear(Excel.ClearApplyTo.formats);
                            cell.values = [[result.response.summary.summaryText]];

                            // match summaries only if they are not null or empty
                            if (expectedSummary[result.rowNum][0] !== null && expectedSummary[result.rowNum][0] !== "") {

                                const score = await calculateSimilarityUsingVertexAI(result.response.summary.summaryText, expectedSummary[result.rowNum][0], config);
                                const score_cell = summaryScoreColumn.getRange().getCell(result.rowNum, 0);
                                score_cell.clear(Excel.ClearApplyTo.formats);
                                //console.log('result.rowNum ' + result.rowNum + ' score: ' + score);

                                if (score.trim() === 'same') {
                                    score_cell.values = [["TRUE"]];

                                } else {
                                    score_cell.values = [["FALSE"]];
                                    score_cell.format.fill.color = '#FFCCCB';
                                    const actualSummarycell = actualSummaryColumn.getRange().getCell(result.rowNum, 0);
                                    actualSummarycell.format.fill.color = '#FFCCCB';

                                }
                            }
                            await context.sync();


                        }
                        if (result.response.hasOwnProperty('results')) {

                            var p0_result;
                            var p2_result;
                            if (result.response.results[0].document.hasOwnProperty('structData')) {
                                link_1_Column.getRange().getCell(result.rowNum, 0).values = [[result.response.results[0].document.structData.title]];
                                link_2_Column.getRange().getCell(result.rowNum, 0).values = [[result.response.results[1].document.structData.title]];
                                link_3_Column.getRange().getCell(result.rowNum, 0).values = [[result.response.results[2].document.structData.title]];
                                p0_result = result.response.results[0].document.structData.title;
                                p2_result = result.response.results[1].document.structData.title;
                            } else {
                                link_1_Column.getRange().getCell(result.rowNum, 0).values = [[result.response.results[0].document.derivedStructData.link]];
                                link_2_Column.getRange().getCell(result.rowNum, 0).values = [[result.response.results[1].document.derivedStructData.link]];
                                link_3_Column.getRange().getCell(result.rowNum, 0).values = [[result.response.results[2].document.derivedStructData.link]];
                                p0_result = result.response.results[0].document.derivedStructData.link;
                                p2_result = result.response.results[1].document.derivedStructData.link;
                            }

                            const link_p0_cell = link_p0Column.getRange().getCell(result.rowNum, 0);
                            link_p0_cell.clear(Excel.ClearApplyTo.formats);


                            const link1_cell = link_1_Column.getRange().getCell(result.rowNum, 0);
                            link1_cell.clear(Excel.ClearApplyTo.formats);
                            const top2_cell = link_top2Column.getRange().getCell(result.rowNum, 0);
                            top2_cell.clear(Excel.ClearApplyTo.formats);
                            // match first link with expected link
                            if (p0_result === expectedLink1[result.rowNum][0]) {
                                link_p0_cell.values = [["TRUE"]];
                            } else {
                                link_p0_cell.values = [["FALSE"]];
                                link_p0_cell.format.fill.color = '#FFCCCB';
                                link1_cell.format.fill.color = '#FFCCCB';

                            }

                            // match top 2 
                            if (p2_result === expectedLink2[result.rowNum][0]
                                || p2_result === expectedLink1[result.rowNum][0]
                                || p0_result === expectedLink1[result.rowNum][0]
                                || p0_result === expectedLink2[result.rowNum][0]) {
                                top2_cell.values = [["TRUE"]];
                            } else {
                                top2_cell.values = [["FALSE"]];
                                top2_cell.format.fill.color = '#FFCCCB';
                            }
                            await context.sync();
                        }
                        // check for error json property
                        if (result.response.hasOwnProperty('error')) {
                            throw Error(result.response.error.message);
                        }

                    }).catch(async function (error) {
                        
                        numfails++;
                        console.log(`Error callVertexAISearch: row: ${result.rowNum} numFails:${numfails} error: ${JSON.stringify(error.message)}`);
                        showStatus(`Error callVertexAISearch: row: ${result.rowNum} numFails:${numfails} error: ${JSON.stringify(error.message)}`, true);
                    });
                    rowNum++;
                }
            }


            currentWorksheet.getUsedRange().format.autofitColumns();
            //currentWorksheet.getUsedRange().format.autofitRows();
            await context.sync();
            showStatus("Calling Vertex AI Search", false);

        } catch (error) {
            console.log('Error in getResults: ' + error);
            throw error;
            //showStatus(`Exception when getting results: ${JSON.stringify(error)}`, true);
        }
    });

}





import { appendError, appendLog, showStatus } from "./ui.js";

import { callGeminiMultitModal, callVertexAI } from "./vertex_ai.js";

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

export async function getSummarizationConfig() {
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
                vertexAIProjectID: valueColumn.values[1][0],
                vertexAILocation: valueColumn.values[2][0],
                model: valueColumn.values[3][0],
                prompt: valueColumn.values[4][0],
                batchSize: valueColumn.values[5][0],
                timeBetweenCallsInSec: valueColumn.values[6][0],
                accessToken: $('#access-token').val(),
                systemInstruction: "",
                responseMimeType: "text/plain",
            };

        } catch (error) {
            appendError(`Caught Exception in Summarization createConfig: ${error} `, error);
            showStatus(`Caught Exception in Summarization createConfig: ${error}`, true);
            return null;
        }

    });
    return config;
}

var stopProcessing;

export async function createSummarizationData(config) {
    
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
            const idColumn = getColumn(testCasesTable, "ID");
            const toSummarizeColumn = getColumn(testCasesTable, "Context");
            const summaryColumn = getColumn(testCasesTable, "Summary");
            const summarization_qualityColumn = getColumn(testCasesTable, "summarization_quality");
            const groundednessColumn = getColumn(testCasesTable, "groundedness");
            const fulfillmentColumn = getColumn(testCasesTable, "fulfillment");
            const summarization_helpfulnesColumn = getColumn(testCasesTable, "summarization_helpfulness");
            const summarization_verbosityTimeColumn = getColumn(testCasesTable, "summarization_verbosity");

            testCasesTable.rows.load('count');
            await context.sync();

            if (config.accessToken === null || config.accessToken === "") {
                showStatus(`Access token is empty`, true);
                appendError(`Error in createSummarizationData: Access token is empty`, null);
                return;
            }


            if (toSummarizeColumn.isNullObject || idColumn.isNullObject) {
                showStatus(`Error in createSummarizationData: No fileUriColumn or ID column found in Test Cases Table. Make sure there is an ID and Query column in the Test Cases Table.`, true);
                return;
            }


            let processedCount = 1;
            let id = idColumn.values;
            let toSummarize = toSummarizeColumn.values;
           
                
            let numfails = 0;
            const countRows = testCasesTable.rows.count;

            // map of promises
            const promiseMap = new Map();
           
            stopProcessing = false;
            // Iterate rows on the  table. Stop when end of table or ID column is empty
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
             
               
                let full_prompt = config.prompt + " Text to summarize: " + toSummarize[processedCount][0];
                appendLog(`sumCaseID:: ${id[processedCount][0]} Start Processing. with prompt ${full_prompt}`);
                showStatus(`Processing sumCaseID:: ${id[processedCount][0]}`, false);

                // Call Vertex AI Search asynchronously and add the promise to promiseMap
                promiseMap.set(processedCount, callGeminiMultitModal(processedCount, full_prompt, null, null, config)
                    .then(async result => {
                        let output = result.output;
                        let status = result.status_code;
                        let rowNum = result.id;
                        
                        // Check the summary first
                        if (status === 200) {
                            await processResponse(
                                rowNum,
                                output,
                                toSummarize[rowNum][0],
                                summaryColumn,
                                summarization_qualityColumn,
                                groundednessColumn,
                                fulfillmentColumn,
                                summarization_helpfulnesColumn,
                                summarization_verbosityTimeColumn,
                                config, context);
                            
                            appendLog(`sumCaseID:: ${rowNum} Summary is: ${output}`);
                        } else {
                            appendError(`sumCaseID:: ${rowNum} Got error in summary: ${output}`, null);
                        }
                        
                       
                    })
                    .catch(error => {
                        numfails++;
                        stopProcessing = true;
                        appendError(`Error for testCaseID: ${error.id} calling callVertexAISearch`, error);
                        
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
                stoppedReason += ` Empty ID encountered after ${processedCount-1} summary cases.`;
            }
            var summary = `Finished! Successful: ${(processedCount - numfails) - 1}. ${stoppedReason}`;
            appendLog(summary);

            showStatus(summary, numfails > 0);
            
            // autofit the content
            currentWorksheet.getUsedRange().format.autofitColumns();
            currentWorksheet.getUsedRange().format.autofitRows();
            await context.sync();


        } catch (error) {
            appendError(`Caught Exception in createSummarizationData `, error);
            showStatus(`Caught Exception in createSummarizationData: ${JSON.stringify(error)}`, true);
            throw error;
        }

    });

}

export async function stopSummarizationData() { 
    stopProcessing = true; // Set the stop signal flag
    appendLog("Cancel Tests Clicked. Stopping  execution...");
}



async function processResponse(rowNum,
    output,
    textToSummarize,
    summaryColumn,
    summarization_qualityColumn,
    groundednessColumn,
    fulfillmentColumn,
    summarization_helpfulnesColumn,
    summarization_verbosityTimeColumn,
    config, context) {
    
    const token = config.accessToken;
    const prompt = config.prompt;
    const projectId = config.vertexAIProjectID;
    const location = config.vertexAILocation;
    const eval_url = `https://${location}-aiplatform.googleapis.com/v1beta1/projects/${projectId}/locations/${location}:evaluateInstances`;

    // Set the summary
    const cell_summary = summaryColumn.getRange().getCell(rowNum, 0);
    cell_summary.clear(Excel.ClearApplyTo.formats);
    cell_summary.values = [[output]];
    context.sync();
    

    try {

        
        // get the summarization_quality
        var summarization_quality_input = {
            summarization_quality_input: {
                metric_spec: {},
                instance: {
                    prediction: `${output}`,
                    instruction: `${prompt}`,
                    context: `${textToSummarize}`,
                }
            }
        };

        const response = await callVertexAI(eval_url, token, summarization_quality_input);
        if (response.status === 200) {
            // Set the summarization_quality
            const cell_summarization_quality = summarization_qualityColumn.getRange().getCell(rowNum, 0);
            cell_summarization_quality.clear(Excel.ClearApplyTo.formats);
            cell_summarization_quality.values = [[response.json_output.summarizationQualityResult.score]];
        } else {
            throw Error(`Error geting summarization_quality. Error code: ${response.status_code}`);
        }
        
    } catch (err) {
        appendError(`sumCaseID: ${rowNum} Error getting Summary Eval. Error: ${err.message} `, err);
        const cell_status = statusColumn.getRange().getCell(rowNum, 0);
        cell_status.clear(Excel.ClearApplyTo.formats);
        cell_status.format.fill.color = '#FFCCCB';
        cell_status.values = [["Failed. Error: " + err.message]];
    } finally {
        context.sync();
    }
   
   
}


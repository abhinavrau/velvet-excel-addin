import { appendError, appendLog, showStatus } from "./ui.js";

import {
  NotAuthenticatedError,
  PermissionDeniedError,
  QuotaError,
  VertexAIError,
} from "./common.js";
import { callGeminiMultitModal } from "./vertex_ai.js";

function getColumn(table, columnName) {
  try {
    const column = table.columns.getItemOrNullObject(columnName);
    column.load();
    return column;
  } catch (error) {
    appendError("Error getColumn:", error);
    showStatus(`Exception when getting column: ${JSON.stringify(error)}`, true);
  }
}

export async function getSyntheticQAConfig() {
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
        systemInstruction: valueColumn.values[4][0],
        batchSize: valueColumn.values[5][0],
        timeBetweenCallsInSec: valueColumn.values[6][0],
        accessToken: $("#access-token").val(),
        responseMimeType: "application/json",
      };
    } catch (error) {
      appendError(`Caught Exception in Gemini createConfig`, error);
      showStatus(`Caught Exception in Gemini createConfig: ${error}`, true);
      return null;
    }
  });
  return config;
}

var stopProcessing;

export async function createSyntheticQAData(config) {
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
      const fileUriColumn = getColumn(testCasesTable, "GCS File URI");
      const mimeTypeColumn = getColumn(testCasesTable, "Mime Type");
      const generatedQuestionColumn = getColumn(testCasesTable, "Generated Question");
      const expectedAnswerColumn = getColumn(testCasesTable, "Expected Answer");
      const reasoningAColumn = getColumn(testCasesTable, "Reasoning");
      const statusColumn = getColumn(testCasesTable, "Status");
      const responseTimeColumn = getColumn(testCasesTable, "Response Time");

      testCasesTable.rows.load("count");
      await context.sync();

      if (config.accessToken === null || config.accessToken === "") {
        showStatus(`Access token is empty`, true);
        appendError(`Error in createSyntheticQAData: Access token is empty`, null);
        return;
      }

      if (fileUriColumn.isNullObject || idColumn.isNullObject) {
        showStatus(
          `Error in createSyntheticQAData: No fileUriColumn or ID column found in Test Cases Table. Make sure there is an ID and Query column in the Test Cases Table.`,
          true,
        );
        return;
      }

      let processedCount = 1;
      let id = idColumn.values;
      let fileUri = fileUriColumn.values;
      let mimeType = mimeTypeColumn.values;
      let prompt = "Generate 1 question and answer";

      let numfails = 0;
      const countRows = testCasesTable.rows.count;

      // map of promises
      const promiseMap = new Map();

      stopProcessing = false;
      // Iterate rows on the  table. Stop when end of table or ID column is empty
      while (
        processedCount <= countRows &&
        id[processedCount][0] !== null &&
        id[processedCount][0] !== ""
      ) {
        // Batch the calls to Vertex AI since there are throuput checks in place.\
        if (processedCount % config.batchSize === 0) {
          // delay calls with apropriate time
          await new Promise((r) => setTimeout(r, config.timeBetweenCallsInSec * 1000));
        }
        // Stop processing if there errors
        if (stopProcessing) {
          appendLog("Stopping execution.", null);
          break;
        }
        appendLog(`genSynQID: ${id[processedCount][0]} Started Processing...`);
        showStatus(`Processing genSynQID: ${id[processedCount][0]}`, false);
        // Call Vertex AI Search asynchronously and add the promise to promiseMap
        promiseMap.set(
          processedCount,
          callGeminiMultitModal(
            processedCount,
            prompt,
            fileUri[processedCount][0],
            mimeType[processedCount][0],
            config,
          )
            .then(async (result) => {
              let output = result.output;
              let status = result.status_code;
              let rowNum = result.id;

              // Check the summary first
              if (status == 200) {
                await processResponse(
                  rowNum,
                  output,
                  generatedQuestionColumn,
                  expectedAnswerColumn,
                  reasoningAColumn,
                  statusColumn,
                  responseTimeColumn,
                  config,
                  context,
                );
                appendLog(`genSynQID: ${rowNum} Generated Question and Answer.`);
              }
            })
            .catch((error) => {
              numfails++;
              stopProcessing = true;
              if (
                error instanceof NotAuthenticatedError ||
                error instanceof QuotaError ||
                error instanceof VertexAIError ||
                error instanceof PermissionDeniedError
              ) {
                appendError(`Error for testCaseID: ${error.id} calling callVertexAISearch`, error);
              } else {
                appendError(`Error calling callVertexAISearch`, error);
              }
            }),
        );

        processedCount++;
      } // end while

      // wait for all the calls to finish
      await Promise.allSettled(promiseMap.values());
      var stoppedReason = "";
      if (numfails > 0) {
        stoppedReason = `Failed: ${numfails}. See logs for details.`;
      }
      if (
        processedCount <= countRows &&
        (id[processedCount][0] === null || id[processedCount][0] === "")
      ) {
        stoppedReason += ` Empty ID encountered after ${processedCount - 1} test cases.`;
      }
      var summary = `Finished! Successful: ${processedCount - numfails - 1}. ${stoppedReason}`;
      appendLog(summary);

      showStatus(summary, numfails > 0);

      // autofit the content
      currentWorksheet.getUsedRange().format.autofitColumns();
      currentWorksheet.getUsedRange().format.autofitRows();
      await context.sync();
    } catch (error) {
      appendError(`Caught Exception in executeTests `, error);
      showStatus(`Caught Exception in executeTests: ${JSON.stringify(error)}`, true);
      throw error;
    }
  });
}

export async function stopSyntheticData() {
  stopProcessing = true; // Set the stop signal flag
  appendLog("Cancel Tests Clicked. Stopping  execution...");
}

async function processResponse(
  rowNum,
  output,
  generatedQuestionColumn,
  expectedAnswerColumn,
  reasoningColumn,
  statusColumn,
  responseTimeColumn,
  config,
  context,
) {
  try {
    // Set the generated question
    var response = JSON.parse(output);

    const cell_generatedQuestion = generatedQuestionColumn.getRange().getCell(rowNum, 0);
    cell_generatedQuestion.clear(Excel.ClearApplyTo.formats);
    cell_generatedQuestion.values = [[response.question]];

    // Set the answer
    const cell_expectedAnswer = expectedAnswerColumn.getRange().getCell(rowNum, 0);
    cell_expectedAnswer.clear(Excel.ClearApplyTo.formats);
    cell_expectedAnswer.values = [[response.answer]];

    // Set the reasoning
    /*   const cell_reasoning = reasoningColumn.getRange().getCell(rowNum, 0);
          cell_reasoning.clear(Excel.ClearApplyTo.formats);
          cell_reasoning.values = [[response.reasoning]]; */

    // Set the reasoning
    const cell_status = statusColumn.getRange().getCell(rowNum, 0);
    cell_status.clear(Excel.ClearApplyTo.formats);
    cell_status.values = [["Success"]];
  } catch (err) {
    appendError(`testCaseID: ${rowNum} Error getting Similarity. Error: ${err.message} `, err);
    const cell_status = statusColumn.getRange().getCell(rowNum, 0);
    cell_status.clear(Excel.ClearApplyTo.formats);
    cell_status.format.fill.color = "#FFCCCB";
    cell_status.values = [["Failed. Error: " + err.message]];
  } finally {
    await context.sync();
  }
}

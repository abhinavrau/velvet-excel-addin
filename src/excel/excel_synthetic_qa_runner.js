import { TaskRunner } from "../task_runner.js";

import { appendError, appendLog, showStatus } from "../ui.js";
import { callGeminiMultitModal } from "../vertex_ai.js";
import { getColumn } from "./excel_common.js";

export class SyntheticQARunner extends TaskRunner {
  constructor() {
    super();
  }

  async getSyntheticQAConfig() {
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
        appendError(`Caught Exception in getSyntheticQAConfig: ${error} `, error);
        showStatus(`Caught Exception in getSyntheticQAConfig: ${error}`, true);
        return null;
      }
    });
    return config;
  }

  async createSyntheticQAData(config) {
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
        this.idColumn = getColumn(testCasesTable, "ID");
        this.fileUriColumn = getColumn(testCasesTable, "GCS File URI");
        this.mimeTypeColumn = getColumn(testCasesTable, "Mime Type");
        this.generatedQuestionColumn = getColumn(testCasesTable, "Generated Question");
        this.expectedAnswerColumn = getColumn(testCasesTable, "Expected Answer");
        this.reasoningAColumn = getColumn(testCasesTable, "Reasoning");
        this.statusColumn = getColumn(testCasesTable, "Status");
        this.responseTimeColumn = getColumn(testCasesTable, "Response Time");

        testCasesTable.rows.load("count");
        await context.sync();

        if (config.accessToken === null || config.accessToken === "") {
          showStatus(`Access token is empty`, true);
          appendError(`Error in createSyntheticQAData: Access token is empty`, null);
          return;
        }

        if (this.fileUriColumn.isNullObject || this.idColumn.isNullObject) {
          showStatus(
            `Error in createSyntheticQAData: No fileUriColumn or ID column found in Test Cases Table. Make sure there is an ID and Query column in the Test Cases Table.`,
            true,
          );
          return;
        }

        const countRows = testCasesTable.rows.count;

        await this.processsAllRows(context, config, countRows, this.idColumn.values);

        // autofit the content
        currentWorksheet.getUsedRange().format.autofitColumns();
        currentWorksheet.getUsedRange().format.autofitRows();
        await context.sync();
      } catch (error) {
        appendError(`Caught Exception in createSyntheticQAData `, error);
        showStatus(`Caught Exception in createSyntheticQAData: ${JSON.stringify(error)}`, true);
        throw error;
      }
    });
  }

  async stopSyntheticData() {
    this.cancelPressed = true;
    appendLog("Cancel Requested. Stopping SearchTests execution...");
  }

  async getResultFromExternalAPI(rowNum, config) {
    let fileUri = this.fileUriColumn.values;
    let mimeType = this.mimeTypeColumn.values;
    let prompt = "Generate 1 question and answer";
    return callGeminiMultitModal(rowNum, prompt, fileUri[rowNum][0], mimeType[rowNum][0], config);
  }

  async processRow(response_json, context, config, rowNum) {
    try {
      const output = response_json.candidates[0].content.parts[0].text;
      // Set the generated question
      const response = JSON.parse(output);

      const cell_generatedQuestion = this.generatedQuestionColumn.getRange().getCell(rowNum, 0);
      cell_generatedQuestion.clear(Excel.ClearApplyTo.formats);
      cell_generatedQuestion.values = [[response.question]];

      // Set the answer
      const cell_expectedAnswer = this.expectedAnswerColumn.getRange().getCell(rowNum, 0);
      cell_expectedAnswer.clear(Excel.ClearApplyTo.formats);
      cell_expectedAnswer.values = [[response.answer]];

      // Set the reasoning
      /*   const cell_reasoning = this.reasoningColumn.getRange().getCell(rowNum, 0);
            cell_reasoning.clear(Excel.ClearApplyTo.formats);
            cell_reasoning.values = [[response.reasoning]]; */

      // Set the reasoning
      const cell_status = this.statusColumn.getRange().getCell(rowNum, 0);
      cell_status.clear(Excel.ClearApplyTo.formats);
      cell_status.values = [["Success"]];
    } catch (err) {
      appendError(`testCaseID: ${rowNum} Error setting QA. Error: ${err.message} `, err);
      const cell_status = this.statusColumn.getRange().getCell(rowNum, 0);
      cell_status.clear(Excel.ClearApplyTo.formats);
      cell_status.format.fill.color = "#FFCCCB";
      cell_status.values = [["Failed. Error: " + err.message]];
    } finally {
      //await context.sync();
    }
  }
}

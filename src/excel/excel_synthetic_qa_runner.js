import { TaskRunner } from "../task_runner.js";

import {
  findIndexByColumnsNameIn2DArray,
  getFileExtensionFromUri,
  mapGeminiSupportedMimeTypes,
  mapQuestionAnsweringScore,
} from "../common.js";
import { appendError, appendLog, showStatus } from "../ui.js";
import { callGeminiMultitModal } from "../vertex_ai.js";
import { getColumn } from "./excel_common.js";

const AVG_QA_QUALITY_CELL = "B20";

export class SyntheticQARunner extends TaskRunner {
  constructor() {
    super();
    this.synthQATaskPromiseSet = new Set();
    this.generateQualityEval_throttled = this.throttle((a, b, c) =>
      this.generateQnAQualityEval(a, b, c),
    );
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
        const configColumn = getColumn(configTable, "Config");
        await context.sync();

        config = {
          vertexAIProjectID:
            valueColumn.values[
              findIndexByColumnsNameIn2DArray(configColumn.values, "Vertex AI Project ID")
            ][0],
          vertexAILocation:
            valueColumn.values[
              findIndexByColumnsNameIn2DArray(configColumn.values, "Vertex AI Location")
            ][0],
          model:
            valueColumn.values[
              findIndexByColumnsNameIn2DArray(configColumn.values, "Gemini Model ID")
            ][0],
          systemInstruction:
            valueColumn.values[
              findIndexByColumnsNameIn2DArray(configColumn.values, "System Instructions")
            ][0],
          prompt:
            valueColumn.values[findIndexByColumnsNameIn2DArray(configColumn.values, "Prompt")][0],
          qaQualityFlag:
            valueColumn.values[
              findIndexByColumnsNameIn2DArray(configColumn.values, "Generate Q & A Quality")
            ][0],
          qAQualityPrompt:
            valueColumn.values[
              findIndexByColumnsNameIn2DArray(configColumn.values, "Q & A Quality Prompt")
            ][0],
          qAQualityModel:
            valueColumn.values[
              findIndexByColumnsNameIn2DArray(configColumn.values, "Q & A Quality Model ID")
            ][0],
          batchSize: parseInt(
            valueColumn.values[
              findIndexByColumnsNameIn2DArray(configColumn.values, "Batch Size (1-10)")
            ][0],
          ),
          timeBetweenCallsInSec: parseInt(
            valueColumn.values[
              findIndexByColumnsNameIn2DArray(
                configColumn.values,
                "Time between Batches in Seconds (1-10)",
              )
            ][0],
          ),
          accessToken: $("#access-token").val(),
          responseMimeType: "application/json",
        };
      } catch (error) {
        appendError(
          `Caught Exception in getSyntheticQAConfig: ${error}. Make sure you are in the right sheet`,
          error,
        );
        showStatus(
          `Caught Exception in getSyntheticQAConfig: ${error}. Make sure you are in the right sheet`,
          true,
        );
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

        // check if the right data table is there
        const synthTestCaseTable = currentWorksheet.tables.getItem(
          `${worksheetName}.SyntheticQATable`,
        );
        await context.sync();
        if (synthTestCaseTable === null) {
          appendLog(
            `Error: ${worksheetName}.SyntheticQATable not found in current sheet. Make sure you are in the right sheet`,
          );
          showStatus(
            `Error: Synthetic Q&A table not found in current sheet. Make sure you are in the right sheet`,
            true,
          );
          return null;
        }
        const testCasesTable = currentWorksheet.tables.getItem(`${worksheetName}.SyntheticQATable`);
        this.idColumn = getColumn(testCasesTable, "ID");
        this.fileUriColumn = getColumn(testCasesTable, "GCS File URI");
        this.generatedQuestionColumn = getColumn(testCasesTable, "Generated Question");
        this.expectedAnswerColumn = getColumn(testCasesTable, "Expected Answer");
        this.qualityColumn = getColumn(testCasesTable, "Q & A Quality");

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

        const run_results = await this.processsAllRows(
          context,
          config,
          countRows,
          this.idColumn.values,
        );

        await this.addSynthQARunToTable(context, config, worksheetName, run_results);

        // autofit the content
        currentWorksheet.getUsedRange().format.autofitColumns();
        //currentWorksheet.getUsedRange().format.autofitRows();

        await context.sync();
      } catch (error) {
        appendError(`Caught Exception in createSyntheticQAData `, error);
        showStatus(`Caught Exception in createSyntheticQAData: ${JSON.stringify(error)}`, true);
        throw error;
      }
    });
  }

  async getResultFromVertexAI(rowNum, config) {
    let fileUri = this.fileUriColumn.values;
    let mimeType = mapGeminiSupportedMimeTypes[getFileExtensionFromUri(fileUri[rowNum][0])];

    return await callGeminiMultitModal(
      rowNum,
      config.prompt,
      config.systemInstruction,
      fileUri[rowNum][0],
      mimeType,
      config.model,
      config.responseMimeType,
      config,
    );
  }
  async waitForTaskstoFinish() {
    await Promise.allSettled(this.synthQATaskPromiseSet.values());
  }

  async cancelAllTasks() {
    // call abort here for any throttled tasks
    appendLog(`Cancel Requested for SyntheticQAData Tasks`);
  }

  async processRow(response_json, context, config, rowNum) {
    let numCallsMade = 0;
    try {
      // Get the output from the response
      const output = response_json.candidates[0].content.parts[0].text;

      // Parse it as json since that's the format we requested
      const response = JSON.parse(output);

      // set the generated question
      const cell_generatedQuestion = this.generatedQuestionColumn.getRange().getCell(rowNum, 0);
      cell_generatedQuestion.clear(Excel.ClearApplyTo.formats);
      cell_generatedQuestion.values = [[response.question]];

      // Set the generated answer
      const cell_expectedAnswer = this.expectedAnswerColumn.getRange().getCell(rowNum, 0);
      cell_expectedAnswer.clear(Excel.ClearApplyTo.formats);
      cell_expectedAnswer.values = [[response.answer]];

      // call to get quality if flag is set
      if (config.qaQualityFlag) {
        this.synthQATaskPromiseSet.add(
          this.generateQualityEval_throttled(config, response, rowNum),
        );
        ++numCallsMade;
      }
    } catch (err) {
      appendError(`testCaseID: ${rowNum} Error setting QA. Error: ${err.message} `, err);
      const cell_status = this.generatedQuestionColumn.getRange().getCell(rowNum, 0);
      cell_status.clear(Excel.ClearApplyTo.formats);
      cell_status.format.fill.color = "#FFCCCB";
      cell_status.values = [["Failed. Error: " + err.message]];
    } finally {
      //await context.sync();
    }
    // execute the tasks
    await Promise.allSettled(this.synthQATaskPromiseSet.values());

    return numCallsMade;
  }

  async generateQnAQualityEval(config, response, rowNum) {
    try {
      appendLog(`testCaseID::${rowNum} generateQualityEval Started..`);
      let fileUri = this.fileUriColumn.values;
      let mimeType = mapGeminiSupportedMimeTypes[getFileExtensionFromUri(fileUri[rowNum][0])];

      const evalPrompt = `${config.qAQualityPrompt} # User Inputs and AI-generated Response
                        ## User Inputs

                        ### Prompt
                        ${config.systemInstruction}
                        ${config.prompt}

                        ## AI-generated Response
                        ${JSON.stringify(response)}`;

      const eval_response = await callGeminiMultitModal(
        rowNum,
        evalPrompt,
        "",
        fileUri[rowNum][0],
        mimeType,
        config.qAQualityModel,
        config.responseMimeType,
        config,
      );

      const eval_output = eval_response.output.candidates[0].content.parts[0].text;
      // since its json we get the rating tag
      const eval_json = JSON.parse(eval_output);

      // Set the eval quality
      const cell_evalQuality = this.qualityColumn.getRange().getCell(rowNum, 0);
      cell_evalQuality.clear(Excel.ClearApplyTo.formats);
      cell_evalQuality.values = [[mapQuestionAnsweringScore.get(eval_json.rating)]];
      appendLog(`testCaseID::${rowNum} generateQualityEval Finished: Rating: ${eval_json.rating}`);
    } catch (err) {
      appendError(`testCaseID: ${rowNum} Error setting Eval QA  Error: ${err.message} `, err);
      const cell_status = this.qualityColumn.getRange().getCell(rowNum, 0);
      cell_status.clear(Excel.ClearApplyTo.formats);
      cell_status.format.fill.color = "#FFCCCB";
      cell_status.values = [["Failed. Error: " + err.message]];
    }
  }

  async addSynthQARunToTable(context, config, worksheetName, run_results) {
    try {
      const synthQARunsSheet = context.workbook.worksheets.getItemOrNullObject("Synthetic QnAs");
      const runsTable = synthQARunsSheet.tables.getItemOrNullObject(
        "SyntheticQnAs.SynthQARunsTable",
      );

      synthQARunsSheet.load("name");
      runsTable.load("name");
      await context.sync();

      if (synthQARunsSheet.isNullObject) {
        appendLog(
          "Could not find 'Synthetic Questions & Answers Eval Runs' sheet.",
          new Error("Synthetic Questions & Answers Eval Runs sheet not found"),
        );
        return;
      }
      if (runsTable.isNullObject) {
        appendLog(
          "Could not find 'SynthQARunsTable' in 'Synthetic Questions & Answers Eval Runs' sheet.",
          new Error("TestRunsTable not found"),
        );
        return;
      }

      const dataRange = runsTable.getDataBodyRange();
      dataRange.load("values, rowCount");
      await context.sync();

      const testCasesSheet = context.workbook.worksheets.getItem(worksheetName);

      const avgQualityScoreRange = testCasesSheet.getRange(AVG_QA_QUALITY_CELL);
      avgQualityScoreRange.load("values");

      await context.sync();

      const avgQualityScore = avgQualityScoreRange.values[0][0];

      const newRowData = [
        worksheetName,
        new Date().toLocaleString(),
        config.vertexAIProjectID,
        run_results.numSuccessful,
        run_results.numFails,
        avgQualityScore,
      ];

      let tableData = dataRange.values;
      let rowIndex = -1;
      for (let i = 0; i < dataRange.rowCount; i++) {
        if (tableData[i][0] === worksheetName) {
          rowIndex = i;
          break;
        }
      }

      if (rowIndex !== -1) {
        // Update existing row data in our local array
        tableData[rowIndex] = newRowData;
        appendLog(
          `Updating Row index:${rowIndex} 'Synthetic Questions & Answers Eval Runs' table for ${worksheetName}.`,
        );
      } else {
        // Add new row data to our local array
        tableData.push(newRowData);
        appendLog(
          `Inserting New Row 'Synthetic Questions & Answers Eval Runs' table for ${worksheetName}.`,
        );
      }

      // filter out the rows in array runsTable.rows  except header row
      runsTable.rows.load("items");
      await context.sync();

      const dataRows = runsTable.rows.items;

      // Clear all the rows in the table exept the header
      if (dataRows.length > 0) {
        runsTable.rows.deleteRows(dataRows);
      }

      await context.sync();
      // remove empty row with empty values
      tableData = tableData.filter((row) => row !== null && row[0] !== "");

      // Add all rows back to the table
      if (tableData.length > 0) {
        runsTable.rows.add(null, tableData);
      }

      await context.sync();

      // Add hyperlinks
      const newDataRange = runsTable.getDataBodyRange();
      newDataRange.load("values, rowCount");
      await context.sync();
      for (let i = 0; i < newDataRange.rowCount; i++) {
        const cellToUpdate = newDataRange.getCell(i, 0);
        const sheetName = newDataRange.values[i][0];
        cellToUpdate.hyperlink = {
          textToDisplay: sheetName,
          screenTip: `Navigate to the '${sheetName}' worksheet`,
          documentReference: `'${sheetName}'!A1`,
        };
      }

      await context.sync();
      appendLog(
        `Finished addSynthQARunToTable 'Synthetic Questions & Answers Eval Runs' table for ${worksheetName}.`,
      );
    } catch (error) {
      appendError(`Error in addSynthQARunToTable: ${error.message}`, error);
      showStatus(
        `Error adding row to Synthetic Questions & Answers Eval Runs table: ${JSON.stringify(error)}`,
        true,
      );
    }
  }
}

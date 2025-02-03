import { findIndexByColumnsNameIn2DArray, mapQuestionAnsweringScore } from "../common.js";
import { TaskRunner } from "../task_runner.js";
import { appendError, appendLog, showStatus } from "../ui.js";
import { callGeminiMultitModal } from "../vertex_ai.js";
import { getColumn } from "./excel_common.js";
export class SummarizationRunner extends TaskRunner {
  constructor() {
    super();
    this.summarizationEval_throttle = this.throttle((a, b, c, d) =>
      this.generateSummarizationQualityEval(a, b, c, d),
    );
    this.summaryEvalTaskPromiseSet = new Set();
  }

  async getSummarizationConfig() {
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
          summarizationQualityFlag:
            valueColumn.values[
              findIndexByColumnsNameIn2DArray(configColumn.values, "Generate Summarization Quality")
            ][0],
          summarizationQualityPrompt:
            valueColumn.values[
              findIndexByColumnsNameIn2DArray(configColumn.values, "Summarization Quality Prompt")
            ][0],
          summarizationQualityModel:
            valueColumn.values[
              findIndexByColumnsNameIn2DArray(configColumn.values, "Summarization Quality Model ID")
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

  async createSummarizationData(config) {
    if (config == null) {
      return;
    }

    await Excel.run(async (context) => {
      try {
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        currentWorksheet.load("name");
        await context.sync();
        const worksheetName = currentWorksheet.name;

        const testCasesTable = currentWorksheet.tables.getItem(
          `${worksheetName}.SummarizationTestCasesTable`,
        );
        this.idColumn = getColumn(testCasesTable, "ID");
        this.toSummarizeColumn = getColumn(testCasesTable, "Context");
        this.summaryColumn = getColumn(testCasesTable, "Summary");
        this.summarization_qualityColumn = getColumn(testCasesTable, "summarization_quality");
        testCasesTable.rows.load("count");
        await context.sync();

        if (config.accessToken === null || config.accessToken === "") {
          showStatus(`Access token is empty`, true);
          appendError(`Error in createSummarizationData: Access token is empty`, null);
          return;
        }

        if (this.toSummarizeColumn.isNullObject || this.idColumn.isNullObject) {
          showStatus(
            `Error in createSummarizationData: No fileUriColumn or ID column found in Test Cases Table. Make sure there is an ID and Query column in the Test Cases Table.`,
            true,
          );
          return;
        }

        const countRows = testCasesTable.rows.count;

        await this.processsAllRows(context, config, countRows, this.idColumn.values);

        // autofit the content
        //currentWorksheet.getUsedRange().format.autofitColumns();
        //currentWorksheet.getUsedRange().format.autofitRows();
        
        await context.sync();
      } catch (error) {
        appendError(`Caught Exception in createSummarizationData `, error);
        showStatus(`Caught Exception in createSummarizationData: ${JSON.stringify(error)}`, true);
        throw error;
      }
    });
  }

  async getResultFromVertexAI(rowNum, config) {
    const toSummarize = this.toSummarizeColumn.values;
    const full_prompt = config.prompt + " Text to summarize: " + toSummarize[rowNum][0];
    return await callGeminiMultitModal(
      rowNum,
      full_prompt,
      null,
      null,
      null,
      config.model,
      config.responseMimeType,
      config,
    );
  }

  async waitForTaskstoFinish() {
    await Promise.allSettled(this.summaryEvalTaskPromiseSet.values());
  }
  async cancelAllTasks() {
    await this.summaryEval_throttle.abort();
    await this.summaryEvalTaskPromiseSet.clear();
    appendLog(`Cancel Requested for Summarization Tasks`);
  }

  async processRow(response_json, context, config, rowNum) {
    const toSummarize = this.toSummarizeColumn.values;
    const textToSummarize = toSummarize[rowNum][0];

    // Set the generated summary
    const generatedSummary = response_json.candidates[0].content.parts[0].text;
    const cell_summary = this.summaryColumn.getRange().getCell(rowNum, 0);
    cell_summary.clear(Excel.ClearApplyTo.formats);
    cell_summary.values = [[generatedSummary]];
    let numCallsMade = 0;

    // summary fulfillment
    if (config.summarizationQualityFlag) {
      this.summaryEvalTaskPromiseSet.add(
        this.summarizationEval_throttle(config, textToSummarize, generatedSummary, rowNum),
      );
      ++numCallsMade;
    }

    // execute the tasks
    await Promise.allSettled(this.summaryEvalTaskPromiseSet.values());

    // return number of calls made
    return numCallsMade;
  }

  async generateSummarizationQualityEval(config, textToSummarize, generatedSummary, rowNum) {
    try {
      appendLog(`testCaseID::${rowNum} generateSummaryQualityEval Started..`);

      const evalPrompt = `${config.summarizationQualityPrompt} # User Inputs and AI-generated Response
                        ## User Inputs
                        ${textToSummarize}
                        ### Prompt
                        ${config.systemInstruction}
                        ${config.prompt}

                        ## AI-generated Response
                        ${generatedSummary}`;

      const eval_response = await callGeminiMultitModal(
        rowNum,
        evalPrompt,
        "",
        null,
        null,
        config.summarizationQualityModel,
        "application/json", // pass this since we want json back
        config,
      );

      const eval_output = eval_response.output.candidates[0].content.parts[0].text;
      // since its json we get the rating tag
      const eval_json = JSON.parse(eval_output);

      // Set the eval quality
      const cell_evalQuality = this.summarization_qualityColumn.getRange().getCell(rowNum, 0);
      cell_evalQuality.clear(Excel.ClearApplyTo.formats);
      cell_evalQuality.values = [[mapQuestionAnsweringScore.get(eval_json.rating)]];
      appendLog(
        `testCaseID::${rowNum} summarizationQualityEval Finished: Rating: ${eval_json.rating}`,
      );
    } catch (err) {
      appendError(
        `testCaseID: ${rowNum} Error setting SummarizationQualityEval  Error: ${err.message} `,
        err,
      );
      const cell_status = this.summarization_qualityColumn.getRange().getCell(rowNum, 0);
      cell_status.clear(Excel.ClearApplyTo.formats);
      cell_status.format.fill.color = "#FFCCCB";
      cell_status.values = [["Failed. Error: " + err.message]];
    }
  }
}

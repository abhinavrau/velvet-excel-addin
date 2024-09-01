import { AbortError } from "../../lib/p-throttle.js";
import {
  mapSummaryHelpfulnessScore,
  mapSummaryQualityScore,
  mapSummaryVerbosityScore,
  mapTextgenFulfillmentScore,
  mapTextgenGroundednessScore,
} from "../common.js";
import { TaskRunner } from "../task_runner.js";
import { appendError, appendLog, showStatus } from "../ui.js";
import { callGeminiMultitModal, callVertexAI } from "../vertex_ai.js";
import { getColumn } from "./excel_common.js";
export class SummarizationRunner extends TaskRunner {
  constructor() {
    super();
    this.summaryEval_throttle = this.throttle((a, b, c, d, e, f, g, h) =>
      this.callSummaryEval(a, b, c, d, e, f, g, h),
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
        await context.sync();

        config = {
          vertexAIProjectID: valueColumn.values[1][0],
          vertexAILocation: valueColumn.values[2][0],
          model: valueColumn.values[3][0],
          prompt: valueColumn.values[4][0],
          generateSummarizationQuality: valueColumn.values[5][0],
          generateSummarizationHelpfulness: valueColumn.values[6][0],
          generateSummarizationVerbosity: valueColumn.values[7][0],
          generateGroundedness: valueColumn.values[8][0],
          generateFulfillment: valueColumn.values[9][0],
          batchSize: parseInt(valueColumn.values[10][0]),
          timeBetweenCallsInSec: parseInt(valueColumn.values[11][0]),
          accessToken: $("#access-token").val(),

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

        const testCasesTable = currentWorksheet.tables.getItem(`${worksheetName}.TestCasesTable`);
        this.idColumn = getColumn(testCasesTable, "ID");
        this.toSummarizeColumn = getColumn(testCasesTable, "Context");
        this.summaryColumn = getColumn(testCasesTable, "Summary");
        this.summarization_qualityColumn = getColumn(testCasesTable, "summarization_quality");
        this.groundednessColumn = getColumn(testCasesTable, "groundedness");
        this.fulfillmentColumn = getColumn(testCasesTable, "fulfillment");
        this.summarization_helpfulnesColumn = getColumn(
          testCasesTable,
          "summarization_helpfulness",
        );
        this.summarization_verbosityColumn = getColumn(testCasesTable, "summarization_verbosity");

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

  async getResultFromVertexAI(rowNum, config) {
    const toSummarize = this.toSummarizeColumn.values;
    const full_prompt = config.prompt + " Text to summarize: " + toSummarize[rowNum][0];
    return await callGeminiMultitModal(rowNum, full_prompt, null, null, null, config.model, config);
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
    const token = config.accessToken;
    const prompt = config.prompt;
    const projectId = config.vertexAIProjectID;
    const location = config.vertexAILocation;
    const eval_url = `https://${location}-aiplatform.googleapis.com/v1beta1/projects/${projectId}/locations/${location}:evaluateInstances`;
    const output = response_json.candidates[0].content.parts[0].text;

    const toSummarize = this.toSummarizeColumn.values;
    const textToSummarize = toSummarize[rowNum][0];
    // Set the summary
    const cell_summary = this.summaryColumn.getRange().getCell(rowNum, 0);
    cell_summary.clear(Excel.ClearApplyTo.formats);
    cell_summary.values = [[output]];
    let numCallsSoFar = 0;
    //context.sync();

    // summary quality
    if (config.generateSummarizationQuality) {
      appendLog(`testCaseID::${rowNum} summarizationQualityResult Started..`);
      var summarization_quality_input = {
        summarization_quality_input: {
          metric_spec: {},
          instance: {
            prediction: `${output}`,
            instruction: `${prompt}`,
            context: `${textToSummarize}`,
          },
        },
      };

      this.summaryEvalTaskPromiseSet.add(
        this.summaryEval_throttle(
          eval_url,
          token,
          summarization_quality_input,
          this.summarization_qualityColumn,
          "summarizationQualityResult",
          rowNum,
          context,
          mapSummaryQualityScore,
        ),
      );
      numCallsSoFar++;
    }

    // summary helpfulness
    if (config.generateSummarizationHelpfulness) {
      appendLog(`testCaseID::${rowNum} summarizationHelpfulness Started..`);

      var summarization_helpfulness_input = {
        summarization_helpfulness_input: {
          metric_spec: {},
          instance: {
            prediction: `${output}`,
            instruction: `${prompt}`,
            context: `${textToSummarize}`,
          },
        },
      };

      this.summaryEvalTaskPromiseSet.add(
        this.summaryEval_throttle(
          eval_url,
          token,
          summarization_helpfulness_input,
          this.summarization_helpfulnesColumn,
          "summarizationHelpfulnessResult",
          rowNum,
          context,
          mapSummaryHelpfulnessScore,
        ),
      );
      numCallsSoFar++;
    }

    // summary verbosity
    if (config.generateSummarizationVerbosity) {
      // Check the flag
      appendLog(`testCaseID::${rowNum} summarizationVerbosity Started..`);

      var summarization_verbosity_input = {
        summarization_verbosity_input: {
          metric_spec: {},
          instance: {
            prediction: `${output}`,
            instruction: `${prompt}`,
            context: `${textToSummarize}`,
          },
        },
      };

      this.summaryEvalTaskPromiseSet.add(
        this.summaryEval_throttle(
          eval_url,
          token,
          summarization_verbosity_input,
          this.summarization_verbosityColumn,
          "summarizationVerbosityResult",
          rowNum,
          context,
          mapSummaryVerbosityScore,
        ),
      );
      numCallsSoFar++;
    }

    // summary groundedness
    if (config.generateGroundedness) {
      // Check the flag
      appendLog(`testCaseID::${rowNum} groundedness Started..`);

      var groundedness_input = {
        groundedness_input: {
          metric_spec: {},
          instance: {
            prediction: `${output}`,
            context: `${textToSummarize}`,
          },
        },
      };

      this.summaryEvalTaskPromiseSet.add(
        this.summaryEval_throttle(
          eval_url,
          token,
          groundedness_input,
          this.groundednessColumn,
          "groundednessResult",
          rowNum,
          context,
          mapTextgenGroundednessScore,
        ),
      );
      numCallsSoFar++;
    }

    // summary fulfillment
    if (config.generateFulfillment) {
      appendLog(`testCaseID::${rowNum} fulfillment Started..`);

      // summary fulfillment
      var fulfillment_input = {
        fulfillment_input: {
          metric_spec: {},
          instance: {
            prediction: `${output}`,
            instruction: `${prompt}`,
          },
        },
      };

      this.summaryEvalTaskPromiseSet.add(
        this.summaryEval_throttle(
          eval_url,
          token,
          fulfillment_input,
          this.fulfillmentColumn,
          "fulfillmentResult",
          rowNum,
          context,
          mapTextgenFulfillmentScore,
        ),
      );
      numCallsSoFar++;
    }

    // execute the tasks
    await Promise.allSettled(this.summaryEvalTaskPromiseSet.values());

    // return number of calls made
    return numCallsSoFar;
  }

  async callSummaryEval(
    eval_url,
    token,
    summarization_eval_input,
    summarization_evalColumn,
    resultpropertyName,
    rowNum,
    context,
    mapScoreString,
  ) {
    try {
      const response = await callVertexAI(eval_url, token, summarization_eval_input);
      // Set the summarization_quality
      const cell_summarization_quality = summarization_evalColumn.getRange().getCell(rowNum, 0);
      cell_summarization_quality.clear(Excel.ClearApplyTo.formats);
      let score = response.json_output[resultpropertyName].score;
      if (mapScoreString !== null) {
        cell_summarization_quality.values = [[mapScoreString.get(score)]];
      } else {
        cell_summarization_quality.values = [[score]];
      }
    } catch (err) {
      if (err instanceof AbortError) {
        appendLog("Aborting SummarizationTask");
      } else {
        appendError(
          `testCaseID: ${rowNum} Error getting Summary Eval. Error: ${err.message} `,
          err,
        );
        const cell_status = summarization_evalColumn.getRange().getCell(rowNum, 0);
        cell_status.clear(Excel.ClearApplyTo.formats);
        cell_status.format.fill.color = "#FFCCCB";
        cell_status.values = [["Failed. Error: " + err.message]];
      }
    }
    appendLog(`testCaseID::${rowNum} ${resultpropertyName} Finished.`);
  }
}

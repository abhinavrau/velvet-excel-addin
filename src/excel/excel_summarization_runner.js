import { TaskRunner } from "../task_runner.js";

import { appendError, appendLog, showStatus } from "../ui.js";
import { callGeminiMultitModal, callVertexAI } from "../vertex_ai.js";
import { getColumn } from "./excel_common.js";

export class SummarizationRunner extends TaskRunner {
  constructor() {
    super();
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
          batchSize: valueColumn.values[5][0],
          timeBetweenCallsInSec: valueColumn.values[6][0],
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

  async stopSummarizationData() {
    this.cancelPressed = true;
    appendLog("Cancel Requested. Stopping SearchTests execution...");
  }

  async getResultFromExternalAPI(rowNum, config) {
    let toSummarize = this.toSummarizeColumn.values;
    let full_prompt = config.prompt + " Text to summarize: " + toSummarize[rowNum][0];
    return callGeminiMultitModal(rowNum, full_prompt, null, null, config);
  }

  async processRow(response_json, context, config, rowNum) {
    const token = config.accessToken;
    const prompt = config.prompt;
    const projectId = config.vertexAIProjectID;
    const location = config.vertexAILocation;
    const eval_url = `https://${location}-aiplatform.googleapis.com/v1beta1/projects/${projectId}/locations/${location}:evaluateInstances`;
    const output = response_json.candidates[0].content.parts[0].text;
    let toSummarize = this.toSummarizeColumn.values;
    const textToSummarize = toSummarize[rowNum][0];
    // Set the summary
    const cell_summary = this.summaryColumn.getRange().getCell(rowNum, 0);
    cell_summary.clear(Excel.ClearApplyTo.formats);
    cell_summary.values = [[output]];
    context.sync();

    appendLog(`testCaseID::${rowNum} summarizationQualityResult Started..`);
    // summary quality
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

    await this.callSummaryEval(
      eval_url,
      token,
      summarization_quality_input,
      this.summarization_qualityColumn,
      "summarizationQualityResult",
      rowNum,
      context,
    );

    appendLog(`testCaseID::${rowNum} summarizationQuality Finished`);

    appendLog(`testCaseID::${rowNum} summarizationHelpfulness Started..`);
    // summary helpfulness
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

    await this.callSummaryEval(
      eval_url,
      token,
      summarization_helpfulness_input,
      this.summarization_helpfulnesColumn,
      "summarizationHelpfulnessResult",
      rowNum,
      context,
    );

    appendLog(`sumCaseID::${rowNum} summarizationHelpfulness Finished`);

    appendLog(`sumCaseID::${rowNum} summarizationVerbosity Started..`);

    // summary verbosity
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

    await this.callSummaryEval(
      eval_url,
      token,
      summarization_verbosity_input,
      this.summarization_verbosityColumn,
      "summarizationVerbosityResult",
      rowNum,
      context,
    );

    appendLog(`sumCaseID::${rowNum} summarizationVerbosity Finished`);

    appendLog(`sumCaseID::${rowNum} groundedness Started..`);
    // summary groundedness
    var groundedness_input = {
      groundedness_input: {
        metric_spec: {},
        instance: {
          prediction: `${output}`,
          context: `${textToSummarize}`,
        },
      },
    };

    await this.callSummaryEval(
      eval_url,
      token,
      groundedness_input,
      this.groundednessColumn,
      "groundednessResult",
      rowNum,
      context,
    );

    appendLog(`sumCaseID::${rowNum} groundedness Finished`);

    appendLog(`sumCaseID::${rowNum} fulfillment Started..`);

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

    await this.callSummaryEval(
      eval_url,
      token,
      fulfillment_input,
      this.fulfillmentColumn,
      "fulfillmentResult",
      rowNum,
      context,
    );
    appendLog(`sumCaseID::${rowNum} fulfillment Finished`);
  }

  async callSummaryEval(
    eval_url,
    token,
    summarization_eval_input,
    summarization_evalColumn,
    resultpropertyName,
    rowNum,
    context,
  ) {
    try {
      const response = await callVertexAI(eval_url, token, summarization_eval_input);
      if (response.status === 200) {
        // Set the summarization_quality
        const cell_summarization_quality = summarization_evalColumn.getRange().getCell(rowNum, 0);
        cell_summarization_quality.clear(Excel.ClearApplyTo.formats);
        cell_summarization_quality.values = [[response.json_output[resultpropertyName].score]];
      } else {
        throw Error(`Error geting summarization_quality. Error code: ${response.status_code}`);
      }
    } catch (err) {
      appendError(`sumCaseID: ${rowNum} Error getting Summary Eval. Error: ${err.message} `, err);
      const cell_status = summarization_evalColumn.getRange().getCell(rowNum, 0);
      cell_status.clear(Excel.ClearApplyTo.formats);
      cell_status.format.fill.color = "#FFCCCB";
      cell_status.values = [["Failed. Error: " + err.message]];
    } finally {
      //context.sync();
    }
  }
}

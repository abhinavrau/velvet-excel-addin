import {
  calculateSimilarityUsingGemini,
  callCheckGrounding,
  callVertexAISearch,
} from "../vertex_ai.js";

import { appendError, appendLog, showStatus } from "../ui.js";

import { TaskRunner } from "../task_runner.js";

import { getAccuracyFormula, getAverageFormula } from "../common.js";
import { getColumn, getSearchConfigFromActiveSheet } from "./excel_common.js";
export class ExcelSearchRunner extends TaskRunner {
  constructor() {
    super();
    this.throttled_process_summary = this.throttle((a, b, c, d, e) =>
      this.processSummary(a, b, c, d, e),
    );
    this.throttled_check_grounding = this.throttle((a, b, c, d) =>
      this.getCheckGroundingData(a, b, c, d),
    );
    this.searchTaskPromiseSet = new Set();
  }

  async getSearchConfig() {
    return getSearchConfigFromActiveSheet(true, true);
  }

  async executeSearchTests(config) {
    if (config == null) {
      return;
    }

    await Excel.run(async (context) => {
      try {
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        currentWorksheet.load("name");
        await context.sync();
        const worksheetName = currentWorksheet.name;

        const testCasesTable = currentWorksheet.tables.getItemOrNullObject(
          `${worksheetName}.TestCasesTable`,
        );
        this.queryColumn = getColumn(testCasesTable, "Query");
        this.idColumn = getColumn(testCasesTable, "ID");
        this.link_1_Column = getColumn(testCasesTable, "Actual Link 1");
        this.link_2_Column = getColumn(testCasesTable, "Actual Link 2");
        this.link_3_Column = getColumn(testCasesTable, "Actual Link 3");
        this.expected_link_1_Column = getColumn(testCasesTable, "Expected Link 1");
        this.expected_link_2_Column = getColumn(testCasesTable, "Expected Link 2");
        this.expected_link_3_Column = getColumn(testCasesTable, "Expected Link 3");
        this.link_p0Column = getColumn(testCasesTable, "First Link Match");
        this.link_top2Column = getColumn(testCasesTable, "Link in Top 2");
        this.actualSummaryColumn = getColumn(testCasesTable, "Actual Summary");
        this.expectedSummaryColumn = getColumn(testCasesTable, "Expected Summary");
        this.summaryScoreColumn = getColumn(testCasesTable, "Summary Match");
        this.checkGroundingScoreColumn = getColumn(testCasesTable, "Grounding Score");
        await context.sync();
        testCasesTable.rows.load("count");
        await context.sync();

        if (config.accessToken === null || config.accessToken === "") {
          showStatus(`Access token is empty`, true);
          appendError(`Error in executeSearchTests: Access token is empty`, null);
          return;
        }

        

        if (this.queryColumn.isNullObject || this.idColumn.isNullObject) {
          showStatus(
            `Error in executeSearchTests: No Query or ID column found in Test Cases Table. Make sure there is an ID and Query column in the Test Cases Table.`,
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

        await this.addSearchRunToTable(context, config, worksheetName, run_results);

        // autofit the content
        //currentWorksheet.getUsedRange().format.autofitColumns();
        //currentWorksheet.getUsedRange().format.autofitRows();

        await context.sync();
      } catch (error) {
        appendLog(`Caught Exception in executeSearchTests: ${error.message} `, error);
        showStatus(`Caught Exception in executeSearchTests: ${JSON.stringify(error)}`, true);
        throw error;
      }
    });
  }

  async getResultFromVertexAI(rowNum, config) {
    var query = this.queryColumn.values;
    return await callVertexAISearch(rowNum, query[rowNum][0], config);
  }

  async waitForTaskstoFinish() {
    await Promise.allSettled(this.searchTaskPromiseSet.values());
  }

  async cancelAllTasks() {
    this.throttled_process_summary.abort();
    appendLog(`Cancel Requested for Search Tasks`);
  }

  async processRow(response_json, context, config, rowNum) {
    let numCalls = 0;
    if (response_json.hasOwnProperty("summary")) {
      // process the summary using throttling since it makes an external call
      const processSummaryPromise = this.throttled_process_summary(
        context,
        config,
        rowNum,
        response_json.summary.summaryText,
        this.expectedSummaryColumn.values,
      ).then(async (callsSoFar) => {
        appendLog(`testCaseID: ${rowNum} Processed Search Summary.`);
      });

      this.searchTaskPromiseSet.add(processSummaryPromise);

      // Also call check grounding API
      if (config.genereateGrounding) {
        const checkGroundingPromise = this.throttled_check_grounding(
          context,
          config,
          rowNum,
          response_json,
        ).then(async (callsSoFar) => {
          appendLog(`testCaseID: ${rowNum} Finished Check Grounding.`);
        });

        this.searchTaskPromiseSet.add(checkGroundingPromise);
      }
    }

    // Check the documents references
    if (response_json.hasOwnProperty("results")) {
      this.checkDocumentLinks(
        context,
        rowNum,
        response_json,
        this.expected_link_1_Column.values,
        this.expected_link_2_Column.values,
      );
      appendLog(`testCaseID: ${rowNum} Processed Doc Links.`);
    }
    // execute the tasks
    await Promise.allSettled(this.searchTaskPromiseSet.values());
    numCalls += 2;
    return numCalls;
  }

  async getCheckGroundingData(context, config, rowNum, response_json) {
    try {
      const factsArray = [];
      let count = 0;
      const resultsUsedForSummary = parseInt(config.summaryResultCount); // get how many items were used to genereate summary
      response_json.results.forEach((result) => {
        // Only add extractive content as much was used by the summary
        if (count >= resultsUsedForSummary) {
          return;
        }
        // Check if 'extractive_answers' exists within 'derivedStructData' and put in factsArray
        if (result.document.derivedStructData.extractive_answers) {
          result.document.derivedStructData.extractive_answers.forEach((answer) => {
            factsArray.push({ factText: `${answer.content}` });
          });
          // Check if 'extractive_segments' exists within 'derivedStructData' and put in factsArray
        } else if (result.document.derivedStructData.extractive_segments) {
          result.document.derivedStructData.extractive_segments.forEach((segment) => {
            factsArray.push({ factText: `${answer.content}` });
          });
        }
        count++;
      });
      // Call the checkGrounding API with the summary texts and factsArray
      const { id, status_code, output } = await callCheckGrounding(
        config,
        response_json.summary.summaryText,
        factsArray,
        rowNum,
      );

      const groudingScorecell = this.checkGroundingScoreColumn.getRange().getCell(rowNum, 0);
      groudingScorecell.clear(Excel.ClearApplyTo.formats);
      if (output.supportScore) {
        groudingScorecell.values = [[output.supportScore.toString()]];
      } else {
        groudingScorecell.values = [["No support score returned"]];
      }
    } catch (error) {
      appendError(`testCaseID: ${rowNum} Error getting Grounding. Error: ${error.message} `, error);
    }
  }

  async processSummary(context, config, rowNum, result, expectedSummary) {
    // Set the actual summary
    const score_cell = this.summaryScoreColumn.getRange().getCell(rowNum, 0);
    const actualSummarycell = this.actualSummaryColumn.getRange().getCell(rowNum, 0);
    try {
     
      actualSummarycell.clear(Excel.ClearApplyTo.formats);
      actualSummarycell.values = [[result]];

      // match summaries only if they are not null or not empty
      if (expectedSummary[rowNum][0] !== null && expectedSummary[rowNum][0] !== "") {
        
        score_cell.clear(Excel.ClearApplyTo.formats);

        const response = await calculateSimilarityUsingGemini(
          rowNum,
          result,
          expectedSummary[rowNum][0],
          config,
        );

        const score = response.output;

        if (score.trim() === "same") {
          score_cell.values = [["TRUE"]];
        } else {
          score_cell.values = [["FALSE"]];
          score_cell.format.fill.color = "#FFCCCB";
          actualSummarycell.format.fill.color = "#FFCCCB";
        }
      }
      // Catch any errors here and report it in the cell. We don't want failures here to stop processing.
    } catch (err) {
      appendError(`testCaseID: ${rowNum} Error in processSummary. Error: ${err.message} `, err);
      // put the error in the cell.
      score_cell.values = [["Failed. Error: " + err.message]];
      score_cell.format.fill.color = "#FFCCCB";
      actualSummarycell.format.fill.color = "#FFCCCB";
    } finally {
      //await context.sync();
    }
  }

  checkDocumentLinks(context, rowNum, result, expectedLink1, expectedLink2) {
    var p0_result = null;
    var p2_result = null;
    const link_1_cell = this.link_1_Column.getRange().getCell(rowNum, 0);
    const link_2_cell = this.link_2_Column.getRange().getCell(rowNum, 0);
    const link_3_cell = this.link_3_Column.getRange().getCell(rowNum, 0);

    // Check for document info and linksin the metadata if it exists
    /* if (result.results[0] && result.results[0].document.hasOwnProperty("structData")) {
      link_1_cell.values = [[result.results[0].document.structData.sharepoint_ref]];
      p0_result = result.results[0].document.structData.title;
    } else  */
    if (result.results[0] && result.results[0].document.hasOwnProperty("derivedStructData")) {
      link_1_cell.values = [[result.results[0].document.derivedStructData.link]];
      p0_result = result.results[0].document.derivedStructData.link;
    }
    /* if (result.results[1] && result.results[1].document.hasOwnProperty("structData")) {
      link_2_cell.values = [[result.results[1].document.structData.sharepoint_ref]];
    } else */
    if (result.results[1] && result.results[1].document.hasOwnProperty("derivedStructData")) {
      link_2_cell.values = [[result.results[1].document.derivedStructData.link]];
      p2_result = result.results[1].document.derivedStructData.link;
    }
    /* if (result.results[2] && result.results[2].document.hasOwnProperty("structData")) {
      link_3_cell.values = [[result.results[2].document.structData.sharepoint_ref]];
    } else */
    if (result.results[2] && result.results[2].document.hasOwnProperty("derivedStructData")) {
      link_3_cell.values = [[result.results[2].document.derivedStructData.link]];
    }

    // clear the formatting in the cells
    const link_p0_cell = this.link_p0Column.getRange().getCell(rowNum, 0);
    link_p0_cell.clear(Excel.ClearApplyTo.formats);
    link_1_cell.clear(Excel.ClearApplyTo.formats);
    const top2_cell = this.link_top2Column.getRange().getCell(rowNum, 0);
    top2_cell.clear(Excel.ClearApplyTo.formats);

    // match first link with expected link
    if (p0_result !== null && p0_result === expectedLink1[rowNum][0]) {
      link_p0_cell.values = [["TRUE"]];
    } else {
      link_p0_cell.values = [["FALSE"]];
      link_p0_cell.format.fill.color = "#FFCCCB";
      link_1_cell.format.fill.color = "#FFCCCB";
    }

    // match if the top 2 links returned are in the top 2 expected links
    if (
      (p2_result !== null && p2_result === expectedLink2[rowNum][0]) ||
      p2_result === expectedLink1[rowNum][0] ||
      p0_result === expectedLink1[rowNum][0] ||
      p0_result === expectedLink2[rowNum][0]
    ) {
      top2_cell.values = [["TRUE"]];
    } else {
      top2_cell.values = [["FALSE"]];
      top2_cell.format.fill.color = "#FFCCCB";
    }
  }

  async addSearchRunToTable(context, config, worksheetName, run_results) {
    try {
      const searchEvalRunsSheet = context.workbook.worksheets.getItemOrNullObject("Search Evals");
      searchEvalRunsSheet.load();
      await context.sync();

      if (searchEvalRunsSheet.isNullObject) {
        appendLog("Could not find searchEvalRunsSheet.", new Error("Search Evals sheet not found"));
        return;
      }

      const runsTable = searchEvalRunsSheet.tables.getItemOrNullObject("SearchEvals.TestRunsTable");
      runsTable.load("name");
      await context.sync();

      if (runsTable.isNullObject) {
        appendLog(
          "Could not find 'Search Evals.TestRunsTable' in 'Search Evals' sheet.",
          new Error("TestRunsTable not found"),
        );
        return;
      }

      const summaryMatchAccuracyFormula = getAccuracyFormula(worksheetName, "Summary Match");
      const firstLinkMatchAccuracyFormula = getAccuracyFormula(worksheetName, "First Link Match");
      const linkInTop2AccuracyFormula = getAccuracyFormula(worksheetName, "Link in Top 2");
      const avgGroundingScoreFormula = getAverageFormula(worksheetName, "Grounding Score");

      // I'll assume num_success and num_errors are available from the TaskRunner base class.
      // In TaskRunner, they are initialized to 0.
      // In processsAllRows, they are incremented.
      const newRowData = [
        [
          worksheetName,
          new Date().toLocaleString(),
          config.vertexAIProjectID,
          config.vertexAISearchAppId,
          run_results.numSuccessful,
          run_results.numFails,
          summaryMatchAccuracyFormula,
          firstLinkMatchAccuracyFormula,
          linkInTop2AccuracyFormula,
          avgGroundingScoreFormula,
          run_results.numCallsMade,
          run_results.timeTakenSeconds,
          run_results.stoppedReason,
        ],
      ];

      const tableBodyRange = runsTable.getDataBodyRange();
      tableBodyRange.load(["values", "rowCount"]);
      await context.sync();

      const tableValues = tableBodyRange.values;
      let rowIndex = -1;
      for (let i = 0; i < tableBodyRange.rowCount; i++) {
        if (tableValues[i][0] === worksheetName) {
          rowIndex = i;
          break;
        }
      }

      if (rowIndex > -1) {
        // Update existing row
        const rowRange = tableBodyRange.getRow(rowIndex);
        rowRange.values = newRowData;
        appendLog(`Updated row for ${worksheetName} in 'Search Eval Runs' table.`);
      } else {
        // Add new row
        runsTable.rows.add(null, newRowData);
        appendLog("Added a new row to 'Search Eval Runs' table.");
      }

      await context.sync();
    } catch (error) {
      appendError(`Error in addSearchRunToTable: ${error.message}`, error);
      showStatus(`Error adding row to Search Eval Runs table: ${JSON.stringify(error)}`, true);
    }
  }
}

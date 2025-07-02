import { appendError, appendLog } from "../ui.js";
import { calculateSimilarityUsingGemini, callVertexAIAnswer } from "../vertex_ai.js";
import { getAnswerConfigFromActiveSheet } from "./excel_common.js";
import { ExcelSearchRunner } from "./excel_search_runner.js";

export class ExcelAnswerRunner extends ExcelSearchRunner {
  async getSearchConfig() {
    return getAnswerConfigFromActiveSheet(true, true);
  }

  async getResultFromVertexAI(rowNum, config) {
    var query = this.queryColumn.values;
    return await callVertexAIAnswer(rowNum, query[rowNum][0], config);
  }

  async processRow(response_json, context, config, rowNum) {
    let numCalls = 0;
    if (response_json.hasOwnProperty("answer")) {
      // process the summary using throttling since it makes an external call
      const processSummaryPromise = this.throttled_process_summary(
        context,
        config,
        rowNum,
        response_json.answer.answerText,
        this.expectedSummaryColumn.values,
      ).then(async (callsSoFar) => {
        appendLog(`testCaseID: ${rowNum} Processed Search Summary.`);
      });

      this.searchTaskPromiseSet.add(processSummaryPromise);

      this.checkDocumentLinks(
        context,
        rowNum,
        response_json,
        this.expected_link_1_Column.values,
        this.expected_link_2_Column.values,
      );
      appendLog(`testCaseID: ${rowNum} Processed Doc Links.`);
      if (response_json.answer.hasOwnProperty("groundingScore")) {
        const groudingScorecell = this.checkGroundingScoreColumn.getRange().getCell(rowNum, 0);
        groudingScorecell.clear(Excel.ClearApplyTo.formats);
        groudingScorecell.values = [[response_json.answer.groundingScore.toString()]];
      }
    }
    
    // execute the tasks
    await Promise.allSettled(this.searchTaskPromiseSet.values());
    numCalls += 1;
    return numCalls;
  }


  checkDocumentLinks(context, rowNum, result, expectedLink1, expectedLink2) {
    var p0_result = null;
    var p2_result = null;
    const link_1_cell = this.link_1_Column.getRange().getCell(rowNum, 0);
    const link_2_cell = this.link_2_Column.getRange().getCell(rowNum, 0);
    const link_3_cell = this.link_3_Column.getRange().getCell(rowNum, 0);

    if (
      result.answer &&
      result.answer.citations &&
      result.answer.citations.length > 0 &&
      result.answer.citations[0].sources &&
      result.answer.citations[0].sources.length > 0
    ) {
      const sources = result.answer.citations[0].sources;
      if (sources[0] && sources[0].referenceId) {
        const referenceId = sources[0].referenceId;
        if (result.answer.references && result.answer.references[referenceId] && result.answer.references[referenceId].structuredDocumentInfo) {
          p0_result = result.answer.references[referenceId].structuredDocumentInfo.uri;
          link_1_cell.values = [[p0_result]];
        }
      }
      if (sources[1] && sources[1].referenceId) {
        const referenceId = sources[1].referenceId;
        if (result.answer.references && result.answer.references[referenceId] && result.answer.references[referenceId].structuredDocumentInfo) {
          p2_result = result.answer.references[referenceId].structuredDocumentInfo.uri;
          link_2_cell.values = [[p2_result]];
        }
      }
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
}

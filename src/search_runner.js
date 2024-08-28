import { callVertexAISearch } from "./vertex_ai.js";

import { appendError, appendLog, showStatus } from "./ui.js";

import {
  NotAuthenticatedError,
  PermissionDeniedError,
  QuotaError,
  ResourceNotFoundError,
  VelvetError,
} from "./common.js";

export class SearchRunner {
  constructor() {
    this.stopProcessing = false;
  }

  async processRow(
    response_json,
    context,
    config,
    rowNum,
    actualSummaryColumn,
    expectedSummary,
    summaryScoreColumn,
    link_1_Column,
    link_2_Column,
    link_3_Column,
    link_p0Column,
    link_top2Column,
    expectedLink1,
    expectedLink2
  ) {
    throw new Error("You have to implement the method processRow");
  }

  async processsAllRows(
    context,
    config,
    countRows,
    idColumn,
    queryColumn,
    expectedSummaryColumn,
    expected_link_1_Column,
    expected_link_2_Column,
    actualSummaryColumn,
    summaryScoreColumn,
    link_1_Column,
    link_2_Column,
    link_3_Column,
    link_p0Column,
    link_top2Column
  ) {
    let processedCount = 1;
    let id = idColumn.values;
    let query = queryColumn.values;
    let expectedSummary = expectedSummaryColumn.values;
    let expectedLink1 = expected_link_1_Column.values;
    let expectedLink2 = expected_link_2_Column.values;

    let numfails = 0;

    // map of promises
    const promiseMap = new Map();

    this.stopProcessing = false;
    // Loop through the test cases table ans run the tests
    while (
      processedCount <= countRows &&
      id[processedCount][0] !== null &&
      id[processedCount][0] !== ""
    ) {
      // Batch the calls to Vertex AI since there are throuput checks in place.\
      if (processedCount % config.batchSize === 0) {
        // delay calls with apropriate time
        await new Promise((r) =>
          setTimeout(r, config.timeBetweenCallsInSec * 1000)
        );
      }
      // Stop processing if there errors
      if (this.stopProcessing) {
        appendLog("Stopping execution.", null);
        break;
      }
      appendLog(`testCaseID: ${id[processedCount][0]} Started Processing...`);
      showStatus(`Processing testCaseID: ${id[processedCount][0]}`, false);
      // Call Vertex AI Search asynchronously and add the promise to promiseMap
      promiseMap.set(
        processedCount,
        callVertexAISearch(processedCount, query[processedCount][0], config)
          .then(async (result) => {
            let response_json = result.output;
            let testCaseRowNum = result.id;

            // Process Each Result
            await this.processRow(
              response_json,
              context,
              config,
              testCaseRowNum,
              actualSummaryColumn,
              expectedSummary,
              summaryScoreColumn,
              link_1_Column,
              link_2_Column,
              link_3_Column,
              link_p0Column,
              link_top2Column,
              expectedLink1,
              expectedLink2
            );
          })
          .catch((error) => {
            numfails++;
            this.stopProcessing = true;
            if (
              error instanceof NotAuthenticatedError ||
              error instanceof QuotaError ||
              error instanceof VelvetError ||
              error instanceof PermissionDeniedError ||
              error instanceof ResourceNotFoundError
            ) {
              appendError(
                ` ${error.name} processing testCaseID: ${error.id} errorCode: ${error.statusCode}`,
                error
              );
            } else {
              appendError(`Unexpected Error stack: ${error.stack}`, error);
            }
          })
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
      stoppedReason += ` Empty ID encountered after ${
        processedCount - 1
      } test cases.`;
    }
    var summary = `Finished! Successful: ${
      processedCount - numfails - 1
    }. ${stoppedReason}`;
    appendLog(summary);

    showStatus(summary, numfails > 0);
  }
}

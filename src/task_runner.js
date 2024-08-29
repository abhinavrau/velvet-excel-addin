import { appendError, appendLog, showStatus } from "./ui.js";

import {
  NotAuthenticatedError,
  PermissionDeniedError,
  QuotaError,
  ResourceNotFoundError,
  VertexAIError,
} from "./common.js";

export class TaskRunner {
  constructor() {
    this.stopProcessing = false;
    this.cancelPressed = false;
  }

  async processRow(response_json, context, config, rowNum) {
    throw new Error("You have to implement the method processRow");
  }

  async getResultFromExternalAPI(rowNum, config) {
    throw new Error("You have to implement the method getResultFromExternalAPI");
  }

  async processsAllRows(context, config, countRows, idArray) {
    let currentRow = 1;
    let numfails = 0;

    // map of promises
    const promiseMap = new Map();

    this.stopProcessing = false;
    this.cancelPressed = false;
    // Loop through the test cases table ans run the tests
    while (
      currentRow <= countRows &&
      idArray[currentRow][0] !== null &&
      idArray[currentRow][0] !== ""
    ) {
      appendLog(`testCaseID: ${idArray[currentRow][0]} Started Processing...`);
      showStatus(`Processing testCaseID: ${idArray[currentRow][0]}`, false);
      // Call Vertex AI Search asynchronously and add the promise to promiseMap

      const searchPromise = this.getResultFromExternalAPI(currentRow, config)
        .then(async (result) => {
          let response_json = result.output;
          let testCaseRowNum = result.id;

          // Process Each Result
          await this.processRow(response_json, context, config, testCaseRowNum);
        })
        .catch((error) => {
          numfails++;
          this.stopProcessing = true;
          if (
            error instanceof NotAuthenticatedError ||
            error instanceof QuotaError ||
            error instanceof VertexAIError ||
            error instanceof PermissionDeniedError ||
            error instanceof ResourceNotFoundError
          ) {
            appendError(` ${error.name} processing testCaseID: ${error.id}`, error);
          } else {
            appendError(`Unexpected Error stacktrace: ${error.stack}`, error);
          }
        });

      promiseMap.set(currentRow, searchPromise);

      // Batch the calls to Vertex AI since there are throuput checks in place.\
      if (currentRow % config.batchSize === 0) {
        // wait for calls so far to finish to finish
        await Promise.allSettled(promiseMap.values());

        // sync the contents to the cells
        await context.sync();

        // delay calls with apropriate time
        await new Promise((r) => setTimeout(r, config.timeBetweenCallsInSec * 1000));
      }
      currentRow++;

      // Stop processing if there errors
      if (this.stopProcessing || this.cancelPressed) {
        appendLog("Stopping execution.", null);
        break;
      }
    } // end while

    // wait for all the calls to finish if there are any remaining
    await Promise.allSettled(promiseMap.values());
    var stoppedReason = "";

    if (numfails > 0) {
      stoppedReason = `Failed: ${numfails}. See logs for details. `;
    }
    if (
      currentRow <= countRows &&
      (idArray[currentRow][0] === null || idArray[currentRow][0] === "")
    ) {
      stoppedReason += ` No content in ID Column after ${currentRow - 1} test cases.`;
    }
    var summary = `Finished! Successful: ${currentRow - numfails - 1}. ${stoppedReason}`;
    if (this.cancelPressed) {
      summary += "\n Cancelled Execution.";
    }
    appendLog(summary);
    showStatus(summary, numfails > 0);
  }
}

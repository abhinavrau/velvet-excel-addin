import { appendError, appendLog, showStatus } from "./ui.js";

//import { default as Bottleneck } from "../lib/bottleneck-2.19.5/lib/index.js";

import { AbortError, default as pThrottle } from "../lib/p-throttle.js";
import {
  NotAuthenticatedError,
  PermissionDeniedError,
  QuotaError,
  ResourceNotFoundError,
  VertexAIError,
} from "./common.js";

export class TaskRunner {
  constructor() {
    this.cancelPressed = false;
    this.throttle = pThrottle({
      limit: 6,
      interval: 1000,
    });
    this.throttled_api_call = this.throttle((a, b) => this.getResultFromVertexAI(a, b));
  }

  // Main call to vertex ai to get the data for earch row
  async getResultFromVertexAI(rowNum, config) {
    throw new Error("You have to implement the method getResultFromVertexAI");
  }

  // Passes the result from getResultFromVertexAI to this function to popoluate the sheet
  // And also to call other external functions for eval purposes
  async processRow(response_json, context, config, rowNum, numCallsSoFar) {
    throw new Error("You have to implement the method processRow");
  }

  // Called to wait or any throttled external calls to finish
  async waitForTaskstoFinish() {
    throw new Error("You have to implement the method waitForTaskstoFinish");
  }

  // request to cancel all throttled external calls since user signalled or error occured
  async cancelAllTasks() {
    throw new Error("You have to implement the method cancelAllTasks");
  }
  async processsAllRows(context, config, countRows, idArray) {
    let currentRow = 1;
    let numFails = 0;
    let numSuccessful = 0;
    let numCallsMade = 0;
    this.cancelPressed = false;

    // Start timer
    const startTime = new Date();
    let promiseSet = new Set();
    // Loop through the test cases table and run the tests
    while (
      currentRow <= countRows &&
      idArray[currentRow][0] !== null &&
      idArray[currentRow][0] !== ""
    ) {
      // Do the first call without throttling since crednetials could be wrong and we don't want to
      // overwhelm the server with bad crednetials.

      // if all good then keep going.

      const apiPromise = this.throttled_api_call(currentRow, config)
        .then(async (result) => {
          let response_json = result.output;
          let testCaseRowNum = result.id;
          numSuccessful++;
          numCallsMade++;

          appendLog(`testCaseID: ${testCaseRowNum} Processing results`);
          showStatus(`${testCaseRowNum} Processing results`, false);

          // Throttle each row since they also make api calls
          let numCallsProcessRow = await this.processRow(
            response_json,
            context,
            config,
            testCaseRowNum,
          );

          // Add to number of calls
          numCallsMade += numCallsProcessRow;
        })
        .catch((error) => {
          appendLog("Stopping task execution since errors encountered.");
          if (
            error instanceof NotAuthenticatedError ||
            error instanceof QuotaError ||
            error instanceof VertexAIError ||
            error instanceof PermissionDeniedError ||
            error instanceof ResourceNotFoundError
          ) {
            numCallsMade++;
            numFails++;
            this.throttled_api_call.abort();
            appendError(` ${error.name} processing testCaseID: ${error.id}`, error);
          } else if (error instanceof AbortError) {
            appendLog("Tasks Stopped processing.");
          } else {
            numFails++;
            this.throttled_api_call.abort();
            appendError(`Unexpected Error stacktrace: ${error.stack}`, error);
          }
        });

      // We resolve the first row since we don't want to overwhelm the server with
      // bad requests if some fo the config or auth tken is bad
      if (currentRow === 1) {
        await Promise.resolve(apiPromise);
        if (numFails > 0) {
          break;
        }
      }
      if (this.cancelPressed) {
        this.cancelAllTasks();
        break;
      }
      // first row is good to keep going.
      // Add the task promise to set
      promiseSet.add(apiPromise);

      // delay the loop so we can have the ability to cancel.\
      if (currentRow % config.batchSize === 0) {
        // delay calls with apropriate time
        await new Promise((r) => setTimeout(r, 500));
      }

      currentRow++;
    } // end while

    // wait for all tasks to resolve
    await Promise.allSettled(promiseSet);

    // wait for other tasks of inherited classes to finish
    await this.waitForTaskstoFinish();

    await context.sync();

    // Calculate time taken
    const endTime = new Date();
    const timeTaken = endTime - startTime; // This gives the time difference in milliseconds
    const timeTakenSeconds = `${(timeTaken / 1000).toFixed(2)} seconds`;

    var stoppedReason = "";

    if (numFails > 0) {
      stoppedReason = `Failed: ${numFails}. See logs for details.`;
    }
    if (
      currentRow <= countRows &&
      (idArray[currentRow][0] === null || idArray[currentRow][0] === "")
    ) {
      stoppedReason += ` No content in ID Column after ${currentRow - 1} test cases.`;
    }
    var summary = `Finished! Successful: ${numSuccessful}. ${stoppedReason}`;
    if (this.cancelPressed) {
      summary += " Cancelled Execution.";
    }

    appendLog(summary);
    appendLog(`Num Calls to Vertex AI:${numCallsMade} Time taken: ${timeTakenSeconds}`);
    showStatus(summary, numFails > 0);
  }

  async cancelProcessing() {
    try {
      this.cancelPressed = true;
      await this.throttled_api_call.abort();
      await this.cancelAllTasks();
      appendLog("Cancel requested...");
      await this.waitForTaskstoFinish();
    } catch (error) {
      appendError("Error Cancelling", error);
    }
  }
}

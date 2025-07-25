import { appendError, appendLog, showStatus } from "./ui.js";

import { AbortError, default as pThrottle } from "../lib/p-throttle.js";
import {
  NotAuthenticatedError,
  PermissionDeniedError,
  QuotaError,
  ResourceNotFoundError,
  VertexAIError,
} from "./common.js";

// Base class implements the logic of iterating through all the rows in the main data table
// and calls inherited methods for all the logic
export class TaskRunner {
  constructor() {
    this.cancelPressed = false;
    this.throttle = pThrottle({
      limit: 10,
      interval: 1000,
    });
    this.throttled_api_call = this.throttle((a, b) => this.getResultFromVertexAI(a, b));
  }

  // Main call to vertex ai to get the data for earch row
  async getResultFromVertexAI(rowNum, config) {
    throw new Error("You have to implement the method getResultFromVertexAI");
  }

  // Passes the result from getResultFromVertexAI to this function to populate the sheet
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

    if (config.batchSize !== null && config.timeBetweenCallsInSec != null) {
      this.throttle = pThrottle({
        limit: config.batchSize > 10 ? 10 : config.batchSize,
        interval: config.timeBetweenCallsInSec > 5 ? 5 * 1000 : config.timeBetweenCallsInSec * 1000,
      });
      this.throttled_api_call = this.throttle((a, b) => this.getResultFromVertexAI(a, b));
    }
    // Start timer
    const startTime = new Date();
    let promiseSet = new Set();
    // Loop through the test cases table and run the tests. Stop if ID column is empty
    while (
      currentRow <= countRows &&
      idArray[currentRow][0] !== null &&
      idArray[currentRow][0] !== ""
    ) {
      
      // Wrap the main call to Vertex AI API with throttling so we don't 
      // run into out of quota responses
      const apiPromise = this.throttled_api_call(currentRow, config)
        .then(async (result) => {
          let response_json = result.output;
          let testCaseRowNum = result.id;
          numSuccessful++;
          numCallsMade++;

          appendLog(`testCaseID: ${testCaseRowNum} Processing results`);
          showStatus(`${testCaseRowNum} Processing results`, false);

          // Process each row now from inherited class. Returns number of addional 
          // calls it make to Vertex AI
          let numCallsProcessRow = await this.processRow(
            response_json,
            context,
            config,
            testCaseRowNum,
          );

          // Add to number of calls so we can report back
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
      // bad requests if some fo the config or auth token is bad
      if (currentRow === 1) {
        await Promise.resolve(apiPromise);
        await this.waitForTaskstoFinish();
        await context.sync();
        if (numFails > 0) {
          break;
        }
      }
      if (this.cancelPressed) {
        this.cancelAllTasks();
        break;
      }
      // first row is good to keep going.
      // Add the task promise set
      promiseSet.add(apiPromise);

      // delay the loop by a little so we can have the ability to cancel.
      if (currentRow % config.batchSize === 0) {
        // delay calls with apropriate time

        await new Promise((r) => setTimeout(r, 300));
      }

      currentRow++;
    } // end while

    // wait for all tasks to resolve
    await Promise.allSettled(promiseSet);

    // wait for other tasks of inherited class to resolve
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

    return {
      numSuccessful,
      numFails,
      numCallsMade,
      timeTakenSeconds,
      stoppedReason,
    };
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

import expect from "expect";
import fetchMock from "fetch-mock";
import fs from "fs";
import { default as $, default as JQuery } from "jquery";
import pkg from "office-addin-mock";
import sinon from "sinon";

import { createSummarizationData, getSummarizationConfig } from "../src/summarization_runner.js";
import {
  createSummarizationEvalConfigTable,
  createSummarizationEvalDataTable,
} from "../src/summarization_tables.js";
import { showStatus } from "../src/ui.js";
import { callGeminiMultitModal } from "../src/vertex_ai.js";

import { summarization_configValues, summarization_TableHeader } from "../src/common.js";

// mock the UI components
global.showStatus = showStatus;
global.$ = $;
global.JQuery = JQuery;

global.callGeminiMultitModal = callGeminiMultitModal;

const { OfficeMockObject } = pkg;
export var testCaseRows = summarization_TableHeader.concat([
  [
    "1", // ID
    "Great magicians are hard working and do a lot of shows. You need to practice to be good at magic. So practice a lot.", //context
    "Becoming a great magician requires dedication and hard work.  Practicing magic regularly is essential to develop skills and perform well in many shows.", //Summary
    3, // summarization_quality
    3, //groundedness
    3, //fulfillment
    3, // summarization_helpfulness
    3, // summarization_verbosity
  ],
]);

describe("When Summarization Eval is clicked ", () => {
  var showStatusSpy;
  var mockTestData;
  var $stub;
  beforeEach(() => {
    // stub out jQuery calls
    $stub = sinon.stub(globalThis, "$").returns({
      empty: sinon.stub(),
      append: sinon.stub(),
      val: sinon.stub(),
      tabulator: sinon.stub(),
    });
    // Spy on showStatus
    showStatusSpy = sinon.spy(globalThis, "showStatus");

    fetchMock.reset();

    mockTestData = {
      ClearApplyTo: {
        formats: {},
      },
      context: {
        workbook: {
          worksheets: {
            name: "WorksheetName",
            getActiveWorksheet: function () {
              return this;
            },
            range: {
              values: [[]],
              format: {
                font: {
                  bold: false,
                },
                fill: {
                  color: "blue",
                },
                size: 16,
                autofitColumns: function () {
                  return true;
                },
                autofitRows: function () {
                  return true;
                },
              },
              getUsedRange: function () {
                return this.format;
              },
              getCell: function (rowNum, colNum) {
                return this.format;
              },
            },
            getRange: function (str) {
              return this.range;
            },
            getUsedRange: function () {
              return this.range;
            },
            tables: {
              add: function (str, flag) {
                return this;
              },
              getHeaderRowRange: function () {
                return this.header_row_range;
              },
              resize: function (str) {},
              getItem: function (str) {
                // check is str ends with string "TestCasesTable"
                if (str.endsWith("TestCasesTable")) {
                  return this.testCaseTable;
                } else if (str.endsWith("ConfigTable")) {
                  return this.configTable;
                }
                return this;
              },
              rows: {
                count: 1,
                values: [[]],
                add: function (str, vals) {
                  this.values = vals;
                },
              },
              configTable: {
                columns: {
                  getItemOrNullObject: function (columnName) {
                    return {
                      values: [[]],
                      load: function () {
                        const columnIndex = summarization_configValues[0].indexOf(columnName);

                        // If the column name is not found, return null
                        if (columnIndex === -1) {
                          return false;
                        }
                        // Extract the values from the specified column
                        this.values = summarization_configValues.map((row) => [row[columnIndex]]);
                        return true;
                      },
                    };
                  },
                },
              },
              testCaseTable: {
                // Initiallize our data object that will get populated
                data: Array(testCaseRows.length)
                  .fill(null)
                  .map(() => Array(testCaseRows.length).fill(null)),
                columns: {
                  // return a column object that is used to popluate the values returned from VertexAI APIs.
                  getItemOrNullObject: function (columnName) {
                    let columnIndex = -1;
                    // BEGIN Column object
                    return {
                      values: [[]],
                      columnIndex: -1,
                      load: function () {
                        columnIndex = testCaseRows[0].indexOf(columnName);

                        // If the column name is not found, return null
                        if (columnIndex === -1) {
                          return false;
                        }

                        // Extract the values from the specified column
                        this.values = testCaseRows.map((row) => [row[columnIndex]]);
                        return true;
                      },
                      getRange: function () {
                        return {
                          getCell: function (row, col) {
                            // create a cell object and keep track of it in the data array
                            // Create a cell object
                            var cell = {
                              values: [[""]],
                              clear: function (arg) {},
                              format: {
                                font: {
                                  bold: false,
                                },
                                fill: {
                                  color: "blue",
                                },
                              },
                            };
                            // access the columnIndex variable here
                            // Assign the cell object to the correct position in the data array
                            mockTestData.context.workbook.worksheets.tables.testCaseTable.data[row][
                              columnIndex
                            ] = cell;
                            return cell;
                          },
                        };
                      },
                    }; // END Coumn Object
                  },
                },
                rows: {
                  count: 1,
                },
              },

              header_row_range: {
                values: [[]],
              },
            },
          },
        },
      },
      // Mock the Excel.run method.
      run: async function (callback) {
        await callback(this.context);
      },
    };
  });

  afterEach(() => {
    $stub.restore();
    showStatusSpy.restore();
    sinon.reset();
  });

  it("should populate the Data Table with the correct values", async () => {
    // Create the final mock object from the seed object.
    const contextMock = new OfficeMockObject(mockTestData);

    global.Excel = contextMock;
    // Spy on the Showstatus function

    // Simulate creating the Config table
    await createSummarizationEvalConfigTable();
    // Fail the test ifshow status is called
    expect(showStatusSpy.notCalled).toBe(true);

    // Simulate creating the Config table
    await createSummarizationEvalDataTable();
    // Fail the test ifshow status is called
    expect(showStatusSpy.notCalled).toBe(true);

    // Get the config parameters from the config table
    const config = await getSummarizationConfig();

    const { request_json: summary_request_json, response_json: summary_response_json } =
      getRequestResponseJsonFromFile(
        "./test/data/summarization/test_summarization_request.json",
        "./test/data/summarization/test_summarization_response.json",
      );

    const url = `https://${config.vertexAILocation}-aiplatform.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/${config.vertexAILocation}/publishers/google/models/${config.model}:generateContent`;
    fetchMock.post(url, {
      status: 200,
      headers: { "Content-Type": `application/json` },
      body: JSON.stringify(summary_response_json),
    });

    // read the  summarization_quality request from json file
    const { request_json: quality_request_json, response_json: quality_response_json } =
      getRequestResponseJsonFromFile(
        "./test/data/summarization/test_summarization_quality_request.json",
        "./test/data/summarization/test_summarization_quality_response.json",
      );

    // read the  summarization_helpfulness request from json file
    const { request_json: helpfulness_request_json, response_json: helpfulness_response_json } =
      getRequestResponseJsonFromFile(
        "./test/data/summarization/test_summarization_helpfulness_request.json",
        "./test/data/summarization/test_summarization_helpfulness_response.json",
      );

    // read the  summarization_verbosity request from json file
    const { request_json: verbosity_request_json, response_json: verbosity_response_json } =
      getRequestResponseJsonFromFile(
        "./test/data/summarization/test_summarization_verbosity_request.json",
        "./test/data/summarization/test_summarization_verbosity_response.json",
      );

    // read the  summarization_quality request from json file
    const { request_json: grounding_request_json, response_json: grounding_response_json } =
      getRequestResponseJsonFromFile(
        "./test/data/summarization/test_summarization_groundedness_request.json",
        "./test/data/summarization/test_summarization_groundedness_response.json",
      );

    // read the  summarization_quality request from json file
    const { request_json: fulfillment_request_json, response_json: fulfillment_response_json } =
      getRequestResponseJsonFromFile(
        "./test/data/summarization/test_summarization_fulfillment_request.json",
        "./test/data/summarization/test_summarization_fulfillment_response.json",
      );

    const eval_url = `https://${config.vertexAILocation}-aiplatform.googleapis.com/v1beta1/projects/${config.vertexAIProjectID}/locations/${config.vertexAILocation}:evaluateInstances`;
    fetchMock.config.overwriteRoutes = false;

    fetchMock.postOnce(eval_url, {
      status: 200,
      headers: { "Content-Type": `application/json` },
      body: JSON.stringify(quality_response_json),
    });
    fetchMock.postOnce(eval_url, {
      status: 200,
      headers: { "Content-Type": `application/json` },
      body: JSON.stringify(helpfulness_response_json),
    });
    fetchMock.postOnce(eval_url, {
      status: 200,
      headers: { "Content-Type": `application/json` },
      body: JSON.stringify(verbosity_response_json),
    });
    fetchMock.postOnce(eval_url, {
      status: 200,
      headers: { "Content-Type": `application/json` },
      body: JSON.stringify(grounding_response_json),
    });
    fetchMock.postOnce(eval_url, {
      status: 200,
      headers: { "Content-Type": `application/json` },
      body: JSON.stringify(fulfillment_response_json),
    });

    // Execute the tests
    await createSummarizationData(config);

    // Verify mocks are called
    expect(fetchMock.called()).toBe(true);

    //  check if vertex ai search is called
    const callsToVertexAI = fetchMock.calls().filter((call) => call[0] === url);
    // Check if body is sent correctly to vertex ai search
    expect(callsToVertexAI[0][1] !== null).toBe(true);
    expect(JSON.parse(callsToVertexAI[0][1].body)).toEqual(summary_request_json);

    //  check if vertex ai search is called for eval
    const callsToVertexAI_EvalService = fetchMock.calls().filter((call) => call[0] === eval_url);

    // ensure call calls are made
    expect(callsToVertexAI_EvalService.length).toBe(5);

    // Check if body is sent correctly to summarization_quality
    expect(callsToVertexAI_EvalService[0][1] !== null).toBe(true);
    expect(JSON.parse(callsToVertexAI_EvalService[0][1].body)).toEqual(quality_request_json);

    // Check if body is sent correctly to summarization_helpfulness
    expect(callsToVertexAI_EvalService[0][1] !== null).toBe(true);
    expect(JSON.parse(callsToVertexAI_EvalService[1][1].body)).toEqual(helpfulness_request_json);

    // Check if body is sent correctly to summarization_verbosity
    expect(callsToVertexAI_EvalService[0][1] !== null).toBe(true);
    expect(JSON.parse(callsToVertexAI_EvalService[2][1].body)).toEqual(verbosity_request_json);

    // Check if body is sent correctly to grouding score
    expect(callsToVertexAI_EvalService[0][1] !== null).toBe(true);
    expect(JSON.parse(callsToVertexAI_EvalService[3][1].body)).toEqual(grounding_request_json);

    // Check if body is sent correctly to fulfillment score
    expect(callsToVertexAI_EvalService[0][1] !== null).toBe(true);
    expect(JSON.parse(callsToVertexAI_EvalService[4][1].body)).toEqual(fulfillment_request_json);

    // Check Summary got populated
    const { cell: summary_cell, col_index: summary_col_index } = getCellAndColumnIndexByName(
      "Summary",
      mockTestData,
    );
    expect(summary_cell[0][0]).toEqual(testCaseRows[1][summary_col_index]);

    // Check if summarization_quality score got populated
    const { cell: summarization_quality_cell, col_index: summarization_quality_col_index } =
      getCellAndColumnIndexByName("summarization_quality", mockTestData);
    expect(summarization_quality_cell[0][0]).toEqual(
      testCaseRows[1][summarization_quality_col_index],
    );

    // Check if helpfulness score got populated
    const { cell: helpfulness_quality_cell, col_index: helpfulness_quality_col_index } =
      getCellAndColumnIndexByName("summarization_helpfulness", mockTestData);
    expect(helpfulness_quality_cell[0][0]).toEqual(testCaseRows[1][helpfulness_quality_col_index]);

    // Check if summarization_verbosity score got populated
    const { cell: verbosity_quality_cell, col_index: verbosity_quality_col_index } =
      getCellAndColumnIndexByName("summarization_verbosity", mockTestData);
    expect(verbosity_quality_cell[0][0]).toEqual(testCaseRows[1][verbosity_quality_col_index]);

    // Check if groundedness score got populated
    const { cell: groundedness_quality_cell, col_index: groundedness_quality_col_index } =
      getCellAndColumnIndexByName("groundedness", mockTestData);
    expect(groundedness_quality_cell[0][0]).toEqual(
      testCaseRows[1][groundedness_quality_col_index],
    );

    // Check if fulfillment score got populated
    const { cell: fulfillment_quality_cell, col_index: fulfillment_quality_col_index } =
      getCellAndColumnIndexByName("fulfillment", mockTestData);
    expect(fulfillment_quality_cell[0][0]).toEqual(testCaseRows[1][fulfillment_quality_col_index]);
  });
});

function getRequestResponseJsonFromFile(requestJsonFilePath, responseJsonFilePath) {
  const request = fs.readFileSync(requestJsonFilePath);
  const response = fs.readFileSync(responseJsonFilePath);
  return {
    response_json: JSON.parse(response),
    request_json: JSON.parse(request),
  };
}

function getCellAndColumnIndexByName(column_name, mockTestData) {
  var col_index = testCaseRows[0].indexOf(column_name);

  var cell =
    mockTestData.context.workbook.worksheets.tables.testCaseTable.data[1][col_index] !== null
      ? mockTestData.context.workbook.worksheets.tables.testCaseTable.data[1][col_index].values
      : null;
  return { cell, col_index };
}

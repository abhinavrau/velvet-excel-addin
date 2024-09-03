import expect from "expect";
import fetchMock from "fetch-mock";
import { default as $, default as JQuery } from "jquery";
import pkg from "office-addin-mock";
import sinon from "sinon";

import { summarization_configValues, summarization_TableHeader } from "../src/common.js";
import { SummarizationRunner } from "../src/excel/excel_summarization_runner.js";
import {
  createSummarizationEvalConfigTable,
  createSummarizationEvalDataTable,
} from "../src/excel/excel_summarization_tables.js";
import { callGeminiMultitModal } from "../src/vertex_ai.js";
import { mockGeminiRequestResponse } from "./test_common.js";
// mock the UI components

global.$ = $;
global.JQuery = JQuery;

global.callGeminiMultitModal = callGeminiMultitModal;

const { OfficeMockObject } = pkg;
export var testCaseRows = summarization_TableHeader.concat([
  [
    "1", // ID
    "Great magicians are hard working and do a lot of shows. You need to practice to be good at magic. So practice a lot.", //context
    "Becoming a great magician requires dedication and hard work.  Practicing magic regularly is essential to develop skills and perform well in many shows.", //Summary
    "5-Very Good", // summarization_quality
  ],
  [
    "2", // ID
    "Great magicians are hard working and do a lot of shows. You need to practice to be good at magic. So practice a lot.", //context
    "Becoming a great magician requires dedication and hard work.  Practicing magic regularly is essential to develop skills and perform well in many shows.", //Summary
    "3-OK", // summarization_quality
  ],
  [
    "3", // ID
    "Great magicians are hard working and do a lot of shows. You need to practice to be good at magic. So practice a lot.", //context
    "Becoming a great magician requires dedication and hard work.  Practicing magic regularly is essential to develop skills and perform well in many shows.", //Summary
    "3-OK", // summarization_quality
  ],
  [
    "4", // ID
    "Great magicians are hard working and do a lot of shows. You need to practice to be good at magic. So practice a lot.", //context
    "Becoming a great magician requires dedication and hard work.  Practicing magic regularly is essential to develop skills and perform well in many shows.", //Summary
    "3-OK", // summarization_quality
  ],
]);

describe("When Summarization Eval is clicked ", () => {
  var mockTestData;
  var $stub;
  beforeEach(() => {
    // stub out jQuery calls
    $stub = sinon.stub(globalThis, "$").returns({
      empty: sinon.stub(),
      append: sinon.stub(),
      val: sinon.stub(),
      tabulator: sinon.stub(),
      prop: sinon.stub().returns(true),
    });

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
    sinon.reset();
  });

  it("should populate the Data Table with the correct values", async () => {
    // Create the final mock object from the seed object.
    const contextMock = new OfficeMockObject(mockTestData);

    global.Excel = contextMock;

    // Simulate creating the Config table
    await createSummarizationEvalConfigTable();

    // Simulate creating the Config table
    await createSummarizationEvalDataTable();

    var summarizationRunner = new SummarizationRunner();

    // Get the config parameters from the config table
    const config = await summarizationRunner.getSummarizationConfig();

    // Mock call for generating summary
    const {
      requestJson: summary_request_json,
      url: url,
      expectedResponse: responseJson,
    } = mockGeminiRequestResponse(
      1,
      200,
      "./test/data/summarization/test_summarization_request.json",
      "./test/data/summarization/test_summarization_response.json",
      config.model,
      config,
    );

    // Mock call for gnerating SummarizationEval
    const {
      requestJson: quality_request_json,
      url: quality_url,
      expectedResponse: quality_expectedResponse,
    } = mockGeminiRequestResponse(
      1,
      200,
      "./test/data/summarization/test_summarization_quality_request.json",
      "./test/data/summarization/test_summarization_quality_response.json",
      config.summarizationQualityModel,
      config,
    );

    config.batchSize = 10; // set batchSize high so test doesn't timeout
    // Execute the tests
    await summarizationRunner.createSummarizationData(config);

    // Verify mocks are called
    expect(fetchMock.called()).toBe(true);

    //  check if gemini is called
    const callsToVertexAI = fetchMock.calls();
    // Check if body is sent correctly to gemini
    expect(callsToVertexAI[0][1] !== null).toBe(true);
    expect(JSON.parse(callsToVertexAI[0][1].body)).toEqual(summary_request_json);

    // Check Summary got populated
    const { cell: summary_cell, col_index: summary_col_index } = getCellAndColumnIndexByName(
      "Summary",
      mockTestData,
    );
    expect(summary_cell[0][0]).toEqual(testCaseRows[1][summary_col_index]);

    // TODO Need to create the right request json for this request
    /* expect(callsToVertexAI[1][1] !== null).toBe(true);
    expect(JSON.parse(callsToVertexAI[1][1].body)).toEqual(quality_request_json);
 */
    // Check if summarization_quality score got populated
    const { cell: summarization_quality_cell, col_index: summarization_quality_col_index } =
      getCellAndColumnIndexByName("summarization_quality", mockTestData);
    expect(summarization_quality_cell[0][0]).toEqual(
      testCaseRows[1][summarization_quality_col_index],
    );
  });

  function getCellAndColumnIndexByName(column_name, mockTestData) {
    var col_index = testCaseRows[0].indexOf(column_name);

    var cell =
      mockTestData.context.workbook.worksheets.tables.testCaseTable.data[1][col_index] !== null
        ? mockTestData.context.workbook.worksheets.tables.testCaseTable.data[1][col_index].values
        : null;
    return { cell, col_index };
  }
});

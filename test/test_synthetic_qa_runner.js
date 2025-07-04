import expect from "expect";
import fetchMock from "fetch-mock";
import { default as $, default as JQuery } from "jquery";
import pkg from "office-addin-mock";
import sinon from "sinon";

import { synth_q_and_a_configValues, synth_q_and_a_TableHeader } from "../src/common.js";
import { SyntheticQARunner } from "../src/excel/excel_synthetic_qa_runner.js";
import {
  createSyntheticQAConfigTable,
  createSyntheticQADataTable,
} from "../src/excel/excel_synthetic_qa_tables.js";
import { showStatus } from "../src/ui.js";
import { callGeminiMultitModal } from "../src/vertex_ai.js";
import { mockGeminiRequestResponse } from "./test_common.js";

// mock the UI components
global.showStatus = showStatus;
global.$ = $;
global.JQuery = JQuery;

global.callGeminiMultitModal = callGeminiMultitModal;

const { OfficeMockObject } = pkg;
export var testCaseRows = synth_q_and_a_TableHeader.concat([
  [
    "1", // ID
    "gs://argolis-arau-gemini-bank/Procedure - Savings Account Opening.pdf", //GCS File URI
    "If I close my new savings account within 30 days of opening it, will I be charged a fee?", // Generated Question
    "Yes, you will be charged a $25 fee unless the closure is due to a Gemini Bank error in account opening, customer dissatisfaction with a product or service disclosed during the opening process, or insufficient funds.", //Expected Answer
    "5-Very Good",
  ],
]);

describe("When Generate Synthetic Q&A is clicked ", () => {
  var showStatusSpy;
  var mockTestData;
  var $stub;
  var appendStub = sinon.stub();
  var addClassStub = sinon.stub();
  var removeClassStub = sinon.stub();
  var emptyStub = sinon.stub();
  var jQueryObject;

  beforeEach(() => {
    // stub out jQuery calls
    jQueryObject = {
      empty: emptyStub,
      append: appendStub,
      val: sinon.stub(),
      tabulator: sinon.stub(),
      prop: sinon.stub().returns(true),
      removeClass: removeClassStub,
      addClass: addClassStub,
    };
    appendStub.returns(jQueryObject);
    $stub = sinon.stub(globalThis, "$").returns(jQueryObject);
    // Spy on showStatus
    showStatusSpy = sinon.spy(globalThis, "showStatus");

    fetchMock.reset();

    mockTestData = {
      ClearApplyTo: {
        formats: {},
      },
      GroupOption: {
        byRows: 1,
      },
      context: {
        workbook: {
          worksheets: {
            name: "WorksheetName",
            getActiveWorksheet: function () {
              return this;
            },
            getItem: function () {
              return this;
            },
            getItemOrNullObject(str) {
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
                return this;
              },
              getCell: function (rowNum, colNum) {
                return this.values;
              },
              clear: function () {
                return true;
              },
              getRange: function (str) {
                return this;
              },
              group: function (GroupOption) {
                return true;
              },
            },
            getRange: function (str) {
              return this.range;
            },
            getUsedRange: function () {
              return this.range;
            },
            clear: function () {
              return true;
            },
            tables: {
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
                  return this;
                },
                getCell: function (rowNum, colNum) {
                  return this.values[rowNum][colNum];
                },
                clear: function () {
                  return true;
                },
                getRange: function (str) {
                  return this;
                },
              },
              getRange: function (str) {
                return this.range;
              },
              add: function (str, flag) {
                return this;
              },
              getHeaderRowRange: function () {
                return this.header_row_range;
              },
              resize: function (str) {},
              getItem: function (str) {
                // check is str ends with string "TestCasesTable"
                if (str.endsWith("SyntheticQATable")) {
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
                        const columnIndex = synth_q_and_a_configValues[0].indexOf(columnName);

                        // If the column name is not found, return null
                        if (columnIndex === -1) {
                          return false;
                        }
                        // Extract the values from the specified column
                        this.values = synth_q_and_a_configValues.map((row) => [row[columnIndex]]);
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
    appendStub.reset();
    addClassStub.reset();
    removeClassStub.reset();
    emptyStub.reset();
    sinon.reset();
  });

  it("should populate the Test Data Table with the correct values", async () => {
    // Create the final mock object from the seed object.
    const contextMock = new OfficeMockObject(mockTestData);

    global.Excel = contextMock;
    // Spy on the Showstatus function


    const data = {
      sheetName: "WorksheetName",
      config: {
        model: "gemini-1.5-flash",
        vertexAIProjectId: "project_name",
        vertexAILocation: "us-central1",

      },
    };

    const worksheetName = data.sheetName;
    // Simulate creating the Config table
    await createSyntheticQAConfigTable(data);

    // Simulate creating the Config table
    await createSyntheticQADataTable(worksheetName);

    var syntheticQuestionAnswerRunner = new SyntheticQARunner();

    // Get the config parameters from the config table
    const config = await syntheticQuestionAnswerRunner.getSyntheticQAConfig();

    expect(config).not.toBe(null);
    // set up mock for file for first query
    // read the request from json file

    const {
      requestJson: requestJson,
      url: url,
      expectedResponse: responseJson,
    } = mockGeminiRequestResponse(
      1,
      200,
      "./test/data/multi_modal/test_multi_modal_request.json",
      "./test/data/multi_modal/test_multi_modal_response.json",
      config.model,
      config,
    );

    const {
      requestJson: quality_request_json,
      url: quality_url,
      expectedResponse: quality_expectedResponse,
    } = mockGeminiRequestResponse(
      1,
      200,
      "./test/data/question_answering/test_qa_quality_request.json",
      "./test/data/question_answering/test_qa_quality_response.json",
      config.qAQualityModel,
      config,
    );

    config.batchSize = 10; // set batchSize high so test doesn't timeout
    // Execute the tests
    await syntheticQuestionAnswerRunner.createSyntheticQAData(config);

    // Verify mocks are called
    expect(fetchMock.called()).toBe(true);

    //  check if vertex ai search is called
    const callsToVertexAI = fetchMock.calls().filter((call) => call[0] === url);
    // Check if body is sent correctly to vertex ai search
    expect(callsToVertexAI[0][1] !== null).toBe(true);
    expect(JSON.parse(callsToVertexAI[0][1].body)).toEqual(requestJson);

    // Check if  values get populated

    // Match  Generated Question
    const { cell: actual_generated_question, col_index: actual_generated_question_col_index } =
      getCellAndColumnIndexByName("Generated Question", mockTestData);
    expect(actual_generated_question[0][0]).toEqual(
      testCaseRows[1][actual_generated_question_col_index],
    );

    // Match  Expected Answer
    const { cell: actual_generated_answer, col_index: actual_generated_answer_col_index } =
      getCellAndColumnIndexByName("Expected Answer", mockTestData);
    expect(actual_generated_answer[0][0]).toEqual(
      testCaseRows[1][actual_generated_answer_col_index],
    );

    // Match  the Question Answer Quality
    const { cell: question_answer_quality, col_index: question_answer_quality_col_index } =
      getCellAndColumnIndexByName("Q & A Quality", mockTestData);
    expect(question_answer_quality[0][0]).toEqual(
      testCaseRows[1][question_answer_quality_col_index],
    );
  });
});

function getCellAndColumnIndexByName(column_name, mockTestData) {
  var col_index = testCaseRows[0].indexOf(column_name);

  var cell =
    mockTestData.context.workbook.worksheets.tables.testCaseTable.data[1][col_index] !== null
      ? mockTestData.context.workbook.worksheets.tables.testCaseTable.data[1][col_index].values
      : null;
  return { cell, col_index };
}

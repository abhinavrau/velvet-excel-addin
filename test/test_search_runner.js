import expect from "expect";
import fetchMock from "fetch-mock";
import { default as $, default as JQuery } from "jquery";
import pkg from "office-addin-mock";
import sinon from "sinon";

import { ExcelSearchRunner } from "../src/excel/excel_search_runner.js";
import { createVAIConfigTable, createVAIDataTable } from "../src/excel/excel_search_tables.js";
import { mockDiscoveryEngineRequestResponse } from "./test_common.js";

import { vertex_ai_search_configValues, vertex_ai_search_testTableHeader } from "../src/common.js";

// mock the UI components
global.$ = $;
global.JQuery = JQuery;

const { OfficeMockObject } = pkg;
export var testCaseRows = vertex_ai_search_testTableHeader.concat([
  [
    "1", // ID
    "What is Google's revenue for the year ending December 31, 2021", //Query
    "Revenue is $2.2 billion", //Expected Summary
    "Google's total revenue for the year ending December 31, 2021 was $257,637 million. This represents a 41% increase from the previous year. The majority of Google's revenue comes from its advertising business, which includes Google Search, YouTube ads, and Google Network. In 2021, Google's advertising revenue was $209,497 million. Google's other revenue streams include Google Cloud, which generated $19,206 million in revenue in 2021, and Other Bets, which generated $753 million.", // Actual Summary
    "TRUE", // Summary Match
    "TRUE", // First Link Match
    "TRUE", // Link in Top 2
    "0.59236944", // Grounding Score
    "link1", //Expected Link 1
    "link2", //Expected Link 2
    "link3", // Expected Link 3
    "link1", // Actual Link 1
    "link2", // Actual Link 2
    "link3", // Actual Link 3
  ],
  [
    "2", // ID
    "What is Google's revenue for the year ending December 31, 2021", //Query
    "Revenue is $2.2 billion", //Expected Summary
    "Google's total revenue for the year ending December 31, 2021 was $257,637 million. This represents a 41% increase from the previous year. The majority of Google's revenue comes from its advertising business, which includes Google Search, YouTube ads, and Google Network. In 2021, Google's advertising revenue was $209,497 million. Google's other revenue streams include Google Cloud, which generated $19,206 million in revenue in 2021, and Other Bets, which generated $753 million.", // Actual Summary
    "TRUE", // Summary Match
    "TRUE", // First Link Match
    "TRUE", // Link in Top 2
    "0.59236944", // Grounding Score
    "link1", //Expected Link 1
    "link2", //Expected Link 2
    "link3", // Expected Link 3
    "link1", // Actual Link 1
    "link2", // Actual Link 2
    "link3", // Actual Link 3
  ],
  [
    "2", // ID
    "What is Google's revenue for the year ending December 31, 2021", //Query
    "Revenue is $2.2 billion", //Expected Summary
    "Google's total revenue for the year ending December 31, 2021 was $257,637 million. This represents a 41% increase from the previous year. The majority of Google's revenue comes from its advertising business, which includes Google Search, YouTube ads, and Google Network. In 2021, Google's advertising revenue was $209,497 million. Google's other revenue streams include Google Cloud, which generated $19,206 million in revenue in 2021, and Other Bets, which generated $753 million.", // Actual Summary
    "TRUE", // Summary Match
    "TRUE", // First Link Match
    "TRUE", // Link in Top 2
    "0.59236944", // Grounding Score
    "link1", //Expected Link 1
    "link2", //Expected Link 2
    "link3", // Expected Link 3
    "link1", // Actual Link 1
    "link2", // Actual Link 2
    "link3", // Actual Link 3
  ],
  [
    "3", // ID
    "What is Google's revenue for the year ending December 31, 2021", //Query
    "Revenue is $2.2 billion", //Expected Summary
    "Google's total revenue for the year ending December 31, 2021 was $257,637 million. This represents a 41% increase from the previous year. The majority of Google's revenue comes from its advertising business, which includes Google Search, YouTube ads, and Google Network. In 2021, Google's advertising revenue was $209,497 million. Google's other revenue streams include Google Cloud, which generated $19,206 million in revenue in 2021, and Other Bets, which generated $753 million.", // Actual Summary
    "TRUE", // Summary Match
    "TRUE", // First Link Match
    "TRUE", // Link in Top 2
    "0.59236944", // Grounding Score
    "link1", //Expected Link 1
    "link2", //Expected Link 2
    "link3", // Expected Link 3
    "link1", // Actual Link 1
    "link2", // Actual Link 2
    "link3", // Actual Link 3
  ],
  [
    "4", // ID
    "What is Google's revenue for the year ending December 31, 2021", //Query
    "Revenue is $2.2 billion", //Expected Summary
    "Google's total revenue for the year ending December 31, 2021 was $257,637 million. This represents a 41% increase from the previous year. The majority of Google's revenue comes from its advertising business, which includes Google Search, YouTube ads, and Google Network. In 2021, Google's advertising revenue was $209,497 million. Google's other revenue streams include Google Cloud, which generated $19,206 million in revenue in 2021, and Other Bets, which generated $753 million.", // Actual Summary
    "TRUE", // Summary Match
    "TRUE", // First Link Match
    "TRUE", // Link in Top 2
    "0.59236944", // Grounding Score
    "link1", //Expected Link 1
    "link2", //Expected Link 2
    "link3", // Expected Link 3
    "link1", // Actual Link 1
    "link2", // Actual Link 2
    "link3", // Actual Link 3
  ],
  [
    "5", // ID
    "What is Google's revenue for the year ending December 31, 2021", //Query
    "Revenue is $2.2 billion", //Expected Summary
    "Google's total revenue for the year ending December 31, 2021 was $257,637 million. This represents a 41% increase from the previous year. The majority of Google's revenue comes from its advertising business, which includes Google Search, YouTube ads, and Google Network. In 2021, Google's advertising revenue was $209,497 million. Google's other revenue streams include Google Cloud, which generated $19,206 million in revenue in 2021, and Other Bets, which generated $753 million.", // Actual Summary
    "TRUE", // Summary Match
    "TRUE", // First Link Match
    "TRUE", // Link in Top 2
    "0.59236944", // Grounding Score
    "link1", //Expected Link 1
    "link2", //Expected Link 2
    "link3", // Expected Link 3
    "link1", // Actual Link 1
    "link2", // Actual Link 2
    "link3", // Actual Link 3
  ],
]);

describe("When Search Run Tests is clicked ", () => {
  var showStatusSpy;
  var appendErrorSpy;
  var mockTestData;
  var $stub;
  var tabulatorStub = sinon.stub();
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
      tabulator: tabulatorStub,
      prop: sinon.stub().returns(true),
      removeClass: removeClassStub,
      addClass: addClassStub,
    };
    appendStub.returns(jQueryObject);
    $stub = sinon.stub(globalThis, "$").returns(jQueryObject);

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
            getRange: function (str) {
              return this.range;
            },
            range: {
              values: [[]],
              getCell: function (x, y) {
                return this;
              },
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
                  return this.format;
                },
                getCell: function (rowNum, colNum) {
                  return this.format;
                },
                clear: function () {
                  return true;
                },
              },
              add: function (str, flag) {
                return this;
              },
              getRange: function (str) {
                return this.range;
              },
              getHeaderRowRange: function () {
                return this.header_row_range;
              },
              resize: function (str) {},
              getItemOrNullObject: function (str) {
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
                isNullObject: false,
                name: "WorksheetName.ConfigTable",
                columns: {
                  getItemOrNullObject: function (columnName) {
                    return {
                      values: [[]],
                      load: function () {
                        const columnIndex = vertex_ai_search_configValues[0].indexOf(columnName);

                        // If the column name is not found, return null
                        if (columnIndex === -1) {
                          return false;
                        }
                        // Extract the values from the specified column
                        this.values = vertex_ai_search_configValues.map((row) => [
                          row[columnIndex],
                        ]);
                        return true;
                      },
                    };
                  },
                },
              },
              testCaseTable: {
                isNullObject: false,
                name: "WorksheetName.TestCasesTable",
                // Initiallize our data object that will get populated
                data: Array(testCaseRows.length)
                  .fill(null)
                  .map(() => Array(13).fill(null)),
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
                  count: testCaseRows.length - 1,
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

    const ui_config = {
      vertexAISearchAppId: "l300-arau_1695783344117",
      vertexAIProjectID: "test_project",
      vertexAILocation: "us-central1",
      model: "gemini-2.0-flash-001/answer_gen/v1",
    };

    const data = {
      sheetName: "WorksheetName",
      config: ui_config,
    };
    const worksheetName = data.sheetName;

    // Simulate creating the Config table
    await createVAIConfigTable(data);

    // Simulate creating the Config table
    await createVAIDataTable(worksheetName);

    var excelSearchRunner = new ExcelSearchRunner();

    // Get the config parameters from the config table
    const config = await excelSearchRunner.getSearchConfig();

    expect(config).not.toBe(null);

    const url = `https://discoveryengine.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/global/collections/default_collection/engines/${config.vertexAISearchAppId}/servingConfigs/default_search:search`;
    // Prepare the request response mock the call to VertexAISearch
    const { requestJson, expectedResponse } = mockDiscoveryEngineRequestResponse(
      1,
      url,
      200,
      "./test/data/search/search_extractive_answer/test_vai_search_extractive_answer_request.json",
      "./test/data/search/search_extractive_answer/test_vai_search_extractive_answer_response.json",
      config,
    );

    // Mock the call for summary similarity
    const { url: summaryMatchUrl, response: summaryResponse } = mockSimilarityUsingVertexAI(
      config,
      "same",
    );

    const grouding_url = `https://discoveryengine.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/global/groundingConfigs/default_grounding_config:check`;
    // Prepare the request response mock the call to VertexAISearch
    const { requestJson: grouding_requestJson, expectedResponse: grouding_expectedResponse } =
      mockDiscoveryEngineRequestResponse(
        1,
        grouding_url,
        200,
        "./test/data/search/eval_check_grounding/2_test_check_grouding_request.json",
        "./test/data/search/eval_check_grounding/2_test_check_grouding_response.json",
        config,
      );

    // Execute the tests
    config.batchSize = 10; // set batchSize high so test doesn't timeout
    await excelSearchRunner.executeSearchTests(config);

    // Verify mocks are called
    expect(fetchMock.called()).toBe(true);

    //  check if vertex ai search is called
    const callsToVertexAISearch = fetchMock.calls().filter((call) => call[0] === url);
    // Check if body is sent correctly to vertex ai search
    expect(callsToVertexAISearch[0][1] !== null).toBe(true);
    expect(JSON.parse(callsToVertexAISearch[0][1].body)).toStrictEqual(requestJson);

    //  check if vertex ai is called for summary match
    const callsToSummaryMatch = fetchMock.calls().filter((call) => call[0] === summaryMatchUrl);
    // Check if body is sent correctly to vertex ai search
    expect(callsToSummaryMatch[0][1].body !== null).toBe(true);

    //  check if vertex ai is called for check grounding  match
    const callsToCheckGrounding = fetchMock.calls().filter((call) => call[0] === grouding_url);
    // Check if body is sent correctly to check groudning
    expect(JSON.parse(callsToCheckGrounding[0][1].body)).toEqual(grouding_requestJson);

    // Check if values get populated for each test case
    for (let i = 1; i < testCaseRows.length; i++) {
      // Match Actual Summary
      const { cell: actual_summary_cell, col_index: actual_summary_col_index } =
        getCellAndColumnIndexByName("Actual Summary", mockTestData, i);
      expect(actual_summary_cell[0][0]).toEqual(testCaseRows[i][actual_summary_col_index]);

      // Match Summary Match
      const { cell: summary_match_cell, col_index: summary_match_col_index } =
        getCellAndColumnIndexByName("Summary Match", mockTestData, i);
      expect(summary_match_cell[0][0]).toEqual(testCaseRows[i][summary_match_col_index]);

      // Match Grouding Score
      const { cell: grounding_score_cell, col_index: grounding_score_col_index } =
        getCellAndColumnIndexByName("Grounding Score", mockTestData, i);
      expect(grounding_score_cell[0][0]).toEqual(testCaseRows[i][grounding_score_col_index]);

      // Match first link match
      const { cell: first_link_match_cell, col_index: first_link_match_col_index } =
        getCellAndColumnIndexByName("First Link Match", mockTestData, i);

      expect(first_link_match_cell[0][0]).toEqual(testCaseRows[i][first_link_match_col_index]);

      // Match link in top 2
      const { cell: top2_link_match_cell, col_index: top2_link_match_col_index } =
        getCellAndColumnIndexByName("Link in Top 2", mockTestData, i);
      expect(top2_link_match_cell[0][0]).toEqual(testCaseRows[i][top2_link_match_col_index]);

      // Match links
      const { col_index: expected_link1_col_index } = getCellAndColumnIndexByName(
        "Expected Link 1",
        mockTestData,
        i,
      );
      const { cell: actual_link1_cell } = getCellAndColumnIndexByName(
        "Actual Link 1",
        mockTestData,
        i,
      );
      expect(actual_link1_cell[0][0]).toEqual(testCaseRows[i][expected_link1_col_index]);

      const { col_index: expected_link2_col_index } = getCellAndColumnIndexByName(
        "Expected Link 2",
        mockTestData,
        i,
      );
      const { cell: actual_link2_cell } = getCellAndColumnIndexByName(
        "Actual Link 2",
        mockTestData,
        i,
      );
      expect(actual_link2_cell[0][0]).toEqual(testCaseRows[i][expected_link2_col_index]);

      const { col_index: expected_link3_col_index } = getCellAndColumnIndexByName(
        "Expected Link 3",
        mockTestData,
        i,
      );
      const { cell: actual_link3_cell } = getCellAndColumnIndexByName(
        "Actual Link 3",
        mockTestData,
        i,
      );
      expect(actual_link3_cell[0][0]).toEqual(testCaseRows[i][expected_link3_col_index]);
    }
    expect($stub.calledWith("#log-pane")).toBe(true);
    expect($stub.calledWith(".status")).toBe(true);
    expect(removeClassStub.called).toBe(true);
    expect(addClassStub.calledWith("bg-green-100 text-green-700")).toBe(true);
    expect(appendStub.called).toBe(true);
  });

  it("should log error when there is error in vertex ai api", async () => {
    // Create the final mock object from the seed object.
    const contextMock = new OfficeMockObject(mockTestData);

    global.Excel = contextMock;

    const ui_config = {
      vertexAISearchAppId: "l300-arau_1695783344117",
      vertexAIProjectID: "test_project",
      vertexAILocation: "us-central1",
      model: "gemini-2.0-flash-001/answer_gen/v1",
    };
    const data = {
      sheetName: "WorksheetName",
      config: ui_config,
    };
    const worksheetName = data.sheetName;
    // Simulate creating the Config table
    await createVAIDataTable(worksheetName);

    // Simulate creating the Config table
    await createVAIConfigTable(data);

    var excelSearchRunner = new ExcelSearchRunner();

    // Get the config parameters from the config table
    const config = await excelSearchRunner.getSearchConfig();

    const url = `https://discoveryengine.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/global/collections/default_collection/engines/${config.vertexAISearchAppId}/servingConfigs/default_search:search`;

    // Prepare the request response mock the call to VertexAISearch
    const { requestJson, expectedResponse } = mockDiscoveryEngineRequestResponse(
      1,
      url,
      405,
      "./test/data/search/search_extractive_answer/test_vai_search_extractive_answer_request.json",
      "./test/data/not_authenticated.json",
      config,
    );

    // Mock the call for summary similarity
    const { url: summaryMatchUrl, response: summaryResponse } = mockSimilarityUsingVertexAI(
      config,
      "same",
    );

    // Execute the tests
    config.timeBetweenCallsInSec = 0; // set timeout to zero so test doesn't timeout
    await excelSearchRunner.executeSearchTests(config);
    // Verify mocks are called
    expect(fetchMock.called()).toBe(true);

    //  check if vertex ai search is called
    const callsToVertexAISearch = fetchMock.calls().filter((call) => call[0] === url);
    // Check if body is sent correctly to vertex ai search
    expect(callsToVertexAISearch[0][1] !== null).toBe(true);
    expect(JSON.parse(callsToVertexAISearch[0][1].body)).toStrictEqual(requestJson);

    // ensure things are written to the log pane
    expect($stub.calledWith("#log-pane")).toBe(true);

    // ensure things are written to the status div
    expect($stub.calledWith(".status")).toBe(true);
    expect(removeClassStub.called).toBe(true);
    expect(addClassStub.calledWith("bg-red-100 text-red-700")).toBe(true);
    expect(appendStub.called).toBe(true);
  });
});

function getCellAndColumnIndexByName(column_name, mockTestData, rowIndex) {
  var col_index = testCaseRows[0].indexOf(column_name);

  var cell =
    mockTestData.context.workbook.worksheets.tables.testCaseTable.data[rowIndex][col_index] !== null
      ? mockTestData.context.workbook.worksheets.tables.testCaseTable.data[rowIndex][col_index]
          .values
      : null;
  return { cell, col_index };
}

function mockSimilarityUsingVertexAI(config, returnVal) {
  var response = {
    candidates: [
      {
        content: {
          role: "model",
          parts: [
            {
              text: "same",
            },
          ],
        },
      },
    ],
    usageMetadata: {
      promptTokenCount: 634,
      candidatesTokenCount: 166,
      totalTokenCount: 800,
    },
  };
  const url = `https://${config.vertexAILocation}-aiplatform.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/${config.vertexAILocation}/publishers/google/models/gemini-2.0-flash-001:generateContent`;
  fetchMock.post(url, {
    status: 200,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(response),
  });

  return { url, response };
}

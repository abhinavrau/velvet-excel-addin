import expect from "expect";
import fetchMock from "fetch-mock";
import { default as $, default as JQuery } from "jquery";
import pkg from "office-addin-mock";
import sinon from "sinon";

import { ExcelAnswerRunner } from "../src/excel/excel_answer_runner.js";
import {
  createVAIAnswerConfigTable,
  createVAIDataTable,
} from "../src/excel/excel_search_tables.js";
import { mockDiscoveryEngineRequestResponse } from "./test_common.js";

import { vertex_ai_answer_configValues, vertex_ai_search_testTableHeader } from "../src/common.js";

// mock the UI components
global.$ = $;
global.JQuery = JQuery;

const { OfficeMockObject } = pkg;
export var testCaseRows = vertex_ai_search_testTableHeader.concat([
  [
    "1", // ID
    "What is the gross revenue currency neutral growth for Q1/24?", //Query
    "201.5%", //Expected Summary
    "The gross revenue currency neutral growth for Q1/24 actual is 201.5%.\n", // Actual Summary
    "TRUE", // Summary Match
    "TRUE", // First Link Match
    "TRUE", // Link in Top 2
    "0.553", // Grounding Score
    "https://drive.google.com/a/arau.altostrat.com/open?id=1d-qJ9bSyYsGnHy0XZtRaO7p0Sd-Y6TEt", //Expected Link 1
    "https://drive.google.com/a/arau.altostrat.com/open?id=1d-qJ9bSyYsGnHy0XZtRaO7p0Sd-Y6TEt", //Expected Link 2
    "", // Expected Link 3
    "link1", // Actual Link 1 - Dummy value. Actual link is fetched from test data file
    "link2", // Actual Link 2- Dummy value. Actual link is fetched from test data file
    "link3", // Actual Link 3 - Dummy value. Actual link is fetched from test data file
  ],
  [
    "2", // ID
    "What is the gross revenue currency neutral growth for Q1/24?", //Query
    "201.5%", //Expected Summary
    "The gross revenue currency neutral growth for Q1/24 actual is 201.5%.\n", // Actual Summary
    "TRUE", // Summary Match
    "TRUE", // First Link Match
    "TRUE", // Link in Top 2
    "0.553", // Grounding Score
    "https://drive.google.com/a/arau.altostrat.com/open?id=1d-qJ9bSyYsGnHy0XZtRaO7p0Sd-Y6TEt", //Expected Link 1
    "https://drive.google.com/a/arau.altostrat.com/open?id=1d-qJ9bSyYsGnHy0XZtRaO7p0Sd-Y6TEt", //Expected Link 2
    "", // Expected Link 3
    "link1", // Actual Link 1 - Dummy value. Actual link is fetched from test data file
    "link2", // Actual Link 2- Dummy value. Actual link is fetched from test data file
    "link3", // Actual Link 3 - Dummy value. Actual link is fetched from test data file
  ],
  [
    "3", // ID
    "What is the gross revenue currency neutral growth for Q1/24?", //Query
    "201.5%", //Expected Summary
    "The gross revenue currency neutral growth for Q1/24 actual is 201.5%.\n", // Actual Summary
    "TRUE", // Summary Match
    "TRUE", // First Link Match
    "TRUE", // Link in Top 2
    "0.553", // Grounding Score
    "https://drive.google.com/a/arau.altostrat.com/open?id=1d-qJ9bSyYsGnHy0XZtRaO7p0Sd-Y6TEt", //Expected Link 1
    "https://drive.google.com/a/arau.altostrat.com/open?id=1d-qJ9bSyYsGnHy0XZtRaO7p0Sd-Y6TEt", //Expected Link 2
    "", // Expected Link 3
    "link1", // Actual Link 1 - Dummy value. Actual link is fetched from test data file
    "link2", // Actual Link 2- Dummy value. Actual link is fetched from test data file
    "link3", // Actual Link 3 - Dummy value. Actual link is fetched from test data file
  ],
  [
    "4", // ID
    "What is the gross revenue currency neutral growth for Q1/24?", //Query
    "201.5%", //Expected Summary
    "The gross revenue currency neutral growth for Q1/24 actual is 201.5%.\n", // Actual Summary
    "TRUE", // Summary Match
    "TRUE", // First Link Match
    "TRUE", // Link in Top 2
    "0.553", // Grounding Score
    "https://drive.google.com/a/arau.altostrat.com/open?id=1d-qJ9bSyYsGnHy0XZtRaO7p0Sd-Y6TEt", //Expected Link 1
    "https://drive.google.com/a/arau.altostrat.com/open?id=1d-qJ9bSyYsGnHy0XZtRaO7p0Sd-Y6TEt", //Expected Link 2
    "", // Expected Link 3
    "link1", // Actual Link 1 - Dummy value. Actual link is fetched from test data file
    "link2", // Actual Link 2- Dummy value. Actual link is fetched from test data file
    "link3", // Actual Link 3 - Dummy value. Actual link is fetched from test data file
  ],
  [
    "5", // ID
    "What is the gross revenue currency neutral growth for Q1/24?", //Query
    "201.5%", //Expected Summary
    "The gross revenue currency neutral growth for Q1/24 actual is 201.5%.\n", // Actual Summary
    "TRUE", // Summary Match
    "TRUE", // First Link Match
    "TRUE", // Link in Top 2
    "0.553", // Grounding Score
    "https://drive.google.com/a/arau.altostrat.com/open?id=1d-qJ9bSyYsGnHy0XZtRaO7p0Sd-Y6TEt", //Expected Link 1
    "https://drive.google.com/a/arau.altostrat.com/open?id=1d-qJ9bSyYsGnHy0XZtRaO7p0Sd-Y6TEt", //Expected Link 2
    "", // Expected Link 3
    "link1", // Actual Link 1 - Dummy value. Actual link is fetched from test data file
    "link2", // Actual Link 2- Dummy value. Actual link is fetched from test data file
    "link3", // Actual Link 3 - Dummy value. Actual link is fetched from test data file
  ], 
]);

describe("When Answer Run Tests is clicked ", () => {
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
                        const columnIndex = vertex_ai_answer_configValues[0].indexOf(columnName);

                        // If the column name is not found, return null
                        if (columnIndex === -1) {
                          return false;
                        }
                        // Extract the values from the specified column
                        this.values = vertex_ai_answer_configValues.map((row) => [
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
      preamble:
        "You are an expert financial analyst. Focus on questions related to financial amounts, dates, and deadlines.",
      model: "stable",
      ignoreAdversarialQuery: true,
      ignoreNonAnswerSeekingQuery: false,
      ignoreLowRelevantContent: true,
      includeCitations: true,
      includeGroundingSupports: true,
    };

    const data = {
      sheetName: "WorksheetName",
      config: ui_config,
    };
    const worksheetName = data.sheetName;

    // Simulate creating the Config table
    await createVAIAnswerConfigTable(data);

    // Simulate creating the Config table
    await createVAIDataTable(worksheetName);

    var excelAnswerRunner = new ExcelAnswerRunner();

    // Get the config parameters from the config table
    const config = await excelAnswerRunner.getSearchConfig();

    expect(config).not.toBe(null);

    const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${config.vertexAIProjectID}/locations/global/collections/default_collection/engines/${config.vertexAISearchAppId}/servingConfigs/default_search:answer`;
    // Prepare the request response mock the call to VertexAISearch
    const { requestJson, expectedResponse } = mockDiscoveryEngineRequestResponse(
      1,
      url,
      200,
      "./test/data/search/answer/test_vai_answer_request.json",
      "./test/data/search/answer/test_vai_answer_response.json",
      config,
    );

    // Mock the call for summary similarity
    const { url: summaryMatchUrl, response: summaryResponse } = mockSimilarityUsingVertexAI(
      config,
      "same",
    );
    // Execute the tests
    config.batchSize = 10; // set batchSize high so test doesn't timeout
    await excelAnswerRunner.executeSearchTests(config);

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

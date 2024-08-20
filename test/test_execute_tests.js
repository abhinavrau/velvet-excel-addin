
import expect from 'expect';
import fetchMock from 'fetch-mock';
import { default as $, default as JQuery } from 'jquery';
import pkg from 'office-addin-mock';
import sinon from 'sinon';

import { showStatus } from '../src/ui.js';
import { executeTests, getConfig } from '../src/velvet_runner.js';
import { createConfigTable, createDataTable } from '../src/velvet_tables.js';
import { calculateSimilarityUsingPalm2, callVertexAISearch } from '../src/vertex_ai.js';
import { mockVertexAISearchRequestResponse } from './test_common.js';

import { configValues, testCaseData } from '../src/common.js';

// mock the UI components
global.showStatus = showStatus;
global.$ = $;
global.JQuery = JQuery;


global.callVertexAISearch = callVertexAISearch;
global.calculateSimilarityUsingVertexAI = calculateSimilarityUsingPalm2;



const { OfficeMockObject } = pkg;
export var testCaseRows = testCaseData.concat([
    [
        "1", // ID
        "What is Google's revenue for the year ending December 31, 2021", //Query
        "Revenue is $2.2 billion",//Expected Summary
        "Google's revenue for the year ending December 31, 2022 was $2.5 billion. This is based on the deferred revenue as of December 31, 2021.", // Actual Summary
        "link1", //Expected Link 1
        "link2", //Expected Link 2
        "link3", // Expected Link 3
        "TRUE", // Summary Match
        "TRUE",// First Link Match
        "TRUE", // Link in Top 2
        "link1", // Actual Link 1
        "link2", // Actual Link 2
        "link3", // Actual Link 3
    ],
    [
        "2", // ID
        "What is Google's revenue for the year ending December 31, 2021", //Query
        "Revenue is $2.2 billion",//Expected Summary
        "Google's revenue for the year ending December 31, 2022 was $2.5 billion. This is based on the deferred revenue as of December 31, 2021.", // Actual Summary
        "link1", //Expected Link 1
        "link2", //Expected Link 2
        "link3", // Expected Link 3
        "TRUE", // Summary Match
        "TRUE",// First Link Match
        "TRUE", // Link in Top 2
        "link1", // Actual Link 1
        "link2", // Actual Link 2
        "link3", // Actual Link 3
    ]
]);

describe("When Execute Test is clicked ", () => {

    var mockTestData;
    var $stub;
    beforeEach(() => {
        // stub out jQuery calls
        $stub = sinon.stub(globalThis, '$').returns({
            empty: sinon.stub(),
            append: sinon.stub(),
            val: sinon.stub(),
            tabulator: sinon.stub(),
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
                                    bold: false
                                },
                                fill: {
                                    color: "blue"
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
                            }
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
                            resize: function (str) {

                            },
                            getItem: function (str) {
                                // check is str ends with string "TestCasesTable"
                                if (str.endsWith("TestCasesTable")) {
                                    return this.testCaseTable;
                                }
                                else if (str.endsWith("ConfigTable")) {
                                    return this.configTable;
                                }
                                return this;
                            },
                            rows: {
                                count: 1,
                                values: [[]],
                                add: function (str, vals) {
                                    this.values = vals;
                                }

                            },
                            configTable: {

                                columns: {

                                    getItemOrNullObject: function (columnName) {
                                        return {
                                            values: [[]],
                                            load: function () {
                                                const columnIndex = configValues[0].indexOf(columnName);

                                                // If the column name is not found, return null
                                                if (columnIndex === -1) {
                                                    return false;
                                                }
                                                // Extract the values from the specified column
                                                this.values = configValues.map(row => [row[columnIndex]]);
                                                return true;
                                            }
                                        };
                                    }

                                },
                            },
                            testCaseTable: {
                                // Initiallize our data object that will get populated 
                                data: Array(testCaseRows.length).fill(null).map(() => Array(13).fill(null)),
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
                                                this.values = testCaseRows.map(row => [row[columnIndex]]);
                                                return true;
                                            },
                                            getRange: function () {
                                                return {
                                                    getCell: function (row, col) {

                                                        // create a cell object and keep track of it in the data array 
                                                        // Create a cell object
                                                        var cell =  {
                                                            values: [[""]],
                                                            clear: function (arg) { },
                                                            format: {
                                                                font: {
                                                                    bold: false
                                                                },
                                                                fill: {
                                                                    color: "blue"
                                                                },
                                                            }
                                                        };
                                                        // access the columnIndex variable here
                                                        // Assign the cell object to the correct position in the data array
                                                        mockTestData.context.workbook.worksheets.tables.testCaseTable.data[row][columnIndex] = cell;
                                                        return cell;

                                                    }
                                                }
                                            }

                                        } // END Coumn Object
                                    },

                                },
                                rows: {
                                    count: 1,
                                }
                            },

                            header_row_range: {
                                values: [[]]
                            },

                        },

                    },

                }
            },
            // Mock the Excel.run method.
            run: async function (callback) {
                await callback(this.context);
            },
        }
    });

    afterEach(() => {
        $stub.restore();
    });


    it("should populate the Test Data Table with the correct values", async () => {


        // Create the final mock object from the seed object.
        const contextMock = new OfficeMockObject(mockTestData);


        global.Excel = contextMock;
        // Spy on the Showstatus function
        // Spy on showStatus
        const showStatusSpy = sinon.spy(globalThis, 'showStatus');

        // Simulate creating the Config table
        await createDataTable();
        // Fail the test ifshow status is called
        expect(showStatusSpy.notCalled).toBe(true);

        // Simulate creating the Config table
        await createConfigTable();
        // Fail the test ifshow status is called
        expect(showStatusSpy.notCalled).toBe(true);


        // Get the config parameters from the config table
        const config = await getConfig();

        // Prepare the request response mock the call to VertexAISearch
        const { requestJson, url, expectedResponse } = mockVertexAISearchRequestResponse(
            1,
            200,
            './test/data/extractive_answer/test_vai_search_extractive_answer_request.json',
            './test/data/extractive_answer/test_vai_search_extractive_answer_response.json',
            config);


        // Mock the call for summary similarity
        const { url: summaryMatchUrl, response: summaryResponse } = await mockSimilarityUsingVertexAI(config, 'same');
        // Execute the tests
        await executeTests(config);

        // Verify mocks are called
        expect(fetchMock.called()).toBe(true);

        //  check if vertex ai search is called
        const callsToVertexAISearch = fetchMock.calls().filter(call => call[0] === url);
        // Check if body is sent correctly to vertex ai search
        expect(callsToVertexAISearch[0][1] !== null).toBe(true);
        expect(JSON.parse(callsToVertexAISearch[0][1].body)).toStrictEqual(requestJson);

        //  check if vertex ai is called for summary match
        const callsToSummaryMatch = fetchMock.calls().filter(call => call[0] === summaryMatchUrl);
        // Check if body is sent correctly to vertex ai search
        expect(callsToSummaryMatch[0][1].body !== null).toBe(true);

        // Check if  values get populated

        // Match  Actual Summary 
        const { cell: actual_summary_cell, col_index: actual_summary_col_index } = getCellAndColumnIndexByName("Actual Summary", mockTestData);
        expect(actual_summary_cell[0][0]).toEqual(testCaseRows[1][actual_summary_col_index]);

        // Match  Summary Match 
        const { cell: summary_match_cell, col_index: summary_match_col_index } = getCellAndColumnIndexByName("Summary Match", mockTestData);
        expect(summary_match_cell[0][0]).toEqual(testCaseRows[1][summary_match_col_index]);

        // Match first link match
        const { cell: first_link_match_cell, col_index: first_link_match_col_index } = getCellAndColumnIndexByName("First Link Match", mockTestData);
        expect(first_link_match_cell[0][0]).toEqual(testCaseRows[1][first_link_match_col_index]);

        // Match link in top 2
        const { cell: top2_link_match_cell, col_index: top2_link_match_col_index } = getCellAndColumnIndexByName("Link in Top 2", mockTestData);
        expect(top2_link_match_cell[0][0]).toEqual(testCaseRows[1][top2_link_match_col_index]);


        // Match links
        const {  col_index: expected_link1_col_index } = getCellAndColumnIndexByName("Expected Link 1", mockTestData);
        const { cell: actual_link1_cell} = getCellAndColumnIndexByName("Actual Link 1", mockTestData);
        expect(actual_link1_cell[0][0]).toEqual(testCaseRows[1][expected_link1_col_index]);

        const { col_index: expected_link2_col_index } = getCellAndColumnIndexByName("Expected Link 2", mockTestData);
        const { cell: actual_link2_cell } = getCellAndColumnIndexByName("Actual Link 2", mockTestData);
        expect(actual_link2_cell[0][0]).toEqual(testCaseRows[1][expected_link2_col_index]);

        const { col_index: expected_link3_col_index } = getCellAndColumnIndexByName("Expected Link 3", mockTestData);
        const { cell: actual_link3_cell } = getCellAndColumnIndexByName("Actual Link 3", mockTestData);
        expect(actual_link3_cell[0][0]).toEqual(testCaseRows[1][expected_link3_col_index]);

      


    });




});
function getCellAndColumnIndexByName(column_name, mockTestData) {
    var col_index = testCaseRows[0].indexOf(column_name);

    var cell = mockTestData.context.workbook.worksheets.tables
        .testCaseTable.data[1][col_index] !== null ? mockTestData.context.workbook.worksheets.tables
            .testCaseTable.data[1][col_index].values : null;
    return { cell, col_index };
}

async function mockSimilarityUsingVertexAI(config, returnVal) {


    var response = {
        predictions: [
            {
                content: `${returnVal}`,
            },
        ],
    };
    const url = `https://${config.vertexAILocation}-aiplatform.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/${config.vertexAILocation}/publishers/google/models/text-bison:predict`;
    fetchMock.postOnce(url, {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(response)
    });

    return { url, response };
}

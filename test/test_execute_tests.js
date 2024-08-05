
import expect from 'expect';
import fetchMock from 'fetch-mock';
import { default as $, default as JQuery } from 'jquery';
import pkg from 'office-addin-mock';
import sinon from 'sinon';

import { showStatus } from '../src/ui.js';
import { getConfig, runTests } from '../src/velvet_runner.js';
import { createConfigTable, createDataTable } from '../src/velvet_tables.js';
import { calculateSimilarityUsingVertexAI, callVertexAISearch } from '../src/vertex_ai.js';
import { mockVertexAISearchRequestResponse } from './test_common.js';

import { configValues, testCaseData } from '../src/common.js';

global.showStatus = showStatus;

global.callVertexAISearch = callVertexAISearch;
global.calculateSimilarityUsingVertexAI = calculateSimilarityUsingVertexAI;

global.$ = $;
global.JQuery = JQuery;

const { OfficeMockObject } = pkg;

describe("When Execute Test is clicked ", () => {

    var mockTestData;

    beforeEach(() => {
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
                                data: Array(testCaseData.length).fill(null).map(() => Array(13).fill(null)),
                                columns: {
                                    // return a column object that is used to popluate the values returned from VertexAI APIs.
                                    getItemOrNullObject: function (columnName) {
                                        let columnIndex = -1;
                                        // BEGIN Column object
                                        return {
                                            values: [[]],
                                            columnIndex: -1,
                                            load: function () {
                                                columnIndex = testCaseData[0].indexOf(columnName);

                                                // If the column name is not found, return null
                                                if (columnIndex === -1) {
                                                    return false;
                                                }

                                                // Extract the values from the specified column
                                                this.values = testCaseData.map(row => [row[columnIndex]]);
                                                return true;
                                            },
                                            getRange: function () {
                                                return {
                                                    getCell: function (row, col) {

                                                        // create a cell object and keep track of it in the data array 
                                                        // Create a cell object
                                                        var cell = {
                                                            values: [[]],
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

    it("should populate the Test Data Table with the correct values", async () => {


        // Create the final mock object from the seed object.
        const contextMock = new OfficeMockObject(mockTestData);


        global.Excel = contextMock;
        // Spy on the Showstatus function
        // Spy on showStatus
        const showStatusSpy = sinon.spy(globalThis, 'showStatus');

        // Simulate creating the Config table
        await createConfigTable();
        // Fail the test ifshow status is called
        expect(showStatusSpy.notCalled).toBe(true);

        // Simulate creating the Data table
        await createDataTable();
        // Fail the test ifshow status is called
        expect(showStatusSpy.notCalled).toBe(true);


        // stub out jQuery calls
        const $stub = sinon.stub(globalThis, '$').returns({
            empty: sinon.stub(),
            append: sinon.stub(),
            val: sinon.stub(),
        });

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
        mockSimilarityUsingVertexAI(config, 'same');   
        // Populate the test table
        await runTests(config);
        
        // Verify InspectionFile
        expect(fetchMock.called()).toBe(true);
        // Assert request body is correct
        expect(JSON.parse(fetchMock.lastCall()[1].body)).toEqual(requestJson);
        // Assert URL is correct
        expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());

        // Check if return values get populated
        // get the column index that is Actual Summary title
        const columnIndex = testCaseData[0].indexOf("Actual Summary");

        const actual_summary_cell = mockTestData.context.workbook.worksheets.tables
            .testCaseTable.data[1][columnIndex].values;
        expect(actual_summary_cell[0][0]).toEqual("Google's revenue for the year ending December 31, 2022 was $2.5 billion. This is based on the deferred revenue as of December 31, 2021.");


        $stub.restore();


    });




});

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

    
}
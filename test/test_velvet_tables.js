
import expect from 'expect';
import fetchMock from 'fetch-mock';
import { default as $, default as JQuery } from 'jquery';
import pkg from 'office-addin-mock';
import sinon from 'sinon';

import { showStatus } from '../src/ui.js';
import { getConfig, runTests } from '../src/velvet_runner.js';
import { createConfigTable, createDataTable } from '../src/velvet_tables.js';
import { calculateSimilarityUsingVertexAI, callVertexAISearch } from '../src/vertex_ai.js';
import { mockVertexAISearchRequestResponse } from '../test/common.js';

global.showStatus = showStatus;

global.callVertexAISearch = callVertexAISearch;
global.calculateSimilarityUsingVertexAI = calculateSimilarityUsingVertexAI;

global.$ = $;
global.JQuery = JQuery;




const { OfficeMockObject } = pkg;

const configValues = [
    ["Config", "Value"],
    ["Vertex AI Search Project Number", "384473000457"],
    ["Vertex AI Search DataStore Name", "alphabet-pdfs_1695783402380"],
    ["Vertex AI Project ID", "argolis-arau"],
    ["Vertex AI Location", "us-central1"],
    ["maxExtractiveAnswerCount (1-5)", "2"], //maxExtractiveAnswerCount
    ["maxExtractiveSegmentCount (1-5)", "0"], //maxExtractiveSegmentCount
    ["maxSnippetCount (1-5)", "0"], //maxSnippetCount
    ["Preamble (Customized Summaries)", ""],
    ["Summarization Model", "gemini-1.0-pro-001/answer_gen/v1"],
    ["summaryResultCount (1-5)", "2"], //summaryResultCount
    ["useSemanticChunks (True or False)", "False"], //useSemanticChunks
    ["ignoreAdversarialQuery (True or False)", "True"], // ignoreAdversarialQuery
    ["ignoreNonSummarySeekingQuery (True or False)", "True"], // ignoreNonSummarySeekingQuery
    ["SummaryMatchingAdditionalPrompt", "If there are monetary numbers in the answers, they should be matched exactly."],
    ["Batch Size (1-10)", "2"], // BatchSize
    ["Time between Batches in Seconds (1-10)", "2"],
];


function testCaseTableColumns() {
    return ["ID", "Query", "Expected Summary", "Actual Summary", "Expected Link 1", "Expected Link 2", "Expected Link 3", "Summary Match", "First Link Match", "Link in Top 2", "Actual Link 1", "Actual Link 2", "Actual Link 3"];
}

const testCaseData = [
    ["ID", "Query", "Expected Summary", "Actual Summary", "Expected Link 1", "Expected Link 2", "Expected Link 3", "Summary Match", "First Link Match", "Link in Top 2", "Actual Link 1", "Actual Link 2", "Actual Link 3"],
    ["1", "query", "", "", "link1", "link2", "link3", "TRUE", "TRUE", "TRUE", "", "", ""],
];
describe("create Test Template Tables", () => {

    // Create a map to store the mock data.
    let configTableMap = new Map();
    configTableMap.set(0, [[]]);

    let dataTableMap = new Map();
    dataTableMap.set(0, [[]]);

    let mockData;

    beforeEach(() => {
         mockData = {
            context: {
                workbook: {
                    worksheets: {
                        name: "WorksheetName",
                        getActiveWorksheet: function () {
                            return this;
                        },
                        range: { // Config Table
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
                            }
                        },
                        getRange: function (str) {
                            return this.range;
                        },
                        getUsedRange: function () {
                            return this.range;
                        },
                        tables: { // Test Cases Table
                            name: "TestCasesTable",
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
                                if (str.endsWith("ConfigTable")) {
                                    return this.rows;
                                }
                                return this;
                            },
                            rows: {

                                values: [[]],
                                count: 1, 
                                add: function (str, values) {
                                    this.values = values;;
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

    it("should create th config table with the correct name and headers", async () => {

        // Create the final mock object from the seed object.
        const contextMock = new OfficeMockObject(mockData);

        global.Excel = contextMock;
        await createConfigTable();

        const worksheetName = contextMock.context.workbook.worksheets.name;
        expect(contextMock.context.workbook.worksheets.range.values).toEqual([["Vertex AI Search Parameters"]]);
        expect(contextMock.context.workbook.worksheets.range.format.font.bold).toEqual(true);
        expect(contextMock.context.workbook.worksheets.range.format.fill.color).toEqual('yellow');
        expect(contextMock.context.workbook.worksheets.range.format.font.size).toEqual(16);

        expect(contextMock.context.workbook.worksheets.tables.name).toEqual(`${worksheetName}.ConfigTable`);
        expect(contextMock.context.workbook.worksheets.tables.getHeaderRowRange().values).toEqual([["Config", "Value"]]);
        expect(contextMock.context.workbook.worksheets.tables.rows.values).toEqual(configValues.slice(1));
    });

    it("should create th Data table with the correct name and headers", async () => {

        // Create the final mock object from the seed object.
        const contextMock = new OfficeMockObject(mockData);
       

        global.Excel = contextMock;
        await createDataTable();
        const worksheetName = contextMock.context.workbook.worksheets.name;
        
        expect(contextMock.context.workbook.worksheets.tables.name).toEqual(`${worksheetName}.TestCasesTable`);

        expect(contextMock.context.workbook.worksheets.tables.getHeaderRowRange().values).toEqual(
            [testCaseTableColumns()]);

    });

    it("should populate the Date Table with the correct values", async () => {

        var mockTestData = {
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
                    
                                columns: {
                                    getItemOrNullObject: function (columnName) {
                                        return {
                                            values: [[]],
                                            load: function () {
                                                const columnIndex = testCaseData[0].indexOf(columnName);

                                                // If the column name is not found, return null
                                                if (columnIndex === -1) {
                                                    return false;
                                                }
                                                // Extract the values from the specified column
                                                this.values = testCaseData.map(row => [row[columnIndex]]);
                                                return true;
                                            }
                                        };
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
        // Create the final mock object from the seed object.
        const contextMock = new OfficeMockObject(mockTestData);


        global.Excel = contextMock;
        // Spy on the Showstatus function
        // Spy on showStatus
        const showStatusSpy = sinon.spy(globalThis, 'showStatus');


        await createConfigTable();
        // Fail the test ifshow status is called
        expect(showStatusSpy.notCalled).toBe(true);;
    
        // stub out jQuery calls
        const $stub = sinon.stub(globalThis, '$').returns({
            empty: sinon.stub(),
            append: sinon.stub(),
            val: sinon.stub(),
        });

        // Get the config
        
        const config = await getConfig();

        // Prepare the request response mock the call to VertexAISearch
        const { requestJson, url, expectedResponse } = mockVertexAISearchRequestResponse(
            1,
            200,
            './test/data/extractive_answer/test_vai_search_extractive_answer_request.json',
            './test/data/extractive_answer/test_vai_search_extractive_answer_response.json',
            config);
        

        // Populate the test table
        await runTests(config);

        expect(fetchMock.called()).toBe(true);
        // Assert request body is correct
        expect(JSON.parse(fetchMock.lastCall()[1].body)).toEqual(requestJson);
        // Assert URL is correct
        expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
       
        // Check if return values get populated
        //const actual_summary = mockTestData.context.workbook.worksheets.tables.getItem("TestCasesTable").columns.getItemOrNullObject('Actual Summary');
       // actual_summary.load();
       // expect(actual_summary.values[1][0]).toEqual("Google's revenue for the year ending December 31, 2022 was $2.5 billion. This is based on the deferred revenue as of December 31, 2021.");

        $stub.restore();
      
        
    });

    


});




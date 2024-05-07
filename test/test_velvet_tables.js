
import expect from 'expect';
import pkg from 'office-addin-mock';
import { createConfigTable, createDataTable } from '../src/velvet_tables.js';
const { OfficeMockObject } = pkg;

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
        expect(contextMock.context.workbook.worksheets.tables.rows.values).toEqual([
            ["Vertex AI Search Project Number", "384473000457"],
            ["Vertex AI Search DataStore Name", "alphabet-pdfs_1695783402380"],
            ["Vertex AI Project ID", "argolis-arau"],
            ["Vertex AI Location", "us-central1"],
            ["maxExtractiveAnswerCount (1-5)", "2"], //maxExtractiveAnswerCount
            ["maxExtractiveSegmentCount (1-5)", "0"], //maxExtractiveSegmentCount
            ["maxSnippetCount (1-5)", "0"], //maxSnippetCount
            ["Preamble (Customized Summaries)", ""],
            ["Summarization Model", "gemini-1.0-pro-002/answer_gen/v1"],
            ["summaryResultCount (1-5)", "2"],   //summaryResultCount
            ["useSemanticChunks (True or False)", "False"],   //useSemanticChunks
            ["ignoreAdversarialQuery (True or False)", "True"], // ignoreAdversarialQuery
            ["ignoreNonSummarySeekingQuery (True or False)", "True"], // ignoreNonSummarySeekingQuery
            ["SummaryMatchingAdditionalPrompt", "If there are monetory numbers in the answers, they should be matched exactly."],
            ["Batch Size (1-10)", "2"], // BatchSize
            ["Time between Batches in Seconds (1-10)", "2"], // BatchSize
        ]);
    });

    it("should create th Data table with the correct name and headers", async () => {

        // Create the final mock object from the seed object.
        const contextMock = new OfficeMockObject(mockData);
       

        global.Excel = contextMock;
        await createDataTable();
        const worksheetName = contextMock.context.workbook.worksheets.name;
        
        expect(contextMock.context.workbook.worksheets.tables.name).toEqual(`${worksheetName}.TestCasesTable`);

        expect(contextMock.context.workbook.worksheets.tables.getHeaderRowRange().values).toEqual(
            [["ID", "Query", "Expected Summary", "Actual Summary", "Expected Link 1", "Expected Link 2", "Expected Link 3", "Summary Match", "First Link Match", "Link in Top 2", "Actual Link 1", "Actual Link 2", "Actual Link 3"]]);

    });

    /* it("should return the correct values from Vertex AI Search", async () => {

       
        // Create the final mock object from the seed object.
        const contextMock = new OfficeMockObject(mockData);


        global.Excel = contextMock;
        await createConfigTable();
        await createDataTable();

        var config = {
            vertexAISearchProjectNumber: "384473000457",
            vertexAISearchDataStoreName: "alphabet-pdfs_1695783402380",
        }

        const { requestJson, url, expectedResponse } =  prepareVertexAISearchRequestResponse(
            './test/data/snippets/test_vai_search_snippet_request.json',
            './test/data/snippets/test_vai_search_snippet_response.json', config);
        
        
        await executeTests();

        expect(fetchMock.called()).toBe(true);
        // Assert request body is correct
        expect(JSON.parse(fetchMock.lastCall()[1].body)).toEqual(requestJson);
        // Assert URL is correct
        expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
        // Assert response is correct
        expect(result).toEqual(expectedResponse);
       
        
    }); */


});



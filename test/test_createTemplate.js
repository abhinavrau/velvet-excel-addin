
import expect from 'expect';
import pkg from 'office-addin-mock';
import { createConfigTable, createDataTable } from '../src/test_template.js';
const { OfficeMockObject } = pkg;

describe("create Test Template Tables", () => {

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
                        range: {
                            values: [["Config", "Value"]],
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
                        tables: {
                            add: function (str, flag) {
                                return this;
                            },
                            getHeaderRowRange: function () {
                                return this.header_row_range;
                            },
                            resize: function (str) {
                                
                            },
                            rows: {

                                values: [[]],
                                add: function (str, values) {
                                    this.values = values;;
                                }
                            },
                            header_row_range: {
                                values: [["Config", "Value"]]
                            }
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
            ["maxExtractiveAnswerCount (1-5)", "1"], //maxExtractiveAnswerCount
            ["Preamble", "Put your preamble here"],
            ["summaryResultCount (1-5)", "1"],   //summaryResultCount
            ["ignoreAdversarialQuery (True or False)", "True"], // ignoreAdversarialQuery
            ["ignoreNonSummarySeekingQuery (True or False)", "True"] // ignoreNonSummarySeekingQuery
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


});

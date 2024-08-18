
import expect from 'expect';
import pkg from 'office-addin-mock';

import { default as $, default as JQuery } from 'jquery';

import sinon from 'sinon';
import { showStatus } from '../src/ui.js';
import { createConfigTable, createDataTable } from '../src/velvet_tables.js';
import { calculateSimilarityUsingVertexAI, callVertexAISearch } from '../src/vertex_ai.js';


import { configValues, testCaseData } from '../src/common.js';
// mock the UI components
global.showStatus = showStatus;
global.$ = $;
global.JQuery = JQuery;

global.callVertexAISearch = callVertexAISearch;
global.calculateSimilarityUsingVertexAI = calculateSimilarityUsingVertexAI;


const { OfficeMockObject } = pkg;


describe("When Create Test Template Tables is clicked", () => {

    let mockData;
    var $stub;
    beforeEach(() => {
         $stub = sinon.stub(globalThis, '$').returns({
            empty: sinon.stub(),
            append: sinon.stub(),
            val: sinon.stub(),
            tabulator: sinon.stub(),
        });

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
    afterEach(() => {
        $stub.restore();
    });

    it("should create the Config table with the correct name and headers", async () => {

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
        expect(contextMock.context.workbook.worksheets.tables.getHeaderRowRange().values).toEqual([configValues[0]]);
        expect(contextMock.context.workbook.worksheets.tables.rows.values).toEqual(configValues.slice(1));
    });

    it("should create the Test Data table with the correct name and headers", async () => {

        // Create the final mock object from the seed object.
        const contextMock = new OfficeMockObject(mockData);


        global.Excel = contextMock;
        await createDataTable();
        const worksheetName = contextMock.context.workbook.worksheets.name;

        expect(contextMock.context.workbook.worksheets.tables.name).toEqual(`${worksheetName}.TestCasesTable`);

        expect(contextMock.context.workbook.worksheets.tables.getHeaderRowRange().values).toEqual(
            [testCaseData[0]]);

    });
});






import expect from "expect";
import pkg from "office-addin-mock";

import { default as $, default as JQuery } from "jquery";

import sinon from "sinon";
import {
  createSummarizationEvalConfigTable,
  createSummarizationEvalDataTable,
} from "../src/excel/excel_summarization_tables.js";
import { showStatus } from "../src/ui.js";

import { summarization_configValues, summarization_TableHeader } from "../src/common.js";
// mock the UI components
global.showStatus = showStatus;
global.$ = $;
global.JQuery = JQuery;

const { OfficeMockObject } = pkg;

describe("When Summarization Eval Tables is clicked", () => {
  let mockData;
  var $stub;
  beforeEach(() => {
    $stub = sinon.stub(globalThis, "$").returns({
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
            range: {
              // Config Table
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
            },
            getRange: function (str) {
              return this.range;
            },
            getUsedRange: function () {
              return this.range;
            },
            tables: {
              // Test Cases Table
              name: "TestCasesTable",
              add: function (str, flag) {
                return this;
              },
              getHeaderRowRange: function () {
                return this.header_row_range;
              },
              getRange: function () {
                return this.range;
              },
              resize: function (str) {},
              getItem: function (str) {
                // check is str ends with string "TestCasesTable"
                if (str.endsWith("ConfigTable")) {
                  return this.rows;
                }
                return this;
              },
              range: {
                format: {
                  font: {
                    bold: false,
                  },
                },
              },
              rows: {
                values: [[]],
                count: 1,
                add: function (str, values) {
                  this.values = values;
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
  });

  it("should create the Config table with the correct name and headers", async () => {
    // Create the final mock object from the seed object.
    const contextMock = new OfficeMockObject(mockData);

    global.Excel = contextMock;
    await createSummarizationEvalConfigTable();

    const worksheetName = contextMock.context.workbook.worksheets.name;
    expect(contextMock.context.workbook.worksheets.range.values).toEqual([
      ["Summarization Evaluation"],
    ]);
    expect(contextMock.context.workbook.worksheets.range.format.font.bold).toEqual(true);
    expect(contextMock.context.workbook.worksheets.range.format.fill.color).toEqual("yellow");
    expect(contextMock.context.workbook.worksheets.range.format.font.size).toEqual(20);

    expect(contextMock.context.workbook.worksheets.tables.name).toEqual(
      `${worksheetName}.ConfigTable`,
    );
    expect(contextMock.context.workbook.worksheets.tables.getHeaderRowRange().values).toEqual([
      summarization_configValues[0],
    ]);
    expect(contextMock.context.workbook.worksheets.tables.rows.values).toEqual(
      summarization_configValues.slice(1),
    );
  });

  it("should create the Test Data table with the correct name and headers", async () => {
    // Create the final mock object from the seed object.
    const contextMock = new OfficeMockObject(mockData);

    global.Excel = contextMock;
    await createSummarizationEvalDataTable();
    const worksheetName = contextMock.context.workbook.worksheets.name;

    expect(contextMock.context.workbook.worksheets.tables.name).toEqual(
      `${worksheetName}.TestCasesTable`,
    );

    expect(contextMock.context.workbook.worksheets.tables.getHeaderRowRange().values).toEqual([
      summarization_TableHeader[0],
    ]);
  });
});

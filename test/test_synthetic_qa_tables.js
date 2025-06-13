import expect from "expect";
import pkg from "office-addin-mock";

import { default as $, default as JQuery } from "jquery";

import sinon from "sinon";
import {
  createSyntheticQAConfigTable,
  createSyntheticQADataTable,
} from "../src/excel/excel_synthetic_qa_tables.js";
import { showStatus } from "../src/ui.js";

import { synth_q_and_a_configValues, synth_q_and_a_TableHeader } from "../src/common.js";
// mock the UI components
global.showStatus = showStatus;
global.$ = $;
global.JQuery = JQuery;

const { OfficeMockObject } = pkg;

describe("When Create Q&A Tables is clicked", () => {
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
            getItemOrNullObject: function (str) {
              return this;
            },
            getRange: function (str) {
              return this.range;
            },
            range: {
              // Config Table
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
                return this.format;
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
              // Test Cases Table
              name: "SyntheticQATable",
              add: function (str, flag) {
                return this;
              },
              getRange: function () {
                return this.range;
              },
              getHeaderRowRange: function () {
                return this.header_row_range;
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
     
    const data = {
      sheetName: "WorksheetName",
      vertexAiAppId: "l300-arau_1695783344117",
      vertexAiProjectNumber: "384473000457",
      vertexAiProjectId: "argolis-arau",
      vertexAiLocation: "us-central1",
    };

    const worksheetName = data.sheetName;
    await createSyntheticQAConfigTable(data);

    
    expect(contextMock.context.workbook.worksheets.range.values).toEqual([
      [`${worksheetName} - Synthetic Questions & Answers`],
    ]);

    expect(contextMock.context.workbook.worksheets.tables.name).toEqual(
      `${worksheetName}.ConfigTable`,
    );
    expect(contextMock.context.workbook.worksheets.tables.getHeaderRowRange().values).toEqual([
      synth_q_and_a_configValues[0],
    ]);
    expect(contextMock.context.workbook.worksheets.tables.rows.values).toEqual(
      synth_q_and_a_configValues.slice(1),
    );
  });

  it("should create the Test Data table with the correct name and headers", async () => {
    // Create the final mock object from the seed object.
    const contextMock = new OfficeMockObject(mockData);

    global.Excel = contextMock;
    await createSyntheticQADataTable("WorksheetName");
    const worksheetName = "WorksheetName";

     expect(contextMock.context.workbook.worksheets.tables.name).toEqual(
      `${worksheetName}.SyntheticQATable`,
    ); 

    expect(contextMock.context.workbook.worksheets.tables.getHeaderRowRange().values).toEqual([
      synth_q_and_a_TableHeader[0],
    ]);
  });
});

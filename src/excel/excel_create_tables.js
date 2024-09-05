import { appendError, showStatus } from "../ui.js";
export async function createExcelTable(
  title,
  titleCellLocation,
  tableType,
  valuesArray,
  tableRangeStart,
  tableRangeEnd,
  fontSize,
) {
  await Excel.run(async (context) => {
    try {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      currentWorksheet.load("name");
      await context.sync();
      const worksheetName = currentWorksheet.name;

      currentWorksheet.getRange().format.font.name = "Aptos";

      if (title) {
        var range = currentWorksheet.getRange(titleCellLocation);
        //range.merge(); // Merge the cells
        range.values = [[title]];
        range.format.font.bold = true;
        if (tableType === "ConfigTable") {
          range.format.fill.color = "yellow";
        }
        range.format.font.size = fontSize > 20 ? fontSize : 20;
        range.format.horizontalAlignment = "Center";
        range.format.verticalAlignment = "Center";
      }

      var excelTable = currentWorksheet.tables.add(tableRangeStart, true /*hasHeaders*/);
      excelTable.name = `${worksheetName}.${tableType}`;
      excelTable.getRange().format.font.size = fontSize;
      excelTable.showFilterButton = false;

      excelTable.getHeaderRowRange().values = [valuesArray[0]];

      if (tableType !== "TestCasesTable") {
        excelTable.rows.add(0, valuesArray.slice(1));
      }
      if (tableType === "TestCasesTable") {
        excelTable.resize(tableRangeEnd);
        excelTable.showFilterButton = false;
      }

      currentWorksheet.getUsedRange().format.autofitColumns();
      currentWorksheet.getUsedRange().format.autofitRows();

      await context.sync();
    } catch (error) {
      showStatus(`Exception when creating ${title}  Table: ${error.message}`, true);
      appendError(`Error creating ${title}  Table:`, error);
    }
  });
}

export async function insertPercentFormulaInSummaryCell(row, col, columnNameToCalc) {
  await Excel.run(async (context) => {
    try {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      currentWorksheet.load("name");
      await context.sync();
      const worksheetName = currentWorksheet.name;
      const summaryTable = currentWorksheet.tables.getItem(`${worksheetName}.SummaryTable`);
      var percentSummaryMatchCell = summaryTable.getRange().getCell(row, col);
      percentSummaryMatchCell.load();
      await context.sync();
      percentSummaryMatchCell.formulas = [
        [
          `=IF(COUNTA(${worksheetName}.TestCasesTable[${columnNameToCalc}])>0, COUNTIF(${worksheetName}.TestCasesTable[${columnNameToCalc}], TRUE)/COUNTA(${worksheetName}.TestCasesTable[${columnNameToCalc}]), 0)`,
        ],
      ];
      percentSummaryMatchCell.numberFormat = "0.00%";
      await context.sync();
    } catch (error) {
      showStatus(`Exception when creating Search Summary Formulas ${error.message}`, true);
      appendError(`Error creating Search Summary Formulas :`, error);
    }
  });
}

export async function insertAvgFormulaInSummaryCell(row, col, columnNameToCalc) {
  await Excel.run(async (context) => {
    try {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      currentWorksheet.load("name");
      await context.sync();
      const worksheetName = currentWorksheet.name;
      const summaryTable = currentWorksheet.tables.getItem(`${worksheetName}.SummaryTable`);
      var percentSummaryMatchCell = summaryTable.getRange().getCell(row, col);
      percentSummaryMatchCell.load();
      await context.sync();
      percentSummaryMatchCell.formulas = [
        [
          `=IF(COUNTA(${worksheetName}.TestCasesTable[${columnNameToCalc}])>0,AVERAGE(${worksheetName}.TestCasesTable[${columnNameToCalc}]), 0)`,
        ],
      ];
      percentSummaryMatchCell.numberFormat = "0.00%";
      await context.sync();
    } catch (error) {
      showStatus(`Exception when creating Search Summary Formulas ${error.message}`, true);
      appendError(`Error creating Search Summary Formulas :`, error);
    }
  });
}

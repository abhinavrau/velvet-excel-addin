import { appendError, showStatus, summaryFontSize, tableTitlesFontSize } from "../ui.js";
export async function createExcelTable(
  title,
  titleCellLocation,
  tableType,
  valuesArray,
  tableRangeStart,
  tableRangeEnd,
  fontSize,
  titlesFontSize = tableTitlesFontSize,
) {
  var worksheetName = "";
  await Excel.run(async (context) => {
    try {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      currentWorksheet.load("name");
      await context.sync();
      worksheetName = currentWorksheet.name;

      currentWorksheet.getRange().format.font.name = "Aptos";

      if (title !== null) {
        var range = currentWorksheet.getRange(titleCellLocation);

        range.values = [[title]];
        range.format.font.bold = true;
        range.format.font.size = titlesFontSize;
      }

      var excelTable = currentWorksheet.tables.add(tableRangeStart, true /*hasHeaders*/);
      excelTable.name = `${worksheetName}.${tableType}`;
      excelTable.getRange().format.font.size = fontSize;

      //excelTable.getRange().format.wrapText = true;
      excelTable.showFilterButton = false;

      excelTable.getHeaderRowRange().values = [valuesArray[0]];
     

      if (tableType === "ConfigTable") {
        excelTable.rows.add(0, valuesArray.slice(1));
      } else {
        excelTable.resize(tableRangeEnd);
        excelTable.showFilterButton = true;
      }
      await context.sync();
    } catch (error) {
      showStatus(`Exception when creating ${title}  Table: ${error.message}`, true);
      appendError(`Error creating ${title}  Table:`, error);
    }
  });

  return worksheetName;
}

export async function createFormula(
  worksheetName,
  labelRange,
  label,
  formulaRange,
  formula,
  fontSize = summaryFontSize,
  percent = true,
) {
  await Excel.run(async (context) => {
    try {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      var labelCell = currentWorksheet.getRange(labelRange);
      labelCell.clear();
      labelCell.format.font.name = "Aptos";
      labelCell.format.font.size = fontSize;
      labelCell.values = [[label]];
      var cell = currentWorksheet.getRange(formulaRange);
      cell.clear();

      cell.formulas = [[formula]];
      cell.format.font.name = "Aptos";
      cell.format.font.size = fontSize;
      if (percent) {
        cell.numberFormat = "0.00%";
      }

      await context.sync();
    } catch (error) {
      showStatus(`Exception createFormula ${error.message} with formula: ${formula}`, true);
      appendError(`Error createFormula with formula: ${formula}`, error);
    }
  });
}

export async function makeRowBold(range) {
  await Excel.run(async (context) => {
    try {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

      const rowRange = currentWorksheet.getRange(range);
      // Apply bold formatting
      rowRange.format.font.bold = true;
      rowRange.format.fill.color = "lightgray";

      await context.sync(); // Synchronize changes
    } catch (error) {
      showStatus(`Exception makeRowBold ${error.message}`, true);
      appendError(`Error makeRowBold :`, error);
    }
  });
}

export async function summaryHeading(range, text, fontSize = summaryFontSize) {
  Excel.run(async (context) => {
    // Get the active worksheet
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    // Get the range you want to format (adjust as needed)
    let cells = sheet.getRange(range); // Change "A1" to your target cell or range

    // Apply the formatting
    cells.format.fill.color = "darkblue";
    cells.format.font.color = "white";
    cells.format.font.size = fontSize;

    cells.getCell(0, 1).values = [[text]];

    // Sync the changes back to Excel
    await context.sync();
  });
}

import { appendError, showStatus, summaryFontSize } from "../ui.js";
export async function createExcelTable(
  title,
  titleCellLocation,
  tableType,
  valuesArray,
  tableRangeStart,
  tableRangeEnd,
  fontSize,
  titlesFontSize,
  sheetName,
) {
  await Excel.run(async (context) => {
    try {
      var currentWorksheet = null;
      if (sheetName !== null) {
        currentWorksheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      } else {
        throw new Error("Sheetname is null");
      }
      await context.sync();
      currentWorksheet.getRange().format.font.name = "Aptos";

      if (title !== null) {
        var range = currentWorksheet.getRange(titleCellLocation);

        range.values = [[title]];
        range.format.font.bold = true;
        range.format.font.size = titlesFontSize;
      }

      var excelTable = currentWorksheet.tables.add(tableRangeStart, true /*hasHeaders*/);
      // remove space from sheetName
      sheetName = sheetName.replace(/\s/g, "");
      excelTable.name = `${sheetName}.${tableType}`;
      excelTable.getRange().format.font.size = fontSize;

      //excelTable.getRange().format.wrapText = true;
      excelTable.showFilterButton = false;

      excelTable.getHeaderRowRange().values = [valuesArray[0]];

      if (tableType === "ConfigTable") {
        excelTable.rows.add(0, valuesArray.slice(1));
        excelTable.resize(tableRangeEnd);
      } else {
        excelTable.resize(tableRangeEnd);
        excelTable.showFilterButton = true;
      }
      await context.sync();
    } catch (error) {
      showStatus(`Exception when creating ${tableType}  Table: ${error.message}`, true);
      appendError(`Error creating ${tableType}  Table:`, error);
    }
  });

  return sheetName;
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
      const currentWorksheet = context.workbook.worksheets.getItemOrNullObject(worksheetName);
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

export async function makeRowBold(worksheetName, range) {
  await Excel.run(async (context) => {
    try {
      const currentWorksheet = context.workbook.worksheets.getItemOrNullObject(worksheetName);

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

export async function groupRows(worksheetName, range) {
  await Excel.run(async (context) => {
    try {
      const worksheet = context.workbook.worksheets.getItemOrNullObject(worksheetName);
      worksheet.getRange(range).group(Excel.GroupOption.byRows);
      // collapse the group

      await context.sync();
    } catch (error) {
      showStatus(`Exception groupRows ${error.message}`, true);
      appendError(`Error groupRows :`, error);
    }
  });
}

export async function summaryHeading(sheetName, range, text, fontSize = summaryFontSize) {
  await Excel.run(async (context) => {
    // Get the active worksheet
    let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);

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
export async function addImageToSheet(sheetName, imageUrl, cell) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
    const image = sheet.getRange(cell).addImageFromBase64(imageUrl, "png");
    image.setHeight(100);
    image.setWidth(100);
    await context.sync();
  }).catch((error) => {
    showStatus(`Error adding image: ${error.message}`, true);
    appendError(`Error adding image:`, error);
  });
}




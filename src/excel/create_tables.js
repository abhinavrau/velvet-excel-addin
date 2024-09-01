import { appendError, showStatus } from "../ui.js";

export async function createConfigTable(
  taskTitle,
  configValuesArray,
  tableRangeStart,
  tableRangeEnd,
) {
  await Excel.run(async (context) => {
    try {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      currentWorksheet.load("name");
      await context.sync();
      const worksheetName = currentWorksheet.name;

      var range = currentWorksheet.getRange("A1");
      range.values = [[taskTitle]];
      range.format.font.bold = true;
      range.format.fill.color = "yellow";
      range.format.font.size = 16;

      var configTable = currentWorksheet.tables.add(tableRangeStart, true /*hasHeaders*/);
      configTable.name = `${worksheetName}.ConfigTable`;

      configTable.getHeaderRowRange().values = [configValuesArray[0]];

      configTable.rows.add(null, configValuesArray.slice(1));

      currentWorksheet.getUsedRange().format.autofitColumns();
      currentWorksheet.getUsedRange().format.autofitRows();
      currentWorksheet.getRange(tableRangeEnd).format.wrapText = true; // wrap system instrcutions
      currentWorksheet.getRange(tableRangeEnd).format.shrinkToFit = true; // shrinkToFit system instrcutions

      await context.sync();
    } catch (error) {
      showStatus(`Exception when creating ${taskTitle} Config Table: ${error.message}`, true);
      appendError(`Error creating ${taskTitle} Config Table:`, error);

      return;
    }
  });
}

export async function createDataTable(
  taskTitle,
  tableHeaderArray,
  tableRangeStart,
  tableRangeEnd,
) {
  await Excel.run(async (context) => {
    try {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      currentWorksheet.load("name");
      await context.sync();
      const worksheetName = currentWorksheet.name;

      var velvetTable = currentWorksheet.tables.add(tableRangeStart, true /*hasHeaders*/);
      velvetTable.name = `${worksheetName}.TestCasesTable`;

      velvetTable.getHeaderRowRange().values = [tableHeaderArray[0]];

      velvetTable.resize(tableRangeEnd);
      currentWorksheet.getUsedRange().format.autofitColumns();
      currentWorksheet.getUsedRange().format.autofitRows();
      currentWorksheet.getUsedRange().format.wrapText = true;
      currentWorksheet.getUsedRange().format.shrinkToFit = true;

      await context.sync();
    } catch (error) {
      showStatus(`Exception when creating ${taskTitle}  DataTable: ${error.message}`, true);
      appendError(`Error creating  ${taskTitle} Data Table:`, error);
      return;
    }
  });
}

import { vertex_ai_search_configValues, vertex_ai_search_testTableHeader } from "../common.js";
import { appendError, showStatus } from "../ui.js";

export async function createVAIConfigTable() {
  await Excel.run(async (context) => {
    try {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      currentWorksheet.load("name");
      await context.sync();
      const worksheetName = currentWorksheet.name;

      var range = currentWorksheet.getRange("A1");
      range.values = [["Vertex AI Search Evaluation"]];
      range.format.font.bold = true;
      range.format.fill.color = "yellow";
      range.format.font.size = 16;

      var configTable = currentWorksheet.tables.add("A2:B2", true /*hasHeaders*/);
      configTable.name = `${worksheetName}.ConfigTable`;

      configTable.getHeaderRowRange().values = [vertex_ai_search_configValues[0]];

      configTable.rows.add(null, vertex_ai_search_configValues.slice(1));

      currentWorksheet.getUsedRange().format.autofitColumns();
      currentWorksheet.getUsedRange().format.autofitRows();
      currentWorksheet.getRange("A10:B10").format.wrapText = true; // wrap preamble
      currentWorksheet.getRange("A10:B10").format.shrinkToFit = true; // shrinkToFit preamble

      await context.sync();
    } catch (error) {
      showStatus(`Exception when creating Vertex AI Search Config Table: ${error.message}`, true);
      appendError("Error creating Config Table:", error);

      return;
    }
  });
}

export async function createVAIDataTable() {
  await Excel.run(async (context) => {
    try {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      currentWorksheet.load("name");
      await context.sync();
      const worksheetName = currentWorksheet.name;

      var velvetTable = currentWorksheet.tables.add("C17:O17", true /*hasHeaders*/);
      velvetTable.name = `${worksheetName}.TestCasesTable`;

      velvetTable.getHeaderRowRange().values = [vertex_ai_search_testTableHeader[0]];

      velvetTable.resize("C17:O118");
      currentWorksheet.getUsedRange().format.autofitColumns();
      currentWorksheet.getUsedRange().format.autofitRows();

      await context.sync();
    } catch (error) {
      showStatus(`Exception when creating  Vertex AI Search Data Table: ${error.message}`, true);
      appendError("Error creating Data Table:", error);
      return;
    }
  });
}

import {
  findIndexByColumnsNameIn2DArray,
  summarization_configValues,
  summarization_TableHeader,
} from "../common.js";
import {
  configTableFontSize,
  dataTableFontSize,
  sheetTitleFontSize,
  summaryFontSize,
  tableTitlesFontSize,
} from "../ui.js";
import { createExcelTable, createFormula, summaryHeading } from "./excel_create_tables.js";
export async function createSummarizationEvalConfigTable(data) {
  summarization_configValues[
    findIndexByColumnsNameIn2DArray(summarization_configValues, "Vertex AI Project ID")
  ][1] = data.vertexAiProjectId;

  summarization_configValues[
    findIndexByColumnsNameIn2DArray(summarization_configValues, "Vertex AI Location")
  ][1] = data.vertexAiLocation;

   summarization_configValues[
     findIndexByColumnsNameIn2DArray(summarization_configValues, "Gemini Model ID")
   ][1] = data.model;

  await createExcelTable(
    data.sheetName + " - Summary Evaluation",
    "A2",
    "ConfigTable",
    summarization_configValues,
    "A3:B3",
    "A14:B14",
    configTableFontSize,
    sheetTitleFontSize,
    data.sheetName,
  );
}

export async function createSummarizationEvalDataTable(sheetName) {
  await Excel.run(async (context) => {
    // Get the active worksheet
    let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
    await context.sync();
    sheet.getRange("A:A").format.columnWidth = 370;
    sheet.getRange("B:B").format.columnWidth = 370;

    sheet.getRange("E:E").format.columnWidth = 455;
    sheet.getRange("F:F").format.columnWidth = 455;

    sheet.getRange("G:G").format.columnWidth = 455;
    sheet.getRange("G:G").format.columnWidth = 455;

    sheet.getRange("E:E").format.wrapText = true;
    sheet.getRange("F:F").format.wrapText = true;

    await context.sync();
  });

  const worksheetName = await createExcelTable(
    "Summarization Test Cases",
    "E10",
    "SummarizationTestCasesTable",
    summarization_TableHeader,
    "E11:H11",
    "E11:H110",
    dataTableFontSize,
    tableTitlesFontSize,
    sheetName,
  );

  await summaryHeading(sheetName, "D8:E8", "Summarization Quallity");

  const summaryMatchCol = "Summary Quality";

  const summaryFormula = `=IFERROR(AVERAGE(IFERROR(--LEFT(${worksheetName}.SummarizationTestCasesTable[${summaryMatchCol}],1),FALSE)),0)`;
  await createFormula(
    worksheetName,
    "E9",
    "Avg. Summarization Quality (0-5)",
    "D9",
    summaryFormula,
    summaryFontSize,
    false,
  );
}

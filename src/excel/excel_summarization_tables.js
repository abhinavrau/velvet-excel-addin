import { summarization_configValues, summarization_TableHeader } from "../common.js";
import {
  configTableFontSize,
  dataTableFontSize,
  sheetTitleFontSize,
  summaryFontSize,
} from "../ui.js";
import { createExcelTable, createFormula, summaryHeading } from "./excel_create_tables.js";
export async function createSummarizationEvalConfigTable() {
  await createExcelTable(
    "Gemini Summarization Evaluation",
    "A2",
    "ConfigTable",
    summarization_configValues,
    "A3:B3",
    "A14:B14",
    configTableFontSize,
    sheetTitleFontSize,
  );
}

export async function createSummarizationEvalDataTable() {
  Excel.run(async (context) => {
    // Get the active worksheet
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    sheet.getRange("A:A").format.columnWidth = 370;
    sheet.getRange("B:B").format.columnWidth = 370;

    sheet.getRange("E:E").format.columnWidth = 455;
    sheet.getRange("F:F").format.columnWidth = 455;

    sheet.getRange("E:E").format.wrapText = true;
    sheet.getRange("F:F").format.wrapText = true;

    await context.sync();
  });

  const worksheetName = await createExcelTable(
    "Summarization Test Cases",
    "E10",
    "TestCasesTable",
    summarization_TableHeader,
    "E11:G11",
    "E11:G110",
    dataTableFontSize,
  );

  await summaryHeading("D4:E4", "Summarization Quallity");

  const summaryMatchCol = "summarization_quality";

  const summaryFormula = `=IFERROR(AVERAGE(IFERROR(--LEFT(${worksheetName}.TestCasesTable[${summaryMatchCol}],1),FALSE)),0)`;
  await createFormula(
    worksheetName,
    "E5",
    "Avg. Summarization Quality (0-5)",
    "D5",
    summaryFormula,
    summaryFontSize,
    false,
  );
}
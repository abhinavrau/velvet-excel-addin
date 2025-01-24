import { synth_q_and_a_configValues, synth_q_and_a_TableHeader } from "../common.js";
import {
  configTableFontSize,
  dataTableFontSize,
  sheetTitleFontSize,
  summaryFontSize,
} from "../ui.js";
import { createExcelTable, createFormula, summaryHeading } from "./excel_create_tables.js";
export async function createSyntheticQAConfigTable() {
  await createExcelTable(
    "Generate Synthetic Questions and Answers",
    "A2",
    "ConfigTable",
    synth_q_and_a_configValues,
    "A3:B3",
    "A13:B13",
    configTableFontSize,
    sheetTitleFontSize,
  );
}

export async function createSyntheticQADataTable() {
  Excel.run(async (context) => {
    // Get the active worksheet
    let sheet = context.workbook.worksheets.getActiveWorksheet();

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
    "Synthetic Questions and Answers",
    "E10",
    "SyntheticQATable",
    synth_q_and_a_TableHeader,
    "D11:H11",
    "D11:H110",
    dataTableFontSize,
  );

  await summaryHeading("D8:E8", "Generate Synthetic Q&A Quality");

  const summaryMatchCol = "Q & A Quality";

  const summaryFormula = `=IFERROR(AVERAGE(IFERROR(--LEFT(${worksheetName}.TestCasesTable[${summaryMatchCol}],1),FALSE)),0)`;
  await createFormula(
    worksheetName,
    "E9",
    "Avg. Synthetic Q&A Quality (0-5)",
    "D9",
    summaryFormula,
    summaryFontSize,
    false,
  );
}

import { synth_q_and_a_configValues, synth_q_and_a_TableHeader } from "../common.js";
import {
  configTableFontSize,
  dataTableFontSize,
  sheetTitleFontSize,
  summaryFontSize,
} from "../ui.js";
import {
  createExcelTable,
  createFormula,
  groupRows,
  makeRowBold,
  summaryHeading,
} from "./excel_create_tables.js";
export async function createSyntheticQAConfigTable(sheetName) {
  const worksheetName = await createExcelTable(
    sheetName + " - Synthetic Questions & Answers",
    "C2",
    "ConfigTable",
    synth_q_and_a_configValues,
    "A3:B3",
    "A3:B17",
    configTableFontSize,
    sheetTitleFontSize,
  );

  Excel.run(async (context) => {
    // Get the active worksheet
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    sheet.getRange("B9").format.wrapText = true;
    sheet.getRange("B10").format.wrapText = true;
    sheet.getRange("B13").format.wrapText = true;

    await context.sync();
  });

  await makeRowBold(worksheetName, "A4:B4");
  await makeRowBold(worksheetName, "A7:B7");
  await makeRowBold(worksheetName, "A11:B11");
  await makeRowBold(worksheetName, "A15:B15");

  await groupRows(worksheetName, "5:6");
  await groupRows(worksheetName, "8:10");
  await groupRows(worksheetName, "12:14");
  await groupRows(worksheetName, "16:17");
  await groupRows(worksheetName, "4:17");
}

export async function createSyntheticQADataTable(sheetName) {
  Excel.run(async (context) => {
    // Get the active worksheet
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    sheet.getRange("A:A").format.columnWidth = 275;

    sheet.getRange("B:B").format.columnWidth = 455;
    sheet.getRange("C:C").format.columnWidth = 455;
    sheet.getRange("D:D").format.columnWidth = 455;

    sheet.getRange("C:C").format.wrapText = true;
    sheet.getRange("D:D").format.wrapText = true;

    await context.sync();
  });
  const worksheetName = await createExcelTable(
    "Synthetic Questions and Answers",
    "A22",
    "SyntheticQATable",
    synth_q_and_a_TableHeader,
    "A23:E23",
    "A23:E124",
    dataTableFontSize,
  );

  await summaryHeading("A19:B19", "Generate Synthetic Q&A Quality");

  const summaryMatchCol = "Q & A Quality";

  const summaryFormula = `=IFERROR(AVERAGE(IFERROR(--LEFT(${worksheetName}.SyntheticQATable[${summaryMatchCol}],1),FALSE)),0)`;
  await createFormula(
    worksheetName,
    "A20",
    "Avg. Synthetic Q&A Quality (0-5)",
    "B20",
    summaryFormula,
    summaryFontSize,
    false,
  );
}

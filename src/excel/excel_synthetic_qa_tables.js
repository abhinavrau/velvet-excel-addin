import {
  findIndexByColumnsNameIn2DArray,
  synth_q_and_a_configValues,
  synth_q_and_a_TableHeader,
} from "../common.js";
import {
  configTableFontSize,
  dataTableFontSize,
  sheetTitleFontSize,
  summaryFontSize,
  tableTitlesFontSize,
} from "../ui.js";
import {
  createExcelTable,
  createFormula,
  groupRows,
  makeRowBold,
  summaryHeading,
} from "./excel_create_tables.js";
export async function createSyntheticQAConfigTable(data) {
  synth_q_and_a_configValues[
    findIndexByColumnsNameIn2DArray(synth_q_and_a_configValues, "Vertex AI Project ID")
  ][1] = data.vertexAiProjectId;

  synth_q_and_a_configValues[
    findIndexByColumnsNameIn2DArray(synth_q_and_a_configValues, "Vertex AI Location")
  ][1] = data.vertexAiLocation;

  const worksheetName = await createExcelTable(
    data.sheetName + " - Synthetic Questions & Answers",
    "C2",
    "ConfigTable",
    synth_q_and_a_configValues,
    "A3:B3",
    "A3:B17",
    configTableFontSize,
    sheetTitleFontSize,
    data.sheetName,
  );

  Excel.run(async (context) => {
    // Get the active worksheet
    let sheet = context.workbook.worksheets.getItemOrNullObject(data.sheetName);

    sheet.getRange("B9").format.wrapText = true;
    sheet.getRange("B10").format.wrapText = true;
    sheet.getRange("B13").format.wrapText = true;

    await context.sync();
  });

  await makeRowBold(data.sheetName, "A4:B4");
  await makeRowBold(data.sheetName, "A7:B7");
  await makeRowBold(data.sheetName, "A11:B11");
  await makeRowBold(data.sheetName, "A15:B15");

  await groupRows(data.sheetName, "5:6");
  await groupRows(data.sheetName, "8:10");
  await groupRows(data.sheetName, "12:14");
  await groupRows(data.sheetName, "16:17");
  await groupRows(data.sheetName, "4:17");
}

export async function createSyntheticQADataTable(sheetName) {
  Excel.run(async (context) => {
    // Get the active worksheet
    let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);

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
    tableTitlesFontSize,
    sheetName,
  );

  await summaryHeading(sheetName, "A19:B19", "Generate Synthetic Q&A Quality");

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

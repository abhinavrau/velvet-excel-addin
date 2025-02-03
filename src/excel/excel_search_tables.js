import { vertex_ai_search_configValues, vertex_ai_search_testTableHeader } from "../common.js";
import { configTableFontSize, dataTableFontSize, sheetTitleFontSize } from "../ui.js";
import {
  createExcelTable,
  createFormula,
  groupRows,
  makeRowBold,
  summaryHeading,
} from "./excel_create_tables.js";
export async function createVAIConfigTable(sheetName) {
  const worksheetName = await createExcelTable(
    sheetName + " - Vertex AI Search Evaluation",
    "C2",
    "ConfigTable",
    vertex_ai_search_configValues,
    "A3:B3",
    "A3:B24",
    configTableFontSize,
    sheetTitleFontSize,
  );

  Excel.run(async (context) => {
    // Get the active worksheet
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    sheet.getRange("B18").format.wrapText = true;
    sheet.getRange("B20").format.wrapText = true;

    await context.sync();
  });

  await makeRowBold(worksheetName, "A4:B4");
  await makeRowBold(worksheetName, "A9:B9");
  await makeRowBold(worksheetName, "A16:B16");
  await makeRowBold(worksheetName, "A22:B22");

  await groupRows(worksheetName, "5:8");
  await groupRows(worksheetName, "10:15");
  await groupRows(worksheetName, "17:21");
  await groupRows(worksheetName, "23:24");
  await groupRows(worksheetName, "4:24");
}

export async function createVAIDataTable(sheetName) {
  Excel.run(async (context) => {
    // Get the active worksheet
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    sheet.getRange("A:A").format.columnWidth = 275;

    sheet.getRange("B:B").format.columnWidth = 455;
    sheet.getRange("C:C").format.columnWidth = 455;
    sheet.getRange("D:D").format.columnWidth = 455;

    sheet.getRange("E:E").format.wrapText = true;
    sheet.getRange("F:F").format.wrapText = true;
    sheet.getRange("G:G").format.wrapText = true;

    await context.sync();
  });

  const worksheetName = await createExcelTable(
    "Search Test Cases",
    "A32",
    "TestCasesTable",
    vertex_ai_search_testTableHeader,
    "A33:N33",
    "A33:N134",
    dataTableFontSize,
  );

  await summaryHeading("A26:B26", "Evaluation Summary");

  const summaryMatchCol = "Summary Match";
  const summaryMatchFormula = `=IF(COUNTIF(${worksheetName}.TestCasesTable[${summaryMatchCol}], TRUE) + COUNTIF(${worksheetName}.TestCasesTable[${summaryMatchCol}], FALSE) > 0, COUNTIF(${worksheetName}.TestCasesTable[${summaryMatchCol}], TRUE) / (COUNTIF(${worksheetName}.TestCasesTable[${summaryMatchCol}], TRUE) + COUNTIF(${worksheetName}.TestCasesTable[${summaryMatchCol}], FALSE)), 0)`;
  await createFormula(worksheetName, "A27", "Summary Match Accuracy", "B27", summaryMatchFormula);

  const firstLinkMatchCol = "First Link Match";
  const firstLinkMatchFormula = `=IF(COUNTIF(${worksheetName}.TestCasesTable[${firstLinkMatchCol}], TRUE) + COUNTIF(${worksheetName}.TestCasesTable[${firstLinkMatchCol}], FALSE) > 0, COUNTIF(${worksheetName}.TestCasesTable[${firstLinkMatchCol}], TRUE) / (COUNTIF(${worksheetName}.TestCasesTable[${firstLinkMatchCol}], TRUE) + COUNTIF(${worksheetName}.TestCasesTable[${firstLinkMatchCol}], FALSE)), 0)`;
  await createFormula(worksheetName, "A28", "First Link Match", "B28", firstLinkMatchFormula);

  const linkInTop2MatchCol = "Link in Top 2";
  const linkInTop2MatchFormula = `=IF(COUNTIF(${worksheetName}.TestCasesTable[${linkInTop2MatchCol}], TRUE) + COUNTIF(${worksheetName}.TestCasesTable[${linkInTop2MatchCol}], FALSE) > 0, COUNTIF(${worksheetName}.TestCasesTable[${linkInTop2MatchCol}], TRUE) / (COUNTIF(${worksheetName}.TestCasesTable[${linkInTop2MatchCol}], TRUE) + COUNTIF(${worksheetName}.TestCasesTable[${linkInTop2MatchCol}], FALSE)), 0)`;
  await createFormula(worksheetName, "A29", "Link in Top 2", "B29", linkInTop2MatchFormula);

  const groundingScoreCol = "Grounding Score";
  const groundingScoreFormula = `=IF(COUNTA(${worksheetName}.TestCasesTable[${groundingScoreCol}])>0,AVERAGE(${worksheetName}.TestCasesTable[${groundingScoreCol}]), 0)`;
  await createFormula(
    worksheetName,
    "A30",
    "Average Grounding Score",
    "B30",
    groundingScoreFormula,
  );
}

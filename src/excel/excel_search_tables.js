import { vertex_ai_search_configValues, vertex_ai_search_testTableHeader } from "../common.js";
import { configTableFontSize, dataTableFontSize, sheetTitleFontSize } from "../ui.js";
import {
  createExcelTable,
  createFormula,
  makeRowBold,
  summaryHeading,
} from "./excel_create_tables.js";
export async function createVAIConfigTable() {
  const worksheetName = await createExcelTable(
    "Agent Builder Search Evaluation",
    "A2",
    "ConfigTable",
    vertex_ai_search_configValues,
    "A3:B3",
    "A3:B29",
    configTableFontSize,
    sheetTitleFontSize,
  );

  await makeRowBold("A4:B4");
  await makeRowBold("A10:B10");
  await makeRowBold("A18:B18");
  await makeRowBold("A24:B24");
  await makeRowBold("A27:B27");
}

export async function createVAIDataTable() {
  Excel.run(async (context) => {
    // Get the active worksheet
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    sheet.getRange("A:A").format.columnWidth = 275;
    sheet.getRange("B:B").format.columnWidth = 275;

    sheet.getRange("E:E").format.columnWidth = 455;
    sheet.getRange("F:F").format.columnWidth = 455;
    sheet.getRange("G:G").format.columnWidth = 455;

    sheet.getRange("E:E").format.wrapText = true;
    sheet.getRange("F:F").format.wrapText = true;
    sheet.getRange("G:G").format.wrapText = true;

    await context.sync();
  });

  const worksheetName = await createExcelTable(
    "Search Test Cases",
    "E10",
    "TestCasesTable",
    vertex_ai_search_testTableHeader,
    "D11:Q11",
    "D11:Q111",
    dataTableFontSize,
  );

  await summaryHeading("D4:E4", "Search Eval Results");

  const summaryMatchCol = "Summary Match";
  const summaryMatchFormula = `=IF(COUNTIF(${worksheetName}.TestCasesTable[${summaryMatchCol}], TRUE) + COUNTIF(${worksheetName}.TestCasesTable[${summaryMatchCol}], FALSE) > 0, COUNTIF(${worksheetName}.TestCasesTable[${summaryMatchCol}], TRUE) / (COUNTIF(${worksheetName}.TestCasesTable[${summaryMatchCol}], TRUE) + COUNTIF(${worksheetName}.TestCasesTable[${summaryMatchCol}], FALSE)), 0)`;
  await createFormula(worksheetName, "E5", "Summary Match Accuracy", "D5", summaryMatchFormula);

  const firstLinkMatchCol = "First Link Match";
  const firstLinkMatchFormula = `=IF(COUNTIF(${worksheetName}.TestCasesTable[${firstLinkMatchCol}], TRUE) + COUNTIF(${worksheetName}.TestCasesTable[${firstLinkMatchCol}], FALSE) > 0, COUNTIF(${worksheetName}.TestCasesTable[${firstLinkMatchCol}], TRUE) / (COUNTIF(${worksheetName}.TestCasesTable[${firstLinkMatchCol}], TRUE) + COUNTIF(${worksheetName}.TestCasesTable[${firstLinkMatchCol}], FALSE)), 0)`;
  await createFormula(worksheetName, "E6", "First Link Match", "D6", firstLinkMatchFormula);

  const linkInTop2MatchCol = "Link in Top 2";
  const linkInTop2MatchFormula = `=IF(COUNTIF(${worksheetName}.TestCasesTable[${linkInTop2MatchCol}], TRUE) + COUNTIF(${worksheetName}.TestCasesTable[${linkInTop2MatchCol}], FALSE) > 0, COUNTIF(${worksheetName}.TestCasesTable[${linkInTop2MatchCol}], TRUE) / (COUNTIF(${worksheetName}.TestCasesTable[${linkInTop2MatchCol}], TRUE) + COUNTIF(${worksheetName}.TestCasesTable[${linkInTop2MatchCol}], FALSE)), 0)`;
  await createFormula(worksheetName, "E7", "Link in Top 2", "D7", linkInTop2MatchFormula);

  const groundingScoreCol = "Grounding Score";
  const groundingScoreFormula = `=IF(COUNTA(${worksheetName}.TestCasesTable[${groundingScoreCol}])>0,AVERAGE(${worksheetName}.TestCasesTable[${groundingScoreCol}]), 0)`;
  await createFormula(worksheetName, "E8", "Average Grounding Score", "D8", groundingScoreFormula);
}

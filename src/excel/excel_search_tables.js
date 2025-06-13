import {
  findIndexByColumnsNameIn2DArray,
  vertex_ai_search_configValues,
  vertex_ai_search_testTableHeader,
} from "../common.js";
import {
  configTableFontSize,
  dataTableFontSize,
  sheetTitleFontSize,
  tableTitlesFontSize,
} from "../ui.js";
import {
  createExcelTable,
  createFormula,
  groupRows,
  makeRowBold,
  summaryHeading,
} from "./excel_create_tables.js";
export async function createVAIConfigTable(data) {
  vertex_ai_search_configValues[
    findIndexByColumnsNameIn2DArray(vertex_ai_search_configValues, "Vertex AI Search App ID")
  ][1] = data.config.vertexAISearchAppId;

  vertex_ai_search_configValues[
    findIndexByColumnsNameIn2DArray(vertex_ai_search_configValues, "Vertex AI Project ID")
  ][1] = data.config.vertexAIProjectID;

  vertex_ai_search_configValues[
    findIndexByColumnsNameIn2DArray(vertex_ai_search_configValues, "Vertex AI Location")
  ][1] = data.config.vertexAILocation;

  vertex_ai_search_configValues[
    findIndexByColumnsNameIn2DArray(vertex_ai_search_configValues, "Answer Model")
  ][1] = data.config.model;

  const worksheetName = await createExcelTable(
    data.sheetName + " - Vertex AI Search Evaluation",
    "C2",
    "ConfigTable",
    vertex_ai_search_configValues,
    "A3:B3",
    "A3:B24",
    configTableFontSize,
    sheetTitleFontSize,
    data.sheetName,
  );

  await Excel.run(async (context) => {
    // Get the active worksheet
    let sheet = context.workbook.worksheets.getItemOrNullObject(data.sheetName);

    sheet.getRange("B18").format.wrapText = true;
    sheet.getRange("B20").format.wrapText = true;

    await context.sync();
  });

  await makeRowBold(data.sheetName, "A4:B4");
  await makeRowBold(data.sheetName, "A9:B9");
  await makeRowBold(data.sheetName, "A16:B16");
  await makeRowBold(data.sheetName, "A22:B22");

  await groupRows(data.sheetName, "5:8");
  await groupRows(data.sheetName, "10:15");
  await groupRows(data.sheetName, "17:21");
  await groupRows(data.sheetName, "23:24");
  await groupRows(data.sheetName, "4:24");
}

export async function createVAIDataTable(sheetName) {
  await Excel.run(async (context) => {
    // Get the active worksheet
    let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);

    sheet.getRange("A:A").format.columnWidth = 275;

    sheet.getRange("B:B").format.columnWidth = 455;
    sheet.getRange("C:C").format.columnWidth = 455;
    sheet.getRange("D:D").format.columnWidth = 455;

    sheet.getRange("E:E").format.wrapText = true;
    sheet.getRange("F:F").format.wrapText = true;
    sheet.getRange("G:G").format.wrapText = true;

    await context.sync();
  });

  await createExcelTable(
    "Search Test Cases",
    "A32",
    "TestCasesTable",
    vertex_ai_search_testTableHeader,
    "A33:N33",
    "A33:N134",
    dataTableFontSize,
    tableTitlesFontSize,
    sheetName,
  );
  const worksheetName = sheetName;
  await summaryHeading(sheetName, "A26:B26", "Evaluation Summary");

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

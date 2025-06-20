import {
  findIndexByColumnsNameIn2DArray,
  synth_q_and_a_configValues,
  synth_q_and_a_TableHeader,
  vertex_ai_search_testTableHeader,
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
  ][1] = data.config.vertexAIProjectID;

  synth_q_and_a_configValues[
    findIndexByColumnsNameIn2DArray(synth_q_and_a_configValues, "Vertex AI Location")
  ][1] = data.config.vertexAILocation;

  synth_q_and_a_configValues[
    findIndexByColumnsNameIn2DArray(synth_q_and_a_configValues, "Gemini Model ID")
  ][1] = data.config.model;
  
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

export async function getSyntheticQAData(syntheticQASheetName) {
  let data = [];
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItemOrNullObject(syntheticQASheetName);
    const table = worksheet.tables.getItem(`${syntheticQASheetName}.SyntheticQATable`);
    const tableRange = table.getRange();
    tableRange.load("values");
    await context.sync();

    const tableValues = tableRange.values;
    const synthHeader = tableValues[0];
    const searchHeader = vertex_ai_search_testTableHeader[0];

    const questionIndex = synthHeader.indexOf("Generated Question");
    const expectedAnswerIndex = synthHeader.indexOf("Expected Answer");
    const gcsUriIndex = synthHeader.indexOf("GCS File URI");

    const searchQuestionIndex = searchHeader.indexOf("Query");
    const searchExpectedAnswerIndex = searchHeader.indexOf("Expected Summary");
    const searchExpectedLink1Index = searchHeader.indexOf("Expected Link 1");

    const alignedRows = tableValues.slice(1).map((row) => {
      const alignedRow = [];
      alignedRow[searchQuestionIndex] = row[questionIndex];
      alignedRow[searchExpectedAnswerIndex] = row[expectedAnswerIndex];
      alignedRow[searchExpectedLink1Index] = row[gcsUriIndex];
      // Fill the rest of the columns with empty strings
      for (let i = 0; i < searchHeader.length; i++) {
        if (
          i !== searchQuestionIndex &&
          i !== searchExpectedAnswerIndex &&
          i !== searchExpectedLink1Index
        ) {
          alignedRow[i] = "";
        }
      }
      return alignedRow;
    });
    data = alignedRows;
  });
  return data;
}

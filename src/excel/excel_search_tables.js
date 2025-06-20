import {
  findIndexByColumnsNameIn2DArray,
  getAccuracyFormula,
  getAverageFormula,
  vertex_ai_search_configValues,
  vertex_ai_search_testTableHeader,
} from "../common.js";
import { getSyntheticQAData } from "./excel_synthetic_qa_tables.js";
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

    sheet.getRange("B17").format.wrapText = true;
    sheet.getRange("B19").format.wrapText = true;

    await context.sync();
  });

  await makeRowBold(data.sheetName, "A4:B4");
  await makeRowBold(data.sheetName, "A8:B8");
  await makeRowBold(data.sheetName, "A15:B15");
  await makeRowBold(data.sheetName, "A21:B21");

  await groupRows(data.sheetName, "5:7");
  await groupRows(data.sheetName, "9:14");
  await groupRows(data.sheetName, "16:20");
  await groupRows(data.sheetName, "22:23");
  await groupRows(data.sheetName, "4:23");
}

export async function createVAIDataTable(sheetName, originalWorksheetName = null, sampleData = null) {
  let csvData = null;
  if (sampleData) {
    csvData = await loadSampleData(originalWorksheetName, sampleData);
  }
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
    csvData,
  );

  const worksheetName = sheetName;
  await summaryHeading(sheetName, "A26:B26", "Evaluation Summary");

  const summaryMatchCol = "Summary Match";
  const summaryMatchFormula = getAccuracyFormula(worksheetName, summaryMatchCol);
  await createFormula(worksheetName, "A27", "Summary Match Accuracy", "B27", summaryMatchFormula);

  const firstLinkMatchCol = "First Link Match";
  const firstLinkMatchFormula = getAccuracyFormula(worksheetName, firstLinkMatchCol);
  await createFormula(worksheetName, "A28", "First Link Match", "B28", firstLinkMatchFormula);

  const linkInTop2MatchCol = "Link in Top 2";
  const linkInTop2MatchFormula = getAccuracyFormula(worksheetName, linkInTop2MatchCol);
  await createFormula(worksheetName, "A29", "Link in Top 2", "B29", linkInTop2MatchFormula);

  const groundingScoreCol = "Grounding Score";
  const groundingScoreFormula = getAverageFormula(worksheetName, groundingScoreCol);
  await createFormula(
    worksheetName,
    "A30",
    "Average Grounding Score",
    "B30",
    groundingScoreFormula,
  );
}

function parseCSV(csv) {
  const rows = [];
  let currentRow = [];
  let currentField = "";
  let inQuotes = false;

  for (let i = 0; i < csv.length; i++) {
    const char = csv[i];

    if (inQuotes) {
      if (char === '"') {
        if (i + 1 < csv.length && csv[i + 1] === '"') {
          currentField += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        currentField += char;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
      } else if (char === ",") {
        currentRow.push(currentField);
        currentField = "";
      } else if (char === "\n" || char === "\r") {
        if (i > 0 && csv[i - 1] !== "\n" && csv[i - 1] !== "\r") {
          currentRow.push(currentField);
          rows.push(currentRow);
          currentRow = [];
          currentField = "";
        }
      } else {
        currentField += char;
      }
    }
  }

  if (currentField) {
    currentRow.push(currentField);
  }
  if (currentRow.length > 0) {
    rows.push(currentRow);
  }

  return rows;
}

async function getCurrentSheetData(originalWorksheetName) {
  let data = [];
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItemOrNullObject(originalWorksheetName);
    const table = worksheet.tables.getItemOrNullObject(`${originalWorksheetName}.TestCasesTable`);
    const tableRange = table.getRange();
    tableRange.load("values");
    await context.sync();
    const tableValues = tableRange.values;
    tableValues.shift();
    data = tableValues;
  });
  return data;
}

async function loadSampleData(originalWorksheetName, sampleData) {
  if (sampleData === "current_sheet") {
    return await getCurrentSheetData(originalWorksheetName);
  }
  let fileName = "";
  if (sampleData === "alphabet") {
    fileName = "alphabet-reports_dataset.csv";
  } else if (sampleData === "gemini_bank") {
    fileName = "gemini-bank_dataset.csv";
  } else if (sampleData === "device_manuals") {
    fileName = "user-manuals_dataset.csv";
  } else {
    return await getSyntheticQAData(sampleData);
  }
  console.log(`Fetching data from: assets/datasets/${fileName}`);

  const response = await fetch(`assets/datasets/${fileName}`);
  const csvData = await response.text();
  console.log("CSV data fetched successfully.");
  const csvRows = parseCSV(csvData).filter(
    (row) => row.length > 1 || (row.length === 1 && row[0] !== ""),
  );
  const csvHeader = csvRows[0];
  const tableHeader = vertex_ai_search_testTableHeader[0];
  console.log("CSV Header: " + JSON.stringify(csvHeader));
  console.log("Table Header: " + JSON.stringify(tableHeader));

  const columnIndexMap = tableHeader.map((headerName) => csvHeader.indexOf(headerName));
  console.log("Column Index Map: " + JSON.stringify(columnIndexMap));

  const alignedRows = csvRows.slice(1).map((row) => {
    const alignedRow = [];
    columnIndexMap.forEach((csvIndex, tableIndex) => {
      if (csvIndex !== -1) {
        alignedRow[tableIndex] = row[csvIndex];
      } else {
        alignedRow[tableIndex] = ""; // Or some default value
      }
    });
    return alignedRow;
  });
  return alignedRows;
}

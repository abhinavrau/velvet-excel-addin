import {
  alphabetInvestorPdfs,
  alphabetSecFilingsPdfs,
  cymalBankEmployeeGuidePdfs,
  findIndexByColumnsNameIn2DArray,
  geminiBankPoliciesProceduresPdfs,
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

  if (data.config.prompt) {
    
    synth_q_and_a_configValues[
      findIndexByColumnsNameIn2DArray(synth_q_and_a_configValues, "Prompt")
    ][1] = data.config.prompt;
  }
  
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


async function loadSampleData(sampleData) {
  if (sampleData === "alphabet-investor-pdfs") {
    return alphabetInvestorPdfs;
  } else if (sampleData === "alphabet-sec-filings-pdfs") {
    return alphabetSecFilingsPdfs;
  } else if (sampleData === "cymal-bank-employee-guide-pdfs") {
    return cymalBankEmployeeGuidePdfs;
  } else if (sampleData === "gemini-bank-policies-procedures-pdfs") {
    return geminiBankPoliciesProceduresPdfs;
  }
  return null;
}
export async function createSyntheticQADataTable(sheetName, sampleData = null) {

  let csvData = null;
  if (sampleData) {
    csvData = await loadSampleData(sampleData);
  }

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
    csvData,
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

    const iDIndex = synthHeader.indexOf("ID");
    const questionIndex = synthHeader.indexOf("Generated Question");
    const expectedAnswerIndex = synthHeader.indexOf("Expected Answer");
    const gcsUriIndex = synthHeader.indexOf("GCS File URI");

    const searchIDIndex = searchHeader.indexOf("ID");
    const searchQuestionIndex = searchHeader.indexOf("Query");
    const searchExpectedAnswerIndex = searchHeader.indexOf("Expected Summary");
    const searchExpectedLink1Index = searchHeader.indexOf("Expected Link 1");

    const alignedRows = tableValues.slice(1).map((row) => {
      const alignedRow = [];
      alignedRow[searchIDIndex] = row[iDIndex];
      alignedRow[searchQuestionIndex] = row[questionIndex];
      alignedRow[searchExpectedAnswerIndex] = row[expectedAnswerIndex];
      alignedRow[searchExpectedLink1Index] = row[gcsUriIndex];
      // Fill the rest of the columns with empty strings
      for (let i = 0; i < searchHeader.length; i++) {
        if (
          i !== searchIDIndex &&
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

export function generatePrompt(options) {
  // --- Input Validation (optional but recommended) ---
  if (
    !options.persona ||
    !options.answerVerbosity
  ) {
    throw new Error("Missing required options for prompt generation.");
  }

  // --- Constructing Optional Guidelines ---
  let additionalConsiderations = "";
  if (options.focusArea) {
    additionalConsiderations += `\n    * ${options.focusArea}`;
  }

  // --- The Prompt Template ---
  const promptTemplate = `
You are a ${options.persona}. Your goal is to find specific information within the provided document(s) that is relevant to your role and responsibilities.
You are detail-oriented and need accurate, contextually-aware answers.

**Generation Guidelines**

When generating the question-answer pairs, you must adhere to the following guidelines:

* 
* **Question Complexity & Style:** The questions should be formulated to how a ${options.persona} would make real-world inquiries. This includes:
    * **Direct Fact-Finding Questions:** (e.g., "What is the maximum liability coverage for X?")
    * **Clarification Questions:** (e.g., "Can you clarify the terms of the contract regarding this facet X?")
    * **Contextual Questions:** (e.g., "How does the liability coverage for X if I have already signed up for Y?")
* **Answer Format:** The answers should be ${options.answerVerbosity}.
* **Use of Context:** Leverage the provided document(s) to ensure that the answers are grounded in the context of the information available.
* **Avoiding Assumptions:** Do not make assumptions beyond the information provided in the document(s). If the answer is not available, indicate that clearly.
* **Relevance:** Ensure that the questions and answers are relevant to the role of a ${options.persona} and the context of the document(s).
* **Additional Considerations:**${additionalConsiderations || "\n    * None."}

`;
  
  /* * **Comparative Questions:** (e.g., "What are the differences in the maintenance schedules for the A-series and B-series equipment?")
    * **Multi-Detail Questions:** These questions should require synthesizing information from multiple parts of the document to form a complete answer. (e.g., "What are the security protocols and the associated reporting procedures for a data breach?")
    * **Scenario-Based Questions:** Frame some questions as if you are facing a real-world problem. (e.g., "I am a new project manager. What are the first three steps I need to take to initiate a project according to the 'Project Initiation' section?")
    * 
    * */

  return promptTemplate.trim();
}

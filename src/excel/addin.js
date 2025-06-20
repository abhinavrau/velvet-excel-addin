import {
  getSearchConfigFromActiveSheet,
  getSummarizationConfigFromActiveSheet,
  getSyntheticQAConfigFromActiveSheet,
} from "./excel_common.js";

import {
  createSearchRunsTable,
  createSummaryRunsTable,
  createSyntheticQnARunsTable,
} from "./excel_eval_runs_tables.js";
import { findSheetsWithTableSuffix } from "./excel_helper.js";
import { ExcelSearchRunner } from "./excel_search_runner.js";
import { SummarizationRunner } from "./excel_summarization_runner.js";
import {
  createSummarizationEvalConfigTable,
  createSummarizationEvalDataTable,
} from "./excel_summarization_tables.js";
import { SyntheticQARunner } from "./excel_synthetic_qa_runner.js";
import {
  createSyntheticQAConfigTable,
  createSyntheticQADataTable,
  generatePrompt,
} from "./excel_synthetic_qa_tables.js";

import { createVAIConfigTable, createVAIDataTable } from "./excel_search_tables.js";

let excelSearchRunner;
let summarizationRunner;
let syntheticQuestionAnswerRunner;
// Initialize Office API
Office.onReady((info) => {
  OfficeExtension.config.extendedErrorLogging = true;
  // Check that we loaded into Excel
  if (info.host === Office.HostType.Excel) {
    excelSearchRunner = new ExcelSearchRunner();
    setupButtonEvents(
      "createSearchTables",
      createSearchTables,
      "executeSearchTests",
      async function () {
        const config = await excelSearchRunner.getSearchConfig();
        if (config == null) return;
        await excelSearchRunner.executeSearchTests(config);
      },
      "cancelSearchTests",
      async function () {
        await excelSearchRunner.cancelProcessing();
      },
    );
    syntheticQuestionAnswerRunner = new SyntheticQARunner();
    setupButtonEvents(
      "createGenQATables",
      createSyntheticQATables,
      "generateQAData",
      async function () {
        const config = await syntheticQuestionAnswerRunner.getSyntheticQAConfig();
        if (config == null) return;
        await syntheticQuestionAnswerRunner.createSyntheticQAData(config);
      },
      "cancelGenerateQAData",
      async function () {
        await syntheticQuestionAnswerRunner.cancelProcessing();
      },
    );

    summarizationRunner = new SummarizationRunner();
    setupButtonEvents(
      "createSummarizationTables",
      createSummarizationTables,
      "genSummarizationData",
      async function () {
        const config = await summarizationRunner.getSummarizationConfig();
        if (config == null) return;
        await summarizationRunner.createSummarizationData(config);
      },
      "cancelSummarizationData",
      async function () {
        await summarizationRunner.cancelProcessing();
      },
    );
  }
});

function setupButtonEvents(
  createTableButtonId,
  fn_createTables,
  runTaskButtonId,
  fn_runTask,
  cancelJobButtonId,
  fn_cancelTask,
) {
  document.getElementById(createTableButtonId).onclick = fn_createTables;
  const runTaskButton = document.getElementById(runTaskButtonId);
  const cancelJobButton = document.getElementById(cancelJobButtonId);

  runTaskButton.addEventListener("click", async () => {
    $("#log-pane").tabulator("clearData");
    runTaskButton.style.visibility = "hidden";
    cancelJobButton.style.visibility = "visible";

    try {
      await fn_runTask();
    } finally {
      runTaskButton.style.visibility = "visible";
      cancelJobButton.style.visibility = "hidden";
    }
  });

  cancelJobButton.addEventListener("click", async () => {
    try {
      await fn_cancelTask();
    } finally {
      runTaskButton.style.visibility = "visible";
      cancelJobButton.style.visibility = "hidden";
    }
  });
}

async function createSearchTables() {
  const config = await getSearchConfigFromActiveSheet();

  await promptSheetName("search", config, async (arg) => {
    //console.log("data:" + arg.message);
    const data = JSON.parse(arg.message);
    const sheetName = data.sheetName;

    await createNewSheet(sheetName, "Search Evals", createSearchRunsTable);
    await createVAIConfigTable(data);
    await createVAIDataTable(sheetName, data.config.originalWorksheetName, data.sampleData);
  });
}

async function createSyntheticQATables() {
  const config = await getSyntheticQAConfigFromActiveSheet();

  await promptSheetName("synthQA", config, async (arg) => {
    const data = JSON.parse(arg.message);
    console.log("data:" + arg.message);
    const sheetName = data.sheetName;
    if (data.config.persona && data.config.verbosity) {
      data.config.prompt = generatePrompt({
        persona: data.config.persona,
        answerVerbosity: data.config.verbosity,
        focusArea: data.config.focusArea,
      });
    }
    await createNewSheet(sheetName, "Synthetic QnAs", createSyntheticQnARunsTable);
    await createSyntheticQAConfigTable(data);
    await createSyntheticQADataTable(sheetName);
  });
}

async function createSummarizationTables() {
  const config = await getSummarizationConfigFromActiveSheet();

  await promptSheetName("summary", config, async (arg) => {
    const data = JSON.parse(arg.message);
    const sheetName = data.sheetName;
    await createNewSheet(sheetName, "Summarization Evals", createSummaryRunsTable);
    await createSummarizationEvalConfigTable(data);
    await createSummarizationEvalDataTable(sheetName);
  });
}

let dialog = null;

async function promptSheetName(type, config, callback) {
  const baseUrl = window.location.origin;
  var page = "";
  var synthQASheets = [];

  switch (type) {
    case "search":
      page = `search-dialog.html`;
      synthQASheets = await findSheetsWithTableSuffix("SyntheticQATable");

      break;
    case "synthQA":
      page = `synth-qa-dialog.html`;
      break;
    case "summary":
      page = `summary-dialog.html`;
      break;
    default:
      page = `search-dialog.html`;
      break;
  }
  // Creae URL object
  const url = new URL(`${baseUrl}/${page}`);

  if (config !== null) {
    const encodedConfig = encodeURIComponent(JSON.stringify(config));
    url.searchParams.set("config", encodedConfig);
  }
  if (synthQASheets && synthQASheets.length > 0) {
    const synthQASheetsEncoded = encodeURIComponent(JSON.stringify(synthQASheets));
    url.searchParams.set("synthQASheets", synthQASheetsEncoded);
  }

  // pass it to the dialog popup.html below
  Office.context.ui.displayDialogAsync(
    url.href,
    { height: 55, width: 35 },

    function (result) {
      dialog = result.value;
      dialog.addEventHandler(
        Microsoft.Office.WebExtension.EventType.DialogMessageReceived,
        callback,
      );
    },
  );
}

async function createNewSheet(sheetName, historySheetName, fn_createResults) {
  console.log("SheetName:" + sheetName);
  dialog.close();

  // Create blank worksheet with sheetName
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;

    // check if History sheet is created.
    const historySheet = context.workbook.worksheets.getItemOrNullObject(historySheetName);
    await context.sync();
    if (historySheet.isNullObject) {
      const newHistorySheet = sheets.add(historySheetName);
      await newHistorySheet.activate(); // Activate the new sheet
      await context.sync(); // Synchronize changes
      await fn_createResults(historySheetName);
    }

    // Add a new worksheet
    const sheet = sheets.add(sheetName);
    await sheet.activate(); // Activate the new sheet
    await context.sync(); // Synchronize changes
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

import { ExcelSearchRunner } from "./excel_search_runner.js";
import { createVAIConfigTable, createVAIDataTable } from "./excel_search_tables.js";
import { SummarizationRunner } from "./excel_summarization_runner.js";
import { SyntheticQARunner } from "./excel_synthetic_qa_runner.js";

import {
  createSyntheticQAConfigTable,
  createSyntheticQADataTable,
} from "./excel_synthetic_qa_tables.js";

import {
  createSummarizationEvalConfigTable,
  createSummarizationEvalDataTable,
} from "./excel_summarization_tables.js";

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
  await promptSheetName(async (arg) => {
    const sheetName = arg.message;
    await createNewSheet(sheetName);
    await createVAIConfigTable(arg.message);
    await createVAIDataTable(arg.message);
  });
}

async function createSyntheticQATables() {
  await promptSheetName(async (arg) => {
    const sheetName = arg.message;
    await createNewSheet(sheetName);
    await createSyntheticQAConfigTable(sheetName);
    await createSyntheticQADataTable(sheetName);
  });
}

async function createSummarizationTables() {
  await promptSheetName(async (arg) => {
    const sheetName = +arg.message;
    await createNewSheet(sheetName);
    await createSummarizationEvalConfigTable(sheetName);
    await createSummarizationEvalDataTable(sheetName);
  });
}

let dialog = null;

async function promptSheetName(callback) {
  Office.context.ui.displayDialogAsync(
    "https://localhost:5500/popup.html",
    { height: 45, width: 55 },

    function (result) {
      dialog = result.value;
      dialog.addEventHandler(
        Microsoft.Office.WebExtension.EventType.DialogMessageReceived,
        callback,
      );
    },
  );
}

async function createNewSheet(sheetName) {
  console.log("SheetName:" + sheetName);
  dialog.close();

  // Create blank worksheet with sheetName
  Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const sheet = sheets.add(sheetName); // Add a new worksheet
    sheet.activate(); // Activate the new sheet
    await context.sync(); // Synchronize changes
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

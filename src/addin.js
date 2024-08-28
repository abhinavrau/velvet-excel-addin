import { ExcelSearchRunner } from "./excel/excel_search_runner.js";
import { createVAIConfigTable, createVAIDataTable } from "./search_tables.js";

import {
  createSyntheticQAData,
  getSyntheticQAConfig,
  stopSyntheticData,
} from "./synthetic_qa_runner.js";
import { createSyntheticQAConfigTable, createSyntheticQADataTable } from "./synthetic_qa_tables.js";

import {
  createSummarizationData,
  getSummarizationConfig,
  stopSummarizationData,
} from "./summarization_runner.js";
import {
  createSummarizationEvalConfigTable,
  createSummarizationEvalDataTable,
} from "./summarization_tables.js";

let excelSearchRunner;

// Initialize Office API
Office.onReady((info) => {
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
        await excelSearchRunner.stopSearchTests();
      },
    );

    setupButtonEvents(
      "createGenQATables",
      createSyntheticQATables,
      "generateQAData",
      async function () {
        await createSyntheticData();
      },
      "cancelGenerateQAData",
      async function () {
        await stopSyntheticData();
      },
    );

    setupButtonEvents(
      "createSummarizationTables",
      createSummarizationTables,
      "genSummarizationData",
      async function () {
        await genSummarizationData();
      },
      "cancelSummarizationData",
      async function () {
        await stopSummarizationData();
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
  await createVAIConfigTable();
  await createVAIDataTable();
}

async function createSyntheticQATables() {
  await createSyntheticQAConfigTable();
  await createSyntheticQADataTable();
}

async function createSyntheticData() {
  const config = await getSyntheticQAConfig();
  if (config == null) return;
  await createSyntheticQAData(config);
}

async function createSummarizationTables() {
  await createSummarizationEvalConfigTable();
  await createSummarizationEvalDataTable();
}

async function genSummarizationData() {
  const config = await getSummarizationConfig();
  if (config == null) return;
  await createSummarizationData(config);
}

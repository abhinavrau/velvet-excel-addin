import { synth_qa_runs_table, test_search_runs_table } from "../common.js";
import { configTableFontSize, sheetTitleFontSize } from "../ui.js";
import { createExcelTable } from "./excel_create_tables.js";
export async function createSearchRunsTable(sheetName) {
  await createExcelTable(
    "Search Eval Runs",
    "A2",
    "TestRunsTable",
    test_search_runs_table,
    "A3:M3",
    "A3:M50",
    configTableFontSize,
    sheetTitleFontSize,
    sheetName,
  );
}

export async function createSyntheticQnARunsTable(sheetName) {
  await createExcelTable(
    "Synthetic Questions & Answers Eval Runs",
    "A2",
    "SynthQARunsTable",
    synth_qa_runs_table,
    "A3:F3",
    "A3:F50",
    configTableFontSize,
    sheetTitleFontSize,
    sheetName,
  );
}


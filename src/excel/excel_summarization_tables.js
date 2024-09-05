import { summarization_configValues, summarization_TableHeader } from "../common.js";
import { createExcelTable } from "./excel_create_tables.js";

export async function createSummarizationEvalConfigTable() {
  createExcelTable(
    "Summarization Evaluation",
    "A2",
    "ConfigTable",
    summarization_configValues,
    "A2:B2",
    "A13:B13",
  );
}

export async function createSummarizationEvalDataTable() {
  createExcelTable(
    "Summarization Test Cases",
    "C13",
    "TestCasesTable",
    summarization_TableHeader,
    "C14:F14",
    "C14:F115",
  );
}

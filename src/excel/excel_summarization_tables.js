import { summarization_configValues, summarization_TableHeader } from "../common.js";
import { createConfigTable, createDataTable } from "./excel_create_tables.js";

export async function createSummarizationEvalConfigTable() {
  createConfigTable("Summarization Evaluation", summarization_configValues, "A2:B2", "A13:B13");
}

export async function createSummarizationEvalDataTable() {
  createDataTable("Summarization Evaluation", summarization_TableHeader, "C14:F14", "C14:F115");
}

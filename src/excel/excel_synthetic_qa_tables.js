import { synth_q_and_a_configValues, synth_q_and_a_TableHeader } from "../common.js";
import { createExcelTable } from "./excel_create_tables.js";

export async function createSyntheticQAConfigTable() {
  createExcelTable(
    "Generate Synthetic Questions and Answers",
    "A2",
    "ConfigTable",
    synth_q_and_a_configValues,
    "A2:B2",
    "A12:B12",
  );
}

export async function createSyntheticQADataTable() {
  createExcelTable(
    "Table Synthetic Questions and Answers",
    "C12",
    "TestCasesTable",
    synth_q_and_a_TableHeader,
    "C13:H13",
    "C13:H113",
  );
}

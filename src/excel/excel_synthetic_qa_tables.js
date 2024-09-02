import { synth_q_and_a_configValues, synth_q_and_a_TableHeader } from "../common.js";
import { createConfigTable, createDataTable } from "./excel_create_tables.js";

export async function createSyntheticQAConfigTable() {
  createConfigTable(
    "Generate Synthetic Questions and Answers",
    synth_q_and_a_configValues,
    "A2:B2",
    "A12:B12",
  );
}

export async function createSyntheticQADataTable() {
  createDataTable(
    "Generate Synthetic Questions and Answers",
    synth_q_and_a_TableHeader,
    "C13:H13",
    "C13:H113",
  );        
}

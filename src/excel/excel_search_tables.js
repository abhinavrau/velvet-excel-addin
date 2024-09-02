import { vertex_ai_search_configValues, vertex_ai_search_testTableHeader } from "../common.js";
import { createConfigTable, createDataTable } from "./excel_create_tables.js";
export async function createVAIConfigTable() {
  createConfigTable(
    "Vertex AI Search Evaluation",
    vertex_ai_search_configValues,
    "A2:B2",
    "A19:B19",
  );
}

export async function createVAIDataTable() {
  createDataTable(
    "Vertex AI Search Evaluation",
    vertex_ai_search_testTableHeader,
    "C19:P19",
    "C19:P119",
  );
}

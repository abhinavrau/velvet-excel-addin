import {
  vertex_ai_search_configValues,
  vertex_ai_search_summary_Table,
  vertex_ai_search_testTableHeader,
} from "../common.js";
import {
  createExcelTable,
  insertAvgFormulaInSummaryCell,
  insertPercentFormulaInSummaryCell,
} from "./excel_create_tables.js";

export async function createVAIConfigTable() {
  createExcelTable(
    "Vertex AI Search Evaluation",
    "A1",
    "ConfigTable",
    vertex_ai_search_configValues,
    "A2:B2",
    "A2:B8",
    15,
  );
}
export async function createSummaryTable() {
  createExcelTable(
    "Search Eval Results",
    "D2",
    "SummaryTable",
    vertex_ai_search_summary_Table,
    "D3:H3",
    "D4:H4",
    25,
  );

  insertPercentFormulaInSummaryCell(1, 1, "Summary Match");
  insertPercentFormulaInSummaryCell(1, 2, "First Link Match");
  insertPercentFormulaInSummaryCell(1, 3, "Link in Top 2");
  insertAvgFormulaInSummaryCell(1, 4, "Grounding Score");
}
export async function createVAIDataTable() {
  createExcelTable(
    "Search Test Cases",
    "D7",
    "TestCasesTable",
    vertex_ai_search_testTableHeader,
    "D8:Q8",
    "D8:Q108",
    15,
  );
}

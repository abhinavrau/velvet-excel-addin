import { appendError, showStatus } from "../ui.js";

export function getColumn(table, columnName) {
  try {
    const column = table.columns.getItemOrNullObject(columnName);
    column.load();
    return column;
  } catch (error) {
    appendError("Error getColumn:", error);
    showStatus(`Exception when getting column: ${JSON.stringify(error)}`, true);
  }
}

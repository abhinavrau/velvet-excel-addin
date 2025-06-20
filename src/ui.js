export const tableTitlesFontSize = 18;
export const configTableFontSize = 14;
export const dataTableFontSize = 14;
export const summaryFontSize = 18;
export const sheetTitleFontSize = 20;

export function showStatus(message, isError) {
  const statusDiv = $(".status");
  statusDiv.empty();

  // remove any existing color classes from previous status messages
  statusDiv.removeClass("bg-gray-50 text-gray-700 bg-green-100 text-green-700 bg-red-100 text-red-700");

  if (isError) {
    statusDiv.addClass("bg-red-100 text-red-700");
    const title = $("<p/>", { class: "font-bold", text: "An error occurred" });
    const msg = $("<p/>", { text: message });
    statusDiv.append(title).append(msg);
  } else {
    statusDiv.addClass("bg-green-100 text-green-700");
    const title = $("<p/>", { class: "font-bold", text: "Success" });
    const msg = $("<p/>", { text: message });
    statusDiv.append(title).append(msg);
  }
}

export function appendLog(message) {
  appendError(message, null);
}
export function appendError(message, error) {
  const newLogEntry = {
    time: new Date().toLocaleTimeString(), // Get only time
    level: error !== null ? "ERROR" : "INFO",
    message: error !== null ? message + "\n" + error.message : message,
  };

  $("#log-pane").tabulator("addRow", newLogEntry, "top"); // Add the new log entry to the top of the table

  if (error !== null) {
    console.error(message + "\n" + error.message);
  } else {
    console.log(message);
  }
}

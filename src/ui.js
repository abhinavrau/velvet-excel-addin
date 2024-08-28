export function showStatus(message, isError) {
  $(".status").empty();

  var element = $("<div/>", {
    class: `status-card ms-depth-4 ${isError ? "error-msg" : "success-msg"}`,
  }).append(
    $("<p/>", {
      class: "ms-fontSize-24 ms-fontWeight-bold",
      text: isError ? "An error occurred" : "Success",
      class: "ms-fontSize-16 ms-fontWeight-regular",
      text: message,
    })
  );

  $(".status").append(element);
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



import { showStatus } from "./ui.js";
import { executeTests } from './velvet_runner.js';
import { createConfigTable, createDataTable } from './velvet_tables.js';

// Initialize Office API
Office.onReady((info) => {
    // Check that we loaded into Excel
    if (info.host === Office.HostType.Excel) {
        document.getElementById("createTables").onclick = createTestTemplateTables;
        document.getElementById("executeTests").onclick = runTests;
    }
});


async function createTestTemplateTables() {

    try {
        await createConfigTable();
    } catch (error) {
        showStatus(`Exception when creating Config Table: ${JSON.stringify(error)}`, true);
    }
    try {
        await createDataTable();
    } catch (error) {
        showStatus(`Exception when creating Data Table: ${JSON.stringify(error)}`, true);
    }

}

async function runTests() {
    const status = await executeTests();
    let message = "";
    if (status.isError) {
        message = `Error: ${status.message}. Number of tests failed: ${status.numFailed}`;
    } else {
        message = `Success: ${status.message} Num of tests done successfully ${status.numDone} Num of tests failed: ${status.numFailed}`;
    }
    showStatus(message, status.isError);
}

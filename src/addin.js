import { executeTests } from './velvet_runner.js';
import { createConfigTable, createDataTable } from './velvet_tables.js';
import { showStatus } from "./ui.js";

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
    try {
        await executeTests();
        showStatus("Finished Successfully!", false);
    } catch (error) {
        showStatus(`Exception when running tests: ${JSON.stringify(error)}`, true);
    }
}

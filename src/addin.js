import { getConfig, executeTests } from './velvet_runner.js';
import { createConfigTable, createDataTable } from './velvet_tables.js';

// Initialize Office API
Office.onReady((info) => {
    // Check that we loaded into Excel
    if (info.host === Office.HostType.Excel) {
        document.getElementById("createTables").onclick = createTestTemplateTables;
        document.getElementById("executeTests").onclick = runTests;
    }
});


export async function createTestTemplateTables() {

    await createConfigTable();
    await createDataTable();
}

export async  function runTests() {
    const config = await getConfig();
    if (config == null)
        return;
     await executeTests(config);
}

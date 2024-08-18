import { executeTests, getConfig, stopTests } from './velvet_runner.js';
import { createConfigTable, createDataTable } from './velvet_tables.js';

// Initialize Office API
Office.onReady((info) => {
    // Check that we loaded into Excel
    if (info.host === Office.HostType.Excel) {
        document.getElementById("createTables").onclick = createTestTemplateTables;
        const executeTestsButton = document.getElementById("executeTests");
        const cancelTestsButton = document.getElementById("cancelTests");

        executeTestsButton.addEventListener("click", async () => {
    
            executeTestsButton.style.visibility = "hidden";
            cancelTestsButton.style.visibility = "visible";
        
            try {
                 await runTests();
            } finally {
                
                executeTestsButton.style.visibility = "visible";
                cancelTestsButton.style.visibility = "hidden";
            }
        });
        

        cancelTestsButton.addEventListener("click", async () => {

            try {
                 await cancelTests();
            } finally {

                executeTestsButton.style.visibility = "visible";
                cancelTestsButton.style.visibility = "hidden";

            }
        });
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

export async function cancelTests() {

    await stopTests();
}

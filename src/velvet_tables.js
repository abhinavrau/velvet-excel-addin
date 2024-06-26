


export async function createConfigTable() {
    await Excel.run(async (context) => {
        try {
            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            currentWorksheet.load("name");
            await context.sync();
            const worksheetName = currentWorksheet.name;
            console.log(`TableName: ${worksheetName}.ConfigTable`);

            var range = currentWorksheet.getRange('A1');
            range.values = [["Vertex AI Search Parameters"]];
            range.format.font.bold = true;
            range.format.fill.color = 'yellow';
            range.format.font.size = 16;

            var configTable = currentWorksheet.tables.add("A2:B2", true /*hasHeaders*/);
            configTable.name = `${worksheetName}.ConfigTable`;

            configTable.getHeaderRowRange().values =
                [["Config", "Value"]];

            configTable.rows.add(null, [
                ["Vertex AI Search Project Number", "384473000457"],
                ["Vertex AI Search DataStore Name", "alphabet-pdfs_1695783402380"],
                ["Vertex AI Project ID", "argolis-arau"],
                ["Vertex AI Location", "us-central1"],
                ["maxExtractiveAnswerCount (1-5)", "2"], //maxExtractiveAnswerCount
                ["maxExtractiveSegmentCount (1-5)", "0"], //maxExtractiveSegmentCount
                ["maxSnippetCount (1-5)", "0"], //maxSnippetCount
                ["Preamble (Customized Summaries)", ""],
                ["Summarization Model", "gemini-1.0-pro-001/answer_gen/v1"],    
                ["summaryResultCount (1-5)", "2"],   //summaryResultCount
                ["useSemanticChunks (True or False)", "False"],   //useSemanticChunks
                ["ignoreAdversarialQuery (True or False)", "True"], // ignoreAdversarialQuery
                ["ignoreNonSummarySeekingQuery (True or False)", "True"], // ignoreNonSummarySeekingQuery
                ["SummaryMatchingAdditionalPrompt", "If there are monetary numbers in the answers, they should be matched exactly."], // ignoreNonSummarySeekingQuery
                ["Batch Size (1-10)", "2"], // BatchSize
                ["Time between Batches in Seconds (1-10)", "2"], // BatchSize
            ]);

            currentWorksheet.getUsedRange().format.autofitColumns();
            currentWorksheet.getUsedRange().format.autofitRows();
           
            await context.sync();

        } catch (error) {
            console.error('Error createTable:' + error);
            throw error;
            //showStatus(`Exception when creating sample data: ${JSON.stringify(error)}`, true);
        }
    });

}

export async function createDataTable() {
    await Excel.run(async (context) => {
        try {
            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            currentWorksheet.load("name");
            await context.sync();
            const worksheetName = currentWorksheet.name;

            var velvetTable = currentWorksheet.tables.add("C17:O17", true /*hasHeaders*/);
            velvetTable.name = `${worksheetName}.TestCasesTable`;

            velvetTable.getHeaderRowRange().values =
                [["ID", "Query", "Expected Summary", "Actual Summary", "Expected Link 1", "Expected Link 2", "Expected Link 3", "Summary Match", "First Link Match", "Link in Top 2", "Actual Link 1", "Actual Link 2", "Actual Link 3"]];

            velvetTable.resize('C17:O118');
            currentWorksheet.getUsedRange().format.autofitColumns();
            currentWorksheet.getUsedRange().format.autofitRows();
            

            await context.sync();

            return
        } catch (error) {
            console.error('Error createTable:' + error);
            throw error;
            //showStatus(`Exception when creating sample data: ${JSON.stringify(error)}`, true);
        }
    });

}





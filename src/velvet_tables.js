
import { configValues, testCaseData } from './common.js';

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
                [configValues[0]];

            configTable.rows.add(null, configValues.slice(1));

            currentWorksheet.getUsedRange().format.autofitColumns();
            currentWorksheet.getUsedRange().format.autofitRows();
           
            await context.sync();

        } catch (error) {
            console.error('Error createTable:' + error);
            showStatus(`Exception when creating sample data: ${JSON.stringify(error)}`, true);
            throw error;
            
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
                [testCaseData[0]];

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





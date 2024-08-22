
import { synth_q_and_a_configValues, synth_q_and_a_TableHeader } from './common.js';
import { appendError, showStatus } from './ui.js';

export async function createSyntheticQAConfigTable() {
    await Excel.run(async (context) => {
        try {
            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            currentWorksheet.load("name");
            await context.sync();
            const worksheetName = currentWorksheet.name;
            

            var range = currentWorksheet.getRange('A1');
            range.values = [["Generate Synthetic Questions and Answers"]];
            range.format.font.bold = true;
            range.format.fill.color = 'yellow';
            range.format.font.size = 16;

            var configTable = currentWorksheet.tables.add("A2:B2", true /*hasHeaders*/);
            configTable.name = `${worksheetName}.ConfigTable`;

            configTable.getHeaderRowRange().values =
                [synth_q_and_a_configValues[0]];

            configTable.rows.add(null, synth_q_and_a_configValues.slice(1));
            
            currentWorksheet.getUsedRange().format.autofitColumns();
            currentWorksheet.getUsedRange().format.autofitRows();
            currentWorksheet.getRange("A6:B6").format.wrapText = true; // wrap system instrcutions
            currentWorksheet.getRange("A6:B6").format.shrinkToFit = true; // shrinkToFit system instrcutions
            
            await context.sync();
           

        } catch (error) {
            showStatus(`Exception when creating createSyntheticQA Config Table: ${error.message}`, true);
            appendError('Error creating Config Table:', error);
            
            return;
            
        }
    });

}

export async function createSyntheticQADataTable() {
    await Excel.run(async (context) => {
        try {
            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            currentWorksheet.load("name");
            await context.sync();
            const worksheetName = currentWorksheet.name;

            var velvetTable = currentWorksheet.tables.add("C9:J9", true /*hasHeaders*/);
            velvetTable.name = `${worksheetName}.TestCasesTable`;

            velvetTable.getHeaderRowRange().values =
                [synth_q_and_a_TableHeader[0]];

            velvetTable.resize('C9:J119');
            currentWorksheet.getUsedRange().format.autofitColumns();
            currentWorksheet.getUsedRange().format.autofitRows();
            currentWorksheet.getUsedRange().format.wrapText = true; 
            currentWorksheet.getUsedRange().format.shrinkToFit = true; 

            await context.sync();

        } catch (error) {
           
            showStatus(`Exception when creating createSyntheticQA DataTable: ${error.message}`, true);
            appendError('Error creating Data Table:', error);
            return;
            
        }
    });

}





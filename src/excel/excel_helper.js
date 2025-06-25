export async function findSheetsWithTableSuffix(tableSuffix) {
  const sheetNames = [];
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    for (let i = 0; i < sheets.items.length; i++) {
      const sheet = sheets.items[i];
      const tableName = `${sheet.name}.${tableSuffix}`;
      const table = sheet.tables.getItemOrNullObject(tableName);
      table.load("name");
      await context.sync();

      if (!table.isNullObject) {
        sheetNames.push(sheet.name);
      }
    }
  });
  return sheetNames;
}

import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
const path = "/Users/jacheflin/Downloads/Renewal_Management_Tool_RBP-3.xlsx";
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
for (const sheetName of ["SETUP","March 2026","April 2026","TEMPLATE"]) {
  const sheet = wb.worksheets.getItem(sheetName);
  const range = sheet.getRange("A1:AB18");
  const data = await wb.inspect({ kind: "table", sheetId: sheet.id, range: "A1:AB18", include: "values,formulas", tableMaxRows: 20, tableMaxCols: 28, tableMaxCellChars: 60 });
  console.log(`SHEET ${sheetName}`);
  console.log(data.ndjson);
}

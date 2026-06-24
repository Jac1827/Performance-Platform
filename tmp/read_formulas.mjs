import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load('/Users/jacheflin/Downloads/Renewal_Management_Tool_RBP-3.xlsx'));
for (const sheetName of ['March 2026','April 2026']) {
  const sheet = wb.worksheets.getItem(sheetName);
  const range = sheet.getRange('A1:AB8');
  console.log('SHEET', sheetName);
  console.log(JSON.stringify(range.formulas));
}

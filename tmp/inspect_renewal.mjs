import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
const path = "/Users/jacheflin/Downloads/Renewal_Management_Tool_RBP-3.xlsx";
const input = await FileBlob.load(path);
const wb = await SpreadsheetFile.importXlsx(input);
console.log(await wb.inspect({ kind: "workbook,sheet,table", maxChars: 12000, tableMaxRows: 4, tableMaxCols: 12, tableMaxCellChars: 40 }));

import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
const path = "/Users/jacheflin/Downloads/Renewal_Management_Tool_RBP-3.xlsx";
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
console.log(await wb.inspect({ kind: "sheet", include: "id,name" }));

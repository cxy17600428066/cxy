import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const src = "C:/Users/admin/Desktop/2026年5月绩效考核表-IT部.xlsx";
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(src));
const overview = await wb.inspect({ kind: "workbook,sheet,table", maxChars: 12000, tableMaxRows: 12, tableMaxCols: 12, tableMaxCellChars: 120 });
console.log(overview.ndjson);
const sheets = await wb.inspect({ kind: "sheet", include: "id,name", maxChars: 4000 });
console.log(sheets.ndjson);
for (let i = 0; i < wb.worksheets.items.length; i++) {
  const s = wb.worksheets.getItemAt(i);
  const used = s.getUsedRange();
  console.log(`SHEET ${s.name} USED ${used?.address ?? "none"}`);
  if (used) {
    const reg = await wb.inspect({ kind: "region", sheetId: s.name, range: used.address.split("!").pop(), maxChars: 14000, tableMaxRows: 80, tableMaxCols: 20, tableMaxCellChars: 160 });
    console.log(reg.ndjson);
    const png = await wb.render({ sheetName: s.name, autoCrop: "all", scale: 1, format: "png" });
    await fs.writeFile(`outputs/source-${i}.png`, new Uint8Array(await png.arrayBuffer()));
  }
}

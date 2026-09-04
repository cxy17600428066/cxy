import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
const path="./outputs/it-performance-templates/2026年5月绩效考核模板-产品测试研发RPA.xlsx";
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path));
for(const name of ["产品","测试","前端","后端","全栈","RPA"]){
  const p=await wb.render({sheetName:name,range:"A1:L18",scale:1,format:"png"});
  await p.save(`./outputs/it-performance-templates/preview-${name}.png`);
}
console.log("rendered");

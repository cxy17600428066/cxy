import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const sourcePath=path.resolve("outputs","requirement_pool_weekly","需求池模板_开发周任务.xlsx");
const outputDir=path.resolve("outputs","requirement_pool_weekly_report");
await fs.mkdir(outputDir,{recursive:true});
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
console.log((await wb.inspect({kind:"workbook,sheet,table",maxChars:4500,tableMaxRows:6,tableMaxCols:14})).ndjson);
if(process.env.PREVIEW_ONLY==="1"){
  const png=await wb.render({sheetName:"开发周任务",range:"A1:J14",scale:1,format:"png"});
  await fs.writeFile(path.join(outputDir,"before_开发周任务.png"),new Uint8Array(await png.arrayBuffer()));
  process.exit(0);
}

const sheet=wb.worksheets.add("周报填报");
const navy="#17324D",blue="#2F75B5",light="#DCE6F1";
sheet.showGridLines=false;
sheet.getRange("A1:N1").merge(); sheet.getRange("A1").values=[["开发人员周报填报"]];
sheet.getRange("A1:N1").format={fill:navy,font:{bold:true,color:"#FFFFFF",size:16},horizontalAlignment:"center",verticalAlignment:"center"};
sheet.getRange("A1:N1").format.rowHeight=32;
sheet.getRange("A2:N2").merge(); sheet.getRange("A2").values=[["填写周开始日期、版本号和开发人员后，任务数据自动关联；文字部分用于周报汇报。"]];
sheet.getRange("A2:N2").format={fill:light,font:{color:navy,italic:true},horizontalAlignment:"center"};
sheet.getRange("A4:N4").values=[["周开始日期","版本号","开发人员","任务总数","开发中","已上线","上线完成率","本周完成事项","本周未完成及原因","下周工作计划","风险与问题","需协调/支持事项","填报日期","备注"]];
sheet.getRange("A4:N4").format={fill:blue,font:{bold:true,color:"#FFFFFF"},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:"#B4C7E7"}};
sheet.getRange("A5:N104").format={borders:{preset:"all",style:"thin",color:"#E7E6E6"},verticalAlignment:"top"};
for(let r=5;r<=104;r++){
  sheet.getRange(`D${r}`).formulas=[[`=IF(OR(A${r}="",B${r}="",C${r}=""),"",COUNTIFS('需求明细'!$N$4:$N$503,A${r},'需求明细'!$M$4:$M$503,B${r},'需求明细'!$I$4:$I$503,C${r}))`]];
  sheet.getRange(`E${r}`).formulas=[[`=IF(D${r}="","",COUNTIFS('需求明细'!$N$4:$N$503,$A${r},'需求明细'!$M$4:$M$503,$B${r},'需求明细'!$I$4:$I$503,$C${r},'需求明细'!$H$4:$H$503,"开发中"))`]];
  sheet.getRange(`F${r}`).formulas=[[`=IF(D${r}="","",COUNTIFS('需求明细'!$N$4:$N$503,$A${r},'需求明细'!$M$4:$M$503,$B${r},'需求明细'!$I$4:$I$503,$C${r},'需求明细'!$H$4:$H$503,"已上线"))`]];
  sheet.getRange(`G${r}`).formulas=[[`=IF(D${r}="","",IF(D${r}=0,0,F${r}/D${r}))`]];
}
sheet.getRange("A5:A104").format.numberFormat="yyyy-mm-dd";
sheet.getRange("M5:M104").format.numberFormat="yyyy-mm-dd";
sheet.getRange("D5:F104").format={numberFormat:"#,##0",horizontalAlignment:"center"};
sheet.getRange("G5:G104").format={numberFormat:"0%",horizontalAlignment:"center"};
sheet.getRange("G5:G104").conditionalFormats.add("colorScale",{colors:["#F8696B","#FFEB84","#63BE7B"],thresholds:["min","50%","max"]});
sheet.getRange("H5:N104").format.wrapText=true;
sheet.getRange("A5:C104").format.fill="#FFF9E6";
sheet.getRange("H5:N104").format.fill="#FFF9E6";
sheet.tables.add("A4:N104",true,"WeeklyReportEntries");
const widths=[16,16,16,12,12,12,14,32,32,32,28,28,14,22];
for(let i=0;i<14;i++) sheet.getRangeByIndexes(3,i,101,1).format.columnWidth=widths[i];
sheet.getRange("A5:N104").format.rowHeight=44;
sheet.freezePanes.freezeRows(4);

console.log((await wb.inspect({kind:"table",range:"周报填报!A1:N10",include:"values,formulas",tableMaxRows:12,tableMaxCols:16})).ndjson);
console.log((await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!",options:{useRegex:true,maxResults:100},summary:"formula errors"})).ndjson);
const outputPath=path.join(outputDir,"需求池模板_含周报填报.xlsx");
await (await SpreadsheetFile.exportXlsx(wb)).save(outputPath); console.log(`OUTPUT=${outputPath}`);
for(const [name,range] of [["需求总览","A1:F15"],["需求明细","A1:N14"],["开发周任务","A1:J14"],["周报填报","A1:N14"]]){const png=await wb.render({sheetName:name,range,scale:1,format:"png"});await fs.writeFile(path.join(outputDir,`preview_${name}.png`),new Uint8Array(await png.arrayBuffer()));}

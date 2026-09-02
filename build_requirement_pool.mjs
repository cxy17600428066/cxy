import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const sourcePath=path.resolve("outputs","requirement_pool_merged","需求池模板_合并明细.xlsx");
const outputDir=path.resolve("outputs","requirement_pool_weekly");
await fs.mkdir(outputDir,{recursive:true});
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
const detail=wb.worksheets.getItem("需求明细");
const navy="#17324D",blue="#2F75B5",light="#DCE6F1";

// Extend the unified detail sheet with version and planned-week fields.
detail.getRange("A1:L1").unmerge(); detail.getRange("A1:N1").merge();
detail.getRange("A2:L2").unmerge(); detail.getRange("A2:N2").merge();
detail.getRange("A2").values=[["所有板块在此统一填写；所属板块、优先级、当前状态请使用下拉选项；版本号和计划周用于开发周任务统计。"]];
detail.getRange("M3:N3").values=[["版本号","计划周（周一）"]];
detail.getRange("M3:N3").format={fill:blue,font:{bold:true,color:"#FFFFFF"},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:"#B4C7E7"}};
detail.getRange("M4:N503").format={borders:{preset:"all",style:"thin",color:"#E7E6E6"},verticalAlignment:"center"};
detail.getRange("M3:M503").format.columnWidth=16;
detail.getRange("N3:N503").format.columnWidth=17;
detail.getRange("N4:N503").format.numberFormat="yyyy-mm-dd";

const summary=wb.worksheets.getItem("需求总览");
for(let r=5;r<=14;r++){
  summary.getRange(`B${r}`).formulas=[[`=COUNTIF('需求明细'!$B$4:$B$503,A${r})`]];
  for(let j=0;j<4;j++){const c=String.fromCharCode(67+j);summary.getRange(`${c}${r}`).formulas=[[`=COUNTIFS('需求明细'!$B$4:$B$503,$A${r},'需求明细'!$H$4:$H$503,${c}$4)`]];}
}
const weekly=wb.worksheets.add("开发周任务");
weekly.showGridLines=false;
weekly.getRange("A1:J1").merge(); weekly.getRange("A1").values=[["开发人员每周任务统计"]];
weekly.getRange("A1:J1").format={fill:navy,font:{bold:true,color:"#FFFFFF",size:16},horizontalAlignment:"center",verticalAlignment:"center"};
weekly.getRange("A1:J1").format.rowHeight=32;
weekly.getRange("A2:J2").merge(); weekly.getRange("A2").values=[["填写周开始日期（周一）、版本号和开发人员，其余指标将从“需求明细”自动统计。"]];
weekly.getRange("A2:J2").format={fill:light,font:{color:navy,italic:true},horizontalAlignment:"center"};
weekly.getRange("A4:J4").values=[["周开始日期","版本号","开发人员","任务总数","未开始","进行中","开发中","已上线","上线完成率","备注"]];
weekly.getRange("A4:J4").format={fill:blue,font:{bold:true,color:"#FFFFFF"},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:"#B4C7E7"}};
weekly.getRange("A5:J104").format={borders:{preset:"all",style:"thin",color:"#E7E6E6"},verticalAlignment:"center"};
weekly.getRange("A5:A104").format.numberFormat="yyyy-mm-dd";
for(let r=5;r<=104;r++){
  weekly.getRange(`D${r}`).formulas=[[`=IF(OR(A${r}="",B${r}="",C${r}=""),"",COUNTIFS('需求明细'!$N$4:$N$503,A${r},'需求明细'!$M$4:$M$503,B${r},'需求明细'!$I$4:$I$503,C${r}))`]];
  for(let j=0;j<4;j++){const c=String.fromCharCode(69+j);const statusCell=String.fromCharCode(69+j)+"$4";weekly.getRange(`${c}${r}`).formulas=[[`=IF(D${r}="","",COUNTIFS('需求明细'!$N$4:$N$503,$A${r},'需求明细'!$M$4:$M$503,$B${r},'需求明细'!$I$4:$I$503,$C${r},'需求明细'!$H$4:$H$503,${statusCell}))`]];}
  weekly.getRange(`I${r}`).formulas=[[`=IF(D${r}="","",IF(D${r}=0,0,H${r}/D${r}))`]];
}
weekly.getRange("D5:H104").format={numberFormat:"#,##0",horizontalAlignment:"center"};
weekly.getRange("I5:I104").format={numberFormat:"0%",horizontalAlignment:"center"};
weekly.getRange("H5:H104").conditionalFormats.add("dataBar",{color:"#70AD47",gradient:true});
weekly.getRange("I5:I104").conditionalFormats.add("colorScale",{colors:["#F8696B","#FFEB84","#63BE7B"],thresholds:["min","50%","max"]});
weekly.tables.add("A4:J104",true,"DeveloperWeeklyStats");
const widths=[16,16,16,12,12,12,12,12,14,28];
for(let i=0;i<10;i++) weekly.getRangeByIndexes(3,i,101,1).format.columnWidth=widths[i];
weekly.freezePanes.freezeRows(4);

console.log((await wb.inspect({kind:"table",range:"需求明细!A1:N8",include:"values,formulas",tableMaxRows:10,tableMaxCols:16})).ndjson);
console.log((await wb.inspect({kind:"table",range:"开发周任务!A1:J10",include:"values,formulas",tableMaxRows:12,tableMaxCols:12})).ndjson);
console.log((await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!",options:{useRegex:true,maxResults:100},summary:"formula errors"})).ndjson);

const outputPath=path.join(outputDir,"需求池模板_开发周任务.xlsx");
await (await SpreadsheetFile.exportXlsx(wb)).save(outputPath);
console.log(`OUTPUT=${outputPath}`);
for(const [name,range] of [["需求总览","A1:F15"],["需求明细","A1:N14"],["开发周任务","A1:J14"]]){const png=await wb.render({sheetName:name,range,scale:1,format:"png"});await fs.writeFile(path.join(outputDir,`preview_${name}.png`),new Uint8Array(await png.arrayBuffer()));}

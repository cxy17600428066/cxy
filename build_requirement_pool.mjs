import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const modules=["OMS","标准库","CRM","BI报表","基础设置","小程序","北森","勤策","PLM","旺店通"];
const statuses=["未开始","进行中","开发中","已上线"];
const sourcePath=path.resolve("outputs","requirement_pool_template","需求池模板.xlsx");
const outputDir=path.resolve("outputs","requirement_pool_merged");
await fs.mkdir(outputDir,{recursive:true});
const source=await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
console.log((await source.inspect({kind:"workbook,sheet,table",maxChars:4000,tableMaxRows:5,tableMaxCols:12})).ndjson);
if(process.env.PREVIEW_ONLY==="1"){
  for(const [name,range] of [["需求总览","A1:F15"],["OMS","A1:K14"]]){
    const png=await source.render({sheetName:name,range,scale:1,format:"png"});
    await fs.writeFile(path.join(outputDir,`before_${name}.png`),new Uint8Array(await png.arrayBuffer()));
  }
  process.exit(0);
}

const migrated=[];
for(const module of modules){
  for(const row of source.worksheets.getItem(module).getRange("B5:K204").values){
    if(row[0]!=null&&String(row[0]).trim()!=="") migrated.push([module,...row]);
  }
}

const wb=Workbook.create();
const summary=wb.worksheets.add("需求总览");
const detail=wb.worksheets.add("需求明细");
const navy="#17324D",blue="#2F75B5",light="#DCE6F1",pale="#F4F7FA";
summary.showGridLines=false;
summary.getRange("A1:F1").merge(); summary.getRange("A1").values=[["需求池总览"]];
summary.getRange("A1:F1").format={fill:navy,font:{bold:true,color:"#FFFFFF",size:18},horizontalAlignment:"center",verticalAlignment:"center"};
summary.getRange("A1:F1").format.rowHeight=34;
summary.getRange("A2:F2").merge(); summary.getRange("A2").values=[["各板块需求数量及状态自动汇总（请在“需求明细”工作表中统一维护）"]];
summary.getRange("A2:F2").format={fill:light,font:{color:navy,italic:true},horizontalAlignment:"center"};
summary.getRange("A4:F4").values=[["板块","需求总数",...statuses]];
summary.getRange("A4:F4").format={fill:blue,font:{bold:true,color:"#FFFFFF"},horizontalAlignment:"center",verticalAlignment:"center",borders:{preset:"all",style:"thin",color:"#B4C7E7"}};
for(let i=0;i<modules.length;i++){
  const r=i+5; summary.getRange(`A${r}`).values=[[modules[i]]];
  summary.getRange(`B${r}`).formulas=[[`=COUNTIF('需求明细'!$B$5:$B$504,A${r})`]];
  for(let j=0;j<4;j++){const c=String.fromCharCode(67+j);summary.getRange(`${c}${r}`).formulas=[[`=COUNTIFS('需求明细'!$B$5:$B$504,$A${r},'需求明细'!$H$5:$H$504,${c}$4)`]];}
}
summary.getRange("A15").values=[["合计"]];
for(const c of ["B","C","D","E","F"]) summary.getRange(`${c}15`).formulas=[[`=SUM(${c}5:${c}14)`]];
summary.getRange("A5:F14").format={borders:{preset:"all",style:"thin",color:"#D9E2F3"},verticalAlignment:"center"};
summary.getRange("A5:A14").format.fill=pale;
summary.getRange("B5:F15").format={numberFormat:"#,##0",horizontalAlignment:"center"};
summary.getRange("A15:F15").format={fill:navy,font:{bold:true,color:"#FFFFFF"},borders:{preset:"all",style:"thin",color:navy}};
summary.getRange("A4:A15").format.columnWidth=16; summary.getRange("B4:F15").format.columnWidth=13; summary.freezePanes.freezeRows(4);

detail.showGridLines=false;
detail.getRange("A1:L1").merge(); detail.getRange("A1").values=[["统一需求明细"]];
detail.getRange("A1:L1").format={fill:navy,font:{bold:true,color:"#FFFFFF",size:16},horizontalAlignment:"center",verticalAlignment:"center"};
detail.getRange("A1:L1").format.rowHeight=32;
detail.getRange("A2:L2").merge(); detail.getRange("A2").values=[["所有板块在此统一填写；所属板块、优先级、当前状态请使用下拉选项，总览将自动汇总。"]];
detail.getRange("A2:L2").format={fill:light,font:{color:navy,italic:true},wrapText:true};
detail.getRange("A4:L4").values=[["需求编号","所属板块","需求名称","需求描述","提出人","提出日期","优先级","当前状态","负责人","计划上线日期","实际上线日期","备注"]];
detail.getRange("A4:L4").format={fill:blue,font:{bold:true,color:"#FFFFFF"},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:"#B4C7E7"}};
detail.getRange("A5:L504").format={borders:{preset:"all",style:"thin",color:"#E7E6E6"},verticalAlignment:"center"};
detail.getRange("A5:A504").formulas=Array.from({length:500},(_,i)=>[`=IF(C${i+5}="","",ROW()-4)`]);
if(migrated.length) detail.getRangeByIndexes(4,1,migrated.length,11).values=migrated;
detail.getRange("B5:B504").dataValidation={rule:{type:"list",values:modules}};
detail.getRange("G5:G504").dataValidation={rule:{type:"list",values:["高","中","低"]}};
detail.getRange("H5:H504").dataValidation={rule:{type:"list",values:statuses}};
const colors=["#A5A5A5","#ED7D31","#8064A2","#70AD47"];
for(let i=0;i<4;i++) detail.getRange("H5:H504").conditionalFormats.add("containsText",{text:statuses[i],format:{fill:colors[i],font:{color:"#FFFFFF",bold:true}}});
for(const [text,fill,color] of [["高","#F4CCCC","#9C0006"],["中","#FFF2CC","#9C6500"],["低","#D9EAD3","#274E13"]]) detail.getRange("G5:G504").conditionalFormats.add("containsText",{text,format:{fill,font:{color,bold:text==="高"}}});
detail.getRange("F5:F504").format.numberFormat="yyyy-mm-dd"; detail.getRange("J5:K504").format.numberFormat="yyyy-mm-dd";
detail.tables.add("A4:L504",true,"RequirementDetails");
const widths=[11,14,24,42,13,14,12,12,13,15,15,28];
for(let i=0;i<12;i++) detail.getRangeByIndexes(3,i,501,1).format.columnWidth=widths[i];
detail.getRange("C5:D504").format.wrapText=true; detail.getRange("L5:L504").format.wrapText=true; detail.freezePanes.freezeRows(4);

console.log((await wb.inspect({kind:"table",range:"需求总览!A1:F15",include:"values,formulas",tableMaxRows:20,tableMaxCols:8})).ndjson);
console.log((await wb.inspect({kind:"table",range:"需求明细!A1:L10",include:"values,formulas",tableMaxRows:12,tableMaxCols:14})).ndjson);
console.log((await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:100},summary:"formula errors"})).ndjson);
const outputPath=path.join(outputDir,"需求池模板_合并明细.xlsx");
await (await SpreadsheetFile.exportXlsx(wb)).save(outputPath); console.log(`OUTPUT=${outputPath}`);
for(const [name,range] of [["需求总览","A1:F15"],["需求明细","A1:L14"]]){const png=await wb.render({sheetName:name,range,scale:1,format:"png"});await fs.writeFile(path.join(outputDir,`after_${name}.png`),new Uint8Array(await png.arrayBuffer()));}

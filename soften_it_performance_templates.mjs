import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const path = "./outputs/it-performance-templates/2026年5月绩效考核模板-产品测试研发RPA.xlsx";
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const sheets = ["产品","测试","前端","后端","全栈","RPA"];
const standards = [
  "完成情况突出，质量和及时性较好，可评90–100分；整体达到要求，可评75–89分；部分未达到预期，可评60–74分；差距较大，可评60分以下。",
  "工作成果质量较好、问题较少，可评90–100分；存在少量问题但不影响整体交付，可评75–89分；问题较多或需要返工，可评60–74分；对交付造成明显影响，可评60分以下。",
  "推进顺畅、关键事项闭环较好，可评90–100分；基本按要求完成，可评75–89分；推进或闭环存在不足，可评60–74分；多次影响项目进展，可评60分以下。",
  "实际效果较好并对业务或项目有明显帮助，可评90–100分；达到基本预期，可评75–89分；效果一般，可评60–74分；未体现明显效果，可评60分以下。",
  "执行规范、成果完整、可维护性较好，可评90–100分；基本符合要求，可评75–89分；存在一定缺失或改进空间，可评60–74分；问题较明显，可评60分以下。",
  "主动性、责任心和协作表现较好，可评90–100分；日常配合基本到位，可评75–89分；响应或配合存在不足，可评60–74分；多次影响协作，可评60分以下。"
];

for (const name of sheets) {
  const s = wb.worksheets.getItem(name);
  s.getRange("F5:F10").values = Array.from({length:6},()=>["结合当月工作安排"]);
  s.getRange("G5:G10").values = standards.map(x=>[x]);
  s.getRange("I5:I10").values = Array.from({length:6},()=>["结合实际完成情况、工作难度、质量、影响及相关反馈综合评分"]);
  s.getRange("G12").values = [["因个人责任造成较大业务、系统、数据或安全影响"]];
  s.getRange("I12").values = [["视影响范围和责任程度建议扣10–50分"]];
  s.getRange("G13").values = [["存在发布、评审、测试、权限或操作流程方面的不规范情况"]];
  s.getRange("I13").values = [["视情节和影响程度建议扣3–20分"]];
  s.getRange("G14").values = [["存在工作逾期、沟通不及时、重复问题或记录不完整等情况"]];
  s.getRange("I14").values = [["视发生频次和实际影响建议扣1–10分"]];
}

const check = await wb.inspect({kind:"table",sheetId:"产品",range:"D4:L15",include:"values,formulas",tableMaxRows:20,tableMaxCols:12,maxChars:12000});
console.log(check.ndjson);
const errors = await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",options:{useRegex:true,maxResults:300},summary:"final formula error scan"});
console.log(errors.ndjson);
const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(path);
console.log(path);

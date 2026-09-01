import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const modules = ["OMS", "标准库", "CRM", "BI报表", "基础设置", "小程序", "北森", "勤策", "PLM", "旺店通"];
const statuses = ["未开始", "进行中", "开发中", "已上线"];
const outputDir = path.resolve("outputs", "requirement_pool_template");
await fs.mkdir(outputDir, { recursive: true });

const wb = Workbook.create();
const summary = wb.worksheets.add("需求总览");
for (const name of modules) wb.worksheets.add(name);

const navy = "#17324D";
const blue = "#2F75B5";
const lightBlue = "#DCE6F1";
const pale = "#F4F7FA";
const green = "#70AD47";
const orange = "#ED7D31";
const purple = "#8064A2";
const gray = "#A5A5A5";
const statusColors = [gray, orange, purple, green];

summary.showGridLines = false;
summary.getRange("A1:F1").merge();
summary.getRange("A1").values = [["需求池总览"]];
summary.getRange("A1:F1").format = {
  fill: navy,
  font: { bold: true, color: "#FFFFFF", size: 18 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
summary.getRange("A1:F1").format.rowHeight = 34;
summary.getRange("A2:F2").merge();
summary.getRange("A2").values = [["各板块需求数量及状态自动汇总（请在各板块工作表中维护需求）"]];
summary.getRange("A2:F2").format = { fill: lightBlue, font: { color: navy, italic: true }, horizontalAlignment: "center" };

summary.getRange("A4:F4").values = [["板块", "需求总数", ...statuses]];
summary.getRange("A4:F4").format = {
  fill: blue,
  font: { bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: "#B4C7E7" },
};

for (let i = 0; i < modules.length; i++) {
  const row = 5 + i;
  const name = modules[i];
  summary.getRange(`A${row}`).values = [[name]];
  summary.getRange(`B${row}`).formulas = [[`=COUNTA('${name}'!$B$5:$B$204)`]];
  for (let j = 0; j < statuses.length; j++) {
    const col = String.fromCharCode(67 + j);
    summary.getRange(`${col}${row}`).formulas = [[`=COUNTIF('${name}'!$G$5:$G$204,${col}$4)`]];
  }
}
summary.getRange("A15").values = [["合计"]];
summary.getRange("B15").formulas = [["=SUM(B5:B14)"]];
summary.getRange("C15").formulas = [["=SUM(C5:C14)"]];
summary.getRange("D15").formulas = [["=SUM(D5:D14)"]];
summary.getRange("E15").formulas = [["=SUM(E5:E14)"]];
summary.getRange("F15").formulas = [["=SUM(F5:F14)"]];
summary.getRange("A5:F14").format = {
  borders: { preset: "all", style: "thin", color: "#D9E2F3" },
  verticalAlignment: "center",
};
summary.getRange("A5:A14").format.fill = pale;
summary.getRange("B5:F15").format.numberFormat = "#,##0";
summary.getRange("B5:F15").format.horizontalAlignment = "center";
summary.getRange("A15:F15").format = {
  fill: navy,
  font: { bold: true, color: "#FFFFFF" },
  borders: { preset: "all", style: "thin", color: navy },
};
summary.getRange("A4:A15").format.columnWidth = 16;
summary.getRange("B4:F15").format.columnWidth = 13;
summary.freezePanes.freezeRows(4);

const headers = ["需求编号", "需求名称", "需求描述", "提出人", "提出日期", "优先级", "当前状态", "负责人", "计划上线日期", "实际上线日期", "备注"];
for (const name of modules) {
  const sh = wb.worksheets.getItem(name);
  sh.showGridLines = false;
  sh.getRange("A1:K1").merge();
  sh.getRange("A1").values = [[`${name}需求池`]];
  sh.getRange("A1:K1").format = {
    fill: navy,
    font: { bold: true, color: "#FFFFFF", size: 16 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  sh.getRange("A1:K1").format.rowHeight = 32;
  sh.getRange("A2:K2").merge();
  sh.getRange("A2").values = [["填写说明：每条需求占一行；需求名称用于总览计数；状态和优先级请使用下拉选项。"]];
  sh.getRange("A2:K2").format = { fill: lightBlue, font: { color: navy, italic: true }, wrapText: true };
  sh.getRange("A4:K4").values = [headers];
  sh.getRange("A4:K4").format = {
    fill: blue,
    font: { bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#B4C7E7" },
  };
  sh.getRange("A5:K204").format = {
    borders: { preset: "all", style: "thin", color: "#E7E6E6" },
    verticalAlignment: "center",
  };
  sh.getRange("A5:A204").formulas = Array.from({ length: 200 }, (_, i) => [`=IF(B${i + 5}="","",ROW()-4)`]);
  sh.getRange("E5:E204").format.numberFormat = "yyyy-mm-dd";
  sh.getRange("I5:J204").format.numberFormat = "yyyy-mm-dd";
  sh.getRange("F5:F204").dataValidation = { rule: { type: "list", values: ["高", "中", "低"] } };
  sh.getRange("G5:G204").dataValidation = { rule: { type: "list", values: statuses } };
  for (let j = 0; j < statuses.length; j++) {
    sh.getRange("G5:G204").conditionalFormats.add("containsText", {
      text: statuses[j], format: { fill: statusColors[j], font: { color: "#FFFFFF", bold: true } },
    });
  }
  sh.getRange("F5:F204").conditionalFormats.add("containsText", { text: "高", format: { fill: "#F4CCCC", font: { color: "#9C0006", bold: true } } });
  sh.getRange("F5:F204").conditionalFormats.add("containsText", { text: "中", format: { fill: "#FFF2CC", font: { color: "#9C6500" } } });
  sh.getRange("F5:F204").conditionalFormats.add("containsText", { text: "低", format: { fill: "#D9EAD3", font: { color: "#274E13" } } });
  sh.tables.add("A4:K204", true, `Req_${modules.indexOf(name) + 1}`);
  sh.getRange("A4:A204").format.columnWidth = 11;
  sh.getRange("B4:B204").format.columnWidth = 24;
  sh.getRange("C4:C204").format.columnWidth = 42;
  sh.getRange("D4:D204").format.columnWidth = 13;
  sh.getRange("E4:E204").format.columnWidth = 14;
  sh.getRange("F4:G204").format.columnWidth = 12;
  sh.getRange("H4:H204").format.columnWidth = 13;
  sh.getRange("I4:J204").format.columnWidth = 15;
  sh.getRange("K4:K204").format.columnWidth = 28;
  sh.getRange("B5:C204").format.wrapText = true;
  sh.getRange("K5:K204").format.wrapText = true;
  sh.freezePanes.freezeRows(4);
}

const inspect = await wb.inspect({ kind: "table", range: "需求总览!A1:F15", include: "values,formulas", tableMaxRows: 20, tableMaxCols: 8 });
console.log(inspect.ndjson);
const errors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula errors" });
console.log(errors.ndjson);

for (const name of ["需求总览", ...modules]) {
  const blob = await wb.render({ sheetName: name, range: name === "需求总览" ? "A1:F15" : "A1:K14", scale: 1, format: "png" });
  await fs.writeFile(path.join(outputDir, `preview_${name}.png`), new Uint8Array(await blob.arrayBuffer()));
}

const file = await SpreadsheetFile.exportXlsx(wb);
const outputPath = path.join(outputDir, "需求池模板.xlsx");
await file.save(outputPath);
console.log(`OUTPUT=${outputPath}`);

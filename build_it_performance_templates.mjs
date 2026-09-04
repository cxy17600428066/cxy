import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const outDir = "C:/Users/admin/Documents/New project/outputs/it-performance-templates";
const outFile = `${outDir}/2026年5月绩效考核模板-产品测试研发RPA.xlsx`;
const wb = Workbook.create();

const roles = {
  "产品": [
    ["需求交付",25,"按计划完成率≥100%得100分；95%–99%得90分；90%–94%得80分；80%–89%得60分；<80%得0分","按期完成需求调研、PRD、原型及评审；以批准计划为准","按计划完成项÷计划项×100%","产品/业务"],
    ["需求质量",20,"上线后因需求遗漏或逻辑错误返工0次得100分；1次得85分；2次得70分；3次得50分；≥4次得0分","统计由需求定义问题造成的返工，不含开发实现缺陷","按月返工次数分档","产品/研发/测试"],
    ["项目推进",20,"里程碑准时率≥100%得100分；95%–99%得90分；90%–94%得80分；80%–89%得60分；<80%得0分","跟进评审、排期、验收和上线，及时处理阻塞","准时完成里程碑÷应完成里程碑×100%","项目负责人"],
    ["业务价值",15,"目标指标达成率≥100%得100分；90%–99%得85分；80%–89%得70分；60%–79%得50分；<60%得0分","采用立项时约定的转化、效率、收入或成本指标","实际值÷目标值×100%（逆向指标按目标值÷实际值）","业务/数据"],
    ["文档与数据",10,"文档完整、版本清晰、数据准确得100分；每发现1处重要缺失或错误扣20分，最低0分","PRD、原型、埋点、验收记录可追溯","100-重要缺失或错误数×20","产品/测试"],
    ["协作与主动性",10,"主动识别并闭环问题，协作反馈及时得90–100分；基本履职得70–89分；多次催办或影响协作得0–69分","由直接主管结合业务方、研发和测试反馈评分","主管评分","直接主管"]
  ],
  "测试": [
    ["测试任务交付",25,"按期完成率≥100%得100分；95%–99%得90分；90%–94%得80分；80%–89%得60分；<80%得0分","按计划完成测试方案、用例、执行和报告","按期完成任务÷计划任务×100%","测试负责人"],
    ["缺陷发现质量",20,"上线后严重漏测0个得100分；1个一般漏测得85分；每增加1个再扣20分；出现1个严重/致命漏测本项最高40分","漏测须经产品、研发、测试共同复盘确认","按漏测数量及等级分档","产品/研发/测试"],
    ["缺陷闭环效率",15,"缺陷按期关闭率≥98%得100分；95%–97%得90分；90%–94%得75分；80%–89%得60分；<80%得0分","跟踪缺陷修复、回归和关闭","按期关闭缺陷÷应关闭缺陷×100%","缺陷平台"],
    ["用例覆盖",15,"需求覆盖率≥98%得100分；95%–97%得90分；90%–94%得75分；<90%得50分","核心流程、异常分支、兼容和权限场景覆盖","已覆盖需求点÷全部需求点×100%","测试负责人"],
    ["质量改进",15,"完成约定的自动化、专项或流程改进目标得100分；完成80%–99%得80分；完成60%–79%得60分；<60%得0分","包括自动化用例、性能、安全、质量复盘等","完成项÷计划项×100%","测试负责人"],
    ["协作与责任心",10,"响应及时、风险前置、问题闭环得90–100分；基本履职得70–89分；多次延误或遗漏得0–69分","由直接主管结合项目成员反馈评分","主管评分","直接主管"]
  ],
  "前端": [
    ["迭代交付",25,"按期完成率≥100%得100分；95%–99%得90分；90%–94%得80分；80%–89%得60分；<80%得0分","按排期完成开发、联调、自测和发布","按期完成任务÷计划任务×100%","研发负责人"],
    ["交付质量",20,"生产环境前端责任严重缺陷0个且一般缺陷≤2个得100分；一般缺陷3–4个得80分；5–6个得60分；>6个得0分；严重缺陷每个扣40分","以缺陷归因结果为准","按缺陷等级和数量分档","测试/运维"],
    ["还原与体验",15,"UI验收一次通过率≥95%得100分；90%–94%得85分；80%–89%得70分；<80%得40分","页面还原、交互、响应式和兼容性符合设计规范","一次验收通过页面÷验收页面×100%","产品/设计"],
    ["性能与稳定性",15,"约定性能指标全部达标且无前端事故得100分；1项未达标得75分；2项得50分；出现责任事故本项0分","关注首屏、资源体积、错误率和兼容性","按目标达成项及事故分档","监控平台/运维"],
    ["代码质量",15,"评审问题及时整改、核心代码有测试、无重复低级问题得90–100分；存在1–2项不足得70–89分；多次重复问题得0–69分","代码规范、可维护性、测试和评审质量","评审记录综合评分","技术负责人"],
    ["协作与改进",10,"联调响应及时并完成至少1项组件/工程化改进得100分；仅完成日常协作得80分；多次延误得0–60分","跨端协作及公共能力建设","主管评分","直接主管"]
  ],
  "后端": [
    ["迭代交付",25,"按期完成率≥100%得100分；95%–99%得90分；90%–94%得80分；80%–89%得60分；<80%得0分","按排期完成设计、开发、联调、自测和发布","按期完成任务÷计划任务×100%","研发负责人"],
    ["交付质量",20,"生产环境后端责任严重缺陷0个且一般缺陷≤2个得100分；一般缺陷3–4个得80分；5–6个得60分；>6个得0分；严重缺陷每个扣40分","以缺陷归因和线上问题记录为准","按缺陷等级和数量分档","测试/运维"],
    ["接口与数据质量",15,"接口验收一次通过率≥98%且无数据差错得100分；95%–97%得85分；90%–94%得70分；<90%得40分；数据事故本项0分","接口规范、幂等、数据一致性和异常处理","一次通过接口÷验收接口×100%","前端/测试/数据"],
    ["性能与稳定性",15,"SLA及约定性能指标全部达标得100分；1项未达标得75分；2项得50分；责任事故本项0分","关注可用性、响应时间、错误率、容量和告警","按目标达成项及事故分档","监控平台/运维"],
    ["代码与架构质量",15,"评审问题及时整改、核心逻辑有测试、无重复低级问题得90–100分；存在1–2项不足得70–89分；多次重复问题得0–69分","代码规范、测试覆盖、可维护性和安全性","评审记录综合评分","技术负责人"],
    ["协作与改进",10,"主动支持联调并完成至少1项性能/架构/工具改进得100分；仅完成日常协作得80分；多次延误得0–60分","跨团队协作及技术改进","主管评分","直接主管"]
  ],
  "全栈": [
    ["端到端交付",25,"按期完成率≥100%得100分；95%–99%得90分；90%–94%得80分；80%–89%得60分；<80%得0分","独立或主导完成前后端设计、开发、联调、自测和发布","按期完成任务÷计划任务×100%","研发负责人"],
    ["交付质量",20,"生产环境责任严重缺陷0个且一般缺陷≤2个得100分；一般缺陷3–4个得80分；5–6个得60分；>6个得0分；严重缺陷每个扣40分","按前后端综合归因统计","按缺陷等级和数量分档","测试/运维"],
    ["方案与一致性",15,"技术方案评审一次通过且接口、数据、交互一致得100分；返工1次得80分；2次得60分；≥3次得0分","覆盖前端、后端、数据库、异常和兼容方案","按评审返工次数分档","技术负责人"],
    ["性能与稳定性",15,"约定前后端性能及SLA全部达标得100分；1项未达标得75分；2项得50分；责任事故本项0分","关注页面、接口、数据库、错误率和告警","按目标达成项及事故分档","监控平台/运维"],
    ["代码与测试",15,"核心逻辑有测试、评审问题及时整改、无重复低级问题得90–100分；存在1–2项不足得70–89分；多次重复问题得0–69分","前后端代码规范、自动化测试和可维护性","评审记录综合评分","技术负责人"],
    ["协作与改进",10,"跨角色沟通顺畅并完成至少1项公共能力改进得100分；仅完成日常协作得80分；多次延误得0–60分","产品、测试、运维协作及技术沉淀","主管评分","直接主管"]
  ],
  "RPA": [
    ["流程交付",25,"按期上线率≥100%得100分；95%–99%得90分；90%–94%得80分；80%–89%得60分；<80%得0分","按计划完成流程分析、开发、测试、部署和文档","按期上线流程÷计划上线流程×100%","项目负责人"],
    ["运行成功率",20,"成功率≥99.5%得100分；99.0%–99.49%得90分；98.0%–98.99%得75分；95.0%–97.99%得50分；<95%得0分","剔除已确认的外部系统不可用和业务主动中止","成功运行次数÷有效运行次数×100%","RPA平台/运维"],
    ["异常响应与恢复",15,"P1/P2异常均在SLA内响应并恢复得100分；每超时1次扣25分；造成业务中断≥4小时本项0分","按约定SLA统计告警确认、定位和恢复","100-超时次数×25，最低0分","运维/业务"],
    ["自动化效益",15,"节省工时或处理量目标达成率≥100%得100分；90%–99%得85分；80%–89%得70分；<80%得40分","以立项基线和业务确认结果为准","实际效益÷目标效益×100%","业务/财务"],
    ["规范与可维护性",15,"代码、配置、凭据、日志、版本及操作手册全部合规得100分；每缺1项扣20分，最低0分","流程可监控、可回滚、可交接，不明文保存敏感信息","100-不合规项数×20","技术负责人/安全"],
    ["协作与持续改进",10,"需求澄清及时并完成至少1项组件复用或稳定性改进得100分；仅完成日常协作得80分；多次延误得0–60分","业务协作、组件沉淀及流程优化","主管评分","直接主管"]
  ]
};

const deductions = [
  ["重大事故","因本人责任造成重大生产事故、严重数据错误、敏感信息泄露或核心业务中断","每次扣20–50分；情节严重可按公司制度另行处理","事故复盘/信息安全记录"],
  ["流程违规","未经审批上线、绕过评审/测试、违规操作生产数据或账号凭据","每次扣5–20分","发布记录/审计记录"],
  ["工作纪律与重复问题","无故逾期且未提前预警；同类问题复盘后再次发生；关键记录或文档缺失","每次扣2–10分","项目记录/主管确认"]
];

function styleSheet(s, role, rows) {
  s.showGridLines = false;
  s.mergeCells("A1:L1"); s.getRange("A1").values = [["2026年 5 月度绩效考核表"]];
  s.mergeCells("A2:C2"); s.mergeCells("E2:F2"); s.mergeCells("G2:I2"); s.mergeCells("J2:L2");
  s.getRange("A2:L2").values = [["单位/部门：IT部",null,null,null,"负责人：",null,"考核周期：2026.5.1-2026.5.31",null,null,"签订日期：",null,null]];
  s.mergeCells("A3:D3"); s.mergeCells("E3:F3"); s.mergeCells("H3:I3"); s.mergeCells("J3:L3");
  s.getRange("A3:L3").values = [["绩 效 计 划 确 认",null,null,null,`被考核岗位：${role}`,null,"被考核人：","直接主管：",null,"间接主管：",null,null]];
  s.getRange("A4:L4").values = [["类别","序号","指标类型","指标名称","权重","目标值","评分标准（0–100分）","指标定义","核算公式","数据提供","实际绩效/得分","折算分"]];
  const vals = rows.map((r,i)=>[i===0?"KPI\n关键业绩指标":"",i+1,"岗位指标",r[0],r[1]/100,"当月目标",r[2],r[3],r[4],r[5],null,null]);
  s.getRange("A5:L10").values = vals;
  for(let row=5;row<=10;row++) s.getRange(`L${row}`).formulas=[[`=IF(K${row}="","",ROUND(K${row}*E${row}/100,2))`]];
  s.getRange("A11:L11").values=[["扣分项",null,null,"扣分事由",null,null,"扣分条件",null,"扣分规则","数据依据","扣分值","扣分"]];
  deductions.forEach((d,i)=>{ const row=12+i; s.getRange(`A${row}:L${row}`).values=[[i===0?"扣分项":"",i+1,"专项扣分",d[0],null,null,d[1],null,d[2],d[3],null,null]]; s.getRange(`L${row}`).formulas=[[`=IF(K${row}="","",-ABS(K${row}))`]]; });
  s.mergeCells("A15:J15"); s.getRange("A15").values=[["绩效总分（满分100分；扣分后最低0分）："]]; s.getRange("K15:L15").merge(); s.getRange("K15").formulas=[["=MAX(0,SUM(L5:L10)+SUM(L12:L14))"]];
  s.mergeCells("A16:F16"); s.mergeCells("G16:I16"); s.mergeCells("J16:L16"); s.getRange("A16:L16").values=[["被考核人意见：",null,null,null,null,null,"被考核人签字：",null,null,"年    月    日",null,null]];
  s.mergeCells("A17:F17"); s.mergeCells("G17:I17"); s.mergeCells("J17:L17"); s.getRange("A17:L17").values=[["考核人意见：",null,null,null,null,null,"考核人签字：",null,null,"年    月    日",null,null]];
  s.mergeCells("A18:F18"); s.mergeCells("G18:I18"); s.mergeCells("J18:L18"); s.getRange("A18:L18").values=[["间接主管意见：",null,null,null,null,null,"间接主管签字：",null,null,"年    月    日",null,null]];
  const all=s.getRange("A1:L18"); all.format.font={name:"Arial",size:10,color:"#222222"}; all.format.verticalAlignment="center"; all.format.wrapText=true;
  s.getRange("A1:L1").format.font={name:"Arial",size:16,bold:true,color:"#1F1F1F"}; s.getRange("A1:L1").format.horizontalAlignment="center";
  s.getRange("A3:L3").format.fill="#D9EAF7"; s.getRange("A3:L3").format.font={name:"Arial",size:10,bold:true,color:"#1F1F1F"};
  s.getRange("A4:L4").format.fill="#2F5597"; s.getRange("A4:L4").format.font={name:"Arial",size:10,bold:true,color:"#FFFFFF"}; s.getRange("A4:L4").format.horizontalAlignment="center";
  s.getRange("A11:L11").format.fill="#C65911"; s.getRange("A11:L11").format.font={name:"Arial",size:10,bold:true,color:"#FFFFFF"}; s.getRange("A11:L11").format.horizontalAlignment="center";
  s.getRange("A15:L15").format.fill="#FFF2CC"; s.getRange("A15:L15").format.font={name:"Arial",size:11,bold:true,color:"#1F1F1F"};
  s.getRange("A4:L18").format.borders={preset:"all",style:"thin",color:"#A6A6A6"};
  s.getRange("E5:E10").format.numberFormat="0%"; s.getRange("K5:K10").format.fill="#FFF2CC"; s.getRange("K12:K14").format.fill="#FCE4D6"; s.getRange("L5:L15").format.numberFormat="0.00";
  s.getRange("K5:K10").dataValidation={rule:{type:"whole",operator:"between",formula1:0,formula2:100}};
  s.getRange("K12:K14").dataValidation={rule:{type:"whole",operator:"between",formula1:0,formula2:100}};
  s.getRange("A5:A10").merge(); s.getRange("A12:A14").merge();
  s.getRange("A5:A15").format.horizontalAlignment="center"; s.getRange("B4:F15").format.horizontalAlignment="center"; s.getRange("J4:L15").format.horizontalAlignment="center";
  const widths={A:12,B:7,C:11,D:18,E:8,F:11,G:48,H:34,I:25,J:16,K:14,L:11}; for(const [c,w] of Object.entries(widths)) s.getRange(`${c}:${c}`).format.columnWidth=w;
  s.getRange("1:1").format.rowHeight=30; s.getRange("2:3").format.rowHeight=24; s.getRange("4:4").format.rowHeight=34; s.getRange("5:10").format.rowHeight=72; s.getRange("11:11").format.rowHeight=28; s.getRange("12:14").format.rowHeight=58; s.getRange("15:15").format.rowHeight=30; s.getRange("16:18").format.rowHeight=44;
  s.freezePanes.freezeRows(4);
}

for (const [role, rows] of Object.entries(roles)) styleSheet(wb.worksheets.add(role), role === "RPA" ? "RPA工程师" : `${role}工程师`, rows);
await fs.mkdir(outDir,{recursive:true});
const inspect=await wb.inspect({kind:"table",sheetId:"产品",range:"A1:L18",include:"values,formulas",tableMaxRows:20,tableMaxCols:12,maxChars:16000}); console.log(inspect.ndjson);
const errors=await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",options:{useRegex:true,maxResults:300},summary:"final formula error scan"}); console.log(errors.ndjson);
for (const role of Object.keys(roles)) { const p=await wb.render({sheetName:role,range:"A1:L18",scale:1,format:"png"}); await fs.writeFile(`${outDir}/preview-${role}.png`,new Uint8Array(await p.arrayBuffer())); }
const output=await SpreadsheetFile.exportXlsx(wb); await output.save(outFile); console.log(outFile);

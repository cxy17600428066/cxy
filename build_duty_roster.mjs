import fs from 'node:fs/promises';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const workbook = Workbook.create();
const sheet = workbook.worksheets.add('值班表');
sheet.showGridLines = false;

sheet.mergeCells('A1:I1');
sheet.getRange('A1').values = [['值 班 表']];
sheet.mergeCells('A2:I2');
sheet.getRange('A2').values = [['值班电话：13800000001']];

sheet.mergeCells('A3:A4');
sheet.mergeCells('B3:C3');
sheet.mergeCells('D3:E3');
sheet.mergeCells('F3:G3');
sheet.mergeCells('H3:H4');
sheet.mergeCells('I3:I4');
sheet.getRange('A3:I4').values = [
  ['日期','白班负责人',null,'夜班负责人',null,'领导带班',null,'值班人数','备注'],
  [null,'8:00-18:00',null,'19:00-06:00',null,'8:00-18:00',null,null,null],
];

const rows = [];
for (let d = 1; d <= 8; d++) rows.push([`20XX/1/${d}`,'稻小亮','13800000001','稻小亮','13800000001','稻小亮','13800000001',3,'']);
for (let i = 0; i < 4; i++) rows.push(['','','','','','','','','']);
sheet.getRange('A5:I16').values = rows;

const all = sheet.getRange('A1:I16');
all.format.font = { name: 'Microsoft YaHei', size: 10, color: '#000000' };
all.format.verticalAlignment = 'center';
all.format.borders = { preset: 'all', style: 'thin', color: '#555555' };
sheet.getRange('A1:I1').format = { font: { name: 'Microsoft YaHei', size: 22, bold: true }, horizontalAlignment: 'center', verticalAlignment: 'center' };
sheet.getRange('A2:I2').format = { font: { name: 'Microsoft YaHei', size: 10, bold: true }, horizontalAlignment: 'left', verticalAlignment: 'center', borders: { preset: 'outside', style: 'thin', color: '#555555' } };
sheet.getRange('A3:I4').format = { font: { name: 'Microsoft YaHei', size: 10, bold: true }, horizontalAlignment: 'center', verticalAlignment: 'center', wrapText: true, borders: { preset: 'all', style: 'thin', color: '#555555' } };
sheet.getRange('A5:I16').format.horizontalAlignment = 'center';
sheet.getRange('B5:B16').format.horizontalAlignment = 'center';
sheet.getRange('D5:D16').format.horizontalAlignment = 'center';
sheet.getRange('F5:F16').format.horizontalAlignment = 'center';
sheet.getRange('C5:C16').format.numberFormat = '@';
sheet.getRange('E5:E16').format.numberFormat = '@';
sheet.getRange('G5:G16').format.numberFormat = '@';
sheet.getRange('H5:H16').format.numberFormat = '0';

const widths = {A:14,B:10,C:14,D:10,E:14,F:10,G:14,H:10,I:9};
for (const [col,width] of Object.entries(widths)) sheet.getRange(`${col}:${col}`).format.columnWidth = width;
sheet.getRange('1:1').format.rowHeight = 38;
sheet.getRange('2:2').format.rowHeight = 24;
sheet.getRange('3:3').format.rowHeight = 30;
sheet.getRange('4:4').format.rowHeight = 25;
sheet.getRange('5:16').format.rowHeight = 25;
sheet.freezePanes.freezeRows(4);

const outDir = 'outputs/6573abd2';
await fs.mkdir(outDir, {recursive:true});
const preview = await workbook.render({sheetName:'值班表', range:'A1:I16', scale:1.5, format:'png'});
await fs.writeFile(`${outDir}/值班表预览.png`, new Uint8Array(await preview.arrayBuffer()));
console.log((await workbook.inspect({kind:'table', range:'值班表!A1:I16', include:'values,formulas', tableMaxRows:16, tableMaxCols:9})).ndjson);
console.log((await workbook.inspect({kind:'match', searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A', options:{useRegex:true,maxResults:100}, summary:'formula scan'})).ndjson);
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outDir}/值班表_可编辑.xlsx`);
